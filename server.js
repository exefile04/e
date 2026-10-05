// Kutubxona library server. Run: node server.js   (Node 18+, no npm install needed)
// Serves library.html and a small API. Passwords are stored hashed (scrypt) and never sent to browsers.
// Every change is saved to data.json AND to two Excel files next to this script:
//   library_data.xlsx   - all library data (members, books, copies, loans, requests, queues, assignments, ...)
//   activity_log.xlsx   - every action (logins, issues, returns, profile updates, assignments, ...)
// Start options:
//   node server.js --from-excel            load everything from library_data.xlsx (e.g. after editing it in Excel)
//   node server.js --restore-from-seed     reset non-admin passwords to the ones in seed.json
const http = require('http'), fs = require('fs'), path = require('path'), crypto = require('crypto'), zlib = require('zlib');
const PORT = process.env.PORT || 3000;
const SEED = path.join(__dirname, 'seed.json'), DB = path.join(__dirname, 'data.json'), PAGE = path.join(__dirname, 'library.html');
const XD = path.join(__dirname, 'library_data.xlsx'), XL = path.join(__dirname, 'activity_log.xlsx'), LJ = path.join(__dirname, 'activity_log.jsonl');
const ROLES = ['admin', 'librarian', 'teacher', 'staff', 'student'];
const staff = u => !!u && (u.role === 'admin' || u.role === 'librarian');

const hash = p => { const s = crypto.randomBytes(16).toString('hex'); return 'scrypt$' + s + '$' + crypto.scryptSync(p, s, 32).toString('hex'); };
const check = (p, h) => { const [, s, x] = String(h).split('$'); if (!x) return false; const a = Buffer.from(x, 'hex'), b = crypto.scryptSync(p, s, 32); return a.length === b.length && crypto.timingSafeEqual(a, b); };
// Passwords are hashed for login AND kept encrypted (AES-256-GCM, key in secret.key) so the admin can view/reset them.
const KF = path.join(__dirname, 'secret.key');
let keyNew = false;
const KEY = (() => {
  const fromEnv = /^[0-9a-f]{64}$/i.test(process.env.SECRET_KEY || '') ? Buffer.from(process.env.SECRET_KEY, 'hex') : null;   // hosts with temporary disks: set SECRET_KEY (64 hex chars)
  if (fromEnv) return fromEnv;
  try { const k = Buffer.from(fs.readFileSync(KF, 'utf8').trim(), 'hex'); if (k.length === 32) return k; } catch (e) { }
  const k = crypto.randomBytes(32); fs.writeFileSync(KF, k.toString('hex'), { mode: 0o600 }); keyNew = true; return k;
})();
const enc = t => { const iv = crypto.randomBytes(12), c = crypto.createCipheriv('aes-256-gcm', KEY, iv), d = Buffer.concat([c.update(String(t), 'utf8'), c.final()]); return [iv, c.getAuthTag(), d].map(b => b.toString('hex')).join('.'); };
const dec = e => { try { const [i, t, d] = String(e).split('.').map(x => Buffer.from(x, 'hex')), c = crypto.createDecipheriv('aes-256-gcm', KEY, i); c.setAuthTag(t); return Buffer.concat([c.update(d), c.final()]).toString('utf8'); } catch (x) { } };
const N2 = s => String(s || '').toLowerCase().replace(/[\s.\-_ʻʼ’‘'`]/g, '');
const canon = v => Array.isArray(v) ? '[' + v.map(canon).join(',') + ']' : v && typeof v === 'object' ? '{' + Object.keys(v).filter(k => v[k] !== undefined).sort().map(k => JSON.stringify(k) + ':' + canon(v[k])).join(',') + '}' : JSON.stringify(v === undefined ? null : v);
const same = (a, b) => canon(a) === canon(b);   // key order does not matter
const nows = p => String(p || '').replace(/\s+/g, '');
const p2 = n => String(n).padStart(2, '0'), fd = t => { const d = new Date(t); return p2(d.getDate()) + '/' + p2(d.getMonth() + 1) + '/' + d.getFullYear(); };

// ---------- usernames: [first name][first 3 letters of last name]0[2-digit year joined], e.g. O'tkirbek Bobojonov, joined 2020 -> otkirbekbob020
const CY = { а: 'a', б: 'b', в: 'v', г: 'g', д: 'd', е: 'e', ё: 'yo', ж: 'j', з: 'z', и: 'i', й: 'y', к: 'k', л: 'l', м: 'm', н: 'n', о: 'o', п: 'p', р: 'r', с: 's', т: 't', у: 'u', ф: 'f', х: 'x', ц: 'ts', ч: 'ch', ш: 'sh', щ: 'sh', ъ: '', ы: 'i', ь: '', э: 'e', ю: 'yu', я: 'ya', ў: 'o', ғ: 'g', қ: 'q', ҳ: 'h' };
const NM = s => String(s || '').toLowerCase().replace(/[ʻʼ’‘'`´]/g, '').replace(/[а-яёўғқҳ]/g, c => CY[c]).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]/g, '');
const acadYear = t => { const d = new Date(t); return d.getMonth() >= 8 ? d.getFullYear() : d.getFullYear() - 1; };   // school year starts in September
const joinYear = (c, t) => { const m = /^\s*(\d{1,2})/.exec(String(c || '')), y = acadYear(t); return m && +m[1] >= 5 ? y - (+m[1] - 5) : y; };   // grade 5 = first year at school
const baseU = m => { const b = NM(m.f) + NM(m.l).slice(0, 3); if (m.role !== 'student') return b || 'user'; return (b || 'student') + '0' + p2((+m.jy || joinYear(m.c, Date.now())) % 100); };
const sfx = i => i < 25 ? String.fromCharCode(98 + i - 1) : 'z' + i;   // b, c, d ... for students with the same name and year
const uniqU = (b, taken) => { if (!taken.has(b)) return b; for (let i = 1; ; i++) if (!taken.has(b + sfx(i))) return b + sfx(i); };

let D = null;
const loadJSON = f => JSON.parse(fs.readFileSync(f, 'utf8'));
let DREV = 1, LREV = 1;   // change counters: browsers only download the state again when something changed
const persist = () => { fs.writeFileSync(DB + '.tmp', JSON.stringify(D)); fs.renameSync(DB + '.tmp', DB); DREV++; queueX('data'); };   // write then rename, so a crash never leaves a broken file

// ---------- activity log (kept in activity_log.jsonl + activity_log.xlsx; the server writes it, browsers cannot change it)
let LOG = [];
const LL = { l_login: 'Logged in', l_lfail: 'Failed login', l_logout: 'Logged out', l_issue: 'Book issued', l_ret: 'Book returned', l_rnw: 'Loan renewed', l_req: 'Borrow request', l_rreq: 'Renewal request', l_rcan: 'Request cancelled', l_rno: 'Request declined', l_led: 'History edited', l_lrm: 'Removed from history', l_qj: 'Joined queue', l_ql: 'Left queue', l_addm: 'Member added', l_delm: 'Member deleted', l_upd: 'Profile updated', l_pw: 'Password changed', l_addb: 'Book added', l_edb: 'Book edited', l_delb: 'Book deleted', l_ftr: 'Featured changed', l_asg: 'Book assigned to classes', l_asgd: 'Assignment removed', l_un: 'Usernames updated', l_viewas: 'Admin viewed site as member', l_promo: 'New school year: classes moved up', l_sys: 'System' };
function addLog(u, k, a, b) {
  const e = { t: Date.now(), by: u ? ((u.f || '') + ' ' + (u.l || '')).trim() : 'System', bu: u ? u.u || '' : '', role: u ? u.role || '' : 'system', k, a: String(a ?? ''), b: String(b ?? '') };
  LOG.push(e); try { fs.appendFileSync(LJ, JSON.stringify(e) + '\n'); } catch (x) { console.error('Could not write activity_log.jsonl:', x.message); }
  LREV++; queueX('log');
}
const mName = m => m ? ((m.f || '') + ' ' + (m.l || '')).trim() + (m.u ? ' (' + m.u + ')' : '') : '';
function diffLog(o, n, u, pwIds) {
  const L = (k, a, b) => addLog(u, k, a, b), om = new Map(o.ms.map(m => [m.id, m])), nmm = new Map(n.ms.map(m => [m.id, m]));
  const ob = new Map(o.bs.map(b => [b.id, b])), nb = new Map(n.bs.map(b => [b.id, b])), title = id => (nb.get(id) || ob.get(id) || {}).t || '?';
  const who = id => mName(nmm.get(id) || om.get(id)) || id;
  const addedM = n.ms.filter(m => !om.has(m.id));
  if (addedM.length > 20) L('l_addm', addedM.length + ' members', addedM.slice(0, 10).map(mName).join('; ') + '; ...');
  for (const m of n.ms) {
    const x = om.get(m.id);
    if (!x) { if (addedM.length <= 20) L('l_addm', mName(m), m.role + (m.c ? ', ' + m.c : '')); continue; }
    const ch = [['f', 'first name'], ['l', 'last name'], ['c', 'class'], ['u', 'username'], ['hs', 'house'], ['role', 'role'], ['jy', 'join year']].filter(([k]) => String(x[k] ?? '') !== String(m[k] ?? '')).map(([k, lab]) => lab + ': ' + (x[k] ?? '') + ' → ' + (m[k] ?? ''));
    if (ch.length) L('l_upd', mName(m), ch.join('; '));
    if (pwIds.has(m.id)) L('l_pw', mName(m), '');
  }
  for (const x of o.ms) if (!nmm.has(x.id)) L('l_delm', mName(x), x.role);
  const bkey = b => canon({ t: b.t, alt: b.alt, a: b.a, aa: b.aa, lang: b.lang, img: b.img || '', c: (b.copies || []).map(c => [c.lib, c.isbn]) });
  const addedB = n.bs.filter(b => !ob.has(b.id));
  if (addedB.length > 20) L('l_addb', addedB.length + ' books', addedB.slice(0, 10).map(b => b.t).join('; ') + '; ...');   // bulk import: one line instead of thousands
  for (const b of n.bs) {
    const x = ob.get(b.id);
    if (!x) { if (addedB.length <= 20) L('l_addb', b.t, (b.copies || []).length + ' copies'); continue; }
    if (bkey(x) !== bkey(b)) L('l_edb', b.t, '');
    if (!!x.wk !== !!b.wk) L('l_ftr', b.t, 'Book of the Week: ' + (b.wk ? 'on' : 'off'));
    if (!!x.mo !== !!b.mo) L('l_ftr', b.t, 'Book of the Month: ' + (b.mo ? 'on' : 'off'));
  }
  for (const x of o.bs) if (!nb.has(x.id)) L('l_delb', x.t, '');
  const ol = new Map(o.loans.map(l => [l.id, l])), newLoan = new Set();
  for (const l of n.loans) {
    const x = ol.get(l.id), it = (l.t || title(l.b)) + (l.lib ? ' [' + l.lib + ']' : '');
    if (!x) { newLoan.add(l.b + '|' + l.m); L('l_issue', it, (l.who || who(l.m)) + ' · due ' + fd(l.due)); }
    else if (!x.ret && l.ret && (l.rn || 0) === (x.rn || 0) && l.iss === x.iss && l.due === x.due) L('l_ret', it, l.who || who(l.m));
    else if (!l.ret && !x.ret && l.due > x.due && (l.rn || 0) > (x.rn || 0)) L('l_rnw', it, (l.who || who(l.m)) + ' · new due date ' + fd(l.due));
    else {   // a librarian/admin corrected the history
      const ch = [['iss', 'issued'], ['due', 'due'], ['ret', 'returned']].filter(([k]) => x[k] !== l[k]).map(([k, lab]) => lab + ': ' + (x[k] ? fd(x[k]) : 'on loan') + ' → ' + (l[k] ? fd(l[k]) : 'on loan'));
      if (x.lib !== l.lib) ch.push('copy: ' + x.lib + ' → ' + l.lib);
      if (ch.length) L('l_led', it, (l.who || who(l.m)) + ' · ' + ch.join('; '));
    }
  }
  const nl = new Set(n.loans.map(l => l.id));
  for (const x of o.loans) if (!nl.has(x.id)) L('l_lrm', (x.t || title(x.b)) + (x.lib ? ' [' + x.lib + ']' : ''), (x.who || who(x.m)) + ' · ' + fd(x.iss) + (x.ret ? ' – ' + fd(x.ret) : ' (was on loan)'));
  const orq = new Map(o.rq.map(r => [r.id, r])), nrq = new Set(n.rq.map(r => r.id));
  for (const r of n.rq) if (!orq.has(r.id)) L(r.k === 'renew' ? 'l_rreq' : 'l_req', title(r.b), who(r.m));
  for (const r of o.rq) if (!nrq.has(r.id) && r.m === u.id) L('l_rcan', title(r.b), who(r.m));
  const ont = new Set(o.nt.map(x => x.id));
  for (const x of n.nt) if (!ont.has(x.id) && (x.k === 'nno' || x.k === 'rnno')) L('l_rno', x.b, who(x.m));
  for (const b of new Set([...Object.keys(o.qs), ...Object.keys(n.qs)])) {
    const A = o.qs[b] || [], B = n.qs[b] || [];
    B.filter(m => !A.includes(m)).forEach(m => L('l_qj', title(b), who(m)));
    A.filter(m => !B.includes(m) && !newLoan.has(b + '|' + m)).forEach(m => L('l_ql', title(b), who(m)));
  }
  const oas = new Map(o.as.map(a => [a.id, a])), nas = new Set(n.as.map(a => a.id));
  for (const a of n.as) if (!oas.has(a.id)) L('l_asg', title(a.b), (a.cl || []).join(', ') + ' · ' + fd(a.from) + ' – ' + fd(a.to));
  for (const a of o.as) if (!nas.has(a.id)) L('l_asgd', title(a.b), (a.cl || []).join(', '));
}

// ---------- tiny .xlsx writer / reader (no npm packages needed)
const crcT = (() => { const t = new Int32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c; } return t; })();
const crc32 = b => { let c = -1; for (let i = 0; i < b.length; i++) c = crcT[(c ^ b[i]) & 255] ^ (c >>> 8); return (c ^ -1) >>> 0; };
function zip(files) {
  const parts = [], cd = []; let off = 0;
  for (const [name, txt] of files) {
    const nb = Buffer.from(name), data = Buffer.from(txt, 'utf8'), comp = zlib.deflateRawSync(data), crc = crc32(data), lh = Buffer.alloc(30), ch = Buffer.alloc(46);
    lh.writeUInt32LE(0x04034b50, 0); lh.writeUInt16LE(20, 4); lh.writeUInt16LE(0x800, 6); lh.writeUInt16LE(8, 8); lh.writeUInt16LE(0x21, 12); lh.writeUInt32LE(crc, 14); lh.writeUInt32LE(comp.length, 18); lh.writeUInt32LE(data.length, 22); lh.writeUInt16LE(nb.length, 26);
    ch.writeUInt32LE(0x02014b50, 0); ch.writeUInt16LE(20, 4); ch.writeUInt16LE(20, 6); ch.writeUInt16LE(0x800, 8); ch.writeUInt16LE(8, 10); ch.writeUInt16LE(0x21, 14); ch.writeUInt32LE(crc, 16); ch.writeUInt32LE(comp.length, 20); ch.writeUInt32LE(data.length, 24); ch.writeUInt16LE(nb.length, 28); ch.writeUInt32LE(off, 42);
    parts.push(lh, nb, comp); cd.push(ch, nb); off += 30 + nb.length + comp.length;
  }
  const cdb = Buffer.concat(cd), e = Buffer.alloc(22);
  e.writeUInt32LE(0x06054b50, 0); e.writeUInt16LE(files.length, 8); e.writeUInt16LE(files.length, 10); e.writeUInt32LE(cdb.length, 12); e.writeUInt32LE(off, 16);
  return Buffer.concat([...parts, cdb, e]);
}
function unzip(buf) {
  let e = buf.length - 22; while (e >= 0 && buf.readUInt32LE(e) !== 0x06054b50) e--;
  if (e < 0) throw new Error('not a zip/xlsx file');
  const n = buf.readUInt16LE(e + 10), out = {}; let p = buf.readUInt32LE(e + 16);
  for (let i = 0; i < n; i++) {
    const m = buf.readUInt16LE(p + 10), cs = buf.readUInt32LE(p + 20), nl = buf.readUInt16LE(p + 28), xl = buf.readUInt16LE(p + 30), cl = buf.readUInt16LE(p + 32), lo = buf.readUInt32LE(p + 42), name = buf.toString('utf8', p + 46, p + 46 + nl);
    const ds = lo + 30 + buf.readUInt16LE(lo + 26) + buf.readUInt16LE(lo + 28), raw = buf.subarray(ds, ds + cs);
    out[name.replace(/^\//, '')] = (m === 8 ? zlib.inflateRawSync(raw) : raw).toString('utf8'); p += 46 + nl + xl + cl;
  }
  return out;
}
const XE = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '');
const XU = s => s.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (_, x) => x[0] === '#' ? String.fromCodePoint(x[1].toLowerCase() === 'x' ? parseInt(x.slice(2), 16) : +x.slice(1)) : { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }[x.toLowerCase()]);
const colL = i => { let s = ''; i++; while (i) { const r = (i - 1) % 26; s = String.fromCharCode(65 + r) + s; i = Math.floor((i - 1) / 26); } return s; };
const toSer = ms => (ms - new Date(ms).getTimezoneOffset() * 6e4) / 864e5 + 25569;   // Excel date number, in this computer's time zone
const fromSer = v => { const t = (v - 25569) * 864e5; return Math.round(t + new Date(t).getTimezoneOffset() * 6e4); };
// sheets: [{ name, cols: [[header, width, type]], rows: [[...]] }]; type 'd' = date (ms), 'n' = number, otherwise text
function xlsx(sheets) {
  const sx = sheets.map(s => {
    const head = `<row r="1">${s.cols.map((c, x) => `<c r="${colL(x)}1" t="inlineStr" s="1"><is><t>${XE(c[0])}</t></is></c>`).join('')}</row>`;
    const body = s.rows.map((r, y) => `<row r="${y + 2}">${r.map((v, x) => {
      if (v === undefined || v === null || v === '') return '';
      const ref = colL(x) + (y + 2), ty = (s.cols[x] || [])[2];
      if (ty === 'd' && typeof v === 'number' && isFinite(v)) return `<c r="${ref}" s="2"><v>${toSer(v)}</v></c>`;
      if (ty === 'n' && typeof v === 'number' && isFinite(v)) return `<c r="${ref}"><v>${v}</v></c>`;
      return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${XE(String(v).slice(0, 32767))}</t></is></c>`;
    }).join('')}</row>`).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>${s.cols.map((c, i) => `<col min="${i + 1}" max="${i + 1}" width="${c[1] || 14}" customWidth="1"/>`).join('')}</cols><sheetData>${head}${body}</sheetData></worksheet>`;
  });
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  return zip([
    ['[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((s, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`],
    ['_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ['xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}"><sheets>${sheets.map((s, i) => `<sheet name="${XE(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>`],
    ['xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="${R}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${sheets.length + 1}" Type="${R}/styles" Target="styles.xml"/></Relationships>`],
    ['xl/styles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="dd/mm/yyyy hh:mm"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF4EBD9"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'],
    ...sx.map((x, i) => [`xl/worksheets/sheet${i + 1}.xml`, x])]);
}
// returns { sheetName: [ {header: value, ...}, ... ] } - works with files saved by Excel too (shared strings)
function readXlsx(buf) {
  const z = unzip(buf), txt = s => XU((s.match(/<t\b[^>]*>([\s\S]*?)<\/t>/g) || []).map(t => t.replace(/<t\b[^>]*>|<\/t>/g, '')).join(''));
  const ss = ((z['xl/sharedStrings.xml'] || '').match(/<si>[\s\S]*?<\/si>/g) || []).map(txt), rels = {}, out = {};
  ((z['xl/_rels/workbook.xml.rels'] || '').match(/<Relationship\b[^>]*>/g) || []).forEach(r => { const id = /Id="([^"]+)"/.exec(r), t = /Target="([^"]+)"/.exec(r); if (id && t) rels[id[1]] = t[1].replace(/^\/?(xl\/)?/, 'xl/'); });
  ((z['xl/workbook.xml'] || '').match(/<sheet\b[^>]*>/g) || []).forEach(s => {
    const name = XU((/name="([^"]*)"/.exec(s) || [])[1] || ''), rid = (/r:id="([^"]+)"/.exec(s) || /\bid="([^"]+)"/.exec(s) || [])[1], xml = z[rels[rid]];
    if (!xml) return;
    const rows = [];
    for (const rm of xml.matchAll(/<row\b[^>]*?(?:\/>|>([\s\S]*?)<\/row>)/g)) {
      const row = []; let ci = 0;
      for (const cm of (rm[1] || '').matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
        const a = cm[1], ref = /\br="([A-Z]+)\d+"/.exec(a), t = (/\bt="([^"]+)"/.exec(a) || [])[1], inner = cm[2] || '';
        if (ref) ci = [...ref[1]].reduce((s, ch) => s * 26 + ch.charCodeAt(0) - 64, 0) - 1;
        const v = (/<v>([\s\S]*?)<\/v>/.exec(inner) || [])[1];
        row[ci++] = t === 's' ? ss[+v] : t === 'inlineStr' ? txt(inner) : t === 'str' || t === 'e' ? XU(v || '') : t === 'b' ? v === '1' : v === undefined ? '' : +v;
      }
      rows.push(row);
    }
    const head = (rows.shift() || []).map(h => String(h ?? '').trim());
    out[name] = rows.filter(r => r.some(v => v !== undefined && v !== '')).map(r => Object.fromEntries(head.map((h, i) => [h, r[i] ?? ''])));
  });
  return out;
}

// ---------- library_data.xlsx: what each sheet holds
const J = o => { const k = Object.keys(o).filter(x => o[x] !== undefined); return k.length ? JSON.stringify(Object.fromEntries(k.map(x => [x, o[x]]))) : ''; };
const yn = b => b ? 'yes' : '';
function dataSheets() {
  const title = id => ((D.bs.find(b => b.id === id) || {}).t) || '', mm = new Map(D.ms.map(m => [m.id, m]));
  const covers = D.bs.filter(b => b.img).map(b => [b.id, ...String(b.img).match(/[\s\S]{1,32000}/g)]), cw = Math.max(1, ...covers.map(r => r.length - 1));
  const { ms, bs, loans, rq, qs, as, nt, log, ...meta } = D;
  return [
    {
      name: 'Members', cols: [['ID', 10], ['Role', 10], ['First name', 16], ['Last name', 18], ['Class / subject', 14], ['House', 20], ['Join year', 10, 'n'], ['Username', 24], ['Previous username', 26], ['Password hash', 20], ['Password (encrypted)', 20], ['Other (JSON)', 20]],
      rows: ms.map(({ id, role, f, l, c, hs, jy, u, ou, p, pe, ...x }) => [id, role, f, l, c, hs, jy === undefined ? '' : +jy, u, ou, p, pe, J(x)])
    },
    {
      name: 'Books', cols: [['Book ID', 10], ['Title', 32], ['Title (Cyrillic)', 26], ['Author', 24], ['Author (Cyrillic)', 22], ['Language', 10], ['Copies', 8, 'n'], ['Available', 9, 'n'], ['Book of the Week', 9], ['Book of the Month', 9], ['Has cover', 8], ['Other (JSON)', 16]],
      rows: bs.map(({ id, t, alt, a, aa, lang, tot, av, wk, mo, img, copies, lib, isbn, ...x }) => [id, t, alt, a, aa, lang, (copies || []).length, (copies || []).filter(c => !c.out).length, yn(wk), yn(mo), yn(img), J(x)])
    },
    {
      name: 'Copies', cols: [['Book ID', 10], ['Title', 32], ['Copy no.', 8, 'n'], ['Library ID', 18], ['ISBN', 20], ['On loan (loan ID)', 14]],
      rows: bs.flatMap(b => (b.copies || []).map((c, i) => [b.id, b.t, i + 1, c.lib, c.isbn, c.out || '']))
    },
    {
      name: 'Loans', cols: [['Loan ID', 10], ['Status', 10], ['Book ID', 10], ['Title', 30], ['Library ID', 16], ['ISBN', 18], ['Member ID', 10], ['Member', 28], ['Username', 22], ['Issued', 17, 'd'], ['Due', 17, 'd'], ['Returned', 17, 'd'], ['Renewals', 9, 'n'], ['Other (JSON)', 14]],
      rows: loans.map(({ id, b, t, lib, isbn, m, who, u, iss, due, ret, rn, ...x }) => [id, ret ? 'returned' : 'on loan', b, t || title(b), lib, isbn, m, who, u, iss, due, ret, rn || 0, J(x)])
    },
    {
      name: 'Requests', cols: [['ID', 10], ['Type', 10], ['Book ID', 10], ['Title', 30], ['Member ID', 10], ['Member', 28], ['Loan ID', 10], ['Requested', 17, 'd'], ['Other (JSON)', 14]],
      rows: rq.map(({ id, k, b, m, loan, t, ...x }) => [id, k || 'borrow', b, title(b), m, mName(mm.get(m)), loan, t, J(x)])
    },
    {
      name: 'Queues', cols: [['Book ID', 10], ['Title', 30], ['Position', 9, 'n'], ['Member ID', 10], ['Member', 28]],
      rows: Object.entries(qs).flatMap(([b, a]) => (a || []).map((m, i) => [b, title(b), i + 1, m, mName(mm.get(m))]))
    },
    {
      name: 'Assignments', cols: [['ID', 10], ['Book ID', 10], ['Title', 30], ['Classes', 22], ['From', 17, 'd'], ['To', 17, 'd'], ['Period', 14], ['By (member ID)', 12], ['Other (JSON)', 14]],
      rows: as.map(({ id, b, cl, from, to, lab, by, ...x }) => [id, b, title(b), (cl || []).join(', '), from, to, lab, by, J(x)])
    },
    {
      name: 'Notifications', cols: [['ID', 10], ['Member ID', 10], ['Type', 8], ['Book', 30], ['Time', 17, 'd'], ['Read', 6], ['Other (JSON)', 14]],
      rows: nt.map(({ id, m, k, b, t, r, ...x }) => [id, m, k, b, t, yn(r), J(x)])
    },
    { name: 'Covers', cols: [['Book ID', 10], ...Array.from({ length: cw }, (_, i) => ['Image part ' + (i + 1), 12])], rows: covers },
    { name: 'Settings', cols: [['Key', 14], ['Value (JSON)', 30]], rows: Object.entries(meta).map(([k, v]) => [k, JSON.stringify(v)]) }];
}
const logSheets = () => [{
  name: 'Activity log', cols: [['Time', 17, 'd'], ['User', 24], ['Username', 24], ['Role', 10], ['Action', 22], ['Item', 36], ['Details', 50]],
  rows: LOG.slice(-1048000).reverse().map(e => [e.t, e.by, e.bu, e.role, LL[e.k] || e.k, e.a, e.b])
}];   // newest first

// rebuilds the library from library_data.xlsx (also accepts members/books typed in by hand)
function dataFromExcel(buf) {
  const S = readXlsx(buf), g = n => S[n] || [], str = v => v === undefined || v === null ? '' : String(v).trim(), opt = v => str(v) || undefined;
  const dt = v => { if (typeof v === 'number' && isFinite(v) && v > 0) return fromSer(v); const m = /^(\d{1,2})[./](\d{1,2})[./](\d{4})(?:[ ,T]+(\d{1,2}):(\d{2}))?/.exec(str(v)); return m ? new Date(+m[3], m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0)).getTime() : undefined; };
  const js = v => { try { const o = JSON.parse(str(v) || '{}'); return o && typeof o === 'object' && !Array.isArray(o) ? o : {}; } catch (e) { return {}; } };
  const rid = () => crypto.randomBytes(4).toString('hex'), yes = v => v === true || /^(yes|y|true|1|ha|да)$/i.test(str(v));
  const out = { ms: [], bs: [], loans: [], qs: {}, as: [], rq: [], nt: [] };
  g('Settings').forEach(r => { const k = str(r['Key']); if (k && !(k in out)) { try { out[k] = JSON.parse(str(r['Value (JSON)'])); } catch (e) { } } });
  g('Members').forEach(r => {
    const role = ROLES.includes(str(r['Role']).toLowerCase()) ? str(r['Role']).toLowerCase() : 'student', jy = parseInt(r['Join year']);
    const m = { id: str(r['ID']) || rid(), role, f: str(r['First name']), l: str(r['Last name']), c: str(r['Class / subject']), ...js(r['Other (JSON)']) };
    if (str(r['House'])) m.hs = str(r['House']); if (jy > 1990) m.jy = jy; if (str(r['Previous username'])) m.ou = str(r['Previous username']);
    m.u = str(r['Username']); m.p = str(r['Password hash']) || str(r['Password']) || crypto.randomBytes(4).toString('hex'); if (str(r['Password (encrypted)'])) m.pe = str(r['Password (encrypted)']);
    if (m.f || m.u) out.ms.push(m);
  });
  const taken = new Set(out.ms.filter(m => m.u).map(m => N2(m.u)));
  out.ms.forEach(m => { if (!m.u) { if (m.role === 'student' && !m.jy) m.jy = joinYear(m.c, Date.now()); m.u = uniqU(baseU(m), taken); taken.add(m.u); } });
  g('Loans').forEach(r => { const l = { id: str(r['Loan ID']) || rid(), b: str(r['Book ID']), t: opt(r['Title']), lib: str(r['Library ID']), isbn: str(r['ISBN']), m: str(r['Member ID']), who: str(r['Member']), u: str(r['Username']), iss: dt(r['Issued']), due: dt(r['Due']), ...js(r['Other (JSON)']) }; const ret = dt(r['Returned']); if (ret) l.ret = ret; if (+r['Renewals']) l.rn = +r['Renewals']; if (l.b && l.m) out.loans.push(l); });
  const covers = {}; g('Covers').forEach(r => { const id = str(r['Book ID']); if (id) covers[id] = Object.keys(r).filter(k => /^Image part/.test(k)).sort((a, b) => parseInt(a.slice(11)) - parseInt(b.slice(11))).map(k => str(r[k])).join(''); });
  const cps = {}; g('Copies').forEach(r => { const id = str(r['Book ID']); if (id) (cps[id] = cps[id] || []).push({ n: +r['Copy no.'] || 0, lib: str(r['Library ID']), isbn: str(r['ISBN']), out: str(r['On loan (loan ID)']) || null }); });
  const active = new Map(out.loans.filter(l => !l.ret).map(l => [l.id, l]));
  g('Books').forEach(r => {
    const id = str(r['Book ID']) || rid(), b = { id, t: str(r['Title']), alt: str(r['Title (Cyrillic)']), a: str(r['Author']), aa: str(r['Author (Cyrillic)']), lang: str(r['Language']) || 'Uzbek', ...js(r['Other (JSON)']) };
    if (!b.t) return;
    let c = (cps[id] || []).sort((x, y) => x.n - y.n).map(({ lib, isbn, out }) => ({ lib, isbn, out: out && active.has(out) ? out : null }));
    if (!c.length) c = Array.from({ length: Math.max(1, +r['Copies'] || 1) }, () => ({ lib: '', isbn: '', out: null }));
    b.copies = c; b.tot = c.length; b.av = c.filter(x => !x.out).length; b.lib = c[0].lib; b.isbn = c[0].isbn;
    if (yes(r['Book of the Week'])) b.wk = true; if (yes(r['Book of the Month'])) b.mo = true; if (covers[id]) b.img = covers[id];
    out.bs.push(b);
  });
  g('Requests').forEach(r => { const x = { id: str(r['ID']) || rid(), b: str(r['Book ID']), m: str(r['Member ID']), t: dt(r['Requested']) || Date.now(), ...js(r['Other (JSON)']) }; if (str(r['Type']) === 'renew') { x.k = 'renew'; x.loan = str(r['Loan ID']); } if (x.b && x.m) out.rq.push(x); });
  g('Queues').sort((a, b) => (+a['Position'] || 0) - (+b['Position'] || 0)).forEach(r => { const b = str(r['Book ID']), m = str(r['Member ID']); if (b && m) (out.qs[b] = out.qs[b] || []).push(m); });
  g('Assignments').forEach(r => { const a = { id: str(r['ID']) || rid(), b: str(r['Book ID']), cl: str(r['Classes']).split(/\s*,\s*/).filter(Boolean), from: dt(r['From']), to: dt(r['To']), lab: str(r['Period']) || 'Custom range', by: str(r['By (member ID)']), ...js(r['Other (JSON)']) }; if (a.b && a.cl.length && a.from && a.to) out.as.push(a); });
  g('Notifications').forEach(r => { const x = { id: str(r['ID']) || rid(), m: str(r['Member ID']), k: str(r['Type']), b: str(r['Book']), t: dt(r['Time']) || Date.now(), ...js(r['Other (JSON)']) }; if (yes(r['Read'])) x.r = 1; if (x.m) out.nt.push(x); });
  if (!out.ms.some(m => m.role === 'admin')) out.ms.unshift({ id: 'a1', role: 'admin', f: 'Admin', l: '', c: '', u: 'admin', p: 'admin' });
  return out;
}

// write the Excel files shortly after each change (if a file is open in Excel and locked, retry until it is closed)
const XQ = {}, XW = {};
function queueX(kind) { clearTimeout(XQ[kind]); XQ[kind] = setTimeout(() => writeX(kind), 300); }
function writeX(kind) {
  const f = kind === 'data' ? XD : XL;
  try { const buf = xlsx(kind === 'data' ? dataSheets() : logSheets()); fs.writeFileSync(f + '.tmp', buf); fs.renameSync(f + '.tmp', f); if (XW[kind]) { console.log('Saved ' + path.basename(f) + ' again.'); XW[kind] = 0; } }
  catch (e) { if (!XW[kind]) console.warn('Could not save ' + path.basename(f) + ' (' + e.code + '). If it is open in Excel, close it; retrying every 10 s. data.json is still saved.'); XW[kind] = 1; try { fs.unlinkSync(f + '.tmp'); } catch (x) { } clearTimeout(XQ[kind]); XQ[kind] = setTimeout(() => writeX(kind), 1e4); }
}

// ---------- load data
const FROMX = process.argv.includes('--from-excel');
let src;
try {
  if (FROMX && fs.existsSync(XD)) { D = dataFromExcel(fs.readFileSync(XD)); src = 'library_data.xlsx'; }
  else if (fs.existsSync(DB)) { D = loadJSON(DB); src = 'data.json'; }
  else if (fs.existsSync(XD)) { D = dataFromExcel(fs.readFileSync(XD)); src = 'library_data.xlsx (data.json was missing)'; }
  else if (fs.existsSync(SEED)) { D = loadJSON(SEED); src = 'seed.json'; }
} catch (e) { console.error('Could not read ' + (FROMX ? 'library_data.xlsx' : 'data.json') + ':', e.message); process.exit(1); }
if (FROMX && !fs.existsSync(XD)) console.warn('--from-excel: library_data.xlsx not found, using ' + src + '.');
D = D || { ms: [{ id: 'a1', role: 'admin', f: 'Admin', l: '', c: '', u: 'admin', p: 'admin' }] }; src = src || 'defaults';
['ms', 'bs', 'loans', 'as', 'rq', 'nt'].forEach(k => { if (!Array.isArray(D[k])) D[k] = []; }); if (!D.qs || typeof D.qs !== 'object') D.qs = {};
try { LOG = fs.readFileSync(LJ, 'utf8').split('\n').filter(Boolean).map(l => { try { return JSON.parse(l); } catch (e) { } }).filter(Boolean); }
catch (e) {   // no log file yet: take the history from activity_log.xlsx if there is one
  try { if (fs.existsSync(XL)) { const K = Object.fromEntries(Object.entries(LL).map(([k, v]) => [v, k])); LOG = (readXlsx(fs.readFileSync(XL))['Activity log'] || []).map(r => ({ t: typeof r['Time'] === 'number' ? fromSer(r['Time']) : Date.now(), by: String(r['User'] || ''), bu: String(r['Username'] || ''), role: String(r['Role'] || ''), k: K[r['Action']] || String(r['Action'] || ''), a: String(r['Item'] || ''), b: String(r['Details'] || '') })).reverse(); fs.writeFileSync(LJ, LOG.map(x => JSON.stringify(x)).join('\n') + (LOG.length ? '\n' : '')); } } catch (x) { console.warn('Could not read activity_log.xlsx:', x.message); }
}
if (Array.isArray(D.log) && D.log.length) {   // older versions kept a short log inside data.json: move it into the log file
  const old = D.log.map(e => ({ t: +e.t || 0, by: String(e.by || ''), bu: '', role: '', k: e.k, a: String(e.a || ''), b: String(e.b || '') })).filter(e => !LOG.some(x => x.t === e.t && x.k === e.k && x.a === e.a)).sort((a, b) => a.t - b.t);
  LOG = old.concat(LOG).sort((a, b) => a.t - b.t); fs.writeFileSync(LJ, LOG.map(x => JSON.stringify(x)).join('\n') + '\n');
}
delete D.log;

const hashAsync = p => new Promise((ok, no) => { const s = crypto.randomBytes(16).toString('hex'); crypto.scrypt(p, s, 32, (e, k) => e ? no(e) : ok('scrypt$' + s + '$' + k.toString('hex'))); });
async function securePasswords() {   // hash plaintext seeds once; give every member an encrypted copy so the admin can view it
  const FORCE = process.argv.includes('--restore-from-seed');   // node server.js --restore-from-seed : reset passwords to the ones in seed.json
  let seedPw = {}; try { loadJSON(SEED).ms.forEach(m => { if (m.p && !String(m.p).startsWith('scrypt$')) { seedPw['id:' + m.id] = nows(m.p); seedPw['u:' + N2(m.u)] = nows(m.p); } }); } catch (e) { }
  const seedOf = m => seedPw['u:' + N2(m.u)] || (m.ou && seedPw['u:' + N2(m.ou)]) || seedPw['id:' + m.id];   // match by username first (ids may differ between seed.json and data.json)
  const forced = m => FORCE && m.role !== 'admin' && seedOf(m);   // restore never touches admin accounts
  const todo = D.ms.filter(m => !String(m.p).startsWith('scrypt$') || !m.pe || dec(m.pe) === undefined || forced(m));   // also redo copies that cannot be decrypted
  if (keyNew) console.log('NOTE: a new secret.key was created. Keep it next to data.json (back it up!).');
  if (!todo.length) { console.log('Readable passwords: ' + D.ms.length + ' of ' + D.ms.length); return; }
  console.log('Securing ' + todo.length + ' passwords, please wait a few seconds...');
  let lost = [];
  await Promise.all(todo.map(async m => {
    const sp = seedOf(m);
    if (forced(m)) { m.p = await hashAsync(sp); m.pe = enc(sp); }
    else if (!String(m.p).startsWith('scrypt$')) { const pw = nows(m.p); m.p = await hashAsync(pw); m.pe = enc(pw); }
    else if (sp && check(sp, m.p)) { m.pe = enc(sp); }   // password unchanged since seeding
    else { delete m.pe; lost.push(m.u); }
  }));
  persist();
  console.log('Readable passwords: ' + D.ms.filter(m => m.pe && dec(m.pe) !== undefined).length + ' of ' + D.ms.length + '.' + (lost.length ? ' Could NOT recover ' + lost.length + ' (they are not in seed.json or their password was changed; in the admin panel use the "Give readable passwords" button): ' + lost.slice(0, 5).join(', ') + (lost.length > 5 ? ', ...' : '') + '.' : ''));
}
function migrateUsernames() {   // one time: switch students to the new username format. Old usernames keep working for login.
  if (D.unv >= 2) return;
  const taken = new Set(D.ms.filter(m => m.role !== 'student').map(m => N2(m.u))), ch = [];
  D.ms.forEach(m => {
    if (m.role !== 'student') return;
    if (!m.jy) m.jy = joinYear(m.c, Date.now());
    const u = uniqU(baseU(m), taken); taken.add(u);
    if (u !== N2(m.u)) { ch.push(m.u + ' → ' + u); m.ou = m.u; m.u = u; }
  });
  D.unv = 2; persist();
  if (ch.length) { addLog(null, 'l_un', ch.length + ' student usernames changed to the new format', ch.slice(0, 40).join('; ') + (ch.length > 40 ? '; ...' : '')); console.log('Changed ' + ch.length + ' student usernames to the new format (old usernames still work for login). Download the new logins from the admin panel.'); }
}

// ---------- every year after 15 July: students move up one grade (5-01 -> 6-01 ... 10-02 -> 11-02); 11th graders become "graduated"
const lastPromoYear = t => { const d = new Date(t), y = d.getFullYear(); return d >= new Date(y, 6, 15) ? y : y - 1; };   // month 6 = July
function promoteClasses() {
  const due = lastPromoYear(Date.now());
  if (D.promo === undefined) { D.promo = due; persist(); return; }   // first run: today's classes are already correct
  if (D.promo >= due) return;
  let up = 0, grad = 0; const ex = [];
  while (D.promo < due) {
    D.promo++;
    D.ms.forEach(m => {
      if (m.role !== 'student' || m.gr) return;
      const g = /^\s*(\d{1,2})(.*)$/.exec(String(m.c || ''));
      if (!g) return;
      const n = +g[1];
      if (n >= 11) { m.gr = D.promo; m.gc = m.c; m.c = ''; grad++; if (ex.length < 30) ex.push(mName(m) + ': ' + m.gc + ' → graduated'); }
      else { const nc = (n + 1) + g[2]; if (ex.length < 30) ex.push(mName(m) + ': ' + m.c + ' → ' + nc); m.c = nc; up++; }
    });
  }
  persist();
  addLog(null, 'l_promo', up + ' students moved up a grade, ' + grad + ' graduated (' + D.promo + ')', ex.join('; ') + (up + grad > ex.length ? '; ...' : ''));
  console.log('New school year ' + D.promo + ': ' + up + ' students moved up, ' + grad + ' graduated.');
}
setInterval(() => { try { promoteClasses(); } catch (e) { console.error('Class promotion failed:', e); } }, 36e5).unref();   // checked every hour, so a server left running also promotes
const view = u => ({ ...D, ms: D.ms.map(({ p, pe, ...m }) => { const can = u.role === 'admin' || m.id === u.id || (u.role === 'librarian' && m.role !== 'admin'); const pl = pe && can ? dec(pe) : undefined; return pl === undefined ? m : { ...m, p: pl }; }), log: staff(u) ? LOG.slice(-1000).reverse() : [] });   // never expose hashes; plaintext only to admin / librarian (not admin passwords) / the owner
const SF = path.join(__dirname, 'sessions.json'), fails = {};      // login failures, counted per IP + username (a whole school can share one IP)
const ipOf = req => String(req.headers['x-forwarded-for'] || req.socket.remoteAddress || '').split(',')[0].trim();
const blocked = (k, max, win) => { const f = fails[k]; return !!f && Date.now() - f.t < win && f.n >= max; };
const bump = (k, win) => { const f = fails[k]; if (!f || Date.now() - f.t >= win) fails[k] = { n: 1, t: Date.now() }; else f.n++; };
setInterval(() => { for (const k in fails) if (Date.now() - fails[k].t > 6e5) delete fails[k]; }, 6e4).unref();
let S = {}; try { S = JSON.parse(fs.readFileSync(SF, 'utf8')); } catch (e) { }   // hashed token -> {id, exp}
const th = t => crypto.createHash('sha256').update(String(t)).digest('hex');
const saveS = () => fs.writeFileSync(SF, JSON.stringify(S));

// What each role may change. Admins can change everything; librarians everything except admin accounts; others only their own records.
function allowed(u, n) {
  if (!n || !Array.isArray(n.ms) || !Array.isArray(n.bs) || !Array.isArray(n.loans) || !Array.isArray(n.rq) || !Array.isArray(n.nt) || !n.qs || typeof n.qs !== 'object' || !Array.isArray(n.as)) return false;
  if (!n.ms.every(m => m && typeof m.id === 'string' && ROLES.includes(m.role))) return false;
  if (u.role === 'admin') return n.ms.some(m => m.role === 'admin');   // never leave the library without an admin
  if (u.role === 'librarian') {
    const me = n.ms.find(m => m.id === u.id); if (!me || me.role !== 'librarian') return false;   // cannot delete themselves or change their own role
    for (const m of D.ms) if (m.role === 'admin') {   // admin accounts (and their passwords) are off limits
      const x = n.ms.find(y => y.id === m.id); if (!x) return false;
      const { p, pe, ...a } = m, { p: xp, pe: _e, ...b } = x;
      if (!same(a, b) || (xp !== undefined && xp !== '')) return false;
    }
    return !n.ms.some(x => x.role === 'admin' && !D.ms.some(m => m.id === x.id && m.role === 'admin'));   // no new admins, no promotions
  }
  if (n.ms.length !== D.ms.length) return false;
  for (const m of D.ms) {
    const x = n.ms.find(y => y.id === m.id); if (!x) return false;
    const { p, pe, ...a } = m, { p: _p, pe: _pe, ...b } = x;
    if (m.id === u.id) { if (['role', 'c', 'hs', 'ld', 'jy', 'ou', 'gr', 'gc'].some(k => !same(x[k], m[k]))) return false; } else if (!same(a, b)) return false;
  }
  const keep = (a, b, f) => same(a.filter(f), b.filter(f));
  if (!same(D.bs, n.bs) || !same(D.loans, n.loans)) return false;
  if (!keep(D.rq, n.rq, r => r.m !== u.id) || n.rq.some(r => r.m === u.id && r.k && r.k !== 'renew')) return false;
  if (!keep(D.nt, n.nt, x => x.m !== u.id)) return false;
  if (u.role === 'teacher' ? !keep(D.as, n.as, a => a.by !== u.id) : !same(D.as, n.as)) return false;
  return Object.keys({ ...D.qs, ...n.qs }).every(b => same((D.qs[b] || []).filter(i => i !== u.id), (n.qs[b] || []).filter(i => i !== u.id)));
}

function apply(n, u) {
  const pwIds = new Set();
  n.ms = n.ms.map(({ pe: _x, p: sent, ...x }) => {
    const old = D.ms.find(m => m.id === x.id), pw = nows(sent);
    if (pw && !pw.startsWith('scrypt$')) {
      if (old && old.pe && dec(old.pe) === pw) return { ...x, p: old.p, pe: old.pe };   // unchanged
      if (old) pwIds.add(x.id);
      return { ...x, p: hash(pw), pe: enc(pw) };
    }
    return old ? { ...x, p: old.p, pe: old.pe } : { ...x, p: hash(crypto.randomBytes(6).toString('hex')) };
  });
  const names = n.ms.map(m => N2(m.u));
  if (names.some(u => !u) || new Set(names).size < names.length) return false;
  delete n.log; n.unv = D.unv; n.promo = D.promo;
  try { diffLog(D, n, u, pwIds); } catch (e) { console.error('Activity log error:', e); }
  D = n; persist(); return true;
}

const send = (res, code, obj) => { res.writeHead(code, { 'Content-Type': 'application/json' }); res.end(JSON.stringify(obj)); };
const body = req => new Promise((ok, no) => { let b = '', n = 0; req.on('data', c => { n += c.length; if (n > 60e6) { req.destroy(); no(); } else b += c; }); req.on('end', () => { try { ok(JSON.parse(b || '{}')); } catch (e) { no(e); } }); });
const sess = req => S[th((req.headers.authorization || '').slice(7))];
const who = req => { const k = th((req.headers.authorization || '').slice(7)), s = S[k]; if (!s) return; if (s.exp < Date.now()) { delete S[k]; saveS(); return; } return D.ms.find(m => m.id === s.id); };
// what visitors who are not signed in may see: books, rankings and class assignments (no passwords, loans details, staff or full student names)
let PUBC = null, PUBR = -1;
const pubState = () => {
  if (PUBR !== DREV) {
    const st = D.ms.filter(m => m.role === 'student'), ids = new Set(st.map(m => m.id));
    PUBC = { dr: DREV, state: { ms: st.map(m => ({ id: m.id, role: 'student', f: m.f, l: m.l ? [...m.l][0] + '.' : '', c: m.c, hs: m.hs, gr: m.gr })), bs: D.bs.map(({ copies, lib, isbn, ...b }) => ({ ...b, copies: [] })), loans: D.loans.map(l => ({ id: l.id, b: l.b, m: ids.has(l.m) ? l.m : '', iss: l.iss, due: l.due, ret: l.ret })), as: D.as.map(({ by, ...a }) => a), qs: {}, rq: [], nt: [] } };
    PUBR = DREV;
  }
  return PUBC;
};
const full = u => ({ uid: u.id, state: view(u), dr: DREV, lr: LREV });

const server = http.createServer(async (req, res) => {
  try {
    const [url, qstr] = req.url.split('?'), qp = new URLSearchParams(qstr || '');
    if (req.method === 'GET' && (url === '/' || url === '/index.html')) { res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-cache' }); return res.end(fs.readFileSync(PAGE)); }
    if (url === '/api/ping') return send(res, 200, { ok: 1 });
    if (req.method === 'GET' && url === '/api/public') return send(res, 200, +qp.get('dr') === DREV ? { same: 1 } : pubState());
    if (req.method === 'POST' && url === '/api/login') {
      const { u, p } = await body(req), nu = N2(u), k1 = ipOf(req) + '|' + nu, k2 = 'user|' + nu;
      if (blocked(k1, 8, 3e5) || blocked(k2, 40, 6e5)) return send(res, 429, { error: 'Too many attempts. Wait a few minutes.' });
      const m = D.ms.find(x => N2(x.u) === nu) || D.ms.find(x => x.ou && N2(x.ou) === nu);   // old usernames still work
      if (!m || !check(nows(p), m.p)) { bump(k1, 3e5); bump(k2, 6e5); console.log('Login failed for "' + String(u).slice(0, 40) + '": ' + (m ? 'wrong password' : 'no such username')); addLog(m || { f: '(unknown user)', u: String(u || '').slice(0, 40), role: '' }, 'l_lfail', String(u || '').slice(0, 60), (m ? 'wrong password' : 'no such username') + ' · IP ' + ipOf(req)); return send(res, 401, { error: 'bad login' }); }
      delete fails[k1];
      const token = crypto.randomBytes(32).toString('hex'); S[th(token)] = { id: m.id, exp: Date.now() + 30 * 864e5 }; saveS();
      addLog(m, 'l_login', mName(m), (N2(m.u) === nu ? '' : 'used old username ' + m.ou + ' · ') + 'IP ' + ipOf(req));
      return send(res, 200, { token, ...full(m) });
    }
    const u = who(req), ss = u && sess(req), adm = ss && ss.by ? D.ms.find(m => m.id === ss.by) : null;
    const actor = adm ? { ...u, f: ((adm.f || '') + ' ' + (adm.l || '')).trim(), l: '(as ' + mName(u) + ')', u: adm.u, role: 'admin' } : u;   // admin "view as": changes are logged under the admin
    if (url.startsWith('/api/') && !u) return send(res, 401, { error: 'login required' });
    if (req.method === 'POST' && url === '/api/logout') { delete S[th((req.headers.authorization || '').slice(7))]; saveS(); if (!adm) addLog(u, 'l_logout', mName(u), ''); return send(res, 200, { ok: 1 }); }
    if (req.method === 'POST' && url === '/api/viewas') {   // admin only: open the site as another member without logging out
      if (u.role !== 'admin' || adm) return send(res, 403, { error: 'not allowed' });
      const { id } = await body(req), m = D.ms.find(x => x.id === id);
      if (!m || m.id === u.id) return send(res, 404, { error: 'no such member' });
      const token = crypto.randomBytes(32).toString('hex'); S[th(token)] = { id: m.id, exp: Date.now() + 8 * 36e5, by: u.id }; saveS();
      addLog(u, 'l_viewas', mName(m), m.role);
      return send(res, 200, { token, ...full(m) });
    }
    if (req.method === 'GET' && url === '/api/state') {
      if (+qp.get('dr') === DREV && (!staff(u) || +qp.get('lr') === LREV)) return send(res, 200, { same: 1 });   // nothing changed: tiny reply
      return send(res, 200, full(u));
    }
    if (req.method === 'PUT' && url === '/api/state') {
      const { state } = await body(req);
      if (state && D.promo !== undefined && (state.promo || 0) < D.promo) return send(res, 412, { error: 'stale: the school year changed, reload' });   // a page still showing last year's classes must not undo the promotion
      if (!allowed(u, state)) return send(res, 403, { error: 'not allowed' });
      if (!apply(state, actor)) return send(res, 409, { error: 'usernames must be unique' });
      return send(res, 200, full(D.ms.find(m => m.id === u.id) || u));
    }
    if (req.method === 'GET' && (url === '/api/excel/data' || url === '/api/excel/log')) {
      const isData = url === '/api/excel/data';
      if (isData ? u.role !== 'admin' : !staff(u)) return send(res, 403, { error: 'not allowed' });
      const buf = xlsx(isData ? dataSheets() : logSheets());
      res.writeHead(200, { 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'Content-Disposition': 'attachment; filename="' + (isData ? 'library_data' : 'activity_log') + '.xlsx"' }); return res.end(buf);
    }
    send(res, 404, { error: 'not found' });
  } catch (e) { send(res, 400, { error: 'bad request' }); }
});
securePasswords().then(() => { migrateUsernames(); promoteClasses(); if (src.startsWith('library_data.xlsx')) { persist(); addLog(null, 'l_sys', 'Library loaded from ' + src, D.ms.length + ' members, ' + D.bs.length + ' books'); } writeX('data'); writeX('log'); server.listen(PORT, () => console.log('Library running at http://localhost:' + PORT + '  (' + D.ms.length + ' members, ' + D.bs.length + ' books loaded from ' + src + ')  Excel files: library_data.xlsx, activity_log.xlsx' + (D.ms.some(m => m.role === 'admin' && m.pe && dec(m.pe) === 'admin') ? '\nWARNING: the admin password is still "admin". Change it in the Profile page!' : ''))); }).catch(e => { console.error('Startup failed:', e); process.exit(1); });
