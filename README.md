# Kutubxona – school library system

A library website for Presidential School in Khiva: students browse and borrow books, librarians manage loans, and the class reading rankings motivate everyone to read more.

No installation of packages is needed. It is one Node.js server (`server.js`) and one web page (`library.html`).

## What it does

- Public home page and book catalogue (no login needed to browse)
- Student, teacher, librarian and admin accounts
- Borrow and renew requests with a chosen number of days, waiting queue for books that are out
- Reading history per student, class and house rankings, assigned reading for classes
- Librarians can do everything an admin can, except time travel and seeing or changing admin accounts
- Every action is written to an activity log
- Classes move up automatically after 15 July each year; 11th graders become "graduated"
- English, Uzbek and Russian

## Run it

1. Install [Node.js](https://nodejs.org) 18 or newer.
2. Put `server.js` and `library.html` in one folder.
3. Open a terminal in that folder and run:

   ```
   node server.js
   ```
4. Open **http://localhost:3000** in a browser.

First login: **admin / admin**. Change this password right away on the Profile page.

Another port: `PORT=3001 node server.js` (Windows: `set PORT=3001` first).

Do not open `library.html` by double-clicking it or through a static host such as Live Server or GitHub Pages: the page needs the server.

## Where the data lives

The code contains no library data. Everything is saved in files next to `server.js`:

| File | What it holds |
|---|---|
| `library_data.xlsx` | The database: members, books, copies, loans, requests, queues, assignments |
| `staff.json` | Admin and librarian logins only |
| `activity_log.xlsx` / `activity_log.jsonl` | Every action |
| `secret.key` | Key for the readable password copies. Back it up together with the Excel file |

- The server saves `library_data.xlsx` within a second after each change on the site.
- You can edit `library_data.xlsx` in Excel while the server runs. Save it and the site loads the changes within a few seconds. If Excel keeps the file locked, changes wait in `library_data.unsaved.xlsx` and are merged when the file is free.
- Passwords are stored hashed (scrypt). A readable copy is encrypted with `secret.key` so the admin can hand out forgotten passwords.

### Starting a new library

Create `library_data.xlsx` with a **Members** sheet:

| Role | First name | Last name | Class / subject | House | Username | New password (type to change) |
|---|---|---|---|---|---|---|

Username and password may stay empty: the server creates them. Student usernames look like `firstname + first 3 letters of last name + 0 + join year` (for example `alibek karimov, joined 2023` becomes `alibekkar023`). Roles are `student`, `teacher` or `staff`; admins and librarians are made in the admin panel.

Books are added in the admin panel (or in the **Books** and **Copies** sheets). Each book needs at least one copy with its own Library ID; the ISBN is 13 digits and starts with 978.

The admin panel can also download all logins plus every student's borrowing history as an Excel file.

## Private data: do not publish

Never upload these to a public place: `library_data.xlsx`, `staff.json`, `secret.key`, `sessions.json`, the activity logs and any logins Excel file. The included `.gitignore` keeps them out of Git.

## Putting it online

Run the server on a computer or server that stays on (a school PC with a tunnel, or a small VPS). Static hosting (GitHub Pages, Netlify) and short-lived serverless hosting (Vercel) cannot keep the data files. On a host with a temporary disk, set the `SECRET_KEY` environment variable (64 hex characters) and keep the Excel file on a persistent disk.
