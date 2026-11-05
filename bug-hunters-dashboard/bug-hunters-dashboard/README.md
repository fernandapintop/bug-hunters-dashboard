# Bug Hunters – Monthly dashboard

Static site (no build step). Open `index.html` locally or deploy the folder to Vercel (`vercel` / import the repo, framework "Other", no build command, output directory = root).

- `index.html` – page; `styles.css` – theme (light/dark); `app.js` – UI; `core.js` – metrics (pure functions).
- Upload the monthly Excel (sheet `Datos`, columns OWNER, ID, DATE, TYPE, HOURS, LOE, QA STATUS, Month). Parsing happens in the browser with SheetJS (loaded from cdnjs); nothing is sent to a server.
- Rules: cases are counted per row; a case's effort (longest case) sums all rows sharing the same ID in the month; error % = (Failed + Critical) / QA-reviewed (shown with 5+ reviewed); owners with fewer than 10 rows overall are excluded.
- Next step: replace the in-browser storage with a database (e.g. Vercel Postgres / Supabase) so months accumulate.
