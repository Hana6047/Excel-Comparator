# Excel Comparator

A web app (HTML, CSS, JavaScript) that compares two Excel or CSV files and shows the differences.
Everything runs in the browser: no file is ever sent to a server.

## Features
- Import `.xlsx`, `.xls`, `.xlsm`, `.csv` files (click or drag and drop)
- Pick the sheet for each file
- Match rows by position or by a key column (identifier)
- Options: ignore leading/trailing spaces, ignore case
- Total number of modified cells, count per column, and a detailed list (cell, row, column, old and new value)
- Count of rows added in B or missing from B
- Automatic dark mode

## Run locally
Open `index.html` in a browser (an internet connection is needed to load the SheetJS library from a CDN).

## Publish on GitHub Pages (to share with colleagues)
1. Create a GitHub repository and push these files to it (`index.html`, `style.css`, `app.js`, `README.md`).
2. In the repository: **Settings → Pages**.
3. Source: **Deploy from a branch**, branch `main`, folder `/ (root)`, then **Save**.
4. After a minute the site is available at `https://<your-account>.github.io/<repository-name>/`.

## Offline use (optional)
Download `xlsx.full.min.js` (SheetJS 0.18.5), add it to the repository, and replace the `<script src="https://cdnjs...">` tag in `index.html` with `<script src="xlsx.full.min.js"></script>`.

## Structure
- `index.html`: page structure
- `style.css`: styles (light/dark theme)
- `app.js`: file reading and comparison logic
