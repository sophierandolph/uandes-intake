# Setup (one time)

1. Create a Google Sheet in your Rice Drive, e.g. "UANDES Intake Responses" (UAndes / Rice folder).
2. In the Sheet: Extensions > Apps Script. Replace the code with `apps-script/Code.gs`. Save.
3. Run `testWrite` once and approve the permission prompt. A "Part 1" tab appears with a TEST row. Delete the row.
4. Deploy > New deployment > type Web app. Execute as: Me. Who has access: Anyone. Deploy and copy the URL.
5. Paste that URL into `endpoint` in `index.html`.
6. Host the folder (GitHub Pages or similar) and send Felipe the link.

Editing code.gs later: Deploy > Manage deployments > edit > New version, so the URL stays the same.
