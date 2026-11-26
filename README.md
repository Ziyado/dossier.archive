# Dossier Archive: The Ephemera Cabinet

An interactive archive of designed ephemera, filed like a card cabinet:
- one drawer per Solution category
- a black guide card per subcategory
- one folder per work

The data comes from a Google Sheet through an Apps Script JSON feed. Images come from Google Drive, in a full-quality original and a compressed web copy.

- `index.html`: the whole site, a single file. Set `DATA_URL` near the top of its script to your Apps Script web-app URL.
- `apps-script/Code.gs`: the Sheets and Drive back end. Paste it into the sheet's Apps Script editor.
- `SETUP.md`: setup and how to use the sheet.
