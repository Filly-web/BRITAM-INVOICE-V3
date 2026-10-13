# TRA EFD Verification Portal

A small Flask web app that adds TRA verification details and a QR code to a Microsoft Word `.docx` invoice, then records the operation in SQLite.

## Requirements
- Python 3.10 or later
- Internet access is not required to run the app locally after dependencies are installed.
- A valid, editable `.docx` invoice (not legacy `.doc` or a password-protected document).

## Install and run (Windows)
1. Extract this folder.
2. Open Command Prompt or PowerShell in this folder.
3. Create a virtual environment:
   `py -m venv .venv`
4. Activate it:
   PowerShell: `.\.venv\Scripts\Activate.ps1`
   Command Prompt: `.venv\Scripts\activate.bat`
5. Install packages:
   `python -m pip install -r requirements.txt`
6. Start:
   `python app.py`
7. Open `http://127.0.0.1:5000` in your browser.

## How to use
1. In your Word invoice, add `<<EFD_SECTION>>` at the location where the verification code, link, and QR code should appear.
2. Upload the `.docx` file.
3. Enter the TRA verification code and receipt time (`HH:MM:SS` or `HHMMSS`).
4. Click **Process and download invoice**.

If the placeholder is absent, the portal appends the details at the end of the document. It also checks ordinary paragraphs, tables, headers, and footers for the placeholder.

## Configuration
- `PORT`: listening port (default `5000`)
- `HOST`: bind address (default `127.0.0.1`; use `0.0.0.0` only if you intentionally need access from another device)
- `SECRET_KEY`: set a strong random value for deployments
- `PORTAL_DATABASE`: optional database file path

## Important production notes
This version is suitable for local/internal testing. Before exposing it to a network or the public internet, add authentication/authorization, HTTPS, CSRF protection, rate limiting, retention controls, and secure secret management. The portal builds a URL in the configured format but does not independently verify whether TRA accepts a code/time combination. Check the generated link against the original receipt.
