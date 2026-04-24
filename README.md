# As-Built Drawing Compiler — HPC HK2794

Automates compilation of As-Built Drawing packages for Exentec Hargreaves ductwork deliveries.

## What it does

1. Reads the TRN PDF to extract all ECS codes and ductbook references
2. Scans ductwork drawing PDFs to find each ECS code's individual drawing page
3. Renders each page, rotates it to portrait, and applies the conformance stamp
4. Fills in the As-Built Word template with header fields, ECS code tables, and embedded drawings

## Files you need to upload each time

| File | Description |
|---|---|
| TRN PDF | Technical Release Note (contains the ECS code table) |
| Drawing PDFs | One or more ductbook PDFs e.g. `ENG-GSC-WS08-02393.pdf` |
| Word template | `E21369-EHL-XX-ZZ-RP-MM-000xxx.docx` |
| Stamp PNG | Conformance stamp image (portrait, red border) |

## Deploying to Streamlit Community Cloud (free)

1. Create a free account at [github.com](https://github.com) if you don't have one
2. Create a new repository and upload all files in this folder
3. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub
4. Click **New app**, select your repository, set the main file to `app.py`
5. Click **Deploy** — you'll get a URL to share with your team

That's it. Anyone with the link can use it — no Python, no install, no Claude needed.

## Running locally (optional)

```bash
pip install -r requirements.txt
streamlit run app.py
```

Requires `poppler-utils` installed on your system:
- **Windows**: Download from https://github.com/oschwartz10612/poppler-windows/releases
- **Mac**: `brew install poppler`
- **Linux**: `sudo apt install poppler-utils`

## Known ductbook references

| Ductbook | Area |
|---|---|
| ENG-GSC-WS08-02385 | HKX-08 |
| ENG-GSC-WS08-02393 | HKX-08 |
| ENG-GSC-WS08-03005 | HVL-03 |
| ENG-GSC-WS08-03046 | HWX-04 |
| ENG-GSC-WS08-02978 | HWX-05 |
