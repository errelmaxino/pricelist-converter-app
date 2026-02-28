# Multi-Format Pricelist PDF to Excel (Online Web App)

This version is built for **more than one supplier format**.

## What it does

- Auto-detects table headers from many **text-based pricelist PDFs**
- Preserves **unknown columns** instead of dropping them
- Lets you **map detected columns once**
- Saves a **supplier template** so future uploads can reuse the same mapping
- Supports two output styles:
  - **Normalized pricelist**: standard columns such as Category, Brand, Part No., OE No., Model, Year, Size, Original Price, Your Price, Use Price, plus extra columns preserved
  - **Exact detected columns**: keeps the uploaded file's detected columns as closely as possible

## Built-in year cleanup rules

The app keeps the special cleanup rules requested for automotive pricelists, for example:

- `'93-up` -> `1993 - Up`
- `'96-` -> `1996 - Up`
- `'93-'95` -> `1993 - 1995`
- fixes common split cases like `1992 - 1996` being broken between Model and Size

## Template memory

Templates are stored in a local SQLite database inside the app package.

This works well for local/self-hosted deployments. On some free cloud platforms, local storage may reset after a rebuild or restart. Because of that, the app also includes:

- **Download template backup**
- **Import template backup**
- **Download current template**

If you want true long-term cloud memory across all devices without manual backup, connect the storage layer to a managed database service such as Supabase, Neon, or Firebase.

## Deploy online with Streamlit Cloud

1. Create a new GitHub repository
2. Upload the files from this package
3. Open https://share.streamlit.io/
4. Create a new app
5. Set the main file path to `streamlit_app.py`

Once deployed, Streamlit gives you a public URL you can open anywhere.

## Run locally

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Files included

- `streamlit_app.py` - web interface
- `parser.py` - PDF extraction, mapping, and Excel generation logic
- `storage.py` - supplier template storage
- `requirements.txt` - Python dependencies
- `runtime.txt` - Python runtime version
- `demo_output.xlsx` - sample normalized output

## Notes

- Best for **text-based PDFs with real selectable text**.
- Scanned image PDFs may need OCR first.
- Very unusual layouts may still need manual column mapping in the app.
