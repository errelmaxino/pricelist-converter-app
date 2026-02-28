# Smart Automotive Pricelist PDF to Excel

This Streamlit app is designed for automotive catalog PDFs where one source line may contain **multiple compatible brands/models**.

## What this version does
- Splits one catalog line into multiple application rows
- Keeps the original source line in the output
- Spots common brand abbreviations like `TOY`, `MIT`, `NIS`
- Expands shorthand years like `07'`, `'07`, `14'`, `93'-UP`
- Carries the brand forward across split models when obvious
- Creates a review dashboard and row detail view based on confidence
- Exports an Excel workbook with:
  - `Applications`
  - `Catalog Rows`
  - `Review Queue`

## Files
- `streamlit_app.py` — main app
- `parser.py` — PDF parsing, line splitting, pattern spotting, workbook export
- `storage.py` — simple supplier profile storage
- `requirements.txt`
- `runtime.txt`
- `demo_output.xlsx`

## Deploy on Streamlit Cloud
1. Create a public GitHub repository
2. Upload all extracted files from this folder
3. Go to https://share.streamlit.io/
4. Sign in with GitHub
5. Create a new app and choose `streamlit_app.py` as the main file
6. Deploy

## Notes
- Best for text-based PDF catalogs
- The parser is rule-based and confidence-based
- It does **not** use external web search or live AI APIs in this package
- Supplier profiles are stored locally in SQLite; use the profile backup button inside the app if needed


## Current output behavior
- One source line can become many exported application rows
- Example: `MIT. MIRAGE G4 14'/LANCER 07'/NISSAN CUBE 1.5` becomes separate rows
- The Excel output keeps the original source line so you can trace each split row back to the catalog
