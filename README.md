# Smart Automotive Pricelist PDF to Excel - Accuracy First

This Streamlit app is for automotive catalog PDFs where one source line may contain multiple compatible vehicles and shorthand catalog text.

## What this version improves
- Layout-aware row extraction using page coordinates
- Separate raw catalog extraction from normalized application output
- Segment-level slash parsing so one line can become multiple compatible rows
- Protected phrase handling so items like `BONGO 2000`, `CUBE 1.5`, and `MIRAGE G4` are not incorrectly treated as years
- Dedicated shorthand-year parsing such as `07'`, `'07`, `93'-UP`, and `78'-88'`
- Brand abbreviation handling such as `TOY`, `MIT`, `NIS`
- Field validators for Year, Engine, Axle, Side, Vertical, and Mount
- Field-level evidence logging and conflict logging
- Supplier memory for custom aliases and protected phrases


## Faster review controls in this package
- Bulk actions: `Accept all 95%+ rows`, `Accept all safe rule matches`
- Quick filters: `Show only conflicts`, `Show only blank important fields`, `Show only low confidence rows`
- Review dashboard row actions: `Accept selected rows`, `Mark selected for review`, `Reject AI proposal`
- Row detail actions: `Save row edits`, `Accept row`, `Accept selected fields`, `Keep selected fields blank`, `Next review row`, `Save current supplier rule/template`
- Session edits are preserved in the current run and included in the exported Excel workbook

## Workbook output
The exported workbook includes:
- `Applications`
- `Catalog Rows`
- `Review Queue`
- `Conflict Log`
- `Evidence Log`

## App files
- `streamlit_app.py` - main app
- `parser.py` - PDF extraction, segmentation, normalization, evidence/conflict generation
- `storage.py` - supplier memory storage
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
- This package is rule-based and confidence-based
- It does not call external AI APIs or live web search
- Supplier memory is stored locally in SQLite; export backups from inside the app if needed
