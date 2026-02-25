# T2 Factor Visualizer

Simple local web app to visualize the `T2 Master.xlsx` dataset by:
- `Sheet` (variable)
- `Country`

## Files
- `scripts/extract_t2_master.py`: Converts the workbook into JSON.
- `data/t2_master.json`: Generated dataset for the frontend.
- `index.html`: UI entrypoint.
- `assets/app.js`: Dropdown logic, quick text parser, chart rendering.
- `assets/styles.css`: Styling.

## Refresh dataset from Excel
Run from this folder's parent (`Amit`):

```bash
python3 app/scripts/extract_t2_master.py \
  --input "/Users/arjundivecha/Dropbox/AAA Backup/A Complete/T2 Factor Timing Fuzzy/T2 Master.xlsx" \
  --output "app/data/t2_master.json"
```

## Start the app
From `Amit`:

```bash
cd app
python3 -m http.server 8000
```

Open:
- `http://localhost:8000`

## Quick pick examples
- `India Trailing PE`
- `U.S. Earnings Yield`
- `Japan 120MA Signal`

The quick pick also normalizes common variants, including `training` -> `trailing` and `P/E` -> `PE`.
