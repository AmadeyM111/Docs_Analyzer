## Docs Analyzer

A small CLI tool to extract embedded images from Excel (.xlsx) files and generate an enriched JSON report with image properties and EXIF metadata.

### Features
- Extracts images from worksheets via openpyxl and directly from `xl/media/` in the XLSX archive
- Saves images to an output directory
- Inspects each image (format, size, DPI, color mode, file size, MIME, SHA‑256)
- Parses key EXIF fields (make, model, date/time original, orientation) and includes full EXIF dump
- Exports a consolidated JSON report to a file

### Requirements
- Python 3.10+
- See `requirements.txt`

### Installation
```bash
python -m venv venv
source venv/bin/activate  # Windows: venv\\Scripts\\activate
pip install -r requirements.txt
```

### Usage
From the project directory:
```bash
python Docs_Analyzer.py --xlsx "MyWorkbook.xlsx" --out export_images --log-level INFO --out-json ./images_report.json
```

Common flags:
- `--xlsx`: path to the source XLSX file
- `--out`: output directory for extracted images
- `--log-level`: one of `CRITICAL|ERROR|WARNING|INFO|DEBUG`
- `--out-json`: path to save the enriched JSON report
- `--json`: additionally print the JSON report to stdout

Examples:
```bash
# Basic run with JSON file output
python Docs_Analyzer.py \
  --xlsx "/absolute/path/to/Workbook.xlsx" \
  --out "/absolute/path/to/export_images" \
  --log-level INFO \
  --out-json "/absolute/path/to/images_report.json"

# Print the report to stdout (in addition to saving)
python Docs_Analyzer.py --xlsx Workbook.xlsx --out export_images --json
```

### Notes
- If you re-run the tool and images already exist, they are still indexed for the report.
- Some EXIF fields may not be present depending on image format/source.

### License
TBD.


