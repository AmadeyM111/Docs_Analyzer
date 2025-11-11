## Docs Analyzer

A CLI tool for analyzing Excel (.xlsx) files with two main modes:
1. **Images extraction**: Extract embedded images and generate enriched JSON reports with image properties and EXIF metadata
2. **Text statistics**: Compute text statistics (characters, words, tokens) per sheet and overall totals, including token estimation for embedded images

### Features

#### Images Mode
- Extracts images from worksheets via openpyxl and directly from `xl/media/` in the XLSX archive
- Saves images to an output directory
- Inspects each image (format, size, DPI, color mode, file size, MIME, SHA‑256)
- Parses key EXIF fields (make, model, date/time original, orientation) and includes full EXIF dump
- Exports a consolidated JSON report to a file

#### Text Statistics Mode
- Counts text cells, characters, words, and tokens per worksheet
- Supports tiktoken for accurate token counting or heuristic estimation (configurable chars per token)
- Processes embedded images: converts to base64 and counts tokens using the same method as text
- Provides comprehensive statistics including combined text + image token totals
- Useful for calculating token usage for AI/LLM processing

### Requirements
- Python 3.10+
- See `requirements.txt`
- Optional: `tiktoken` for accurate token counting (install separately if needed)

### Installation
```bash
python -m venv venv
source venv/bin/activate  # Windows: venv\\Scripts\\activate
pip install -r requirements.txt
```

### Usage

#### Images Extraction Mode
From the project directory:
```bash
# Legacy style (backward compatible)
python Docs_Analyzer.py --xlsx "MyWorkbook.xlsx" --out export_images --out-json ./images_report.json

# Explicit subcommand
python Docs_Analyzer.py images --xlsx "MyWorkbook.xlsx" --out export_images --out-json ./images_report.json
```

Common flags for images mode:
- `--xlsx`: path to the source XLSX file
- `--out`: output directory for extracted images
- `--log-level`: one of `CRITICAL|ERROR|WARNING|INFO|DEBUG`
- `--out-json`: path to save the enriched JSON report
- `--json`: additionally print the JSON report to stdout

#### Text Statistics Mode
```bash
python Docs_Analyzer.py text-stats --xlsx "MyWorkbook.xlsx" --out-json ./text_stats.json
```

Options for text-stats mode:
- `--xlsx`: path to the source XLSX file (required)
- `--out-json`: path to save the text statistics JSON report (required)
- `--use-tiktoken`: use tiktoken library for accurate token counting (default: heuristic)
- `--encoding`: tiktoken encoding name (default: `cl100k_base`)
- `--chars-per-token`: characters per token for heuristic estimation (default: `3.0`, i.e., 3 chars = 1 token)
- `--log-level`: logging level (default: `INFO`)

### Examples

#### Images Extraction
```bash
# Basic run with JSON file output
python Docs_Analyzer.py images \
  --xlsx "/absolute/path/to/Workbook.xlsx" \
  --out "/absolute/path/to/export_images" \
  --log-level INFO \
  --out-json "/absolute/path/to/images_report.json"

# Print the report to stdout (in addition to saving)
python Docs_Analyzer.py images --xlsx Workbook.xlsx --out export_images --json
```

#### Text Statistics
```bash
# Basic text statistics with heuristic token counting (3 chars = 1 token)
python Docs_Analyzer.py text-stats \
  --xlsx "Workbook.xlsx" \
  --out-json ./text_stats.json

# Use tiktoken for accurate token counting
python Docs_Analyzer.py text-stats \
  --xlsx "Workbook.xlsx" \
  --out-json ./text_stats.json \
  --use-tiktoken \
  --encoding cl100k_base

# Custom heuristic (4 chars = 1 token)
python Docs_Analyzer.py text-stats \
  --xlsx "Workbook.xlsx" \
  --out-json ./text_stats.json \
  --chars-per-token 4.0
```

### How Image Token Counting Works

For text statistics mode, embedded images are processed as follows:
1. Images are read from the `xl/media/` directory in the XLSX archive
2. Each image is converted to a base64-encoded string (as it would be sent to Vision API endpoints)
3. Tokens are counted for the base64 string using the same method as text:
   - If `--use-tiktoken` is set: tiktoken encoder is used
   - Otherwise: heuristic estimation (default: 3 characters = 1 token)
4. Token counts are summed across all images and combined with text tokens

This approach accurately reflects how images are processed in AI/LLM pipelines that use base64 encoding.

### Notes
- If you re-run the images extraction tool and images already exist, they are still indexed for the report.
- Some EXIF fields may not be present depending on image format/source.
- Text statistics mode processes images without extracting them to disk (reads directly from the XLSX archive).
- The default heuristic (3 chars = 1 token) is a reasonable approximation for many tokenizers, but tiktoken provides more accurate counts.

### License
TBD.


