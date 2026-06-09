# ai-compare

Compare answers from multiple LLMs against expected responses defined in an Excel workbook.

Questions and models are configured in spreadsheet sheets; results are written back with grades from Claude Opus 4.8.

## Setup

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
```

Set `APIFY_API_TOKEN` and `ANTHROPIC_API_KEY` in `.env`.

## Usage

```bash
# Create a template workbook
python compare_models.py --create-template comparison.xlsx

# Run a comparison (any workbook with questions/models sheets)
python compare_models.py input.xlsx
```
