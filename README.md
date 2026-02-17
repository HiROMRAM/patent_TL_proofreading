# 📘 JP–EN Patent Proofreading Tools

Standalone quality-checking tools for Japanese–English patent translation. All sample files are synthetic—no real patent content is included.

## 🔍 Features

### Reference Sign Checker (`ref_sign_checker.py`)

- Validates bidirectional JP↔EN reference-sign consistency in bilingual DOCX tables
- Detects mismatches at token level
- Outputs a mismatch-only DOCX with highlighted differences

### Translation Proofreading Pipeline (`pipeline_docx_to_check.py`)

- Bidirectional glossary validation (JP→EN & EN→JP)
- Word repetition detection (EN)
- Spacing checks
- Outputs an Excel file with:
  - All — every JP/EN pair + issue summary
  - IssuesOnly — rows where at least one issue was found

### LLM-Based Final QA Check (`patent_tl_final_check.py`)

- Sends translation segments to Claude API for error detection
- Auto-detects translation direction (JP→EN or EN→JP) from the input DOCX
- Checks for: omissions, number/unit mismatches, mistranslations, grammar errors, subject/object confusion
- Separates patent claims from non-claim segments with different batching strategies
- Uses prompt caching for cost reduction
- Distinguishes definite vs uncertain issues in the output
- Outputs an Excel file with All and IssuesOnly sheets

## ⚡ Quick Start

```
pip install -r requirements.txt
python -m spacy download en_core_web_sm
```

### Reference-sign checking

```
python ref_sign_checker.py
```

### Proofreading pipeline

```
python pipeline_docx_to_check.py
```

### LLM final check (requires Anthropic API key)

```
export ANTHROPIC_API_KEY=your_key_here
python patent_tl_final_check.py input.docx
```

Or interactive mode:

```
python patent_tl_final_check.py
```

## 📂 Repository Structure

```
patent_TL_proofreading/
├── ref_sign_checker.py          # Reference-sign validation (DOCX → DOCX)
├── proofreading_checker_full.py # Core proofreading logic
├── TL_docx_to_txt.py            # DOCX → JP/EN text extraction
├── pipeline_docx_to_check.py    # End-to-end proofreading pipeline
├── patent_tl_final_check.py     # LLM-based final QA check (Claude API)
├── examples/
│   ├── sample_bilingual.docx
│   └── sample_glossary.xlsx
├── requirements.txt
└── README.md
```

## 🧪 Usage Details

### Reference Sign Checker

- Input: bilingual DOCX with JP/EN in table columns
- Output: new DOCX containing only mismatched rows, highlighted
- Direction: inherently bidirectional (JP↔EN)

### Proofreading Pipeline

- Extracts JP/EN text pairs → runs all checks → exports Excel
- The output sheets contain:
  - JP text
  - EN text
  - Issue summary (glossary mismatches, repetition, spacing, etc.)
- Designed to be run independently from the reference-sign checker

### LLM Final Check

- Input: bilingual DOCX with 2-column table (source | target)
- Output: Excel with issue annotations per segment
- Non-claim segments: 7-row context window, 5-row check blocks with 2-row overlap
- Claim segments: grouped by 【請求項n】 markers, 3 claims per batch
- Direction auto-detected from character-ratio scoring of the DOCX columns

## ⚙️ Key Design Notes

- Bidirectional checks: all tools support JP→EN and EN→JP scenarios
- Standalone scripts: no tool auto-invokes the other
- Demo-focused: intended for workflow experimentation
- Safe: examples are minimal and synthetic

## 🚧 Limitations

- English repetition detection only (in the rule-based pipeline)
- Japanese linguistic handling is intentionally minimal
- Not optimized for large production-scale documents
- LLM final check requires an Anthropic API key and incurs API costs
