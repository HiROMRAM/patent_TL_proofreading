📘 JP–EN Patent Proofreading Tools

Standalone quality-checking tools for Japanese–English patent translation.
All sample files are synthetic—no real patent content is included.

🔍 Features
1. Reference Sign Checker (ref_sign_checker.py)

Validates bidirectional JP↔EN reference-sign consistency in bilingual DOCX tables

Detects mismatches at token level

Outputs a mismatch-only DOCX with highlighted differences

2. Translation Proofreading Pipeline (pipeline_docx_to_check.py)

Bidirectional glossary validation (JP→EN & EN→JP)

Word repetition detection (EN)

Spacing checks

Outputs an Excel file with:

All — every JP/EN pair + issue summary

IssuesOnly — rows where at least one issue was found

⚡ Quick Start
pip install -r requirements.txt
python -m spacy download en_core_web_sm

Reference-sign checking
python ref_sign_checker.py


Prompt example:

Enter DOCX path: examples/sample_bilingual.docx

Proofreading pipeline
python pipeline_docx_to_check.py


Prompts:

DOCX path: examples/sample_bilingual.docx
Glossary path: examples/sample_glossary.xlsx

📂 Repository Structure
patent_TL_proofreading/
├── ref_sign_checker.py          # Reference-sign validation (DOCX → DOCX)
├── proofreading_checker_full.py # Core proofreading logic
├── TL_docx_to_txt.py            # DOCX → JP/EN text extraction
├── pipeline_docx_to_check.py    # End-to-end proofreading pipeline
├── examples/
│   ├── sample_bilingual.docx
│   └── sample_glossary.xlsx
├── requirements.txt
└── README.md

🧪 Usage Details
Reference Sign Checker

Input: bilingual DOCX with JP/EN in table columns

Output: new DOCX containing only mismatched rows, highlighted

Direction: inherently bidirectional (JP↔EN)

Proofreading Pipeline

Extracts JP/EN text pairs → runs all checks → exports Excel

The output sheets contain:

JP text

EN text

Issue summary (glossary mismatches, repetition, spacing, etc.)

Designed to be run independently from the reference-sign checker

⚙️ Key Design Notes

Bidirectional checks: both tools support JP→EN and EN→JP scenarios

Standalone scripts: no tool auto-invokes the other

Demo-focused: intended for workflow experimentation

Safe: examples are minimal and synthetic

🚧 Limitations

English repetition detection only

Japanese linguistic handling is intentionally minimal

Not optimized for large production-scale documents

📈 Future Improvements

Unified CLI (e.g., python qc.py --all)

Package structure (pip install)

Improved Japanese morphological analysis

Optional integration with CAT tool APIs