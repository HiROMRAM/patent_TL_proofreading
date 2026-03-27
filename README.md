# JP–EN Patent Proofreading Tools

Rule-based quality-checking tools for Japanese–English patent translation.

## Setup
```
pip install -r requirements.txt
python -m spacy download en_core_web_sm
python -c "import nltk; nltk.download('wordnet')"
```

## Tools

### Reference Sign Checker (`ref_sign_checker.py`)

Checks reference-sign consistency between JP and EN columns in a bilingual DOCX table. Outputs a new DOCX with mismatched rows highlighted.

```
python ref_sign_checker.py
```

### Proofreading Checker (`proofreading_checker_full.py`)

English-side checks (JP2EN only):
- Consecutive word repetition
- Double spaces
- Space before punctuation
- a/an article correctness (via `inflect`, with acronym handling)
- Bare verb after "for" (e.g., "for detect" → "for detecting"), using WordNet + patent-domain allowlist
- Missing final period when JP source ends with 。
- Subject-verb agreement (output to separate sheet due to low accuracy)

Glossary compliance (both directions):
- Term matching with noun-lemma fallback (e.g., "trajectories" matches "trajectory")

Input: bilingual TXT (tab-separated) + glossary XLSX.
Output: Excel with three sheets (All, IssuesOnly, SV_Agreement).

### Pipeline (`pipeline_docx_to_check.py`)

Runs both the proofreading checker and reference sign checker on a bilingual DOCX in one step.

```
python pipeline_docx_to_check.py
```

Prompts for a DOCX path and a glossary XLSX path. Translation direction is auto-detected from character-ratio scoring.

### DOCX Extractor (`TL_docx_to_txt.py`)

Extracts JP/EN text pairs from a 2-column DOCX table to a tab-separated TXT file.

## Repository Structure
```
├── ref_sign_checker.py
├── proofreading_checker_full.py
├── TL_docx_to_txt.py
├── pipeline_docx_to_check.py
├── requirements.txt
└── README.md
```

## Limitations

- No Japanese-side linguistic checks (grammar, morphology, etc.)
- SV agreement detection has limited accuracy
