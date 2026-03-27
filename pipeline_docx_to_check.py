from TL_docx_to_txt import extract_pairs_from_docx
from proofreading_checker_full import (
    normalize_bilingual_df,
    load_glossary_xlsx,
    run_all_checks,
    export_with_format,
)
import pandas as pd
import os

from ref_sign_checker import process_docx as process_refsign_docx
from ref_sign_checker import detect_column_lang as detect_refsign_columns
from ref_sign_checker import direction_tag as refsign_direction_tag

from docx import Document


def run_ref_sign_checker(in_docx_path: str) -> tuple[str, str]:
    """
    Run ref_sign_checker on the same input DOCX, producing a highlighted DOCX next to it.

    Returns:
        (out_docx_path, direction_tag) where direction_tag is 'JP2EN' or 'EN2JP'
    """
    # Detect direction for output naming (same as ref_sign_checker CLI)
    doc = Document(in_docx_path)
    jp_col, en_col = detect_refsign_columns(doc, max_rows=50)
    tag = refsign_direction_tag(jp_col, en_col)

    out_docx_path = os.path.splitext(in_docx_path)[0] + f"_checked_{tag}.docx"
    process_refsign_docx(in_docx_path, out_docx_path)
    return out_docx_path, tag


def main():
    docx_path = input("DOCX path: ").strip().strip('"')
    glossary_path = input("Glossary xlsx path: ").strip().strip('"')

    if not os.path.isfile(docx_path):
        print("DOCX file not found.")
        return
    if not os.path.isfile(glossary_path):
        print("Glossary file not found.")
        return

    # 1) DOCX → in-memory pairs (for proofreading checks)
    pairs = extract_pairs_from_docx(docx_path)
    df_raw = pd.DataFrame(pairs, columns=["col0", "col1"])

    # Detect JP/EN column roles (rough)
    df, doc_direction = normalize_bilingual_df(df_raw)
    print(f"Detected document direction (rough): {doc_direction}")

    # 2) Load glossary and detect its direction (JP2EN or EN2JP)
    glossary, glossary_direction = load_glossary_xlsx(glossary_path)
    print(f"Detected glossary direction: {glossary_direction}")

    # 3) Proofreading checks using glossary_direction
    out_df = run_all_checks(df, glossary, glossary_direction)

    # 4) Export Excel next to the original DOCX
    out_file = export_with_format(out_df, docx_path, glossary_direction)
    print(f"Done. Proofreading result: {out_file}")

    # 5) Also run ref_sign_checker on the same DOCX (outputs a checked DOCX)
    checked_docx, refsign_tag = run_ref_sign_checker(docx_path)
    print(f"Done. Ref-sign checked DOCX ({refsign_tag}): {checked_docx}")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        import traceback

        traceback.print_exc()
        input("\n[ERROR] Press Enter to close...")
    else:
        input("\n[OK] Press Enter to close...")
