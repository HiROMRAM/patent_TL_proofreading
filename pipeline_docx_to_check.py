from TL_docx_to_txt import extract_pairs_from_docx
from proofreading_checker_full import (
    normalize_bilingual_df,
    load_glossary_xlsx,
    run_all_checks,
    export_with_format,
)
import pandas as pd


def main():
    docx_path = input("DOCX path: ").strip().strip('"')
    glossary_path = input("Glossary xlsx path: ").strip().strip('"')

    # 1) DOCX → in-memory pairs
    pairs = extract_pairs_from_docx(docx_path)
    df_raw = pd.DataFrame(pairs, columns=["col0", "col1"])

    # Detect JP/EN column roles (similar idea to ref_sign_checker)
    df, doc_direction = normalize_bilingual_df(df_raw)
    print(f"Detected document direction (rough): {doc_direction}")

    # 2) Load glossary and detect its direction (JP2EN or EN2JP)
    glossary, glossary_direction = load_glossary_xlsx(glossary_path)
    print(f"Detected glossary direction: {glossary_direction}")

    # 3) Proofreading checks using glossary_direction
    out_df = run_all_checks(df, glossary, glossary_direction)

    # 4) Export Excel next to the original DOCX
    #    File name includes glossary_direction: JP2EN or EN2JP
    out_file = export_with_format(out_df, docx_path, glossary_direction)

    print(f"Done. Proofreading result: {out_file}")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        import traceback

        traceback.print_exc()
        input("\n[ERROR] Press Enter to close...")
    else:
        input("\n[OK] Press Enter to close...")
