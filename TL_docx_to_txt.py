# requirements: pip install python-docx
from docx import Document
import os
import re


def normalize_cell_text(s: str) -> str:
    # Handle common DOCX flattening of breaks:
    s = s.replace("\r\n", " ").replace("\r", " ").replace("\n", " ")

    # Replace uncommon whitespace marks with standard spaces (safe)
    s = s.replace("\x0b", " ")      # vertical tab
    s = s.replace("\u000b", " ")    # same as above
    s = s.replace("\u2028", " ")    # Unicode line separator
    s = s.replace("\u00A0", " ")    # non-breaking space → normal space

    # Replace tabs with spaces to avoid TSV corruption
    s = s.replace("\t", " ")

    # *** DO NOT collapse whitespace ***
    # return original spacing exactly as seen
    return s


def extract_pairs_from_docx(docx_path: str) -> list[tuple[str, str]]:
    doc = Document(docx_path)
    pairs: list[tuple[str, str]] = []

    for tbl in doc.tables:
        if not tbl.rows:
            continue
        for row in tbl.rows:
            cells = row.cells
            if len(cells) < 2:
                continue
            jp = normalize_cell_text(cells[0].text or "")
            en = normalize_cell_text(cells[1].text or "")
            if jp or en:
                pairs.append((jp, en))
    return pairs


def save_pairs_to_tsv(pairs, out_path: str):
    with open(out_path, "w", encoding="utf-8", newline="") as f:
        for jp, en in pairs:
            f.write(f"{jp}\t{en}\n")


def main():
    in_path = input("Input DOCX path: ").strip().strip('"')
    if not os.path.isfile(in_path):
        print("File not found.")
        return

    default_out = os.path.splitext(in_path)[0] + "_pairs.txt"
    out_path = input(f"Output TXT (TSV) path [{default_out}]: ").strip().strip('"')
    if not out_path:
        out_path = default_out

    pairs = extract_pairs_from_docx(in_path)
    if not pairs:
        print("No pairs found (check the table structure).")
        return

    save_pairs_to_tsv(pairs, out_path)
    print(f"Done. Wrote {len(pairs)} lines → {out_path}")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        import traceback
        traceback.print_exc()
        input("\n[ERROR] Press Enter to close...")
    else:
        input("\n[OK] Press Enter to close...")
