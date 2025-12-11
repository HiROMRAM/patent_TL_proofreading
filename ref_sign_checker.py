from docx import Document
from docx.shared import RGBColor
import unicodedata
import re
from collections import Counter
import os

# ========== 1. Constants / Patterns ==========

EN_NUM_MAP = {
    "1st": "1", "2nd": "2", "3rd": "3", "4th": "4", "5th": "5",
    "6th": "6", "7th": "7", "8th": "8", "9th": "9",
    "first": "1", "second": "2", "third": "3", "fourth": "4", "fifth": "5",
    "sixth": "6", "seventh": "7", "eighth": "8", "ninth": "9",
    "one": "1", "two": "2", "three": "3", "four": "4", "five": "5",
    "six": "6", "seven": "7", "eight": "8", "nine": "9", "ten": "10",
    "zero": "0",
}

# EN tokens
pat_en_1 = re.compile(r"[A-Za-z0-9Α-Ωα-ω]*[0-9Α-Ωα-ω]+[A-Za-z0-9Α-Ωα-ω]*")
pat_en_2 = re.compile(r"\b[A-Z_]{2,}\b|(?:(?<=\W)|^)[A-Z_](?=\W)")
pat_en_num_commas = re.compile(r"\d{1,3}(?:,\d{3})+")
pat_en_caps_plural = re.compile(r"\b([A-Z_]{2,})s\b")

# JP tokens
pat_jp_num = re.compile(r"\d+")
pat_jp_alnum = re.compile(r"[A-Za-z0-9Α-Ωα-ω]*[0-9Α-Ωα-ω]+[A-Za-z0-9Α-Ωα-ω]*")
pat_jp_alpha = re.compile(r"[A-Za-zΑ-Ωα-ω_]+")
pat_jp_zero = re.compile(r"ゼロ")


# ========== 2. Utilities ==========

def nfkc(s: str) -> str:
    return unicodedata.normalize("NFKC", s or "")


def strip_num_commas(s: str) -> str:
    return (s or "").replace(",", "")


def first_alpha_match(text: str):
    m = re.search(r"[A-Za-z]+", text)
    return (m.group(0), m.span()) if m else (None, None)


# ========== 3. TL direction detection ==========

def is_japanese_char(ch: str) -> bool:
    code = ord(ch)
    if 0x3040 <= code <= 0x309F:
        return True  # Hiragana
    if 0x30A0 <= code <= 0x30FF:
        return True  # Katakana
    if 0x4E00 <= code <= 0x9FFF:
        return True  # Kanji
    if 0xFF00 <= code <= 0xFFEF:
        return True  # Full-width forms
    return False


def jp_score(text: str) -> float:
    chars = [c for c in text if not c.isspace()]
    if not chars:
        return 0.0
    jp = sum(1 for c in chars if is_japanese_char(c))
    return jp / len(chars)


def en_score(text: str) -> float:
    chars = [c for c in text if not c.isspace()]
    if not chars:
        return 0.0
    en = sum(1 for c in chars if ("A" <= c <= "Z") or ("a" <= c <= "z"))
    return en / len(chars)


def detect_column_lang(doc: Document, max_rows=50):
    for tbl in doc.tables:
        col_scores = [{"jp": 0.0, "en": 0.0}, {"jp": 0.0, "en": 0.0}]
        checked = 0

        for row in tbl.rows:
            if len(row.cells) < 2:
                continue

            t0 = nfkc(row.cells[0].text)
            t1 = nfkc(row.cells[1].text)

            col_scores[0]["jp"] += jp_score(t0)
            col_scores[0]["en"] += en_score(t0)
            col_scores[1]["jp"] += jp_score(t1)
            col_scores[1]["en"] += en_score(t1)

            checked += 1
            if checked >= max_rows:
                break

        if checked > 0:
            jp_col = 0 if col_scores[0]["jp"] >= col_scores[1]["jp"] else 1
            en_col = 1 - jp_col
            return jp_col, en_col

    # fallback
    return 0, 1


def direction_tag(jp_col, en_col):
    return "JP2EN" if jp_col < en_col else "EN2JP"


# ========== 4. JP Tokenizer ==========

def tokenize_jp(text: str):
    text = nfkc(text)
    tokens = []
    seen = set()

    # ゼロ → 0
    for m in pat_jp_zero.finditer(text):
        span = (m.start(), m.end())
        if span in seen:
            continue
        seen.add(span)
        tokens.append({
            "surface": m.group(0),
            "norm": "0",
            "start": m.start(),
            "end": m.end(),
            "cls": "jp_number",
        })

    # numbers with commas
    for m in pat_en_num_commas.finditer(text):
        span = (m.start(), m.end())
        if span in seen:
            continue
        seen.add(span)
        tok = m.group(0)
        tokens.append({
            "surface": tok,
            "norm": strip_num_commas(tok),
            "start": m.start(),
            "end": m.end(),
            "cls": "jp_number",
        })

    # plain numbers
    for m in pat_jp_num.finditer(text):
        span = (m.start(), m.end())
        if span in seen:
            continue
        seen.add(span)
        tok = m.group(0)
        tokens.append({
            "surface": tok,
            "norm": strip_num_commas(tok),
            "start": m.start(),
            "end": m.end(),
            "cls": "jp_number",
        })

    # alnum with digits (S10, RAM100, 5A, etc.)
    for m in pat_jp_alnum.finditer(text):
        span = (m.start(), m.end())
        if span in seen:
            continue
        seen.add(span)
        tok = m.group(0)
        tokens.append({
            "surface": tok,
            "norm": strip_num_commas(tok),
            "start": m.start(),
            "end": m.end(),
            "cls": "jp_alnum_digit",
        })

    # roman alpha in JP text
    for m in pat_jp_alpha.finditer(text):
        span = (m.start(), m.end())
        if span in seen:
            continue
        seen.add(span)
        tok = m.group(0)
        tokens.append({
            "surface": tok,
            "norm": tok,
            "start": m.start(),
            "end": m.end(),
            "cls": "jp_alpha",
        })

    return tokens


def mark_embedded_numbers(jp_tokens):
    """
    Mark jp_number tokens that are embedded in jp_alnum_digit spans
    (e.g. '10' inside 'S10').
    """
    alnum_spans = [
        (t["start"], t["end"])
        for t in jp_tokens
        if t["cls"] == "jp_alnum_digit"
    ]
    for t in jp_tokens:
        t["embedded"] = False
        if t["cls"] == "jp_number":
            s, e = t["start"], t["end"]
            for a_s, a_e in alnum_spans:
                if a_s <= s and e <= a_e:
                    t["embedded"] = True
                    break


# ========== 5. EN Tokenizer ==========

def tokenize_en(text: str):
    text = nfkc(text)
    tokens = []
    seen = set()

    first_alpha, first_span = first_alpha_match(text)

    # plural ALLCAPS like APIs → normalize to API
    for m in pat_en_caps_plural.finditer(text):
        span = (m.start(), m.end())
        if span in seen:
            continue
        seen.add(span)
        base = m.group(1)
        surface = m.group(0)
        tokens.append({
            "surface": surface,
            "norm": base,
            "start": m.start(),
            "end": m.end(),
            "cls": "en_caps_plural",
        })

    # numbers with commas
    for m in pat_en_num_commas.finditer(text):
        span = (m.start(), m.end())
        if span in seen:
            continue
        seen.add(span)
        tok = m.group(0)
        tokens.append({
            "surface": tok,
            "norm": strip_num_commas(tok),
            "start": m.start(),
            "end": m.end(),
            "cls": "en_alnum_digit",
        })

    # alnum with digits
    for m in pat_en_1.finditer(text):
        span = (m.start(), m.end())
        if span in seen:
            continue
        seen.add(span)
        tok = m.group(0)
        tokens.append({
            "surface": tok,
            "norm": strip_num_commas(tok),
            "start": m.start(),
            "end": m.end(),
            "cls": "en_alnum_digit",
        })

    # ALLCAPS tokens
    for m in pat_en_2.finditer(text):
        span = (m.start(), m.end())
        tok = m.group(0)
        # skip leading 'A' if it's the first alpha in the text
        if tok == "A" and first_alpha == "A" and first_span == span:
            continue
        if span in seen:
            continue
        seen.add(span)
        tokens.append({
            "surface": tok,
            "norm": tok,
            "start": m.start(),
            "end": m.end(),
            "cls": "en_caps",
        })

    # English number words / ordinals
    word_pat = re.compile(r"\b[a-zA-Z]+\b")
    for m in word_pat.finditer(text):
        w = m.group(0).lower()
        if w in EN_NUM_MAP:
            span = (m.start(), m.end())
            if span in seen:
                continue
            seen.add(span)
            tokens.append({
                "surface": m.group(0),
                "norm": EN_NUM_MAP[w],
                "start": m.start(),
                "end": m.end(),
                "cls": "en_num_word",
            })

    return tokens


def counter_from_tokens(tokens):
    return Counter(t["norm"] for t in tokens)


# ========== 6. Alpha cross-match ==========

def find_alpha_in_en(alpha: str, en_text: str):
    """
    Word-boundary search for alpha in EN sentence (case-insensitive).
    """
    pat = r"(?<![A-Za-z0-9])" + re.escape(alpha) + r"(?![A-Za-z0-9])"
    return re.search(pat, en_text, flags=re.I)


def cross_match_jp_alpha(jp_tokens, en_text):
    """
    For JP alpha tokens that are NOT all-caps:
      - if found as a word in EN, treat as shared
      - otherwise as JP-only
    """
    jp_shared, jp_only, en_shared = [], [], []

    for t in jp_tokens:
        if t["cls"] != "jp_alpha":
            continue
        tok = t["surface"]
        if not tok:
            continue
        if tok.isupper():  # CAPS handled elsewhere
            continue

        m = find_alpha_in_en(tok, en_text)
        if m:
            jp_shared.append((t["start"], t["end"]))
            en_shared.append((m.start(), m.end()))
        else:
            jp_only.append((t["start"], t["end"]))

    return jp_shared, jp_only, en_shared


# ========== 7. Evidence-based split for alnum tokens ==========

def evidence_split_jp_alnum(
    jp_tokens, en_tokens, en_text,
    shared_spans_jp, only_spans_jp,
    shared_spans_en
):
    """
    Evidence-based split of alnum-digit tokens:

      - For JP alnum-digit tokens like:
          10S, S10, RAM100, etc.
      - Check whether there is "evidence" on EN side (numbers / alpha).
      - If yes, move corresponding spans from JP-only → shared.
    """
    en_norms = set(t["norm"] for t in en_tokens)
    to_remove = set()

    def add_shared(sp):
        if sp not in shared_spans_jp:
            shared_spans_jp.append(sp)

    for t in jp_tokens:
        if t["cls"] != "jp_alnum_digit":
            continue

        s_full = t["surface"]
        base = t["start"]
        full_sp = (t["start"], t["end"])

        # Case A: 10S → '10' + 'S'
        mA = re.match(r"^(\d+)([A-Za-zΑ-Ωα-ω]+)$", s_full)
        if mA:
            num, alpha = mA.group(1), mA.group(2)
            if num in en_norms:
                m2 = find_alpha_in_en(alpha, en_text)
                if m2:
                    alpha_jp = (base + len(num), t["end"])
                    add_shared(alpha_jp)
                    shared_spans_en.append((m2.start(), m2.end()))
                    to_remove.add(full_sp)
                    to_remove.add(alpha_jp)
            continue

        # Case B: S10 → 'S' + '10'
        mB = re.match(r"^([A-Za-zΑ-Ωα-ω]+)(\d+)$", s_full)
        if mB:
            alpha, num = mB.group(1), mB.group(2)
            if alpha in en_norms and num in en_norms:
                alpha_sp = (base, base + len(alpha))
                num_sp = (base + len(alpha), t["end"])
                add_shared(alpha_sp)
                add_shared(num_sp)
                to_remove.update([full_sp, alpha_sp, num_sp])
            continue

    if to_remove:
        only_spans_jp[:] = [sp for sp in only_spans_jp if sp not in to_remove]


# ========== 8. Drop mismatch spans fully covered by shared spans ==========

def filter_inner_only(only_spans, shared_spans):
    """
    Remove mismatch spans that are completely covered by a shared span.
    This emulates the behavior where matched spans override mismatches inside them.

    """
    if not only_spans or not shared_spans:
        return only_spans

    filtered = []
    for s, e in only_spans:
        contained = False
        for S, E in shared_spans:
            if S <= s and e <= E:
                contained = True
                break
        if not contained:
            filtered.append((s, e))
    return filtered


# ========== 9. Highlight mismatches only (red + bold) ==========

def highlight_span_in_cell(cell, spans):
    """
    spans: list of (start, end) in NFKC-normalized coordinates.
    Only those spans are highlighted as mismatches:
      - bold
      - red
    All other text stays plain.
    """
    if not spans:
        return
    text = cell.text
    if not text:
        return

    para = cell.paragraphs[0]

    # Clear existing runs
    for r in list(para.runs):
        r.clear()
        r.text = ""

    # Sort and merge overlapping/adjacent spans
    spans_sorted = sorted(set(spans), key=lambda x: (x[0], x[1]))
    merged = []
    for s, e in spans_sorted:
        if not merged or s > merged[-1][1]:
            merged.append([s, e])
        else:
            merged[-1][1] = max(merged[-1][1], e)

    pos = 0
    for s, e in merged:
        if pos < s:
            para.add_run(text[pos:s])  # normal text
        run = para.add_run(text[s:e])
        run.font.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
        pos = e

    if pos < len(text):
        para.add_run(text[pos:])


# ========== 10. Main processing ==========

def process_docx(in_path, out_path):
    doc = Document(in_path)

    # JP / EN columns
    jp_col, en_col = detect_column_lang(doc, max_rows=50)

    for tbl in doc.tables:
        rows = list(tbl.rows)
        rows_to_delete = []

        for row in rows:
            if len(row.cells) < 2:
                continue

            jp_cell = row.cells[jp_col]
            en_cell = row.cells[en_col]

            jp_text = nfkc(jp_cell.text)
            en_text = nfkc(en_cell.text)

            jp_t = tokenize_jp(jp_text)
            en_t = tokenize_en(en_text)
            mark_embedded_numbers(jp_t)

            # Base counts
            jp_cnt = counter_from_tokens(jp_t)
            en_cnt = counter_from_tokens(en_t)

            shared = jp_cnt & en_cnt
            jp_only_cnt = jp_cnt - shared
            en_only_cnt = en_cnt - shared

            # ----- Build JP spans (shared / only) -----
            shared_jp, only_jp = [], []
            tmp_s = shared.copy()
            tmp_jp = jp_only_cnt.copy()

            # shared_jp: non-embedded first
            for t in jp_t:
                if t.get("embedded"):
                    continue
                n = t["norm"]
                if tmp_s.get(n, 0) > 0:
                    shared_jp.append((t["start"], t["end"]))
                    tmp_s[n] -= 1

            # shared_jp: then embedded
            for t in jp_t:
                if not t.get("embedded"):
                    continue
                n = t["norm"]
                if tmp_s.get(n, 0) > 0:
                    shared_jp.append((t["start"], t["end"]))
                    tmp_s[n] -= 1

            # only_jp spans
            for t in jp_t:
                n = t["norm"]
                if jp_only_cnt.get(n, 0) > 0:
                    only_jp.append((t["start"], t["end"]))
                    jp_only_cnt[n] -= 1

            # ----- EN spans (shared / only) -----
            tmp_s2 = shared.copy()
            tmp_en = en_only_cnt.copy()
            shared_en, only_en = [], []

            for t in en_t:
                n = t["norm"]
                if tmp_s2.get(n, 0) > 0:
                    shared_en.append((t["start"], t["end"]))
                    tmp_s2[n] -= 1
                elif tmp_en.get(n, 0) > 0:
                    only_en.append((t["start"], t["end"]))
                    tmp_en[n] -= 1

            # ----- JP alpha cross-match (Mg, Ca, etc.) -----
            jp_alpha_s, jp_alpha_o, en_alpha_s = cross_match_jp_alpha(jp_t, en_text)

            # Overwrite semantics: shared wins over mismatch
            #   - remove shared alpha spans from only_jp
            only_jp = [sp for sp in only_jp if sp not in jp_alpha_s]
            shared_jp.extend(jp_alpha_s)

            # unmatched alpha
            only_jp.extend(jp_alpha_o)

            # EN shared alpha spans
            shared_en.extend(en_alpha_s)

            # ----- Alnum split (S10 / RAM100, etc.) -----
            evidence_split_jp_alnum(
                jp_t, en_t, en_text,
                shared_jp, only_jp,
                shared_en
            )

            # ----- Drop mismatch spans that lie fully inside shared spans -----
            only_jp = filter_inner_only(only_jp, shared_jp)
            only_en = filter_inner_only(only_en, shared_en)

            # ----- Final decision: mismatches? -----
            # only_jp / only_en after all overwrites & splits & filtering
            has_mismatch = bool(only_jp or only_en)
            if not has_mismatch:
                # perfect row -> remove later
                rows_to_delete.append(row)
                continue

            # ----- Highlight mismatches only (red + bold) -----
            highlight_span_in_cell(jp_cell, only_jp)
            highlight_span_in_cell(en_cell, only_en)

        # Remove rows with no mismatches
        for row in rows_to_delete:
            tbl._tbl.remove(row._tr)

    doc.save(out_path)


# ========== 11. CLI ==========

def main():
    path = input("DOCX path: ").strip().strip('"')
    if not os.path.isfile(path):
        print("File not found.")
        return

    # detect direction BEFORE processing
    doc = Document(path)
    jp_col, en_col = detect_column_lang(doc, max_rows=50)
    tag = direction_tag(jp_col, en_col)

    out = os.path.splitext(path)[0] + f"_checked_{tag}.docx"
    process_docx(path, out)

    print(f"Detected TL direction: {tag}")
    print("written ->", out)


if __name__ == "__main__":
    main()
