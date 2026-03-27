# -*- coding: utf-8 -*-
"""
Full Proofreading Checker for Japanese/English Translation (Bidirectional)
Input: 
 - Bilingual translation file (.txt, tab-separated, 2 columns per line)
 - Glossary file (.xlsx, col A = source language, col B = target language)
Output:
 - Excel file with flagged issues per sentence
   (file name includes detected direction: JP2EN or EN2JP)

Note:
 - Requires the spaCy model 'en_core_web_sm'.
   Install via:  python -m spacy download en_core_web_sm
 - Requires the 'inflect' library for a/an article checking.
   Install via:  pip install inflect
 - Requires 'nltk' with WordNet data for the bare-verb checker.
   Install via:  pip install nltk
   Then once:    python -c "import nltk; nltk.download('wordnet')"
"""

import re
import spacy
import inflect
import pandas as pd
from pathlib import Path
from datetime import datetime
from openpyxl.styles import Alignment
import os
import unicodedata
import nltk
from nltk.corpus import wordnet as wn

# Ensure WordNet data is available
try:
    wn.synsets("test")
except LookupError:
    nltk.download("wordnet", quiet=True)

nlp = spacy.load("en_core_web_sm")
_inflect_engine = inflect.engine()


# ========== 1. Language heuristics (similar to ref_sign_checker) ==========

def is_japanese_char(ch: str) -> bool:
    code = ord(ch)
    # Hiragana, Katakana, CJK, full-width forms (rough but sufficient)
    if 0x3040 <= code <= 0x309F:
        return True  # Hiragana
    if 0x30A0 <= code <= 0x30FF:
        return True  # Katakana
    if 0x4E00 <= code <= 0x9FFF:
        return True  # Kanji
    if 0xFF00 <= code <= 0xFFEF:
        return True  # Full-width forms
    return False


def is_english_char(ch: str) -> bool:
    return ("A" <= ch <= "Z") or ("a" <= ch <= "z")


def language_scores_for_text(s: str) -> tuple[int, int]:
    jp = en = 0
    for ch in s:
        if ch.isspace():
            continue
        if is_japanese_char(ch):
            jp += 1
        elif is_english_char(ch):
            en += 1
    return jp, en


def language_scores_for_series(series: pd.Series, max_rows: int = 50) -> tuple[int, int]:
    parts = []
    for val in series.head(max_rows):
        parts.append(str(val))
    text = "".join(parts)
    return language_scores_for_text(text)


def normalize_bilingual_df(df_raw: pd.DataFrame) -> tuple[pd.DataFrame, str]:
    """
    Given a 2-column DataFrame from a bilingual source, detect which column
    is Japanese and which is English, then return:
      - df with columns ['Japanese', 'English']
      - direction tag 'JP2EN' or 'EN2JP' based on column order
    """
    if df_raw.shape[1] != 2:
        raise ValueError("Expected exactly 2 columns in bilingual data")

    col0, col1 = df_raw.columns
    jp0, en0 = language_scores_for_series(df_raw[col0])
    jp1, en1 = language_scores_for_series(df_raw[col1])

    lang0 = "JP" if jp0 >= en0 else "EN"
    lang1 = "JP" if jp1 >= en1 else "EN"

    if lang0 == "JP" and lang1 == "EN":
        jp_col, en_col = col0, col1
    elif lang0 == "EN" and lang1 == "JP":
        jp_col, en_col = col1, col0
    else:
        # Fallback: assume original order [JP, EN]
        jp_col, en_col = col0, col1

    # Match ref_sign_checker style: JP2EN if JP column index < EN column index
    jp_index = 0 if jp_col == col0 else 1
    en_index = 1 - jp_index
    direction = "JP2EN" if jp_index < en_index else "EN2JP"

    df = pd.DataFrame(
        {
            "Japanese": df_raw[jp_col],
            "English": df_raw[en_col],
        }
    )
    return df, direction


# ========== 2. Loaders & normalizers ==========

def load_bilingual_txt(path: str) -> pd.DataFrame:
    """
    Load a tab-separated bilingual TXT and normalize to ['Japanese', 'English'].
    Direction detection is done but ignored here (mainly for CLI use).
    """
    lines = Path(path).read_text(encoding="utf-8").splitlines()
    pairs = [tuple(line.split("\t")) for line in lines if "\t" in line]
    df_raw = pd.DataFrame(pairs, columns=["col0", "col1"])
    df, _ = normalize_bilingual_df(df_raw)
    return df


def normalize_en(s: str) -> str:
    return str(s).lower()


def normalize_ja(s: str) -> str:
    # NFKC to unify full-/half-width, but DO NOT remove spaces
    s = unicodedata.normalize("NFKC", str(s))
    # Convert full-width space to normal space (keep spaces)
    s = s.replace("\u3000", " ")
    return s


def split_variants(s: str) -> list[str]:
    # Split on half- or full-width semicolon
    return [v.strip() for v in re.split(r"[;；]", str(s)) if v.strip()]


def load_glossary_xlsx(path: str) -> tuple[list[tuple[str, list[str]]], str]:
    """
    Load a glossary Excel file and detect its direction:

      - Column A: source language terms (JP or EN)
      - Column B: target language terms (variants separated by ';' or '；')

    Direction detection:
      - If col A looks JP and col B looks EN → 'JP2EN'
      - If col A looks EN and col B looks JP → 'EN2JP'
      - Otherwise default to 'JP2EN'

    Returns:
      glossary: list of (src_term, [tgt_variants])
      direction: 'JP2EN' or 'EN2JP'
    """
    df = pd.read_excel(path)
    if df.shape[1] < 2:
        raise ValueError("Glossary must have at least two columns (source, target)")

    col0 = df.iloc[:, 0]
    col1 = df.iloc[:, 1]

    jp0, en0 = language_scores_for_series(col0)
    jp1, en1 = language_scores_for_series(col1)

    lang0 = "JP" if jp0 >= en0 else "EN"
    lang1 = "JP" if jp1 >= en1 else "EN"

    if lang0 == "JP" and lang1 == "EN":
        direction = "JP2EN"
        src_series = col0
        tgt_series = col1
    elif lang0 == "EN" and lang1 == "JP":
        direction = "EN2JP"
        src_series = col0
        tgt_series = col1
    else:
        # Fallback: assume JP→EN
        direction = "JP2EN"
        src_series = col0
        tgt_series = col1

    glossary: list[tuple[str, list[str]]] = []
    for src, tgt in zip(src_series, tgt_series):
        if pd.isna(src) or pd.isna(tgt):
            continue
        src_str = str(src)
        tgt_variants = split_variants(tgt)
        if src_str and tgt_variants:
            glossary.append((src_str, tgt_variants))

    return glossary, direction


# ========== 2b. Noun-only lemmatization helpers ==========

def lemmatize_last_noun(term: str) -> str:
    """
    Lemmatize only the last token of an English term, and only if
    spaCy tags it as NOUN.  All other tokens are lowercased as-is.

    Examples:
      'travel trajectories' → 'travel trajectory'
      'control signals'     → 'control signal'
      'signal processing'   → 'signal processing'  (last word not NOUN)
    """
    doc = nlp(term)
    if not doc:
        return term.lower()
    last_tok = doc[-1]
    if last_tok.pos_ == "NOUN":
        parts = [tok.text.lower() for tok in doc[:-1]]
        parts.append(last_tok.lemma_.lower())
        return " ".join(parts)
    return term.lower()


def lemmatize_nouns_in_text(text: str) -> str:
    """
    Lemmatize every token tagged as NOUN in an English text;
    leave all other tokens lowercased as-is.

    Used to normalize the English side of a translation row so that
    a glossary variant like 'trajectory' can match 'trajectories'.
    """
    doc = nlp(text)
    return " ".join(
        tok.lemma_.lower() if tok.pos_ == "NOUN" else tok.text.lower()
        for tok in doc
    )


# ========== 3. Core English-side checks ==========

def check_word_repetition(sentence: str) -> list[str]:
    """
    Detect only *strict* consecutive repetition in English:
      e.g., 'policy policy', 'the   the'
    Ignore cases like 'policy (policy', 'policy, policy', etc.
    """
    if not isinstance(sentence, str):
        return []

    doc = nlp(sentence)
    issues: list[str] = []

    for i in range(len(doc) - 1):
        t1 = doc[i]
        t2 = doc[i + 1]

        if not (t1.is_alpha and t2.is_alpha):
            continue

        if t1.text.lower() != t2.text.lower():
            continue

        between = sentence[t1.idx + len(t1.text) : t2.idx]
        if between.strip() != "":
            continue

        issues.append(f"Consecutive repetition: '{t1.text} {t2.text}'")

    return issues


def check_double_space(sentence: str) -> list[str]:
    """
    Detect runs of two or more consecutive half-width spaces in English.
    - Flags "  ", "   ", "    ", etc.
    - Each run is reported once.
    """
    if not isinstance(sentence, str):
        return []

    issues: list[str] = []
    pattern = re.compile(r" {2,}")

    for m in pattern.finditer(sentence):
        run_len = len(m.group())
        idx = m.start()
        context = sentence[max(0, idx - 10) : idx + run_len + 10].replace("\n", "\\n")
        shown = " " * min(run_len, 4)
        issues.append(
            f"Consecutive spaces ({run_len}): '{shown}' (context: …{context}…)"
        )

    return issues


def check_whitespace_before_punctuation(sentence: str) -> list[str]:
    """
    Flag a half-width space immediately before . , ; : ? !

    Catches common typos like 'the device .' or 'a method , comprising'.
    Skips:
      - Ellipsis-like runs ('...' or '. . .')
      - Decimal numbers ('3 .5' is unlikely but '3.5' wouldn't match anyway)
      - Leading period at start of segment (numbered lists, e.g. '. 1')
    """
    if not isinstance(sentence, str):
        return []

    issues: list[str] = []

    for m in re.finditer(r"(?<=\S) +([.,;:?!])", sentence):
        punct = m.group(1)
        # Grab context: a few chars on each side
        start = max(0, m.start() - 15)
        end = min(len(sentence), m.end() + 10)
        ctx = sentence[start:end].replace("\n", "\\n")
        issues.append(
            f"Space before '{punct}': …{ctx}…"
        )

    return issues


def check_missing_final_period(jp: str, en: str) -> list[str]:
    """
    Flag when the Japanese source ends with '。' but the English target
    does not end with a period (after stripping trailing whitespace).

    Common in JP→EN when the translator restructures a sentence and
    forgets to close with '.'.

    Tolerates other terminal punctuation that could legitimately replace
    a period: ; : ? ! (but not comma or nothing).
    Also tolerates segments that end with a closing paren/bracket/quote
    immediately after one of those punctuation marks.
    """
    if not isinstance(jp, str) or not isinstance(en, str):
        return []

    jp_stripped = jp.rstrip()
    en_stripped = en.rstrip()

    if not jp_stripped or not en_stripped:
        return []

    if not jp_stripped.endswith("。"):
        return []

    # Accept . ; : ? ! as valid terminal punctuation,
    # optionally followed by closing brackets/quotes.
    if re.search(r'[.;:?!][)"\'\]）"』】]*$', en_stripped):
        return []

    return [
        f"Missing final period: source ends with '。' "
        f"but target ends with '{en_stripped[-1]}'"
    ]


# ========== 3c. "for + bare verb" checker ==========

def _naive_gerund(lemma: str) -> str:
    """Best-effort gerund from a verb lemma. Covers common spelling rules."""
    if lemma.endswith("ie"):                    # die → dying
        return lemma[:-2] + "ying"
    if lemma.endswith("e") and not lemma.endswith("ee"):  # generate → generating
        return lemma[:-1] + "ing"
    return lemma + "ing"


# C: Hardcoded allowlist of verb/noun homographs common in patent English.
# Words in this set can legitimately follow "for" as nouns:
#   "for control of", "for use in", "for transfer of", etc.
_PATENT_NOUN_VERB_HOMOGRAPHS = {
    "access", "account", "address", "advance", "aid",
    "balance", "benefit", "block", "bridge", "broadcast",
    "cache", "capture", "cause", "challenge", "change", "charge",
    "claim", "code", "command", "contact", "control",
    "copy", "cost", "count", "cover", "cure",
    "damage", "deal", "decay", "decrease", "delay", "demand",
    "design", "desire", "discharge", "discount", "display",
    "drain", "drift", "drive", "drop",
    "effect", "end", "escape", "estimate", "exchange",
    "experience", "export",
    "fall", "fault", "feed", "file", "filter", "finance",
    "flow", "focus", "force", "form", "function",
    "gain", "grip", "group", "guard", "guide",
    "handle", "harvest", "heat", "help", "hold", "host",
    "impact", "import", "increase", "influence", "input",
    "interest", "issue",
    "judge", "jump",
    "lack", "lead", "license", "lift", "limit", "link",
    "load", "lock", "look", "loop", "loss",
    "manufacture", "map", "mark", "match", "measure",
    "merge", "mix", "model", "monitor", "move",
    "need", "note", "notice",
    "offer", "offset", "order", "output", "overlap", "override",
    "permit", "picture", "pitch", "place", "plan", "plant",
    "play", "plot", "point", "position", "power",
    "practice", "press", "process", "produce", "program",
    "promise", "proof", "purchase", "push", "purpose",
    "range", "rate", "reach", "record", "reference",
    "reform", "regard", "release", "relief", "repair",
    "report", "request", "reserve", "result", "return",
    "review", "rise", "roll", "run",
    "sample", "scale", "search", "select", "sense", "sequence",
    "service", "set", "shape", "share", "shift", "signal",
    "slice", "sort", "source", "split", "spread", "stage",
    "start", "step", "stop", "store", "stream", "strike",
    "structure", "study", "supply", "support", "survey", "switch",
    "target", "test", "trade", "transfer", "transform",
    "transport", "trigger", "trim", "turn",
    "update", "upgrade", "use",
    "value", "view", "visit",
    "watch", "work", "wrap",
}


def _has_wordnet_pos(word: str, pos: str) -> bool:
    """Check if the word has any synsets for the given POS in WordNet."""
    return len(wn.synsets(word, pos=pos)) > 0


def check_for_bare_verb(sentence: str) -> list[str]:
    """
    Detect 'for + bare verb' errors common in JP→EN patent translation.

    Japanese ～するための maps to 'for [verb]-ing' in English, but
    translators often produce 'for [base verb]' or 'for [verb]-s'.

    Error examples:
      'for perform processing'     → 'for performing processing'
      'for generate a signal'      → 'for generating a signal'
      'for performs the method'     → 'for performing the method'
      'a method for detect errors'  → 'a method for detecting errors'

    Correct (not flagged):
      'for manufacture of'     — 'manufacture' is a noun
      'for control purposes'   — 'control' is a noun
      'for use in the method'  — 'use' is a noun
      'for performing the step' — gerund is correct
      'for each element'       — 'each' is a determiner
      'for improved accuracy'  — past participle as adjective
      'for further analysis'   — adjective

    Strategy:
      spaCy is used only for tokenization and lemmatization.
      The verb/noun distinction is handled entirely by lexicon
      lookups, avoiding spaCy's unreliable contextual POS tags:

      1. Skip gerunds (-ing) and past forms (-ed) by surface form.
      2. Skip words that have NO verb reading in WordNet (pure
         nouns, determiners, adjectives, etc.).
      3. C (allowlist): skip known patent verb/noun homographs.
      4. B (WordNet noun): skip words that can be a noun.
      5. Skip words that can be an adjective (e.g. "further").
      6. Whatever remains is a pure verb → flag.
    """
    if not isinstance(sentence, str):
        return []

    doc = nlp(sentence)
    issues: list[str] = []

    for i, tok in enumerate(doc):
        if tok.lower_ != "for":
            continue

        # Skip "for" used as a conjunction
        # (rare in patents, but e.g. "..., for the device operates ...")
        if tok.dep_ in ("cc", "mark"):
            continue

        if i + 1 >= len(doc):
            continue

        nxt = doc[i + 1]

        # Only care about alphabetic tokens
        if not nxt.is_alpha:
            continue

        token_text = nxt.text.lower()
        lemma = nxt.lemma_.lower()

        # --- Skip gerunds and past participles by surface form ---
        # "performing" → lemma "perform": -ing form, correct usage → skip
        # "improved"   → lemma "improve": -ed form, participle adj → skip
        # "string"     → lemma "string": ends in -ing but lemma does too → don't skip
        # "speed"      → lemma "speed": ends in -ed? no → don't skip
        if token_text.endswith("ing") and not lemma.endswith("ing"):
            continue
        if token_text.endswith("ed") and not lemma.endswith("ed"):
            continue

        # --- Does this word have any verb reading at all? ---
        # Words with no verb synsets can't be bare-verb errors:
        # "each", "example", "optimal", "the", "further" (if no verb sense), etc.
        if not _has_wordnet_pos(lemma, wn.VERB):
            continue

        # --- C: allowlist of known verb/noun homographs ---
        if lemma in _PATENT_NOUN_VERB_HOMOGRAPHS:
            continue

        # --- B: WordNet noun check ---
        if _has_wordnet_pos(lemma, wn.NOUN):
            continue

        # --- WordNet adjective check ---
        # Catches words like "further" (verb + adj but not noun)
        if _has_wordnet_pos(lemma, wn.ADJ):
            continue

        # Passed all checks — this is a pure verb after "for" → flag.
        issues.append(
            f"Bare verb after 'for': 'for {nxt.text}' "
            f"→ consider 'for {_naive_gerund(lemma)}'"
        )

    return issues


# ========== 3b. Indefinite article (a/an) checker ==========

# Letters whose English *name* starts with a vowel sound:
#   A(ay) E(ee) F(eff) H(aitch) I(eye) L(ell) M(em) N(en)
#   O(oh) R(ar) S(ess) X(eks)
_VOWEL_SOUND_LETTERS = set("AEFHILMNORSX")

# Regex to capture "a" or "an" immediately followed by a word.
# Uses a word boundary before the article and expects whitespace before the
# next word.  The look-behind (?<![A-Za-z]) prevents matching the tail of
# words like "extra" or "than".
_A_AN_RE = re.compile(
    r"(?<![A-Za-z])\b(a|an)\s+([A-Za-z][\w-]*)",
    re.IGNORECASE,
)


def _is_acronym(word: str) -> bool:
    """
    Heuristic: treat a token as a spelled-out acronym if it is
    ≥2 characters and all-uppercase (digits allowed: 'H264', '3GPP'
    are borderline but rare after "a/an").
    Hyphenated forms like 'N-BOX' also count.
    """
    core = word.replace("-", "")
    return len(core) >= 2 and core.isascii() and core.isupper()


def _expected_article(word: str) -> str:
    """
    Return 'a' or 'an' — the article that should precede *word*.

    For acronyms (all-uppercase), uses the English letter-name
    pronunciation of the first letter, because inflect sometimes
    tries to pronounce them as regular words (giving 'a FEC'
    instead of 'an FEC', etc.).

    For all other words, delegates to inflect.engine().a(), which
    handles phonetic edge cases: silent-h (hour), /juː/ onsets
    (university, unit), and more.
    """
    if _is_acronym(word):
        first = word.lstrip("-")[0].upper() if word.lstrip("-") else ""
        return "an" if first in _VOWEL_SOUND_LETTERS else "a"

    # inflect.a() returns e.g. "an element" or "a unit"
    result = _inflect_engine.a(word)
    return result.split()[0].lower()


def check_a_an(sentence: str) -> list[str]:
    """
    Detect incorrect indefinite article usage ('a' vs 'an').

    Uses the *inflect* library for regular words (phonetic-aware:
    handles silent-h, /juː/-onset, etc.) and a letter-name lookup
    for all-uppercase acronyms common in patent text.

    Skips:
      - Tokens that start with a digit or non-alpha character
      - Parenthesised insertions like 'a (first) element' where 'a'
        actually governs a later noun — these are too ambiguous to
        flag reliably.
    """
    if not isinstance(sentence, str):
        return []

    issues: list[str] = []

    for m in _A_AN_RE.finditer(sentence):
        article_used = m.group(1).lower()
        next_word = m.group(2)

        # Skip if the very next char after the article is '(' — the
        # real governed noun may be further away.
        between_start = m.start(1) + len(m.group(1))
        between = sentence[between_start : m.start(2)]
        if "(" in between:
            continue

        expected = _expected_article(next_word)

        if article_used != expected:
            issues.append(
                f"Article: '{m.group(1)} {next_word}' → "
                f"should be '{expected} {next_word}'"
            )

    return issues


# ========== 4. Subject-verb number agreement checker ==========

# Verbs where spaCy's Number morph is unreliable or not applicable
_MODAL_VERBS = {"can", "could", "may", "might", "shall", "should", "will", "would", "must"}


def _find_true_subject_number(subj_token) -> str | None:
    """
    Walk from the nsubj token to find the morphological Number
    of the *real* grammatical subject head.

    For relative pronouns ('that', 'which', 'who') acting as nsubj,
    we return None since agreement traces back to the antecedent.
    """
    # Relative pronouns — skip to avoid spurious flags
    if subj_token.pos_ == "PRON" and subj_token.lower_ in (
        "that", "which", "who", "whose", "whom"
    ):
        return None

    # Determiners/pronouns that are semantically singular
    if subj_token.lower_ in ("each", "every", "one", "none", "neither", "either"):
        return "Sing"

    number = subj_token.morph.get("Number")
    if number:
        return number[0]  # "Sing" or "Plur"
    return None


def _find_verb_number(verb_token) -> str | None:
    """
    Determine the morphological number of the finite verb.

    For auxiliary chains ('does not comprise'), number marking is on
    the auxiliary.  We check the verb first, then its aux children.

    Returns None for modals, infinitives, past tense (except was/were),
    where English does not mark number.
    """
    if verb_token.lower_ in _MODAL_VERBS:
        return None

    tense = verb_token.morph.get("Tense")
    # Past tense verbs don't show number (except was/were)
    if tense == ["Past"] and verb_token.lower_ not in ("was", "were"):
        return None

    number = verb_token.morph.get("Number")
    if number:
        return number[0]

    # Check aux children (handles "does comprise", "is connected")
    for child in verb_token.children:
        if child.dep_ in ("aux", "auxpass"):
            if child.lower_ in _MODAL_VERBS:
                return None
            child_number = child.morph.get("Number")
            if child_number:
                return child_number[0]

    return None


def check_subject_verb_agreement(sentence: str) -> list[str]:
    """
    Detect subject-verb number disagreement in English sentences.

    Uses spaCy's dependency parse to find nsubj relations, then
    compares the Number morphology of the subject head noun with
    that of the governing verb.

    Designed to catch patent-translation errors like:
      - "The device comprise ..."       (Sing subj + Plur verb)
      - "The devices comprises ..."     (Plur subj + Sing verb)
      - "The apparatus for processing wafers comprise ..."
        (subject head is 'apparatus', not 'wafers')
    """
    if not isinstance(sentence, str):
        return []

    doc = nlp(sentence)
    issues: list[str] = []
    seen_pairs: set[tuple[int, int]] = set()  # avoid duplicate flags

    for token in doc:
        if token.dep_ != "nsubj":
            continue

        verb = token.head

        # Skip if verb is not actually a verb POS
        if verb.pos_ not in ("VERB", "AUX"):
            continue

        pair_key = (token.i, verb.i)
        if pair_key in seen_pairs:
            continue
        seen_pairs.add(pair_key)

        subj_num = _find_true_subject_number(token)
        verb_num = _find_verb_number(verb)

        # Can only flag when both sides have a determinable number
        if subj_num is None or verb_num is None:
            continue

        if subj_num != verb_num:
            issues.append(
                f"SV agreement: '{token.text}' ({subj_num}) ↔ "
                f"'{verb.text}' ({verb_num})"
            )

    return issues


# ========== 5. Bidirectional glossary checker ==========


def check_glossary_terms(
    src_text: str,
    tgt_text: str,
    glossary: list[tuple[str, list[str], list[str] | str | None]],
    src_lang: str,
    tgt_text_lemma: str = "",
    src_text_lemma: str = "",
) -> list[str]:
    """
    Generic glossary checker with optional lemma fallback.

    glossary entries:
      JP2EN: (src_term, [tgt_variants], [tgt_variants_lemma])
      EN2JP: (src_term, [tgt_variants], src_term_lemma)

    When src_lang == 'JP' (JP2EN):
      First checks exact normalized match; if no hit, falls back to
      lemmatized variant vs. lemmatized target text.

    When src_lang == 'EN' (EN2JP):
      First checks exact normalized match on source side; if no hit,
      falls back to lemmatized source term vs. lemmatized source text.
      Target (JP) side is always exact match.
    """
    flags: list[str] = []

    if not isinstance(src_text, str) or not isinstance(tgt_text, str):
        return flags

    if src_lang == "JP":
        src_text_norm = normalize_ja(src_text)
        tgt_text_norm = normalize_en(tgt_text)
        for entry in glossary:
            src_term, tgt_variants, tgt_variants_lemma = entry
            src_term_norm = normalize_ja(src_term)
            if src_term_norm and src_term_norm in src_text_norm:
                # Exact match first
                if any(normalize_en(v) in tgt_text_norm for v in tgt_variants):
                    continue
                # Lemma fallback
                if tgt_variants_lemma and tgt_text_lemma:
                    if any(vl in tgt_text_lemma for vl in tgt_variants_lemma):
                        continue
                flags.append(
                    f"Glossary missing: '{src_term}' → {', '.join(tgt_variants)}"
                )
    else:  # src_lang == "EN"
        src_text_norm = normalize_en(src_text)
        tgt_text_norm = normalize_ja(tgt_text)
        for entry in glossary:
            src_term, tgt_variants, src_term_lemma = entry
            src_term_norm = normalize_en(src_term)
            # Check if source term appears (exact or lemma)
            found_in_src = False
            if src_term_norm and src_term_norm in src_text_norm:
                found_in_src = True
            elif src_term_lemma and src_text_lemma:
                if src_term_lemma in src_text_lemma:
                    found_in_src = True

            if found_in_src:
                if not any(
                    normalize_ja(v) in tgt_text_norm for v in tgt_variants
                ):
                    flags.append(
                        f"Glossary missing: '{src_term}' → {', '.join(tgt_variants)}"
                    )

    return flags


# ========== 6. Master runner ==========

def run_all_checks(
    df: pd.DataFrame,
    glossary: list[tuple[str, list[str]]],
    glossary_direction: str,
) -> pd.DataFrame:
    """
    df: DataFrame with columns ['Japanese', 'English']
    glossary: list of (src_term, [tgt_variants])
    glossary_direction: 'JP2EN' or 'EN2JP'

    SV agreement results are stored in a separate 'SV_Issues' column
    (kept out of the main 'Issues' column due to low accuracy).

    Glossary matching uses lemma fallback for singular/plural tolerance.
    """
    results = []
    src_lang = "JP" if glossary_direction == "JP2EN" else "EN"

    # Precompute lemmatized glossary entries (one-time cost)
    glossary_ext: list[tuple[str, list[str], list[str] | str | None]] = []
    for src_term, tgt_variants in glossary:
        if glossary_direction == "JP2EN":
            # Target is EN: precompute lemmatized variants
            tgt_lemmas = [lemmatize_last_noun(v) for v in tgt_variants]
            glossary_ext.append((src_term, tgt_variants, tgt_lemmas))
        else:
            # Source is EN: precompute lemmatized source term
            src_lemma = lemmatize_last_noun(src_term)
            glossary_ext.append((src_term, tgt_variants, src_lemma))

    for _, row in df.iterrows():
        jp = row["Japanese"]
        en = row["English"]

        flags: list[str] = []
        sv_flags: list[str] = []

        # Per-row lemmatized English text (computed once, used by glossary)
        en_lemma = ""
        if isinstance(en, str) and glossary_direction == "JP2EN":
            en_lemma = lemmatize_nouns_in_text(en)

        en_src_lemma = ""
        if isinstance(en, str) and glossary_direction == "EN2JP":
            en_src_lemma = lemmatize_nouns_in_text(en)

        # English-side form checks:
        # Only run when English is the TARGET (JP2EN).
        if glossary_direction == "JP2EN":
            flags += check_word_repetition(en)
            flags += check_double_space(en)
            flags += check_whitespace_before_punctuation(en)
            flags += check_a_an(en)
            flags += check_for_bare_verb(en)
            flags += check_missing_final_period(jp, en)
            sv_flags += check_subject_verb_agreement(en)

        # Glossary checks (directional)
        if src_lang == "JP":
            flags += check_glossary_terms(
                jp, en, glossary_ext, src_lang="JP",
                tgt_text_lemma=en_lemma,
            )
        else:
            flags += check_glossary_terms(
                en, jp, glossary_ext, src_lang="EN",
                src_text_lemma=en_src_lemma,
            )

        results.append({
            "Japanese": jp,
            "English": en,
            "Issues": "; ".join(flags),
            "SV_Issues": "; ".join(sv_flags),
        })

    return pd.DataFrame(results)


# ========== 7. Excel export (direction in file name) ==========

def export_with_format(
    out_df: pd.DataFrame,
    src_path: str,
    direction_tag: str
) -> str:
    """
    Export results to a single Excel file with three sheets:
      - 'All'        — all rows with main Issues column
      - 'IssuesOnly' — rows that have at least one main issue
      - 'SV_Agreement' — rows that have SV agreement flags (separate due to low accuracy)

    Column order:
      - JP2EN: A=Japanese, B=English
      - EN2JP: A=English,  B=Japanese
    """

    # Decide column order based on direction
    if direction_tag == "EN2JP":
        main_cols = ["English", "Japanese", "Issues"]
        sv_cols = ["English", "Japanese", "SV_Issues"]
    else:  # JP2EN
        main_cols = ["Japanese", "English", "Issues"]
        sv_cols = ["Japanese", "English", "SV_Issues"]

    all_df = out_df[main_cols]

    issue_mask = all_df["Issues"].astype(str).str.strip() != ""
    issues_df = all_df[issue_mask]

    sv_mask = out_df["SV_Issues"].astype(str).str.strip() != ""
    sv_df = out_df.loc[sv_mask, sv_cols]

    out_dir = os.path.dirname(src_path)
    stamp = datetime.now().strftime("%Y%m%d")
    out_file = os.path.join(
        out_dir,
        f"proofreading_result_{direction_tag}_{stamp}.xlsx"
    )

    with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
        all_df.to_excel(writer, index=False, sheet_name="All")
        issues_df.to_excel(writer, index=False, sheet_name="IssuesOnly")
        sv_df.to_excel(writer, index=False, sheet_name="SV_Agreement")

        book = writer.book
        for sheet in book.worksheets:
            sheet.column_dimensions["A"].width = 60
            sheet.column_dimensions["B"].width = 60
            sheet.column_dimensions["C"].width = 60

            for row in sheet.iter_rows(
                min_row=1,
                max_row=sheet.max_row,
                min_col=1,
                max_col=sheet.max_column,
            ):
                for cell in row:
                    if cell.row == 1 and cell.column in (1, 2, 3):
                        cell.alignment = Alignment(
                            wrapText=True,
                            horizontal="center"
                        )
                    else:
                        cell.alignment = Alignment(wrapText=True)

    return out_file


# ========== 8. CLI entry (for TXT input) ==========

def main():
    txt = input("Enter path of bilingual text: ").strip('"')
    df = load_bilingual_txt(txt)

    term_list_file = input("Enter path of term list: ").strip('"')
    glossary, glossary_direction = load_glossary_xlsx(term_list_file)
    print(f"Detected glossary direction: {glossary_direction}")

    out_df = run_all_checks(df, glossary, glossary_direction)
    out_file = export_with_format(out_df, txt, glossary_direction)
    print(f"Saved: {out_file}")


if __name__ == "__main__":
    main()
