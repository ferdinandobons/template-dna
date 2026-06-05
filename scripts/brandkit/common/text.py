# SPDX-License-Identifier: MIT
"""Text utilities: rich-run helpers, markdown-literal detection, slugs, and the
multilingual name-token lexicon used by role inference.

This module is intentionally dependency-free (stdlib only) so every layer of the
engine can import it without pulling in lxml / python-docx.

Four concerns:

1. **Rich runs** - the inline text model shared by the IntermediateDocument and
   the resolver. A *run* is a ``{"t": str, "b"?: bool, "i"?: bool, "u"?: bool,
   "code"?: bool, "link"?: str}`` dict. Helpers normalize loose input (a bare
   string, a list of runs, the ``text:`` sugar) into a canonical run list and
   flatten runs back to plain text.
2. **Markdown-literal detection** - the L0 ``markdown_literal`` checker must
   catch literal ``**bold**``, ``*i*``, `` `code` ``, ``## heading``, and table
   ``|`` pipes that leaked into rendered text. :data:`MARKDOWN_LITERAL_RE` and
   :func:`find_markdown_literals` are the shared detector.
3. **Slug / safe filename** - deterministic, filesystem-safe identifiers for
   brand-kit directory names and asset filenames.
4. **Name-token lexicon (a WEAK PRIOR, never a gate)** - multilingual word lists
   the role-inference scorer consults as its *weakest* (0.20) signal and only as
   a LAST RESORT. The PRIMARY signals are structural, language-invariant OOXML
   facts (builtin style ids, field codes, SDT flags, placeholder TYPES,
   named-range geometry); the lexicon merely ADDS weak positive evidence when no
   structural signal is available and never gates output. See
   :data:`NAME_TOKEN_LEXICON` for the full anti-overfitting contract. Tokens
   cover EN/IT/FR/DE/ES.
"""
from __future__ import annotations

import re
import unicodedata
from typing import Iterable, Optional, Union

# A canonical inline run. Only "t" is required.
Run = dict
RunInput = Union[str, dict, Iterable[dict], None]


# ---------------------------------------------------------------------------
# Rich runs
# ---------------------------------------------------------------------------
# The boolean toggle keys a run may carry, plus the string "link" key.
RUN_TOGGLE_KEYS: tuple[str, ...] = ("b", "i", "u", "strike", "code", "sup", "sub")
RUN_STRING_KEYS: tuple[str, ...] = ("link",)


def normalize_runs(value: RunInput, *, text: Optional[str] = None) -> list[Run]:
    """Coerce loose inline input into a canonical list of run dicts.

    Accepts, in priority order:
      - ``value`` already a list/tuple of run dicts -> validated & copied.
      - ``value`` a single run dict -> wrapped in a one-element list.
      - ``value`` a plain ``str`` -> ``[{"t": value}]``.
      - ``value`` falsy and ``text`` given (the ``text:`` block sugar) ->
        ``[{"t": text}]``.
      - everything falsy -> ``[]``.

    Each produced run keeps only recognised keys (``t`` + the toggle/link keys),
    drops falsy toggles, and guarantees ``t`` is a ``str``.
    """
    if value is None or value == "":
        if text:
            return [{"t": str(text)}]
        return []
    if isinstance(value, str):
        return [{"t": value}] if value else []
    if isinstance(value, dict):
        runs_in: Iterable[dict] = [value]
    else:
        runs_in = list(value)  # type: ignore[arg-type]
    out: list[Run] = []
    for r in runs_in:
        if not isinstance(r, dict):
            # Tolerate a stray plain string inside a list.
            if isinstance(r, str) and r:
                out.append({"t": r})
            continue
        t = str(r.get("t", ""))
        run: Run = {"t": t}
        for k in RUN_TOGGLE_KEYS:
            if r.get(k):
                run[k] = True
        for k in RUN_STRING_KEYS:
            if r.get(k):
                run[k] = str(r[k])
        out.append(run)
    return out


def runs_to_text(runs: Iterable[Run]) -> str:
    """Flatten a run list to its concatenated plain text."""
    return "".join(str(r.get("t", "")) for r in runs)


def plain_run(text: str) -> list[Run]:
    """Return a single-run list for plain ``text`` (``[]`` if empty)."""
    return [{"t": text}] if text else []


# ---------------------------------------------------------------------------
# Markdown-literal detection
# ---------------------------------------------------------------------------
# Each alternative names a leaked markdown construct. The detector is
# deliberately conservative: it targets the markers most likely to survive into
# a *rendered* document (where they are always wrong), not every theoretical
# markdown token.
_MD_PATTERNS: tuple[tuple[str, str], ...] = (
    ("bold", r"\*\*[^\s*][^*]*?\*\*"),                 # **bold**
    ("bold_underscore", r"__[^\s_][^_]*?__"),          # __bold__
    ("italic", r"(?<![\*\w])\*[^\s*][^*]*?\*(?![\*\w])"),  # *italic*
    ("code", r"`[^`\n]+?`"),                            # `code`
    ("heading", r"^\s{0,3}#{1,6}\s+\S"),               # ## heading (line start)
    ("bullet", r"^\s{0,3}[-*+]\s+\S"),                 # - bullet (line start)
    ("link", r"\[[^\]]+\]\([^)]+\)"),                  # [text](url)
    ("table_pipe", r"^\s*\|.+\|\s*$"),                 # | a | b | table row
)

# MULTILINE so the line-anchored alternatives (heading/bullet/table_pipe) match
# at every embedded line start, not just the string start. The inline ``(?m)``
# flag is disallowed mid-pattern on Python 3.11+, so the flag is set here.
MARKDOWN_LITERAL_RE = re.compile(
    "|".join(f"(?P<{name}>{pat})" for name, pat in _MD_PATTERNS),
    re.MULTILINE,
)


def find_markdown_literals(text: str) -> list[dict]:
    """Return every markdown-literal match in ``text``.

    Each result is ``{"kind": <pattern-name>, "match": <substring>,
    "start": int, "end": int}``. An empty list means the text is clean. This is
    the shared detector behind the L0 ``markdown_literal`` checker; keeping it
    here makes the same rule format-agnostic across docx/pptx/xlsx.
    """
    out: list[dict] = []
    for m in MARKDOWN_LITERAL_RE.finditer(text or ""):
        out.append(
            {
                "kind": m.lastgroup,
                "match": m.group(),
                "start": m.start(),
                "end": m.end(),
            }
        )
    return out


def has_markdown_literal(text: str) -> bool:
    """Return True if ``text`` contains any markdown literal."""
    return MARKDOWN_LITERAL_RE.search(text or "") is not None


# ---------------------------------------------------------------------------
# Slug / safe filename
# ---------------------------------------------------------------------------
_SLUG_STRIP_RE = re.compile(r"[^a-z0-9]+")
_SLUG_TRIM_RE = re.compile(r"^-+|-+$")
_UNSAFE_FILENAME_RE = re.compile(r"[^A-Za-z0-9._-]+")


def slugify(value: str, *, max_len: int = 64, default: str = "untitled") -> str:
    """Return a lowercase, hyphenated, ASCII slug suitable for a directory name.

    Unicode is transliterated to its ASCII approximation (NFKD + drop combining
    marks); runs of non-alphanumerics collapse to a single hyphen; leading and
    trailing hyphens are trimmed. An empty result falls back to ``default``.
    """
    norm = unicodedata.normalize("NFKD", value or "")
    ascii_only = norm.encode("ascii", "ignore").decode("ascii").lower()
    slug = _SLUG_STRIP_RE.sub("-", ascii_only)
    slug = _SLUG_TRIM_RE.sub("", slug)
    if max_len and len(slug) > max_len:
        slug = _SLUG_TRIM_RE.sub("", slug[:max_len])
    return slug or default


def safe_filename(value: str, *, default: str = "file") -> str:
    """Return a filesystem-safe filename, preserving a single extension.

    Unlike :func:`slugify`, this keeps the case and the dot before a short
    extension. Path separators and unsafe characters collapse to ``_``; the
    result never starts with a dot (no accidental hidden files) and is never
    empty.
    """
    name = (value or "").strip().replace("/", "_").replace("\\", "_")
    name = _UNSAFE_FILENAME_RE.sub("_", name)
    name = name.lstrip(".")
    return name or default


# ---------------------------------------------------------------------------
# Multilingual name-token lexicon for role inference (WEAK PRIOR ONLY).
# ---------------------------------------------------------------------------
# IMPORTANT (plan §5 / M-i-8 lexicon demotion): this lexicon is the engine's
# *weakest* signal and a LAST-RESORT TIEBREAKER, retained ONLY for the
# comprehension-absent deterministic path. It is deliberately NOT deleted - it is
# the only thing left to nudge an ambiguous custom-named style when no structural
# evidence exists - but it has been firmly DEMOTED below every structural signal:
#
#   PRIMARY (language-invariant, structural OOXML facts; these decide the role):
#     - docx: builtin style ids (``Heading N`` / ``Title`` / ``Caption`` /
#       ``Quote`` / ``Normal``), field codes (``w:instrText`` ``TOC`` /
#       ``TOC \c`` / ``TOC \f``), SDT flags (alias / dataBinding /
#       docPartGallery / showingPlcHdr).
#     - pptx: placeholder *types* (TITLE / SUBTITLE / BODY ...) and named layouts.
#     - xlsx: named-range geometry (single vs multi cell, merged header, freeze,
#       tables) and number formats.
#   WEAK PRIOR (this lexicon): lowercased substring containment in a style's
#     *display name*. It only ever ADDS weak positive evidence; it NEVER strips,
#     NEVER overrides a structural signal, and NEVER matches on rendered body
#     text. A role recognised *only* via this lexicon is stamped best_effort /
#     low-confidence (<= the structural floor) so it can never gate output: the
#     deterministic ``resolver_targets_exist`` / ``comprehension_targets_exist``
#     guards reject any load-bearing ref that is not a verbatim structural id.
#
# Because the signal is language-invariant at the structural layer, removing or
# editing tokens here can only ever weaken a heuristic tiebreaker - it can never
# change which output is brand-valid. That is the whole point of the demotion.
#
# Maps a semantic *role family* -> set of lowercase substrings that, when found
# in a style's display name, weakly suggest that role. EN / IT / FR / DE / ES.
# The scorer matches by lowercased substring containment, so multi-word phrases
# ("elenco puntato") and single tokens ("puce") both work.
NAME_TOKEN_LEXICON: dict[str, frozenset[str]] = {
    "heading": frozenset({
        "heading", "title", "titolo", "titre", "uberschrift", "überschrift",
        "titulo", "título", "head",
    }),
    "callout": frozenset({
        "callout", "box", "riquadro", "encadre", "encadré", "kasten",
        "caja", "cuadro", "note", "nota", "highlight", "evidenza",
    }),
    "list.bullet": frozenset({
        "bullet", "elenco puntato", "elenchi puntati", "puce", "puces",
        "aufzahlung", "aufzählung", "vineta", "viñeta", "list", "elenco",
        "liste", "lista",
    }),
    "list.number": frozenset({
        "number", "numbered", "numerato", "numerique", "numérique",
        "nummeriert", "numerada", "ordered", "ordinato", "ordnung",
    }),
    "table": frozenset({
        "table", "tabella", "tableau", "tabelle", "tabla", "grid", "griglia",
    }),
    "caption": frozenset({
        "caption", "didascalia", "legende", "légende", "beschriftung",
        "leyenda", "figure", "figura",
    }),
    "quote": frozenset({
        "quote", "citazione", "citation", "zitat", "cita", "blockquote",
    }),
    "cover": frozenset({
        "cover", "copertina", "couverture", "deckblatt", "portada",
        "title page", "frontespizio",
    }),
    "toc": frozenset({
        "toc", "contents", "sommario", "indice", "table of contents",
        "inhaltsverzeichnis", "tabla de contenido",
    }),
    "kpi": frozenset({
        "kpi", "metric", "metrica", "metrik", "stat", "indicator",
        "indicatore", "indicador",
    }),
}

# Generic company-name noise tokens are NOT hardcoded here (the noise-token
# stripper is profile-gated and learned from the template per §5.1.2). This
# lexicon only ever *adds* weak positive evidence; it never strips.


def name_token_score(style_name: str, role_family: str) -> float:
    """Return 1.0 if any lexicon token for ``role_family`` is in ``style_name``.

    Comparison is case-insensitive substring containment. Returns 0.0 when the
    role family is unknown or no token matches. This is the raw (pre-weight)
    name-token signal; the scorer multiplies it by the 0.20 WEAK-PRIOR weight.

    This is a tiebreaker, not a decision: a non-zero score only ever ADDS to a
    role's evidence and is dominated by any structural signal (builtin style id,
    field code, placeholder type, named-range geometry). It never gates output -
    a role won purely on this score is best_effort/low-confidence and is still
    subject to the deterministic ``resolver_targets_exist`` guard. Returning 0.0
    here can never *remove* a structurally-established role.
    """
    tokens = NAME_TOKEN_LEXICON.get(role_family)
    if not tokens:
        return 0.0
    low = (style_name or "").lower()
    return 1.0 if any(tok in low for tok in tokens) else 0.0


def tokenize_name(style_name: str) -> list[str]:
    """Split a style display name into lowercased alphanumeric tokens.

    Used by the noise-token frequency analysis (e.g. detecting that a company
    name token recurs across many style names). Splits on any non-alphanumeric.
    """
    return [t for t in re.split(r"[^0-9A-Za-z]+", (style_name or "").lower()) if t]
