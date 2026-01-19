import re
import argparse
from dataclasses import dataclass
from collections import Counter, defaultdict

import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn


# ============================================================
# IMPORTANT NOTE
# ============================================================
# Bullet mode functionality is preserved as-is.
# The existing bullet-mode pipeline (text extraction, coordinate-based
# bullet indent inference, and postprocess_formatting) is kept unchanged.
#
# We add a new optional flag `no_bullets` to `convert(...)` that switches
# to a separate "handout/no bullets" pipeline based on coordinates for
# *all* lines, WITHOUT using bold for classification.
# ============================================================


# ----------------------------
# BULLET DETECTION (PDF TEXT)  (UNCHANGED)
# ----------------------------
BULLET_CHARS = [
    "➣", "➤", "➢", "➔",
    "•", "◦", "·", "∙", "‣", "⁃",
    "▪", "■", "◼", "◾", "◻", "□",
    "-", "–", "—",
]

BULLET_RE = re.compile(rf"^\s*(?:{'|'.join(re.escape(c) for c in BULLET_CHARS)})\s+(.*\S)\s*$")


def set_aptos_12(doc: Document) -> None:
    style = doc.styles["Normal"]
    style.font.name = "Aptos (Body)"
    style.font.size = Pt(12)

    rfonts = style.element.rPr.rFonts
    rfonts.set(qn("w:ascii"), "Aptos (Body)")
    rfonts.set(qn("w:hAnsi"), "Aptos (Body)")
    rfonts.set(qn("w:cs"), "Aptos (Body)")
    rfonts.set(qn("w:eastAsia"), "Aptos (Body)")


def is_footer_noise(line: str) -> bool:
    # (UNCHANGED)
    l = line.strip()
    if not l:
        return True
    if re.fullmatch(r"\d+", l):
        return True
    if "@" in l:
        return True
    if l.lower().startswith("further reading"):
        return True
    return False


def looks_like_heading(line: str) -> bool:
    # (UNCHANGED)
    l = line.strip()
    if not l:
        return False
    # Don't treat bullets as headings
    if BULLET_RE.match(l):
        return False

    if len(l) > 80:
        return False

    if l.endswith(":"):
        return True

    words = re.findall(r"[A-Za-z']+", l)
    if not words:
        return False

    if l.isupper() and len(l) <= 60:
        return True

    title_like = sum(1 for w in words if w[0].isupper()) / max(1, len(words))
    if title_like >= 0.7 and len(l) <= 70:
        return True

    return False


def add_bold_line(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.name = "Aptos (Body)"
    run.font.size = Pt(12)


def add_bullet(doc: Document, text: str, level: int = 0) -> None:
    """
    level 0 -> List Bullet
    level 1 -> List Bullet 2
    level 2 -> List Bullet 3
    """
    level = max(0, min(level, 2))
    style = "List Bullet" if level == 0 else f"List Bullet {level + 1}"
    p = doc.add_paragraph(text, style=style)
    for r in p.runs:
        r.font.name = "Aptos (Body)"
        r.font.size = Pt(12)


def normalize_lines(text: str):
    # (UNCHANGED)
    lines = []
    for ln in text.splitlines():
        ln = re.sub(r"\s+", " ", ln.strip())
        if ln and not is_footer_noise(ln):
            lines.append(ln)
    return lines


def bullet_text(line: str):
    """
    Return the bullet content if line starts with a recognised bullet marker, else None.
    """
    m = BULLET_RE.match(line)
    if not m:
        return None
    return m.group(1).strip()


# ----------------------------
# COORDINATE-BASED BULLET LEVELS  (UNCHANGED)
# ----------------------------
def _is_bullet_start_text(t: str) -> bool:
    if not t:
        return False
    t = t.lstrip()
    return bool(t) and t[0] in BULLET_CHARS


def _extract_bullet_x_positions(page: fitz.Page) -> list[float]:
    """
    Returns a list of x-positions (bullet glyph x) for bullet lines on the page,
    in reading order (top-to-bottom, left-to-right).

    This uses page.get_text("dict") (coordinates) but does NOT change the text pipeline.
    """
    d = page.get_text("dict")
    hits = []

    # Walk blocks/lines/spans; sort lines by y then x for stable reading order.
    for block in d.get("blocks", []):
        if block.get("type") != 0:  # 0 = text block
            continue
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue

            # Build the line text in a conservative way
            line_text = "".join(s.get("text", "") for s in spans)
            if not _is_bullet_start_text(line_text):
                continue

            # Find the span where the first non-space char lives; use its x0
            bullet_x = None
            for s in spans:
                st = s.get("text", "")
                if not st:
                    continue
                # Skip spans that are only whitespace
                if not st.strip():
                    continue

                # The bullet glyph is the first non-space char of the full line
                # If this span begins (after lstrip) with a bullet, take its x0.
                if st.lstrip() and st.lstrip()[0] in BULLET_CHARS:
                    bullet_x = float(s["bbox"][0])
                    break

                # Otherwise, sometimes the bullet is glued after spaces in the same span
                # We still treat the first non-space char as the bullet. If it is bullet, use this span x0.
                first = st.lstrip()[0] if st.lstrip() else ""
                if first in BULLET_CHARS:
                    bullet_x = float(s["bbox"][0])
                    break

            if bullet_x is None:
                # Fallback: use line bbox x0 if span parsing fails
                bbox = line.get("bbox")
                if bbox:
                    bullet_x = float(bbox[0])
                else:
                    continue

            bbox = line.get("bbox") or spans[0].get("bbox")
            y0 = float(bbox[1]) if bbox else 0.0
            x0 = float(bbox[0]) if bbox else bullet_x
            hits.append((y0, x0, bullet_x))

    hits.sort(key=lambda t: (t[0], t[1]))
    return [bx for _, __, bx in hits]


def _cluster_x_positions(xs: list[float], tol: float = 4.0) -> list[float]:
    """
    Cluster x positions into columns using a simple tolerance (points).
    Returns cluster centers sorted ascending.
    Deterministic and robust for slide bullets.
    """
    if not xs:
        return []
    xs_sorted = sorted(xs)
    clusters = [[xs_sorted[0]]]
    for x in xs_sorted[1:]:
        if abs(x - clusters[-1][-1]) <= tol:
            clusters[-1].append(x)
        else:
            clusters.append([x])

    centers = [sum(c) / len(c) for c in clusters]
    centers.sort()
    return centers


def _levels_for_bullets_on_page(bullet_xs: list[float]) -> list[int]:
    """
    Convert bullet x positions to levels 0/1/2 by clustering.
    Leftmost cluster => level 0, next => level 1, etc.
    """
    if not bullet_xs:
        return []
    centers = _cluster_x_positions(bullet_xs, tol=4.0)
    # Map each bullet x to the nearest cluster center index
    levels = []
    for x in bullet_xs:
        idx = min(range(len(centers)), key=lambda i: abs(x - centers[i]))
        levels.append(max(0, min(idx, 2)))
    return levels


# ----------------------------
# POSTPROCESSING ONLY (UNCHANGED)
# ----------------------------
def postprocess_formatting(docx_path: str) -> None:
    """
    Postprocess the produced DOCX WITHOUT changing content:
    1) Remove bold-only heading paragraphs that have no bullet paragraphs beneath them (before next heading).
    2) Convert remaining headings into level-0 bullets ("List Bullet"), keeping the heading text bold.
    3) Force ALL bullets to use "List Bullet" (filled dot) and simulate nesting via indentation only.
       Bullets under a heading are shifted +1 level deeper (cap at 2).
    """
    doc = Document(docx_path)
    paras = doc.paragraphs

    def is_heading(p) -> bool:
        txt = (p.text or "").strip()
        if not txt:
            return False
        # don't treat list paragraphs as headings
        if p.style and p.style.name and p.style.name.startswith("List"):
            return False
        # Heading = all non-empty runs are bold
        any_nonempty_run = False
        for run in p.runs:
            if run.text and run.text.strip():
                any_nonempty_run = True
                if not run.bold:
                    return False
        return any_nonempty_run

    def is_bullet(p) -> bool:
        return bool(p.style and p.style.name and p.style.name.startswith("List"))

    def bullet_level_from_style_name(name: str) -> int:
        if name == "List Bullet":
            return 0
        if name == "List Bullet 2":
            return 1
        if name == "List Bullet 3":
            return 2
        return 0

    HANG = Pt(18)
    STEP = Pt(18)
    BASE_LEFT = Pt(18)

    def apply_level_indent(p, level: int) -> None:
        level = max(0, min(level, 2))
        pf = p.paragraph_format
        pf.left_indent = Pt(BASE_LEFT.pt + STEP.pt * level)
        pf.first_line_indent = Pt(-HANG.pt)

    def enforce_aptos_12(p) -> None:
        for r in p.runs:
            r.font.name = "Aptos (Body)"
            r.font.size = Pt(12)

    # (1) Remove headings with no bullets beneath them
    to_delete_idxs = []
    for i, p in enumerate(paras):
        if not is_heading(p):
            continue

        has_bullet = False
        for j in range(i + 1, len(paras)):
            nxt = paras[j]
            if is_heading(nxt):
                break
            if is_bullet(nxt):
                has_bullet = True
                break

        if not has_bullet:
            to_delete_idxs.append(i)

    for i in reversed(to_delete_idxs):
        p = paras[i]
        p._element.getparent().remove(p._element)

    paras = doc.paragraphs

    # (2)(3) Headings => List Bullet; bullets => List Bullet always, indent for nesting
    in_heading_block = False

    for p in paras:
        if is_heading(p):
            p.style = doc.styles["List Bullet"]
            apply_level_indent(p, 0)

            for r in p.runs:
                if r.text and r.text.strip():
                    r.bold = True
            enforce_aptos_12(p)

            in_heading_block = True
            continue

        if in_heading_block and is_bullet(p):
            old_name = p.style.name if p.style else ""
            old_level = bullet_level_from_style_name(old_name)
            new_level = min(old_level + 1, 2)

            p.style = doc.styles["List Bullet"]
            apply_level_indent(p, new_level)
            enforce_aptos_12(p)
            continue

        if (p.text or "").strip() == "":
            in_heading_block = False
        else:
            if not is_bullet(p):
                in_heading_block = False

        if is_bullet(p):
            old_name = p.style.name if p.style else ""
            lvl = bullet_level_from_style_name(old_name)

            p.style = doc.styles["List Bullet"]
            apply_level_indent(p, lvl)
            enforce_aptos_12(p)
        else:
            enforce_aptos_12(p)

    doc.save(docx_path)


# ============================================================
# HANDOUT / NO-BULLETS MODE (NEW)
# ============================================================

@dataclass(frozen=True)
class LineObj:
    page_index: int
    text: str
    x0: float
    y0: float
    y1: float
    font_size: float
    span_count: int


def _norm_text_for_repeat(s: str) -> str:
    s = re.sub(r"\s+", " ", (s or "").strip())
    # Normalize common page counters like "3 / 18" or "3/18"
    s = re.sub(r"\b\d+\s*/\s*\d+\b", "{PAGECOUNT}", s)
    # Normalize standalone numbers
    s = re.sub(r"\b\d+\b", "{N}", s)
    return s.lower()


def _extract_all_lines_with_coords(pdf: fitz.Document) -> list[LineObj]:
    out: list[LineObj] = []
    for pi in range(pdf.page_count):
        page = pdf.load_page(pi)
        d = page.get_text("dict")
        for block in d.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                spans = line.get("spans", []) or []
                if not spans:
                    continue
                text = "".join(s.get("text", "") for s in spans)
                text = re.sub(r"\s+", " ", (text or "").strip())
                if not text:
                    continue

                # Coordinates
                bbox = line.get("bbox") or spans[0].get("bbox")
                if not bbox:
                    continue
                x0 = float(bbox[0])
                y0 = float(bbox[1])
                y1 = float(bbox[3])

                # Font size: robust median-ish via average
                sizes = [float(s.get("size", 0.0)) for s in spans if s.get("size") is not None]
                fs = (sum(sizes) / len(sizes)) if sizes else 0.0

                out.append(
                    LineObj(
                        page_index=pi,
                        text=text,
                        x0=x0,
                        y0=y0,
                        y1=y1,
                        font_size=fs,
                        span_count=len(spans),
                    )
                )

    # Sort global reading order: page, then y, then x
    out.sort(key=lambda l: (l.page_index, l.y0, l.x0))
    return out


def _detect_repeating_header_footer_lines(
    pdf: fitz.Document,
    all_lines: list[LineObj],
    top_frac: float = 0.10,
    bot_frac: float = 0.12,
    repeat_threshold: float = 0.55,
) -> set[tuple[int, str]]:
    """
    Identify (page_index, normalized_text) pairs to drop, based on repetition across pages
    and being located in the top/bottom bands of the page.

    We intentionally do NOT use bold. This is purely coordinate + repetition.
    """
    # Bucket lines by page
    lines_by_page: dict[int, list[LineObj]] = defaultdict(list)
    for l in all_lines:
        lines_by_page[l.page_index].append(l)

    total_pages = max(1, pdf.page_count)
    norm_counts = Counter()

    candidates: list[tuple[int, str]] = []

    for pi in range(pdf.page_count):
        page = pdf.load_page(pi)
        height = float(page.rect.height) if page.rect else 1.0
        top_y = height * top_frac
        bot_y = height * (1.0 - bot_frac)

        for l in lines_by_page.get(pi, []):
            band = "mid"
            if l.y0 <= top_y:
                band = "top"
            elif l.y1 >= bot_y:
                band = "bot"
            else:
                continue

            nt = _norm_text_for_repeat(l.text)
            if not nt or len(nt) <= 2:
                continue

            norm_counts[nt] += 1
            candidates.append((pi, nt))

    # Which normalized strings repeat on enough pages?
    min_pages = max(2, int(total_pages * repeat_threshold))
    repeating_norms = {nt for nt, c in norm_counts.items() if c >= min_pages}

    to_drop = {(pi, nt) for (pi, nt) in candidates if nt in repeating_norms}
    return to_drop


def _cluster_centers(xs: list[float], tol: float = 6.0) -> list[float]:
    # Similar to bullet clustering but with a slightly wider tolerance for handout layout.
    if not xs:
        return []
    xs_sorted = sorted(xs)
    clusters = [[xs_sorted[0]]]
    for x in xs_sorted[1:]:
        if abs(x - clusters[-1][-1]) <= tol:
            clusters[-1].append(x)
        else:
            clusters.append([x])
    centers = [sum(c) / len(c) for c in clusters]
    centers.sort()
    return centers


def _assign_indent_level(x0: float, centers: list[float], max_level: int = 2) -> int:
    if not centers:
        return 0
    idx = min(range(len(centers)), key=lambda i: abs(x0 - centers[i]))
    return max(0, min(idx, max_level))


def _is_tableish(line: LineObj) -> bool:
    """
    Heuristic: table-ish lines often have multiple spans and short token clusters.
    We avoid overfitting; this only gates "don't force bullets".
    """
    if line.span_count >= 4:
        return True
    # Many column-like separators / repeated spacing patterns in text
    if re.search(r"\s{2,}", line.text):
        return True
    # Truth-table-ish: lots of single-char tokens
    tokens = re.findall(r"\S+", line.text)
    if len(tokens) >= 6 and sum(1 for t in tokens if len(t) == 1) / len(tokens) >= 0.6:
        return True
    return False


def _handout_detect_headings(
    page_lines: list[LineObj],
    centers: list[float],
) -> set[int]:
    """
    Return indices (into page_lines) that are headings.
    We do NOT use bold. We use font size + spacing + position.
    """
    if not page_lines:
        return set()

    # Body font size: median-ish via simple robust percentile
    sizes = sorted([l.font_size for l in page_lines if l.font_size > 0.0])
    if not sizes:
        return set()
    body = sizes[len(sizes) // 2]

    # A heading is typically noticeably larger than body.
    # Use a small additive margin to handle small fonts.
    size_thresh = max(body + 1.0, body * 1.12)

    heading_idxs: set[int] = set()

    # Compute y-gaps
    for i, l in enumerate(page_lines):
        # Font size signal
        if l.font_size < size_thresh:
            continue

        # Position signal: closer to left margin band
        lvl = _assign_indent_level(l.x0, centers, max_level=2)
        if lvl > 0:
            # headings usually start near left edge; allow but require stronger spacing
            pass

        # Spacing signal: big gap above or below
        gap_above = None
        gap_below = None
        if i > 0:
            gap_above = l.y0 - page_lines[i - 1].y1
        if i + 1 < len(page_lines):
            gap_below = page_lines[i + 1].y0 - l.y1

        # Typical line height ~ (y1-y0)
        lh = max(1.0, (l.y1 - l.y0))
        big_gap = ((gap_above is not None and gap_above >= 0.8 * lh) or
                   (gap_below is not None and gap_below >= 0.8 * lh))

        # Text-shape helper (not required, but stabilizes):
        shortish = len(l.text) <= 90
        colonish = l.text.endswith(":")
        upperish = l.text.isupper() and len(l.text) <= 80

        if big_gap or colonish or upperish or shortish:
            heading_idxs.add(i)

    # De-duplicate consecutive headings (keep the first if they are effectively the same region)
    cleaned: set[int] = set()
    last_y1 = None
    for i in sorted(heading_idxs):
        l = page_lines[i]
        if last_y1 is not None and l.y0 - last_y1 < 3.0:
            # Very close; likely part of the same heading block; keep earliest only
            continue
        cleaned.add(i)
        last_y1 = l.y1

    return cleaned


def _handout_convert(pdf: fitz.Document, out_docx_path: str) -> None:
    doc = Document()
    set_aptos_12(doc)

    all_lines = _extract_all_lines_with_coords(pdf)

    # Identify repeating header/footer lines to drop
    to_drop = _detect_repeating_header_footer_lines(pdf, all_lines)

    # Group lines by page (post drop)
    by_page: dict[int, list[LineObj]] = defaultdict(list)
    for l in all_lines:
        nt = _norm_text_for_repeat(l.text)
        if (l.page_index, nt) in to_drop:
            continue
        # Basic whitespace cleanup (do not collapse internal symbols)
        txt = re.sub(r"\s+", " ", (l.text or "").strip())
        if not txt:
            continue
        by_page[l.page_index].append(
            LineObj(
                page_index=l.page_index,
                text=txt,
                x0=l.x0,
                y0=l.y0,
                y1=l.y1,
                font_size=l.font_size,
                span_count=l.span_count,
            )
        )

    for pi in range(pdf.page_count):
        page_lines = by_page.get(pi, [])
        if not page_lines:
            continue

        # Build indent bands (x0 clusters) using *all* non-empty lines
        xs = [l.x0 for l in page_lines]
        centers = _cluster_centers(xs, tol=6.0)

        # Detect headings without bold
        heading_idxs = _handout_detect_headings(page_lines, centers)

        # If the first line looks like a title and is a heading, keep it (like bullet mode keeps titles)
        # We will emit headings as bold lines in the DOCX (same heading representation as bullet mode output).
        current_heading_active = False

        # Group lines into blocks using y-gaps + indent level
        i = 0
        while i < len(page_lines):
            l = page_lines[i]

            if i in heading_idxs:
                # Emit heading as-is (like bullet mode "title/heading line")
                add_bold_line(doc, l.text)
                current_heading_active = True
                i += 1
                continue

            # Table-ish region: emit as plain paragraphs (do not force bullets)
            if _is_tableish(l):
                # Group consecutive table-ish lines with small gaps
                j = i + 1
                while j < len(page_lines):
                    nxt = page_lines[j]
                    if j in heading_idxs:
                        break
                    if not _is_tableish(nxt):
                        break
                    gap = nxt.y0 - page_lines[j - 1].y1
                    if gap > 10.0:
                        break
                    j += 1

                for k in range(i, j):
                    p = doc.add_paragraph(page_lines[k].text)
                    for r in p.runs:
                        r.font.name = "Aptos (Body)"
                        r.font.size = Pt(12)
                i = j
                continue

            # Otherwise: treat as a "list/paragraph-like" block.
            # To keep output consistent with your notes style, we render these as bullets.
            # Nesting comes from indent band levels.
            level = _assign_indent_level(l.x0, centers, max_level=2)

            # Build a block by collecting continuation lines that are visually part of this item:
            # same indent band (or very close), and small vertical gaps.
            item_text = l.text
            j = i + 1
            while j < len(page_lines):
                nxt = page_lines[j]
                if j in heading_idxs:
                    break
                if _is_tableish(nxt):
                    break

                nxt_level = _assign_indent_level(nxt.x0, centers, max_level=2)

                gap = nxt.y0 - page_lines[j - 1].y1

                # Continuation rule: same level and small gap -> append
                if nxt_level == level and gap <= 6.0:
                    item_text += " " + nxt.text
                    j += 1
                    continue

                # If next line is deeper indent and very close, treat it as a new bullet,
                # not continuation.
                break

            # Emit as bullet (even at level 0), because in no-bullets mode you still want
            # "notes style" output. Headings remain included above, as in bullet mode.
            add_bullet(doc, item_text.strip(), level)

            i = j

        # Page separator like bullet mode
        doc.add_paragraph("")

    doc.save(out_docx_path)
    print(f"Saved (handout mode): {out_docx_path}")


# ============================================================
# BULLET MODE CONVERTER (EXISTING) — UNCHANGED LOGIC
# ============================================================

def _bullet_convert(pdf: fitz.Document, out_docx_path: str) -> None:
    doc = Document()
    set_aptos_12(doc)

    for page_index in range(pdf.page_count):
        page = pdf.load_page(page_index)

        # IMPORTANT: keep your existing text extraction so output content stays identical
        raw = page.get_text("text") or ""
        lines = normalize_lines(raw)
        if not lines:
            continue

        # Compute bullet levels using coordinates (does not affect 'lines')
        bullet_xs = _extract_bullet_x_positions(page)
        bullet_levels = _levels_for_bullets_on_page(bullet_xs)
        bullet_level_idx = 0

        # Slide title: first non-bullet line (or short fallback)
        title = None
        for ln in lines:
            if bullet_text(ln) is not None:
                continue
            if looks_like_heading(ln) or len(ln) <= 60:
                title = ln
                break

        # Skip slides titled Outline or Summary (case-insensitive)
        if title and title.strip().lower() in {"outline", "summary"}:
            continue

        if title:
            add_bold_line(doc, title)

        current_bullet = None
        current_level = 0

        def flush_bullet():
            nonlocal current_bullet, current_level
            if current_bullet:
                add_bullet(doc, current_bullet.strip(), current_level)
                current_bullet = None
                current_level = 0

        for ln in lines:
            if title and ln == title:
                continue

            bt = bullet_text(ln)
            is_head = looks_like_heading(ln)

            if bt is not None:
                # new bullet starts
                flush_bullet()

                # Pull the next coordinate-derived level if available; else default to 0
                if bullet_level_idx < len(bullet_levels):
                    current_level = bullet_levels[bullet_level_idx]
                    bullet_level_idx += 1
                else:
                    current_level = 0

                current_bullet = bt
                continue

            if is_head:
                flush_bullet()
                add_bold_line(doc, ln)
                continue

            # continuation line: append to existing bullet
            if current_bullet:
                current_bullet += " " + ln

        flush_bullet()
        doc.add_paragraph("")

    # Save then postprocess formatting (ONLY formatting changes)
    doc.save(out_docx_path)
    postprocess_formatting(out_docx_path)
    print(f"Saved (formatted): {out_docx_path}")


# ============================================================
# PUBLIC API
# ============================================================

def convert(pdf_path: str, out_docx_path: str, no_bullets: bool = False) -> None:
    """
    Convert a PDF into a DOCX notes format.

    - no_bullets=False (default): ORIGINAL bullet-mode behavior (unchanged).
    - no_bullets=True: Handout/no-bullets mode using coordinates on all lines
      (no bold classification; headings kept and emitted like bullet mode).
    """
    pdf = fitz.open(pdf_path)
    if no_bullets:
        _handout_convert(pdf, out_docx_path)
    else:
        _bullet_convert(pdf, out_docx_path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf", help="Input PDF")
    ap.add_argument("out", help="Output docx")
    ap.add_argument(
        "--no-bullets",
        action="store_true",
        help="Handout/no-bullets mode: infer structure from coordinates on all lines",
    )
    args = ap.parse_args()
    convert(args.pdf, args.out, no_bullets=args.no_bullets)


if __name__ == "__main__":
    main()
