import re
import argparse

import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn


# Common bullet markers that appear in extracted PDF text
BULLET_CHARS = [
    "➣", "➤", "➢", "➔",
    "•", "◦", "·", "∙", "‣", "⁃",
    "▪", "■", "◼", "◾", "◻", "□",
    "-", "–", "—",
]

# Regex: optional leading spaces, then one of the bullet chars, then the bullet text
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


def _normalize_line_text_for_match(s: str) -> str:
    # Used to match 'dict' lines to 'text' lines, without changing output content.
    s = re.sub(r"\s+", " ", (s or "").strip())
    return s


def normalize_lines(text: str):
    lines = []
    for ln in text.splitlines():
        ln2 = _normalize_line_text_for_match(ln)
        if ln2 and not is_footer_noise(ln2):
            lines.append(ln2)
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
# COORDINATE-BASED LEVELS
# ----------------------------
def _cluster_x_positions(xs: list[float], tol: float = 4.0) -> list[float]:
    """
    Cluster x positions into columns using a simple tolerance (points).
    Returns cluster centers sorted ascending.
    Deterministic and robust for slide bullets/indents.
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


def _levels_from_xs(xs: list[float], tol: float = 4.0) -> list[int]:
    """
    Map each x in xs to a level 0/1/2 using clustering.
    Leftmost cluster => level 0, next => level 1, etc.
    """
    if not xs:
        return []
    centers = _cluster_x_positions(xs, tol=tol)
    levels = []
    for x in xs:
        idx = min(range(len(centers)), key=lambda i: abs(x - centers[i]))
        levels.append(max(0, min(idx, 2)))
    return levels


def _extract_bullet_x_positions(page: fitz.Page) -> list[float]:
    """
    Original behaviour support: x-positions for lines that start with a bullet glyph.
    """
    d = page.get_text("dict")
    hits = []

    for block in d.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue

            line_text = "".join(s.get("text", "") for s in spans)
            lt_norm = _normalize_line_text_for_match(line_text)
            if not lt_norm or is_footer_noise(lt_norm):
                continue

            # detect bullet at first non-space char of line_text
            t = line_text.lstrip()
            if not t or t[0] not in BULLET_CHARS:
                continue

            bbox = line.get("bbox") or spans[0].get("bbox")
            if not bbox:
                continue
            bullet_x = float(bbox[0])
            hits.append((float(bbox[1]), float(bbox[0]), bullet_x))

    hits.sort(key=lambda t: (t[0], t[1]))
    return [bx for _, __, bx in hits]


def _extract_all_line_xs_in_reading_order(page: fitz.Page) -> list[tuple[str, float]]:
    """
    NEW (used only in force_all_lines_bullets mode):
    Returns a reading-order list of (normalized_line_text, x0) for ALL text lines.

    We still filter footer noise here, so footer noise never gets assigned a level.
    """
    d = page.get_text("dict")
    items: list[tuple[float, float, str, float]] = []

    for block in d.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue

            raw_text = "".join(s.get("text", "") for s in spans)
            txt = _normalize_line_text_for_match(raw_text)
            if not txt or is_footer_noise(txt):
                continue

            bbox = line.get("bbox") or spans[0].get("bbox")
            if not bbox:
                continue

            y0 = float(bbox[1])
            x0 = float(bbox[0])
            items.append((y0, x0, txt, x0))

    # reading order
    items.sort(key=lambda t: (t[0], t[1]))
    return [(txt, x0) for _, __, txt, x0 in items]


def _align_levels_to_text_lines(text_lines: list[str], dict_lines_with_x: list[tuple[str, float]]) -> list[int]:
    """
    Given the output 'text_lines' (from page.get_text("text") -> normalize_lines),
    align coordinate-derived x positions from dict lines in reading order.

    This avoids changing the original text extraction pipeline.
    """
    dict_texts = [t for t, _ in dict_lines_with_x]
    dict_xs = [x for _, x in dict_lines_with_x]
    dict_levels = _levels_from_xs(dict_xs, tol=4.0) if dict_xs else []

    # Build a mapping from normalized text to a queue of indices in dict order
    positions: dict[str, list[int]] = {}
    for i, t in enumerate(dict_texts):
        positions.setdefault(t, []).append(i)

    aligned: list[int] = []
    last_level = 0

    for ln in text_lines:
        idx_list = positions.get(ln)
        if idx_list:
            idx = idx_list.pop(0)
            lvl = dict_levels[idx] if idx < len(dict_levels) else last_level
            last_level = lvl
            aligned.append(lvl)
        else:
            # If we can't match, keep the last known level as a stable fallback.
            aligned.append(last_level)

    return aligned


# ----------------------------
# POSTPROCESSING ONLY (unchanged)
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
        if p.style and p.style.name and p.style.name.startswith("List"):
            return False
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

    # (2)(3) enforce filled-dot bullet style + indent nesting
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


def convert(pdf_path: str, out_docx_path: str, force_all_lines_bullets: bool = False) -> None:
    """
    IMPORTANT:
      - When force_all_lines_bullets=False (default), behaviour is identical to your original script.
      - When force_all_lines_bullets=True, we still:
          * detect headings the same way as before
          * ignore footer noise the same way as before
          * use coordinate-based clustering for indentation levels
        but we treat every non-heading line as a bullet (even if it has no bullet glyph).
    """
    pdf = fitz.open(pdf_path)
    doc = Document()
    set_aptos_12(doc)

    for page_index in range(pdf.page_count):
        page = pdf.load_page(page_index)

        # Keep existing text extraction pipeline
        raw = page.get_text("text") or ""
        lines = normalize_lines(raw)
        if not lines:
            continue

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

        if not force_all_lines_bullets:
            # -----------------------------
            # ORIGINAL MODE (UNCHANGED)
            # -----------------------------
            bullet_xs = _extract_bullet_x_positions(page)
            bullet_levels = _levels_from_xs(bullet_xs, tol=4.0)
            bullet_level_idx = 0

            for ln in lines:
                if title and ln == title:
                    continue

                bt = bullet_text(ln)
                is_head = looks_like_heading(ln)

                if bt is not None:
                    flush_bullet()

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
            continue

        # -----------------------------
        # NEW MODE (OPT-IN)
        # Treat every non-heading line as bullet,
        # but still use coordinate clustering for indentation.
        # -----------------------------
        dict_lines_with_x = _extract_all_line_xs_in_reading_order(page)
        aligned_levels = _align_levels_to_text_lines(lines, dict_lines_with_x)

        for idx, ln in enumerate(lines):
            if title and ln == title:
                continue

            is_head = looks_like_heading(ln)
            if is_head:
                flush_bullet()
                add_bold_line(doc, ln)
                continue

            # Every non-heading line becomes its own bullet.
            flush_bullet()

            # If the line starts with a bullet glyph, strip it (so content matches normal bullet_text behaviour).
            bt = bullet_text(ln)
            bullet_content = bt if bt is not None else ln

            lvl = aligned_levels[idx] if idx < len(aligned_levels) else 0
            current_level = max(0, min(int(lvl), 2))
            current_bullet = bullet_content

        flush_bullet()
        doc.add_paragraph("")

    doc.save(out_docx_path)
    postprocess_formatting(out_docx_path)
    print(f"Saved (formatted): {out_docx_path}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf", help="Input slides PDF")
    ap.add_argument("out", help="Output docx")
    ap.add_argument(
        "--all-bullets",
        action="store_true",
        help="Treat every non-heading line as a bullet, still using coordinate-based indentation.",
    )
    args = ap.parse_args()
    convert(args.pdf, args.out, force_all_lines_bullets=args.all_bullets)


if __name__ == "__main__":
    main()
