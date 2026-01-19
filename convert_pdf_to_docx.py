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


def normalize_lines(text: str):
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
# COORDINATE-BASED LEVELS
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

    for block in d.get("blocks", []):
        if block.get("type") != 0:  # 0 = text block
            continue
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue

            line_text = "".join(s.get("text", "") for s in spans)
            if not _is_bullet_start_text(line_text):
                continue

            bullet_x = None
            for s in spans:
                st = s.get("text", "")
                if not st:
                    continue
                if not st.strip():
                    continue

                if st.lstrip() and st.lstrip()[0] in BULLET_CHARS:
                    bullet_x = float(s["bbox"][0])
                    break

                first = st.lstrip()[0] if st.lstrip() else ""
                if first in BULLET_CHARS:
                    bullet_x = float(s["bbox"][0])
                    break

            if bullet_x is None:
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


def _extract_line_x_positions(page: fitz.Page) -> list[float]:
    """
    Returns a list of x-positions (line left edge) for ALL text lines on the page,
    in reading order (top-to-bottom, left-to-right).
    """
    d = page.get_text("dict")
    hits = []

    for block in d.get("blocks", []):
        if block.get("type") != 0:  # text block
            continue
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue

            line_text = "".join(s.get("text", "") for s in spans)
            if not line_text or not line_text.strip():
                continue

            xs = [float(s["bbox"][0]) for s in spans if s.get("text", "").strip()]
            if not xs:
                continue
            line_x = min(xs)

            bbox = line.get("bbox") or spans[0].get("bbox")
            y0 = float(bbox[1]) if bbox else 0.0
            x0 = float(bbox[0]) if bbox else line_x

            hits.append((y0, x0, line_x))

    hits.sort(key=lambda t: (t[0], t[1]))
    return [lx for _, __, lx in hits]


def _extract_lines_xy_text_excluding_footer(page: fitz.Page) -> list[tuple[float, float, str]]:
    """
    ONLY USED FOR all_lines MODE:
    Return ALL text lines on the page as (y0, x0, text), sorted by (y0, x0),
    excluding content in the footer area at the bottom using y-coordinate.
    """
    FOOTER_CUTOFF_RATIO = 0.90  # bottom 10% treated as footer
    y_cutoff = float(page.rect.height) * FOOTER_CUTOFF_RATIO

    d = page.get_text("dict")
    hits: list[tuple[float, float, str]] = []

    for block in d.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue

            bbox = line.get("bbox") or spans[0].get("bbox")
            if not bbox:
                continue

            y0 = float(bbox[1])
            x0 = float(bbox[0])

            # Ignore bottom footer region in all-lines mode
            if y0 >= y_cutoff:
                continue

            text = "".join(s.get("text", "") for s in spans)
            text = re.sub(r"\s+", " ", text.strip())
            if not text or is_footer_noise(text):
                continue

            hits.append((y0, x0, text))

    hits.sort(key=lambda t: (t[0], t[1]))
    return hits


def _extract_first_line_text(page: fitz.Page) -> str | None:
    """
    ONLY USED FOR all_lines MODE:
    Determine the slide title as the first text line on the page by (y, x) order,
    using coordinates (not looks_like_heading), and return its normalized text.
    """
    hits = _extract_lines_xy_text_excluding_footer(page)
    if not hits:
        return None
    return hits[0][2]


def _cluster_x_positions(xs: list[float], tol: float = 4.0) -> list[float]:
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
    if not bullet_xs:
        return []
    centers = _cluster_x_positions(bullet_xs, tol=4.0)
    levels = []
    for x in bullet_xs:
        idx = min(range(len(centers)), key=lambda i: abs(x - centers[i]))
        levels.append(max(0, min(idx, 2)))
    return levels


# ----------------------------
# POSTPROCESSING ONLY
# ----------------------------
def postprocess_formatting(docx_path: str) -> None:
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


def convert(pdf_path: str, out_docx_path: str, mode: str = "bullets_only") -> None:
    pdf = fitz.open(pdf_path)
    doc = Document()
    set_aptos_12(doc)

    for page_index in range(pdf.page_count):
        page = pdf.load_page(page_index)

        raw = page.get_text("text") or ""
        lines = normalize_lines(raw)
        if not lines:
            continue

        bullet_xs = _extract_bullet_x_positions(page)
        bullet_levels = _levels_for_bullets_on_page(bullet_xs)
        bullet_level_idx = 0

        line_xs = _extract_line_x_positions(page)
        line_levels = _levels_for_bullets_on_page(line_xs)
        line_level_idx = 0

        # Slide title:
        # - bullets_only: original heuristic behaviour (unchanged)
        # - all_lines: title is first line by (y,x) coordinates (footer excluded by y)
        if mode == "all_lines":
            title = _extract_first_line_text(page)
        else:
            title = None
            for ln in lines:
                if bullet_text(ln) is not None:
                    continue
                if looks_like_heading(ln) or len(ln) <= 60:
                    title = ln
                    break

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

        if mode == "all_lines":
            # ONLY CHANGE REQUESTED: ignore footer content at bottom of slide (by y coordinate)
            coord_lines = _extract_lines_xy_text_excluding_footer(page)

            # Remove the title line from the list we bullet (if present)
            if title:
                coord_lines = [t for t in coord_lines if t[2] != title]

            # Recompute levels for the remaining lines on this page so indices align
            coord_xs = [x0 for _, x0, __ in coord_lines]
            coord_lvls = _levels_for_bullets_on_page(coord_xs)
            coord_idx = 0

            for _, __, txt in coord_lines:
                flush_bullet()
                lvl = coord_lvls[coord_idx] if coord_idx < len(coord_lvls) else 0
                coord_idx += 1

                bt = bullet_text(txt)
                text_out = bt if bt is not None else txt
                add_bullet(doc, text_out, lvl)

            flush_bullet()
            doc.add_paragraph("")
            continue

        # ---- bullets_only (original behaviour) ----
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

            if current_bullet:
                current_bullet += " " + ln

        flush_bullet()
        doc.add_paragraph("")

    doc.save(out_docx_path)

    # Run postprocessing (as before) for both modes
    postprocess_formatting(out_docx_path)

    print(f"Saved (formatted): {out_docx_path}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf", help="Input slides PDF")
    ap.add_argument("out", help="Output docx")
    args = ap.parse_args()
    convert(args.pdf, args.out)


if __name__ == "__main__":
    main()
