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
# COORDINATE-BASED BULLET LEVELS (existing)
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
                if not st or not st.strip():
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
# NEW: COORDINATE-BASED LINE EXTRACTION (for "all_lines" mode)
# ----------------------------
def _extract_text_lines_with_coords(page: fitz.Page) -> list[dict]:
    """
    Extract text lines with coordinates using page.get_text("dict").
    Returns a list of dicts:
      {"text": str, "x0": float, "y0": float, "size": float}
    in stable reading order.

    size = max font size among spans on that line.
    """
    d = page.get_text("dict")
    out = []

    for block in d.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue

            text = "".join(s.get("text", "") for s in spans)
            text = re.sub(r"\s+", " ", (text or "").strip())
            if not text or is_footer_noise(text):
                continue

            bbox = line.get("bbox")
            if not bbox:
                continue

            x0 = float(bbox[0])
            y0 = float(bbox[1])

            # max font size among spans (robust title/heading detection)
            max_size = 0.0
            for s in spans:
                sz = s.get("size")
                if isinstance(sz, (int, float)):
                    if float(sz) > max_size:
                        max_size = float(sz)

            out.append({"text": text, "x0": x0, "y0": y0, "size": max_size})

    out.sort(key=lambda r: (r["y0"], r["x0"]))
    return out


def _levels_for_line_xs(xs: list[float], tol: float = 4.0) -> list[int]:
    if not xs:
        return []
    centers = _cluster_x_positions(xs, tol=tol)
    levels = []
    for x in xs:
        idx = min(range(len(centers)), key=lambda i: abs(x - centers[i]))
        levels.append(max(0, min(idx, 2)))
    return levels


def _looks_like_course_header(text: str) -> bool:
    # Common repeating header on your slides, e.g. "CS262 Logic and Verification"
    t = (text or "").strip()
    return bool(re.match(r"^CS\d{3}\b", t))


# ----------------------------
# POSTPROCESSING ONLY (unchanged)
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
    """
    mode:
      - "bullets_only" (default): existing behaviour (only lines with bullet glyphs become bullets)
      - "all_lines": treat every non-heading line as a bullet, with levels inferred from x-coordinates
    """
    mode = (mode or "bullets_only").strip().lower()
    if mode not in {"bullets_only", "all_lines"}:
        raise ValueError("mode must be 'bullets_only' or 'all_lines'")

    pdf = fitz.open(pdf_path)
    doc = Document()
    set_aptos_12(doc)

    for page_index in range(pdf.page_count):
        page = pdf.load_page(page_index)

        # ----------------------------
        # EXISTING MODE (UNCHANGED)
        # ----------------------------
        if mode == "bullets_only":
            raw = page.get_text("text") or ""
            lines = normalize_lines(raw)
            if not lines:
                continue

            bullet_xs = _extract_bullet_x_positions(page)
            bullet_levels = _levels_for_bullets_on_page(bullet_xs)
            bullet_level_idx = 0

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
            continue

        # ----------------------------
        # ALL_LINES MODE (improved title/heading detection using font sizes)
        # ----------------------------
        recs = _extract_text_lines_with_coords(page)
        if not recs:
            continue

        # Determine "large font" threshold for headings on this page
        max_size = max((r["size"] for r in recs), default=0.0)
        # Headings are usually the top lines with the largest font; allow small tolerance
        heading_size_cutoff = max_size - 0.6  # ~same font size as title

        # Candidate title lines: near top AND large font
        top_band = [r for r in recs if r["y0"] <= 150 and r["size"] >= heading_size_cutoff]

        title_idx = None
        if top_band:
            # If first is a course header (CS### ...) and second is also large-font nearby, prefer second
            # This fixes missing real slide titles.
            # Find their indices in recs
            idxs = [recs.index(r) for r in top_band]
            idxs.sort()

            if len(idxs) >= 2:
                i0, i1 = idxs[0], idxs[1]
                if _looks_like_course_header(recs[i0]["text"]) and abs(recs[i1]["y0"] - recs[i0]["y0"]) <= 40:
                    title_idx = i1
                else:
                    title_idx = i0
            else:
                title_idx = idxs[0]
        else:
            # fallback to old logic (but only if needed)
            for i, r in enumerate(recs):
                ln = r["text"]
                if looks_like_heading(ln) or len(ln) <= 60:
                    title_idx = i
                    break

        title = recs[title_idx]["text"] if title_idx is not None else None

        if title and title.strip().lower() in {"outline", "summary"}:
            continue

        if title:
            add_bold_line(doc, title)

        # Decide which non-title lines are headings:
        # - looks_like_heading(text) OR
        # - large font (same family as title)
        def is_heading_line(i: int, r: dict) -> bool:
            if title_idx is not None and i == title_idx:
                return False
            txt = r["text"]
            if looks_like_heading(txt):
                return True
            if r["size"] >= heading_size_cutoff and r["y0"] <= 250:
                # large-font headings usually live in the top half; avoid footer/header junk
                return True
            return False

        # Build x-list for bullet candidates: all non-heading lines (excluding title)
        bullet_candidate_xs = []
        bullet_candidate_idxs = []
        for i, r in enumerate(recs):
            if title_idx is not None and i == title_idx:
                continue
            if is_heading_line(i, r):
                continue
            bullet_candidate_xs.append(r["x0"])
            bullet_candidate_idxs.append(i)

        levels = _levels_for_line_xs(bullet_candidate_xs, tol=4.0)
        idx_to_level = {bullet_candidate_idxs[k]: levels[k] for k in range(len(levels))}

        # Emit
        for i, r in enumerate(recs):
            if title_idx is not None and i == title_idx:
                continue

            if is_heading_line(i, r):
                add_bold_line(doc, r["text"])
            else:
                lvl = idx_to_level.get(i, 0)
                add_bullet(doc, r["text"], lvl)

        doc.add_paragraph("")

    doc.save(out_docx_path)
    postprocess_formatting(out_docx_path)
    print(f"Saved (formatted): {out_docx_path}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf", help="Input slides PDF")
    ap.add_argument("out", help="Output docx")
    ap.add_argument(
        "--mode",
        default="bullets_only",
        choices=["bullets_only", "all_lines"],
        help="Conversion mode (default: bullets_only)",
    )
    args = ap.parse_args()
    convert(args.pdf, args.out, mode=args.mode)


if __name__ == "__main__":
    main()
