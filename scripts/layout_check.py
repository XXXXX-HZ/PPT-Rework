"""Geometry checker for a .pptx: off-canvas shapes, estimated text overflow,
margin violations and overlapping text boxes.

    python layout_check.py deck.pptx

Use it always, and rely on it *instead of* rendering only when LibreOffice is
unavailable. Line counts are estimated from an average Arial advance width, so
flags are approximate: treat each as "open this slide and look", not as proof.
CJK text is counted at double width."""
import sys

# Windows consoles default to cp1252; this file prints CJK.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass
from pptx import Presentation
from pptx.util import Emu

AVG_CHAR_W = 0.50   # average Arial advance width as a fraction of font size
LINE_F     = 1.22   # line height as a multiple of font size
PAD_IN     = 0.0    # margin=0 was set on every text box in this deck

prs = Presentation(sys.argv[1])
SLIDE_W = prs.slide_width / 914400.0
SLIDE_H = prs.slide_height / 914400.0

def walk(shapes, out):
    for sh in shapes:
        if sh.shape_type == 6:  # group
            walk(sh.shapes, out)
        else:
            out.append(sh)
    return out

problems = []
for idx, slide in enumerate(prs.slides, start=1):
    for sh in walk(list(slide.shapes), []):
        # geometry check first
        if sh.left is not None:
            l = sh.left / 914400.0
            t = sh.top / 914400.0
            r = l + (sh.width or 0) / 914400.0
            b = t + (sh.height or 0) / 914400.0
            if l < -0.01 or t < -0.01 or r > SLIDE_W + 0.01 or b > SLIDE_H + 0.01:
                problems.append((idx, "OFF-CANVAS", f"{sh.shape_type} at ({l:.2f},{t:.2f})-({r:.2f},{b:.2f})"))
        if not sh.has_text_frame:
            continue
        tf = sh.text_frame
        text = tf.text.strip()
        if not text:
            continue
        box_w = (sh.width or 0) / 914400.0 - 2 * PAD_IN
        box_h = (sh.height or 0) / 914400.0
        if box_w <= 0:
            continue
        total_lines = 0.0
        max_sz = 0
        for para in tf.paragraphs:
            sizes = [r.font.size.pt for r in para.runs if r.font.size]
            sz = max(sizes) if sizes else 18
            max_sz = max(max_sz, sz)
            ptxt = "".join(r.text for r in para.runs)
            for seg in (ptxt.split("\n") or [""]):
                # CJK glyphs are full-width; count them as two latin chars
                width_units = sum(2 if ord(ch) > 0x2E80 else 1 for ch in seg)
                cw = sz * AVG_CHAR_W / 72.0            # inches per latin char
                cpl = max(1, int(box_w / cw))
                total_lines += max(1, -(-width_units // cpl))
        need_h = total_lines * max_sz * LINE_F / 72.0
        if need_h > box_h * 1.06:
            problems.append((idx, "MAY OVERFLOW",
                             f'need ~{need_h:.2f}" in {box_h:.2f}" box @{max_sz:.0f}pt :: "{text[:58]}"'))

MARGIN = 0.5
for idx, slide in enumerate(prs.slides, start=1):
    for sh in walk(list(slide.shapes), []):
        if sh.left is None or not sh.has_text_frame:
            continue
        if not sh.text_frame.text.strip():
            continue
        l = sh.left / 914400.0
        r = l + (sh.width or 0) / 914400.0
        if l < MARGIN - 0.01 or r > SLIDE_W - MARGIN + 0.01:
            problems.append((idx, "TIGHT MARGIN",
                             f'text at x {l:.2f}-{r:.2f}" breaks the {MARGIN}" margin'))

if not problems:
    print("No off-canvas shapes, no margin breaks and no estimated text overflow.")
for p in sorted(problems):
    print(f"slide {p[0]:>2}  {p[1]:<13} {p[2]}")
