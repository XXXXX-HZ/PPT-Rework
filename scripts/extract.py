#!/usr/bin/env python
"""Dump everything needed to rebuild a deck from an existing .pptx.

    python extract.py input.pptx workdir/

Writes:
    workdir/content.md     per-slide text, tables, speaker notes, geometry
    workdir/media/         every embedded image, named by the slide that uses it
    workdir/palette.txt    colour usage frequency (for deriving the brand colour)

Look at every file in media/ before rebuilding. The source deck's only genuinely
valuable content is often locked inside a chart image, and that is exactly what
must survive the rework.
"""
import collections
import os
import shutil
import sys

# Windows consoles default to cp1252; this file prints CJK.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass
import zipfile

from pptx import Presentation

EMU = 914400.0


def iter_shapes(shapes):
    for sh in shapes:
        try:
            if sh.shape_type == 6:  # GROUP
                for inner in iter_shapes(sh.shapes):
                    yield inner
                continue
        except Exception:
            pass
        yield sh


def fill_hex(sh):
    try:
        f = sh.fill
        if f.type == 1:
            return str(f.fore_color.rgb).upper()
    except Exception:
        pass
    return None


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(2)
    src, out = sys.argv[1], sys.argv[2]
    media = os.path.join(out, "media")
    os.makedirs(media, exist_ok=True)

    prs = Presentation(src)
    SW, SH = prs.slide_width / EMU, prs.slide_height / EMU
    colors = collections.Counter()
    lines = [
        f"# Extracted from `{os.path.basename(src)}`",
        "",
        f"- canvas: {SW:.3f} x {SH:.3f} in"
        f"  ({'16:9' if abs(SW / SH - 16 / 9) < 0.02 else f'{SW/SH:.2f}:1'})",
        f"- slides: {len(prs.slides)}",
        "",
    ]

    img_no = 0
    for idx, slide in enumerate(prs.slides, start=1):
        lines += ["", "---", "", f"## Slide {idx}", ""]
        shapes = sorted(iter_shapes(list(slide.shapes)),
                        key=lambda s: ((s.top or 0), (s.left or 0)))
        for sh in shapes:
            x = (sh.left or 0) / EMU
            y = (sh.top or 0) / EMU
            w = (sh.width or 0) / EMU
            h = (sh.height or 0) / EMU
            geo = f"[{x:.2f},{y:.2f} {w:.2f}x{h:.2f}]"

            f = fill_hex(sh)
            if f:
                colors[f] += 1

            if sh.shape_type == 13 or sh.__class__.__name__ == "Picture":
                try:
                    blob = sh.image.blob
                    ext = sh.image.ext
                    img_no += 1
                    name = f"slide{idx:02d}_img{img_no:02d}.{ext}"
                    with open(os.path.join(media, name), "wb") as fh:
                        fh.write(blob)
                    lines.append(f"- IMAGE {geo} -> `media/{name}`")
                except Exception as e:
                    lines.append(f"- IMAGE {geo} (extract failed: {e})")
                continue

            if sh.has_table:
                lines.append(f"- TABLE {geo}")
                tbl = sh.table
                for r in tbl.rows:
                    cells = [c.text.replace("\n", " ").strip() for c in r.cells]
                    lines.append("  | " + " | ".join(cells) + " |")
                continue

            if getattr(sh, "has_chart", False):
                ch = sh.chart
                lines.append(f"- CHART {geo} type={ch.chart_type}")
                try:
                    cats = list(ch.plots[0].categories)
                    lines.append(f"  categories: {cats}")
                    for s in ch.series:
                        lines.append(f"  series {s.name}: {list(s.values)}")
                except Exception:
                    pass
                continue

            if sh.has_text_frame and sh.text_frame.text.strip():
                tf = sh.text_frame
                sizes, fonts, cols, bolds = set(), set(), set(), set()
                for p in tf.paragraphs:
                    for r in p.runs:
                        if r.font.size:
                            sizes.add(round(r.font.size.pt))
                        if r.font.name:
                            fonts.add(r.font.name)
                        if r.font.bold is not None:
                            bolds.add(bool(r.font.bold))
                        try:
                            if r.font.color and r.font.color.rgb:
                                c = str(r.font.color.rgb).upper()
                                cols.add(c)
                                colors[c] += 1
                        except Exception:
                            pass
                meta = []
                if sizes:
                    meta.append(f"{sorted(sizes)}pt")
                if fonts:
                    meta.append("/".join(sorted(fonts)))
                if cols:
                    meta.append("#" + ",#".join(sorted(cols)))
                if True in bolds:
                    meta.append("bold")
                lines.append(f"- TEXT {geo} {' '.join(meta)}")
                for ln in tf.text.strip().split("\n"):
                    lines.append(f"    {ln}")

        if slide.has_notes_slide:
            nt = slide.notes_slide.notes_text_frame.text.strip()
            if nt:
                lines += ["", f"  NOTES: {nt}"]

    # media the shape walk missed (backgrounds, grouped art)
    with zipfile.ZipFile(src) as z:
        for n in z.namelist():
            if n.startswith("ppt/media/"):
                base = os.path.basename(n)
                dst = os.path.join(media, "_pkg_" + base)
                if not os.path.exists(dst):
                    with z.open(n) as s, open(dst, "wb") as d:
                        shutil.copyfileobj(s, d)

    os.makedirs(out, exist_ok=True)
    with open(os.path.join(out, "content.md"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")
    with open(os.path.join(out, "palette.txt"), "w", encoding="utf-8") as fh:
        fh.write("colour usage (most frequent first)\n")
        for c, n in colors.most_common(30):
            fh.write(f"  {c}   x{n}\n")

    imgs = sorted(os.listdir(media))
    print(f"wrote {out}/content.md   ({len(lines)} lines)")
    print(f"wrote {out}/palette.txt  ({len(colors)} distinct colours)")
    print(f"wrote {out}/media/       ({len(imgs)} files)")
    if imgs:
        print("\n看完每一张图再动手重建：")
        for i in imgs[:24]:
            print(f"  {os.path.abspath(os.path.join(media, i))}")


if __name__ == "__main__":
    main()
