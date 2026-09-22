#!/usr/bin/env python
"""Automated AI-tell scanner for a .pptx.

    python smell.py deck.pptx [--verbose]

Prints itemised findings keyed to references/ai-tells.md and a total score.
Run it on the INPUT deck and again on the OUTPUT deck: the output's score must
be lower, or the rework did not actually change anything that matters.

The score is a triage aid, not a verdict. Low score + no argument still fails —
"is there an argument" and "do the titles chain" cannot be detected mechanically.

Cover and divider slides are classified and exempted from the checks that only
make sense on content slides (short titles, heavy display type, provenance).
"""
import re
import sys

# Windows consoles default to cp1252; this file prints CJK.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from pptx import Presentation

EMU = 914400.0

# --- A5: the generic "SaaS blue" cluster and purple-gradient stand-ins -------
SLOP_COLORS = {
    "2563EB", "3B82F6", "1E40AF", "1D4ED8", "60A5FA", "2E5FE8",
    "6366F1", "818CF8", "8B5CF6", "A78BFA", "7C3AED", "4F46E5",
}

# --- B1: titles that spend the most valuable line on a category name --------
LABEL_TITLES = {
    "overview", "introduction", "background", "agenda", "outline", "contents",
    "results", "findings", "data", "methodology", "methods", "analysis",
    "discussion", "conclusion", "conclusions", "summary", "key takeaways",
    "takeaways", "next steps", "recommendations", "thank you", "questions",
    "key insights", "insights", "highlights", "关键洞察", "核心洞察",
    "about us", "our team", "problem", "solution", "challenges", "opportunities",
    "概述", "简介", "背景", "目录", "议程", "现状", "现状分析", "数据分析",
    "结果", "研究方法", "方法论", "分析", "讨论", "结论", "总结", "小结",
    "关键要点", "核心要点", "下一步", "建议", "问题", "解决方案", "谢谢",
    "谢谢聆听", "感谢聆听", "提问环节", "团队介绍", "关于我们", "挑战", "机遇",
}

# --- B3: high-frequency, zero-information vocabulary ------------------------
FILLER = [
    "leverage", "leveraging", "streamline", "streamlining", "robust",
    "comprehensive", "seamless", "holistic", "empower", "empowering",
    "synergy", "synergies", "cutting-edge", "state-of-the-art", "best-in-class",
    "paradigm", "unlock", "unlocking", "game-chang", "revolutioniz",
    "赋能", "抓手", "闭环", "全方位", "多维度", "深度融合", "持续优化",
    "全面提升", "有效提升", "进一步加强", "夯实", "沉淀", "打法", "组合拳",
]

# Pictographs only. Arrows (U+2190-21FF) and typographic dashes are NOT emoji —
# including them fires on every "5.24% -> 7.29%" in a legitimate table.
EMOJI = re.compile(
    "[\U0001F300-\U0001FAFF\U0001F000-\U0001F0FF"
    "\U00002600-\U000026FF\U00002700-\U000027BF✅❌⭐️]"
)
HAS_NUM = re.compile(r"\d[\d,.]*\s*%|[$￥€]\s*[\d,.]+|\b\d{2,}\b")
# A bare 4-digit year or a page number is not a claim that needs provenance.
ONLY_YEAR = re.compile(r"^\D*((19|20)\d{2}|\d{1,2})\D*$")
SOURCE_HINT = re.compile(
    r"source|sources|based on|measured|來源|来源|数据来源|口径|样本|统计|"
    r"\bn\s*=|\bN\s*=|cohort|截至|注[:：]|excl(uding|uded)|window|as of",
    re.I,
)
# B5 accepts any honest statement of scope or limitation, in either language.
LIMIT_HINT = re.compile(
    r"limitation|caveat|constraint|boundar|cannot say|can not say|not tested|"
    r"censor|assumption|directional|局限|边界|不能说明|说不了|注意事项|"
    r"未做|未经|仅供参考|不代表|前提|取舍",
    re.I,
)


def shape_fill_hex(sh):
    try:
        f = sh.fill
        if f.type is not None and f.type == 1:  # solid
            return str(f.fore_color.rgb).upper()
    except Exception:
        pass
    return None


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


def is_dark(hexcol):
    try:
        r, g, b = (int(hexcol[i:i + 2], 16) for i in (0, 2, 4))
        return (0.299 * r + 0.587 * g + 0.114 * b) < 110
    except Exception:
        return False


def slide_bg_hex(slide):
    """Best-effort read of an explicitly set solid slide background."""
    try:
        el = slide.background.fill
        if el.type == 1:
            return str(el.fore_color.rgb).upper()
    except Exception:
        pass
    return None


def analyse(path):
    prs = Presentation(path)
    SW = prs.slide_width / EMU
    findings = []          # (code, slide_no, detail)
    para_counts = []
    n_slides = len(prs.slides)

    for idx, slide in enumerate(prs.slides, start=1):
        shapes = list(iter_shapes(list(slide.shapes)))
        title_text, title_top = None, None
        has_number = has_source = False
        para_count = 0
        text_shapes = 0
        has_exhibit = False

        bg = slide_bg_hex(slide)
        # A cover or divider: dark ground, or a handful of text shapes and no
        # table/chart/picture. These legitimately carry short titles and big type.
        for sh in shapes:
            if sh.has_text_frame and sh.text_frame.text.strip():
                text_shapes += 1
            if sh.has_table or getattr(sh, "has_chart", False) or \
                    sh.__class__.__name__ == "Picture":
                has_exhibit = True
        structural = (bg is not None and is_dark(bg)) or \
                     (text_shapes <= 4 and not has_exhibit)

        # ---------- HS2: pure white ground (house-style §4) ----------
        # "Avoid pure white #FFFFFF as the dominant background unless required
        #  by an existing corporate template." Reported once per slide.
        if bg == "FFFFFF":
            findings.append(("HS2", idx, "纯白底 FFFFFF，house-style 默认 F7F6F2"))

        for sh in shapes:
            w = (sh.width or 0) / EMU
            h = (sh.height or 0) / EMU
            top = (sh.top or 0) / EMU
            fill = shape_fill_hex(sh)
            has_text = sh.has_text_frame and sh.text_frame.text.strip()

            # ---------- A1 / A2: decorative bars and accent stripes ----------
            if not has_text and w > 0 and h > 0 and fill:
                ratio = max(w, h) / max(min(w, h), 0.001)
                if ratio >= 12 and min(w, h) <= 0.22:
                    if w >= SW * 0.55:
                        findings.append(("A2", idx, f'通栏色条 {w:.1f}x{h:.2f}" fill={fill}'))
                    elif top < 2.0:
                        findings.append(("A1", idx, f'标题下装饰线 {w:.1f}x{h:.2f}" fill={fill}'))
                    else:
                        findings.append(("A2", idx, f'装饰条 {w:.1f}x{h:.2f}" fill={fill}'))

                # ---------- A3: icon circles ----------
                st = str(getattr(sh, "shape_type", "")).upper()
                nm = str(getattr(sh, "name", ""))
                if ("OVAL" in st or "Oval" in nm or "Ellipse" in nm) and \
                        0.15 <= w <= 1.2 and abs(w - h) < 0.12:
                    findings.append(("A3", idx, f'彩色圆底图标 d={w:.2f}" fill={fill}'))

            # ---------- A5: slop palette ----------
            if fill and fill in SLOP_COLORS:
                findings.append(("A5", idx, f"通用 SaaS 蓝/紫 填充 {fill}"))

            # ---------- A6: drop shadows ----------
            try:
                if sh.shadow is not None and sh.shadow.inherit is False:
                    findings.append(("A6", idx, "显式投影"))
            except Exception:
                pass

            if not sh.has_text_frame:
                continue
            tf = sh.text_frame
            t = tf.text.strip()
            if not t:
                continue

            if top is not None and (title_top is None or top < title_top):
                title_top, title_text = top, t.split("\n")[0].strip()

            sizes = []
            for p in tf.paragraphs:
                body = "".join(r.text for r in p.runs).strip()
                if body:
                    para_count += 1
                for r in p.runs:
                    if r.font.size:
                        sizes.append(r.font.size.pt)
                    try:
                        if r.font.color and r.font.color.rgb and \
                                str(r.font.color.rgb).upper() in SLOP_COLORS:
                            findings.append(
                                ("A5", idx, f"通用 SaaS 蓝/紫 文字 {r.font.color.rgb}"))
                    except Exception:
                        pass

            # ---------- A4: emoji ----------
            if EMOJI.search(t):
                findings.append(("A4", idx, "emoji: " + "".join(EMOJI.findall(t))[:12]))

            # ---------- B3: filler vocabulary ----------
            low = t.lower()
            for f in FILLER:
                if f in low:
                    findings.append(("B3", idx, f"空洞词「{f}」"))
                    break

            # ---------- A9: centred body text ----------
            for p in tf.paragraphs:
                body = "".join(r.text for r in p.runs).strip()
                if len(body) > 45 and str(p.alignment) == "CENTER (2)":
                    findings.append(("A9", idx, f"正文居中: {body[:34]}…"))
                    break

            # ---------- A8: heavy display type on a CONTENT slide ----------
            if not structural:
                for p in tf.paragraphs:
                    line = "".join(x.text for x in p.runs).strip()
                    for r in p.runs:
                        if r.font.size and r.font.size.pt >= 32 and r.font.bold \
                                and len(line) > 30:
                            findings.append(("A8", idx, f"{r.font.size.pt:.0f}pt 粗体长文"))
                            break
                    break

            # ---------- SZ: body prose below house-style §6's 14pt floor ----
            # Only wide prose blocks. Narrow annotation columns, table cells,
            # chart labels (11-13pt) and footnotes (9-10pt) are legitimately smaller.
            prose = [s for s in sizes if 1 < s < 12]
            in_source_zone = top > 6.3          # footnotes belong at 9-10pt
            if prose and w > 6.0 and len(t) > 120 and not in_source_zone:
                findings.append(("SZ", idx, f"宽栏正文 {max(prose):.0f}pt < 12pt"))

            # ---------- HS1: pure black text (house-style §4) ----------
            for p in tf.paragraphs:
                for r in p.runs:
                    try:
                        if r.font.color and r.font.color.rgb and                                 str(r.font.color.rgb).upper() == "000000" and len(t) > 30:
                            findings.append(("HS1", idx, "纯黑 000000 文字，应为 242424"))
                            raise StopIteration
                    except StopIteration:
                        break
                    except Exception:
                        pass
                else:
                    continue
                break

            # ---------- HS3: ALL-CAPS (house-style §6 sentence case) --------
            letters = [c for c in t if c.isalpha() and ord(c) < 128]
            if len(letters) >= 12 and all(c.isupper() for c in letters):
                findings.append(("HS3", idx, f"全大写: {t[:32]}…"))

            # ---------- HS4: oversized title on a content slide (§6) --------
            if not structural:
                for p in tf.paragraphs:
                    for r in p.runs:
                        if r.font.size and r.font.size.pt > 30 and top < 1.6:
                            findings.append(
                                ("HS4", idx, f"正文页标题 {r.font.size.pt:.0f}pt > 30pt 上限"))
                            break
                    break

            if HAS_NUM.search(t) and not ONLY_YEAR.match(t.strip()):
                has_number = True
            if SOURCE_HINT.search(t):
                has_source = True
            if [s for s in sizes if s <= 14] and top > 6.3 and len(t) > 25:
                has_source = True   # a small low line is a source line by position

        # ---------- B1: label titles (content slides only) ----------
        if title_text and not structural:
            key = re.sub(r"[\s:：·\-—0-9.、]+", "", title_text).lower()
            if key in {re.sub(r"[\s:：·\-—0-9.、]+", "", x).lower() for x in LABEL_TITLES}:
                findings.append(("B1", idx, f"主题标签标题「{title_text}」"))
            elif (len(title_text) <= 14 and not re.search(r"[，。,.]", title_text)
                  and not HAS_NUM.search(title_text)):
                findings.append(("B1?", idx, f"标题疑似话题标签「{title_text}」"))

        # ---------- B4: numbers with no provenance (content slides only) -----
        if has_number and not has_source and not structural:
            findings.append(("B4", idx, "本页有数字但没有出处行"))

        para_counts.append(para_count)

    # ---------- B2: the rule-of-three ----------
    body_pages = [c for c in para_counts if c > 0]
    if len(body_pages) >= 5:
        threes = sum(1 for c in body_pages if c == 3)
        if threes / len(body_pages) >= 0.45:
            findings.append(("B2", 0, f"{threes}/{len(body_pages)} 页恰好 3 段——排比过度规整"))

    # ---------- B8: how it ends ----------
    if n_slides:
        last = list(prs.slides)[-1]
        lt = " ".join(sh.text_frame.text for sh in iter_shapes(list(last.shapes))
                      if sh.has_text_frame).strip()
        if not lt:
            findings.append(("B8", n_slides, "最后一页是空白页"))
        elif re.search(r"thank\s*you|谢谢|感谢|q\s*&\s*a|any questions", lt, re.I) \
                and len(lt) < 80:
            findings.append(("B8", n_slides, f"以「{lt[:24]}」结尾，不是结论页"))

    # ---------- B5: any statement of scope or limitation ----------
    blob = "\n".join(sh.text_frame.text for s in prs.slides
                     for sh in iter_shapes(list(s.shapes)) if sh.has_text_frame)
    if not LIMIT_HINT.search(blob):
        findings.append(("B5", 0, "全篇没有局限/边界的任何表述"))

    return findings, n_slides


WEIGHT = {"A1": 3, "A2": 3, "A3": 2, "A4": 3, "A5": 3, "A6": 2, "A8": 1, "A9": 2,
          "B1": 4, "B1?": 1, "B2": 3, "B3": 2, "B4": 3, "B5": 5, "B8": 4, "SZ": 2,
          "HS1": 2, "HS2": 2, "HS3": 2, "HS4": 2}

TITLES = {
    "A1": "标题装饰横线", "A2": "通栏/侧边色条", "A3": "彩色圆底图标",
    "A4": "emoji", "A5": "通用 SaaS 蓝/紫", "A6": "投影", "A8": "大字过重",
    "A9": "正文居中", "B1": "主题标签当标题", "B1?": "标题疑似话题标签",
    "B2": "每页恰好三条", "B3": "空洞词", "B4": "数字无出处",
    "B5": "无局限/边界表述", "B8": "以谢谢/空白结尾", "SZ": "宽栏正文过小",
    "HS1": "纯黑文字", "HS2": "纯白底", "HS3": "全大写", "HS4": "标题超 30pt",
}


def pad(s, n):
    """Pad to n display columns, counting CJK glyphs as two."""
    wide = sum(1 for ch in s if ord(ch) > 0x2E80)
    return s + " " * max(0, n - len(s) - wide)


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    path = sys.argv[1]
    verbose = "--verbose" in sys.argv or "-v" in sys.argv

    findings, n = analyse(path)
    groups = {}
    for code, slide, detail in findings:
        groups.setdefault(code, []).append((slide, detail))

    print(f"AI-tell scan  ::  {path}  ({n} slides)")
    print("=" * 68)
    if not findings:
        print("  没有命中任何自动化检测项。")
    score = 0
    for code in sorted(groups, key=lambda c: -WEIGHT.get(c, 1) * len(groups[c])):
        items = groups[code]
        score += WEIGHT.get(code, 1) * len(items)
        pages = sorted({s for s, _ in items if s})
        loc = ("p" + ",".join(str(p) for p in pages[:12])) if pages else "全篇"
        print(f"  [{pad(code, 3)}] {pad(TITLES.get(code, code), 18)} x{len(items):<3} {loc}")
        if verbose:
            for s, d in items[:14]:
                print(f"          p{s or '-'}: {d}")

    per = score / max(n, 1)
    band = "干净" if per < 1.0 else "轻度" if per < 2.5 else "明显" if per < 5 else "严重"
    print("=" * 68)
    print(f"总分 {score}   每页 {per:.1f}   ->  AI 痕迹：{band}")
    print()
    print("分数只是分诊。「有没有论证」「标题能不能串成一条线」测不出来，")
    print("必须人工过 references/ai-tells.md 的 C 组。")


if __name__ == "__main__":
    main()
