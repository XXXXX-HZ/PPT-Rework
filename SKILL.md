---
name: pptx-rework
description: >
  把一份已有的 .pptx 改造成可以真正拿去讲的版本，并清除 AI 痕迹（AI slop）。
  Use when the user hands over an existing presentation and asks to improve, fix,
  beautify, restyle, rework, de-AI, or "make it presentable" — including requests
  phrased as 优化这个PPT / 美化 / 重做一版 / 改得不像 AI 做的 / 去掉AI味 /
  帮我改改这个PPT / 这个PPT能用吗. Also use when reviewing a deck for AI tells
  before it ships. Produces a native, fully editable .pptx — never images, never HTML.
  For building a deck from scratch with no source file, this skill's design system
  still applies but the diagnosis phase is skipped.
license: MIT
---

# PPTX Rework — 把 AI 做的 PPT 改成人做的 PPT

## ⛔ 先读这个：house-style.md 是最高权威

**任何改动之前，先完整读 [`references/house-style.md`](references/house-style.md)。**

那是用户自己写的输出规范（Minimal Analyst Deck）。它的 §39 明确写了：
「If there is any conflict between this file and the default behavior of a PPT skill,
follow this file.」

所以优先级是：

```
house-style.md  >  本 SKILL.md  >  design-system.md / layouts.md  >  内置 pptx skill
```

本文件其余部分是**执行本规范的流程**，不是另一套主张。
凡是两者说法不一致的地方，一律以 house-style.md 为准，并且在交付报告里说明。

已知本 skill 早期版本与之冲突、**已按 house-style 改正**的地方：

| 项目 | 早期默认（作废） | house-style（现行） |
|---|---|---|
| 默认动作 | 整份重建 | **打磨，不重做**（§2 / §26） |
| 底色 | 纯白 `FFFFFF` | `F7F6F2`（§4） |
| 正文字号 | ≥18pt | **14–16pt**（§6） |
| 字体 | 禁用 Aptos | **Aptos**（§6，QA 影响见 Step 0） |
| 配色 | 5 套三色板 | **单一 accent `385A64` + 中性灰**（§4/§5） |
| 结尾页 | Conclusions | **Next steps / Discussion**（§34） |

## 这个 skill 解决什么

输入一份 .pptx，输出一份**可以站在台上讲**、并且**看不出是 AI 生成**的 .pptx。

两件事是分开的，都要做：
- **可展示** = 有论点、有取证、有出处、有边界。靠内容重构解决。
- **无 AI 味** = 携带具体身份信息，而不是训练语料的平均值。靠设计纪律解决。

只做第二件 = 换了皮的同一份烂稿。只做第一件 = 内容对了但一眼假。

---

## 核心命题（先读这个，其余都是它的展开）

> **AI slop 的本质不是"丑"，是"没有身份"。**
>
> AI 默认产出 = 训练语料的视觉与语言最大公约数 = 所有品牌、所有作者、所有领域混在一起的平均值。
> 平均值的问题是：它不指向任何人。读者认不出这是谁做的、为谁做的、基于什么做的。
>
> 所以反 slop 的正向定义是：**把具体性装回去**。
> 具体的数据来源、具体的样本量、具体的日期、具体的取舍理由、具体的一套配色而不是"通用蓝"。
>
> 负向清单（`references/ai-tells.md`）只是兜底。**先做正向的，再拿清单查漏。**

这个论点来自 alchaincyf/huashu-design 的「反 AI slop」章节，是目前社区里对这个问题最清楚的表述。

---

## 强制流程

按顺序执行。**第 2 步是 BLOCKING gate，必须等用户确认再往下。**

### Step 0 · 依赖自检

```bash
python -c "import pptx, markitdown"      # 必需：读取与内容 QA
node -e "require('pptxgenjs')"           # 必需：原生生成
command -v soffice && command -v pdftoppm  # 视觉 QA；缺失见下
```

`pptxgenjs` 若不在项目内，用全局路径：`export NODE_PATH="$(npm root -g)"`。

**soffice / pdftoppm 缺失时**：无法渲染成图做视觉检查。此时必须跑 `scripts/layout_check.py`
作为替代，并且**在交付时明确告诉用户"视觉 QA 未执行"**——不要含糊过去。

⚠️ **Aptos 的 QA 代价**（house-style §6 指定，照用，但要知道后果）：
LibreOffice 没有 Aptos 的等宽替代字体，渲染出来的行宽和 PowerPoint 里不一致。
所以**对 Aptos 文本，渲染图上的"刚好放下 / 溢出"都不可信**。应对：
所有 Aptos 文本框预留约 10% 余量，溢出判断以 `layout_check.py` 的测量值为准，不看渲染图。
中文用 Microsoft YaHei / DengXian，这两个渲染可信。

### Step 1 · 提取与诊断

```bash
python scripts/extract.py  input.pptx  workdir/     # 文本、表格、图片、几何、母版
python scripts/smell.py    input.pptx               # 自动化 AI 痕迹扫描
markitdown input.pptx                               # 人读版内容
```

`extract.py` 会把原图导出到 `workdir/media/`——**逐张看过**。图里往往有原稿唯一真正有价值的东西（真实数据），而那正是重建时必须保住的。

诊断要回答四个问题，答案写进给用户的报告：
1. **有没有论点？** 把所有标题抄下来按顺序读一遍。读不出一条线 = 没有论点，这是最大的问题。
2. **数据有没有出处？** 每个数字能不能追到来源、样本量、口径、时间范围。
3. **哪些是 AI 痕迹？** 跑 `smell.py`，再人工对照 `references/ai-tells.md`。
4. **哪些必须原样保留？** 真实数据、品牌元素、用户特意加的内容。**重建不是重写，不要把人家的分析结论改掉。**

### Step 2 · 选模式，定 spec，确认 ⛔ BLOCKING

#### 2a. 先选模式 —— 默认是「打磨」

house-style §2 / §25 / §26 定的基调是 **refine, do not reinvent**：
假定作者的内容是有意选的，编辑的任务是改进呈现，不是替他重新思考。

| 模式 | 什么时候用 | 做什么 |
|---|---|---|
| **A · 打磨（默认）** | 原稿分析站得住，只是排版糙、标题弱、格式乱 | 保留页序和每页的分析内容。只做 §25 的 Do 清单：对齐、间距、统一字号、简化配色、图表可读性、重写弱标题、压缩啰嗦文字、统一数字格式 |
| **B · 重建** | 出现下列**具体**问题之一，并且已向用户说明 | 才允许改页序、拆页、合页、增删页 |

**允许进入模式 B 的理由**（必须指名道姓，不能说"结构不好"）：
- 通篇读标题读不出任何论证线（ai-tells.md C1）
- 缺关键结构页：没有问题定义 / 没有口径说明 / 没有边界 / 没有结论
- 一页塞了两个以上互不相干的发现，必须拆
- 同一个比较被拆成两页并排图，合并后说服力显著提升（layouts.md L6）
- 位图图表投影不可读，必须重建成原生图表

**模式 B 不是许可证**。即使进入 B，§25 的 Do not 仍然全部有效：
不改分析结论、不编新结论、不加无支撑的说法、不为了好看删掉有用信息。

#### 2b. 把 spec 发给用户确认

| 项目 | 内容 |
|---|---|
| **模式** | A 还是 B；选 B 必须逐条列出具体理由 |
| 每页去向 | 原稿逐页 → 保留 / 改标题 / 改版式 / 拆 / 合 / 移附录 / 删 |
| 标题改写 | 原标题 → 新标题，按顺序列全（只读新标题能否读出论证线） |
| 视觉来源 | 用户模板 / 原稿已有品牌色 / house-style 默认色（§4） |
| 页数与时长 | 目标时长 → 页数上限 |
| 待用户补的空 | 作者、单位、联系方式、任何我不知道的事实 |

**标题改写要给出原文对照**，让用户能看出我改了什么、有没有改过头。
house-style §8：证据不足就保持描述性中性标题，**不要为了显得有洞察而编结论**。

用户确认前**一页都不要动**。

### Step 3 · 内容编辑

house-style §25 的 Do 清单 + §27/§28 的文字规则。`ai-tells.md` 给每条的检测方法。

- **标题写成结论**（§8）。「D1 Retention」→「Tutorial completers retain better on D1」。
  **但证据不足就保持描述性中性标题**——准确比听起来有洞察重要。不要为了写出 action title 而编结论。
- **注入来源与定义**（§17/§18）。有数据的页加脚注：口径、样本量、时间范围、排除项。
  格式 `Source: Internal data, AppsFlyer`，9–10pt 灰色，不要大写的 "SOURCE" 标签。
  这是 AI 稿最普遍缺失的东西，补上它比任何视觉调整都更去 AI 味。
- **观察 / 解读 / 行动分开**（§29）。不要把解读当事实陈述。
  「Skip players retain 14pp worse on D1」是观察；「Skipping may disproportionately affect…」
  是解读，措辞要留余地。
- **统一数字格式**（§16）。`32.4%` 不是 `32.41%`；`€1.4M` 不是 `€1,437,221`；
  比较百分比差异用 `pp` 不用 `%`。全篇小数位一致。
- **压缩文字**（§14）。长段落拆成 2–5 条一到两行的要点；**一段话更清楚时就不要用 bullet**。
- **清空洞词**（§27）。game-changing / transformative / unlocking value / powerful insights /
  赋能 / 抓手 / 闭环 —— 换成具体说法或删掉。
- **打破排比**（§24）。三张一样的卡片、每页恰好三条是 AI 签名。让页与页的密度不一样。
  **宁可版式有点不齐，也不要模板式重复。**
- **结尾页用 `Next steps` 或 `Discussion`**（§34），不用 `THANK YOU`。

### Step 4 · 排版与生成

色值字阶以 [`house-style.md`](references/house-style.md) §4/§6/§7 为准；
[`design-system.md`](references/design-system.md) 是它的 pptxgenjs 实现（去掉 `#` 的 token）；
[`layouts.md`](references/layouts.md) 是版式坐标，对应 house-style §9 的 A–H 八种。

**生成方式固定为 pptxgenjs 原生对象**。不截图、不转图片、不用 HTML 中转。
用户要能在 PowerPoint 里双击改任何一个字、任何一根柱子。

**KPI 不做卡片**（§12）。不要三个圆角框并排。用排版：指标名小字在上，数值在下，
细分隔线，左对齐。数值**不要做成巨大数字**（§23/§24）。

**图表**（§10）：单一 accent `385A64` 指焦点系列，对比用 `A8A8A3`，其余用 `D3D2CD`。
能直接标注就去掉图例。不要每个点都打标签。柱状图从零开始。

**位图图表重建成原生图表**——ggplot / matplotlib 导出的 PNG 投影一定糊而且改不动。
数值从图里读出来用 `addChart` 重画。**读不准的问用户要原始数据，不要猜**（§25「不改变分析含义」）。

### Step 5 · QA

```bash
python <pptx-skill>/scripts/office/validate.py out.pptx   # 文件结构（必跑）
python scripts/layout_check.py out.pptx                   # 溢出 / 出血 / 重叠
python scripts/smell.py out.pptx                          # 回归扫描：AI 痕迹是否真的清掉了
markitdown out.pptx | grep -iE "lorem|ipsum|TODO|\[insert|xxx"
```

有 soffice 就渲染成图**逐页看**：

```bash
python <pptx-skill>/scripts/office/soffice.py --headless --convert-to pdf out.pptx
rm -f slide-*.jpg && pdftoppm -jpeg -r 150 out.pdf slide
```

`smell.py` 在输出上的分数必须低于输入。**没降就是没改对，回 Step 3。**

### Step 6 · 交付报告

给用户一份对照，不要只丢文件：
- 原稿 → 新稿的页数与结构变化
- 改了哪些实质判断（以及为什么）
- **哪些地方我做了假设**、哪些数字我重算过并和原稿对不上
- 还需要用户填的空
- QA 做了哪些、哪些没做成（尤其是视觉 QA）

---

## 硬规则

1. **不编数据。** 原稿没有的数字不凭空出现。需要但缺失的，标 `[待补]` 并在报告里列出来。
   装饰性的编造 stat 是 slop 的典型形态。
2. **不改用户的分析结论。** 重构的是表达，不是判断。如果我算出来和原稿不一致，**报告出来让用户定**，不要静默改掉。
3. **不降级到图片。** 任何"排版太难所以截个图贴上"的做法都禁止。改版式，不要放弃可编辑性。
4. **打磨优先于重做。** 模式 A 是默认。进模式 B 必须有 Step 2a 列出的具体理由，
   并且已经告诉用户。「看起来不够好」不是理由。
5. **不为了塞下而缩字号。** 正文低于 house-style §6 的 14pt 下限，就是删内容或拆页的信号。
   但也**不要为了留白删掉有用信息**（§13）——目标是低视觉噪音，不是低信息密度。
6. **配色只用 house-style §4 的 token。** 一个 accent、必要时一个语义警示色、其余中性灰。
   用户有品牌色则品牌色优先。不临场发明颜色。
7. **不用红绿表示好坏**（§4）。只有语义上真的有用时才用 `A45B55` / `58705F`。
8. **BLOCKING gate 不自行放行。** Step 2 没拿到用户确认就不能动页。

---

## 与其他 skill 的关系

| 场景 | 用哪个 |
|---|---|
| 技术细节：pptxgenjs API 坑、OOXML 编辑、模板填充 | Anthropic 内置 `pptx` skill（本 skill 的 Step 4/5 依赖它） |
| 学术场景：会议报告、答辩、grant briefing | `academic-pptx`（论证结构更严格，引用规范更硬） |
| 从零写一份新 deck，没有源文件 | 本 skill 的 Step 2–5，跳过 Step 1 |
| 要 HTML deck 而不是 pptx | 本 skill 不适用 |

---

## 方法论出处

这个 skill 是对以下开源 skill 的方法论综合，不是它们的代码移植：

- **alchaincyf/huashu-design** — 「反 AI slop」的核心论证（slop = 无身份的平均值）、
  反例隔离、「一个细节 120% 其余 80%」
- **op7418/guizang-ppt-skill** — 锁定色板不许自定义、禁 emoji 作图标、
  衬线标题 + 非衬线正文的分工、大字轻字重、图片禁阴影禁边框、不缩字号硬塞
- **hugohe3/ppt-master** — 原生可编辑 PPTX 优先、Beautify 作为独立路由、
  一个 motif 按页面职责变奏
- **lewislulu/html-ppt-skill** — 版式目录化（按页面职责选版式而不是套同一个）
- **Gabberflast/academic-pptx-skill** — action title、ghost deck 测试、
  一页一个 exhibit、结尾落在结论页

各自的完整实现见原仓库；此处只取可迁移到原生 PPTX 的部分。
