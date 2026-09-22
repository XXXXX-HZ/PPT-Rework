# Design System — house-style.md 的 pptxgenjs 实现

> **这个文件没有主张。** 它把 [`house-style.md`](house-style.md) §4–§7、§10–§12、§16–§22
> 翻译成可以直接写进 pptxgenjs 的常量和参数。
> 两者不一致时以 house-style.md 为准，并且把不一致报告给用户。

画布 **13.333 × 7.5 英寸**（`pres.layout = 'LAYOUT_WIDE'`，16:9）。
坐标单位英寸。颜色为**不带 `#` 的六位十六进制**——带 `#` 或写成八位（含 alpha）会损坏文件。

---

## 1. Token（house-style §4）

```js
// ---- 底色 ----
const PAPER    = 'F7F6F2';   // §4 默认底色。不要用 FFFFFF
const PAPER_2  = 'EFEDE7';   // §4 分节页 / 偶尔的视觉分隔，仅此两处

// ---- 文字 ----
const TEXT     = '242424';   // §4 主文字。不要用 000000
const TEXT_2   = '686868';   // §4 次级文字
const MUTED    = '8A8A86';   // §4 脚注 / 注释 / 页码

// ---- 线 ----
const DIVIDER  = 'D9D7D1';   // §4 分隔线、表格分隔线
const GRIDLINE = 'E4E2DC';   // §4 图表网格线

// ---- 唯一 accent ----
const ACCENT   = '385A64';   // §4 主系列 / 主结论 / 最重要的数字 / 被选中的组
                             //    §4「Do not use the accent color decoratively」

// ---- 对比序列 ----
const COMPARE  = 'A8A8A3';   // §4 对比组
const OTHER    = 'D3D2CD';   // §4 背景序列
const NEG      = 'A45B55';   // §4 仅在语义上确实需要时
const POS      = '58705F';   // §4 仅在语义上确实需要时
```

**用色哲学（§5）**：一页上只有一个东西被强调，其余全部安静下来。
同一页不出现两个 accent 色。不用红绿表示「好/坏」。

---

## 2. 字体（house-style §6）

```js
const FONT    = 'Aptos';            // §6 英文与数字
const FONT_CN = 'Microsoft YaHei';  // §6 中文（或 'DengXian'）
```

**原稿已经在一致地用某个专业字体时保留它**（§6 最后一句）——
先看 `extract.py` 输出的字体统计再决定。

⚠️ **Aptos 的 QA 影响**：LibreOffice 没有等宽替代，渲染图上的行宽不等于 PowerPoint 里的行宽。
**对 Aptos 文本，渲染图的"刚好放下/溢出"都不可信。** 预留 ~10% 余量，
溢出以 `layout_check.py` 的测量值为准。中文的 Microsoft YaHei / DengXian 渲染可信。

---

## 3. 字阶（house-style §6）

| 元素 | 字号 | 字重 | 颜色 |
|---|---|---|---|
| 封面标题 | 30–40pt | Semibold | TEXT 或白 |
| 分节页标题 | 28–34pt | Semibold | TEXT 或白 |
| 页面标题 | **26–30pt** | **Semibold（不是 Bold）** | TEXT |
| 正文 | **14–16pt** | Regular | TEXT |
| 次级正文 | 12–14pt | Regular | TEXT_2 |
| 图表标签 | 11–13pt | Regular | TEXT_2 |
| 脚注 / source / 定义 | **9–10pt** | Regular | MUTED |
| 页码 | 9–10pt | Regular | MUTED |

- **正文页标题不超过 30pt**（§6「Avoid oversized titles」）。40pt+ 只给封面和分节页。
- pptxgenjs 没有 semibold 开关：用 `bold: true` 配 Aptos 会偏重，
  **标题可以用 `bold: true`，但正文和小标题不要再加粗**（§6「Avoid excessive bolding」）。
  加粗只给：关键数字、关键词、主结论片段。**不要整段加粗。**
- 句式用 sentence case，**不要全大写**（§6）。

大数字和正文放同一个文本框时行高按最大字号算，**一定溢出**——分成两个 `addText`。

---

## 4. 网格（house-style §7）

```
左右外边距   0.65"        （§7 允许 0.55–0.75，取中值）
上边距       0.50"        （§7 0.4–0.6）
下边距       0.45"        （§7 0.35–0.55）
标题基线     y = 0.50"，h = 0.80"
内容区       y = 1.55"  →  6.55"
脚注行       y = 6.85"，h = 0.40"
页码         右下，y = 7.00"
可用宽度     W = 13.333 - 2 × 0.65 = 12.033"
```

- 一律左对齐（§7）。居中只给封面、分节页、偶尔的单句页
- 内容块间距统一 0.3" 或 0.4"，**选一个用到底**
- 两栏推荐 5.9" + 0.4" + 5.7"（略不对称，比 50/50 自然）
- 留白服务于层级，**不是删信息的理由**（§13）

---

## 5. 图表（house-style §10）

```js
{
  barDir: 'col', barGapWidthPct: 55,
  showLegend: false,                 // §10 能直接标注就去图例
  showValue: true, dataLabelPosition: 'outEnd',
  dataLabelFontSize: 11, dataLabelColor: TEXT_2, dataLabelFontFace: FONT,
  catAxisLabelFontSize: 12, catAxisLabelColor: TEXT_2, catAxisLabelFontFace: FONT,
  valAxisLabelFontSize: 11, valAxisLabelColor: MUTED, valAxisLabelFontFace: FONT,
  valGridLine: { color: GRIDLINE, size: 1 },   // §10 细、淡、不抢视觉
  catGridLine: { style: 'none' },
  valAxisLineShow: false, catAxisLineShow: true, catAxisLineColor: DIVIDER,
  showTitle: false,                  // 标题由页面标题承担
  chartColors: [ACCENT, COMPARE, OTHER]
}
```

- **柱状图从零开始**（§10）。折线图可以收窄区间，但要有分析理由
- **不要每个点都打标签**（§10）——只标有分析价值的
- 小数位按 §16：`58.7%` 不是 `58.65%`
- 禁止：3D、仪表盘、雷达图、装饰性信息图、彩虹配色（§10/§23）
- 桶宽不等的柱状图**必须在脚注说明**，否则视觉上会骗人

**坑**：堆叠柱的 `dataLabelPosition` 只能 `ctr`/`inEnd`/`inBase`，用 `outEnd` **会损坏文件**。
双轴组合图必须同时给 `valAxes` 和 `catAxes` 各两项，否则 PowerPoint 丢弃图表并报损坏。

---

## 6. 表格（house-style §11）

```js
{
  fontFace: FONT, fontSize: 12,
  border: { type: 'solid', pt: 0.75, color: DIVIDER },   // 只要淡分隔线
  autoPage: false
}
```

- **表头不要重填充**（§11「Avoid strong header fills」）。
  用 `bold: true` + `TEXT` 色 + 底部一条 `DIVIDER` 线，不要整行深色底
- 只高亮一行或一列（§11），用 `ACCENT` 加粗；弱化的行用 `MUTED`
- 小数对齐，数字格式全表一致
- 不要每个单元格四边描边、不要隔行变色

---

## 7. KPI（house-style §12）

**不做卡片。** 不要圆角框并排。用排版：

```js
// 指标名：11pt MUTED，在上
// 数值：  20–24pt Semibold ACCENT 或 TEXT，在下     ← 不做成巨大数字（§23/§24）
// 上下文：11pt TEXT_2，再下一行
// 多个 KPI：等间距，需要时插 0.75pt DIVIDER 竖线
```

---

## 8. 形状与效果（house-style §21/§22）

- 只用矩形、直线、圆、箭头
- **直角或极小圆角**，不要夸张圆角卡片
- **默认无阴影**。真要用就极淡、低不透明度、小模糊
- 禁止：发光、浮雕、3D、玻璃质感、重透明效果、渐变

pptxgenjs 注意：`shadow.offset` 为负数**会损坏文件**；要向上投影用 `angle: 270` + 正 offset。
渐变本来就不支持。

---

## 9. 明确禁止（house-style §23/§24）

渐变 · 玻璃质感 · 霓虹色 · emoji · 装饰插画 · AI 生成的通用配图 · 浮动卡片 ·
过度圆角 · 无上下文的巨大数字 · 励志引言 · 编造的引述或客户证言 · 未来感 UI ·
发光元素 · 无谓图标 · 大装饰圆 · 随机抽象形状 · 创业口号 · 彩虹图表 · 剪贴画 ·
三张一样的卡片 · 图标+标题+段落的重复结构 · 通用「Key Insights」页 · 过度对称的版面

§24 最后一句值得单独记：
**「Prefer small imperfections in layout over generic template repetition.」**
版面略有不齐好过模板式的整齐划一。
