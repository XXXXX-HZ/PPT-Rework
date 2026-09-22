# 打磨模式 — 模式 A 的具体手法

模式 A（打磨）不是"把字号调统一就完事"。这份文件给的是**在不动页序、不动分析结论的前提下，
把一页报告变成一页可讲的幻灯**的具体手法。

来源：部分模式吸收自 `analyst-deck-polisher`（一份基于同一套 house-style 的独立实现），
在本 skill 的实测中验证并补上了可测量的判据。

---

## 0. 单页自检：一眼落点

> **观众第一眼落在哪里？是不是标题所声称的那个证据？**

这是整份文件里最有用的一句话。

标题说「Google Ads 付费率从 0.91% 涨到 3.43%」，那么这两个数字必须是页面上最先被看到的东西。
如果观众第一眼落在表头、落在别的行、或者根本没有落点，这一页就没做完——**不管排版多干净。**

这一条排在所有视觉规则之前。

---

## 1. 报告页 → 幻灯页

**原始形态**（最常见的待打磨形态）：

```
标题
一张密表
一大段解释文字
脚注
```

**不要直接推倒重画。** 先建立视觉焦点，顺序如下：

1. **找出证明标题的那一个值 / 行 / 列**
2. **把它高亮**（accent + 加粗），同时**把次要数据压下去**（`TEXT_2` 或 `MUTED`）
   —— 压暗次要项比再加一个颜色有效得多
3. **把解释段落压成 2–4 条可扫读的陈述**（见 §3）
4. **口径、定义、caveat 移到脚注**，不要占正文
5. 合适时**把一两个数字提成无框 callout**（§12：排版，不做卡片）

第 2 步是关键：**大多数人只做了高亮，没做压暗。** 只高亮不压暗，页面仍然是平的。

---

## 2. 平表 → 证据层级

标题如果点名了某个渠道 / cohort / 实验组 / 分段：

- 高亮**那一行或那一列**
- **before / after 两个数字都要加粗**——只标 after，观众看不到变化幅度
- 其余行保持中性，不要每个类别配一个颜色
- 表头不用深色底（house-style §11），用加粗 + 一条底部细线

---

## 3. 长段落 → 短陈述

**只在确实更好扫读时才转。** house-style §14：一段话更清楚时就不要用 bullet。

值得转的典型信号：**这段话在罗列同一个维度的多个分位 / 分组 / 时间点**。

实例（本 skill 实测中被改的一段）：

```
改前（一段）：
The median last purchase falls on Day 4 in both cohorts, but the distribution is
heavily skewed rather than merely short. A quarter of early payers make their last
purchase on Day 1. Another quarter keep purchasing to around D29–D33, and the top
10% are still buying after D75–D78. The mean of 21 days describes almost nobody.

改后（label / value 陈述栈）：
Median last purchase   D4 in both cohorts
Bottom quartile        25% stop on D1 and never return
Top quartile           Still purchasing at D29–D33; top 10% continue past D75
Mean of 21 days        Not representative — the distribution is skewed, not short
```

左列 13pt `TEXT_2` 标签，右列 14pt `TEXT` 内容，行间一条 `DIVIDER` 细线。
**不做卡片、不加 bullet 点**（house-style §12/§20）。

⚠️ **caveat、口径、样本限制、右截断说明不能在这一步被压掉。** 压的是啰嗦，不是严谨。

---

## 4. 版式节奏 —— 可测量的

**反模式**：整份 deck 大多数页都是 `标题 → 表 → 段落 → 脚注`。

这是打磨类任务最容易掉进去的坑，因为逐页看每一页都合格，**只有把版式序列排出来才看得见**。

检测方法（`scripts/smell.py` 的 `C3` 检查会自动跑）：
给每页生成一个版式签名，统计重复。

```
判据：
  同一签名出现 > 40% 的页数        → 单调
  连续 3 页以上同签名              → 单调
  每一页都以长段落收尾              → 单调
```

实测例（本 skill 打磨一份 11 页分析稿）：

```
改前：5× TITLE+TABLE+PARA，11/11 页都以长段落收尾   ← 单调
改后：3× TITLE+TABLE，最长连续 2 页                 ← 通过
```

**怎么破**：不是为了变化而变化（house-style §25「不要无理由重构」），
而是在内容支持的地方换手法——某几页的段落改成陈述栈（§3），
某一页把关键数字提成无框 callout（§1 第 5 步），双图页做并排对比。

可用的交替形态：

- 单张主导图
- 表 + 高亮证据
- 图 + 短评（图占左 2/3，解读在右）
- 双图对比
- 3–5 条发现的小结页
- 稀疏分节页
- 无卡片的指标 callout

---

## 5. 分节页

保持稀疏（house-style §33）：大号弱化的节号 + 短标题 + 可选一句范围说明。
**不加插画、图标、抽象装饰。**

分节页很适合承载方法论交代——比如「本节只覆盖早期付费者，因为晚付费者会被 90 天窗口右截断」。
这种话放在分节页比单开一页自然。

---

## 6. 小结页

3–5 条**真实发现**，不是「Key Insights」这种分类名（house-style §32）。
编号陈述或简单竖排堆叠，**不要每条塞进一个圆角卡片**。

标题写成 `Three findings explain the retention gap` 这种带信息量的句子。

---

## 7. 双图对比

- **两张图的坐标轴范围、标签、视觉语法必须一致**，否则对比不成立
- **如果论点是"两个分布反转"，就必须把反转标出来**——靠排序、标注、或统一坐标轴。
  不标注的话，观众要在脑子里做对齐，说服力掉一半
- 每张图下面一行 12pt 说明它是什么，整体结论放在两图下方

> 判断题：如果两张并排图其实是**同一个比较的两个侧面**，考虑合成一张双系列图。
> 但这属于结构改动——**模式 A 下要先问用户**，不要自作主张。

---

## 8. 标注

用于：发版/变更日期、异常点、实验边界、异常峰值、重要 caveat。

小字 + 细引导线。**不用气泡框、不用漫画式 callout**（house-style §15）。

斜体 + accent 色的一行小字通常就够了，例如在双图页上写 `tallest bar` 和
`…becomes the second shortest`。

---

## 9. 观察 / 解读 / 行动

一页里混了这三者时必须区分开（house-style §29）：

```
Observation      Skip players retain 14pp worse on D1.
Interpretation   Skipping may disproportionately affect users who still need onboarding.
Action           Test a shorter tutorial path rather than a full skip.
```

**绝不把解读或建议写成观察到的事实。** 措辞上 Interpretation 要留余地（may / suggests），
Observation 不留余地。

---

## 10. 最终判据

> **成品应该像"一个认真的分析师打磨了自己的稿子"，
> 而不是"一个模板被套在了这份分析上"。**

套模板的迹象：每页一样齐、每页一样满、高亮位置和标题无关、
所有段落一样长、看不出哪一页是重点。
