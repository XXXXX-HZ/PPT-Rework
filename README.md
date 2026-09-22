# pptx-rework

一个 Claude Skill：**输入一份 .pptx，输出一份能真正拿去讲、并且看不出 AI 痕迹的 .pptx。**

输出始终是**原生可编辑的 PowerPoint 文件**——不是图片，不是 HTML，不是截图拼版。
每一个字、每一根柱子都能在 PowerPoint 里双击改。

---

## 为什么不是又一个"美化 PPT"工具

大部分 PPT 工具在解决"不好看"。这个 skill 认为不好看是次要问题。

它的核心主张是：

> **AI slop 的本质不是"丑"，是"没有身份"。**
>
> AI 默认产出 = 训练语料的视觉与语言最大公约数 = 所有品牌、所有作者、所有领域
> 混在一起的平均值。平均值的问题是它不指向任何人——读者认不出这是谁做的、
> 为谁做的、基于什么做的。
>
> 所以反 slop 的正向定义不是"避开丑东西"，是**把具体性装回去**：
> 具体的数据口径、具体的样本量、具体的时间范围、具体的取舍理由。

实践后果是反直觉的：**去 AI 味最有效的一招不是调配色，是给每个数字补上出处。**
一页写着「转化率提升 30%」而没有口径、没有样本量、没有时间范围，
比任何配色问题都更像机器写的——因为没有人对这个数字负责。

这个论证取自 [alchaincyf/huashu-design](https://github.com/alchaincyf/huashu-design) 的
「反 AI slop」章节。

---

## 特点

**1. 默认是打磨，不是重做。**
假定作者的内容是有意选的。只有出现具体的结构问题（读标题读不出论证线、
缺口径说明、一页塞了两个不相干的发现……）才允许改结构，而且必须先告诉用户理由。

**2. 有一个可执行的规范，不是一堆形容词。**
`references/house-style.md` 定死了色值、字阶、边距、图表参数、表格样式、
禁用清单。所有主观判断都落到具体 token 上。

**3. 三个真的能跑的检查脚本。**

| 脚本 | 作用 |
|---|---|
| `scripts/extract.py` | 导出原稿的文本、表格、图表数值、图片、配色统计 |
| `scripts/smell.py` | 自动扫 AI 痕迹并打分（输入/输出对比，分数必须降） |
| `scripts/layout_check.py` | 出血、溢出、边距违规检查（无 LibreOffice 时的替代方案） |

`smell.py` 检测 20 类痕迹：标题装饰横线、通栏色条、彩色圆底图标、emoji、
通用 SaaS 蓝、投影、正文居中、主题标签当标题、每页恰好三条、空洞词、
**数字无出处**、无局限页、以「谢谢」结尾、纯白底、纯黑字、全大写、标题超尺寸……

它对封面和分节页做分类豁免——那些页本来就该有短标题和大字。

**4. 一个 BLOCKING gate。**
重写后的全部标题（含原文对照）和每页去向必须先发给用户确认，确认前一页都不动。
重构会动结构，方向错了后面全白做。

---

## 安装

```bash
git clone https://github.com/<you>/pptx-rework ~/.claude/skills/pptx-rework
```

Windows：

```powershell
git clone https://github.com/<you>/pptx-rework "$env:USERPROFILE\.claude\skills\pptx-rework"
```

重启 Claude Code 会话后生效（skill 列表在启动时扫描）。

### 依赖

```bash
pip install "markitdown[pptx]" python-pptx
npm install -g pptxgenjs
```

可选但强烈建议（否则无法渲染成图做视觉检查）：

- **LibreOffice** — 提供 `soffice`，转 PDF
- **Poppler** — 提供 `pdftoppm`，PDF 转图

缺这两个时 skill 会自动改用 `layout_check.py` 的几何测量，并在交付时
**明确告知"视觉 QA 未执行"**。

---

## 用法

```
帮我把这个 PPT 改一下 @deck.pptx
```

流程：

```
Step 0  依赖自检
Step 1  提取与诊断（extract.py + smell.py + markitdown）
Step 2  ⛔ 选模式（打磨 / 重建）+ 标题改写对照 → 等用户确认
Step 3  内容编辑（标题、出处、观察/解读/行动分离、数字格式、压缩文字）
Step 4  排版生成（pptxgenjs 原生对象）
Step 5  QA（validate + layout_check + smell 回归 + 渲染逐页看）
Step 6  交付报告（改了什么、哪里做了假设、哪些 QA 没做成）
```

---

## 仓库结构

```
pptx-rework/
├── SKILL.md                      入口：优先级、六步流程、硬规则
├── references/
│   ├── house-style.md            ★ 最高权威：输出规范（色值/字阶/边距/禁用清单）
│   ├── ai-tells.md               22 条 AI 痕迹：现象 → 为什么 → 改法
│   ├── design-system.md          house-style 的 pptxgenjs 实现（token/参数）
│   └── layouts.md                版式目录，带实测坐标
└── scripts/
    ├── extract.py
    ├── smell.py
    └── layout_check.py
```

优先级链（写在 SKILL.md 开头，冲突时按此裁决）：

```
house-style.md > SKILL.md > design-system.md / layouts.md > 内置 pptx skill
```

`house-style.md` 是可替换的。**改成你自己的规范，整个 skill 的行为就跟着变。**

---

## 方法论出处

这是对以下开源 skill 的方法论综合，**不是代码移植**——
它们大多输出 HTML deck，本 skill 输出原生 pptx，只取可迁移的部分。

| 仓库 | 取了什么 |
|---|---|
| [alchaincyf/huashu-design](https://github.com/alchaincyf/huashu-design) | 反 AI slop 的核心论证；禁编造 stat；「一个细节 120% 其余 80%」 |
| [op7418/guizang-ppt-skill](https://github.com/op7418/guizang-ppt-skill) | 锁定色板不许自定义；禁 emoji 作图标；衬线标题配非衬线正文；大字轻字重；图片禁阴影禁边框；不缩字号硬塞 |
| [hugohe3/ppt-master](https://github.com/hugohe3/ppt-master) | 原生可编辑 PPTX 优先；Beautify 作为独立路由；一个 motif 按页面职责变奏 |
| [lewislulu/html-ppt-skill](https://github.com/lewislulu/html-ppt-skill) | 版式目录化——按页面职责选版式，而不是全篇套一个 |
| [Gabberflast/academic-pptx-skill](https://github.com/Gabberflast/academic-pptx-skill) | action title；ghost deck 测试；一页一个 exhibit |

技术实现依赖 Anthropic 内置的 `pptx` skill（pptxgenjs API、OOXML 编辑、文件校验）。

---

## 已知限制

- **字体 QA**：`house-style.md` 默认指定 Aptos。LibreOffice 没有 Aptos 的等宽替代，
  渲染图上的行宽和 PowerPoint 里不一致，所以对 Aptos 文本**渲染图的溢出判断不可信**，
  要以 `layout_check.py` 的测量值为准。中文的 Microsoft YaHei / DengXian 渲染可信。
- **`smell.py` 测不出论证。**「有没有一条论证线」「标题能不能串起来」
  必须人工过 `ai-tells.md` 的 C 组。分数低不等于能讲。
- **图表数值需要人读。** 原稿如果是位图图表，数值要从图里读出来重建。
  读不准的地方 skill 会停下来问你要原始数据，**不会猜**。

---

## License

MIT
