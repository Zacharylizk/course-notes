---
name: course-notes
description: "把一节课的 PPTX/PDF 变成单页离线双语 HTML 学习页（默认交付单文件、图片 base64 内嵌）：总述与公式清单、按理解重讲、抽图/重绘、分步推导、折叠例题、自测与易错清单、左侧导航。"
whenToUse: "用户发来课件（.pptx/.pdf）并要求学习、讲解、整理成学习页/笔记，或明确说用 course-notes 技能时。"
---

# 课件 → 学习页 生成器（Course Notes Builder）

> 本技能基目录：`~/.dsh/skills/course-notes`（Windows 下是 `%USERPROFILE%\.dsh\skills\course-notes`）。
> 下面所有相对路径都相对于这个基目录；先 `cd` 到基目录，或把 `tools\` 换成基目录下的绝对路径。

> 这是一个「可重复的工作流 + 模板 + 工具」，不是一次性提示词。
> 每次用户发来一节课的 PPT/PDF，按本文件执行，产出**一个**自包含 HTML 学习页。

---

## 0. 先判断：这是「学习页」不是「复习页」

| | 复习页（旧） | 学习页（本流程） |
|---|---|---|
| 目的 | 考前 2 小时串完 | 课后第一次真正学懂 |
| 结构 | 按题型/公式归类 | 按认知顺序：问题 → 直觉 → 定义 → 推导 → 例题 |
| 深度 | 结论 + 速查 | 每一步「为什么」，符号逐个解释 |
| 例题 | 展示解法 | 先给思考入口，再折叠解答 |
| 语气 | 提纲 | 讲课，像坐在旁边讲 |

**硬规则：不要把 PPT 的句子抄成要点。** 先读懂，再用自己的话重讲；PPT 只提供知识点清单、图、原题和数据。

---

## 1. 输入与工作目录

```
course-notes/
  shell.html                 # 页面外壳（CSS/JS/导航/打印），不要改内容结构
  tools/
    extract_deck.py          # pptx/pdf → text.md + slides.json + pages/*.png + images/*
    crop.py                  # 从整页渲染里裁图
    tile.py                  # 4 页一张 contact sheet，用于快速扫全 deck
    build.py                 # body.html + shell → index.html（自动生成侧栏）
    inline.py                # index.html + img/ → index-single.html（图 base64 内嵌）
  lectures/
    <course>/<lecture-slug>/
      body.html              # 我写的内容片段
      index.html             # 构建产物（分文件版，改内容时用）
      index-single.html      # 交付产物：图全部内嵌的单文件
      img/                   # 裁出来的图 / 重绘的 SVG
      work/                  # 抽取产物（text.md, pages/, slides.json）
```

输出文件命名：`<两位课次>-<英文短标题>.html`，例如 `05-plane-stress.html`。

**默认交付「单文件版」**：`index-single.html` 重命名后拷到用户指定目录（默认课件同目录）。
原因：分文件版的图靠 `img/…` 相对路径，用户把 html 单独挪走（或只转发 html）图就全空白——这是本技能历史上最高频的问题，别再重复。
`img/` 与分文件版留在 `lectures/` 里供后续修改，**不往交付目录塞**（除非用户明确要「方便改」的版本，那也要说清楚 html 和 `img\` 必须一起动）。
页面必须离线可开（无 CDN、无 MathJax）；单文件版连图片都不依赖磁盘上的其它文件。

---

## 2. 流水线（每次照做）

### 2.1 抽取

```powershell
# 任意带 python-pptx / Pdfium 的解释器；需要固定路径就写绝对路径
$py = "python"
& $py tools\extract_deck.py "<deck.pptx|pdf>" --out "lectures\<course>\<slug>\work"
```

- `.pptx`：python-pptx 取文字/表格/备注/内嵌图，PowerPoint COM 导出整页 PNG。
- `.pdf`：pdfium 取文字层并渲染每页；**没有文字层的 PDF 很常见**（扫描/导出图），此时只能读 `pages/*.png`。
- 先看 `work/text.md` 了解全课；再 `tile.py` 出 2×2 联系表快速扫版式；关键页单独 `read_image` 看细节。

### 2.2 读图（不可跳过）

必须逐页确认三类东西：

1. **图**：受力图、几何图、曲线、莫尔圆、电路、流程图 → 决定用原图还是重绘。
2. **非文字强调**：箭头、红框、高亮、动画分步（常以多张近似页出现）、手写批注。
3. **原题**：PPT 里的例题/数据/答案，优先复用并独立验算。

> 课件里「一张图 + 三个箭头」往往就是老师真正讲的内容。只读文字层会漏掉一半。

### 2.3 列大纲（先写给自己看）

在 `body.html` 里先放占位标题，确认：

- 这节课的**一个核心问题**是什么？
- 知识点之间的**依赖顺序**（哪个必须先讲）？
- 哪 2–3 个点必考？哪 1–2 个点最容易错？
- 需要补哪些先修（数学/物理），补到什么程度？

再按第 3 节的结构填内容。

### 2.4 写内容 → 构建

```powershell
& $py tools\build.py "lectures\<course>\<slug>\body.html" "lectures\<course>\<slug>\index.html" `
  --course "5104MECH" --lecture "Lecture 05 · Plane Stress"
```

`build.py` 会：

- 从第一个 `<h1>` 取 `<title>`；
- 从每个 `<h2>`（或它所在 `<section id>`）生成左侧导航；
- 检查重复 id、失效锚点、缺失图片；
- 填入 `{{TITLE}}` / `{{NAV}}` / `{{COURSE}}` / `{{LECTURE}}` / `{{BODY}}`。

### 2.5 打包成单文件（默认交付形态）

```powershell
& $py tools\inline.py "lectures\<course>\<slug>\index.html" "lectures\<course>\<slug>\index-single.html"
```

- 把每个本地 `<img src="img/…">` 换成 `data:` URI，其余一律不动；结束会打印内嵌数量和体积（常见 1–3 MB）。
- 交付的是这个文件，重命名成 `<两位课次>-<短标题>.html` 再拷出去。
- 分文件版（`index.html` + `img/`）只在需要我继续改内容时用；如果用户也想要，明确提醒「html 必须和 `img\` 一起移动」。

### 2.6 验收

- Edge 无头截图（1440 宽 + 500 宽）看排版、折行、溢出、折叠态。
- 打印预览：`--print-to-pdf`，确认 A4 下不截断、导航隐藏。
- 每道例题独立复算一遍；答案写进 `.answer`，并用不变量/极限/量纲自查。
- 所有图片路径存在；页面断网可开。
- **单文件版拷到桌面（离开 `img/`）双击打开，图必须全在**——这是唯一要保证的形态。

---

## 3. 页面结构（学习页的标准骨架）

按顺序，不要求每节等长，但**不要缺**：

### 3.0 Hero（首屏）

- 课程代码 + 课次 + 标题（英文主标题 + 中文副标题）。
- 3 句话：这节课回答什么问题、为什么难、学完能做什么。
- 一颗「核心公式」或「核心图」。
- 4 个 chips：核心概念 / 工具 / 考试重点 / 预计阅读时间。

### 3.1 总览：先看终点（Learning outcomes）

紧接 hero 的第一个 card 必须是「总述」，用户明确要求先知道这节课讲了啥、对应哪些公式：

- **一句话**：把整节课压缩成一句人话。
- **路线图**（`.roadmap`）：3–5 个概念节点，箭头连接，显示依赖顺序。
- **公式清单**（`.fstrip`）：3–6 个公式卡，每张 = 英文名 + 中文名 + 公式。只放本课最核心的，不要把所有变形都列上。
- **学习目标**：3–4 条，用「能……」「能解释……」「能判断……」写，不要写「了解……」。
- **先修**：只补真正卡住的那一个点。

> 这个 card 是「课前 30 秒」：看不懂后面时，回到这里看路线图和公式清单。

### 3.2 导入：为什么需要它

一个具体场景/失败案例/反直觉问题。**先不给定义。**
例如：薄板为什么能忽略厚度方向应力？先想自由表面。

### 3.3 知识点（按依赖顺序，可偏离 PPT 顺序）

每个知识点固定四件套：

1. **一句话**（`.sec-lede`）：这个点到底在说什么。
2. **直觉**：用日常语言/几何图像解释，先文字后符号。
3. **定义/公式**（`.def` + `.eq.key`）：英文定义 + 中文解释。
4. **推导或用法**（`ol.steps`）：每步配 `.why`，写「为什么可以这样做」。

再按需加：

- 图（`figure` + `figcaption` + `.fig-note` 看图要点）。
- 对比表（易混概念）。
- 一个 30 秒小检查（`.quiz`）或口头自问。
- `.en-line`：考试会怎么用英文问。

### 3.4 例题（每课 3–6 道）

- 1 道**完整带路**（新概念刚学完，步骤写全）。
- 1–2 道**考试级**（只给思路入口，解答折叠）。
- 1 道**易错型**（故意设置符号/单位/边界条件陷阱）。
- 英文原题（保留原表述）+ 中文翻译；解答用中文，关键步骤给英文术语。
- 难度星级、出处（slide/page）标在 header。
- `details.sol` 默认折叠；答案放 `.answer`；最后写「考点」和「自查」。

### 3.5 易错点

只写**这节课真的会错**的，3–8 条。每条：错法 → 为什么错 → 正确做法。
不要放通用鸡汤（「认真审题」）。

### 3.6 一页纸带走

- `.takeaway`：一句话总结。
- `.fgrid`：4–8 张公式卡（名称 + 公式 + 什么时候用）。
- 术语表：English / 中文 / 一句话。
- `.checklist`：3–5 条「我是否能……」，可勾选（localStorage 记忆）。
- 3–5 道自测（`.quiz`，点选即时反馈）。

### 3.7 课件对照（可选但推荐）

表格：本页小节 ↔ 课件页码。方便用户上课后对照，也方便我下次增量修改。

---

## 4. 双语规则（英语课件强制）

用户是中外合办，**考试是英文题，理解靠中文**。不要全文机翻，按信息层级分配：

| 位置 | 语言 |
|---|---|
| 标题 / 小节标题 | 英文 + 中文（`h2` 内 `.en` 副标题） |
| 术语首次出现 | `term` 组件：**English** + 中文 |
| 定义 | 英文定义（完整、可背）+ 中文解释（口语化） |
| 公式名 | 英文 + 中文（如 `stress transformation 应力变换`） |
| 例题题干 | 英文原文 + 中文翻译 |
| 解答步骤 | 中文为主，关键动词/术语保留英文 |
| 考试问法 | `.en-line`：英文原句 + 中文题意 |
| 正文讲解 | 中文（必要时夹英文术语） |

写英文时用**考试会出现的句式**，不要用文学化表达：
`Determine …`, `Show that …`, `Calculate …`, `State the assumption …`, `Sketch …`.

中文解释要「说人话」，允许用「你可以把它想成……」；英文定义要严谨，考试可直接背。

---

## 5. 图与公式

### 图

- 优先用**原始清晰图**：从 PPT 内嵌图提取，或从整页 PNG 裁切（`crop.py`）。
- 整页截图不算图：必须裁到只含目标图，去掉页眉页脚。
- 图里的英文标签保留；在图注里给中文解释。
- 示意图/受力图如果原图糊、或需要分步高亮，用**手写 SVG** 重绘（参考 `sample-plane-stress/img/plane-stress.svg`）。
- 每张图必须有存在理由：要么解释机制，要么给数据，要么是题图。装饰图删掉。
- 图注格式：`图 N　中文说明（English term）`；下面可加 `.fig-note`「看图三步」。

### 公式

- 不联网：用 Unicode + `.frac` / `.sqrt` / `.abs` 等 CSS 组件，不用 MathJax/KaTeX。
- 每个符号第一次出现必须定义（符号表或行内）。
- 关键公式用 `.eq.key`；最终答案用 `.eq.ans` 或 `.answer`。
- 推导不要跳步：每一步写「用了什么」+「为什么成立」。
- 数值例题统一 SI，写单位；最后做量纲/极限/不变量检查。

---

## 6. 写作语气

- 像助教坐在旁边讲，不要像教科书。
- 允许「先别急着记公式」「这里有个坑」这类口语。
- 每 3–4 段必须有一个视觉断点（图、表、公式块、callout、例题）。
- 不写「本页介绍了……」这种元话语；直接讲内容。
- 不堆砌「显然」「易得」；如果显然，说明为什么显然。
- 中文标点用全角；英文术语前后加空格。

---

## 7. 组件速查（shell.html 已内置）

```html
<section class="card" id="xxx">
  <h2><span class="n">01</span> 中文标题 <span class="en">English Title</span></h2>
  <p class="sec-lede">一句话。</p>
  ...
</section>
```

- 术语：`<span class="term"><b>plane stress</b><i>平面应力</i></span>`
- 英文定义：`<div class="def"><div class="t">Definition · 定义</div><div class="en">…</div><div class="zh">…</div></div>`
- 步骤：`<ol class="steps"><li><p>…</p><div class="why"><b>为什么</b>…</div></li></ol>`
- 提示：`<div class="note tip|warn|info|def|amber"><b>标题</b> …</div>`
- 例题：`<article class="example"><header>…</header><div class="q">…</div><details class="sol"><summary>…</summary><div class="sol-body">…</div></details><footer>…</footer></article>`
- 公式卡：`<div class="fgrid"><div class="fcard"><div class="fk">…</div><div class="ft">…</div><div class="eq">…</div></div></div>`
- 自测：`<div class="quiz" data-exp="解析" data-hint="提示"><p class="quiz-q">…</p><div class="opts"><button class="opt" data-ok="1">A. …</button>…</div><div class="quiz-fb"></div></div>`
- 勾选：`<div class="checklist"><label><input type="checkbox"><span>…</span></label></div>`

不要改 `shell.html` 里的 JS 行为；新增样式优先复用变量。

---

## 8. 质量检查清单（交付前逐项打勾）

- [ ] 开头 30 秒能看懂「这节课讲什么、要会哪几个公式」。
- [ ] 每个知识点都有「为什么」，不是只有结论。
- [ ] 每个符号都定义过；每个公式都说了适用条件。
- [ ] 例题数据独立复算，答案与题干自洽。
- [ ] 英语题干保留原文；关键术语双语。
- [ ] 图都裁干净、有中文图注、有「看图要点」。
- [ ] 易错点是本课特有的，不是通用套话。
- [ ] 自测题能区分「记住」和「理解」。
- [ ] 侧栏锚点全通；无重复 id；无缺图。
- [ ] 1440px 与 500px 宽无横向溢出；A4 打印不截断。
- [ ] 断网打开正常（无外部字体/脚本/图片）。
- [ ] 交付的是**单文件版**（图 base64 内嵌）：单独把 html 拷到别处，图仍然齐全。

---

## 9. 常见坑

1. **PPT 文字层是空的**：扫描版 PDF 很常见，直接读 `pages/*.png`，别硬 OCR 等工具。
2. **公式是图片**：文字层会丢；以渲染页为准，必要时重打公式。
3. **PPT 分步动画**：同一页导出多张时，内容会重复/缺步；以最终页 + 备注为准。
4. **符号约定不同**：剪应力正负、θ 转向，先声明再套公式。
5. **单位**：rpm→rad/s、mm→m、MPa→Pa，例题里至少埋一次。
6. **不要照搬 PPT 顺序**：如果 PPT 先给公式再解释，学习页应反过来。
7. **页面过长**：超过 ~12 节时拆成「上/下」两页，或在侧栏分组；不要把所有东西塞一页。
8. **旧复习页的写法**：不要把「五类振动」那种速查表直接搬来当学习页；学习页要讲清第一性原理。
9. **只交付 html + img/**：用户挪走 html、图全丢（已经发生过一次）。交付一律走 2.5 的单文件版；分文件版只留在 `lectures/` 里当工作副本。

---

## 10. 已知环境

- Python 3 + `python-pptx`、`Pillow`；走 PDF 路线再加 `pypdfium2`（`pip install python-pptx pillow pypdfium2`）。
- 如果机器上有两个解释器（一个装了 python-pptx，另一个装了 pdfium），用环境变量 `CN_PY` / `CN_PY_ALT` 指定，`extract_deck.py` 缺库时会自动切过去重跑。
- 渲染 PPTX 用 PowerPoint COM（VBScript 晚绑定），需要本机 Office；窗口会闪一下。
- 无网络依赖；不要引入 CDN。
