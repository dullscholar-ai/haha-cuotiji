---
name: haha-cuotiji
description: 「哈哈错题集」——把试卷/作业截图与题目数据排成一题一页的错题集 docx。每题一张对应卷面截图，题干·小问·选项·检索表·解析·答案分区排布，并按实测版面反解每张图能有多大，做到「题数 == 页数」且截图完整可读。当用户说「做成错题集」「错题本排版」「每题一页」「把截图和题目对上」「重新排版这份习题/试卷」「康奈尔错题笔记」，或拿到一份带图片的习题 docx、要求图题一一对应并重新排版时使用。
whenToUse: 有习题/试卷素材（截图或 docx）需要排成可打印的错题集；或已有错题集版式错乱（图片全堆在开头、题目重复、跨页混乱、截图太小或显示不全），要重排成一题一页的干净稿。
---

# 哈哈错题集 · 一题一页排版

一句话：**把「题目数据 + 卷面截图」写成 JSON，渲染成 docx，再用排版引擎实测反解每张图能有多大，直到「题数 == 页数」且每题截图完整、可读。**

产物：A4 docx（可导出 PDF）。**一题一页，只留「题头 / ① 原题 / ② 解题过程（右侧并排卷面原件）/ ③ 答案」。**

---

## 0. 四件事先定死（除非用户逐条推翻）

| 问题 | 默认答案 | 依据 |
|---|---|---|
| 留哪些区块？ | **只留** 题头、① 原题、② 解题过程、③ 答案 | 用户连删四轮：封面、错题纲要、编排说明、题目索引、线索栏、总结栏、错因、错因归纳、核心问题、学习建议、订正区 —— 全部不要 |
| 一题一页？ | **是**，但用 `1×1 表格单元格 + cantSplit` 实现，不要用分页符 | 分页符会留半页空白；cantSplit 让整题「整体搬家」，短题还能共处一页 |
| 截图放哪？ | **与 ② 解题过程并排**（左 5.6cm 文字 / 右 11.2cm 图） | 通栏图会让整题超一页；并排后图能到 11×16cm（A4 原卷的 54%） |
| 图注怎么写？ | 用原稿自己的说法：`第 7 张 · 上半 / 中部 / 中下部 / 下半` | 写「(1/4)(2/4)」会被理解成「只给了四分之一」 |

**硬性不做**：封面页、目录/索引、编排说明、线索栏、总结栏、错因、订正区、附录。文档第一页就是第一道题（文档名放页眉）。

---

## 1. 最终版式速查（实测值，A4）

| 项 | 值 |
|---|---|
| 页面 | 21 × 29.7 cm，上 1.8 / 下 1.5 / 左右 1.8，可用高度 **26.4cm** |
| 正文 | 微软雅黑 9.5pt，行距 1.2 |
| 页眉 | 8.5pt 灰：`初中科学典型习题 · 详细解析`（左）＋副题（右，右制表位 17.4cm），下加细线 |
| 页脚 | `— 页码 —`（`w:fldSimple` PAGE 域，右制表位 17.4cm） |
| 题头 | 题号 13pt 粗 ＋ `｜分类 · 考点` 9.5pt 灰，下加 0.75pt 粗线 |
| 小节标题 | `① 原题` / `② 解题过程` / `③ 答案` 10.5pt 粗，下加 0.5pt 浅灰线 |
| ③ 答案 | 整段浅灰底纹（`w:shd F4F5F7`） |
| 并列栏 | 左 5.6cm（解析）／右 11.2cm（图） |
| 图 | 宽 ≤ 11.2cm，高按题反解，竖长图上限 16cm；`width` 与 `height` 两个参数都给，保证不变形 |

详见 `references/final-layout.md`（几何、反解公式、四轮改动记录）。

---

## 2. 最短路径（校准闭环，2 轮收敛）

```bash
# ① 数据 → HTML → docx（cal 模式：每题强制独立一页、图高固定 6cm，只为测量）
python assets/production/postprocess.py                      # 生成精简 HTML/MD
python assets/production/build_docx_local.py in.html _cal.docx --cal

# ② 实测每题「文字高度」，反解出每题图高上限
pwsh -File assets/production/calibrate.ps1 -Docx _cal.docx -Figcaps figcaps.json

# ③ 正式排
python assets/production/build_docx_local.py in.html 错题集.docx

# ④ 验收（页数 / 每页首行 / 图尺寸）
pwsh -File assets/production/calibrate.ps1 -Docx 错题集.docx -Audit
```

反解公式（cal 模式实测得到每题文字高 `text_cm`，含题头/①/②/③）：

```
cap_cm = clamp(6.0, 16.0, 26.4 - text_cm - 0.85)
```

`0.85` = 图注 + 段落间距 + 安全余量。结果写 `figcaps.json`，键为题序号。

**验收判据**：`pages == 题数`；每页首行是题头；没有题被拆到两页；PDF 页数与 docx 一致。

### 两条渲染路线

| 路线 | 入口 | 适用 |
|---|---|---|
| **A. 直出**（推荐新项目） | `assets/build_docx.py --content content.json --figs figs --out x.docx` | 只有 JSON + 截图时；要收敛需配合 `measure.ps1` + `fit.py` |
| **B. 生产**（本次成品用的） | `assets/production/*` | 已有一份康奈尔 HTML 要精简重排；图高用 `figcaps.json` 逐题反解 |

路线 A 的默认值已按最终版式调过（微软雅黑、并列栏 5.6/11.2、默认全部并列），但**单页预算仍靠 `fit.py` 收敛**：没有 `fit.json` 时示例可能溢出一页，属正常，跑一遍反解即可。

---

## 3. 四条硬约束（都踩过）

1. **`w:pPr` 子元素必须按 schema 顺序插入。** 直接把 `w:keepNext` append 到 `pPr` 末尾，Word 会**静默忽略**（表现为题头孤零零掉在页尾）。要么用 python-docx 的 `p.paragraph_format.keep_with_next = True`，要么按 `pStyle→keepNext→keepLines→…→pBdr→shd→…→spacing→ind→jc` 的顺序 `insert_element_before()`。
2. **keepNext 链挡不住「整题超过一页」。** 链太长时 Word 照样断，还把题头留在页尾当孤行。可靠做法：整题塞进 `1 行 1 列` 表格并 `row.cantSplit`，内部段落不设 keepNext。
3. **图放表格里必须比单元格窄。** 图宽 > 单元格宽会被**裁掉**（不是缩放）；`width` 和 `height` 两个参数都要给，只给 width 时按 DPI 猜，可能变形。
4. **docx 被 WPS 打开时写入会 `PermissionError`。** 先写临时名再 `Copy-Item` 覆盖；目录里有 `~$*.docx` 就说明还开着。

---

## 4. 整页长截图（最容易做成「显示不全」）

手机截长题会得到 **989 × 3500** 这种「一页被截成 3 段拼起来」的图。**只贴第一段 = 丢掉 2/3 的题面**，而图注还写着「1/3」——用户一眼就看出来。

判定：`h / w > 2.2`。

处理：
1. 用**最空白行**切片（在 40%~62% 高度范围内找墨迹最少的行），不要等分；
2. 两栏放在**单独一页图版页**：`page_break_before` + 图注 + 2 列无框表格，每栏 8.4cm 宽；
3. 文字页末尾加一行小灰字：`▸ 卷面原图见下页：整页原图已完整收录（左右两栏并排，左栏读完接右栏）`；
4. 图注写 `整页原图`，不要写 `(1/3)`。

为什么两栏：整页等比缩到 26.4cm 高时只有 7.4cm 宽（正文约 4pt，看不清）；切两栏后每栏 8.4cm，完整且可读。

**验证切片没切到字**：切完检查首/末 6 行墨迹占比，必须接近 0%（切在空白行上）。

---

## 5. 内容 JSON 速览

```json
{
  "subject": "科学", "code": "科学-20261031-001",
  "questions": [{
    "q": "7", "sec": "生物 · 分类检索", "topic": "二歧检索表 · 脊椎与无脊椎",
    "img": "figures/orig-01.jpg", "imgcap": "卷面原件 第 1 张 · 第 7 题",
    "imgpage": {"a": "figures/orig-06-a.jpg", "b": "figures/orig-06-b.jpg",
                "cap": "卷面原件 第 6 张 · 第 10 题（整页原图，左栏读完接右栏）"},
    "body": [["① 原题", ["题干…", "(1) …"]],
             ["② 解题过程", ["▸ 检索：…", "▸ 归类：…"]],
             ["③ 答案", ["**答案**：(1) E　(2) ②③⑤"]]]
  }]
}
```

- `img`：并列在 ② 右侧的单张图；`imgpage`：整页长截图（走图版页，可覆盖 `img`）
- 富文本：`**粗体**`、`@@等宽@@`；`"————"` → 分隔线；`[[TABLE:KEY]]` → 数据表
- 旧字段 `star / src / cue / summ / outline / cause_rows / core_issues / advice / correction_box` 可留在 JSON 当档案，**默认不渲染**（见 `assets/production/postprocess.py`）

---

## 6. 从一份"错版"docx 起步的顺序

1. dump 原文件所有段落/表格块，把图片、重复段、坏公式找出来；
2. 建「题 → 图」映射表（**全流程最关键的产物**，原文图片通常与题目没有绑定）；对不上时用原稿的「第 N 张」说明 + 切片墨迹做交叉验证；
3. 去重 → 断句（小问/选项）→ 表格化（检索表、数据表）→ 修公式 → 提炼答案；
4. 按 §2 校准闭环排、测、收敛。

踩坑表见 `references/workflow.md`，几何与四轮改动见 `references/final-layout.md`。

---

## 7. 目录

```
haha-cuotiji/
├── SKILL.md
├── assets/
│   ├── build_docx.py            ← 通用 JSON→docx 引擎（--side / --fit）
│   ├── prep_figs.py             ← 截图准备：切长图 / 拼图 / 去白边 / 量尺寸
│   ├── measure.ps1              ← 旧版结构（3 行表）实测，留作参考
│   ├── fit.py                   ← 反解 fit.json
│   └── production/              ← 产出「最终版式」的那一套（改路径即可复用）
│       ├── postprocess.py       ← 精简版：删区块、图并到 ② 右侧、长图转图版页
│       ├── build_docx_local.py  ← HTML → docx（表格包题、图高读 figcaps.json）
│       ├── calibrate.ps1        ← cal 实测 + 反解 figcaps.json + 验收
│       └── README.md
├── references/
│   ├── final-layout.md          ← 最终版式几何 + 反解 + 四轮改动记录（先看这个）
│   ├── layout-spec.md           ← 通用版式规格（旧，作为备选样式）
│   └── workflow.md              ← 全流程 + 踩坑表 + 验收清单
└── examples/
    ├── sample-content.json
    └── figs/
```
