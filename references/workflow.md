# 工作流与踩坑（workflow）

> ⚠️ 先看 `final-layout.md`：默认版式是「题头 + ① 原题 + ② 解题过程(右并截图) + ③ 答案」，
> 不要封面/索引/错因/订正区，一题一页用 `1×1 表格 + cantSplit` 实现。

## 全流程

```
素材 ──①清点──> 映射表 ──②截图准备──> figs/ ──③建内容──> content.json
                                                            │
                                          ┌─────────────────┘
                                          ▼
                                    ④首排 build_docx.py ─> A.docx
                                          │
                            ⑤实测 measure.ps1 ─> 谁跨页？谁的字号小？
                                          │
                                     fit.py ─> fit.json
                                          │
                              build_docx.py --fit fit.json ─┐
                                          ▲                  │
                                          └── 复测（2~3 轮）──┘
                                          ▼
                                    ⑥定稿 + PDF + 校对
```

### ① 素材清点（最容易被跳过、最容易返工）

打开原始 docx，把每个段落/表格块 dump 出来（用 python-docx 遍历 `body` 的子元素），**显式建一张"题 → 图"映射表**：

```
题号 | 题目标题 | 卷面来源(如"第 7 题") | 截图文件名 | 图注
```

原始素材里的图片经常是**一整段里堆着 9 张图**、而正文只有一句"如第 7 张图所示"。这种结构里图片与题目**没有任何内在绑定**，只能靠"第几张图对应第几题"的人工对应关系。**先写死这张表，再动手排版**；否则后面每题都要重新猜一次。

### ② 截图准备（`assets/prep_figs.py`）

| 情况 | 命令 |
|---|---|
| 一张长图叠了 5 道题 | `prep_figs.py auto --src scan.jpg --outdir figs --prefix fig --gap 40` |
| 切点要手调 | `prep_figs.py split --src scan.jpg --cuts 1200,2400 --outdir figs` |
| 竖长条太窄、缩完看不清 | `prep_figs.py join --a l.jpg --b r.jpg --out fig06.jpg`（左右分栏并排，同宽下清晰度翻倍） |
| 四周白边太多 | `prep_figs.py trim --src x.jpg --out y.jpg --pad 8` |
| 想知道该走通栏还是侧栏 | `prep_figs.py info --glob "figs/*.jpg"` |

`info` 会直接给出"通栏高 / 侧栏高"两个值，比例 ≥ 0.9 的优先考虑并列版式。

### ③ 内容建模（content.json）

一个题目一个对象，字段：

| 字段 | 必填 | 说明 |
|---|---|---|
| `no` | ✓ | 两位序号 `"01"`，同时是 fit.json 的键 |
| `src` | ✓ | 卷面出处，如 `"第 7 题"`、`"典例 2"` |
| `title` | ✓ | 题目短名，进题头 |
| `kind` | ✓ | `选择题 / 填空 / 实验题 / 探究填空` |
| `note` | | 题头下方的补充说明（如"本卷有两道第 12 题"） |
| `q` | ✓ | 题目块数组，元素是 `["p"/"sub"/"opt"/"deep"/"tbl", "文字"]` |
| `qtable` | | 检索表/数据表规格，`q` 里用 `["tbl",""]` 占位 |
| `fig` | ✓ | `[文件名, 图注]` |
| `anal` | ✓ | 解析块，`"p"` 普通段、`"li"` 带 ■ 的条目 |
| `summ` | ✓ | 答案行数组，**只写答案** |

整理原稿时按这个顺序修：**去重 → 断句 → 补表格 → 修公式 → 提炼答案**。常见源文件缺陷：

- 同一段题目/检索表重复出现两遍 → 只留一份；
- 每题都有一行 `配图：第 N 张（…）` 的占位提示 → 删掉，改由图注承担；
- `\dfrac`、`Command("dfrac")` 之类解析失败的公式残留 → 改写成 `V = (V总 − V水) ÷ n` 这种纯文本 + `{#}`/`{^}` 标记；
- 题干、小问、选项连成一大段 → 按 `（1）（2）`、`A. B. C.` 拆成独立段落类型。

### ④ 首排

```bash
python assets/build_docx.py --content content.json --figs figs --out 错题集.docx
```

### ⑤ 实测收敛（核心，别跳）

python-docx 不知道"页"是什么，必须借排版引擎量：

```bash
pwsh -File assets/measure.ps1 -Docx 错题集.docx          # 每题的 Tq/Ta/yans/p/pa/img
python assets/fit.py --measure m.txt --content content.json --figs figs \
                     --fit fit.json --side 01,02,04
python assets/build_docx.py ... --fit fit.json --side 01,02,04
```

判据只有一条：**每题的 `p == pa`（题目页 = 答案页），且 `pages == 题数`**。不满足就再跑一轮 `measure → fit → build`。

`fit.py` 的取舍顺序是固定的：**先压截图高度（≥4.5cm），压到底还超才压字号（≥0.78）**。字号是读者体验，图是佐证材料，所以图先让。

### ⑥ 定稿

- `--side` 的题号一旦变化，`fit.json` 里的 `sidew` 要重新解一轮；
- PDF：`measure.ps1 -Pdf`（WPS 导出，页数应与 docx 一致，不一致说明字体缺失）；
- 校对：dump 一遍成品文本，重点看**图题是否一一对应**（最容易错的地方，因为图与题本来没有绑定关系）。

---

## 踩坑表

| # | 现象 | 原因 | 解法 |
|---|---|---|---|
| 1 | 页数从 13 涨到 29 | 给主笔记区行加了 `cantSplit`，超高行被整块推到下一页 | 只给题头行、答案区行加 `cantSplit` |
| 2 | Word 打开提示"内容有问题" | 手工 append 的 `w:tblPr`/`w:tcPr`/`w:pPr` 子元素顺序不符合 schema（如 `tblBorders` 写在 `tblLayout` 之后） | 保存前跑一遍 `canonicalize()`，按 schema 顺序重排五类属性容器 |
| 3 | `tblLayout` 重复 | `table.autofit = False` 已经写了一次，又手动 append 一次 | 只让 python-docx 写 |
| 4 | 脚本报"文档打开失败"/路径变乱码 | Windows PowerShell 5.1 把 UTF-8 无 BOM 的脚本按 ANSI 读，中文路径损坏 | 一律用 **pwsh 7** 跑脚本 |
| 5 | `PermissionError` / `文件被占用` | 上一轮 WPS COM 没退干净，句柄还开着 | `Get-Process wps | Stop-Process -Force`；仍占用就 `Remove-Item` 后再写 |
| 6 | 截图被压成 3cm 宽的一小条 | 为了塞进一页，把图高一路压到底 | 比例 ≥ 0.9 的图改用**并列版式**：图进解析右侧独立栏，通栏图带整条省掉，图反而能放大到 8.5cm 且更矮 |
| 7 | 并列版式后文字挤成"一柱擎天" | 右栏给太宽，左栏只剩 5cm | 右栏上限 9.0cm（左栏 ≥ 8cm），超出就换回通栏 |
| 8 | 答案行贴着页脚没了 | 块高算到 26.8 正好用完，答案行溢出 | 目标值取 `limit = 26.15`，留 0.65cm 余量 |
| 9 | 打印出现方框/豆腐块 | `✓ ✗ ▸` 等符号在部分中文字库缺字形 | 统一换 `√ × ■ ▲ □`（GBK 全覆盖） |
| 10 | 图片超出单元格右边框 | 嵌套表宽 = 笔记区可用宽 + 图片加边距 | 图宽按 `栏宽 − 0.45`、`avail = 内容宽 − 0.75` 留边 |

## 验收清单

- [ ] `pages == 题数`，且每题 `p == pa`
- [ ] 每题恰好 1 张图，图注写明卷面来源
- [ ] 所有图宽 ≥ 6cm，横长条图走通栏（≥9cm）
- [ ] 每页答案区都在页面内，不贴页脚
- [ ] docx 与 PDF 页数一致
- [ ] 抽样 3 题：题干/小问/选项/表格/解析/答案与原卷核对无误
