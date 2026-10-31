# production/ —— 产出「最终版式」的那一套

这一套是**实际跑出 14 页成品**的代码（13 题 + 第 10 题整页图版页），不是示意。

```
content.json ──build_notebook.py──▶ 康奈尔 HTML ──postprocess.py──▶ 精简 HTML/MD
                                                          │
                                        build_docx_local.py (--cal) ──▶ _cal.docx
                                                          │
                                                  calibrate.ps1 ──▶ figcaps.json
                                                          │
                                        build_docx_local.py ──▶ 错题集.docx ──▶ PDF
```

## 环境变量 / 参数

| 项 | 说明 |
|---|---|
| `CJT_DIR` | 工作目录（默认脚本所在目录）。里面要有 `<code>.html`、`content.json`、`figures/` |
| `CJT_CODE` | 本册代号（如 `科学-20261031-001`）；目录里只有一个 html 时可省略 |
| `CJT_FIGCAPS` | `figcaps.json` 路径（默认当前目录） |

## 四个脚本

| 文件 | 作用 | 入口 |
|---|---|---|
| `postprocess.py` | 删封面/索引/编排说明/错因归纳/线索栏/总结栏/订正区；截图并到 ② 右侧；整页长图转图版页；③ 只留答案；重写 MD | `python postprocess.py` |
| `build_docx_local.py` | HTML → docx：整题包进 `1×1 表格 + cantSplit`；图高读 `figcaps.json`；页眉页脚；`--cal` 为测量模式 | `python build_docx_local.py in.html out.docx [--cal]` |
| `calibrate.ps1` | cal 稿实测 → 反解 `figcaps.json`；`-Audit` 验收；`-Pdf` 导 PDF | `pwsh -File calibrate.ps1 -Docx x.docx -Figcaps figcaps.json` |
| （上游）`build_notebook.py` | 来自 `康奈尔错题集` skill：`content.json` → HTML | `python build_notebook.py content.json .` |

## 一份 figcaps.json 长什么样

```json
{"1": 14.2, "2": 16.0, "3": 16.0, "4": 16.0, "5": 16.0,
 "7": 16.0, "8": 16.0, "9": 16.0, "10": 16.0, "11": 16.0,
 "12": 16.0, "13": 12.2}
```

键是**题序号（从 1 开始，按 JSON 里的顺序）**，值是这张图的高度上限（cm）。缺省题走 `est_cap_cm()` 的文字量估算，估得偏大时会溢出，所以正式排之前一定跑一遍校准。

## 已知取舍

- 图宽上限 11.2cm（左栏 5.6cm 给解析），竖长图高上限 16cm；两者同时卡住时取小的那个，保证「整张图都在」。
- 单题若天生超过一页（整页长截图），走**图版页**（单独一页、左右两栏 8.4cm），文字页末尾留 `▸ 卷面原图见下页` 指向它。
- 页数目标 = 题数，但整页长图的题会多占一页；这是为了完整性主动接受的代价。
