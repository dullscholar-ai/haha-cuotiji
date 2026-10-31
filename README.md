# 哈哈错题集 · haha-cuotiji

一个 Agent Skill：**把试卷/作业的「题目数据 + 卷面截图」排成一题一页的错题集 docx**（可导出 PDF），
每题一张对应截图、截图完整且可读、整题不跨页。

- 默认版式：`题头 → ① 原题 → ② 解题过程（右侧并排卷面原件）→ ③ 答案`
- 不要封面、目录、编排说明、线索栏、总结栏、错因、订正区
- 一题一页用「1×1 表格 + `cantSplit`」实现，不是分页符
- 图能放多大由**实测反解**决定：`cap_cm = 26.4 − 文字高 − 0.85`（上限 16cm）
- 整页长截图（`h/w > 2.2`）自动转「图版页」，两栏并排、绝不截断

## 实测效果（13 题·初中科学）

| 版本 | 页数 | 图尺寸 | 备注 |
|---|---|---|---|
| 初版康奈尔三栏 | 31 | 附录统一收录 | 封面+说明+索引+错因+14 页附录 |
| 精简后 | 13 | 8 × 8.5cm | 只留题头/①/②/③ |
| **当前** | **14** | **9–11 × 10.5–16cm** | 图与解析并排；第 10 题整页图版 |

图面积从 68cm² 提到 130–176cm²（约 2.6 倍），并修掉「长截图只贴了 1/3」的显示不全问题。

## 快速开始

```bash
# A. 直接 JSON → docx
python assets/build_docx.py --content content.json --figs figs --out 错题集.docx

# B. 生产路线（本仓库成品用的，含实测反解）
python assets/production/postprocess.py                              # 先出精简 HTML
python assets/production/build_docx_local.py in.html _cal.docx --cal # 校准稿
pwsh -File assets/production/calibrate.ps1 -Docx _cal.docx -Figcaps figcaps.json
python assets/production/build_docx_local.py in.html 错题集.docx     # 正式稿
pwsh -File assets/production/calibrate.ps1 -Docx 错题集.docx -Audit  # 验收
```

## 文档

| 文件 | 内容 |
|---|---|
| `SKILL.md` | 技能说明：默认决策、版式速查、四条硬约束 |
| `references/final-layout.md` | 最终版式几何、反解公式、**六轮返工记录**、验收清单 |
| `references/workflow.md` | 全流程 + 踩坑表 |
| `references/layout-spec.md` | 康奈尔三栏备选样式 |

## 环境

Windows + Word/WPS（COM 实测分页）；Python 3.11+（python-docx、beautifulsoup4、Pillow）。
脚本一律用 PowerShell 7（`pwsh`）运行。

## License

MIT
