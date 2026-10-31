# measure.ps1 —— 用 WPS/Word COM 实测 docx 的真实分页与版面坐标
#
# 为什么需要它：python-docx 写出来的 docx 没有"页"的概念，一题是否跨页、
# 截图实际多大、解析栏有多高，只有排版引擎知道。本脚本把渲染结果量化成
# 数字，供 fit.py 反解"该压多少"。
#
# 用法（必须用 PowerShell 7 / pwsh 运行，见 references/workflow.md 的坑 4）：
#   pwsh -File measure.ps1 -Docx "错题集.docx"
#   pwsh -File measure.ps1 -Docx "错题集.docx" -Pdf "错题集.pdf"
#
# 输出（每行一条，供 fit.py 解析）：
#   pages=13
#   Q01 Tq=8.42 Ta=12.09 yans=26.13 p=1 pa=1 img=6.94x11.19
#     Tq    【题目】到【解析】之间的高度（cm，含通栏截图与表格）
#     Ta    【解析】到答案区之间的高度（cm，并列版式下=解析栏高度）
#     yans  答案区标签的纵坐标（cm，从页面顶端算）
#     p / pa  【题目】所在页 / 答案区所在页（两者不同 ⇒ 该题跨页，必须压）
#     img   该题截图的实际宽×高（cm）

param(
  [Parameter(Mandatory = $true)][string]$Docx,
  [string]$Pdf = '',
  [string]$AnchorStart = '【题目】',
  [string]$AnchorMid = '【解析】',
  [string]$AnchorAnswer = '答案区'
)

$ErrorActionPreference = 'Stop'
$TOPC = 1.55                 # 上边距，与 build_docx.py 保持一致
$USE = 29.7 - $TOPC - 1.35   # 单页可用高度 26.8cm

$app = New-Object -ComObject KWps.Application
Start-Sleep -Milliseconds 1500
$doc = $app.Documents.Open((Resolve-Path -LiteralPath $Docx).Path)

$total = $doc.ComputeStatistics(2)
"pages=$total"

# 图片尺寸按文档顺序取（每题恰好 1 张，与表格顺序一致）
$imgSizes = @()
foreach ($shape in $doc.InlineShapes) {
  $imgSizes += ('{0:N2}x{1:N2}' -f ($shape.Width / 28.3465), ($shape.Height / 28.3465))
}

# 题目表：3 行 1 列（题头 / 主笔记区 / 答案区）
$i = 0
foreach ($t in $doc.Tables) {
  if ($t.Rows.Count -ne 3 -or $t.Columns.Count -ne 1) { continue }
  $i++
  $notes = $t.Cell(2, 1).Range
  $ansCell = $t.Cell(3, 1).Range

  $a1 = $notes.Duplicate
  $a1.Find.ClearFormatting(); $a1.Find.Text = $AnchorStart; $a1.Find.Forward = $true
  $a2 = $notes.Duplicate
  $a2.Find.ClearFormatting(); $a2.Find.Text = $AnchorMid; $a2.Find.Forward = $true
  $a3 = $ansCell.Duplicate
  $a3.Find.ClearFormatting(); $a3.Find.Text = $AnchorAnswer; $a3.Find.Forward = $true
  if (-not ($a1.Find.Execute() -and $a2.Find.Execute() -and $a3.Find.Execute())) { continue }

  $p1 = $a1.Information(3); $y1 = [math]::Round($a1.Information(6) / 28.3465, 2)
  $p2 = $a2.Information(3); $y2 = [math]::Round($a2.Information(6) / 28.3465, 2)
  $p3 = $a3.Information(3); $y3 = [math]::Round($a3.Information(6) / 28.3465, 2)

  $tq = ($p2 - $p1) * $USE + ($y2 - $y1)
  $ta = ($p3 - $p2) * $USE + ($y3 - $y2)

  $img = '-'
  if ($i -le $imgSizes.Count) { $img = $imgSizes[$i - 1] }
  'Q{0:d2} Tq={1:N2} Ta={2:N2} yans={3:N2} p={4} pa={5} img={6}' -f $i, $tq, $ta, $y3, $p1, $p3, $img
}

if ($Pdf -ne '') {
  $doc.ExportAsFixedFormat((Join-Path (Get-Location) $Pdf), 17)
  "pdf=$Pdf"
}

$doc.Close(0)
try { $app.Quit() } catch { }
