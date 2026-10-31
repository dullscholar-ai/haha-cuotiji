# calibrate.ps1 —— 一题一页 docx 的实测 / 反解 / 验收
#
# 三种用法：
#   1) 反解 figcaps.json（先用 --cal 模式排一版：每题强制独立一页、图高固定 6cm）
#      pwsh -File calibrate.ps1 -Docx _cal.docx -Figcaps figcaps.json
#   2) 验收正式稿
#      pwsh -File calibrate.ps1 -Docx 错题集.docx -Audit
#   3) 只导 PDF
#      pwsh -File calibrate.ps1 -Docx 错题集.docx -Pdf 错题集.pdf
#
# 为什么必须实测：python-docx 写出来的 docx 没有「页」的概念，图能放多大、
# 哪题跨页，只有排版引擎知道。这里把结果量化成数字，再按公式反解。
#
# 反解公式（每题文字高 text_cm = 内容高 - 校准图高 - 图注）：
#   cap_cm = clamp(6.0, 16.0, Avail - text_cm - 0.85)
# 其中 Avail = 页面高度 - 上下边距（默认 26.4cm），0.85 是图注 + 段落间距 + 余量。
#
# 注意：必须用 PowerShell 7（pwsh）运行；Windows PowerShell 5.1 会把无 BOM 的
# UTF-8 脚本按 ANSI 读，中文路径直接坏掉。

param(
  [Parameter(Mandatory = $true)][string]$Docx,
  [string]$Figcaps = 'figcaps.json',
  [string]$Pdf = '',
  [switch]$Audit,
  [double]$Avail = 26.4,      # 单页可用高度 cm
  [double]$CalImg = 6.0,      # --cal 模式下的固定图高 cm
  [double]$CapMin = 6.0,
  [double]$CapMax = 16.0
)

$ErrorActionPreference = 'Stop'
$PT = 28.3465
$docxPath = (Resolve-Path -LiteralPath $Docx).Path

# Word 优先，没装 Word 就退回 WPS
try { $app = New-Object -ComObject Word.Application }
catch { $app = New-Object -ComObject KWps.Application }
$app.Visible = $false
try { $app.DisplayAlerts = 0 } catch { }

$doc = $app.Documents.Open($docxPath, $false, $true)
$doc.Repaginate()
$pages = $doc.ComputeStatistics(2)

if ($Audit) {
  "pages=$pages"
  $sizes = @()
  foreach ($sh in $doc.InlineShapes) {
    $sizes += ('{0:N1}x{1:N1}' -f ($sh.Width / $PT), ($sh.Height / $PT))
  }
  "images=$($sizes.Count)  " + ($sizes -join '  ')

  $per = @{}
  for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
    $r = $doc.Paragraphs.Item($i).Range
    $t = ($r.Text -replace "[\r\n\t]", ' ').Trim()
    if ($t.Length -eq 0) { continue }
    try { $pg = $r.Information(3); $y = $r.Information(6) } catch { break }
    if (-not $per.ContainsKey($pg)) { $per[$pg] = @{ first = $t; max = $y } }
    elseif ($y -gt $per[$pg].max) { $per[$pg].max = $y }
  }
  foreach ($k in ($per.Keys | Sort-Object)) {
    $head = $per[$k].first
    if ($head.Length -gt 28) { $head = $head.Substring(0, 28) }
    'p{0} bottom={1:N0}pt | {2}' -f $k, $per[$k].max, $head
  }
  $doc.Close(0)
  try { $app.Quit() } catch { }
  return
}

# ---- 反解模式：cal 稿每题一页，按页序对应题序 -----------------------------
$per = @{}
for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
  $r = $doc.Paragraphs.Item($i).Range
  $t = ($r.Text -replace "[\r\n\t]", ' ').Trim()
  if ($t.Length -eq 0) { continue }
  try { $pg = $r.Information(3); $y = $r.Information(6) } catch { break }
  if (-not $per.ContainsKey($pg)) { $per[$pg] = @{ min = $y; max = $y } }
  else {
    if ($y -lt $per[$pg].min) { $per[$pg].min = $y }
    if ($y -gt $per[$pg].max) { $per[$pg].max = $y }
  }
}

if ($Pdf -ne '') {
  $doc.ExportAsFixedFormat((Join-Path (Get-Location) $Pdf), 17)
  "pdf=$Pdf"
}

$pagesSorted = $per.Keys | Sort-Object
$out = [ordered]@{}
$idx = 0
"{0,-4} {1,8} {2,8} {3,8}" -f 'no', 'content', 'text', 'cap'
foreach ($k in $pagesSorted) {
  $idx++
  $content = ($per[$k].max - $per[$k].min) / $PT + 0.45      # 末行行高补偿
  $text = $content - $CalImg - 0.45                          # 扣掉校准图 + 图注
  $cap = [math]::Round([math]::Min($CapMax, [math]::Max($CapMin, $Avail - $text - 0.85)), 2)
  $out["$idx"] = $cap
  "{0,-4} {1,8:N2} {2,8:N2} {3,8:N2}" -f $idx, $content, $text, $cap
}

($out | ConvertTo-Json -Compress) | Out-File -FilePath $Figcaps -Encoding utf8
"figcaps -> $Figcaps ($($out.Count) 题)"
"pages=$pages"

$doc.Close(0)
try { $app.Quit() } catch { }
