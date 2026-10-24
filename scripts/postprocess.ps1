<#
.SYNOPSIS
    用 Word COM 更新文档域与目录、保存、导出 PDF，最后把视图固定为页面视图。

.DESCRIPTION
    对应《学术论文 Word 排版与生成规范》§17.4 与 §18：
      1. Fields.Update()  → 更新所有域（目录、页码、交叉引用）
      2. TablesOfContents.Item(1).Update() → 更新目录
      3. 保存 + 导出 PDF
      4. Word 保存可能把视图改回 Web 版式，所以最后调用 set_view.py 重新写 settings.xml

.PARAMETER Docx
    要处理的 .docx 路径。

.PARAMETER Pdf
    导出的 PDF 路径，默认与 docx 同名的 .pdf。

.EXAMPLE
    pwsh -File postprocess.ps1 -Docx out\report.docx
    pwsh -File postprocess.ps1 -Docx out\report.docx -Pdf out\report.pdf
#>
param(
    [Parameter(Mandatory = $true)][string]$Docx,
    [string]$Pdf,
    [switch]$NoPdf,
    [switch]$NoViewReset
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $Docx)) {
    throw "找不到 docx：$Docx"
}
$Docx = (Resolve-Path -LiteralPath $Docx).Path
if (-not $Pdf) {
    $Pdf = [System.IO.Path]::ChangeExtension($Docx, ".pdf")
}
$PdfDir = Split-Path -Parent ([System.IO.Path]::GetFullPath($Pdf))
if ($PdfDir -and -not (Test-Path -LiteralPath $PdfDir)) {
    New-Item -ItemType Directory -Path $PdfDir -Force | Out-Null
}

$word = $null
$doc = $null
$pages = 0

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    try { $word.ScreenUpdating = $false } catch { }

    $doc = $word.Documents.Open($Docx, $false, $false)
    $doc.Fields.Update() | Out-Null

    # 页眉页脚里的域（PAGE / NUMPAGES）单独更新
    foreach ($sec in $doc.Sections) {
        foreach ($story in @($sec.Headers, $sec.Footers)) {
            foreach ($hf in $story) {
                try { $hf.Range.Fields.Update() | Out-Null } catch { }
            }
        }
    }

    if ($doc.TablesOfContents.Count -gt 0) {
        $doc.TablesOfContents.Item(1).Update()
    }
    $doc.Fields.Update() | Out-Null

    # 先在 Word 里切到页面视图 + 100%（后面还会用 python 兜底一次）
    try {
        $doc.ActiveWindow.View.Type = 3
        $doc.ActiveWindow.View.Zoom.Percentage = 100
    } catch { }

    $doc.Save()

    if (-not $NoPdf) {
        $doc.ExportAsFixedFormat($Pdf, 17)
    }
    try { $pages = $doc.ComputeStatistics(2) } catch { $pages = 0 }
    $doc.Close(0)
    $doc = $null
}
finally {
    if ($doc) { try { $doc.Close(0) } catch { } }
    if ($word) {
        try { $word.Quit() } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch { }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host ("Word 已更新域与目录：{0} 页" -f $pages)
if (-not $NoPdf) { Write-Host ("PDF：{0}" -f $Pdf) }

if (-not $NoViewReset) {
    python (Join-Path $PSScriptRoot "set_view.py") $Docx
    if ($LASTEXITCODE -ne 0) { throw "set_view.py 执行失败（视图未设置为页面视图）" }
}
