$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = $true

$presentation = $ppt.Presentations.Open(
"C:\Users\nikhi\DashboardPortfolioSource\May 2026 Analytics Overview_Data through April 2026.pptx"
)

$outDir = "C:\Users\nikhi\Indhu6.github.io\projects\images\communications"

if (!(Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir
}

$presentation.SaveAs($outDir,18)

$presentation.Close()
$ppt.Quit()