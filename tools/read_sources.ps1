$ErrorActionPreference = 'Stop'
$sourceRoot = 'F:\study\методичка'
$targetRoot = 'F:\study\inf-labs\tmp\sources'
New-Item -ItemType Directory -Path $targetRoot -Force | Out-Null
$sources = @(
    @('График лабораторных работ.doc', 'schedule'),
    @('Титульный лист.doc', 'title'),
    @('mathcad\Методичка по MathCAD(2-е работы).doc', 'mathcad'),
    @('HTML\Лабораторная работа. Форматирование. Титульный лист\Лабораторная работа. Форматирование. Титульный лист.doc', 'html')
)
$word = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    foreach ($entry in $sources) {
        $doc = $null
        try {
            $source = Join-Path $sourceRoot $entry[0]
            $doc = $word.Documents.Open($source, $false, $true, $false)
            $pdf = Join-Path $targetRoot ($entry[1] + '.pdf')
            $doc.ExportAsFixedFormat($pdf, 17)
            [System.IO.File]::WriteAllText((Join-Path $targetRoot ($entry[1] + '.txt')), $doc.Content.Text)
            Write-Output $pdf
        } finally {
            if ($null -ne $doc) { $doc.Close(0); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
        }
    }
} finally {
    if ($null -ne $word) { $word.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
}
