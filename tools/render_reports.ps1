param([string[]]$Paths)
$ErrorActionPreference='Stop'
$qaRoot='F:\study\inf-labs\tmp\reports'
$poppler='C:\Users\artem\.cache\codex-runtimes\codex-primary-runtime\dependencies\native\poppler\Library\bin\pdftoppm.exe'
New-Item -ItemType Directory -Path $qaRoot -Force | Out-Null
$word=$null
try {
    $word=New-Object -ComObject Word.Application
    $word.Visible=$false;$word.DisplayAlerts=0;$word.AutomationSecurity=3
    foreach($path in $Paths) {
        $resolved=(Resolve-Path -LiteralPath $path).Path
        $name=(Get-Item -LiteralPath $resolved).Directory.Name
        $dir=Join-Path $qaRoot $name
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        $doc=$null
        try {
            $doc=$word.Documents.Open($resolved,$false,$true,$false)
            $pdf=Join-Path $dir 'report.pdf'
            $doc.ExportAsFixedFormat($pdf,17)
        } finally {if($null -ne $doc){$doc.Close(0);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc)}}
        & $poppler -scale-to 1600 -png $pdf (Join-Path $dir 'page')
        Write-Output $dir
    }
} finally {if($null -ne $word){$word.Quit();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)}}
