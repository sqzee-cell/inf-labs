$ErrorActionPreference='Stop'
$root='F:\study\inf-labs'
$word=$null;$doc=$null
try {
 $word=New-Object -ComObject Word.Application;$word.Visible=$false;$word.DisplayAlerts=0;$word.AutomationSecurity=3
 $doc=$word.Documents.Open((Join-Path $root '05_Word\Дубограй_А_Ю_Word.docx'))
 $headings=0
 foreach($p in $doc.Paragraphs){
  $t=$p.Range.Text.Trim()
  if($t -match '^Задание [1-7] ' -and -not $t.Contains("`t")){
   $p.Range.Style=-2;$p.OutlineLevel=1;$p.Range.Font.Name='Times New Roman';$p.Range.Font.Size=14;$p.Range.Font.Bold=$true;$p.Range.Font.Color=0;$headings++
   $p.Range.ParagraphFormat.PageBreakBefore=if($t -match '^Задание 1 '){0}else{-1}
   $p.Range.ParagraphFormat.KeepWithNext=-1
  }elseif($t -eq 'Оглавление'){$p.Range.Style=-1;$p.OutlineLevel=10;$p.Range.Font.Name='Times New Roman';$p.Range.Font.Size=14;$p.Range.Font.Bold=$true;$p.Range.ParagraphFormat.PageBreakBefore=-1;$p.Range.ParagraphFormat.KeepWithNext=-1}
 }
 foreach($shape in $doc.Shapes){
  if($shape.Name -eq 'Circuit_Variant6' -or $shape.Name -eq 'WordArt_Dubogray'){
   $shape.RelativeVerticalPosition=2;$shape.RelativeHorizontalPosition=2;$shape.Top=8;$shape.Left=20;$shape.LockAnchor=$true;$shape.WrapFormat.Type=3
   $shape.Anchor.ParagraphFormat.SpaceAfter=$shape.Height+20
   $shape.Anchor.ParagraphFormat.LineSpacingRule=4;$shape.Anchor.ParagraphFormat.LineSpacing=1
   if($shape.Name -eq 'Circuit_Variant6'){
    $parts=@();foreach($part in $shape.CanvasItems){$parts+=$part.Name}
    if($parts -notcontains 'PowerTopConnection'){$line=$shape.CanvasItems.AddLine(245,135,245,158);$line.Name='PowerTopConnection';$line.Line.Weight=1.2;$line.Line.ForeColor.RGB=0}
    if($parts -notcontains 'PowerBottomConnection'){$line=$shape.CanvasItems.AddLine(245,188,245,211);$line.Name='PowerBottomConnection';$line.Line.Weight=1.2;$line.Line.ForeColor.RGB=0}
    foreach($part in $shape.CanvasItems){
     if([Math]::Abs($part.Left-245)-lt 1 -and [Math]::Abs($part.Top-35)-lt 1 -and [Math]::Abs($part.Height-100)-lt 1){$part.Height=123}
     if([Math]::Abs($part.Left-245)-lt 1 -and [Math]::Abs($part.Top-211)-lt 1 -and [Math]::Abs($part.Height-84)-lt 1){$part.Top=188;$part.Height=107}
    }
   }
  }
 }
 $symbol=$doc.Content.Duplicate
 if($symbol.Find.Execute([string][char]0xF0C9)){$symbol.Text=[string][char]0x260E;$symbol.Font.Name='Segoe UI Symbol';$symbol.Font.Size=22}
 $text=$doc.Content.Duplicate
 if($text.Find.Execute('сгруппированной схемы')){$text.Text='схемы на полотне рисунка'}
 $doc.Repaginate()
 foreach($toc in $doc.TablesOfContents){$toc.Update()}
 $doc.Repaginate()
 foreach($toc in $doc.TablesOfContents){$toc.UpdatePageNumbers();Write-Output $toc.Range.Text}
 if($headings-ne 7){throw ('Expected 7 headings, got '+$headings)}
 if($doc.Content.Text -match 'Элементы оглавления не найдены|No table of contents entries found'){throw 'Contents still empty'}
 $doc.Save()
 $doc.ExportAsFixedFormat((Join-Path $root 'tmp\reports\05_Word\report.pdf'),17)
 $doc.SaveAs2((Join-Path $root '05_Word\Дубограй_А_Ю_Word.doc'),0)
 Write-Output 'Both Word formats saved with seven contents entries.'
}finally{
 if($null-ne $doc){try{$doc.Close(0)}catch{};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc)}
 if($null-ne $word){try{$word.Quit()}catch{};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)}
}
