param([switch]$WordOnly,[switch]$ExcelOnly)
$ErrorActionPreference='Stop'
$root='F:\study\inf-labs'

if(-not $ExcelOnly){
 $word=$null;$doc=$null
 try{
  $word=New-Object -ComObject Word.Application;$word.Visible=$false;$word.DisplayAlerts=0;$word.AutomationSecurity=3
  $doc=$word.Documents.Open((Join-Path $root '05_Word\Дубограй_А_Ю_Word.docx'))
  $find=$doc.Content.Duplicate;$find.Find.ClearFormatting()
  if(-not $find.Find.Execute('[[WORDART]]')){throw 'WordArt anchor missing'}
  $find.Text='';$find.ParagraphFormat.SpaceAfter=130
  $wa=$doc.Shapes.AddTextEffect(0,'Дубограй Артем','Arial',36,$true,$false,25,10,$find)
  $wa.Name='WordArt_Dubogray';$wa.TextEffect.PresetShape=9;$wa.Width=430;$wa.Height=100
  $wa.Fill.PresetGradient(1,1,9);$wa.Line.ForeColor.RGB=65535;$wa.Line.Weight=2.25;$wa.Line.DashStyle=1;$wa.WrapFormat.Type=4
  $anchor=$doc.Content.Duplicate;$anchor.Find.ClearFormatting()
  if(-not $anchor.Find.Execute('[[CIRCUIT]]')){throw 'Circuit anchor missing'}
  $anchor.Text='';$anchor.ParagraphFormat.SpaceAfter=330
  $canvas=$doc.Shapes.AddCanvas(25,0,450,340,$anchor);$canvas.Name='Circuit_Variant6';$canvas.WrapFormat.Type=4
  $script:items=$canvas.CanvasItems
  $script:partNames=[System.Collections.Generic.List[string]]::new()
  function Part($sh){$sh.Name=('circuit_part_'+$script:partNames.Count);$sh.Line.ForeColor.RGB=0;$sh.Line.Weight=1.2;$script:partNames.Add($sh.Name);return $sh}
  function Wire([single]$x1,[single]$y1,[single]$x2,[single]$y2){$sh=$script:items.AddLine($x1,$y1,$x2,$y2);return (Part $sh)}
  function Box([int]$kind,[single]$x,[single]$y,[single]$w,[single]$h){$sh=$script:items.AddShape($kind,$x,$y,$w,$h);$sh.Fill.ForeColor.RGB=16777215;return(Part $sh)}
  function Label([string]$text,[single]$x,[single]$y,[single]$w=55,[single]$h=23){$sh=$script:items.AddTextbox(1,$x,$y,$w,$h);$sh.Name=('circuit_part_'+$script:partNames.Count);$sh.TextFrame.TextRange.Text=$text;$sh.TextFrame.TextRange.Font.Name='Arial';$sh.TextFrame.TextRange.Font.Size=10;$sh.TextFrame.MarginLeft=1;$sh.TextFrame.MarginTop=1;$sh.Line.Visible=0;$sh.Fill.Visible=0;$script:partNames.Add($sh.Name);return $sh}
  [void](Wire 85 35 245 35);[void](Wire 85 35 85 92);[void](Wire 165 35 165 65)
  [void](Box 9 67 92 36 36);[void](Label 'LDR' 26 98)
  [void](Wire 85 128 85 222);[void](Wire 85 128 85 145)
  [void](Wire 81 99 89 106);[void](Wire 89 106 80 114);[void](Wire 80 114 88 121)
  $ar=Wire 55 81 72 98;$ar.Line.EndArrowheadStyle=3
  $ar=Wire 48 88 65 105;$ar.Line.EndArrowheadStyle=3
  [void](Box 1 79 222 12 50);[void](Label 'R1' 44 235)
  [void](Wire 85 272 85 295);$ar=Wire 115 295 89 250;$ar.Line.EndArrowheadStyle=3
  [void](Wire 165 105 165 222);[void](Box 1 159 65 12 40);[void](Label 'R2' 178 76)
  [void](Box 1 159 222 12 50);[void](Label 'R3' 178 235);[void](Wire 165 272 165 295)
  [void](Wire 85 295 245 295);[void](Wire 165 295 165 309)
  [void](Wire 152 309 178 309);[void](Wire 156 315 174 315);[void](Wire 160 321 170 321)
  $tri=Box 7 215 135 60 76;$tri.Rotation=90
  [void](Label 'U1' 270 135);[void](Label '+' 217 153 14 18);[void](Label '−' 217 181 14 18)
  [void](Wire 165 145 198 145);[void](Wire 198 145 198 160);[void](Wire 198 160 215 160)
  [void](Wire 85 190 198 190);[void](Wire 198 190 198 186);[void](Wire 198 186 215 186)
  [void](Wire 245 35 245 135);[void](Wire 245 211 245 295)
  [void](Wire 275 173 325 173);[void](Box 1 325 167 45 12);[void](Label 'R4' 333 144)
  [void](Wire 370 173 408 173);[void](Box 9 408 170 6 6);[void](Label 'OUT' 405 146)
  [void](Wire 165 17 165 35);[void](Box 9 162 12 6 6);[void](Label '+V' 153 -8)
  foreach($xy in @(@(165,35),@(165,145),@(85,190),@(85,295),@(165,295))){$pt=Box 9 ($xy[0]-2) ($xy[1]-2) 4 4;$pt.Fill.ForeColor.RGB=0}
  $shapeNames=[object[]]$script:partNames.ToArray()
  $group=$canvas
  [void]$doc.Fields.Update()
  if($doc.TablesOfContents.Count -eq 0){$tocAnchor=$doc.Content.Duplicate;$tocAnchor.Collapse(0);[void]$doc.TablesOfContents.Add($tocAnchor,$true,1,1)}
  foreach($toc in $doc.TablesOfContents){$toc.Update()}
  $doc.SaveAs2((Join-Path $root '05_Word\Дубограй_А_Ю_Word.docx'),16)
  $doc.SaveAs2((Join-Path $root '05_Word\Дубограй_А_Ю_Word.doc'),0)
  Write-Output ('Word drawings: '+$doc.Shapes.Count+'; circuit canvas elements: '+$canvas.CanvasItems.Count)
 }finally{if($null -ne $doc){$doc.Close(0);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc)};if($null -ne $word){$word.Quit();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)}}
}

if(-not $WordOnly){
 $excel=$null;$book=$null
 try{
  $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AutomationSecurity=3
  $book=$excel.Workbooks.Open((Join-Path $root 'tmp\excel\base.xlsx'))
  $cal=$book.Worksheets.Item('Календарь')
  [void]$cal.Range('D8').AddComment('4 ноября — День народного единства. Государственный нерабочий праздник. Источник: https://mintrud.gov.ru/social/social/1366')
  [void]$cal.Range('C9').AddComment('10 ноября — День сотрудника органов внутренних дел Российской Федерации. Профессиональный праздник. Источник: https://mvdmedia.ru/news/official/pozdravlenie-prezidenta-rossiyskoy-federatsii-vladimira-vladimirovicha-putina-s-dnem-sotrudnika-orga/')
  [void]$cal.Range('G10').AddComment('21 ноября — День работника налоговых органов Российской Федерации. Источник: https://www.nalog.gov.ru/rn05/news/activities_fts/16583225/')
  $wordart=$cal.Shapes.AddTextEffect(0,'ОСЕНЬ','Arial',28,$true,$false,420,$cal.Range('E28').Top);$wordart.Fill.ForeColor.RGB=4163028;$wordart.Line.Visible=0
  # Insert an ordinary picture, drawn from vector geometry in this task.
  Add-Type -AssemblyName System.Drawing
  $bmp=[Drawing.Bitmap]::new(120,95);$g=[Drawing.Graphics]::FromImage($bmp);$g.Clear([Drawing.Color]::White)
  $g.SmoothingMode=[Drawing.Drawing2D.SmoothingMode]::AntiAlias
  $leaf=[Drawing.Point[]]@([Drawing.Point]::new(60,5),[Drawing.Point]::new(72,29),[Drawing.Point]::new(99,18),[Drawing.Point]::new(88,45),[Drawing.Point]::new(112,55),[Drawing.Point]::new(78,69),[Drawing.Point]::new(60,86),[Drawing.Point]::new(42,69),[Drawing.Point]::new(8,55),[Drawing.Point]::new(32,45),[Drawing.Point]::new(21,18),[Drawing.Point]::new(48,29))
  $brush=[Drawing.SolidBrush]::new([Drawing.Color]::FromArgb(190,99,30));$g.FillPolygon($brush,$leaf)
  $pen=[Drawing.Pen]::new([Drawing.Color]::FromArgb(109,63,18),2);$g.DrawLine($pen,60,25,60,94)
  $pic=Join-Path $root 'tmp\excel\autumn.png';$bmp.Save($pic,[Drawing.Imaging.ImageFormat]::Png);$pen.Dispose();$brush.Dispose();$g.Dispose();$bmp.Dispose()
  [void]$cal.Shapes.AddPicture($pic,$false,$true,110,$cal.Range('B27').Top,100,80)
  $pay=$book.Worksheets.Item('Таблица');$pay.Range('C5:E14').VerticalAlignment=-4130
  $surf=$book.Worksheets.Item('Поверхность');$chart=$surf.ChartObjects().Add(20,550,780,380).Chart
  $chart.SetSourceData($surf.Range('A7:N20'));$chart.ChartType=83;$chart.HasTitle=$true;$chart.ChartTitle.Text='z = x²/a² − y²/b²';$chart.Axes(1).HasTitle=$true;$chart.Axes(1).AxisTitle.Text='x, градусы';$chart.Axes(2).HasTitle=$true;$chart.Axes(2).AxisTitle.Text='y, градусы';$chart.Rotation=35;$chart.Elevation=25
  $chart.SetSourceData($surf.Range('B8:N20'),1)
  for($i=1;$i -le $chart.SeriesCollection().Count;$i++){$sr=$chart.SeriesCollection($i);$row=7+$i;$sr.Name=('=Поверхность!$A$'+$row);$sr.XValues=$surf.Range('B7:N7');$sr.Values=$surf.Range(('B'+$row+':N'+$row))}
  $chart.Axes(2).AxisTitle.Text='z';$chart.Axes(3).HasTitle=$true;$chart.Axes(3).AxisTitle.Text='y, градусы'
  $fun=$book.Worksheets.Item('Функция');$ch=$fun.ChartObjects(1).Chart;$ch.ChartType=74;$ch.Axes(1).CrossesAt=0;$ch.Axes(2).CrossesAt=0;$ch.Axes(1).MinimumScale=0
  foreach($c in @($ch,$chart)){
   $c.ChartArea.Format.Fill.Solid();$c.ChartArea.Format.Fill.ForeColor.RGB=16777215;$c.PlotArea.Format.Fill.Solid();$c.PlotArea.Format.Fill.ForeColor.RGB=16777215
   $c.ChartArea.Font.Color=0;$c.ChartTitle.Font.Color=0
   foreach($axis in $c.Axes()){$axis.TickLabels.Font.Color=0;if($axis.HasTitle){$axis.AxisTitle.Font.Color=0}}
   if($c.HasLegend){$c.Legend.Font.Color=0}
  }
  $m=$book.Worksheets.Item('Матрицы')
  foreach($name in @('A','B','C','D')){$address=@{A='B4:D6';B='F4:H6';C='J4:L6';D='N4:P6'}[$name];[void]$book.Names.Add(('Матрица_'+$name),('=Матрицы!$'+$address.Replace(':',':$').Replace('4','$4').Replace('6','$6')))}
  $excel.CalculateFullRebuild()
  if([Math]::Abs([double]$pay.Range('D5').Value2-330) -gt 0.00001){throw 'Payroll calculation mismatch'}
  $pay.Range('J5').Value2=0.12;$excel.CalculateFullRebuild();if([double]$pay.Range('D5').Value2-ne 360){throw 'Payroll input recalculation failed'};$pay.Range('J5').Value2=0.11
  $before=[double]$m.Range('J10').Value2;$m.Range('B4').Value2=4;$excel.CalculateFullRebuild();if([Math]::Abs([double]$m.Range('J10').Value2-$before-1)-gt 0.00001){throw 'Matrix input recalculation failed'};$m.Range('B4').Value2=3
  $excel.CalculateFullRebuild()
  foreach($sh in $book.Worksheets){
   $sh.UsedRange.Font.Underline=-4142
   $sh.PageSetup.Orientation=2;$sh.PageSetup.Zoom=$false;$sh.PageSetup.FitToPagesWide=1;$sh.PageSetup.FitToPagesTall=1
   $sh.PageSetup.LeftMargin=20;$sh.PageSetup.RightMargin=20;$sh.PageSetup.TopMargin=25;$sh.PageSetup.BottomMargin=25
   $sh.PageSetup.PrintArea=@{'Календарь'='B2:H31';'Таблица'='A1:M20';'Функция'='A1:N45';'Поверхность'='A1:N48';'Матрицы'='A1:P20'}[$sh.Name]
  }
  $cal.PageSetup.Orientation=1;$cal.Range('B2:H3').VerticalAlignment=-4108;$cal.Range('B2:H3').HorizontalAlignment=-4108
  $out=Join-Path $root '06_Excel';New-Item -ItemType Directory -Path $out -Force|Out-Null
  $book.CheckCompatibility=$false;$book.SaveAs((Join-Path $out 'Дубограй_А_Ю_Excel_1.xls'),56)
  $book.ExportAsFixedFormat(0,(Join-Path $root 'tmp\excel\final.pdf'))
  $errCount=0
  foreach($sh in $book.Worksheets){foreach($c in $sh.UsedRange.Cells){if($c.HasFormula -and $c.Text -match '^#'){$errCount++;Write-Output ($sh.Name+'!'+$c.Address()+': '+$c.Text)}}}
  if($errCount-ne 0){throw 'Formula errors in final workbook'}
  Write-Output ('Excel: '+$book.Worksheets.Count+' worksheets, '+$book.Names.Count+' named ranges, no formula errors.')
  Write-Output ('Payroll totals: '+$pay.Range('C15').Text+'; '+$pay.Range('G15').Text)
 }finally{if($null-ne $book){$book.Close($false);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)};if($null-ne $excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)}}
}
