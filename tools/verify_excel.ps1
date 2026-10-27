$ErrorActionPreference='Stop'
$excel=$null;$book=$null
try{
 $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AutomationSecurity=3
 $book=$excel.Workbooks.Open('F:\study\inf-labs\06_Excel\Дубограй_А_Ю_Excel_1.xls',0,$true)
 $fun=$book.Worksheets.Item('Функция');$surf=$book.Worksheets.Item('Поверхность');$mat=$book.Worksheets.Item('Матрицы')
 foreach($row in @(10,27,45)){
  $x=[double]$fun.Range('A'+$row).Value2
  $expected=2*[Math]::Sin(3*$x*[Math]::PI/180)/[Math]::Pow([Math]::Cos($x*[Math]::PI/180)+2,1.0/3.0)+[Math]::Log($x,3)
  if([Math]::Abs([double]$fun.Range('B'+$row).Value2-$expected)-gt 0.00000001){throw 'Function does not match independent calculation'}
 }
 if([Math]::Abs([double]$surf.Range('B8').Value2+5)-gt 0.000001){throw 'Surface corner mismatch'}
 if([double]$surf.Range('H14').Value2-ne 0){throw 'Surface origin mismatch'}
 $expected=@(@(6,11,11),@(5,4,10),@(11,11,17))
 for($i=0;$i-lt 3;$i++){for($j=0;$j-lt 3;$j++){if([double]$mat.Cells.Item((10+$i),(10+$j)).Value2-ne $expected[$i][$j]){throw 'Matrix product mismatch'}}}
 $cal=$book.Worksheets.Item('Календарь')
 if($cal.Range('H7').Value2-ne 1 -or $cal.Range('B12').Value2-ne 30){throw 'Calendar date position mismatch'}
 if($cal.Comments.Count-ne 3){throw 'Calendar notes missing'}
 if($surf.ChartObjects().Count-ne 1 -or $surf.ChartObjects(1).Chart.ChartType-ne 83){throw 'Native surface chart missing'}
 if($surf.ChartObjects(1).Chart.SeriesCollection().Count-ne 13){throw 'Surface series mismatch'}
 foreach($n in @('Матрица_A','Матрица_B','Матрица_C','Матрица_D')){[void]$book.Names.Item($n)}
 Write-Output 'Saved XLS: independent function, surface, matrix, calendar, comments, native charts and named ranges verified.'
}finally{if($null-ne $book){$book.Close($false);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)};if($null-ne $excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)}}
