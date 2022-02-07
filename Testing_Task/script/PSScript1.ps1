
#передача параметров
$FileExict=$true
$PathExcel="C:\Users\lord-\Desktop\project\UIpath_Test_Task\Testing_Task\temp_excel\123.xlsx"
$myDT = new-object System.Data.DataTable

#добавляем/открываем эксель
$xl = new-object -comobject excel.application
$xl.visible = $true
if($FileExict){
    $wb = $xl.Workbooks.Open($PathExcel)
    $ws = $wb.Worksheets.Item(1)
    $ws.UsedRange.Clear()
}
else{
    $wb = $xl.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = 'курс валют'
}

#вызов функиции для записи
#выравнивание
  $ren=$ws.UsedRange
  $ren.EntireColumn.AutoFit() |  Out-Null
  #устанавливаем фомат для столбца с
  $r=$ws.Range($ws.Cells.Item(2, 3),$ws.Cells.Item(($ws.UsedRange.Rows).count, 3))
  $r.NumberFormatLocal='0,00'  
  #устанавливаем фомат для столбца f
  $r=$ws.Range($ws.Cells.Item(2, 6),$ws.Cells.Item(($ws.UsedRange.Rows).count, 6))
  $r.NumberFormatLocal='0,00'  
  #устанавливаем фомат для столбца g
  $r=$ws.Range($ws.Cells.Item(2, 7),$ws.Cells.Item(($ws.UsedRange.Rows).count, 7))
  $r.NumberFormatLocal='0,00'  
  #устанавливаем фомат для столбца b
  $r=$ws.Range($ws.Cells.Item(2, 2),$ws.Cells.Item(($ws.UsedRange.Rows).count, 2))
  $r.NumberFormatLocal='_-* # ##0,00 ₽_-;-* # ##0,00 ₽_-;_-* "-"?? ₽_-;_-@_-' 
  #устанавливаем фомат для столбца e
  $r=$ws.Range($ws.Cells.Item(2, 5),$ws.Cells.Item(($ws.UsedRange.Rows).count, 5))
  $r.NumberFormatLocal='_-* # ##0,00 ₽_-;-* # ##0,00 ₽_-;_-* "-"?? ₽_-;_-@_-' 
 
#сохраняем и закрываем данные
if($FileExict){
    $wb.save()
}
else{
    $wb.saveas($PathExcel)
}
$wb.Close()
$xl.Quit()

#проверяем работает ли эксель иначе все закрываем
try{
    if(Get-Process -ProcessName "EXCEL")
    {Stop-Process -Name "EXCEL"}
}
catch
{#
}