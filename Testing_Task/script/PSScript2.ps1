
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

#вызов функции записи дельты
   $RowsRenge=($ws.UsedRange.Rows).count
   $r=$ws.Range($ws.Cells.Item(2, 3),$ws.Cells.Item(($ws.UsedRange.Rows).count, 3))
   $r.FormulaLocal='=(B2-B3)*ЕСЛИ(B3>0;1;0)'
   $r=$ws.Range($ws.Cells.Item(2, 6),$ws.Cells.Item(($ws.UsedRange.Rows).count, 6))
   $r.FormulaLocal='=(E2-E3)*ЕСЛИ(E3>0;1;0)'
   $r=$ws.Range($ws.Cells.Item(2, 7),$ws.Cells.Item(($ws.UsedRange.Rows).count, 7))
   $r.Formula='=C2/F2'

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