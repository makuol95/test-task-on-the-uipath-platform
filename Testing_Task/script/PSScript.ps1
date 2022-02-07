
#передача параметров
Param (
    [Parameter (Mandatory=$true)][System.Data.DataTable]$myDT,
    [Parameter (Mandatory=$true)][Boolean]$FileExict,
    [Parameter (Mandatory=$true)][string]$PathExcel
)

#функция записи данный
function Set-Data
{
    #запись шапки
    $ws.Cells.Item(1, 1)="Дата"
    $ws.Cells.Item(1, 2)="Курс"
    $ws.Cells.Item(1, 3)="Изменение"
    $ws.Cells.Item(1, 4)="Дата"
    $ws.Cells.Item(1, 5)="Курс"
    $ws.Cells.Item(1, 6)="Изменение"
    $ws.Cells.Item(1, 7)="Дельта"

    $Renge=2
    foreach($Row in $myDT)
    {
        
        $ws.Cells.Item($Renge, 1)=$Row.moment
        $ws.Cells.Item($Renge, 2)=$Row.value
        $ws.Cells.Item($Renge, 4)=$Row.moment
        $ws.Cells.Item($Renge, 5)=$Row.value_1
        $Renge++
    }
}

#функция записи дельты
function Сhange-Data{
   $RowsRenge=($ws.UsedRange.Rows).count
   $r=$ws.Range($ws.Cells.Item(2, 3),$ws.Cells.Item(($ws.UsedRange.Rows).count, 3))
   $r.FormulaLocal='=(B2-B3)*ЕСЛИ(B3>0;1;0)'
   $r=$ws.Range($ws.Cells.Item(2, 6),$ws.Cells.Item(($ws.UsedRange.Rows).count, 6))
   $r.FormulaLocal='=(E2-E3)*ЕСЛИ(E3>0;1;0)'
   $r=$ws.Range($ws.Cells.Item(2, 7),$ws.Cells.Item(($ws.UsedRange.Rows).count, 7))
   $r.FormulaLocal='=ЕСЛИОШИБКА(C2/F2;0)'
}

#функция записи дельты
function Сhange-Format{
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
 }

#добавляем/открываем эксель
$xl = new-object -comobject excel.application
$xl.visible = $true
if($FileExict){
    $wb = $xl.Workbooks.Open($PathExcel)
    $ws = $wb.Worksheets.Item(1)
    $ws.UsedRange.Clear() | Out-Null
}
else{
    $wb = $xl.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = 'курс валют'
}

#вызов функиции для записи
Set-Data

Сhange-Format

Сhange-Data


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