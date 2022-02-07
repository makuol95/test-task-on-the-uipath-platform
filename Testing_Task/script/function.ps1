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
        
        $ws.Cells.Item($Renge, 1)=$Row.one
        $ws.Cells.Item($Renge, 2)=$Row.two
        $ws.Cells.Item($Renge, 5)=$Row.tree
        $ws.Cells.Item($Renge, 4)=$Row.one
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
   $r.Formula='=C2/F2'
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