'VBA for Excel
'Функция разборки строки на столбцы в зависимости от сдвига от начала строки
'Перед запуском необходимо выбрать первую ячейку (верхнюю), по столбцу которой будет идти анализ
'Добавляет текст справа в столбец, начиная с `startColumn`
Sub offset2stolb_by_selectedStolb_right()
    
    'Columns("A:A").Select
    
    Dim startColumn As Integer
    startColumn = 9 ' Right column to copy data
    
    Dim i As Integer ' for loop
    Dim off As Integer ' offset of a string
    Dim ind As Integer ' index if offsets array
    Dim arrayMax As Integer ' offsets array max index
    arrayMax = 1 ' zero offset is first element

    Dim offsets(1 To 100) As Integer ' offsets array
    offsets(1) = 0 ' zero is no offset
    
    Dim ex As Boolean ' offsets exists

    Dim eCount As Integer ' empty counter
    eCount = 0 ' not empty
    Dim eMax As Integer ' max empty cells
    eMax = 3 ' max 3 cells
        
    Do Until eCount >= eMax
       If IsEmpty(ActiveCell) Then 'check is empty
         eCount = eCount + 1 'if empty
       Else 'if not empty
         eCount = 0 'reset empty counter
         off = ActiveCell.IndentLevel 'get offset
         ex = False 'not exists in array
         ind = 0 'array index
         For i = 1 To arrayMax ' search in array
           If offsets(i) = off Then
             ex = True
             ind = i
           End If
         Next i
         If ex = False Then ' not found
            arrayMax = arrayMax + 1 'add element in array
            offsets(arrayMax) = off 'write offset
            ind = arrayMax 'index of element in array with that offset
         End If
         Cells(ActiveCell.Row, startColumn + ind - 1).Value = ActiveCell.Value ' copy value to new cell
       End If
       ActiveCell.Offset(1, 0).Select ' move down
    Loop

End Sub



