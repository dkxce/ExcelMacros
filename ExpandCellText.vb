'VBA for Excel
Sub ExpandCellValueByGroupOrOffsetToSeveralColumns()
  '
  'Макрос: разбивка строки по столбцам в зависимости от группировки или отступа
  '        с добавление ячеек справа или слева с автозаполнение
  'Excel VBA Macro
  'Author: Milok Zbrozek <milokz@gmail.com>
  '
  
  'https://vremya-ne-zhdet.ru/vba-excel/funktsiya-msgbox-parametry/
  Dim dlgRes As Integer ' MsgBox result
   
  'Диалог запуска
  dlgRes = MsgBox("Запустить макрос разбивки строки по столбцам в зависимости от группировки или отступа?", vbOKCancel + vbQuestion, "Запуск макроса")
  If dlgRes = vbCancel Then
    Exit Sub 'Отмена
  End If
    
  ' Разбирать в зависимости от уровня группировки или отступа строки
  Dim byGroup As Boolean
  byGroup = True 'по умолчанию по уровню группировки
  
  'Диалог выбора группировки или отступа
  dlgRes = MsgBox("Разбивать строку в зависимости от группировки (да)" & vbNewLine & "или в зависимости от отступа (нет)?", vbYesNoCancel + vbQuestion, "Запуск макроса")
  If dlgRes = vbCancel Then
    Exit Sub 'Отмена
  End If
  If dlgRes = vbNo Then
    byGroup = False 'по отступу строки
  End If
  
  'Выводить разобранную информацию справа или слева
  Dim toRight As Boolean
  toRight = True 'по умолчанию справа
  
  'Диалог выбора вправо или влево
  dlgRes = MsgBox("Заполнять столбцы справа (да) - быстрее" & vbNewLine & "или добавлять слева (нет) - медленнее?", vbYesNoCancel + vbQuestion, "Запуск макроса")
  If dlgRes = vbCancel Then
    Exit Sub 'Отмена
  End If
  If dlgRes = vbNo Then
    toRight = False 'слева
  End If
    
  'Заполнять пустые ячейкаи значениями сверху
  Dim fillTop As Boolean
  fillTop = True 'по умолчанию заполнять пустые ячейкаи значениями сверху
  
  'Диалог заполнения пустых ячеек
  dlgRes = MsgBox("Заполнять пустые ячейки значениями сверху?", vbYesNoCancel + vbQuestion, "Запуск макроса")
  If dlgRes = vbCancel Then
    Exit Sub 'Отмена
  End If
  If dlgRes = vbNo Then
    fillTop = False 'не заполнять пустые ячейкаи значениями сверху
  End If
  
  'Объявление переменных
  ' i,k - энумераторы циклов
  ' off - уровень иерархии строки
  ' ind - номер в массиве уровня иерархии строки
  Dim i, k, off, ind As Integer
  Dim aYY(1 To 99) As Integer 'массив значений иерархий строки
  
  'Объявление переменных
  Dim rNo, rMax, cRead, cWrite, aMX As Integer
  rNo = 1 'стартовая строка
  rMax = Range("A" & Rows.Count).End(xlUp).Row 'конечная строка
  cRead = 1 'из какой колонки читать строку
  cWrite = 0 'отступ в какую колонку писать строку
  aMX = 0 'число заполненных ячеек в массиве иерархие
  
  If toRight Then 'при выводе справа ищем последнюю непустую колонку
    For i = 1 To 30 'перебираем первые 30 строк
        k = Cells(i, Columns.Count).End(xlToLeft).Column ' последняя непустая колонка
        If k > cWrite Then 'если последняя непустая строка дальше(правее)
            cWrite = k 'отступаем к ней
        End If
    Next i
  End If
  
  ' >>A
  For rNo = rNo To rMax 'перебираем все строки в таблице
    
    ' >>B
    If byGroups Then 'в зависимости от чего разбираем иерархию строки
      off = Cells(rNo, cRead).Rows(1).OutlineLevel 'получаем уровень группировки
    Else
      off = Cells(rNo, cRead).IndentLevel 'получаем отступ
    End If
    ' <<B
    
    ind = -1 'индекс иерархии в массиве (-1 - нет в массиве)
    
    ' >>C
    For i = 1 To aMX 'ищем по массиву
      If aYY(i) = off Then 'нашли
        ind = i 'присваиваем индекс иерархии
      End If
    Next i
    ' <CC
    
    ' >>D
    If ind = -1 Then  ' если значение отсутствует в массиве
      aMX = aMX + 1 'добавляем новый элемент в массив
      aYY(aMX) = off 'пишем в новый элемент значение иерархии
      ind = aMX 'присваиваем индекс иерархии
      If toRight = False Then 'если добавляем колонки слева, то делаем это
        Columns(aMX).EntireColumn.Insert 'добавляем колонку слева
        
        cRead = cRead + 1 'смещаем читаемую колонку на 1 вправо
      End If
    End If
    ' <<D
    
    'копируем текст из читаемой ячейки в новую
    Cells(rNo, cWrite + ind).Value = Cells(rNo, cRead).Value
    
    ' >>E
    If fillTop And rNo > 1 Then 'заполняем пустые ячейки значениями сверху
      For k = cWrite + 1 To cWrite + aMX 'проверяем все новые ячейки
        If IsEmpty(Cells(rNo, k)) Then 'если пустая
          Cells(rNo, k).Value = Cells(rNo - 1, k).Value 'пишем значение из верхней ячейки
        Else ' если нет
          k = 99 'обрываем цикл
        End If
      Next k
    End If
    ' <<E
    
  Next rNo
  ' <<A
  
  Dim res As String
  res = "Готово " & vbNewLine
  If toRight Then
     res = res & "Справа "
  Else
     res = res & "Слева "
  End If
  res = res & "добавлено " & aMX & " новых ячеек" & vbNewLine
  res = res & "Ячейки сформированы в зависимости от "
  If byGroup Then
     res = res & "группировки " & vbNewLine
  Else
     res = res & "отступа " & vbNewLine
  End If
  If fillTop Then
     res = res & "Пустые ячейки заполнены информцией сверху " & vbNewLine
  End If
  
  dlgRes = MsgBox(res, vbOKOnly + vbInformation) ', "Макрос выполнен")
End Sub


