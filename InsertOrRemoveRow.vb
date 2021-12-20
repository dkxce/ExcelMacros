Sub РаздвинутьЯчейки()
'
' Макрос - Развинуть ячейки с сохранением формул
' Author: Milok Zbrozek <milokz@gmail.com>
'
  Dim i, X, Y As Integer
  Dim A() As String
  Dim sP, sC As String
  
  X = ActiveCell.Row
  Y = Cells(X, Columns.Count).End(xlToLeft).Column
  
  ReDim A(Y)
  For i = 1 To Y
    sP = Cells(X - 1, i).FormulaR1C1
    sC = Cells(X, i).FormulaR1C1
    If sP = sC Then
        A(i) = sP
    End If
  Next i
  
  Rows(X & ":" & X).Select
  Selection.Insert Shift:=xlDown
  For i = 1 To Y
    If Not A(i) = "" Then
       Cells(X - 1, i).FormulaR1C1 = A(i)
       Cells(X, i).FormulaR1C1 = A(i)
       Cells(X + 1, i).FormulaR1C1 = A(i)
    Else
       Cells(X, i).Value = Cells(X - 1, i).Value
    End If
  Next i
  
End Sub


Sub УдалитьСтроку()
'
' Макрос - Удалить строку с сохранение формул
' Author: Milok Zbrozek <milokz@gmail.com>
'
  Dim i, X, Y As Integer
  Dim A() As String
  Dim sP, sC As String
  
  X = ActiveCell.Row
  Y = Cells(X, Columns.Count).End(xlToLeft).Column
  
  ReDim A(Y)
  For i = 1 To Y
    sP = Cells(X - 1, i).FormulaR1C1
    sC = Cells(X + 1, i).FormulaR1C1
    If sP = sC Then
        A(i) = sP
    End If
  Next i
  
  Rows(X & ":" & X).Select
  Selection.Delete Shift:=xlUp
  For i = 1 To Y
    If Not A(i) = "" Then
       Cells(X - 1, i).FormulaR1C1 = A(i)
       Cells(X, i).FormulaR1C1 = A(i)
    End If
  Next i
  
End Sub



