Attribute VB_Name = "Macros"
Function getLastValue() As Long
' Определение номера последней непустой стрки
    getLastValue = Cells(Rows.Count, 1).End(xlUp).Row
End Function


Function checkString(SearchString, SearchChar) As Boolean
'
' Определение наличия символа (последовательности символов) в строке
'

Dim position As Integer
position = InStr(1, SearchString, SearchChar, vbTextCompare)

    If position = 0 Then
        checkString = False
    Else
        checkString = True
    End If

End Function


Function setSearchText() As String
'
' Текст фильтра для сравниваемых ячеек
'
    If MainForm.CheckBox1.Value = True Then
        setSearchText = MainForm.TextBox1.Text
    Else
        setSearchText = ""
    End If

End Function


Sub SearchDoublesRow()
'
' Поиск номеров строк с повторами в столбце А
'
Dim i, j As Integer
Dim lLastRow As Long
Dim sDesiredValue, sSearchAreaValue As String
lLastRow = getLastValue()
Application.ScreenUpdating = False

    For i = 2 To lLastRow
        sDesiredValue = Cells(i, 1)
        If checkString(sDesiredValue, setSearchText()) Then
            For j = i + 1 To lLastRow
                sSearchAreaValue = Cells(j, 1)
                If sDesiredValue = sSearchAreaValue Then
                    Cells(j, 2) = "да"
                    Cells(j, 3) = "# " + CStr(i)
                  
                    Dim k As Integer
                    For k = 1 To 3
                        Cells(j, k).Interior.Color = 10132207
                    Next k
                End If
            Next j
        End If
        Application.StatusBar = "Проверено: " & CStr(i) & " из " & CStr(lLastRow) & " строк"
    Next i

Application.StatusBar = False
Application.ScreenUpdating = True

End Sub


Sub ShowDoubl()
'
' Отображение повторяющихся строк
'
    SearchDoublesRow
End Sub


Sub delDoubl()
'
' Удаление повторяющихся строк
'
Dim i, j, k As Integer
SearchDoublesRow
i = getLastValue()
j = 0
    
Application.ScreenUpdating = False

    Do While i > 1
        If Cells(i, 2) = "да" Then
        
            ' Проверка и, при необходимости, удаление следующий за найденной строк
            If MainForm.ComboBox2.Value <> 0 Then
                For k = 1 To MainForm.ComboBox2.Value
                    Cells(i + k, 2).Select
                    Selection.EntireRow.Delete
                    j = j + 1
                Next k
            End If
            Cells(i, 2).Select
            Selection.EntireRow.Delete
            
            ' Проверка и, при необходимости, удаление следующий за найденной строк
            If MainForm.ComboBox1.Value <> 0 Then
                For k = 1 To MainForm.ComboBox1.Value
                    Cells(i - k, 2).Select
                    Selection.EntireRow.Delete
                    j = j + 1
                Next k
            End If
            j = j + 1
        End If
        i = i - 1
        Application.StatusBar = "Удаляется строка: " & CStr(i)
    Loop

Application.StatusBar = False
Application.ScreenUpdating = True
MsgBox ("Удалено " + CStr(j) + " строк")

End Sub
