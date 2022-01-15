Attribute VB_Name = "Settings"
Sub workbookOpen()
'
' Запуск управляющей формы
'
    For i = 0 To 9
        MainForm.ComboBox1.AddItem CStr(i)
        MainForm.ComboBox2.AddItem CStr(i)
    Next i
   
    MainForm.Show
   
End Sub


Sub ResetRange()
'
' Сброс скрола
'
   ActiveSheet.UsedRange
End Sub


Sub clearAll()
Attribute clearAll.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Очистка значений и свойств
'
Dim lLastRow As Long

    lLastRow = Macros.getLastValue()
    
    If lLastRow > 1 Then
        ActiveSheet.Range(Cells(2, 1), Cells(lLastRow, 3)).Select
        Selection.Clear
        ResetRange
    End If
    Cells(2, 1).Select
End Sub

