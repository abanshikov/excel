VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DialogForm 
   Caption         =   "Создание заявки"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4485
   OleObjectBlob   =   "DialogForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DialogForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
    Unload Me
End Sub


Private Sub Ok_Click()
'
' Проверка для последней или нескольких заявок получить данные и 
' передача строк для создания заявок в главный модуль MainModule
'
    If CheckSearchLastString Then
        GetLastString
    Else
        GetManyPage
    End If
    Unload Me
End Sub


Private Sub GetLastString()
'
' Получение значения строки только для последней заявки. 
' Первая строка = 0
'
    NumberFirstString = 0
    NumberLastString = WorksheetFunction.CountA(Range("A:A")) + 1
End Sub


Private Sub GetManyPage()
'
' Получение номеров строк для нескольких заявок,
' проверка правильности ввода номеров строк.
'
    If Check_EnteringStrings Then
        NumberFirstString = DialogForm.FirstStringBox.Text
        NumberLastString = DialogForm.LastStringBox.Text
    End If
End Sub


Function Check_EnteringStrings() As Boolean
'
' Проверка правильности ввода строк
'
    If DialogForm.OptionManyStrings Then
        
        ' Проверка введения каких-либо данных
        If DialogForm.FirstStringBox.Text = "" Or DialogForm.LastStringBox.Text = "" Then
            MsgBox ("Необходимо ввести все номера строк")
            Check_EnteringStrings = False
        Else
            Check_EnteringStrings = True
        End If
        
        ' Проверка что значение последней строки больше первого
        If CInt(DialogForm.FirstStringBox.Text) > CInt(DialogForm.LastStringBox.Text) Then
            MsgBox ("Последняя строка должна быть больше начальной")
            Check_EnteringStrings = False
        Else
            Check_EnteringStrings = True
        End If
            
    End If
End Function

Private Sub OptionOneString_Click()
'
' Выбор опции "Для последней заявки".
'
    DialogForm.FrameNumbersStrings.Enabled = False
    DialogForm.FrameNumbersStrings.Visible = False
End Sub


Private Sub OptionManyStrings_Click()
'
' Выбор опции "Для нескольких заявок".
'
    DialogForm.FrameNumbersStrings.Enabled = True
    DialogForm.FrameNumbersStrings.Visible = True
End Sub

Private Sub FirstStringBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'
' Ввод только чисел в поле начальной строки
'
    Debug.Print KeyAscii
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        Debug.Print "number"
    Else
        Debug.Print "other"
        KeyAscii = 0
    End If
End Sub

Private Sub LastStringBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'
' Ввод только чисел в поле последней строки
'
    Debug.Print KeyAscii
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        Debug.Print "number"
    Else
        Debug.Print "other"
        KeyAscii = 0
    End If
End Sub

Function CheckSearchLastString() As Boolean
'
' Проверка что нужно искать последнюю строку
'
    If DialogForm.OptionOneString.Value Then
        CheckSearchLastString = True
    End If
End Function

