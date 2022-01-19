VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Управление"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7875
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckBox1_Click()
'
' Выбор необходимости поиска символов в строке
'
    If MainForm.CheckBox1.Value = True Then
        MainForm.TextBox1.Enabled = True
        MainForm.TextBox1.SetFocus
    Else
        MainForm.TextBox1.Enabled = False
    End If
End Sub


Private Sub CommandButton1_Click()
'
' Поиск повторов
'
    Macros.ShowDoubl
    Unload MainForm
End Sub

Private Sub CommandButton2_Click()
'
' Удаление повторов
'
    Macros.delDoubl
    Unload MainForm
End Sub

Private Sub CommandButton3_Click()
'
' Очистка таблицы
'
    Settings.clearAll
    Unload MainForm
End Sub

Private Sub CommandButton4_Click()
'
' Закрытие окна формы
'
    Unload MainForm
End Sub
