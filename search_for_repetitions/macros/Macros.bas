Attribute VB_Name = "Macros"
Function getLastValue() As Long
' ����������� ������ ��������� �������� �����
    getLastValue = Cells(Rows.Count, 1).End(xlUp).Row
End Function


Function checkString(SearchString, SearchChar) As Boolean
'
' ����������� ������� ������� (������������������ ��������) � ������
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
' ����� ������� ��� ������������ �����
'
    If MainForm.CheckBox1.Value = True Then
        setSearchText = MainForm.TextBox1.Text
    Else
        setSearchText = ""
    End If

End Function


Sub SearchDoublesRow()
'
' ����� ������� ����� � ��������� � ������� �
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
                    Cells(j, 2) = "��"
                    Cells(j, 3) = "# " + CStr(i)
                  
                    Dim k As Integer
                    For k = 1 To 3
                        Cells(j, k).Interior.Color = 10132207
                    Next k
                End If
            Next j
        End If
        Application.StatusBar = "���������: " & CStr(i) & " �� " & CStr(lLastRow) & " �����"
    Next i

Application.StatusBar = False
Application.ScreenUpdating = True

End Sub


Sub ShowDoubl()
'
' ����������� ������������� �����
'
    SearchDoublesRow
End Sub


Sub delDoubl()
'
' �������� ������������� �����
'
Dim i, j, k As Integer
SearchDoublesRow
i = getLastValue()
j = 0
    
Application.ScreenUpdating = False

    Do While i > 1
        If Cells(i, 2) = "��" Then
        
            ' �������� �, ��� �������������, �������� ��������� �� ��������� �����
            If MainForm.ComboBox2.Value <> 0 Then
                For k = 1 To MainForm.ComboBox2.Value
                    Cells(i + k, 2).Select
                    Selection.EntireRow.Delete
                    j = j + 1
                Next k
            End If
            Cells(i, 2).Select
            Selection.EntireRow.Delete
            
            ' �������� �, ��� �������������, �������� ��������� �� ��������� �����
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
        Application.StatusBar = "��������� ������: " & CStr(i)
    Loop

Application.StatusBar = False
Application.ScreenUpdating = True
MsgBox ("������� " + CStr(j) + " �����")

End Sub
