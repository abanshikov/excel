Attribute VB_Name = "macros"
Sub main()
Attribute main.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Перенос данных из листа "ДАННЫЕ" на созданные листы
'

Dim Object_Str(50) As String
Dim TipColIS As String
Dim i, j As Integer
Dim MaxIS As Integer

Application.ScreenUpdating = False

    Sheets("ДАННЫЕ").Select
    MaxIS = Application.WorksheetFunction.Max(Range("A2:A100"))
        
    Sheets("ДАННЫЕ").Select
    Range("F2").Select
    Object_Str(0) = ActiveCell
    If (Object_Str(0) <> "") Then
        AddList (1)
        Sheets("ДАННЫЕ").Select
        For i = 0 To 49
            Object_Str(i) = Cells(2 + i, 6)
        Next
        ' запись из массива названия объекта
        Sheets("Акт 1").Select
        For i = 0 To 9
            Cells(11 + i, 1) = Object_Str(i)
        Next
        ' массив с типом и колличеством оборудования
        For i = 1 To MaxIS
            TipColIS = TipColIS + Sheets("ДАННЫЕ").Cells(i + 1, 2) + ", " & CStr(Sheets("ДАННЫЕ").Cells(i + 1, 3)) + "шт.; "
        Next
        ' название организации и колличество/тип оборудования
        Sheets("Акт 1").Select
        Cells(23, 1) = Sheets("ДАННЫЕ").Cells(2, 5)
        Cells(10, 1) = TipColIS
    ElseIf (Cells(2, 4) <> "") Then
        For j = 2 To 99
            If (Cells(j, 4) <> "") Then
                AddList (j - 1)
               Dim NameNewList As String
                NameNewList = "Акт " & CStr(j - 1)
                ' копирование данных
                With ActiveWorkbook.Sheets(NameNewList)
                    .Cells(10, 1) = Sheets("ДАННЫЕ").Cells(j, 2) + ", " & CStr(Sheets("ДАННЫЕ").Cells(j, 3)) + "шт.; "
                    .Cells(11, 1) = Sheets("ДАННЫЕ").Cells(j, 4)
                    .Cells(23, 1) = Sheets("ДАННЫЕ").Cells(2, 5)
                End With
            End If
        Next
    End If
    
Application.ScreenUpdating = True
    
End Sub


Function AddList(NumList As Integer)
    '
    ' добавление листа
    '
        
    Dim NameNewList As String
    NameNewList = "Акт " & CStr(NumList)
        
    ' добавление нового листа (копия шаблона)
    Sheets("ШАБЛОН").Copy After:=Sheets(NumList + 1)
    Sheets("ШАБЛОН (2)").Name = NameNewList

    ' цвета ярлычка
    With ActiveWorkbook.Sheets(NameNewList).Tab
        .Color = 65535
        .TintAndShade = 0
    End With
    
    Range("A1").Select
    Sheets("ДАННЫЕ").Select
    
End Function

' выбор подписывающего справку

Sub set_employee_1()
    Sheets("ШАБЛОН").Cells(39, 7) = "Начальник отдела Иванов И.И.."
End Sub

Sub set_employee_2()
    Sheets("ШАБЛОН").Cells(39, 7) = "Мастер отдела Петров П.П."
End Sub

Sub set_employee_3()
    Sheets("ШАБЛОН").Cells(39, 7) = "Мастер отдела Сидоров С.С."
End Sub

Sub set_employee_4()
    Sheets("ШАБЛОН").Cells(39, 7) = "Техник отдела Иванова Е.И."
End Sub
