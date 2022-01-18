Attribute VB_Name = "Main"
Public statement_name, statement_file, statement_path As String
Public data_name, data_file As String
Public manager_name As String

Const STATEMENT_SHEET = "TDSheet"
Const COL_STATEMENT_WORKER = 13
Const ROW_STATEMENT_WORKER = 16
Const COL_STATEMENT_ORDER = 15

Const DATA_SHEET = "По сотрудникам"
Const COL_DATA_ORDER = 1
Const ROW_DATA_ORDER = 2
Const COL_DATA_WORKER = 2



Sub Main()
'
' Управление переносом данных
'

    Dim last_worker_row_statement, last_order_row_data As Integer
    Dim last_worker_col_data As Integer
    
    If manager_name = "" Then
        manager_name = ThisWorkbook.Name
    End If
    
    Set Manager = Workbooks(manager_name).Sheets("Управление")
    
    If Not isOpenFile(statement_file) Then
        MsgBox ("Не найден файл " & Chr(34) & "Ведомости" & Chr(34) & Chr(13) & _
                "Попробуйте повторно выбрать файл.")
        Manager.Cells(1, 2) = "Файл не выбран"
        Manager.Cells(1, 2).Style = "Плохой"
        Exit Sub
    End If

    If Not isOpenFile(data_file) Then
        MsgBox ("Не найден файл " & Chr(34) & "Данных для ведомости" & Chr(34) & Chr(13) & _
                "Попробуйте повторно выбрать файл.")
        Manager.Cells(2, 2) = "Файл не выбран"
        Manager.Cells(2, 2).Style = "Плохой"
        Exit Sub
    End If
    
    Manager.Cells(3, 2) = "Идёт перенос файлов..."
    Manager.Cells(3, 2).Style = "Нейтральный"
        
    Application.ScreenUpdating = False
    
        last_worker_row_statement = getLastRow((statement_name), (STATEMENT_SHEET), _
                                 (ROW_STATEMENT_WORKER), (COL_STATEMENT_WORKER))
        last_order_row_data = getLastRow((data_name), (DATA_SHEET), _
                                 (ROW_DATA_ORDER), (COL_DATA_ORDER))
        last_worker_col_data = getLastCol((data_name), (DATA_SHEET), 1, 2)
          
        Call addRows((last_worker_row_statement), _
                     (last_order_row_data), _
                     (last_worker_col_data))
        
        Workbooks(manager_name).Activate
        new_name = "С ДАННЫМИ " + statement_name
        full_new_name = statement_path + new_name
        Workbooks(statement_name).SaveAs (full_new_name)
        Workbooks(new_name).Close
        Workbooks(data_name).Close
                
        Application.StatusBar = "Данные перенесены в файл: " + full_new_name
        Manager.Cells(3, 2) = "Данные перенесены и сохранены в файл: " & _
                              Chr(32) & Chr(10) & Chr(34) & new_name & Chr(34)
        Manager.Cells(3, 2).Style = "Хороший"
        
        Manager.Cells(1, 2) = "Выберите файл ведомости..."
        Manager.Cells(1, 2).Style = "Нейтральный"
        
        Manager.Cells(2, 2) = "Выберите файл данных для ведомости..."
        Manager.Cells(2, 2).Style = "Нейтральный"
        
    Application.ScreenUpdating = True

End Sub


Private Sub addRows(last_worker_row_statement As Integer, _
                    last_order_row_data As Integer, _
                    last_worker_col_data As Integer)
                    
'
' Добавление строк с данными в ведомость
'
Dim flag_copy As Boolean
Set Statement = Workbooks(statement_name).Sheets(STATEMENT_SHEET)
Set Data = Workbooks(data_name).Sheets(DATA_SHEET)

' Перебор всех работников из ведомости
For index_row_statement = last_worker_row_statement _
            To ROW_STATEMENT_WORKER Step -1
    current_worker_statement = Statement.Cells(index_row_statement, _
                                               COL_STATEMENT_WORKER)
    Application.StatusBar = "Поиск данных для: " & current_worker_statement
    
    ' Перебор всех сотрудников из "данных для ведомости"
    For index_col_data = last_worker_col_data To COL_DATA_WORKER Step -1
        flag_copy = False
        current_worker_data = Data.Cells(1, index_col_data)
        If current_worker_statement = current_worker_data Then
            ' Перебор всех заказов из "данных для ведомости"
            For index_current_order_data = last_order_row_data _
                    To ROW_DATA_ORDER Step -1:
                current_order_data = Data.Cells(index_current_order_data, _
                                                index_col_data)
                ' При наличии процента заказа перенос данных
                If current_order_data Then
                    If flag_copy Then
                        Statement.Rows(index_row_statement).Copy
                        Statement.Rows(index_row_statement + 1).Insert Shift:=xlDown
                        Statement.Cells(index_row_statement + 1, 1) = ""
                        Statement.Cells(index_row_statement + 1, 15) = _
                            Data.Cells(index_current_order_data, 1)
                        Statement.Cells(index_row_statement + 1, 16) = _
                            Data.Cells(index_current_order_data, index_col_data)
                    Else
                        Statement.Columns(15).ColumnWidth = 15
                        Statement.Cells(index_row_statement, 15) = _
                            Data.Cells(index_current_order_data, 1)
                        Statement.Cells(index_row_statement, 16) = _
                            Data.Cells(index_current_order_data, index_col_data)
                    End If
                    flag_copy = True
                End If
            Next index_current_order_data
            Exit For
        ' Не найдено ни одного совпадения сотрудников
        ElseIf index_col_data = COL_DATA_WORKER Then
            MsgBox (current_worker_statement & " не найден в файле:" & _
                    Chr(13) & Chr(34) & data_name & Chr(34) & Chr(13) & _
                    "ФИО должны в точности совпадать.")
        End If
    Next index_col_data
Next index_row_statement

End Sub



Private Function getLastCol(workbook_name As String, _
                            sheet_name As String, _
                            row_index As Integer, _
                            col_index As Integer) As Integer
'
' Получение номера последнего столбца относительно переданной ячейкм
'
getLastCol = Workbooks(workbook_name).Sheets(sheet_name). _
                Cells(row_index, col_index).End(xlToRight).Column
End Function


Private Function getLastRow(workbook_name As String, _
                            sheet_name As String, _
                            row_index As Integer, _
                            col_index As Integer) As Integer
'
' Получение номера последней строки относительно переданной ячейкм
'
getLastRow = Workbooks(workbook_name).Sheets(sheet_name). _
                Cells(row_index, col_index).End(xlDown).Row
End Function


Private Function isOpenFile(ByVal filePath As String) As Boolean
'
' Проверка существования и открытие файла файла.
'
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(filePath) = True Then
        Workbooks.Open filePath
        isOpenFile = True
    Else
        isOpenFile = False
    End If
End Function


Private Sub Auto_Open()
'
' Действия при открытии файла управления.
'
    manager_name = ThisWorkbook.Name
    Set Manager = Workbooks(manager_name).Sheets("Управление")
    
    Manager.Cells(1, 2) = "Выберите файл ведомости..."
    Manager.Cells(1, 2).Style = "Нейтральный"
    
    Manager.Cells(2, 2) = "Выберите файл данных для ведомости..."
    Manager.Cells(2, 2).Style = "Нейтральный"
    
    Manager.Cells(3, 2) = ""
    Manager.Cells(3, 2).Select
    Selection.Style = "Normal"
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Manager.Cells(1, 1).Select
End Sub
