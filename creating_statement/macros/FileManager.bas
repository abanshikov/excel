Attribute VB_Name = "FileManager"
Public Sub OpenForm()
Attribute OpenForm.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Выбор файла ведомости
'

    Dim avFiles, address
        
    'по умолчанию к выбору доступны файлы Excel(xls,xlsx,xlsm,xlsb)
    avFiles = Application.GetOpenFilename _
                ("Excel files(*.xls*),*.xls*", 1, _
                "Выбрать Excel файлы", , False)
    If VarType(avFiles) = vbBoolean Then
        'была нажата кнопка отмены - выход из процедуры
        Cells(1, 2) = "Файл не выбран"
        statement_name = ""
        statement_file = ""
        Cells(1, 2).Style = "Плохой"
        Exit Sub
    End If
       
    address = Split(avFiles, "\")
    statement_name = address(UBound(address))
    ReDim Preserve address(UBound(address) - 1)
    statement_path = Join(address, "\") + "\"
    statement_file = avFiles
    
    Cells(1, 2) = "Выбран файл " & Chr(34) & statement_name & Chr(34)
    Cells(1, 2).Style = "Хороший"
    
    Cells(3, 2) = "Данные не перенесены"
    Cells(3, 2).Style = "Нейтральный"
    Application.StatusBar = "Данные не перенесены"
    
End Sub


Public Sub OpenData()
'
' Выбор файла с данными
'
    Dim avFiles, address
    
    'по умолчанию к выбору доступны файлы Excel(xls,xlsx,xlsm,xlsb)
    avFiles = Application.GetOpenFilename _
                ("Excel files(*.xls*),*.xls*", 1, "Выбрать Excel файлы", , False)
    If VarType(avFiles) = vbBoolean Then
        'была нажата кнопка отмены - выход из процедуры
        Cells(2, 2) = "Файл не выбран"
        data_name = ""
        data_file = ""
        Cells(2, 2).Style = "Плохой"
        Exit Sub
    End If
    
    address = Split(avFiles, "\")
    data_name = address(UBound(address))
    data_file = avFiles
    
    Cells(2, 2) = "Выбран файл " & Chr(34) & data_name & Chr(34)
    Cells(2, 2).Style = "Хороший"
    
    Cells(3, 2) = "Данные не перенесены"
    Cells(3, 2).Style = "Нейтральный"
    Application.StatusBar = "Данные не перенесены"
End Sub

