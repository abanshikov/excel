Attribute VB_Name = "MainModule"
' Номера столбцов на листе "Регистрация заявок"
Const COL_NUM_REQUEST = 1
Const COL_DATE_REQUEST = 2
Const COL_NUM_INVOICE = 3
Const COL_NUM_INVOICE_ACTUAL = 4
Const COL_DATE_INVOICE = 5
Const COL_SUMM_AMOUNT = 6
Const COL_VAT_RATE = 7
Const COL_VAT_AMOUNT = 8
Const COL_DATE_AMOUNT = 13
Const COL_CONTRACT_NUMBER = 14
Const COL_REMARK = 15
Const COL_RESPONSIBLE = 16

' Номера столбцов на листе "Реестр контрагентов"
Const COL_NAME_OF_RECIPIENT = 1
Const COL_CONTRACT_NUMBER_REGICTRY = 2
Const COL_DATE_CONTRACT = 3
Const COL_TERMS_PAYMENT = 4
Const COL_PURPOSE_PAYMENT = 5
Const COL_INN = 6
Const COL_KPP = 7
Const COL_PAYMENT_ACCOUNT = 8
Const COL_BIK = 9
Const COL_NAMR_BANK = 10
Const COL_KBK = 11
Const COL_OKTMO = 12
Const COL_TAXABLE_PERIOD = 13
Const COL_UIN = 14

Const NAME_SHEET = "Заявка"
Public NumberFirstString As Integer
Public NumberLastString As Integer


Sub Main()
'
' Основная функция управления переносом и сохранения данных.
'
    Dim Folder As String
    Folder = ActiveWorkbook.Path + "\заявки\"
    
    ' Получение номеров строк для составления заявок.
    DialogForm.Show
    Application.ScreenUpdating = False
    Dim count, i, lenght  As Integer
    
        If NumberFirstString = 0 And NumberLastString = 0 Then
            ' Не выбраны никакие данные.
        ElseIf NumberFirstString = 0 Then
            ' Формирование заявки для последней строки.
            Application.StatusBar = "Формируется заявка для последней строки"
            Call CopyData(NumberLastString, Folder)
        Else
            ' Формирование заявок для нескольких строк
            For count = NumberFirstString To NumberLastString
                lenght = NumberLastString - NumberFirstString + 1
                i = count - NumberFirstString + 1
                Application.StatusBar = "Формируется заявка: " & CStr(i) & _
                                        " из " & CStr(lenght)
                Call CopyData(CInt(count), Folder)
            Next
        End If
        
    ThisWorkbook.Worksheets("Регистрация заявок").Select
    Application.StatusBar = "Все заявки сформированы"
    Application.ScreenUpdating = True
End Sub


Private Sub CopyData(NumStrRegistr As Integer, Folder As String)
'
' Копирование данных в лист заявки NAME_SHEET
' из строки NumStrRegistr листа "Регистрация заявок".
'

    ' Инициализация листов.
    Dim SheetRegistr As Worksheet
    Dim SheetCatalog As Worksheet
    Set SheetRegistr = ThisWorkbook.Worksheets("Регистрация заявок")
    Set SheetCatalog = ThisWorkbook.Worksheets("Реестр контрагентов")
    
    ' Определение строки организации из реестра организаций.
    Dim NumContract As String
    Dim NumStrCatalog As Integer
    NumContract = CStr(SheetRegistr.Cells(NumStrRegistr, COL_CONTRACT_NUMBER))
    NumStrCatalog = CInt(GetStringCatalog(NumContract))

    If NumStrCatalog <> 0 Then
        ' Копирование шаблона в лист заявки.
        CreateSheetRequest
        Dim SheetRequest As Worksheet
        Set SheetRequest = ThisWorkbook.Worksheets(NAME_SHEET)
    
        ' Дата платежа
        SheetRequest.Cells(7, 27) = CStr(SheetRegistr.Cells(NumStrRegistr, _
                                         COL_DATE_AMOUNT))
        ' Сумма платежа
        SheetRequest.Cells(8, 27) = CStr(SheetRegistr.Cells(NumStrRegistr, _
                                         COL_SUMM_AMOUNT))
        ' Ставка НДС
        SheetRequest.Cells(9, 27) = CStr(SheetRegistr.Cells(NumStrRegistr, _
                                         COL_VAT_RATE))
        ' Сумма НДС
        SheetRequest.Cells(10, 27) = CStr(SheetRegistr.Cells(NumStrRegistr, _
                                          COL_VAT_AMOUNT))
        ' Наименование получателя денежных средств
        SheetRequest.Cells(11, 27) = CStr(SheetCatalog.Cells(NumStrCatalog, _
                                          COL_NAME_OF_RECIPIENT))
        ' Дата и номер договора
        SheetRequest.Cells(12, 27) = "№" + CStr(SheetCatalog.Cells(NumStrCatalog, _
                                     COL_CONTRACT_NUMBER_REGICTRY)) + " от " + _
                                     CStr(SheetCatalog.Cells(NumStrCatalog, _
                                          COL_DATE_CONTRACT))
        ' Условия (срок) оплаты по договору
        SheetRequest.Cells(13, 27) = CStr(SheetCatalog.Cells(NumStrCatalog, COL_TERMS_PAYMENT))
        SheetRequest.Cells(13, 27).Rows.RowHeight = SetHeightRow(Len(SheetRequest.Cells(13, 27)))
        
        ' Документ основание (наименование дата номер)
        If SheetRegistr.Cells(NumStrRegistr, COL_NUM_INVOICE) <> "" Then
            SheetRequest.Cells(14, 27) = CStr(SheetRegistr.Cells(2, COL_NUM_INVOICE)) + _
                                         " №" + CStr(SheetRegistr.Cells(NumStrRegistr, _
                                         COL_NUM_INVOICE)) + " от " + _
                                         CStr(SheetRegistr.Cells(NumStrRegistr, _
                                              COL_DATE_INVOICE))
        Else
            SheetRequest.Cells(14, 27) = CStr(SheetRegistr.Cells(2, COL_NUM_INVOICE_ACTUAL)) + _
                                         " №" + CStr(SheetRegistr.Cells(NumStrRegistr, _
                                         COL_NUM_INVOICE_ACTUAL)) + " от " + _
                                         CStr(SheetRegistr.Cells(NumStrRegistr, _
                                              COL_DATE_INVOICE))
        End If
        
        ' Назначение платежа.
        SheetRequest.Cells(16, 27) = CStr(SheetCatalog.Cells(NumStrCatalog, _
                                          COL_PURPOSE_PAYMENT))
        ' Примечание
        SheetRequest.Cells(17, 27) = CStr(SheetRegistr.Cells(NumStrRegistr, _
                                          COL_REMARK))
        SheetRequest.Cells(17, 27).Rows.RowHeight = SetHeightRow(Len(SheetRequest.Cells(17, 27)))
        
        ' ИНН
        SheetRequest.Cells(19, 27) = CStr(SheetCatalog.Cells(NumStrCatalog, COL_INN))
        ' КПП
        SheetRequest.Cells(20, 27) = CStr(SheetCatalog.Cells(NumStrCatalog, COL_KPP))
        ' Расчетный счет Банка получателя
        SheetRequest.Cells(21, 27) = CStr(SheetCatalog.Cells(NumStrCatalog, _
                                          COL_PAYMENT_ACCOUNT))
        ' БИК Банка получателя
        SheetRequest.Cells(22, 27) = CStr(SheetCatalog.Cells(NumStrCatalog, _
                                          COL_BIK))
        ' Наименование Банка получателя
        SheetRequest.Cells(23, 27) = CStr(SheetCatalog.Cells(NumStrCatalog, _
                                          COL_NAMR_BANK))
        SheetRequest.Cells(23, 27).Rows.RowHeight = SetHeightRow(Len(SheetRequest.Cells(23, 27)))
        ' КБК
        SheetRequest.Cells(24, 27) = CStr(SheetCatalog.Cells(NumStrCatalog, _
                                                             COL_KBK))
        ' ОКТМО
        SheetRequest.Cells(25, 27) = CStr(SheetCatalog.Cells(NumStrCatalog, _
                                          COL_OKTMO))
        ' Налоговый период
        SheetRequest.Cells(26, 27) = CStr(SheetCatalog.Cells(NumStrCatalog, _
                                          COL_TAXABLE_PERIOD))
        ' УИН
        SheetRequest.Cells(27, 27) = CStr(SheetCatalog.Cells(NumStrCatalog, _
                                          COL_UIN))
        ' Ответственный
        SheetRequest.Cells(28, 27) = CStr(SheetRegistr.Cells(NumStrRegistr, _
                                          COL_RESPONSIBLE))
        ' Сохранениу листа NAME_SHEET в книгу
        Dim NumRequest As String
        Dim DateRequest As String
        NumRequest = CStr(SheetRegistr.Cells(NumStrRegistr, COL_NUM_REQUEST))
        DateRequest = CStr(SheetRegistr.Cells(NumStrRegistr, COL_DATE_REQUEST))
        Call MoveRequest(NumRequest, DateRequest, Folder)
    Else
        MsgBox ("Не верно введён номер договора." & vbCr & "Номер договора на листе " & _
                Chr(34) & "Регистрация заявок" & Chr(34) & _
                " из столбца L должен ТОЧНО соответствовать номеру договора из листа " & _
                Chr(34) & "Реестр контрагентов" & Chr(34) & " столбца B")
    End If
            
End Sub


Function SetHeightRow(Len_Row As Integer) As Integer
'
' Выравнивание высоты строки по содержимому
'
    Const MAX_WIDTH = 45
    Const BASE_HEIGHT = 15
    If Len_Row > MAX_WIDTH Then
        SetHeightRow = (Len_Row \ MAX_WIDTH + 1) * BASE_HEIGHT
    Else
        SetHeightRow = BASE_HEIGHT
    End If
End Function


Function GetStringCatalog(NumContract As String) As Integer
'
' Получение номера строки договора из реестра договоров
'
    Set fcell = ThisWorkbook.Worksheets("Реестр контрагентов"). _
                Columns("B:B").Find(CStr(NumContract), _
                LookIn:=xlValues, LookAt:=xlWhole)
    If Not fcell Is Nothing Then
        GetStringCatalog = fcell.Row
    Else
        GetStringCatalog = 0
    End If
    
    If NumContract = "" Then
        GetStringCatalog = 0
    End If
End Function


Private Sub CreateSheetRequest()
    '
    ' Копирование листа "Шаблон заявки" в новый и
    ' переименование его в NAME_SHEET
    '
    ThisWorkbook.Worksheets("Шаблон заявки").Copy After:=Worksheets(Worksheets.count)
    ThisWorkbook.Worksheets("Шаблон заявки (2)").Name = NAME_SHEET
End Sub


Private Sub MoveRequest(NumRequest As String, _
                        DateRequest As String, _
                        Folder As String)
'
' Перемещение листа NAME_SHEET в новую книгу.
'
    Dim FileName As String
    FileName = "Заявка №" + NumRequest + " от " + DateRequest + ".xlsx"
    
    Workbooks.Add
    ActiveWorkbook.SaveAs FileName:=Folder + FileName
    
    ThisWorkbook.Worksheets(NAME_SHEET).Move _
        Before:=Workbooks(FileName).Worksheets("Лист1")
    
    Application.DisplayAlerts = False
    Workbooks(FileName).Worksheets("Лист1").Delete
    Workbooks(FileName).Worksheets("Лист2").Delete
    Workbooks(FileName).Worksheets("Лист3").Delete
    Application.DisplayAlerts = True
    
    Workbooks(FileName).Save
    Workbooks(FileName).Close
End Sub

