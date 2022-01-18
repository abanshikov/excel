Attribute VB_Name = "FileManager"
Public Sub OpenForm()
Attribute OpenForm.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ����� ����� ���������
'

    Dim avFiles, address
        
    '�� ��������� � ������ �������� ����� Excel(xls,xlsx,xlsm,xlsb)
    avFiles = Application.GetOpenFilename _
                ("Excel files(*.xls*),*.xls*", 1, _
                "������� Excel �����", , False)
    If VarType(avFiles) = vbBoolean Then
        '���� ������ ������ ������ - ����� �� ���������
        Cells(1, 2) = "���� �� ������"
        statement_name = ""
        statement_file = ""
        Cells(1, 2).Style = "������"
        Exit Sub
    End If
       
    address = Split(avFiles, "\")
    statement_name = address(UBound(address))
    ReDim Preserve address(UBound(address) - 1)
    statement_path = Join(address, "\") + "\"
    statement_file = avFiles
    
    Cells(1, 2) = "������ ���� " & Chr(34) & statement_name & Chr(34)
    Cells(1, 2).Style = "�������"
    
    Cells(3, 2) = "������ �� ����������"
    Cells(3, 2).Style = "�����������"
    Application.StatusBar = "������ �� ����������"
    
End Sub


Public Sub OpenData()
'
' ����� ����� � �������
'
    Dim avFiles, address
    
    '�� ��������� � ������ �������� ����� Excel(xls,xlsx,xlsm,xlsb)
    avFiles = Application.GetOpenFilename _
                ("Excel files(*.xls*),*.xls*", 1, "������� Excel �����", , False)
    If VarType(avFiles) = vbBoolean Then
        '���� ������ ������ ������ - ����� �� ���������
        Cells(2, 2) = "���� �� ������"
        data_name = ""
        data_file = ""
        Cells(2, 2).Style = "������"
        Exit Sub
    End If
    
    address = Split(avFiles, "\")
    data_name = address(UBound(address))
    data_file = avFiles
    
    Cells(2, 2) = "������ ���� " & Chr(34) & data_name & Chr(34)
    Cells(2, 2).Style = "�������"
    
    Cells(3, 2) = "������ �� ����������"
    Cells(3, 2).Style = "�����������"
    Application.StatusBar = "������ �� ����������"
End Sub

