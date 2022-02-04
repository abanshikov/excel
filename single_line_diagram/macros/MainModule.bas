Attribute VB_Name = "MainModule"
Option Explicit
Public app As nanoCAD.Application
Public ThisDrawing As nanoCAD.Document
Const FONT = "{\fGOST 2.304 type A|b0|i0|c0|p34;"


Sub add_all_parameters()
'
' Управление добавлением данных
'

    ' Массив переменных, содержащих числа двойной точности
    Dim insert_point(2) As Double
    
    ' Хранение передаваемого в NanoCAD текста
    Dim text As String
    Dim txt As AcadMText
    
    ' Размер текста в формате "\H4;" - высота текста 4
    Dim font_size As String


    ' Добавление данных на лист
    Set app = GetObject("", "nanoCAD.Application")
    app.Visible = True
    Set ThisDrawing = app.ActiveDocument

    ' Создадание нового слоя, назначение для него толщины и цвета
    Dim layer As AcadLayer
    Set layer = ThisDrawing.Layers.Add("Автоматические построения")
    layer.Color = 150
    layer.Lineweight = acLnWt020
    ThisDrawing.ActiveLayer = layer


    ' -------------------------------------------------------------------------
    ' Добавление информации о вводном напряжении и мощности

    ' Координаты вставки объекта
    insert_point(0) = 30
    insert_point(1) = 275
    
    ' Добавление текста
    font_size = "\H4;"
    If CStr(Range("B1").Value) <> "" Then
        text = FONT & font_size & "U=" & CStr(Range("B1").Value) & " кВ" & Chr(10)
    End If
    If CStr(Range("B2").Value) <> "" Then
        text = text & "Pуст=" & CStr(Range("B2").Value) & " кВт"
    End If
    ' второй параметр определяет длину Мтекста, третий параметр – текст Мтекста.
    ' "}" - для закрытия свойств шрифта
    Set txt = ThisDrawing.ModelSpace.AddMText(insert_point, 30, CStr(text + "}"))

    ' -------------------------------------------------------------------------
    ' Добавление информации о точке подключениея
    insert_point(0) = 107
    insert_point(1) = 266.5
    font_size = "\H3;"
    If CStr(Range("B3").Value) <> "" Then
        text = FONT & font_size & CStr(Range("B3").Value)
        Set txt = ThisDrawing.ModelSpace.AddMText(insert_point, 37, CStr(text + "}"))
        ' изменение свойства Мтекста "Выравнивание" на "Середина по центру"
        txt.AttachmentPoint = acAttachmentPointMiddleCenter
    End If

    ' -------------------------------------------------------------------------
    ' Добавление информации о вводном кабеле
    insert_point(0) = 123
    insert_point(1) = 248
    font_size = "\H4;"
    If CStr(Range("B4").Value) <> "" Then
        text = FONT & font_size & CStr(Range("B4").Value)
        Set txt = ThisDrawing.ModelSpace.AddMText(insert_point, 37, CStr(text + "}"))
    End If
    
    ' -------------------------------------------------------------------------
    ' Добавление информации о марке электроустановки
    insert_point(0) = 62
    insert_point(1) = 220
    font_size = "\H4;"
    If CStr(Range("B5").Value) <> "" Then
        text = FONT & font_size & CStr(Range("B5").Value)
        Set txt = ThisDrawing.ModelSpace.AddMText(insert_point, 40, CStr(text + "}"))
    End If
   
    
    ' -------------------------------------------------------------------------
    ' Добавление информации о вводном автомате
    insert_point(0) = 112
    insert_point(1) = 210
    font_size = "\H4;"
    If CStr(Range("B6").Value) <> "" Then
        text = FONT & font_size & CStr(Range("B6").Value)
        Set txt = ThisDrawing.ModelSpace.AddMText(insert_point, 40, CStr(text + "}"))
    End If

    ' -------------------------------------------------------------------------
    ' Добавление информации о приборе учёта
    insert_point(0) = 115
    insert_point(1) = 188
    font_size = "\H4;"
    If CStr(Range("B7").Value) <> "" Then
        text = FONT & font_size & CStr(Range("B7").Value)
        Set txt = ThisDrawing.ModelSpace.AddMText(insert_point, 40, CStr(text + "}"))
    End If

    ' -------------------------------------------------------------------------
    ' Добавление информации о шине заземления
    insert_point(0) = 25
    insert_point(1) = 175
    font_size = "\H3;"
    If CStr(Range("B8").Value) <> "" Then
        text = FONT & font_size & CStr(Range("B8").Value)
        Set txt = ThisDrawing.ModelSpace.AddMText(insert_point, 30, CStr(text + "}"))
    End If

    ' -------------------------------------------------------------------------
    ' Добавление информации о заземлителях
    insert_point(0) = 71
    insert_point(1) = 76
    font_size = "\H3;"
    If CStr(Range("B9").Value) <> "" Then
        text = FONT & font_size & CStr(Range("B9").Value)
        Set txt = ThisDrawing.ModelSpace.AddMText(insert_point, 60, CStr(text + "}"))
    End If

    ' -------------------------------------------------------------------------
    ' Добавление информации о типе установки и инвентарном номере
    insert_point(0) = 139.023
    insert_point(1) = 43.605
    font_size = "\H3;"
    If CStr(Range("B10").Value) <> "" Then
        text = FONT & font_size & CStr(Range("B10").Value)
        Set txt = ThisDrawing.ModelSpace.AddMText(insert_point, 110, CStr(text + "}"))
        txt.AttachmentPoint = acAttachmentPointMiddleCenter
    End If

    ' -------------------------------------------------------------------------
    ' Добавление информации об адресе
    insert_point(0) = 113.373
    insert_point(1) = 28.215
    font_size = "\H3;"
    If CStr(Range("B11").Value) <> "" Then
        text = FONT & font_size & CStr(Range("B11").Value)
        Set txt = ThisDrawing.ModelSpace.AddMText(insert_point, 60, CStr(text + "}"))
        txt.AttachmentPoint = acAttachmentPointMiddleCenter
    End If
    
End Sub


