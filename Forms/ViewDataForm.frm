VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ViewDataForm 
   Caption         =   "Управление видами"
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10845
   OleObjectBlob   =   "ViewDataForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ViewDataForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Dim ViewData() As ViewType: ReDim ViewData(2000) 'Создание динамического массива
    'ViewData.StyleName
'    With ViewData(0)
'        .StyleName = TextBox1.Text
'        .FontSize = TextBox2.Text
'        .InsertSymbol_FontName = TextBox3.Text
'        .InsertSymbol_CharacterNumber1 = TextBox4.Text
'        .InsertSymbol_CharacterNumber2 = TextBox5.Text
'        .TegText = TextBox6.Text
'        .ClearTime = CheckBox1.Value
'        .SaveDoc = CheckBox2.Value
'    End With
'DataBoolToList(Var) ' обратная ListToDataBool
    With ListBox1
        .AddItem TextBox1.Text
        Count = .ListCount - 1
        .List(Count, 1) = TextBox2.Text
        .List(Count, 2) = TextBox3.Text
        .List(Count, 3) = TextBox4.Text
        .List(Count, 4) = TextBox5.Text
        .List(Count, 5) = TextBox6.Text
        .List(Count, 6) = DataBoolToList(CheckBox1.Value)
        .List(Count, 7) = DataBoolToList(CheckBox2.Value)
    End With
End Sub

Private Function CheckCheckBoxValue(CheckBoxValue)
    If CheckBoxValue = False Then
        CheckCheckBoxValue = False
    End If
    If CheckBoxValue = True Or IsNull(CheckBoxValue) Then
        CheckCheckBoxValue = True
    End If
End Function
Private Sub CommandButton2_Click()
    If ListBox1.ListIndex = -1 Then Exit Sub
    'DataBoolToList(Var) ' обратная ListToDataBool
    ListItemIndex = ListBox1.ListIndex
    With ListBox1
        '.List(ListItemIndex, 0) = TextBox1.Text
        .List(ListItemIndex, 1) = TextBox2.Text
        .List(ListItemIndex, 2) = TextBox3.Text
        .List(ListItemIndex, 3) = TextBox4.Text
        .List(ListItemIndex, 4) = TextBox5.Text
        .List(ListItemIndex, 5) = TextBox6.Text
        .List(ListItemIndex, 6) = DataBoolToList(CheckBox1.Value)
        .List(ListItemIndex, 7) = DataBoolToList(CheckBox2.Value)
        .List(ListItemIndex, 0) = TextBox1.Text 'Не понятно почему, но оно автоматически вызывает событие ListBox1_Click
    End With
End Sub

Private Sub CommandButton3_Click()
    If ListBox1.ListIndex = -1 Then Exit Sub
    If MsgBox("Удалить выбранный стиль: " & vbNewLine & _
    ListBox1.List(ListBox1.ListIndex, 0) & ", " & _
    ListBox1.List(ListBox1.ListIndex, 1) & ", " & _
    ListBox1.List(ListBox1.ListIndex, 2) & ", " & _
    ListBox1.List(ListBox1.ListIndex, 3) & ", " & _
    ListBox1.List(ListBox1.ListIndex, 4) & ", " & _
    ListBox1.List(ListBox1.ListIndex, 5) & ", " & _
    ListBox1.List(ListBox1.ListIndex, 6) & ", " & _
    ListBox1.List(ListBox1.ListIndex, 7) & "" _
    , vbYesNo, "Удаление элемента") = vbYes Then
        ListBox1.RemoveItem (ListBox1.ListIndex)
    End If
End Sub

Private Sub CommandButton4_Click()
    Dim ViewData() As ViewType
    ReDim Preserve ViewData(ListBox1.ListCount - 1)
    For i = 0 To UBound(ViewData)
        With ListBox1
            ViewData(i).StyleName = .List(i, 0)
            ViewData(i).FontSize = .List(i, 1)
            ViewData(i).InsertSymbol_FontName = .List(i, 2)
            ViewData(i).InsertSymbol_CharacterNumber1 = .List(i, 3)
            ViewData(i).InsertSymbol_CharacterNumber2 = .List(i, 4)
            ViewData(i).TegText = .List(i, 5)
            ViewData(i).ClearTime = ListToDataBool(.List(i, 6))
            ViewData(i).SaveDoc = ListToDataBool(.List(i, 7))
        End With
    Next i
    'Сохранение
    Path = ActiveDocument.Path & "\" & ViewDataFileName & ".bin"
    Open Path For Binary As #1
    Put #1, 1, ViewData 'Сохранение данных в файл
    Close #1
End Sub

Private Sub CommandButton5_Click()
    If MsgBox("Восстановить значения по умолчанию?", vbYesNo, "Значения по умолчанию") = vbNo Then
        Exit Sub
    End If
    Dim ViewData() As ViewType
    ReDim Preserve ViewData(0)
    With ViewData(UBound(ViewData))
    'ВидПрименить "Воспоминание", 22, "Webdings", -4003, "Воспоминание", False, True
        .StyleName = "Воспоминание"
        .FontSize = 22
        .InsertSymbol_FontName = "Webdings"
        .InsertSymbol_CharacterNumber1 = -4003
        .InsertSymbol_CharacterNumber2 = 0
        .TegText = "Воспоминание"
        .ClearTime = False
        .SaveDoc = True
    End With
    
    ReDim Preserve ViewData(UBound(ViewData) + 1)
    With ViewData(UBound(ViewData))
    'ВидПрименить "Идея", 26, "Wingdings", -4033, "Идея", False, True
        .StyleName = "Идея"
        .FontSize = 26
        .InsertSymbol_FontName = "Wingdings"
        .InsertSymbol_CharacterNumber1 = -4033
        .InsertSymbol_CharacterNumber2 = 0
        .TegText = "Идея"
        .ClearTime = False
        .SaveDoc = True
    End With
    
    ReDim Preserve ViewData(UBound(ViewData) + 1)
    With ViewData(UBound(ViewData))
    'ВидПрименить "Особый", 26, "Segoe UI Emoji", -10180, "Гриб", False, True, -8380
        .StyleName = "Особый"
        .FontSize = 26
        .InsertSymbol_FontName = "Segoe UI Emoji"
        .InsertSymbol_CharacterNumber1 = -10180
        .InsertSymbol_CharacterNumber2 = -8380
        .TegText = "Гриб"
        .ClearTime = False
        .SaveDoc = True
    End With
    
    ReDim Preserve ViewData(UBound(ViewData) + 1)
    With ViewData(UBound(ViewData))
    'ВидПрименить "Идея", 26, "Wingdings 2", -4062, "Событие", False, True
        .StyleName = "Идея"
        .FontSize = 26
        .InsertSymbol_FontName = "Wingdings 2"
        .InsertSymbol_CharacterNumber1 = -4062
        .InsertSymbol_CharacterNumber2 = 0
        .TegText = "Событие"
        .ClearTime = False
        .SaveDoc = True
    End With
    
    'Заполнение таблицы
    ListBox1.Clear
    For i = 0 To UBound(ViewData)
        With ListBox1
            .AddItem ViewData(i).StyleName
            .List(i, 0) = ViewData(i).StyleName
            .List(i, 1) = ViewData(i).FontSize
            .List(i, 2) = ViewData(i).InsertSymbol_FontName
            .List(i, 3) = ViewData(i).InsertSymbol_CharacterNumber1
            .List(i, 4) = ViewData(i).InsertSymbol_CharacterNumber2
            .List(i, 5) = ViewData(i).TegText
            .List(i, 6) = DataBoolToList(ViewData(i).ClearTime)
            .List(i, 7) = DataBoolToList(ViewData(i).SaveDoc)
        End With
    Next i
End Sub

Private Function DataBoolToList(Var) ' обратная ListToDataBool
    If Var = True Then
        DataBoolToList = "+"
    Else
        DataBoolToList = "-"
    End If
End Function

Private Function ListToDataBool(Var)
    If Var = "+" Then
        ListToDataBool = True
    Else
        ListToDataBool = False
    End If
End Function

Private Sub CommandButton6_Click()
    ind = ListBox1.ListIndex
    StyleName = ListBox1.List(ind, 0)
    FontSize = ListBox1.List(ind, 1)
    InsertSymbol_FontName = ListBox1.List(ind, 2)
    InsertSymbol_CharacterNumber1 = ListBox1.List(ind, 3)
    InsertSymbol_CharacterNumber2 = ListBox1.List(ind, 4)
    TegText = ListBox1.List(ind, 5)
    ClearTime = ListToDataBool(ListBox1.List(ind, 6))
    SaveDoc = ListToDataBool(ListBox1.List(ind, 7))
    ВидПрименить StyleName, FontSize, InsertSymbol_FontName, InsertSymbol_CharacterNumber1, TegText, ClearTime, SaveDoc, InsertSymbol_CharacterNumber2
End Sub

Private Sub ListBox1_Click()
'DataBoolToList(Var) ' обратная ListToDataBool
    If ListBox1.ListCount > 1 Then
        TextBox1.Text = ListBox1.List(ListBox1.ListIndex, 0)
        TextBox2.Text = ListBox1.List(ListBox1.ListIndex, 1)
        TextBox3.Text = ListBox1.List(ListBox1.ListIndex, 2)
        TextBox4.Text = ListBox1.List(ListBox1.ListIndex, 3)
        TextBox5.Text = ListBox1.List(ListBox1.ListIndex, 4)
        TextBox6.Text = ListBox1.List(ListBox1.ListIndex, 5)
        CheckBox1.Value = ListToDataBool(ListBox1.List(ListBox1.ListIndex, 6))
        CheckBox2.Value = ListToDataBool(ListBox1.List(ListBox1.ListIndex, 7))
    End If
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    Path = ActiveDocument.Path & "\" & ViewDataFileName & ".bin" 'ViewDataFileName из основного модуля
    If Not СуществуетФайл(Path) Then
        MsgBox "Не найден файл видов '" & ViewDataFileName & "' в папке" & vbNewLine & ActiveDocument.Path
        Exit Sub
    End If
    Dim ViewData() As ViewType
    ReDim Preserve ViewData(100)
    Close #1
    Open Path For Binary As #1
    Get #1, 1, ViewData
    Close #1
    'Заполнение таблицы
    ListBox1.Clear
    'For i = 0 To UBound(ViewData)
    'Next i
    i = 0
    Do While ViewData(i).StyleName <> "" And ViewData(i + 1).StyleName <> ""
        With ListBox1
            .AddItem ViewData(i).StyleName
            .List(i, 0) = ViewData(i).StyleName
            .List(i, 1) = ViewData(i).FontSize
            .List(i, 2) = ViewData(i).InsertSymbol_FontName
            .List(i, 3) = ViewData(i).InsertSymbol_CharacterNumber1
            .List(i, 4) = ViewData(i).InsertSymbol_CharacterNumber2
            .List(i, 5) = ViewData(i).TegText
            .List(i, 6) = DataBoolToList(ViewData(i).ClearTime)
            .List(i, 7) = DataBoolToList(ViewData(i).SaveDoc)
        End With
        i = i + 1
    Loop
    With ListBox1
        .AddItem ViewData(i).StyleName
        .List(i, 0) = ViewData(i).StyleName
        .List(i, 1) = ViewData(i).FontSize
        .List(i, 2) = ViewData(i).InsertSymbol_FontName
        .List(i, 3) = ViewData(i).InsertSymbol_CharacterNumber1
        .List(i, 4) = ViewData(i).InsertSymbol_CharacterNumber2
        .List(i, 5) = ViewData(i).TegText
        .List(i, 6) = DataBoolToList(ViewData(i).ClearTime)
        .List(i, 7) = DataBoolToList(ViewData(i).SaveDoc)
    End With
End Sub
