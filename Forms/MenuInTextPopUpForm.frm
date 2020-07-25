VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MenuInTextPopUpForm 
   Caption         =   "Дополнительное меню правого клика на тексте"
   ClientHeight    =   9990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11670
   OleObjectBlob   =   "MenuInTextPopUpForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MenuInTextPopUpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ViewData() As ViewType
Dim OtherViewDic
Const Mname As String = "MyPopUpMenu"
Const ViewPopUpDataMaxSize = 100

'CommandBarName     - Название CommandBar'а
'Standard-вкладка, Text-в меню при ПКМ на тексте, _
 Header Area Popup-верхний колонтитул, Headings-заголовки
    
Private Sub AddOtherView() 'Добавление прочих (особых) видов
    Set OtherViewDic = CreateObject("Scripting.Dictionary") ' создаём словарь
    OtherViewDic.Add "Мысль", "ВидМысль"
    OtherViewDic.Add "Музыка", "ВидМузыка"
    OtherViewDic.Add "Цитата", "ВидЦитата"
    OtherViewDic.Add "Рамка с засечками", "ВидРамкаСЗасечками"
    OtherViewDic.Add "Вопрос-Ответ", "ВидВопросОтвет"
    OtherViewDic.Add "Записка", "ВидЗаписка"
    OtherViewDic.Add "Дополнение", "ВидДополнение"
    OtherViewDic.Add "Обычный", "ВидОбычный"
    OtherViewDic.Add "Фаза", "ВидФаза"
    OtherViewDic.Add "Код", "ВидКод"
    OtherViewDic.Add "Список книг", "ВидСписокКниг"
End Sub

Private Function GetCmdStringFromViewData(i)
    If i <= UBound(ViewData) Then
        S = ViewData(i).StyleName
        S = S & ", " & ViewData(i).FontSize
        S = S & ", " & ViewData(i).InsertSymbol_FontName
        S = S & ", " & ViewData(i).InsertSymbol_CharacterNumber1
        S = S & ", " & ViewData(i).InsertSymbol_CharacterNumber2
        S = S & ", " & ViewData(i).TegText
        S = S & ", " & DataBoolToList(ViewData(i).ClearTime) 'DataBoolToList в модуле WordDiaryMacros
        S = S & ", " & DataBoolToList(ViewData(i).SaveDoc)
        GetCmdStringFromViewData = S
    Else
        GetCmdStringFromViewData = ""
    End If
End Function
    
Private Sub CommandButton1_Click()
    If Trim(TextBox1.Text) = "" Then
        MsgBox "Введите название"
        TextBox1.SetFocus
        Exit Sub
    End If
    If Trim(TextBox2.Text) = "" Then
        MsgBox "Введите номер иконки"
        TextBox2.SetFocus
        Exit Sub
    End If
    If ListBox1.ListIndex >= 0 Then 'Добавление пользовательских
        With ListBox3
            .AddItem TextBox1.Text
            Count = .ListCount - 1
            .List(Count, 1) = ListBox1.List(ListBox1.ListIndex)
            .List(Count, 2) = TextBox2.Text
            .List(Count, 3) = GetCmdStringFromViewData(ListBox1.ListIndex)
        End With
    End If
    If ListBox2.ListIndex >= 0 Then 'Добавление особых
        With ListBox3
            .AddItem TextBox1.Text
            Count = .ListCount - 1
            .List(Count, 1) = ListBox2.List(ListBox2.ListIndex)
            .List(Count, 2) = TextBox2.Text
            .List(Count, 3) = "Код: " & OtherViewDic.Items()(ListBox2.ListIndex)
        End With
    End If
End Sub

Private Sub CommandButton10_Click()
    Call DeletePopUpMenu
    Call ShowFaceIdIcons_PopUpMenu(3001, 3500)
    On Error Resume Next
    Application.CommandBars(Mname).ShowPopUp
    On Error GoTo 0
End Sub

Private Sub CommandButton11_Click()
    Call DeletePopUpMenu
    Call ShowFaceIdIcons_PopUpMenu(3501, 4000)
    On Error Resume Next
    Application.CommandBars(Mname).ShowPopUp
    On Error GoTo 0
End Sub

Private Sub CommandButton12_Click()
    Call DeletePopUpMenu
    Call ShowFaceIdIcons_PopUpMenu(4001, 4500)
    On Error Resume Next
    Application.CommandBars(Mname).ShowPopUp
    On Error GoTo 0
End Sub

Private Sub CommandButton13_Click()
    Call DeletePopUpMenu
    Call ShowFaceIdIcons_PopUpMenu(5401, 5500)
    On Error Resume Next
    Application.CommandBars(Mname).ShowPopUp
    On Error GoTo 0
End Sub

Private Sub CommandButton14_Click()
    Call DeletePopUpMenu
    Call ShowFaceIdIcons_PopUpMenu(5501, 5685)
    On Error Resume Next
    Application.CommandBars(Mname).ShowPopUp
    On Error GoTo 0
End Sub

Private Sub CommandButton15_Click()
    If Trim(TextBox3.Text) <> "" Then
        index = Int(TextBox3.Text)
        Call DeletePopUpMenu
        Call ShowFaceIdIcons_PopUpMenu(index, index)
        On Error Resume Next
        Application.CommandBars(Mname).ShowPopUp
        On Error GoTo 0
    End If
End Sub

Private Sub AddCommandBarPopUpMenu(ByVal CommandBarName, ByVal CommandBarTeg, ByVal CommandBarFuncName, ByVal CommandBarCaption)
'Одно меню с выпадающей подменюшкой в виде массива
    Dim ContextMenu As CommandBar
    Dim ModuleName, FuncName As String
    ModuleName = "WordDiaryMacros" ' название модуля из которого будет вызвана функция
    FuncName = CommandBarFuncName '"ВидПрименитьИзМеню" ' Имя Вызываемой функции
    DeleteFromCommandBarMenu CommandBarName, CommandBarTeg  'Удаление Вида из меню
    
    Set ContextMenu = Application.CommandBars(CommandBarName)
    Set MenuItem = ContextMenu.Controls.Add(Type:=msoControlPopup, Before:=1)
    With MenuItem
        .Tag = CommandBarTeg
        .Caption = CommandBarCaption '"Вид" Заголовок показываемого подменю в меню ПКМ
        
        If CommandBarTeg = ViewPopUpMenuTegText Then ' Если тег элемента для Вида
            For i = 0 To ListBox3.ListCount - 1
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = ListBox3.List(i, 0) 'Название
                    .FaceID = ListBox3.List(i, 2) 'Иконка
                    .Parameter = ListBox3.List(i, 3) 'Строка команд
                    .onaction = ModuleName & "." & FuncName
                End With
            Next i
        ElseIf CommandBarTeg = TegsPopUpMenuTegText Then ' Если тег элемента для Тега
            For i = 0 To ListBox4.ListCount - 1
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = ListBox4.List(i, 0) 'Название
                    .FaceID = ListBox4.List(i, 2) 'Иконка
                    .Parameter = ListBox4.List(i, 3) 'Строка команд (Название)
                    .onaction = ModuleName & "." & FuncName
                End With
            Next i
        End If
    End With
End Sub

Private Sub CommandButton16_Click()
    ApplyPopUpOptions ' Приминение настроек ПКМ
    SavePopUpItemsToFile ' Сохранение данных выводимых при ПКМ
    SavePopUpOptionsToFile ' Сохранение настроек вывода при ПКМ
End Sub

Private Sub ApplyPopUpOptions()
    ShowViewMenu = CheckBox1.Value ' Показывать меню "Вид" (при ПКМ на обычном тексте)
    ShowTegMenuOnText = CheckBox2.Value ' Показывать меню "Тег" (при ПКМ на обычном тексте)
    ShowTegMenuOnHandle = CheckBox3.Value ' Показывать меню "Тег" (при ПКМ на заголовке)
    
    If ShowViewMenu Then ' Показывать меню "Вид" (при ПКМ на обычном тексте)
        If ListBox3.ListCount > 0 Then
            AddCommandBarPopUpMenu "Text", ViewPopUpMenuTegText, "ВидПрименитьИзМеню", "Вид"
        End If
    Else
        DeleteFromCommandBarMenu "Text", ViewPopUpMenuTegText
    End If
    
    If ShowTegMenuOnText Then ' Показывать меню "Тег" (при ПКМ на обычном тексте)
        If ListBox4.ListCount > 0 Then
            AddCommandBarPopUpMenu "Text", TegsPopUpMenuTegText, "ТегИзТекстаПрименитьИзМеню", "Тег"
        End If
    Else
        DeleteFromCommandBarMenu "Text", TegsPopUpMenuTegText
    End If
    
    If ShowTegMenuOnHandle Then ' Показывать меню "Тег" (при ПКМ на заголовке)
        If ListBox4.ListCount > 0 Then
            AddCommandBarPopUpMenu "Headings", TegsPopUpMenuTegText, "ТегИзЗаголовкаПрименитьИзМеню", "Тег"
        End If
    Else
        DeleteFromCommandBarMenu "Headings", TegsPopUpMenuTegText
    End If
    
End Sub

Private Sub CommandButton17_Click()
    If ListBox3.ListIndex = -1 Then Exit Sub
    'DataBoolToList(Var) ' обратная ListToDataBool
    ListItemIndex = ListBox3.ListIndex
    With ListBox3
        '.List(ListItemIndex, 0) = TextBox1.Text
        .List(ListItemIndex, 2) = TextBox2.Text
        .List(ListItemIndex, 0) = TextBox1.Text 'Не понятно почему, но оно автоматически вызывает событие ListBox1_Click
    End With
End Sub

Private Sub CommandButton18_Click()
    If ListBox3.ListIndex = -1 Then Exit Sub
    MsgBox ("Подробный вид выбранного элемента: " & vbNewLine & _
    ListBox3.List(ListBox3.ListIndex, 0) & ", " & _
    ListBox3.List(ListBox3.ListIndex, 1) & ", " & _
    ListBox3.List(ListBox3.ListIndex, 2) & ", " & _
    ListBox3.List(ListBox3.ListIndex, 3) & "")
End Sub

Private Sub CommandButton19_Click()
    If MsgBox("Удалить все элементы из меню ПКМ? (" & CStr(ListBox3.ListCount) & ")" & vbNewLine & _
    "Если да, то после надо будет вручную сохранить изменения, нажав кнопку 'Применить'", vbYesNo, "Полное очищение") = vbYes Then
        ListBox3.Clear
    End If
End Sub

Private Sub CommandButton2_Click()
    If ListBox3.ListIndex = -1 Then Exit Sub
    If MsgBox("Удалить выбранный пункт меню: " & vbNewLine & _
    ListBox3.List(ListBox3.ListIndex, 0) & ", " & _
    ListBox3.List(ListBox3.ListIndex, 1) & ", " & _
    ListBox3.List(ListBox3.ListIndex, 2) & ", " & _
    ListBox3.List(ListBox3.ListIndex, 3) & "" _
    , vbYesNo, "Удаление элемента") = vbYes Then
        ListBox3.RemoveItem (ListBox3.ListIndex)
    End If
End Sub

Private Sub CommandButton20_Click()
    SavePopUpItemsToFile
End Sub

Private Sub CommandButton21_Click()
    LoadPopUpItemsFromFile
End Sub

Private Sub CommandButton22_Click()
    If ListBox3.ListCount >= 0 Then
        For i = 0 To ListBox3.ListCount - 1
            КурсорВКонец
            Selection.TypeParagraph
            S = ListBox3.List(i, 0) & vbTab & _
            ListBox3.List(i, 1) & vbTab & _
            ListBox3.List(i, 2) & vbTab & _
            ListBox3.List(i, 3)
            Selection.Text = S
        Next i
        КурсорВКонец
        Selection.TypeParagraph
    End If
End Sub

Private Sub CommandButton23_Click()
    If ListBox3.ListIndex = -1 Then Exit Sub
    CmdStrPast = ListBox3.List(ListBox3.ListIndex, 3)
    CmdStr = InputBox("Измениете кода. Будьте аккуратны и ничего не изменяйте, если не понимаете, что делаете.", "Изменение кода", CmdStrPast)
    If CmdStr <> "" Then
        ListBox3.List(ListBox3.ListIndex, 3) = CmdStr
    End If
End Sub

Private Sub CommandButton24_Click()
    ViewDataForm.Show
End Sub

Private Sub CommandButton25_Click()
    If ListBox4.ListIndex = -1 Then Exit Sub
    If MsgBox("Удалить выбранный пункт: " & ListBox4.List(ListBox4.ListIndex), vbYesNo, "Удаление элемента") = vbYes Then
        ListBox4.RemoveItem (ListBox4.ListIndex)
    End If
End Sub

Private Sub CommandButton26_Click()
    If ListBox4.ListCount = 0 Then
        Exit Sub
    End If
    Dim TegsPopUpData() As ViewPopUpType
    ReDim Preserve TegsPopUpData(ListBox4.ListCount - 1)
    For i = 0 To ListBox4.ListCount - 1
        With ListBox4
            TegsPopUpData(i).Name = .List(i, 0)
            TegsPopUpData(i).Teg = .List(i, 1)
            TegsPopUpData(i).FaceID = .List(i, 2)
            TegsPopUpData(i).Code = .List(i, 3)
        End With
    Next i
    'Сохранение
    'ViewPopUpDataMaxSize
    Path = ActiveDocument.Path & "\" & TegsDataFileName 'TegsDataFileName из основного модуля. Для сохранения тегов показываемых в меню ПКМ
    Open Path For Binary As #1
    Put #1, 1, TegsPopUpData 'Сохранение данных в файл
    Close #1
End Sub

Private Sub CommandButton27_Click()
    TegLoadToForm
End Sub

Private Sub CommandButton28_Click()
    If Trim(TextBox4.Text) = "" Then
        MsgBox "Введите название"
        TextBox4.SetFocus
        Exit Sub
    End If
    If Trim(TextBox5.Text) = "" Then
        MsgBox "Введите номер иконки"
        TextBox5.SetFocus
        Exit Sub
    End If
    If ListBox5.ListIndex >= 0 Then 'Добавление
        With ListBox4
            .AddItem TextBox4.Text
            Count = .ListCount - 1
            .List(Count, 1) = ListBox5.List(ListBox5.ListIndex)
            .List(Count, 2) = TextBox5.Text
            .List(Count, 3) = ListBox5.List(ListBox5.ListIndex) & ", " & CheckBox4.Value
        End With
    End If
End Sub

Private Sub CommandButton3_Click()
'5501-5685
'5401-5500
'4001-4500
'3501-4000
'3001-3500
'2501-3000
'2001-2500
'1501-2000
'1001-1500
'501-1000
'1-500
    Call DeletePopUpMenu
    Call Help_PopUpMenu
    On Error Resume Next
    Application.CommandBars(Mname).ShowPopUp
    On Error GoTo 0
End Sub

Public Sub ShowFaceIdIcons_PopUpMenu(ByVal IconStart As Integer, ByVal IconEnd As Integer)
    Dim MenuItem As CommandBarPopup
    ' Add the popup menu. ModuleName & "." & FuncName
    Dim ModuleName, FuncName As String
    ModuleName = "MenuInTextPopUpForm"
    FuncName = "CopyTextFromHelpPopUp"
    With Application.CommandBars.Add(Name:=Mname, Position:=msoBarPopup, _
        MenuBar:=False, Temporary:=True)
        For i = IconStart To IconEnd
            With .Controls.Add(Type:=msoControlButton)
            .Caption = Str(i)
            .FaceID = i
            .Parameter = i
            .onaction = ""
            End With
        Next i

'        Set MenuItem = .Controls.Add(Type:=msoControlPopup)
'        With MenuItem
'            .caption = IconStart & "-" & IconEnd
'            For i = IconStart To IconEnd
'                With .Controls.Add(Type:=msoControlButton)
'                    .caption = Str(i)
'                    .FaceId = i
'                    .Parameter = i
'                    .onaction = ""
'                End With
'            Next i
'        End With
    End With
End Sub

Sub Help_PopUpMenu()
    Dim MenuItem As CommandBarPopup
    ' Add the popup menu. ModuleName & "." & FuncName
    Dim ModuleName, FuncName As String
    ModuleName = "MenuInTextPopUpForm"
    FuncName = "CopyTextFromHelpPopUp"
    With Application.CommandBars.Add(Name:=Mname, Position:=msoBarPopup, _
        MenuBar:=False, Temporary:=True)

        Set MenuItem = .Controls.Add(Type:=msoControlPopup)
        With MenuItem
            .Caption = "1-500"
            For i = 1 To 500
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = Str(i)
                    .FaceID = i
                    .Parameter = i
                    .onaction = ""
                End With
            Next i
        End With
        
        Set MenuItem = .Controls.Add(Type:=msoControlPopup)
        With MenuItem
            .Caption = "501-1000"
            For i = 501 To 1000
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = Str(i)
                    .FaceID = i
                    .Parameter = i
                    .onaction = ""
                End With
            Next i
        End With

        Set MenuItem = .Controls.Add(Type:=msoControlPopup)
        With MenuItem
            .Caption = "1001-1500"
            For i = 1001 To 1500
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = Str(i)
                    .FaceID = i
                    .Parameter = i
                    .onaction = ""
                End With
            Next i
        End With
        
        Set MenuItem = .Controls.Add(Type:=msoControlPopup)
        With MenuItem
            .Caption = "1501-2000"
            For i = 1501 To 2000
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = Str(i)
                    .FaceID = i
                    .Parameter = i
                    .onaction = ""
                End With
            Next i
        End With
        
        Set MenuItem = .Controls.Add(Type:=msoControlPopup)
        With MenuItem
            .Caption = "2001-2500"
            For i = 2001 To 2500
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = Str(i)
                    .FaceID = i
                    .Parameter = i
                    .onaction = ""
                End With
            Next i
        End With
        
        Set MenuItem = .Controls.Add(Type:=msoControlPopup)
        With MenuItem
            .Caption = "2501-3000"
            For i = 2501 To 3000
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = Str(i)
                    .FaceID = i
                    .Parameter = i
                    .onaction = ""
                End With
            Next i
        End With
        
        Set MenuItem = .Controls.Add(Type:=msoControlPopup)
        With MenuItem
            .Caption = "3001-3500"
            For i = 3001 To 3500
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = Str(i)
                    .FaceID = i
                    .Parameter = i
                    .onaction = ""
                End With
            Next i
        End With
        
        Set MenuItem = .Controls.Add(Type:=msoControlPopup)
        With MenuItem
            .Caption = "3501-4000"
            For i = 3501 To 4000
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = Str(i)
                    .FaceID = i
                    .Parameter = i
                    .onaction = ""
                End With
            Next i
        End With
         
        Set MenuItem = .Controls.Add(Type:=msoControlPopup)
        With MenuItem
            .Caption = "4001-4500"
            For i = 4001 To 4500
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = Str(i)
                    .FaceID = i
                    .Parameter = i
                    .onaction = ""
                End With
            Next i
        End With
        
        Set MenuItem = .Controls.Add(Type:=msoControlPopup)
        With MenuItem
            .Caption = "5401-5500"
            For i = 5401 To 5500
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = Str(i)
                    .FaceID = i
                    .Parameter = i
                    .onaction = ""
                End With
            Next i
        End With
        
        Set MenuItem = .Controls.Add(Type:=msoControlPopup)
        With MenuItem
            .Caption = "5501-5685"
            For i = 5501 To 5685
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = Str(i)
                    .FaceID = i
                    .Parameter = i
                    .onaction = ""
                End With
            Next i
        End With
    End With
End Sub

Sub DeletePopUpMenu()
    ' Delete the popup menu if it already exists.
    On Error Resume Next
    Application.CommandBars(Mname).Delete
    On Error GoTo 0
End Sub

Public Sub CreateDisplayPopUpMenu()

End Sub

Private Sub CommandButton31_Click()
    If MsgBox("Удалить все элементы из меню Тегов ПКМ? (" & CStr(ListBox4.ListCount) & ")" & vbNewLine & _
    "Если да, то после надо будет вручную сохранить изменения, нажав кнопку 'Применить'", vbYesNo, "Полное очищение") = vbYes Then
        ListBox4.Clear
    End If
End Sub

Private Sub CommandButton33_Click()
    If ListBox4.ListCount >= 0 Then
        For i = 0 To ListBox4.ListCount - 1
            КурсорВКонец
            Selection.TypeParagraph
            S = ListBox4.List(i, 0) & vbTab & _
            ListBox4.List(i, 1) & vbTab & _
            ListBox4.List(i, 2) & vbTab & _
            ListBox4.List(i, 3) & vbTab
            Selection.Text = S
        Next i
        КурсорВКонец
        Selection.TypeParagraph
    End If
End Sub

Private Sub CommandButton34_Click()
    TegsForm.Show
End Sub

Private Sub CommandButton35_Click()
    If ListBox4.ListIndex = -1 Then Exit Sub
    'DataBoolToList(Var) ' обратная ListToDataBool
    ListItemIndex = ListBox4.ListIndex
    With ListBox4
        '.List(ListItemIndex, 0) = TextBox1.Text
        .List(ListItemIndex, 2) = TextBox5.Text
        .List(ListItemIndex, 3) = TextBox4.Text & ", " & CheckBox4.Value
        .List(ListItemIndex, 0) = TextBox4.Text 'Не понятно почему, но оно автоматически вызывает событие ListBox1_Click
    End With
End Sub

Private Sub CommandButton36_Click()
    If ListBox4.ListIndex = -1 Then Exit Sub
    CmdStrPast = ListBox4.List(ListBox4.ListIndex, 3)
    CmdStr = InputBox("Изменение кода. Будьте аккуратны и ничего не изменяйте, если не понимаете, что делаете.", "Изменение кода", CmdStrPast)
    If CmdStr <> "" Then
        ListBox4.List(ListBox4.ListIndex, 3) = CmdStr
    End If
End Sub

Private Sub CommandButton38_Click()
    If ListBox4.ListIndex = -1 Then Exit Sub
    MsgBox ("Подробный вид выбранного элемента: " & vbNewLine & _
    ListBox4.List(ListBox4.ListIndex, 0) & ", " & _
    ListBox4.List(ListBox4.ListIndex, 1) & ", " & _
    ListBox4.List(ListBox4.ListIndex, 2) & ", " & _
    ListBox4.List(ListBox4.ListIndex, 3) & "")
End Sub

Private Sub CommandButton4_Click()
    Call DeletePopUpMenu
    Call ShowFaceIdIcons_PopUpMenu(1, 500)
    On Error Resume Next
    Application.CommandBars(Mname).ShowPopUp
    On Error GoTo 0
End Sub

Private Sub CommandButton5_Click()
    Call DeletePopUpMenu
    Call ShowFaceIdIcons_PopUpMenu(501, 1000)
    On Error Resume Next
    Application.CommandBars(Mname).ShowPopUp
    On Error GoTo 0
End Sub

Private Sub CommandButton6_Click()
    Call DeletePopUpMenu
    Call ShowFaceIdIcons_PopUpMenu(1001, 1500)
    On Error Resume Next
    Application.CommandBars(Mname).ShowPopUp
    On Error GoTo 0
End Sub

Private Sub CommandButton7_Click()
    Call DeletePopUpMenu
    Call ShowFaceIdIcons_PopUpMenu(1501, 2500)
    On Error Resume Next
    Application.CommandBars(Mname).ShowPopUp
    On Error GoTo 0
End Sub

Private Sub CommandButton8_Click()
    Call DeletePopUpMenu
    Call ShowFaceIdIcons_PopUpMenu(2001, 2500)
    On Error Resume Next
    Application.CommandBars(Mname).ShowPopUp
    On Error GoTo 0
End Sub

Private Sub CommandButton9_Click()
    Call DeletePopUpMenu
    Call ShowFaceIdIcons_PopUpMenu(2501, 3000)
    On Error Resume Next
    Application.CommandBars(Mname).ShowPopUp
    On Error GoTo 0
End Sub

Private Sub ListBox1_Click()
    ListBox2.ListIndex = -1
    If ListBox1.ListCount >= 0 Then
        TextBox1.Text = ListBox1.List(ListBox1.ListIndex)
    End If
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    CommandButton1_Click
End Sub

Private Sub ListBox2_Click()
    ListBox1.ListIndex = -1
    If ListBox2.ListCount >= 0 Then
        TextBox1.Text = ListBox2.List(ListBox2.ListIndex)
    End If
End Sub

Private Sub ListBox3_Click()
    With ListBox3
        If .ListCount >= 0 Then
            TextBox1.Text = .List(.ListIndex, 0)
            TextBox2.Text = .List(.ListIndex, 2)
            TextBox3.Text = .List(.ListIndex, 2) ' Для показывания иконки
        End If
    End With
End Sub

Private Sub ListBox4_Click()
    With ListBox4
        If .ListCount >= 0 Then
            TextBox4.Text = .List(.ListIndex, 0)
            TextBox5.Text = .List(.ListIndex, 2)
            TextBox3.Text = .List(.ListIndex, 2) ' Для показывания иконки
        End If
    End With
End Sub

Private Sub ListBox5_Click()
    If ListBox5.ListCount >= 0 Then
        TextBox4.Text = ListBox5.List(ListBox5.ListIndex)
    End If
End Sub

Private Sub ListBox5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    CommandButton28_Click
End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    ViewLoadToForm ' Загрузка Видов на форму
    TegLoadToForm ' Загрузка Тегов на форму
End Sub

Private Sub ViewLoadToForm()
    Path = ActiveDocument.Path & "\" & ViewDataFileName 'ViewDataFileName из основного модуля
    If Not СуществуетФайл(Path) Then
        'MsgBox "Не найден файл видов '" & ViewDataFileName & "' в папке" & vbNewLine & ActiveDocument.Path
        Exit Sub
    End If
    
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
        ListBox1.AddItem ViewData(i).TegText
        i = i + 1
    Loop
    ListBox1.AddItem ViewData(i).TegText
    
    AddOtherView
    With OtherViewDic 'Считываем особые виды из словаря
        For Each varKey In .Keys ' или For Each varItem In .Items
          ListBox2.AddItem varKey
        Next
    End With
    LoadPopUpItemsFromFile ' Загрузка данных для меню ПКМ
    LoadPopUpOptionsFromFile ' Загрузка настроек меню ПКМ
    
    Dim ContextMenu As CommandBar
    Dim ctrl As CommandBarControl
    On Error Resume Next
    Set ContextMenu = Application.CommandBars("Text")
    For Each ctrl In ContextMenu.Controls
        If ctrl.Tag = ViewPopUpMenuTegText Then
            CheckBox1.Value = True
        End If
        If ctrl.Tag = TegsPopUpMenuTegText Then
            CheckBox2.Value = True
        End If
    Next ctrl
End Sub

Private Sub TegLoadToForm()
    Path = ActiveDocument.Path & "\" & TegsFileName 'TegsFileName из основного модуля
    If СуществуетФайл(Path) Then
        Call SaveLoadListbox(ListBox5, Path, "load")
    End If
   
    Path = ActiveDocument.Path & "\" & TegsDataFileName 'TegsDataFileName из основного модуля
    If Not СуществуетФайл(Path) Then
        Exit Sub
    End If
    
    Dim TegsPopUpData() As ViewPopUpType
    ReDim Preserve TegsPopUpData(ViewPopUpDataMaxSize)
    Close #1
    Open Path For Binary As #1
    Get #1, 1, TegsPopUpData
    Close #1
    'Заполнение таблицы
    ListBox4.Clear
    'For i = 0 To UBound(ViewData)
    'Next i
    i = 0
    Do While TegsPopUpData(i).FaceID <> 0
        With ListBox4
            .AddItem TegsPopUpData(i).Name
            .List(i, 0) = TegsPopUpData(i).Name
            .List(i, 1) = TegsPopUpData(i).Teg
            .List(i, 2) = TegsPopUpData(i).FaceID
            .List(i, 3) = TegsPopUpData(i).Code
        End With
        i = i + 1
    Loop
'    LoadPopUpItemsFromFile ' Загрузка данных для меню ПКМ
'    LoadPopUpOptionsFromFile ' Загрузка настроек меню ПКМ
'
'    Dim ContextMenu As CommandBar
'    Dim ctrl As CommandBarControl
'    On Error Resume Next
'    Set ContextMenu = Application.CommandBars("Text")
'    For Each ctrl In ContextMenu.Controls
'        If ctrl.Tag = ViewPopUpMenuTegText Then
'            CheckBox1.Value = True
'        End If
'    Next ctrl
End Sub

Public Sub CopyTextFromHelpPopUp()
    TextBox2.Text = CommandBars.ActionControl.Parameter
End Sub

Private Sub SavePopUpItemsToFile()
    If ListBox3.ListCount = 0 Then
        Exit Sub
    End If
    Dim ViewPopUpData() As ViewPopUpType
    ReDim Preserve ViewPopUpData(ListBox3.ListCount - 1)
    For i = 0 To ListBox3.ListCount - 1
        With ListBox3
            ViewPopUpData(i).Name = .List(i, 0)
            ViewPopUpData(i).Teg = .List(i, 1)
            ViewPopUpData(i).FaceID = .List(i, 2)
            ViewPopUpData(i).Code = .List(i, 3)
        End With
    Next i
    'Сохранение
    'ViewPopUpDataMaxSize
    Path = ActiveDocument.Path & "\" & ViewPopUpDataFileName
    Open Path For Binary As #1
    Put #1, 1, ViewPopUpData 'Сохранение данных в файл
    Close #1
End Sub

Private Sub LoadPopUpItemsFromFile()
    Path = ActiveDocument.Path & "\" & ViewPopUpDataFileName 'ViewPopUpDataFileName из основного модуля
    If Not СуществуетФайл(Path) Then
        MsgBox "Не найден файл видов в сплывающем меню '" & ViewPopUpDataFileName & "' в папке" & vbNewLine & ActiveDocument.Path
        Exit Sub
    End If
    Dim ViewPopUpData() As ViewPopUpType
    ReDim Preserve ViewPopUpData(ViewPopUpDataMaxSize)
    Close #1
    Open Path For Binary As #1
    Get #1, 1, ViewPopUpData
    Close #1
    'Заполнение таблицы
    ListBox3.Clear
    'For i = 0 To UBound(ViewData)
    'Next i
    i = 0
    Do While ViewPopUpData(i).FaceID <> 0
        With ListBox3
            .AddItem ViewPopUpData(i).Name
            .List(i, 0) = ViewPopUpData(i).Name
            .List(i, 1) = ViewPopUpData(i).Teg
            .List(i, 2) = ViewPopUpData(i).FaceID
            .List(i, 3) = ViewPopUpData(i).Code
        End With
        i = i + 1
    Loop
'    With ListBox3
'            .AddItem ViewPopUpData(i).Name
'            .List(i, 0) = ViewPopUpData(i).Name
'            .List(i, 1) = ViewPopUpData(i).Teg
'            .List(i, 2) = ViewPopUpData(i).FaceID
'            .List(i, 3) = ViewPopUpData(i).Code
'    End With
End Sub

Private Sub SavePopUpOptionsToFile()
    Dim PopUpOptionsData As PopUpOptions 'Файл с настройками отображения меню Вида и Тегов при ПКМ
    ShowViewMenu = CheckBox1.Value ' Показывать меню "Вид" (при ПКМ на обычном тексте)
    ShowTegMenuOnText = CheckBox2.Value ' Показывать меню "Тег" (при ПКМ на обычном тексте)
    ShowTegMenuOnHandle = CheckBox3.Value ' Показывать меню "Тег" (при ПКМ на заголовке)
    
    PopUpOptionsData.ShowViewPopUp = ShowViewMenu
    PopUpOptionsData.ShowTegMenuOnText = ShowTegMenuOnText
    PopUpOptionsData.ShowTegMenuOnHandle = ShowTegMenuOnHandle
    'Сохранение
    Path = ActiveDocument.Path & "\" & PopUpOptionsFileName
    Open Path For Binary As #1
    Put #1, 1, PopUpOptionsData 'Сохранение данных в файл
    Close #1
End Sub

Private Sub DefaultPopUpOptionsToFile()
    Dim PopUpOptionsData As PopUpOptions 'Файл с настройками отображения меню Вида и Тегов при ПКМ
    ShowViewMenu = False ' Показывать меню "Вид" (при ПКМ на обычном тексте)
    ShowTegMenuOnText = False ' Показывать меню "Тег" (при ПКМ на обычном тексте)
    ShowTegMenuOnHandle = False ' Показывать меню "Тег" (при ПКМ на заголовке)
    
    PopUpOptionsData.ShowViewPopUp = ShowViewMenu
    PopUpOptionsData.ShowTegMenuOnText = ShowTegMenuOnText
    PopUpOptionsData.ShowTegMenuOnHandle = ShowTegMenuOnHandle
    'Сохранение
    Path = ActiveDocument.Path & "\" & PopUpOptionsFileName
    Open Path For Binary As #1
    Put #1, 1, PopUpOptionsData 'Сохранение данных в файл
    Close #1
End Sub

Public Sub LoadPopUpOptionsFromFile()
    Path = ActiveDocument.Path & "\" & PopUpOptionsFileName 'PopUpOptionsFileName из основного модуля
    If Not СуществуетФайл(Path) Then
        MsgBox "Не найден файл настройки видов в сплывающем меню '" & PopUpOptionsFileName & "' в папке" & vbNewLine & ActiveDocument.Path
        DefaultPopUpOptionsToFile
        Exit Sub
    End If
    Dim PopUpOptionsData As PopUpOptions 'Файл с настройками отображения меню Вида и Тегов при ПКМ
    Close #1
    Open Path For Binary As #1
    Get #1, 1, PopUpOptionsData
    Close #1
    CheckBox1.Value = PopUpOptionsData.ShowViewPopUp ' Показывать меню "Вид" (при ПКМ на обычном тексте)
    CheckBox2.Value = PopUpOptionsData.ShowTegMenuOnText ' Показывать меню "Тег" (при ПКМ на обычном тексте)
    CheckBox3.Value = PopUpOptionsData.ShowTegMenuOnHandle ' Показывать меню "Тег" (при ПКМ на заголовке)
    
    ApplyPopUpOptions ' Приминение настроек
End Sub
