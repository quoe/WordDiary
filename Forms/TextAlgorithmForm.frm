VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TextAlgorithmForm 
   Caption         =   "Управление алгоритмами"
   ClientHeight    =   9705.001
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13650
   OleObjectBlob   =   "TextAlgorithmForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TextAlgorithmForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function GetHTTPResponse(ByVal sURL As String) As String
    On Error Resume Next
    Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP")
    With oXMLHTTP
        .Open "GET", sURL, False
        ' раскомментируйте следующие строки и подставьте верные IP, логин и пароль
        ' если вы сидите за proxy
        ' .setProxy 2, "192.168.100.1:3128"
        ' .setProxyCredentials "user", "password"
        .Send
        GetHTTPResponse = .ResponseText
    End With
    Set oXMLHTTP = Nothing
End Function

Function B() 'Двойные кавычки
    B = Chr(34)
End Function

Private Sub CommandButton1_Click()
    Dim Site As String
    Dim StartDate As Date
    TextBox3.Text = ""
    StartDate = "03.06.2017"
    'Site = TextBox1.Text
    While StartDate <= Date
        Site = "http://vkonline.info/user/125071411/?date=" & StartDate
        SiteText = GetHTTPResponse(Site)
        TextBox3.Text = TextBox3.Text & GetTags(SiteText, "td", "class", "online-day", "data-day")
        SiteText = GetTags(SiteText, "td", "class", "online-day", "innerHTML")
        SiteText = Replace(SiteText, " &mdash; ", vbTab)
        SiteText = Replace(SiteText, "item-long" & B & ">", "item-long" & B & ">Компьютер" & vbTab)
        SiteText = Replace(SiteText, "time-morning" & B & ">", "time-morning" & B & ">Компьютер" & vbTab)
        SiteText = Replace(SiteText, "time-day" & B & ">", "time-day" & B & ">Компьютер" & vbTab)
        SiteText = Replace(SiteText, "time-evening" & B & ">", "time-evening" & B & ">Компьютер" & vbTab)
        SiteText = Replace(SiteText, "time-night" & B & ">", "time-night" & B & ">Компьютер" & vbTab)
        SiteText = Replace(SiteText, "online-phone" & B & ">", "online-phone" & B & ">Телефон" & vbTab)
        SiteText = GetTags(SiteText, "div", "class", "online-item*", "innerHTML")
        SiteText = GetTags(SiteText, "span", "class", "right", "DeleteTags")
        'GetTags(SiteText, "td", "class", "online-day", "data-day")
        TextBox3.Text = TextBox3.Text & vbNewLine & SiteText & vbNewLine
        StartDate = Format(DateAdd("d", 3, StartDate), "dd.mm.yyyy") 'Плюс 3 дня
    Wend
    TextBox3.Text = Replace(TextBox3.Text, "%~$", vbNewLine)
    MsgBox ("Готово!")
    'TextBox2.Text = GetTags(SiteText, "td", "class", "online-day", "data-day 2")
    'TextBox3.Text = GetTags(TextBox2.Text, "span", "class", "right", "DeleteTags")
End Sub

Function CheckSuperChar(Str)
    CheckSuperChar = Str
    If InStr(Str, "#Tab") > 0 Then
        CheckSuperChar = Replace(Str, "#Tab", vbTab)
    End If
    If InStr(Str, "#T") > 0 Then
        CheckSuperChar = Replace(Str, "#T", vbTab)
    End If
    If InStr(Str, "#NewLine") > 0 Then
        CheckSuperChar = Replace(Str, "#Tab", vbNewLine)
    End If
    If InStr(Str, "#NL") > 0 Then 'Сначала длинное название, потом сокращение, иначе беда
        CheckSuperChar = Replace(Str, "#NL", vbNewLine)
    End If
End Function

Private Function ЗаменитьВИсходном(AStr, BStr)
    cmdString = "Заменить в исходном '#А' на '#Б'"
    If AStr = "" Then
        AStr = InputBox("Введите что нужно заменить. Например, если нужно заменить '#А' на '*Б', то сейчас введите '#A'", "Заменить в исходном А на Б", "")
    Else
        AStr = InputBox("Введите что нужно заменить. Например, если нужно заменить '#А' на '*Б', то сейчас введите '#A'", "Заменить в исходном А на Б", AStr)
    End If
    If AStr = "" Then
        ЗаменитьВИсходном = ""
        Exit Function
    End If
    
    If BStr = "" Then
        BStr = InputBox("Введите на что заменять. Например, если нужно заменить '#А' на '*Б', то сейчас введите '*Б'", "Заменить в исходном А на Б", "#Tab")
    Else
        BStr = InputBox("Введите на что заменять. Например, если нужно заменить '#А' на '*Б', то сейчас введите '*Б'", "Заменить в исходном А на Б", BStr)
    End If

    cmdString = Replace(cmdString, "#А", AStr)
    cmdString = Replace(cmdString, "#Б", BStr)
    ЗаменитьВИсходном = cmdString
End Function

Private Sub CommandButton2_Click() 'Добавление команд
    cmd = ListBox1.List(ListBox1.ListIndex)
    cmdString = cmd
    If cmdString = "Заменить в исходном '#А' на '#Б'" Then
        cmdString = ЗаменитьВИсходном("", "")
        If cmdString = "" Then Exit Sub
    ElseIf cmdString = "Загрузить из интернета: '#А'" Then
        '"Заменить в обработанном '#А' на '#Б'"
        AStr = InputBox("Введите ссылку на страницу. Например, https://pikabu.ru/", "Загрузить из интернета", "")
        If AStr = "" Then Exit Sub
        cmdString = Replace(cmdString, "#А", AStr)
    ElseIf cmdString = "Получить HTML теги: '#А', '#Б', '#В', '#Г'" Then
        '"Получить HTML теги: '#А', '#Б', '#В', '#Г'". 'TagName', 'AttrName', 'AttrValue', 'Result'
        HTMLHelpStr = vbNewLine & "Например, ищем div id=" & B & "mod-lists" & B & ", и берем его начинку (innerHTML):" & vbNewLine & _
        "Тогда в качестве '#А' надо ввести 'div' (без кавычек), в качестве '#Б' - 'id', '#В' - 'mod-lists', '#Г' - 'innerHTML 1')." & vbNewLine & _
        "Индекс 1 после innerHTML означает, что если будет найдено несколько таких тегов, - макрос возьмет только первый"
        AStr = InputBox("Введите '#А' (TagName)." & HTMLHelpStr, "Получить HTML теги: '#А', '#Б', '#В', '#Г'. Введите '#А'", "")
        If AStr = "" Then Exit Sub
        BStr = InputBox("Введите '#Б' (AttrName)." & HTMLHelpStr, "Получить HTML теги: '#А', '#Б', '#В', '#Г'. Введите '#Б'", "")
        VStr = InputBox("Введите '#В' (AttrValue)." & HTMLHelpStr, "Получить HTML теги: '#А', '#Б', '#В', '#Г'. Введите '#В'", "")
        GStr = InputBox("Введите '#Г' (Result)." & HTMLHelpStr, "Получить HTML теги: '#А', '#Б', '#В', '#Г'. Введите '#Г'", "innerHTML")
        
        cmdString = Replace(cmdString, "#А", AStr)
        cmdString = Replace(cmdString, "#Б", BStr)
        cmdString = Replace(cmdString, "#В", VStr)
        cmdString = Replace(cmdString, "#Г", GStr)
    ElseIf cmdString = "Обработанный в обрабатываемый" Then
        '"Обработанный в обрабатываемый" 'Перенести текст из обработанного в обрабатываемый
        cmdString = "Обработанный в обрабатываемый"
    ElseIf cmdString = "РВ. Заменить '#Шаблон' на '#Б'" Then
        '"РВ. Заменить '#Шаблон' на '#Б'" RegExp
        cmdString = "РВ. Заменить '#Шаблон' на '#Б'"
        AStr = InputBox("Введите шаблон регулярного выражения.", "Регулярные выражения. Заменить '#Шаблон' на '#Б'", "")
        If AStr = "" Then Exit Sub
        BStr = InputBox("Введите, на что заменять шаблон.", "Регулярные выражения. Заменить '#Шаблон' на '#Б'", "")
        cmdString = Replace(cmdString, "#Шаблон", AStr)
        cmdString = Replace(cmdString, "#Б", BStr)
    ElseIf cmdString = "Документ. Загрузить весь текст" Then
        'Документ. Загрузить весь текст
        cmdString = "Документ. Загрузить весь текст"
    ElseIf cmdString = "Документ. Загрузить выделенный текст" Then
        'Документ. Загрузить весь текст
        cmdString = "Документ. Загрузить выделенный текст"
    Else
        MsgBox ("Не удалось. Команда не была распознана.")
    End If
    ДобавитьКоманду (cmdString)
    'If MsgBox("Разделять табуляцией?", vbYesNo, "Настройки") = vbYes Then
End Sub

Private Sub CommandButton3_Click()
    УдалитьКоманду
End Sub

Private Sub CommandButton4_Click() 'Выполнение команд
    TextBox3.Text = ""
    Dim AStrIndex, BStrIndex As Integer
    For i = 0 To ListBox2.ListCount - 1
        cmd = ListBox2.List(i)
        cmdString = cmd
        '"Заменить в исходном '#А' на '#Б'"
        If InStr(cmdString, "Заменить в исходном ") > 0 Then
            AStrIndex = InStr(cmdString, "'")
            AStr = Mid(cmdString, AStrIndex + 1, InStr(AStrIndex + 1, cmdString, "'") - AStrIndex - 1)
            BStrIndex = InStr(cmdString, "на '") + Len("на '") - 1
            ВStr = Mid(cmdString, BStrIndex + 1, InStr(BStrIndex + 1, cmdString, "'") - BStrIndex - 1)
            
            AStr = CheckSuperChar(AStr)
            ВStr = CheckSuperChar(ВStr)
            If TextBox3.Text = "" Then
                TextBox3.Text = TextBox2.Text
                TextBox3.Text = Replace(TextBox3.Text, AStr, ВStr)
            Else
                TextBox3.Text = Replace(TextBox3.Text, AStr, ВStr)
            End If
        End If
        '"Загрузить из интернета: '#А'"
        If InStr(cmdString, "Загрузить из интернета: ") > 0 Then
            AStrIndex = InStr(cmdString, "'")
            AStr = Mid(cmdString, AStrIndex + 1, InStr(AStrIndex + 1, cmdString, "'") - AStrIndex - 1)
            TextBox2.Text = GetHTTPResponse(AStr)
            If TextBox3.Text = "" Then
                TextBox3.Text = TextBox2.Text
                TextBox3.Text = Replace(TextBox3.Text, AStr, ВStr)
            Else
                TextBox3.Text = Replace(TextBox3.Text, AStr, ВStr)
            End If
        End If
        '"Получить HTML теги: '#А', '#Б', '#В', '#Г'"
        If InStr(cmdString, "Получить HTML теги: ") > 0 Then
            AStrIndex = InStr(cmdString, "'")
            AStr = Mid(cmdString, AStrIndex + 1, InStr(AStrIndex + 1, cmdString, "'") - AStrIndex - 1)
            BStrIndex = InStr(cmdString, ", '") + Len(", '") - 1
            ВStr = Mid(cmdString, BStrIndex + 1, InStr(BStrIndex + 1, cmdString, "'") - BStrIndex - 1)
            VStrIndex = InStr(BStrIndex, cmdString, ", '") + Len(", '") - 1
            VStr = Mid(cmdString, VStrIndex + 1, InStr(VStrIndex + 1, cmdString, "'") - VStrIndex - 1)
            GStrIndex = InStr(VStrIndex, cmdString, ", '") + Len(", '") - 1
            GStr = Mid(cmdString, GStrIndex + 1, InStr(GStrIndex + 1, cmdString, "'") - GStrIndex - 1)
            If TextBox3.Text = "" Then
                TextBox3.Text = GetTags(TextBox2.Text, AStr, BStr, VStr, GStr)
            Else
                TextBox3.Text = GetTags(TextBox2.Text, AStr, BStr, VStr, GStr)
            End If
        End If
        
        '"Обработанный в обрабатываемый"
        If InStr(cmdString, "Обработанный в обрабатываемый") > 0 Then
            TextBox2.Text = TextBox3.Text
            'ListBox2.Clear
        End If
        
        '"РВ. Заменить '#Шаблон' на '#Б'"
        If InStr(cmdString, "РВ. Заменить '") > 0 Then
            AStrIndex = InStr(cmdString, "'")
            AStr = Mid(cmdString, AStrIndex + 1, InStr(AStrIndex + 1, cmdString, "'") - AStrIndex - 1)
            BStrIndex = InStr(cmdString, "на '") + Len("на '") - 1
            ВStr = Mid(cmdString, BStrIndex + 1, InStr(BStrIndex + 1, cmdString, "'") - BStrIndex - 1)
            
            AStr = CheckSuperChar(AStr)
            ВStr = CheckSuperChar(ВStr)
            Set objRegExp = CreateObject("VBScript.RegExp")
            objRegExp.Global = True
            objRegExp.MultiLine = True
            objRegExp.Pattern = AStr
            If TextBox3.Text = "" Then
                TextBox3.Text = TextBox2.Text
                TextBox3.Text = objRegExp.Replace(TextBox3.Text, ВStr)
            Else
                TextBox3.Text = objRegExp.Replace(TextBox3.Text, ВStr)
            End If
        End If
        
        'Документ. Загрузить весь текст
        If InStr(cmdString, "Документ. Загрузить весь текст") > 0 Then
            TextBox2.Text = ActiveDocument.content.Text
        End If
        
        'Документ. Загрузить выделенный текст
        If InStr(cmdString, "Документ. Загрузить выделенный текст") > 0 Then
            TextBox2.Text = Selection.Text
        End If
    Next i
    'If MsgBox("Разделять табуляцией?", vbYesNo, "Настройки") = vbYes Then
End Sub

Private Sub CommandButton5_Click() 'Сохранить алгоритм
    Dim oFD As FileDialog
    Dim X, lf As Long
    'назначаем переменной ссылку на экземпляр диалога
    Set oFD = Application.FileDialog(msoFileDialogSaveAs)
    With oFD 'используем короткое обращение к объекту
    'так же можно без oFD
    'With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .FilterIndex = 13
        .Title = "Сохранить алгоритм (Word Diary Algorithm files)" 'заголовок окна диалога
        .InitialFileName = ActiveDocument.Path + "\" & "Word Diary Algorithm.txt" 'назначаем папку отображения и имя файла по умолчанию
        .InitialView = msoFileDialogViewDetails 'вид диалогового окна(доступно 9 вариантов)
        If oFD.Show = 0 Then Exit Sub 'показывает диалог
        'цикл по коллекции выбранных в диалоге файлов
        For lf = 1 To .SelectedItems.Count
            Path = .SelectedItems(lf) 'считываем полный путь к файлу
            Call SaveLoadListbox(ListBox2, Path, "save")
            'можно также без Path
            'Workbooks.Open .SelectedItems(lf)
        Next
    End With
End Sub

Private Sub CommandButton6_Click() 'Открыть алгоритм
    Dim oFD As FileDialog
    Dim X, lf As Long
    'назначаем переменной ссылку на экземпляр диалога
    Set oFD = Application.FileDialog(msoFileDialogFilePicker)
    With oFD 'используем короткое обращение к объекту
    'так же можно без oFD
    'With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = "Выбрать файлы отчетов" 'заголовок окна диалога
        .Filters.Clear 'очищаем установленные ранее типы файлов
        .Filters.Add "Word Diary Algorithm files", "*.wda*;*.txt*", 1 'устанавливаем возможность выбора только файлов Excel
        .Filters.Add "Text files", "*.txt", 2 'добавляем возможность выбора текстовых файлов
        .FilterIndex = 1 'устанавливаем тип файлов по умолчанию - Text files(Текстовые файлы)
        .InitialFileName = ActiveDocument.Path + "\" & "Алгоритм.wda" 'назначаем папку отображения и имя файла по умолчанию
        .InitialView = msoFileDialogViewDetails 'вид диалогового окна(доступно 9 вариантов)
        If oFD.Show = 0 Then Exit Sub 'показывает диалог
        'цикл по коллекции выбранных в диалоге файлов
        For lf = 1 To .SelectedItems.Count
            Path = .SelectedItems(lf) 'считываем полный путь к файлу
            Call SaveLoadListbox(ListBox2, Path, "load")
            'можно также без х
            'Workbooks.Open .SelectedItems(lf)
        Next
    End With
End Sub

Private Sub CommandButton7_Click() ' Изменить команду
    If ListBox2.ListIndex = -1 Then Exit Sub
    CmdStrPast = ListBox2.List(ListBox2.ListIndex)
    CmdStr = InputBox("Измениете команду", "Изменение команды", CmdStrPast)
    If CmdStr <> "" Then
        ListBox2.List(ListBox2.ListIndex) = CmdStr
        CommandButton4_Click
    End If
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub CommandButton8_Click()
    If MsgBox("Удалить все команды", vbYesNo, "Подтверждение удаления") = vbYes Then
        ListBox2.Clear
    End If
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    CommandButton2_Click
End Sub

Private Sub ListBox2_Click()

End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    CommandButton7_Click
End Sub

Private Sub ListBox2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Button ПКМ=2
    If Button = 2 Then CreateDisplayPopUpMenu
End Sub

Private Sub UserForm_Initialize()
    ListBox1.AddItem "Заменить в исходном '#А' на '#Б'"
    ListBox1.AddItem "Загрузить из интернета: '#А'"
    ListBox1.AddItem "Получить HTML теги: '#А', '#Б', '#В', '#Г'" ''#TagName', '#AttrName', '#AttrValue', '#Result'"
    ListBox1.AddItem "Обработанный в обрабатываемый"
    ListBox1.AddItem "РВ. Заменить '#Шаблон' на '#Б'"
    ListBox1.AddItem "Документ. Загрузить весь текст"
    ListBox1.AddItem "Документ. Загрузить выделенный текст"
End Sub

Private Sub ДобавитьКоманду(cmd)
    ListBox2.AddItem (Trim(Str(ListBox2.ListCount + 1)) & ". " & cmd)
    CommandButton4_Click
End Sub

Private Sub УдалитьКоманду()
    If ListBox2.ListIndex = -1 Then Exit Sub
    If MsgBox("Удалить выбранную команду:" & vbNewLine & ListBox2.List(ListBox2.ListIndex), vbYesNo, "Настройка вывода") = vbYes Then
        ListBox2.RemoveItem (ListBox2.ListIndex)
    End If
    CommandButton4_Click
End Sub
