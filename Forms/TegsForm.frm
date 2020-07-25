VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TegsForm 
   Caption         =   "Теги"
   ClientHeight    =   4890
   ClientLeft      =   195
   ClientTop       =   585
   ClientWidth     =   8805.001
   OleObjectBlob   =   "TegsForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TegsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Path = ActiveDocument.Path + "\" + TegsFileName + ".txt"
    Call SaveLoadListbox(ListBox1, Path, "save")
End Sub

Private Sub CommandButton10_Click()
    ListBox1.AddItem "Идея"
    ListBox1.AddItem "Мысль"
    ListBox1.AddItem "Событие"
    ListBox1.AddItem "Воспоминание"
    ListBox1.AddItem "Список дел"
    ListBox1.AddItem "Диалог из ВК"
    ListBox1.AddItem "Диалог из Skype"
    ListBox1.AddItem "Сон"
    ListBox1.AddItem "Цитата"
    ListBox1.AddItem "Фильм"
    ListBox1.AddItem "Сериал"
    ListBox1.AddItem "Аниме"
    ListBox1.AddItem "Карта"
    ListBox1.AddItem "Числа"
    ListBox1.AddItem "Таблица"
    ListBox1.AddItem "Стихи"
    ListBox1.AddItem "Картинка"
    ListBox1.AddItem "Дополнение"
    ListBox1.AddItem "Записка"
    ListBox1.AddItem "Код"
    ListBox1.AddItem "Дописать"
    ListBox1.AddItem "Посмотреть"
    ListBox1.AddItem "Книга"
    ListBox1.AddItem "Статистика"
    
    ОбновитьИнформацию
End Sub

Private Sub CommandButton11_Click()
    КурсорВКонец
    Set cursorBackup = Selection.Range
    Teg = TegOpen + TextBox1.Text + TegClose
    sngStart = Timer                               ' Начало отсчёта
    ЭкспортТегов (Teg)
    sngEnd = Timer                                 ' Конец
    sngElapsed = Format(sngEnd - sngStart, "Fixed") ' Приращение.
    КурсорВКонец
    Selection.TypeParagraph
    Selection.ClearFormatting
    Selection.TypeText Text:="Экспорт тегов занял " & sngElapsed & " секунд."
    cursorBackup.Select
    ВыделитьОтКурсораДоКонца
    Selection.Copy
    Selection.TypeBackspace

    Dim WordApp As Word.Application '  экземпляр приложения
    Dim DocWord As Word.Document '  экземпляр документа
    
    'создаём  новый экземпляр Word-a
    Set WordApp = New Word.Application
    
    'определяем видимость Word-a по True - видимый,
    'по False - не видимый (работает только ядро)
    WordApp.Visible = True
    
    'создаём новый документ в Word-e
    Set DocWord = WordApp.Documents.Add
    
    'активируем его
    DocWord.Activate
    Selection.Paste
End Sub

Private Sub CommandButton12_Click()
    TegOpenOld = TextBox2.Text
    TegOpenNew = TextBox3.Text
    TegCloseOld = TextBox4.Text
    TegCloseNew = TextBox5.Text

    If TegOpenOld <> "" Then
        Selection.find.ClearFormatting
        Selection.find.Replacement.ClearFormatting
        Selection.find.Replacement.Style = ActiveDocument.Styles("Заголовок 4;З_Момент")
        With Selection.find
            .Text = TegOpenOld
            .Replacement.Text = TegOpenNew
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.find.Execute Replace:=wdReplaceAll
        
        If TegOpenNew = " " Then
            Selection.find.ClearFormatting
            Selection.find.Replacement.ClearFormatting
            Selection.find.Replacement.Style = ActiveDocument.Styles("Заголовок 4;З_Момент")
            With Selection.find
                .Text = ",  "
                .Replacement.Text = ", "
                .Forward = True
                .Wrap = wdFindContinue
                .Format = True
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.find.Execute Replace:=wdReplaceAll
        End If
    End If
    
    If TegCloseOld <> "" Then
        Selection.find.ClearFormatting
        Selection.find.Replacement.ClearFormatting
        Selection.find.Replacement.Style = ActiveDocument.Styles("Заголовок 4;З_Момент")
        With Selection.find
            .Text = TegCloseOld
            .Replacement.Text = TegCloseNew
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.find.Execute Replace:=wdReplaceAll
        
        If TegCloseNew = " " Then
            Selection.find.ClearFormatting
            Selection.find.Replacement.ClearFormatting
            Selection.find.Replacement.Style = ActiveDocument.Styles("Заголовок 4;З_Момент")
            With Selection.find
                .Text = " ,"
                .Replacement.Text = ","
                .Forward = True
                .Wrap = wdFindContinue
                .Format = True
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.find.Execute Replace:=wdReplaceAll
        End If
    End If
End Sub

Private Sub CommandButton2_Click()
    UserForm_Initialize
End Sub


Private Sub CommandButton3_Click()
    If TextBox1.Text = "" Then Exit Sub
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.List(i) = TextBox1.Text Then
            MsgBox "Такой элемент уже есть!"
            Exit Sub
        End If
    Next i
    ListBox1.AddItem TextBox1.Text
    ListBox1.ListIndex = ListBox1.ListCount - 1 ' Выделить последний
    ОбновитьИнформацию
End Sub

Private Sub CommandButton4_Click()
    If ListBox1.ListIndex = -1 Then Exit Sub
    ListBox1.RemoveItem (ListBox1.ListIndex)
    ОбновитьИнформацию
End Sub

Private Sub CommandButton5_Click()
If ListBox1.ListIndex = -1 Then Exit Sub
For i = 0 To ListBox1.ListCount - 1
    If ListBox1.List(i) = TextBox1.Text Then
        MsgBox "Такой элемент уже есть!"
        Exit Sub
    End If
Next i
    ListBox1.List(ListBox1.ListIndex) = TextBox1.Text
    ОбновитьИнформацию
End Sub

Private Sub CommandButton6_Click()
    Selection.TypeText Text:=TextBox1.Text
End Sub

Private Sub CommandButton7_Click()
    Selection.TypeText Text:=", " + TegOpen + TextBox1.Text + TegClose
End Sub

Private Sub CommandButton8_Click()
    Set Paragraph = Selection.Range 'Запомнить текущие положение курсора
    If ДобавлениеТегаКВремени(TextBox1.Text, False) Then 'Добавить тег к времени не затирая время
    End If
    'CommandButton7_Click  'Текст тега
End Sub

Private Sub CommandButton9_Click()
    ListBox1.Clear
End Sub

Private Sub Frame2_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub ListBox1_Click()
    TextBox1.Text = ListBox1.List(ListBox1.ListIndex)
End Sub

Private Sub ОбновитьИнформацию()
    Label1.caption = "Количество: " + CStr(ListBox1.ListCount)
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    CommandButton8_Click
    If CheckBox1.Value = True Then Unload Me
End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub UserForm_Initialize()
    Path = ActiveDocument.Path + "\" + TegsFileName + ".txt" ' TegsFileName из основного модуля
    If Not СуществуетФайл(Path) Then ТекстФайлЗапись Path, ""
    Call SaveLoadListbox(ListBox1, Path, "load")
    'If ListBox1.ListCount > 0 Then ListBox1.ListIndex = 0
    ОбновитьИнформацию
End Sub

Public Function НомерСтроки()
' Выдаёт номер строки на текущей странице
    НомерСтроки = Selection.Information(wdFirstCharacterLineNumber)
End Function

Public Function НомерСекции()
' Выдаёт номер секции где курсор
    НомерСекции = Selection.Information(wdActiveEndSectionNumber)
End Function

Public Sub ЭкспортТегов(Teg)
'ЭкспортТегов
    Dim Times() As String
    Dim rgePages As Range
    Dim cursorEnd As Range
    'Teg = "[Цитата]"
    MaxCount = СобытийСТегом(Teg) 'Максимальное количество заголовков
    КурсорВКонец
    Selection.TypeParagraph
    
    КурсорВНачало
    Selection.find.ClearFormatting
    Selection.find.Style = ActiveDocument.Styles("Заголовок 4;З_Момент")
    With Selection.find
        .Text = Teg
        .Execute
    End With
    i = 0
    'ActiveDocument.Range(0, Selection.Start).Paragraphs.Count
    'MsgBox (ActiveDocument.Range(0, Selection.Start).Paragraphs.Count)

    Do While Selection.find.Found = True And i < MaxCount
        i = i + 1
        Ls = НомерСтроки
        Selection.HomeKey Unit:=wdLine
        'Selection.GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Count:=Ls
        Set rgePages = Selection.Range
        Selection.GoTo What:=wdGoToHeading, which:=wdGoToNext
        Set cursorEnd = Selection.Range
        rgePages.End = Selection.Range.End
        rgePages.Select
        Selection.Copy
        КурсорВКонец
        Selection.Paste
        cursorEnd.Select
        Selection.find.ClearFormatting
        Selection.find.Style = ActiveDocument.Styles("Заголовок 4;З_Момент")
        With Selection.find
            .Text = Teg
            .Execute
        End With
    Loop
    КурсорВКонец
    Selection.TypeParagraph
    Selection.ClearFormatting
    Selection.TypeText Text:="Экспорт тегов '" + Teg + "' (" + Trim(Str(MaxCount)) + ")."
    Selection.TypeParagraph
End Sub
