VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StatsFromFolderForm 
   Caption         =   "Статистика файлов из папки"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7335
   OleObjectBlob   =   "StatsFromFolderForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StatsFromFolderForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function ВыбратьПапку(Title)
    Dim oFD As FileDialog
    Dim X, lf As Long
    'назначаем переменной ссылку на экземпляр диалога
    Set oFD = Application.FileDialog(msoFileDialogFolderPicker)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    currentPath = FSO.GetAbsolutePathName(".") 'Путь туда, где лежит файл
    With oFD 'используем короткое обращение к объекту
    'так же можно без oFD
    'With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = Title 'заголовок окна диалога
        .ButtonName = "Выбрать папку"
        .Filters.Clear 'очищаем установленные ранее типы файлов
        .InitialFileName = currentPath 'назначаем первую папку отображения
        .InitialView = msoFileDialogViewList 'вид диалогового окна(доступно 9 вариантов)
        If oFD.Show = 0 Then Exit Function 'показывает диалог
        'цикл по коллекции выбранных в диалоге файлов
        X = .SelectedItems(1) 'считываем путь к папке
        ВыбратьПапку = X
        'MsgBox "Выбрана папка: '" & x & "'", vbInformation, "Заголовок"
    End With
End Function

Public Function ФайловВПапке(currentPath, FileFormat)
' Подсчёт всех RTF документов в указанном месте
' FileFormat = "RTF" или "DOC" или "DOCM" или "DOCX"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    'currentPath = FSO.GetAbsolutePathName(".") 'Путь туда, где лежит файл
    'currentPath = "I:\Disk_G\torrents\Книги\Терри Пратчетт - Собрание сочинений\RTF\Плоский мир" 'Или путь вручную
    Set FLD = FSO.GetFolder(currentPath)
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = False
    Count = 0
    For Each Fil In FLD.Files
        If UCase(FSO.GetExtensionName(Fil.Name)) = FileFormat Then
            Count = Count + 1
        End If
    Next
    objWord.Quit
    Set oShell = Nothing
    Set FLD = Nothing
    Set FSO = Nothing
    ФайловВПапке = Count
End Function

Private Sub СтатистикаФайловИзПапки(currentPath, FileFormat)
' Статистика всех RTF документов в указанном месте
' FileFormat = "RTF" или "DOC" или "DOCM" или "DOCX"
    Const wdStatisticCharacters = 3
    Const wdStatisticCharactersWithSpaces = 5
    Const wdStatisticFarEastCharacters = 6
    Const wdStatisticLines = 1
    Const wdStatisticPages = 2
    Const wdStatisticParagraphs = 4
    Const wdStatisticWords = 0

    Selection.TypeText Text:="Файл"
    Selection.TypeText Text:=vbTab
    Selection.TypeText Text:="Символов"
    Selection.TypeText Text:=vbTab
    Selection.TypeText Text:="Символов с пробелами"
    Selection.TypeText Text:=vbTab
    'Selection.TypeText Text:="Far East characters: "
    'Selection.TypeText Text:=vbTab
    Selection.TypeText Text:="Линий"
    Selection.TypeText Text:=vbTab
    Selection.TypeText Text:="Страниц"
    Selection.TypeText Text:=vbTab
    Selection.TypeText Text:="Параграфов"
    Selection.TypeText Text:=vbTab
    Selection.TypeText Text:="Слов"
    Selection.TypeText Text:=vbTab
    Selection.TypeText Text:="Картинок"
    Selection.TypeParagraph 'С новой строки
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    'currentPath = FSO.GetAbsolutePathName(".") 'Путь туда, где лежит файл
    'currentPath = "I:\Disk_G\torrents\Книги\Терри Пратчетт - Собрание сочинений\RTF\Плоский мир" 'Или путь вручную
    Set FLD = FSO.GetFolder(currentPath)
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = False

    For Each Fil In FLD.Files
        If UCase(FSO.GetExtensionName(Fil.Name)) = FileFormat Then
            Set objDoc = objWord.Documents.Open(currentPath & "\" & Fil.Name)
            Selection.TypeText Text:=Fil.Name
            Selection.TypeText Text:=vbTab
            Selection.TypeText Text:=objDoc.ComputeStatistics(wdStatisticCharacters)
            Selection.TypeText Text:=vbTab
            Selection.TypeText Text:=objDoc.ComputeStatistics(wdStatisticCharactersWithSpaces)
            Selection.TypeText Text:=vbTab
            'Selection.TypeText Text:=objDoc.ComputeStatistics(wdStatisticFarEastCharacters)
            'Selection.TypeText Text:=vbTab
            Selection.TypeText Text:=objDoc.ComputeStatistics(wdStatisticLines)
            Selection.TypeText Text:=vbTab
            Selection.TypeText Text:=objDoc.ComputeStatistics(wdStatisticPages)
            Selection.TypeText Text:=vbTab
            Selection.TypeText Text:=objDoc.ComputeStatistics(wdStatisticParagraphs)
            Selection.TypeText Text:=vbTab
            Selection.TypeText Text:=objDoc.ComputeStatistics(wdStatisticWords)
            Selection.TypeText Text:=vbTab
            Selection.TypeText Text:=objDoc.InlineShapes.Count 'Картинок
            Selection.TypeParagraph 'С новой строки
            objDoc.Saved = True
            objDoc.Close
        End If
    Next
    objWord.Quit
    Set oShell = Nothing
    Set FLD = Nothing
    Set FSO = Nothing
End Sub

Private Sub CheckBox2_Click()

End Sub

Private Sub CommandButton1_Click()
    Path = ВыбратьПапку("Папка для обработки")
    TextBox1.Value = Path
    CommandButton3_Click
End Sub

Private Sub ОбработатьФайл(FileFormat)
' FileFormat = "RTF" или "DOC" или "DOCM" или "DOCX"
    КурсорВКонец
    Selection.TypeParagraph
    Set currentPosition = Selection.Range
    СтатистикаФайловИзПапки TextBox1.Value, FileFormat '"RTF" или "DOC" или "DOCM"
    If CheckBox1.Value = True Then
        currentPosition.Select
        ВыделитьОтКурсораДоКонца
        Selection.Cut
        Dim xlApp As Object
        Dim xlBook As Object
        Dim xlSheet As Object
        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Add
        Set xlSheet = xlBook.Worksheets(1)
        xlSheet.Application.Visible = True
        xlSheet.Application.Cells(1, 1).Activate
        xlSheet.Application.ActiveSheet.Paste ' Вставить текст из буфера
        If TextBox2.Value <> "" Then
            xlSheet.Application.Columns("A:A").ColumnWidth = TextBox2.Value 'Установить ширину первого столбца
            xlSheet.Application.Cells.Select 'Выделить все ячейки
            xlSheet.Application.Selection.Rows.AutoFit 'Подогнать по высоте
            xlSheet.Application.Cells(1, 1).Select
        End If
        Selection.TypeBackspace
    End If
End Sub

Private Sub CommandButton2_Click()
    If TextBox1.Value <> "" Then
        If CheckBox2.Value = True And Label3.caption > 0 Then ОбработатьФайл ("RTF") '"RTF" или "DOC" или "DOCM" или "DOCX"
        If CheckBox3.Value = True And Label5.caption > 0 Then ОбработатьФайл ("DOC") '"RTF" или "DOC" или "DOCM" или "DOCX"
        If CheckBox4.Value = True And Label7.caption > 0 Then ОбработатьФайл ("DOCM") '"RTF" или "DOC" или "DOCM" или "DOCX"
        If CheckBox5.Value = True And Label9.caption > 0 Then ОбработатьФайл ("DOCX") '"RTF" или "DOC" или "DOCM" или "DOCX"
    End If
    'ВыделитьМеждуПозициями(PosStart, PosEnd)
End Sub

Private Sub CommandButton3_Click()
    Path = TextBox1.Value
    If Path <> "" Then
        CommandButton2.Enabled = False
        CommandButton2.caption = "Подождите..."
        Label3.caption = ФайловВПапке(Path, "RTF") '"RTF" или "DOC" или "DOCM"
        Label5.caption = ФайловВПапке(Path, "DOC") '"RTF" или "DOC" или "DOCM"
        Label7.caption = ФайловВПапке(Path, "DOCM") '"RTF" или "DOC" или "DOCM"
        Label9.caption = ФайловВПапке(Path, "DOCX") '"RTF" или "DOC" или "DOCM" или "DOCX"
        CommandButton2.caption = "Обработать"
        CommandButton2.Enabled = True
    End If
End Sub

Private Sub Label3_Click()

End Sub

