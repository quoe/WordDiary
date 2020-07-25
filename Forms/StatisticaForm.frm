VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StatisticaForm 
   Caption         =   "Статистика"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "StatisticaForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StatisticaForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub СтатистикаКаждогоСобытия(Teg)
'ЭкспортТегов
    Dim Times() As String
    Dim rgePages As Range
    Dim cursorEnd As Range
    'Teg = "[Цитата]"
    MaxCount = СобытийСТегом(Teg) 'Максимальное количество заголовков
    Dim Stat() As String
    ReDim Stat(MaxCount - 1)
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
        Set myRange = Selection.Range
        NumPages = myRange.ComputeStatistics(wdStatisticPages)
        NumWords = myRange.ComputeStatistics(wdStatisticWords)
        NumCharacters = myRange.ComputeStatistics(wdStatisticCharacters)
        NumCharactersWithSpaces = myRange.ComputeStatistics(wdStatisticCharactersWithSpaces)
        NumParagraphs = myRange.ComputeStatistics(wdStatisticParagraphs)
        NumLines = myRange.ComputeStatistics(wdStatisticLines)
        Stat(i - 1) = (i) & vbTab & NumPages & vbTab & NumWords & vbTab & NumCharacters & vbTab & NumCharactersWithSpaces & vbTab & NumParagraphs & vbTab & NumLines & vbNewLine
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
    Selection.TypeText Text:="Статистика событий '" + Teg + "' (" + Trim(Str(MaxCount)) + ")."
    Selection.TypeParagraph
    Selection.TypeText Text:="№" & vbTab & "Страниц" & vbTab & "Слов" & vbTab & "Знаков (без пробелов)" & vbTab & "Знаков (с пробелами)" & vbTab & "Абзацев" & vbTab & "Строк" & vbNewLine
    For i = 0 To UBound(Stat) - 1 ' "-1" т.к. последнего события нет следующего заголовка
        Selection.TypeText Text:=Stat(i)
    Next i
End Sub

Private Sub CommandButton1_Click()
    sngStart = Timer                               ' Начало отсчёта
    СтатистикаКаждогоСобытия ("")
    sngEnd = Timer                                 ' Конец
    sngElapsed = Format(sngEnd - sngStart, "Fixed") ' Приращение.
    КурсорВКонец
    Selection.TypeParagraph
    Selection.ClearFormatting
    Selection.TypeText Text:="Экспорт занял " & sngElapsed & " секунд."
End Sub
