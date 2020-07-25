VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StatisticaForm 
   Caption         =   "����������"
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
Public Sub ������������������������(Teg)
'������������
    Dim Times() As String
    Dim rgePages As Range
    Dim cursorEnd As Range
    'Teg = "[������]"
    MaxCount = �������������(Teg) '������������ ���������� ����������
    Dim Stat() As String
    ReDim Stat(MaxCount - 1)
    ������������
    Selection.TypeParagraph
    
    �������������
    Selection.find.ClearFormatting
    Selection.find.Style = ActiveDocument.Styles("��������� 4;�_������")
    With Selection.find
        .Text = Teg
        .Execute
    End With
    i = 0
    'ActiveDocument.Range(0, Selection.Start).Paragraphs.Count
    'MsgBox (ActiveDocument.Range(0, Selection.Start).Paragraphs.Count)

    Do While Selection.find.Found = True And i < MaxCount
        i = i + 1
        Ls = �����������
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
        Selection.find.Style = ActiveDocument.Styles("��������� 4;�_������")
        With Selection.find
            .Text = Teg
            .Execute
        End With
    Loop
    ������������
    Selection.TypeParagraph
    Selection.ClearFormatting
    Selection.TypeText Text:="���������� ������� '" + Teg + "' (" + Trim(Str(MaxCount)) + ")."
    Selection.TypeParagraph
    Selection.TypeText Text:="�" & vbTab & "�������" & vbTab & "����" & vbTab & "������ (��� ��������)" & vbTab & "������ (� ���������)" & vbTab & "�������" & vbTab & "�����" & vbNewLine
    For i = 0 To UBound(Stat) - 1 ' "-1" �.�. ���������� ������� ��� ���������� ���������
        Selection.TypeText Text:=Stat(i)
    Next i
End Sub

Private Sub CommandButton1_Click()
    sngStart = Timer                               ' ������ �������
    ������������������������ ("")
    sngEnd = Timer                                 ' �����
    sngElapsed = Format(sngEnd - sngStart, "Fixed") ' ����������.
    ������������
    Selection.TypeParagraph
    Selection.ClearFormatting
    Selection.TypeText Text:="������� ����� " & sngElapsed & " ������."
End Sub
