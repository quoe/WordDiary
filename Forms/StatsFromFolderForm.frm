VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StatsFromFolderForm 
   Caption         =   "���������� ������ �� �����"
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
Function ������������(Title)
    Dim oFD As FileDialog
    Dim X, lf As Long
    '��������� ���������� ������ �� ��������� �������
    Set oFD = Application.FileDialog(msoFileDialogFolderPicker)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    currentPath = FSO.GetAbsolutePathName(".") '���� ����, ��� ����� ����
    With oFD '���������� �������� ��������� � �������
    '��� �� ����� ��� oFD
    'With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = Title '��������� ���� �������
        .ButtonName = "������� �����"
        .Filters.Clear '������� ������������� ����� ���� ������
        .InitialFileName = currentPath '��������� ������ ����� �����������
        .InitialView = msoFileDialogViewList '��� ����������� ����(�������� 9 ���������)
        If oFD.Show = 0 Then Exit Function '���������� ������
        '���� �� ��������� ��������� � ������� ������
        X = .SelectedItems(1) '��������� ���� � �����
        ������������ = X
        'MsgBox "������� �����: '" & x & "'", vbInformation, "���������"
    End With
End Function

Public Function ������������(currentPath, FileFormat)
' ������� ���� RTF ���������� � ��������� �����
' FileFormat = "RTF" ��� "DOC" ��� "DOCM" ��� "DOCX"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    'currentPath = FSO.GetAbsolutePathName(".") '���� ����, ��� ����� ����
    'currentPath = "I:\Disk_G\torrents\�����\����� �������� - �������� ���������\RTF\������� ���" '��� ���� �������
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
    ������������ = Count
End Function

Private Sub �����������������������(currentPath, FileFormat)
' ���������� ���� RTF ���������� � ��������� �����
' FileFormat = "RTF" ��� "DOC" ��� "DOCM" ��� "DOCX"
    Const wdStatisticCharacters = 3
    Const wdStatisticCharactersWithSpaces = 5
    Const wdStatisticFarEastCharacters = 6
    Const wdStatisticLines = 1
    Const wdStatisticPages = 2
    Const wdStatisticParagraphs = 4
    Const wdStatisticWords = 0

    Selection.TypeText Text:="����"
    Selection.TypeText Text:=vbTab
    Selection.TypeText Text:="��������"
    Selection.TypeText Text:=vbTab
    Selection.TypeText Text:="�������� � ���������"
    Selection.TypeText Text:=vbTab
    'Selection.TypeText Text:="Far East characters: "
    'Selection.TypeText Text:=vbTab
    Selection.TypeText Text:="�����"
    Selection.TypeText Text:=vbTab
    Selection.TypeText Text:="�������"
    Selection.TypeText Text:=vbTab
    Selection.TypeText Text:="����������"
    Selection.TypeText Text:=vbTab
    Selection.TypeText Text:="����"
    Selection.TypeText Text:=vbTab
    Selection.TypeText Text:="��������"
    Selection.TypeParagraph '� ����� ������
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    'currentPath = FSO.GetAbsolutePathName(".") '���� ����, ��� ����� ����
    'currentPath = "I:\Disk_G\torrents\�����\����� �������� - �������� ���������\RTF\������� ���" '��� ���� �������
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
            Selection.TypeText Text:=objDoc.InlineShapes.Count '��������
            Selection.TypeParagraph '� ����� ������
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
    Path = ������������("����� ��� ���������")
    TextBox1.Value = Path
    CommandButton3_Click
End Sub

Private Sub ��������������(FileFormat)
' FileFormat = "RTF" ��� "DOC" ��� "DOCM" ��� "DOCX"
    ������������
    Selection.TypeParagraph
    Set currentPosition = Selection.Range
    ����������������������� TextBox1.Value, FileFormat '"RTF" ��� "DOC" ��� "DOCM"
    If CheckBox1.Value = True Then
        currentPosition.Select
        ������������������������
        Selection.Cut
        Dim xlApp As Object
        Dim xlBook As Object
        Dim xlSheet As Object
        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Add
        Set xlSheet = xlBook.Worksheets(1)
        xlSheet.Application.Visible = True
        xlSheet.Application.Cells(1, 1).Activate
        xlSheet.Application.ActiveSheet.Paste ' �������� ����� �� ������
        If TextBox2.Value <> "" Then
            xlSheet.Application.Columns("A:A").ColumnWidth = TextBox2.Value '���������� ������ ������� �������
            xlSheet.Application.Cells.Select '�������� ��� ������
            xlSheet.Application.Selection.Rows.AutoFit '��������� �� ������
            xlSheet.Application.Cells(1, 1).Select
        End If
        Selection.TypeBackspace
    End If
End Sub

Private Sub CommandButton2_Click()
    If TextBox1.Value <> "" Then
        If CheckBox2.Value = True And Label3.caption > 0 Then �������������� ("RTF") '"RTF" ��� "DOC" ��� "DOCM" ��� "DOCX"
        If CheckBox3.Value = True And Label5.caption > 0 Then �������������� ("DOC") '"RTF" ��� "DOC" ��� "DOCM" ��� "DOCX"
        If CheckBox4.Value = True And Label7.caption > 0 Then �������������� ("DOCM") '"RTF" ��� "DOC" ��� "DOCM" ��� "DOCX"
        If CheckBox5.Value = True And Label9.caption > 0 Then �������������� ("DOCX") '"RTF" ��� "DOC" ��� "DOCM" ��� "DOCX"
    End If
    '����������������������(PosStart, PosEnd)
End Sub

Private Sub CommandButton3_Click()
    Path = TextBox1.Value
    If Path <> "" Then
        CommandButton2.Enabled = False
        CommandButton2.caption = "���������..."
        Label3.caption = ������������(Path, "RTF") '"RTF" ��� "DOC" ��� "DOCM"
        Label5.caption = ������������(Path, "DOC") '"RTF" ��� "DOC" ��� "DOCM"
        Label7.caption = ������������(Path, "DOCM") '"RTF" ��� "DOC" ��� "DOCM"
        Label9.caption = ������������(Path, "DOCX") '"RTF" ��� "DOC" ��� "DOCM" ��� "DOCX"
        CommandButton2.caption = "����������"
        CommandButton2.Enabled = True
    End If
End Sub

Private Sub Label3_Click()

End Sub

