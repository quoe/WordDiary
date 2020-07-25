VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TextAlgorithmForm 
   Caption         =   "���������� �����������"
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
        ' ���������������� ��������� ������ � ���������� ������ IP, ����� � ������
        ' ���� �� ������ �� proxy
        ' .setProxy 2, "192.168.100.1:3128"
        ' .setProxyCredentials "user", "password"
        .Send
        GetHTTPResponse = .ResponseText
    End With
    Set oXMLHTTP = Nothing
End Function

Function B() '������� �������
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
        SiteText = Replace(SiteText, "item-long" & B & ">", "item-long" & B & ">���������" & vbTab)
        SiteText = Replace(SiteText, "time-morning" & B & ">", "time-morning" & B & ">���������" & vbTab)
        SiteText = Replace(SiteText, "time-day" & B & ">", "time-day" & B & ">���������" & vbTab)
        SiteText = Replace(SiteText, "time-evening" & B & ">", "time-evening" & B & ">���������" & vbTab)
        SiteText = Replace(SiteText, "time-night" & B & ">", "time-night" & B & ">���������" & vbTab)
        SiteText = Replace(SiteText, "online-phone" & B & ">", "online-phone" & B & ">�������" & vbTab)
        SiteText = GetTags(SiteText, "div", "class", "online-item*", "innerHTML")
        SiteText = GetTags(SiteText, "span", "class", "right", "DeleteTags")
        'GetTags(SiteText, "td", "class", "online-day", "data-day")
        TextBox3.Text = TextBox3.Text & vbNewLine & SiteText & vbNewLine
        StartDate = Format(DateAdd("d", 3, StartDate), "dd.mm.yyyy") '���� 3 ���
    Wend
    TextBox3.Text = Replace(TextBox3.Text, "%~$", vbNewLine)
    MsgBox ("������!")
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
    If InStr(Str, "#NL") > 0 Then '������� ������� ��������, ����� ����������, ����� ����
        CheckSuperChar = Replace(Str, "#NL", vbNewLine)
    End If
End Function

Private Function �����������������(AStr, BStr)
    cmdString = "�������� � �������� '#�' �� '#�'"
    If AStr = "" Then
        AStr = InputBox("������� ��� ����� ��������. ��������, ���� ����� �������� '#�' �� '*�', �� ������ ������� '#A'", "�������� � �������� � �� �", "")
    Else
        AStr = InputBox("������� ��� ����� ��������. ��������, ���� ����� �������� '#�' �� '*�', �� ������ ������� '#A'", "�������� � �������� � �� �", AStr)
    End If
    If AStr = "" Then
        ����������������� = ""
        Exit Function
    End If
    
    If BStr = "" Then
        BStr = InputBox("������� �� ��� ��������. ��������, ���� ����� �������� '#�' �� '*�', �� ������ ������� '*�'", "�������� � �������� � �� �", "#Tab")
    Else
        BStr = InputBox("������� �� ��� ��������. ��������, ���� ����� �������� '#�' �� '*�', �� ������ ������� '*�'", "�������� � �������� � �� �", BStr)
    End If

    cmdString = Replace(cmdString, "#�", AStr)
    cmdString = Replace(cmdString, "#�", BStr)
    ����������������� = cmdString
End Function

Private Sub CommandButton2_Click() '���������� ������
    cmd = ListBox1.List(ListBox1.ListIndex)
    cmdString = cmd
    If cmdString = "�������� � �������� '#�' �� '#�'" Then
        cmdString = �����������������("", "")
        If cmdString = "" Then Exit Sub
    ElseIf cmdString = "��������� �� ���������: '#�'" Then
        '"�������� � ������������ '#�' �� '#�'"
        AStr = InputBox("������� ������ �� ��������. ��������, https://pikabu.ru/", "��������� �� ���������", "")
        If AStr = "" Then Exit Sub
        cmdString = Replace(cmdString, "#�", AStr)
    ElseIf cmdString = "�������� HTML ����: '#�', '#�', '#�', '#�'" Then
        '"�������� HTML ����: '#�', '#�', '#�', '#�'". 'TagName', 'AttrName', 'AttrValue', 'Result'
        HTMLHelpStr = vbNewLine & "��������, ���� div id=" & B & "mod-lists" & B & ", � ����� ��� ������� (innerHTML):" & vbNewLine & _
        "����� � �������� '#�' ���� ������ 'div' (��� �������), � �������� '#�' - 'id', '#�' - 'mod-lists', '#�' - 'innerHTML 1')." & vbNewLine & _
        "������ 1 ����� innerHTML ��������, ��� ���� ����� ������� ��������� ����� �����, - ������ ������� ������ ������"
        AStr = InputBox("������� '#�' (TagName)." & HTMLHelpStr, "�������� HTML ����: '#�', '#�', '#�', '#�'. ������� '#�'", "")
        If AStr = "" Then Exit Sub
        BStr = InputBox("������� '#�' (AttrName)." & HTMLHelpStr, "�������� HTML ����: '#�', '#�', '#�', '#�'. ������� '#�'", "")
        VStr = InputBox("������� '#�' (AttrValue)." & HTMLHelpStr, "�������� HTML ����: '#�', '#�', '#�', '#�'. ������� '#�'", "")
        GStr = InputBox("������� '#�' (Result)." & HTMLHelpStr, "�������� HTML ����: '#�', '#�', '#�', '#�'. ������� '#�'", "innerHTML")
        
        cmdString = Replace(cmdString, "#�", AStr)
        cmdString = Replace(cmdString, "#�", BStr)
        cmdString = Replace(cmdString, "#�", VStr)
        cmdString = Replace(cmdString, "#�", GStr)
    ElseIf cmdString = "������������ � ��������������" Then
        '"������������ � ��������������" '��������� ����� �� ������������� � ��������������
        cmdString = "������������ � ��������������"
    ElseIf cmdString = "��. �������� '#������' �� '#�'" Then
        '"��. �������� '#������' �� '#�'" RegExp
        cmdString = "��. �������� '#������' �� '#�'"
        AStr = InputBox("������� ������ ����������� ���������.", "���������� ���������. �������� '#������' �� '#�'", "")
        If AStr = "" Then Exit Sub
        BStr = InputBox("�������, �� ��� �������� ������.", "���������� ���������. �������� '#������' �� '#�'", "")
        cmdString = Replace(cmdString, "#������", AStr)
        cmdString = Replace(cmdString, "#�", BStr)
    ElseIf cmdString = "��������. ��������� ���� �����" Then
        '��������. ��������� ���� �����
        cmdString = "��������. ��������� ���� �����"
    ElseIf cmdString = "��������. ��������� ���������� �����" Then
        '��������. ��������� ���� �����
        cmdString = "��������. ��������� ���������� �����"
    Else
        MsgBox ("�� �������. ������� �� ���� ����������.")
    End If
    ��������������� (cmdString)
    'If MsgBox("��������� ����������?", vbYesNo, "���������") = vbYes Then
End Sub

Private Sub CommandButton3_Click()
    ��������������
End Sub

Private Sub CommandButton4_Click() '���������� ������
    TextBox3.Text = ""
    Dim AStrIndex, BStrIndex As Integer
    For i = 0 To ListBox2.ListCount - 1
        cmd = ListBox2.List(i)
        cmdString = cmd
        '"�������� � �������� '#�' �� '#�'"
        If InStr(cmdString, "�������� � �������� ") > 0 Then
            AStrIndex = InStr(cmdString, "'")
            AStr = Mid(cmdString, AStrIndex + 1, InStr(AStrIndex + 1, cmdString, "'") - AStrIndex - 1)
            BStrIndex = InStr(cmdString, "�� '") + Len("�� '") - 1
            �Str = Mid(cmdString, BStrIndex + 1, InStr(BStrIndex + 1, cmdString, "'") - BStrIndex - 1)
            
            AStr = CheckSuperChar(AStr)
            �Str = CheckSuperChar(�Str)
            If TextBox3.Text = "" Then
                TextBox3.Text = TextBox2.Text
                TextBox3.Text = Replace(TextBox3.Text, AStr, �Str)
            Else
                TextBox3.Text = Replace(TextBox3.Text, AStr, �Str)
            End If
        End If
        '"��������� �� ���������: '#�'"
        If InStr(cmdString, "��������� �� ���������: ") > 0 Then
            AStrIndex = InStr(cmdString, "'")
            AStr = Mid(cmdString, AStrIndex + 1, InStr(AStrIndex + 1, cmdString, "'") - AStrIndex - 1)
            TextBox2.Text = GetHTTPResponse(AStr)
            If TextBox3.Text = "" Then
                TextBox3.Text = TextBox2.Text
                TextBox3.Text = Replace(TextBox3.Text, AStr, �Str)
            Else
                TextBox3.Text = Replace(TextBox3.Text, AStr, �Str)
            End If
        End If
        '"�������� HTML ����: '#�', '#�', '#�', '#�'"
        If InStr(cmdString, "�������� HTML ����: ") > 0 Then
            AStrIndex = InStr(cmdString, "'")
            AStr = Mid(cmdString, AStrIndex + 1, InStr(AStrIndex + 1, cmdString, "'") - AStrIndex - 1)
            BStrIndex = InStr(cmdString, ", '") + Len(", '") - 1
            �Str = Mid(cmdString, BStrIndex + 1, InStr(BStrIndex + 1, cmdString, "'") - BStrIndex - 1)
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
        
        '"������������ � ��������������"
        If InStr(cmdString, "������������ � ��������������") > 0 Then
            TextBox2.Text = TextBox3.Text
            'ListBox2.Clear
        End If
        
        '"��. �������� '#������' �� '#�'"
        If InStr(cmdString, "��. �������� '") > 0 Then
            AStrIndex = InStr(cmdString, "'")
            AStr = Mid(cmdString, AStrIndex + 1, InStr(AStrIndex + 1, cmdString, "'") - AStrIndex - 1)
            BStrIndex = InStr(cmdString, "�� '") + Len("�� '") - 1
            �Str = Mid(cmdString, BStrIndex + 1, InStr(BStrIndex + 1, cmdString, "'") - BStrIndex - 1)
            
            AStr = CheckSuperChar(AStr)
            �Str = CheckSuperChar(�Str)
            Set objRegExp = CreateObject("VBScript.RegExp")
            objRegExp.Global = True
            objRegExp.MultiLine = True
            objRegExp.Pattern = AStr
            If TextBox3.Text = "" Then
                TextBox3.Text = TextBox2.Text
                TextBox3.Text = objRegExp.Replace(TextBox3.Text, �Str)
            Else
                TextBox3.Text = objRegExp.Replace(TextBox3.Text, �Str)
            End If
        End If
        
        '��������. ��������� ���� �����
        If InStr(cmdString, "��������. ��������� ���� �����") > 0 Then
            TextBox2.Text = ActiveDocument.content.Text
        End If
        
        '��������. ��������� ���������� �����
        If InStr(cmdString, "��������. ��������� ���������� �����") > 0 Then
            TextBox2.Text = Selection.Text
        End If
    Next i
    'If MsgBox("��������� ����������?", vbYesNo, "���������") = vbYes Then
End Sub

Private Sub CommandButton5_Click() '��������� ��������
    Dim oFD As FileDialog
    Dim X, lf As Long
    '��������� ���������� ������ �� ��������� �������
    Set oFD = Application.FileDialog(msoFileDialogSaveAs)
    With oFD '���������� �������� ��������� � �������
    '��� �� ����� ��� oFD
    'With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .FilterIndex = 13
        .Title = "��������� �������� (Word Diary Algorithm files)" '��������� ���� �������
        .InitialFileName = ActiveDocument.Path + "\" & "Word Diary Algorithm.txt" '��������� ����� ����������� � ��� ����� �� ���������
        .InitialView = msoFileDialogViewDetails '��� ����������� ����(�������� 9 ���������)
        If oFD.Show = 0 Then Exit Sub '���������� ������
        '���� �� ��������� ��������� � ������� ������
        For lf = 1 To .SelectedItems.Count
            Path = .SelectedItems(lf) '��������� ������ ���� � �����
            Call SaveLoadListbox(ListBox2, Path, "save")
            '����� ����� ��� Path
            'Workbooks.Open .SelectedItems(lf)
        Next
    End With
End Sub

Private Sub CommandButton6_Click() '������� ��������
    Dim oFD As FileDialog
    Dim X, lf As Long
    '��������� ���������� ������ �� ��������� �������
    Set oFD = Application.FileDialog(msoFileDialogFilePicker)
    With oFD '���������� �������� ��������� � �������
    '��� �� ����� ��� oFD
    'With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = "������� ����� �������" '��������� ���� �������
        .Filters.Clear '������� ������������� ����� ���� ������
        .Filters.Add "Word Diary Algorithm files", "*.wda*;*.txt*", 1 '������������� ����������� ������ ������ ������ Excel
        .Filters.Add "Text files", "*.txt", 2 '��������� ����������� ������ ��������� ������
        .FilterIndex = 1 '������������� ��� ������ �� ��������� - Text files(��������� �����)
        .InitialFileName = ActiveDocument.Path + "\" & "��������.wda" '��������� ����� ����������� � ��� ����� �� ���������
        .InitialView = msoFileDialogViewDetails '��� ����������� ����(�������� 9 ���������)
        If oFD.Show = 0 Then Exit Sub '���������� ������
        '���� �� ��������� ��������� � ������� ������
        For lf = 1 To .SelectedItems.Count
            Path = .SelectedItems(lf) '��������� ������ ���� � �����
            Call SaveLoadListbox(ListBox2, Path, "load")
            '����� ����� ��� �
            'Workbooks.Open .SelectedItems(lf)
        Next
    End With
End Sub

Private Sub CommandButton7_Click() ' �������� �������
    If ListBox2.ListIndex = -1 Then Exit Sub
    CmdStrPast = ListBox2.List(ListBox2.ListIndex)
    CmdStr = InputBox("��������� �������", "��������� �������", CmdStrPast)
    If CmdStr <> "" Then
        ListBox2.List(ListBox2.ListIndex) = CmdStr
        CommandButton4_Click
    End If
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub CommandButton8_Click()
    If MsgBox("������� ��� �������", vbYesNo, "������������� ��������") = vbYes Then
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
    'Button ���=2
    If Button = 2 Then CreateDisplayPopUpMenu
End Sub

Private Sub UserForm_Initialize()
    ListBox1.AddItem "�������� � �������� '#�' �� '#�'"
    ListBox1.AddItem "��������� �� ���������: '#�'"
    ListBox1.AddItem "�������� HTML ����: '#�', '#�', '#�', '#�'" ''#TagName', '#AttrName', '#AttrValue', '#Result'"
    ListBox1.AddItem "������������ � ��������������"
    ListBox1.AddItem "��. �������� '#������' �� '#�'"
    ListBox1.AddItem "��������. ��������� ���� �����"
    ListBox1.AddItem "��������. ��������� ���������� �����"
End Sub

Private Sub ���������������(cmd)
    ListBox2.AddItem (Trim(Str(ListBox2.ListCount + 1)) & ". " & cmd)
    CommandButton4_Click
End Sub

Private Sub ��������������()
    If ListBox2.ListIndex = -1 Then Exit Sub
    If MsgBox("������� ��������� �������:" & vbNewLine & ListBox2.List(ListBox2.ListIndex), vbYesNo, "��������� ������") = vbYes Then
        ListBox2.RemoveItem (ListBox2.ListIndex)
    End If
    CommandButton4_Click
End Sub
