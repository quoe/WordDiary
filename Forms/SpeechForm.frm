VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SpeechForm 
   Caption         =   "����� ������ � �������"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8400.001
   OleObjectBlob   =   "SpeechForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SpeechForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub GetVoices()
    Dim WshShell As Object
    Set WshShell = CreateObject("wscript.Shell")
    FilePath = TextBox1.Value
    filePrm = "-l"
    Set WshExec = WshShell.Exec("""" & FilePath & """ " & filePrm)
    With WshExec.StdIn
        '.WriteLine "chcp 1251"
        '.WriteLine """" & filePath & """" & filePrm
        '.WriteLine "exit"
    End With
'    Set s = WshExec.StdOut
'    Do While Not s.AtEndOfStream
'        MsgBox s.ReadLine & "LOL"
'    Loop
    Do While Not WshExec.StdOut.AtEndOfStream
        S = WshExec.StdOut.ReadLine
        If S <> " " Then
            ComboBox1.AddItem S
        Else
            ComboBox1.AddItem S
        End If
    Loop
    ComboBox1.RemoveItem 0
    If ComboBox1.ListCount > 0 Then ComboBox1.ListIndex = 0
End Sub

Private Sub SaveLoadComboBox(ByVal plstLB As ComboBox, _
ByVal pstrFileName As String, _
ByVal pstrSaveOrLoad As String)
    Dim strListItems As String
    Dim i As Long
    sel = plstLB.ListIndex
    Select Case pstrSaveOrLoad
        Case "save"
        Open pstrFileName For Output As #1
        For i = 0 To plstLB.ListCount - 1
            plstLB.ListIndex = i
            Print #1, plstLB.List(i)
        Next
        Print #1, sel
        Close #1
        plstLB.ListIndex = sel
        
       Case "load"
       plstLB.Clear
        Open pstrFileName For Input As #1
        While Not EOF(1)
          Line Input #1, strListItems
          plstLB.AddItem strListItems
        Wend
        Close #1
        plstLB.ListIndex = plstLB.List(plstLB.ListCount - 1)
        plstLB.RemoveItem plstLB.ListCount - 1
    End Select
End Sub

Private Sub ComboBox1_Change()

End Sub

Private Sub CommandButton1_Click()
    ComboBox1.Clear
    GetVoices
End Sub

Private Sub CommandButton2_Click()
    FilePath = TextBox1.Text
    filePrm = " -k"
    RunFileFromCMD FilePath, filePrm, 0
End Sub

Private Function GetVoice()
    Voice = Trim(ComboBox1.Text)
    If Voice = "" Then
        MsgBox "��������� ������ � ������ �������� ���������� ������", , "�� ��������� ������"
        GetVoice = ""
    End If
    If InStr(Voice, "SAPI ") > 0 Then
        MsgBox "�������� �������� ������ � ����������, � �� '" & Voice & "'", , "�� ������ �����"
        GetVoice = ""
    End If
    GetVoice = Voice
End Function

Private Sub CommandButton3_Click()
    Voice = GetVoice
    If Voice = "" Then Exit Sub
    FilePath = TextBox1.Text
    Text = TextBox2.Text
    filePrm = ""
    'filePrm = "-tray" '  �������� ������ ��������� � ������� �����������
    filePrm = filePrm + " -n """ & "" & Voice & """" ' n - ������� �����
    filePrm = filePrm + " -t """ & "" & Text & """" ' t - ������������ ����� �� ��������� ������
    RunFileFromCMD FilePath, filePrm, 0
End Sub

Private Sub CommandButton4_Click()
    Voice = GetVoice
    If Voice = "" Then Exit Sub
    FilePath = TextBox1.Text
    Text = TextBox2.Text
    filePrm = ""
    'filePrm = "-tray" '  �������� ������ ��������� � ������� �����������
    filePrm = filePrm + " -n """ & "" & Voice & """" ' n - ������� �����
    userPrm = Trim(TextBox3.Value)
    filePrm = filePrm & " " & userPrm
    'filePrm = filePrm + " -t """ & "" & Text & """" ' t - ������������ ����� �� ��������� ������
    RunFileFromCMD FilePath, filePrm, 0
End Sub

Private Sub CommandButton5_Click()
    Voice = GetVoice
    If Voice = "" Then Exit Sub
    FilePath = TextBox1.Text
    filePrm = ""
    'filePrm = filePrm + " -tray"
    filePrm = filePrm + " -c" ' c - ������������ ����� �� ������ ������.
    filePrm = filePrm + " -n """ & "" & Voice & """" ' n - ������� �����
    RunFileFromCMD FilePath, filePrm, 0
End Sub

Private Sub CommandButton6_Click()
    CommandButton5_Click
End Sub

Private Sub CommandButton7_Click()
    Voice = GetVoice
    If Voice = "" Then Exit Sub
    FilePath = TextBox1.Text
    Text = Selection.Text
    filePrm = ""
    'filePrm = "-tray" '  �������� ������ ��������� � ������� �����������
    filePrm = filePrm + " -n """ & "" & Voice & """" ' n - ������� �����
    filePrm = filePrm + " -t """ & "" & Text & """" ' t - ������������ ����� �� ��������� ������
    RunFileFromCMD FilePath, filePrm, 0
End Sub



Private Sub CommandButton8_Click()
    Voice = GetVoice
    If Voice = "" Then Exit Sub
    CfgSpeechFilePath = �����������(TextBox1.Text) + CfgSpeechFileName + ".cfg"
    If ��������������(CfgSpeechFilePath) Then
        If MsgBox("���� ��� ����������. ��������?", vbYesNo, "���� ��� ����������. ��������?") = vbYes Then
            Open CfgSpeechFilePath For Output As #1
            Print #1, "-n " & Voice
            Close #1
            CommandButton9_Click
        End If
    End If
End Sub

Private Sub CommandButton9_Click()
    Voice = GetVoice
    If Voice = "" Then Exit Sub
    Path = ActiveDocument.Path + "\" + VoiceFileName + ".txt"
    Call SaveLoadComboBox(ComboBox1, Path, "save")
End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub UserForm_Activate()
    TextBox1.Value = BalabolkaConsoleFilePath
'ActiveDocument.Name
'ActiveDocument.FullName
'ActiveDocument.Path
    Path = ActiveDocument.Path + "\" + VoiceFileName + ".txt" ' TegsFileName �� ��������� ������
    'If Not ��������������(Path) Then ��������������� Path, ""
    If ��������������(Path) Then Call SaveLoadComboBox(ComboBox1, Path, "load")
End Sub

Private Sub UserForm_Click()

End Sub
