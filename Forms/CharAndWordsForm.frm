VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CharAndWordsForm 
   Caption         =   "������ � ��������� � �������"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9645
   OleObjectBlob   =   "CharAndWordsForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CharAndWordsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Text ' ��������� ��������� (��������, � Like)
Option Explicit

Sub �������_���������_����()
    '����������� ����� ������������� ������ �����
    Const OsnovaLen% = 2    ' =2 ��� �������� ��������� ���� � ������� �����������, =100 ��� �������� ��������� ���������� ���� � ������ ����������
    Const LyuboyPoryadok As Boolean = True ' = True - ��� ������������� ������� ���� � ���������, = False - ������ ��� ������� ������� ���� � ���������
    Const IgnorirovatZnakiVnutriPredlozheniya As Boolean = True '  =True  - �� ����������� ����� ���������� ������ �����������, ����������� ����� � ���������
    Const IgnorirovatPredlogi As Boolean = True '  =True  - ��������, ������������� � ������� EtoPredlog, �� ���������� � ������, � ������������ ����� �������.
    Const IgnorirovatSoyuzy As Boolean = True '  =True  - �����, ������������� � ������� EtoSoyuz, �� ���������� � ������, � ������������ ����� �������.
    Const PechatSPonizheniemChastoty As Boolean = True ' ����� ����������� � ������� ��������� ������� ��� ����. False - � ���������� �������.
    Const freq% = 2         ' ����������� ������� ��������� ��������� ����, � ������� ��� �����������
    Dim Slova               ' ���������� ��� ������� ����
    Dim PredPrediduschee$, SlovoPrediduschee$, SlovoPrediduschee2$, Slovo$ ' ���������� ��� ������ ���������������, ����������� � ���������� "�����"
    Dim i&, j&, LenS%, UbOk%, iCr%, dic As Object, Key, S$, Z  ' ��������������� ����������
    ' ������ ��������� ���������, ������ ���������� �� 3 ��������� �� 1 ��������� :
    Dim ok: ok = Split("��� ��� ��� ��� ��� ��� ��� ��� ��� ��� ��� ��� ��� ��� ��� ��� �� �� �� ax �� �� � �� �� ex �� �� �� �� �� �� �� �� �� �� �� �� �� �� �� �� �� �� �� �� �� �� �� � �� �� �� �� �� �� �� �� �� �� �� �� � � � � � � � � � �", " ")
    UbOk = UBound(ok)
    ' ������ ���� ���������
    Dim k%(): ReDim k(UbOk): For i = 0 To UbOk: k(i) = Len(ok(i)): Next
    ' ������ ����������� ���� ����, � ������� ����� �������� ���������
    Dim minLen%(): ReDim minLen(UbOk):  For i = 0 To UbOk: minLen(i) = k(i) + OsnovaLen: Next
    ' ������ ������ ����������. � ������� �� ������� �������_����� ��� ������� ����������� ������ � ������ ������ ���������� �� �� ����� �������� � ������� ����� ������ 1 ������ � ��������� ���. ������ ����� ������ ������ ���������� �� ������� ������� �������.
    Dim Punkt: Punkt = Split(". . ? ! , ; : ( )", " "): Punkt(0) = Chr$(13) ' vbCr - ������ ������
    S = ActiveDocument.Range.Text
    ' ������� ������ �� ������������ �������� � �������������� ��������
    S = Replace(S, "�", ""): S = Replace(S, "�", ""): S = Replace(S, """", ""): S = Replace(S, "'", "")
    S = Replace$(S, Chr$(10), " "): S = Replace$(S, Chr$(9), " "): S = Replace(S, Chr$(160), " ") ' ����������� ������
    S = Replace(S, Chr$(150), "-"): S = Replace(S, Chr$(151), "-"): S = Replace(S, Chr$(30), "-")
    S = Replace(S, "�", "."): S = Replace(S, "...", ".")
    ' ���������� �������� ������ ������ ����������
    For i = 0 To UBound(Punkt): S = Replace$(S, Punkt(i), " " & Punkt(i) & " "): Next
    For i = 1 To 5: S = Replace$(S, "  ", " "): Next '������������ ������ 2-��� �������� �� 1
    S = LCase$(S) ' ��������� ���� ����� � ��������� �����
    Slova = Split(S, " ") ' ������� ���������� ���������� � ������� � ������, ����������� - ������.
    S = ""
    Set dic = CreateObject("Scripting.Dictionary") ' ������ ������� ��� ��������� � �������� ��������� ����
    With dic
        .CompareMode = 1   ' ���������� ���������������� � �������� � �������.
        For i = 0 To UBound(Slova)
            Slovo = Slova(i) ' ����� ��������� ����� ��� ������ �������� �� �������
            If EtoSoyuz(PredPrediduschee, SlovoPrediduschee, SlovoPrediduschee2, Slovo, IgnorirovatSoyuzy) And IgnorirovatSoyuzy Then
                iCr = 0
            ElseIf EtoPredlog(PredPrediduschee, SlovoPrediduschee, SlovoPrediduschee2, Slovo, IgnorirovatPredlogi) And IgnorirovatPredlogi Then
                iCr = 0
            Else
                S = Left$(Slovo, 1) ' �������� 1 ����� ������, ���� ��� �����, �� ���������� � � ���������.
                Select Case S
                    Case "" '������ �� ������, 2 ������� ���� ������
                    Case ".", "?", "!"   ' ����� �����������, ����� ����������� ����� ������� �� ������� ����������.
                        SlovoPrediduschee = "" ' ���������� ����� ��������� �� ����� � ��������� �� ��������� ������.
                        PredPrediduschee = ""
                        iCr = 0
                    Case ",", ";", ":", "(", ")", "-"
                        If Not IgnorirovatZnakiVnutriPredlozheniya Then
                            SlovoPrediduschee = ""  ' ���������� ����� ��������� �� ����� � ��������� �� ��������� ������.
                            PredPrediduschee = ""
                        End If
                        iCr = 0
                    Case Chr$(13)  ' vbCr - ������ ������
                        iCr = iCr + 1 ' ������� ������ ������ �������� ������
                        If iCr > 1 Then
                            SlovoPrediduschee = "" ' �����, ����������� ������ �������, ��������� �� ����� ���������.
                            PredPrediduschee = ""
                        End If
                    Case "a" To "z", "�" To "�", "�", "0" To "9", "�", "�", "&", "$", "�"
                        If OsnovaLen < 100 Then
                            If EtoSoyuz("", "", SlovoPrediduschee2, Slovo, False) Then
                            ElseIf EtoPredlog("", "", SlovoPrediduschee2, Slovo, False) Then
                            ElseIf Not Isklyucheniye(Slovo) Then
                                ' ���������� � ������� ���������
                                LenS = Len(Slovo)
                                S = LCase(Slovo)
                                For j = 0 To UbOk
                                    If LenS >= minLen(j) Then
                                        If Right$(S, k(j)) = ok(j) Then
                                            Slovo = Left$(Slovo, LenS - k(j))
                                            Exit For
                                        End If
                                    End If
                                Next j
                            End If
                        End If
                        If SlovoPrediduschee <> "" Then
                            .Item(SlovoPrediduschee & " " & Slovo) = .Item(SlovoPrediduschee & " " & Slovo) + 1
                        End If
                        PredPrediduschee = SlovoPrediduschee
                        SlovoPrediduschee = Slovo
                        SlovoPrediduschee2 = Slovo
                        iCr = 0
                End Select
            End If
        Next
        Erase Slova
        If LyuboyPoryadok Then
            ' ������������ ���� ���� � ������ � �������� �������
            For Each Key In .Keys
                If .Item(Key) > 0 Then
                    Z = Split(Key, " ")
                    S = Z(1) & " " & Z(0)
                    .Item(Key) = .Item(Key) + .Item(S)
                    .Item(S) = 0
                End If
            Next
        End If
        DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents ' ������ � �������� �� W�rd
        Application.ScreenUpdating = False
        Documents.Add: ActiveWindow.ActivePane.View.Type = wdPrintView  ' ��� "�������� ��������"
        For Each Key In .Keys
            If .Item(Key) >= freq Then
                If OsnovaLen < 100 Then
                    Z = Split(Key, " ")
                    S = ""
                    For i = 0 To 1
                        S = S & Z(i)
                        If EtoSoyuz("", "", "", Z(i), False) Then
                        ElseIf EtoPredlog("", "", "", Z(i), False) Then
                        ElseIf Isklyucheniye(Z(i)) Then
                        Else
                           S = S & IIf(Len(Z(i)) >= OsnovaLen, "-", "")
                        End If
                        S = S & " "
                    Next
                    S = RTrim$(S)
                Else
                    S = Key
                End If
                Selection.TypeText S & Chr$(9) & .Item(Key) & Chr$(13)
            End If
        Next
    End With
    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents
    Set Key = Nothing: Set dic = Nothing
    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents
    With Selection
        If PechatSPonizheniemChastoty Then
            ' ���������� �� ������� � ����������.
            .Sort FieldNumber:="����� 2", SortFieldType:=wdSortFieldNumeric, SortOrder:=wdSortOrderDescending, Separator:=wdSortSeparateByTabs
            .Collapse Direction:=wdCollapseEnd: .Delete
        Else
            ' ���������� �� ��������.
            .Sort FieldNumber:="�������", SortFieldType:=wdSortFieldAlphanumeric, SortOrder:=wdSortOrderAscending
            .Collapse: .Delete
        End If
    End With
    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents
    With ActiveDocument.PageSetup.TextColumns
        .SetCount NumColumns:=3 ' � 3 �������.
        .LineBetween = True     ' ����� ����� �������.
    End With
    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents
    ActiveDocument.Paragraphs.TabStops.Add CentimetersToPoints(4), wdAlignTabRight
End Sub
Function EtoSoyuz(PredPrediduschee, SlovoPrediduschee, SlovoPrediduschee2, Slovo, Ignorirovat As Boolean) As Boolean  ' �������� ��������� ������
    ' ����� �� 3-� � ����� ���� ����� ��������.
    Select Case SlovoPrediduschee2 & " " & Slovo
        Case "� ��������", "� ������", "� �����", "� ��", "� �", "��������� ����", "� ����������", "� �����", "� ���� ����", "��� ��", "�� ��", "�� ��������", "�� ���", "�� ���", "�� �", "����� ���", "��� ����", "���� ��", "���� ��", "� ������", "� ������", "� �������", "� ������", "� ���-����", "� ��-����", "� �������������", "� ��", "� �����", "� ���", "� ��������", "��-�� ����", "��� �����", "��� �����", "��� ������", "��� ������", "���� ��", "����� ����", "����� �����", "���� ��", "���� ������"
            EtoSoyuz = True
            If Ignorirovat Then SlovoPrediduschee = PredPrediduschee: SlovoPrediduschee2 = Slovo
            Exit Function
        Case "����� ���", "�� �������", "�� ��", "�� ������", "�������� ��", "���������� ��", "�������� ��", "�� �", "�� ����", "����� ���", "�� ����", "�� �������", "������� ����", "���� ��", "����� ����", "������ ���", "��� ����", "��� ���", "��� �������", "���� ����", "������ ���", "� ���", "��� ���", "��� �", "��� ��", "��� ���", "��� ���", "��� �����", "��� ���", "����� ���", "���� ���", "�� ����", "�� ��", "�� ���", "������ ��", "������ ���", "������ ����", "������ ����", "��� ���"
            EtoSoyuz = True
            If Ignorirovat Then SlovoPrediduschee = PredPrediduschee: SlovoPrediduschee2 = Slovo
            Exit Function
    End Select
    EtoSoyuz = True
    Select Case Slovo
        Case "�", "�����", "����", "�����", "��������", "�����", "����������", "��", "����", "����", "��", "����"
        Case "�����", "����", "�����", "�����", "����", "�����", "�", "���", "��� ", "����", "���", "���-��"
        Case "�����", "����", "��", "����", "����", "������", "��", "��", "������", "��������", "������", "������"
        Case "����", "��������", "������", "���������", "������", "������", "������", "������", "������", "�����"
        Case "���", "�������", "������", "�������", "�����", "���", "����", "��", "����", "������", "�����", "����"
        Case "���", "����", "���", "����", "�����"
        Case Else: EtoSoyuz = False
    End Select
    If Ignorirovat Then
        PredPrediduschee = SlovoPrediduschee
        SlovoPrediduschee2 = Slovo
    End If
End Function
Function EtoPredlog(PredPrediduschee, SlovoPrediduschee, SlovoPrediduschee2, Slovo, Ignorirovat As Boolean) As Boolean  ' �������� ��������� ���������
    Select Case SlovoPrediduschee2 & " " & Slovo
        Case "� �������", "� �����������", "�������� ��", "����� �", "�� �������", "�� �������", "� �����", "� �����", "�� ������", "�� ������", "� ����", "� �������", "� ����������", "�� ����������"
            EtoPredlog = True
            If Ignorirovat Then SlovoPrediduschee = PredPrediduschee: SlovoPrediduschee2 = Slovo
            Exit Function
    End Select
    EtoPredlog = True
    Select Case Slovo
        Case "��" ' ���� ��� � �� �������, �� ��� ������� ������ �� ������� "��", ������� �� ����.
        Case "�", "�", "�", "�", "�", "��", "��", "��", "��", "��", "��", "���", "���", "��-���"
        Case "���", "���", "���", "����", "�����", "��-��", "��-���", "��-��", "����", "�����"
        Case "�����", "�����", "������", "������", "�����", "������", "�������", "�����������", "������"
        Case "����", "������", "�������", "�����", "�������", "������", "�����"
        Case "������", "�����", "������", "�����", "������", "�����", "������", "������", "���������"
        Case "��������", "���������", "��������", "���������", "��������", "�������", "������", "����������", "���������", "�����", "������"
        Case Else: EtoPredlog = False
    End Select
    If Ignorirovat Then
        PredPrediduschee = SlovoPrediduschee
        SlovoPrediduschee2 = Slovo
    End If
End Function
Function Isklyucheniye(Slovo) As Boolean
' �������� ����, � ������� �� ����� ���������� ���������, ������� �� ���������.
    Isklyucheniye = True
    Select Case Slovo
        ' � 1-�� ����� "Case"  ���������� ��� ���������� ��������� ������ �� ���� ����.
        ' ���������� � 1-2 ����� � ����������, ����������� � ������������ ���������� � ������� ������� �� ����.
        Case "���������", "���", "��", "�����", "�����", "����", "�����", "��������", "����������", "��������", "�������", "����", "������", "�����", "����", "�����", "������"
        Case "��������"
        Case Else: Isklyucheniye = False
    End Select
End Function

Public Function funcSortKeysByLengthDesc(dctList As Object) As Object
    Dim arrTemp() As String
    Dim curKey As Variant
    Dim itX As Integer
    Dim itY As Integer

    'Only sort if more than one item in the dict
    If dctList.Count > 1 Then

        'Populate the array
        ReDim arrTemp(dctList.Count)
        itX = 0
        For Each curKey In dctList
            arrTemp(itX) = curKey
            itX = itX + 1
        Next

        'Do the sort in the array
        For itX = 0 To (dctList.Count - 2)
            For itY = (itX + 1) To (dctList.Count - 1)
                If Len(arrTemp(itX)) < Len(arrTemp(itY)) Then
                    curKey = arrTemp(itY)
                    arrTemp(itY) = arrTemp(itX)
                    arrTemp(itX) = curKey
                End If
            Next
        Next

        'Create the new dictionary
        Set funcSortKeysByLengthDesc = CreateObject("Scripting.Dictionary")
        For itX = 0 To (dctList.Count - 1)
            funcSortKeysByLengthDesc.Add arrTemp(itX), dctList(arrTemp(itX))
        Next

    Else
        Set funcSortKeysByLengthDesc = dctList
    End If
End Function

Sub �������_����(Selected As Boolean, frecConst As Integer, Registr As Integer, Sort As Boolean)  ' cyberforum.ru > KoGG > Sasha_Smirnov
Dim freq: freq = frecConst   ' ������� ��������� ������� oWord (����� ��� �����), ������ ������� ��� �����������
Dim oWord As Range, dic As Object, vX As Variant, S As String, TextWords As Object
Dim t, TextStr As String             ' ���������� ��� ������ ���������� "�����" (oWord.Text)
    Set dic = CreateObject("Scripting.Dictionary")
            Application.ScreenUpdating = False

With dic
        .CompareMode = Registr        ' ��������� ���������������� � �������� � �������
    If Selected = True Then
        Set TextWords = Selection.Range.Words
    Else
        Set TextWords = ActiveDocument.Range.Words
    End If
    
    For Each oWord In TextWords
        t = oWord.Text
        ' ������ (� ����� ��� �������) ������������ ������� �������
        If Right(t, 1) = ChrW(160) Then t = Mid(t, 1, Len(t) - 1): t = t & " " ': MsgBox t & vbCr & _
                                                                            "Len(t) = " & Len(t)
    ' ���������� � "�����" (oWord.Text) �������, ���� ��� ���, ����� ���� = ����� � �������� � ���
        If Right(t, 1) Like "[!" & Chr(10) & Chr(11) & Chr(13) & Chr(30) & Chr(32) & "]" Then t = t & " "
        
        S = UCase(t)
        
            Select Case AscW(S) ' ��������� ������ 1-� �����
            Case 30, 33 To 90, 132, 133, 145 To 148, 168, 171, 184, 187, 1025, 1040 To 1071 ', _
            8211, 8212 ' �����, ���. ����� ���������� (� �. �. ����: 8212), ���. � ���. ����. �����
            .Item(t) = .Item(t) + 1
            End Select
            
        If AscW(t) = 8211 Then .Item(t) = .Item(t) + 1  ' ������� ������� ������� (N-dash:�)
        If AscW(t) = 8212 Then .Item(t) = .Item(t) + 1  ' ������� ������� ���� (M-dash:�)
    Next
    
    '���������� � �������� ������ ��������� ��� ������ ���� ���������� �������
    Documents.Add
    Options.CheckGrammarAsYouType = False               ' ������ ��������������� �������� �����
    Options.CheckSpellingAsYouType = False              ' ������ ���������������� �������� �����

    With ActiveDocument    '��������� �����; ��������� ����� (���� .LineNumbering.Active = True)
        With .PageSetup
    '    .LineNumbering.Active = True                    ' ��������� ����� ������� (�� �������!)
    '    .LineNumbering.RestartMode = wdRestartContinuous ' �������� ��������� �����
    '    .LineNumbering.DistanceFromText = 4 'pt
        .TopMargin = CentimetersToPoints(1)
        .BottomMargin = CentimetersToPoints(1)
        .LeftMargin = CentimetersToPoints(1.1)
        .RightMargin = CentimetersToPoints(0)
        End With
    .Paragraphs.TabStops.Add CentimetersToPoints(3.5), wdAlignTabRight ' ��� ������ 35 ��
    End With
    
    For Each vX In .Keys
        If .Item(vX) > freq Then Selection.TypeText RTrim(vX) & Chr(9) & .Item(vX) & Chr(13)
    Next
End With

Set dic = Nothing
    
    With Selection.PageSetup
    .TextColumns.SetCount NumColumns:=4                     ' � 4 �������
    .TextColumns.LineBetween = True                         ' ����� ����� �������? ��!
    .TextColumns.Spacing = CentimetersToPoints(1.9)          ' 19 �� ��� �������
    End With
    
    With Selection
        If Sort Then
            .Sort      ' ���������� ������� ���������
        End If
        .Font.Size = 10
        .Collapse
        .Delete
    End With
    ActiveDocument.UndoClear ' ������� ������ ������� �� ����� ��������� ���������
End Sub
Public Sub �����������������������������()
    '���������� � �������� ������ ��������� ��� ������ ���� ���������� �������
    Documents.Add
    Options.CheckGrammarAsYouType = False               ' ������ ��������������� �������� �����
    Options.CheckSpellingAsYouType = False              ' ������ ���������������� �������� �����

    With ActiveDocument    '��������� �����; ��������� ����� (���� .LineNumbering.Active = True)
        With .PageSetup
    '    .LineNumbering.Active = True                    ' ��������� ����� ������� (�� �������!)
    '    .LineNumbering.RestartMode = wdRestartContinuous ' �������� ��������� �����
    '    .LineNumbering.DistanceFromText = 4 'pt
        .TopMargin = CentimetersToPoints(1)
        .BottomMargin = CentimetersToPoints(1)
        .LeftMargin = CentimetersToPoints(1.1)
        .RightMargin = CentimetersToPoints(0)
        End With
    .Paragraphs.TabStops.Add CentimetersToPoints(3.5), wdAlignTabRight ' ��� ������ 35 ��
    End With
End Sub
Public Sub ��������������������������������(NColumns As Integer, Sort As Boolean)
    With Selection.PageSetup
    .TextColumns.SetCount NumColumns:=NColumns                     ' � N �������
    .TextColumns.LineBetween = True                         ' ����� ����� �������? ��!
    .TextColumns.Spacing = CentimetersToPoints(1.9)          ' 19 �� ��� �������
    End With
    
    With Selection
        If Sort Then
            .Sort      ' ���������� ������� ���������
        End If
        .Font.Size = 10
        .Collapse
        .Delete
    End With
    ActiveDocument.UndoClear ' ������� ������ ������� �� ����� ��������� ���������
End Sub

Sub ���������������������(Selected As Boolean, frecConst As Integer, Registr As Integer, Sort As Boolean)
Dim oWord As Range, vX As Variant, vX2 As Variant, S As String, TextWords As Object
Dim t, TextStr, TestStr As String             ' ���������� ��� ������ ���������� "�����" (oWord.Text)
Dim i As Integer, MaxInd As Integer
Dim dicWords As Object
Set dicWords = CreateObject("Scripting.Dictionary")
Application.ScreenUpdating = False
TextStr = ""
i = 0

If Selected = True Then
    Set TextWords = Selection.Range.Words
Else
    Set TextWords = ActiveDocument.Range.Words
End If

With dicWords
    .CompareMode = Registr        ' ��������� ���������������� � �������� � �������
    For Each oWord In TextWords
        i = i + 1
        t = oWord.Text
        ' ������ (� ����� ��� �������) ������������ ������� �������
        If Right(t, 1) = ChrW(160) Then t = Mid(t, 1, Len(t) - 1): t = t & " " ': MsgBox t & vbCr & _
                                                                            "Len(t) = " & Len(t)
    ' ���������� � "�����" (oWord.Text) �������, ���� ��� ���, ����� ���� = ����� � �������� � ���
        If Right(t, 1) Like "[!" & Chr(10) & Chr(11) & Chr(13) & Chr(30) & Chr(32) & "]" Then t = t & " "
        
        S = UCase(t)
        
        Select Case AscW(S) ' ��������� ������ 1-� �����
            Case 30, 33 To 90, 132, 133, 145 To 148, 168, 171, 184, 187, 1025, 1040 To 1071 ', _
            8211, 8212 ' �����, ���. ����� ���������� (� �. �. ����: 8212), ���. � ���. ����. �����
            If Not .Exists(t) Then
                .Item(t) = i
            Else
                i = i - 1
            End If
        End Select
            
        If AscW(t) = 8211 Then             ' ���� ������� ������� (N-dash:�)
            If Not .Exists(t) Then
                .Item(t) = i
            Else
                i = i - 1
            End If
        End If
        If AscW(t) = 8212 Then ' ���� ������� ���� (M-dash:�)
            If Not .Exists(t) Then
                .Item(t) = i
            Else
                i = i - 1
            End If
        End If
        TextStr = TextStr + CStr(.Item(t)) + " "
        If CheckBox21.Value = True Then '� ����� ������
            TextStr = TextStr + Chr(13)
        End If
    Next
    
    TextStr = Replace(TextStr, "    ", " ")
    TextStr = Replace(TextStr, "   ", " ")
    TextStr = Replace(TextStr, "  ", " ")
    ������������
    Selection.TypeParagraph
    Selection.Text = TextStr

    ����������������������������� '��������������������������������
    For Each vX In .Keys
        Selection.TypeText RTrim(vX) & Chr(9) & .Item(vX) & Chr(13)
    Next
End With

Set dicWords = Nothing
    �������������������������������� 4, Sort
End Sub

Sub ���������������������������(Selected As Boolean, frecConst As Integer, Registr As Integer, Sort As Boolean)
Dim oWord As Range, vX As Variant, vX2 As Variant, S As String, TextWords As Object
Dim t, TextStr, TestStr As String             ' ���������� ��� ������ ���������� "�����" (oWord.Text)
Dim i As Integer, MaxInd As Integer
Dim dicWords As Object
Set dicWords = CreateObject("Scripting.Dictionary")
Application.ScreenUpdating = False
TextStr = ""
i = 0

If Selected = True Then
    Set TextWords = Selection.Sentences
Else
    Set TextWords = ActiveDocument.Sentences
End If

With dicWords
    .CompareMode = Registr        ' ��������� ���������������� � �������� � �������
    For Each oWord In TextWords
        i = i + 1
        t = oWord.Text
        ' ������ (� ����� ��� �������) ������������ ������� �������
        If Right(t, 1) = ChrW(160) Then t = Mid(t, 1, Len(t) - 1): t = t & " " ': MsgBox t & vbCr & _
                                                                            "Len(t) = " & Len(t)
    ' ���������� � "�����" (oWord.Text) �������, ���� ��� ���, ����� ���� = ����� � �������� � ���
        If Right(t, 1) Like "[!" & Chr(10) & Chr(11) & Chr(13) & Chr(30) & Chr(32) & "]" Then t = t & " "
        
        S = UCase(t)
        
        Select Case AscW(S) ' ��������� ������ 1-� �����
            Case 30, 33 To 90, 132, 133, 145 To 148, 168, 171, 184, 187, 1025, 1040 To 1071 ', _
            8211, 8212 ' �����, ���. ����� ���������� (� �. �. ����: 8212), ���. � ���. ����. �����
            If Not .Exists(t) Then
                .Item(t) = i
            Else
                i = i - 1
            End If
        End Select
            
        If AscW(t) = 8211 Then             ' ���� ������� ������� (N-dash:�)
            If Not .Exists(t) Then
                .Item(t) = i
            Else
                i = i - 1
            End If
        End If
        If AscW(t) = 8212 Then ' ���� ������� ���� (M-dash:�)
            If Not .Exists(t) Then
                .Item(t) = i
            Else
                i = i - 1
            End If
        End If
        TextStr = TextStr + CStr(.Item(t)) + " "
        If CheckBox21.Value = True Then '� ����� ������
            TextStr = TextStr + Chr(13)
        End If
    Next
    
    TextStr = Replace(TextStr, "    ", " ")
    TextStr = Replace(TextStr, "   ", " ")
    TextStr = Replace(TextStr, "  ", " ")
    ������������
    Selection.TypeParagraph
    Selection.Text = TextStr

    ����������������������������� '��������������������������������
    For Each vX In .Keys
        Selection.TypeText RTrim(vX) & Chr(9) & .Item(vX) & Chr(13)
    Next
End With

Set dicWords = Nothing
    �������������������������������� 3, Sort
End Sub

Public Function ��������������(ByVal Str As String)
    Dim ArrayStr: ArrayStr = Split("�, �, �, �, �, �, �, �, �, �", ", ")
    Dim i As Long, j As Long, n As Long
    n = 0
    For i = 1 To Len(Str)
        '� ������� "Mid" ���� ��������� ������ �� ������.
        '���� ������ �������� �������, �� �������,
        '��� ��������� ����� �������.
        If Mid(Str, i, 1) Like "[�-�]" Then
            For j = 0 To UBound(ArrayStr)
                If Mid(Str, i, 1) = ArrayStr(j) Then
                    n = n + 1
                End If
            Next j
        End If
    Next i
    �������������� = n
End Function

Public Function ����������������(ByVal Str As String)
    Dim ArrayStr: ArrayStr = Split("�, �, �, �, �, �, �, �, �, �, �, �, �, �, �, �, �, �, �, �, �", ", ")
    Dim i As Long, j As Long, n As Long
    n = 0
    For i = 1 To Len(Str)
        '� ������� "Mid" ���� ��������� ������ �� ������.
        '���� ������ �������� �������, �� �������,
        '��� ��������� ����� �������.
        If Mid(Str, i, 1) Like "[�-�]" Then
            For j = 0 To UBound(ArrayStr)
                If Mid(Str, i, 1) = ArrayStr(j) Then
                    n = n + 1
                End If
            Next j
        End If
    Next i
    ���������������� = n
End Function

Public Function �������������(ByVal TextStr As String)
    Dim OutStr As String
    OutStr = TextStr
    OutStr = Replace(OutStr, "    ", " ")
    OutStr = Replace(OutStr, "   ", " ")
    OutStr = Replace(OutStr, "  ", " ")
    ������������� = OutStr
End Function

Public Function ���������������(ByVal TextStr As String)
    Dim OutStr As String
    Dim i
    OutStr = TextStr
    For i = 0 To 3
        OutStr = Replace(OutStr, vbTab & vbTab & vbTab, vbTab)
        OutStr = Replace(OutStr, vbTab & vbTab, vbTab)
        OutStr = Replace(OutStr, vbNewLine & vbTab, vbNewLine)
        OutStr = Replace(OutStr, vbCr & vbTab, vbCr)
        OutStr = Replace(OutStr, vbCrLf & vbTab, vbCrLf)
    Next i
    ��������������� = OutStr
End Function

Public Sub ��������������(ByVal HandleTextStr As String)
    Selection.Style = ActiveDocument.Styles("��������� 1") '����� ������
    Selection.TypeText HandleTextStr
    Selection.TypeParagraph
    Selection.ClearFormatting
End Sub

    
Sub ������������(Selected As Boolean, frecConst As Integer, Registr As Integer, Sort As Boolean)
Dim oWord As Range, vX As Variant, vX2 As Variant, S As String, TextWords As Object
Dim t, TextStr, TextStrSeq, TextFromDict, TestStr As String             ' ���������� ��� ������ ���������� "�����" (oWord.Text)
Dim i As Double, MaxInd As Double
Dim dicFrec As Object, dicWords As Object, dicTemp As Object
Dim WordsCount As Double
Dim WordsCountTest As Double
Set dicFrec = CreateObject("Scripting.Dictionary")
Set dicWords = CreateObject("Scripting.Dictionary")
Set dicTemp = CreateObject("Scripting.Dictionary")
Application.ScreenUpdating = False

TextStr = ""
TextStrSeq = ""
i = 0
WordsCount = 0
If Selected = True Then
    Set TextWords = Selection.Range.Words
Else
    Set TextWords = ActiveDocument.Range.Words
End If

With dicFrec '���������� ������� � ��������� ����
    .CompareMode = Registr        ' ��������� ���������������� � �������� � �������
    For Each oWord In TextWords '������� ����
        t = oWord.Text
        ' ������ (� ����� ��� �������) ������������ ������� �������
        If Right(t, 1) = ChrW(160) Then t = Mid(t, 1, Len(t) - 1): t = t & " " ': MsgBox t & vbCr & _
                                                                            "Len(t) = " & Len(t)
    ' ���������� � "�����" (oWord.Text) �������, ���� ��� ���, ����� ���� = ����� � �������� � ���
        If Right(t, 1) Like "[!" & Chr(10) & Chr(11) & Chr(13) & Chr(30) & Chr(32) & "]" Then t = t & " "
        
        S = UCase(t)
        
            Select Case AscW(S) ' ��������� ������ 1-� �����
            Case 30, 33 To 90, 132, 133, 145 To 148, 168, 171, 184, 187, 1025, 1040 To 1071 ', _
            8211, 8212 ' �����, ���. ����� ���������� (� �. �. ����: 8212), ���. � ���. ����. �����
            .Item(t) = .Item(t) + 1
            End Select
            
        If AscW(t) = 8211 Then .Item(t) = .Item(t) + 1  ' ������� ������� ������� (N-dash:�)
        If AscW(t) = 8212 Then .Item(t) = .Item(t) + 1  ' ������� ������� ���� (M-dash:�)
    Next
    
    '�������� ���� � ������ ���������
    For Each vX In .Keys
        If CInt(.Item(vX)) < frecConst Then
            .Remove (vX)
        End If
    Next
    
    dicTemp.CompareMode = Registr        ' ��������� ���������������� � �������� � �������
    For Each vX In .Keys '����� ������� ������, ����� � �� �������
        dicTemp(vX) = .Item(vX)
        WordsCountTest = CInt(.Item(vX))
        WordsCount = WordsCount + WordsCountTest '�� ���� ������� ���������� ����
    Next
End With

    i = 0
    dicWords.CompareMode = Registr        ' ��������� ���������������� � �������� � �������
    For Each vX In dicTemp.Keys '������� ������ ������������ �������, ����������, ������� ������� �� �������
        i = i + 1
        TestStr = vX
        TestStr = dicTemp.Item(vX)
        MaxInd = FindMaxInd(dicTemp.Items())
        dicWords.Item(dicTemp.Keys()(MaxInd)) = i '���������� �������� ��������� �� ������ �������. ������ ������� �� 1
        dicTemp.Remove (dicTemp.Keys()(MaxInd))
    Next
    dicTemp.RemoveAll '������� �������, ������ ��� ������ ���-�� �������

    '����� � ����� ���������� ����
    For Each oWord In TextWords '������� ����
        t = oWord.Text
        ' ������ (� ����� ��� �������) ������������ ������� �������
        If Right(t, 1) = ChrW(160) Then t = Mid(t, 1, Len(t) - 1): t = t & " " ': MsgBox t & vbCr & _
                                                                            "Len(t) = " & Len(t)
    ' ���������� � "�����" (oWord.Text) �������, ���� ��� ���, ����� ���� = ����� � �������� � ���
        If Right(t, 1) Like "[!" & Chr(10) & Chr(11) & Chr(13) & Chr(30) & Chr(32) & "]" Then t = t & " "
        
        If CheckBox20.Value = True Then '����������� ������ �������
            TextFromDict = CStr(dicFrec.Item(t) / WordsCount)
        Else
            TextFromDict = CStr(dicWords.Item(t))
        End If
        If TextFromDict = " " Or TextFromDict = "0" Then '���� ����� ������� �� ����, �� ����������� ������ ���� 0
            TextFromDict = ""
        End If
        TextStrSeq = TextStrSeq & TextFromDict & " "
        If CheckBox20.Value = True Then '����������� ������ �������
            If t = ". " Then
                TextStrSeq = TextStrSeq & Chr(13)
            End If
        End If
                
        TextStr = TextStr & t & Chr(9) & TextFromDict & " "
        If CheckBox21.Value = True Then '� ����� ������
            TextStr = TextStr & Chr(13)
        End If
    Next
    TextStr = �������������(TextStr)
    TextStrSeq = �������������(TextStrSeq)
    TextStr = ���������������(TextStr)
    TextStrSeq = ���������������(TextStrSeq)
    
    ����������������������������� '��������������������������������
    �������������� "�������"
    Selection.TypeText "�����" & Chr(9) & "������ �� �������" & Chr(9) & "�������" & Chr(9) & "�����������" & Chr(9) & "�����" & Chr(13)
    For Each vX In dicWords.Keys
        If dicWords.Item(vX) <> "" Then
            Selection.TypeText RTrim(vX) & Chr(9) & dicWords.Item(vX) & Chr(9) & dicFrec.Item(vX) & Chr(9) & CStr(dicFrec.Item(vX) / WordsCount) & Chr(9) & Len(RTrim(vX)) & Chr(13)
        End If
    Next
    
    Selection.InsertBreak Type:=wdPageBreak
    �������������� "�����"
    '������������
    'Selection.TypeParagraph
    Selection.TypeText TextStr
    
    Selection.InsertBreak Type:=wdPageBreak
    �������������� "������"
    Selection.TypeText TextStr 'TextStrSeq
    
    Selection.InsertBreak Type:=wdPageBreak
    �������������� "�����������"
    TestStr = Replace(TextStrSeq, " " & dicWords.Item(". ") & " ", Chr(13))
    TestStr = Replace(TestStr, " ", Chr(9))
    Selection.TypeText TestStr
    
    �������������������������������� 2, Sort
        
    Set dicFrec = Nothing
    Set dicWords = Nothing
    Set dicTemp = Nothing
End Sub

Sub ������������������(Selected As Boolean, frecConst As Integer, Registr As Integer, Sort As Boolean)
Dim oWord As Range, vX As Variant, vX2 As Variant, S As String, TextWords As Object
Dim t, TextStr, TextStrSeq, TextFromDict, TestStr As String             ' ���������� ��� ������ ���������� "�����" (oWord.Text)
Dim i As Double, MaxInd As Double
Dim dicFrec As Object, dicWords As Object, dicTemp As Object
Dim WordsCount As Double
Dim WordsCountTest As Double
Set dicFrec = CreateObject("Scripting.Dictionary")
Set dicWords = CreateObject("Scripting.Dictionary")
Set dicTemp = CreateObject("Scripting.Dictionary")
Application.ScreenUpdating = False

TextStr = ""
TextStrSeq = ""
i = 0
WordsCount = 0
If Selected = True Then
    Set TextWords = Selection.Sentences
Else
    Set TextWords = ActiveDocument.Sentences
End If

With dicFrec '���������� ������� � ��������� ����
    .CompareMode = Registr        ' ��������� ���������������� � �������� � �������
    For Each oWord In TextWords '������� ����
        t = oWord.Text
        t = ����������������������(t)
        ' ������ (� ����� ��� �������) ������������ ������� �������
        If Right(t, 1) = ChrW(160) Then t = Mid(t, 1, Len(t) - 1): t = t & " " ': MsgBox t & vbCr & _
                                                                            "Len(t) = " & Len(t)
    ' ���������� � "�����" (oWord.Text) �������, ���� ��� ���, ����� ���� = ����� � �������� � ���
        If Right(t, 1) Like "[!" & Chr(10) & Chr(11) & Chr(13) & Chr(30) & Chr(32) & "]" Then t = t & " "
        
        S = UCase(t)
        
            Select Case AscW(S) ' ��������� ������ 1-� �����
            Case 30, 33 To 90, 132, 133, 145 To 148, 168, 171, 184, 187, 1025, 1040 To 1071 ', _
            8211, 8212 ' �����, ���. ����� ���������� (� �. �. ����: 8212), ���. � ���. ����. �����
            .Item(t) = .Item(t) + 1
            End Select
            
        If AscW(t) = 8211 Then .Item(t) = .Item(t) + 1  ' ������� ������� ������� (N-dash:�)
        If AscW(t) = 8212 Then .Item(t) = .Item(t) + 1  ' ������� ������� ���� (M-dash:�)
    Next
    
    '�������� ���� � ������ ���������
    For Each vX In .Keys
        If CInt(.Item(vX)) < frecConst Then
            .Remove (vX)
        End If
    Next
    
    dicTemp.CompareMode = Registr        ' ��������� ���������������� � �������� � �������
    For Each vX In .Keys '����� ������� ������, ����� � �� �������
        dicTemp(vX) = .Item(vX)
        WordsCountTest = CInt(.Item(vX))
        WordsCount = WordsCount + WordsCountTest '�� ���� ������� ���������� ����
    Next
End With

    i = 0
    dicWords.CompareMode = Registr        ' ��������� ���������������� � �������� � �������
    For Each vX In dicTemp.Keys '������� ������ ������������ �������, ����������, ������� ������� �� �������
        i = i + 1
        TestStr = vX
        TestStr = dicTemp.Item(vX)
        MaxInd = FindMaxInd(dicTemp.Items())
        dicWords.Item(dicTemp.Keys()(MaxInd)) = i '���������� �������� ��������� �� ������ �������. ������ ������� �� 1
        dicTemp.Remove (dicTemp.Keys()(MaxInd))
    Next
    dicTemp.RemoveAll '������� �������, ������ ��� ������ ���-�� �������

    '����� � ����� ���������� ����
    For Each oWord In TextWords '������� ����
        t = oWord.Text
        t = ����������������������(t)
        ' ������ (� ����� ��� �������) ������������ ������� �������
        If Right(t, 1) = ChrW(160) Then t = Mid(t, 1, Len(t) - 1): t = t & " " ': MsgBox t & vbCr & _
                                                                            "Len(t) = " & Len(t)
    ' ���������� � "�����" (oWord.Text) �������, ���� ��� ���, ����� ���� = ����� � �������� � ���
        If Right(t, 1) Like "[!" & Chr(10) & Chr(11) & Chr(13) & Chr(30) & Chr(32) & "]" Then t = t & " "
        
        If CheckBox20.Value = True Then '����������� ������ �������
            TextFromDict = CStr(dicFrec.Item(t) / WordsCount)
        Else
            TextFromDict = CStr(dicWords.Item(t))
        End If
        If TextFromDict = " " Or TextFromDict = "0" Then '���� ����� ������� �� ����, �� ����������� ������ ���� 0
            TextFromDict = ""
        End If
        TextStrSeq = TextStrSeq & TextFromDict & " "
        If CheckBox20.Value = True Then '����������� ������ �������
            If t = ". " Then
                TextStrSeq = TextStrSeq & Chr(13)
            End If
        End If
                
        TextStr = TextStr & t & Chr(9) & TextFromDict & " "
        If CheckBox21.Value = True Then '� ����� ������
            TextStr = TextStr & Chr(13)
        End If
    Next
    TextStr = �������������(TextStr)
    TextStrSeq = �������������(TextStrSeq)
    TextStr = ���������������(TextStr)
    TextStrSeq = ���������������(TextStrSeq)
    
    ����������������������������� '��������������������������������
    �������������� "�������"
    Selection.TypeText "�����" & Chr(9) & "������ �� �������" & Chr(9) & "�������" & Chr(9) & "�����������" & Chr(9) & "�����" & Chr(13)
    For Each vX In dicWords.Keys
        If dicWords.Item(vX) <> "" Then
            Selection.TypeText RTrim(vX) & Chr(9) & dicWords.Item(vX) & Chr(9) & dicFrec.Item(vX) & Chr(9) & CStr(dicFrec.Item(vX) / WordsCount) & Chr(9) & Len(RTrim(vX)) & Chr(13)
        End If
    Next
    
    Selection.InsertBreak Type:=wdPageBreak
    �������������� "�����"
    '������������
    'Selection.TypeParagraph
    Selection.TypeText TextStr
    
    Selection.InsertBreak Type:=wdPageBreak
    �������������� "������"
    Selection.TypeText TextStr 'TextStrSeq
    
    Selection.InsertBreak Type:=wdPageBreak
    �������������� "�����������"
    'TestStr = Replace(TextStrSeq, ". ", Chr(13))
    'TestStr = Replace(TextStrSeq, ". ", "")
    TestStr = Replace(TextStrSeq, " ", Chr(13))
    Selection.TypeText TestStr
    
    �������������������������������� 2, Sort
        
    Set dicFrec = Nothing
    Set dicWords = Nothing
    Set dicTemp = Nothing
End Sub

Public Function ����������������(ByVal TextStr As String, ByVal DelPharagraph)
'Dim AlfU() As String
'Dim AlfL() As String
'Dim EngAlfL() As String
'Dim AlfPrep(0 To 7) As String
    Dim i
    Dim AlfPrep: AlfPrep = Split(". . ? ! ; : ( ) * -", " "): AlfPrep(0) = Chr$(13) ' vbCr - ������ ������
    Dim AlfSpecChars: AlfSpecChars = Split("@ # � $ % ^ & _ + = \ / | ~ ` < > "" ' � [ ] { } �", " ")
    Dim AlfNums: AlfNums = Split("1 2 3 4 5 6 7 8 9 0", " ")
    Dim AlfNumsStr: AlfNumsStr = Split("���� ��� ��� ������ ���� ����� ���� ������ ������ ����", " ")
    
    Dim AlfU: AlfU = Split("� � � � � � � � � � � � � � � � � � � � � � � � � � � � � � � � �", " ")
    Dim AlfL: AlfL = Split("� � � � � � � � � � � � � � � � � � � � � � � � � � � � � � � � �", " ")
    
    Dim EngAlfU: EngAlfU = Split("A B C D E F G H I J K L M N O P Q R S T U V W X Y Z", " ")
    Dim EngAlfL: EngAlfL = Split("A B C D E F G H I J K L M N O P Q R S T U V W X Y Z", " ")
    
    For i = 0 To UBound(AlfU)
        AlfL(i) = LCase(AlfU(i))
    Next i
    For i = 0 To UBound(EngAlfU)
        EngAlfL(i) = LCase(EngAlfU(i))
    Next i

    Dim Str: Str = TextStr
    '������� �������� ������� "., ", �.�. ����� ��� ����� ����� ����������
    For i = 0 To UBound(AlfNums) '�������� ������
        Str = Replace(Str, AlfNums(i), AlfNumsStr(i))
    Next i
    For i = 0 To UBound(AlfPrep) '�������� ����������.
        Str = Replace(Str, AlfPrep(i), Asc(AlfPrep(i)) & ", ")
    Next i
    For i = 0 To UBound(AlfSpecChars) '�������� ���� �������
        Str = Replace(Str, AlfSpecChars(i), Asc(AlfSpecChars(i)) & ", ")
    Next i
    
    
    For i = 0 To UBound(AlfU) '�������� ����� ����� �������
        Str = Replace(Str, AlfL(i), Asc(AlfL(i)) & ", ", , , vbBinaryCompare)
    Next i

    For i = 0 To UBound(AlfU) '�������� ������� ����� �������
        Str = Replace(Str, AlfU(i), Asc(AlfU(i)) & ", ", , , vbBinaryCompare)
    Next i
    
    For i = 0 To UBound(EngAlfU) '�������� ����� ����� ����
        Str = Replace(Str, EngAlfL(i), Asc(EngAlfL(i)) & ", ", , , vbBinaryCompare)
    Next i
    
    For i = 0 To UBound(EngAlfU) '�������� ������� ����� ����
        Str = Replace(Str, EngAlfU(i), Asc(EngAlfU(i)) & ", ", , , vbBinaryCompare)
    Next i
    
    Str = Replace(Str, "  ", " ")
    Str = Replace(Str, ", ,", ", " & Asc(",") & ",")
    Str = Replace(Str, " , ", ", ")
    Str = Replace(Str, " ,", ",")
    If DelPharagraph Then
        Str = Replace(Str, "^p", "")
    End If
    
    ���������������� = Str
    'Asc(s) '����� ��������� �������� ������� "�", ������ 193
    'Chr(Asc(s)) '����������� ��������� �������� ������� � ��� �����
End Function

Private Sub CheckBox7_Click()
    CheckBox1.Value = False
End Sub

Private Sub CommandButton1_Click()
    Dim Selected As Boolean, frecConst As Integer, Registr As Integer, Sort As Boolean
    Selected = CheckBox1.Value
    frecConst = Int(TextBox3.Text)
    If CheckBox3.Value = True Then
        Registr = 1
    Else
        Registr = 0
    End If
    Sort = CheckBox2.Value
    �������_���� Selected, frecConst, Registr, Sort
End Sub

Private Sub CommandButton10_Click()
    Dim XDoc As Object, root As Object, elem As Object
    Set XDoc = CreateObject("MSXML2.DOMDocument")
    Set root = XDoc.createElement("Root")
    XDoc.appendChild root
    'Add child to root
    Set elem = XDoc.createElement("Child")
    root.appendChild elem
    'Add Attribute to the child
    Dim rel As Object
    Set rel = XDoc.createAttribute("Attrib")
    rel.NodeValue = "Attrib value"
    elem.setAttributeNode rel
    'Save the XML file
    XDoc.Save "E:\my_file.xml"
End Sub

Private Sub CommandButton11_Click()
    Dim StrText: StrText = Selection.Text
    If Len(Trim(StrText)) = 0 Or StrText = Chr(13) Then Exit Sub
    StrText = ����������������������(StrText)
    MsgBox ("�������:   " & Chr(9) & ��������������(StrText) & Chr(13) & "���������: " & Chr(9) & ����������������(StrText))
End Sub

Private Sub CommandButton12_Click()
    Dim Selected As Boolean, frecConst As Integer, Registr As Integer, Sort As Boolean
    Selected = CheckBox1.Value
    frecConst = Int(TextBox3.Text)
    If CheckBox3.Value = True Then
        Registr = 1
    Else
        Registr = 0
    End If
    Sort = CheckBox2.Value
    ��������������������� Selected, frecConst, Registr, Sort
End Sub

Private Sub CommandButton13_Click()
    Dim Selected As Boolean, frecConst As Integer, Registr As Integer, Sort As Boolean
    Selected = CheckBox1.Value
    frecConst = Int(TextBox3.Text)
    If CheckBox3.Value = True Then
        Registr = 1
    Else
        Registr = 0
    End If
    Sort = CheckBox2.Value
    ��������������������������� Selected, frecConst, Registr, Sort
End Sub

Private Sub CommandButton14_Click()
    Dim Selected As Boolean, frecConst As Integer, Registr As Integer, Sort As Boolean
    Selected = CheckBox1.Value
    frecConst = Int(TextBox3.Text)
    If CheckBox3.Value = True Then
        Registr = 1
    Else
        Registr = 0
    End If
    Sort = CheckBox2.Value
    ������������������ Selected, frecConst, Registr, Sort
End Sub

Private Sub CommandButton15_Click()
    �������������
    Dim Handles, RemoveParam, RemoveTextParam, i, Replacement
    RemoveTextParam = TextBox5.Text
    Replacement = TextBox6.Text
    
    RemoveParam = Split(RemoveTextParam, " ")
    For i = 0 To UBound(RemoveParam)
        If RemoveParam(i) <> "" Then
            Selection.find.ClearFormatting
            Selection.find.Replacement.ClearFormatting
            With Selection.find
                .Text = RemoveParam(i)
                .Replacement.Text = Replacement
                .Forward = True
                .Wrap = wdFindContinue
                .Format = True
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
        End If
        Selection.find.Execute Replace:=wdReplaceAll
    Next i
End Sub

Public Sub �����������(Glasnie As Boolean)
Dim symbol As String
Dim Array_Str() As String
Dim TextWords As String
Dim i, j
Dim Array_Glasn: Array_Glasn = Split("�, �, �, �, �, �, �, �, �, �", ", ")
Dim Array_Sogl: Array_Sogl = Split("�, �, �, �, �, �, �, �, �, �, �, �, �, �, �, �, �, �, �, �, �", ", ")
    
    If Glasnie = True Then
        Array_Str = Array_Glasn
    Else
        Array_Str = Array_Sogl
    End If

    TextWords = ActiveDocument.Range.Text

'    Selection.WholeStory
'    Selection.MoveUp Unit:=wdLine, Count:=1
'    Selection.MoveRight Unit:=wdCharacter, Count:=1
    For i = Len(TextWords) To 1 Step -1
      symbol = LCase(Mid(TextWords, i, 1))
      For j = LBound(Array_Str) To UBound(Array_Str)
        If symbol = Array_Str(j) Then
          ActiveDocument.Range(i - 1, i).Select
          With Selection.Font.Borders(1)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            If UBound(Array_Str) = UBound(Array_Glasn) Then
                .Color = wdColorRed 'Options.DefaultBorderColor
            Else
                .Color = wdColorBlue 'Options.DefaultBorderColor
            End If
              
          End With
          Exit For
        End If
'        Exit For
'        Exit For
      Next
    Next
End Sub

Private Sub CommandButton16_Click()
    ����������� CheckBox23.Value
End Sub

Private Sub CommandButton2_Click()
    �������_���������_����
End Sub

Private Sub CommandButton3_Click()
    Dim StrText: StrText = Selection.Text
    If MsgBox("�� � ���� ������� (������� ������ ������)?", vbYesNo, "�� � ���� �������?") = vbYes Then
        Dim DelParagraph: DelParagraph = True
    End If
    Dim Str: Str = ����������������(StrText, DelParagraph)
    ������������
    Selection.TypeParagraph
    Selection.Text = Str
End Sub

Function �������_����������(WordFind As String, Selected As Boolean, Registr As Integer)
Dim oWord As Range, dic As Object, vX As Variant, S As String, TextWords As Object
Dim t, WordText As String             ' ���������� ��� ������ ���������� "�����" (oWord.Text)
Set dic = CreateObject("Scripting.Dictionary")
Application.ScreenUpdating = False

With dic
        .CompareMode = Registr        ' ��������� ���������������� � �������� � �������
    If Selected = True Then
        Set TextWords = Selection.Range.Words
    Else
        Set TextWords = ActiveDocument.Range.Words
    End If
    WordFind = Trim(WordFind) & " "
    Registr = 1
    
    For Each oWord In TextWords
        t = oWord.Text
        ' ������ (� ����� ��� �������) ������������ ������� �������
        If Right(t, 1) = ChrW(160) Then t = Mid(t, 1, Len(t) - 1): t = t & " " ': MsgBox t & vbCr & _
                                                                            "Len(t) = " & Len(t)
    ' ���������� � "�����" (oWord.Text) �������, ���� ��� ���, ����� ���� = ����� � �������� � ���
        If Right(t, 1) Like "[!" & Chr(10) & Chr(11) & Chr(13) & Chr(30) & Chr(32) & "]" Then t = t & " "
        S = t
        If Registr = 1 Then
            S = UCase(t)
            WordText = UCase(WordFind)
        End If
        If InStr(Trim(S), Trim(WordText)) > 0 Then 'WordFind - ������� �����. '���� ��� ����������, �� ����� 0
            .Item(WordFind) = .Item(WordFind) + 1
        End If
    Next
    
    �������_���������� = .Item(WordFind)
End With

Set dic = Nothing
End Function

Function �������_�����������(ByVal WordFind As String, ByVal SearchPart As Boolean, ByVal Selected As Boolean, ByVal Registr As Integer)
Dim oWord As Range, dic As Object, vX As Variant, S As String, TextWords As Object
Dim t, WordText As String             ' ���������� ��� ������ ���������� "�����" (oWord.Text)
Set dic = CreateObject("Scripting.Dictionary")
Application.ScreenUpdating = False

With dic
        .CompareMode = Registr        ' ��������� ���������������� � �������� � �������
    If Selected = True Then
        Set TextWords = Selection.Range.Words
    Else
        Set TextWords = ActiveDocument.Range.Words
    End If
    WordFind = Trim(WordFind) & " "
    Registr = 1
    
    For Each oWord In TextWords
        t = oWord.Text
        ' ������ (� ����� ��� �������) ������������ ������� �������
        If Right(t, 1) = ChrW(160) Then t = Mid(t, 1, Len(t) - 1): t = t & " " ': MsgBox t & vbCr & _
                                                                            "Len(t) = " & Len(t)
    ' ���������� � "�����" (oWord.Text) �������, ���� ��� ���, ����� ���� = ����� � �������� � ���
        If Right(t, 1) Like "[!" & Chr(10) & Chr(11) & Chr(13) & Chr(30) & Chr(32) & "]" Then t = t & " "
        S = t
        If Registr = 1 Then
            S = UCase(t)
            WordText = UCase(WordFind)
        End If
        If SearchPart Then '���� ������ ����� �����, �.�. �����, ������� �������� �������� �����
            If InStr(Trim(S), Trim(WordText)) > 0 Then 'WordFind - ������� �����. '���� ��� ����������, �� ����� 0
                .Item(WordFind) = .Item(WordFind) + 1
            End If
        Else
            If S = WordText Then 'WordFind - ������� �����
                .Item(WordFind) = .Item(WordFind) + 1
            End If
        End If
    Next
    
    �������_����������� = .Item(WordFind)
End With

Set dic = Nothing
End Function

Sub ������������������������������(ByVal WordFind As String, ByVal SearchPart As Boolean, ByRef xlSheet As Object, ByVal MatchCase As Boolean)
    Dim FindStyle
    FindStyle = "��������� 3;�_����"
    ������������������������������������ WordFind, SearchPart, FindStyle, xlSheet, MatchCase
End Sub

Sub �������������������������������(ByVal WordFind As String, ByVal SearchPart As Boolean, ByRef xlSheet As Object, ByVal MatchCase As Boolean)
    Dim FindStyle
    FindStyle = "��������� 2;�_���"
    ������������������������������������ WordFind, SearchPart, FindStyle, xlSheet, MatchCase
End Sub

Sub ������������������������������������(ByVal WordFind As String, ByVal SearchPart As Boolean, ByVal FindStyle As String, ByRef xlSheet As Object, ByVal MatchCase As Boolean)
    '���������� ������� ��������� ����� ���������� ������
    Dim Words(), MaxCount, i, n_Words As Integer
    Dim WordsDate(), S, s_Words ', FindStyle As String
    '����������������������
    Dim rgePages As Range
    Dim cursorEnd As Range
    MaxCount = �������������������(FindStyle) '������������ ���������� ����������
    'FindStyle = "��������� 2;�_���"
    �������������
    Selection.find.ClearFormatting
    Selection.find.Style = ActiveDocument.Styles(FindStyle)
    With Selection.find
        .Forward = True
        .MatchCase = MatchCase
        .Execute
    End With
    i = 0
    'ActiveDocument.Range(0, Selection.Start).Paragraphs.Count
    'MsgBox (ActiveDocument.Range(0, Selection.Start).Paragraphs.Count)
    ReDim Preserve Words(0): Words(0) = 0
    ReDim Preserve WordsDate(0): WordsDate(0) = ""
    Do While Selection.find.Found = True And i < MaxCount
        i = i + 1
        S = Selection.Text
        s_Words = Trim(Left(S, Len(S) - 1)) '�������� ������� �������� ����� (����� ������ ��������� �� ������ ������)
        Selection.HomeKey Unit:=wdLine
        'Selection.GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Count:=Ls
        Set rgePages = Selection.Range
        Selection.EndKey
        Selection.find.ClearFormatting
        Selection.find.Style = ActiveDocument.Styles(FindStyle)
        With Selection.find
            .Forward = True
            .MatchCase = MatchCase
            .Execute
        End With
        Selection.HomeKey Unit:=wdLine
        Set cursorEnd = Selection.Range
        rgePages.End = Selection.Range.End
        rgePages.Select
        n_Words = �������_�����������(WordFind, SearchPart, True, 1)
        If i = 1 Then
            Words(0) = n_Words
            WordsDate(0) = s_Words
        Else
            ReDim Preserve Words(UBound(Words) + 1): Words(UBound(Words)) = n_Words
            ReDim Preserve WordsDate(UBound(WordsDate) + 1): WordsDate(UBound(WordsDate)) = s_Words
        End If
        cursorEnd.Select
        Selection.find.ClearFormatting
        Selection.find.Style = ActiveDocument.Styles(FindStyle)
        With Selection.find
            .Forward = True
            .MatchCase = MatchCase
            .Execute
        End With
    Loop
    
    If CheckBox9.Value = True Then
        ������������
        Selection.TypeParagraph
        If SearchPart Then
            Selection.TypeText Text:="������� ���� '*" & Trim(WordFind) & "*' �� ����������: " & "(" + Trim(Str(MaxCount)) + ")"
        Else
            Selection.TypeText Text:="������� ���� '" & Trim(WordFind) & "' �� ����������: " & "(" + Trim(Str(MaxCount)) + ")"
        End If
        Selection.TypeParagraph
        For i = 0 To UBound(Words)
            Selection.TypeText Text:=i + 1
            Selection.TypeText Text:=vbTab
            Selection.TypeText Text:=WordsDate(i)
            Selection.TypeText Text:=vbTab
            Selection.TypeText Text:=Words(i)
            Selection.TypeParagraph
        Next i
    End If
    
    Dim x, y As Integer
    x = 2: y = 1
    If CheckBox8.Value = True Then
        Do While xlSheet.Application.Cells(y, 1).Value <> ""
            y = y + 1
            'Exit Do
        Loop
        xlSheet.Application.Cells(y, 1).Value = WordFind
        For i = 0 To UBound(Words)
            xlSheet.Application.Cells(y, x + i).Value = Words(i)
            xlSheet.Application.Cells(1, 2 + i).Value = WordsDate(i)
        Next i
    End If
End Sub

Private Sub �����������(ByVal WordFind As String, ByRef xlSheet As Object)  '������� ��������� �����
    'WordFind = Trim(TextBox4.Text) '����� ����� ������
    Dim Selected, SearchPart As Boolean, WordFrec As Integer, Registr As Integer
    Dim MatchCase As Boolean
    If WordFind = "" Then
        MsgBox ("�� ������ ������� �����")
        Exit Sub
    End If
    Dim rgePages As Range
    Set rgePages = Selection.Range
    
    Selected = CheckBox1.Value '���� � ���������� �����
    SearchPart = CheckBox5.Value '���� ����� �����
    If Selected And Not CheckBox4.Value And Len(Selection.Text) = 1 Then
        MsgBox ("�������� ����� ��� ������� ����� � '������ ���������� �����'.")
        Exit Sub
    End If
    
    If CheckBox4.Value = True Then '����� �� ���� ��� ����� ������
        ������������������������������ WordFind, SearchPart, xlSheet, MatchCase
        'MsgBox (Trim(WordFind) & ": " & WordFrec)
    ElseIf CheckBox6.Value = True Then '����� �� ����� ��� ����� ������
        ������������������������������� WordFind, SearchPart, xlSheet, MatchCase
        'MsgBox (Trim(WordFind) & ": " & WordFrec)
    Else
        WordFrec = �������_�����������(WordFind, SearchPart, Selected, 1)
        MsgBox (Trim(WordFind) & ": " & WordFrec)
    End If
    rgePages.Select
End Sub

Private Sub CommandButton4_Click() '������� ��������� �����
    Dim Words, wordString
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    xlSheet.Application.Cells(1, 1).Value = "�����\������� �� ����"
    If CheckBox8.Value = True Then
        If CheckBox10.Value = True Then
            xlSheet.Application.Visible = True
        End If
    End If
    
    Dim sngStart: sngStart = Timer              ' ������ �������
    If CheckBox7.Value = True Then '��������� ����
        Words = Split(Selection.Text, vbCr) '�������� ����� �������. ��� ���� vbNewLine vbCrLf
        For Each wordString In Words
            If wordString <> "" Then
                ����������� wordString, xlSheet
            End If
        Next wordString
    Else
        ����������� Trim(TextBox4.Text), xlSheet
    End If
    Dim sngEnd: sngEnd = Timer              ' ����� �������
    Dim sngElapsed: sngElapsed = Format(sngEnd - sngStart, "Fixed") ' ����������.
    Label4.Caption = "������� �����: " & sngElapsed & " ������."
    If CheckBox8.Value = True And xlApp <> Null Then
        If CheckBox10.Value = True Then
            xlSheet.Application.Visible = True
        End If
        xlSheet.Application.Visible = True
    End If
End Sub

Sub ����������������������()
    Selection.find.ClearFormatting
    Selection.find.Replacement.ClearFormatting
    With Selection.find
        .Text = "�"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.find.Execute Replace:=wdReplaceAll

    Selection.find.ClearFormatting
    Selection.find.Replacement.ClearFormatting
    With Selection.find
        .Text = "�"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.find.Execute Replace:=wdReplaceAll

    Selection.find.ClearFormatting
    Selection.find.Replacement.ClearFormatting
    With Selection.find
        .Text = "^t"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.find.Execute Replace:=wdReplaceAll
    
    Selection.find.ClearFormatting
    Selection.find.Replacement.ClearFormatting
    With Selection.find
        .Text = "^="
        .Replacement.Text = "-"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.find.Execute Replace:=wdReplaceAll
End Sub


Private Sub CommandButton5_Click()
    ����������������������
    Dim StrText: StrText = Selection.Text
    Dim StrNumsIO
    Dim Normalize
    If MsgBox("�� � ���� ������� (������� ������ ������)?", vbYesNo, "�� � ���� �������?") = vbYes Then
        Dim DelParagraph: DelParagraph = True
    End If
    Normalize = CheckBox11.Value
    StrNumsIO = ��������������������������(StrText, DelParagraph, Normalize)
    
    If CheckBox12.Value Then
        ������������������������
        Selection.TypeText StrNumsIO
    Else
        ������������
        Selection.TypeParagraph
        Selection.Text = StrNumsIO
    End If
End Sub

Public Function FindMaxInd(Arr) '�� ���� � ����� ������
    On Error Resume Next
    Dim myMax As Double
    Dim i As Double
    
    For i = LBound(Arr, 1) To UBound(Arr, 1)
      If Arr(i) > myMax And Arr(i) <> "" Then
        myMax = Arr(i)
        FindMaxInd = i
      End If
    Next i
End Function

Public Function FindMaxVal(Arr)
    On Error Resume Next
    Dim myMax As Long
    Dim i As Long
    
    For i = LBound(Arr, 1) To UBound(Arr, 1)
      If Arr(i) > myMax And Arr(i) <> "" Then
        myMax = Arr(i)
        FindMaxVal = myMax
      End If
    Next i
End Function

Public Function ��������������������������(ByVal StrText As String, ByVal DelParagraph As Boolean, ByVal Normalize As Boolean)
    Dim i
    Dim Str: Str = StrText
    Str = ����������������(Str, DelParagraph)

    Dim StrNums: StrNums = Split(Str, ", ") '������ ����� � ���� �����
    'Dim StrNumsSize: StrNumsSize =
    Dim StrNumsIO
    Dim MaxValInd
    Dim MaxVal
    
    For i = 0 To UBound(StrNums) - 2
        If Normalize Then
            MaxValInd = FindMaxInd(StrNums)
            MaxVal = FindMaxVal(StrNums)
        Else
            MaxVal = 1
        End If
        StrNumsIO = StrNumsIO + CStr(CDbl(StrNums(i)) / MaxVal) + vbTab + CStr(CDbl(StrNums(i + 1)) / MaxVal) + vbNewLine
    Next i
    �������������������������� = StrNumsIO
End Function

Public Function ��������������(ByVal Str As String)
    'Dim StrNums: StrNums = Str
    Dim StrNumsLines: StrNumsLines = Split(Str, vbCr)
    Dim StrNums ': StrNums = Split(Str, ", ")
    Dim StrNumStr, i, j
    For i = 0 To UBound(StrNumsLines)
        If (StrNumsLines(i) <> "") And (StrNumsLines(i) <> vbNewLine) Then
            StrNums = Split(StrNumsLines(i), ", ")
            For j = 0 To UBound(StrNums)
                StrNums(j) = Trim(StrNums(j)) '����������������������(StrNums(i))
                If (StrNums(j) <> "") Then
                    StrNums(j) = Chr(CLng(StrNums(j)))
                    If (StrNums(j) = vbCr) Then '���� � ������ ������� ������
                        StrNumStr = StrNumStr + "\n"
                    Else
                        StrNumStr = StrNumStr + StrNums(j)
                    End If
                End If
            Next j
            StrNumStr = StrNumStr + vbNewLine
        End If
    Next i
    'StrNumStr = Replace(StrNumStr, vbCr, "\n")
    �������������� = StrNumStr
End Function

Public Function ����������������������(ByVal Str As String)
    Dim StrText: StrText = Str
    If Mid(StrText, Len(StrText), 1) = vbNewLine Then
        StrText = Left(StrText, Len(StrText) - 1) '������� ��������� ������
    End If
    If Mid(StrText, Len(StrText), 1) = vbCr Then
        StrText = Left(StrText, Len(StrText) - 1) '������� ��������� ������
    End If
    If Mid(StrText, Len(StrText), 1) = vbCrLf Then
        StrText = Left(StrText, Len(StrText) - 1) '������� ��������� ������
    End If
    ���������������������� = StrText
End Function

Private Sub CommandButton6_Click()
    Dim StrText: StrText = Selection.Text
    StrText = ����������������������(StrText)
    Dim Str: Str = ��������������(StrText)
    ������������
    Selection.TypeParagraph
    Selection.Text = Str
    'Asc(s) '����� ��������� �������� ������� "�", ������ 193
    'Chr(Asc(s)) '����������� ��������� �������� ������� � ��� �����
End Sub

Public Function ����������������������(Selected As Boolean, frecConst As Integer, Registr As Integer, Sort As Boolean)
Dim freq: freq = frecConst   ' ������� ��������� ������� oWord (����� ��� �����), ������ ������� ��� �����������
Dim oWord As Range, dic As Object, vX As Variant, S As String, TextWords As Object
Dim t, ResultText As String             ' ���������� ��� ������ ���������� "�����" (oWord.Text)
Dim i
Dim StrIO()

    If Selected = True Then
        Set TextWords = Selection.Range.Words
    Else
        Set TextWords = ActiveDocument.Range.Words
    End If
    
    ReDim Preserve StrIO(0): StrIO(0) = ""

    i = 0
    For Each oWord In TextWords
        i = i + 1
        t = oWord.Text
        If t Like "*" Then '"*[�-�A-z]*" @#�$%^&_+=\/|~`<>""'�{}�.?!;:()*-1234567890
            ' ������ (� ����� ��� �������) ������������ ������� �������
            If Right(t, 1) = ChrW(160) Then t = Mid(t, 1, Len(t) - 1): t = t & " " ': MsgBox t & vbCr & _
                                                                                "Len(t) = " & Len(t)
        ' ���������� � "�����" (oWord.Text) �������, ���� ��� ���, ����� ���� = ����� � �������� � ���
            If Right(t, 1) Like "[!" & Chr(10) & Chr(11) & Chr(13) & Chr(30) & Chr(32) & "]" Then t = t & " "
    
            If i = 1 Then
                StrIO(0) = t
            Else
                ReDim Preserve StrIO(UBound(StrIO) + 1): StrIO(UBound(StrIO)) = t
            End If
        End If
    Next
    
    For i = 0 To UBound(StrIO) - 2
        ResultText = ResultText + StrIO(i) + vbTab + StrIO(i + 1) + vbNewLine
    Next i
    
    ���������������������� = ResultText
End Function

Private Sub ������������������������()
    '���������� � �������� ������ ��������� ��� ������ ���� ���������� �������
    Documents.Add
    Options.CheckGrammarAsYouType = False               ' ������ ��������������� �������� �����
    Options.CheckSpellingAsYouType = False              ' ������ ���������������� �������� �����

    With ActiveDocument    '��������� �����; ��������� ����� (���� .LineNumbering.Active = True)
        With .PageSetup
    '    .LineNumbering.Active = True                    ' ��������� ����� ������� (�� �������!)
    '    .LineNumbering.RestartMode = wdRestartContinuous ' �������� ��������� �����
    '    .LineNumbering.DistanceFromText = 4 'pt
        .TopMargin = CentimetersToPoints(1)
        .BottomMargin = CentimetersToPoints(1)
        .LeftMargin = CentimetersToPoints(1.1)
        .RightMargin = CentimetersToPoints(0)
        End With
    .Paragraphs.TabStops.Add CentimetersToPoints(3.5), wdAlignTabRight ' ��� ������ 35 ��
    End With
    
    With Selection.PageSetup
    .TextColumns.SetCount NumColumns:=4                     ' � 4 �������
    .TextColumns.LineBetween = True                         ' ����� ����� �������? ��!
    .TextColumns.Spacing = CentimetersToPoints(1.9)          ' 19 �� ��� �������
    End With
    
    With Selection
        .Font.Size = 10
        .Collapse
        .Delete
    End With
    ActiveDocument.UndoClear ' ������� ������ ������� �� ����� ��������� ���������
End Sub

Private Sub CommandButton7_Click()
    ����������������������
    Dim TextStr
    TextStr = ����������������������(True, 0, True, False)
    If CheckBox12.Value Then
        ������������������������
        Selection.TypeText TextStr
    Else
        ������������
        Selection.TypeParagraph
        Selection.Text = TextStr
    End If
End Sub

Private Sub CommandButton8_Click()
    Dim Selected As Boolean, frecConst As Integer, Registr As Integer, Sort As Boolean
    Selected = CheckBox1.Value
    frecConst = Int(TextBox3.Text)
    If CheckBox3.Value = True Then
        Registr = 1
    Else
        Registr = 0
    End If
    Sort = CheckBox2.Value
    ������������ Selected, frecConst, Registr, Sort
End Sub

Private Sub CommandButton9_Click()
    �������������
    Dim Handles, RemoveParam, RemoveTextParam, i
    Handles = Split("��������� 4;�_������#��������� 3;�_����#��������� 2;�_���#��������� 1;�_���", "#")
    If CheckBox13.Value = True Then '������� ���������
        For i = 0 To UBound(Handles)
            Selection.find.ClearFormatting
            Selection.find.Style = ActiveDocument.Styles(Handles(i))
            Selection.find.Replacement.ClearFormatting
            With Selection.find
                .Text = ""
                .Replacement.Text = ""
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
        Next i
    End If
    RemoveTextParam = ""
    If CheckBox14.Value = True Then RemoveTextParam = RemoveTextParam + "^p^g "
    If CheckBox15.Value = True Then RemoveTextParam = RemoveTextParam + "^t "
    If CheckBox16.Value = True Then RemoveTextParam = RemoveTextParam + "^p "
    If CheckBox22.Value = True Then RemoveTextParam = RemoveTextParam + ". "
    If CheckBox17.Value = True Then RemoveTextParam = RemoveTextParam + ", ? ! ; : ( ) * - ^= "
    If CheckBox18.Value = True Then RemoveTextParam = RemoveTextParam + "@ # � $ % ^ & _ + = \ / | ~ ` < > "" ' � [ ] { } � "
    If CheckBox19.Value = True Then RemoveTextParam = RemoveTextParam + "1 2 3 4 5 6 7 8 9 0 "
    
    RemoveParam = Split(RemoveTextParam, " ")
    For i = 0 To UBound(RemoveParam)
        If RemoveParam(i) <> "" Then
            Selection.find.ClearFormatting
            Selection.find.Replacement.ClearFormatting
            With Selection.find
                .Text = RemoveParam(i)
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Format = True
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
        End If
        Selection.find.Execute Replace:=wdReplaceAll
    Next i
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Frame2_Click()

End Sub

Private Sub Frame3_Click()

End Sub

Private Sub Frame4_Click()

End Sub


Private Sub Frame5_Click()

End Sub

Private Sub TextBox5_Change()

End Sub
