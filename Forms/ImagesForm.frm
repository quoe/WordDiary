VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImagesForm 
   Caption         =   "������ � �������������"
   ClientHeight    =   3990
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5895
   OleObjectBlob   =   "ImagesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ImagesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub ���������������()
    Debug.Print "Shapes: " & ActiveDocument.Shapes.Count '������������ ���������
    Debug.Print "InlineShapes: " & ActiveDocument.InlineShapes.Count '��������
End Sub

Sub ShapeBrowser()
    For Each objPaint In ActiveDocument.InlineShapes
      objPaint.Select      '�������� ����
    Next objPaint
    For i = 1 To ActiveDocument.InlineShapes.Count
        ActiveDocument.InlineShapes(i).Select '������� i-�� ��������
    Next i
End Sub

Sub AllPictSize()
       Dim PercentSize As Integer
       Dim oIshp As InlineShape
       Dim oshp As Shape
    
       PercentSize = InputBox("Enter percent of full size", "Resize Picture", 100)
    
       For Each oIshp In ActiveDocument.InlineShapes
           With oIshp
               .ScaleHeight = PercentSize
               .ScaleWidth = PercentSize
           End With
       Next oIshp
    
       For Each oshp In ActiveDocument.Shapes
           With oshp
               .ScaleHeight Factor:=Round(PercentSize / 100), _
                 RelativeToOriginalSize:=msoCTrue
               .ScaleWidth Factor:=Round(PercentSize / 100), _
                 RelativeToOriginalSize:=msoCTrue
           End With
       Next oshp
   End Sub

Sub changeImagesWidth()
    Dim iShape As InlineShape
    newW = InputBox("������� ������", "��������� ���� ��������", "100")
    For Each iShape In ActiveDocument.InlineShapes
        'newW = 100 - ��� 100 �� (10 ��)
        WH = iShape.Width / iShape.Height
        iShape.Width = MillimetersToPoints(newW)
        iShape.Height = MillimetersToPoints(newW / WH)
    Next iShape
End Sub

Sub changeImagesScaleLockAspectRatio()
    Dim iShape As InlineShape
    newW = InputBox("������� ������", "��������� ���� ��������", "100")
    For Each iShape In ActiveDocument.InlineShapes
        iShape.LockAspectRatio = msoFalse
        iShape.ScaleWidth 1.4, msoTrue
        iShape.ScaleHeight 0.5, msoFalse
    Next iShape
End Sub

Sub ��������������������()
    Dim iShape As InlineShape
    For Each iShape In ActiveDocument.InlineShapes
        iShape.Select
        With CaptionLabels("�������")
            .NumberStyle = wdCaptionNumberStyleArabic
            .IncludeChapterNumber = False
        End With
        Selection.InsertCaption Label:="�������", TitleAutoText:="", Title:=" � ", _
            Position:=wdCaptionPositionBelow, ExcludeLabel:=0
    Next iShape
End Sub

Private Sub CommandButton1_Click()
    If MsgBox("������� ��� �����������?", vbYesNo, "������������� ��������") = vbYes Then
        For Each iShape In ActiveDocument.InlineShapes
            iShape.Delete
        Next iShape
    End If
End Sub

Private Sub CommandButton2_Click()
    For Each iShape In ActiveDocument.InlineShapes
        iShape.Select
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Next iShape
End Sub

Private Sub CommandButton3_Click()
    ��������������������
End Sub

Private Sub Label3_Click()

End Sub

Private Sub UserForm_Activate()
    Label2.caption = ActiveDocument.Shapes.Count '������������ ���������
    Label3.caption = ActiveDocument.InlineShapes.Count '��������
    i = 0
    For Each iShape In ActiveDocument.InlineShapes
        If iShape.Height > 1.5 Then
            i = i + 1
        End If
    Next iShape
    Label4.caption = i '����������� � ������� ����� 1.5, �.� �� �������������� �����
End Sub

Private Sub UserForm_Click()

End Sub
