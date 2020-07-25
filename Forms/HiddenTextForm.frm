VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HiddenTextForm 
   Caption         =   "������� �����"
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4305
   OleObjectBlob   =   "HiddenTextForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HiddenTextForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Selection.Style = ActiveDocument.Styles("������� ����")
End Sub

Private Sub CommandButton2_Click()
    Selection.Style = ActiveDocument.Styles("�������") '����� �������
    Selection.ClearFormatting
End Sub

Private Sub CommandButton3_Click()
    With ActiveDocument.Styles("�������").Font
        .Hidden = False
    End With
    UserForm_Activate
End Sub

Private Sub CommandButton4_Click()
    With ActiveDocument.Styles("�������").Font
        .Hidden = True
    End With
    UserForm_Activate
End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Activate()
    If ActiveDocument.Styles("�������").Font.Hidden Then
        Label1.caption = "������ �����"
    Else
        Label1.caption = "������ �� �����"
    End If
End Sub

Private Sub UserForm_Click()

End Sub
