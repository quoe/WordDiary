VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "ׁלאיכû"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5295
   OleObjectBlob   =   "UserForm2.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type POINTAPI
x As Long
y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" _
(lpPoint As POINTAPI) As Long

Dim pos As POINTAPI ' Declare variable

Private Sub Image1_Click()

End Sub

Sub InsertSmile(Num1 As Integer, Num2 As Integer)
Selection.TypeText Text:=" "
Selection.InsertSymbol Font:="Segoe UI Emoji", CharacterNumber:=-Num1, _
        Unicode:=True
    Selection.InsertSymbol Font:="Segoe UI Emoji", CharacterNumber:=-Num2, _
        Unicode:=True
Selection.TypeText Text:=" "
UserForm2.Hide
End Sub

Public Sub Image2_Click()
InsertSmile 10179, 8704
End Sub

Public Sub Image3_Click()
InsertSmile 10179, 8703
End Sub

Public Sub Image4_Click()
InsertSmile 10179, 8702
End Sub

Public Sub Image5_Click()
InsertSmile 10179, 8701
End Sub

Public Sub Image6_Click()
InsertSmile 10179, 8700
End Sub

Public Sub Image7_Click()
InsertSmile 10179, 8699
End Sub

Public Sub Image8_Click()
InsertSmile 10179, 8698
End Sub

Public Sub Image9_Click()
InsertSmile 10179, 8697
End Sub

Public Sub Image10_Click()
InsertSmile 10179, 8696
End Sub

Public Sub Image11_Click()
InsertSmile 10179, 8695
End Sub

Public Sub Image12_Click()
InsertSmile 10179, 8694
End Sub

Public Sub Image13_Click()
InsertSmile 10179, 8693
End Sub

Public Sub Image14_Click()
InsertSmile 10179, 8692
End Sub

Public Sub Image15_Click()
InsertSmile 10179, 8691
End Sub

Public Sub Image16_Click()
InsertSmile 10179, 8690
End Sub

Public Sub Image17_Click()
InsertSmile 10179, 8689
End Sub

Public Sub Image18_Click()
InsertSmile 10179, 8688
End Sub

Public Sub Image19_Click()
InsertSmile 10179, 8687
End Sub

Public Sub Image20_Click()
InsertSmile 10179, 8686
End Sub

Public Sub Image21_Click()
InsertSmile 10179, 8685
End Sub

Public Sub Image22_Click()
InsertSmile 10179, 8684
End Sub

Public Sub Image23_Click()
InsertSmile 10179, 8683
End Sub

Public Sub Image24_Click()
InsertSmile 10179, 8682
End Sub

Public Sub Image25_Click()
InsertSmile 10179, 8681
End Sub

Public Sub Image26_Click()
InsertSmile 10179, 8680
End Sub

Public Sub Image27_Click()
InsertSmile 10179, 8679
End Sub

Public Sub Image28_Click()
InsertSmile 10179, 8678
End Sub

Public Sub Image29_Click()
InsertSmile 10179, 8677
End Sub

Public Sub Image30_Click()
InsertSmile 10179, 8676
End Sub

Public Sub Image31_Click()
InsertSmile 10179, 8675
End Sub

Public Sub Image32_Click()
InsertSmile 10179, 8674
End Sub

Public Sub Image33_Click()
InsertSmile 10179, 8673
End Sub

Public Sub Image34_Click()
InsertSmile 10179, 8672
End Sub

Public Sub Image35_Click()
InsertSmile 10179, 8671
End Sub

Public Sub Image36_Click()
InsertSmile 10179, 8670
End Sub

Public Sub Image37_Click()
InsertSmile 10179, 8669
End Sub

Public Sub Image38_Click()
InsertSmile 10179, 8668
End Sub

Public Sub Image39_Click()
InsertSmile 10179, 8667
End Sub

Public Sub Image40_Click()
InsertSmile 10179, 8666
End Sub

Public Sub Image41_Click()
InsertSmile 10179, 8665
End Sub

Public Sub Image42_Click()
InsertSmile 10179, 8664
End Sub

Public Sub Image43_Click()
InsertSmile 10179, 8663
End Sub

Public Sub Image44_Click()
InsertSmile 10179, 8662
End Sub

Public Sub Image45_Click()
InsertSmile 10179, 8661
End Sub

Public Sub Image46_Click()
InsertSmile 10179, 8660
End Sub

Public Sub Image47_Click()
InsertSmile 10179, 8659
End Sub

Public Sub Image48_Click()
InsertSmile 10179, 8658
End Sub

Public Sub Image49_Click()
InsertSmile 10179, 8657
End Sub

Public Sub Image50_Click()
InsertSmile 10179, 8656
End Sub

Public Sub Image51_Click()
InsertSmile 10179, 8655
End Sub

Public Sub Image52_Click()
InsertSmile 10179, 8654
End Sub

Public Sub Image53_Click()
InsertSmile 10179, 8653
End Sub

Public Sub Image54_Click()
InsertSmile 10179, 8652
End Sub

Public Sub Image55_Click()
InsertSmile 10179, 8651
End Sub

Public Sub Image56_Click()
InsertSmile 10179, 8650
End Sub

Public Sub Image57_Click()
InsertSmile 10179, 8649
End Sub

Private Sub UserForm_Click()
'GetCursorPos pos
'UserForm2.Caption = "x:=" & pos.x & vbNewLine _
'& "y:=" & pos.y
End Sub

Private Sub UserForm_Initialize()

End Sub
