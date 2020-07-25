VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GoldenRatioForm 
   Caption         =   "Золотое сечение"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6960
   OleObjectBlob   =   "GoldenRatioForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GoldenRatioForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Dim G, l, S, sum As Double 'Give, long, short
    If TextBox1.Value = "" Then Exit Sub
    TextBox1.Value = Replace(TextBox1.Value, ".", ",")
    G = TextBox1.Value ' Дан отирезок
    F = (1 + (5 ^ (1 / 2))) / 2
    Select Case ComboBox1.ListIndex
        Case 0 ' Дан Суммарный отрезок
            sum = G
            l = sum / F ' длинный
            S = sum - l ' короткий
        Case 1 ' Дан Длинный отрезок
            l = G
            sum = l * F
            S = sum - l
        Case 2 ' Дан Короткий отрезок
            S = G
            l = S * F
            sum = l + S
        Case Else
            ComboBox1.ListIndex = 0
            CommandButton1_Click
    End Select
    TextBox2.Value = l
    TextBox3.Value = S
    TextBox4.Value = sum
    TextBox2.Value = Replace(TextBox2.Value, ".", ",")
    TextBox3.Value = Replace(TextBox3.Value, ".", ",")
    TextBox4.Value = Replace(TextBox4.Value, ".", ",")
    'MsgBox F
End Sub

Private Sub CommandButton2_Click()
' Копировать в буфер
   With TextBox2
      .SelStart = 0
      .SelLength = Len(.Text)
      .Copy
   End With
End Sub

Private Sub CommandButton3_Click()
' Копировать в буфер
   With TextBox3
      .SelStart = 0
      .SelLength = Len(.Text)
      .Copy
   End With
End Sub

Private Sub CommandButton4_Click()
' Копировать в буфер
   With TextBox4
      .SelStart = 0
      .SelLength = Len(.Text)
      .Copy
   End With
End Sub

Private Sub UserForm_Activate()
    ComboBox1.AddItem "Суммарный отрезок"
    ComboBox1.AddItem "Длинный отрезок"
    ComboBox1.AddItem "Короткий отрезок"
    ComboBox1.ListIndex = 0
End Sub

Private Sub UserForm_Click()

End Sub
