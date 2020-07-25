VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HeaderFooterForm 
   Caption         =   "Работа с колонтитулами"
   ClientHeight    =   4140
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5925
   OleObjectBlob   =   "HeaderFooterForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HeaderFooterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ИзменитьРазмерВерхнегоКолонтитула(NewSize)
    Dim i As Long
    For i = 1 To ActiveDocument.Sections.Count
        With ActiveDocument.Sections(i)
            .Headers(wdHeaderFooterPrimary).Range.Font.Size = NewSize
            '.Headers(wdHeaderFooterPrimary).Range.Text = "Новый текст"
        End With
    Next
End Sub

Sub ИзменитьРазмерНижнегоКолонтитула(NewSize)
    Dim i As Long
    For i = 1 To ActiveDocument.Sections.Count
        With ActiveDocument.Sections(i)
            .Footers(wdHeaderFooterPrimary).Range.Font.Size = NewSize
        End With
    Next
End Sub

Sub ИзменитьШрифтВерхнегоКолонтитула(NewFont)
    Dim i As Long
    For i = 1 To ActiveDocument.Sections.Count
        With ActiveDocument.Sections(i)
            .Headers(wdHeaderFooterPrimary).Range.Font.Name = NewFont
            '.Headers(wdHeaderFooterPrimary).Range.Text = "Новый текст"
        End With
    Next
End Sub

Sub ИзменитьШрифтНижнегоКолонтитула(NewFont)
    Dim i As Long
    For i = 1 To ActiveDocument.Sections.Count
        With ActiveDocument.Sections(i)
            .Footers(wdHeaderFooterPrimary).Range.Font.Name = NewFont
        End With
    Next
End Sub

Sub ИзменитьТекстВерхнегоКолонтитула(NewText)
    Dim i As Long
    For i = 1 To ActiveDocument.Sections.Count
        With ActiveDocument.Sections(i)
            .Headers(wdHeaderFooterPrimary).Range.Text = NewText
        End With
    Next
End Sub

Sub ИзменитьТекстНижнегоКолонтитула(NewText)
    Dim i As Long
    For i = 1 To ActiveDocument.Sections.Count
        With ActiveDocument.Sections(i)
            .Footers(wdHeaderFooterPrimary).Range.Text = NewText
        End With
    Next
End Sub

Sub УдалитьВсеВерхниеКолонтитулы()
    Dim sec As Section
    Dim hf As HeaderFooter
    Dim rng As Range
    For Each sec In ActiveDocument.Sections
        For Each hf In sec.Headers
            hf.Range.Delete
        Next hf
    Next sec
End Sub

Sub УдалитьВсеНижниеКолонтитулы()
    Dim sec As Section
    Dim hf As HeaderFooter
    Dim rng As Range
    For Each sec In ActiveDocument.Sections
        For Each hf In sec.Footers
            hf.Range.Delete
        Next hf
    Next sec
End Sub


Private Sub CommandButton1_Click()
    ИзменитьРазмерВерхнегоКолонтитула (TextBox1.Value)
End Sub

Private Sub CommandButton2_Click()
    If OptionButton1.Value Then
        УдалитьВсеВерхниеКолонтитулы
    End If
    If OptionButton2.Value Then
        УдалитьВсеНижниеКолонтитулы
    End If
    If OptionButton3.Value Then
        УдалитьВсеВерхниеКолонтитулы
        УдалитьВсеНижниеКолонтитулы
    End If
End Sub

Private Sub CommandButton3_Click()
    ИзменитьРазмерНижнегоКолонтитула (TextBox2.Value)
End Sub


Private Sub CommandButton4_Click()
    ИзменитьТекстВерхнегоКолонтитула (TextBox3.Value)
End Sub

Private Sub CommandButton5_Click()
    ИзменитьТекстНижнегоКолонтитула (TextBox4.Value)
End Sub

Private Sub CommandButton6_Click()
    ИзменитьШрифтВерхнегоКолонтитула (TextBox5.Value)
End Sub

Private Sub CommandButton7_Click()
    ИзменитьШрифтНижнегоКолонтитула (TextBox6.Value)
End Sub

Private Sub SpinButton1_Change()
    TextBox1.Value = SpinButton1.Value
End Sub

Private Sub SpinButton2_Change()
    TextBox2.Value = SpinButton2.Value
End Sub

Private Sub TextBox1_Change()
    SpinButton1.Value = TextBox1.Value
End Sub

Private Sub TextBox2_Change()
    SpinButton2.Value = TextBox2.Value
End Sub

Private Sub UserForm_Activate()
    SpinButton1.Value = DefaultFontSize
    SpinButton2.Value = DefaultFontSize
    TextBox1.Value = SpinButton1.Value
    TextBox2.Value = SpinButton2.Value
    TextBox5.Value = DefaultFontName
    TextBox6.Value = DefaultFontName
End Sub

Private Sub UserForm_Click()
    
End Sub
