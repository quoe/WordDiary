VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DiaryMainForm 
   Caption         =   "Дневник"
   ClientHeight    =   8445.001
   ClientLeft      =   150
   ClientTop       =   5370
   ClientWidth     =   8295.001
   OleObjectBlob   =   "DiaryMainForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DiaryMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()  'День
'"yyyy" Год
'"q" Квартал
'"m" Месяц
'"y" День года
'"d" День
'"w" День недели
'"ww" Неделя
'"h" Час
'"n" Минута
'"s" Секунда
    If OptionButton5.Value = True Then 'В текст
        If CheckBox3.Value = True Then 'Текущая дата
            If CheckBox1.Value = True Then 'Текущее время
                НовыйДень
                Exit Sub
            Else
                d = Format(Date, "d MMMM, dddd")
                If CheckBox2.Value = True Then
                    t = TextBox2.Value + ":" + TextBox3.Value + ":" + Format(РандомЧислоДоЗаданногоВернуть(59), "00")
                Else
                    t = TextBox2.Value + ":" + TextBox3.Value + ":" + TextBox4.Value
                End If
            End If
        Else
            d = Format(TextBox1.Value, "d MMMM, dddd")
            If CheckBox1.Value = True Then 'Текущее время
                t = Time
            Else
                If CheckBox2.Value = True Then
                    t = TextBox2.Value + ":" + TextBox3.Value + ":" + Format(РандомЧислоДоЗаданногоВернуть(59), "00")
                Else
                    t = TextBox2.Value + ":" + TextBox3.Value + ":" + TextBox4.Value
                End If
            End If
        НовыйДеньДатаВремя d, t
        End If
    End If
    
    If OptionButton6.Value = True Then
        КурсорВКонец
        If CheckBox3.Value = True Then 'Текущая дата
            If CheckBox1.Value = True Then 'Текущее время
                НовыйДень
                Exit Sub
            Else
                d = Format(Date, "d MMMM, dddd")
                If CheckBox2.Value = True Then
                    t = TextBox2.Value + ":" + TextBox3.Value + ":" + Format(РандомЧислоДоЗаданногоВернуть(59), "00")
                Else
                    t = TextBox2.Value + ":" + TextBox3.Value + ":" + TextBox4.Value
                End If
            End If
        Else
            d = Format(TextBox1.Value, "d MMMM, dddd")
            If CheckBox1.Value = True Then 'Текущее время
                t = Time
            Else
                If CheckBox2.Value = True Then
                    t = TextBox2.Value + ":" + TextBox3.Value + ":" + Format(РандомЧислоДоЗаданногоВернуть(59), "00")
                Else
                    t = TextBox2.Value + ":" + TextBox3.Value + ":" + TextBox4.Value
                End If
            End If
        End If
        НовыйДеньДатаВремя d, t
    End If
End Sub

Private Sub CommandButton10_Click()
    TextBox1.Value = Format(DateAdd("d", -1, TextBox1.Value), "dd.mm.yyyy")
End Sub

Private Sub CommandButton11_Click()
    TextBox5.Value = TextBox5.Value + 1
End Sub

Private Sub CommandButton12_Click()
    TextBox5.Value = TextBox5.Value - 1
End Sub

Private Sub CommandButton13_Click()
    If TextBox2.Value < 23 Then
        TextBox2.Value = Format(TextBox2.Value + 1, "00")
    Else
        TextBox2.Value = Format(0, "00")
    End If
End Sub

Private Sub CommandButton14_Click()
    If TextBox2.Value > 0 Then
        TextBox2.Value = Format(TextBox2.Value - 1, "00")
    Else
        TextBox2.Value = Format(23, "00")
    End If
End Sub

Private Sub CommandButton15_Click()
    If TextBox3.Value < 59 Then
        TextBox3.Value = Format(TextBox3.Value + 1, "00")
    Else
        TextBox3.Value = Format(0, "00")
    End If
End Sub

Private Sub CommandButton16_Click()
    If TextBox3.Value > 0 Then
        TextBox3.Value = Format(TextBox3.Value - 1, "00")
    Else
        TextBox3.Value = Format(59, "00")
    End If
End Sub

Private Sub CommandButton17_Click()
    If TextBox4.Value < 59 Then
        TextBox4.Value = Format(TextBox4.Value + 1, "00")
    Else
        TextBox4.Value = Format(0, "00")
    End If
End Sub

Private Sub CommandButton18_Click()
    If TextBox4.Value > 0 Then
        TextBox4.Value = Format(TextBox4.Value - 1, "00")
    Else
        TextBox4.Value = Format(59, "00")
    End If
End Sub

Private Sub CommandButton19_Click()
    ActiveDocument.Save
End Sub

Private Sub CommandButton2_Click()  'Год
    If OptionButton7.Value = True Then
        If CheckBox4.Value = True Then
            НовыйГод
        Else
            НовыйГодГод (TextBox5.Value)
        End If
    End If
    
    If OptionButton8.Value = True Then
        КурсорВКонец
        If CheckBox4.Value = True Then
            НовыйГод
        Else
            НовыйГодГод (TextBox5.Value)
        End If
    End If
End Sub

Private Sub CommandButton20_Click()
    Set Paragraph = Selection.Range 'Запомнить текущие положение курсора
    Selection.GoTo What:=wdGoToHeading, which:=wdGoToPrevious
    Selection.MoveRight Unit:=wdCharacter, Count:=8, Extend:=wdExtend
    'Selection.EndKey Unit:=wdLine, Extend:=wdExtend 'Курсор в конец текста с выделением
    sngEnd = Selection ' Конец отсчёта
    Selection.GoTo What:=wdGoToHeading, which:=wdGoToPrevious
    Selection.MoveRight Unit:=wdCharacter, Count:=8, Extend:=wdExtend
    'Selection.EndKey Unit:=wdLine, Extend:=wdExtend 'Курсор в конец текста с выделением
    sngStart = Selection ' Начало отсчёта
    sngElapsed_m = DateDiff("n", CDate(sngStart), CDate(sngEnd)) 'разница в минутах
    If OptionButton10 = True Then 'ч:м
        sngElapsed_h = sngElapsed_m / 60
        sngElapsed_m = sngElapsed_m - Fix(sngElapsed_h) * 60 ' Fix() отбрасывает дробную часть
        'sngElapsed = Format(CDate(sngEnd) - CDate(sngStart)) ' Приращение.
        Paragraph.Select    'Возвращение к начальному положению курсора
        Selection.TypeText Text:=Format(sngElapsed_h, "00") & ":" & Format(sngElapsed_m, "00")
    End If
    If OptionButton9 = True Then 'минут
        Paragraph.Select    'Возвращение к начальному положению курсора
        Selection.TypeText Text:=sngElapsed_m & " минут"
    End If
    If CheckBox5.Value = True Then DiaryMainForm.Hide
End Sub

Private Sub CommandButton21_Click()
    КурсорНаСобытиеНазад
End Sub

Private Sub CommandButton22_Click()
    КурсорНаСобытиеВперед
End Sub

Private Sub CommandButton23_Click()
    КурсорВНачало
End Sub

Private Sub CommandButton24_Click()
    КурсорВКонец
End Sub

Private Sub CommandButton25_Click()
    ПоказатьОшибки
End Sub

Private Sub CommandButton4_Click()  'Момент
    If OptionButton3.Value = True Then
        If CheckBox1.Value = True Then
            НовыйМомент
        Else
            If CheckBox2.Value = True Then
                S = TextBox2.Value + ":" + TextBox3.Value + ":" + Format(РандомЧислоДоЗаданногоВернуть(59), "00")
            Else
                S = TextBox2.Value + ":" + TextBox3.Value + ":" + TextBox4.Value
            End If
            НовыйМоментВремя (S)
        End If
    End If
    
    If OptionButton4.Value = True Then
        КурсорВКонец
        If CheckBox1.Value = True Then
            НовыйМомент
        Else
            If CheckBox2.Value = True Then
                S = TextBox2.Value + ":" + TextBox3.Value + ":" + Format(РандомЧислоДоЗаданногоВернуть(59), "00")
            Else
                S = TextBox2.Value + ":" + TextBox3.Value + ":" + TextBox4.Value
            End If
            НовыйМоментВремя (S)
        End If
    End If
End Sub

Private Sub CommandButton5_Click()  'Фильм
    CommandButton8_Click
End Sub

Private Sub CommandButton6_Click()  'Сериал
    CommandButton8_Click
End Sub

Private Sub CommandButton7_Click()  'Аниме
    CommandButton8_Click
End Sub

Private Sub CommandButton8_Click() 'Разделит. линия
    If OptionButton1.Value = True Then
        Линия
    End If
    
    If OptionButton2.Value = True Then
        КурсорВКонец
        Линия
    End If
End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub CommandButton9_Click()
'"yyyy" Год
'"q" Квартал
'"m" Месяц
'"y" День года
'"d" День
'"w" День недели
'"ww" Неделя
'"h" Час
'"n" Минута
'"s" Секунда
    TextBox1.Value = Format(DateAdd("d", 1, TextBox1.Value), "dd.mm.yyyy")
End Sub

Private Sub Label4_Click() 'Загрузка системы: ЦП и память в %
    StrCPU = Format(Win32_Processor_LoadPercentage)
    StrMem = Int(Win32_PhysicalMemory_LoadPercentage) 'Format(Win32_PhysicalMemory_LoadPercentage, "Standard")
    Label4.caption = "Загрузка ЦП: " & StrCPU & "%" & vbNewLine & _
                     "Загрузка памяти: " & StrMem & "%"
End Sub



Private Sub OptionButton2_Click()

End Sub

Private Sub TextBox1_Change()
    Weekday_num = Weekday(TextBox1.Value, vbMonday)
    Select Case Weekday_num
        Case 1: Weekday_str = "понедельник"
        Case 2: Weekday_str = "вторник"
        Case 3: Weekday_str = "среда"
        Case 4: Weekday_str = "четверг"
        Case 5: Weekday_str = "пятница"
        Case 6: Weekday_str = "суббота"
        Case 7: Weekday_str = "воскресенье"
        Case Else: Weekday_str = "не определено"
    End Select
    Label2.caption = Weekday_str
End Sub

Private Sub TextBox2_Change()
    'MyStr = Format(334.9, "###0.00")    ' Returns "334.90".
    'TextBox2.Value = Format(TextBox2.Value, "00")
End Sub

Private Sub TextBox5_Change()
    check_year = TextBox5.Value
    If IsDate("29.02." & check_year) = True Then
        Label3.caption = "Високосный" 'проверка высокосен ли год, высокосным считаетсся тот год у которого имеется 29 февраля
    Else
        Label3.caption = "Не високосный"
    End If
End Sub

Private Sub UserForm_Activate()
    curr_time = Time
    curr_date = Date
    TextBox1.Value = Format(curr_date, "dd.mm.yyyy")
    TextBox2.Value = Format(curr_time, "hh") ' часы
    TextBox3.Value = Format(curr_time, "nn") ' минуты
    TextBox4.Value = Format(curr_time, "ss") ' секунды
    TextBox5.Value = Format(curr_date, "yyyy")
    TextBox1_Change
    TextBox5_Change
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    Label4.caption = "Загрузка ЦП: " & vbNewLine & "(нажмите на этот текст)"
End Sub

