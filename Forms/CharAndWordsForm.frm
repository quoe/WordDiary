VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CharAndWordsForm 
   Caption         =   "Работа с символами и словами"
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


Option Compare Text ' текстовое сравнение (например, в Like)
Option Explicit

Sub Частота_сочетаний_слов()
    'Минимальная длина несокращаемой основы слова
    Const OsnovaLen% = 2    ' =2 для подсчета сочетаний слов с разными окончаниями, =100 для подсчета сочетаний конкретных слов с данным окончанием
    Const LyuboyPoryadok As Boolean = True ' = True - для произвольного порядка слов в сочетании, = False - только для прямого порядка слов в сочетании
    Const IgnorirovatZnakiVnutriPredlozheniya As Boolean = True '  =True  - не учитываются знаки препинания внутри предложения, разделяющие слова в сочетании
    Const IgnorirovatPredlogi As Boolean = True '  =True  - предлоги, перечисленные в функции EtoPredlog, не учавствуют в отборе, и игнорируются между словами.
    Const IgnorirovatSoyuzy As Boolean = True '  =True  - союзы, перечисленные в функции EtoSoyuz, не учавствуют в отборе, и игнорируются между словами.
    Const PechatSPonizheniemChastoty As Boolean = True ' Вывод результатов в порядке понижения частоты пар слов. False - в алфавитном порядке.
    Const freq% = 2         ' минимальная частота появления сочетания слов, с которой оно учитывается
    Dim Slova               ' переменная для массива слов
    Dim PredPrediduschee$, SlovoPrediduschee$, SlovoPrediduschee2$, Slovo$ ' переменные для текста предпредыдущего, предыдущего и очередного "слова"
    Dim i&, j&, LenS%, UbOk%, iCr%, dic As Object, Key, S$, Z  ' Вспомогательные переменные
    ' Массив удаляемых окончаний, массив упорядочен от 3 буквенных до 1 буквенных :
    Dim ok: ok = Split("ами еми емя ёте ете ёшь ешь ими ите ишь ому ого умя ыми ыте ышь ам ас ат ax ая ее её ей ем ex ею ёт ет ёх ех ии ие ий им ит их ию ыи ые ый ым ыт ых ыю ми мя ов ое оё ой ом ою ум ут ух ую шь ье ьё ью ья а е ё и о у ы ь ю я", " ")
    UbOk = UBound(ok)
    ' Массив длин окончаний
    Dim k%(): ReDim k(UbOk): For i = 0 To UbOk: k(i) = Len(ok(i)): Next
    ' Массив минимальных длин слов, у которых будем отрезать окончание
    Dim minLen%(): ReDim minLen(UbOk):  For i = 0 To UbOk: minLen(i) = k(i) + OsnovaLen: Next
    ' Массив знаков пунктуации. В отличии от макроса Частота_основ для анализа примыкающих справа к словам знаков пунктуации мы не будем отрезать у каждого слова справа 1 символ и проверять его. Вместо этого вокруг знаков пунктуации мы заранее добавим пробелы.
    Dim Punkt: Punkt = Split(". . ? ! , ; : ( )", " "): Punkt(0) = Chr$(13) ' vbCr - символ абзаца
    S = ActiveDocument.Range.Text
    ' Очистка текста от игнорируемых символов и стандартизация символов
    S = Replace(S, "»", ""): S = Replace(S, "«", ""): S = Replace(S, """", ""): S = Replace(S, "'", "")
    S = Replace$(S, Chr$(10), " "): S = Replace$(S, Chr$(9), " "): S = Replace(S, Chr$(160), " ") ' Неразрывный пробел
    S = Replace(S, Chr$(150), "-"): S = Replace(S, Chr$(151), "-"): S = Replace(S, Chr$(30), "-")
    S = Replace(S, "…", "."): S = Replace(S, "...", ".")
    ' Добавление пробелов вокруг знаков пунктуации
    For i = 0 To UBound(Punkt): S = Replace$(S, Punkt(i), " " & Punkt(i) & " "): Next
    For i = 1 To 5: S = Replace$(S, "  ", " "): Next 'Многократная замена 2-ных пробелов на 1
    S = LCase$(S) ' Переводим весь текст в маленькие буквы
    Slova = Split(S, " ") ' Перенос символьной переменной с текстом в массив, разделитель - пробел.
    S = ""
    Set dic = CreateObject("Scripting.Dictionary") ' Объект словарь для занесения и подсчета сочетаний слов
    With dic
        .CompareMode = 1   ' Отключение чувствительности к регистру в словаре.
        For i = 0 To UBound(Slova)
            Slovo = Slova(i) ' Берем очередное слово или группу символов из массива
            If EtoSoyuz(PredPrediduschee, SlovoPrediduschee, SlovoPrediduschee2, Slovo, IgnorirovatSoyuzy) And IgnorirovatSoyuzy Then
                iCr = 0
            ElseIf EtoPredlog(PredPrediduschee, SlovoPrediduschee, SlovoPrediduschee2, Slovo, IgnorirovatPredlogi) And IgnorirovatPredlogi Then
                iCr = 0
            Else
                S = Left$(Slovo, 1) ' Отрезаем 1 левый символ, если это буква, то превращаем её в заглавную.
                Select Case S
                    Case "" 'Ничего не делаем, 2 пробела идут подряд
                    Case ".", "?", "!"   ' Конец предложения, слова разделенные этими знаками не считаем сочетанием.
                        SlovoPrediduschee = "" ' Предыдущее слово исключаем из учета в сочетании со следующим словом.
                        PredPrediduschee = ""
                        iCr = 0
                    Case ",", ";", ":", "(", ")", "-"
                        If Not IgnorirovatZnakiVnutriPredlozheniya Then
                            SlovoPrediduschee = ""  ' Предыдущее слово исключаем из учета в сочетании со следующим словом.
                            PredPrediduschee = ""
                        End If
                        iCr = 0
                    Case Chr$(13)  ' vbCr - символ абзаца
                        iCr = iCr + 1 ' Счетчик подряд идущих символов абзаца
                        If iCr > 1 Then
                            SlovoPrediduschee = "" ' Слова, разделенные пустой строкой, исключаем из учета сочетаний.
                            PredPrediduschee = ""
                        End If
                    Case "a" To "z", "а" To "я", "ё", "0" To "9", "§", "№", "&", "$", "€"
                        If OsnovaLen < 100 Then
                            If EtoSoyuz("", "", SlovoPrediduschee2, Slovo, False) Then
                            ElseIf EtoPredlog("", "", SlovoPrediduschee2, Slovo, False) Then
                            ElseIf Not Isklyucheniye(Slovo) Then
                                ' Перебираем и удаляем окончания
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
            ' Подсчитываем пары слов в прямом и обратном порядке
            For Each Key In .Keys
                If .Item(Key) > 0 Then
                    Z = Split(Key, " ")
                    S = Z(1) & " " & Z(0)
                    .Item(Key) = .Item(Key) + .Item(S)
                    .Item(S) = 0
                End If
            Next
        End If
        DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents ' Борьба с вылетами из Wоrd
        Application.ScreenUpdating = False
        Documents.Add: ActiveWindow.ActivePane.View.Type = wdPrintView  ' вид "Разметка страницы"
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
            ' Сортировка по частоте с понижением.
            .Sort FieldNumber:="полям 2", SortFieldType:=wdSortFieldNumeric, SortOrder:=wdSortOrderDescending, Separator:=wdSortSeparateByTabs
            .Collapse Direction:=wdCollapseEnd: .Delete
        Else
            ' Сортировка по алфавиту.
            .Sort FieldNumber:="абзацам", SortFieldType:=wdSortFieldAlphanumeric, SortOrder:=wdSortOrderAscending
            .Collapse: .Delete
        End If
    End With
    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents
    With ActiveDocument.PageSetup.TextColumns
        .SetCount NumColumns:=3 ' В 3 колонки.
        .LineBetween = True     ' Линии между колонок.
    End With
    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents
    ActiveDocument.Paragraphs.TabStops.Add CentimetersToPoints(4), wdAlignTabRight
End Sub
Function EtoSoyuz(PredPrediduschee, SlovoPrediduschee, SlovoPrediduschee2, Slovo, Ignorirovat As Boolean) As Boolean  ' Перечень известных союзов
    ' Союзы из 3-х и более слов здесь упрощены.
    Select Case SlovoPrediduschee2 & " " & Slovo
        Case "а вдобавок", "а именно", "а также", "а то", "а и", "благодаря тому", "в результате", "в связи", "в силу того", "все же", "всё же", "да вдобавок", "да еще", "да ещё", "да и", "даром что", "для того", "если бы", "если не", "и значит", "а именно", "и поэтому", "и притом", "и все-таки", "и всё-таки", "и следовательно", "и то", "и тогда", "и еще", "и вдобавок", "из-за того", "как скоро", "как будто", "как словно", "как только", "тому же", "кроме того", "кроме этого", "лишь бы", "лишь только"
            EtoSoyuz = True
            If Ignorirovat Then SlovoPrediduschee = PredPrediduschee: SlovoPrediduschee2 = Slovo
            Exit Function
        Case "между тем", "не столько", "не то", "не только", "невзирая на", "независимо от", "несмотря на", "но и", "но даже", "перед тем", "по мере", "по причине", "подобно тому", "пока не", "после того", "прежде чем", "при всем", "при всём", "при условии", "ради того", "раньше чем", "с тем", "тех пор", "так и", "так же", "так как", "так что", "тем более", "тех пор", "тогда как", "того что", "то есть", "то ли", "то что", "только бы", "только что", "только лишь", "только чуть", "том что"
            EtoSoyuz = True
            If Ignorirovat Then SlovoPrediduschee = PredPrediduschee: SlovoPrediduschee2 = Slovo
            Exit Function
    End Select
    EtoSoyuz = True
    Select Case Slovo
        Case "а", "благо", "буде", "будто", "вдобавок", "ввиду", "вследствие", "да", "дабы", "даже", "же", "едва"
        Case "ежели", "если", "затем", "затем", "зато", "зачем", "и", "ибо", "или ", "кабы", "как", "как-то"
        Case "когда", "коли", "ли", "либо", "лишь", "нежели", "ни", "но", "однако", "особенно", "оттого", "отчего"
        Case "пока", "покамест", "покуда", "поскольку", "потому", "почему", "притом", "причем", "пускай", "пусть"
        Case "раз", "сколько", "словно", "столько", "также", "тем", "того", "то", "тоже", "только", "точно", "хотя"
        Case "чем", "чего", "что", "чтоб", "чтобы"
        Case Else: EtoSoyuz = False
    End Select
    If Ignorirovat Then
        PredPrediduschee = SlovoPrediduschee
        SlovoPrediduschee2 = Slovo
    End If
End Function
Function EtoPredlog(PredPrediduschee, SlovoPrediduschee, SlovoPrediduschee2, Slovo, Ignorirovat As Boolean) As Boolean  ' Перечень известных предлогов
    Select Case SlovoPrediduschee2 & " " & Slovo
        Case "в течение", "в продолжение", "несмотря на", "связи с", "со стороны", "по причине", "в целях", "с целью", "по поводу", "по случаю", "в силу", "в отличие", "в заключение", "на протяжении"
            EtoPredlog = True
            If Ignorirovat Then SlovoPrediduschee = PredPrediduschee: SlovoPrediduschee2 = Slovo
            Exit Function
    End Select
    EtoPredlog = True
    Select Case Slovo
        Case "не" ' Хотя это и не предлог, но для очистки текста от частицы "не", добавил ее сюда.
        Case "у", "о", "в", "с", "к", "за", "до", "на", "по", "из", "от", "над", "под", "по-над"
        Case "про", "без", "для", "ради", "через", "из-за", "из-под", "по-за", "мимо", "вдоль"
        Case "около", "возле", "вокруг", "кругом", "после", "поперёк", "поперек", "посредством", "спустя"
        Case "близ", "вблизи", "врепеди", "перед", "спереди", "позади", "сзади"
        Case "вверху", "внизу", "сверху", "снизу", "вместо", "вроде", "насчёт", "насчет", "навстречу"
        Case "напротив", "наперекор", "накануне", "благодаря", "согласно", "вопреки", "внутри", "вследствие", "наподобие", "ввиду", "сквозь"
        Case Else: EtoPredlog = False
    End Select
    If Ignorirovat Then
        PredPrediduschee = SlovoPrediduschee
        SlovoPrediduschee2 = Slovo
    End If
End Function
Function Isklyucheniye(Slovo) As Boolean
' Перечень слов, в которых не будут обрезаться сочетания, похожие на окончания.
    Isklyucheniye = True
    Select Case Slovo
        ' В 1-ом блоке "Case"  исключения для правильной обработки союзов из двух слов.
        ' исключения в 1-2 буквы и исключения, совпадающие с однословными предлогами и союзами вносить не надо.
        Case "благодаря", "все", "всё", "даром", "кроме", "лишь", "между", "невзирая", "независимо", "несмотря", "подобно", "ради", "раньше", "тогда", "тому", "после", "прежде"
        Case "ёёёёёёёё"
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

Sub Частота_слов(Selected As Boolean, frecConst As Integer, Registr As Integer, Sort As Boolean)  ' cyberforum.ru > KoGG > Sasha_Smirnov
Dim freq: freq = frecConst   ' частота появления объекта oWord (слова или знака), больше которой они учитываются
Dim oWord As Range, dic As Object, vX As Variant, S As String, TextWords As Object
Dim t, TextStr As String             ' переменная для текста очередного "слова" (oWord.Text)
    Set dic = CreateObject("Scripting.Dictionary")
            Application.ScreenUpdating = False

With dic
        .CompareMode = Registr        ' включение чувствительности к регистру в словаре
    If Selected = True Then
        Set TextWords = Selection.Range.Words
    Else
        Set TextWords = ActiveDocument.Range.Words
    End If
    
    For Each oWord In TextWords
        t = oWord.Text
        ' замена (в слове для словаря) неразрывного пробела обычным
        If Right(t, 1) = ChrW(160) Then t = Mid(t, 1, Len(t) - 1): t = t & " " ': MsgBox t & vbCr & _
                                                                            "Len(t) = " & Len(t)
    ' добавление к "слову" (oWord.Text) пробела, если его нет, чтобы были = слова с пробелом и без
        If Right(t, 1) Like "[!" & Chr(10) & Chr(11) & Chr(13) & Chr(30) & Chr(32) & "]" Then t = t & " "
        
        S = UCase(t)
        
            Select Case AscW(S) ' проверяет юникод 1-й буквы
            Case 30, 33 To 90, 132, 133, 145 To 148, 168, 171, 184, 187, 1025, 1040 To 1071 ', _
            8211, 8212 ' цифры, нек. знаки препинания (в т. ч. тире: 8212), лат. и рус. загл. буквы
            .Item(t) = .Item(t) + 1
            End Select
            
        If AscW(t) = 8211 Then .Item(t) = .Item(t) + 1  ' подсчёт длинных дефисов (N-dash:–)
        If AscW(t) = 8212 Then .Item(t) = .Item(t) + 1  ' подсчёт длинных тире (M-dash:—)
    Next
    
    'добавление и открытие нового документа для печати туда частотного словаря
    Documents.Add
    Options.CheckGrammarAsYouType = False               ' отмена грамматического контроля ввода
    Options.CheckSpellingAsYouType = False              ' отмена орфографического контроля ввода

    With ActiveDocument    'установка полей; нумерация строк (если .LineNumbering.Active = True)
        With .PageSetup
    '    .LineNumbering.Active = True                    ' нумерация строк словаря (не абзацев!)
    '    .LineNumbering.RestartMode = wdRestartContinuous ' сплошная нумерация строк
    '    .LineNumbering.DistanceFromText = 4 'pt
        .TopMargin = CentimetersToPoints(1)
        .BottomMargin = CentimetersToPoints(1)
        .LeftMargin = CentimetersToPoints(1.1)
        .RightMargin = CentimetersToPoints(0)
        End With
    .Paragraphs.TabStops.Add CentimetersToPoints(3.5), wdAlignTabRight ' таб вправо 35 мм
    End With
    
    For Each vX In .Keys
        If .Item(vX) > freq Then Selection.TypeText RTrim(vX) & Chr(9) & .Item(vX) & Chr(13)
    Next
End With

Set dic = Nothing
    
    With Selection.PageSetup
    .TextColumns.SetCount NumColumns:=4                     ' в 4 колонки
    .TextColumns.LineBetween = True                         ' линии между колонок? Да!
    .TextColumns.Spacing = CentimetersToPoints(1.9)          ' 19 мм меж колонок
    End With
    
    With Selection
        If Sort Then
            .Sort      ' сортировка абзацев документа
        End If
        .Font.Size = 10
        .Collapse
        .Delete
    End With
    ActiveDocument.UndoClear ' очистка списка откатов во вновь созданном документе
End Sub
Public Sub ОформлениеДокументаВСтолбцыДо()
    'добавление и открытие нового документа для печати туда частотного словаря
    Documents.Add
    Options.CheckGrammarAsYouType = False               ' отмена грамматического контроля ввода
    Options.CheckSpellingAsYouType = False              ' отмена орфографического контроля ввода

    With ActiveDocument    'установка полей; нумерация строк (если .LineNumbering.Active = True)
        With .PageSetup
    '    .LineNumbering.Active = True                    ' нумерация строк словаря (не абзацев!)
    '    .LineNumbering.RestartMode = wdRestartContinuous ' сплошная нумерация строк
    '    .LineNumbering.DistanceFromText = 4 'pt
        .TopMargin = CentimetersToPoints(1)
        .BottomMargin = CentimetersToPoints(1)
        .LeftMargin = CentimetersToPoints(1.1)
        .RightMargin = CentimetersToPoints(0)
        End With
    .Paragraphs.TabStops.Add CentimetersToPoints(3.5), wdAlignTabRight ' таб вправо 35 мм
    End With
End Sub
Public Sub ОформлениеДокументаВСтолбцыПосле(NColumns As Integer, Sort As Boolean)
    With Selection.PageSetup
    .TextColumns.SetCount NumColumns:=NColumns                     ' в N колонки
    .TextColumns.LineBetween = True                         ' линии между колонок? Да!
    .TextColumns.Spacing = CentimetersToPoints(1.9)          ' 19 мм меж колонок
    End With
    
    With Selection
        If Sort Then
            .Sort      ' сортировка абзацев документа
        End If
        .Font.Size = 10
        .Collapse
        .Delete
    End With
    ActiveDocument.UndoClear ' очистка списка откатов во вновь созданном документе
End Sub

Sub СловаВНомераПоПорядку(Selected As Boolean, frecConst As Integer, Registr As Integer, Sort As Boolean)
Dim oWord As Range, vX As Variant, vX2 As Variant, S As String, TextWords As Object
Dim t, TextStr, TestStr As String             ' переменная для текста очередного "слова" (oWord.Text)
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
    .CompareMode = Registr        ' включение чувствительности к регистру в словаре
    For Each oWord In TextWords
        i = i + 1
        t = oWord.Text
        ' замена (в слове для словаря) неразрывного пробела обычным
        If Right(t, 1) = ChrW(160) Then t = Mid(t, 1, Len(t) - 1): t = t & " " ': MsgBox t & vbCr & _
                                                                            "Len(t) = " & Len(t)
    ' добавление к "слову" (oWord.Text) пробела, если его нет, чтобы были = слова с пробелом и без
        If Right(t, 1) Like "[!" & Chr(10) & Chr(11) & Chr(13) & Chr(30) & Chr(32) & "]" Then t = t & " "
        
        S = UCase(t)
        
        Select Case AscW(S) ' проверяет юникод 1-й буквы
            Case 30, 33 To 90, 132, 133, 145 To 148, 168, 171, 184, 187, 1025, 1040 To 1071 ', _
            8211, 8212 ' цифры, нек. знаки препинания (в т. ч. тире: 8212), лат. и рус. загл. буквы
            If Not .Exists(t) Then
                .Item(t) = i
            Else
                i = i - 1
            End If
        End Select
            
        If AscW(t) = 8211 Then             ' учёт длинных дефисов (N-dash:–)
            If Not .Exists(t) Then
                .Item(t) = i
            Else
                i = i - 1
            End If
        End If
        If AscW(t) = 8212 Then ' учёт длинных тире (M-dash:—)
            If Not .Exists(t) Then
                .Item(t) = i
            Else
                i = i - 1
            End If
        End If
        TextStr = TextStr + CStr(.Item(t)) + " "
        If CheckBox21.Value = True Then 'С новой строки
            TextStr = TextStr + Chr(13)
        End If
    Next
    
    TextStr = Replace(TextStr, "    ", " ")
    TextStr = Replace(TextStr, "   ", " ")
    TextStr = Replace(TextStr, "  ", " ")
    КурсорВКонец
    Selection.TypeParagraph
    Selection.Text = TextStr

    ОформлениеДокументаВСтолбцыДо 'ОформлениеДокументаВСтолбцыПосле
    For Each vX In .Keys
        Selection.TypeText RTrim(vX) & Chr(9) & .Item(vX) & Chr(13)
    Next
End With

Set dicWords = Nothing
    ОформлениеДокументаВСтолбцыПосле 4, Sort
End Sub

Sub ПредложенияВНомераПоПорядку(Selected As Boolean, frecConst As Integer, Registr As Integer, Sort As Boolean)
Dim oWord As Range, vX As Variant, vX2 As Variant, S As String, TextWords As Object
Dim t, TextStr, TestStr As String             ' переменная для текста очередного "слова" (oWord.Text)
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
    .CompareMode = Registr        ' включение чувствительности к регистру в словаре
    For Each oWord In TextWords
        i = i + 1
        t = oWord.Text
        ' замена (в слове для словаря) неразрывного пробела обычным
        If Right(t, 1) = ChrW(160) Then t = Mid(t, 1, Len(t) - 1): t = t & " " ': MsgBox t & vbCr & _
                                                                            "Len(t) = " & Len(t)
    ' добавление к "слову" (oWord.Text) пробела, если его нет, чтобы были = слова с пробелом и без
        If Right(t, 1) Like "[!" & Chr(10) & Chr(11) & Chr(13) & Chr(30) & Chr(32) & "]" Then t = t & " "
        
        S = UCase(t)
        
        Select Case AscW(S) ' проверяет юникод 1-й буквы
            Case 30, 33 To 90, 132, 133, 145 To 148, 168, 171, 184, 187, 1025, 1040 To 1071 ', _
            8211, 8212 ' цифры, нек. знаки препинания (в т. ч. тире: 8212), лат. и рус. загл. буквы
            If Not .Exists(t) Then
                .Item(t) = i
            Else
                i = i - 1
            End If
        End Select
            
        If AscW(t) = 8211 Then             ' учёт длинных дефисов (N-dash:–)
            If Not .Exists(t) Then
                .Item(t) = i
            Else
                i = i - 1
            End If
        End If
        If AscW(t) = 8212 Then ' учёт длинных тире (M-dash:—)
            If Not .Exists(t) Then
                .Item(t) = i
            Else
                i = i - 1
            End If
        End If
        TextStr = TextStr + CStr(.Item(t)) + " "
        If CheckBox21.Value = True Then 'С новой строки
            TextStr = TextStr + Chr(13)
        End If
    Next
    
    TextStr = Replace(TextStr, "    ", " ")
    TextStr = Replace(TextStr, "   ", " ")
    TextStr = Replace(TextStr, "  ", " ")
    КурсорВКонец
    Selection.TypeParagraph
    Selection.Text = TextStr

    ОформлениеДокументаВСтолбцыДо 'ОформлениеДокументаВСтолбцыПосле
    For Each vX In .Keys
        Selection.TypeText RTrim(vX) & Chr(9) & .Item(vX) & Chr(13)
    Next
End With

Set dicWords = Nothing
    ОформлениеДокументаВСтолбцыПосле 3, Sort
End Sub

Public Function СколькоГласных(ByVal Str As String)
    Dim ArrayStr: ArrayStr = Split("о, и, а, ы, ю, я, э, ё, у, е", ", ")
    Dim i As Long, j As Long, n As Long
    n = 0
    For i = 1 To Len(Str)
        'С помощью "Mid" берём отдельный символ из строки.
        'Если символ является запятой, то смотрим,
        'что находится после запятой.
        If Mid(Str, i, 1) Like "[А-я]" Then
            For j = 0 To UBound(ArrayStr)
                If Mid(Str, i, 1) = ArrayStr(j) Then
                    n = n + 1
                End If
            Next j
        End If
    Next i
    СколькоГласных = n
End Function

Public Function СколькоСогласных(ByVal Str As String)
    Dim ArrayStr: ArrayStr = Split("б, в, г, д, ж, з, й, к, л, м, н, п, р, с, т, ф, х, ц, ч, ш, щ", ", ")
    Dim i As Long, j As Long, n As Long
    n = 0
    For i = 1 To Len(Str)
        'С помощью "Mid" берём отдельный символ из строки.
        'Если символ является запятой, то смотрим,
        'что находится после запятой.
        If Mid(Str, i, 1) Like "[А-я]" Then
            For j = 0 To UBound(ArrayStr)
                If Mid(Str, i, 1) = ArrayStr(j) Then
                    n = n + 1
                End If
            Next j
        End If
    Next i
    СколькоСогласных = n
End Function

Public Function УбратьПробелы(ByVal TextStr As String)
    Dim OutStr As String
    OutStr = TextStr
    OutStr = Replace(OutStr, "    ", " ")
    OutStr = Replace(OutStr, "   ", " ")
    OutStr = Replace(OutStr, "  ", " ")
    УбратьПробелы = OutStr
End Function

Public Function УбратьТабуляции(ByVal TextStr As String)
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
    УбратьТабуляции = OutStr
End Function

Public Sub НовыйЗаголовок(ByVal HandleTextStr As String)
    Selection.Style = ActiveDocument.Styles("Заголовок 1") 'Стиль текста
    Selection.TypeText HandleTextStr
    Selection.TypeParagraph
    Selection.ClearFormatting
End Sub

    
Sub СловаВНомера(Selected As Boolean, frecConst As Integer, Registr As Integer, Sort As Boolean)
Dim oWord As Range, vX As Variant, vX2 As Variant, S As String, TextWords As Object
Dim t, TextStr, TextStrSeq, TextFromDict, TestStr As String             ' переменная для текста очередного "слова" (oWord.Text)
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

With dicFrec 'Заполнение словаря с частотами слов
    .CompareMode = Registr        ' включение чувствительности к регистру в словаре
    For Each oWord In TextWords 'Частоты слов
        t = oWord.Text
        ' замена (в слове для словаря) неразрывного пробела обычным
        If Right(t, 1) = ChrW(160) Then t = Mid(t, 1, Len(t) - 1): t = t & " " ': MsgBox t & vbCr & _
                                                                            "Len(t) = " & Len(t)
    ' добавление к "слову" (oWord.Text) пробела, если его нет, чтобы были = слова с пробелом и без
        If Right(t, 1) Like "[!" & Chr(10) & Chr(11) & Chr(13) & Chr(30) & Chr(32) & "]" Then t = t & " "
        
        S = UCase(t)
        
            Select Case AscW(S) ' проверяет юникод 1-й буквы
            Case 30, 33 To 90, 132, 133, 145 To 148, 168, 171, 184, 187, 1025, 1040 To 1071 ', _
            8211, 8212 ' цифры, нек. знаки препинания (в т. ч. тире: 8212), лат. и рус. загл. буквы
            .Item(t) = .Item(t) + 1
            End Select
            
        If AscW(t) = 8211 Then .Item(t) = .Item(t) + 1  ' подсчёт длинных дефисов (N-dash:–)
        If AscW(t) = 8212 Then .Item(t) = .Item(t) + 1  ' подсчёт длинных тире (M-dash:—)
    Next
    
    'Удаление слов с малыми частотами
    For Each vX In .Keys
        If CInt(.Item(vX)) < frecConst Then
            .Remove (vX)
        End If
    Next
    
    dicTemp.CompareMode = Registr        ' включение чувствительности к регистру в словаре
    For Each vX In .Keys 'Копия словаря частот, чтобы её не портить
        dicTemp(vX) = .Item(vX)
        WordsCountTest = CInt(.Item(vX))
        WordsCount = WordsCount + WordsCountTest 'За одно подсчёт количества слов
    Next
End With

    i = 0
    dicWords.CompareMode = Registr        ' включение чувствительности к регистру в словаре
    For Each vX In dicTemp.Keys 'Находим индекс максимальной частоты, используем, удаляем элемент из словаря
        i = i + 1
        TestStr = vX
        TestStr = dicTemp.Item(vX)
        MaxInd = FindMaxInd(dicTemp.Items())
        dicWords.Item(dicTemp.Keys()(MaxInd)) = i 'Заполнение итоговой нумерации на основе частоты. Частые начиная от 1
        dicTemp.Remove (dicTemp.Keys()(MaxInd))
    Next
    dicTemp.RemoveAll 'Очищаем словарь, потому что иногда что-то остаётся

    'Вывод в текст замененных слов
    For Each oWord In TextWords 'Частоты слов
        t = oWord.Text
        ' замена (в слове для словаря) неразрывного пробела обычным
        If Right(t, 1) = ChrW(160) Then t = Mid(t, 1, Len(t) - 1): t = t & " " ': MsgBox t & vbCr & _
                                                                            "Len(t) = " & Len(t)
    ' добавление к "слову" (oWord.Text) пробела, если его нет, чтобы были = слова с пробелом и без
        If Right(t, 1) Like "[!" & Chr(10) & Chr(11) & Chr(13) & Chr(30) & Chr(32) & "]" Then t = t & " "
        
        If CheckBox20.Value = True Then 'Вероятность вместо номеров
            TextFromDict = CStr(dicFrec.Item(t) / WordsCount)
        Else
            TextFromDict = CStr(dicWords.Item(t))
        End If
        If TextFromDict = " " Or TextFromDict = "0" Then 'Если слова найдено не было, то присваиваем вместо него 0
            TextFromDict = ""
        End If
        TextStrSeq = TextStrSeq & TextFromDict & " "
        If CheckBox20.Value = True Then 'Вероятность вместо номеров
            If t = ". " Then
                TextStrSeq = TextStrSeq & Chr(13)
            End If
        End If
                
        TextStr = TextStr & t & Chr(9) & TextFromDict & " "
        If CheckBox21.Value = True Then 'С новой строки
            TextStr = TextStr & Chr(13)
        End If
    Next
    TextStr = УбратьПробелы(TextStr)
    TextStrSeq = УбратьПробелы(TextStrSeq)
    TextStr = УбратьТабуляции(TextStr)
    TextStrSeq = УбратьТабуляции(TextStrSeq)
    
    ОформлениеДокументаВСтолбцыДо 'ОформлениеДокументаВСтолбцыПосле
    НовыйЗаголовок "Словарь"
    Selection.TypeText "Слово" & Chr(9) & "Номера по частоте" & Chr(9) & "Частота" & Chr(9) & "Вероятность" & Chr(9) & "Длина" & Chr(13)
    For Each vX In dicWords.Keys
        If dicWords.Item(vX) <> "" Then
            Selection.TypeText RTrim(vX) & Chr(9) & dicWords.Item(vX) & Chr(9) & dicFrec.Item(vX) & Chr(9) & CStr(dicFrec.Item(vX) / WordsCount) & Chr(9) & Len(RTrim(vX)) & Chr(13)
        End If
    Next
    
    Selection.InsertBreak Type:=wdPageBreak
    НовыйЗаголовок "Слова"
    'КурсорВКонец
    'Selection.TypeParagraph
    Selection.TypeText TextStr
    
    Selection.InsertBreak Type:=wdPageBreak
    НовыйЗаголовок "Номера"
    Selection.TypeText TextStr 'TextStrSeq
    
    Selection.InsertBreak Type:=wdPageBreak
    НовыйЗаголовок "Предложения"
    TestStr = Replace(TextStrSeq, " " & dicWords.Item(". ") & " ", Chr(13))
    TestStr = Replace(TestStr, " ", Chr(9))
    Selection.TypeText TestStr
    
    ОформлениеДокументаВСтолбцыПосле 2, Sort
        
    Set dicFrec = Nothing
    Set dicWords = Nothing
    Set dicTemp = Nothing
End Sub

Sub ПредложенияВНомера(Selected As Boolean, frecConst As Integer, Registr As Integer, Sort As Boolean)
Dim oWord As Range, vX As Variant, vX2 As Variant, S As String, TextWords As Object
Dim t, TextStr, TextStrSeq, TextFromDict, TestStr As String             ' переменная для текста очередного "слова" (oWord.Text)
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

With dicFrec 'Заполнение словаря с частотами слов
    .CompareMode = Registr        ' включение чувствительности к регистру в словаре
    For Each oWord In TextWords 'Частоты слов
        t = oWord.Text
        t = ПроверитьПереносСтроки(t)
        ' замена (в слове для словаря) неразрывного пробела обычным
        If Right(t, 1) = ChrW(160) Then t = Mid(t, 1, Len(t) - 1): t = t & " " ': MsgBox t & vbCr & _
                                                                            "Len(t) = " & Len(t)
    ' добавление к "слову" (oWord.Text) пробела, если его нет, чтобы были = слова с пробелом и без
        If Right(t, 1) Like "[!" & Chr(10) & Chr(11) & Chr(13) & Chr(30) & Chr(32) & "]" Then t = t & " "
        
        S = UCase(t)
        
            Select Case AscW(S) ' проверяет юникод 1-й буквы
            Case 30, 33 To 90, 132, 133, 145 To 148, 168, 171, 184, 187, 1025, 1040 To 1071 ', _
            8211, 8212 ' цифры, нек. знаки препинания (в т. ч. тире: 8212), лат. и рус. загл. буквы
            .Item(t) = .Item(t) + 1
            End Select
            
        If AscW(t) = 8211 Then .Item(t) = .Item(t) + 1  ' подсчёт длинных дефисов (N-dash:–)
        If AscW(t) = 8212 Then .Item(t) = .Item(t) + 1  ' подсчёт длинных тире (M-dash:—)
    Next
    
    'Удаление слов с малыми частотами
    For Each vX In .Keys
        If CInt(.Item(vX)) < frecConst Then
            .Remove (vX)
        End If
    Next
    
    dicTemp.CompareMode = Registr        ' включение чувствительности к регистру в словаре
    For Each vX In .Keys 'Копия словаря частот, чтобы её не портить
        dicTemp(vX) = .Item(vX)
        WordsCountTest = CInt(.Item(vX))
        WordsCount = WordsCount + WordsCountTest 'За одно подсчёт количества слов
    Next
End With

    i = 0
    dicWords.CompareMode = Registr        ' включение чувствительности к регистру в словаре
    For Each vX In dicTemp.Keys 'Находим индекс максимальной частоты, используем, удаляем элемент из словаря
        i = i + 1
        TestStr = vX
        TestStr = dicTemp.Item(vX)
        MaxInd = FindMaxInd(dicTemp.Items())
        dicWords.Item(dicTemp.Keys()(MaxInd)) = i 'Заполнение итоговой нумерации на основе частоты. Частые начиная от 1
        dicTemp.Remove (dicTemp.Keys()(MaxInd))
    Next
    dicTemp.RemoveAll 'Очищаем словарь, потому что иногда что-то остаётся

    'Вывод в текст замененных слов
    For Each oWord In TextWords 'Частоты слов
        t = oWord.Text
        t = ПроверитьПереносСтроки(t)
        ' замена (в слове для словаря) неразрывного пробела обычным
        If Right(t, 1) = ChrW(160) Then t = Mid(t, 1, Len(t) - 1): t = t & " " ': MsgBox t & vbCr & _
                                                                            "Len(t) = " & Len(t)
    ' добавление к "слову" (oWord.Text) пробела, если его нет, чтобы были = слова с пробелом и без
        If Right(t, 1) Like "[!" & Chr(10) & Chr(11) & Chr(13) & Chr(30) & Chr(32) & "]" Then t = t & " "
        
        If CheckBox20.Value = True Then 'Вероятность вместо номеров
            TextFromDict = CStr(dicFrec.Item(t) / WordsCount)
        Else
            TextFromDict = CStr(dicWords.Item(t))
        End If
        If TextFromDict = " " Or TextFromDict = "0" Then 'Если слова найдено не было, то присваиваем вместо него 0
            TextFromDict = ""
        End If
        TextStrSeq = TextStrSeq & TextFromDict & " "
        If CheckBox20.Value = True Then 'Вероятность вместо номеров
            If t = ". " Then
                TextStrSeq = TextStrSeq & Chr(13)
            End If
        End If
                
        TextStr = TextStr & t & Chr(9) & TextFromDict & " "
        If CheckBox21.Value = True Then 'С новой строки
            TextStr = TextStr & Chr(13)
        End If
    Next
    TextStr = УбратьПробелы(TextStr)
    TextStrSeq = УбратьПробелы(TextStrSeq)
    TextStr = УбратьТабуляции(TextStr)
    TextStrSeq = УбратьТабуляции(TextStrSeq)
    
    ОформлениеДокументаВСтолбцыДо 'ОформлениеДокументаВСтолбцыПосле
    НовыйЗаголовок "Словарь"
    Selection.TypeText "Слово" & Chr(9) & "Номера по частоте" & Chr(9) & "Частота" & Chr(9) & "Вероятность" & Chr(9) & "Длина" & Chr(13)
    For Each vX In dicWords.Keys
        If dicWords.Item(vX) <> "" Then
            Selection.TypeText RTrim(vX) & Chr(9) & dicWords.Item(vX) & Chr(9) & dicFrec.Item(vX) & Chr(9) & CStr(dicFrec.Item(vX) / WordsCount) & Chr(9) & Len(RTrim(vX)) & Chr(13)
        End If
    Next
    
    Selection.InsertBreak Type:=wdPageBreak
    НовыйЗаголовок "Слова"
    'КурсорВКонец
    'Selection.TypeParagraph
    Selection.TypeText TextStr
    
    Selection.InsertBreak Type:=wdPageBreak
    НовыйЗаголовок "Номера"
    Selection.TypeText TextStr 'TextStrSeq
    
    Selection.InsertBreak Type:=wdPageBreak
    НовыйЗаголовок "Предложения"
    'TestStr = Replace(TextStrSeq, ". ", Chr(13))
    'TestStr = Replace(TextStrSeq, ". ", "")
    TestStr = Replace(TextStrSeq, " ", Chr(13))
    Selection.TypeText TestStr
    
    ОформлениеДокументаВСтолбцыПосле 2, Sort
        
    Set dicFrec = Nothing
    Set dicWords = Nothing
    Set dicTemp = Nothing
End Sub

Public Function СимволыВИхНомера(ByVal TextStr As String, ByVal DelPharagraph)
'Dim AlfU() As String
'Dim AlfL() As String
'Dim EngAlfL() As String
'Dim AlfPrep(0 To 7) As String
    Dim i
    Dim AlfPrep: AlfPrep = Split(". . ? ! ; : ( ) * -", " "): AlfPrep(0) = Chr$(13) ' vbCr - символ абзаца
    Dim AlfSpecChars: AlfSpecChars = Split("@ # № $ % ^ & _ + = \ / | ~ ` < > "" ' ’ [ ] { } …", " ")
    Dim AlfNums: AlfNums = Split("1 2 3 4 5 6 7 8 9 0", " ")
    Dim AlfNumsStr: AlfNumsStr = Split("Один Два Три Четыре Пять Шесть Семь Восемь Девять Ноль", " ")
    
    Dim AlfU: AlfU = Split("А Б В Г Д Е Ё Ж З И Й К Л М Н О П Р С Т У Ф Х Ц Ч Ш Щ Ъ Ы Ь Э Ю Я", " ")
    Dim AlfL: AlfL = Split("А Б В Г Д Е Ё Ж З И Й К Л М Н О П Р С Т У Ф Х Ц Ч Ш Щ Ъ Ы Ь Э Ю Я", " ")
    
    Dim EngAlfU: EngAlfU = Split("A B C D E F G H I J K L M N O P Q R S T U V W X Y Z", " ")
    Dim EngAlfL: EngAlfL = Split("A B C D E F G H I J K L M N O P Q R S T U V W X Y Z", " ")
    
    For i = 0 To UBound(AlfU)
        AlfL(i) = LCase(AlfU(i))
    Next i
    For i = 0 To UBound(EngAlfU)
        EngAlfL(i) = LCase(EngAlfU(i))
    Next i

    Dim Str: Str = TextStr
    'Вначале заменяем символы "., ", т.к. иначе они будут снова заменяться
    For i = 0 To UBound(AlfNums) 'Заменяем номера
        Str = Replace(Str, AlfNums(i), AlfNumsStr(i))
    Next i
    For i = 0 To UBound(AlfPrep) 'Заменяем препинания.
        Str = Replace(Str, AlfPrep(i), Asc(AlfPrep(i)) & ", ")
    Next i
    For i = 0 To UBound(AlfSpecChars) 'Заменяем спец символы
        Str = Replace(Str, AlfSpecChars(i), Asc(AlfSpecChars(i)) & ", ")
    Next i
    
    
    For i = 0 To UBound(AlfU) 'Заменяем малые буквы русские
        Str = Replace(Str, AlfL(i), Asc(AlfL(i)) & ", ", , , vbBinaryCompare)
    Next i

    For i = 0 To UBound(AlfU) 'Заменяем большие буквы русские
        Str = Replace(Str, AlfU(i), Asc(AlfU(i)) & ", ", , , vbBinaryCompare)
    Next i
    
    For i = 0 To UBound(EngAlfU) 'Заменяем малые буквы англ
        Str = Replace(Str, EngAlfL(i), Asc(EngAlfL(i)) & ", ", , , vbBinaryCompare)
    Next i
    
    For i = 0 To UBound(EngAlfU) 'Заменяем большие буквы англ
        Str = Replace(Str, EngAlfU(i), Asc(EngAlfU(i)) & ", ", , , vbBinaryCompare)
    Next i
    
    Str = Replace(Str, "  ", " ")
    Str = Replace(Str, ", ,", ", " & Asc(",") & ",")
    Str = Replace(Str, " , ", ", ")
    Str = Replace(Str, " ,", ",")
    If DelPharagraph Then
        Str = Replace(Str, "^p", "")
    End If
    
    СимволыВИхНомера = Str
    'Asc(s) 'Вернёт численное значение символа "Б", равное 193
    'Chr(Asc(s)) 'Преобразует численное значение символа в его букву
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
    Частота_слов Selected, frecConst, Registr, Sort
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
    StrText = ПроверитьПереносСтроки(StrText)
    MsgBox ("Гласных:   " & Chr(9) & СколькоГласных(StrText) & Chr(13) & "Согласных: " & Chr(9) & СколькоСогласных(StrText))
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
    СловаВНомераПоПорядку Selected, frecConst, Registr, Sort
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
    ПредложенияВНомераПоПорядку Selected, frecConst, Registr, Sort
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
    ПредложенияВНомера Selected, frecConst, Registr, Sort
End Sub

Private Sub CommandButton15_Click()
    КурсорВНачало
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

Public Sub БуквыВРамку(Glasnie As Boolean)
Dim symbol As String
Dim Array_Str() As String
Dim TextWords As String
Dim i, j
Dim Array_Glasn: Array_Glasn = Split("о, и, а, ы, ю, я, э, ё, у, е", ", ")
Dim Array_Sogl: Array_Sogl = Split("б, в, г, д, ж, з, й, к, л, м, н, п, р, с, т, ф, х, ц, ч, ш, щ", ", ")
    
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
    БуквыВРамку CheckBox23.Value
End Sub

Private Sub CommandButton2_Click()
    Частота_сочетаний_слов
End Sub

Private Sub CommandButton3_Click()
    Dim StrText: StrText = Selection.Text
    If MsgBox("Всё в одну строчку (удалить символ абзаца)?", vbYesNo, "Всё в одну строчку?") = vbYes Then
        Dim DelParagraph: DelParagraph = True
    End If
    Dim Str: Str = СимволыВИхНомера(StrText, DelParagraph)
    КурсорВКонец
    Selection.TypeParagraph
    Selection.Text = Str
End Sub

Function Частота_ЧастиСлова(WordFind As String, Selected As Boolean, Registr As Integer)
Dim oWord As Range, dic As Object, vX As Variant, S As String, TextWords As Object
Dim t, WordText As String             ' переменная для текста очередного "слова" (oWord.Text)
Set dic = CreateObject("Scripting.Dictionary")
Application.ScreenUpdating = False

With dic
        .CompareMode = Registr        ' включение чувствительности к регистру в словаре
    If Selected = True Then
        Set TextWords = Selection.Range.Words
    Else
        Set TextWords = ActiveDocument.Range.Words
    End If
    WordFind = Trim(WordFind) & " "
    Registr = 1
    
    For Each oWord In TextWords
        t = oWord.Text
        ' замена (в слове для словаря) неразрывного пробела обычным
        If Right(t, 1) = ChrW(160) Then t = Mid(t, 1, Len(t) - 1): t = t & " " ': MsgBox t & vbCr & _
                                                                            "Len(t) = " & Len(t)
    ' добавление к "слову" (oWord.Text) пробела, если его нет, чтобы были = слова с пробелом и без
        If Right(t, 1) Like "[!" & Chr(10) & Chr(11) & Chr(13) & Chr(30) & Chr(32) & "]" Then t = t & " "
        S = t
        If Registr = 1 Then
            S = UCase(t)
            WordText = UCase(WordFind)
        End If
        If InStr(Trim(S), Trim(WordText)) > 0 Then 'WordFind - искомое слово. 'Если нет совпадений, то вернёт 0
            .Item(WordFind) = .Item(WordFind) + 1
        End If
    Next
    
    Частота_ЧастиСлова = .Item(WordFind)
End With

Set dic = Nothing
End Function

Function Частота_ОдногоСлова(ByVal WordFind As String, ByVal SearchPart As Boolean, ByVal Selected As Boolean, ByVal Registr As Integer)
Dim oWord As Range, dic As Object, vX As Variant, S As String, TextWords As Object
Dim t, WordText As String             ' переменная для текста очередного "слова" (oWord.Text)
Set dic = CreateObject("Scripting.Dictionary")
Application.ScreenUpdating = False

With dic
        .CompareMode = Registr        ' включение чувствительности к регистру в словаре
    If Selected = True Then
        Set TextWords = Selection.Range.Words
    Else
        Set TextWords = ActiveDocument.Range.Words
    End If
    WordFind = Trim(WordFind) & " "
    Registr = 1
    
    For Each oWord In TextWords
        t = oWord.Text
        ' замена (в слове для словаря) неразрывного пробела обычным
        If Right(t, 1) = ChrW(160) Then t = Mid(t, 1, Len(t) - 1): t = t & " " ': MsgBox t & vbCr & _
                                                                            "Len(t) = " & Len(t)
    ' добавление к "слову" (oWord.Text) пробела, если его нет, чтобы были = слова с пробелом и без
        If Right(t, 1) Like "[!" & Chr(10) & Chr(11) & Chr(13) & Chr(30) & Chr(32) & "]" Then t = t & " "
        S = t
        If Registr = 1 Then
            S = UCase(t)
            WordText = UCase(WordFind)
        End If
        If SearchPart Then 'Ищем только часть слова, т.е. слова, которые включают введённую часть
            If InStr(Trim(S), Trim(WordText)) > 0 Then 'WordFind - искомое слово. 'Если нет совпадений, то вернёт 0
                .Item(WordFind) = .Item(WordFind) + 1
            End If
        Else
            If S = WordText Then 'WordFind - искомое слово
                .Item(WordFind) = .Item(WordFind) + 1
            End If
        End If
    Next
    
    Частота_ОдногоСлова = .Item(WordFind)
End With

Set dic = Nothing
End Function

Sub СтатистикаПоявленияСловаЗаДень(ByVal WordFind As String, ByVal SearchPart As Boolean, ByRef xlSheet As Object, ByVal MatchCase As Boolean)
    Dim FindStyle
    FindStyle = "Заголовок 3;З_День"
    СтатистикаПоявленияСловаЗаПромежуток WordFind, SearchPart, FindStyle, xlSheet, MatchCase
End Sub

Sub СтатистикаПоявленияСловаЗаМесяц(ByVal WordFind As String, ByVal SearchPart As Boolean, ByRef xlSheet As Object, ByVal MatchCase As Boolean)
    Dim FindStyle
    FindStyle = "Заголовок 2;З_Мес"
    СтатистикаПоявленияСловаЗаПромежуток WordFind, SearchPart, FindStyle, xlSheet, MatchCase
End Sub

Sub СтатистикаПоявленияСловаЗаПромежуток(ByVal WordFind As String, ByVal SearchPart As Boolean, ByVal FindStyle As String, ByRef xlSheet As Object, ByVal MatchCase As Boolean)
    'Промежуток задаётся названием стиля оформления текста
    Dim Words(), MaxCount, i, n_Words As Integer
    Dim WordsDate(), S, s_Words ', FindStyle As String
    'ВыделитьМеждуПозициями
    Dim rgePages As Range
    Dim cursorEnd As Range
    MaxCount = ЗаголовковВДневнике(FindStyle) 'Максимальное количество заголовков
    'FindStyle = "Заголовок 2;З_Мес"
    КурсорВНачало
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
        s_Words = Trim(Left(S, Len(S) - 1)) 'удаление символа переноса строк (новой строке присвоить на символ меньше)
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
        n_Words = Частота_ОдногоСлова(WordFind, SearchPart, True, 1)
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
        КурсорВКонец
        Selection.TypeParagraph
        If SearchPart Then
            Selection.TypeText Text:="Частота слов '*" & Trim(WordFind) & "*' за промежуток: " & "(" + Trim(Str(MaxCount)) + ")"
        Else
            Selection.TypeText Text:="Частота слов '" & Trim(WordFind) & "' за промежуток: " & "(" + Trim(Str(MaxCount)) + ")"
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

Private Sub ИскатьСлово(ByVal WordFind As String, ByRef xlSheet As Object)  'Частота заданного слова
    'WordFind = Trim(TextBox4.Text) 'Какое слово искать
    Dim Selected, SearchPart As Boolean, WordFrec As Integer, Registr As Integer
    Dim MatchCase As Boolean
    If WordFind = "" Then
        MsgBox ("Не задано искомое слово")
        Exit Sub
    End If
    Dim rgePages As Range
    Set rgePages = Selection.Range
    
    Selected = CheckBox1.Value 'Ищем в выделенной части
    SearchPart = CheckBox5.Value 'Ищем часть слова
    If Selected And Not CheckBox4.Value And Len(Selection.Text) = 1 Then
        MsgBox ("Выделите текст или снимите галку с 'Только выделенный текст'.")
        Exit Sub
    End If
    
    If CheckBox4.Value = True Then 'Слово за день для всего текста
        СтатистикаПоявленияСловаЗаДень WordFind, SearchPart, xlSheet, MatchCase
        'MsgBox (Trim(WordFind) & ": " & WordFrec)
    ElseIf CheckBox6.Value = True Then 'Слово за месяц для всего текста
        СтатистикаПоявленияСловаЗаМесяц WordFind, SearchPart, xlSheet, MatchCase
        'MsgBox (Trim(WordFind) & ": " & WordFrec)
    Else
        WordFrec = Частота_ОдногоСлова(WordFind, SearchPart, Selected, 1)
        MsgBox (Trim(WordFind) & ": " & WordFrec)
    End If
    rgePages.Select
End Sub

Private Sub CommandButton4_Click() 'Частота заданного слова
    Dim Words, wordString
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    xlSheet.Application.Cells(1, 1).Value = "Слово\Частота за дату"
    If CheckBox8.Value = True Then
        If CheckBox10.Value = True Then
            xlSheet.Application.Visible = True
        End If
    End If
    
    Dim sngStart: sngStart = Timer              ' Начало отсчёта
    If CheckBox7.Value = True Then 'Несколько слов
        Words = Split(Selection.Text, vbCr) 'Разбитие новой строкой. Ещё есть vbNewLine vbCrLf
        For Each wordString In Words
            If wordString <> "" Then
                ИскатьСлово wordString, xlSheet
            End If
        Next wordString
    Else
        ИскатьСлово Trim(TextBox4.Text), xlSheet
    End If
    Dim sngEnd: sngEnd = Timer              ' Конец отсчёта
    Dim sngElapsed: sngElapsed = Format(sngEnd - sngStart, "Fixed") ' Приращение.
    Label4.Caption = "Подсчёт занял: " & sngElapsed & " секунд."
    If CheckBox8.Value = True And xlApp <> Null Then
        If CheckBox10.Value = True Then
            xlSheet.Application.Visible = True
        End If
        xlSheet.Application.Visible = True
    End If
End Sub

Sub УдалитьСтранныеСимволы()
    Selection.find.ClearFormatting
    Selection.find.Replacement.ClearFormatting
    With Selection.find
        .Text = "«"
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
        .Text = "»"
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
    УдалитьСтранныеСимволы
    Dim StrText: StrText = Selection.Text
    Dim StrNumsIO
    Dim Normalize
    If MsgBox("Всё в одну строчку (удалить символ абзаца)?", vbYesNo, "Всё в одну строчку?") = vbYes Then
        Dim DelParagraph: DelParagraph = True
    End If
    Normalize = CheckBox11.Value
    StrNumsIO = РазбитьСловаНаЦифровыеПары(StrText, DelParagraph, Normalize)
    
    If CheckBox12.Value Then
        НовыйОформленныйДокумент
        Selection.TypeText StrNumsIO
    Else
        КурсорВКонец
        Selection.TypeParagraph
        Selection.Text = StrNumsIO
    End If
End Sub

Public Function FindMaxInd(Arr) 'От нуля в общем случае
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

Public Function РазбитьСловаНаЦифровыеПары(ByVal StrText As String, ByVal DelParagraph As Boolean, ByVal Normalize As Boolean)
    Dim i
    Dim Str: Str = StrText
    Str = СимволыВИхНомера(Str, DelParagraph)

    Dim StrNums: StrNums = Split(Str, ", ") 'Массив чисел в виде строк
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
    РазбитьСловаНаЦифровыеПары = StrNumsIO
End Function

Public Function НомераВСимволы(ByVal Str As String)
    'Dim StrNums: StrNums = Str
    Dim StrNumsLines: StrNumsLines = Split(Str, vbCr)
    Dim StrNums ': StrNums = Split(Str, ", ")
    Dim StrNumStr, i, j
    For i = 0 To UBound(StrNumsLines)
        If (StrNumsLines(i) <> "") And (StrNumsLines(i) <> vbNewLine) Then
            StrNums = Split(StrNumsLines(i), ", ")
            For j = 0 To UBound(StrNums)
                StrNums(j) = Trim(StrNums(j)) 'ПроверитьПереносСтроки(StrNums(i))
                If (StrNums(j) <> "") Then
                    StrNums(j) = Chr(CLng(StrNums(j)))
                    If (StrNums(j) = vbCr) Then 'Если в строке перенос строки
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
    НомераВСимволы = StrNumStr
End Function

Public Function ПроверитьПереносСтроки(ByVal Str As String)
    Dim StrText: StrText = Str
    If Mid(StrText, Len(StrText), 1) = vbNewLine Then
        StrText = Left(StrText, Len(StrText) - 1) 'удалить последний символ
    End If
    If Mid(StrText, Len(StrText), 1) = vbCr Then
        StrText = Left(StrText, Len(StrText) - 1) 'удалить последний символ
    End If
    If Mid(StrText, Len(StrText), 1) = vbCrLf Then
        StrText = Left(StrText, Len(StrText) - 1) 'удалить последний символ
    End If
    ПроверитьПереносСтроки = StrText
End Function

Private Sub CommandButton6_Click()
    Dim StrText: StrText = Selection.Text
    StrText = ПроверитьПереносСтроки(StrText)
    Dim Str: Str = НомераВСимволы(StrText)
    КурсорВКонец
    Selection.TypeParagraph
    Selection.Text = Str
    'Asc(s) 'Вернёт численное значение символа "Б", равное 193
    'Chr(Asc(s)) 'Преобразует численное значение символа в его букву
End Sub

Public Function РазбитьТекстНаПарыСлов(Selected As Boolean, frecConst As Integer, Registr As Integer, Sort As Boolean)
Dim freq: freq = frecConst   ' частота появления объекта oWord (слова или знака), больше которой они учитываются
Dim oWord As Range, dic As Object, vX As Variant, S As String, TextWords As Object
Dim t, ResultText As String             ' переменная для текста очередного "слова" (oWord.Text)
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
        If t Like "*" Then '"*[А-яA-z]*" @#№$%^&_+=\/|~`<>""'’{}….?!;:()*-1234567890
            ' замена (в слове для словаря) неразрывного пробела обычным
            If Right(t, 1) = ChrW(160) Then t = Mid(t, 1, Len(t) - 1): t = t & " " ': MsgBox t & vbCr & _
                                                                                "Len(t) = " & Len(t)
        ' добавление к "слову" (oWord.Text) пробела, если его нет, чтобы были = слова с пробелом и без
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
    
    РазбитьТекстНаПарыСлов = ResultText
End Function

Private Sub НовыйОформленныйДокумент()
    'добавление и открытие нового документа для печати туда частотного словаря
    Documents.Add
    Options.CheckGrammarAsYouType = False               ' отмена грамматического контроля ввода
    Options.CheckSpellingAsYouType = False              ' отмена орфографического контроля ввода

    With ActiveDocument    'установка полей; нумерация строк (если .LineNumbering.Active = True)
        With .PageSetup
    '    .LineNumbering.Active = True                    ' нумерация строк словаря (не абзацев!)
    '    .LineNumbering.RestartMode = wdRestartContinuous ' сплошная нумерация строк
    '    .LineNumbering.DistanceFromText = 4 'pt
        .TopMargin = CentimetersToPoints(1)
        .BottomMargin = CentimetersToPoints(1)
        .LeftMargin = CentimetersToPoints(1.1)
        .RightMargin = CentimetersToPoints(0)
        End With
    .Paragraphs.TabStops.Add CentimetersToPoints(3.5), wdAlignTabRight ' таб вправо 35 мм
    End With
    
    With Selection.PageSetup
    .TextColumns.SetCount NumColumns:=4                     ' в 4 колонки
    .TextColumns.LineBetween = True                         ' линии между колонок? Да!
    .TextColumns.Spacing = CentimetersToPoints(1.9)          ' 19 мм меж колонок
    End With
    
    With Selection
        .Font.Size = 10
        .Collapse
        .Delete
    End With
    ActiveDocument.UndoClear ' очистка списка откатов во вновь созданном документе
End Sub

Private Sub CommandButton7_Click()
    УдалитьСтранныеСимволы
    Dim TextStr
    TextStr = РазбитьТекстНаПарыСлов(True, 0, True, False)
    If CheckBox12.Value Then
        НовыйОформленныйДокумент
        Selection.TypeText TextStr
    Else
        КурсорВКонец
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
    СловаВНомера Selected, frecConst, Registr, Sort
End Sub

Private Sub CommandButton9_Click()
    КурсорВНачало
    Dim Handles, RemoveParam, RemoveTextParam, i
    Handles = Split("Заголовок 4;З_Момент#Заголовок 3;З_День#Заголовок 2;З_Мес#Заголовок 1;З_Год", "#")
    If CheckBox13.Value = True Then 'Удалить заголовки
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
    If CheckBox18.Value = True Then RemoveTextParam = RemoveTextParam + "@ # № $ % ^ & _ + = \ / | ~ ` < > "" ' ’ [ ] { } … "
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
