
Option Explicit

' Модуль для відмінювання прізвищ, імен та по-батькові у всіх 7 відмінках


Option Explicit

' --- Логіка з файлу Option_Explicit_Corrected.bas ---

Option Explicit


' Список голосних української мови
Private Const vowels As String = "аеиоуіїєюя"

' Список приголосних української мови
Private Const consonant As String = "бвгджзйклмнпрстфхцчшщ"

' Українські шиплячі приголосні
Private Const shyplyachi As String = "жчшщ"

' Українські завжди м’які звуки
Private Const myaki As String = "ьюяєї"

' Українські губні звуки
Private Const gubni As String = "мвпбф"

' Глухі приголосні
Private Const gluhi As String = "кпстфхцчшщ"

' Дзвінкі приголосні
Private Const dzvinki As String = "бвгджзлмнр"

' Функція перевірки, чи є символ голосною
Private Function IsVowel(char As String) As Boolean
    IsVowel = InStr(1, vowels, char, vbTextCompare) > 0
End Function

' Функція перевірки, чи є символ приголосним
Private Function IsConsonant(char As String) As Boolean
    IsConsonant = InStr(1, consonant, char, vbTextCompare) > 0
End Function

' Функція перевірки, чи є символ шиплячим
Private Function IsShyplyachi(char As String) As Boolean
    IsShyplyachi = InStr(1, shyplyachi, char, vbTextCompare) > 0
End Function

' Функція перевірки, чи є символ губним
Private Function IsGubni(char As String) As Boolean
    IsGubni = InStr(1, gubni, char, vbTextCompare) > 0
End Function

' Функція перевірки, чи є символ апострофом
Private Function IsApostrof(char As String) As Boolean
    IsApostrof = char = "'"
End Function

' Функція для чергування г-к-х > з-ц-с
Private Function InverseGKH(letter As String) As String
    Select Case letter
        Case "г": InverseGKH = "з"
        Case "к": InverseGKH = "ц"
        Case "х": InverseGKH = "с"
        Case Else: InverseGKH = letter
    End Select
End Function

' Функція для чергування г-к > ж-ч
Private Function Inverse2(letter As String) As String
    Select Case letter
        Case "г": Inverse2 = "ж"
        Case "к": Inverse2 = "ч"
        Case Else: Inverse2 = letter
    End Select
End Function

' Функція для чергування дзвінких та глухих приголосних
Private Function ToggleSound(letter As String) As String
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")
    
    map.Add "б", "п": map.Add "п", "б"
    map.Add "в", "ф": map.Add "ф", "в"
    map.Add "г", "к": map.Add "к", "г"
    map.Add "д", "т": map.Add "т", "д"
    map.Add "з", "с": map.Add "с", "з"
    
    If map.exists(letter) Then
        ToggleSound = map(letter)
    Else
        ToggleSound = letter
    End If
End Function

' Функція визначення групи для іменників 2-ї відміни
Private Function Detect2Group(ByVal word As String) As Integer
    Dim osnova As String
    Dim stack() As String
    Dim i As Integer, stackSize As Integer
    Dim osnovaEnd As String, lastVowel As String
    Dim vowelsAndSoft As String
    Dim osnovaLen As Integer
    
    osnova = word
    vowelsAndSoft = vowels & "ь" ' Голосні + м'який знак
    
    ' Ініціалізуємо стек
    ReDim stack(Len(word))
    i = 0

    ' Відокремлюємо голосні та м'який знак
    Do While InStr(1, vowelsAndSoft, Right(osnova, 1), vbTextCompare) > 0
        stack(i) = Right(osnova, 1)
        osnova = Left(osnova, Len(osnova) - 1)
        i = i + 1
    Loop

    stackSize = i
    lastVowel = "Z" ' Нульове закінчення за замовчуванням

    ' Визначення останнього елемента у стеку (голосної або м'якого знака)
    If stackSize > 0 Then
        lastVowel = stack(stackSize - 1)
    End If

    ' Отримуємо останній символ основи
    osnovaEnd = Right(osnova, 1)
    
    ' Визначаємо групу
    If InStr(1, neshyplyachi, osnovaEnd, vbTextCompare) > 0 And InStr(1, myaki, lastVowel, vbTextCompare) = 0 Then
        Detect2Group = 1 ' Тверда група
    ElseIf InStr(1, shyplyachi, osnovaEnd, vbTextCompare) > 0 And InStr(1, myaki, lastVowel, vbTextCompare) = 0 Then
        Detect2Group = 2 ' Мішана група
    Else
        Detect2Group = 3 ' М'яка група
    End If
End Function

' Функція для пошуку першої голосної з кінця
Private Function FirstLastVowel(ByVal word As String, ByVal vowels As String) As String
    Dim i As Integer
    For i = Len(word) To 1 Step -1
        If InStr(1, vowels, Mid(word, i, 1), vbTextCompare) > 0 Then
            FirstLastVowel = Mid(word, i, 1)
            Exit Function
        End If
    Next i
    FirstLastVowel = "" ' Повертає порожній рядок, якщо голосна не знайдена
End Function

' Функція для визначення основи слова
Private Function GetOsnova(ByVal word As String) As String
    Dim osnova As String
    osnova = word

    ' Видаляємо всі голосні та м'який знак із кінця слова
    Do While InStr(1, vowels & "ь", Right(osnova, 1), vbTextCompare) > 0
        osnova = Left(osnova, Len(osnova) - 1)
    Loop

    GetOsnova = osnova
End Function


' Функція відмінювання чоловічих та жіночих імен, що закінчуються на -а, -я
Private Function ManRule1(ByVal workingWord As String) As Boolean
    Dim BeforeLast As String
    BeforeLast = Mid(workingWord, Len(workingWord) - 1, 1)

    ' Остання літера -а
    If Right(workingWord, 1) = "а" Then
        Call WordForms(workingWord, Array(BeforeLast & "и", InverseGKH(BeforeLast) & "і", BeforeLast & "у", BeforeLast & "ою", InverseGKH(BeforeLast) & "і", BeforeLast & "о"), 2)
        Call Rule(101)
        ManRule1 = True
        Exit Function
    End If

    ' Остання літера -я
    If Right(workingWord, 1) = "я" Then
        If BeforeLast = "і" Then
            Call WordForms(workingWord, Array("ї", "ї", "ю", "єю", "ї", "є"), 1)
            Call Rule(102)
            ManRule1 = True
        Else
            Call WordForms(workingWord, Array(BeforeLast & "і", InverseGKH(BeforeLast) & "і", BeforeLast & "ю", BeforeLast & "ею", InverseGKH(BeforeLast) & "і", BeforeLast & "е"), 2)
            Call Rule(103)
            ManRule1 = True
        End If
    End If
End Function

' Функція відмінювання чоловічих імен, що закінчуються на -р
Private Function ManRule2(ByVal workingWord As String) As Boolean
    If Right(workingWord, 1) = "р" Then
        If InNames(workingWord, Array("Ігор", "Лазар")) Then
            Call WordForms(workingWord, Array("я", "еві", "я", "ем", "еві", "е"))
            Call Rule(201)
            ManRule2 = True
        Else
            Dim osnova As String
            osnova = workingWord
            If Mid(osnova, Len(osnova) - 1, 1) = "і" Then
                osnova = Left(osnova, Len(osnova) - 2) & "о" & Right(osnova, 1)
            End If
            Call WordForms(osnova, Array("а", "ові", "а", "ом", "ові", "е"))
            Call Rule(202)
            ManRule2 = True
        End If
    End If
End Function

' Функція відмінювання чоловічих імен, що закінчуються на приголосний або -о
Private Function ManRule3(ByVal workingWord As String) As Boolean
    Dim BeforeLast As String, osnova As String, osLast As String
    Dim group As Integer

    BeforeLast = Mid(workingWord, Len(workingWord) - 1, 1)

    If InStr(1, consonant & "оь", Right(workingWord, 1), vbTextCompare) > 0 Then
        group = Detect2Group(workingWord)
        osnova = GetOsnova(workingWord)
        osLast = Right(osnova, 1)

        ' Чергування і -> о
        If osLast <> "й" And Mid(osnova, Len(osnova) - 1, 1) = "і" And Not InStr(1, "світцвіт", LCase(osnova)) And workingWord <> "Гліб" Then
            osnova = Left(osnova, Len(osnova) - 2) & "о" & Right(osnova, 1)
        End If

        ' Випадання букви е
        If Left(osnova, 1) = "о" And FirstLastVowel(osnova, vowels & "гк") = "е" And Right(workingWord, 2) <> "сь" Then
            Dim delim As Integer
            delim = InStrRev(osnova, "е")
            osnova = Left(osnova, delim - 1) & Mid(osnova, delim + 1)
        End If

        ' Групи відмінювання
        Select Case group
            Case 1 ' Тверда група
                If Right(workingWord, 2) = "ок" And Right(workingWord, 3) <> "оок" Then
                    Call WordForms(workingWord, Array("ка", "кові", "ка", "ком", "кові", "че"), 2)
                    Call Rule(301)
                    ManRule3 = True
                Else
                    Call WordForms(osnova, Array(osLast & "а", osLast & "ові", osLast & "а", osLast & "ом", osLast & "ові", Inverse2(osLast) & "е"), 1)
                    Call Rule(304)
                    ManRule3 = True
                End If
            Case 2 ' Мішана група
                Call WordForms(osnova, Array("а", "еві", "а", "ем", "еві", "е"))
                Call Rule(305)
                ManRule3 = True
            Case 3 ' М’яка група
                If Right(workingWord, 2) = "ей" And InStr(1, gubni, Mid(workingWord, Len(workingWord) - 2, 1), vbTextCompare) > 0 Then
                    osnova = Left(workingWord, Len(workingWord) - 2) & "’"
                    Call WordForms(osnova, Array("я", "єві", "я", "єм", "єві", "ю"))
                    Call Rule(306)
                    ManRule3 = True
                ElseIf Right(workingWord, 1) = "й" Or BeforeLast = "і" Then
                    Call WordForms(workingWord, Array("я", "єві", "я", "єм", "єві", "ю"), 1)
                    Call Rule(307)
                    ManRule3 = True
                End If
        End Select
    End If
End Function


' Функція: Якщо слово закінчується на "і", відмінюємо як множину
Private Function ManRule4(ByVal workingWord As String) As Boolean
    If Right(workingWord, 1) = "і" Then
        Call WordForms(workingWord, Array("их", "им", "их", "ими", "их", "і"), 1)
        Call Rule(4)
        ManRule4 = True
    Else
        ManRule4 = False
    End If
End Function

' Функція: Якщо слово закінчується на "ий" або "ой"
Private Function ManRule5(ByVal workingWord As String) As Boolean
    If Right(workingWord, 2) = "ий" Or Right(workingWord, 2) = "ой" Then
        Call WordForms(workingWord, Array("ого", "ому", "ого", "им", "ому", "ий"), 2)
        Call Rule(5)
        ManRule5 = True
    Else
        ManRule5 = False
    End If
End Function

' Функція: Відмінювання жіночих імен, що закінчуються на -а або -я
Private Function WomanRule1(ByVal workingWord As String) As Boolean
    Dim BeforeLast As String
    BeforeLast = Mid(workingWord, Len(workingWord) - 1, 1)

    ' Якщо закінчується на "ніга", змінюємо на "нога"
    If Right(workingWord, 4) = "ніга" Then
        Dim osnova As String
        osnova = Left(workingWord, Len(workingWord) - 3) & "о"
        Call WordForms(osnova, Array("ги", "зі", "гу", "гою", "зі", "го"))
        Call Rule(101)
        WomanRule1 = True
        Exit Function
    End If

    ' Якщо закінчується на "а"
    If Right(workingWord, 1) = "а" Then
        Call WordForms(workingWord, Array(BeforeLast & "и", InverseGKH(BeforeLast) & "і", BeforeLast & "у", BeforeLast & "ою", InverseGKH(BeforeLast) & "і", BeforeLast & "о"), 2)
        Call Rule(102)
        WomanRule1 = True
        Exit Function
    End If

    ' Якщо закінчується на "я"
    If Right(workingWord, 1) = "я" Then
        If InStr(1, vowels & "’", BeforeLast, vbTextCompare) > 0 Then
            Call WordForms(workingWord, Array("ї", "ї", "ю", "єю", "ї", "є"), 1)
            Call Rule(103)
        Else
            Call WordForms(workingWord, Array(BeforeLast & "і", InverseGKH(BeforeLast) & "і", BeforeLast & "ю", BeforeLast & "ею", InverseGKH(BeforeLast) & "і", BeforeLast & "е"), 2)
            Call Rule(104)
        End If
        WomanRule1 = True
        Exit Function
    End If

    WomanRule1 = False
End Function

' Функція: Відмінювання жіночих імен, що закінчуються на приголосний або "ь"
Private Function WomanRule2(ByVal workingWord As String) As Boolean
    If InStr(1, consonant & "ь", Right(workingWord, 1), vbTextCompare) > 0 Then
        Dim osnova As String, apostrof As String, duplicate As String
        Dim osLast As String, osBeforeLast As String
        osnova = GetOsnova(workingWord)
        osLast = Right(osnova, 1)
        osBeforeLast = Mid(osnova, Len(osnova) - 1, 1)
        apostrof = ""
        duplicate = ""

        ' Чи треба ставити апостроф
        If InStr(1, gubni, osLast, vbTextCompare) > 0 And InStr(1, vowels, osBeforeLast, vbTextCompare) > 0 Then
            apostrof = "’"
        End If

        ' Чи треба подвоювати
        If InStr(1, "дтзсцлн", osLast, vbTextCompare) > 0 Then
            duplicate = osLast
        End If

        ' Відмінюємо
        If Right(workingWord, 1) = "ь" Then
            Call WordForms(osnova, Array("і", "і", "ь", duplicate & apostrof & "ю", "і", "е"))
            Call Rule(201)
        Else
            Call WordForms(osnova, Array("і", "і", "", duplicate & apostrof & "ю", "і", "е"))
            Call Rule(202)
        End If
        WomanRule2 = True
        Exit Function
    End If

    WomanRule2 = False
End Function

' Функція: Відмінювання жіночих імен, що закінчуються на -ая або -ськ
Private Function WomanRule3(ByVal workingWord As String) As Boolean
    Dim BeforeLast As String
    BeforeLast = Mid(workingWord, Len(workingWord) - 1, 1)

    ' Прізвища, що закінчуються на -ая (напр., Донская)
    If Right(workingWord, 2) = "ая" Then
        Call WordForms(workingWord, Array("ої", "ій", "ую", "ою", "ій", "ая"), 2)
        Call Rule(301)
        WomanRule3 = True
        Exit Function
    End If

    ' Прізвища на -ськ (або з кінцевими буквами "чнв" перед -а)
    If Right(workingWord, 1) = "а" And _
       (InStr(1, "чнв", BeforeLast, vbTextCompare) > 0 Or Mid(workingWord, Len(workingWord) - 2, 2) = "ьк") Then
        Call WordForms(workingWord, Array(BeforeLast & "ої", BeforeLast & "ій", BeforeLast & "у", BeforeLast & "ою", BeforeLast & "ій", BeforeLast & "о"), 2)
        Call Rule(302)
        WomanRule3 = True
        Exit Function
    End If

    WomanRule3 = False
End Function

' Функція: Застосовує ланцюг правил для чоловічих імен
Private Function ManFirstName(ByVal workingWord As String) As Boolean
    ManFirstName = RulesChain("man", Array(1, 2, 3), workingWord)
End Function

' Функція: Застосовує ланцюг правил для жіночих імен
Private Function WomanFirstName(ByVal workingWord As String) As Boolean
    WomanFirstName = RulesChain("woman", Array(1, 2), workingWord)
End Function

' Функція: Застосовує ланцюг правил для чоловічих прізвищ
Private Function ManSecondName(ByVal workingWord As String) As Boolean
    ManSecondName = RulesChain("man", Array(5, 1, 2, 3, 4), workingWord)
End Function

' Функція: Застосовує ланцюг правил для жіночих прізвищ
Private Function WomanSecondName(ByVal workingWord As String) As Boolean
    WomanSecondName = RulesChain("woman", Array(3, 1), workingWord)
End Function

' Функція: Застосовує правила з ланцюга
Private Function RulesChain(ByVal gender As String, ByVal rules() As Integer, ByVal workingWord As String) As Boolean
    Dim ruleNumber As Integer
    Dim result As Boolean
    Dim i As Integer

    result = False

    For i = LBound(rules) To UBound(rules)
        ruleNumber = rules(i)
        Select Case gender
            Case "man"
                Select Case ruleNumber
                    Case 1: result = ManRule1(workingWord)
                    Case 2: result = ManRule2(workingWord)
                    Case 3: result = ManRule3(workingWord)
                    Case 4: result = ManRule4(workingWord)
                    Case 5: result = ManRule5(workingWord)
                End Select
            Case "woman"
                Select Case ruleNumber
                    Case 1: result = WomanRule1(workingWord)
                    Case 2: result = WomanRule2(workingWord)
                    Case 3: result = WomanRule3(workingWord)
                End Select
        End Select

        If result Then
            RulesChain = True
            Exit Function
        End If
    Next i

    RulesChain = False
End Function


' Функція: Відмінювання чоловічих по-батькові
Private Function ManFatherName(ByVal workingWord As String) As Boolean
    If Right(workingWord, 2) = "ич" Or Right(workingWord, 2) = "іч" Then
        Call WordForms(workingWord, Array("а", "у", "а", "ем", "у", "у"))
        ManFatherName = True
    Else
        ManFatherName = False
    End If
End Function

' Функція: Відмінювання жіночих по-батькові
Private Function WomanFatherName(ByVal workingWord As String) As Boolean
    If Right(workingWord, 3) = "вна" Then
        Call WordForms(workingWord, Array("и", "і", "у", "ою", "і", "о"), 1)
        WomanFatherName = True
    Else
        WomanFatherName = False
    End If
End Function

' Функція: Визначення статі за ім’ям
Private Sub GenderByFirstName(ByRef genderMan As Double, ByRef genderWoman As Double, ByVal workingWord As String)
    ' Якщо ім’я закінчується на "й", то скоріше за все чоловік
    If Right(workingWord, 1) = "й" Then
        genderMan = genderMan + 0.9
    End If

    ' Чоловічі імена за списком
    If InStr(1, "Петро,Микола", workingWord, vbTextCompare) > 0 Then
        genderMan = genderMan + 30
    End If

    ' Чоловічі закінчення
    If InStr(1, "он,ов,ав,ам,ол,ан,рд,мп,ко,ло", Right(workingWord, 2), vbTextCompare) > 0 Then
        genderMan = genderMan + 0.5
    End If

    ' Жіночі закінчення
    If InStr(1, "бов,нка,яра,ила,опа", Right(workingWord, 3), vbTextCompare) > 0 Then
        genderWoman = genderWoman + 0.5
    End If

    ' Закінчення на приголосний
    If InStr(1, consonant, Right(workingWord, 1), vbTextCompare) > 0 Then
        genderMan = genderMan + 0.01
    End If

    ' Закінчення на "ь"
    If Right(workingWord, 1) = "ь" Then
        genderMan = genderMan + 0.02
    End If

    ' Жіночі закінчення на "дь"
    If Right(workingWord, 2) = "дь" Then
        genderWoman = genderWoman + 0.1
    End If

    ' Закінчення "ель" або "бов"
    If InStr(1, "ель,бов", Right(workingWord, 3), vbTextCompare) > 0 Then
        genderWoman = genderWoman + 0.4
    End If
End Sub

' Функція: Визначення статі за прізвищем
Private Sub GenderBySecondName(ByRef genderMan As Double, ByRef genderWoman As Double, ByVal workingWord As String)
    ' Чоловічі закінчення
    If InStr(1, "ов,ин,ев,єв,ін,їн,ий,їв,ів,ой,ей", Right(workingWord, 2), vbTextCompare) > 0 Then
        genderMan = genderMan + 0.4
    End If

    ' Жіночі закінчення
    If InStr(1, "ова,ина,ева,єва,іна,мін", Right(workingWord, 3), vbTextCompare) > 0 Then
        genderWoman = genderWoman + 0.4
    End If

    If Right(workingWord, 2) = "ая" Then
        genderWoman = genderWoman + 0.4
    End If
End Sub

' Функція: Визначення статі за по-батькові
Private Sub GenderByFatherName(ByRef genderMan As Double, ByRef genderWoman As Double, ByVal workingWord As String)
    If Right(workingWord, 2) = "ич" Then
        genderMan = genderMan + 10
    End If
    If Right(workingWord, 2) = "на" Then
        genderWoman = genderWoman + 12
    End If
End Sub

' Функція: Ідентифікує слово як ім’я, прізвище або по-батькові
Private Sub DetectNamePart(ByVal workingWord As String, ByRef namePart As String)
    Dim first As Double, second As Double, father As Double
    Dim maxVal As Double

    first = 0
    second = 0
    father = 0

    ' Якщо схоже на по-батькові
    If InStr(1, "вна,чна,ліч", Right(workingWord, 3), vbTextCompare) > 0 Or _
       InStr(1, "ьмич,ович", Right(workingWord, 4), vbTextCompare) > 0 Then
        father = father + 3
    End If

    ' Якщо схоже на ім’я
    If InStr(1, "тин,ьмич,юбов,івна,явка,орив,кіян", Right(workingWord, 3), vbTextCompare) > 0 Or _
       InStr(1, "ьмич,юбов,івна,явка,орив,кіян", Right(workingWord, 4), vbTextCompare) > 0 Then
        first = first + 0.5
    End If

    ' Виключення для імен
    If InStr(1, "Лев,Гаїна,Афіна,Антоніна,Ангеліна,Альвіна,Альбіна,Аліна,Павло,Олесь,Микола,Мая,Англеліна,Елькін,Мерлін", workingWord, vbTextCompare) > 0 Then
        first = first + 10
    End If

    ' Якщо схоже на прізвище
    If InStr(1, "ов,ін,ев,єв,ий,ин,ой,ко,ук,як,ца,их,ик,ун,ок,ша,ая,га,єк,аш,ив,юк,ус,це,ак,бр,яр,іл,ів,ич,сь,ей,нс,яс,ер,ай,ян,ах,ць,ющ,іс,ач,уб,ох,юх,ут,ча,ул,вк,зь,уц,їн,де,уз,юр,ік,іч,ро", Right(workingWord, 2), vbTextCompare) > 0 Then
        second = second + 0.4
    End If

    If InStr(1, "ова,ева,єва,тих,рик,вач,аха,шен,мей,арь,вка,шир,бан,чий,іна,їна,ька,ань,ива,аль,ура,ран,ало,ола,кур,оба,оль,нта,зій,ґан,іло,шта,юпа,рна,бла,еїн,има,мар,кар,оха,чур,ниш,ета,тна,зур,нір,йма,орж,рба,іла,лас,дід,роз,аба,чан,ган", Right(workingWord, 3), vbTextCompare) > 0 Then
        second = second + 0.4
    End If

    If InStr(1, "ьник,нчук,тник,кирь,ский,шена,шина,вина,нина,гана,хній,зюба,орош,орон,сило,руба,лест,мара,обка,рока,сика,одна,нчар,вата,ндар,грій", Right(workingWord, 4), vbTextCompare) > 0 Then
        second = second + 0.4
    End If

    ' Якщо закінчується на "і"
    If Right(workingWord, 1) = "і" Then
        second = second + 0.2
    End If

    ' Визначення частини імені
    maxVal = Application.WorksheetFunction.Max(first, second, father)

    If first = maxVal Then
        namePart = "N" ' Ім’я
    ElseIf second = maxVal Then
        namePart = "S" ' Прізвище
    Else
        namePart = "F" ' По-батькові
    End If
End Sub

' Функції для кожного відмінка
Public Function Називний_vidminok() As String
    Називний_vidminok = ProcessDeclension("називний")
End Function

Public Function Родовий_vidminok() As String
    Родовий_vidminok = ProcessDeclension("родовий")
End Function

Public Function Давальний_vidminok() As String
    Давальний_vidminok = ProcessDeclension("давальний")
End Function

Public Function Знахідний_vidminok() As String
    Знахідний_vidminok = ProcessDeclension("знахідний")
End Function

Public Function Орудний_vidminok() As String
    Орудний_vidminok = ProcessDeclension("орудний")
End Function

Public Function Місцевий_vidminok() As String
    Місцевий_vidminok = ProcessDeclension("місцевий")
End Function

Public Function Кличний_vidminok() As String
    Кличний_vidminok = ProcessDeclension("кличний")
End Function

' Основна функція для обробки вибраного відмінка
Private Function ProcessDeclension(ByVal caseName As String) As String
    Dim rng As Range
    Dim inputText As String

    ' Запит на вибір комірки
    On Error Resume Next
    Set rng = Application.InputBox("Виберіть комірку з текстом для відмінювання:", Type:=8)
    On Error GoTo 0

    If rng Is Nothing Then
        ProcessDeclension = "Скасовано користувачем."
        Exit Function
    End If

    inputText = rng.Value

    ' Відмінювання тексту (шаблонна логіка)
    ProcessDeclension = DeclineName(inputText, caseName)
End Function

' Шаблон функції для відмінювання (доповнити реальною логікою)
Private Function DeclineName(ByVal fullName As String, ByVal caseName As String) As String
    ' Тимчасова логіка для демонстрації
    DeclineName = fullName & " (" & caseName & ")"
End Function


' --- Адаптована логіка з файлу NCLNameCaseCore.php ---
' Оригінальний PHP-код був розроблений як об'єктно-орієнтований. Нижче представлено його адаптацію у VBA.

' Тут ви можете додати функції з PHP у вигляді підпроцедур або функцій VBA.
' Для прикладу (це слід замінити на реальні конвертовані функції):
'
' Public Function DeclineName(name As String, caseNumber As Integer) As String
'    ' Логіка з PHP, перетворена на VBA
'    DeclineName = name & "_case" & caseNumber
' End Function

' --- Вставте адаптацію функцій з NCLNameCaseCore.php ---

<?php

/**
 * @license Dual licensed under the MIT or GPL Version 2 licenses.
 * @package NameCaseLib
 */
/**
 *
 */
if (!defined('NCL_DIR'))
{
		define('NCL_DIR', dirname(__FILE__));
}

require_once NCL_DIR . '/NCL.php';
require_once NCL_DIR . '/NCLStr.php';
require_once NCL_DIR . '/NCLNameCaseWord.php';

/**
 * <b>NCL NameCase Core</b>
 *
 * Набор основных функций, который позволяют сделать интерфейс слонения русского и украниского языка
 * абсолютно одинаковым. Содержит все функции для внешнего взаимодействия с библиотекой.
 *
 * @author Андрей Чайка <bymer3@gmail.com>
 * @version 0.4.1
 * @package NameCaseLib
 */
class NCLNameCaseCore extends NCL
{

		/**
		 * Версия библиотеки
		 * @var string
		 */
		protected $version = '0.4.1';
		/**
		 * Версия языкового файла
		 * @var string
		 */
		protected $languageBuild = '0';
		/**
		 * Готовность системы:
		 * - Все слова идентифицированы (известо к какой части ФИО относится слово)
		 * - У всех слов определен пол
		 * Если все сделано стоит флаг true, при добавлении нового слова флаг сбрасывается на false
		 * @var bool
		 */
		private $ready = false;
		/**
		 * Если все текущие слова было просклонены и в каждом слове уже есть результат склонения,
		 * тогда true. Если было добавлено новое слово флаг збрасывается на false
		 * @var bool
		 */
		private $finished = false;
		/**
		 * Массив содержит елементы типа NCLNameCaseWord. Это все слова которые нужно обработать и просклонять
		 * @var array
		 */
		private $words = array();
		/**
		 * Переменная, в которую заносится слово с которым сейчас идет работа
		 * @var string
		 */
		protected $workingWord = '';
		/**
		 * Метод Last() вырезает подстроки разной длины. Посколько одинаковых вызовов бывает несколько,
		 * то все результаты выполнения кешируются в этом массиве.
		 * @var array
		 */
		protected $workindLastCache = array();
		/**
		 * Номер последнего использованого правила, устанавливается методом Rule()
		 * @var int
		 */
		private $lastRule = 0;
		/**
		 * Массив содержит результат склонения слова - слово во всех падежах
		 * @var array
		 */
		protected $lastResult = array();
		/**
		 * Массив содержит информацию о том какие слова из массива <var>$this->words</var> относятся к
		 * фамилии, какие к отчеству а какие к имени. Массив нужен потому, что при добавлении слов мы не
		 * всегда знаем какая часть ФИО сейчас, поэтому после идентификации всех слов генерируется массив
		 * индексов для быстрого поиска в дальнейшем.
		 * @var array
		 */
		private $index = array();

		public $gender_koef=0;//вероятность автоопредления пола [0..10]. Достаточно точно при 0.1

		/**
		 * Метод очищает результаты последнего склонения слова. Нужен при склонении нескольких слов.
		 */
		private function reset()
		{
				$this->lastRule = 0;
				$this->lastResult = array();
		}

		/**
		 * Сбрасывает все информацию на начальную. Очищает все слова добавленые в систему.
		 * После выполнения система готова работать с начала.
		 * @return NCLNameCaseCore
		 */
		public function fullReset()
		{
				$this->words = array();
				$this->index = array('N' => array(), 'F' => array(), 'S' => array());
				$this->reset();
				$this->notReady();
				return $this;
		}

		/**
		 * Устанавливает флаги о том, что система не готово и слова еще не были просклонены
		 */
		private function notReady()
		{
				$this->ready = false;
				$this->finished = false;
		}

		/**
		 * Устанавливает номер последнего правила
		 * @param int $index номер правила которое нужно установить
		 */
		protected function Rule($index)
		{
				$this->lastRule = $index;
		}

		/**
		 * Устанавливает слово текущим для работы системы. Очищает кеш слова.
		 * @param string $word слово, которое нужно установить
		 */
		protected function setWorkingWord($word)
		{
				//Сбрасываем настройки
				$this->reset();
				//Ставим слово
				$this->workingWord = $word;
				//Чистим кеш
				$this->workindLastCache = array();
		}

		/**
		 * Если не нужно склонять слово, делает результат таким же как и именительный падеж
		 */
		protected function makeResultTheSame()
		{
				$this->lastResult = array_fill(0, $this->CaseCount, $this->workingWord);
		}

		/**
		 * Если <var>$stopAfter</var> = 0, тогда вырезает $length последних букв с текущего слова (<var>$this->workingWord</var>)
		 * Если нет, тогда вырезает <var>$stopAfter</var> букв начиная от <var>$length</var> с конца
		 * @param int $length количество букв с конца
		 * @param int $stopAfter количество букв которые нужно вырезать (0 - все)
		 * @return string требуемая подстрока
		 */
		protected function Last($length=1, $stopAfter=0)
		{
				//Сколько букв нужно вырезать все или только часть
				if (!$stopAfter)
				{
						$cut = $length;
				}
				else
				{
						$cut = $stopAfter;
				}

				//Проверяем кеш
				if (!isset($this->workindLastCache[$length][$stopAfter]))
				{
						$this->workindLastCache[$length][$stopAfter] = NCLStr::substr($this->workingWord, -$length, $cut);
				}
				return $this->workindLastCache[$length][$stopAfter];
		}

		/**
		 * Над текущим словом (<var>$this->workingWord</var>) выполняются правила в порядке указаном в <var>$rulesArray</var>.
		 * <var>$gender</var> служит для указания какие правила использовать мужские ('man') или женские ('woman')
		 * @param string $gender - префикс мужских/женских правил
		 * @param array $rulesArray - массив, порядок выполнения правил
		 * @return boolean если правило было задествовано, тогда true, если нет - тогда false
		 */
		protected function RulesChain($gender, $rulesArray)
		{
				foreach ($rulesArray as $ruleID)
				{
						$ruleMethod = $gender . 'Rule' . $ruleID;
						if ($this->$ruleMethod())
						{
								return true;
						}
				}
				return false;
		}

		/**
		 * Если <var>$string</var> строка, тогда проверяется входит ли буква <var>$letter</var> в строку <var>$string</var>
		 * Если <var>$string</var> массив, тогда проверяется входит ли строка <var>$letter</var> в массив <var>$string</var>
		 * @param string $letter буква или строка, которую нужно искать
		 * @param mixed $string строка или массив, в котором нужно искать
		 * @return bool true если искомое значение найдено
		 */
		protected function in($letter, $string)
		{
				//Если второй параметр массив
				if (is_array($string))
				{
						return in_array($letter, $string);
				}
				else
				{
						if (!$letter or NCLStr::strpos($string, $letter) === false)
						{
								return false;
						}
						else
						{
								return true;
						}
				}
		}

		/**
		 * Функция проверяет, входит ли имя <var>$nameNeedle</var> в перечень имен <var>$names</var>.
		 * @param string $nameNeedle - имя которое нужно найти
		 * @param array $names - перечень имен в котором нужно найти имя
		 */
		protected function inNames($nameNeedle, $names)
		{
				if (!is_array($names))
				{
						$names = array($names);
				}

				foreach ($names as $name)
				{
						if (NCLStr::strtolower($nameNeedle) == NCLStr::strtolower($name))
						{
								return true;
						}
				}
				return false;
		}

		/**
		 * Склоняет слово <var>$word</var>, удаляя из него <var>$replaceLast</var> последних букв
		 * и добавляя в каждый падеж окончание из массива <var>$endings</var>.
		 * @param string $word слово, к которому нужно добавить окончания
		 * @param array $endings массив окончаний
		 * @param int $replaceLast сколько последних букв нужно убрать с начального слова
		 */
		protected function wordForms($word, $endings, $replaceLast=0)
		{
				//Создаем массив с именительный падежом
				$result = array($this->workingWord);
				//Убираем в окончание лишние буквы
				$word = NCLStr::substr($word, 0, NCLStr::strlen($word) - $replaceLast);

				//Добавляем окончания
				for ($padegIndex = 1; $padegIndex < $this->CaseCount; $padegIndex++)
				{
						$result[$padegIndex] = $word . $endings[$padegIndex - 1];
				}

				$this->lastResult = $result;
		}

		/**
		 * В массив <var>$this->words</var> добавляется новый об’єкт класса NCLNameCaseWord
		 * со словом <var>$firstname</var> и пометкой, что это имя
		 * @param string $firstname имя
		 * @return NCLNameCaseCore
		 */
		public function setFirstName($firstname="")
		{
				if ($firstname)
				{
						$index = count($this->words);
						$this->words[$index] = new NCLNameCaseWord($firstname);
						$this->words[$index]->setNamePart('N');
						$this->notReady();
				}
				return $this;
		}

		/**
		 * В массив <var>$this->words</var> добавляется новый об’єкт класса NCLNameCaseWord
		 * со словом <var>$secondname</var> и пометкой, что это фамилия
		 * @param string $secondname фамилия
		 * @return NCLNameCaseCore
		 */
		public function setSecondName($secondname="")
		{
				if ($secondname)
				{
						$index = count($this->words);
						$this->words[$index] = new NCLNameCaseWord($secondname);
						$this->words[$index]->setNamePart('S');
						$this->notReady();
				}
				return $this;
		}

		/**
		 * В массив <var>$this->words</var> добавляется новый об’єкт класса NCLNameCaseWord
		 * со словом <var>$fathername</var> и пометкой, что это отчество
		 * @param string $fathername отчество
		 * @return NCLNameCaseCore
		 */
		public function setFatherName($fathername="")
		{
				if ($fathername)
				{
						$index = count($this->words);
						$this->words[$index] = new NCLNameCaseWord($fathername);
						$this->words[$index]->setNamePart('F');
						$this->notReady();
				}
				return $this;
		}

		/**
		 * Всем словам устанавливается пол, который может иметь следующие значения
		 * - 0 - не определено
		 * - NCL::$MAN - мужчина
		 * - NCL::$WOMAN - женщина
		 * @param int $gender пол, который нужно установить
		 * @return NCLNameCaseCore
		 */
		public function setGender($gender=0)
		{
				foreach ($this->words as $word)
				{
						$word->setTrueGender($gender);
				}
				return $this;
		}

		/**
		 * В система заносится сразу фамилия, имя, отчество
		 * @param string $secondName фамилия
		 * @param string $firstName имя
		 * @param string $fatherName отчество
		 * @return NCLNameCaseCore
		 */
		public function setFullName($secondName="", $firstName="", $fatherName="")
		{
				$this->setFirstName($firstName);
				$this->setSecondName($secondName);
				$this->setFatherName($fatherName);
				return $this;
		}

		/**
		 * В массив <var>$this->words</var> добавляется новый об’єкт класса NCLNameCaseWord
		 * со словом <var>$firstname</var> и пометкой, что это имя
		 * @param string $firstname имя
		 * @return NCLNameCaseCore
		 */
		public function setName($firstname="")
		{
				return $this->setFirstName($firstname);
		}

		/**
		 * В массив <var>$this->words</var> добавляется новый об’єкт класса NCLNameCaseWord
		 * со словом <var>$secondname</var> и пометкой, что это фамилия
		 * @param string $secondname фамилия
		 * @return NCLNameCaseCore
		 */
		public function setLastName($secondname="")
		{
				return $this->setSecondName($secondname);
		}

		/**
		 * В массив <var>$this->words</var> добавляется новый об’єкт класса NCLNameCaseWord
		 * со словом <var>$secondname</var> и пометкой, что это фамилия
		 * @param string $secondname фамилия
		 * @return NCLNameCaseCore
		 */
		public function setSirName($secondname="")
		{
				return $this->setSecondName($secondname);
		}

		/**
		 * Если слово <var>$word</var> не идентифицировано, тогда определяется это имя, фамилия или отчество
		 * @param NCLNameCaseWord $word слово которое нужно идентифицировать
		 */
		private function prepareNamePart(NCLNameCaseWord $word)
		{
				if (!$word->getNamePart())
				{
						$this->detectNamePart($word);
				}
		}

		/**
		 * Проверяет все ли слова идентифицированы, если нет тогда для каждого определяется это имя, фамилия или отчество
		 */
		private function prepareAllNameParts()
		{
				foreach ($this->words as $word)
				{
						$this->prepareNamePart($word);
				}
		}

		/**
		 * Определяет пол для слова <var>$word</var>
		 * @param NCLNameCaseWord $word слово для которого нужно определить пол
		 */
		private function prepareGender(NCLNameCaseWord $word)
		{
				if (!$word->isGenderSolved())
				{
						$namePart = $word->getNamePart();
						switch ($namePart)
						{
								case 'N': $this->GenderByFirstName($word);
										break;
								case 'F': $this->GenderByFatherName($word);
										break;
								case 'S': $this->GenderBySecondName($word);
										break;
						}
				}
		}

		/**
		 * Для всех слов проверяет определен ли пол, если нет - определяет его
		 * После этого расчитывает пол для всех слов и устанавливает такой пол всем словам
		 * @return bool был ли определен пол
		 */
		private function solveGender()
		{
				//Ищем, может гдето пол уже установлен
				foreach ($this->words as $word)
				{
						if ($word->isGenderSolved())
						{
								$this->setGender($word->gender());
								return true;
						}
				}

				//Если нет тогда определяем у каждого слова и потом сумируем
				$man = 0;
				$woman = 0;

				foreach ($this->words as $word)
				{
						$this->prepareGender($word);
						$gender = $word->getGender();
						$man+=$gender[NCL::$MAN];
						$woman+=$gender[NCL::$WOMAN];
				}

				if ($man > $woman)
				{
						$this->setGender(NCL::$MAN);
				}
				else
				{
						$this->setGender(NCL::$WOMAN);
				}

				return true;
		}

		/**
		 * Генерируется массив, который содержит информацию о том какие слова из массива <var>$this->words</var> относятся к
		 * фамилии, какие к отчеству а какие к имени. Массив нужен потому, что при добавлении слов мы не
		 * всегда знаем какая часть ФИО сейчас, поэтому после идентификации всех слов генерируется массив
		 * индексов для быстрого поиска в дальнейшем.
		 */
		private function generateIndex()
		{
				$this->index = array('N' => array(), 'S' => array(), 'F' => array());
				foreach ($this->words as $index => $word)
				{
						$namepart = $word->getNamePart();
						$this->index[$namepart][] = $index;
				}
		}

		/**
		 * Выполнет все необходимые подготовления для склонения.
		 * Все слова идентфицируются. Определяется пол.
		 * Обновляется индекс.
		 */
		private function prepareEverything()
		{
				if (!$this->ready)
				{
						$this->prepareAllNameParts();
						$this->solveGender();
						$this->generateIndex();
						$this->ready = true;
				}
		}

		/**
		 * По указаным словам определяется пол человека:
		 * - 0 - не определено
		 * - NCL::$MAN - мужчина
		 * - NCL::$WOMAN - женщина
		 * @return int текущий пол человека
		 */
		public function genderAutoDetect()
		{
				$this->prepareEverything();

				if (!empty($this->words)){
					$n=-1;
					$max_koef=-1;
					foreach ($this->words as $k=>$word){
						$genders=$word->getGender();
						$min=min( $genders );
						$max=max( $genders );
						$koef=$max-$min;
						if ($koef>$max_koef) {
							$max_koef=$koef;
							$n=$k;
						}
					}

					if ($n>=0){
						if (isset($this->words[$n]))
				{
							$genders=$this->words[$n]->getGender();
							$min=min( $genders );
							$max=max( $genders );
							$this->gender_koef=$max-$min;

							return $this->words[$n]->gender();
						}
					}
				}
				return false;
		}

		/**
		 * Разбивает строку <var>$fullname</var> на слова и возвращает формат в котором записано имя
		 * <b>Формат:</b>
		 * - S - Фамилия
		 * - N - Имя
		 * - F - Отчество
		 * @param string $fullname строка, для которой необходимо определить формат
		 * @return array формат в котором записано имя массив типа <var>$this->words</var>
		 */
		private function splitFullName($fullname)
		{

				$fullname = trim($fullname);
				$list = explode(' ', $fullname);

				foreach ($list as $word)
				{
						$this->words[] = new NCLNameCaseWord($word);
				}

				$this->prepareEverything();
				$formatArr = array();

				foreach ($this->words as $word)
				{
						$formatArr[] = $word->getNamePart();
				}

				return $this->words;
		}

		/**
		 * Разбивает строку <var>$fullname</var> на слова и возвращает формат в котором записано имя
		 * <b>Формат:</b>
		 * - S - Фамилия
		 * - N - Имя
		 * - F - Отчество
		 * @param string $fullname строка, для которой необходимо определить формат
		 * @return string формат в котором записано имя
		 */
		public function getFullNameFormat($fullname)
		{
				$this->fullReset();
				$words = $this->splitFullName($fullname);
				$format = '';
				foreach ($words as $word)
				{
						$format .= $word->getNamePart() . ' ';
				}
				return $format;
		}

		/**
		 * Склоняет слово <var>$word</var> по нужным правилам в зависимости от пола и типа слова
		 * @param NCLNameCaseWord $word слово, которое нужно просклонять
		 */
		private function WordCase(NCLNameCaseWord $word)
		{
				$gender = ($word->gender() == NCL::$MAN ? 'man' : 'woman');

				$namepart = '';

				$name_part_letter=$word->getNamePart();
				switch ($name_part_letter)
				{
						case 'F': $namepart = 'Father';
								break;
						case 'N': $namepart = 'First';
								break;
						case 'S': $namepart = 'Second';
								break;
				}

				$method = $gender . $namepart . 'Name';

				//если фамилия из 2х слов через дефис
				//http://new.gramota.ru/spravka/buro/search-answer?s=273912

				//рабоиваем слово с дефисами на части
				$tmp=$word->getWordOrig();
				$cur_words=explode('-', $tmp);
				$o_cur_words=array();

				$result=array();
				$last_rule=-1;

				$cnt=count($cur_words);
				foreach ($cur_words as $k=>$cur_word){
					$is_norm_rules=true;

					$o_ncw=new NCLNameCaseWord($cur_word);
					if ( $name_part_letter=='S' && $cnt>1 && $k<$cnt-1 ){
						//если первая часть фамилии тоже фамилия, то склоняем по общим правилам
						//иначе не склоняется

						$exclusion=array('тулуз');//исключения
						$cur_word_=mb_strtolower($cur_word);
						if ( !in_array($cur_word_, $exclusion ) ){
							$o_nc = new NCLNameCaseRu();
							$o_nc->detectNamePart( $o_ncw );
							$is_norm_rules=( $o_ncw->getNamePart()=='S' );
						}
						else {
							$is_norm_rules=false;
						}
					}

					$this->setWorkingWord($cur_word);

					if ($is_norm_rules && $this->$method())
					{
						//склоняется
						$result_tmp=$this->lastResult;
						$last_rule=$this->lastRule;
					}
					else
					{
						//не склоняется. Заполняем что есть
						$result_tmp=array_fill(0, $this->CaseCount, $cur_word);
						$last_rule=-1;
					}

					$o_ncw->setNameCases($result_tmp);
					$o_cur_words[]=$o_ncw;
				}

				//объединение пачку частей слова в одно слово по каждому падежу
				foreach ($o_cur_words as $o_ncw){
					$namecases=$o_ncw->getNameCases();
					foreach ($namecases as $k=>$namecase){
						if ( key_exists($k, $result) ) $result[$k]=$result[$k].'-'.$namecase;
						else $result[$k]=$namecase;
					}
				}

				//устанавливаем падежи для целого слова
				$word->setNameCases($result, false);
				$word->setRule($last_rule);
		}

		/**
		 * Производит склонение всех слов, который хранятся в массиве <var>$this->words</var>
		 */
		private function AllWordCases()
		{
				if (!$this->finished)
				{
						$this->prepareEverything();

						foreach ($this->words as $word)
						{
								$this->WordCase($word);
						}

						$this->finished = true;
				}
		}

		/**
		 * Если указан номер падежа <var>$number</var>, тогда возвращается строка с таким номером падежа,
		 * если нет, тогда возвращается массив со всеми падежами текущего слова.
		 * @param NCLNameCaseWord $word слово для котрого нужно вернуть падеж
		 * @param int $number номер падежа, который нужно вернуть
		 * @return mixed массив или строка с нужным падежом
		 */
		private function getWordCase(NCLNameCaseWord $word, $number=null)
		{
				$cases = $word->getNameCases();
				if (is_null($number) or $number < 0 or $number > ($this->CaseCount - 1))
				{
						return $cases;
				}
				else
				{
						return $cases[$number];
				}
		}

		/**
		 * Если нужно было просклонять несколько слов, то их необходимо собрать в одну строку.
		 * Эта функция собирает все слова указаные в <var>$indexArray</var>  в одну строку.
		 * @param array $indexArray индексы слов, которые необходимо собрать вместе
		 * @param int $number номер падежа
		 * @return mixed либо массив со всеми падежами, либо строка с одним падежом
		 */
		private function getCasesConnected($indexArray, $number=null)
		{
				$readyArr = array();
				foreach ($indexArray as $index)
				{
						$readyArr[] = $this->getWordCase($this->words[$index], $number);
				}

				$all = count($readyArr);
				if ($all)
				{
						if (is_array($readyArr[0]))
						{
								//Масив нужно скелить каждый падеж
								$resultArr = array();
								for ($case = 0; $case < $this->CaseCount; $case++)
								{
										$tmp = array();
										for ($i = 0; $i < $all; $i++)
										{
												$tmp[] = $readyArr[$i][$case];
										}
										$resultArr[$case] = implode(' ', $tmp);
								}
								return $resultArr;
						}
						else
						{
								return implode(' ', $readyArr);
						}
				}
				return '';
		}

		/**
		 * Функция ставит имя в нужный падеж.
		 *
		 * Если указан номер падежа <var>$number</var>, тогда возвращается строка с таким номером падежа,
		 * если нет, тогда возвращается массив со всеми падежами текущего слова.
		 * @param int $number номер падежа
		 * @return mixed массив или строка с нужным падежом
		 */
		public function getFirstNameCase($number=null)
		{
				$this->AllWordCases();

				return $this->getCasesConnected($this->index['N'], $number);
		}

		/**
		 * Функция ставит фамилию в нужный падеж.
		 *
		 * Если указан номер падежа <var>$number</var>, тогда возвращается строка с таким номером падежа,
		 * если нет, тогда возвращается массив со всеми падежами текущего слова.
		 * @param int $number номер падежа
		 * @return mixed массив или строка с нужным падежом
		 */
		public function getSecondNameCase($number=null)
		{
				$this->AllWordCases();

				return $this->getCasesConnected($this->index['S'], $number);
		}

		/**
		 * Функция ставит отчество в нужный падеж.
		 *
		 * Если указан номер падежа <var>$number</var>, тогда возвращается строка с таким номером падежа,
		 * если нет, тогда возвращается массив со всеми падежами текущего слова.
		 * @param int $number номер падежа
		 * @return mixed массив или строка с нужным падежом
		 */
		public function getFatherNameCase($number=null)
		{
				$this->AllWordCases();

				return $this->getCasesConnected($this->index['F'], $number);
		}

		/**
		 * Функция ставит имя <var>$firstName</var> в нужный падеж <var>$CaseNumber</var> по правилам пола <var>$gender</var>.
		 *
		 * Если указан номер падежа <var>$CaseNumber</var>, тогда возвращается строка с таким номером падежа,
		 * если нет, тогда возвращается массив со всеми падежами текущего слова.
		 * @param string $firstName имя, которое нужно просклонять
		 * @param int $CaseNumber номер падежа
		 * @param int $gender пол, который нужно использовать
		 * @return mixed массив или строка с нужным падежом
		 */
		public function qFirstName($firstName, $CaseNumber=null, $gender=0)
		{
				$this->fullReset();
				$this->setFirstName($firstName);
				if ($gender)
				{
						$this->setGender($gender);
				}
				return $this->getFirstNameCase($CaseNumber);
		}

		/**
		 * Функция ставит фамилию <var>$secondName</var> в нужный падеж <var>$CaseNumber</var> по правилам пола <var>$gender</var>.
		 *
		 * Если указан номер падежа <var>$CaseNumber</var>, тогда возвращается строка с таким номером падежа,
		 * если нет, тогда возвращается массив со всеми падежами текущего слова.
		 * @param string $secondName фамилия, которую нужно просклонять
		 * @param int $CaseNumber номер падежа
		 * @param int $gender пол, который нужно использовать
		 * @return mixed массив или строка с нужным падежом
		 */
		public function qSecondName($secondName, $CaseNumber=null, $gender=0)
		{
				$this->fullReset();
				$this->setSecondName($secondName);
				if ($gender)
				{
						$this->setGender($gender);
				}

				return $this->getSecondNameCase($CaseNumber);
		}

		/**
		 * Функция ставит отчество <var>$fatherName</var> в нужный падеж <var>$CaseNumber</var> по правилам пола <var>$gender</var>.
		 *
		 * Если указан номер падежа <var>$CaseNumber</var>, тогда возвращается строка с таким номером падежа,
		 * если нет, тогда возвращается массив со всеми падежами текущего слова.
		 * @param string $fatherName отчество, которое нужно просклонять
		 * @param int $CaseNumber номер падежа
		 * @param int $gender пол, который нужно использовать
		 * @return mixed массив или строка с нужным падежом
		 */
		public function qFatherName($fatherName, $CaseNumber=null, $gender=0)
		{
				$this->fullReset();
				$this->setFatherName($fatherName);
				if ($gender)
				{
						$this->setGender($gender);
				}
				return $this->getFatherNameCase($CaseNumber);
		}

		/**
		 * Склоняет текущие слова во все падежи и форматирует слово по шаблону <var>$format</var>
		 * <b>Формат:</b>
		 * - S - Фамилия
		 * - N - Имя
		 * - F - Отчество
		 * @param string $format строка формат
		 * @return array массив со всеми падежами
		 */
		public function getFormattedArray($format)
		{
				if (is_array($format))
				{
						return $this->getFormattedArrayHard($format);
				}

				$length = NCLStr::strlen($format);
				$result = array();
				$cases = array();
				$cases['S'] = $this->getCasesConnected($this->index['S']);
				$cases['N'] = $this->getCasesConnected($this->index['N']);
				$cases['F'] = $this->getCasesConnected($this->index['F']);

				for ($curCase = 0; $curCase < $this->CaseCount; $curCase++)
				{
						$line = "";
						for ($i = 0; $i < $length; $i++)
						{
								$symbol = NCLStr::substr($format, $i, 1);
								if ($symbol == 'S')
								{
										$line.=$cases['S'][$curCase];
								}
								elseif ($symbol == 'N')
								{
										$line.=$cases['N'][$curCase];
								}
								elseif ($symbol == 'F')
								{
										$line.=$cases['F'][$curCase];
								}
								else
								{
										$line.=$symbol;
								}
						}
						$result[] = $line;
				}
				return $result;
		}

		/**
		 * Склоняет текущие слова во все падежи и форматирует слово по шаблону <var>$format</var>
		 * <b>Формат:</b>
		 * - S - Фамилия
		 * - N - Имя
		 * - F - Отчество
		 * @param array $format массив с форматом
		 * @return array массив со всеми падежами
		 */
		public function getFormattedArrayHard($format)
		{

				$result = array();
				$cases = array();
				foreach ($format as $word)
				{
						$cases[] = $word->getNameCases();
				}

				for ($curCase = 0; $curCase < $this->CaseCount; $curCase++)
				{
						$line = "";
						foreach ($cases as $value)
						{
								$line.=$value[$curCase] . ' ';
						}
						$result[] = trim($line);
				}
				return $result;
		}

		/**
		 * Склоняет текущие слова в падеж <var>$caseNum</var> и форматирует слово по шаблону <var>$format</var>
		 * <b>Формат:</b>
		 * - S - Фамилия
		 * - N - Имя
		 * - F - Отчество
		 * @param array $format массив с форматом
		 * @return string строка в нужном падеже
		 */
		public function getFormattedHard($caseNum=0, $format=array())
		{
				$result = "";
				foreach ($format as $word)
				{
						$cases = $word->getNameCases();
						$result.= $cases[$caseNum] . ' ';
				}
				return trim($result);
		}

		/**
		 * Склоняет текущие слова в падеж <var>$caseNum</var> и форматирует слово по шаблону <var>$format</var>
		 * <b>Формат:</b>
		 * - S - Фамилия
		 * - N - Имя
		 * - F - Отчество
		 * @param string $format строка с форматом
		 * @return string строка в нужном падеже
		 */
		public function getFormatted($caseNum=0, $format="S N F")
		{
				$this->AllWordCases();
				//Если не указан падеж используем другую функцию
				if (is_null($caseNum) or !$caseNum)
				{
						return $this->getFormattedArray($format);
				}
				//Если формат сложный
				elseif (is_array($format))
				{
						return $this->getFormattedHard($caseNum, $format);
				}
				else
				{
						$length = NCLStr::strlen($format);
						$result = "";
						for ($i = 0; $i < $length; $i++)
						{
								$symbol = NCLStr::substr($format, $i, 1);
								if ($symbol == 'S')
								{
										$result.=$this->getSecondNameCase($caseNum);
								}
								elseif ($symbol == 'N')
								{
										$result.=$this->getFirstNameCase($caseNum);
								}
								elseif ($symbol == 'F')
								{
										$result.=$this->getFatherNameCase($caseNum);
								}
								else
								{
										$result.=$symbol;
								}
						}
						return $result;
				}
		}

		/**
		 * Склоняет фамилию <var>$secondName</var>, имя <var>$firstName</var>, отчество <var>$fatherName</var>
		 * в падеж <var>$caseNum</var> по правилам пола <var>$gender</var> и форматирует результат по шаблону <var>$format</var>
		 * <b>Формат:</b>
		 * - S - Фамилия
		 * - N - Имя
		 * - F - Отчество
		 * @param string $secondName фамилия
		 * @param string $firstName имя
		 * @param string $fatherName отчество
		 * @param int $gender пол
		 * @param int $caseNum номер падежа
		 * @param string $format формат
		 * @return mixed либо массив со всеми падежами, либо строка
		 */
		public function qFullName($secondName="", $firstName="", $fatherName="", $gender=0, $caseNum=0, $format="S N F")
		{
				$this->fullReset();
				$this->setFirstName($firstName);
				$this->setSecondName($secondName);
				$this->setFatherName($fatherName);
				if ($gender)
				{
						$this->setGender($gender);
				}

				return $this->getFormatted($caseNum, $format);
		}

		/**
		 * Склоняет ФИО <var>$fullname</var> в падеж <var>$caseNum</var> по правилам пола <var>$gender</var>.
		 * Возвращает результат в таком же формате, как он и был.
		 * @param string $fullname ФИО
		 * @param int $caseNum номер падежа
		 * @param int $gender пол человека
		 * @return mixed либо массив со всеми падежами, либо строка
		 */
		public function q($fullname, $caseNum=null, $gender=null)
		{
				$this->fullReset();
				$format = $this->splitFullName($fullname);
				if ($gender)
				{
						$this->setGender($gender);
				}

				return $this->getFormatted($caseNum, $format);
		}

		/**
		 * Определяет пол человека по ФИО
		 * @param string $fullname ФИО
		 * @return int пол человека
		 */
		public function genderDetect($fullname)
		{
				$this->fullReset();
				$this->splitFullName($fullname);
				return $this->genderAutoDetect();
		}

		/**
		 * Возвращает внутренний массив $this->words каждая запись имеет тип NCLNameCaseWord
		 * @return array Массив всех слов в системе
		 */
		public function getWordsArray()
		{
				return $this->words;
		}

		/**
		 * Функция пытается применить цепочку правил для мужских имен
		 * @return boolean true - если было использовано правило из списка, false - если правило не было найденым
		 */
		protected function manFirstName()
		{
				return false;
		}

		/**
		 * Функция пытается применить цепочку правил для женских имен
		 * @return boolean true - если было использовано правило из списка, false - если правило не было найденым
		 */
		protected function womanFirstName()
		{
				return false;
		}

		/**
		 * Функция пытается применить цепочку правил для мужских фамилий
		 * @return boolean true - если было использовано правило из списка, false - если правило не было найденым
		 */
		protected function manSecondName()
		{
				return false;
		}

		/**
		 * Функция пытается применить цепочку правил для женских фамилий
		 * @return boolean true - если было использовано правило из списка, false - если правило не было найденым
		 */
		protected function womanSecondName()
		{
				return false;
		}

		/**
		 * Функция склоняет мужский отчества
		 * @return boolean true - если слово было успешно изменено, false - если не получилось этого сделать
		 */
		protected function manFatherName()
		{
				return false;
		}

		/**
		 * Функция склоняет женские отчества
		 * @return boolean true - если слово было успешно изменено, false - если не получилось этого сделать
		 */
		protected function womanFatherName()
		{
				return false;
		}

		/**
		 * Определение пола по правилам имен
		 * @param NCLNameCaseWord $word обьект класса слов, для которого нужно определить пол
		 */
		protected function GenderByFirstName(NCLNameCaseWord $word)
		{

		}

		/**
		 * Определение пола по правилам фамилий
		 * @param NCLNameCaseWord $word обьект класса слов, для которого нужно определить пол
		 */
		protected function GenderBySecondName(NCLNameCaseWord $word)
		{

		}

		/**
		 * Определение пола по правилам отчеств
		 * @param NCLNameCaseWord $word обьект класса слов, для которого нужно определить пол
		 */
		protected function GenderByFatherName(NCLNameCaseWord $word)
		{

		}

		/**
		 * Идетифицирует слово определяе имя это, или фамилия, или отчество
		 * - <b>N</b> - имя
		 * - <b>S</b> - фамилия
		 * - <b>F</b> - отчество
		 * @param NCLNameCaseWord $word обьект класса слов, который необходимо идентифицировать
		 */
		protected function detectNamePart(NCLNameCaseWord $word)
		{

		}

		/**
		 * Возвращает версию библиотеки
		 * @return string версия библиотеки
		 */
		public function version()
		{
				return $this->version;
		}

		/**
		 * Возвращает версию использованого языкового файла
		 * @return string версия языкового файла
		 */
		public function languageVersion()
		{
				return $this->languageBuild;
		}

}

?>


' --- Кінець об'єднаного модуля ---


' Функція для виклику з Excel для відмінювання повного імені
Public Function DeclineNameExcel(fullName As String, caseNumber As Integer) As String
    ' Виклик основної функції з оригінального коду
    DeclineNameExcel = DeclineFullName(fullName, caseNumber)
End Function

