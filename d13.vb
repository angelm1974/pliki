Sub SubstringExample()
    Dim tekst As String
    tekst = "Visual Basic for Applications"
    
    ' Wyciągnięcie podciągu od 8 do 12 znaku
    MsgBox "Podciąg od 8 do 12 znaku: " & Mid(tekst, 8, 5)
End Sub

Sub SingleCharacterExample()
    Dim tekst As String
    tekst = "Excel VBA"
    
    ' Wyciągnięcie 1 znaku
    MsgBox "Pierwszy znak ciągu: " & Mid(tekst, 1, 1)
    
    ' Wyciągnięcie ostatniego znaku
    MsgBox "Ostatni znak ciągu: " & Mid(tekst, Len(tekst), 1)
End Sub

Sub ConcatenateStrings()
    Dim tekst1 As String
    Dim tekst2 As String
    tekst1 = "Visual Basic"
    tekst2 = " for Excel"
    
    MsgBox "Połączone ciągi: " & tekst1 & tekst2
End Sub

Sub StringLength()
    Dim tekst As String
    tekst = "Visual Basic"
    
    MsgBox "Długość ciągu: " & Len(tekst)
End Sub

Sub ChangeCase()
    Dim tekst As String
    tekst = "Visual Basic for Applications"
    
    MsgBox "Wielkie litery: " & UCase(tekst)
    MsgBox "Małe litery: " & LCase(tekst)
End Sub

Sub FindCharacterPosition()
    Dim tekst As String
    tekst = "Visual Basic for Applications"
    
    ' Znalezienie pozycji pierwszego wystąpienia znaku "B"
    MsgBox "Pozycja litery 'B': " & InStr(tekst, "B")
    
    ' Znalezienie pozycji pierwszego wystąpienia podciągu "for"
    MsgBox "Pozycja podciągu 'for': " & InStr(tekst, "for")
End Sub

Sub ReplaceSubstring()
    Dim tekst As String
    tekst = "Visual Basic for Applications"
    
    ' Zamiana "Basic" na "Advanced"
    tekst = Replace(tekst, "Basic", "Advanced")
    
    MsgBox "Po zamianie: " & tekst
End Sub

Sub TrimSpaces()
    Dim tekst As String
    tekst = "   Visual Basic   "
    
    MsgBox "Tekst po usunięciu spacji: '" & Trim(tekst) & "'"
End Sub

Sub CheckStringStartEnd()
    Dim tekst As String
    tekst = "Visual Basic"
    
    ' Sprawdzenie, czy ciąg zaczyna się na "Visual"
    If Left(tekst, 6) = "Visual" Then
        MsgBox "Ciąg zaczyna się od 'Visual'"
    End If
    
    ' Sprawdzenie, czy ciąg kończy się na "Basic"
    If Right(tekst, 5) = "Basic" Then
        MsgBox "Ciąg kończy się na 'Basic'"
    End If
End Sub

Sub SplitStringExample()
    Dim tekst As String
    Dim czesci() As String
    tekst = "Visual,Basic,for,Applications"
    
    ' Rozdzielenie ciągu na tablicę w oparciu o przecinek
    czesci = Split(tekst, ",")
    
    MsgBox "Pierwszy element tablicy: " & czesci(0)
    MsgBox "Drugi element tablicy: " & czesci(1)
End Sub

Sub RepeatStringExample()
    Dim tekst As String
    tekst = "VB"
    
    ' Powtarzanie ciągu 5 razy
    MsgBox "Powtórzony ciąg: " & String(5, tekst)
End Sub

Sub ReverseString()
    Dim tekst As String
    Dim reversedTekst As String
    Dim i As Integer
    tekst = "Visual Basic"
    
    ' Odwracanie ciągu
    For i = Len(tekst) To 1 Step -1
        reversedTekst = reversedTekst & Mid(tekst, i, 1)
    Next i
    
    MsgBox "Odwrócony ciąg: " & reversedTekst
End Sub

Sub ArithmeticOperations()
    Dim num1 As Integer
    Dim num2 As Integer
    Dim wynik As Integer
    
    num1 = 10
    num2 = 5
    
    wynik = num1 + num2
    MsgBox "Dodawanie: " & num1 & " + " & num2 & " = " & wynik
    
    wynik = num1 - num2
    MsgBox "Odejmowanie: " & num1 & " - " & num2 & " = " & wynik
    
    wynik = num1 * num2
    MsgBox "Mnożenie: " & num1 & " * " & num2 & " = " & wynik
    
    wynik = num1 / num2
    MsgBox "Dzielenie: " & num1 & " / " & num2 & " = " & wynik
End Sub

Sub IntegerDivisionAndModulo()
    Dim num1 As Integer
    Dim num2 As Integer
    
    num1 = 17
    num2 = 5
    
    MsgBox "Dzielenie całkowite: " & num1 & " \ " & num2 & " = " & (num1 \ num2)
    MsgBox "Reszta z dzielenia: " & num1 & " Mod " & num2 & " = " & (num1 Mod num2)
End Sub

Sub ChangeSign()
    Dim liczba As Integer
    liczba = 10
    
    MsgBox "Liczba dodatnia: " & liczba
    liczba = -liczba
    MsgBox "Po zmianie znaku: " & liczba
End Sub

Sub RoundingNumbers()
    Dim liczba As Double
    liczba = 3.14159
    
    MsgBox "Liczba pierwotna: " & liczba
    MsgBox "Zaokrąglona: " & Round(liczba, 2) ' Zaokrąglenie do 2 miejsc po przecinku
End Sub

Sub RandomNumbers()
    Dim liczbaLosowa As Double
    
    ' Generowanie liczby losowej między 0 a 1
    liczbaLosowa = Rnd
    MsgBox "Liczba losowa między 0 a 1: " & liczbaLosowa
    
    ' Generowanie liczby losowej między 1 a 100
    liczbaLosowa = Int((100 * Rnd) + 1)
    MsgBox "Liczba losowa między 1 a 100: " & liczbaLosowa
End Sub

Sub MathFunctions()
    Dim liczba As Double
    liczba = 16
    
    MsgBox "Pierwiastek kwadratowy z 16: " & Sqr(liczba)
    
    ' Potęgowanie
    Dim wynik As Double
    wynik = 2 ^ 3 ' 2 do potęgi 3
    MsgBox "2 do potęgi 3: " & wynik
End Sub

Sub TrigonometricFunctions()
    Dim kat As Double
    kat = 45 ' Wartość kąta w stopniach
    
    ' Konwersja kąta na radiany (funkcje trygonometryczne w VBA używają radianów)
    kat = kat * (3.14159 / 180)
    
    MsgBox "Sinus 45 stopni: " & Sin(kat)
    MsgBox "Cosinus 45 stopni: " & Cos(kat)
    MsgBox "Tangens 45 stopni: " & Tan(kat)
End Sub

Sub CurrencyOperations()
    Dim kwota1 As Currency
    Dim kwota2 As Currency
    Dim suma As Currency
    
    kwota1 = 100.75
    kwota2 = 200.50
    
    suma = kwota1 + kwota2
    MsgBox "Suma kwot: " & Format(suma, "Currency")
End Sub
