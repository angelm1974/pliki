Sub SingleArrayExample()
    Dim liczby(1 To 5) As Integer
    Dim i As Integer
    
    ' Przypisanie wartości do tablicy
    For i = 1 To 5
        liczby(i) = i * 10
    Next i
    
    MsgBox "Wartość z indeksu 3: " & liczby(3)
End Sub

Sub MultiDimensionalArrayExample()
    Dim macierz(1 To 2, 1 To 2) As Integer
    
    macierz(1, 1) = 10
    macierz(1, 2) = 20
    macierz(2, 1) = 30
    macierz(2, 2) = 40
    
    MsgBox "Wartość z wiersza 2, kolumny 1: " & macierz(2, 1)
End Sub

Sub DynamicArrayExample()
    Dim liczby() As Integer
    Dim i As Integer
    
    ' Dynamiczne przydzielenie rozmiaru tablicy
    ReDim liczby(1 To 5)
    
    For i = 1 To 5
        liczby(i) = i * 10
    Next i
    
    MsgBox "Wartość z indeksu 5: " & liczby(5)
End Sub

Sub ForNextLoopExample()
    Dim i As Integer
    Dim wynik As Integer
    wynik = 0
    
    For i = 1 To 5
        wynik = wynik + i
    Next i
    
    MsgBox "Suma liczb od 1 do 5: " & wynik
End Sub

Sub DoWhileLoopExample()
    Dim i As Integer
    i = 1
    Do While i <= 5
        MsgBox "Wartość i: " & i
        i = i + 1
    Loop
End Sub

Sub IfThenElseExample()
    Dim liczba As Integer
    liczba = 10
    
    If liczba > 5 Then
        MsgBox "Liczba większa niż 5"
    Else
        MsgBox "Liczba mniejsza lub równa 5"
    End If
End Sub

Sub SelectCaseExample()
    Dim dzienTygodnia As Integer
    dzienTygodnia = 3
    
    Select Case dzienTygodnia
        Case 1
            MsgBox "Poniedziałek"
        Case 2
            MsgBox "Wtorek"
        Case 3
            MsgBox "Środa"
        Case Else
            MsgBox "Inny dzień"
    End Select
End Sub

Function DodajLiczby(ByVal liczba1 As Integer, ByVal liczba2 As Integer) As Integer
    DodajLiczby = liczba1 + liczba2
End Function

Sub UseFunctionExample()
    Dim suma As Integer
    suma = DodajLiczby(5, 10)
    MsgBox "Suma: " & suma
End Sub

Function Powitaj(Optional imie As String = "Gość") As String
    Powitaj = "Witaj, " & imie & "!"
End Function

Sub UseOptionalParamsExample()
    MsgBox Powitaj() ' Wywołanie bez parametru
    MsgBox Powitaj("Mariusz") ' Wywołanie z parametrem
End Sub
