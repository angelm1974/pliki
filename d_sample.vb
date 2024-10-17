''' <summary>
''' Klasa testowa
''' </summary>
private Sub tekst()
    Dim liczba As Integer
    liczba = 10
    liczba1 = 20
    suma = liczba + liczba1
    MsgBox "Suma: " & suma
End Sub

Sub SingleArrayExmpl()
Dim liczby(1 To 5) As Integer
Dim i As Integer

For i = 1 To 5
    liczby(i) = i * 10
Next i
MsgBox "Wartosc dla indeksu = 3 " & liczby(3)
End Sub