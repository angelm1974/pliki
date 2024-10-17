Option Explicit

Private Sub UserForm_Initialize()
Dim lista_klientow As Variant
Dim licznik As Long
Dim klient As Variant
Dim isUnique As Boolean
Dim i As Integer

licznik = Sheets("Zmienne").Cells(1, 2).Value
lista_klientow = Sheets("Dane").Range("A2:A" & licznik).Value2

For Each klient In lista_klientow
    isUnique = True
    
    ' Sprawdzenie, czy wartość już istnieje w ComboBoxie
    For i = 0 To cmbKlient.ListCount - 1
        If cmbKlient.List(i) = klient Then
            isUnique = False
            Exit For
        End If
    Next i
    
    ' Dodanie tylko unikalnych wartości
    If isUnique Then
        cmbKlient.AddItem klient
    End If
Next klient

    
End Sub