
' -------------------------------
' Plik VB - Operacje na Zakresach 2
' -------------------------------

' Ćwiczenie 6: Usunięcie filtra z zakresu danych
Sub UsunFiltr()
    If ThisWorkbook.Sheets("Arkusz1").AutoFilterMode Then
        ThisWorkbook.Sheets("Arkusz1").AutoFilterMode = False
    End If
    MsgBox "Filtr został wyłączony"
End Sub

' Ćwiczenie 7: Zaznaczanie dynamicznego zakresu danych
Sub DynamicznyZakres()
    Dim zakres As Range
    Set zakres = ThisWorkbook.Sheets("Arkusz1").Range("A1", ThisWorkbook.Sheets("Arkusz1").Range("A1").End(xlDown))
    MsgBox "Dynamiczny zakres został zaznaczony: " & zakres.Address
End Sub

' Ćwiczenie 8: Podświetlanie komórek, które spełniają określony warunek
Sub PodswietlanieKomorek()
    Dim zakres As Range
    Dim komorka As Range
    Set zakres = ThisWorkbook.Sheets("Arkusz1").Range("A1:A10")
    For Each komorka In zakres
        If komorka.Value > 50 Then
            komorka.Interior.Color = RGB(255, 255, 0)
        End If
    Next komorka
    MsgBox "Komórki z wartościami większymi niż 50 zostały podświetlone"
End Sub

' Ćwiczenie 9: Kopiowanie zakresu danych do innego arkusza
Sub KopiowanieZakresu()
    Dim zakres As Range
    Set zakres = ThisWorkbook.Sheets("Arkusz1").Range("A1:B10")
    zakres.Copy Destination:=ThisWorkbook.Sheets("Arkusz2").Range("A1")
    MsgBox "Zakres A1:B10 został skopiowany do Arkusza 2"
End Sub

sub Autosumowanie(ByVal Nazwisko as String,ByVal Cena Currency )
    dim wartoscZamowiena as Currency
    Sheets("test").cells(1,1).value = Nazwisko
    Sheets("test").cells(1,2).value = Cena
end sub