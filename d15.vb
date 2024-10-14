
' -------------------------------
' Plik VB - Operacje na Zakresach 1
' -------------------------------

' Ćwiczenie 1: Sumowanie zakresu komórek
Sub SumowanieZakresu()
    Dim suma As Double
    suma = Application.WorksheetFunction.Sum(ThisWorkbook.Sheets("Arkusz1").Range("A1:A10"))
    ThisWorkbook.Sheets("Arkusz1").Range("A11").Value = suma
    MsgBox "Suma komórek A1:A10 wynosi: " & suma
End Sub

' Ćwiczenie 2: Tworzenie tabeli z zakresu danych
Sub UtworzTabele()
    Dim zakres As Range
    Dim tabela As ListObject
    Set zakres = ThisWorkbook.Sheets("Arkusz1").Range("A1:B10")
    Set tabela = ThisWorkbook.Sheets("Arkusz1").ListObjects.Add(xlSrcRange, zakres, , xlYes)
    tabela.Name = "MojaTabela"
    MsgBox "Tabela została utworzona z zakresu A1:B10"
End Sub

' Ćwiczenie 3: Obliczanie sum częściowych
Sub SumyCzesciowe()
    Dim zakres As Range
    Set zakres = ThisWorkbook.Sheets("Arkusz1").Range("B1:B10")
    zakres.Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(2), Replace:=True, PageBreaks:=False, SummaryBelowData:=True
    MsgBox "Sumy częściowe zostały obliczone dla kolumny B"
End Sub

' Ćwiczenie 4: Filtrowanie danych
Sub FiltrowanieDanych()
    ThisWorkbook.Sheets("Arkusz1").Range("A1:B10").AutoFilter
    ThisWorkbook.Sheets("Arkusz1").Range("A1:B10").AutoFilter Field:=1, Criteria1:="Test"
    MsgBox "Dane zostały przefiltrowane, aby pokazać tylko wartości 'Test' w kolumnie A"
End Sub

' Ćwiczenie 5: Sortowanie danych
Sub SortowanieDanych()
    Dim zakres As Range
    Set zakres = ThisWorkbook.Sheets("Arkusz1").Range("A1:B10")
    zakres.Sort Key1:=ThisWorkbook.Sheets("Arkusz1").Range("A1"), Order1:=xlAscending, Header:=xlYes
    MsgBox "Dane zostały posortowane według kolumny A"
End Sub
