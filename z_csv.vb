Sub ImportujDaneZCSV()
    Dim ws As Worksheet
    Dim PlikCSV As String
    Dim Wiersz As Integer

    ' Ścieżka do pliku CSV
    PlikCSV = Application.GetOpenFilename("Pliki CSV (*.csv), *.csv", , "Wybierz plik CSV do importu")
    
    ' Sprawdź, czy wybrano plik
    If PlikCSV = "False" Then Exit Sub
    
    ' Utwórz nowy arkusz do importowania danych
    Set ws = Worksheets.Add
    
    ' Otwórz plik CSV
    Open PlikCSV For Input As #1
    
    ' Czytaj plik linia po linii
    Wiersz = 1
    Do Until EOF(1)
        Line Input #1, Linia
        ws.Cells(Wiersz, 1).Value = Linia
        Wiersz = Wiersz + 1
    Loop
    
    ' Zamknij plik
    Close #1
End Sub