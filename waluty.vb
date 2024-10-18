Option Explicit

Sub PobierzKursyWalut()
    Dim http As Object
    Dim url As String
    Dim json As String
    Dim data As Object
    Dim i As Integer
    Dim ws As Worksheet
    
    Set ws = Workbooks("Zeszyt1").Sheets.Add
    ws.Name = "Kursy walut"
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = "https://api.nbp.pl/api/exchangerates/tables/A/"
    
    http.Open "GET", url, False
    http.Send
    
    json = http.responseText
    
    Set data = JsonConverter.ParseJson(json)
    
    For i = 1 To data(1)("rates").Count
        ws.Cells(i + 1, 1).Value = data(1)("rates")(i)("currency")
        ws.Cells(i + 1, 2).Value = data(1)("rates")(i)("code")
        ws.Cells(i + 1, 3).Value = data(1)("rates")(i)("mid")
    Next i
    
    Set http = Nothing
    
    
End Sub
