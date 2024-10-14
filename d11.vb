Sub TypyZmiennych()
    Dim liczbaCalkowita As Integer
    Dim liczbaDuza As Long
    Dim liczbaZmiennoprzecinkowa As Double
    Dim tekst As String
    Dim dataCzas As Date
    Dim wartoscLogiczna As Boolean
    
    liczbaCalkowita = 100
    liczbaDuza = 123456789
    liczbaZmiennoprzecinkowa = 3.14159
    tekst = "Przykładowy tekst"
    dataCzas = Now
    wartoscLogiczna = True
    
    MsgBox "Liczba całkowita: " & liczbaCalkowita
    MsgBox "Duża liczba: " & liczbaDuza
    MsgBox "Liczba zmiennoprzecinkowa: " & liczbaZmiennoprzecinkowa
    MsgBox "Tekst: " & tekst
    MsgBox "Data i czas: " & dataCzas
    MsgBox "Wartość logiczna: " & wartoscLogiczna
End Sub
