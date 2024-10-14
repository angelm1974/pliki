Sub ByteExample()
    Dim smallNumber As Byte
    smallNumber = 200
    MsgBox "Wartość Byte: " & smallNumber
End Sub

Sub BooleanExample()
    Dim isReady As Boolean
    isReady = True
    If isReady Then
        MsgBox "Zmiennej Boolean przypisano wartość True"
    Else
        MsgBox "Zmiennej Boolean przypisano wartość False"
    End If
End Sub

Sub IntegerExample()
    Dim number As Integer
    number = 300
    MsgBox "Liczba całkowita: " & number
End Sub

Sub LongExample()
    Dim largeNumber As Long
    largeNumber = 123456789
    MsgBox "Duża liczba: " & largeNumber
End Sub

Sub SingleExample()
    Dim singlePrecision As Single
    singlePrecision = 123.456
    MsgBox "Liczba zmiennoprzecinkowa Single: " & singlePrecision
End Sub

Sub DoubleExample()
    Dim doublePrecision As Double
    doublePrecision = 123456.789123
    MsgBox "Liczba zmiennoprzecinkowa Double: " & doublePrecision
End Sub

Sub CurrencyExample()
    Dim money As Currency
    money = 1234.56
    MsgBox "Kwota: " & money
End Sub

Sub StringExample()
    Dim tekst As String
    tekst = "Visual Basic dla Excela"
    MsgBox "Tekst: " & tekst
End Sub

Sub DateExample()
    Dim currentDate As Date
    currentDate = Now
    MsgBox "Aktualna data i czas: " & currentDate
End Sub

Sub ObjectExample()
    Dim arkusz As Worksheet
    Set arkusz = ThisWorkbook.Sheets(1)
    MsgBox "Nazwa pierwszego arkusza: " & arkusz.Name
End Sub

Sub VariantExample()
    Dim zmienna As Variant
    zmienna = "Dowolny typ danych"
    MsgBox "Typ Variant: " & zmienna
    zmienna = 123
    MsgBox "Typ Variant (liczba): " & zmienna
End Sub
