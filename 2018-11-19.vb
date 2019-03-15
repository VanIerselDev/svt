Sub ToonWoordPerLetter

    Dim strWoord As String
    Dim strSpelling As String
    Dim strTekst As String
    Dim intTeller As Integer
    
    strTekst = ""
    strSpelling = ""
    
    strWoord = Inputbox("Geef een woord", "Vraag", "El-Pueblo-de-Nuestra Señora-la-Reina-De-los-Ángeles-del-Río-de-Porciúncula")
    
    For intTeller = 1 To Len(StrWoord) Step 1
        strSpelling = strSpelling & Mid(Right(Left(strWoord, intTeller), 1), 1) & " "
    Next 'intTeller
    
    Msgbox strSpelling

End Sub 'ToonWoordPerLetter
