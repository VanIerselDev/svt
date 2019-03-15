Sub ToonMaaltafels

    Dim intEersteGetal As Integer
    Dim intTweedeGetal As Integer
    Dim intUitkomst As Integer

    Dim strTekst As String

    strTekst = ""
    
    For intEersteGetal = 1 To 10
        For intTweedeGetal = 1 To 10
            intUitkomst = intEersteGetal * intTweedeGetal
            strTekst = strTekst & Format(intEersteGetal, "0") & " * " & Format(intTweedeGetal, "0") & " = " & Format(intUitkomst, "0") & " | " 
        Next 'intTweedeGetal
    Next 'intEersteGetal
    
    Msgbox strTekst
    
End Sub 'ToonMaaltafels



Function Faculteit

    Const intKleinste As Integer = -323758

    Dim strTekst As String

    strTekst = ""
    
    
    
    Msgbox strTekst

End Function 'Faculteit



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


' 1) =ALS((A14+1)<=$D$4;A14+1;"")
' 2) =ALS(ISGETAL(A15);$B$4*5;"")
' 3) ---
' 4) ---
' 5) ---
