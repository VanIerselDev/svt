Sub LeesGetallenTotTweemaalHetzelfdeGetal

    Dim intGetal As Integer
    Dim intVorigGetal As Integer
    Dim strTekst As String
    
    intVorigGetal = -32768
    strTekst = ""
    
    intGetal = cint(val(inputbox("Geef getal (tweemaal hetzelfde om te stoppen)")))
    Do While intGetal <> intVorigGetal
        intVorigGetal = intGetal
        strTekst = strTekst & Format(intGetal, "0") & " "        
        intGetal = cint(val(inputbox("Geef getal (tweemaal hetzelfde om te stoppen)")))
    Loop 'intGetal <> intVorigGetal
    'intGetal = intVorigGetal
    
    
    MsgBox strTekst

End Sub 'LeesGetallenTotTweemaalHetzelfdeGetal

Sub LeesGetallenTotLaatsteKleinerIsDanVorige

    Dim intGetal As Integer
    Dim intVorigGetal As Integer
    Dim strTekst As String
    
    intVorigGetal = -32768
    strTekst = ""
    
    intGetal = cint(val(inputbox("Geef getal (kleiner dan het vorig om te stoppen)")))
    Do While intGetal > intVorigGetal
        intVorigGetal = intGetal
        strTekst = strTekst & Format(intGetal, "0") & " "        
        intGetal = cint(val(inputbox("Geef getal (kleiner dan het vorig om te stoppen)")))
    Loop 'intGetal <> intVorigGetal
    'intGetal = intVorigGetal
    
    
    MsgBox strTekst

End Sub 'LeesGetallenTotLaatsteKleinerIsDanVorige

Function isEven(intGetal) As Boolean
    Dim blnVoorlopigIsEven As Boolean
    
    If (intGetal mod 2) Then
        blnVoorlopigIsEven = true
    Else 
        blnVoorlopigIsEven = false
    End If '(intGetal mod 2)
    
    isEven = blnVoorlopigIsEven
End Function 'isEven()

Sub LeesGetallenTotLaatsteEnVoorlaatstegetalAllebeiEvenOfOnevenZijn

    Dim intGetal As Integer
    Dim intVorigGetal As Integer
    Dim strTekst As String
    
    intVorigGetal = 0
    strTekst = ""
    
    intGetal = cint(val(inputbox("Geef getal (even of oneven net als het vorig om te stoppen)")))
    Do While isEven(intVorigGetal) <> isEven(intGetal)
        intVorigGetal = intGetal
        strTekst = strTekst & Format(intGetal, "0") & " "        
        intGetal = cint(val(inputbox("Geef getal (even of oneven net als het vorig om te stoppen)")))
    Loop 'intGetal <> intVorigGetal
    'intGetal = intVorigGetal
    
    
    MsgBox strTekst

End Sub 'LeesGetallenTotLaatsteEnVoorlaatstegetalAllebeiEvenOfOnevenZijn
