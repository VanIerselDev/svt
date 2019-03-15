Option Explicit

Function loCRLF()
    loCRLF = Chr(13) & Chr(10)
End Function 'loCRLF()

Sub BerekenTotaalVanAantal
    Const intMinAantalBoeken As Integer = 1
    Const strError As String = "ERROR: Aantal boeken is kleiner dan 1; het moet 1 of groter zijn. Pas dit a.u.b. aan."   
    
    Dim intAantalBoeken As Integer
    Dim intTeller As Integer
    Dim strTekst As String
    
    Dim sngPrijsBoek As Integer
    'init
    strTekst = ""
    'lees
    intAantalBoeken = Cint(Val(Inputbox("Aantal boeken?")
    If (intAantalBoeken <= 0) Then
        MsgBox strError
        intAantalBoeken = Cint(Val(Inputbox("Aantal boeken?")
    Else 'intaantalBoeken > 0
        For intTeller = intMinAantalBoeken To intAantalBoeken
            sngPrijsBoek = Cint(Val(Inputbox("Prijs boek " & Format(intTeller, "0"))))
            strTekst = "Prijs boek " & Format(intTeller, "0") & ": " & Format(sngPrijsBoek, "0.00")& loCRLF()
        Next 'intTeller
    End If '(intAantalBoeken <= 0)
    'toon
    MsgBox strTekst
    
End Sub 'BerekenTotaalVanAantal
