Option Explicit

Function loCRLF() As String
    loCRLF = Chr(13) & Chr(10)
End Function 'loCRLF()

Function OnderElkaar() As String
    OnderElkaar = loCRLF() & "--------------------------------------" & loCRLF()
End Function 'OnderElkaar()


Sub ToonTicketVanBoeken
    Dim sngPrijs As Single
    Dim intAantal As Integer
    Dim intTeller As Integer
    Dim sngTotaal As Single
    Dim strTekst As String
    'init
    strTekst = ""
    sngTotaal = 0
    'lees & verwerk
    intAantal = Cint(Val(Inputbox("Aantal de zelfde boeken")))
    Do While intAantal > 0
        sngPrijs = Cint(Val(Inputbox("Prijs van het vorige aantal gegeven boek")))
'VERWERK>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    	sngTotaal = sngTotaal + (intAantal * sngPrijs)
        strTekst = strTekst & Format(intAantal,"0") & " * " & Format(sngPrijs,"0.00") & loCRLF()
'VERWERK<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        intAantal = Cint(Val(Inputbox("Aantal boeken")))
    Loop 'intAantal <> 0
    'intAantal = 0
    strTekst = strTekst & OnderElkaar() "Totaal: " & Format(sngTotaal,"0.00") & " Euro"
    'toon
    Msgbox strTekst
End Sub 'ToonTicketVanBoeken
