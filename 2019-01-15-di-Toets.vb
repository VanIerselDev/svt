Option Explicit

Function loCRLF() As String
	
	loCRLF = Chr(13) & Chr(10)
	
End Function 'loCRLF

Function loStreep() As String

    loStreep = "---------------" & loCRLF()

End Function 'loStreep()

Sub ToonTicketVanBoekenMetKorting
    'defineer
    Const csngGeenKorting As Single = 0
	Const csngHonderdEuroKorting As Single = 5
	Const csngHonderVijftigEuroKorting As Single = 7.5
	Const csngTweeHonderdEuroKorting As Single = 10
	
	Const cintMax As Integer = 32767
	
    Dim strTekst As String
	
	Dim sngPrijs As Single
	Dim intAantal As Integer
	Dim sngTotaal As Single
	Dim sngSubTotaal As Single
	Dim sngTotaalBoek As Single
	Dim intKortingPercentage As Integer
	Dim sngKortingBedrag As Single
	
	Dim intTeller As Integer
	
	'init
	sngSubTotaal = 0
	intKortingPercentage = 0
	sngKortingBedrag = 0
	strTekst = ""
	
	'lees
	intAantal = Cint(Val(Inputbox("Aantal boeken?")))
	Do While intAantal <> 0
		sngPrijs = Csng(Val(Inputbox("Prijs per boek?")))
		
		'bereken
		sngTotaalBoek = intAantal * sngPrijs
		sngSubTotaal = sngSubTotaal + sngTotaalBoek
		
		strTekst = strTekst & Format(intAantal, "0") & " * " & Format(sngPrijs, "0.00") & " Euro = " & Format(sngTotaalBoek, "0.00") & " Euro" & loCRLF() 
		
		intAantal = Cint(Val(Inputbox("Aantal boeken?")))
	Loop 'intAantal <> 0
	'intAantal = 0
	
	strTekst = strTekst & loStreep()
	strTekst = strTekst & "Totaal: " & Format(sngSubTotaal, "0.00") & " Euro" & loCRLF()
	
	Select Case sngSubTotaal
        Case 0 To 99.99
            intKortingPercentage = csngGeenKorting
        Case 100 To 149.99
            intKortingPercentage = csngHonderdEuroKorting
        Case 150 To 199.99
            intKortingPercentage = csngHonderVijftigEuroKorting      
        Case 200 To cintMax
            intKortingPercentage = csngTweeHonderdEuroKorting       
	End Select
	sngKortingBedrag = sngSubTotaal * (0 + (intKortingPercentage / 100)) 
	sngTotaal = sngSubTotaal * (1 - (intKortingPercentage / 100))
    strTekst = strTekst & "Korting %: " & Format(intKortingPercentage, "0") & loCRLF()
    strTekst = strTekst & "Korting €: " & Format(sngKortingBedrag, "0.00") & " Euro" & loCRLF()
    strTekst = strTekst & "Eindtotaal: " & Format(sngTotaal, "0.00") & " Euro" & loCRLF()
	strTekst = strTekst & loStreep()
	
	'toon
	Msgbox strTekst
	
End Sub 'ToonTicketVanBoekenMetKorting
