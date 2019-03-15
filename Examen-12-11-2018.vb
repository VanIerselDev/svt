Option Explicit

Function loCRLF() As String
	
	loCRLF = Chr(13) & Chr(10)
	
End Function 'loCRLF

Function Omgekeerd(strWoord As String) As String

    Dim strVoorlopigOmgekeerd As String
    Dim strWoordOmgekeerd As String
    Dim intWoordLengte As Integer
    Dim intTeller As Integer
    
    strWoordOmgekeerd = ""
    intWoordLengte = Len(strWoord)
    
    For intTeller = intWoordLengte To 1 Step -1
        strWoordOmgekeerd = strWoordOmgekeerd & Mid(strWoord, intTeller, 1)
    Next 'intTeller
    
    strVoorlopigOmgekeerd = strWoordOmgekeerd
    Omgekeerd = strVoorlopigOmgekeerd

End Function 'Omgekeerd

Function ZonderKlinkers(strWoord As String) As String

	Dim strVoorlopigZonderKlinkers As String
	Dim strWoordZonderKlinkers As String
	
	
	
    ZonderKlinkers = strVoorlopigZonderKlinkers

End Function 'ZonderKlinkers

Sub Woordspelletje

    Dim strTekst As String
    Dim strWoord As String
    Dim intGetal As Integer
    
    strTekst = ""
    
    Do While strWoord <> "stop"
    	strWoord = Inputbox("Keuze Woord?")
    	intGetal = Cint(Val(Inputbox("Keuze Getal?")))
    	Select Case intGetal
    		Case 1
    			strTekst = strTekst & Omgekeerd(strWoord) & loCRLF()
    		Case 2
    			strTekst = strTekst & ZonderKlinkers(strWoord) & loCRLF()
    		Case 3
    			strTekst = strTekst & AlleenKlinkers(strWoord) & loCRLF()
    	End Select 'intGetal
    Loop ' strWoord <> "stop"
    'strWoord = "stop"
         
    Msgbox strTekst
    
End Sub 'Woordspelletje









Sub Maaltafels

    Dim strTekst As String
    Dim intTeller As Integer
    Dim intMaal As Integer
    Dim intUitkomst As Integer
    
    For intTeller = 1 To 10
        For intMaal = 1 To 10
            intUitkomst = intTeller * intMaal
            strTekst = strTekst & Format(intTeller, "0") & " * " & Format(intMaal, "0") & " = " & Format(intUitkomst, "0") & " "
        Next 'intTeller
        strTekst = strTekst & loCRLF()
    Next 'intTeller
    
    Msgbox strTekst

End Sub 'Maaltafels






Function ZijnBeideNegatief(intGetal1 As Integer, intGetal2 As Integer) As Boolean
	
	Dim blnVoorlopigZijnBeideNegatief As Boolean
	
	blnVoorlopigZijnBeideNegatief = False
	
    If (intGetal1 < 0) Then
        If (intGetal2 < 0) Then
            blnVoorlopigZijnBeideNegatief = True
        End If '(intGetal < 0)
    End If '(intGetal < 0)
    
	ZijnBeideNegatief = blnVoorlopigZijnBeideNegatief
	
End Function 'ZijnBeideNegatief

Sub ToonAlleGetallenTotTweeNegatieveNaElkaar
    
	Const intHoogste As Integer = 32767
	Dim intGetal As Integer
	Dim intVorigGetal As Integer
	Dim blnStop As Boolean
	
	blnStop = false
	
    intVorigGetal = intHoogste
        
    intGetal = Cint(Val(Inputbox("Geef Getal!!!")))
    Do Until blnStop = true
        blnStop = ZijnBeideNegatief(intGetal, intVorigGetal)
        intVorigGetal = intGetal
        intGetal = Cint(Val(Inputbox("Geef Getal!!!")))
    Loop 'blnStop = false
    'blnStop = true
    
End Sub 'ToonAlleGetallenTotTweeNegatieveNaElkaar




Sub ToonTicketVanBoekenMetKorting
	
	Const cintHonderdEuroKorting As Integer = 5
	Const cintHonderVijftigEuroKorting As Integer = 7.5
	Const cintTweeHonderdEuroKorting As Integer = 10
	
	Dim strTekst As String
	
	Dim sngPrijs As Single
	Dim intAantal As Integer
	Dim sngTotaal As Single
	
	Dim intTeller As Integer
	
	intAantal = Cint(Val(Inputbox("Aantal boeken?")))
	Do While intAantal <> 0
		sngPrijs = Csng(Val(Inputbox("Prijs per boek?")))
		
		sngTotaal = intAantal * sngPrijs
		
		intAantal = Cint(Val(Inputbox("Aantal boeken?")))
	Loop 'intAantal <> 0
	'intAantal = 0
	
	Msgbox strTekst
	
End Sub 'ToonTickerVanBoekenMetKorting








Sub AutoDelen

    Const csngSTARTUurPrijs As Single = 2.00
    Const csngSTARTKilometerPrijs As Single = 0.33
    Const csngSTARTMaandelijksAbonnement As Single = 8.00
    Const csngSTARTElektronischeKorting As Single = -1.00
    
    Const csngBONUSUurPrijs As Single = 1.75
    Const csngBONUSKilometerPrijs As Single = 0.24
    Const csngBONUSMaandelijksAbonnement As Single = 8.00
    Const csngBONUSElektronischeKorting As Single = -1.00
    
    Const csngCOMFORTUurPrijs As Single = 1.55
    Const csngCOMFORTKilometerPrijs As Single = 0.21
    Const csngCOMFORTMaandelijksAbonnement As Single = 22.00
    Const csngCOMFORTElektronischeKorting As Single = -1.00
    
    Dim sngSTARTKosten As Single
    Dim sngBONUSKosten As Single
    Dim sngCOMFORTKosten As Single

    Dim sngSTARTUrenTotaalPrijs As Single
    Dim sngSTARTKilometersTotaalPrijs As Single
    Dim sngBONUSUrenTotaalPrijs As Single
    Dim sngBONUSKilometersTotaalPrijs As Single
    Dim sngCOMFORTUrenTotaalPrijs As Single
    Dim sngCOMFORTKilometersTotaalPrijs As Single
    
    Dim sngUren As Single
    Dim sngKilometers As Single
    
    Dim strTekst As String
    strTekst = ""
    
    sngUren = Csng(Val(Inputbox("Aantal uren rijdend per maand?")
    sngKilometers = Csng(Val(Inputbox("Aantal kilometers rijdend per maand?")
    
    sngSTARTUrenTotaalPrijs = sngUren * csngSTARTUurPrijs
    sngSTARTKilometersTotaalPrijs = sngKilometers * csngSTARTKilometerPrijs
    
    sngBONUSUrenTotaalPrijs = sngUren * csngBONUSUurPrijs
    sngBONUSKilometersTotaalPrijs = sngKilometers * csngBONUSKilometerPrijs
    
    sngCOMFORTUrenTotaalPrijs = sngUren * csngCOMFORTUurPrijs
    sngCOMFORTKilometersTotaalPrijs = sngKilometers * csngCOMFORTKilometerPrijs
    
    sngSTARTKosten = sngStartUrenTotaalPrijs + sngSTARTKilometersTotaalPrijs + csngSTARTMaandelijksAbonnement + csngSTARTElektronischeKorting
    sngBONUSKosten = sngStartUrenTotaalPrijs + sngBONUSKilometersTotaalPrijs + csngBONUSMaandelijksAbonnement + csngBONUSElektronischeKorting
    sngCOMFORTKosten = sngStartUrenTotaalPrijs + sngCOMFORTKilometersTotaalPrijs + csngCOMFORTMaandelijksAbonnement + csngCOMFORTElektronischeKorting
    
    Select Case sngKilometers
    	Case <= 50
	        strTekst = strTekst & "Het beste tarief voor u, die " & Format(sngKilometers, "0.00 Km") & " en " & Format(sngUren, "0 u") & ", is het START tarief" & loCRLF()
	        strTekst = strTekst & "Kostenoverzicht voor één maand:" & loCRLF()
	        strTekst = strTekst & "Uurprijs: " & Format(sngSTARTUrenTotaalPrijs, "€0.00-") & loCRLF()
	        strTekst = strTekst & "KilometerPrijs: " & Format(sngSTARTKilometersTotaalPrijs, "€0.00-") & loCRLF()
	        strTekst = strTekst & "Maandelijks abonnement: " & Format(csngSTARTMaandelijksAbonnement, "€0.00-") & loCRLF()
	        strTekst = strTekst & "Elektronische Korting: €" & Format(csngSTARTElektronischeKorting, "0.00-") & loCRLF()
	        strTekst = strTekst & "Kosten: " & Format(sngSTARTKosten, "€0.00-") & loCRLF()
'	        strTekst = strTekst & "START-tarief: " & Format(sngSTARTKosten, "€0.00-") & loCRLF()
	        strTekst = strTekst & "BONUS-tarief: " & Format(sngBONUSKosten, "€0.00-") & loCRLF() 
	        strTekst = strTekst & "COMFORT-tarief: " & Format(sngCOMFORTKosten, "€0.00-" & loCRLF() 
    	Case <= 300 
	        strTekst = strTekst & "Het beste tarief voor u, die " & Format(sngKilometers, "0.00 Km") & " en " & Format(sngUren, "0 u") & ", is het BONUS tarief" & loCRLF()
	        strTekst = strTekst & "Kostenoverzicht voor één maand:" & loCRLF()
	        strTekst = strTekst & "Uurprijs: " & Format(sngBONUSUrenTotaalPrijs, "€0.00-") & loCRLF()
	        strTekst = strTekst & "KilometerPrijs: " & Format(sngBONUSKilometersTotaalPrijs, "€0.00-") & loCRLF()
	        strTekst = strTekst & "Maandelijks abonnement: " & Format(csngBONUSMaandelijksAbonnement, "€0.00-") & loCRLF()
	        strTekst = strTekst & "Elektronische Korting: €" & Format(csngBONUSElektronischeKorting, "0.00-") & loCRLF()
	        strTekst = strTekst & "Kosten: " & Format(sngBONUSKosten, "€0.00-") & loCRLF()
	        strTekst = strTekst & "START-tarief: " & Format(sngSTARTKosten, "€0.00-") & loCRLF()
'	        strTekst = strTekst & "BONUS-tarief: " & Format(sngBONUSKosten, "€0.00-") & loCRLF() 
	        strTekst = strTekst & "COMFORT-tarief: " & Format(sngCOMFORTKosten, "€0.00-" & loCRLF() 
    	Case > 300
	        strTekst = strTekst & "Het beste tarief voor u, die " & Format(sngKilometers, "0.00 Km") & " en " & Format(sngUren, "0 u") & ", is het COMFORT tarief" & loCRLF()
	        strTekst = strTekst & "Kostenoverzicht voor één maand:" & loCRLF()
	        strTekst = strTekst & "Uurprijs: " & Format(sngCOMFORTUrenTotaalPrijs, "€0.00-") & loCRLF()
	        strTekst = strTekst & "KilometerPrijs: " & Format(sngCOMFORTKilometersTotaalPrijs, "€0.00-") & loCRLF()
	        strTekst = strTekst & "Maandelijks abonnement: " & Format(csngCOMFORTMaandelijksAbonnement, "€0.00-") & loCRLF()
	        strTekst = strTekst & "Elektronische Korting: €" & Format(csngCOMFORTElektronischeKorting, "0.00-") & loCRLF()
	        strTekst = strTekst & "Kosten: " & Format(sngCOMFORTKosten, "€0.00-") & loCRLF()
	        strTekst = strTekst & "START-tarief: " & Format(sngSTARTKosten, "€0.00-") & loCRLF()
	        strTekst = strTekst & "BONUS-tarief: " & Format(sngBONUSKosten, "€0.00-") & loCRLF() 
'	        strTekst = strTekst & "COMFORT-tarief: " & Format(sngCOMFORTKosten, "€0.00-" & loCRLF() 
    End Select 'sngKilometers
    
    Msgbox strTekst
    
End Sub 'AutoDelen
