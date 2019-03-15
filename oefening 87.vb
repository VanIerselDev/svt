Option Explicit

Function MaakOfferteLoodgieter(intTotaalPrijsMateriaal As Integer, intWerkuren As Integer, intAfstand As Integer, intLeeftijdWoning As Integer) As String
	'Initialisatie
	Const csngMateriaalVermenigvuldiger As Single = 1.33 '+33%
	Const csngUurPrijs As Single = 75 'Euro 75,-
	Const cintGratisAfstandKM As Integer = 10 '10 Kilometer
	Const csngPrijsPerKM As Single = 0.75 'Euro 0,75- per kilometer
	Const cintBTWKortingLeeftijd As Integer = 15 'Woning leeftijd is 15 jaar
' 	Const csngBTWKorting As Single = 1.06 '+6%  
' 	Const csngBTW As Single = 1.21 '+21%
	Const csngAlleenBTWKorting As Single = 0.06 '6%
	Const csngAlleenBTW As Single = 0.21 '21%
	
	Dim loCRLF As String
	
	Dim sngTotaalMateriaalKosten As Single
	Dim sngTotaalArbeidsKosten As Single
	Dim sngTotaalAfstandsKosten as Single
	Dim sngTotaalBelastingKosten as Single
	Dim sngTotaal As Single
	
	Dim strTekst As String
	
	strTekst = ""
	loCRLF = Chr(13) & Chr(10)
	
	'Bereken
	sngTotaalMateriaalKosten = intTotaalPrijsMateriaal * csngMateriaalVermenigvuldiger
	strTekst = strTekst & "Berekening: €" & Format(sngTotaalMateriaalKosten, "0.00") & "-" & " = €" & Format(intTotaalPrijsMateriaal, "0.00") & "-" & " * " & Format(csngMateriaalVermenigvuldiger, "0.00") & loCRLF	
	strTekst = strTekst & "Totaal bedrag materialen: €" & Format(sngTotaalMateriaalKosten, "0.00") & "-" & loCRLF
	
	sngTotaalArbeidsKosten = intWerkuren * csngUurPrijs
	strTekst = strTekst & "Berekening: €" & Format(sngTotaalArbeidsKosten, "0.00") & "-" & " = " & Format(intWerkuren, "0") & " * €" & Format(csngUurPrijs, "0.00") & "-"  & loCRLF	
	strTekst = strTekst & "Totaal bedrag arbeid: €" & Format(sngTotaalArbeidsKosten, "0.00") & "-" & loCRLF
    
	If (intAfstand > 10) Then
		sngTotaalAfstandsKosten = intAfstand * csngPrijsPerKM
        strTekst = strTekst & "Berekening: €" & Format(sngTotaalAfstandsKosten, "0.00") & "-" & " = " & Format(intAfstand, "0.00") & "KM" & " * €" & Format(csngPrijsPerKM, "0.00") & "-"  & loCRLF	
        strTekst = strTekst & "Totaal bedrag afstand: €" & Format(sngTotaalAfstandsKosten, "0.00") & "-" & loCRLF
    Else '(intAfstand <= 10)
		sngTotaalAfstandsKosten = 0
        strTekst = strTekst & "Berekening: De afstand is kleiner dan 10KM dus de kosten zijn €0.00-"  & loCRLF	
        strTekst = strTekst & "Totaal bedrag afstand: €" & Format(sngTotaalAfstandsKosten, "0.00") & "-" & loCRLF
	End If '(intAfstand > 10)
	
    sngTotaal = sngTotaalMateriaalKosten + sngTotaalArbeidsKosten + sngTotaalAfstandsKosten
    
	If (intLeeftijdWoning >= cintBTWKortingLeeftijd) Then '15 Jaar of jonger
        sngTotaalBelastingKosten = sngTotaal * csngAlleenBTWKorting
        strTekst = strTekst & "Berekening: €" & Format(sngTotaalBelastingKosten, "0.00") & "-" & " = €" & Format(sngTotaal, "0.00") & "-" & " * " & Format(csngAlleenBTWKorting, "0.00") & loCRLF	
        strTekst = strTekst & "Totaal bedrag BTW: €" & Format(sngTotaalBelastingKosten, "0.00") & "-" & loCRLF
	Else '(intLeeftijdWoning < cintBTWKortingLeeftijd) 'Ouder dan 15 jaar		
        sngTotaalBelastingKosten = sngTotaal * csngAlleenBTW
        strTekst = strTekst & "Berekening: €" & Format(sngTotaalBelastingKosten, "0.00") & "-" & " = €" & Format(sngTotaal, "0.00") & "-" & " * " & Format(csngAlleenBTW, "0.00") & loCRLF	
        strTekst = strTekst & "Totaal bedrag BTW: €" & Format(sngTotaalBelastingKosten, "0.00") & "-" & loCRLF
	End If '(intLeeftijdWoning >= cintBTWKortingLeeftijd)
	
End Function 'MaakOfferteLoodgieter

Sub OfferteLoodgieter

	'Initialisatie en inlezen
	Dim intTotaalPrijsMateriaal As Integer
	intTotaalPrijsMateriaal = Cint(Val(Inputbox("Geef de totaalprijs van het materiaal", "Totaal Prijs Materiaal", "10000")))

	Dim intWerkuren As Integer
	intWerkuren = Cint(Val(Inputbox("Geef het totaal aantal uren gewerkt", "Totaal Uren gewerkt", "8")))
	
	Dim intAfstand As Integer
	intAfstand = Cint(Val(Inputbox("Geef de totale afstand in KM", "Totale Afstand", "25")))
	
	Dim intLeeftijdWoning As Integer
	intLeeftijdWoning = Cint(Val(Inputbox("Geef de leeftijd van de woning in jaren", "Leeftijd Woning", "3")))
	
	Dim strTekst As String
	strTekst = MaakOfferteLoodgieter(intTotaalPrijsMateriaal, intWerkuren, intAfstand, intLeeftijdWoning)
    
	MsgBox strTekst
	
End Sub 'OfferteLoodgieter
