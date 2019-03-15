Const cintEersteType As Integer = 1
Const cintCompact As Integer = 1
Const cintMiddenklAsse As Integer = 2
Const cintLuxe As Integer = 3
Const cintJeep As Integer = 4
Const cintMonovolume As Integer = 5
Const cintLaatsteType As Integer = 5

Const cstrCompact As String = "Compact"
Const cstrMiddenklAsse As String = "MiddenklAsse"
Const cstrLuxe As String = "Luxe"
Const cstrJeep As String = "Jeep"
Const cstrMonovolume As String = "Monovolume"

Const csngDagPrijsCompact As Single = 15.50
Const csngDagPrijsMiddenklAsse As Single = 17.72
Const csngDagPrijsLuxe As Single = 22.30
Const csngDagPrijsJeep As Single = 22.55
Const csngDagPrijsMonovolume As Single = 26.65

Const csngWeekPrijsCompact As Single = 100.00
Const csngWeekPrijsMiddenklAsse As Single = 112.00
Const csngWeekPrijsLuxe As Single = 149.00
Const csngWeekPrijsJeep As Single = 163.00
Const csngWeekPrijsMonovolume As Single = 180.00

Const csngPrijsPerExtraKmCompact As Single = 0.02
Const csngPrijsPerExtraKmMiddenklAsse As Single = 0.03
Const csngPrijsPerExtraKmLuxe As Single = 0.04
Const csngPrijsPerExtraKmJeep As Single = 0.05
Const csngPrijsPerExtraKmMonovolume As Single = 0.05

Const csngVerzekeringPerDagCompact As Single = 3.10
Const csngVerzekeringPerDagMiddenklAsse As Single = 4.21
Const csngVerzekeringPerDagLuxe As Single = 4.96
Const csngVerzekeringPerDagJeep As Single = 5.58
Const csngVerzekeringPerDagMonovolume As Single = 5.58

Const cintAantalKmInPrijsInbegrepen As Integer = 200

Dim strTypeWagen(cintEersteType to cintLaatsteType) As String
Dim sngDagPrijs(cintEersteType to cintLaatsteType) As Single
Dim sngWeekPrijs(cintEersteType to cintLaatsteType) As Single
Dim sngPrijsPerExtraKm(cintEersteType to cintLaatsteType) As Single
Dim sngVerzekeringPerDag(cintEersteType to cintLaatsteType) As Single

Function AantalVolledigeWeken(intDagen As Integer, intWeken As Integer) As Integer

    Dim intTEMPAantalVolledigeWeken As Integer
    
    intTEMPAantalVolledigeWeken = intWeken + ( intDagen / 7)
    
    AantalVolledigeWeken = intTEMPAantalVolledigeWeken

End Function 'AantalVolledigeWeken

Function AantalDagenInLaatsteWeek(intDagen As Integer, intWeken As Integer) As Integer

    Dim intTEMPAantalDagenInLaatsteWeek As Integer
    
    intTEMPAantalDagenInLaatsteWeek = intDagen mod 7
    
    AantalDagenInLaatsteWeek = intTEMPAantalDagenInLaatsteWeek

End Function 'AantalDagenInLaatsteWeek


Function AantalTotaalDagen(intDagen As Integer, intWeken As Integer) As Integer

    Dim intTEMPAantalDagenInLaatsteWeekAantalTotaalDagen As Integer
    
    intTEMPAantalTotaalDagen = (intWeken * 7) + intDagen
    
    AantalTotaalDagen = intTEMPAantalTotaalDagen

End Function 'AantalTotaalDagen

Sub Init

    strTypeWagen(cintCompact) = cstrCompact
    strTypeWagen(cintMiddenklAsse) = cstrMiddenklAsse
    strTypeWagen(cintLuxe) = cstrLuxe
    strTypeWagen(cintJeep) = cstrJeep
    strTypeWagen(cintMonovolume) = cstrMonovolume

    sngDagPrijs(cintCompact) = csngDagPrijsCompact
    sngDagPrijs(cintMiddenklAsse) = csngDagPrijsMiddenklAsse
    sngDagPrijs(cintLuxe) = csngDagPrijsLuxe
    sngDagPrijs(cintJeep) = csngDagPrijsJeep
    sngDagPrijs(cintMonovolume) = csngDagPrijsMonovolume

    sngWeekPrijs(cintCompact) = csngWeekPrijsCompact
    sngWeekPrijs(cintMiddenklAsse) = csngWeekPrijsMiddenklAsse
    sngWeekPrijs(cintLuxe) = csngWeekPrijsLuxe
    sngWeekPrijs(cintJeep) = csngWeekPrijsJeep
    sngWeekPrijs(cintMonovolume) = csngWeekPrijsMonovolume

    sngPrijsPerExtraKm(cintCompact) = csngPrijsPerExtraKmCompact
    sngPrijsPerExtraKm(cintMiddenklAsse) = csngPrijsPerExtraKmMiddenklAsse
    sngPrijsPerExtraKm(cintLuxe) = csngPrijsPerExtraKmLuxe
    sngPrijsPerExtraKm(cintJeep) = csngPrijsPerExtraKmJeep
    sngPrijsPerExtraKm(cintMonovolume) = csngPrijsPerExtraKmMonovolume

    sngVerzekeringPerDag(cintCompact) = csngVerzekeringPerDagCompact
    sngVerzekeringPerDag(cintMiddenklAsse) = csngVerzekeringPerDagMiddenklAsse
    sngVerzekeringPerDag(cintLuxe) = csngVerzekeringPerDagLuxe
    sngVerzekeringPerDag(cintJeep) = csngVerzekeringPerDagJeep
    sngVerzekeringPerDag(cintMonovolume) = csngVerzekeringPerDagMonovolume
  
    Dim intType As Integer
    Dim intAantalDagen As Integer
    Dim intAantalWeken As Integer
    Dim intUiteidelijkAantalWeken As Integer
    Dim intUiteidelijkAantalDagen As Integer
    Dim intAantalKm As Integer
    Dim blnVerzekeringNodig As Boolean
    Dim intMsgBoxVerzekeringRetour As Integer
    Dim intTotaalAantalDagen As Integer
    
    Dim intUitEindelijkAantalKm As Integer
    
    Dim sngPrijsVerzekering As Integer
    Dim sngTotaalPrijs As Single
    Dim strTekst As String
    
    strTekst = ""

End Sub 'Init

Sub WagenVerhuur    
    'init  
    Call Init
    
    'lees
    intType = Cint(Val(InputBox("Kies 1: " & cstrCompact & "; Kies 2: " & cstrMiddenklAsse & "; Kies 3: " & cstrLuxe & "; Kies 4: " & cstrJeep & "; Kies 4: " & cstrMonovolume, "Geef het type nummer van de wagen:")))
    intAantalWeken = Cint(Val(InputBox("Voor hoelang zou u deze wagen willen huren? (Weken)", "Aantal weken dat u wilt huren:")))
    intAantalDagen = Cint(Val(InputBox("Voor hoelang zou u deze wagen willen huren? (Dagen)", "Aantal dagen dat u wilt huren:")))
    intAantalKm = Cint(Val(InputBox("Hoeveel kilometer wilt u gaan rijden met deze wagen?", "Aantal kilometer dat u wilt gaan rijden:")))
    
    intMsgBoxVerzekeringRetour = MsgBox("Kies ja als u een bijkomende verzekering en nee als u dit niet wilt", MB_YESNO, "Bijkomende verzekering:")
    
    'verwerk
    
    intTotaalAantalDagen = AantalTotaalDagen(intAantalDagen, intAantalWeken)
    
	intUiteidelijkAantalWeken = AantalVolledigeWeken(intAantalDagen, intAantalWeken)
	intUiteidelijkAantalDagen = AantalDagenInLaatsteWeek(intAantalDagen, intAantalWeken)
	
	If intAantalKm <= 200 Then
        intUitEindelijkAantalKm = 0
    Else
        intUitEindelijkAantalKm = intAantalKm - 200
    End If 'intAantalKm <= 200
	
	If intMsgBoxVerzekeringRetour = IDYES Then
        sngPrijsVerzekering = intTotaalAantalDagen * sngVerzekeringPerDag(intType)
	Else 
        sngPrijsVerzekering = 0
	End If' intMsgBoxVerzekeringRetour <> IDYES || IDNO
    sngTotaalPrijs = (intUiteidelijkAantalWeken * sngWeekPrijs(intType)) + (intUiteidelijkAantalDagen * sngDagPrijs(intType)) + (intUitEindelijkAantalKm * sngPrijsPerExtraKm(intType)) + sngPrijsVerzekering
    'toon
    strTekst = strTekst & "U kiest de wagen: " & strTypeWagen(intType) & loCRLF
    strTekst = strTekst & "U rijdt met de wagen: " & Format(intAantalKm, "0") & " kilometer" & loCRLF
    strTekst = strTekst & "Dit kost: " & Format(intUitEindelijkAantalKm * sngPrijsPerExtraKm(intType), "0.00") & " euro" & loCRLF
    strTekst = strTekst & "U rijdt met de wagen: " & Format(intUiteidelijkAantalDagen, "0") & " dagen" & loCRLF
    strTekst = strTekst & "Dit kost: " & Format(intUiteidelijkAantalDagen * sngDagPrijs(intType), "0.00") & " euro" & loCRLF
    strTekst = strTekst & "U rijdt met de wagen: " & Format(intUiteidelijkAantalWeken, "0") & " weken" & loCRLF
    strTekst = strTekst & "Dit kost: " & Format(intUiteidelijkAantalWeken * sngWeekPrijs(intType), "0.00") & " euro" & loCRLF
    strTekst = strTekst & "Uw verzekering kost: " & Format(sngPrijsVerzekering, "0.00") & " euro" & loCRLF
    strTekst = strTekst & "De totaal prijs is: " & Format(sngTotaalPrijs, "0.00") & " euro"
    
    MsgBox strTekst

End Sub 'WagenVerhuur
