Const cstrAantalKmPerMaand As String = "Geef aantal Km per maand"
Const cstrAantalurenPerMaand As String = "Geef aantal uren per maand"
Const csngKortingElektronisch As Single = 1

Const cintEersteFormule As integer = 1
Const cintFormuleStart As integer = 1
Const cintFormuleBonus As integer = 2
Const cintFormuleComfort As integer = 3
Const cintLaatsteFormule As integer = 3

Const cstrFormuleStart As string = "Start"
Const cstrFormuleBonus As string = "Bonus"
Const cstrFormuleComfort As string = "Comfort"

Const csngUurprijsStart As Single = 2
Const csngUurprijsBonus As Single = 1.75
Const csngUurprijsComfort As Single = 1.55

Const csngKmprijsStart As Single = 0.33
Const csngKmprijsBonus As Single = 0.24
Const csngKmprijsComfort As Single = 0.21

Const csngMaandelijksAbonnementStart As Single = 4
Const csngMaandelijksAbonnementBonus As Single = 8
Const csngMaandelijksAbonnementComfort As Single = 22

Dim strFormule(cintEersteFormule to cintLaatsteFormule) As String
Dim sngUurprijs(cintEersteFormule to cintLaatsteFormule) As Single
Dim sngKmprijs(cintEersteFormule to cintLaatsteFormule) As Single
Dim sngMaandelijksAbonnement(cintEersteFormule to cintLaatsteFormule) As Single
Dim sngTarief(cintEersteFormule to cintLaatsteFormule) As Single

Dim sngUurprijsPerMaand(cintEersteFormule to cintLaatsteFormule) As Single
Dim sngKmprijsPerMaand(cintEersteFormule to cintLaatsteFormule) As Single
Dim sngPrijsAbonnementPerMaand(cintEersteFormule to cintLaatsteFormule) As Single
Dim KortingPerMaand(cintEersteFormule to cintLaatsteFormule) As Single
Dim sngTotaalPerMaand(cintEersteFormule to cintLaatsteFormule) As Single



Sub Initialiseer
  strFormule(cintFormuleStart) = cstrFormuleStart
  strFormule(cintFormuleBonus) = cstrFormuleBonus
  strFormule(cintFormuleComfort) = cstrFormuleComfort

  sngUurprijs(cintFormuleStart) = csngUurprijsStart
  sngUurprijs(cintFormuleBonus) = csngUurprijsBonus
  sngUurprijs(cintFormuleComfort) = csngUurprijsComfort
  
  sngKmprijs(cintFormuleStart) = csngKmprijsStart
  sngKmprijs(cintFormuleBonus) = csngKmprijsBonus
  sngKmprijs(cintFormuleComfort) = csngKmprijsComfort
  
  sngMaandelijksAbonnement(cintFormuleStart) = csngMaandelijksAbonnementStart
  sngMaandelijksAbonnement(cintFormuleBonus) = csngMaandelijksAbonnementBonus
  sngMaandelijksAbonnement(cintFormuleComfort) = csngMaandelijksAbonnementComfort
  
End Sub 'Initialiseer

Function Goedkoopste(sngTotaalPerMaand As Single) As Integer
  Dim intTeller1 As Integer
  Dim intTeller2 As Integer
  
  For intTeller1 = cintEersteFormule To cintLaatsteFormule - 1
      For intTeller2 = cintEersteFormule + 1 To cintLaatsteFormule
        If sngTotaalPerMaand(intTeller1) < sngTotaalPerMaand(intTeller2) Then
            Goedkoopste = intTeller1
        End if 'sngTotaalPerMaand(intTeller1) >= sngTotaalPerMaand(intTeller2)
    Next 'intTeller2
  Next 'intTeller1
End Function 'Goedkoopste

Function UurprijsPerMaand(sngUurprijs As Single, intAantalUrenPerMaand As integer) As Single
  Dim sngUurprijsPerMaand As Single
  sngUurprijsPerMaand = intAantalUrenPerMaand * sngUurprijs
    strTekst  = strTekst & "Uurprijs per maand: " & Format(sngUurprijsPerMaand, "0.00") & " = " & Format(intAantalUrenPerMaand, "0") & " * " & Format(sngUurprijs, "0.00") & loCRLF
    
    UurprijsPerMaand = sngUurprijsPerMaand
End Function 'UurprijsPerMaand

Function KmprijsPerMaand(sngKmprijs As Single, intAantalKmPerMaand As integer) As Single
  Dim sngKmprijsPerMaand As Single
  sngKmprijsPerMaand = intAantalKmPerMaand * sngKmprijs
  strTekst = strTekst & "KM prijs per maand: " & Format(sngKmprijsPerMaand, "0.00") & " = " & Format(intAantalKmPerMaand, "0") & " * " & Format(sngKmprijs, "0") & loCRLF
  
  KmPrijsPerMaand = sngKmprijsPerMaand
End Function '

Function PrijsAbonnementPerMaand(sngMaandelijksAbonnement As Single) As Single
'  PrijsAbonnementPerMaand = sngMaandelijksAbonnement(intFormule)
  Dim sngPrijsAbonnementPerMaand As Single
  sngPrijsAbonnementPerMaand = sngMaandelijksAbonnement
    strTekst = strTekst & "Abonnementprijs per maand: " & Format(sngPrijsAbonnementPerMaand, "0.00") & " = " & Format(sngMaandelijksAbonnement, "0.00") & loCRLF

PrijsAbonnementPerMaand = sngPrijsAbonnementPerMaand
End Function 'PrijsAbonnementPerMaand

Function PrijsKortingPerMaand(sngKortingPerMaand As Single) As Single
  Dim sngKortingPerMaand As Single
  sngKortingPerMaand = csngKortingElektronisch
  strTekst = strTekst & "Korting per maand: " & Format(sngKortingPerMaand, "0.00") & loCRLF

  PrijsKortingPerMaand = sngKortingPerMaand
End Function 'KortingPerMaand

Sub AutoDelen
  Dim intAantalUrenPerMaand As Integer
  Dim intAantalKmPerMaand As Integer
  Dim intFormule As Integer
  
  Dim sngUurprijsPerMaand As Single
  Dim sngKmprijsPerMaand As Single
  Dim sngPrijsAbonnementPerMaand As Single
  Dim sngKortingPerMaand As Single
  Dim sngTariefPerMaand As Single
  
  Dim strNaam As String
  Dim strTekst As String
  
  'init
  strTekst = ""
  Call Initialiseer
  
  'lees
  intAantalKmPerMaand = cint(val(inputbox(cstrAantalKmPerMaand)))
  intAantalUrenPerMaand = cint(val(inputbox(cstrAantalurenPerMaand)))
  
  'Verwerk
  For intFormule = cintEersteFormule To cintLaatsteFormule
 	strNaam = UCase(strFormule(intFormule))
 	
    strTekst = strTekst & strNaam & loCRLF
    
    sngUurprijsPerMaand = UurprijsPerMaand(sngUurprijs(intFormule), intAantalUrenPerMaand)
    strTekst  = strTekst & "Uurprijs per maand: " & Format(sngUurprijsPerMaand, "0.00") & " = " & Format(intAantalUrenPerMaand, "0") & " * " & Format(sngUurprijs(intFormule), "0.00") & loCRLF
    
    sngKmprijsPerMaand = KmprijsPerMaand(sngKmprijs(intFormule), intAantalKmPerMaand)
    strTekst = strTekst & "KM prijs per maand: " & Format(sngKmprijsPerMaand, "0.00") & " = " & Format(intAantalKmPerMaand, "0") & " * " & Format(sngKmprijs(intFormule), "0") & loCRLF
    
    sngPrijsAbonnementPerMaand = PrijsAbonnementPerMaand(sngMaandelijksAbonnement(intFormule))
    strTekst = strTekst & "Abonnementprijs per maand: " & Format(sngPrijsAbonnementPerMaand, "0.00") & " = " & Format(sngMaandelijksAbonnement(intFormule), "0.00") & loCRLF
    
    sngKortingPerMaand = PrijsKortingPerMaand(csngKortingElektronisch)
    strTekst = strTekst & "Korting per maand: " & Format(sngKortingPerMaand, "0.00") & loCRLF

    
'     
'     sngUurprijsPerMaand = intAantalUrenPerMaand * sngUurprijs(intFormule)
'     strTekst  = strTekst & "Uurprijs per maand: " & Format(sngUurprijsPerMaand, "0.00") & " = " & Format(intAantalUrenPerMaand, "0") & " * " & Format(sngUurprijs(intFormule), "0.00") & loCRLF
'     
'     sngKmprijsPerMaand = intAantalKmPerMaand * sngKmprijs(intFormule)
'     strTekst = strTekst & "KM prijs per maand: " & Format(sngKmprijsPerMaand, "0.00") & " = " & Format(intAantalKmPerMaand, "0") & " * " & Format(sngKmprijs(intFormule), "0") & loCRLF
'     
'     sngPrijsAbonnementPerMaand = sngMaandelijksAbonnement(intFormule)
'     strTekst = strTekst & "Abonnementprijs per maand: " & Format(sngPrijsAbonnementPerMaand, "0.00") & " = " & Format(sngMaandelijksAbonnement(intFormule), "0.00") & loCRLF
'     
'     sngTariefPerMaand = sngUurprijsPerMaand + sngKmprijsPerMaand + sngPrijsAbonnementPerMaand - sngKortingPerMaand
'     strTekst = strTekst & "Totaal per maand voor " & strNaam & ": " & Format(sngTariefPerMaand, "0.00") & " = " & Format(sngUurprijsPerMaand, "0.00") & " + " & Format(sngKmprijsPerMaand, "0.00") & " + " & Format(sngPrijsAbonnementPerMaand, "0.00") & " + " & Format(sngKortingPerMaand, "0.00") & loCRLF & loCRLF

  Next 'intFormule
  
  'Toon
'   strTekst = strTekst & "Goedkoopste: " & strFormule(Goedkoopste(sngTariefPerMaand))
'   strTekst = strTekst & "Goedkoopste optie is: " & UCase(strFormule(intGoedkoopste))
   
  Msgbox strTekst
  
End Sub 'AutoDelen
