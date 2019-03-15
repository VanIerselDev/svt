Const cintEersteGetalPositieMatrix As Integer = 1
Const cintLaatsteGetalPositieMatrix As Integer = 1000
Const cintAantalGetallen As Integer = (cintLaatsteGetalPositieMatrix + 1) - cintEersteGetalPositieMatrix

Sub Matrix1000

Dim intMatrix(cintEersteGetalPositieMatrix To cintLaatsteGetalPositieMatrix) As Integer

Dim intTellerLeesGetallen As Integer 'intTeller voor de getallen in te lezen
Dim intTeller As Integer 'intTeller voor de kleinste, grootste en op volgorde te zetten
Dim intTeller2 As Integer
Dim intTeller3 As Integer

Dim intTellerToonSorteer As Integer

Dim intBeginPositie As Integer
Dim intPositieGrootste2 As Integer

Dim intPositieKleinste As Integer
Dim intPositieGrootste As Integer

Dim intPositieBeginSorteer As Integer
Dim intWissel As Integer

Dim strTekst As String

'init
intBeginPositie = cintEersteGetalPositieMatrix

strTekst = ""

' Voor testen 
' intMatrix(1) = 4
' intMatrix(2) = 1
' intMatrix(3) = 3
' intMatrix(4) = 5
' intMatrix(5) = 8
' intMatrix(6) = 9
' intMatrix(7) = 2
' intMatrix(8) = 7
' intMatrix(9) = 10
' intMatrix(10) = 6
'lees

For intTellerLeesGetallen = 1 to cintAantalGetallen
 
     intMatrix(intTellerLeesGetallen) = Cint(Val(InputBox("Geef het getal van de positie: " & Format(intTellerLeesGetallen) & " van de " & Format(cintAantalGetallen, "0") & " getallen.")))
 
Next 'intTellerLeesGetallen

intPositieBeginSorteer = Cint(Val(InputBox("Geef de positie vanaf waar u de matrix van " & Format(cintAantalGetallen, "0") & " getallen gesorteerd wilt hebben (moet tussen:" & Format(cintEersteGetalPositieMatrix, "0") & " en " & Format(cintLaatsteGetalPositieMatrix, "0") & " zijn)")))

'verwerk
intPositieKleinste = 1
intPositieGrootste = 1

For intTeller = (cintEersteGetalPositieMatrix + 1) to cintAantalGetallen
    'Zoek het kleinste getal
    If intMatrix(intTeller) < intMatrix(intPositieKleinste) Then
        
        intPositieKleinste = intTeller
        
    End If 'intGetallen(intTeller) > intGetallen(intPositieKleinste)
Next 'intTeller

For intTeller = (cintEersteGetalPositieMatrix + 1) to cintAantalGetallen
     'Zoek het grootste getal
    If intMatrix(intTeller) > intMatrix(intPositieGrootste) Then
        
        intPositieGrootste = intTeller
        
    End If 'intGetallen(intTeller) < intGetallen(intPositieGrootste)
Next 'intTeller

 For intTeller2 = intPositieBeginSorteer To cintAantalGetallen -1
    intPositieGrootste2 = intTeller2
    For intTeller3 = intTeller2 + 1 To cintAantalGetallen 
        If intMatrix(intTeller3) < intMatrix(intPositieGrootste2) then
            intPositieGrootste2 = intTeller3
        End If 'intMatrix(intTeller3) <= intMatrix(intPositieGrootste2)
    Next 'intTeller3
        If intTeller2 <> intPositieGrootste2 Then
            intWissel = intMatrix(intPositieGrootste2)
            intMatrix(intPositieGrootste2) = intMatrix(intTeller2)
            intMatrix(intTeller2) = intWissel  
        End If 'intMatrix(intTeller3) <= intMatrix(intPositieGrootste2)
 Next 'intTeller2

'toon
'Pas strTekst aan met tekst over kleinste getal
strTekst = strTekst & "Het kleinste getal van uw reeks is: " & Format(intMatrix(intPositieKleinste), "0") & " en bevindt zich op positie: " & Format(intPositieKleinste, "0") & loCRLF

'Pas strTekst aan met tekst over grootste getal
strTekst = strTekst & "Het grootste getal van uw reeks is: " & Format(intMatrix(intPositieGrootste), "0") & " en bevindt zich op positie: " & Format(intPositieGrootste, "0") & loCRLF

For intTellerToonSorteer = intBeginPositie To cintAantalGetallen
    strTekst = strTekst & Format(intMatrix(intTellerToonSorteer), "0") & " "
  Next 'intTeller

MsgBox strTekst

End Sub 'Matrix1000
