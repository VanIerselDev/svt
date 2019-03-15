Sub lees10getallenzoekkleinste

Const cintAantalGetallen = 10
Dim intGetallen(1 to 10) As Integer
Dim intPositieKleinste As Integer

Dim intTeller As Integer

Dim strTekst As String

'init
strTekst = "Het kleinste getal van uw reeks is: "

'lees
For intPositie = 1 to cintAantalGetallen

    intGetallen(intPositie) = Cint(Val(InputBox("Geef het getal van de positie: " & Format(intpositie))))

Next 'intPositie

'verwerk
intPositieKleinste = 1

For intTeller = 2 to cintAantalGetallen

    If intGetallen(intTeller) < intGetallen(intPositieKleinste) Then
        
        intPositieKleinste = intTeller
        
    End If 'intGetallen(intTeller) > intGetallen(intPositieKleinste)

Next 'intTeller
strTekst = strTekst & Cint(Val(intGetallen(intPositieKleinste))) & " bevindt zich op positie: " & Cint(Val(intPositieKleinste))

'show

MsgBox strTekst
End Sub 'lees10getallenzoekkleinste

Sub lees10getallenzoekgrootste


Const cintAantalGetallen = 10
Dim intGetallen(1 to 10) As Integer
Dim intPositieGrootste As Integer

Dim intTeller As Integer

Dim strTekst As String

'init
strTekst = "Het grootste getal van uw reeks is: "

'lees
For intPositie = 1 to cintAantalGetallen

    intGetallen(intPositie) = Cint(Val(InputBox("Geef het getal van de positie: " & Format(intpositie))))

Next 'intPositie

'verwerk
intPositieGrootste = 1

For intTeller = 2 to cintAantalGetallen

    If intGetallen(intTeller) > intGetallen(intPositieGrootste) Then
        
        intPositieGrootste = intTeller
        
    End If 'intGetallen(intTeller) < intGetallen(intPositieGrootste)

Next 'intTeller
strTekst = strTekst & Cint(Val(intGetallen(intPositieGrootste))) & " bevindt zich op positie: " & Cint(Val(intPositieGrootste))

'show

MsgBox strTekst
End Sub 'lees10getallenzoekgrootste

Sub Lees10GetallenInEenMatrixEnZoekHetKleinsteMetBeginPositie
  Const cintAantalGetallen As integer = 10

  Dim intPositieKleinste As integer
  Dim intTeller As integer
  Dim intGetal(1 to cintAantalGetallen) As integer
  Dim intBeginPositie As Integer
  Dim strTekst As String
'init
  intGetal(1) = 10
  intGetal(2) = 1
  intGetal(3) = 9
  intGetal(4) = 2
  intGetal(5) = 8
  intGetal(6) = 3
  intGetal(7) = 7
  intGetal(8) = 4
  intGetal(9) = 6
  intGetal(10) = 5
  strTekst = ""
'lees
  intBeginPositie = Cint(Val(InputBox("Geef de begin positie van de matrix")))
'verwerk  
  intPositieKleinste = intBeginPositie
  For intTeller = intBeginPositie to cintAantalGetallen
    If intGetal(intTeller) < intGetal(intPositieKleinste) Then
      intPositieKleinste = intTeller
    End If 'intGetal(intTeller) < intGetal(intPositieKleinste)  
  Next 'intTeller
'toon
  For intTeller = intBeginPositie to cintAantalGetallen
    strTekst = strTekst & Format(intGetal(intTeller), "0") & " "
  Next 'intTeller
  strTekst = strTekst & loCRLF & "op positie " & format(intPositieKleinste, "0") & " staat " & format(intGetal(intPositieKleinste), "0") & "."
  MsgBox strTekst
End Sub 'Lees10GetallenInEenMatrixEnZoekHetKleinste

Sub Sorteer
  Const cintAantalGetallen As integer = 10

  Dim intPositieKleinste As integer
  Dim intTeller As integer
  Dim intGetallen(1 to cintAantalGetallen) As integer
  Dim intWissel As Integer
  Dim intBeginPositie As Integer
  Dim strTekst As String
'init
  intGetallen(1) = 10
  intGetallen(2) = 1
  intGetallen(3) = 9
  intGetallen(4) = 2
  intGetallen(5) = 8
  intGetallen(6) = 3
  intGetallen(7) = 7
  intGetallen(8) = 4
  intGetallen(9) = 6
  intGetallen(10) = 5
  strTekst = ""
'lees
  intBeginPositie = 1
'verwerk  
  intPositieKleinste = intBeginPositie
	For intTeller = cintAantalGetallen - 1 To intBeginPositie Step -1
		For intTeller2 = 1 To intTeller  
            If intGetallen(intTeller2) > intGetallen(intTeller2 + 1) Then
                intWissel = intGetallen(intTeller2)  
                intGetallen(intTeller2) = intGetallen(intTeller2 + 1)
                intGetallen(intTeller2 + 1) = intWissel  
  
            End If  
        Next  
    Next  
'toon
  For intTeller = intBeginPositie To cintAantalGetallen
    strTekst = strTekst & Format(intGetallen(intTeller), "0") & " "
  Next 'intTeller
  MsgBox strTekst
End Sub 'Sorteer

Sub Sorteer2
  Const cintAantalGetallen As integer = 10

  Dim intPositieKleinste As integer
  Dim intTeller As integer
  Dim intGetallen(1 to cintAantalGetallen) As integer
  Dim intWissel As Integer
  Dim intBeginPositie As Integer
  Dim strTekst As String
'init
  intGetallen(1) = 10
  intGetallen(2) = 1
  intGetallen(3) = 9
  intGetallen(4) = 2
  intGetallen(5) = 8
  intGetallen(6) = 3
  intGetallen(7) = 7
  intGetallen(8) = 4
  intGetallen(9) = 6
  intGetallen(10) = 5
  strTekst = ""
'lees
  intBeginPositie = 1
'verwerk  
  intPositieKleinste = intBeginPositie
	For intTeller = cintAantalGetallen - 1 To intBeginPositie Step -1
		For intTeller2 = 1 To intTeller  
            If intGetallen(intTeller2) < intGetallen(intTeller2 + 1) Then
                intWissel = intGetallen(intTeller2)  
                intGetallen(intTeller2) = intGetallen(intTeller2 + 1)
                intGetallen(intTeller2 + 1) = intWissel  
            End If  
        Next  
    Next  
'toon
  For intTeller = intBeginPositie To cintAantalGetallen
    strTekst = strTekst & Format(intGetallen(intTeller), "0") & " "
  Next 'intTeller
  MsgBox strTekst
End Sub 'Sorteer2
