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
