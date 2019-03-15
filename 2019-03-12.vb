Const cintVastePrijsPerMaand As Integer = 3
Const csngPrijsPerKg As Single = 0.25

Sub Init

    Dim strOpdracht As String
    Dim sngSaldo As Single
    Dim intDatum(1 To 365, 0 To 0) As Integer
    Dim intAantalKgAfval As Integer 

    strOpdracht = ""
    sngSaldo = 0

End Sub 'Init

Sub AfvalVerwerking

'init
    Call Init
'lees
    strOpdracht = InputBox("Kies 'o' voor het ophalen; Kies 'p' om provisie te betalen; Kies 'r' voor uw rekening stand; Kies 'e' om te stoppen", "Wat wilt u doen")
'verwerk
    Do While strOpdracht <> "e"
        Do While sngSaldo <= 10
	    strOpdracht = "p"
	Loop 'sngSaldo > 10
	For intTeller = 1 To 365
        Select Case strOpdracht
            Case "o"
            intAantalKgAfval = Cint(Val(InputBox("Hoeveel Kg afval geeft u mee?", "Afval ophalen")))
'             intDatum(intTeller, 0) = 
            sngSaldo = sngSaldo - (cintVastePrijsPerMaand * intAantalKgAfval)
            Case "p"
            
            Case "r"
            
        Loop 'strOpdracht = "e"
    Next 'intTeller
'toon

End Sub 'AfvalVerwerking
