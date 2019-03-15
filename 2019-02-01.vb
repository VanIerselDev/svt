REM  *****  BASIC  *****
Option Explicit

Function CRLF() As String
	CRLF = Chr(13) & Chr(10)
End Function 'CRLF

Sub LeesWoorden

	Dim intMaxAantalKarakters As Integer
	Dim intAantalKarakters As Integer
	Dim intTeller As Integer
	Dim strWoord As String
	Dim strTekst As String
	Dim strMsgTekst As String
	
	'init
	strWoord = ""
    strTekst = ""
    strMsgTekst = ""
    
    intAantalKarakters = 0
    
    'lees
    intMaxAantalKarakters = Cint(Val(Inputbox("Aantal Karakters om te stoppen?")))
    
    'bereken
    Do Until intMaxAantalKarakters <= intAantalKarakters
        strWoord = Inputbox("Geef een woord")
        For intTeller = 1 To Len(strWoord)
        	Select Case Mid(strWoord, intTeller)
        		Case " "
        			
        		Case 1 To 9
        			
        	End Select 'Mid(strWoord, intTeller)
        Next 'intTeller
        intAantalKarakters = intAantalKarakters + Len(strWoord)
    Loop 'intMaxAantalKarakters <= intAantalKarakters
    strMsgTekst = strTekst
    
    'toon
    MsgBox strMsgTekst
	
End Sub 'LeesWoorden
