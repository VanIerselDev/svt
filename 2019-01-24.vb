Option Explicit

'Extra functies
Function loCRLF() As String
	
	loCRLF = Chr(13) & Chr(10)
	
End Function 'loCRLF


'Begin programma functies
Function KleineLetters(strTekst As String) As String
    Dim strKleineLettersVL As String
    
    strKleineLettersVL = LCase(strTekst)
    
    KleineLetters = strKleineLettersVL
End Function 'KleineLetters

Function GroteLetters(strTekst As String) As String
    Dim strGroteLettersVL As String
    
    strGroteLettersVL = UCase(strTekst)
    
    GroteLetters = strGroteLettersVL
End Function 'KleineLetters

Function GeenSpaties(strTekst As String) As String
    Dim strGeenSpatiesVL As String
    
    Dim intTeller As Integer
    
    For intTeller = 1 To Len(strTekst)
    	strGeenSpatiesVL = strTekst.Replace(" ", "")
    Next 'intTeller
        
    GeenSpaties = strGeenSpatiesVL
End Function 'KleineLetters


'Begin programma
Sub Woordspelletje    

	Const cstrError As String = "[ERROR]: Voer iets in!"
    
    Dim strUitkomst As String
	Dim strInvoer As String
    
    Dim intKeuze As Integer
    
    'init
    strUitkomst = ""
    'lees
    strInvoer = InputBox("Uw tekst om te bewerken.")
    intKeuze = Cint(Val(InputBox("1. Kleine Letters; 2. Grote Letters; 3. Geen Spaties")
    'bereken
    Select Case intKeuze
    	Case 1
 			strUitkomst = strUitkomst & strInvoer & " is in kleine letters: " & KleineLetters(strInvoer) & loCRLF()
 		Case 2
			strUitkomst = strUitkomst & strInvoer & " is in grote letters: " & GroteLetters(strInvoer) & loCRLF()
		Case 3
			strUitkomst = strUitkomst & strInvoer & " is zonder spaties: " & GeenSpaties(strInvoer) & loCRLF()
		Case Else
			strUitkomst = cstrError
    'toon
    Msgbox strUitkomst

End Sub 'Woordspelletje
