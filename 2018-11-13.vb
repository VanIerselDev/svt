Sub Nigr
'Initialize constants
const intBeginGetal As Integer = 0
const intEindGetal As Integer = 10000
'Initialize Variables
Dim intStep As Integer 
Dim intSom As Integer
Dim intGetal As Integer
Dim loCRLF As String
Dim strTekst As String

intStep = 1
intSom = 0

loCRLF = Chr(13) & Chr(10)
strTekst = ""
'Exec
For intGetal = intBeginGetal = to intEindGetal
    intSom = intSom + intGetal
    strTekst = strTekst & Format(intSom, "0")
    

End Sub 'Nigr
