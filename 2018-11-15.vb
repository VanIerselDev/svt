Function loCRLF() As String
    loCRLF = Chr(13) & Chr(10)
End Function 'loCRLF

Sub ToonSomVanBeginTotEinde

    Dim intTeller As Integer

    Dim intBegin As Integer
    Dim intEind As Integer
    
'init
    intTeller = 0
'lees
    intBegin = Cint(Val(InputBox("Geef een begin getal")))
    intEind = Cint(Val(InputBox("Geef een eind getal")))
'verwerk
    For intTeller = intBegin to intEind
    
    Next 'intTeller
'toon
    MsgBox strTekst
End Sub 'ToonSomVanBeginTotEinde
