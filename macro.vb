Sub gaanMetDieBanaan()
Dim tekst As String
Dim naam  As String
Dim eindpos As Integer
Dim wb As Workbook
Dim ws As Worksheet
Dim FoundCell As Range
Set wb = ActiveWorkbook
Set ws = Sheets(2) ''ActiveSheet
Dim regel As Integer
Dim aantal As Integer

naamKolom = "D"
infoKolom = "U"
regel = 1
aantal = 0
niksjes = 0
Do
    regel = regel + 1
    naam = ws.Range(naamKolom & regel).Value
    
    startpos = InStr(naam, ",") + 2
    voornamen = Split(Mid(naam, startpos), " ")
    
    voornaam = "(" & voornamen(0) & ")"
    naam = Mid(naam, 1, startpos - 3)
    
    
    tekst = ws.Range(infoKolom & regel).Value
    
    startpos = InStr(tekst, "Afdeling en ruimtenummer") + 3
    eindpos = InStr(tekst, "Formaat laptop")
    ruimtenr = Mid(tekst, startpos + Len("Afdeling en ruimtenummer"), eindpos - Len("Afdeling en ruimtenummer") - startpos)
    
    
    ruimtenr = Replace(ruimtenr, Chr(13), ",")
    ruimtenr = Replace(ruimtenr, Chr(12), ",")
    ruimtenr = Replace(ruimtenr, Chr(10), ",")
    ruimtenr = Replace(ruimtenr, ",,", "")
    ruimtenr = Replace(ruimtenr, " ,", ", ")
    ruimtenr = Replace(ruimtenr, " - ", ", ")
    ruimtenr = Replace(ruimtenr, ",", ", ")
    ruimtenr = Replace(ruimtenr, ",  ", ", ")
    ruimtenr = Replace(ruimtenr, ".,", ",")
    
    If (InStr(tekst, "Groot, met numeriek")) Then formaat = "Groot" Else formaat = "Klein"
    
    If (InStr(tekst, "Aktetas")) Then drager = "Aktetas"
    If (InStr(tekst, "Rugzak")) Then drager = "Rugzak"
    If (InStr(tekst, "Sleeve")) Then drager = "Sleeve"
    
    If (InStr(tekst, "- nee")) Then
        laptop = "Nee"
    Else
        startpos = InStr(tekst, "Laptop nummer (LC)") + 4
        laptop = Mid(tekst, startpos + Len("Laptop nummer (LC)"))
    End If
    
    
    eruit = False
    Set userregel = Sheets(1).Range("A:A").Find(naam, , , , , , True)
    Do
    
        
        If Not (userregel Is Nothing) Then
        
            beginstuk = Left(Sheets(1).Range("A" & userregel.Row).Value, Len(naam))
                        
            If (beginstuk = naam And InStr(Sheets(1).Range("A" & userregel.Row).Value, voornaam)) Then
                eruit = True
            Else
                Set userregel = Sheets(1).Range("A:A").FindNext(userregel)
            End If
        Else
            niksjes = niksjes + 1
        End If
        
        
    Loop While (eruit = False) '' & Not userregel Is Nothing)
    

  
    '' Vul gegevens in in cellen
    ''
    If Not userregel Is Nothing Then
    
        Sheets(1).Range("N" & userregel.Row).Value = 1                              ''Nieuwe laptop
        ''Sheets(1).Range("O" & userregel.Row).Value =  ''Laptop herinrichten
        If (laptop <> "Nee") Then
            If (InStr(Sheets(1).Range("P" & userregel.Row).Value, laptop)) Then
                
            Else
                Sheets(1).Range("P" & userregel.Row).Value = laptop & "  " & Sheets(1).Range("P" & userregel.Row).Value ''Opmerkingen
            End If
        End If
        ''Sheets(1).Range("Q" & userregel.Row).Value =  ''
        Sheets(1).Range("R" & userregel.Row).Value = formaat                        ''Laptop grootte
        Sheets(1).Range("S" & userregel.Row).Value = drager                         ''Drager
        ''Sheets(1).Range("T" & userregel.Row).Value = 1 ''
        If (formaat = "Groot") Then Sheets(1).Range("U" & userregel.Row).Value = 1  ''Laptop groot
        If (formaat = "Klein") Then Sheets(1).Range("V" & userregel.Row).Value = 1  ''Laptop Klein
        If (drager = "Aktetas") Then Sheets(1).Range("W" & userregel.Row).Value = 1 ''Aktetas
        If (drager = "Rugzak") Then Sheets(1).Range("X" & userregel.Row).Value = 1  ''Rugzak
        If (drager = "Sleeve") Then Sheets(1).Range("Y" & userregel.Row).Value = 1  ''Sleeve
        Sheets(1).Range("Z" & userregel.Row).Value = ruimtenr                       ''Ruimtenummer
        
        aantal = aantal + 1
    Else
    
        MsgBox ("userregel is Nothing, dit zou niet moeten gebeuren.  Script wordt gestopt op item nummer " & aantal & Chr(13) & "De excel is nu wel een beetje gewijzigd dus wellicht is opnieuw openen verstandig.")
        End
    End If
    
    
Loop While ws.Range(infoKolom & regel + 1).Value <> ""

MsgBox (aantal & "  regels doorgelopen" & Chr(13) & Chr(13) & niksjes & " niet kunnen vinden")
        
End Sub