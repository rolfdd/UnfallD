Sub BINAER_Export_Standard_MitIndex_OhneHeader()
    ' Deepseek "Excel Makro Umsetzen Unfalldaten zu CSV" 07-2026
    ' Header zum kopieren
    ' "Gem","yy","mm","dd","hh","Kat","Art","Typ","Licht","Zust","Rad","Pkw","Fuss","Krad","Lkw","Sonst","PLST","lon","lat"
    ' Header Unfallviewer - Kreis statt Gemeinde, Wochentag dabei, PLST fehlt
    ' UKREIS,UJAHR,UMONAT,UTAG,USTUNDE,UWOCHENTAG,UKATEGORIE,UART,UTYP,LICHT,USTRZUSTAND,IstRad,IstPKW,IstFuss,IstKrad,IstGkfz,IstSonstige,@lon,@lat
    
    Dim i As Integer, p As String, pIndex As String, r As Long, c As Long, j As Long, nextRow As Long, blatt As Variant
    Dim hasData As Boolean, firstData As Boolean, currentGemSchl As String, gemDict As Object, offsetZeilen As Long, d As Variant
    
    ' ** legt zwei leere CSV-Dateien auf dem Desktop an, wo die Daten und der Index hingeschrieben werden
    p = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\unfall.bin"
    pIndex = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\unfall_index.csv"
    
    Set gemDict = CreateObject("Scripting.Dictionary")
    offsetZeilen = 0
    firstData = (Dir(p) = "")
    
    ' äußere Schleife für Bundesländer - Sachsen ist 14
    
    For i = 1 To 16
    
        ' *****************
        ' *** Schritt 1 ***
        ' *****************
        
        ' ** temoräres Arbeitsblatt Export, wird zum Schluss wieder gelöscht
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Export"
        hasData = False
        
        ' ** innere Schleife über drei Arbeitsblätter mit dem Bundesland-Filter,
        ' ** ggf. weitere Filter (nur Rad, nur Dresden etc.)
        ' ** enthält Routinen, dass der Code weiterläuft, wenn keine Daten gefunden werden
        ' ** kopiert die gefilterten Daten in das Export-Datenblatt, die Spalten wie im Original
        For Each blatt In Array("16-17-18-19", "20-21-22-23", "24-25-26-27")
            If Application.WorksheetFunction.CountIf(Sheets(blatt).Columns(3), CStr(i)) > 0 Then
                With Sheets(blatt)
                    .UsedRange.AutoFilter Field:=3, Criteria1:=CStr(i)       ' Bundesland - Filter muss sein
                    ' .UsedRange.AutoFilter Field:=16, Criteria1:="1"        ' nur Rad - Filter optional
                    ' .UsedRange.AutoFilter Field:=32, Criteria1:="14612000" ' nur Dresden - Filter optional
            
                    ' ** Prüfen, ob sichtbare Zellen vorhanden sind für "nur Rad" etc.
                    Dim rngVisible As Range
                    On Error Resume Next
                    Set rngVisible = .UsedRange.Offset(1).Resize(.UsedRange.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
                    On Error GoTo 0
            
                    ' ** Kopieren der gefilterten Daten und anfügen in der ersten leeren Zeile in "Export" - falls Daten existieren
                    If Not rngVisible Is Nothing Then
                        nextRow = Sheets("Export").Cells(Sheets("Export").Rows.Count, "A").End(xlUp).Row + 1
                        rngVisible.Copy Destination:=Sheets("Export").Range("A" & nextRow)
                        hasData = True
                    End If
            
                    .AutoFilterMode = False
                End With
            End If
        Next blatt
        
        ' ** Routine, falls es für ein Bundesland komplett keine Daten gibt (für Testdaten o.ä.) - wird dann übersprungen,
        ' ** "Export" wird bereits hier gelöscht, die Sortierschritte etc. werden übersprungen
        If Not hasData Then
            Application.DisplayAlerts = False
            Sheets("Export").Delete
            Application.DisplayAlerts = True
            GoTo NextBundesland
        End If
        
        ' ** Sortieren der gefilterten Original-Daten (alle Spalten vorhanden und in der ursprünglichen Reeihenfolge)
        ' ** Sortierung so, dass sich ggf. ein eindeutiger Index bilden lässt Gemeindeschlüsssel + Jahr + Indexnummer
        With Sheets("Export").Sort
            .SortFields.Clear
            .SortFields.Add key:=Sheets("Export").Range("AF1"), Order:=xlAscending  ' Gemeindeschlüssel
            .SortFields.Add key:=Sheets("Export").Range("G1"), Order:=xlAscending   ' Jahr
            .SortFields.Add key:=Sheets("Export").Range("H1"), Order:=xlAscending   ' Monat
            .SortFields.Add key:=Sheets("Export").Range("AE1"), Order:=xlAscending  ' Tag
            .SortFields.Add key:=Sheets("Export").Range("I1"), Order:=xlAscending   ' Stunde
            .SortFields.Add key:=Sheets("Export").Range("X1"), Order:=xlAscending   ' X
            .SortFields.Add key:=Sheets("Export").Range("Y1"), Order:=xlAscending   ' Y
            .SortFields.Add key:=Sheets("Export").Range("K1"), Order:=xlAscending   ' Kategorie
            .SortFields.Add key:=Sheets("Export").Range("L1"), Order:=xlAscending   ' Art
            .SortFields.Add key:=Sheets("Export").Range("M1"), Order:=xlAscending   ' Typ1
            .SortFields.Add key:=Sheets("Export").Range("N1"), Order:=xlAscending   ' Lichtverhältnisse
            .SortFields.Add key:=Sheets("Export").Range("O1"), Order:=xlAscending   ' Straßenzustand
            .SortFields.Add key:=Sheets("Export").Range("P1"), Order:=xlAscending   ' Rad
            .SortFields.Add key:=Sheets("Export").Range("Q1"), Order:=xlAscending   ' PKW
            .SortFields.Add key:=Sheets("Export").Range("R1"), Order:=xlAscending   ' Fuß
            .SortFields.Add key:=Sheets("Export").Range("S1"), Order:=xlAscending   ' Krad
            .SortFields.Add key:=Sheets("Export").Range("T1"), Order:=xlAscending   ' GKfz
            .SortFields.Add key:=Sheets("Export").Range("U1"), Order:=xlAscending   ' Sonstige
            .SortFields.Add key:=Sheets("Export").Range("A1"), Order:=xlAscending   ' OBJECTID
            .SetRange Sheets("Export").UsedRange
            .Header = xlNo
            .Apply
        End With
        
        ' *****************
        ' *** Schritt 2 ***
        ' *****************
        
        ' *********************************************************************
        ' *** entweder Standard oder Unfallviewer - jeweils Auskommentieren ***
        ' *********************************************************************
        
        ' ** Variante STANDARD
        ' ** Überflüssige Spalten löschen:
        ' ** A:F   -> OBJECTID;UIDENTSTLA;ULAND;UREGBEZ;UKREIS;UGEMEINDE
        ' ** J:J   -> UWOCHENTAG
        ' ** V:W   -> LINREFX;LINREFY
        ' ** AA:AD -> IstSonstig + IstGkfz 2017;UTAG aus UIDENSTLA;Kontrolle UTAG;fake UTAG für 16/18/19
        ' ** AG:AG -> ID
        Sheets("Export").Range("A:F,J:J,V:W,AA:AD,AG:AG").Delete Shift:=xlToLeft
        
        ' ** Spalten umsortieren
        ' ** Spalte "UTAG inkl Fake für 16-18-19" verschieben rechts von "UMONAT"
        ' ** Spalte "PLST" verschieben rechts von "Sonst"
        ' ** Spalte "GemSchl errechnet" verschieben ganz an den Anfang
        Sheets("Export").Columns("R").Cut
        Sheets("Export").Columns("C").Insert Shift:=xlToRight
        Sheets("Export").Columns("R").Cut
        Sheets("Export").Columns("P").Insert Shift:=xlToRight
        Sheets("Export").Columns("S").Cut
        Sheets("Export").Columns("A").Insert Shift:=xlToRight
        Application.CutCopyMode = False
        
        
        ' ** Variante UNFALLVIEWER
        ' ** Überflüssige Spalten löschen - für Unfallviewer:
        ' ** A:D   -> OBJECTID;UIDENTSTLA;ULAND;UREGBEZ
        ' ** F:F   -> UGEMEINDE
        ' ** V:W   -> LINREFX;LINREFY
        ' ** Z:AD  -> PLST;IstSonstig + IstGkfz 2017;UTAG aus UIDENSTLA;Kontrolle UTAG;fake UTAG für 16/18/19
        ' ** AF:AG -> GemSchl errechnet;ID

        'Sheets("Export").Range("A:D,F:F,V:W,Z:AD,AF:AG").Delete Shift:=xlToLeft
    
        ' ** Spalten umsortieren
        ' ** Spalte "UTAG inkl Fake für 16-18-19" verschieben nach "UMONAT"
        'Sheets("Export").Columns("S").Cut
        'Sheets("Export").Columns("D").Insert Shift:=xlToRight
        'Application.CutCopyMode = False
        
' *****************
' *** Schritt 3 - BINÄREXPORT (16 BYTE) ***
' *****************

d = Sheets("Export").UsedRange.Value
r = UBound(d, 1)
c = UBound(d, 2)

' ** Index bauen
For j = 1 To r
    currentGemSchl = d(j, 1)
    If currentGemSchl <> "" Then
        gemDict(currentGemSchl) = j + offsetZeilen
    End If
Next j

' ** Binärdatei öffnen
Open p For Binary As #1

If Not firstData Then
    Seek #1, LOF(1) + 1
End If
firstData = False

Dim b As Byte
Dim lVal As Long
Dim koordinate As Double
Dim art As Integer, typ As Integer
Dim kat As Integer, licht As Integer, zust As Integer
Dim flags As Byte

For j = 1 To r
    ' ===================================================
    ' 1. yy -> Byte (1 Byte)
    ' ===================================================
    b = CByte(d(j, 2))
    Put #1, , b
    
    ' ===================================================
    ' 2. mm -> Byte (1 Byte)
    ' ===================================================
    b = CByte(d(j, 3))
    Put #1, , b
    
    ' ===================================================
    ' 3. dd -> Byte (1 Byte)
    ' ===================================================
    b = CByte(d(j, 4))
    Put #1, , b
    
    ' ===================================================
    ' 4. hh -> Byte (1 Byte)
    ' ===================================================
    b = CByte(d(j, 5))
    Put #1, , b
    
    ' ===================================================
    ' 5. Flags1: Art (4 Bit) + Typ (4 Bit)
    ' ===================================================
    art = CInt(d(j, 7))   ' Art: 0-9
    typ = CInt(d(j, 8))   ' Typ: 1-7
    b = (art And 15) Or ((typ And 15) * 16)  ' 16 = 2^4
    Put #1, , b
    
    ' ===================================================
    ' 6. Flags2: Kat (2 Bit) + Licht (2 Bit) + Zust (2 Bit) + Reserve (2 Bit)
    ' ===================================================
    kat = CInt(d(j, 6))   ' Kat: 1-3
    licht = CInt(d(j, 9)) ' Licht: 0-2
    zust = CInt(d(j, 10)) ' Zust: 0-3
    b = (kat And 3) Or ((licht And 3) * 4) Or ((zust And 3) * 16)  ' 4=2^2, 16=2^4
    Put #1, , b
    
    ' ===================================================
    ' 7. Flags3: Rad, Pkw, Fuss, Krad, Lkw, Sonst, PLST (7 Bit) + Reserve (1 Bit)
    ' ===================================================
    flags = 0
    If CByte(d(j, 11)) = 1 Then flags = flags Or 2 ^ 0   ' Rad
    If CByte(d(j, 12)) = 1 Then flags = flags Or 2 ^ 1   ' Pkw
    If CByte(d(j, 13)) = 1 Then flags = flags Or 2 ^ 2   ' Fuss
    If CByte(d(j, 14)) = 1 Then flags = flags Or 2 ^ 3   ' Krad
    If CByte(d(j, 15)) = 1 Then flags = flags Or 2 ^ 4   ' Lkw
    If CByte(d(j, 16)) = 1 Then flags = flags Or 2 ^ 5   ' Sonst
    If CByte(d(j, 17)) = 1 Then flags = flags Or 2 ^ 6   ' PLST
    ' Bit 7 bleibt Reserve (0)
    Put #1, , flags
    
    ' ===================================================
    ' 8. Reserve-Byte (immer 0)
    ' ===================================================
    b = 0
    Put #1, , b
    
    ' ===================================================
    ' 9. lon × 1.000.000 -> Long (4 Bytes)
    ' ===================================================
    koordinate = CDbl(d(j, 18))
    lVal = CLng(koordinate * 1000000)
    Put #1, , lVal
    
    ' ===================================================
    ' 10. lat × 1.000.000 -> Long (4 Bytes)
    ' ===================================================
    koordinate = CDbl(d(j, 19))
    lVal = CLng(koordinate * 1000000)
    Put #1, , lVal
Next j

Close #1

offsetZeilen = offsetZeilen + r
        
        Application.DisplayAlerts = False
        Sheets("Export").Delete
        Application.DisplayAlerts = True
        
NextBundesland:
    Next i
    
    ' ** Index-Datei schreiben (mit vbLf-Trennung, ohne abschließenden Zeilenumbruch)
    Open pIndex For Append As #2
    Dim key As Variant
    Dim keyCount As Long
    keyCount = gemDict.Count
    Dim currentKey As Long
    currentKey = 0

    For Each key In gemDict.Keys
        currentKey = currentKey + 1
        Print #2, key & "," & gemDict(key);
        If currentKey < keyCount Then Print #2, vbLf;
    Next key
    Close #2

    Set gemDict = Nothing
End Sub
