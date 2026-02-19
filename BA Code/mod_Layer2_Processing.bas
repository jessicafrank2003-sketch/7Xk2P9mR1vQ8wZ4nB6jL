Attribute VB_Name = "mod_Layer2_Processing"
' =================================================================================
' MODUL: mod_Layer2_Processing
' LAYER: 2 (Core Processing Layer)
' ZWECK: Klassifizierung, Selektion & Fokus auf Datenkörper
' =================================================================================
Option Explicit

Public Sub Layer2_ProcessSelection(wsL1 As Worksheet)
    Dim wsNord As Worksheet, wsSued As Worksheet, wsResiduen As Worksheet
    L2_Internal_SplitByRegion wsL1, wsNord, wsSued, wsResiduen
    If Not wsNord Is Nothing Then L2_Internal_ApplySampling wsNord, "Nord", True
    If Not wsSued Is Nothing Then L2_Internal_ApplySampling wsSued, "Sued", True
    If Not wsResiduen Is Nothing Then L2_Internal_FormatAsTable wsResiduen, "Residuen"
End Sub

Private Sub L2_Internal_SplitByRegion(wsSource As Worksheet, ByRef wsN As Worksheet, ByRef wsS As Worksheet, ByRef wsR As Worksheet)
    Dim tbl As ListObject: Set tbl = wsSource.ListObjects(1)
    Dim nSt As Variant, sSt As Variant, i As Long, brTxt As String, reg As String, st As Variant
    
    nSt = Array("Hamburg", "Hannover", "Bremen", "Kiel", "Rostock", "Berlin", "Potsdam", "Magdeburg", "Dresden", "Leipzig", "Düsseldorf", "Mönchengladbach", "Dortmund", "Bielefeld", "Köln")
    sSt = Array("Frankfurt", "Mannheim", "Saarbrücken", "Kassel", "Stuttgart", "Baden-Baden", "Freiburg", "Neu-Ulm", "Heilbronn", "Ravensburg", "München", "Nürnberg", "Regensburg", "Erfurt", "Augsburg", "Würzburg")
    
    Set wsN = ThisWorkbook.Sheets.Add(After:=Sheets("L1_Normalisierte_Daten")): wsN.Name = "L2_Stratum_Nord"
    Set wsS = ThisWorkbook.Sheets.Add(After:=wsN): wsS.Name = "L2_Stratum_Sued"
    Set wsR = ThisWorkbook.Sheets.Add(After:=wsS): wsR.Name = "L2_Residuen"
    
    wsSource.Rows(1).Copy wsN.Rows(1): wsSource.Rows(1).Copy wsS.Rows(1): wsSource.Rows(1).Copy wsR.Rows(1)
    
    Dim nR As Long: nR = 2: Dim sR As Long: sR = 2: Dim rR As Long: rR = 2
    For i = 1 To tbl.ListRows.Count
        brTxt = tbl.DataBodyRange(i, 2).Value: reg = "Unbekannt"
        
        For Each st In nSt
            If InStr(1, brTxt, st, vbTextCompare) > 0 Then: reg = "Nord": Exit For
        Next
        
        If reg = "Unbekannt" Then
            For Each st In sSt
                If InStr(1, brTxt, st, vbTextCompare) > 0 Then
                    reg = "Süd"
                    Exit For
                End If
            Next
        End If
        
        If reg = "Nord" Then
            L2_Internal_CopyClean tbl.ListRows(i).Range, wsN.Rows(nR), True: nR = nR + 1
        ElseIf reg = "Süd" Then
            L2_Internal_CopyClean tbl.ListRows(i).Range, wsS.Rows(sR), True: sR = sR + 1
        Else
            L2_Internal_CopyClean tbl.ListRows(i).Range, wsR.Rows(rR), False: rR = rR + 1
        End If
    Next i
End Sub

Private Sub L2_Internal_CopyClean(src As Range, dest As Range, clean As Boolean)
    src.Copy: dest.PasteSpecial xlPasteValues
    If clean Then
        Dim s As String: s = dest.Cells(1, 2).Value: Dim res As String: res = "": Dim j As Long
        For j = 1 To Len(s): If IsNumeric(Mid(s, j, 1)) Then res = res & Mid(s, j, 1)
        Next j: dest.Cells(1, 2).Value = res
    End If
End Sub

Private Sub L2_Internal_ApplySampling(ws As Worksheet, regionName As String, clean As Boolean)
    If ws.Cells(2, 1).Value = "" Then Exit Sub
    Dim tbl As ListObject: Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
    tbl.Name = "Tabelle_L2_" & Replace(regionName, "ü", "ue"): tbl.TableStyle = "TableStyleMedium2"
    
    Dim colZ As ListColumn: Set colZ = tbl.ListColumns.Add: colZ.Name = "Zufallswert"
    colZ.DataBodyRange.FormulaLocal = "=ZUFALLSZAHL()": colZ.DataBodyRange.Value = colZ.DataBodyRange.Value
    
    With tbl.Sort: .SortFields.Clear: .SortFields.Add Key:=colZ.Range, Order:=xlAscending: .Apply: End With
    ws.Columns.AutoFit
    
    ' --- AKTUALISIERTER AUDIT TRAIL L2 ---
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    With ws.Cells(lastR + 4, 1)
        .Value = "AUDIT TRAIL - LAYER 2 (PROCESSING):": .Font.Bold = True: .Font.Size = 12
    End With
    ws.Cells(lastR + 5, 1).Value = "1. Stratifizierung: Datensatz dem Stratum '" & regionName & "' zugeordnet."
    ws.Cells(lastR + 6, 1).Value = "2. Selektion: Ermittlung der Grundgesamtheit pro Stratum für die Stichprobenziehung."
    If clean Then ws.Cells(lastR + 7, 1).Value = "3. Bereinigung: Städtenamen entfernt, nur NL-Nummern beibehalten."
    ws.Cells(lastR + 8, 1).Value = "4. Selektion: Fixierte Zufallswerte aufsteigend sortiert."
    
    ' MANIPULATIONSSCHUTZ & FOKUS
    ws.Cells.Locked = True: ws.Protect Password:="FraJes", UserInterfaceOnly:=True, AllowFiltering:=True
    ws.Activate: tbl.DataBodyRange.Select
End Sub

Private Sub L2_Internal_FormatAsTable(ws As Worksheet, tblName As String)
    If ws.Cells(2, 1).Value = "" Then Exit Sub
    Dim tbl As ListObject: Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
    tbl.Name = "Tabelle_L2_" & tblName: tbl.TableStyle = "TableStyleMedium2": ws.Columns.AutoFit
    
    ' --- VOLLSTÄNDIGER AUDIT TRAIL RESIDUEN ---
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    With ws.Cells(lastR + 3, 1)
        .Value = "TRANSPARENZ-CHECK (RESIDUEN):": .Font.Bold = True: .Font.Size = 12
    End With
    ws.Cells(lastR + 4, 1).Value = "Diese Datensätze wurden keiner Region zugeordnet. Städtenamen zur Fehleranalyse im Original belassen."
    
    ws.Cells.Locked = True: ws.Protect Password:="FraJes", UserInterfaceOnly:=True, AllowFiltering:=True
    ws.Activate: tbl.DataBodyRange.Select
End Sub
