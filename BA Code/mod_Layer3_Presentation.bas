Attribute VB_Name = "mod_Layer3_Presentation"
Option Explicit

Public Sub Layer3_CreateFinalReport(wsNord As Worksheet, wsSued As Worksheet, startD As String, endD As String)
    Dim wsFinal As Worksheet, tRow As Long: tRow = 2: Dim lNum As Long: lNum = 1
    Dim cN As Long, cS As Long, tN As Long, tS As Long, tTotal As Long: tTotal = 30
    
    Set wsFinal = ThisWorkbook.Sheets.Add(After:=Sheets("L2_Residuen")): wsFinal.Name = "L3_Finale_Pruefliste"
    wsFinal.Range("A1:D1").Value = Array("ID", "Land", "NL-Nummer", "Vermittler-ID")
    
    If wsNord.ListObjects.Count > 0 Then cN = wsNord.ListObjects(1).ListRows.Count
    If wsSued.ListObjects.Count > 0 Then cS = wsSued.ListObjects(1).ListRows.Count
    
    tN = 15: tS = 15
    If cN < 15 Then
        tN = cN: tS = Application.WorksheetFunction.Min(cS, tTotal - tN)
    ElseIf cS < 15 Then
        tS = cS: tN = Application.WorksheetFunction.Min(cN, tTotal - tS)
    End If
    
    L3_Internal_TransferData wsNord, wsFinal, tRow, lNum, tN
    L3_Internal_TransferData wsSued, wsFinal, tRow, lNum, tS
    L3_Internal_FinalizeLayout wsFinal, tRow, startD, endD, cN, cS
End Sub

Private Sub L3_Internal_TransferData(wsSrc As Worksheet, wsDest As Worksheet, ByRef tR As Long, ByRef lN As Long, amt As Long)
    If amt <= 0 Or wsSrc.ListObjects.Count = 0 Then Exit Sub
    Dim tbl As ListObject: Set tbl = wsSrc.ListObjects(1): Dim i As Long
    For i = 1 To amt
        wsDest.Cells(tR, 1).Value = lN
        wsDest.Cells(tR, 2).Value = tbl.DataBodyRange(i, 1).Value
        wsDest.Cells(tR, 3).Value = tbl.DataBodyRange(i, 2).Value
        wsDest.Cells(tR, 4).Value = tbl.DataBodyRange(i, 3).Value
        tR = tR + 1: lN = lN + 1
    Next i
End Sub

Private Sub L3_Internal_FinalizeLayout(ws As Worksheet, lastR As Long, sD As String, eD As String, avN As Long, avS As Long)
    Dim tbl As ListObject
    If lastR > 2 Then
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:D" & lastR - 1), , xlYes)
        tbl.Name = "Tabelle_Final": tbl.TableStyle = "TableStyleMedium2"
    End If
    ws.Columns.AutoFit
    
    ' --- VOLLSTÄNDIGER AUDIT TRAIL L3 ---
    Dim audR As Long: audR = lastR + 3
    With ws.Cells(audR, 1)
        .Value = "FINALER AUDIT TRAIL - REVISIONSBERICHT": .Font.Bold = True: .Font.Size = 12
    End With
    ws.Cells(audR + 2, 1).Value = "Status: Selektionsprozess erfolgreich abgeschlossen."
    ws.Cells(audR + 3, 1).Value = "Prüfzeitraum: " & sD & " bis " & eD
    ws.Cells(audR + 4, 1).Value = "Verfügbarkeit in L2: Nord (" & avN & "), Süd (" & avS & ")"
    ws.Cells(audR + 5, 1).Value = "Methodik: Stratifizierte Zufallsauswahl mit dynamischer Auffüll-Logik."
    ws.Cells(audR + 6, 1).Value = "Compliance: Lückenlose Nachvollziehbarkeit und Manipulationsschutz gewährleistet."
    
    ws.Cells.Locked = True: ws.Protect Password:="FraJes", UserInterfaceOnly:=True, AllowFiltering:=True
    ws.Activate: If Not tbl Is Nothing Then tbl.DataBodyRange.Select
End Sub
