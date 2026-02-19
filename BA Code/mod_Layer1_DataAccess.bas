Attribute VB_Name = "mod_Layer1_DataAccess"
Option Explicit

Public Function Layer1_PrepareData(startD As String, endD As String) As Worksheet
    Dim srcWS As Worksheet: Set srcWS = ThisWorkbook.Sheets("Detail1")
    Dim hRow As Long
    If Not L1_Internal_Validate(srcWS, hRow) Then Exit Function
    Set Layer1_PrepareData = L1_Internal_ExecutePipeline(srcWS, hRow, startD, endD)
End Function

Private Function L1_Internal_Validate(ws As Worksheet, ByRef hRow As Long) As Boolean
    Dim testR As Range: Set testR = ws.Cells.Find("Country", LookAt:=xlWhole)
    If testR Is Nothing Then: L1_Internal_Validate = False: Exit Function
    hRow = testR.Row: L1_Internal_Validate = True
End Function

Private Function L1_Internal_ExecutePipeline(src As Worksheet, hRow As Long, sD As String, eD As String) As Worksheet
    Dim wsTmp As Worksheet, i As Long, targetRow As Long: targetRow = 2
    Dim dupCount As Long
    
    Set wsTmp = ThisWorkbook.Sheets.Add(After:=Sheets("Detail1")): wsTmp.Name = "L1_Normalisierte_Daten"
    wsTmp.Range("A1:C1").Value = Array("Country", "Branch", "Broker no.")
    
    Dim cCountry As Long: cCountry = src.Rows(hRow).Find("Country").Column
    Dim cDatum As Long: cDatum = src.Rows(hRow).Find("FirstContract").Column
    Dim cBranch As Long: cBranch = src.Rows(hRow).Find("Branch").Column
    Dim cBroker As Long: cBroker = src.Rows(hRow).Find("Broker no.").Column
    Dim cAct As Long: cAct = src.Rows(hRow).Find("Activity").Column
    
    For i = hRow + 1 To src.Cells(src.Rows.Count, 1).End(xlUp).Row
        If (src.Cells(i, cDatum).Text >= sD And src.Cells(i, cDatum).Text <= eD) And _
           Trim(src.Cells(i, cAct).Value) <> "D (no contract)" Then
            wsTmp.Cells(targetRow, 1).Value = src.Cells(i, cCountry).Value
            wsTmp.Cells(targetRow, 2).Value = src.Cells(i, cBranch).Value
            wsTmp.Cells(targetRow, 3).Value = src.Cells(i, cBroker).Value
            targetRow = targetRow + 1
        End If
    Next
    
    If targetRow > 2 Then
        Dim tbl As ListObject: Set tbl = wsTmp.ListObjects.Add(xlSrcRange, wsTmp.Range("A1").CurrentRegion, , xlYes)
        tbl.Name = "Tabelle_L1_Normalisiert": tbl.TableStyle = "TableStyleMedium2"
        
        Dim rowsBefore As Long: rowsBefore = tbl.DataBodyRange.Rows.Count
        wsTmp.Range("A1").CurrentRegion.RemoveDuplicates Columns:=3, Header:=xlYes
        Dim rowsAfter As Long: rowsAfter = wsTmp.Range("A1").CurrentRegion.Rows.Count - 1
        dupCount = rowsBefore - rowsAfter
        
        wsTmp.Columns.AutoFit
        
        ' --- VOLLSTÄNDIGER AUDIT TRAIL L1 ---
        Dim auditR As Long: auditR = rowsAfter + 4
        With wsTmp.Cells(auditR, 1)
            .Value = "AUDIT TRAIL - LAYER 1 (DATA ACCESS):": .Font.Bold = True: .Font.Size = 12
        End With
        wsTmp.Cells(auditR + 1, 1).Value = "1. Filterung: Zeitraum (" & sD & " bis " & eD & ") erfolgreich auf 'FirstContract' angewendet."
        wsTmp.Cells(auditR + 2, 1).Value = "2. Projektion: Spalten reduziert auf Country, Branch, Broker no."
        wsTmp.Cells(auditR + 3, 1).Value = "3. Deduplizierung: " & dupCount & " Dubletten via 'Broker no.' entfernt."
        
        wsTmp.Cells.Locked = True
        wsTmp.Protect Password:="FraJes", UserInterfaceOnly:=True, AllowFiltering:=True
        wsTmp.Activate: tbl.DataBodyRange.Select
    End If
    Set L1_Internal_ExecutePipeline = wsTmp
End Function

