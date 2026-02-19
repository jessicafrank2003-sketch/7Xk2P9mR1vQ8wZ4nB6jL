Attribute VB_Name = "mod_Main_Control"
Option Explicit

Sub Datenverarbeitung_Starten()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' --- RESET-ROUTINE ---
    On Error Resume Next
    Sheets(Array("L1_Normalisierte_Daten", "L2_Stratum_Nord", "L2_Stratum_Sued", "L2_Residuen", "L3_Finale_Pruefliste")).Delete
    On Error GoTo 0
    
    ' --- USERFORM ---
    Dim eingabeForm As New frmEingabe
    eingabeForm.Show
    If eingabeForm.WurdeAbgebrochen Then: Unload eingabeForm: GoTo Cleanup
    
    Dim sD As String: sD = eingabeForm.StartDatum
    Dim eD As String: eD = eingabeForm.EndDatum
    Dim analyse As Boolean: analyse = eingabeForm.AnalyseModus
    Unload eingabeForm

    ' --- LAYER 1: DATA ACCESS ---
    Dim dtCleaned As Worksheet
    Set dtCleaned = mod_Layer1_DataAccess.Layer1_PrepareData(sD, eD)
    If dtCleaned Is Nothing Then GoTo Cleanup
    
    ' --- LAYER 2: PROCESSING ---
    mod_Layer2_Processing.Layer2_ProcessSelection dtCleaned
    
    ' --- LAYER 3: PRESENTATION ---
    Dim wsN As Worksheet: Set wsN = ThisWorkbook.Sheets("L2_Stratum_Nord")
    Dim wsS As Worksheet: Set wsS = ThisWorkbook.Sheets("L2_Stratum_Sued")
    mod_Layer3_Presentation.Layer3_CreateFinalReport wsN, wsS, sD, eD
    
    ' --- OPTIONALER CLEANUP ---
    If Not analyse Then
        On Error Resume Next
        dtCleaned.Delete: wsN.Delete: wsS.Delete
        ThisWorkbook.Sheets("L2_Residuen").Delete
        On Error GoTo 0
    End If
    
    MsgBox "Stichprobenziehung erfolgreich abgeschlossen!", vbInformation

Cleanup:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
