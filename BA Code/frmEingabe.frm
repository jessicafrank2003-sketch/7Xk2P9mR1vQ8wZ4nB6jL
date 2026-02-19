VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEingabe 
   Caption         =   "Eingabeaufforderung"
   ClientHeight    =   5640
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   8265.001
   OleObjectBlob   =   "frmEingabe.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmEingabe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =================================================================================
' CODE FÜR DIE USERFORM "frmEingabe"
' =================================================================================

Option Explicit

' ADMINISTRATOR-EINSTELLUNG: Hier den Wert ändern (True/False)
Private Const MODUS_AN As Boolean = True

Public StartDatum As String
Public EndDatum As String
Public AnalyseModus As Boolean
Public WurdeAbgebrochen As Boolean

Private Sub UserForm_Initialize()
    Me.WurdeAbgebrochen = True
    Me.txtStartDatum.Value = "2025-10-07"
    Me.txtEndDatum.Value = "2025-10-20"
    
    ' Fixe Zuweisung des Modus ohne Nutzerinteraktion
    Me.AnalyseModus = MODUS_AN
    
    ' Visuelle Statusanzeige unten rechts
    If MODUS_AN Then
        Me.lblAnalyseStatus.Caption = "Analysemodus: AN"
        Me.lblAnalyseStatus.ForeColor = RGB(0, 150, 0) ' Grün
    Else
        Me.lblAnalyseStatus.Caption = "Analysemodus: AUS"
        Me.lblAnalyseStatus.ForeColor = RGB(150, 0, 0) ' Rot
    End If
End Sub

Private Sub cmdStarten_Click()
    If Not IsDate(Me.txtStartDatum.Value) Or Not IsDate(Me.txtEndDatum.Value) Then
        MsgBox "Bitte gültige Daten eingeben.", vbExclamation
        Exit Sub
    End If
    Me.StartDatum = Me.txtStartDatum.Value
    Me.EndDatum = Me.txtEndDatum.Value
    Me.WurdeAbgebrochen = False
    Me.Hide
End Sub

Private Sub cmdAbbrechen_Click()
    Me.WurdeAbgebrochen = True
    Me.Hide
End Sub
