VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmChoixDateImpressionFT 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportation feuille de temps"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3900
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   3900
   Begin MSComCtl2.MonthView mvwDate 
      Height          =   2370
      Left            =   1200
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   0
      Appearance      =   1
      ShowToday       =   0   'False
      StartOfWeek     =   152633345
      CurrentDate     =   37735
   End
   Begin VB.CommandButton cmdDateFin 
      Caption         =   "..."
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton cmdDateDebut 
      Caption         =   "..."
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdExporter 
      Caption         =   "Exporter"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSMask.MaskEdBox mskDateDebut 
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskDateFin 
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date début :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date fin :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "AA-MM-JJ"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "frmChoixDateImpressionFT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enumDate
 AUCUNE = 0
 DEBUT = 1
 Fin = 2
End Enum

Private m_eDate As enumDate
Private m_xlsApp As Excel.Application

Private Sub cmdAnnuler_Click()

 On Error GoTo Oups

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixDateImpressionFT", "cmdAnnuler_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdexporter_Click()

 On Error GoTo Oups

 If Len(mskDateDebut.Text) =   Then
 Call mskDateDebut_LostFocus
 End If

 If Len(mskDateFin.Text) =   Then
 Call mskDateFin_LostFocus
 End If

 If ValiderDate(mskDateDebut.Text) = False Then
 Call MsgBox("Date de début invalide!", vbOKOnly, "Erreur")

 Exit Sub
 End If

  If ValiderDate(mskDateFin.Text) = False Then
  Call MsgBox("Date de fin invalide!", vbOKOnly, "Erreur")

  Exit Sub
  End If

  If mskDateFin.Text < mskDateDebut.Text Then
  Call MsgBox("La date de fin doit être plus grande que la date de début!", vbOKOnly, "Erreur")

  Exit Sub
  End If

10 Call ExporterPunch

Call Unload(Me)

Exit Sub

Oups:

wOups "frmChoixDateImpressionFT", "cmdExporter_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Click()

 On Error GoTo Oups

 mvwDate.Visible = False

 Exit Sub

Oups:

 wOups "frmChoixDateImpressionFT", "Form_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 m_eDate = AUCUNE

 Exit Sub

Oups:

 wOups "frmChoixDateImpressionFT", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub mvwDate_LostFocus()

 On Error GoTo Oups

 mvwDate.Visible = False

 Exit Sub

Oups:

 wOups "frmChoixDateImpressionFT", "mvwDate_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mvwDate_DateClick(ByVal DateClicked As Date)

 On Error GoTo Oups

 Select Case m_eDate
 Case DEBUT: mskDateDebut.Text = ConvertDate(DateClicked)
 Case Fin: mskDateFin.Text = ConvertDate(DateClicked)
 End Select
 
 m_eDate = AUCUNE
 
 'Enlever le calendrier
 mvwDate.Visible = False

 Exit Sub

Oups:

 wOups "frmChoixDateImpressionFT", "mvwDate_DateClick", Err, Err.number, Err.Description
End Sub

Private Sub mskDateDebut_GotFocus()

 On Error GoTo Oups
 
 'Met l'année sur 2 chiffres
 If Len(mskDateDebut.Text) = 10 Then
 mskDateDebut.Text = Right$(mskDateDebut.Text, 8)
 End If
 
 mskDateDebut.mask = "##-##-##"

 Exit Sub

Oups:

 wOups "frmChoixDateImpressionFT", "mskDateDebut_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskDateFin_GotFocus()

 On Error GoTo Oups
 
 'Met l'année sur 2 chiffres
 If Len(mskDateFin.Text) = 10 Then
 mskDateFin.Text = Right$(mskDateFin.Text, 8)
 End If
 
 mskDateFin.mask = "##-##-##"

 Exit Sub

Oups:

 wOups "frmChoixDateImpressionFT", "mskDateFin_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskDateDebut_LostFocus()

 On Error GoTo Oups
 
 'Enlève le mask
 mskDateDebut.mask = vbNullString
 
 'Vide le champs si l'utilisateur n'a rien écrit
 If mskDateDebut.Text = "__-__-__" Then
 mskDateDebut.Text = vbNullString
 Else
 'Remet l'année sur   chiffres
 If Len(mskDateDebut.Text) =   Then
 If IsDate(mskDateDebut.Text) Then
 mskDateDebut.Text = Year(DateSerial(Left$(mskDateDebut.Text, 2), Mid$(mskDateDebut.Text, 4, 2), Right$(mskDateDebut.Text, 2))) & Mid$(mskDateDebut.Text, 3, 8)
 End If
 End If
 End If

  Exit Sub

Oups:

  wOups "frmChoixDateImpressionFT", "mskDateDebut_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskDateFin_LostFocus()

 On Error GoTo Oups
 
 'Enlève le mask
 mskDateFin.mask = vbNullString
 
 'Vide le champs si l'utilisateur n'a rien écrit
 If mskDateFin.Text = "__-__-__" Then
 mskDateFin.Text = vbNullString
 Else
 'Remet l'année sur   chiffres
 If Len(mskDateFin.Text) =   Then
 If IsDate(mskDateFin.Text) Then
 mskDateFin.Text = Year(DateSerial(Left$(mskDateFin.Text, 2), Mid$(mskDateFin.Text, 4, 2), Right$(mskDateFin.Text, 2))) & Mid$(mskDateFin.Text, 3, 8)
 End If
 End If
 End If

  Exit Sub

Oups:

  wOups "frmChoixDateImpressionFT", "mskDateFin_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub cmdDateDebut_Click()

 On Error GoTo Oups
 'Ouverture du calendrier
 
 'Si il y a une date valide, la date par défaut est celle-ci, sinon c'est la date
 'd'aujourd'hui
 If Trim$(mskDateDebut.Text) <> vbNullString Then
 If ValiderDate(mskDateDebut.Text) = True Then
 mvwDate.Value = mskDateDebut.Text
 Else
 mvwDate.Value = Date
 End If
 Else
 mvwDate.Value = Date
 End If
 
 m_eDate = DEBUT
 
  mvwDate.Visible = True
 
  Call mvwDate.SetFocus

  Exit Sub

Oups:

  wOups "frmChoixDateImpressionFT", "cmdDateDebut_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdDateFin_Click()

 On Error GoTo Oups
 'Ouverture du calendrier
 
 'Si il y a une date valide, la date par défaut est celle-ci, sinon c'est la date
 'd'aujourd'hui
 If Trim$(mskDateFin.Text) <> vbNullString Then
 If ValiderDate(mskDateFin.Text) = True Then
 mvwDate.Value = mskDateFin.Text
 Else
 mvwDate.Value = Date
 End If
 Else
 mvwDate.Value = Date
 End If
 
 m_eDate = Fin
 
  mvwDate.Visible = True
 
  Call mvwDate.SetFocus

  Exit Sub

Oups:

  wOups "frmChoixDateImpressionFT", "cmdDateFin_Click", Err, Err.number, Err.Description
End Sub

Private Function ValiderDate(ByVal sDate As String) As Boolean

 On Error GoTo Oups

 'Validation des dates
 If Not IsDate(sDate) Then
 ValiderDate = False
 Else
 ValiderDate = True
 End If

 Exit Function

Oups:

 wOups "frmChoixDateImpressionFT", "ValiderDate", Err, Err.number, Err.Description
End Function

Private Sub ExporterPunch()
 
 On Error GoTo Oups
 
 'Impression de la feuille de la liste des pièces
 Dim rstEmployes As ADODB.Recordset
 Dim rstProjets As ADODB.Recordset
 Dim xlsWorkBook As Excel.Workbook
 Dim iNbreEmployes As Integer
 Dim iNbreProjets As Integer
 Dim iNbrePages As Integer
 Dim iPage As Integer
 Dim sDateDebut As String
 Dim sDateFin As String
 
 Screen.MousePointer = vbHourglass

  Set rstEmployes = New ADODB.Recordset
  Set rstProjets = New ADODB.Recordset

  rstEmployes.CursorLocation = adUseClient
  rstProjets.CursorLocation = adUseClient
 
  sDateDebut = mskDateDebut.Text
  sDateFin = mskDateFin.Text
 
  Call rstProjets.Open("SELECT DISTINCT NoProjet, RIGHT(NoProjet, 9) FROM GrbPunch WHERE Date BETWEEN '" & sDateDebut & "' AND '" & sDateFin & "' ORDER BY RIGHT(NoProjet, 9)", g_connData, adOpenDynamic, adLockOptimistic)

  Call rstEmployes.Open("SELECT DISTINCT Employe FROM GrbPunch INNER JOIN GrbEmployés ON GrbEmployés.NoEmploye = GrbPunch.NoEmploye WHERE Date BETWEEN '" & sDateDebut & "' AND '" & sDateFin & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

10 iNbreEmployes = rstEmployes.RecordCount
iNbreProjets = rstProjets.RecordCount

If iNbreProjets > 2 Then
 iNbrePages = Int(iNbreProjets / 26)
 
 If iNbrePages * 2 < iNbreProjets Then
 iNbrePages = iNbrePages + 1
 End If
Else
 iNbrePages = 1
End If
 
Set m_xlsApp = New Excel.Application

Set xlsWorkBook = m_xlsApp.Workbooks.Add
 
1  For iPage = 1 To iNbrePages
 Call CreerTableau(rstProjets, rstEmployes, (iPage * 43) - 42, sDateDebut, sDateFin)
 Next
 
Call rstProjets.Close
 Call rstEmployes.Close
 
Set rstProjets = Nothing
 Set rstEmployes = Nothing
 
1  Call TransfererValeurs(sDateDebut, sDateFin)
 
 Call RemplirValeurs(iNbrePages)
 
 m_xlsApp.ActiveSheet.PageSetup.LeftMargin = m_xlsApp.Application.InchesToPoints(0)
m_xlsApp.ActiveSheet.PageSetup.RightMargin = m_xlsApp.Application.InchesToPoints(0)
m_xlsApp.ActiveSheet.PageSetup.TopMargin = m_xlsApp.Application.InchesToPoints(0)
m_xlsApp.ActiveSheet.PageSetup.BottomMargin = m_xlsApp.Application.InchesToPoints(0)
m_xlsApp.ActiveSheet.PageSetup.HeaderMargin = m_xlsApp.Application.InchesToPoints(0)
m_xlsApp.ActiveSheet.PageSetup.FooterMargin = m_xlsApp.Application.InchesToPoints(0)
m_xlsApp.ActiveSheet.PageSetup.CenterHorizontally = True
m_xlsApp.ActiveSheet.PageSetup.CenterVertically = False
m_xlsApp.ActiveSheet.PageSetup.Orientation = xlLandscape
m_xlsApp.ActiveSheet.PageSetup.PaperSize = xlPaperLegal
 
Screen.MousePointer = vbDefault

2  m_xlsApp.Visible = True

Exit Sub

Oups:

2  wOups "frmChoixDateImpressionFT", "ExporterPunch", Err, Err.number, Err.Description
End Sub

Private Sub CreerTableau(ByRef rstProjets As ADODB.Recordset, ByRef rstEmployes As ADODB.Recordset, ByVal iDebut As Integer, ByVal sDateDebut As String, ByVal sDateFin As String)
 
 On Error GoTo Oups
 
 Dim iNbreProjets As Integer
 Dim iNbreEmployes As Integer
 Dim iCompteur As Integer
 Dim sLettre As String
 
 iNbreProjets = rstProjets.RecordCount
 iNbreEmployes = rstEmployes.RecordCount
 
 m_xlsApp.Cells(iDebut, 1) = "DU " & UCase(GetDateTexte(sDateDebut)) & " AU " & UCase(GetDateTexte(sDateFin))
 m_xlsApp.range("A" & iDebut).Font.Bold = True
 m_xlsApp.range("A" & iDebut).Font.Underline = xlUnderlineStyleSingle
 m_xlsApp.range("A" & iDebut).HorizontalAlignment = xlCenter
  m_xlsApp.range("A" & iDebut).Font.SIZE = 18

  Call m_xlsApp.range("A" & iDebut, "AB" & iDebut).Merge

  m_xlsApp.Columns("A:A").ColumnWidth = 21
  m_xlsApp.Columns("B:AB").ColumnWidth = 5
  m_xlsApp.Columns("AB:AB").ColumnWidth = 6.29
 
  m_xlsApp.range("B" & iDebut + 3, "AB" & iDebut + 3).HorizontalAlignment = xlCenter
  m_xlsApp.range("B" & iDebut + 3, "AB" & iDebut + 3).VerticalAlignment = xlCenter
  m_xlsApp.range("B" & iDebut + 3, "AB" & iDebut + 3).Orientation = 90
 
10 For iCompteur = 2 To 27
1 If rstProjets.EOF = True Then
 Exit For
 Else
 m_xlsApp.Cells(iDebut + 3, iCompteur) = rstProjets.Fields("NoProjet")
 
 Call rstProjets.MoveNext
 End If
Next
 
m_xlsApp.Cells(iDebut + 3, 28) = "TOTAL"
 
Call rstEmployes.MoveFirst
 
For iCompteur = 4 To iNbreEmployes + 3
 m_xlsApp.Cells(iDebut + iCompteur, 1) = rstEmployes.Fields("Employe")
 
Call rstEmployes.MoveNext
Next
 
 m_xlsApp.Cells(iDebut + iCompteur, 1) = "TOTAL"

 'Ajout des lignes
m_xlsApp.range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeLeft).LineStyle = xlContinuous
 m_xlsApp.range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeLeft).Weight = xlMedium
m_xlsApp.range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
 
 m_xlsApp.range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeTop).LineStyle = xlContinuous
1  m_xlsApp.range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeTop).Weight = xlMedium
 m_xlsApp.range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeTop).ColorIndex = xlAutomatic
 
 m_xlsApp.range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeBottom).LineStyle = xlContinuous
m_xlsApp.range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeBottom).Weight = xlMedium
m_xlsApp.range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
 
m_xlsApp.range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeRight).LineStyle = xlContinuous
m_xlsApp.range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeRight).Weight = xlMedium
m_xlsApp.range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeRight).ColorIndex = xlAutomatic
 
m_xlsApp.range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlInsideVertical).LineStyle = xlContinuous
m_xlsApp.range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlInsideVertical).Weight = xlThin
m_xlsApp.range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlInsideVertical).ColorIndex = xlAutomatic
 
m_xlsApp.range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlInsideHorizontal).LineStyle = xlContinuous
m_xlsApp.range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlInsideHorizontal).Weight = xlThin
2  m_xlsApp.range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
 
m_xlsApp.range("B" & iDebut + 3 & ":AA" & iDebut + iNbreEmployes + 4).Borders(xlEdgeLeft).LineStyle = xlContinuous
2  m_xlsApp.range("B" & iDebut + 3 & ":AA" & iDebut + iNbreEmployes + 4).Borders(xlEdgeLeft).Weight = xlMedium
m_xlsApp.range("B" & iDebut + 3 & ":AA" & iDebut + iNbreEmployes + 4).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
 
2  m_xlsApp.range("B" & iDebut + 3 & ":AA" & iDebut + iNbreEmployes + 4).Borders(xlEdgeRight).LineStyle = xlContinuous
m_xlsApp.range("B" & iDebut + 3 & ":AA" & iDebut + iNbreEmployes + 4).Borders(xlEdgeRight).Weight = xlMedium
2  m_xlsApp.range("B" & iDebut + 3 & ":AA" & iDebut + iNbreEmployes + 4).Borders(xlEdgeRight).ColorIndex = xlAutomatic
 
m_xlsApp.range("A" & iDebut + 3 & ":AB" & iDebut + 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
30 m_xlsApp.range("A" & iDebut + 3 & ":AB" & iDebut + 3).Borders(xlEdgeBottom).Weight = xlMedium
m_xlsApp.range("A" & iDebut + 3 & ":AB" & iDebut + 3).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
 
m_xlsApp.range("A" & iDebut + iNbreEmployes + 4 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeTop).LineStyle = xlContinuous
m_xlsApp.range("A" & iDebut + iNbreEmployes + 4 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeTop).Weight = xlMedium
m_xlsApp.range("A" & iDebut + iNbreEmployes + 4 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeTop).ColorIndex = xlAutomatic
 
For iCompteur = iDebut + 4 To iNbreEmployes + iDebut + 4
 m_xlsApp.range("AB" & iCompteur) = "=SUM(B" & iCompteur & ":AA" & iCompteur & ")"
Next
 
For iCompteur = 2 To 27
 Select Case iCompteur
 Case 2: sLettre = "B"
 Case 3: sLettre = "C"
 Case 4: sLettre = "D"
 Case 5: sLettre = "E"
 Case 6: sLettre = "F"
 Case 7: sLettre = "G"
 Case 8: sLettre = "H"
 Case 9: sLettre = "I"
 Case 10: sLettre = "J"
 Case 11: sLettre = "K"
 Case 12: sLettre = "L"
 Case 13: sLettre = "M"
 Case 14: sLettre = "N"
 Case 15: sLettre = "O"
 Case 16: sLettre = "P"
 Case 17: sLettre = "Q"
 Case 18: sLettre = "R"
 Case 19: sLettre = "S"
 Case 20: sLettre = "T"
 Case 21: sLettre = "U"
 Case 22: sLettre = "V"
 Case 23: sLettre = "W"
 Case 24: sLettre = "X"
 Case 25: sLettre = "Y"
 Case 26: sLettre = "Z"
 Case 27: sLettre = "AA"
 End Select
 
 m_xlsApp.range(sLettre & iDebut + iNbreEmployes + 4) = "=SUM(" & sLettre & (iDebut + 4) & ":" & sLettre & (iDebut + iNbreEmployes + 3) & ")"
Next

3  Exit Sub

Oups:

wOups "frmChoixDateImpressionFT", "CreerTableau", Err, Err.number, Err.Description
End Sub

Private Sub TransfererValeurs(ByVal sDateDebut As String, ByVal sDateFin As String)
 
 On Error GoTo Oups

 Dim rstSource As ADODB.Recordset
 Dim rstDestination As ADODB.Recordset
 Dim sDate As String
 Dim sHeure As String
 Dim sMinute As String
 Dim dblResult As Double
 Dim datTemp As Date
 
 Set rstSource = New ADODB.Recordset
 
 Call rstSource.Open("SELECT * FROM GrbPunch WHERE Date BETWEEN '" & sDateDebut & "' AND '" & sDateFin & "' AND HeureFin Is Not Null", g_connData, adOpenForwardOnly, adLockReadOnly)
 
 If Not rstSource.EOF Then
  Call g_connData.Execute("DELETE * FROM GrbPunchExcel")
 
  Set rstDestination = New ADODB.Recordset
 
  Call rstDestination.Open("SELECT * FROM GrbPunchExcel", g_connData, adOpenDynamic, adLockOptimistic)
 
  Do While Not rstSource.EOF
  Call rstDestination.AddNew
 
  rstDestination.Fields("NoEmploye") = rstSource.Fields("NoEmploye")
  rstDestination.Fields("NoProjet") = rstSource.Fields("NoProjet")
 
  sHeure = Left(rstSource.Fields("HeureDébut"), 2)
 sMinute = Right(rstSource.Fields("HeureDébut"), 2)
 
dblResult = CInt(sMinute) / 60
 
 If dblResult <> 0 Then
 rstDestination.Fields("HeureDébut") = sHeure & "," & Right$(dblResult, Len(CStr(dblResult)) - InStr(1, dblResult, ","))
 Else
 rstDestination.Fields("HeureDébut") = CInt(sHeure)
 End If
 
 sHeure = Left(rstSource.Fields("HeureFin"), 2)
 sMinute = Right(rstSource.Fields("HeureFin"), 2)
 
 dblResult = CInt(sMinute) / 60
 
 If dblResult <> 0 Then
 rstDestination.Fields("HeureFin") = CInt(sHeure) & "," & CInt(Right$(dblResult, Len(CStr(dblResult)) - InStr(1, dblResult, ",")))
 Else
 rstDestination.Fields("HeureFin") = sHeure
 End If
 
 Call rstDestination.Update
 
 Call rstSource.MoveNext
 Loop
 
 Call rstDestination.Close
1  Set rstDestination = Nothing
 End If
 
 Call rstSource.Close
Set rstSource = Nothing

Exit Sub

Oups:

wOups "frmChoixDateImpressionFT", "TransfererValeurs", Err, Err.number, Err.Description
End Sub

Private Sub RemplirValeurs(ByVal iNbrePages As Integer)
 
 On Error GoTo Oups

 Dim rstPunch As ADODB.Recordset
 Dim sNom As String
 Dim iIndexNom As Integer
 Dim iCompteur As Integer
 Dim iPage As Integer
 Dim iPageRendu As Integer
 Dim iNoRendu As Integer
 Dim iDebutPage As Integer
 Dim iIndexProjet As Integer
 Dim bProjetTrouve As Boolean

  Set rstPunch = New ADODB.Recordset

  Call rstPunch.Open("SELECT Employe, NoProjet, SUM(HeureFin - HeureDébut) As Total FROM GrbPunchExcel INNER JOIN GrbEmployés ON GrbPunchExcel.NoEmploye = GrbEmployés.NoEmploye GROUP BY Employe, NoProjet ORDER BY Employe, RIGHT(NoProjet, 9)", g_connData, adOpenForwardOnly, adLockReadOnly)
 
  If Not rstPunch.EOF Then
  iIndexNom = 4
 
  sNom = rstPunch.Fields("Employe")
 
  iPageRendu = 1
  iNoRendu = 2
 
  Do While Not rstPunch.EOF
 If sNom <> rstPunch.Fields("Employe") Then
 sNom = rstPunch.Fields("Employe")
 
 iIndexNom = iIndexNom + 1
 
 iPageRendu = 1
 iNoRendu = 2
 Else
 If iNoRendu > 2 Then
 iNoRendu = iNoRendu - 1
 Else
 If iPageRendu > 1 Then
 iPageRendu = iPageRendu - 1
 iNoRendu = 27
 End If
 End If
 End If
 
 bProjetTrouve = False
 
 For iPage = iPageRendu To iNbrePages
 If iPageRendu <> iPage Then
 iNoRendu = 2
1  End If

 iDebutPage = (iPage * 43) - 42
 
 For iCompteur = iNoRendu To 27
 If m_xlsApp.Cells(iDebutPage + 3, iCompteur) = rstPunch.Fields("NoProjet") Then
 iIndexProjet = iCompteur
 
 iPageRendu = iPage
 iNoRendu = iCompteur
 
 bProjetTrouve = True
 
 Exit For
 End If
 Next
 
 If bProjetTrouve = True Then
 Exit For
 End If
 Next
 
 If bProjetTrouve = False Then
 Call MsgBox("Le # " & rstPunch.Fields("NoProjet") & " n'a pas pu être trouvé pour l'employé " & sNom & "." & vbNewLine & _
 "Son temps de " & rstPunch.Fields("Total") & " heures sera ajouté à cet endroit : " & vbNewLine & _
 "Page : " & iPageRendu & vbNewLine & _
 "Rangée : " & iIndexProjet)
 
 End If
 
 m_xlsApp.Cells(iDebutPage + iIndexNom, iIndexProjet) = rstPunch.Fields("Total")
 
 Call rstPunch.MoveNext
 Loop
30 End If
 
Call rstPunch.Close
Set rstPunch = Nothing

Exit Sub

Oups:

wOups "frmChoixDateImpressionFT", "RemplirValeurs", Err, Err.number, Err.Description
End Sub
