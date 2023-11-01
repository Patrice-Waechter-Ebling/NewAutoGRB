VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChoixDateSommairePunch 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sommaire des projets"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   7800
   Begin MSComCtl2.MonthView mvwDate 
      Height          =   2370
      Left            =   5040
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   0
      Appearance      =   1
      ShowToday       =   0   'False
      StartOfWeek     =   152305665
      CurrentDate     =   37735
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   4800
      TabIndex        =   6
      Top             =   960
      Width           =   2895
      Begin VB.CommandButton cmdDateFin 
         Caption         =   "..."
         Height          =   255
         Left            =   2280
         TabIndex        =   10
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton cmdDateDebut 
         Caption         =   "..."
         Height          =   255
         Left            =   2280
         TabIndex        =   7
         Top             =   600
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskDateDebut 
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDateFin 
         Height          =   255
         Left            =   960
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Début :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fin :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "AA-MM-JJ"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fraFamille 
      BackColor       =   &H00000000&
      Caption         =   "Famille"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   2400
      TabIndex        =   5
      Top             =   960
      Width           =   2295
      Begin MSComctlLib.ListView lvwFamille 
         Height          =   1575
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   2778
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Famille"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Projets"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2175
      Begin VB.OptionButton optProjets 
         BackColor       =   &H00000000&
         Caption         =   "Sommaire des heures"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   2025
      End
      Begin VB.OptionButton optProjets 
         BackColor       =   &H00000000&
         Caption         =   "Projets GRB seulement"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   2025
      End
      Begin VB.OptionButton optProjets 
         BackColor       =   &H00000000&
         Caption         =   "Tous les projets"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   2025
      End
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimer 
      Caption         =   "Imprimer"
      Default         =   -1  'True
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
   End
End
Attribute VB_Name = "frmChoixDateSommairePunch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const I_OPT_TOUS_LES_PROJETS As Integer = 0
Private Const I_OPT_PROJETS_GRB As Integer = 1
Private Const I_OPT_SOMMAIRE_HEURES As Integer = 2


Private Enum enumDate
 AUCUNE = 0
 DEBUT = 1
 Fin = 2
End Enum

Private m_eDate As enumDate

Private Sub cmdAnnuler_Click()

 On Error GoTo Oups

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixDateSommairePunch", "cmdAnnuler_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdImprimer_Click()

 On Error GoTo Oups

 Dim iCompteur As Integer
 Dim bChecked As Boolean

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

 'Si "Projets GRB seulement" est choisi
  If optProjets(I_OPT_PROJETS_GRB).Value = True Then
 'Si aucune famille est sélectionnée
  For iCompteur = 1 To lvwFamille.ListItems.count
  If lvwFamille.ListItems(iCompteur).Checked = True Then
  bChecked = True

 Exit For
End If
 Next
 
 If bChecked = False Then
 Call MsgBox("Vous devez choisir au moins une famille d'employés!", vbOKOnly, "Erreur")
 
 Exit Sub
 End If
End If

Screen.MousePointer = vbHourglass

If optProjets(I_OPT_TOUS_LES_PROJETS).Value = True Then
 Call ImprimerPunchGeneral
Else
If optProjets(I_OPT_PROJETS_GRB).Value = True Then
 Call ImprimerPunchGRB
 Else
 Call ImprimerSommaireHeures
 End If
End If

 Screen.MousePointer = vbDefault

1  Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixDateSommairePunch", "cmdExporter_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Click()

 On Error GoTo Oups

 mvwDate.Visible = False

 Exit Sub

Oups:

 wOups "frmChoixDateSommairePunch", "Form_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 m_eDate = AUCUNE

 Call RemplirListViewFamille

 optProjets(I_OPT_TOUS_LES_PROJETS).Value = vbChecked

 Exit Sub

Oups:

 wOups "frmChoixDateSommairePunch", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewFamille()

 On Error GoTo Oups

 Dim rstFamille As ADODB.Recordset
 Dim itmFamille As ListItem

 Set rstFamille = New ADODB.Recordset
 
 Call rstFamille.Open("SELECT * FROM GrbFamille ORDER BY Famille", g_connData, adOpenDynamic, adLockOptimistic)
 
 Do While Not rstFamille.EOF
 Set itmFamille = lvwFamille.ListItems.Add
 
 itmFamille.Text = rstFamille.Fields("Famille")
 itmFamille.Tag = rstFamille.Fields("IDFamille")
 
 Call rstFamille.MoveNext
 Loop
 
  Call rstFamille.Close
  Set rstFamille = Nothing

  Exit Sub

Oups:

  wOups "frmChoixDateSommairePunch", "RemplirListViewFamille", Err, Err.number, Err.Description
End Sub

Private Sub mvwDate_LostFocus()

 On Error GoTo Oups

 mvwDate.Visible = False

 Exit Sub

Oups:

 wOups "frmChoixDateSommairePunch", "mvwDate_LostFocus", Err, Err.number, Err.Description
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

 wOups "frmChoixDateSommairePunch", "mvwDate_DateClick", Err, Err.number, Err.Description
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

 wOups "frmChoixDateSommairePunch", "mskDateDebut_GotFocus", Err, Err.number, Err.Description
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

 wOups "frmChoixDateSommairePunch", "mskDateFin_GotFocus", Err, Err.number, Err.Description
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

  wOups "frmChoixDateSommairePunch", "mskDateDebut_LostFocus", Err, Err.number, Err.Description
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

  wOups "frmChoixDateSommairePunch", "mskDateFin_LostFocus", Err, Err.number, Err.Description
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

  wOups "frmChoixDateSommairePunch", "cmdDateDebut_Click", Err, Err.number, Err.Description
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

  wOups "frmChoixDateSommairePunch", "cmdDateFin_Click", Err, Err.number, Err.Description
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

 wOups "frmChoixDateSommairePunch", "ValiderDate", Err, Err.number, Err.Description
End Function

Private Sub ImprimerPunchGeneral()

 On Error GoTo Oups

 Dim rstProjets As ADODB.Recordset
 Dim rstHeures As ADODB.Recordset
 Dim rstSommaire As ADODB.Recordset
 Dim datDebut As Date
 Dim datFin As Date
 Dim dblTotal As Double
 Dim dblSecondes As Double
 Dim dblHeures As Double

 Call g_connData.Execute("DELETE * FROM GrbImpressionSommairePunchGeneral")

 Set rstSommaire = New ADODB.Recordset
  Set rstProjets = New ADODB.Recordset
  Set rstHeures = New ADODB.Recordset

  Call rstSommaire.Open("SELECT * FROM GrbImpressionSommairePunchGeneral", g_connData, adOpenDynamic, adLockOptimistic)

  Call rstProjets.Open("SELECT DISTINCT NoProjet FROM GrbPunch WHERE Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

  Do While Not rstProjets.EOF
  Call rstSommaire.AddNew
 
  rstSommaire.Fields("NoProjet") = rstProjets.Fields("NoProjet")
 
  Call rstHeures.Open("SELECT HeureDébut, HeureFin FROM GrbPunch WHERE NoProjet = '" & rstProjets.Fields("NoProjet") & "' AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin is not Null", g_connData, adOpenDynamic, adLockOptimistic)
 
dblTotal = 0
 
1 Do While Not rstHeures.EOF
 datDebut = rstHeures.Fields("HeureDébut")

 If rstHeures.Fields("HeureFin") = "24:00" Then
 datFin = TimeSerial(24, 0, 0)
 Else
 datFin = rstHeures.Fields("HeureFin")
 End If
 
 dblSecondes = DateDiff("s", datDebut, datFin)

 dblHeures = Int(dblSecondes / 3600)
 
 dblSecondes = dblSecondes - (3600 * dblHeures)
 
 dblTotal = dblTotal + dblHeures + Round((dblSecondes / 3600), 2)
 
 Call rstHeures.MoveNext
 Loop
 
 Call rstHeures.Close
 
 rstSommaire.Fields("Total") = dblTotal
 
 Call rstSommaire.Update
 
 Call rstProjets.MoveNext
 Loop

1  Set DR_ListeProjets.DataSource = rstSommaire

 DR_ListeProjets.Sections("Section4").Controls("lblDateDebut").Caption = mskDateDebut.Text
 DR_ListeProjets.Sections("Section4").Controls("lblDateFin").Caption = mskDateFin.Text

Call DR_ListeProjets.Show(vbModal)

Call rstSommaire.Close
 
Set rstSommaire = Nothing
Set rstHeures = Nothing
Set rstProjets = Nothing

Exit Sub

Oups:

wOups "frmChoixDateSommairePunch", "ImprimerPunch", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerPunchGRB()

 On Error GoTo Oups

 Dim rstPunch As ADODB.Recordset
 Dim rstSommaire As ADODB.Recordset
 Dim datDebut As Date
 Dim datFin As Date
 Dim dblTotal As Double
 Dim dblSecondes As Double
 Dim dblHeures As Double
 Dim dblTotalHeures As Double
 Dim dblTotalKM As Double
 Dim sWhere As String
  Dim sWhereFamille As String
  Dim iCompteur As Integer

  Call g_connData.Execute("DELETE * FROM GrbImpressionSommairePunchGRB")

  Set rstSommaire = New ADODB.Recordset
  Set rstPunch = New ADODB.Recordset

  Call rstSommaire.Open("SELECT * FROM GrbImpressionSommairePunchGRB", g_connData, adOpenDynamic, adLockOptimistic)

  sWhere = "((Left(NoProjet, 6) = 'E" & Right$(Year(Date), 1) & "3000' OR Left(NoProjet, 6) = 'M" & Right$(Year(Date), 1) & "3000') AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin Is Not Null)"

  For iCompteur = 1 To lvwFamille.ListItems.count
If lvwFamille.ListItems(iCompteur).Checked = True Then
If sWhereFamille = "" Then
 sWhereFamille = " AND (Famille = " & lvwFamille.ListItems(iCompteur).Tag
 Else
 sWhereFamille = sWhereFamille & " OR Famille = " & lvwFamille.ListItems(iCompteur).Tag
 End If
 End If
Next

sWhere = sWhere & sWhereFamille & ")"

Call rstPunch.Open("SELECT employe, NoProjet, Date, HeureDébut, HeureFin, NbreKM, Commentaire FROM Grbemployés INNER JOIN GrbPunch ON Grbemployés.noemploye = GrbPunch.NoEmploye WHERE " & sWhere & " ORDER BY employe, Date, HeureDébut, HeureFin", g_connData, adOpenDynamic, adLockOptimistic)

Do While Not rstPunch.EOF
 Call rstSommaire.AddNew

rstSommaire.Fields("Employé") = rstPunch.Fields("employe")
 rstSommaire.Fields("NoProjet") = rstPunch.Fields("NoProjet")
 rstSommaire.Fields("Date") = rstPunch.Fields("Date")
 rstSommaire.Fields("Commentaire") = rstPunch.Fields("Commentaire")
 rstSommaire.Fields("HeureDébut") = rstPunch.Fields("HeureDébut")
 rstSommaire.Fields("HeureFin") = rstPunch.Fields("HeureFin")
 rstSommaire.Fields("NbreKM") = rstPunch.Fields("NbreKM")

1  datDebut = rstPunch.Fields("HeureDébut")

 If rstPunch.Fields("HeureFin") = "24:00" Then
 datFin = TimeSerial(24, 0, 0)
 Else
 datFin = rstPunch.Fields("HeureFin")
 End If
 
 dblSecondes = DateDiff("s", datDebut, datFin)

 dblHeures = Int(dblSecondes / 3600)
 
 dblSecondes = dblSecondes - (3600 * dblHeures)
 
 dblTotal = dblHeures + Round((dblSecondes / 3600), 2)
 
 dblTotalHeures = dblTotalHeures + dblTotal

 If Not IsNull(rstPunch.Fields("NbreKM")) Then
 If Trim(rstPunch.Fields("NbreKM")) <> "" Then
 dblTotalKM = dblTotalKM + rstPunch.Fields("NbreKM")
 End If
End If
 
 rstSommaire.Fields("Total") = dblTotal
 
Call rstSommaire.Update
 
 Call rstPunch.MoveNext
2  Loop

Set DR_SommairePunchGRB.DataSource = rstSommaire

30 DR_SommairePunchGRB.Orientation = rptOrientLandscape

DR_SommairePunchGRB.Sections("Section4").Controls("lblDateDebut").Caption = mskDateDebut.Text
DR_SommairePunchGRB.Sections("Section4").Controls("lblDateFin").Caption = mskDateFin.Text

DR_SommairePunchGRB.Sections("Section5").Controls("lblGrandTotal").Caption = dblTotalHeures
DR_SommairePunchGRB.Sections("Section5").Controls("lblGrandTotalKM").Caption = dblTotalKM

Call DR_SommairePunchGRB.Show(vbModal)

Call rstSommaire.Close
 
Set rstSommaire = Nothing
Set rstPunch = Nothing

Exit Sub

Oups:

wOups "frmChoixDateSommairePunch", "ImprimerPunch", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerSommaireHeures()

 On Error GoTo Oups

 Dim rstPunch As ADODB.Recordset
 Dim datDebut As Date
 Dim datFin As Date
 Dim dblSoumElec As Double
 Dim dblSoumMec As Double
 Dim dblProjGRBElec As Double
 Dim dblProjGRBMec As Double
 Dim dblProjElec As Double
 Dim dblProjMec As Double
 Dim dblFabElec As Double
  Dim dblFabMec As Double
  Dim dblRechElec As Double
  Dim dblRechMec As Double
  Dim dblAppelsElec As Double
  Dim dblAppelsMec As Double
  Dim dblGrandTotal As Double
  Dim dblSecondes As Double
  Dim dblHeures As Double
10 Dim dblTotalHeures As Double
Dim sWhere As String

Set rstPunch = New ADODB.Recordset

 'Soumissions électriques
sWhere = "(LEFT(NoProjet, 1) = 'E' AND MID(NoProjet, 3, 1) = '1' AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin Is Not Null)"

Call rstPunch.Open("SELECT Date, HeureDébut, HeureFin FROM GrbPunch WHERE " & sWhere, g_connData, adOpenForwardOnly, adLockReadOnly)

Do While Not rstPunch.EOF
 datDebut = rstPunch.Fields("HeureDébut")

 If rstPunch.Fields("HeureFin") = "24:00" Then
 datFin = TimeSerial(24, 0, 0)
 Else
 datFin = rstPunch.Fields("HeureFin")
 End If
 
dblSecondes = DateDiff("s", datDebut, datFin)

 dblHeures = Int(dblSecondes / 3600)
 
 dblSecondes = dblSecondes - (3600 * dblHeures)
 
 dblSoumElec = dblSoumElec + (dblHeures + Round((dblSecondes / 3600), 2))
 
 Call rstPunch.MoveNext
Loop

 Call rstPunch.Close

 'Soumissions mécaniques
1  sWhere = "(LEFT(NoProjet, 1) = 'M' AND MID(NoProjet, 3, 1) = '1' AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin Is Not Null)"

 Call rstPunch.Open("SELECT Date, HeureDébut, HeureFin FROM GrbPunch WHERE " & sWhere, g_connData, adOpenForwardOnly, adLockReadOnly)

 Do While Not rstPunch.EOF
 datDebut = rstPunch.Fields("HeureDébut")

 If rstPunch.Fields("HeureFin") = "24:00" Then
 datFin = TimeSerial(24, 0, 0)
 Else
 datFin = rstPunch.Fields("HeureFin")
 End If
 
 dblSecondes = DateDiff("s", datDebut, datFin)

 dblHeures = Int(dblSecondes / 3600)
 
 dblSecondes = dblSecondes - (3600 * dblHeures)
 
 dblSoumMec = dblSoumMec + (dblHeures + Round((dblSecondes / 3600), 2))
 
Call rstPunch.MoveNext
Loop

2  Call rstPunch.Close

 'Projets GRB électriques
sWhere = "(LEFT(NoProjet, 1) = 'E' AND MID(NoProjet, 3, 4) = '3000' AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin Is Not Null)"

2  Call rstPunch.Open("SELECT Date, HeureDébut, HeureFin FROM GrbPunch WHERE " & sWhere, g_connData, adOpenForwardOnly, adLockReadOnly)

Do While Not rstPunch.EOF
datDebut = rstPunch.Fields("HeureDébut")

 If rstPunch.Fields("HeureFin") = "24:00" Then
 datFin = TimeSerial(24, 0, 0)
3 Else
 datFin = rstPunch.Fields("HeureFin")
 End If
 
 dblSecondes = DateDiff("s", datDebut, datFin)

 dblHeures = Int(dblSecondes / 3600)
 
 dblSecondes = dblSecondes - (3600 * dblHeures)
 
 dblProjGRBElec = dblProjGRBElec + (dblHeures + Round((dblSecondes / 3600), 2))
 
 Call rstPunch.MoveNext
Loop

Call rstPunch.Close

 'Projets GRB mécaniques
sWhere = "(LEFT(NoProjet, 1) = 'M' AND MID(NoProjet, 3, 4) = '3000' AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin Is Not Null)"

3  Call rstPunch.Open("SELECT Date, HeureDébut, HeureFin FROM GrbPunch WHERE " & sWhere, g_connData, adOpenForwardOnly, adLockReadOnly)

Do While Not rstPunch.EOF
datDebut = rstPunch.Fields("HeureDébut")

 If rstPunch.Fields("HeureFin") = "24:00" Then
 datFin = TimeSerial(24, 0, 0)
 Else
 datFin = rstPunch.Fields("HeureFin")
 End If
 
dblSecondes = DateDiff("s", datDebut, datFin)

4 dblHeures = Int(dblSecondes / 3600)
 
4 dblSecondes = dblSecondes - (3600 * dblHeures)
 
4 dblProjGRBMec = dblProjGRBMec + (dblHeures + Round((dblSecondes / 3600), 2))
 
4 Call rstPunch.MoveNext
4 Loop

4 Call rstPunch.Close

 'Projets électriques
4 sWhere = "(LEFT(NoProjet, 1) = 'E' AND MID(NoProjet, 3, 1) = '3' AND MID(NoProjet, 3, 4) <> '3000' AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin Is Not Null)"

4 Call rstPunch.Open("SELECT Date, HeureDébut, HeureFin FROM GrbPunch WHERE " & sWhere, g_connData, adOpenForwardOnly, adLockReadOnly)

4 Do While Not rstPunch.EOF
4 datDebut = rstPunch.Fields("HeureDébut")

4 If rstPunch.Fields("HeureFin") = "24:00" Then
4  datFin = TimeSerial(24, 0, 0)
4  Else
4  datFin = rstPunch.Fields("HeureFin")
4  End If
 
4  dblSecondes = DateDiff("s", datDebut, datFin)

4  dblHeures = Int(dblSecondes / 3600)
 
4  dblSecondes = dblSecondes - (3600 * dblHeures)
 
4  dblProjElec = dblProjElec + (dblHeures + Round((dblSecondes / 3600), 2))
 
50 Call rstPunch.MoveNext
50 Loop

 Call rstPunch.Close

 'Projets mécaniques
 sWhere = "(LEFT(NoProjet, 1) = 'M' AND MID(NoProjet, 3, 1) = '3' AND MID(NoProjet, 3, 4) <> '3000' AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin Is Not Null)"

 Call rstPunch.Open("SELECT Date, HeureDébut, HeureFin FROM GrbPunch WHERE " & sWhere, g_connData, adOpenForwardOnly, adLockReadOnly)

 Do While Not rstPunch.EOF
 datDebut = rstPunch.Fields("HeureDébut")

 If rstPunch.Fields("HeureFin") = "24:00" Then
 datFin = TimeSerial(24, 0, 0)
 Else
 datFin = rstPunch.Fields("HeureFin")
 End If
 
5  dblSecondes = DateDiff("s", datDebut, datFin)

5  dblHeures = Int(dblSecondes / 3600)
 
5  dblSecondes = dblSecondes - (3600 * dblHeures)
 
5  dblProjMec = dblProjMec + (dblHeures + Round((dblSecondes / 3600), 2))
 
5  Call rstPunch.MoveNext
5  Loop

5  Call rstPunch.Close

 'Fabrication électrique
5  sWhere = "(LEFT(NoProjet, 1) = 'E' AND MID(NoProjet, 3, 1) = '4' AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin Is Not Null)"

60 Call rstPunch.Open("SELECT Date, HeureDébut, HeureFin FROM GrbPunch WHERE " & sWhere, g_connData, adOpenForwardOnly, adLockReadOnly)

60 Do While Not rstPunch.EOF
  datDebut = rstPunch.Fields("HeureDébut")

  If rstPunch.Fields("HeureFin") = "24:00" Then
  datFin = TimeSerial(24, 0, 0)
  Else
  datFin = rstPunch.Fields("HeureFin")
  End If
 
  dblSecondes = DateDiff("s", datDebut, datFin)

  dblHeures = Int(dblSecondes / 3600)
 
  dblSecondes = dblSecondes - (3600 * dblHeures)
 
  dblFabElec = dblFabElec + (dblHeures + Round((dblSecondes / 3600), 2))
 
6  Call rstPunch.MoveNext
6  Loop

6  Call rstPunch.Close

 'Fabrication mécanique
6  sWhere = "(LEFT(NoProjet, 1) = 'M' AND MID(NoProjet, 3, 1) = '4' AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin Is Not Null)"

6  Call rstPunch.Open("SELECT Date, HeureDébut, HeureFin FROM GrbPunch WHERE " & sWhere, g_connData, adOpenForwardOnly, adLockReadOnly)

6  Do While Not rstPunch.EOF
6  datDebut = rstPunch.Fields("HeureDébut")

6  If rstPunch.Fields("HeureFin") = "24:00" Then
70 datFin = TimeSerial(24, 0, 0)
  Else
  datFin = rstPunch.Fields("HeureFin")
  End If
 
  dblSecondes = DateDiff("s", datDebut, datFin)

  dblHeures = Int(dblSecondes / 3600)
 
  dblSecondes = dblSecondes - (3600 * dblHeures)
 
  dblFabMec = dblFabMec + (dblHeures + Round((dblSecondes / 3600), 2))
 
  Call rstPunch.MoveNext
  Loop

  Call rstPunch.Close

 'Recherche et développement électrique
  sWhere = "(LEFT(NoProjet, 1) = 'E' AND MID(NoProjet, 3, 1) = '9' AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin Is Not Null)"

   Call rstPunch.Open("SELECT Date, HeureDébut, HeureFin FROM GrbPunch WHERE " & sWhere, g_connData, adOpenForwardOnly, adLockReadOnly)

   Do While Not rstPunch.EOF
7  datDebut = rstPunch.Fields("HeureDébut")

7  If rstPunch.Fields("HeureFin") = "24:00" Then
7  datFin = TimeSerial(24, 0, 0)
7  Else
7  datFin = rstPunch.Fields("HeureFin")
7  End If
 
80 dblSecondes = DateDiff("s", datDebut, datFin)

  dblHeures = Int(dblSecondes / 3600)
 
  dblSecondes = dblSecondes - (3600 * dblHeures)
 
  dblRechElec = dblRechElec + (dblHeures + Round((dblSecondes / 3600), 2))
 
  Call rstPunch.MoveNext
  Loop

  Call rstPunch.Close

 'Recherche et développement mécanique
  sWhere = "(LEFT(NoProjet, 1) = 'M' AND MID(NoProjet, 3, 1) = '9' AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin Is Not Null)"

  Call rstPunch.Open("SELECT Date, HeureDébut, HeureFin FROM GrbPunch WHERE " & sWhere, g_connData, adOpenForwardOnly, adLockReadOnly)

  Do While Not rstPunch.EOF
  datDebut = rstPunch.Fields("HeureDébut")

  If rstPunch.Fields("HeureFin") = "24:00" Then
   datFin = TimeSerial(24, 0, 0)
   Else
   datFin = rstPunch.Fields("HeureFin")
   End If
 
8  dblSecondes = DateDiff("s", datDebut, datFin)

8  dblHeures = Int(dblSecondes / 3600)
 
8  dblSecondes = dblSecondes - (3600 * dblHeures)
 
8  dblRechMec = dblRechMec + (dblHeures + Round((dblSecondes / 3600), 2))
 
90 Call rstPunch.MoveNext
90 Loop

  Call rstPunch.Close

 'Appels de services électriques
  sWhere = "(LEFT(NoProjet, 1) = 'E' AND (MID(NoProjet, 3, 1) = '5' OR MID(NoProjet, 3, 1) = '7') AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin Is Not Null)"

  Call rstPunch.Open("SELECT Date, HeureDébut, HeureFin FROM GrbPunch WHERE " & sWhere, g_connData, adOpenForwardOnly, adLockReadOnly)

  Do While Not rstPunch.EOF
  datDebut = rstPunch.Fields("HeureDébut")

  If rstPunch.Fields("HeureFin") = "24:00" Then
  datFin = TimeSerial(24, 0, 0)
  Else
  datFin = rstPunch.Fields("HeureFin")
  End If
 
 dblSecondes = DateDiff("s", datDebut, datFin)

   dblHeures = Int(dblSecondes / 3600)
 
 dblSecondes = dblSecondes - (3600 * dblHeures)
 
   dblAppelsElec = dblAppelsElec + (dblHeures + Round((dblSecondes / 3600), 2))
 
 Call rstPunch.MoveNext
   Loop

 Call rstPunch.Close

 'Appels de services mécaniques
9  sWhere = "(LEFT(NoProjet, 1) = 'M' AND (MID(NoProjet, 3, 1) = '5' OR MID(NoProjet, 3, 1) = '7') AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin Is Not Null)"

 Call rstPunch.Open("SELECT Date, HeureDébut, HeureFin FROM GrbPunch WHERE " & sWhere, g_connData, adOpenForwardOnly, adLockReadOnly)

100 Do While Not rstPunch.EOF
1datDebut = rstPunch.Fields("HeureDébut")

1 If rstPunch.Fields("HeureFin") = "24:00" Then
 datFin = TimeSerial(24, 0, 0)
1Else
 datFin = rstPunch.Fields("HeureFin")
1End If
 
 dblSecondes = DateDiff("s", datDebut, datFin)

1dblHeures = Int(dblSecondes / 3600)
 
 dblSecondes = dblSecondes - (3600 * dblHeures)
 
1dblAppelsMec = dblAppelsMec + (dblHeures + Round((dblSecondes / 3600), 2))
 
10  Call rstPunch.MoveNext
10  Loop

10  Call rstPunch.Close

10  dblGrandTotal = dblSoumElec + _
 dblSoumMec + _
 dblProjGRBElec + _
 dblProjGRBMec + _
 dblProjElec + _
 dblProjMec + _
 dblFabElec + _
 dblFabMec + _
 dblRechElec + _
 dblRechMec + _
 dblAppelsElec + _
 dblAppelsMec

 'Ce recordset sert à rien. Il sert seulement à l'ouverture du DataReport.
 'Un DataReport ne peut pas ouvrir s'il n'a pas de DataSource
10  Call rstPunch.Open("SELECT * FROM GrbPunch WHERE NoProjet = '0000000'", g_connData, adOpenForwardOnly, adLockReadOnly)

10  Set DR_SommaireHeures.DataSource = rstPunch

109DR_SommaireHeures.Orientation = rptOrientPortrait

10  DR_SommaireHeures.Sections("Section2").Controls("lblDateDebut").Caption = mskDateDebut.Text
110DR_SommaireHeures.Sections("Section2").Controls("lblDateFin").Caption = mskDateFin.Text

110 DR_SommaireHeures.Sections("Section2").Controls("lblSoumElec").Caption = dblSoumElec
11 DR_SommaireHeures.Sections("Section2").Controls("lblSoumMec").Caption = dblSoumMec

11 DR_SommaireHeures.Sections("Section2").Controls("lblTotalSoum").Caption = dblSoumElec + dblSoumMec

11 DR_SommaireHeures.Sections("Section2").Controls("lblProjGRBElec").Caption = dblProjGRBElec
11 DR_SommaireHeures.Sections("Section2").Controls("lblProjGRBMec").Caption = dblProjGRBMec

11 DR_SommaireHeures.Sections("Section2").Controls("lblTotalProjGRB").Caption = dblProjGRBElec + dblProjGRBMec

11 DR_SommaireHeures.Sections("Section2").Controls("lblProjElec").Caption = dblProjElec
11 DR_SommaireHeures.Sections("Section2").Controls("lblProjMec").Caption = dblProjMec

11 DR_SommaireHeures.Sections("Section2").Controls("lblTotalProj").Caption = dblProjElec + dblProjMec

11 DR_SommaireHeures.Sections("Section2").Controls("lblFabElec").Caption = dblFabElec
11 DR_SommaireHeures.Sections("Section2").Controls("lblFabMec").Caption = dblFabMec

116DR_SommaireHeures.Sections("Section2").Controls("lblTotalFab").Caption = dblFabElec + dblFabMec

11  DR_SommaireHeures.Sections("Section2").Controls("lblRechElec").Caption = dblRechElec
1 DR_SommaireHeures.Sections("Section2").Controls("lblRechMec").Caption = dblRechMec

11  DR_SommaireHeures.Sections("Section2").Controls("lblTotalRech").Caption = dblRechElec + dblRechMec

1 DR_SommaireHeures.Sections("Section2").Controls("lblAppelsElec").Caption = dblAppelsElec
11  DR_SommaireHeures.Sections("Section2").Controls("lblAppelsMec").Caption = dblAppelsMec

1 DR_SommaireHeures.Sections("Section2").Controls("lblTotalAppels").Caption = dblAppelsElec + dblAppelsMec

11  DR_SommaireHeures.Sections("Section2").Controls("lblGrandTotal").Caption = dblGrandTotal

1 Call DR_SommaireHeures.Show(vbModal)

1 Call rstPunch.Close
 
12 Set rstPunch = Nothing

12 Exit Sub

Oups:

12 wOups "frmChoixDateSommairePunch", "ImprimerSommaireHeures", Err, Err.number, Err.Description
End Sub

Private Sub optProjets_Click(Index As Integer)
 If optProjets(I_OPT_PROJETS_GRB).Value = True Then
 fraFamille.Enabled = True
 Else
 fraFamille.Enabled = False
 End If
End Sub
