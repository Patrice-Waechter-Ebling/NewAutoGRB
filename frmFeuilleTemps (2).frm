VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFeuilleTemps 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Feuilles de temps"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   8445
   Begin VB.CommandButton CmdModifier 
      Caption         =   "Modifier Type"
      Height          =   495
      Left            =   6000
      TabIndex        =   34
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdexcel 
      Caption         =   "Excel"
      Height          =   495
      Left            =   2450
      TabIndex        =   33
      Top             =   6600
      Width           =   1095
   End
   Begin MSComCtl2.MonthView mvwDate 
      Height          =   2370
      Left            =   720
      TabIndex        =   7
      Top             =   1920
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
   Begin VB.OptionButton optTypePunch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Mécanique"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   30
      Top             =   4200
      Width           =   1095
   End
   Begin VB.OptionButton optTypePunch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Électrique"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   29
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdExporter 
      Caption         =   "Exporter"
      Height          =   495
      Left            =   120
      TabIndex        =   28
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CheckBox chkKM 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "KM :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   720
      TabIndex        =   19
      Top             =   6000
      Width           =   735
   End
   Begin VB.TextBox txtSemaine 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdDateSemaine 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   315
      Left            =   7920
      TabIndex        =   4
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton cmdImprimer 
      Caption         =   "Imprimer"
      Height          =   495
      Left            =   1290
      TabIndex        =   23
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdDate 
      Caption         =   "..."
      Height          =   255
      Left            =   2880
      TabIndex        =   13
      Top             =   4920
      Width           =   375
   End
   Begin VB.TextBox txtCommentaires 
      Height          =   765
      Left            =   4560
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   5520
      Width           =   3735
   End
   Begin VB.ComboBox cmbEmployé 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   840
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cmdEnregistrer 
      Caption         =   "Enregistrer"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4800
      TabIndex        =   25
      Top             =   6600
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwPunch 
      Height          =   3015
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Projet"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Début"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Fin"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Client"
         Object.Width           =   3889
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Commentaire"
         Object.Width           =   2752
      EndProperty
   End
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   495
      Left            =   7200
      TabIndex        =   27
      Top             =   6600
      Width           =   1095
   End
   Begin MSMask.MaskEdBox mskHeureFin 
      Height          =   255
      Left            =   1560
      TabIndex        =   18
      Top             =   5640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskDate 
      Height          =   255
      Left            =   1560
      TabIndex        =   12
      Top             =   4920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskHeureDebut 
      Height          =   255
      Left            =   1560
      TabIndex        =   15
      Top             =   5280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdAjouter 
      Caption         =   "Ajouter"
      Height          =   495
      Left            =   3600
      TabIndex        =   24
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      Height          =   495
      Left            =   7200
      TabIndex        =   26
      Top             =   6600
      Width           =   1095
   End
   Begin MSMask.MaskEdBox mskNoProjet 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Top             =   4560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "#####-##"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtKM 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   21
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox txtClient 
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   4320
      Width           =   3735
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   4920
      Width           =   3735
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   "Type :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   32
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label lblPrefixe 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   31
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Km"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   22
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Client :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4560
      TabIndex        =   9
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label lblSemaine 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Semaine du :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5280
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Commentaires :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4560
      TabIndex        =   16
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Employé :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Heure de fin :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Heure de début :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date (AA-MM-JJ):"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Numéro de projet :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Width           =   1335
   End
End
Attribute VB_Name = "frmFeuilleTemps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Types quand c'est un 51
Private Const I_TYPE_ELEC_INSTALLATION As Integer = 0
Private Const I_TYPE_ELEC_MISE_SERVICE As Integer = 1

'Types quand c'est pas un 51
Private Const I_TYPE_ELEC_DESSIN As Integer = 0
Private Const I_TYPE_ELEC_FABRICATION As Integer = 1
Private Const I_TYPE_ELEC_ASSEMBLAGE As Integer = 2
Private Const I_TYPE_ELEC_PROG_INTERFACE As Integer = 3
Private Const I_TYPE_ELEC_PROG_AUTOMATE As Integer = 4
Private Const I_TYPE_ELEC_PROG_ROBOT As Integer = 5
Private Const I_TYPE_ELEC_VISION As Integer = 6
Private Const I_TYPE_ELEC_TEST As Integer = 7
Private Const I_TYPE_ELEC_FORMATION As Integer = 8
Private Const I_TYPE_ELEC_GESTION As Integer = 9
Private Const I_TYPE_ELEC_SHIPPING As Integer = 10
Private Const I_TYPE_ELEC_prototypage As Integer = 11

'Types quand c'est un 51
Private Const I_TYPE_MEC_INSTALLATION As Integer = 0

'Types quand c'est pas un 51
Private Const I_TYPE_MEC_DESSIN As Integer = 0
Private Const I_TYPE_MEC_COUPE As Integer = 1
Private Const I_TYPE_MEC_MACHINAGE As Integer = 2
Private Const I_TYPE_MEC_SOUDURE As Integer = 3
Private Const I_TYPE_MEC_ASSEMBLAGE As Integer = 4
Private Const I_TYPE_MEC_PEINTURE As Integer = 5
Private Const I_TYPE_MEC_TEST As Integer = 6
Private Const I_TYPE_MEC_FORMATION As Integer = 7
Private Const I_TYPE_MEC_GESTION As Integer = 8
Private Const I_TYPE_MEC_SHIPPING As Integer = 9
Private Const I_TYPE_MEC_prototypage As Integer = 10

Private Const I_LVW_PROJET As Integer = 0
Private Const I_LVW_DATE As Integer = 1
Private Const I_LVW_DEBUT As Integer = 2
Private Const I_LVW_FIN As Integer = 3
Private Const I_LVW_CLIENT As Integer = 4
Private Const I_LVW_TYPE As Integer = 5
Private Const I_LVW_COMMENTAIRE As Integer = 6

Private Const I_OPT_ELECTRIQUE As Integer = 0
Private Const I_OPT_MECANIQUE As Integer = 1

Private Enum enumMode
 MODE_INACTIF = 0
 MODE_MODIF = 1
 MODE_AJOUT = 2
End Enum

Private m_lIDPunch As Long
Public m_datSemaine As Date
Private m_bModifProj As Boolean
Private m_eMode As enumMode
Private m_bClick As Boolean

Private Sub ActiverControles(ByVal eMode As enumMode)

 On Error GoTo Oups

 'Activation des controles dépendamment du mode choisi
 Dim bListView As Boolean 'Pour le ListView
 Dim bEmploye As Boolean 'Pour la liste des employés
 Dim bSemaine As Boolean 'Pour le bouton de la semaine
 Dim bChamps As Boolean 'Pour les champs
 Dim bImprimer As Boolean 'Pour le bouton "Imprimer"
 Dim bAjouter As Boolean 'Pour le bouton "Ajouter"
 Dim bAnnuler As Boolean 'Pour le bouton "Annuler"
 Dim bEnregistrer As Boolean 'Pour le bouton "Enregistrer"
 Dim bFermer As Boolean 'Pour le bouton "Fermer"
 
 m_eMode = eMode
 
  Select Case eMode
 'Mode pour ouverture et après ajout et modif
 Case MODE_INACTIF:
  bListView = True
  bEmploye = True
  bSemaine = True
  bImprimer = True
  bAjouter = True
  bFermer = True
 
 'Mode pour la modification
  Case MODE_MODIF:
 bListView = True
bEmploye = True
 bSemaine = True
 bChamps = True
 bImprimer = True
 bAjouter = True
 bAnnuler = True
 bEnregistrer = True
 bFermer = True
 
 'Mode pour l'ajout
 Case MODE_AJOUT:
 bChamps = True
 bAnnuler = True
 bEnregistrer = True
End Select
 
 lvwPunch.Enabled = bListView
cmbemployé.Enabled = bEmploye
 cmdDateSemaine.Enabled = bSemaine
mskNoProjet.Enabled = bChamps
 mskDate.Enabled = bChamps
1  cmdDate.Enabled = bChamps
 mskHeureDebut.Enabled = bChamps
 mskHeureFin.Enabled = bChamps
txtCommentaires.Enabled = bChamps
cmbType.Enabled = bChamps
optTypePunch(I_OPT_ELECTRIQUE).Enabled = bChamps
optTypePunch(I_OPT_MECANIQUE).Enabled = bChamps

If g_bModificationFeuillesTemps = True Or (cmbemployé.Text = g_sEmploye) Then
 cmdEnregistrer.Enabled = bEnregistrer
 cmdexcel.Enabled = bImprimer
 cmdImprimer.Enabled = bImprimer
 Cmdajouter.Enabled = bAjouter
Else
 cmdEnregistrer.Enabled = False
 cmdexcel.Enabled = False
cmdImprimer.Enabled = False
 Cmdajouter.Enabled = False
2  End If

cmdAnnuler.Visible = bAnnuler
2  Cmdfermer.Visible = bFermer

Exit Sub

Oups:

2  wOups "frmFeuilleTemps", "ActiverControles", Err, Err.number, Err.Description
End Sub

Private Sub chkKM_Click()

 On Error GoTo Oups

 If chkKM.Value = vbChecked Then
 txtKM.Enabled = True
 Else
 txtKM.Text = ""
 txtKM.Enabled = False
 End If

 Exit Sub

Oups:

 wOups "frmFeuilleTemps", "chkKM_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdexcel_Click()
Dim xlworksheet As Excel.Workbook
Dim xlsheet As Excel.Application
Dim info As ADODB.Recordset
Dim row As Integer
Dim i As Integer



Set xlsheet = New Excel.Application
Set xlworksheet = xlsheet.Workbooks.Add

xlsheet.Cells(1, 1).Value = "Employé"
xlsheet.Cells(1, 4).Value = "Semaine Du:"
xlsheet.Cells(3, 1).Value = "Projet"
xlsheet.Cells(3, 2).Value = "Date"
xlsheet.Cells(3, 3).Value = "Début"
xlsheet.Cells(3, 4).Value = "Fin"
xlsheet.Cells(3, 5).Value = "Client"
xlsheet.Cells(3, 6).Value = "Type"
xlsheet.Cells(3, 7).Value = "Commentaire"

With xlsheet.range("A1;D1;A3:G3")
 .Font.Bold = True
 .Font.SIZE = 11
End With

xlsheet.Cells(1, 2).Value = cmbemployé.Text
xlsheet.Cells(1, 5).Value = txtSemaine.Text



row = 4
For i = 1 To lvwPunch.ListItems.count
 xlsheet.Cells(row, 1).Value = lvwPunch.ListItems(i).Text
 xlsheet.Cells(row, 2).Value = lvwPunch.ListItems(i).SubItems(1)
 xlsheet.Cells(row, 3).Value = lvwPunch.ListItems(i).SubItems(2)
 xlsheet.Cells(row, 4).Value = lvwPunch.ListItems(i).SubItems(3)
 xlsheet.Cells(row, 5).Value = lvwPunch.ListItems(i).SubItems(4)
 xlsheet.Cells(row, 6).Value = lvwPunch.ListItems(i).SubItems(5)
 xlsheet.Cells(row, 7).Value = lvwPunch.ListItems(i).SubItems(6)
 row = row + 1
Next
xlsheet.range("A:G").Columns.AutoFit







xlsheet.Visible = True

Set xlsheet = Nothing



End Sub

Private Sub cmdModifier_Click()
Call FrmModType.Show(vbModal)



End Sub

Private Sub optTypePunch_Click(Index As Integer)
 
 On Error GoTo Oups

 If Index = I_OPT_ELECTRIQUE Then
 lblPrefixe.Caption = "E"
 Else
 lblPrefixe.Caption = "M"
 End If
 
 Call RemplirComboType

 Exit Sub

Oups:

 wOups "frmFeuilleTemps", "optTypePunch_Click", Err, Err.number, Err.Description
End Sub

Private Sub Cmdajouter_Click()

 On Error GoTo Oups

 Call ViderChamps

 Call ActiverControles(MODE_AJOUT)

 Exit Sub

Oups:

 wOups "frmFeuilleTemps", "cmdAjouter_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnuler_Click()

 On Error GoTo Oups

 'Annuler l'ajout ou la modif
 Call ViderChamps
 
 Call ActiverControles(MODE_INACTIF)

 Exit Sub

Oups:

 wOups "frmFeuilleTemps", "cmdAnnuler_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdDateSemaine_Click()

 On Error GoTo Oups

 'Affichage du calendrier pour choisir une semaine
 Call OuvrirForm(frmFeuilleTempsCalendrier, True)

 Exit Sub

Oups:

 wOups "frmFeuilleTemps", "cmdDateSemaine_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmbEmployé_Click()

 On Error GoTo Oups

 'Rempli le listview si la semaine a été choisie
 Call ViderChamps
 
 Call ActiverControles(MODE_INACTIF)
 
 'Il faut remplir le listview dépendant la semaine
 If txtSemaine.Text <> vbNullString Then
 Call RemplirListView
 End If
 
 cmdEnregistrer.Enabled = False

 Exit Sub

Oups:

 wOups "frmFeuilleTemps", "cmbEmployé_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdDate_Click()

 On Error GoTo Oups
 'Ouverture du calendrier
 
 'Si il y a une date valide, la date par défaut est celle-ci, sinon c'est la date
 'd'aujourd'hui
 If Trim$(mskDate.Text) <> vbNullString Then
 If ValiderDate(mskDate.Text) = True Then
 mvwDate.Value = mskDate.Text
 Else
 mvwDate.Value = Date
 End If
 Else
 mvwDate.Value = Date
 End If
 
 mvwDate.Visible = True
 
  Call mvwDate.SetFocus

  Exit Sub

Oups:

  wOups "frmFeuilleTemps", "cmdDate_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdEnregistrer_Click()

 On Error GoTo Oups
 
 'Enregistrer l'ajout ou la modif
 Dim rstPunch As ADODB.Recordset
 Dim rstProjSoum As ADODB.Recordset
 Dim bInstallation As Boolean
 
 Set rstProjSoum = New ADODB.Recordset

 Call rstProjSoum.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & lblPrefixe.Caption & mskNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 If Not rstProjSoum.EOF Then
 If rstProjSoum.Fields("Ouvert") = False Then
 Call MsgBox("Ce projet est fermé!", vbOKOnly, "Erreur")
 
 Call rstProjSoum.Close
 Set rstProjSoum = Nothing
 
  Exit Sub
  End If
  Else
  Call MsgBox("Projet inexistant!", vbOKOnly, "Erreur")
 
  Call rstProjSoum.Close
  Set rstProjSoum = Nothing
 
  Exit Sub
  End If

10 Call rstProjSoum.Close
Set rstProjSoum = Nothing
 
 'Valider l'heure de début
If ValiderHeure(mskHeureDebut.Text) = False Then
 Call MsgBox("L'heure de début est invalide!", vbOKOnly, "Erreur")
 
 Exit Sub
Else
 If mskHeureFin.Text <> "24:00" Then
 If mskHeureFin.Text <> vbNullString Then
 'Valider l'heure de fin
 If ValiderHeure(mskHeureFin.Text) = False Then
 Call MsgBox("L'heure de fin est invalide!", vbOKOnly, "Erreur")
 
 Exit Sub
 End If
 End If
 End If
 End If
 
 'Valider la date
If ValiderDate(mskDate.Text) = False Then
 Call MsgBox("La date est invalide!", vbOKOnly, "Erreur")
 
 Exit Sub
 End If
 
 'Si les champs importants ont été rempli
1  If mskNoProjet.Text = vbNullString Or mskDate.Text = vbNullString Or mskHeureDebut.Text = vbNullString Then
 Call MsgBox("Champs vide!", vbOKOnly, "Erreur")
 
 Exit Sub
End If
 
 'Si les champs importants ont été rempli
If Right$(mskNoProjet.Text, 1) = "_" Then
 Call MsgBox("Le numéro de projet est invalide!", vbOKOnly, "Erreur")
 
 Exit Sub
End If

If cmbType.Text = "" And cmbType.Visible = True Then
 Call MsgBox("Le type est obligatoire!", vbOKOnly, "Erreur")

 Exit Sub
End If
 
Set rstPunch = New ADODB.Recordset
 
2  If m_eMode = MODE_AJOUT Then
 Call rstPunch.Open("SELECT * FROM GrbPunch", g_connData, adOpenDynamic, adLockOptimistic)
 
Call rstPunch.AddNew
Else
Call rstPunch.Open("SELECT * FROM GrbPunch WHERE IDPunch = " & m_lIDPunch, g_connData, adOpenDynamic, adLockOptimistic)
End If
 
2  rstPunch.Fields("NoEmploye") = cmbemployé.ItemData(cmbemployé.ListIndex)
 
rstPunch.Fields("NoProjet") = lblPrefixe.Caption & mskNoProjet.Text
30 rstPunch.Fields("Date") = mskDate.Text
rstPunch.Fields("HeureDébut") = mskHeureDebut.Text
rstPunch.Fields("HeureFin") = mskHeureFin.Text

If chkKM.Value = vbChecked Then
 rstPunch.Fields("KM") = True

 If txtKM.Text <> "" Then
 txtKM.Text = Replace(txtKM.Text, ".", ",")

 If IsNumeric(txtKM.Text) Then
 rstPunch.Fields("NbreKM") = txtKM.Text
 Else
 rstPunch.Fields("NbreKM") = 0
 End If
Else
 txtKM.Text = 0
End If
Else
rstPunch.Fields("KM") = False
 rstPunch.Fields("NbreKM") = ""
3  End If
 
 If txtClient.Text <> vbNullString Then
rstPunch.Fields("NoClient") = txtClient.Tag
Else
4 rstPunch.Fields("NoClient") = vbNullString
4 End If

4 If IsNumeric(Right$(mskNoProjet.Text, 2)) Then
4 If CInt(Right$(mskNoProjet.Text, 2)) >= 51 And CInt(Right$(mskNoProjet.Text, 2)) <= 5 Then
4 bInstallation = True
4 Else
4 bInstallation = False
4 End If
4 Else
4 bInstallation = False
4  End If
 
4  If bInstallation = True Then
4  If lblPrefixe.Caption = "E" Then
4  Select Case cmbType.ListIndex
 Case I_TYPE_ELEC_INSTALLATION: rstPunch.Fields("Type") = "Installation"
4  Case I_TYPE_ELEC_MISE_SERVICE: rstPunch.Fields("Type") = "MiseService"
4  End Select
4  Else
4  Select Case cmbType.ListIndex
 Case I_TYPE_MEC_INSTALLATION: rstPunch.Fields("Type") = "Installation"
50 End Select
5 End If
 Else
 If lblPrefixe.Caption = "E" Then
 rstPunch.Fields("Type") = cmbType.Text
5  Else
 rstPunch.Fields("Type") = cmbType.Text
  End If
  End If
 
  rstPunch.Fields("Commentaire") = txtCommentaires.Text
 
  Call rstPunch.Update
 
6  Call rstPunch.Close
6  Set rstPunch = Nothing
 
6  Call ViderChamps
 
6  Call ActiverControles(MODE_INACTIF)
 
6  Call RemplirListView

6  Exit Sub

Oups:

6  wOups "frmFeuilleTemps", "cmdEnregistrer_Click", Err, Err.number, Err.Description
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

 wOups "frmFeuilleTemps", "ValiderDate", Err, Err.number, Err.Description
End Function

Private Function ValiderHeure(ByVal sHeure As String) As Boolean

 On Error GoTo Oups

 'validation des heures
 Dim sHour As String
 Dim sMinute As String
 Dim sSecond As String
 
 sHour = Left$(sHeure, 2)
 
 sMinute = Mid$(sHeure, 4, 2)
 
 sSecond = Right$(sHeure, 2)
 
 ValiderHeure = True
 
 'Si numérique
 If Not IsNumeric(sHour) Then
 ValiderHeure = False
 
 Exit Function
  Else
 'Si entre 0 et 23
  If sHour < 0 Or sHour > 23 Then
  ValiderHeure = False

  Exit Function
  End If
  End If
 
 'Si numérique
  If Not IsNumeric(sMinute) Then
  ValiderHeure = False
 
Exit Function
Else
 'Si entre 0 et 59
 If sMinute < 0 Or sMinute > 5 Then
 ValiderHeure = False
 
 Exit Function
 End If
End If

 'Si numérique
If Not IsNumeric(sSecond) Then
 ValiderHeure = False
 
 Exit Function
Else
 'Si entre 0 et 59
 If sSecond < 0 Or sSecond > 5 Then
 ValiderHeure = False
 
 Exit Function
 End If
End If

 Exit Function

Oups:

wOups "frmFeuilleTemps", "ValiderHeure", Err, Err.number, Err.Description
End Function

Private Sub ViderChamps()

 On Error GoTo Oups

 'Vider les champs
 mskNoProjet.Text = "_____-__"
 
 mskDate.Text = vbNullString
 mskHeureDebut.Text = vbNullString
 mskHeureFin.Text = vbNullString
 txtCommentaires.Text = vbNullString
 txtClient.Text = vbNullString
 chkKM.Value = vbUnchecked
 cmbType.ListIndex = -1

 txtKM.Text = vbNullString

 Exit Sub

Oups:

  wOups "frmFeuilleTemps", "ViderChamps", Err, Err.number, Err.Description
End Sub

Private Sub cmdexporter_Click()

 On Error GoTo Oups

 Call frmChoixDateImpressionFT.Show(vbModal)

 Exit Sub

Oups:

 wOups "frmFeuilleTemps", "cmdExporter_Click", Err, Err.number, Err.Description
End Sub

Private Sub Cmdfermer_Click()

 On Error GoTo Oups

 'Fermeture de la fenêtre
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmFeuilleTemps", "cmdFermer_Click", Err, Err.number, Err.Description
End Sub

Public Sub RemplirListView()

 On Error GoTo Oups
 
 'Rempli le listview selon la semaine et l'employé choisi
 Dim rstPunch As ADODB.Recordset
 Dim rstNomClient As ADODB.Recordset
 Dim itmPunch As ListItem
 Dim sTemp As String
 Dim iTemp As Integer
 Dim sDateDebut As String
 Dim sDateFin As String
 
 'Il faut vider le ListView pour ne pas le remplir plein de fois
 Call lvwPunch.ListItems.Clear
 
 sDateDebut = ConvertDate(m_datSemaine)
 sDateFin = ConvertDate(GetLastDay(m_datSemaine))

  Set rstNomClient = New ADODB.Recordset
  Set rstPunch = New ADODB.Recordset

  Call rstPunch.Open("SELECT * FROM GrbPunch WHERE noEmploye = " & cmbemployé.ItemData(cmbemployé.ListIndex) & " AND Date BETWEEN '" & sDateDebut & "' AND '" & sDateFin & "' ORDER BY Date, HeureDébut", g_connData, adOpenDynamic, adLockOptimistic)
 
  Do While Not rstPunch.EOF
  Set itmPunch = lvwPunch.ListItems.Add
 
 'Numéro du projet
  itmPunch.Text = rstPunch.Fields("noProjet")
 
 'IDPunch dans le tag
  itmPunch.Tag = rstPunch.Fields("IDPunch")
 
 'Date
  itmPunch.SubItems(I_LVW_DATE) = rstPunch.Fields("Date")
 
 'Heure de début
itmPunch.SubItems(I_LVW_DEBUT) = rstPunch.Fields("HeureDébut")
 
 'Heure de fin
1 If Not IsNull(rstPunch.Fields("HeureFin")) Then
 itmPunch.SubItems(I_LVW_FIN) = rstPunch.Fields("HeureFin")
 Else
 itmPunch.SubItems(I_LVW_FIN) = vbNullString
 End If
 
 If Not IsNull(rstPunch.Fields("NoClient")) And rstPunch.Fields("NoClient") > "0" Then
 Call rstNomClient.Open("SELECT NomClient FROM GrbClient WHERE IDClient = " & rstPunch.Fields("NoClient"), g_connData, adOpenDynamic, adLockOptimistic)
 
 itmPunch.SubItems(I_LVW_CLIENT) = rstNomClient.Fields("NomClient")
 
 Call rstNomClient.Close
 Else
 itmPunch.SubItems(I_LVW_CLIENT) = vbNullString
End If

 If Not IsNull(rstPunch.Fields("Type")) Then
 If Left$(itmPunch.Text, 1) = "E" Then
 itmPunch.SubItems(I_LVW_TYPE) = rstPunch.Fields("Type")
 Else
 itmPunch.SubItems(I_LVW_TYPE) = rstPunch.Fields("Type")
 End If
 End If

 If Not IsNull(rstPunch.Fields("Commentaire")) Then
 itmPunch.SubItems(I_LVW_COMMENTAIRE) = rstPunch.Fields("Commentaire")
 Else
 itmPunch.SubItems(I_LVW_COMMENTAIRE) = vbNullString
 End If
 
 Call rstPunch.MoveNext
Loop
 
Call rstPunch.Close
3  Set rstPunch = Nothing

Set rstNomClient = Nothing

3  Exit Sub

Oups:

wOups "frmFeuilleTemps", "RemplirListView", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 Dim datTemp As Date
 Dim iNoJour As Integer
 Call frmPunch.Table_exist
 Call RemplirComboEmploye

 If g_admin = False Then
 frmFeuilleTemps.width = 7350
 txtCommentaires.Left = 3360
 cmbType.Left = 3360
 Label7.Left = 3360
 lblType.Left = 3360
 txtClient.Left = 3360
 Label8.Left = 3360
 lvwPunch.width = 6975
 cmdDateSemaine.Left = 6720
 Cmdfermer.Left = 6000
 cmdAnnuler.Left = 6000
 txtSemaine.Left = 5040
 lblSemaine.Left = 4080
 cmdModifier.Visible = False
 End If
 
 
 datTemp = Date

 iNoJour = 1

 'Pour avoir la date de la semaine précédente
 Do While iNoJour < 8
 datTemp = datTemp - TimeSerial(24, 0, 0)

 iNoJour = iNoJour + 1
 Loop
 
 m_datSemaine = GetFirstDay(datTemp)

 'On affiche le dimanche de la semaine précédente
  txtSemaine.Text = GetDateTexte(m_datSemaine)
 
  Call RemplirListView

  Screen.MousePointer = vbDefault

  Exit Sub

Oups:

  wOups "frmFeuilleTemps", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboEmploye()

 On Error GoTo Oups
 
 'Rempli le combo des employés
 Dim rstEmploye As ADODB.Recordset
 
 'Il faut vider le combo pour ne pas le remplir plein de fois
 Call cmbemployé.Clear
 
 Set rstEmploye = New ADODB.Recordset
 
 Call rstEmploye.Open("SELECT * FROM GrbEmployés WHERE Actif = true", g_connData, adOpenDynamic, adLockOptimistic)
 
 Do While Not rstEmploye.EOF
 'Ajout du nom de l'employé dans le combo
 Call cmbemployé.AddItem(rstEmploye.Fields("employe"))
 
 'Ajout du numéro de l'employé dans l'ItemData
 cmbemployé.ItemData(cmbemployé.newIndex) = rstEmploye.Fields("noemploye")
 
 Call rstEmploye.MoveNext
 Loop
 
 Call rstEmploye.Close
  Set rstEmploye = Nothing
 
 'Si le combo n'est pas vide, on sélectionne le premier
  If cmbemployé.ListCount > 0 Then
  cmbemployé.ListIndex = 0
  End If

  Exit Sub

Oups:

  wOups "frmFeuilleTemps", "RemplirComboEmploye", Err, Err.number, Err.Description
End Sub

Private Sub lvwPunch_ItemClick(ByVal Item As MSComctlLib.ListItem)

 On Error GoTo Oups

 Dim rstPunch As ADODB.Recordset
 Dim iCompteur As Integer
 Dim G As Integer
 
 
 'Si le ListView n'est pas vide
 If lvwPunch.ListItems.count > 0 Then
 'Pour aller chercher l'index du "punch" cliqué
 m_lIDPunch = Item.Tag
 
 Set rstPunch = New ADODB.Recordset
 
 Call rstPunch.Open("SELECT * FROM GrbPunch WHERE IDPunch = " & m_lIDPunch, g_connData, adOpenDynamic, adLockOptimistic)
 
 If Left(rstPunch.Fields("NoProjet"), 1) = "E" Then
 optTypePunch(I_OPT_ELECTRIQUE).Value = True
 Else
 optTypePunch(I_OPT_MECANIQUE).Value = True
  End If
 
  m_bClick = True

  mskNoProjet.Text = Right(rstPunch.Fields("NoProjet"), 8)

  m_bClick = False
 
  If mskDate.mask = "##-##-##" Then
  mskDate.mask = vbNullString
  End If
 
  mskDate.Text = rstPunch.Fields("Date")
 
mskHeureDebut.Text = rstPunch.Fields("HeureDébut")
 
1 If Not IsNull(rstPunch.Fields("HeureFin")) Then
 If rstPunch.Fields("HeureFin") <> "" Then
 mskHeureFin.Text = rstPunch.Fields("HeureFin")
 Else
 mskHeureFin.Text = "__:__"
 End If
 Else
 mskHeureFin.Text = "__:__"
 End If

 txtClient.Text = vbNullString

 Call AfficherClient

If Not IsNull(rstPunch.Fields("Type")) Then
 cmbType.ListIndex = -1
 
 If rstPunch.Fields("Type") = "Installation" Then
 cmbType.ListIndex = 0
 GoTo Fin_De_Type
 End If
 If rstPunch.Fields("Type") = "MiseService" Then
 cmbType.ListIndex = 1
 GoTo Fin_De_Type
 End If
 
 
 If Left(rstPunch.Fields("NoProjet"), 1) = "E" Then
 For G = 0 To cmbType.ListCount
 If cmbType.LIST(G) = rstPunch.Fields("Type") Then
 cmbType.ListIndex = G
 Exit For
 End If
 Next
 Else
 For G = 0 To cmbType.ListCount
 If cmbType.LIST(G) = rstPunch.Fields("Type") Then
 cmbType.ListIndex = G
 Exit For
 End If
 Next
 End If
315 Fin_De_Type:
 End If

 If Not IsNull(rstPunch.Fields("Commentaire")) Then
 txtCommentaires.Text = rstPunch.Fields("Commentaire")
 Else
 txtCommentaires.Text = vbNullString
 End If

 If rstPunch.Fields("KM") = True Then
 chkKM.Value = vbChecked

 txtKM.Text = rstPunch.Fields("NbreKM")
 Else
 chkKM.Value = vbUnchecked
 txtKM.Text = vbNullString
End If

 Call rstPunch.Close
 Set rstPunch = Nothing
 
 Call ActiverControles(MODE_MODIF)
40 End If

Exit Sub

Oups:

4 wOups "frmFeuilleTemps", "lvwPunch_ItemClick", Err, Err.number, Err.Description
End Sub

Private Sub lvwPunch_KeyDown(KeyCode As Integer, Shift As Integer)

 On Error GoTo Oups
 'Pour permettre d'effacer un enregistrement en appuyant sur "Delete" ou "Suppr"
 
 'Si il y a un enregistrement dans le listview
 If lvwPunch.ListItems.count > 0 Then
 'Si la touche appuyée est "Delete"
 If KeyCode = vbKeyDelete Then
 'Si l'utilisateur répond oui à "Etes vous sur?"
 If MsgBox("Voulez-vous vraiment effacer le punch?", vbYesNo) = vbYes Then
 'Efface
 Call g_connData.Execute("DELETE * FROM GrbPunch WHERE IDPunch = " & lvwPunch.SelectedItem.Tag)
 
 Call RemplirListView
 
 Call ViderChamps
 
 Call ActiverControles(MODE_INACTIF)
 
 If lvwPunch.ListItems.count > 0 Then
 'Sélection du premier ListItem
 lvwPunch.ListItems(1).Selected = True
 Call lvwPunch_ItemClick(lvwPunch.SelectedItem)
  End If
  End If
  End If
  End If

  Exit Sub

Oups:

  wOups "frmFeuilleTemps", "lvwPunch_KeyDown", Err, Err.number, Err.Description
End Sub

Private Sub mskNoProjet_Change()

 On Error GoTo Oups

 If InStr(1, mskNoProjet.Text, "_") = 0 Then
 Call AfficherClient
 Else
 txtClient.Text = vbNullString
 End If

 Call RemplirComboType

 Exit Sub

Oups:

 wOups "frmFeuilleTemps", "mskNoProjet_Change", Err, Err.number, Err.Description
End Sub

Private Sub AfficherClient()

 On Error GoTo Oups

 Dim rstProjSoum As ADODB.Recordset
 Dim rstClient As ADODB.Recordset
 Dim iCompteur As Integer
 
 If m_bClick = False Then
 Set rstProjSoum = New ADODB.Recordset
 Set rstClient = New ADODB.Recordset
 
 Call rstProjSoum.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & lblPrefixe.Caption & mskNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 If Not rstProjSoum.EOF Then
 Call rstClient.Open("SELECT NomClient FROM GrbClient WHERE IDClient = " & rstProjSoum.Fields("NoClient"), g_connData, adOpenDynamic, adLockOptimistic)
 
 txtClient.Text = rstClient.Fields("NomClient")
  txtClient.Tag = rstProjSoum.Fields("NoClient")
 
  Call rstClient.Close
  Set rstClient = Nothing
 
  If rstProjSoum.Fields("Ouvert") = False Then
  Call MsgBox("Ce projet n'est pas ouvert!", vbOKOnly, "Erreur")
  End If
  Else
  Call MsgBox("Projet inexistant!", vbOKOnly, "Erreur")
End If
 
1 Call rstProjSoum.Close
 Set rstProjSoum = Nothing
End If

Exit Sub

Oups:

wOups "frmFeuilleTemps", "AfficherClient", Err, Err.number, Err.Description
End Sub

Private Sub mvwDate_DateClick(ByVal DateClicked As Date)

 On Error GoTo Oups

 mskDate.Text = ConvertDate(DateClicked)
 
 'Enlever le calendrier
 mvwDate.Visible = False

 Exit Sub

Oups:

 wOups "frmFeuilleTemps", "mvwDate_DateClick", Err, Err.number, Err.Description
End Sub

Private Sub mvwDate_LostFocus()

 On Error GoTo Oups
 'Quand le calendrier perd le focus, il faut l'enlever

 mvwDate.Visible = False

 Exit Sub

Oups:

 wOups "frmFeuilleTemps", "mvwDate_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub cmdImprimer_Click()
 
 On Error GoTo Oups

 Call frmChoixImpressionFT.Afficher(m_datSemaine, GetLastDay(m_datSemaine))

 Exit Sub

Oups:

 wOups "frmFeuilleTemps", "cmdImprimer_Click", Err, Err.number, Err.Description
End Sub

Private Sub mskDate_GotFocus()

 On Error GoTo Oups
 
 'Met l'année sur 2 chiffres
 If Len(mskDate.Text) = 10 Then
 mskDate.Text = Right$(mskDate.Text, 8)
 End If
 
 mskDate.mask = "##-##-##"

 Exit Sub

Oups:

 wOups "frmFeuilleTemps", "mskDate_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskHeureDebut_GotFocus()

 On Error GoTo Oups
 
 'Format d'heure
 mskHeureDebut.mask = "##:##"

 Exit Sub

Oups:

 wOups "frmFeuilleTemps", "mskHeureDebut_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskHeureFin_GotFocus()

 On Error GoTo Oups
 
 'Format d'heure
 mskHeureFin.mask = "##:##"

 Exit Sub

Oups:

 wOups "frmFeuilleTemps", "mskHeureFin_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskDate_LostFocus()

 On Error GoTo Oups
 
 'Enlève le mask
 mskDate.mask = vbNullString
 
 'Vide le champs si l'utilisateur n'a rien écrit
 If mskDate.Text = "__-__-__" Then
 mskDate.Text = vbNullString
 Else
 'Remet l'année sur   chiffres
 If Len(mskDate.Text) =   Then
 If IsDate(mskDate.Text) Then
 mskDate.Text = Year(DateSerial(Left$(mskDate, 2), Mid$(mskDate, 4, 2), Right$(mskDate.Text, 2))) & Mid$(mskDate.Text, 3, 8)
 End If
 End If
 End If

  Exit Sub

Oups:

  wOups "frmFeuilleTemps", "mskDate_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskHeureDebut_LostFocus()
 Dim heure As Date
 On Error GoTo Oups
 
 'Enlève le mask
 mskHeureDebut.mask = vbNullString
 
 'Vide le champs si l'utilisateur n'a rien écrit
 If mskHeureDebut.Text = "__:__" Then
 mskHeureDebut.Text = vbNullString
25
 Else

 heure = CDate(mskHeureDebut.Text)
 
 If Minute(heure) <= 5 Then
 heure = TimeSerial(Hour(heure), 0, 0)
 Else
 If Minute(heure) <= 24 Then
 heure = TimeSerial(Hour(heure), 15, 0)
 Else
 If Minute(heure) <= 35 Then
 heure = TimeSerial(Hour(heure), 30, 0)
 Else
 If Minute(heure) <= 54 Then
 heure = TimeSerial(Hour(heure), 45, 0)
 Else
  heure = TimeSerial(Hour(heure) + 1, 0, 0)
 End If
 End If
 End If
 End If
 
  mskHeureDebut.Text = Right$("0" & Hour(heure), 2) + ":" + Right$("0" & Minute(heure), 2)
 
 
 
 End If
 
 
 
 
 
  Exit Sub

Oups:

  wOups "frmFeuilleTemps", "mskHeureDebut_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskHeureFin_LostFocus()
 Dim heure As Date
 On Error GoTo Oups
 
 'Enlève le mask
 mskHeureFin.mask = vbNullString
 
 'Vide le champs si l'utilisateur n'a rien écrit
 If mskHeureFin.Text = "__:__" Then
 mskHeureFin.Text = vbNullString
 Else
 heure = CDate(mskHeureFin.Text)
 
 If Minute(heure) <= 5 Then
 heure = TimeSerial(Hour(heure), 0, 0)
 Else
 If Minute(heure) <= 24 Then
 heure = TimeSerial(Hour(heure), 15, 0)
 Else
 If Minute(heure) <= 35 Then
 heure = TimeSerial(Hour(heure), 30, 0)
 Else
 If Minute(heure) <= 54 Then
 heure = TimeSerial(Hour(heure), 45, 0)
 Else
  heure = TimeSerial(Hour(heure) + 1, 0, 0)
 End If
 End If
 End If
 End If
 
  mskHeureFin.Text = Right$("0" & Hour(heure), 2) + ":" + Right$("0" & Minute(heure), 2)



 End If

  Exit Sub

Oups:

  wOups "frmFeuilleTemps", "mskHeureFin_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtKM_LostFocus()

 On Error GoTo Oups

 txtKM.Text = Replace(txtKM.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmFeuilleTemps", "txtKM_LostFocus", Err, Err.number, Err.Description
End Sub

Public Sub RemplirComboType()
 
 On Error GoTo Oups

 Dim bInstallation As Boolean
 Dim bTypeInutile As Boolean
 Dim tbltype As ADODB.Recordset
 Set tbltype = New ADODB.Recordset
 
 Call cmbType.Clear

 If Mid$(mskNoProjet.Text, 2, 1) = "1" Then
 bTypeInutile = True
 End If

 If IsNumeric(Right$(mskNoProjet.Text, 2)) Then
 If Mid$(mskNoProjet.Text, 2, 4) <> "3000" Then
 If CInt(Right$(mskNoProjet.Text, 2)) >= 51 And CInt(Right$(mskNoProjet.Text, 2)) <= 5 Then
 bInstallation = True
  End If
  Else
  bTypeInutile = True
  End If
  End If
 
  If bTypeInutile = False Then
  lblType.Visible = True
  cmbType.Visible = True

If bInstallation = True Then
If lblPrefixe.Caption = "E" Then
 Call cmbType.AddItem("Installation")
 Call cmbType.AddItem("Mise en service")
 Else
 Call cmbType.AddItem("Installation")
 End If
 Else
 If lblPrefixe.Caption = "E" Then
 Call tbltype.Open("select * from TBL_Punch_Type Where Mode = 'E' Order by name", g_connData, adOpenDynamic, adLockOptimistic)
 Do While Not tbltype.EOF
 cmbType.AddItem (tbltype.Fields("Name"))
 Call tbltype.MoveNext
 Loop
 Call tbltype.Close
 Set tbltype = Nothing
 
 
 
 
 Else
205
 Call tbltype.Open("select * from TBL_Punch_Type Where Mode = 'M' Order by name ", g_connData, adOpenDynamic, adLockOptimistic)
 Do While Not tbltype.EOF
 cmbType.AddItem (tbltype.Fields("Name"))
 Call tbltype.MoveNext
 Loop
 Call tbltype.Close
 Set tbltype = Nothing
 End If
End If
Else
lblType.Visible = False
 cmbType.Visible = False
2  End If

Exit Sub

Oups:

2  wOups "frmFeuilleTemps", "RemplirComboType", Err, Err.number, Err.Description
End Sub

