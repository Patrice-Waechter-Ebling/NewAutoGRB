VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFeuilleTemps 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Feuilles de temps"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmFeuilleTemps.frx":0000
   ScaleHeight     =   7845
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdModifier 
      Caption         =   "Modifier Type"
      Height          =   495
      Left            =   6000
      TabIndex        =   34
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmdexcel 
      Caption         =   "Excel"
      Height          =   495
      Left            =   2450
      TabIndex        =   33
      Top             =   7200
      Width           =   1095
   End
   Begin MSComCtl2.MonthView mvwDate 
      Height          =   2370
      Left            =   960
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   0
      Appearance      =   1
      ShowToday       =   0   'False
      StartOfWeek     =   90243073
      CurrentDate     =   37735
   End
   Begin VB.OptionButton optTypePunch 
      BackColor       =   &H00000000&
      Caption         =   "Mécanique"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   30
      Top             =   4800
      Width           =   1095
   End
   Begin VB.OptionButton optTypePunch 
      BackColor       =   &H00000000&
      Caption         =   "Électrique"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   29
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdExporter 
      Caption         =   "Exporter"
      Height          =   495
      Left            =   120
      TabIndex        =   28
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CheckBox chkKM 
      BackColor       =   &H00000000&
      Caption         =   "KM :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   19
      Top             =   6600
      Width           =   735
   End
   Begin VB.TextBox txtSemaine 
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdDateSemaine 
      Caption         =   "..."
      Height          =   315
      Left            =   7920
      TabIndex        =   4
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdImprimer 
      Caption         =   "Imprimer"
      Height          =   495
      Left            =   1290
      TabIndex        =   23
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmdDate 
      Caption         =   "..."
      Height          =   255
      Left            =   2880
      TabIndex        =   13
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox txtCommentaires 
      Height          =   765
      Left            =   4560
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   6120
      Width           =   3735
   End
   Begin VB.ComboBox cmbEmployé 
      Height          =   315
      Left            =   840
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton cmdEnregistrer 
      Caption         =   "Enregistrer"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4800
      TabIndex        =   25
      Top             =   7200
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwPunch 
      Height          =   3015
      Left            =   120
      TabIndex        =   5
      Top             =   1440
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
      Appearance      =   1
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
      Top             =   7200
      Width           =   1095
   End
   Begin MSMask.MaskEdBox mskHeureFin 
      Height          =   255
      Left            =   1560
      TabIndex        =   18
      Top             =   6240
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
      Top             =   5520
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
      Top             =   5880
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
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      Height          =   495
      Left            =   7200
      TabIndex        =   26
      Top             =   7200
      Width           =   1095
   End
   Begin MSMask.MaskEdBox mskNoProjet 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Top             =   5160
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
      Top             =   6600
      Width           =   615
   End
   Begin VB.TextBox txtClient 
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   4920
      Width           =   3735
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   5520
      Width           =   3735
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   "Type :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   32
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label lblPrefixe 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   31
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Km"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   22
      Top             =   6600
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Client :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   9
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label lblSemaine 
      BackStyle       =   0  'Transparent
      Caption         =   "Semaine du :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Commentaires :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   16
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Employé :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Heure de fin :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Heure de début :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date (AA-MM-JJ):"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Numéro de projet :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5160
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
Private Const I_TYPE_ELEC_INSTALLATION   As Integer = 0
Private Const I_TYPE_ELEC_MISE_SERVICE   As Integer = 1

'Types quand c'est pas un 51
Private Const I_TYPE_ELEC_DESSIN         As Integer = 0
Private Const I_TYPE_ELEC_FABRICATION    As Integer = 1
Private Const I_TYPE_ELEC_ASSEMBLAGE     As Integer = 2
Private Const I_TYPE_ELEC_PROG_INTERFACE As Integer = 3
Private Const I_TYPE_ELEC_PROG_AUTOMATE  As Integer = 4
Private Const I_TYPE_ELEC_PROG_ROBOT     As Integer = 5
Private Const I_TYPE_ELEC_VISION         As Integer = 6
Private Const I_TYPE_ELEC_TEST           As Integer = 7
Private Const I_TYPE_ELEC_FORMATION      As Integer = 8
Private Const I_TYPE_ELEC_GESTION        As Integer = 9
Private Const I_TYPE_ELEC_SHIPPING       As Integer = 10
Private Const I_TYPE_ELEC_prototypage       As Integer = 11

'Types quand c'est un 51
Private Const I_TYPE_MEC_INSTALLATION    As Integer = 0

'Types quand c'est pas un 51
Private Const I_TYPE_MEC_DESSIN          As Integer = 0
Private Const I_TYPE_MEC_COUPE           As Integer = 1
Private Const I_TYPE_MEC_MACHINAGE       As Integer = 2
Private Const I_TYPE_MEC_SOUDURE         As Integer = 3
Private Const I_TYPE_MEC_ASSEMBLAGE      As Integer = 4
Private Const I_TYPE_MEC_PEINTURE        As Integer = 5
Private Const I_TYPE_MEC_TEST            As Integer = 6
Private Const I_TYPE_MEC_FORMATION       As Integer = 7
Private Const I_TYPE_MEC_GESTION         As Integer = 8
Private Const I_TYPE_MEC_SHIPPING        As Integer = 9
Private Const I_TYPE_MEC_prototypage        As Integer = 10

Private Const I_LVW_PROJET               As Integer = 0
Private Const I_LVW_DATE                 As Integer = 1
Private Const I_LVW_DEBUT                As Integer = 2
Private Const I_LVW_FIN                  As Integer = 3
Private Const I_LVW_CLIENT               As Integer = 4
Private Const I_LVW_TYPE                 As Integer = 5
Private Const I_LVW_COMMENTAIRE          As Integer = 6

Private Const I_OPT_ELECTRIQUE           As Integer = 0
Private Const I_OPT_MECANIQUE            As Integer = 1

Private Enum enumMode
  MODE_INACTIF = 0
  MODE_MODIF = 1
  MODE_AJOUT = 2
End Enum

Private m_lIDPunch   As Long
Public m_datSemaine  As Date
Private m_bModifProj As Boolean
Private m_eMode      As enumMode
Private m_bClick     As Boolean

Private Sub ActiverControles(ByVal eMode As enumMode)

5       On Error GoTo AfficherErreur

        'Activation des controles dépendamment du mode choisi
10      Dim bListView    As Boolean 'Pour le ListView
15      Dim bEmploye     As Boolean 'Pour la liste des employés
20      Dim bSemaine     As Boolean 'Pour le bouton de la semaine
25      Dim bChamps      As Boolean 'Pour les champs
30      Dim bImprimer    As Boolean 'Pour le bouton "Imprimer"
35      Dim bAjouter     As Boolean 'Pour le bouton "Ajouter"
40      Dim bAnnuler     As Boolean 'Pour le bouton "Annuler"
45      Dim bEnregistrer As Boolean 'Pour le bouton "Enregistrer"
50      Dim bFermer      As Boolean 'Pour le bouton "Fermer"
    
55      m_eMode = eMode
    
60      Select Case eMode
          'Mode pour ouverture et après ajout et modif
          Case MODE_INACTIF:
65          bListView = True
70          bEmploye = True
75          bSemaine = True
80          bImprimer = True
85          bAjouter = True
90          bFermer = True
      
          'Mode pour la modification
95        Case MODE_MODIF:
100         bListView = True
105         bEmploye = True
110         bSemaine = True
115         bChamps = True
120         bImprimer = True
125         bAjouter = True
130         bAnnuler = True
135         bEnregistrer = True
140         bFermer = True
      
          'Mode pour l'ajout
145       Case MODE_AJOUT:
150         bChamps = True
155         bAnnuler = True
160         bEnregistrer = True
165     End Select
      
170     lvwPunch.Enabled = bListView
175     cmbemployé.Enabled = bEmploye
180     cmdDateSemaine.Enabled = bSemaine
185     mskNoProjet.Enabled = bChamps
190     mskDate.Enabled = bChamps
195     cmdDate.Enabled = bChamps
200     mskHeureDebut.Enabled = bChamps
205     mskHeureFin.Enabled = bChamps
210     txtCommentaires.Enabled = bChamps
215     cmbType.Enabled = bChamps
220     optTypePunch(I_OPT_ELECTRIQUE).Enabled = bChamps
225     optTypePunch(I_OPT_MECANIQUE).Enabled = bChamps

230     If g_bModificationFeuillesTemps = True Or (cmbemployé.Text = g_sEmploye) Then
235       cmdEnregistrer.Enabled = bEnregistrer
          cmdexcel.Enabled = bImprimer
240       cmdImprimer.Enabled = bImprimer
245       Cmdajouter.Enabled = bAjouter
250     Else
255       cmdEnregistrer.Enabled = False
          cmdexcel.Enabled = False
260       cmdImprimer.Enabled = False
265       Cmdajouter.Enabled = False
270     End If

275     cmdAnnuler.Visible = bAnnuler
280     Cmdfermer.Visible = bFermer

285     Exit Sub

AfficherErreur:

290     woups "frmFeuilleTemps", "ActiverControles", Err, Erl
End Sub

Private Sub chkKM_Click()

5       On Error GoTo AfficherErreur

10      If chkKM.Value = vbChecked Then
15        txtKM.Enabled = True
20      Else
25        txtKM.Text = ""
30        txtKM.Enabled = False
35      End If

40      Exit Sub

AfficherErreur:

45      woups "frmFeuilleTemps", "chkKM_Click", Err, Erl
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

With xlsheet.Range("A1;D1;A3:G3")
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
xlsheet.Range("A:G").Columns.AutoFit







xlsheet.Visible = True

Set xlsheet = Nothing



End Sub

Private Sub cmdModifier_Click()
Call FrmModType.Show(vbModal)



End Sub

Private Sub optTypePunch_Click(Index As Integer)
        
5       On Error GoTo AfficherErreur

10      If Index = I_OPT_ELECTRIQUE Then
15        lblPrefixe.Caption = "E"
20      Else
25        lblPrefixe.Caption = "M"
30      End If
        
35      Call RemplirComboType

40      Exit Sub

AfficherErreur:

45      woups "frmFeuilleTemps", "optTypePunch_Click", Err, Erl
End Sub

Private Sub Cmdajouter_Click()

5       On Error GoTo AfficherErreur

10      Call ViderChamps

15      Call ActiverControles(MODE_AJOUT)

20      Exit Sub

AfficherErreur:

25      woups "frmFeuilleTemps", "cmdAjouter_Click", Err, Erl
End Sub

Private Sub cmdAnnuler_Click()

5       On Error GoTo AfficherErreur

        'Annuler l'ajout ou la modif
10      Call ViderChamps
  
15      Call ActiverControles(MODE_INACTIF)

20      Exit Sub

AfficherErreur:

25      woups "frmFeuilleTemps", "cmdAnnuler_Click", Err, Erl
End Sub

Private Sub cmdDateSemaine_Click()

5       On Error GoTo AfficherErreur

        'Affichage du calendrier pour choisir une semaine
10      Call OuvrirForm(frmFeuilleTempsCalendrier, True)

15      Exit Sub

AfficherErreur:

20      woups "frmFeuilleTemps", "cmdDateSemaine_Click", Err, Erl
End Sub

Private Sub cmbEmployé_Click()

5       On Error GoTo AfficherErreur

        'Rempli le listview si la semaine a été choisie
10      Call ViderChamps
  
15      Call ActiverControles(MODE_INACTIF)
  
        'Il faut remplir le listview dépendant la semaine
20      If txtSemaine.Text <> vbNullString Then
25        Call RemplirListView
30      End If
  
35      cmdEnregistrer.Enabled = False

40      Exit Sub

AfficherErreur:

45      woups "frmFeuilleTemps", "cmbEmployé_Click", Err, Erl
End Sub

Private Sub cmdDate_Click()

5       On Error GoTo AfficherErreur
        'Ouverture du calendrier
  
        'Si il y a une date valide, la date par défaut est celle-ci, sinon c'est la date
        'd'aujourd'hui
10      If Trim$(mskDate.Text) <> vbNullString Then
15        If ValiderDate(mskDate.Text) = True Then
20          mvwDate.Value = mskDate.Text
25        Else
30          mvwDate.Value = Date
35        End If
40      Else
45        mvwDate.Value = Date
50      End If
  
55      mvwDate.Visible = True
  
60      Call mvwDate.SetFocus

65      Exit Sub

AfficherErreur:

70      woups "frmFeuilleTemps", "cmdDate_Click", Err, Erl
End Sub

Private Sub cmdEnregistrer_Click()

5       On Error GoTo AfficherErreur
        
        'Enregistrer l'ajout ou la modif
10      Dim rstPunch      As ADODB.Recordset
15      Dim rstProjSoum   As ADODB.Recordset
20      Dim bInstallation As Boolean
  
25      Set rstProjSoum = New ADODB.Recordset

30      Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & lblPrefixe.Caption & mskNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
35      If Not rstProjSoum.EOF Then
40        If rstProjSoum.Fields("Ouvert") = False Then
45          Call MsgBox("Ce projet est fermé!", vbOKOnly, "Erreur")
        
50          Call rstProjSoum.Close
55          Set rstProjSoum = Nothing
        
60          Exit Sub
65        End If
70      Else
75        Call MsgBox("Projet inexistant!", vbOKOnly, "Erreur")
      
80        Call rstProjSoum.Close
85        Set rstProjSoum = Nothing
      
90        Exit Sub
95      End If

100     Call rstProjSoum.Close
105     Set rstProjSoum = Nothing
  
        'Valider l'heure de début
110     If ValiderHeure(mskHeureDebut.Text) = False Then
115       Call MsgBox("L'heure de début est invalide!", vbOKOnly, "Erreur")
    
120       Exit Sub
125     Else
130       If mskHeureFin.Text <> "24:00" Then
135         If mskHeureFin.Text <> vbNullString Then
              'Valider l'heure de fin
140           If ValiderHeure(mskHeureFin.Text) = False Then
145             Call MsgBox("L'heure de fin est invalide!", vbOKOnly, "Erreur")
         
150             Exit Sub
155           End If
160         End If
165       End If
170     End If
  
        'Valider la date
175     If ValiderDate(mskDate.Text) = False Then
180       Call MsgBox("La date est invalide!", vbOKOnly, "Erreur")
    
185       Exit Sub
190     End If
  
        'Si les champs importants ont été rempli
195     If mskNoProjet.Text = vbNullString Or mskDate.Text = vbNullString Or mskHeureDebut.Text = vbNullString Then
200       Call MsgBox("Champs vide!", vbOKOnly, "Erreur")
    
205       Exit Sub
210     End If
  
        'Si les champs importants ont été rempli
215     If Right$(mskNoProjet.Text, 1) = "_" Then
220       Call MsgBox("Le numéro de projet est invalide!", vbOKOnly, "Erreur")
    
225       Exit Sub
230     End If

235     If cmbType.Text = "" And cmbType.Visible = True Then
240       Call MsgBox("Le type est obligatoire!", vbOKOnly, "Erreur")

245       Exit Sub
250     End If
  
255     Set rstPunch = New ADODB.Recordset
  
260     If m_eMode = MODE_AJOUT Then
265       Call rstPunch.Open("SELECT * FROM  GRB_Punch", g_connData, adOpenDynamic, adLockOptimistic)
      
270       Call rstPunch.AddNew
275     Else
280       Call rstPunch.Open("SELECT * FROM GRB_Punch WHERE IDPunch = " & m_lIDPunch, g_connData, adOpenDynamic, adLockOptimistic)
285     End If
  
290     rstPunch.Fields("NoEmploye") = cmbemployé.ItemData(cmbemployé.ListIndex)
    
295     rstPunch.Fields("NoProjet") = lblPrefixe.Caption & mskNoProjet.Text
300     rstPunch.Fields("Date") = mskDate.Text
305     rstPunch.Fields("HeureDébut") = mskHeureDebut.Text
310     rstPunch.Fields("HeureFin") = mskHeureFin.Text

315     If chkKM.Value = vbChecked Then
320       rstPunch.Fields("KM") = True

325       If txtKM.Text <> "" Then
330         txtKM.Text = Replace(txtKM.Text, ".", ",")

335         If IsNumeric(txtKM.Text) Then
340           rstPunch.Fields("NbreKM") = txtKM.Text
345         Else
350           rstPunch.Fields("NbreKM") = 0
355         End If
360       Else
365         txtKM.Text = 0
370       End If
375     Else
380       rstPunch.Fields("KM") = False
385       rstPunch.Fields("NbreKM") = ""
390     End If
     
395     If txtClient.Text <> vbNullString Then
400       rstPunch.Fields("NoClient") = txtClient.Tag
405     Else
410       rstPunch.Fields("NoClient") = vbNullString
415     End If

420     If IsNumeric(Right$(mskNoProjet.Text, 2)) Then
425       If CInt(Right$(mskNoProjet.Text, 2)) >= 51 And CInt(Right$(mskNoProjet.Text, 2)) <= 59 Then
430         bInstallation = True
435       Else
440         bInstallation = False
445       End If
450     Else
455       bInstallation = False
460     End If
  
465     If bInstallation = True Then
470       If lblPrefixe.Caption = "E" Then
475         Select Case cmbType.ListIndex
              Case I_TYPE_ELEC_INSTALLATION: rstPunch.Fields("Type") = "Installation"
480           Case I_TYPE_ELEC_MISE_SERVICE: rstPunch.Fields("Type") = "MiseService"
485         End Select
490       Else
495         Select Case cmbType.ListIndex
              Case I_TYPE_MEC_INSTALLATION: rstPunch.Fields("Type") = "Installation"
500         End Select
505       End If
510     Else
515       If lblPrefixe.Caption = "E" Then
            rstPunch.Fields("Type") = cmbType.Text
580       Else
            rstPunch.Fields("Type") = cmbType.Text
640       End If
645     End If
  
650     rstPunch.Fields("Commentaire") = txtCommentaires.Text
    
655     Call rstPunch.Update
    
660     Call rstPunch.Close
665     Set rstPunch = Nothing
    
670     Call ViderChamps
    
675     Call ActiverControles(MODE_INACTIF)
    
680     Call RemplirListView

685     Exit Sub

AfficherErreur:

690     woups "frmFeuilleTemps", "cmdEnregistrer_Click", Err, Erl
End Sub

Private Function ValiderDate(ByVal sDate As String) As Boolean

5       On Error GoTo AfficherErreur

        'Validation des dates
10      If Not IsDate(sDate) Then
15        ValiderDate = False
20      Else
25        ValiderDate = True
30      End If

35      Exit Function

AfficherErreur:

40      woups "frmFeuilleTemps", "ValiderDate", Err, Erl
End Function

Private Function ValiderHeure(ByVal sHeure As String) As Boolean

5       On Error GoTo AfficherErreur

        'validation des heures
10      Dim sHour   As String
15      Dim sMinute As String
20      Dim sSecond As String
  
25      sHour = Left$(sHeure, 2)
  
30      sMinute = Mid$(sHeure, 4, 2)
  
35      sSecond = Right$(sHeure, 2)
  
40      ValiderHeure = True
  
        'Si numérique
45      If Not IsNumeric(sHour) Then
50        ValiderHeure = False
    
55        Exit Function
60      Else
          'Si entre 0 et 23
65        If sHour < 0 Or sHour > 23 Then
70          ValiderHeure = False

75          Exit Function
80        End If
85      End If
  
        'Si numérique
90      If Not IsNumeric(sMinute) Then
95        ValiderHeure = False
    
100       Exit Function
105     Else
          'Si entre 0 et 59
110       If sMinute < 0 Or sMinute > 59 Then
115         ValiderHeure = False
      
120         Exit Function
125       End If
130     End If

        'Si numérique
135     If Not IsNumeric(sSecond) Then
140       ValiderHeure = False
    
145       Exit Function
150     Else
                'Si entre 0 et 59
155       If sSecond < 0 Or sSecond > 59 Then
160         ValiderHeure = False
      
165         Exit Function
170       End If
175     End If

180     Exit Function

AfficherErreur:

185     woups "frmFeuilleTemps", "ValiderHeure", Err, Erl
End Function

Private Sub ViderChamps()

5       On Error GoTo AfficherErreur

        'Vider les champs
10      mskNoProjet.Text = "_____-__"
  
15      mskDate.Text = vbNullString
20      mskHeureDebut.Text = vbNullString
25      mskHeureFin.Text = vbNullString
30      txtCommentaires.Text = vbNullString
35      txtClient.Text = vbNullString
40      chkKM.Value = vbUnchecked
45      cmbType.ListIndex = -1

50      txtKM.Text = vbNullString

55      Exit Sub

AfficherErreur:

60      woups "frmFeuilleTemps", "ViderChamps", Err, Erl
End Sub

Private Sub cmdexporter_Click()

5       On Error GoTo AfficherErreur

10      Call frmChoixDateImpressionFT.Show(vbModal)

15      Exit Sub

AfficherErreur:

20      woups "frmFeuilleTemps", "cmdExporter_Click", Err, Erl
End Sub

Private Sub Cmdfermer_Click()

5       On Error GoTo AfficherErreur

        'Fermeture de la fenêtre
10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmFeuilleTemps", "cmdFermer_Click", Err, Erl
End Sub

Public Sub RemplirListView()

5       On Error GoTo AfficherErreur
        
        'Rempli le listview selon la semaine et l'employé choisi
10      Dim rstPunch      As ADODB.Recordset
15      Dim rstNomClient  As ADODB.Recordset
20      Dim itmPunch      As ListItem
25      Dim sTemp         As String
30      Dim iTemp         As Integer
35      Dim sDateDebut    As String
40      Dim sDateFin      As String
    
        'Il faut vider le ListView pour ne pas le remplir plein de fois
45      Call lvwPunch.ListItems.Clear
  
50      sDateDebut = ConvertDate(m_datSemaine)
55      sDateFin = ConvertDate(GetLastDay(m_datSemaine))

60      Set rstNomClient = New ADODB.Recordset
65      Set rstPunch = New ADODB.Recordset

70      Call rstPunch.Open("SELECT * FROM GRB_Punch WHERE noEmploye = " & cmbemployé.ItemData(cmbemployé.ListIndex) & " AND Date BETWEEN '" & sDateDebut & "' AND '" & sDateFin & "' ORDER BY Date, HeureDébut", g_connData, adOpenDynamic, adLockOptimistic)
    
75      Do While Not rstPunch.EOF
80        Set itmPunch = lvwPunch.ListItems.Add
      
          'Numéro du projet
85        itmPunch.Text = rstPunch.Fields("noProjet")
      
          'IDPunch dans le tag
90        itmPunch.Tag = rstPunch.Fields("IDPunch")
          
          'Date
95        itmPunch.SubItems(I_LVW_DATE) = rstPunch.Fields("Date")
      
          'Heure de début
100       itmPunch.SubItems(I_LVW_DEBUT) = rstPunch.Fields("HeureDébut")
      
          'Heure de fin
105       If Not IsNull(rstPunch.Fields("HeureFin")) Then
110         itmPunch.SubItems(I_LVW_FIN) = rstPunch.Fields("HeureFin")
115       Else
120         itmPunch.SubItems(I_LVW_FIN) = vbNullString
125       End If
      
130       If Not IsNull(rstPunch.Fields("NoClient")) And rstPunch.Fields("NoClient") > "0" Then
135         Call rstNomClient.Open("SELECT NomClient FROM GRB_Client WHERE IDClient = " & rstPunch.Fields("NoClient"), g_connData, adOpenDynamic, adLockOptimistic)
      
140         itmPunch.SubItems(I_LVW_CLIENT) = rstNomClient.Fields("NomClient")
     
145         Call rstNomClient.Close
150       Else
155         itmPunch.SubItems(I_LVW_CLIENT) = vbNullString
160       End If

165       If Not IsNull(rstPunch.Fields("Type")) Then
170         If Left$(itmPunch.Text, 1) = "E" Then
                itmPunch.SubItems(I_LVW_TYPE) = rstPunch.Fields("Type")
245         Else
                 itmPunch.SubItems(I_LVW_TYPE) = rstPunch.Fields("Type")
310         End If
315       End If

320       If Not IsNull(rstPunch.Fields("Commentaire")) Then
325         itmPunch.SubItems(I_LVW_COMMENTAIRE) = rstPunch.Fields("Commentaire")
330       Else
335         itmPunch.SubItems(I_LVW_COMMENTAIRE) = vbNullString
340       End If
      
345       Call rstPunch.MoveNext
350     Loop
    
355     Call rstPunch.Close
360     Set rstPunch = Nothing

365     Set rstNomClient = Nothing

370     Exit Sub

AfficherErreur:

375     woups "frmFeuilleTemps", "RemplirListView", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      Dim datTemp As Date
15      Dim iNoJour As Integer
        Call frmPunch.Table_exist
20      Call RemplirComboEmploye

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
        
        
25      datTemp = Date

30      iNoJour = 1

        'Pour avoir la date de la semaine précédente
35      Do While iNoJour < 8
40        datTemp = datTemp - TimeSerial(24, 0, 0)

45        iNoJour = iNoJour + 1
50      Loop
  
55      m_datSemaine = GetFirstDay(datTemp)

        'On affiche le dimanche de la semaine précédente
60      txtSemaine.Text = GetDateTexte(m_datSemaine)
    
65      Call RemplirListView

70      Screen.MousePointer = vbDefault

75      Exit Sub

AfficherErreur:

80      woups "frmFeuilleTemps", "Form_Load", Err, Erl
End Sub

Private Sub RemplirComboEmploye()

5       On Error GoTo AfficherErreur
        
        'Rempli le combo des employés
10      Dim rstEmploye As ADODB.Recordset
  
        'Il faut vider le combo pour ne pas le remplir plein de fois
15      Call cmbemployé.Clear
  
20      Set rstEmploye = New ADODB.Recordset
  
25      Call rstEmploye.Open("SELECT * FROM GRB_Employés WHERE Actif = true", g_connData, adOpenDynamic, adLockOptimistic)
    
30      Do While Not rstEmploye.EOF
          'Ajout du nom de l'employé dans le combo
35        Call cmbemployé.AddItem(rstEmploye.Fields("employe"))
      
          'Ajout du numéro de l'employé dans l'ItemData
40        cmbemployé.ItemData(cmbemployé.newIndex) = rstEmploye.Fields("noemploye")
      
45        Call rstEmploye.MoveNext
50      Loop
    
55      Call rstEmploye.Close
60      Set rstEmploye = Nothing
  
        'Si le combo n'est pas vide, on sélectionne le premier
65      If cmbemployé.ListCount > 0 Then
70        cmbemployé.ListIndex = 0
75      End If

80      Exit Sub

AfficherErreur:

85      woups "frmFeuilleTemps", "RemplirComboEmploye", Err, Erl
End Sub

Private Sub lvwPunch_ItemClick(ByVal Item As MSComctlLib.ListItem)

5       On Error GoTo AfficherErreur

10      Dim rstPunch  As ADODB.Recordset
15      Dim iCompteur As Integer
        Dim G As Integer
        
  
        'Si le ListView n'est pas vide
20      If lvwPunch.ListItems.count > 0 Then
          'Pour aller chercher l'index du "punch" cliqué
25        m_lIDPunch = Item.Tag
    
30        Set rstPunch = New ADODB.Recordset
    
35        Call rstPunch.Open("SELECT * FROM GRB_Punch WHERE IDPunch = " & m_lIDPunch, g_connData, adOpenDynamic, adLockOptimistic)
      
40        If Left(rstPunch.Fields("NoProjet"), 1) = "E" Then
45          optTypePunch(I_OPT_ELECTRIQUE).Value = True
50        Else
55          optTypePunch(I_OPT_MECANIQUE).Value = True
60        End If
      
65        m_bClick = True

70        mskNoProjet.Text = Right(rstPunch.Fields("NoProjet"), 8)

75        m_bClick = False
      
80        If mskDate.mask = "##-##-##" Then
85          mskDate.mask = vbNullString
90        End If
      
95        mskDate.Text = rstPunch.Fields("Date")
      
100       mskHeureDebut.Text = rstPunch.Fields("HeureDébut")
      
105       If Not IsNull(rstPunch.Fields("HeureFin")) Then
110         If rstPunch.Fields("HeureFin") <> "" Then
115           mskHeureFin.Text = rstPunch.Fields("HeureFin")
120         Else
125           mskHeureFin.Text = "__:__"
130         End If
135       Else
140         mskHeureFin.Text = "__:__"
145       End If

150       txtClient.Text = vbNullString

155       Call AfficherClient

160       If Not IsNull(rstPunch.Fields("Type")) Then
                cmbType.ListIndex = -1
                
                If rstPunch.Fields("Type") = "Installation" Then
                    cmbType.ListIndex = 0
                    GoTo Fin_De_Type
                End If
                If rstPunch.Fields("Type") = "MiseService" Then
                    cmbType.ListIndex = 1
                    GoTo Fin_De_Type
                End If
                
                
165             If Left(rstPunch.Fields("NoProjet"), 1) = "E" Then
                    For G = 0 To cmbType.ListCount
                        If cmbType.LIST(G) = rstPunch.Fields("Type") Then
                            cmbType.ListIndex = G
                            Exit For
                        End If
                    Next
245             Else
                    For G = 0 To cmbType.ListCount
                        If cmbType.LIST(G) = rstPunch.Fields("Type") Then
                            cmbType.ListIndex = G
                            Exit For
                        End If
                    Next
                End If
315 Fin_De_Type:
320       End If

325       If Not IsNull(rstPunch.Fields("Commentaire")) Then
330         txtCommentaires.Text = rstPunch.Fields("Commentaire")
335       Else
340         txtCommentaires.Text = vbNullString
345       End If

350       If rstPunch.Fields("KM") = True Then
355         chkKM.Value = vbChecked

360         txtKM.Text = rstPunch.Fields("NbreKM")
365       Else
370         chkKM.Value = vbUnchecked
375         txtKM.Text = vbNullString
380       End If

385       Call rstPunch.Close
390       Set rstPunch = Nothing
      
395       Call ActiverControles(MODE_MODIF)
400     End If

405     Exit Sub

AfficherErreur:

410     woups "frmFeuilleTemps", "lvwPunch_ItemClick", Err, Erl
End Sub

Private Sub lvwPunch_KeyDown(KeyCode As Integer, Shift As Integer)

5       On Error GoTo AfficherErreur
        'Pour permettre d'effacer un enregistrement en appuyant sur "Delete" ou "Suppr"
  
        'Si il y a un enregistrement dans le listview
10      If lvwPunch.ListItems.count > 0 Then
          'Si la touche appuyée est "Delete"
15        If KeyCode = vbKeyDelete Then
            'Si l'utilisateur répond oui à "Etes vous sur?"
20          If MsgBox("Voulez-vous vraiment effacer le punch?", vbYesNo) = vbYes Then
              'Efface
25            Call g_connData.Execute("DELETE * FROM GRB_Punch WHERE IDPunch = " & lvwPunch.SelectedItem.Tag)
          
30            Call RemplirListView
        
35            Call ViderChamps
          
40            Call ActiverControles(MODE_INACTIF)
          
45            If lvwPunch.ListItems.count > 0 Then
                'Sélection du premier ListItem
50              lvwPunch.ListItems(1).Selected = True
55              Call lvwPunch_ItemClick(lvwPunch.SelectedItem)
60            End If
65          End If
70        End If
75      End If

80      Exit Sub

AfficherErreur:

85      woups "frmFeuilleTemps", "lvwPunch_KeyDown", Err, Erl
End Sub

Private Sub mskNoProjet_Change()

5       On Error GoTo AfficherErreur

10      If InStr(1, mskNoProjet.Text, "_") = 0 Then
15        Call AfficherClient
20      Else
25        txtClient.Text = vbNullString
30      End If

35      Call RemplirComboType

40      Exit Sub

AfficherErreur:

45      woups "frmFeuilleTemps", "mskNoProjet_Change", Err, Erl
End Sub

Private Sub AfficherClient()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset
15      Dim rstClient   As ADODB.Recordset
20      Dim iCompteur   As Integer
        
25      If m_bClick = False Then
30        Set rstProjSoum = New ADODB.Recordset
35        Set rstClient = New ADODB.Recordset
    
40        Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & lblPrefixe.Caption & mskNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
  
45        If Not rstProjSoum.EOF Then
50          Call rstClient.Open("SELECT NomClient FROM GRB_Client WHERE IDClient = " & rstProjSoum.Fields("NoClient"), g_connData, adOpenDynamic, adLockOptimistic)
    
55          txtClient.Text = rstClient.Fields("NomClient")
60          txtClient.Tag = rstProjSoum.Fields("NoClient")
    
65          Call rstClient.Close
70          Set rstClient = Nothing
    
75          If rstProjSoum.Fields("Ouvert") = False Then
80            Call MsgBox("Ce projet n'est pas ouvert!", vbOKOnly, "Erreur")
85          End If
90        Else
95          Call MsgBox("Projet inexistant!", vbOKOnly, "Erreur")
100       End If
  
105       Call rstProjSoum.Close
110       Set rstProjSoum = Nothing
115     End If

120     Exit Sub

AfficherErreur:

125     woups "frmFeuilleTemps", "AfficherClient", Err, Erl
End Sub

Private Sub mvwDate_DateClick(ByVal DateClicked As Date)

5       On Error GoTo AfficherErreur

10      mskDate.Text = ConvertDate(DateClicked)
  
        'Enlever le calendrier
15      mvwDate.Visible = False

20      Exit Sub

AfficherErreur:

25      woups "frmFeuilleTemps", "mvwDate_DateClick", Err, Erl
End Sub

Private Sub mvwDate_LostFocus()

5       On Error GoTo AfficherErreur
        'Quand le calendrier perd le focus, il faut l'enlever

10      mvwDate.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmFeuilleTemps", "mvwDate_LostFocus", Err, Erl
End Sub

Private Sub cmdImprimer_Click()
        
5       On Error GoTo AfficherErreur

10      Call frmChoixImpressionFT.Afficher(m_datSemaine, GetLastDay(m_datSemaine))

15      Exit Sub

AfficherErreur:

20      woups "frmFeuilleTemps", "cmdImprimer_Click", Err, Erl
End Sub

Private Sub mskDate_GotFocus()

5       On Error GoTo AfficherErreur
        
        'Met l'année sur 2 chiffres
10      If Len(mskDate.Text) = 10 Then
15        mskDate.Text = Right$(mskDate.Text, 8)
20      End If
  
25      mskDate.mask = "##-##-##"

30      Exit Sub

AfficherErreur:

35      woups "frmFeuilleTemps", "mskDate_GotFocus", Err, Erl
End Sub

Private Sub mskHeureDebut_GotFocus()

5       On Error GoTo AfficherErreur
        
        'Format d'heure
10      mskHeureDebut.mask = "##:##"

15      Exit Sub

AfficherErreur:

20      woups "frmFeuilleTemps", "mskHeureDebut_GotFocus", Err, Erl
End Sub

Private Sub mskHeureFin_GotFocus()

5       On Error GoTo AfficherErreur
        
        'Format d'heure
10      mskHeureFin.mask = "##:##"

15      Exit Sub

AfficherErreur:

20      woups "frmFeuilleTemps", "mskHeureFin_GotFocus", Err, Erl
End Sub

Private Sub mskDate_LostFocus()

5       On Error GoTo AfficherErreur
        
        'Enlève le mask
10      mskDate.mask = vbNullString
  
        'Vide le champs si l'utilisateur n'a rien écrit
15      If mskDate.Text = "__-__-__" Then
20        mskDate.Text = vbNullString
25      Else
          'Remet l'année sur 8 chiffres
30        If Len(mskDate.Text) = 8 Then
35          If IsDate(mskDate.Text) Then
40            mskDate.Text = Year(DateSerial(Left$(mskDate, 2), Mid$(mskDate, 4, 2), Right$(mskDate.Text, 2))) & Mid$(mskDate.Text, 3, 8)
45          End If
50        End If
55      End If

60      Exit Sub

AfficherErreur:

65      woups "frmFeuilleTemps", "mskDate_LostFocus", Err, Erl
End Sub

Private Sub mskHeureDebut_LostFocus()
    Dim heure As Date
5       On Error GoTo AfficherErreur
        
        'Enlève le mask
10      mskHeureDebut.mask = vbNullString
  
        'Vide le champs si l'utilisateur n'a rien écrit
15      If mskHeureDebut.Text = "__:__" Then
20        mskHeureDebut.Text = vbNullString
25
        Else

30      heure = CDate(mskHeureDebut.Text)
        
        If Minute(heure) <= 5 Then
35          heure = TimeSerial(Hour(heure), 0, 0)
        Else
40          If Minute(heure) <= 24 Then
45              heure = TimeSerial(Hour(heure), 15, 0)
            Else
                If Minute(heure) <= 35 Then
50                heure = TimeSerial(Hour(heure), 30, 0)
                Else
                    If Minute(heure) <= 54 Then
55                      heure = TimeSerial(Hour(heure), 45, 0)
                    Else
60                      heure = TimeSerial(Hour(heure) + 1, 0, 0)
                    End If
                End If
            End If
        End If
        
65      mskHeureDebut.Text = Right$("0" & Hour(heure), 2) + ":" + Right$("0" & Minute(heure), 2)
    
    
    
    End If
        
        
        
        
        
70      Exit Sub

AfficherErreur:

75      woups "frmFeuilleTemps", "mskHeureDebut_LostFocus", Err, Erl
End Sub

Private Sub mskHeureFin_LostFocus()
    Dim heure As Date
5       On Error GoTo AfficherErreur
        
        'Enlève le mask
10      mskHeureFin.mask = vbNullString
  
        'Vide le champs si l'utilisateur n'a rien écrit
15      If mskHeureFin.Text = "__:__" Then
20        mskHeureFin.Text = vbNullString
        Else
25          heure = CDate(mskHeureFin.Text)
        
            If Minute(heure) <= 5 Then
30              heure = TimeSerial(Hour(heure), 0, 0)
            Else
35              If Minute(heure) <= 24 Then
40              heure = TimeSerial(Hour(heure), 15, 0)
            Else
45              If Minute(heure) <= 35 Then
50                  heure = TimeSerial(Hour(heure), 30, 0)
                Else
                    If Minute(heure) <= 54 Then
55                      heure = TimeSerial(Hour(heure), 45, 0)
                    Else
60                      heure = TimeSerial(Hour(heure) + 1, 0, 0)
                    End If
                End If
            End If
        End If
        
65      mskHeureFin.Text = Right$("0" & Hour(heure), 2) + ":" + Right$("0" & Minute(heure), 2)



      End If

70      Exit Sub

AfficherErreur:

75      woups "frmFeuilleTemps", "mskHeureFin_LostFocus", Err, Erl
End Sub

Private Sub txtKM_LostFocus()

5       On Error GoTo AfficherErreur

10      txtKM.Text = Replace(txtKM.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmFeuilleTemps", "txtKM_LostFocus", Err, Erl
End Sub

Public Sub RemplirComboType()
  
5       On Error GoTo AfficherErreur

10      Dim bInstallation As Boolean
15      Dim bTypeInutile  As Boolean
        Dim tbltype As ADODB.Recordset
        Set tbltype = New ADODB.Recordset
  
20      Call cmbType.Clear

25      If Mid$(mskNoProjet.Text, 2, 1) = "1" Then
30        bTypeInutile = True
35      End If

40      If IsNumeric(Right$(mskNoProjet.Text, 2)) Then
45        If Mid$(mskNoProjet.Text, 2, 4) <> "3000" Then
50          If CInt(Right$(mskNoProjet.Text, 2)) >= 51 And CInt(Right$(mskNoProjet.Text, 2)) <= 59 Then
55            bInstallation = True
60          End If
65        Else
70          bTypeInutile = True
75        End If
80      End If
  
85      If bTypeInutile = False Then
90        lblType.Visible = True
95        cmbType.Visible = True

100       If bInstallation = True Then
105         If lblPrefixe.Caption = "E" Then
110           Call cmbType.AddItem("Installation")
115           Call cmbType.AddItem("Mise en service")
120         Else
125           Call cmbType.AddItem("Installation")
130         End If
135       Else
140         If lblPrefixe.Caption = "E" Then
145             Call tbltype.Open("select * from TBL_Punch_Type Where Mode = 'E' Order by name", g_connData, adOpenDynamic, adLockOptimistic)
150             Do While Not tbltype.EOF
155                 cmbType.AddItem (tbltype.Fields("Name"))
160                 Call tbltype.MoveNext
165             Loop
170             Call tbltype.Close
175             Set tbltype = Nothing
                
                
                
                
200         Else
205
210             Call tbltype.Open("select * from TBL_Punch_Type Where Mode = 'M' Order by name ", g_connData, adOpenDynamic, adLockOptimistic)
215             Do While Not tbltype.EOF
220                 cmbType.AddItem (tbltype.Fields("Name"))
225                 Call tbltype.MoveNext
230             Loop
235             Call tbltype.Close
240             Set tbltype = Nothing
255         End If
260       End If
265     Else
270       lblType.Visible = False
275       cmbType.Visible = False
280     End If

285     Exit Sub

AfficherErreur:

290     woups "frmFeuilleTemps", "RemplirComboType", Err, Erl
End Sub

