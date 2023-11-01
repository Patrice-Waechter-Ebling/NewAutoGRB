VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoutils 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Machinerie & Outillage"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   7890
   Begin VB.CommandButton cmdRechercher 
      Caption         =   "Rechercher"
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtRecherche 
      Height          =   285
      Left            =   4560
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
   End
   Begin VB.ComboBox cmbdepartement 
      Height          =   315
      ItemData        =   "frmoutils.frx":0000
      Left            =   120
      List            =   "frmoutils.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton CmdAdd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Ajouter"
      Height          =   495
      Left            =   1680
      TabIndex        =   28
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton CmdSupp 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Supprimer"
      Height          =   495
      Left            =   3240
      TabIndex        =   31
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton CmdFerme 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Fermer"
      Height          =   495
      Left            =   6360
      TabIndex        =   33
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton CmdModif 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Modifier"
      Height          =   495
      Left            =   4800
      TabIndex        =   32
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdreport 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Impression"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   27
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton CmdEnr 
      Caption         =   "&Enregistrer"
      Height          =   495
      Left            =   1680
      TabIndex        =   29
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton CmdAnul 
      Caption         =   "A&nnuler"
      Height          =   495
      Left            =   3240
      TabIndex        =   30
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Frame fraModif 
      BackColor       =   &H00404040&
      Height          =   3615
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   7695
      Begin MSComCtl2.MonthView mvwDate 
         Height          =   2370
         Left            =   1800
         TabIndex        =   11
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
         StartOfWeek     =   152633345
         CurrentDate     =   37726
      End
      Begin VB.CommandButton cmdDateAchat 
         Caption         =   "..."
         Height          =   255
         Left            =   2640
         TabIndex        =   19
         Top             =   2040
         Width           =   375
      End
      Begin VB.CommandButton cmdDateHorsfonction 
         Caption         =   "..."
         Height          =   255
         Left            =   2640
         TabIndex        =   23
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox txtno 
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox cmbetiquette 
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   600
         Width           =   2535
      End
      Begin VB.ComboBox cmbdepartement_modif 
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtcommentaire 
         Height          =   645
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   2760
         Width           =   5895
      End
      Begin VB.TextBox txtoutils 
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Top             =   1320
         Width           =   4815
      End
      Begin MSMask.MaskEdBox txtcout 
         Height          =   300
         Left            =   1440
         TabIndex        =   16
         Top             =   1680
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txthorsfonction 
         Height          =   300
         Left            =   1440
         TabIndex        =   22
         Top             =   2400
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtachat 
         Height          =   300
         Left            =   1440
         TabIndex        =   18
         Top             =   2040
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "AA-MM-JJ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   2220
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Outil"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Commentaire"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Disposition"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Date achat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Étiquette"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Coût"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Département"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Outil"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView lstoutils 
      Height          =   3375
      Left            =   120
      TabIndex        =   26
      Top             =   1680
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5953
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
         Text            =   "No"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nom"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Achat"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Disposition"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Coût"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Étiquette"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Commentaire"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label lblRecherche 
      BackStyle       =   0  'Transparent
      Caption         =   "Recherche :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lbldepartement 
      BackStyle       =   0  'Transparent
      Caption         =   "Département"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
End
Attribute VB_Name = "frmoutils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bDateAchat As Boolean 'date d'achat=true ou horsfonction = false
Private m_bModeAjouter As Boolean

Private Sub cmbetiquette_Click()

 On Error GoTo Oups

 ''''''''''''''''''''''''''''''''''''''''''
 'trouve le prochain no_outils automatique
 ''''''''''''''''''''''''''''''''''''''''''
 Dim rstOutils As ADODB.Recordset
 
 Set rstOutils = New ADODB.Recordset
 
 Call rstOutils.Open("SELECT * FROM GrbOutils WHERE type_étiquette = '" & cmbetiquette.Text & "' ORDER BY no_outils DESC", g_connData, adOpenDynamic, adLockOptimistic)
 
 'incremente de un le dernier no outils
 If Not rstOutils.EOF Then
 If IsNumeric(rstOutils.Fields("no_outils")) Then
 txtNo.Text = rstOutils.Fields("no_outils") + 1
 Else
 txtNo.Text = ""
 End If
 End If
 
 'ferme la table
  Call rstOutils.Close
  Set rstOutils = Nothing

  Exit Sub

Oups:

  wOups "frmoutils", "cmbetiquette_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdAnul_Click()

 On Error GoTo Oups

 ''''''''''''''''''''''''''''''
 'affiche en mode visualisation
 '''''''''''''''''''''''''''''''
 Call AfficherListe

 Exit Sub

Oups:

 wOups "frmoutils", "CmdAnul_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdDateAchat_Click()

 On Error GoTo Oups

 '''''''''''''''''''''''''''''''
 'affiche pour choisir une date
 ''''''''''''''''''''''''''''''''
 m_bDateAchat = True
 mvwDate.Visible = True

 'si pas de date met la date du jour
 If IsDate(txtachat.Text) = True Then
 mvwDate.Year = Left$(txtachat.Text, 4)
 mvwDate.Month = Mid$(txtachat.Text, 6, 2)
 mvwDate.Day = Right$(txtachat.Text, 2)
 Else
 mvwDate.Value = Date
 End If

 Call mvwDate.SetFocus

  Exit Sub

Oups:

  wOups "frmoutils", "cmdDateAchat_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdDateHorsfonction_Click()

 On Error GoTo Oups

 '''''''''''''''''''''''''''''''
 'affiche pour choisir une date
 ''''''''''''''''''''''''''''''''
 m_bDateAchat = False
 
 mvwDate.Visible = True

 'si pas de date met la date du jour
 If IsDate(txthorsfonction.Text) = True Then
 mvwDate.Year = Left$(txthorsfonction.Text, 4)
 mvwDate.Month = Mid$(txthorsfonction.Text, 6, 2)
 mvwDate.Day = Right$(txthorsfonction.Text, 2)
 Else
 mvwDate.Value = Date
 End If
 
 Call mvwDate.SetFocus

  Exit Sub

Oups:

  wOups "frmoutils", "cmdDateHorsfonction_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdFerme_Click()

 On Error GoTo Oups

 'quitte fenetre
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmoutils", "CmdFerme_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdRechercher_Click()

 On Error GoTo Oups

 Dim rstRecherche As ADODB.Recordset
 Dim itmOutils As ListItem

 Screen.MousePointer = vbHourglass

 Call lstoutils.ListItems.Clear

 Set rstRecherche = New ADODB.Recordset

 Call rstRecherche.Open("SELECT * FROM GrbOutils WHERE (Instr(1,CStr(no_outils),'" & txtRecherche.Text & "') > 0 OR Instr(1,nom_outils,'" & txtRecherche.Text & "') > 0) AND Departement = '" & Replace(cmbdepartement.Text, "'", "''") & "' ORDER BY no_outils", g_connData, adOpenDynamic, adLockOptimistic)

 Do While Not rstRecherche.EOF
 Set itmOutils = lstoutils.ListItems.Add
 
 itmOutils.Text = rstRecherche.Fields("no_outils")
 
 If Not IsNull(rstRecherche.Fields("nom_outils")) Then
  itmOutils.SubItems(1) = rstRecherche.Fields("nom_outils")
  End If
 
  If Not IsNull(rstRecherche.Fields("date_achat")) Then
  itmOutils.SubItems(2) = rstRecherche.Fields("date_achat")
  End If
 
  If Not IsNull(rstRecherche.Fields("date_hors_fonction")) Then
  itmOutils.SubItems(3) = rstRecherche.Fields("date_hors_fonction")
  End If
 
If Not IsNull(rstRecherche.Fields("cout")) Then
itmOutils.SubItems(4) = Conversion(rstRecherche.Fields("cout"), MODE_ARGENT)
 End If
 
 If Not IsNull(rstRecherche.Fields("type_étiquette")) Then
 itmOutils.SubItems(5) = rstRecherche.Fields("type_étiquette")
 End If
 
 If Not IsNull(rstRecherche.Fields("commentaire")) Then
 itmOutils.SubItems(6) = rstRecherche.Fields("commentaire")
 End If
 
 Call rstRecherche.MoveNext
Loop
 
Call rstRecherche.Close
1  Set rstRecherche = Nothing
 
Screen.MousePointer = vbDefault

 Exit Sub

Oups:

wOups "frmoutils", "cmdRechercher_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdreport_Click()

 On Error GoTo Oups

 Dim rstOutils As ADODB.Recordset

 Screen.MousePointer = vbHourglass

 Set rstOutils = New ADODB.Recordset

 Call rstOutils.Open("SELECT * FROM Grboutils WHERE departement = '" & Replace(cmbdepartement.Text, "'", "''") & "' ORDER BY no_outils", g_connData, adOpenDynamic, adLockOptimistic)
 
 ''''''''''''''''''''''''''''''''''''''''''
 'rapport liste d'outil pour un departement
 ''''''''''''''''''''''''''''''''''''''''''
 
 'set le rapport
 Set DR_Outils_machinerie.DataSource = rstOutils
 
 'contenu label
 DR_Outils_machinerie.Sections("section2").Controls("lbldepartement").Caption = "Outils & machinerie " + LCase(rstOutils!departement)
 DR_Outils_machinerie.Sections("section3").Controls("lbldate").Caption = ConvertDate(Date)
 DR_Outils_machinerie.Orientation = rptOrientLandscape
 
 'affiche rapport
 Call DR_Outils_machinerie.Show(vbModal)
 
 Screen.MousePointer = vbDefault

  Exit Sub

Oups:

  wOups "frmoutils", "cmdreport_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 Call RemplirDepartement

 Exit Sub

Oups:

 wOups "frmoutils", "Form_Load", Err, Err.number, Err.Description
End Sub

Public Sub RemplirDepartement()

 On Error GoTo Oups

 '''''''''''''''''''''''''''''''''''''''''''''''''
 'remplis combo departement en mode visualisation
 '''''''''''''''''''''''''''''''''''''''''''''''''
 Dim rstDepartement As ADODB.Recordset
 
 Set rstDepartement = New ADODB.Recordset
 
 Call rstDepartement.Open("SELECT DISTINCT departement FROM Grboutils ORDER BY departement", g_connData, adOpenDynamic, adLockOptimistic)
 
 Call cmbdepartement.Clear
 
 'rempli tant il y a des departement
 Do While Not rstDepartement.EOF
 Call cmbdepartement.AddItem(rstDepartement.Fields("departement"))

 Call rstDepartement.MoveNext
 Loop
 
 'ferme la table
 Call rstDepartement.Close
 Set rstDepartement = Nothing
 
 'si il y a des departement ,selectionne par defaut
  If cmbdepartement.ListCount > 0 Then
  cmbdepartement.ListIndex = 0
  End If

  Exit Sub

Oups:

  wOups "frmoutils", "RemplirDepartement", Err, Err.number, Err.Description
End Sub

Public Sub RemplirEtiquette()

 On Error GoTo Oups

 ''''''''''''''''''''''''''''''''''''''''''''',
 'remplis combo etiquette en mode modification
 ''''''''''''''''''''''''''''''''''''''''''''''
 Dim rstEtiquette As ADODB.Recordset
 
 Set rstEtiquette = New ADODB.Recordset
 
 Call rstEtiquette.Open("SELECT DISTINCT type_étiquette FROM Grboutils ORDER BY type_étiquette", g_connData, adOpenDynamic, adLockOptimistic)
 
 Call cmbetiquette.Clear
 
 'rempli tant il y a des type_étiquette
 Do While Not rstEtiquette.EOF
 Call cmbetiquette.AddItem(rstEtiquette.Fields("type_étiquette"))

 Call rstEtiquette.MoveNext
 Loop
 
 'ferme la table
 Call rstEtiquette.Close
 Set rstEtiquette = Nothing

  Exit Sub

Oups:

  wOups "frmoutils", "RemplirEtiquette", Err, Err.number, Err.Description
End Sub

Public Sub RemplirDepartementModif()

 On Error GoTo Oups

 '''''''''''''''''''''''''''''''''''''''''''''''
 'remplis combo departement en mode modification
 '''''''''''''''''''''''''''''''''''''''''''''''
 Dim rstDepartement As ADODB.Recordset
 
 Set rstDepartement = New ADODB.Recordset
 
 Call rstDepartement.Open("SELECT DISTINCT departement FROM Grboutils ORDER BY departement", g_connData, adOpenDynamic, adLockOptimistic)
 
 Call cmbdepartement_modif.Clear
 
 'rempli tant il y a des employé
 Do While Not rstDepartement.EOF
 Call cmbdepartement_modif.AddItem(rstDepartement.Fields("departement"))

 Call rstDepartement.MoveNext
 Loop
 
 'ferme la table
 Call rstDepartement.Close
 Set rstDepartement = Nothing

  Exit Sub

Oups:

  wOups "frmoutils", "RemplirDepartementModif", Err, Err.number, Err.Description
End Sub

Public Sub RemplirListViewOutils()

 On Error GoTo Oups

 'remplis lister une journée
 Dim rstOutils As ADODB.Recordset
 Dim itmOutils As ListItem
 
 'vide le lister
 Call lstoutils.ListItems.Clear

 lstoutils.Sorted = False

 Set rstOutils = New ADODB.Recordset

 Call rstOutils.Open("SELECT * FROM Grboutils WHERE departement = '" & Replace(cmbdepartement.Text, "'", "''") & "' ORDER BY no_outils", g_connData, adOpenDynamic, adLockOptimistic)
 
 'tant il y a de employé cedulé , ajoute dans lister
 Do While Not rstOutils.EOF
 Set itmOutils = lstoutils.ListItems.Add
 
 itmOutils.Text = rstOutils.Fields("no_outils")
 
 If IsNull(rstOutils.Fields("nom_outils")) Then
  Call itmOutils.ListSubItems.Add(, , vbNullString)
  Else
  Call itmOutils.ListSubItems.Add(, , rstOutils.Fields("nom_outils"))
  End If
 
  If IsNull(rstOutils.Fields("date_achat")) Then
  Call itmOutils.ListSubItems.Add(, , vbNullString)
  Else
  Call itmOutils.ListSubItems.Add(, , rstOutils.Fields("date_achat"))
End If
 
1 If IsNull(rstOutils.Fields("date_hors_fonction")) Then
 Call itmOutils.ListSubItems.Add(, , vbNullString)
 Else
 Call itmOutils.ListSubItems.Add(, , rstOutils.Fields("date_hors_fonction"))
 End If
 
 If IsNull(rstOutils.Fields("cout")) Or rstOutils.Fields("Cout") = vbNullString Then
 Call itmOutils.ListSubItems.Add(, , vbNullString)
 Else
 Call itmOutils.ListSubItems.Add(, , Conversion(rstOutils.Fields("cout"), MODE_ARGENT))
 End If
 
 If IsNull(rstOutils.Fields("type_étiquette")) Then
 Call itmOutils.ListSubItems.Add(, , vbNullString)
 Else
 Call itmOutils.ListSubItems.Add(, , rstOutils.Fields("type_étiquette"))
 End If
 
 If IsNull(rstOutils.Fields("commentaire")) Then
 Call itmOutils.ListSubItems.Add(, , vbNullString)
 Else
1  Call itmOutils.ListSubItems.Add(, , rstOutils.Fields("commentaire"))
 End If
 
 Call rstOutils.MoveNext
Loop
 
 'ferme la table
Call rstOutils.Close
Set rstOutils = Nothing

Exit Sub

Oups:

wOups "frmoutils", "RemplirListViewOutils", Err, Err.number, Err.Description
End Sub

Private Sub cmbdepartement_Click()

 On Error GoTo Oups

 Screen.MousePointer = vbHourglass
 
 'remplis le lister outil dependant le departement selectionné
 Call RemplirListViewOutils
 
 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmoutils", "cmbdepartement_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdAdd_Click()

 On Error GoTo Oups

 Screen.MousePointer = vbHourglass

 'proc qui permet d'ajouter un outils
 m_bModeAjouter = True
 
 'affiche en mode modification
 Call AfficherModif
 
 'remplis les combo en mode modification
 Call RemplirEtiquette
 Call RemplirDepartementModif
 
 'vide les champs
 txtachat.Text = vbNullString
 txtcommentaire.Text = vbNullString
 txtcout.Text = vbNullString
 txthorsfonction.Text = vbNullString
 txtNo.Text = vbNullString
  txtoutils.Text = vbNullString
 
  Screen.MousePointer = vbDefault

  Exit Sub

Oups:

  wOups "frmoutils", "CmdAdd_Click", Err, Err.number, Err.Description
End Sub

Private Sub AfficherModif()

 On Error GoTo Oups
 
 'met visible les champ pour modifié ou ajouté
 fraModif.Visible = True
 lstoutils.Visible = False
 cmbdepartement.Visible = False
 lbldepartement.Visible = False
 CmdAdd.Visible = False
 CmdSupp.Visible = False
 CmdModif.Visible = False
 CmdEnr.Visible = True
 CmdAnul.Visible = True
 lblRecherche.Visible = False
  txtRecherche.Visible = False
  cmdRechercher.Visible = False
  cmdReport.Visible = False
  CmdFerme.Visible = False

  Exit Sub

Oups:

  wOups "frmoutils", "AfficherModif", Err, Err.number, Err.Description
End Sub

Private Sub AfficherListe()

 On Error GoTo Oups

 'met visible les champ pour modifié ou ajouté
 fraModif.Visible = False
 lstoutils.Visible = True
 cmbdepartement.Visible = True
 lbldepartement.Visible = True
 CmdAdd.Visible = True
 CmdSupp.Visible = True
 CmdModif.Visible = True
 CmdEnr.Visible = False
 CmdAnul.Visible = False
 lblRecherche.Visible = True
  txtRecherche.Visible = True
  cmdRechercher.Visible = True
  cmdReport.Visible = True
  CmdFerme.Visible = True

  Exit Sub

Oups:

  wOups "frmoutils", "AfficherListe", Err, Err.number, Err.Description
End Sub

Private Sub CmdModif_Click()

 On Error GoTo Oups

 Dim rstOutils As ADODB.Recordset
 
 Screen.MousePointer = vbHourglass
 
 'proc qui permet d'ajouter un outils
 m_bModeAjouter = False
 
 'affiche en mode modification
 Call AfficherModif
 
 'remplis les combo en mode modification
 Call RemplirEtiquette
 Call RemplirDepartementModif

 If lstoutils.ListItems.count > 0 Then
 'ouvre la table
 Set rstOutils = New ADODB.Recordset
 
 Call rstOutils.Open("SELECT * FROM Grboutils WHERE no_outils = '" & lstoutils.SelectedItem.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 frmoutils.cmbdepartement_modif.Text = rstOutils.Fields("departement")
  frmoutils.cmbetiquette.Text = rstOutils.Fields("type_étiquette")
  frmoutils.txtNo.Text = rstOutils.Fields("no_outils")
 
  If IsNull(rstOutils.Fields("nom_outils")) Then
  txtoutils.Text = vbNullString
  Else
  txtoutils.Text = rstOutils.Fields("nom_outils")
  End If
 
  If IsNull(rstOutils.Fields("cout")) Then
 txtcout.Text = vbNullString
1 Else
 txtcout.Text = rstOutils.Fields("cout")
 End If
 
 If Not IsNull(rstOutils.Fields("date_achat")) Then
 txtachat.Text = rstOutils.Fields("date_achat")
 Else
 txtachat.Text = vbNullString
 End If
 
 If Not IsNull(rstOutils.Fields("date_hors_fonction")) Then
 txthorsfonction.Text = rstOutils.Fields("date_hors_fonction")
 Else
 txthorsfonction.Text = vbNullString
 End If

 If IsNull(rstOutils.Fields("commentaire")) Then
 txtcommentaire.Text = vbNullString
 Else
 txtcommentaire.Text = rstOutils.Fields("commentaire")
 End If
 
 'ferme la table
1  Call rstOutils.Close
 Set rstOutils = Nothing
 End If

Screen.MousePointer = vbDefault

Exit Sub

Oups:

wOups "frmoutils", "CmdModif_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdSupp_Click()

 On Error GoTo Oups

 ''''''''''''''''''''''''''''''''
 'Supprime l'outils selectionné
 ''''''''''''''''''''''''''''''''
 If lstoutils.ListItems.count > 0 Then
 If MsgBox("Voulez-vous supprimer cette enregistrement?", vbYesNo, "Supprimer") = vbYes Then
 Screen.MousePointer = vbHourglass
 
 Call g_connData.Execute("DELETE * FROM Grboutils WHERE no_outils = '" & lstoutils.SelectedItem.Text & "'")
 
 'mise a jour des lister
 Call RemplirListViewOutils
 
 Call RemplirDepartement
 
 Screen.MousePointer = vbDefault
 End If
 End If

 Exit Sub

Oups:

  wOups "frmoutils", "CmdSupp_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdEnr_Click()

 On Error GoTo Oups

 'enregistre
 Dim rstOutils As ADODB.Recordset
 Dim rstVerifModif As ADODB.Recordset

 Screen.MousePointer = vbHourglass
 
 If cmbdepartement.Text = vbNullString Or cmbetiquette.Text = vbNullString Or txtNo.Text = vbNullString Then
 Call MsgBox("Champs vide!", vbOKOnly, "Erreur")
 
 Screen.MousePointer = vbDefault

 Exit Sub
 End If
 
 If (Len(txtachat.Text) = 0 Or (Len(txtachat.Text) > 1 And IsDate(txtachat.Text))) And (Len(txthorsfonction.Text) = 0 Or (Len(txthorsfonction.Text) > 1 And IsDate(txthorsfonction.Text))) Then
 'ouvre la table
 Set rstOutils = New ADODB.Recordset
 
  If m_bModeAjouter = True Then
  Call rstOutils.Open("SELECT * FROM Grboutils WHERE no_outils = '" & txtNo.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
  If Not rstOutils.EOF Then
  Call MsgBox("Le numéro d'outils existe déjà!", vbOKOnly, "Erreur")
 
  Call rstOutils.Close
  Set rstOutils = Nothing
 
  Screen.MousePointer = vbDefault
 
  Exit Sub
 End If
 
Call rstOutils.AddNew
 Else
 Call rstOutils.Open("SELECT * FROM Grboutils WHERE no_outils = '" & txtNo.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 End If
 
 ''''''''''''''''''''''''''
 'ajoute l'enregistrement
 ''''''''''''''''''''''''''
 rstOutils.Fields("departement") = cmbdepartement_modif.Text
 rstOutils.Fields("type_étiquette") = cmbetiquette.Text
 rstOutils.Fields("no_outils") = txtNo.Text
 rstOutils.Fields("nom_outils") = txtoutils.Text
 rstOutils.Fields("cout") = txtcout.Text
 rstOutils.Fields("date_achat") = txtachat.Text
 rstOutils.Fields("date_hors_fonction") = txthorsfonction.Text
rstOutils.Fields("commentaire") = txtcommentaire.Text
 
 Call rstOutils.Update
 
 'quitte ecran pour ajouté ou modifié
 Call AfficherListe
 
 Call rstOutils.Close
 Set rstOutils = Nothing
  
 'met a jour l'écran
 Call RemplirListViewOutils
 
 Call RemplirDepartement
1  Else
 Call MsgBox("La date est invalide! (aaaa-mm-jj)", , "Erreur")
 End If
 
Screen.MousePointer = vbDefault

Exit Sub

Oups:

wOups "frmoutils", "CmdEnr_Click", Err, Err.number, Err.Description
End Sub

Private Sub lstoutils_DblClick()

 On Error GoTo Oups

 'affiche l'écran en mode modification
 Call CmdModif_Click

 Exit Sub

Oups:

 wOups "frmoutils", "lstoutils_DblClick", Err, Err.number, Err.Description
End Sub

Private Sub lstoutils_KeyDown(KeyCode As Integer, Shift As Integer)

 On Error GoTo Oups

 If lstoutils.ListItems.count > 0 Then
 If KeyCode = vbKeyDelete Then
 If MsgBox("Voulez-vous supprimer cette enregistrement?", vbYesNo, "Supprimer") = vbYes Then
 Call g_connData.Execute("DELETE * FROM Grboutils WHERE no_outils = '" & lstoutils.SelectedItem.Text & "'")
 
 'mise a jour des lister
 Call RemplirListViewOutils
 End If
 End If
 End If

 Exit Sub

Oups:

 wOups "frmoutils", "lstoutils_KeyDown", Err, Err.number, Err.Description
End Sub

Private Sub mvwDate_DateClick(ByVal DateClicked As Date)

 On Error GoTo Oups

 '''''''''''''''''''''''''''''''''''''''''
 'affiche dans l'écran la date sélectionné
 '''''''''''''''''''''''''''''''''''''''''
 Dim sDate As String
 
 'ajoute dans le champ text la date selectionné
 If m_bDateAchat = True Then
 txtachat.Text = ConvertDate(DateClicked)
 Else
 txthorsfonction.Text = ConvertDate(DateClicked)
 End If
 
 mvwDate.Visible = False

 Exit Sub

Oups:

 wOups "frmoutils", "mvwDate_DateClick", Err, Err.number, Err.Description
End Sub

Private Sub mvwDate_LostFocus()

 On Error GoTo Oups

 ''''''''''''''''''''''''''''''
 'lorsque clique ailleur cache
 '''''''''''''''''''''''''''''''
 mvwDate.Visible = False

 Exit Sub

Oups:

 wOups "frmoutils", "mvwDate_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtachat_GotFocus()

 On Error GoTo Oups

 'met l'année 2caratere
 If Len(txtachat.Text) = 10 Then
 txtachat.Text = Right$(txtachat.Text, 8)
 End If
 
 'met le mask
 txtachat.mask = "##-##-##"

 Exit Sub

Oups:

 wOups "frmoutils", "txtachat_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtachat_LostFocus()

 On Error GoTo Oups

 'enleve le mask
 txtachat.mask = vbNullString
 
 'losque est vide , enleve les caratere du masque
 If txtachat.Text = "__-__-__" Then
 txtachat.Text = vbNullString
 End If
 
 'met l'année 4 caractere
 If Len(txtachat.Text) =   Then
 If IsDate(txtachat.Text) Then
 txtachat.Text = Trim$(Year(DateSerial(Mid$(txtachat.Text, 1, 2), Mid$(txtachat.Text, 4, 2), Mid$(txtachat.Text, 7, 2)))) + Mid$(txtachat.Text, 3, 8)
 End If
 End If

 Exit Sub

Oups:

  wOups "frmoutils", "txtachat_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtcout_LostFocus()

 On Error GoTo Oups

 txtcout.Text = Replace(txtcout.Text, ".", ",")

 If Not IsNumeric(txtcout.Text) Then
 Call MsgBox("Valeur non numérique!", vbOKOnly, "Erreur")

 txtcout.Text = ""
 End If

 Exit Sub

Oups:

 wOups "frmoutils", "txtcout_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txthorsfonction_GotFocus()

 On Error GoTo Oups

 'met l'année 2caratere
 If Len(txthorsfonction.Text) = 10 Then
 txthorsfonction.Text = Right$(txthorsfonction.Text, 8)
 End If

 'met le mask
 txthorsfonction.mask = "##-##-##"

 Exit Sub

Oups:

 wOups "frmoutils", "txthorsfonction_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub txthorsfonction_LostFocus()

 On Error GoTo Oups

 'enleve le mask
 txthorsfonction.mask = vbNullString
 
 'losque est vide , enleve les caratere du masque
 If txthorsfonction.Text = "__-__-__" Then
 txthorsfonction.Text = vbNullString
 End If
 
 'met l'année 4 caractere
 If Len(txthorsfonction.Text) =   Then
 If IsDate(txthorsfonction.Text) Then
 txthorsfonction.Text = Trim$(Year(DateSerial(Mid$(txthorsfonction.Text, 1, 2), Mid$(txthorsfonction.Text, 4, 2), Mid$(txthorsfonction.Text, 7, 2)))) + Mid$(txthorsfonction.Text, 3, 8)
 End If
 End If

 Exit Sub

Oups:

  wOups "frmoutils", "txthorsfonction_LostFocus", Err, Err.number, Err.Description
End Sub
