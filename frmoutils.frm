VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoutils 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Machinerie & Outillage"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmoutils.frx":0000
   ScaleHeight     =   5865
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
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
      ItemData        =   "frmoutils.frx":2F0D
      Left            =   120
      List            =   "frmoutils.frx":2F0F
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
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   120
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
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   0
         Appearance      =   1
         ShowToday       =   0   'False
         StartOfWeek     =   90243073
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

Private m_bDateAchat   As Boolean     'date d'achat=true ou horsfonction = false
Private m_bModeAjouter As Boolean

Private Sub cmbetiquette_Click()

5       On Error GoTo AfficherErreur

        ''''''''''''''''''''''''''''''''''''''''''
        'trouve le prochain no_outils automatique
        ''''''''''''''''''''''''''''''''''''''''''
10      Dim rstOutils As ADODB.Recordset
  
15      Set rstOutils = New ADODB.Recordset
  
20      Call rstOutils.Open("SELECT * FROM GRB_Outils WHERE type_étiquette = '" & cmbetiquette.Text & "' ORDER BY no_outils DESC", g_connData, adOpenDynamic, adLockOptimistic)
    
        'incremente de un le dernier no outils
25      If Not rstOutils.EOF Then
30        If IsNumeric(rstOutils.Fields("no_outils")) Then
35          txtNo.Text = rstOutils.Fields("no_outils") + 1
40        Else
45          txtNo.Text = ""
50        End If
55      End If
    
        'ferme la table
60      Call rstOutils.Close
65      Set rstOutils = Nothing

70      Exit Sub

AfficherErreur:

75      woups "frmoutils", "cmbetiquette_Click", Err, Erl
End Sub

Private Sub CmdAnul_Click()

5       On Error GoTo AfficherErreur

        ''''''''''''''''''''''''''''''
        'affiche en mode visualisation
        '''''''''''''''''''''''''''''''
10      Call AfficherListe

15      Exit Sub

AfficherErreur:

20      woups "frmoutils", "CmdAnul_Click", Err, Erl
End Sub

Private Sub cmdDateAchat_Click()

5       On Error GoTo AfficherErreur

        '''''''''''''''''''''''''''''''
        'affiche pour choisir une date
        ''''''''''''''''''''''''''''''''
10      m_bDateAchat = True
15      mvwDate.Visible = True

        'si pas de date met la date du jour
20      If IsDate(txtachat.Text) = True Then
25        mvwDate.Year = Left$(txtachat.Text, 4)
30        mvwDate.Month = Mid$(txtachat.Text, 6, 2)
35        mvwDate.Day = Right$(txtachat.Text, 2)
40      Else
45        mvwDate.Value = Date
50      End If

55      Call mvwDate.SetFocus

60      Exit Sub

AfficherErreur:

65      woups "frmoutils", "cmdDateAchat_Click", Err, Erl
End Sub

Private Sub cmdDateHorsfonction_Click()

5       On Error GoTo AfficherErreur

        '''''''''''''''''''''''''''''''
        'affiche pour choisir une date
        ''''''''''''''''''''''''''''''''
10      m_bDateAchat = False
  
15      mvwDate.Visible = True

        'si pas de date met la date du jour
20      If IsDate(txthorsfonction.Text) = True Then
25        mvwDate.Year = Left$(txthorsfonction.Text, 4)
30        mvwDate.Month = Mid$(txthorsfonction.Text, 6, 2)
35        mvwDate.Day = Right$(txthorsfonction.Text, 2)
40      Else
45        mvwDate.Value = Date
50      End If
  
55      Call mvwDate.SetFocus

60      Exit Sub

AfficherErreur:

65      woups "frmoutils", "cmdDateHorsfonction_Click", Err, Erl
End Sub

Private Sub CmdFerme_Click()

5       On Error GoTo AfficherErreur

        'quitte fenetre
10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmoutils", "CmdFerme_Click", Err, Erl
End Sub

Private Sub cmdRechercher_Click()

5       On Error GoTo AfficherErreur

10      Dim rstRecherche As ADODB.Recordset
15      Dim itmOutils    As ListItem

20      Screen.MousePointer = vbHourglass

25      Call lstoutils.ListItems.Clear

30      Set rstRecherche = New ADODB.Recordset

35      Call rstRecherche.Open("SELECT * FROM GRB_Outils WHERE (Instr(1,CStr(no_outils),'" & txtRecherche.Text & "') > 0 OR Instr(1,nom_outils,'" & txtRecherche.Text & "') > 0) AND Departement = '" & Replace(cmbdepartement.Text, "'", "''") & "' ORDER BY no_outils", g_connData, adOpenDynamic, adLockOptimistic)

40      Do While Not rstRecherche.EOF
45        Set itmOutils = lstoutils.ListItems.Add
        
50        itmOutils.Text = rstRecherche.Fields("no_outils")
        
55        If Not IsNull(rstRecherche.Fields("nom_outils")) Then
60          itmOutils.SubItems(1) = rstRecherche.Fields("nom_outils")
65        End If
           
70        If Not IsNull(rstRecherche.Fields("date_achat")) Then
75          itmOutils.SubItems(2) = rstRecherche.Fields("date_achat")
80        End If
            
85        If Not IsNull(rstRecherche.Fields("date_hors_fonction")) Then
90          itmOutils.SubItems(3) = rstRecherche.Fields("date_hors_fonction")
95        End If
            
100       If Not IsNull(rstRecherche.Fields("cout")) Then
105         itmOutils.SubItems(4) = Conversion(rstRecherche.Fields("cout"), MODE_ARGENT)
110       End If
        
115       If Not IsNull(rstRecherche.Fields("type_étiquette")) Then
120         itmOutils.SubItems(5) = rstRecherche.Fields("type_étiquette")
125       End If
        
130       If Not IsNull(rstRecherche.Fields("commentaire")) Then
135         itmOutils.SubItems(6) = rstRecherche.Fields("commentaire")
140       End If
      
145       Call rstRecherche.MoveNext
150     Loop
    
155     Call rstRecherche.Close
160     Set rstRecherche = Nothing
  
165     Screen.MousePointer = vbDefault

170     Exit Sub

AfficherErreur:

175     woups "frmoutils", "cmdRechercher_Click", Err, Erl
End Sub

Private Sub cmdreport_Click()

5       On Error GoTo AfficherErreur

10      Dim rstOutils As ADODB.Recordset

15      Screen.MousePointer = vbHourglass

20      Set rstOutils = New ADODB.Recordset

25      Call rstOutils.Open("SELECT * FROM GRB_outils WHERE departement = '" & Replace(cmbdepartement.Text, "'", "''") & "' ORDER BY no_outils", g_connData, adOpenDynamic, adLockOptimistic)
    
        ''''''''''''''''''''''''''''''''''''''''''
        'rapport liste d'outil pour un departement
        ''''''''''''''''''''''''''''''''''''''''''
  
        'set le rapport
30      Set DR_Outils_machinerie.DataSource = rstOutils
    
        'contenu label
35      DR_Outils_machinerie.Sections("section2").Controls("lbldepartement").Caption = "Outils & machinerie " + LCase(rstOutils!departement)
40      DR_Outils_machinerie.Sections("section3").Controls("lbldate").Caption = ConvertDate(Date)
45      DR_Outils_machinerie.Orientation = rptOrientLandscape
    
        'affiche rapport
50      Call DR_Outils_machinerie.Show(vbModal)
    
55      Screen.MousePointer = vbDefault

60      Exit Sub

AfficherErreur:

65      woups "frmoutils", "cmdreport_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      Call RemplirDepartement

15      Exit Sub

AfficherErreur:

20      woups "frmoutils", "Form_Load", Err, Erl
End Sub

Public Sub RemplirDepartement()

5       On Error GoTo AfficherErreur

        '''''''''''''''''''''''''''''''''''''''''''''''''
        'remplis combo departement en mode visualisation
        '''''''''''''''''''''''''''''''''''''''''''''''''
10      Dim rstDepartement As ADODB.Recordset
  
15      Set rstDepartement = New ADODB.Recordset
  
20      Call rstDepartement.Open("SELECT DISTINCT departement FROM GRB_outils ORDER BY departement", g_connData, adOpenDynamic, adLockOptimistic)
  
25      Call cmbdepartement.Clear
  
        'rempli tant il y a des departement
30      Do While Not rstDepartement.EOF
35        Call cmbdepartement.AddItem(rstDepartement.Fields("departement"))

40        Call rstDepartement.MoveNext
45      Loop
     
        'ferme la table
50      Call rstDepartement.Close
55      Set rstDepartement = Nothing
  
        'si il y a des departement ,selectionne par defaut
60      If cmbdepartement.ListCount > 0 Then
65        cmbdepartement.ListIndex = 0
70      End If

75      Exit Sub

AfficherErreur:

80      woups "frmoutils", "RemplirDepartement", Err, Erl
End Sub

Public Sub RemplirEtiquette()

5       On Error GoTo AfficherErreur

        ''''''''''''''''''''''''''''''''''''''''''''',
        'remplis combo etiquette en mode modification
        ''''''''''''''''''''''''''''''''''''''''''''''
10      Dim rstEtiquette As ADODB.Recordset
  
15      Set rstEtiquette = New ADODB.Recordset
  
20      Call rstEtiquette.Open("SELECT DISTINCT type_étiquette FROM GRB_outils ORDER BY type_étiquette", g_connData, adOpenDynamic, adLockOptimistic)
  
25      Call cmbetiquette.Clear
  
        'rempli tant il y a des type_étiquette
30      Do While Not rstEtiquette.EOF
35        Call cmbetiquette.AddItem(rstEtiquette.Fields("type_étiquette"))

40        Call rstEtiquette.MoveNext
45      Loop
  
        'ferme la table
50      Call rstEtiquette.Close
55      Set rstEtiquette = Nothing

60      Exit Sub

AfficherErreur:

65      woups "frmoutils", "RemplirEtiquette", Err, Erl
End Sub

Public Sub RemplirDepartementModif()

5       On Error GoTo AfficherErreur

        '''''''''''''''''''''''''''''''''''''''''''''''
        'remplis combo departement en mode modification
        '''''''''''''''''''''''''''''''''''''''''''''''
10      Dim rstDepartement As ADODB.Recordset
  
15      Set rstDepartement = New ADODB.Recordset
  
20      Call rstDepartement.Open("SELECT DISTINCT departement FROM GRB_outils ORDER BY departement", g_connData, adOpenDynamic, adLockOptimistic)
    
25      Call cmbdepartement_modif.Clear
    
        'rempli tant il y a des employé
30      Do While Not rstDepartement.EOF
35        Call cmbdepartement_modif.AddItem(rstDepartement.Fields("departement"))

40        Call rstDepartement.MoveNext
45      Loop
    
        'ferme la table
50      Call rstDepartement.Close
55      Set rstDepartement = Nothing

60      Exit Sub

AfficherErreur:

65      woups "frmoutils", "RemplirDepartementModif", Err, Erl
End Sub

Public Sub RemplirListViewOutils()

5       On Error GoTo AfficherErreur

        'remplis lister une journée
10      Dim rstOutils As ADODB.Recordset
15      Dim itmOutils As ListItem
  
        'vide le lister
20      Call lstoutils.ListItems.Clear

25      lstoutils.Sorted = False

30      Set rstOutils = New ADODB.Recordset

35      Call rstOutils.Open("SELECT * FROM GRB_outils WHERE departement = '" & Replace(cmbdepartement.Text, "'", "''") & "' ORDER BY no_outils", g_connData, adOpenDynamic, adLockOptimistic)
    
        'tant il y a de employé cedulé , ajoute dans lister
40      Do While Not rstOutils.EOF
45        Set itmOutils = lstoutils.ListItems.Add
            
50        itmOutils.Text = rstOutils.Fields("no_outils")
        
55        If IsNull(rstOutils.Fields("nom_outils")) Then
60          Call itmOutils.ListSubItems.Add(, , vbNullString)
65        Else
70          Call itmOutils.ListSubItems.Add(, , rstOutils.Fields("nom_outils"))
75        End If
            
80        If IsNull(rstOutils.Fields("date_achat")) Then
85          Call itmOutils.ListSubItems.Add(, , vbNullString)
90        Else
95          Call itmOutils.ListSubItems.Add(, , rstOutils.Fields("date_achat"))
100       End If
            
105       If IsNull(rstOutils.Fields("date_hors_fonction")) Then
110         Call itmOutils.ListSubItems.Add(, , vbNullString)
115       Else
120         Call itmOutils.ListSubItems.Add(, , rstOutils.Fields("date_hors_fonction"))
125       End If
            
130       If IsNull(rstOutils.Fields("cout")) Or rstOutils.Fields("Cout") = vbNullString Then
135         Call itmOutils.ListSubItems.Add(, , vbNullString)
140       Else
145         Call itmOutils.ListSubItems.Add(, , Conversion(rstOutils.Fields("cout"), MODE_ARGENT))
150       End If
        
155       If IsNull(rstOutils.Fields("type_étiquette")) Then
160         Call itmOutils.ListSubItems.Add(, , vbNullString)
165       Else
170         Call itmOutils.ListSubItems.Add(, , rstOutils.Fields("type_étiquette"))
175       End If
        
180       If IsNull(rstOutils.Fields("commentaire")) Then
185         Call itmOutils.ListSubItems.Add(, , vbNullString)
190       Else
195         Call itmOutils.ListSubItems.Add(, , rstOutils.Fields("commentaire"))
200       End If
            
205       Call rstOutils.MoveNext
210     Loop
    
        'ferme la table
215     Call rstOutils.Close
220     Set rstOutils = Nothing

225     Exit Sub

AfficherErreur:

230     woups "frmoutils", "RemplirListViewOutils", Err, Erl
End Sub

Private Sub cmbdepartement_Click()

5       On Error GoTo AfficherErreur

10      Screen.MousePointer = vbHourglass
  
        'remplis le lister outil dependant le departement selectionné
15      Call RemplirListViewOutils
  
20      Screen.MousePointer = vbDefault

25      Exit Sub

AfficherErreur:

30      woups "frmoutils", "cmbdepartement_Click", Err, Erl
End Sub

Private Sub CmdAdd_Click()

5       On Error GoTo AfficherErreur

10      Screen.MousePointer = vbHourglass

        'proc qui permet d'ajouter un outils
15      m_bModeAjouter = True
  
        'affiche en mode modification
20      Call AfficherModif
  
        'remplis les combo en mode modification
25      Call RemplirEtiquette
30      Call RemplirDepartementModif
  
        'vide les champs
35      txtachat.Text = vbNullString
40      txtcommentaire.Text = vbNullString
45      txtcout.Text = vbNullString
50      txthorsfonction.Text = vbNullString
55      txtNo.Text = vbNullString
60      txtoutils.Text = vbNullString
  
65      Screen.MousePointer = vbDefault

70      Exit Sub

AfficherErreur:

75      woups "frmoutils", "CmdAdd_Click", Err, Erl
End Sub

Private Sub AfficherModif()

5       On Error GoTo AfficherErreur
        
        'met visible les champ pour modifié ou ajouté
10      fraModif.Visible = True
15      lstoutils.Visible = False
20      cmbdepartement.Visible = False
25      lbldepartement.Visible = False
30      CmdAdd.Visible = False
35      CmdSupp.Visible = False
40      CmdModif.Visible = False
45      CmdEnr.Visible = True
50      CmdAnul.Visible = True
55      lblRecherche.Visible = False
60      txtRecherche.Visible = False
65      cmdRechercher.Visible = False
70      cmdReport.Visible = False
75      CmdFerme.Visible = False

80      Exit Sub

AfficherErreur:

85      woups "frmoutils", "AfficherModif", Err, Erl
End Sub

Private Sub AfficherListe()

5       On Error GoTo AfficherErreur

        'met visible les champ pour modifié ou ajouté
10      fraModif.Visible = False
15      lstoutils.Visible = True
20      cmbdepartement.Visible = True
25      lbldepartement.Visible = True
30      CmdAdd.Visible = True
35      CmdSupp.Visible = True
40      CmdModif.Visible = True
45      CmdEnr.Visible = False
50      CmdAnul.Visible = False
55      lblRecherche.Visible = True
60      txtRecherche.Visible = True
65      cmdRechercher.Visible = True
70      cmdReport.Visible = True
75      CmdFerme.Visible = True

80      Exit Sub

AfficherErreur:

85      woups "frmoutils", "AfficherListe", Err, Erl
End Sub

Private Sub CmdModif_Click()

5       On Error GoTo AfficherErreur

10      Dim rstOutils As ADODB.Recordset
  
15      Screen.MousePointer = vbHourglass
  
        'proc qui permet d'ajouter un outils
20      m_bModeAjouter = False
  
        'affiche en mode modification
25      Call AfficherModif
  
        'remplis les combo en mode modification
30      Call RemplirEtiquette
35      Call RemplirDepartementModif

40      If lstoutils.ListItems.count > 0 Then
          'ouvre la table
45        Set rstOutils = New ADODB.Recordset
          
50        Call rstOutils.Open("SELECT * FROM GRB_outils WHERE no_outils = '" & lstoutils.SelectedItem.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
55        frmoutils.cmbdepartement_modif.Text = rstOutils.Fields("departement")
60        frmoutils.cmbetiquette.Text = rstOutils.Fields("type_étiquette")
65        frmoutils.txtNo.Text = rstOutils.Fields("no_outils")
    
70        If IsNull(rstOutils.Fields("nom_outils")) Then
75          txtoutils.Text = vbNullString
80        Else
85          txtoutils.Text = rstOutils.Fields("nom_outils")
90        End If
      
95        If IsNull(rstOutils.Fields("cout")) Then
100         txtcout.Text = vbNullString
105       Else
110         txtcout.Text = rstOutils.Fields("cout")
115       End If
      
120       If Not IsNull(rstOutils.Fields("date_achat")) Then
125         txtachat.Text = rstOutils.Fields("date_achat")
130       Else
135         txtachat.Text = vbNullString
140       End If
    
145       If Not IsNull(rstOutils.Fields("date_hors_fonction")) Then
150         txthorsfonction.Text = rstOutils.Fields("date_hors_fonction")
155       Else
160         txthorsfonction.Text = vbNullString
165       End If

170       If IsNull(rstOutils.Fields("commentaire")) Then
175         txtcommentaire.Text = vbNullString
180       Else
185         txtcommentaire.Text = rstOutils.Fields("commentaire")
190       End If
      
          'ferme la table
195       Call rstOutils.Close
200       Set rstOutils = Nothing
205     End If

210     Screen.MousePointer = vbDefault

215     Exit Sub

AfficherErreur:

220     woups "frmoutils", "CmdModif_Click", Err, Erl
End Sub

Private Sub CmdSupp_Click()

5       On Error GoTo AfficherErreur

        ''''''''''''''''''''''''''''''''
        'Supprime l'outils selectionné
        ''''''''''''''''''''''''''''''''
10      If lstoutils.ListItems.count > 0 Then
15        If MsgBox("Voulez-vous supprimer cette enregistrement?", vbYesNo, "Supprimer") = vbYes Then
20          Screen.MousePointer = vbHourglass
    
25          Call g_connData.Execute("DELETE * FROM GRB_outils WHERE no_outils = '" & lstoutils.SelectedItem.Text & "'")
        
            'mise a jour des lister
30          Call RemplirListViewOutils
        
35          Call RemplirDepartement
      
40          Screen.MousePointer = vbDefault
45        End If
50      End If

55      Exit Sub

AfficherErreur:

60      woups "frmoutils", "CmdSupp_Click", Err, Erl
End Sub

Private Sub CmdEnr_Click()

5       On Error GoTo AfficherErreur

        'enregistre
10      Dim rstOutils     As ADODB.Recordset
15      Dim rstVerifModif As ADODB.Recordset

20      Screen.MousePointer = vbHourglass
  
25      If cmbdepartement.Text = vbNullString Or cmbetiquette.Text = vbNullString Or txtNo.Text = vbNullString Then
30        Call MsgBox("Champs vide!", vbOKOnly, "Erreur")
    
35        Screen.MousePointer = vbDefault

40        Exit Sub
45      End If
  
50      If (Len(txtachat.Text) = 0 Or (Len(txtachat.Text) > 1 And IsDate(txtachat.Text))) And (Len(txthorsfonction.Text) = 0 Or (Len(txthorsfonction.Text) > 1 And IsDate(txthorsfonction.Text))) Then
          'ouvre la table
55        Set rstOutils = New ADODB.Recordset
          
60        If m_bModeAjouter = True Then
65          Call rstOutils.Open("SELECT * FROM GRB_outils WHERE no_outils = '" & txtNo.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
        
70          If Not rstOutils.EOF Then
75            Call MsgBox("Le numéro d'outils existe déjà!", vbOKOnly, "Erreur")
          
80            Call rstOutils.Close
85            Set rstOutils = Nothing
          
90            Screen.MousePointer = vbDefault
          
95            Exit Sub
100         End If
        
105         Call rstOutils.AddNew
110       Else
115         Call rstOutils.Open("SELECT * FROM GRB_outils WHERE no_outils = '" & txtNo.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
120       End If
                       
          ''''''''''''''''''''''''''
          'ajoute l'enregistrement
          ''''''''''''''''''''''''''
125       rstOutils.Fields("departement") = cmbdepartement_modif.Text
130       rstOutils.Fields("type_étiquette") = cmbetiquette.Text
135       rstOutils.Fields("no_outils") = txtNo.Text
140       rstOutils.Fields("nom_outils") = txtoutils.Text
145       rstOutils.Fields("cout") = txtcout.Text
150       rstOutils.Fields("date_achat") = txtachat.Text
155       rstOutils.Fields("date_hors_fonction") = txthorsfonction.Text
160       rstOutils.Fields("commentaire") = txtcommentaire.Text
                
165       Call rstOutils.Update
                        
          'quitte ecran pour ajouté ou modifié
170       Call AfficherListe
                          
175       Call rstOutils.Close
180       Set rstOutils = Nothing
                                  
          'met a jour l'écran
185       Call RemplirListViewOutils
      
190       Call RemplirDepartement
195     Else
200       Call MsgBox("La date est invalide! (aaaa-mm-jj)", , "Erreur")
205     End If
  
210     Screen.MousePointer = vbDefault

215     Exit Sub

AfficherErreur:

220     woups "frmoutils", "CmdEnr_Click", Err, Erl
End Sub

Private Sub lstoutils_DblClick()

5       On Error GoTo AfficherErreur

        'affiche l'écran en mode modification
10      Call CmdModif_Click

15      Exit Sub

AfficherErreur:

20      woups "frmoutils", "lstoutils_DblClick", Err, Erl
End Sub

Private Sub lstoutils_KeyDown(KeyCode As Integer, Shift As Integer)

5       On Error GoTo AfficherErreur

10      If lstoutils.ListItems.count > 0 Then
15        If KeyCode = vbKeyDelete Then
20          If MsgBox("Voulez-vous supprimer cette enregistrement?", vbYesNo, "Supprimer") = vbYes Then
25            Call g_connData.Execute("DELETE * FROM GRB_outils WHERE no_outils = '" & lstoutils.SelectedItem.Text & "'")
        
              'mise a jour des lister
30            Call RemplirListViewOutils
35          End If
40        End If
45      End If

50      Exit Sub

AfficherErreur:

55      woups "frmoutils", "lstoutils_KeyDown", Err, Erl
End Sub

Private Sub mvwDate_DateClick(ByVal DateClicked As Date)

5       On Error GoTo AfficherErreur

        '''''''''''''''''''''''''''''''''''''''''
        'affiche dans l'écran la date sélectionné
        '''''''''''''''''''''''''''''''''''''''''
10      Dim sDate As String
  
        'ajoute dans le champ text la date selectionné
15      If m_bDateAchat = True Then
20        txtachat.Text = ConvertDate(DateClicked)
25      Else
30        txthorsfonction.Text = ConvertDate(DateClicked)
35      End If
  
40      mvwDate.Visible = False

45      Exit Sub

AfficherErreur:

50      woups "frmoutils", "mvwDate_DateClick", Err, Erl
End Sub

Private Sub mvwDate_LostFocus()

5       On Error GoTo AfficherErreur

        ''''''''''''''''''''''''''''''
        'lorsque clique ailleur cache
        '''''''''''''''''''''''''''''''
10      mvwDate.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmoutils", "mvwDate_LostFocus", Err, Erl
End Sub

Private Sub txtachat_GotFocus()

5       On Error GoTo AfficherErreur

        'met l'année 2caratere
10      If Len(txtachat.Text) = 10 Then
15        txtachat.Text = Right$(txtachat.Text, 8)
20      End If
  
        'met le mask
25      txtachat.mask = "##-##-##"

30      Exit Sub

AfficherErreur:

35      woups "frmoutils", "txtachat_GotFocus", Err, Erl
End Sub

Private Sub txtachat_LostFocus()

5       On Error GoTo AfficherErreur

        'enleve le mask
10      txtachat.mask = vbNullString
  
        'losque est vide , enleve les caratere du masque
15      If txtachat.Text = "__-__-__" Then
20        txtachat.Text = vbNullString
25      End If
  
        'met l'année 4 caractere
30      If Len(txtachat.Text) = 8 Then
35        If IsDate(txtachat.Text) Then
40          txtachat.Text = Trim$(Year(DateSerial(Mid$(txtachat.Text, 1, 2), Mid$(txtachat.Text, 4, 2), Mid$(txtachat.Text, 7, 2)))) + Mid$(txtachat.Text, 3, 8)
45        End If
50      End If

55      Exit Sub

AfficherErreur:

60      woups "frmoutils", "txtachat_LostFocus", Err, Erl
End Sub

Private Sub txtcout_LostFocus()

5       On Error GoTo AfficherErreur

10      txtcout.Text = Replace(txtcout.Text, ".", ",")

15      If Not IsNumeric(txtcout.Text) Then
20        Call MsgBox("Valeur non numérique!", vbOKOnly, "Erreur")

25        txtcout.Text = ""
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmoutils", "txtcout_LostFocus", Err, Erl
End Sub

Private Sub txthorsfonction_GotFocus()

5       On Error GoTo AfficherErreur

        'met l'année 2caratere
10      If Len(txthorsfonction.Text) = 10 Then
15        txthorsfonction.Text = Right$(txthorsfonction.Text, 8)
20      End If

        'met le mask
25      txthorsfonction.mask = "##-##-##"

30      Exit Sub

AfficherErreur:

35      woups "frmoutils", "txthorsfonction_GotFocus", Err, Erl
End Sub

Private Sub txthorsfonction_LostFocus()

5       On Error GoTo AfficherErreur

        'enleve le mask
10      txthorsfonction.mask = vbNullString
  
        'losque est vide , enleve les caratere du masque
15      If txthorsfonction.Text = "__-__-__" Then
20        txthorsfonction.Text = vbNullString
25      End If
  
        'met l'année 4 caractere
30      If Len(txthorsfonction.Text) = 8 Then
35        If IsDate(txthorsfonction.Text) Then
40          txthorsfonction.Text = Trim$(Year(DateSerial(Mid$(txthorsfonction.Text, 1, 2), Mid$(txthorsfonction.Text, 4, 2), Mid$(txthorsfonction.Text, 7, 2)))) + Mid$(txthorsfonction.Text, 3, 8)
45        End If
50      End If

55      Exit Sub

AfficherErreur:

60     woups "frmoutils", "txthorsfonction_LostFocus", Err, Erl
End Sub
