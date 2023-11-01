VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmOutils_InOut 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Magasin"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9315
   Icon            =   "FrmOutils_InOut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "FrmOutils_InOut.frx":0442
   ScaleHeight     =   5730
   ScaleWidth      =   9315
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chknonRetour 
      BackColor       =   &H00000000&
      Caption         =   "Outils non retournés"
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
      Height          =   375
      Left            =   6960
      TabIndex        =   19
      Top             =   840
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CommandButton cmdsortie 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Sortie Outils"
      Height          =   495
      Left            =   1680
      TabIndex        =   22
      Top             =   5040
      Width           =   1455
   End
   Begin VB.ComboBox cmbemployé 
      Height          =   315
      Left            =   4440
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton cmdConfig 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Configuration Outils"
      Height          =   495
      Left            =   5760
      TabIndex        =   24
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton CmdFerme 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Fermer"
      Height          =   495
      Left            =   7680
      TabIndex        =   25
      Top             =   5040
      Width           =   1455
   End
   Begin MSComctlLib.ListView lstoutils 
      Height          =   3375
      Left            =   120
      TabIndex        =   20
      Top             =   1440
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
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
         Text            =   "Employé"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Sortie"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Retour"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Département"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Commentaire"
         Object.Width           =   5292
      EndProperty
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
      TabIndex        =   21
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdretour 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Retour Outils"
      Height          =   495
      Left            =   3240
      TabIndex        =   23
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Frame fraOutils 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   9015
      Begin VB.TextBox txtcommentaire 
         Height          =   765
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   3240
         Width           =   4575
      End
      Begin VB.TextBox txtemploye 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   15
         Top             =   2760
         Width           =   1815
      End
      Begin VB.CommandButton CmdAnul 
         Caption         =   "&Afficher liste"
         Height          =   495
         Left            =   7080
         TabIndex        =   11
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton CmdEnr 
         Caption         =   "&Enregistrer"
         Height          =   495
         Left            =   7080
         TabIndex        =   8
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtdepartement 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtsortie 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txtdepart 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   9
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtnom 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   4
         Top             =   840
         Width           =   4815
      End
      Begin VB.TextBox txtno 
         Height          =   285
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label7 
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
         Left            =   600
         TabIndex        =   16
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Employé"
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
         Left            =   600
         TabIndex        =   14
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label5 
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
         Left            =   600
         TabIndex        =   7
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Retour"
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
         Left            =   600
         TabIndex        =   12
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Sortie"
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
         Left            =   600
         TabIndex        =   10
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nom"
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
         Left            =   600
         TabIndex        =   5
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "No Outil"
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
         Left            =   600
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Label lblemployé 
      BackStyle       =   0  'Transparent
      Caption         =   "Employé"
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
      Left            =   4440
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "FrmOutils_InOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chknonRetour_Click()

5       On Error GoTo AfficherErreur

10      Screen.MousePointer = vbHourglass
  
        'remplis le lister dependant si retourné ou pas
15      Call remplir_lister_outils
  
20      Screen.MousePointer = vbDefault

25      Exit Sub

AfficherErreur:

30      woups "frmOutils_InOut", "chknonRetour_Click", Err, Erl
End Sub

Private Sub cmbEmployé_Click()

5       On Error GoTo AfficherErreur

10      Screen.MousePointer = vbHourglass

        'remplis le lister dependant l'employé selectionné
15      Call remplir_lister_outils
  
20      Screen.MousePointer = vbDefault

25      Exit Sub

AfficherErreur:

30      woups "frmOutils_InOut", "cmbEmployé_Click", Err, Erl
End Sub

Private Sub CmdAnul_Click()

5       On Error GoTo AfficherErreur

        'affiche mode liste
10      fraOutils.Visible = False
15      lstoutils.Visible = True
20      cmbemployé.Visible = True
25      lblemployé.Visible = True
30      chknonRetour.Visible = True

        'remplis le lister
35      Call remplir_lister_outils

40      Exit Sub

AfficherErreur:

45      woups "frmOutils_InOut", "CmdAnul_Click", Err, Erl
End Sub

Private Sub cmdConfig_Click()

5       On Error GoTo AfficherErreur

10      Screen.MousePointer = vbHourglass
  
        'affiche ecran pour ajouter modifier un outils
15      Call OuvrirForm(frmoutils, True)
  
20      Screen.MousePointer = vbDefault

25      Exit Sub

AfficherErreur:

30      woups "frmOutils_InOut", "cmdConfig_Click", Err, Erl
End Sub

Private Sub CmdEnr_Click()

5       On Error GoTo AfficherErreur

        'enregistre
10      Dim rstOutils      As ADODB.Recordset
15      Dim rstOutilsInOut As ADODB.Recordset

20      Screen.MousePointer = vbHourglass

25      If IsNumeric(txtNo.Text) Then
          'ouvre la table
30        Set rstOutils = New ADODB.Recordset
          
35        Call rstOutils.Open("SELECT * FROM GRB_outils WHERE no_outils = '" & txtNo.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
               
          'si existe on peut l'ajouter
40        If Not rstOutils.EOF Then
            'ouvre la table avec une recherche sur l'outils étant non-retourné
45          Set rstOutilsInOut = New ADODB.Recordset
            
50          Call rstOutilsInOut.Open("SELECT GRB_Outils_In_out.*, GRB_employés.employe FROM GRB_Outils_In_out INNER JOIN GRB_employés ON GRB_Outils_In_out.no_employe = GRB_employés.noemploye WHERE no_outils = " & txtNo.Text & " AND retour_date_heure is null", g_connData, adOpenDynamic, adLockOptimistic)
        
55          If rstOutilsInOut.EOF Then
              'ajoute
60            Call rstOutilsInOut.AddNew
                
65            rstOutilsInOut!no_outils = txtNo.Text
70            rstOutilsInOut!no_employe = txtEmploye.Tag
75            rstOutilsInOut!depart_date_heure = txtdepart.Text
80            rstOutilsInOut!commentaire = txtcommentaire.Text
          
85            Call rstOutilsInOut.Update
          
              'vide les champs
90            Call vider_champs
95          Else
100           If MsgBox("L'outils n'a pas été retourné par " & rstOutilsInOut.Fields("employe") & ", voulez-vous retourner l'outil pour cet employé?", vbYesNo, CStr(rstOutils!no_outils) + " " + rstOutils!nom_outils) = vbYes Then
                'retourne l'outils pour l'employé
105             rstOutilsInOut!commentaire = "(Retourné par: " + CStr(g_sEmploye) + ") " + CStr(rstOutilsInOut!commentaire)
110             rstOutilsInOut!retour_date_heure = CStr(Year(Now)) + "-" + Right$("0" + CStr(Month(Now)), 2) + "-" + Right$("0" + CStr(Day(Now)), 2) + " " + CStr(Time)
          
115             Call rstOutilsInOut.Update
          
                'ajoute
120             Call rstOutilsInOut.AddNew
                   
125             rstOutilsInOut!no_outils = txtNo.Text
130             rstOutilsInOut!no_employe = txtEmploye.Tag
135             rstOutilsInOut!depart_date_heure = txtdepart.Text
140             rstOutilsInOut!commentaire = txtcommentaire.Text
          
145             Call rstOutilsInOut.Update
            
                'vide les champs
150             Call vider_champs
155           End If
160         End If
          
            'quitte la table
165         Call rstOutilsInOut.Close
170         Set rstOutilsInOut = Nothing
                      
175         Call txtNo.SetFocus
180       Else
185         Call MsgBox("Le numéro de l'outil n'existe pas!", , "Erreur")
190       End If
195     Else
200       Call MsgBox("Le numéro doit être numérique!", , "Erreur")
205     End If
  
210     Screen.MousePointer = vbDefault

215     Exit Sub

AfficherErreur:

220     woups "frmOutils_InOut", "CmdEnr_Click", Err, Erl
End Sub

Private Sub CmdFerme_Click()

5       On Error GoTo AfficherErreur

        'quitte la fenetre
10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmOutils_InOut", "CmdFerme_Click", Err, Erl
End Sub

Public Sub remplir_lister_outils()

5       On Error GoTo AfficherErreur

        'remplis lister une journée
10      Dim rstOutils      As ADODB.Recordset
15      Dim rstOutilsInOut As ADODB.Recordset
20      Dim rstEmploye     As ADODB.Recordset
25      Dim itmOutils      As ListItem
  
        'vide le lister
30      Call lstoutils.ListItems.Clear
35      lstoutils.Sorted = False

40      Set rstOutilsInOut = New ADODB.Recordset
        'affiche tout les employes si = *
45      If cmbemployé.Text = "*" Then
          'affiche les non-retournés si est coché, sinon affiche retourné et non
50        If chknonRetour.Value = vbChecked Then
55          Call rstOutilsInOut.Open("SELECT * FROM GRB_outils_in_out WHERE retour_date_heure is null ORDER BY no_enreg", g_connData, adOpenDynamic, adLockOptimistic)
60        Else
65          Call rstOutilsInOut.Open("SELECT * FROM GRB_outils_in_out ORDER BY no_enreg", g_connData, adOpenDynamic, adLockOptimistic)
70        End If
75      Else
          'affiche pour un employé et
          'affiche les non-retourné si est coché, sinon affiche retourné et non retourné
80        If chknonRetour.Value = vbChecked Then
85          Call rstOutilsInOut.Open("SELECT * FROM GRB_outils_in_out WHERE CStr(no_employe) = CStr('" & cmbemployé.ItemData(cmbemployé.ListIndex) & "') AND retour_date_heure is null ORDER BY no_enreg", g_connData, adOpenDynamic, adLockOptimistic)
90        Else
95          Call rstOutilsInOut.Open("SELECT * FROM GRB_outils_in_out WHERE CStr(no_employe) = CStr('" & cmbemployé.ItemData(cmbemployé.ListIndex) & "') ORDER BY no_enreg", g_connData, adOpenDynamic, adLockOptimistic)
100       End If
105     End If
    
110     Set rstOutils = New ADODB.Recordset
115     Set rstEmploye = New ADODB.Recordset
    
120     Do While Not rstOutilsInOut.EOF
125       Call rstOutils.Open("SELECT * FROM GRB_outils WHERE CStr(no_outils) = CStr('" & rstOutilsInOut!no_outils & "') ORDER BY no_outils", g_connData, adOpenDynamic, adLockOptimistic)
130       Call rstEmploye.Open("SELECT * FROM GRB_employés WHERE CStr(noemploye) = CStr('" & rstOutilsInOut!no_employe & "') ORDER BY noemploye", g_connData, adOpenDynamic, adLockOptimistic)
        
          'ajoute dans lister
135       Set itmOutils = lstoutils.ListItems.Add
            
140       itmOutils.Text = (rstOutils!no_outils)
145       itmOutils.Tag = rstOutilsInOut!no_enreg
    
150       Call itmOutils.ListSubItems.Add(, , rstOutils!nom_outils)
155       Call itmOutils.ListSubItems.Add(, , rstEmploye!employe)
160       itmOutils.ListSubItems(2).Tag = rstEmploye!loginname
165       Call itmOutils.ListSubItems.Add(, , rstOutilsInOut!depart_date_heure)
        
170       If IsNull(rstOutilsInOut!retour_date_heure) Then
175         Call itmOutils.ListSubItems.Add(, , vbNullString)
180       Else
185         Call itmOutils.ListSubItems.Add(, , rstOutilsInOut!retour_date_heure)
190       End If
        
195       Call itmOutils.ListSubItems.Add(, , rstOutils!departement)
200       Call itmOutils.ListSubItems.Add(, , rstOutilsInOut!commentaire)

205       Call rstOutilsInOut.MoveNext
     
          'ferme la table
210       Call rstOutils.Close
215       Call rstEmploye.Close
220     Loop
      
225     Call lstoutils.Refresh
    
230     Set rstOutils = Nothing
235     Set rstEmploye = Nothing

        'ferme la table
240     Call rstOutilsInOut.Close
245     Set rstOutilsInOut = Nothing

250     Exit Sub

AfficherErreur:

255     woups "frmOutils_InOut", "remplir_lister_outils", Err, Erl
End Sub

Public Sub remplir_cmbemploye()

5       On Error GoTo AfficherErreur

        ''''''''''''''''''''''''''''''''''''''''''''',
        'remplis combo etiquette en mode modification
        ''''''''''''''''''''''''''''''''''''''''''''''
10      Dim rstEmploye As ADODB.Recordset
  
15      Set rstEmploye = New ADODB.Recordset
  
20      Call rstEmploye.Open("SELECT NoEmploye, Employe FROM GRB_employés WHERE Actif = true ORDER BY employe", g_connData, adOpenDynamic, adLockOptimistic)
  
25      Call cmbemployé.Clear
  
30      Call cmbemployé.AddItem("*")
        'rempli tant il y a des type_étiquette
    
35      Do While Not rstEmploye.EOF
40        Call cmbemployé.AddItem(rstEmploye!employe)
45        cmbemployé.ItemData(cmbemployé.newIndex) = rstEmploye!noEmploye

50        Call rstEmploye.MoveNext
55      Loop

        'ferme la table
60      Call rstEmploye.Close
65      Set rstEmploye = Nothing

70      Exit Sub

AfficherErreur:

75      woups "frmOutils_InOut", "remplir_cmbemploye", Err, Erl
End Sub

Private Sub cmdreport_Click()

5       On Error GoTo AfficherErreur

        'affiche rapport filtré par les selection a l'ecran
10      Dim rstOutilsInOut As ADODB.Recordset

15      Set rstOutilsInOut = New ADODB.Recordset

        'si tout les employé
20      If cmbemployé.Text = "*" Then
          'affiche les non-retournés si est coché, sinon affiche retourné et non
25        If chknonRetour.Value = vbChecked Then
30          Call rstOutilsInOut.Open("SELECT GRB_Outils_In_out.*, GRB_employés.employe, GRB_Outils.nom_outils, GRB_Outils.departement FROM GRB_employés INNER JOIN (GRB_Outils_In_out INNER JOIN GRB_Outils ON CStr(GRB_Outils_In_out.no_outils) = GRB_Outils.no_outils) ON GRB_employés.noemploye = GRB_Outils_In_out.no_employe WHERE retour_date_heure is null ORDER BY no_enreg", g_connData, adOpenDynamic, adLockOptimistic)
35        Else
40          Call rstOutilsInOut.Open("SELECT GRB_Outils_In_out.*, GRB_employés.employe, GRB_Outils.nom_outils, GRB_Outils.departement FROM GRB_employés INNER JOIN (GRB_Outils_In_out INNER JOIN GRB_Outils ON CStr(GRB_Outils_In_out.no_outils) = GRB_Outils.no_outils) ON GRB_employés.noemploye = GRB_Outils_In_out.no_employe ORDER BY no_enreg", g_connData, adOpenDynamic, adLockOptimistic)
45        End If
50      Else
          'affiche pour un employé et
          'affiche les non-retourné si est coché, sinon affiche retourné et non retourné
55        If chknonRetour.Value = vbChecked Then
60          Call rstOutilsInOut.Open("SELECT GRB_Outils_In_out.*, GRB_employés.employe, GRB_Outils.nom_outils, GRB_Outils.departement FROM GRB_employés INNER JOIN (GRB_Outils_In_out INNER JOIN GRB_Outils ON CStr(GRB_Outils_In_out.no_outils) = GRB_Outils.no_outils) ON GRB_employés.noemploye = GRB_Outils_In_out.no_employe WHERE CStr(no_employe) = CStr('" & cmbemployé.ItemData(cmbemployé.ListIndex) & "') AND retour_date_heure is null ORDER BY no_enreg", g_connData, adOpenDynamic, adLockOptimistic)
65        Else
70          Call rstOutilsInOut.Open("SELECT GRB_Outils_In_out.*, GRB_employés.employe, GRB_Outils.nom_outils, GRB_Outils.departement FROM GRB_employés INNER JOIN (GRB_Outils_In_out INNER JOIN GRB_Outils ON CStr(GRB_Outils_In_out.no_outils) = GRB_Outils.no_outils) ON GRB_employés.noemploye = GRB_Outils_In_out.no_employe WHERE CStr(no_employe) = CStr('" & cmbemployé.ItemData(cmbemployé.ListIndex) & "') ORDER BY no_enreg", g_connData, adOpenDynamic, adLockOptimistic)
75        End If
80      End If

        ''''''''''''''''''''''''''''''''''''''''''
        'rapport liste d'outil pour un departement
        ''''''''''''''''''''''''''''''''''''''''''
        'set le rapport
85      Set DR_Outils_in_out.DataSource = rstOutilsInOut
    
        'contenu label
90      If chknonRetour.Value = vbChecked Then
95        DR_Outils_in_out.Sections("section2").Controls("lbldepartement").Caption = "Liste des outils non-retourné"
100     Else
105       DR_Outils_in_out.Sections("section2").Controls("lbldepartement").Caption = "Liste des outils empruntés"
110     End If
    
115     DR_Outils_in_out.Sections("section3").Controls("lbldate").Caption = CStr(Year(Now)) + "-" + CStr(Month(Now)) + "-" + CStr(Day(Now))
   
120     DR_Outils_in_out.Orientation = rptOrientLandscape
    
        'affiche rapport
125     Call DR_Outils_in_out.Show(vbModal)

130     Exit Sub

AfficherErreur:

135     woups "frmOutils_InOut", "cmdreport_Click", Err, Erl
End Sub

Private Sub cmdRetour_Click()

5       On Error GoTo AfficherErreur

        '''''''''''''''''''''''''''''''''''
        'retourne les outils selectionné
        '''''''''''''''''''''''''''''''''''
10      Dim rstOutilsInOut As ADODB.Recordset
15      Dim iCompteur      As Integer
  
20      Screen.MousePointer = vbHourglass
  
25      Set rstOutilsInOut = New ADODB.Recordset
  
        'tant que pas fin du lister
30      For iCompteur = 1 To lstoutils.ListItems.count
          'si l'enreg du lister est selectionné, on modifie dans la bd
35        If lstoutils.ListItems(iCompteur).Selected = True Then
            'ouvre la table sur l'enreg selectionné dans le lister
        
40          Call rstOutilsInOut.Open("SELECT * FROM GRB_outils_in_out WHERE CStr(no_enreg) = CStr('" & lstoutils.ListItems(iCompteur).Tag & "') ORDER BY no_outils", g_connData, adOpenDynamic, adLockOptimistic)
            'met la date d'aujourd'hui si pas null
      
45          If IsNull(rstOutilsInOut!retour_date_heure) Then
              'si c'est un autre employé, on ajoute dans commentaire l'employé qui a retourné l'outil
50            If g_sUserID <> lstoutils.ListItems(iCompteur).ListSubItems(2).Tag Then
55              rstOutilsInOut.Fields("Commentaire") = "(Retourné par: " + g_sEmploye + ")" + rstOutilsInOut.Fields("Commentaire")
60            End If
                
65            rstOutilsInOut!retour_date_heure = CStr(Year(Now)) + "-" + Right$("0" + CStr(Month(Now)), 2) + "-" + Right$("0" + CStr(Day(Now)), 2) + " " + CStr(Time)
         
70            Call rstOutilsInOut.Update
75          End If
        
            'ferme la table
80          Call rstOutilsInOut.Close
85          Set rstOutilsInOut = Nothing
90        End If
95      Next
    
        'met a jour le lister
100     Call remplir_lister_outils
  
105     Screen.MousePointer = vbDefault

110     Exit Sub

AfficherErreur:

115     woups "frmOutils_InOut", "cmdretour_Click", Err, Erl
End Sub

Private Sub cmdsortie_Click()

5       On Error GoTo AfficherErreur

10      Screen.MousePointer = vbHourglass
  
              'affiche en mode ajout
15      fraOutils.Visible = True
20      lstoutils.Visible = False
25      cmbemployé.Visible = False
30      lblemployé.Visible = False
35      chknonRetour.Visible = False
40      Call txtNo.SetFocus
  
45      Screen.MousePointer = vbDefault

50      Exit Sub

AfficherErreur:

55      woups "frmOutils_InOut", "cmdsortie_Click", Err, Erl
End Sub

Private Sub ActiverBoutonsGroupe()

5       On Error GoTo AfficherErreur

10      cmdConfig.Enabled = g_bModificationOutils
  
15      Exit Sub

AfficherErreur:

20      woups "frmOutils_InOut", "ActiverBoutonsGroupe", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      Dim rstEmploye As ADODB.Recordset
15      Dim iCompteur  As Integer
20      Dim iNoEmploye As Integer
25      Dim bEmpExiste As Boolean
  
30      Call Unload(frmChoixInventaire)
  
        'rempli le combo employe
35      Call remplir_cmbemploye
  
        'si il y a des employes ,selectionne par defaut
40      If cmbemployé.ListCount > 0 Then
45        Set rstEmploye = New ADODB.Recordset

50        Call rstEmploye.Open("SELECT * FROM GRB_employés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)
      
          'select l'employé logué sinon, affiche tout
55        If rstEmploye.EOF Then
60          bEmpExiste = False
65        Else
70          bEmpExiste = True
        
75          iNoEmploye = rstEmploye.Fields("noEmploye")
80        End If
      
85        Call rstEmploye.Close
90        Set rstEmploye = Nothing
      
95        If bEmpExiste = False Then
100         cmbemployé.ListIndex = 0
105       Else
            'on trouve l'index dans le combo qui contient le noemploye, pour le selectionné
110         For iCompteur = 0 To cmbemployé.ListCount - 1
              'si trouvé noemploye, on selectione cet index
115           If cmbemployé.ItemData(iCompteur) = iNoEmploye Then
120             cmbemployé.ListIndex = iCompteur
                  
125             Exit For
130           End If
135         Next
140       End If
145     End If
  
              'met enabled le bouton config dependant le user
150     Call ActiverBoutonsGroupe

155     Exit Sub

AfficherErreur:

160     woups "frmOutils_InOut", "Form_Load", Err, Erl
End Sub

Private Sub txtno_LostFocus()

5       On Error GoTo AfficherErreur


        'affiche automatique le nom de l'outils et l'employé et la date de sortie
10      Dim rstOutils  As ADODB.Recordset
15      Dim rstEmploye As ADODB.Recordset

20      If IsNumeric(txtNo.Text) Then
25        Set rstOutils = New ADODB.Recordset

30        Call rstOutils.Open("SELECT * FROM GRB_outils WHERE no_outils = '" & txtNo.Text & "' ORDER BY no_outils", g_connData, adOpenDynamic, adLockOptimistic)
      
          'si no outils existe
35        If Not rstOutils.EOF Then
            'affiche nom outils et date sortie
40          txtnom.Text = rstOutils!nom_outils
45          txtdepart.Text = CStr(Year(Now)) + "-" + Right$("0" + CStr(Month(Now)), 2) + "-" + Right$("0" + CStr(Day(Now)), 2) + " " + CStr(Time)
50          txtdepartement.Text = rstOutils!departement
        
            'affiche l'employe
55          Set rstEmploye = New ADODB.Recordset
            
60          Call rstEmploye.Open("SELECT * FROM GRB_employés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)

65          txtEmploye.Text = rstEmploye!employe
70          txtEmploye.Tag = rstEmploye!noEmploye
        
            'ferme la table
75          Call rstEmploye.Close
80          Set rstEmploye = Nothing
85        Else
90          Call vider_champs
95        End If
      
          'ferme la table
100       Call rstOutils.Close
105       Set rstOutils = Nothing
110     Else
115       Call vider_champs
120     End If

125     Exit Sub

AfficherErreur:

130     woups "frmOutils_InOut", "txtno_LostFocus", Err, Erl
End Sub

Private Sub vider_champs()

5       On Error GoTo AfficherErreur
              'vide les champs
10      txtcommentaire.Text = vbNullString
15      txtdepart.Text = vbNullString
20      txtdepartement.Text = vbNullString
25      txtEmploye.Text = vbNullString
30      txtNo.Text = vbNullString
35      txtnom.Text = vbNullString
40      txtsortie.Text = vbNullString

45      Exit Sub

AfficherErreur:

50      woups "frmOutils_InOut", "vider_champs", Err, Erl
End Sub
