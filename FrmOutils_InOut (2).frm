VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmOutils_InOut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Magasin"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9315
   Icon            =   "FrmOutils_InOut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5730
   ScaleWidth      =   9315
   Begin VB.CheckBox chknonRetour 
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
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
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

 On Error GoTo Oups

 Screen.MousePointer = vbHourglass
 
 'remplis le lister dependant si retourné ou pas
 Call remplir_lister_outils
 
 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmOutils_InOut", "chknonRetour_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmbEmployé_Click()

 On Error GoTo Oups

 Screen.MousePointer = vbHourglass

 'remplis le lister dependant l'employé selectionné
 Call remplir_lister_outils
 
 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmOutils_InOut", "cmbEmployé_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdAnul_Click()

 On Error GoTo Oups

 'affiche mode liste
 fraOutils.Visible = False
 lstoutils.Visible = True
 cmbemployé.Visible = True
 lblemployé.Visible = True
 chknonRetour.Visible = True

 'remplis le lister
 Call remplir_lister_outils

 Exit Sub

Oups:

 wOups "frmOutils_InOut", "CmdAnul_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdConfig_Click()

 On Error GoTo Oups

 Screen.MousePointer = vbHourglass
 
 'affiche ecran pour ajouter modifier un outils
 Call OuvrirForm(frmoutils, True)
 
 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmOutils_InOut", "cmdConfig_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdEnr_Click()

 On Error GoTo Oups

 'enregistre
 Dim rstOutils As ADODB.Recordset
 Dim rstOutilsInOut As ADODB.Recordset

 Screen.MousePointer = vbHourglass

 If IsNumeric(txtNo.Text) Then
 'ouvre la table
 Set rstOutils = New ADODB.Recordset
 
 Call rstOutils.Open("SELECT * FROM Grboutils WHERE no_outils = '" & txtNo.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 'si existe on peut l'ajouter
 If Not rstOutils.EOF Then
 'ouvre la table avec une recherche sur l'outils étant non-retourné
 Set rstOutilsInOut = New ADODB.Recordset
 
 Call rstOutilsInOut.Open("SELECT GrbOutils_In_out.*, Grbemployés.employe FROM GrbOutils_In_out INNER JOIN Grbemployés ON GrbOutils_In_out.no_employe = Grbemployés.noemploye WHERE no_outils = " & txtNo.Text & " AND retour_date_heure is null", g_connData, adOpenDynamic, adLockOptimistic)
 
 If rstOutilsInOut.EOF Then
 'ajoute
  Call rstOutilsInOut.AddNew
 
  rstOutilsInOut!no_outils = txtNo.Text
  rstOutilsInOut!no_employe = txtEmploye.Tag
  rstOutilsInOut!depart_date_heure = txtdepart.Text
  rstOutilsInOut!commentaire = txtcommentaire.Text
 
  Call rstOutilsInOut.Update
 
 'vide les champs
  Call vider_champs
  Else
 If MsgBox("L'outils n'a pas été retourné par " & rstOutilsInOut.Fields("employe") & ", voulez-vous retourner l'outil pour cet employé?", vbYesNo, CStr(rstOutils!no_outils) + " " + rstOutils!nom_outils) = vbYes Then
 'retourne l'outils pour l'employé
 rstOutilsInOut!commentaire = "(Retourné par: " + CStr(g_sEmploye) + ") " + CStr(rstOutilsInOut!commentaire)
 rstOutilsInOut!retour_date_heure = CStr(Year(Now)) + "-" + Right$("0" + CStr(Month(Now)), 2) + "-" + Right$("0" + CStr(Day(Now)), 2) + " " + CStr(Time)
 
 Call rstOutilsInOut.Update
 
 'ajoute
 Call rstOutilsInOut.AddNew
 
 rstOutilsInOut!no_outils = txtNo.Text
 rstOutilsInOut!no_employe = txtEmploye.Tag
 rstOutilsInOut!depart_date_heure = txtdepart.Text
 rstOutilsInOut!commentaire = txtcommentaire.Text
 
 Call rstOutilsInOut.Update
 
 'vide les champs
 Call vider_champs
 End If
 End If
 
 'quitte la table
 Call rstOutilsInOut.Close
 Set rstOutilsInOut = Nothing
 
 Call txtNo.SetFocus
 Else
 Call MsgBox("Le numéro de l'outil n'existe pas!", , "Erreur")
 End If
1  Else
 Call MsgBox("Le numéro doit être numérique!", , "Erreur")
 End If
 
Screen.MousePointer = vbDefault

Exit Sub

Oups:

wOups "frmOutils_InOut", "CmdEnr_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdFerme_Click()

 On Error GoTo Oups

 'quitte la fenetre
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmOutils_InOut", "CmdFerme_Click", Err, Err.number, Err.Description
End Sub

Public Sub remplir_lister_outils()

 On Error GoTo Oups

 'remplis lister une journée
 Dim rstOutils As ADODB.Recordset
 Dim rstOutilsInOut As ADODB.Recordset
 Dim rstEmploye As ADODB.Recordset
 Dim itmOutils As ListItem
 
 'vide le lister
 Call lstoutils.ListItems.Clear
 lstoutils.Sorted = False

 Set rstOutilsInOut = New ADODB.Recordset
 'affiche tout les employes si = *
 If cmbemployé.Text = "*" Then
 'affiche les non-retournés si est coché, sinon affiche retourné et non
 If chknonRetour.Value = vbChecked Then
 Call rstOutilsInOut.Open("SELECT * FROM Grboutils_in_out WHERE retour_date_heure is null ORDER BY no_enreg", g_connData, adOpenDynamic, adLockOptimistic)
  Else
  Call rstOutilsInOut.Open("SELECT * FROM Grboutils_in_out ORDER BY no_enreg", g_connData, adOpenDynamic, adLockOptimistic)
  End If
  Else
 'affiche pour un employé et
 'affiche les non-retourné si est coché, sinon affiche retourné et non retourné
  If chknonRetour.Value = vbChecked Then
  Call rstOutilsInOut.Open("SELECT * FROM Grboutils_in_out WHERE CStr(no_employe) = CStr('" & cmbemployé.ItemData(cmbemployé.ListIndex) & "') AND retour_date_heure is null ORDER BY no_enreg", g_connData, adOpenDynamic, adLockOptimistic)
  Else
  Call rstOutilsInOut.Open("SELECT * FROM Grboutils_in_out WHERE CStr(no_employe) = CStr('" & cmbemployé.ItemData(cmbemployé.ListIndex) & "') ORDER BY no_enreg", g_connData, adOpenDynamic, adLockOptimistic)
End If
End If
 
Set rstOutils = New ADODB.Recordset
Set rstEmploye = New ADODB.Recordset
 
Do While Not rstOutilsInOut.EOF
 Call rstOutils.Open("SELECT * FROM Grboutils WHERE CStr(no_outils) = CStr('" & rstOutilsInOut!no_outils & "') ORDER BY no_outils", g_connData, adOpenDynamic, adLockOptimistic)
 Call rstEmploye.Open("SELECT * FROM Grbemployés WHERE CStr(noemploye) = CStr('" & rstOutilsInOut!no_employe & "') ORDER BY noemploye", g_connData, adOpenDynamic, adLockOptimistic)
 
 'ajoute dans lister
 Set itmOutils = lstoutils.ListItems.Add
 
 itmOutils.Text = (rstOutils!no_outils)
 itmOutils.Tag = rstOutilsInOut!no_enreg
 
 Call itmOutils.ListSubItems.Add(, , rstOutils!nom_outils)
 Call itmOutils.ListSubItems.Add(, , rstEmploye!employe)
itmOutils.ListSubItems(2).Tag = rstEmploye!loginname
 Call itmOutils.ListSubItems.Add(, , rstOutilsInOut!depart_date_heure)
 
 If IsNull(rstOutilsInOut!retour_date_heure) Then
 Call itmOutils.ListSubItems.Add(, , vbNullString)
 Else
 Call itmOutils.ListSubItems.Add(, , rstOutilsInOut!retour_date_heure)
 End If
 
1  Call itmOutils.ListSubItems.Add(, , rstOutils!departement)
 Call itmOutils.ListSubItems.Add(, , rstOutilsInOut!commentaire)

 Call rstOutilsInOut.MoveNext
 
 'ferme la table
 Call rstOutils.Close
 Call rstEmploye.Close
Loop
 
Call lstoutils.Refresh
 
Set rstOutils = Nothing
Set rstEmploye = Nothing

 'ferme la table
Call rstOutilsInOut.Close
Set rstOutilsInOut = Nothing

Exit Sub

Oups:

wOups "frmOutils_InOut", "remplir_lister_outils", Err, Err.number, Err.Description
End Sub

Public Sub remplir_cmbemploye()

 On Error GoTo Oups

 ''''''''''''''''''''''''''''''''''''''''''''',
 'remplis combo etiquette en mode modification
 ''''''''''''''''''''''''''''''''''''''''''''''
 Dim rstEmploye As ADODB.Recordset
 
 Set rstEmploye = New ADODB.Recordset
 
 Call rstEmploye.Open("SELECT NoEmploye, Employe FROM Grbemployés WHERE Actif = true ORDER BY employe", g_connData, adOpenDynamic, adLockOptimistic)
 
 Call cmbemployé.Clear
 
 Call cmbemployé.AddItem("*")
 'rempli tant il y a des type_étiquette
 
 Do While Not rstEmploye.EOF
 Call cmbemployé.AddItem(rstEmploye!employe)
 cmbemployé.ItemData(cmbemployé.newIndex) = rstEmploye!noEmploye

 Call rstEmploye.MoveNext
 Loop

 'ferme la table
  Call rstEmploye.Close
  Set rstEmploye = Nothing

  Exit Sub

Oups:

  wOups "frmOutils_InOut", "remplir_cmbemploye", Err, Err.number, Err.Description
End Sub

Private Sub cmdreport_Click()

 On Error GoTo Oups

 'affiche rapport filtré par les selection a l'ecran
 Dim rstOutilsInOut As ADODB.Recordset

 Set rstOutilsInOut = New ADODB.Recordset

 'si tout les employé
 If cmbemployé.Text = "*" Then
 'affiche les non-retournés si est coché, sinon affiche retourné et non
 If chknonRetour.Value = vbChecked Then
 Call rstOutilsInOut.Open("SELECT GrbOutils_In_out.*, Grbemployés.employe, GrbOutils.nom_outils, GrbOutils.departement FROM Grbemployés INNER JOIN (GrbOutils_In_out INNER JOIN GrbOutils ON CStr(GrbOutils_In_out.no_outils) = GrbOutils.no_outils) ON Grbemployés.noemploye = GrbOutils_In_out.no_employe WHERE retour_date_heure is null ORDER BY no_enreg", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstOutilsInOut.Open("SELECT GrbOutils_In_out.*, Grbemployés.employe, GrbOutils.nom_outils, GrbOutils.departement FROM Grbemployés INNER JOIN (GrbOutils_In_out INNER JOIN GrbOutils ON CStr(GrbOutils_In_out.no_outils) = GrbOutils.no_outils) ON Grbemployés.noemploye = GrbOutils_In_out.no_employe ORDER BY no_enreg", g_connData, adOpenDynamic, adLockOptimistic)
 End If
 Else
 'affiche pour un employé et
 'affiche les non-retourné si est coché, sinon affiche retourné et non retourné
 If chknonRetour.Value = vbChecked Then
  Call rstOutilsInOut.Open("SELECT GrbOutils_In_out.*, Grbemployés.employe, GrbOutils.nom_outils, GrbOutils.departement FROM Grbemployés INNER JOIN (GrbOutils_In_out INNER JOIN GrbOutils ON CStr(GrbOutils_In_out.no_outils) = GrbOutils.no_outils) ON Grbemployés.noemploye = GrbOutils_In_out.no_employe WHERE CStr(no_employe) = CStr('" & cmbemployé.ItemData(cmbemployé.ListIndex) & "') AND retour_date_heure is null ORDER BY no_enreg", g_connData, adOpenDynamic, adLockOptimistic)
  Else
  Call rstOutilsInOut.Open("SELECT GrbOutils_In_out.*, Grbemployés.employe, GrbOutils.nom_outils, GrbOutils.departement FROM Grbemployés INNER JOIN (GrbOutils_In_out INNER JOIN GrbOutils ON CStr(GrbOutils_In_out.no_outils) = GrbOutils.no_outils) ON Grbemployés.noemploye = GrbOutils_In_out.no_employe WHERE CStr(no_employe) = CStr('" & cmbemployé.ItemData(cmbemployé.ListIndex) & "') ORDER BY no_enreg", g_connData, adOpenDynamic, adLockOptimistic)
  End If
  End If

 ''''''''''''''''''''''''''''''''''''''''''
 'rapport liste d'outil pour un departement
 ''''''''''''''''''''''''''''''''''''''''''
 'set le rapport
  Set DR_Outils_in_out.DataSource = rstOutilsInOut
 
 'contenu label
  If chknonRetour.Value = vbChecked Then
  DR_Outils_in_out.Sections("section2").Controls("lbldepartement").Caption = "Liste des outils non-retourné"
10 Else
1 DR_Outils_in_out.Sections("section2").Controls("lbldepartement").Caption = "Liste des outils empruntés"
End If
 
DR_Outils_in_out.Sections("section3").Controls("lbldate").Caption = CStr(Year(Now)) + "-" + CStr(Month(Now)) + "-" + CStr(Day(Now))
 
DR_Outils_in_out.Orientation = rptOrientLandscape
 
 'affiche rapport
Call DR_Outils_in_out.Show(vbModal)

Exit Sub

Oups:

wOups "frmOutils_InOut", "cmdreport_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdRetour_Click()

 On Error GoTo Oups

 '''''''''''''''''''''''''''''''''''
 'retourne les outils selectionné
 '''''''''''''''''''''''''''''''''''
 Dim rstOutilsInOut As ADODB.Recordset
 Dim iCompteur As Integer
 
 Screen.MousePointer = vbHourglass
 
 Set rstOutilsInOut = New ADODB.Recordset
 
 'tant que pas fin du lister
 For iCompteur = 1 To lstoutils.ListItems.count
 'si l'enreg du lister est selectionné, on modifie dans la bd
 If lstoutils.ListItems(iCompteur).Selected = True Then
 'ouvre la table sur l'enreg selectionné dans le lister
 
 Call rstOutilsInOut.Open("SELECT * FROM Grboutils_in_out WHERE CStr(no_enreg) = CStr('" & lstoutils.ListItems(iCompteur).Tag & "') ORDER BY no_outils", g_connData, adOpenDynamic, adLockOptimistic)
 'met la date d'aujourd'hui si pas null
 
 If IsNull(rstOutilsInOut!retour_date_heure) Then
 'si c'est un autre employé, on ajoute dans commentaire l'employé qui a retourné l'outil
 If g_sUserID <> lstoutils.ListItems(iCompteur).ListSubItems(2).Tag Then
 rstOutilsInOut.Fields("Commentaire") = "(Retourné par: " + g_sEmploye + ")" + rstOutilsInOut.Fields("Commentaire")
  End If
 
  rstOutilsInOut!retour_date_heure = CStr(Year(Now)) + "-" + Right$("0" + CStr(Month(Now)), 2) + "-" + Right$("0" + CStr(Day(Now)), 2) + " " + CStr(Time)
 
  Call rstOutilsInOut.Update
  End If
 
 'ferme la table
  Call rstOutilsInOut.Close
  Set rstOutilsInOut = Nothing
  End If
  Next
 
 'met a jour le lister
10 Call remplir_lister_outils
 
Screen.MousePointer = vbDefault

Exit Sub

Oups:

wOups "frmOutils_InOut", "cmdretour_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdsortie_Click()

 On Error GoTo Oups

 Screen.MousePointer = vbHourglass
 
 'affiche en mode ajout
 fraOutils.Visible = True
 lstoutils.Visible = False
 cmbemployé.Visible = False
 lblemployé.Visible = False
 chknonRetour.Visible = False
 Call txtNo.SetFocus
 
 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmOutils_InOut", "cmdsortie_Click", Err, Err.number, Err.Description
End Sub

Private Sub ActiverBoutonsGroupe()

 On Error GoTo Oups

 cmdConfig.Enabled = g_bModificationOutils
 
 Exit Sub

Oups:

 wOups "frmOutils_InOut", "ActiverBoutonsGroupe", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 Dim rstEmploye As ADODB.Recordset
 Dim iCompteur As Integer
 Dim iNoEmploye As Integer
 Dim bEmpExiste As Boolean
 
 Call Unload(frmChoixInventaire)
 
 'rempli le combo employe
 Call remplir_cmbemploye
 
 'si il y a des employes ,selectionne par defaut
 If cmbemployé.ListCount > 0 Then
 Set rstEmploye = New ADODB.Recordset

 Call rstEmploye.Open("SELECT * FROM Grbemployés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 'select l'employé logué sinon, affiche tout
 If rstEmploye.EOF Then
  bEmpExiste = False
  Else
  bEmpExiste = True
 
  iNoEmploye = rstEmploye.Fields("noEmploye")
  End If
 
  Call rstEmploye.Close
  Set rstEmploye = Nothing
 
  If bEmpExiste = False Then
 cmbemployé.ListIndex = 0
1 Else
 'on trouve l'index dans le combo qui contient le noemploye, pour le selectionné
 For iCompteur = 0 To cmbemployé.ListCount - 1
 'si trouvé noemploye, on selectione cet index
 If cmbemployé.ItemData(iCompteur) = iNoEmploye Then
 cmbemployé.ListIndex = iCompteur
 
 Exit For
 End If
 Next
 End If
End If
 
 'met enabled le bouton config dependant le user
Call ActiverBoutonsGroupe

Exit Sub

Oups:

1  wOups "frmOutils_InOut", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub txtno_LostFocus()

 On Error GoTo Oups


 'affiche automatique le nom de l'outils et l'employé et la date de sortie
 Dim rstOutils As ADODB.Recordset
 Dim rstEmploye As ADODB.Recordset

 If IsNumeric(txtNo.Text) Then
 Set rstOutils = New ADODB.Recordset

 Call rstOutils.Open("SELECT * FROM Grboutils WHERE no_outils = '" & txtNo.Text & "' ORDER BY no_outils", g_connData, adOpenDynamic, adLockOptimistic)
 
 'si no outils existe
 If Not rstOutils.EOF Then
 'affiche nom outils et date sortie
 txtnom.Text = rstOutils!nom_outils
 txtdepart.Text = CStr(Year(Now)) + "-" + Right$("0" + CStr(Month(Now)), 2) + "-" + Right$("0" + CStr(Day(Now)), 2) + " " + CStr(Time)
 txtdepartement.Text = rstOutils!departement
 
 'affiche l'employe
 Set rstEmploye = New ADODB.Recordset
 
  Call rstEmploye.Open("SELECT * FROM Grbemployés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)

  txtEmploye.Text = rstEmploye!employe
  txtEmploye.Tag = rstEmploye!noEmploye
 
 'ferme la table
  Call rstEmploye.Close
  Set rstEmploye = Nothing
  Else
  Call vider_champs
  End If
 
 'ferme la table
Call rstOutils.Close
1 Set rstOutils = Nothing
Else
 Call vider_champs
End If

Exit Sub

Oups:

wOups "frmOutils_InOut", "txtno_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub vider_champs()

 On Error GoTo Oups
 'vide les champs
 txtcommentaire.Text = vbNullString
 txtdepart.Text = vbNullString
 txtdepartement.Text = vbNullString
 txtEmploye.Text = vbNullString
 txtNo.Text = vbNullString
 txtnom.Text = vbNullString
 txtsortie.Text = vbNullString

 Exit Sub

Oups:

 wOups "frmOutils_InOut", "vider_champs", Err, Err.number, Err.Description
End Sub
