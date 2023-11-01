VERSION 5.00
Begin VB.Form FrmDispatch 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automation GRB"
   ClientHeight    =   4500
   ClientLeft      =   -150
   ClientTop       =   -240
   ClientWidth     =   7575
   Icon            =   "FrmDispatch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   7575
   Begin VB.CommandButton CmdChangerDB 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Changer de base de données"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2640
      TabIndex        =   17
      Top             =   3840
      Width           =   2325
   End
   Begin VB.CommandButton cmdDistList 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Listes de distribution"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2640
      TabIndex        =   15
      Top             =   3120
      Width           =   2325
   End
   Begin VB.CommandButton cmdFormulaire 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Rapports"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   2295
   End
   Begin VB.CommandButton cmdEmploye 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Employés"
      Enabled         =   0   'False
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
      TabIndex        =   11
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton cmdContact 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Co&ntacts"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   2295
   End
   Begin VB.CommandButton cmdFournisseur 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Fournisseurs"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Timer tmrAlarme 
      Interval        =   5000
      Left            =   2280
      Top             =   360
   End
   Begin VB.CommandButton cmdInventaire 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Inventaire"
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
      Left            =   5160
      TabIndex        =   4
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton cmdPunch 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Punch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2640
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton cmdConfiguration 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Confi&guration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   5160
      TabIndex        =   12
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton cmdCedule 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cé&dule"
      Enabled         =   0   'False
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
      Left            =   2640
      TabIndex        =   6
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton cmdVendeur 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Contacts pour &vendeurs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2640
      TabIndex        =   9
      Top             =   2400
      Width           =   2325
   End
   Begin VB.CommandButton cmdquitter 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Quitter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   5160
      TabIndex        =   14
      Top             =   3840
      Width           =   2295
   End
   Begin VB.CommandButton cmdClient 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Clients"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton cmdCatalogue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "C&atalogue"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   5160
      TabIndex        =   7
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton cmdProjSoum 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Projets / &Soumissions"
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
      Left            =   5160
      TabIndex        =   10
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label lbldb 
      BackColor       =   &H00000000&
      Caption         =   "Base de donné:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   16
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label lblDerniereVersion 
      BackColor       =   &H00000000&
      Caption         =   "Dernière Version : "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00000000&
      Caption         =   "Version "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6240
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "FrmDispatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiverBoutonsGroupe()

 On Error GoTo Oups

 'Dépendant le groupe de sécurité contenu dans la table GrbGroupes
 'met enabled les boutons
 cmdClient.Enabled = g_bAffichageClients
 cmdFournisseur.Enabled = g_bAffichageFournisseurs
 cmdContact.Enabled = g_bAffichageContacts
 'GLLcmdVendeur.Enabled = g_bAffichageContactsVendeurs
 cmdFormulaire.Enabled = g_bAffichageRapports
 cmdEmploye.Enabled = g_bAffichageEmployes
 cmdCedule.Enabled = g_bAffichageCedule
 cmdConfiguration.Enabled = g_bAffichageConfiguration
 cmdDistList.Enabled = g_bModificationListeDistribution
 
 If g_bAffichagePunch = True Or g_bModificationFeuillesTemps = True Or g_bModificationFacturation = True Then
  cmdPunch.Enabled = True
  Else
  cmdPunch.Enabled = False
  End If

  If g_bAffichageSoumissionsMec = True Or g_bAffichageSoumissionsElec = True Or g_bAffichageProjetsMec = True Or g_bAffichageProjetsElec = True Then
  cmdProjSoum.Enabled = True
  Else
  cmdProjSoum.Enabled = False
10 End If
 
If g_bAffichageCatalogueMec = True Or g_bAffichageCatalogueElec = True Then
 cmdCatalogue.Enabled = True
Else
 cmdCatalogue.Enabled = False
End If
 
If g_bAffichageInventaireMec = True Or g_bAffichageInventaireElec = True Or g_bAffichageOutils = True Then
 cmdInventaire.Enabled = True
Else
 cmdInventaire.Enabled = False
End If

Exit Sub

Oups:

1  wOups "frmDispatch", "ActiverBoutonsGroupe", Err, Err.number, Err.Description
End Sub

Private Sub cmdCedule_Click()

 On Error GoTo Oups

 'Cédule
 Screen.MousePointer = vbHourglass

 Call OuvrirForm(frmCédule, False)

 Exit Sub

Oups:

 wOups "frmDispatch", "cmdCedule_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdChangerDB_Click()
Call FermerConnection 'Fermer la connection pour pouvoir ouvrir l'autre
'Si la Base de donné est présentement la base de donné actuel on ouvre l'ancienne
If BdMaintenant = True Then
 Call OuvrirOldConnection
 lbldb.Caption = "Base de donné:Ancienne"
 BdMaintenant = False
Else
'Si la Base de donné est présentement l'ancienne base de donné on ouvre la base de donné actuel
 Call OuvrirConnection
 BdMaintenant = True
 lbldb.Caption = "Base de donné:Actuel"
End If

End Sub

Private Sub cmdConfiguration_Click()

 On Error GoTo Oups

 'Configuration
 Screen.MousePointer = vbHourglass
 
 Call OuvrirForm(FrmPara, True)

 Exit Sub

Oups:

 wOups "frmDispatch", "cmdConfiguration_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdDistList_Click()

 On Error GoTo Oups

 Screen.MousePointer = vbHourglass

 Call OuvrirForm(frmDistList, False)

 Exit Sub

Oups:

 wOups "frmDispatch", "cmdDistList_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdEmploye_Click()

 On Error GoTo Oups

 'Employés
 Screen.MousePointer = vbHourglass

 Call OuvrirForm(frmemploye, False)

 Exit Sub

Oups:

 wOups "frmDispatch", "cmdEmploye_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdInventaire_Click()
 'Magasin
 On Error GoTo Oups

 Dim lSize As Long
 Dim lLCID As Long
 Dim sBuffer As String
 
 'Vérifie si bons paramètres régionaux
 lLCID = GetUserDefaultLCID
 
 lSize = GetLocaleInfo(lLCID, LOCALE_SDECIMAL, StrPtr(sBuffer), lSize)
 
 sBuffer = Space$(lSize)
 
 lSize = GetLocaleInfo(lLCID, LOCALE_SDECIMAL, sBuffer, lSize)
 
 sBuffer = Trim$(Replace(sBuffer, Chr(0), ""))
 
 If sBuffer = "," Then
 Screen.MousePointer = vbHourglass

  Call OuvrirForm(frmChoixInventaire, True)
  Else
  Call MsgBox("Vos paramètres régionaux sont incorrects!" & vbNewLine & _
 "Vous devez avoir la virgule (,) comme symbole de décimal!" & vbNewLine & _
 "Des erreurs vont se produire dans ce formulaire car il contient des montants d'argent!", vbOKOnly, "Erreur")
  End If

  Exit Sub

Oups:

  wOups "frmDispatch", "cmdInventaire_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdProjSoum_Click()

 On Error GoTo Oups

 Dim lSize As Long
 Dim lLCID As Long
 Dim sBuffer As String
 
 'Vérifie si bons paramètres régionaux
 lLCID = GetUserDefaultLCID
 
 lSize = GetLocaleInfo(lLCID, LOCALE_SDECIMAL, StrPtr(sBuffer), lSize)
 
 sBuffer = Space$(lSize)
 
 lSize = GetLocaleInfo(lLCID, LOCALE_SDECIMAL, sBuffer, lSize)
 
 sBuffer = Trim$(Replace(sBuffer, Chr(0), ""))
 
 If sBuffer = "," Then
  Screen.MousePointer = vbHourglass

  Call OuvrirForm(frmChoixProjSoum, True)
  Else
  Call MsgBox("Vos paramètres régionaux sont incorrects!" & vbNewLine & _
 "Vous devez avoir la virgule (,) comme symbole de décimal!" & vbNewLine & _
 "Des erreurs vont se produire dans ce formulaire car il contient des montants d'argent!", vbOKOnly, "Erreur")
  End If

  Exit Sub

Oups:

  wOups "frmDispatch", "cmdProjSoum_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdPunch_Click()

 On Error GoTo Oups

 'Punch
 Screen.MousePointer = vbHourglass
 
 Call OuvrirForm(frmChoixPunch, True)
 
 Exit Sub

Oups:

 wOups "frmDispatch", "cmdPunch_Click", Err, Err.number, Err.Description
End Sub

Private Sub ActiverBoutons(ByVal bEnabled As Boolean)

 On Error GoTo Oups

 cmdCatalogue.Enabled = bEnabled
 cmdCedule.Enabled = bEnabled
 cmdClient.Enabled = bEnabled
 cmdConfiguration.Enabled = bEnabled
 cmdContact.Enabled = bEnabled
 cmdEmploye.Enabled = bEnabled
 cmdFormulaire.Enabled = bEnabled
 cmdFournisseur.Enabled = bEnabled
 cmdInventaire.Enabled = bEnabled
 cmdProjSoum.Enabled = bEnabled
  cmdPunch.Enabled = bEnabled
  cmdVendeur.Enabled = bEnabled

  Exit Sub

Oups:

  wOups "frmDispatch", "ActiverBoutons", Err, Err.number, Err.Description
End Sub

Private Sub cmdVendeur_Click()

 On Error GoTo Oups
 
 'Contacts pour vendeur
 Screen.MousePointer = vbHourglass
 
 Call OuvrirForm(frmvendeur, False)
 
 Exit Sub

Oups:

 wOups "frmDispatch", "cmdVendeur_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdClient_Click()

 On Error GoTo Oups

 'Clients
 Screen.MousePointer = vbHourglass
 
 Call OuvrirForm(FrmClient, False)
 
 Exit Sub

Oups:

 wOups "frmDispatch", "cmdClient_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdFournisseur_Click()

 On Error GoTo Oups

 'Founisseurs
 Screen.MousePointer = vbHourglass
 
 Call OuvrirForm(FrmFRS, False)
 
 Exit Sub

Oups:

 wOups "frmDispatch", "cmdFournisseur_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdContact_Click()

 On Error GoTo Oups

 'Contacts
 Screen.MousePointer = vbHourglass
 
 Call OuvrirForm(FrmContact, False)

 Exit Sub

Oups:

 wOups "frmDispatch", "cmdContact_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdCatalogue_Click()
 'Magasin
 On Error GoTo Oups

 Dim lSize As Long
 Dim lLCID As Long
 Dim sBuffer As String
 
 'Vérifie si bons paramètres régionaux
 lLCID = GetUserDefaultLCID
 
 lSize = GetLocaleInfo(lLCID, LOCALE_SDECIMAL, StrPtr(sBuffer), lSize)
 
 sBuffer = Space$(lSize)
 
 lSize = GetLocaleInfo(lLCID, LOCALE_SDECIMAL, sBuffer, lSize)
 
 sBuffer = Trim$(Replace(sBuffer, Chr(0), ""))
 
 If sBuffer = "," Then
 Screen.MousePointer = vbHourglass

  Call OuvrirForm(frmChoixCatalogue, True)
  Else
  Call MsgBox("Vos paramètres régionaux sont incorrects!" & vbNewLine & _
 "Vous devez avoir la virgule (,) comme symbole de décimal!" & vbNewLine & _
 "Des erreurs vont se produire dans ce formulaire car il contient des montants d'argent!", vbOKOnly, "Erreur")
  End If
 
  Screen.MousePointer = vbDefault

  Exit Sub

Oups:

  wOups "frmDispatch", "cmdCatalogue_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdQuitter_Click()

 On Error GoTo Oups

 'Quitte l'application
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmDispatch", "cmdQuitter_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdFormulaire_Click()

 On Error GoTo Oups

 'Rapport
 Screen.MousePointer = vbHourglass

 Call OuvrirForm(frmreport, False)

 Exit Sub

Oups:

 wOups "frmDispatch", "cmdFormulaire_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups
 'Vérifie si l'utilisateur peux voir le bouton changement de base de donné
 If g_iNoGroupe = 2 Or g_iNoGroupe = 24 Then 'Or g_iNoGroupe = 22
 lbldb.Caption = "Base de données:Actuel" 'ajoute l'information de quel base de donné on active GLL
 lbldb.Visible = True
 g_admin = True
 CmdChangerDB.Visible = True
 Else
 lbldb.Visible = False
 CmdChangerDB.Visible = False
 g_admin = False
 End If
 Dim sVersion As String
 Dim rstConfig As ADODB.Recordset
 
 Call ActiverBoutonsGroupe
 
 'Caption = Programme + Nom de l'employé
 Me.Caption = "Solution GRB inc. (" & g_sEmploye & ")"

 sVersion = App.Major & "." & Right$("0" & App.Minor, 2) & "." & Right$("0" & App.Revision, 4)
 
 lblVersion.Caption = "Version " & sVersion

 Set rstConfig = New ADODB.Recordset

 Call rstConfig.Open("SELECT DerniereVersion FROM GrbConfig", g_connData, adOpenDynamic, adLockOptimistic)

 If Not IsNull(rstConfig.Fields("DerniereVersion")) Then
 If rstConfig.Fields("DerniereVersion") <> "" Then
  lblDerniereVersion.Caption = "Dernière Version : " & rstConfig.Fields("DerniereVersion")
  Else
  lblDerniereVersion.Caption = ""
  End If
  Else
  lblDerniereVersion.Caption = ""
  End If

  Call rstConfig.Close
10 Set rstConfig = Nothing

If Trim$(Replace(lblDerniereVersion.Caption, "Dernière Version : ", "")) = Trim$(Replace(lblVersion.Caption, "Version", "")) Then
 lblVersion.ForeColor = vbGreen
Else
 lblVersion.ForeColor = vbRed
End If

Screen.MousePointer = vbDefault

Exit Sub

Oups:

wOups "frmDispatch", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)

 On Error GoTo Oups

 Dim iCompteur As Integer

 'Ferme les forms invisibles
 If Forms.count > 1 Then
 iCompteur = 0

 Do While iCompteur <= Forms.count - 1
 If Forms(iCompteur).Visible = False Then
 Call Unload(Forms(iCompteur))
 Else
 iCompteur = iCompteur + 1
 End If
 Loop
  End If

  If Forms.count > 1 Then
  If MsgBox("Il y a des formulaires ouverts, êtes-vous certains de vouloir quitter?", vbYesNo) = vbYes Then
  If FermerTousLesForms = True Then
  Call FermerConnection

  End
  Else
  Call MsgBox("Impossible de fermer, un formulaire est encore en modification!", vbOKOnly, "Erreur")

 Cancel = 1
End If
 Else
 Cancel = 1
 End If
Else
 Call Unload(Me)

 End
End If

Exit Sub

Oups:

wOups "frmDispatch", "Form_Unload", Err, Err.number, Err.Description
End Sub

Private Function FermerTousLesForms() As Boolean

 On Error GoTo Oups

 Dim objForm As Form
 Dim bFermer As Boolean

 bFermer = True

 For Each objForm In Forms
 If objForm.Name <> Me.Name Then
 If UCase(objForm.Name) = "FRMPROJSOUMELEC" Or UCase(objForm.Name) = "FRMPROJSOUMMEC" Then
 bFermer = objForm.PeutFermer

 Exit For
 End If
 End If
  Next

  If bFermer = True Then
  For Each objForm In Forms
  If objForm.Name <> Me.Name Then
  Call Unload(objForm)
  End If
  Next
  End If

10 FermerTousLesForms = bFermer

Exit Function

Oups:

wOups "frmDispatch", "FermerTousLesForms", Err, Err.number, Err.Description
End Function

Private Sub Label1_Click()

End Sub

Private Sub tmrAlarme_Timer()

 On Error GoTo Oups

 Dim rstAlarme As ADODB.Recordset
 Dim rstEmploye As ADODB.Recordset
 Dim rstConfig As ADODB.Recordset
 Dim bAfficher As Boolean
 Dim iNoEmploye As Integer

 Set rstConfig = New ADODB.Recordset

 Call rstConfig.Open("SELECT DerniereVersion FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

 If Not IsNull(rstConfig.Fields("DerniereVersion")) Then
 If rstConfig.Fields("DerniereVersion") <> "" Then
 lblDerniereVersion.Caption = "Dernière Version : " & rstConfig.Fields("DerniereVersion")
  Else
  lblDerniereVersion.Caption = ""
  End If
  Else
  lblDerniereVersion.Caption = ""
  End If

  Call rstConfig.Close
  Set rstConfig = Nothing

10 If Trim$(Replace(lblDerniereVersion.Caption, "Dernière Version : ", "")) = Trim$(Replace(lblVersion.Caption, "Version", "")) Then
1 lblVersion.ForeColor = vbGreen
Else
 lblVersion.ForeColor = vbRed
End If

Set rstEmploye = New ADODB.Recordset

Call rstEmploye.Open("SELECT * FROM GrbEmployés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

iNoEmploye = rstEmploye.Fields("NoEmploye")

Call rstEmploye.Close
Set rstEmploye = Nothing
 
 'Ouverture de la table
Set rstAlarme = New ADODB.Recordset
 
Call rstAlarme.Open("SELECT * FROM GrbAlarmes WHERE NoEmploye = " & iNoEmploye, g_connData, adOpenForwardOnly, adLockReadOnly)
 
 'Tant qu'il y a des enregistrements
1  Do While Not rstAlarme.EOF
 bAfficher = False

 If rstAlarme.Fields("Date") < ConvertDate(Date) Then
 bAfficher = True
 Else
 If rstAlarme.Fields("Date") = ConvertDate(Date) Then
 If CDate(rstAlarme.Fields("Heure")) <= Time Then
1  bAfficher = True
 End If
 End If
 End If

 If bAfficher = True Then
 Call frmAlarme.Afficher(rstAlarme.Fields("IDAlarme"))
 End If

 Call rstAlarme.MoveNext
Loop

Call rstAlarme.Close
Set rstAlarme = Nothing

Exit Sub

Oups:

wOups "frmDispatch", "tmrAlarme_Timer", Err, Err.number, Err.Description
End Sub
