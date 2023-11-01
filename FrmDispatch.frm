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
   Picture         =   "FrmDispatch.frx":0E42
   ScaleHeight     =   4500
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
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

5       On Error GoTo AfficherErreur

        'Dépendant le groupe de sécurité contenu dans la table GRB_Groupes
        'met enabled les boutons
10      cmdClient.Enabled = g_bAffichageClients
15      cmdFournisseur.Enabled = g_bAffichageFournisseurs
20      cmdContact.Enabled = g_bAffichageContacts
25      'GLLcmdVendeur.Enabled = g_bAffichageContactsVendeurs
30      cmdFormulaire.Enabled = g_bAffichageRapports
35      cmdEmploye.Enabled = g_bAffichageEmployes
40      cmdCedule.Enabled = g_bAffichageCedule
45      cmdConfiguration.Enabled = g_bAffichageConfiguration
50      cmdDistList.Enabled = g_bModificationListeDistribution
        
55      If g_bAffichagePunch = True Or g_bModificationFeuillesTemps = True Or g_bModificationFacturation = True Then
60        cmdPunch.Enabled = True
65      Else
70        cmdPunch.Enabled = False
75      End If

80      If g_bAffichageSoumissionsMec = True Or g_bAffichageSoumissionsElec = True Or g_bAffichageProjetsMec = True Or g_bAffichageProjetsElec = True Then
85        cmdProjSoum.Enabled = True
90      Else
95        cmdProjSoum.Enabled = False
100     End If
    
105     If g_bAffichageCatalogueMec = True Or g_bAffichageCatalogueElec = True Then
110       cmdCatalogue.Enabled = True
115     Else
120       cmdCatalogue.Enabled = False
125     End If
    
130     If g_bAffichageInventaireMec = True Or g_bAffichageInventaireElec = True Or g_bAffichageOutils = True Then
135       cmdInventaire.Enabled = True
140     Else
145       cmdInventaire.Enabled = False
150     End If

155     Exit Sub

AfficherErreur:

160     woups "frmDispatch", "ActiverBoutonsGroupe", Err, Erl
End Sub

Private Sub cmdCedule_Click()

5       On Error GoTo AfficherErreur

        'Cédule
10      Screen.MousePointer = vbHourglass

15      Call OuvrirForm(frmCédule, False)

20      Exit Sub

AfficherErreur:

25      woups "frmDispatch", "cmdCedule_Click", Err, Erl
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

5       On Error GoTo AfficherErreur

        'Configuration
10      Screen.MousePointer = vbHourglass
  
15      Call OuvrirForm(FrmPara, True)

20      Exit Sub

AfficherErreur:

25      woups "frmDispatch", "cmdConfiguration_Click", Err, Erl
End Sub

Private Sub cmdDistList_Click()

5       On Error GoTo AfficherErreur

10      Screen.MousePointer = vbHourglass

15      Call OuvrirForm(frmDistList, False)

20      Exit Sub

AfficherErreur:

25      woups "frmDispatch", "cmdDistList_Click", Err, Erl
End Sub

Private Sub cmdEmploye_Click()

5       On Error GoTo AfficherErreur

        'Employés
10      Screen.MousePointer = vbHourglass

15      Call OuvrirForm(frmemploye, False)

20      Exit Sub

AfficherErreur:

25      woups "frmDispatch", "cmdEmploye_Click", Err, Erl
End Sub

Private Sub cmdInventaire_Click()
        'Magasin
5       On Error GoTo AfficherErreur

10      Dim lSize   As Long
15      Dim lLCID   As Long
20      Dim sBuffer As String
  
        'Vérifie si bons paramètres régionaux
25      lLCID = GetUserDefaultLCID
  
30      lSize = GetLocaleInfo(lLCID, LOCALE_SDECIMAL, StrPtr(sBuffer), lSize)
  
35      sBuffer = Space$(lSize)
  
40      lSize = GetLocaleInfo(lLCID, LOCALE_SDECIMAL, sBuffer, lSize)
  
45      sBuffer = Trim$(Replace(sBuffer, Chr(0), ""))
  
50      If sBuffer = "," Then
55        Screen.MousePointer = vbHourglass

60        Call OuvrirForm(frmChoixInventaire, True)
65      Else
70        Call MsgBox("Vos paramètres régionaux sont incorrects!" & vbNewLine & _
                      "Vous devez avoir la virgule (,) comme symbole de décimal!" & vbNewLine & _
                      "Des erreurs vont se produire dans ce formulaire car il contient des montants d'argent!", vbOKOnly, "Erreur")
75      End If

80      Exit Sub

AfficherErreur:

85      woups "frmDispatch", "cmdInventaire_Click", Err, Erl
End Sub

Private Sub cmdProjSoum_Click()

5       On Error GoTo AfficherErreur

10      Dim lSize   As Long
15      Dim lLCID   As Long
20      Dim sBuffer As String
  
25      'Vérifie si bons paramètres régionaux
30      lLCID = GetUserDefaultLCID
  
35      lSize = GetLocaleInfo(lLCID, LOCALE_SDECIMAL, StrPtr(sBuffer), lSize)
  
40      sBuffer = Space$(lSize)
  
45      lSize = GetLocaleInfo(lLCID, LOCALE_SDECIMAL, sBuffer, lSize)
  
50      sBuffer = Trim$(Replace(sBuffer, Chr(0), ""))
  
55      If sBuffer = "," Then
60        Screen.MousePointer = vbHourglass

65        Call OuvrirForm(frmChoixProjSoum, True)
70      Else
75        Call MsgBox("Vos paramètres régionaux sont incorrects!" & vbNewLine & _
                      "Vous devez avoir la virgule (,) comme symbole de décimal!" & vbNewLine & _
                      "Des erreurs vont se produire dans ce formulaire car il contient des montants d'argent!", vbOKOnly, "Erreur")
80      End If

85      Exit Sub

AfficherErreur:

90      woups "frmDispatch", "cmdProjSoum_Click", Err, Erl
End Sub

Private Sub cmdPunch_Click()

5       On Error GoTo AfficherErreur

        'Punch
10      Screen.MousePointer = vbHourglass
  
15      Call OuvrirForm(frmChoixPunch, True)
  
20      Exit Sub

AfficherErreur:

25      woups "frmDispatch", "cmdPunch_Click", Err, Erl
End Sub

Private Sub ActiverBoutons(ByVal bEnabled As Boolean)

5       On Error GoTo AfficherErreur

10      cmdCatalogue.Enabled = bEnabled
15      cmdCedule.Enabled = bEnabled
20      cmdClient.Enabled = bEnabled
25      cmdConfiguration.Enabled = bEnabled
30      cmdContact.Enabled = bEnabled
35      cmdEmploye.Enabled = bEnabled
40      cmdFormulaire.Enabled = bEnabled
45      cmdFournisseur.Enabled = bEnabled
50      cmdInventaire.Enabled = bEnabled
55      cmdProjSoum.Enabled = bEnabled
60      cmdPunch.Enabled = bEnabled
65      cmdVendeur.Enabled = bEnabled

70      Exit Sub

AfficherErreur:

75      woups "frmDispatch", "ActiverBoutons", Err, Erl
End Sub

Private Sub cmdVendeur_Click()

5       On Error GoTo AfficherErreur
              
        'Contacts pour vendeur
10      Screen.MousePointer = vbHourglass
  
15      Call OuvrirForm(frmvendeur, False)
  
20      Exit Sub

AfficherErreur:

25      woups "frmDispatch", "cmdVendeur_Click", Err, Erl
End Sub

Private Sub cmdClient_Click()

5       On Error GoTo AfficherErreur

        'Clients
10      Screen.MousePointer = vbHourglass
    
15      Call OuvrirForm(FrmClient, False)
  
20      Exit Sub

AfficherErreur:

25      woups "frmDispatch", "cmdClient_Click", Err, Erl
End Sub

Private Sub cmdFournisseur_Click()

5       On Error GoTo AfficherErreur

        'Founisseurs
10      Screen.MousePointer = vbHourglass
    
15      Call OuvrirForm(FrmFRS, False)
  
20      Exit Sub

AfficherErreur:

25      woups "frmDispatch", "cmdFournisseur_Click", Err, Erl
End Sub

Private Sub cmdContact_Click()

5       On Error GoTo AfficherErreur

        'Contacts
10      Screen.MousePointer = vbHourglass
    
15      Call OuvrirForm(FrmContact, False)

20      Exit Sub

AfficherErreur:

25      woups "frmDispatch", "cmdContact_Click", Err, Erl
End Sub

Private Sub cmdCatalogue_Click()
        'Magasin
5       On Error GoTo AfficherErreur

10      Dim lSize   As Long
15      Dim lLCID   As Long
20      Dim sBuffer As String
  
        'Vérifie si bons paramètres régionaux
25      lLCID = GetUserDefaultLCID
  
30      lSize = GetLocaleInfo(lLCID, LOCALE_SDECIMAL, StrPtr(sBuffer), lSize)
  
35      sBuffer = Space$(lSize)
  
40      lSize = GetLocaleInfo(lLCID, LOCALE_SDECIMAL, sBuffer, lSize)
  
45      sBuffer = Trim$(Replace(sBuffer, Chr(0), ""))
  
50      If sBuffer = "," Then
55        Screen.MousePointer = vbHourglass

60        Call OuvrirForm(frmChoixCatalogue, True)
65      Else
70        Call MsgBox("Vos paramètres régionaux sont incorrects!" & vbNewLine & _
                      "Vous devez avoir la virgule (,) comme symbole de décimal!" & vbNewLine & _
                      "Des erreurs vont se produire dans ce formulaire car il contient des montants d'argent!", vbOKOnly, "Erreur")
75      End If
 
80      Screen.MousePointer = vbDefault

85      Exit Sub

AfficherErreur:

90      woups "frmDispatch", "cmdCatalogue_Click", Err, Erl
End Sub

Private Sub cmdQuitter_Click()

5       On Error GoTo AfficherErreur

        'Quitte l'application
10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmDispatch", "cmdQuitter_Click", Err, Erl
End Sub

Private Sub cmdFormulaire_Click()

5       On Error GoTo AfficherErreur

        'Rapport
10      Screen.MousePointer = vbHourglass

15      Call OuvrirForm(frmreport, False)

20      Exit Sub

AfficherErreur:

25      woups "frmDispatch", "cmdFormulaire_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur
        'Vérifie si l'utilisateur peux voir le bouton changement de base de donné
        If g_iNoGroupe = 2 Or g_iNoGroupe = 24 Then    'Or g_iNoGroupe = 22
            lbldb.Caption = "Base de données:Actuel" 'ajoute l'information de quel base de donné on active GLL
            lbldb.Visible = True
            g_admin = True
            CmdChangerDB.Visible = True
         Else
           lbldb.Visible = False
           CmdChangerDB.Visible = False
           g_admin = False
        End If
10      Dim sVersion  As String
15      Dim rstConfig As ADODB.Recordset
  
20      Call ActiverBoutonsGroupe
    
        'Caption = Programme + Nom de l'employé
25      Me.Caption = "Solution GRB inc. (" & g_sEmploye & ")"

30      sVersion = App.Major & "." & Right$("0" & App.Minor, 2) & "." & Right$("0" & App.Revision, 4)
  
35      lblVersion.Caption = "Version " & sVersion

40      Set rstConfig = New ADODB.Recordset

45      Call rstConfig.Open("SELECT DerniereVersion FROM GRB_Config", g_connData, adOpenDynamic, adLockOptimistic)

50      If Not IsNull(rstConfig.Fields("DerniereVersion")) Then
55        If rstConfig.Fields("DerniereVersion") <> "" Then
60          lblDerniereVersion.Caption = "Dernière Version : " & rstConfig.Fields("DerniereVersion")
65        Else
70          lblDerniereVersion.Caption = ""
75        End If
80      Else
85        lblDerniereVersion.Caption = ""
90      End If

95      Call rstConfig.Close
100     Set rstConfig = Nothing

105     If Trim$(Replace(lblDerniereVersion.Caption, "Dernière Version : ", "")) = Trim$(Replace(lblVersion.Caption, "Version", "")) Then
110       lblVersion.ForeColor = vbGreen
115     Else
120       lblVersion.ForeColor = vbRed
125     End If

130     Screen.MousePointer = vbDefault

135     Exit Sub

AfficherErreur:

140     woups "frmDispatch", "Form_Load", Err, Erl
End Sub

Private Sub Form_Unload(Cancel As Integer)

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer

        'Ferme les forms invisibles
15      If Forms.count > 1 Then
20        iCompteur = 0

25        Do While iCompteur <= Forms.count - 1
30          If Forms(iCompteur).Visible = False Then
35            Call Unload(Forms(iCompteur))
40          Else
45            iCompteur = iCompteur + 1
50          End If
55        Loop
60      End If

65      If Forms.count > 1 Then
70        If MsgBox("Il y a des formulaires ouverts, êtes-vous certains de vouloir quitter?", vbYesNo) = vbYes Then
75          If FermerTousLesForms = True Then
80            Call FermerConnection

85            End
90          Else
95            Call MsgBox("Impossible de fermer, un formulaire est encore en modification!", vbOKOnly, "Erreur")

100           Cancel = 1
105         End If
110       Else
115         Cancel = 1
120       End If
125     Else
130       Call Unload(Me)

135       End
140     End If

145     Exit Sub

AfficherErreur:

150     woups "frmDispatch", "Form_Unload", Err, Erl
End Sub

Private Function FermerTousLesForms() As Boolean

5       On Error GoTo AfficherErreur

10      Dim objForm As Form
15      Dim bFermer As Boolean

20      bFermer = True

25      For Each objForm In Forms
30        If objForm.Name <> Me.Name Then
35          If UCase(objForm.Name) = "FRMPROJSOUMELEC" Or UCase(objForm.Name) = "FRMPROJSOUMMEC" Then
40            bFermer = objForm.PeutFermer

45            Exit For
50          End If
55        End If
60      Next

65      If bFermer = True Then
70        For Each objForm In Forms
75          If objForm.Name <> Me.Name Then
80            Call Unload(objForm)
85          End If
90        Next
95      End If

100     FermerTousLesForms = bFermer

105     Exit Function

AfficherErreur:

110     woups "frmDispatch", "FermerTousLesForms", Err, Erl
End Function

Private Sub Label1_Click()

End Sub

Private Sub tmrAlarme_Timer()

5       On Error GoTo AfficherErreur

10      Dim rstAlarme  As ADODB.Recordset
15      Dim rstEmploye As ADODB.Recordset
20      Dim rstConfig  As ADODB.Recordset
25      Dim bAfficher  As Boolean
30      Dim iNoEmploye As Integer

35      Set rstConfig = New ADODB.Recordset

40      Call rstConfig.Open("SELECT DerniereVersion FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

45      If Not IsNull(rstConfig.Fields("DerniereVersion")) Then
50        If rstConfig.Fields("DerniereVersion") <> "" Then
55          lblDerniereVersion.Caption = "Dernière Version : " & rstConfig.Fields("DerniereVersion")
60        Else
65          lblDerniereVersion.Caption = ""
70        End If
75      Else
80        lblDerniereVersion.Caption = ""
85      End If

90      Call rstConfig.Close
95      Set rstConfig = Nothing

100     If Trim$(Replace(lblDerniereVersion.Caption, "Dernière Version : ", "")) = Trim$(Replace(lblVersion.Caption, "Version", "")) Then
105       lblVersion.ForeColor = vbGreen
110     Else
115       lblVersion.ForeColor = vbRed
120     End If

125     Set rstEmploye = New ADODB.Recordset

130     Call rstEmploye.Open("SELECT * FROM GRB_Employés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

135     iNoEmploye = rstEmploye.Fields("NoEmploye")

140     Call rstEmploye.Close
145     Set rstEmploye = Nothing
        
        'Ouverture de la table
150     Set rstAlarme = New ADODB.Recordset
        
155     Call rstAlarme.Open("SELECT * FROM GRB_Alarmes WHERE NoEmploye = " & iNoEmploye, g_connData, adOpenForwardOnly, adLockReadOnly)
        
        'Tant qu'il y a des enregistrements
160     Do While Not rstAlarme.EOF
165       bAfficher = False

170       If rstAlarme.Fields("Date") < ConvertDate(Date) Then
175         bAfficher = True
180       Else
185         If rstAlarme.Fields("Date") = ConvertDate(Date) Then
190           If CDate(rstAlarme.Fields("Heure")) <= Time Then
195             bAfficher = True
200           End If
205         End If
210       End If

215       If bAfficher = True Then
220         Call frmAlarme.Afficher(rstAlarme.Fields("IDAlarme"))
225       End If

230       Call rstAlarme.MoveNext
235     Loop

240     Call rstAlarme.Close
245     Set rstAlarme = Nothing

250     Exit Sub

AfficherErreur:

255     woups "frmDispatch", "tmrAlarme_Timer", Err, Erl
End Sub
