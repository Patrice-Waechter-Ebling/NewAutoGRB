VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacturation 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturation"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFacturation.frx":0000
   ScaleHeight     =   8520
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_export 
      Caption         =   "Exporter vers Excel"
      Height          =   375
      Left            =   120
      TabIndex        =   39
      Top             =   6240
      Width           =   2055
   End
   Begin VB.TextBox txtDescription 
      Height          =   495
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   840
      Width           =   3615
   End
   Begin VB.Frame fraType 
      BackColor       =   &H00000000&
      Caption         =   "Type"
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
      Height          =   1095
      Left            =   1560
      TabIndex        =   33
      Top             =   1080
      Width           =   1695
      Begin VB.OptionButton optType 
         BackColor       =   &H00000000&
         Caption         =   "Tous"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H00000000&
         Caption         =   "Électrique"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H00000000&
         Caption         =   "Mécanique"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   34
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdVerrouiller 
      Caption         =   "Verrouiller Soum"
      Height          =   375
      Left            =   2880
      TabIndex        =   32
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdSommaire 
      Caption         =   "Sommaire"
      Height          =   375
      Left            =   4800
      TabIndex        =   31
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox txtDateOuverture 
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdCommentaire 
      Caption         =   "Commentaires"
      Height          =   375
      Left            =   2160
      TabIndex        =   29
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdModifier 
      Caption         =   "Modifier"
      Height          =   375
      Left            =   8280
      TabIndex        =   28
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdNCRectifier 
      Caption         =   "NC"
      Height          =   375
      Left            =   8280
      TabIndex        =   21
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton cmdReouverture 
      Caption         =   "Annuler Fermeture"
      Height          =   375
      Left            =   6600
      TabIndex        =   24
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton cmdImprimer 
      Caption         =   "Imprimer"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox txtRaisonFermeture 
      Height          =   495
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox txtDateFermeture 
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtClient 
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton cmdSupprimer 
      Caption         =   "Supprimer"
      Height          =   375
      Left            =   6120
      TabIndex        =   19
      Top             =   5760
      Width           =   975
   End
   Begin VB.Frame fraMontrer 
      BackColor       =   &H00000000&
      Caption         =   "Soumissions"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
      Begin VB.OptionButton optMontrer 
         BackColor       =   &H00000000&
         Caption         =   "Fermées"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optMontrer 
         BackColor       =   &H00000000&
         Caption         =   "Ouvertes"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optMontrer 
         BackColor       =   &H00000000&
         Caption         =   "Toutes"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdOuvrirProjSoum 
      Caption         =   "Ouvrir Soum"
      Height          =   375
      Left            =   4800
      TabIndex        =   18
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdFermerProjSoum 
      Caption         =   "Fermer Soum"
      Height          =   375
      Left            =   3480
      TabIndex        =   17
      Top             =   5760
      Width           =   1215
   End
   Begin VB.ComboBox cmbProjSoum 
      Height          =   315
      ItemData        =   "frmFacturation.frx":2F0D
      Left            =   6000
      List            =   "frmFacturation.frx":2F17
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdFacturerRectifier 
      Caption         =   "Facturer"
      Height          =   375
      Left            =   7200
      TabIndex        =   20
      Top             =   5760
      Width           =   975
   End
   Begin VB.ComboBox cmbNoProjSoum 
      Height          =   315
      Left            =   7680
      TabIndex        =   14
      Text            =   "cmbNoProjSoum"
      Top             =   1800
      Width           =   1575
   End
   Begin MSComctlLib.ListView lvwProjets 
      Height          =   3375
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   9135
      _ExtentX        =   16113
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Employé"
         Object.Width           =   1376
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   1746
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Début"
         Object.Width           =   1085
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Fin"
         Object.Width           =   1085
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Description"
         Object.Width           =   6826
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Total"
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "No. Facture"
         Object.Width           =   1799
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Type"
         Object.Width           =   3000
      EndProperty
   End
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   8280
      TabIndex        =   25
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox txtNoProjSoum 
      Height          =   285
      Left            =   7680
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5640
      TabIndex        =   38
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblRaisonFermeture 
      BackStyle       =   0  'Transparent
      Caption         =   "Raison de la fermeture :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblDateFermeture 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date de fermeture : "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblDateOuverture 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date d'ouverture : "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblClient 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Client : "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblHeuresNonFacturees 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1920
      TabIndex        =   27
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label lblHeuresFacturees 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1920
      TabIndex        =   23
      Top             =   6720
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Heures non facturées :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Heures facturées :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label lblTitreProjSoum 
      BackStyle       =   0  'Transparent
      Caption         =   "Numéro de projet :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7680
      TabIndex        =   10
      Top             =   1560
      Width           =   1815
   End
End
Attribute VB_Name = "frmFacturation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Option Explicit

'Index des colonnes du listview
Private Const I_LVW_EMPLOYE             As Integer = 0
Private Const I_LVW_DATE                As Integer = 1
Private Const I_LVW_DEBUT               As Integer = 2
Private Const I_LVW_FIN                 As Integer = 3
Private Const I_LVW_DESCRIPTION         As Integer = 4
Private Const I_LVW_TOTAL               As Integer = 5
Private Const I_LVW_NO_FACTURE          As Integer = 6
Private Const I_LVW_TYPE                As Integer = 7


'Index de optType
Private Const I_OPT_TYPE_ELECTRIQUE     As Integer = 0
Private Const I_OPT_TYPE_MECANIQUE      As Integer = 1
Private Const I_OPT_TYPE_TOUS           As Integer = 2

'Index du combo
Private Const I_CMB_PROJET              As Integer = 0
Private Const I_CMB_SOUMISSION          As Integer = 1

'Caption du bouton cmdFacturerRectifier
Private Const S_FACTURER                As String = "Facturer"
Private Const S_RECTIFIER               As String = "Rectifier"
Private Const S_NC                      As String = "NC"

'Caption des Option Buttons
'Si c'est un projet
Private Const S_PROJ_OUVERT             As String = "Ouverts"
Private Const S_PROJ_FERME              As String = "Fermés"
Private Const S_PROJ_TOUS               As String = "Tous"

'Si c'est une soumission
Private Const S_SOUM_OUVERT             As String = "Ouvertes"
Private Const S_SOUM_FERME              As String = "Fermées"
Private Const S_SOUM_TOUS               As String = "Toutes"

'Caption de fraMontrer
Private Const S_FRA_PROJ                As String = "Projets"
Private Const S_FRA_SOUM                As String = "Soumissions"

'Index des Option Buttons
Private Const I_OPT_TOUS                As Integer = 0
Private Const I_OPT_OUVERT              As Integer = 1
Private Const I_OPT_FERME               As Integer = 2

Private Enum enumType
  TYPE_PROJET = 0
  TYPE_SOUMISSION = 1
End Enum

Private m_eType       As enumType

Public m_iIDClient    As Integer
Public m_sDescription As String

Public m_bModifClient As Boolean

Private m_bLoading    As Boolean

Private Sub cmbNoProjSoum_Click()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset

15      Set rstProjSoum = New ADODB.Recordset

20      txtNoProjSoum.Text = cmbNoProjSoum.Text
  
25      Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
  
30      If rstProjSoum.Fields("Ouvert") = True Then
35        cmdFermerProjSoum.Enabled = True
40      Else
45        cmdFermerProjSoum.Enabled = False
50      End If

55      If rstProjSoum.Fields("Verrouillé") = True Then
60        If m_eType = TYPE_SOUMISSION Then
65          cmdVerrouiller.Caption = "Déverrouiller Soum"
70        Else
75          cmdVerrouiller.Caption = "Déverrouiller Proj"
80        End If
85      Else
90        If m_eType = TYPE_SOUMISSION Then
95          cmdVerrouiller.Caption = "Verrouiller Soum"
100       Else
105         cmdVerrouiller.Caption = "Verrouiller Proj"
110       End If
115     End If

120     Call rstProjSoum.Close
125     Set rstProjSoum = Nothing
  
130     Call AfficherProjSoum

135     Exit Sub

AfficherErreur:

140     woups "frmFacturation", "cmbNoProjSoum_Click", Err, Erl
End Sub

Private Sub cmbNoProjSoum_KeyUp(KeyCode As Integer, Shift As Integer)
  
5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer
  
15      For iCompteur = 0 To cmbNoProjSoum.ListCount - 1
20        If UCase(cmbNoProjSoum.LIST(iCompteur)) = UCase(cmbNoProjSoum.Text) Then
25          cmbNoProjSoum.ListIndex = iCompteur
      
30          Exit For
35        End If
40      Next

45      Exit Sub

AfficherErreur:

50      woups "frmFacturation", "cmbProjSoum_KeyUp", Err, Erl
End Sub


Private Sub AfficherProjSoum()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset
15      Dim rstClient   As ADODB.Recordset
  
20      Set rstProjSoum = New ADODB.Recordset
25      Set rstClient = New ADODB.Recordset
  
30      Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
  
35      Call rstClient.Open("SELECT NomClient FROM GRB_Client WHERE IDClient = " & rstProjSoum.Fields("NoClient"), g_connData, adOpenDynamic, adLockOptimistic)
  
40      If Not rstClient.EOF Then
45        txtClient.Text = rstClient.Fields("NomClient")
50        txtClient.Tag = rstProjSoum.Fields("NoClient")
55      End If

60      Call rstClient.Close
65      Set rstClient = Nothing
  
70      If Not IsNull(rstProjSoum.Fields("DateOuverture")) Then
75        txtDateOuverture.Text = rstProjSoum.Fields("DateOuverture")
80      Else
85        txtDateOuverture.Text = vbNullString
90      End If
    
95      If optMontrer(I_OPT_TOUS).Value = True Or optMontrer(I_OPT_FERME).Value = True Then
100       If Not IsNull(rstProjSoum.Fields("DateFermeture")) Then
105         txtDateFermeture.Text = rstProjSoum.Fields("DateFermeture")
110       Else
115         txtDateFermeture.Text = vbNullString
120       End If
    
125       If Not IsNull(rstProjSoum.Fields("RaisonFermeture")) Then
130         txtRaisonFermeture.Text = rstProjSoum.Fields("RaisonFermeture")
135       Else
140         txtRaisonFermeture.Text = vbNullString
145       End If
150     End If

155     If Not IsNull(rstProjSoum.Fields("Description")) Then
160       txtDescription.Text = rstProjSoum.Fields("Description")
165     Else
170       txtDescription.Text = vbNullString
175     End If

180     Call rstProjSoum.Close
185     Set rstProjSoum = Nothing
  
190     Call RemplirListView

195     Exit Sub

AfficherErreur:

200     woups "frmFacturation", "AfficherProjSoum", Err, Erl
End Sub

Private Sub cmd_export_Click()
Call vb_to_excel

End Sub

Private Sub cmdCommentaire_Click()

5       On Error GoTo AfficherErreur

10      If cmbProjSoum.ListIndex = I_CMB_PROJET Then
15        Call frmCommentairesProjSoum.Afficher(cmbNoProjSoum.Text, True)
20      Else
25        Call frmCommentairesProjSoum.Afficher(cmbNoProjSoum.Text, False)
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmFacturation", "cmdCommentaire_Click", Err, Erl
End Sub

Private Sub cmdFacturerRectifier_Click()

5       On Error GoTo AfficherErreur

10      Dim sNoFacture As String
15      Dim rstFacture As ADODB.Recordset
20      Dim sWhere     As String
25      Dim iCompteur  As Integer
  
30      If cmdFacturerRectifier.Caption = S_FACTURER Then
          'Change la valeur du champs "Facturé" dans la table GRB_Punch pour True et
          'ajoute le numéro de la facture dans le champs "NoFacture"
   
35        sNoFacture = InputBox("Entrez le numéro de la facture")
    
          'Le numéro de facture peut être vide, mais si il ne l'est pas, il doit
          'être numérique
40        If sNoFacture <> vbNullString Then
45          If Not IsNumeric(sNoFacture) Then
50            Call MsgBox("Le numéro de facture est invalide", vbOKOnly, "Erreur")
      
55            Exit Sub
60          End If
  
65          sWhere = "IDPunch In ("
  
70          For iCompteur = 1 To lvwProjets.ListItems.count
              'Si l'élément est sélectionné
75            If lvwProjets.ListItems(iCompteur).Selected = True Then
                'Si la condition where est vide, c'est parce que c'est le premier élément
                'sélectionné
80              If sWhere = "IDPunch In (" Then
85                sWhere = sWhere & lvwProjets.ListItems(iCompteur).Tag
90              Else
95                sWhere = sWhere & "," & lvwProjets.ListItems(iCompteur).Tag
100             End If
105           End If
110         Next

115         sWhere = sWhere & ")"
    
120         Set rstFacture = New ADODB.Recordset
    
            'Ouverture des enregistrements sélectionnés dans le ListView
125         Call rstFacture.Open("SELECT * FROM GRB_Punch WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
    
130         Do While Not rstFacture.EOF
              'Mettre la facturation à true et remplir le numéro de facture
135           rstFacture.Fields("Facturé") = True
140           rstFacture.Fields("NoFacture") = sNoFacture
      
145           Call rstFacture.Update
        
150           Call rstFacture.MoveNext
155         Loop
      
160         Call rstFacture.Close
165         Set rstFacture = Nothing
      
170         Call RemplirListView(lvwProjets.SelectedItem.Index)
175       End If
180     Else
          'Change la valeur du champs "Facturé" dans la table GRB_Punch pour False et
          'ajoute le numéro de la facture dans le champs "NoFacture"
  
185       sWhere = "IDPunch In ("
  
190       For iCompteur = 1 To lvwProjets.ListItems.count
            'Si l'élément est sélectionné
195         If lvwProjets.ListItems(iCompteur).Selected = True Then
              'Si la condition where est vide, c'est parce que c'est le premier élément
              'sélectionné
200           If sWhere = "IDPunch In (" Then
205             sWhere = sWhere & lvwProjets.ListItems(iCompteur).Tag
210           Else
215             sWhere = sWhere & "," & lvwProjets.ListItems(iCompteur).Tag
220           End If
225         End If
230       Next
    
235       sWhere = sWhere & ")"
    
240       Set rstFacture = New ADODB.Recordset
    
          'Ouverture des enregistrements sélectionnés dans le ListView
245       Call rstFacture.Open("SELECT Facturé, NoFacture FROM GRB_Punch WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
       
          'Tant que ce n'est pas la fin du recordset
250       Do While Not rstFacture.EOF
            'Mettre la facturation à false et vider le numéro de facture
255         rstFacture.Fields("Facturé") = False
260         rstFacture.Fields("NoFacture") = vbNullString
    
265         Call rstFacture.Update
    
270         Call rstFacture.MoveNext
275      Loop
    
280       Call rstFacture.Close
285       Set rstFacture = Nothing
  
290       Call RemplirListView(lvwProjets.SelectedItem.Index)
295     End If

300     Exit Sub

AfficherErreur:

305     woups "frmFacturation", "cmdFacturerRectifier_Click", Err, Erl
End Sub

Private Sub cmdModifier_Click()

5       On Error GoTo AfficherErreur

10      Dim rstPunch    As ADODB.Recordset
15      Dim rstProjSoum As ADODB.Recordset

20      If PeutModifier = True Then
25        m_bModifClient = True

30        Call frmChoixClient.Show(vbModal)

35        m_bModifClient = False

40        If m_iIDClient <> txtClient.Tag Then
45          Set rstProjSoum = New ADODB.Recordset

50          Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

55          rstProjSoum.Fields("NoClient") = m_iIDClient
60          rstProjSoum.Fields("Description") = m_sDescription

65          Call rstProjSoum.Update

70          Call rstProjSoum.Close
75          Set rstProjSoum = Nothing

80          Set rstPunch = New ADODB.Recordset

85          Call rstPunch.Open("SELECT * FROM GRB_Punch WHERE NoProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

90          If Not rstPunch.EOF Then
95            If MsgBox("Voulez-vous modifier tous les punch ?", vbYesNo) = vbYes Then
100             Do While Not rstPunch.EOF
105               rstPunch.Fields("NoClient") = m_iIDClient

110               Call rstPunch.Update

115               Call rstPunch.MoveNext
120             Loop
125           End If
130         End If

135         Call rstPunch.Close
140         Set rstPunch = Nothing

145         Call AfficherProjSoum
150       End If
155     End If

160     Exit Sub

AfficherErreur:

165     woups "frmFacturation", "cmdModifier_Click", Err, Erl
End Sub

Private Function PeutModifier() As Boolean

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum   As ADODB.Recordset
15      Dim rstProjet     As ADODB.Recordset
20      Dim rstSoumission As ADODB.Recordset
25      Dim bPeutModifier As Boolean

30      Set rstProjSoum = New ADODB.Recordset

35      Call rstProjSoum.Open("SELECT Ouvert, Type FROM GRB_ProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

40      If rstProjSoum.Fields("Ouvert") = True Then
45        If rstProjSoum.Fields("Type") = "P" Then
50          Set rstProjet = New ADODB.Recordset

55          Call rstProjet.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

60          If rstProjet.EOF Then
65            Call rstProjet.Close

70            Call rstProjet.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

75            If rstProjet.EOF Then
80              bPeutModifier = True
85            Else
90              Call MsgBox("Le client doit être modifié dans l'écran des projets mécaniques!", vbOKOnly, "Erreur")

95              bPeutModifier = False
100           End If

105           Call rstProjet.Close
110         Else
115           Call MsgBox("Le client doit être modifié dans l'écran des projets électriques!", vbOKOnly, "Erreur")

120           Call rstProjet.Close

125           bPeutModifier = False
130         End If

135         Set rstProjet = Nothing
140       Else
145          Set rstSoumission = New ADODB.Recordset

150         Call rstSoumission.Open("SELECT * FROM GRB_SoumissionElec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

155         If rstSoumission.EOF Then
160           Call rstSoumission.Close

165           Call rstSoumission.Open("SELECT * FROM GRB_SoumissionMec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

170           If rstSoumission.EOF Then
175             bPeutModifier = True
180           Else
185             Call MsgBox("Le client doit être modifié dans l'écran des soumissions mécaniques!", vbOKOnly, "Erreur")

190             bPeutModifier = False
195           End If

200           Call rstSoumission.Close
205         Else
210           Call MsgBox("Le client doit être modifié dans l'écran des soumissions électriques!", vbOKOnly, "Erreur")

215           Call rstSoumission.Close

220           bPeutModifier = False
225         End If

230         Set rstSoumission = Nothing
235       End If
240     Else
245       If rstProjSoum.Fields("Type") = "P" Then
250         Call MsgBox("Ce projet est fermé!", vbOKOnly, "Erreur")
255       Else
260         Call MsgBox("Cette soumission est fermée!", vbOKOnly, "Erreur")
265       End If

270       bPeutModifier = False
275     End If

280     Call rstProjSoum.Close
285     Set rstProjSoum = Nothing

290     PeutModifier = bPeutModifier

295     Exit Function

AfficherErreur:

300     woups "frmFacturation", "PeutModifier", Err, Erl
End Function

Private Sub cmdNCRectifier_Click()

5       On Error GoTo AfficherErreur

10      Dim rstFacture As ADODB.Recordset
15      Dim sWhere     As String
20      Dim iCompteur  As Integer
  
25      If cmdNCRectifier.Caption = S_NC Then
          'Change la valeur du champs "Facturé" dans la table GRB_Punch pour True et
          'ajoute NC dans le champs "NoFacture"
30        sWhere = "IDPunch In ("
          
35        For iCompteur = 1 To lvwProjets.ListItems.count
            'Si l'élément est sélectionné
40          If lvwProjets.ListItems(iCompteur).Selected = True Then
              'Si la condition where est vide, c'est parce que c'est le premier élément
              'sélectionné
45            If sWhere = "IDPunch In (" Then
50              sWhere = sWhere & lvwProjets.ListItems(iCompteur).Tag
55            Else
60              sWhere = sWhere & "," & lvwProjets.ListItems(iCompteur).Tag
65            End If
70          End If
75        Next
    
80        sWhere = sWhere & ")"

85        Set rstFacture = New ADODB.Recordset
    
          'Ouverture des enregistrements sélectionnés dans le ListView
90        Call rstFacture.Open("SELECT * FROM GRB_Punch WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
    
95        Do While Not rstFacture.EOF
            'Mettre la facturation à true et remplir le numéro de facture
100         rstFacture.Fields("Facturé") = True
105         rstFacture.Fields("NoFacture") = "NC"
      
110         Call rstFacture.Update
        
115         Call rstFacture.MoveNext
120       Loop
      
125       Call rstFacture.Close
130       Set rstFacture = Nothing
      
135       Call RemplirListView(lvwProjets.SelectedItem.Index)
140     Else
          'Change la valeur du champs "Facturé" dans la table GRB_Punch pour False et
          'enlève NC
145       sWhere = "IDPunch In ("
  
150       For iCompteur = 1 To lvwProjets.ListItems.count
            'Si l'élément est sélectionné
155         If lvwProjets.ListItems(iCompteur).Selected = True Then
              'Si la condition where est vide, c'est parce que c'est le premier élément
              'sélectionné
160           If sWhere = "IDPunch In (" Then
165             sWhere = sWhere & lvwProjets.ListItems(iCompteur).Tag
170           Else
175             sWhere = sWhere & "," & lvwProjets.ListItems(iCompteur).Tag
180           End If
185         End If
190       Next
    
195       sWhere = sWhere & ")"

200       Set rstFacture = New ADODB.Recordset
    
          'Ouverture des enregistrements sélectionnés dans le ListView
205       Call rstFacture.Open("SELECT * FROM GRB_Punch WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
       
          'Tant que ce n'est pas la fin du recordset
210       Do While Not rstFacture.EOF
            'Mettre la facturation à false et vider le numéro de facture
215         rstFacture.Fields("Facturé") = False
220         rstFacture.Fields("NoFacture") = vbNullString
    
225         Call rstFacture.Update
    
230         Call rstFacture.MoveNext
235      Loop
    
240       Call rstFacture.Close
245       Set rstFacture = Nothing
  
250       Call RemplirListView(lvwProjets.SelectedItem.Index)
255     End If

260     Exit Sub

AfficherErreur:

265     woups "frmFacturation", "cmdFacturerRectifier_Click", Err, Erl
End Sub

Private Sub Cmdfermer_Click()

5       On Error GoTo AfficherErreur

        'Fermer de la fênêtre
10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmFacturation", "cmdFermer_Click", Err, Erl
End Sub

Private Sub cmdImprimer_Click()
Dim intdummie As Integer



5       On Error GoTo AfficherErreur

10      If m_eType = TYPE_PROJET Then
15        Call frmChoixDateImpressionFacturation.Afficher(txtNoProjSoum.Text, True, txtClient.Text, txtDescription.Text)
20      Else
25        Call frmChoixDateImpressionFacturation.Afficher(txtNoProjSoum.Text, False, txtClient.Text, txtDescription.Text)
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmFacturation", "cmdImprimer_Click", Err, Erl
End Sub

Private Function vb_to_excel()



6       Dim iCount As Integer
10      Dim oXLApp As Excel.Application         'Declare the object variables
15      Dim oXLBook As Excel.Workbook
20      Dim oXLSheet As Excel.Worksheet
        Dim data_array(1 To 1500, 1 To 7) As Variant 'modifier pour intégré une nouvelle onglet
        Dim r As Integer
25      Set oXLApp = New Excel.Application    'Create a new instance of Excel
30      Set oXLBook = oXLApp.Workbooks.Add    'Add a new workbook
35      Set oXLSheet = oXLBook.Worksheets(1)  'Work with the first worksheet
        oXLApp.Visible = False

'on inscrit les valeurs du listbox dans un tableau
r = 1
Do While r <= lvwProjets.ListItems.count <> Empty
        data_array(r, 1) = lvwProjets.ListItems(r)
        data_array(r, 2) = lvwProjets.ListItems(r).SubItems(1)
        data_array(r, 3) = lvwProjets.ListItems(r).SubItems(2)
        data_array(r, 4) = lvwProjets.ListItems(r).SubItems(3)
        data_array(r, 5) = lvwProjets.ListItems(r).SubItems(4) 'Ajouter la description a la table excel 2017-06-26 GLL
        data_array(r, 6) = CDbl(lvwProjets.ListItems(r).SubItems(5))
        data_array(r, 7) = lvwProjets.ListItems(r).SubItems(7) 'Ajouter pour avoir le tableau complet en Excel
        
        r = r + 1
       
Loop




'creation en-tête de colonne
oXLSheet.Range("A1: G1").Font.Bold = True
oXLSheet.Range("A:G").HorizontalAlignment = xlCenter
oXLSheet.Range("A1: G1").Value = Array("Employé", "Date", "Debut", "Fin", "Description", "Total", "Type") 'GLL


'inscription des valeur du tableau dans excel
oXLSheet.Range("A2").Resize(r, 7).Value = data_array
'ajustement largeur des colonne
oXLSheet.Range("A:G").Columns.AutoFit
oXLApp.Visible = True

        







End Function


Private Sub cmdOuvrirProjSoum_Click()

5       On Error GoTo AfficherErreur

10      Dim sNumero     As String
15      Dim rstProjSoum As Recordset
20      Dim sQuestion   As String
25      Dim sType       As String
30      Dim bNoValide   As Boolean

35      Select Case m_eType
          Case TYPE_PROJET:
40          sQuestion = "Quel est le numéro du projet?"
45          sType = "P"
50        Case TYPE_SOUMISSION:
55          sQuestion = "Quel est le numéro de la soumission?"
60          sType = "S"
65      End Select
    
70      sNumero = InputBox(sQuestion)
  
75      If sNumero <> vbNullString Then
80        bNoValide = True

85        If ValiderFormatNumeroProjSoum(sNumero) = False Then
90          bNoValide = False
95        End If

100       If bNoValide = True Then
105         If m_eType = TYPE_PROJET Then
110           If ValiderFormatProjet(sNumero) = False Then
115             bNoValide = False
120           End If
125         Else
130           If ValiderFormatSoumission(sNumero) = False Then
135             bNoValide = False
140           End If
145         End If
150       End If

155       If bNoValide = False Then
160         Exit Sub
165       End If

170       Call frmChoixClient.Show(vbModal)
  
175       Set rstProjSoum = New ADODB.Recordset
  
180       Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & sNumero & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
185       If rstProjSoum.EOF Then
190         Call rstProjSoum.AddNew
    
195         rstProjSoum.Fields("IDProjSoum") = sNumero
200         rstProjSoum.Fields("NoClient") = m_iIDClient
205         rstProjSoum.Fields("Description") = m_sDescription
210         rstProjSoum.Fields("DateOuverture") = ConvertDate(Date)
215         rstProjSoum.Fields("Ouvert") = True
220         rstProjSoum.Fields("Type") = sType
      
225         Call rstProjSoum.Update
    
230         Call RemplirComboProjSoum
235       Else
240         Call MsgBox("Ce numéro existe déjà!", vbOKOnly, "Erreur")
245       End If
    
250       Call rstProjSoum.Close
255       Set rstProjSoum = Nothing
260     End If

265     Exit Sub

AfficherErreur:

270     woups "frmFacturation", "cmdOuvrirProjSoum_Click", Err, Erl
End Sub

Private Sub cmdFermerProjSoum_Click()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As Recordset
15      Dim sQuestion   As String
20      Dim sRaison     As String
  
25      Select Case m_eType
          Case TYPE_PROJET:     sQuestion = "Voulez-vous vraiment fermer ce projet?"
30        Case TYPE_SOUMISSION: sQuestion = "Voulez-vous vraiment fermer cette soumission?"
35      End Select
      
40      If MsgBox(sQuestion, vbYesNo) = vbYes Then
45        Set rstProjSoum = New ADODB.Recordset

50        Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
55        rstProjSoum.Fields("Ouvert") = False
60        rstProjSoum.Fields("DateFermeture") = ConvertDate(Date)
    
65        If m_eType = TYPE_SOUMISSION Then
70          sRaison = InputBox("Quelle est la raison de la fermeture?")
    
75          rstProjSoum.Fields("RaisonFermeture") = sRaison
80        End If
            
85        Call rstProjSoum.Update
    
90        Call rstProjSoum.Close
95        Set rstProjSoum = Nothing
    
100       Call RemplirComboProjSoum
105     End If

110     Exit Sub

AfficherErreur:

115     woups "frmFacturation", "cmdFermerProjSoum_Click", Err, Erl
End Sub

Private Sub cmdReouverture_Click()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As Recordset
15      Dim sQuestion   As String
  
20      If cmbNoProjSoum.ListIndex <> -1 Then
25        Select Case m_eType
            Case TYPE_PROJET:     sQuestion = "Voulez-vous vraiment annuler la fermeture de ce projet?"
30          Case TYPE_SOUMISSION: sQuestion = "Voulez-vous vraiment annuler la fermeture de cette soumission?"
35        End Select
      
40        If MsgBox(sQuestion, vbYesNo) = vbYes Then
45          Set rstProjSoum = New ADODB.Recordset

50          Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
55          rstProjSoum.Fields("Ouvert") = True
60          rstProjSoum.Fields("RaisonFermeture") = Null
            
65          Call rstProjSoum.Update
    
70          Call rstProjSoum.Close
75          Set rstProjSoum = Nothing

80          Call ViderValeurs
    
85          Call RemplirComboProjSoum
90        End If
95      End If

100     Exit Sub

AfficherErreur:

105     woups "frmFacturation", "cmdReouverture_Click", Err, Erl
End Sub

Private Sub cmdSommaire_Click()

5       On Error GoTo AfficherErreur

10      Call frmChoixDateSommairePunch.Show(vbModal)

15      Exit Sub

AfficherErreur:

20      woups "frmFacturation", "cmdImprimer_Click", Err, Erl
End Sub

Private Sub cmdsupprimer_Click()

5       On Error GoTo AfficherErreur

10      Dim sMessage As String
15      Dim sErreur  As String
  
20      Call RemplirListView
  
25      If m_eType = TYPE_PROJET Then
30        sMessage = "Voulez-vous vraiment effacer le projet " & txtNoProjSoum.Text & "?"
35        sErreur = "Impossible de supprimer ce projet car il y a déjà des punchs!"
40      Else
45        sMessage = "Voulez-vous vraiment effacer la soumission " & txtNoProjSoum.Text & "?"
50        sErreur = "Impossible de supprimer cette soumission car il y a déjà des punchs!"
55      End If

60      If lvwProjets.ListItems.count = 0 Then
65        If MsgBox(sMessage, vbYesNo) = vbYes Then
70          Call g_connData.Execute("DELETE * FROM GRB_ProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'")
      
75          Call ViderValeurs
      
80          Call RemplirComboProjSoum

85          If cmbNoProjSoum.ListCount = 0 Then
90            Call ViderValeurs
95          End If
100       End If
105     Else
110       Call MsgBox(sErreur, vbOKOnly, "Erreur")
115     End If

120     Exit Sub

AfficherErreur:

125     woups "frmFacturation", "cmdSupprimer_Click", Err, Erl
End Sub

Private Sub ViderValeurs()

5       On Error GoTo AfficherErreur

10      txtClient.Text = vbNullString
15      txtDescription.Text = vbNullString
20      txtDateOuverture.Text = vbNullString
25      txtDateFermeture.Text = vbNullString
30      txtRaisonFermeture.Text = vbNullString

35      Exit Sub

AfficherErreur:

40      woups "frmFacturation", "ViderValeurs", Err, Erl
End Sub

Private Sub cmdVerrouiller_Click()
  
5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset
  
15      Set rstProjSoum = New ADODB.Recordset
  
20      Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

25      Select Case cmdVerrouiller.Caption
          Case "Verrouiller Soum":   rstProjSoum.Fields("Verrouillé") = True
30        Case "Verrouiller Proj":   rstProjSoum.Fields("Verrouillé") = True
35        Case "Déverrouiller Soum": rstProjSoum.Fields("Verrouillé") = False
40        Case "Déverrouiller Proj": rstProjSoum.Fields("Verrouillé") = False
45      End Select
  
50      Call rstProjSoum.Update
  
55      Call rstProjSoum.Close
60      Set rstProjSoum = Nothing
  
65      Call cmbNoProjSoum_Click

70      Exit Sub

AfficherErreur:

75      woups "frmFacturation", "cmdVerrouiller_Click", Err, Erl
End Sub

Private Sub Command1_Click()

Call vb_to_excel

End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      m_bLoading = True

15      cmbProjSoum.ListIndex = I_CMB_PROJET
20      cmdFacturerRectifier.Enabled = False
25      cmdNCRectifier.Enabled = False
30      optMontrer(I_OPT_OUVERT).Value = True
35      optType(I_OPT_TYPE_TOUS).Value = True

40      m_bLoading = False

45      Call RemplirComboProjSoum

50      Screen.MousePointer = vbDefault

55      Exit Sub

AfficherErreur:

60      woups "frmFacturation", "Form_Load", Err, Erl
End Sub

Private Sub cmbProjSoum_Click()

5       On Error GoTo AfficherErreur

        'Rempli le cmbNoProjet avec les numéros de projet
10      Select Case cmbProjSoum.ListIndex
          Case I_CMB_PROJET:
15          m_eType = TYPE_PROJET
20          lblTitreProjSoum.Caption = "Numéro de projet"
25          cmdOuvrirProjSoum.Caption = "Ouvrir Projet"
30          cmdFermerProjSoum.Caption = "Fermer Projet"
35          fraMontrer.Caption = S_FRA_PROJ
40          optMontrer(I_OPT_TOUS).Caption = S_PROJ_TOUS
45          optMontrer(I_OPT_OUVERT).Caption = S_PROJ_OUVERT
50          optMontrer(I_OPT_FERME).Caption = S_PROJ_FERME
      
55        Case I_CMB_SOUMISSION:
60          m_eType = TYPE_SOUMISSION
65          lblTitreProjSoum.Caption = "Numéro de soumission"
70          cmdOuvrirProjSoum.Caption = "Ouvrir Soum"
75          cmdFermerProjSoum.Caption = "Fermer Soum"
80          fraMontrer.Caption = S_FRA_SOUM
85          optMontrer(I_OPT_TOUS).Caption = S_SOUM_TOUS
90          optMontrer(I_OPT_OUVERT).Caption = S_SOUM_OUVERT
95          optMontrer(I_OPT_FERME).Caption = S_SOUM_FERME
100     End Select
  
105     Call RemplirComboProjSoum

110     Exit Sub

AfficherErreur:

115     woups "frmFacturation", "cmbProjSoum_Click", Err, Erl
End Sub

Private Sub RemplirComboProjSoum()

5       On Error GoTo AfficherErreur

        'Rempli le cmbNoProjet avec les numéros de projet et soumissions
        'Si bOuvert est à True, on affiche seulement ceux qui sont ouverts actuellement
10      Dim rstProjet As ADODB.Recordset
15      Dim sType     As String
20      Dim sWhere    As String
  
25      If m_bLoading = False Then
30        Select Case m_eType
            Case TYPE_PROJET:     sType = "P"
35          Case TYPE_SOUMISSION: sType = "S"
40        End Select
  
45        If optMontrer(I_OPT_TOUS).Value = True Then
50          sWhere = "Type = '" & sType & "'"
55        Else
60          If optMontrer(I_OPT_OUVERT).Value = True Then
65            sWhere = "Ouvert = True AND Type = '" & sType & "'"
70          Else
75            sWhere = "Ouvert = False AND Type = '" & sType & "'"
80          End If
85        End If
    
90        If optType(I_OPT_TYPE_ELECTRIQUE).Value = True Then
95          sWhere = sWhere & " AND Left(IDProjSoum, 1) = 'E'"
100       Else
105         If optType(I_OPT_TYPE_MECANIQUE).Value = True Then
110           sWhere = sWhere & " AND Left(IDProjSoum, 1) = 'M'"
115         End If
120       End If
    
          'Il faut vider le Combo avant de le remplir
125       Call cmbNoProjSoum.Clear
    
130       Set rstProjet = New ADODB.Recordset
    
          'Ouverture d'un recordset contenant les NoProjet
135       Call rstProjet.Open("SELECT IDProjSoum, Ouvert FROM GRB_ProjSoum WHERE " & sWhere & " ORDER BY IDProjSoum", g_connData, adOpenDynamic, adLockOptimistic)
      
140       Do While Not rstProjet.EOF
            'Ajout du numéro de projet dans le Combo
145         Call cmbNoProjSoum.AddItem(rstProjet.Fields("IDProjSoum"))
            
150         Call rstProjet.MoveNext
155       Loop
      
160       Call rstProjet.Close
165       Set rstProjet = Nothing
    
          'Si il y a des éléments dans le combo, on sélectionne le premier
170       If cmbNoProjSoum.ListCount > 0 Then
175         cmbNoProjSoum.ListIndex = 0
180       Else
185         Call lvwProjets.ListItems.Clear
190       End If
195     End If
  
200     Exit Sub

AfficherErreur:

205     woups "frmFacturation", "RemplirComboProjSoum", Err, Erl
End Sub

Private Sub RemplirListView(Optional ByVal o_iIndex As Integer = 1)

5       On Error GoTo AfficherErreur

        'Remplissage du listView dépendamment du no de projet choisi
10      Dim rstProjet   As ADODB.Recordset
15      Dim itmProjet   As ListItem
20      Dim lColor      As Long
25      Dim sDateDebut  As String
30      Dim sDateFin    As String
35      Dim sTotal      As String
  
        'Il faut vider le listView avant de le remplir
40      Call lvwProjets.ListItems.Clear
    
45      sDateDebut = "TIMESERIAL(Left(GRB_Punch.HeureDébut,2),RIGHT(GRB_Punch.HeureDébut,2),0)"
  
50      sDateFin = "TIMESERIAL(Left(GRB_Punch.HeureFin,2),RIGHT(GRB_Punch.HeureFin,2),0)"
  
55      sTotal = "((" & sDateFin & " - " & sDateDebut & ")* 24) As Total"
  
        'Ouverture des enregistrements avec comme filtre, le numéro du projet
60      Set rstProjet = New ADODB.Recordset
        
65      rstProjet.CursorLocation = adUseServer
        
70      Call rstProjet.Open("SELECT GRB_Punch.*, " & sTotal & ", GRB_employés.initiale FROM GRB_employés INNER JOIN GRB_Punch ON GRB_employés.noemploye = GRB_Punch.NoEmploye WHERE NoProjet = '" & txtNoProjSoum.Text & "' ORDER BY [Date], HeureDébut, HeureFin", g_connData, adOpenDynamic, adLockOptimistic)
    
        'Tant que ce n'est pas la fin des enregistrements
75      Do While Not rstProjet.EOF
          'Vérification du champs "Facturé", si il est à vrai, il faut l'inscrire en
          'rouge dans le ListView, sinon, il faut l'inscrire en noir
80        If rstProjet.Fields("Facturé") = "Vrai" Then
85          lColor = COLOR_ROUGE
90        Else
95          lColor = COLOR_NOIR
100       End If
      
105       Set itmProjet = lvwProjets.ListItems.Add
      
110       itmProjet.Tag = rstProjet.Fields("IDPunch")
      
          'Initiales de l'employé
115       itmProjet.Text = rstProjet.Fields("Initiale")
120       itmProjet.ForeColor = lColor
      
          'Date
125       itmProjet.SubItems(I_LVW_DATE) = rstProjet.Fields("Date")
130       itmProjet.ListSubItems(I_LVW_DATE).ForeColor = lColor
      
          'Début
135       If Not IsNull(rstProjet.Fields("HeureDébut")) Then
140         itmProjet.SubItems(I_LVW_DEBUT) = rstProjet.Fields("HeureDébut")
145       Else
150         itmProjet.SubItems(I_LVW_DEBUT) = vbNullString
155       End If
    
160       itmProjet.ListSubItems(I_LVW_DEBUT).ForeColor = lColor
                       
          'Fin
165       If Not IsNull(rstProjet.Fields("HeureFin")) Then
170         itmProjet.SubItems(I_LVW_FIN) = rstProjet.Fields("HeureFin")
175       Else
180         itmProjet.SubItems(I_LVW_FIN) = vbNullString
185       End If
                       
190       itmProjet.ListSubItems(I_LVW_FIN).ForeColor = lColor
    
          'Description
195       If Not IsNull(rstProjet.Fields("Commentaire")) Then
200         itmProjet.SubItems(I_LVW_DESCRIPTION) = rstProjet.Fields("Commentaire")
205       Else
210         itmProjet.SubItems(I_LVW_DESCRIPTION) = vbNullString
215       End If
      
220       itmProjet.ListSubItems(I_LVW_DESCRIPTION).ForeColor = lColor
     
          'Total
225       If Not IsNull(rstProjet.Fields("Total")) Then
230         itmProjet.SubItems(I_LVW_TOTAL) = Round(rstProjet.Fields("Total"), 2)
235       Else
240         itmProjet.SubItems(I_LVW_TOTAL) = vbNullString
245       End If
    
250       itmProjet.ListSubItems(I_LVW_TOTAL).ForeColor = lColor
     
          'Numéro de facture
255       If Not IsNull(rstProjet.Fields("NoFacture")) Then
260         itmProjet.SubItems(I_LVW_NO_FACTURE) = rstProjet.Fields("NoFacture")
265       Else
270         itmProjet.SubItems(I_LVW_NO_FACTURE) = vbNullString
275       End If
      
280       itmProjet.ListSubItems(I_LVW_NO_FACTURE).ForeColor = lColor
      
            'Type
         If Not IsNull(rstProjet.Fields("Type")) Then
           itmProjet.SubItems(I_LVW_TYPE) = rstProjet.Fields("Type")
         Else
           itmProjet.SubItems(I_LVW_TYPE) = vbNullString
         End If
        
         itmProjet.ListSubItems(I_LVW_TYPE).ForeColor = lColor

285       Call rstProjet.MoveNext
290     Loop

295     If lvwProjets.ListItems.count > 0 Then
300       Call lvwProjets.ListItems(o_iIndex).EnsureVisible
305     End If
    
310     Call rstProjet.Close
315     Set rstProjet = Nothing
  
320     Call CalculerTotaux

325     Exit Sub

AfficherErreur:

330     woups "frmFacturation", "RemplirListView", Err, Erl
End Sub

Private Sub CalculerTotaux()

5       On Error GoTo AfficherErreur

10      Dim dblTotalFacture    As Double
15      Dim dblTotalNonFacture As Double
20      Dim iCompteur          As Integer
  
25      For iCompteur = 1 To lvwProjets.ListItems.count
30        If lvwProjets.ListItems(iCompteur).SubItems(I_LVW_NO_FACTURE) <> vbNullString Then
35          dblTotalFacture = dblTotalFacture + CDbl(lvwProjets.ListItems(iCompteur).SubItems(I_LVW_TOTAL))
40        Else
45          dblTotalNonFacture = dblTotalNonFacture + CDbl(lvwProjets.ListItems(iCompteur).SubItems(I_LVW_TOTAL))
50        End If
55      Next
  
60      lblHeuresFacturees.Caption = Round(dblTotalFacture, 2)
65      lblHeuresNonFacturees.Caption = Round(dblTotalNonFacture, 2)

70      Exit Sub

AfficherErreur:

75      woups "frmFacturation", "CalculerTotaux", Err, Erl
End Sub

Private Sub VerifierSelection()

5       On Error GoTo AfficherErreur

        'D'après les éléments sélectionner dans le ListView, cette méthode active
        'le bon bouton
10      Dim iCompteur As Integer
15      Dim iSelected As Integer
20      Dim iFacture  As Integer
25      Dim iNC       As Integer
30      Dim iNon      As Integer
  
        'Boucle servant à compter le nombre d'éléments sélectionnés dans le ListView,
        'le nombre d'éléments facturés et nombre d'éléments non facturés
35      For iCompteur = 1 To lvwProjets.ListItems.count
40        If lvwProjets.ListItems(iCompteur).Selected Then
            'Compte les éléments sélectionnés
45          iSelected = iSelected + 1
      
50          If lvwProjets.ListItems(iCompteur).SubItems(I_LVW_NO_FACTURE) = "NC" Then
              'Compte les nc
55            iNC = iNC + 1
60          Else
65            If lvwProjets.ListItems(iCompteur).SubItems(I_LVW_NO_FACTURE) <> vbNullString Then
                'Compte les factures
70              iFacture = iFacture + 1
75            Else
                'Compte les non
80              iNon = iNon + 1
85            End If
90          End If
95        End If
100     Next
    
        'Si tous les éléments sélectionnés ont été facturés
105     If iSelected = iFacture Then
110       cmdFacturerRectifier.Enabled = True
115       cmdNCRectifier.Enabled = False

120       cmdFacturerRectifier.Caption = S_RECTIFIER
125     Else
          'Si tous les éléments sélectionnés sont NC
130       If iSelected = iNC Then
135         cmdFacturerRectifier.Enabled = False
140         cmdNCRectifier.Enabled = True

145         cmdNCRectifier.Caption = S_RECTIFIER
150       Else
            'Si tous les éléments sélectionnés n'ont pas été facturés
155         If iSelected = iNon Then
160           cmdFacturerRectifier.Enabled = True
165           cmdNCRectifier.Enabled = True

170           cmdFacturerRectifier.Caption = S_FACTURER
175           cmdNCRectifier.Caption = S_NC
180         Else
              'Si les éléments sélectionnés sont facturés ou non
185           cmdFacturerRectifier.Enabled = False
190           cmdNCRectifier.Enabled = False
195         End If
200       End If
205     End If

210     Exit Sub

AfficherErreur:

215     woups "frmFacturation", "VerifierSelection", Err, Erl
End Sub

Private Sub lvwProjets_ItemClick(ByVal Item As MSComctlLib.ListItem)

5       On Error GoTo AfficherErreur

        'Vérification de la sélection lorsque un Item dans le ListView est cliqué
10      Call VerifierSelection

15      Exit Sub

AfficherErreur:

20      woups "frmFacturation", "lvwProjets_ItemClick", Err, Erl
End Sub

Private Sub lvwProjets_Click()

5       On Error GoTo AfficherErreur

        'Vérification de la sélection lorsque un Item dans le ListView est cliqué
        'Cette méthode est importante puisque si l'utilisateur déclique un Item en tenant
        'la touche "Ctrl" enfoncé, ça ne passe pas dans l'événement ItemClick
10      Call VerifierSelection

15      Exit Sub

AfficherErreur:

20      woups "frmFacturation", "lvwProjets_Click", Err, Erl
End Sub

Private Sub optMontrer_Click(Index As Integer)

5       On Error GoTo AfficherErreur

10      Dim bFermeture As Boolean

15      Call ViderValeurs

20      Select Case Index
          Case I_OPT_TOUS:
25          Call RemplirComboProjSoum

30          bFermeture = True
          
          Case I_OPT_OUVERT:
35          Call RemplirComboProjSoum

40          bFermeture = False
          
          Case I_OPT_FERME:
45          Call RemplirComboProjSoum

50          bFermeture = True
55      End Select

60      lblDateFermeture.Visible = bFermeture
65      txtDateFermeture.Visible = bFermeture
70      lblRaisonFermeture.Visible = bFermeture
75      txtRaisonFermeture.Visible = bFermeture

80      cmdReouverture.Visible = bFermeture

85      Exit Sub

AfficherErreur:

90      woups "frmFacturation", "optMontrer_Click", Err, Erl
End Sub

Private Sub optType_Click(Index As Integer)
        
5       On Error GoTo AfficherErreur

10      Call RemplirComboProjSoum

15      Exit Sub

AfficherErreur:

20      woups "frmFacturation", "optType_Click", Err, Erl
End Sub

Private Function ValiderFormatSoumission(ByVal sNoSoumission As String) As Boolean
        
5       On Error GoTo AfficherErreur

10      If Mid$(sNoSoumission, 3, 1) = "1" Then
15        ValiderFormatSoumission = True
20      Else
25        Call MsgBox("Une soumission doit absolument avoir '1' comme 3e caractère !", vbOKOnly, "Erreur")

30        ValiderFormatSoumission = False
35      End If

40      Exit Function

AfficherErreur:

45      woups "FrmFacturation", "ValiderFormatSoumission", Err, Erl
End Function

Private Function ValiderFormatProjet(ByVal sNoProjet As String) As Boolean
        
5       On Error GoTo AfficherErreur

10      Dim iType As Integer

15      iType = Mid$(sNoProjet, 3, 1)

20      If iType = 4 Or iType = 5 Or iType = 7 Or iType = 9 Then
25        ValiderFormatProjet = True
30      Else
35        Call MsgBox("Un projet ouvert dans cet écran doit absolument avoir '4', '5', '7' ou '9' comme 3e caractère !", vbOKOnly, "Erreur")

40        ValiderFormatProjet = False
45      End If

50      Exit Function

AfficherErreur:

55      woups "FrmFacturationElec", "ValiderFormatProjet", Err, Erl
End Function
