VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChoixBonCommande 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choix des pièces à commander"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   Icon            =   "frmChoixBonCommande.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   10380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDeselectAll 
      Caption         =   "Aucun"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Tous"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Frame fraFournisseur 
      BackColor       =   &H00000000&
      Caption         =   "Fournisseurs"
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
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   10095
      Begin VB.CommandButton cmdAnnulerFRS 
         Caption         =   "Annuler"
         Height          =   375
         Left            =   7680
         TabIndex        =   2
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdOKFRS 
         Caption         =   "OK"
         Height          =   375
         Left            =   8880
         TabIndex        =   3
         Top             =   1920
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvwFournisseur 
         Height          =   1575
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   2778
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fournisseur"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Pers. Ress."
            Object.Width           =   1984
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   1746
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Par"
            Object.Width           =   805
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Valide"
            Object.Width           =   1746
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Prix listé"
            Object.Width           =   1614
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Escompte"
            Object.Width           =   1561
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Prix net"
            Object.Width           =   1614
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Prix spécial"
            Object.Width           =   1720
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Quoter"
            Object.Width           =   1191
         EndProperty
      End
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   7920
      TabIndex        =   7
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdCommander 
      Caption         =   "Commander"
      Height          =   375
      Left            =   9120
      TabIndex        =   8
      Top             =   6000
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwPiece 
      Height          =   5055
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   8916
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Qté"
         Object.Width           =   794
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "No. Item"
         Object.Width           =   3228
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   6720
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Manufacturier"
         Object.Width           =   2037
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Fournisseur"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Stock"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmChoixBonCommande"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Index des colonnes de lvwPiece
Private Const I_COL_QTE            As Integer = 0
Private Const I_COL_NO_ITEM        As Integer = 1
Private Const I_COL_DESCRIPTION    As Integer = 2
Private Const I_COL_MANUFACTURIER  As Integer = 3
Private Const I_COL_FOURNISSEUR    As Integer = 4
Private Const I_COL_QTE_STOCK      As Integer = 5

'Index des colonnes de lvwFournisseur
Private Const I_COL_FRS_FRS        As Integer = 0
Private Const I_COL_FRS_PERS_RESS  As Integer = 1
Private Const I_COL_FRS_DATE       As Integer = 2
Private Const I_COL_FRS_ENTRER_PAR As Integer = 3
Private Const I_COL_FRS_VALIDE     As Integer = 4
Private Const I_COL_FRS_PRIX_LIST  As Integer = 5
Private Const I_COL_FRS_ESCOMPTE   As Integer = 6
Private Const I_COL_FRS_PRIX_NET   As Integer = 7
Private Const I_COL_FRS_PRIX_SP    As Integer = 8
Private Const I_COL_FRS_QUOTER     As Integer = 9

'Énumération servant à savoir si le form est en anglais ou en francais
Private Enum enumLangage
  FRANCAIS = 0
  ANGLAIS = 1
End Enum

Private m_sNoProjet       As String
Private m_frmSource       As Form
Private m_sType           As String

Private m_sIDAchat        As String
Private m_iIndexAchat     As Integer

Private m_collPiece       As Collection
Private m_collNoLigne     As Collection

Private m_eLangage        As enumLangage

Private m_collNoLigneFRS  As Collection
Private m_collPrixList    As Collection
Private m_collPrixOrigine As Collection
Private m_collPrixNet     As Collection
Private m_collEscompte    As Collection
Private m_collPrixSP      As Collection

Public Sub AfficherAchat(ByVal sIDAchat As String, ByVal iIndexAchat As Integer, ByVal eType As enumCatalogue)

5       On Error GoTo AfficherErreur

10      m_sIDAchat = sIDAchat

15      m_iIndexAchat = iIndexAchat

20      Set m_frmSource = frmAchat

25      cmdSelectAll.Visible = True

30      If eType = ELECTRIQUE Then
35        m_sType = "E"
40      Else
45        m_sType = "M"
50      End If

55      Call lvwPiece.ColumnHeaders.Remove(I_COL_QTE_STOCK)

60      Call RemplirListViewPieceAchat

65      Call Me.Show(vbModal)

70      Exit Sub

AfficherErreur:

75      woups "frmChoixBonCommande", "AfficherAchat", Err, Erl
End Sub

Public Sub Afficher(ByVal sNoProjet As String, ByVal frmSource As Form, ByVal iLangage As Integer)

5       On Error GoTo AfficherErreur

        'Méthode pour afficher le form
10      m_sNoProjet = sNoProjet

15      m_eLangage = iLangage

20      Set m_frmSource = frmSource

25      If frmSource.Name = "FrmProjSoumElec" Then
30        m_sType = "E"
35      Else
40        m_sType = "M"
45      End If
  
50      Call RemplirListViewPieces
  
55      Call Me.Show(vbModal)

60      Exit Sub

AfficherErreur:

65      woups "frmChoixBonCommande", "Afficher", Err, Erl
End Sub

Private Sub RemplirListViewPieceAchat()

5       On Error GoTo AfficherErreur

        'Remplis les pièces de l'achat avec la BD
10      Dim rstAchat    As ADODB.Recordset
15      Dim rstFRS      As ADODB.Recordset
20      Dim itmAchat    As ListItem
25      Dim lColor      As Long
    
30      Call lvwPiece.ListItems.Clear
  
35      Set rstFRS = New ADODB.Recordset
40      Set rstAchat = New ADODB.Recordset
  
45      Call rstAchat.Open("SELECT * FROM GRB_Achat_Pieces WHERE IDAchat = '" & m_sIDAchat & "' AND IndexAchat = " & m_iIndexAchat & " ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
  
50      Do While Not rstAchat.EOF
55        If rstAchat.Fields("Recu") = True Then
60          lColor = COLOR_GRIS 'Gris
65        Else
70          If rstAchat.Fields("Commandé") = True Then
75            lColor = COLOR_ORANGE     'COLOR_ORANGE
80          Else
85            lColor = COLOR_NOIR
90          End If
95        End If

100       Set itmAchat = lvwPiece.ListItems.Add
          
          'Quantité
105       If Not IsNull(rstAchat.Fields("Qté")) Then
110         itmAchat.Text = rstAchat.Fields("Qté")
115       Else
120         itmAchat.Text = vbNullString
125       End If

130       itmAchat.ForeColor = lColor
    
135       itmAchat.Tag = rstAchat.Fields("DateRéception")

          'Numéro d'item
140       If Not IsNull(rstAchat.Fields("PIECE")) Then
145         itmAchat.SubItems(I_COL_NO_ITEM) = rstAchat.Fields("PIECE")
150       Else
155         itmAchat.SubItems(I_COL_NO_ITEM) = vbNullString
160       End If

165       itmAchat.ListSubItems(I_COL_NO_ITEM).ForeColor = lColor

170       itmAchat.ListSubItems(I_COL_NO_ITEM).Tag = rstAchat.Fields("NuméroLigne")
            
          'Description en francais
175       If Not IsNull(rstAchat.Fields("DESC_FR")) Then
180         itmAchat.SubItems(I_COL_DESCRIPTION) = rstAchat.Fields("DESC_FR")
185       Else
190         itmAchat.SubItems(I_COL_DESCRIPTION) = vbNullString
195       End If

200       itmAchat.ListSubItems(I_COL_DESCRIPTION).ForeColor = lColor
    
          'On met la description en anglais dans le tag de la description en francais
205       If Not IsNull(rstAchat.Fields("Desc_EN")) Then
210         itmAchat.ListSubItems(I_COL_DESCRIPTION).Tag = rstAchat.Fields("Desc_EN")
215       Else
220         itmAchat.ListSubItems(I_COL_DESCRIPTION).Tag = vbNullString
225       End If
   
          'Fabricant
230       If Not IsNull(rstAchat.Fields("Manufact")) Then
235         itmAchat.SubItems(I_COL_MANUFACTURIER) = rstAchat.Fields("Manufact")
240       Else
245         itmAchat.SubItems(I_COL_MANUFACTURIER) = vbNullString
250       End If

255       itmAchat.ListSubItems(I_COL_MANUFACTURIER).ForeColor = lColor

260       itmAchat.ListSubItems(I_COL_MANUFACTURIER).Tag = rstAchat.Fields("NoRetour")
          
          'Fournisseur
265       If Not IsNull(rstAchat.Fields("IDFRS")) Then
270         If rstAchat.Fields("IDFRS") <> 0 Then
275           Call rstFRS.Open("SELECT NomFournisseur FROM GRB_Fournisseur WHERE IDFRS = " & rstAchat.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
   
              'On affiche le nom dans la colonne
280           itmAchat.SubItems(I_COL_FOURNISSEUR) = rstFRS.Fields("NomFournisseur")
        
              'On affiche l'Id dans le tag
285           itmAchat.ListSubItems(I_COL_FOURNISSEUR).Tag = rstAchat.Fields("IDFRS")
        
290           Call rstFRS.Close
295         Else
300           itmAchat.SubItems(I_COL_FOURNISSEUR) = " "
305         End If
310       Else
315         itmAchat.SubItems(I_COL_FOURNISSEUR) = vbNullString
320       End If

325       itmAchat.ListSubItems(I_COL_FOURNISSEUR).ForeColor = lColor

330       Call rstAchat.MoveNext
335     Loop
  
340     Call rstAchat.Close
345     Set rstAchat = Nothing

350     Set rstFRS = Nothing

355     Exit Sub

AfficherErreur:

360     woups "frmChoixBonCommande", "RemplirListViewAchat", Err, Erl
End Sub

Private Sub RemplirListViewPieces()

5       On Error GoTo AfficherErreur

        'Rempli le ListView selon le no. du projet
10      Dim rstPieces     As ADODB.Recordset
15      Dim rstSection    As ADODB.Recordset
20      Dim rstInventaire As ADODB.Recordset
25      Dim rstFRS        As Recordset
30      Dim itmPieces     As ListItem
35      Dim iCompteur     As Integer
40      Dim bPremierEnr   As Boolean
45      Dim iOrdreSection As Integer
50      Dim sSousSection  As String
55      Dim sSection      As String
60      Dim lCouleur      As Long
  
65      bPremierEnr = True
 
70      If m_eLangage = ANGLAIS Then
75        sSection = "NomSectionEN"
80      Else
85        sSection = "NomSectionFR"
90      End If
  
95      lvwPiece.Sorted = False

100     Set rstFRS = New ADODB.Recordset
105     Set rstPieces = New ADODB.Recordset
110     Set rstSection = New ADODB.Recordset
115     Set rstInventaire = New ADODB.Recordset

        'Ouverture du recordset
120     Call rstPieces.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & m_sNoProjet & "' AND Type = '" & m_sType & "' AND PieceExtraChargeable = False AND PieceExtraNonChargeable = False ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
  
125     Do While Not rstPieces.EOF
130       Set itmPieces = lvwPiece.ListItems.Add
                    
          'Si c'est le premier enregistrement, il faut ajouter la section et la sous-section
135       If bPremierEnr = True Then
140         sSousSection = rstPieces.Fields("SousSection")
145         iOrdreSection = rstPieces.Fields("OrdreSection")
     
            'Pour avoir le nom de la section
            'Si c'est un projet électrique
150         If m_sType = "E" Then
155           Call rstSection.Open("SELECT " & sSection & " FROM GRB_SoumProjSectionElec WHERE IDSection = " & rstPieces.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
160         Else
165           Call rstSection.Open("SELECT " & sSection & " FROM GRB_SoumProjSectionMec WHERE IDSection = " & rstPieces.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
170         End If

            'Ajout du nom de la section
175         If Not IsNull(rstSection.Fields(sSection)) Then
180           itmPieces.SubItems(I_COL_NO_ITEM) = rstSection.Fields(sSection)
185         Else
190           itmPieces.SubItems(I_COL_NO_ITEM) = vbNullString
195         End If
      
200         itmPieces.ListSubItems(I_COL_NO_ITEM).Bold = True
                    
205         Call rstSection.Close
        
210         Set itmPieces = lvwPiece.ListItems.Add
      
            'Ajout du nom de la sous-section
215         If sSousSection = "PAS DE SOUS-SECTION" Then
220           itmPieces.SubItems(I_COL_DESCRIPTION) = vbNullString
225         Else
230           itmPieces.SubItems(I_COL_DESCRIPTION) = sSousSection
235         End If
             
240         itmPieces.Tag = "PAS UNE SECTION"
             
245         itmPieces.ListSubItems(I_COL_DESCRIPTION).Bold = True
      
250         Set itmPieces = lvwPiece.ListItems.Add
      
255         bPremierEnr = False
260       Else
            'Si c'est pas le premier enregistrement, il faut vérifier avec l'ancienne section
265         If iOrdreSection <> rstPieces.Fields("OrdreSection") Then
270           iOrdreSection = rstPieces.Fields("OrdreSection")
        
275           If m_sType = "E" Then
280             Call rstSection.Open("SELECT " & sSection & " FROM GRB_SoumProjSectionElec WHERE IDSection = " & rstPieces.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
285           Else
290             Call rstSection.Open("SELECT " & sSection & " FROM GRB_SoumProjSectionMec WHERE IDSection = " & rstPieces.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
295           End If
        
300           If Not IsNull(rstSection.Fields(sSection)) Then
305             itmPieces.SubItems(I_COL_NO_ITEM) = rstSection.Fields(sSection)
310           Else
315             itmPieces.SubItems(I_COL_NO_ITEM) = vbNullString
320           End If
        
325           itmPieces.ListSubItems(I_COL_NO_ITEM).Bold = True
        
330           Call rstSection.Close
              
335           Set itmPieces = lvwPiece.ListItems.Add
        
340           sSousSection = rstPieces.Fields("SousSection")
        
345           If sSousSection = "PAS DE SOUS-SECTION" Then
350             itmPieces.SubItems(I_COL_DESCRIPTION) = vbNullString
355           Else
360             itmPieces.SubItems(I_COL_DESCRIPTION) = rstPieces.Fields("SousSection")
365           End If
        
370           itmPieces.Tag = "PAS UNE SECTION"
        
375           itmPieces.ListSubItems(I_COL_DESCRIPTION).Bold = True
        
380           Set itmPieces = lvwPiece.ListItems.Add
385         Else
              'il faut vérifier avec l'ancienne sous-section
390           If sSousSection <> rstPieces.Fields("SousSection") Then
395             sSousSection = rstPieces.Fields("SousSection")
          
400             If sSousSection = "PAS DE SOUS-SECTION" Then
405               itmPieces.SubItems(I_COL_DESCRIPTION) = vbNullString
410             Else
415               itmPieces.SubItems(I_COL_DESCRIPTION) = sSousSection
420             End If
        
425             itmPieces.ListSubItems(I_COL_DESCRIPTION).Bold = True

430             itmPieces.Tag = "PAS UNE SECTION"
        
435             Set itmPieces = lvwPiece.ListItems.Add
440           End If
445         End If
450       End If
                      
455       If rstPieces.Fields("Commandé") = True Then
460         lCouleur = COLOR_ORANGE
465       Else
470         If rstPieces.Fields("Recu") = True Then
475           lCouleur = COLOR_GRIS
480         Else
485           lCouleur = COLOR_NOIR
490         End If
495       End If
          
          'Quantité
500       If Not IsNull(rstPieces.Fields("Qté")) Then
505         itmPieces.Text = rstPieces.Fields("Qté")
510       Else
515         itmPieces.Text = vbNullString
520       End If

525       itmPieces.ForeColor = lCouleur
    
530       itmPieces.Tag = "PAS UNE SECTION"
    
          'Numéro d'item
535       If Not IsNull(rstPieces.Fields("NumItem")) Then
540         itmPieces.SubItems(I_COL_NO_ITEM) = rstPieces.Fields("NumItem")
545       Else
550         itmPieces.SubItems(I_COL_NO_ITEM) = vbNullString
555       End If

560       itmPieces.ListSubItems(I_COL_NO_ITEM).ForeColor = lCouleur

565       itmPieces.ListSubItems(I_COL_NO_ITEM).Tag = rstPieces.Fields("NuméroLigne")
    
570       If m_eLangage = FRANCAIS Then
            'Description en francais
575         If Not IsNull(rstPieces.Fields("Desc_FR")) Then
580           itmPieces.SubItems(I_COL_DESCRIPTION) = rstPieces.Fields("Desc_FR")
585         Else
590           itmPieces.SubItems(I_COL_DESCRIPTION) = vbNullString
595         End If
600       Else
            'Description en anglais
605         If Not IsNull(rstPieces.Fields("Desc_EN")) Then
610           itmPieces.SubItems(I_COL_DESCRIPTION) = rstPieces.Fields("Desc_EN")
615         Else
620           itmPieces.SubItems(I_COL_DESCRIPTION) = vbNullString
625         End If
630       End If
    
635       itmPieces.ListSubItems(I_COL_DESCRIPTION).ForeColor = lCouleur
    
          'Fabricant
640       If Not IsNull(rstPieces.Fields("Manufact")) Then
645         itmPieces.SubItems(I_COL_MANUFACTURIER) = rstPieces.Fields("Manufact")
650       Else
655         itmPieces.SubItems(I_COL_MANUFACTURIER) = vbNullString
660       End If
          
665       itmPieces.ListSubItems(I_COL_MANUFACTURIER).ForeColor = lCouleur
          
          'Fournisseur
670       If Not IsNull(rstPieces.Fields("IDFRS")) And rstPieces.Fields("IDFRS") > 0 Then
675         If itmPieces.SubItems(I_COL_NO_ITEM) <> "Texte" Then
680           Call rstFRS.Open("SELECT NomFournisseur FROM GRB_Fournisseur WHERE IDFRS = " & rstPieces.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
    
              'On affiche le nom dans la colonne
685           itmPieces.SubItems(I_COL_FOURNISSEUR) = rstFRS.Fields("NomFournisseur")
       
690           Call rstFRS.Close
695         End If
700       Else
705         itmPieces.SubItems(I_COL_FOURNISSEUR) = vbNullString
710       End If
          
715       itmPieces.ListSubItems(I_COL_FOURNISSEUR).ForeColor = lCouleur
                
720       If m_sType = "E" Then
725         Call rstInventaire.Open("SELECT * FROM GRB_InventaireElec WHERE NoItem = '" & Replace(rstPieces.Fields("NumItem"), "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
730       Else
735         Call rstInventaire.Open("SELECT * FROM GRB_InventaireMec WHERE NoItem = '" & Replace(rstPieces.Fields("NumItem"), "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
740       End If

745       If Not rstInventaire.EOF Then
750         itmPieces.SubItems(I_COL_QTE_STOCK) = rstInventaire.Fields("QuantitéStock")
755       End If

760       Call rstInventaire.Close
    
765       Call rstPieces.MoveNext
770     Loop
  
775     Call rstPieces.Close
780     Set rstPieces = Nothing

785     Set rstFRS = Nothing
790     Set rstInventaire = Nothing
795     Set rstSection = Nothing

800     Exit Sub

AfficherErreur:

805     woups "frmChoixBonCommande", "RemplirListViewPieces", Err, Erl
End Sub

Private Sub cmdAnnuler_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmChoixBonCommande", "cmdAnnuler_Click", Err, Erl
End Sub

Private Sub cmdCommander_Click()

5       On Error GoTo AfficherErreur

10      Dim bChecked  As Boolean
15      Dim iCompteur As Integer
20      Dim sNoBC     As String
25      Dim rstProjet As ADODB.Recordset
    
30      bChecked = False
  
35      For iCompteur = 1 To lvwPiece.ListItems.count
40        If lvwPiece.ListItems(iCompteur).Checked = True Then
45          bChecked = True
      
50          Exit For
55        End If
60      Next
  
65      If bChecked = True Then
70        Set m_collPiece = New Collection
75        Set m_collNoLigne = New Collection

80        If m_frmSource.Name <> "frmAchat" Then
85          Call ModifierFournisseurBD
90        End If
    
95        For iCompteur = 1 To lvwPiece.ListItems.count
100         If lvwPiece.ListItems(iCompteur).Checked = True Then
105           Call m_collPiece.Add(lvwPiece.ListItems(iCompteur).SubItems(I_COL_NO_ITEM))
110           Call m_collNoLigne.Add(lvwPiece.ListItems(iCompteur).ListSubItems(I_COL_NO_ITEM).Tag)
115         End If
120       Next
    
125       If m_frmSource.Name <> "frmAchat" Then
130         Set rstProjet = New ADODB.Recordset

135         If m_sType = "E" Then
140           Call rstProjet.Open("SELECT ProchaineCommande FROM GRB_ProjetElec WHERE IDProjet = '" & m_sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
145         Else
150           Call rstProjet.Open("SELECT ProchaineCommande FROM GRB_ProjetMec WHERE IDProjet = '" & m_sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
155         End If

160         If Not IsNull(rstProjet.Fields("ProchaineCommande")) Then
165           sNoBC = m_sNoProjet & "-" & Right$("00" & rstProjet.Fields("ProchaineCommande"), 3)
170         Else
175           sNoBC = m_sNoProjet
180         End If

185         Call rstProjet.Close
190         Set rstProjet = Nothing
195       Else
200         sNoBC = m_sIDAchat & "-" & Right$("00" & m_iIndexAchat, 3)
205       End If

210       If sNoBC <> vbNullString Then
215         If m_frmSource.Name = "FrmProjSoumElec" Then
220           Call frmBonCommande.AfficherFormProjetAchat(m_sNoProjet, sNoBC, m_collPiece, m_collNoLigne, I_PROJET_ELEC, m_eLangage)
225         Else
230           If m_frmSource.Name = "FrmProjSoumMec" Then
235             Call frmBonCommande.AfficherFormProjetAchat(m_sNoProjet, sNoBC, m_collPiece, m_collNoLigne, I_PROJET_MEC, m_eLangage)
240           Else
245             If m_sType = "E" Then
250               Call frmBonCommande.AfficherFormProjetAchat(m_sIDAchat & "-" & Right$("00" & m_iIndexAchat, 3), sNoBC, m_collPiece, m_collNoLigne, I_ACHAT_ELEC, 0)
255             Else
                
260               Call frmBonCommande.AfficherFormProjetAchat(m_sIDAchat & "-" & Right$("00" & m_iIndexAchat, 3), sNoBC, m_collPiece, m_collNoLigne, I_ACHAT_MEC, 0)
265             End If
270           End If
275         End If

280         Call Unload(Me)
285       End If
290     Else
295       Call MsgBox("Aucune pièce n'est sélectionnée!", vbOKOnly, "Erreur")
300     End If

305     Exit Sub

AfficherErreur:

310     woups "frmChoixBonCommande", "cmdCommander_Click", Err, Erl
End Sub

Private Sub cmdSelectAll_Click()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer

15      For iCompteur = 1 To lvwPiece.ListItems.count
20        If lvwPiece.ListItems(iCompteur).SubItems(I_COL_NO_ITEM) <> "Texte" And lvwPiece.ListItems(iCompteur).SubItems(I_COL_NO_ITEM) <> "Text" Then
25          If m_frmSource.Name <> "frmAchat" Then
30            If lvwPiece.ListItems(iCompteur).Tag <> vbNullString Then
35              If lvwPiece.ListItems(iCompteur).SubItems(I_COL_NO_ITEM) <> vbNullString Then
40                If CDbl(lvwPiece.ListItems(iCompteur).Text) > 0 Then
45                  lvwPiece.ListItems(iCompteur).Checked = True
50                End If
55              End If
60            End If
65          Else
70            If CDbl(lvwPiece.ListItems(iCompteur).Text) > 0 Then
75              lvwPiece.ListItems(iCompteur).Checked = True
80            End If
85          End If
90        End If
95      Next

100     Exit Sub

AfficherErreur:

105     woups "frmChoixBonCommande", "cmdSelectAll_Click", Err, Erl
End Sub

Private Sub cmdDeSelectAll_Click()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer

15      For iCompteur = 1 To lvwPiece.ListItems.count
20        lvwPiece.ListItems(iCompteur).Checked = False
25      Next

30      Exit Sub

AfficherErreur:

35      woups "frmChoixBonCommande", "cmdDeselectAll_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      Set m_collNoLigneFRS = New Collection
15      Set m_collEscompte = New Collection
20      Set m_collPrixList = New Collection
25      Set m_collPrixNet = New Collection
30      Set m_collPrixOrigine = New Collection
35      Set m_collPrixSP = New Collection

40      Exit Sub

AfficherErreur:

45      woups "frmChoixBonCommande", "Form_Load", Err, Erl
End Sub

Private Sub lvwPiece_DblClick()

5       On Error GoTo AfficherErreur

        
10      If m_frmSource.Name <> "frmAchat" Then
          'Si ce n'est pas une section
15        If lvwPiece.SelectedItem.Tag <> "" Then
            'Si ce n'est pas une sous-section
20          If lvwPiece.SelectedItem.SubItems(I_COL_NO_ITEM) <> "" Then
25            Call ChangerFournisseur
30          End If
35        End If
40      End If

45      Exit Sub

AfficherErreur:

50      woups "frmChoixBonCommande", "lvwPiece_DblClick", Err, Erl
End Sub

Private Sub lvwPiece_ItemCheck(ByVal Item As MSComctlLib.ListItem)

5       On Error GoTo AfficherErreur
  
10      If m_frmSource.Name <> "frmAchat" Then
          'Si c'est du texte
15        If Item.SubItems(I_COL_NO_ITEM) = "Texte" Or Item.Tag = vbNullString Or Item.SubItems(I_COL_NO_ITEM) = vbNullString Then
            'On enlève le check
20          Item.Checked = False
25        Else
30          If CDbl(Replace(Item.Text, "*", "")) <= 0 Then
35            'On enlève le check
40            Item.Checked = False
45          End If
50        End If
55      End If

60      Exit Sub

AfficherErreur:

65      woups "frmChoixBonCommande", "lvwPiece_ItemCheck", Err, Erl
End Sub

Private Sub ChangerFournisseur()

5       On Error GoTo AfficherErreur

10      Call AfficherListeFournisseurs

15      If lvwfournisseur.ListItems.count = 0 Then
20        Call MsgBox("Il n'y a aucun fournisseur pour cette pièce!", vbOKOnly, "Erreur")
 
25        Exit Sub
30      Else
35        frafournisseur.Visible = True
40      End If

45      Exit Sub

AfficherErreur:

50      woups "frmChoixBonCommande", "ChangerFournisseur", Err, Erl
End Sub

Private Sub AfficherListeFournisseurs()

5       On Error GoTo AfficherErreur

        'Méthode qui sert à afficher la liste des fournisseurs
        'Affiche le frame seulement s'il y a des items dans le ListView
10      Call RemplirListViewFournisseur
  
15      If lvwfournisseur.ListItems.count > 1 Then
25        frafournisseur.Visible = True
30        Call lvwfournisseur.SetFocus
35      End If

40      Exit Sub

AfficherErreur:

45      woups "frmChoixBonCommande", "AfficherListeFournisseurs", Err, Erl
End Sub

Private Sub lvwFournisseur_DblClick()

5       On Error GoTo AfficherErreur

10      Call ChoisirFournisseur

15      Exit Sub

AfficherErreur:

20      woups "frmChoixBonCommande", "lvwFournisseur_DblClick", Err, Erl
End Sub

Private Sub ChoisirFournisseur()

5       On Error GoTo AfficherErreur

10      Dim itmBC  As ListItem
15      Dim itmFRS As ListItem

20      Set itmBC = lvwPiece.SelectedItem
25      Set itmFRS = lvwfournisseur.SelectedItem

30      itmBC.SubItems(I_COL_FOURNISSEUR) = itmFRS.Text

35      itmBC.ListSubItems(I_COL_FOURNISSEUR).Tag = itmFRS.Tag

40      Call m_collNoLigneFRS.Add(itmBC.ListSubItems(I_COL_NO_ITEM).Tag)
45      Call m_collEscompte.Add(itmFRS.SubItems(I_COL_FRS_ESCOMPTE))
50      Call m_collPrixList.Add(itmFRS.SubItems(I_COL_FRS_PRIX_LIST))
55      Call m_collPrixNet.Add(itmFRS.SubItems(I_COL_FRS_PRIX_NET))
60      Call m_collPrixOrigine.Add(itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).Tag)
65      Call m_collPrixSP.Add(itmFRS.SubItems(I_COL_FRS_PRIX_SP))

        'On cache le listview
70      frafournisseur.Visible = False

75      Exit Sub

AfficherErreur:

80      woups "frmChoixBonCommande", "ChoisirFournisseur", Err, Erl
End Sub

Private Sub cmdOKFRS_Click()

5       On Error GoTo AfficherErreur

10      Call ChoisirFournisseur

15      Exit Sub

AfficherErreur:

20      woups "frmChoixBonCommande", "cmdOKFRS_Click", Err, Erl
End Sub

Private Sub cmdAnnulerFRS_Click()

5       On Error GoTo AfficherErreur

10      frafournisseur.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmChoixBonCommande", "cmdAnnulerFRS_Click", Err, Erl
End Sub

Private Sub RemplirListViewFournisseur()

5       On Error GoTo AfficherErreur

        'Rempli le listview des distributeur pour une pièce choisie
10      Dim rstPieceFRS As ADODB.Recordset
15      Dim rstContact  As ADODB.Recordset
20      Dim rstFRS      As Recordset
25      Dim iCompteur   As Integer
30      Dim itmFRS      As ListItem
35      Dim iNoClient   As Integer
40      Dim sDevise     As String
  
        'vide le lister
45      Call lvwfournisseur.ListItems.Clear

50      Set rstPieceFRS = New ADODB.Recordset
55      Set rstContact = New ADODB.Recordset
60      Set rstFRS = New ADODB.Recordset

65      Call rstFRS.Open("SELECT IDFRS FROM GRB_Fournisseur WHERE NomFournisseur = 'FOURNI PAR LE CLIENT'", g_connData, adOpenDynamic, adLockOptimistic)
      
70      iNoClient = rstFRS.Fields("IDFRS")

75      Call rstFRS.Close
80      Set rstFRS = Nothing
      
85      Call rstPieceFRS.Open("SELECT GRB_PiecesFRS.*, GRB_Fournisseur.NomFournisseur FROM GRB_PiecesFRS INNER JOIN GRB_Fournisseur ON GRB_PiecesFRS.IDFRS = GRB_Fournisseur.IDFRS WHERE PIECE = '" & Replace(lvwPiece.SelectedItem.SubItems(I_COL_NO_ITEM), "'", "''") & "' AND Type = '" & m_sType & "' ORDER BY PrixReel", g_connData, adOpenDynamic, adLockOptimistic)
  
        'tant il y a des fournisseur de la piece , ajoute dans lister
90      Do While Not rstPieceFRS.EOF
          'on change la couleur de l'enregistrement selon la devise monétaire.
          'CAN = rouge, USA ou ESP = bleu
95        If rstPieceFRS.Fields("DeviseMonétaire") = "CAN" Then
100         sDevise = "CAN"
105       Else
110         If rstPieceFRS.Fields("DeviseMonétaire") = "USA" Then
115           sDevise = "USA"
120         Else
125           sDevise = "SPA"
130         End If
135       End If
     
140       Set itmFRS = lvwfournisseur.ListItems.Add
       
145       itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = " "
150       itmFRS.SubItems(I_COL_FRS_ESCOMPTE) = " "
155       itmFRS.SubItems(I_COL_FRS_PRIX_NET) = " "
160       itmFRS.SubItems(I_COL_FRS_PRIX_SP) = " "
       
          'Nom du FRS
165       itmFRS.Text = rstPieceFRS.Fields("NomFournisseur")
           
170       itmFRS.Tag = rstPieceFRS.Fields("IDFRS")
      
          'Personne ressource
175       If Trim(rstPieceFRS.Fields("PERS_RESS")) <> vbNullString Then
180         Call rstContact.Open("SELECT NomContact FROM GRB_Contact WHERE IDContact = " & rstPieceFRS.Fields("PERS_RESS"), g_connData, adOpenDynamic, adLockOptimistic)
        
185         If Not rstContact.EOF Then
190           itmFRS.SubItems(I_COL_FRS_PERS_RESS) = rstContact.Fields("NomContact")
195         End If

200         Call rstContact.Close
205       End If
                     
          'Date
210       If Not IsNull(rstPieceFRS.Fields("Date")) Then
215         itmFRS.SubItems(I_COL_FRS_DATE) = rstPieceFRS.Fields("Date")
220       Else
225         itmFRS.SubItems(I_COL_FRS_DATE) = vbNullString
230       End If
                          
          'Entrer par
235       If Not IsNull(rstPieceFRS.Fields("Entrer_Par")) Then
240         itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = rstPieceFRS.Fields("Entrer_Par")
245       Else
250         itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = vbNullString
255       End If
                                 
          'Valide
260       If Not IsNull(rstPieceFRS.Fields("Valide")) Then
265         itmFRS.SubItems(I_COL_FRS_VALIDE) = rstPieceFRS.Fields("Valide")
270       Else
275         itmFRS.SubItems(I_COL_FRS_VALIDE) = vbNullString
280       End If
                             
          'Prix listé
285       If rstPieceFRS.Fields("PRIX_LIST") <> vbNullString Then
290         If sDevise = "USA" Then
295           itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = Conversion(CStr(Round(CDbl(rstPieceFRS.Fields("PRIX_LIST")), 4)), MODE_ARGENT, 4)
300         Else
305           If sDevise = "SPA" Then
310             itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = Conversion(CStr(Round(CDbl(rstPieceFRS.Fields("PRIX_LIST")), 4)), MODE_ARGENT, 4)
315           Else
320             itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = Conversion(Round(Replace(rstPieceFRS.Fields("PRIX_LIST"), ".", ","), 4), MODE_ARGENT, 4)
325           End If
330         End If
335       End If
                             
          'Escompte
340       If rstPieceFRS.Fields("ESCOMPTE") <> vbNullString Then
345         itmFRS.SubItems(I_COL_FRS_ESCOMPTE) = Conversion(Replace(Replace(rstPieceFRS.Fields("ESCOMPTE"), "_", vbNullString), ".", ",") * 100, MODE_POURCENT)
350       End If
   
          'Prix net
355       If rstPieceFRS.Fields("PRIX_NET") <> vbNullString Then
360         If sDevise = "USA" Then
365           itmFRS.SubItems(I_COL_FRS_PRIX_NET) = Conversion(CStr(Round(CDbl(rstPieceFRS.Fields("PRIX_NET")), 4)), MODE_ARGENT, 4)
370         Else
375           If sDevise = "SPA" Then
380             itmFRS.SubItems(I_COL_FRS_PRIX_NET) = Conversion(CStr(Round(CDbl(rstPieceFRS.Fields("PRIX_NET")), 4)), MODE_ARGENT, 4)
385           Else
390             itmFRS.SubItems(I_COL_FRS_PRIX_NET) = Conversion(Round(Replace(rstPieceFRS.Fields("PRIX_NET"), ".", ","), 4), MODE_ARGENT, 4)
395           End If
400         End If
405       End If
      
          'Prix spécial
410       If rstPieceFRS.Fields("PRIX_SP") <> vbNullString Then
415         If sDevise = "USA" Then
420           itmFRS.SubItems(I_COL_FRS_PRIX_SP) = Conversion(CStr(Round(CDbl(rstPieceFRS.Fields("PRIX_SP")), 4)), MODE_ARGENT, 4)
425         Else
430           If sDevise = "SPA" Then
435             itmFRS.SubItems(I_COL_FRS_PRIX_SP) = Conversion(Round(CDbl(rstPieceFRS.Fields("PRIX_SP")), 4), MODE_ARGENT, 4)
440           Else
445             itmFRS.SubItems(I_COL_FRS_PRIX_SP) = Conversion(Round(rstPieceFRS.Fields("PRIX_SP"), 4), MODE_ARGENT, 4)
450           End If
455         End If
460       End If
       
          'Quoter
465       If rstPieceFRS.Fields("QUOTER") = True Then
470         itmFRS.SubItems(I_COL_FRS_QUOTER) = "Oui"
475       Else
480         itmFRS.SubItems(I_COL_FRS_QUOTER) = "Non"
485       End If
        
          'Pour garder en mémoire le prix d'origine, je le mets dans le
          'tag de la colonne Prix Listé
490       If itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = vbNullString Then
495         itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = " "
500       End If
    
505       If rstPieceFRS.Fields("PRIX_NET") <> vbNullString Then
510         itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).Tag = Replace(rstPieceFRS.Fields("PRIX_NET"), ".", ",")
515       Else
520         itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).Tag = Replace(rstPieceFRS.Fields("PRIX_SP"), ".", ",")
525       End If

530       'Pour avoir le no d'enregistrement de PiecesFRS

535       If itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = vbNullString Then
540         itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = " "
545       End If

550       itmFRS.ListSubItems(I_COL_FRS_ENTRER_PAR).Tag = rstPieceFRS.Fields("NoEnreg")
 
555       Call rstPieceFRS.MoveNext
560     Loop
    
        'ferme la table
565     Call rstPieceFRS.Close
570     Set rstPieceFRS = Nothing

575     Exit Sub

AfficherErreur:

580     woups "frmChoixBonCommande", "RemplirListViewFournisseur", Err, Erl
End Sub

Private Sub CalculerTotalRecordsetElec(ByVal sNoProjSoum As String)

5       On Error GoTo AfficherErreur

        'Méthode pour calculer le prix
10      Dim rstProjet             As ADODB.Recordset
15      Dim rstPiece              As ADODB.Recordset
20      Dim rstPunch              As ADODB.Recordset
25      Dim dblTotalDessin        As Double
30      Dim dblTotalFabrication   As Double
35      Dim dblTotalAssemblage    As Double
40      Dim dblTotalProgInterface As Double
45      Dim dblTotalProgAutomate  As Double
50      Dim dblTotalProgRobot     As Double
55      Dim dblTotalVision        As Double
60      Dim dblTotalTest          As Double
65      Dim dblTotalInstallation  As Double
70      Dim dblTotalMiseService   As Double
75      Dim dblTotalFormation     As Double
80      Dim dblTotalGestion       As Double
85      Dim dblTotalShipping      As Double
90      Dim dblHebergement        As Double
95      Dim dblRepas              As Double
100     Dim dblTransport          As Double
105     Dim dblUniteMobile        As Double
110     Dim dblPrixEmballage      As Double
115     Dim dblTotalResteTemps    As Double
120     Dim dblPrixPieces         As Double
125     Dim dblPrixTotal          As Double
130     Dim dblCommission         As Double
135     Dim dblTotalTemps         As Double
140     Dim dblProfit             As Double
145     Dim dblTotalManuel        As Double
150     Dim dblTotalPieceImprevue As Double
155     Dim dblGrandTotal         As Double
160     Dim sDateDebut            As String
165     Dim sDateFin              As String
170     Dim sTotal                As String

175     Set rstProjet = New ADODB.Recordset

180     Call rstProjet.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)

185     If Not rstProjet.EOF Then
190       Set rstPunch = New ADODB.Recordset

          'Total des temps
195       sDateDebut = "TIMESERIAL(Left(GRB_Punch.HeureDébut,2),RIGHT(GRB_Punch.HeureDébut,2),0)"

200       sDateFin = "TIMESERIAL(Left(GRB_Punch.HeureFin,2),RIGHT(GRB_Punch.HeureFin,2),0)"

205       sTotal = "(SUM(" & sDateFin & " - " & sDateDebut & ")* 24) As Total"


210       Call rstPunch.Open("SELECT Type, " & sTotal & " FROM GRB_Punch WHERE NoProjet = '" & sNoProjSoum & "' And HeureFin is not null AND HeureDébut is not null GROUP BY Type", g_connData, adOpenDynamic, adLockOptimistic)

215       dblTotalDessin = 0
220       dblTotalFabrication = 0
225       dblTotalAssemblage = 0
230       dblTotalProgInterface = 0
235       dblTotalProgAutomate = 0
240       dblTotalProgRobot = 0
245       dblTotalVision = 0
250       dblTotalTest = 0
255       dblTotalInstallation = 0
260       dblTotalMiseService = 0
265       dblTotalFormation = 0
270       dblTotalGestion = 0
275       dblTotalShipping = 0

280       Do While Not rstPunch.EOF
285         If Not IsNull(rstPunch.Fields("Total")) Then
290           Select Case rstPunch.Fields("Type")
                Case "Dessin":        dblTotalDessin = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxDessin"))
295             Case "Fabrication":   dblTotalFabrication = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxFabrication"))
300             Case "Assemblage":    dblTotalAssemblage = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxAssemblage"))
305             Case "ProgInterface": dblTotalProgInterface = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxProgInterface"))
310             Case "ProgAutomate":  dblTotalProgAutomate = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxProgAutomate"))
315             Case "ProgRobot":     dblTotalProgRobot = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxProgRobot"))
320             Case "Vision":        dblTotalVision = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxVision"))
325             Case "Test":          dblTotalTest = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxTest"))
330             Case "Installation":  dblTotalInstallation = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxInstallation"))
335             Case "MiseService":   dblTotalMiseService = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxMiseService"))
340             Case "Formation":     dblTotalFormation = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxFormation"))
345             Case "Gestion":       dblTotalGestion = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxGestion"))
350             Case "Shipping":      dblTotalShipping = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxShipping"))
355           End Select
360         End If
              
365         Call rstPunch.MoveNext
370       Loop

375       Call rstPunch.Close
380       Set rstPunch = Nothing

385       dblTotalTemps = dblTotalDessin + _
                          dblTotalFabrication + _
                          dblTotalAssemblage + _
                          dblTotalProgInterface + _
                          dblTotalProgAutomate + _
                          dblTotalProgRobot + _
                          dblTotalVision + _
                          dblTotalTest + _
                          dblTotalInstallation + _
                          dblTotalMiseService + _
                          dblTotalFormation + _
                          dblTotalGestion + _
                          dblTotalShipping
                          
390       Set rstPiece = New ADODB.Recordset

395       Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjSoum & "' AND Type = 'E'", g_connData, adOpenDynamic, adLockOptimistic)

          'Pour chaque élément du recordset
400       Do While Not rstPiece.EOF
405         If Trim(rstPiece.Fields("Prix_total")) <> vbNullString Then
              'On additionne le prix total
410           dblPrixPieces = dblPrixPieces + CDbl(rstPiece.Fields("Prix_total")) - CDbl(rstPiece.Fields("Profit_Argent"))

              'On additionne le profit
415           dblProfit = dblProfit + CDbl(rstPiece.Fields("Profit_Argent"))
420         End If

425         Call rstPiece.MoveNext
430       Loop

435       Call rstPiece.Close
440       Set rstPiece = Nothing

445       dblHebergement = 0
450       dblRepas = 0
455       dblTransport = 0
460       dblUniteMobile = 0

          'Correction d'un bug de Type Incompatible
465       If IsNumeric(rstProjet.Fields("PrixEmballage")) Then
470         dblPrixEmballage = CDbl(rstProjet.Fields("PrixEmballage"))
475       Else
480         dblPrixEmballage = 0
485       End If
     
490       dblTotalResteTemps = dblHebergement + dblRepas + dblTransport + dblUniteMobile + dblPrixEmballage

495       If IsNumeric(rstProjet.Fields("total_manuel")) Then
500         dblTotalManuel = CDbl(rstProjet.Fields("total_manuel"))
505       Else
510         dblTotalManuel = 0
515       End If

520       dblTotalPieceImprevue = (dblPrixPieces + dblProfit) * (1 + CDbl(rstProjet.Fields("Imprevue")))

525       dblPrixTotal = dblTotalTemps + dblTotalManuel + dblTotalPieceImprevue + dblTotalResteTemps

          'Le calcul de la commission est sur les manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps
530       dblCommission = dblPrixTotal * CDbl(rstProjet.Fields("Commission"))

          'Le prix total est le calcul des manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps + total de la commission
535       dblGrandTotal = dblPrixTotal + dblCommission

          'Format monétaires avec 2 chiffres après la virgule
540       rstProjet.Fields("total_commission") = dblCommission
545       rstProjet.Fields("Total_manuel") = dblTotalManuel
550       rstProjet.Fields("Total_temps") = dblTotalTemps
555       rstProjet.Fields("total_imprevue") = dblTotalPieceImprevue - (dblPrixPieces + dblProfit)
560       rstProjet.Fields("total_piece") = dblPrixPieces
565       rstProjet.Fields("Total_Commission") = Conversion(CStr(Round(dblCommission, 2)), MODE_ARGENT)
570       rstProjet.Fields("PrixTotal") = Conversion(CStr(Round(dblGrandTotal, 2)), MODE_ARGENT)
575       rstProjet.Fields("Total_Profit") = Conversion(CStr(Round(dblProfit, 2)), MODE_ARGENT)

580       Call rstProjet.Update
585     Else
590       Call MsgBox("Le projet " & sNoProjSoum & " est inexistant!", vbOKOnly, "Erreur")
595     End If

600     Call rstProjet.Close
605     Set rstProjet = Nothing

610     Exit Sub

AfficherErreur:

615     woups "frmChoixBonCommande", "CalculerTotalRecordset", Err, Erl
End Sub

Private Sub CalculerTotalRecordsetMec(ByVal sNoProjSoum As String)

5       On Error GoTo AfficherErreur

        'Méthode pour calculer le prix
10      Dim rstProjet            As ADODB.Recordset
15      Dim rstPiece             As ADODB.Recordset
20      Dim rstPunch             As ADODB.Recordset
25      Dim dblPrixPieces        As Double
30      Dim dblPrixTotal         As Double
35      Dim dblCommission        As Double
40      Dim dblTotalTemps        As Double
45      Dim dblProfit            As Double
50      Dim dblTotalManuel       As Double
55      Dim dblTotalImprevue     As Double
60      Dim dblGrandTotal        As Double
65      Dim dblTotalDessin       As Double
70      Dim dblTotalCoupe        As Double
75      Dim dblTotalMachinage    As Double
80      Dim dblTotalSoudure      As Double
85      Dim dblTotalAssemblage   As Double
90      Dim dblTotalPeinture     As Double
95      Dim dblTotalTest         As Double
100     Dim dblTotalInstallation As Double
105     Dim dblTotalFormation    As Double
110     Dim dblTotalGestion      As Double
115     Dim dblTotalShipping     As Double
120     Dim dblPrixEmballage     As Double
125     Dim dblTotalResteTemps   As Double
130     Dim sDateDebut           As String
135     Dim sDateFin             As String
140     Dim sTotal               As String

145     Set rstProjet = New ADODB.Recordset

150     Call rstProjet.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)

155     If Not rstProjet.EOF Then
160       Set rstPiece = New ADODB.Recordset

165       Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjSoum & "' AND Type = 'M'", g_connData, adOpenDynamic, adLockOptimistic)

          'Pour chaque élément du recordset
170       Do While Not rstPiece.EOF
175         If Trim(rstPiece.Fields("Prix_total")) <> vbNullString Then
              'On additionne le prix total
180           dblPrixPieces = dblPrixPieces + rstPiece.Fields("Prix_total") - rstPiece.Fields("Profit_Argent")
          
              'On additionne le profit
185           dblProfit = dblProfit + rstPiece.Fields("Profit_Argent")
190         End If

195         Call rstPiece.MoveNext
200       Loop
    
          'Total des temps
205       sDateDebut = "TIMESERIAL(Left(GRB_Punch.HeureDébut,2),RIGHT(GRB_Punch.HeureDébut,2),0)"

210       sDateFin = "TIMESERIAL(Left(GRB_Punch.HeureFin,2),RIGHT(GRB_Punch.HeureFin,2),0)"

215       sTotal = "(SUM(" & sDateFin & " - " & sDateDebut & ")* 24) As Total"

220       Set rstPunch = New ADODB.Recordset

225       Call rstPunch.Open("SELECT Type, " & sTotal & " FROM GRB_Punch WHERE NoProjet = '" & sNoProjSoum & "' And HeureFin is not null AND HeureDébut is not null GROUP BY Type", g_connData, adOpenDynamic, adLockOptimistic)
          
230       dblTotalDessin = 0
235       dblTotalCoupe = 0
240       dblTotalMachinage = 0
245       dblTotalSoudure = 0
250       dblTotalAssemblage = 0
255       dblTotalPeinture = 0
260       dblTotalTest = 0
265       dblTotalInstallation = 0
270       dblTotalFormation = 0
275       dblTotalGestion = 0
280       dblTotalShipping = 0

285       Do While Not rstPunch.EOF
290         If Not IsNull(rstPunch.Fields("Total")) Then
295           Select Case rstPunch.Fields("Type")
                Case "Dessin":       dblTotalDessin = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxDessin"))
300             Case "Coupe":        dblTotalCoupe = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxCoupe"))
305             Case "Machinage":    dblTotalMachinage = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxMachinage"))
310             Case "Soudure":      dblTotalSoudure = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxSoudure"))
315             Case "Assemblage":   dblTotalAssemblage = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxAssemblage"))
320             Case "Peinture":     dblTotalPeinture = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxPeinture"))
325             Case "Test":         dblTotalTest = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxTest"))
330             Case "Installation": dblTotalInstallation = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxInstallation"))
335             Case "Formation":    dblTotalFormation = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxFormation"))
340             Case "Gestion":      dblTotalGestion = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxGestion"))
345             Case "Shipping":     dblTotalShipping = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjet.Fields("TauxShipping"))
350           End Select
355         End If

360         Call rstPunch.MoveNext
365       Loop

370       Call rstPunch.Close
375       Set rstPunch = Nothing
                 
380       dblTotalTemps = dblTotalDessin + _
                          dblTotalCoupe + _
                          dblTotalMachinage + _
                          dblTotalSoudure + _
                          dblTotalAssemblage + _
                          dblTotalPeinture + _
                          dblTotalTest + _
                          dblTotalInstallation + _
                          dblTotalFormation + _
                          dblTotalGestion + _
                          dblTotalShipping
                        
          'Correction d'un bug de Type Incompatible
385       If IsNumeric(rstProjet.Fields("PrixEmballage")) Then
390         dblPrixEmballage = CDbl(rstProjet.Fields("PrixEmballage"))
395       Else
400         dblPrixEmballage = 0
405       End If
      
410       dblTotalResteTemps = dblPrixEmballage
                                                              
415       If IsNumeric(rstProjet.Fields("total_manuel")) Then
420         dblTotalManuel = CDbl(rstProjet.Fields("total_manuel"))
425       Else
430         dblTotalManuel = 0
435       End If
                        
440       dblTotalImprevue = Round((dblPrixPieces + dblProfit) * CDbl(rstProjet.Fields("Imprevue")), 2)
   
445       dblPrixTotal = dblPrixPieces + dblProfit + dblTotalTemps + dblTotalImprevue + dblTotalManuel + dblTotalResteTemps
                          
          'Le calcul de la commission est sur les manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps
450       dblCommission = Round(dblPrixTotal * CDbl(rstProjet.Fields("Commission")), 2)
        
          'Le prix total est le calcul des manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps + total de la commission
455       dblGrandTotal = dblPrixTotal + dblCommission

          'Format monétaires avec 2 chiffres après la virgule
460       rstProjet.Fields("Total_Piece") = Conversion(CStr(Round(dblPrixPieces, 2)), MODE_ARGENT)
465       rstProjet.Fields("Total_Temps") = Conversion(CStr(Round(dblTotalTemps, 2)), MODE_ARGENT)
470       rstProjet.Fields("PrixTotal") = Conversion(CStr(Round(dblGrandTotal, 2)), MODE_ARGENT)
475       rstProjet.Fields("total_Imprevue") = Conversion(CStr(Round(dblTotalImprevue, 2)), MODE_ARGENT)
480       rstProjet.Fields("total_commission") = Conversion(CStr(Round(dblCommission, 2)), MODE_ARGENT)
485       rstProjet.Fields("total_profit") = Conversion(CStr(Round(dblProfit, 2)), MODE_ARGENT)

490       Call rstProjet.Update

495       Call rstPiece.Close
500       Set rstPiece = Nothing
505     Else
510       Call MsgBox("Le projet " & sNoProjSoum & " est inexistant!", vbOKOnly, "Erreur")
515     End If

520     Call rstProjet.Close
525     Set rstProjet = Nothing

530     Exit Sub

AfficherErreur:

535     woups "frmChoixBonCommande", "CalculerTotalRecordset", Err, Erl
End Sub

Private Sub ModifierFournisseurBD()

5       On Error GoTo AfficherErreur

10      Dim rstPiece      As ADODB.Recordset
15      Dim rstProjet     As ADODB.Recordset
20      Dim sProfit       As String
25      Dim iCompteur     As Integer
30      Dim bModif        As Boolean
35      Dim iCompteurColl As Integer
40      Dim iIndexColl    As Integer
45      Dim sLiaison      As String

50      Set rstProjet = New ADODB.Recordset

55      If m_sType = "E" Then
60        Call rstProjet.Open("SELECT Profit, LiaisonChargeable FROM GRB_ProjetElec WHERE IDProjet = '" & m_sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
65      Else
70        Call rstProjet.Open("SELECT Profit, LiaisonChargeable FROM GRB_ProjetMec WHERE IDProjet = '" & m_sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
75      End If

80      sProfit = rstProjet.Fields("Profit")

85      If Not IsNull(rstProjet.Fields("LiaisonChargeable")) Then
90        sLiaison = rstProjet.Fields("LiaisonChargeable")
95      Else
100       sLiaison = ""
105     End If

110     Call rstProjet.Close
115     Set rstProjet = Nothing

120     Set rstPiece = New ADODB.Recordset

125     For iCompteur = 1 To lvwPiece.ListItems.count
130       If lvwPiece.ListItems(iCompteur).Checked = True Then
135         If lvwPiece.ListItems(iCompteur).ListSubItems(I_COL_FOURNISSEUR).Tag <> "" Then
140           bModif = True

145           Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & m_sNoProjet & "' AND NuméroLigne = " & lvwPiece.ListItems(iCompteur).ListSubItems(I_COL_NO_ITEM).Tag, g_connData, adOpenDynamic, adLockOptimistic)

150           For iCompteurColl = 1 To m_collNoLigneFRS.count
155             If m_collNoLigneFRS(iCompteurColl) = lvwPiece.ListItems(iCompteur).ListSubItems(I_COL_NO_ITEM).Tag Then
160               iIndexColl = iCompteurColl

165               Exit For
170             End If
175           Next

180           rstPiece.Fields("IDFRS") = lvwPiece.ListItems(iCompteur).ListSubItems(I_COL_FOURNISSEUR).Tag

              'Prix listé
185           If Trim$(m_collPrixList(iIndexColl)) = vbNullString Then
190             rstPiece.Fields("PRIX_LIST") = Conversion("0", MODE_PAS_FORMAT, 4)
195           Else
200             rstPiece.Fields("PRIX_LIST") = Conversion(m_collPrixList(iIndexColl), MODE_PAS_FORMAT, 4)
205             rstPiece.Fields("PrixOrigine") = Conversion(m_collPrixOrigine(iIndexColl), MODE_PAS_FORMAT, 4)
210           End If
      
              'S'il y a un prix net, on le met l'escompte et le prix net sinon, on prend le prix
              'spécial pour mettre dans le prix net
215           If Trim$(m_collPrixNet(iIndexColl)) <> vbNullString Then
220             rstPiece.Fields("ESCOMPTE") = Conversion(m_collEscompte(iIndexColl), MODE_PAS_FORMAT)
225             rstPiece.Fields("PRIX_NET") = Conversion(m_collPrixNet(iIndexColl), MODE_PAS_FORMAT, 4)
230           Else
235             If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP)) <> vbNullString Then
240               rstPiece.Fields("PRIX_NET") = Conversion(m_collPrixSP(iIndexColl), MODE_PAS_FORMAT, 4)
245             Else
250               rstPiece.Fields("ESCOMPTE") = Conversion("0", MODE_PAS_FORMAT)
255               rstPiece.Fields("PRIX_NET") = Conversion("0", MODE_PAS_FORMAT, 4)
260             End If
265           End If
     
              'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
270           rstPiece.Fields("Prix_Total") = Conversion(CStr(Round(Replace(rstPiece.Fields("Qté"), "*", vbNullString) * rstPiece.Fields("PRIX_NET") * CSng(sProfit), 2)), MODE_PAS_FORMAT)
           
              'Pour le profit, c'est le prix total - (prix net * quantité)
275           rstPiece.Fields("Profit_argent") = Conversion(CStr(Round(rstPiece.Fields("Prix_Total") - (rstPiece.Fields("PRIX_NET") * Replace(rstPiece.Fields("Qté"), "*", vbNullString)), 2)), MODE_PAS_FORMAT)

280           Call rstPiece.Update

285           Call rstPiece.Close

290           If sLiaison <> "" Then
295             Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & Left$(m_sNoProjet, Len(m_sNoProjet) - 2) & sLiaison & "' AND Provenance = '" & Right$(m_sNoProjet, 2) & "' AND NumItem = '" & lvwPiece.ListItems(iCompteur).SubItems(I_COL_NO_ITEM) & "' AND Qté = '" & lvwPiece.ListItems(iCompteur).Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

300             For iCompteurColl = 1 To m_collNoLigneFRS.count
305               If m_collNoLigneFRS(iCompteurColl) = lvwPiece.ListItems(iCompteur).ListSubItems(I_COL_NO_ITEM).Tag Then
310                 iIndexColl = iCompteurColl

315                 Exit For
320               End If
325             Next

330             rstPiece.Fields("IDFRS") = lvwPiece.ListItems(iCompteur).ListSubItems(I_COL_FOURNISSEUR).Tag

                'Prix listé
335             If Trim$(m_collPrixList(iIndexColl)) = vbNullString Then
340               rstPiece.Fields("PRIX_LIST") = Conversion("0", MODE_PAS_FORMAT, 4)
345             Else
350               rstPiece.Fields("PRIX_LIST") = Conversion(m_collPrixList(iIndexColl), MODE_PAS_FORMAT, 4)
355               rstPiece.Fields("PrixOrigine") = Conversion(m_collPrixOrigine(iIndexColl), MODE_PAS_FORMAT, 4)
360             End If
      
                'S'il y a un prix net, on le met l'escompte et le prix net sinon, on prend le prix
                'spécial pour mettre dans le prix net
365             If Trim$(m_collPrixNet(iIndexColl)) <> vbNullString Then
370               rstPiece.Fields("ESCOMPTE") = Conversion(m_collEscompte(iIndexColl), MODE_PAS_FORMAT)
375               rstPiece.Fields("PRIX_NET") = Conversion(m_collPrixNet(iIndexColl), MODE_PAS_FORMAT, 4)
380             Else
385               If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP)) <> vbNullString Then
390                 rstPiece.Fields("PRIX_NET") = Conversion(m_collPrixSP(iIndexColl), MODE_PAS_FORMAT, 4)
395               Else
400                 rstPiece.Fields("ESCOMPTE") = Conversion("0", MODE_PAS_FORMAT)
405                 rstPiece.Fields("PRIX_NET") = Conversion("0", MODE_PAS_FORMAT, 4)
410               End If
415             End If
     
                'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
420             rstPiece.Fields("Prix_Total") = Conversion(CStr(Round(Replace(rstPiece.Fields("Qté"), "*", vbNullString) * rstPiece.Fields("PRIX_NET") * CSng(sProfit), 2)), MODE_PAS_FORMAT)
           
                'Pour le profit, c'est le prix total - (prix net * quantité)
425             rstPiece.Fields("Profit_argent") = Conversion(CStr(Round(rstPiece.Fields("Prix_Total") - (rstPiece.Fields("PRIX_NET") * Replace(rstPiece.Fields("Qté"), "*", vbNullString)), 2)), MODE_PAS_FORMAT)

430             Call rstPiece.Update

435             Call rstPiece.Close
440           End If
445         End If
450       End If
455     Next

460     Set rstPiece = Nothing

465     If bModif = True Then
470       If m_sType = "E" Then
475         Call CalculerTotalRecordsetElec(m_sNoProjet)

480         If sLiaison <> "" Then
485           Call CalculerTotalRecordsetElec(Left$(m_sNoProjet, Len(m_sNoProjet) - 2) & sLiaison)
490         End If

495         FrmProjSoumElec.m_bModifFournisseurBC = True
500       Else
505         Call CalculerTotalRecordsetMec(m_sNoProjet)

510         If sLiaison <> "" Then
515           Call CalculerTotalRecordsetMec(Left$(m_sNoProjet, Len(m_sNoProjet) - 2) & sLiaison)
520         End If

525         FrmProjSoumMec.m_bModifFournisseurBC = True
530       End If
535     Else
540       If m_sType = "E" Then
545         FrmProjSoumElec.m_bModifFournisseurBC = False
550       Else
555         FrmProjSoumMec.m_bModifFournisseurBC = False
560       End If
565     End If

570     Exit Sub

AfficherErreur:

575     woups "frmChoixBonCommande", "ModifierFournisseurBD", Err, Erl
End Sub
