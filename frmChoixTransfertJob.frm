VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChoixTransfertJob 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choix des pièces à transférer dans le projet"
   ClientHeight    =   9315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmChoixTransfertJob.frx":0000
   ScaleHeight     =   9315
   ScaleWidth      =   12315
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDeselectAll 
      Caption         =   "Aucun"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   8760
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Tous"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   8760
      Width           =   1095
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   9600
      TabIndex        =   3
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton cmdCreer 
      Caption         =   "Créer le projet"
      Height          =   375
      Left            =   10920
      TabIndex        =   4
      Top             =   8760
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwPiece 
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   12938
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
      NumItems        =   5
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
   End
End
Attribute VB_Name = "frmChoixTransfertJob"
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

Private m_sNoSoumission As String
Private m_sType     As String

Public Sub Afficher(ByVal sNoSoumission As String, ByVal sType As String)

5       On Error GoTo AfficherErreur

        'Méthode pour afficher le form
10      m_sNoSoumission = sNoSoumission

15      m_sType = sType
  
20      Call RemplirListViewPieces
  
25      Call Me.Show(vbModal)

30      Exit Sub

AfficherErreur:

35      woups "frmChoixTransfertJob", "Afficher", Err, Erl
End Sub

Private Sub RemplirListViewPieces()

5       On Error GoTo AfficherErreur

        'Rempli le ListView selon le no. du projet
10      Dim rstPieces     As ADODB.Recordset
15      Dim rstSection    As ADODB.Recordset
20      Dim rstFRS        As Recordset
25      Dim itmPieces     As ListItem
30      Dim bPremierEnr   As Boolean
35      Dim iOrdreSection As Integer
40      Dim sSousSection  As String
  
45      bPremierEnr = True
   
50      lvwPiece.Sorted = False

55      Set rstFRS = New ADODB.Recordset
60      Set rstPieces = New ADODB.Recordset
65      Set rstSection = New ADODB.Recordset

        'Ouverture du recordset
70      Call rstPieces.Open("SELECT * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & m_sNoSoumission & "' AND Type = '" & m_sType & "' AND PieceExtraChargeable = False AND PieceExtraNonChargeable = False ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
  
75      Do While Not rstPieces.EOF
80        Set itmPieces = lvwPiece.ListItems.Add
                    
          'Si c'est le premier enregistrement, il faut ajouter la section et la sous-section
85        If bPremierEnr = True Then
90          sSousSection = rstPieces.Fields("SousSection")
95          iOrdreSection = rstPieces.Fields("OrdreSection")
     
            'Pour avoir le nom de la section
            'Si c'est un projet électrique
100         If m_sType = "E" Then
105           Call rstSection.Open("SELECT NomSectionFR FROM GRB_SoumProjSectionElec WHERE IDSection = " & rstPieces.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
110         Else
115           Call rstSection.Open("SELECT NomSectionFR FROM GRB_SoumProjSectionMec WHERE IDSection = " & rstPieces.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
120         End If

            'Ajout du nom de la section
125         If Not IsNull(rstSection.Fields("NomSectionFR")) Then
130           itmPieces.SubItems(I_COL_NO_ITEM) = rstSection.Fields("NomSectionFR")
135         Else
140           itmPieces.SubItems(I_COL_NO_ITEM) = vbNullString
145         End If
      
150         itmPieces.ListSubItems(I_COL_NO_ITEM).Bold = True
                    
155         Call rstSection.Close
        
160         Set itmPieces = lvwPiece.ListItems.Add
      
            'Ajout du nom de la sous-section
165         If sSousSection = "PAS DE SOUS-SECTION" Then
170           itmPieces.SubItems(I_COL_DESCRIPTION) = vbNullString
175         Else
180           itmPieces.SubItems(I_COL_DESCRIPTION) = sSousSection
185         End If
            
190         itmPieces.ListSubItems(I_COL_DESCRIPTION).Bold = True
      
195         Set itmPieces = lvwPiece.ListItems.Add
      
200         bPremierEnr = False
205       Else
            'Si c'est pas le premier enregistrement, il faut vérifier avec l'ancienne section
210         If iOrdreSection <> rstPieces.Fields("OrdreSection") Then
215           iOrdreSection = rstPieces.Fields("OrdreSection")
        
220           If m_sType = "E" Then
225             Call rstSection.Open("SELECT NomSectionFR FROM GRB_SoumProjSectionElec WHERE IDSection = " & rstPieces.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
230           Else
235             Call rstSection.Open("SELECT NomSectionFR FROM GRB_SoumProjSectionMec WHERE IDSection = " & rstPieces.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
240           End If
        
245           If Not IsNull(rstSection.Fields("NomSectionFR")) Then
250             itmPieces.SubItems(I_COL_NO_ITEM) = rstSection.Fields("NomSectionFR")
255           Else
260             itmPieces.SubItems(I_COL_NO_ITEM) = vbNullString
265           End If
        
270           itmPieces.ListSubItems(I_COL_NO_ITEM).Bold = True
        
275           Call rstSection.Close
              
280           Set itmPieces = lvwPiece.ListItems.Add
        
285           sSousSection = rstPieces.Fields("SousSection")
       
290           If sSousSection = "PAS DE SOUS-SECTION" Then
295             itmPieces.SubItems(I_COL_DESCRIPTION) = vbNullString
300           Else
305             itmPieces.SubItems(I_COL_DESCRIPTION) = rstPieces.Fields("SousSection")
310           End If
        
315           itmPieces.ListSubItems(I_COL_DESCRIPTION).Bold = True
        
320           Set itmPieces = lvwPiece.ListItems.Add
325         Else
              'il faut vérifier avec l'ancienne sous-section
330           If sSousSection <> rstPieces.Fields("SousSection") Then
335             sSousSection = rstPieces.Fields("SousSection")
          
340             If sSousSection = "PAS DE SOUS-SECTION" Then
345               itmPieces.SubItems(I_COL_DESCRIPTION) = vbNullString
350             Else
355               itmPieces.SubItems(I_COL_DESCRIPTION) = sSousSection
360             End If
        
365             itmPieces.ListSubItems(I_COL_DESCRIPTION).Bold = True
        
370             Set itmPieces = lvwPiece.ListItems.Add
375           End If
380         End If
385       End If
                                
          'Quantité
390       If Not IsNull(rstPieces.Fields("Qté")) Then
395         itmPieces.Text = rstPieces.Fields("Qté")
400       Else
405         itmPieces.Text = vbNullString
410       End If
    
415       itmPieces.Tag = rstPieces.Fields("NoEnreg")
    
          'Numéro d'item
420       If Not IsNull(rstPieces.Fields("NumItem")) Then
425         itmPieces.SubItems(I_COL_NO_ITEM) = rstPieces.Fields("NumItem")
430       Else
435         itmPieces.SubItems(I_COL_NO_ITEM) = vbNullString
440       End If

445       itmPieces.ListSubItems(I_COL_NO_ITEM).Tag = rstPieces.Fields("NuméroLigne")
    
          'Description en francais
450       If Not IsNull(rstPieces.Fields("Desc_FR")) Then
455         itmPieces.SubItems(I_COL_DESCRIPTION) = rstPieces.Fields("Desc_FR")
460       Else
465         itmPieces.SubItems(I_COL_DESCRIPTION) = vbNullString
470       End If
    
          'Fabricant
475       If Not IsNull(rstPieces.Fields("Manufact")) Then
480         itmPieces.SubItems(I_COL_MANUFACTURIER) = rstPieces.Fields("Manufact")
485       Else
490         itmPieces.SubItems(I_COL_MANUFACTURIER) = vbNullString
495       End If
          
          'Fournisseur
500       If Not IsNull(rstPieces.Fields("IDFRS")) And rstPieces.Fields("IDFRS") > 0 Then
505         If itmPieces.SubItems(I_COL_NO_ITEM) <> "Texte" Then
510           Call rstFRS.Open("SELECT NomFournisseur FROM GRB_Fournisseur WHERE IDFRS = " & rstPieces.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
    
              'On affiche le nom dans la colonne
515           itmPieces.SubItems(I_COL_FOURNISSEUR) = rstFRS.Fields("NomFournisseur")
       
520           Call rstFRS.Close
525         End If
530       Else
535         itmPieces.SubItems(I_COL_FOURNISSEUR) = vbNullString
540       End If
    
545       Call rstPieces.MoveNext
550     Loop
  
555     Call rstPieces.Close
560     Set rstPieces = Nothing

565     Set rstFRS = Nothing
570     Set rstSection = Nothing

575     Exit Sub

AfficherErreur:

580     woups "frmChoixTransfertJob", "RemplirListViewPieces", Err, Erl
End Sub

Private Sub cmdAnnuler_Click()

5       On Error GoTo AfficherErreur

10      If m_sType = "E" Then
15        FrmProjSoumElec.m_bTransfertJobCancel = True
20      Else
25        FrmProjSoumMec.m_bTransfertJobCancel = True
30      End If

35      Call Unload(Me)

40      Exit Sub

AfficherErreur:

45      woups "frmChoixTransfertJob", "cmdAnnuler_Click", Err, Erl
End Sub

Private Sub cmdCreer_Click()
  
5       On Error GoTo AfficherErreur

10      Dim rstSoum   As ADODB.Recordset
15      Dim iCompteur As Integer
  
20      Set rstSoum = New ADODB.Recordset
  
25      Call rstSoum.Open("SELECT * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & m_sNoSoumission & "'", g_connData, adOpenDynamic, adLockOptimistic)
  
30      Do While Not rstSoum.EOF
35        For iCompteur = 1 To lvwPiece.ListItems.count
40          If lvwPiece.ListItems(iCompteur).Tag = rstSoum.Fields("NoEnreg") Then
45            If lvwPiece.ListItems(iCompteur).Checked = True Then
50              rstSoum.Fields("TransfertJob") = True
55            Else
60              rstSoum.Fields("TransfertJob") = False
65            End If
        
70            Call rstSoum.Update
        
75            Exit For
80          End If
85        Next
    
90        Call rstSoum.MoveNext
95      Loop
  
100     Call rstSoum.Close
105     Set rstSoum = Nothing
  
110     If m_sType = "E" Then
115       FrmProjSoumElec.m_bTransfertJobCancel = False
120     Else
125       FrmProjSoumMec.m_bTransfertJobCancel = False
130     End If
  
135     Call Unload(Me)

140     Exit Sub

AfficherErreur:

145     woups "frmChoixTransfertJob", "cmdCreer_Click", Err, Erl
End Sub

Private Sub cmdSelectAll_Click()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer

15      For iCompteur = 1 To lvwPiece.ListItems.count
20        If lvwPiece.ListItems(iCompteur).Tag <> vbNullString Then
25          If lvwPiece.ListItems(iCompteur).SubItems(I_COL_NO_ITEM) <> vbNullString Then
30            lvwPiece.ListItems(iCompteur).Checked = True
35          End If
40        End If
45      Next

50      Exit Sub

AfficherErreur:

55      woups "frmChoixTransfertJob", "cmdSelectAll_Click", Err, Erl
End Sub

Private Sub cmdDeSelectAll_Click()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer

15      For iCompteur = 1 To lvwPiece.ListItems.count
20        lvwPiece.ListItems(iCompteur).Checked = False
25      Next

30      Exit Sub

AfficherErreur:

35      woups "frmChoixTransfertJob", "cmdDeselectAll_Click", Err, Erl
End Sub

Private Sub Form_Load()
  
5       On Error GoTo AfficherErreur

10      Screen.MousePointer = vbDefault

15      Exit Sub

AfficherErreur:

20      woups "frmChoixTransfertJob", "Form_Load", Err, Erl
End Sub

Private Sub lvwPiece_ItemCheck(ByVal Item As MSComctlLib.ListItem)

5       On Error GoTo AfficherErreur
  
10      If Item.Tag = vbNullString Or Item.SubItems(I_COL_NO_ITEM) = vbNullString Then
          'On enlève le check
15        Item.Checked = False
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmChoixTransfertJob", "lvwPiece_ItemCheck", Err, Erl
End Sub
