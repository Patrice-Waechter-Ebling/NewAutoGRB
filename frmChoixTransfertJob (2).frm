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
   MDIChild        =   -1  'True
   ScaleHeight     =   9315
   ScaleWidth      =   12315
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
Private Const I_COL_QTE As Integer = 0
Private Const I_COL_NO_ITEM As Integer = 1
Private Const I_COL_DESCRIPTION As Integer = 2
Private Const I_COL_MANUFACTURIER As Integer = 3
Private Const I_COL_FOURNISSEUR As Integer = 4

Private m_sNoSoumission As String
Private m_sType As String

Public Sub Afficher(ByVal sNoSoumission As String, ByVal sType As String)

 On Error GoTo Oups

 'Méthode pour afficher le form
 m_sNoSoumission = sNoSoumission

 m_sType = sType
 
 Call RemplirListViewPieces
 
 Call Me.Show(vbModal)

 Exit Sub

Oups:

 wOups "frmChoixTransfertJob", "Afficher", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewPieces()

 On Error GoTo Oups

 'Rempli le ListView selon le no. du projet
 Dim rstPieces As ADODB.Recordset
 Dim rstSection As ADODB.Recordset
 Dim rstFRS As Recordset
 Dim itmPieces As ListItem
 Dim bPremierEnr As Boolean
 Dim iOrdreSection As Integer
 Dim sSousSection As String
 
 bPremierEnr = True
 
 lvwPiece.Sorted = False

 Set rstFRS = New ADODB.Recordset
  Set rstPieces = New ADODB.Recordset
  Set rstSection = New ADODB.Recordset

 'Ouverture du recordset
  Call rstPieces.Open("SELECT * FROM GrbSoumission_Pieces WHERE IDSoumission = '" & m_sNoSoumission & "' AND Type = '" & m_sType & "' AND PieceExtraChargeable = False AND PieceExtraNonChargeable = False ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
 
  Do While Not rstPieces.EOF
  Set itmPieces = lvwPiece.ListItems.Add
 
 'Si c'est le premier enregistrement, il faut ajouter la section et la sous-section
  If bPremierEnr = True Then
  sSousSection = rstPieces.Fields("SousSection")
  iOrdreSection = rstPieces.Fields("OrdreSection")
 
 'Pour avoir le nom de la section
 'Si c'est un projet électrique
 If m_sType = "E" Then
 Call rstSection.Open("SELECT NomSectionFR FROM GrbSoumProjSectionElec WHERE IDSection = " & rstPieces.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstSection.Open("SELECT NomSectionFR FROM GrbSoumProjSectionMec WHERE IDSection = " & rstPieces.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
 End If

 'Ajout du nom de la section
 If Not IsNull(rstSection.Fields("NomSectionFR")) Then
 itmPieces.SubItems(I_COL_NO_ITEM) = rstSection.Fields("NomSectionFR")
 Else
 itmPieces.SubItems(I_COL_NO_ITEM) = vbNullString
 End If
 
 itmPieces.ListSubItems(I_COL_NO_ITEM).Bold = True
 
 Call rstSection.Close
 
 Set itmPieces = lvwPiece.ListItems.Add
 
 'Ajout du nom de la sous-section
 If sSousSection = "PAS DE SOUS-SECTION" Then
 itmPieces.SubItems(I_COL_DESCRIPTION) = vbNullString
 Else
 itmPieces.SubItems(I_COL_DESCRIPTION) = sSousSection
 End If
 
 itmPieces.ListSubItems(I_COL_DESCRIPTION).Bold = True
 
1  Set itmPieces = lvwPiece.ListItems.Add
 
 bPremierEnr = False
 Else
 'Si c'est pas le premier enregistrement, il faut vérifier avec l'ancienne section
 If iOrdreSection <> rstPieces.Fields("OrdreSection") Then
 iOrdreSection = rstPieces.Fields("OrdreSection")
 
 If m_sType = "E" Then
 Call rstSection.Open("SELECT NomSectionFR FROM GrbSoumProjSectionElec WHERE IDSection = " & rstPieces.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstSection.Open("SELECT NomSectionFR FROM GrbSoumProjSectionMec WHERE IDSection = " & rstPieces.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
 End If
 
 If Not IsNull(rstSection.Fields("NomSectionFR")) Then
 itmPieces.SubItems(I_COL_NO_ITEM) = rstSection.Fields("NomSectionFR")
 Else
 itmPieces.SubItems(I_COL_NO_ITEM) = vbNullString
 End If
 
 itmPieces.ListSubItems(I_COL_NO_ITEM).Bold = True
 
 Call rstSection.Close
 
 Set itmPieces = lvwPiece.ListItems.Add
 
 sSousSection = rstPieces.Fields("SousSection")
 
 If sSousSection = "PAS DE SOUS-SECTION" Then
 itmPieces.SubItems(I_COL_DESCRIPTION) = vbNullString
 Else
 itmPieces.SubItems(I_COL_DESCRIPTION) = rstPieces.Fields("SousSection")
 End If
 
 itmPieces.ListSubItems(I_COL_DESCRIPTION).Bold = True
 
 Set itmPieces = lvwPiece.ListItems.Add
 Else
 'il faut vérifier avec l'ancienne sous-section
 If sSousSection <> rstPieces.Fields("SousSection") Then
 sSousSection = rstPieces.Fields("SousSection")
 
 If sSousSection = "PAS DE SOUS-SECTION" Then
 itmPieces.SubItems(I_COL_DESCRIPTION) = vbNullString
 Else
 itmPieces.SubItems(I_COL_DESCRIPTION) = sSousSection
 End If
 
 itmPieces.ListSubItems(I_COL_DESCRIPTION).Bold = True
 
 Set itmPieces = lvwPiece.ListItems.Add
 End If
 End If
 End If
 
 'Quantité
 If Not IsNull(rstPieces.Fields("Qté")) Then
 itmPieces.Text = rstPieces.Fields("Qté")
Else
4 itmPieces.Text = vbNullString
4 End If
 
4 itmPieces.Tag = rstPieces.Fields("NoEnreg")
 
 'Numéro d'item
4 If Not IsNull(rstPieces.Fields("NumItem")) Then
4 itmPieces.SubItems(I_COL_NO_ITEM) = rstPieces.Fields("NumItem")
4 Else
4 itmPieces.SubItems(I_COL_NO_ITEM) = vbNullString
4 End If

4 itmPieces.ListSubItems(I_COL_NO_ITEM).Tag = rstPieces.Fields("NuméroLigne")
 
 'Description en francais
4 If Not IsNull(rstPieces.Fields("Desc_FR")) Then
4 itmPieces.SubItems(I_COL_DESCRIPTION) = rstPieces.Fields("Desc_FR")
4  Else
4  itmPieces.SubItems(I_COL_DESCRIPTION) = vbNullString
4  End If
 
 'Fabricant
4  If Not IsNull(rstPieces.Fields("Manufact")) Then
4  itmPieces.SubItems(I_COL_MANUFACTURIER) = rstPieces.Fields("Manufact")
4  Else
4  itmPieces.SubItems(I_COL_MANUFACTURIER) = vbNullString
4  End If
 
 'Fournisseur
50 If Not IsNull(rstPieces.Fields("IDFRS")) And rstPieces.Fields("IDFRS") > 0 Then
If itmPieces.SubItems(I_COL_NO_ITEM) <> "Texte" Then
 Call rstFRS.Open("SELECT NomFournisseur FROM GrbFournisseur WHERE IDFRS = " & rstPieces.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
 
 'On affiche le nom dans la colonne
 itmPieces.SubItems(I_COL_FOURNISSEUR) = rstFRS.Fields("NomFournisseur")
 
 Call rstFRS.Close
 End If
 Else
 itmPieces.SubItems(I_COL_FOURNISSEUR) = vbNullString
 End If
 
 Call rstPieces.MoveNext
 Loop
 
 Call rstPieces.Close
5  Set rstPieces = Nothing

5  Set rstFRS = Nothing
5  Set rstSection = Nothing

5  Exit Sub

Oups:

5  wOups "frmChoixTransfertJob", "RemplirListViewPieces", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnuler_Click()

 On Error GoTo Oups

 If m_sType = "E" Then
 FrmProjSoumElec.m_bTransfertJobCancel = True
 Else
 FrmProjSoumMec.m_bTransfertJobCancel = True
 End If

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixTransfertJob", "cmdAnnuler_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdCreer_Click()
 
 On Error GoTo Oups

 Dim rstSoum As ADODB.Recordset
 Dim iCompteur As Integer
 
 Set rstSoum = New ADODB.Recordset
 
 Call rstSoum.Open("SELECT * FROM GrbSoumission_Pieces WHERE IDSoumission = '" & m_sNoSoumission & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 Do While Not rstSoum.EOF
 For iCompteur = 1 To lvwPiece.ListItems.count
 If lvwPiece.ListItems(iCompteur).Tag = rstSoum.Fields("NoEnreg") Then
 If lvwPiece.ListItems(iCompteur).Checked = True Then
 rstSoum.Fields("TransfertJob") = True
 Else
  rstSoum.Fields("TransfertJob") = False
  End If
 
  Call rstSoum.Update
 
  Exit For
  End If
  Next
 
  Call rstSoum.MoveNext
  Loop
 
10 Call rstSoum.Close
Set rstSoum = Nothing
 
If m_sType = "E" Then
 FrmProjSoumElec.m_bTransfertJobCancel = False
Else
 FrmProjSoumMec.m_bTransfertJobCancel = False
End If
 
Call Unload(Me)

Exit Sub

Oups:

wOups "frmChoixTransfertJob", "cmdCreer_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdSelectAll_Click()

 On Error GoTo Oups

 Dim iCompteur As Integer

 For iCompteur = 1 To lvwPiece.ListItems.count
 If lvwPiece.ListItems(iCompteur).Tag <> vbNullString Then
 If lvwPiece.ListItems(iCompteur).SubItems(I_COL_NO_ITEM) <> vbNullString Then
 lvwPiece.ListItems(iCompteur).Checked = True
 End If
 End If
 Next

 Exit Sub

Oups:

 wOups "frmChoixTransfertJob", "cmdSelectAll_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdDeSelectAll_Click()

 On Error GoTo Oups

 Dim iCompteur As Integer

 For iCompteur = 1 To lvwPiece.ListItems.count
 lvwPiece.ListItems(iCompteur).Checked = False
 Next

 Exit Sub

Oups:

 wOups "frmChoixTransfertJob", "cmdDeselectAll_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()
 
 On Error GoTo Oups

 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmChoixTransfertJob", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub lvwPiece_ItemCheck(ByVal Item As MSComctlLib.ListItem)

 On Error GoTo Oups
 
 If Item.Tag = vbNullString Or Item.SubItems(I_COL_NO_ITEM) = vbNullString Then
 'On enlève le check
 Item.Checked = False
 End If

 Exit Sub

Oups:

 wOups "frmChoixTransfertJob", "lvwPiece_ItemCheck", Err, Err.number, Err.Description
End Sub
