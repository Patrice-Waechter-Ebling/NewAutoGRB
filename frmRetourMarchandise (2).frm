VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRetourMarchandise 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retour de marchandise"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmRetourMarchandise.frx":0000
   ScaleHeight     =   7680
   ScaleWidth      =   11895
   Begin MSComCtl2.MonthView mvwRetour 
      Height          =   2370
      Left            =   9120
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   0
      Appearance      =   1
      StartOfWeek     =   152633345
      CurrentDate     =   38170
   End
   Begin VB.TextBox txtDateRetour 
      Height          =   285
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdDate 
      Caption         =   "..."
      Height          =   285
      Left            =   11400
      TabIndex        =   5
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton cmdImprimer 
      Caption         =   "Imprimer"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   10560
      TabIndex        =   8
      Top             =   7200
      Width           =   1215
   End
   Begin VB.ComboBox cmbNoProjet 
      Height          =   315
      Left            =   3360
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin MSComctlLib.ListView lvwProjet 
      Height          =   6255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   11033
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Qté"
         Object.Width           =   1147
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
         Text            =   "Distributeur"
         Object.Width           =   1773
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "No. Retour"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtNoProjet 
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date de retour :"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8400
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmRetourMarchandise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Index des colonnes de lvwSoumission
Private Const I_COL_SOUM_QUANTITE As Integer = 0
Private Const I_COL_SOUM_PIECE As Integer = 1
Private Const I_COL_SOUM_DESCR As Integer = 2
Private Const I_COL_SOUM_MANUFACT As Integer = 3
Private Const I_COL_SOUM_DISTRIB As Integer = 4
Private Const I_COL_SOUM_NO_RETOUR As Integer = 5
Private Const I_COL_SOUM_DATE As Integer = 6

Private Enum enumTypeRetour
 PROJET = 0
 ACHAT = 1
End Enum

Private m_sUserID As String

'Pour l'impression
Public m_bAnnuleImpression As Boolean
Public m_eTypeImpression As enumImpressionRetour

Private m_eTypeRetour As enumTypeRetour

Public Sub Afficher(ByVal sNoProjet As String, ByVal sUserID As String)

 On Error GoTo Oups

 m_eTypeRetour = PROJET

 m_sUserID = sUserID

 Call RemplirComboProjetElec
 Call RemplirComboProjetMec

 Call NouveauRetour(sNoProjet)

 Call Me.Show(vbModal)

 Exit Sub

Oups:

 wOups "frmRetourMarchandise", "Afficher", Err, Err.number, Err.Description
End Sub

Public Sub AfficherAchat(ByVal sNoAchat As String, ByVal iIndexAchat As Integer, ByVal sUserID As String)

 On Error GoTo Oups

 m_eTypeRetour = ACHAT

 m_sUserID = sUserID

 Call RemplirComboAchats

 Call NouveauRetourAchat(sNoAchat, iIndexAchat)

 Call Me.Show(vbModal)

 Exit Sub

Oups:

 wOups "frmRetourMarchandise", "AfficherAchat", Err, Err.number, Err.Description
End Sub

Private Sub cmbNoProjet_Click()

 On Error GoTo Oups
 
 Screen.MousePointer = vbHourglass
 
 txtnoprojet.Text = cmbNoProjet.Text
 
 If m_eTypeRetour = ACHAT Then
 'Rempli les valeurs de l'achat sélectionné
 Call RemplirListViewAchat
 Else
 'Rempli les valeurs du projet sélectionné
 Call RemplirListViewProjet
 End If

 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmRetourMarchandise", "cmbNoProjet_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdImprimer_Click()

 On Error GoTo Oups

 Dim iCompteur As Integer
 Dim bChecked As Boolean
 Dim rstProjet As ADODB.Recordset
 Dim bRetourPermis As Boolean

 If cmbNoProjet.ListIndex > -1 Then
 Set rstProjet = New ADODB.Recordset

 If m_eTypeRetour = PROJET Then
 If cmbNoProjet.ItemData(cmbNoProjet.ListIndex) = 0 Then
 Call rstProjet.Open("SELECT Modification, Par FROM GrbProjetElec WHERE IDProjet = '" & Right$(txtnoprojet.Text, Len(txtnoprojet.Text) - 1) & "'", g_connData, adOpenDynamic, adLockOptimistic)
 Else
  Call rstProjet.Open("SELECT Modification, Par FROM GrbProjetMec WHERE IDProjet = '" & Right$(txtnoprojet.Text, Len(txtnoprojet.Text) - 1) & "'", g_connData, adOpenDynamic, adLockOptimistic)
  End If

  If rstProjet.Fields("Modification") = False Then
  bRetourPermis = True
  Else
  bRetourPermis = False
  End If

  Call rstProjet.Close
 Set rstProjet = Nothing
1 Else
 bRetourPermis = True
 End If

 If bRetourPermis = True Then
 For iCompteur = 1 To lvwProjet.ListItems.count
 If lvwProjet.ListItems(iCompteur).Checked = True Then
 bChecked = True
 
 Exit For
 End If
 Next

 If bChecked = True Then
 Call OuvrirForm(frmChoixImpressionRetourMarchandise, True)

 If m_bAnnuleImpression = False Then
 If m_eTypeImpression = MODE_DEMANDE_RETOUR Then
 Call ImprimerDemandeRetour
 Else
 Call ImprimerRetour
 End If
1  End If
 End If
 Else
 Call MsgBox("Ce projet est en modification par " & rstProjet.Fields("Par") & "!", vbOKOnly, "Erreur")
 End If
Else
 Call MsgBox("Vous devez choisir un numéro de retour!", vbOKOnly, "Erreur")
End If

Exit Sub

Oups:

wOups "frmRetourMarchandise", "cmdImprimer_Click", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerRetour()

 On Error GoTo Oups

 Dim collPiece As Collection
 Dim collNoLigne As Collection
 Dim iCompteur As Integer
 Dim sNoBC As String

 sNoBC = InputBox("Quel est le numéro du retour?", , txtnoprojet.Text)

 If sNoBC <> vbNullString Then
 Set collPiece = New Collection
 Set collNoLigne = New Collection

 For iCompteur = 1 To lvwProjet.ListItems.count
 If lvwProjet.ListItems(iCompteur).Checked = True Then
  Call collPiece.Add(lvwProjet.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE))
  Call collNoLigne.Add(lvwProjet.ListItems(iCompteur).Tag)
  End If
  Next

  If m_eTypeRetour = ACHAT Then
  Call frmBonCommande.AfficherFormRetourMarchandiseAchat(Replace(Trim$(Left$(txtnoprojet.Text, InStrRev(txtnoprojet.Text, "-") - 1)), "R", ""), CInt(Right$(txtnoprojet.Text, 3)), sNoBC, collPiece, collNoLigne, m_sUserID, MODE_RETOUR)
  Else
  Call frmBonCommande.AfficherFormRetourMarchandiseProjet(txtnoprojet.Text, sNoBC, collPiece, collNoLigne, m_sUserID, MODE_RETOUR)
End If
End If

Exit Sub

Oups:

wOups "frmRetourMarchandise", "ImprimerRetour", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerDemandeRetour()

 On Error GoTo Oups

 Dim collPiece As Collection
 Dim collNoLigne As Collection
 Dim iCompteur As Integer
 Dim sNoBC As String

 sNoBC = InputBox("Quel est le numéro de la demande de retour?", , txtnoprojet.Text)

 If sNoBC <> vbNullString Then
 Set collPiece = New Collection
 Set collNoLigne = New Collection

 For iCompteur = 1 To lvwProjet.ListItems.count
 If lvwProjet.ListItems(iCompteur).Checked = True Then
  Call collPiece.Add(lvwProjet.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE))
  Call collNoLigne.Add(lvwProjet.ListItems(iCompteur).Tag)
  End If
  Next

  If m_eTypeRetour = ACHAT Then
  Call frmBonCommande.AfficherFormRetourMarchandiseAchat(Trim$(Replace(Left$(txtnoprojet.Text, InStrRev(txtnoprojet.Text, "-") - 1), "R", "")), CInt(Right$(txtnoprojet.Text, 3)), sNoBC, collPiece, collNoLigne, m_sUserID, MODE_DEMANDE_RETOUR)
  Else
  Call frmBonCommande.AfficherFormRetourMarchandiseProjet(txtnoprojet.Text, sNoBC, collPiece, collNoLigne, m_sUserID, MODE_DEMANDE_RETOUR)
End If
End If

Exit Sub

Oups:

wOups "frmRetourMarchandise", "ImprimerDemandeRetour", Err, Err.number, Err.Description
End Sub

Private Sub NouveauRetour(ByVal sNoProjet As String)

 On Error GoTo Oups

 Dim rstProjet As ADODB.Recordset
 Dim iCompteur As Integer
 Dim bExiste As Boolean
 Dim eType As enumCatalogue

 bExiste = False

 If ComboContient(cmbNoProjet, "R" & sNoProjet) = False Then
 Set rstProjet = New ADODB.Recordset

 Call rstProjet.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If Not rstProjet.EOF Then
 bExiste = True

  eType = ELECTRIQUE
  End If

  Call rstProjet.Close

  If bExiste = False Then
  Call rstProjet.Open("SELECT * FROM GrbProjetMec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

  If Not rstProjet.EOF Then
  bExiste = True

  eType = MECANIQUE
 End If

Call rstProjet.Close
 End If

 If bExiste = True Then
 Call cmbNoProjet.AddItem("R" & sNoProjet)

 cmbNoProjet.ItemData(cmbNoProjet.newIndex) = eType

 cmbNoProjet.ListIndex = -1
 Else
 Call MsgBox("Projet inexistant!", vbOKOnly, "Erreur")
 End If

 Set rstProjet = Nothing
End If

1  For iCompteur = 0 To cmbNoProjet.ListCount - 1
 If cmbNoProjet.LIST(iCompteur) = "R" & sNoProjet Then
 cmbNoProjet.ListIndex = iCompteur

 Exit Sub
 End If
Next

 Exit Sub

Oups:

1  wOups "frmRetourMarchandise", "NouveauRetour", Err, Err.number, Err.Description
End Sub

Private Sub NouveauRetourAchat(ByVal sNoAchat As String, ByVal iIndexAchat As Integer)

 On Error GoTo Oups

 Dim rstAchat As ADODB.Recordset
 Dim iCompteur As Integer
 Dim bExiste As Boolean
 Dim eType As enumCatalogue

 bExiste = False

 If ComboContient(cmbNoProjet, "R" & sNoAchat & "-" & Right$("000" & iIndexAchat, 3)) = False Then
 Set rstAchat = New ADODB.Recordset

 Call rstAchat.Open("SELECT * FROM GrbAchat WHERE IDAchat = '" & sNoAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenDynamic, adLockOptimistic)

 If Not rstAchat.EOF Then
 bExiste = True

  If rstAchat.Fields("Type") = "M" Then
  eType = MECANIQUE
  Else
  eType = ELECTRIQUE
  End If
  End If

  Call rstAchat.Close
  Set rstAchat = Nothing

If bExiste = True Then
Call cmbNoProjet.AddItem("R" & sNoAchat & "-" & Right$("000" & iIndexAchat, 3))

 cmbNoProjet.ItemData(cmbNoProjet.newIndex) = eType
 Else
 Call MsgBox("Projet inexistant!", vbOKOnly, "Erreur")
 End If
End If

For iCompteur = 0 To cmbNoProjet.ListCount - 1
 If cmbNoProjet.LIST(iCompteur) = "R" & sNoAchat & "-" & Right$("000" & iIndexAchat, 3) Then
 cmbNoProjet.ListIndex = iCompteur

 Exit For
 End If
1  Next


Exit Sub

Oups:

 wOups "frmRetourMarchandise", "NouveauRetourAchat", Err, Err.number, Err.Description
End Sub

Private Sub Cmdfermer_Click()

 On Error GoTo Oups
 
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmRetourMarchandise", "cmdFermer_Click", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboProjetElec()

 On Error GoTo Oups

 'Rempli le combo des projets
 Dim rstProjet As ADODB.Recordset
 
 Set rstProjet = New ADODB.Recordset
 
 'Ouvre le recordset selon le type
 Call rstProjet.Open("SELECT DISTINCT GrbProjetElec.IDProjet FROM GrbProjetElec INNER JOIN GrbProjet_Pieces ON GrbProjetElec.IDProjet = GrbProjet_Pieces.IDProjet WHERE Retour = True ORDER BY GrbProjetElec.IDProjet", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Tant que ce n'est pas la fin des enregistrements
 Do While Not rstProjet.EOF
 Call cmbNoProjet.AddItem("R" & rstProjet.Fields("IDProjet"))

 cmbNoProjet.ItemData(cmbNoProjet.newIndex) = 0

 Call rstProjet.MoveNext
 Loop

 Call rstProjet.Close
 Set rstProjet = Nothing

  Exit Sub

Oups:

  wOups "frmRetourMarchandise", "RemplirComboProjetElec", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboProjetMec()

 On Error GoTo Oups

 'Rempli le combo des projets
 Dim rstProjet As ADODB.Recordset
 
 Set rstProjet = New ADODB.Recordset
 
 'Ouvre le recordset selon le type
 Call rstProjet.Open("SELECT DISTINCT GrbProjetMec.IDProjet FROM GrbProjetMec INNER JOIN GrbProjet_Pieces ON GrbProjetMec.IDProjet = GrbProjet_Pieces.IDProjet WHERE Retour = True ORDER BY GrbProjetMec.IDProjet", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Tant que ce n'est pas la fin des enregistrements
 Do While Not rstProjet.EOF
 Call cmbNoProjet.AddItem("R" & rstProjet.Fields("IDProjet"))

 cmbNoProjet.ItemData(cmbNoProjet.newIndex) = 1

 Call rstProjet.MoveNext
 Loop

 Call rstProjet.Close
 Set rstProjet = Nothing

  Exit Sub

Oups:

  wOups "frmRetourMarchandise", "RemplirComboProjetMec", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboAchats()

 On Error GoTo Oups

 'Rempli le combo des projets
 Dim rstAchat As ADODB.Recordset
 
 Set rstAchat = New ADODB.Recordset
 
 'Ouvre le recordset selon le type
 Call rstAchat.Open("SELECT DISTINCT GrbAchat.IDAchat, GrbAchat.IndexAchat, GrbAchat.Type FROM GrbAchat INNER JOIN GrbAchat_Pieces ON GrbAchat.IDAchat = GrbAchat_Pieces.IDAchat AND GrbAchat.IndexAchat = GrbAchat_Pieces.IndexAchat WHERE GrbAchat_Pieces.Retour = True ORDER BY GrbAchat.IDAchat, GrbAchat.IndexAchat", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Tant que ce n'est pas la fin des enregistrements
 Do While Not rstAchat.EOF
 Call cmbNoProjet.AddItem("R" & rstAchat.Fields("IDAchat") & "-" & Right$("000" & rstAchat.Fields("IndexAchat"), 3))

 If rstAchat.Fields("Type") = "E" Then
 cmbNoProjet.ItemData(cmbNoProjet.newIndex) = 0
 Else
 cmbNoProjet.ItemData(cmbNoProjet.newIndex) = 1
 End If
 
  Call rstAchat.MoveNext
  Loop

  Call rstAchat.Close
  Set rstAchat = Nothing

  Exit Sub

Oups:

  wOups "frmRetourMarchandise", "RemplirComboAchats", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewProjet()

 On Error GoTo Oups

 'Remplis les pièces de la soumission avec la BD
 Dim rstProjet As ADODB.Recordset
 Dim rstFRS As ADODB.Recordset
 Dim itmProjet As ListItem
 Dim lColor As Long
 
 If cmbNoProjet.ListIndex <> -1 Then
 Call lvwProjet.ListItems.Clear
 
 Set rstProjet = New ADODB.Recordset
 Set rstFRS = New ADODB.Recordset
 
 Call rstProjet.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & Right$(txtnoprojet.Text, Len(txtnoprojet.Text) - 1) & "' AND Left$(Qté,1) = '-' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

 Do While Not rstProjet.EOF
  Set itmProjet = lvwProjet.ListItems.Add

  itmProjet.Checked = False
 
  If rstProjet.Fields("Retour") = True Then
  lColor = COLOR_ROUGE
  Else
  If rstProjet.Fields("Commandé") = True Then
  lColor = COLOR_ORANGE 'COLOR_ORANGE
  Else
 If rstProjet.Fields("Recu") = True Then
 lColor = COLOR_GRIS 'Gris
 Else
 If rstProjet.Fields("MatérielInutile") = True Then
 lColor = COLOR_BRUN
 Else
 lColor = COLOR_NOIR
 End If
 End If
 End If
 End If

 'No Ligne
1  itmProjet.Tag = rstProjet.Fields("NuméroLigne")
 
 'Quantité
 If Not IsNull(rstProjet.Fields("Qté")) Then
 itmProjet.Text = rstProjet.Fields("Qté")
 Else
 itmProjet.Text = vbNullString
 End If

 itmProjet.ForeColor = lColor
 
 'Numéro d'item
 If Not IsNull(rstProjet.Fields("NumItem")) Then
 itmProjet.SubItems(I_COL_SOUM_PIECE) = rstProjet.Fields("NumItem")
 Else
 itmProjet.SubItems(I_COL_SOUM_PIECE) = vbNullString
 End If
 
 itmProjet.ListSubItems(I_COL_SOUM_PIECE).ForeColor = lColor
 
 'On met le nom de la sous-section dans le tag du numéro d'item
 itmProjet.ListSubItems(I_COL_SOUM_PIECE).Tag = rstProjet.Fields("SousSection")
 
 'Description en francais
 If Not IsNull(rstProjet.Fields("Desc_FR")) Then
 itmProjet.SubItems(I_COL_SOUM_DESCR) = rstProjet.Fields("Desc_FR")
 Else
 itmProjet.SubItems(I_COL_SOUM_DESCR) = vbNullString
 End If
 
 itmProjet.ListSubItems(I_COL_SOUM_DESCR).ForeColor = lColor
 
 'On met la description en anglais dans le tag de la description en francais
 If Not IsNull(rstProjet.Fields("DESC_EN")) Then
 itmProjet.ListSubItems(I_COL_SOUM_DESCR).Tag = rstProjet.Fields("Desc_EN")
Else
 itmProjet.ListSubItems(I_COL_SOUM_DESCR).Tag = vbNullString
 End If
 
 'Fabricant
 If Not IsNull(rstProjet.Fields("Manufact")) Then
 itmProjet.SubItems(I_COL_SOUM_MANUFACT) = rstProjet.Fields("Manufact")
 Else
 itmProjet.SubItems(I_COL_SOUM_MANUFACT) = vbNullString
 End If
 
 itmProjet.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = lColor
 
 'Fournisseur
 If Not IsNull(rstProjet.Fields("IDFRS")) And rstProjet.Fields("IDFRS") > 0 Then
 If itmProjet.SubItems(I_COL_SOUM_PIECE) <> "Texte" Then
 Call rstFRS.Open("SELECT NomFournisseur FROM GrbFournisseur WHERE IDFRS = " & rstProjet.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
 
 'On affiche le nom dans la colonne
 itmProjet.SubItems(I_COL_SOUM_DISTRIB) = rstFRS.Fields("NomFournisseur")
 
 itmProjet.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = lColor
 
 'On affiche l'Id dans le tag
 itmProjet.ListSubItems(I_COL_SOUM_DISTRIB).Tag = rstProjet.Fields("IDFRS")
 
 Call rstFRS.Close
 End If
 Else
 itmProjet.SubItems(I_COL_SOUM_DISTRIB) = vbNullString
 End If

4 If Not IsNull(rstProjet.Fields("NoRetour")) Then
4 itmProjet.SubItems(I_COL_SOUM_NO_RETOUR) = rstProjet.Fields("NoRetour")
4 Else
4 itmProjet.SubItems(I_COL_SOUM_NO_RETOUR) = vbNullString
4 End If

4 itmProjet.ListSubItems(I_COL_SOUM_NO_RETOUR).ForeColor = lColor

4 If Not IsNull(rstProjet.Fields("DateRetour")) Then
4 itmProjet.SubItems(I_COL_SOUM_DATE) = rstProjet.Fields("DateRetour")
4 Else
4 itmProjet.SubItems(I_COL_SOUM_DATE) = vbNullString
4 End If

4  itmProjet.ListSubItems(I_COL_SOUM_DATE).ForeColor = lColor
 
4  Call rstProjet.MoveNext
4  Loop
 
4  Call rstProjet.Close
4  Set rstProjet = Nothing

4  Set rstFRS = Nothing
4  End If
 
4  Exit Sub

Oups:

50 wOups "frmRetourMarchandise", "RemplirListViewProjet", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewAchat()

 On Error GoTo Oups

 'Remplis les pièces de la soumission avec la BD
 Dim rstAchat As ADODB.Recordset
 Dim rstFRS As ADODB.Recordset
 Dim itmAchat As ListItem
 Dim lColor As Long
 
 If cmbNoProjet.ListIndex <> -1 Then
 Call lvwProjet.ListItems.Clear
 
 Set rstAchat = New ADODB.Recordset
 Set rstFRS = New ADODB.Recordset
 
 Call rstAchat.Open("SELECT * FROM GrbAchat_Pieces WHERE IDAchat = '" & Replace(Trim$(Left$(txtnoprojet.Text, InStrRev(txtnoprojet.Text, "-") - 1)), "R", "") & "' AND IndexAchat = " & CInt(Right$(txtnoprojet.Text, 3)) & " AND Left$(Qté,1) = '-' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

 Do While Not rstAchat.EOF
  Set itmAchat = lvwProjet.ListItems.Add

  itmAchat.Checked = False
 
  If rstAchat.Fields("Retour") = True Then
  lColor = COLOR_ROUGE
  Else
  If rstAchat.Fields("Commandé") = True Then
  lColor = COLOR_ORANGE 'COLOR_ORANGE
  End If
 End If

 'No Ligne
 itmAchat.Tag = rstAchat.Fields("NuméroLigne")
 
 'Quantité
 If Not IsNull(rstAchat.Fields("Qté")) Then
 itmAchat.Text = rstAchat.Fields("Qté")
 Else
 itmAchat.Text = vbNullString
 End If

 itmAchat.ForeColor = lColor
 
 'Numéro d'item
 If Not IsNull(rstAchat.Fields("PIECE")) Then
 itmAchat.SubItems(I_COL_SOUM_PIECE) = rstAchat.Fields("PIECE")
 Else
 itmAchat.SubItems(I_COL_SOUM_PIECE) = vbNullString
 End If
 
 itmAchat.ListSubItems(I_COL_SOUM_PIECE).ForeColor = lColor
 
 'Description en francais
 If Not IsNull(rstAchat.Fields("DESC_FR")) Then
 itmAchat.SubItems(I_COL_SOUM_DESCR) = rstAchat.Fields("DESC_FR")
 Else
 itmAchat.SubItems(I_COL_SOUM_DESCR) = vbNullString
 End If
 
1  itmAchat.ListSubItems(I_COL_SOUM_DESCR).ForeColor = lColor
 
 'On met la description en anglais dans le tag de la description en francais
 If Not IsNull(rstAchat.Fields("DESC_EN")) Then
 itmAchat.ListSubItems(I_COL_SOUM_DESCR).Tag = rstAchat.Fields("DESC_EN")
 Else
 itmAchat.ListSubItems(I_COL_SOUM_DESCR).Tag = vbNullString
 End If
 
 'Fabricant
 If Not IsNull(rstAchat.Fields("Manufact")) Then
 itmAchat.SubItems(I_COL_SOUM_MANUFACT) = rstAchat.Fields("Manufact")
 Else
 itmAchat.SubItems(I_COL_SOUM_MANUFACT) = vbNullString
 End If
 
 itmAchat.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = lColor
 
 'Fournisseur
 If Not IsNull(rstAchat.Fields("IDFRS")) And rstAchat.Fields("IDFRS") > 0 Then
 Call rstFRS.Open("SELECT NomFournisseur FROM GrbFournisseur WHERE IDFRS = " & rstAchat.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
 
 'On affiche le nom dans la colonne
 itmAchat.SubItems(I_COL_SOUM_DISTRIB) = rstFRS.Fields("NomFournisseur")
 
 itmAchat.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = lColor
 
 'On affiche l'Id dans le tag
 itmAchat.ListSubItems(I_COL_SOUM_DISTRIB).Tag = rstAchat.Fields("IDFRS")
 
 Call rstFRS.Close
 Else
 itmAchat.SubItems(I_COL_SOUM_DISTRIB) = vbNullString
 End If

 If Not IsNull(rstAchat.Fields("NoRetour")) Then
 itmAchat.SubItems(I_COL_SOUM_NO_RETOUR) = rstAchat.Fields("NoRetour")
 Else
 itmAchat.SubItems(I_COL_SOUM_NO_RETOUR) = vbNullString
 End If

 itmAchat.ListSubItems(I_COL_SOUM_NO_RETOUR).ForeColor = lColor

 If Not IsNull(rstAchat.Fields("DateRetour")) Then
 itmAchat.SubItems(I_COL_SOUM_DATE) = rstAchat.Fields("DateRetour")
 Else
 itmAchat.SubItems(I_COL_SOUM_DATE) = vbNullString
 End If

 itmAchat.ListSubItems(I_COL_SOUM_DATE).ForeColor = lColor
 
 Call rstAchat.MoveNext
 Loop
 
Call rstAchat.Close
 Set rstAchat = Nothing

Set rstFRS = Nothing
End If
 
3  Exit Sub

Oups:

 wOups "frmRetourMarchandise", "RemplirListViewAchat", Err, Err.number, Err.Description
End Sub

Public Sub Retour()

 On Error GoTo Oups

 Dim rstBC As ADODB.Recordset
 Dim rstBCPiece As ADODB.Recordset
 Dim rstPiece As ADODB.Recordset
 Dim rstModif As ADODB.Recordset
 Dim rstInventaire As ADODB.Recordset
 Dim rstInvModif As ADODB.Recordset
 Dim rstEmploye As ADODB.Recordset
 Dim sWhere As String
 Dim sWherePiece As String
 Dim sWhereNoLigne As String
  Dim bPremier As Boolean
  Dim bPremierNoLigne As Boolean
  Dim bRetourFait As Boolean
  Dim sPiece As String
  Dim sNoLigne As String
  Dim sNoRetour As String

  sNoRetour = DR_Commande.Sections("Section2").Controls("lblNoSoum").Caption
 
  Set rstBC = New ADODB.Recordset
10 Set rstBCPiece = New ADODB.Recordset
Set rstPiece = New ADODB.Recordset
 
Call rstBC.Open("SELECT * FROM GrbBonsCommandes WHERE NoBonCommande = '" & sNoRetour & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Pour chaque enregistrement
Call rstBCPiece.Open("SELECT NoItem, NuméroLigne FROM GrbBonsCommandes_Pieces WHERE NoBonCommande = '" & sNoRetour & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Tant que ce n'est pas la fin des enregistrements
If m_eTypeRetour = ACHAT Then
 sWhere = "(IDAchat = '" & Replace(Trim$(Left$(txtnoprojet.Text, InStrRev(txtnoprojet.Text, "-") - 1)), "R", "") & "' AND IndexAchat = " & Int(Right$(txtnoprojet.Text, 3)) & ")"

 sWherePiece = "PIECE In ("
 sWhereNoLigne = "NuméroLigne In ("

 bPremier = True

 Do While Not rstBCPiece.EOF
 If Not IsNull(rstBCPiece.Fields("NoItem")) Then
 sNoLigne = rstBCPiece.Fields("NuméroLigne")

 If bPremier = True Then
 If InStr(1, sNoLigne, ",") = 0 Then
 sWherePiece = sWherePiece & "'" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
 sWhereNoLigne = sWhereNoLigne & rstBCPiece.Fields("NuméroLigne")
 Else
 bPremierNoLigne = True

 Do While InStr(1, sNoLigne, ",") > 0
1  If bPremierNoLigne = True Then
 sWherePiece = sWherePiece & "'" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
 sWhereNoLigne = sWhereNoLigne & Left$(sNoLigne, InStr(1, sNoLigne, ",") - 1)

 bPremierNoLigne = False
 Else
 sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
 sWhereNoLigne = sWhereNoLigne & ", " & Left$(sNoLigne, InStr(1, sNoLigne, ",") - 1)
 End If

 sNoLigne = Right$(sNoLigne, Len(sNoLigne) - (InStr(1, sNoLigne, ",") + 1))
 Loop

 If Trim$(sNoLigne) <> "" Then
 sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
 sWhereNoLigne = sWhereNoLigne & ", " & sNoLigne
 End If
 End If

 bPremier = False
 Else
 If InStr(1, sNoLigne, ",") = 0 Then
 sWherePiece = sWherePiece & ", '" & rstBCPiece.Fields("NoItem") & "'"
 sWhereNoLigne = sWhereNoLigne & ", " & rstBCPiece.Fields("NuméroLigne")
 Else
 Do While InStr(1, sNoLigne, ",") > 0
 sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
 sWhereNoLigne = sWhereNoLigne & ", " & Left$(sNoLigne, InStr(1, sNoLigne, ",") - 1)

 sNoLigne = Right$(sNoLigne, Len(sNoLigne) - (InStr(1, sNoLigne, ",") + 1))
 Loop

 If Trim$(sNoLigne) <> "" Then
 sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
 sWhereNoLigne = sWhereNoLigne & ", " & sNoLigne
 End If
 End If
 End If
 End If

 Call rstBCPiece.MoveNext
 Loop

sWherePiece = sWherePiece & ")"
 sWhereNoLigne = sWhereNoLigne & ")"

sWhere = sWhere & " AND " & sWherePiece & " AND " & sWhereNoLigne

 Call rstPiece.Open("SELECT * FROM GrbAchat_Pieces WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
3  Else
 sWhere = "(IDProjet = '" & Right$(txtnoprojet.Text, Len(txtnoprojet.Text) - 1) & "')"

sWherePiece = "NumItem In ("
4 sWhereNoLigne = "NuméroLigne In ("

4 bPremier = True

4 Do While Not rstBCPiece.EOF
4 If Not IsNull(rstBCPiece.Fields("NoItem")) Then
4 sNoLigne = rstBCPiece.Fields("NuméroLigne")

4 If bPremier = True Then
4 If InStr(1, sNoLigne, ",") = 0 Then
4 sWherePiece = sWherePiece & "'" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
4 sWhereNoLigne = sWhereNoLigne & rstBCPiece.Fields("NuméroLigne")
4 Else
4 bPremierNoLigne = True

4  Do While InStr(1, sNoLigne, ",") > 0
4  If bPremierNoLigne = True Then
4  sWherePiece = sWherePiece & "'" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
4  sWhereNoLigne = sWhereNoLigne & Left$(sNoLigne, InStr(1, sNoLigne, ",") - 1)

4  bPremierNoLigne = False
4  Else
4  sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
4  sWhereNoLigne = sWhereNoLigne & ", " & Left$(sNoLigne, InStr(1, sNoLigne, ",") - 1)
50 End If

 sNoLigne = Right$(sNoLigne, Len(sNoLigne) - (InStr(1, sNoLigne, ",") + 1))
 Loop

 If Trim$(sNoLigne) <> "" Then
 sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
 sWhereNoLigne = sWhereNoLigne & ", " & sNoLigne
 End If
 End If

 bPremier = False
 Else
 If InStr(1, sNoLigne, ",") = 0 Then
 sWherePiece = sWherePiece & ", '" & rstBCPiece.Fields("NoItem") & "'"
5  sWhereNoLigne = sWhereNoLigne & ", " & rstBCPiece.Fields("NuméroLigne")
5  Else
5  Do While InStr(1, sNoLigne, ",") > 0
5  sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
5  sWhereNoLigne = sWhereNoLigne & ", " & Left$(sNoLigne, InStr(1, sNoLigne, ",") - 1)

5  sNoLigne = Right$(sNoLigne, Len(sNoLigne) - (InStr(1, sNoLigne, ",") + 1))
5  Loop

5  If Trim$(sNoLigne) <> "" Then
60 sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
  sWhereNoLigne = sWhereNoLigne & ", " & sNoLigne
  End If
  End If
  End If
  End If

  Call rstBCPiece.MoveNext
  Loop

  sWherePiece = sWherePiece & ")"
  sWhereNoLigne = sWhereNoLigne & ")"

  sWhere = sWhere & " AND " & sWherePiece & " AND " & sWhereNoLigne

  Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
6  End If

6  Call rstBCPiece.Close
6  Set rstBCPiece = Nothing

6  Set rstInventaire = New ADODB.Recordset
6  Set rstInvModif = New ADODB.Recordset

6  Do While Not rstPiece.EOF
6  Call rstBC.MoveFirst

6  Do While Not rstBC.EOF
70 If rstBC.Fields("NoFournisseur") = rstPiece.Fields("IDFRS") Then
  If rstPiece.Fields("Retour") = True Then
  bRetourFait = True
  Else
  bRetourFait = False
  End If

  rstPiece.Fields("DateRetour") = txtDateRetour.Text

  rstPiece.Fields("Retour") = True
  rstPiece.Fields("NoRetour") = rstBC.Fields("NoBonCommande")

  If m_eTypeRetour = PROJET Then
  rstPiece.Fields("MatérielInutile") = False
  End If

   Call rstPiece.Update

   If bRetourFait = False Then
7  If rstPiece.Fields("IDFRS") = 71 Then
7  If m_eTypeRetour = ACHAT Then
7  sPiece = rstPiece.Fields("PIECE")
7  Else
7  sPiece = rstPiece.Fields("NumItem")
7  End If

80 If MsgBox("Voulez vous modifier l'inventaire pour la pièce " & sPiece & " ?", vbYesNo) = vbYes Then
  If cmbNoProjet.ItemData(cmbNoProjet.ListIndex) = 0 Then
  Call rstInventaire.Open("SELECT * FROM GrbInventaireElec WHERE NoItem = '" & Replace(sPiece, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
  Else
  Call rstInventaire.Open("SELECT * FROM GrbInventaireMec WHERE NoItem = '" & Replace(sPiece, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
  End If

  If rstInventaire.EOF Then
  Call rstInventaire.AddNew

  rstInventaire.Fields("NoItem") = sPiece
  rstInventaire.Fields("Description") = rstPiece.Fields("Desc_FR")
  rstInventaire.Fields("Manufacturier") = rstPiece.Fields("Manufact")

  Call frmChoixQteBoite.Afficher(rstPiece.Fields("NumItem"))

   rstInventaire.Fields("CommandeParBoite") = g_bQteBoite
   rstInventaire.Fields("QteBoite") = g_sQteBoite

   rstInventaire.Fields("QuantitéStock") = 0
   rstInventaire.Fields("Commentaires") = ""

8  If cmbNoProjet.ItemData(cmbNoProjet.ListIndex) = 0 Then
8  Call frmChoixLocalisation.Afficher(ELECTRIQUE, rstPiece.Fields("NumItem"))
8  Else
8  Call frmChoixLocalisation.Afficher(MECANIQUE, rstPiece.Fields("NumItem"))
90 End If

  rstInventaire.Fields("Localisation") = g_sLocalisation
  rstInventaire.Fields("Minimum") = False
  rstInventaire.Fields("QuantitéMinimum") = ""
  rstInventaire.Fields("Commande") = ""
  End If

  If rstInventaire.Fields("CommandeParBoite") = True Then
  rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) + (CDbl(Replace(rstPiece.Fields("Qté"), "-", "")) * CDbl(rstInventaire.Fields("QteBoite"))), ".", ",")
  Else
  rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) + CDbl(Replace(rstPiece.Fields("Qté"), "-", "")), ".", ",")
  End If

  If rstPiece.Fields("Prix_List") = "" Then
 rstInventaire.Fields("Prix Liste") = " "
   Else
 rstInventaire.Fields("Prix Liste") = rstPiece.Fields("Prix_List")
   End If

 rstInventaire.Fields("Escompte") = rstPiece.Fields("Escompte")
   rstInventaire.Fields("Prix net") = rstPiece.Fields("Prix_Net")

 Call rstInventaire.Update

9  Call rstInventaire.Close

 If cmbNoProjet.ItemData(cmbNoProjet.ListIndex) = 0 Then
 Call rstInvModif.Open("SELECT * FROM GrbInventaireElecModif", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstInvModif.Open("SELECT * FROM GrbInventaireMecModif", g_connData, adOpenDynamic, adLockOptimistic)
 End If

 Call rstInvModif.AddNew

 rstInvModif.Fields("Date") = txtDateRetour.Text
 rstInvModif.Fields("IDProjet") = txtnoprojet.Text
 rstInvModif.Fields("NoItem") = sPiece

 rstInvModif.Fields("Quantité") = Replace(rstPiece.Fields("Qté"), "-", "")

 rstInvModif.Fields("User") = g_sInitiale

 Call rstInvModif.Update

10  Call rstInvModif.Close
10  End If
10  End If
10  End If

10  Exit Do
10  End If
 
10  Call rstBC.MoveNext
10  Loop

110 Call rstPiece.MoveNext
110 Loop

11 Set rstInventaire = Nothing
11 Set rstInvModif = Nothing
 
11 Call rstPiece.Close
11 Set rstPiece = Nothing
 
11 Call rstBC.Close
11 Set rstBC = Nothing

11 If m_eTypeRetour = ACHAT Then
1 Call RemplirListViewAchat
11 Else
1 Call RemplirListViewProjet
 
 'Ajout aux modifs
11  Set rstModif = New ADODB.Recordset
1 Set rstEmploye = New ADODB.Recordset
 
 Call rstModif.Open("SELECT * FROM GrbProjet_Modif", g_connData, adOpenDynamic, adLockOptimistic)
 
1 Call rstModif.AddNew

 Call rstEmploye.Open("SELECT noEmploye FROM GrbEmployés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
1 rstModif.Fields("Type") = "E"
 rstModif.Fields("IDProjet") = Right$(txtnoprojet.Text, Len(txtnoprojet.Text) - 1)
11  rstModif.Fields("noEmployé") = rstEmploye.Fields("noEmploye")
 rstModif.Fields("Date") = ConvertDate(Date)
1 rstModif.Fields("Heure") = Time
1 rstModif.Fields("TypeModif") = "RETOUR"

1 Call rstEmploye.Close
1 Set rstEmploye = Nothing
 
1 Call rstModif.Update
 
1 Call rstModif.Close
1 Set rstModif = Nothing
12 End If
 
12 Exit Sub

Oups:

12 wOups "frmRetourMarchandise", "Retour", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 txtDateRetour.Text = ConvertDate(Date)

 Exit Sub

Oups:

 wOups "frmRetourMarchandise", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub lvwProjet_ItemCheck(ByVal Item As MSComctlLib.ListItem)

 On Error GoTo Oups

 If Item.Text <> vbNullString Then
 If Item.Text > 0 Then
 Item.Checked = False
 End If
 Else
 Item.Checked = False
 End If

 Exit Sub

Oups:

 wOups "frmRetourMarchandise", "lvwProjet_ItemCheck", Err, Err.number, Err.Description
End Sub

Private Sub mvwRetour_DateClick(ByVal DateClicked As Date)

 On Error GoTo Oups

 txtDateRetour.Text = ConvertDate(DateClicked)

 'Enlever le calendrier
 mvwRetour.Visible = False

 Exit Sub

Oups:

 wOups "frmRetourMarchandise", "mvwRetour_DateClick", Err, Err.number, Err.Description
End Sub

Private Sub mvwRetour_LostFocus()

 On Error GoTo Oups

 mvwRetour.Visible = False

 Exit Sub

Oups:

 wOups "frmRetourMarchandise", "mvwRetour_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub cmdDate_Click()

 On Error GoTo Oups

 'Ouverture du calendrier
 If txtDateRetour.Text <> vbNullString Then
 mvwRetour.Value = txtDateRetour.Text
 Else
 mvwRetour.Value = Date
 End If

 mvwRetour.Visible = True

 Call mvwRetour.SetFocus

 Exit Sub

Oups:

 wOups "frmRetourMarchandise", "cmdDate_Click", Err, Err.number, Err.Description
End Sub
