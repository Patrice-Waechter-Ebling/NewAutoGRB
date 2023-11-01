VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmReceptionMec 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Réception mécanique"
   ClientHeight    =   7755
   ClientLeft      =   210
   ClientTop       =   630
   ClientWidth     =   11925
   Icon            =   "FrmReceptionMec.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7755
   ScaleWidth      =   11925
   Begin VB.CommandButton cmdNonRecu 
      Caption         =   "Pièces non recues"
      Height          =   375
      Left            =   3600
      TabIndex        =   17
      Top             =   7320
      Width           =   1695
   End
   Begin MSComCtl2.MonthView mvwReception 
      Height          =   2370
      Left            =   9240
      TabIndex        =   3
      Top             =   240
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
   Begin VB.Frame fraPiecesNonRecues 
      BackColor       =   &H00000000&
      Height          =   6255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   11655
      Begin MSComCtl2.MonthView mvwDateRequise 
         Height          =   2370
         Left            =   2760
         TabIndex        =   8
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
      Begin VB.CheckBox chkProjetAchat 
         BackColor       =   &H00000000&
         Caption         =   "Projet / Achat :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox chkDateRequise 
         BackColor       =   &H00000000&
         Caption         =   "Date Requise :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtProjetAchat 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   20
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtDateRequise 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdFermerPieces 
         Caption         =   "Fermer"
         Height          =   375
         Left            =   9960
         TabIndex        =   13
         Top             =   5760
         Width           =   1575
      End
      Begin VB.CommandButton cmdImprimerPieces 
         Caption         =   "Imprimer"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   5760
         Width           =   1575
      End
      Begin VB.CommandButton cmdAfficher 
         Caption         =   "Afficher"
         Height          =   405
         Left            =   4080
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvwPieces 
         Height          =   4455
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   7858
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
            Text            =   "# Projet"
            Object.Width           =   1588
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Qté"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "# Item"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Description"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Fournisseur"
            Object.Width           =   4762
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Date Commande"
            Object.Width           =   2408
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Date Requise"
            Object.Width           =   2408
         EndProperty
      End
      Begin VB.CommandButton cmdDateRequise 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Left            =   3240
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdImprimer 
      Caption         =   "Imprimer"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   7320
      Width           =   1575
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      ItemData        =   "FrmReceptionMec.frx":2CFA
      Left            =   3120
      List            =   "FrmReceptionMec.frx":2D04
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdDate 
      Caption         =   "..."
      Height          =   285
      Left            =   11160
      TabIndex        =   6
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox txtDateReception 
      Height          =   285
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler la réception"
      Height          =   375
      Left            =   1800
      TabIndex        =   16
      Top             =   7320
      Width           =   1695
   End
   Begin VB.ComboBox cmbNoProjet 
      Height          =   315
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox txtNoProjet 
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSComctlLib.ListView lvwProjet 
      Height          =   6255
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   11033
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
      NumItems        =   8
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
         Text            =   "Date Réception"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Date Commande"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Date Requise"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   10560
      TabIndex        =   18
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date de réception :"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8160
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "FrmReceptionMec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Index des colonnes de lvwSoumission
Private Const I_COL_QUANTITE As Integer = 0
Private Const I_COL_PIECE As Integer = 1
Private Const I_COL_DESCRIPTION As Integer = 2
Private Const I_COL_MANUFACTURIER As Integer = 3
Private Const I_COL_DISTRIBUTEUR As Integer = 4
Private Const I_COL_DATE_RECEPTION As Integer = 5
Private Const I_COL_DATE_COMMANDE As Integer = 6
Private Const I_COL_DATE_REQUISE As Integer = 7

Private Const I_LVW_PROJET As Integer = 0
Private Const I_LVW_QUANTITE As Integer = 1
Private Const I_LVW_PIECE As Integer = 2
Private Const I_LVW_DESCRIPTION As Integer = 3
Private Const I_LVW_FOURNISSEUR As Integer = 4
Private Const I_LVW_DATE_COMMANDE As Integer = 5
Private Const I_LVW_DATE_REQUISE As Integer = 6

Private Enum enumType
 PROJET = 0
 ACHAT = 1
End Enum

Private m_sUserID As String
Private m_sNoProjet As String
Private m_sNoAchat As String
Private m_eType As enumType
Private m_iIndexReception As Integer

Private Sub chkDateRequise_Click()

 On Error GoTo Oups

 If chkDateRequise.Value = vbChecked Then
 txtDateRequise.Enabled = True
 cmdDateRequise.Enabled = True
 Else
 txtDateRequise.Enabled = False
 cmdDateRequise.Enabled = False
 End If

 Exit Sub

Oups:

 wOups "frmReceptionMec", "chkDateRequise_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkProjetAchat_Click()

 On Error GoTo Oups

 If chkProjetAchat.Value = vbChecked Then
 txtProjetAchat.Enabled = True
 Else
 txtProjetAchat.Enabled = False
 End If

 Exit Sub

Oups:

 wOups "frmReceptionMec", "chkProjetAchat_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmbNoProjet_KeyUp(KeyCode As Integer, Shift As Integer)
 
 On Error GoTo Oups

 Dim iCompteur As Integer
 
 For iCompteur = 0 To cmbNoProjet.ListCount - 1
 If UCase(cmbNoProjet.LIST(iCompteur)) = UCase(cmbNoProjet.Text) Then
 cmbNoProjet.ListIndex = iCompteur
 
 Exit For
 End If
 Next

 Exit Sub

Oups:

 wOups "frmReceptionMec", "cmbNoProjet_KeyUp", Err, Err.number, Err.Description
End Sub

Private Sub cmbNoProjet_Click()

 On Error GoTo Oups

 Dim rstProjAchat As ADODB.Recordset
 Dim sNumero As String

 Set rstProjAchat = New ADODB.Recordset

 sNumero = txtnoprojet.Text

 If m_eType = ACHAT Then
 Call rstProjAchat.Open("SELECT * FROM GrbAchat WHERE IDAchat = '" & Left$(cmbNoProjet.Text, 9) & "' AND IndexAchat = " & CInt(Right$(cmbNoProjet.Text, 3)), g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstProjAchat.Open("SELECT * FROM GrbProjetMec WHERE IDProjet = '" & cmbNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 End If

 If rstProjAchat.Fields("Modification") = True Then
  If m_eType = ACHAT Then
  Call MsgBox("Cet achat est en modification par " & rstProjAchat.Fields("Par") & "!", vbOKOnly, "Erreur")
  Else
  Call MsgBox("Ce projet est en modification par " & rstProjAchat.Fields("Par") & "!", vbOKOnly, "Erreur")
  End If

  Call rstProjAchat.Close
  Set rstProjAchat = Nothing

  cmbNoProjet.Text = sNumero

Exit Sub
End If
 
Screen.MousePointer = vbHourglass

m_iIndexReception = 0

txtnoprojet.Text = cmbNoProjet.Text
 
If m_eType = PROJET Then
 'Rempli les valeurs du projet sélectionné
 Call RemplirListViewProjet(txtnoprojet.Text)
Else
 Call RemplirListViewAchat(txtnoprojet.Text)
End If

Call VerifierBoutonAnnuler

Screen.MousePointer = vbDefault

1  Exit Sub

Oups:

wOups "frmReceptionMec", "cmbNoProjet_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAfficher_Click()

 On Error GoTo Oups

 Dim bRemplir As Boolean

 If chkProjetAchat.Value = vbChecked Then
 If Trim$(txtProjetAchat.Text) <> "" Then
 If m_eType = ACHAT Then
 If Len(Trim$(txtProjetAchat.Text)) = 13 Then
 bRemplir = True
 Else
 Call MsgBox("Format de numéro d'achat incorrect!", vbOKOnly, "Erreur")
 End If
 Else
  If Len(Trim$(txtProjetAchat.Text)) =   Then
  bRemplir = True
  Else
  Call MsgBox("Format de numéro de projet incorrect!", vbOKOnly, "Erreur")
  End If
  End If
  Else
  If m_eType = ACHAT Then
 Call MsgBox("Le numéro de l'achat est obligatoire!", vbOKOnly, "Erreur")
Else
 Call MsgBox("Le numéro de projet est obligatoire!", vbOKOnly, "Erreur")
 End If
 End If
Else
 bRemplir = True
End If

If bRemplir = True Then
 Screen.MousePointer = vbHourglass

 Call RemplirListePiecesNonRecues

 Screen.MousePointer = vbDefault
1  End If

Exit Sub

Oups:

 wOups "frmReceptionMec", "cmdAfficher_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnuler_Click()

 On Error GoTo Oups

 If m_eType = PROJET Then
 Call AnnulerReceptionProjet
 Else
 Call AnnulerReceptionAchat
 End If

 Exit Sub

Oups:

 wOups "frmReceptionMec", "cmdAnnuler_Click", Err, Err.number, Err.Description
End Sub

Private Sub AnnulerReceptionProjet()

 On Error GoTo Oups

 Dim rstProjet As ADODB.Recordset
 Dim rstPiece As ADODB.Recordset
 Dim rstModif As ADODB.Recordset
 Dim rstEmploye As ADODB.Recordset
 
 'S'il y a des enregistrements
 If lvwProjet.ListItems.count > 0 Then
 Set rstProjet = New ADODB.Recordset

 Call rstProjet.Open("SELECT Modification, Par FROM GrbProjetMec WHERE IDProjet = '" & txtnoprojet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If rstProjet.Fields("Modification") = False Then
 Set rstPiece = New ADODB.Recordset

 Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & txtnoprojet.Text & "' AND Type = 'M' AND NuméroLigne = " & lvwProjet.SelectedItem.ListSubItems(I_COL_MANUFACTURIER).Tag, g_connData, adOpenDynamic, adLockOptimistic)

  rstPiece.Fields("Recu") = False
  rstPiece.Fields("Commandé") = True

  rstPiece.Fields("DateRéception") = ""

  Call rstPiece.Update

 'Ajout aux modifs
  Set rstModif = New ADODB.Recordset
 
  Call rstModif.Open("SELECT * FROM GrbProjet_Modif", g_connData, adOpenDynamic, adLockOptimistic)
 
  Call rstModif.AddNew

 Set rstEmploye = New ADODB.Recordset

Call rstEmploye.Open("SELECT noEmploye FROM GrbEmployés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 rstModif.Fields("Type") = "M"
 rstModif.Fields("IDProjet") = txtnoprojet.Text
 rstModif.Fields("noEmployé") = rstEmploye.Fields("noEmploye")
 rstModif.Fields("Date") = ConvertDate(Date)
 rstModif.Fields("Heure") = Time
 rstModif.Fields("TypeModif") = "RECEPTION"

 Call rstEmploye.Close
 Set rstEmploye = Nothing
 
 Call rstModif.Update
 
 Call rstModif.Close
 Set rstModif = Nothing

 Call rstPiece.Close
 Set rstPiece = Nothing

 If UCase(lvwProjet.SelectedItem.SubItems(I_COL_DISTRIBUTEUR)) = "SOLUTION GRB INC." Then
 Call AjouterInventaireProjet
 End If

 Call SupprimerHistorique

1  m_iIndexReception = lvwProjet.SelectedItem.Index

 Call RemplirListViewProjet(txtnoprojet.Text)
 Else
 Call MsgBox("Ce projet est en modification par " & rstProjet.Fields("Par") & "!", vbOKOnly, "Erreur")
 End If

 Call rstProjet.Close
 Set rstProjet = Nothing
End If

Exit Sub

Oups:

wOups "frmReceptionMec", "AnnulerReceptionProjet", Err, Err.number, Err.Description
End Sub

Private Sub AnnulerReceptionAchat()

 On Error GoTo Oups

 Dim rstPiece As ADODB.Recordset
 Dim rstAchat As ADODB.Recordset
 Dim sIDAchat As String
 Dim iIndexAchat As Integer
 
 'S'il y a des enregistrements
 If lvwProjet.ListItems.count > 0 Then
 sIDAchat = Left$(txtnoprojet.Text, 9)
 iIndexAchat = CInt(Right$(txtnoprojet.Text, 3))

 Set rstAchat = New ADODB.Recordset

 Call rstAchat.Open("SELECT Modification, Par FROM GrbAchat WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenDynamic, adLockOptimistic)

 If rstAchat.Fields("Modification") = False Then
  Set rstPiece = New ADODB.Recordset

  Call rstPiece.Open("SELECT * FROM GrbAchat_Pieces WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat & " And NuméroLigne = " & lvwProjet.SelectedItem.ListSubItems(I_COL_MANUFACTURIER).Tag, g_connData, adOpenDynamic, adLockOptimistic)

  rstPiece.Fields("Recu") = False
  rstPiece.Fields("Commandé") = True

  rstPiece.Fields("DateRéception") = ""

  Call rstPiece.Update

  If (CInt(Right$(sIDAchat, 2)) >= 0 And CInt(Right$(sIDAchat, 2)) <= 12) Or rstPiece.Fields("IDFRS") = 71 Then
 Call EnleverInventaireAchat
End If

 Call rstPiece.Close
 Set rstPiece = Nothing

 m_iIndexReception = lvwProjet.SelectedItem.Index

 Call RemplirListViewAchat(txtnoprojet.Text)
 Else
 Call MsgBox("Ce projet est en modification par " & rstAchat.Fields("Par") & "!", vbOKOnly, "Erreur")
 End If

 Call rstAchat.Close
 Set rstAchat = Nothing
End If

1  Exit Sub

Oups:

wOups "frmReceptionMec", "AnnulerReceptionAchat", Err, Err.number, Err.Description
End Sub

Private Sub EnleverInventaireAchat()
 
 On Error GoTo Oups
 
 Dim rstInventaire As ADODB.Recordset
 Dim sQuantite As String

 If MsgBox("Voulez-vous modifier l'inventaire?", vbYesNo) = vbYes Then
 Set rstInventaire = New ADODB.Recordset

 Call rstInventaire.Open("SELECT * FROM GrbInventaireMec WHERE NoItem = '" & lvwProjet.SelectedItem.SubItems(I_COL_PIECE) & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If Not rstInventaire.EOF Then
 If rstInventaire.Fields("CommandeParBoite") = True Then
 If lvwProjet.SelectedItem.ListSubItems(I_COL_DISTRIBUTEUR).Tag = 71 Then
 rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) + (CDbl(lvwProjet.SelectedItem.Text) * CDbl(rstInventaire.Fields("QteBoite"))), "", ",")
 Else
  rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) - (CDbl(lvwProjet.SelectedItem.Text) * CDbl(rstInventaire.Fields("QteBoite"))), ".", ",")
  End If

  sQuantite = CDbl(lvwProjet.SelectedItem.Text) * CDbl(rstInventaire.Fields("QteBoite"))
  Else
  If lvwProjet.SelectedItem.ListSubItems(I_COL_DISTRIBUTEUR).Tag = 71 Then
  rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) + CDbl(lvwProjet.SelectedItem.Text), "", ",")
  Else
  rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) - CDbl(lvwProjet.SelectedItem.Text), ".", ",")
 End If

 sQuantite = CDbl(lvwProjet.SelectedItem.Text)
 End If
 
 Call rstInventaire.Update
 End If

 Call rstInventaire.Close
 Set rstInventaire = Nothing
 
 Call SupprimerHistorique(sQuantite)
End If

Exit Sub

Oups:
 
wOups "frmReceptionMec", "EnleverInventaireAchat", Err, Err.number, Err.Description
End Sub

Private Sub AjouterInventaireProjet()

 On Error GoTo Oups

 Dim rstInventaire As ADODB.Recordset
 Dim sQuantite As String

 If MsgBox("Voulez-vous modifier l'inventaire?", vbYesNo) = vbYes Then
 Set rstInventaire = New ADODB.Recordset

 Call rstInventaire.Open("SELECT * FROM GrbInventaireMec WHERE NoItem = '" & Replace(lvwProjet.SelectedItem.SubItems(I_COL_PIECE), "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If Not rstInventaire.EOF Then
 If rstInventaire.Fields("CommandeParBoite") = True Then
 rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) + (CDbl(lvwProjet.SelectedItem.Text) * CDbl(rstInventaire.Fields("QteBoite"))), ".", ",")

 sQuantite = CDbl(lvwProjet.SelectedItem.Text) * CDbl(rstInventaire.Fields("QteBoite"))
 Else
  rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) + CDbl(lvwProjet.SelectedItem.Text), ".", ",")

  sQuantite = CDbl(lvwProjet.SelectedItem.Text)
  End If

  Call rstInventaire.Update
  End If

  Call rstInventaire.Close
  Set rstInventaire = Nothing

  Call SupprimerHistorique(sQuantite)
10 End If

Exit Sub

Oups:

wOups "frmReceptionMec", "AjouterInventaireProjet", Err, Err.number, Err.Description
End Sub

Private Sub cmdDateRequise_Click()

 On Error GoTo Oups

 'Ouverture du calendrier
 If txtDateRequise.Text <> vbNullString Then
 mvwDateRequise.Value = txtDateRequise.Text
 Else
 mvwDateRequise.Value = Date
 End If

 mvwDateRequise.Visible = True

 Call mvwDateRequise.SetFocus

 Exit Sub

Oups:

 wOups "frmReceptionMec", "cmdDateRequise_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdFermerPieces_Click()

 On Error GoTo Oups

 fraPiecesNonRecues.Visible = False

 Exit Sub

Oups:

 wOups "frmReceptionMec", "cmdFermerPieces_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdImprimerPieces_Click()

 On Error GoTo Oups

 Dim rstPiece As ADODB.Recordset

 Set rstPiece = New ADODB.Recordset

 If m_eType = PROJET Then
 If chkDateRequise.Value = vbChecked And chkProjetAchat.Value = vbChecked Then
 Call rstPiece.Open("SELECT GrbProjet_Pieces.*, GrbFournisseur.NomFournisseur FROM GrbProjet_Pieces INNER JOIN GrbFournisseur ON GrbProjet_Pieces.IDFRS = GrbFournisseur.IDFRS WHERE Type = 'M' AND Commandé = True AND DateRequise <= '" & txtDateRequise.Text & "' AND IDProjet = '" & txtProjetAchat.Text & "' AND DateRequise <> '' ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 If chkDateRequise.Value = vbChecked Then
 Call rstPiece.Open("SELECT GrbProjet_Pieces.*, GrbFournisseur.NomFournisseur FROM GrbProjet_Pieces INNER JOIN GrbFournisseur ON GrbProjet_Pieces.IDFRS = GrbFournisseur.IDFRS WHERE Type = 'M' AND Commandé = True AND DateRequise <= '" & txtDateRequise.Text & "' AND DateRequise <> '' ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 If chkProjetAchat.Value = vbChecked Then
  Call rstPiece.Open("SELECT GrbProjet_Pieces.*, GrbFournisseur.NomFournisseur FROM GrbProjet_Pieces INNER JOIN GrbFournisseur ON GrbProjet_Pieces.IDFRS = GrbFournisseur.IDFRS WHERE Type = 'M' AND Commandé = True AND IDProjet = '" & txtProjetAchat.Text & "' ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
  Else
  Call rstPiece.Open("SELECT GrbProjet_Pieces.*, GrbFournisseur.NomFournisseur FROM GrbProjet_Pieces INNER JOIN GrbFournisseur ON GrbProjet_Pieces.IDFRS = GrbFournisseur.IDFRS WHERE Type = 'M' AND Commandé = True ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
  End If
  End If
  End If
  Else
  If chkDateRequise.Value = vbChecked And chkProjetAchat.Value = vbChecked Then
 Call rstPiece.Open("SELECT (GrbAchat_Pieces.IDAchat & '-' & RIGHT('00' & GrbAchat_Pieces.IndexAchat,3)) AS NoAchat, GrbAchat_Pieces.*, GrbFournisseur.NomFournisseur FROM GrbAchat_Pieces INNER JOIN GrbFournisseur ON GrbAchat_Pieces.IDFRS = GrbFournisseur.IDFRS WHERE LEN(IDAchat) =   AND Commandé = True AND DateRequise <= '" & txtDateRequise.Text & "' AND IDAchat = '" & Left$(txtProjetAchat.Text, 9) & "' AND IndexAchat = " & CInt(Right$(txtProjetAchat.Text, 3)) & " AND DateRequise <> '' ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
1 Else
 If chkDateRequise.Value = vbChecked Then
 Call rstPiece.Open("SELECT (GrbAchat_Pieces.IDAchat & '-' & RIGHT('00' & GrbAchat_Pieces.IndexAchat,3)) AS NoAchat, GrbAchat_Pieces.*, GrbFournisseur.NomFournisseur FROM GrbAchat_Pieces INNER JOIN GrbFournisseur ON GrbAchat_Pieces.IDFRS = GrbFournisseur.IDFRS WHERE LEN(IDAchat) =   AND Commandé = True AND DateRequise <= '" & txtDateRequise.Text & "' AND DateRequise <> '' ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 If chkProjetAchat.Value = vbChecked Then
 Call rstPiece.Open("SELECT (GrbAchat_Pieces.IDAchat & '-' & RIGHT('00' & GrbAchat_Pieces.IndexAchat,3)) AS NoAchat, GrbAchat_Pieces.*, GrbFournisseur.NomFournisseur FROM GrbAchat_Pieces INNER JOIN GrbFournisseur ON GrbAchat_Pieces.IDFRS = GrbFournisseur.IDFRS WHERE LEN(IDAchat) =   AND Commandé = True AND IDAchat = '" & Left$(txtProjetAchat.Text, 9) & "' AND IndexAchat = " & CInt(Right$(txtProjetAchat.Text, 3)) & " ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstPiece.Open("SELECT (GrbAchat_Pieces.IDAchat & '-' & RIGHT('00' & GrbAchat_Pieces.IndexAchat,3)) AS NoAchat, GrbAchat_Pieces.*, GrbFournisseur.NomFournisseur FROM GrbAchat_Pieces INNER JOIN GrbFournisseur ON GrbAchat_Pieces.IDFRS = GrbFournisseur.IDFRS WHERE LEN(IDAchat) =   AND Commandé = True ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
 End If
 End If
 End If
1  End If

Set DR_BackOrder.DataSource = rstPiece

 DR_BackOrder.Orientation = rptOrientLandscape

If m_eType = PROJET Then
 DR_BackOrder.Sections("Section4").Controls("lblTitre").Caption = "Projets mécaniques : Pièces non reçues"

 DR_BackOrder.Sections("Section4").Controls("lblTitreProjetAchat").Caption = "Projet : "
 Else
1  DR_BackOrder.Sections("Section4").Controls("lblTitre").Caption = "Achats mécaniques : Pièces non reçues"

 DR_BackOrder.Sections("Section4").Controls("lblTitreProjetAchat").Caption = "Achat : "
 End If

DR_BackOrder.Sections("Section4").Controls("lblDate").Caption = txtDateRequise.Text

DR_BackOrder.Sections("Section4").Controls("lblProjetAchat").Caption = txtProjetAchat.Text

If m_eType = ACHAT Then
 DR_BackOrder.Sections("Section2").Controls("lblTitreNoProjet").Caption = "# Achat"

 DR_BackOrder.Sections("Section1").Controls("txtNoProjAchat").DataField = "NoAchat"
 DR_BackOrder.Sections("Section1").Controls("txtNoItem").DataField = "PIECE"
End If

Call DR_BackOrder.Show(vbModal)

Call rstPiece.Close
Set rstPiece = Nothing

2  Exit Sub

Oups:

wOups "frmReceptionMec", "cmdImprimerPieces_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdNonRecu_Click()

 On Error GoTo Oups

 Call lvwPieces.ListItems.Clear

 If m_eType = ACHAT Then
 chkProjetAchat.Caption = "No achat : "
 Else
 chkProjetAchat.Caption = "No projet : "
 End If

 fraPiecesNonRecues.Visible = True

 Exit Sub

Oups:

 wOups "frmReceptionMec", "cmdNonRecu_Click", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListePiecesNonRecues()

 On Error GoTo Oups

 Dim itmPiece As ListItem
 Dim rstPiece As ADODB.Recordset

 Call lvwPieces.ListItems.Clear

 Set rstPiece = New ADODB.Recordset

 If m_eType = PROJET Then
 If chkDateRequise.Value = vbChecked And chkProjetAchat.Value = vbChecked Then
 Call rstPiece.Open("SELECT GrbProjet_Pieces.*, GrbFournisseur.NomFournisseur FROM GrbProjet_Pieces INNER JOIN GrbFournisseur ON GrbProjet_Pieces.IDFRS = GrbFournisseur.IDFRS WHERE Type = 'M' AND Commandé = True AND DateRequise <= '" & txtDateRequise.Text & "' AND IDProjet = '" & txtProjetAchat.Text & "' AND DateRequise <> '' ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 If chkDateRequise.Value = vbChecked Then
 Call rstPiece.Open("SELECT GrbProjet_Pieces.*, GrbFournisseur.NomFournisseur FROM GrbProjet_Pieces INNER JOIN GrbFournisseur ON GrbProjet_Pieces.IDFRS = GrbFournisseur.IDFRS WHERE Type = 'M' AND Commandé = True AND DateRequise <= '" & txtDateRequise.Text & "' AND DateRequise <> '' ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
  Else
  If chkProjetAchat.Value = vbChecked Then
  Call rstPiece.Open("SELECT GrbProjet_Pieces.*, GrbFournisseur.NomFournisseur FROM GrbProjet_Pieces INNER JOIN GrbFournisseur ON GrbProjet_Pieces.IDFRS = GrbFournisseur.IDFRS WHERE Type = 'M' AND Commandé = True AND IDProjet = '" & txtProjetAchat.Text & "' ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
  Else
  Call rstPiece.Open("SELECT GrbProjet_Pieces.*, GrbFournisseur.NomFournisseur FROM GrbProjet_Pieces INNER JOIN GrbFournisseur ON GrbProjet_Pieces.IDFRS = GrbFournisseur.IDFRS WHERE Type = 'M' AND Commandé = True ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
  End If
  End If
  End If
10 Else
1 If chkDateRequise.Value = vbChecked And chkProjetAchat.Value = vbChecked Then
 Call rstPiece.Open("SELECT GrbAchat_Pieces.*, GrbFournisseur.NomFournisseur FROM GrbAchat_Pieces INNER JOIN GrbFournisseur ON GrbAchat_Pieces.IDFRS = GrbFournisseur.IDFRS WHERE LEN(IDAchat) =   AND Commandé = True AND DateRequise <= '" & txtDateRequise.Text & "' AND IDAchat = '" & Left$(txtProjetAchat.Text, 9) & "' AND IndexAchat = " & CInt(Right$(txtProjetAchat.Text, 3)) & " AND DateRequise <> '' ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 If chkDateRequise.Value = vbChecked Then
 Call rstPiece.Open("SELECT GrbAchat_Pieces.*, GrbFournisseur.NomFournisseur FROM GrbAchat_Pieces INNER JOIN GrbFournisseur ON GrbAchat_Pieces.IDFRS = GrbFournisseur.IDFRS WHERE LEN(IDAchat) =   AND Commandé = True AND DateRequise <= '" & txtDateRequise.Text & "' AND DateRequise <> '' ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 If chkProjetAchat.Value = vbChecked Then
 Call rstPiece.Open("SELECT GrbAchat_Pieces.*, GrbFournisseur.NomFournisseur FROM GrbAchat_Pieces INNER JOIN GrbFournisseur ON GrbAchat_Pieces.IDFRS = GrbFournisseur.IDFRS WHERE LEN(IDAchat) =   AND Commandé = True AND IDAchat = '" & Left$(txtProjetAchat.Text, 9) & "' AND IndexAchat = " & CInt(Right$(txtProjetAchat.Text, 3)) & " ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstPiece.Open("SELECT GrbAchat_Pieces.*, GrbFournisseur.NomFournisseur FROM GrbAchat_Pieces INNER JOIN GrbFournisseur ON GrbAchat_Pieces.IDFRS = GrbFournisseur.IDFRS WHERE LEN(IDAchat) =   AND Commandé = True ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
 End If
 End If
 End If
 End If

Do While Not rstPiece.EOF
 Set itmPiece = lvwPieces.ListItems.Add

 If m_eType = PROJET Then
 itmPiece.Text = rstPiece.Fields("IDProjet")
1  Else
 itmPiece.Text = rstPiece.Fields("IDAchat") & "-" & Right$("00" & rstPiece.Fields("IndexAchat"), 3)
 End If

 itmPiece.SubItems(I_LVW_QUANTITE) = rstPiece.Fields("Qté")

 If m_eType = PROJET Then
 itmPiece.SubItems(I_LVW_PIECE) = rstPiece.Fields("NumItem")
 Else
 itmPiece.SubItems(I_LVW_PIECE) = rstPiece.Fields("PIECE")
 End If

 itmPiece.SubItems(I_LVW_DESCRIPTION) = rstPiece.Fields("Desc_FR")
 itmPiece.SubItems(I_LVW_FOURNISSEUR) = rstPiece.Fields("NomFournisseur")

 If Not IsNull(rstPiece.Fields("DateCommande")) Then
 itmPiece.SubItems(I_LVW_DATE_COMMANDE) = rstPiece.Fields("DateCommande")
Else
 itmPiece.SubItems(I_LVW_DATE_COMMANDE) = ""
End If

 If Not IsNull(rstPiece.Fields("DateRequise")) Then
 itmPiece.SubItems(I_LVW_DATE_REQUISE) = rstPiece.Fields("DateRequise")
 Else
 itmPiece.SubItems(I_LVW_DATE_REQUISE) = ""
 End If

Call rstPiece.MoveNext
Loop

Call rstPiece.Close
Set rstPiece = Nothing

Exit Sub

Oups:

wOups "frmReceptionMec", "RemplirListePiecesNonRecues", Err, Err.number, Err.Description
End Sub

Private Sub lvwProjet_Click()

 On Error GoTo Oups

 Call VerifierBoutonAnnuler

 Exit Sub

Oups:

 wOups "frmReceptionMec", "lvwProjet_Click", Err, Err.number, Err.Description
End Sub

Private Sub lvwProjet_DblClick()

 On Error GoTo Oups

 If m_eType = PROJET Then
 Call ReceptionProjet
 Else
 Call ReceptionAchat
 End If

 Exit Sub

Oups:

 wOups "frmReceptionMec", "Reception", Err, Err.number, Err.Description
End Sub

Private Sub ReceptionProjet()

 On Error GoTo Oups

 Dim rstPiece As ADODB.Recordset
 Dim rstCopiePiece As ADODB.Recordset
 Dim rstProjet As ADODB.Recordset
 Dim rstModif As ADODB.Recordset
 Dim rstEmploye As ADODB.Recordset
 Dim sQuantite As String
 Dim sTotal As String
 Dim sProfit As String
 Dim bSkip As Boolean

 'Si il y a des enregistrements
 If lvwProjet.ListItems.count > 0 Then
  Set rstProjet = New ADODB.Recordset

  Call rstProjet.Open("SELECT Modification, Par, Profit FROM GrbProjetMec WHERE IDProjet = '" & txtnoprojet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

  If rstProjet.Fields("Modification") = False Then
  If lvwProjet.SelectedItem.ForeColor = COLOR_ORANGE Or lvwProjet.SelectedItem.ForeColor = COLOR_BLEU Then 'COLOR_ORANGE ou bleu
  sQuantite = InputBox("Quelle est la quantité recue?")

  sQuantite = Replace(sQuantite, ".", ",")

  sProfit = rstProjet.Fields("Profit")

  If sQuantite <> "" Then
 If IsNumeric(sQuantite) Then
 If CDbl(sQuantite) > 0 Then
 Set rstPiece = New ADODB.Recordset
 Set rstModif = New ADODB.Recordset
 Set rstEmploye = New ADODB.Recordset

 If CDbl(sQuantite) = CDbl(lvwProjet.SelectedItem.Text) Then
 Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & txtnoprojet.Text & "' AND Type = 'M' AND NuméroLigne = " & lvwProjet.SelectedItem.ListSubItems(I_COL_MANUFACTURIER).Tag, g_connData, adOpenDynamic, adLockOptimistic)

 rstPiece.Fields("Recu") = True
 rstPiece.Fields("Commandé") = False
 rstPiece.Fields("DateRéception") = txtDateReception.Text

 Call rstPiece.Update

 'Ajout aux modifs
 Call rstModif.Open("SELECT * FROM GrbProjet_Modif", g_connData, adOpenDynamic, adLockOptimistic)
 
 Call rstModif.AddNew

 Call rstEmploye.Open("SELECT noEmploye FROM GrbEmployés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 rstModif.Fields("Type") = "M"
 rstModif.Fields("IDProjet") = txtnoprojet.Text
 rstModif.Fields("noEmployé") = rstEmploye.Fields("noEmploye")
 rstModif.Fields("Date") = ConvertDate(Date)
1  rstModif.Fields("Heure") = Time
 rstModif.Fields("TypeModif") = "RECEPTION"

 Call rstEmploye.Close
 
 Call rstModif.Update
 
 Call rstModif.Close

 Call rstPiece.Close

 If UCase(lvwProjet.SelectedItem.SubItems(I_COL_DISTRIBUTEUR)) = "SOLUTION GRB INC." Then
 Call EnleverInventaireProjet(sQuantite)
 End If

 m_iIndexReception = lvwProjet.SelectedItem.Index

 Call RemplirListViewProjet(txtnoprojet.Text)
 Else
 If CDbl(sQuantite) < CDbl(lvwProjet.SelectedItem.Text) Then
 Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & txtnoprojet.Text & "' AND Type = 'M' AND NuméroLigne = " & lvwProjet.SelectedItem.ListSubItems(I_COL_MANUFACTURIER).Tag, g_connData, adOpenDynamic, adLockOptimistic)

 sTotal = rstPiece.Fields("Qté")

 rstPiece.Fields("Qté") = sQuantite
 
 'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
 rstPiece.Fields("Prix_Total") = Conversion(CStr(Round(Replace(rstPiece.Fields("Qté"), "*", vbNullString) * rstPiece.Fields("Prix_net") * CSng(sProfit), 2)), MODE_PAS_FORMAT)
 
 'Pour le profit, c'est le prix total - (prix net * quantité)
 rstPiece.Fields("Profit_argent") = Conversion(CStr(Round(rstPiece.Fields("Prix_total") - (rstPiece.Fields("Prix_net") * Replace(rstPiece.Fields("Qté"), "*", vbNullString)), 2)), MODE_PAS_FORMAT)

 rstPiece.Fields("Recu") = True
 rstPiece.Fields("Commandé") = False
 rstPiece.Fields("DateRéception") = txtDateReception.Text

 Call rstPiece.Update

 Set rstCopiePiece = New ADODB.Recordset

 Call rstCopiePiece.Open("SELECT * FROM GrbProjet_Pieces", g_connData, adOpenDynamic, adLockOptimistic)

 Call rstCopiePiece.AddNew

 rstCopiePiece.Fields("IDProjet") = rstPiece.Fields("IDProjet")
 rstCopiePiece.Fields("IDSection") = rstPiece.Fields("IDSection")
 rstCopiePiece.Fields("NumItem") = rstPiece.Fields("NumItem")
 rstCopiePiece.Fields("Qté") = CDbl(sTotal) - CDbl(sQuantite)
 rstCopiePiece.Fields("Desc_FR") = rstPiece.Fields("Desc_FR")
 rstCopiePiece.Fields("Desc_EN") = rstPiece.Fields("Desc_EN")
 rstCopiePiece.Fields("Manufact") = rstPiece.Fields("Manufact")
 rstCopiePiece.Fields("Prix_List") = rstPiece.Fields("Prix_List")
 rstCopiePiece.Fields("Escompte") = rstPiece.Fields("Escompte")
 rstCopiePiece.Fields("Prix_net") = rstPiece.Fields("Prix_net")
 rstCopiePiece.Fields("IDFRS") = rstPiece.Fields("IDFRS")

 'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
 rstCopiePiece.Fields("Prix_Total") = Conversion(CStr(Round(Replace(rstCopiePiece.Fields("Qté"), "*", vbNullString) * rstCopiePiece.Fields("Prix_net") * CSng(sProfit), 2)), MODE_PAS_FORMAT)
 
 rstCopiePiece.Fields("Profit_Pourcent") = rstPiece.Fields("Profit_Pourcent")
 
 'Pour le profit, c'est le prix total - (prix net * quantité)
 rstCopiePiece.Fields("Profit_argent") = Conversion(CStr(Round(rstCopiePiece.Fields("Prix_total") - (rstCopiePiece.Fields("Prix_net") * Replace(rstCopiePiece.Fields("Qté"), "*", vbNullString)), 2)), MODE_PAS_FORMAT)

 rstCopiePiece.Fields("SousSection") = rstPiece.Fields("SousSection")
 rstCopiePiece.Fields("OrdreSection") = rstPiece.Fields("OrdreSection")
4 rstCopiePiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne") + 1
4 rstCopiePiece.Fields("PrixOrigine") = rstPiece.Fields("PrixOrigine")
4 rstCopiePiece.Fields("Type") = rstPiece.Fields("Type")
4 rstCopiePiece.Fields("Visible") = rstPiece.Fields("Visible")
4 rstCopiePiece.Fields("Commandé") = True
4 rstCopiePiece.Fields("Quoté") = rstPiece.Fields("Quoté")
4 rstCopiePiece.Fields("Recu") = False
4 rstCopiePiece.Fields("Retour") = False
4 rstCopiePiece.Fields("NoRetour") = vbNullString
4 rstCopiePiece.Fields("CommandeAnnulée") = False
4 rstCopiePiece.Fields("DateRéception") = vbNullString
4  rstCopiePiece.Fields("Facturation") = rstPiece.Fields("Facturation")
4  rstCopiePiece.Fields("ID") = ""
4  rstCopiePiece.Fields("PieceExtra") = rstPiece.Fields("PieceExtra")
4  rstCopiePiece.Fields("DateCommande") = rstPiece.Fields("DateCommande")
4  rstCopiePiece.Fields("DateRequise") = rstPiece.Fields("DateRequise")
4  rstCopiePiece.Fields("MatérielInutile") = False

4  Call rstCopiePiece.Update

50 Call rstCopiePiece.Close
 Set rstCopiePiece = Nothing

 Call rstPiece.Close

 Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & txtnoprojet.Text & "' AND Type = 'M' AND NuméroLigne >= " & lvwProjet.SelectedItem.ListSubItems(I_COL_MANUFACTURIER).Tag + 1, g_connData, adOpenDynamic, adLockOptimistic)

 bSkip = False

 Do While Not rstPiece.EOF
 If ((rstPiece.Fields("NumItem") <> lvwProjet.SelectedItem.SubItems(I_COL_PIECE)) Or (rstPiece.Fields("Qté") <> CDbl(sTotal) - CDbl(sQuantite))) Or bSkip = True Then
 rstPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne") + 1

 Call rstPiece.Update
 Else
 bSkip = True
 End If

5  Call rstPiece.MoveNext
5  Loop

5  Call rstPiece.Close

 'Ajout aux modifs
5  Call rstModif.Open("SELECT * FROM GrbProjet_Modif", g_connData, adOpenDynamic, adLockOptimistic)

5  Call rstModif.AddNew

5  Call rstEmploye.Open("SELECT noEmploye FROM GrbEmployés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
5  rstModif.Fields("Type") = "M"
5  rstModif.Fields("IDProjet") = txtnoprojet.Text
60 rstModif.Fields("noEmployé") = rstEmploye.Fields("noEmploye")
  rstModif.Fields("Date") = ConvertDate(Date)
  rstModif.Fields("Heure") = Time
  rstModif.Fields("TypeModif") = "RECEPTION"

  Call rstEmploye.Close

  Call rstModif.Update

  Call rstModif.Close

  If UCase(lvwProjet.SelectedItem.SubItems(I_COL_DISTRIBUTEUR)) = "SOLUTION GRB INC." Then
  Call EnleverInventaireProjet(sQuantite)
  End If

  m_iIndexReception = lvwProjet.SelectedItem.Index

  Call RemplirListViewProjet(txtnoprojet.Text)
6  Else
6  Call MsgBox("La quantité est trop grande!", vbOKOnly, "Erreur")
6  End If
6  End If

6  Set rstPiece = Nothing
6  Set rstModif = Nothing
6  Set rstEmploye = Nothing
6  Else
70 Call MsgBox("La quantité doit être plus grande que 0!", vbOKOnly, "Erreur")
  End If
  Else
  Call MsgBox("Quantité non numérique!", vbOKOnly, "Erreur")
  End If
  End If
  End If
  Else
  Call MsgBox("Ce projet est en modification par " & rstProjet.Fields("Par") & "!", vbOKOnly, "Erreur")
  End If

  Call rstProjet.Close
  Set rstProjet = Nothing
   End If

   Exit Sub

Oups:

7  wOups "frmReceptionMec", "Reception", Err, Err.number, Err.Description
End Sub

Private Sub ReceptionAchat()

 On Error GoTo Oups

 Dim rstPiece As ADODB.Recordset
 Dim rstCopiePiece As ADODB.Recordset
 Dim rstAchat As ADODB.Recordset
 Dim sQuantite As String
 Dim sIDAchat As String
 Dim sTotal As String
 Dim bSkip As Boolean
 Dim iIndexAchat As Integer
 Dim iIDFRS As Integer

 'Si il y a des enregistrements
 If lvwProjet.ListItems.count > 0 Then
  sIDAchat = Left$(txtnoprojet.Text, 9)
  iIndexAchat = CInt(Right$(txtnoprojet.Text, 3))

  Set rstAchat = New ADODB.Recordset

  Call rstAchat.Open("SELECT Modification, Par FROM GrbAchat WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenDynamic, adLockOptimistic)

  If rstAchat.Fields("Modification") = False Then
  If lvwProjet.SelectedItem.ForeColor = COLOR_ORANGE Or lvwProjet.SelectedItem.ForeColor = COLOR_BLEU Then 'COLOR_ORANGE ou bleu
  sQuantite = InputBox("Quelle est la quantité reçue?")

  sQuantite = Replace(sQuantite, ".", ",")

 If sQuantite <> "" Then
 If IsNumeric(sQuantite) Then
 If CDbl(sQuantite) > 0 Then
 Set rstPiece = New ADODB.Recordset

 If CDbl(sQuantite) = CDbl(lvwProjet.SelectedItem.Text) Then
 Call rstPiece.Open("SELECT * FROM GrbAchat_Pieces WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat & " AND NuméroLigne = " & lvwProjet.SelectedItem.ListSubItems(I_COL_MANUFACTURIER).Tag, g_connData, adOpenDynamic, adLockOptimistic)

 rstPiece.Fields("Recu") = True
 rstPiece.Fields("Commandé") = False

 rstPiece.Fields("DateRéception") = txtDateReception.Text

 Call rstPiece.Update

 If (CInt(Right$(sIDAchat, 2)) >= 0 And CInt(Right$(sIDAchat, 2)) <= 12) Or rstPiece.Fields("IDFRS") = 71 Then
 Call AjouterInventaireAchat(sQuantite)
 End If

 Call rstPiece.Close

 m_iIndexReception = lvwProjet.SelectedItem.Index

 Call RemplirListViewAchat(txtnoprojet.Text)
 Else
 If CDbl(sQuantite) < CDbl(lvwProjet.SelectedItem.Text) Then
1  Call rstPiece.Open("SELECT * FROM GrbAchat_Pieces WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat & " AND NuméroLigne = " & lvwProjet.SelectedItem.ListSubItems(I_COL_MANUFACTURIER).Tag, g_connData, adOpenDynamic, adLockOptimistic)

 sTotal = rstPiece.Fields("Qté")

 rstPiece.Fields("Qté") = sQuantite
 
 'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
 rstPiece.Fields("Prix_Total") = Conversion(CStr(Round(Replace(rstPiece.Fields("Qté"), "*", vbNullString) * rstPiece.Fields("Prix_net"), 2)), MODE_PAS_FORMAT)

 rstPiece.Fields("Recu") = True
 rstPiece.Fields("Commandé") = False
 rstPiece.Fields("DateRéception") = txtDateReception.Text

 Call rstPiece.Update

 Set rstCopiePiece = New ADODB.Recordset

 Call rstCopiePiece.Open("SELECT * FROM GrbAchat_Pieces", g_connData, adOpenDynamic, adLockOptimistic)

 Call rstCopiePiece.AddNew

 rstCopiePiece.Fields("IDAchat") = rstPiece.Fields("IDAchat")
 rstCopiePiece.Fields("IndexAchat") = rstPiece.Fields("IndexAchat")
 rstCopiePiece.Fields("PIECE") = rstPiece.Fields("PIECE")
 rstCopiePiece.Fields("Qté") = CDbl(sTotal) - CDbl(sQuantite)
 rstCopiePiece.Fields("Desc_FR") = rstPiece.Fields("Desc_FR")
 rstCopiePiece.Fields("Desc_EN") = rstPiece.Fields("Desc_EN")
 rstCopiePiece.Fields("Manufact") = rstPiece.Fields("Manufact")
 rstCopiePiece.Fields("Prix_List") = rstPiece.Fields("Prix_List")
 rstCopiePiece.Fields("Escompte") = rstPiece.Fields("Escompte")
 rstCopiePiece.Fields("Prix_net") = rstPiece.Fields("Prix_net")
 rstCopiePiece.Fields("IDFRS") = rstPiece.Fields("IDFRS")

 'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
 rstCopiePiece.Fields("Prix_Total") = Conversion(CStr(Round(Replace(rstCopiePiece.Fields("Qté"), "*", vbNullString) * rstCopiePiece.Fields("Prix_net"), 2)), MODE_PAS_FORMAT)
 
 rstCopiePiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne") + 1
 rstCopiePiece.Fields("Type") = rstPiece.Fields("Type")
 rstCopiePiece.Fields("Commandé") = True
 rstCopiePiece.Fields("Recu") = False
 rstCopiePiece.Fields("Retour") = False
 rstCopiePiece.Fields("NoRetour") = vbNullString
 rstCopiePiece.Fields("DateRéception") = vbNullString
 rstCopiePiece.Fields("DateCommande") = rstPiece.Fields("DateCommande")
 rstCopiePiece.Fields("DateRequise") = rstPiece.Fields("DateRequise")

 Call rstCopiePiece.Update

 Call rstCopiePiece.Close
 Set rstCopiePiece = Nothing

 iIDFRS = 717

 Call rstPiece.Close

 Call rstPiece.Open("SELECT * FROM GrbAchat_Pieces WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat & " AND NuméroLigne >= " & lvwProjet.SelectedItem.ListSubItems(I_COL_MANUFACTURIER).Tag + 1, g_connData, adOpenDynamic, adLockOptimistic)

 bSkip = False

 Do While Not rstPiece.EOF
4 If ((rstPiece.Fields("PIECE") <> lvwProjet.SelectedItem.SubItems(I_COL_PIECE)) Or (rstPiece.Fields("Qté") <> CDbl(sTotal) - CDbl(sQuantite))) Or bSkip = True Then
4 rstPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne") + 1

4 Call rstPiece.Update
4 Else
4 bSkip = True
4 End If

4 Call rstPiece.MoveNext
4 Loop

4 Call rstPiece.Close

4 If CInt(Right$(sIDAchat, 2)) >= 0 And CInt(Right$(sIDAchat, 2)) <= 12 Or iIDFRS = 71 Then
4 Call AjouterInventaireAchat(sQuantite)
4  End If

4  m_iIndexReception = lvwProjet.SelectedItem.Index

4  Call RemplirListViewAchat(txtnoprojet.Text)
4  Else
4  Call MsgBox("La quantité est trop grande!", vbOKOnly, "Erreur")
4  End If
4  End If

4  Set rstPiece = Nothing
50 Else
 Call MsgBox("La quantité doit être plus grande que 0!", vbOKOnly, "Erreur")
 End If
 Else
 Call MsgBox("Quantité non numérique!", vbOKOnly, "Erreur")
 End If
 End If
 End If
 Else
 Call MsgBox("Ce projet est en modification par " & rstAchat.Fields("Par") & "!", vbOKOnly, "Erreur")
 End If

 Call rstAchat.Close
5  Set rstAchat = Nothing
5  End If

5  Exit Sub

Oups:

5  wOups "frmReceptionMec", "ReceptionAchat", Err, Err.number, Err.Description
End Sub

Private Sub EnleverInventaireProjet(ByVal sQuantite As String)

 On Error GoTo Oups

 Dim rstInventaire As ADODB.Recordset
 Dim rstPieceFRS As ADODB.Recordset
 Dim rstProjet As ADODB.Recordset

 If MsgBox("Voulez-vous modifier l'inventaire?", vbYesNo) = vbYes Then
 Set rstInventaire = New ADODB.Recordset

 Call rstInventaire.Open("SELECT * FROM GrbInventaireMec WHERE NoItem = '" & Replace(lvwProjet.SelectedItem.SubItems(I_COL_PIECE), "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If rstInventaire.EOF Then
 Call rstInventaire.AddNew
 
 rstInventaire.Fields("NoItem") = lvwProjet.SelectedItem.SubItems(I_COL_PIECE)
 rstInventaire.Fields("Description") = lvwProjet.SelectedItem.SubItems(I_COL_DESCRIPTION)
  rstInventaire.Fields("Manufacturier") = lvwProjet.SelectedItem.SubItems(I_COL_MANUFACTURIER)

  Call frmChoixQteBoite.Afficher(lvwProjet.SelectedItem.SubItems(I_COL_PIECE))

  rstInventaire.Fields("CommandeParBoite") = g_bQteBoite
  rstInventaire.Fields("QteBoite") = g_sQteBoite

  rstInventaire.Fields("Commentaires") = ""
  rstInventaire.Fields("QuantitéStock") = "0"
 
  Call frmChoixLocalisation.Afficher(MECANIQUE, lvwProjet.SelectedItem.SubItems(I_COL_PIECE))
 
  rstInventaire.Fields("Localisation") = g_sLocalisation
 rstInventaire.Fields("Minimum") = False
rstInventaire.Fields("QuantitéMinimum") = ""
 rstInventaire.Fields("Commande") = ""
 End If

 Set rstProjet = New ADODB.Recordset

 Call rstProjet.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & txtnoprojet.Text & "' AND NuméroLigne = " & lvwProjet.SelectedItem.ListSubItems(I_COL_MANUFACTURIER).Tag, g_connData, adOpenDynamic, adLockOptimistic)
 
 If rstInventaire.Fields("CommandeParBoite") = True Then
 rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) - (CDbl(sQuantite) * CDbl(rstInventaire.Fields("QteBoite"))), ".", ",")

 If rstProjet.Fields("Prix_List") <> "" Then
 If rstInventaire.Fields("QteBoite") <> "" Then
 rstInventaire.Fields("Prix Liste") = Replace(rstProjet.Fields("Prix_List") / rstInventaire.Fields("QteBoite"), ".", ",")
 Else
 rstInventaire.Fields("Prix Liste") = rstProjet.Fields("Prix_List")
 End If
 Else
 rstInventaire.Fields("Prix Liste") = "0"
 End If

 rstInventaire.Fields("Escompte") = rstProjet.Fields("Escompte")

 If rstInventaire.Fields("QteBoite") <> "" Then
1  rstInventaire.Fields("Prix net") = Replace(rstProjet.Fields("Prix_Net") / rstInventaire.Fields("QteBoite"), ".", ",")
 Else
 rstInventaire.Fields("Prix net") = rstProjet.Fields("Prix_Net")
 End If
 Else
 rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) - CDbl(sQuantite), ".", ",")

 If rstProjet.Fields("Prix_List") <> "" Then
 rstInventaire.Fields("Prix Liste") = rstProjet.Fields("Prix_List")
 Else
 rstInventaire.Fields("Prix Liste") = ""
 End If

 rstInventaire.Fields("Escompte") = rstProjet.Fields("Escompte")
 rstInventaire.Fields("Prix net") = rstProjet.Fields("Prix_Net")
End If
 
 Call rstInventaire.Update
 
Set rstPieceFRS = New ADODB.Recordset
 
 Call rstPieceFRS.Open("SELECT * FROM GrbPiecesFRS WHERE PIECE = '" & Replace(lvwProjet.SelectedItem.SubItems(I_COL_PIECE), "'", "''") & "' AND IDFRS = 717", g_connData, adOpenDynamic, adLockOptimistic)
 
If rstPieceFRS.EOF Then
 Call rstPieceFRS.AddNew
 
 rstPieceFRS.Fields("IDFRS") = 717
 rstPieceFRS.Fields("PIECE") = lvwProjet.SelectedItem.SubItems(I_COL_PIECE)
 rstPieceFRS.Fields("PERS_RESS") = 901
rstPieceFRS.Fields("ENTRER_PAR") = g_sInitiale
 rstPieceFRS.Fields("DeviseMonétaire") = "CAN"
 rstPieceFRS.Fields("Type") = "M"
 End If
 
 rstPieceFRS.Fields("PRIX_LIST") = rstProjet.Fields("Prix_List")
 rstPieceFRS.Fields("ESCOMPTE") = rstProjet.Fields("Escompte")
 rstPieceFRS.Fields("PRIX_NET") = rstProjet.Fields("Prix_net")
 rstPieceFRS.Fields("DATE") = txtDateReception.Text

 Call rstPieceFRS.Update
 
 Call rstPieceFRS.Close
 Set rstPieceFRS = Nothing
 
Call rstProjet.Close
 Set rstProjet = Nothing
 
If rstInventaire.Fields("CommandeParBoite") = True Then
 sQuantite = Replace(CDbl(sQuantite) * CDbl(rstInventaire.Fields("QteBoite")), ".", ",")
End If
 
 Call rstInventaire.Close
 Set rstInventaire = Nothing
 
 Call AjouterHistorique(sQuantite)
40 End If

Exit Sub

Oups:

4 wOups "frmReceptionMec", "EnleverInventaireProjet", Err, Err.number, Err.Description
End Sub

Private Sub AjouterInventaireAchat(ByVal sQuantite As String)

 On Error GoTo Oups

 Dim rstInventaire As ADODB.Recordset
 Dim rstPieceFRS As ADODB.Recordset
 Dim rstAchat As ADODB.Recordset
 Dim sIDAchat As String
 Dim iIndexAchat As Integer
 Dim iCompteur As Integer

 If MsgBox("Voulez-vous modifier l'inventaire?", vbYesNo) = vbYes Then
 Set rstInventaire = New ADODB.Recordset

 Call rstInventaire.Open("SELECT * FROM GrbInventaireMec WHERE NoItem = '" & Replace(lvwProjet.SelectedItem.SubItems(I_COL_PIECE), "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If rstInventaire.EOF Then
  Call rstInventaire.AddNew

  rstInventaire.Fields("NoItem") = lvwProjet.SelectedItem.SubItems(I_COL_PIECE)
  rstInventaire.Fields("Description") = lvwProjet.SelectedItem.SubItems(I_COL_DESCRIPTION)
  rstInventaire.Fields("Manufacturier") = lvwProjet.SelectedItem.SubItems(I_COL_MANUFACTURIER)

  Call frmChoixQteBoite.Afficher(lvwProjet.SelectedItem.SubItems(I_COL_PIECE))

  rstInventaire.Fields("CommandeParBoite") = g_bQteBoite
  rstInventaire.Fields("QteBoite") = g_sQteBoite

  rstInventaire.Fields("Commentaires") = ""
 rstInventaire.Fields("QuantitéStock") = "0"

Call frmChoixLocalisation.Afficher(MECANIQUE, lvwProjet.SelectedItem.SubItems(I_COL_PIECE))

 rstInventaire.Fields("Localisation") = g_sLocalisation
 rstInventaire.Fields("Minimum") = False
 rstInventaire.Fields("QuantitéMinimum") = ""
 rstInventaire.Fields("Commande") = ""
 End If
 
 sIDAchat = Left$(txtnoprojet.Text, 9)
 iIndexAchat = CInt(Right$(txtnoprojet.Text, 3))
 
 Set rstAchat = New ADODB.Recordset
 
 Call rstAchat.Open("SELECT * FROM GrbAchat_Pieces WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat & " AND NuméroLigne = " & lvwProjet.SelectedItem.ListSubItems(I_COL_MANUFACTURIER).Tag, g_connData, adOpenDynamic, adLockOptimistic)
 
 If rstInventaire.Fields("CommandeParBoite") = True Then
 If rstAchat.Fields("IDFRS") = 71 Then
 rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) - (CDbl(sQuantite) * CDbl(rstInventaire.Fields("QteBoite"))), ".", ",")
 Else
 rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) + (CDbl(sQuantite) * CDbl(rstInventaire.Fields("QteBoite"))), ".", ",")
 End If

 If rstAchat.Fields("Prix_List") <> "" Then
 If rstInventaire.Fields("QteBoite") <> "" Then
1  rstInventaire.Fields("Prix Liste") = Replace(rstAchat.Fields("Prix_List") / rstInventaire.Fields("QteBoite"), ".", ",")
 Else
 rstInventaire.Fields("Prix Liste") = rstAchat.Fields("Prix_List")
 End If
 Else
 rstInventaire.Fields("Prix Liste") = "0"
 End If

 rstInventaire.Fields("Escompte") = rstAchat.Fields("Escompte")

 If rstInventaire.Fields("QteBoite") <> "" Then
 rstInventaire.Fields("Prix net") = Replace(rstAchat.Fields("Prix_Net") / rstInventaire.Fields("QteBoite"), ".", ",")
 Else
 rstInventaire.Fields("Prix net") = rstAchat.Fields("Prix_Net")
 End If
Else
 If rstAchat.Fields("IDFRS") = 71 Then
 rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) - CDbl(sQuantite), ".", ",")
 Else
 rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) + CDbl(sQuantite), ".", ",")
 End If

 If rstAchat.Fields("Prix_List") <> "" Then
 rstInventaire.Fields("Prix Liste") = rstAchat.Fields("Prix_List")
 Else
 rstInventaire.Fields("Prix Liste") = "0"
 End If

 rstInventaire.Fields("Escompte") = rstAchat.Fields("Escompte")
 rstInventaire.Fields("Prix net") = rstAchat.Fields("Prix_Net")
 End If
 
 Call rstInventaire.Update

 Set rstPieceFRS = New ADODB.Recordset
 
 Call rstPieceFRS.Open("SELECT * FROM GrbPiecesFRS WHERE PIECE = '" & Replace(lvwProjet.SelectedItem.SubItems(I_COL_PIECE), "'", "''") & "' AND IDFRS = 717", g_connData, adOpenDynamic, adLockOptimistic)
 
 If rstPieceFRS.EOF Then
 Call rstPieceFRS.AddNew
 
 rstPieceFRS.Fields("IDFRS") = 717
 rstPieceFRS.Fields("PIECE") = lvwProjet.SelectedItem.SubItems(I_COL_PIECE)
 rstPieceFRS.Fields("PERS_RESS") = 901
 rstPieceFRS.Fields("ENTRER_PAR") = g_sInitiale
 rstPieceFRS.Fields("DeviseMonétaire") = "CAN"
 rstPieceFRS.Fields("Type") = "M"
 End If

 rstPieceFRS.Fields("PRIX_LIST") = rstAchat.Fields("Prix_List")
 rstPieceFRS.Fields("ESCOMPTE") = rstAchat.Fields("Escompte")
rstPieceFRS.Fields("PRIX_NET") = rstAchat.Fields("Prix_net")
4 rstPieceFRS.Fields("DATE") = txtDateReception.Text

4 Call rstPieceFRS.Update
 
4 Call rstPieceFRS.Close
4 Set rstPieceFRS = Nothing
 
4 Call rstAchat.Close
4 Set rstAchat = Nothing
 
4 If rstInventaire.Fields("CommandeParBoite") = True Then
4 sQuantite = CDbl(sQuantite) * CDbl(rstInventaire.Fields("QteBoite"))
4 End If
 
4 Call rstInventaire.Close
4 Set rstInventaire = Nothing
 
4  Call AjouterHistorique(sQuantite)
4  End If

4  Exit Sub

Oups:

4  wOups "frmReceptionMec", "AjouterInventaireAchat", Err, Err.number, Err.Description
End Sub

Private Sub AjouterHistorique(ByVal sQuantite As String)

 On Error GoTo Oups

 Dim rstHist As ADODB.Recordset

 Set rstHist = New ADODB.Recordset

 Call rstHist.Open("SELECT * FROM GrbInventaireMecModif", g_connData, adOpenDynamic, adLockOptimistic)

 Call rstHist.AddNew

 rstHist.Fields("Date") = txtDateReception.Text
 rstHist.Fields("IDProjet") = txtnoprojet.Text
 rstHist.Fields("NoItem") = lvwProjet.SelectedItem.SubItems(I_COL_PIECE)

 If m_eType = ACHAT Then
 If lvwProjet.SelectedItem.ListSubItems(I_COL_DISTRIBUTEUR).Tag = 71 Then
 rstHist.Fields("Quantité") = "-" & sQuantite
  Else
  rstHist.Fields("Quantité") = sQuantite
  End If
  Else
  rstHist.Fields("Quantité") = "-" & sQuantite
  End If

  rstHist.Fields("User") = g_sInitiale

  Call rstHist.Update

10 Call rstHist.Close
Set rstHist = Nothing

Exit Sub

Oups:

wOups "frmReceptionMec", "AjouterHistorique", Err, Err.number, Err.Description
End Sub

Private Sub SupprimerHistorique(Optional ByVal sQuantite As String = "")

 On Error GoTo Oups

 Dim rstHist As ADODB.Recordset

 Set rstHist = New ADODB.Recordset

 If m_eType = ACHAT Then
 If sQuantite <> "" Then
 If lvwProjet.SelectedItem.ListSubItems(I_COL_DISTRIBUTEUR).Tag = 71 Then
 Call rstHist.Open("SELECT * FROM GrbInventaireMecModif WHERE [Date] = '" & lvwProjet.SelectedItem.SubItems(I_COL_DATE_RECEPTION) & "' AND IDProjet = '" & txtnoprojet.Text & "' AND NoItem = '" & Replace(lvwProjet.SelectedItem.SubItems(I_COL_PIECE), "'", "''") & "' AND Quantité = '-" & sQuantite & "'", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstHist.Open("SELECT * FROM GrbInventaireMecModif WHERE [Date] = '" & lvwProjet.SelectedItem.SubItems(I_COL_DATE_RECEPTION) & "' AND IDProjet = '" & txtnoprojet.Text & "' AND NoItem = '" & Replace(lvwProjet.SelectedItem.SubItems(I_COL_PIECE), "'", "''") & "' AND Quantité = '" & sQuantite & "'", g_connData, adOpenDynamic, adLockOptimistic)
 End If
 Else
  If lvwProjet.SelectedItem.ListSubItems(I_COL_DISTRIBUTEUR).Tag = 71 Then
  Call rstHist.Open("SELECT * FROM GrbInventaireMecModif WHERE [Date] = '" & lvwProjet.SelectedItem.SubItems(I_COL_DATE_RECEPTION) & "' AND IDProjet = '" & txtnoprojet.Text & "' AND NoItem = '" & Replace(lvwProjet.SelectedItem.SubItems(I_COL_PIECE), "'", "''") & "' AND Quantité = '-" & lvwProjet.SelectedItem.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
  Else
  Call rstHist.Open("SELECT * FROM GrbInventaireMecModif WHERE [Date] = '" & lvwProjet.SelectedItem.SubItems(I_COL_DATE_RECEPTION) & "' AND IDProjet = '" & txtnoprojet.Text & "' AND NoItem = '" & Replace(lvwProjet.SelectedItem.SubItems(I_COL_PIECE), "'", "''") & "' AND Quantité = '" & lvwProjet.SelectedItem.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
  End If
  End If
  Else
  If sQuantite <> "" Then
 Call rstHist.Open("SELECT * FROM GrbInventaireMecModif WHERE [Date] = '" & lvwProjet.SelectedItem.SubItems(I_COL_DATE_RECEPTION) & "' AND IDProjet = '" & txtnoprojet.Text & "' AND NoItem = '" & Replace(lvwProjet.SelectedItem.SubItems(I_COL_PIECE), "'", "''") & "' AND Quantité = '" & "-" & sQuantite & "'", g_connData, adOpenDynamic, adLockOptimistic)
1 Else
 Call rstHist.Open("SELECT * FROM GrbInventaireMecModif WHERE [Date] = '" & lvwProjet.SelectedItem.SubItems(I_COL_DATE_RECEPTION) & "' AND IDProjet = '" & txtnoprojet.Text & "' AND NoItem = '" & Replace(lvwProjet.SelectedItem.SubItems(I_COL_PIECE), "'", "''") & "' AND Quantité = '" & "-" & lvwProjet.SelectedItem.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 End If
End If

If Not rstHist.EOF Then
 Call rstHist.Delete
End If

Call rstHist.Close
Set rstHist = Nothing

Exit Sub

Oups:

wOups "frmReceptionMec", "SupprimerHistorique", Err, Err.number, Err.Description
End Sub

Private Sub Cmdfermer_Click()

 On Error GoTo Oups
 
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmReceptionMec", "cmdFermer_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 Dim iCompteur As Integer

 Call Unload(frmChoixProjSoum)

 txtDateReception.Text = ConvertDate(Date)
 txtDateRequise.Text = ConvertDate(Date)

 If m_sNoProjet <> "" Then
 cmbType.ListIndex = 0

 For iCompteur = 0 To cmbNoProjet.ListCount - 1
 If cmbNoProjet.LIST(iCompteur) = m_sNoProjet Then
 cmbNoProjet.ListIndex = iCompteur

 Exit For
  End If
  Next
  Else
  If m_sNoAchat <> "" Then
  cmbType.ListIndex = 1

  For iCompteur = 0 To cmbNoProjet.ListCount - 1
  If cmbNoProjet.LIST(iCompteur) = m_sNoAchat Then
  cmbNoProjet.ListIndex = iCompteur

 Exit For
 End If
 Next
 Else
 cmbType.ListIndex = 0
 End If
End If

Exit Sub

Oups:

wOups "frmReceptionMec", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboProjet()

 On Error GoTo Oups

 'Rempli le combo des soumissions
 Dim rstProjet As ADODB.Recordset
 
 'Il faut vider le combo avant de le remplir
 Call cmbNoProjet.Clear
 
 'Ouvre le recordset selon le type
 Set rstProjet = New ADODB.Recordset
 
 Call rstProjet.Open("SELECT IDProjet FROM GrbProjetMec ORDER BY IDProjet DESC", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Tant que ce n'est pas la fin des enregistrements
 Do While Not rstProjet.EOF
 Call cmbNoProjet.AddItem(rstProjet.Fields("IDProjet"))

 Call rstProjet.MoveNext
 Loop
 
 Call rstProjet.Close
 Set rstProjet = Nothing
 
 'Si le combo n'est pas vide, on sélectionne l'item voulu ou le premier
  If cmbNoProjet.ListCount > 0 Then
 'On sélectionne le premier
  cmbNoProjet.ListIndex = 0
  End If

  Exit Sub

Oups:

  wOups "frmReceptionMec", "RemplirComboProjet", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboAchat()

 On Error GoTo Oups

 'Rempli le combo des soumissions
 Dim rstAchat As ADODB.Recordset
 
 'Il faut vider le combo avant de le remplir
 Call cmbNoProjet.Clear
 
 'Ouvre le recordset selon le type
 Set rstAchat = New ADODB.Recordset
 
 Call rstAchat.Open("SELECT IDAchat, IndexAchat FROM GrbAchat WHERE Type = 'M' ORDER BY IDAchat DESC, IndexAchat DESC", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Tant que ce n'est pas la fin des enregistrements
 Do While Not rstAchat.EOF
 Call cmbNoProjet.AddItem(rstAchat.Fields("IDAchat") & "-" & Right$("000" & rstAchat.Fields("IndexAchat"), 3))

 Call rstAchat.MoveNext
 Loop
 
 Call rstAchat.Close
 Set rstAchat = Nothing
 
 'Si le combo n'est pas vide, on sélectionne l'item voulu ou le premier
  If cmbNoProjet.ListCount > 0 Then
 'On sélectionne le premier
  cmbNoProjet.ListIndex = 0
  End If

  Exit Sub

Oups:

  wOups "frmReceptionMec", "RemplirComboAchat", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewProjet(ByVal sNoProjet As String)

 On Error GoTo Oups

 'Remplis les pièces de la soumission avec la BD
 Dim rstProjet As ADODB.Recordset
 Dim rstSection As ADODB.Recordset
 Dim rstFRS As ADODB.Recordset
 Dim itmProjet As ListItem
 Dim bPremierEnr As Boolean
 Dim iOrdreSection As Integer
 Dim sSousSection As String
 Dim lColor As Long
 
 Call lvwProjet.ListItems.Clear
 
 bPremierEnr = True
 
  Set rstProjet = New ADODB.Recordset
  Set rstFRS = New ADODB.Recordset
  Set rstSection = New ADODB.Recordset
 
  Call rstProjet.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & sNoProjet & "' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

  Do While Not rstProjet.EOF
  Set itmProjet = lvwProjet.ListItems.Add
 
 'Si c'est le premier enregistrement, il faut ajouter la section et la sous-section
  If bPremierEnr = True Then
  iOrdreSection = rstProjet.Fields("OrdreSection")
 sSousSection = rstProjet.Fields("SousSection")
 
 'Pour avoir le nom de la section
Call rstSection.Open("SELECT NomSectionFR FROM GrbSoumProjSectionMec WHERE IDSection = " & rstProjet.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
 
 'Ajout du nom de la section
 If Not IsNull(rstSection.Fields("NomSectionFR")) Then
 itmProjet.SubItems(I_COL_PIECE) = rstSection.Fields("NomSectionFR")
 Else
 itmProjet.SubItems(I_COL_PIECE) = vbNullString
 End If
 
 itmProjet.ListSubItems(I_COL_PIECE).Bold = True
 
 Call rstSection.Close
 
 Set itmProjet = lvwProjet.ListItems.Add
 
 'Ajout du nom de la sous-section
 If sSousSection = "PAS DE SOUS-SECTION" Then
 itmProjet.SubItems(I_COL_DESCRIPTION) = vbNullString
 Else
 itmProjet.SubItems(I_COL_DESCRIPTION) = sSousSection
 End If
 
 'Le tag ne peut pas être remplis si la colonne est vide
 itmProjet.SubItems(I_COL_MANUFACTURIER) = " "
 
 itmProjet.ListSubItems(I_COL_DESCRIPTION).Bold = True
 
 itmProjet.Tag = rstProjet.Fields("IDSection")
 
 Set itmProjet = lvwProjet.ListItems.Add
 
1  bPremierEnr = False
 Else
 'Si c'est pas le premier enregistrement, il faut vérifier avec l'ancienne section
 If iOrdreSection <> rstProjet.Fields("OrdreSection") Then
 iOrdreSection = rstProjet.Fields("OrdreSection")
 
 Call rstSection.Open("SELECT NomSectionFR FROM GrbSoumProjSectionMec WHERE IDSection = " & rstProjet.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
 
 If Not IsNull(rstSection.Fields("NomSectionFR")) Then
 itmProjet.SubItems(I_COL_PIECE) = rstSection.Fields("NomSectionFR")
 Else
 itmProjet.SubItems(I_COL_PIECE) = vbNullString
 End If
 
 itmProjet.ListSubItems(I_COL_PIECE).Bold = True
 
 Call rstSection.Close
 
 Set itmProjet = lvwProjet.ListItems.Add
 
 sSousSection = rstProjet.Fields("SousSection")
 
 If sSousSection = "PAS DE SOUS-SECTION" Then
 itmProjet.SubItems(I_COL_DESCRIPTION) = vbNullString
 Else
 itmProjet.SubItems(I_COL_DESCRIPTION) = rstProjet.Fields("SousSection")
 End If
 
 'Le tag ne peut pas être remplis si la colonne est vide
 itmProjet.SubItems(I_COL_MANUFACTURIER) = " "
 
 itmProjet.ListSubItems(I_COL_DESCRIPTION).Bold = True
 
 Set itmProjet = lvwProjet.ListItems.Add
Else
 'il faut vérifier avec l'ancienne sous-section
 If sSousSection <> rstProjet.Fields("SousSection") Then
 sSousSection = rstProjet.Fields("SousSection")
 
 If sSousSection = "PAS DE SOUS-SECTION" Then
 itmProjet.SubItems(I_COL_DESCRIPTION) = vbNullString
 Else
 itmProjet.SubItems(I_COL_DESCRIPTION) = sSousSection
 End If
 
 itmProjet.ListSubItems(I_COL_DESCRIPTION).Bold = True
 
 'Le tag ne peut pas être remplis si la colonne est vide
 itmProjet.SubItems(I_COL_MANUFACTURIER) = " "
 
 Set itmProjet = lvwProjet.ListItems.Add
 End If
 End If
End If
 
 If rstProjet.Fields("Commandé") = True Then
 lColor = COLOR_ORANGE 'COLOR_ORANGE
 Else
 If rstProjet.Fields("Recu") = True Then
 lColor = COLOR_GRIS
 Else
4 If rstProjet.Fields("Retour") = True Then
4 lColor = COLOR_ROUGE
4 Else
4 lColor = COLOR_NOIR
4 End If
4 End If
4 End If

4 itmProjet.Tag = rstProjet.Fields("IDSection")

 'Quantité
4 If Not IsNull(rstProjet.Fields("Qté")) Then
4 itmProjet.Text = rstProjet.Fields("Qté")
4 Else
4  itmProjet.Text = vbNullString
4  End If

4  itmProjet.ForeColor = lColor
 
 'Numéro d'item
4  If Not IsNull(rstProjet.Fields("NumItem")) Then
4  itmProjet.SubItems(I_COL_PIECE) = rstProjet.Fields("NumItem")
4  Else
4  itmProjet.SubItems(I_COL_PIECE) = vbNullString
4  End If
 
50 itmProjet.ListSubItems(I_COL_PIECE).ForeColor = lColor
 
 'On met le nom de la sous-section dans le tag du numéro d'item
5 itmProjet.ListSubItems(I_COL_PIECE).Tag = rstProjet.Fields("SousSection")
 
 'Description en francais
 If Not IsNull(rstProjet.Fields("Desc_FR")) Then
 itmProjet.SubItems(I_COL_DESCRIPTION) = rstProjet.Fields("Desc_FR")
 Else
 itmProjet.SubItems(I_COL_DESCRIPTION) = vbNullString
 End If
 
 itmProjet.ListSubItems(I_COL_DESCRIPTION).ForeColor = lColor
 
 'On met la description en anglais dans le tag de la description en francais
 If Not IsNull(rstProjet.Fields("Desc_EN")) Then
 itmProjet.ListSubItems(I_COL_DESCRIPTION).Tag = rstProjet.Fields("Desc_EN")
 Else
 itmProjet.ListSubItems(I_COL_DESCRIPTION).Tag = vbNullString
5  End If
 
 'Fabricant
5  If Not IsNull(rstProjet.Fields("Manufact")) Then
5  itmProjet.SubItems(I_COL_MANUFACTURIER) = rstProjet.Fields("Manufact")
5  Else
5  itmProjet.SubItems(I_COL_MANUFACTURIER) = vbNullString
5  End If
 
5  itmProjet.ListSubItems(I_COL_MANUFACTURIER).ForeColor = lColor
 
 'On met l'ordre de la section dans le tag du fabricant
5  itmProjet.ListSubItems(I_COL_MANUFACTURIER).Tag = rstProjet.Fields("NuméroLigne")
 
 'Fournisseur
60 If Not IsNull(rstProjet.Fields("IDFRS")) And rstProjet.Fields("IDFRS") > 0 Then
  If itmProjet.SubItems(I_COL_PIECE) <> "Texte" Then
  Call rstFRS.Open("SELECT NomFournisseur FROM GrbFournisseur WHERE IDFRS = " & rstProjet.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
 
 'On affiche le nom dans la colonne
  itmProjet.SubItems(I_COL_DISTRIBUTEUR) = rstFRS.Fields("NomFournisseur")
 
  itmProjet.ListSubItems(I_COL_DISTRIBUTEUR).ForeColor = lColor
 
 'On affiche l'Id dans le tag
  itmProjet.ListSubItems(I_COL_DISTRIBUTEUR).Tag = rstProjet.Fields("IDFRS")
 
  Call rstFRS.Close
  End If
  Else
  itmProjet.SubItems(I_COL_DISTRIBUTEUR) = vbNullString
  End If

  If Not IsNull(rstProjet.Fields("DateRéception")) Then
6  itmProjet.SubItems(I_COL_DATE_RECEPTION) = rstProjet.Fields("DateRéception")
6  Else
6  itmProjet.SubItems(I_COL_DATE_RECEPTION) = vbNullString
6  End If

6  itmProjet.ListSubItems(I_COL_DATE_RECEPTION).ForeColor = lColor
 
6  If Not IsNull(rstProjet.Fields("DateCommande")) Then
6  itmProjet.SubItems(I_COL_DATE_COMMANDE) = rstProjet.Fields("DateCommande")
6  Else
70 itmProjet.SubItems(I_COL_DATE_COMMANDE) = vbNullString
  End If

  itmProjet.ListSubItems(I_COL_DATE_COMMANDE).ForeColor = lColor
 
  If Not IsNull(rstProjet.Fields("DateRequise")) Then
  itmProjet.SubItems(I_COL_DATE_REQUISE) = rstProjet.Fields("DateRequise")
  Else
  itmProjet.SubItems(I_COL_DATE_REQUISE) = vbNullString
  End If

  itmProjet.ListSubItems(I_COL_DATE_REQUISE).ForeColor = lColor
 
  Call rstProjet.MoveNext
  Loop
 
  Call rstProjet.Close
   Set rstProjet = Nothing

   Set rstFRS = Nothing
7  Set rstSection = Nothing

7  If m_iIndexReception > 0 Then
7  lvwProjet.ListItems(m_iIndexReception).Selected = True

7  Call lvwProjet.SetFocus

7  Call lvwProjet.SelectedItem.EnsureVisible
7  End If

80 Exit Sub

Oups:

80 wOups "frmReceptionMec", "RemplirListViewProjet", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewAchat(ByVal sNoProjet As String)

 On Error GoTo Oups

 'Remplis les pièces de la soumission avec la BD
 Dim rstAchat As ADODB.Recordset
 Dim rstFRS As ADODB.Recordset
 Dim itmAchat As ListItem
 Dim lColor As Long
 Dim sNoAchat As String
 Dim iIndexAchat As Integer

 sNoAchat = Left$(sNoProjet, 9)

 iIndexAchat = CInt(Right$(sNoProjet, 3))
 
 Call lvwProjet.ListItems.Clear
 
 Set rstAchat = New ADODB.Recordset
  Set rstFRS = New ADODB.Recordset
 
  Call rstAchat.Open("SELECT * FROM GrbAchat_Pieces WHERE IDAchat = '" & sNoAchat & "' AND IndexAchat = " & Right$("000" & iIndexAchat, 3) & " ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

  Do While Not rstAchat.EOF
  Set itmAchat = lvwProjet.ListItems.Add
 
  If rstAchat.Fields("Commandé") = True Then
  lColor = COLOR_ORANGE 'COLOR_ORANGE
  Else
  If rstAchat.Fields("Recu") = True Then
 lColor = COLOR_GRIS
Else
 If rstAchat.Fields("Retour") = True Then
 lColor = COLOR_ROUGE
 Else
 lColor = COLOR_NOIR
 End If
 End If
 End If
 
 'Quantité
 If Not IsNull(rstAchat.Fields("Qté")) Then
 itmAchat.Text = rstAchat.Fields("Qté")
 Else
 itmAchat.Text = vbNullString
 End If

 itmAchat.ForeColor = lColor
 
 'Numéro d'item
 If Not IsNull(rstAchat.Fields("PIECE")) Then
 itmAchat.SubItems(I_COL_PIECE) = rstAchat.Fields("PIECE")
 Else
 itmAchat.SubItems(I_COL_PIECE) = vbNullString
1  End If
 
 itmAchat.ListSubItems(I_COL_PIECE).ForeColor = lColor
 
 'Description en francais
 If Not IsNull(rstAchat.Fields("Desc_FR")) Then
 itmAchat.SubItems(I_COL_DESCRIPTION) = rstAchat.Fields("Desc_FR")
 Else
 itmAchat.SubItems(I_COL_DESCRIPTION) = vbNullString
 End If
 
 itmAchat.ListSubItems(I_COL_DESCRIPTION).ForeColor = lColor
 
 'On met la description en anglais dans le tag de la description en francais
 If Not IsNull(rstAchat.Fields("DESC_EN")) Then
 itmAchat.ListSubItems(I_COL_DESCRIPTION).Tag = rstAchat.Fields("Desc_EN")
 Else
 itmAchat.ListSubItems(I_COL_DESCRIPTION).Tag = vbNullString
 End If
 
 'Fabricant
If Not IsNull(rstAchat.Fields("Manufact")) Then
 itmAchat.SubItems(I_COL_MANUFACTURIER) = rstAchat.Fields("Manufact")
Else
 itmAchat.SubItems(I_COL_MANUFACTURIER) = vbNullString
End If
 
 itmAchat.ListSubItems(I_COL_MANUFACTURIER).ForeColor = lColor

 'On met l'ordre de la section dans le tag du fabricant
itmAchat.ListSubItems(I_COL_MANUFACTURIER).Tag = rstAchat.Fields("NuméroLigne")
 
 'Fournisseur
 If Not IsNull(rstAchat.Fields("IDFRS")) And rstAchat.Fields("IDFRS") > 0 Then
 If itmAchat.SubItems(I_COL_PIECE) <> "Texte" Then
 Call rstFRS.Open("SELECT NomFournisseur FROM GrbFournisseur WHERE IDFRS = " & rstAchat.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
 
 'On affiche le nom dans la colonne
 itmAchat.SubItems(I_COL_DISTRIBUTEUR) = rstFRS.Fields("NomFournisseur")
 
 itmAchat.ListSubItems(I_COL_DISTRIBUTEUR).ForeColor = lColor
 
 'On affiche l'Id dans le tag
 itmAchat.ListSubItems(I_COL_DISTRIBUTEUR).Tag = rstAchat.Fields("IDFRS")
 
 Call rstFRS.Close
 End If
 Else
 itmAchat.SubItems(I_COL_DISTRIBUTEUR) = vbNullString
 End If

 If Not IsNull(rstAchat.Fields("DateRéception")) Then
 itmAchat.SubItems(I_COL_DATE_RECEPTION) = rstAchat.Fields("DateRéception")
Else
 itmAchat.SubItems(I_COL_DATE_RECEPTION) = vbNullString
End If

 itmAchat.ListSubItems(I_COL_DATE_RECEPTION).ForeColor = lColor

If Not IsNull(rstAchat.Fields("DateCommande")) Then
 itmAchat.SubItems(I_COL_DATE_COMMANDE) = rstAchat.Fields("DateCommande")
 Else
 itmAchat.SubItems(I_COL_DATE_COMMANDE) = vbNullString
End If

4 itmAchat.ListSubItems(I_COL_DATE_COMMANDE).ForeColor = lColor

4 If Not IsNull(rstAchat.Fields("DateRequise")) Then
4 itmAchat.SubItems(I_COL_DATE_REQUISE) = rstAchat.Fields("DateRequise")
4 Else
4 itmAchat.SubItems(I_COL_DATE_REQUISE) = vbNullString
4 End If

4 itmAchat.ListSubItems(I_COL_DATE_REQUISE).ForeColor = lColor

4 Call rstAchat.MoveNext
4 Loop
 
4 Call rstAchat.Close
4 Set rstAchat = Nothing

4  Set rstFRS = Nothing

4  If m_iIndexReception > 0 Then
4  lvwProjet.ListItems(m_iIndexReception).Selected = True

4  Call lvwProjet.SetFocus

4  Call lvwProjet.SelectedItem.EnsureVisible
4  End If

4  Exit Sub

Oups:

4  wOups "frmReceptionMec", "RemplirListViewAchat", Err, Err.number, Err.Description
End Sub

Private Sub lvwProjet_ItemClick(ByVal Item As MSComctlLib.ListItem)

 On Error GoTo Oups

 Call VerifierBoutonAnnuler

 Exit Sub

Oups:

 wOups "frmReceptionMec", "lvwProjet_ItemClick", Err, Err.number, Err.Description
End Sub

Private Sub VerifierBoutonAnnuler()

 On Error GoTo Oups
 
 If lvwProjet.ListItems.count > 0 Then
 If lvwProjet.SelectedItem.ForeColor = COLOR_GRIS Or lvwProjet.SelectedItem.ForeColor = COLOR_BLEU Then 'Gris ou bleu
 cmdAnnuler.Enabled = True
 Else
 cmdAnnuler.Enabled = False
 End If
 Else
 cmdAnnuler.Enabled = False
 End If

 Exit Sub

Oups:

  wOups "frmReceptionMec", "VerifierBoutonAnnuler", Err, Err.number, Err.Description
End Sub

Public Sub Afficher(ByVal sUserID As String)
 
 On Error GoTo Oups

 m_sUserID = sUserID

 Call Me.Show

 Exit Sub

Oups:

 wOups "frmReceptionMec", "Afficher", Err, Err.number, Err.Description
End Sub

Public Sub AfficherProjet(ByVal sUserID As String, ByVal sNoProjet As String)

 On Error GoTo Oups

 m_sUserID = sUserID

 m_sNoProjet = sNoProjet

 Call Me.Show(vbModal)

 Exit Sub

Oups:

 wOups "frmReceptionMec", "AfficherProjet", Err, Err.number, Err.Description
End Sub

Public Sub AfficherAchat(ByVal sUserID As String, ByVal sNoAchat As String)

 On Error GoTo Oups

 m_sUserID = sUserID

 m_sNoAchat = sNoAchat

 Call Me.Show(vbModal)

 Exit Sub

Oups:

 wOups "frmReceptionMec", "AfficherAchat", Err, Err.number, Err.Description
End Sub


Private Sub lvwProjet_KeyDown(KeyCode As Integer, Shift As Integer)

 On Error GoTo Oups

 If KeyCode = vbKeyReturn Then
 If m_eType = PROJET Then
 Call ReceptionProjet
 Else
 Call ReceptionAchat
 End If
 End If

 Exit Sub

Oups:

 wOups "frmReceptionMec", "lvwProjet_KeyDown", Err, Err.number, Err.Description
End Sub

Private Sub mvwDateRequise_DateClick(ByVal DateClicked As Date)

 On Error GoTo Oups

 txtDateRequise.Text = ConvertDate(DateClicked)

 'Enlever le calendrier
 mvwDateRequise.Visible = False

 Exit Sub

Oups:

 wOups "frmReceptionMec", "mvwDateRequise_DateClick", Err, Err.number, Err.Description
End Sub

Private Sub mvwDateRequise_LostFocus()

 On Error GoTo Oups

 mvwDateRequise.Visible = False

 Exit Sub

Oups:

 wOups "frmReceptionMec", "mvwDateRequise_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mvwReception_DateClick(ByVal DateClicked As Date)

 On Error GoTo Oups

 txtDateReception.Text = ConvertDate(DateClicked)

 'Enlever le calendrier
 mvwReception.Visible = False

 Exit Sub

Oups:

 wOups "frmReceptionMec", "mvwReception_DateClick", Err, Err.number, Err.Description
End Sub

Private Sub mvwReception_LostFocus()

 On Error GoTo Oups

 mvwReception.Visible = False

 Exit Sub

Oups:

 wOups "frmReceptionMec", "mvwReception_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub Form_Click()

 On Error GoTo Oups

 mvwReception.Visible = False
 mvwDateRequise.Visible = False

 Exit Sub

Oups:

 wOups "frmReceptionMec", "Form_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdDate_Click()

 On Error GoTo Oups

 'Ouverture du calendrier
 If txtDateReception.Text <> vbNullString Then
 mvwReception.Value = txtDateReception.Text
 Else
 mvwReception.Value = Date
 End If

 mvwReception.Visible = True

 Call mvwReception.SetFocus

 Exit Sub

Oups:

 wOups "frmReceptionMec", "cmdDate_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmbType_Click()

 On Error GoTo Oups

 If cmbType.ListIndex = 0 Then
 m_eType = PROJET

 Call RemplirComboProjet
 Else
 m_eType = ACHAT

 Call RemplirComboAchat
 End If

 If fraPiecesNonRecues.Visible = True Then
 If m_eType = ACHAT Then
 chkProjetAchat.Caption = "No achat : "
  Else
  chkProjetAchat.Caption = "No projet : "
  End If
  End If

  Exit Sub

Oups:

  wOups "frmReceptionMec", "cmbType_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdImprimer_Click()

 On Error GoTo Oups

 Call ImprimerReception

 Exit Sub

Oups:

 wOups "frmReceptionMec", "cmdImprimer_Click", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerReception()

 On Error GoTo Oups

 If m_eType = ACHAT Then
 Call frmChoixDateImpressionReception.Afficher(txtnoprojet.Text, MECANIQUE, ACHAT)
 Else
 Call frmChoixDateImpressionReception.Afficher(txtnoprojet.Text, MECANIQUE, PROJET)
 End If

 Exit Sub

Oups:

 wOups "frmReceptionMec", "ImprimerReception", Err, Err.number, Err.Description
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 'CTRL-I pour imprimer
 
 On Error GoTo Oups

 If Shift = vbCtrlMask Then
 If KeyCode = vbKeyI Then
 Call ImprimerReception
 End If
 End If

 Exit Sub

Oups:

 wOups "frmReceptionMec", "Form_KeyDown", Err, Err.number, Err.Description
End Sub
