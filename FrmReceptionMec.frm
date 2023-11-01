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
   Picture         =   "FrmReceptionMec.frx":2CFA
   ScaleHeight     =   7755
   ScaleWidth      =   11925
   StartUpPosition =   2  'CenterScreen
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
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   0
      Appearance      =   1
      StartOfWeek     =   90243073
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
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   0
         Appearance      =   1
         StartOfWeek     =   90243073
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
      ItemData        =   "FrmReceptionMec.frx":206AC
      Left            =   3120
      List            =   "FrmReceptionMec.frx":206B6
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
Private Const I_COL_QUANTITE       As Integer = 0
Private Const I_COL_PIECE          As Integer = 1
Private Const I_COL_DESCRIPTION    As Integer = 2
Private Const I_COL_MANUFACTURIER  As Integer = 3
Private Const I_COL_DISTRIBUTEUR   As Integer = 4
Private Const I_COL_DATE_RECEPTION As Integer = 5
Private Const I_COL_DATE_COMMANDE  As Integer = 6
Private Const I_COL_DATE_REQUISE   As Integer = 7

Private Const I_LVW_PROJET         As Integer = 0
Private Const I_LVW_QUANTITE       As Integer = 1
Private Const I_LVW_PIECE          As Integer = 2
Private Const I_LVW_DESCRIPTION    As Integer = 3
Private Const I_LVW_FOURNISSEUR    As Integer = 4
Private Const I_LVW_DATE_COMMANDE  As Integer = 5
Private Const I_LVW_DATE_REQUISE   As Integer = 6

Private Enum enumType
  PROJET = 0
  ACHAT = 1
End Enum

Private m_sUserID         As String
Private m_sNoProjet       As String
Private m_sNoAchat        As String
Private m_eType           As enumType
Private m_iIndexReception As Integer

Private Sub chkDateRequise_Click()

5       On Error GoTo AfficherErreur

10      If chkDateRequise.Value = vbChecked Then
15        txtDateRequise.Enabled = True
20        cmdDateRequise.Enabled = True
25      Else
30        txtDateRequise.Enabled = False
35        cmdDateRequise.Enabled = False
40      End If

45      Exit Sub

AfficherErreur:

50      woups "frmReceptionMec", "chkDateRequise_Click", Err, Erl
End Sub

Private Sub chkProjetAchat_Click()

5       On Error GoTo AfficherErreur

10      If chkProjetAchat.Value = vbChecked Then
15        txtProjetAchat.Enabled = True
20      Else
25        txtProjetAchat.Enabled = False
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmReceptionMec", "chkProjetAchat_Click", Err, Erl
End Sub

Private Sub cmbNoProjet_KeyUp(KeyCode As Integer, Shift As Integer)
  
5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer
  
15      For iCompteur = 0 To cmbNoProjet.ListCount - 1
20        If UCase(cmbNoProjet.LIST(iCompteur)) = UCase(cmbNoProjet.Text) Then
25          cmbNoProjet.ListIndex = iCompteur
      
30          Exit For
35        End If
40      Next

45      Exit Sub

AfficherErreur:

50      woups "frmReceptionMec", "cmbNoProjet_KeyUp", Err, Erl
End Sub

Private Sub cmbNoProjet_Click()

5       On Error GoTo AfficherErreur

10      Dim rstProjAchat As ADODB.Recordset
15      Dim sNumero      As String

20      Set rstProjAchat = New ADODB.Recordset

25      sNumero = txtnoprojet.Text

30      If m_eType = ACHAT Then
35        Call rstProjAchat.Open("SELECT * FROM GRB_Achat WHERE IDAchat = '" & Left$(cmbNoProjet.Text, 9) & "' AND IndexAchat = " & CInt(Right$(cmbNoProjet.Text, 3)), g_connData, adOpenDynamic, adLockOptimistic)
40      Else
45        Call rstProjAchat.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & cmbNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
50      End If

55      If rstProjAchat.Fields("Modification") = True Then
60        If m_eType = ACHAT Then
65          Call MsgBox("Cet achat est en modification par " & rstProjAchat.Fields("Par") & "!", vbOKOnly, "Erreur")
70        Else
75          Call MsgBox("Ce projet est en modification par " & rstProjAchat.Fields("Par") & "!", vbOKOnly, "Erreur")
80        End If

85        Call rstProjAchat.Close
90        Set rstProjAchat = Nothing

95        cmbNoProjet.Text = sNumero

100       Exit Sub
105     End If
  
110     Screen.MousePointer = vbHourglass

115     m_iIndexReception = 0

120     txtnoprojet.Text = cmbNoProjet.Text
  
125     If m_eType = PROJET Then
          'Rempli les valeurs du projet sélectionné
130       Call RemplirListViewProjet(txtnoprojet.Text)
135     Else
140       Call RemplirListViewAchat(txtnoprojet.Text)
145     End If

150     Call VerifierBoutonAnnuler

155     Screen.MousePointer = vbDefault

160     Exit Sub

AfficherErreur:

165     woups "frmReceptionMec", "cmbNoProjet_Click", Err, Erl
End Sub

Private Sub cmdAfficher_Click()

5       On Error GoTo AfficherErreur

10      Dim bRemplir As Boolean

15      If chkProjetAchat.Value = vbChecked Then
20        If Trim$(txtProjetAchat.Text) <> "" Then
25          If m_eType = ACHAT Then
30            If Len(Trim$(txtProjetAchat.Text)) = 13 Then
35              bRemplir = True
40            Else
45              Call MsgBox("Format de numéro d'achat incorrect!", vbOKOnly, "Erreur")
50            End If
55          Else
60            If Len(Trim$(txtProjetAchat.Text)) = 9 Then
65              bRemplir = True
70            Else
75              Call MsgBox("Format de numéro de projet incorrect!", vbOKOnly, "Erreur")
80            End If
85          End If
90        Else
95          If m_eType = ACHAT Then
100           Call MsgBox("Le numéro de l'achat est obligatoire!", vbOKOnly, "Erreur")
105         Else
110           Call MsgBox("Le numéro de projet est obligatoire!", vbOKOnly, "Erreur")
115         End If
120       End If
125     Else
130       bRemplir = True
135     End If

140     If bRemplir = True Then
145       Screen.MousePointer = vbHourglass

150       Call RemplirListePiecesNonRecues

155       Screen.MousePointer = vbDefault
160     End If

165     Exit Sub

AfficherErreur:

170     woups "frmReceptionMec", "cmdAfficher_Click", Err, Erl
End Sub

Private Sub cmdAnnuler_Click()

5       On Error GoTo AfficherErreur

10      If m_eType = PROJET Then
15        Call AnnulerReceptionProjet
20      Else
25        Call AnnulerReceptionAchat
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmReceptionMec", "cmdAnnuler_Click", Err, Erl
End Sub

Private Sub AnnulerReceptionProjet()

5       On Error GoTo AfficherErreur

10      Dim rstProjet  As ADODB.Recordset
15      Dim rstPiece   As ADODB.Recordset
20      Dim rstModif   As ADODB.Recordset
25      Dim rstEmploye As ADODB.Recordset
    
        'S'il y a des enregistrements
30      If lvwProjet.ListItems.count > 0 Then
35        Set rstProjet = New ADODB.Recordset

40        Call rstProjet.Open("SELECT Modification, Par FROM GRB_ProjetMec WHERE IDProjet = '" & txtnoprojet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

45        If rstProjet.Fields("Modification") = False Then
50          Set rstPiece = New ADODB.Recordset

55          Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & txtnoprojet.Text & "' AND Type = 'M' AND NuméroLigne = " & lvwProjet.SelectedItem.ListSubItems(I_COL_MANUFACTURIER).Tag, g_connData, adOpenDynamic, adLockOptimistic)

65          rstPiece.Fields("Recu") = False
70          rstPiece.Fields("Commandé") = True

75          rstPiece.Fields("DateRéception") = ""

80          Call rstPiece.Update

            'Ajout aux modifs
85          Set rstModif = New ADODB.Recordset
            
90          Call rstModif.Open("SELECT * FROM GRB_Projet_Modif", g_connData, adOpenDynamic, adLockOptimistic)
      
95          Call rstModif.AddNew

100         Set rstEmploye = New ADODB.Recordset

105         Call rstEmploye.Open("SELECT noEmploye FROM GRB_Employés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)
      
110         rstModif.Fields("Type") = "M"
115         rstModif.Fields("IDProjet") = txtnoprojet.Text
120         rstModif.Fields("noEmployé") = rstEmploye.Fields("noEmploye")
125         rstModif.Fields("Date") = ConvertDate(Date)
130         rstModif.Fields("Heure") = Time
135         rstModif.Fields("TypeModif") = "RECEPTION"

140         Call rstEmploye.Close
145         Set rstEmploye = Nothing
      
150         Call rstModif.Update
    
155         Call rstModif.Close
160         Set rstModif = Nothing

165         Call rstPiece.Close
170         Set rstPiece = Nothing

175         If UCase(lvwProjet.SelectedItem.SubItems(I_COL_DISTRIBUTEUR)) = "SOLUTION GRB INC." Then
180           Call AjouterInventaireProjet
185         End If

190         Call SupprimerHistorique

195         m_iIndexReception = lvwProjet.SelectedItem.Index

200         Call RemplirListViewProjet(txtnoprojet.Text)
205       Else
210         Call MsgBox("Ce projet est en modification par " & rstProjet.Fields("Par") & "!", vbOKOnly, "Erreur")
215       End If

220       Call rstProjet.Close
225       Set rstProjet = Nothing
230     End If

235     Exit Sub

AfficherErreur:

240     woups "frmReceptionMec", "AnnulerReceptionProjet", Err, Erl
End Sub

Private Sub AnnulerReceptionAchat()

5       On Error GoTo AfficherErreur

10      Dim rstPiece    As ADODB.Recordset
15      Dim rstAchat    As ADODB.Recordset
20      Dim sIDAchat    As String
25      Dim iIndexAchat As Integer
    
        'S'il y a des enregistrements
30      If lvwProjet.ListItems.count > 0 Then
35        sIDAchat = Left$(txtnoprojet.Text, 9)
40        iIndexAchat = CInt(Right$(txtnoprojet.Text, 3))

45        Set rstAchat = New ADODB.Recordset

50        Call rstAchat.Open("SELECT Modification, Par FROM GRB_Achat WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenDynamic, adLockOptimistic)

55        If rstAchat.Fields("Modification") = False Then
60          Set rstPiece = New ADODB.Recordset

65          Call rstPiece.Open("SELECT * FROM GRB_Achat_Pieces WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat & " And NuméroLigne = " & lvwProjet.SelectedItem.ListSubItems(I_COL_MANUFACTURIER).Tag, g_connData, adOpenDynamic, adLockOptimistic)

75          rstPiece.Fields("Recu") = False
80          rstPiece.Fields("Commandé") = True

85          rstPiece.Fields("DateRéception") = ""

90          Call rstPiece.Update

95          If (CInt(Right$(sIDAchat, 2)) >= 0 And CInt(Right$(sIDAchat, 2)) <= 12) Or rstPiece.Fields("IDFRS") = 717 Then
100           Call EnleverInventaireAchat
105         End If

110         Call rstPiece.Close
115         Set rstPiece = Nothing

120         m_iIndexReception = lvwProjet.SelectedItem.Index

125         Call RemplirListViewAchat(txtnoprojet.Text)
130       Else
135         Call MsgBox("Ce projet est en modification par " & rstAchat.Fields("Par") & "!", vbOKOnly, "Erreur")
140       End If

145       Call rstAchat.Close
150       Set rstAchat = Nothing
155     End If

160     Exit Sub

AfficherErreur:

165     woups "frmReceptionMec", "AnnulerReceptionAchat", Err, Erl
End Sub

Private Sub EnleverInventaireAchat()
 
5       On Error GoTo AfficherErreur
 
10      Dim rstInventaire As ADODB.Recordset
15      Dim sQuantite     As String

20      If MsgBox("Voulez-vous modifier l'inventaire?", vbYesNo) = vbYes Then
25        Set rstInventaire = New ADODB.Recordset

30        Call rstInventaire.Open("SELECT * FROM GRB_InventaireMec WHERE NoItem = '" & lvwProjet.SelectedItem.SubItems(I_COL_PIECE) & "'", g_connData, adOpenDynamic, adLockOptimistic)

35        If Not rstInventaire.EOF Then
40          If rstInventaire.Fields("CommandeParBoite") = True Then
45            If lvwProjet.SelectedItem.ListSubItems(I_COL_DISTRIBUTEUR).Tag = 717 Then
50              rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) + (CDbl(lvwProjet.SelectedItem.Text) * CDbl(rstInventaire.Fields("QteBoite"))), "", ",")
55            Else
60              rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) - (CDbl(lvwProjet.SelectedItem.Text) * CDbl(rstInventaire.Fields("QteBoite"))), ".", ",")
65            End If

70            sQuantite = CDbl(lvwProjet.SelectedItem.Text) * CDbl(rstInventaire.Fields("QteBoite"))
75          Else
80            If lvwProjet.SelectedItem.ListSubItems(I_COL_DISTRIBUTEUR).Tag = 717 Then
85              rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) + CDbl(lvwProjet.SelectedItem.Text), "", ",")
90            Else
95              rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) - CDbl(lvwProjet.SelectedItem.Text), ".", ",")
100           End If

105           sQuantite = CDbl(lvwProjet.SelectedItem.Text)
110         End If
 
115         Call rstInventaire.Update
120       End If

125       Call rstInventaire.Close
130       Set rstInventaire = Nothing
 
135       Call SupprimerHistorique(sQuantite)
140     End If

145     Exit Sub

AfficherErreur:
 
150     woups "frmReceptionMec", "EnleverInventaireAchat", Err, Erl
End Sub

Private Sub AjouterInventaireProjet()

5       On Error GoTo AfficherErreur

10      Dim rstInventaire As ADODB.Recordset
15      Dim sQuantite     As String

20      If MsgBox("Voulez-vous modifier l'inventaire?", vbYesNo) = vbYes Then
25        Set rstInventaire = New ADODB.Recordset

30        Call rstInventaire.Open("SELECT * FROM GRB_InventaireMec WHERE NoItem = '" & Replace(lvwProjet.SelectedItem.SubItems(I_COL_PIECE), "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)

35        If Not rstInventaire.EOF Then
40          If rstInventaire.Fields("CommandeParBoite") = True Then
45            rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) + (CDbl(lvwProjet.SelectedItem.Text) * CDbl(rstInventaire.Fields("QteBoite"))), ".", ",")

50            sQuantite = CDbl(lvwProjet.SelectedItem.Text) * CDbl(rstInventaire.Fields("QteBoite"))
55          Else
60            rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) + CDbl(lvwProjet.SelectedItem.Text), ".", ",")

65            sQuantite = CDbl(lvwProjet.SelectedItem.Text)
70          End If

75          Call rstInventaire.Update
80        End If

85        Call rstInventaire.Close
90        Set rstInventaire = Nothing

95        Call SupprimerHistorique(sQuantite)
100     End If

105     Exit Sub

AfficherErreur:

110     woups "frmReceptionMec", "AjouterInventaireProjet", Err, Erl
End Sub

Private Sub cmdDateRequise_Click()

5       On Error GoTo AfficherErreur

        'Ouverture du calendrier
10      If txtDateRequise.Text <> vbNullString Then
15        mvwDateRequise.Value = txtDateRequise.Text
20      Else
25        mvwDateRequise.Value = Date
35      End If

40      mvwDateRequise.Visible = True

45      Call mvwDateRequise.SetFocus

50      Exit Sub

AfficherErreur:

55      woups "frmReceptionMec", "cmdDateRequise_Click", Err, Erl
End Sub

Private Sub cmdFermerPieces_Click()

5       On Error GoTo AfficherErreur

10      fraPiecesNonRecues.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmReceptionMec", "cmdFermerPieces_Click", Err, Erl
End Sub

Private Sub cmdImprimerPieces_Click()

5       On Error GoTo AfficherErreur

10      Dim rstPiece As ADODB.Recordset

15      Set rstPiece = New ADODB.Recordset

20      If m_eType = PROJET Then
25        If chkDateRequise.Value = vbChecked And chkProjetAchat.Value = vbChecked Then
30          Call rstPiece.Open("SELECT GRB_Projet_Pieces.*, GRB_Fournisseur.NomFournisseur FROM GRB_Projet_Pieces INNER JOIN GRB_Fournisseur ON GRB_Projet_Pieces.IDFRS = GRB_Fournisseur.IDFRS WHERE Type = 'M' AND Commandé = True AND DateRequise <= '" & txtDateRequise.Text & "' AND IDProjet = '" & txtProjetAchat.Text & "' AND DateRequise <> '' ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
35        Else
40          If chkDateRequise.Value = vbChecked Then
45            Call rstPiece.Open("SELECT GRB_Projet_Pieces.*, GRB_Fournisseur.NomFournisseur FROM GRB_Projet_Pieces INNER JOIN GRB_Fournisseur ON GRB_Projet_Pieces.IDFRS = GRB_Fournisseur.IDFRS WHERE Type = 'M' AND Commandé = True AND DateRequise <= '" & txtDateRequise.Text & "' AND DateRequise <> '' ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
50          Else
55            If chkProjetAchat.Value = vbChecked Then
60              Call rstPiece.Open("SELECT GRB_Projet_Pieces.*, GRB_Fournisseur.NomFournisseur FROM GRB_Projet_Pieces INNER JOIN GRB_Fournisseur ON GRB_Projet_Pieces.IDFRS = GRB_Fournisseur.IDFRS WHERE Type = 'M' AND Commandé = True AND IDProjet = '" & txtProjetAchat.Text & "' ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
65            Else
70              Call rstPiece.Open("SELECT GRB_Projet_Pieces.*, GRB_Fournisseur.NomFournisseur FROM GRB_Projet_Pieces INNER JOIN GRB_Fournisseur ON GRB_Projet_Pieces.IDFRS = GRB_Fournisseur.IDFRS WHERE Type = 'M' AND Commandé = True ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
75            End If
80          End If
85        End If
90      Else
95        If chkDateRequise.Value = vbChecked And chkProjetAchat.Value = vbChecked Then
100         Call rstPiece.Open("SELECT (GRB_Achat_Pieces.IDAchat &  '-' & RIGHT('00' & GRB_Achat_Pieces.IndexAchat,3)) AS NoAchat, GRB_Achat_Pieces.*, GRB_Fournisseur.NomFournisseur FROM GRB_Achat_Pieces INNER JOIN GRB_Fournisseur ON GRB_Achat_Pieces.IDFRS = GRB_Fournisseur.IDFRS WHERE LEN(IDAchat) = 9 AND Commandé = True AND DateRequise <= '" & txtDateRequise.Text & "' AND IDAchat = '" & Left$(txtProjetAchat.Text, 9) & "' AND IndexAchat = " & CInt(Right$(txtProjetAchat.Text, 3)) & " AND DateRequise <> '' ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
105       Else
110         If chkDateRequise.Value = vbChecked Then
115           Call rstPiece.Open("SELECT (GRB_Achat_Pieces.IDAchat &  '-' & RIGHT('00' & GRB_Achat_Pieces.IndexAchat,3)) AS NoAchat, GRB_Achat_Pieces.*, GRB_Fournisseur.NomFournisseur FROM GRB_Achat_Pieces INNER JOIN GRB_Fournisseur ON GRB_Achat_Pieces.IDFRS = GRB_Fournisseur.IDFRS WHERE LEN(IDAchat) = 9 AND Commandé = True AND DateRequise <= '" & txtDateRequise.Text & "' AND DateRequise <> '' ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
120         Else
125           If chkProjetAchat.Value = vbChecked Then
130             Call rstPiece.Open("SELECT (GRB_Achat_Pieces.IDAchat &  '-' & RIGHT('00' & GRB_Achat_Pieces.IndexAchat,3)) AS NoAchat, GRB_Achat_Pieces.*, GRB_Fournisseur.NomFournisseur FROM GRB_Achat_Pieces INNER JOIN GRB_Fournisseur ON GRB_Achat_Pieces.IDFRS = GRB_Fournisseur.IDFRS WHERE LEN(IDAchat) = 9 AND Commandé = True AND IDAchat = '" & Left$(txtProjetAchat.Text, 9) & "' AND IndexAchat = " & CInt(Right$(txtProjetAchat.Text, 3)) & " ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
135           Else
140             Call rstPiece.Open("SELECT (GRB_Achat_Pieces.IDAchat &  '-' & RIGHT('00' & GRB_Achat_Pieces.IndexAchat,3)) AS NoAchat, GRB_Achat_Pieces.*, GRB_Fournisseur.NomFournisseur FROM GRB_Achat_Pieces INNER JOIN GRB_Fournisseur ON GRB_Achat_Pieces.IDFRS = GRB_Fournisseur.IDFRS WHERE LEN(IDAchat) = 9 AND Commandé = True ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
145           End If
150         End If
155       End If
160     End If

165     Set DR_BackOrder.DataSource = rstPiece

170     DR_BackOrder.Orientation = rptOrientLandscape

175     If m_eType = PROJET Then
180       DR_BackOrder.Sections("Section4").Controls("lblTitre").Caption = "Projets mécaniques : Pièces non reçues"

185       DR_BackOrder.Sections("Section4").Controls("lblTitreProjetAchat").Caption = "Projet : "
190     Else
195       DR_BackOrder.Sections("Section4").Controls("lblTitre").Caption = "Achats mécaniques : Pièces non reçues"

200       DR_BackOrder.Sections("Section4").Controls("lblTitreProjetAchat").Caption = "Achat : "
205     End If

210     DR_BackOrder.Sections("Section4").Controls("lblDate").Caption = txtDateRequise.Text

215     DR_BackOrder.Sections("Section4").Controls("lblProjetAchat").Caption = txtProjetAchat.Text

220     If m_eType = ACHAT Then
225       DR_BackOrder.Sections("Section2").Controls("lblTitreNoProjet").Caption = "# Achat"

230       DR_BackOrder.Sections("Section1").Controls("txtNoProjAchat").DataField = "NoAchat"
235       DR_BackOrder.Sections("Section1").Controls("txtNoItem").DataField = "PIECE"
240     End If

245     Call DR_BackOrder.Show(vbModal)

250     Call rstPiece.Close
255     Set rstPiece = Nothing

260     Exit Sub

AfficherErreur:

265     woups "frmReceptionMec", "cmdImprimerPieces_Click", Err, Erl
End Sub

Private Sub cmdNonRecu_Click()

5       On Error GoTo AfficherErreur

10      Call lvwPieces.ListItems.Clear

15      If m_eType = ACHAT Then
20        chkProjetAchat.Caption = "No achat : "
25      Else
30        chkProjetAchat.Caption = "No projet : "
35      End If

40      fraPiecesNonRecues.Visible = True

45      Exit Sub

AfficherErreur:

50      woups "frmReceptionMec", "cmdNonRecu_Click", Err, Erl
End Sub

Private Sub RemplirListePiecesNonRecues()

5       On Error GoTo AfficherErreur

10      Dim itmPiece As ListItem
15      Dim rstPiece As ADODB.Recordset

20      Call lvwPieces.ListItems.Clear

25      Set rstPiece = New ADODB.Recordset

30      If m_eType = PROJET Then
35        If chkDateRequise.Value = vbChecked And chkProjetAchat.Value = vbChecked Then
40          Call rstPiece.Open("SELECT GRB_Projet_Pieces.*, GRB_Fournisseur.NomFournisseur FROM GRB_Projet_Pieces INNER JOIN GRB_Fournisseur ON GRB_Projet_Pieces.IDFRS = GRB_Fournisseur.IDFRS WHERE Type = 'M' AND Commandé = True AND DateRequise <= '" & txtDateRequise.Text & "' AND IDProjet = '" & txtProjetAchat.Text & "' AND DateRequise <> '' ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
45        Else
50          If chkDateRequise.Value = vbChecked Then
55            Call rstPiece.Open("SELECT GRB_Projet_Pieces.*, GRB_Fournisseur.NomFournisseur FROM GRB_Projet_Pieces INNER JOIN GRB_Fournisseur ON GRB_Projet_Pieces.IDFRS = GRB_Fournisseur.IDFRS WHERE Type = 'M' AND Commandé = True AND DateRequise <= '" & txtDateRequise.Text & "' AND DateRequise <> '' ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
60          Else
65            If chkProjetAchat.Value = vbChecked Then
70              Call rstPiece.Open("SELECT GRB_Projet_Pieces.*, GRB_Fournisseur.NomFournisseur FROM GRB_Projet_Pieces INNER JOIN GRB_Fournisseur ON GRB_Projet_Pieces.IDFRS = GRB_Fournisseur.IDFRS WHERE Type = 'M' AND Commandé = True AND IDProjet = '" & txtProjetAchat.Text & "' ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
75            Else
80              Call rstPiece.Open("SELECT GRB_Projet_Pieces.*, GRB_Fournisseur.NomFournisseur FROM GRB_Projet_Pieces INNER JOIN GRB_Fournisseur ON GRB_Projet_Pieces.IDFRS = GRB_Fournisseur.IDFRS WHERE Type = 'M' AND Commandé = True ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
85            End If
90          End If
95        End If
100     Else
105       If chkDateRequise.Value = vbChecked And chkProjetAchat.Value = vbChecked Then
110         Call rstPiece.Open("SELECT GRB_Achat_Pieces.*, GRB_Fournisseur.NomFournisseur FROM GRB_Achat_Pieces INNER JOIN GRB_Fournisseur ON GRB_Achat_Pieces.IDFRS = GRB_Fournisseur.IDFRS WHERE LEN(IDAchat) = 9 AND Commandé = True AND DateRequise <= '" & txtDateRequise.Text & "' AND IDAchat = '" & Left$(txtProjetAchat.Text, 9) & "' AND IndexAchat = " & CInt(Right$(txtProjetAchat.Text, 3)) & " AND DateRequise <> '' ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
115       Else
120         If chkDateRequise.Value = vbChecked Then
125           Call rstPiece.Open("SELECT GRB_Achat_Pieces.*, GRB_Fournisseur.NomFournisseur FROM GRB_Achat_Pieces INNER JOIN GRB_Fournisseur ON GRB_Achat_Pieces.IDFRS = GRB_Fournisseur.IDFRS WHERE LEN(IDAchat) = 9 AND Commandé = True AND DateRequise <= '" & txtDateRequise.Text & "' AND DateRequise <> '' ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
130         Else
135           If chkProjetAchat.Value = vbChecked Then
140             Call rstPiece.Open("SELECT GRB_Achat_Pieces.*, GRB_Fournisseur.NomFournisseur FROM GRB_Achat_Pieces INNER JOIN GRB_Fournisseur ON GRB_Achat_Pieces.IDFRS = GRB_Fournisseur.IDFRS WHERE LEN(IDAchat) = 9 AND Commandé = True AND IDAchat = '" & Left$(txtProjetAchat.Text, 9) & "' AND IndexAchat = " & CInt(Right$(txtProjetAchat.Text, 3)) & " ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
145           Else
150             Call rstPiece.Open("SELECT GRB_Achat_Pieces.*, GRB_Fournisseur.NomFournisseur FROM GRB_Achat_Pieces INNER JOIN GRB_Fournisseur ON GRB_Achat_Pieces.IDFRS = GRB_Fournisseur.IDFRS WHERE LEN(IDAchat) = 9 AND Commandé = True ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
155           End If
160         End If
165       End If
170     End If

175     Do While Not rstPiece.EOF
180       Set itmPiece = lvwPieces.ListItems.Add

185       If m_eType = PROJET Then
190         itmPiece.Text = rstPiece.Fields("IDProjet")
195       Else
200         itmPiece.Text = rstPiece.Fields("IDAchat") & "-" & Right$("00" & rstPiece.Fields("IndexAchat"), 3)
205       End If

210       itmPiece.SubItems(I_LVW_QUANTITE) = rstPiece.Fields("Qté")

215       If m_eType = PROJET Then
220         itmPiece.SubItems(I_LVW_PIECE) = rstPiece.Fields("NumItem")
225       Else
230         itmPiece.SubItems(I_LVW_PIECE) = rstPiece.Fields("PIECE")
235       End If

240       itmPiece.SubItems(I_LVW_DESCRIPTION) = rstPiece.Fields("Desc_FR")
245       itmPiece.SubItems(I_LVW_FOURNISSEUR) = rstPiece.Fields("NomFournisseur")

250       If Not IsNull(rstPiece.Fields("DateCommande")) Then
255         itmPiece.SubItems(I_LVW_DATE_COMMANDE) = rstPiece.Fields("DateCommande")
260       Else
265         itmPiece.SubItems(I_LVW_DATE_COMMANDE) = ""
270       End If

275       If Not IsNull(rstPiece.Fields("DateRequise")) Then
280         itmPiece.SubItems(I_LVW_DATE_REQUISE) = rstPiece.Fields("DateRequise")
285       Else
290         itmPiece.SubItems(I_LVW_DATE_REQUISE) = ""
295       End If

300       Call rstPiece.MoveNext
305     Loop

310     Call rstPiece.Close
315     Set rstPiece = Nothing

320     Exit Sub

AfficherErreur:

325     woups "frmReceptionMec", "RemplirListePiecesNonRecues", Err, Erl
End Sub

Private Sub lvwProjet_Click()

5       On Error GoTo AfficherErreur

10      Call VerifierBoutonAnnuler

15      Exit Sub

AfficherErreur:

20      woups "frmReceptionMec", "lvwProjet_Click", Err, Erl
End Sub

Private Sub lvwProjet_DblClick()

5       On Error GoTo AfficherErreur

10      If m_eType = PROJET Then
15        Call ReceptionProjet
20      Else
25        Call ReceptionAchat
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmReceptionMec", "Reception", Err, Erl
End Sub

Private Sub ReceptionProjet()

5       On Error GoTo AfficherErreur

10      Dim rstPiece      As ADODB.Recordset
15      Dim rstCopiePiece As ADODB.Recordset
20      Dim rstProjet     As ADODB.Recordset
25      Dim rstModif      As ADODB.Recordset
30      Dim rstEmploye    As ADODB.Recordset
35      Dim sQuantite     As String
40      Dim sTotal        As String
45      Dim sProfit       As String
50      Dim bSkip         As Boolean

        'Si il y a des enregistrements
55      If lvwProjet.ListItems.count > 0 Then
60        Set rstProjet = New ADODB.Recordset

65        Call rstProjet.Open("SELECT Modification, Par, Profit FROM GRB_ProjetMec WHERE IDProjet = '" & txtnoprojet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

70        If rstProjet.Fields("Modification") = False Then
75          If lvwProjet.SelectedItem.ForeColor = COLOR_ORANGE Or lvwProjet.SelectedItem.ForeColor = COLOR_BLEU Then 'COLOR_ORANGE ou bleu
80            sQuantite = InputBox("Quelle est la quantité recue?")

85            sQuantite = Replace(sQuantite, ".", ",")

90            sProfit = rstProjet.Fields("Profit")

95            If sQuantite <> "" Then
100             If IsNumeric(sQuantite) Then
105               If CDbl(sQuantite) > 0 Then
110                 Set rstPiece = New ADODB.Recordset
115                 Set rstModif = New ADODB.Recordset
120                 Set rstEmploye = New ADODB.Recordset

125                 If CDbl(sQuantite) = CDbl(lvwProjet.SelectedItem.Text) Then
130                   Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & txtnoprojet.Text & "' AND Type = 'M' AND NuméroLigne = " & lvwProjet.SelectedItem.ListSubItems(I_COL_MANUFACTURIER).Tag, g_connData, adOpenDynamic, adLockOptimistic)

140                   rstPiece.Fields("Recu") = True
145                   rstPiece.Fields("Commandé") = False
150                   rstPiece.Fields("DateRéception") = txtDateReception.Text

155                   Call rstPiece.Update

                      'Ajout aux modifs
160                   Call rstModif.Open("SELECT * FROM GRB_Projet_Modif", g_connData, adOpenDynamic, adLockOptimistic)
      
165                   Call rstModif.AddNew

170                   Call rstEmploye.Open("SELECT noEmploye FROM GRB_Employés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)
      
175                   rstModif.Fields("Type") = "M"
180                   rstModif.Fields("IDProjet") = txtnoprojet.Text
185                   rstModif.Fields("noEmployé") = rstEmploye.Fields("noEmploye")
190                   rstModif.Fields("Date") = ConvertDate(Date)
195                   rstModif.Fields("Heure") = Time
200                   rstModif.Fields("TypeModif") = "RECEPTION"

205                   Call rstEmploye.Close
      
210                   Call rstModif.Update
    
215                   Call rstModif.Close

220                   Call rstPiece.Close

225                   If UCase(lvwProjet.SelectedItem.SubItems(I_COL_DISTRIBUTEUR)) = "SOLUTION GRB INC." Then
230                     Call EnleverInventaireProjet(sQuantite)
235                   End If

240                   m_iIndexReception = lvwProjet.SelectedItem.Index

245                   Call RemplirListViewProjet(txtnoprojet.Text)
250                 Else
255                   If CDbl(sQuantite) < CDbl(lvwProjet.SelectedItem.Text) Then
260                     Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & txtnoprojet.Text & "' AND Type = 'M' AND NuméroLigne = " & lvwProjet.SelectedItem.ListSubItems(I_COL_MANUFACTURIER).Tag, g_connData, adOpenDynamic, adLockOptimistic)

265                     sTotal = rstPiece.Fields("Qté")

270                     rstPiece.Fields("Qté") = sQuantite
     
                        'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
280                     rstPiece.Fields("Prix_Total") = Conversion(CStr(Round(Replace(rstPiece.Fields("Qté"), "*", vbNullString) * rstPiece.Fields("Prix_net") * CSng(sProfit), 2)), MODE_PAS_FORMAT)
      
                        'Pour le profit, c'est le prix total - (prix net * quantité)
285                     rstPiece.Fields("Profit_argent") = Conversion(CStr(Round(rstPiece.Fields("Prix_total") - (rstPiece.Fields("Prix_net") * Replace(rstPiece.Fields("Qté"), "*", vbNullString)), 2)), MODE_PAS_FORMAT)

290                     rstPiece.Fields("Recu") = True
295                     rstPiece.Fields("Commandé") = False
300                     rstPiece.Fields("DateRéception") = txtDateReception.Text

305                     Call rstPiece.Update

310                     Set rstCopiePiece = New ADODB.Recordset

315                     Call rstCopiePiece.Open("SELECT * FROM GRB_Projet_Pieces", g_connData, adOpenDynamic, adLockOptimistic)

320                     Call rstCopiePiece.AddNew

325                     rstCopiePiece.Fields("IDProjet") = rstPiece.Fields("IDProjet")
330                     rstCopiePiece.Fields("IDSection") = rstPiece.Fields("IDSection")
335                     rstCopiePiece.Fields("NumItem") = rstPiece.Fields("NumItem")
340                     rstCopiePiece.Fields("Qté") = CDbl(sTotal) - CDbl(sQuantite)
345                     rstCopiePiece.Fields("Desc_FR") = rstPiece.Fields("Desc_FR")
350                     rstCopiePiece.Fields("Desc_EN") = rstPiece.Fields("Desc_EN")
355                     rstCopiePiece.Fields("Manufact") = rstPiece.Fields("Manufact")
360                     rstCopiePiece.Fields("Prix_List") = rstPiece.Fields("Prix_List")
365                     rstCopiePiece.Fields("Escompte") = rstPiece.Fields("Escompte")
370                     rstCopiePiece.Fields("Prix_net") = rstPiece.Fields("Prix_net")
375                     rstCopiePiece.Fields("IDFRS") = rstPiece.Fields("IDFRS")

                        'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
380                     rstCopiePiece.Fields("Prix_Total") = Conversion(CStr(Round(Replace(rstCopiePiece.Fields("Qté"), "*", vbNullString) * rstCopiePiece.Fields("Prix_net") * CSng(sProfit), 2)), MODE_PAS_FORMAT)
     
385                     rstCopiePiece.Fields("Profit_Pourcent") = rstPiece.Fields("Profit_Pourcent")
     
                        'Pour le profit, c'est le prix total - (prix net * quantité)
390                     rstCopiePiece.Fields("Profit_argent") = Conversion(CStr(Round(rstCopiePiece.Fields("Prix_total") - (rstCopiePiece.Fields("Prix_net") * Replace(rstCopiePiece.Fields("Qté"), "*", vbNullString)), 2)), MODE_PAS_FORMAT)

395                     rstCopiePiece.Fields("SousSection") = rstPiece.Fields("SousSection")
400                     rstCopiePiece.Fields("OrdreSection") = rstPiece.Fields("OrdreSection")
405                     rstCopiePiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne") + 1
410                     rstCopiePiece.Fields("PrixOrigine") = rstPiece.Fields("PrixOrigine")
415                     rstCopiePiece.Fields("Type") = rstPiece.Fields("Type")
420                     rstCopiePiece.Fields("Visible") = rstPiece.Fields("Visible")
425                     rstCopiePiece.Fields("Commandé") = True
430                     rstCopiePiece.Fields("Quoté") = rstPiece.Fields("Quoté")
435                     rstCopiePiece.Fields("Recu") = False
440                     rstCopiePiece.Fields("Retour") = False
445                     rstCopiePiece.Fields("NoRetour") = vbNullString
450                     rstCopiePiece.Fields("CommandeAnnulée") = False
455                     rstCopiePiece.Fields("DateRéception") = vbNullString
465                     rstCopiePiece.Fields("Facturation") = rstPiece.Fields("Facturation")
470                     rstCopiePiece.Fields("ID") = ""
475                     rstCopiePiece.Fields("PieceExtra") = rstPiece.Fields("PieceExtra")
480                     rstCopiePiece.Fields("DateCommande") = rstPiece.Fields("DateCommande")
485                     rstCopiePiece.Fields("DateRequise") = rstPiece.Fields("DateRequise")
490                     rstCopiePiece.Fields("MatérielInutile") = False

495                     Call rstCopiePiece.Update

500                     Call rstCopiePiece.Close
505                     Set rstCopiePiece = Nothing

510                     Call rstPiece.Close

515                     Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & txtnoprojet.Text & "' AND Type = 'M' AND NuméroLigne >= " & lvwProjet.SelectedItem.ListSubItems(I_COL_MANUFACTURIER).Tag + 1, g_connData, adOpenDynamic, adLockOptimistic)

520                     bSkip = False

525                     Do While Not rstPiece.EOF
530                       If ((rstPiece.Fields("NumItem") <> lvwProjet.SelectedItem.SubItems(I_COL_PIECE)) Or (rstPiece.Fields("Qté") <> CDbl(sTotal) - CDbl(sQuantite))) Or bSkip = True Then
535                         rstPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne") + 1

540                         Call rstPiece.Update
545                       Else
550                         bSkip = True
555                       End If

560                       Call rstPiece.MoveNext
565                     Loop

570                     Call rstPiece.Close

                        'Ajout aux modifs
575                     Call rstModif.Open("SELECT * FROM GRB_Projet_Modif", g_connData, adOpenDynamic, adLockOptimistic)

580                     Call rstModif.AddNew

585                     Call rstEmploye.Open("SELECT noEmploye FROM GRB_Employés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
590                     rstModif.Fields("Type") = "M"
595                     rstModif.Fields("IDProjet") = txtnoprojet.Text
600                     rstModif.Fields("noEmployé") = rstEmploye.Fields("noEmploye")
605                     rstModif.Fields("Date") = ConvertDate(Date)
610                     rstModif.Fields("Heure") = Time
615                     rstModif.Fields("TypeModif") = "RECEPTION"

620                     Call rstEmploye.Close

625                     Call rstModif.Update

630                     Call rstModif.Close

635                     If UCase(lvwProjet.SelectedItem.SubItems(I_COL_DISTRIBUTEUR)) = "SOLUTION GRB INC." Then
640                       Call EnleverInventaireProjet(sQuantite)
645                     End If

650                     m_iIndexReception = lvwProjet.SelectedItem.Index

655                     Call RemplirListViewProjet(txtnoprojet.Text)
660                   Else
665                     Call MsgBox("La quantité est trop grande!", vbOKOnly, "Erreur")
670                   End If
675                 End If

680                 Set rstPiece = Nothing
685                 Set rstModif = Nothing
690                 Set rstEmploye = Nothing
695               Else
700                 Call MsgBox("La quantité doit être plus grande que 0!", vbOKOnly, "Erreur")
705               End If
710             Else
715               Call MsgBox("Quantité non numérique!", vbOKOnly, "Erreur")
720             End If
725           End If
730         End If
735       Else
740         Call MsgBox("Ce projet est en modification par " & rstProjet.Fields("Par") & "!", vbOKOnly, "Erreur")
745       End If

750       Call rstProjet.Close
755       Set rstProjet = Nothing
760     End If

765     Exit Sub

AfficherErreur:

770     woups "frmReceptionMec", "Reception", Err, Erl
End Sub

Private Sub ReceptionAchat()

5       On Error GoTo AfficherErreur

10      Dim rstPiece      As ADODB.Recordset
15      Dim rstCopiePiece As ADODB.Recordset
20      Dim rstAchat      As ADODB.Recordset
25      Dim sQuantite     As String
30      Dim sIDAchat      As String
35      Dim sTotal        As String
40      Dim bSkip         As Boolean
45      Dim iIndexAchat   As Integer
50      Dim iIDFRS        As Integer

        'Si il y a des enregistrements
55      If lvwProjet.ListItems.count > 0 Then
60        sIDAchat = Left$(txtnoprojet.Text, 9)
65        iIndexAchat = CInt(Right$(txtnoprojet.Text, 3))

70        Set rstAchat = New ADODB.Recordset

75        Call rstAchat.Open("SELECT Modification, Par FROM GRB_Achat WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenDynamic, adLockOptimistic)

80        If rstAchat.Fields("Modification") = False Then
85          If lvwProjet.SelectedItem.ForeColor = COLOR_ORANGE Or lvwProjet.SelectedItem.ForeColor = COLOR_BLEU Then 'COLOR_ORANGE ou bleu
90            sQuantite = InputBox("Quelle est la quantité reçue?")

95            sQuantite = Replace(sQuantite, ".", ",")

100           If sQuantite <> "" Then
105             If IsNumeric(sQuantite) Then
110               If CDbl(sQuantite) > 0 Then
115                 Set rstPiece = New ADODB.Recordset

120                 If CDbl(sQuantite) = CDbl(lvwProjet.SelectedItem.Text) Then
125                   Call rstPiece.Open("SELECT * FROM GRB_Achat_Pieces WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat & " AND NuméroLigne = " & lvwProjet.SelectedItem.ListSubItems(I_COL_MANUFACTURIER).Tag, g_connData, adOpenDynamic, adLockOptimistic)

135                   rstPiece.Fields("Recu") = True
140                   rstPiece.Fields("Commandé") = False

145                   rstPiece.Fields("DateRéception") = txtDateReception.Text

150                   Call rstPiece.Update

155                   If (CInt(Right$(sIDAchat, 2)) >= 0 And CInt(Right$(sIDAchat, 2)) <= 12) Or rstPiece.Fields("IDFRS") = 717 Then
160                     Call AjouterInventaireAchat(sQuantite)
165                   End If

170                   Call rstPiece.Close

175                   m_iIndexReception = lvwProjet.SelectedItem.Index

180                   Call RemplirListViewAchat(txtnoprojet.Text)
185                 Else
190                   If CDbl(sQuantite) < CDbl(lvwProjet.SelectedItem.Text) Then
195                     Call rstPiece.Open("SELECT * FROM GRB_Achat_Pieces WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat & " AND NuméroLigne = " & lvwProjet.SelectedItem.ListSubItems(I_COL_MANUFACTURIER).Tag, g_connData, adOpenDynamic, adLockOptimistic)

200                     sTotal = rstPiece.Fields("Qté")

205                     rstPiece.Fields("Qté") = sQuantite
   
                        'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
215                     rstPiece.Fields("Prix_Total") = Conversion(CStr(Round(Replace(rstPiece.Fields("Qté"), "*", vbNullString) * rstPiece.Fields("Prix_net"), 2)), MODE_PAS_FORMAT)

220                     rstPiece.Fields("Recu") = True
225                     rstPiece.Fields("Commandé") = False
230                     rstPiece.Fields("DateRéception") = txtDateReception.Text

235                     Call rstPiece.Update

240                     Set rstCopiePiece = New ADODB.Recordset

245                     Call rstCopiePiece.Open("SELECT * FROM GRB_Achat_Pieces", g_connData, adOpenDynamic, adLockOptimistic)

250                     Call rstCopiePiece.AddNew

255                     rstCopiePiece.Fields("IDAchat") = rstPiece.Fields("IDAchat")
260                     rstCopiePiece.Fields("IndexAchat") = rstPiece.Fields("IndexAchat")
265                     rstCopiePiece.Fields("PIECE") = rstPiece.Fields("PIECE")
270                     rstCopiePiece.Fields("Qté") = CDbl(sTotal) - CDbl(sQuantite)
275                     rstCopiePiece.Fields("Desc_FR") = rstPiece.Fields("Desc_FR")
280                     rstCopiePiece.Fields("Desc_EN") = rstPiece.Fields("Desc_EN")
285                     rstCopiePiece.Fields("Manufact") = rstPiece.Fields("Manufact")
290                     rstCopiePiece.Fields("Prix_List") = rstPiece.Fields("Prix_List")
295                     rstCopiePiece.Fields("Escompte") = rstPiece.Fields("Escompte")
300                     rstCopiePiece.Fields("Prix_net") = rstPiece.Fields("Prix_net")
305                     rstCopiePiece.Fields("IDFRS") = rstPiece.Fields("IDFRS")

                        'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
310                     rstCopiePiece.Fields("Prix_Total") = Conversion(CStr(Round(Replace(rstCopiePiece.Fields("Qté"), "*", vbNullString) * rstCopiePiece.Fields("Prix_net"), 2)), MODE_PAS_FORMAT)
   
315                     rstCopiePiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne") + 1
320                     rstCopiePiece.Fields("Type") = rstPiece.Fields("Type")
325                     rstCopiePiece.Fields("Commandé") = True
330                     rstCopiePiece.Fields("Recu") = False
335                     rstCopiePiece.Fields("Retour") = False
340                     rstCopiePiece.Fields("NoRetour") = vbNullString
345                     rstCopiePiece.Fields("DateRéception") = vbNullString
355                     rstCopiePiece.Fields("DateCommande") = rstPiece.Fields("DateCommande")
360                     rstCopiePiece.Fields("DateRequise") = rstPiece.Fields("DateRequise")

365                     Call rstCopiePiece.Update

370                     Call rstCopiePiece.Close
375                     Set rstCopiePiece = Nothing

380                     iIDFRS = 717

385                     Call rstPiece.Close

390                     Call rstPiece.Open("SELECT * FROM GRB_Achat_Pieces WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat & " AND NuméroLigne >= " & lvwProjet.SelectedItem.ListSubItems(I_COL_MANUFACTURIER).Tag + 1, g_connData, adOpenDynamic, adLockOptimistic)

395                     bSkip = False

400                     Do While Not rstPiece.EOF
405                       If ((rstPiece.Fields("PIECE") <> lvwProjet.SelectedItem.SubItems(I_COL_PIECE)) Or (rstPiece.Fields("Qté") <> CDbl(sTotal) - CDbl(sQuantite))) Or bSkip = True Then
410                         rstPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne") + 1

415                         Call rstPiece.Update
420                       Else
425                         bSkip = True
430                       End If

435                       Call rstPiece.MoveNext
440                     Loop

445                     Call rstPiece.Close

450                     If CInt(Right$(sIDAchat, 2)) >= 0 And CInt(Right$(sIDAchat, 2)) <= 12 Or iIDFRS = 717 Then
455                       Call AjouterInventaireAchat(sQuantite)
460                     End If

465                     m_iIndexReception = lvwProjet.SelectedItem.Index

470                     Call RemplirListViewAchat(txtnoprojet.Text)
475                   Else
480                     Call MsgBox("La quantité est trop grande!", vbOKOnly, "Erreur")
485                   End If
490                 End If

495                 Set rstPiece = Nothing
500               Else
505                 Call MsgBox("La quantité doit être plus grande que 0!", vbOKOnly, "Erreur")
510               End If
515             Else
520               Call MsgBox("Quantité non numérique!", vbOKOnly, "Erreur")
525             End If
530           End If
535         End If
540       Else
545         Call MsgBox("Ce projet est en modification par " & rstAchat.Fields("Par") & "!", vbOKOnly, "Erreur")
550       End If

555       Call rstAchat.Close
560       Set rstAchat = Nothing
565     End If

570     Exit Sub

AfficherErreur:

575     woups "frmReceptionMec", "ReceptionAchat", Err, Erl
End Sub

Private Sub EnleverInventaireProjet(ByVal sQuantite As String)

5       On Error GoTo AfficherErreur

10      Dim rstInventaire As ADODB.Recordset
15      Dim rstPieceFRS   As ADODB.Recordset
20      Dim rstProjet     As ADODB.Recordset

25      If MsgBox("Voulez-vous modifier l'inventaire?", vbYesNo) = vbYes Then
30        Set rstInventaire = New ADODB.Recordset

35        Call rstInventaire.Open("SELECT * FROM GRB_InventaireMec WHERE NoItem = '" & Replace(lvwProjet.SelectedItem.SubItems(I_COL_PIECE), "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)

40        If rstInventaire.EOF Then
45          Call rstInventaire.AddNew
  
50          rstInventaire.Fields("NoItem") = lvwProjet.SelectedItem.SubItems(I_COL_PIECE)
55          rstInventaire.Fields("Description") = lvwProjet.SelectedItem.SubItems(I_COL_DESCRIPTION)
60          rstInventaire.Fields("Manufacturier") = lvwProjet.SelectedItem.SubItems(I_COL_MANUFACTURIER)

65          Call frmChoixQteBoite.Afficher(lvwProjet.SelectedItem.SubItems(I_COL_PIECE))

70          rstInventaire.Fields("CommandeParBoite") = g_bQteBoite
75          rstInventaire.Fields("QteBoite") = g_sQteBoite

80          rstInventaire.Fields("Commentaires") = ""
85          rstInventaire.Fields("QuantitéStock") = "0"
  
90          Call frmChoixLocalisation.Afficher(MECANIQUE, lvwProjet.SelectedItem.SubItems(I_COL_PIECE))
  
95          rstInventaire.Fields("Localisation") = g_sLocalisation
100         rstInventaire.Fields("Minimum") = False
105         rstInventaire.Fields("QuantitéMinimum") = ""
110         rstInventaire.Fields("Commande") = ""
115       End If

120       Set rstProjet = New ADODB.Recordset

125       Call rstProjet.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & txtnoprojet.Text & "' AND NuméroLigne = " & lvwProjet.SelectedItem.ListSubItems(I_COL_MANUFACTURIER).Tag, g_connData, adOpenDynamic, adLockOptimistic)
    
130       If rstInventaire.Fields("CommandeParBoite") = True Then
135         rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) - (CDbl(sQuantite) * CDbl(rstInventaire.Fields("QteBoite"))), ".", ",")

140         If rstProjet.Fields("Prix_List") <> "" Then
145           If rstInventaire.Fields("QteBoite") <> "" Then
150             rstInventaire.Fields("Prix Liste") = Replace(rstProjet.Fields("Prix_List") / rstInventaire.Fields("QteBoite"), ".", ",")
155           Else
160             rstInventaire.Fields("Prix Liste") = rstProjet.Fields("Prix_List")
165           End If
170         Else
175           rstInventaire.Fields("Prix Liste") = "0"
180         End If

185         rstInventaire.Fields("Escompte") = rstProjet.Fields("Escompte")

190         If rstInventaire.Fields("QteBoite") <> "" Then
195           rstInventaire.Fields("Prix net") = Replace(rstProjet.Fields("Prix_Net") / rstInventaire.Fields("QteBoite"), ".", ",")
200         Else
205           rstInventaire.Fields("Prix net") = rstProjet.Fields("Prix_Net")
210         End If
215       Else
220         rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) - CDbl(sQuantite), ".", ",")

225         If rstProjet.Fields("Prix_List") <> "" Then
230           rstInventaire.Fields("Prix Liste") = rstProjet.Fields("Prix_List")
235         Else
240           rstInventaire.Fields("Prix Liste") = ""
245         End If

250         rstInventaire.Fields("Escompte") = rstProjet.Fields("Escompte")
255         rstInventaire.Fields("Prix net") = rstProjet.Fields("Prix_Net")
260       End If
  
265       Call rstInventaire.Update
  
270       Set rstPieceFRS = New ADODB.Recordset
  
275       Call rstPieceFRS.Open("SELECT * FROM GRB_PiecesFRS WHERE PIECE = '" & Replace(lvwProjet.SelectedItem.SubItems(I_COL_PIECE), "'", "''") & "' AND IDFRS = 717", g_connData, adOpenDynamic, adLockOptimistic)
  
280       If rstPieceFRS.EOF Then
285         Call rstPieceFRS.AddNew
  
290         rstPieceFRS.Fields("IDFRS") = 717
295         rstPieceFRS.Fields("PIECE") = lvwProjet.SelectedItem.SubItems(I_COL_PIECE)
300         rstPieceFRS.Fields("PERS_RESS") = 901
305         rstPieceFRS.Fields("ENTRER_PAR") = g_sInitiale
310         rstPieceFRS.Fields("DeviseMonétaire") = "CAN"
315         rstPieceFRS.Fields("Type") = "M"
320       End If
  
325       rstPieceFRS.Fields("PRIX_LIST") = rstProjet.Fields("Prix_List")
330       rstPieceFRS.Fields("ESCOMPTE") = rstProjet.Fields("Escompte")
335       rstPieceFRS.Fields("PRIX_NET") = rstProjet.Fields("Prix_net")
340       rstPieceFRS.Fields("DATE") = txtDateReception.Text

345       Call rstPieceFRS.Update
  
350       Call rstPieceFRS.Close
355       Set rstPieceFRS = Nothing
  
360       Call rstProjet.Close
365       Set rstProjet = Nothing
  
370       If rstInventaire.Fields("CommandeParBoite") = True Then
375         sQuantite = Replace(CDbl(sQuantite) * CDbl(rstInventaire.Fields("QteBoite")), ".", ",")
380       End If
  
385       Call rstInventaire.Close
390       Set rstInventaire = Nothing
  
395       Call AjouterHistorique(sQuantite)
400     End If

405     Exit Sub

AfficherErreur:

410     woups "frmReceptionMec", "EnleverInventaireProjet", Err, Erl
End Sub

Private Sub AjouterInventaireAchat(ByVal sQuantite As String)

5       On Error GoTo AfficherErreur

10      Dim rstInventaire As ADODB.Recordset
15      Dim rstPieceFRS   As ADODB.Recordset
20      Dim rstAchat      As ADODB.Recordset
25      Dim sIDAchat      As String
30      Dim iIndexAchat   As Integer
35      Dim iCompteur     As Integer

40      If MsgBox("Voulez-vous modifier l'inventaire?", vbYesNo) = vbYes Then
45        Set rstInventaire = New ADODB.Recordset

50        Call rstInventaire.Open("SELECT * FROM GRB_InventaireMec WHERE NoItem = '" & Replace(lvwProjet.SelectedItem.SubItems(I_COL_PIECE), "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)

55        If rstInventaire.EOF Then
60          Call rstInventaire.AddNew

65          rstInventaire.Fields("NoItem") = lvwProjet.SelectedItem.SubItems(I_COL_PIECE)
70          rstInventaire.Fields("Description") = lvwProjet.SelectedItem.SubItems(I_COL_DESCRIPTION)
75          rstInventaire.Fields("Manufacturier") = lvwProjet.SelectedItem.SubItems(I_COL_MANUFACTURIER)

80          Call frmChoixQteBoite.Afficher(lvwProjet.SelectedItem.SubItems(I_COL_PIECE))

85          rstInventaire.Fields("CommandeParBoite") = g_bQteBoite
90          rstInventaire.Fields("QteBoite") = g_sQteBoite

95          rstInventaire.Fields("Commentaires") = ""
100         rstInventaire.Fields("QuantitéStock") = "0"

105         Call frmChoixLocalisation.Afficher(MECANIQUE, lvwProjet.SelectedItem.SubItems(I_COL_PIECE))

110         rstInventaire.Fields("Localisation") = g_sLocalisation
115         rstInventaire.Fields("Minimum") = False
120         rstInventaire.Fields("QuantitéMinimum") = ""
125         rstInventaire.Fields("Commande") = ""
130       End If
  
135       sIDAchat = Left$(txtnoprojet.Text, 9)
140       iIndexAchat = CInt(Right$(txtnoprojet.Text, 3))
  
145       Set rstAchat = New ADODB.Recordset
  
150       Call rstAchat.Open("SELECT * FROM GRB_Achat_Pieces WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat & " AND NuméroLigne = " & lvwProjet.SelectedItem.ListSubItems(I_COL_MANUFACTURIER).Tag, g_connData, adOpenDynamic, adLockOptimistic)
    
155       If rstInventaire.Fields("CommandeParBoite") = True Then
160         If rstAchat.Fields("IDFRS") = 717 Then
165           rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) - (CDbl(sQuantite) * CDbl(rstInventaire.Fields("QteBoite"))), ".", ",")
170         Else
175           rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) + (CDbl(sQuantite) * CDbl(rstInventaire.Fields("QteBoite"))), ".", ",")
180         End If

185         If rstAchat.Fields("Prix_List") <> "" Then
190           If rstInventaire.Fields("QteBoite") <> "" Then
195             rstInventaire.Fields("Prix Liste") = Replace(rstAchat.Fields("Prix_List") / rstInventaire.Fields("QteBoite"), ".", ",")
200           Else
205             rstInventaire.Fields("Prix Liste") = rstAchat.Fields("Prix_List")
210           End If
215         Else
220           rstInventaire.Fields("Prix Liste") = "0"
225         End If

230         rstInventaire.Fields("Escompte") = rstAchat.Fields("Escompte")

235         If rstInventaire.Fields("QteBoite") <> "" Then
240           rstInventaire.Fields("Prix net") = Replace(rstAchat.Fields("Prix_Net") / rstInventaire.Fields("QteBoite"), ".", ",")
245         Else
250           rstInventaire.Fields("Prix net") = rstAchat.Fields("Prix_Net")
255         End If
260       Else
265         If rstAchat.Fields("IDFRS") = 717 Then
270           rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) - CDbl(sQuantite), ".", ",")
275         Else
280           rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) + CDbl(sQuantite), ".", ",")
285         End If

290         If rstAchat.Fields("Prix_List") <> "" Then
295           rstInventaire.Fields("Prix Liste") = rstAchat.Fields("Prix_List")
300         Else
305           rstInventaire.Fields("Prix Liste") = "0"
310         End If

315         rstInventaire.Fields("Escompte") = rstAchat.Fields("Escompte")
320         rstInventaire.Fields("Prix net") = rstAchat.Fields("Prix_Net")
325       End If
    
330       Call rstInventaire.Update

335       Set rstPieceFRS = New ADODB.Recordset
  
340       Call rstPieceFRS.Open("SELECT * FROM GRB_PiecesFRS WHERE PIECE = '" & Replace(lvwProjet.SelectedItem.SubItems(I_COL_PIECE), "'", "''") & "' AND IDFRS = 717", g_connData, adOpenDynamic, adLockOptimistic)
  
345       If rstPieceFRS.EOF Then
350         Call rstPieceFRS.AddNew
  
355         rstPieceFRS.Fields("IDFRS") = 717
360         rstPieceFRS.Fields("PIECE") = lvwProjet.SelectedItem.SubItems(I_COL_PIECE)
365         rstPieceFRS.Fields("PERS_RESS") = 901
370         rstPieceFRS.Fields("ENTRER_PAR") = g_sInitiale
375         rstPieceFRS.Fields("DeviseMonétaire") = "CAN"
380         rstPieceFRS.Fields("Type") = "M"
385       End If

390       rstPieceFRS.Fields("PRIX_LIST") = rstAchat.Fields("Prix_List")
395       rstPieceFRS.Fields("ESCOMPTE") = rstAchat.Fields("Escompte")
400       rstPieceFRS.Fields("PRIX_NET") = rstAchat.Fields("Prix_net")
405       rstPieceFRS.Fields("DATE") = txtDateReception.Text

410       Call rstPieceFRS.Update
  
415       Call rstPieceFRS.Close
420       Set rstPieceFRS = Nothing
  
425       Call rstAchat.Close
430       Set rstAchat = Nothing
  
435       If rstInventaire.Fields("CommandeParBoite") = True Then
440         sQuantite = CDbl(sQuantite) * CDbl(rstInventaire.Fields("QteBoite"))
445       End If
  
450       Call rstInventaire.Close
455       Set rstInventaire = Nothing
  
460       Call AjouterHistorique(sQuantite)
465     End If

470     Exit Sub

AfficherErreur:

475     woups "frmReceptionMec", "AjouterInventaireAchat", Err, Erl
End Sub

Private Sub AjouterHistorique(ByVal sQuantite As String)

5       On Error GoTo AfficherErreur

10      Dim rstHist As ADODB.Recordset

15      Set rstHist = New ADODB.Recordset

20      Call rstHist.Open("SELECT * FROM GRB_InventaireMecModif", g_connData, adOpenDynamic, adLockOptimistic)

25      Call rstHist.AddNew

30      rstHist.Fields("Date") = txtDateReception.Text
35      rstHist.Fields("IDProjet") = txtnoprojet.Text
40      rstHist.Fields("NoItem") = lvwProjet.SelectedItem.SubItems(I_COL_PIECE)

45      If m_eType = ACHAT Then
50        If lvwProjet.SelectedItem.ListSubItems(I_COL_DISTRIBUTEUR).Tag = 717 Then
55          rstHist.Fields("Quantité") = "-" & sQuantite
60        Else
65          rstHist.Fields("Quantité") = sQuantite
70        End If
75      Else
80        rstHist.Fields("Quantité") = "-" & sQuantite
85      End If

90      rstHist.Fields("User") = g_sInitiale

95      Call rstHist.Update

100     Call rstHist.Close
105     Set rstHist = Nothing

110     Exit Sub

AfficherErreur:

115     woups "frmReceptionMec", "AjouterHistorique", Err, Erl
End Sub

Private Sub SupprimerHistorique(Optional ByVal sQuantite As String = "")

5       On Error GoTo AfficherErreur

10      Dim rstHist As ADODB.Recordset

15      Set rstHist = New ADODB.Recordset

20      If m_eType = ACHAT Then
25        If sQuantite <> "" Then
30          If lvwProjet.SelectedItem.ListSubItems(I_COL_DISTRIBUTEUR).Tag = 717 Then
35            Call rstHist.Open("SELECT * FROM GRB_InventaireMecModif WHERE [Date] = '" & lvwProjet.SelectedItem.SubItems(I_COL_DATE_RECEPTION) & "' AND IDProjet = '" & txtnoprojet.Text & "' AND NoItem = '" & Replace(lvwProjet.SelectedItem.SubItems(I_COL_PIECE), "'", "''") & "' AND Quantité = '-" & sQuantite & "'", g_connData, adOpenDynamic, adLockOptimistic)
40          Else
45            Call rstHist.Open("SELECT * FROM GRB_InventaireMecModif WHERE [Date] = '" & lvwProjet.SelectedItem.SubItems(I_COL_DATE_RECEPTION) & "' AND IDProjet = '" & txtnoprojet.Text & "' AND NoItem = '" & Replace(lvwProjet.SelectedItem.SubItems(I_COL_PIECE), "'", "''") & "' AND Quantité = '" & sQuantite & "'", g_connData, adOpenDynamic, adLockOptimistic)
50          End If
55        Else
60          If lvwProjet.SelectedItem.ListSubItems(I_COL_DISTRIBUTEUR).Tag = 717 Then
65            Call rstHist.Open("SELECT * FROM GRB_InventaireMecModif WHERE [Date] = '" & lvwProjet.SelectedItem.SubItems(I_COL_DATE_RECEPTION) & "' AND IDProjet = '" & txtnoprojet.Text & "' AND NoItem = '" & Replace(lvwProjet.SelectedItem.SubItems(I_COL_PIECE), "'", "''") & "' AND Quantité = '-" & lvwProjet.SelectedItem.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
70          Else
75            Call rstHist.Open("SELECT * FROM GRB_InventaireMecModif WHERE [Date] = '" & lvwProjet.SelectedItem.SubItems(I_COL_DATE_RECEPTION) & "' AND IDProjet = '" & txtnoprojet.Text & "' AND NoItem = '" & Replace(lvwProjet.SelectedItem.SubItems(I_COL_PIECE), "'", "''") & "' AND Quantité = '" & lvwProjet.SelectedItem.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
80          End If
85        End If
90      Else
95        If sQuantite <> "" Then
100         Call rstHist.Open("SELECT * FROM GRB_InventaireMecModif WHERE [Date] = '" & lvwProjet.SelectedItem.SubItems(I_COL_DATE_RECEPTION) & "' AND IDProjet = '" & txtnoprojet.Text & "' AND NoItem = '" & Replace(lvwProjet.SelectedItem.SubItems(I_COL_PIECE), "'", "''") & "' AND Quantité = '" & "-" & sQuantite & "'", g_connData, adOpenDynamic, adLockOptimistic)
105       Else
110         Call rstHist.Open("SELECT * FROM GRB_InventaireMecModif WHERE [Date] = '" & lvwProjet.SelectedItem.SubItems(I_COL_DATE_RECEPTION) & "' AND IDProjet = '" & txtnoprojet.Text & "' AND NoItem = '" & Replace(lvwProjet.SelectedItem.SubItems(I_COL_PIECE), "'", "''") & "' AND Quantité = '" & "-" & lvwProjet.SelectedItem.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
115       End If
120     End If

125     If Not rstHist.EOF Then
130       Call rstHist.Delete
135     End If

140     Call rstHist.Close
145     Set rstHist = Nothing

150     Exit Sub

AfficherErreur:

155     woups "frmReceptionMec", "SupprimerHistorique", Err, Erl
End Sub

Private Sub Cmdfermer_Click()

5       On Error GoTo AfficherErreur
 
10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmReceptionMec", "cmdFermer_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer

15      Call Unload(frmChoixProjSoum)

20      txtDateReception.Text = ConvertDate(Date)
25      txtDateRequise.Text = ConvertDate(Date)

30      If m_sNoProjet <> "" Then
35        cmbType.ListIndex = 0

40        For iCompteur = 0 To cmbNoProjet.ListCount - 1
45          If cmbNoProjet.LIST(iCompteur) = m_sNoProjet Then
50            cmbNoProjet.ListIndex = iCompteur

55            Exit For
60          End If
65        Next
70      Else
75        If m_sNoAchat <> "" Then
80          cmbType.ListIndex = 1

85          For iCompteur = 0 To cmbNoProjet.ListCount - 1
90            If cmbNoProjet.LIST(iCompteur) = m_sNoAchat Then
95              cmbNoProjet.ListIndex = iCompteur

100             Exit For
105           End If
110         Next
115       Else
120         cmbType.ListIndex = 0
125       End If
130     End If

135     Exit Sub

AfficherErreur:

140     woups "frmReceptionMec", "Form_Load", Err, Erl
End Sub

Private Sub RemplirComboProjet()

5       On Error GoTo AfficherErreur

        'Rempli le combo des soumissions
10      Dim rstProjet As ADODB.Recordset
  
        'Il faut vider le combo avant de le remplir
15      Call cmbNoProjet.Clear
  
        'Ouvre le recordset selon le type
20      Set rstProjet = New ADODB.Recordset
        
25      Call rstProjet.Open("SELECT IDProjet FROM GRB_ProjetMec ORDER BY IDProjet DESC", g_connData, adOpenDynamic, adLockOptimistic)
    
        'Tant que ce n'est pas la fin des enregistrements
30      Do While Not rstProjet.EOF
35        Call cmbNoProjet.AddItem(rstProjet.Fields("IDProjet"))

40        Call rstProjet.MoveNext
45      Loop
      
50      Call rstProjet.Close
55      Set rstProjet = Nothing
  
        'Si le combo n'est pas vide, on sélectionne l'item voulu ou le premier
60      If cmbNoProjet.ListCount > 0 Then
          'On sélectionne le premier
65        cmbNoProjet.ListIndex = 0
70      End If

75      Exit Sub

AfficherErreur:

80      woups "frmReceptionMec", "RemplirComboProjet", Err, Erl
End Sub

Private Sub RemplirComboAchat()

5       On Error GoTo AfficherErreur

        'Rempli le combo des soumissions
10      Dim rstAchat As ADODB.Recordset
  
        'Il faut vider le combo avant de le remplir
15      Call cmbNoProjet.Clear
  
        'Ouvre le recordset selon le type
20      Set rstAchat = New ADODB.Recordset
        
25      Call rstAchat.Open("SELECT IDAchat, IndexAchat FROM GRB_Achat WHERE Type = 'M' ORDER BY IDAchat DESC, IndexAchat DESC", g_connData, adOpenDynamic, adLockOptimistic)
    
        'Tant que ce n'est pas la fin des enregistrements
30      Do While Not rstAchat.EOF
35        Call cmbNoProjet.AddItem(rstAchat.Fields("IDAchat") & "-" & Right$("000" & rstAchat.Fields("IndexAchat"), 3))

40        Call rstAchat.MoveNext
45      Loop
      
50      Call rstAchat.Close
55      Set rstAchat = Nothing
  
        'Si le combo n'est pas vide, on sélectionne l'item voulu ou le premier
60      If cmbNoProjet.ListCount > 0 Then
          'On sélectionne le premier
65        cmbNoProjet.ListIndex = 0
70      End If

75      Exit Sub

AfficherErreur:

80      woups "frmReceptionMec", "RemplirComboAchat", Err, Erl
End Sub

Private Sub RemplirListViewProjet(ByVal sNoProjet As String)

5       On Error GoTo AfficherErreur

        'Remplis les pièces de la soumission avec la BD
10      Dim rstProjet     As ADODB.Recordset
15      Dim rstSection    As ADODB.Recordset
20      Dim rstFRS        As ADODB.Recordset
25      Dim itmProjet     As ListItem
30      Dim bPremierEnr   As Boolean
35      Dim iOrdreSection As Integer
40      Dim sSousSection  As String
45      Dim lColor        As Long
  
50      Call lvwProjet.ListItems.Clear
  
55      bPremierEnr = True
  
60      Set rstProjet = New ADODB.Recordset
65      Set rstFRS = New ADODB.Recordset
70      Set rstSection = New ADODB.Recordset
  
75      Call rstProjet.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjet & "' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

80      Do While Not rstProjet.EOF
85        Set itmProjet = lvwProjet.ListItems.Add
          
          'Si c'est le premier enregistrement, il faut ajouter la section et la sous-section
90        If bPremierEnr = True Then
95          iOrdreSection = rstProjet.Fields("OrdreSection")
100         sSousSection = rstProjet.Fields("SousSection")
     
            'Pour avoir le nom de la section
105         Call rstSection.Open("SELECT NomSectionFR FROM GRB_SoumProjSectionMec WHERE IDSection = " & rstProjet.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
      
            'Ajout du nom de la section
110         If Not IsNull(rstSection.Fields("NomSectionFR")) Then
115           itmProjet.SubItems(I_COL_PIECE) = rstSection.Fields("NomSectionFR")
120         Else
125           itmProjet.SubItems(I_COL_PIECE) = vbNullString
130         End If
      
135         itmProjet.ListSubItems(I_COL_PIECE).Bold = True
                    
140         Call rstSection.Close
        
145         Set itmProjet = lvwProjet.ListItems.Add
      
            'Ajout du nom de la sous-section
150         If sSousSection = "PAS DE SOUS-SECTION" Then
155           itmProjet.SubItems(I_COL_DESCRIPTION) = vbNullString
160         Else
165           itmProjet.SubItems(I_COL_DESCRIPTION) = sSousSection
170         End If
      
            'Le tag ne peut pas être remplis si la colonne est vide
175         itmProjet.SubItems(I_COL_MANUFACTURIER) = " "
             
180         itmProjet.ListSubItems(I_COL_DESCRIPTION).Bold = True
      
185         itmProjet.Tag = rstProjet.Fields("IDSection")
      
190         Set itmProjet = lvwProjet.ListItems.Add
      
195         bPremierEnr = False
200       Else
            'Si c'est pas le premier enregistrement, il faut vérifier avec l'ancienne section
205         If iOrdreSection <> rstProjet.Fields("OrdreSection") Then
210           iOrdreSection = rstProjet.Fields("OrdreSection")
        
215           Call rstSection.Open("SELECT NomSectionFR FROM GRB_SoumProjSectionMec WHERE IDSection = " & rstProjet.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
        
220           If Not IsNull(rstSection.Fields("NomSectionFR")) Then
225             itmProjet.SubItems(I_COL_PIECE) = rstSection.Fields("NomSectionFR")
230           Else
235             itmProjet.SubItems(I_COL_PIECE) = vbNullString
240           End If
        
245           itmProjet.ListSubItems(I_COL_PIECE).Bold = True
        
250           Call rstSection.Close
              
255           Set itmProjet = lvwProjet.ListItems.Add
        
260           sSousSection = rstProjet.Fields("SousSection")
        
265           If sSousSection = "PAS DE SOUS-SECTION" Then
270             itmProjet.SubItems(I_COL_DESCRIPTION) = vbNullString
275           Else
280             itmProjet.SubItems(I_COL_DESCRIPTION) = rstProjet.Fields("SousSection")
285           End If
        
              'Le tag ne peut pas être remplis si la colonne est vide
290           itmProjet.SubItems(I_COL_MANUFACTURIER) = " "
        
295           itmProjet.ListSubItems(I_COL_DESCRIPTION).Bold = True
        
300           Set itmProjet = lvwProjet.ListItems.Add
305         Else
              'il faut vérifier avec l'ancienne sous-section
310           If sSousSection <> rstProjet.Fields("SousSection") Then
315             sSousSection = rstProjet.Fields("SousSection")
          
320             If sSousSection = "PAS DE SOUS-SECTION" Then
325               itmProjet.SubItems(I_COL_DESCRIPTION) = vbNullString
330             Else
335               itmProjet.SubItems(I_COL_DESCRIPTION) = sSousSection
340             End If
        
345             itmProjet.ListSubItems(I_COL_DESCRIPTION).Bold = True
          
                'Le tag ne peut pas être remplis si la colonne est vide
350             itmProjet.SubItems(I_COL_MANUFACTURIER) = " "
        
355             Set itmProjet = lvwProjet.ListItems.Add
360           End If
365         End If
370       End If
    
375       If rstProjet.Fields("Commandé") = True Then
380         lColor = COLOR_ORANGE     'COLOR_ORANGE
385       Else
390         If rstProjet.Fields("Recu") = True Then
395           lColor = COLOR_GRIS
400         Else
405           If rstProjet.Fields("Retour") = True Then
410             lColor = COLOR_ROUGE
415           Else
420             lColor = COLOR_NOIR
425           End If
430         End If
435       End If

440       itmProjet.Tag = rstProjet.Fields("IDSection")

          'Quantité
445       If Not IsNull(rstProjet.Fields("Qté")) Then
450         itmProjet.Text = rstProjet.Fields("Qté")
455       Else
460         itmProjet.Text = vbNullString
465       End If

470       itmProjet.ForeColor = lColor
    
          'Numéro d'item
475       If Not IsNull(rstProjet.Fields("NumItem")) Then
480         itmProjet.SubItems(I_COL_PIECE) = rstProjet.Fields("NumItem")
485       Else
490         itmProjet.SubItems(I_COL_PIECE) = vbNullString
495       End If
    
500       itmProjet.ListSubItems(I_COL_PIECE).ForeColor = lColor
    
          'On met le nom de la sous-section dans le tag du numéro d'item
505       itmProjet.ListSubItems(I_COL_PIECE).Tag = rstProjet.Fields("SousSection")
    
          'Description en francais
510       If Not IsNull(rstProjet.Fields("Desc_FR")) Then
515         itmProjet.SubItems(I_COL_DESCRIPTION) = rstProjet.Fields("Desc_FR")
520       Else
525         itmProjet.SubItems(I_COL_DESCRIPTION) = vbNullString
530       End If
    
535       itmProjet.ListSubItems(I_COL_DESCRIPTION).ForeColor = lColor
    
          'On met la description en anglais dans le tag de la description en francais
540       If Not IsNull(rstProjet.Fields("Desc_EN")) Then
545         itmProjet.ListSubItems(I_COL_DESCRIPTION).Tag = rstProjet.Fields("Desc_EN")
550       Else
555         itmProjet.ListSubItems(I_COL_DESCRIPTION).Tag = vbNullString
560       End If
    
          'Fabricant
565       If Not IsNull(rstProjet.Fields("Manufact")) Then
570         itmProjet.SubItems(I_COL_MANUFACTURIER) = rstProjet.Fields("Manufact")
575       Else
580         itmProjet.SubItems(I_COL_MANUFACTURIER) = vbNullString
585       End If
    
590       itmProjet.ListSubItems(I_COL_MANUFACTURIER).ForeColor = lColor
    
          'On met l'ordre de la section dans le tag du fabricant
595       itmProjet.ListSubItems(I_COL_MANUFACTURIER).Tag = rstProjet.Fields("NuméroLigne")
    
          'Fournisseur
600       If Not IsNull(rstProjet.Fields("IDFRS")) And rstProjet.Fields("IDFRS") > 0 Then
605         If itmProjet.SubItems(I_COL_PIECE) <> "Texte" Then
610           Call rstFRS.Open("SELECT NomFournisseur FROM GRB_Fournisseur WHERE IDFRS = " & rstProjet.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
    
              'On affiche le nom dans la colonne
615           itmProjet.SubItems(I_COL_DISTRIBUTEUR) = rstFRS.Fields("NomFournisseur")
          
620           itmProjet.ListSubItems(I_COL_DISTRIBUTEUR).ForeColor = lColor
       
              'On affiche l'Id dans le tag
625           itmProjet.ListSubItems(I_COL_DISTRIBUTEUR).Tag = rstProjet.Fields("IDFRS")
      
630           Call rstFRS.Close
635         End If
640       Else
645         itmProjet.SubItems(I_COL_DISTRIBUTEUR) = vbNullString
650       End If

655       If Not IsNull(rstProjet.Fields("DateRéception")) Then
660         itmProjet.SubItems(I_COL_DATE_RECEPTION) = rstProjet.Fields("DateRéception")
665       Else
670         itmProjet.SubItems(I_COL_DATE_RECEPTION) = vbNullString
675       End If

680       itmProjet.ListSubItems(I_COL_DATE_RECEPTION).ForeColor = lColor
           
685       If Not IsNull(rstProjet.Fields("DateCommande")) Then
690         itmProjet.SubItems(I_COL_DATE_COMMANDE) = rstProjet.Fields("DateCommande")
695       Else
700         itmProjet.SubItems(I_COL_DATE_COMMANDE) = vbNullString
705       End If

710       itmProjet.ListSubItems(I_COL_DATE_COMMANDE).ForeColor = lColor
           
715       If Not IsNull(rstProjet.Fields("DateRequise")) Then
720         itmProjet.SubItems(I_COL_DATE_REQUISE) = rstProjet.Fields("DateRequise")
725       Else
730         itmProjet.SubItems(I_COL_DATE_REQUISE) = vbNullString
735       End If

740       itmProjet.ListSubItems(I_COL_DATE_REQUISE).ForeColor = lColor
           
745       Call rstProjet.MoveNext
750     Loop
  
755     Call rstProjet.Close
760     Set rstProjet = Nothing

765     Set rstFRS = Nothing
770     Set rstSection = Nothing

775     If m_iIndexReception > 0 Then
780       lvwProjet.ListItems(m_iIndexReception).Selected = True

785       Call lvwProjet.SetFocus

790       Call lvwProjet.SelectedItem.EnsureVisible
795     End If

800     Exit Sub

AfficherErreur:

805     woups "frmReceptionMec", "RemplirListViewProjet", Err, Erl
End Sub

Private Sub RemplirListViewAchat(ByVal sNoProjet As String)

5       On Error GoTo AfficherErreur

        'Remplis les pièces de la soumission avec la BD
10      Dim rstAchat    As ADODB.Recordset
15      Dim rstFRS      As ADODB.Recordset
20      Dim itmAchat    As ListItem
25      Dim lColor      As Long
30      Dim sNoAchat    As String
35      Dim iIndexAchat As Integer

40      sNoAchat = Left$(sNoProjet, 9)

45      iIndexAchat = CInt(Right$(sNoProjet, 3))
  
50      Call lvwProjet.ListItems.Clear
 
55      Set rstAchat = New ADODB.Recordset
60      Set rstFRS = New ADODB.Recordset
 
65      Call rstAchat.Open("SELECT * FROM GRB_Achat_Pieces WHERE IDAchat = '" & sNoAchat & "' AND IndexAchat = " & Right$("000" & iIndexAchat, 3) & " ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

70      Do While Not rstAchat.EOF
75        Set itmAchat = lvwProjet.ListItems.Add
    
80        If rstAchat.Fields("Commandé") = True Then
85          lColor = COLOR_ORANGE     'COLOR_ORANGE
90        Else
95          If rstAchat.Fields("Recu") = True Then
100           lColor = COLOR_GRIS
105         Else
110           If rstAchat.Fields("Retour") = True Then
115             lColor = COLOR_ROUGE
120           Else
125             lColor = COLOR_NOIR
130           End If
135         End If
140       End If
  
          'Quantité
145       If Not IsNull(rstAchat.Fields("Qté")) Then
150         itmAchat.Text = rstAchat.Fields("Qté")
155       Else
160         itmAchat.Text = vbNullString
165       End If

170       itmAchat.ForeColor = lColor
          
          'Numéro d'item
175       If Not IsNull(rstAchat.Fields("PIECE")) Then
180         itmAchat.SubItems(I_COL_PIECE) = rstAchat.Fields("PIECE")
185       Else
190         itmAchat.SubItems(I_COL_PIECE) = vbNullString
195       End If
    
200       itmAchat.ListSubItems(I_COL_PIECE).ForeColor = lColor
    
          'Description en francais
205       If Not IsNull(rstAchat.Fields("Desc_FR")) Then
210         itmAchat.SubItems(I_COL_DESCRIPTION) = rstAchat.Fields("Desc_FR")
215       Else
220         itmAchat.SubItems(I_COL_DESCRIPTION) = vbNullString
225       End If
   
230       itmAchat.ListSubItems(I_COL_DESCRIPTION).ForeColor = lColor
    
          'On met la description en anglais dans le tag de la description en francais
235       If Not IsNull(rstAchat.Fields("DESC_EN")) Then
240         itmAchat.ListSubItems(I_COL_DESCRIPTION).Tag = rstAchat.Fields("Desc_EN")
245       Else
250         itmAchat.ListSubItems(I_COL_DESCRIPTION).Tag = vbNullString
255       End If
    
          'Fabricant
260       If Not IsNull(rstAchat.Fields("Manufact")) Then
265         itmAchat.SubItems(I_COL_MANUFACTURIER) = rstAchat.Fields("Manufact")
270       Else
275         itmAchat.SubItems(I_COL_MANUFACTURIER) = vbNullString
280       End If
    
285       itmAchat.ListSubItems(I_COL_MANUFACTURIER).ForeColor = lColor

          'On met l'ordre de la section dans le tag du fabricant
290       itmAchat.ListSubItems(I_COL_MANUFACTURIER).Tag = rstAchat.Fields("NuméroLigne")
 
          'Fournisseur
295       If Not IsNull(rstAchat.Fields("IDFRS")) And rstAchat.Fields("IDFRS") > 0 Then
300         If itmAchat.SubItems(I_COL_PIECE) <> "Texte" Then
305           Call rstFRS.Open("SELECT NomFournisseur FROM GRB_Fournisseur WHERE IDFRS = " & rstAchat.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
  
              'On affiche le nom dans la colonne
310           itmAchat.SubItems(I_COL_DISTRIBUTEUR) = rstFRS.Fields("NomFournisseur")
          
315           itmAchat.ListSubItems(I_COL_DISTRIBUTEUR).ForeColor = lColor
        
              'On affiche l'Id dans le tag
320           itmAchat.ListSubItems(I_COL_DISTRIBUTEUR).Tag = rstAchat.Fields("IDFRS")
      
325           Call rstFRS.Close
330         End If
335       Else
340         itmAchat.SubItems(I_COL_DISTRIBUTEUR) = vbNullString
345       End If

350       If Not IsNull(rstAchat.Fields("DateRéception")) Then
355         itmAchat.SubItems(I_COL_DATE_RECEPTION) = rstAchat.Fields("DateRéception")
360       Else
365         itmAchat.SubItems(I_COL_DATE_RECEPTION) = vbNullString
370       End If

375       itmAchat.ListSubItems(I_COL_DATE_RECEPTION).ForeColor = lColor

380       If Not IsNull(rstAchat.Fields("DateCommande")) Then
385         itmAchat.SubItems(I_COL_DATE_COMMANDE) = rstAchat.Fields("DateCommande")
390       Else
395         itmAchat.SubItems(I_COL_DATE_COMMANDE) = vbNullString
400       End If

405       itmAchat.ListSubItems(I_COL_DATE_COMMANDE).ForeColor = lColor

410       If Not IsNull(rstAchat.Fields("DateRequise")) Then
415         itmAchat.SubItems(I_COL_DATE_REQUISE) = rstAchat.Fields("DateRequise")
420       Else
425         itmAchat.SubItems(I_COL_DATE_REQUISE) = vbNullString
430       End If

435       itmAchat.ListSubItems(I_COL_DATE_REQUISE).ForeColor = lColor

440       Call rstAchat.MoveNext
445     Loop
  
450     Call rstAchat.Close
455     Set rstAchat = Nothing

460     Set rstFRS = Nothing

465     If m_iIndexReception > 0 Then
470       lvwProjet.ListItems(m_iIndexReception).Selected = True

475       Call lvwProjet.SetFocus

480       Call lvwProjet.SelectedItem.EnsureVisible
485     End If

490     Exit Sub

AfficherErreur:

495     woups "frmReceptionMec", "RemplirListViewAchat", Err, Erl
End Sub

Private Sub lvwProjet_ItemClick(ByVal Item As MSComctlLib.ListItem)

5       On Error GoTo AfficherErreur

10      Call VerifierBoutonAnnuler

15      Exit Sub

AfficherErreur:

20      woups "frmReceptionMec", "lvwProjet_ItemClick", Err, Erl
End Sub

Private Sub VerifierBoutonAnnuler()

5       On Error GoTo AfficherErreur
        
10      If lvwProjet.ListItems.count > 0 Then
15        If lvwProjet.SelectedItem.ForeColor = COLOR_GRIS Or lvwProjet.SelectedItem.ForeColor = COLOR_BLEU Then 'Gris ou bleu
20          cmdAnnuler.Enabled = True
25        Else
30          cmdAnnuler.Enabled = False
35        End If
40      Else
45        cmdAnnuler.Enabled = False
50      End If

55      Exit Sub

AfficherErreur:

60      woups "frmReceptionMec", "VerifierBoutonAnnuler", Err, Erl
End Sub

Public Sub Afficher(ByVal sUserID As String)
        
5       On Error GoTo AfficherErreur

10      m_sUserID = sUserID

15      Call Me.Show

20      Exit Sub

AfficherErreur:

25      woups "frmReceptionMec", "Afficher", Err, Erl
End Sub

Public Sub AfficherProjet(ByVal sUserID As String, ByVal sNoProjet As String)

5       On Error GoTo AfficherErreur

10      m_sUserID = sUserID

15      m_sNoProjet = sNoProjet

20      Call Me.Show(vbModal)

25      Exit Sub

AfficherErreur:

30      woups "frmReceptionMec", "AfficherProjet", Err, Erl
End Sub

Public Sub AfficherAchat(ByVal sUserID As String, ByVal sNoAchat As String)

5       On Error GoTo AfficherErreur

10      m_sUserID = sUserID

15      m_sNoAchat = sNoAchat

20      Call Me.Show(vbModal)

25      Exit Sub

AfficherErreur:

30      woups "frmReceptionMec", "AfficherAchat", Err, Erl
End Sub


Private Sub lvwProjet_KeyDown(KeyCode As Integer, Shift As Integer)

5       On Error GoTo AfficherErreur

10      If KeyCode = vbKeyReturn Then
15        If m_eType = PROJET Then
20          Call ReceptionProjet
25        Else
30          Call ReceptionAchat
35        End If
40      End If

45      Exit Sub

AfficherErreur:

50      woups "frmReceptionMec", "lvwProjet_KeyDown", Err, Erl
End Sub

Private Sub mvwDateRequise_DateClick(ByVal DateClicked As Date)

5       On Error GoTo AfficherErreur

10      txtDateRequise.Text = ConvertDate(DateClicked)

        'Enlever le calendrier
15      mvwDateRequise.Visible = False

20      Exit Sub

AfficherErreur:

25      woups "frmReceptionMec", "mvwDateRequise_DateClick", Err, Erl
End Sub

Private Sub mvwDateRequise_LostFocus()

5       On Error GoTo AfficherErreur

10      mvwDateRequise.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmReceptionMec", "mvwDateRequise_LostFocus", Err, Erl
End Sub

Private Sub mvwReception_DateClick(ByVal DateClicked As Date)

5       On Error GoTo AfficherErreur

10      txtDateReception.Text = ConvertDate(DateClicked)

        'Enlever le calendrier
15      mvwReception.Visible = False

20      Exit Sub

AfficherErreur:

25      woups "frmReceptionMec", "mvwReception_DateClick", Err, Erl
End Sub

Private Sub mvwReception_LostFocus()

5       On Error GoTo AfficherErreur

10      mvwReception.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmReceptionMec", "mvwReception_LostFocus", Err, Erl
End Sub

Private Sub Form_Click()

5       On Error GoTo AfficherErreur

10      mvwReception.Visible = False
15      mvwDateRequise.Visible = False

20      Exit Sub

AfficherErreur:

25      woups "frmReceptionMec", "Form_Click", Err, Erl
End Sub

Private Sub cmdDate_Click()

5       On Error GoTo AfficherErreur

        'Ouverture du calendrier
10      If txtDateReception.Text <> vbNullString Then
15        mvwReception.Value = txtDateReception.Text
20      Else
25        mvwReception.Value = Date
35      End If

40      mvwReception.Visible = True

45      Call mvwReception.SetFocus

50      Exit Sub

AfficherErreur:

55      woups "frmReceptionMec", "cmdDate_Click", Err, Erl
End Sub

Private Sub cmbType_Click()

5       On Error GoTo AfficherErreur

10      If cmbType.ListIndex = 0 Then
15        m_eType = PROJET

20        Call RemplirComboProjet
25      Else
30        m_eType = ACHAT

35        Call RemplirComboAchat
40      End If

45      If fraPiecesNonRecues.Visible = True Then
50        If m_eType = ACHAT Then
55          chkProjetAchat.Caption = "No achat : "
60        Else
65          chkProjetAchat.Caption = "No projet : "
70        End If
75      End If

80      Exit Sub

AfficherErreur:

85      woups "frmReceptionMec", "cmbType_Click", Err, Erl
End Sub

Private Sub cmdImprimer_Click()

5       On Error GoTo AfficherErreur

10      Call ImprimerReception

15      Exit Sub

AfficherErreur:

20      woups "frmReceptionMec", "cmdImprimer_Click", Err, Erl
End Sub

Private Sub ImprimerReception()

5       On Error GoTo AfficherErreur

10      If m_eType = ACHAT Then
15        Call frmChoixDateImpressionReception.Afficher(txtnoprojet.Text, MECANIQUE, ACHAT)
20      Else
25        Call frmChoixDateImpressionReception.Afficher(txtnoprojet.Text, MECANIQUE, PROJET)
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmReceptionMec", "ImprimerReception", Err, Erl
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        'CTRL-I pour imprimer
      
5       On Error GoTo AfficherErreur

10      If Shift = vbCtrlMask Then
15        If KeyCode = vbKeyI Then
20          Call ImprimerReception
25        End If
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmReceptionMec", "Form_KeyDown", Err, Erl
End Sub
