VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRetourMateriel 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retour de matériel"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmRetourMateriel.frx":0000
   ScaleHeight     =   5760
   ScaleWidth      =   8250
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   6840
      TabIndex        =   19
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdEnregistrer 
      Caption         =   "Enregistrer"
      Height          =   375
      Left            =   5520
      TabIndex        =   18
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Frame fraAjout 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   7935
      Begin VB.CheckBox chkMecanique 
         BackColor       =   &H00000000&
         Caption         =   "Mécanique"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   2520
         Width           =   1575
      End
      Begin VB.ComboBox cmbEmployé 
         Height          =   315
         Left            =   960
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtQte 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   1680
         Width           =   975
      End
      Begin VB.Frame fraRecherche 
         BackColor       =   &H00000000&
         Caption         =   "Recherche"
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
         Height          =   3135
         Left            =   2640
         TabIndex        =   11
         Top             =   840
         Width           =   5295
         Begin MSComctlLib.ListView lvwRecherche 
            Height          =   1935
            Left            =   120
            TabIndex        =   17
            Top             =   1080
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   3413
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "No Item"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Manufacturier"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.TextBox txtRecherche 
            Height          =   285
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   1575
         End
         Begin VB.ComboBox cmbRecherche 
            Height          =   315
            ItemData        =   "frmRetourMateriel.frx":2F0D
            Left            =   2160
            List            =   "frmRetourMateriel.frx":2F1A
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton cmdRechercher 
            Caption         =   "Afficher"
            Height          =   375
            Left            =   3960
            TabIndex        =   16
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Texte à rechercher : "
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Rechercher dans : "
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2160
            TabIndex        =   13
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.TextBox txtNoItem 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   1200
         Width           =   1335
      End
      Begin MSMask.MaskEdBox mskNoProjet 
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   2160
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "#####-##"
         PromptChar      =   "_"
      End
      Begin VB.Label lblprojet 
         BackColor       =   &H00000000&
         Caption         =   "No. Projet :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Employé : "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "---->"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Qté retournée :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Item : "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmRetourMateriel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const I_CMB_NO_ITEM As Integer = 0
Private Const I_CMB_DESCRIPTION As Integer = 1
Private Const I_CMB_MANUFACTURIER As Integer = 2

Private Const I_LVW_RECHERCHE_NO_ITEM As Integer = 0
Private Const I_LVW_RECHERCHE_DESCRIPTION As Integer = 1
Private Const I_LVW_RECHERCHE_MANUFACTURIER As Integer = 2

Private Enum enumExtra
 AUCUN_EXTRA = 0
 EXTRA_CHARGEABLE = 1
 EXTRA_NON_CHARGEABLE = 2
End Enum

Private m_eCatalogue As enumCatalogue

Private Sub cmdEnregistrer_Click()

 On Error GoTo Oups

 Dim rstInv As ADODB.Recordset
 Dim rstSortie As ADODB.Recordset
 Dim rstProjet As ADODB.Recordset
 Dim rstHistInv As ADODB.Recordset
 Dim rstInitiale As ADODB.Recordset

 If txtNoItem.Text <> "" Then
 If IsNumeric(txtQte.Text) Then
 If mskNoProjet.Text <> "_____-__" And mskNoProjet.Text <> "M_____-__" Then
 If ProjetExiste = True Then
 If VerifierSortie = True Then
  Set rstProjet = New ADODB.Recordset

  If m_eCatalogue = ELECTRIQUE Then
  Call rstProjet.Open("SELECT Modification, Par FROM GrbProjetElec WHERE IDProjet = '" & mskNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
  Else
  Call rstProjet.Open("SELECT Modification, Par FROM GrbProjetMec WHERE IDProjet = '" & mskNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
  End If

  If rstProjet.Fields("Modification") = False Then
  Set rstInv = New ADODB.Recordset

 If m_eCatalogue = ELECTRIQUE Then
 Call rstInv.Open("SELECT * FROM GrbInventaireElec WHERE NoItem = '" & Replace(txtNoItem.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstInv.Open("SELECT * FROM GrbInventaireMec WHERE NoItem = '" & Replace(txtNoItem.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
 End If

 If Not rstInv.EOF Then
 Set rstSortie = New ADODB.Recordset

 Call rstSortie.Open("SELECT * FROM GrbSortieMatériel", g_connData, adOpenDynamic, adLockOptimistic)

 Call rstSortie.AddNew

 rstSortie.Fields("Qté") = "-" & Abs(txtQte.Text)
 rstSortie.Fields("Nom") = cmbemployé.Text
 rstSortie.Fields("NoProjet") = mskNoProjet.Text
 rstSortie.Fields("NoItem") = txtNoItem.Text
 rstSortie.Fields("Date") = ConvertDate(Date)

 If m_eCatalogue = ELECTRIQUE Then
 rstSortie.Fields("Type") = "E"
 Else
 rstSortie.Fields("Type") = "M"
 End If

1  Call rstSortie.Update

 Call rstSortie.Close
 Set rstSortie = Nothing

 rstInv.Fields("QuantitéStock") = CDbl(rstInv.Fields("QuantitéStock")) + CDbl(Abs(txtQte.Text))

 Call rstInv.Update

 Set rstHistInv = New ADODB.Recordset

 If m_eCatalogue = ELECTRIQUE Then
 Call rstHistInv.Open("SELECT * FROM GrbInventaireElecModif", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstHistInv.Open("SELECT * FROM GrbInventaireMecModif", g_connData, adOpenDynamic, adLockOptimistic)
 End If

 Call rstHistInv.AddNew

 rstHistInv.Fields("Date") = ConvertDate(Date)
 rstHistInv.Fields("IDProjet") = mskNoProjet.Text
 rstHistInv.Fields("NoItem") = txtNoItem.Text
 rstHistInv.Fields("Quantité") = Abs(txtQte.Text)

 Set rstInitiale = New ADODB.Recordset

 Call rstInitiale.Open("SELECT * FROM GrbEmployés WHERE NoEmploye = " & cmbemployé.ItemData(cmbemployé.ListIndex), g_connData, adOpenDynamic, adLockOptimistic)

 rstHistInv.Fields("User") = rstInitiale.Fields("Initiale")

 Call rstInitiale.Close
 Set rstInitiale = Nothing

 Call rstHistInv.Update

 Call rstHistInv.Close
 Set rstHistInv = Nothing

 Call AjouterDansProjet(mskNoProjet.Text, AUCUN_EXTRA, "")

 Call rstProjet.Close

 If m_eCatalogue = ELECTRIQUE Then
 If Right$(mskNoProjet.Text, 2) >= 61 And Right$(mskNoProjet.Text, 2) <= 80 Then
 Call rstProjet.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & mskNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

 Call AjouterDansProjet(Left$(mskNoProjet.Text, Len(mskNoProjet.Text) - 2) & rstProjet.Fields("LiaisonChargeable"), EXTRA_CHARGEABLE, Right$(mskNoProjet.Text, 2))

 Call rstProjet.Close
 Else
 If Right$(mskNoProjet.Text, 2) >= 81 And Right$(mskNoProjet.Text, 2) <=   Then
 Call AjouterDansProjet(Left$(mskNoProjet.Text, Len(mskNoProjet.Text) - 2) & Right$("0" & Right$(mskNoProjet.Text, 2) - 80, 2), EXTRA_NON_CHARGEABLE, "")
 End If
 End If
 Else
 If Right$(mskNoProjet.Text, 2) >= 61 And Right$(mskNoProjet.Text, 2) <= 80 Then
 Call rstProjet.Open("SELECT * FROM GrbProjetMec WHERE IDProjet = '" & mskNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

 Call AjouterDansProjet(Left$(mskNoProjet.Text, Len(mskNoProjet.Text) - 2) & rstProjet.Fields("LiaisonChargeable"), EXTRA_CHARGEABLE, Right$(mskNoProjet.Text, 2))

 Call rstProjet.Close
 Else
4 If Right$(mskNoProjet.Text, 2) >= 81 And Right$(mskNoProjet.Text, 2) <=   Then
4 Call AjouterDansProjet(Left$(mskNoProjet.Text, Len(mskNoProjet.Text) - 2) & Right$("0" & Right$(mskNoProjet.Text, 2) - 80, 2), EXTRA_NON_CHARGEABLE, "")
4 End If
4 End If
4 End If

4 Call MsgBox("Le retour de matériel a été enregistrée!", vbOKOnly, "Erreur")

4 Call ViderChamps
4 Else
4 Call MsgBox("Cette pièce n'existe pas dans l'inventaire!", vbOKOnly, "Erreur")
4 End If

4 Call rstInv.Close
4  Set rstInv = Nothing
4  Else
4  Call MsgBox("Ce projet est en modification par " & rstProjet.Fields("Par") & "!", vbOKOnly, "Erreur")

4  Call rstProjet.Close
4  End If

4  Set rstProjet = Nothing
4  Else
4  Call MsgBox("Pas assez de pièces ont été sortie pour en retourner " & txtQte.Text & "!", vbOKOnly, "Erreur")
50 End If
 Else
 Call MsgBox("Projet inexistant!", vbOKOnly, "Erreur")
 End If
 Else
 Call MsgBox("Le numéro de projet est obligatoire!", vbOKOnly, "Erreur")
 End If
 Else
 Call MsgBox("Quantité non numérique!", vbOKOnly, "Erreur")
 End If
 Else
 Call MsgBox("Le numéro d'item est obligatoire!", vbOKOnly, "Erreur")
5  End If

5  Exit Sub

Oups:

5  wOups "frmRetourMateriel", "cmdEnregistrer_Click", Err, Err.number, Err.Description
End Sub

Private Sub Cmdfermer_Click()

 On Error GoTo Oups

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmRetourMateriel", "cmdFermer_Click", Err, Err.number, Err.Description
End Sub

Private Sub ViderChamps()

 On Error GoTo Oups

 Dim iCompteur As Integer

 txtNoItem.Text = ""
 txtQte.Text = ""
 txtRecherche.Text = ""
 cmbRecherche.ListIndex = 0

 chkMecanique.Value = vbUnchecked

 mskNoProjet.Text = "_____-__"

 For iCompteur = 0 To cmbemployé.ListCount - 1
 If cmbemployé.LIST(iCompteur) = g_sEmploye Then
 cmbemployé.ListIndex = iCompteur

  Exit For
  End If
  Next

  Exit Sub

Oups:

  wOups "frmRetourMateriel", "ViderChamps", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewRecherche()

 On Error GoTo Oups

 Dim rstInv As ADODB.Recordset
 Dim itmInv As ListItem
 Dim sWhere As String

 Screen.MousePointer = vbHourglass

 Call lvwRecherche.ListItems.Clear

 Select Case cmbRecherche.ListIndex
 Case I_CMB_NO_ITEM: sWhere = "Instr(1,NoItem,'" & txtRecherche.Text & "') > 0"
 Case I_CMB_DESCRIPTION: sWhere = "Instr(1,Description,'" & txtRecherche.Text & "') > 0"
 Case I_CMB_MANUFACTURIER: sWhere = "Instr(1,Manufacturier,'" & txtRecherche.Text & "') > 0"
 End Select

 Set rstInv = New ADODB.Recordset

  If m_eCatalogue = ELECTRIQUE Then
  Call rstInv.Open("SELECT * FROM GrbInventaireElec WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
  Else
  Call rstInv.Open("SELECT * FROM GrbInventaireMec WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
  End If

  Do While Not rstInv.EOF
  Set itmInv = lvwRecherche.ListItems.Add

  itmInv.Text = rstInv.Fields("NoItem")
itmInv.SubItems(I_LVW_RECHERCHE_DESCRIPTION) = rstInv.Fields("Description")
1 itmInv.SubItems(I_LVW_RECHERCHE_MANUFACTURIER) = rstInv.Fields("Manufacturier")

 Call rstInv.MoveNext
Loop

Call rstInv.Close
Set rstInv = Nothing

Screen.MousePointer = vbDefault

If lvwRecherche.ListItems.count = 0 Then
 Call MsgBox("Aucun enregistrement trouvé!", vbOKOnly, "Erreur")
End If

Exit Sub

Oups:

wOups "frmRetourMateriel", "RemplirListViewRecherche", Err, Err.number, Err.Description
End Sub

Private Sub cmdRechercher_Click()

 On Error GoTo Oups

 If txtRecherche.Text <> "" Then
 Call RemplirListViewRecherche
 Else
 Call MsgBox("Rien à rechercher!", vbOKOnly, "Erreur")
 End If

 Exit Sub

Oups:

 wOups "frmRetourMateriel", "cmdRechercher_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 Call RemplirComboEmployes

 Call ViderChamps

 Exit Sub

Oups:

 wOups "frmRetourMateriel", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboEmployes()

 On Error GoTo Oups

 Dim rstEmploye As ADODB.Recordset

 Set rstEmploye = New ADODB.Recordset

 Call rstEmploye.Open("SELECT * FROM GrbEmployés WHERE Actif = True", g_connData, adOpenDynamic, adLockOptimistic)

 Do While Not rstEmploye.EOF
 Call cmbemployé.AddItem(rstEmploye.Fields("Employe"))

 cmbemployé.ItemData(cmbemployé.newIndex) = rstEmploye.Fields("NoEmploye")

 Call rstEmploye.MoveNext
 Loop

 Call rstEmploye.Close
 Set rstEmploye = Nothing

  Exit Sub

Oups:

  wOups "frmRetourMateriel", "RemplirComboEmployes", Err, Err.number, Err.Description
End Sub

Private Sub lvwRecherche_DblClick()

 On Error GoTo Oups

 If lvwRecherche.ListItems.count > 0 Then
 txtNoItem.Text = lvwRecherche.SelectedItem.Text
 End If

 Exit Sub

Oups:

 wOups "frmRetourMateriel", "lvwRecherche_DblClick", Err, Err.number, Err.Description
End Sub

Private Sub mskNoProjet_Change()

 On Error GoTo Oups

 If fraAjout.Visible = True Then
 If InStr(1, mskNoProjet.Text, "_") = 0 Then
 Call ProjetExiste
 End If
 End If

 Exit Sub

Oups:

 wOups "frmRetourMateriel", "mskNoProjet_Change", Err, Err.number, Err.Description
End Sub

Private Function ProjetExiste() As Boolean

 On Error GoTo Oups

 Dim rstProjSoum As ADODB.Recordset
 
 If Right$(mskNoProjet.Text, 2) >= 51 And Right$(mskNoProjet.Text, 2) <=   Then
 Set rstProjSoum = New ADODB.Recordset

 Call rstProjSoum.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & mskNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 If Not rstProjSoum.EOF Then
 If rstProjSoum.Fields("Ouvert") = False Then
 Call MsgBox("Ce projet n'est pas ouvert!", vbOKOnly, "Erreur")

 ProjetExiste = False
 Else
 ProjetExiste = True
  End If
  Else
  Call MsgBox("Projet inexistant!", vbOKOnly, "Erreur")

  ProjetExiste = False
  End If

  Call rstProjSoum.Close
  Set rstProjSoum = Nothing
  Else
Call MsgBox("Impossible de faire une sortie de matériel sur ce numéro!", vbOKOnly, "Erreur")

1 ProjetExiste = False
End If
 
Exit Function

Oups:

wOups "frmRetourMateriel", "ProjetExiste", Err, Err.number, Err.Description
End Function

Private Function VerifierSortie() As Boolean

 On Error GoTo Oups

 Dim rstProjet As ADODB.Recordset
 Dim rstSection As ADODB.Recordset
 Dim sIDSection As String
 Dim dblQteProjet As Double
 
 Set rstSection = New ADODB.Recordset

 If m_eCatalogue = ELECTRIQUE Then
 Call rstSection.Open("SELECT * FROM GrbSoumProjSectionElec WHERE NomSectionFR = 'Externe'", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstSection.Open("SELECT * FROM GrbSoumProjSectionMec WHERE NomSectionFR = 'Externe'", g_connData, adOpenDynamic, adLockOptimistic)
 End If

  sIDSection = rstSection.Fields("IDSection")

  Call rstSection.Close
  Set rstSection = Nothing

  Set rstProjet = New ADODB.Recordset
 
  Call rstProjet.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & mskNoProjet.Text & "' AND IDSection = " & sIDSection & " AND SousSection = 'PAS DE SOUS-SECTION' AND NumItem = '" & Replace(txtNoItem.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
  If Not rstProjet.EOF Then
  Do While Not rstProjet.EOF
  dblQteProjet = dblQteProjet + rstProjet.Fields("Qté")

 Call rstProjet.MoveNext
1 Loop

 If dblQteProjet >= Abs(txtQte.Text) Then
 VerifierSortie = True
 Else
 Call MsgBox("Il n'y a pas assez de " & txtNoItem.Text & " dans le projet " & mskNoProjet.Text & " pour en enlever " & Abs(txtQte.Text), vbOKOnly, "Erreur")

 VerifierSortie = False
 End If
Else
 Call MsgBox("La pièce " & txtNoItem.Text & " n'a pas été sortie pour le projet " & mskNoProjet.Text, vbOKOnly, "Erreur")

 VerifierSortie = False
End If

1  Call rstProjet.Close
Set rstProjet = Nothing
 
 Exit Function

Oups:

wOups "frmRetourMateriel", "VerifierSortie", Err, Err.number, Err.Description
End Function


Private Sub chkMecanique_Click()

 On Error GoTo Oups

 Dim sTampon As String

 sTampon = mskNoProjet.Text
 
 'dépendant si coché mécanique affiche le mask
 If chkMecanique.Value = vbChecked Then
 mskNoProjet.mask = "\M#####-##"
 'ajoute le M
 If Len(sTampon) =   Then
 mskNoProjet.Text = "M" + sTampon
 End If
 Else
 'enleve le m
 mskNoProjet.mask = "#####-##"
 mskNoProjet.Text = Right$(sTampon, 9)
  End If
 
  If fraAjout.Visible = True Then
  Call mskNoProjet.SetFocus
  End If

  Exit Sub

Oups:

  wOups "frmRetourMateriel", "chkMecanique_Click", Err, Err.number, Err.Description
End Sub

Public Sub Afficher(ByVal eCatalogue As enumCatalogue)
 
 On Error GoTo Oups

 m_eCatalogue = eCatalogue

 Call Unload(frmChoixRetourMateriel)

 Call Me.Show

 Exit Sub

Oups:

 wOups "frmRetourMateriel", "Afficher", Err, Err.number, Err.Description
End Sub

Private Sub AjouterDansProjet(ByVal sNoProjet As String, ByVal eExtra As enumExtra, ByVal sProvenance As String)
 
 On Error GoTo Oups

 Dim rstProjet As ADODB.Recordset
 Dim rstPiece As ADODB.Recordset
 Dim rstSection As ADODB.Recordset
 Dim rstInv As ADODB.Recordset
 Dim iCompteur As Integer
 Dim sSection As String
 Dim bSkip As Boolean
 Dim sIDSection As String
 Dim sOrdre As String
 Dim sProfit As String
 
  Set rstProjet = New ADODB.Recordset
  Set rstSection = New ADODB.Recordset
 
  If m_eCatalogue = ELECTRIQUE Then
  Call rstSection.Open("SELECT * FROM GrbSoumProjSectionElec WHERE NomSectionFR = 'Externe'", g_connData, adOpenDynamic, adLockOptimistic)

  Call rstProjet.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

  sProfit = rstProjet.Fields("Profit")

  Call rstProjet.Close
  Else
Call rstSection.Open("SELECT * FROM GrbSoumProjSectionMec WHERE NomSectionFR = 'Externe'", g_connData, adOpenDynamic, adLockOptimistic)

1 Call rstProjet.Open("SELECT * FROM GrbProjetMec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockBatchOptimistic)

 sProfit = rstProjet.Fields("Profit")

 Call rstProjet.Close
End If

sIDSection = rstSection.Fields("IDSection")
sOrdre = rstSection.Fields("Ordre")

Call rstSection.Close
Set rstSection = Nothing

 'Ouverture du recordset sur le projet original
Set rstPiece = New ADODB.Recordset
 
Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & sNoProjet & "' AND IDSection = " & sIDSection & " AND SousSection = 'PAS DE SOUS-SECTION' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

If Not rstPiece.EOF Then
Call rstPiece.MoveLast

 iCompteur = rstPiece.Fields("NuméroLigne") + 1
 Else
 iCompteur = 1
 End If

Set rstInv = New ADODB.Recordset

 If m_eCatalogue = ELECTRIQUE Then
1  Call rstInv.Open("SELECT * FROM GrbInventaireElec WHERE NoItem = '" & Replace(txtNoItem.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstInv.Open("SELECT * FROM GrbInventaireMec WHERE NoItem = '" & Replace(txtNoItem.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
End If

Call rstPiece.AddNew

rstPiece.Fields("IDProjet") = sNoProjet
 
If m_eCatalogue = ELECTRIQUE Then
 rstPiece.Fields("Type") = "E"
Else
 rstPiece.Fields("Type") = "M"
End If

rstPiece.Fields("Visible") = True

rstPiece.Fields("Facturation") = ""
 
rstPiece.Fields("IDSection") = sIDSection
2  rstPiece.Fields("NumItem") = rstInv.Fields("NoItem")
rstPiece.Fields("Qté") = "-" & Abs(txtQte.Text)
2  rstPiece.Fields("Desc_FR") = rstInv.Fields("Description")
rstPiece.Fields("Desc_EN") = ""
2  rstPiece.Fields("Manufact") = rstInv.Fields("Manufacturier")
rstPiece.Fields("Prix_list") = Conversion(rstInv.Fields("Prix liste"), MODE_PAS_FORMAT, 4)
30 rstPiece.Fields("Escompte") = Conversion(rstInv.Fields("Escompte"), MODE_PAS_FORMAT)
rstPiece.Fields("Prix_net") = Conversion(rstInv.Fields("Prix net"), MODE_PAS_FORMAT, 4)
rstPiece.Fields("OrdreSection") = sOrdre
rstPiece.Fields("NuméroLigne") = iCompteur
 
rstPiece.Fields("IDFRS") = 717
 
rstPiece.Fields("Prix_Total") = Conversion(rstInv.Fields("Prix net") * rstPiece.Fields("Qté") * CDbl(sProfit), MODE_PAS_FORMAT)
rstPiece.Fields("Profit_argent") = Conversion(rstPiece.Fields("Prix_Total") - (rstInv.Fields("Prix net") * rstPiece.Fields("Qté")), MODE_PAS_FORMAT)
rstPiece.Fields("SousSection") = "PAS DE SOUS-SECTION"
 
rstPiece.Fields("PrixOrigine") = Replace(Round(CDbl(Replace(rstInv.Fields("Prix liste"), ".", ",")), 2), ".", ",")

Select Case eExtra
 Case EXTRA_CHARGEABLE:
 rstPiece.Fields("PieceExtraChargeable") = True
 rstPiece.Fields("PieceExtraNonChargeable") = False

Case EXTRA_NON_CHARGEABLE:
 rstPiece.Fields("PieceExtraChargeable") = False
 rstPiece.Fields("PieceExtraNonChargeable") = True

 Case AUCUN_EXTRA:
 rstPiece.Fields("PieceExtraChargeable") = False
 rstPiece.Fields("PieceExtraNonChargeable") = False
3  End Select

 rstPiece.Fields("Provenance") = sProvenance

40 Call rstPiece.Update

Call rstPiece.Close

4 Call rstInv.Close
4 Set rstInv = Nothing

4 rstPiece.CursorLocation = adUseServer

4 Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & sNoProjet & "' AND NuméroLigne >= " & iCompteur & " ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

 'Tant qu'il y a des enregistrements dans le recordset
4 Do While Not rstPiece.EOF
4 If bSkip = False Then
4 bSkip = True
4 Else
4 rstPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne") + 1

4 Call rstPiece.Update
4  End If

4  Call rstPiece.MoveNext
4  Loop

4  Call rstPiece.Close
4  Set rstPiece = Nothing

4  If m_eCatalogue = ELECTRIQUE Then
4  Call CalculerTempsMecRecordset(sNoProjet)

4  Call CalculerTotalRecordsetElec(sNoProjet)
50 Else
5 Call CalculerTotalRecordsetMec(sNoProjet)
 End If

 Exit Sub

Oups:

 wOups "frmRetourMateriel", "AjouterDansProjet", Err, Err.number, Err.Description
End Sub

Private Sub CalculerTempsMecRecordset(ByVal sNoProjet As String)

 On Error GoTo Oups

 Dim rstProjet As ADODB.Recordset
 Dim rstPiece As ADODB.Recordset
 Dim dblTempsMec As Double

 'Ouverture des tables
 Set rstProjet = New ADODB.Recordset
 Set rstPiece = New ADODB.Recordset

 Call rstProjet.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

 Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet ='" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

 'Pour chaque enregistrement du recordset
 Do While Not rstPiece.EOF
 'Si le temps total n'est pas vide
 If Trim(rstPiece.Fields("Temps_total")) <> vbNullString Then
 'On additionne le temps
 dblTempsMec = dblTempsMec + CDbl(Replace(Trim(rstPiece.Fields("Temps_total")), ".", ","))
  End If

  Call rstPiece.MoveNext
  Loop
 
  rstProjet.Fields("temp_mec") = Replace(dblTempsMec / 10, ".", ",")

  Call rstProjet.Update

  Call rstPiece.Close
  Set rstPiece = Nothing

  Call rstProjet.Close
10 Set rstProjet = Nothing

Exit Sub

Oups:

wOups "frmRetourMateriel", "CalculerTempsMecRecordset", Err, Err.number, Err.Description
End Sub

Private Sub CalculerTotalRecordsetElec(ByVal sNoProjet As String)

 On Error GoTo Oups

 'Méthode pour calculer le prix
 Dim dblManuel As Double
 Dim dblCopies As Double
 Dim dblTempsDessin As Double
 Dim dblTempsProg As Double
 Dim dblTempsMec As Double
 Dim dblTempsElec As Double
 Dim dblTempsTest As Double
 Dim dblTempsVision As Double
 Dim dblPrixPieces As Double
 Dim dblPrixTotal As Double
  Dim dblCommission As Double
  Dim dblTotalTemps As Double
  Dim dblProfit As Double
  Dim dblTotalManuel As Double
  Dim dblTotalPieceImprevue As Double
  Dim dblGrandTotal As Double
  Dim rstProjet As ADODB.Recordset
  Dim rstPiece As ADODB.Recordset
10 Dim rstConfig As ADODB.Recordset
Dim sCommission As String
Dim sCopieManuel As String
Dim sImprevue As String

Set rstProjet = New ADODB.Recordset
Set rstPiece = New ADODB.Recordset
Set rstConfig = New ADODB.Recordset

Call rstConfig.Open("SELECT * FROM GrbConfig", g_connData, adOpenDynamic, adLockOptimistic)

sCommission = rstConfig.Fields("Commission")
sImprevue = rstConfig.Fields("Imprévus")
sCopieManuel = rstConfig.Fields("PrixPagesManuel")

Call rstConfig.Close
1  Set rstConfig = Nothing

Call rstProjet.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If Not rstProjet.EOF Then
 'Validation du nombre de pages
 If IsNumeric(rstProjet.Fields("Manuel")) Then
 dblManuel = rstProjet.Fields("Manuel")
 Else
 dblManuel = 0
1  End If

 'Validation du nombre de copies
 If IsNumeric(rstProjet.Fields("copies")) Then
 dblCopies = rstProjet.Fields("copies")
 Else
 dblCopies = 0
 End If

 'Validation des heures de dessin
 If IsNumeric(rstProjet.Fields("temp_dessin")) Then
 dblTempsDessin = rstProjet.Fields("temp_dessin")
 Else
 dblTempsDessin = 0
 End If

 'Validation des heures de prog
 If IsNumeric(rstProjet.Fields("temp_prog")) Then
 dblTempsProg = rstProjet.Fields("temp_prog")
Else
 dblTempsProg = 0
End If

 'Validation des heures de mec
 If rstProjet.Fields("SansTemps") = True Then
 dblTempsMec = 0
 Else
 dblTempsMec = rstProjet.Fields("temp_mec")
 End If

 'Validation des heures d'élec
If IsNumeric(rstProjet.Fields("temp_elec")) Then
dblTempsElec = rstProjet.Fields("temp_elec")
 Else
 dblTempsElec = 0
 End If

 'Validation des heures de test
 If IsNumeric(rstProjet.Fields("temp_test")) Then
 dblTempsTest = rstProjet.Fields("temp_test")
 Else
 dblTempsTest = 0
 End If

 'Validation des heures de vision
 If IsNumeric(rstProjet.Fields("temp_vision")) Then
 dblTempsVision = rstProjet.Fields("temp_vision")
Else
 dblTempsVision = 0
End If

 Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & sNoProjet & "' AND Type = 'E'", g_connData, adOpenDynamic, adLockOptimistic)

 'Pour chaque élément du recordset
Do While Not rstPiece.EOF
 If Trim(rstPiece.Fields("Prix_total")) <> vbNullString Then
 'On additionne le prix total
 dblPrixPieces = dblPrixPieces + rstPiece.Fields("Prix_total")

 'On additionne le profit
 dblProfit = dblProfit + rstPiece.Fields("Profit_Argent")
4 End If

4 Call rstPiece.MoveNext
4 Loop

4 Call rstPiece.Close
4 Set rstPiece = Nothing

 'On additionne les (temps * taux)
4 dblTotalTemps = (dblTempsDessin * CDbl(rstProjet.Fields("taux_dessin"))) + (dblTempsProg * CDbl(rstProjet.Fields("taux_prog"))) + (dblTempsMec * CDbl(rstProjet.Fields("taux_mec"))) + (dblTempsElec * CDbl(rstProjet.Fields("taux_elec"))) + (dblTempsTest * CDbl(rstProjet.Fields("taux_test"))) + (dblTempsVision * CDbl(rstProjet.Fields("taux_vision")))

4 dblTotalManuel = dblManuel * dblCopies * CDbl(sCopieManuel)

4 dblTotalPieceImprevue = dblPrixPieces * (1 + CDbl(sImprevue))

4 dblPrixTotal = dblTotalTemps + dblTotalManuel + dblTotalPieceImprevue

 'Le calcul de la commission est sur les manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps
4 dblCommission = dblPrixTotal * CDbl(sCommission)

 'Le prix total est le calcul des manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps + total de la commission
4 dblGrandTotal = dblPrixTotal + dblCommission

 'Format monétaires avec 2 chiffres après la virgule
4  rstProjet.Fields("Total_Commission") = Conversion(CStr(Round(dblCommission, 2)), MODE_ARGENT)
4  rstProjet.Fields("PrixTotal") = Conversion(CStr(Round(dblGrandTotal, 2)), MODE_ARGENT)
4  rstProjet.Fields("Total_Profit") = Conversion(CStr(Round(dblProfit, 2)), MODE_ARGENT)

4  Call rstProjet.Update
4  End If

4  Call rstProjet.Close
4  Set rstProjet = Nothing

4  Exit Sub

Oups:

50 wOups "frmRetourMateriel", "CalculerTotalRecordset", Err, Err.number, Err.Description
End Sub

Private Sub CalculerTotalRecordsetMec(ByVal sNoProjet As String)

 On Error GoTo Oups

 'Méthode pour calculer le prix
 Dim dblPrixPieces As Double
 Dim dblPrixTotal As Double
 Dim dblCommission As Double
 Dim dblTotalTemps As Double
 Dim dblProfit As Double
 Dim dblTotalManuel As Double
 Dim dblTotalImprevue As Double
 Dim dblGrandTotal As Double
 Dim dblTotalMachinage As Double
 Dim dblTotalCoupe As Double
  Dim dblTotalSoudure As Double
  Dim dblTotalAssemblage As Double
  Dim dblTotalPeinture As Double
  Dim dblTotalTest As Double
  Dim dblTotalDessin As Double
  Dim dblTotalFormation As Double
  Dim dblTotalInstallation As Double
  Dim dblHebergement As Double
10 Dim dblRepas As Double
Dim dblTransport As Double
Dim dblUniteMobile As Double
Dim dblPrixEmballage As Double
Dim dblTotalResteTemps As Double
Dim iNbrePersonne As Integer
Dim rstProjet As ADODB.Recordset
Dim rstPiece As ADODB.Recordset
Dim rstConfig As ADODB.Recordset
Dim sCommission As String
Dim sImprevue As String

Set rstConfig = New ADODB.Recordset

1  Call rstConfig.Open("SELECT * FROM GrbConfig", g_connData, adOpenDynamic, adLockOptimistic)

sCommission = rstConfig.Fields("Commission")
 sImprevue = rstConfig.Fields("Imprévus")

Call rstConfig.Close
 Set rstConfig = Nothing

Set rstProjet = New ADODB.Recordset
 Set rstPiece = New ADODB.Recordset

1  Call rstProjet.Open("SELECT * FROM GrbProjetMec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

 Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & sNoProjet & "' AND Type = 'M'", g_connData, adOpenDynamic, adLockOptimistic)

 'Pour chaque élément du recordset
 Do While Not rstPiece.EOF
 If Trim(rstPiece.Fields("Prix_total")) <> vbNullString Then
 'On additionne le prix total
 dblPrixPieces = dblPrixPieces + rstPiece.Fields("Prix_total") - rstPiece.Fields("Profit_Argent")
 
 'On additionne le profit
 dblProfit = dblProfit + rstPiece.Fields("Profit_Argent")
 End If

 Call rstPiece.MoveNext
Loop
 
 'Total des temps
dblTotalMachinage = CDbl(rstProjet.Fields("TempsMachinage")) * CDbl(rstProjet.Fields("TauxMachinage"))
dblTotalCoupe = CDbl(rstProjet.Fields("TempsCoupe")) * CDbl(rstProjet.Fields("TauxCoupePréparation"))
dblTotalSoudure = CDbl(rstProjet.Fields("TempsSoudure")) * CDbl(rstProjet.Fields("TauxAssemblageSoudure"))
dblTotalAssemblage = CDbl(rstProjet.Fields("TempsAssemblage")) * CDbl(rstProjet.Fields("TauxAssemblageSystèmes"))
2  dblTotalPeinture = CDbl(rstProjet.Fields("TempsPeinture")) * CDbl(rstProjet.Fields("TauxPeintureFinition"))
dblTotalTest = CDbl(rstProjet.Fields("TempsTest")) * CDbl(rstProjet.Fields("TauxTestsFinaux"))
2  dblTotalDessin = CDbl(rstProjet.Fields("TempsDessin")) * CDbl(rstProjet.Fields("TauxConceptionDessins"))
dblTotalFormation = CDbl(rstProjet.Fields("TempsFormation")) * CDbl(rstProjet.Fields("TauxFormation"))
2  dblTotalInstallation = CDbl(rstProjet.Fields("TempsInstallation")) * CDbl(rstProjet.Fields("TauxInstallation"))
 
dblTotalTemps = dblTotalMachinage + dblTotalCoupe + dblTotalSoudure + _
 dblTotalAssemblage + dblTotalPeinture + dblTotalTest + _
 dblTotalDessin + dblTotalFormation + dblTotalInstallation
 
2  If Not IsNull(rstProjet.Fields("NbrePersonne")) Then
 If Trim(rstProjet.Fields("NbrePersonne")) <> "" Then
 iNbrePersonne = Int(rstProjet.Fields("NbrePersonne"))
3 Else
 iNbrePersonne = 0
 End If
Else
 iNbrePersonne = 0
End If
 
Do While iNbrePersonne > 0
 If iNbrePersonne >= 2 Then
 dblHebergement = dblHebergement + CDbl(rstProjet.Fields("TempsHebergement")) * CDbl(rstProjet.Fields("TauxHebergement2"))
 
 iNbrePersonne = iNbrePersonne - 2
 Else
 dblHebergement = dblHebergement + CDbl(rstProjet.Fields("TempsHebergement")) * CDbl(rstProjet.Fields("TauxHebergement1"))
 
 iNbrePersonne = iNbrePersonne - 1
End If
Loop
 
3  If Not IsNull(rstProjet.Fields("NbrePersonne")) Then
 If Trim(rstProjet.Fields("NbrePersonne")) <> "" Then
 dblRepas = CDbl(rstProjet.Fields("TempsRepas")) * CDbl(rstProjet.Fields("TauxRepas")) * CDbl(rstProjet.Fields("NbrePersonne"))
 Else
 dblRepas = 0
4 End If
4 Else
4 dblRepas = 0
4 End If

4 dblTransport = CDbl(rstProjet.Fields("TempsTransport")) * CDbl(rstProjet.Fields("TauxTransport"))
4 dblUniteMobile = CDbl(rstProjet.Fields("TempsUniteMobile")) * CDbl(rstProjet.Fields("TauxUniteMobile"))

 'Correction d'un bug de Type Incompatible
4 If IsNumeric(rstProjet.Fields("PrixEmballage")) Then
4 dblPrixEmballage = CDbl(rstProjet.Fields("PrixEmballage"))
4 Else
4 dblPrixEmballage = 0
4 End If
 
4  dblTotalResteTemps = dblHebergement + dblRepas + dblTransport + dblUniteMobile + dblPrixEmballage
  
4  If IsNumeric(rstProjet.Fields("total_manuel")) Then
4  dblTotalManuel = CDbl(rstProjet.Fields("total_manuel"))
4  Else
4  dblTotalManuel = 0
4  End If
 
4  dblTotalImprevue = (dblPrixPieces + dblProfit) * CDbl(sImprevue)
 
4  dblPrixTotal = dblPrixPieces + dblProfit + dblTotalTemps + dblTotalImprevue + dblTotalManuel + dblTotalResteTemps
 
 'Le calcul de la commission est sur les manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps
50 dblCommission = dblPrixTotal * CDbl(sCommission)
 
 'Le prix total est le calcul des manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps + total de la commission
50 dblGrandTotal = dblPrixTotal + dblCommission
 
 'Format monétaires avec 2 chiffres après la virgule
 rstProjet.Fields("Total_Piece") = Conversion(CStr(Round(dblPrixPieces, 2)), MODE_ARGENT)
 rstProjet.Fields("Total_Temps") = Conversion(CStr(Round(dblTotalTemps, 2)), MODE_ARGENT)
 rstProjet.Fields("PrixTotal") = Conversion(CStr(Round(dblGrandTotal, 2)), MODE_ARGENT)
 rstProjet.Fields("total_Imprevue") = Conversion(CStr(Round(dblTotalImprevue, 2)), MODE_ARGENT)
 rstProjet.Fields("total_commission") = Conversion(CStr(Round(dblCommission, 2)), MODE_ARGENT)
 rstProjet.Fields("total_profit") = Conversion(CStr(Round(dblProfit, 2)), MODE_ARGENT)

 Call rstProjet.Update

 Call rstPiece.Close
 Set rstPiece = Nothing

 Call rstProjet.Close
5  Set rstProjet = Nothing

5  Exit Sub

Oups:

5  wOups "frmRetourMateriel", "CalculerTotalRecordset", Err, Err.number, Err.Description
End Sub
