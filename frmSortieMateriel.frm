VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSortieMateriel 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sortie de matériel"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmSortieMateriel.frx":0000
   ScaleHeight     =   5760
   ScaleWidth      =   8250
   StartUpPosition =   2  'CenterScreen
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
            ItemData        =   "frmSortieMateriel.frx":2F0D
            Left            =   2160
            List            =   "frmSortieMateriel.frx":2F1A
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
         Caption         =   "Qté sortie : "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   975
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
Attribute VB_Name = "frmSortieMateriel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const I_CMB_NO_ITEM                 As Integer = 0
Private Const I_CMB_DESCRIPTION             As Integer = 1
Private Const I_CMB_MANUFACTURIER           As Integer = 2

Private Const I_LVW_RECHERCHE_NO_ITEM       As Integer = 0
Private Const I_LVW_RECHERCHE_DESCRIPTION   As Integer = 1
Private Const I_LVW_RECHERCHE_MANUFACTURIER As Integer = 2

Private Enum enumExtra
  AUCUN_EXTRA = 0
  EXTRA_CHARGEABLE = 1
  EXTRA_NON_CHARGEABLE = 2
End Enum

Private m_eCatalogue As enumCatalogue

Private Sub cmdEnregistrer_Click()

5       On Error GoTo AfficherErreur

10      Dim rstInv      As ADODB.Recordset
15      Dim rstSortie   As ADODB.Recordset
20      Dim rstProjet   As ADODB.Recordset
25      Dim rstHistInv  As ADODB.Recordset
30      Dim rstInitiale As ADODB.Recordset

35      If txtNoItem.Text <> "" Then
40        If IsNumeric(txtQte.Text) Then
45          If mskNoProjet.Text <> "_____-__" And mskNoProjet.Text <> "M_____-__" Then
50            If ProjetExiste = True Then
55              Set rstProjet = New ADODB.Recordset

60              If m_eCatalogue = ELECTRIQUE Then
65                Call rstProjet.Open("SELECT Modification, Par FROM GRB_ProjetElec WHERE IDProjet = '" & mskNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
70              Else
75                Call rstProjet.Open("SELECT Modification, Par FROM GRB_ProjetMec WHERE IDProjet = '" & mskNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
80              End If

85              If rstProjet.Fields("Modification") = False Then
90                Set rstInv = New ADODB.Recordset

95                If m_eCatalogue = ELECTRIQUE Then
100                 Call rstInv.Open("SELECT * FROM GRB_InventaireElec WHERE NoItem = '" & Replace(txtNoItem.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
105               Else
110                 Call rstInv.Open("SELECT * FROM GRB_InventaireMec WHERE NoItem = '" & Replace(txtNoItem.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
115               End If

120               If Not rstInv.EOF Then
125                 Set rstSortie = New ADODB.Recordset

130                 Call rstSortie.Open("SELECT * FROM GRB_SortieMatériel", g_connData, adOpenDynamic, adLockOptimistic)

135                 Call rstSortie.AddNew

140                 rstSortie.Fields("Qté") = txtQte.Text
145                 rstSortie.Fields("Nom") = cmbemployé.Text
150                 rstSortie.Fields("NoProjet") = mskNoProjet.Text
155                 rstSortie.Fields("NoItem") = txtNoItem.Text
160                 rstSortie.Fields("Date") = ConvertDate(Date)

165                 If m_eCatalogue = ELECTRIQUE Then
170                   rstSortie.Fields("Type") = "E"
175                 Else
180                   rstSortie.Fields("Type") = "M"
185                 End If

190                 Call rstSortie.Update

195                 Call rstSortie.Close
200                 Set rstSortie = Nothing

205                 rstInv.Fields("QuantitéStock") = CDbl(rstInv.Fields("QuantitéStock")) - CDbl(txtQte.Text)

210                 Call rstInv.Update

215                 Set rstHistInv = New ADODB.Recordset

220                 If m_eCatalogue = ELECTRIQUE Then
225                   Call rstHistInv.Open("SELECT * FROM GRB_InventaireElecModif", g_connData, adOpenDynamic, adLockOptimistic)
230                 Else
235                   Call rstHistInv.Open("SELECT * FROM GRB_InventaireMecModif", g_connData, adOpenDynamic, adLockOptimistic)
240                 End If

245                 Call rstHistInv.AddNew

250                 rstHistInv.Fields("Date") = ConvertDate(Date)
255                 rstHistInv.Fields("IDProjet") = mskNoProjet.Text
260                 rstHistInv.Fields("NoItem") = txtNoItem.Text
265                 rstHistInv.Fields("Quantité") = "-" & Abs(txtQte.Text)

270                 Set rstInitiale = New ADODB.Recordset

275                 Call rstInitiale.Open("SELECT * FROM GRB_Employés WHERE NoEmploye = " & cmbemployé.ItemData(cmbemployé.ListIndex), g_connData, adOpenDynamic, adLockOptimistic)

280                 rstHistInv.Fields("User") = rstInitiale.Fields("Initiale")

285                 Call rstInitiale.Close
290                 Set rstInitiale = Nothing

295                 Call rstHistInv.Update

300                 Call rstHistInv.Close
305                 Set rstHistInv = Nothing

310                 Call AjouterDansProjet(mskNoProjet.Text, AUCUN_EXTRA, "")

315                 Call rstProjet.Close

320                 If m_eCatalogue = ELECTRIQUE Then
325                   If Right$(mskNoProjet.Text, 2) >= 61 And Right$(mskNoProjet.Text, 2) <= 80 Then
330                     Call rstProjet.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & mskNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

335                     Call AjouterDansProjet(Left$(mskNoProjet.Text, Len(mskNoProjet.Text) - 2) & rstProjet.Fields("LiaisonChargeable"), EXTRA_CHARGEABLE, Right$(mskNoProjet.Text, 2))

340                     Call rstProjet.Close
345                   Else
350                     If Right$(mskNoProjet.Text, 2) >= 81 And Right$(mskNoProjet.Text, 2) <= 98 Then
355                       Call AjouterDansProjet(Left$(mskNoProjet.Text, Len(mskNoProjet.Text) - 2) & Right$("0" & Right$(mskNoProjet.Text, 2) - 80, 2), EXTRA_NON_CHARGEABLE, "")
360                     End If
365                   End If
370                 Else
375                   If Right$(mskNoProjet.Text, 2) >= 61 And Right$(mskNoProjet.Text, 2) <= 80 Then
380                     Call rstProjet.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & mskNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

385                     Call AjouterDansProjet(Left$(mskNoProjet.Text, Len(mskNoProjet.Text) - 2) & rstProjet.Fields("LiaisonChargeable"), EXTRA_CHARGEABLE, Right$(mskNoProjet.Text, 2))

390                     Call rstProjet.Close
395                   Else
400                     If Right$(mskNoProjet.Text, 2) >= 81 And Right$(mskNoProjet.Text, 2) <= 98 Then
405                       Call AjouterDansProjet(Left$(mskNoProjet.Text, Len(mskNoProjet.Text) - 2) & Right$("0" & Right$(mskNoProjet.Text, 2) - 80, 2), EXTRA_NON_CHARGEABLE, "")
410                     End If
415                   End If
420                 End If

425                 Set rstProjet = Nothing

430                 Call MsgBox("La sortie de matériel a été enregistrée!", vbOKOnly, "Erreur")

435                 Call ViderChamps
440               Else
445                 Call MsgBox("Cette pièce n'existe pas dans l'inventaire!", vbOKOnly, "Erreur")
450               End If

455               Call rstInv.Close
460               Set rstInv = Nothing
465             Else
470               Call MsgBox("Ce projet est en modification par " & rstProjet.Fields("Par") & "!", vbOKOnly, "Erreur")

475               Call rstProjet.Close
480               Set rstProjet = Nothing
485             End If
490           Else
495             Call MsgBox("Projet inexistant!", vbOKOnly, "Erreur")
500           End If
505         Else
510           Call MsgBox("Le numéro de projet est obligatoire!", vbOKOnly, "Erreur")
515         End If
520       Else
525         Call MsgBox("Quantité non numérique!", vbOKOnly, "Erreur")
530       End If
535     Else
540       Call MsgBox("Le numéro d'item est obligatoire!", vbOKOnly, "Erreur")
545     End If

550     Exit Sub

AfficherErreur:

555     woups "frmSortieMateriel", "cmdEnregistrer_Click", Err, Erl
End Sub

Private Sub Cmdfermer_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmSortieMateriel", "cmdFermer_Click", Err, Erl
End Sub

Private Sub ViderChamps()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer

15      txtNoItem.Text = ""
20      txtQte.Text = ""
25      txtRecherche.Text = ""
30      cmbRecherche.ListIndex = 0

35      chkMecanique.Value = vbUnchecked

40      mskNoProjet.Text = "_____-__"

45      For iCompteur = 0 To cmbemployé.ListCount - 1
50        If cmbemployé.LIST(iCompteur) = g_sEmploye Then
55          cmbemployé.ListIndex = iCompteur

60          Exit For
65        End If
70      Next

75      Exit Sub

AfficherErreur:

80      woups "frmSortieMateriel", "ViderChamps", Err, Erl
End Sub

Private Sub RemplirListViewRecherche()

5       On Error GoTo AfficherErreur

10      Dim rstInv As ADODB.Recordset
15      Dim itmInv As ListItem
20      Dim sWhere As String

25      Screen.MousePointer = vbHourglass

30      Call lvwRecherche.ListItems.Clear

35      Select Case cmbRecherche.ListIndex
          Case I_CMB_NO_ITEM:       sWhere = "Instr(1,NoItem,'" & txtRecherche.Text & "') > 0"
40        Case I_CMB_DESCRIPTION:   sWhere = "Instr(1,Description,'" & txtRecherche.Text & "') > 0"
45        Case I_CMB_MANUFACTURIER: sWhere = "Instr(1,Manufacturier,'" & txtRecherche.Text & "') > 0"
50      End Select

55      Set rstInv = New ADODB.Recordset

60      If m_eCatalogue = ELECTRIQUE Then
65        Call rstInv.Open("SELECT * FROM GRB_InventaireElec WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
70      Else
75        Call rstInv.Open("SELECT * FROM GRB_InventaireMec WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
80      End If

85      Do While Not rstInv.EOF
90        Set itmInv = lvwRecherche.ListItems.Add

95        itmInv.Text = rstInv.Fields("NoItem")
100       itmInv.SubItems(I_LVW_RECHERCHE_DESCRIPTION) = rstInv.Fields("Description")
105       itmInv.SubItems(I_LVW_RECHERCHE_MANUFACTURIER) = rstInv.Fields("Manufacturier")

110       Call rstInv.MoveNext
115     Loop

120     Call rstInv.Close
125     Set rstInv = Nothing

130     Screen.MousePointer = vbDefault

135     If lvwRecherche.ListItems.count = 0 Then
140       Call MsgBox("Aucun enregistrement trouvé!", vbOKOnly, "Erreur")
145     End If

150     Exit Sub

AfficherErreur:

155     woups "frmSortieMateriel", "RemplirListViewRecherche", Err, Erl
End Sub

Private Sub cmdRechercher_Click()

5       On Error GoTo AfficherErreur

10      If txtRecherche.Text <> "" Then
15        Call RemplirListViewRecherche
20      Else
25        Call MsgBox("Rien à rechercher!", vbOKOnly, "Erreur")
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmSortieMateriel", "cmdRechercher_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      Call RemplirComboEmployes

15      Call ViderChamps

20      Exit Sub

AfficherErreur:

25      woups "frmSortieMateriel", "Form_Load", Err, Erl
End Sub

Private Sub RemplirComboEmployes()

5       On Error GoTo AfficherErreur

10      Dim rstEmploye As ADODB.Recordset

15      Set rstEmploye = New ADODB.Recordset

20      Call rstEmploye.Open("SELECT * FROM GRB_Employés WHERE Actif = True", g_connData, adOpenDynamic, adLockOptimistic)

25      Do While Not rstEmploye.EOF
30        Call cmbemployé.AddItem(rstEmploye.Fields("Employe"))

35        cmbemployé.ItemData(cmbemployé.newIndex) = rstEmploye.Fields("NoEmploye")

40        Call rstEmploye.MoveNext
45      Loop

50      Call rstEmploye.Close
55      Set rstEmploye = Nothing

60      Exit Sub

AfficherErreur:

65      woups "frmSortieMateriel", "RemplirComboEmployes", Err, Erl
End Sub

Private Sub lvwRecherche_DblClick()

5       On Error GoTo AfficherErreur

10      If lvwRecherche.ListItems.count > 0 Then
15        txtNoItem.Text = lvwRecherche.SelectedItem.Text
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmSortieMateriel", "lvwRecherche_DblClick", Err, Erl
End Sub

Private Sub mskNoProjet_Change()

5       On Error GoTo AfficherErreur

10      If fraAjout.Visible = True Then
15        If InStr(1, mskNoProjet.Text, "_") = 0 Then
20          Call ProjetExiste
25        End If
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmSortieMateriel", "mskNoProjet_Change", Err, Erl
End Sub

Private Function ProjetExiste() As Boolean

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset
15      Dim iCompteur   As Integer
    
20      If Right$(mskNoProjet.Text, 2) >= 51 And Right$(mskNoProjet.Text, 2) <= 98 Then
25        Set rstProjSoum = New ADODB.Recordset

30        Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & mskNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
  
35        If Not rstProjSoum.EOF Then
40          If rstProjSoum.Fields("Ouvert") = False Then
45            Call MsgBox("Ce projet n'est pas ouvert!", vbOKOnly, "Erreur")

50            ProjetExiste = False
55          Else
60            ProjetExiste = True
65          End If
70        Else
75          Call MsgBox("Projet inexistant!", vbOKOnly, "Erreur")

80          ProjetExiste = False
85        End If
  
90        Call rstProjSoum.Close
95        Set rstProjSoum = Nothing
100     Else
105       Call MsgBox("Impossible de faire une sortie de matériel sur ce numéro!", vbOKOnly, "Erreur")

110       ProjetExiste = False
115     End If

120     Exit Function

AfficherErreur:

125     woups "frmSortieMateriel", "AfficherClient", Err, Erl
End Function

Private Sub chkMecanique_Click()

5       On Error GoTo AfficherErreur

10      Dim sTampon As String

15      sTampon = mskNoProjet.Text
  
        'dépendant si coché mécanique affiche le mask
20      If chkMecanique.Value = vbChecked Then
25        mskNoProjet.mask = "\M#####-##"
          'ajoute le M
30        If Len(sTampon) = 8 Then
35          mskNoProjet.Text = "M" + sTampon
40        End If
45      Else
          'enleve le m
50        mskNoProjet.mask = "#####-##"
55        mskNoProjet.Text = Right$(sTampon, 9)
60      End If
  
65      If fraAjout.Visible = True Then
70        Call mskNoProjet.SetFocus
75      End If

80      Exit Sub

AfficherErreur:

85      woups "frmSortieMateriel", "chkMecanique_Click", Err, Erl
End Sub

Public Sub Afficher(ByVal eCatalogue As enumCatalogue)
  
5       On Error GoTo AfficherErreur

10      m_eCatalogue = eCatalogue

15      Call Unload(frmChoixSortieMateriel)

20      Call Me.Show

25      Exit Sub

AfficherErreur:

30      woups "frmSortieMateriel", "Afficher", Err, Erl
End Sub

Private Sub AjouterDansProjet(ByVal sNoProjet As String, ByVal eExtra As enumExtra, ByVal sProvenance As String)
  
5       On Error GoTo AfficherErreur

10      Dim rstProjet  As ADODB.Recordset
15      Dim rstPiece   As ADODB.Recordset
20      Dim rstSection As ADODB.Recordset
25      Dim rstInv     As ADODB.Recordset
30      Dim iCompteur  As Integer
35      Dim sSection   As String
40      Dim bSkip      As Boolean
45      Dim sIDSection As String
50      Dim sOrdre     As String
55      Dim sProfit    As String
        
60      Set rstProjet = New ADODB.Recordset
65      Set rstSection = New ADODB.Recordset
        
70      If m_eCatalogue = ELECTRIQUE Then
75        Call rstSection.Open("SELECT * FROM GRB_SoumProjSectionElec WHERE NomSectionFR = 'Externe'", g_connData, adOpenDynamic, adLockOptimistic)

80        Call rstProjet.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

85        sProfit = rstProjet.Fields("Profit")

90        Call rstProjet.Close
95      Else
100       Call rstSection.Open("SELECT * FROM GRB_SoumProjSectionMec WHERE NomSectionFR = 'Externe'", g_connData, adOpenDynamic, adLockOptimistic)

105       Call rstProjet.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockBatchOptimistic)

110       sProfit = rstProjet.Fields("Profit")

115       Call rstProjet.Close
120     End If

125     sIDSection = rstSection.Fields("IDSection")
130     sOrdre = rstSection.Fields("Ordre")

135     Call rstSection.Close
140     Set rstSection = Nothing

        'Ouverture du recordset sur le projet original
145     Set rstPiece = New ADODB.Recordset
          
150     Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjet & "' AND IDSection = " & sIDSection & " AND SousSection = 'PAS DE SOUS-SECTION' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

155     If Not rstPiece.EOF Then
160       Call rstPiece.MoveLast

165       iCompteur = rstPiece.Fields("NuméroLigne") + 1
170     Else
175       iCompteur = 1
180     End If

185     Set rstInv = New ADODB.Recordset

190     If m_eCatalogue = ELECTRIQUE Then
195       Call rstInv.Open("SELECT * FROM GRB_InventaireElec WHERE NoItem = '" & Replace(txtNoItem.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
200     Else
205       Call rstInv.Open("SELECT * FROM GRB_InventaireMec WHERE NoItem = '" & Replace(txtNoItem.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
210     End If

215     Call rstPiece.AddNew

220     rstPiece.Fields("IDProjet") = sNoProjet
       
225     If m_eCatalogue = ELECTRIQUE Then
230       rstPiece.Fields("Type") = "E"
235     Else
240       rstPiece.Fields("Type") = "M"
245     End If

250     rstPiece.Fields("Visible") = True

255     rstPiece.Fields("Facturation") = ""
              
265     rstPiece.Fields("IDSection") = sIDSection
270     rstPiece.Fields("NumItem") = rstInv.Fields("NoItem")
275     rstPiece.Fields("Qté") = txtQte.Text
280     rstPiece.Fields("Desc_FR") = rstInv.Fields("Description")
285     rstPiece.Fields("Desc_EN") = ""
290     rstPiece.Fields("Manufact") = rstInv.Fields("Manufacturier")
295     rstPiece.Fields("Prix_list") = Conversion(rstInv.Fields("Prix liste"), MODE_PAS_FORMAT, 4)
300     rstPiece.Fields("Escompte") = Conversion(rstInv.Fields("Escompte"), MODE_PAS_FORMAT)
305     rstPiece.Fields("Prix_net") = Conversion(rstInv.Fields("Prix net"), MODE_PAS_FORMAT, 4)
310     rstPiece.Fields("OrdreSection") = sOrdre
315     rstPiece.Fields("NuméroLigne") = iCompteur
      
320     rstPiece.Fields("IDFRS") = 717
       
325     rstPiece.Fields("Prix_Total") = Conversion(rstInv.Fields("Prix net") * txtQte.Text * CDbl(sProfit), MODE_PAS_FORMAT)
330     rstPiece.Fields("Profit_argent") = Conversion(rstPiece.Fields("Prix_Total") - (rstInv.Fields("Prix net") * txtQte.Text), MODE_PAS_FORMAT)
335     rstPiece.Fields("SousSection") = "PAS DE SOUS-SECTION"
            
340     rstPiece.Fields("PrixOrigine") = Replace(Round(CDbl(Replace(rstInv.Fields("Prix liste"), ".", ",")), 2), ".", ",")

345     Select Case eExtra
          Case EXTRA_CHARGEABLE:
350          rstPiece.Fields("PieceExtraChargeable") = True
355          rstPiece.Fields("PieceExtraNonChargeable") = False

360       Case EXTRA_NON_CHARGEABLE:
365          rstPiece.Fields("PieceExtraChargeable") = False
370          rstPiece.Fields("PieceExtraNonChargeable") = True

375       Case AUCUN_EXTRA:
380          rstPiece.Fields("PieceExtraChargeable") = False
385          rstPiece.Fields("PieceExtraNonChargeable") = False
390     End Select

395     rstPiece.Fields("Provenance") = sProvenance

400     Call rstPiece.Update

405     Call rstPiece.Close

410     Call rstInv.Close
415     Set rstInv = Nothing

420     rstPiece.CursorLocation = adUseServer

425     Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjet & "' AND NuméroLigne >= " & iCompteur & " ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

        'Tant qu'il y a des enregistrements dans le recordset
430     Do While Not rstPiece.EOF
435       If bSkip = False Then
440         bSkip = True
445       Else
450         rstPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne") + 1

455         Call rstPiece.Update
460       End If

465       Call rstPiece.MoveNext
470     Loop

475     Call rstPiece.Close
480     Set rstPiece = Nothing

485     If m_eCatalogue = ELECTRIQUE Then
490       Call CalculerTempsMecRecordset(sNoProjet)

495       Call CalculerTotalRecordsetElec(sNoProjet)
500     Else
505       Call CalculerTotalRecordsetMec(sNoProjet)
510     End If

515     Exit Sub

AfficherErreur:

520     woups "frmSortieMateriel", "AjouterDansProjet", Err, Erl
End Sub

Private Sub CalculerTempsMecRecordset(ByVal sNoProjet As String)

5       On Error GoTo AfficherErreur

10      Dim rstProjet   As ADODB.Recordset
15      Dim rstPiece    As ADODB.Recordset
20      Dim dblTempsMec As Double

        'Ouverture des tables
25      Set rstProjet = New ADODB.Recordset
30      Set rstPiece = New ADODB.Recordset

35      Call rstProjet.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

40      Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet ='" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

        'Pour chaque enregistrement du recordset
45      Do While Not rstPiece.EOF
          'Si le temps total n'est pas vide
50        If Trim(rstPiece.Fields("Temps_total")) <> vbNullString Then
            'On additionne le temps
55          dblTempsMec = dblTempsMec + CDbl(Replace(Trim(rstPiece.Fields("Temps_total")), ".", ","))
60        End If

65        Call rstPiece.MoveNext
70      Loop
                
75      rstProjet.Fields("temp_mec") = Replace(dblTempsMec / 10, ".", ",")

80      Call rstProjet.Update

85      Call rstPiece.Close
90      Set rstPiece = Nothing

95      Call rstProjet.Close
100     Set rstProjet = Nothing

105     Exit Sub

AfficherErreur:

110     woups "frmSortieMateriel", "CalculerTempsMecRecordset", Err, Erl
End Sub

Private Sub CalculerTotalRecordsetElec(ByVal sNoProjet As String)

5       On Error GoTo AfficherErreur

        'Méthode pour calculer le prix
10      Dim dblManuel             As Double
15      Dim dblCopies             As Double
20      Dim dblTempsDessin        As Double
25      Dim dblTempsProg          As Double
30      Dim dblTempsMec           As Double
35      Dim dblTempsElec          As Double
40      Dim dblTempsTest          As Double
45      Dim dblTempsVision        As Double
50      Dim dblPrixPieces         As Double
55      Dim dblPrixTotal          As Double
60      Dim dblCommission         As Double
65      Dim dblTotalTemps         As Double
70      Dim dblProfit             As Double
75      Dim dblTotalManuel        As Double
80      Dim dblTotalPieceImprevue As Double
85      Dim dblGrandTotal         As Double
90      Dim rstProjet             As ADODB.Recordset
95      Dim rstPiece              As ADODB.Recordset
100     Dim rstConfig             As ADODB.Recordset
105     Dim sCommission           As String
110     Dim sCopieManuel          As String
115     Dim sImprevue             As String

120     Set rstProjet = New ADODB.Recordset
125     Set rstPiece = New ADODB.Recordset
130     Set rstConfig = New ADODB.Recordset

135     Call rstConfig.Open("SELECT * FROM GRB_Config", g_connData, adOpenDynamic, adLockOptimistic)

140     sCommission = rstConfig.Fields("Commission")
145     sImprevue = rstConfig.Fields("Imprévus")
150     sCopieManuel = rstConfig.Fields("PrixPagesManuel")

155     Call rstConfig.Close
160     Set rstConfig = Nothing

165     Call rstProjet.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

170     If Not rstProjet.EOF Then
          'Validation du nombre de pages
175       If IsNumeric(rstProjet.Fields("Manuel")) Then
180         dblManuel = rstProjet.Fields("Manuel")
185       Else
190         dblManuel = 0
195       End If

          'Validation du nombre de copies
200       If IsNumeric(rstProjet.Fields("copies")) Then
205         dblCopies = rstProjet.Fields("copies")
210       Else
215         dblCopies = 0
220       End If

          'Validation des heures de dessin
225       If IsNumeric(rstProjet.Fields("temp_dessin")) Then
230         dblTempsDessin = rstProjet.Fields("temp_dessin")
235       Else
240         dblTempsDessin = 0
245       End If

          'Validation des heures de prog
250       If IsNumeric(rstProjet.Fields("temp_prog")) Then
255         dblTempsProg = rstProjet.Fields("temp_prog")
260       Else
265         dblTempsProg = 0
270       End If

          'Validation des heures de mec
275       If rstProjet.Fields("SansTemps") = True Then
280         dblTempsMec = 0
285       Else
290         dblTempsMec = rstProjet.Fields("temp_mec")
295       End If

          'Validation des heures d'élec
300       If IsNumeric(rstProjet.Fields("temp_elec")) Then
305         dblTempsElec = rstProjet.Fields("temp_elec")
310       Else
315         dblTempsElec = 0
320       End If

          'Validation des heures de test
325       If IsNumeric(rstProjet.Fields("temp_test")) Then
330         dblTempsTest = rstProjet.Fields("temp_test")
335       Else
340         dblTempsTest = 0
345       End If

          'Validation des heures de vision
350       If IsNumeric(rstProjet.Fields("temp_vision")) Then
355         dblTempsVision = rstProjet.Fields("temp_vision")
360       Else
365         dblTempsVision = 0
370       End If

375       Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjet & "' AND Type = 'E'", g_connData, adOpenDynamic, adLockOptimistic)

          'Pour chaque élément du recordset
380       Do While Not rstPiece.EOF
385         If Trim(rstPiece.Fields("Prix_total")) <> vbNullString Then
390           'On additionne le prix total
395           dblPrixPieces = dblPrixPieces + rstPiece.Fields("Prix_total")

              'On additionne le profit
400           dblProfit = dblProfit + rstPiece.Fields("Profit_Argent")
405         End If

410         Call rstPiece.MoveNext
415       Loop

420       Call rstPiece.Close
425       Set rstPiece = Nothing

          'On additionne les (temps * taux)
430       dblTotalTemps = (dblTempsDessin * CDbl(rstProjet.Fields("taux_dessin"))) + (dblTempsProg * CDbl(rstProjet.Fields("taux_prog"))) + (dblTempsMec * CDbl(rstProjet.Fields("taux_mec"))) + (dblTempsElec * CDbl(rstProjet.Fields("taux_elec"))) + (dblTempsTest * CDbl(rstProjet.Fields("taux_test"))) + (dblTempsVision * CDbl(rstProjet.Fields("taux_vision")))

435       dblTotalManuel = dblManuel * dblCopies * CDbl(sCopieManuel)

440       dblTotalPieceImprevue = dblPrixPieces * (1 + CDbl(sImprevue))

445       dblPrixTotal = dblTotalTemps + dblTotalManuel + dblTotalPieceImprevue

          'Le calcul de la commission est sur les manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps
450       dblCommission = dblPrixTotal * CDbl(sCommission)

          'Le prix total est le calcul des manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps + total de la commission
455       dblGrandTotal = dblPrixTotal + dblCommission

          'Format monétaires avec 2 chiffres après la virgule
460       rstProjet.Fields("total_commission") = dblCommission
465       rstProjet.Fields("Total_manuel") = dblTotalManuel
470       rstProjet.Fields("Total_temps") = dblTotalTemps
475       rstProjet.Fields("total_imprevue") = dblTotalPieceImprevue - dblPrixPieces
480       rstProjet.Fields("total_piece") = dblPrixPieces
485       rstProjet.Fields("Total_Commission") = Conversion(CStr(Round(dblCommission, 2)), MODE_ARGENT)
490       rstProjet.Fields("PrixTotal") = Conversion(CStr(Round(dblGrandTotal, 2)), MODE_ARGENT)
495       rstProjet.Fields("Total_Profit") = Conversion(CStr(Round(dblProfit, 2)), MODE_ARGENT)

500       Call rstProjet.Update
505     End If

510     Call rstProjet.Close
515     Set rstProjet = Nothing

520     Exit Sub

AfficherErreur:

525     woups "frmSortieMateriel", "CalculerTotalRecordset", Err, Erl
End Sub

Private Sub CalculerTotalRecordsetMec(ByVal sNoProjet As String)

5       On Error GoTo AfficherErreur

        'Méthode pour calculer le prix
10      Dim dblPrixPieces        As Double
15      Dim dblPrixTotal         As Double
20      Dim dblCommission        As Double
25      Dim dblTotalTemps        As Double
30      Dim dblProfit            As Double
35      Dim dblTotalManuel       As Double
40      Dim dblTotalImprevue     As Double
45      Dim dblGrandTotal        As Double
50      Dim dblTotalDessin       As Double
55      Dim dblTotalCoupe        As Double
60      Dim dblTotalMachinage    As Double
65      Dim dblTotalSoudure      As Double
70      Dim dblTotalAssemblage   As Double
75      Dim dblTotalPeinture     As Double
80      Dim dblTotalTest         As Double
85      Dim dblTotalInstallation As Double
90      Dim dblTotalFormation    As Double
95      Dim dblTotalGestion      As Double
100     Dim dblHebergement       As Double
105     Dim dblRepas             As Double
110     Dim dblTransport         As Double
115     Dim dblUniteMobile       As Double
120     Dim dblPrixEmballage     As Double
125     Dim dblTotalResteTemps   As Double
130     Dim iNbrePersonne        As Integer
135     Dim rstProjet            As ADODB.Recordset
140     Dim rstPiece             As ADODB.Recordset
145     Dim rstConfig            As ADODB.Recordset
150     Dim sCommission          As String
155     Dim sImprevue            As String

160     Set rstConfig = New ADODB.Recordset

165     Call rstConfig.Open("SELECT * FROM GRB_Config", g_connData, adOpenDynamic, adLockOptimistic)

170     sCommission = rstConfig.Fields("Commission")
175     sImprevue = rstConfig.Fields("Imprévus")

180     Call rstConfig.Close
185     Set rstConfig = Nothing

190     Set rstProjet = New ADODB.Recordset
195     Set rstPiece = New ADODB.Recordset

200     Call rstProjet.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

205     Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjet & "' AND Type = 'M'", g_connData, adOpenDynamic, adLockOptimistic)

        'Pour chaque élément du recordset
210     Do While Not rstPiece.EOF
215       If Trim(rstPiece.Fields("Prix_total")) <> vbNullString Then
            'On additionne le prix total
220         dblPrixPieces = dblPrixPieces + rstPiece.Fields("Prix_total") - rstPiece.Fields("Profit_Argent")
          
            'On additionne le profit
225         dblProfit = dblProfit + rstPiece.Fields("Profit_Argent")
230       End If

235       Call rstPiece.MoveNext
240     Loop
    
        'Total des temps
245     dblTotalMachinage = CDbl(rstProjet.Fields("TempsMachinage")) * CDbl(rstProjet.Fields("TauxMachinage"))
250     dblTotalCoupe = CDbl(rstProjet.Fields("TempsCoupe")) * CDbl(rstProjet.Fields("TauxCoupePréparation"))
255     dblTotalSoudure = CDbl(rstProjet.Fields("TempsSoudure")) * CDbl(rstProjet.Fields("TauxAssemblageSoudure"))
260     dblTotalAssemblage = CDbl(rstProjet.Fields("TempsAssemblage")) * CDbl(rstProjet.Fields("TauxAssemblageSystèmes"))
265     dblTotalPeinture = CDbl(rstProjet.Fields("TempsPeinture")) * CDbl(rstProjet.Fields("TauxPeintureFinition"))
270     dblTotalTest = CDbl(rstProjet.Fields("TempsTest")) * CDbl(rstProjet.Fields("TauxTestsFinaux"))
275     dblTotalDessin = CDbl(rstProjet.Fields("TempsDessin")) * CDbl(rstProjet.Fields("TauxConceptionDessins"))
280     dblTotalFormation = CDbl(rstProjet.Fields("TempsFormation")) * CDbl(rstProjet.Fields("TauxFormation"))
285     dblTotalInstallation = CDbl(rstProjet.Fields("TempsInstallation")) * CDbl(rstProjet.Fields("TauxInstallation"))
290     dblTotalGestion = CDbl(rstProjet.Fields("TempsGestion")) * CDbl(rstProjet.Fields("TauxGestion"))
          
295     dblTotalTemps = dblTotalMachinage + dblTotalCoupe + dblTotalSoudure + _
                        dblTotalAssemblage + dblTotalPeinture + dblTotalTest + _
                        dblTotalDessin + dblTotalFormation + dblTotalInstallation
      
300     If Not IsNull(rstProjet.Fields("NbrePersonne")) Then
305       If Trim(rstProjet.Fields("NbrePersonne")) <> "" Then
310         iNbrePersonne = Int(rstProjet.Fields("NbrePersonne"))
315       Else
320         iNbrePersonne = 0
325       End If
330     Else
335       iNbrePersonne = 0
340     End If
           
345     Do While iNbrePersonne > 0
350       If iNbrePersonne >= 2 Then
355         dblHebergement = dblHebergement + CDbl(rstProjet.Fields("TempsHebergement")) * CDbl(rstProjet.Fields("TauxHebergement2"))
              
360         iNbrePersonne = iNbrePersonne - 2
365       Else
370         dblHebergement = dblHebergement + CDbl(rstProjet.Fields("TempsHebergement")) * CDbl(rstProjet.Fields("TauxHebergement1"))
             
375         iNbrePersonne = iNbrePersonne - 1
380       End If
385     Loop
      
390     If Not IsNull(rstProjet.Fields("NbrePersonne")) Then
395       If Trim(rstProjet.Fields("NbrePersonne")) <> "" Then
400         dblRepas = CDbl(rstProjet.Fields("TempsRepas")) * CDbl(rstProjet.Fields("TauxRepas")) * CDbl(rstProjet.Fields("NbrePersonne"))
405       Else
410         dblRepas = 0
415       End If
420     Else
425       dblRepas = 0
430     End If

435     dblTransport = CDbl(rstProjet.Fields("TempsTransport")) * CDbl(rstProjet.Fields("TauxTransport"))
440     dblUniteMobile = CDbl(rstProjet.Fields("TempsUniteMobile")) * CDbl(rstProjet.Fields("TauxUniteMobile"))

        'Correction d'un bug de Type Incompatible
445     If IsNumeric(rstProjet.Fields("PrixEmballage")) Then
450       dblPrixEmballage = CDbl(rstProjet.Fields("PrixEmballage"))
455     Else
460       dblPrixEmballage = 0
465     End If
      
470     dblTotalResteTemps = dblHebergement + dblRepas + dblTransport + dblUniteMobile + dblPrixEmballage
                                                              
475     If IsNumeric(rstProjet.Fields("total_manuel")) Then
480       dblTotalManuel = CDbl(rstProjet.Fields("total_manuel"))
485     Else
490       dblTotalManuel = 0
495     End If
                        
500     dblTotalImprevue = (dblPrixPieces + dblProfit) * CDbl(sImprevue)
    
505     dblPrixTotal = dblPrixPieces + dblProfit + dblTotalTemps + dblTotalImprevue + dblTotalManuel + dblTotalResteTemps
                        
        'Le calcul de la commission est sur les manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps
510     dblCommission = dblPrixTotal * CDbl(sCommission)
        
        'Le prix total est le calcul des manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps + total de la commission
515     dblGrandTotal = dblPrixTotal + dblCommission
                
        'Format monétaires avec 2 chiffres après la virgule
520     rstProjet.Fields("Total_Piece") = Conversion(CStr(Round(dblPrixPieces, 2)), MODE_ARGENT)
525     rstProjet.Fields("Total_Temps") = Conversion(CStr(Round(dblTotalTemps, 2)), MODE_ARGENT)
530     rstProjet.Fields("PrixTotal") = Conversion(CStr(Round(dblGrandTotal, 2)), MODE_ARGENT)
535     rstProjet.Fields("total_Imprevue") = Conversion(CStr(Round(dblTotalImprevue, 2)), MODE_ARGENT)
540     rstProjet.Fields("total_commission") = Conversion(CStr(Round(dblCommission, 2)), MODE_ARGENT)
545     rstProjet.Fields("total_profit") = Conversion(CStr(Round(dblProfit, 2)), MODE_ARGENT)

550     Call rstProjet.Update

555     Call rstPiece.Close
560     Set rstPiece = Nothing

565     Call rstProjet.Close
570     Set rstProjet = Nothing

575     Exit Sub

AfficherErreur:

580     woups "frmSortieMateriel", "CalculerTotalRecordset", Err, Erl
End Sub
