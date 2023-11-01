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
   Picture         =   "frmRetourMateriel.frx":0000
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
55              If VerifierSortie = True Then
60                Set rstProjet = New ADODB.Recordset

65                If m_eCatalogue = ELECTRIQUE Then
70                  Call rstProjet.Open("SELECT Modification, Par FROM GRB_ProjetElec WHERE IDProjet = '" & mskNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
75                Else
80                  Call rstProjet.Open("SELECT Modification, Par FROM GRB_ProjetMec WHERE IDProjet = '" & mskNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
85                End If

90                If rstProjet.Fields("Modification") = False Then
95                  Set rstInv = New ADODB.Recordset

100                 If m_eCatalogue = ELECTRIQUE Then
105                   Call rstInv.Open("SELECT * FROM GRB_InventaireElec WHERE NoItem = '" & Replace(txtNoItem.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
110                 Else
115                   Call rstInv.Open("SELECT * FROM GRB_InventaireMec WHERE NoItem = '" & Replace(txtNoItem.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
120                 End If

125                 If Not rstInv.EOF Then
130                   Set rstSortie = New ADODB.Recordset

135                   Call rstSortie.Open("SELECT * FROM GRB_SortieMatériel", g_connData, adOpenDynamic, adLockOptimistic)

140                   Call rstSortie.AddNew

145                   rstSortie.Fields("Qté") = "-" & Abs(txtQte.Text)
150                   rstSortie.Fields("Nom") = cmbemployé.Text
155                   rstSortie.Fields("NoProjet") = mskNoProjet.Text
160                   rstSortie.Fields("NoItem") = txtNoItem.Text
165                   rstSortie.Fields("Date") = ConvertDate(Date)

170                   If m_eCatalogue = ELECTRIQUE Then
175                     rstSortie.Fields("Type") = "E"
180                   Else
185                     rstSortie.Fields("Type") = "M"
190                   End If

195                   Call rstSortie.Update

200                   Call rstSortie.Close
205                   Set rstSortie = Nothing

210                   rstInv.Fields("QuantitéStock") = CDbl(rstInv.Fields("QuantitéStock")) + CDbl(Abs(txtQte.Text))

215                   Call rstInv.Update

220                   Set rstHistInv = New ADODB.Recordset

225                   If m_eCatalogue = ELECTRIQUE Then
230                     Call rstHistInv.Open("SELECT * FROM GRB_InventaireElecModif", g_connData, adOpenDynamic, adLockOptimistic)
235                   Else
240                     Call rstHistInv.Open("SELECT * FROM GRB_InventaireMecModif", g_connData, adOpenDynamic, adLockOptimistic)
245                   End If

250                   Call rstHistInv.AddNew

255                   rstHistInv.Fields("Date") = ConvertDate(Date)
260                   rstHistInv.Fields("IDProjet") = mskNoProjet.Text
265                   rstHistInv.Fields("NoItem") = txtNoItem.Text
270                   rstHistInv.Fields("Quantité") = Abs(txtQte.Text)

275                   Set rstInitiale = New ADODB.Recordset

280                   Call rstInitiale.Open("SELECT * FROM GRB_Employés WHERE NoEmploye = " & cmbemployé.ItemData(cmbemployé.ListIndex), g_connData, adOpenDynamic, adLockOptimistic)

285                   rstHistInv.Fields("User") = rstInitiale.Fields("Initiale")

290                   Call rstInitiale.Close
295                   Set rstInitiale = Nothing

300                   Call rstHistInv.Update

305                   Call rstHistInv.Close
310                   Set rstHistInv = Nothing

315                   Call AjouterDansProjet(mskNoProjet.Text, AUCUN_EXTRA, "")

320                   Call rstProjet.Close

325                   If m_eCatalogue = ELECTRIQUE Then
330                     If Right$(mskNoProjet.Text, 2) >= 61 And Right$(mskNoProjet.Text, 2) <= 80 Then
335                       Call rstProjet.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & mskNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

340                       Call AjouterDansProjet(Left$(mskNoProjet.Text, Len(mskNoProjet.Text) - 2) & rstProjet.Fields("LiaisonChargeable"), EXTRA_CHARGEABLE, Right$(mskNoProjet.Text, 2))

345                       Call rstProjet.Close
350                     Else
355                       If Right$(mskNoProjet.Text, 2) >= 81 And Right$(mskNoProjet.Text, 2) <= 98 Then
360                         Call AjouterDansProjet(Left$(mskNoProjet.Text, Len(mskNoProjet.Text) - 2) & Right$("0" & Right$(mskNoProjet.Text, 2) - 80, 2), EXTRA_NON_CHARGEABLE, "")
365                       End If
370                     End If
375                   Else
380                     If Right$(mskNoProjet.Text, 2) >= 61 And Right$(mskNoProjet.Text, 2) <= 80 Then
385                       Call rstProjet.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & mskNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

390                       Call AjouterDansProjet(Left$(mskNoProjet.Text, Len(mskNoProjet.Text) - 2) & rstProjet.Fields("LiaisonChargeable"), EXTRA_CHARGEABLE, Right$(mskNoProjet.Text, 2))

395                       Call rstProjet.Close
400                     Else
405                       If Right$(mskNoProjet.Text, 2) >= 81 And Right$(mskNoProjet.Text, 2) <= 98 Then
410                         Call AjouterDansProjet(Left$(mskNoProjet.Text, Len(mskNoProjet.Text) - 2) & Right$("0" & Right$(mskNoProjet.Text, 2) - 80, 2), EXTRA_NON_CHARGEABLE, "")
415                       End If
420                     End If
425                   End If

430                   Call MsgBox("Le retour de matériel a été enregistrée!", vbOKOnly, "Erreur")

435                   Call ViderChamps
440                 Else
445                   Call MsgBox("Cette pièce n'existe pas dans l'inventaire!", vbOKOnly, "Erreur")
450                 End If

455                 Call rstInv.Close
460                 Set rstInv = Nothing
465               Else
470                 Call MsgBox("Ce projet est en modification par " & rstProjet.Fields("Par") & "!", vbOKOnly, "Erreur")

475                 Call rstProjet.Close
480               End If

485               Set rstProjet = Nothing
490             Else
495               Call MsgBox("Pas assez de pièces ont été sortie pour en retourner " & txtQte.Text & "!", vbOKOnly, "Erreur")
500             End If
505           Else
510             Call MsgBox("Projet inexistant!", vbOKOnly, "Erreur")
515           End If
520         Else
525           Call MsgBox("Le numéro de projet est obligatoire!", vbOKOnly, "Erreur")
530         End If
535       Else
540         Call MsgBox("Quantité non numérique!", vbOKOnly, "Erreur")
545       End If
550     Else
555       Call MsgBox("Le numéro d'item est obligatoire!", vbOKOnly, "Erreur")
560     End If

565     Exit Sub

AfficherErreur:

570     woups "frmRetourMateriel", "cmdEnregistrer_Click", Err, Erl
End Sub

Private Sub Cmdfermer_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmRetourMateriel", "cmdFermer_Click", Err, Erl
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

80      woups "frmRetourMateriel", "ViderChamps", Err, Erl
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

155     woups "frmRetourMateriel", "RemplirListViewRecherche", Err, Erl
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

40      woups "frmRetourMateriel", "cmdRechercher_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      Call RemplirComboEmployes

15      Call ViderChamps

20      Exit Sub

AfficherErreur:

25      woups "frmRetourMateriel", "Form_Load", Err, Erl
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

65      woups "frmRetourMateriel", "RemplirComboEmployes", Err, Erl
End Sub

Private Sub lvwRecherche_DblClick()

5       On Error GoTo AfficherErreur

10      If lvwRecherche.ListItems.count > 0 Then
15        txtNoItem.Text = lvwRecherche.SelectedItem.Text
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmRetourMateriel", "lvwRecherche_DblClick", Err, Erl
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

40      woups "frmRetourMateriel", "mskNoProjet_Change", Err, Erl
End Sub

Private Function ProjetExiste() As Boolean

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset
    
15      If Right$(mskNoProjet.Text, 2) >= 51 And Right$(mskNoProjet.Text, 2) <= 98 Then
20        Set rstProjSoum = New ADODB.Recordset

25        Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & mskNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
  
30        If Not rstProjSoum.EOF Then
35          If rstProjSoum.Fields("Ouvert") = False Then
40            Call MsgBox("Ce projet n'est pas ouvert!", vbOKOnly, "Erreur")

45            ProjetExiste = False
50          Else
55            ProjetExiste = True
60          End If
65        Else
70          Call MsgBox("Projet inexistant!", vbOKOnly, "Erreur")

75          ProjetExiste = False
80        End If

85        Call rstProjSoum.Close
90        Set rstProjSoum = Nothing
95      Else
100       Call MsgBox("Impossible de faire une sortie de matériel sur ce numéro!", vbOKOnly, "Erreur")

105       ProjetExiste = False
110     End If
  
115     Exit Function

AfficherErreur:

120     woups "frmRetourMateriel", "ProjetExiste", Err, Erl
End Function

Private Function VerifierSortie() As Boolean

5       On Error GoTo AfficherErreur

10      Dim rstProjet    As ADODB.Recordset
15      Dim rstSection   As ADODB.Recordset
20      Dim sIDSection   As String
25      Dim dblQteProjet As Double
    
30      Set rstSection = New ADODB.Recordset

35      If m_eCatalogue = ELECTRIQUE Then
40        Call rstSection.Open("SELECT * FROM GRB_SoumProjSectionElec WHERE NomSectionFR = 'Externe'", g_connData, adOpenDynamic, adLockOptimistic)
45      Else
50        Call rstSection.Open("SELECT * FROM GRB_SoumProjSectionMec WHERE NomSectionFR = 'Externe'", g_connData, adOpenDynamic, adLockOptimistic)
55      End If

60      sIDSection = rstSection.Fields("IDSection")

65      Call rstSection.Close
70      Set rstSection = Nothing

75      Set rstProjet = New ADODB.Recordset
  
80      Call rstProjet.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & mskNoProjet.Text & "' AND IDSection = " & sIDSection & " AND SousSection = 'PAS DE SOUS-SECTION' AND NumItem = '" & Replace(txtNoItem.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
  
85      If Not rstProjet.EOF Then
90        Do While Not rstProjet.EOF
95          dblQteProjet = dblQteProjet + rstProjet.Fields("Qté")

100         Call rstProjet.MoveNext
105       Loop

110       If dblQteProjet >= Abs(txtQte.Text) Then
115         VerifierSortie = True
120       Else
125         Call MsgBox("Il n'y a pas assez de " & txtNoItem.Text & " dans le projet " & mskNoProjet.Text & " pour en enlever " & Abs(txtQte.Text), vbOKOnly, "Erreur")

130         VerifierSortie = False
135       End If
140     Else
145       Call MsgBox("La pièce " & txtNoItem.Text & " n'a pas été sortie pour le projet " & mskNoProjet.Text, vbOKOnly, "Erreur")

150       VerifierSortie = False
155     End If

160     Call rstProjet.Close
165     Set rstProjet = Nothing
  
170     Exit Function

AfficherErreur:

175     woups "frmRetourMateriel", "VerifierSortie", Err, Erl
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

85      woups "frmRetourMateriel", "chkMecanique_Click", Err, Erl
End Sub

Public Sub Afficher(ByVal eCatalogue As enumCatalogue)
  
5       On Error GoTo AfficherErreur

10      m_eCatalogue = eCatalogue

15      Call Unload(frmChoixRetourMateriel)

20      Call Me.Show

25      Exit Sub

AfficherErreur:

30      woups "frmRetourMateriel", "Afficher", Err, Erl
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
275     rstPiece.Fields("Qté") = "-" & Abs(txtQte.Text)
280     rstPiece.Fields("Desc_FR") = rstInv.Fields("Description")
285     rstPiece.Fields("Desc_EN") = ""
290     rstPiece.Fields("Manufact") = rstInv.Fields("Manufacturier")
295     rstPiece.Fields("Prix_list") = Conversion(rstInv.Fields("Prix liste"), MODE_PAS_FORMAT, 4)
300     rstPiece.Fields("Escompte") = Conversion(rstInv.Fields("Escompte"), MODE_PAS_FORMAT)
305     rstPiece.Fields("Prix_net") = Conversion(rstInv.Fields("Prix net"), MODE_PAS_FORMAT, 4)
310     rstPiece.Fields("OrdreSection") = sOrdre
315     rstPiece.Fields("NuméroLigne") = iCompteur
      
320     rstPiece.Fields("IDFRS") = 717
       
325     rstPiece.Fields("Prix_Total") = Conversion(rstInv.Fields("Prix net") * rstPiece.Fields("Qté") * CDbl(sProfit), MODE_PAS_FORMAT)
330     rstPiece.Fields("Profit_argent") = Conversion(rstPiece.Fields("Prix_Total") - (rstInv.Fields("Prix net") * rstPiece.Fields("Qté")), MODE_PAS_FORMAT)
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

520     woups "frmRetourMateriel", "AjouterDansProjet", Err, Erl
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

110     woups "frmRetourMateriel", "CalculerTempsMecRecordset", Err, Erl
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
460       rstProjet.Fields("Total_Commission") = Conversion(CStr(Round(dblCommission, 2)), MODE_ARGENT)
465       rstProjet.Fields("PrixTotal") = Conversion(CStr(Round(dblGrandTotal, 2)), MODE_ARGENT)
470       rstProjet.Fields("Total_Profit") = Conversion(CStr(Round(dblProfit, 2)), MODE_ARGENT)

475       Call rstProjet.Update
480     End If

485     Call rstProjet.Close
490     Set rstProjet = Nothing

495     Exit Sub

AfficherErreur:

500     woups "frmRetourMateriel", "CalculerTotalRecordset", Err, Erl
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
50      Dim dblTotalMachinage    As Double
55      Dim dblTotalCoupe        As Double
60      Dim dblTotalSoudure      As Double
65      Dim dblTotalAssemblage   As Double
70      Dim dblTotalPeinture     As Double
75      Dim dblTotalTest         As Double
80      Dim dblTotalDessin       As Double
85      Dim dblTotalFormation    As Double
90      Dim dblTotalInstallation As Double
95      Dim dblHebergement       As Double
100     Dim dblRepas             As Double
105     Dim dblTransport         As Double
110     Dim dblUniteMobile       As Double
115     Dim dblPrixEmballage     As Double
120     Dim dblTotalResteTemps   As Double
125     Dim iNbrePersonne        As Integer
130     Dim rstProjet            As ADODB.Recordset
135     Dim rstPiece             As ADODB.Recordset
140     Dim rstConfig            As ADODB.Recordset
145     Dim sCommission          As String
150     Dim sImprevue            As String

155     Set rstConfig = New ADODB.Recordset

160     Call rstConfig.Open("SELECT * FROM GRB_Config", g_connData, adOpenDynamic, adLockOptimistic)

165     sCommission = rstConfig.Fields("Commission")
170     sImprevue = rstConfig.Fields("Imprévus")

175     Call rstConfig.Close
180     Set rstConfig = Nothing

185     Set rstProjet = New ADODB.Recordset
190     Set rstPiece = New ADODB.Recordset

195     Call rstProjet.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

200     Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjet & "' AND Type = 'M'", g_connData, adOpenDynamic, adLockOptimistic)

        'Pour chaque élément du recordset
205     Do While Not rstPiece.EOF
210       If Trim(rstPiece.Fields("Prix_total")) <> vbNullString Then
            'On additionne le prix total
215         dblPrixPieces = dblPrixPieces + rstPiece.Fields("Prix_total") - rstPiece.Fields("Profit_Argent")
          
            'On additionne le profit
220         dblProfit = dblProfit + rstPiece.Fields("Profit_Argent")
225       End If

230       Call rstPiece.MoveNext
235     Loop
    
        'Total des temps
240     dblTotalMachinage = CDbl(rstProjet.Fields("TempsMachinage")) * CDbl(rstProjet.Fields("TauxMachinage"))
245     dblTotalCoupe = CDbl(rstProjet.Fields("TempsCoupe")) * CDbl(rstProjet.Fields("TauxCoupePréparation"))
250     dblTotalSoudure = CDbl(rstProjet.Fields("TempsSoudure")) * CDbl(rstProjet.Fields("TauxAssemblageSoudure"))
255     dblTotalAssemblage = CDbl(rstProjet.Fields("TempsAssemblage")) * CDbl(rstProjet.Fields("TauxAssemblageSystèmes"))
260     dblTotalPeinture = CDbl(rstProjet.Fields("TempsPeinture")) * CDbl(rstProjet.Fields("TauxPeintureFinition"))
265     dblTotalTest = CDbl(rstProjet.Fields("TempsTest")) * CDbl(rstProjet.Fields("TauxTestsFinaux"))
270     dblTotalDessin = CDbl(rstProjet.Fields("TempsDessin")) * CDbl(rstProjet.Fields("TauxConceptionDessins"))
275     dblTotalFormation = CDbl(rstProjet.Fields("TempsFormation")) * CDbl(rstProjet.Fields("TauxFormation"))
280     dblTotalInstallation = CDbl(rstProjet.Fields("TempsInstallation")) * CDbl(rstProjet.Fields("TauxInstallation"))
          
285     dblTotalTemps = dblTotalMachinage + dblTotalCoupe + dblTotalSoudure + _
                        dblTotalAssemblage + dblTotalPeinture + dblTotalTest + _
                        dblTotalDessin + dblTotalFormation + dblTotalInstallation
      
290     If Not IsNull(rstProjet.Fields("NbrePersonne")) Then
295       If Trim(rstProjet.Fields("NbrePersonne")) <> "" Then
300         iNbrePersonne = Int(rstProjet.Fields("NbrePersonne"))
305       Else
310         iNbrePersonne = 0
315       End If
320     Else
325       iNbrePersonne = 0
330     End If
           
335     Do While iNbrePersonne > 0
340       If iNbrePersonne >= 2 Then
345         dblHebergement = dblHebergement + CDbl(rstProjet.Fields("TempsHebergement")) * CDbl(rstProjet.Fields("TauxHebergement2"))
              
350         iNbrePersonne = iNbrePersonne - 2
355       Else
360         dblHebergement = dblHebergement + CDbl(rstProjet.Fields("TempsHebergement")) * CDbl(rstProjet.Fields("TauxHebergement1"))
             
365         iNbrePersonne = iNbrePersonne - 1
370       End If
375     Loop
      
380     If Not IsNull(rstProjet.Fields("NbrePersonne")) Then
385       If Trim(rstProjet.Fields("NbrePersonne")) <> "" Then
390         dblRepas = CDbl(rstProjet.Fields("TempsRepas")) * CDbl(rstProjet.Fields("TauxRepas")) * CDbl(rstProjet.Fields("NbrePersonne"))
395       Else
400         dblRepas = 0
405       End If
410     Else
415       dblRepas = 0
420     End If

425     dblTransport = CDbl(rstProjet.Fields("TempsTransport")) * CDbl(rstProjet.Fields("TauxTransport"))
430     dblUniteMobile = CDbl(rstProjet.Fields("TempsUniteMobile")) * CDbl(rstProjet.Fields("TauxUniteMobile"))

        'Correction d'un bug de Type Incompatible
435     If IsNumeric(rstProjet.Fields("PrixEmballage")) Then
440       dblPrixEmballage = CDbl(rstProjet.Fields("PrixEmballage"))
445     Else
450       dblPrixEmballage = 0
455     End If
      
460     dblTotalResteTemps = dblHebergement + dblRepas + dblTransport + dblUniteMobile + dblPrixEmballage
                                                              
465     If IsNumeric(rstProjet.Fields("total_manuel")) Then
470       dblTotalManuel = CDbl(rstProjet.Fields("total_manuel"))
475     Else
480       dblTotalManuel = 0
485     End If
                        
490     dblTotalImprevue = (dblPrixPieces + dblProfit) * CDbl(sImprevue)
    
495     dblPrixTotal = dblPrixPieces + dblProfit + dblTotalTemps + dblTotalImprevue + dblTotalManuel + dblTotalResteTemps
                        
        'Le calcul de la commission est sur les manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps
500     dblCommission = dblPrixTotal * CDbl(sCommission)
        
        'Le prix total est le calcul des manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps + total de la commission
505     dblGrandTotal = dblPrixTotal + dblCommission
                
        'Format monétaires avec 2 chiffres après la virgule
510     rstProjet.Fields("Total_Piece") = Conversion(CStr(Round(dblPrixPieces, 2)), MODE_ARGENT)
515     rstProjet.Fields("Total_Temps") = Conversion(CStr(Round(dblTotalTemps, 2)), MODE_ARGENT)
520     rstProjet.Fields("PrixTotal") = Conversion(CStr(Round(dblGrandTotal, 2)), MODE_ARGENT)
525     rstProjet.Fields("total_Imprevue") = Conversion(CStr(Round(dblTotalImprevue, 2)), MODE_ARGENT)
530     rstProjet.Fields("total_commission") = Conversion(CStr(Round(dblCommission, 2)), MODE_ARGENT)
535     rstProjet.Fields("total_profit") = Conversion(CStr(Round(dblProfit, 2)), MODE_ARGENT)

540     Call rstProjet.Update

545     Call rstPiece.Close
550     Set rstPiece = Nothing

555     Call rstProjet.Close
560     Set rstProjet = Nothing

565     Exit Sub

AfficherErreur:

570     woups "frmRetourMateriel", "CalculerTotalRecordset", Err, Erl
End Sub
