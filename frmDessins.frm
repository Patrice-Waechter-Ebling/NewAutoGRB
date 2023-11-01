VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDessins 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dessins"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmDessins.frx":0000
   ScaleHeight     =   7245
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPièce 
      BackColor       =   &H00000000&
      Caption         =   "Pièce"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   240
      TabIndex        =   9
      Top             =   4080
      Width           =   7815
      Begin VB.CommandButton cmdModifierPiece 
         Caption         =   "Modifier"
         Height          =   375
         Left            =   5640
         TabIndex        =   14
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton cmdAjouterPiece 
         Caption         =   "Ajouter"
         Height          =   375
         Left            =   4560
         TabIndex        =   11
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton cmdSupprimerPiece 
         Caption         =   "Supprimer"
         Height          =   375
         Left            =   6720
         TabIndex        =   10
         Top             =   2520
         Width           =   975
      End
      Begin MSComctlLib.ListView lvwPiece 
         Height          =   2055
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   3625
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "# Dessin"
            Object.Width           =   3413
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   9411
         EndProperty
      End
   End
   Begin VB.Frame fraSousAssemblage 
      BackColor       =   &H00000000&
      Caption         =   "Sous-Assemblage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   7815
      Begin VB.CommandButton cmdModifierSousAssemblage 
         Caption         =   "Modifier"
         Height          =   375
         Left            =   5640
         TabIndex        =   13
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton cmdAjouterSousAssemblage 
         Caption         =   "Ajouter"
         Height          =   375
         Left            =   4560
         TabIndex        =   7
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton cmdSupprimerSousAssemblage 
         Caption         =   "Supprimer"
         Height          =   375
         Left            =   6720
         TabIndex        =   6
         Top             =   2520
         Width           =   975
      End
      Begin MSComctlLib.ListView lvwSousAssemblage 
         Height          =   2055
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   3625
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "# Dessin"
            Object.Width           =   3413
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   9412
         EndProperty
      End
   End
   Begin VB.CommandButton cmdSupprimerProjet 
      Caption         =   "Supprimer"
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdAjouterProjet 
      Caption         =   "Ajouter"
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox cmbProjet 
      Height          =   315
      Left            =   4320
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtProjet 
      Height          =   285
      Left            =   4320
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Projet :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frmDessins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const I_COL_NO_DESSIN   As Integer = 0
Private Const I_COL_DESCRIPTION As Integer = 1

Private Enum enumListe
  SOUS_ASSEMBLAGE = 0
  PIECE = 1
End Enum

Public m_sDessin      As String
Public m_sDescription As String
Public m_bAnnuleAjout As Boolean

Private Sub cmbProjet_Click()
        'Affiche les dessins selon le projet sélectionné
5       On Error GoTo AfficherErreur

10      txtProjet.Text = cmbProjet.Text

15      Call AfficherDessins

20      Exit Sub

AfficherErreur:

25      Call AfficherErreur(Me, "cmbProjet_Click", Err, Erl)
End Sub

Private Sub cmdAjouterSousAssemblage_Click()
        'Ajoute un sous-assemblage
5       On Error GoTo AfficherErreur

10      If cmbProjet.ListCount > 0 Then
15        Call AjouterDessin(SOUS_ASSEMBLAGE)
20      Else
25        Call MsgBox("Il n'y a pas de projet!", vbOKOnly, "Erreur")
30      End If

35      Exit Sub

AfficherErreur:

40      Call AfficherErreur(Me, "cmdAjouterSousAssemblage_Click", Err, Erl)
End Sub

Private Sub cmdAjouterPiece_Click()
        'Ajoute une piece
5       On Error GoTo AfficherErreur

10      If cmbProjet.ListCount > 0 Then
15        Call AjouterDessin(PIECE)
20      Else
25        Call MsgBox("Il n'y a pas de projet!", vbOKOnly, "Erreur")
30      End If

35      Exit Sub

AfficherErreur:

40      Call AfficherErreur(Me, "cmdAjouterPiece_Click", Err, Erl)
End Sub

Private Sub AjouterDessin(ByVal eListe As enumListe)
        'Ajoute un dessin
5       On Error GoTo AfficherErreur

10      Dim rstDessin As ADODB.Recordset
15      Dim itmDessin As ListItem

20      Call OuvrirForm(frmAjoutDessin, True)

        'Si l'ajout n'a pas été annulé
25      If m_bAnnuleAjout = False Then
          'Ouverture de la table
30        Set rstDessin = New ADODB.Recordset
          
35        Call rstDessin.Open("SELECT * FROM GRB_Dessins WHERE NoDessin = '" & m_sDessin & "'", g_connData, adOpenDynamic, adLockOptimistic)

40        If rstDessin.EOF Then
            'Ajoute dans la table
45          Call rstDessin.AddNew

50          rstDessin.Fields("NoProjet") = txtProjet.Text
55          rstDessin.Fields("NoDessin") = m_sDessin
60          rstDessin.Fields("Description") = m_sDescription

65          If eListe = SOUS_ASSEMBLAGE Then
70            rstDessin.Fields("Type") = "SA"
75          Else
80            rstDessin.Fields("Type") = "P"
85          End If

90          Call rstDessin.Update

95          Call AfficherDessins
100       Else
105         Call MsgBox("Ce dessin existe déjà!", vbOKOnly, "Erreur")
110       End If
115     End If
        
120     Exit Sub

AfficherErreur:

125     Call AfficherErreur(Me, "AjouterDessin", Err, Erl)
End Sub

Private Sub cmdModifierSousAssemblage_Click()
        'Modifie un Sous-Assemblage
5       On Error GoTo AfficherErreur

10      If lvwSousAssemblage.ListItems.Count > 0 Then
15        Call ModifierDessin(SOUS_ASSEMBLAGE)
20      End If

25      Exit Sub

AfficherErreur:

30      Call AfficherErreur(Me, "cmdModifierSousAssemblage_Click", Err, Erl)
End Sub

Private Sub cmdModifierPiece_Click()
        'Modifie une pièce
10      On Error GoTo AfficherErreur

20      If lvwPiece.ListItems.Count > 0 Then
30        Call ModifierDessin(PIECE)
40      End If

50      Exit Sub

AfficherErreur:

60      Call AfficherErreur(Me, "cmdModifierPiece_Click", Err, Erl)
End Sub

Private Sub lvwSousAssemblage_DblClick()
        'Modifie un Sous-Assemblage
10      On Error GoTo AfficherErreur

20      If lvwSousAssemblage.ListItems.Count > 0 Then
30        Call ModifierDessin(SOUS_ASSEMBLAGE)
40      End If

50      Exit Sub

AfficherErreur:

60      Call AfficherErreur(Me, "lvwSousAssemblage_DblClick", Err, Erl)
End Sub

Private Sub lvwPiece_DblClick()
        'Modifie une pièce
10      On Error GoTo AfficherErreur

20      If lvwPiece.ListItems.Count > 0 Then
30        Call ModifierDessin(PIECE)
40      End If

50      Exit Sub

AfficherErreur:

60      Call AfficherErreur(Me, "lvwPiece_DblClick", Err, Erl)
End Sub

Private Sub ModifierDessin(ByVal eListe As enumListe)
        'Modifie un dessin
5       On Error GoTo AfficherErreur

10      Dim rstDessin As ADODB.Recordset

15      If eListe = PIECE Then
20        Call frmAjoutDessin.Afficher(lvwPiece.SelectedItem.Text, lvwPiece.SelectedItem.SubItems(I_COL_DESCRIPTION))
25      Else
30        Call frmAjoutDessin.Afficher(lvwSousAssemblage.SelectedItem.Text, lvwSousAssemblage.SelectedItem.SubItems(I_COL_DESCRIPTION))
35      End If

40      If m_bAnnuleAjout = False Then
45        Set rstDessin = New ADODB.Recordset

50        If eListe = PIECE Then
55          Call rstDessin.Open("SELECT * FROM GRB_Dessins WHERE NoEnreg = " & lvwPiece.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
60        Else
65          Call rstDessin.Open("SELECT * FROM GRB_Dessins WHERE NoEnreg = " & lvwSousAssemblage.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
70        End If

75        rstDessin.Fields("NoDessin") = m_sDessin
80        rstDessin.Fields("Description") = m_sDescription

85        Call rstDessin.Update

90        Call rstDessin.Close
95        Set rstDessin = Nothing

100       Call AfficherDessins
105     End If

110     Exit Sub

AfficherErreur:

115     Call AfficherErreur(Me, "ModifierDessin", Err, Erl)
End Sub

Private Sub cmdSupprimerSousAssemblage_Click()
        'Supprime un sous-assemblage
10      On Error GoTo AfficherErreur

20      If lvwSousAssemblage.ListItems.Count > 0 Then
30        Call SupprimerDessin(SOUS_ASSEMBLAGE)
40      End If

50      Exit Sub

AfficherErreur:

60      Call AfficherErreur(Me, "cmdSupprimerSousAssemblage_Click", Err, Erl)
End Sub

Private Sub cmdSupprimerPiece_Click()
        'Supprimer une piece
10      On Error GoTo AfficherErreur

20      If lvwPiece.ListItems.Count > 0 Then
30        Call SupprimerDessin(PIECE)
40      End If

50      Exit Sub

AfficherErreur:

60      Call AfficherErreur(Me, "cmdSupprimerPiece_Click", Err, Erl)
End Sub

Private Sub lvwSousAssemblage_KeyDown(KeyCode As Integer, Shift As Integer)
        
10      On Error GoTo AfficherErreur

20      If lvwSousAssemblage.ListItems.Count > 0 Then
30        If KeyCode = vbKeyDelete Then
40          Call SupprimerDessin(SOUS_ASSEMBLAGE)
50        End If
60      End If

70      Exit Sub

AfficherErreur:

80      Call AfficherErreur(Me, "lvwSousAssemblage_KeyDown", Err, Erl)
End Sub

Private Sub lvwPiece_KeyDown(KeyCode As Integer, Shift As Integer)

10      On Error GoTo AfficherErreur

20      If lvwPiece.ListItems.Count > 0 Then
30        If KeyCode = vbKeyDelete Then
40          Call SupprimerDessin(PIECE)
50        End If
60      End If

70      Exit Sub

AfficherErreur:

80      Call AfficherErreur(Me, "lvwPiece_KeyDown", Err, Erl)
End Sub

Private Sub SupprimerDessin(ByVal eListe As enumListe)
        'Efface un dessin
5       On Error GoTo AfficherErreur

10      If MsgBox("Voulez-vous vraiment effacer ce dessin?", vbYesNo) = vbYes Then
15        If eListe = PIECE Then
20          Call g_connData.Execute("DELETE * FROM GRB_Dessins WHERE NoEnreg = " & lvwPiece.SelectedItem.Tag)
25        Else
30          Call g_connData.Execute("DELETE * FROM GRB_Dessins WHERE NoEnreg = " & lvwSousAssemblage.SelectedItem.Tag)
35        End If

40        Call AfficherDessins
45      End If

50      Exit Sub

AfficherErreur:

55      Call AfficherErreur(Me, "SupprimerDessin", Err, Erl)
End Sub

Private Sub Form_Load()
        'Ouverture du formulaire
10      On Error GoTo AfficherErreur

20      Call AfficherBoutonsGroupe

30      Call RemplirComboProjets

40      Exit Sub

AfficherErreur:

50      Call AfficherErreur(Me, "Form_Load", Err, Erl)
End Sub

Private Sub AfficherBoutonsGroupe()

10      On Error GoTo AfficherErreur

20      cmdAjouterProjet.Enabled = g_bModificationDessin
30      cmdSupprimerProjet.Enabled = g_bModificationDessin
40      cmdAjouterSousAssemblage.Enabled = g_bModificationDessin
50      cmdSupprimerSousAssemblage.Enabled = g_bModificationDessin
60      cmdAjouterPiece.Enabled = g_bModificationDessin
70      cmdSupprimerPiece.Enabled = g_bModificationDessin

80      Exit Sub

AfficherErreur:

90      Call AfficherErreur(Me, "AfficherBoutonsGroupe", Err, Erl)
End Sub

Private Sub RemplirComboProjets()
        'Rempli le combo des projets
5       On Error GoTo AfficherErreur

10      Dim rstProjet As ADODB.Recordset

        'Vide le combo
15      Call cmbProjet.Clear

        'Ouverture de la table
20      Set rstProjet = New ADODB.Recordset
        
25      Call rstProjet.Open("SELECT NoProjet FROM GRB_ProjetsDessins", g_connData, adOpenDynamic, adLockOptimistic)
        
        'Tant qu'il y a des enregistrements
30      Do While Not rstProjet.EOF
          'On ajoute le projet
35        Call cmbProjet.AddItem(rstProjet.Fields("NoProjet"))

40        Call rstProjet.MoveNext
45      Loop

        'Fermeture de la table
50      Call rstProjet.Close
55      Set rstProjet = Nothing

        'Si le combo n'est pas vide
60      If cmbProjet.ListCount > 0 Then
          'On sélectionne le premier élément
65        cmbProjet.ListIndex = 0
70      Else
          'Sinon, on vide les listes
75        Call lvwSousAssemblage.ListItems.Clear
80        Call lvwPiece.ListItems.Clear
85      End If

90      Exit Sub

AfficherErreur:

95      Call AfficherErreur(Me, "RemplirComboProjets", Err, Erl)
End Sub

Private Sub AfficherDessins()
        'Affiche les dessins selon le projet sélectionné
5       On Error GoTo AfficherErreur

10      Dim rstDessin As ADODB.Recordset
15      Dim itmDessin As ListItem

        'Vide les listes
20      Call lvwSousAssemblage.ListItems.Clear
25      Call lvwPiece.ListItems.Clear

        'Ouverture de la table
30      Set rstDessin = New ADODB.Recordset
        
35      Call rstDessin.Open("SELECT * FROM GRB_Dessins WHERE NoProjet = '" & txtProjet.Text & "' ORDER BY NoDessin", g_connData, adOpenDynamic, adLockOptimistic)

        'Tant qu'il y a des enregistrements
40      Do While Not rstDessin.EOF
          'Si c'est un sous-assemblage
45        If rstDessin.Fields("Type") = "SA" Then
            'On ajoute dans lvwSousAssemblage
50          Set itmDessin = lvwSousAssemblage.ListItems.Add
55        Else
            'Sinon, on ajoute dans lvwPiece
60          Set itmDessin = lvwPiece.ListItems.Add
65        End If

70        itmDessin.Text = rstDessin.Fields("NoDessin")
75        itmDessin.SubItems(I_COL_DESCRIPTION) = rstDessin.Fields("Description")

80        itmDessin.Tag = rstDessin.Fields("NoEnreg")

85        Call rstDessin.MoveNext
90      Loop
        
        'Fermeture de la table
95      Call rstDessin.Close
100     Set rstDessin = Nothing

        'Met le focus sur le dernier enregistrement
105     If lvwSousAssemblage.ListItems.Count > 0 Then
110       lvwSousAssemblage.ListItems(lvwSousAssemblage.ListItems.Count).Selected = True
115       Call lvwSousAssemblage.ListItems(lvwSousAssemblage.ListItems.Count).EnsureVisible
120     End If

125     If lvwPiece.ListItems.Count > 0 Then
130       lvwPiece.ListItems(lvwPiece.ListItems.Count).Selected = True
135       Call lvwPiece.ListItems(lvwPiece.ListItems.Count).EnsureVisible
140     End If

145     Exit Sub

AfficherErreur:

150     Call AfficherErreur(Me, "AfficherDessins", Err, Erl)
End Sub

Private Sub cmdAjouterProjet_Click()
        'Ajoute un projet
5       On Error GoTo AfficherErreur

10      Dim rstProjet As ADODB.Recordset
15      Dim sNoProjet As String

        'Saisie du numéro de projet
20      sNoProjet = InputBox("Quel est le numéro du projet?")

        'Si le numéro n'est pas vide
25      If Trim(sNoProjet) <> "" Then
          'Ouverture de la projet
30        Set rstProjet = New ADODB.Recordset
          
35        Call rstProjet.Open("SELECT * FROM GRB_ProjetsDessins WHERE NoProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

40        If rstProjet.EOF Then
45          Call rstProjet.AddNew

50          rstProjet.Fields("NoProjet") = sNoProjet

55          Call rstProjet.Update

60          Call RemplirComboProjets

65          Call rstProjet.Close
70          Set rstProjet = Nothing
75        Else
80          Call MsgBox("Ce numéro existe déjà!", vbOKOnly, "Erreur")
85        End If
90      End If

95      Exit Sub

AfficherErreur:

100      Call AfficherErreur(Me, "cmdAjouterProjet_Click", Err, Erl)
End Sub

Private Sub cmdSupprimerProjet_Click()
        'Supprimer le projet sélectionné
10      On Error GoTo AfficherErreur

20      If cmbProjet.ListCount > 0 Then
          'Demande de confirmation
30        If MsgBox("Voulez-vous vraiment effacer le projet " & txtProjet.Text, vbYesNo) = vbYes Then
            'Efface le projet
40          Call g_connData.Execute("DELETE * FROM GRB_ProjetsDessins WHERE NoProjet = '" & txtProjet.Text & "'")

            'Efface les dessins
50          Call g_connData.Execute("DELETE * FROM GRB_Dessins WHERE NoProjet = '" & txtProjet.Text & "'")

60          Call RemplirComboProjets
70        End If
80      End If

90      Exit Sub

AfficherErreur:

100     Call AfficherErreur(Me, "cmdSupprimerProjet_Click", Err, Erl)
End Sub
