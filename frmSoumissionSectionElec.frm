VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSoumissionSectionElec 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sections électriques"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSoumissionSectionElec.frx":0000
   ScaleHeight     =   4110
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAjout 
      BackColor       =   &H00000000&
      Caption         =   "Ajout de nouvelles sections"
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
      Height          =   1215
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton cmdAnnuler 
         Caption         =   "Annuler"
         Height          =   375
         Left            =   3480
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   3480
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtAnglais 
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtFrancais 
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Anglais"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Français"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdModifier 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Modifier"
      Height          =   495
      Left            =   4440
      TabIndex        =   11
      Top             =   2280
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvwSection 
      Height          =   2895
      Left            =   720
      TabIndex        =   13
      Top             =   1080
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5106
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
         Text            =   "Français"
         Object.Width           =   2924
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Anglais"
         Object.Width           =   2924
      EndProperty
   End
   Begin VB.CommandButton CmdAdd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Ajouter"
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "È"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Ç"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton CmdFerme 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Fermer"
      Height          =   495
      Left            =   4440
      TabIndex        =   14
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton CmdSupp 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Supprimer"
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Section"
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
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmSoumissionSectionElec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const I_COL_FRANCAIS As Integer = 0
Private Const I_COL_ANGLAIS  As Integer = 1

Private m_bAjout     As Boolean

Private Sub Sauvegarde()

5       On Error GoTo AfficherErreur
        
        '''''''''''''''''''''''''''''''''''''''''''''
        'sauvegarde des données dans l'ordre du lister
        '''''''''''''''''''''''''''''''''''''''''''''
10      Dim rstSection As ADODB.Recordset
15      Dim iCompteur  As Integer
  
        'pour tout les donnée dans lister
20      Set rstSection = New ADODB.Recordset
        
25      For iCompteur = 1 To lvwSection.ListItems.count
30        Call rstSection.Open("SELECT NomSectionFR, NomSectionEN, Ordre FROM GRB_SoumProjSectionElec WHERE IDSection = " & lvwSection.ListItems(iCompteur).Tag, g_connData, adOpenDynamic, adLockOptimistic)
      
35        rstSection.Fields("Ordre") = iCompteur
          
40        Call rstSection.Update
      
          'ferme la table
45        Call rstSection.Close
50      Next

55      Set rstSection = Nothing

60      Exit Sub

AfficherErreur:

65      woups "frmSoumissionSectionElec", "Sauvegarde", Err, Erl
End Sub

Private Sub RemplirListViewSection()

5       On Error GoTo AfficherErreur

        'Rempli le ListView des Sections
10      Dim rstSection As ADODB.Recordset
15      Dim itmSection As ListItem

20      Set rstSection = New ADODB.Recordset

25      Call rstSection.Open("SELECT * FROM GRB_SoumProjSectionElec ORDER BY Ordre", g_connData, adOpenDynamic, adLockOptimistic)

        'Il faut vider le ListView avant de le remplir
30      Call lvwSection.ListItems.Clear

        'Tant que ce n'est pas la fin des enregistrements
35      Do While Not rstSection.EOF
40        Set itmSection = lvwSection.ListItems.Add

45        itmSection.Tag = rstSection.Fields("IDSection")

          'Nom en francais
50        itmSection.Text = rstSection.Fields("NomSectionFR")

          'Nom en anglais
55        If Not IsNull(rstSection.Fields("NomSectionEN")) Then
60          itmSection.SubItems(I_COL_ANGLAIS) = rstSection.Fields("NomSectionEN")
65        Else
70          itmSection.SubItems(I_COL_ANGLAIS) = vbNullString
75        End If

80        Call rstSection.MoveNext
85      Loop

90      Call rstSection.Close
95      Set rstSection = Nothing

100     Exit Sub

AfficherErreur:

105     woups "frmSoumissionSectionElec", "RemplirListViewSection", Err, Erl
End Sub

Private Sub CmdAdd_Click()

5       On Error GoTo AfficherErreur

10      m_bAjout = True

15      txtAnglais.Text = vbNullString
20      txtFrancais.Text = vbNullString

25      fraAjout.Visible = True

30      Call txtFrancais.SetFocus

35      Exit Sub

AfficherErreur:

40      woups "frmSoumissionSectionElec", "CmdAdd_Click", Err, Erl
End Sub

Private Sub cmdAnnuler_Click()

5       On Error GoTo AfficherErreur

10      fraAjout.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmSoumissionSectionElec", "cmdAnnuler_Click", Err, Erl
End Sub

Private Sub cmdOK_Click()

5       On Error GoTo AfficherErreur
        
        'proc qui permet d'ajouter un contact à la BD
10      Dim rstSection  As ADODB.Recordset
15      Dim rstMaxOrdre As ADODB.Recordset
   
20      If Trim$(txtFrancais.Text) = vbNullString Or Trim$(txtAnglais.Text) = vbNullString Then
25        Call MsgBox("Le nom en français et en anglais est obligatoire!", vbOKOnly, "Erreur")
    
30        Exit Sub
35      End If
  
40      Screen.MousePointer = vbHourglass
                            
45      Set rstSection = New ADODB.Recordset
                            
50      If m_bAjout = True Then
          'ouvre la table
55        Call rstSection.Open("SELECT * FROM GRB_SoumProjSectionElec WHERE NomSectionFR = '" & Replace(txtFrancais.Text, "'", "''") & "' OR NomSectionEN = '" & Replace(txtAnglais.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
  
          'si n'existe pas
60        If rstSection.EOF Then
65          Set rstMaxOrdre = New ADODB.Recordset

70          Call rstMaxOrdre.Open("SELECT Max(Ordre) as MaxOrdre FROM GRB_SoumProjSectionElec", g_connData, adOpenDynamic, adLockOptimistic)
        
            'ajoute la section
75          Call rstSection.AddNew
            
80          rstSection.Fields("NomSectionFR") = txtFrancais.Text
85          rstSection.Fields("NomSectionEN") = txtAnglais.Text
90          rstSection.Fields("Ordre") = rstMaxOrdre.Fields("MaxOrdre") + 1
            
95          Call rstMaxOrdre.Close
100         Set rstMaxOrdre = Nothing

105         Call rstSection.Update

110         m_bAjout = False
115       Else
120         Call MsgBox("Cette section existe déjà!")
125       End If

130       Call rstSection.Close
135     Else
140       Call rstSection.Open("SELECT * FROM GRB_SoumProjSectionElec WHERE IDSection = " & lvwSection.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)

145       rstSection.Fields("NomSectionFR") = txtFrancais.Text
150       rstSection.Fields("NomSectionEN") = txtAnglais.Text

155       Call rstSection.Update

          'ferme la table
160       Call rstSection.Close
165     End If

170     Set rstSection = Nothing

175     Call RemplirListViewSection

180     fraAjout.Visible = False

185     Screen.MousePointer = vbDefault

190     Exit Sub

AfficherErreur:

195     woups "frmSoumissionSectionElec", "CmdOK_Click", Err, Erl
End Sub

Private Sub cmdDown_Click()

5       On Error GoTo AfficherErreur

        '''''''''''''''''''''''''''''''''''''''''
        'descend la selection d'une ligne
        '''''''''''''''''''''''''''''''''''''''''
10      Dim sTagAvant      As String
15      Dim sTagApres      As String
20      Dim sFrancaisAvant As String
25      Dim sFrancaisApres As String
30      Dim sAnglaisAvant  As String
35      Dim sAnglaisApres  As String
40      Dim iIndex         As Integer
  
45      iIndex = lvwSection.SelectedItem.Index
  
50      If iIndex < lvwSection.ListItems.count Then
          'garde en memoire les données qui vont se repositionné dans la list
55        sTagAvant = lvwSection.ListItems(iIndex).Tag
60        sFrancaisAvant = lvwSection.ListItems(iIndex).Text
65        sAnglaisAvant = lvwSection.ListItems(iIndex).SubItems(I_COL_ANGLAIS)
  
70        sTagApres = lvwSection.ListItems(iIndex + 1).Tag
75        sFrancaisApres = lvwSection.ListItems(iIndex + 1).Text
80        sAnglaisApres = lvwSection.ListItems(iIndex + 1).SubItems(I_COL_ANGLAIS)
  
          'reposition dans la liste
85        lvwSection.ListItems(iIndex).Tag = sTagApres
90        lvwSection.ListItems(iIndex).Text = sFrancaisApres
95        lvwSection.ListItems(iIndex).SubItems(I_COL_ANGLAIS) = sAnglaisApres
  
100       lvwSection.ListItems(iIndex + 1).Tag = sTagAvant
105       lvwSection.ListItems(iIndex + 1).Text = sFrancaisAvant
110       lvwSection.ListItems(iIndex + 1).SubItems(I_COL_ANGLAIS) = sAnglaisAvant

          'descend la selection
115       lvwSection.ListItems(iIndex + 1).Selected = True
    
120       Call lvwSection.SelectedItem.EnsureVisible
125     End If

130     Exit Sub

AfficherErreur:

135     woups "frmSoumissionSectionElec", "cmdDown_Click", Err, Erl
End Sub

Private Sub cmdUp_Click()

5       On Error GoTo AfficherErreur

        ''''''''''''''''''''''''''''''''''
        ' monte la selection d'une ligne '
        ''''''''''''''''''''''''''''''''''
10      Dim sTagAvant      As String
15      Dim sTagApres      As String
20      Dim sFrancaisAvant As String
25      Dim sFrancaisApres As String
30      Dim sAnglaisAvant  As String
35      Dim sAnglaisApres  As String
40      Dim iIndex         As Integer
  
45      iIndex = lvwSection.SelectedItem.Index
  
50      If iIndex > 1 Then
          'garde en memoire les données qui vont se repositionné dans la list
55        sTagAvant = lvwSection.ListItems(iIndex).Tag
60        sFrancaisAvant = lvwSection.ListItems(iIndex).Text
65        sAnglaisAvant = lvwSection.ListItems(iIndex).SubItems(I_COL_ANGLAIS)
    
70        sTagApres = lvwSection.ListItems(iIndex - 1).Tag
75        sFrancaisApres = lvwSection.ListItems(iIndex - 1).Text
80        sAnglaisApres = lvwSection.ListItems(iIndex - 1).SubItems(I_COL_ANGLAIS)
  
          'reposition dans la liste
85        lvwSection.ListItems(iIndex).Tag = sTagApres
90        lvwSection.ListItems(iIndex).Text = sFrancaisApres
95        lvwSection.ListItems(iIndex).SubItems(I_COL_ANGLAIS) = sAnglaisApres
  
100       lvwSection.ListItems(iIndex - 1).Tag = sTagAvant
105       lvwSection.ListItems(iIndex - 1).Text = sFrancaisAvant
110       lvwSection.ListItems(iIndex - 1).SubItems(I_COL_ANGLAIS) = sAnglaisAvant

          'monte la selection
115       lvwSection.ListItems(iIndex - 1).Selected = True
    
120       Call lvwSection.SelectedItem.EnsureVisible
125     End If

130     Exit Sub

AfficherErreur:

135     woups "frmSoumissionSectionElec", "cmdUp_Click", Err, Erl
End Sub

Private Sub CmdFerme_Click()

5       On Error GoTo AfficherErreur

10      Call Sauvegarde

15      Call Unload(Me)

20      Exit Sub

AfficherErreur:

25     woups "frmSoumissionSectionElec", "CmdFerme_Click", Err, Erl
End Sub

Private Sub cmdModifier_Click()

5       On Error GoTo AfficherErreur

10      txtFrancais.Text = lvwSection.SelectedItem.Text
15      txtAnglais.Text = lvwSection.SelectedItem.SubItems(I_COL_ANGLAIS)
 
20      fraAjout.Visible = True

25      Exit Sub

AfficherErreur:

30      woups "frmSoumissionSectionElec", "cmdModifier_Click", Err, Erl
End Sub

Private Sub CmdSupp_Click()

5       On Error GoTo AfficherErreur

10      Call SupprimerSection

15      Exit Sub

AfficherErreur:

20      woups "frmSoumissionSectionElec", "CmdSupp_Click", Err, Erl
End Sub

Private Sub SupprimerSection()

5       On Error GoTo AfficherErreur

10      Dim rstSoumission As ADODB.Recordset

        'fonction qui supprime lenregistrement courant
15      If lvwSection.ListItems.count > 0 Then
20        If MsgBox("Etes-vous sur de supprimer cette enregistrement?", vbYesNo, "Supprimer") = vbYes Then
25          Screen.MousePointer = vbHourglass
               
30          Set rstSoumission = New ADODB.Recordset
               
35          Call rstSoumission.Open("SELECT IDSection FROM GRB_Soumission_pieces WHERE IDSection = " & lvwSection.SelectedItem.Tag & " AND Type = 'E'", g_connData, adOpenDynamic, adLockOptimistic)
      
40          If rstSoumission.EOF Then
              'efface l'enregistrement
45            Call g_connData.Execute("DELETE * FROM GRB_SoumProjSectionElec WHERE IDsection = " & lvwSection.SelectedItem.Tag)
          
              'Si le combo n'est pas vide, on sélectionne le premier
50            If lvwSection.ListItems.count > 0 Then
55              lvwSection.ListItems(1).Selected = True
60            End If
65          Else
70            Call MsgBox("Impossible de supprimer une section déjà utilisé dans une soumission!", vbOKOnly, "Erreur")
75          End If
        
            'ferm la table
80          Call rstSoumission.Close
85          Set rstSoumission = Nothing
        
90          Call RemplirListViewSection
95        End If
100     Else
105       Call MsgBox("Aucun enregistrement sélectionné!")
110     End If
        
115     Screen.MousePointer = vbDefault

120     Exit Sub

AfficherErreur:

125     woups "frmSoumissionSectionElec", "SupprimerSection", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

        'Ouverture de la fenêtre
10      Call RemplirListViewSection

15      Exit Sub

AfficherErreur:

20      woups "frmSoumissionSectionElec", "Form_Load", Err, Erl
End Sub

Private Sub lvwSection_Click()

5       On Error GoTo AfficherErreur
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'met les fleche enabled depandant si au debut ou a la fin
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
10      If lvwSection.ListItems.count > 0 Then
15        If lvwSection.SelectedItem.Index = lvwSection.ListItems.count Then
20          cmdDown.Enabled = False
25        Else
30          cmdDown.Enabled = True
35        End If
    
40        If lvwSection.SelectedItem.Index = 1 Then
45          cmdUp.Enabled = False
50        Else
55          cmdUp.Enabled = True
60        End If
65      End If

70      Exit Sub

AfficherErreur:

75      woups "frmSoumissionSectionElec", "lvwSection_Click", Err, Erl
End Sub

Private Sub lvwSection_DblClick()

5       On Error GoTo AfficherErreur

10      txtFrancais.Text = lvwSection.SelectedItem.Text
15      txtAnglais.Text = lvwSection.SelectedItem.SubItems(I_COL_ANGLAIS)
  
20      fraAjout.Visible = True

25      Exit Sub

AfficherErreur:

30     woups "frmSoumissionSectionElec", "lvwSection_DblClick", Err, Erl
End Sub

Private Sub lvwSection_KeyDown(KeyCode As Integer, Shift As Integer)

5       On Error GoTo AfficherErreur

10      If KeyCode = vbKeyDelete Then
15        Call SupprimerSection
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmSoumissionSectionElec", "lvwSection_KeyDown", Err, Erl
End Sub
