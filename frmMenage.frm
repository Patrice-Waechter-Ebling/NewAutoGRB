VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMenage 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ménage dans le catalogue"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMenage.frx":0000
   ScaleHeight     =   5745
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   6360
      TabIndex        =   21
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   7560
      TabIndex        =   22
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Frame fraModif 
      BackColor       =   &H00000000&
      Caption         =   "Modification de la pièce"
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
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   8535
      Begin VB.CommandButton cmdAnnulerModif 
         Caption         =   "Annuler"
         Height          =   375
         Left            =   6120
         TabIndex        =   19
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton cmdEnregistrer 
         Caption         =   "Enregistrer"
         Height          =   375
         Left            =   7320
         TabIndex        =   20
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CheckBox chkUR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "UR"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4140
         TabIndex        =   18
         Top             =   2160
         Width           =   855
      End
      Begin VB.CheckBox chkUL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "UL"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4140
         TabIndex        =   17
         Top             =   1920
         Width           =   855
      End
      Begin VB.CheckBox chkCE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "CE"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4140
         TabIndex        =   16
         Top             =   1680
         Width           =   855
      End
      Begin VB.CheckBox chkCUL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "CUL"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4140
         TabIndex        =   15
         Top             =   1440
         Width           =   855
      End
      Begin VB.CheckBox chkCSA 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "CSA"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4140
         TabIndex        =   14
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtTemps 
         Height          =   285
         Left            =   1680
         MaxLength       =   7
         TabIndex        =   8
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox txtLargeur 
         Height          =   285
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   9
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox txtHauteur 
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Top             =   3720
         Width           =   1575
      End
      Begin VB.TextBox txtEpaisseur 
         Height          =   285
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   10
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox txtPageCat 
         Height          =   285
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   7
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtComment 
         Height          =   285
         Left            =   1680
         MaxLength       =   41
         TabIndex        =   6
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtDescrEN 
         Height          =   285
         Left            =   4800
         MaxLength       =   61
         TabIndex        =   13
         Top             =   840
         Width           =   3615
      End
      Begin VB.TextBox txtDescrFR 
         Height          =   285
         Left            =   4800
         MaxLength       =   61
         TabIndex        =   12
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox txtFabricant 
         Height          =   285
         Left            =   1680
         MaxLength       =   31
         TabIndex        =   5
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtCategorie 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtPieceGRB 
         Height          =   285
         Left            =   1680
         MaxLength       =   21
         TabIndex        =   4
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtPiece 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "TEMPS :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "LARGEUR :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "HAUTEUR :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "EPAISSEUR :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "PAGE_CAT :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "COMMENT :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "DESCR_EN :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3600
         TabIndex        =   29
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "DESCR_FR :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3600
         TabIndex        =   28
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "FABRICANT :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "CATEGORIE :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PIECE_GRB:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PIECE :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   840
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView lvwMenage 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7223
      View            =   3
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
      NumItems        =   17
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "PIECE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "PIECE_GRB"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CATEGORIE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "FABRICANT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "DESCR_FR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "DESCR_EN"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "COMMENT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "PAGE_CAT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "EPAISSEUR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "HAUTEUR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "LARGEUR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "TEMPS"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "CSA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Text            =   "CUL"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   14
         Text            =   "CE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   15
         Text            =   "UL"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   16
         Text            =   "UR"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblPiece 
      BackStyle       =   0  'Transparent
      Caption         =   "Quelle pièce est la bonne ?"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   840
      Width           =   2655
   End
End
Attribute VB_Name = "frmMenage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Index des colonnes de lvwMenage
Private Const I_COL_PIECE       As Integer = 0
Private Const I_COL_PIECE_GRB   As Integer = 1
Private Const I_COL_CATEGORIE   As Integer = 2
Private Const I_COL_FABRICANT   As Integer = 3
Private Const I_COL_DESCR_FR    As Integer = 4
Private Const I_COL_DESCR_EN    As Integer = 5
Private Const I_COL_COMMENTAIRE As Integer = 6
Private Const I_COL_PAGE_CAT    As Integer = 7
Private Const I_COL_EPAISSEUR   As Integer = 8
Private Const I_COL_HAUTEUR     As Integer = 9
Private Const I_COL_LARGEUR     As Integer = 10
Private Const I_COL_TEMPS       As Integer = 11
Private Const I_COL_CSA         As Integer = 12
Private Const I_COL_CUL         As Integer = 13
Private Const I_COL_CE          As Integer = 14
Private Const I_COL_UR          As Integer = 15
Private Const I_COL_UL          As Integer = 16

Private Const S_SELECT_NORMAL   As String = "FABRICANT,PIECE,PIECE_GRB," & _
                                            "DESCR_FR,DESCR_EN,COMMENT," & _
                                            "PAGE_CAT,EPAISSEUR,HAUTEUR," & _
                                            "LARGEUR,TEMPS,CSA,CUL,CE,UR,UL," & _
                                            "VIGNETTE, IMPLANT, [{SEE_SET}]"
                                                
Private Const S_SELECT_CABLE    As String = "FABRICANT,PIECE,PIECE_GRB," & _
                                            "DESCR_FR,DESCR_EN,COMMENT," & _
                                            "PAGE_CAT,TEMPS,CSA,CUL,CE,UR,UL," & _
                                            "[{SEE_SET}]"
                                            
Private Enum enumMode
  MODE_INACTIF = 0
  MODE_MODIF = 1
End Enum
                                            
'Pour savoir l'index du listItem qui est en train de se faire modifier
Private m_iIndexModif As Integer

Public Sub Afficher(ByVal sPiece As String)
  
5       On Error GoTo AfficherErreur

10      Call RemplirListView(sPiece)

15      Call AfficherControles(MODE_INACTIF)

20      Call Me.Show(vbModal)

25      Exit Sub

AfficherErreur:

30      Call AfficherErreur(Me, "Afficher", Err, Erl)
End Sub

Private Sub AfficherControles(ByVal eMode As enumMode)

5       On Error GoTo AfficherErreur

10      Dim bFrame    As Boolean
15      Dim bListView As Boolean

        Select Case eMode
          Case MODE_MODIF:   bFrame = True
          Case MODE_INACTIF: bListView = True
        End Select
        
20      cmdOK.Visible = bListView
25      cmdAnnuler.Visible = bListView
30      lblPiece.Visible = bListView
35      lvwMenage.Visible = bListView

40      fraModif.Visible = bFrame

45      Exit Sub

AfficherErreur:

50      Call AfficherErreur(Me, "AfficherControles", Err, Erl)
End Sub

Private Sub RemplirListView(ByVal sPiece As String)

5       On Error GoTo AfficherErreur

10      Dim itmMenage As ListItem
15      Dim rstPiece  As ADODB.Recordset
20      Dim iCompteur As Integer
25      Dim sTable    As String

'30      If OuvrirConnectionEquipment = True Then
          'Pour chaque table électrique
35        Set rstPiece = New ADODB.Recordset
          
40        For iCompteur = 0 To FrmCatalogueElec.cmbTable.ListCount - 1
45          sTable = FrmCatalogueElec.cmbTable.List(iCompteur)
                    
            'Si ce n'est pas la table Cable
50          If sTable <> "CABLE" Then
55            Call rstPiece.Open("SELECT " & S_SELECT_NORMAL & " FROM [" & sTable & "] WHERE Trim(PIECE) = '" & Replace(Trim(sPiece), "'", "''") & "'", g_connEquipment, adOpenDynamic, adLockOptimistic)
60          Else
65            Call rstPiece.Open("SELECT " & S_SELECT_CABLE & " FROM [" & sTable & "] WHERE Trim(PIECE) = '" & Replace(Trim(sPiece), "'", "''") & "'", g_connEquipment, adOpenDynamic, adLockOptimistic)
70          End If
          
            'Tant que ce n'est pas la fin des enregistrements
75          Do While Not rstPiece.EOF
              'On rempli le ListView
80            Set itmMenage = lvwMenage.ListItems.Add

              'PIECE
85            itmMenage.Text = rstPiece.Fields("PIECE")

              'PIECE_GRB
90            If Not IsNull(rstPiece.Fields("PIECE_GRB")) Then
95              itmMenage.SubItems(I_COL_PIECE_GRB) = rstPiece.Fields("PIECE_GRB")
100           Else
105             itmMenage.SubItems(I_COL_PIECE_GRB) = vbNullString
110           End If

              'CATEGORIE
115           itmMenage.SubItems(I_COL_CATEGORIE) = sTable

              'FABRICANT
120           If Not IsNull(rstPiece.Fields("FABRICANT")) Then
125             itmMenage.SubItems(I_COL_FABRICANT) = rstPiece.Fields("FABRICANT")
130           Else
135             itmMenage.SubItems(I_COL_FABRICANT) = vbNullString
140           End If

              'DESCR_FR
145           If Not IsNull(rstPiece.Fields("DESCR_FR")) Then
150             itmMenage.SubItems(I_COL_DESCR_FR) = rstPiece.Fields("DESCR_FR")
155           Else
160             itmMenage.SubItems(I_COL_DESCR_FR) = vbNullString
165           End If

              'DESCR_EN
170           If Not IsNull(rstPiece.Fields("DESCR_EN")) Then
175             itmMenage.SubItems(I_COL_DESCR_EN) = rstPiece.Fields("DESCR_EN")
180           Else
185             itmMenage.SubItems(I_COL_DESCR_EN) = vbNullString
190           End If

              'COMMENT
195           If Not IsNull(rstPiece.Fields("COMMENT")) Then
200             itmMenage.SubItems(I_COL_COMMENTAIRE) = rstPiece.Fields("COMMENT")
205           Else
210             itmMenage.SubItems(I_COL_COMMENTAIRE) = vbNullString
215           End If

              'PAGE_CAT
220           If Not IsNull(rstPiece.Fields("PAGE_CAT")) Then
225             itmMenage.SubItems(I_COL_PAGE_CAT) = rstPiece.Fields("PAGE_CAT")
230           Else
235             itmMenage.SubItems(I_COL_PAGE_CAT) = vbNullString
240           End If

              'Si ce n'est pas CABLE
245           If sTable <> "CABLE" Then
                'EPAISSEUR
250             If Not IsNull(rstPiece.Fields("EPAISSEUR")) Then
255               itmMenage.SubItems(I_COL_EPAISSEUR) = rstPiece.Fields("EPAISSEUR")
260             Else
265               itmMenage.SubItems(I_COL_EPAISSEUR) = vbNullString
270             End If

                'HAUTEUR
275             If Not IsNull(rstPiece.Fields("HAUTEUR")) Then
280               itmMenage.SubItems(I_COL_HAUTEUR) = rstPiece.Fields("HAUTEUR")
285             Else
290               itmMenage.SubItems(I_COL_HAUTEUR) = vbNullString
295             End If

                'LARGEUR
300             If Not IsNull(rstPiece.Fields("LARGEUR")) Then
305               itmMenage.SubItems(I_COL_LARGEUR) = rstPiece.Fields("LARGEUR")
310             Else
315               itmMenage.SubItems(I_COL_LARGEUR) = vbNullString
320             End If
325           End If

              'TEMPS
330           If Not IsNull(rstPiece.Fields("TEMPS")) Then
335             itmMenage.SubItems(I_COL_TEMPS) = rstPiece.Fields("TEMPS")
340           Else
345             itmMenage.SubItems(I_COL_TEMPS) = vbNullString
350           End If

              'CSA
355           If Not IsNull(rstPiece.Fields("CSA")) Then
360             If rstPiece.Fields("CSA") = "1" Then
365               itmMenage.SubItems(I_COL_CSA) = "X"
370             Else
375               itmMenage.SubItems(I_COL_CSA) = " "
380             End If
385           Else
390             itmMenage.SubItems(I_COL_CSA) = " "
395           End If

              'CUL
400           If Not IsNull(rstPiece.Fields("CUL")) Then
405             If rstPiece.Fields("CUL") = "1" Then
410               itmMenage.SubItems(I_COL_CUL) = "X"
415             Else
420               itmMenage.SubItems(I_COL_CUL) = " "
425             End If
430           Else
435             itmMenage.SubItems(I_COL_CUL) = " "
440           End If

              'CE
445           If Not IsNull(rstPiece.Fields("CE")) Then
450             If rstPiece.Fields("CE") = "1" Then
455               itmMenage.SubItems(I_COL_CE) = "X"
460             Else
465               itmMenage.SubItems(I_COL_CE) = " "
470             End If
475           Else
480             itmMenage.SubItems(I_COL_CE) = " "
485           End If

              'UR
490           If Not IsNull(rstPiece.Fields("UR")) Then
495             If rstPiece.Fields("UR") = "1" Then
500               itmMenage.SubItems(I_COL_UR) = "X"
505             Else
510               itmMenage.SubItems(I_COL_UR) = " "
515             End If
520           Else
525             itmMenage.SubItems(I_COL_UR) = " "
530           End If

              'UL
535           If Not IsNull(rstPiece.Fields("UL")) Then
540             If rstPiece.Fields("UL") = "1" Then
545               itmMenage.SubItems(I_COL_UL) = "X"
550             Else
555               itmMenage.SubItems(I_COL_UL) = " "
560             End If
565           Else
570             itmMenage.SubItems(I_COL_UL) = " "
575           End If

580           If sTable <> "CABLE" Then
                'IMPLANT
585             If Not IsNull(rstPiece.Fields("IMPLANT")) Then
590               itmMenage.Tag = rstPiece.Fields("IMPLANT")
595             Else
600               itmMenage.Tag = "Null"
605             End If

                'VIGNETTE
610             If Not IsNull(rstPiece.Fields("VIGNETTE")) Then
615               itmMenage.ListSubItems(I_COL_PIECE_GRB).Tag = rstPiece.Fields("VIGNETTE")
620             Else
625               itmMenage.ListSubItems(I_COL_PIECE_GRB).Tag = "Null"
630             End If
635           End If

              '{SEE_SET}
640           If Not IsNull(rstPiece.Fields("{SEE_SET}")) Then
645             itmMenage.ListSubItems(I_COL_CATEGORIE).Tag = rstPiece.Fields("{SEE_SET}")
650           Else
655             itmMenage.ListSubItems(I_COL_CATEGORIE).Tag = "Null"
660           End If

665           Call rstPiece.MoveNext
670         Loop

675         Call rstPiece.Close
680       Next

685       Set rstPiece = Nothing

'690       Call FermerConnectionEquipment
'695     End If

700     Exit Sub

AfficherErreur:

705     Call AfficherErreur(Me, "RemplirListView", Err, Erl)
End Sub

Private Sub cmdAnnuler_Click()

5       On Error GoTo AfficherErreur

10      FrmCatalogueElec.m_bPieceEffacée = False

15      Call Unload(Me)

20      Exit Sub

AfficherErreur:

25      Call AfficherErreur(Me, "cmdAnnuler_Click", Err, Erl)
End Sub

Private Sub cmdAnnulerModif_Click()

5       On Error GoTo AfficherErreur

10      Call AfficherControles(MODE_INACTIF)

15      Exit Sub

AfficherErreur:

20      Call AfficherErreur(Me, "cmdAnnulerModif_Click", Err, Erl)
End Sub

Private Sub cmdEnregistrer_Click()

5       On Error GoTo AfficherErreur

10      Dim itmMenage As ListItem

15      Set itmMenage = lvwMenage.ListItems(m_iIndexModif)

20      itmMenage.Text = txtPiece.Text
25      itmMenage.SubItems(I_COL_PIECE_GRB) = txtPieceGRB.Text
30      itmMenage.SubItems(I_COL_CATEGORIE) = txtCategorie.Text
35      itmMenage.SubItems(I_COL_FABRICANT) = txtFabricant.Text
40      itmMenage.SubItems(I_COL_DESCR_FR) = txtDescrFR.Text
45      itmMenage.SubItems(I_COL_DESCR_EN) = txtDescrEN.Text
50      itmMenage.SubItems(I_COL_COMMENTAIRE) = txtComment.Text
55      itmMenage.SubItems(I_COL_PAGE_CAT) = txtPageCat.Text
60      itmMenage.SubItems(I_COL_EPAISSEUR) = txtEpaisseur.Text
65      itmMenage.SubItems(I_COL_HAUTEUR) = txtHauteur.Text
70      itmMenage.SubItems(I_COL_LARGEUR) = txtLargeur.Text
75      itmMenage.SubItems(I_COL_TEMPS) = txtTemps.Text

80      If chkCSA.Value = vbChecked Then
85        itmMenage.SubItems(I_COL_CSA) = "X"
90      Else
95        itmMenage.SubItems(I_COL_CSA) = " "
100     End If

105     If chkCUL.Value = vbChecked Then
110       itmMenage.SubItems(I_COL_CUL) = "X"
115     Else
120       itmMenage.SubItems(I_COL_CUL) = " "
125     End If
            
130     If chkCE.Value = vbChecked Then
135       itmMenage.SubItems(I_COL_CE) = "X"
140     Else
145       itmMenage.SubItems(I_COL_CE) = " "
150     End If

155     If chkUR.Value = vbChecked Then
160       itmMenage.SubItems(I_COL_UR) = "X"
165     Else
170       itmMenage.SubItems(I_COL_UR) = " "
175     End If

180     If chkUL.Value = vbChecked Then
185       itmMenage.SubItems(I_COL_UL) = "X"
190     Else
195       itmMenage.SubItems(I_COL_UL) = " "
200     End If

205     Call AfficherControles(MODE_INACTIF)

210     Exit Sub

AfficherErreur:

215     Call AfficherErreur(Me, "cmdEnregistrer_Click", Err, Erl)
End Sub

Private Sub lvwMenage_DblClick()

5       On Error GoTo AfficherErreur

10      Dim itmMenage As ListItem

15      Call ViderChamps

20      m_iIndexModif = lvwMenage.SelectedItem.Index

25      Set itmMenage = lvwMenage.SelectedItem

30      txtPiece.Text = itmMenage.Text
35      txtPieceGRB.Text = itmMenage.SubItems(I_COL_PIECE_GRB)
40      txtCategorie.Text = itmMenage.SubItems(I_COL_CATEGORIE)
45      txtFabricant.Text = itmMenage.SubItems(I_COL_FABRICANT)
50      txtDescrFR.Text = itmMenage.SubItems(I_COL_DESCR_FR)
55      txtDescrEN.Text = itmMenage.SubItems(I_COL_DESCR_EN)
60      txtComment.Text = itmMenage.SubItems(I_COL_COMMENTAIRE)
65      txtPageCat.Text = itmMenage.SubItems(I_COL_PAGE_CAT)
70      txtEpaisseur.Text = itmMenage.SubItems(I_COL_EPAISSEUR)
75      txtHauteur.Text = itmMenage.SubItems(I_COL_HAUTEUR)
80      txtLargeur.Text = itmMenage.SubItems(I_COL_LARGEUR)
85      txtTemps.Text = itmMenage.SubItems(I_COL_TEMPS)

90      If itmMenage.SubItems(I_COL_CSA) = "X" Then
95        chkCSA.Value = vbChecked
100     End If

105     If itmMenage.SubItems(I_COL_CUL) = "X" Then
110       chkCUL.Value = vbChecked
115     End If

120     If itmMenage.SubItems(I_COL_CE) = "X" Then
125       chkCE.Value = vbChecked
130     End If

135     If itmMenage.SubItems(I_COL_UL) = "X" Then
140       chkUL.Value = vbChecked
145     End If

150     If itmMenage.SubItems(I_COL_UR) = "X" Then
155       chkUR.Value = vbChecked
160     End If

165     If itmMenage.SubItems(I_COL_CATEGORIE) = "CABLE" Then
170       txtEpaisseur.Enabled = False
175       txtHauteur.Enabled = False
180       txtLargeur.Enabled = False
185     Else
190       txtEpaisseur.Enabled = True
195       txtHauteur.Enabled = True
200       txtLargeur.Enabled = True
205     End If

210     Call AfficherControles(MODE_MODIF)

215     Exit Sub

AfficherErreur:

220     Call AfficherErreur(Me, "lvwMenage_DblClick", Err, Erl)
End Sub

Private Sub lvwMenage_ItemCheck(ByVal Item As MSComctlLib.ListItem)

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer

15      If Item.Checked = True Then
          'Enlève tous les crochets
20        For iCompteur = 1 To lvwMenage.ListItems.Count
25          lvwMenage.ListItems(iCompteur).Checked = False
30        Next

          'Remet le crochet de celui sélectionné
35        Item.Checked = True
40      Else
45        Item.Checked = False
50      End If

55      Exit Sub

AfficherErreur:

60      Call AfficherErreur(Me, "lvwMenage_ItemCheck", Err, Erl)
End Sub

Private Sub cmdOK_Click()

5       On Error GoTo AfficherErreur

10      Dim itmMenage   As ListItem
15      Dim rstPiece    As ADODB.Recordset
20      Dim rstPieceFRS As ADODB.Recordset
25      Dim bChecked    As Boolean
30      Dim sCategorie  As String
35      Dim iIndexCheck As Integer
40      Dim iCompteur   As Integer

        'Vérifie qu'il y ait un élément de coché
45      For iCompteur = 1 To lvwMenage.ListItems.Count
50        If lvwMenage.ListItems(iCompteur).Checked = True Then
55          bChecked = True

60          Exit For
65        End If
70      Next

        'Si aucun n'est coché, on sort de la méthode
75      If bChecked = False Then
80        Call MsgBox("Aucune pièce n'a été sélectionnée!", vbOKOnly, "Erreur")

85        Exit Sub
90      End If

95      MousePointer = vbHourglass

'100     If OuvrirConnectionEquipment = True Then
          'Pour chaque élément du listview
105       For iCompteur = 1 To lvwMenage.ListItems.Count
110         Set itmMenage = lvwMenage.ListItems(iCompteur)

            'S'il est coché
115         If itmMenage.Checked = True Then
120           iIndexCheck = iCompteur
125         End If

            'On l'efface
130         Call g_connEquipment.Execute("DELETE * FROM [" & itmMenage.SubItems(I_COL_CATEGORIE) & "] WHERE PIECE = '" & Replace(itmMenage.Text, "'", "''") & "'")
135       Next

          'On ajoute l'élément choisi
140       Set itmMenage = lvwMenage.ListItems(iIndexCheck)

145       sCategorie = itmMenage.SubItems(I_COL_CATEGORIE)

150       Set rstPiece = New ADODB.Recordset

155       If sCategorie <> "CABLE" Then
160         Call rstPiece.Open("SELECT " & S_SELECT_NORMAL & " FROM [" & sCategorie & "]", g_connEquipment, adOpenDynamic, adLockOptimistic)
165       Else
170         Call rstPiece.Open("SELECT " & S_SELECT_CABLE & " FROM [" & sCategorie & "]", g_connEquipment, adOpenDynamic, adLockOptimistic)
175       End If

180       Call rstPiece.AddNew
        
185       rstPiece.Fields("PIECE") = itmMenage.Text
190       rstPiece.Fields("PIECE_GRB") = itmMenage.SubItems(I_COL_PIECE_GRB)

195       If itmMenage.SubItems(I_COL_FABRICANT) <> vbNullString Then
200         rstPiece.Fields("FABRICANT") = itmMenage.SubItems(I_COL_FABRICANT)
205       Else
210         rstPiece.Fields("FABRICANT") = " "
215       End If

220       If itmMenage.SubItems(I_COL_DESCR_FR) <> vbNullString Then
225         rstPiece.Fields("DESCR_FR") = itmMenage.SubItems(I_COL_DESCR_FR)
230       Else
235         rstPiece.Fields("DESCR_FR") = " "
240       End If

245       If itmMenage.SubItems(I_COL_DESCR_EN) <> vbNullString Then
250         rstPiece.Fields("DESCR_EN") = itmMenage.SubItems(I_COL_DESCR_EN)
255       Else
260         rstPiece.Fields("DESCR_EN") = " "
265       End If

270       If itmMenage.SubItems(I_COL_COMMENTAIRE) <> vbNullString Then
275         rstPiece.Fields("COMMENT") = itmMenage.SubItems(I_COL_COMMENTAIRE)
280       Else
285         rstPiece.Fields("COMMENT") = " "
290       End If

295       If itmMenage.SubItems(I_COL_PAGE_CAT) <> vbNullString Then
300         rstPiece.Fields("PAGE_CAT") = itmMenage.SubItems(I_COL_PAGE_CAT)
305       Else
310         rstPiece.Fields("PAGE_CAT") = " "
315      End If

320       If sCategorie <> "CABLE" Then
325         If itmMenage.SubItems(I_COL_EPAISSEUR) <> vbNullString Then
330           rstPiece.Fields("EPAISSEUR") = itmMenage.SubItems(I_COL_EPAISSEUR)
335         Else
340           rstPiece.Fields("EPAISSEUR") = "0"
345         End If

350         If itmMenage.SubItems(I_COL_HAUTEUR) <> vbNullString Then
355           rstPiece.Fields("HAUTEUR") = itmMenage.SubItems(I_COL_HAUTEUR)
360         Else
365           rstPiece.Fields("HAUTEUR") = "0"
370         End If

375         If itmMenage.SubItems(I_COL_LARGEUR) <> vbNullString Then
380           rstPiece.Fields("LARGEUR") = itmMenage.SubItems(I_COL_LARGEUR)
385         Else
390           rstPiece.Fields("LARGEUR") = "0"
395         End If
400       End If

405       rstPiece.Fields("TEMPS") = itmMenage.SubItems(I_COL_TEMPS)

410       If itmMenage.SubItems(I_COL_CSA) = "X" Then
415         rstPiece.Fields("CSA") = "1"
420       Else
425         rstPiece.Fields("CSA") = "0"
430       End If

435       If itmMenage.SubItems(I_COL_CUL) = "X" Then
440         rstPiece.Fields("CUL") = "1"
445       Else
450         rstPiece.Fields("CUL") = "0"
455       End If

460       If itmMenage.SubItems(I_COL_CE) = "X" Then
465         rstPiece.Fields("CE") = "1"
470       Else
475         rstPiece.Fields("CE") = "0"
480       End If

485       If itmMenage.SubItems(I_COL_UR) = "X" Then
490         rstPiece.Fields("UR") = "1"
495       Else
500         rstPiece.Fields("UR") = "0"
505       End If

510       If itmMenage.SubItems(I_COL_UL) = "X" Then
515         rstPiece.Fields("UL") = "1"
520       Else
525         rstPiece.Fields("UL") = "0"
530       End If

535       If sCategorie <> "CABLE" Then
540         If itmMenage.Tag <> "Null" Then
545           rstPiece.Fields("IMPLANT") = itmMenage.Tag
550         End If
  
555         If itmMenage.ListSubItems(I_COL_PIECE_GRB).Tag <> "Null" Then
560           rstPiece.Fields("VIGNETTE") = itmMenage.ListSubItems(I_COL_PIECE_GRB).Tag
565         End If
570       End If

575       If itmMenage.ListSubItems(I_COL_CATEGORIE).Tag <> "Null" Then
580         rstPiece.Fields("{SEE_SET}") = itmMenage.ListSubItems(I_COL_CATEGORIE).Tag
585       End If

590       Call rstPiece.Update

595       Call rstPiece.Close
600       Set rstPiece = Nothing

'605       Call FermerConnectionEquipment
'610     End If

        'Il faut modifier le champs TableElec dans GRB_PiecesFRS
615     Set rstPieceFRS = New ADODB.Recordset
        
620     rstPieceFRS.CursorLocation = adUseServer
        
625     Call rstPieceFRS.Open("SELECT TableElec FROM GRB_PiecesFRS WHERE PIECE = '" & Replace(itmMenage.Text, "'", "''") & "' AND Type = 'E'", g_connData, adOpenDynamic, adLockOptimistic)

630     Do While Not rstPieceFRS.EOF
635       rstPieceFRS.Fields("TableElec") = sCategorie

640       Call rstPieceFRS.Update

645       Call rstPieceFRS.MoveNext
650     Loop

655     Call rstPieceFRS.Close
660     Set rstPieceFRS = Nothing

        'On rempli le combo des fabricants pour remettre à jour la liste
665     Call FrmCatalogueElec.RemplirComboFabricant

670     MousePointer = vbDefault

675     FrmCatalogueElec.m_bPieceEffacée = True

680     Call Unload(Me)

685     Exit Sub

AfficherErreur:

690     Call AfficherErreur(Me, "cmdOK_Click", Err, Erl)
End Sub

Private Sub ViderChamps()
        'Vide les valeurs pour la modif
5       On Error GoTo AfficherErreur

10      txtPiece.Text = vbNullString
15      txtPieceGRB.Text = vbNullString
20      txtCategorie.Text = vbNullString
25      txtFabricant.Text = vbNullString
30      txtDescrFR.Text = vbNullString
35      txtDescrEN.Text = vbNullString
40      txtComment.Text = vbNullString
45      txtPageCat.Text = vbNullString
50      txtEpaisseur.Text = vbNullString
55      txtHauteur.Text = vbNullString
60      txtTemps.Text = vbNullString
65      chkCSA.Value = vbUnchecked
70      chkCUL.Value = vbUnchecked
75      chkCE.Value = vbUnchecked
80      chkUR.Value = vbUnchecked
85      chkUL.Value = vbUnchecked

90      Exit Sub

AfficherErreur:

95      Call AfficherErreur(Me, "ViderChamps", Err, Erl)
End Sub
