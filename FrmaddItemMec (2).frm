VERSION 5.00
Begin VB.Form FrmaddItemMec 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "."
   ClientHeight    =   3645
   ClientLeft      =   3570
   ClientTop       =   3240
   ClientWidth     =   5805
   ControlBox      =   0   'False
   Icon            =   "FrmaddItemMec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbcategorie 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   2160
      Width           =   3975
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnnuler 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Annuler"
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtNoItem 
      BackColor       =   &H00FFFFFF&
      Height          =   288
      Left            =   1680
      TabIndex        =   1
      Top             =   2640
      Width           =   2412
   End
   Begin VB.ComboBox cmbFabricant 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   3120
      Width           =   2412
   End
   Begin VB.Label lblStainless 
      BackStyle       =   0  'Transparent
      Caption         =   "STAINLESS : "
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
      Left            =   2400
      TabIndex        =   9
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label lblAluminium 
      BackStyle       =   0  'Transparent
      Caption         =   "ALUMINIUM : "
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
      Left            =   2400
      TabIndex        =   10
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblPlastique 
      BackStyle       =   0  'Transparent
      Caption         =   "PLASTIQUE : "
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
      Left            =   2400
      TabIndex        =   8
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblAcier 
      BackStyle       =   0  'Transparent
      Caption         =   "ACIER : "
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
      Left            =   2400
      TabIndex        =   7
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Prochain numéro pour :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Catégorie :"
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
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero d'item :"
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
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturier :"
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
      Left            =   240
      TabIndex        =   13
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label lblTitre 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmaddItemMec.frx":030A
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
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "FrmaddItemMec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbCategorie_KeyUp(KeyCode As Integer, Shift As Integer)

 On Error GoTo Oups

 Dim iCompteur As Integer
 
 For iCompteur = 0 To cmbCategorie.ListCount - 1
 If UCase(cmbCategorie.LIST(iCompteur)) = UCase(cmbCategorie.Text) Then
 cmbCategorie.ListIndex = iCompteur
 
 Exit For
 End If
 Next

 Exit Sub

Oups:

 wOups "FrmaddItemMec", "cmbProjSoum_KeyUp", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboCategorie()

 On Error GoTo Oups

 'Remplir le combo categorie
 Dim rstCatalogueMec As ADODB.Recordset
 
 'Il faut vider le combo avant de le remplir
 Call cmbCategorie.Clear
 
 Set rstCatalogueMec = New ADODB.Recordset
 
 'Cette méthode crée un recordset contenant les categorie
 'le nom de toutes les tables de la BD
 Call rstCatalogueMec.Open("SELECT DISTINCT Categorie FROM GrbCatalogueMec ORDER BY Categorie", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Tant que ce n'est pas la fin des enregistrements
 Do While Not rstCatalogueMec.EOF
 Call cmbCategorie.AddItem(rstCatalogueMec.Fields("categorie"))
 
 Call rstCatalogueMec.MoveNext
 Loop
 
 Call rstCatalogueMec.Close
 Set rstCatalogueMec = Nothing
 
 'Si le combo n'est pas vide, on sélectionne la catégorie sélectionnée dans
 'le catalogue
  If cmbCategorie.ListCount > 0 Then
  cmbCategorie.Text = FrmCatalogueMec.cmbCategorie.Text
  End If

  Exit Sub

Oups:

  wOups "FrmaddItemMec", "RemplirComboCategorie", Err, Err.number, Err.Description
End Sub

Private Sub cmbFabricant_KeyPress(KeyAscii As Integer)

 On Error GoTo Oups
 
 'Pour mettre les lettres en majuscule
 If KeyAscii <= 122 And KeyAscii >=   Then
 KeyAscii = KeyAscii - 32
 End If

 Exit Sub

Oups:

 wOups "FrmaddItemMec", "cmbFabricant_KeyPress", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnuler_Click()

 On Error GoTo Oups

 'fermeture de la fenêtre
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "FrmaddItemMec", "cmdAnnuler_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 'Affiche les numéros suivants
 
 Call AfficherNoAcier
 
 Call AfficherNoPlastique
 
 Call AfficherNoStainless
25
 Call AfficherNoAluminium
 
 'Sur l'ouverture, il faut remplir les combos
 
 Call RemplirComboCategorie
35
 Call RemplirComboManufacturier
 
 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "FrmaddItemMec", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub AfficherNoAcier()
 'Affiche le numéro de la prochaine pièce ACIER
 On Error GoTo Oups

 Dim rstAcier As ADODB.Recordset
 Dim sNoAcier As String
 Dim iNoAcier As Integer
 Dim iNoAcierNow As Integer
 Set rstAcier = New ADODB.Recordset
 iNoAcier = 0
 Call rstAcier.Open("SELECT PIECE FROM GrbCatalogueMec WHERE PIECE LIKE 'ACIER%' ORDER BY PIECE", g_connData, adOpenDynamic, adLockOptimistic)
 Do While Not rstAcier.EOF
 sNoAcier = Right(rstAcier.Fields("PIECE"), Len(rstAcier.Fields("PIECE")) - 5)
 sNoAcier = Left(sNoAcier, 4)
 If IsNumeric(sNoAcier) Then iNoAcierNow = CInt(sNoAcier)
 If iNoAcierNow > iNoAcier Then iNoAcier = iNoAcierNow
 Call rstAcier.MoveNext
 Loop
 
 
 
 lblAcier.Caption = "ACIER : " & iNoAcier + 1


  Call rstAcier.Close
  Set rstAcier = Nothing

  Exit Sub

Oups:

  wOups "FrmaddItemMec", "AfficherNoAcier", Err, Err.number, Err.Description
End Sub

Private Sub AfficherNoPlastique()
 'Affiche le numéro de la prochaine pièce ACIER
 On Error GoTo Oups

 Dim rstPlastique As ADODB.Recordset
 Dim sNoPlastique As String
 Dim iNoPlastique As Integer
 Dim iNoPlastiqueNow As Integer
 Set rstPlastique = New ADODB.Recordset
 
 Call rstPlastique.Open("SELECT PIECE FROM GrbCatalogueMec WHERE PIECE LIKE 'PLASTIQUE%' ORDER BY PIECE", g_connData, adOpenDynamic, adLockOptimistic)
 
 Do While Not rstPlastique.EOF
40
 
 sNoPlastique = Right(rstPlastique.Fields("PIECE"), Len(rstPlastique.Fields("PIECE")) - 9)
 sNoPlastique = Left(sNoPlastique, 4)
 If IsNumeric(sNoPlastique) Then iNoPlastiqueNow = CInt(sNoPlastique)
 If iNoPlastiqueNow > iNoPlastique Then iNoPlastique = iNoPlastiqueNow
 Call rstPlastique.MoveNext
 Loop
 
 lblPlastique.Caption = "PLASTIQUE : " & iNoPlastique + 1


  Call rstPlastique.Close
  Set rstPlastique = Nothing

  Exit Sub

Oups:

  wOups "FrmaddItemMec", "AfficherNoPlastique", Err, Err.number, Err.Description
End Sub

Private Sub AfficherNoStainless()
 'Affiche le numéro de la prochaine pièce ACIER
 On Error GoTo Oups

 Dim rstStainless As ADODB.Recordset
 Dim sNoStainless As String
 Dim iNoStainless As Integer
 Dim iNoStainlessNow As Integer
 Set rstStainless = New ADODB.Recordset
 
 Call rstStainless.Open("SELECT PIECE FROM GrbCatalogueMec WHERE PIECE LIKE 'STAINLESS%' ORDER BY PIECE", g_connData, adOpenDynamic, adLockOptimistic)
 
 Do While Not rstStainless.EOF
 sNoStainless = Right(rstStainless.Fields("PIECE"), Len(rstStainless.Fields("PIECE")) - 9)
 If IsNumeric(sNoStainless) Then iNoStainlessNow = CInt(sNoStainless)
 If iNoStainlessNow > iNoStainless Then iNoStainless = iNoStainlessNow
 Call rstStainless.MoveNext
 Loop
50
 
 lblStainless.Caption = "STAINLESS : " & iNoStainless + 1
 

  Call rstStainless.Close
  Set rstStainless = Nothing

  Exit Sub

Oups:

  wOups "FrmaddItemMec", "AfficherNoStainless", Err, Err.number, Err.Description
End Sub

Private Sub AfficherNoAluminium()
 'Affiche le numéro de la prochaine pièce ACIER
 On Error GoTo Oups

 Dim rstAluminium As ADODB.Recordset
 Dim sNoAluminium As String
 Dim iNoAluminium As Integer
 Dim inoAluminiumNow As Integer
 Set rstAluminium = New ADODB.Recordset
 iNoAluminium = 0
 Call rstAluminium.Open("SELECT PIECE FROM GrbCatalogueMec WHERE PIECE LIKE 'ALUMINIUM%' ORDER BY PIECE", g_connData, adOpenDynamic, adLockOptimistic)
 
 Do While Not rstAluminium.EOF
40
 
 sNoAluminium = Right(rstAluminium.Fields("PIECE"), Len(rstAluminium.Fields("PIECE")) - 9)
 sNoAluminium = Left(sNoAluminium, 4)
 If IsNumeric(sNoAluminium) Then inoAluminiumNow = CInt(sNoAluminium)
 If inoAluminiumNow > iNoAluminium Then iNoAluminium = inoAluminiumNow
 Call rstAluminium.MoveNext
 Loop
 lblAluminium.Caption = "ALUMINIUM : " & iNoAluminium + 1

  Call rstAluminium.Close
  Set rstAluminium = Nothing

  Exit Sub

Oups:

  wOups "FrmaddItemMec", "AfficherNoAluminium", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboManufacturier()

 On Error GoTo Oups

 'Rempli le combo des manufacturiers selon la table choisie
 Dim rstManufacturier As ADODB.Recordset
 
 Set rstManufacturier = New ADODB.Recordset
 
 Call rstManufacturier.Open("SELECT DISTINCT FABRICANT FROM GrbCatalogueMec", g_connData, adOpenDynamic, adLockOptimistic)

 'Tant que c'est pas la fin des enregistrements
 Do While Not rstManufacturier.EOF
 If Not IsNull(rstManufacturier.Fields("FABRICANT")) Then
 'Ajout du nom du manufacturier au Combo
 Call cmbFabricant.AddItem(rstManufacturier.Fields("FABRICANT"))
 End If
 
 Call rstManufacturier.MoveNext
 Loop

 Call rstManufacturier.Close
  Set rstManufacturier = Nothing

  Exit Sub

Oups:

  wOups "FrmaddItemMec", "RemplirComboManufacturier", Err, Err.number, Err.Description
End Sub

Private Sub cmdOK_Click()

 On Error GoTo Oups

 'Proc qui permet dajouter un item a la BD
 Dim rstItem As ADODB.Recordset
 Dim rstFRS As ADODB.Recordset
 Dim rstPieceFRS As ADODB.Recordset
 Dim iCompteur As Integer
 Dim iFRS As Integer
 Dim sPieceModif As String
 Dim sLettre As String
 
 'Si aucun champs est vide
 If Trim$(txtNoItem.Text) <> vbNullString And Trim$(cmbFabricant.Text) <> vbNullString And Trim$(cmbCategorie.Text) <> vbNullString Then
 Screen.MousePointer = vbHourglass

 Set rstItem = New ADODB.Recordset
 
  Call rstItem.Open("SELECT * FROM GrbCatalogueMec WHERE PIECE = '" & Replace(txtNoItem.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
  If Not rstItem.EOF Then
  Call MsgBox("Attention! L'item # " & txtNoItem.Text & " existe déjà!", vbOKOnly, "Erreur")

  Call rstItem.Close
  Set rstItem = Nothing
 
  Screen.MousePointer = vbDefault
  Else
 'Si elle n'existe pas
 'On l'ajoute
  Call rstItem.AddNew
 
 rstItem.Fields("CATEGORIE").Value = cmbCategorie.Text
rstItem.Fields("FABRICANT").Value = cmbFabricant.Text
 rstItem.Fields("PIECE").Value = Trim$(txtNoItem.Text)

 For iCompteur = 1 To Len(Trim$(txtNoItem.Text))
 sLettre = Mid$(Trim$(txtNoItem.Text), iCompteur, 1)
 
 If (Asc(sLettre) >= 4 And Asc(sLettre) <= 57) Or _
 (Asc(sLettre) >= 65 And Asc(sLettre) <= 90) Or _
 (Asc(sLettre) >=   And Asc(sLettre) <= 122) Then
 sPieceModif = sPieceModif & sLettre
 End If
 Next
 
 rstItem.Fields("PIECE_MODIF") = sPieceModif
 rstItem.Fields("PIECE_GRB").Value = Trim$(txtNoItem.Text) & "GRB"
 
 Call rstItem.Update
 
 Call rstItem.Close
 Set rstItem = Nothing

 Set rstFRS = New ADODB.Recordset

 Call rstFRS.Open("SELECT * FROM GrbFournisseur WHERE NomFournisseur = 'FOURNI PAR LE CLIENT'", g_connData, adOpenDynamic, adLockOptimistic)

 iFRS = rstFRS.Fields("IDFRS")

 Call rstFRS.Close
 Set rstFRS = Nothing

1  Set rstPieceFRS = New ADODB.Recordset
 
 Call rstPieceFRS.Open("SELECT * FROM GrbPiecesFRS", g_connData, adOpenDynamic, adLockOptimistic)

 'Ajout du fournisseur 'FOURNI PAR LE CLIENT'
 Call rstPieceFRS.AddNew

 rstPieceFRS.Fields("IDFRS") = iFRS
 rstPieceFRS.Fields("PIECE") = Trim$(txtNoItem.Text)
 rstPieceFRS.Fields("PRIX_LIST") = 0
 rstPieceFRS.Fields("ESCOMPTE") = 0
 rstPieceFRS.Fields("PRIX_NET") = 0
 rstPieceFRS.Fields("DATE") = ConvertDate(Date)
 rstPieceFRS.Fields("ENTRER_PAR") = g_sInitiale
 rstPieceFRS.Fields("DeviseMonétaire") = "CAN"
 rstPieceFRS.Fields("Type") = "M"

 Call rstPieceFRS.Update

 'Ajout du fournisseur 'SOLUTION GRB INC.'
 Call rstPieceFRS.AddNew

 rstPieceFRS.Fields("IDFRS") = 717
 rstPieceFRS.Fields("PIECE") = Trim$(txtNoItem.Text)
 rstPieceFRS.Fields("PRIX_LIST") = 0
 rstPieceFRS.Fields("ESCOMPTE") = 0
 rstPieceFRS.Fields("PRIX_NET") = 0
 rstPieceFRS.Fields("DATE") = ConvertDate(Date)
 rstPieceFRS.Fields("ENTRER_PAR") = g_sInitiale
 rstPieceFRS.Fields("DeviseMonétaire") = "CAN"
rstPieceFRS.Fields("Type") = "M"

 Call rstPieceFRS.Update

 Call rstPieceFRS.Close
 Set rstPieceFRS = Nothing

 FrmCatalogueMec.m_sSelectCategorie = cmbCategorie.Text
 FrmCatalogueMec.m_sSelectFabricant = cmbFabricant.Text
 FrmCatalogueMec.m_sSelectNoItem = txtNoItem.Text
 
 'Remplis le combo catégorie dans le form catalogue mécanique
 Call FrmCatalogueMec.RemplirComboCategorie
 
 'Montre seulement les boutons pour enregistrer
 Call FrmCatalogueMec.MontrerControles(MODE_AJOUT_MODIF_MEC)

 FrmCatalogueMec.txtNoItemGRB.Text = txtNoItem.Text & "GRB"

 FrmCatalogueMec.txtDescriptionFR.Text = vbNullString

 Call FrmCatalogueMec.BarrerChamps_piece(False)

 'On redonne le focus au catalogue
 Call Unload(Me)
End If
Else
Call MsgBox("Vous devez remplir tous les champs", vbOKOnly, "Erreur")
End If

3  Exit Sub

Oups:

 wOups "FrmaddItemMec", "cmdOK_Click", Err, Err.number, Err.Description
End Sub

