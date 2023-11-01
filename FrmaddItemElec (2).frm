VERSION 5.00
Begin VB.Form FrmaddItemElec 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ajout d'items"
   ClientHeight    =   2730
   ClientLeft      =   3570
   ClientTop       =   3240
   ClientWidth     =   5790
   ControlBox      =   0   'False
   Icon            =   "FrmaddItemElec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbCategorie 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   2412
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnnuler 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Annuler"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtNoItem 
      BackColor       =   &H00FFFFFF&
      Height          =   288
      Left            =   1560
      MaxLength       =   37
      TabIndex        =   1
      Top             =   1680
      Width           =   2412
   End
   Begin VB.ComboBox cmbFabricant 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   2160
      Width           =   2412
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Catégorie"
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
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero d'item:"
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
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturier"
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
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmaddItemElec.frx":0442
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
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "FrmaddItemElec"
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

 wOups "FrmaddItemElec", "cmbProjSoum_KeyUp", Err, Err.number, Err.Description
End Sub

Private Sub cmbFabricant_KeyPress(KeyAscii As Integer)

 On Error GoTo Oups

 If KeyAscii <= 122 And KeyAscii >=   Then
 KeyAscii = KeyAscii - 32
 End If

 Exit Sub

Oups:

 wOups "FrmaddItemElec", "cmbFabricant_KeyPress", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnuler_Click()

 On Error GoTo Oups

 'Fermeture de la fenêtre
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "FrmaddItemElec", "cmdAnnuler_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 'Rempli le combo des catégories avec le nom des tables
 Call RemplirComboCategorie
 
 'Sur l'ouverture, il faut remplir le combo des manufacturiers
 Call RemplirComboManufacturier

 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "FrmaddItemElec", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboCategorie()

 On Error GoTo Oups

 'Remplir le combo catégorie
 Dim rstCatalogueElec As ADODB.Recordset
 
 'Il faut vider le combo avant de le remplir
 Call cmbCategorie.Clear
 
 Set rstCatalogueElec = New ADODB.Recordset
 
 'Cette méthode crée un recordset contenant les categorie
 'le nom de toutes les tables de la BD
 Call rstCatalogueElec.Open("SELECT DISTINCT CATEGORIE FROM GrbCatalogueElec ORDER BY Categorie", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Tant que ce n'est pas la fin des enregistrements
 Do While Not rstCatalogueElec.EOF
 Call cmbCategorie.AddItem(rstCatalogueElec.Fields("CATEGORIE"))
 
 Call rstCatalogueElec.MoveNext
 Loop
 
 Call rstCatalogueElec.Close
 Set rstCatalogueElec = Nothing
 
 'Si le combo n'est pas vide, on sélectionne la catégorie sélectionnée dans
 'le catalogue
 If cmbCategorie.ListCount > 0 Then
 cmbCategorie.Text = FrmCatalogueElec.cmbCategorie.Text
 End If

 Exit Sub

Oups:

 wOups "FrmAddItemElec", "RemplirComboCategorie", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboManufacturier()

 On Error GoTo Oups

 'Rempli le combo des manufacturiers selon la table choisie
 Dim rstManufacturier As ADODB.Recordset
 
 Set rstManufacturier = New ADODB.Recordset
 
 Call rstManufacturier.Open("SELECT DISTINCT FABRICANT FROM GrbCatalogueElec", g_connData, adOpenDynamic, adLockOptimistic)

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

 wOups "FrmaddItemElec", "RemplirComboManufacturier", Err, Err.number, Err.Description
End Sub

Private Sub cmdOK_Click()

 On Error GoTo Oups

 'Proc qui permet d'ajouter un item a la BD
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

  Call rstItem.Open("SELECT * FROM GrbCatalogueElec WHERE PIECE = '" & Replace(txtNoItem.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)

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
rstItem.Fields("FABRICANT").Value = Trim$(cmbFabricant.Text)
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
 rstItem.Fields("PIECE_GRB").Value = Trim$(txtNoItem) & "GRB"
 rstItem.Fields("TEMPS").Value = 0

 Call rstItem.Update
 
 Call rstItem.Close
 Set rstItem = Nothing

 Set rstFRS = New ADODB.Recordset

 Call rstFRS.Open("SELECT * FROM GrbFournisseur WHERE NomFournisseur = 'FOURNI PAR LE CLIENT'", g_connData, adOpenDynamic, adLockOptimistic)

 iFRS = rstFRS.Fields("IDFRS")

 Call rstFRS.Close
 Set rstFRS = Nothing

 Set rstPieceFRS = New ADODB.Recordset

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
 rstPieceFRS.Fields("Type") = "E"

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
 rstPieceFRS.Fields("Type") = "E"

 Call rstPieceFRS.Update

 Call rstPieceFRS.Close
 Set rstPieceFRS = Nothing

 FrmCatalogueElec.m_sSelectCategorie = cmbCategorie.Text
 FrmCatalogueElec.m_sSelectFabricant = cmbFabricant.Text
 FrmCatalogueElec.m_sSelectNoItem = txtNoItem.Text
 
 'Remplis le combo catégorie dans le form électrique
 Call FrmCatalogueElec.RemplirComboCategorie
 
 'Montre seulement les boutons pour enregistrer
 Call FrmCatalogueElec.MontrerControles(MODE_AJOUT_MODIF_ELEC)

 FrmCatalogueElec.txtNoItemGRB.Text = txtNoItem.Text & "GRB"

 FrmCatalogueElec.txtDescriptionFR.Text = vbNullString

 Call FrmCatalogueElec.BarrerChamps_piece(False)

 'on redonne le controle au catalogue
 Call Unload(Me)
 End If
 Else
 Call MsgBox("Vous devez remplir tous les champs", vbOKOnly, "Erreur")
 End If

 Exit Sub

Oups:

 wOups "FrmaddItemElec", "cmdOK_Click", Err, Err.number, Err.Description
End Sub
