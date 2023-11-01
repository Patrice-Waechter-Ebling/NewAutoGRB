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

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer
  
15      For iCompteur = 0 To cmbCategorie.ListCount - 1
20        If UCase(cmbCategorie.LIST(iCompteur)) = UCase(cmbCategorie.Text) Then
25          cmbCategorie.ListIndex = iCompteur
      
30          Exit For
35        End If
40      Next

45      Exit Sub

AfficherErreur:

50      woups "FrmaddItemElec", "cmbProjSoum_KeyUp", Err, Erl
End Sub

Private Sub cmbFabricant_KeyPress(KeyAscii As Integer)

5       On Error GoTo AfficherErreur

10      If KeyAscii <= 122 And KeyAscii >= 97 Then
15        KeyAscii = KeyAscii - 32
20      End If

25      Exit Sub

AfficherErreur:

30      woups "FrmaddItemElec", "cmbFabricant_KeyPress", Err, Erl
End Sub

Private Sub cmdAnnuler_Click()

5       On Error GoTo AfficherErreur

        'Fermeture de la fenêtre
10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "FrmaddItemElec", "cmdAnnuler_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

        'Rempli le combo des catégories avec le nom des tables
10      Call RemplirComboCategorie
  
        'Sur l'ouverture, il faut remplir le combo des manufacturiers
15      Call RemplirComboManufacturier

20      Screen.MousePointer = vbDefault

25      Exit Sub

AfficherErreur:

30      woups "FrmaddItemElec", "Form_Load", Err, Erl
End Sub

Private Sub RemplirComboCategorie()

5       On Error GoTo AfficherErreur

        'Remplir le combo catégorie
10      Dim rstCatalogueElec As ADODB.Recordset
  
        'Il faut vider le combo avant de le remplir
15      Call cmbCategorie.Clear
  
20      Set rstCatalogueElec = New ADODB.Recordset
  
        'Cette méthode crée un recordset contenant les categorie
        'le nom de toutes les tables de la BD
25      Call rstCatalogueElec.Open("SELECT DISTINCT CATEGORIE FROM GRB_CatalogueElec ORDER BY Categorie", g_connData, adOpenDynamic, adLockOptimistic)
  
        'Tant que ce n'est pas la fin des enregistrements
30      Do While Not rstCatalogueElec.EOF
35        Call cmbCategorie.AddItem(rstCatalogueElec.Fields("CATEGORIE"))
    
40        Call rstCatalogueElec.MoveNext
45      Loop
  
50      Call rstCatalogueElec.Close
55      Set rstCatalogueElec = Nothing
    
        'Si le combo n'est pas vide, on sélectionne la catégorie sélectionnée dans
        'le catalogue
60      If cmbCategorie.ListCount > 0 Then
65        cmbCategorie.Text = FrmCatalogueElec.cmbCategorie.Text
70      End If

75      Exit Sub

AfficherErreur:

80      woups "FrmAddItemElec", "RemplirComboCategorie", Err, Erl
End Sub

Private Sub RemplirComboManufacturier()

5       On Error GoTo AfficherErreur

        'Rempli le combo des manufacturiers selon la table choisie
10      Dim rstManufacturier As ADODB.Recordset
  
15      Set rstManufacturier = New ADODB.Recordset
  
20      Call rstManufacturier.Open("SELECT DISTINCT FABRICANT FROM GRB_CatalogueElec", g_connData, adOpenDynamic, adLockOptimistic)

        'Tant que c'est pas la fin des enregistrements
25      Do While Not rstManufacturier.EOF
30        If Not IsNull(rstManufacturier.Fields("FABRICANT")) Then
            'Ajout du nom du manufacturier au Combo
35          Call cmbFabricant.AddItem(rstManufacturier.Fields("FABRICANT"))
40        End If
  
45        Call rstManufacturier.MoveNext
50      Loop

55      Call rstManufacturier.Close
60      Set rstManufacturier = Nothing
    
65      Exit Sub

AfficherErreur:

70      woups "FrmaddItemElec", "RemplirComboManufacturier", Err, Erl
End Sub

Private Sub cmdOK_Click()

5       On Error GoTo AfficherErreur

        'Proc qui permet d'ajouter un item a la BD
10      Dim rstItem     As ADODB.Recordset
15      Dim rstFRS      As ADODB.Recordset
20      Dim rstPieceFRS As ADODB.Recordset
25      Dim iCompteur   As Integer
30      Dim iFRS        As Integer
35      Dim sPieceModif As String
40      Dim sLettre     As String

        'Si aucun champs est vide
45      If Trim$(txtNoItem.Text) <> vbNullString And Trim$(cmbFabricant.Text) <> vbNullString And Trim$(cmbCategorie.Text) <> vbNullString Then
50        Screen.MousePointer = vbHourglass
          
55        Set rstItem = New ADODB.Recordset

60        Call rstItem.Open("SELECT * FROM GRB_CatalogueElec WHERE PIECE = '" & Replace(txtNoItem.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)

65        If Not rstItem.EOF Then
70          Call MsgBox("Attention! L'item # " & txtNoItem.Text & " existe déjà!", vbOKOnly, "Erreur")
            
75          Call rstItem.Close
80          Set rstItem = Nothing

85          Screen.MousePointer = vbDefault
90        Else
            'Si elle n'existe pas
            'On l'ajoute
95          Call rstItem.AddNew
          
100         rstItem.Fields("CATEGORIE").Value = cmbCategorie.Text
105         rstItem.Fields("FABRICANT").Value = Trim$(cmbFabricant.Text)
110         rstItem.Fields("PIECE").Value = Trim$(txtNoItem.Text)

115         For iCompteur = 1 To Len(Trim$(txtNoItem.Text))
120           sLettre = Mid$(Trim$(txtNoItem.Text), iCompteur, 1)
        
125           If (Asc(sLettre) >= 48 And Asc(sLettre) <= 57) Or _
                 (Asc(sLettre) >= 65 And Asc(sLettre) <= 90) Or _
                 (Asc(sLettre) >= 97 And Asc(sLettre) <= 122) Then
130             sPieceModif = sPieceModif & sLettre
135           End If
140         Next
      
145         rstItem.Fields("PIECE_MODIF") = sPieceModif
150         rstItem.Fields("PIECE_GRB").Value = Trim$(txtNoItem) & "GRB"
155         rstItem.Fields("TEMPS").Value = 0

160         Call rstItem.Update
         
165         Call rstItem.Close
170         Set rstItem = Nothing

175         Set rstFRS = New ADODB.Recordset

180         Call rstFRS.Open("SELECT * FROM GRB_Fournisseur WHERE NomFournisseur = 'FOURNI PAR LE CLIENT'", g_connData, adOpenDynamic, adLockOptimistic)

185         iFRS = rstFRS.Fields("IDFRS")

190         Call rstFRS.Close
195         Set rstFRS = Nothing

200         Set rstPieceFRS = New ADODB.Recordset

205         Call rstPieceFRS.Open("SELECT * FROM GRB_PiecesFRS", g_connData, adOpenDynamic, adLockOptimistic)

            'Ajout du fournisseur 'FOURNI PAR LE CLIENT'
210         Call rstPieceFRS.AddNew

215         rstPieceFRS.Fields("IDFRS") = iFRS
220         rstPieceFRS.Fields("PIECE") = Trim$(txtNoItem.Text)
225         rstPieceFRS.Fields("PRIX_LIST") = 0
230         rstPieceFRS.Fields("ESCOMPTE") = 0
235         rstPieceFRS.Fields("PRIX_NET") = 0
240         rstPieceFRS.Fields("DATE") = ConvertDate(Date)
245         rstPieceFRS.Fields("ENTRER_PAR") = g_sInitiale
250         rstPieceFRS.Fields("DeviseMonétaire") = "CAN"
255         rstPieceFRS.Fields("Type") = "E"

260         Call rstPieceFRS.Update

            'Ajout du fournisseur 'SOLUTION GRB INC.'
265         Call rstPieceFRS.AddNew

270         rstPieceFRS.Fields("IDFRS") = 717
275         rstPieceFRS.Fields("PIECE") = Trim$(txtNoItem.Text)
280         rstPieceFRS.Fields("PRIX_LIST") = 0
285         rstPieceFRS.Fields("ESCOMPTE") = 0
290         rstPieceFRS.Fields("PRIX_NET") = 0
295         rstPieceFRS.Fields("DATE") = ConvertDate(Date)
300         rstPieceFRS.Fields("ENTRER_PAR") = g_sInitiale
305         rstPieceFRS.Fields("DeviseMonétaire") = "CAN"
310         rstPieceFRS.Fields("Type") = "E"

315         Call rstPieceFRS.Update

320         Call rstPieceFRS.Close
325         Set rstPieceFRS = Nothing

335         FrmCatalogueElec.m_sSelectCategorie = cmbCategorie.Text
340         FrmCatalogueElec.m_sSelectFabricant = cmbFabricant.Text
345         FrmCatalogueElec.m_sSelectNoItem = txtNoItem.Text
        
            'Remplis le combo catégorie dans le form électrique
330         Call FrmCatalogueElec.RemplirComboCategorie
                
            'Montre seulement les boutons pour enregistrer
350         Call FrmCatalogueElec.MontrerControles(MODE_AJOUT_MODIF_ELEC)

355         FrmCatalogueElec.txtNoItemGRB.Text = txtNoItem.Text & "GRB"

360         FrmCatalogueElec.txtDescriptionFR.Text = vbNullString

365         Call FrmCatalogueElec.BarrerChamps_piece(False)

            'on redonne le controle au catalogue
370         Call Unload(Me)
375       End If
380     Else
385       Call MsgBox("Vous devez remplir tous les champs", vbOKOnly, "Erreur")
390     End If

395     Exit Sub

AfficherErreur:

400     woups "FrmaddItemElec", "cmdOK_Click", Err, Erl
End Sub
