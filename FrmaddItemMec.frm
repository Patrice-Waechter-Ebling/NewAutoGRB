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

50      woups "FrmaddItemMec", "cmbProjSoum_KeyUp", Err, Erl
End Sub

Private Sub RemplirComboCategorie()

5       On Error GoTo AfficherErreur

        'Remplir le combo categorie
10      Dim rstCatalogueMec As ADODB.Recordset
  
        'Il faut vider le combo avant de le remplir
15      Call cmbCategorie.Clear
  
20      Set rstCatalogueMec = New ADODB.Recordset
  
        'Cette méthode crée un recordset contenant les categorie
        'le nom de toutes les tables de la BD
25      Call rstCatalogueMec.Open("SELECT DISTINCT Categorie FROM GRB_CatalogueMec ORDER BY Categorie", g_connData, adOpenDynamic, adLockOptimistic)
  
        'Tant que ce n'est pas la fin des enregistrements
30      Do While Not rstCatalogueMec.EOF
35        Call cmbCategorie.AddItem(rstCatalogueMec.Fields("categorie"))
    
40        Call rstCatalogueMec.MoveNext
45      Loop
  
50      Call rstCatalogueMec.Close
55      Set rstCatalogueMec = Nothing
    
        'Si le combo n'est pas vide, on sélectionne la catégorie sélectionnée dans
        'le catalogue
60      If cmbCategorie.ListCount > 0 Then
65        cmbCategorie.Text = FrmCatalogueMec.cmbCategorie.Text
70      End If

75      Exit Sub

AfficherErreur:

80      woups "FrmaddItemMec", "RemplirComboCategorie", Err, Erl
End Sub

Private Sub cmbFabricant_KeyPress(KeyAscii As Integer)

5       On Error GoTo AfficherErreur
        
        'Pour mettre les lettres en majuscule
10      If KeyAscii <= 122 And KeyAscii >= 97 Then
15        KeyAscii = KeyAscii - 32
20      End If

25      Exit Sub

AfficherErreur:

30      woups "FrmaddItemMec", "cmbFabricant_KeyPress", Err, Erl
End Sub

Private Sub cmdAnnuler_Click()

5       On Error GoTo AfficherErreur

        'fermeture de la fenêtre
10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "FrmaddItemMec", "cmdAnnuler_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

        'Affiche les numéros suivants
        
10      Call AfficherNoAcier
        
15      Call AfficherNoPlastique
        
20      Call AfficherNoStainless
25
        Call AfficherNoAluminium
  
        'Sur l'ouverture, il faut remplir les combos
        
30      Call RemplirComboCategorie
35
        Call RemplirComboManufacturier
        
40      Screen.MousePointer = vbDefault

45      Exit Sub

AfficherErreur:

50      woups "FrmaddItemMec", "Form_Load", Err, Erl
End Sub

Private Sub AfficherNoAcier()
        'Affiche le numéro de la prochaine pièce ACIER
5       On Error GoTo AfficherErreur

10      Dim rstAcier As ADODB.Recordset
15      Dim sNoAcier As String
20      Dim iNoAcier As Integer
        Dim iNoAcierNow As Integer
25      Set rstAcier = New ADODB.Recordset
        iNoAcier = 0
30      Call rstAcier.Open("SELECT PIECE FROM GRB_CatalogueMec WHERE PIECE LIKE 'ACIER%' ORDER BY PIECE", g_connData, adOpenDynamic, adLockOptimistic)
35      Do While Not rstAcier.EOF
40          sNoAcier = Right(rstAcier.Fields("PIECE"), Len(rstAcier.Fields("PIECE")) - 5)
            sNoAcier = Left(sNoAcier, 4)
            If IsNumeric(sNoAcier) Then iNoAcierNow = CInt(sNoAcier)
45          If iNoAcierNow > iNoAcier Then iNoAcier = iNoAcierNow
            Call rstAcier.MoveNext
        Loop
        
        
  
55      lblAcier.Caption = "ACIER : " & iNoAcier + 1


75      Call rstAcier.Close
80      Set rstAcier = Nothing

85      Exit Sub

AfficherErreur:

90      woups "FrmaddItemMec", "AfficherNoAcier", Err, Erl
End Sub

Private Sub AfficherNoPlastique()
        'Affiche le numéro de la prochaine pièce ACIER
5       On Error GoTo AfficherErreur

10      Dim rstPlastique As ADODB.Recordset
15      Dim sNoPlastique As String
20      Dim iNoPlastique As Integer
        Dim iNoPlastiqueNow As Integer
25      Set rstPlastique = New ADODB.Recordset
  
30      Call rstPlastique.Open("SELECT PIECE FROM GRB_CatalogueMec WHERE PIECE LIKE 'PLASTIQUE%' ORDER BY PIECE", g_connData, adOpenDynamic, adLockOptimistic)
  
35      Do While Not rstPlastique.EOF
40
  
45        sNoPlastique = Right(rstPlastique.Fields("PIECE"), Len(rstPlastique.Fields("PIECE")) - 9)
          sNoPlastique = Left(sNoPlastique, 4)
          If IsNumeric(sNoPlastique) Then iNoPlastiqueNow = CInt(sNoPlastique)
          If iNoPlastiqueNow > iNoPlastique Then iNoPlastique = iNoPlastiqueNow
           Call rstPlastique.MoveNext
        Loop
          
55        lblPlastique.Caption = "PLASTIQUE : " & iNoPlastique + 1


75      Call rstPlastique.Close
80      Set rstPlastique = Nothing

85      Exit Sub

AfficherErreur:

90      woups "FrmaddItemMec", "AfficherNoPlastique", Err, Erl
End Sub

Private Sub AfficherNoStainless()
        'Affiche le numéro de la prochaine pièce ACIER
5       On Error GoTo AfficherErreur

10      Dim rstStainless As ADODB.Recordset
15      Dim sNoStainless As String
20      Dim iNoStainless As Integer
        Dim iNoStainlessNow As Integer
25      Set rstStainless = New ADODB.Recordset
  
30      Call rstStainless.Open("SELECT PIECE FROM GRB_CatalogueMec WHERE PIECE LIKE 'STAINLESS%' ORDER BY PIECE", g_connData, adOpenDynamic, adLockOptimistic)
  
35      Do While Not rstStainless.EOF
40          sNoStainless = Right(rstStainless.Fields("PIECE"), Len(rstStainless.Fields("PIECE")) - 9)
            If IsNumeric(sNoStainless) Then iNoStainlessNow = CInt(sNoStainless)
45          If iNoStainlessNow > iNoStainless Then iNoStainless = iNoStainlessNow
        Call rstStainless.MoveNext
        Loop
50
  
55        lblStainless.Caption = "STAINLESS : " & iNoStainless + 1
 

75      Call rstStainless.Close
80      Set rstStainless = Nothing

85      Exit Sub

AfficherErreur:

90      woups "FrmaddItemMec", "AfficherNoStainless", Err, Erl
End Sub

Private Sub AfficherNoAluminium()
        'Affiche le numéro de la prochaine pièce ACIER
5       On Error GoTo AfficherErreur

10      Dim rstAluminium As ADODB.Recordset
15      Dim sNoAluminium As String
20      Dim iNoAluminium As Integer
        Dim inoAluminiumNow As Integer
25      Set rstAluminium = New ADODB.Recordset
        iNoAluminium = 0
30      Call rstAluminium.Open("SELECT PIECE FROM GRB_CatalogueMec WHERE PIECE LIKE 'ALUMINIUM%' ORDER BY PIECE", g_connData, adOpenDynamic, adLockOptimistic)
  
35     Do While Not rstAluminium.EOF
40
  
45        sNoAluminium = Right(rstAluminium.Fields("PIECE"), Len(rstAluminium.Fields("PIECE")) - 9)
          sNoAluminium = Left(sNoAluminium, 4)
50        If IsNumeric(sNoAluminium) Then inoAluminiumNow = CInt(sNoAluminium)
        If inoAluminiumNow > iNoAluminium Then iNoAluminium = inoAluminiumNow
        Call rstAluminium.MoveNext
        Loop
55        lblAluminium.Caption = "ALUMINIUM : " & iNoAluminium + 1

75      Call rstAluminium.Close
80      Set rstAluminium = Nothing

85      Exit Sub

AfficherErreur:

90      woups "FrmaddItemMec", "AfficherNoAluminium", Err, Erl
End Sub

Private Sub RemplirComboManufacturier()

5       On Error GoTo AfficherErreur

        'Rempli le combo des manufacturiers selon la table choisie
10      Dim rstManufacturier As ADODB.Recordset
  
15      Set rstManufacturier = New ADODB.Recordset
  
20      Call rstManufacturier.Open("SELECT DISTINCT FABRICANT FROM GRB_CatalogueMec", g_connData, adOpenDynamic, adLockOptimistic)

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

70      woups "FrmaddItemMec", "RemplirComboManufacturier", Err, Erl
End Sub

Private Sub cmdOK_Click()

5       On Error GoTo AfficherErreur

        'Proc qui permet dajouter un item a la BD
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
      
60        Call rstItem.Open("SELECT * FROM GRB_CatalogueMec WHERE PIECE = '" & Replace(txtNoItem.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
      
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
105         rstItem.Fields("FABRICANT").Value = cmbFabricant.Text
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
150         rstItem.Fields("PIECE_GRB").Value = Trim$(txtNoItem.Text) & "GRB"
  
155         Call rstItem.Update
    
160         Call rstItem.Close
165         Set rstItem = Nothing

170         Set rstFRS = New ADODB.Recordset

175         Call rstFRS.Open("SELECT * FROM GRB_Fournisseur WHERE NomFournisseur = 'FOURNI PAR LE CLIENT'", g_connData, adOpenDynamic, adLockOptimistic)

180         iFRS = rstFRS.Fields("IDFRS")

185         Call rstFRS.Close
190         Set rstFRS = Nothing

195         Set rstPieceFRS = New ADODB.Recordset
  
200         Call rstPieceFRS.Open("SELECT * FROM GRB_PiecesFRS", g_connData, adOpenDynamic, adLockOptimistic)

            'Ajout du fournisseur 'FOURNI PAR LE CLIENT'
205         Call rstPieceFRS.AddNew

210         rstPieceFRS.Fields("IDFRS") = iFRS
215         rstPieceFRS.Fields("PIECE") = Trim$(txtNoItem.Text)
220         rstPieceFRS.Fields("PRIX_LIST") = 0
225         rstPieceFRS.Fields("ESCOMPTE") = 0
230         rstPieceFRS.Fields("PRIX_NET") = 0
235         rstPieceFRS.Fields("DATE") = ConvertDate(Date)
240         rstPieceFRS.Fields("ENTRER_PAR") = g_sInitiale
245         rstPieceFRS.Fields("DeviseMonétaire") = "CAN"
250         rstPieceFRS.Fields("Type") = "M"

255         Call rstPieceFRS.Update

            'Ajout du fournisseur 'SOLUTION GRB INC.'
260         Call rstPieceFRS.AddNew

265         rstPieceFRS.Fields("IDFRS") = 717
270         rstPieceFRS.Fields("PIECE") = Trim$(txtNoItem.Text)
275         rstPieceFRS.Fields("PRIX_LIST") = 0
280         rstPieceFRS.Fields("ESCOMPTE") = 0
285         rstPieceFRS.Fields("PRIX_NET") = 0
290         rstPieceFRS.Fields("DATE") = ConvertDate(Date)
295         rstPieceFRS.Fields("ENTRER_PAR") = g_sInitiale
300         rstPieceFRS.Fields("DeviseMonétaire") = "CAN"
305         rstPieceFRS.Fields("Type") = "M"

310         Call rstPieceFRS.Update

315         Call rstPieceFRS.Close
320         Set rstPieceFRS = Nothing

330         FrmCatalogueMec.m_sSelectCategorie = cmbCategorie.Text
335         FrmCatalogueMec.m_sSelectFabricant = cmbFabricant.Text
340         FrmCatalogueMec.m_sSelectNoItem = txtNoItem.Text
                       
            'Remplis le combo catégorie dans le form catalogue mécanique
325         Call FrmCatalogueMec.RemplirComboCategorie
                       
            'Montre seulement les boutons pour enregistrer
345         Call FrmCatalogueMec.MontrerControles(MODE_AJOUT_MODIF_MEC)

350         FrmCatalogueMec.txtNoItemGRB.Text = txtNoItem.Text & "GRB"

355         FrmCatalogueMec.txtDescriptionFR.Text = vbNullString

360         Call FrmCatalogueMec.BarrerChamps_piece(False)

            'On redonne le focus au catalogue
365         Call Unload(Me)
370       End If
375     Else
380       Call MsgBox("Vous devez remplir tous les champs", vbOKOnly, "Erreur")
385     End If

390     Exit Sub

AfficherErreur:

395     woups "FrmaddItemMec", "cmdOK_Click", Err, Erl
End Sub

