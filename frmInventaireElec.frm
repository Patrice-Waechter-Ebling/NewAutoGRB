VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInventaireElec 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventaire électrique"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9990
   Icon            =   "frmInventaireElec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmInventaireElec.frx":0442
   ScaleHeight     =   6615
   ScaleWidth      =   9990
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExport 
      Caption         =   "Exporter vers Excel"
      Height          =   375
      Left            =   1560
      TabIndex        =   18
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton cmdAfficherHistorique 
      Caption         =   "Afficher"
      Default         =   -1  'True
      Height          =   375
      Left            =   9000
      TabIndex        =   14
      Top             =   1320
      Width           =   855
   End
   Begin VB.CheckBox chkChoix 
      BackColor       =   &H00000000&
      Caption         =   "Pièce"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.CheckBox chkChoix 
      BackColor       =   &H00000000&
      Caption         =   "Projet"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.CheckBox chkChoix 
      BackColor       =   &H00000000&
      Caption         =   "Date"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   6120
      TabIndex        =   7
      Top             =   840
      Width           =   660
   End
   Begin VB.CommandButton cmdImprimer 
      Caption         =   "Imprimer"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   8760
      TabIndex        =   17
      Top             =   6120
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwHistorique 
      Height          =   4215
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   7435
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   1667
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Utilisateur"
         Object.Width           =   1561
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Projet"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Pièce"
         Object.Width           =   2566
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Quantité"
         Object.Width           =   1482
      EndProperty
   End
   Begin VB.Frame fraDate 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   6000
      TabIndex        =   8
      Top             =   840
      Width           =   2895
      Begin MSMask.MaskEdBox mskDateDebut 
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDateFin 
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   540
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "AA-MM-JJ"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Du :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Au :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   540
         Width           =   375
      End
   End
   Begin VB.Frame fraPiece 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   2880
      TabIndex        =   5
      Top             =   840
      Width           =   3015
      Begin VB.TextBox txtNoPiece 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.Frame fraProjet 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2655
      Begin VB.TextBox txtNoProjet 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Label lblTitre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Historique de l'inventaire"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   240
      Width           =   6735
   End
End
Attribute VB_Name = "frmInventaireElec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Index des colonnes de lvwHistorique
Private Const I_HIST_DATE           As Integer = 0
Private Const I_HIST_USER           As Integer = 1
Private Const I_HIST_PROJET         As Integer = 2
Private Const I_HIST_PIECE          As Integer = 3
Private Const I_HIST_QUANTITE       As Integer = 4

'Index des chkChoix
Private Const I_CHK_CHOIX_PROJET    As Integer = 0
Private Const I_CHK_CHOIX_PIECE     As Integer = 1
Private Const I_CHK_CHOIX_DATE      As Integer = 2

'Pour l'impression
Public m_bAnnuleImpression As Boolean
Public m_eTypeImpression   As enumImpressionInventaire
Public m_typeImpressionExel As Boolean

Private Sub RemplirListViewHistorique()

5       On Error GoTo AfficherErreur

        'Rempli le ListView lvwHistorique
10      Dim rstHist As ADODB.Recordset
15      Dim sWhere  As String
20      Dim itmHist As ListItem
     
        'Si c'est une recherche avec le no de projet
25      If chkChoix(I_CHK_CHOIX_PROJET).Value = vbChecked Then
30        sWhere = "Left(IDProjet, " & Len(txtnoprojet.Text) & ") = '" & txtnoprojet.Text & "' "
35      End If
  
        'Si c'est une recherche avec le no de la piece
40      If chkChoix(I_CHK_CHOIX_PIECE).Value = vbChecked Then
45        If sWhere = vbNullString Then
50          sWhere = "INSTR(1,NoItem, '" & Replace(txtNoPiece.Text, "'", "''") & "') > 0 "
55        Else
60          sWhere = sWhere & " AND INSTR(1,NoItem,'" & Replace(txtNoPiece.Text, "'", "''") & "') > 0 "
65        End If
70      End If
  
        'Si c'est une recherche par date
75      If chkChoix(I_CHK_CHOIX_DATE).Value = vbChecked Then
80        If sWhere = vbNullString Then
85          sWhere = "Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "'"
90        Else
95          sWhere = sWhere & " AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "'"
100       End If
105     End If
                    
110     Set rstHist = New ADODB.Recordset
          
115     If sWhere <> "" Then
120       Call rstHist.Open("SELECT * FROM GRB_InventaireElecModif WHERE " & sWhere & " ORDER BY NoEnreg DESC", g_connData, adOpenDynamic, adLockOptimistic)
125     Else
130       Call rstHist.Open("SELECT * FROM GRB_InventaireElecModif ORDER BY NoEnreg DESC", g_connData, adOpenDynamic, adLockOptimistic)
135     End If
  
        'Il faut vider le ListView avant de le remplir
140     Call lvwHistorique.ListItems.Clear
  
        'Tant que ce n'est pas la fin des enregistrements
145     Do While Not rstHist.EOF
         'On l'ajoute
150       Set itmHist = lvwHistorique.ListItems.Add
    
          'Date
155       itmHist.Text = rstHist.Fields("Date")
160       itmHist.Tag = rstHist.Fields("NoEnreg")
    
          'User
165       itmHist.SubItems(I_HIST_USER) = rstHist.Fields("User")
    
          'IDProjet
170       If Not IsNull(rstHist.Fields("IDProjet")) Then
175         If Trim(rstHist.Fields("IDProjet")) <> "" Then
180           itmHist.SubItems(I_HIST_PROJET) = rstHist.Fields("IDProjet")
185         Else
190           itmHist.SubItems(I_HIST_PROJET) = "N/A"
195         End If
200       Else
205         itmHist.SubItems(I_HIST_PROJET) = "N/A"
210       End If
    
          'No Item
215       itmHist.SubItems(I_HIST_PIECE) = rstHist.Fields("NoItem")
    
          'Quantité
220       itmHist.SubItems(I_HIST_QUANTITE) = rstHist.Fields("Quantité")
    
225       Call rstHist.MoveNext
230     Loop
  
235     Call rstHist.Close
240     Set rstHist = Nothing

245     Exit Sub

AfficherErreur:

250     woups "frmInventaireElec", "RemplirListViewHistorique", Err, Erl
End Sub

Private Sub chkChoix_Click(Index As Integer)

5       On Error GoTo AfficherErreur

        'Méthode qui met enabled ou disabled les controles selon le type de recherche
10      Dim bEnabled As Boolean
    
15      If chkChoix(Index).Value = vbChecked Then
20        bEnabled = True
25      Else
30        bEnabled = False
35      End If
  
40      Select Case Index
          Case I_CHK_CHOIX_PROJET:
45          fraProjet.Enabled = bEnabled
      
          Case I_CHK_CHOIX_PIECE:
50          fraPiece.Enabled = bEnabled

          Case I_CHK_CHOIX_DATE:
55          fraDate.Enabled = bEnabled
60      End Select

65      Exit Sub

AfficherErreur:

70      woups "frmInventaireElec", "chkChoix_Click", Err, Erl
End Sub

Private Sub cmdAfficherHistorique_Click()

5       On Error GoTo AfficherErreur

        'Affiche l'historique des modifications de l'inventaire

10      If chkChoix(I_CHK_CHOIX_DATE).Value = vbChecked Then
15        If mskDateDebut.Text <> vbNullString And mskDateFin.Text <> vbNullString Then
20          If mskDateDebut.Text > mskDateFin.Text Then
25            Call MsgBox("La date de fin doit être plus petite que la date de début!", vbOKOnly, "Erreur")
        
30            Exit Sub
35          End If
40        Else
45          Call MsgBox("Vous devez remplir tous les champs!", vbOKOnly, "Erreur")
     
50          Exit Sub
55        End If
60      End If
  
65      Screen.MousePointer = vbHourglass
  
70      If chkChoix(I_CHK_CHOIX_DATE).Value = vbChecked Or _
           chkChoix(I_CHK_CHOIX_PIECE).Value = vbChecked Or _
           chkChoix(I_CHK_CHOIX_PROJET).Value = vbChecked Then
75        Call RemplirListViewHistorique
80      Else
85        Call MsgBox("Vous devez choisir au moins une option de recherche!", vbOKOnly, "Erreur")
90      End If
  
95      Screen.MousePointer = vbDefault

100     Exit Sub

AfficherErreur:

105     woups "frmInventaireElec", "cmdAfficherHistorique_Click", Err, Erl
End Sub

Private Sub cmdExport_Click()
m_typeImpressionExel = True 'pour l'affichage de Imprimer au lieu de Exporter dans frmchoixImpressionInventaire
      On Error GoTo AfficherErreur

      Call frmChoixImpressionInventaire.Afficher(Me)

    If m_bAnnuleImpression = False Then
        If m_eTypeImpression = MODE_AJUST_INV Then
            Call ExporterAjustementInventaire
        Else
            Call ExporterValeurComptable
        End If
    End If

Exit Sub

AfficherErreur:

    woups "frmInventaireElec", "cmdExport_Click", Err, Erl
End Sub

Private Sub cmdImprimer_Click()
m_typeImpressionExel = False 'pour l'affichage de Imprimer au lieu de Exporter dans frmchoixImpressionInventaire
5       On Error GoTo AfficherErreur

10      Call frmChoixImpressionInventaire.Afficher(Me)

15      If m_bAnnuleImpression = False Then
20        If m_eTypeImpression = MODE_AJUST_INV Then
25          Call ImprimerAjustementInventaire
30        Else
35          Call ImprimerValeurComptable
40        End If
45      End If

50      Exit Sub

AfficherErreur:

55      woups "frmInventaireElec", "cmdImprimer_Click", Err, Erl
End Sub

Private Sub Cmdfermer_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmInventaireElec", "cmdFermer_Click", Err, Erl
End Sub



Private Sub Form_Load()
  
5       On Error GoTo AfficherErreur

10      Call Unload(frmChoixInventaire)

15      Exit Sub

AfficherErreur:

20      woups "frmInventaireElec", "Form_Load", Err, Erl
End Sub

Private Sub lvwHistorique_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

5       On Error GoTo AfficherErreur

10      If lvwHistorique.ListItems.count > 0 Then
15        lvwHistorique.Sorted = True

20        lvwHistorique.SortKey = ColumnHeader.Index - 1
  
25        If lvwHistorique.SortOrder = lvwAscending Then
30          lvwHistorique.SortOrder = lvwDescending
35        Else
40          lvwHistorique.SortOrder = lvwAscending
45        End If
50      End If

55      Exit Sub

AfficherErreur:

60      woups "frmInventaireElec", "lvwHistorique_ColumnClick", Err, Erl
End Sub

Private Sub lvwHistorique_KeyDown(KeyCode As Integer, Shift As Integer)

5       On Error GoTo AfficherErreur

10      If KeyCode = vbKeyDelete Then
15        If MsgBox("Voulez-vous vraiment effacer cet historique?", vbYesNo) = vbYes Then
20          Call g_connData.Execute("DELETE * FROM GRB_InventaireElecModif WHERE NoEnreg = " & lvwHistorique.SelectedItem.Tag)

25          If MsgBox("Voulez-vous modifier la quantité dans l'inventaire?", vbYesNo) = vbYes Then
30            Call ModifierInventaire(lvwHistorique.SelectedItem.SubItems(I_HIST_PIECE), lvwHistorique.SelectedItem.SubItems(I_HIST_QUANTITE))
35          End If
40        End If

45        Call RemplirListViewHistorique
50      End If

55      Exit Sub

AfficherErreur:

60      woups "frmInventaireElec", "lvwHistorique_KeyDown", Err, Erl
End Sub

Private Sub ModifierInventaire(ByVal sPiece As String, ByVal sQuantite As String)

5       On Error GoTo AfficherErreur

10      Dim rstInv As ADODB.Recordset

15      If sQuantite > 0 Then
20        sQuantite = "-" & sQuantite
25      Else
30        sQuantite = Replace(sQuantite, "-", "")
35      End If

40      Set rstInv = New ADODB.Recordset

45      Call rstInv.Open("SELECT * FROM GRB_InventaireElec WHERE NoItem = '" & Replace(sPiece, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)

50      rstInv.Fields("QuantitéStock") = CDbl(rstInv.Fields("QuantitéStock")) + CDbl(sQuantite)

55      Call rstInv.Update

60      Call rstInv.Close
65      Set rstInv = Nothing

70      Exit Sub

AfficherErreur:

75      woups "frmInventaireElec", "ModifierInventaire", Err, Erl
End Sub

Private Sub mskDateDebut_GotFocus()

5       On Error GoTo AfficherErreur

10      If Len(mskDateDebut.Text) = 10 Then
15        mskDateDebut.Text = Right$(mskDateDebut.Text, 8)
20      End If
  
25      mskDateDebut.mask = "##-##-##"

30      Exit Sub

AfficherErreur:

35      woups "frmInventaireElec", "mskDateDebut_GotFocus", Err, Erl
End Sub

Private Sub mskDateDebut_LostFocus()

5       On Error GoTo AfficherErreur

10      mskDateDebut.mask = vbNullString
  
15      If mskDateDebut.Text = "__-__-__" Then
20        mskDateDebut.Text = vbNullString
25      Else
30        If Len(mskDateDebut.Text) = 8 Then
35          If IsDate(mskDateDebut.Text) Then
40            mskDateDebut.Text = Year(DateSerial(Left$(mskDateDebut.Text, 2), Mid$(mskDateDebut.Text, 4, 2), Right$(mskDateDebut.Text, 2))) & Mid$(mskDateDebut.Text, 3, 8)
45          End If
50        End If
55      End If

60      Exit Sub

AfficherErreur:

65      woups "frmInventaireElec", "mskDateDebut_LostFocus", Err, Erl
End Sub

Private Sub mskDateFin_GotFocus()

5       On Error GoTo AfficherErreur

10      If Len(mskDateFin.Text) = 10 Then
15        mskDateFin.Text = Right$(mskDateFin.Text, 8)
20      End If
  
25      mskDateFin.mask = "##-##-##"

30      Exit Sub

AfficherErreur:

35      woups "frmInventaireElec", "mskDateFin_GotFocus", Err, Erl
End Sub

Private Sub mskDateFin_LostFocus()

5       On Error GoTo AfficherErreur

10      mskDateFin.mask = vbNullString
  
15      If mskDateFin.Text = "__-__-__" Then
20        mskDateFin.Text = vbNullString
25      Else
30        If Len(mskDateFin.Text) = 8 Then
35          If IsDate(mskDateFin.Text) Then
40            mskDateFin.Text = Year(DateSerial(Left$(mskDateFin.Text, 2), Mid$(mskDateFin.Text, 4, 2), Right$(mskDateFin.Text, 2))) & Mid$(mskDateFin.Text, 3, 8)
45          End If
50        End If
55      End If

60      Exit Sub

AfficherErreur:

65      woups "frmInventaireElec", "mskDateFin_LostFocus", Err, Erl
End Sub

Private Sub ImprimerAjustementInventaire()

5       On Error GoTo AfficherErreur

        'Impression de l'inventaire trié par localisation et par no d'item
10      Dim rstInv  As ADODB.Recordset
15      Dim sChamps As String
  
20      Set rstInv = New ADODB.Recordset
  
25      Call rstInv.Open("SELECT NoItem, QuantitéStock, Description, Manufacturier, Localisation FROM GRB_InventaireElec ORDER BY Localisation, NoItem", g_connData, adOpenDynamic, adLockOptimistic)
        
30      Set DR_ListeInventaire.DataSource = rstInv
    
35      DR_ListeInventaire.Sections("Section3").Controls("lblDate").Caption = ConvertDate(Date)
40      DR_ListeInventaire.Sections("Section4").Controls("lblTitre").Caption = "Inventaire Électrique"
  
45      Call DR_ListeInventaire.Show(vbModal)
  
50      Call rstInv.Close
55      Set rstInv = Nothing

60      Exit Sub

AfficherErreur:

65      woups "frmInventaireElec", "cmdImprimer_Click", Err, Erl
End Sub

Private Sub ImprimerValeurComptable()

5       On Error GoTo AfficherErreur

        'Impression de la l'inventaire trié par localisation et par no d'item
10      Dim rstInv   As ADODB.Recordset
15      Dim rstTotal As ADODB.Recordset
        
20      Set rstInv = New ADODB.Recordset
25      Set rstTotal = New ADODB.Recordset
        
30      Call rstInv.Open("SELECT NoItem, Description, Manufacturier, [Prix Liste], Escompte, [Prix Net], Localisation, QuantitéStock, [Prix Net] * QuantitéStock As Total FROM GRB_InventaireElec WHERE QuantitéStock <> '0' ORDER BY NoItem", g_connData, adOpenDynamic, adLockOptimistic)
        
35      Set DR_Inventaire.DataSource = rstInv
    
40      Call rstTotal.Open("SELECT SUM([Prix net] * QuantitéStock) As Total FROM GRB_InventaireElec", g_connData, adOpenDynamic, adLockOptimistic)
    
45      DR_Inventaire.Sections("Section3").Controls("lblDate").Caption = ConvertDate(Date)
50      DR_Inventaire.Sections("Section5").Controls("lblTotal").Caption = Conversion(rstTotal.Fields("Total"), MODE_ARGENT)
55      DR_Inventaire.Sections("Section4").Controls("lblTitre").Caption = "Inventaire Électrique"
  
60      Call rstTotal.Close
65      Set rstTotal = Nothing

70      DR_Inventaire.Orientation = rptOrientLandscape
  
75      Call DR_Inventaire.Show(vbModal)
  
80      Call rstInv.Close
85      Set rstInv = Nothing

90      Exit Sub

AfficherErreur:

95      woups "frmInventaireElec", "cmdImprimer_Click", Err, Erl
End Sub
Private Function ExporterAjustementInventaire()
      On Error GoTo AfficherErreur

        'Impression de l'inventaire trié par localisation et par no d'item
        Dim rstInv  As ADODB.Recordset
        Dim sChamps As String
        Dim oXLApp As Excel.Application         'Declare the object variables
        Dim oXLBook As Excel.Workbook
        Dim oXLSheet As Excel.Worksheet
        Dim data_array(1 To 2500, 1 To 5) As Variant
        Dim r As Integer
        Set oXLApp = New Excel.Application    'Create a new instance of Excel
        Set oXLBook = oXLApp.Workbooks.Add    'Add a new workbook
        Set oXLSheet = oXLBook.Worksheets(1)  'Work with the first worksheet
        oXLApp.Visible = False
  
        Set rstInv = New ADODB.Recordset
  
        Call rstInv.Open("SELECT NoItem, QuantitéStock, Description, Manufacturier, Localisation FROM GRB_InventaireElec ORDER BY Localisation, NoItem", g_connData, adOpenDynamic, adLockOptimistic)
        r = 1
        
        'Fait un tableau pour envoyer dans excel.
        Do While Not rstInv.EOF
            data_array(r, 1) = rstInv.Fields("NoItem")
            data_array(r, 2) = rstInv.Fields("Description")
            data_array(r, 3) = rstInv.Fields("Manufacturier")
            data_array(r, 4) = rstInv.Fields("Localisation")
            data_array(r, 5) = rstInv.Fields("QuantitéStock")
            r = r + 1
            Call rstInv.MoveNext
         Loop

         



'creation en-tête de colonne
oXLSheet.Range("A1: E1").Font.Bold = True
oXLSheet.Range("A:E").HorizontalAlignment = xlCenter
oXLSheet.Range("A1: E1").Value = Array("NoItem", "Description", "Manufacturier", "Localisation", "QuantitéStock") 'GLL
'Fige la premiere ligne du tableau

oXLApp.ActiveSheet.Range("a2").Select
With oXLApp.ActiveWindow
    .FreezePanes = False
    .ScrollRow = 1
    .ScrollColumn = 1
    .FreezePanes = True
    .ScrollRow = 2
End With


'inscription des valeur du tableau dans excel
oXLSheet.Range("A2").Resize(r, 5).Value = data_array

'ajustement largeur des colonne
oXLSheet.Range("A:I").Columns.AutoFit
oXLApp.Visible = True

        
        
'Fermeture de la table de la base de donné
  Call rstInv.Close
  Set rstInv = Nothing

    Exit Function

AfficherErreur:

 woups "frmInventaireElec", "cmdExport_Click", Err, Erl
End Function
Private Function ExporterValeurComptable()

        On Error GoTo AfficherErreur

        'Impression de la l'inventaire trié par localisation et par no d'item
        Dim rstTotal As ADODB.Recordset
        Dim rstInv   As ADODB.Recordset
        Set rstInv = New ADODB.Recordset
        Set rstTotal = New ADODB.Recordset
        Dim oXLApp As Excel.Application         'Declare the object variables
        Dim oXLBook As Excel.Workbook
        Dim oXLSheet As Excel.Worksheet
        Dim data_array(1 To 2500, 1 To 9) As Variant
        Dim r As Integer
        Set oXLApp = New Excel.Application    'Create a new instance of Excel
        Set oXLBook = oXLApp.Workbooks.Add    'Add a new workbook
        Set oXLSheet = oXLBook.Worksheets(1)  'Work with the first worksheet
        oXLApp.Visible = False
        
        Call rstInv.Open("SELECT NoItem, Description, Manufacturier, [Prix Liste], Escompte, [Prix Net], Localisation, QuantitéStock, [Prix Net] * QuantitéStock As Total FROM GRB_InventaireElec WHERE QuantitéStock <> '0' ORDER BY NoItem", g_connData, adOpenDynamic, adLockOptimistic)
        
        Set DR_Inventaire.DataSource = rstInv
    
        Call rstTotal.Open("SELECT SUM([Prix net] * QuantitéStock) As Total FROM GRB_InventaireElec", g_connData, adOpenDynamic, adLockOptimistic)
        
        r = 1
        'Crée le tableau de donné
        Do While Not rstInv.EOF
            data_array(r, 1) = rstInv.Fields("NoItem")
            data_array(r, 2) = rstInv.Fields("Description")
            data_array(r, 3) = rstInv.Fields("manufacturier")
            data_array(r, 4) = rstInv.Fields("Localisation")
            data_array(r, 5) = rstInv.Fields("QuantitéStock")
            data_array(r, 6) = rstInv.Fields("prix liste")
            data_array(r, 7) = rstInv.Fields("escompte")
            data_array(r, 8) = rstInv.Fields("prix net")
            data_array(r, 9) = rstInv.Fields("Total")
            r = r + 1
            
        rstInv.MoveNext
        Loop
        'Ajoute le Grand Total a la fin du tableau
            data_array(r, 8) = "Grand Total:":
            data_array(r, 9) = rstTotal.Fields("Total")
    
    


        'creation en-tête de colonne
    oXLSheet.Range("A1: I1").Font.Bold = True
    oXLSheet.Range("A:i").HorizontalAlignment = xlCenter
    oXLSheet.Range("A1: I1").Value = Array("NoItem", "Description", "Manufacturier", "Localisation", "QuantitéStock", "Prix Lister", "Escompte", "Prix Net", "Total") 'GLL
        
        'Fige la premiere ligne de la colonne
    oXLApp.ActiveSheet.Range("a2").Select
        With oXLApp.ActiveWindow
        .FreezePanes = False
        .ScrollRow = 1
        .ScrollColumn = 1
        .FreezePanes = True
        .ScrollRow = 2
    End With
        'transfer le tableau dans la table d'excel
        oXLSheet.Range("A2").Resize(r, 9).Value = data_array
        'ajustement largeur des colonne
        oXLSheet.Range("A:K").Columns.AutoFit
        oXLSheet.Range("F:F").NumberFormat = "#,##0.00 $"
        oXLSheet.Range("H:I").NumberFormat = "#,##0.00 $"
        oXLApp.Visible = True
  
        Call rstTotal.Close
        Set rstTotal = Nothing
        Call rstInv.Close
        Set rstInv = Nothing

        Exit Function

AfficherErreur:

        woups "frmInventaireElec", "cmdexport_Click", Err, Erl
End Function
