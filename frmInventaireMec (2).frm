VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInventaireMec 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventaire mécanique"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmInventaireMec.frx":0000
   ScaleHeight     =   6615
   ScaleWidth      =   9990
   Begin VB.CommandButton cmdexporter 
      Caption         =   "Export vers excel"
      Height          =   375
      Left            =   1560
      TabIndex        =   18
      Top             =   6120
      Width           =   1695
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
   End
   Begin VB.Label lblTitrePlusMoins 
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
Attribute VB_Name = "frmInventaireMec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Index des colonnes de lvwHistorique
Private Const I_HIST_DATE As Integer = 0
Private Const I_HIST_USER As Integer = 1
Private Const I_HIST_PROJET As Integer = 2
Private Const I_HIST_PIECE As Integer = 3
Private Const I_HIST_QUANTITE As Integer = 4

'Index des chkChoix
Private Const I_CHK_CHOIX_PROJET As Integer = 0
Private Const I_CHK_CHOIX_PIECE As Integer = 1
Private Const I_CHK_CHOIX_DATE As Integer = 2

'Pour l'impression
Public m_bAnnuleImpression As Boolean
Public m_eTypeImpression As enumImpressionInventaire
Public m_typeImpressionExel As Boolean

Private Sub cmdexporter_Click()
m_typeImpressionExel = True 'pour l'affichage de Imprimer au lieu de Exporter dans frmchoixImpressionInventaire
 On Error GoTo Oups

 Call frmChoixImpressionInventaire.Afficher(Me)

 If m_bAnnuleImpression = False Then
 If m_eTypeImpression = MODE_AJUST_INV Then
 Call ExporterAjustementInventaire
 Else
 Call ExporterValeurComptable
 End If
 End If

 Exit Sub

Oups:

 wOups "frmInventaireMec", "cmdImprimer_Click", Err, Err.number, Err.Description
End Sub

Private Sub lvwHistorique_KeyDown(KeyCode As Integer, Shift As Integer)

 On Error GoTo Oups

 If KeyCode = vbKeyDelete Then
 If MsgBox("Voulez-vous vraiment effacer cet historique?", vbYesNo) = vbYes Then
 Call g_connData.Execute("DELETE * FROM GrbInventaireMecModif WHERE NoEnreg = " & lvwHistorique.SelectedItem.Tag)

 If MsgBox("Voulez-vous modifier la quantité dans l'inventaire?", vbYesNo) = vbYes Then
 Call ModifierInventaire(lvwHistorique.SelectedItem.SubItems(I_HIST_PIECE), lvwHistorique.SelectedItem.SubItems(I_HIST_QUANTITE))
 End If
 End If

 Call RemplirListViewHistorique
 End If

 Exit Sub

Oups:

  wOups "frmInventaireMec", "lvwHistorique_KeyDown", Err, Err.number, Err.Description
End Sub

Private Sub ModifierInventaire(ByVal sPiece As String, ByVal sQuantite As String)

 On Error GoTo Oups

 Dim rstInv As ADODB.Recordset

 If sQuantite > 0 Then
 sQuantite = "-" & sQuantite
 Else
 sQuantite = Replace(sQuantite, "-", "")
 End If

 Set rstInv = New ADODB.Recordset

 Call rstInv.Open("SELECT * FROM GrbInventaireMec WHERE NoItem = '" & Replace(sPiece, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)

 rstInv.Fields("QuantitéStock") = CDbl(rstInv.Fields("QuantitéStock")) + CDbl(sQuantite)

 Call rstInv.Update

  Call rstInv.Close
  Set rstInv = Nothing

  Exit Sub

Oups:

  wOups "frmInventaireMec", "ModifierInventaire", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewHistorique()

 On Error GoTo Oups

 'Rempli le ListView lvwHistorique
 Dim rstHist As ADODB.Recordset
 Dim sWhere As String
 Dim itmHist As ListItem
 
 'Si c'est une recherche avec le no de projet
 If chkChoix(I_CHK_CHOIX_PROJET).Value = vbChecked Then
 sWhere = "Left(IDProjet," & Len(txtnoprojet.Text) & ") = '" & txtnoprojet.Text & "' "
 End If
 
 'Si c'est une recherche avec le no de la piece
 If chkChoix(I_CHK_CHOIX_PIECE).Value = vbChecked Then
 If sWhere = vbNullString Then
 sWhere = "INSTR(1,NoItem,'" & Replace(txtNoPiece.Text, "'", "''") & "') > 0 "
 Else
  sWhere = sWhere & " AND INSTR(1,NoItem,'" & Replace(txtNoPiece.Text, "'", "''") & "') > 0 "
  End If
  End If
 
 'Si c'est une recherche par date
  If chkChoix(I_CHK_CHOIX_DATE).Value = vbChecked Then
  If sWhere = vbNullString Then
  sWhere = "Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "'"
  Else
  sWhere = sWhere & " AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "'"
End If
End If
 
Set rstHist = New ADODB.Recordset
 
Call rstHist.Open("SELECT * FROM GrbInventaireMecModif WHERE " & sWhere & " ORDER BY NoEnreg DESC", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Il faut vider le ListView avant de le remplir
Call lvwHistorique.ListItems.Clear
 
 'Tant que ce n'est pas la fin des enregistrements
Do While Not rstHist.EOF
 'On l'ajoute
 Set itmHist = lvwHistorique.ListItems.Add
 
 'Date
 itmHist.Text = rstHist.Fields("Date")
 itmHist.Tag = rstHist.Fields("NoEnreg")
 
 'User
 itmHist.SubItems(I_HIST_USER) = rstHist.Fields("User")
 
 'IDProjet
 If Not IsNull(rstHist.Fields("IDProjet")) Then
 If Trim(rstHist.Fields("IDProjet")) <> "" Then
 itmHist.SubItems(I_HIST_PROJET) = rstHist.Fields("IDProjet")
 Else
 itmHist.SubItems(I_HIST_PROJET) = "N/A"
 End If
 Else
 itmHist.SubItems(I_HIST_PROJET) = "N/A"
 End If
 
 'No Item
1  itmHist.SubItems(I_HIST_PIECE) = rstHist.Fields("NoItem")
 
 'Quantité
 itmHist.SubItems(I_HIST_QUANTITE) = rstHist.Fields("Quantité")
 
 Call rstHist.MoveNext
Loop
 
Call rstHist.Close
Set rstHist = Nothing

Exit Sub

Oups:

wOups "frmInventaireMec", "RemplirListViewHistorique", Err, Err.number, Err.Description
End Sub

Private Sub chkChoix_Click(Index As Integer)

 On Error GoTo Oups

 'Méthode qui met enabled ou disabled les controles selon le type de recherche
 Dim bEnabled As Boolean
 
 If chkChoix(Index).Value = vbChecked Then
 bEnabled = True
 Else
 bEnabled = False
 End If
 
 Select Case Index
 Case I_CHK_CHOIX_PROJET:
 fraProjet.Enabled = bEnabled
 
 Case I_CHK_CHOIX_PIECE:
 fraPiece.Enabled = bEnabled

 Case I_CHK_CHOIX_DATE:
 fraDate.Enabled = bEnabled
  End Select

  Exit Sub

Oups:

  wOups "frmInventaireMec", "chkChoix_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAfficherHistorique_Click()

 On Error GoTo Oups

 'Affiche l'historique des modifications de l'inventaire

 If chkChoix(I_CHK_CHOIX_DATE).Value = vbChecked Then
 If mskDateDebut.Text <> vbNullString And mskDateFin.Text <> vbNullString Then
 If mskDateDebut.Text > mskDateFin.Text Then
 Call MsgBox("La date de fin doit être plus petite que la date de début!", vbOKOnly, "Erreur")
 
 Exit Sub
 End If
 Else
 Call MsgBox("Vous devez remplir tous les champs!", vbOKOnly, "Erreur")
 
 Exit Sub
 End If
  End If
 
  Screen.MousePointer = vbHourglass
 
  If chkChoix(I_CHK_CHOIX_DATE).Value = vbChecked Or _
 chkChoix(I_CHK_CHOIX_PIECE).Value = vbChecked Or _
 chkChoix(I_CHK_CHOIX_PROJET).Value = vbChecked Then
  Call RemplirListViewHistorique
  Else
  Call MsgBox("Vous devez choisir au moins une option de recherche!", vbOKOnly, "Erreur")
  End If

  Screen.MousePointer = vbDefault

10 Exit Sub

Oups:

wOups "frmInventaireMec", "cmdAfficherHistorique_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdImprimer_Click()
 m_typeImpressionExel = False 'pour l'affichage de Imprimer au lieu de Exporter dans frmchoixImpressionInventaire
 On Error GoTo Oups

 Call frmChoixImpressionInventaire.Afficher(Me)

 If m_bAnnuleImpression = False Then
 If m_eTypeImpression = MODE_AJUST_INV Then
 Call ImprimerAjustementInventaire
 Else
 Call ImprimerValeurComptable
 End If
 End If

 Exit Sub

Oups:

 wOups "frmInventaireMec", "cmdImprimer_Click", Err, Err.number, Err.Description
End Sub

Private Sub Cmdfermer_Click()

 On Error GoTo Oups

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmInventaireMec", "cmdFermer_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()
 
 On Error GoTo Oups

 Call Unload(frmChoixInventaire)
 
 Exit Sub

Oups:

 wOups "frmInventaireMec", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub lvwHistorique_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

 On Error GoTo Oups

 If lvwHistorique.ListItems.count > 0 Then
 lvwHistorique.Sorted = True

 lvwHistorique.SortKey = ColumnHeader.Index - 1
 
 If lvwHistorique.SortOrder = lvwAscending Then
 lvwHistorique.SortOrder = lvwDescending
 Else
 lvwHistorique.SortOrder = lvwAscending
 End If
 End If

 Exit Sub

Oups:

  wOups "frmInventaireMec", "lvwHistorique_ColumnClick", Err, Err.number, Err.Description
End Sub

Private Sub mskDateDebut_GotFocus()

 On Error GoTo Oups

 If Len(mskDateDebut.Text) = 10 Then
 mskDateDebut.Text = Right$(mskDateDebut.Text, 8)
 End If
 
 mskDateDebut.mask = "##-##-##"

 Exit Sub

Oups:

 wOups "frmInventaireMec", "mskDateDebut_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskDateDebut_LostFocus()

 On Error GoTo Oups

 mskDateDebut.mask = vbNullString
 
 If mskDateDebut.Text = "__-__-__" Then
 mskDateDebut.Text = vbNullString
 Else
 If Len(mskDateDebut.Text) =   Then
 If IsDate(mskDateDebut.Text) Then
 mskDateDebut.Text = Year(DateSerial(Left$(mskDateDebut.Text, 2), Mid$(mskDateDebut.Text, 4, 2), Right$(mskDateDebut.Text, 2))) & Mid$(mskDateDebut.Text, 3, 8)
 End If
 End If
 End If

  Exit Sub

Oups:

  wOups "frmInventaireMec", "mskDateDebut_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskDateFin_GotFocus()

 On Error GoTo Oups

 If Len(mskDateFin.Text) = 10 Then
 mskDateFin.Text = Right$(mskDateFin.Text, 8)
 End If
 
 mskDateFin.mask = "##-##-##"

 Exit Sub

Oups:

 wOups "frmInventaireMec", "mskDateFin_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskDateFin_LostFocus()
 
 On Error GoTo Oups

 mskDateFin.mask = vbNullString
 
 If mskDateFin.Text = "__-__-__" Then
 mskDateFin.Text = vbNullString
 Else
 If Len(mskDateFin.Text) =   Then
 If IsDate(mskDateFin.Text) Then
 mskDateFin.Text = Year(DateSerial(Left$(mskDateFin.Text, 2), Mid$(mskDateFin.Text, 4, 2), Right$(mskDateFin.Text, 2))) & Mid$(mskDateFin.Text, 3, 8)
 End If
 End If
 End If

  Exit Sub

Oups:

  wOups "frmInventaireMec", "mskDateFin_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerAjustementInventaire()

 On Error GoTo Oups

 'Impression de la l'inventaire trié par localisation et par no d'item
 Dim rstInv As ADODB.Recordset
 Dim sChamps As String
 
 Set rstInv = New ADODB.Recordset
 
 Call rstInv.Open("SELECT NoItem, QuantitéStock, Description, Manufacturier, Localisation FROM GrbInventaireMec ORDER BY Localisation, NoItem", g_connData, adOpenDynamic, adLockOptimistic)
 
 Set DR_ListeInventaire.DataSource = rstInv
 
 DR_ListeInventaire.Sections("Section3").Controls("lblDate").Caption = ConvertDate(Date)
 DR_ListeInventaire.Sections("Section4").Controls("lblTitre").Caption = "Inventaire Mécanique"
 
 Call DR_ListeInventaire.Show(vbModal)
 
 Call rstInv.Close
 Set rstInv = Nothing

  Exit Sub

Oups:

  wOups "frmInventaireMec", "ImprimerAjustementInventaire", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerValeurComptable()

 On Error GoTo Oups

 'Impression de la l'inventaire trié par localisation et par no d'item
 Dim rstInv As ADODB.Recordset
 Dim rstTotal As ADODB.Recordset
 
 Set rstInv = New ADODB.Recordset
 Set rstTotal = New ADODB.Recordset
 
 Call rstInv.Open("SELECT NoItem, Description, Manufacturier, [Prix Liste], Escompte, [Prix Net], Localisation, QuantitéStock, [Prix Net] * QuantitéStock As Total FROM GrbInventaireMec WHERE QuantitéStock <> '0' ORDER BY NoItem", g_connData, adOpenDynamic, adLockOptimistic)
 
 Set DR_Inventaire.DataSource = rstInv
 
 Call rstTotal.Open("SELECT SUM([Prix net] * QuantitéStock) As Total FROM GrbInventaireMec", g_connData, adOpenDynamic, adLockOptimistic)
 
 DR_Inventaire.Sections("Section3").Controls("lblDate").Caption = ConvertDate(Date)
 DR_Inventaire.Sections("Section5").Controls("lblTotal").Caption = Conversion(rstTotal.Fields("Total"), MODE_ARGENT)
 DR_Inventaire.Sections("Section4").Controls("lblTitre").Caption = "Inventaire Mécanique"
 
  Call rstTotal.Close
  Set rstTotal = Nothing

  DR_Inventaire.Orientation = rptOrientLandscape
 
  Call DR_Inventaire.Show(vbModal)
 
  Call rstInv.Close
  Set rstInv = Nothing

  Exit Sub

Oups:

  wOups "frmInventaireMec", "ImprimerValeurComptable", Err, Err.number, Err.Description
End Sub
Private Function ExporterAjustementInventaire()
 On Error GoTo Oups

 'Impression de l'inventaire trié par localisation et par no d'item
 Dim rstInv As ADODB.Recordset
 Dim sChamps As String
 Dim oXLApp As Excel.Application 'Declare the object variables
 Dim oXLBook As Excel.Workbook
 Dim oXLSheet As Excel.Worksheet
 Dim data_array(1 To 2500, 1 To 5) As Variant
 Dim r As Integer
 Set oXLApp = New Excel.Application 'Create a new instance of Excel
 Set oXLBook = oXLApp.Workbooks.Add 'Add a new workbook
 Set oXLSheet = oXLBook.Worksheets(1) 'Work with the first worksheet
 oXLApp.Visible = False
 
 Set rstInv = New ADODB.Recordset
 
 Call rstInv.Open("SELECT NoItem, QuantitéStock, Description, Manufacturier, Localisation FROM GrbInventaireMec ORDER BY Localisation, NoItem", g_connData, adOpenDynamic, adLockOptimistic)
 r = 1
 
 'ajoute les donné dans un tableau
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
oXLSheet.range("A1: E1").Font.Bold = True
oXLSheet.range("A:E").HorizontalAlignment = xlCenter
oXLSheet.range("A1: E1").Value = Array("NoItem", "Description", "Manufacturier", "Localisation", "QuantitéStock") 'GLL
'figer la premiere ligne de la table
oXLApp.ActiveSheet.range("a2").Select
With oXLApp.ActiveWindow
 .FreezePanes = False
 .ScrollRow = 1
 .ScrollColumn = 1
 .FreezePanes = True
 .ScrollRow = 2
End With


'inscription des valeur du tableau dans excel
oXLSheet.range("A2").Resize(r, 5).Value = data_array

'ajustement largeur des colonne
oXLSheet.range("A:I").Columns.AutoFit
oXLApp.Visible = True

 
 

 Call rstInv.Close
 Set rstInv = Nothing

  Exit Function

Oups:

  wOups "frmInventaireElec", "cmdExport_Click", Err, Err.number, Err.Description
End Function
Private Function ExporterValeurComptable()

 On Error GoTo Oups

 'Impression de la l'inventaire trié par localisation et par no d'item
 Dim rstTotal As ADODB.Recordset
 Dim rstInv As ADODB.Recordset
 Set rstInv = New ADODB.Recordset
 Set rstTotal = New ADODB.Recordset
 Dim oXLApp As Excel.Application 'Declare the object variables
 Dim oXLBook As Excel.Workbook
 Dim oXLSheet As Excel.Worksheet
 Dim data_array(1 To 2500, 1 To 9) As Variant
 Dim r As Integer
 Set oXLApp = New Excel.Application 'Create a new instance of Excel
 Set oXLBook = oXLApp.Workbooks.Add 'Add a new workbook
 Set oXLSheet = oXLBook.Worksheets(1) 'Work with the first worksheet
 oXLApp.Visible = False
 
 Call rstInv.Open("SELECT NoItem, Description, Manufacturier, [Prix Liste], Escompte, [Prix Net], Localisation, QuantitéStock, [Prix Net] * QuantitéStock As Total FROM GrbInventaireMec WHERE QuantitéStock <> '0' ORDER BY NoItem", g_connData, adOpenDynamic, adLockOptimistic)
 
 Set DR_Inventaire.DataSource = rstInv
 
 Call rstTotal.Open("SELECT SUM([Prix net] * QuantitéStock) As Total FROM GrbInventaireMec", g_connData, adOpenDynamic, adLockOptimistic)
 
 r = 1
 'Ajouter les valeur lu au tableau
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
 data_array(r, 8) = "Grand Total"
 data_array(r, 9) = rstTotal.Fields("Total")
 
 
 


 'creation en-tête de colonne
 oXLSheet.range("A1: I1").Font.Bold = True
 oXLSheet.range("A:E").HorizontalAlignment = xlCenter
 oXLSheet.range("A1: I1").Value = Array("NoItem", "Description", "Manufacturier", "Localisation", "QuantitéStock", "Prix Lister", "Escompte", "Prix Net", "Total") 'GLL
 
 'Figer la premiere ligne de la table
 oXLApp.ActiveSheet.range("a2").Select
 With oXLApp.ActiveWindow
 .FreezePanes = False
 .ScrollRow = 1
 .ScrollColumn = 1
 .FreezePanes = True
 .ScrollRow = 2
 End With
 'transfer le tableau dans la table de excel
 oXLSheet.range("A2").Resize(r, 9).Value = data_array
 oXLSheet.range("F:F").NumberFormat = "#,##0.00 $"
 oXLSheet.range("H:I").NumberFormat = "#,##0.00 $"
 
 'ajustement largeur des colonne
 oXLSheet.range("A:I").Columns.AutoFit
 oXLApp.Visible = True
 
 Call rstTotal.Close
 Set rstTotal = Nothing
 Call rstInv.Close
 Set rstInv = Nothing

 Exit Function

Oups:

 wOups "frmInventaireElec", "cmdexport_Click", Err, Err.number, Err.Description
End Function

