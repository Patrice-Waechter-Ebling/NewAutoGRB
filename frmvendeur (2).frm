VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmvendeur 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contacts pour vendeurs"
   ClientHeight    =   6345
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8100
   Icon            =   "frmvendeur.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6345
   ScaleWidth      =   8100
   Begin MSComctlLib.ListView lstclient 
      Height          =   1815
      Left            =   1200
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Catalogue"
         Object.Width           =   4471
      EndProperty
   End
   Begin VB.CommandButton cmdrechercheclient 
      Height          =   375
      Left            =   4080
      Picture         =   "frmvendeur.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton CmdExporter 
      Caption         =   "&Exporter vers Excel"
      Height          =   495
      Left            =   5280
      TabIndex        =   22
      Top             =   5760
      Width           =   2175
   End
   Begin VB.CommandButton cmdcherche 
      Caption         =   "Cherche par date"
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin MSMask.MaskEdBox mskDateCherche 
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.ComboBox cmbClient 
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "cmbClient"
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&Ajouter"
      Height          =   495
      Left            =   360
      TabIndex        =   17
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton CmdSupp 
      Caption         =   "&Supprimer"
      Height          =   495
      Left            =   2040
      TabIndex        =   18
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "&Fermer"
      Height          =   495
      Left            =   3600
      TabIndex        =   19
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Frame fracontact 
      Caption         =   "Contact"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   7935
      Begin VB.ComboBox cmbtype 
         Height          =   315
         ItemData        =   "frmvendeur.frx":0784
         Left            =   4080
         List            =   "frmvendeur.frx":0786
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox txtNomCompagny 
         Enabled         =   0   'False
         Height          =   375
         Left            =   480
         TabIndex        =   20
         Top             =   1200
         Width           =   2295
      End
      Begin MSMask.MaskEdBox txtdate 
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtcommentaire 
         Height          =   1365
         Left            =   480
         TabIndex        =   14
         Top             =   1800
         Width           =   6975
      End
      Begin VB.TextBox txtcontact 
         Height          =   375
         Left            =   4200
         TabIndex        =   12
         Top             =   1200
         Width           =   3255
      End
      Begin VB.CommandButton cmdfermercontact 
         Caption         =   "Fermer"
         Height          =   495
         Left            =   5880
         TabIndex        =   16
         Top             =   3240
         Width           =   1575
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "Sauvegarde"
         Height          =   495
         Left            =   3960
         TabIndex        =   15
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "État"
         Height          =   255
         Left            =   4080
         TabIndex        =   26
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblNomcompagnie 
         Caption         =   "Nom de Compagnie"
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Commentaire"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Contact"
         Height          =   255
         Left            =   4080
         TabIndex        =   10
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label label2 
         Caption         =   "Date"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.ListView lister 
      Height          =   3855
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   6800
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nom de Compagnie"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Contact"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "État"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Commentaire"
         Object.Width           =   10583
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Enregistrer par"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lbladresse 
      Caption         =   "lbladresse"
      Height          =   495
      Left            =   1200
      TabIndex        =   25
      Top             =   720
      Width           =   6735
   End
   Begin VB.Label Label3 
      Caption         =   "Date  AA-MM-JJ"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lbltelephone 
      Caption         =   "lblTELEPHONE"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   7815
   End
   Begin VB.Label Label1 
      Caption         =   "NomClient"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Adresse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "frmvendeur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enumModeCherche
 MODE_DATE = 0
 MODE_CLIENT = 1
End Enum

Private m_bModeAjouter As Boolean
Private m_eModeCherche As enumModeCherche
Private numéroCompagnie As Integer
Private FieldOk As Boolean
Private Sub FindFieldsExist(Name As String)
 On Error GoTo Oups
 Dim strName As String
 Dim Findfield As ADODB.Recordset
 
 Dim i As Integer
 FieldOk = False
 Set Findfield = New ADODB.Recordset
 Call Findfield.Open("Select * from GrbVendeur", g_connData, adOpenDynamic, adLockOptimistic)
 For i = 0 To Findfield.Fields.count - 1
 strName = Findfield.Fields(i).Name
 If strName = Name Then
 FieldOk = True
 Call Findfield.Close
 Set Findfield = Nothing
 Exit Sub
 End If
 Next
 Call Findfield.Close
 Set Findfield = Nothing
 Call g_connData.Execute("ALTER TABLE GrbVendeur Add " & Name & " Text(25);")
 FieldOk = False
 Exit Sub
Oups:
 wOups "frmvendeur", "FindFieldsExist()", Err, Err.number, Err.Description
 End Sub




Private Sub remplir_lister()

 On Error GoTo Oups

 '''''''''''''''''',
 'rempli le lister
 ''''''''''''''''''''',
 Dim rstVendeur As ADODB.Recordset
 Dim itmVendeur As ListItem

 Call FindFieldsExist("EnregPar")
 Call FindFieldsExist("Etat")
 m_eModeCherche = MODE_CLIENT
 
 CmdAdd.Visible = True
 
 'vide lister
 Call lister.ListItems.Clear
 
 'ouvre la table pour client
 Set rstVendeur = New ADODB.Recordset
 
 Call rstVendeur.Open("SELECT * FROM Grbvendeur WHERE IDClient = " & numéroCompagnie & " ORDER BY no", g_connData, adOpenDynamic, adLockOptimistic)
 
 'temp que pas a la fin de la table
 Do While Not rstVendeur.EOF
 'ajoute au lister
 Set itmVendeur = lister.ListItems.Add
 
 'no du client
 itmVendeur.Tag = rstVendeur.Fields("no")
 
 'vérifie les champs vide avant d'inséré
 'date
  If IsNull(rstVendeur.Fields("Date")) Then
  itmVendeur.Text = " "
  Else
  itmVendeur.Text = ConvertDate(DateSerial(Left(rstVendeur.Fields("Date"), 2), Mid(rstVendeur.Fields("Date"), 4, 2), Right(rstVendeur.Fields("Date"), 2)))
  End If
 '
 If IsNull(rstVendeur.Fields("Contact")) And IsNull(rstVendeur.Fields("commentaire")) And IsNull(rstVendeur.Fields("Date")) Then
 Call itmVendeur.ListSubItems.Add(, , vbNullString)
 Else
 Call itmVendeur.ListSubItems.Add(, , cmbclient.Text)
 End If
 
 
 
 
 
 
 'contact
  If IsNull(rstVendeur.Fields("Contact")) Then
  Call itmVendeur.ListSubItems.Add(, , vbNullString)
  Else
 Call itmVendeur.ListSubItems.Add(, , rstVendeur.Fields("Contact"))
1 End If

 'Type
 If IsNull(rstVendeur.Fields("Etat")) Then
 Call itmVendeur.ListSubItems.Add(, , vbNullString)
 Else
 Call itmVendeur.ListSubItems.Add(, , rstVendeur.Fields("Etat"))
 End If
 
 
 'commentaire
 If IsNull(rstVendeur.Fields("commentaire")) Then
 Call itmVendeur.ListSubItems.Add(, , vbNullString)
 Else
 Call itmVendeur.ListSubItems.Add(, , rstVendeur.Fields("commentaire"))
 End If
 'Enregistrer par
 If IsNull(rstVendeur.Fields("EnregPar")) Then
 Call itmVendeur.ListSubItems.Add(, , vbNullString)
 Else
 Call itmVendeur.ListSubItems.Add(, , rstVendeur.Fields("EnregPar"))
 End If
 
 
 
 'prochaine enreg
 Call rstVendeur.MoveNext
Loop
 
 'fermeture table et bd
Call rstVendeur.Close
Set rstVendeur = Nothing

Exit Sub

Oups:

1  wOups "frmvendeur", "remplir_lister", Err, Err.number, Err.Description
End Sub

Private Sub cmbclient_Click()

 On Error GoTo Oups

 '''''''''''''''''''''''''''''
 'lorsque on select un client
 '''''''''''''''''''''''''''''
 Dim rstClient As ADODB.Recordset

 If cmbclient.ListIndex <> -1 Then
 
 'met visible fenetre pour ajouter
 fracontact.Visible = False
 
 'mode ajouter ou editer
 m_bModeAjouter = False
 
 'set le rapport
 Set rstClient = New ADODB.Recordset
 
 Call rstClient.Open("SELECT * FROM GrbClient WHERE IDClient = " & cmbclient.ItemData(cmbclient.ListIndex) & " ORDER BY NomClient", g_connData, adOpenDynamic, adLockOptimistic)

 'initialise label adress et teleph
 lblAdresse.Caption = vbNullString
 lbltelephone.Caption = vbNullString
 
 numéroCompagnie = rstClient.Fields("idclient")
 'si client existe
 ''''''''''''''''''''''''''''''''''''''''''''
 'rempli adresse pays ville prov et codepostal si pas vide
 '''''''''''''''''''''''''''''''''''''''''''''
 'adresse
 If Not rstClient.Fields("adresseliv") = "" Then
 lblAdresse.Caption = lblAdresse.Caption + rstClient.Fields("adresseliv")
  End If
 
 'ville
  If Not rstClient.Fields("villeliv") = "" Then
  lblAdresse.Caption = lblAdresse.Caption + ", " + rstClient.Fields("villeliv")
  End If
 
 'pays
  If Not rstClient.Fields("paysliv") = "" Then
  lblAdresse.Caption = lblAdresse.Caption + ", " + rstClient.Fields("paysliv")
  End If
 
 'province
  If Not rstClient.Fields("prov/etatliv") = "" Then
 lblAdresse.Caption = lblAdresse.Caption + ", " + rstClient.Fields("prov/etatliv")
1 End If
 
 'codepostal
 If Not rstClient.Fields("codepostalliv") = "" Then
 lblAdresse.Caption = lblAdresse.Caption + ", " + rstClient.Fields("codepostalliv")
 End If
 
 ''''''''''''''''''''''''''''''''''
 'rempli tel fax pagette cell email si pas vide
 ''''''''''''''''''''''''''''''''''
 'telephone
 If Not rstClient.Fields("telephonne") = "" Then
 lbltelephone.Caption = lbltelephone.Caption + "TÉL: " + rstClient.Fields("telephonne")
 End If
 
 'fax
 If Not rstClient.Fields("fax") = "" Then
 lbltelephone.Caption = lbltelephone.Caption + " FAX: " + rstClient.Fields("fax")
 End If
 
 'pagette
 If Not rstClient.Fields("pagette") = "" Then
 lbltelephone.Caption = lbltelephone.Caption + " PAGE: " + rstClient.Fields("pagette")
 End If
 
 'cellulaire
 If Not rstClient.Fields("cellulaire") = "" Then
 lbltelephone.Caption = lbltelephone.Caption + " CELL: " + rstClient.Fields("cellulaire")
 End If
 
 'email
 If Not rstClient.Fields("email") = "" Then
 lbltelephone.Caption = lbltelephone.Caption + " EMAIL: " + rstClient.Fields("email")
1  End If
 txtNomCompagny.Text = rstClient.Fields("NomClient")
 Call rstClient.Close
 Set rstClient = Nothing
 
 'rempli le lister
 
 Call remplir_lister
End If
 
Exit Sub

Oups:

wOups "frmvendeur", "cmbclient_Click", Err, Err.number, Err.Description
End Sub





Private Sub CmdAdd_Click()

 On Error GoTo Oups

 '''''''''''''''''''''''''''''''''
 'ajoute un contact
 ''''''''''''''''''''''''''''''''
 'met visible fenetre pour ajouter
 fracontact.Visible = True
 fracontact.Tag = numéroCompagnie
 'mode ajouter ou editer
 m_bModeAjouter = True

 'valeur par defaut sur l'ouverture
 txtDate.Text = Year(Date) & "-" & Right$("0" & Month(Date), 2) & "-" & Right$("0" & Day(Date), 2)
 txtNomCompagny.Text = cmbclient.Text
 txtcommentaire.Text = vbNullString
 txtcontact.Text = vbNullString

 Exit Sub

Oups:

 wOups "frmvendeur", "CmdAdd_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdcherche_Click()

 On Error GoTo Oups
 fracontact.Visible = False

 Call remplir_lister_date

 Exit Sub

Oups:

 wOups "frmvendeur", "cmdcherche_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdexporter_Click()
Dim xlworksheet As Excel.Workbook
Dim xlsheet As Excel.Application
Dim info As ADODB.Recordset
Dim row As Integer
Dim i As Integer



Set xlsheet = New Excel.Application
Set xlworksheet = xlsheet.Workbooks.Add
row = 1


If m_eModeCherche = MODE_CLIENT Then
 row = 6
 xlsheet.Cells(1, 1) = "Client:"
 xlsheet.Cells(2, 1) = "Adresse:"
 xlsheet.Cells(3, 1) = "Téléphone:"
 xlsheet.Cells(2, 3) = "Ville:"
 xlsheet.Cells(3, 3) = "Fax:"
 xlsheet.Cells(2, 5) = "Pays:"
 xlsheet.Cells(3, 5) = "Page:"
 xlsheet.Cells(2, 7) = "Province/État:"
 xlsheet.Cells(3, 7) = "Cell:"
 xlsheet.Cells(2, 9) = "Codepostal:"
 xlsheet.Cells(3, 9) = "Email:"
 xlsheet.Cells(5, 1) = "Date:"
 xlsheet.Cells(5, 2) = "Nom de la Compagnie"
 xlsheet.Cells(5, 3) = "Nom du Contact"
 xlsheet.Cells(5, 4) = "État"
 xlsheet.Cells(5, 5) = "Commentaire"
 xlsheet.Cells(5, 9) = "Enregister Par"
 
 With xlsheet.range("A1:A3;C2:C3;E2:E3;G2:G3;I2:I3")
 .Font.Bold = True
 .HorizontalAlignment = xlRight
 .Font.SIZE = 11
 End With
 With xlsheet.range("A5:I5")
 .Font.Bold = True
 .HorizontalAlignment = xlLeft
 .Font.SIZE = 11
 End With
 
 
 
 
 
 Set info = New ADODB.Recordset
 Call info.Open("Select * From Grbclient where IDClient = " & numéroCompagnie, g_connData, adOpenDynamic, adLockOptimistic)
 Do While Not info.EOF
 xlsheet.Cells(1, 2) = info.Fields("Nomclient")
 xlsheet.Cells(2, 2) = info.Fields("AdresseLiv")
 xlsheet.Cells(3, 2) = info.Fields("Telephonne")
 xlsheet.Cells(2, 4) = info.Fields("VilleLiv")
 xlsheet.Cells(3, 4) = info.Fields("Fax")
 xlsheet.Cells(2, 6) = info.Fields("PaysLiv")
 xlsheet.Cells(3, 6) = info.Fields("Pagette")
 xlsheet.Cells(2, 8) = info.Fields("Prov/EtatLiv")
 xlsheet.Cells(3, 8) = info.Fields("Cellulaire")
 xlsheet.Cells(3, 10) = info.Fields("CodePostalLiv")
 xlsheet.Cells(3, 10) = info.Fields("Email")
 Call info.MoveNext
 Loop
 For i = 1 To lister.ListItems.count
 xlsheet.Cells(row, 1) = lister.ListItems(i).Text
 xlsheet.Cells(row, 2) = lister.ListItems(i).SubItems(1)
 xlsheet.Cells(row, 3) = lister.ListItems(i).SubItems(2)
 xlsheet.Cells(row, 4) = lister.ListItems(i).SubItems(3)
 xlsheet.Cells(row, 5) = lister.ListItems(i).SubItems(4)
 xlsheet.Cells(row, 9) = lister.ListItems(i).SubItems(5)
 xlsheet.range("E" & row & ":H" & row).Merge
 row = row + 1
 Next
 xlsheet.range("A:J").Columns.AutoFit
 info.Close
 Set info = Nothing
Else
 If lister.ListItems.count <= 0 Then Exit Sub
 row = 3
 xlsheet.range("A1:D1").Merge
 xlsheet.Cells(1, 1) = "Liste des contacts en date du " & lister.ListItems(1).Text
 xlsheet.Cells(2, 1) = "Date:"
 xlsheet.Cells(2, 2) = "Nom de la Compagnie"
 xlsheet.Cells(2, 3) = "Nom du Contact"
 xlsheet.Cells(2, 4) = "État"
 xlsheet.Cells(2, 5) = "Commentaire"
 xlsheet.Cells(2, 6) = "Enregister Par"
 With xlsheet.range("A1;A2;A2:F2")
 .Font.Bold = True
 .Font.SIZE = 11
 End With
 For i = 1 To lister.ListItems.count
 xlsheet.Cells(row, 1) = lister.ListItems(i).Text
 xlsheet.Cells(row, 2) = lister.ListItems(i).SubItems(1)
 xlsheet.Cells(row, 3) = lister.ListItems(i).SubItems(2)
 xlsheet.Cells(row, 4) = lister.ListItems(i).SubItems(3)
 xlsheet.Cells(row, 5) = lister.ListItems(i).SubItems(4)
 xlsheet.Cells(row, 6) = lister.ListItems(i).SubItems(5)
 row = row + 1
 Next
 xlsheet.range("A:F").Columns.AutoFit
End If




xlsheet.Visible = True



End Sub

Private Sub cmdfermercontact_Click()

 On Error GoTo Oups

 ''''''''''''''''''''''
 'Quitte liste contact'
 ''''''''''''''''''''''
 'cache fenêtre
 fracontact.Visible = False
 If m_eModeCherche = MODE_CLIENT Then
 Call remplir_lister
 Else
 Call remplir_lister_date
 End If

 Exit Sub

Oups:

 wOups "frmvendeur", "cmdfermercontact_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdquit_Click()

 On Error GoTo Oups

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmvendeur", "cmdquit_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdrechercheclient_Click()

 On Error GoTo Oups

 Dim rstcatalog As ADODB.Recordset
 Dim sDescription As String
 Dim itmDescription As ListItem
 
 sDescription = InputBox("Quelle est la description à rechercher")
 
 If sDescription <> vbNullString Then
 Call lstclient.ListItems.Clear
 
 sDescription = Replace(sDescription, "'", "''")
 
 sDescription = "%" & sDescription & "%"

 Set rstcatalog = New ADODB.Recordset

55
 Call rstcatalog.Open("SELECT DISTINCT NomClient FROM GrbClient WHERE NomClient LIKE '" & sDescription & "' ORDER BY NomClient", g_connData, adOpenDynamic, adLockOptimistic)
  Do While Not rstcatalog.EOF
  Set itmDescription = lstclient.ListItems.Add()
 
  itmDescription.Tag = rstcatalog.Fields("NomClient")
 itmDescription.Text = rstcatalog.Fields("NomClient")

 Call rstcatalog.MoveNext
Loop
 
 Call rstcatalog.Close
 Set rstcatalog = Nothing

 If lstclient.ListItems.count > 0 Then
 lstclient.Visible = True

 Call lstclient.SetFocus
 Else
1  Call MsgBox("Aucun enregistrement trouvé!")
 End If
 End If

Exit Sub

Oups:

wOups "frmvendeur", "cmdrechercheclient_Click", Err, Err.number, Err.Description


End Sub

Private Sub cmdsave_Click()

 On Error GoTo Oups

 Dim rstVendeur As ADODB.Recordset
 
 'table vendeur ouvert
 Set rstVendeur = New ADODB.Recordset
 Call FindFieldsExist("Enregistrerpar")
 Call FindFieldsExist("Type")
 
 If m_bModeAjouter = True Then
 Call rstVendeur.Open("SELECT * FROM Grbvendeur", g_connData, adOpenDynamic, adLockOptimistic)
 
 Call rstVendeur.AddNew
 
 m_bModeAjouter = False
 Else
 Call rstVendeur.Open("SELECT * FROM GrbVendeur WHERE [no] = " & lister.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
 End If
 
 
 
 
 rstVendeur.Fields("IDClient").Value = fracontact.Tag
  rstVendeur.Fields("Date").Value = Right$(txtDate.Text, 8)
  rstVendeur.Fields("Contact").Value = txtcontact.Text
  rstVendeur.Fields("commentaire").Value = txtcommentaire.Text
 rstVendeur.Fields("EnregPar").Value = CStr(g_sUserID)
 rstVendeur.Fields("Etat").Value = CStr(cmbType.Text)
 
  Call rstVendeur.Update
 
 'ferme la table
  Call rstVendeur.Close
  Set rstVendeur = Nothing
 
  fracontact.Visible = False
 
 'rempli le lister
  If m_eModeCherche = MODE_CLIENT Then
Call remplir_lister
Else
 Call remplir_lister_date
End If

Exit Sub

Oups:

wOups "frmvendeur", "cmdsave_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdSupp_Click()

 On Error GoTo Oups

 'Supprime l'enregistrement sélectionné
 If lister.ListItems.count > 0 Then
 Call g_connData.Execute("DELETE * FROM GrbVendeur WHERE [no] = " & lister.SelectedItem.Tag)
 
 Call remplir_lister
 End If

 Exit Sub

Oups:

 wOups "frmvendeur", "CmdSupp_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 'initialise les labels adresse et telephone
 lblAdresse.Caption = vbNullString
 lbltelephone.Caption = vbNullString

 cmbType.AddItem ("Piste de vente")
 cmbType.AddItem ("Opportunité")
 cmbType.AddItem ("Soumission")
 cmbType.AddItem ("Gagner")
 cmbType.AddItem ("Perdu")
 cmbType.AddItem ("En attente")
 cmbType.ListIndex = 0
 

 'rempli le comboclient et lister
 Call remplir_comboclient

 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmvendeur", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub lister_DblClick()

 On Error GoTo Oups

 'Sur DblClick, affiche fenêtre pour modifié l'enreg selectionné dans lister
 Dim rstVendeur As ADODB.Recordset

 If lister.ListItems.count <> 0 Then
 If m_eModeCherche = MODE_CLIENT Then
 'si lister pas vide
 'met fenetre visible
 fracontact.Visible = True
 
 'affiche les valeur dans la fenetre pour modifié
 fracontact.Tag = numéroCompagnie
 txtNomCompagny.Text = lister.SelectedItem.SubItems(1)
 txtDate.Text = lister.SelectedItem.Text
 txtcontact.Text = lister.SelectedItem.SubItems(2)
 If lister.SelectedItem.SubItems(3) = "" Then
 cmbType.ListIndex = 0
 Else
 cmbType.Text = lister.SelectedItem.SubItems(3)
 End If
 txtcommentaire.Text = lister.SelectedItem.SubItems(4)
 
 'met en mode edition
 m_bModeAjouter = False
 Else
 'trouve le client pour afficher information
  Set rstVendeur = New ADODB.Recordset
 
  Call rstVendeur.Open("SELECT Grbvendeur.etat , Grbvendeur.no ,Grbvendeur.idclient , Grbvendeur.date ,Grbvendeur.contact, Grbvendeur.commentaire, Grbclient.nomclient FROM Grbvendeur inner join Grbclient on Grbvendeur.idclient = Grbclient.idclient WHERE [no] = " & lister.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
 
  fracontact.Visible = True
 
  fracontact.Tag = rstVendeur.Fields("idclient")
  txtDate.Text = rstVendeur.Fields("Date")
 txtNomCompagny.Text = rstVendeur.Fields("nomClient")
  txtcontact.Text = rstVendeur.Fields("Contact")
  txtcommentaire.Text = rstVendeur.Fields("commentaire")
 If IsNull(rstVendeur.Fields("Etat")) Then
 cmbType.ListIndex = 0
 Else
 cmbType.Text = rstVendeur.Fields("Etat")
 End If

 
 Call rstVendeur.Close
Set rstVendeur = Nothing
 End If
End If

Exit Sub

Oups:

wOups "frmvendeur", "lister_DblClick", Err, Err.number, Err.Description
End Sub

Private Sub lister_KeyDown(KeyCode As Integer, Shift As Integer)

 On Error GoTo Oups

 'Supprime l'enregistrement sélectionné
 If lister.ListItems.count > 0 Then
 If KeyCode = vbKeyDelete Then
 Call g_connData.Execute("DELETE * FROM GrbVendeur WHERE [no] = " & lister.SelectedItem.Tag)
 
 Call remplir_lister
 End If
 End If

 Exit Sub

Oups:

 wOups "frmvendeur", "lister_KeyDown", Err, Err.number, Err.Description
End Sub



Private Sub lstclient_DblClick()
10
 Dim iCompteur As Integer
 Dim rstclientinfo As ADODB.Recordset
 Dim rstVendeur As ADODB.Recordset

 
 fracontact.Visible = False
 If lstclient.ListItems.count > 0 Then
 Screen.MousePointer = vbHourglass

30


  cmbclient.Text = lstclient.SelectedItem.Text
75
 
 
  lstclient.Visible = False
 Set rstclientinfo = New ADODB.Recordset
 Call rstclientinfo.Open("Select * From Grbclient where nomclient= '" & lstclient.SelectedItem.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 lblAdresse.Caption = vbNullString
 lbltelephone.Caption = vbNullString
 If Not rstclientinfo.EOF Then numéroCompagnie = rstclientinfo.Fields("Idclient")
 'Ajoute les information de la base de donné au lbl correspondant
 Do While Not rstclientinfo.EOF
 'adresse
 If Not IsNull(rstclientinfo.Fields("AdresseLiv")) Then lblAdresse.Caption = lblAdresse.Caption + rstclientinfo.Fields("AdresseLiv")
 'Ville
 If Not IsNull(rstclientinfo.Fields("VilleLiv")) Then lblAdresse.Caption = lblAdresse.Caption + ", " + rstclientinfo.Fields("VilleLiv")
 'Pays
 If Not IsNull(rstclientinfo.Fields("PaysLiv")) Then lblAdresse.Caption = lblAdresse.Caption + ", " + rstclientinfo.Fields("PaysLiv")
 'Prov
 If Not IsNull(rstclientinfo.Fields("Prov/EtatLiv")) Then lblAdresse.Caption = lblAdresse.Caption + ", " + rstclientinfo.Fields("Prov/EtatLiv")
 'code postale
 If Not IsNull(rstclientinfo.Fields("CodePostalLiv")) Then lblAdresse.Caption = lblAdresse.Caption + ", " + rstclientinfo.Fields("CodePostalLiv")
 'Téléphone
 If Not rstclientinfo.Fields("Telephonne") = "" Then lbltelephone.Caption = lbltelephone.Caption + "Tél: " + rstclientinfo.Fields("Telephonne")
 'Fax
 If Not rstclientinfo.Fields("Fax") = "" Then lbltelephone.Caption = lbltelephone.Caption + " Fax: " + rstclientinfo.Fields("Fax")
 'Pagette
 If Not rstclientinfo.Fields("Pagette") = "" Then lbltelephone.Caption = lbltelephone.Caption + " Page: " + rstclientinfo.Fields("Pagette")
 'Cellulaire
 If Not rstclientinfo.Fields("Cellulaire") = "" Then lbltelephone.Caption = lbltelephone.Caption + " Cell: " + rstclientinfo.Fields("Cellulaire")
 'Email
 If Not rstclientinfo.Fields("Email") = "" Then lbltelephone.Caption = lbltelephone.Caption + " Email: " + rstclientinfo.Fields("Email")
 Call rstclientinfo.MoveNext
 Loop

 

 

 
 'vide lister
 Call lister.ListItems.Clear
 m_eModeCherche = MODE_CLIENT
 
 CmdAdd.Visible = True
 Dim itmvendeurrec As ListItem
 
 'ouvre la table pour client
 Set rstVendeur = New ADODB.Recordset
 
 Call rstVendeur.Open("SELECT Grbvendeur.etat , Grbvendeur.EnregPar, Grbvendeur.Contact, Grbvendeur.commentaire, Grbvendeur.no, Grbvendeur.Date FROM Grbvendeur INNER JOIN GrbClient ON GrbClient.IDClient = Grbvendeur.IDclient where nomclient= '" & lstclient.SelectedItem.Text & "' ORDER BY no", g_connData, adOpenDynamic, adLockOptimistic)
 
 'temp que pas a la fin de la table
 Do While Not rstVendeur.EOF
 'ajoute au lister
 Set itmvendeurrec = lister.ListItems.Add
 
 'no du client
 itmvendeurrec.Tag = rstVendeur.Fields("no")
 
 'vérifie les champs vide avant d'inséré
 'date
 If IsNull(rstVendeur.Fields("Date")) Then
 itmvendeurrec.Text = " "
 Else
 itmvendeurrec.Text = ConvertDate(DateSerial(Left(rstVendeur.Fields("Date"), 2), Mid(rstVendeur.Fields("Date"), 4, 2), Right(rstVendeur.Fields("Date"), 2)))
 End If
 '
 If IsNull(rstVendeur.Fields("Contact")) And IsNull(rstVendeur.Fields("commentaire")) And IsNull(rstVendeur.Fields("Date")) Then
 Call itmvendeurrec.ListSubItems.Add(, , vbNullString)
 Else
 Call itmvendeurrec.ListSubItems.Add(, , cmbclient.Text)
 End If
 'contact
 If IsNull(rstVendeur.Fields("Contact")) Then
 Call itmvendeurrec.ListSubItems.Add(, , vbNullString)
 Else
 Call itmvendeurrec.ListSubItems.Add(, , rstVendeur.Fields("Contact"))
 End If
 If IsNull(rstVendeur.Fields("etat")) Then
 Call itmvendeurrec.ListSubItems.Add(, , vbNullString)
 Else
 Call itmvendeurrec.ListSubItems.Add(, , rstVendeur.Fields("etat"))
 End If
 
 
 'commentaire
 If IsNull(rstVendeur.Fields("commentaire")) Then
 Call itmvendeurrec.ListSubItems.Add(, , vbNullString)
 Else
 Call itmvendeurrec.ListSubItems.Add(, , rstVendeur.Fields("commentaire"))
 End If
 
 If IsNull(rstVendeur.Fields("EnregPar")) Then
 Call itmvendeurrec.ListSubItems.Add(, , vbNullString)
 Else
 Call itmvendeurrec.ListSubItems.Add(, , "Enregpar")
 End If
 
 
 
 'prochaine enreg
 Call rstVendeur.MoveNext
 Loop
 
 'fermeture table et bd
 Call rstVendeur.Close
 Set rstVendeur = Nothing
 
 Call rstclientinfo.Close
 Set rstclientinfo = Nothing
 
 Screen.MousePointer = vbDefault
 End If
End Sub

Private Sub lstClient_LostFocus()
lstclient.Visible = False
End Sub

Private Sub mskDateCherche_GotFocus()

 On Error GoTo Oups

 If Len(mskDateCherche.Text) = 10 Then
 mskDateCherche.Text = Right$(mskDateCherche.Text, 8)
 End If
 
 mskDateCherche.mask = "##-##-##"

 Exit Sub

Oups:

 wOups "frmvendeur", "mskDateCherche_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskDateCherche_LostFocus()

 On Error GoTo Oups

 mskDateCherche.mask = vbNullString
 
 If mskDateCherche.Text = "__-__-__" Then
 mskDateCherche.Text = vbNullString
 Else
 If Len(mskDateCherche.Text) =   Then
 If IsDate(mskDateCherche.Text) Then
 mskDateCherche.Text = Year(DateSerial(Left$(mskDateCherche.Text, 2), Mid$(mskDateCherche.Text, 4, 2), Right$(mskDateCherche.Text, 2))) & Mid$(mskDateCherche.Text, 3, 8)
 End If
 End If
 End If

  Exit Sub

Oups:

  wOups "frmvendeur", "mskDateCherche_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub remplir_comboclient()

 On Error GoTo Oups

 '''''''''''''''''''''''''
 'rempli le combo client '
 '''''''''''''''''''''''''

 Dim rstClient As ADODB.Recordset

 'Set les tables
 Set rstClient = New ADODB.Recordset
 
 Call rstClient.Open("SELECT * FROM GrbClient WHERE Supprimé = False ORDER BY NomClient", g_connData, adOpenDynamic, adLockOptimistic)

 'Vide combo
 Call cmbclient.Clear
 If Not rstClient.EOF Then numéroCompagnie = rstClient.Fields("idclient")

 'Rempli les combo tant que pas fin d'enregistrement
 Do While Not rstClient.EOF
 Call cmbclient.AddItem(rstClient.Fields("NomClient"))
 
 cmbclient.ItemData(cmbclient.newIndex) = rstClient.Fields("IDClient")

 'Prochaine enreg
 Call rstClient.MoveNext
 Loop
 
 'Ferme table
 Call rstClient.Close
  Set rstClient = Nothing
 
  If cmbclient.ListCount > 0 Then
  cmbclient.ListIndex = 0
 
  End If

  Exit Sub

Oups:

  wOups "frmvendeur", "remplir_comboclient", Err, Err.number, Err.Description
End Sub
Private Sub remplir_lister_date()

 On Error GoTo Oups

 ''''''''''''''''''''
 ' rempli le lister '
 ''''''''''''''''''''
 Dim itmVendeur As ListItem
 Dim rstVendeur As ADODB.Recordset
 
 m_eModeCherche = MODE_DATE
 
 CmdAdd.Visible = False
 
 'vide lister
 Call lister.ListItems.Clear
 
 'ouvre la table pour client
 Set rstVendeur = New ADODB.Recordset
 
 Call rstVendeur.Open("SELECT Grbvendeur.no , Grbvendeur.date , Grbvendeur.etat , Grbvendeur.EnregPar , Grbvendeur.Contact , Grbvendeur.commentaire , Grbclient.NomClient FROM Grbclient Inner join Grbvendeur on Grbvendeur.IDClient = Grbclient.IDClient WHERE Grbvendeur.Date = '" & Right$(mskDateCherche.Text, 8) & "' ORDER BY Grbvendeur.no", g_connData, adOpenDynamic, adLockOptimistic)
 
 'temp que pas a la fin de la table
 Do While Not rstVendeur.EOF
 'ajoute au lister
 Set itmVendeur = lister.ListItems.Add
 
 'no du client
 itmVendeur.Tag = rstVendeur.Fields("no")
 
 'vérifie les champs vide avant d'inséré
 'date
  If IsNull(rstVendeur.Fields("Date")) Then
  itmVendeur.Text = vbNullString
  Else
  itmVendeur.Text = ConvertDate(DateSerial(Left(rstVendeur.Fields("Date"), 2), Mid(rstVendeur.Fields("Date"), 4, 2), Right(rstVendeur.Fields("Date"), 2)))
  End If
 'Nom Compagnie
 If IsNull(rstVendeur.Fields("nomclient")) Then
 Call itmVendeur.ListSubItems.Add(, , vbNullString)
 Else
 Call itmVendeur.ListSubItems.Add(, , rstVendeur.Fields("nomclient"))
 End If
 
 
 
 'contact
  If IsNull(rstVendeur.Fields("Contact")) Then
  Call itmVendeur.ListSubItems.Add(, , vbNullString)
  Else
 Call itmVendeur.ListSubItems.Add(, , rstVendeur.Fields("Contact"))
1 End If
 'État
 If IsNull(rstVendeur.Fields("etat")) Then
 Call itmVendeur.ListSubItems.Add(, , vbNullString)
 Else
 Call itmVendeur.ListSubItems.Add(, , rstVendeur.Fields("etat"))
 End If
 
 'commentaire
 If IsNull(rstVendeur.Fields("commentaire")) Then
 Call itmVendeur.ListSubItems.Add(, , vbNullString)
 Else
 Call itmVendeur.ListSubItems.Add(, , rstVendeur.Fields("commentaire"))
 End If
 If IsNull(rstVendeur.Fields("EnregPar")) Then
 Call itmVendeur.ListSubItems.Add(, , vbNullString)
 Else
 Call itmVendeur.ListSubItems.Add(, , rstVendeur.Fields("EnregPar"))
 End If
 
 'prochaine enreg
 Call rstVendeur.MoveNext
Loop
 
 'fermeture table et bd
Call rstVendeur.Close
Set rstVendeur = Nothing

Exit Sub

Oups:

1  wOups "frmvendeur", "remplir_lister_date", Err, Err.number, Err.Description
End Sub

