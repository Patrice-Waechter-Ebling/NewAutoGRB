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
   ScaleHeight     =   6345
   ScaleWidth      =   8100
   StartUpPosition =   2  'CenterScreen
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
    On Error GoTo AfficherErreur
    Dim strName As String
    Dim Findfield As ADODB.Recordset
    
    Dim i As Integer
    FieldOk = False
    Set Findfield = New ADODB.Recordset
    Call Findfield.Open("Select * from Grb_Vendeur", g_connData, adOpenDynamic, adLockOptimistic)
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
    Call g_connData.Execute("ALTER TABLE GRB_Vendeur Add " & Name & " Text(25);")
    FieldOk = False
    Exit Sub
AfficherErreur:
    woups "frmvendeur", "FindFieldsExist()", Err, Erl
    End Sub




Private Sub remplir_lister()

5       On Error GoTo AfficherErreur

        '''''''''''''''''',
        'rempli le lister
        ''''''''''''''''''''',
10      Dim rstVendeur As ADODB.Recordset
15      Dim itmVendeur As ListItem

        Call FindFieldsExist("EnregPar")
        Call FindFieldsExist("Etat")
20      m_eModeCherche = MODE_CLIENT
  
25      CmdAdd.Visible = True
        
        'vide lister
30      Call lister.ListItems.Clear
  
        'ouvre la table pour client
35      Set rstVendeur = New ADODB.Recordset
        
40      Call rstVendeur.Open("SELECT * FROM GRB_vendeur WHERE IDClient = " & numéroCompagnie & " ORDER BY no", g_connData, adOpenDynamic, adLockOptimistic)
  
        'temp que pas a la fin de la table
45      Do While Not rstVendeur.EOF
          'ajoute au lister
50        Set itmVendeur = lister.ListItems.Add
        
          'no du client
55        itmVendeur.Tag = rstVendeur.Fields("no")
      
          'vérifie les champs vide avant d'inséré
          'date
60        If IsNull(rstVendeur.Fields("Date")) Then
65          itmVendeur.Text = " "
70        Else
75          itmVendeur.Text = ConvertDate(DateSerial(Left(rstVendeur.Fields("Date"), 2), Mid(rstVendeur.Fields("Date"), 4, 2), Right(rstVendeur.Fields("Date"), 2)))
80        End If
          '
          If IsNull(rstVendeur.Fields("Contact")) And IsNull(rstVendeur.Fields("commentaire")) And IsNull(rstVendeur.Fields("Date")) Then
            Call itmVendeur.ListSubItems.Add(, , vbNullString)
          Else
            Call itmVendeur.ListSubItems.Add(, , cmbclient.Text)
          End If
          
          
          
          
          
          
          'contact
85        If IsNull(rstVendeur.Fields("Contact")) Then
90          Call itmVendeur.ListSubItems.Add(, , vbNullString)
95        Else
100         Call itmVendeur.ListSubItems.Add(, , rstVendeur.Fields("Contact"))
105       End If

          'Type
          If IsNull(rstVendeur.Fields("Etat")) Then
             Call itmVendeur.ListSubItems.Add(, , vbNullString)
          Else
             Call itmVendeur.ListSubItems.Add(, , rstVendeur.Fields("Etat"))
          End If
          
        
          'commentaire
110       If IsNull(rstVendeur.Fields("commentaire")) Then
115         Call itmVendeur.ListSubItems.Add(, , vbNullString)
120       Else
125         Call itmVendeur.ListSubItems.Add(, , rstVendeur.Fields("commentaire"))
130       End If
          'Enregistrer par
          If IsNull(rstVendeur.Fields("EnregPar")) Then
            Call itmVendeur.ListSubItems.Add(, , vbNullString)
          Else
            Call itmVendeur.ListSubItems.Add(, , rstVendeur.Fields("EnregPar"))
          End If
          
          
          
          'prochaine enreg
135       Call rstVendeur.MoveNext
140     Loop
    
        'fermeture table et bd
145     Call rstVendeur.Close
150     Set rstVendeur = Nothing

155     Exit Sub

AfficherErreur:

160     woups "frmvendeur", "remplir_lister", Err, Erl
End Sub

Private Sub cmbclient_Click()

5       On Error GoTo AfficherErreur

        '''''''''''''''''''''''''''''
        'lorsque on select un client
        '''''''''''''''''''''''''''''
10      Dim rstClient As ADODB.Recordset

15      If cmbclient.ListIndex <> -1 Then
   
          'met visible fenetre pour ajouter
20        fracontact.Visible = False
    
          'mode ajouter ou editer
25        m_bModeAjouter = False
        
          'set le rapport
30        Set rstClient = New ADODB.Recordset
          
35        Call rstClient.Open("SELECT * FROM GRB_Client WHERE IDClient = " & cmbclient.ItemData(cmbclient.ListIndex) & " ORDER BY NomClient", g_connData, adOpenDynamic, adLockOptimistic)

          'initialise label adress et teleph
40        lblAdresse.Caption = vbNullString
45        lbltelephone.Caption = vbNullString
          
          numéroCompagnie = rstClient.Fields("idclient")
          'si client existe
          ''''''''''''''''''''''''''''''''''''''''''''
          'rempli adresse pays ville prov et codepostal si pas vide
          '''''''''''''''''''''''''''''''''''''''''''''
          'adresse
50        If Not rstClient.Fields("adresseliv") = "" Then
55          lblAdresse.Caption = lblAdresse.Caption + rstClient.Fields("adresseliv")
60        End If
    
          'ville
65        If Not rstClient.Fields("villeliv") = "" Then
70          lblAdresse.Caption = lblAdresse.Caption + ", " + rstClient.Fields("villeliv")
75        End If
    
          'pays
80        If Not rstClient.Fields("paysliv") = "" Then
85          lblAdresse.Caption = lblAdresse.Caption + ", " + rstClient.Fields("paysliv")
90        End If
    
         'province
95        If Not rstClient.Fields("prov/etatliv") = "" Then
100          lblAdresse.Caption = lblAdresse.Caption + ", " + rstClient.Fields("prov/etatliv")
105       End If
    
         'codepostal
110       If Not rstClient.Fields("codepostalliv") = "" Then
115         lblAdresse.Caption = lblAdresse.Caption + ", " + rstClient.Fields("codepostalliv")
120       End If
    
          ''''''''''''''''''''''''''''''''''
          'rempli tel fax pagette cell email si pas vide
          ''''''''''''''''''''''''''''''''''
          'telephone
125       If Not rstClient.Fields("telephonne") = "" Then
130         lbltelephone.Caption = lbltelephone.Caption + "TÉL: " + rstClient.Fields("telephonne")
135       End If
    
         'fax
140       If Not rstClient.Fields("fax") = "" Then
145         lbltelephone.Caption = lbltelephone.Caption + "      FAX: " + rstClient.Fields("fax")
150       End If
    
          'pagette
155       If Not rstClient.Fields("pagette") = "" Then
160         lbltelephone.Caption = lbltelephone.Caption + "      PAGE: " + rstClient.Fields("pagette")
165       End If
    
         'cellulaire
170       If Not rstClient.Fields("cellulaire") = "" Then
175         lbltelephone.Caption = lbltelephone.Caption + "      CELL: " + rstClient.Fields("cellulaire")
180       End If
    
          'email
185       If Not rstClient.Fields("email") = "" Then
190         lbltelephone.Caption = lbltelephone.Caption + "      EMAIL: " + rstClient.Fields("email")
195       End If
          txtNomCompagny.Text = rstClient.Fields("NomClient")
200       Call rstClient.Close
205       Set rstClient = Nothing
   
         'rempli le lister
          
210       Call remplir_lister
215     End If
        
220     Exit Sub

AfficherErreur:

225     woups "frmvendeur", "cmbclient_Click", Err, Erl
End Sub





Private Sub CmdAdd_Click()

5       On Error GoTo AfficherErreur

        '''''''''''''''''''''''''''''''''
        'ajoute un contact
        ''''''''''''''''''''''''''''''''
        'met visible fenetre pour ajouter
10      fracontact.Visible = True
15      fracontact.Tag = numéroCompagnie
        'mode ajouter ou editer
20      m_bModeAjouter = True

        'valeur par defaut sur l'ouverture
25      txtDate.Text = Year(Date) & "-" & Right$("0" & Month(Date), 2) & "-" & Right$("0" & Day(Date), 2)
        txtNomCompagny.Text = cmbclient.Text
30      txtcommentaire.Text = vbNullString
35      txtcontact.Text = vbNullString

40      Exit Sub

AfficherErreur:

45      woups "frmvendeur", "CmdAdd_Click", Err, Erl
End Sub

Private Sub cmdcherche_Click()

5       On Error GoTo AfficherErreur
        fracontact.Visible = False

10      Call remplir_lister_date

15      Exit Sub

AfficherErreur:

20      woups "frmvendeur", "cmdcherche_Click", Err, Erl
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
    
    With xlsheet.Range("A1:A3;C2:C3;E2:E3;G2:G3;I2:I3")
        .Font.Bold = True
        .HorizontalAlignment = xlRight
        .Font.SIZE = 11
    End With
    With xlsheet.Range("A5:I5")
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .Font.SIZE = 11
    End With
    
    
    
    
    
    Set info = New ADODB.Recordset
    Call info.Open("Select * From Grb_client where IDClient = " & numéroCompagnie, g_connData, adOpenDynamic, adLockOptimistic)
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
        xlsheet.Range("E" & row & ":H" & row).Merge
        row = row + 1
    Next
        xlsheet.Range("A:J").Columns.AutoFit
        info.Close
        Set info = Nothing
Else
    If lister.ListItems.count <= 0 Then Exit Sub
    row = 3
    xlsheet.Range("A1:D1").Merge
    xlsheet.Cells(1, 1) = "Liste des contacts en date du " & lister.ListItems(1).Text
    xlsheet.Cells(2, 1) = "Date:"
    xlsheet.Cells(2, 2) = "Nom de la Compagnie"
    xlsheet.Cells(2, 3) = "Nom du Contact"
    xlsheet.Cells(2, 4) = "État"
    xlsheet.Cells(2, 5) = "Commentaire"
    xlsheet.Cells(2, 6) = "Enregister Par"
    With xlsheet.Range("A1;A2;A2:F2")
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
        xlsheet.Range("A:F").Columns.AutoFit
End If




xlsheet.Visible = True



End Sub

Private Sub cmdfermercontact_Click()

5       On Error GoTo AfficherErreur

        ''''''''''''''''''''''
        'Quitte liste contact'
        ''''''''''''''''''''''
        'cache fenêtre
10      fracontact.Visible = False
15     If m_eModeCherche = MODE_CLIENT Then
20      Call remplir_lister
25    Else
30       Call remplir_lister_date
35    End If

40     Exit Sub

AfficherErreur:

45      woups "frmvendeur", "cmdfermercontact_Click", Err, Erl
End Sub

Private Sub cmdquit_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmvendeur", "cmdquit_Click", Err, Erl
End Sub

Private Sub cmdrechercheclient_Click()

5       On Error GoTo AfficherErreur

10      Dim rstcatalog As ADODB.Recordset
15      Dim sDescription   As String
20      Dim itmDescription As ListItem
  
25      sDescription = InputBox("Quelle est la description à rechercher")
  
30      If sDescription <> vbNullString Then
35        Call lstclient.ListItems.Clear
  
40        sDescription = Replace(sDescription, "'", "''")
  
45        sDescription = "%" & sDescription & "%"

50        Set rstcatalog = New ADODB.Recordset

55
    Call rstcatalog.Open("SELECT DISTINCT NomClient FROM GRB_Client WHERE  NomClient LIKE '" & sDescription & "'  ORDER BY NomClient", g_connData, adOpenDynamic, adLockOptimistic)
60        Do While Not rstcatalog.EOF
65          Set itmDescription = lstclient.ListItems.Add()
        
70          itmDescription.Tag = rstcatalog.Fields("NomClient")
            itmDescription.Text = rstcatalog.Fields("NomClient")

155         Call rstcatalog.MoveNext
160       Loop
    
165       Call rstcatalog.Close
170       Set rstcatalog = Nothing

175       If lstclient.ListItems.count > 0 Then
180         lstclient.Visible = True

185         Call lstclient.SetFocus
190       Else
195         Call MsgBox("Aucun enregistrement trouvé!")
200       End If
205     End If

210     Exit Sub

AfficherErreur:

215     woups "frmvendeur", "cmdrechercheclient_Click", Err, Erl


End Sub

Private Sub cmdsave_Click()

5       On Error GoTo AfficherErreur

10      Dim rstVendeur As ADODB.Recordset
  
        'table vendeur ouvert
15      Set rstVendeur = New ADODB.Recordset
        Call FindFieldsExist("Enregistrerpar")
        Call FindFieldsExist("Type")
        
20      If m_bModeAjouter = True Then
25        Call rstVendeur.Open("SELECT * FROM GRB_vendeur", g_connData, adOpenDynamic, adLockOptimistic)
      
30        Call rstVendeur.AddNew
      
35        m_bModeAjouter = False
40      Else
45        Call rstVendeur.Open("SELECT * FROM GRB_Vendeur WHERE [no] = " & lister.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
50      End If
        
        
        
        
55      rstVendeur.Fields("IDClient").Value = fracontact.Tag
60      rstVendeur.Fields("Date").Value = Right$(txtDate.Text, 8)
65      rstVendeur.Fields("Contact").Value = txtcontact.Text
70      rstVendeur.Fields("commentaire").Value = txtcommentaire.Text
        rstVendeur.Fields("EnregPar").Value = CStr(g_sUserID)
        rstVendeur.Fields("Etat").Value = CStr(cmbType.Text)
        
75      Call rstVendeur.Update
         
        'ferme la table
80      Call rstVendeur.Close
85      Set rstVendeur = Nothing
          
90      fracontact.Visible = False
      
        'rempli le lister
95      If m_eModeCherche = MODE_CLIENT Then
100       Call remplir_lister
105     Else
110       Call remplir_lister_date
115     End If

120     Exit Sub

AfficherErreur:

125     woups "frmvendeur", "cmdsave_Click", Err, Erl
End Sub

Private Sub CmdSupp_Click()

5       On Error GoTo AfficherErreur

        'Supprime l'enregistrement sélectionné
10      If lister.ListItems.count > 0 Then
15        Call g_connData.Execute("DELETE * FROM GRB_Vendeur WHERE [no] = " & lister.SelectedItem.Tag)
      
20        Call remplir_lister
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmvendeur", "CmdSupp_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

        'initialise les labels adresse et telephone
10      lblAdresse.Caption = vbNullString
15      lbltelephone.Caption = vbNullString

        cmbType.AddItem ("Piste de vente")
            cmbType.AddItem ("Opportunité")
            cmbType.AddItem ("Soumission")
            cmbType.AddItem ("Gagner")
            cmbType.AddItem ("Perdu")
            cmbType.AddItem ("En attente")
         cmbType.ListIndex = 0
         

        'rempli le comboclient et lister
20      Call remplir_comboclient

25      Screen.MousePointer = vbDefault

30      Exit Sub

AfficherErreur:

35      woups "frmvendeur", "Form_Load", Err, Erl
End Sub

Private Sub lister_DblClick()

5       On Error GoTo AfficherErreur

        'Sur DblClick, affiche fenêtre pour modifié l'enreg selectionné dans lister
10      Dim rstVendeur As ADODB.Recordset

15      If lister.ListItems.count <> 0 Then
20        If m_eModeCherche = MODE_CLIENT Then
            'si lister pas vide
            'met fenetre visible
25          fracontact.Visible = True
        
            'affiche les valeur dans la fenetre pour modifié
30          fracontact.Tag = numéroCompagnie
            txtNomCompagny.Text = lister.SelectedItem.SubItems(1)
35          txtDate.Text = lister.SelectedItem.Text
40          txtcontact.Text = lister.SelectedItem.SubItems(2)
            If lister.SelectedItem.SubItems(3) = "" Then
                cmbType.ListIndex = 0
            Else
                cmbType.Text = lister.SelectedItem.SubItems(3)
45          End If
            txtcommentaire.Text = lister.SelectedItem.SubItems(4)
        
            'met en mode edition
50          m_bModeAjouter = False
55        Else
            'trouve le client pour afficher information
60          Set rstVendeur = New ADODB.Recordset
          
65          Call rstVendeur.Open("SELECT grb_vendeur.etat , grb_vendeur.no ,grb_vendeur.idclient , grb_vendeur.date ,grb_vendeur.contact, grb_vendeur.commentaire, grb_client.nomclient FROM GRB_vendeur inner join grb_client on grb_vendeur.idclient = grb_client.idclient WHERE [no] = " & lister.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
        
70          fracontact.Visible = True
    
75          fracontact.Tag = rstVendeur.Fields("idclient")
80          txtDate.Text = rstVendeur.Fields("Date")
            txtNomCompagny.Text = rstVendeur.Fields("nomClient")
85          txtcontact.Text = rstVendeur.Fields("Contact")
90          txtcommentaire.Text = rstVendeur.Fields("commentaire")
            If IsNull(rstVendeur.Fields("Etat")) Then
                cmbType.ListIndex = 0
            Else
                cmbType.Text = rstVendeur.Fields("Etat")
            End If

   
100         Call rstVendeur.Close
105         Set rstVendeur = Nothing
110       End If
115     End If

120     Exit Sub

AfficherErreur:

125     woups "frmvendeur", "lister_DblClick", Err, Erl
End Sub

Private Sub lister_KeyDown(KeyCode As Integer, Shift As Integer)

5       On Error GoTo AfficherErreur

        'Supprime l'enregistrement sélectionné
10      If lister.ListItems.count > 0 Then
15        If KeyCode = vbKeyDelete Then
20          Call g_connData.Execute("DELETE * FROM GRB_Vendeur WHERE [no] = " & lister.SelectedItem.Tag)
      
25          Call remplir_lister
30        End If
35      End If

40      Exit Sub

AfficherErreur:

45      woups "frmvendeur", "lister_KeyDown", Err, Erl
End Sub



Private Sub lstclient_DblClick()
10
15      Dim iCompteur      As Integer
        Dim rstclientinfo As ADODB.Recordset
        Dim rstVendeur As ADODB.Recordset

        
        fracontact.Visible = False
20      If lstclient.ListItems.count > 0 Then
25        Screen.MousePointer = vbHourglass

30


65          cmbclient.Text = lstclient.SelectedItem.Text
75
            
            
80        lstclient.Visible = False
        Set rstclientinfo = New ADODB.Recordset
        Call rstclientinfo.Open("Select * From Grb_client where nomclient= '" & lstclient.SelectedItem.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
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
        
     Call rstVendeur.Open("SELECT grb_vendeur.etat , grb_vendeur.EnregPar, grb_vendeur.Contact, grb_vendeur.commentaire, grb_vendeur.no, grb_vendeur.Date FROM GRB_vendeur INNER JOIN GRB_Client ON Grb_Client.IDClient = grb_vendeur.IDclient where nomclient= '" & lstclient.SelectedItem.Text & "' ORDER BY no", g_connData, adOpenDynamic, adLockOptimistic)
  
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

5       On Error GoTo AfficherErreur

10      If Len(mskDateCherche.Text) = 10 Then
15        mskDateCherche.Text = Right$(mskDateCherche.Text, 8)
20      End If
  
25      mskDateCherche.mask = "##-##-##"

30      Exit Sub

AfficherErreur:

35      woups "frmvendeur", "mskDateCherche_GotFocus", Err, Erl
End Sub

Private Sub mskDateCherche_LostFocus()

5       On Error GoTo AfficherErreur

10      mskDateCherche.mask = vbNullString
  
15      If mskDateCherche.Text = "__-__-__" Then
20        mskDateCherche.Text = vbNullString
25      Else
30        If Len(mskDateCherche.Text) = 8 Then
35          If IsDate(mskDateCherche.Text) Then
40            mskDateCherche.Text = Year(DateSerial(Left$(mskDateCherche.Text, 2), Mid$(mskDateCherche.Text, 4, 2), Right$(mskDateCherche.Text, 2))) & Mid$(mskDateCherche.Text, 3, 8)
45          End If
50        End If
55      End If

60      Exit Sub

AfficherErreur:

65      woups "frmvendeur", "mskDateCherche_LostFocus", Err, Erl
End Sub

Private Sub remplir_comboclient()

5       On Error GoTo AfficherErreur

        '''''''''''''''''''''''''
        'rempli le combo client '
        '''''''''''''''''''''''''

10      Dim rstClient As ADODB.Recordset

        'Set les tables
15      Set rstClient = New ADODB.Recordset
        
20      Call rstClient.Open("SELECT * FROM GRB_Client WHERE Supprimé = False ORDER BY NomClient", g_connData, adOpenDynamic, adLockOptimistic)

        'Vide combo
25      Call cmbclient.Clear
        If Not rstClient.EOF Then numéroCompagnie = rstClient.Fields("idclient")

        'Rempli les combo tant que pas fin d'enregistrement
30      Do While Not rstClient.EOF
35        Call cmbclient.AddItem(rstClient.Fields("NomClient"))
      
40        cmbclient.ItemData(cmbclient.newIndex) = rstClient.Fields("IDClient")

          'Prochaine enreg
45        Call rstClient.MoveNext
50      Loop
  
        'Ferme table
55      Call rstClient.Close
60      Set rstClient = Nothing
  
65      If cmbclient.ListCount > 0 Then
70        cmbclient.ListIndex = 0
    
75      End If

80      Exit Sub

AfficherErreur:

85      woups "frmvendeur", "remplir_comboclient", Err, Erl
End Sub
Private Sub remplir_lister_date()

5       On Error GoTo AfficherErreur

        ''''''''''''''''''''
        ' rempli le lister '
        ''''''''''''''''''''
10      Dim itmVendeur As ListItem
15      Dim rstVendeur As ADODB.Recordset
  
20      m_eModeCherche = MODE_DATE
  
25      CmdAdd.Visible = False
  
        'vide lister
30      Call lister.ListItems.Clear
  
        'ouvre la table pour client
35      Set rstVendeur = New ADODB.Recordset
        
40      Call rstVendeur.Open("SELECT grb_vendeur.no , grb_vendeur.date , grb_vendeur.etat , grb_vendeur.EnregPar , grb_vendeur.Contact , grb_vendeur.commentaire , grb_client.NomClient FROM GRB_client Inner join grb_vendeur on grb_vendeur.IDClient = grb_client.IDClient WHERE grb_vendeur.Date = '" & Right$(mskDateCherche.Text, 8) & "' ORDER BY grb_vendeur.no", g_connData, adOpenDynamic, adLockOptimistic)
   
        'temp que pas a la fin de la table
45      Do While Not rstVendeur.EOF
          'ajoute au lister
50        Set itmVendeur = lister.ListItems.Add
        
          'no du client
55        itmVendeur.Tag = rstVendeur.Fields("no")
          
          'vérifie les champs vide avant d'inséré
          'date
60        If IsNull(rstVendeur.Fields("Date")) Then
65          itmVendeur.Text = vbNullString
70        Else
75          itmVendeur.Text = ConvertDate(DateSerial(Left(rstVendeur.Fields("Date"), 2), Mid(rstVendeur.Fields("Date"), 4, 2), Right(rstVendeur.Fields("Date"), 2)))
80        End If
          'Nom Compagnie
          If IsNull(rstVendeur.Fields("nomclient")) Then
            Call itmVendeur.ListSubItems.Add(, , vbNullString)
          Else
            Call itmVendeur.ListSubItems.Add(, , rstVendeur.Fields("nomclient"))
          End If
          
          
          
          'contact
85        If IsNull(rstVendeur.Fields("Contact")) Then
90          Call itmVendeur.ListSubItems.Add(, , vbNullString)
95        Else
100         Call itmVendeur.ListSubItems.Add(, , rstVendeur.Fields("Contact"))
105       End If
          'État
          If IsNull(rstVendeur.Fields("etat")) Then
            Call itmVendeur.ListSubItems.Add(, , vbNullString)
          Else
            Call itmVendeur.ListSubItems.Add(, , rstVendeur.Fields("etat"))
          End If
          
          'commentaire
110       If IsNull(rstVendeur.Fields("commentaire")) Then
115         Call itmVendeur.ListSubItems.Add(, , vbNullString)
120       Else
125         Call itmVendeur.ListSubItems.Add(, , rstVendeur.Fields("commentaire"))
130       End If
          If IsNull(rstVendeur.Fields("EnregPar")) Then
            Call itmVendeur.ListSubItems.Add(, , vbNullString)
          Else
            Call itmVendeur.ListSubItems.Add(, , rstVendeur.Fields("EnregPar"))
          End If
          
          'prochaine enreg
135       Call rstVendeur.MoveNext
140     Loop
  
        'fermeture table et bd
145     Call rstVendeur.Close
150     Set rstVendeur = Nothing

155     Exit Sub

AfficherErreur:

160     woups "frmvendeur", "remplir_lister_date", Err, Erl
End Sub

