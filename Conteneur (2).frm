VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm Conteneur 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H00000000&
   Caption         =   "NewAutoGRB"
   ClientHeight    =   11205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19515
   Icon            =   "Conteneur.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imgMenu 
      Left            =   1800
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Conteneur.frx":3F02
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Conteneur.frx":62B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Conteneur.frx":6706
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Conteneur.frx":6B58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Conteneur.frx":8EDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Conteneur.frx":91F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Conteneur.frx":9646
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Conteneur.frx":BDF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Conteneur.frx":CCD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Conteneur.frx":D124
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Conteneur.frx":D576
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Conteneur.frx":D9C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Conteneur.frx":DCE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Conteneur.frx":E5BC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      Picture         =   "Conteneur.frx":EE96
      ScaleHeight     =   375
      ScaleWidth      =   19515
      TabIndex        =   2
      Top             =   420
      Width           =   19515
      Begin VB.Label lbldb 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Base de donné:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   9000
         TabIndex        =   5
         Top             =   0
         Width           =   1425
      End
      Begin VB.Label lblDerniereVersion 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Dernière Version : "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   2760
         TabIndex        =   4
         Top             =   0
         Width           =   1740
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Version "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   6840
         TabIndex        =   3
         Top             =   0
         Width           =   765
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   10590
      Width           =   19515
      _ExtentX        =   34422
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   12
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   22016
            MinWidth        =   1590
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1376
            MinWidth        =   776
            Text            =   "Utilisateur"
            TextSave        =   "Utilisateur"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   1376
            MinWidth        =   776
            Text            =   "username"
            TextSave        =   "username"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   776
            MinWidth        =   776
            Text            =   "ID"
            TextSave        =   "ID"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   1085
            MinWidth        =   776
            Text            =   "Groupe"
            TextSave        =   "Groupe"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   1005
            MinWidth        =   776
            Text            =   "Famille"
            TextSave        =   "Famille"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1693
            MinWidth        =   953
            TextSave        =   "18/10/2023"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   953
            MinWidth        =   953
            TextSave        =   "22:26"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   767
            MinWidth        =   776
            TextSave        =   "Maj"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   776
            MinWidth        =   776
            TextSave        =   "Num"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   7
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   926
            MinWidth        =   776
            TextSave        =   "KANA"
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Alignment       =   1
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   884
            MinWidth        =   884
            TextSave        =   "DÉFIL"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Conteneur.frx":FB38
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Conteneur.frx":FF8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Conteneur.frx":10C64
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   10845
      Width           =   19515
      _ExtentX        =   34422
      _ExtentY        =   635
      ButtonWidth     =   2249
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Connexion"
            ImageIndex      =   1
            Style           =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Àpropos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Quitter"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Tables"
            ImageIndex      =   1
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar MenuPrincipal 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   19515
      _ExtentX        =   34422
      _ExtentY        =   741
      ButtonWidth     =   2937
      ButtonHeight    =   582
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "imgMenu"
      DisabledImageList=   "imgMenu"
      HotImageList    =   "imgMenu"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&Clients"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&Fournisseurs"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Co&ntacts"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&Employés"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&Rapports"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&Punch"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Cé&dule"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "cVendeurs"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Newletter"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Importer MDB"
            Object.ToolTipText     =   "Requiere le niveau 99"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&Inventaire"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "C&atalogue"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Soumissions"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Confi&guration"
            Object.ToolTipText     =   "Requiere le rang de groupe 2"
            ImageIndex      =   14
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Conteneur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()
On Error GoTo Oups
 Dim g_connData As New ADODB.Connection
 Dim rs As New ADODB.Recordset
 Dim fld As ADODB.Field, alignment As Integer
 Dim recCount As Long, i As Long, fldName As String
 g_connData.Open "Driver={SQL Server};Server=TOUR-PATRICE\SQLEXPRESS;Database=WebGRB;Trusted_Connection=Yes;"
 If g_iNoGroupe = 2 Or g_iNoGroupe = 24 Or g_iNoGroupe = 2 Then
 lbldb.Caption = "Base de données:Actuelle" 'ajoute l'information de quel base de donné on active GLL
 lbldb.Visible = True
 g_admin = True
 MenuPrincipal.Buttons(10).Visible = True
 Else
 lbldb.Visible = False
 MenuPrincipal.Buttons(10).Visible = False
 g_admin = False
 End If
 If TesterVersion() = False Then MsgBox "Votre logiciel en corespond pas la version en usage", vbInformation + vbOKOnly, Titre
 rs.Open "[@_menu]", g_connData, adOpenForwardOnly, adLockReadOnly
 rs.MoveFirst
 Do Until rs.EOF
 recCount = recCount + 1
 Toolbar1.Buttons(4).ButtonMenus.Add , , rs.Fields("NormalizedName")
 If recCount = MaxRecords Then Exit Do
 rs.MoveNext
 Loop
 Exit Sub
Oups:
 MsgBox Err.Description + vbCrLf + Err.Source, vbCritical, Me.Caption
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Index
 Case 1: frmLogin.Show
 Case 2: ShellAbout Me.hwnd, Titre, "Création Patrice Waechter-Ebling 2023", Me.Icon
 Case 3: If MsgBox("Voulez vous vraiment quitter", vbDefaultButton2 + vbYesNo, Titre) = vbYes Then End
 Case 4: MsgBox Button.ButtonMenus(4).Text
 
 End Select
End Sub
Private Function FermerTousLesForms() As Boolean
 On Error GoTo Oups
 Dim objForm As Form
 Dim bFermer As Boolean
 bFermer = True
 For Each objForm In Forms
 If objForm.Name <> Me.Name Then
 If UCase(objForm.Name) = "FRMPROJSOUMELEC" Or UCase(objForm.Name) = "FRMPROJSOUMMEC" Then
 bFermer = objForm.PeutFermer
 Exit For
 End If
 End If
 Next
 If bFermer = True Then
 For Each objForm In Forms
 If objForm.Name <> Me.Name Then
 Call Unload(objForm)
 End If
 Next
 End If
 FermerTousLesForms = bFermer
 Exit Function
Oups:
 wOups "frmDispatch", "FermerTousLesForms", Err, Err.number, Err.Description
End Function
