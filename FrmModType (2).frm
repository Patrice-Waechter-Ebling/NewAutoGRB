VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmModType 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modifier les Types "
   ClientHeight    =   4305
   ClientLeft      =   8640
   ClientTop       =   6735
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrmADD 
      Caption         =   "Ajouter"
      Height          =   735
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   3615
      Begin VB.TextBox txtAdd 
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Text            =   "Entré le Nom Ici"
         Top             =   240
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmModType.frx":0000
         Left            =   240
         List            =   "FrmModType.frx":000A
         TabIndex        =   10
         Text            =   "E"
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Afficher 
      BackColor       =   &H00000000&
      Caption         =   "Afficher"
      ForeColor       =   &H8000000F&
      Height          =   1935
      Left            =   3840
      TabIndex        =   0
      Top             =   2160
      Width           =   1455
      Begin VB.OptionButton OptElec 
         BackColor       =   &H00000000&
         Caption         =   "Électrique"
         ForeColor       =   &H8000000F&
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton OptMec 
         BackColor       =   &H00000000&
         Caption         =   "Mecanique"
         ForeColor       =   &H8000000F&
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Opttous 
         BackColor       =   &H00000000&
         Caption         =   "Tous"
         ForeColor       =   &H8000000F&
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      ForeColor       =   &H8000000F&
      Height          =   2115
      Left            =   3840
      TabIndex        =   5
      Top             =   120
      Width           =   1455
      Begin VB.CommandButton Cmdfermer 
         Caption         =   "Fermer"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton CmdValider 
         Caption         =   "Valider"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdsupprimer 
         Caption         =   "Supprimer"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "Annuler"
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Cmdajouter 
         Caption         =   "Ajouter"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView lsttype 
      Height          =   3855
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "E/M"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   4939
      EndProperty
   End
End
Attribute VB_Name = "FrmModType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Cmdajouter_Click()
FrmADD.Visible = True
Cmdajouter.Visible = False
CmdValider.Visible = True
cmdsupprimer.Visible = False
CmdCancel.Visible = True
cmdsupprimer.Enabled = False

End Sub

Private Sub cmdCancel_Click()
FrmADD.Visible = False
Cmdajouter.Visible = True
CmdValider.Visible = False
cmdsupprimer.Visible = True
CmdCancel.Visible = False
End Sub

Private Sub Cmdfermer_Click()
Call Unload(Me)

End Sub

Private Sub cmdsupprimer_Click()
Dim strtest As String
Dim Name As String
Name = lsttype.SelectedItem.SubItems(1)
Name = Replace(Name, "'", "''")


strtest = "DELETE * FROM TBL_Punch_Type WHERE mode = '" & lsttype.SelectedItem.Text & "' And name = '" & Name & " ' "
Call g_connData.Execute(strtest)
If Opttous.Value = True Then Call Opttous_Click
If OptMec.Value = True Then Call OptMec_Click
If OptElec.Value = True Then Call OptElec_Click

cmdsupprimer.Enabled = False
End Sub

Private Sub CmdValider_Click()
Dim strtest As String
txtAdd.Text = Replace(txtAdd.Text, "'", "''")
strtest = "Insert into TBL_Punch_Type (mode, Name) Values ('" & Combo1.Text & "','" & txtAdd.Text & "');"
Call g_connData.Execute(strtest)
CmdValider.Visible = False
Cmdajouter.Visible = True
FrmADD.Visible = False
CmdCancel.Visible = False
cmdsupprimer.Visible = True
If Opttous.Value = True Then Call Opttous_Click
If OptMec.Value = True Then Call OptMec_Click
If OptElec.Value = True Then Call OptElec_Click



End Sub

Private Sub Form_Load()
Dim tbltype As ADODB.Recordset
Dim LIST As ListItem
Set tbltype = New ADODB.Recordset
Call tbltype.Open("Select * from Tbl_punch_Type Order by name ", g_connData, adOpenDynamic, adLockOptimistic)

Do While Not tbltype.EOF
 
 Set LIST = lsttype.ListItems.Add()
 LIST.Text = tbltype.Fields("mode")
 Call LIST.ListSubItems.Add(, , tbltype.Fields("name"))
 
 
 Call tbltype.MoveNext
Loop
Call tbltype.Close
Set tbltype = Nothing



End Sub

Private Sub Form_Unload(Cancel As Integer)
Call frmFeuilleTemps.RemplirComboType
End Sub

Private Sub lsttype_Click()
cmdsupprimer.Enabled = True



End Sub

Private Sub OptElec_Click()
OptMec.Value = False
cmdsupprimer.Enabled = False

Opttous.Value = False
Dim tbltype As ADODB.Recordset
Dim LIST As ListItem
Set tbltype = New ADODB.Recordset
lsttype.ListItems.Clear
Call tbltype.Open("Select * from Tbl_punch_Type where mode = 'E' Order by name ", g_connData, adOpenDynamic, adLockOptimistic)

Do While Not tbltype.EOF
 
 Set LIST = lsttype.ListItems.Add()
 LIST.Text = tbltype.Fields("mode")
 Call LIST.ListSubItems.Add(, , tbltype.Fields("name"))
 
 
 Call tbltype.MoveNext
Loop
Call tbltype.Close
Set tbltype = Nothing
End Sub

Private Sub OptMec_Click()
Opttous.Value = False
cmdsupprimer.Enabled = False
OptElec.Value = False
Dim tbltype As ADODB.Recordset
Dim LIST As ListItem
Set tbltype = New ADODB.Recordset
lsttype.ListItems.Clear
Call tbltype.Open("Select * from Tbl_punch_Type where mode = 'M' Order by name ", g_connData, adOpenDynamic, adLockOptimistic)

Do While Not tbltype.EOF
 
 Set LIST = lsttype.ListItems.Add()
 LIST.Text = tbltype.Fields("mode")
 Call LIST.ListSubItems.Add(, , tbltype.Fields("name"))
 
 
 Call tbltype.MoveNext
Loop
Call tbltype.Close
Set tbltype = Nothing
End Sub


Private Sub Opttous_Click()


OptMec.Value = False
OptElec.Value = False
cmdsupprimer.Enabled = False
Dim tbltype As ADODB.Recordset
Dim LIST As ListItem
Set tbltype = New ADODB.Recordset
lsttype.ListItems.Clear
Call tbltype.Open("Select * from Tbl_punch_Type Order by name ", g_connData, adOpenDynamic, adLockOptimistic)

Do While Not tbltype.EOF
 
 Set LIST = lsttype.ListItems.Add()
 LIST.Text = tbltype.Fields("mode")
 Call LIST.ListSubItems.Add(, , tbltype.Fields("name"))
 
 
 Call tbltype.MoveNext
Loop
Call tbltype.Close
Set tbltype = Nothing
End Sub
