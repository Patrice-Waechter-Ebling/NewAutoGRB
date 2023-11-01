VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAjoutPlusieursItemMec 
   Caption         =   "Form1"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNouveau 
      Caption         =   "Nouveau"
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      Top             =   5040
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   8281
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Test"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Test2"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmAjoutPlusieursItemMec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNouveau_Click()
  Call ListView1.ListItems.Add
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Set ListView1.DropHighlight = ListView1.HitTest(x, y)

  If Not ListView1.DropHighlight Is Nothing Then
    If x >= ListView1.ColumnHeaders(1).Left And x <= ListView1.ColumnHeaders(1).Left + ListView1.ColumnHeaders(1).Width Then
      Call MsgBox(ListView1.ColumnHeaders(1).Text)
    Else
      If x >= ListView1.ColumnHeaders(2).Left And x <= ListView1.ColumnHeaders(2).Left + ListView1.ColumnHeaders(2).Width Then
        Call MsgBox(ListView1.ColumnHeaders(2).Text)
      End If
    End If
  End If
End Sub
