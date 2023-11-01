VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ListitemsView 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   7320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13485
   ForeColor       =   &H0000FF00&
   Icon            =   "ListitemsView.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   13485
   Begin VB.VScrollBar VScroll1 
      Height          =   2655
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6945
      Width           =   13485
      _ExtentX        =   23786
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14923
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   4657
            MinWidth        =   776
            Text            =   "Nombre d'enregistements: "
            TextSave        =   "Nombre d'enregistements: "
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   3625
            MinWidth        =   2893
            Text            =   "Exporter les données"
            TextSave        =   "Exporter les données"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6495
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   11456
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   16777152
      BackColor       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "ListitemsView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Table As String

Private Sub Form_Load()
 Me.Caption = Titre + " Table[" + Table + "]"
 SetLayeredWindow Me.hwnd, True
 SetLayeredWindowAttributes Me.hwnd, &HFFFFFF, 255 * 0.8, 1
 ListView1.ListItems.Clear
 On Error Resume Next
 Dim g_connData As New ADODB.Connection, rs As New ADODB.Recordset
 g_connData.Open "Driver={SQL Server};Server=TOUR-PATRICE\SQLEXPRESS;Database=WebGRB;Trusted_Connection=Yes;"
 rs.Open Table, g_connData, adOpenForwardOnly, adLockReadOnly
 LoadListViewFromRecordset ListView1, rs
 ListViewAdjustColumnWidth ListView1, True
 StatusBar1.Panels(1).Text = g_connData.Provider + "\" + g_connData.DefaultDatabase + "\" + Table
 StatusBar1.Panels(2).Text = ListView1.ListItems.count
 Me.VScroll1.Max = ListView1.ListItems.count + 1

End Sub
Private Sub Form_Resize()
If ScaleHeight > 400 Then 'protege en mode fenete minimal
 ListView1.Move 375, 0, ScaleWidth - 375, ScaleHeight - 375
 Else
 ListView1.Move 0, 0, ScaleWidth, ScaleHeight
 End If
 VScroll1.Height = ListView1.Height
End Sub
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
 If ColumnHeader.Text = "Freight" Then
 ListViewSortOnNonStringField ListView1, ColumnHeader.Index
 Exit Sub
 ElseIf ColumnHeader.Text = "OrderDate" Or ColumnHeader.Text = "RequiredDate" Then
 ListViewSortOnNonStringField ListView1, ColumnHeader.Index, , True
 Exit Sub
 End If
 If ListView1.Sorted And ColumnHeader.Index - 1 = ListView1.SortKey Then
 ListView1.SortOrder = 1 - ListView1.SortOrder
 Else
 ListView1.SortOrder = lvwAscending
 ListView1.SortKey = ColumnHeader.Index - 1
 End If
 ListView1.Sorted = True
End Sub
Public Function MakeRegion(picSkin As PictureBox) As Long
 Dim X As Long, Y As Long, StartLineX As Long
 Dim FullRegion As Long, LineRegion As Long
 Dim TransparentColor As Long
 Dim InFirstRegion As Boolean
 Dim InLine As Boolean
 Dim hDC As Long
 Dim PicWidth As Long
 Dim PicHeight As Long
 hDC = picSkin.hDC
 PicWidth = picSkin.ScaleWidth
 PicHeight = picSkin.ScaleHeight
 InFirstRegion = True: InLine = False
 X = Y = StartLineX = 0
 TransparentColor = GetPixel(hDC, 0, 0)
 For Y = 0 To PicHeight - 1
 For X = 0 To PicWidth - 1
 If GetPixel(hDC, X, Y) = TransparentColor Or X = PicWidth Then
 If InLine Then
 InLine = False
 LineRegion = CreateRectRgn(StartLineX, Y, X, Y + 1)
 If InFirstRegion Then
 FullRegion = LineRegion
 InFirstRegion = False
 Else
 CombineRgn FullRegion, FullRegion, LineRegion, RGN_OR
 DeleteObject LineRegion
 End If
 End If
 Else
 If Not InLine Then
 InLine = True
 StartLineX = X
 End If
 End If
 Next
 Next
 MakeRegion = FullRegion
End Function
Private Sub ListView1_DblClick()
 Dim li As ListItem
 Dim X As Integer, Y As Integer
 Dim str As String * 256
 Y = 0
 For Each li In ListView1.ListItems
 Y = Y + 1
 str = li.Text
 li.Bold = True
 Debug.Print li
 For X = 1 To li.ListSubItems.count
 Debug.Print " " + li.SubItems(X)
 Next
 Next
 MsgBox str, vbInformation, Me.Caption

End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)
 If Panel.Index = 3 Then frmLoginAdmin.Show
End Sub

Private Sub VScroll1_Change()
ListView1.ListItems(VScroll1.Value).Bold = True
End Sub
