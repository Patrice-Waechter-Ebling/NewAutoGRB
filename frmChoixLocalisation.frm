VERSION 5.00
Begin VB.Form frmChoixLocalisation 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChoixLocalisation.frx":0000
   ScaleHeight     =   2400
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbLocalisation 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   3615
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblQuestion 
      BackStyle       =   0  'Transparent
      Caption         =   "Dans quelle localisation ?"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3615
   End
End
Attribute VB_Name = "frmChoixLocalisation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_eCatalogue As enumCatalogue

Public Sub Afficher(ByVal eCatalogue As enumCatalogue, ByVal sNoPiece As String)
  
5       On Error GoTo AfficherErreur

10      m_eCatalogue = eCatalogue

15      lblQuestion.Caption = "Quelle est la localisation de la pièce " & sNoPiece & "?"

20      Call RemplirComboLocalisation

25      Call Me.Show(vbModal)

30      Exit Sub

AfficherErreur:
  
35      woups "frmChoixLocalisation", "Afficher", Err, Erl
End Sub

Private Sub cmdOK_Click()

5       On Error GoTo AfficherErreur

10      g_sLocalisation = cmbLocalisation.Text

15      Call Unload(Me)

20      Exit Sub

AfficherErreur:

25      woups "frmChoixLocalisation", "cmdOK_Click", Err, Erl
End Sub

Private Sub RemplirComboLocalisation()

5       On Error GoTo AfficherErreur

        'Remplir le combo des localisations
10      Dim rstInv As ADODB.Recordset
  
        'Il faut vider le combo avant de le remplir
15      Call cmbLocalisation.Clear

20      Call cmbLocalisation.AddItem("")

25      Set rstInv = New ADODB.Recordset
      
30      If m_eCatalogue = ELECTRIQUE Then
35        Call rstInv.Open("SELECT DISTINCT Localisation FROM GRB_InventaireElec", g_connData, adOpenDynamic, adLockOptimistic)
40      Else
45        Call rstInv.Open("SELECT DISTINCT Localisation FROM GRB_InventaireMec", g_connData, adOpenDynamic, adLockOptimistic)
50      End If

        'Tant que ce n'est pas la fin des enregistrements
55      Do While Not rstInv.EOF
60        If Trim(rstInv.Fields("Localisation")) <> "" Then
65          Call cmbLocalisation.AddItem(rstInv.Fields("Localisation"))
70        End If
    
75        Call rstInv.MoveNext
80      Loop
  
85      Call rstInv.Close
90      Set rstInv = Nothing
  
95      Exit Sub

AfficherErreur:

100     woups "frmChoixLocalisation", "RemplirComboContactFRS", Err, Erl
End Sub

