VERSION 5.00
Begin VB.Form frmChoixContactFRS 
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
   Picture         =   "frmChoixContactFRS.frx":0000
   ScaleHeight     =   2400
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.ComboBox cmbContactFRS 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "cmbContactFRS"
      Top             =   1440
      Width           =   3615
   End
   Begin VB.Label lblQuestion 
      BackStyle       =   0  'Transparent
      Caption         =   "Qui est le contact pour ... ?"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3615
   End
End
Attribute VB_Name = "frmChoixContactFRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_iNoFRS As Integer

Private Sub cmdAnnuler_Click()

5       On Error GoTo AfficherErreur

10      frmChoixDemande.m_bAnnulerContact = True

15      Call Unload(Me)

20      Exit Sub

AfficherErreur:

25      woups "frmChoixContactFRS", "cmdAnnuler_Click", Err, Erl
End Sub

Private Sub cmdOK_Click()

5       On Error GoTo AfficherErreur

10      frmChoixDemande.m_bAnnulerContact = False

15      frmChoixDemande.m_sContact = cmbContactFRS.Text
  
20      Call Unload(Me)

25      Exit Sub

AfficherErreur:

30      woups "frmChoixContactFRS", "cmdOK_Click", Err, Erl
End Sub

Private Sub RemplirComboContactFRS()

5       On Error GoTo AfficherErreur

        'Remplir le combo des contacts
10      Dim rstContactFRS As ADODB.Recordset
15      Dim rstContact    As ADODB.Recordset
  
        'Il faut vider le combo avant de le remplir
20      Call cmbContactFRS.Clear
      
25      Set rstContactFRS = New ADODB.Recordset
30      Set rstContact = New ADODB.Recordset
      
35      Call rstContactFRS.Open("SELECT * FROM GRB_ContactFRS WHERE NoFRS = " & m_iNoFRS, g_connData, adOpenDynamic, adLockOptimistic)
  
        'Tant que ce n'est pas la fin des enregistrements
40      Do While Not rstContactFRS.EOF
45        Call rstContact.Open("SELECT * FROM GRB_Contact WHERE IDContact = " & rstContactFRS.Fields("NoContact"), g_connData, adOpenDynamic, adLockOptimistic)
                    
50        If Not rstContact.EOF Then
55          Call cmbContactFRS.AddItem(rstContact.Fields("NomContact"))
60        End If
    
65        Call rstContact.Close
    
70        Call rstContactFRS.MoveNext
75      Loop
  
80      Set rstContact = Nothing
  
85      Call rstContactFRS.Close
90      Set rstContactFRS = Nothing
  
95      Exit Sub

AfficherErreur:

100     woups "frmChoixContactFRS", "RemplirComboContactFRS", Err, Erl
End Sub

Public Sub Afficher(ByVal iNoFRS As Integer)

5       On Error GoTo AfficherErreur

10      Dim rstFRS As ADODB.Recordset
 
15      Set rstFRS = New ADODB.Recordset
 
20      Call rstFRS.Open("SELECT * FROM GRB_Fournisseur WHERE IDFRS = " & iNoFRS, g_connData, adOpenDynamic, adLockOptimistic)

25      lblQuestion.Caption = "Qui est le contact pour " & Replace(rstFRS.Fields("NomFournisseur"), "&", "&&") & "?"
 
30      Call rstFRS.Close
35      Set rstFRS = Nothing
 
40      m_iNoFRS = iNoFRS

45      Call RemplirComboContactFRS

50      Call Me.Show(vbModal)

55      Exit Sub

AfficherErreur:

60      woups "frmChoixContactFRS", "Form_Load", Err, Erl
End Sub

