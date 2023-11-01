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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
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

 On Error GoTo Oups

 frmChoixDemande.m_bAnnulerContact = True

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixContactFRS", "cmdAnnuler_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdOK_Click()

 On Error GoTo Oups

 frmChoixDemande.m_bAnnulerContact = False

 frmChoixDemande.m_sContact = cmbContactFRS.Text
 
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixContactFRS", "cmdOK_Click", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboContactFRS()

 On Error GoTo Oups

 'Remplir le combo des contacts
 Dim rstContactFRS As ADODB.Recordset
 Dim rstContact As ADODB.Recordset
 
 'Il faut vider le combo avant de le remplir
 Call cmbContactFRS.Clear
 
 Set rstContactFRS = New ADODB.Recordset
 Set rstContact = New ADODB.Recordset
 
 Call rstContactFRS.Open("SELECT * FROM GrbContactFRS WHERE NoFRS = " & m_iNoFRS, g_connData, adOpenDynamic, adLockOptimistic)
 
 'Tant que ce n'est pas la fin des enregistrements
 Do While Not rstContactFRS.EOF
 Call rstContact.Open("SELECT * FROM GrbContact WHERE IDContact = " & rstContactFRS.Fields("NoContact"), g_connData, adOpenDynamic, adLockOptimistic)
 
 If Not rstContact.EOF Then
 Call cmbContactFRS.AddItem(rstContact.Fields("NomContact"))
  End If
 
  Call rstContact.Close
 
  Call rstContactFRS.MoveNext
  Loop
 
  Set rstContact = Nothing
 
  Call rstContactFRS.Close
  Set rstContactFRS = Nothing
 
  Exit Sub

Oups:

10 wOups "frmChoixContactFRS", "RemplirComboContactFRS", Err, Err.number, Err.Description
End Sub

Public Sub Afficher(ByVal iNoFRS As Integer)

 On Error GoTo Oups

 Dim rstFRS As ADODB.Recordset
 
 Set rstFRS = New ADODB.Recordset
 
 Call rstFRS.Open("SELECT * FROM GrbFournisseur WHERE IDFRS = " & iNoFRS, g_connData, adOpenDynamic, adLockOptimistic)

 lblQuestion.Caption = "Qui est le contact pour " & Replace(rstFRS.Fields("NomFournisseur"), "&", "&&") & "?"
 
 Call rstFRS.Close
 Set rstFRS = Nothing
 
 m_iNoFRS = iNoFRS

 Call RemplirComboContactFRS

 Call Me.Show(vbModal)

 Exit Sub

Oups:

  wOups "frmChoixContactFRS", "Form_Load", Err, Err.number, Err.Description
End Sub

