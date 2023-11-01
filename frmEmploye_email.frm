VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEmploye_email 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Courriels"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmEmploye_email.frx":0000
   ScaleHeight     =   5340
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSupprimerExterne 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Supprimer"
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton cmdAjoutExterne 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ajouter"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton cmdSupprimerInterne 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Supprimer"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton cmdAjoutInterne 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ajouter"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton cmdFermer 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fermer"
      Height          =   495
      Left            =   4320
      TabIndex        =   9
      Top             =   4680
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwExterne 
      Height          =   2175
      Left            =   3120
      TabIndex        =   4
      Top             =   1800
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Email"
         Object.Width           =   4410
      EndProperty
   End
   Begin MSComctlLib.ListView lvwInterne 
      Height          =   2175
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Email"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Courriels Externes"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Courriels Internes"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label lblEmploye 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Employe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   960
      Width           =   2535
   End
End
Attribute VB_Name = "frmEmploye_email"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAjoutExterne_Click()

5       On Error GoTo AfficherErreur

10      Dim sEmail          As String
15      Dim rstEmailExterne As ADODB.Recordset

        'msg pour entree le nouveau email externe
20      sEmail = InputBox("Veuillez entrer le courriel (ex:andre@hotmail.com)!", "Courriel")
        
        'si email d'entré
25      If sEmail <> vbNullString Then
30        Set rstEmailExterne = New ADODB.Recordset

35        Call rstEmailExterne.Open("SELECT * FROM GRB_employe_email_externe WHERE email_name = '" & sEmail & "' AND UserID = '" & frmemploye.txtuserid.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
      
40        If rstEmailExterne.EOF Then
            'ajoute le email dans bd
45          Call rstEmailExterne.AddNew
        
50          rstEmailExterne.Fields("userid").Value = frmemploye.txtuserid.Text
55          rstEmailExterne.Fields("email_name").Value = sEmail
        
60          Call rstEmailExterne.Update
        
            'ferme table
65          Call rstEmailExterne.Close
70          Set rstEmailExterne = Nothing
      
            'rempli les lister
75          Call RemplirListViewExterne
80        Else
85          Call MsgBox("Le courriel est déjà existant pour cet employé!", vbOKOnly, "Erreur")
90        End If
95      End If

100     Exit Sub

AfficherErreur:

105     Call AfficherErreur(Me, "cmdAjoutExterne_Click", Err, Erl)
End Sub

Private Sub cmdAjoutInterne_Click()

5       On Error GoTo AfficherErreur

10      Dim sEmail          As String
15      Dim rstEmailInterne As ADODB.Recordset
  
        'MSG pour entree le nouveau email externe
20      sEmail = InputBox("Veuillez entrer le courriel (ex:andre.roy@grb-inc.com)!", "Courriel")
 
        'Si email d'entré
25      If sEmail <> vbNullString Then
30        Set rstEmailInterne = New ADODB.Recordset

35        Call rstEmailInterne.Open("SELECT * FROM GRB_employe_email_interne WHERE email_name = '" & sEmail & "' AND UserID = '" & frmemploye.txtuserid.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
     
40        If rstEmailInterne.EOF Then
            'Ajoute le email dans bd
45          Call rstEmailInterne.AddNew
        
50          rstEmailInterne.Fields("userid").Value = frmemploye.txtuserid.Text
55          rstEmailInterne.Fields("email_name").Value = sEmail
        
60          Call rstEmailInterne.Update
        
            'Ferme table
65          Call rstEmailInterne.Close
70          Set rstEmailInterne = Nothing
                 
            'rempli les lister
75          Call RemplirListViewInterne
80        Else
85          Call MsgBox("Le courriel est déjà existant pour cet employé!")
90        End If
95      End If

100     Exit Sub

AfficherErreur:

105     Call AfficherErreur(Me, "cmdAjoutInterne_Click", Err, Erl)
End Sub

Private Sub cmdFermer_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      Call AfficherErreur(Me, "cmdFermer_Click", Err, Erl)
End Sub

Private Sub RemplirListViewExterne()

5       On Error GoTo AfficherErreur

10      Dim rstEmailExterne As ADODB.Recordset
15      Dim itmEmail        As ListItem

        'vide le lister
20      Call lvwExterne.ListItems.Clear
  
25      Set rstEmailExterne = New ADODB.Recordset
  
30      Call rstEmailExterne.Open("SELECT * FROM GRB_employe_email_externe WHERE UserID = '" & frmemploye.txtuserid.Text & "' ORDER BY email_name", g_connData, adOpenDynamic, adLockOptimistic)
    
        'tant il y a de employé cedulé , ajoute dans lister
35      Do While Not rstEmailExterne.EOF
40        Set itmEmail = lvwExterne.ListItems.Add
      
45        itmEmail.Text = rstEmailExterne.Fields("email_name")
      
50        Call rstEmailExterne.MoveNext
55      Loop
       
60      Call rstEmailExterne.Close
65      Set rstEmailExterne = Nothing

70      Exit Sub

AfficherErreur:

75      Call AfficherErreur(Me, "RemplirListViewExterne", Err, Erl)
End Sub

Private Sub RemplirListViewInterne()

5       On Error GoTo AfficherErreur

10      Dim rstEmailInterne As ADODB.Recordset
15      Dim itmEmail        As ListItem
  
        'vide le lister
20      Call lvwInterne.ListItems.Clear
  
25      Set rstEmailInterne = New ADODB.Recordset
  
30      Call rstEmailInterne.Open("SELECT * FROM GRB_employe_email_interne WHERE UserID = '" & frmemploye.txtuserid.Text & "' ORDER BY email_name", g_connData, adOpenDynamic, adLockOptimistic)
    
35      Do While Not rstEmailInterne.EOF
40        Set itmEmail = lvwInterne.ListItems.Add
      
45        itmEmail.Text = rstEmailInterne.Fields("email_name")
      
50        Call rstEmailInterne.MoveNext
55      Loop
    
60      Call rstEmailInterne.Close
65      Set rstEmailInterne = Nothing

70      Exit Sub

AfficherErreur:

75      Call AfficherErreur(Me, "RemplirListViewInterne", Err, Erl)
End Sub

Private Sub cmdSupprimerExterne_Click()

5       On Error GoTo AfficherErreur
        
        ''''''''''''''''''''''''''''''''
        'Supprime l'adresse email selectionné
        ''''''''''''''''''''''''''''''''
   
        'si il y a des enreg
10      If lvwExterne.ListItems.Count > 0 Then
          'supprime l'enreg
15        If MsgBox("Voulez-vous supprimer cet enregistrement?", vbYesNo) = vbYes Then
20          Call g_connData.Execute("DELETE * FROM GRB_employe_email_externe WHERE email_name = '" & lvwExterne.SelectedItem.Text & "' AND userid = '" & frmemploye.txtuserid.Text & "'")
          
            'mise a jour des lister
25          Call RemplirListViewExterne
30        End If
35      End If

40      Exit Sub

AfficherErreur:

45      Call AfficherErreur(Me, "cmdSupprimerExterne_Click", Err, Erl)
End Sub

Private Sub cmdSupprimerInterne_Click()

5       On Error GoTo AfficherErreur
              
        ''''''''''''''''''''''''''''''''
        'Supprime l'adresse email selectionné
        ''''''''''''''''''''''''''''''''
  
        'si ya enreg
10      If lvwInterne.ListItems.Count > 0 Then
          'supprime l'enreg
15        If MsgBox("Voulez-vous supprimer cet enregistrement?", vbYesNo) = vbYes Then
20          Call g_connData.Execute("DELETE * FROM GRB_employe_email_interne WHERE email_name = '" & lvwInterne.SelectedItem.Text & "' AND userid = '" & frmemploye.txtuserid.Text & "'")
                
            'mise a jour des lister
25          Call RemplirListViewInterne
30        End If
35      End If

40      Exit Sub

AfficherErreur:

45      Call AfficherErreur(Me, "cmdSupprimerInterne_Click", Err, Erl)
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur
        
        '''''''''''''''''''''''''''''''
        'initialise affichage a l'écran
        '''''''''''''''''''''''''''''''
10      lblEmploye.Caption = frmemploye.cmbEmploye.Text
              
        'rempli les lister
15      Call RemplirListViewExterne
20      Call RemplirListViewInterne
  
25      Call ActiverBoutonsGroupe

30      Exit Sub

AfficherErreur:

35      Call AfficherErreur(Me, "Form_Load", Err, Erl)
End Sub

Private Sub ActiverBoutonsGroupe()

5       On Error GoTo AfficherErreur
        
        'Activation des boutons selon le groupe
10      cmdAjoutInterne.Enabled = g_bModificationCourriels
15      cmdSupprimerInterne.Enabled = g_bModificationCourriels
20      cmdAjoutExterne.Enabled = g_bModificationCourriels
25      cmdSupprimerExterne.Enabled = g_bModificationCourriels

30      Exit Sub

AfficherErreur:

35      Call AfficherErreur(Me, "ActiverBoutonsGroupe", Err, Erl)
End Sub

Private Sub lvwExterne_KeyDown(KeyCode As Integer, Shift As Integer)

5       On Error GoTo AfficherErreur

10      If lvwExterne.ListItems.Count > 0 Then
15        If KeyCode = vbKeyDelete Then
            'supprime l'enreg
20          If MsgBox("Voulez-vous supprimer cet enregistrement?", vbYesNo) = vbYes Then
25            Call g_connData.Execute("DELETE * FROM GRB_employe_email_externe WHERE email_name = '" & lvwExterne.SelectedItem.Text & "' AND userID = '" & frmemploye.txtuserid.Text & "'")
     
                    'mise a jour des lister
30            Call RemplirListViewExterne
35          End If
40        End If
45      End If

50      Exit Sub

AfficherErreur:

55      Call AfficherErreur(Me, "lvwExterne_KeyDown", Err, Erl)
End Sub

Private Sub lvwInterne_KeyDown(KeyCode As Integer, Shift As Integer)

5       On Error GoTo AfficherErreur

10      If lvwInterne.ListItems.Count > 0 Then
15        If KeyCode = vbKeyDelete Then
            'supprime l'enreg
20          If MsgBox("Voulez-vous supprimer cet enregistrements?", vbYesNo) = vbYes Then
25            Call g_connData.Execute("DELETE * FROM GRB_employe_email_interne WHERE email_name = '" & lvwInterne.SelectedItem.Text & "' AND userID = '" & frmemploye.txtuserid.Text & "'")
          
              'mise a jour des listers
30            Call RemplirListViewInterne
35          End If
40        End If
45      End If

50      Exit Sub

AfficherErreur:

55      Call AfficherErreur(Me, "lvwInterne_KeyDown", Err, Erl)
End Sub
