VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEmail 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Courriels"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9735
   Icon            =   "frmEmail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmEmail.frx":0442
   ScaleHeight     =   7695
   ScaleWidth      =   9735
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrVerif 
      Interval        =   1000
      Left            =   5040
      Top             =   240
   End
   Begin VB.Frame fraLecture 
      BackColor       =   &H00000000&
      Height          =   6975
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   9735
      Begin VB.CommandButton cmdNouveau 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nouveau message"
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   6480
         Width           =   1695
      End
      Begin MSComctlLib.ListView lstemail 
         Height          =   5655
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   9975
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "no email"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "De"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Objet"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Reçu"
            Object.Width           =   3351
         EndProperty
      End
      Begin VB.CommandButton cmdFermer 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fermer"
         Height          =   375
         Left            =   6360
         TabIndex        =   7
         Top             =   6480
         Width           =   1575
      End
      Begin VB.CommandButton cmdSupprimer 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Supprimer message"
         Height          =   375
         Left            =   4440
         TabIndex        =   6
         Top             =   6480
         Width           =   1695
      End
      Begin VB.Label lblNbreMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   6480
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Lecture des messages reçus"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   120
         Width           =   3615
      End
   End
   Begin VB.OLE OLE1 
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_iSeconde As Integer
Private m_bNewMail As Boolean

Private Sub cmdNouveau_Click()

5       On Error GoTo AfficherErreur

10      Dim otlMessage  As Outlook.Application
15      Dim myItem      As Outlook.MailItem
20      Dim myNameSpace As Outlook.NameSpace
25      Dim myFolder    As Outlook.MAPIFolder
  
30      MousePointer = vbHourglass
  
35      Set otlMessage = CreateObject("Outlook.Application")
  
        'Ouverture d'Outlook sur la boite d'envoi
40      Set myNameSpace = otlMessage.GetNamespace("MAPI")
          
45      Set myFolder = myNameSpace.GetDefaultFolder(olFolderOutbox)
  
50      Call myFolder.Display
  
        'Ouverture d'un nouveau message
55      Set myItem = otlMessage.CreateItem(olMailItem)
  
60      Call myItem.Display
  
65      MousePointer = vbDefault

70      Exit Sub

AfficherErreur:

75      Call AfficherErreur(Me, "cmdNouveau_Click", Err, Erl)
End Sub

Private Sub cmdFermer_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      Call AfficherErreur(Me, "cmdFermer_Click", Err, Erl)
End Sub

Private Sub cmdSupprimer_Click()

5       On Error GoTo AfficherErreur

10      Dim fso       As FileSystemObject
15      Dim iCompteur As Integer

20      Set fso = CreateObject("Scripting.FileSystemObject")

25      If lstemail.ListItems.Count > 0 Then
          'si reponse est oui pour supprimer
30        If MsgBox("Voulez-vous supprimer ce(s) message(s)?", vbYesNo, "Supprimer message") = vbYes Then
35          MousePointer = vbHourglass

40          For iCompteur = 1 To lstemail.ListItems.Count
45            If lstemail.ListItems(iCompteur).Selected = True Then

                'Supprime le repertoire
50              If fso.FolderExists(lstemail.ListItems(iCompteur).Tag) = True Then
55                Call fso.GetFolder(lstemail.ListItems(iCompteur).Tag).Delete
60              End If

                'Supprime dans la table
65              Call g_connData.Execute("DELETE * FROM GRB_email_recu WHERE noemail = " & CInt(lstemail.ListItems(iCompteur).Text))
70            End If
75          Next
80        End If

          'Rempli le lister email
85        Call RemplirListView

90        MousePointer = vbDefault
95      End If

100     Exit Sub

AfficherErreur:

105     Call AfficherErreur(Me, "cmdSupprimer_Click", Err, Erl)
End Sub

Private Sub Form_Resize()

5       On Error GoTo AfficherErreur

10      m_bNewMail = False

15      Exit Sub

AfficherErreur:

20      Call AfficherErreur(Me, "Form_Resize", Err, Erl)
End Sub

Private Sub lstEmail_DblClick()

5       On Error GoTo AfficherErreur

        '''''''''''''''''''''''''''''''''
        ' Exécute le fichier sélectionné
        '''''''''''''''''''''''''''''''''
10      Dim fso As FileSystemObject
    
15      Set fso = CreateObject("Scripting.FileSystemObject")
  
        'si fichier de sélectionné
        'execute le fichier sélectionné
20      If lstemail.ListItems.Count > 0 Then
25        If fso.FolderExists(lstemail.SelectedItem.Tag) = True Then
30          Call OLE1.CreateLink(lstemail.SelectedItem.Tag & "\email.msg")
      
35          If OLE1.OLEType <> vbOLENone Then
40            Call OLE1.DoVerb(0)
45          Else
50            Call MsgBox("Impossible de créer une liaison avec le courriel!", vbOKOnly, "Erreur")
55          End If
60        Else
65          Call MsgBox("Ce courriel n'existe plus ou perte de communication avec le serveur!", vbOKOnly, "Erreur")
70        End If
75      End If

80      Exit Sub

AfficherErreur:

85      Call AfficherErreur(Me, "lstEmail_DblClick", Err, Erl)
End Sub

Private Sub RemplirListView()

5       On Error GoTo AfficherErreur

        ''''''''''''''''''''''''''''''''''''''''
        'remplis le lister email
        ''''''''''''''''''''''''''''''''''''
10      Dim rstEmail    As ADODB.Recordset
15      Dim itmEmail    As ListItem
20      Dim sDestiEmail As String
25      Dim iCmpEmail   As Integer

30      Set rstEmail = New ADODB.Recordset

35      Call rstEmail.Open("SELECT * FROM GRB_email_recu WHERE UserID = '" & g_sUserID & "' ORDER BY NoEmail DESC", g_connData, adOpenDynamic, adLockOptimistic)

        'vide
40      Call lstemail.ListItems.Clear
  
45      iCmpEmail = 0
  
50      Do While Not rstEmail.EOF
55        Set itmEmail = lstemail.ListItems.Add
            
60        itmEmail.Text = rstEmail.Fields("noemail")
       
          'chemin du serveur
65        sDestiEmail = "\\Serveur\" + Mid(rstEmail.Fields("attachement"), 8, Len(rstEmail.Fields("attachement")) - 7)
    
70        itmEmail.Tag = sDestiEmail
        
75        If IsNull(rstEmail!De) Then
80          Call itmEmail.ListSubItems.Add(, , vbNullString)
85        Else
90          Call itmEmail.ListSubItems.Add(, , rstEmail!De)
95        End If
           
100       If IsNull(rstEmail!Objet) Then
105         Call itmEmail.ListSubItems.Add(, , vbNullString)
110       Else
115         Call itmEmail.ListSubItems.Add(, , rstEmail!Objet)
120       End If
            
125       If IsNull(rstEmail!Date) Then
130         Call itmEmail.ListSubItems.Add(, , vbNullString)
135       Else
140         Call itmEmail.ListSubItems.Add(, , rstEmail!Date)
145       End If
           
150       Call rstEmail.MoveNext
        
155       iCmpEmail = iCmpEmail + 1
160     Loop
  
165     Call rstEmail.Close
170     Set rstEmail = Nothing
  
        'si il y n'y a qu'un message, on écrit : "1 message"
175     If iCmpEmail = 1 Then
180       lblNbreMsg.Caption = "1 message"
185     Else
          'si il y a 0 message, on écrit : "Aucun message"
190       If iCmpEmail = 0 Then
195         lblNbreMsg.Caption = "Aucun message"
            'Sinon, on écrit par exemple : "5 messages"
200       Else
205         lblNbreMsg.Caption = iCmpEmail & " messages"
210       End If
215     End If

220     Exit Sub

AfficherErreur:

225     Call AfficherErreur(Me, "RemplirListView", Err, Erl)
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      Call RemplirListView

15      Exit Sub

AfficherErreur:

20      Call AfficherErreur(Me, "Form_Load", Err, Erl)
End Sub

Private Sub lstemail_KeyDown(KeyCode As Integer, Shift As Integer)

5       On Error GoTo AfficherErreur

10      Dim fso       As FileSystemObject
15      Dim iCompteur As Integer
  
20      Set fso = CreateObject("Scripting.FileSystemObject")
  
25      If lstemail.ListItems.Count > 0 Then
30        If KeyCode = vbKeyDelete Then
            'si reponse est oui pour supprimer
35          If MsgBox("Voulez-vous supprimer ce(s) message(s)?", vbYesNo, "Supprimer message") = vbYes Then
40            MousePointer = vbHourglass
    
45            For iCompteur = 1 To lstemail.ListItems.Count
50              If lstemail.ListItems(iCompteur).Selected = True Then
      
                  'supprime le repertoire
55                If fso.FolderExists(lstemail.ListItems(iCompteur).Tag) = True Then
60                  Call fso.GetFolder(lstemail.ListItems(iCompteur).Tag).Delete
65                End If
      
                  'supprime dans la table
70                Call g_connData.Execute("DELETE * FROM GRB_email_recu WHERE noemail = " & CInt(lstemail.ListItems(iCompteur).Text))
75              End If
80            Next
85          End If
      
            'rempli le lister email
90          Call RemplirListView
95        End If
            
100       MousePointer = vbDefault
105     End If

110     Exit Sub

AfficherErreur:

115     Call AfficherErreur(Me, "lstemail_KeyDown", Err, Erl)
End Sub

Private Sub tmrVerif_Timer()

5       On Error GoTo AfficherErreur

10      Dim iEmailAvant As Integer
15      Dim iEmailApres As Integer

20      m_iSeconde = m_iSeconde + 1

        'Flash
25      If m_bNewMail = True Then
30        Call FlashWindow(hwnd, True)
35      End If

        'Remplis la liste d'email
40      If m_iSeconde = 300 Then
45        iEmailAvant = lstemail.ListItems.Count
    
50        Call RemplirListView
    
55        iEmailApres = lstemail.ListItems.Count
        
60        If iEmailApres > iEmailAvant Then
65          m_bNewMail = True
70          Beep
75        End If
    
80        m_iSeconde = 0
85      End If

90      Exit Sub

AfficherErreur:

95      Call AfficherErreur(Me, "tmrVerif_Timer", Err, Erl)
End Sub
