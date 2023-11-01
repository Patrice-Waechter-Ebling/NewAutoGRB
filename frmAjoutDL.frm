VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAjoutDL 
   BackColor       =   &H00000000&
   Caption         =   "Création des listes de distribution"
   ClientHeight    =   5880
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   9465
   Icon            =   "frmAjoutDL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   9465
   Begin VB.CommandButton cmdExceptions 
      Caption         =   "Liste des exceptions ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   16
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Frame fraChoixDossier 
      BackColor       =   &H00000000&
      Caption         =   "Choix du dossier dans Outlook"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   4920
      TabIndex        =   10
      Top             =   120
      Width           =   4335
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   360
         ScaleHeight     =   1815
         ScaleWidth      =   3855
         TabIndex        =   11
         Top             =   360
         Width           =   3855
         Begin VB.OptionButton optChoixDossier 
            BackColor       =   &H00000000&
            Caption         =   "Fournisseurs GRB"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   14
            Top             =   1440
            Width           =   2175
         End
         Begin VB.OptionButton optChoixDossier 
            BackColor       =   &H00000000&
            Caption         =   "Clients GRB"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   1695
         End
         Begin VB.OptionButton optChoixDossier 
            BackColor       =   &H00000000&
            Caption         =   "Contacts GRB"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1695
         End
      End
   End
   Begin VB.Frame fraChoixListe 
      BackColor       =   &H00000000&
      Caption         =   "Choix de la liste"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   4335
      Begin VB.OptionButton optChoix 
         BackColor       =   &H00000000&
         Caption         =   "Tous les membres du Meat Processing"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   5
         Left            =   240
         TabIndex        =   15
         Top             =   2280
         Width           =   3855
      End
      Begin VB.OptionButton optChoix 
         BackColor       =   &H00000000&
         Caption         =   "Tous les fournisseurs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   9
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton optChoix 
         BackColor       =   &H00000000&
         Caption         =   "Tous les contacts"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton optChoix 
         BackColor       =   &H00000000&
         Caption         =   "Tous les clients"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   1695
      End
      Begin VB.OptionButton optChoix 
         BackColor       =   &H00000000&
         Caption         =   "Tous les clients facturés"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   6
         Top             =   1200
         Width           =   2175
      End
      Begin VB.OptionButton optChoix 
         BackColor       =   &H00000000&
         Caption         =   "Tous les membres du groupement des chefs d'entreprise"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Width           =   3855
      End
   End
   Begin MSComctlLib.ProgressBar pgb 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   4560
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtPrefixe 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4080
      TabIndex        =   2
      Top             =   3840
      Width           =   3855
   End
   Begin VB.CommandButton cmdCreer 
      Caption         =   "Créer la liste"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   0
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Préfixe de la liste de distribution :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   3900
      Width           =   3735
   End
End
Attribute VB_Name = "frmAjoutDL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_rstData As ADODB.Recordset

Private Const I_OPT_CONTACTS         As Integer = 0
Private Const I_OPT_CLIENTS          As Integer = 1
Private Const I_OPT_CLIENTS_FACTURES As Integer = 2
Private Const I_OPT_FRS              As Integer = 3
Private Const I_OPT_GROUPEMENT       As Integer = 4
Private Const I_OPT_MEAT_PROCESSING  As Integer = 5

Private Const I_OPT_DOSSIER_CONTACTS As Integer = 0
Private Const I_OPT_DOSSIER_CLIENTS  As Integer = 1
Private Const I_OPT_DOSSIER_FRS      As Integer = 2

Private m_otlApp           As Outlook.Application
Private m_bDejaOuvert      As Boolean
Private m_arr_sException() As String

Private Sub cmdCreer_Click()

5       On Error GoTo AfficherErreur

10      Dim folDestination As Outlook.MAPIFolder
15      Dim folSource      As Outlook.MAPIFolder
20      Dim itmContact     As Outlook.ContactItem
25      Dim otlDistList    As Outlook.DistListItem
30      Dim otlRecipient   As Outlook.Recipient
35      Dim myItems        As Outlook.Items
40      Dim objItem        As Object
45      Dim iIndexListe    As Integer
50      Dim rstData        As ADODB.Recordset

55      If optChoix(0).Value = True Or optChoix(1).Value = True Or optChoix(2).Value = True Or optChoix(3).Value = True Or optChoix(4).Value = True Or optChoix(5).Value = True Then
60        If optChoixDossier(0).Value = True Or optChoixDossier(1).Value = True Or optChoixDossier(2).Value = True Then
65          If Trim$(txtPrefixe.Text) <> "" Then
70            fraChoixListe.Enabled = False
75            fraChoixDossier.Enabled = False
80            cmdCreer.Enabled = False
85            cmdExceptions.Enabled = False
90            txtPrefixe.Enabled = False

95            Call RemplirArrayExceptions

100           If optChoix(I_OPT_CLIENTS_FACTURES).Value = True Then
105             Call MsgBox("Veuillez noter que cette liste peut prendre plusieurs minutes avant de débuter!", vbInformation)
110           End If

115           If optChoix(I_OPT_GROUPEMENT).Value = True Or optChoix(I_OPT_MEAT_PROCESSING).Value = True Then
120             Call AjouterGroupementMeatProcessing
                                
125             fraChoixListe.Enabled = True
130             fraChoixDossier.Enabled = True
135             cmdCreer.Enabled = True
140             cmdExceptions.Enabled = True
145             txtPrefixe.Enabled = True
                
150             Exit Sub
155           End If
                            
160           Set rstData = New ADODB.Recordset
              
165           rstData.CursorLocation = adUseClient
              
170           If optChoix(I_OPT_CONTACTS).Value = True Then
175             Call rstData.Open("SELECT * FROM GRB_Contact WHERE [E-mail] Is Not null And [E-mail] <> '' ORDER BY NomContact", g_connData, adOpenForwardOnly, adLockReadOnly)

180             Set folSource = GetFolder(m_otlApp, "Contacts GRB")
185           Else
190             If optChoix(I_OPT_FRS).Value = True Then
195               Call rstData.Open("SELECT * FROM GRB_Fournisseur WHERE [E-mail] Is Not null And [E-mail] <> '' ORDER BY NomFournisseur", g_connData, adOpenForwardOnly, adLockReadOnly)

200               Set folSource = GetFolder(m_otlApp, "Fournisseurs GRB")
205             Else
210               If optChoix(I_OPT_CLIENTS).Value = True Then
215                 Call rstData.Open("SELECT * FROM GRB_Client WHERE Email Is Not null And Email <> '' ORDER BY NomClient", g_connData, adOpenForwardOnly, adLockReadOnly)
220               Else
225                 Call rstData.Open("SELECT DISTINCT(GRB_Punch.NoClient), GRB_Client.NomClient FROM GRB_Punch INNER JOIN GRB_Client ON CInt(GRB_Punch.NoClient) = CInt(GRB_Client.IDClient) WHERE GRB_Punch.NoClient <> '' AND GRB_Punch.NoClient Is Not Null AND GRB_Punch.Facturé = True AND Email Is Not null And Email <> '' ORDER BY GRB_Client.NomClient", g_connData, adOpenForwardOnly, adLockReadOnly)
230               End If

235               Set folSource = GetFolder(m_otlApp, "Clients GRB")
240             End If
245           End If
      
250           If optChoixDossier(I_OPT_CONTACTS).Value = True Then
255             Set folDestination = GetFolder(m_otlApp, "Contacts GRB")
260           Else
265             If optChoixDossier(I_OPT_CLIENTS).Value = True Then
270               Set folDestination = GetFolder(m_otlApp, "Clients GRB")
275             Else
280               Set folDestination = GetFolder(m_otlApp, "Fournisseurs GRB")
285             End If
290           End If
           
295           pgb.Min = 0
300           pgb.Max = rstData.RecordCount
305           pgb.Value = 0
             
310           iIndexListe = 1
      
315           Do While Not rstData.EOF
320             Set otlDistList = Nothing
              
325             Set myItems = folDestination.Items.Restrict("[MessageClass] = 'IPM.DistList'")
                
330             For Each objItem In myItems
335               If objItem = txtPrefixe.Text & Right$("00" & iIndexListe, 3) Then
340                 Set otlDistList = objItem
               
345                 Exit For
350               End If
355             Next
             
360             If otlDistList Is Nothing Then
365               Set otlDistList = folDestination.Items.Add(olDistributionListItem)
                  
370               otlDistList.DLName = txtPrefixe.Text & Right$("00" & iIndexListe, 3)
     
375               Call otlDistList.Save
380             Else
385               If otlDistList.MemberCount = 10 Then
390                 iIndexListe = iIndexListe + 1
                  
395                 Set otlDistList = folDestination.Items.Add(olDistributionListItem)
                    
400                 otlDistList.DLName = txtPrefixe.Text & Right$("00" & iIndexListe, 3)
                    
405                 Call otlDistList.Save
410               End If
415             End If
                
420             If optChoix(I_OPT_CONTACTS).Value = True Then
425               Set itmContact = folSource.Items.Find("[User1] = " & rstData.Fields("IDContact"))
430             Else
435               If optChoix(I_OPT_FRS).Value = True Then
440                 Set itmContact = folSource.Items.Find("[User1]= " & rstData.Fields("IDFRS"))
445               Else
450                 If optChoix(I_OPT_CLIENTS).Value = True Then
455                   Set itmContact = folSource.Items.Find("[User1] = " & rstData.Fields("IDClient"))
460                 Else
465                   Set itmContact = folSource.Items.Find("[User1] = " & rstData.Fields("NoClient"))
470                 End If
475               End If
480             End If
                  
485             If Not itmContact Is Nothing Then
490               If IsException(itmContact.Email1Address) = False Then
495                 Set otlRecipient = m_otlApp.Session.CreateRecipient(itmContact.Email1DisplayName)
          
500                 If otlRecipient.Resolve = True Then
505                   Call otlDistList.AddMember(otlRecipient)
             
510                   Call otlDistList.Save
515                 End If
520               End If
525             Else
530               If optChoix(I_OPT_CONTACTS).Value = True Then
535                 Call MsgBox("Le contact " & rstData.Fields("NomContact") & " n'a pas été trouvé dans outlook.")
540               Else
545                 If optChoix(I_OPT_FRS).Value = True Then
550                   Call MsgBox("Le fournisseur " & rstData.Fields("NomFournisseur") & " n'a pas été trouvé dans outlook.")
555                 Else
560                   Call MsgBox("Le client " & rstData.Fields("NomClient") & " n'a pas été trouvé dans outlook.")
565                 End If
570               End If
575             End If
             
580             pgb.Value = pgb.Value + 1
                
585             DoEvents
                
590             Call rstData.MoveNext
595           Loop
              
600           Call rstData.Close
605           Set rstData = Nothing
            
610           Call MsgBox("Terminé!")

615           fraChoixDossier.Enabled = True
620           fraChoixListe.Enabled = True
625           cmdCreer.Enabled = True
630           cmdExceptions.Enabled = True
635           txtPrefixe.Enabled = True
640         Else
645           Call MsgBox("Le préfixe de la liste à créer ne doit pas être vide!", vbOKOnly, "Erreur")
650         End If
655       Else
660         Call MsgBox("Vous devez choisir un dossier de destination dans Outlook!", vbOKOnly, "Erreur")
665       End If
670     Else
675       Call MsgBox("Vous devez choisir une liste à faire!", vbOKOnly, "Erreur")
680     End If

685     Exit Sub

AfficherErreur:

690     woups "frmAjoutDL", "cmdCreer_Click", Err, Erl
End Sub

Private Sub AjouterGroupementMeatProcessing()
  
5       On Error GoTo AfficherErreur

10      Dim folContact     As Outlook.MAPIFolder
15      Dim folClient      As Outlook.MAPIFolder
20      Dim folDestination As Outlook.MAPIFolder
25      Dim itmContact     As Outlook.ContactItem
30      Dim otlDistList    As Outlook.DistListItem
35      Dim otlRecipient   As Outlook.Recipient
40      Dim myItems        As Outlook.Items
45      Dim objItem        As Object
50      Dim iIndexListe    As Integer
55      Dim rstData        As ADODB.Recordset
  
60      Set rstData = New ADODB.Recordset
  
65      rstData.CursorLocation = adUseClient
  
        'E-mail = Courriel du contact
        'Email = Courriel du client
70      If optChoix(I_OPT_GROUPEMENT).Value = True Then
75        Call rstData.Open("SELECT IDContact, [E-mail], IDClient, Email FROM (GRB_ContactClient INNER JOIN GRB_Client ON GRB_ContactClient.noclient = GRB_Client.IDClient) INNER JOIN GRB_contact ON GRB_ContactClient.nocontact = GRB_contact.IDContact WHERE (INSTR(Titre,'(Groupement des chefs') > 0) AND ([E-mail] <> '' OR Email <> '') ORDER BY GRB_Contact.NomContact", g_connData, adOpenForwardOnly, adLockReadOnly)
80      Else
85        Call rstData.Open("SELECT IDContact, [E-mail], IDClient, Email FROM (GRB_ContactClient INNER JOIN GRB_Client ON GRB_ContactClient.noclient = GRB_Client.IDClient) INNER JOIN GRB_contact ON GRB_ContactClient.nocontact = GRB_contact.IDContact WHERE (INSTR(Titre,'(Meat Processing)') > 0) AND ([E-mail] <> '' OR Email <> '') ORDER BY GRB_Contact.NomContact", g_connData, adOpenForwardOnly, adLockReadOnly)
90      End If
    
95      Set folContact = GetFolder(m_otlApp, "Contacts GRB")
100     Set folClient = GetFolder(m_otlApp, "Clients GRB")

105     If optChoixDossier(I_OPT_DOSSIER_CONTACTS).Value = True Then
110       Set folDestination = folContact
115     Else
120       If optChoixDossier(I_OPT_DOSSIER_CLIENTS).Value = True Then
125         Set folDestination = folClient
130       Else
135         Set folDestination = GetFolder(m_otlApp, "Fournisseurs GRB")
140       End If
145     End If
  
150     pgb.Min = 0
155     pgb.Max = rstData.RecordCount
160     pgb.Value = 0
  
165     iIndexListe = 1

170     Do While Not rstData.EOF
175       Set otlDistList = Nothing
  
180       Set myItems = folDestination.Items.Restrict("[MessageClass] = 'IPM.DistList'")
    
185       For Each objItem In myItems
190         If objItem = txtPrefixe.Text & Right$("00" & iIndexListe, 3) Then
195           Set otlDistList = objItem
    
200           Exit For
205         End If
210       Next
  
215       If otlDistList Is Nothing Then
220         Set otlDistList = folDestination.Items.Add(olDistributionListItem)
      
225         otlDistList.DLName = txtPrefixe.Text & Right$("00" & iIndexListe, 3)
      
230         Call otlDistList.Save
235       Else
240         If otlDistList.MemberCount = 10 Then
245           iIndexListe = iIndexListe + 1
        
250           Set otlDistList = folDestination.Items.Add(olDistributionListItem)
        
255           otlDistList.DLName = txtPrefixe.Text & Right$("00" & iIndexListe, 3)
        
260           Call otlDistList.Save
265         End If
270       End If
    
275       If rstData.Fields("E-mail") <> "" Then 'Si le contact a un e-mail then
280         Set itmContact = folContact.Items.Find("[User1] = " & rstData.Fields("IDContact"))
285       Else
            'Sinon, on prend le courriel du client
290         Set itmContact = folClient.Items.Find("[User1] = " & rstData.Fields("IDClient"))
295       End If
        
300       If Not itmContact Is Nothing Then
305         If IsException(itmContact.Email1Address) = False Then
310           Set otlRecipient = m_otlApp.Session.CreateRecipient(itmContact.Email1DisplayName)
  
315           If otlRecipient.Resolve = True Then
320             Call otlDistList.AddMember(otlRecipient)
  
325             Call otlDistList.Save
330           End If
335         End If
340       End If
    
345       pgb.Value = pgb.Value + 1
    
350       DoEvents
    
355       Call rstData.MoveNext
360     Loop
  
365     Call rstData.Close
370     Set rstData = Nothing
  
375     Call MsgBox("Terminé!")

380     Exit Sub

AfficherErreur:

385     woups "frmAjoutDL", "AjouterGroupementMeatProcessing", Err, Erl
End Sub

Private Sub cmdExceptions_Click()

5       On Error GoTo AfficherErreur

10      Call OuvrirForm(frmExceptionsDL, False)

15      Exit Sub

AfficherErreur:

20      woups "frmAjoutDL", "cmdExceptions_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      Set m_otlApp = OuvrirOutlook(m_bDejaOuvert)

15      Exit Sub

AfficherErreur:

20      woups "frmAjoutDL", "Form_Load", Err, Erl
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
5       On Error GoTo AfficherErreur
  
10      If m_bDejaOuvert = False Then
15        Call m_otlApp.Quit
20      End If
  
25      Set m_otlApp = Nothing

30      Exit Sub

AfficherErreur:

35      woups "frmAjoutDL", "Form_Load", Err, Erl
End Sub

Private Function IsException(ByVal sAdresse As String) As Boolean

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer

15      For iCompteur = 0 To UBound(m_arr_sException)
20        If m_arr_sException(iCompteur) = sAdresse Then
25          IsException = True

30          Exit For
35        End If
40      Next

45      Exit Function

AfficherErreur:

50      woups "frmAjoutDL", "IsException", Err, Erl
End Function

Private Sub RemplirArrayExceptions()

5       On Error GoTo AfficherErreur

10      Dim rstExceptions As ADODB.Recordset
15      Dim iNombre       As Integer

20      ReDim m_arr_sException(0)

25      Set rstExceptions = New ADODB.Recordset

30      Call rstExceptions.Open("SELECT * FROM GRB_ExceptionsDL", g_connData, adOpenForwardOnly, adLockReadOnly)

35      Do While Not rstExceptions.EOF
40        ReDim Preserve m_arr_sException(0 To iNombre)

45        m_arr_sException(iNombre) = rstExceptions.Fields("Exception")

50        iNombre = iNombre + 1

55        Call rstExceptions.MoveNext
60      Loop

65      Call rstExceptions.Close
70      Set rstExceptions = Nothing

75      Exit Sub

AfficherErreur:

80      woups "frmAjoutDL", "RemplirArrayExceptions", Err, Erl
End Sub
