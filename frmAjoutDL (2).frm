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



Private Sub cmdCreer_Click()

 On Error GoTo Oups

 Dim folDestination As Outlook.MAPIFolder
 Dim folSource As Outlook.MAPIFolder
 Dim itmContact As Outlook.ContactItem
 Dim otlDistList As Outlook.DistListItem
 Dim otlRecipient As Outlook.Recipient
 Dim myItems As Outlook.Items
 Dim objItem As Object
 Dim iIndexListe As Integer
 Dim rstData As ADODB.Recordset

 If optChoix(0).Value = True Or optChoix(1).Value = True Or optChoix(2).Value = True Or optChoix(3).Value = True Or optChoix(4).Value = True Or optChoix(5).Value = True Then
  If optChoixDossier(0).Value = True Or optChoixDossier(1).Value = True Or optChoixDossier(2).Value = True Then
  If Trim$(txtPrefixe.Text) <> "" Then
  fraChoixListe.Enabled = False
  fraChoixDossier.Enabled = False
  cmdCreer.Enabled = False
  cmdExceptions.Enabled = False
  txtPrefixe.Enabled = False

  Call RemplirArrayExceptions

 If optChoix(I_OPT_CLIENTS_FACTURES).Value = True Then
 Call MsgBox("Veuillez noter que cette liste peut prendre plusieurs minutes avant de débuter!", vbInformation)
 End If

 If optChoix(I_OPT_GROUPEMENT).Value = True Or optChoix(I_OPT_MEAT_PROCESSING).Value = True Then
 Call AjouterGroupementMeatProcessing
 
 fraChoixListe.Enabled = True
 fraChoixDossier.Enabled = True
 cmdCreer.Enabled = True
 cmdExceptions.Enabled = True
 txtPrefixe.Enabled = True
 
 Exit Sub
 End If
 
 Set rstData = New ADODB.Recordset
 
 rstData.CursorLocation = adUseClient
 
 If optChoix(I_OPT_CONTACTS).Value = True Then
 Call rstData.Open("SELECT * FROM GrbContact WHERE [E-mail] Is Not null And [E-mail] <> '' ORDER BY NomContact", g_connData, adOpenForwardOnly, adLockReadOnly)

 Set folSource = GetFolder(m_otlApp, "Contacts GRB")
 Else
 If optChoix(I_OPT_FRS).Value = True Then
1  Call rstData.Open("SELECT * FROM GrbFournisseur WHERE [E-mail] Is Not null And [E-mail] <> '' ORDER BY NomFournisseur", g_connData, adOpenForwardOnly, adLockReadOnly)

 Set folSource = GetFolder(m_otlApp, "Fournisseurs GRB")
 Else
 If optChoix(I_OPT_CLIENTS).Value = True Then
 Call rstData.Open("SELECT * FROM GrbClient WHERE Email Is Not null And Email <> '' ORDER BY NomClient", g_connData, adOpenForwardOnly, adLockReadOnly)
 Else
 Call rstData.Open("SELECT DISTINCT(GrbPunch.NoClient), GrbClient.NomClient FROM GrbPunch INNER JOIN GrbClient ON CInt(GrbPunch.NoClient) = CInt(GrbClient.IDClient) WHERE GrbPunch.NoClient <> '' AND GrbPunch.NoClient Is Not Null AND GrbPunch.Facturé = True AND Email Is Not null And Email <> '' ORDER BY GrbClient.NomClient", g_connData, adOpenForwardOnly, adLockReadOnly)
 End If

 Set folSource = GetFolder(m_otlApp, "Clients GRB")
 End If
 End If
 
 If optChoixDossier(I_OPT_CONTACTS).Value = True Then
 Set folDestination = GetFolder(m_otlApp, "Contacts GRB")
 Else
 If optChoixDossier(I_OPT_CLIENTS).Value = True Then
 Set folDestination = GetFolder(m_otlApp, "Clients GRB")
 Else
 Set folDestination = GetFolder(m_otlApp, "Fournisseurs GRB")
 End If
 End If
 
 pgb.Min = 0
 pgb.Max = rstData.RecordCount
 pgb.Value = 0
 
 iIndexListe = 1
 
 Do While Not rstData.EOF
 Set otlDistList = Nothing
 
 Set myItems = folDestination.Items.Restrict("[MessageClass] = 'IPM.DistList'")
 
 For Each objItem In myItems
 If objItem = txtPrefixe.Text & Right$("00" & iIndexListe, 3) Then
 Set otlDistList = objItem
 
 Exit For
 End If
 Next
 
 If otlDistList Is Nothing Then
 Set otlDistList = folDestination.Items.Add(olDistributionListItem)
 
 otlDistList.DLName = txtPrefixe.Text & Right$("00" & iIndexListe, 3)
 
 Call otlDistList.Save
 Else
 If otlDistList.MemberCount = 10 Then
 iIndexListe = iIndexListe + 1
 
 Set otlDistList = folDestination.Items.Add(olDistributionListItem)
 
 otlDistList.DLName = txtPrefixe.Text & Right$("00" & iIndexListe, 3)
 
4 Call otlDistList.Save
4 End If
4 End If
 
4 If optChoix(I_OPT_CONTACTS).Value = True Then
4 Set itmContact = folSource.Items.Find("[User1] = " & rstData.Fields("IDContact"))
4 Else
4 If optChoix(I_OPT_FRS).Value = True Then
4 Set itmContact = folSource.Items.Find("[User1]= " & rstData.Fields("IDFRS"))
4 Else
4 If optChoix(I_OPT_CLIENTS).Value = True Then
4 Set itmContact = folSource.Items.Find("[User1] = " & rstData.Fields("IDClient"))
4  Else
4  Set itmContact = folSource.Items.Find("[User1] = " & rstData.Fields("NoClient"))
4  End If
4  End If
4  End If
 
4  If Not itmContact Is Nothing Then
4  If IsException(itmContact.Email1Address) = False Then
4  Set otlRecipient = m_otlApp.Session.CreateRecipient(itmContact.Email1DisplayName)
 
50 If otlRecipient.Resolve = True Then
 Call otlDistList.AddMember(otlRecipient)
 
 Call otlDistList.Save
 End If
 End If
 Else
 If optChoix(I_OPT_CONTACTS).Value = True Then
 Call MsgBox("Le contact " & rstData.Fields("NomContact") & " n'a pas été trouvé dans outlook.")
 Else
 If optChoix(I_OPT_FRS).Value = True Then
 Call MsgBox("Le fournisseur " & rstData.Fields("NomFournisseur") & " n'a pas été trouvé dans outlook.")
 Else
5  Call MsgBox("Le client " & rstData.Fields("NomClient") & " n'a pas été trouvé dans outlook.")
5  End If
5  End If
5  End If
 
5  pgb.Value = pgb.Value + 1
 
5  DoEvents
 
5  Call rstData.MoveNext
5  Loop
 
60 Call rstData.Close
  Set rstData = Nothing
 
  Call MsgBox("Terminé!")

  fraChoixDossier.Enabled = True
  fraChoixListe.Enabled = True
  cmdCreer.Enabled = True
  cmdExceptions.Enabled = True
  txtPrefixe.Enabled = True
  Else
  Call MsgBox("Le préfixe de la liste à créer ne doit pas être vide!", vbOKOnly, "Erreur")
  End If
  Else
6  Call MsgBox("Vous devez choisir un dossier de destination dans Outlook!", vbOKOnly, "Erreur")
6  End If
6  Else
6  Call MsgBox("Vous devez choisir une liste à faire!", vbOKOnly, "Erreur")
6  End If

6  Exit Sub

Oups:

6  wOups "frmAjoutDL", "cmdCreer_Click", Err, Err.number, Err.Description
End Sub

Private Sub AjouterGroupementMeatProcessing()
 
 On Error GoTo Oups

 Dim folContact As Outlook.MAPIFolder
 Dim folClient As Outlook.MAPIFolder
 Dim folDestination As Outlook.MAPIFolder
 Dim itmContact As Outlook.ContactItem
 Dim otlDistList As Outlook.DistListItem
 Dim otlRecipient As Outlook.Recipient
 Dim myItems As Outlook.Items
 Dim objItem As Object
 Dim iIndexListe As Integer
 Dim rstData As ADODB.Recordset
 
  Set rstData = New ADODB.Recordset
 
  rstData.CursorLocation = adUseClient
 
 'E-mail = Courriel du contact
 'Email = Courriel du client
  If optChoix(I_OPT_GROUPEMENT).Value = True Then
  Call rstData.Open("SELECT IDContact, [E-mail], IDClient, Email FROM (GrbContactClient INNER JOIN GrbClient ON GrbContactClient.noclient = GrbClient.IDClient) INNER JOIN Grbcontact ON GrbContactClient.nocontact = Grbcontact.IDContact WHERE (INSTR(Titre,'(Groupement des chefs') > 0) AND ([E-mail] <> '' OR Email <> '') ORDER BY GrbContact.NomContact", g_connData, adOpenForwardOnly, adLockReadOnly)
  Else
  Call rstData.Open("SELECT IDContact, [E-mail], IDClient, Email FROM (GrbContactClient INNER JOIN GrbClient ON GrbContactClient.noclient = GrbClient.IDClient) INNER JOIN Grbcontact ON GrbContactClient.nocontact = Grbcontact.IDContact WHERE (INSTR(Titre,'(Meat Processing)') > 0) AND ([E-mail] <> '' OR Email <> '') ORDER BY GrbContact.NomContact", g_connData, adOpenForwardOnly, adLockReadOnly)
  End If
 
  Set folContact = GetFolder(m_otlApp, "Contacts GRB")
10 Set folClient = GetFolder(m_otlApp, "Clients GRB")

If optChoixDossier(I_OPT_DOSSIER_CONTACTS).Value = True Then
 Set folDestination = folContact
Else
 If optChoixDossier(I_OPT_DOSSIER_CLIENTS).Value = True Then
 Set folDestination = folClient
 Else
 Set folDestination = GetFolder(m_otlApp, "Fournisseurs GRB")
 End If
End If
 
pgb.Min = 0
pgb.Max = rstData.RecordCount
1  pgb.Value = 0
 
iIndexListe = 1

 Do While Not rstData.EOF
 Set otlDistList = Nothing
 
 Set myItems = folDestination.Items.Restrict("[MessageClass] = 'IPM.DistList'")
 
 For Each objItem In myItems
 If objItem = txtPrefixe.Text & Right$("00" & iIndexListe, 3) Then
1  Set otlDistList = objItem
 
 Exit For
 End If
 Next
 
 If otlDistList Is Nothing Then
 Set otlDistList = folDestination.Items.Add(olDistributionListItem)
 
 otlDistList.DLName = txtPrefixe.Text & Right$("00" & iIndexListe, 3)
 
 Call otlDistList.Save
 Else
 If otlDistList.MemberCount = 10 Then
 iIndexListe = iIndexListe + 1
 
 Set otlDistList = folDestination.Items.Add(olDistributionListItem)
 
 otlDistList.DLName = txtPrefixe.Text & Right$("00" & iIndexListe, 3)
 
 Call otlDistList.Save
 End If
End If
 
 If rstData.Fields("E-mail") <> "" Then 'Si le contact a un e-mail then
 Set itmContact = folContact.Items.Find("[User1] = " & rstData.Fields("IDContact"))
 Else
 'Sinon, on prend le courriel du client
 Set itmContact = folClient.Items.Find("[User1] = " & rstData.Fields("IDClient"))
 End If
 
If Not itmContact Is Nothing Then
If IsException(itmContact.Email1Address) = False Then
 Set otlRecipient = m_otlApp.Session.CreateRecipient(itmContact.Email1DisplayName)
 
 If otlRecipient.Resolve = True Then
 Call otlDistList.AddMember(otlRecipient)
 
 Call otlDistList.Save
 End If
 End If
 End If
 
 pgb.Value = pgb.Value + 1
 
 DoEvents
 
 Call rstData.MoveNext
3  Loop
 
Call rstData.Close
3  Set rstData = Nothing
 
Call MsgBox("Terminé!")

3  Exit Sub

Oups:

wOups "frmAjoutDL", "AjouterGroupementMeatProcessing", Err, Err.number, Err.Description
End Sub

Private Sub cmdExceptions_Click()

 On Error GoTo Oups

 Call OuvrirForm(frmExceptionsDL, False)

 Exit Sub

Oups:

 wOups "frmAjoutDL", "cmdExceptions_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 Set m_otlApp = OuvrirOutlook(m_bDejaOuvert)

 Exit Sub

Oups:

 wOups "frmAjoutDL", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
 
 On Error GoTo Oups
 
 If m_bDejaOuvert = False Then
 Call m_otlApp.Quit
 End If
 
 Set m_otlApp = Nothing

 Exit Sub

Oups:

 wOups "frmAjoutDL", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Function IsException(ByVal sAdresse As String) As Boolean

 On Error GoTo Oups

 Dim iCompteur As Integer

 For iCompteur = 0 To UBound(m_arr_sException)
 If m_arr_sException(iCompteur) = sAdresse Then
 IsException = True

 Exit For
 End If
 Next

 Exit Function

Oups:

 wOups "frmAjoutDL", "IsException", Err, Err.number, Err.Description
End Function

Private Sub RemplirArrayExceptions()

 On Error GoTo Oups

 Dim rstExceptions As ADODB.Recordset
 Dim iNombre As Integer

 ReDim m_arr_sException(0)

 Set rstExceptions = New ADODB.Recordset

 Call rstExceptions.Open("SELECT * FROM GrbExceptionsDL", g_connData, adOpenForwardOnly, adLockReadOnly)

 Do While Not rstExceptions.EOF
 ReDim Preserve m_arr_sException(0 To iNombre)

 m_arr_sException(iNombre) = rstExceptions.Fields("Exception")

 iNombre = iNombre + 1

 Call rstExceptions.MoveNext
  Loop

  Call rstExceptions.Close
  Set rstExceptions = Nothing

  Exit Sub

Oups:

  wOups "frmAjoutDL", "RemplirArrayExceptions", Err, Err.number, Err.Description
End Sub
