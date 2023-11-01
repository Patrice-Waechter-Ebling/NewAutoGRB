VERSION 5.00
Begin VB.Form frmExportToOutLook 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export to Outlook"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraEtatOutlook 
      Caption         =   "Tranfère en cours"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   5655
      Begin VB.Label lblnbre 
         Alignment       =   2  'Center
         Caption         =   "Label5"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   5175
      End
      Begin VB.Label lblEtatOutlook 
         Alignment       =   2  'Center
         Caption         =   "export data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   5175
      End
   End
   Begin VB.CheckBox ckFRS 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox ckClient 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox ckContact 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Fermer"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Exécuter"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   1455
      Left            =   120
      Top             =   720
      Width           =   5655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Exporter les Fournisseurs"
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   1800
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Exporter les Clients"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   1320
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Exporter les Contacts"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   840
      Width           =   5055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Choisir les listes à exporter."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmExportToOutLook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.ckClient.Value = False
    Me.ckContact.Value = False
    Me.ckFRS.Value = False
    
End Sub


Private Sub OKButton_Click()

    If Me.ckContact.Value = 1 Then
        'export les contacts
        If VerifierSiBesoinExport("SELECT * FROM GRB_Contact", "Contacts GRB") Then
            Call SupprimerContactExchange("Contacts GRB", "contacts")
            Call ExportContactExchange("SELECT * FROM GRB_Contact", "Contacts GRB")
        End If
    End If
    If Me.ckClient.Value = 1 Then
        'export les clients
        If VerifierSiBesoinExport("SELECT * FROM GRB_Client", "Clients GRB") Then
            Call SupprimerContactExchange("Clients GRB", "clients")
            Call ExportClientExchange("SELECT * FROM GRB_Client", "Clients GRB")
        End If
    End If
    If Me.ckFRS.Value = 1 Then
        'export les fournisseurs
        If VerifierSiBesoinExport("SELECT * FROM GRB_Fournisseur", "Fournisseurs GRB") Then
            Call SupprimerContactExchange("Fournisseurs GRB", "fournisseurs")
            Call ExportFournisseursExchange("SELECT * FROM GRB_Fournisseur", "Fournisseurs GRB")
        End If
    End If
    
    MsgBox "Exportation des données réussi."
    
    
End Sub
Private Function VerifierSiBesoinExport(ByVal strQuery As String, ByVal strFolder As String) As Boolean

5       On Error GoTo AfficherErreur

6       VerifierSiBesoinExport = False

7       Dim dummie As Integer

10      Dim otlApp      As Outlook.Application
15      Dim otlContact  As Outlook.ContactItem
20      Dim folContact  As Outlook.MAPIFolder
25      Dim bDejaOuvert As Boolean
30      Dim sNom()      As String
31      Dim rstContact  As ADODB.Recordset
32      Dim i           As Integer
33      Dim Y           As Integer


34      Screen.MousePointer = vbHourglass

35      lblEtatOutlook.Caption = "Validation ..."
36      lblnbre.Caption = "Vérifier si nous avons besoin de faire l'exportation."
37      fraEtatOutlook.Visible = True

40      Set rstContact = New ADODB.Recordset
41      Call rstContact.Open(strQuery, g_connData, adOpenDynamic, adLockOptimistic)

43      'nombre total de records dans GRB
44      i = 0
45      rstContact.MoveFirst
46      Do While Not rstContact.EOF
47          i = i + 1
48          rstContact.MoveNext
49      Loop

64      'nombre total de records dans Outlook
65      Y = 0
66      Set otlApp = OuvrirOutlook(bDejaOuvert)
70      Set folContact = GetFolder(otlApp, strFolder)
71      Y = folContact.Items.count


80      dummie = MsgBox("Nous avons " & i & " records dans GRB et " & Y & " records dans Outlook." & Chr(13) & Chr(13) & _
              "Désirez-vous toujours faire l'exportation dans Outlook?", vbYesNo, "Exportation dans Outlook")

85      If dummie = vbYes Then
86          VerifierSiBesoinExport = True
87      Else
88          VerifierSiBesoinExport = False
89      End If

          
320     Call rstContact.Close
325     Set rstContact = Nothing

330     If bDejaOuvert = False Then
335         Call otlApp.Quit
340     End If

341     Set otlApp = Nothing

342     Screen.MousePointer = vbDefault

343     fraEtatOutlook.Visible = False

345     DoEvents

350     Exit Function




AfficherErreur:

355     woups "frmExportToOutlook", "VerifierSiBesoinExport", Err, Erl
356     Call rstContact.Close
357     Set rstContact = Nothing

End Function
Private Function ExportContactExchange(ByVal strQuery As String, ByVal strFolder As String)

5       On Error GoTo AfficherErreur

10      Dim otlApp      As Outlook.Application
15      Dim otlContact  As Outlook.ContactItem
20      Dim folContact  As Outlook.MAPIFolder
25      Dim bDejaOuvert As Boolean
30      Dim sNom()      As String
35      Dim rstContact  As ADODB.Recordset
36      Dim i           As Integer

37      Screen.MousePointer = vbHourglass

40      Set rstContact = New ADODB.Recordset
45      Call rstContact.Open(strQuery, g_connData, adOpenDynamic, adLockOptimistic)

42      i = 0
43      rstContact.MoveFirst
46      Do While Not rstContact.EOF
47          i = i + 1
48          rstContact.MoveNext
49      Loop

50      lblEtatOutlook.Caption = "Ajout des contacts dans Outlook ..."
55      lblnbre.Caption = "Nombre de contact restant à transférer : " & i
60      fraEtatOutlook.Visible = True

65      Set otlApp = OuvrirOutlook(bDejaOuvert)
70      Set folContact = GetFolder(otlApp, strFolder)

90      rstContact.MoveFirst
100     Do While Not rstContact.EOF

120         Set otlContact = folContact.Items.Add(olContactItem)

130         otlContact.User1 = rstContact.Fields("IDContact")

131         If Not IsNull(rstContact.Fields("NomContact")) Then
135             sNom = Split(Trim$(rstContact.Fields("NomContact")), " ")

140             Select Case UBound(sNom)
                Case 0:
145                 otlContact.FirstName = sNom(0)
  
                Case 1:
150                 otlContact.FirstName = sNom(0)
155                 otlContact.LastName = sNom(1)

                Case 2:
160                 otlContact.FirstName = sNom(0)
165                 otlContact.MiddleName = sNom(1)
170                 otlContact.LastName = sNom(2)
175             End Select
176         End If

180         otlContact.Title = ""

181         If Not IsNull(rstContact.Fields("Compagnie")) Then
184             otlContact.CompanyName = rstContact.Fields("Compagnie")
185         End If
186         If Not IsNull(rstContact.Fields("Titre")) Or Not rstContact.Fields("Titre") = "" Then
190             otlContact.JobTitle = rstContact.Fields("Titre")
191         End If

195         If rstContact.Fields("Telephonne") <> "(___) ___-____" Then
                If Not IsNull(rstContact.Fields("NoPoste")) Then
200                 If Trim$(rstContact.Fields("NoPoste")) <> "" Then
205                     otlContact.BusinessTelephoneNumber = rstContact.Fields("Telephonne") & " Ext : " & rstContact.Fields("NoPoste")
210                 Else
215                     otlContact.BusinessTelephoneNumber = rstContact.Fields("Telephonne")
220                 End If
                Else
                    otlContact.BusinessTelephoneNumber = rstContact.Fields("Telephonne")
                End If
225         End If

230         If rstContact.Fields("Fax") <> "(___) ___-____" Then
235             otlContact.BusinessFaxNumber = rstContact.Fields("Fax")
240         End If

245         If rstContact.Fields("Cellulaire") <> "(___) ___-____" Then
250             otlContact.MobileTelephoneNumber = rstContact.Fields("Cellulaire")
255         End If

260         If rstContact.Fields("Pagette") <> "(___) ___-____" Then
265             otlContact.PagerNumber = rstContact.Fields("Pagette")
270         End If

274         If Not IsNull(rstContact.Fields("E-mail")) And Not rstContact.Fields("E-mail") = "" Then
275             otlContact.Email1Address = rstContact.Fields("E-mail")
276         End If

280         If rstContact.Fields("TelDomicile") <> "(___) ___-____" Then
285             otlContact.HomeTelephoneNumber = rstContact.Fields("TelDomicile")
290         End If

295         If rstContact.Fields("Commentaire") <> "" Then
300             otlContact.Body = rstContact.Fields("Commentaire")
305         End If

309         Call otlContact.Save

310         rstContact.Fields("DateModification") = ConvertDate(Date)
311         rstContact.Fields("UserModification") = g_sInitiale
            
312         rstContact.Fields("EntryIDOutlook") = otlContact.EntryID
313         rstContact.Update

315         rstContact.MoveNext
316         i = i - 1
317         lblnbre.Caption = "Nombre de contact restant à transférer : " & i
318         Me.Refresh
319     Loop
        
320     Call rstContact.Close
325     Set rstContact = Nothing

330     If bDejaOuvert = False Then
335         Call otlApp.Quit
340     End If

341     Set otlApp = Nothing

342     Screen.MousePointer = vbDefault

343     fraEtatOutlook.Visible = False

345     DoEvents

350     Exit Function

AfficherErreur:

355     woups "frmExportToOutlook", "ExportContactExchange", Err, Erl, "iContactID = " & rstContact.Fields("IDContact"))
356     Call rstContact.Close
357     Set rstContact = Nothing
360     fraEtatOutlook.Visible = False
End Function
Private Function ExportClientExchange(ByVal strQuery As String, ByVal strFolder As String)

5       On Error GoTo AfficherErreur

10      Dim otlApp      As Outlook.Application
15      Dim otlClient   As Outlook.ContactItem
20      Dim folClient   As Outlook.MAPIFolder
25      Dim bDejaOuvert As Boolean
30      Dim sNom()      As String
35      Dim rstClient   As ADODB.Recordset
36      Dim i           As Integer

37      Screen.MousePointer = vbHourglass

40      Set rstClient = New ADODB.Recordset
41      Call rstClient.Open(strQuery, g_connData, adOpenDynamic, adLockOptimistic)

42      i = 0
43      rstClient.MoveFirst
46      Do While Not rstClient.EOF
47          i = i + 1
48          rstClient.MoveNext
49      Loop

50      lblEtatOutlook.Caption = "Ajout des clients dans Outlook ..."
55      lblnbre.Caption = "Nombre de client restant à transférer : " & i
60      fraEtatOutlook.Visible = True

65      Set otlApp = OuvrirOutlook(bDejaOuvert)
70      Set folClient = GetFolder(otlApp, strFolder)

90      rstClient.MoveFirst
100     Do While Not rstClient.EOF

120         Set otlClient = folClient.Items.Add(olContactItem)

130         otlClient.User1 = rstClient.Fields("IDClient")

131         If Not IsNull(rstClient.Fields("NomClient")) Then
132             otlClient.CompanyName = rstClient.Fields("NomClient")
133         End If
    
135         If rstClient.Fields("Telephonne") <> "(___) ___-____" Then
140             otlClient.BusinessTelephoneNumber = rstClient.Fields("Telephonne")
145         End If
  
150         If rstClient.Fields("Fax") <> "(___) ___-____" Then
155             otlClient.BusinessFaxNumber = rstClient.Fields("Fax")
160         End If
164         If Not IsNull(rstClient.Fields("Email")) Then
165             otlClient.Email1Address = rstClient.Fields("Email")
166         End If
169         If Not IsNull(rstClient.Fields("AdresseLiv")) Then
170             otlClient.BusinessAddressStreet = rstClient.Fields("AdresseLiv")
171         End If
174         If Not IsNull(rstClient.Fields("VilleLiv")) Then
175             otlClient.BusinessAddressCity = rstClient.Fields("VilleLiv")
176         End If
179         If Not IsNull(rstClient.Fields("Prov/EtatLiv")) Then
180             otlClient.BusinessAddressState = rstClient.Fields("Prov/EtatLiv")
181         End If
184         If Not IsNull(rstClient.Fields("PaysLiv")) Then
185             otlClient.BusinessAddressCountry = rstClient.Fields("PaysLiv")
186         End If
189         If Not IsNull(rstClient.Fields("CodePostalLiv")) Then
190             otlClient.BusinessAddressPostalCode = rstClient.Fields("CodePostalLiv")
191         End If
194         If Not IsNull(rstClient.Fields("Commentaire")) Then
195             otlClient.Body = rstClient.Fields("Commentaire")
196         End If
199         If Not IsNull(rstClient.Fields("SiteWeb")) Then
200             otlClient.WebPage = rstClient.Fields("SiteWeb")
201         End If

309         Call otlClient.Save

310         rstClient.Fields("DateModification") = ConvertDate(Date)
311         rstClient.Fields("UserModification") = g_sInitiale
            
312         rstClient.Fields("EntryIDOutlook") = otlClient.EntryID
313         rstClient.Update

315         rstClient.MoveNext
316         i = i - 1
317         lblnbre.Caption = "Nombre de client restant à transférer : " & i
318         Me.Refresh
319     Loop
        
320     Call rstClient.Close
325     Set rstClient = Nothing

330     If bDejaOuvert = False Then
335         Call otlApp.Quit
340     End If

341     Set otlApp = Nothing

342     Screen.MousePointer = vbDefault

343     fraEtatOutlook.Visible = False

345     DoEvents

350     Exit Function

AfficherErreur:

355     woups "frmExportToOutlook", "ExportClientExchange", Err, Erl, "iClientID = " & rstClient.Fields("IDClient"))
356     Call rstClient.Close
357     Set rstClient = Nothing
360     fraEtatOutlook.Visible = False
End Function
Private Function ExportFournisseursExchange(ByVal strQuery As String, ByVal strFolder As String)

5       On Error GoTo AfficherErreur

10      Dim otlApp      As Outlook.Application
15      Dim otlFRS      As Outlook.ContactItem
20      Dim folFRS      As Outlook.MAPIFolder
25      Dim bDejaOuvert As Boolean
30      Dim sNom()      As String
35      Dim rstFRS      As ADODB.Recordset
36      Dim i           As Integer

37      Screen.MousePointer = vbHourglass

40      Set rstFRS = New ADODB.Recordset
41      Call rstFRS.Open(strQuery, g_connData, adOpenDynamic, adLockOptimistic)

42      i = 0
43      rstFRS.MoveFirst
46      Do While Not rstFRS.EOF
47          i = i + 1
48          rstFRS.MoveNext
49      Loop

50      lblEtatOutlook.Caption = "Ajout des fournisseurs dans Outlook ..."
55      lblnbre.Caption = "Nombre de fournisseur restant à transférer : " & i
60      fraEtatOutlook.Visible = True

65      Set otlApp = OuvrirOutlook(bDejaOuvert)
70      Set folFRS = GetFolder(otlApp, strFolder)

90      rstFRS.MoveFirst
100     Do While Not rstFRS.EOF

120         Set otlFRS = folFRS.Items.Add(olContactItem)
130         otlFRS.User1 = rstFRS.Fields("IDFRS")

131         If Not IsNull(rstFRS.Fields("NomFournisseur")) Then
132             otlFRS.CompanyName = rstFRS.Fields("NomFournisseur")
133         End If
    
135         If rstFRS.Fields("Telephonne") <> "(___) ___-____" Then
140             otlFRS.BusinessTelephoneNumber = rstFRS.Fields("Telephonne")
145         End If
  
150         If rstFRS.Fields("Fax") <> "(___) ___-____" Then
155             otlFRS.BusinessFaxNumber = rstFRS.Fields("Fax")
160         End If
164         If Not IsNull(rstFRS.Fields("E-mail")) Then
165             otlFRS.Email1Address = rstFRS.Fields("E-mail")
166         End If
169         If Not IsNull(rstFRS.Fields("Adresse")) Then
170             otlFRS.BusinessAddressStreet = rstFRS.Fields("Adresse")
171         End If
174         If Not IsNull(rstFRS.Fields("Ville")) Then
175             otlFRS.BusinessAddressCity = rstFRS.Fields("Ville")
176         End If
179         If Not IsNull(rstFRS.Fields("Prov/Etat")) Then
180             otlFRS.BusinessAddressState = rstFRS.Fields("Prov/Etat")
181         End If
184         If Not IsNull(rstFRS.Fields("Pays")) Then
185             otlFRS.BusinessAddressCountry = rstFRS.Fields("Pays")
186         End If
189         If Not IsNull(rstFRS.Fields("CodePostal")) Then
190             otlFRS.BusinessAddressPostalCode = rstFRS.Fields("CodePostal")
191         End If
194         If Not IsNull(rstFRS.Fields("Commentaire")) Then
195             otlFRS.Body = rstFRS.Fields("Commentaire")
196         End If
199         If Not IsNull(rstFRS.Fields("SiteWeb")) Then
200             otlFRS.WebPage = rstFRS.Fields("SiteWeb")
201         End If

309         Call otlFRS.Save

310         rstFRS.Fields("DateModification") = ConvertDate(Date)
311         rstFRS.Fields("UserModification") = g_sInitiale
            
312         rstFRS.Fields("EntryIDOutlook") = otlFRS.EntryID
313         rstFRS.Update

315         rstFRS.MoveNext
316         i = i - 1
317         lblnbre.Caption = "Nombre de fournisseur restant à transférer : " & i
318         Me.Refresh
319     Loop
        
320     Call rstFRS.Close
325     Set rstFRS = Nothing

330     If bDejaOuvert = False Then
335         Call otlApp.Quit
340     End If

341     Set otlApp = Nothing

342     Screen.MousePointer = vbDefault

343     fraEtatOutlook.Visible = False

345     DoEvents

350     Exit Function

AfficherErreur:

355     woups "frmExportToOutlook", "ExportClientExchange", Err, Erl, "iFRSID = " & rstFRS.Fields("IDFRS"))
356     Call rstFRS.Close
357     Set rstFRS = Nothing
360     fraEtatOutlook.Visible = False
End Function

Private Function SupprimerContactExchange(ByVal strFolder As String, ByVal strName As String)

5       On Error GoTo AfficherErreur

10      Dim otlApp      As Outlook.Application
15      Dim otlContact  As Outlook.ContactItem
20      Dim folContact  As MAPIFolder
25      Dim bDejaOuvert As Boolean
26      Dim i           As Integer

28      Screen.MousePointer = vbHourglass

30      lblEtatOutlook.Caption = "Suppression des " & strName & " dans Outlook ..."
35      fraEtatOutlook.Visible = True

40      Set otlApp = OuvrirOutlook(bDejaOuvert)

45      Set folContact = GetFolder(otlApp, strFolder)

47      i = folContact.Items.count
48      Do While Not folContact.Items.count = 0
            Set otlContact = folContact.Items.GetFirst
54          Call otlContact.Delete
55          i = i - 1
57          lblnbre.Caption = i & " " & strName & " restant à supprimer."
58          Me.Refresh
60      Loop
70      If bDejaOuvert = False Then
75        Call otlApp.Quit
80      End If

85      Set otlApp = Nothing

87      Screen.MousePointer = vbDefault

90      fraEtatOutlook.Visible = False

95      DoEvents

100     Exit Function

AfficherErreur:

105     woups "frmExportToOutlook", "SupprimerContactExchange", Err, Erl

110     fraEtatOutlook.Visible = False
End Function

