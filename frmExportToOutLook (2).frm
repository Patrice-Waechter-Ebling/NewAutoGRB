VERSION 5.00
Begin VB.Form frmExportToOutLook 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export to Outlook"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraEtatOutlook 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   5655
      Begin VB.Label lblnbre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Label5"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   5175
      End
      Begin VB.Label lblEtatOutlook 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
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
 If VerifierSiBesoinExport("SELECT * FROM GrbContact", "Contacts GRB") Then
 Call SupprimerContactExchange("Contacts GRB", "contacts")
 Call ExportContactExchange("SELECT * FROM GrbContact", "Contacts GRB")
 End If
 End If
 If Me.ckClient.Value = 1 Then
 'export les clients
 If VerifierSiBesoinExport("SELECT * FROM GrbClient", "Clients GRB") Then
 Call SupprimerContactExchange("Clients GRB", "clients")
 Call ExportClientExchange("SELECT * FROM GrbClient", "Clients GRB")
 End If
 End If
 If Me.ckFRS.Value = 1 Then
 'export les fournisseurs
 If VerifierSiBesoinExport("SELECT * FROM GrbFournisseur", "Fournisseurs GRB") Then
 Call SupprimerContactExchange("Fournisseurs GRB", "fournisseurs")
 Call ExportFournisseursExchange("SELECT * FROM GrbFournisseur", "Fournisseurs GRB")
 End If
 End If
 
 MsgBox "Exportation des données réussi."
 
 
End Sub
Private Function VerifierSiBesoinExport(ByVal strQuery As String, ByVal strFolder As String) As Boolean

 On Error GoTo Oups

  VerifierSiBesoinExport = False

  Dim dummie As Integer

 Dim otlApp As Outlook.Application
 Dim otlContact As Outlook.ContactItem
 Dim folContact As Outlook.MAPIFolder
 Dim bDejaOuvert As Boolean
 Dim sNom() As String
Dim rstContact As ADODB.Recordset
Dim i As Integer
3 Dim Y As Integer


34 Screen.MousePointer = vbHourglass

 lblEtatOutlook.Caption = "Validation ..."
 lblnbre.Caption = "Vérifier si nous avons besoin de faire l'exportation."
 fraEtatOutlook.Visible = True

 Set rstContact = New ADODB.Recordset
4 Call rstContact.Open(strQuery, g_connData, adOpenDynamic, adLockOptimistic)

4 'nombre total de records dans GRB
44 i = 0
 rstContact.MoveFirst
4  Do While Not rstContact.EOF
4  i = i + 1
4  rstContact.MoveNext
4  Loop

64 'nombre total de records dans Outlook
  Y = 0
6  Set otlApp = OuvrirOutlook(bDejaOuvert)
  Set folContact = GetFolder(otlApp, strFolder)
  Y = folContact.Items.count


  dummie = MsgBox("Nous avons " & i & " records dans GRB et " & Y & " records dans Outlook." & Chr(13) & Chr(13) & _
 "Désirez-vous toujours faire l'exportation dans Outlook?", vbYesNo, "Exportation dans Outlook")

  If dummie = vbYes Then
   VerifierSiBesoinExport = True
   Else
8  VerifierSiBesoinExport = False
8  End If

 
Call rstContact.Close
Set rstContact = Nothing

If bDejaOuvert = False Then
 Call otlApp.Quit
End If

34 Set otlApp = Nothing

34 Screen.MousePointer = vbDefault

34 fraEtatOutlook.Visible = False

DoEvents

Exit Function




Oups:

wOups "frmExportToOutlook", "VerifierSiBesoinExport", Err, Err.number, Err.Description
35  Call rstContact.Close
35  Set rstContact = Nothing

End Function
Private Function ExportContactExchange(ByVal strQuery As String, ByVal strFolder As String)

 On Error GoTo Oups

 Dim otlApp As Outlook.Application
 Dim otlContact As Outlook.ContactItem
 Dim folContact As Outlook.MAPIFolder
 Dim bDejaOuvert As Boolean
 Dim sNom() As String
 Dim rstContact As ADODB.Recordset
 Dim i As Integer

 Screen.MousePointer = vbHourglass

 Set rstContact = New ADODB.Recordset
 Call rstContact.Open(strQuery, g_connData, adOpenDynamic, adLockOptimistic)

4 i = 0
4 rstContact.MoveFirst
4  Do While Not rstContact.EOF
4  i = i + 1
4  rstContact.MoveNext
4  Loop

 lblEtatOutlook.Caption = "Ajout des contacts dans Outlook ..."
 lblnbre.Caption = "Nombre de contact restant à transférer : " & i
  fraEtatOutlook.Visible = True

  Set otlApp = OuvrirOutlook(bDejaOuvert)
  Set folContact = GetFolder(otlApp, strFolder)

  rstContact.MoveFirst
10 Do While Not rstContact.EOF

 Set otlContact = folContact.Items.Add(olContactItem)

 otlContact.User1 = rstContact.Fields("IDContact")

1 If Not IsNull(rstContact.Fields("NomContact")) Then
 sNom = Split(Trim$(rstContact.Fields("NomContact")), " ")

 Select Case UBound(sNom)
 Case 0:
 otlContact.FirstName = sNom(0)
 
 Case 1:
 otlContact.FirstName = sNom(0)
 otlContact.LastName = sNom(1)

 Case 2:
 otlContact.FirstName = sNom(0)
 otlContact.MiddleName = sNom(1)
 otlContact.LastName = sNom(2)
 End Select
1   End If

 otlContact.Title = ""

 If Not IsNull(rstContact.Fields("Compagnie")) Then
184 otlContact.CompanyName = rstContact.Fields("Compagnie")
 End If
1   If Not IsNull(rstContact.Fields("Titre")) Or Not rstContact.Fields("Titre") = "" Then
 otlContact.JobTitle = rstContact.Fields("Titre")
1  End If

1  If rstContact.Fields("Telephonne") <> "(___) ___-____" Then
 If Not IsNull(rstContact.Fields("NoPoste")) Then
 If Trim$(rstContact.Fields("NoPoste")) <> "" Then
 otlContact.BusinessTelephoneNumber = rstContact.Fields("Telephonne") & " Ext : " & rstContact.Fields("NoPoste")
 Else
 otlContact.BusinessTelephoneNumber = rstContact.Fields("Telephonne")
 End If
 Else
 otlContact.BusinessTelephoneNumber = rstContact.Fields("Telephonne")
 End If
 End If

 If rstContact.Fields("Fax") <> "(___) ___-____" Then
 otlContact.BusinessFaxNumber = rstContact.Fields("Fax")
 End If

 If rstContact.Fields("Cellulaire") <> "(___) ___-____" Then
 otlContact.MobileTelephoneNumber = rstContact.Fields("Cellulaire")
 End If

 If rstContact.Fields("Pagette") <> "(___) ___-____" Then
 otlContact.PagerNumber = rstContact.Fields("Pagette")
 End If

274 If Not IsNull(rstContact.Fields("E-mail")) And Not rstContact.Fields("E-mail") = "" Then
 otlContact.Email1Address = rstContact.Fields("E-mail")
2   End If

 If rstContact.Fields("TelDomicile") <> "(___) ___-____" Then
 otlContact.HomeTelephoneNumber = rstContact.Fields("TelDomicile")
 End If

 If rstContact.Fields("Commentaire") <> "" Then
 otlContact.Body = rstContact.Fields("Commentaire")
End If

30  Call otlContact.Save

 rstContact.Fields("DateModification") = ConvertDate(Date)
31 rstContact.Fields("UserModification") = g_sInitiale
 
31 rstContact.Fields("EntryIDOutlook") = otlContact.EntryID
31 rstContact.Update

 rstContact.MoveNext
 i = i - 1
 lblnbre.Caption = "Nombre de contact restant à transférer : " & i
 Me.Refresh
31  Loop
 
Call rstContact.Close
Set rstContact = Nothing

If bDejaOuvert = False Then
 Call otlApp.Quit
End If

34 Set otlApp = Nothing

34 Screen.MousePointer = vbDefault

34 fraEtatOutlook.Visible = False

DoEvents

Exit Function

Oups:

woups"frmExportToOutlook", "ExportContactExchange", Err, Erl, "iContactID = " & rstContact.Fields("IDContact"))
35  Call rstContact.Close
35  Set rstContact = Nothing
3  fraEtatOutlook.Visible = False
End Function
Private Function ExportClientExchange(ByVal strQuery As String, ByVal strFolder As String)

 On Error GoTo Oups

 Dim otlApp As Outlook.Application
 Dim otlClient As Outlook.ContactItem
 Dim folClient As Outlook.MAPIFolder
 Dim bDejaOuvert As Boolean
 Dim sNom() As String
 Dim rstClient As ADODB.Recordset
 Dim i As Integer

 Screen.MousePointer = vbHourglass

 Set rstClient = New ADODB.Recordset
4 Call rstClient.Open(strQuery, g_connData, adOpenDynamic, adLockOptimistic)

4 i = 0
4 rstClient.MoveFirst
4  Do While Not rstClient.EOF
4  i = i + 1
4  rstClient.MoveNext
4  Loop

 lblEtatOutlook.Caption = "Ajout des clients dans Outlook ..."
 lblnbre.Caption = "Nombre de client restant à transférer : " & i
  fraEtatOutlook.Visible = True

  Set otlApp = OuvrirOutlook(bDejaOuvert)
  Set folClient = GetFolder(otlApp, strFolder)

  rstClient.MoveFirst
10 Do While Not rstClient.EOF

 Set otlClient = folClient.Items.Add(olContactItem)

 otlClient.User1 = rstClient.Fields("IDClient")

1 If Not IsNull(rstClient.Fields("NomClient")) Then
1 otlClient.CompanyName = rstClient.Fields("NomClient")
13 End If
 
 If rstClient.Fields("Telephonne") <> "(___) ___-____" Then
 otlClient.BusinessTelephoneNumber = rstClient.Fields("Telephonne")
 End If
 
 If rstClient.Fields("Fax") <> "(___) ___-____" Then
 otlClient.BusinessFaxNumber = rstClient.Fields("Fax")
 End If
164 If Not IsNull(rstClient.Fields("Email")) Then
 otlClient.Email1Address = rstClient.Fields("Email")
16  End If
16  If Not IsNull(rstClient.Fields("AdresseLiv")) Then
 otlClient.BusinessAddressStreet = rstClient.Fields("AdresseLiv")
 End If
174 If Not IsNull(rstClient.Fields("VilleLiv")) Then
 otlClient.BusinessAddressCity = rstClient.Fields("VilleLiv")
1   End If
17  If Not IsNull(rstClient.Fields("Prov/EtatLiv")) Then
 otlClient.BusinessAddressState = rstClient.Fields("Prov/EtatLiv")
 End If
184 If Not IsNull(rstClient.Fields("PaysLiv")) Then
 otlClient.BusinessAddressCountry = rstClient.Fields("PaysLiv")
1   End If
18  If Not IsNull(rstClient.Fields("CodePostalLiv")) Then
 otlClient.BusinessAddressPostalCode = rstClient.Fields("CodePostalLiv")
1  End If
194 If Not IsNull(rstClient.Fields("Commentaire")) Then
1  otlClient.Body = rstClient.Fields("Commentaire")
1   End If
19  If Not IsNull(rstClient.Fields("SiteWeb")) Then
 otlClient.WebPage = rstClient.Fields("SiteWeb")
20 End If

30  Call otlClient.Save

 rstClient.Fields("DateModification") = ConvertDate(Date)
31 rstClient.Fields("UserModification") = g_sInitiale
 
31 rstClient.Fields("EntryIDOutlook") = otlClient.EntryID
31 rstClient.Update

 rstClient.MoveNext
 i = i - 1
 lblnbre.Caption = "Nombre de client restant à transférer : " & i
 Me.Refresh
31  Loop
 
Call rstClient.Close
Set rstClient = Nothing

If bDejaOuvert = False Then
 Call otlApp.Quit
End If

34 Set otlApp = Nothing

34 Screen.MousePointer = vbDefault

34 fraEtatOutlook.Visible = False

DoEvents

Exit Function

Oups:

woups"frmExportToOutlook", "ExportClientExchange", Err, Erl, "iClientID = " & rstClient.Fields("IDClient"))
35  Call rstClient.Close
35  Set rstClient = Nothing
3  fraEtatOutlook.Visible = False
End Function
Private Function ExportFournisseursExchange(ByVal strQuery As String, ByVal strFolder As String)

 On Error GoTo Oups

 Dim otlApp As Outlook.Application
 Dim otlFRS As Outlook.ContactItem
 Dim folFRS As Outlook.MAPIFolder
 Dim bDejaOuvert As Boolean
 Dim sNom() As String
 Dim rstFRS As ADODB.Recordset
 Dim i As Integer

 Screen.MousePointer = vbHourglass

 Set rstFRS = New ADODB.Recordset
4 Call rstFRS.Open(strQuery, g_connData, adOpenDynamic, adLockOptimistic)

4 i = 0
4 rstFRS.MoveFirst
4  Do While Not rstFRS.EOF
4  i = i + 1
4  rstFRS.MoveNext
4  Loop

 lblEtatOutlook.Caption = "Ajout des fournisseurs dans Outlook ..."
 lblnbre.Caption = "Nombre de fournisseur restant à transférer : " & i
  fraEtatOutlook.Visible = True

  Set otlApp = OuvrirOutlook(bDejaOuvert)
  Set folFRS = GetFolder(otlApp, strFolder)

  rstFRS.MoveFirst
10 Do While Not rstFRS.EOF

 Set otlFRS = folFRS.Items.Add(olContactItem)
 otlFRS.User1 = rstFRS.Fields("IDFRS")

1 If Not IsNull(rstFRS.Fields("NomFournisseur")) Then
1 otlFRS.CompanyName = rstFRS.Fields("NomFournisseur")
13 End If
 
 If rstFRS.Fields("Telephonne") <> "(___) ___-____" Then
 otlFRS.BusinessTelephoneNumber = rstFRS.Fields("Telephonne")
 End If
 
 If rstFRS.Fields("Fax") <> "(___) ___-____" Then
 otlFRS.BusinessFaxNumber = rstFRS.Fields("Fax")
 End If
164 If Not IsNull(rstFRS.Fields("E-mail")) Then
 otlFRS.Email1Address = rstFRS.Fields("E-mail")
16  End If
16  If Not IsNull(rstFRS.Fields("Adresse")) Then
 otlFRS.BusinessAddressStreet = rstFRS.Fields("Adresse")
 End If
174 If Not IsNull(rstFRS.Fields("Ville")) Then
 otlFRS.BusinessAddressCity = rstFRS.Fields("Ville")
1   End If
17  If Not IsNull(rstFRS.Fields("Prov/Etat")) Then
 otlFRS.BusinessAddressState = rstFRS.Fields("Prov/Etat")
 End If
184 If Not IsNull(rstFRS.Fields("Pays")) Then
 otlFRS.BusinessAddressCountry = rstFRS.Fields("Pays")
1   End If
18  If Not IsNull(rstFRS.Fields("CodePostal")) Then
 otlFRS.BusinessAddressPostalCode = rstFRS.Fields("CodePostal")
1  End If
194 If Not IsNull(rstFRS.Fields("Commentaire")) Then
1  otlFRS.Body = rstFRS.Fields("Commentaire")
1   End If
19  If Not IsNull(rstFRS.Fields("SiteWeb")) Then
 otlFRS.WebPage = rstFRS.Fields("SiteWeb")
20 End If

30  Call otlFRS.Save

 rstFRS.Fields("DateModification") = ConvertDate(Date)
31 rstFRS.Fields("UserModification") = g_sInitiale
 
31 rstFRS.Fields("EntryIDOutlook") = otlFRS.EntryID
31 rstFRS.Update

 rstFRS.MoveNext
 i = i - 1
 lblnbre.Caption = "Nombre de fournisseur restant à transférer : " & i
 Me.Refresh
31  Loop
 
Call rstFRS.Close
Set rstFRS = Nothing

If bDejaOuvert = False Then
 Call otlApp.Quit
End If

34 Set otlApp = Nothing

34 Screen.MousePointer = vbDefault

34 fraEtatOutlook.Visible = False

DoEvents

Exit Function

Oups:

woups"frmExportToOutlook", "ExportClientExchange", Err, Erl, "iFRSID = " & rstFRS.Fields("IDFRS"))
35  Call rstFRS.Close
35  Set rstFRS = Nothing
3  fraEtatOutlook.Visible = False
End Function

Private Function SupprimerContactExchange(ByVal strFolder As String, ByVal strName As String)

 On Error GoTo Oups

 Dim otlApp As Outlook.Application
 Dim otlContact As Outlook.ContactItem
 Dim folContact As MAPIFolder
 Dim bDejaOuvert As Boolean
 Dim i As Integer

 Screen.MousePointer = vbHourglass

 lblEtatOutlook.Caption = "Suppression des " & strName & " dans Outlook ..."
 fraEtatOutlook.Visible = True

 Set otlApp = OuvrirOutlook(bDejaOuvert)

 Set folContact = GetFolder(otlApp, strFolder)

4  i = folContact.Items.count
4  Do While Not folContact.Items.count = 0
 Set otlContact = folContact.Items.GetFirst
54 Call otlContact.Delete
 i = i - 1
5  lblnbre.Caption = i & " " & strName & " restant à supprimer."
5  Me.Refresh
  Loop
  If bDejaOuvert = False Then
  Call otlApp.Quit
  End If

  Set otlApp = Nothing

   Screen.MousePointer = vbDefault

  fraEtatOutlook.Visible = False

  DoEvents

10 Exit Function

Oups:

wOups "frmExportToOutlook", "SupprimerContactExchange", Err, Err.number, Err.Description

fraEtatOutlook.Visible = False
End Function

