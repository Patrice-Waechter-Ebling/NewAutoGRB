VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChoixImpressionFT 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impression des feuilles de temps"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4260
   ControlBox      =   0   'False
   Icon            =   "frmChoixImpressionFT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   4260
   Begin VB.CommandButton cmdImprimer 
      Caption         =   "Imprimer"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdAucun 
      Caption         =   "Aucun"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdTous 
      Caption         =   "Tous"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5520
      Width           =   975
   End
   Begin MSComctlLib.ListView lvwEmploye 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nom"
         Object.Width           =   6879
      EndProperty
   End
End
Attribute VB_Name = "frmChoixImpressionFT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_datDebut As Date
Private m_datFin As Date

Public Sub Afficher(ByVal datDebut As Date, ByVal datFin As Date)
 
 On Error GoTo Oups

 m_datDebut = datDebut
 m_datFin = datFin
 
 Call RemplirListViewEmploye

 Call Me.Show(vbModal)

 Exit Sub

Oups:

 wOups "frmChoixImpressionFT", "Afficher", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnuler_Click()
 
 On Error GoTo Oups

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixImpressionFT", "cmdAnnuler_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdImprimer_Click()

 On Error GoTo Oups

 'Impression de la feuille de temps
 Dim rstImpPunch As ADODB.Recordset
 Dim rstSommeKM As ADODB.Recordset
 Dim rstPunch As ADODB.Recordset
 Dim bAnnuler As Boolean
 Dim iCompteur As Integer
 Dim dblTotal As Double

 Screen.MousePointer = vbHourglass

 Set rstImpPunch = New ADODB.Recordset
 Set rstSommeKM = New ADODB.Recordset
 Set rstPunch = New ADODB.Recordset

  For iCompteur = 1 To lvwEmploye.ListItems.count
  If lvwEmploye.ListItems(iCompteur).Checked = True Then
  dblTotal = 0

  bAnnuler = False

  Call g_connData.Execute("DELETE * FROM GrbImpressionPunch")

  Call rstPunch.Open("SELECT * FROM GrbPunch WHERE Date BETWEEN '" & ConvertDate(m_datDebut) & "' AND '" & ConvertDate(m_datFin) & "' AND NoEmploye = " & lvwEmploye.ListItems(iCompteur).Tag & " ORDER BY Date, HeureDébut, HeureFin", g_connData, adOpenDynamic, adLockOptimistic)

 'Vérifie s'il y a des punchs
  If Not rstPunch.EOF Then
 'Vérifie si le punch out a été fait
  Do While Not rstPunch.EOF
 If IsNull(rstPunch.Fields("HeureFin")) Or rstPunch.Fields("HeureFin") = vbNullString Then
 Call MsgBox("Un punch out n'a pas été fait pour " & lvwEmploye.ListItems(iCompteur).Text & " !", vbOKOnly, "Erreur")

 bAnnuler = True

 Exit Do
 End If

 Call rstPunch.MoveNext
 Loop
 Else
 Screen.MousePointer = vbDefault

 Call MsgBox("Il n'y a aucun punch à imprimer pour " & lvwEmploye.ListItems(iCompteur).Text & " !", vbOKOnly, "Erreur")

 bAnnuler = True
 End If

 Call rstPunch.Close

 If bAnnuler = False Then
 Call RemplirTableImpressionPunch(iCompteur)

 Call AjouterNomJour


 Call AjouterSéparateur

 Call CalculerTotal

 'Le nom de l'employé
 DR_FeuilleTemps.Sections("Section4").Controls("lblNom").Caption = lvwEmploye.ListItems(iCompteur).Text

 'La date de début et de fin de la semaine
1  DR_FeuilleTemps.Sections("Section4").Controls("lblDate").Caption = "Semaine du " & GetDateTexte(m_datDebut) & " au " & GetDateTexte(m_datFin)

 'Date d'aujourd'hui
 DR_FeuilleTemps.Sections("Section3").Controls("lblDatePrint").Caption = ConvertDate(Date)

 Call rstImpPunch.Open("SELECT * FROM GrbImpressionPunch ORDER BY Date, HeureDébut", g_connData, adOpenDynamic, adLockOptimistic)

 Do While Not rstImpPunch.EOF
 If Not IsNull(rstImpPunch.Fields("Total")) Then
 If rstImpPunch.Fields("Total") <> "" Then
 dblTotal = dblTotal + CDbl(Left(rstImpPunch.Fields("Total"), 2))
 dblTotal = dblTotal + CDbl(CDbl(Right(rstImpPunch.Fields("Total"), 2)) / 60)
 End If
 End If

 Call rstImpPunch.MoveNext
 Loop

 Call rstImpPunch.MoveFirst

 Call rstSommeKM.Open("SELECT SUM(NbreKM) As TotalKM FROM GrbPunch WHERE Date BETWEEN '" & ConvertDate(m_datDebut) & "' AND '" & ConvertDate(m_datFin) & "' AND NoEmploye = " & lvwEmploye.ListItems(iCompteur).Tag & " AND KM = True", g_connData, adOpenDynamic, adLockOptimistic)
 
 Set DR_FeuilleTemps.DataSource = rstImpPunch

 'Le total d'heure dans une semaine
 DR_FeuilleTemps.Sections("Section5").Controls("lblGrandTotal").Caption = Round(dblTotal, 2)

 If Not IsNull(rstSommeKM.Fields("TotalKM")) Then
 DR_FeuilleTemps.Sections("Section5").Controls("lblGrandTotalKM").Caption = Round(rstSommeKM.Fields("TotalKM"), 2)
 Else
 DR_FeuilleTemps.Sections("Section5").Controls("lblGrandTotalKM").Caption = "0"
 End If

 DR_FeuilleTemps.Orientation = rptOrientLandscape

 Call DR_FeuilleTemps.Show(vbModal)
 
 Call rstImpPunch.Close
 Call rstSommeKM.Close
 End If
 End If
Next

Screen.MousePointer = vbDefault

Call Unload(Me)

Exit Sub

Oups:

wOups "frmChoixImpressionFT", "cmdImprimer_Click", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewEmploye()
 
 On Error GoTo Oups

 Dim rstEmploye As ADODB.Recordset
 Dim itmEmploye As ListItem
 
 Set rstEmploye = New ADODB.Recordset
 
 Call rstEmploye.Open("SELECT NoEmploye, Employe FROM GrbEmployés WHERE Actif = True ORDER BY Employe", g_connData, adOpenForwardOnly, adLockReadOnly)
 
 Do While Not rstEmploye.EOF
 Set itmEmploye = lvwEmploye.ListItems.Add
 
 itmEmploye.Tag = rstEmploye.Fields("NoEmploye")
 itmEmploye.Text = rstEmploye.Fields("Employe")
 
 Call rstEmploye.MoveNext
 Loop
 
  Call rstEmploye.Close
  Set rstEmploye = Nothing

  Exit Sub

Oups:

  wOups "frmChoixImpressionFT", "RemplirListViewEmploye", Err, Err.number, Err.Description
End Sub

Private Sub RemplirTableImpressionPunch(ByVal iIndexListView As Integer)

 On Error GoTo Oups

 Dim rstImpPunch As ADODB.Recordset
 Dim rstPunch As ADODB.Recordset
 Dim rstClient As ADODB.Recordset
 
 Set rstPunch = New ADODB.Recordset
 Set rstImpPunch = New ADODB.Recordset
 Set rstClient = New ADODB.Recordset
 
 Call rstPunch.Open("SELECT * FROM GrbPunch WHERE NoEmploye = " & lvwEmploye.ListItems(iIndexListView).Tag & " AND Date BETWEEN '" & ConvertDate(m_datDebut) & "' AND '" & ConvertDate(m_datFin) & "' ORDER BY Date, HeureDébut", g_connData, adOpenDynamic, adLockOptimistic)
 
 Call rstImpPunch.Open("SELECT * FROM GrbImpressionPunch", g_connData, adOpenDynamic, adLockOptimistic)
 
 Do While Not rstPunch.EOF
 Call rstImpPunch.AddNew
 
  rstImpPunch.Fields("NoProjet") = rstPunch.Fields("NoProjet")
  rstImpPunch.Fields("Date") = rstPunch.Fields("Date")
 
  If rstPunch.Fields("ModifDébut") = True Then
  rstImpPunch.Fields("HeureDébut") = Right$("0" & rstPunch.Fields("HeureDébut"), 5) & "*"
  Else
  rstImpPunch.Fields("HeureDébut") = Right$("0" & rstPunch.Fields("HeureDébut"), 5)
  End If
 
  If rstPunch.Fields("ModifFin") = True Then
 rstImpPunch.Fields("HeureFin") = Right$("0" & rstPunch.Fields("HeureFin"), 5) & "*"
1 Else
 rstImpPunch.Fields("HeureFin") = Right$("0" & rstPunch.Fields("HeureFin"), 5)
 End If
 
 If Not IsNull(rstPunch.Fields("NoClient")) And rstPunch.Fields("NoClient") > "0" Then
 Call rstClient.Open("SELECT NomClient FROM GrbClient WHERE IDClient = " & rstPunch.Fields("NoClient"), g_connData, adOpenDynamic, adLockOptimistic)
 
 rstImpPunch.Fields("Client") = rstClient.Fields("NomClient")
 
 Call rstClient.Close
 End If
 
 rstImpPunch.Fields("Commentaire") = rstPunch.Fields("Commentaire")


 '**************************************************************************
 'MODIFIÉ PAR GAÉTAN GINGRAS LE 0  FÉVRIER 2010
 '**************************************************************************
 'ajout
 rstImpPunch.Fields("Type") = rstPunch.Fields("Type")
 '**************************************************************************
 
 
 rstImpPunch.Fields("NbreKM") = rstPunch.Fields("NbreKM")
 
 Call rstImpPunch.Update
 
Call rstPunch.MoveNext
Loop
 
 Call rstPunch.Close
Set rstPunch = Nothing
 
 Call rstImpPunch.Close
Set rstImpPunch = Nothing

 Set rstClient = Nothing

1  Exit Sub

Oups:

 wOups "frmChoixImpressionFT", "RemplirTableImpressionPunch", Err, Err.number, Err.Description
End Sub

Private Sub AjouterNomJour()

 On Error GoTo Oups

 Dim rstImpPunch As ADODB.Recordset
 Dim sJour As String
 Dim datTemp As Date
 Dim sDate As String
 
 'Ouverture du recordset pour ajouter le nom de la journée
 Set rstImpPunch = New ADODB.Recordset
 
 rstImpPunch.CursorLocation = adUseServer
 
 Call rstImpPunch.Open("SELECT * FROM GrbImpressionPunch ORDER BY Date, HeureDébut", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Loop pour ajouter le nom de la journée
 Do While Not rstImpPunch.EOF
 sDate = rstImpPunch.Fields("Date")
 
 datTemp = DateSerial(Left$(sDate, 4), Mid$(sDate, 6, 2), Right$(sDate, 2))
 
  If sJour <> WeekdayName(Weekday(datTemp)) Then
  rstImpPunch.Fields("NomJour") = UCase(WeekdayName(Weekday(datTemp)))
  sJour = WeekdayName(Weekday(datTemp))
  Else
  rstImpPunch.Fields("NomJour") = ""
  End If
 
  Call rstImpPunch.Update
 
  Call rstImpPunch.MoveNext
10 Loop
 
Call rstImpPunch.Close
Set rstImpPunch = Nothing

Exit Sub

Oups:

wOups "frmChoixImpressionFT", "AjouterNomJour", Err, Err.number, Err.Description
End Sub

Private Sub AjouterSéparateur()

 On Error GoTo Oups

 Dim rstImpPunch As ADODB.Recordset
 Dim iNoRec As Integer
 Dim sJour As String
 Dim collDate As Collection
 Dim sDate As String
 Dim iCompteur As Integer
 
 Set collDate = New Collection
 
 Set rstImpPunch = New ADODB.Recordset
 
 Call rstImpPunch.Open("SELECT * FROM GrbImpressionPunch ORDER BY Date, HeureDébut", g_connData, adOpenDynamic, adLockOptimistic)
 
 iNoRec = 1
 
  Do While Not rstImpPunch.EOF
  If sJour <> rstImpPunch.Fields("NomJour") Then
  If rstImpPunch.Fields("NomJour") <> vbNullString Then
  If iNoRec > 1 Then
  sDate = rstImpPunch.Fields("Date")
 
  Call collDate.Add(sDate)
  End If
  End If
 
 If rstImpPunch.Fields("NomJour") <> vbNullString Then
 sJour = rstImpPunch.Fields("NomJour")
 End If
 End If
 
 iNoRec = iNoRec + 1
 
 Call rstImpPunch.MoveNext
Loop
 
For iCompteur = 1 To collDate.count
 Call rstImpPunch.AddNew
 
 rstImpPunch.Fields("Date") = collDate(iCompteur)
 rstImpPunch.Fields("NomJour") = vbNullString
 rstImpPunch.Fields("NoProjet") = " "
rstImpPunch.Fields("Client") = vbNullString
 rstImpPunch.Fields("Commentaire") = vbNullString
 rstImpPunch.Fields("HeureDébut") = " "
 rstImpPunch.Fields("HeureFin") = vbNullString
 
 Call rstImpPunch.Update
Next
 
 Call rstImpPunch.Close
1  Set rstImpPunch = Nothing

 Exit Sub

Oups:

 wOups "frmChoixImpressionFT", "AjouterSéparateur", Err, Err.number, Err.Description
End Sub

Private Sub CalculerTotal()

 On Error GoTo Oups

 'Calcul le temps entre chaque punch in et punch out
 Dim rstImpPunch As ADODB.Recordset
 Dim datDébut As Date
 Dim datFin As Date
 Dim datTotal As Date
 Dim sDate As String
 Dim sDébut As String
 Dim sFin As String
 
 'Ouverture de tous les punchs à imprimer
 Set rstImpPunch = New ADODB.Recordset

 rstImpPunch.CursorLocation = adUseServer
 
 Call rstImpPunch.Open("SELECT * FROM GrbImpressionPunch", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Tant que ce n'est pas la fin des enregistrements
  Do While Not rstImpPunch.EOF
 'Si ce n'est pas un séparateur
  If Trim(rstImpPunch.Fields("HeureDébut")) <> vbNullString Then
  sDate = rstImpPunch.Fields("Date")
  sDébut = Left(rstImpPunch.Fields("HeureDébut"), 5)
  sFin = Left(rstImpPunch.Fields("HeureFin"), 5)
 
  datDébut = DateSerial(Left$(sDate, 4), Mid$(sDate, 6, 2), Right$(sDate, 2)) + TimeSerial(Left$(sDébut, 2), Mid$(sDébut, 4, 2), 0)
  datFin = DateSerial(Left$(sDate, 4), Mid$(sDate, 6, 2), Right$(sDate, 2)) + TimeSerial(Left$(sFin, 2), Mid$(sFin, 4, 2), 0)
 
  If Hour(datDébut) = 0 And Minute(datDébut) = 0 And Second(datDébut) = 0 And _
 Hour(datFin) = 0 And Minute(datFin) = 0 And Second(datFin) = 0 And _
 DateDiff("d", datDébut, datFin) = 1 Then
 rstImpPunch.Fields("Total") = "24:00"
Else
 datTotal = datFin - datDébut
 
 rstImpPunch.Fields("Total") = Right$("0" & Hour(datTotal), 2) & ":" & Right$("0" & Minute(datTotal), 2)
 End If

 Call rstImpPunch.Update
 End If
 
 Call rstImpPunch.MoveNext
Loop

Exit Sub

Oups:

wOups "frmChoixImpressionFT", "CalculerTotal", Err, Err.number, Err.Description
End Sub

Private Sub cmdTous_Click()
 
 On Error GoTo Oups

 Dim iCompteur As Integer

 For iCompteur = 1 To lvwEmploye.ListItems.count
 lvwEmploye.ListItems(iCompteur).Checked = True
 Next

 Exit Sub

Oups:

 wOups "frmChoixImpressionFT", "cmdTous_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAucun_Click()
 
 On Error GoTo Oups

 Dim iCompteur As Integer
 
 For iCompteur = 1 To lvwEmploye.ListItems.count
 lvwEmploye.ListItems(iCompteur).Checked = False
 Next

 Exit Sub

Oups:

 wOups "frmChoixImpressionFT", "cmdAucun_Click", Err, Err.number, Err.Description
End Sub
