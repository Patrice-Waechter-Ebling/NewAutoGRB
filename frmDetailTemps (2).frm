VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDetailTemps 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Détail des temps"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImprimer 
      Caption         =   "Imprimer"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   4440
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwTemps 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7223
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Employé"
         Object.Width           =   6033
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Heures"
         Object.Width           =   5477
      EndProperty
   End
End
Attribute VB_Name = "frmDetailTemps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const I_COL_EMPLOYE As Integer = 0
Private Const I_COL_TYPE As Integer = 1
Private Const I_COL_HEURES As Integer = 2

Private m_sNoProjet As String
Private m_bProjet As Boolean
Private m_eType As enumCatalogue

Public Sub Afficher(ByVal sNoProjet As String, ByVal eType As enumCatalogue, ByVal bProjet As Boolean)
 
 On Error GoTo Oups

 m_eType = eType
 
 m_sNoProjet = sNoProjet

 m_bProjet = bProjet
 
 Call RemplirListViewTemps(sNoProjet)

 Call Show(vbModal)

 Exit Sub

Oups:
 
 wOups "frmDetailTemps", "Afficher", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewTemps(ByVal sNoProjet As String)

 On Error GoTo Oups

 Dim rstPunch As ADODB.Recordset
 Dim itmPunch As ListItem
 Dim sFilterNoProjet As String

 If Right$(sNoProjet, 2) = "99" Then
 sFilterNoProjet = "LEFT(NoProjet, 6) = '" & Left$(sNoProjet, 6) & "'"
 Else
 sFilterNoProjet = "NoProjet = '" & sNoProjet & "'"
 End If

 Set rstPunch = New ADODB.Recordset

 Call rstPunch.Open("SELECT Employe, Type, (Sum(TimeSerial(LEFT(HeureFin, 2), RIGHT(HeureFin, 2), 0) - TimeSerial(LEFT(HeureDébut, 2), RIGHT(HeureDébut, 2), 0)) * 24) AS TotalHeures FROM GrbPunch INNER JOIN GrbEmployés ON GrbPunch.NoEmploye = GrbEmployés.NoEmploye WHERE HeureDébut is Not Null And HeureFin is Not Null AND " & sFilterNoProjet & " GROUP BY Employe, Type", g_connData, adOpenDynamic, adLockOptimistic)

  Do While Not rstPunch.EOF
  Set itmPunch = lvwTemps.ListItems.Add

  itmPunch.Text = rstPunch.Fields("Employe")
 'Retire pour afficher le type vu au lieu des spécifique GLL
 
  'If Not IsNull(rstPunch.Fields("Type")) Then
  ' If m_eType = ELECTRIQUE Then
  ' Select Case rstPunch.Fields("Type")
 ' Case "Dessin": itmPunch.SubItems(I_COL_TYPE) = "Dessin"
  ' Case "Fabrication": itmPunch.SubItems(I_COL_TYPE) = "Fabrication"
  ' Case "Assemblage": itmPunch.SubItems(I_COL_TYPE) = "Assemblage"
' Case "ProgInterface": itmPunch.SubItems(I_COL_TYPE) = "Programmation d'interface"
1 ' Case "ProgAutomate": itmPunch.SubItems(I_COL_TYPE) = "Programmation d'automate"
 ' Case "ProgRobot": itmPunch.SubItems(I_COL_TYPE) = "Programmation de robot"
 ' Case "Vision": itmPunch.SubItems(I_COL_TYPE) = "Vision"
 ' Case "Test": itmPunch.SubItems(I_COL_TYPE) = "Test"
 ' Case "Installation": itmPunch.SubItems(I_COL_TYPE) = "Installation"
 ' Case "MiseService": itmPunch.SubItems(I_COL_TYPE) = "Mise en service"
 ' Case "Formation": itmPunch.SubItems(I_COL_TYPE) = "Formation du personnel"
 ' Case "Gestion": itmPunch.SubItems(I_COL_TYPE) = "Gestion du projet"
 ' Case "Shipping": itmPunch.SubItems(I_COL_TYPE) = "Expédition"
 ' Case "Prototypage-Dévelloppement expérimental": itmPunch.SubItems(I_COL_TYPE) = "Prototypage-Dévelloppement expérimental"
 ' End Select
 ' Else
' Select Case rstPunch.Fields("Type")
 ' Case "Dessin": itmPunch.SubItems(I_COL_TYPE) = "Conception et dessins"
 ' Case "Coupe": itmPunch.SubItems(I_COL_TYPE) = "Coupe et préparation (sauf soudage)"
 ' Case "Machinage": itmPunch.SubItems(I_COL_TYPE) = "Machinage"
 ' Case "Soudure": itmPunch.SubItems(I_COL_TYPE) = "Coupe, soudure et meulage"
 ' Case "Assemblage": itmPunch.SubItems(I_COL_TYPE) = "Assemblage des systèmes"
 ' Case "Peinture": itmPunch.SubItems(I_COL_TYPE) = "Peinture et finition"
 ' Case "Test": itmPunch.SubItems(I_COL_TYPE) = "Tests finaux"
1  ' Case "Installation": itmPunch.SubItems(I_COL_TYPE) = "Installation"
 ' Case "Formation": itmPunch.SubItems(I_COL_TYPE) = "Formation du formation"
 ' Case "Gestion": itmPunch.SubItems(I_COL_TYPE) = "Gestion du projet"
 ' Case "Shipping": itmPunch.SubItems(I_COL_TYPE) = "Expédition"
 ' Case "Prototypage-Dévelloppement expérimental": itmPunch.SubItems(I_COL_TYPE) = "Prototypage-Dévelloppement expérimental"
 ' End Select
 ' End If
 'Else
 itmPunch.SubItems(I_COL_TYPE) = rstPunch.Fields("Type")
 'End If
 
 itmPunch.SubItems(I_COL_HEURES) = Round(rstPunch.Fields("TotalHeures"), 2)

 Call rstPunch.MoveNext
Loop

Call rstPunch.Close
2  Set rstPunch = Nothing

Exit Sub

Oups:

2  wOups "frmDetailTemps", "RemplirListViewTemps", Err, Err.number, Err.Description
End Sub

Private Sub cmdImprimer_Click()
Dim intdummie As Integer

 On Error GoTo Oups

 If m_eType = ELECTRIQUE Then

 'demande d'ecriture dans excel
 intdummie = MsgBox("Désirez-vous exporter les données dans Excel ?", vbYesNo + vbInformation, "Exportation dans Excel")
 If intdummie = vbYes Then
 Call vb_to_excel
 
 End If

 Call ImprimerDetailTempsElectriques
 Else

 'demande d'ecriture dans excel
 intdummie = MsgBox("Désirez-vous exporter les données dans Excel ?", vbYesNo + vbInformation, "Exportation dans Excel")
 If intdummie = vbYes Then
 Call vb_to_excel
 
 End If

 Call ImprimerDetailTempsMecaniques
 End If



 Exit Sub

Oups:

 wOups "frmDetailTemps", "cmdImprimer_Click", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerDetailTempsElectriques()
 
 On Error GoTo Oups

 Dim rstEmploye As ADODB.Recordset
 Dim rstImpTemps As ADODB.Recordset
 Dim rstProjSoum As ADODB.Recordset
 Dim dblTotal As Double
 Dim sFilterNoProjet As String

 If Right$(m_sNoProjet, 2) = "99" Then
 sFilterNoProjet = "LEFT(NoProjet, 6) = '" & Left$(m_sNoProjet, 6) & "'"
 Else
 sFilterNoProjet = "NoProjet = '" & m_sNoProjet & "'"
 End If

  Set rstEmploye = New ADODB.Recordset

  Call rstEmploye.Open("SELECT Employe, Type, (Sum(TimeSerial(Left(HeureFin,2), RIGHT(HeureFin,2),0) - TimeSerial(Left(HeureDébut,2), RIGHT(HeureDébut,2),0)) * 24) AS TotalHeures FROM GrbPunch INNER JOIN GrbEmployés ON GrbPunch.NoEmploye = GrbEmployés.NoEmploye WHERE HeureDébut is Not Null And HeureFin is Not Null AND " & sFilterNoProjet & " GROUP BY Employe, Type", g_connData, adOpenDynamic, adLockOptimistic)

  Call g_connData.Execute("DELETE * FROM GrbImpressionDetailTemps")

  Set rstImpTemps = New ADODB.Recordset

  Call rstImpTemps.Open("SELECT * FROM GrbImpressionDetailTemps", g_connData, adOpenDynamic, adLockOptimistic)

  Do While Not rstEmploye.EOF
  Call rstImpTemps.AddNew

  rstImpTemps.Fields("Employe") = rstEmploye.Fields("Employe")

If Not IsNull(rstEmploye.Fields("Type")) Then
'Retirer pour afficher que le type vu GLL v41 2017-08-23
 'Select Case rstEmploye.Fields("Type")
 ' Case "Dessin": rstImpTemps.Fields("Type") = "Dessin"
 ' Case "Fabrication": rstImpTemps.Fields("Type") = "Fabrication"
 ' Case "Assemblage": rstImpTemps.Fields("Type") = "Assemblage"
 ' Case "ProgInterface": rstImpTemps.Fields("Type") = "Programmation d'interface"
 ' Case "ProgAutomate": rstImpTemps.Fields("Type") = "Programmation d'automate"
 ' Case "ProgRobot": rstImpTemps.Fields("Type") = "Programmation de robot"
 ' Case "Vision": rstImpTemps.Fields("Type") = "Vision"
 ' Case "Test": rstImpTemps.Fields("Type") = "Test"
 ' Case "Installation": rstImpTemps.Fields("Type") = "Installation"
 ' Case "MiseService": rstImpTemps.Fields("Type") = "Mise en service"
 ' Case "Formation": rstImpTemps.Fields("Type") = "Formation du personnel"
 ' Case "Gestion": rstImpTemps.Fields("Type") = "Gestion du projet"
 ' Case "Shipping": rstImpTemps.Fields("Type") = "Expédition"
 ' Case "Prototypage-Dévelloppement expérimental": rstImpTemps.Fields("Type") = "Prototypage-Dévelloppement expérimental"
 'End Select
 rstImpTemps.Fields("Type") = rstEmploye.Fields("Type")
1   Else
 rstImpTemps.Fields("Type") = ""
 End If

 rstImpTemps.Fields("TotalHeures") = rstEmploye.Fields("TotalHeures")

1  Call rstImpTemps.Update

 Call rstEmploye.MoveNext
 Loop

Call rstEmploye.Close
Set rstEmploye = Nothing

Set DR_TempsElec.DataSource = rstImpTemps

Set rstProjSoum = New ADODB.Recordset

If m_bProjet = True Then
 Call rstProjSoum.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & m_sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
Else
 Call rstProjSoum.Open("SELECT * FROM GrbSoumissionElec WHERE IDSoumission = '" & m_sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
End If

 'Affichage du # de projet ou soumissin
DR_TempsElec.Sections("Section4").Controls("lblNoProjet").Caption = m_sNoProjet

2  If Not IsNull(rstProjSoum.Fields("TempsDessin")) Then
 DR_TempsElec.Sections("Section4").Controls("lblTempsDessinEstime").Caption = Round(rstProjSoum.Fields("TempsDessin"), 2)
2  Else
 DR_TempsElec.Sections("Section4").Controls("lblTempsDessinEstime").Caption = "0"
2  End If

If Not IsNull(rstProjSoum.Fields("TempsFabrication")) Then
DR_TempsElec.Sections("Section4").Controls("lblTempsFabricationEstime").Caption = Round(rstProjSoum.Fields("TempsFabrication"), 2)
Else
DR_TempsElec.Sections("Section4").Controls("lblTempsFabricationEstime").Caption = "0"
End If

If Not IsNull(rstProjSoum.Fields("TempsAssemblage")) Then
 DR_TempsElec.Sections("Section4").Controls("lblTempsAssemblageEstime").Caption = Round(rstProjSoum.Fields("TempsAssemblage"), 2)
Else
 DR_TempsElec.Sections("Section4").Controls("lblTempsAssemblageEstime").Caption = "0"
End If

If Not IsNull(rstProjSoum.Fields("TempsProgInterface")) Then
 DR_TempsElec.Sections("Section4").Controls("lblTempsProgInterfaceEstime").Caption = Round(rstProjSoum.Fields("TempsProgInterface"), 2)
Else
 DR_TempsElec.Sections("Section4").Controls("lblTempsProgInterfaceEstime").Caption = "0"
End If

3  If Not IsNull(rstProjSoum.Fields("TempsProgAutomate")) Then
 DR_TempsElec.Sections("Section4").Controls("lblTempsProgAutomateEstime").Caption = Round(rstProjSoum.Fields("TempsProgAutomate"), 2)
3  Else
 DR_TempsElec.Sections("Section4").Controls("lblTempsProgAutomateEstime").Caption = "0"
3  End If

If Not IsNull(rstProjSoum.Fields("TempsProgRobot")) Then
 DR_TempsElec.Sections("Section4").Controls("lblTempsProgRobotEstime").Caption = Round(rstProjSoum.Fields("TempsProgRobot"), 2)
 Else
DR_TempsElec.Sections("Section4").Controls("lblTempsProgRobotEstime").Caption = "0"
End If

4 If Not IsNull(rstProjSoum.Fields("TempsVision")) Then
4 DR_TempsElec.Sections("Section4").Controls("lblTempsVisionEstime").Caption = Round(rstProjSoum.Fields("TempsVision"), 2)
4 Else
4 DR_TempsElec.Sections("Section4").Controls("lblTempsVisionEstime").Caption = "0"
4 End If

4 If Not IsNull(rstProjSoum.Fields("TempsTest")) Then
4 DR_TempsElec.Sections("Section4").Controls("lblTempsTestEstime").Caption = Round(rstProjSoum.Fields("TempsTest"), 2)
4 Else
4 DR_TempsElec.Sections("Section4").Controls("lblTempsTestEstime").Caption = "0"
4 End If

4  If Not IsNull(rstProjSoum.Fields("TempsInstallation")) Then
4  DR_TempsElec.Sections("Section4").Controls("lblTempsInstallationEstime").Caption = Round(rstProjSoum.Fields("TempsInstallation"), 2)
4  Else
4  DR_TempsElec.Sections("Section4").Controls("lblTempsInstallationEstime").Caption = "0"
4  End If

4  If Not IsNull(rstProjSoum.Fields("TempsMiseService")) Then
4  DR_TempsElec.Sections("Section4").Controls("lblTempsMiseServiceEstime").Caption = Round(rstProjSoum.Fields("TempsMiseService"), 2)
4  Else
50 DR_TempsElec.Sections("Section4").Controls("lblTempsMiseServiceEstime").Caption = "0"
50 End If

 If Not IsNull(rstProjSoum.Fields("TempsFormation")) Then
 DR_TempsElec.Sections("Section4").Controls("lblTempsFormationEstime").Caption = Round(rstProjSoum.Fields("TempsFormation"), 2)
 Else
 DR_TempsElec.Sections("Section4").Controls("lblTempsFormationEstime").Caption = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsGestion")) Then
 DR_TempsElec.Sections("Section4").Controls("lblTempsGestionEstime").Caption = Round(rstProjSoum.Fields("TempsGestion"), 2)
 Else
 DR_TempsElec.Sections("Section4").Controls("lblTempsGestionEstime").Caption = "0"
 End If

5  If Not IsNull(rstProjSoum.Fields("TempsShipping")) Then
5  DR_TempsElec.Sections("Section4").Controls("lblTempsShippingEstime").Caption = Round(rstProjSoum.Fields("TempsShipping"), 2)
5  Else
5  DR_TempsElec.Sections("Section4").Controls("lblTempsShippingEstime").Caption = "0"
5  End If




5  dblTotal = CDbl(DR_TempsElec.Sections("Section4").Controls("lblTempsDessinEstime").Caption) + _
 CDbl(DR_TempsElec.Sections("Section4").Controls("lblTempsFabricationEstime").Caption) + _
 CDbl(DR_TempsElec.Sections("Section4").Controls("lblTempsAssemblageEstime").Caption) + _
 CDbl(DR_TempsElec.Sections("Section4").Controls("lblTempsProgInterfaceEstime").Caption) + _
 CDbl(DR_TempsElec.Sections("Section4").Controls("lblTempsProgAutomateEstime").Caption) + _
 CDbl(DR_TempsElec.Sections("Section4").Controls("lblTempsProgRobotEstime").Caption) + _
 CDbl(DR_TempsElec.Sections("Section4").Controls("lblTempsVisionEstime").Caption) + _
 CDbl(DR_TempsElec.Sections("Section4").Controls("lblTempsTestEstime").Caption) + _
 CDbl(DR_TempsElec.Sections("Section4").Controls("lblTempsInstallationEstime").Caption) + _
 CDbl(DR_TempsElec.Sections("Section4").Controls("lblTempsMiseServiceEstime").Caption) + _
 CDbl(DR_TempsElec.Sections("Section4").Controls("lblTempsFormationEstime").Caption) + _
 CDbl(DR_TempsElec.Sections("Section4").Controls("lblTempsGestionEstime").Caption) + _
 CDbl(DR_TempsElec.Sections("Section4").Controls("lblTempsShippingEstime").Caption)

5  DR_TempsElec.Sections("Section4").Controls("lblTotalTempsEstime").Caption = dblTotal

5  Call rstProjSoum.Close
60 Set rstProjSoum = Nothing

60 Call CalculerTempsReelsElec

  Call DR_TempsElec.Show(vbModal)

  Call rstImpTemps.Close
  Set rstImpTemps = Nothing

  Exit Sub

Oups:

  wOups "frmDetailTemps", "ImprimerDetailTempsElectriques", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerDetailTempsMecaniques()
 
 On Error GoTo Oups

 Dim rstEmploye As ADODB.Recordset
 Dim rstImpTemps As ADODB.Recordset
 Dim rstProjSoum As ADODB.Recordset
 Dim rstSoum As ADODB.Recordset
 Dim dblTotal As Double
 Dim sFilterNoProjet As String

 If Right$(m_sNoProjet, 2) = "99" Then
 sFilterNoProjet = "LEFT(NoProjet, 6) = '" & Left$(m_sNoProjet, 6) & "'"
 Else
 sFilterNoProjet = "NoProjet = '" & m_sNoProjet & "'"
  End If

  Set rstEmploye = New ADODB.Recordset

  Call rstEmploye.Open("SELECT Employe, Type, (Sum(TimeSerial(Left(HeureFin,2), RIGHT(HeureFin,2),0) - TimeSerial(Left(HeureDébut,2), RIGHT(HeureDébut,2),0)) * 24) AS TotalHeures FROM GrbPunch INNER JOIN GrbEmployés ON GrbPunch.NoEmploye = GrbEmployés.NoEmploye WHERE HeureDébut is Not Null And HeureFin is Not Null AND " & sFilterNoProjet & " GROUP BY Employe, Type", g_connData, adOpenDynamic, adLockOptimistic)

  Call g_connData.Execute("DELETE * FROM GrbImpressionDetailTemps")

  Set rstImpTemps = New ADODB.Recordset

  Call rstImpTemps.Open("SELECT * FROM GrbImpressionDetailTemps", g_connData, adOpenDynamic, adLockOptimistic)

  Do While Not rstEmploye.EOF
  Call rstImpTemps.AddNew

rstImpTemps.Fields("Employe") = rstEmploye.Fields("Employe")

1 If Not IsNull(rstEmploye.Fields("Type")) Then
 'Retirer par GLL V41 2017-08-23
 'Select Case rstEmploye.Fields("Type")
 ' Case "Dessin": rstImpTemps.Fields("Type") = "Conception et dessins"
 ' Case "Coupe": rstImpTemps.Fields("Type") = "Coupe et préparation (sauf soudage)"
 ' Case "Machinage": rstImpTemps.Fields("Type") = "Machinage"
 ' Case "Soudure": rstImpTemps.Fields("Type") = "Coupe, soudure et meulage"
 ' Case "Assemblage": rstImpTemps.Fields("Type") = "Assemblage des systèmes"
 ' Case "Peinture": rstImpTemps.Fields("Type") = "Peinture et finition"
 ' Case "Test": rstImpTemps.Fields("Type") = "Tests finaux"
 ' Case "Installation": rstImpTemps.Fields("Type") = "Installation"
 ' Case "Formation": rstImpTemps.Fields("Type") = "Formation du personnel"
 ' Case "Gestion": rstImpTemps.Fields("Type") = "Gestion du projet"
 ' Case "Shipping": rstImpTemps.Fields("Type") = "Expédition"
 'End Select
 rstImpTemps.Fields("Type") = rstEmploye.Fields("Type")
1  Else
 rstImpTemps.Fields("Type") = ""
 End If

 rstImpTemps.Fields("TotalHeures") = rstEmploye.Fields("TotalHeures")

 Call rstImpTemps.Update

1  Call rstEmploye.MoveNext
 Loop

 Call rstEmploye.Close
Set rstEmploye = Nothing

Set DR_TempsMec.DataSource = rstImpTemps

Set rstProjSoum = New ADODB.Recordset

If m_bProjet = True Then
 Call rstProjSoum.Open("SELECT * FROM GrbProjetMec WHERE IDProjet = '" & m_sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
Else
 Call rstProjSoum.Open("SELECT * FROM GrbSoumissionMec WHERE IDSoumission = '" & m_sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
End If

 'Affichage du # de projet ou soumission
DR_TempsMec.Sections("Section4").Controls("lblNoProjet").Caption = m_sNoProjet

 'Si soumission
If m_bProjet = False Then
If Not IsNull(rstProjSoum.Fields("TempsDessin")) Then
 DR_TempsMec.Sections("Section4").Controls("lblTempsDessinSoum").Caption = Round(rstProjSoum.Fields("TempsDessin"), 2)
Else
 DR_TempsMec.Sections("Section4").Controls("lblTempsDessinSoum").Caption = "0"
End If

 If Not IsNull(rstProjSoum.Fields("TempsCoupe")) Then
 DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeSoum").Caption = Round(rstProjSoum.Fields("TempsCoupe"), 2)
 Else
 DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeSoum").Caption = "0"
3 End If

 If Not IsNull(rstProjSoum.Fields("TempsMachinage")) Then
 DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageSoum").Caption = Round(rstProjSoum.Fields("TempsMachinage"), 2)
 Else
 DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageSoum").Caption = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsSoudure")) Then
 DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureSoum").Caption = Round(rstProjSoum.Fields("TempsSoudure"), 2)
 Else
 DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureSoum").Caption = "0"
 End If

If Not IsNull(rstProjSoum.Fields("TempsAssemblage")) Then
 DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageSoum").Caption = Round(rstProjSoum.Fields("TempsAssemblage"), 2)
Else
 DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageSoum").Caption = "0"
End If

 If Not IsNull(rstProjSoum.Fields("TempsPeinture")) Then
 DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureSoum").Caption = Round(rstProjSoum.Fields("TempsPeinture"), 2)
 Else
 DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureSoum").Caption = "0"
4 End If

4 If Not IsNull(rstProjSoum.Fields("TempsTest")) Then
4 DR_TempsMec.Sections("Section4").Controls("lblTempsTestSoum").Caption = Round(rstProjSoum.Fields("TempsTest"), 2)
4 Else
4 DR_TempsMec.Sections("Section4").Controls("lblTempsTestSoum").Caption = "0"
4 End If

4 If Not IsNull(rstProjSoum.Fields("TempsInstallation")) Then
4 DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationSoum").Caption = Round(rstProjSoum.Fields("TempsInstallation"), 2)
4 Else
4 DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationSoum").Caption = "0"
4 End If

4  If Not IsNull(rstProjSoum.Fields("TempsFormation")) Then
4  DR_TempsMec.Sections("Section4").Controls("lblTempsFormationSoum").Caption = Round(rstProjSoum.Fields("TempsFormation"), 2)
4  Else
4  DR_TempsMec.Sections("Section4").Controls("lblTempsFormationSoum").Caption = "0"
4  End If

4  If Not IsNull(rstProjSoum.Fields("TempsGestion")) Then
4  DR_TempsMec.Sections("Section4").Controls("lblTempsGestionSoum").Caption = Round(rstProjSoum.Fields("TempsGestion"), 2)
4  Else
50 DR_TempsMec.Sections("Section4").Controls("lblTempsGestionSoum").Caption = "0"
5 End If

 If Not IsNull(rstProjSoum.Fields("TempsShipping")) Then
 DR_TempsMec.Sections("Section4").Controls("lblTempsShippingSoum").Caption = Round(rstProjSoum.Fields("TempsShipping"), 2)
 Else
 DR_TempsMec.Sections("Section4").Controls("lblTempsShippingSoum").Caption = "0"
 End If

 DR_TempsMec.Sections("Section4").Controls("lblTempsDessinProj").Caption = "---"
 DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeProj").Caption = "---"
 DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageProj").Caption = "---"
 DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureProj").Caption = "---"
 DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageProj").Caption = "---"
5  DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureProj").Caption = "---"
5  DR_TempsMec.Sections("Section4").Controls("lblTempsTestProj").Caption = "---"
5  DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationProj").Caption = "---"
5  DR_TempsMec.Sections("Section4").Controls("lblTempsFormationProj").Caption = "---"
5  DR_TempsMec.Sections("Section4").Controls("lblTempsGestionProj").Caption = "---"
5  DR_TempsMec.Sections("Section4").Controls("lblTempsShippingProj").Caption = "---"

5  DR_TempsMec.Sections("Section4").Controls("lblTempsDessinConc").Caption = "---"
5  DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeConc").Caption = "---"
60 DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageConc").Caption = "---"
  DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureConc").Caption = "---"
  DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageConc").Caption = "---"
  DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureConc").Caption = "---"
  DR_TempsMec.Sections("Section4").Controls("lblTempsTestConc").Caption = "---"
  DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationConc").Caption = "---"
  DR_TempsMec.Sections("Section4").Controls("lblTempsFormationConc").Caption = "---"
  DR_TempsMec.Sections("Section4").Controls("lblTempsGestionConc").Caption = "---"
  DR_TempsMec.Sections("Section4").Controls("lblTempsShippingConc").Caption = "---"
  Else
  If Not IsNull(rstProjSoum.Fields("IDSoumission")) Then
  Set rstSoum = New ADODB.Recordset

6  Call rstSoum.Open("SELECT * FROM GrbSoumissionMec WHERE IDSoumission = '" & m_sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

6  If Not rstSoum.EOF Then
6  If Not IsNull(rstSoum.Fields("TempsDessin")) Then
6  DR_TempsMec.Sections("Section4").Controls("lblTempsDessinSoum").Caption = Round(rstSoum.Fields("TempsDessin"), 2)
6  Else
6  DR_TempsMec.Sections("Section4").Controls("lblTempsDessinSoum").Caption = "0"
6  End If

6  If Not IsNull(rstSoum.Fields("TempsCoupe")) Then
70 DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeSoum").Caption = Round(rstSoum.Fields("TempsCoupe"), 2)
  Else
  DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeSoum").Caption = "0"
  End If

  If Not IsNull(rstSoum.Fields("TempsMachinage")) Then
  DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageSoum").Caption = Round(rstSoum.Fields("TempsMachinage"), 2)
  Else
  DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageSoum").Caption = "0"
  End If

  If Not IsNull(rstSoum.Fields("TempsSoudure")) Then
  DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureSoum").Caption = Round(rstSoum.Fields("TempsSoudure"), 2)
  Else
   DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureSoum").Caption = "0"
   End If

7  If Not IsNull(rstSoum.Fields("TempsAssemblage")) Then
7  DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageSoum").Caption = Round(rstSoum.Fields("TempsAssemblage"), 2)
7  Else
7  DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageSoum").Caption = "0"
7  End If

7  If Not IsNull(rstSoum.Fields("TempsPeinture")) Then
80 DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureSoum").Caption = Round(rstSoum.Fields("TempsPeinture"), 2)
  Else
  DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureSoum").Caption = "0"
  End If

  If Not IsNull(rstSoum.Fields("TempsTest")) Then
  DR_TempsMec.Sections("Section4").Controls("lblTempsTestSoum").Caption = Round(rstSoum.Fields("TempsTest"), 2)
  Else
  DR_TempsMec.Sections("Section4").Controls("lblTempsTestSoum").Caption = "0"
  End If

  If Not IsNull(rstSoum.Fields("TempsInstallation")) Then
  DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationSoum").Caption = Round(rstSoum.Fields("TempsInstallation"), 2)
  Else
   DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationSoum").Caption = "0"
   End If

   If Not IsNull(rstSoum.Fields("TempsFormation")) Then
   DR_TempsMec.Sections("Section4").Controls("lblTempsFormationSoum").Caption = Round(rstSoum.Fields("TempsFormation"), 2)
8  Else
8  DR_TempsMec.Sections("Section4").Controls("lblTempsFormationSoum").Caption = "0"
8  End If

8  If Not IsNull(rstSoum.Fields("TempsGestion")) Then
90 DR_TempsMec.Sections("Section4").Controls("lblTempsGestionSoum").Caption = Round(rstSoum.Fields("TempsGestion"), 2)
  Else
  DR_TempsMec.Sections("Section4").Controls("lblTempsGestionSoum").Caption = "0"
  End If

  If Not IsNull(rstSoum.Fields("TempsShipping")) Then
  DR_TempsMec.Sections("Section4").Controls("lblTempsShippingSoum").Caption = Round(rstSoum.Fields("TempsShipping"), 2)
  Else
  DR_TempsMec.Sections("Section4").Controls("lblTempsShippingSoum").Caption = "0"
  End If
  Else
  DR_TempsMec.Sections("Section4").Controls("lblTempsDessinSoum").Caption = "---"
  DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeSoum").Caption = "---"
 DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageSoum").Caption = "---"
   DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureSoum").Caption = "---"
 DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageSoum").Caption = "---"
   DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureSoum").Caption = "---"
 DR_TempsMec.Sections("Section4").Controls("lblTempsTestSoum").Caption = "---"
   DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationSoum").Caption = "---"
 DR_TempsMec.Sections("Section4").Controls("lblTempsFormationSoum").Caption = "---"
9  DR_TempsMec.Sections("Section4").Controls("lblTempsGestionSoum").Caption = "---"
 DR_TempsMec.Sections("Section4").Controls("lblTempsShippingSoum").Caption = "---"
10 End If

1 Call rstSoum.Close
1 Set rstSoum = Nothing
 Else
1 DR_TempsMec.Sections("Section4").Controls("lblTempsDessinSoum").Caption = "---"
 DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeSoum").Caption = "---"
1 DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageSoum").Caption = "---"
 DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureSoum").Caption = "---"
1 DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageSoum").Caption = "---"
 DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureSoum").Caption = "---"
1 DR_TempsMec.Sections("Section4").Controls("lblTempsTestSoum").Caption = "---"
10  DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationSoum").Caption = "---"
10  DR_TempsMec.Sections("Section4").Controls("lblTempsFormationSoum").Caption = "---"
10  DR_TempsMec.Sections("Section4").Controls("lblTempsGestionSoum").Caption = "---"
10  DR_TempsMec.Sections("Section4").Controls("lblTempsShippingSoum").Caption = "---"
10  End If

10  If Not IsNull(rstProjSoum.Fields("TempsDessinProj")) Then
10  DR_TempsMec.Sections("Section4").Controls("lblTempsDessinProj").Caption = Round(rstProjSoum.Fields("TempsDessinProj"), 2)
10  Else
1 DR_TempsMec.Sections("Section4").Controls("lblTempsDessinProj").Caption = "0"
11End If

1 If Not IsNull(rstProjSoum.Fields("TempsCoupeProj")) Then
1 DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeProj").Caption = Round(rstProjSoum.Fields("TempsCoupeProj"), 2)
1 Else
1 DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeProj").Caption = "0"
1 End If

1 If Not IsNull(rstProjSoum.Fields("TempsMachinageProj")) Then
1 DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageProj").Caption = Round(rstProjSoum.Fields("TempsMachinageProj"), 2)
1 Else
1 DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageProj").Caption = "0"
1 End If

11  If Not IsNull(rstProjSoum.Fields("TempsSoudureProj")) Then
1 DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureProj").Caption = Round(rstProjSoum.Fields("TempsSoudureProj"), 2)
 Else
1 DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureProj").Caption = "0"
 End If

1 If Not IsNull(rstProjSoum.Fields("TempsAssemblageProj")) Then
 DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageProj").Caption = Round(rstProjSoum.Fields("TempsAssemblageProj"), 2)
11  Else
 DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageProj").Caption = "0"
1 End If

1 If Not IsNull(rstProjSoum.Fields("TempsPeintureProj")) Then
1 DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureProj").Caption = Round(rstProjSoum.Fields("TempsPeintureProj"), 2)
1 Else
1 DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureProj").Caption = "0"
1 End If

1 If Not IsNull(rstProjSoum.Fields("TempsTestProj")) Then
1 DR_TempsMec.Sections("Section4").Controls("lblTempsTestProj").Caption = Round(rstProjSoum.Fields("TempsTestProj"), 2)
1 Else
1 DR_TempsMec.Sections("Section4").Controls("lblTempsTestProj").Caption = "0"
1 End If

12  If Not IsNull(rstProjSoum.Fields("TempsInstallationProj")) Then
1 DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationProj").Caption = Round(rstProjSoum.Fields("TempsInstallationProj"), 2)
12  Else
1 DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationProj").Caption = "0"
12  End If

1 If Not IsNull(rstProjSoum.Fields("TempsFormationProj")) Then
1 DR_TempsMec.Sections("Section4").Controls("lblTempsFormationProj").Caption = Round(rstProjSoum.Fields("TempsFormationProj"), 2)
1 Else
1 DR_TempsMec.Sections("Section4").Controls("lblTempsFormationProj").Caption = "0"
13End If

1 If Not IsNull(rstProjSoum.Fields("TempsGestionProj")) Then
1 DR_TempsMec.Sections("Section4").Controls("lblTempsGestionProj").Caption = Round(rstProjSoum.Fields("TempsGestionProj"), 2)
1 Else
1 DR_TempsMec.Sections("Section4").Controls("lblTempsGestionProj").Caption = "0"
1 End If

1 If Not IsNull(rstProjSoum.Fields("TempsShippingProj")) Then
1 DR_TempsMec.Sections("Section4").Controls("lblTempsShippingProj").Caption = Round(rstProjSoum.Fields("TempsShippingProj"), 2)
1 Else
1 DR_TempsMec.Sections("Section4").Controls("lblTempsShippingProj").Caption = "0"
1 End If

13  If rstProjSoum.Fields("TempsProjBarré") = True Then
1 If Not IsNull(rstProjSoum.Fields("TempsDessinConc")) Then
1 DR_TempsMec.Sections("Section4").Controls("lblTempsDessinConc").Caption = Round(rstProjSoum.Fields("TempsDessinConc"), 2)
1 Else
1 DR_TempsMec.Sections("Section4").Controls("lblTempsDessinConc").Caption = "0"
1 End If

1 If Not IsNull(rstProjSoum.Fields("TempsCoupeConc")) Then
1 DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeConc").Caption = Round(rstProjSoum.Fields("TempsCoupeConc"), 2)
1 Else
14 DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeConc").Caption = "0"
14 End If

14 If Not IsNull(rstProjSoum.Fields("TempsMachinageConc")) Then
14 DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageConc").Caption = Round(rstProjSoum.Fields("TempsMachinageConc"), 2)
14 Else
14 DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageConc").Caption = "0"
14 End If

14 If Not IsNull(rstProjSoum.Fields("TempsSoudureConc")) Then
14 DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureConc").Caption = Round(rstProjSoum.Fields("TempsSoudureConc"), 2)
14 Else
14 DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureConc").Caption = "0"
14  End If

14  If Not IsNull(rstProjSoum.Fields("TempsAssemblageConc")) Then
14  DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageConc").Caption = Round(rstProjSoum.Fields("TempsAssemblageConc"), 2)
14  Else
14  DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageConc").Caption = "0"
14  End If

14  If Not IsNull(rstProjSoum.Fields("TempsPeintureConc")) Then
14  DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureConc").Caption = Round(rstProjSoum.Fields("TempsPeintureConc"), 2)
150 Else
1DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureConc").Caption = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsTestConc")) Then
 DR_TempsMec.Sections("Section4").Controls("lblTempsTestConc").Caption = Round(rstProjSoum.Fields("TempsTestConc"), 2)
 Else
 DR_TempsMec.Sections("Section4").Controls("lblTempsTestConc").Caption = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsInstallationConc")) Then
 DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationConc").Caption = Round(rstProjSoum.Fields("TempsInstallationConc"), 2)
 Else
 DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationConc").Caption = "0"
15  End If

15  If Not IsNull(rstProjSoum.Fields("TempsFormationConc")) Then
15  DR_TempsMec.Sections("Section4").Controls("lblTempsFormationConc").Caption = Round(rstProjSoum.Fields("TempsFormationConc"), 2)
15  Else
15  DR_TempsMec.Sections("Section4").Controls("lblTempsFormationConc").Caption = "0"
15  End If

15  If Not IsNull(rstProjSoum.Fields("TempsGestionConc")) Then
15  DR_TempsMec.Sections("Section4").Controls("lblTempsGestionConc").Caption = Round(rstProjSoum.Fields("TempsGestionConc"), 2)
160 Else
 DR_TempsMec.Sections("Section4").Controls("lblTempsGestionConc").Caption = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsShippingConc")) Then
 DR_TempsMec.Sections("Section4").Controls("lblTempsShippingConc").Caption = Round(rstProjSoum.Fields("TempsShippingConc"), 2)
 Else
 DR_TempsMec.Sections("Section4").Controls("lblTempsShippingConc").Caption = "0"
 End If
 Else
 DR_TempsMec.Sections("Section4").Controls("lblTempsDessinConc").Caption = "---"
 DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeConc").Caption = "---"
 DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageConc").Caption = "---"
16  DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureConc").Caption = "---"
16  DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageConc").Caption = "---"
16  DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureConc").Caption = "---"
16  DR_TempsMec.Sections("Section4").Controls("lblTempsTestConc").Caption = "---"
16  DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationConc").Caption = "---"
16  DR_TempsMec.Sections("Section4").Controls("lblTempsFormationConc").Caption = "---"
16  DR_TempsMec.Sections("Section4").Controls("lblTempsGestionConc").Caption = "---"
16  DR_TempsMec.Sections("Section4").Controls("lblTempsShippingConc").Caption = "---"
170 End If
End If

1  If DR_TempsMec.Sections("Section4").Controls("lblTempsDessinSoum").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeSoum").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageSoum").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureSoum").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageSoum").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureSoum").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsTestSoum").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationSoum").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsFormationSoum").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsGestionSoum").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsShippingSoum").Caption = "---" Then
 
 DR_TempsMec.Sections("Section4").Controls("lblTotalTempsSoum").Caption = "---"
1  Else
 dblTotal = CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsDessinSoum").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeSoum").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageSoum").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureSoum").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageSoum").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureSoum").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsTestSoum").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationSoum").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsFormationSoum").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsGestionSoum").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsShippingSoum").Caption)
 
 DR_TempsMec.Sections("Section4").Controls("lblTotalTempsSoum").Caption = dblTotal
1  End If


1  If DR_TempsMec.Sections("Section4").Controls("lblTempsDessinProj").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeProj").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageProj").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureProj").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageProj").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureProj").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsTestProj").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationProj").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsFormationProj").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsGestionProj").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsGestionProj").Caption = "---" Then
 
 DR_TempsMec.Sections("Section4").Controls("lblTotalTempsProj").Caption = "---"
1  Else
 dblTotal = CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsDessinProj").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeProj").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageProj").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureProj").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageProj").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureProj").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsTestProj").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationProj").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsFormationProj").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsGestionProj").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsShippingProj").Caption)

1   DR_TempsMec.Sections("Section4").Controls("lblTotalTempsProj").Caption = dblTotal
1   End If

17  If DR_TempsMec.Sections("Section4").Controls("lblTempsDessinConc").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeConc").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageConc").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureConc").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageConc").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureConc").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsTestConc").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationConc").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsFormationConc").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsGestionConc").Caption = "---" And _
 DR_TempsMec.Sections("Section4").Controls("lblTempsShippingConc").Caption = "---" Then
 
17  DR_TempsMec.Sections("Section4").Controls("lblTotalTempsConc").Caption = "---"
178Else
17  dblTotal = CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsDessinConc").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeConc").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageConc").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureConc").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageConc").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureConc").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsTestConc").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationConc").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsFormationConc").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsGestionConc").Caption) + _
 CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsShippingConc").Caption)
 
17  DR_TempsMec.Sections("Section4").Controls("lblTotalTempsConc").Caption = dblTotal
17  End If

180 Call rstProjSoum.Close
Set rstProjSoum = Nothing

1  Call CalculerTempsReelsMec

1  Call DR_TempsMec.Show(vbModal)

1  Call rstImpTemps.Close
1  Set rstImpTemps = Nothing

1  Exit Sub

Oups:

1  wOups "frmDetailTemps", "ImprimerDetailTempsMecaniques", Err, Err.number, Err.Description
End Sub

Private Sub CalculerTempsReelsElec()

 On Error GoTo Oups

 Dim rstTotal As ADODB.Recordset
 Dim sDateDebut As String
 Dim sDateFin As String
 Dim sTotal As String
 Dim sFilterNoProjet As String

 If Right$(m_sNoProjet, 2) = "99" Then
 sFilterNoProjet = "LEFT(NoProjet, 6) = '" & Left$(m_sNoProjet, 6) & "'"
 Else
 sFilterNoProjet = "NoProjet = '" & m_sNoProjet & "'"
 End If

  Set rstTotal = New ADODB.Recordset
 
  sDateDebut = "TIMESERIAL(Left(GrbPunch.HeureDébut,2),RIGHT(GrbPunch.HeureDébut,2),0)"
 
  sDateFin = "TIMESERIAL(Left(GrbPunch.HeureFin,2),RIGHT(GrbPunch.HeureFin,2),0)"
 
  sTotal = "(SUM(" & sDateFin & " - " & sDateDebut & ")* 24) As Total"

  Call rstTotal.Open("SELECT Type, " & sTotal & " FROM GrbPunch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null GROUP BY Type", g_connData, adOpenDynamic, adLockOptimistic)

  DR_TempsElec.Sections("Section4").Controls("lblTempsDessinReel").Caption = "0"
  DR_TempsElec.Sections("Section4").Controls("lblTempsFabricationReel").Caption = "0"
  DR_TempsElec.Sections("Section4").Controls("lblTempsAssemblageReel").Caption = "0"
10 DR_TempsElec.Sections("Section4").Controls("lblTempsProgInterfaceReel").Caption = "0"
DR_TempsElec.Sections("Section4").Controls("lblTempsProgAutomateReel").Caption = "0"
DR_TempsElec.Sections("Section4").Controls("lblTempsProgRobotReel").Caption = "0"
DR_TempsElec.Sections("Section4").Controls("lblTempsVisionReel").Caption = "0"
DR_TempsElec.Sections("Section4").Controls("lblTempsTestReel").Caption = "0"
DR_TempsElec.Sections("Section4").Controls("lblTempsInstallationReel").Caption = "0"
DR_TempsElec.Sections("Section4").Controls("lblTempsMiseServiceReel").Caption = "0"
DR_TempsElec.Sections("Section4").Controls("lblTempsFormationReel").Caption = "0"
DR_TempsElec.Sections("Section4").Controls("lblTempsGestionReel").Caption = "0"
DR_TempsElec.Sections("Section4").Controls("lblTempsShippingReel").Caption = "0"

Do While Not rstTotal.EOF
 If Not IsNull(rstTotal.Fields("Total")) Then
 Select Case rstTotal.Fields("Type")
 Case "Dessin": DR_TempsElec.Sections("Section4").Controls("lblTempsDessinReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "Fabrication": DR_TempsElec.Sections("Section4").Controls("lblTempsFabricationReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "Assemblage": DR_TempsElec.Sections("Section4").Controls("lblTempsAssemblageReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "ProgInterface": DR_TempsElec.Sections("Section4").Controls("lblTempsProgInterfaceReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "ProgAutomate": DR_TempsElec.Sections("Section4").Controls("lblTempsProgAutomateReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "ProgRobot": DR_TempsElec.Sections("Section4").Controls("lblTempsProgRobotReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "Vision": DR_TempsElec.Sections("Section4").Controls("lblTempsVisionReel").Caption = Round(rstTotal.Fields("Total"), 2)
1  Case "Test": DR_TempsElec.Sections("Section4").Controls("lblTempsTestReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "Installation": DR_TempsElec.Sections("Section4").Controls("lblTempsInstallationReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "MiseService": DR_TempsElec.Sections("Section4").Controls("lblTempsMiseServiceReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "Formation": DR_TempsElec.Sections("Section4").Controls("lblTempsFormationReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "Gestion": DR_TempsElec.Sections("Section4").Controls("lblTempsGestionReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "Shipping": DR_TempsElec.Sections("Section4").Controls("lblTempsShippingReel").Caption = Round(rstTotal.Fields("Total"), 2)
2 Case "Prototypage-Dévelloppement expérimental": DR_TempsElec.Sections("Section4").Controls("lblTempsprototypeReel").Caption = Round(rstTotal.Fields("Total"), 2)

 End Select
 End If

 Call rstTotal.MoveNext
Loop

Call rstTotal.Close
 
Call rstTotal.Open("SELECT " & sTotal & " FROM GrbPunch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null", g_connData, adOpenDynamic, adLockOptimistic)

If Not IsNull(rstTotal.Fields("Total")) Then
DR_TempsElec.Sections("Section4").Controls("lblTotalTempsReel").Caption = Round(rstTotal.Fields("Total"), 2)
Else
DR_TempsElec.Sections("Section4").Controls("lblTotalTempsReel").Caption = "0"
End If

2  Call rstTotal.Close
Set rstTotal = Nothing

2  Exit Sub

Oups:

wOups "frmDetailTemps", "CalculerTempsReelsElec", Err, Err.number, Err.Description
End Sub

Private Sub CalculerTempsReelsMec()

 On Error GoTo Oups

 Dim rstTotal As ADODB.Recordset
 Dim sDateDebut As String
 Dim sDateFin As String
 Dim sTotal As String
 Dim sFilterNoProjet As String

 If Right$(m_sNoProjet, 2) = "99" Then
 sFilterNoProjet = "LEFT(NoProjet, 6) = '" & Left$(m_sNoProjet, 6) & "'"
 Else
 sFilterNoProjet = "NoProjet = '" & m_sNoProjet & "'"
 End If

  Set rstTotal = New ADODB.Recordset
 
  sDateDebut = "TIMESERIAL(Left(GrbPunch.HeureDébut,2),RIGHT(GrbPunch.HeureDébut,2),0)"
 
  sDateFin = "TIMESERIAL(Left(GrbPunch.HeureFin,2),RIGHT(GrbPunch.HeureFin,2),0)"
 
  sTotal = "(SUM(" & sDateFin & " - " & sDateDebut & ")* 24) As Total"

  Call rstTotal.Open("SELECT Type, " & sTotal & " FROM GrbPunch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null GROUP BY Type", g_connData, adOpenDynamic, adLockOptimistic)

  DR_TempsMec.Sections("Section4").Controls("lblTempsDessinReel").Caption = "0"
  DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeReel").Caption = "0"
  DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageReel").Caption = "0"
10 DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureReel").Caption = "0"
DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageReel").Caption = "0"
DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureReel").Caption = "0"
DR_TempsMec.Sections("Section4").Controls("lblTempsTestReel").Caption = "0"
DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationReel").Caption = "0"
DR_TempsMec.Sections("Section4").Controls("lblTempsFormationReel").Caption = "0"
DR_TempsMec.Sections("Section4").Controls("lblTempsGestionReel").Caption = "0"
DR_TempsMec.Sections("Section4").Controls("lblTempsShippingReel").Caption = "0"
 DR_TempsMec.Sections("Section4").Controls("lblTempsprototypeReel").Caption = "0"


Do While Not rstTotal.EOF
 If Not IsNull(rstTotal.Fields("Total")) Then
 Select Case rstTotal.Fields("Type")
 Case "Dessin": DR_TempsMec.Sections("Section4").Controls("lblTempsDessinReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "Coupe": DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "Machinage": DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "Soudure": DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "Assemblage": DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "Peinture": DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "Test": DR_TempsMec.Sections("Section4").Controls("lblTempsTestReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "Installation": DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "Formation": DR_TempsMec.Sections("Section4").Controls("lblTempsFormationReel").Caption = Round(rstTotal.Fields("Total"), 2)
1  Case "Gestion": DR_TempsMec.Sections("Section4").Controls("lblTempsGestionReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "Shipping": DR_TempsMec.Sections("Section4").Controls("lblTempsShippingReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "Prototypage-Dévelloppement expérimental": DR_TempsMec.Sections("Section4").Controls("lblTempsPrototypeReel").Caption = Round(rstTotal.Fields("Total"), 2)
 
 End Select
 End If

 Call rstTotal.MoveNext
Loop

Call rstTotal.Close
 
Call rstTotal.Open("SELECT " & sTotal & " FROM GrbPunch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null", g_connData, adOpenDynamic, adLockOptimistic)

If Not IsNull(rstTotal.Fields("Total")) Then
 DR_TempsMec.Sections("Section4").Controls("lblTotalTempsReel").Caption = Round(rstTotal.Fields("Total"), 2)
Else
 DR_TempsMec.Sections("Section4").Controls("lblTotalTempsReel").Caption = "0"
End If

2  Call rstTotal.Close
Set rstTotal = Nothing

2  Exit Sub

Oups:

wOups "frmDetailTemps", "CalculerTempsReelsMec", Err, Err.number, Err.Description
End Sub

Private Sub cmdOK_Click()

 On Error GoTo Oups

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmDetailTemps", "cmdOK_Click", Err, Err.number, Err.Description
End Sub


Private Function vb_to_excel()


5

  Dim iCount As Integer
 Dim oXLApp As Excel.Application 'Declare the object variables
 Dim oXLBook As Excel.Workbook
 Dim oXLSheet As Excel.Worksheet
 Dim data_array(1 To 500, 1 To 4) As Variant
 Dim r As Integer
 Set oXLApp = New Excel.Application 'Create a new instance of Excel
 Set oXLBook = oXLApp.Workbooks.Add 'Add a new workbook
 Set oXLSheet = oXLBook.Worksheets(1) 'Work with the first worksheet
 oXLApp.Visible = False

'on inscrit les valeurs du listbox dans un tableau
r = 1
Do While r <= lvwTemps.ListItems.count <> Empty
 data_array(r, 1) = lvwTemps.ListItems(r)
 data_array(r, 2) = lvwTemps.ListItems(r).SubItems(1)
 data_array(r, 3) = CDbl(lvwTemps.ListItems(r).SubItems(2))
 r = r + 1
 
Loop

'ajustement largeur des colonne
oXLSheet.Columns(1).ColumnWidth = 30
oXLSheet.Columns(2).ColumnWidth = 30
oXLSheet.Columns(3).ColumnWidth = 10

'creation en-tête de colonne
oXLSheet.range("A1: C1").Font.Bold = True
oXLSheet.range("A1: C1").Value = Array("Employé", "Type", "heures")

'inscription des valeur du tableau dans excel
oXLSheet.range("A2").Resize(r, 3).Value = data_array

oXLApp.Visible = True

 



End Function

Private Sub Command1_Click()

Call vb_to_excel










End Sub

