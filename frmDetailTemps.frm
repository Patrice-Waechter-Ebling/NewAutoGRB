VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDetailTemps 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Détail des temps"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
Private Const I_COL_TYPE    As Integer = 1
Private Const I_COL_HEURES  As Integer = 2

Private m_sNoProjet As String
Private m_bProjet   As Boolean
Private m_eType     As enumCatalogue

Public Sub Afficher(ByVal sNoProjet As String, ByVal eType As enumCatalogue, ByVal bProjet As Boolean)
  
5       On Error GoTo AfficherErreur

10      m_eType = eType
 
15      m_sNoProjet = sNoProjet

20      m_bProjet = bProjet
 
25      Call RemplirListViewTemps(sNoProjet)

30      Call Show(vbModal)

35      Exit Sub

AfficherErreur:
  
40      woups "frmDetailTemps", "Afficher", Err, Erl
End Sub

Private Sub RemplirListViewTemps(ByVal sNoProjet As String)

5       On Error GoTo AfficherErreur

10      Dim rstPunch        As ADODB.Recordset
15      Dim itmPunch        As ListItem
20      Dim sFilterNoProjet As String

25      If Right$(sNoProjet, 2) = "99" Then
30        sFilterNoProjet = "LEFT(NoProjet, 6) = '" & Left$(sNoProjet, 6) & "'"
35      Else
40        sFilterNoProjet = "NoProjet = '" & sNoProjet & "'"
45      End If

50      Set rstPunch = New ADODB.Recordset

55      Call rstPunch.Open("SELECT Employe, Type, (Sum(TimeSerial(LEFT(HeureFin, 2), RIGHT(HeureFin, 2), 0) - TimeSerial(LEFT(HeureDébut, 2), RIGHT(HeureDébut, 2), 0)) * 24) AS TotalHeures FROM GRB_Punch INNER JOIN GRB_Employés ON GRB_Punch.NoEmploye = GRB_Employés.NoEmploye WHERE HeureDébut is Not Null And HeureFin is Not Null AND " & sFilterNoProjet & " GROUP BY Employe, Type", g_connData, adOpenDynamic, adLockOptimistic)

60      Do While Not rstPunch.EOF
65        Set itmPunch = lvwTemps.ListItems.Add

70        itmPunch.Text = rstPunch.Fields("Employe")
        'Retire pour afficher le type vu au lieu des spécifique GLL
        
75        'If Not IsNull(rstPunch.Fields("Type")) Then
80        '  If m_eType = ELECTRIQUE Then
85        '    Select Case rstPunch.Fields("Type")
          '      Case "Dessin":        itmPunch.SubItems(I_COL_TYPE) = "Dessin"
90        '      Case "Fabrication":   itmPunch.SubItems(I_COL_TYPE) = "Fabrication"
95        '      Case "Assemblage":    itmPunch.SubItems(I_COL_TYPE) = "Assemblage"
100       '      Case "ProgInterface": itmPunch.SubItems(I_COL_TYPE) = "Programmation d'interface"
105       '      Case "ProgAutomate":  itmPunch.SubItems(I_COL_TYPE) = "Programmation d'automate"
110       '      Case "ProgRobot":     itmPunch.SubItems(I_COL_TYPE) = "Programmation de robot"
115       '      Case "Vision":        itmPunch.SubItems(I_COL_TYPE) = "Vision"
120       '      Case "Test":          itmPunch.SubItems(I_COL_TYPE) = "Test"
125       '      Case "Installation":  itmPunch.SubItems(I_COL_TYPE) = "Installation"
130       '      Case "MiseService":   itmPunch.SubItems(I_COL_TYPE) = "Mise en service"
135       '      Case "Formation":     itmPunch.SubItems(I_COL_TYPE) = "Formation du personnel"
140       '      Case "Gestion":       itmPunch.SubItems(I_COL_TYPE) = "Gestion du projet"
145       '      Case "Shipping":      itmPunch.SubItems(I_COL_TYPE) = "Expédition"
          '      Case "Prototypage-Dévelloppement expérimental":      itmPunch.SubItems(I_COL_TYPE) = "Prototypage-Dévelloppement expérimental"
150       '    End Select
155       '  Else
160       '    Select Case rstPunch.Fields("Type")
          '     Case "Dessin":       itmPunch.SubItems(I_COL_TYPE) = "Conception et dessins"
165       '      Case "Coupe":        itmPunch.SubItems(I_COL_TYPE) = "Coupe et préparation (sauf soudage)"
170       '      Case "Machinage":    itmPunch.SubItems(I_COL_TYPE) = "Machinage"
175       '      Case "Soudure":      itmPunch.SubItems(I_COL_TYPE) = "Coupe, soudure et meulage"
180       '      Case "Assemblage":   itmPunch.SubItems(I_COL_TYPE) = "Assemblage des systèmes"
185       '      Case "Peinture":     itmPunch.SubItems(I_COL_TYPE) = "Peinture et finition"
190       '      Case "Test":         itmPunch.SubItems(I_COL_TYPE) = "Tests finaux"
195       '      Case "Installation": itmPunch.SubItems(I_COL_TYPE) = "Installation"
200       '      Case "Formation":    itmPunch.SubItems(I_COL_TYPE) = "Formation du formation"
205       '      Case "Gestion":      itmPunch.SubItems(I_COL_TYPE) = "Gestion du projet"
210       '      Case "Shipping":     itmPunch.SubItems(I_COL_TYPE) = "Expédition"
          '      Case "Prototypage-Dévelloppement expérimental":     itmPunch.SubItems(I_COL_TYPE) = "Prototypage-Dévelloppement expérimental"
215       '    End Select
220       '  End If
225       'Else
230         itmPunch.SubItems(I_COL_TYPE) = rstPunch.Fields("Type")
235       'End If
            
240       itmPunch.SubItems(I_COL_HEURES) = Round(rstPunch.Fields("TotalHeures"), 2)

245       Call rstPunch.MoveNext
250     Loop

255     Call rstPunch.Close
260     Set rstPunch = Nothing

265     Exit Sub

AfficherErreur:

270     woups "frmDetailTemps", "RemplirListViewTemps", Err, Erl
End Sub

Private Sub cmdImprimer_Click()
Dim intdummie As Integer

5       On Error GoTo AfficherErreur

10      If m_eType = ELECTRIQUE Then

            'demande d'ecriture dans excel
             intdummie = MsgBox("Désirez-vous exporter les données dans Excel ?", vbYesNo + vbInformation, "Exportation dans Excel")
            If intdummie = vbYes Then
                Call vb_to_excel
            
            End If

15        Call ImprimerDetailTempsElectriques
20      Else

            'demande d'ecriture dans excel
             intdummie = MsgBox("Désirez-vous exporter les données dans Excel ?", vbYesNo + vbInformation, "Exportation dans Excel")
            If intdummie = vbYes Then
                Call vb_to_excel
            
            End If

25        Call ImprimerDetailTempsMecaniques
30      End If



35      Exit Sub

AfficherErreur:

40      woups "frmDetailTemps", "cmdImprimer_Click", Err, Erl
End Sub

Private Sub ImprimerDetailTempsElectriques()
  
5       On Error GoTo AfficherErreur

10      Dim rstEmploye      As ADODB.Recordset
15      Dim rstImpTemps     As ADODB.Recordset
20      Dim rstProjSoum     As ADODB.Recordset
25      Dim dblTotal        As Double
30      Dim sFilterNoProjet As String

35      If Right$(m_sNoProjet, 2) = "99" Then
40        sFilterNoProjet = "LEFT(NoProjet, 6) = '" & Left$(m_sNoProjet, 6) & "'"
45      Else
50        sFilterNoProjet = "NoProjet = '" & m_sNoProjet & "'"
55      End If

60      Set rstEmploye = New ADODB.Recordset

65      Call rstEmploye.Open("SELECT Employe, Type, (Sum(TimeSerial(Left(HeureFin,2), RIGHT(HeureFin,2),0) - TimeSerial(Left(HeureDébut,2), RIGHT(HeureDébut,2),0)) * 24) AS TotalHeures FROM GRB_Punch INNER JOIN GRB_Employés ON GRB_Punch.NoEmploye = GRB_Employés.NoEmploye WHERE HeureDébut is Not Null And HeureFin is Not Null AND " & sFilterNoProjet & " GROUP BY Employe, Type", g_connData, adOpenDynamic, adLockOptimistic)

70      Call g_connData.Execute("DELETE * FROM GRB_ImpressionDetailTemps")

75      Set rstImpTemps = New ADODB.Recordset

80      Call rstImpTemps.Open("SELECT * FROM GRB_ImpressionDetailTemps", g_connData, adOpenDynamic, adLockOptimistic)

85      Do While Not rstEmploye.EOF
90        Call rstImpTemps.AddNew

95        rstImpTemps.Fields("Employe") = rstEmploye.Fields("Employe")

100       If Not IsNull(rstEmploye.Fields("Type")) Then
105         'Retirer pour afficher que le type vu GLL v41 2017-08-23
            'Select Case rstEmploye.Fields("Type")
            '  Case "Dessin":        rstImpTemps.Fields("Type") = "Dessin"
110         '  Case "Fabrication":   rstImpTemps.Fields("Type") = "Fabrication"
115         '  Case "Assemblage":    rstImpTemps.Fields("Type") = "Assemblage"
120         '  Case "ProgInterface": rstImpTemps.Fields("Type") = "Programmation d'interface"
125         '  Case "ProgAutomate":  rstImpTemps.Fields("Type") = "Programmation d'automate"
130         '  Case "ProgRobot":     rstImpTemps.Fields("Type") = "Programmation de robot"
135         '  Case "Vision":        rstImpTemps.Fields("Type") = "Vision"
140         '  Case "Test":          rstImpTemps.Fields("Type") = "Test"
145         '  Case "Installation":  rstImpTemps.Fields("Type") = "Installation"
150         '  Case "MiseService":   rstImpTemps.Fields("Type") = "Mise en service"
155         '  Case "Formation":     rstImpTemps.Fields("Type") = "Formation du personnel"
160         '  Case "Gestion":       rstImpTemps.Fields("Type") = "Gestion du projet"
165         '  Case "Shipping":      rstImpTemps.Fields("Type") = "Expédition"
            '  Case "Prototypage-Dévelloppement expérimental":      rstImpTemps.Fields("Type") = "Prototypage-Dévelloppement expérimental"
170         'End Select
175         rstImpTemps.Fields("Type") = rstEmploye.Fields("Type")
176        Else
180         rstImpTemps.Fields("Type") = ""
185       End If

190       rstImpTemps.Fields("TotalHeures") = rstEmploye.Fields("TotalHeures")

195       Call rstImpTemps.Update

200       Call rstEmploye.MoveNext
205     Loop

210     Call rstEmploye.Close
215     Set rstEmploye = Nothing

220     Set DR_TempsElec.DataSource = rstImpTemps

225     Set rstProjSoum = New ADODB.Recordset

230     If m_bProjet = True Then
235       Call rstProjSoum.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & m_sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
240     Else
245       Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionElec WHERE IDSoumission = '" & m_sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
250     End If

        'Affichage du # de projet ou soumissin
255     DR_TempsElec.Sections("Section4").Controls("lblNoProjet").Caption = m_sNoProjet

260     If Not IsNull(rstProjSoum.Fields("TempsDessin")) Then
265       DR_TempsElec.Sections("Section4").Controls("lblTempsDessinEstime").Caption = Round(rstProjSoum.Fields("TempsDessin"), 2)
270     Else
275       DR_TempsElec.Sections("Section4").Controls("lblTempsDessinEstime").Caption = "0"
280     End If

285     If Not IsNull(rstProjSoum.Fields("TempsFabrication")) Then
290       DR_TempsElec.Sections("Section4").Controls("lblTempsFabricationEstime").Caption = Round(rstProjSoum.Fields("TempsFabrication"), 2)
295     Else
300       DR_TempsElec.Sections("Section4").Controls("lblTempsFabricationEstime").Caption = "0"
305     End If

310     If Not IsNull(rstProjSoum.Fields("TempsAssemblage")) Then
315       DR_TempsElec.Sections("Section4").Controls("lblTempsAssemblageEstime").Caption = Round(rstProjSoum.Fields("TempsAssemblage"), 2)
320     Else
325       DR_TempsElec.Sections("Section4").Controls("lblTempsAssemblageEstime").Caption = "0"
330     End If

335     If Not IsNull(rstProjSoum.Fields("TempsProgInterface")) Then
340       DR_TempsElec.Sections("Section4").Controls("lblTempsProgInterfaceEstime").Caption = Round(rstProjSoum.Fields("TempsProgInterface"), 2)
345     Else
350       DR_TempsElec.Sections("Section4").Controls("lblTempsProgInterfaceEstime").Caption = "0"
355     End If

360     If Not IsNull(rstProjSoum.Fields("TempsProgAutomate")) Then
365       DR_TempsElec.Sections("Section4").Controls("lblTempsProgAutomateEstime").Caption = Round(rstProjSoum.Fields("TempsProgAutomate"), 2)
370     Else
375       DR_TempsElec.Sections("Section4").Controls("lblTempsProgAutomateEstime").Caption = "0"
380     End If

385     If Not IsNull(rstProjSoum.Fields("TempsProgRobot")) Then
390       DR_TempsElec.Sections("Section4").Controls("lblTempsProgRobotEstime").Caption = Round(rstProjSoum.Fields("TempsProgRobot"), 2)
395     Else
400       DR_TempsElec.Sections("Section4").Controls("lblTempsProgRobotEstime").Caption = "0"
405     End If

410     If Not IsNull(rstProjSoum.Fields("TempsVision")) Then
415       DR_TempsElec.Sections("Section4").Controls("lblTempsVisionEstime").Caption = Round(rstProjSoum.Fields("TempsVision"), 2)
420     Else
425       DR_TempsElec.Sections("Section4").Controls("lblTempsVisionEstime").Caption = "0"
430     End If

435     If Not IsNull(rstProjSoum.Fields("TempsTest")) Then
440       DR_TempsElec.Sections("Section4").Controls("lblTempsTestEstime").Caption = Round(rstProjSoum.Fields("TempsTest"), 2)
445     Else
450       DR_TempsElec.Sections("Section4").Controls("lblTempsTestEstime").Caption = "0"
455     End If

460     If Not IsNull(rstProjSoum.Fields("TempsInstallation")) Then
465       DR_TempsElec.Sections("Section4").Controls("lblTempsInstallationEstime").Caption = Round(rstProjSoum.Fields("TempsInstallation"), 2)
470     Else
475       DR_TempsElec.Sections("Section4").Controls("lblTempsInstallationEstime").Caption = "0"
480     End If

485     If Not IsNull(rstProjSoum.Fields("TempsMiseService")) Then
490       DR_TempsElec.Sections("Section4").Controls("lblTempsMiseServiceEstime").Caption = Round(rstProjSoum.Fields("TempsMiseService"), 2)
495     Else
500       DR_TempsElec.Sections("Section4").Controls("lblTempsMiseServiceEstime").Caption = "0"
505     End If

510     If Not IsNull(rstProjSoum.Fields("TempsFormation")) Then
515       DR_TempsElec.Sections("Section4").Controls("lblTempsFormationEstime").Caption = Round(rstProjSoum.Fields("TempsFormation"), 2)
520     Else
525       DR_TempsElec.Sections("Section4").Controls("lblTempsFormationEstime").Caption = "0"
530     End If

535     If Not IsNull(rstProjSoum.Fields("TempsGestion")) Then
540       DR_TempsElec.Sections("Section4").Controls("lblTempsGestionEstime").Caption = Round(rstProjSoum.Fields("TempsGestion"), 2)
545     Else
550       DR_TempsElec.Sections("Section4").Controls("lblTempsGestionEstime").Caption = "0"
555     End If

560     If Not IsNull(rstProjSoum.Fields("TempsShipping")) Then
565       DR_TempsElec.Sections("Section4").Controls("lblTempsShippingEstime").Caption = Round(rstProjSoum.Fields("TempsShipping"), 2)
570     Else
575       DR_TempsElec.Sections("Section4").Controls("lblTempsShippingEstime").Caption = "0"
580     End If




585     dblTotal = CDbl(DR_TempsElec.Sections("Section4").Controls("lblTempsDessinEstime").Caption) + _
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

590     DR_TempsElec.Sections("Section4").Controls("lblTotalTempsEstime").Caption = dblTotal

595     Call rstProjSoum.Close
600     Set rstProjSoum = Nothing

605     Call CalculerTempsReelsElec

610     Call DR_TempsElec.Show(vbModal)

615     Call rstImpTemps.Close
620     Set rstImpTemps = Nothing

625     Exit Sub

AfficherErreur:

630     woups "frmDetailTemps", "ImprimerDetailTempsElectriques", Err, Erl
End Sub

Private Sub ImprimerDetailTempsMecaniques()
  
5       On Error GoTo AfficherErreur

10      Dim rstEmploye      As ADODB.Recordset
15      Dim rstImpTemps     As ADODB.Recordset
20      Dim rstProjSoum     As ADODB.Recordset
25      Dim rstSoum         As ADODB.Recordset
30      Dim dblTotal        As Double
35      Dim sFilterNoProjet As String

40      If Right$(m_sNoProjet, 2) = "99" Then
45        sFilterNoProjet = "LEFT(NoProjet, 6) = '" & Left$(m_sNoProjet, 6) & "'"
50      Else
55        sFilterNoProjet = "NoProjet = '" & m_sNoProjet & "'"
60      End If

65      Set rstEmploye = New ADODB.Recordset

70      Call rstEmploye.Open("SELECT Employe, Type, (Sum(TimeSerial(Left(HeureFin,2), RIGHT(HeureFin,2),0) - TimeSerial(Left(HeureDébut,2), RIGHT(HeureDébut,2),0)) * 24) AS TotalHeures FROM GRB_Punch INNER JOIN GRB_Employés ON GRB_Punch.NoEmploye = GRB_Employés.NoEmploye WHERE HeureDébut is Not Null And HeureFin is Not Null AND " & sFilterNoProjet & " GROUP BY Employe, Type", g_connData, adOpenDynamic, adLockOptimistic)

75      Call g_connData.Execute("DELETE * FROM GRB_ImpressionDetailTemps")

80      Set rstImpTemps = New ADODB.Recordset

85      Call rstImpTemps.Open("SELECT * FROM GRB_ImpressionDetailTemps", g_connData, adOpenDynamic, adLockOptimistic)

90      Do While Not rstEmploye.EOF
95        Call rstImpTemps.AddNew

100       rstImpTemps.Fields("Employe") = rstEmploye.Fields("Employe")

105       If Not IsNull(rstEmploye.Fields("Type")) Then
110         'Retirer par GLL V41 2017-08-23
            'Select Case rstEmploye.Fields("Type")
            '  Case "Dessin":       rstImpTemps.Fields("Type") = "Conception et dessins"
115         '  Case "Coupe":        rstImpTemps.Fields("Type") = "Coupe et préparation (sauf soudage)"
120         '  Case "Machinage":    rstImpTemps.Fields("Type") = "Machinage"
125         '  Case "Soudure":      rstImpTemps.Fields("Type") = "Coupe, soudure et meulage"
130         '  Case "Assemblage":   rstImpTemps.Fields("Type") = "Assemblage des systèmes"
135         '  Case "Peinture":     rstImpTemps.Fields("Type") = "Peinture et finition"
140         '  Case "Test":         rstImpTemps.Fields("Type") = "Tests finaux"
145         '  Case "Installation": rstImpTemps.Fields("Type") = "Installation"
150         '  Case "Formation":    rstImpTemps.Fields("Type") = "Formation du personnel"
155         '  Case "Gestion":      rstImpTemps.Fields("Type") = "Gestion du projet"
160         '  Case "Shipping":     rstImpTemps.Fields("Type") = "Expédition"
165         'End Select
170         rstImpTemps.Fields("Type") = rstEmploye.Fields("Type")
173          Else
175         rstImpTemps.Fields("Type") = ""
180       End If

185       rstImpTemps.Fields("TotalHeures") = rstEmploye.Fields("TotalHeures")

190       Call rstImpTemps.Update

195       Call rstEmploye.MoveNext
200     Loop

205     Call rstEmploye.Close
210     Set rstEmploye = Nothing

215     Set DR_TempsMec.DataSource = rstImpTemps

220     Set rstProjSoum = New ADODB.Recordset

225     If m_bProjet = True Then
230       Call rstProjSoum.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & m_sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
235     Else
240       Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionMec WHERE IDSoumission = '" & m_sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
245     End If

        'Affichage du # de projet ou soumission
250     DR_TempsMec.Sections("Section4").Controls("lblNoProjet").Caption = m_sNoProjet

        'Si soumission
255     If m_bProjet = False Then
260       If Not IsNull(rstProjSoum.Fields("TempsDessin")) Then
265         DR_TempsMec.Sections("Section4").Controls("lblTempsDessinSoum").Caption = Round(rstProjSoum.Fields("TempsDessin"), 2)
270       Else
275         DR_TempsMec.Sections("Section4").Controls("lblTempsDessinSoum").Caption = "0"
280       End If

285       If Not IsNull(rstProjSoum.Fields("TempsCoupe")) Then
290         DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeSoum").Caption = Round(rstProjSoum.Fields("TempsCoupe"), 2)
295       Else
300         DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeSoum").Caption = "0"
305       End If

310       If Not IsNull(rstProjSoum.Fields("TempsMachinage")) Then
315         DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageSoum").Caption = Round(rstProjSoum.Fields("TempsMachinage"), 2)
320       Else
325         DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageSoum").Caption = "0"
330       End If

335       If Not IsNull(rstProjSoum.Fields("TempsSoudure")) Then
340         DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureSoum").Caption = Round(rstProjSoum.Fields("TempsSoudure"), 2)
345       Else
350         DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureSoum").Caption = "0"
355       End If

360       If Not IsNull(rstProjSoum.Fields("TempsAssemblage")) Then
365         DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageSoum").Caption = Round(rstProjSoum.Fields("TempsAssemblage"), 2)
370       Else
375         DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageSoum").Caption = "0"
380       End If

385       If Not IsNull(rstProjSoum.Fields("TempsPeinture")) Then
390         DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureSoum").Caption = Round(rstProjSoum.Fields("TempsPeinture"), 2)
395       Else
400         DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureSoum").Caption = "0"
405       End If

410       If Not IsNull(rstProjSoum.Fields("TempsTest")) Then
415         DR_TempsMec.Sections("Section4").Controls("lblTempsTestSoum").Caption = Round(rstProjSoum.Fields("TempsTest"), 2)
420       Else
425         DR_TempsMec.Sections("Section4").Controls("lblTempsTestSoum").Caption = "0"
430       End If

435       If Not IsNull(rstProjSoum.Fields("TempsInstallation")) Then
440         DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationSoum").Caption = Round(rstProjSoum.Fields("TempsInstallation"), 2)
445       Else
450         DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationSoum").Caption = "0"
455       End If

460       If Not IsNull(rstProjSoum.Fields("TempsFormation")) Then
465         DR_TempsMec.Sections("Section4").Controls("lblTempsFormationSoum").Caption = Round(rstProjSoum.Fields("TempsFormation"), 2)
470       Else
475         DR_TempsMec.Sections("Section4").Controls("lblTempsFormationSoum").Caption = "0"
480       End If

485       If Not IsNull(rstProjSoum.Fields("TempsGestion")) Then
490         DR_TempsMec.Sections("Section4").Controls("lblTempsGestionSoum").Caption = Round(rstProjSoum.Fields("TempsGestion"), 2)
495       Else
500         DR_TempsMec.Sections("Section4").Controls("lblTempsGestionSoum").Caption = "0"
505       End If

510       If Not IsNull(rstProjSoum.Fields("TempsShipping")) Then
515         DR_TempsMec.Sections("Section4").Controls("lblTempsShippingSoum").Caption = Round(rstProjSoum.Fields("TempsShipping"), 2)
520       Else
525         DR_TempsMec.Sections("Section4").Controls("lblTempsShippingSoum").Caption = "0"
530       End If

535       DR_TempsMec.Sections("Section4").Controls("lblTempsDessinProj").Caption = "---"
540       DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeProj").Caption = "---"
545       DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageProj").Caption = "---"
550       DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureProj").Caption = "---"
555       DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageProj").Caption = "---"
560       DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureProj").Caption = "---"
565       DR_TempsMec.Sections("Section4").Controls("lblTempsTestProj").Caption = "---"
570       DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationProj").Caption = "---"
575       DR_TempsMec.Sections("Section4").Controls("lblTempsFormationProj").Caption = "---"
580       DR_TempsMec.Sections("Section4").Controls("lblTempsGestionProj").Caption = "---"
585       DR_TempsMec.Sections("Section4").Controls("lblTempsShippingProj").Caption = "---"

590       DR_TempsMec.Sections("Section4").Controls("lblTempsDessinConc").Caption = "---"
595       DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeConc").Caption = "---"
600       DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageConc").Caption = "---"
605       DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureConc").Caption = "---"
610       DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageConc").Caption = "---"
615       DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureConc").Caption = "---"
620       DR_TempsMec.Sections("Section4").Controls("lblTempsTestConc").Caption = "---"
625       DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationConc").Caption = "---"
630       DR_TempsMec.Sections("Section4").Controls("lblTempsFormationConc").Caption = "---"
635       DR_TempsMec.Sections("Section4").Controls("lblTempsGestionConc").Caption = "---"
640       DR_TempsMec.Sections("Section4").Controls("lblTempsShippingConc").Caption = "---"
645     Else
650       If Not IsNull(rstProjSoum.Fields("IDSoumission")) Then
655         Set rstSoum = New ADODB.Recordset

660         Call rstSoum.Open("SELECT * FROM GRB_SoumissionMec WHERE IDSoumission = '" & m_sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

665         If Not rstSoum.EOF Then
670           If Not IsNull(rstSoum.Fields("TempsDessin")) Then
675             DR_TempsMec.Sections("Section4").Controls("lblTempsDessinSoum").Caption = Round(rstSoum.Fields("TempsDessin"), 2)
680           Else
685             DR_TempsMec.Sections("Section4").Controls("lblTempsDessinSoum").Caption = "0"
690           End If

695           If Not IsNull(rstSoum.Fields("TempsCoupe")) Then
700             DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeSoum").Caption = Round(rstSoum.Fields("TempsCoupe"), 2)
705           Else
710             DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeSoum").Caption = "0"
715           End If

720           If Not IsNull(rstSoum.Fields("TempsMachinage")) Then
725             DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageSoum").Caption = Round(rstSoum.Fields("TempsMachinage"), 2)
730           Else
735             DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageSoum").Caption = "0"
740           End If

745           If Not IsNull(rstSoum.Fields("TempsSoudure")) Then
750             DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureSoum").Caption = Round(rstSoum.Fields("TempsSoudure"), 2)
755           Else
760             DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureSoum").Caption = "0"
765           End If

770           If Not IsNull(rstSoum.Fields("TempsAssemblage")) Then
775             DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageSoum").Caption = Round(rstSoum.Fields("TempsAssemblage"), 2)
780           Else
785             DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageSoum").Caption = "0"
790           End If

795           If Not IsNull(rstSoum.Fields("TempsPeinture")) Then
800             DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureSoum").Caption = Round(rstSoum.Fields("TempsPeinture"), 2)
805           Else
810             DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureSoum").Caption = "0"
815           End If

820           If Not IsNull(rstSoum.Fields("TempsTest")) Then
825             DR_TempsMec.Sections("Section4").Controls("lblTempsTestSoum").Caption = Round(rstSoum.Fields("TempsTest"), 2)
830           Else
835             DR_TempsMec.Sections("Section4").Controls("lblTempsTestSoum").Caption = "0"
840           End If

845           If Not IsNull(rstSoum.Fields("TempsInstallation")) Then
850             DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationSoum").Caption = Round(rstSoum.Fields("TempsInstallation"), 2)
855           Else
860             DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationSoum").Caption = "0"
865           End If

870           If Not IsNull(rstSoum.Fields("TempsFormation")) Then
875             DR_TempsMec.Sections("Section4").Controls("lblTempsFormationSoum").Caption = Round(rstSoum.Fields("TempsFormation"), 2)
880           Else
885             DR_TempsMec.Sections("Section4").Controls("lblTempsFormationSoum").Caption = "0"
890           End If

895           If Not IsNull(rstSoum.Fields("TempsGestion")) Then
900             DR_TempsMec.Sections("Section4").Controls("lblTempsGestionSoum").Caption = Round(rstSoum.Fields("TempsGestion"), 2)
905           Else
910             DR_TempsMec.Sections("Section4").Controls("lblTempsGestionSoum").Caption = "0"
915           End If

920           If Not IsNull(rstSoum.Fields("TempsShipping")) Then
925             DR_TempsMec.Sections("Section4").Controls("lblTempsShippingSoum").Caption = Round(rstSoum.Fields("TempsShipping"), 2)
930           Else
935             DR_TempsMec.Sections("Section4").Controls("lblTempsShippingSoum").Caption = "0"
940           End If
945         Else
950           DR_TempsMec.Sections("Section4").Controls("lblTempsDessinSoum").Caption = "---"
955           DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeSoum").Caption = "---"
960           DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageSoum").Caption = "---"
965           DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureSoum").Caption = "---"
970           DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageSoum").Caption = "---"
975           DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureSoum").Caption = "---"
980           DR_TempsMec.Sections("Section4").Controls("lblTempsTestSoum").Caption = "---"
985           DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationSoum").Caption = "---"
990           DR_TempsMec.Sections("Section4").Controls("lblTempsFormationSoum").Caption = "---"
995           DR_TempsMec.Sections("Section4").Controls("lblTempsGestionSoum").Caption = "---"
1000          DR_TempsMec.Sections("Section4").Controls("lblTempsShippingSoum").Caption = "---"
1005        End If

1010        Call rstSoum.Close
1015        Set rstSoum = Nothing
1020      Else
1025        DR_TempsMec.Sections("Section4").Controls("lblTempsDessinSoum").Caption = "---"
1030        DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeSoum").Caption = "---"
1035        DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageSoum").Caption = "---"
1040        DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureSoum").Caption = "---"
1045        DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageSoum").Caption = "---"
1050        DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureSoum").Caption = "---"
1055        DR_TempsMec.Sections("Section4").Controls("lblTempsTestSoum").Caption = "---"
1060        DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationSoum").Caption = "---"
1065        DR_TempsMec.Sections("Section4").Controls("lblTempsFormationSoum").Caption = "---"
1070        DR_TempsMec.Sections("Section4").Controls("lblTempsGestionSoum").Caption = "---"
1075        DR_TempsMec.Sections("Section4").Controls("lblTempsShippingSoum").Caption = "---"
1080      End If

1085      If Not IsNull(rstProjSoum.Fields("TempsDessinProj")) Then
1090        DR_TempsMec.Sections("Section4").Controls("lblTempsDessinProj").Caption = Round(rstProjSoum.Fields("TempsDessinProj"), 2)
1095      Else
1100        DR_TempsMec.Sections("Section4").Controls("lblTempsDessinProj").Caption = "0"
1105      End If

1110      If Not IsNull(rstProjSoum.Fields("TempsCoupeProj")) Then
1115        DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeProj").Caption = Round(rstProjSoum.Fields("TempsCoupeProj"), 2)
1120      Else
1125        DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeProj").Caption = "0"
1130      End If

1135      If Not IsNull(rstProjSoum.Fields("TempsMachinageProj")) Then
1140        DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageProj").Caption = Round(rstProjSoum.Fields("TempsMachinageProj"), 2)
1145      Else
1150        DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageProj").Caption = "0"
1155      End If

1160      If Not IsNull(rstProjSoum.Fields("TempsSoudureProj")) Then
1165        DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureProj").Caption = Round(rstProjSoum.Fields("TempsSoudureProj"), 2)
1170      Else
1175        DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureProj").Caption = "0"
1180      End If

1185      If Not IsNull(rstProjSoum.Fields("TempsAssemblageProj")) Then
1190        DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageProj").Caption = Round(rstProjSoum.Fields("TempsAssemblageProj"), 2)
1195      Else
1200        DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageProj").Caption = "0"
1205      End If

1210      If Not IsNull(rstProjSoum.Fields("TempsPeintureProj")) Then
1215        DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureProj").Caption = Round(rstProjSoum.Fields("TempsPeintureProj"), 2)
1220      Else
1225        DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureProj").Caption = "0"
1230      End If

1235      If Not IsNull(rstProjSoum.Fields("TempsTestProj")) Then
1240        DR_TempsMec.Sections("Section4").Controls("lblTempsTestProj").Caption = Round(rstProjSoum.Fields("TempsTestProj"), 2)
1245      Else
1250        DR_TempsMec.Sections("Section4").Controls("lblTempsTestProj").Caption = "0"
1255      End If

1260      If Not IsNull(rstProjSoum.Fields("TempsInstallationProj")) Then
1265        DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationProj").Caption = Round(rstProjSoum.Fields("TempsInstallationProj"), 2)
1270      Else
1275        DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationProj").Caption = "0"
1280      End If

1285      If Not IsNull(rstProjSoum.Fields("TempsFormationProj")) Then
1290        DR_TempsMec.Sections("Section4").Controls("lblTempsFormationProj").Caption = Round(rstProjSoum.Fields("TempsFormationProj"), 2)
1295      Else
1300        DR_TempsMec.Sections("Section4").Controls("lblTempsFormationProj").Caption = "0"
1305      End If

1310      If Not IsNull(rstProjSoum.Fields("TempsGestionProj")) Then
1315        DR_TempsMec.Sections("Section4").Controls("lblTempsGestionProj").Caption = Round(rstProjSoum.Fields("TempsGestionProj"), 2)
1320      Else
1325        DR_TempsMec.Sections("Section4").Controls("lblTempsGestionProj").Caption = "0"
1330      End If

1335      If Not IsNull(rstProjSoum.Fields("TempsShippingProj")) Then
1340        DR_TempsMec.Sections("Section4").Controls("lblTempsShippingProj").Caption = Round(rstProjSoum.Fields("TempsShippingProj"), 2)
1345      Else
1350        DR_TempsMec.Sections("Section4").Controls("lblTempsShippingProj").Caption = "0"
1355      End If

1360      If rstProjSoum.Fields("TempsProjBarré") = True Then
1365        If Not IsNull(rstProjSoum.Fields("TempsDessinConc")) Then
1370          DR_TempsMec.Sections("Section4").Controls("lblTempsDessinConc").Caption = Round(rstProjSoum.Fields("TempsDessinConc"), 2)
1375        Else
1380          DR_TempsMec.Sections("Section4").Controls("lblTempsDessinConc").Caption = "0"
1385        End If

1390        If Not IsNull(rstProjSoum.Fields("TempsCoupeConc")) Then
1395          DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeConc").Caption = Round(rstProjSoum.Fields("TempsCoupeConc"), 2)
1400        Else
1405          DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeConc").Caption = "0"
1410        End If

1415        If Not IsNull(rstProjSoum.Fields("TempsMachinageConc")) Then
1420          DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageConc").Caption = Round(rstProjSoum.Fields("TempsMachinageConc"), 2)
1425        Else
1430          DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageConc").Caption = "0"
1435        End If

1440        If Not IsNull(rstProjSoum.Fields("TempsSoudureConc")) Then
1445          DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureConc").Caption = Round(rstProjSoum.Fields("TempsSoudureConc"), 2)
1450        Else
1455          DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureConc").Caption = "0"
1460        End If

1465        If Not IsNull(rstProjSoum.Fields("TempsAssemblageConc")) Then
1470          DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageConc").Caption = Round(rstProjSoum.Fields("TempsAssemblageConc"), 2)
1475        Else
1480          DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageConc").Caption = "0"
1485        End If

1490        If Not IsNull(rstProjSoum.Fields("TempsPeintureConc")) Then
1495          DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureConc").Caption = Round(rstProjSoum.Fields("TempsPeintureConc"), 2)
1500        Else
1505          DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureConc").Caption = "0"
1510        End If

1515        If Not IsNull(rstProjSoum.Fields("TempsTestConc")) Then
1520          DR_TempsMec.Sections("Section4").Controls("lblTempsTestConc").Caption = Round(rstProjSoum.Fields("TempsTestConc"), 2)
1525        Else
1530          DR_TempsMec.Sections("Section4").Controls("lblTempsTestConc").Caption = "0"
1535        End If

1540        If Not IsNull(rstProjSoum.Fields("TempsInstallationConc")) Then
1545          DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationConc").Caption = Round(rstProjSoum.Fields("TempsInstallationConc"), 2)
1550        Else
1555          DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationConc").Caption = "0"
1560        End If

1565        If Not IsNull(rstProjSoum.Fields("TempsFormationConc")) Then
1570          DR_TempsMec.Sections("Section4").Controls("lblTempsFormationConc").Caption = Round(rstProjSoum.Fields("TempsFormationConc"), 2)
1575        Else
1580          DR_TempsMec.Sections("Section4").Controls("lblTempsFormationConc").Caption = "0"
1585        End If

1590        If Not IsNull(rstProjSoum.Fields("TempsGestionConc")) Then
1595          DR_TempsMec.Sections("Section4").Controls("lblTempsGestionConc").Caption = Round(rstProjSoum.Fields("TempsGestionConc"), 2)
1600        Else
1605          DR_TempsMec.Sections("Section4").Controls("lblTempsGestionConc").Caption = "0"
1610        End If

1615        If Not IsNull(rstProjSoum.Fields("TempsShippingConc")) Then
1620          DR_TempsMec.Sections("Section4").Controls("lblTempsShippingConc").Caption = Round(rstProjSoum.Fields("TempsShippingConc"), 2)
1625        Else
1630          DR_TempsMec.Sections("Section4").Controls("lblTempsShippingConc").Caption = "0"
1635        End If
1640      Else
1645        DR_TempsMec.Sections("Section4").Controls("lblTempsDessinConc").Caption = "---"
1650        DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeConc").Caption = "---"
1655        DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageConc").Caption = "---"
1660        DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureConc").Caption = "---"
1665        DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageConc").Caption = "---"
1670        DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureConc").Caption = "---"
1675        DR_TempsMec.Sections("Section4").Controls("lblTempsTestConc").Caption = "---"
1680        DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationConc").Caption = "---"
1685        DR_TempsMec.Sections("Section4").Controls("lblTempsFormationConc").Caption = "---"
1690        DR_TempsMec.Sections("Section4").Controls("lblTempsGestionConc").Caption = "---"
1695        DR_TempsMec.Sections("Section4").Controls("lblTempsShippingConc").Caption = "---"
1700      End If
1705    End If

1710    If DR_TempsMec.Sections("Section4").Controls("lblTempsDessinSoum").Caption = "---" And _
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
           
1715      DR_TempsMec.Sections("Section4").Controls("lblTotalTempsSoum").Caption = "---"
1720    Else
1725      dblTotal = CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsDessinSoum").Caption) + _
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
                     
1730      DR_TempsMec.Sections("Section4").Controls("lblTotalTempsSoum").Caption = dblTotal
1735    End If


1740    If DR_TempsMec.Sections("Section4").Controls("lblTempsDessinProj").Caption = "---" And _
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
           
1745      DR_TempsMec.Sections("Section4").Controls("lblTotalTempsProj").Caption = "---"
1750    Else
1755      dblTotal = CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsDessinProj").Caption) + _
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

1760      DR_TempsMec.Sections("Section4").Controls("lblTotalTempsProj").Caption = dblTotal
1765    End If

1770    If DR_TempsMec.Sections("Section4").Controls("lblTempsDessinConc").Caption = "---" And _
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
           
1775      DR_TempsMec.Sections("Section4").Controls("lblTotalTempsConc").Caption = "---"
1780    Else
1785      dblTotal = CDbl(DR_TempsMec.Sections("Section4").Controls("lblTempsDessinConc").Caption) + _
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
                     
1790      DR_TempsMec.Sections("Section4").Controls("lblTotalTempsConc").Caption = dblTotal
1795    End If

1800    Call rstProjSoum.Close
1805    Set rstProjSoum = Nothing

1810    Call CalculerTempsReelsMec

1815    Call DR_TempsMec.Show(vbModal)

1820    Call rstImpTemps.Close
1825    Set rstImpTemps = Nothing

1830    Exit Sub

AfficherErreur:

1835    woups "frmDetailTemps", "ImprimerDetailTempsMecaniques", Err, Erl
End Sub

Private Sub CalculerTempsReelsElec()

5       On Error GoTo AfficherErreur

10      Dim rstTotal        As ADODB.Recordset
15      Dim sDateDebut      As String
20      Dim sDateFin        As String
25      Dim sTotal          As String
30      Dim sFilterNoProjet As String

35      If Right$(m_sNoProjet, 2) = "99" Then
40        sFilterNoProjet = "LEFT(NoProjet, 6) = '" & Left$(m_sNoProjet, 6) & "'"
45      Else
50        sFilterNoProjet = "NoProjet = '" & m_sNoProjet & "'"
55      End If

60      Set rstTotal = New ADODB.Recordset
  
65      sDateDebut = "TIMESERIAL(Left(GRB_Punch.HeureDébut,2),RIGHT(GRB_Punch.HeureDébut,2),0)"
  
70      sDateFin = "TIMESERIAL(Left(GRB_Punch.HeureFin,2),RIGHT(GRB_Punch.HeureFin,2),0)"
  
75      sTotal = "(SUM(" & sDateFin & " - " & sDateDebut & ")* 24) As Total"

80      Call rstTotal.Open("SELECT Type, " & sTotal & " FROM GRB_Punch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null GROUP BY Type", g_connData, adOpenDynamic, adLockOptimistic)

85      DR_TempsElec.Sections("Section4").Controls("lblTempsDessinReel").Caption = "0"
90      DR_TempsElec.Sections("Section4").Controls("lblTempsFabricationReel").Caption = "0"
95      DR_TempsElec.Sections("Section4").Controls("lblTempsAssemblageReel").Caption = "0"
100     DR_TempsElec.Sections("Section4").Controls("lblTempsProgInterfaceReel").Caption = "0"
105     DR_TempsElec.Sections("Section4").Controls("lblTempsProgAutomateReel").Caption = "0"
110     DR_TempsElec.Sections("Section4").Controls("lblTempsProgRobotReel").Caption = "0"
115     DR_TempsElec.Sections("Section4").Controls("lblTempsVisionReel").Caption = "0"
120     DR_TempsElec.Sections("Section4").Controls("lblTempsTestReel").Caption = "0"
125     DR_TempsElec.Sections("Section4").Controls("lblTempsInstallationReel").Caption = "0"
130     DR_TempsElec.Sections("Section4").Controls("lblTempsMiseServiceReel").Caption = "0"
135     DR_TempsElec.Sections("Section4").Controls("lblTempsFormationReel").Caption = "0"
140     DR_TempsElec.Sections("Section4").Controls("lblTempsGestionReel").Caption = "0"
145     DR_TempsElec.Sections("Section4").Controls("lblTempsShippingReel").Caption = "0"

150     Do While Not rstTotal.EOF
155       If Not IsNull(rstTotal.Fields("Total")) Then
160         Select Case rstTotal.Fields("Type")
              Case "Dessin":        DR_TempsElec.Sections("Section4").Controls("lblTempsDessinReel").Caption = Round(rstTotal.Fields("Total"), 2)
165           Case "Fabrication":   DR_TempsElec.Sections("Section4").Controls("lblTempsFabricationReel").Caption = Round(rstTotal.Fields("Total"), 2)
170           Case "Assemblage":    DR_TempsElec.Sections("Section4").Controls("lblTempsAssemblageReel").Caption = Round(rstTotal.Fields("Total"), 2)
175           Case "ProgInterface": DR_TempsElec.Sections("Section4").Controls("lblTempsProgInterfaceReel").Caption = Round(rstTotal.Fields("Total"), 2)
180           Case "ProgAutomate":  DR_TempsElec.Sections("Section4").Controls("lblTempsProgAutomateReel").Caption = Round(rstTotal.Fields("Total"), 2)
185           Case "ProgRobot":     DR_TempsElec.Sections("Section4").Controls("lblTempsProgRobotReel").Caption = Round(rstTotal.Fields("Total"), 2)
190           Case "Vision":        DR_TempsElec.Sections("Section4").Controls("lblTempsVisionReel").Caption = Round(rstTotal.Fields("Total"), 2)
195           Case "Test":          DR_TempsElec.Sections("Section4").Controls("lblTempsTestReel").Caption = Round(rstTotal.Fields("Total"), 2)
200           Case "Installation":  DR_TempsElec.Sections("Section4").Controls("lblTempsInstallationReel").Caption = Round(rstTotal.Fields("Total"), 2)
205           Case "MiseService":   DR_TempsElec.Sections("Section4").Controls("lblTempsMiseServiceReel").Caption = Round(rstTotal.Fields("Total"), 2)
210           Case "Formation":     DR_TempsElec.Sections("Section4").Controls("lblTempsFormationReel").Caption = Round(rstTotal.Fields("Total"), 2)
215           Case "Gestion":       DR_TempsElec.Sections("Section4").Controls("lblTempsGestionReel").Caption = Round(rstTotal.Fields("Total"), 2)
220           Case "Shipping":      DR_TempsElec.Sections("Section4").Controls("lblTempsShippingReel").Caption = Round(rstTotal.Fields("Total"), 2)
221           Case "Prototypage-Dévelloppement expérimental":      DR_TempsElec.Sections("Section4").Controls("lblTempsprototypeReel").Caption = Round(rstTotal.Fields("Total"), 2)

225         End Select
230       End If

235       Call rstTotal.MoveNext
240     Loop

245     Call rstTotal.Close
  
250     Call rstTotal.Open("SELECT " & sTotal & " FROM GRB_Punch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null", g_connData, adOpenDynamic, adLockOptimistic)

255     If Not IsNull(rstTotal.Fields("Total")) Then
260       DR_TempsElec.Sections("Section4").Controls("lblTotalTempsReel").Caption = Round(rstTotal.Fields("Total"), 2)
265     Else
270       DR_TempsElec.Sections("Section4").Controls("lblTotalTempsReel").Caption = "0"
275     End If

280     Call rstTotal.Close
285     Set rstTotal = Nothing

290     Exit Sub

AfficherErreur:

295     woups "frmDetailTemps", "CalculerTempsReelsElec", Err, Erl
End Sub

Private Sub CalculerTempsReelsMec()

5       On Error GoTo AfficherErreur

10      Dim rstTotal        As ADODB.Recordset
15      Dim sDateDebut      As String
20      Dim sDateFin        As String
25      Dim sTotal          As String
30      Dim sFilterNoProjet As String

35      If Right$(m_sNoProjet, 2) = "99" Then
40        sFilterNoProjet = "LEFT(NoProjet, 6) = '" & Left$(m_sNoProjet, 6) & "'"
45      Else
50        sFilterNoProjet = "NoProjet = '" & m_sNoProjet & "'"
55      End If

60      Set rstTotal = New ADODB.Recordset
  
65      sDateDebut = "TIMESERIAL(Left(GRB_Punch.HeureDébut,2),RIGHT(GRB_Punch.HeureDébut,2),0)"
  
70      sDateFin = "TIMESERIAL(Left(GRB_Punch.HeureFin,2),RIGHT(GRB_Punch.HeureFin,2),0)"
  
75      sTotal = "(SUM(" & sDateFin & " - " & sDateDebut & ")* 24) As Total"

80      Call rstTotal.Open("SELECT Type, " & sTotal & " FROM GRB_Punch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null GROUP BY Type", g_connData, adOpenDynamic, adLockOptimistic)

85      DR_TempsMec.Sections("Section4").Controls("lblTempsDessinReel").Caption = "0"
90      DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeReel").Caption = "0"
95      DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageReel").Caption = "0"
100     DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureReel").Caption = "0"
105     DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageReel").Caption = "0"
110     DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureReel").Caption = "0"
115     DR_TempsMec.Sections("Section4").Controls("lblTempsTestReel").Caption = "0"
120     DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationReel").Caption = "0"
125     DR_TempsMec.Sections("Section4").Controls("lblTempsFormationReel").Caption = "0"
130     DR_TempsMec.Sections("Section4").Controls("lblTempsGestionReel").Caption = "0"
135     DR_TempsMec.Sections("Section4").Controls("lblTempsShippingReel").Caption = "0"
        DR_TempsMec.Sections("Section4").Controls("lblTempsprototypeReel").Caption = "0"


140     Do While Not rstTotal.EOF
145       If Not IsNull(rstTotal.Fields("Total")) Then
150         Select Case rstTotal.Fields("Type")
              Case "Dessin":       DR_TempsMec.Sections("Section4").Controls("lblTempsDessinReel").Caption = Round(rstTotal.Fields("Total"), 2)
155           Case "Coupe":        DR_TempsMec.Sections("Section4").Controls("lblTempsCoupeReel").Caption = Round(rstTotal.Fields("Total"), 2)
160           Case "Machinage":    DR_TempsMec.Sections("Section4").Controls("lblTempsMachinageReel").Caption = Round(rstTotal.Fields("Total"), 2)
165           Case "Soudure":      DR_TempsMec.Sections("Section4").Controls("lblTempsSoudureReel").Caption = Round(rstTotal.Fields("Total"), 2)
170           Case "Assemblage":   DR_TempsMec.Sections("Section4").Controls("lblTempsAssemblageReel").Caption = Round(rstTotal.Fields("Total"), 2)
175           Case "Peinture":     DR_TempsMec.Sections("Section4").Controls("lblTempsPeintureReel").Caption = Round(rstTotal.Fields("Total"), 2)
180           Case "Test":         DR_TempsMec.Sections("Section4").Controls("lblTempsTestReel").Caption = Round(rstTotal.Fields("Total"), 2)
185           Case "Installation": DR_TempsMec.Sections("Section4").Controls("lblTempsInstallationReel").Caption = Round(rstTotal.Fields("Total"), 2)
190           Case "Formation":    DR_TempsMec.Sections("Section4").Controls("lblTempsFormationReel").Caption = Round(rstTotal.Fields("Total"), 2)
195           Case "Gestion":      DR_TempsMec.Sections("Section4").Controls("lblTempsGestionReel").Caption = Round(rstTotal.Fields("Total"), 2)
200           Case "Shipping":     DR_TempsMec.Sections("Section4").Controls("lblTempsShippingReel").Caption = Round(rstTotal.Fields("Total"), 2)
205           Case "Prototypage-Dévelloppement expérimental":     DR_TempsMec.Sections("Section4").Controls("lblTempsPrototypeReel").Caption = Round(rstTotal.Fields("Total"), 2)
 
            End Select
210       End If

215       Call rstTotal.MoveNext
220     Loop

225     Call rstTotal.Close
  
230     Call rstTotal.Open("SELECT " & sTotal & " FROM GRB_Punch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null", g_connData, adOpenDynamic, adLockOptimistic)

235     If Not IsNull(rstTotal.Fields("Total")) Then
240       DR_TempsMec.Sections("Section4").Controls("lblTotalTempsReel").Caption = Round(rstTotal.Fields("Total"), 2)
245     Else
250       DR_TempsMec.Sections("Section4").Controls("lblTotalTempsReel").Caption = "0"
255     End If

260     Call rstTotal.Close
265     Set rstTotal = Nothing

270     Exit Sub

AfficherErreur:

275     woups "frmDetailTemps", "CalculerTempsReelsMec", Err, Erl
End Sub

Private Sub cmdOK_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmDetailTemps", "cmdOK_Click", Err, Erl
End Sub


Private Function vb_to_excel()


5

6       Dim iCount As Integer
10      Dim oXLApp As Excel.Application         'Declare the object variables
15      Dim oXLBook As Excel.Workbook
20      Dim oXLSheet As Excel.Worksheet
        Dim data_array(1 To 500, 1 To 4) As Variant
        Dim r As Integer
25      Set oXLApp = New Excel.Application    'Create a new instance of Excel
30      Set oXLBook = oXLApp.Workbooks.Add    'Add a new workbook
35      Set oXLSheet = oXLBook.Worksheets(1)  'Work with the first worksheet
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
oXLSheet.Range("A1: C1").Font.Bold = True
oXLSheet.Range("A1: C1").Value = Array("Employé", "Type", "heures")

'inscription des valeur du tableau dans excel
oXLSheet.Range("A2").Resize(r, 3).Value = data_array

oXLApp.Visible = True

        



End Function

Private Sub Command1_Click()

Call vb_to_excel










End Sub

