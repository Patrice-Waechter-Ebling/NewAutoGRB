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
   MinButton       =   0   'False
   Picture         =   "frmChoixImpressionFT.frx":000C
   ScaleHeight     =   6735
   ScaleWidth      =   4260
   StartUpPosition =   2  'CenterScreen
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
Private m_datFin   As Date

Public Sub Afficher(ByVal datDebut As Date, ByVal datFin As Date)
        
5       On Error GoTo AfficherErreur

10      m_datDebut = datDebut
15      m_datFin = datFin
        
20      Call RemplirListViewEmploye

25      Call Me.Show(vbModal)

30      Exit Sub

AfficherErreur:

35      woups "frmChoixImpressionFT", "Afficher", Err, Erl
End Sub

Private Sub cmdAnnuler_Click()
  
5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmChoixImpressionFT", "cmdAnnuler_Click", Err, Erl
End Sub

Private Sub cmdImprimer_Click()

5       On Error GoTo AfficherErreur

        'Impression de la feuille de temps
10      Dim rstImpPunch As ADODB.Recordset
15      Dim rstSommeKM  As ADODB.Recordset
20      Dim rstPunch    As ADODB.Recordset
25      Dim bAnnuler    As Boolean
30      Dim iCompteur   As Integer
35      Dim dblTotal    As Double

40      Screen.MousePointer = vbHourglass

45      Set rstImpPunch = New ADODB.Recordset
50      Set rstSommeKM = New ADODB.Recordset
55      Set rstPunch = New ADODB.Recordset

60      For iCompteur = 1 To lvwEmploye.ListItems.count
65        If lvwEmploye.ListItems(iCompteur).Checked = True Then
70          dblTotal = 0

75          bAnnuler = False

80          Call g_connData.Execute("DELETE * FROM GRB_ImpressionPunch")

85          Call rstPunch.Open("SELECT * FROM GRB_Punch WHERE Date BETWEEN '" & ConvertDate(m_datDebut) & "' AND '" & ConvertDate(m_datFin) & "' AND NoEmploye = " & lvwEmploye.ListItems(iCompteur).Tag & " ORDER BY Date, HeureDébut, HeureFin", g_connData, adOpenDynamic, adLockOptimistic)

            'Vérifie s'il y a des punchs
90          If Not rstPunch.EOF Then
              'Vérifie si le punch out a été fait
95            Do While Not rstPunch.EOF
100             If IsNull(rstPunch.Fields("HeureFin")) Or rstPunch.Fields("HeureFin") = vbNullString Then
105               Call MsgBox("Un punch out n'a pas été fait pour " & lvwEmploye.ListItems(iCompteur).Text & " !", vbOKOnly, "Erreur")

110               bAnnuler = True

115               Exit Do
120             End If

125             Call rstPunch.MoveNext
130           Loop
135         Else
140           Screen.MousePointer = vbDefault

145           Call MsgBox("Il n'y a aucun punch à imprimer pour " & lvwEmploye.ListItems(iCompteur).Text & " !", vbOKOnly, "Erreur")

150           bAnnuler = True
155         End If

160         Call rstPunch.Close

165         If bAnnuler = False Then
170           Call RemplirTableImpressionPunch(iCompteur)

175           Call AjouterNomJour


180           Call AjouterSéparateur

185           Call CalculerTotal

              'Le nom de l'employé
190           DR_FeuilleTemps.Sections("Section4").Controls("lblNom").Caption = lvwEmploye.ListItems(iCompteur).Text

              'La date de début et de fin de la semaine
195           DR_FeuilleTemps.Sections("Section4").Controls("lblDate").Caption = "Semaine du " & GetDateTexte(m_datDebut) & " au " & GetDateTexte(m_datFin)

              'Date d'aujourd'hui
200           DR_FeuilleTemps.Sections("Section3").Controls("lblDatePrint").Caption = ConvertDate(Date)

205           Call rstImpPunch.Open("SELECT * FROM GRB_ImpressionPunch ORDER BY Date, HeureDébut", g_connData, adOpenDynamic, adLockOptimistic)

210           Do While Not rstImpPunch.EOF
215             If Not IsNull(rstImpPunch.Fields("Total")) Then
220               If rstImpPunch.Fields("Total") <> "" Then
225                 dblTotal = dblTotal + CDbl(Left(rstImpPunch.Fields("Total"), 2))
230                 dblTotal = dblTotal + CDbl(CDbl(Right(rstImpPunch.Fields("Total"), 2)) / 60)
235               End If
240             End If

245             Call rstImpPunch.MoveNext
250           Loop

255           Call rstImpPunch.MoveFirst

260           Call rstSommeKM.Open("SELECT SUM(NbreKM) As TotalKM FROM GRB_Punch WHERE Date BETWEEN '" & ConvertDate(m_datDebut) & "' AND '" & ConvertDate(m_datFin) & "' AND NoEmploye = " & lvwEmploye.ListItems(iCompteur).Tag & " AND KM = True", g_connData, adOpenDynamic, adLockOptimistic)
  
265           Set DR_FeuilleTemps.DataSource = rstImpPunch

              'Le total d'heure dans une semaine
270           DR_FeuilleTemps.Sections("Section5").Controls("lblGrandTotal").Caption = Round(dblTotal, 2)

275           If Not IsNull(rstSommeKM.Fields("TotalKM")) Then
280             DR_FeuilleTemps.Sections("Section5").Controls("lblGrandTotalKM").Caption = Round(rstSommeKM.Fields("TotalKM"), 2)
285           Else
290             DR_FeuilleTemps.Sections("Section5").Controls("lblGrandTotalKM").Caption = "0"
295           End If

300           DR_FeuilleTemps.Orientation = rptOrientLandscape

305           Call DR_FeuilleTemps.Show(vbModal)
  
310           Call rstImpPunch.Close
315           Call rstSommeKM.Close
320         End If
325       End If
330     Next

335     Screen.MousePointer = vbDefault

340     Call Unload(Me)

345     Exit Sub

AfficherErreur:

350     woups "frmChoixImpressionFT", "cmdImprimer_Click", Err, Erl
End Sub

Private Sub RemplirListViewEmploye()
  
5       On Error GoTo AfficherErreur

10      Dim rstEmploye As ADODB.Recordset
15      Dim itmEmploye As ListItem
        
20      Set rstEmploye = New ADODB.Recordset
        
25      Call rstEmploye.Open("SELECT NoEmploye, Employe FROM GRB_Employés WHERE Actif = True ORDER BY Employe", g_connData, adOpenForwardOnly, adLockReadOnly)
        
30      Do While Not rstEmploye.EOF
35        Set itmEmploye = lvwEmploye.ListItems.Add
          
40        itmEmploye.Tag = rstEmploye.Fields("NoEmploye")
45        itmEmploye.Text = rstEmploye.Fields("Employe")
            
50        Call rstEmploye.MoveNext
55      Loop
        
60      Call rstEmploye.Close
65      Set rstEmploye = Nothing

70      Exit Sub

AfficherErreur:

75      woups "frmChoixImpressionFT", "RemplirListViewEmploye", Err, Erl
End Sub

Private Sub RemplirTableImpressionPunch(ByVal iIndexListView As Integer)

5       On Error GoTo AfficherErreur

10      Dim rstImpPunch As ADODB.Recordset
15      Dim rstPunch    As ADODB.Recordset
20      Dim rstClient   As ADODB.Recordset
  
25      Set rstPunch = New ADODB.Recordset
30      Set rstImpPunch = New ADODB.Recordset
35      Set rstClient = New ADODB.Recordset
  
40      Call rstPunch.Open("SELECT * FROM GRB_Punch WHERE NoEmploye = " & lvwEmploye.ListItems(iIndexListView).Tag & " AND Date BETWEEN '" & ConvertDate(m_datDebut) & "' AND '" & ConvertDate(m_datFin) & "' ORDER BY Date, HeureDébut", g_connData, adOpenDynamic, adLockOptimistic)
  
45      Call rstImpPunch.Open("SELECT * FROM GRB_ImpressionPunch", g_connData, adOpenDynamic, adLockOptimistic)
  
50      Do While Not rstPunch.EOF
55        Call rstImpPunch.AddNew
    
60        rstImpPunch.Fields("NoProjet") = rstPunch.Fields("NoProjet")
65        rstImpPunch.Fields("Date") = rstPunch.Fields("Date")
    
70        If rstPunch.Fields("ModifDébut") = True Then
75          rstImpPunch.Fields("HeureDébut") = Right$("0" & rstPunch.Fields("HeureDébut"), 5) & "*"
80        Else
85          rstImpPunch.Fields("HeureDébut") = Right$("0" & rstPunch.Fields("HeureDébut"), 5)
90        End If
    
95        If rstPunch.Fields("ModifFin") = True Then
100         rstImpPunch.Fields("HeureFin") = Right$("0" & rstPunch.Fields("HeureFin"), 5) & "*"
105       Else
110         rstImpPunch.Fields("HeureFin") = Right$("0" & rstPunch.Fields("HeureFin"), 5)
115       End If
    
120       If Not IsNull(rstPunch.Fields("NoClient")) And rstPunch.Fields("NoClient") > "0" Then
125         Call rstClient.Open("SELECT NomClient FROM GRB_Client WHERE IDClient = " & rstPunch.Fields("NoClient"), g_connData, adOpenDynamic, adLockOptimistic)
      
130         rstImpPunch.Fields("Client") = rstClient.Fields("NomClient")
      
135         Call rstClient.Close
140       End If
    
145       rstImpPunch.Fields("Commentaire") = rstPunch.Fields("Commentaire")


        '**************************************************************************
        'MODIFIÉ PAR GAÉTAN GINGRAS LE 07 FÉVRIER 2010
        '**************************************************************************
        'ajout
        rstImpPunch.Fields("Type") = rstPunch.Fields("Type")
        '**************************************************************************
        
    
150       rstImpPunch.Fields("NbreKM") = rstPunch.Fields("NbreKM")
    
155       Call rstImpPunch.Update
    
160       Call rstPunch.MoveNext
165     Loop
  
170     Call rstPunch.Close
175     Set rstPunch = Nothing
  
180     Call rstImpPunch.Close
185     Set rstImpPunch = Nothing

190     Set rstClient = Nothing

195     Exit Sub

AfficherErreur:

200     woups "frmChoixImpressionFT", "RemplirTableImpressionPunch", Err, Erl
End Sub

Private Sub AjouterNomJour()

5       On Error GoTo AfficherErreur

10      Dim rstImpPunch As ADODB.Recordset
15      Dim sJour       As String
20      Dim datTemp     As Date
25      Dim sDate       As String
  
        'Ouverture du recordset pour ajouter le nom de la journée
30      Set rstImpPunch = New ADODB.Recordset
        
35      rstImpPunch.CursorLocation = adUseServer
        
40      Call rstImpPunch.Open("SELECT * FROM GRB_ImpressionPunch ORDER BY Date, HeureDébut", g_connData, adOpenDynamic, adLockOptimistic)
      
        'Loop pour ajouter le nom de la journée
45      Do While Not rstImpPunch.EOF
50        sDate = rstImpPunch.Fields("Date")
    
55        datTemp = DateSerial(Left$(sDate, 4), Mid$(sDate, 6, 2), Right$(sDate, 2))
      
60        If sJour <> WeekdayName(Weekday(datTemp)) Then
65          rstImpPunch.Fields("NomJour") = UCase(WeekdayName(Weekday(datTemp)))
70          sJour = WeekdayName(Weekday(datTemp))
75        Else
80          rstImpPunch.Fields("NomJour") = ""
85        End If
      
90        Call rstImpPunch.Update
      
95        Call rstImpPunch.MoveNext
100     Loop
    
105     Call rstImpPunch.Close
110     Set rstImpPunch = Nothing

115     Exit Sub

AfficherErreur:

120     woups "frmChoixImpressionFT", "AjouterNomJour", Err, Erl
End Sub

Private Sub AjouterSéparateur()

5       On Error GoTo AfficherErreur

10      Dim rstImpPunch As ADODB.Recordset
15      Dim iNoRec      As Integer
20      Dim sJour       As String
25      Dim collDate    As Collection
30      Dim sDate       As String
35      Dim iCompteur   As Integer
  
40      Set collDate = New Collection
  
45      Set rstImpPunch = New ADODB.Recordset
  
50      Call rstImpPunch.Open("SELECT * FROM GRB_ImpressionPunch ORDER BY Date, HeureDébut", g_connData, adOpenDynamic, adLockOptimistic)
  
55      iNoRec = 1
  
60      Do While Not rstImpPunch.EOF
65        If sJour <> rstImpPunch.Fields("NomJour") Then
70          If rstImpPunch.Fields("NomJour") <> vbNullString Then
75            If iNoRec > 1 Then
80              sDate = rstImpPunch.Fields("Date")
        
85              Call collDate.Add(sDate)
90            End If
95          End If
      
100         If rstImpPunch.Fields("NomJour") <> vbNullString Then
105           sJour = rstImpPunch.Fields("NomJour")
110         End If
115       End If
    
120       iNoRec = iNoRec + 1
    
125       Call rstImpPunch.MoveNext
130     Loop
  
135     For iCompteur = 1 To collDate.count
140       Call rstImpPunch.AddNew
    
145       rstImpPunch.Fields("Date") = collDate(iCompteur)
150       rstImpPunch.Fields("NomJour") = vbNullString
155       rstImpPunch.Fields("NoProjet") = " "
160       rstImpPunch.Fields("Client") = vbNullString
165       rstImpPunch.Fields("Commentaire") = vbNullString
170       rstImpPunch.Fields("HeureDébut") = " "
175       rstImpPunch.Fields("HeureFin") = vbNullString
    
180       Call rstImpPunch.Update
185     Next
  
190     Call rstImpPunch.Close
195     Set rstImpPunch = Nothing

200     Exit Sub

AfficherErreur:

205     woups "frmChoixImpressionFT", "AjouterSéparateur", Err, Erl
End Sub

Private Sub CalculerTotal()

5       On Error GoTo AfficherErreur

        'Calcul le temps entre chaque punch in et punch out
10      Dim rstImpPunch As ADODB.Recordset
15      Dim datDébut    As Date
20      Dim datFin      As Date
25      Dim datTotal    As Date
30      Dim sDate       As String
35      Dim sDébut      As String
40      Dim sFin        As String
  
        'Ouverture de tous les punchs à imprimer
45      Set rstImpPunch = New ADODB.Recordset

50      rstImpPunch.CursorLocation = adUseServer
        
55      Call rstImpPunch.Open("SELECT * FROM GRB_ImpressionPunch", g_connData, adOpenDynamic, adLockOptimistic)
  
        'Tant que ce n'est pas la fin des enregistrements
60      Do While Not rstImpPunch.EOF
          'Si ce n'est pas un séparateur
65        If Trim(rstImpPunch.Fields("HeureDébut")) <> vbNullString Then
70          sDate = rstImpPunch.Fields("Date")
75          sDébut = Left(rstImpPunch.Fields("HeureDébut"), 5)
80          sFin = Left(rstImpPunch.Fields("HeureFin"), 5)
      
85          datDébut = DateSerial(Left$(sDate, 4), Mid$(sDate, 6, 2), Right$(sDate, 2)) + TimeSerial(Left$(sDébut, 2), Mid$(sDébut, 4, 2), 0)
90          datFin = DateSerial(Left$(sDate, 4), Mid$(sDate, 6, 2), Right$(sDate, 2)) + TimeSerial(Left$(sFin, 2), Mid$(sFin, 4, 2), 0)
      
95          If Hour(datDébut) = 0 And Minute(datDébut) = 0 And Second(datDébut) = 0 And _
               Hour(datFin) = 0 And Minute(datFin) = 0 And Second(datFin) = 0 And _
               DateDiff("d", datDébut, datFin) = 1 Then
100           rstImpPunch.Fields("Total") = "24:00"
105         Else
110           datTotal = datFin - datDébut
      
115           rstImpPunch.Fields("Total") = Right$("0" & Hour(datTotal), 2) & ":" & Right$("0" & Minute(datTotal), 2)
120         End If

125         Call rstImpPunch.Update
130       End If
    
135       Call rstImpPunch.MoveNext
140     Loop

145     Exit Sub

AfficherErreur:

150     woups "frmChoixImpressionFT", "CalculerTotal", Err, Erl
End Sub

Private Sub cmdTous_Click()
  
5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer

15      For iCompteur = 1 To lvwEmploye.ListItems.count
20        lvwEmploye.ListItems(iCompteur).Checked = True
25      Next

30      Exit Sub

AfficherErreur:

35      woups "frmChoixImpressionFT", "cmdTous_Click", Err, Erl
End Sub

Private Sub cmdAucun_Click()
  
5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer
  
15      For iCompteur = 1 To lvwEmploye.ListItems.count
20        lvwEmploye.ListItems(iCompteur).Checked = False
25      Next

30      Exit Sub

AfficherErreur:

35      woups "frmChoixImpressionFT", "cmdAucun_Click", Err, Erl
End Sub
