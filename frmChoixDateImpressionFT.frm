VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmChoixDateImpressionFT 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportation feuille de temps"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3900
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChoixDateImpressionFT.frx":0000
   ScaleHeight     =   3285
   ScaleWidth      =   3900
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.MonthView mvwDate 
      Height          =   2370
      Left            =   1200
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   0
      Appearance      =   1
      ShowToday       =   0   'False
      StartOfWeek     =   90243073
      CurrentDate     =   37735
   End
   Begin VB.CommandButton cmdDateFin 
      Caption         =   "..."
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton cmdDateDebut 
      Caption         =   "..."
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdExporter 
      Caption         =   "Exporter"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSMask.MaskEdBox mskDateDebut 
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskDateFin 
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date début :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date fin :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "AA-MM-JJ"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "frmChoixDateImpressionFT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enumDate
  AUCUNE = 0
  DEBUT = 1
  Fin = 2
End Enum

Private m_eDate  As enumDate
Private m_xlsApp As Excel.Application

Private Sub cmdAnnuler_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmChoixDateImpressionFT", "cmdAnnuler_Click", Err, Erl
End Sub

Private Sub cmdexporter_Click()

5       On Error GoTo AfficherErreur

10      If Len(mskDateDebut.Text) = 8 Then
15        Call mskDateDebut_LostFocus
20      End If

25      If Len(mskDateFin.Text) = 8 Then
30        Call mskDateFin_LostFocus
35      End If

40      If ValiderDate(mskDateDebut.Text) = False Then
45        Call MsgBox("Date de début invalide!", vbOKOnly, "Erreur")

50        Exit Sub
55      End If

60      If ValiderDate(mskDateFin.Text) = False Then
65        Call MsgBox("Date de fin invalide!", vbOKOnly, "Erreur")

70        Exit Sub
75      End If

80      If mskDateFin.Text < mskDateDebut.Text Then
85        Call MsgBox("La date de fin doit être plus grande que la date de début!", vbOKOnly, "Erreur")

90        Exit Sub
95      End If

100     Call ExporterPunch

105     Call Unload(Me)

110     Exit Sub

AfficherErreur:

115     woups "frmChoixDateImpressionFT", "cmdExporter_Click", Err, Erl
End Sub

Private Sub Form_Click()

5       On Error GoTo AfficherErreur

10      mvwDate.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmChoixDateImpressionFT", "Form_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      m_eDate = AUCUNE

15      Exit Sub

AfficherErreur:

20      woups "frmChoixDateImpressionFT", "Form_Load", Err, Erl
End Sub

Private Sub mvwDate_LostFocus()

5       On Error GoTo AfficherErreur

10      mvwDate.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmChoixDateImpressionFT", "mvwDate_LostFocus", Err, Erl
End Sub

Private Sub mvwDate_DateClick(ByVal DateClicked As Date)

5       On Error GoTo AfficherErreur

10      Select Case m_eDate
          Case DEBUT: mskDateDebut.Text = ConvertDate(DateClicked)
          Case Fin:   mskDateFin.Text = ConvertDate(DateClicked)
15      End Select
  
20      m_eDate = AUCUNE
  
        'Enlever le calendrier
25      mvwDate.Visible = False

30      Exit Sub

AfficherErreur:

35      woups "frmChoixDateImpressionFT", "mvwDate_DateClick", Err, Erl
End Sub

Private Sub mskDateDebut_GotFocus()

5       On Error GoTo AfficherErreur
        
        'Met l'année sur 2 chiffres
10      If Len(mskDateDebut.Text) = 10 Then
15        mskDateDebut.Text = Right$(mskDateDebut.Text, 8)
20      End If
  
25      mskDateDebut.mask = "##-##-##"

30      Exit Sub

AfficherErreur:

35      woups "frmChoixDateImpressionFT", "mskDateDebut_GotFocus", Err, Erl
End Sub

Private Sub mskDateFin_GotFocus()

5       On Error GoTo AfficherErreur
        
        'Met l'année sur 2 chiffres
10      If Len(mskDateFin.Text) = 10 Then
15        mskDateFin.Text = Right$(mskDateFin.Text, 8)
20      End If
  
25      mskDateFin.mask = "##-##-##"

30      Exit Sub

AfficherErreur:

35      woups "frmChoixDateImpressionFT", "mskDateFin_GotFocus", Err, Erl
End Sub

Private Sub mskDateDebut_LostFocus()

5       On Error GoTo AfficherErreur
        
        'Enlève le mask
10      mskDateDebut.mask = vbNullString
  
        'Vide le champs si l'utilisateur n'a rien écrit
15      If mskDateDebut.Text = "__-__-__" Then
20        mskDateDebut.Text = vbNullString
25      Else
          'Remet l'année sur 8 chiffres
30        If Len(mskDateDebut.Text) = 8 Then
35          If IsDate(mskDateDebut.Text) Then
40            mskDateDebut.Text = Year(DateSerial(Left$(mskDateDebut.Text, 2), Mid$(mskDateDebut.Text, 4, 2), Right$(mskDateDebut.Text, 2))) & Mid$(mskDateDebut.Text, 3, 8)
45          End If
50        End If
55      End If

60      Exit Sub

AfficherErreur:

65      woups "frmChoixDateImpressionFT", "mskDateDebut_LostFocus", Err, Erl
End Sub

Private Sub mskDateFin_LostFocus()

5       On Error GoTo AfficherErreur
        
        'Enlève le mask
10      mskDateFin.mask = vbNullString
  
        'Vide le champs si l'utilisateur n'a rien écrit
15      If mskDateFin.Text = "__-__-__" Then
20        mskDateFin.Text = vbNullString
25      Else
          'Remet l'année sur 8 chiffres
30        If Len(mskDateFin.Text) = 8 Then
35          If IsDate(mskDateFin.Text) Then
40            mskDateFin.Text = Year(DateSerial(Left$(mskDateFin.Text, 2), Mid$(mskDateFin.Text, 4, 2), Right$(mskDateFin.Text, 2))) & Mid$(mskDateFin.Text, 3, 8)
45          End If
50        End If
55      End If

60      Exit Sub

AfficherErreur:

65      woups "frmChoixDateImpressionFT", "mskDateFin_LostFocus", Err, Erl
End Sub

Private Sub cmdDateDebut_Click()

5       On Error GoTo AfficherErreur
        'Ouverture du calendrier
  
        'Si il y a une date valide, la date par défaut est celle-ci, sinon c'est la date
        'd'aujourd'hui
10      If Trim$(mskDateDebut.Text) <> vbNullString Then
15        If ValiderDate(mskDateDebut.Text) = True Then
20          mvwDate.Value = mskDateDebut.Text
25        Else
30          mvwDate.Value = Date
35        End If
40      Else
45        mvwDate.Value = Date
50      End If
  
55      m_eDate = DEBUT
  
60      mvwDate.Visible = True
  
65      Call mvwDate.SetFocus

70      Exit Sub

AfficherErreur:

75      woups "frmChoixDateImpressionFT", "cmdDateDebut_Click", Err, Erl
End Sub

Private Sub cmdDateFin_Click()

5       On Error GoTo AfficherErreur
        'Ouverture du calendrier
  
        'Si il y a une date valide, la date par défaut est celle-ci, sinon c'est la date
        'd'aujourd'hui
10      If Trim$(mskDateFin.Text) <> vbNullString Then
15        If ValiderDate(mskDateFin.Text) = True Then
20          mvwDate.Value = mskDateFin.Text
25        Else
30          mvwDate.Value = Date
35        End If
40      Else
45        mvwDate.Value = Date
50      End If
  
55      m_eDate = Fin
  
60      mvwDate.Visible = True
  
65      Call mvwDate.SetFocus

70      Exit Sub

AfficherErreur:

75      woups "frmChoixDateImpressionFT", "cmdDateFin_Click", Err, Erl
End Sub

Private Function ValiderDate(ByVal sDate As String) As Boolean

5       On Error GoTo AfficherErreur

        'Validation des dates
10      If Not IsDate(sDate) Then
15        ValiderDate = False
20      Else
25        ValiderDate = True
30      End If

35      Exit Function

AfficherErreur:

40      woups "frmChoixDateImpressionFT", "ValiderDate", Err, Erl
End Function

Private Sub ExporterPunch()
  
5       On Error GoTo AfficherErreur
  
        'Impression de la feuille de la liste des pièces
10      Dim rstEmployes   As ADODB.Recordset
15      Dim rstProjets    As ADODB.Recordset
20      Dim xlsWorkBook   As Excel.Workbook
25      Dim iNbreEmployes As Integer
30      Dim iNbreProjets  As Integer
35      Dim iNbrePages    As Integer
40      Dim iPage         As Integer
45      Dim sDateDebut    As String
50      Dim sDateFin      As String
  
55      Screen.MousePointer = vbHourglass

60      Set rstEmployes = New ADODB.Recordset
65      Set rstProjets = New ADODB.Recordset

70      rstEmployes.CursorLocation = adUseClient
75      rstProjets.CursorLocation = adUseClient
  
80      sDateDebut = mskDateDebut.Text
85      sDateFin = mskDateFin.Text
  
90      Call rstProjets.Open("SELECT DISTINCT NoProjet, RIGHT(NoProjet, 9) FROM GRB_Punch WHERE Date BETWEEN '" & sDateDebut & "' AND '" & sDateFin & "' ORDER BY RIGHT(NoProjet, 9)", g_connData, adOpenDynamic, adLockOptimistic)

95      Call rstEmployes.Open("SELECT DISTINCT Employe FROM GRB_Punch INNER JOIN GRB_Employés ON GRB_Employés.NoEmploye = GRB_Punch.NoEmploye WHERE Date BETWEEN '" & sDateDebut & "' AND '" & sDateFin & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

100     iNbreEmployes = rstEmployes.RecordCount
105     iNbreProjets = rstProjets.RecordCount

110     If iNbreProjets > 26 Then
115       iNbrePages = Int(iNbreProjets / 26)
    
120       If iNbrePages * 26 < iNbreProjets Then
125         iNbrePages = iNbrePages + 1
130       End If
135     Else
140       iNbrePages = 1
145     End If
        
150     Set m_xlsApp = New Excel.Application

155     Set xlsWorkBook = m_xlsApp.Workbooks.Add
      
160     For iPage = 1 To iNbrePages
165       Call CreerTableau(rstProjets, rstEmployes, (iPage * 43) - 42, sDateDebut, sDateFin)
170     Next
  
175     Call rstProjets.Close
180     Call rstEmployes.Close
  
185     Set rstProjets = Nothing
190     Set rstEmployes = Nothing
  
195     Call TransfererValeurs(sDateDebut, sDateFin)
  
200     Call RemplirValeurs(iNbrePages)
    
205     m_xlsApp.ActiveSheet.PageSetup.LeftMargin = m_xlsApp.Application.InchesToPoints(0)
210     m_xlsApp.ActiveSheet.PageSetup.RightMargin = m_xlsApp.Application.InchesToPoints(0)
215     m_xlsApp.ActiveSheet.PageSetup.TopMargin = m_xlsApp.Application.InchesToPoints(0)
220     m_xlsApp.ActiveSheet.PageSetup.BottomMargin = m_xlsApp.Application.InchesToPoints(0)
225     m_xlsApp.ActiveSheet.PageSetup.HeaderMargin = m_xlsApp.Application.InchesToPoints(0)
230     m_xlsApp.ActiveSheet.PageSetup.FooterMargin = m_xlsApp.Application.InchesToPoints(0)
235     m_xlsApp.ActiveSheet.PageSetup.CenterHorizontally = True
240     m_xlsApp.ActiveSheet.PageSetup.CenterVertically = False
245     m_xlsApp.ActiveSheet.PageSetup.Orientation = xlLandscape
250     m_xlsApp.ActiveSheet.PageSetup.PaperSize = xlPaperLegal
  
255     Screen.MousePointer = vbDefault

260     m_xlsApp.Visible = True

265     Exit Sub

AfficherErreur:

270     woups "frmChoixDateImpressionFT", "ExporterPunch", Err, Erl
End Sub

Private Sub CreerTableau(ByRef rstProjets As ADODB.Recordset, ByRef rstEmployes As ADODB.Recordset, ByVal iDebut As Integer, ByVal sDateDebut As String, ByVal sDateFin As String)
  
5       On Error GoTo AfficherErreur
  
10      Dim iNbreProjets   As Integer
15      Dim iNbreEmployes  As Integer
20      Dim iCompteur      As Integer
25      Dim sLettre        As String
  
30      iNbreProjets = rstProjets.RecordCount
35      iNbreEmployes = rstEmployes.RecordCount
    
40      m_xlsApp.Cells(iDebut, 1) = "DU " & UCase(GetDateTexte(sDateDebut)) & " AU " & UCase(GetDateTexte(sDateFin))
45      m_xlsApp.Range("A" & iDebut).Font.Bold = True
50      m_xlsApp.Range("A" & iDebut).Font.Underline = xlUnderlineStyleSingle
55      m_xlsApp.Range("A" & iDebut).HorizontalAlignment = xlCenter
60      m_xlsApp.Range("A" & iDebut).Font.SIZE = 18

65      Call m_xlsApp.Range("A" & iDebut, "AB" & iDebut).Merge

70      m_xlsApp.Columns("A:A").ColumnWidth = 21
75      m_xlsApp.Columns("B:AB").ColumnWidth = 5
80      m_xlsApp.Columns("AB:AB").ColumnWidth = 6.29
    
85      m_xlsApp.Range("B" & iDebut + 3, "AB" & iDebut + 3).HorizontalAlignment = xlCenter
90      m_xlsApp.Range("B" & iDebut + 3, "AB" & iDebut + 3).VerticalAlignment = xlCenter
95      m_xlsApp.Range("B" & iDebut + 3, "AB" & iDebut + 3).Orientation = 90
  
100     For iCompteur = 2 To 27
105       If rstProjets.EOF = True Then
110         Exit For
115       Else
120         m_xlsApp.Cells(iDebut + 3, iCompteur) = rstProjets.Fields("NoProjet")
      
125         Call rstProjets.MoveNext
130       End If
135     Next
  
140     m_xlsApp.Cells(iDebut + 3, 28) = "TOTAL"
  
145     Call rstEmployes.MoveFirst
  
150     For iCompteur = 4 To iNbreEmployes + 3
155       m_xlsApp.Cells(iDebut + iCompteur, 1) = rstEmployes.Fields("Employe")
    
160       Call rstEmployes.MoveNext
165     Next
  
170     m_xlsApp.Cells(iDebut + iCompteur, 1) = "TOTAL"

        'Ajout des lignes
175     m_xlsApp.Range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeLeft).LineStyle = xlContinuous
180     m_xlsApp.Range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeLeft).Weight = xlMedium
185     m_xlsApp.Range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
190     m_xlsApp.Range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeTop).LineStyle = xlContinuous
195     m_xlsApp.Range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeTop).Weight = xlMedium
200     m_xlsApp.Range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeTop).ColorIndex = xlAutomatic
    
205     m_xlsApp.Range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeBottom).LineStyle = xlContinuous
210     m_xlsApp.Range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeBottom).Weight = xlMedium
215     m_xlsApp.Range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    
220     m_xlsApp.Range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeRight).LineStyle = xlContinuous
225     m_xlsApp.Range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeRight).Weight = xlMedium
230     m_xlsApp.Range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeRight).ColorIndex = xlAutomatic
    
235     m_xlsApp.Range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlInsideVertical).LineStyle = xlContinuous
240     m_xlsApp.Range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlInsideVertical).Weight = xlThin
245     m_xlsApp.Range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlInsideVertical).ColorIndex = xlAutomatic
    
250     m_xlsApp.Range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlInsideHorizontal).LineStyle = xlContinuous
255     m_xlsApp.Range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlInsideHorizontal).Weight = xlThin
260     m_xlsApp.Range("A" & iDebut + 3 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
  
265     m_xlsApp.Range("B" & iDebut + 3 & ":AA" & iDebut + iNbreEmployes + 4).Borders(xlEdgeLeft).LineStyle = xlContinuous
270     m_xlsApp.Range("B" & iDebut + 3 & ":AA" & iDebut + iNbreEmployes + 4).Borders(xlEdgeLeft).Weight = xlMedium
275     m_xlsApp.Range("B" & iDebut + 3 & ":AA" & iDebut + iNbreEmployes + 4).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
  
280     m_xlsApp.Range("B" & iDebut + 3 & ":AA" & iDebut + iNbreEmployes + 4).Borders(xlEdgeRight).LineStyle = xlContinuous
285     m_xlsApp.Range("B" & iDebut + 3 & ":AA" & iDebut + iNbreEmployes + 4).Borders(xlEdgeRight).Weight = xlMedium
290     m_xlsApp.Range("B" & iDebut + 3 & ":AA" & iDebut + iNbreEmployes + 4).Borders(xlEdgeRight).ColorIndex = xlAutomatic
    
295     m_xlsApp.Range("A" & iDebut + 3 & ":AB" & iDebut + 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
300     m_xlsApp.Range("A" & iDebut + 3 & ":AB" & iDebut + 3).Borders(xlEdgeBottom).Weight = xlMedium
305     m_xlsApp.Range("A" & iDebut + 3 & ":AB" & iDebut + 3).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    
310     m_xlsApp.Range("A" & iDebut + iNbreEmployes + 4 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeTop).LineStyle = xlContinuous
315     m_xlsApp.Range("A" & iDebut + iNbreEmployes + 4 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeTop).Weight = xlMedium
320     m_xlsApp.Range("A" & iDebut + iNbreEmployes + 4 & ":AB" & iDebut + iNbreEmployes + 4).Borders(xlEdgeTop).ColorIndex = xlAutomatic
    
325     For iCompteur = iDebut + 4 To iNbreEmployes + iDebut + 4
330       m_xlsApp.Range("AB" & iCompteur) = "=SUM(B" & iCompteur & ":AA" & iCompteur & ")"
335     Next
  
340     For iCompteur = 2 To 27
345       Select Case iCompteur
            Case 2:  sLettre = "B"
            Case 3:  sLettre = "C"
            Case 4:  sLettre = "D"
            Case 5:  sLettre = "E"
            Case 6:  sLettre = "F"
            Case 7:  sLettre = "G"
            Case 8:  sLettre = "H"
            Case 9:  sLettre = "I"
            Case 10: sLettre = "J"
            Case 11: sLettre = "K"
            Case 12: sLettre = "L"
            Case 13: sLettre = "M"
            Case 14: sLettre = "N"
            Case 15: sLettre = "O"
            Case 16: sLettre = "P"
            Case 17: sLettre = "Q"
            Case 18: sLettre = "R"
            Case 19: sLettre = "S"
            Case 20: sLettre = "T"
            Case 21: sLettre = "U"
            Case 22: sLettre = "V"
            Case 23: sLettre = "W"
            Case 24: sLettre = "X"
            Case 25: sLettre = "Y"
            Case 26: sLettre = "Z"
            Case 27: sLettre = "AA"
          End Select
    
350       m_xlsApp.Range(sLettre & iDebut + iNbreEmployes + 4) = "=SUM(" & sLettre & (iDebut + 4) & ":" & sLettre & (iDebut + iNbreEmployes + 3) & ")"
355     Next

360     Exit Sub

AfficherErreur:

365     woups "frmChoixDateImpressionFT", "CreerTableau", Err, Erl
End Sub

Private Sub TransfererValeurs(ByVal sDateDebut As String, ByVal sDateFin As String)
  
5       On Error GoTo AfficherErreur

10      Dim rstSource      As ADODB.Recordset
15      Dim rstDestination As ADODB.Recordset
20      Dim sDate          As String
25      Dim sHeure         As String
30      Dim sMinute        As String
35      Dim dblResult      As Double
40      Dim datTemp        As Date
  
45      Set rstSource = New ADODB.Recordset
  
50      Call rstSource.Open("SELECT * FROM GRB_Punch WHERE Date BETWEEN '" & sDateDebut & "' AND '" & sDateFin & "' AND HeureFin Is Not Null", g_connData, adOpenForwardOnly, adLockReadOnly)
  
55      If Not rstSource.EOF Then
60        Call g_connData.Execute("DELETE * FROM GRB_PunchExcel")
  
65        Set rstDestination = New ADODB.Recordset
    
70        Call rstDestination.Open("SELECT * FROM GRB_PunchExcel", g_connData, adOpenDynamic, adLockOptimistic)
  
75        Do While Not rstSource.EOF
80          Call rstDestination.AddNew
      
85          rstDestination.Fields("NoEmploye") = rstSource.Fields("NoEmploye")
90          rstDestination.Fields("NoProjet") = rstSource.Fields("NoProjet")
      
95          sHeure = Left(rstSource.Fields("HeureDébut"), 2)
100         sMinute = Right(rstSource.Fields("HeureDébut"), 2)
      
105         dblResult = CInt(sMinute) / 60
      
110         If dblResult <> 0 Then
115           rstDestination.Fields("HeureDébut") = sHeure & "," & Right$(dblResult, Len(CStr(dblResult)) - InStr(1, dblResult, ","))
120         Else
125           rstDestination.Fields("HeureDébut") = CInt(sHeure)
130         End If
      
135         sHeure = Left(rstSource.Fields("HeureFin"), 2)
140         sMinute = Right(rstSource.Fields("HeureFin"), 2)
      
145         dblResult = CInt(sMinute) / 60
      
150         If dblResult <> 0 Then
155           rstDestination.Fields("HeureFin") = CInt(sHeure) & "," & CInt(Right$(dblResult, Len(CStr(dblResult)) - InStr(1, dblResult, ",")))
160         Else
165           rstDestination.Fields("HeureFin") = sHeure
170         End If
      
175         Call rstDestination.Update
      
180         Call rstSource.MoveNext
185       Loop
  
190       Call rstDestination.Close
195       Set rstDestination = Nothing
200     End If
  
205     Call rstSource.Close
210     Set rstSource = Nothing

215     Exit Sub

AfficherErreur:

220     woups "frmChoixDateImpressionFT", "TransfererValeurs", Err, Erl
End Sub

Private Sub RemplirValeurs(ByVal iNbrePages As Integer)
  
5       On Error GoTo AfficherErreur

10      Dim rstPunch      As ADODB.Recordset
15      Dim sNom          As String
20      Dim iIndexNom     As Integer
25      Dim iCompteur     As Integer
30      Dim iPage         As Integer
35      Dim iPageRendu    As Integer
40      Dim iNoRendu      As Integer
45      Dim iDebutPage    As Integer
50      Dim iIndexProjet  As Integer
55      Dim bProjetTrouve As Boolean

60      Set rstPunch = New ADODB.Recordset

65      Call rstPunch.Open("SELECT Employe, NoProjet, SUM(HeureFin - HeureDébut) As Total FROM GRB_PunchExcel INNER JOIN GRB_Employés ON GRB_PunchExcel.NoEmploye = GRB_Employés.NoEmploye GROUP BY Employe, NoProjet ORDER BY Employe, RIGHT(NoProjet, 9)", g_connData, adOpenForwardOnly, adLockReadOnly)
  
70      If Not rstPunch.EOF Then
75        iIndexNom = 4
    
80        sNom = rstPunch.Fields("Employe")
      
85        iPageRendu = 1
90        iNoRendu = 2
  
95        Do While Not rstPunch.EOF
100         If sNom <> rstPunch.Fields("Employe") Then
105           sNom = rstPunch.Fields("Employe")
        
110           iIndexNom = iIndexNom + 1
        
115           iPageRendu = 1
120           iNoRendu = 2
125         Else
130           If iNoRendu > 2 Then
135             iNoRendu = iNoRendu - 1
140           Else
145             If iPageRendu > 1 Then
150               iPageRendu = iPageRendu - 1
155               iNoRendu = 27
160             End If
165           End If
170         End If
      
175         bProjetTrouve = False
           
180         For iPage = iPageRendu To iNbrePages
185           If iPageRendu <> iPage Then
190             iNoRendu = 2
195           End If

200           iDebutPage = (iPage * 43) - 42
      
205           For iCompteur = iNoRendu To 27
210             If m_xlsApp.Cells(iDebutPage + 3, iCompteur) = rstPunch.Fields("NoProjet") Then
215               iIndexProjet = iCompteur
            
220               iPageRendu = iPage
225               iNoRendu = iCompteur
            
230               bProjetTrouve = True
          
235               Exit For
240             End If
245           Next
        
250           If bProjetTrouve = True Then
255             Exit For
260           End If
265         Next
            
270         If bProjetTrouve = False Then
275           Call MsgBox("Le # " & rstPunch.Fields("NoProjet") & " n'a pas pu être trouvé pour l'employé " & sNom & "." & vbNewLine & _
                          "Son temps de " & rstPunch.Fields("Total") & " heures sera ajouté à cet endroit : " & vbNewLine & _
                          "Page   : " & iPageRendu & vbNewLine & _
                          "Rangée : " & iIndexProjet)
                          
280         End If
            
285         m_xlsApp.Cells(iDebutPage + iIndexNom, iIndexProjet) = rstPunch.Fields("Total")
      
290         Call rstPunch.MoveNext
295       Loop
300     End If
    
305     Call rstPunch.Close
310     Set rstPunch = Nothing

315     Exit Sub

AfficherErreur:

320     woups "frmChoixDateImpressionFT", "RemplirValeurs", Err, Erl
End Sub
