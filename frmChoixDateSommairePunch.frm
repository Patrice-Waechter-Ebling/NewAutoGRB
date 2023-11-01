VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChoixDateSommairePunch 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sommaire des projets"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChoixDateSommairePunch.frx":0000
   ScaleHeight     =   3690
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.MonthView mvwDate 
      Height          =   2370
      Left            =   5040
      TabIndex        =   14
      Top             =   720
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Date"
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
      Height          =   2055
      Left            =   4800
      TabIndex        =   6
      Top             =   960
      Width           =   2895
      Begin VB.CommandButton cmdDateFin 
         Caption         =   "..."
         Height          =   255
         Left            =   2280
         TabIndex        =   10
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton cmdDateDebut 
         Caption         =   "..."
         Height          =   255
         Left            =   2280
         TabIndex        =   7
         Top             =   600
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskDateDebut 
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDateFin 
         Height          =   255
         Left            =   960
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Début :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fin :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "AA-MM-JJ"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fraFamille 
      BackColor       =   &H00000000&
      Caption         =   "Famille"
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
      Height          =   2055
      Left            =   2400
      TabIndex        =   5
      Top             =   960
      Width           =   2295
      Begin MSComctlLib.ListView lvwFamille 
         Height          =   1575
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   2778
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Text            =   "Famille"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Projets"
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
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2175
      Begin VB.OptionButton optProjets 
         BackColor       =   &H00000000&
         Caption         =   "Sommaire des heures"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   2025
      End
      Begin VB.OptionButton optProjets 
         BackColor       =   &H00000000&
         Caption         =   "Projets GRB seulement"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   2025
      End
      Begin VB.OptionButton optProjets 
         BackColor       =   &H00000000&
         Caption         =   "Tous les projets"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   2025
      End
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimer 
      Caption         =   "Imprimer"
      Default         =   -1  'True
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
   End
End
Attribute VB_Name = "frmChoixDateSommairePunch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const I_OPT_TOUS_LES_PROJETS As Integer = 0
Private Const I_OPT_PROJETS_GRB      As Integer = 1
Private Const I_OPT_SOMMAIRE_HEURES  As Integer = 2


Private Enum enumDate
  AUCUNE = 0
  DEBUT = 1
  Fin = 2
End Enum

Private m_eDate As enumDate

Private Sub cmdAnnuler_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmChoixDateSommairePunch", "cmdAnnuler_Click", Err, Erl
End Sub

Private Sub cmdImprimer_Click()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer
15      Dim bChecked  As Boolean

20      If ValiderDate(mskDateDebut.Text) = False Then
25        Call MsgBox("Date de début invalide!", vbOKOnly, "Erreur")

30        Exit Sub
35      End If

40      If ValiderDate(mskDateFin.Text) = False Then
45        Call MsgBox("Date de fin invalide!", vbOKOnly, "Erreur")

50        Exit Sub
55      End If

60      If mskDateFin.Text < mskDateDebut.Text Then
65        Call MsgBox("La date de fin doit être plus grande que la date de début!", vbOKOnly, "Erreur")

70        Exit Sub
75      End If

        'Si "Projets GRB seulement" est choisi
80      If optProjets(I_OPT_PROJETS_GRB).Value = True Then
          'Si aucune famille est sélectionnée
85        For iCompteur = 1 To lvwFamille.ListItems.count
90          If lvwFamille.ListItems(iCompteur).Checked = True Then
95            bChecked = True

100           Exit For
105         End If
110       Next
          
115       If bChecked = False Then
120         Call MsgBox("Vous devez choisir au moins une famille d'employés!", vbOKOnly, "Erreur")
          
125         Exit Sub
130       End If
135     End If

140     Screen.MousePointer = vbHourglass

145     If optProjets(I_OPT_TOUS_LES_PROJETS).Value = True Then
150       Call ImprimerPunchGeneral
155     Else
160       If optProjets(I_OPT_PROJETS_GRB).Value = True Then
165         Call ImprimerPunchGRB
170       Else
175         Call ImprimerSommaireHeures
180       End If
185     End If

190     Screen.MousePointer = vbDefault

195     Call Unload(Me)

200     Exit Sub

AfficherErreur:

205     woups "frmChoixDateSommairePunch", "cmdExporter_Click", Err, Erl
End Sub

Private Sub Form_Click()

5       On Error GoTo AfficherErreur

10      mvwDate.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmChoixDateSommairePunch", "Form_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      m_eDate = AUCUNE

15      Call RemplirListViewFamille

20      optProjets(I_OPT_TOUS_LES_PROJETS).Value = vbChecked

25      Exit Sub

AfficherErreur:

30      woups "frmChoixDateSommairePunch", "Form_Load", Err, Erl
End Sub

Private Sub RemplirListViewFamille()

5       On Error GoTo AfficherErreur

10      Dim rstFamille As ADODB.Recordset
15      Dim itmFamille As ListItem

20      Set rstFamille = New ADODB.Recordset
  
25      Call rstFamille.Open("SELECT * FROM GRB_Famille ORDER BY Famille", g_connData, adOpenDynamic, adLockOptimistic)
  
30      Do While Not rstFamille.EOF
35        Set itmFamille = lvwFamille.ListItems.Add
    
40        itmFamille.Text = rstFamille.Fields("Famille")
45        itmFamille.Tag = rstFamille.Fields("IDFamille")
    
50        Call rstFamille.MoveNext
55      Loop
    
60      Call rstFamille.Close
65      Set rstFamille = Nothing

70      Exit Sub

AfficherErreur:

75      woups "frmChoixDateSommairePunch", "RemplirListViewFamille", Err, Erl
End Sub

Private Sub mvwDate_LostFocus()

5       On Error GoTo AfficherErreur

10      mvwDate.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmChoixDateSommairePunch", "mvwDate_LostFocus", Err, Erl
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

35      woups "frmChoixDateSommairePunch", "mvwDate_DateClick", Err, Erl
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

35      woups "frmChoixDateSommairePunch", "mskDateDebut_GotFocus", Err, Erl
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

35      woups "frmChoixDateSommairePunch", "mskDateFin_GotFocus", Err, Erl
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

65      woups "frmChoixDateSommairePunch", "mskDateDebut_LostFocus", Err, Erl
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

65      woups "frmChoixDateSommairePunch", "mskDateFin_LostFocus", Err, Erl
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

75      woups "frmChoixDateSommairePunch", "cmdDateDebut_Click", Err, Erl
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

75      woups "frmChoixDateSommairePunch", "cmdDateFin_Click", Err, Erl
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

40      woups "frmChoixDateSommairePunch", "ValiderDate", Err, Erl
End Function

Private Sub ImprimerPunchGeneral()

5       On Error GoTo AfficherErreur

10      Dim rstProjets  As ADODB.Recordset
15      Dim rstHeures   As ADODB.Recordset
20      Dim rstSommaire As ADODB.Recordset
25      Dim datDebut    As Date
30      Dim datFin      As Date
35      Dim dblTotal    As Double
40      Dim dblSecondes As Double
45      Dim dblHeures   As Double

50      Call g_connData.Execute("DELETE * FROM GRB_ImpressionSommairePunchGeneral")

55      Set rstSommaire = New ADODB.Recordset
60      Set rstProjets = New ADODB.Recordset
65      Set rstHeures = New ADODB.Recordset

70      Call rstSommaire.Open("SELECT * FROM GRB_ImpressionSommairePunchGeneral", g_connData, adOpenDynamic, adLockOptimistic)

75      Call rstProjets.Open("SELECT DISTINCT NoProjet FROM GRB_Punch WHERE Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

80      Do While Not rstProjets.EOF
85        Call rstSommaire.AddNew
          
90        rstSommaire.Fields("NoProjet") = rstProjets.Fields("NoProjet")
          
95        Call rstHeures.Open("SELECT HeureDébut, HeureFin FROM GRB_Punch WHERE NoProjet = '" & rstProjets.Fields("NoProjet") & "' AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin is not Null", g_connData, adOpenDynamic, adLockOptimistic)
          
100       dblTotal = 0
          
105       Do While Not rstHeures.EOF
110         datDebut = rstHeures.Fields("HeureDébut")

115         If rstHeures.Fields("HeureFin") = "24:00" Then
120           datFin = TimeSerial(24, 0, 0)
125         Else
130           datFin = rstHeures.Fields("HeureFin")
135         End If
            
140         dblSecondes = DateDiff("s", datDebut, datFin)

145         dblHeures = Int(dblSecondes / 3600)
            
150         dblSecondes = dblSecondes - (3600 * dblHeures)
            
155         dblTotal = dblTotal + dblHeures + Round((dblSecondes / 3600), 2)
            
160         Call rstHeures.MoveNext
165       Loop
          
170       Call rstHeures.Close
         
175       rstSommaire.Fields("Total") = dblTotal
          
180       Call rstSommaire.Update
        
185       Call rstProjets.MoveNext
190     Loop

195     Set DR_ListeProjets.DataSource = rstSommaire

200     DR_ListeProjets.Sections("Section4").Controls("lblDateDebut").Caption = mskDateDebut.Text
205     DR_ListeProjets.Sections("Section4").Controls("lblDateFin").Caption = mskDateFin.Text

210     Call DR_ListeProjets.Show(vbModal)

215     Call rstSommaire.Close
        
220     Set rstSommaire = Nothing
225     Set rstHeures = Nothing
230     Set rstProjets = Nothing

235     Exit Sub

AfficherErreur:

240     woups "frmChoixDateSommairePunch", "ImprimerPunch", Err, Erl
End Sub

Private Sub ImprimerPunchGRB()

5       On Error GoTo AfficherErreur

10      Dim rstPunch       As ADODB.Recordset
15      Dim rstSommaire    As ADODB.Recordset
20      Dim datDebut       As Date
25      Dim datFin         As Date
30      Dim dblTotal       As Double
35      Dim dblSecondes    As Double
40      Dim dblHeures      As Double
45      Dim dblTotalHeures As Double
50      Dim dblTotalKM     As Double
55      Dim sWhere         As String
60      Dim sWhereFamille  As String
65      Dim iCompteur      As Integer

70      Call g_connData.Execute("DELETE * FROM GRB_ImpressionSommairePunchGRB")

75      Set rstSommaire = New ADODB.Recordset
80      Set rstPunch = New ADODB.Recordset

85      Call rstSommaire.Open("SELECT * FROM GRB_ImpressionSommairePunchGRB", g_connData, adOpenDynamic, adLockOptimistic)

90      sWhere = "((Left(NoProjet, 6) = 'E" & Right$(Year(Date), 1) & "3000' OR Left(NoProjet, 6) = 'M" & Right$(Year(Date), 1) & "3000') AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin Is Not Null)"

95      For iCompteur = 1 To lvwFamille.ListItems.count
100       If lvwFamille.ListItems(iCompteur).Checked = True Then
105         If sWhereFamille = "" Then
110           sWhereFamille = " AND (Famille = " & lvwFamille.ListItems(iCompteur).Tag
115         Else
120           sWhereFamille = sWhereFamille & " OR Famille = " & lvwFamille.ListItems(iCompteur).Tag
125         End If
130       End If
135     Next

140     sWhere = sWhere & sWhereFamille & ")"

145     Call rstPunch.Open("SELECT employe, NoProjet, Date, HeureDébut, HeureFin, NbreKM, Commentaire FROM GRB_employés INNER JOIN GRB_Punch ON GRB_employés.noemploye = GRB_Punch.NoEmploye WHERE " & sWhere & " ORDER BY employe, Date, HeureDébut, HeureFin", g_connData, adOpenDynamic, adLockOptimistic)

150     Do While Not rstPunch.EOF
155       Call rstSommaire.AddNew

160       rstSommaire.Fields("Employé") = rstPunch.Fields("employe")
165       rstSommaire.Fields("NoProjet") = rstPunch.Fields("NoProjet")
170       rstSommaire.Fields("Date") = rstPunch.Fields("Date")
175       rstSommaire.Fields("Commentaire") = rstPunch.Fields("Commentaire")
180       rstSommaire.Fields("HeureDébut") = rstPunch.Fields("HeureDébut")
185       rstSommaire.Fields("HeureFin") = rstPunch.Fields("HeureFin")
190       rstSommaire.Fields("NbreKM") = rstPunch.Fields("NbreKM")

195       datDebut = rstPunch.Fields("HeureDébut")

200       If rstPunch.Fields("HeureFin") = "24:00" Then
205         datFin = TimeSerial(24, 0, 0)
210       Else
215         datFin = rstPunch.Fields("HeureFin")
220       End If
            
225       dblSecondes = DateDiff("s", datDebut, datFin)

230       dblHeures = Int(dblSecondes / 3600)
            
235       dblSecondes = dblSecondes - (3600 * dblHeures)
            
240       dblTotal = dblHeures + Round((dblSecondes / 3600), 2)
                               
245       dblTotalHeures = dblTotalHeures + dblTotal

250       If Not IsNull(rstPunch.Fields("NbreKM")) Then
255         If Trim(rstPunch.Fields("NbreKM")) <> "" Then
260           dblTotalKM = dblTotalKM + rstPunch.Fields("NbreKM")
265         End If
270       End If
                               
275       rstSommaire.Fields("Total") = dblTotal
          
280       Call rstSommaire.Update
        
285       Call rstPunch.MoveNext
290     Loop

295     Set DR_SommairePunchGRB.DataSource = rstSommaire

300     DR_SommairePunchGRB.Orientation = rptOrientLandscape

305     DR_SommairePunchGRB.Sections("Section4").Controls("lblDateDebut").Caption = mskDateDebut.Text
310     DR_SommairePunchGRB.Sections("Section4").Controls("lblDateFin").Caption = mskDateFin.Text

315     DR_SommairePunchGRB.Sections("Section5").Controls("lblGrandTotal").Caption = dblTotalHeures
320     DR_SommairePunchGRB.Sections("Section5").Controls("lblGrandTotalKM").Caption = dblTotalKM

325     Call DR_SommairePunchGRB.Show(vbModal)

330     Call rstSommaire.Close
        
335     Set rstSommaire = Nothing
340     Set rstPunch = Nothing

345     Exit Sub

AfficherErreur:

350     woups "frmChoixDateSommairePunch", "ImprimerPunch", Err, Erl
End Sub

Private Sub ImprimerSommaireHeures()

5       On Error GoTo AfficherErreur

10      Dim rstPunch       As ADODB.Recordset
15      Dim datDebut       As Date
20      Dim datFin         As Date
25      Dim dblSoumElec    As Double
30      Dim dblSoumMec     As Double
35      Dim dblProjGRBElec As Double
40      Dim dblProjGRBMec  As Double
45      Dim dblProjElec    As Double
50      Dim dblProjMec     As Double
55      Dim dblFabElec     As Double
60      Dim dblFabMec      As Double
65      Dim dblRechElec    As Double
70      Dim dblRechMec     As Double
75      Dim dblAppelsElec  As Double
80      Dim dblAppelsMec   As Double
85      Dim dblGrandTotal  As Double
90      Dim dblSecondes    As Double
95      Dim dblHeures      As Double
100     Dim dblTotalHeures As Double
105     Dim sWhere         As String

110     Set rstPunch = New ADODB.Recordset

        'Soumissions électriques
115     sWhere = "(LEFT(NoProjet, 1) = 'E' AND MID(NoProjet, 3, 1) = '1' AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin Is Not Null)"

120     Call rstPunch.Open("SELECT Date, HeureDébut, HeureFin FROM GRB_Punch WHERE " & sWhere, g_connData, adOpenForwardOnly, adLockReadOnly)

125     Do While Not rstPunch.EOF
130       datDebut = rstPunch.Fields("HeureDébut")

135       If rstPunch.Fields("HeureFin") = "24:00" Then
140         datFin = TimeSerial(24, 0, 0)
145       Else
150         datFin = rstPunch.Fields("HeureFin")
155       End If
            
160       dblSecondes = DateDiff("s", datDebut, datFin)

165       dblHeures = Int(dblSecondes / 3600)
            
170       dblSecondes = dblSecondes - (3600 * dblHeures)
            
175       dblSoumElec = dblSoumElec + (dblHeures + Round((dblSecondes / 3600), 2))
        
180       Call rstPunch.MoveNext
185     Loop

190     Call rstPunch.Close

        'Soumissions mécaniques
195     sWhere = "(LEFT(NoProjet, 1) = 'M' AND MID(NoProjet, 3, 1) = '1' AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin Is Not Null)"

200     Call rstPunch.Open("SELECT Date, HeureDébut, HeureFin FROM GRB_Punch WHERE " & sWhere, g_connData, adOpenForwardOnly, adLockReadOnly)

205     Do While Not rstPunch.EOF
210       datDebut = rstPunch.Fields("HeureDébut")

215       If rstPunch.Fields("HeureFin") = "24:00" Then
220         datFin = TimeSerial(24, 0, 0)
225       Else
230         datFin = rstPunch.Fields("HeureFin")
235       End If
            
240       dblSecondes = DateDiff("s", datDebut, datFin)

245       dblHeures = Int(dblSecondes / 3600)
            
250       dblSecondes = dblSecondes - (3600 * dblHeures)
            
255       dblSoumMec = dblSoumMec + (dblHeures + Round((dblSecondes / 3600), 2))
        
260       Call rstPunch.MoveNext
265     Loop

270     Call rstPunch.Close

        'Projets GRB électriques
275     sWhere = "(LEFT(NoProjet, 1) = 'E' AND MID(NoProjet, 3, 4) = '3000' AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin Is Not Null)"

280     Call rstPunch.Open("SELECT Date, HeureDébut, HeureFin FROM GRB_Punch WHERE " & sWhere, g_connData, adOpenForwardOnly, adLockReadOnly)

285     Do While Not rstPunch.EOF
290       datDebut = rstPunch.Fields("HeureDébut")

295       If rstPunch.Fields("HeureFin") = "24:00" Then
300         datFin = TimeSerial(24, 0, 0)
305       Else
310         datFin = rstPunch.Fields("HeureFin")
315       End If
            
320       dblSecondes = DateDiff("s", datDebut, datFin)

325       dblHeures = Int(dblSecondes / 3600)
            
330       dblSecondes = dblSecondes - (3600 * dblHeures)
            
335       dblProjGRBElec = dblProjGRBElec + (dblHeures + Round((dblSecondes / 3600), 2))
        
340       Call rstPunch.MoveNext
345     Loop

350     Call rstPunch.Close

        'Projets GRB mécaniques
355     sWhere = "(LEFT(NoProjet, 1) = 'M' AND MID(NoProjet, 3, 4) = '3000' AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin Is Not Null)"

360     Call rstPunch.Open("SELECT Date, HeureDébut, HeureFin FROM GRB_Punch WHERE " & sWhere, g_connData, adOpenForwardOnly, adLockReadOnly)

365     Do While Not rstPunch.EOF
370       datDebut = rstPunch.Fields("HeureDébut")

375       If rstPunch.Fields("HeureFin") = "24:00" Then
380         datFin = TimeSerial(24, 0, 0)
385       Else
390         datFin = rstPunch.Fields("HeureFin")
395       End If
            
400       dblSecondes = DateDiff("s", datDebut, datFin)

405       dblHeures = Int(dblSecondes / 3600)
            
410       dblSecondes = dblSecondes - (3600 * dblHeures)
            
415       dblProjGRBMec = dblProjGRBMec + (dblHeures + Round((dblSecondes / 3600), 2))
        
420       Call rstPunch.MoveNext
425     Loop

430     Call rstPunch.Close

        'Projets électriques
435     sWhere = "(LEFT(NoProjet, 1) = 'E' AND MID(NoProjet, 3, 1) = '3' AND MID(NoProjet, 3, 4) <> '3000' AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin Is Not Null)"

440     Call rstPunch.Open("SELECT Date, HeureDébut, HeureFin FROM GRB_Punch WHERE " & sWhere, g_connData, adOpenForwardOnly, adLockReadOnly)

445     Do While Not rstPunch.EOF
450       datDebut = rstPunch.Fields("HeureDébut")

455       If rstPunch.Fields("HeureFin") = "24:00" Then
460         datFin = TimeSerial(24, 0, 0)
465       Else
470         datFin = rstPunch.Fields("HeureFin")
475       End If
            
480       dblSecondes = DateDiff("s", datDebut, datFin)

485       dblHeures = Int(dblSecondes / 3600)
            
490       dblSecondes = dblSecondes - (3600 * dblHeures)
            
495       dblProjElec = dblProjElec + (dblHeures + Round((dblSecondes / 3600), 2))
        
500       Call rstPunch.MoveNext
505     Loop

510     Call rstPunch.Close

        'Projets mécaniques
515     sWhere = "(LEFT(NoProjet, 1) = 'M' AND MID(NoProjet, 3, 1) = '3' AND MID(NoProjet, 3, 4) <> '3000' AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin Is Not Null)"

520     Call rstPunch.Open("SELECT Date, HeureDébut, HeureFin FROM GRB_Punch WHERE " & sWhere, g_connData, adOpenForwardOnly, adLockReadOnly)

525     Do While Not rstPunch.EOF
530       datDebut = rstPunch.Fields("HeureDébut")

535       If rstPunch.Fields("HeureFin") = "24:00" Then
540         datFin = TimeSerial(24, 0, 0)
545       Else
550         datFin = rstPunch.Fields("HeureFin")
555       End If
            
560       dblSecondes = DateDiff("s", datDebut, datFin)

565       dblHeures = Int(dblSecondes / 3600)
            
570       dblSecondes = dblSecondes - (3600 * dblHeures)
            
575       dblProjMec = dblProjMec + (dblHeures + Round((dblSecondes / 3600), 2))
        
580       Call rstPunch.MoveNext
585     Loop

590     Call rstPunch.Close

        'Fabrication électrique
595     sWhere = "(LEFT(NoProjet, 1) = 'E' AND MID(NoProjet, 3, 1) = '4' AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin Is Not Null)"

600     Call rstPunch.Open("SELECT Date, HeureDébut, HeureFin FROM GRB_Punch WHERE " & sWhere, g_connData, adOpenForwardOnly, adLockReadOnly)

605     Do While Not rstPunch.EOF
610       datDebut = rstPunch.Fields("HeureDébut")

615       If rstPunch.Fields("HeureFin") = "24:00" Then
620         datFin = TimeSerial(24, 0, 0)
625       Else
630         datFin = rstPunch.Fields("HeureFin")
635       End If
            
640       dblSecondes = DateDiff("s", datDebut, datFin)

645       dblHeures = Int(dblSecondes / 3600)
            
650       dblSecondes = dblSecondes - (3600 * dblHeures)
            
655       dblFabElec = dblFabElec + (dblHeures + Round((dblSecondes / 3600), 2))
        
660       Call rstPunch.MoveNext
665     Loop

670     Call rstPunch.Close

        'Fabrication mécanique
675     sWhere = "(LEFT(NoProjet, 1) = 'M' AND MID(NoProjet, 3, 1) = '4' AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin Is Not Null)"

680     Call rstPunch.Open("SELECT Date, HeureDébut, HeureFin FROM GRB_Punch WHERE " & sWhere, g_connData, adOpenForwardOnly, adLockReadOnly)

685     Do While Not rstPunch.EOF
690       datDebut = rstPunch.Fields("HeureDébut")

695       If rstPunch.Fields("HeureFin") = "24:00" Then
700         datFin = TimeSerial(24, 0, 0)
705       Else
710         datFin = rstPunch.Fields("HeureFin")
715       End If
            
720       dblSecondes = DateDiff("s", datDebut, datFin)

725       dblHeures = Int(dblSecondes / 3600)
            
730       dblSecondes = dblSecondes - (3600 * dblHeures)
            
735       dblFabMec = dblFabMec + (dblHeures + Round((dblSecondes / 3600), 2))
        
740       Call rstPunch.MoveNext
745     Loop

750     Call rstPunch.Close

        'Recherche et développement électrique
755     sWhere = "(LEFT(NoProjet, 1) = 'E' AND MID(NoProjet, 3, 1) = '9' AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin Is Not Null)"

760     Call rstPunch.Open("SELECT Date, HeureDébut, HeureFin FROM GRB_Punch WHERE " & sWhere, g_connData, adOpenForwardOnly, adLockReadOnly)

765     Do While Not rstPunch.EOF
770       datDebut = rstPunch.Fields("HeureDébut")

775       If rstPunch.Fields("HeureFin") = "24:00" Then
780         datFin = TimeSerial(24, 0, 0)
785       Else
790         datFin = rstPunch.Fields("HeureFin")
795       End If
            
800       dblSecondes = DateDiff("s", datDebut, datFin)

805       dblHeures = Int(dblSecondes / 3600)
            
810       dblSecondes = dblSecondes - (3600 * dblHeures)
            
815       dblRechElec = dblRechElec + (dblHeures + Round((dblSecondes / 3600), 2))
        
820       Call rstPunch.MoveNext
825     Loop

830     Call rstPunch.Close

        'Recherche et développement mécanique
835     sWhere = "(LEFT(NoProjet, 1) = 'M' AND MID(NoProjet, 3, 1) = '9' AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin Is Not Null)"

840     Call rstPunch.Open("SELECT Date, HeureDébut, HeureFin FROM GRB_Punch WHERE " & sWhere, g_connData, adOpenForwardOnly, adLockReadOnly)

845     Do While Not rstPunch.EOF
850       datDebut = rstPunch.Fields("HeureDébut")

855       If rstPunch.Fields("HeureFin") = "24:00" Then
860         datFin = TimeSerial(24, 0, 0)
865       Else
870         datFin = rstPunch.Fields("HeureFin")
875       End If
            
880       dblSecondes = DateDiff("s", datDebut, datFin)

885       dblHeures = Int(dblSecondes / 3600)
            
890       dblSecondes = dblSecondes - (3600 * dblHeures)
            
895       dblRechMec = dblRechMec + (dblHeures + Round((dblSecondes / 3600), 2))
        
900       Call rstPunch.MoveNext
905     Loop

910     Call rstPunch.Close

        'Appels de services électriques
915     sWhere = "(LEFT(NoProjet, 1) = 'E' AND (MID(NoProjet, 3, 1) = '5' OR MID(NoProjet, 3, 1) = '7') AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin Is Not Null)"

920     Call rstPunch.Open("SELECT Date, HeureDébut, HeureFin FROM GRB_Punch WHERE " & sWhere, g_connData, adOpenForwardOnly, adLockReadOnly)

925     Do While Not rstPunch.EOF
930       datDebut = rstPunch.Fields("HeureDébut")

935       If rstPunch.Fields("HeureFin") = "24:00" Then
940         datFin = TimeSerial(24, 0, 0)
945       Else
950         datFin = rstPunch.Fields("HeureFin")
955       End If
            
960       dblSecondes = DateDiff("s", datDebut, datFin)

965       dblHeures = Int(dblSecondes / 3600)
            
970       dblSecondes = dblSecondes - (3600 * dblHeures)
            
975       dblAppelsElec = dblAppelsElec + (dblHeures + Round((dblSecondes / 3600), 2))
        
980       Call rstPunch.MoveNext
985     Loop

990     Call rstPunch.Close

        'Appels de services mécaniques
995     sWhere = "(LEFT(NoProjet, 1) = 'M' AND (MID(NoProjet, 3, 1) = '5' OR MID(NoProjet, 3, 1) = '7') AND Date BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND HeureFin Is Not Null)"

1000    Call rstPunch.Open("SELECT Date, HeureDébut, HeureFin FROM GRB_Punch WHERE " & sWhere, g_connData, adOpenForwardOnly, adLockReadOnly)

1005    Do While Not rstPunch.EOF
1010      datDebut = rstPunch.Fields("HeureDébut")

1015      If rstPunch.Fields("HeureFin") = "24:00" Then
1020        datFin = TimeSerial(24, 0, 0)
1025      Else
1030        datFin = rstPunch.Fields("HeureFin")
1035      End If
            
1040      dblSecondes = DateDiff("s", datDebut, datFin)

1045      dblHeures = Int(dblSecondes / 3600)
            
1050      dblSecondes = dblSecondes - (3600 * dblHeures)
            
1055      dblAppelsMec = dblAppelsMec + (dblHeures + Round((dblSecondes / 3600), 2))
        
1060      Call rstPunch.MoveNext
1065    Loop

1070    Call rstPunch.Close

1075    dblGrandTotal = dblSoumElec + _
                        dblSoumMec + _
                        dblProjGRBElec + _
                        dblProjGRBMec + _
                        dblProjElec + _
                        dblProjMec + _
                        dblFabElec + _
                        dblFabMec + _
                        dblRechElec + _
                        dblRechMec + _
                        dblAppelsElec + _
                        dblAppelsMec

        'Ce recordset sert à rien. Il sert seulement à l'ouverture du DataReport.
        'Un DataReport ne peut pas ouvrir s'il n'a pas de DataSource
1080    Call rstPunch.Open("SELECT * FROM GRB_Punch WHERE NoProjet = '0000000'", g_connData, adOpenForwardOnly, adLockReadOnly)

1085    Set DR_SommaireHeures.DataSource = rstPunch

1090    DR_SommaireHeures.Orientation = rptOrientPortrait

1095    DR_SommaireHeures.Sections("Section2").Controls("lblDateDebut").Caption = mskDateDebut.Text
1100    DR_SommaireHeures.Sections("Section2").Controls("lblDateFin").Caption = mskDateFin.Text

1105    DR_SommaireHeures.Sections("Section2").Controls("lblSoumElec").Caption = dblSoumElec
1110    DR_SommaireHeures.Sections("Section2").Controls("lblSoumMec").Caption = dblSoumMec

1115    DR_SommaireHeures.Sections("Section2").Controls("lblTotalSoum").Caption = dblSoumElec + dblSoumMec

1120    DR_SommaireHeures.Sections("Section2").Controls("lblProjGRBElec").Caption = dblProjGRBElec
1125    DR_SommaireHeures.Sections("Section2").Controls("lblProjGRBMec").Caption = dblProjGRBMec

1130    DR_SommaireHeures.Sections("Section2").Controls("lblTotalProjGRB").Caption = dblProjGRBElec + dblProjGRBMec

1135    DR_SommaireHeures.Sections("Section2").Controls("lblProjElec").Caption = dblProjElec
1140    DR_SommaireHeures.Sections("Section2").Controls("lblProjMec").Caption = dblProjMec

1145    DR_SommaireHeures.Sections("Section2").Controls("lblTotalProj").Caption = dblProjElec + dblProjMec

1150    DR_SommaireHeures.Sections("Section2").Controls("lblFabElec").Caption = dblFabElec
1155    DR_SommaireHeures.Sections("Section2").Controls("lblFabMec").Caption = dblFabMec

1160    DR_SommaireHeures.Sections("Section2").Controls("lblTotalFab").Caption = dblFabElec + dblFabMec

1165    DR_SommaireHeures.Sections("Section2").Controls("lblRechElec").Caption = dblRechElec
1170    DR_SommaireHeures.Sections("Section2").Controls("lblRechMec").Caption = dblRechMec

1175    DR_SommaireHeures.Sections("Section2").Controls("lblTotalRech").Caption = dblRechElec + dblRechMec

1180    DR_SommaireHeures.Sections("Section2").Controls("lblAppelsElec").Caption = dblAppelsElec
1185    DR_SommaireHeures.Sections("Section2").Controls("lblAppelsMec").Caption = dblAppelsMec

1190    DR_SommaireHeures.Sections("Section2").Controls("lblTotalAppels").Caption = dblAppelsElec + dblAppelsMec

1195    DR_SommaireHeures.Sections("Section2").Controls("lblGrandTotal").Caption = dblGrandTotal

1200    Call DR_SommaireHeures.Show(vbModal)

1205    Call rstPunch.Close
        
1210    Set rstPunch = Nothing

1215    Exit Sub

AfficherErreur:

1220    woups "frmChoixDateSommairePunch", "ImprimerSommaireHeures", Err, Erl
End Sub

Private Sub optProjets_Click(Index As Integer)
  If optProjets(I_OPT_PROJETS_GRB).Value = True Then
    fraFamille.Enabled = True
  Else
    fraFamille.Enabled = False
  End If
End Sub
