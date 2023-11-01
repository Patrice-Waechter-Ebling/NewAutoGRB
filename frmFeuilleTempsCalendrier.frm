VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFeuilleTempsCalendrier 
   BackColor       =   &H00000000&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   3720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   1320
      Width           =   855
   End
   Begin MSComCtl2.MonthView mvwDate 
      Height          =   2370
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      ScrollRate      =   1
      ShowToday       =   0   'False
      StartOfWeek     =   90243073
      CurrentDate     =   37735
      MaxDate         =   2958464
   End
End
Attribute VB_Name = "frmFeuilleTempsCalendrier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

5       On Error GoTo AfficherErreur
        
        'Ajout de la date
  
10      frmFeuilleTemps.txtSemaine.Text = GetDateTexte(GetFirstDay(mvwDate.Value))
  
15      frmFeuilleTemps.m_datSemaine = mvwDate.Value
  
        'Rempli le listview
20      Call frmFeuilleTemps.RemplirListView
  
25      Call Unload(Me)

30      Exit Sub

AfficherErreur:

35      woups "frmFeuilleTempsCalendrier", "cmdOK_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur
        
        'Ouverture de la fenêtre
10      Dim datDate As Date
15      Dim sDate   As String
20      Dim sYear   As String
25      Dim sMonth  As String
30      Dim sDay    As String
  
35      sDate = frmFeuilleTemps.txtSemaine.Text
  
        'Année
40      sYear = Right$(sDate, 4)
  
        'Pour enlever l'année de la date
45      sDate = Left$(sDate, Len(sDate) - 4)
  
        'Jour
50      sDay = Trim$(Left$(sDate, 2))
  
        'Pour enlever le jour de la date
55      sDate = Right$(sDate, Len(sDate) - 2)
  
        'Pour avoir le no du mois
60      sMonth = GetNoMois(Trim$(sDate))
            
        'Conversion en date
65      datDate = DateSerial(sYear, sMonth, sDay)
  
        'Met la date actuelle comme valeur
70      mvwDate.Value = datDate
  
75      Call SelectionnerLigne(datDate)

80      Exit Sub

AfficherErreur:

85      woups "frmFeuilleTempsCalendrier", "Form_Load", Err, Erl
End Sub

Private Function GetNoMois(ByVal sMois As String) As Integer

5       On Error GoTo AfficherErreur
        
        'Procédure pour avoir le numéro du mois
10      Dim iNoMois As Integer
  
15      sMois = UCase(sMois)
  
20      Select Case sMois
          Case "JANVIER":   iNoMois = 1
          Case "FÉVRIER":   iNoMois = 2
          Case "MARS":      iNoMois = 3
          Case "AVRIL":     iNoMois = 4
          Case "MAI":       iNoMois = 5
          Case "JUIN":      iNoMois = 6
          Case "JUILLET":   iNoMois = 7
          Case "AOÛT":      iNoMois = 8
          Case "SEPTEMBRE": iNoMois = 9
          Case "OCTOBRE":   iNoMois = 10
          Case "NOVEMBRE":  iNoMois = 11
          Case "DÉCEMBRE":  iNoMois = 12
25      End Select
  
30      GetNoMois = iNoMois

35      Exit Function

AfficherErreur:

40      woups "frmFeuilleTempsCalendrier", "GetNoMois", Err, Erl
End Function

Private Sub SelectionnerLigne(ByVal DateClicked As Date)

5       On Error GoTo AfficherErreur
              'Procédure pour sélectionner la ligne au complet
10      Dim datDim  As Date
15      Dim datSam  As Date
  
20      mvwDate.MultiSelect = True
  
25      datDim = GetFirstDay(DateClicked)
    
30      datSam = GetLastDay(DateClicked)
  
35      mvwDate.Value = datDim
  
40      mvwDate.SelStart = datDim
  
45      mvwDate.SelEnd = datSam

50      Exit Sub

AfficherErreur:

55      woups "frmFeuilleTempsCalendrier", "SelectionnerLigne", Err, Erl
End Sub

Private Sub mvwDate_Click()

5       On Error GoTo AfficherErreur
              'Sélectionner la semaine au complet
10      mvwDate.MultiSelect = False
  
15      Call SelectionnerLigne(mvwDate.Value)
    
20      Call cmdOk.SetFocus

25      Exit Sub

AfficherErreur:

30      woups "frmFeuilleTempsCalendrier", "mvwDate_Click", Err, Erl
End Sub
