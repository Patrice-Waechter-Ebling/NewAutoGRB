VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFeuilleTempsCalendrier 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   3720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
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
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      ScrollRate      =   1
      ShowToday       =   0   'False
      StartOfWeek     =   152305665
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

 On Error GoTo Oups
 
 'Ajout de la date
 
 frmFeuilleTemps.txtSemaine.Text = GetDateTexte(GetFirstDay(mvwDate.Value))
 
 frmFeuilleTemps.m_datSemaine = mvwDate.Value
 
 'Rempli le listview
 Call frmFeuilleTemps.RemplirListView
 
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmFeuilleTempsCalendrier", "cmdOK_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups
 
 'Ouverture de la fenêtre
 Dim datDate As Date
 Dim sDate As String
 Dim sYear As String
 Dim sMonth As String
 Dim sDay As String
 
 sDate = frmFeuilleTemps.txtSemaine.Text
 
 'Année
 sYear = Right$(sDate, 4)
 
 'Pour enlever l'année de la date
 sDate = Left$(sDate, Len(sDate) - 4)
 
 'Jour
 sDay = Trim$(Left$(sDate, 2))
 
 'Pour enlever le jour de la date
 sDate = Right$(sDate, Len(sDate) - 2)
 
 'Pour avoir le no du mois
  sMonth = GetNoMois(Trim$(sDate))
 
 'Conversion en date
  datDate = DateSerial(sYear, sMonth, sDay)
 
 'Met la date actuelle comme valeur
  mvwDate.Value = datDate
 
  Call SelectionnerLigne(datDate)

  Exit Sub

Oups:

  wOups "frmFeuilleTempsCalendrier", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Function GetNoMois(ByVal sMois As String) As Integer

 On Error GoTo Oups
 
 'Procédure pour avoir le numéro du mois
 Dim iNoMois As Integer
 
 sMois = UCase(sMois)
 
 Select Case sMois
 Case "JANVIER": iNoMois = 1
 Case "FÉVRIER": iNoMois = 2
 Case "MARS": iNoMois = 3
 Case "AVRIL": iNoMois = 4
 Case "MAI": iNoMois = 5
 Case "JUIN": iNoMois = 6
 Case "JUILLET": iNoMois = 7
 Case "AOÛT": iNoMois = 8
 Case "SEPTEMBRE": iNoMois = 9
 Case "OCTOBRE": iNoMois = 10
 Case "NOVEMBRE": iNoMois = 11
 Case "DÉCEMBRE": iNoMois = 12
 End Select
 
 GetNoMois = iNoMois

 Exit Function

Oups:

 wOups "frmFeuilleTempsCalendrier", "GetNoMois", Err, Err.number, Err.Description
End Function

Private Sub SelectionnerLigne(ByVal DateClicked As Date)

 On Error GoTo Oups
 'Procédure pour sélectionner la ligne au complet
 Dim datDim As Date
 Dim datSam As Date
 
 mvwDate.MultiSelect = True
 
 datDim = GetFirstDay(DateClicked)
 
 datSam = GetLastDay(DateClicked)
 
 mvwDate.Value = datDim
 
 mvwDate.SelStart = datDim
 
 mvwDate.SelEnd = datSam

 Exit Sub

Oups:

 wOups "frmFeuilleTempsCalendrier", "SelectionnerLigne", Err, Err.number, Err.Description
End Sub

Private Sub mvwDate_Click()

 On Error GoTo Oups
 'Sélectionner la semaine au complet
 mvwDate.MultiSelect = False
 
 Call SelectionnerLigne(mvwDate.Value)
 
 Call cmdOk.SetFocus

 Exit Sub

Oups:

 wOups "frmFeuilleTempsCalendrier", "mvwDate_Click", Err, Err.number, Err.Description
End Sub
