VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmChoixDateImpressionReception 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impression réception"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChoixDateImpressionReception.frx":0000
   ScaleHeight     =   3255
   ScaleWidth      =   3840
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimer 
      Caption         =   "Imprimer"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
   End
   Begin MSComCtl2.MonthView mvwDate 
      Height          =   2370
      Left            =   1080
      TabIndex        =   0
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
   Begin VB.CommandButton cmdDateDebut 
      Caption         =   "..."
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   1560
      Width           =   375
   End
   Begin MSMask.MaskEdBox mskDateDebut 
      Height          =   255
      Left            =   1440
      TabIndex        =   3
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
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdDateFin 
      Caption         =   "..."
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "AA-MM-JJ"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date fin :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date début :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
End
Attribute VB_Name = "frmChoixDateImpressionReception"
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

Public Enum enumTypeReception
  PROJET = 0
  ACHAT = 1
End Enum

Private m_eDate          As enumDate
Private m_eCatalogue     As enumCatalogue
Private m_eTypeReception As enumTypeReception
Private m_sNoProjet      As String
Private m_sIDAchat       As String
Private m_iIndexAchat    As Integer

Private Sub cmdAnnuler_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmChoixDateImpressionReception", "cmdAnnuler_Click", Err, Erl
End Sub

Private Sub cmdImprimer_Click()

5       On Error GoTo AfficherErreur

10      Dim rstReception As ADODB.Recordset
15      Dim rstTotal     As ADODB.Recordset

16      If Len(mskDateDebut.Text) = 8 Then
17        Call mskDateDebut_LostFocus
18      End If

19      If Len(mskDateFin.Text) = 8 Then
20        Call mskDateFin_LostFocus
21      End If

22      If ValiderDate(mskDateDebut.Text) = False Then
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

80      Set rstReception = New ADODB.Recordset

85      If m_eTypeReception = PROJET Then
90        Call rstReception.Open("SELECT GRB_Projet_Pieces.*, (Escompte / 100) As ModifEscompte, (Prix_Net * Qté) As TotalReception FROM GRB_Projet_Pieces WHERE IDProjet = '" & m_sNoProjet & "' AND ((Recu = True AND DateRéception BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "') OR (Retour = True AND DateRetour BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "'))", g_connData, adOpenDynamic, adLockOptimistic)
95      Else
100       Call rstReception.Open("SELECT GRB_Achat_Pieces.*, (Escompte / 100) As ModifEscompte, (Prix_Net * Qté) As TotalReception FROM GRB_Achat_Pieces WHERE IDAchat = '" & m_sIDAchat & "' AND IndexAchat = " & m_iIndexAchat & " AND ((Recu = True AND DateRéception BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "') OR (Retour = True AND DateRetour BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "'))", g_connData, adOpenDynamic, adLockOptimistic)
105     End If

110     Set DR_Reception.DataSource = rstReception

115     DR_Reception.Sections("Section4").Controls("lblDate").Caption = "Du " & mskDateDebut.Text & " Au " & mskDateFin.Text

120     Set rstTotal = New ADODB.Recordset

125     If m_eTypeReception = ACHAT Then
130       DR_Reception.Sections("Section4").Controls("lblTitreProjetAchat").Caption = "Achat :"
135       DR_Reception.Sections("Section4").Controls("lblProjetAchat").Caption = m_sIDAchat & "-" & Right$("00" & m_iIndexAchat, 3)

140       DR_Reception.Sections("Section1").Controls("txtDate").DataField = "DateRéception"
145       DR_Reception.Sections("Section1").Controls("txtQuantite").DataField = "Qté"
150       DR_Reception.Sections("Section1").Controls("txtPiece").DataField = "PIECE"
155       DR_Reception.Sections("Section1").Controls("txtPrixListe").DataField = "Prix_List"
160       DR_Reception.Sections("Section1").Controls("txtEscompte").DataField = "ModifEscompte"
165       DR_Reception.Sections("Section1").Controls("txtPrixNet").DataField = "Prix_Net"
170       DR_Reception.Sections("Section1").Controls("txtTotal").DataField = "TotalReception"

175       Call rstTotal.Open("SELECT SUM(Qté * Prix_Net) As Total FROM GRB_Achat_Pieces WHERE IDAchat = '" & m_sIDAchat & "' AND IndexAchat = " & m_iIndexAchat & " AND ((Recu = True AND DateRéception BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "') OR (Retour = True AND DateRetour BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "'))", g_connData, adOpenDynamic, adLockOptimistic)

180       If Not IsNull(rstTotal.Fields("Total")) Then
185         DR_Reception.Sections("Section5").Controls("lblTotal").Caption = Conversion(rstTotal.Fields("Total"), MODE_ARGENT)
190       Else
195         DR_Reception.Sections("Section5").Controls("lblTotal").Caption = Conversion("0", MODE_ARGENT)
200       End If
205     Else
210       DR_Reception.Sections("Section4").Controls("lblTitreProjetAchat").Caption = "Projet :"
215       DR_Reception.Sections("Section4").Controls("lblProjetAchat").Caption = m_sNoProjet

220       DR_Reception.Sections("Section1").Controls("txtDate").DataField = "DateRéception"
225       DR_Reception.Sections("Section1").Controls("txtQuantite").DataField = "Qté"
230       DR_Reception.Sections("Section1").Controls("txtPiece").DataField = "NumItem"
235       DR_Reception.Sections("Section1").Controls("txtPrixListe").DataField = "Prix_List"
240       DR_Reception.Sections("Section1").Controls("txtEscompte").DataField = "ModifEscompte"
245       DR_Reception.Sections("Section1").Controls("txtPrixNet").DataField = "Prix_Net"
250       DR_Reception.Sections("Section1").Controls("txtTotal").DataField = "TotalReception"

255       Call rstTotal.Open("SELECT SUM(Qté * Prix_Net) As Total FROM GRB_Projet_Pieces WHERE IDProjet = '" & m_sNoProjet & "' AND ((Recu = True AND DateRéception BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "') OR (Retour = True AND DateRetour BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "'))", g_connData, adOpenDynamic, adLockOptimistic)

260       If Not IsNull(rstTotal.Fields("Total")) Then
265         DR_Reception.Sections("Section5").Controls("lblTotal").Caption = Conversion(rstTotal.Fields("Total"), MODE_ARGENT)
270       Else
275         DR_Reception.Sections("Section5").Controls("lblTotal").Caption = Conversion("0", MODE_ARGENT)
280       End If
285     End If

290     Call rstTotal.Close
295     Set rstTotal = Nothing

300     Call DR_Reception.Show(vbModal)

305     Call Unload(Me)

310     Exit Sub

AfficherErreur:

315     woups "frmChoixDateImpressionReception", "cmdImprimer_Click", Err, Erl
End Sub

Private Sub Form_Click()

5       On Error GoTo AfficherErreur

10      mvwDate.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmChoixDateImpressionReception", "Form_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      m_eDate = AUCUNE

15      Exit Sub

AfficherErreur:

20      woups "frmChoixDateImpressionReception", "Form_Load", Err, Erl
End Sub

Private Sub mvwDate_LostFocus()

5       On Error GoTo AfficherErreur

10      mvwDate.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmChoixDateImpressionReception", "mvwDate_LostFocus", Err, Erl
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

35      woups "frmChoixDateImpressionReception", "mvwDate_DateClick", Err, Erl
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

35      woups "frmChoixDateImpressionReception", "mskDateDebut_GotFocus", Err, Erl
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

35      woups "frmChoixDateImpressionReception", "mskDateFin_GotFocus", Err, Erl
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

65      woups "frmChoixDateImpressionReception", "mskDateDebut_LostFocus", Err, Erl
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

65      woups "frmChoixDateImpressionReception", "mskDateFin_LostFocus", Err, Erl
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

75      woups "frmChoixDateImpressionReception", "cmdDateDebut_Click", Err, Erl
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

75      woups "frmChoixDateImpressionReception", "cmdDateFin_Click", Err, Erl
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

40      woups "frmChoixDateImpressionReception", "ValiderDate", Err, Erl
End Function

Public Sub Afficher(ByVal sNoProjet As String, ByVal eCatalogue As enumCatalogue, ByVal eTypeReception As enumTypeReception)

5       On Error GoTo AfficherErreur

10      m_eTypeReception = eTypeReception

15      Select Case eTypeReception
          Case PROJET:
20          m_sNoProjet = sNoProjet

25        Case ACHAT:
30          m_sIDAchat = Left$(sNoProjet, 9)
35          m_iIndexAchat = CInt(Right$(sNoProjet, 3))
40      End Select

45      m_eCatalogue = eCatalogue

50      Call Me.Show(vbModal)

55      Exit Sub

AfficherErreur:

60      woups "frmChoixDateImpressionReception", "Afficher", Err, Erl
End Sub
