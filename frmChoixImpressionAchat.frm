VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmChoixImpressionAchat 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impression des achats"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmChoixImpressionAchat.frx":0000
   ScaleHeight     =   5175
   ScaleWidth      =   4275
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.MonthView mvwDate 
      Height          =   2370
      Left            =   1440
      TabIndex        =   17
      Top             =   2640
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   0
      Appearance      =   1
      StartOfWeek     =   90243073
      CurrentDate     =   37952
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   1320
      TabIndex        =   25
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdImprimer 
      Caption         =   "Imprimer"
      Height          =   375
      Left            =   2760
      TabIndex        =   26
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Frame fraDate 
      BackColor       =   &H00000000&
      Caption         =   "Date (AA-MM-JJ)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   18
      Top             =   3360
      Width           =   3975
      Begin VB.CommandButton cmdDateFin 
         Caption         =   "..."
         Height          =   255
         Left            =   2160
         TabIndex        =   24
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton cmdDateDebut 
         Caption         =   "..."
         Height          =   255
         Left            =   2160
         TabIndex        =   21
         Top             =   360
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskDateDebut 
         Height          =   255
         Left            =   960
         TabIndex        =   20
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDateFin 
         Height          =   255
         Left            =   960
         TabIndex        =   23
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "Au :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Du :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame fraCategorie 
      BackColor       =   &H00000000&
      Caption         =   "Catégories d'achat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3975
      Begin VB.OptionButton optCategorie 
         BackColor       =   &H00000000&
         Caption         =   "Formation"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   5
         Top             =   840
         Width           =   2055
      End
      Begin VB.OptionButton optCategorie 
         BackColor       =   &H00000000&
         Caption         =   "Équipement && outillage"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   15
         Top             =   2040
         Width           =   2055
      End
      Begin VB.OptionButton optCategorie 
         BackColor       =   &H00000000&
         Caption         =   "Équipement && outillage PPE"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   13
         Top             =   1800
         Width           =   2295
      End
      Begin VB.OptionButton optCategorie 
         BackColor       =   &H00000000&
         Caption         =   "Équipement de bureau"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   11
         Top             =   1560
         Width           =   2055
      End
      Begin VB.OptionButton optCategorie 
         BackColor       =   &H00000000&
         Caption         =   "Bâtiment"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   9
         Top             =   1320
         Width           =   2055
      End
      Begin VB.OptionButton optCategorie 
         BackColor       =   &H00000000&
         Caption         =   "Logiciel interne"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   7
         Top             =   1080
         Width           =   2055
      End
      Begin VB.OptionButton optCategorie 
         BackColor       =   &H00000000&
         Caption         =   "Réparation outils GRB"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton optCategorie 
         BackColor       =   &H00000000&
         Caption         =   "Stocks plancher"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "83"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3000
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "99"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3000
         TabIndex        =   16
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "98"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3000
         TabIndex        =   14
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "97"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3000
         TabIndex        =   12
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "95"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3000
         TabIndex        =   10
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "85"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3000
         TabIndex        =   8
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "80"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3000
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "01 à 12"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3000
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmChoixImpressionAchat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const I_CATEGORIE_MOIS As Integer = 0
Private Const I_CATEGORIE_80   As Integer = 1
Private Const I_CATEGORIE_83   As Integer = 2
Private Const I_CATEGORIE_85   As Integer = 3
Private Const I_CATEGORIE_95   As Integer = 4
Private Const I_CATEGORIE_97   As Integer = 5
Private Const I_CATEGORIE_98   As Integer = 6
Private Const I_CATEGORIE_99   As Integer = 7

Private Enum enumDate
  I_DATE_DEBUT = 0
  I_DATE_FIN = 1
End Enum

Private m_eDate      As enumDate
Private m_eCatalogue As enumCatalogue
Private m_iCategorie As Integer

Private Sub cmdAnnuler_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmChoixImpressionAchat", "cmdAnnuler_Click", Err, Erl
End Sub

Private Sub cmdDateDebut_Click()

5       On Error GoTo AfficherErreur

10      Dim bAfficherDate As Boolean

15      m_eDate = I_DATE_DEBUT
  
20      If mskDateDebut.Text <> vbNullString Then
25        If InStr(1, mskDateDebut.Text, "_") = 0 Then
30          mvwDate.Year = Left$(mskDateDebut.Text, 4)
35          mvwDate.Month = Mid$(mskDateDebut.Text, 6, 2)
40          mvwDate.Day = Right$(mskDateDebut.Text, 2)
45        Else
50          bAfficherDate = True
55        End If
60      Else
65        bAfficherDate = True
70      End If
      
75      If bAfficherDate = True Then
80        mvwDate.Year = Year(Date)
85        mvwDate.Month = Month(Date)
90        mvwDate.Day = Day(Date)
95      End If
      
100     mvwDate.Visible = True
  
105     Call mvwDate.SetFocus

110     Exit Sub

AfficherErreur:

115     woups "frmChoixImpressionAchat", "cmdDateDebut_Click", Err, Erl
End Sub

Private Sub cmdDateFin_Click()

5       On Error GoTo AfficherErreur

10      Dim bAfficherDate As Boolean
  
15      m_eDate = I_DATE_FIN
  
20      If mskDateFin.Text <> vbNullString Then
25        If InStr(1, mskDateFin.Text, "_") = 0 Then
30          mvwDate.Year = Left$(mskDateDebut.Text, 4)
35          mvwDate.Month = Mid$(mskDateDebut.Text, 6, 2)
40          mvwDate.Day = Right$(mskDateDebut.Text, 2)
45        Else
50          bAfficherDate = True
55        End If
60      Else
65        bAfficherDate = True
70      End If
  
75      If bAfficherDate = True Then
80        mvwDate.Year = Year(Date)
85        mvwDate.Month = Month(Date)
90        mvwDate.Day = Day(Date)
95      End If
  
100     mvwDate.Visible = True
  
105     Call mvwDate.SetFocus

110     Exit Sub

AfficherErreur:

115     woups "frmChoixImpressionAchat", "cmdDateFin_Click", Err, Erl
End Sub

Private Sub cmdImprimer_Click()

5       On Error GoTo AfficherErreur

10      Dim rstPiece     As ADODB.Recordset
15      Dim rstTotal     As ADODB.Recordset
20      Dim sSelect      As String
25      Dim sFrom        As String
30      Dim sWhere       As String
35      Dim sNumeroDebut As String
40      Dim sNumeroFin   As String
    
45      If Len(mskDateDebut.Text) <> 10 Or Len(mskDateFin.Text) <> 10 Then
50        Call MsgBox("Date invalide!", vbOKOnly, "Erreur")

55        Exit Sub
60      End If
    
65      Select Case m_eCatalogue
          Case ELECTRIQUE:
70          sNumeroDebut = "E" & Mid$(mskDateDebut.Text, 4, 1) & "3000-"
75          sNumeroFin = "E" & Mid$(mskDateFin.Text, 4, 1) & "3000-"
    
          Case MECANIQUE:
80          sNumeroDebut = "M" & Mid$(mskDateDebut.Text, 4, 1) & "3000-"
85          sNumeroFin = "M" & Mid$(mskDateFin.Text, 4, 1) & "3000-"
90      End Select
  
95      sSelect = "GRB_Achat.IDAchat, GRB_Achat.IndexAchat, GRB_Achat.Raison, " & _
                  "GRB_Achat.DateAchat, GRB_Employés.initiale, " & _
                  "GRB_Achat_Pieces.PIECE, GRB_Achat_Pieces.Qté, " & _
                  "GRB_Achat_Pieces.Desc_FR, GRB_Achat_Pieces.Manufact, " & _
                  "GRB_Achat_Pieces.Prix_list , GRB_Achat_Pieces.Escompte, " & _
                  "GRB_Achat_Pieces.Prix_net, GRB_Fournisseur.NomFournisseur, " & _
                  "GRB_Achat_Pieces.Prix_total"
  
100     sFrom = "GRB_Fournisseur INNER JOIN " & _
                "(GRB_Employés INNER JOIN " & _
                "(GRB_Achat INNER JOIN GRB_Achat_Pieces ON " & _
                "(GRB_Achat.IndexAchat = GRB_Achat_Pieces.IndexAchat) " & _
                "AND (GRB_Achat.IDAchat = GRB_Achat_Pieces.IDAchat)) " & _
                "ON GRB_Employés.noEmploye = GRB_Achat.Acheteur) " & _
                "ON GRB_Fournisseur.IDFRS = GRB_Achat_Pieces.IDFRS"
  
105     Select Case m_iCategorie
          Case I_CATEGORIE_MOIS:
110         sWhere = "GRB_Achat.IDAchat BETWEEN '" & _
                     sNumeroDebut & Mid$(mskDateDebut.Text, 6, 2) & _
                     "' AND '" & sNumeroFin & Mid$(mskDateFin.Text, 6, 2) & _
                     "' AND DateAchat BETWEEN '" & mskDateDebut.Text & _
                     "' AND '" & mskDateFin.Text & "'"
    
          Case I_CATEGORIE_80:
115         sWhere = "GRB_Achat.IDAchat BETWEEN '" & sNumeroDebut & "80' AND '" & _
                     sNumeroFin & "80' AND DateAchat BETWEEN '" & mskDateDebut.Text & _
                     "' AND '" & mskDateFin.Text & "'"
                     
          Case I_CATEGORIE_83:
120         sWhere = "GRB_Achat.IDAchat BETWEEN '" & sNumeroDebut & "83' AND '" & _
                     sNumeroFin & "83' AND DateAchat BETWEEN '" & mskDateDebut.Text & _
                     "' AND '" & mskDateFin.Text & "'"
       
          Case I_CATEGORIE_85:
125         sWhere = "GRB_Achat.IDAchat BETWEEN '" & sNumeroDebut & "85' AND '" & _
                     sNumeroFin & "85' AND DateAchat BETWEEN '" & mskDateDebut.Text & _
                     "' AND '" & mskDateFin.Text & "'"
    
          Case I_CATEGORIE_95:
130         sWhere = "GRB_Achat.IDAchat BETWEEN '" & sNumeroDebut & "95' AND '" & _
                     sNumeroFin & "95' AND DateAchat BETWEEN '" & mskDateDebut.Text & _
                     "' AND '" & mskDateFin.Text & "'"
   
          Case I_CATEGORIE_97:
135         sWhere = "GRB_Achat.IDAchat BETWEEN '" & sNumeroDebut & "97' AND '" & _
                     sNumeroFin & "97' AND DateAchat BETWEEN '" & mskDateDebut.Text & _
                     "' AND '" & mskDateFin.Text & "'"
    
          Case I_CATEGORIE_98:
140         sWhere = "GRB_Achat.IDAchat BETWEEN '" & sNumeroDebut & "98' AND '" & _
                     sNumeroFin & "98' AND DateAchat BETWEEN '" & mskDateDebut.Text & _
                     "' AND '" & mskDateFin.Text & "'"
    
          Case I_CATEGORIE_99:
145         sWhere = "GRB_Achat.IDAchat BETWEEN '" & sNumeroDebut & "99' AND '" & _
                     sNumeroFin & "99' AND DateAchat BETWEEN '" & mskDateDebut.Text & _
                     "' AND '" & mskDateFin.Text & "'"
150     End Select
  
155     Set rstTotal = New ADODB.Recordset
160     Set rstPiece = New ADODB.Recordset
  
165     Call rstTotal.Open("SELECT SUM(Prix_total) As PrixTotal FROM GRB_Achat_Pieces INNER JOIN GRB_Achat ON (GRB_Achat.IDAchat = GRB_Achat_Pieces.IDAchat) AND (GRB_Achat.IndexAchat = GRB_Achat_Pieces.IndexAchat) WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)

170     Call rstPiece.Open("SELECT " & sSelect & " FROM " & sFrom & " WHERE " & sWhere & " ORDER BY GRB_Achat.IndexAchat, PIECE", g_connData, adOpenDynamic, adLockOptimistic)

175     If rstPiece.EOF = True Then
180       Screen.MousePointer = vbDefault

185       Call MsgBox("Aucun achat à imprimer!", vbOKOnly, "Erreur")
          
190       Call rstTotal.Close
195       Set rstTotal = Nothing
          
200       Call rstPiece.Close
205       Set rstPiece = Nothing
          
210       Exit Sub
215     End If
           
220     DR_ListeAchats.Orientation = rptOrientLandscape
    
225     Set DR_ListeAchats.DataSource = rstPiece
        
230     DR_ListeAchats.Sections("Section5").Controls("lblTotal").Caption = Conversion(Replace(rstTotal.Fields("PrixTotal"), ".", ","), MODE_ARGENT)
  
235     Call rstTotal.Close
240     Set rstTotal = Nothing
  
245     Select Case m_iCategorie
          Case I_CATEGORIE_MOIS:
250         If Mid$(mskDateDebut.Text, 6, 2) = Mid$(mskDateFin.Text, 6, 2) And Left$(mskDateDebut.Text, 4) = Left$(mskDateFin.Text, 4) Then
255           DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & Mid$(mskDateDebut.Text, 6, 2)
260         Else
265           DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & Mid$(mskDateDebut.Text, 6, 2) & " à " & sNumeroFin & Mid$(mskDateFin.Text, 6, 2)
270         End If
    
          Case I_CATEGORIE_80:
275         If Left$(mskDateDebut.Text, 4) = Left$(mskDateFin.Text, 4) Then
280           DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "80"
285         Else
290           DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "80 à " & sNumeroFin & "80"
295         End If

300       Case I_CATEGORIE_83:
305         If Left$(mskDateDebut.Text, 4) = Left$(mskDateFin.Text, 4) Then
310           DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "83"
315         Else
320           DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "83 à " & sNumeroFin & "83"
325         End If

          Case I_CATEGORIE_85:
330         If Left$(mskDateDebut.Text, 4) = Left$(mskDateFin.Text, 4) Then
335           DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "85"
340         Else
345           DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "85 à " & sNumeroFin & "85"
350         End If
    
          Case I_CATEGORIE_95:
355         If Left$(mskDateDebut.Text, 4) = Left$(mskDateFin.Text, 4) Then
360           DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "95"
365         Else
370           DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "95 à " & sNumeroFin & "95"
375         End If
    
          Case I_CATEGORIE_97:
380         If Left$(mskDateDebut.Text, 4) = Left$(mskDateFin.Text, 4) Then
385           DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "97"
390         Else
395           DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "97 à " & sNumeroFin & "97"
400         End If
      
          Case I_CATEGORIE_98:
405         If Left$(mskDateDebut.Text, 4) = Left$(mskDateFin.Text, 4) Then
410           DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "98"
415         Else
420           DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "98 à " & sNumeroFin & "98"
425         End If
    
          Case I_CATEGORIE_99:
430         If Left$(mskDateDebut.Text, 4) = Left$(mskDateFin.Text, 4) Then
435           DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "99"
440         Else
445           DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "99 à " & sNumeroFin & "99"
450         End If
455     End Select
  
460     DR_ListeAchats.Sections("Section4").Controls("lblDate").Caption = "Du " & mskDateDebut.Text & " Au " & mskDateFin.Text
  
465     DR_ListeAchats.Sections("Section3").Controls("lblDateImpression").Caption = ConvertDate(Date)

470     Call DR_ListeAchats.Show(vbModal)
  
475     Call rstPiece.Close
480     Set rstPiece = Nothing
  
485     Call Unload(Me)

490     Exit Sub

AfficherErreur:

495     woups "frmChoixImpressionAchat", "cmdImprimer_Click", Err, Erl
End Sub

Public Sub Afficher(ByVal eCatalogue As enumCatalogue)

5       On Error GoTo AfficherErreur

10      optCategorie(I_CATEGORIE_MOIS).Value = True
  
15      m_eCatalogue = eCatalogue
  
20      Call Me.Show(vbModal)

25      Exit Sub

AfficherErreur:

30      woups "frmChoixImpressionAchat", "Afficher", Err, Erl
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

35      woups "frmChoixImpressionAchat", "mskDateDebut_GotFocus", Err, Erl
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

35      woups "frmChoixImpressionAchat", "mskDateFin_GotFocus", Err, Erl
End Sub

Private Sub mskDateDebut_LostFocus()

5       On Error GoTo AfficherErreur

10      Dim sDate As String

        'Enlève le mask
15      mskDateDebut.mask = vbNullString
  
20      sDate = mskDateDebut.Text
  
        'Vide le champs si l'utilisateur n'a rien écrit
25      If sDate = "__-__-__" Then
30        mskDateDebut.Text = vbNullString
35      Else
          'Remet l'année sur 8 chiffres
40        If Len(sDate) = 8 Then
45          If IsDate(sDate) Then
50            mskDateDebut.Text = Year(DateSerial(Left$(sDate, 2), Mid$(sDate, 4, 2), Right$(sDate, 2))) & Mid$(sDate, 3, 8)
55          End If
60        End If
65      End If

70      Exit Sub

AfficherErreur:

75      woups "frmChoixImpressionAchat", "mskDateDebut_LostFocus", Err, Erl
End Sub

Private Sub mskDateFin_LostFocus()

5       On Error GoTo AfficherErreur

10      Dim sDate As String
  
        'Enlève le mask
15      mskDateFin.mask = vbNullString
  
20      sDate = mskDateFin.Text
  
        'Vide le champs si l'utilisateur n'a rien écrit
25      If sDate = "__-__-__" Then
30        mskDateFin.Text = vbNullString
35      Else
          'Remet l'année sur 8 chiffres
40        If Len(sDate) = 8 Then
45          If IsDate(sDate) Then
50            mskDateFin.Text = Year(DateSerial(Left$(sDate, 2), Mid$(sDate, 4, 2), Right$(sDate, 2))) & Mid$(sDate, 3, 8)
55          End If
60        End If
65      End If

70      Exit Sub

AfficherErreur:

75      woups "frmChoixImpressionAchat", "mskDateFin_LostFocus", Err, Erl
End Sub

Private Sub mvwDate_DateClick(ByVal DateClicked As Date)

5       On Error GoTo AfficherErreur

10      Select Case m_eDate
          Case I_DATE_DEBUT: mskDateDebut.Text = ConvertDate(DateClicked)
          Case I_DATE_FIN:   mskDateFin.Text = ConvertDate(DateClicked)
15      End Select
  
20      mvwDate.Visible = False

25      Exit Sub

AfficherErreur:

30      woups "frmChoixImpressionAchat", "mvwDate_DateClick", Err, Erl
End Sub

Private Sub mvwDate_LostFocus()

5       On Error GoTo AfficherErreur

10      mvwDate.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmChoixImpressionAchat", "mvwDate_LostFocus", Err, Erl
End Sub

Private Sub optCategorie_Click(Index As Integer)

5       On Error GoTo AfficherErreur

10      m_iCategorie = Index

15      Exit Sub

AfficherErreur:

20      woups "frmChoixImpressionAchat", "optCategorie_Click", Err, Erl
End Sub
