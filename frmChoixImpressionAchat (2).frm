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
   MDIChild        =   -1  'True
   ScaleHeight     =   5175
   ScaleWidth      =   4275
   Begin MSComCtl2.MonthView mvwDate 
      Height          =   2370
      Left            =   1440
      TabIndex        =   17
      Top             =   2640
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   0
      Appearance      =   1
      StartOfWeek     =   152633345
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


Private Enum enumDate
 I_DATE_DEBUT = 0
 I_DATE_FIN = 1
End Enum

Private m_eDate As enumDate
Private m_eCatalogue As enumCatalogue
Private m_iCategorie As Integer

Private Sub cmdAnnuler_Click()

 On Error GoTo Oups

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixImpressionAchat", "cmdAnnuler_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdDateDebut_Click()

 On Error GoTo Oups

 Dim bAfficherDate As Boolean

 m_eDate = I_DATE_DEBUT
 
 If mskDateDebut.Text <> vbNullString Then
 If InStr(1, mskDateDebut.Text, "_") = 0 Then
 mvwDate.Year = Left$(mskDateDebut.Text, 4)
 mvwDate.Month = Mid$(mskDateDebut.Text, 6, 2)
 mvwDate.Day = Right$(mskDateDebut.Text, 2)
 Else
 bAfficherDate = True
 End If
  Else
  bAfficherDate = True
  End If
 
  If bAfficherDate = True Then
  mvwDate.Year = Year(Date)
  mvwDate.Month = Month(Date)
  mvwDate.Day = Day(Date)
  End If
 
10 mvwDate.Visible = True
 
Call mvwDate.SetFocus

Exit Sub

Oups:

wOups "frmChoixImpressionAchat", "cmdDateDebut_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdDateFin_Click()

 On Error GoTo Oups

 Dim bAfficherDate As Boolean
 
 m_eDate = I_DATE_FIN
 
 If mskDateFin.Text <> vbNullString Then
 If InStr(1, mskDateFin.Text, "_") = 0 Then
 mvwDate.Year = Left$(mskDateDebut.Text, 4)
 mvwDate.Month = Mid$(mskDateDebut.Text, 6, 2)
 mvwDate.Day = Right$(mskDateDebut.Text, 2)
 Else
 bAfficherDate = True
 End If
  Else
  bAfficherDate = True
  End If
 
  If bAfficherDate = True Then
  mvwDate.Year = Year(Date)
  mvwDate.Month = Month(Date)
  mvwDate.Day = Day(Date)
  End If
 
10 mvwDate.Visible = True
 
Call mvwDate.SetFocus

Exit Sub

Oups:

wOups "frmChoixImpressionAchat", "cmdDateFin_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdImprimer_Click()

 On Error GoTo Oups

 Dim rstPiece As ADODB.Recordset
 Dim rstTotal As ADODB.Recordset
 Dim sSelect As String
 Dim sFrom As String
 Dim sWhere As String
 Dim sNumeroDebut As String
 Dim sNumeroFin As String
 
 If Len(mskDateDebut.Text) <> 10 Or Len(mskDateFin.Text) <> 10 Then
 Call MsgBox("Date invalide!", vbOKOnly, "Erreur")

 Exit Sub
  End If
 
  Select Case m_eCatalogue
 Case ELECTRIQUE:
  sNumeroDebut = "E" & Mid$(mskDateDebut.Text, 4, 1) & "3000-"
  sNumeroFin = "E" & Mid$(mskDateFin.Text, 4, 1) & "3000-"
 
 Case MECANIQUE:
  sNumeroDebut = "M" & Mid$(mskDateDebut.Text, 4, 1) & "3000-"
  sNumeroFin = "M" & Mid$(mskDateFin.Text, 4, 1) & "3000-"
  End Select
 
  sSelect = "GrbAchat.IDAchat, GrbAchat.IndexAchat, GrbAchat.Raison, " & _
 "GrbAchat.DateAchat, GrbEmployés.initiale, " & _
 "GrbAchat_Pieces.PIECE, GrbAchat_Pieces.Qté, " & _
 "GrbAchat_Pieces.Desc_FR, GrbAchat_Pieces.Manufact, " & _
 "GrbAchat_Pieces.Prix_list , GrbAchat_Pieces.Escompte, " & _
 "GrbAchat_Pieces.Prix_net, GrbFournisseur.NomFournisseur, " & _
 "GrbAchat_Pieces.Prix_total"
 
10 sFrom = "GrbFournisseur INNER JOIN " & _
 "(GrbEmployés INNER JOIN " & _
 "(GrbAchat INNER JOIN GrbAchat_Pieces ON " & _
 "(GrbAchat.IndexAchat = GrbAchat_Pieces.IndexAchat) " & _
 "AND (GrbAchat.IDAchat = GrbAchat_Pieces.IDAchat)) " & _
 "ON GrbEmployés.noEmploye = GrbAchat.Acheteur) " & _
 "ON GrbFournisseur.IDFRS = GrbAchat_Pieces.IDFRS"
 
Select Case m_iCategorie
 Case I_CATEGORIE_MOIS:
 sWhere = "GrbAchat.IDAchat BETWEEN '" & _
 sNumeroDebut & Mid$(mskDateDebut.Text, 6, 2) & _
 "' AND '" & sNumeroFin & Mid$(mskDateFin.Text, 6, 2) & _
 "' AND DateAchat BETWEEN '" & mskDateDebut.Text & _
 "' AND '" & mskDateFin.Text & "'"
 
 Case I_CATEGORIE_80:
 sWhere = "GrbAchat.IDAchat BETWEEN '" & sNumeroDebut & "80' AND '" & _
 sNumeroFin & "80' AND DateAchat BETWEEN '" & mskDateDebut.Text & _
 "' AND '" & mskDateFin.Text & "'"
 
 Case I_CATEGORIE_83:
 sWhere = "GrbAchat.IDAchat BETWEEN '" & sNumeroDebut & "83' AND '" & _
 sNumeroFin & "83' AND DateAchat BETWEEN '" & mskDateDebut.Text & _
 "' AND '" & mskDateFin.Text & "'"
 
 Case I_CATEGORIE_85:
 sWhere = "GrbAchat.IDAchat BETWEEN '" & sNumeroDebut & "85' AND '" & _
 sNumeroFin & "85' AND DateAchat BETWEEN '" & mskDateDebut.Text & _
 "' AND '" & mskDateFin.Text & "'"
 
 Case I_CATEGORIE_95:
 sWhere = "GrbAchat.IDAchat BETWEEN '" & sNumeroDebut & "95' AND '" & _
 sNumeroFin & "95' AND DateAchat BETWEEN '" & mskDateDebut.Text & _
 "' AND '" & mskDateFin.Text & "'"
 
 Case I_CATEGORIE_97:
 sWhere = "GrbAchat.IDAchat BETWEEN '" & sNumeroDebut & "97' AND '" & _
 sNumeroFin & "97' AND DateAchat BETWEEN '" & mskDateDebut.Text & _
 "' AND '" & mskDateFin.Text & "'"
 
 Case I_CATEGORIE_98:
 sWhere = "GrbAchat.IDAchat BETWEEN '" & sNumeroDebut & "98' AND '" & _
 sNumeroFin & "98' AND DateAchat BETWEEN '" & mskDateDebut.Text & _
 "' AND '" & mskDateFin.Text & "'"
 
 Case I_CATEGORIE_99:
 sWhere = "GrbAchat.IDAchat BETWEEN '" & sNumeroDebut & "99' AND '" & _
 sNumeroFin & "99' AND DateAchat BETWEEN '" & mskDateDebut.Text & _
 "' AND '" & mskDateFin.Text & "'"
End Select
 
Set rstTotal = New ADODB.Recordset
1  Set rstPiece = New ADODB.Recordset
 
Call rstTotal.Open("SELECT SUM(Prix_total) As PrixTotal FROM GrbAchat_Pieces INNER JOIN GrbAchat ON (GrbAchat.IDAchat = GrbAchat_Pieces.IDAchat) AND (GrbAchat.IndexAchat = GrbAchat_Pieces.IndexAchat) WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)

 Call rstPiece.Open("SELECT " & sSelect & " FROM " & sFrom & " WHERE " & sWhere & " ORDER BY GrbAchat.IndexAchat, PIECE", g_connData, adOpenDynamic, adLockOptimistic)

If rstPiece.EOF = True Then
 Screen.MousePointer = vbDefault

 Call MsgBox("Aucun achat à imprimer!", vbOKOnly, "Erreur")
 
 Call rstTotal.Close
1  Set rstTotal = Nothing
 
 Call rstPiece.Close
 Set rstPiece = Nothing
 
 Exit Sub
End If
 
DR_ListeAchats.Orientation = rptOrientLandscape
 
Set DR_ListeAchats.DataSource = rstPiece
 
DR_ListeAchats.Sections("Section5").Controls("lblTotal").Caption = Conversion(Replace(rstTotal.Fields("PrixTotal"), ".", ","), MODE_ARGENT)
 
Call rstTotal.Close
Set rstTotal = Nothing
 
Select Case m_iCategorie
 Case I_CATEGORIE_MOIS:
 If Mid$(mskDateDebut.Text, 6, 2) = Mid$(mskDateFin.Text, 6, 2) And Left$(mskDateDebut.Text, 4) = Left$(mskDateFin.Text, 4) Then
 DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & Mid$(mskDateDebut.Text, 6, 2)
 Else
 DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & Mid$(mskDateDebut.Text, 6, 2) & " à " & sNumeroFin & Mid$(mskDateFin.Text, 6, 2)
 End If
 
 Case I_CATEGORIE_80:
 If Left$(mskDateDebut.Text, 4) = Left$(mskDateFin.Text, 4) Then
 DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "80"
 Else
 DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "80 à " & sNumeroFin & "80"
 End If

Case I_CATEGORIE_83:
If Left$(mskDateDebut.Text, 4) = Left$(mskDateFin.Text, 4) Then
 DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "83"
 Else
 DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "83 à " & sNumeroFin & "83"
 End If

 Case I_CATEGORIE_85:
 If Left$(mskDateDebut.Text, 4) = Left$(mskDateFin.Text, 4) Then
 DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "85"
 Else
 DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "85 à " & sNumeroFin & "85"
 End If
 
 Case I_CATEGORIE_95:
 If Left$(mskDateDebut.Text, 4) = Left$(mskDateFin.Text, 4) Then
 DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "95"
 Else
 DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "95 à " & sNumeroFin & "95"
 End If
 
 Case I_CATEGORIE_97:
 If Left$(mskDateDebut.Text, 4) = Left$(mskDateFin.Text, 4) Then
 DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "97"
 Else
 DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "   à " & sNumeroFin & "97"
 End If
 
 Case I_CATEGORIE_98:
4 If Left$(mskDateDebut.Text, 4) = Left$(mskDateFin.Text, 4) Then
4 DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "98"
4 Else
4 DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "   à " & sNumeroFin & "98"
4 End If
 
 Case I_CATEGORIE_99:
4 If Left$(mskDateDebut.Text, 4) = Left$(mskDateFin.Text, 4) Then
4 DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "99"
4 Else
4 DR_ListeAchats.Sections("Section4").Controls("lblNumero").Caption = sNumeroDebut & "9  à " & sNumeroFin & "99"
4 End If
4 End Select
 
4  DR_ListeAchats.Sections("Section4").Controls("lblDate").Caption = "Du " & mskDateDebut.Text & " Au " & mskDateFin.Text
 
4  DR_ListeAchats.Sections("Section3").Controls("lblDateImpression").Caption = ConvertDate(Date)

4  Call DR_ListeAchats.Show(vbModal)
 
4  Call rstPiece.Close
4  Set rstPiece = Nothing
 
4  Call Unload(Me)

4  Exit Sub

Oups:

4  wOups "frmChoixImpressionAchat", "cmdImprimer_Click", Err, Err.number, Err.Description
End Sub

Public Sub Afficher(ByVal eCatalogue As enumCatalogue)

 On Error GoTo Oups

 optCategorie(I_CATEGORIE_MOIS).Value = True
 
 m_eCatalogue = eCatalogue
 
 Call Me.Show(vbModal)

 Exit Sub

Oups:

 wOups "frmChoixImpressionAchat", "Afficher", Err, Err.number, Err.Description
End Sub

Private Sub mskDateDebut_GotFocus()

 On Error GoTo Oups

 'Met l'année sur 2 chiffres
 If Len(mskDateDebut.Text) = 10 Then
 mskDateDebut.Text = Right$(mskDateDebut.Text, 8)
 End If
 
 mskDateDebut.mask = "##-##-##"

 Exit Sub

Oups:

 wOups "frmChoixImpressionAchat", "mskDateDebut_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskDateFin_GotFocus()

 On Error GoTo Oups

 'Met l'année sur 2 chiffres
 If Len(mskDateFin.Text) = 10 Then
 mskDateFin.Text = Right$(mskDateFin.Text, 8)
 End If
 
 mskDateFin.mask = "##-##-##"

 Exit Sub

Oups:

 wOups "frmChoixImpressionAchat", "mskDateFin_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskDateDebut_LostFocus()

 On Error GoTo Oups

 Dim sDate As String

 'Enlève le mask
 mskDateDebut.mask = vbNullString
 
 sDate = mskDateDebut.Text
 
 'Vide le champs si l'utilisateur n'a rien écrit
 If sDate = "__-__-__" Then
 mskDateDebut.Text = vbNullString
 Else
 'Remet l'année sur   chiffres
 If Len(sDate) =   Then
 If IsDate(sDate) Then
 mskDateDebut.Text = Year(DateSerial(Left$(sDate, 2), Mid$(sDate, 4, 2), Right$(sDate, 2))) & Mid$(sDate, 3, 8)
 End If
  End If
  End If

  Exit Sub

Oups:

  wOups "frmChoixImpressionAchat", "mskDateDebut_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskDateFin_LostFocus()

 On Error GoTo Oups

 Dim sDate As String
 
 'Enlève le mask
 mskDateFin.mask = vbNullString
 
 sDate = mskDateFin.Text
 
 'Vide le champs si l'utilisateur n'a rien écrit
 If sDate = "__-__-__" Then
 mskDateFin.Text = vbNullString
 Else
 'Remet l'année sur   chiffres
 If Len(sDate) =   Then
 If IsDate(sDate) Then
 mskDateFin.Text = Year(DateSerial(Left$(sDate, 2), Mid$(sDate, 4, 2), Right$(sDate, 2))) & Mid$(sDate, 3, 8)
 End If
  End If
  End If

  Exit Sub

Oups:

  wOups "frmChoixImpressionAchat", "mskDateFin_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mvwDate_DateClick(ByVal DateClicked As Date)

 On Error GoTo Oups

 Select Case m_eDate
 Case I_DATE_DEBUT: mskDateDebut.Text = ConvertDate(DateClicked)
 Case I_DATE_FIN: mskDateFin.Text = ConvertDate(DateClicked)
 End Select
 
 mvwDate.Visible = False

 Exit Sub

Oups:

 wOups "frmChoixImpressionAchat", "mvwDate_DateClick", Err, Err.number, Err.Description
End Sub

Private Sub mvwDate_LostFocus()

 On Error GoTo Oups

 mvwDate.Visible = False

 Exit Sub

Oups:

 wOups "frmChoixImpressionAchat", "mvwDate_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub optCategorie_Click(Index As Integer)

 On Error GoTo Oups

 m_iCategorie = Index

 Exit Sub

Oups:

 wOups "frmChoixImpressionAchat", "optCategorie_Click", Err, Err.number, Err.Description
End Sub
