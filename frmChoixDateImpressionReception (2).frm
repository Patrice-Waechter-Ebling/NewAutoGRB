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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   3840
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
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   0
      Appearance      =   1
      ShowToday       =   0   'False
      StartOfWeek     =   152633345
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

Private m_eDate As enumDate
Private m_eCatalogue As enumCatalogue
Private m_eTypeReception As enumTypeReception
Private m_sNoProjet As String
Private m_sIDAchat As String
Private m_iIndexAchat As Integer

Private Sub cmdAnnuler_Click()

 On Error GoTo Oups

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixDateImpressionReception", "cmdAnnuler_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdImprimer_Click()

 On Error GoTo Oups

 Dim rstReception As ADODB.Recordset
 Dim rstTotal As ADODB.Recordset

 If Len(mskDateDebut.Text) =   Then
 Call mskDateDebut_LostFocus
 End If

1  If Len(mskDateFin.Text) =   Then
 Call mskDateFin_LostFocus
End If

2 If ValiderDate(mskDateDebut.Text) = False Then
 Call MsgBox("Date de début invalide!", vbOKOnly, "Erreur")

 Exit Sub
 End If

 If ValiderDate(mskDateFin.Text) = False Then
 Call MsgBox("Date de fin invalide!", vbOKOnly, "Erreur")

 Exit Sub
 End If

  If mskDateFin.Text < mskDateDebut.Text Then
  Call MsgBox("La date de fin doit être plus grande que la date de début!", vbOKOnly, "Erreur")

  Exit Sub
  End If

  Set rstReception = New ADODB.Recordset

  If m_eTypeReception = PROJET Then
  Call rstReception.Open("SELECT GrbProjet_Pieces.*, (Escompte / 100) As ModifEscompte, (Prix_Net * Qté) As TotalReception FROM GrbProjet_Pieces WHERE IDProjet = '" & m_sNoProjet & "' AND ((Recu = True AND DateRéception BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "') OR (Retour = True AND DateRetour BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "'))", g_connData, adOpenDynamic, adLockOptimistic)
  Else
Call rstReception.Open("SELECT GrbAchat_Pieces.*, (Escompte / 100) As ModifEscompte, (Prix_Net * Qté) As TotalReception FROM GrbAchat_Pieces WHERE IDAchat = '" & m_sIDAchat & "' AND IndexAchat = " & m_iIndexAchat & " AND ((Recu = True AND DateRéception BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "') OR (Retour = True AND DateRetour BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "'))", g_connData, adOpenDynamic, adLockOptimistic)
End If

Set DR_Reception.DataSource = rstReception

DR_Reception.Sections("Section4").Controls("lblDate").Caption = "Du " & mskDateDebut.Text & " Au " & mskDateFin.Text

Set rstTotal = New ADODB.Recordset

If m_eTypeReception = ACHAT Then
 DR_Reception.Sections("Section4").Controls("lblTitreProjetAchat").Caption = "Achat :"
 DR_Reception.Sections("Section4").Controls("lblProjetAchat").Caption = m_sIDAchat & "-" & Right$("00" & m_iIndexAchat, 3)

 DR_Reception.Sections("Section1").Controls("txtDate").DataField = "DateRéception"
 DR_Reception.Sections("Section1").Controls("txtQuantite").DataField = "Qté"
 DR_Reception.Sections("Section1").Controls("txtPiece").DataField = "PIECE"
 DR_Reception.Sections("Section1").Controls("txtPrixListe").DataField = "Prix_List"
DR_Reception.Sections("Section1").Controls("txtEscompte").DataField = "ModifEscompte"
 DR_Reception.Sections("Section1").Controls("txtPrixNet").DataField = "Prix_Net"
 DR_Reception.Sections("Section1").Controls("txtTotal").DataField = "TotalReception"

 Call rstTotal.Open("SELECT SUM(Qté * Prix_Net) As Total FROM GrbAchat_Pieces WHERE IDAchat = '" & m_sIDAchat & "' AND IndexAchat = " & m_iIndexAchat & " AND ((Recu = True AND DateRéception BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "') OR (Retour = True AND DateRetour BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "'))", g_connData, adOpenDynamic, adLockOptimistic)

 If Not IsNull(rstTotal.Fields("Total")) Then
 DR_Reception.Sections("Section5").Controls("lblTotal").Caption = Conversion(rstTotal.Fields("Total"), MODE_ARGENT)
 Else
1  DR_Reception.Sections("Section5").Controls("lblTotal").Caption = Conversion("0", MODE_ARGENT)
 End If
 Else
 DR_Reception.Sections("Section4").Controls("lblTitreProjetAchat").Caption = "Projet :"
 DR_Reception.Sections("Section4").Controls("lblProjetAchat").Caption = m_sNoProjet

 DR_Reception.Sections("Section1").Controls("txtDate").DataField = "DateRéception"
 DR_Reception.Sections("Section1").Controls("txtQuantite").DataField = "Qté"
 DR_Reception.Sections("Section1").Controls("txtPiece").DataField = "NumItem"
 DR_Reception.Sections("Section1").Controls("txtPrixListe").DataField = "Prix_List"
 DR_Reception.Sections("Section1").Controls("txtEscompte").DataField = "ModifEscompte"
 DR_Reception.Sections("Section1").Controls("txtPrixNet").DataField = "Prix_Net"
 DR_Reception.Sections("Section1").Controls("txtTotal").DataField = "TotalReception"

 Call rstTotal.Open("SELECT SUM(Qté * Prix_Net) As Total FROM GrbProjet_Pieces WHERE IDProjet = '" & m_sNoProjet & "' AND ((Recu = True AND DateRéception BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "') OR (Retour = True AND DateRetour BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "'))", g_connData, adOpenDynamic, adLockOptimistic)

If Not IsNull(rstTotal.Fields("Total")) Then
 DR_Reception.Sections("Section5").Controls("lblTotal").Caption = Conversion(rstTotal.Fields("Total"), MODE_ARGENT)
Else
 DR_Reception.Sections("Section5").Controls("lblTotal").Caption = Conversion("0", MODE_ARGENT)
End If
End If

2  Call rstTotal.Close
Set rstTotal = Nothing

30 Call DR_Reception.Show(vbModal)

Call Unload(Me)

Exit Sub

Oups:

wOups "frmChoixDateImpressionReception", "cmdImprimer_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Click()

 On Error GoTo Oups

 mvwDate.Visible = False

 Exit Sub

Oups:

 wOups "frmChoixDateImpressionReception", "Form_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 m_eDate = AUCUNE

 Exit Sub

Oups:

 wOups "frmChoixDateImpressionReception", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub mvwDate_LostFocus()

 On Error GoTo Oups

 mvwDate.Visible = False

 Exit Sub

Oups:

 wOups "frmChoixDateImpressionReception", "mvwDate_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mvwDate_DateClick(ByVal DateClicked As Date)

 On Error GoTo Oups

 Select Case m_eDate
 Case DEBUT: mskDateDebut.Text = ConvertDate(DateClicked)
 Case Fin: mskDateFin.Text = ConvertDate(DateClicked)
 End Select
 
 m_eDate = AUCUNE
 
 'Enlever le calendrier
 mvwDate.Visible = False

 Exit Sub

Oups:

 wOups "frmChoixDateImpressionReception", "mvwDate_DateClick", Err, Err.number, Err.Description
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

 wOups "frmChoixDateImpressionReception", "mskDateDebut_GotFocus", Err, Err.number, Err.Description
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

 wOups "frmChoixDateImpressionReception", "mskDateFin_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskDateDebut_LostFocus()

 On Error GoTo Oups
 
 'Enlève le mask
 mskDateDebut.mask = vbNullString
 
 'Vide le champs si l'utilisateur n'a rien écrit
 If mskDateDebut.Text = "__-__-__" Then
 mskDateDebut.Text = vbNullString
 Else
 'Remet l'année sur   chiffres
 If Len(mskDateDebut.Text) =   Then
 If IsDate(mskDateDebut.Text) Then
 mskDateDebut.Text = Year(DateSerial(Left$(mskDateDebut.Text, 2), Mid$(mskDateDebut.Text, 4, 2), Right$(mskDateDebut.Text, 2))) & Mid$(mskDateDebut.Text, 3, 8)
 End If
 End If
 End If

  Exit Sub

Oups:

  wOups "frmChoixDateImpressionReception", "mskDateDebut_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskDateFin_LostFocus()

 On Error GoTo Oups
 
 'Enlève le mask
 mskDateFin.mask = vbNullString
 
 'Vide le champs si l'utilisateur n'a rien écrit
 If mskDateFin.Text = "__-__-__" Then
 mskDateFin.Text = vbNullString
 Else
 'Remet l'année sur   chiffres
 If Len(mskDateFin.Text) =   Then
 If IsDate(mskDateFin.Text) Then
 mskDateFin.Text = Year(DateSerial(Left$(mskDateFin.Text, 2), Mid$(mskDateFin.Text, 4, 2), Right$(mskDateFin.Text, 2))) & Mid$(mskDateFin.Text, 3, 8)
 End If
 End If
 End If

  Exit Sub

Oups:

  wOups "frmChoixDateImpressionReception", "mskDateFin_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub cmdDateDebut_Click()

 On Error GoTo Oups
 'Ouverture du calendrier
 
 'Si il y a une date valide, la date par défaut est celle-ci, sinon c'est la date
 'd'aujourd'hui
 If Trim$(mskDateDebut.Text) <> vbNullString Then
 If ValiderDate(mskDateDebut.Text) = True Then
 mvwDate.Value = mskDateDebut.Text
 Else
 mvwDate.Value = Date
 End If
 Else
 mvwDate.Value = Date
 End If
 
 m_eDate = DEBUT
 
  mvwDate.Visible = True
 
  Call mvwDate.SetFocus

  Exit Sub

Oups:

  wOups "frmChoixDateImpressionReception", "cmdDateDebut_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdDateFin_Click()

 On Error GoTo Oups
 'Ouverture du calendrier
 
 'Si il y a une date valide, la date par défaut est celle-ci, sinon c'est la date
 'd'aujourd'hui
 If Trim$(mskDateFin.Text) <> vbNullString Then
 If ValiderDate(mskDateFin.Text) = True Then
 mvwDate.Value = mskDateFin.Text
 Else
 mvwDate.Value = Date
 End If
 Else
 mvwDate.Value = Date
 End If
 
 m_eDate = Fin
 
  mvwDate.Visible = True
 
  Call mvwDate.SetFocus

  Exit Sub

Oups:

  wOups "frmChoixDateImpressionReception", "cmdDateFin_Click", Err, Err.number, Err.Description
End Sub

Private Function ValiderDate(ByVal sDate As String) As Boolean

 On Error GoTo Oups

 'Validation des dates
 If Not IsDate(sDate) Then
 ValiderDate = False
 Else
 ValiderDate = True
 End If

 Exit Function

Oups:

 wOups "frmChoixDateImpressionReception", "ValiderDate", Err, Err.number, Err.Description
End Function

Public Sub Afficher(ByVal sNoProjet As String, ByVal eCatalogue As enumCatalogue, ByVal eTypeReception As enumTypeReception)

 On Error GoTo Oups

 m_eTypeReception = eTypeReception

 Select Case eTypeReception
 Case PROJET:
 m_sNoProjet = sNoProjet

 Case ACHAT:
 m_sIDAchat = Left$(sNoProjet, 9)
 m_iIndexAchat = CInt(Right$(sNoProjet, 3))
 End Select

 m_eCatalogue = eCatalogue

 Call Me.Show(vbModal)

 Exit Sub

Oups:

  wOups "frmChoixDateImpressionReception", "Afficher", Err, Err.number, Err.Description
End Sub
