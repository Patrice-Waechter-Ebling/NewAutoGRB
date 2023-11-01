VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmChoixDateImpressionFacturation 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportation feuilles de temps"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6540
   Begin MSComCtl2.MonthView mvwDate 
      Height          =   2370
      Left            =   3720
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   0
      Appearance      =   1
      ShowToday       =   0   'False
      StartOfWeek     =   152305665
      CurrentDate     =   37735
   End
   Begin VB.OptionButton optChoix 
      BackColor       =   &H00000000&
      Caption         =   "Entre 2 dates"
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
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   11
      Top             =   1560
      Width           =   1475
   End
   Begin VB.Frame fra2Dates 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
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
      Height          =   1335
      Left            =   3120
      TabIndex        =   3
      Top             =   1560
      Width           =   3255
      Begin VB.CommandButton cmdDateFin 
         Caption         =   "..."
         Height          =   255
         Left            =   2760
         TabIndex        =   7
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton cmdDateDebut 
         Caption         =   "..."
         Height          =   255
         Left            =   2760
         TabIndex        =   4
         Top             =   480
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskDateDebut 
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   480
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
         Top             =   960
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
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Date fin :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "AA-MM-JJ"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimer 
      Caption         =   "Imprimer"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Frame fraProjetEntier 
      BackColor       =   &H00000000&
      Caption         =   "Choix de l'impression"
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
      Height          =   1815
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   2775
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   120
         ScaleHeight     =   1455
         ScaleWidth      =   2415
         TabIndex        =   13
         Top             =   240
         Width           =   2415
         Begin VB.OptionButton optChoixProjetEntier 
            BackColor       =   &H00000000&
            Caption         =   "Prix coûtant du projet"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   15
            Top             =   960
            Width           =   1815
         End
         Begin VB.OptionButton optChoixProjetEntier 
            BackColor       =   &H00000000&
            Caption         =   "Liste des punchs"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   14
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
      End
   End
   Begin VB.OptionButton optChoix 
      BackColor       =   &H00000000&
      Caption         =   "Projet entier"
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
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   16
      Top             =   1080
      Value           =   -1  'True
      Width           =   1935
   End
End
Attribute VB_Name = "frmChoixDateImpressionFacturation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enumDate
 DEBUT = 0
 Fin = 1
End Enum

Private Const I_OPT_PROJET_ENTIER As Integer = 0
Private Const I_OPT_2_DATES As Integer = 1

Private Const I_OPT_LISTE_PUNCH As Integer = 0
Private Const I_OPT_COUTANT As Integer = 1

Private m_eDate As enumDate
Private m_sNoProjSoum As String
Private m_bProjet As Boolean
Private m_sClient As String
Private m_sDescription As String

Public Sub Afficher(ByVal sNoProjSoum As String, ByVal bProjet As Boolean, ByVal sClient As String, ByVal sDescription As String)

 On Error GoTo Oups

 m_sNoProjSoum = sNoProjSoum

 m_bProjet = bProjet

 If bProjet = True Then
 optChoix(I_OPT_COUTANT).Enabled = True
 Else
 optChoix(I_OPT_COUTANT).Enabled = False
 End If

 m_sClient = sClient

 m_sDescription = sDescription

 Call Me.Show(vbModal)

  Exit Sub

Oups:

  wOups "frmChoixDateImpressionFacturation", "Afficher", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnuler_Click()

 On Error GoTo Oups

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixDateImpressionFacturation", "cmdAnnuler_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdImprimer_Click()

 On Error GoTo Oups

 If optChoixProjetEntier(I_OPT_LISTE_PUNCH).Value = True Then
 Call ImprimerListePunch
 Else
 Call ImprimerPrixCoutant
 End If

 Exit Sub

Oups:

 wOups "frmChoixDateImpressionFacturation", "cmdImprimer_Click", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerListePunch()
 
 On Error GoTo Oups

 Dim rstPunch As ADODB.Recordset
 Dim rstSomme As ADODB.Recordset
 Dim iCompteur As Integer
 Dim bNonComplet As Boolean
 
 If optChoix(I_OPT_2_DATES).Value = True Then
 If mskDateDebut.Text <> "" Then
 If mskDateFin.Text <> "" Then
 If ValiderDate(mskDateDebut.Text) = True Then
 If ValiderDate(mskDateFin.Text) = True Then
 If mskDateDebut.Text > mskDateFin.Text Then
  Call MsgBox("La date de début doit être plus petite que la date de fin!", vbOKOnly, "Erreur")

  Exit Sub
  End If
  Else
  Call MsgBox("Date de fin non valide!", vbOKOnly, "Erreur")

  Exit Sub
  End If
  Else
 Call MsgBox("Date de début non valide!", vbOKOnly, "Erreur")

 Exit Sub
 End If
 Else
 Call MsgBox("La date de fin est obligatoire!", vbOKOnly, "Erreur")

 Exit Sub
 End If
 Else
 Call MsgBox("La date de début est obligatoire!", vbOKOnly, "Erreur")

 Exit Sub
 End If
End If

 'Si il y a des projets ou des soumissions
 If m_sNoProjSoum <> "" Then
 If optChoix(I_OPT_PROJET_ENTIER).Value = True Then
 For iCompteur = 1 To frmFacturation.lvwProjets.ListItems.count
1  If frmFacturation.lvwProjets.ListItems(iCompteur).SubItems(3) = "" Then
 bNonComplet = True

 Exit For
 End If
 Next

 If bNonComplet = True Then
 If MsgBox("Les punchs ne sont pas complets!" & vbNewLine & "Voulez-vous imprimer seulement les punchs complets?", vbYesNo) = vbNo Then
 Exit Sub
 End If
 End If

 Set rstPunch = New ADODB.Recordset

 rstPunch.CursorLocation = adUseServer

 '*************************************************************************
 'ajout du champ type dans la requête PAR GAÉTAN GINGRAS LE 0  FÉVRIER 2010
 If MsgBox("Désirez-vous afficher les commentaires avec le type des travaux?", vbYesNo, "Choix d'affichage") = vbYes Then
 Call rstPunch.Open("SELECT (GrbPunch.Type & ' - ' & GrbPunch.Commentaire) AS Comment, GrbPunch.Date, GrbPunch.HeureDébut, GrbPunch.HeureFin, GrbPunch.Facturé, GrbPunch.NoFacture, GrbEmployés.Initiale, Round((TimeSerial(Left(GrbPunch.HeureFin,2), RIGHT(GrbPunch.HeureFin,2),0) - TimeSerial(Left(GrbPunch.HeureDébut,2), RIGHT(GrbPunch.HeureDébut,2),0)) * 24, 2) As Total FROM GrbPunch INNER JOIN GrbEmployés ON GrbPunch.NoEmploye = GrbEmployés.noEmploye WHERE GrbPunch.NoProjet = '" & m_sNoProjSoum & "' AND HeureFin IS NOT NULL ORDER BY [Date]", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstPunch.Open("SELECT GrbPunch.Type AS Comment, GrbPunch.Date, GrbPunch.HeureDébut, GrbPunch.HeureFin, GrbPunch.Commentaire, GrbPunch.Facturé, GrbPunch.NoFacture, GrbEmployés.Initiale, Round((TimeSerial(Left(GrbPunch.HeureFin,2), RIGHT(GrbPunch.HeureFin,2),0) - TimeSerial(Left(GrbPunch.HeureDébut,2), RIGHT(GrbPunch.HeureDébut,2),0)) * 24, 2) As Total FROM GrbPunch INNER JOIN GrbEmployés ON GrbPunch.NoEmploye = GrbEmployés.noEmploye WHERE GrbPunch.NoProjet = '" & m_sNoProjSoum & "' AND HeureFin IS NOT NULL ORDER BY [Date]", g_connData, adOpenDynamic, adLockOptimistic)
 End If
 '*************************************************************************
 
Else
 For iCompteur = 1 To frmFacturation.lvwProjets.ListItems.count
 If frmFacturation.lvwProjets.ListItems(iCompteur).SubItems(3) = "" Then
 If frmFacturation.lvwProjets.ListItems(iCompteur).SubItems(1) >= mskDateDebut.Text And frmFacturation.lvwProjets.ListItems(iCompteur).SubItems(1) >= mskDateFin.Text Then
 bNonComplet = True

 Exit For
 End If
 End If
 Next

If bNonComplet = True Then
 If MsgBox("Les punchs ne sont pas complets!" & vbNewLine & "Voulez-vous imprimer seulement les punchs complets?", vbYesNo) = vbNo Then
 Exit Sub
 End If
 End If

 Set rstPunch = New ADODB.Recordset

 rstPunch.CursorLocation = adUseServer

 '**************************************************************************
 'ajout du champ type dans la requête PAR GAÉTAN GINGRAS LE 0  FÉVRIER 2010
 If MsgBox("Désirez-vous afficher les commentaires avec le type des travaux?", vbYesNo, "Choix d'affichage") = vbYes Then
 Call rstPunch.Open("SELECT (GrbPunch.Type & ' - ' & GrbPunch.Commentaire) AS Comment, GrbPunch.Date, GrbPunch.HeureDébut, GrbPunch.HeureFin, GrbPunch.Commentaire, GrbPunch.Facturé, GrbPunch.NoFacture, GrbEmployés.Initiale, Round((TimeSerial(Left(GrbPunch.HeureFin,2), RIGHT(GrbPunch.HeureFin,2),0) - TimeSerial(Left(GrbPunch.HeureDébut,2), RIGHT(GrbPunch.HeureDébut,2),0)) * 24, 2) As Total FROM GrbPunch INNER JOIN GrbEmployés ON GrbPunch.NoEmploye = GrbEmployés.noEmploye WHERE GrbPunch.NoProjet = '" & m_sNoProjSoum & "' AND HeureFin IS NOT NULL AND [GrbPunch.Date] BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' ORDER BY [Date]", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstPunch.Open("SELECT GrbPunch.Type AS Comment, GrbPunch.Date, GrbPunch.HeureDébut, GrbPunch.HeureFin, GrbPunch.Commentaire, GrbPunch.Facturé, GrbPunch.NoFacture, GrbEmployés.Initiale, Round((TimeSerial(Left(GrbPunch.HeureFin,2), RIGHT(GrbPunch.HeureFin,2),0) - TimeSerial(Left(GrbPunch.HeureDébut,2), RIGHT(GrbPunch.HeureDébut,2),0)) * 24, 2) As Total FROM GrbPunch INNER JOIN GrbEmployés ON GrbPunch.NoEmploye = GrbEmployés.noEmploye WHERE GrbPunch.NoProjet = '" & m_sNoProjSoum & "' AND HeureFin IS NOT NULL AND [GrbPunch.Date] BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' ORDER BY [Date]", g_connData, adOpenDynamic, adLockOptimistic)
 End If
 '**************************************************************************
 
 End If

 Set DR_Facturation.DataSource = rstPunch

 DR_Facturation.Orientation = rptOrientLandscape

If m_bProjet = True Then
 DR_Facturation.Sections("Section4").Controls("lblTitreNumero").Caption = "Numéro de projet :"
Else
 DR_Facturation.Sections("Section4").Controls("lblTitreNumero").Caption = "Numéro de soumission :"
End If

 DR_Facturation.Sections("Section4").Controls("lblNumero").Caption = m_sNoProjSoum
 DR_Facturation.Sections("Section4").Controls("lblClient").Caption = m_sClient

 'affiche la date
 '**************************************************
 'ajout par Gaétan Gingras le 20 mai 2009
394 If MsgBox("Désirez-vous afficher la date en bas de page ?", vbYesNo + vbInformation, "Affichage de la date") = vbYes Then
 DR_Facturation.Sections("Section3").Controls("lblDate").Caption = ConvertDate(Date)
3   Else
3   DR_Facturation.Sections("Section3").Controls("lblDate").Caption = " "
3   End If

 '**************************************************
 
 'affichage des colonnes facturé et no. de facture
 '**************************************************
 'ajout de Gaétan Gingras le 20 mai 2009
39  If MsgBox("Désirez-vous afficher les colonnes 'facturé' et 'no. facture'?", vbYesNo + vbInformation, "Affichage de la date") = vbYes Then
 DR_Facturation.Sections("Section1").Controls("text1").Visible = True
 DR_Facturation.Sections("Section1").Controls("text4").Visible = True
 DR_Facturation.Sections("Section2").Controls("label4").Visible = True
 DR_Facturation.Sections("Section2").Controls("label14").Visible = True
404 Else
4 DR_Facturation.Sections("Section1").Controls("text1").Visible = False
40  DR_Facturation.Sections("Section1").Controls("text4").Visible = False
40  DR_Facturation.Sections("Section2").Controls("label4").Visible = False
40  DR_Facturation.Sections("Section2").Controls("label14").Visible = False
40  End If
 '**************************************************
 
4 Set rstSomme = New ADODB.Recordset

41 If optChoix(I_OPT_PROJET_ENTIER).Value = True Then
414 Call rstSomme.Open("SELECT SUM(TimeSerial(Left(HeureFin,2),RIGHT(HeureFin,2),0) - TimeSerial(Left(HeureDébut,2), RIGHT(HeureDébut,2),0)) As Total FROM GrbPunch WHERE NoProjet = '" & m_sNoProjSoum & "' AND Facturé = True AND HeureFin IS NOT NULL", g_connData, adOpenDynamic, adLockOptimistic)
4 Else
4 Call rstSomme.Open("SELECT SUM(TimeSerial(Left(HeureFin,2),RIGHT(HeureFin,2),0) - TimeSerial(Left(HeureDébut,2), RIGHT(HeureDébut,2),0)) As Total FROM GrbPunch WHERE NoProjet = '" & m_sNoProjSoum & "' AND Facturé = True AND HeureFin IS NOT NULL AND [Date] BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
4 End If

4 If Not IsNull(rstSomme.Fields("Total")) Then
4 DR_Facturation.Sections("Section5").Controls("lblHeuresFacturees").Caption = Round(rstSomme.Fields("Total") * 24, 4)
4 Else
4 DR_Facturation.Sections("Section5").Controls("lblHeuresFacturees").Caption = "0"
4 End If

4 Call rstSomme.Close

4  rstSomme.CursorLocation = adUseServer

4  If optChoix(I_OPT_PROJET_ENTIER).Value = True Then
4  Call rstSomme.Open("SELECT SUM(TimeSerial(Left(HeureFin,2), RIGHT(HeureFin,2),0) - TimeSerial(Left(HeureDébut,2),RIGHT(HeureDébut,2),0)) As Total FROM GrbPunch WHERE NoProjet = '" & m_sNoProjSoum & "' AND Facturé = False AND HeureFin IS NOT NULL", g_connData, adOpenDynamic, adLockOptimistic)
4  Else
4  Call rstSomme.Open("SELECT SUM(TimeSerial(Left(HeureFin,2), RIGHT(HeureFin,2),0) - TimeSerial(Left(HeureDébut,2),RIGHT(HeureDébut,2),0)) As Total FROM GrbPunch WHERE NoProjet = '" & m_sNoProjSoum & "' AND Facturé = False AND HeureFin IS NOT NULL AND [Date] BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
4  End If

4  If Not IsNull(rstSomme.Fields("Total")) Then
4  DR_Facturation.Sections("Section5").Controls("lblHeuresNonFacturees").Caption = Round(rstSomme.Fields("Total") * 24, 4)
50 Else
DR_Facturation.Sections("Section5").Controls("lblHeuresNonFacturees").Caption = "0"
 End If

 Call rstSomme.Close
 Set rstSomme = Nothing

 DR_Facturation.Sections("Section5").Controls("lblGrandTotal").Caption = CDbl(DR_Facturation.Sections("Section5").Controls("lblHeuresFacturees").Caption) + CDbl(DR_Facturation.Sections("Section5").Controls("lblHeuresNonFacturees").Caption)

 If optChoix(I_OPT_PROJET_ENTIER).Value = True Then
 DR_Facturation.Sections("Section4").Controls("lblDateDebut").Caption = "N/A"
 DR_Facturation.Sections("Section4").Controls("lblDateFin").Caption = "N/A"
 Else
 DR_Facturation.Sections("Section4").Controls("lblDateDebut").Caption = mskDateDebut.Text
 DR_Facturation.Sections("Section4").Controls("lblDateFin").Caption = mskDateFin.Text
5  End If

5  Call DR_Facturation.Show(vbModal)

5  Call rstPunch.Close
5  Set rstPunch = Nothing
5  End If

5  Call Unload(Me)

5  Exit Sub

Oups:

5  wOups "frmChoixDateImpressionFacturation", "ImprimerListePunch", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerPrixCoutant()

 On Error GoTo Oups

 Dim dblTotal As Double
 Dim sProjet As String
 Dim rstDS As ADODB.Recordset

 If optChoix(I_OPT_2_DATES).Value = True Then
 If mskDateDebut.Text <> "" Then
 If mskDateFin.Text <> "" Then
 If ValiderDate(mskDateDebut.Text) = True Then
 If ValiderDate(mskDateFin.Text) = True Then
 If mskDateDebut.Text > mskDateFin.Text Then
 Call MsgBox("La date de début doit être plus petite que la date de fin!", vbOKOnly, "Erreur")

  Exit Sub
  End If
  Else
  Call MsgBox("Date de fin non valide!", vbOKOnly, "Erreur")

  Exit Sub
  End If
  Else
  Call MsgBox("Date de début non valide!", vbOKOnly, "Erreur")

 Exit Sub
 End If
 Else
 Call MsgBox("La date de fin est obligatoire!", vbOKOnly, "Erreur")

 Exit Sub
 End If
 Else
 Call MsgBox("La date de début est obligatoire!", vbOKOnly, "Erreur")

 Exit Sub
 End If
End If

If Len(m_sNoProjSoum) =   Then
sProjet = Right$(m_sNoProjSoum, 8)
Else
 Call MsgBox("Numéro de projet non valide!", vbOKOnly, "Erreur")

 Exit Sub
 End If

 'Ce recordset ne sert absolument à rien,
 'il est seulement utiliser parce que le DR a besoin d'un DataSource pour ouvrir
Set rstDS = New ADODB.Recordset

 Call rstDS.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & m_sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)

1  Set DR_ApercuProjet.DataSource = rstDS

 DR_ApercuProjet.Sections("Section2").Controls("lblNumero").Caption = sProjet
 DR_ApercuProjet.Sections("Section2").Controls("lblClient").Caption = m_sClient
DR_ApercuProjet.Sections("Section2").Controls("lblDescription").Caption = m_sDescription

If optChoix(I_OPT_PROJET_ENTIER).Value = True Then
 DR_ApercuProjet.Sections("Section2").Controls("lblDate").Caption = ConvertDate(Date)
Else
 DR_ApercuProjet.Sections("Section2").Controls("lblDate").Caption = "Du " & mskDateDebut.Text & " au " & mskDateFin.Text
End If

Call RemplirRapportElectrique(sProjet)
Call RemplirRapportMecanique(sProjet)

If IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitElecSoum").Caption) And IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitMecSoum").Caption) Then
 dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitElecSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitMecSoum").Caption)
2  Else
 If IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitElecSoum").Caption) Then
 dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitElecSoum").Caption)
 Else
 If IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitMecSoum").Caption) Then
 dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitMecSoum").Caption)
 Else
 dblTotal = 0
 End If
3 End If
End If

DR_ApercuProjet.Sections("Section2").Controls("lblTotalForfaitSoum").Caption = Conversion(dblTotal, MODE_ARGENT)

If IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitElecProj").Caption) And IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitMecProj").Caption) Then
 dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitElecProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitMecProj").Caption)
Else
 If IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitElecProj").Caption) Then
 dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitElecProj").Caption)
 Else
 If IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitMecProj").Caption) Then
 dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitMecProj").Caption)
 Else
 dblTotal = 0
 End If
 End If
3  End If

DR_ApercuProjet.Sections("Section2").Controls("lblTotalForfaitProj").Caption = Conversion(dblTotal, MODE_ARGENT)

3  If IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTotalSoum").Caption) And IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTotalSoum").Caption) Then
 dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTotalSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTotalSoum").Caption)

DR_ApercuProjet.Sections("Section2").Controls("lblTotalHeuresSoum").Caption = dblTotal
Else
4 If Not IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTotalSoum").Caption) And Not IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTotalSoum").Caption) Then
4 DR_ApercuProjet.Sections("Section2").Controls("lblTotalHeuresSoum").Caption = "---"
4 Else
4 If Not IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTotalSoum").Caption) Then
4 DR_ApercuProjet.Sections("Section2").Controls("lblTotalHeuresSoum").Caption = DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTotalSoum").Caption
4 Else
4 DR_ApercuProjet.Sections("Section2").Controls("lblTotalHeuresSoum").Caption = DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTotalSoum").Caption
4 End If
4 End If
4 End If

4  If IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblTotalElecSoum").Caption) And IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblTotalMecSoum").Caption) Then
4  dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblTotalElecSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblTotalMecSoum").Caption)

4  DR_ApercuProjet.Sections("Section2").Controls("lblGrandTotalSoum").Caption = Conversion(dblTotal, MODE_ARGENT)
4  Else
4  If Not IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblTotalElecSoum").Caption) And Not IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblTotalMecSoum").Caption) Then
4  DR_ApercuProjet.Sections("Section2").Controls("lblGrandTotalSoum").Caption = "---"
4  Else
4  If Not IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblTotalElecSoum").Caption) Then
50 DR_ApercuProjet.Sections("Section2").Controls("lblGrandTotalSoum").Caption = DR_ApercuProjet.Sections("Section2").Controls("lblTotalMecSoum").Caption
Else
 DR_ApercuProjet.Sections("Section2").Controls("lblGrandTotalSoum").Caption = DR_ApercuProjet.Sections("Section2").Controls("lblTotalElecSoum").Caption
 End If
 End If
 End If

 If IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTotalProj").Caption) And IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTotalProj").Caption) Then
 dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTotalProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTotalProj").Caption)

 DR_ApercuProjet.Sections("Section2").Controls("lblTotalHeuresProj").Caption = dblTotal
 Else
 If Not IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTotalProj").Caption) And Not IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTotalProj").Caption) Then
 DR_ApercuProjet.Sections("Section2").Controls("lblTotalHeuresProj").Caption = "---"
5  Else
5  If Not IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTotalProj").Caption) Then
5  DR_ApercuProjet.Sections("Section2").Controls("lblTotalHeuresProj").Caption = DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTotalProj").Caption
5  Else
5  DR_ApercuProjet.Sections("Section2").Controls("lblTotalHeuresProj").Caption = DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTotalProj").Caption
5  End If
5  End If
5  End If

60 If IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblTotalElecProj").Caption) And IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblTotalMecProj").Caption) Then
  dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblTotalElecProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblTotalMecProj").Caption)

  DR_ApercuProjet.Sections("Section2").Controls("lblGrandTotalProj").Caption = Conversion(dblTotal, MODE_ARGENT)
  Else
  If Not IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblTotalElecProj").Caption) And Not IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblTotalMecProj").Caption) Then
  DR_ApercuProjet.Sections("Section2").Controls("lblGrandTotalProj").Caption = "---"
  Else
  If Not IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblTotalElecProj").Caption) Then
  DR_ApercuProjet.Sections("Section2").Controls("lblGrandTotalProj").Caption = DR_ApercuProjet.Sections("Section2").Controls("lblTotalMecProj").Caption
  Else
  DR_ApercuProjet.Sections("Section2").Controls("lblGrandTotalProj").Caption = DR_ApercuProjet.Sections("Section2").Controls("lblTotalElecProj").Caption
  End If
6  End If
6  End If

6  If DR_ApercuProjet.Sections("Section2").Controls("lblTotalForfaitSoum").Caption <> "---" And _
 DR_ApercuProjet.Sections("Section2").Controls("lblGrandTotalSoum").Caption <> "---" Then
6  dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblTotalForfaitSoum").Caption) - CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblGrandTotalSoum").Caption)
6  Else
6  dblTotal = 0
6  End If

6  DR_ApercuProjet.Sections("Section2").Controls("lblProfitSoum").Caption = Conversion(dblTotal, MODE_ARGENT)

70 If DR_ApercuProjet.Sections("Section2").Controls("lblTotalForfaitProj").Caption <> "---" And _
 DR_ApercuProjet.Sections("Section2").Controls("lblGrandTotalProj").Caption <> "---" Then
  dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblTotalForfaitProj").Caption) - CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblGrandTotalProj").Caption)
  Else
  dblTotal = 0
  End If

  DR_ApercuProjet.Sections("Section2").Controls("lblProfitProj").Caption = Conversion(dblTotal, MODE_ARGENT)

  Call DR_ApercuProjet.Show(vbModal)

  Call rstDS.Close
  Set rstDS = Nothing

  Call Unload(Me)

  Exit Sub

Oups:

  wOups "frmChoixDateImpressionFacturation", "ImprimerPrixCoutant", Err, Err.number, Err.Description
End Sub

Private Sub RemplirRapportElectrique(ByVal sProjet As String)
 
 On Error GoTo Oups

 Dim rstProjetElec As ADODB.Recordset
 Dim rstSoumElec As ADODB.Recordset
 Dim rstProjetPieces As ADODB.Recordset
 Dim dblTotal As Double
 Dim bSoumission As Boolean
 Dim iNbrePersonne As Integer
 Dim dblHebergement As Double
 Dim dblRepas As Double
 Dim dblTransport As Double
 Dim dblUniteMobile As Double
  Dim dblPrixEmballage As Double
  Dim dblTotalResteTemps As Double
  Dim dblTotalManuel As Double
  Dim dblTotalPieces As Double
 
  Set rstProjetElec = New ADODB.Recordset

  Call rstProjetElec.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = 'E" & sProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

  DR_ApercuProjet.Sections("Section2").Controls("lblProjetElec").Caption = "E" & sProjet

  If Not rstProjetElec.EOF Then
bSoumission = False

1 If Not IsNull(rstProjetElec.Fields("IDSoumission")) Then
 Set rstSoumElec = New ADODB.Recordset

 Call rstSoumElec.Open("SELECT * FROM GrbSoumissionElec WHERE IDSoumission = '" & rstProjetElec.Fields("IDSoumission") & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If Not rstSoumElec.EOF Then
 bSoumission = True
 Else
 Call rstSoumElec.Close
 Set rstSoumElec = Nothing
 End If
 End If

 If bSoumission = True Then
 If Not IsNull(rstSoumElec.Fields("MontantForfait")) Then
 DR_ApercuProjet.Sections("Section2").Controls("lblForfaitElecSoum").Caption = Conversion(rstSoumElec.Fields("MontantForfait"), MODE_ARGENT)
 Else
 DR_ApercuProjet.Sections("Section2").Controls("lblForfaitElecSoum").Caption = Conversion("0", MODE_ARGENT)
 End If

 If Not IsNull(rstSoumElec.Fields("TempsDessin")) Then
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecDessinSoum").Caption = rstSoumElec.Fields("TempsDessin")
1  DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecDessinSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsDessin")) * CDbl(rstSoumElec.Fields("TauxDessin")), MODE_ARGENT)
 Else
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecDessinSoum").Caption = "0"
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecDessinSoum").Caption = Conversion("0", MODE_ARGENT)
 End If

 If Not IsNull(rstSoumElec.Fields("TempsFabrication")) Then
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFabricationSoum").Caption = rstSoumElec.Fields("TempsFabrication")
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFabricationSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsFabrication")) * CDbl(rstSoumElec.Fields("TauxFabrication")), MODE_ARGENT)
 Else
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFabricationSoum").Caption = "0"
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFabricationSoum").Caption = Conversion("0", MODE_ARGENT)
 End If

 If Not IsNull(rstSoumElec.Fields("TempsAssemblage")) Then
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecAssemblageSoum").Caption = rstSoumElec.Fields("TempsAssemblage")
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecAssemblageSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsAssemblage")) * CDbl(rstSoumElec.Fields("TauxAssemblage")), MODE_ARGENT)
 Else
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecAssemblageSoum").Caption = "0"
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecAssemblageSoum").Caption = Conversion("0", MODE_ARGENT)
 End If

 If Not IsNull(rstSoumElec.Fields("TempsProgInterface")) Then
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgInterfaceSoum").Caption = rstSoumElec.Fields("TempsProgInterface")
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgInterfaceSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsProgInterface")) * CDbl(rstSoumElec.Fields("TauxProgInterface")), MODE_ARGENT)
Else
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgInterfaceSoum").Caption = "0"
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgInterfaceSoum").Caption = Conversion("0", MODE_ARGENT)
 End If

 If Not IsNull(rstSoumElec.Fields("TempsProgAutomate")) Then
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgAutomateSoum").Caption = rstSoumElec.Fields("TempsProgAutomate")
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgAutomateSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsProgAutomate")) * CDbl(rstSoumElec.Fields("TauxProgAutomate")), MODE_ARGENT)
 Else
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgAutomateSoum").Caption = "0"
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgAutomateSoum").Caption = Conversion("0", MODE_ARGENT)
 End If

 If Not IsNull(rstSoumElec.Fields("TempsProgRobot")) Then
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgRobotSoum").Caption = rstSoumElec.Fields("TempsProgRobot")
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgRobotSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsProgRobot")) * CDbl(rstSoumElec.Fields("TauxProgRobot")), MODE_ARGENT)
 Else
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgRobotSoum").Caption = "0"
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgRobotSoum").Caption = Conversion("0", MODE_ARGENT)
 End If

 If Not IsNull(rstSoumElec.Fields("TempsVision")) Then
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecVisionSoum").Caption = rstSoumElec.Fields("TempsVision")
4 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecVisionSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsVision")) * CDbl(rstSoumElec.Fields("TauxVision")), MODE_ARGENT)
4 Else
4 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecVisionSoum").Caption = "0"
4 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecVisionSoum").Caption = Conversion("0", MODE_ARGENT)
4 End If

4 If Not IsNull(rstSoumElec.Fields("TempsTest")) Then
4 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTestSoum").Caption = rstSoumElec.Fields("TempsTest")
4 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecTestSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsTest")) * CDbl(rstSoumElec.Fields("TauxTest")), MODE_ARGENT)
4 Else
4 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTestSoum").Caption = rstSoumElec.Fields("TempsTest")
4 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecTestSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsTest")) * CDbl(rstSoumElec.Fields("TauxTest")), MODE_ARGENT)
4  End If

4  If Not IsNull(rstSoumElec.Fields("TempsInstallation")) Then
4  DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecInstallationSoum").Caption = rstSoumElec.Fields("TempsInstallation")
4  DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecInstallationSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsInstallation")) * CDbl(rstSoumElec.Fields("TauxInstallation")), MODE_ARGENT)
4  Else
4  DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecInstallationSoum").Caption = "0"
4  DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecInstallationSoum").Caption = Conversion("0", MODE_ARGENT)
4  End If

50 If Not IsNull(rstSoumElec.Fields("TempsMiseService")) Then
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecMiseServiceSoum").Caption = rstSoumElec.Fields("TempsMiseService")
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecMiseServiceSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsMiseService")) * CDbl(rstSoumElec.Fields("TauxMiseService")), MODE_ARGENT)
 Else
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecMiseServiceSoum").Caption = "0"
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecMiseServiceSoum").Caption = Conversion("0", MODE_ARGENT)
 End If

 If Not IsNull(rstSoumElec.Fields("TempsFormation")) Then
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFormationSoum").Caption = rstSoumElec.Fields("TempsFormation")
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFormationSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsFormation")) * CDbl(rstSoumElec.Fields("TauxFormation")), MODE_ARGENT)
 Else
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFormationSoum").Caption = "0"
5  DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFormationSoum").Caption = Conversion("0", MODE_ARGENT)
5  End If

5  If Not IsNull(rstSoumElec.Fields("TempsGestion")) Then
5  DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecGestionSoum").Caption = rstSoumElec.Fields("TempsGestion")
5  DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecGestionSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsGestion")) * CDbl(rstSoumElec.Fields("TauxGestion")), MODE_ARGENT)
5  Else
5  DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecGestionSoum").Caption = "0"
5  DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecGestionSoum").Caption = Conversion("0", MODE_ARGENT)
60 End If

  If Not IsNull(rstSoumElec.Fields("TempsShipping")) Then
  DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecShippingSoum").Caption = rstSoumElec.Fields("TempsShipping")
  DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecShippingSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsShipping")) * CDbl(rstSoumElec.Fields("TauxShipping")), MODE_ARGENT)
  Else
  DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecShippingSoum").Caption = "0"
  DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecShippingSoum").Caption = Conversion("0", MODE_ARGENT)
  End If

  DR_ApercuProjet.Sections("Section2").Controls("lblPiecesElecSoum").Caption = Conversion(rstSoumElec.Fields("total_piece"), MODE_ARGENT)
  DR_ApercuProjet.Sections("Section2").Controls("lblImprevuElecSoum").Caption = Conversion(rstSoumElec.Fields("total_imprevue"), MODE_ARGENT)

  If Not IsNull(rstSoumElec.Fields("NbrePersonne")) Then
  iNbrePersonne = rstSoumElec.Fields("NbrePersonne")
6  Else
6  iNbrePersonne = 0
6  End If
 
6  Do While iNbrePersonne > 0
6  If iNbrePersonne >= 2 Then
6  dblHebergement = dblHebergement + rstSoumElec.Fields("TempsHebergement") * rstSoumElec.Fields("TauxHebergement2")
 
6  iNbrePersonne = iNbrePersonne - 2
6  Else
70 dblHebergement = dblHebergement + rstSoumElec.Fields("TempsHebergement") * rstSoumElec.Fields("TauxHebergement1")
 
  iNbrePersonne = iNbrePersonne - 1
  End If
  Loop
 
  If Not IsNull(rstSoumElec.Fields("TempsRepas")) Then
  dblRepas = CDbl(rstSoumElec.Fields("TempsRepas")) * CDbl(rstSoumElec.Fields("TauxRepas")) * CDbl(rstSoumElec.Fields("NbrePersonne"))
  Else
  dblRepas = 0
  End If

  If Not IsNull(rstSoumElec.Fields("TempsTransport")) Then
  dblTransport = CDbl(rstSoumElec.Fields("TempsTransport")) * CDbl(rstSoumElec.Fields("TauxTransport"))
  Else
   dblTransport = 0
   End If

7  If Not IsNull(rstSoumElec.Fields("TempsUniteMobile")) Then
7  dblUniteMobile = CDbl(rstSoumElec.Fields("TempsUniteMobile")) * CDbl(rstSoumElec.Fields("TauxUniteMobile"))
7  Else
7  dblUniteMobile = 0
7  End If

7  If Not IsNull(rstSoumElec.Fields("PrixEmballage")) Then
80 dblPrixEmballage = CDbl(rstSoumElec.Fields("PrixEmballage"))
  Else
  dblPrixEmballage = 0
  End If
 
  dblTotalResteTemps = dblHebergement + dblRepas + dblTransport + dblUniteMobile + dblPrixEmballage
  
  If IsNumeric(rstSoumElec.Fields("total_manuel")) Then
  dblTotalManuel = CDbl(rstSoumElec.Fields("total_manuel"))
  Else
  dblTotalManuel = 0
  End If

  DR_ApercuProjet.Sections("Section2").Controls("lblAutresElecSoum").Caption = Conversion(dblTotalResteTemps + dblTotalManuel, MODE_ARGENT)

  Call rstSoumElec.Close
   Set rstSoumElec = Nothing
   Else
   DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecDessinSoum").Caption = "---"
   DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecDessinSoum").Caption = "---"

8  DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFabricationSoum").Caption = "---"
8  DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFabricationSoum").Caption = "---"

8  DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecAssemblageSoum").Caption = "---"
8  DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecAssemblageSoum").Caption = "---"

90 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgInterfaceSoum").Caption = "---"
  DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgInterfaceSoum").Caption = "---"

  DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgAutomateSoum").Caption = "---"
  DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgAutomateSoum").Caption = "---"

  DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgRobotSoum").Caption = "---"
  DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgRobotSoum").Caption = "---"

  DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecVisionSoum").Caption = "---"
  DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecVisionSoum").Caption = "---"

  DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTestSoum").Caption = "---"
  DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecTestSoum").Caption = "---"

  DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecInstallationSoum").Caption = "---"
  DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecInstallationSoum").Caption = "---"

 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecMiseServiceSoum").Caption = "---"
   DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecMiseServiceSoum").Caption = "---"

 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFormationSoum").Caption = "---"
   DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFormationSoum").Caption = "---"

 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecGestionSoum").Caption = "---"
   DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecGestionSoum").Caption = "---"

 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecGestionSoum").Caption = "---"
9  DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecGestionSoum").Caption = "---"

 DR_ApercuProjet.Sections("Section2").Controls("lblPiecesElecSoum").Caption = "---"
10 DR_ApercuProjet.Sections("Section2").Controls("lblImprevuElecSoum").Caption = "---"
1 DR_ApercuProjet.Sections("Section2").Controls("lblAutresElecSoum").Caption = "---"
1End If

 Call RemplirTempsReelsElec("E" & sProjet)

1 If Not IsNull(rstProjetElec.Fields("MontantForfait")) Then
 DR_ApercuProjet.Sections("Section2").Controls("lblForfaitElecProj").Caption = Conversion(rstProjetElec.Fields("MontantForfait"), MODE_ARGENT)
1Else
 DR_ApercuProjet.Sections("Section2").Controls("lblForfaitElecProj").Caption = Conversion("0", MODE_ARGENT)
1End If

 If Not IsNull(rstProjetElec.Fields("TauxDessin")) Then
1 'DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecDessinProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecDessinProj").Caption) * CDbl(rstProjetElec.Fields("TauxDessin")), MODE_ARGENT)
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecDessinProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecDessinProj").Caption) * CDbl(50), MODE_ARGENT)
10  Else
10  DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecDessinProj").Caption = Conversion(0, MODE_ARGENT)
10  End If

10  If Not IsNull(rstProjetElec.Fields("TauxFabrication")) Then
10  'DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFabricationProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFabricationProj").Caption) * CDbl(rstProjetElec.Fields("TauxFabrication")), MODE_ARGENT)
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFabricationProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFabricationProj").Caption) * CDbl(50), MODE_ARGENT)
10  Else
10  DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFabricationProj").Caption = Conversion(0, MODE_ARGENT)
10  End If

110 If Not IsNull(rstProjetElec.Fields("TauxAssemblage")) Then
11 'DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecAssemblageProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecAssemblageProj").Caption) * CDbl(rstProjetElec.Fields("TauxAssemblage")), MODE_ARGENT)
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecAssemblageProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecAssemblageProj").Caption) * CDbl(50), MODE_ARGENT)
1 Else
1 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecAssemblageProj").Caption = Conversion(0, MODE_ARGENT)
1 End If

1 If Not IsNull(rstProjetElec.Fields("TauxProgInterface")) Then
1 'DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgInterfaceProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgInterfaceProj").Caption) * CDbl(rstProjetElec.Fields("TauxProgInterface")), MODE_ARGENT)
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgInterfaceProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgInterfaceProj").Caption) * CDbl(50), MODE_ARGENT)
1 Else
1 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgInterfaceProj").Caption = Conversion(0, MODE_ARGENT)
1 End If

1 If Not IsNull(rstProjetElec.Fields("TauxProgAutomate")) Then
1 'DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgAutomateProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgAutomateProj").Caption) * CDbl(rstProjetElec.Fields("TauxProgAutomate")), MODE_ARGENT)
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgAutomateProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgAutomateProj").Caption) * CDbl(50), MODE_ARGENT)
11  Else
1 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgAutomateProj").Caption = Conversion(0, MODE_ARGENT)
 End If

1 If Not IsNull(rstProjetElec.Fields("TauxProgRobot")) Then
 'DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgRobotProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgRobotProj").Caption) * CDbl(rstProjetElec.Fields("TauxProgRobot")), MODE_ARGENT)
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgRobotProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgRobotProj").Caption) * CDbl(50), MODE_ARGENT)
1 Else
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgRobotProj").Caption = Conversion(0, MODE_ARGENT)
11  End If

 If Not IsNull(rstProjetElec.Fields("TauxVision")) Then
 'DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecVisionProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecVisionProj").Caption) * CDbl(rstProjetElec.Fields("TauxVision")), MODE_ARGENT)
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecVisionProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecVisionProj").Caption) * CDbl(50), MODE_ARGENT)
1 Else
1 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecVisionProj").Caption = Conversion(0, MODE_ARGENT)
1 End If

1 If Not IsNull(rstProjetElec.Fields("TauxTest")) Then
1 'DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecTestProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTestProj").Caption) * CDbl(rstProjetElec.Fields("TauxTest")), MODE_ARGENT)
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecTestProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTestProj").Caption) * CDbl(50), MODE_ARGENT)
1 Else
1 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecTestProj").Caption = Conversion(0, MODE_ARGENT)
1 End If

1 If Not IsNull(rstProjetElec.Fields("TauxInstallation")) Then
1 'DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecInstallationProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecInstallationProj").Caption) * CDbl(rstProjetElec.Fields("TauxInstallation")), MODE_ARGENT)
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecInstallationProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecInstallationProj").Caption) * CDbl(50), MODE_ARGENT)
12  Else
1 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecInstallationProj").Caption = Conversion(0, MODE_ARGENT)
12  End If

1 If Not IsNull(rstProjetElec.Fields("TauxMiseService")) Then
1 'DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecMiseServiceProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecMiseServiceProj").Caption) * CDbl(rstProjetElec.Fields("TauxMiseService")), MODE_ARGENT)
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecMiseServiceProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecMiseServiceProj").Caption) * CDbl(50), MODE_ARGENT)
1 Else
1 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecMiseServiceProj").Caption = Conversion(0, MODE_ARGENT)
1 End If

130 If Not IsNull(rstProjetElec.Fields("TauxFormation")) Then
13 'DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFormationProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFormationProj").Caption) * CDbl(rstProjetElec.Fields("TauxFormation")), MODE_ARGENT)
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFormationProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFormationProj").Caption) * CDbl(50), MODE_ARGENT)
1 Else
1 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFormationProj").Caption = Conversion(0, MODE_ARGENT)
1 End If

1 If Not IsNull(rstProjetElec.Fields("TauxGestion")) Then
1 'DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecGestionProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecGestionProj").Caption) * CDbl(rstProjetElec.Fields("TauxGestion")), MODE_ARGENT)
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecGestionProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecGestionProj").Caption) * CDbl(50), MODE_ARGENT)
1 Else
1 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecGestionProj").Caption = Conversion(0, MODE_ARGENT)
1 End If

1 If Not IsNull(rstProjetElec.Fields("TauxShipping")) Then
1 'DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecShippingProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecShippingProj").Caption) * CDbl(rstProjetElec.Fields("TauxShipping")), MODE_ARGENT)
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecShippingProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecShippingProj").Caption) * CDbl(50), MODE_ARGENT)
13  Else
1 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecShippingProj").Caption = Conversion(0, MODE_ARGENT)
13  End If
 '''''''''''''''''''''''''''''''''
 '''''''''''''''''''''''''''''''''
 ' ajout pour développement expérimental

 If Not IsNull(rstProjetElec.Fields("TauxGestion")) Then
 'DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecRechercheProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecRechercheProj").Caption) * CDbl(rstProjetElec.Fields("TauxGestion")), MODE_ARGENT)
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecRechercheProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecRechercheProj").Caption) * CDbl(50), MODE_ARGENT)
 Else
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecRechercheProj").Caption = Conversion(0, MODE_ARGENT)
 End If
 ''''''''''''''''''''''''''''''''''''''
 ''''''''''''''''''''''''




1 Set rstProjetPieces = New ADODB.Recordset

1 If optChoix(I_OPT_PROJET_ENTIER).Value = True Then
1 Call rstProjetPieces.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = 'E" & sProjet & "'", g_connData, adOpenForwardOnly, adLockReadOnly)
1Else
1 If Right(sProjet, 2) = "99" Then
1 Call rstProjetPieces.Open("SELECT * FROM GrbProjet_Pieces WHERE Left(IDProjet, 6) = '" & Left$("E" & sProjet, 6) & "' AND DateRéception BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND PieceExtraNonChargeable = False AND PieceExtraChargeable = False", g_connData, adOpenForwardOnly, adLockReadOnly)
14 Else
14 Call rstProjetPieces.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = 'E" & sProjet & "' AND DateRéception BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "'", g_connData, adOpenForwardOnly, adLockReadOnly)
14 End If
14 End If

14 Do While Not rstProjetPieces.EOF
14 If Trim(rstProjetPieces.Fields("Prix_total")) <> vbNullString Then
 'On additionne le prix total
14 dblTotalPieces = dblTotalPieces + CDbl(rstProjetPieces.Fields("Prix_total")) - CDbl(rstProjetPieces.Fields("Profit_Argent"))
14 End If

14 Call rstProjetPieces.MoveNext
14 Loop

14 Call rstProjetPieces.Close
14  Set rstProjetPieces = Nothing

14  DR_ApercuProjet.Sections("Section2").Controls("lblPiecesElecProj").Caption = Conversion(dblTotalPieces, MODE_ARGENT)
'  DR_ApercuProjet.Sections("Section2").Controls("lblImprevuElecProj").Caption = Conversion(rstProjetElec.Fields("total_imprevue"), MODE_ARGENT)
14  DR_ApercuProjet.Sections("Section2").Controls("lblImprevuElecProj").Caption = Conversion(0, MODE_ARGENT)

14  If Not IsNull(rstProjetElec.Fields("PrixEmballage")) Then
14  dblPrixEmballage = CDbl(rstProjetElec.Fields("PrixEmballage"))
14  Else
14  dblPrixEmballage = 0
14  End If
 
150 dblTotalResteTemps = dblPrixEmballage
  
15 If IsNumeric(rstProjetElec.Fields("total_manuel")) Then
 dblTotalManuel = CDbl(rstProjetElec.Fields("total_manuel"))
1 Else
 dblTotalManuel = 0
 End If

 DR_ApercuProjet.Sections("Section2").Controls("lblAutresElecProj").Caption = Conversion(dblTotalResteTemps + dblTotalManuel, MODE_ARGENT)

 'Calcul des totaux

 'Total des temps
 If DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecDessinSoum").Caption <> "---" Then
 dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecDessinSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFabricationSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecAssemblageSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgInterfaceSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgAutomateSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgRobotSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecVisionSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTestSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecInstallationSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecMiseServiceSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFormationSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecGestionSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecShippingSoum").Caption)
 
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTotalSoum").Caption = dblTotal
 Else
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTotalSoum").Caption = "---"
15  End If

15  If DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecDessinSoum").Caption <> "---" Then
15  dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecDessinSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFabricationSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecAssemblageSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgInterfaceSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgAutomateSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgRobotSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecVisionSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecTestSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecInstallationSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecMiseServiceSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFormationSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecGestionSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecShippingSoum").Caption)
 
15  DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecTotalSoum").Caption = Conversion(dblTotal, MODE_ARGENT)
15  Else
15  DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecTotalSoum").Caption = "---"
15  End If
''''''''''''''
' calcul du total des heures
'''''''''''''''''''''
15  dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecDessinProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFabricationProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecAssemblageProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgInterfaceProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgAutomateProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgRobotProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecVisionProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTestProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecInstallationProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecMiseServiceProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFormationProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecGestionProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecShippingProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecRechercheProj").Caption)
 
160 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTotalProj").Caption = dblTotal
''''''''''''''''''''''''''''''
'Calcul du total de l'argent
'''''''''''''''''''''''''''''''
16dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecDessinProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFabricationProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecAssemblageProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgInterfaceProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgAutomateProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgRobotProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecVisionProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecTestProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecInstallationProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecMiseServiceProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFormationProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecGestionProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecShippingProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecRechercheProj").Caption)
 
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecTotalProj").Caption = Conversion(dblTotal, MODE_ARGENT)

 'Calcul des prix totaux
 If DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecTotalSoum").Caption <> "---" Then
 dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecTotalSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblPiecesElecSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblImprevuElecSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblAutresElecSoum").Caption)

 DR_ApercuProjet.Sections("Section2").Controls("lblTotalElecSoum").Caption = Conversion(dblTotal, MODE_ARGENT)
 Else
 DR_ApercuProjet.Sections("Section2").Controls("lblTotalElecSoum").Caption = "---"
 End If

 dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecTotalProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblPiecesElecProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblImprevuElecProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblAutresElecProj").Caption)

 DR_ApercuProjet.Sections("Section2").Controls("lblTotalElecProj").Caption = Conversion(dblTotal, MODE_ARGENT)
1  End If

16  Call rstProjetElec.Close
16  Set rstProjetElec = Nothing

167Exit Sub

Oups:

16  wOups "frmChoixDateImpressionFacturation", "RemplirRapportElectrique", Err, Err.number, Err.Description
End Sub

Private Sub RemplirRapportMecanique(ByVal sProjet As String)
 
 On Error GoTo Oups

 Dim rstProjetMec As ADODB.Recordset
 Dim rstSoumMec As ADODB.Recordset
 Dim rstProjetPieces As ADODB.Recordset
 Dim dblTotal As Double
 Dim bSoumission As Boolean
 Dim iNbrePersonne As Integer
 Dim dblHebergement As Double
 Dim dblRepas As Double
 Dim dblTransport As Double
 Dim dblUniteMobile As Double
  Dim dblPrixEmballage As Double
  Dim dblTotalResteTemps As Double
  Dim dblTotalManuel As Double
  Dim dblTotalPieces As Double
 
  Set rstProjetMec = New ADODB.Recordset

  Call rstProjetMec.Open("SELECT * FROM GrbProjetMec WHERE IDProjet = 'M" & sProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

  DR_ApercuProjet.Sections("Section2").Controls("lblProjetMec").Caption = "M" & sProjet

  If Not rstProjetMec.EOF Then
bSoumission = False

1 If Not IsNull(rstProjetMec.Fields("IDSoumission")) Then
 Set rstSoumMec = New ADODB.Recordset

 Call rstSoumMec.Open("SELECT * FROM GrbSoumissionMec WHERE IDSoumission = '" & rstProjetMec.Fields("IDSoumission") & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If Not rstSoumMec.EOF Then
 bSoumission = True
 Else
 Call rstSoumMec.Close
 Set rstSoumMec = Nothing
 End If
 End If

 If bSoumission = True Then
 If Not IsNull(rstSoumMec.Fields("MontantForfait")) Then
 DR_ApercuProjet.Sections("Section2").Controls("lblForfaitMecSoum").Caption = Conversion(rstSoumMec.Fields("MontantForfait"), MODE_ARGENT)
 Else
 DR_ApercuProjet.Sections("Section2").Controls("lblForfaitMecSoum").Caption = Conversion("0", MODE_ARGENT)
 End If

 If Not IsNull(rstSoumMec.Fields("TempsDessin")) Then
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecDessinSoum").Caption = rstSoumMec.Fields("TempsDessin")
1  DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecDessinSoum").Caption = Conversion(CDbl(rstSoumMec.Fields("TempsDessin")) * CDbl(rstSoumMec.Fields("TauxDessin")), MODE_ARGENT)
 Else
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecDessinSoum").Caption = "0"
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecDessinSoum").Caption = Conversion("0", MODE_ARGENT)
 End If

 If Not IsNull(rstSoumMec.Fields("TempsCoupe")) Then
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecCoupeSoum").Caption = rstSoumMec.Fields("TempsCoupe")
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecCoupeSoum").Caption = Conversion(CDbl(rstSoumMec.Fields("TempsCoupe")) * CDbl(rstSoumMec.Fields("TauxCoupe")), MODE_ARGENT)
 Else
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecCoupeSoum").Caption = "0"
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecCoupeSoum").Caption = Conversion("0", MODE_ARGENT)
 End If

 If Not IsNull(rstSoumMec.Fields("TempsMachinage")) Then
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecMachinageSoum").Caption = rstSoumMec.Fields("TempsMachinage")
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecMachinageSoum").Caption = Conversion(CDbl(rstSoumMec.Fields("TempsMachinage")) * CDbl(rstSoumMec.Fields("TauxMachinage")), MODE_ARGENT)
 Else
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecMachinageSoum").Caption = "0"
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecMachinageSoum").Caption = Conversion("0", MODE_ARGENT)
 End If

 If Not IsNull(rstSoumMec.Fields("TempsSoudure")) Then
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecSoudureSoum").Caption = rstSoumMec.Fields("TempsSoudure")
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecSoudureSoum").Caption = Conversion(CDbl(rstSoumMec.Fields("TempsSoudure")) * CDbl(rstSoumMec.Fields("TauxSoudure")), MODE_ARGENT)
Else
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecSoudureSoum").Caption = "0"
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecSoudureSoum").Caption = Conversion("0", MODE_ARGENT)
 End If

 If Not IsNull(rstSoumMec.Fields("TempsAssemblage")) Then
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecAssemblageSoum").Caption = rstSoumMec.Fields("TempsAssemblage")
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecAssemblageSoum").Caption = Conversion(CDbl(rstSoumMec.Fields("TempsAssemblage")) * CDbl(rstSoumMec.Fields("TauxAssemblage")), MODE_ARGENT)
 Else
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecAssemblageSoum").Caption = "0"
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecAssemblageSoum").Caption = Conversion("0", MODE_ARGENT)
 End If

 If Not IsNull(rstSoumMec.Fields("TempsPeinture")) Then
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecPeintureSoum").Caption = rstSoumMec.Fields("TempsPeinture")
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecPeintureSoum").Caption = Conversion(CDbl(rstSoumMec.Fields("TempsPeinture")) * CDbl(rstSoumMec.Fields("TauxPeinture")), MODE_ARGENT)
 Else
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecPeintureSoum").Caption = "0"
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecPeintureSoum").Caption = Conversion("0", MODE_ARGENT)
 End If

 If Not IsNull(rstSoumMec.Fields("TempsTest")) Then
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTestSoum").Caption = rstSoumMec.Fields("TempsTest")
4 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecTestSoum").Caption = Conversion(CDbl(rstSoumMec.Fields("TempsTest")) * CDbl(rstSoumMec.Fields("TauxTest")), MODE_ARGENT)
4 Else
4 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTestSoum").Caption = "0"
4 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecTestSoum").Caption = Conversion("0", MODE_ARGENT)
4 End If

4 If Not IsNull(rstSoumMec.Fields("TempsInstallation")) Then
4 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecInstallationSoum").Caption = rstSoumMec.Fields("TempsInstallation")
4 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecInstallationSoum").Caption = Conversion(CDbl(rstSoumMec.Fields("TempsInstallation")) * CDbl(rstSoumMec.Fields("TauxInstallation")), MODE_ARGENT)
4 Else
4 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecInstallationSoum").Caption = "0"
4 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecInstallationSoum").Caption = Conversion("0", MODE_ARGENT)
4  End If

4  If Not IsNull(rstSoumMec.Fields("TempsFormation")) Then
4  DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecFormationSoum").Caption = rstSoumMec.Fields("TempsFormation")
4  DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecFormationSoum").Caption = Conversion(CDbl(rstSoumMec.Fields("TempsFormation")) * CDbl(rstSoumMec.Fields("TauxFormation")), MODE_ARGENT)
4  Else
4  DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecFormationSoum").Caption = "0"
4  DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecFormationSoum").Caption = Conversion("0", MODE_ARGENT)
4  End If

50 If Not IsNull(rstSoumMec.Fields("TempsGestion")) Then
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecGestionSoum").Caption = rstSoumMec.Fields("TempsGestion")
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecGestionSoum").Caption = Conversion(CDbl(rstSoumMec.Fields("TempsGestion")) * CDbl(rstSoumMec.Fields("TauxGestion")), MODE_ARGENT)
 Else
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecGestionSoum").Caption = "0"
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecGestionSoum").Caption = Conversion("0", MODE_ARGENT)
 End If

 If Not IsNull(rstSoumMec.Fields("TempsShipping")) Then
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecShippingSoum").Caption = rstSoumMec.Fields("TempsShipping")
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecShippingSoum").Caption = Conversion(CDbl(rstSoumMec.Fields("TempsShipping")) * CDbl(rstSoumMec.Fields("TauxShipping")), MODE_ARGENT)
 Else
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecShippingSoum").Caption = "0"
5  DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecShippingSoum").Caption = Conversion("0", MODE_ARGENT)
5  End If

5  DR_ApercuProjet.Sections("Section2").Controls("lblPiecesMecSoum").Caption = Conversion(rstSoumMec.Fields("total_piece"), MODE_ARGENT)
5  DR_ApercuProjet.Sections("Section2").Controls("lblImprevuMecSoum").Caption = Conversion(rstSoumMec.Fields("total_imprevue"), MODE_ARGENT)

5  iNbrePersonne = rstSoumMec.Fields("NbrePersonne")
 
5  Do While iNbrePersonne > 0
5  If iNbrePersonne >= 2 Then
5  dblHebergement = dblHebergement + rstSoumMec.Fields("TempsHebergement") * rstSoumMec.Fields("TauxHebergement2")
 
60 iNbrePersonne = iNbrePersonne - 2
  Else
  dblHebergement = dblHebergement + rstSoumMec.Fields("TempsHebergement") * rstSoumMec.Fields("TauxHebergement1")
 
  iNbrePersonne = iNbrePersonne - 1
  End If
  Loop
 
  dblRepas = CDbl(rstSoumMec.Fields("TempsRepas")) * CDbl(rstSoumMec.Fields("TauxRepas")) * CDbl(rstSoumMec.Fields("NbrePersonne"))
  dblTransport = CDbl(rstSoumMec.Fields("TempsTransport")) * CDbl(rstSoumMec.Fields("TauxTransport"))
  dblUniteMobile = CDbl(rstSoumMec.Fields("TempsUniteMobile")) * CDbl(rstSoumMec.Fields("TauxUniteMobile"))

  If IsNumeric(rstSoumMec.Fields("PrixEmballage")) Then
  dblPrixEmballage = CDbl(rstSoumMec.Fields("PrixEmballage"))
  Else
6  dblPrixEmballage = 0
6  End If
 
6  dblTotalResteTemps = dblHebergement + dblRepas + dblTransport + dblUniteMobile + dblPrixEmballage
  
6  If IsNumeric(rstSoumMec.Fields("total_manuel")) Then
6  dblTotalManuel = CDbl(rstSoumMec.Fields("total_manuel"))
6  Else
6  dblTotalManuel = 0
6  End If

70 DR_ApercuProjet.Sections("Section2").Controls("lblAutresMecSoum").Caption = Conversion(dblTotalResteTemps + dblTotalManuel, MODE_ARGENT)

  Call rstSoumMec.Close
  Set rstSoumMec = Nothing
  Else
  DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecDessinSoum").Caption = "---"
  DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecDessinSoum").Caption = "---"

  DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecCoupeSoum").Caption = "---"
  DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecCoupeSoum").Caption = "---"

  DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecMachinageSoum").Caption = "---"
  DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecMachinageSoum").Caption = "---"

  DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecSoudureSoum").Caption = "---"
  DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecSoudureSoum").Caption = "---"

   DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecAssemblageSoum").Caption = "---"
   DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecAssemblageSoum").Caption = "---"

7  DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecPeintureSoum").Caption = "---"
7  DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecPeintureSoum").Caption = "---"

7  DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTestSoum").Caption = "---"
7  DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecTestSoum").Caption = "---"

7  DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecInstallationSoum").Caption = "---"
7  DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecInstallationSoum").Caption = "---"

80 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecFormationSoum").Caption = "---"
  DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecFormationSoum").Caption = "---"

  DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecGestionSoum").Caption = "---"
  DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecGestionSoum").Caption = "---"

  DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecShippingSoum").Caption = "---"
  DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecShippingSoum").Caption = "---"

  DR_ApercuProjet.Sections("Section2").Controls("lblPiecesMecSoum").Caption = "---"
  DR_ApercuProjet.Sections("Section2").Controls("lblImprevuMecSoum").Caption = "---"
  DR_ApercuProjet.Sections("Section2").Controls("lblAutresMecSoum").Caption = "---"
  End If

  Call RemplirTempsReelsMec("M" & sProjet)

'''''''''''''''''''
'''', calcul de l'argent


  If Not IsNull(rstProjetMec.Fields("MontantForfait")) Then
   DR_ApercuProjet.Sections("Section2").Controls("lblForfaitMecProj").Caption = Conversion(rstProjetMec.Fields("MontantForfait"), MODE_ARGENT)
   Else
   DR_ApercuProjet.Sections("Section2").Controls("lblForfaitMecProj").Caption = Conversion("0", MODE_ARGENT)
   End If

8  If Not IsNull(rstProjetMec.Fields("TauxDessin")) Then
8  'DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecDessinProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecDessinProj").Caption) * CDbl(rstProjetMec.Fields("TauxDessin")), MODE_ARGENT)
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecDessinProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecDessinProj").Caption) * CDbl(50), MODE_ARGENT)
8  Else
8  DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecDessinProj").Caption = Conversion("0", MODE_ARGENT)
90 End If

  If Not IsNull(rstProjetMec.Fields("TauxCoupe")) Then
  'DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecCoupeProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecCoupeProj").Caption) * CDbl(rstProjetMec.Fields("TauxCoupe")), MODE_ARGENT)
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecCoupeProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecCoupeProj").Caption) * CDbl(50), MODE_ARGENT)
  Else
  DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecCoupeProj").Caption = Conversion("0", MODE_ARGENT)
  End If

  If Not IsNull(rstProjetMec.Fields("TauxMachinage")) Then
  'DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecMachinageProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecMachinageProj").Caption) * CDbl(rstProjetMec.Fields("TauxMachinage")), MODE_ARGENT)
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecMachinageProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecMachinageProj").Caption) * CDbl(50), MODE_ARGENT)
  Else
  DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecMachinageProj").Caption = Conversion("0", MODE_ARGENT)
  End If

  If Not IsNull(rstProjetMec.Fields("TauxSoudure")) Then
 'DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecSoudureProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecSoudureProj").Caption) * CDbl(rstProjetMec.Fields("TauxSoudure")), MODE_ARGENT)
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecSoudureProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecSoudureProj").Caption) * CDbl(50), MODE_ARGENT)
   Else
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecSoudureProj").Caption = Conversion("0", MODE_ARGENT)
   End If

 If Not IsNull(rstProjetMec.Fields("TauxAssemblage")) Then
   'DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecAssemblageProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecAssemblageProj").Caption) * CDbl(rstProjetMec.Fields("TauxAssemblage")), MODE_ARGENT)
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecAssemblageProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecAssemblageProj").Caption) * CDbl(50), MODE_ARGENT)
 Else
9  DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecAssemblageProj").Caption = Conversion("0", MODE_ARGENT)
 End If

10 If Not IsNull(rstProjetMec.Fields("TauxPeinture")) Then
1 'DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecPeintureProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecPeintureProj").Caption) * CDbl(rstProjetMec.Fields("TauxPeinture")), MODE_ARGENT)
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecPeintureProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecPeintureProj").Caption) * CDbl(50), MODE_ARGENT)
1Else
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecPeintureProj").Caption = Conversion("0", MODE_ARGENT)
1End If

 If Not IsNull(rstProjetMec.Fields("TauxTest")) Then
1 'DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecTestProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTestProj").Caption) * CDbl(rstProjetMec.Fields("TauxTest")), MODE_ARGENT)
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecTestProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTestProj").Caption) * CDbl(50), MODE_ARGENT)
 Else
1 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecTestProj").Caption = Conversion("0", MODE_ARGENT)
 End If

1 If Not IsNull(rstProjetMec.Fields("TauxInstallation")) Then
10  'DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecInstallationProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecInstallationProj").Caption) * CDbl(rstProjetMec.Fields("TauxInstallation")), MODE_ARGENT)
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecInstallationProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecInstallationProj").Caption) * CDbl(50), MODE_ARGENT)
10  Else
10  DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecInstallationProj").Caption = Conversion("0", MODE_ARGENT)
10  End If

10  If Not IsNull(rstProjetMec.Fields("TauxFormation")) Then
10  'DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecFormationProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecFormationProj").Caption) * CDbl(rstProjetMec.Fields("TauxFormation")), MODE_ARGENT)
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecFormationProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecFormationProj").Caption) * CDbl(50), MODE_ARGENT)
10  Else
10  DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecFormationProj").Caption = Conversion("0", MODE_ARGENT)
110 End If

11 If Not IsNull(rstProjetMec.Fields("TauxGestion")) Then
1 'DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecGestionProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecGestionProj").Caption) * CDbl(rstProjetMec.Fields("TauxGestion")), MODE_ARGENT)
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecGestionProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecGestionProj").Caption) * CDbl(50), MODE_ARGENT)
1 Else
1 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecGestionProj").Caption = Conversion("0", MODE_ARGENT)
1 End If

1 If Not IsNull(rstProjetMec.Fields("TauxShipping")) Then
1 'DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecShippingProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecShippingProj").Caption) * CDbl(rstProjetMec.Fields("TauxShipping")), MODE_ARGENT)
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecShippingProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecShippingProj").Caption) * CDbl(50), MODE_ARGENT)
1 Else
1 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecShippingProj").Caption = Conversion("0", MODE_ARGENT)
1 End If



''''''',''''''''''''''''''''''''
''''ajout pour develloppement experimental
'''''''''''''''''''''''''''''''''''''''
 If Not IsNull(rstProjetMec.Fields("TauxGestion")) Then
 'DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecRechercheProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecRechercheProj").Caption) * CDbl(rstProjetMec.Fields("TauxGestion")), MODE_ARGENT)
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecRechercheProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecRechercheProj").Caption) * CDbl(50), MODE_ARGENT)
 Else
 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecRechercheProj").Caption = Conversion("0", MODE_ARGENT)
 End If

'''''''''''''''''''''''''''
'''''''''''''''''''''''''



1 Set rstProjetPieces = New ADODB.Recordset

11  If optChoix(I_OPT_PROJET_ENTIER).Value = True Then
1 Call rstProjetPieces.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = 'M" & sProjet & "'", g_connData, adOpenForwardOnly, adLockReadOnly)
 Else
1 If Right(sProjet, 2) = "99" Then
 Call rstProjetPieces.Open("SELECT * FROM GrbProjet_Pieces WHERE Left(IDProjet, 6) = '" & Left$("M" & sProjet, 6) & "' AND DateRéception BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND PieceExtraNonChargeable = False AND PieceExtraChargeable = False", g_connData, adOpenForwardOnly, adLockReadOnly)
1 Else
 Call rstProjetPieces.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = 'M" & sProjet & "' AND DateRéception BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "'", g_connData, adOpenForwardOnly, adLockReadOnly)
11  End If
 End If

1 Do While Not rstProjetPieces.EOF
1 If Trim(rstProjetPieces.Fields("Prix_total")) <> vbNullString Then
 'On additionne le prix total
1 dblTotalPieces = dblTotalPieces + CDbl(rstProjetPieces.Fields("Prix_total")) - CDbl(rstProjetPieces.Fields("Profit_argent"))
1 End If

1 Call rstProjetPieces.MoveNext
1 Loop

1 Call rstProjetPieces.Close
1 Set rstProjetPieces = Nothing

1 DR_ApercuProjet.Sections("Section2").Controls("lblPiecesMecProj").Caption = Conversion(dblTotalPieces, MODE_ARGENT)
'6  DR_ApercuProjet.Sections("Section2").Controls("lblImprevuMecProj").Caption = Conversion(rstProjetMec.Fields("total_imprevue"), MODE_ARGENT)
1 DR_ApercuProjet.Sections("Section2").Controls("lblImprevuMecProj").Caption = Conversion(0, MODE_ARGENT)

1 If IsNumeric(rstProjetMec.Fields("PrixEmballage")) Then
1 dblPrixEmballage = CDbl(rstProjetMec.Fields("PrixEmballage"))
1 Else
1 dblPrixEmballage = 0
1 End If
 
12  dblTotalResteTemps = dblPrixEmballage
  
1 If IsNumeric(rstProjetMec.Fields("total_manuel")) Then
1 dblTotalManuel = CDbl(rstProjetMec.Fields("total_manuel"))
1 Else
1 dblTotalManuel = 0
13End If

1 DR_ApercuProjet.Sections("Section2").Controls("lblAutresMecProj").Caption = Conversion(dblTotalResteTemps + dblTotalManuel, MODE_ARGENT)

 'Calcul des totaux

 'Total des temps
1 If DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecDessinSoum").Caption <> "---" Then
1 dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecDessinSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecCoupeSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecMachinageSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecSoudureSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecAssemblageSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecPeintureSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTestSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecInstallationSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecFormationSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecGestionSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecShippingSoum").Caption)
 
1 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTotalSoum").Caption = dblTotal
1 Else
1 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTotalSoum").Caption = "---"
1 End If

1 If DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecDessinSoum").Caption <> "---" Then
1 dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecDessinSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecCoupeSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecMachinageSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecSoudureSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecAssemblageSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecPeintureSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecTestSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecInstallationSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecFormationSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecGestionSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecShippingSoum").Caption)
 
1 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecTotalSoum").Caption = Conversion(dblTotal, MODE_ARGENT)
13  Else
1 DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecTotalSoum").Caption = "---"
13  End If

'''''''
''' Total heure mécanique
''''''''''''''''''''''
1 dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecDessinProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecCoupeProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecMachinageProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecSoudureProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecAssemblageProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecPeintureProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTestProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecInstallationProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecFormationProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecGestionProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecShippingProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecRechercheProj").Caption)
 
 
1DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTotalProj").Caption = dblTotal
'''''''''''''''''''''''
'''' total argent mécanique
'''''''''''''''''''''

1 dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecDessinProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecCoupeProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecMachinageProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecSoudureProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecAssemblageProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecPeintureProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecTestProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecInstallationProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecFormationProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecGestionProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecShippingProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecRechercheProj").Caption)
 
 
1DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecTotalProj").Caption = Conversion(dblTotal, MODE_ARGENT)

 'Calcul des prix totaux
1 If DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecTotalSoum").Caption <> "---" Then
1 dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecTotalSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblPiecesMecSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblImprevuMecSoum").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblAutresMecSoum").Caption)

14 DR_ApercuProjet.Sections("Section2").Controls("lblTotalMecSoum").Caption = Conversion(dblTotal, MODE_ARGENT)
14 Else
14 DR_ApercuProjet.Sections("Section2").Controls("lblTotalMecSoum").Caption = "---"
14 End If

14 dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecTotalProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblPiecesMecProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblImprevuMecProj").Caption) + _
 CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblAutresMecProj").Caption)

14 DR_ApercuProjet.Sections("Section2").Controls("lblTotalMecProj").Caption = Conversion(dblTotal, MODE_ARGENT)
14 End If

14 Exit Sub

Oups:

14 wOups "frmChoixDateImpressionFacturation", "RemplirRapportMecanique", Err, Err.number, Err.Description
End Sub

Private Sub RemplirTempsReelsElec(ByVal sProjet As String)
 
 On Error GoTo Oups

 Dim rstPunch As ADODB.Recordset
 Dim sDateDebut As String
 Dim sDateFin As String
 Dim sTotal As String
 Dim sFilterNoProjet As String
 Dim Compile1 As String
 Dim Compile2 As String
 
 Compile1 = 0
 Compile2 = 0
 

 If Right$(sProjet, 2) = "99" Then
 sFilterNoProjet = "LEFT(NoProjet, 6) = '" & Left$(sProjet, 6) & "'"
 Else
 sFilterNoProjet = "NoProjet = '" & sProjet & "'"
 End If

  sDateDebut = "TIMESERIAL(Left(GrbPunch.HeureDébut,2),RIGHT(GrbPunch.HeureDébut,2),0)"

  sDateFin = "TIMESERIAL(Left(GrbPunch.HeureFin,2),RIGHT(GrbPunch.HeureFin,2),0)"

  sTotal = "(SUM(" & sDateFin & " - " & sDateDebut & ")* 24) As Total"

  Set rstPunch = New ADODB.Recordset

  If optChoix(I_OPT_PROJET_ENTIER).Value = True Then
  Call rstPunch.Open("SELECT Type, " & sTotal & " FROM GrbPunch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null GROUP BY Type", g_connData, adOpenForwardOnly, adLockReadOnly)
  Else
  Call rstPunch.Open("SELECT Type, " & sTotal & " FROM GrbPunch WHERE " & sFilterNoProjet & " AND HeureFin Is Not Null AND HeureDébut Is not Null AND [Date] BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' GROUP BY Type", g_connData, adOpenForwardOnly, adLockReadOnly)
10 End If

DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecDessinProj").Caption = "0"
DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFabricationProj").Caption = "0"
DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecAssemblageProj").Caption = "0"
DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgInterfaceProj").Caption = "0"
DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgAutomateProj").Caption = "0"
DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgRobotProj").Caption = "0"
DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecVisionProj").Caption = "0"
DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTestProj").Caption = "0"
DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecInstallationProj").Caption = "0"
DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecMiseServiceProj").Caption = "0"
DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFormationProj").Caption = "0"
1  DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecGestionProj").Caption = "0"
DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecShippingProj").Caption = "0"
 DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecRechercheProj").Caption = "0"

 Do While Not rstPunch.EOF
 If Not IsNull(rstPunch.Fields("Total")) Then
 Select Case rstPunch.Fields("Type")
 Case "Dessin": DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecDessinProj").Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Fabrication": DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFabricationProj").Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Assemblage": DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecAssemblageProj").Caption = Round(rstPunch.Fields("Total"), 2)
1  Case "ProgInterface": DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgInterfaceProj").Caption = Round(rstPunch.Fields("Total"), 2)
 Case "ProgAutomate": DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgAutomateProj").Caption = Round(rstPunch.Fields("Total"), 2)
 Case "ProgRobot": DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgRobotProj").Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Vision": DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecVisionProj").Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Test": DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTestProj").Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Installation": DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecInstallationProj").Caption = Round(rstPunch.Fields("Total"), 2)
 Case "MiseService": DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecMiseServiceProj").Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Formation": DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFormationProj").Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Gestion": DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecGestionProj").Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Shipping": DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecShippingProj").Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Prototypage-Dévelloppement expérimental": Compile1 = Round(rstPunch.Fields("Total"), 2)
 Case "": Compile2 = Round(rstPunch.Fields("Total"), 2)
 
 ''''''''''''''''''''''''''''''''''''''''''''
 ''''''''''''''''''''''''''''''''''''''''''''
 'ajout prototypage expérimental
 
 
 
 '''''''''''''''''''''' modif alex   fevrier 2012
 '''''''''''''''''''''''''''''''''''''''''''''''''''''
 
 
 End Select
 End If

 Call rstPunch.MoveNext
2  Loop

'''''''''''''''''''''''
' On addtionne develloppement avec aucun type ensemble
'''''''''''''''
DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecRechercheProj").Caption = CStr(CDbl(Compile1) + CDbl(Compile2))



Call rstPunch.Close
2  Set rstPunch = Nothing

Exit Sub

Oups:

2  wOups "frmProjSoumElecTemps", "RemplirTempsReelsElec", Err, Err.number, Err.Description
End Sub

Private Sub RemplirTempsReelsMec(ByVal sProjet As String)
 
 On Error GoTo Oups

 Dim rstPunch As ADODB.Recordset
 Dim sDateDebut As String
 Dim sDateFin As String
 Dim sTotal As String
 Dim sFilterNoProjet As String
 Dim test As String
 Dim Compile1 As String
 Dim Compile2 As String
 
 Compile1 = 0
 Compile2 = 0
 
 If Right$(sProjet, 2) = "99" Then
 sFilterNoProjet = "LEFT(NoProjet, 6) = '" & Left$(sProjet, 6) & "'"
 Else
 sFilterNoProjet = "NoProjet = '" & sProjet & "'"
 End If

  sDateDebut = "TIMESERIAL(Left(GrbPunch.HeureDébut,2),RIGHT(GrbPunch.HeureDébut,2),0)"

  sDateFin = "TIMESERIAL(Left(GrbPunch.HeureFin,2),RIGHT(GrbPunch.HeureFin,2),0)"

  sTotal = "(SUM(" & sDateFin & " - " & sDateDebut & ")* 24) As Total"

  Set rstPunch = New ADODB.Recordset

  If optChoix(I_OPT_PROJET_ENTIER).Value = True Then
  Call rstPunch.Open("SELECT Type, " & sTotal & " FROM GrbPunch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null GROUP BY Type", g_connData, adOpenForwardOnly, adLockReadOnly)
  Else
  Call rstPunch.Open("SELECT Type, " & sTotal & " FROM GrbPunch WHERE " & sFilterNoProjet & " AND HeureFin Is Not Null AND HeureDébut Is Not Null AND [Date] BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' GROUP BY Type", g_connData, adOpenForwardOnly, adLockReadOnly)
10 End If

DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecDessinProj").Caption = "0"
DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecCoupeProj").Caption = "0"
DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecMachinageProj").Caption = "0"
DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecSoudureProj").Caption = "0"
DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecAssemblageProj").Caption = "0"
DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecPeintureProj").Caption = "0"
DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTestProj").Caption = "0"
DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecInstallationProj").Caption = "0"
DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecFormationProj").Caption = "0"
DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecGestionProj").Caption = "0"
DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecShippingProj").Caption = "0"

1  Do While Not rstPunch.EOF
 If Not IsNull(rstPunch.Fields("Total")) Then
 Select Case rstPunch.Fields("Type")
 Case "Dessin": DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecDessinProj").Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Coupe": DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecCoupeProj").Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Machinage": DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecMachinageProj").Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Soudure": DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecSoudureProj").Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Assemblage": DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecAssemblageProj").Caption = Round(rstPunch.Fields("Total"), 2)
1  Case "Peinture": DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecPeintureProj").Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Test": DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTestProj").Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Installation": DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecInstallationProj").Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Formation": DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecFormationProj").Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Gestion": DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecGestionProj").Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Shipping": DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecShippingProj").Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Prototypage-Dévelloppement expérimental": Compile1 = Round(rstPunch.Fields("Total"), 2)
 Case "": Compile2 = Round(rstPunch.Fields("Total"), 2)
 

 End Select
 End If

 Call rstPunch.MoveNext
Loop
'''''''''''''''''''''''''''''''''
' s'il y a des enregistrement sans type , compile sans develloppement


DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecRechercheProj").Caption = CStr(CDbl(Compile1) + CDbl(Compile2))
''''''''''''''''''''''''''''''''''''


Call rstPunch.Close
Set rstPunch = Nothing

Exit Sub

Oups:

2  wOups "frmChoixDateImpressionFacturation", "RemplirTempsReelsElec", Err, Err.number, Err.Description
End Sub


Private Sub Form_Click()

 On Error GoTo Oups

 mvwDate.Visible = False

 Exit Sub

Oups:

 wOups "frmChoixDateImpressionFacturation", "Form_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()
 
 On Error GoTo Oups

 optChoix(I_OPT_PROJET_ENTIER).Value = True
 optChoixProjetEntier(0).Value = True

 Exit Sub

Oups:

 wOups "frmChoixDateImpressionFacturation", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub mvwDate_LostFocus()

 On Error GoTo Oups

 mvwDate.Visible = False

 Exit Sub

Oups:

 wOups "frmChoixDateImpressionFacturation", "mvwDate_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mvwDate_DateClick(ByVal DateClicked As Date)

 On Error GoTo Oups

 Select Case m_eDate
 Case DEBUT: mskDateDebut.Text = ConvertDate(DateClicked)
 Case Fin: mskDateFin.Text = ConvertDate(DateClicked)
 End Select
 
 'Enlever le calendrier
 mvwDate.Visible = False

 Exit Sub

Oups:

 wOups "frmChoixDateImpressionFacturation", "mvwDate_DateClick", Err, Err.number, Err.Description
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

 wOups "frmChoixDateImpressionFacturation", "mskDateDebut_GotFocus", Err, Err.number, Err.Description
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

 wOups "frmChoixDateImpressionFacturation", "mskDateFin_GotFocus", Err, Err.number, Err.Description
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

  wOups "frmChoixDateImpressionFacturation", "mskDateDebut_LostFocus", Err, Err.number, Err.Description
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

  wOups "frmChoixDateImpressionFacturation", "mskDateFin_LostFocus", Err, Err.number, Err.Description
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

  wOups "frmChoixDateImpressionFacturation", "cmdDateDebut_Click", Err, Err.number, Err.Description
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

  wOups "frmChoixDateImpressionFacturation", "cmdDateFin_Click", Err, Err.number, Err.Description
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

 wOups "frmChoixDateImpressionFacturation", "ValiderDate", Err, Err.number, Err.Description
End Function

Private Sub optChoix_Click(Index As Integer)

 On Error GoTo Oups

 If optChoix(I_OPT_PROJET_ENTIER).Value = True Then
 fra2Dates.Enabled = False
 Else
 fra2Dates.Enabled = True
 End If

 Exit Sub

Oups:

 wOups "frmChoixDateImpressionFacturation", "optChoix_Click", Err, Err.number, Err.Description
End Sub

