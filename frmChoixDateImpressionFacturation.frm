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
   MinButton       =   0   'False
   Picture         =   "frmChoixDateImpressionFacturation.frx":0000
   ScaleHeight     =   3915
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.MonthView mvwDate 
      Height          =   2370
      Left            =   3720
      TabIndex        =   0
      Top             =   840
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
Private Const I_OPT_2_DATES       As Integer = 1

Private Const I_OPT_LISTE_PUNCH   As Integer = 0
Private Const I_OPT_COUTANT       As Integer = 1

Private m_eDate        As enumDate
Private m_sNoProjSoum  As String
Private m_bProjet      As Boolean
Private m_sClient      As String
Private m_sDescription As String

Public Sub Afficher(ByVal sNoProjSoum As String, ByVal bProjet As Boolean, ByVal sClient As String, ByVal sDescription As String)

5       On Error GoTo AfficherErreur

10      m_sNoProjSoum = sNoProjSoum

15      m_bProjet = bProjet

20      If bProjet = True Then
25        optChoix(I_OPT_COUTANT).Enabled = True
30      Else
35        optChoix(I_OPT_COUTANT).Enabled = False
40      End If

45      m_sClient = sClient

50      m_sDescription = sDescription

55      Call Me.Show(vbModal)

60      Exit Sub

AfficherErreur:

65      woups "frmChoixDateImpressionFacturation", "Afficher", Err, Erl
End Sub

Private Sub cmdAnnuler_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmChoixDateImpressionFacturation", "cmdAnnuler_Click", Err, Erl
End Sub

Private Sub cmdImprimer_Click()

5       On Error GoTo AfficherErreur

10      If optChoixProjetEntier(I_OPT_LISTE_PUNCH).Value = True Then
15        Call ImprimerListePunch
20      Else
25        Call ImprimerPrixCoutant
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmChoixDateImpressionFacturation", "cmdImprimer_Click", Err, Erl
End Sub

Private Sub ImprimerListePunch()
  
5       On Error GoTo AfficherErreur

10      Dim rstPunch    As ADODB.Recordset
15      Dim rstSomme    As ADODB.Recordset
20      Dim iCompteur   As Integer
25      Dim bNonComplet As Boolean
        
30      If optChoix(I_OPT_2_DATES).Value = True Then
35        If mskDateDebut.Text <> "" Then
40          If mskDateFin.Text <> "" Then
45            If ValiderDate(mskDateDebut.Text) = True Then
50              If ValiderDate(mskDateFin.Text) = True Then
55                If mskDateDebut.Text > mskDateFin.Text Then
60                  Call MsgBox("La date de début doit être plus petite que la date de fin!", vbOKOnly, "Erreur")

65                  Exit Sub
70                End If
75              Else
80                Call MsgBox("Date de fin non valide!", vbOKOnly, "Erreur")

85                Exit Sub
90              End If
95            Else
100             Call MsgBox("Date de début non valide!", vbOKOnly, "Erreur")

105             Exit Sub
110           End If
115         Else
120           Call MsgBox("La date de fin est obligatoire!", vbOKOnly, "Erreur")

125           Exit Sub
130         End If
135       Else
140         Call MsgBox("La date de début est obligatoire!", vbOKOnly, "Erreur")

145         Exit Sub
150       End If
175     End If

        'Si il y a des projets ou des soumissions
180     If m_sNoProjSoum <> "" Then
185       If optChoix(I_OPT_PROJET_ENTIER).Value = True Then
190         For iCompteur = 1 To frmFacturation.lvwProjets.ListItems.count
195           If frmFacturation.lvwProjets.ListItems(iCompteur).SubItems(3) = "" Then
200             bNonComplet = True

205             Exit For
210           End If
215         Next

220         If bNonComplet = True Then
225           If MsgBox("Les punchs ne sont pas complets!" & vbNewLine & "Voulez-vous imprimer seulement les punchs complets?", vbYesNo) = vbNo Then
230             Exit Sub
235           End If
240         End If

245         Set rstPunch = New ADODB.Recordset

250         rstPunch.CursorLocation = adUseServer

            '*************************************************************************
            'ajout du champ type dans la requête PAR GAÉTAN GINGRAS LE 07 FÉVRIER 2010
            If MsgBox("Désirez-vous afficher les commentaires avec le type des travaux?", vbYesNo, "Choix d'affichage") = vbYes Then
255             Call rstPunch.Open("SELECT (GRB_Punch.Type & ' - ' & GRB_Punch.Commentaire) AS Comment, GRB_Punch.Date, GRB_Punch.HeureDébut, GRB_Punch.HeureFin, GRB_Punch.Facturé, GRB_Punch.NoFacture, GRB_Employés.Initiale, Round((TimeSerial(Left(GRB_Punch.HeureFin,2), RIGHT(GRB_Punch.HeureFin,2),0) - TimeSerial(Left(GRB_Punch.HeureDébut,2), RIGHT(GRB_Punch.HeureDébut,2),0)) * 24, 2) As Total FROM GRB_Punch INNER JOIN GRB_Employés ON GRB_Punch.NoEmploye = GRB_Employés.noEmploye WHERE GRB_Punch.NoProjet = '" & m_sNoProjSoum & "' AND HeureFin IS NOT NULL ORDER BY [Date]", g_connData, adOpenDynamic, adLockOptimistic)
            Else
                Call rstPunch.Open("SELECT GRB_Punch.Type AS Comment, GRB_Punch.Date, GRB_Punch.HeureDébut, GRB_Punch.HeureFin, GRB_Punch.Commentaire, GRB_Punch.Facturé, GRB_Punch.NoFacture, GRB_Employés.Initiale, Round((TimeSerial(Left(GRB_Punch.HeureFin,2), RIGHT(GRB_Punch.HeureFin,2),0) - TimeSerial(Left(GRB_Punch.HeureDébut,2), RIGHT(GRB_Punch.HeureDébut,2),0)) * 24, 2) As Total FROM GRB_Punch INNER JOIN GRB_Employés ON GRB_Punch.NoEmploye = GRB_Employés.noEmploye WHERE GRB_Punch.NoProjet = '" & m_sNoProjSoum & "' AND HeureFin IS NOT NULL ORDER BY [Date]", g_connData, adOpenDynamic, adLockOptimistic)
            End If
            '*************************************************************************
            
260       Else
265         For iCompteur = 1 To frmFacturation.lvwProjets.ListItems.count
270           If frmFacturation.lvwProjets.ListItems(iCompteur).SubItems(3) = "" Then
275             If frmFacturation.lvwProjets.ListItems(iCompteur).SubItems(1) >= mskDateDebut.Text And frmFacturation.lvwProjets.ListItems(iCompteur).SubItems(1) >= mskDateFin.Text Then
280               bNonComplet = True

285               Exit For
290             End If
295           End If
300         Next

305         If bNonComplet = True Then
310           If MsgBox("Les punchs ne sont pas complets!" & vbNewLine & "Voulez-vous imprimer seulement les punchs complets?", vbYesNo) = vbNo Then
315             Exit Sub
320           End If
325         End If

330         Set rstPunch = New ADODB.Recordset

335         rstPunch.CursorLocation = adUseServer

            '**************************************************************************
            'ajout du champ type dans la requête PAR GAÉTAN GINGRAS LE 07 FÉVRIER 2010
            If MsgBox("Désirez-vous afficher les commentaires avec le type des travaux?", vbYesNo, "Choix d'affichage") = vbYes Then
340             Call rstPunch.Open("SELECT (GRB_Punch.Type & ' - ' & GRB_Punch.Commentaire) AS Comment, GRB_Punch.Date, GRB_Punch.HeureDébut, GRB_Punch.HeureFin, GRB_Punch.Commentaire, GRB_Punch.Facturé, GRB_Punch.NoFacture, GRB_Employés.Initiale, Round((TimeSerial(Left(GRB_Punch.HeureFin,2), RIGHT(GRB_Punch.HeureFin,2),0) - TimeSerial(Left(GRB_Punch.HeureDébut,2), RIGHT(GRB_Punch.HeureDébut,2),0)) * 24, 2) As Total FROM GRB_Punch INNER JOIN GRB_Employés ON GRB_Punch.NoEmploye = GRB_Employés.noEmploye WHERE GRB_Punch.NoProjet = '" & m_sNoProjSoum & "' AND HeureFin IS NOT NULL AND [GRB_Punch.Date] BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' ORDER BY [Date]", g_connData, adOpenDynamic, adLockOptimistic)
            Else
                Call rstPunch.Open("SELECT GRB_Punch.Type AS Comment, GRB_Punch.Date, GRB_Punch.HeureDébut, GRB_Punch.HeureFin, GRB_Punch.Commentaire, GRB_Punch.Facturé, GRB_Punch.NoFacture, GRB_Employés.Initiale, Round((TimeSerial(Left(GRB_Punch.HeureFin,2), RIGHT(GRB_Punch.HeureFin,2),0) - TimeSerial(Left(GRB_Punch.HeureDébut,2), RIGHT(GRB_Punch.HeureDébut,2),0)) * 24, 2) As Total FROM GRB_Punch INNER JOIN GRB_Employés ON GRB_Punch.NoEmploye = GRB_Employés.noEmploye WHERE GRB_Punch.NoProjet = '" & m_sNoProjSoum & "' AND HeureFin IS NOT NULL AND [GRB_Punch.Date] BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' ORDER BY [Date]", g_connData, adOpenDynamic, adLockOptimistic)
            End If
            '**************************************************************************
            
345       End If

350       Set DR_Facturation.DataSource = rstPunch

355       DR_Facturation.Orientation = rptOrientLandscape

360       If m_bProjet = True Then
365         DR_Facturation.Sections("Section4").Controls("lblTitreNumero").Caption = "Numéro de projet :"
370       Else
375         DR_Facturation.Sections("Section4").Controls("lblTitreNumero").Caption = "Numéro de soumission :"
380       End If

385       DR_Facturation.Sections("Section4").Controls("lblNumero").Caption = m_sNoProjSoum
390       DR_Facturation.Sections("Section4").Controls("lblClient").Caption = m_sClient

        'affiche la date
        '**************************************************
        'ajout par Gaétan Gingras le 20 mai 2009
394     If MsgBox("Désirez-vous afficher la date en bas de page ?", vbYesNo + vbInformation, "Affichage de la date") = vbYes Then
395         DR_Facturation.Sections("Section3").Controls("lblDate").Caption = ConvertDate(Date)
396     Else
397         DR_Facturation.Sections("Section3").Controls("lblDate").Caption = " "
398     End If

        '**************************************************
        
        'affichage des colonnes facturé et no. de facture
        '**************************************************
        'ajout de Gaétan Gingras le 20 mai 2009
399     If MsgBox("Désirez-vous afficher les colonnes 'facturé' et 'no. facture'?", vbYesNo + vbInformation, "Affichage de la date") = vbYes Then
400         DR_Facturation.Sections("Section1").Controls("text1").Visible = True
401         DR_Facturation.Sections("Section1").Controls("text4").Visible = True
402         DR_Facturation.Sections("Section2").Controls("label4").Visible = True
403         DR_Facturation.Sections("Section2").Controls("label14").Visible = True
404     Else
405         DR_Facturation.Sections("Section1").Controls("text1").Visible = False
406         DR_Facturation.Sections("Section1").Controls("text4").Visible = False
407         DR_Facturation.Sections("Section2").Controls("label4").Visible = False
408         DR_Facturation.Sections("Section2").Controls("label14").Visible = False
409     End If
        '**************************************************
        
410       Set rstSomme = New ADODB.Recordset

413       If optChoix(I_OPT_PROJET_ENTIER).Value = True Then
414         Call rstSomme.Open("SELECT SUM(TimeSerial(Left(HeureFin,2),RIGHT(HeureFin,2),0) - TimeSerial(Left(HeureDébut,2), RIGHT(HeureDébut,2),0)) As Total FROM GRB_Punch WHERE NoProjet = '" & m_sNoProjSoum & "' AND Facturé = True AND HeureFin IS NOT NULL", g_connData, adOpenDynamic, adLockOptimistic)
415       Else
420         Call rstSomme.Open("SELECT SUM(TimeSerial(Left(HeureFin,2),RIGHT(HeureFin,2),0) - TimeSerial(Left(HeureDébut,2), RIGHT(HeureDébut,2),0)) As Total FROM GRB_Punch WHERE NoProjet = '" & m_sNoProjSoum & "' AND Facturé = True AND HeureFin IS NOT NULL AND [Date] BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
425       End If

430       If Not IsNull(rstSomme.Fields("Total")) Then
435         DR_Facturation.Sections("Section5").Controls("lblHeuresFacturees").Caption = Round(rstSomme.Fields("Total") * 24, 4)
440       Else
445         DR_Facturation.Sections("Section5").Controls("lblHeuresFacturees").Caption = "0"
450       End If

455       Call rstSomme.Close

460       rstSomme.CursorLocation = adUseServer

465       If optChoix(I_OPT_PROJET_ENTIER).Value = True Then
470         Call rstSomme.Open("SELECT SUM(TimeSerial(Left(HeureFin,2), RIGHT(HeureFin,2),0) - TimeSerial(Left(HeureDébut,2),RIGHT(HeureDébut,2),0)) As Total FROM GRB_Punch WHERE NoProjet = '" & m_sNoProjSoum & "' AND Facturé = False AND HeureFin IS NOT NULL", g_connData, adOpenDynamic, adLockOptimistic)
475       Else
480         Call rstSomme.Open("SELECT SUM(TimeSerial(Left(HeureFin,2), RIGHT(HeureFin,2),0) - TimeSerial(Left(HeureDébut,2),RIGHT(HeureDébut,2),0)) As Total FROM GRB_Punch WHERE NoProjet = '" & m_sNoProjSoum & "' AND Facturé = False AND HeureFin IS NOT NULL AND [Date] BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
485       End If

490       If Not IsNull(rstSomme.Fields("Total")) Then
495         DR_Facturation.Sections("Section5").Controls("lblHeuresNonFacturees").Caption = Round(rstSomme.Fields("Total") * 24, 4)
500       Else
505         DR_Facturation.Sections("Section5").Controls("lblHeuresNonFacturees").Caption = "0"
510       End If

515       Call rstSomme.Close
520       Set rstSomme = Nothing

525       DR_Facturation.Sections("Section5").Controls("lblGrandTotal").Caption = CDbl(DR_Facturation.Sections("Section5").Controls("lblHeuresFacturees").Caption) + CDbl(DR_Facturation.Sections("Section5").Controls("lblHeuresNonFacturees").Caption)

530       If optChoix(I_OPT_PROJET_ENTIER).Value = True Then
535         DR_Facturation.Sections("Section4").Controls("lblDateDebut").Caption = "N/A"
540         DR_Facturation.Sections("Section4").Controls("lblDateFin").Caption = "N/A"
545       Else
550         DR_Facturation.Sections("Section4").Controls("lblDateDebut").Caption = mskDateDebut.Text
555         DR_Facturation.Sections("Section4").Controls("lblDateFin").Caption = mskDateFin.Text
560       End If

565       Call DR_Facturation.Show(vbModal)

570       Call rstPunch.Close
575       Set rstPunch = Nothing
580     End If

585     Call Unload(Me)

590     Exit Sub

AfficherErreur:

595     woups "frmChoixDateImpressionFacturation", "ImprimerListePunch", Err, Erl
End Sub

Private Sub ImprimerPrixCoutant()

5       On Error GoTo AfficherErreur

10      Dim dblTotal As Double
15      Dim sProjet  As String
20      Dim rstDS    As ADODB.Recordset

25      If optChoix(I_OPT_2_DATES).Value = True Then
30        If mskDateDebut.Text <> "" Then
35          If mskDateFin.Text <> "" Then
40            If ValiderDate(mskDateDebut.Text) = True Then
45              If ValiderDate(mskDateFin.Text) = True Then
50                If mskDateDebut.Text > mskDateFin.Text Then
55                  Call MsgBox("La date de début doit être plus petite que la date de fin!", vbOKOnly, "Erreur")

60                  Exit Sub
65                End If
70              Else
75                Call MsgBox("Date de fin non valide!", vbOKOnly, "Erreur")

80                Exit Sub
85              End If
90            Else
95              Call MsgBox("Date de début non valide!", vbOKOnly, "Erreur")

100             Exit Sub
105           End If
110         Else
115           Call MsgBox("La date de fin est obligatoire!", vbOKOnly, "Erreur")

120           Exit Sub
125         End If
130       Else
135         Call MsgBox("La date de début est obligatoire!", vbOKOnly, "Erreur")

140         Exit Sub
145       End If
150     End If

155     If Len(m_sNoProjSoum) = 9 Then
160       sProjet = Right$(m_sNoProjSoum, 8)
165     Else
170       Call MsgBox("Numéro de projet non valide!", vbOKOnly, "Erreur")

175       Exit Sub
180     End If

        'Ce recordset ne sert absolument à rien,
        'il est seulement utiliser parce que le DR a besoin d'un DataSource pour ouvrir
185     Set rstDS = New ADODB.Recordset

190     Call rstDS.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & m_sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)

195     Set DR_ApercuProjet.DataSource = rstDS

200     DR_ApercuProjet.Sections("Section2").Controls("lblNumero").Caption = sProjet
205     DR_ApercuProjet.Sections("Section2").Controls("lblClient").Caption = m_sClient
210     DR_ApercuProjet.Sections("Section2").Controls("lblDescription").Caption = m_sDescription

215     If optChoix(I_OPT_PROJET_ENTIER).Value = True Then
220       DR_ApercuProjet.Sections("Section2").Controls("lblDate").Caption = ConvertDate(Date)
225     Else
230       DR_ApercuProjet.Sections("Section2").Controls("lblDate").Caption = "Du " & mskDateDebut.Text & " au " & mskDateFin.Text
235     End If

240     Call RemplirRapportElectrique(sProjet)
245     Call RemplirRapportMecanique(sProjet)

250     If IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitElecSoum").Caption) And IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitMecSoum").Caption) Then
255       dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitElecSoum").Caption) + _
                     CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitMecSoum").Caption)
260     Else
265       If IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitElecSoum").Caption) Then
270         dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitElecSoum").Caption)
275       Else
280         If IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitMecSoum").Caption) Then
285           dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitMecSoum").Caption)
290         Else
295           dblTotal = 0
300         End If
305       End If
310     End If

315     DR_ApercuProjet.Sections("Section2").Controls("lblTotalForfaitSoum").Caption = Conversion(dblTotal, MODE_ARGENT)

320     If IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitElecProj").Caption) And IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitMecProj").Caption) Then
325       dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitElecProj").Caption) + _
                     CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitMecProj").Caption)
330     Else
335       If IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitElecProj").Caption) Then
340         dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitElecProj").Caption)
345       Else
350         If IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitMecProj").Caption) Then
355           dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblForfaitMecProj").Caption)
360         Else
365           dblTotal = 0
370         End If
375       End If
380     End If

385     DR_ApercuProjet.Sections("Section2").Controls("lblTotalForfaitProj").Caption = Conversion(dblTotal, MODE_ARGENT)

390     If IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTotalSoum").Caption) And IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTotalSoum").Caption) Then
395       dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTotalSoum").Caption) + _
                     CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTotalSoum").Caption)

400       DR_ApercuProjet.Sections("Section2").Controls("lblTotalHeuresSoum").Caption = dblTotal
405     Else
410       If Not IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTotalSoum").Caption) And Not IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTotalSoum").Caption) Then
415         DR_ApercuProjet.Sections("Section2").Controls("lblTotalHeuresSoum").Caption = "---"
420       Else
425         If Not IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTotalSoum").Caption) Then
430           DR_ApercuProjet.Sections("Section2").Controls("lblTotalHeuresSoum").Caption = DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTotalSoum").Caption
435         Else
440           DR_ApercuProjet.Sections("Section2").Controls("lblTotalHeuresSoum").Caption = DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTotalSoum").Caption
445         End If
450       End If
455     End If

460     If IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblTotalElecSoum").Caption) And IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblTotalMecSoum").Caption) Then
465       dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblTotalElecSoum").Caption) + _
                     CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblTotalMecSoum").Caption)

470       DR_ApercuProjet.Sections("Section2").Controls("lblGrandTotalSoum").Caption = Conversion(dblTotal, MODE_ARGENT)
475     Else
480       If Not IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblTotalElecSoum").Caption) And Not IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblTotalMecSoum").Caption) Then
485         DR_ApercuProjet.Sections("Section2").Controls("lblGrandTotalSoum").Caption = "---"
490       Else
495         If Not IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblTotalElecSoum").Caption) Then
500           DR_ApercuProjet.Sections("Section2").Controls("lblGrandTotalSoum").Caption = DR_ApercuProjet.Sections("Section2").Controls("lblTotalMecSoum").Caption
505         Else
510           DR_ApercuProjet.Sections("Section2").Controls("lblGrandTotalSoum").Caption = DR_ApercuProjet.Sections("Section2").Controls("lblTotalElecSoum").Caption
515         End If
520       End If
525     End If

530     If IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTotalProj").Caption) And IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTotalProj").Caption) Then
535       dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTotalProj").Caption) + _
                     CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTotalProj").Caption)

540       DR_ApercuProjet.Sections("Section2").Controls("lblTotalHeuresProj").Caption = dblTotal
545     Else
550       If Not IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTotalProj").Caption) And Not IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTotalProj").Caption) Then
555         DR_ApercuProjet.Sections("Section2").Controls("lblTotalHeuresProj").Caption = "---"
560       Else
565         If Not IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTotalProj").Caption) Then
570           DR_ApercuProjet.Sections("Section2").Controls("lblTotalHeuresProj").Caption = DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTotalProj").Caption
575         Else
580           DR_ApercuProjet.Sections("Section2").Controls("lblTotalHeuresProj").Caption = DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTotalProj").Caption
585         End If
590       End If
595     End If

600     If IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblTotalElecProj").Caption) And IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblTotalMecProj").Caption) Then
605       dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblTotalElecProj").Caption) + _
                     CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblTotalMecProj").Caption)

610       DR_ApercuProjet.Sections("Section2").Controls("lblGrandTotalProj").Caption = Conversion(dblTotal, MODE_ARGENT)
615     Else
620       If Not IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblTotalElecProj").Caption) And Not IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblTotalMecProj").Caption) Then
625         DR_ApercuProjet.Sections("Section2").Controls("lblGrandTotalProj").Caption = "---"
630       Else
635         If Not IsNumeric(DR_ApercuProjet.Sections("Section2").Controls("lblTotalElecProj").Caption) Then
640           DR_ApercuProjet.Sections("Section2").Controls("lblGrandTotalProj").Caption = DR_ApercuProjet.Sections("Section2").Controls("lblTotalMecProj").Caption
645         Else
650           DR_ApercuProjet.Sections("Section2").Controls("lblGrandTotalProj").Caption = DR_ApercuProjet.Sections("Section2").Controls("lblTotalElecProj").Caption
655         End If
660       End If
665     End If

670     If DR_ApercuProjet.Sections("Section2").Controls("lblTotalForfaitSoum").Caption <> "---" And _
           DR_ApercuProjet.Sections("Section2").Controls("lblGrandTotalSoum").Caption <> "---" Then
675       dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblTotalForfaitSoum").Caption) - CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblGrandTotalSoum").Caption)
680     Else
685       dblTotal = 0
690     End If

695     DR_ApercuProjet.Sections("Section2").Controls("lblProfitSoum").Caption = Conversion(dblTotal, MODE_ARGENT)

700     If DR_ApercuProjet.Sections("Section2").Controls("lblTotalForfaitProj").Caption <> "---" And _
           DR_ApercuProjet.Sections("Section2").Controls("lblGrandTotalProj").Caption <> "---" Then
705       dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblTotalForfaitProj").Caption) - CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblGrandTotalProj").Caption)
710     Else
715       dblTotal = 0
720     End If

725     DR_ApercuProjet.Sections("Section2").Controls("lblProfitProj").Caption = Conversion(dblTotal, MODE_ARGENT)

730     Call DR_ApercuProjet.Show(vbModal)

735     Call rstDS.Close
740     Set rstDS = Nothing

745     Call Unload(Me)

750     Exit Sub

AfficherErreur:

755     woups "frmChoixDateImpressionFacturation", "ImprimerPrixCoutant", Err, Erl
End Sub

Private Sub RemplirRapportElectrique(ByVal sProjet As String)
  
5       On Error GoTo AfficherErreur

10      Dim rstProjetElec      As ADODB.Recordset
15      Dim rstSoumElec        As ADODB.Recordset
20      Dim rstProjetPieces    As ADODB.Recordset
25      Dim dblTotal           As Double
30      Dim bSoumission        As Boolean
35      Dim iNbrePersonne      As Integer
40      Dim dblHebergement     As Double
45      Dim dblRepas           As Double
50      Dim dblTransport       As Double
55      Dim dblUniteMobile     As Double
60      Dim dblPrixEmballage   As Double
65      Dim dblTotalResteTemps As Double
70      Dim dblTotalManuel     As Double
75      Dim dblTotalPieces     As Double
                
80      Set rstProjetElec = New ADODB.Recordset

85      Call rstProjetElec.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = 'E" & sProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

90      DR_ApercuProjet.Sections("Section2").Controls("lblProjetElec").Caption = "E" & sProjet

95      If Not rstProjetElec.EOF Then
100       bSoumission = False

105       If Not IsNull(rstProjetElec.Fields("IDSoumission")) Then
110         Set rstSoumElec = New ADODB.Recordset

115         Call rstSoumElec.Open("SELECT * FROM GRB_SoumissionElec WHERE IDSoumission = '" & rstProjetElec.Fields("IDSoumission") & "'", g_connData, adOpenDynamic, adLockOptimistic)

120         If Not rstSoumElec.EOF Then
125           bSoumission = True
130         Else
135           Call rstSoumElec.Close
140           Set rstSoumElec = Nothing
145         End If
150       End If

155       If bSoumission = True Then
160         If Not IsNull(rstSoumElec.Fields("MontantForfait")) Then
165           DR_ApercuProjet.Sections("Section2").Controls("lblForfaitElecSoum").Caption = Conversion(rstSoumElec.Fields("MontantForfait"), MODE_ARGENT)
170         Else
175           DR_ApercuProjet.Sections("Section2").Controls("lblForfaitElecSoum").Caption = Conversion("0", MODE_ARGENT)
180         End If

185         If Not IsNull(rstSoumElec.Fields("TempsDessin")) Then
190           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecDessinSoum").Caption = rstSoumElec.Fields("TempsDessin")
195           DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecDessinSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsDessin")) * CDbl(rstSoumElec.Fields("TauxDessin")), MODE_ARGENT)
200         Else
205           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecDessinSoum").Caption = "0"
210           DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecDessinSoum").Caption = Conversion("0", MODE_ARGENT)
215         End If

220         If Not IsNull(rstSoumElec.Fields("TempsFabrication")) Then
225           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFabricationSoum").Caption = rstSoumElec.Fields("TempsFabrication")
230           DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFabricationSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsFabrication")) * CDbl(rstSoumElec.Fields("TauxFabrication")), MODE_ARGENT)
235         Else
240           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFabricationSoum").Caption = "0"
245           DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFabricationSoum").Caption = Conversion("0", MODE_ARGENT)
250         End If

255         If Not IsNull(rstSoumElec.Fields("TempsAssemblage")) Then
260           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecAssemblageSoum").Caption = rstSoumElec.Fields("TempsAssemblage")
265           DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecAssemblageSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsAssemblage")) * CDbl(rstSoumElec.Fields("TauxAssemblage")), MODE_ARGENT)
270         Else
275           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecAssemblageSoum").Caption = "0"
280           DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecAssemblageSoum").Caption = Conversion("0", MODE_ARGENT)
285         End If

290         If Not IsNull(rstSoumElec.Fields("TempsProgInterface")) Then
295           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgInterfaceSoum").Caption = rstSoumElec.Fields("TempsProgInterface")
300           DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgInterfaceSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsProgInterface")) * CDbl(rstSoumElec.Fields("TauxProgInterface")), MODE_ARGENT)
305         Else
310           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgInterfaceSoum").Caption = "0"
315           DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgInterfaceSoum").Caption = Conversion("0", MODE_ARGENT)
320         End If

325         If Not IsNull(rstSoumElec.Fields("TempsProgAutomate")) Then
330           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgAutomateSoum").Caption = rstSoumElec.Fields("TempsProgAutomate")
335           DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgAutomateSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsProgAutomate")) * CDbl(rstSoumElec.Fields("TauxProgAutomate")), MODE_ARGENT)
340         Else
345           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgAutomateSoum").Caption = "0"
350           DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgAutomateSoum").Caption = Conversion("0", MODE_ARGENT)
355         End If

360         If Not IsNull(rstSoumElec.Fields("TempsProgRobot")) Then
365           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgRobotSoum").Caption = rstSoumElec.Fields("TempsProgRobot")
370           DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgRobotSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsProgRobot")) * CDbl(rstSoumElec.Fields("TauxProgRobot")), MODE_ARGENT)
375         Else
380           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgRobotSoum").Caption = "0"
385           DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgRobotSoum").Caption = Conversion("0", MODE_ARGENT)
390         End If

395         If Not IsNull(rstSoumElec.Fields("TempsVision")) Then
400           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecVisionSoum").Caption = rstSoumElec.Fields("TempsVision")
405           DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecVisionSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsVision")) * CDbl(rstSoumElec.Fields("TauxVision")), MODE_ARGENT)
410         Else
415           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecVisionSoum").Caption = "0"
420           DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecVisionSoum").Caption = Conversion("0", MODE_ARGENT)
425         End If

430         If Not IsNull(rstSoumElec.Fields("TempsTest")) Then
435           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTestSoum").Caption = rstSoumElec.Fields("TempsTest")
440           DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecTestSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsTest")) * CDbl(rstSoumElec.Fields("TauxTest")), MODE_ARGENT)
445         Else
450           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTestSoum").Caption = rstSoumElec.Fields("TempsTest")
455           DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecTestSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsTest")) * CDbl(rstSoumElec.Fields("TauxTest")), MODE_ARGENT)
460         End If

465         If Not IsNull(rstSoumElec.Fields("TempsInstallation")) Then
470           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecInstallationSoum").Caption = rstSoumElec.Fields("TempsInstallation")
475           DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecInstallationSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsInstallation")) * CDbl(rstSoumElec.Fields("TauxInstallation")), MODE_ARGENT)
480         Else
485           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecInstallationSoum").Caption = "0"
490           DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecInstallationSoum").Caption = Conversion("0", MODE_ARGENT)
495         End If

500         If Not IsNull(rstSoumElec.Fields("TempsMiseService")) Then
505           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecMiseServiceSoum").Caption = rstSoumElec.Fields("TempsMiseService")
510           DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecMiseServiceSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsMiseService")) * CDbl(rstSoumElec.Fields("TauxMiseService")), MODE_ARGENT)
515         Else
520           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecMiseServiceSoum").Caption = "0"
525           DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecMiseServiceSoum").Caption = Conversion("0", MODE_ARGENT)
530         End If

535         If Not IsNull(rstSoumElec.Fields("TempsFormation")) Then
540           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFormationSoum").Caption = rstSoumElec.Fields("TempsFormation")
545           DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFormationSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsFormation")) * CDbl(rstSoumElec.Fields("TauxFormation")), MODE_ARGENT)
550         Else
555           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFormationSoum").Caption = "0"
560           DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFormationSoum").Caption = Conversion("0", MODE_ARGENT)
565         End If

570         If Not IsNull(rstSoumElec.Fields("TempsGestion")) Then
575           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecGestionSoum").Caption = rstSoumElec.Fields("TempsGestion")
580           DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecGestionSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsGestion")) * CDbl(rstSoumElec.Fields("TauxGestion")), MODE_ARGENT)
585         Else
590           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecGestionSoum").Caption = "0"
595           DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecGestionSoum").Caption = Conversion("0", MODE_ARGENT)
600         End If

605         If Not IsNull(rstSoumElec.Fields("TempsShipping")) Then
610           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecShippingSoum").Caption = rstSoumElec.Fields("TempsShipping")
615           DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecShippingSoum").Caption = Conversion(CDbl(rstSoumElec.Fields("TempsShipping")) * CDbl(rstSoumElec.Fields("TauxShipping")), MODE_ARGENT)
620         Else
625           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecShippingSoum").Caption = "0"
630           DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecShippingSoum").Caption = Conversion("0", MODE_ARGENT)
635         End If

640         DR_ApercuProjet.Sections("Section2").Controls("lblPiecesElecSoum").Caption = Conversion(rstSoumElec.Fields("total_piece"), MODE_ARGENT)
645         DR_ApercuProjet.Sections("Section2").Controls("lblImprevuElecSoum").Caption = Conversion(rstSoumElec.Fields("total_imprevue"), MODE_ARGENT)

650         If Not IsNull(rstSoumElec.Fields("NbrePersonne")) Then
655           iNbrePersonne = rstSoumElec.Fields("NbrePersonne")
660         Else
665           iNbrePersonne = 0
670         End If
           
675         Do While iNbrePersonne > 0
680           If iNbrePersonne >= 2 Then
685             dblHebergement = dblHebergement + rstSoumElec.Fields("TempsHebergement") * rstSoumElec.Fields("TauxHebergement2")
              
690             iNbrePersonne = iNbrePersonne - 2
695           Else
700             dblHebergement = dblHebergement + rstSoumElec.Fields("TempsHebergement") * rstSoumElec.Fields("TauxHebergement1")
             
705             iNbrePersonne = iNbrePersonne - 1
710           End If
715         Loop
            
720         If Not IsNull(rstSoumElec.Fields("TempsRepas")) Then
725           dblRepas = CDbl(rstSoumElec.Fields("TempsRepas")) * CDbl(rstSoumElec.Fields("TauxRepas")) * CDbl(rstSoumElec.Fields("NbrePersonne"))
730         Else
735           dblRepas = 0
740         End If

745         If Not IsNull(rstSoumElec.Fields("TempsTransport")) Then
750           dblTransport = CDbl(rstSoumElec.Fields("TempsTransport")) * CDbl(rstSoumElec.Fields("TauxTransport"))
755         Else
760           dblTransport = 0
765         End If

770         If Not IsNull(rstSoumElec.Fields("TempsUniteMobile")) Then
775           dblUniteMobile = CDbl(rstSoumElec.Fields("TempsUniteMobile")) * CDbl(rstSoumElec.Fields("TauxUniteMobile"))
780         Else
785           dblUniteMobile = 0
790         End If

795         If Not IsNull(rstSoumElec.Fields("PrixEmballage")) Then
800           dblPrixEmballage = CDbl(rstSoumElec.Fields("PrixEmballage"))
805         Else
810           dblPrixEmballage = 0
815         End If
      
820         dblTotalResteTemps = dblHebergement + dblRepas + dblTransport + dblUniteMobile + dblPrixEmballage
                                                              
825         If IsNumeric(rstSoumElec.Fields("total_manuel")) Then
830           dblTotalManuel = CDbl(rstSoumElec.Fields("total_manuel"))
835         Else
840           dblTotalManuel = 0
845         End If

850         DR_ApercuProjet.Sections("Section2").Controls("lblAutresElecSoum").Caption = Conversion(dblTotalResteTemps + dblTotalManuel, MODE_ARGENT)

855         Call rstSoumElec.Close
860         Set rstSoumElec = Nothing
865       Else
870         DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecDessinSoum").Caption = "---"
875         DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecDessinSoum").Caption = "---"

880         DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFabricationSoum").Caption = "---"
885         DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFabricationSoum").Caption = "---"

890         DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecAssemblageSoum").Caption = "---"
895         DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecAssemblageSoum").Caption = "---"

900         DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgInterfaceSoum").Caption = "---"
905         DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgInterfaceSoum").Caption = "---"

910         DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgAutomateSoum").Caption = "---"
915         DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgAutomateSoum").Caption = "---"

920         DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgRobotSoum").Caption = "---"
925         DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgRobotSoum").Caption = "---"

930         DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecVisionSoum").Caption = "---"
935         DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecVisionSoum").Caption = "---"

940         DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTestSoum").Caption = "---"
945         DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecTestSoum").Caption = "---"

950         DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecInstallationSoum").Caption = "---"
955         DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecInstallationSoum").Caption = "---"

960         DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecMiseServiceSoum").Caption = "---"
965         DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecMiseServiceSoum").Caption = "---"

970         DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFormationSoum").Caption = "---"
975         DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFormationSoum").Caption = "---"

980         DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecGestionSoum").Caption = "---"
985         DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecGestionSoum").Caption = "---"

990         DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecGestionSoum").Caption = "---"
995         DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecGestionSoum").Caption = "---"

1000        DR_ApercuProjet.Sections("Section2").Controls("lblPiecesElecSoum").Caption = "---"
1005        DR_ApercuProjet.Sections("Section2").Controls("lblImprevuElecSoum").Caption = "---"
1010        DR_ApercuProjet.Sections("Section2").Controls("lblAutresElecSoum").Caption = "---"
1015      End If

1020      Call RemplirTempsReelsElec("E" & sProjet)

1025      If Not IsNull(rstProjetElec.Fields("MontantForfait")) Then
1030        DR_ApercuProjet.Sections("Section2").Controls("lblForfaitElecProj").Caption = Conversion(rstProjetElec.Fields("MontantForfait"), MODE_ARGENT)
1035      Else
1040        DR_ApercuProjet.Sections("Section2").Controls("lblForfaitElecProj").Caption = Conversion("0", MODE_ARGENT)
1045      End If

1050      If Not IsNull(rstProjetElec.Fields("TauxDessin")) Then
1055        'DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecDessinProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecDessinProj").Caption) * CDbl(rstProjetElec.Fields("TauxDessin")), MODE_ARGENT)
            DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecDessinProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecDessinProj").Caption) * CDbl(50), MODE_ARGENT)
1060      Else
1065        DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecDessinProj").Caption = Conversion(0, MODE_ARGENT)
1070      End If

1075      If Not IsNull(rstProjetElec.Fields("TauxFabrication")) Then
1080        'DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFabricationProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFabricationProj").Caption) * CDbl(rstProjetElec.Fields("TauxFabrication")), MODE_ARGENT)
            DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFabricationProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFabricationProj").Caption) * CDbl(50), MODE_ARGENT)
1085      Else
1090        DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFabricationProj").Caption = Conversion(0, MODE_ARGENT)
1095      End If

1100      If Not IsNull(rstProjetElec.Fields("TauxAssemblage")) Then
1105        'DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecAssemblageProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecAssemblageProj").Caption) * CDbl(rstProjetElec.Fields("TauxAssemblage")), MODE_ARGENT)
            DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecAssemblageProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecAssemblageProj").Caption) * CDbl(50), MODE_ARGENT)
1110      Else
1115        DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecAssemblageProj").Caption = Conversion(0, MODE_ARGENT)
1120      End If

1125      If Not IsNull(rstProjetElec.Fields("TauxProgInterface")) Then
1130        'DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgInterfaceProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgInterfaceProj").Caption) * CDbl(rstProjetElec.Fields("TauxProgInterface")), MODE_ARGENT)
            DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgInterfaceProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgInterfaceProj").Caption) * CDbl(50), MODE_ARGENT)
1135      Else
1140        DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgInterfaceProj").Caption = Conversion(0, MODE_ARGENT)
1145      End If

1150      If Not IsNull(rstProjetElec.Fields("TauxProgAutomate")) Then
1155        'DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgAutomateProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgAutomateProj").Caption) * CDbl(rstProjetElec.Fields("TauxProgAutomate")), MODE_ARGENT)
            DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgAutomateProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgAutomateProj").Caption) * CDbl(50), MODE_ARGENT)
1160      Else
1165        DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgAutomateProj").Caption = Conversion(0, MODE_ARGENT)
1170      End If

1175      If Not IsNull(rstProjetElec.Fields("TauxProgRobot")) Then
1180        'DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgRobotProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgRobotProj").Caption) * CDbl(rstProjetElec.Fields("TauxProgRobot")), MODE_ARGENT)
            DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgRobotProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgRobotProj").Caption) * CDbl(50), MODE_ARGENT)
1185      Else
1190        DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecProgRobotProj").Caption = Conversion(0, MODE_ARGENT)
1195      End If

1200      If Not IsNull(rstProjetElec.Fields("TauxVision")) Then
1205        'DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecVisionProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecVisionProj").Caption) * CDbl(rstProjetElec.Fields("TauxVision")), MODE_ARGENT)
            DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecVisionProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecVisionProj").Caption) * CDbl(50), MODE_ARGENT)
1210      Else
1215        DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecVisionProj").Caption = Conversion(0, MODE_ARGENT)
1220      End If

1225      If Not IsNull(rstProjetElec.Fields("TauxTest")) Then
1230        'DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecTestProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTestProj").Caption) * CDbl(rstProjetElec.Fields("TauxTest")), MODE_ARGENT)
            DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecTestProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTestProj").Caption) * CDbl(50), MODE_ARGENT)
1235      Else
1240        DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecTestProj").Caption = Conversion(0, MODE_ARGENT)
1245      End If

1250      If Not IsNull(rstProjetElec.Fields("TauxInstallation")) Then
1255        'DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecInstallationProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecInstallationProj").Caption) * CDbl(rstProjetElec.Fields("TauxInstallation")), MODE_ARGENT)
            DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecInstallationProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecInstallationProj").Caption) * CDbl(50), MODE_ARGENT)
1260      Else
1265        DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecInstallationProj").Caption = Conversion(0, MODE_ARGENT)
1270      End If

1275      If Not IsNull(rstProjetElec.Fields("TauxMiseService")) Then
1280        'DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecMiseServiceProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecMiseServiceProj").Caption) * CDbl(rstProjetElec.Fields("TauxMiseService")), MODE_ARGENT)
            DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecMiseServiceProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecMiseServiceProj").Caption) * CDbl(50), MODE_ARGENT)
1285      Else
1290        DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecMiseServiceProj").Caption = Conversion(0, MODE_ARGENT)
1295      End If

1300      If Not IsNull(rstProjetElec.Fields("TauxFormation")) Then
1305        'DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFormationProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFormationProj").Caption) * CDbl(rstProjetElec.Fields("TauxFormation")), MODE_ARGENT)
            DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFormationProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFormationProj").Caption) * CDbl(50), MODE_ARGENT)
1310      Else
1315        DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecFormationProj").Caption = Conversion(0, MODE_ARGENT)
1320      End If

1325      If Not IsNull(rstProjetElec.Fields("TauxGestion")) Then
1330        'DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecGestionProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecGestionProj").Caption) * CDbl(rstProjetElec.Fields("TauxGestion")), MODE_ARGENT)
            DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecGestionProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecGestionProj").Caption) * CDbl(50), MODE_ARGENT)
1335      Else
1340        DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecGestionProj").Caption = Conversion(0, MODE_ARGENT)
1345      End If

1350      If Not IsNull(rstProjetElec.Fields("TauxShipping")) Then
1355        'DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecShippingProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecShippingProj").Caption) * CDbl(rstProjetElec.Fields("TauxShipping")), MODE_ARGENT)
            DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecShippingProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecShippingProj").Caption) * CDbl(50), MODE_ARGENT)
1360      Else
1365        DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecShippingProj").Caption = Conversion(0, MODE_ARGENT)
1370      End If
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




1375      Set rstProjetPieces = New ADODB.Recordset

1380      If optChoix(I_OPT_PROJET_ENTIER).Value = True Then
1385        Call rstProjetPieces.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = 'E" & sProjet & "'", g_connData, adOpenForwardOnly, adLockReadOnly)
1390      Else
1395        If Right(sProjet, 2) = "99" Then
1400          Call rstProjetPieces.Open("SELECT * FROM GRB_Projet_Pieces WHERE Left(IDProjet, 6) = '" & Left$("E" & sProjet, 6) & "' AND DateRéception BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND PieceExtraNonChargeable = False AND PieceExtraChargeable = False", g_connData, adOpenForwardOnly, adLockReadOnly)
1405        Else
1410          Call rstProjetPieces.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = 'E" & sProjet & "' AND DateRéception BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "'", g_connData, adOpenForwardOnly, adLockReadOnly)
1415        End If
1420      End If

1425      Do While Not rstProjetPieces.EOF
1430        If Trim(rstProjetPieces.Fields("Prix_total")) <> vbNullString Then
              'On additionne le prix total
1435          dblTotalPieces = dblTotalPieces + CDbl(rstProjetPieces.Fields("Prix_total")) - CDbl(rstProjetPieces.Fields("Profit_Argent"))
1440        End If

1445        Call rstProjetPieces.MoveNext
1450      Loop

1455      Call rstProjetPieces.Close
1460      Set rstProjetPieces = Nothing

1465      DR_ApercuProjet.Sections("Section2").Controls("lblPiecesElecProj").Caption = Conversion(dblTotalPieces, MODE_ARGENT)
'930       DR_ApercuProjet.Sections("Section2").Controls("lblImprevuElecProj").Caption = Conversion(rstProjetElec.Fields("total_imprevue"), MODE_ARGENT)
1470      DR_ApercuProjet.Sections("Section2").Controls("lblImprevuElecProj").Caption = Conversion(0, MODE_ARGENT)

1475      If Not IsNull(rstProjetElec.Fields("PrixEmballage")) Then
1480        dblPrixEmballage = CDbl(rstProjetElec.Fields("PrixEmballage"))
1485      Else
1490        dblPrixEmballage = 0
1495      End If
     
1500      dblTotalResteTemps = dblPrixEmballage
                                                              
1505      If IsNumeric(rstProjetElec.Fields("total_manuel")) Then
1510        dblTotalManuel = CDbl(rstProjetElec.Fields("total_manuel"))
1515      Else
1520        dblTotalManuel = 0
1525      End If

1530      DR_ApercuProjet.Sections("Section2").Controls("lblAutresElecProj").Caption = Conversion(dblTotalResteTemps + dblTotalManuel, MODE_ARGENT)

          'Calcul des totaux

          'Total des temps
1535      If DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecDessinSoum").Caption <> "---" Then
1540        dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecDessinSoum").Caption) + _
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
                       
1545        DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTotalSoum").Caption = dblTotal
1550      Else
1555        DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTotalSoum").Caption = "---"
1560      End If

1565      If DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecDessinSoum").Caption <> "---" Then
1570        dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecDessinSoum").Caption) + _
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
                       
1575        DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecTotalSoum").Caption = Conversion(dblTotal, MODE_ARGENT)
1580      Else
1585        DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecTotalSoum").Caption = "---"
1590      End If
''''''''''''''
' calcul du total des heures
'''''''''''''''''''''
1595      dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecDessinProj").Caption) + _
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
                       
1600      DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTotalProj").Caption = dblTotal
''''''''''''''''''''''''''''''
'Calcul du total de l'argent
'''''''''''''''''''''''''''''''
1605      dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecDessinProj").Caption) + _
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
                     
1610      DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecTotalProj").Caption = Conversion(dblTotal, MODE_ARGENT)

          'Calcul des prix totaux
1615      If DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecTotalSoum").Caption <> "---" Then
1620        dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecTotalSoum").Caption) + _
                       CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblPiecesElecSoum").Caption) + _
                       CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblImprevuElecSoum").Caption) + _
                       CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblAutresElecSoum").Caption)

1625        DR_ApercuProjet.Sections("Section2").Controls("lblTotalElecSoum").Caption = Conversion(dblTotal, MODE_ARGENT)
1630      Else
1635        DR_ApercuProjet.Sections("Section2").Controls("lblTotalElecSoum").Caption = "---"
1640      End If

1645      dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentElecTotalProj").Caption) + _
                     CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblPiecesElecProj").Caption) + _
                     CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblImprevuElecProj").Caption) + _
                     CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblAutresElecProj").Caption)

1650      DR_ApercuProjet.Sections("Section2").Controls("lblTotalElecProj").Caption = Conversion(dblTotal, MODE_ARGENT)
1655    End If

1660    Call rstProjetElec.Close
1665    Set rstProjetElec = Nothing

1670    Exit Sub

AfficherErreur:

1675    woups "frmChoixDateImpressionFacturation", "RemplirRapportElectrique", Err, Erl
End Sub

Private Sub RemplirRapportMecanique(ByVal sProjet As String)
  
5       On Error GoTo AfficherErreur

10      Dim rstProjetMec       As ADODB.Recordset
15      Dim rstSoumMec         As ADODB.Recordset
20      Dim rstProjetPieces    As ADODB.Recordset
25      Dim dblTotal           As Double
30      Dim bSoumission        As Boolean
35      Dim iNbrePersonne      As Integer
40      Dim dblHebergement     As Double
45      Dim dblRepas           As Double
50      Dim dblTransport       As Double
55      Dim dblUniteMobile     As Double
60      Dim dblPrixEmballage   As Double
65      Dim dblTotalResteTemps As Double
70      Dim dblTotalManuel     As Double
75      Dim dblTotalPieces     As Double
                
80      Set rstProjetMec = New ADODB.Recordset

85      Call rstProjetMec.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = 'M" & sProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

90      DR_ApercuProjet.Sections("Section2").Controls("lblProjetMec").Caption = "M" & sProjet

95      If Not rstProjetMec.EOF Then
100       bSoumission = False

105       If Not IsNull(rstProjetMec.Fields("IDSoumission")) Then
110         Set rstSoumMec = New ADODB.Recordset

115         Call rstSoumMec.Open("SELECT * FROM GRB_SoumissionMec WHERE IDSoumission = '" & rstProjetMec.Fields("IDSoumission") & "'", g_connData, adOpenDynamic, adLockOptimistic)

120         If Not rstSoumMec.EOF Then
125           bSoumission = True
130         Else
135           Call rstSoumMec.Close
140           Set rstSoumMec = Nothing
145         End If
150       End If

155       If bSoumission = True Then
160         If Not IsNull(rstSoumMec.Fields("MontantForfait")) Then
165           DR_ApercuProjet.Sections("Section2").Controls("lblForfaitMecSoum").Caption = Conversion(rstSoumMec.Fields("MontantForfait"), MODE_ARGENT)
170         Else
175           DR_ApercuProjet.Sections("Section2").Controls("lblForfaitMecSoum").Caption = Conversion("0", MODE_ARGENT)
180         End If

185         If Not IsNull(rstSoumMec.Fields("TempsDessin")) Then
190           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecDessinSoum").Caption = rstSoumMec.Fields("TempsDessin")
195           DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecDessinSoum").Caption = Conversion(CDbl(rstSoumMec.Fields("TempsDessin")) * CDbl(rstSoumMec.Fields("TauxDessin")), MODE_ARGENT)
200         Else
205           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecDessinSoum").Caption = "0"
210           DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecDessinSoum").Caption = Conversion("0", MODE_ARGENT)
215         End If

220         If Not IsNull(rstSoumMec.Fields("TempsCoupe")) Then
225           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecCoupeSoum").Caption = rstSoumMec.Fields("TempsCoupe")
230           DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecCoupeSoum").Caption = Conversion(CDbl(rstSoumMec.Fields("TempsCoupe")) * CDbl(rstSoumMec.Fields("TauxCoupe")), MODE_ARGENT)
235         Else
240           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecCoupeSoum").Caption = "0"
245           DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecCoupeSoum").Caption = Conversion("0", MODE_ARGENT)
250         End If

255         If Not IsNull(rstSoumMec.Fields("TempsMachinage")) Then
260           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecMachinageSoum").Caption = rstSoumMec.Fields("TempsMachinage")
265           DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecMachinageSoum").Caption = Conversion(CDbl(rstSoumMec.Fields("TempsMachinage")) * CDbl(rstSoumMec.Fields("TauxMachinage")), MODE_ARGENT)
270         Else
275           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecMachinageSoum").Caption = "0"
280           DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecMachinageSoum").Caption = Conversion("0", MODE_ARGENT)
285         End If

290         If Not IsNull(rstSoumMec.Fields("TempsSoudure")) Then
295           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecSoudureSoum").Caption = rstSoumMec.Fields("TempsSoudure")
300           DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecSoudureSoum").Caption = Conversion(CDbl(rstSoumMec.Fields("TempsSoudure")) * CDbl(rstSoumMec.Fields("TauxSoudure")), MODE_ARGENT)
305         Else
310           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecSoudureSoum").Caption = "0"
315           DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecSoudureSoum").Caption = Conversion("0", MODE_ARGENT)
320         End If

325         If Not IsNull(rstSoumMec.Fields("TempsAssemblage")) Then
330           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecAssemblageSoum").Caption = rstSoumMec.Fields("TempsAssemblage")
335           DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecAssemblageSoum").Caption = Conversion(CDbl(rstSoumMec.Fields("TempsAssemblage")) * CDbl(rstSoumMec.Fields("TauxAssemblage")), MODE_ARGENT)
340         Else
345           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecAssemblageSoum").Caption = "0"
350           DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecAssemblageSoum").Caption = Conversion("0", MODE_ARGENT)
355         End If

360         If Not IsNull(rstSoumMec.Fields("TempsPeinture")) Then
365           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecPeintureSoum").Caption = rstSoumMec.Fields("TempsPeinture")
370           DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecPeintureSoum").Caption = Conversion(CDbl(rstSoumMec.Fields("TempsPeinture")) * CDbl(rstSoumMec.Fields("TauxPeinture")), MODE_ARGENT)
375         Else
380           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecPeintureSoum").Caption = "0"
385           DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecPeintureSoum").Caption = Conversion("0", MODE_ARGENT)
390         End If

395         If Not IsNull(rstSoumMec.Fields("TempsTest")) Then
400           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTestSoum").Caption = rstSoumMec.Fields("TempsTest")
405           DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecTestSoum").Caption = Conversion(CDbl(rstSoumMec.Fields("TempsTest")) * CDbl(rstSoumMec.Fields("TauxTest")), MODE_ARGENT)
410         Else
415           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTestSoum").Caption = "0"
420           DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecTestSoum").Caption = Conversion("0", MODE_ARGENT)
425         End If

430         If Not IsNull(rstSoumMec.Fields("TempsInstallation")) Then
435           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecInstallationSoum").Caption = rstSoumMec.Fields("TempsInstallation")
440           DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecInstallationSoum").Caption = Conversion(CDbl(rstSoumMec.Fields("TempsInstallation")) * CDbl(rstSoumMec.Fields("TauxInstallation")), MODE_ARGENT)
445         Else
450           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecInstallationSoum").Caption = "0"
455           DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecInstallationSoum").Caption = Conversion("0", MODE_ARGENT)
460         End If

465         If Not IsNull(rstSoumMec.Fields("TempsFormation")) Then
470           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecFormationSoum").Caption = rstSoumMec.Fields("TempsFormation")
475           DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecFormationSoum").Caption = Conversion(CDbl(rstSoumMec.Fields("TempsFormation")) * CDbl(rstSoumMec.Fields("TauxFormation")), MODE_ARGENT)
480         Else
485           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecFormationSoum").Caption = "0"
490           DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecFormationSoum").Caption = Conversion("0", MODE_ARGENT)
495         End If

500         If Not IsNull(rstSoumMec.Fields("TempsGestion")) Then
505           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecGestionSoum").Caption = rstSoumMec.Fields("TempsGestion")
510           DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecGestionSoum").Caption = Conversion(CDbl(rstSoumMec.Fields("TempsGestion")) * CDbl(rstSoumMec.Fields("TauxGestion")), MODE_ARGENT)
515         Else
520           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecGestionSoum").Caption = "0"
525           DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecGestionSoum").Caption = Conversion("0", MODE_ARGENT)
530         End If

535         If Not IsNull(rstSoumMec.Fields("TempsShipping")) Then
540           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecShippingSoum").Caption = rstSoumMec.Fields("TempsShipping")
545           DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecShippingSoum").Caption = Conversion(CDbl(rstSoumMec.Fields("TempsShipping")) * CDbl(rstSoumMec.Fields("TauxShipping")), MODE_ARGENT)
550         Else
555           DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecShippingSoum").Caption = "0"
560           DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecShippingSoum").Caption = Conversion("0", MODE_ARGENT)
565         End If

570         DR_ApercuProjet.Sections("Section2").Controls("lblPiecesMecSoum").Caption = Conversion(rstSoumMec.Fields("total_piece"), MODE_ARGENT)
575         DR_ApercuProjet.Sections("Section2").Controls("lblImprevuMecSoum").Caption = Conversion(rstSoumMec.Fields("total_imprevue"), MODE_ARGENT)

580         iNbrePersonne = rstSoumMec.Fields("NbrePersonne")
           
585         Do While iNbrePersonne > 0
590           If iNbrePersonne >= 2 Then
595             dblHebergement = dblHebergement + rstSoumMec.Fields("TempsHebergement") * rstSoumMec.Fields("TauxHebergement2")
              
600             iNbrePersonne = iNbrePersonne - 2
605           Else
610             dblHebergement = dblHebergement + rstSoumMec.Fields("TempsHebergement") * rstSoumMec.Fields("TauxHebergement1")
             
615             iNbrePersonne = iNbrePersonne - 1
620           End If
625         Loop
            
630         dblRepas = CDbl(rstSoumMec.Fields("TempsRepas")) * CDbl(rstSoumMec.Fields("TauxRepas")) * CDbl(rstSoumMec.Fields("NbrePersonne"))
635         dblTransport = CDbl(rstSoumMec.Fields("TempsTransport")) * CDbl(rstSoumMec.Fields("TauxTransport"))
640         dblUniteMobile = CDbl(rstSoumMec.Fields("TempsUniteMobile")) * CDbl(rstSoumMec.Fields("TauxUniteMobile"))

645         If IsNumeric(rstSoumMec.Fields("PrixEmballage")) Then
650           dblPrixEmballage = CDbl(rstSoumMec.Fields("PrixEmballage"))
655         Else
660           dblPrixEmballage = 0
665         End If
      
670         dblTotalResteTemps = dblHebergement + dblRepas + dblTransport + dblUniteMobile + dblPrixEmballage
                                                              
675         If IsNumeric(rstSoumMec.Fields("total_manuel")) Then
680           dblTotalManuel = CDbl(rstSoumMec.Fields("total_manuel"))
685         Else
690           dblTotalManuel = 0
695         End If

700         DR_ApercuProjet.Sections("Section2").Controls("lblAutresMecSoum").Caption = Conversion(dblTotalResteTemps + dblTotalManuel, MODE_ARGENT)

705         Call rstSoumMec.Close
710         Set rstSoumMec = Nothing
715       Else
720         DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecDessinSoum").Caption = "---"
725         DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecDessinSoum").Caption = "---"

730         DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecCoupeSoum").Caption = "---"
735         DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecCoupeSoum").Caption = "---"

740         DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecMachinageSoum").Caption = "---"
745         DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecMachinageSoum").Caption = "---"

750         DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecSoudureSoum").Caption = "---"
755         DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecSoudureSoum").Caption = "---"

760         DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecAssemblageSoum").Caption = "---"
765         DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecAssemblageSoum").Caption = "---"

770         DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecPeintureSoum").Caption = "---"
775         DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecPeintureSoum").Caption = "---"

780         DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTestSoum").Caption = "---"
785         DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecTestSoum").Caption = "---"

790         DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecInstallationSoum").Caption = "---"
795         DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecInstallationSoum").Caption = "---"

800         DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecFormationSoum").Caption = "---"
805         DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecFormationSoum").Caption = "---"

810         DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecGestionSoum").Caption = "---"
815         DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecGestionSoum").Caption = "---"

820         DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecShippingSoum").Caption = "---"
825         DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecShippingSoum").Caption = "---"

830         DR_ApercuProjet.Sections("Section2").Controls("lblPiecesMecSoum").Caption = "---"
835         DR_ApercuProjet.Sections("Section2").Controls("lblImprevuMecSoum").Caption = "---"
840         DR_ApercuProjet.Sections("Section2").Controls("lblAutresMecSoum").Caption = "---"
845       End If

850       Call RemplirTempsReelsMec("M" & sProjet)

'''''''''''''''''''
'''', calcul de l'argent


855       If Not IsNull(rstProjetMec.Fields("MontantForfait")) Then
860         DR_ApercuProjet.Sections("Section2").Controls("lblForfaitMecProj").Caption = Conversion(rstProjetMec.Fields("MontantForfait"), MODE_ARGENT)
865       Else
870         DR_ApercuProjet.Sections("Section2").Controls("lblForfaitMecProj").Caption = Conversion("0", MODE_ARGENT)
875       End If

880       If Not IsNull(rstProjetMec.Fields("TauxDessin")) Then
885         'DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecDessinProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecDessinProj").Caption) * CDbl(rstProjetMec.Fields("TauxDessin")), MODE_ARGENT)
            DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecDessinProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecDessinProj").Caption) * CDbl(50), MODE_ARGENT)
890       Else
895         DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecDessinProj").Caption = Conversion("0", MODE_ARGENT)
900       End If

905       If Not IsNull(rstProjetMec.Fields("TauxCoupe")) Then
910         'DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecCoupeProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecCoupeProj").Caption) * CDbl(rstProjetMec.Fields("TauxCoupe")), MODE_ARGENT)
            DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecCoupeProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecCoupeProj").Caption) * CDbl(50), MODE_ARGENT)
915       Else
920         DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecCoupeProj").Caption = Conversion("0", MODE_ARGENT)
925       End If

930       If Not IsNull(rstProjetMec.Fields("TauxMachinage")) Then
935         'DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecMachinageProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecMachinageProj").Caption) * CDbl(rstProjetMec.Fields("TauxMachinage")), MODE_ARGENT)
            DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecMachinageProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecMachinageProj").Caption) * CDbl(50), MODE_ARGENT)
940       Else
945         DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecMachinageProj").Caption = Conversion("0", MODE_ARGENT)
950       End If

955       If Not IsNull(rstProjetMec.Fields("TauxSoudure")) Then
960         'DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecSoudureProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecSoudureProj").Caption) * CDbl(rstProjetMec.Fields("TauxSoudure")), MODE_ARGENT)
            DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecSoudureProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecSoudureProj").Caption) * CDbl(50), MODE_ARGENT)
965       Else
970         DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecSoudureProj").Caption = Conversion("0", MODE_ARGENT)
975       End If

980       If Not IsNull(rstProjetMec.Fields("TauxAssemblage")) Then
985         'DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecAssemblageProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecAssemblageProj").Caption) * CDbl(rstProjetMec.Fields("TauxAssemblage")), MODE_ARGENT)
            DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecAssemblageProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecAssemblageProj").Caption) * CDbl(50), MODE_ARGENT)
990       Else
995         DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecAssemblageProj").Caption = Conversion("0", MODE_ARGENT)
1000      End If

1005      If Not IsNull(rstProjetMec.Fields("TauxPeinture")) Then
1010        'DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecPeintureProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecPeintureProj").Caption) * CDbl(rstProjetMec.Fields("TauxPeinture")), MODE_ARGENT)
            DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecPeintureProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecPeintureProj").Caption) * CDbl(50), MODE_ARGENT)
1015      Else
1020        DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecPeintureProj").Caption = Conversion("0", MODE_ARGENT)
1025      End If

1030      If Not IsNull(rstProjetMec.Fields("TauxTest")) Then
1035        'DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecTestProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTestProj").Caption) * CDbl(rstProjetMec.Fields("TauxTest")), MODE_ARGENT)
            DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecTestProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTestProj").Caption) * CDbl(50), MODE_ARGENT)
1040      Else
1045        DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecTestProj").Caption = Conversion("0", MODE_ARGENT)
1050      End If

1055      If Not IsNull(rstProjetMec.Fields("TauxInstallation")) Then
1060        'DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecInstallationProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecInstallationProj").Caption) * CDbl(rstProjetMec.Fields("TauxInstallation")), MODE_ARGENT)
            DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecInstallationProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecInstallationProj").Caption) * CDbl(50), MODE_ARGENT)
1065      Else
1070        DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecInstallationProj").Caption = Conversion("0", MODE_ARGENT)
1075      End If

1080      If Not IsNull(rstProjetMec.Fields("TauxFormation")) Then
1085        'DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecFormationProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecFormationProj").Caption) * CDbl(rstProjetMec.Fields("TauxFormation")), MODE_ARGENT)
            DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecFormationProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecFormationProj").Caption) * CDbl(50), MODE_ARGENT)
1090      Else
1095        DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecFormationProj").Caption = Conversion("0", MODE_ARGENT)
1100      End If

1105      If Not IsNull(rstProjetMec.Fields("TauxGestion")) Then
1110        'DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecGestionProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecGestionProj").Caption) * CDbl(rstProjetMec.Fields("TauxGestion")), MODE_ARGENT)
            DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecGestionProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecGestionProj").Caption) * CDbl(50), MODE_ARGENT)
1115      Else
1120        DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecGestionProj").Caption = Conversion("0", MODE_ARGENT)
1125      End If

1130      If Not IsNull(rstProjetMec.Fields("TauxShipping")) Then
1135        'DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecShippingProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecShippingProj").Caption) * CDbl(rstProjetMec.Fields("TauxShipping")), MODE_ARGENT)
            DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecShippingProj").Caption = Conversion(CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecShippingProj").Caption) * CDbl(50), MODE_ARGENT)
1140      Else
1145        DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecShippingProj").Caption = Conversion("0", MODE_ARGENT)
1150      End If



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



1155      Set rstProjetPieces = New ADODB.Recordset

1160      If optChoix(I_OPT_PROJET_ENTIER).Value = True Then
1165        Call rstProjetPieces.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = 'M" & sProjet & "'", g_connData, adOpenForwardOnly, adLockReadOnly)
1170      Else
1175        If Right(sProjet, 2) = "99" Then
1180          Call rstProjetPieces.Open("SELECT * FROM GRB_Projet_Pieces WHERE Left(IDProjet, 6) = '" & Left$("M" & sProjet, 6) & "' AND DateRéception BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' AND PieceExtraNonChargeable = False AND PieceExtraChargeable = False", g_connData, adOpenForwardOnly, adLockReadOnly)
1185        Else
1190          Call rstProjetPieces.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = 'M" & sProjet & "' AND DateRéception BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "'", g_connData, adOpenForwardOnly, adLockReadOnly)
1195        End If
1200      End If

1205      Do While Not rstProjetPieces.EOF
1210        If Trim(rstProjetPieces.Fields("Prix_total")) <> vbNullString Then
              'On additionne le prix total
1215          dblTotalPieces = dblTotalPieces + CDbl(rstProjetPieces.Fields("Prix_total")) - CDbl(rstProjetPieces.Fields("Profit_argent"))
1220        End If

1225        Call rstProjetPieces.MoveNext
1230      Loop

1235      Call rstProjetPieces.Close
1240      Set rstProjetPieces = Nothing

1245      DR_ApercuProjet.Sections("Section2").Controls("lblPiecesMecProj").Caption = Conversion(dblTotalPieces, MODE_ARGENT)
'675       DR_ApercuProjet.Sections("Section2").Controls("lblImprevuMecProj").Caption = Conversion(rstProjetMec.Fields("total_imprevue"), MODE_ARGENT)
1250      DR_ApercuProjet.Sections("Section2").Controls("lblImprevuMecProj").Caption = Conversion(0, MODE_ARGENT)

1255      If IsNumeric(rstProjetMec.Fields("PrixEmballage")) Then
1260        dblPrixEmballage = CDbl(rstProjetMec.Fields("PrixEmballage"))
1265      Else
1270        dblPrixEmballage = 0
1275      End If
      
1280      dblTotalResteTemps = dblPrixEmballage
                                                              
1285      If IsNumeric(rstProjetMec.Fields("total_manuel")) Then
1290        dblTotalManuel = CDbl(rstProjetMec.Fields("total_manuel"))
1295      Else
1300        dblTotalManuel = 0
1305      End If

1310      DR_ApercuProjet.Sections("Section2").Controls("lblAutresMecProj").Caption = Conversion(dblTotalResteTemps + dblTotalManuel, MODE_ARGENT)

          'Calcul des totaux

          'Total des temps
1315      If DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecDessinSoum").Caption <> "---" Then
1320        dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecDessinSoum").Caption) + _
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
                       
1325        DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTotalSoum").Caption = dblTotal
1330      Else
1335        DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTotalSoum").Caption = "---"
1340      End If

1345      If DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecDessinSoum").Caption <> "---" Then
1350        dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecDessinSoum").Caption) + _
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
                       
1355        DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecTotalSoum").Caption = Conversion(dblTotal, MODE_ARGENT)
1360      Else
1365        DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecTotalSoum").Caption = "---"
1370      End If

'''''''
''' Total heure mécanique
''''''''''''''''''''''
1375      dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecDessinProj").Caption) + _
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
                     
                     
1380      DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTotalProj").Caption = dblTotal
'''''''''''''''''''''''
'''' total argent mécanique
'''''''''''''''''''''

1385      dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecDessinProj").Caption) + _
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
                     
                       
1390      DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecTotalProj").Caption = Conversion(dblTotal, MODE_ARGENT)

          'Calcul des prix totaux
1395      If DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecTotalSoum").Caption <> "---" Then
1400        dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecTotalSoum").Caption) + _
                       CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblPiecesMecSoum").Caption) + _
                       CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblImprevuMecSoum").Caption) + _
                       CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblAutresMecSoum").Caption)

1405        DR_ApercuProjet.Sections("Section2").Controls("lblTotalMecSoum").Caption = Conversion(dblTotal, MODE_ARGENT)
1410      Else
1415        DR_ApercuProjet.Sections("Section2").Controls("lblTotalMecSoum").Caption = "---"
1420      End If

1425      dblTotal = CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblArgentMecTotalProj").Caption) + _
                     CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblPiecesMecProj").Caption) + _
                     CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblImprevuMecProj").Caption) + _
                     CDbl(DR_ApercuProjet.Sections("Section2").Controls("lblAutresMecProj").Caption)

1430      DR_ApercuProjet.Sections("Section2").Controls("lblTotalMecProj").Caption = Conversion(dblTotal, MODE_ARGENT)
1435    End If

1440    Exit Sub

AfficherErreur:

1445    woups "frmChoixDateImpressionFacturation", "RemplirRapportMecanique", Err, Erl
End Sub

Private Sub RemplirTempsReelsElec(ByVal sProjet As String)
  
5       On Error GoTo AfficherErreur

10      Dim rstPunch        As ADODB.Recordset
15      Dim sDateDebut      As String
20      Dim sDateFin        As String
25      Dim sTotal          As String
30      Dim sFilterNoProjet As String
        Dim Compile1 As String
        Dim Compile2 As String
        
        Compile1 = 0
        Compile2 = 0
        

35      If Right$(sProjet, 2) = "99" Then
40        sFilterNoProjet = "LEFT(NoProjet, 6) = '" & Left$(sProjet, 6) & "'"
45      Else
50        sFilterNoProjet = "NoProjet = '" & sProjet & "'"
55      End If

60      sDateDebut = "TIMESERIAL(Left(GRB_Punch.HeureDébut,2),RIGHT(GRB_Punch.HeureDébut,2),0)"

65      sDateFin = "TIMESERIAL(Left(GRB_Punch.HeureFin,2),RIGHT(GRB_Punch.HeureFin,2),0)"

70      sTotal = "(SUM(" & sDateFin & " - " & sDateDebut & ")* 24) As Total"

75      Set rstPunch = New ADODB.Recordset

80      If optChoix(I_OPT_PROJET_ENTIER).Value = True Then
85        Call rstPunch.Open("SELECT Type, " & sTotal & " FROM GRB_Punch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null GROUP BY Type", g_connData, adOpenForwardOnly, adLockReadOnly)
90      Else
95        Call rstPunch.Open("SELECT Type, " & sTotal & " FROM GRB_Punch WHERE " & sFilterNoProjet & " AND HeureFin Is Not Null AND HeureDébut Is not Null AND [Date] BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' GROUP BY Type", g_connData, adOpenForwardOnly, adLockReadOnly)
100     End If

105     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecDessinProj").Caption = "0"
110     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFabricationProj").Caption = "0"
115     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecAssemblageProj").Caption = "0"
120     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgInterfaceProj").Caption = "0"
125     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgAutomateProj").Caption = "0"
130     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgRobotProj").Caption = "0"
135     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecVisionProj").Caption = "0"
140     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTestProj").Caption = "0"
145     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecInstallationProj").Caption = "0"
150     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecMiseServiceProj").Caption = "0"
155     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFormationProj").Caption = "0"
160     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecGestionProj").Caption = "0"
165     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecShippingProj").Caption = "0"
        DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecRechercheProj").Caption = "0"

170     Do While Not rstPunch.EOF
175       If Not IsNull(rstPunch.Fields("Total")) Then
180         Select Case rstPunch.Fields("Type")
              Case "Dessin":        DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecDessinProj").Caption = Round(rstPunch.Fields("Total"), 2)
185           Case "Fabrication":   DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFabricationProj").Caption = Round(rstPunch.Fields("Total"), 2)
190           Case "Assemblage":    DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecAssemblageProj").Caption = Round(rstPunch.Fields("Total"), 2)
195           Case "ProgInterface": DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgInterfaceProj").Caption = Round(rstPunch.Fields("Total"), 2)
200           Case "ProgAutomate":  DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgAutomateProj").Caption = Round(rstPunch.Fields("Total"), 2)
205           Case "ProgRobot":     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecProgRobotProj").Caption = Round(rstPunch.Fields("Total"), 2)
210           Case "Vision":        DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecVisionProj").Caption = Round(rstPunch.Fields("Total"), 2)
215           Case "Test":          DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecTestProj").Caption = Round(rstPunch.Fields("Total"), 2)
220           Case "Installation":  DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecInstallationProj").Caption = Round(rstPunch.Fields("Total"), 2)
225           Case "MiseService":   DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecMiseServiceProj").Caption = Round(rstPunch.Fields("Total"), 2)
230           Case "Formation":     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecFormationProj").Caption = Round(rstPunch.Fields("Total"), 2)
235           Case "Gestion":       DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecGestionProj").Caption = Round(rstPunch.Fields("Total"), 2)
240           Case "Shipping":      DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecShippingProj").Caption = Round(rstPunch.Fields("Total"), 2)
              Case "Prototypage-Dévelloppement expérimental": Compile1 = Round(rstPunch.Fields("Total"), 2)
              Case "": Compile2 = Round(rstPunch.Fields("Total"), 2)
                
            ''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''
            'ajout prototypage expérimental
            
            
            
            '''''''''''''''''''''' modif alex 6 fevrier 2012
            '''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            
245         End Select
250       End If

255       Call rstPunch.MoveNext
260     Loop

'''''''''''''''''''''''
' On addtionne develloppement avec aucun type ensemble
'''''''''''''''
DR_ApercuProjet.Sections("Section2").Controls("lblHeuresElecRechercheProj").Caption = CStr(CDbl(Compile1) + CDbl(Compile2))



265     Call rstPunch.Close
270     Set rstPunch = Nothing

275     Exit Sub

AfficherErreur:

280     woups "frmProjSoumElecTemps", "RemplirTempsReelsElec", Err, Erl
End Sub

Private Sub RemplirTempsReelsMec(ByVal sProjet As String)
  
5       On Error GoTo AfficherErreur

10      Dim rstPunch        As ADODB.Recordset
15      Dim sDateDebut      As String
20      Dim sDateFin        As String
25      Dim sTotal          As String
30      Dim sFilterNoProjet As String
        Dim test As String
        Dim Compile1 As String
        Dim Compile2 As String
        
        Compile1 = 0
        Compile2 = 0
        
35      If Right$(sProjet, 2) = "99" Then
40        sFilterNoProjet = "LEFT(NoProjet, 6) = '" & Left$(sProjet, 6) & "'"
45      Else
50        sFilterNoProjet = "NoProjet = '" & sProjet & "'"
55      End If

60      sDateDebut = "TIMESERIAL(Left(GRB_Punch.HeureDébut,2),RIGHT(GRB_Punch.HeureDébut,2),0)"

65      sDateFin = "TIMESERIAL(Left(GRB_Punch.HeureFin,2),RIGHT(GRB_Punch.HeureFin,2),0)"

70      sTotal = "(SUM(" & sDateFin & " - " & sDateDebut & ")* 24) As Total"

75      Set rstPunch = New ADODB.Recordset

80      If optChoix(I_OPT_PROJET_ENTIER).Value = True Then
85        Call rstPunch.Open("SELECT Type, " & sTotal & " FROM GRB_Punch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null GROUP BY Type", g_connData, adOpenForwardOnly, adLockReadOnly)
90      Else
95        Call rstPunch.Open("SELECT Type, " & sTotal & " FROM GRB_Punch WHERE " & sFilterNoProjet & " AND HeureFin Is Not Null AND HeureDébut Is Not Null AND [Date] BETWEEN '" & mskDateDebut.Text & "' AND '" & mskDateFin.Text & "' GROUP BY Type", g_connData, adOpenForwardOnly, adLockReadOnly)
100     End If

105     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecDessinProj").Caption = "0"
110     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecCoupeProj").Caption = "0"
115     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecMachinageProj").Caption = "0"
120     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecSoudureProj").Caption = "0"
125     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecAssemblageProj").Caption = "0"
130     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecPeintureProj").Caption = "0"
135     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTestProj").Caption = "0"
140     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecInstallationProj").Caption = "0"
145     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecFormationProj").Caption = "0"
150     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecGestionProj").Caption = "0"
155     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecShippingProj").Caption = "0"

160     Do While Not rstPunch.EOF
165       If Not IsNull(rstPunch.Fields("Total")) Then
170         Select Case rstPunch.Fields("Type")
              Case "Dessin":       DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecDessinProj").Caption = Round(rstPunch.Fields("Total"), 2)
175           Case "Coupe":        DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecCoupeProj").Caption = Round(rstPunch.Fields("Total"), 2)
180           Case "Machinage":    DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecMachinageProj").Caption = Round(rstPunch.Fields("Total"), 2)
185           Case "Soudure":      DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecSoudureProj").Caption = Round(rstPunch.Fields("Total"), 2)
190           Case "Assemblage":   DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecAssemblageProj").Caption = Round(rstPunch.Fields("Total"), 2)
195           Case "Peinture":     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecPeintureProj").Caption = Round(rstPunch.Fields("Total"), 2)
200           Case "Test":         DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecTestProj").Caption = Round(rstPunch.Fields("Total"), 2)
205           Case "Installation": DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecInstallationProj").Caption = Round(rstPunch.Fields("Total"), 2)
210           Case "Formation":    DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecFormationProj").Caption = Round(rstPunch.Fields("Total"), 2)
215           Case "Gestion":      DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecGestionProj").Caption = Round(rstPunch.Fields("Total"), 2)
220           Case "Shipping":     DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecShippingProj").Caption = Round(rstPunch.Fields("Total"), 2)
225           Case "Prototypage-Dévelloppement expérimental":     Compile1 = Round(rstPunch.Fields("Total"), 2)
              Case "":     Compile2 = Round(rstPunch.Fields("Total"), 2)
                

            End Select
230       End If

235       Call rstPunch.MoveNext
240     Loop
'''''''''''''''''''''''''''''''''
' s'il y a  des enregistrement sans type , compile sans develloppement


DR_ApercuProjet.Sections("Section2").Controls("lblHeuresMecRechercheProj").Caption = CStr(CDbl(Compile1) + CDbl(Compile2))
''''''''''''''''''''''''''''''''''''


245     Call rstPunch.Close
250     Set rstPunch = Nothing

255     Exit Sub

AfficherErreur:

260     woups "frmChoixDateImpressionFacturation", "RemplirTempsReelsElec", Err, Erl
End Sub


Private Sub Form_Click()

5       On Error GoTo AfficherErreur

10      mvwDate.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmChoixDateImpressionFacturation", "Form_Click", Err, Erl
End Sub

Private Sub Form_Load()
        
5       On Error GoTo AfficherErreur

10      optChoix(I_OPT_PROJET_ENTIER).Value = True
15      optChoixProjetEntier(0).Value = True

20      Exit Sub

AfficherErreur:

25      woups "frmChoixDateImpressionFacturation", "Form_Load", Err, Erl
End Sub

Private Sub mvwDate_LostFocus()

5       On Error GoTo AfficherErreur

10      mvwDate.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmChoixDateImpressionFacturation", "mvwDate_LostFocus", Err, Erl
End Sub

Private Sub mvwDate_DateClick(ByVal DateClicked As Date)

5       On Error GoTo AfficherErreur

10      Select Case m_eDate
          Case DEBUT: mskDateDebut.Text = ConvertDate(DateClicked)
          Case Fin:   mskDateFin.Text = ConvertDate(DateClicked)
15      End Select
  
        'Enlever le calendrier
20      mvwDate.Visible = False

25      Exit Sub

AfficherErreur:

30      woups "frmChoixDateImpressionFacturation", "mvwDate_DateClick", Err, Erl
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

35      woups "frmChoixDateImpressionFacturation", "mskDateDebut_GotFocus", Err, Erl
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

35      woups "frmChoixDateImpressionFacturation", "mskDateFin_GotFocus", Err, Erl
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

65      woups "frmChoixDateImpressionFacturation", "mskDateDebut_LostFocus", Err, Erl
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

65      woups "frmChoixDateImpressionFacturation", "mskDateFin_LostFocus", Err, Erl
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

75      woups "frmChoixDateImpressionFacturation", "cmdDateDebut_Click", Err, Erl
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

75      woups "frmChoixDateImpressionFacturation", "cmdDateFin_Click", Err, Erl
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

40      woups "frmChoixDateImpressionFacturation", "ValiderDate", Err, Erl
End Function

Private Sub optChoix_Click(Index As Integer)

5       On Error GoTo AfficherErreur

10      If optChoix(I_OPT_PROJET_ENTIER).Value = True Then
20        fra2Dates.Enabled = False
25      Else
35        fra2Dates.Enabled = True
40      End If

45      Exit Sub

AfficherErreur:

50      woups "frmChoixDateImpressionFacturation", "optChoix_Click", Err, Erl
End Sub

