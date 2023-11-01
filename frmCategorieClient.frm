VERSION 5.00
Begin VB.Form frmCategorieClient 
   BackColor       =   &H00000000&
   Caption         =   "Catégories"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6795
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3495
   ScaleWidth      =   6795
   Begin VB.Frame fraCategories 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   -120
      Width           =   6495
      Begin VB.CheckBox chkProduitsChimiques 
         BackColor       =   &H00000000&
         Caption         =   "Produits chimiques"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3480
         TabIndex        =   13
         Top             =   1920
         Width           =   2895
      End
      Begin VB.CheckBox chkICPI 
         BackColor       =   &H00000000&
         Caption         =   "ICPI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3480
         TabIndex        =   12
         Top             =   1560
         Width           =   2895
      End
      Begin VB.CheckBox chkAsphalte 
         BackColor       =   &H00000000&
         Caption         =   "Asphalte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3480
         TabIndex        =   11
         Top             =   1200
         Width           =   2895
      End
      Begin VB.CheckBox chkConsultant 
         BackColor       =   &H00000000&
         Caption         =   "Consultant"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3480
         TabIndex        =   10
         Top             =   840
         Width           =   2895
      End
      Begin VB.CheckBox chkBeton 
         BackColor       =   &H00000000&
         Caption         =   "Béton"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   480
         Width           =   2895
      End
      Begin VB.CheckBox chkPave 
         BackColor       =   &H00000000&
         Caption         =   "Pavé"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   840
         Width           =   2895
      End
      Begin VB.CheckBox chkPharmaceutique 
         BackColor       =   &H00000000&
         Caption         =   "Pharmaceutique"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   1200
         Width           =   2895
      End
      Begin VB.CheckBox chkMeuble 
         BackColor       =   &H00000000&
         Caption         =   "Meuble"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   1920
         Width           =   2895
      End
      Begin VB.CheckBox chkMeunerie 
         BackColor       =   &H00000000&
         Caption         =   "Meunerie, grain, engrais"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   2280
         Width           =   2895
      End
      Begin VB.CheckBox chkAgroalimentaire 
         BackColor       =   &H00000000&
         Caption         =   "Agroalimentaire"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   1560
         Width           =   2895
      End
      Begin VB.CheckBox chkManufacturier 
         BackColor       =   &H00000000&
         Caption         =   "Manufacturier"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3480
         TabIndex        =   3
         Top             =   480
         Width           =   2895
      End
      Begin VB.CheckBox chkAutre 
         BackColor       =   &H00000000&
         Caption         =   "Autre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3480
         TabIndex        =   2
         Top             =   2280
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   2760
      Width           =   1335
   End
End
Attribute VB_Name = "frmCategorieClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum enumTypeOuverture
  CLIENT = 0
  IMPRESSION = 1
End Enum

Private m_eOuverture As enumTypeOuverture

Private Sub Cmdfermer_Click()

5       On Error GoTo AfficherErreur

10      If m_eOuverture = CLIENT Then
15        If chkBeton.Value = vbChecked Then
20          FrmClient.m_bCategorieBeton = True
25        Else
30          FrmClient.m_bCategorieBeton = False
35        End If

40        If chkPave.Value = vbChecked Then
45          FrmClient.m_bCategoriePave = True
50        Else
55          FrmClient.m_bCategoriePave = False
60        End If

65        If chkPharmaceutique.Value = vbChecked Then
70          FrmClient.m_bCategoriePharmaceutique = True
75        Else
80          FrmClient.m_bCategoriePharmaceutique = False
85        End If

90        If chkAgroalimentaire.Value = vbChecked Then
95          FrmClient.m_bCategorieAgroalimentaire = True
100       Else
105         FrmClient.m_bCategorieAgroalimentaire = False
110       End If

115       If chkMeuble.Value = vbChecked Then
120         FrmClient.m_bCategorieMeuble = True
125       Else
130         FrmClient.m_bCategorieMeuble = False
135       End If

140       If chkMeunerie.Value = vbChecked Then
145         FrmClient.m_bCategorieMeunerie = True
150       Else
155         FrmClient.m_bCategorieMeunerie = False
160       End If

165       If chkManufacturier.Value = vbChecked Then
170         FrmClient.m_bCategorieManufacturier = True
175       Else
180         FrmClient.m_bCategorieManufacturier = False
185       End If

190       If chkConsultant.Value = vbChecked Then
195         FrmClient.m_bCategorieConsultant = True
200       Else
205         FrmClient.m_bCategorieConsultant = False
210       End If

215       If chkAsphalte.Value = vbChecked Then
220         FrmClient.m_bCategorieAsphalte = True
225       Else
230         FrmClient.m_bCategorieAsphalte = False
235       End If

240       If chkICPI.Value = vbChecked Then
245         FrmClient.m_bCategorieICPI = True
250       Else
255         FrmClient.m_bCategorieICPI = False
260       End If

265       If chkProduitsChimiques.Value = vbChecked Then
270         FrmClient.m_bCategorieProduitsChimiques = True
275       Else
280         FrmClient.m_bCategorieProduitsChimiques = False
285       End If

290       If chkAutre.Value = vbChecked Then
295         FrmClient.m_bCategorieAutre = True
300       Else
305         FrmClient.m_bCategorieAutre = False
310       End If
315     Else
320       If chkBeton.Value = vbChecked Then
325         FrmClient.m_bImpressionBeton = True
330       Else
335         FrmClient.m_bImpressionBeton = False
340       End If

345       If chkPave.Value = vbChecked Then
350         FrmClient.m_bImpressionPave = True
355       Else
360         FrmClient.m_bImpressionPave = False
365       End If

370       If chkPharmaceutique.Value = vbChecked Then
375         FrmClient.m_bImpressionPharmaceutique = True
380       Else
385         FrmClient.m_bImpressionPharmaceutique = False
390       End If

395       If chkAgroalimentaire.Value = vbChecked Then
400         FrmClient.m_bImpressionAgroAlimentaire = True
405       Else
410         FrmClient.m_bImpressionAgroAlimentaire = False
415       End If

420       If chkMeuble.Value = vbChecked Then
425         FrmClient.m_bImpressionMeuble = True
430       Else
435         FrmClient.m_bImpressionMeuble = False
440       End If

445       If chkMeunerie.Value = vbChecked Then
450         FrmClient.m_bImpressionMeunerie = True
455       Else
460         FrmClient.m_bImpressionMeunerie = False
465       End If

470       If chkManufacturier.Value = vbChecked Then
475         FrmClient.m_bImpressionManufacturier = True
480       Else
485         FrmClient.m_bImpressionManufacturier = False
490       End If

495       If chkConsultant.Value = vbChecked Then
500         FrmClient.m_bImpressionConsultant = True
505       Else
510         FrmClient.m_bImpressionConsultant = False
515       End If

520       If chkAsphalte.Value = vbChecked Then
525         FrmClient.m_bImpressionAsphalte = True
530       Else
535         FrmClient.m_bImpressionAsphalte = False
540       End If

545       If chkICPI.Value = vbChecked Then
550         FrmClient.m_bImpressionICPI = True
555       Else
560         FrmClient.m_bImpressionICPI = False
565       End If

570       If chkProduitsChimiques.Value = vbChecked Then
575         FrmClient.m_bImpressionProduitsChimiques = True
580       Else
585         FrmClient.m_bImpressionProduitsChimiques = False
590       End If

595       If chkAutre.Value = vbChecked Then
600         FrmClient.m_bImpressionAutre = True
605       Else
610         FrmClient.m_bImpressionAutre = False
615       End If
620     End If

625     Call Unload(Me)

630     Exit Sub

AfficherErreur:

635     woups "frmCategorieClient", "cmdFermer_Click", Err, Erl
End Sub

Public Sub AfficherClient()

5       On Error GoTo AfficherErreur

10      Cmdfermer.Caption = "Fermer"

15      If FrmClient.m_bCategorieBeton = True Then
20        chkBeton.Value = vbChecked
25      Else
30        chkBeton.Value = vbUnchecked
35      End If

40      If FrmClient.m_bCategoriePave = True Then
45        chkPave.Value = vbChecked
50      Else
55        chkPave.Value = vbUnchecked
60      End If

65      If FrmClient.m_bCategoriePharmaceutique = True Then
70        chkPharmaceutique.Value = vbChecked
75      Else
80        chkPharmaceutique.Value = vbUnchecked
85      End If

90      If FrmClient.m_bCategorieAgroalimentaire = True Then
95        chkAgroalimentaire.Value = vbChecked
100     Else
105       chkAgroalimentaire.Value = vbUnchecked
110     End If

115     If FrmClient.m_bCategorieMeuble = True Then
120       chkMeuble.Value = vbChecked
125     Else
130       chkMeuble.Value = vbUnchecked
135     End If

140     If FrmClient.m_bCategorieMeunerie = True Then
145       chkMeunerie.Value = vbChecked
150     Else
155       chkMeunerie.Value = vbUnchecked
160     End If

165     If FrmClient.m_bCategorieManufacturier = True Then
170       chkManufacturier.Value = vbChecked
175     Else
180       chkManufacturier.Value = vbUnchecked
185     End If

190     If FrmClient.m_bCategorieConsultant = True Then
195       chkConsultant.Value = vbChecked
200     Else
205       chkConsultant.Value = vbUnchecked
210     End If

215     If FrmClient.m_bCategorieAsphalte = True Then
220       chkAsphalte.Value = vbChecked
225     Else
230       chkAsphalte.Value = vbUnchecked
235     End If

240     If FrmClient.m_bCategorieICPI = True Then
245       chkICPI.Value = vbChecked
250     Else
255       chkICPI.Value = vbUnchecked
260     End If

265     If FrmClient.m_bCategorieProduitsChimiques = True Then
270       chkProduitsChimiques.Value = vbChecked
275     Else
280       chkProduitsChimiques.Value = vbUnchecked
285     End If

290     If FrmClient.m_bCategorieAutre = True Then
295       chkAutre.Value = vbChecked
300     Else
305       chkAutre.Value = vbUnchecked
310     End If
           
315     m_eOuverture = CLIENT
           
320     Call Me.Show(vbModal)

325     Exit Sub

AfficherErreur:

330     woups "frmCategorieClient", "AfficherClient", Err, Erl
End Sub

Public Sub AfficherImpression()

5       On Error GoTo AfficherErreur

10      Cmdfermer.Caption = "Imprimer"

15      chkBeton.Value = vbUnchecked
20      chkPave.Value = vbUnchecked
25      chkPharmaceutique.Value = vbUnchecked
30      chkAgroalimentaire.Value = vbUnchecked
35      chkMeuble.Value = vbUnchecked
40      chkMeunerie.Value = vbUnchecked
45      chkManufacturier.Value = vbUnchecked
50      chkConsultant.Value = vbUnchecked
55      chkAsphalte.Value = vbUnchecked
60      chkICPI.Value = vbUnchecked
65      chkProduitsChimiques.Value = vbUnchecked
70      chkAutre.Value = vbUnchecked

75      m_eOuverture = IMPRESSION

80      Call Me.Show(vbModal)

85      Exit Sub

AfficherErreur:

90      woups "frmCategorieClient", "AfficherImpression", Err, Erl
End Sub
