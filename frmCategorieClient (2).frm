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

 On Error GoTo Oups

 If m_eOuverture = CLIENT Then
 If chkBeton.Value = vbChecked Then
 FrmClient.m_bCategorieBeton = True
 Else
 FrmClient.m_bCategorieBeton = False
 End If

 If chkPave.Value = vbChecked Then
 FrmClient.m_bCategoriePave = True
 Else
 FrmClient.m_bCategoriePave = False
  End If

  If chkPharmaceutique.Value = vbChecked Then
  FrmClient.m_bCategoriePharmaceutique = True
  Else
  FrmClient.m_bCategoriePharmaceutique = False
  End If

  If chkAgroalimentaire.Value = vbChecked Then
  FrmClient.m_bCategorieAgroalimentaire = True
Else
FrmClient.m_bCategorieAgroalimentaire = False
 End If

 If chkMeuble.Value = vbChecked Then
 FrmClient.m_bCategorieMeuble = True
 Else
 FrmClient.m_bCategorieMeuble = False
 End If

 If chkMeunerie.Value = vbChecked Then
 FrmClient.m_bCategorieMeunerie = True
 Else
 FrmClient.m_bCategorieMeunerie = False
End If

 If chkManufacturier.Value = vbChecked Then
 FrmClient.m_bCategorieManufacturier = True
 Else
 FrmClient.m_bCategorieManufacturier = False
 End If

 If chkConsultant.Value = vbChecked Then
1  FrmClient.m_bCategorieConsultant = True
 Else
 FrmClient.m_bCategorieConsultant = False
 End If

 If chkAsphalte.Value = vbChecked Then
 FrmClient.m_bCategorieAsphalte = True
 Else
 FrmClient.m_bCategorieAsphalte = False
 End If

 If chkICPI.Value = vbChecked Then
 FrmClient.m_bCategorieICPI = True
 Else
 FrmClient.m_bCategorieICPI = False
End If

 If chkProduitsChimiques.Value = vbChecked Then
 FrmClient.m_bCategorieProduitsChimiques = True
 Else
 FrmClient.m_bCategorieProduitsChimiques = False
 End If

If chkAutre.Value = vbChecked Then
 FrmClient.m_bCategorieAutre = True
Else
FrmClient.m_bCategorieAutre = False
 End If
Else
 If chkBeton.Value = vbChecked Then
 FrmClient.m_bImpressionBeton = True
 Else
 FrmClient.m_bImpressionBeton = False
 End If

 If chkPave.Value = vbChecked Then
 FrmClient.m_bImpressionPave = True
 Else
 FrmClient.m_bImpressionPave = False
 End If

If chkPharmaceutique.Value = vbChecked Then
 FrmClient.m_bImpressionPharmaceutique = True
Else
 FrmClient.m_bImpressionPharmaceutique = False
 End If

 If chkAgroalimentaire.Value = vbChecked Then
 FrmClient.m_bImpressionAgroAlimentaire = True
4 Else
4 FrmClient.m_bImpressionAgroAlimentaire = False
4 End If

4 If chkMeuble.Value = vbChecked Then
4 FrmClient.m_bImpressionMeuble = True
4 Else
4 FrmClient.m_bImpressionMeuble = False
4 End If

4 If chkMeunerie.Value = vbChecked Then
4 FrmClient.m_bImpressionMeunerie = True
4 Else
4  FrmClient.m_bImpressionMeunerie = False
4  End If

4  If chkManufacturier.Value = vbChecked Then
4  FrmClient.m_bImpressionManufacturier = True
4  Else
4  FrmClient.m_bImpressionManufacturier = False
4  End If

4  If chkConsultant.Value = vbChecked Then
50 FrmClient.m_bImpressionConsultant = True
5 Else
 FrmClient.m_bImpressionConsultant = False
 End If

 If chkAsphalte.Value = vbChecked Then
 FrmClient.m_bImpressionAsphalte = True
 Else
 FrmClient.m_bImpressionAsphalte = False
 End If

 If chkICPI.Value = vbChecked Then
 FrmClient.m_bImpressionICPI = True
 Else
5  FrmClient.m_bImpressionICPI = False
5  End If

5  If chkProduitsChimiques.Value = vbChecked Then
5  FrmClient.m_bImpressionProduitsChimiques = True
5  Else
5  FrmClient.m_bImpressionProduitsChimiques = False
5  End If

5  If chkAutre.Value = vbChecked Then
60 FrmClient.m_bImpressionAutre = True
  Else
  FrmClient.m_bImpressionAutre = False
  End If
  End If

  Call Unload(Me)

  Exit Sub

Oups:

  wOups "frmCategorieClient", "cmdFermer_Click", Err, Err.number, Err.Description
End Sub

Public Sub AfficherClient()

 On Error GoTo Oups

 Cmdfermer.Caption = "Fermer"

 If FrmClient.m_bCategorieBeton = True Then
 chkBeton.Value = vbChecked
 Else
 chkBeton.Value = vbUnchecked
 End If

 If FrmClient.m_bCategoriePave = True Then
 chkPave.Value = vbChecked
 Else
 chkPave.Value = vbUnchecked
  End If

  If FrmClient.m_bCategoriePharmaceutique = True Then
  chkPharmaceutique.Value = vbChecked
  Else
  chkPharmaceutique.Value = vbUnchecked
  End If

  If FrmClient.m_bCategorieAgroalimentaire = True Then
  chkAgroalimentaire.Value = vbChecked
10 Else
1 chkAgroalimentaire.Value = vbUnchecked
End If

If FrmClient.m_bCategorieMeuble = True Then
 chkMeuble.Value = vbChecked
Else
 chkMeuble.Value = vbUnchecked
End If

If FrmClient.m_bCategorieMeunerie = True Then
 chkMeunerie.Value = vbChecked
Else
 chkMeunerie.Value = vbUnchecked
1  End If

If FrmClient.m_bCategorieManufacturier = True Then
 chkManufacturier.Value = vbChecked
Else
 chkManufacturier.Value = vbUnchecked
End If

 If FrmClient.m_bCategorieConsultant = True Then
1  chkConsultant.Value = vbChecked
 Else
 chkConsultant.Value = vbUnchecked
End If

If FrmClient.m_bCategorieAsphalte = True Then
 chkAsphalte.Value = vbChecked
Else
 chkAsphalte.Value = vbUnchecked
End If

If FrmClient.m_bCategorieICPI = True Then
 chkICPI.Value = vbChecked
Else
 chkICPI.Value = vbUnchecked
2  End If

If FrmClient.m_bCategorieProduitsChimiques = True Then
chkProduitsChimiques.Value = vbChecked
Else
chkProduitsChimiques.Value = vbUnchecked
End If

2  If FrmClient.m_bCategorieAutre = True Then
 chkAutre.Value = vbChecked
30 Else
3 chkAutre.Value = vbUnchecked
End If
 
m_eOuverture = CLIENT
 
Call Me.Show(vbModal)

Exit Sub

Oups:

wOups "frmCategorieClient", "AfficherClient", Err, Err.number, Err.Description
End Sub

Public Sub AfficherImpression()

 On Error GoTo Oups

 Cmdfermer.Caption = "Imprimer"

 chkBeton.Value = vbUnchecked
 chkPave.Value = vbUnchecked
 chkPharmaceutique.Value = vbUnchecked
 chkAgroalimentaire.Value = vbUnchecked
 chkMeuble.Value = vbUnchecked
 chkMeunerie.Value = vbUnchecked
 chkManufacturier.Value = vbUnchecked
 chkConsultant.Value = vbUnchecked
 chkAsphalte.Value = vbUnchecked
  chkICPI.Value = vbUnchecked
  chkProduitsChimiques.Value = vbUnchecked
  chkAutre.Value = vbUnchecked

  m_eOuverture = IMPRESSION

  Call Me.Show(vbModal)

  Exit Sub

Oups:

  wOups "frmCategorieClient", "AfficherImpression", Err, Err.number, Err.Description
End Sub
