VERSION 5.00
Begin VB.Form frmGroupes 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuration des groupes de sécurité"
   ClientHeight    =   7560
   ClientLeft      =   4905
   ClientTop       =   2445
   ClientWidth     =   11220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGroupes.frx":0000
   ScaleHeight     =   7560
   ScaleWidth      =   11220
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstUser 
      BackColor       =   &H00FFFFFF&
      Height          =   1230
      ItemData        =   "frmGroupes.frx":2F0D
      Left            =   9240
      List            =   "frmGroupes.frx":2F0F
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdFermer 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fermer"
      Height          =   375
      Left            =   9480
      TabIndex        =   53
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton cmdSupprimer 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Supprimer"
      Height          =   375
      Left            =   6120
      TabIndex        =   50
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton cmdModifier 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Modifier"
      Height          =   375
      Left            =   7800
      TabIndex        =   52
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton cmdAjouter 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ajouter"
      Height          =   375
      Left            =   4440
      TabIndex        =   49
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Frame fraModification 
      BackColor       =   &H80000007&
      Caption         =   "Modification"
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
      Height          =   5415
      Left            =   5160
      TabIndex        =   25
      Top             =   1440
      Width           =   5895
      Begin VB.CheckBox chkDeverrouillageTempsProjet 
         BackColor       =   &H80000012&
         Caption         =   "Déverrouillage du temps de projet"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   56
         ToolTipText     =   "Permet de faire les retours de marchandise"
         Top             =   3960
         Width           =   2775
      End
      Begin VB.CheckBox chkVerrouillageTempsProjet 
         BackColor       =   &H80000012&
         Caption         =   "Verrouillage du temps de projet"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   55
         ToolTipText     =   "Permet de faire les retours de marchandise"
         Top             =   3600
         Width           =   2535
      End
      Begin VB.CheckBox chkPunchSemaineAnterieure 
         BackColor       =   &H80000012&
         Caption         =   "Punchs dans une semaine antérieure"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   54
         ToolTipText     =   "Permet de faire les retours de marchandise"
         Top             =   3240
         Width           =   3015
      End
      Begin VB.CheckBox chkMailingList 
         BackColor       =   &H80000012&
         Caption         =   "Liste de distribution Outlook"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   47
         ToolTipText     =   "Permet de faire les retours de marchandise"
         Top             =   2880
         Width           =   2295
      End
      Begin VB.CheckBox chkRetourMarchandise 
         BackColor       =   &H80000012&
         Caption         =   "Retour de marchandise"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   45
         ToolTipText     =   "Permet de faire les retours de marchandise"
         Top             =   2520
         Width           =   2175
      End
      Begin VB.CheckBox chkReception 
         BackColor       =   &H80000012&
         Caption         =   "Réception"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   43
         ToolTipText     =   "Permet de faire la réception de marchandise"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CheckBox chkSupprimerProjets 
         BackColor       =   &H80000012&
         Caption         =   "Suppression de projets"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   41
         ToolTipText     =   "Permet de supprimer les projets"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CheckBox chkModificationPunchEmployes 
         BackColor       =   &H80000012&
         Caption         =   "Punch employés"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   44
         ToolTipText     =   "Permet de modifier la liste des employés pour qui on peut puncher"
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CheckBox chkModificationInventaireMec 
         BackColor       =   &H80000012&
         Caption         =   "Inventaire mécanique"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   46
         ToolTipText     =   "Modification de l'inventaire mécanique"
         Top             =   3960
         Width           =   1935
      End
      Begin VB.CheckBox chkModificationInventaireElec 
         BackColor       =   &H80000012&
         Caption         =   "Inventaire électrique"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   32
         ToolTipText     =   "Modification de l'inventaire électrique"
         Top             =   360
         Width           =   1935
      End
      Begin VB.CheckBox chkModificationCatalogueElec 
         BackColor       =   &H80000012&
         Caption         =   "Catalogue électrique"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   35
         ToolTipText     =   "Modication du catalogue électrique"
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox chkModificationCatalogueMec 
         BackColor       =   &H80000012&
         Caption         =   "Catalogue mécanique"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         ToolTipText     =   "Modification du catalogue mécanique"
         Top             =   4320
         Width           =   1935
      End
      Begin VB.CheckBox chkModificationBonsCommandes 
         BackColor       =   &H80000012&
         Caption         =   "Bons de commandes"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   42
         ToolTipText     =   "Modification des bons de commandes"
         Top             =   3240
         Width           =   1815
      End
      Begin VB.CheckBox chkModificationSoumissionElec 
         BackColor       =   &H80000012&
         Caption         =   "Soumissions électriques"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   37
         ToolTipText     =   "Modification des soumissions électriques"
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox chkModificationProjetElec 
         BackColor       =   &H80000012&
         Caption         =   "Projets électriques"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   39
         ToolTipText     =   "Modification des projets électriques"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox chkModificationProjetMec 
         BackColor       =   &H80000012&
         Caption         =   "Projets mécaniques"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         ToolTipText     =   "Modification des projets mécaniques"
         Top             =   5040
         Width           =   1695
      End
      Begin VB.CheckBox chkModificationSoumissionMec 
         BackColor       =   &H80000012&
         Caption         =   "Soumissions mécaniques"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         ToolTipText     =   "Modification des soumissions mécaniques mécaniques"
         Top             =   4680
         Width           =   2065
      End
      Begin VB.CheckBox chkModificationFacturation 
         BackColor       =   &H80000012&
         Caption         =   "Facturation"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   40
         ToolTipText     =   "Modification de la facturation"
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CheckBox chkModificationOutils 
         BackColor       =   &H80000012&
         Caption         =   "Outils et machinerie"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         ToolTipText     =   "Modification des outils et machinerie"
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CheckBox chkModificationFeuillesTemps 
         BackColor       =   &H80000012&
         Caption         =   "Feuilles de temps"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         ToolTipText     =   "Modification des feuilles de temps"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CheckBox chkModificationGroupes 
         BackColor       =   &H80000012&
         Caption         =   "Groupes de sécurité"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         ToolTipText     =   "Modification des groupes de sécurité"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CheckBox chkModificationEmployes 
         BackColor       =   &H80000012&
         Caption         =   "Employés"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         ToolTipText     =   "Modification des employés"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CheckBox chkModificationContacts 
         BackColor       =   &H80000012&
         Caption         =   "Contacts"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         ToolTipText     =   "Modification des contacts"
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox chkModificationFRS 
         BackColor       =   &H80000012&
         Caption         =   "Fournisseurs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         ToolTipText     =   "Modification des fournisseurs"
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox chkModificationClients 
         BackColor       =   &H80000012&
         Caption         =   "Clients"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         ToolTipText     =   "Modification des clients"
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.ComboBox cmbGroupe 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3720
      TabIndex        =   2
      Text            =   "cmbGroupe"
      Top             =   360
      Width           =   2055
   End
   Begin VB.Frame fraAffichage 
      BackColor       =   &H80000007&
      Caption         =   "Affichage"
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
      Height          =   5415
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   4935
      Begin VB.CheckBox chkAchat 
         BackColor       =   &H80000012&
         Caption         =   "Achats"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   23
         ToolTipText     =   "Affichage des achats"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CheckBox chkInventaireMec 
         BackColor       =   &H80000012&
         Caption         =   "Inventaire mécanique"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Affichage de l'inventaire mécanique"
         Top             =   3960
         Width           =   2175
      End
      Begin VB.CheckBox chkInventaireElec 
         BackColor       =   &H80000012&
         Caption         =   "Inventaire électrique"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   15
         ToolTipText     =   "Affichage de l'inventaire électrique"
         Top             =   360
         Width           =   1935
      End
      Begin VB.CheckBox chkCatalogueElec 
         BackColor       =   &H80000012&
         Caption         =   "Catalogue électrique"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   17
         ToolTipText     =   "Affichage du catalogue électrique"
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox chkSoumissionElec 
         BackColor       =   &H80000012&
         Caption         =   "Soumissions électriques"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   19
         ToolTipText     =   "Affichage des soumissions électriques"
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox chkProjetElec 
         BackColor       =   &H80000012&
         Caption         =   "Projets électriques"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   21
         ToolTipText     =   "Affichage des projets électriques"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox chkProjetMec 
         BackColor       =   &H80000012&
         Caption         =   "Projets mécaniques"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Affichage des projets mécaniques"
         Top             =   5040
         Width           =   2175
      End
      Begin VB.CheckBox chkSoumissionMec 
         BackColor       =   &H80000012&
         Caption         =   "Soumissions mécaniques"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Affichage des soumissions mécaniques"
         Top             =   4680
         Width           =   2175
      End
      Begin VB.CheckBox chkOutils 
         BackColor       =   &H80000012&
         Caption         =   "Outils entrée-sortie"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "Affichage du magasin"
         Top             =   3600
         Width           =   2055
      End
      Begin VB.CheckBox chkPunch 
         BackColor       =   &H80000012&
         Caption         =   "Punch"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Affichage du punch"
         Top             =   3240
         Width           =   2175
      End
      Begin VB.CheckBox chkConfiguration 
         BackColor       =   &H80000012&
         Caption         =   "Configuration"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Affichage de la configuration"
         Top             =   2880
         Width           =   2295
      End
      Begin VB.CheckBox chkCedule 
         BackColor       =   &H80000012&
         Caption         =   "Cédule"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Affichage de la cédule"
         Top             =   2520
         Width           =   1935
      End
      Begin VB.CheckBox chkEmployes 
         BackColor       =   &H80000012&
         Caption         =   "Employés"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Affichage des employés"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CheckBox chkCatalogueMec 
         BackColor       =   &H80000012&
         Caption         =   "Catalogue mécanique"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Affichage du catalogue mécaniqe"
         Top             =   4320
         Width           =   2175
      End
      Begin VB.CheckBox chkRapports 
         BackColor       =   &H80000012&
         Caption         =   "Rapports"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "Affichage des rapports"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CheckBox chkContactsVendeurs 
         BackColor       =   &H80000012&
         Caption         =   "Contacts pour vendeur"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Affichage des contacts pour les vendeurs"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CheckBox chkContacts 
         BackColor       =   &H80000012&
         Caption         =   "Contacts"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Affichage des contacts"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CheckBox chkFournisseurs 
         BackColor       =   &H80000012&
         Caption         =   "Fournisseurs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Affichage des fournisseurs"
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chkClients 
         BackColor       =   &H80000012&
         Caption         =   "Clients"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Affichage des clients"
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdEnregistrer 
      Caption         =   "Enregistrer"
      Height          =   375
      Left            =   4440
      TabIndex        =   48
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   6120
      TabIndex        =   51
      Top             =   7080
      Width           =   1575
   End
   Begin VB.TextBox txtGroupe 
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Liste des employés dans le groupe :"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7560
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblGroupes 
      BackStyle       =   0  'Transparent
      Caption         =   "Groupes"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmGroupes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enumMode
  MODE_AJOUT = 0
  MODE_MODIF = 1
  MODE_INACTIF = 2
End Enum

Private m_bModeAjout As Boolean
Private m_iNoGroupe  As Integer
Private m_iModif     As Integer

Private Sub chkClients_Click()

5       On Error GoTo AfficherErreur

        'Si chkClient est cliqué, la modification des clients est permise
10      If chkClients.Value = vbChecked Then
15        chkModificationClients.Enabled = True
20      Else
          'Enlève les crochets
25        chkModificationClients.Value = vbUnchecked
            
30        chkModificationClients.Enabled = False
35      End If

40      Exit Sub

AfficherErreur:

45      woups "frmGroupes", "chkClients_Click", Err, Erl
End Sub

Private Sub chkPunch_Click()

5       On Error GoTo AfficherErreur

10      If chkPunch.Value = vbChecked Then
15        chkPunchSemaineAnterieure.Enabled = True
20      Else
25        chkPunchSemaineAnterieure.Value = vbUnchecked
        
30        chkPunchSemaineAnterieure.Enabled = False
35      End If

40      Exit Sub

AfficherErreur:

45      woups "frmGroupes", "chkPunch_Click", Err, Erl
End Sub

Private Sub chkSoumissionMec_Click()

5       On Error GoTo AfficherErreur

        'Si chkSoumissionMec est cliqué, la modification des soumissions est permise
10      If chkSoumissionMec.Value = vbChecked Then
15        chkModificationSoumissionMec.Enabled = True
20      Else
          'Enlève les crochets
25        chkModificationSoumissionMec.Value = vbUnchecked
      
30        chkModificationSoumissionMec.Enabled = False
35      End If

40      Exit Sub

AfficherErreur:

45      woups "frmGroupes", "chkSoumissionMec_Click", Err, Erl
End Sub

Private Sub chkProjetMec_Click()

5       On Error GoTo AfficherErreur

        'Si chkProjetMec est cliqué, la modification des projets est permise
10      If chkProjetMec.Value = vbChecked Then
15        chkModificationProjetMec.Enabled = True
20        chkVerrouillageTempsProjet.Enabled = True
25        chkDeverrouillageTempsProjet.Enabled = True
30      Else
          'Enlève les crochets
35        chkModificationProjetMec.Value = vbUnchecked
40        chkVerrouillageTempsProjet.Value = vbUnchecked
45        chkDeverrouillageTempsProjet.Value = vbUnchecked
    
50        chkModificationProjetMec.Enabled = False
55        chkVerrouillageTempsProjet.Enabled = False
60        chkDeverrouillageTempsProjet.Enabled = False
65      End If

70      Exit Sub

AfficherErreur:

75      woups "frmGroupes", "chkProjetMec_Click", Err, Erl
End Sub

Private Sub chkSoumissionElec_Click()

5       On Error GoTo AfficherErreur

        'Si chkSoumissionElec est cliqué, la modification des soumissions est permise
10      If chkSoumissionElec.Value = vbChecked Then
15        chkModificationSoumissionElec.Enabled = True
20      Else
          'Enlève les crochets
25        chkModificationSoumissionElec.Value = vbUnchecked
      
30        chkModificationSoumissionElec.Enabled = False
35      End If

40      Exit Sub

AfficherErreur:

45      woups "frmGroupes", "chkSoumissionElec_Click", Err, Erl
End Sub

Private Sub chkProjetElec_Click()

5       On Error GoTo AfficherErreur

        'Si chkProjetElec est cliqué, la modification des projets est permise
10      If chkProjetElec.Value = vbChecked Then
15        chkModificationProjetElec.Enabled = True
20      Else
          'Enlève les crochets
25        chkModificationProjetElec.Value = vbUnchecked
    
30        chkModificationProjetElec.Enabled = False
35      End If

40      Exit Sub

AfficherErreur:

45      woups "frmGroupes", "chkProjetElec_Click", Err, Erl
End Sub

Private Sub chkOutils_Click()

5       On Error GoTo AfficherErreur

        'Si chkOutils est cliqué, la modification des outils est permise
10      If chkOutils.Value = vbChecked Then
15        chkModificationOutils.Enabled = True
20      Else
          'Enlève les crochets
25        chkModificationOutils.Value = vbUnchecked
        
30        chkModificationOutils.Enabled = False
35      End If

40      Exit Sub

AfficherErreur:

45      woups "frmGroupes", "chkOutils_Click", Err, Erl
End Sub

Private Sub chkFournisseurs_Click()

5       On Error GoTo AfficherErreur

        'Si chkFournisseur est cliqué, la modification des fournisseurs est permise
10      If chkFournisseurs.Value = vbChecked Then
15        chkModificationFRS.Enabled = True
20      Else
          'Enlève les crochets
25        chkModificationFRS.Value = vbUnchecked
           
30        chkModificationFRS.Enabled = False
35      End If

40      Exit Sub

AfficherErreur:

45      woups "frmGroupes", "chkFournisseurs_Click", Err, Erl
End Sub

Private Sub chkContacts_Click()

5       On Error GoTo AfficherErreur

        'Si chkContacts est cliqué, la modification des contacts est permise
10      If chkContacts.Value = vbChecked Then
15        chkModificationContacts.Enabled = True
20      Else
          'Enlève les crochets
25        chkModificationContacts.Value = vbUnchecked
        
30        chkModificationContacts.Enabled = False
35      End If

40      Exit Sub

AfficherErreur:

45      woups "frmGroupes", "chkContacts_Click", Err, Erl
End Sub

Private Sub chkEmployes_Click()

5       On Error GoTo AfficherErreur

        'Si chkEmployes est cliqué, la modification des employes,
        'des groupes et de la liste des punch sont permises
10      If chkEmployes.Value = vbChecked Then
15        chkModificationEmployes.Enabled = True
20        chkModificationGroupes.Enabled = True
25        chkModificationPunchEmployes.Enabled = True
30      Else
          'Enlève les crochets
35        chkModificationEmployes.Value = vbUnchecked
40        chkModificationGroupes.Value = vbUnchecked
45        chkModificationPunchEmployes.Value = vbUnchecked
        
50        chkModificationEmployes.Enabled = False
55        chkModificationGroupes.Enabled = False
60        chkModificationPunchEmployes.Enabled = False
65      End If

70      Exit Sub

AfficherErreur:

75      woups "frmGroupes", "chkEmployes_Click", Err, Erl
End Sub

Private Sub chkCatalogueElec_Click()

5       On Error GoTo AfficherErreur

        'Si chkCatalogueElec est cliqué, la modification des catalogueElec est permise
10      If chkCatalogueElec.Value = vbChecked Then
15        chkModificationCatalogueElec.Enabled = True
20      Else
25        chkModificationCatalogueElec.Value = vbUnchecked
  
30        chkModificationCatalogueElec.Enabled = False
35      End If

40      Exit Sub

AfficherErreur:

45      woups "frmGroupes", "chkCatalogueElec_Click", Err, Erl
End Sub

Private Sub chkCatalogueMec_Click()

5       On Error GoTo AfficherErreur

        'Si chkCatalogueMec est cliqué, la modification des catalogueMec est permise
10      If chkCatalogueMec.Value = vbChecked Then
15        chkModificationCatalogueMec.Enabled = True
20      Else
25        chkModificationCatalogueMec.Value = vbUnchecked
  
30        chkModificationCatalogueMec.Enabled = False
35      End If

40      Exit Sub

AfficherErreur:

45      woups "frmGroupes", "chkCatalogueMec_Click", Err, Erl
End Sub

Private Sub chkInventaireMec_Click()

5       On Error GoTo AfficherErreur

        'Si chkInventaireMec est cliqué, la modification de l'inventaireMec est permise
10      If chkInventaireMec.Value = vbChecked Then
15        chkModificationInventaireMec.Enabled = True
20      Else
25        chkModificationInventaireMec.Value = vbUnchecked
    
30        chkModificationInventaireMec.Enabled = False
35      End If

40      Exit Sub

AfficherErreur:

45      woups "frmGroupes", "chkInventaireMec_Click", Err, Erl
End Sub

Private Sub chkInventaireElec_Click()

5       On Error GoTo AfficherErreur

        'Si chkInventairElec est cliqué, la modification de l'inventaireElec est permise
10      If chkInventaireElec.Value = vbChecked Then
15        chkModificationInventaireElec.Enabled = True
20      Else
25        chkModificationInventaireElec.Value = vbUnchecked
    
30        chkModificationInventaireElec.Enabled = False
35      End If

40      Exit Sub

AfficherErreur:

45      woups "frmGroupes", "chkInventaireElec_Click", Err, Erl
End Sub

Private Sub cmbGroupe_Click()

5       On Error GoTo AfficherErreur

        'Affiche le groupe sélectionné
10      txtGroupe.Text = cmbGroupe.Text
  
15      m_iNoGroupe = cmbGroupe.ItemData(cmbGroupe.ListIndex)

20      Call AfficherGroupe
  
25      cmdModifier.Enabled = True
30      cmdsupprimer.Enabled = True

35      Exit Sub

AfficherErreur:

40      woups "frmGroupes", "cmbGroupe_Click", Err, Erl
End Sub

Private Sub Cmdajouter_Click()

5       On Error GoTo AfficherErreur

        'Met en mode ajout
10      m_bModeAjout = True
  
15      Call UncheckedAll
  
20      Call MontrerControles(MODE_AJOUT)
    
25      txtGroupe.Text = vbNullString

30      Exit Sub

AfficherErreur:

35      woups "frmGroupes", "cmdAjouter_Click", Err, Erl
End Sub

Private Sub cmdModifier_Click()

5       On Error GoTo AfficherErreur

10      If cmbGroupe.ItemData(cmbGroupe.ListIndex) <> g_iNoGroupe Then
          'Met en mode modif
15        Call MontrerControles(MODE_MODIF)
    
20        Call AfficherGroupe
    
25        m_iModif = cmbGroupe.ListIndex
    
30        m_bModeAjout = False
35      Else
40        Call MsgBox("Impossible de modifier le groupe actuel!", vbOKOnly, "Erreur")
45      End If

50      Exit Sub

AfficherErreur:

55      woups "frmGroupes", "cmdModifier_Click", Err, Erl
End Sub

Private Sub cmdAnnuler_Click()

5       On Error GoTo AfficherErreur

        'Annule l'ajout ou la modif
10      Call MontrerControles(MODE_INACTIF)
  
15      Call AfficherGroupe

20      Exit Sub

AfficherErreur:

25      woups "frmGroupes", "cmdAnnuler_Click", Err, Erl
End Sub

Private Sub cmdsupprimer_Click()

5       On Error GoTo AfficherErreur

10      If cmbGroupe.ListCount > 0 Then
          'Met la liste Up to date au cas où il y aurait des nouveaux enregistrements
15        Call AfficherUtilisateurs
    
20        If txtGroupe.Text <> S_GROUPE_ADMIN And txtGroupe.Text <> S_GROUPE_DEFAUT Then
            'Il ne faut pas effacer un groupe si il y a des utilisateurs dedans
25          If lstUser.ListCount = 0 Then
              'Efface le groupe
30            Call g_connData.Execute("DELETE * FROM GRB_Groupes WHERE NomGroupe = '" & Replace(cmbGroupe.Text, "'", "''") & "'")
      
35            Call RemplirComboGroupes
40          Else
45            Call MsgBox("Il y a des utilisateurs dans ce groupe!", vbOKOnly, "Erreur")
50          End If
55        Else
60          Call MsgBox("Vous ne pouvez pas effacer ce groupe!", vbOKOnly, "Erreur")
65        End If
70      End If

75      Exit Sub

AfficherErreur:

80      woups "frmGroupes", "cmdSupprimer_Click", Err, Erl
End Sub

Private Sub cmdEnregistrer_Click()

5       On Error GoTo AfficherErreur

        'Enregistre un ajout ou une modif
10      If txtGroupe.Text <> vbNullString Then
15        Call EnregistrerGroupe
  
20        Call MontrerControles(MODE_INACTIF)
  
25        Call RemplirComboGroupes
    
30        cmbGroupe.ListIndex = m_iModif
35      Else
40        Call MsgBox("Le nom du groupe ne peut pas être vide!", vbOKOnly, "Erreur")
45      End If

50      Exit Sub

AfficherErreur:

55      woups "frmGroupes", "cmdEnregistrer_Click", Err, Erl
End Sub

Private Sub EnregistrerGroupe()

5       On Error GoTo AfficherErreur

        'Enregistre un groupe
10      Dim rstGroupes As ADODB.Recordset
  
15      Set rstGroupes = New ADODB.Recordset
  
        'Si en mode ajout
20      If m_bModeAjout = True Then
          'Ouverture de la table GRB_Groupes
25        Call rstGroupes.Open("SELECT * FROM GRB_Groupes", g_connData, adOpenDynamic, adLockOptimistic)
    
30        Call rstGroupes.AddNew
    
35        m_bModeAjout = False
40      Else
          'Ouverture de la table GRB_Groupes avec le numéro du groupe
45        Call rstGroupes.Open("SELECT * FROM GRB_Groupes WHERE IDGroupe = " & m_iNoGroupe, g_connData, adOpenDynamic, adLockOptimistic)
50      End If
    
        'Enregistrement des valeurs
55      rstGroupes.Fields("NomGroupe").Value = txtGroupe.Text
60      rstGroupes.Fields("Clients").Value = chkClients.Value
65      rstGroupes.Fields("Fournisseurs").Value = chkFournisseurs.Value
70      rstGroupes.Fields("Contacts").Value = chkContacts.Value
75      rstGroupes.Fields("ContactsVendeurs").Value = chkContactsVendeurs.Value
80      rstGroupes.Fields("Rapport").Value = chkRapports.Value
85      rstGroupes.Fields("CatalogueMec").Value = chkCatalogueMec.Value
90      rstGroupes.Fields("CatalogueElec").Value = chkCatalogueElec.Value
95      rstGroupes.Fields("Employes").Value = chkEmployes.Value
100     rstGroupes.Fields("Cedule").Value = chkCedule.Value
105     rstGroupes.Fields("Configuration").Value = chkConfiguration.Value
110     rstGroupes.Fields("Punch").Value = chkPunch.Value
115     rstGroupes.Fields("Outils").Value = chkOutils.Value
120     rstGroupes.Fields("InventaireMec").Value = chkInventaireMec.Value
125     rstGroupes.Fields("SoumissionMec").Value = chkSoumissionMec.Value
130     rstGroupes.Fields("ProjetMec").Value = chkProjetMec.Value
135     rstGroupes.Fields("InventaireElec").Value = chkInventaireElec.Value
140     rstGroupes.Fields("SoumissionElec").Value = chkSoumissionElec.Value
145     rstGroupes.Fields("ProjetElec").Value = chkProjetElec.Value
150     rstGroupes.Fields("Achat").Value = chkAchat.Value
155     rstGroupes.Fields("ModificationClients").Value = chkModificationClients.Value
160     rstGroupes.Fields("ModificationFournisseurs").Value = chkModificationFRS.Value
165     rstGroupes.Fields("ModificationContacts").Value = chkModificationContacts.Value
170     rstGroupes.Fields("ModificationEmployes").Value = chkModificationEmployes.Value
175     rstGroupes.Fields("ModificationGroupes").Value = chkModificationGroupes.Value
180     rstGroupes.Fields("ModificationFeuillesTemps").Value = chkModificationFeuillesTemps.Value
185     rstGroupes.Fields("ModificationOutils").Value = chkModificationOutils.Value
190     rstGroupes.Fields("ModificationFacturation").Value = chkModificationFacturation.Value
195     rstGroupes.Fields("ModificationInventaireMec").Value = chkModificationInventaireMec.Value
200     rstGroupes.Fields("ModificationSoumissionsMec").Value = chkModificationSoumissionMec.Value
205     rstGroupes.Fields("ModificationProjetsMec").Value = chkModificationProjetMec.Value
210     rstGroupes.Fields("ModificationInventaireElec").Value = chkModificationInventaireElec.Value
215     rstGroupes.Fields("ModificationSoumissionsElec").Value = chkModificationSoumissionElec.Value
220     rstGroupes.Fields("ModificationProjetsElec").Value = chkModificationProjetElec.Value
225     rstGroupes.Fields("ModificationBonsCommandes").Value = chkModificationBonsCommandes.Value
230     rstGroupes.Fields("ModificationCatalogueElec").Value = chkModificationCatalogueElec.Value
235     rstGroupes.Fields("ModificationCatalogueMec").Value = chkModificationCatalogueMec.Value
240     rstGroupes.Fields("ModificationPunchEmployes").Value = chkModificationPunchEmployes.Value
245     rstGroupes.Fields("SuppressionProjets").Value = chkSupprimerProjets.Value
250     rstGroupes.Fields("ModificationReception").Value = chkReception.Value
255     rstGroupes.Fields("ModificationRetourMarchandise").Value = chkRetourMarchandise.Value
260     rstGroupes.Fields("ListeDistribution").Value = chkMailingList.Value
265     rstGroupes.Fields("PunchSemaineAntérieure").Value = chkPunchSemaineAnterieure.Value
270     rstGroupes.Fields("VerrouillageTempsProjet").Value = chkVerrouillageTempsProjet.Value
275     rstGroupes.Fields("DéverrouillageTempsProjet").Value = chkDeverrouillageTempsProjet.Value

280     Call rstGroupes.Update
  
285     Call rstGroupes.Close
290     Set rstGroupes = Nothing

295     Exit Sub

AfficherErreur:

300     woups "frmGroupes", "EnregistrerGroupe", Err, Erl
End Sub

Private Sub Cmdfermer_Click()

5       On Error GoTo AfficherErreur

        'Fermeture de la fenêtre
10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmGroupes", "cmdFermer_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

        'Ouverture de la fenêtre
10      Call RemplirComboGroupes
  
15      Call MontrerControles(MODE_INACTIF)

20      Exit Sub

AfficherErreur:

25      woups "frmGroupes", "Form_Load", Err, Erl
End Sub

Private Sub RemplirComboGroupes()

5       On Error GoTo AfficherErreur

        'Rempli le combo des groupes
10      Dim rstGroupes As ADODB.Recordset
  
        'Il faut vider le combo avant de le remplir
15      Call cmbGroupe.Clear
  
20      Set rstGroupes = New ADODB.Recordset
  
25      Call rstGroupes.Open("SELECT * FROM GRB_Groupes ORDER BY NomGroupe", g_connData, adOpenDynamic, adLockOptimistic)
    
        'Tant que ce n'est pas la fin des enregistrements
30      Do While Not rstGroupes.EOF
          'Ajout du nom du groupe dans le combo
35        Call cmbGroupe.AddItem(rstGroupes.Fields("NomGroupe"))
      
          'Ajout du numéro du groupe dans l'ItemData du combo
40        cmbGroupe.ItemData(cmbGroupe.newIndex) = rstGroupes.Fields("IDGroupe")
    
45        Call rstGroupes.MoveNext
50      Loop
    
55      Call rstGroupes.Close
60      Set rstGroupes = Nothing
   
        'Si le combo n'est pas vide, on sélectionne le premier élément
65      If cmbGroupe.ListCount > 0 Then
70        cmbGroupe.ListIndex = 0
75      End If

80      Exit Sub

AfficherErreur:

85      woups "frmGroupes", "RemplirComboGroupes", Err, Erl
End Sub

Private Sub MontrerControles(ByVal eMode As enumMode)

5       On Error GoTo AfficherErreur

        'Met les controles Enabled/Disabled
        '                  Visible/Invisible
        '                  Locked /Unlocked
10      Dim bCmbGroupe    As Boolean
15      Dim bTxtGroupe    As Boolean
20      Dim bAnnuler      As Boolean
25      Dim bEnregistrer  As Boolean
30      Dim bQuitter      As Boolean
35      Dim bAjouter      As Boolean
40      Dim bModifier     As Boolean
45      Dim bSupprimer    As Boolean
50      Dim bAffichage    As Boolean
55      Dim bModification As Boolean
60      Dim bLockedGroupe As Boolean
  
  
65      Select Case eMode
          'En mode ajout, on montre TxtGroupe,les boutons annuler et
          'enregistrer
          Case MODE_AJOUT:
70          bTxtGroupe = True
75          bAnnuler = True
80          bEnregistrer = True
85          bAffichage = True
90          bModification = True
      
         Case MODE_MODIF:
95          bTxtGroupe = True
      
100         If txtGroupe.Text = S_GROUPE_DEFAUT Then
105           bLockedGroupe = True
110         End If
      
115         bAnnuler = True
120         bEnregistrer = True
125         bAffichage = True
130         bModification = True
          
          'En mode inactif, on montre cmbGroupe, les boutons Ajouter,
          'Modifier, Supprimer et Quitter
          Case MODE_INACTIF:
135         bCmbGroupe = True
140         bQuitter = True
145         bAjouter = True
150         bModifier = True
155         bSupprimer = True
160     End Select
  
165     txtGroupe.Visible = bTxtGroupe
170     txtGroupe.Locked = bLockedGroupe
175     cmbGroupe.Visible = bCmbGroupe
180     cmdAnnuler.Visible = bAnnuler
185     cmdEnregistrer.Visible = bEnregistrer
190     Cmdfermer.Visible = bQuitter
195     Cmdajouter.Visible = bAjouter
200     cmdModifier.Visible = bModifier
205     cmdsupprimer.Visible = bSupprimer
210     fraAffichage.Enabled = bAffichage
215     fraModification.Enabled = bModification

220     Exit Sub

AfficherErreur:

225     woups "frmGroupes", "MontrerControles", Err, Erl
End Sub

Private Sub AfficherGroupe()

5       On Error GoTo AfficherErreur

        'Affiche le groupe selon la sélection dans le combo
10      Dim rstGroupes As ADODB.Recordset
 
15      Set rstGroupes = New ADODB.Recordset
 
20      Call rstGroupes.Open("SELECT * FROM GRB_Groupes WHERE IDGroupe = " & m_iNoGroupe, g_connData, adOpenDynamic, adLockOptimistic)
    
25      chkClients.Value = Abs(CInt(rstGroupes.Fields("Clients").Value))
30      chkFournisseurs.Value = Abs(CInt(rstGroupes.Fields("Fournisseurs").Value))
35      chkContacts.Value = Abs(CInt(rstGroupes.Fields("Contacts").Value))
40      chkContactsVendeurs.Value = Abs(CInt(rstGroupes.Fields("ContactsVendeurs").Value))
45      chkRapports.Value = Abs(CInt(rstGroupes.Fields("Rapport").Value))
50      chkCatalogueMec.Value = Abs(CInt(rstGroupes.Fields("CatalogueMec").Value))
55      chkCatalogueElec.Value = Abs(CInt(rstGroupes.Fields("CatalogueElec").Value))
60      chkEmployes.Value = Abs(CInt(rstGroupes.Fields("Employes").Value))
65      chkCedule.Value = Abs(CInt(rstGroupes.Fields("Cedule").Value))
70      chkConfiguration.Value = Abs(CInt(rstGroupes.Fields("Configuration").Value))
75      chkPunch.Value = Abs(CInt(rstGroupes.Fields("Punch").Value))
80      chkOutils.Value = Abs(CInt(rstGroupes.Fields("Outils").Value))
85      chkSoumissionMec.Value = Abs(CInt(rstGroupes.Fields("SoumissionMec").Value))
90      chkProjetMec.Value = Abs(CInt(rstGroupes.Fields("ProjetMec").Value))
95      chkSoumissionElec.Value = Abs(CInt(rstGroupes.Fields("SoumissionElec").Value))
100     chkProjetElec.Value = Abs(CInt(rstGroupes.Fields("ProjetElec").Value))
105     chkInventaireMec.Value = Abs(CInt(rstGroupes.Fields("InventaireMec").Value))
110     chkInventaireElec.Value = Abs(CInt(rstGroupes.Fields("InventaireElec").Value))
115     chkAchat.Value = Abs(CInt(rstGroupes.Fields("Achat").Value))
120     chkModificationFacturation.Value = Abs(CInt(rstGroupes.Fields("ModificationFacturation").Value))
125     chkModificationClients.Value = Abs(CInt(rstGroupes.Fields("ModificationClients").Value))
130     chkModificationFRS.Value = Abs(CInt(rstGroupes.Fields("ModificationFournisseurs").Value))
135     chkModificationContacts.Value = Abs(CInt(rstGroupes.Fields("ModificationContacts").Value))
140     chkModificationEmployes.Value = Abs(CInt(rstGroupes.Fields("ModificationEmployes").Value))
145     chkModificationGroupes.Value = Abs(CInt(rstGroupes.Fields("ModificationGroupes").Value))
150     chkModificationFeuillesTemps.Value = Abs(CInt(rstGroupes.Fields("ModificationFeuillesTemps").Value))
155     chkModificationOutils.Value = Abs(CInt(rstGroupes.Fields("ModificationOutils").Value))
160     chkModificationSoumissionMec.Value = Abs(CInt(rstGroupes.Fields("ModificationSoumissionsMec").Value))
165     chkModificationProjetMec.Value = Abs(CInt(rstGroupes.Fields("ModificationProjetsMec").Value))
170     chkModificationSoumissionElec.Value = Abs(CInt(rstGroupes.Fields("ModificationSoumissionsElec").Value))
175     chkModificationProjetElec.Value = Abs(CInt(rstGroupes.Fields("ModificationProjetsElec").Value))
180     chkModificationBonsCommandes.Value = Abs(CInt(rstGroupes.Fields("ModificationBonsCommandes").Value))
185     chkModificationCatalogueElec.Value = Abs(CInt(rstGroupes.Fields("ModificationCatalogueElec").Value))
190     chkModificationCatalogueMec.Value = Abs(CInt(rstGroupes.Fields("ModificationCatalogueMec").Value))
195     chkModificationInventaireMec.Value = Abs(CInt(rstGroupes.Fields("ModificationInventaireMec").Value))
200     chkModificationInventaireElec.Value = Abs(CInt(rstGroupes.Fields("ModificationInventaireElec").Value))
205     chkModificationPunchEmployes.Value = Abs(CInt(rstGroupes.Fields("ModificationPunchEmployes").Value))
210     chkSupprimerProjets.Value = Abs(CInt(rstGroupes.Fields("SuppressionProjets").Value))
215     chkReception.Value = Abs(CInt(rstGroupes.Fields("ModificationReception").Value))
220     chkRetourMarchandise.Value = Abs(CInt(rstGroupes.Fields("ModificationRetourMarchandise").Value))
225     chkMailingList.Value = Abs(CInt(rstGroupes.Fields("ListeDistribution").Value))
230     chkPunchSemaineAnterieure.Value = Abs(CInt(rstGroupes.Fields("PunchSemaineAntérieure").Value))
235     chkVerrouillageTempsProjet.Value = Abs(CInt(rstGroupes.Fields("VerrouillageTempsProjet").Value))
240     chkDeverrouillageTempsProjet.Value = Abs(CInt(rstGroupes.Fields("DéverrouillageTempsProjet").Value))

245     Call rstGroupes.Close
250     Set rstGroupes = Nothing
  
255     Call AfficherUtilisateurs

260     Exit Sub

AfficherErreur:

265     woups "frmGroupes", "AfficherGroupe", Err, Erl
End Sub

Private Sub AfficherUtilisateurs()

5       On Error GoTo AfficherErreur

        'Affiche les utilisateurs compris dans le groupe
10      Dim rstEmployes As ADODB.Recordset
  
        'Vide la liste
15      Call lstUser.Clear
  
20      Set rstEmployes = New ADODB.Recordset
  
25      Call rstEmployes.Open("SELECT * FROM GRB_employés WHERE Groupe = " & cmbGroupe.ItemData(cmbGroupe.ListIndex) & " AND Actif = true ORDER BY employe", g_connData, adOpenDynamic, adLockOptimistic)
    
        'Tant que ce n'est pas la fin des enregistrements
30      Do While Not rstEmployes.EOF
          'Ajout du nom de l'employé dans la liste
35        Call lstUser.AddItem(rstEmployes.Fields("employe"))
    
40        Call rstEmployes.MoveNext
45      Loop
  
50      Call rstEmployes.Close
55      Set rstEmployes = Nothing

60      Exit Sub

AfficherErreur:

65      woups "frmGroupes", "AfficherUtilisateurs", Err, Erl
End Sub

Private Sub UncheckedAll()

5       On Error GoTo AfficherErreur
        
        'Enlève tous les crochets des checkbox
10      Dim objControl As Object
  
15      For Each objControl In Me
20        If TypeOf objControl Is CheckBox Then
25          objControl.Value = vbUnchecked
30        End If
35      Next

40      Exit Sub

AfficherErreur:

45      woups "frmGroupes", "UncheckedAll", Err, Erl
End Sub
