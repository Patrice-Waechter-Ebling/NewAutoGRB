VERSION 5.00
Begin VB.Form frmGroupes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuration des groupes de sécurité"
   ClientHeight    =   7560
   ClientLeft      =   4905
   ClientTop       =   2445
   ClientWidth     =   11220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   11220
   Begin VB.ListBox lstUser 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   1200
      ItemData        =   "frmGroupes.frx":0000
      Left            =   9240
      List            =   "frmGroupes.frx":0002
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
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   5160
      TabIndex        =   25
      Top             =   1440
      Width           =   5895
      Begin VB.CheckBox chkDeverrouillageTempsProjet 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Déverrouillage du temps de projet"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   56
         ToolTipText     =   "Permet de faire les retours de marchandise"
         Top             =   3960
         Width           =   2775
      End
      Begin VB.CheckBox chkVerrouillageTempsProjet 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Verrouillage du temps de projet"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   55
         ToolTipText     =   "Permet de faire les retours de marchandise"
         Top             =   3600
         Width           =   2535
      End
      Begin VB.CheckBox chkPunchSemaineAnterieure 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Punchs dans une semaine antérieure"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   54
         ToolTipText     =   "Permet de faire les retours de marchandise"
         Top             =   3240
         Width           =   3015
      End
      Begin VB.CheckBox chkMailingList 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Liste de distribution Outlook"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   47
         ToolTipText     =   "Permet de faire les retours de marchandise"
         Top             =   2880
         Width           =   2295
      End
      Begin VB.CheckBox chkRetourMarchandise 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Retour de marchandise"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   45
         ToolTipText     =   "Permet de faire les retours de marchandise"
         Top             =   2520
         Width           =   2175
      End
      Begin VB.CheckBox chkReception 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Réception"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   43
         ToolTipText     =   "Permet de faire la réception de marchandise"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CheckBox chkSupprimerProjets 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Suppression de projets"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   41
         ToolTipText     =   "Permet de supprimer les projets"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CheckBox chkModificationPunchEmployes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Punch employés"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   44
         ToolTipText     =   "Permet de modifier la liste des employés pour qui on peut puncher"
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CheckBox chkModificationInventaireMec 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Inventaire mécanique"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   46
         ToolTipText     =   "Modification de l'inventaire mécanique"
         Top             =   3960
         Width           =   1935
      End
      Begin VB.CheckBox chkModificationInventaireElec 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Inventaire électrique"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   32
         ToolTipText     =   "Modification de l'inventaire électrique"
         Top             =   360
         Width           =   1935
      End
      Begin VB.CheckBox chkModificationCatalogueElec 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Catalogue électrique"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   35
         ToolTipText     =   "Modication du catalogue électrique"
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox chkModificationCatalogueMec 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Catalogue mécanique"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         ToolTipText     =   "Modification du catalogue mécanique"
         Top             =   4320
         Width           =   1935
      End
      Begin VB.CheckBox chkModificationBonsCommandes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Bons de commandes"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   42
         ToolTipText     =   "Modification des bons de commandes"
         Top             =   3240
         Width           =   1815
      End
      Begin VB.CheckBox chkModificationSoumissionElec 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Soumissions électriques"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   37
         ToolTipText     =   "Modification des soumissions électriques"
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox chkModificationProjetElec 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Projets électriques"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   39
         ToolTipText     =   "Modification des projets électriques"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox chkModificationProjetMec 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Projets mécaniques"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         ToolTipText     =   "Modification des projets mécaniques"
         Top             =   5040
         Width           =   1695
      End
      Begin VB.CheckBox chkModificationSoumissionMec 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Soumissions mécaniques"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         ToolTipText     =   "Modification des soumissions mécaniques mécaniques"
         Top             =   4680
         Width           =   2065
      End
      Begin VB.CheckBox chkModificationFacturation 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Facturation"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   40
         ToolTipText     =   "Modification de la facturation"
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CheckBox chkModificationOutils 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Outils et machinerie"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         ToolTipText     =   "Modification des outils et machinerie"
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CheckBox chkModificationFeuillesTemps 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Feuilles de temps"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         ToolTipText     =   "Modification des feuilles de temps"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CheckBox chkModificationGroupes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Groupes de sécurité"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         ToolTipText     =   "Modification des groupes de sécurité"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CheckBox chkModificationEmployes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Employés"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         ToolTipText     =   "Modification des employés"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CheckBox chkModificationContacts 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Contacts"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         ToolTipText     =   "Modification des contacts"
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox chkModificationFRS 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fournisseurs"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         ToolTipText     =   "Modification des fournisseurs"
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox chkModificationClients 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Clients"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         ToolTipText     =   "Modification des clients"
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.ComboBox cmbGroupe 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3720
      TabIndex        =   2
      Text            =   "cmbGroupe"
      Top             =   360
      Width           =   2055
   End
   Begin VB.Frame fraAffichage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   4935
      Begin VB.CheckBox chkAchat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Achats"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   23
         ToolTipText     =   "Affichage des achats"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CheckBox chkInventaireMec 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Inventaire mécanique"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Affichage de l'inventaire mécanique"
         Top             =   3960
         Width           =   2175
      End
      Begin VB.CheckBox chkInventaireElec 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Inventaire électrique"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   15
         ToolTipText     =   "Affichage de l'inventaire électrique"
         Top             =   360
         Width           =   1935
      End
      Begin VB.CheckBox chkCatalogueElec 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Catalogue électrique"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   17
         ToolTipText     =   "Affichage du catalogue électrique"
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox chkSoumissionElec 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Soumissions électriques"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   19
         ToolTipText     =   "Affichage des soumissions électriques"
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox chkProjetElec 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Projets électriques"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   21
         ToolTipText     =   "Affichage des projets électriques"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox chkProjetMec 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Projets mécaniques"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Affichage des projets mécaniques"
         Top             =   5040
         Width           =   2175
      End
      Begin VB.CheckBox chkSoumissionMec 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Soumissions mécaniques"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Affichage des soumissions mécaniques"
         Top             =   4680
         Width           =   2175
      End
      Begin VB.CheckBox chkOutils 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Outils entrée-sortie"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "Affichage du magasin"
         Top             =   3600
         Width           =   2055
      End
      Begin VB.CheckBox chkPunch 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Punch"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Affichage du punch"
         Top             =   3240
         Width           =   2175
      End
      Begin VB.CheckBox chkConfiguration 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Configuration"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Affichage de la configuration"
         Top             =   2880
         Width           =   2295
      End
      Begin VB.CheckBox chkCedule 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cédule"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Affichage de la cédule"
         Top             =   2520
         Width           =   1935
      End
      Begin VB.CheckBox chkEmployes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Employés"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Affichage des employés"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CheckBox chkCatalogueMec 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Catalogue mécanique"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Affichage du catalogue mécaniqe"
         Top             =   4320
         Width           =   2175
      End
      Begin VB.CheckBox chkRapports 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Rapports"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "Affichage des rapports"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CheckBox chkContactsVendeurs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Contacts pour vendeur"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Affichage des contacts pour les vendeurs"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CheckBox chkContacts 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Contacts"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Affichage des contacts"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CheckBox chkFournisseurs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fournisseurs"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Affichage des fournisseurs"
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chkClients 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Clients"
         ForeColor       =   &H80000008&
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
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Liste des employés dans le groupe :"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7560
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblGroupes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Groupes"
      ForeColor       =   &H80000008&
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
Private m_iNoGroupe As Integer
Private m_iModif As Integer

Private Sub chkClients_Click()

 On Error GoTo Oups

 'Si chkClient est cliqué, la modification des clients est permise
 If chkClients.Value = vbChecked Then
 chkModificationClients.Enabled = True
 Else
 'Enlève les crochets
 chkModificationClients.Value = vbUnchecked
 
 chkModificationClients.Enabled = False
 End If

 Exit Sub

Oups:

 wOups "frmGroupes", "chkClients_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkPunch_Click()

 On Error GoTo Oups

 If chkPunch.Value = vbChecked Then
 chkPunchSemaineAnterieure.Enabled = True
 Else
 chkPunchSemaineAnterieure.Value = vbUnchecked
 
 chkPunchSemaineAnterieure.Enabled = False
 End If

 Exit Sub

Oups:

 wOups "frmGroupes", "chkPunch_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkSoumissionMec_Click()

 On Error GoTo Oups

 'Si chkSoumissionMec est cliqué, la modification des soumissions est permise
 If chkSoumissionMec.Value = vbChecked Then
 chkModificationSoumissionMec.Enabled = True
 Else
 'Enlève les crochets
 chkModificationSoumissionMec.Value = vbUnchecked
 
 chkModificationSoumissionMec.Enabled = False
 End If

 Exit Sub

Oups:

 wOups "frmGroupes", "chkSoumissionMec_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkProjetMec_Click()

 On Error GoTo Oups

 'Si chkProjetMec est cliqué, la modification des projets est permise
 If chkProjetMec.Value = vbChecked Then
 chkModificationProjetMec.Enabled = True
 chkVerrouillageTempsProjet.Enabled = True
 chkDeverrouillageTempsProjet.Enabled = True
 Else
 'Enlève les crochets
 chkModificationProjetMec.Value = vbUnchecked
 chkVerrouillageTempsProjet.Value = vbUnchecked
 chkDeverrouillageTempsProjet.Value = vbUnchecked
 
 chkModificationProjetMec.Enabled = False
 chkVerrouillageTempsProjet.Enabled = False
  chkDeverrouillageTempsProjet.Enabled = False
  End If

  Exit Sub

Oups:

  wOups "frmGroupes", "chkProjetMec_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkSoumissionElec_Click()

 On Error GoTo Oups

 'Si chkSoumissionElec est cliqué, la modification des soumissions est permise
 If chkSoumissionElec.Value = vbChecked Then
 chkModificationSoumissionElec.Enabled = True
 Else
 'Enlève les crochets
 chkModificationSoumissionElec.Value = vbUnchecked
 
 chkModificationSoumissionElec.Enabled = False
 End If

 Exit Sub

Oups:

 wOups "frmGroupes", "chkSoumissionElec_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkProjetElec_Click()

 On Error GoTo Oups

 'Si chkProjetElec est cliqué, la modification des projets est permise
 If chkProjetElec.Value = vbChecked Then
 chkModificationProjetElec.Enabled = True
 Else
 'Enlève les crochets
 chkModificationProjetElec.Value = vbUnchecked
 
 chkModificationProjetElec.Enabled = False
 End If

 Exit Sub

Oups:

 wOups "frmGroupes", "chkProjetElec_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkOutils_Click()

 On Error GoTo Oups

 'Si chkOutils est cliqué, la modification des outils est permise
 If chkOutils.Value = vbChecked Then
 chkModificationOutils.Enabled = True
 Else
 'Enlève les crochets
 chkModificationOutils.Value = vbUnchecked
 
 chkModificationOutils.Enabled = False
 End If

 Exit Sub

Oups:

 wOups "frmGroupes", "chkOutils_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkFournisseurs_Click()

 On Error GoTo Oups

 'Si chkFournisseur est cliqué, la modification des fournisseurs est permise
 If chkFournisseurs.Value = vbChecked Then
 chkModificationFRS.Enabled = True
 Else
 'Enlève les crochets
 chkModificationFRS.Value = vbUnchecked
 
 chkModificationFRS.Enabled = False
 End If

 Exit Sub

Oups:

 wOups "frmGroupes", "chkFournisseurs_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkContacts_Click()

 On Error GoTo Oups

 'Si chkContacts est cliqué, la modification des contacts est permise
 If chkContacts.Value = vbChecked Then
 chkModificationContacts.Enabled = True
 Else
 'Enlève les crochets
 chkModificationContacts.Value = vbUnchecked
 
 chkModificationContacts.Enabled = False
 End If

 Exit Sub

Oups:

 wOups "frmGroupes", "chkContacts_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkEmployes_Click()

 On Error GoTo Oups

 'Si chkEmployes est cliqué, la modification des employes,
 'des groupes et de la liste des punch sont permises
 If chkEmployes.Value = vbChecked Then
 chkModificationEmployes.Enabled = True
 chkModificationGroupes.Enabled = True
 chkModificationPunchEmployes.Enabled = True
 Else
 'Enlève les crochets
 chkModificationEmployes.Value = vbUnchecked
 chkModificationGroupes.Value = vbUnchecked
 chkModificationPunchEmployes.Value = vbUnchecked
 
 chkModificationEmployes.Enabled = False
 chkModificationGroupes.Enabled = False
  chkModificationPunchEmployes.Enabled = False
  End If

  Exit Sub

Oups:

  wOups "frmGroupes", "chkEmployes_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkCatalogueElec_Click()

 On Error GoTo Oups

 'Si chkCatalogueElec est cliqué, la modification des catalogueElec est permise
 If chkCatalogueElec.Value = vbChecked Then
 chkModificationCatalogueElec.Enabled = True
 Else
 chkModificationCatalogueElec.Value = vbUnchecked
 
 chkModificationCatalogueElec.Enabled = False
 End If

 Exit Sub

Oups:

 wOups "frmGroupes", "chkCatalogueElec_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkCatalogueMec_Click()

 On Error GoTo Oups

 'Si chkCatalogueMec est cliqué, la modification des catalogueMec est permise
 If chkCatalogueMec.Value = vbChecked Then
 chkModificationCatalogueMec.Enabled = True
 Else
 chkModificationCatalogueMec.Value = vbUnchecked
 
 chkModificationCatalogueMec.Enabled = False
 End If

 Exit Sub

Oups:

 wOups "frmGroupes", "chkCatalogueMec_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkInventaireMec_Click()

 On Error GoTo Oups

 'Si chkInventaireMec est cliqué, la modification de l'inventaireMec est permise
 If chkInventaireMec.Value = vbChecked Then
 chkModificationInventaireMec.Enabled = True
 Else
 chkModificationInventaireMec.Value = vbUnchecked
 
 chkModificationInventaireMec.Enabled = False
 End If

 Exit Sub

Oups:

 wOups "frmGroupes", "chkInventaireMec_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkInventaireElec_Click()

 On Error GoTo Oups

 'Si chkInventairElec est cliqué, la modification de l'inventaireElec est permise
 If chkInventaireElec.Value = vbChecked Then
 chkModificationInventaireElec.Enabled = True
 Else
 chkModificationInventaireElec.Value = vbUnchecked
 
 chkModificationInventaireElec.Enabled = False
 End If

 Exit Sub

Oups:

 wOups "frmGroupes", "chkInventaireElec_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmbGroupe_Click()

 On Error GoTo Oups

 'Affiche le groupe sélectionné
 txtGroupe.Text = cmbGroupe.Text
 
 m_iNoGroupe = cmbGroupe.ItemData(cmbGroupe.ListIndex)

 Call AfficherGroupe
 
 cmdModifier.Enabled = True
 cmdsupprimer.Enabled = True

 Exit Sub

Oups:

 wOups "frmGroupes", "cmbGroupe_Click", Err, Err.number, Err.Description
End Sub

Private Sub Cmdajouter_Click()

 On Error GoTo Oups

 'Met en mode ajout
 m_bModeAjout = True
 
 Call UncheckedAll
 
 Call MontrerControles(MODE_AJOUT)
 
 txtGroupe.Text = vbNullString

 Exit Sub

Oups:

 wOups "frmGroupes", "cmdAjouter_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdModifier_Click()

 On Error GoTo Oups

 If cmbGroupe.ItemData(cmbGroupe.ListIndex) <> g_iNoGroupe Then
 'Met en mode modif
 Call MontrerControles(MODE_MODIF)
 
 Call AfficherGroupe
 
 m_iModif = cmbGroupe.ListIndex
 
 m_bModeAjout = False
 Else
 Call MsgBox("Impossible de modifier le groupe actuel!", vbOKOnly, "Erreur")
 End If

 Exit Sub

Oups:

 wOups "frmGroupes", "cmdModifier_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnuler_Click()

 On Error GoTo Oups

 'Annule l'ajout ou la modif
 Call MontrerControles(MODE_INACTIF)
 
 Call AfficherGroupe

 Exit Sub

Oups:

 wOups "frmGroupes", "cmdAnnuler_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdsupprimer_Click()

 On Error GoTo Oups

 If cmbGroupe.ListCount > 0 Then
 'Met la liste Up to date au cas où il y aurait des nouveaux enregistrements
 Call AfficherUtilisateurs
 
 If txtGroupe.Text <> S_GROUPE_ADMIN And txtGroupe.Text <> S_GROUPE_DEFAUT Then
 'Il ne faut pas effacer un groupe si il y a des utilisateurs dedans
 If lstUser.ListCount = 0 Then
 'Efface le groupe
 Call g_connData.Execute("DELETE * FROM GrbGroupes WHERE NomGroupe = '" & Replace(cmbGroupe.Text, "'", "''") & "'")
 
 Call RemplirComboGroupes
 Else
 Call MsgBox("Il y a des utilisateurs dans ce groupe!", vbOKOnly, "Erreur")
 End If
 Else
  Call MsgBox("Vous ne pouvez pas effacer ce groupe!", vbOKOnly, "Erreur")
  End If
  End If

  Exit Sub

Oups:

  wOups "frmGroupes", "cmdSupprimer_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdEnregistrer_Click()

 On Error GoTo Oups

 'Enregistre un ajout ou une modif
 If txtGroupe.Text <> vbNullString Then
 Call EnregistrerGroupe
 
 Call MontrerControles(MODE_INACTIF)
 
 Call RemplirComboGroupes
 
 cmbGroupe.ListIndex = m_iModif
 Else
 Call MsgBox("Le nom du groupe ne peut pas être vide!", vbOKOnly, "Erreur")
 End If

 Exit Sub

Oups:

 wOups "frmGroupes", "cmdEnregistrer_Click", Err, Err.number, Err.Description
End Sub

Private Sub EnregistrerGroupe()

 On Error GoTo Oups

 'Enregistre un groupe
 Dim rstGroupes As ADODB.Recordset
 
 Set rstGroupes = New ADODB.Recordset
 
 'Si en mode ajout
 If m_bModeAjout = True Then
 'Ouverture de la table GrbGroupes
 Call rstGroupes.Open("SELECT * FROM GrbGroupes", g_connData, adOpenDynamic, adLockOptimistic)
 
 Call rstGroupes.AddNew
 
 m_bModeAjout = False
 Else
 'Ouverture de la table GrbGroupes avec le numéro du groupe
 Call rstGroupes.Open("SELECT * FROM GrbGroupes WHERE IDGroupe = " & m_iNoGroupe, g_connData, adOpenDynamic, adLockOptimistic)
 End If
 
 'Enregistrement des valeurs
 rstGroupes.Fields("NomGroupe").Value = txtGroupe.Text
  rstGroupes.Fields("Clients").Value = chkClients.Value
  rstGroupes.Fields("Fournisseurs").Value = chkFournisseurs.Value
  rstGroupes.Fields("Contacts").Value = chkContacts.Value
  rstGroupes.Fields("ContactsVendeurs").Value = chkContactsVendeurs.Value
  rstGroupes.Fields("Rapport").Value = chkRapports.Value
  rstGroupes.Fields("CatalogueMec").Value = chkCatalogueMec.Value
  rstGroupes.Fields("CatalogueElec").Value = chkCatalogueElec.Value
  rstGroupes.Fields("Employes").Value = chkEmployes.Value
10 rstGroupes.Fields("Cedule").Value = chkCedule.Value
rstGroupes.Fields("Configuration").Value = chkConfiguration.Value
rstGroupes.Fields("Punch").Value = chkPunch.Value
rstGroupes.Fields("Outils").Value = chkOutils.Value
rstGroupes.Fields("InventaireMec").Value = chkInventaireMec.Value
rstGroupes.Fields("SoumissionMec").Value = chkSoumissionMec.Value
rstGroupes.Fields("ProjetMec").Value = chkProjetMec.Value
rstGroupes.Fields("InventaireElec").Value = chkInventaireElec.Value
rstGroupes.Fields("SoumissionElec").Value = chkSoumissionElec.Value
rstGroupes.Fields("ProjetElec").Value = chkProjetElec.Value
rstGroupes.Fields("Achat").Value = chkAchat.Value
rstGroupes.Fields("ModificationClients").Value = chkModificationClients.Value
1  rstGroupes.Fields("ModificationFournisseurs").Value = chkModificationFRS.Value
rstGroupes.Fields("ModificationContacts").Value = chkModificationContacts.Value
 rstGroupes.Fields("ModificationEmployes").Value = chkModificationEmployes.Value
rstGroupes.Fields("ModificationGroupes").Value = chkModificationGroupes.Value
 rstGroupes.Fields("ModificationFeuillesTemps").Value = chkModificationFeuillesTemps.Value
rstGroupes.Fields("ModificationOutils").Value = chkModificationOutils.Value
 rstGroupes.Fields("ModificationFacturation").Value = chkModificationFacturation.Value
1  rstGroupes.Fields("ModificationInventaireMec").Value = chkModificationInventaireMec.Value
 rstGroupes.Fields("ModificationSoumissionsMec").Value = chkModificationSoumissionMec.Value
 rstGroupes.Fields("ModificationProjetsMec").Value = chkModificationProjetMec.Value
rstGroupes.Fields("ModificationInventaireElec").Value = chkModificationInventaireElec.Value
rstGroupes.Fields("ModificationSoumissionsElec").Value = chkModificationSoumissionElec.Value
rstGroupes.Fields("ModificationProjetsElec").Value = chkModificationProjetElec.Value
rstGroupes.Fields("ModificationBonsCommandes").Value = chkModificationBonsCommandes.Value
rstGroupes.Fields("ModificationCatalogueElec").Value = chkModificationCatalogueElec.Value
rstGroupes.Fields("ModificationCatalogueMec").Value = chkModificationCatalogueMec.Value
rstGroupes.Fields("ModificationPunchEmployes").Value = chkModificationPunchEmployes.Value
rstGroupes.Fields("SuppressionProjets").Value = chkSupprimerProjets.Value
rstGroupes.Fields("ModificationReception").Value = chkReception.Value
rstGroupes.Fields("ModificationRetourMarchandise").Value = chkRetourMarchandise.Value
2  rstGroupes.Fields("ListeDistribution").Value = chkMailingList.Value
rstGroupes.Fields("PunchSemaineAntérieure").Value = chkPunchSemaineAnterieure.Value
2  rstGroupes.Fields("VerrouillageTempsProjet").Value = chkVerrouillageTempsProjet.Value
rstGroupes.Fields("DéverrouillageTempsProjet").Value = chkDeverrouillageTempsProjet.Value

2  Call rstGroupes.Update
 
Call rstGroupes.Close
2  Set rstGroupes = Nothing

Exit Sub

Oups:

30 wOups "frmGroupes", "EnregistrerGroupe", Err, Err.number, Err.Description
End Sub

Private Sub Cmdfermer_Click()

 On Error GoTo Oups

 'Fermeture de la fenêtre
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmGroupes", "cmdFermer_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 'Ouverture de la fenêtre
 Call RemplirComboGroupes
 
 Call MontrerControles(MODE_INACTIF)

 Exit Sub

Oups:

 wOups "frmGroupes", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboGroupes()

 On Error GoTo Oups

 'Rempli le combo des groupes
 Dim rstGroupes As ADODB.Recordset
 
 'Il faut vider le combo avant de le remplir
 Call cmbGroupe.Clear
 
 Set rstGroupes = New ADODB.Recordset
 
 Call rstGroupes.Open("SELECT * FROM GrbGroupes ORDER BY NomGroupe", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Tant que ce n'est pas la fin des enregistrements
 Do While Not rstGroupes.EOF
 'Ajout du nom du groupe dans le combo
 Call cmbGroupe.AddItem(rstGroupes.Fields("NomGroupe"))
 
 'Ajout du numéro du groupe dans l'ItemData du combo
 cmbGroupe.ItemData(cmbGroupe.newIndex) = rstGroupes.Fields("IDGroupe")
 
 Call rstGroupes.MoveNext
 Loop
 
 Call rstGroupes.Close
  Set rstGroupes = Nothing
 
 'Si le combo n'est pas vide, on sélectionne le premier élément
  If cmbGroupe.ListCount > 0 Then
  cmbGroupe.ListIndex = 0
  End If

  Exit Sub

Oups:

  wOups "frmGroupes", "RemplirComboGroupes", Err, Err.number, Err.Description
End Sub

Private Sub MontrerControles(ByVal eMode As enumMode)

 On Error GoTo Oups

 'Met les controles Enabled/Disabled
 ' Visible/Invisible
 ' Locked /Unlocked
 Dim bCmbGroupe As Boolean
 Dim bTxtGroupe As Boolean
 Dim bAnnuler As Boolean
 Dim bEnregistrer As Boolean
 Dim bQuitter As Boolean
 Dim bAjouter As Boolean
 Dim bModifier As Boolean
 Dim bSupprimer As Boolean
 Dim bAffichage As Boolean
 Dim bModification As Boolean
  Dim bLockedGroupe As Boolean
 
 
  Select Case eMode
 'En mode ajout, on montre TxtGroupe,les boutons annuler et
 'enregistrer
 Case MODE_AJOUT:
  bTxtGroupe = True
  bAnnuler = True
  bEnregistrer = True
  bAffichage = True
  bModification = True
 
 Case MODE_MODIF:
  bTxtGroupe = True
 
 If txtGroupe.Text = S_GROUPE_DEFAUT Then
 bLockedGroupe = True
 End If
 
 bAnnuler = True
 bEnregistrer = True
 bAffichage = True
 bModification = True
 
 'En mode inactif, on montre cmbGroupe, les boutons Ajouter,
 'Modifier, Supprimer et Quitter
 Case MODE_INACTIF:
 bCmbGroupe = True
 bQuitter = True
 bAjouter = True
 bModifier = True
 bSupprimer = True
1  End Select
 
txtGroupe.Visible = bTxtGroupe
 txtGroupe.Locked = bLockedGroupe
cmbGroupe.Visible = bCmbGroupe
 cmdAnnuler.Visible = bAnnuler
cmdEnregistrer.Visible = bEnregistrer
 Cmdfermer.Visible = bQuitter
1  Cmdajouter.Visible = bAjouter
 cmdModifier.Visible = bModifier
 cmdsupprimer.Visible = bSupprimer
fraAffichage.Enabled = bAffichage
fraModification.Enabled = bModification

Exit Sub

Oups:

wOups "frmGroupes", "MontrerControles", Err, Err.number, Err.Description
End Sub

Private Sub AfficherGroupe()

 On Error GoTo Oups

 'Affiche le groupe selon la sélection dans le combo
 Dim rstGroupes As ADODB.Recordset
 
 Set rstGroupes = New ADODB.Recordset
 
 Call rstGroupes.Open("SELECT * FROM GrbGroupes WHERE IDGroupe = " & m_iNoGroupe, g_connData, adOpenDynamic, adLockOptimistic)
 
 chkClients.Value = Abs(CInt(rstGroupes.Fields("Clients").Value))
 chkFournisseurs.Value = Abs(CInt(rstGroupes.Fields("Fournisseurs").Value))
 chkContacts.Value = Abs(CInt(rstGroupes.Fields("Contacts").Value))
 chkContactsVendeurs.Value = Abs(CInt(rstGroupes.Fields("ContactsVendeurs").Value))
 chkRapports.Value = Abs(CInt(rstGroupes.Fields("Rapport").Value))
 chkCatalogueMec.Value = Abs(CInt(rstGroupes.Fields("CatalogueMec").Value))
 chkCatalogueElec.Value = Abs(CInt(rstGroupes.Fields("CatalogueElec").Value))
  chkEmployes.Value = Abs(CInt(rstGroupes.Fields("Employes").Value))
  chkCedule.Value = Abs(CInt(rstGroupes.Fields("Cedule").Value))
  chkConfiguration.Value = Abs(CInt(rstGroupes.Fields("Configuration").Value))
  chkPunch.Value = Abs(CInt(rstGroupes.Fields("Punch").Value))
  chkOutils.Value = Abs(CInt(rstGroupes.Fields("Outils").Value))
  chkSoumissionMec.Value = Abs(CInt(rstGroupes.Fields("SoumissionMec").Value))
  chkProjetMec.Value = Abs(CInt(rstGroupes.Fields("ProjetMec").Value))
  chkSoumissionElec.Value = Abs(CInt(rstGroupes.Fields("SoumissionElec").Value))
10 chkProjetElec.Value = Abs(CInt(rstGroupes.Fields("ProjetElec").Value))
chkInventaireMec.Value = Abs(CInt(rstGroupes.Fields("InventaireMec").Value))
chkInventaireElec.Value = Abs(CInt(rstGroupes.Fields("InventaireElec").Value))
chkAchat.Value = Abs(CInt(rstGroupes.Fields("Achat").Value))
chkModificationFacturation.Value = Abs(CInt(rstGroupes.Fields("ModificationFacturation").Value))
chkModificationClients.Value = Abs(CInt(rstGroupes.Fields("ModificationClients").Value))
chkModificationFRS.Value = Abs(CInt(rstGroupes.Fields("ModificationFournisseurs").Value))
chkModificationContacts.Value = Abs(CInt(rstGroupes.Fields("ModificationContacts").Value))
chkModificationEmployes.Value = Abs(CInt(rstGroupes.Fields("ModificationEmployes").Value))
chkModificationGroupes.Value = Abs(CInt(rstGroupes.Fields("ModificationGroupes").Value))
chkModificationFeuillesTemps.Value = Abs(CInt(rstGroupes.Fields("ModificationFeuillesTemps").Value))
chkModificationOutils.Value = Abs(CInt(rstGroupes.Fields("ModificationOutils").Value))
1  chkModificationSoumissionMec.Value = Abs(CInt(rstGroupes.Fields("ModificationSoumissionsMec").Value))
chkModificationProjetMec.Value = Abs(CInt(rstGroupes.Fields("ModificationProjetsMec").Value))
 chkModificationSoumissionElec.Value = Abs(CInt(rstGroupes.Fields("ModificationSoumissionsElec").Value))
chkModificationProjetElec.Value = Abs(CInt(rstGroupes.Fields("ModificationProjetsElec").Value))
 chkModificationBonsCommandes.Value = Abs(CInt(rstGroupes.Fields("ModificationBonsCommandes").Value))
chkModificationCatalogueElec.Value = Abs(CInt(rstGroupes.Fields("ModificationCatalogueElec").Value))
 chkModificationCatalogueMec.Value = Abs(CInt(rstGroupes.Fields("ModificationCatalogueMec").Value))
1  chkModificationInventaireMec.Value = Abs(CInt(rstGroupes.Fields("ModificationInventaireMec").Value))
 chkModificationInventaireElec.Value = Abs(CInt(rstGroupes.Fields("ModificationInventaireElec").Value))
 chkModificationPunchEmployes.Value = Abs(CInt(rstGroupes.Fields("ModificationPunchEmployes").Value))
chkSupprimerProjets.Value = Abs(CInt(rstGroupes.Fields("SuppressionProjets").Value))
chkReception.Value = Abs(CInt(rstGroupes.Fields("ModificationReception").Value))
chkRetourMarchandise.Value = Abs(CInt(rstGroupes.Fields("ModificationRetourMarchandise").Value))
chkMailingList.Value = Abs(CInt(rstGroupes.Fields("ListeDistribution").Value))
chkPunchSemaineAnterieure.Value = Abs(CInt(rstGroupes.Fields("PunchSemaineAntérieure").Value))
chkVerrouillageTempsProjet.Value = Abs(CInt(rstGroupes.Fields("VerrouillageTempsProjet").Value))
chkDeverrouillageTempsProjet.Value = Abs(CInt(rstGroupes.Fields("DéverrouillageTempsProjet").Value))

Call rstGroupes.Close
Set rstGroupes = Nothing
 
Call AfficherUtilisateurs

2  Exit Sub

Oups:

wOups "frmGroupes", "AfficherGroupe", Err, Err.number, Err.Description
End Sub

Private Sub AfficherUtilisateurs()

 On Error GoTo Oups

 'Affiche les utilisateurs compris dans le groupe
 Dim rstEmployes As ADODB.Recordset
 
 'Vide la liste
 Call lstUser.Clear
 
 Set rstEmployes = New ADODB.Recordset
 
 Call rstEmployes.Open("SELECT * FROM Grbemployés WHERE Groupe = " & cmbGroupe.ItemData(cmbGroupe.ListIndex) & " AND Actif = true ORDER BY employe", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Tant que ce n'est pas la fin des enregistrements
 Do While Not rstEmployes.EOF
 'Ajout du nom de l'employé dans la liste
 Call lstUser.AddItem(rstEmployes.Fields("employe"))
 
 Call rstEmployes.MoveNext
 Loop
 
 Call rstEmployes.Close
 Set rstEmployes = Nothing

  Exit Sub

Oups:

  wOups "frmGroupes", "AfficherUtilisateurs", Err, Err.number, Err.Description
End Sub

Private Sub UncheckedAll()

 On Error GoTo Oups
 
 'Enlève tous les crochets des checkbox
 Dim objControl As Object
 
 For Each objControl In Me
 If TypeOf objControl Is CheckBox Then
 objControl.Value = vbUnchecked
 End If
 Next

 Exit Sub

Oups:

 wOups "frmGroupes", "UncheckedAll", Err, Err.number, Err.Description
End Sub
