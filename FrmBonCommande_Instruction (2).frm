VERSION 5.00
Begin VB.Form FrmBonCommande_Instruction 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bon Commande - Configuration Instruction"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   5925
   Begin VB.Frame fraLabel 
      BackColor       =   &H00000000&
      Caption         =   "Étiquette"
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   5655
      Begin VB.TextBox txtPays 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox txtAdresse 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox txtCompagnie 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label lblPays 
         BackStyle       =   0  'Transparent
         Caption         =   "Pays"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblAdresse 
         BackStyle       =   0  'Transparent
         Caption         =   "Adresse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblcompagnie 
         BackStyle       =   0  'Transparent
         Caption         =   "Compagnie"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.TextBox txtAssistance 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Top             =   3240
      Width           =   4095
   End
   Begin VB.TextBox txtEtat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton CmdEnr 
      Caption         =   "&Enregistrer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   11
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton CmdFerme 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Fermer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   12
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label lblassistance 
      BackStyle       =   0  'Transparent
      Caption         =   "Assistance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblEtat 
      BackStyle       =   0  'Transparent
      Caption         =   "État"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   1815
   End
End
Attribute VB_Name = "FrmBonCommande_Instruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdEnr_Click()

 On Error GoTo Oups

 ''''''''''''''''''''''''''''
 'Enregistrement d'une modif
 ''''''''''''''''''''''''''''
 Dim rstConfig As ADODB.Recordset

 If txtCompagnie.Text <> vbNullString And txtAdresse.Text <> vbNullString And txtEtat.Text <> vbNullString And txtAssistance.Text <> vbNullString Then
 Set rstConfig = New ADODB.Recordset

 Call rstConfig.Open("SELECT * FROM GrbConfig", g_connData, adOpenDynamic, adLockOptimistic)
 
 'enreg les donnees
 rstConfig.Fields("parcel_label_line1") = txtCompagnie.Text
 rstConfig.Fields("parcel_label_line2") = txtAdresse.Text
 rstConfig.Fields("parcel_label_line3") = txtPays.Text
 rstConfig.Fields("parcelassist") = txtAssistance.Text
 rstConfig.Fields("parceletat") = txtEtat.Text
 
 Call rstConfig.Update
 
 'ferme table
  Call rstConfig.Close
  Set rstConfig = Nothing
  Else
  Call MsgBox("Champs vides!", , "Erreur")
  End If

  Exit Sub

Oups:

  wOups "frmBonCommande_Instruction", "CmdEnr_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdFerme_Click()

 On Error GoTo Oups
 
 'quitte le form
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmBonCommande_Instruction", "CmdFerme_Click", Err, Err.number, Err.Description
End Sub

Private Sub AfficherDonnees()

 On Error GoTo Oups
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 'Affiche les données pour le rapport bon commande instruction
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Dim rstConfig As ADODB.Recordset

 Set rstConfig = New ADODB.Recordset

 Call rstConfig.Open("SELECT * FROM GrbConfig", g_connData, adOpenDynamic, adLockOptimistic)
 
 'met les donnees dans les controls
 txtCompagnie.Text = rstConfig.Fields("parcel_label_line1")
 txtAdresse.Text = rstConfig.Fields("parcel_label_line2")
 txtPays.Text = rstConfig.Fields("parcel_label_line3")
 txtAssistance.Text = rstConfig.Fields("parcelassist")
 txtEtat.Text = rstConfig.Fields("parceletat")
 
 'ferme table
 Call rstConfig.Close
 Set rstConfig = Nothing

  Exit Sub

Oups:

  wOups "frmBonCommande_Instruction", "AfficherDonnees", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups
 
 'affichage en mode visualisation
 Call AfficherDonnees

 Exit Sub

Oups:

 wOups "frmBonCommande_Instruction", "Form_Load", Err, Err.number, Err.Description
End Sub
