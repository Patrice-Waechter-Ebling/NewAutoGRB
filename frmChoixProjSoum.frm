VERSION 5.00
Begin VB.Form frmChoixProjSoum 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Projets / Soumissions"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10365
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChoixProjSoum.frx":0000
   ScaleHeight     =   7500
   ScaleWidth      =   10365
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraReception 
      BackColor       =   &H00000000&
      Caption         =   "Réception"
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
      Height          =   1335
      Left            =   8040
      TabIndex        =   30
      Top             =   3480
      Width           =   2175
      Begin VB.CommandButton cmdReceptionElec 
         Caption         =   "Électrique"
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdReceptionMec 
         Caption         =   "Mécanique"
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.Frame fraAchat 
      BackColor       =   &H00000000&
      Caption         =   "Achat"
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
      Height          =   1335
      Left            =   8040
      TabIndex        =   45
      Top             =   5160
      Width           =   2175
      Begin VB.CommandButton cmdAchatMec 
         Caption         =   "Mécanique"
         Height          =   375
         Left            =   240
         TabIndex        =   47
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton cmdAchatElec 
         Caption         =   "Électrique"
         Height          =   375
         Left            =   240
         TabIndex        =   46
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame fraProjSoum 
      BackColor       =   &H00000000&
      Caption         =   "Projet / Soumission"
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
      Height          =   1335
      Left            =   8040
      TabIndex        =   12
      Top             =   1800
      Width           =   2175
      Begin VB.CommandButton cmdProjSoumElec 
         Caption         =   "Électrique"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdProjSoumMec 
         Caption         =   "Mécanique"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   8280
      TabIndex        =   51
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(Possibilité de 999)"
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
      Left            =   5880
      TabIndex        =   55
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label52 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
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
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label51 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "="
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
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label lblAnnee 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
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
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label49 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "="
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
      Left            =   600
      TabIndex        =   19
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label48 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "="
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
      Left            =   600
      TabIndex        =   17
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label47 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "="
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
      Left            =   600
      TabIndex        =   5
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label46 
      BackStyle       =   0  'Transparent
      Caption         =   "ZZ"
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
      Left            =   120
      TabIndex        =   20
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "YYY"
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
      Left            =   120
      TabIndex        =   16
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label44 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(Possibilité de 99)"
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
      Left            =   6000
      TabIndex        =   53
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label Label43 
      BackStyle       =   0  'Transparent
      Caption         =   "# Révision"
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
      Left            =   4080
      TabIndex        =   54
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label42 
      BackStyle       =   0  'Transparent
      Caption         =   "Dessin"
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
      Height          =   255
      Left            =   1560
      TabIndex        =   52
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label Label41 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(Possibilité de 99)"
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
      Left            =   6000
      TabIndex        =   50
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   "# Version"
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
      Left            =   4080
      TabIndex        =   49
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "Programmation"
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
      Height          =   255
      Left            =   1560
      TabIndex        =   48
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Label Label38 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(80 à 99)"
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
      Left            =   6840
      TabIndex        =   44
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "# Visite Non Facturée"
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
      Left            =   4080
      TabIndex        =   43
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Label Label36 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(Possibilité de 79)"
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
      Left            =   6000
      TabIndex        =   42
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "# Visite"
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
      Left            =   4080
      TabIndex        =   41
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "Technicien && Matériel"
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
      Height          =   255
      Left            =   1560
      TabIndex        =   40
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label Label33 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(Possibilité de 99)"
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
      Left            =   6000
      TabIndex        =   39
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "# Visite"
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
      Left            =   4080
      TabIndex        =   38
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Technicien"
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
      Height          =   255
      Left            =   1560
      TabIndex        =   37
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(80 à 99)"
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
      Left            =   6840
      TabIndex        =   35
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "# Extra Non Facturé"
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
      Left            =   4080
      TabIndex        =   36
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(60 à 79)"
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
      Left            =   6840
      TabIndex        =   33
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "# Extra Facturé"
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
      Left            =   4080
      TabIndex        =   34
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(50 à 59)"
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
      Left            =   6840
      TabIndex        =   28
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "# Mise en service"
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
      Left            =   4080
      TabIndex        =   29
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(01 à 49)"
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
      Left            =   6840
      TabIndex        =   27
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "# du Panneau"
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
      Left            =   4080
      TabIndex        =   26
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Fabrication"
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
      Height          =   255
      Left            =   1560
      TabIndex        =   25
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(Possibilité de 99)"
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
      Left            =   6000
      TabIndex        =   24
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "# Révision"
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
      Left            =   4080
      TabIndex        =   23
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Soumission"
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
      Height          =   255
      Left            =   1560
      TabIndex        =   22
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Ex."
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
      Left            =   960
      TabIndex        =   21
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "# Séquentiel de 3 chiffres"
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
      Height          =   255
      Left            =   960
      TabIndex        =   18
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "9 = Dessin"
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
      Left            =   2040
      TabIndex        =   15
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "7 = Technicien && Matériel"
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
      Left            =   2040
      TabIndex        =   11
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "5 = Technicien"
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
      Left            =   2040
      TabIndex        =   10
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "3 = Fabrication"
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
      Left            =   2040
      TabIndex        =   9
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "1 = Soumission"
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
      Left            =   2040
      TabIndex        =   8
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Année"
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
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblExemple 
      BackStyle       =   0  'Transparent
      Caption         =   "Exemple : 3XYYY-ZZ"
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
      Left            =   4440
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NUMÉROTATION DE DOSSIER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4095
   End
End
Attribute VB_Name = "frmChoixProjSoum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_sUserID   As String
Public m_iNoGroupe As Integer

Private Sub cmdAchatElec_Click()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer
15      Dim objAchat  As frmAchat
20      Dim bOuvert   As Boolean

25      Screen.MousePointer = vbHourglass

30      For iCompteur = 0 To Forms.count - 1
35        If Forms(iCompteur).Caption = "Achat électrique" Then
40          bOuvert = True

45          Exit For
50        End If
55      Next

60      If bOuvert = False Then
65        Set objAchat = New frmAchat

70        Call objAchat.Afficher(ELECTRIQUE)

75        Set g_objAchatElec = objAchat
80      Else
85        Forms(iCompteur).WindowState = vbNormal

90        Call Forms(iCompteur).ZOrder(0)

95        Call Unload(Me)
100     End If

105     Exit Sub

AfficherErreur:

110     woups "frmChoixProjSoum", "cmdAchatElec_Click", Err, Erl
End Sub

Private Sub cmdAchatMec_Click()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer
15      Dim objAchat  As frmAchat
20      Dim bOuvert   As Boolean

25      Screen.MousePointer = vbHourglass

30      For iCompteur = 0 To Forms.count - 1
35        If Forms(iCompteur).Caption = "Achat mécanique" Then
40          bOuvert = True

45          Exit For
50        End If
55      Next

60      If bOuvert = False Then
65        Set objAchat = New frmAchat

70        Call objAchat.Afficher(MECANIQUE)

75        Set g_objAchatMec = objAchat
80      Else
85        Forms(iCompteur).WindowState = vbNormal

90        Call Forms(iCompteur).ZOrder(0)

95        Call Unload(Me)
100     End If

105     Exit Sub

AfficherErreur:

110     woups "frmChoixProjSoum", "cmdAchatMec_Click", Err, Erl
End Sub

Private Sub Cmdfermer_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmChoixProjSoum", "cmdFermer_Click", Err, Erl
End Sub

Private Sub cmdProjSoumElec_Click()

5       On Error GoTo AfficherErreur

10      Screen.MousePointer = vbHourglass
  
15      Call OuvrirForm(FrmProjSoumElec, False)
  
20      Call Unload(Me)

25      Exit Sub

AfficherErreur:

30      woups "frmChoixProjSoum", "cmdProjSoumElec_Click", Err, Erl
End Sub

Private Sub cmdProjSoumMec_Click()

5       On Error GoTo AfficherErreur

10      Screen.MousePointer = vbHourglass
  
15      Call OuvrirForm(FrmProjSoumMec, False)
  
20      Call Unload(Me)

25      Exit Sub

AfficherErreur:

30      woups "frmChoixProjSoum", "cmdProjSoumMec_Click", Err, Erl
End Sub

Private Sub cmdReceptionElec_Click()

5       On Error GoTo AfficherErreur

10      Dim rstGroupe As ADODB.Recordset

        'Il faut afficher le login pour faire la réception
15      Call frmLogin.Afficher(Me)

        'Si bon password
20      If g_bBonPasswd = True Then
25        g_bBonPasswd = False

30        Set rstGroupe = New ADODB.Recordset

35        Call rstGroupe.Open("SELECT ModificationReception FROM GRB_Groupes WHERE IDGroupe = " & m_iNoGroupe, g_connData, adOpenDynamic, adLockOptimistic)
   
40        If rstGroupe.Fields("ModificationReception") = True Then

            'Ouverture des réceptions
45          Call FrmReceptionElec.Afficher(m_sUserID)

50          Call rstGroupe.Close
55          Set rstGroupe = Nothing

60          Call Unload(Me)
65        Else
70          Call MsgBox("Accès refusé!", vbOKOnly, "Erreur")
            
75          Call rstGroupe.Close
80          Set rstGroupe = Nothing
85        End If
90      End If

95      Exit Sub

AfficherErreur:

100     woups "frmChoixProjSoum", "cmdReceptionElec_Click", Err, Erl
End Sub

Private Sub cmdReceptionMec_Click()

5       On Error GoTo AfficherErreur

10      Dim rstGroupe As ADODB.Recordset

        'Il faut afficher le login pour faire la réception
15      Call frmLogin.Afficher(Me)

        'Si bon password
20      If g_bBonPasswd = True Then
25        g_bBonPasswd = False

30        Set rstGroupe = New ADODB.Recordset

35        Call rstGroupe.Open("SELECT ModificationReception FROM GRB_Groupes WHERE IDGroupe = " & m_iNoGroupe, g_connData, adOpenDynamic, adLockOptimistic)

40        If rstGroupe.Fields("ModificationReception") = True Then
            'Ouverture des réceptions
45          Call FrmReceptionMec.Afficher(m_sUserID)

50          Call rstGroupe.Close
55          Set rstGroupe = Nothing

60          Call Unload(Me)
65        Else
70          Call MsgBox("Accès refusé!", vbOKOnly, "Erreur")

75          Call rstGroupe.Close
80          Set rstGroupe = Nothing
85        End If
90      End If

95      Exit Sub

AfficherErreur:

100     woups "frmChoixProjSoum", "cmdReceptionMec_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      lblExemple.Caption = "Exemple : " & Right$(Year(Date), 1) & "XYYY-ZZ"

15      lblAnnee.Caption = Right$(Year(Date), 1)

20      Call ActiverBoutonsGroupe

25      Screen.MousePointer = vbDefault

30      Exit Sub

AfficherErreur:

35      woups "frmChoixProjSoum", "Form_Load", Err, Erl
End Sub

Private Sub ActiverBoutonsGroupe()

5       On Error GoTo AfficherErreur
    
10      If g_bAffichageSoumissionsMec = True Or g_bAffichageProjetsMec = True Then
15        cmdProjSoumMec.Enabled = True
20      Else
25        cmdProjSoumMec.Enabled = False
30      End If
    
35      If g_bAffichageSoumissionsElec = True Or g_bAffichageProjetsElec = True Then
40        cmdProjSoumElec.Enabled = True
45      Else
50        cmdProjSoumElec.Enabled = False
55      End If
    
60      If g_bAffichageAchats = True Then
65        cmdAchatElec.Enabled = True
70        cmdAchatMec.Enabled = True
75      Else
80        cmdAchatElec.Enabled = False
85        cmdAchatMec.Enabled = False
90     End If

95      If g_bModificationReception = True Then
100       cmdReceptionElec.Enabled = True
105       cmdReceptionMec.Enabled = True
110     Else
115       cmdReceptionElec.Enabled = False
120       cmdReceptionMec.Enabled = False
125     End If

130     Exit Sub

AfficherErreur:

135     woups "frmChoixProjSoum", "ActiverBoutonsGroupe", Err, Erl
End Sub

