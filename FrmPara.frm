VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPara 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuration"
   ClientHeight    =   8955
   ClientLeft      =   3240
   ClientTop       =   1710
   ClientWidth     =   8565
   ControlBox      =   0   'False
   Icon            =   "FrmPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConfig 
      Appearance      =   0  'Flat
      Caption         =   "."
      Height          =   195
      Left            =   0
      TabIndex        =   109
      TabStop         =   0   'False
      Top             =   8760
      Width           =   135
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   7095
      Index           =   4
      Left            =   360
      TabIndex        =   57
      Top             =   720
      Visible         =   0   'False
      Width           =   7695
      Begin VB.TextBox txtDixMoins 
         Height          =   285
         Left            =   2520
         TabIndex        =   59
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtDixQuinze 
         Height          =   285
         Left            =   2520
         TabIndex        =   61
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtQuinzeVingt 
         Height          =   285
         Left            =   2520
         TabIndex        =   63
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtVingtVingtCinq 
         Height          =   285
         Left            =   2520
         TabIndex        =   65
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtVingtCinqCinquante 
         Height          =   285
         Left            =   2520
         TabIndex        =   67
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtCinquanteCent 
         Height          =   285
         Left            =   2520
         TabIndex        =   69
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtCentPlus 
         Height          =   285
         Left            =   2520
         TabIndex        =   71
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "10,000$ et -"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   58
         Top             =   480
         Width           =   1020
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "15,000$ à 20,000$"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   62
         Top             =   1200
         Width           =   1545
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "20,000$ à 25,000$"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   64
         Top             =   1560
         Width           =   1545
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "25,000$ à 50,000$"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   66
         Top             =   1920
         Width           =   1545
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "50,000$ à 100,000$ "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   68
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "100,000$ et +"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   70
         Top             =   2640
         Width           =   1080
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "10,000$ à 15,000$ "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   60
         Top             =   840
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3972
      Index           =   3
      Left            =   360
      TabIndex        =   48
      Top             =   720
      Visible         =   0   'False
      Width           =   5532
      Begin VB.TextBox txtLeGrand 
         Height          =   285
         Left            =   1320
         TabIndex        =   50
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtLamine 
         Height          =   285
         Left            =   1320
         TabIndex        =   52
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtThermo 
         Height          =   285
         Left            =   1320
         TabIndex        =   54
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txt4em 
         Height          =   285
         Left            =   1320
         TabIndex        =   56
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "LeGrand"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   480
         TabIndex        =   49
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Laminé"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   480
         TabIndex        =   51
         Top             =   960
         Width           =   612
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Thermo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   480
         TabIndex        =   53
         Top             =   1320
         Width           =   648
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "4em"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   480
         TabIndex        =   55
         Top             =   1680
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   7335
      Index           =   2
      Left            =   240
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   8055
      Begin VB.TextBox txtShippingMec 
         Height          =   285
         Left            =   7200
         TabIndex        =   106
         Top             =   3960
         Width           =   735
      End
      Begin VB.TextBox txtShippingElec 
         Height          =   285
         Left            =   2520
         TabIndex        =   104
         Top             =   4680
         Width           =   735
      End
      Begin VB.TextBox txtRepas 
         Height          =   285
         Left            =   7080
         TabIndex        =   102
         Top             =   5520
         Width           =   852
      End
      Begin VB.TextBox txtUniteMobile 
         Height          =   285
         Left            =   7080
         TabIndex        =   99
         Top             =   6960
         Width           =   852
      End
      Begin VB.TextBox txtStandard 
         Height          =   285
         Left            =   7080
         TabIndex        =   97
         Top             =   6600
         Width           =   852
      End
      Begin VB.TextBox txtHebergement2 
         Height          =   285
         Left            =   2640
         TabIndex        =   94
         Top             =   6960
         Width           =   852
      End
      Begin VB.TextBox txtHebergement1 
         Height          =   285
         Left            =   2640
         TabIndex        =   92
         Top             =   6600
         Width           =   852
      End
      Begin VB.TextBox txtFormationElec 
         Height          =   285
         Left            =   2520
         TabIndex        =   90
         Top             =   3960
         Width           =   735
      End
      Begin VB.TextBox txtGestionProjetsElec 
         Height          =   285
         Left            =   2520
         TabIndex        =   88
         Top             =   4320
         Width           =   735
      End
      Begin VB.TextBox txtMiseEnService 
         Height          =   285
         Left            =   2520
         TabIndex        =   86
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox txtInstallationElec 
         Height          =   285
         Left            =   2520
         TabIndex        =   84
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox txtProgRobot 
         Height          =   285
         Left            =   2520
         TabIndex        =   81
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtProgAutomate 
         Height          =   285
         Left            =   2520
         TabIndex        =   79
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtGestionProjetsMec 
         Height          =   285
         Left            =   7200
         TabIndex        =   75
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox txtInstallationMec 
         Height          =   285
         Left            =   7200
         TabIndex        =   45
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox txtFormationMec 
         Height          =   285
         Left            =   7200
         TabIndex        =   41
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox txtDessinMec 
         Height          =   285
         Left            =   7200
         TabIndex        =   39
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtTestMec 
         Height          =   285
         Left            =   7200
         TabIndex        =   37
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtPeinture 
         Height          =   285
         Left            =   7200
         TabIndex        =   33
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtAssemblageMec 
         Height          =   285
         Left            =   7200
         TabIndex        =   29
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtSoudure 
         Height          =   285
         Left            =   7200
         TabIndex        =   25
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtCoupe 
         Height          =   285
         Left            =   7200
         TabIndex        =   21
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtMachinage 
         Height          =   285
         Left            =   7200
         TabIndex        =   17
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtProgVision 
         Height          =   285
         Left            =   2520
         TabIndex        =   35
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtTauxEspagnol 
         Height          =   285
         Left            =   2640
         TabIndex        =   47
         Top             =   5880
         Width           =   852
      End
      Begin VB.TextBox txtTauxAmericain 
         Height          =   285
         Left            =   2640
         TabIndex        =   43
         Top             =   5520
         Width           =   852
      End
      Begin VB.TextBox txtDessinElec 
         Height          =   285
         Left            =   2520
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtProgInterface 
         Height          =   285
         Left            =   2520
         TabIndex        =   19
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtAssemblageElec 
         Height          =   285
         Left            =   2520
         TabIndex        =   23
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtFabrication 
         Height          =   285
         Left            =   2520
         TabIndex        =   27
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtTestElec 
         Height          =   285
         Left            =   2520
         TabIndex        =   31
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         Caption         =   "Shipping :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3840
         TabIndex        =   107
         Top             =   3960
         Width           =   870
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "Shipping :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   105
         Top             =   4680
         Width           =   870
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         Caption         =   "Repas pour 1 journée :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3960
         TabIndex        =   103
         Top             =   5520
         Width           =   2955
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         Caption         =   "Unité mobile :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3840
         TabIndex        =   101
         Top             =   6960
         Width           =   3105
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         Caption         =   "Véhicule standard :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3840
         TabIndex        =   100
         Top             =   6600
         Width           =   3105
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         Caption         =   "Prix / KM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3720
         TabIndex        =   98
         Top             =   6360
         Width           =   3225
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         Caption         =   "Chambre à 2 lits :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   96
         Top             =   6960
         Width           =   2385
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         Caption         =   "Chambre à 1 lit :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   95
         Top             =   6600
         Width           =   2385
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Prix de l'hébergement"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   93
         Top             =   6360
         Width           =   2385
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "Formation du personnel :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   91
         Top             =   3960
         Width           =   2115
      End
      Begin VB.Line Line1 
         X1              =   3600
         X2              =   3600
         Y1              =   0
         Y2              =   5040
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "Gestion des projets :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   89
         Top             =   4320
         Width           =   1770
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "Mise en service :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   87
         Top             =   3600
         Width           =   1470
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Installation :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   85
         Top             =   3240
         Width           =   1065
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Général"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   83
         Top             =   5160
         Width           =   7815
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "Programmation de robot :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   82
         Top             =   2160
         Width           =   2145
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Programmation d'automate :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   80
         Top             =   1800
         Width           =   2370
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Taux mécaniques"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3840
         TabIndex        =   78
         Top             =   0
         Width           =   4035
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Taux électriques"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   77
         Top             =   0
         Width           =   3195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gestion des projets :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3885
         TabIndex        =   76
         Top             =   3600
         Width           =   1770
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Installation :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3885
         TabIndex        =   44
         Top             =   2880
         Width           =   1065
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Formation du personnel :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3885
         TabIndex        =   40
         Top             =   3240
         Width           =   2115
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conception et dessins :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3885
         TabIndex        =   38
         Top             =   360
         Width           =   2010
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tests Finaux :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3885
         TabIndex        =   36
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Peinture et finition :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3885
         TabIndex        =   32
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assemblage des systèmes :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3885
         TabIndex        =   28
         Top             =   1800
         Width           =   2325
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coupe, soudure et meulage :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3885
         TabIndex        =   24
         Top             =   1440
         Width           =   2460
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coupe et préparation (sauf soudage) :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3885
         TabIndex        =   20
         Top             =   720
         Width           =   3240
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Machinage :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3885
         TabIndex        =   16
         Top             =   1080
         Width           =   1065
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Programmation de vision :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   2520
         Width           =   2205
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Taux de change espagnol :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   5880
         Width           =   2340
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Taux de change américain :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   5520
         Width           =   2385
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Dessin :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Programmation d'interface :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   2340
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Assemblage :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fabrication :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Test :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   2880
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7335
      Index           =   1
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   7815
      Begin VB.CommandButton cmdExportToOutlook 
         Caption         =   "Export dans Outlook"
         Height          =   615
         Left            =   480
         TabIndex        =   108
         Top             =   3360
         Width           =   3375
      End
      Begin VB.TextBox txtProfitMec 
         Height          =   285
         Left            =   3480
         TabIndex        =   4
         Top             =   480
         Width           =   852
      End
      Begin VB.TextBox txtCheminSEE4000 
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   2400
         Width           =   4095
      End
      Begin VB.TextBox txtIndice 
         Height          =   285
         Left            =   3480
         TabIndex        =   8
         Top             =   1200
         Width           =   852
      End
      Begin VB.TextBox txtCommission 
         Height          =   285
         Left            =   3480
         TabIndex        =   6
         Top             =   840
         Width           =   852
      End
      Begin VB.TextBox txtProfitElec 
         Height          =   285
         Left            =   3480
         TabIndex        =   2
         Top             =   120
         Width           =   852
      End
      Begin VB.TextBox txtImprevus 
         Height          =   285
         Left            =   3480
         TabIndex        =   10
         Top             =   1560
         Width           =   852
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Profit mécanique:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1785
         TabIndex        =   3
         Top             =   480
         Width           =   1500
      End
      Begin VB.Label Label30 
         Caption         =   "Chemin de la base de données de SEE4000:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   4215
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Indice de dessin (% du tmp dess):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   1200
         Width           =   2865
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Commission (% du totale de la soum):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   3150
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Profit électrique:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1860
         TabIndex        =   1
         Top             =   120
         Width           =   1425
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pourcentage d'imprévus:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   9
         Top             =   1560
         Width           =   2115
      End
   End
   Begin MSComctlLib.TabStrip tbsPara 
      Height          =   7935
      Left            =   120
      TabIndex        =   74
      Top             =   240
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   13996
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Générale"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Taux Horaire"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Marqueur"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "FloorStock"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdFermer 
      Caption         =   "&Fermer"
      Height          =   375
      Left            =   7200
      TabIndex        =   73
      Top             =   8400
      Width           =   1215
   End
   Begin VB.CommandButton cmdAppliquer 
      Caption         =   "&Appliquer"
      Height          =   375
      Left            =   6000
      TabIndex        =   72
      Top             =   8400
      Width           =   1095
   End
End
Attribute VB_Name = "FrmPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_iCurrFrame As Integer

Private Sub cmdAppliquer_Click()

5       On Error GoTo AfficherErreur
        
        'Enregistrement des paramètres dans la table GRB_Config
10      If VerifierChamps = True Then
          'Enregistrer les configuration
15        Call EnregistrerConfiguration
                  
          'Fermeture du form
20        Call Unload(Me)
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmPara", "cmdAppliquer_Click", Err, Erl
End Sub

Private Function VerifierChamps() As Boolean

5       On Error GoTo AfficherErreur

10      Dim objControl As Object
  
15      VerifierChamps = True
  
        'Si champs vide
20      For Each objControl In Me
25        If TypeOf objControl Is TextBox Then
30          If Trim$(objControl.Text) = vbNullString Then
35            Call MsgBox("Vous devez remplir tous les champs!", vbOKOnly, "Erreur")
        
40            VerifierChamps = False
        
45            Exit Function
50          Else
55           If objControl.Name <> "txtCheminSEE4000" Then
60              objControl.Text = Replace(objControl.Text, ".", ",")
          
65              If Not IsNumeric(objControl.Text) Then
70                Call MsgBox("Champs non numérique!", vbOKOnly, "Erreur")
        
75                VerifierChamps = False
        
80                Exit Function
85              End If
90            End If
95          End If
100       End If
105     Next
  
        'Profit électrique
110     If txtProfitElec.Text < 1 Then
115       Call MsgBox("Le pourcentage de profit électrique doit être plus grand que 1 !", vbOKOnly, "Erreur")
      
120       VerifierChamps = False
      
125       Exit Function
130     End If
   
        'Profit mécanique
135     If txtProfitMec.Text < 1 Then
140       Call MsgBox("Le pourcentage de profit mécanique doit être plus grand que 1 !", vbOKOnly, "Erreur")

145       VerifierChamps = False

150       Exit Function
155     End If
   
160     If txtCommission.Text > 1 Then
165       Call MsgBox("Le pourcentage de commission doit être plus petit que 1 !", vbOKOnly, "Erreur")
    
170       VerifierChamps = False
    
175       Exit Function
180     End If
  
185     If txtImprevus.Text > 1 Then
190       Call MsgBox("Le pourcentage d'imprévus doit être plus petit que 1 !", vbOKOnly, "Erreur")
    
195       VerifierChamps = False
    
200       Exit Function
205     End If

210     Exit Function

AfficherErreur:

215     woups "frmPara", "VerifierChamps", Err, Erl
End Function

Private Sub EnregistrerConfiguration()

5       On Error GoTo AfficherErreur

        'Enregistrement des configurations
10      Dim rstPara As ADODB.Recordset
  
        'Initialisation des variables
15      Call InitialiserVariablesConfiguration
  
20      Set rstPara = New ADODB.Recordset
  
        'Enregistrement dans la BD
25      Call rstPara.Open("SELECT * FROM GRB_Config", g_connData, adOpenDynamic, adLockOptimistic)
    
30      rstPara.Fields("ProfitElec") = txtProfitElec.Text
35      rstPara.Fields("ProfitMec") = txtProfitMec.Text
40      rstPara.Fields("Commission") = txtCommission.Text
45      rstPara.Fields("Imprévus") = txtImprevus.Text
50      rstPara.Fields("IndiceDessin") = txtIndice.Text
55      rstPara.Fields("TauxAmericain") = txtTauxAmericain.Text
60      rstPara.Fields("TauxEspagnol") = txtTauxEspagnol.Text
65      rstPara.Fields("TauxDessinElec") = txtDessinElec.Text
70      rstPara.Fields("TauxProgInterface") = txtProgInterface.Text
75      rstPara.Fields("TauxProgAutomate") = txtProgAutomate.Text
80      rstPara.Fields("TauxProgRobot") = txtProgRobot.Text
85      rstPara.Fields("TauxVision") = txtProgVision.Text
90      rstPara.Fields("TauxAssemblageElec") = txtAssemblageElec.Text
95      rstPara.Fields("TauxFabrication") = txtFabrication.Text
100     rstPara.Fields("TauxTestElec") = txtTestElec.Text
105     rstPara.Fields("TauxGestionProjetsElec") = txtGestionProjetsElec.Text
110     rstPara.Fields("TauxInstallationElec") = txtInstallationElec.Text
115     rstPara.Fields("TauxMiseService") = txtMiseEnService.Text
120     rstPara.Fields("TauxFormationElec") = txtFormationElec.Text
125     rstPara.Fields("TauxShippingElec") = txtShippingElec.Text
130     rstPara.Fields("TauxMachinage") = txtMachinage.Text
135     rstPara.Fields("TauxCoupe") = txtCoupe.Text
140     rstPara.Fields("TauxSoudure") = txtSoudure.Text
145     rstPara.Fields("TauxAssemblageMec") = txtAssemblageMec.Text
150     rstPara.Fields("TauxPeinture") = txtPeinture.Text
155     rstPara.Fields("TauxTestMec") = txtTestMec.Text
160     rstPara.Fields("TauxGestionProjetsMec") = txtGestionProjetsMec
165     rstPara.Fields("TauxDessinMec") = txtDessinMec.Text
170     rstPara.Fields("TauxFormationMec") = txtFormationMec.Text
175     rstPara.Fields("TauxInstallationMec") = txtInstallationMec.Text
180     rstPara.Fields("TauxShippingMec") = txtShippingMec.Text
185     rstPara.Fields("LeGrand") = txtLeGrand.Text
190     rstPara.Fields("Lamine") = txtLamine.Text
195     rstPara.Fields("Thermo") = txtThermo.Text
200     rstPara.Fields("4em") = txt4em.Text
205     rstPara.Fields("fsDixMoins") = txtDixMoins.Text
210     rstPara.Fields("fsDix") = txtDixQuinze.Text
215     rstPara.Fields("fsQuinze") = txtQuinzeVingt.Text
220     rstPara.Fields("fsVingt") = txtVingtVingtCinq.Text
225     rstPara.Fields("fsVingtCinq") = txtVingtCinqCinquante.Text
230     rstPara.Fields("fsCinquante") = txtCinquanteCent.Text
235     rstPara.Fields("fsCent") = txtCentPlus.Text
240     rstPara.Fields("CheminSee4000") = txtCheminSEE4000.Text
245     rstPara.Fields("Hebergement1") = txtHebergement1.Text
250     rstPara.Fields("Hebergement2") = txtHebergement2.Text
255     rstPara.Fields("Repas") = txtRepas.Text
260     rstPara.Fields("Standard") = txtStandard.Text
265     rstPara.Fields("UniteMobile") = txtUniteMobile.Text
  
270     Call rstPara.Update
   
275     Call rstPara.Close
280     Set rstPara = Nothing

285     Exit Sub

AfficherErreur:

290     woups "frmPara", "EnregistrerConfiguration", Err, Erl
End Sub


Private Sub cmdConfig_Click()

    Dim sVersion As String
    Dim rstPara As ADODB.Recordset
    
    sVersion = InputBox("Entrer le mot de passe.", "Version")
    If sVersion = "gaetan" Then
        Set rstPara = New ADODB.Recordset
        sVersion = ""
        Call rstPara.Open("SELECT DerniereVersion FROM GRB_Config", g_connData, adOpenDynamic, adLockOptimistic)
        sVersion = rstPara("DerniereVersion")
        sVersion = InputBox("Entrer le numéro de version.", "Version", sVersion)
        If Not sVersion = "" Then
            rstPara("DerniereVersion") = sVersion
            rstPara.Update
        End If
        rstPara.Close
        Set rstPara = Nothing
    End If
    
End Sub

Private Sub cmdExportToOutlook_Click()
    Call OuvrirForm(frmExportToOutLook, True)
End Sub

Private Sub Cmdfermer_Click()

5       On Error GoTo AfficherErreur

        'Fermer la fenêtre
10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmPara", "cmdFermer_Click", Err, Erl
End Sub

Private Sub RemplirValue()

5       On Error GoTo AfficherErreur

        'On remplir les champs à l'aide de la table GRB_Config
10      Dim rstPara As ADODB.Recordset
    
15      Set rstPara = New ADODB.Recordset
    
20      Call rstPara.Open("SELECT * FROM GRB_Config", g_connData, adOpenDynamic, adLockOptimistic)
    
25      If Not IsNull(rstPara.Fields("ProfitElec")) Then
30        txtProfitElec.Text = rstPara.Fields("ProfitElec")
35      End If
  
40      If Not IsNull(rstPara.Fields("ProfitMec")) Then
45        txtProfitMec.Text = rstPara.Fields("ProfitMec")
50      End If

55      If Not IsNull(rstPara.Fields("Commission")) Then
60        txtCommission.Text = rstPara.Fields("Commission")
65      End If

70      If Not IsNull(rstPara.Fields("Imprévus")) Then
75        txtImprevus.Text = rstPara.Fields("Imprévus")
80      End If

85      If Not IsNull(rstPara.Fields("IndiceDessin")) Then
90        txtIndice.Text = rstPara.Fields("IndiceDessin")
95      End If

100     If Not IsNull(rstPara.Fields("TauxAmericain")) Then
105        txtTauxAmericain.Text = rstPara.Fields("TauxAmericain")
110     End If

115     If Not IsNull(rstPara.Fields("TauxEspagnol")) Then
120       txtTauxEspagnol.Text = rstPara.Fields("TauxEspagnol")
125     End If

130     If Not IsNull(rstPara.Fields("TauxDessinElec")) Then
135       txtDessinElec.Text = rstPara.Fields("TauxDessinElec")
140     End If

145     If Not IsNull(rstPara.Fields("TauxProgInterface")) Then
150       txtProgInterface.Text = rstPara.Fields("TauxProgInterface")
155     End If

160     If Not IsNull(rstPara.Fields("TauxProgAutomate")) Then
165       txtProgAutomate.Text = rstPara.Fields("TauxProgAutomate")
170     End If

175     If Not IsNull(rstPara.Fields("TauxProgRobot")) Then
180       txtProgRobot.Text = rstPara.Fields("TauxProgRobot")
185     End If

190     If Not IsNull(rstPara.Fields("TauxVision")) Then
195       txtProgVision.Text = rstPara.Fields("TauxVision")
200     End If

205     If Not IsNull(rstPara.Fields("TauxAssemblageElec")) Then
210       txtAssemblageElec.Text = rstPara.Fields("TauxAssemblageElec")
215     End If

220     If Not IsNull(rstPara.Fields("TauxFabrication")) Then
225       txtFabrication.Text = rstPara.Fields("TauxFabrication")
230     End If

235     If Not IsNull(rstPara.Fields("TauxTestElec")) Then
240       txtTestElec.Text = rstPara.Fields("TauxTestElec")
245     End If

250     If Not IsNull(rstPara.Fields("TauxGestionProjetsElec")) Then
255       txtGestionProjetsElec.Text = rstPara.Fields("TauxGestionProjetsElec")
260     End If

265     If Not IsNull(rstPara.Fields("TauxInstallationElec")) Then
270       txtInstallationElec.Text = rstPara.Fields("TauxInstallationElec")
275     End If

280     If Not IsNull(rstPara.Fields("TauxMiseService")) Then
285       txtMiseEnService.Text = rstPara.Fields("TauxMiseService")
290     End If

295     If Not IsNull(rstPara.Fields("TauxFormationElec")) Then
300       txtFormationElec.Text = rstPara.Fields("TauxFormationElec")
305     End If

310     If Not IsNull(rstPara.Fields("TauxShippingElec")) Then
315       txtShippingElec.Text = rstPara.Fields("TauxShippingElec")
320     End If

325     If Not IsNull(rstPara.Fields("TauxMachinage")) Then
330       txtMachinage.Text = rstPara.Fields("TauxMachinage")
335     End If

340     If Not IsNull(rstPara.Fields("TauxCoupe")) Then
345       txtCoupe.Text = rstPara.Fields("TauxCoupe")
350     End If

355     If Not IsNull(rstPara.Fields("TauxSoudure")) Then
360       txtSoudure.Text = rstPara.Fields("TauxSoudure")
365     End If

370     If Not IsNull(rstPara.Fields("TauxAssemblageMec")) Then
375       txtAssemblageMec.Text = rstPara.Fields("TauxAssemblageMec")
380     End If

385     If Not IsNull(rstPara.Fields("TauxPeinture")) Then
390       txtPeinture.Text = rstPara.Fields("TauxPeinture")
395     End If

400     If Not IsNull(rstPara.Fields("TauxTestMec")) Then
405       txtTestMec.Text = rstPara.Fields("TauxTestMec")
410     End If

415     If Not IsNull(rstPara.Fields("TauxGestionProjetsMec")) Then
420       txtGestionProjetsMec = rstPara.Fields("TauxGestionProjetsMec")
425     End If

430     If Not IsNull(rstPara.Fields("TauxDessinMec")) Then
435       txtDessinMec.Text = rstPara.Fields("TauxDessinMec")
440     End If

445     If Not IsNull(rstPara.Fields("TauxFormationMec")) Then
450       txtFormationMec.Text = rstPara.Fields("TauxFormationMec")
455     End If

460     If Not IsNull(rstPara.Fields("TauxInstallationMec")) Then
465       txtInstallationMec.Text = rstPara.Fields("TauxInstallationMec")
470     End If

475     If Not IsNull(rstPara.Fields("TauxShippingMec")) Then
480       txtShippingMec.Text = rstPara.Fields("TauxShippingMec")
485     End If

490     If Not IsNull(rstPara.Fields("LeGrand")) Then
495       txtLeGrand.Text = rstPara.Fields("LeGrand")
500     End If

505     If Not IsNull(rstPara.Fields("Lamine")) Then
510       txtLamine.Text = rstPara.Fields("Lamine")
515     End If

520     If Not IsNull(rstPara.Fields("Thermo")) Then
525       txtThermo.Text = rstPara.Fields("Thermo")
530     End If

535     If Not IsNull(rstPara.Fields("4em")) Then
540       txt4em.Text = rstPara.Fields("4em")
545     End If

550     If Not IsNull(rstPara.Fields("fsDixMoins")) Then
555       txtDixMoins.Text = rstPara.Fields("fsDixMoins")
560     End If

565     If Not IsNull(rstPara.Fields("fsDix")) Then
570       txtDixQuinze.Text = rstPara.Fields("fsDix")
575     End If

580     If Not IsNull(rstPara.Fields("fsQuinze")) Then
585       txtQuinzeVingt.Text = rstPara.Fields("fsQuinze")
590     End If

595     If Not IsNull(rstPara.Fields("fsVingt")) Then
600       txtVingtVingtCinq.Text = rstPara.Fields("fsVingt")
605     End If

610     If Not IsNull(rstPara.Fields("fsVingtCinq")) Then
615       txtVingtCinqCinquante.Text = rstPara.Fields("fsVingtCinq")
620     End If

625     If Not IsNull(rstPara.Fields("fsCinquante")) Then
630       txtCinquanteCent.Text = rstPara.Fields("fsCinquante")
635     End If

640     If Not IsNull(rstPara.Fields("fsCent")) Then
645       txtCentPlus.Text = rstPara.Fields("fsCent")
650     End If

655     If Not IsNull(rstPara.Fields("CheminSee4000")) Then
660       txtCheminSEE4000.Text = rstPara.Fields("CheminSee4000")
665     End If

670     If Not IsNull(rstPara.Fields("Hebergement1")) Then
675       txtHebergement1.Text = rstPara.Fields("Hebergement1")
680     End If

685     If Not IsNull(rstPara.Fields("Hebergement2")) Then
690       txtHebergement2.Text = rstPara.Fields("Hebergement2")
695     End If

700     If Not IsNull(rstPara.Fields("Repas")) Then
705       txtRepas.Text = rstPara.Fields("Repas")
710     End If

715     If Not IsNull(rstPara.Fields("Standard")) Then
720       txtStandard.Text = rstPara.Fields("Standard")
725     End If

730     If Not IsNull(rstPara.Fields("UniteMobile")) Then
735       txtUniteMobile.Text = rstPara.Fields("UniteMobile")
740     End If

745     Call rstPara.Close
750     Set rstPara = Nothing
    
755     Exit Sub

AfficherErreur:

760     woups "frmPara", "RemplirValue", Err, Erl
End Sub


Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      Call RemplirValue
  
15      Frame1(1).Visible = True
  
20      m_iCurrFrame = 1
  
25      Call TbsPara_Click

30      Screen.MousePointer = vbDefault

35      Exit Sub

AfficherErreur:

40      woups "frmPara", "Form_Load", Err, Erl
End Sub

Private Sub TbsPara_Click()

5       On Error GoTo AfficherErreur

        'si on est deja sur le l'onglet voulu on ne fait rien
10      If tbsPara.SelectedItem.Index <> m_iCurrFrame Then
15        Frame1(tbsPara.SelectedItem.Index).Visible = True
20        Frame1(m_iCurrFrame).Visible = False
25        m_iCurrFrame = tbsPara.SelectedItem.Index
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmPara", "TbsPara_Click", Err, Erl
End Sub

