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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   8565
   Begin VB.CommandButton cmdConfig 
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
         Caption         =   "15,000$ � 20,000$"
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
         Caption         =   "20,000$ � 25,000$"
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
         Caption         =   "25,000$ � 50,000$"
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
         Caption         =   "50,000$ � 100,000$ "
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
         Caption         =   "10,000$ � 15,000$ "
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
         Caption         =   "Lamin�"
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
         Caption         =   "Repas pour 1 journ�e :"
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
         Caption         =   "Unit� mobile :"
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
         Caption         =   "V�hicule standard :"
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
         Caption         =   "Chambre � 2 lits :"
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
         Caption         =   "Chambre � 1 lit :"
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
         Caption         =   "Prix de l'h�bergement"
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
         Caption         =   "G�n�ral"
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
         Caption         =   "Taux m�caniques"
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
         Caption         =   "Taux �lectriques"
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
         Caption         =   "Assemblage des syst�mes :"
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
         Caption         =   "Coupe et pr�paration (sauf soudage) :"
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
         Caption         =   "Taux de change am�ricain :"
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
         Caption         =   "Profit m�canique:"
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
         Caption         =   "Chemin de la base de donn�es de SEE4000:"
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
         Caption         =   "Profit �lectrique:"
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
         Caption         =   "Pourcentage d'impr�vus:"
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
            Caption         =   "G�n�rale"
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

 On Error GoTo Oups
 
 'Enregistrement des param�tres dans la table GrbConfig
 If VerifierChamps = True Then
 'Enregistrer les configuration
 Call EnregistrerConfiguration
 
 'Fermeture du form
 Call Unload(Me)
 End If

 Exit Sub

Oups:

 wOups "frmPara", "cmdAppliquer_Click", Err, Err.number, Err.Description
End Sub

Private Function VerifierChamps() As Boolean

 On Error GoTo Oups

 Dim objControl As Object
 
 VerifierChamps = True
 
 'Si champs vide
 For Each objControl In Me
 If TypeOf objControl Is TextBox Then
 If Trim$(objControl.Text) = vbNullString Then
 Call MsgBox("Vous devez remplir tous les champs!", vbOKOnly, "Erreur")
 
 VerifierChamps = False
 
 Exit Function
 Else
 If objControl.Name <> "txtCheminSEE4000" Then
  objControl.Text = Replace(objControl.Text, ".", ",")
 
  If Not IsNumeric(objControl.Text) Then
  Call MsgBox("Champs non num�rique!", vbOKOnly, "Erreur")
 
  VerifierChamps = False
 
  Exit Function
  End If
  End If
  End If
End If
Next
 
 'Profit �lectrique
If txtProfitElec.Text < 1 Then
 Call MsgBox("Le pourcentage de profit �lectrique doit �tre plus grand que 1 !", vbOKOnly, "Erreur")
 
 VerifierChamps = False
 
 Exit Function
End If
 
 'Profit m�canique
If txtProfitMec.Text < 1 Then
 Call MsgBox("Le pourcentage de profit m�canique doit �tre plus grand que 1 !", vbOKOnly, "Erreur")

 VerifierChamps = False

 Exit Function
End If
 
1  If txtCommission.Text > 1 Then
 Call MsgBox("Le pourcentage de commission doit �tre plus petit que 1 !", vbOKOnly, "Erreur")
 
 VerifierChamps = False
 
 Exit Function
 End If
 
If txtImprevus.Text > 1 Then
 Call MsgBox("Le pourcentage d'impr�vus doit �tre plus petit que 1 !", vbOKOnly, "Erreur")
 
1  VerifierChamps = False
 
 Exit Function
 End If

Exit Function

Oups:

wOups "frmPara", "VerifierChamps", Err, Err.number, Err.Description
End Function

Private Sub EnregistrerConfiguration()

 On Error GoTo Oups

 'Enregistrement des configurations
 Dim rstPara As ADODB.Recordset
 
 'Initialisation des variables
 Call InitialiserVariablesConfiguration
 
 Set rstPara = New ADODB.Recordset
 
 'Enregistrement dans la BD
 Call rstPara.Open("SELECT * FROM GrbConfig", g_connData, adOpenDynamic, adLockOptimistic)
 
 rstPara.Fields("ProfitElec") = txtProfitElec.Text
 rstPara.Fields("ProfitMec") = txtProfitMec.Text
 rstPara.Fields("Commission") = txtCommission.Text
 rstPara.Fields("Impr�vus") = txtImprevus.Text
 rstPara.Fields("IndiceDessin") = txtIndice.Text
 rstPara.Fields("TauxAmericain") = txtTauxAmericain.Text
  rstPara.Fields("TauxEspagnol") = txtTauxEspagnol.Text
  rstPara.Fields("TauxDessinElec") = txtDessinElec.Text
  rstPara.Fields("TauxProgInterface") = txtProgInterface.Text
  rstPara.Fields("TauxProgAutomate") = txtProgAutomate.Text
  rstPara.Fields("TauxProgRobot") = txtProgRobot.Text
  rstPara.Fields("TauxVision") = txtProgVision.Text
  rstPara.Fields("TauxAssemblageElec") = txtAssemblageElec.Text
  rstPara.Fields("TauxFabrication") = txtFabrication.Text
10 rstPara.Fields("TauxTestElec") = txtTestElec.Text
rstPara.Fields("TauxGestionProjetsElec") = txtGestionProjetsElec.Text
rstPara.Fields("TauxInstallationElec") = txtInstallationElec.Text
rstPara.Fields("TauxMiseService") = txtMiseEnService.Text
rstPara.Fields("TauxFormationElec") = txtFormationElec.Text
rstPara.Fields("TauxShippingElec") = txtShippingElec.Text
rstPara.Fields("TauxMachinage") = txtMachinage.Text
rstPara.Fields("TauxCoupe") = txtCoupe.Text
rstPara.Fields("TauxSoudure") = txtSoudure.Text
rstPara.Fields("TauxAssemblageMec") = txtAssemblageMec.Text
rstPara.Fields("TauxPeinture") = txtPeinture.Text
rstPara.Fields("TauxTestMec") = txtTestMec.Text
1  rstPara.Fields("TauxGestionProjetsMec") = txtGestionProjetsMec
rstPara.Fields("TauxDessinMec") = txtDessinMec.Text
 rstPara.Fields("TauxFormationMec") = txtFormationMec.Text
rstPara.Fields("TauxInstallationMec") = txtInstallationMec.Text
 rstPara.Fields("TauxShippingMec") = txtShippingMec.Text
rstPara.Fields("LeGrand") = txtLeGrand.Text
 rstPara.Fields("Lamine") = txtLamine.Text
1  rstPara.Fields("Thermo") = txtThermo.Text
 rstPara.Fields("4em") = txt4em.Text
 rstPara.Fields("fsDixMoins") = txtDixMoins.Text
rstPara.Fields("fsDix") = txtDixQuinze.Text
rstPara.Fields("fsQuinze") = txtQuinzeVingt.Text
rstPara.Fields("fsVingt") = txtVingtVingtCinq.Text
rstPara.Fields("fsVingtCinq") = txtVingtCinqCinquante.Text
rstPara.Fields("fsCinquante") = txtCinquanteCent.Text
rstPara.Fields("fsCent") = txtCentPlus.Text
rstPara.Fields("CheminSee4000") = txtCheminSEE4000.Text
rstPara.Fields("Hebergement1") = txtHebergement1.Text
rstPara.Fields("Hebergement2") = txtHebergement2.Text
rstPara.Fields("Repas") = txtRepas.Text
2  rstPara.Fields("Standard") = txtStandard.Text
rstPara.Fields("UniteMobile") = txtUniteMobile.Text
 
2  Call rstPara.Update
 
Call rstPara.Close
2  Set rstPara = Nothing

Exit Sub

Oups:

2  wOups "frmPara", "EnregistrerConfiguration", Err, Err.number, Err.Description
End Sub


Private Sub cmdConfig_Click()

 Dim sVersion As String
 Dim rstPara As ADODB.Recordset
 
 sVersion = InputBox("Entrer le mot de passe.", "Version")
 If sVersion = "gaetan" Then
 Set rstPara = New ADODB.Recordset
 sVersion = ""
 Call rstPara.Open("SELECT DerniereVersion FROM GrbConfig", g_connData, adOpenDynamic, adLockOptimistic)
 sVersion = rstPara("DerniereVersion")
 sVersion = InputBox("Entrer le num�ro de version.", "Version", sVersion)
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

 On Error GoTo Oups

 'Fermer la fen�tre
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmPara", "cmdFermer_Click", Err, Err.number, Err.Description
End Sub

Private Sub RemplirValue()

 On Error GoTo Oups

 'On remplir les champs � l'aide de la table GrbConfig
 Dim rstPara As ADODB.Recordset
 
 Set rstPara = New ADODB.Recordset
 
 Call rstPara.Open("SELECT * FROM GrbConfig", g_connData, adOpenDynamic, adLockOptimistic)
 
 If Not IsNull(rstPara.Fields("ProfitElec")) Then
 txtProfitElec.Text = rstPara.Fields("ProfitElec")
 End If
 
 If Not IsNull(rstPara.Fields("ProfitMec")) Then
 txtProfitMec.Text = rstPara.Fields("ProfitMec")
 End If

 If Not IsNull(rstPara.Fields("Commission")) Then
  txtCommission.Text = rstPara.Fields("Commission")
  End If

  If Not IsNull(rstPara.Fields("Impr�vus")) Then
  txtImprevus.Text = rstPara.Fields("Impr�vus")
  End If

  If Not IsNull(rstPara.Fields("IndiceDessin")) Then
  txtIndice.Text = rstPara.Fields("IndiceDessin")
  End If

10 If Not IsNull(rstPara.Fields("TauxAmericain")) Then
1 txtTauxAmericain.Text = rstPara.Fields("TauxAmericain")
End If

If Not IsNull(rstPara.Fields("TauxEspagnol")) Then
 txtTauxEspagnol.Text = rstPara.Fields("TauxEspagnol")
End If

If Not IsNull(rstPara.Fields("TauxDessinElec")) Then
 txtDessinElec.Text = rstPara.Fields("TauxDessinElec")
End If

If Not IsNull(rstPara.Fields("TauxProgInterface")) Then
 txtProgInterface.Text = rstPara.Fields("TauxProgInterface")
End If

1  If Not IsNull(rstPara.Fields("TauxProgAutomate")) Then
 txtProgAutomate.Text = rstPara.Fields("TauxProgAutomate")
 End If

If Not IsNull(rstPara.Fields("TauxProgRobot")) Then
 txtProgRobot.Text = rstPara.Fields("TauxProgRobot")
End If

 If Not IsNull(rstPara.Fields("TauxVision")) Then
1  txtProgVision.Text = rstPara.Fields("TauxVision")
 End If

 If Not IsNull(rstPara.Fields("TauxAssemblageElec")) Then
 txtAssemblageElec.Text = rstPara.Fields("TauxAssemblageElec")
End If

If Not IsNull(rstPara.Fields("TauxFabrication")) Then
 txtFabrication.Text = rstPara.Fields("TauxFabrication")
End If

If Not IsNull(rstPara.Fields("TauxTestElec")) Then
 txtTestElec.Text = rstPara.Fields("TauxTestElec")
End If

If Not IsNull(rstPara.Fields("TauxGestionProjetsElec")) Then
 txtGestionProjetsElec.Text = rstPara.Fields("TauxGestionProjetsElec")
2  End If

If Not IsNull(rstPara.Fields("TauxInstallationElec")) Then
txtInstallationElec.Text = rstPara.Fields("TauxInstallationElec")
End If

2  If Not IsNull(rstPara.Fields("TauxMiseService")) Then
 txtMiseEnService.Text = rstPara.Fields("TauxMiseService")
2  End If

If Not IsNull(rstPara.Fields("TauxFormationElec")) Then
txtFormationElec.Text = rstPara.Fields("TauxFormationElec")
End If

If Not IsNull(rstPara.Fields("TauxShippingElec")) Then
 txtShippingElec.Text = rstPara.Fields("TauxShippingElec")
End If

If Not IsNull(rstPara.Fields("TauxMachinage")) Then
 txtMachinage.Text = rstPara.Fields("TauxMachinage")
End If

If Not IsNull(rstPara.Fields("TauxCoupe")) Then
 txtCoupe.Text = rstPara.Fields("TauxCoupe")
End If

If Not IsNull(rstPara.Fields("TauxSoudure")) Then
txtSoudure.Text = rstPara.Fields("TauxSoudure")
End If

3  If Not IsNull(rstPara.Fields("TauxAssemblageMec")) Then
 txtAssemblageMec.Text = rstPara.Fields("TauxAssemblageMec")
3  End If

If Not IsNull(rstPara.Fields("TauxPeinture")) Then
 txtPeinture.Text = rstPara.Fields("TauxPeinture")
 End If

40 If Not IsNull(rstPara.Fields("TauxTestMec")) Then
4 txtTestMec.Text = rstPara.Fields("TauxTestMec")
4 End If

4 If Not IsNull(rstPara.Fields("TauxGestionProjetsMec")) Then
4 txtGestionProjetsMec = rstPara.Fields("TauxGestionProjetsMec")
4 End If

4 If Not IsNull(rstPara.Fields("TauxDessinMec")) Then
4 txtDessinMec.Text = rstPara.Fields("TauxDessinMec")
4 End If

4 If Not IsNull(rstPara.Fields("TauxFormationMec")) Then
4 txtFormationMec.Text = rstPara.Fields("TauxFormationMec")
4 End If

4  If Not IsNull(rstPara.Fields("TauxInstallationMec")) Then
4  txtInstallationMec.Text = rstPara.Fields("TauxInstallationMec")
4  End If

4  If Not IsNull(rstPara.Fields("TauxShippingMec")) Then
4  txtShippingMec.Text = rstPara.Fields("TauxShippingMec")
4  End If

4  If Not IsNull(rstPara.Fields("LeGrand")) Then
4  txtLeGrand.Text = rstPara.Fields("LeGrand")
50 End If

50 If Not IsNull(rstPara.Fields("Lamine")) Then
 txtLamine.Text = rstPara.Fields("Lamine")
 End If

 If Not IsNull(rstPara.Fields("Thermo")) Then
 txtThermo.Text = rstPara.Fields("Thermo")
 End If

 If Not IsNull(rstPara.Fields("4em")) Then
 txt4em.Text = rstPara.Fields("4em")
 End If

 If Not IsNull(rstPara.Fields("fsDixMoins")) Then
 txtDixMoins.Text = rstPara.Fields("fsDixMoins")
5  End If

5  If Not IsNull(rstPara.Fields("fsDix")) Then
5  txtDixQuinze.Text = rstPara.Fields("fsDix")
5  End If

5  If Not IsNull(rstPara.Fields("fsQuinze")) Then
5  txtQuinzeVingt.Text = rstPara.Fields("fsQuinze")
5  End If

5  If Not IsNull(rstPara.Fields("fsVingt")) Then
60 txtVingtVingtCinq.Text = rstPara.Fields("fsVingt")
60 End If

  If Not IsNull(rstPara.Fields("fsVingtCinq")) Then
  txtVingtCinqCinquante.Text = rstPara.Fields("fsVingtCinq")
  End If

  If Not IsNull(rstPara.Fields("fsCinquante")) Then
  txtCinquanteCent.Text = rstPara.Fields("fsCinquante")
  End If

  If Not IsNull(rstPara.Fields("fsCent")) Then
  txtCentPlus.Text = rstPara.Fields("fsCent")
  End If

  If Not IsNull(rstPara.Fields("CheminSee4000")) Then
6  txtCheminSEE4000.Text = rstPara.Fields("CheminSee4000")
6  End If

6  If Not IsNull(rstPara.Fields("Hebergement1")) Then
6  txtHebergement1.Text = rstPara.Fields("Hebergement1")
6  End If

6  If Not IsNull(rstPara.Fields("Hebergement2")) Then
6  txtHebergement2.Text = rstPara.Fields("Hebergement2")
6  End If

70 If Not IsNull(rstPara.Fields("Repas")) Then
  txtRepas.Text = rstPara.Fields("Repas")
  End If

  If Not IsNull(rstPara.Fields("Standard")) Then
  txtStandard.Text = rstPara.Fields("Standard")
  End If

  If Not IsNull(rstPara.Fields("UniteMobile")) Then
  txtUniteMobile.Text = rstPara.Fields("UniteMobile")
  End If

  Call rstPara.Close
  Set rstPara = Nothing
 
  Exit Sub

Oups:

   wOups "frmPara", "RemplirValue", Err, Err.number, Err.Description
End Sub


Private Sub Form_Load()

 On Error GoTo Oups

 Call RemplirValue
 
 Frame1(1).Visible = True
 
 m_iCurrFrame = 1
 
 Call TbsPara_Click

 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmPara", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub TbsPara_Click()

 On Error GoTo Oups

 'si on est deja sur le l'onglet voulu on ne fait rien
 If tbsPara.SelectedItem.Index <> m_iCurrFrame Then
 Frame1(tbsPara.SelectedItem.Index).Visible = True
 Frame1(m_iCurrFrame).Visible = False
 m_iCurrFrame = tbsPara.SelectedItem.Index
 End If

 Exit Sub

Oups:

 wOups "frmPara", "TbsPara_Click", Err, Err.number, Err.Description
End Sub

