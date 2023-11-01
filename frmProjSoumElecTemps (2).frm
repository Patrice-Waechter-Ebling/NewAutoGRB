VERSION 5.00
Begin VB.Form frmProjSoumElecTemps 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Temps"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11025
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmProjSoumElecTemps.frx":0000
   ScaleHeight     =   8100
   ScaleWidth      =   11025
   Begin VB.Frame fraRessourcesHumaines 
      BackColor       =   &H00000000&
      Caption         =   "Ressources humaines"
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
      Height          =   6015
      Left            =   120
      TabIndex        =   31
      Top             =   840
      Width           =   5775
      Begin VB.TextBox txtTempsprototypeSoum 
         Height          =   285
         Left            =   2520
         TabIndex        =   123
         Top             =   5520
         Width           =   735
      End
      Begin VB.TextBox txtTempsShippingSoum 
         Height          =   285
         Left            =   2520
         TabIndex        =   117
         Top             =   5160
         Width           =   735
      End
      Begin VB.TextBox txtTempsFormationSoum 
         Height          =   285
         Left            =   2520
         TabIndex        =   111
         Top             =   4440
         Width           =   735
      End
      Begin VB.TextBox txtTempsGestionSoum 
         Height          =   285
         Left            =   2520
         TabIndex        =   104
         Top             =   4800
         Width           =   735
      End
      Begin VB.TextBox txtTempsDessinSoum 
         Height          =   285
         Left            =   2520
         TabIndex        =   40
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtTempsAssemblageSoum 
         Height          =   285
         Left            =   2520
         TabIndex        =   39
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtTempsProgInterfaceSoum 
         Height          =   285
         Left            =   2520
         TabIndex        =   38
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtTempsProgAutomateSoum 
         Height          =   285
         Left            =   2520
         TabIndex        =   37
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtTempsProgRobotSoum 
         Height          =   285
         Left            =   2520
         TabIndex        =   36
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtTempsTestSoum 
         Height          =   285
         Left            =   2520
         TabIndex        =   35
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox txtTempsInstallationSoum 
         Height          =   285
         Left            =   2520
         TabIndex        =   34
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox txtTempsMiseServiceSoum 
         Height          =   285
         Left            =   2520
         TabIndex        =   33
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox txtTempsVisionSoum 
         Height          =   285
         Left            =   2520
         TabIndex        =   32
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label lblPrixPrototype 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   4680
         TabIndex        =   128
         Top             =   5520
         Width           =   735
      End
      Begin VB.Label lblTempsPrototypeReel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3360
         TabIndex        =   127
         Top             =   5520
         Width           =   735
      End
      Begin VB.Label Label64 
         BackStyle       =   0  'Transparent
         Caption         =   "Prototypage-Développement:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   126
         Top             =   5520
         Width           =   2055
      End
      Begin VB.Label Label56 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4200
         TabIndex        =   125
         Top             =   5520
         Width           =   255
      End
      Begin VB.Label Label55 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5520
         TabIndex        =   124
         Top             =   5520
         Width           =   135
      End
      Begin VB.Label Label63 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5520
         TabIndex        =   122
         Top             =   5160
         Width           =   135
      End
      Begin VB.Label Label62 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4200
         TabIndex        =   121
         Top             =   5160
         Width           =   255
      End
      Begin VB.Label Label61 
         BackStyle       =   0  'Transparent
         Caption         =   "Expédition :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   120
         Top             =   5160
         Width           =   2055
      End
      Begin VB.Label lblTempsShippingReel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3360
         TabIndex        =   119
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label lblPrixShipping 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   4680
         TabIndex        =   118
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label Label60 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5520
         TabIndex        =   116
         Top             =   4440
         Width           =   135
      End
      Begin VB.Label Label59 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4200
         TabIndex        =   115
         Top             =   4440
         Width           =   255
      End
      Begin VB.Label Label57 
         BackStyle       =   0  'Transparent
         Caption         =   "Formation du personnel :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   114
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label lblTempsFormationReel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3360
         TabIndex        =   113
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label lblPrixFormation 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   4680
         TabIndex        =   112
         Top             =   4440
         Width           =   735
      End
      Begin VB.Line Croix1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         X1              =   2760
         X2              =   3000
         Y1              =   1440
         Y2              =   1200
      End
      Begin VB.Line Croix2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         X1              =   2760
         X2              =   3000
         Y1              =   1200
         Y2              =   1440
      End
      Begin VB.Label lblPrixGestion 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   4680
         TabIndex        =   109
         Top             =   4800
         Width           =   735
      End
      Begin VB.Label lblTempsGestionReel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3360
         TabIndex        =   108
         Top             =   4800
         Width           =   735
      End
      Begin VB.Label Label53 
         BackStyle       =   0  'Transparent
         Caption         =   "Gestion du projet :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   107
         Top             =   4800
         Width           =   2055
      End
      Begin VB.Label Label51 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4200
         TabIndex        =   106
         Top             =   4800
         Width           =   255
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5520
         TabIndex        =   105
         Top             =   4800
         Width           =   135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Dessin :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   94
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fabrication"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   93
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Assemblage :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   92
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Programmation d'automate :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   91
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Programmation d'interface :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   90
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Programmation de robot :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Test :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   88
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Temps"
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
         Left            =   2520
         TabIndex        =   87
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Prix"
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
         Left            =   4680
         TabIndex        =   86
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4200
         TabIndex        =   85
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4200
         TabIndex        =   84
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4200
         TabIndex        =   83
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4200
         TabIndex        =   82
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4200
         TabIndex        =   81
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4200
         TabIndex        =   80
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4200
         TabIndex        =   79
         Top             =   3360
         Width           =   255
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5520
         TabIndex        =   78
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5520
         TabIndex        =   77
         Top             =   1200
         Width           =   135
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5520
         TabIndex        =   76
         Top             =   1560
         Width           =   135
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5520
         TabIndex        =   75
         Top             =   1920
         Width           =   135
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5520
         TabIndex        =   74
         Top             =   2280
         Width           =   135
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5520
         TabIndex        =   73
         Top             =   2640
         Width           =   135
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5520
         TabIndex        =   72
         Top             =   3360
         Width           =   135
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5520
         TabIndex        =   71
         Top             =   3720
         Width           =   135
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4200
         TabIndex        =   70
         Top             =   3720
         Width           =   255
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "Installation :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5520
         TabIndex        =   68
         Top             =   4080
         Width           =   135
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4200
         TabIndex        =   67
         Top             =   4080
         Width           =   255
      End
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Caption         =   "Mise en service :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label Label54 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Temps"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   65
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label58 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Réel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   64
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblTempsDessinReel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3360
         TabIndex        =   63
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblTempsFabricationReel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3360
         TabIndex        =   62
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblTempsAssemblageReel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3360
         TabIndex        =   61
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblTempsProgInterfaceReel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3360
         TabIndex        =   60
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label lblTempsProgAutomateReel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3360
         TabIndex        =   59
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label lblTempsProgRobotReel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3360
         TabIndex        =   58
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label lblTempsTestReel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3360
         TabIndex        =   57
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label lblTempsInstallationReel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3360
         TabIndex        =   56
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label lblTempsMiseServiceReel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3360
         TabIndex        =   55
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label65 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5520
         TabIndex        =   54
         Top             =   3000
         Width           =   135
      End
      Begin VB.Label Label66 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4200
         TabIndex        =   53
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label Label67 
         BackStyle       =   0  'Transparent
         Caption         =   "Vision :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label lblTempsVisionReel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3360
         TabIndex        =   51
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label lblPrixDessin 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   4680
         TabIndex        =   50
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblPrixFabrication 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   4680
         TabIndex        =   49
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblPrixAssemblage 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   4680
         TabIndex        =   48
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblPrixProgInterface 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   4680
         TabIndex        =   47
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label lblPrixProgAutomate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   4680
         TabIndex        =   46
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label lblPrixProgRobot 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   4680
         TabIndex        =   45
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label lblPrixVision 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   4680
         TabIndex        =   44
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label lblPrixTest 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   4680
         TabIndex        =   43
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label lblPrixInstallation 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   4680
         TabIndex        =   42
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label lblPrixMiseService 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   4680
         TabIndex        =   41
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label lblTempsFabricationSoum 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2520
         TabIndex        =   110
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.Frame fraFraisSubsistences 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Frais de subsistances"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   6000
      TabIndex        =   7
      Top             =   840
      Width           =   4815
      Begin VB.TextBox txtTempsUniteMobile 
         Height          =   285
         Left            =   2400
         TabIndex        =   12
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtTempsRepas 
         Height          =   285
         Left            =   2400
         TabIndex        =   11
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtTempsHebergement 
         Height          =   285
         Left            =   2400
         TabIndex        =   10
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtTempsDeplacement 
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtNbrePersonne 
         Height          =   285
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   8
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4560
         TabIndex        =   30
         Top             =   1560
         Width           =   135
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4560
         TabIndex        =   29
         Top             =   1200
         Width           =   135
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4560
         TabIndex        =   28
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4560
         TabIndex        =   27
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Km"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   26
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Km"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   25
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Jours"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   24
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Jours"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   23
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Prix"
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
         Left            =   3720
         TabIndex        =   22
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   "Transport de l'unité mobile :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Transport / déplacement : "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "Repas :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "Hébergement :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "pers."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   17
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblPrixHebergement 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   3720
         TabIndex        =   16
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblPrixRepas 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   3720
         TabIndex        =   15
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblPrixDeplacement 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   3720
         TabIndex        =   14
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblPrixUniteMobile 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   3720
         TabIndex        =   13
         Top             =   1560
         Width           =   735
      End
   End
   Begin VB.Frame fraManutention 
      BackColor       =   &H00000000&
      Caption         =   "Manutention"
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
      Height          =   855
      Left            =   6000
      TabIndex        =   2
      Top             =   2880
      Width           =   4815
      Begin VB.TextBox txtPrixEmballage 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3720
         TabIndex        =   3
         Text            =   "0"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Frais de transport / emballage :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Prix"
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
         Left            =   3720
         TabIndex        =   5
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4560
         TabIndex        =   4
         Top             =   480
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   9600
      TabIndex        =   1
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdDetail 
      Appearance      =   0  'Flat
      Caption         =   "Détails"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   7560
      Width           =   735
   End
   Begin VB.Label Label52 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total de la ressource humaine :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   103
      Top             =   7080
      Width           =   2295
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8880
      TabIndex        =   102
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10560
      TabIndex        =   101
      Top             =   3960
      Width           =   135
   End
   Begin VB.Label Label50 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Hrs"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   100
      Top             =   7080
      Width           =   255
   End
   Begin VB.Label lblDollarRH 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5640
      TabIndex        =   99
      Top             =   7080
      Width           =   135
   End
   Begin VB.Label lblTotalTempsRHSoum 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2640
      TabIndex        =   98
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label lblTotalTempsRHReel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3480
      TabIndex        =   97
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label lblTotalPrixRH 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4800
      TabIndex        =   96
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   285
      Left            =   9720
      TabIndex        =   95
      Top             =   3960
      Width           =   735
   End
End
Attribute VB_Name = "frmProjSoumElecTemps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub Afficher(ByVal sNoProjSoum As String, ByVal iType As Integer, ByVal iMode As Integer, ByVal bNouveauTaux As Boolean)

 On Error GoTo Oups
 
 m_eType = iType
 
 m_eMode = iMode
 
 m_sNoProjSoum = sNoProjSoum
 
 m_bNouveauTaux = bNouveauTaux
 
 If bNouveauTaux = True Then
 Call InitialiserVariablesConfig
 Else
 Call InitialiserVariablesProjSoum
 End If
 
 If m_eMode = MODE_AJOUT_MODIF Then
  Call BarrerChamps(False)
  Else
  Call BarrerChamps(True)
  End If
 
  Call AfficherEnregistrement
 
  Call Me.Show(vbModal)

  Exit Sub

Oups:

  wOups "frmProjSoumElecTemps", "Afficher", Err, Err.number, Err.Description
End Sub

Private Sub AfficherEnregistrement()

 On Error GoTo Oups

 Dim rstProjSoum As ADODB.Recordset
 Dim rstSoum As ADODB.Recordset
 Dim sChamps As String
 Dim sTable As String
 
 If m_eType = TYPE_PROJET Then
 sChamps = "IDProjet"
 sTable = "GrbProjetElec"
 Else
 sChamps = "IDSoumission"
 sTable = "GrbSoumissionElec"
  End If
 
  Set rstProjSoum = New ADODB.Recordset
 
  Call rstProjSoum.Open("SELECT * FROM " & sTable & " WHERE " & sChamps & " = '" & m_sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
  If Not rstProjSoum.EOF And FrmProjSoumElec.m_bTempsDejaOuvert = False And m_eMode = MODE_INACTIF Then
  m_bSansTemps = rstProjSoum.Fields("SansTemps")

  If Not IsNull(rstProjSoum.Fields("TempsDessin")) Then
  txtTempsDessinSoum.Text = rstProjSoum.Fields("TempsDessin")
  Else
 txtTempsDessinSoum.Text = "0"
1 End If

 If m_eType = TYPE_SOUMISSION Then
 If Not IsNull(rstProjSoum.Fields("TempsFabrication")) Then
 lblTempsFabricationSoum.Caption = rstProjSoum.Fields("TempsFabrication")
 Else
 lblTempsFabricationSoum.Caption = "0"
 End If
 Else
 lblTempsFabricationSoum.Caption = CalculerTempsFabrication
 End If

 If Not IsNull(rstProjSoum.Fields("TempsAssemblage")) Then
 txtTempsAssemblageSoum.Text = rstProjSoum.Fields("TempsAssemblage")
 Else
 txtTempsAssemblageSoum.Text = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsProgInterface")) Then
 txtTempsProgInterfaceSoum.Text = rstProjSoum.Fields("TempsProgInterface")
 Else
1  txtTempsProgInterfaceSoum.Text = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsProgAutomate")) Then
 txtTempsProgAutomateSoum.Text = rstProjSoum.Fields("TempsProgAutomate")
 Else
 txtTempsProgAutomateSoum.Text = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsProgRobot")) Then
 txtTempsProgRobotSoum.Text = rstProjSoum.Fields("TempsProgRobot")
 Else
 txtTempsProgRobotSoum.Text = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsVision")) Then
 txtTempsVisionSoum.Text = rstProjSoum.Fields("TempsVision")
 Else
 txtTempsVisionSoum.Text = "0"
 End If

If Not IsNull(rstProjSoum.Fields("TempsTest")) Then
 txtTempsTestSoum.Text = rstProjSoum.Fields("TempsTest")
Else
 txtTempsTestSoum.Text = "0"
End If

3 If Not IsNull(rstProjSoum.Fields("TempsInstallation")) Then
 txtTempsInstallationSoum.Text = rstProjSoum.Fields("TempsInstallation")
 Else
 txtTempsInstallationSoum.Text = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsMiseService")) Then
 txtTempsMiseServiceSoum.Text = rstProjSoum.Fields("TempsMiseService")
 Else
 txtTempsMiseServiceSoum.Text = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsFormation")) Then
 txtTempsFormationSoum.Text = rstProjSoum.Fields("TempsFormation")
 Else
 txtTempsFormationSoum.Text = "0"
 End If

If Not IsNull(rstProjSoum.Fields("TempsGestion")) Then
 txtTempsGestionSoum.Text = rstProjSoum.Fields("TempsGestion")
 Else
 txtTempsGestionSoum.Text = "0"
End If

4 If Not IsNull(rstProjSoum.Fields("TempsShipping")) Then
4 txtTempsShippingSoum.Text = rstProjSoum.Fields("TempsShipping")
4 Else
4 txtTempsShippingSoum.Text = "0"
4 End If
 txtTempsprototypeSoum.Text = "0"

4 If m_bSansTemps = True Then
4 Croix1.Visible = True
4 Croix2.Visible = True
4 Else
4 Croix1.Visible = False
4 Croix2.Visible = False
4  End If

4  If Right$(m_sNoProjSoum, 2) <> "99" Then
4  If Not IsNull(rstProjSoum.Fields("NbrePersonne")) Then
4  txtNbrePersonne.Text = rstProjSoum.Fields("NbrePersonne")
4  Else
4  txtNbrePersonne.Text = "0"
4  End If

4  If Not IsNull(rstProjSoum.Fields("TempsHebergement")) Then
50 txtTempsHebergement.Text = rstProjSoum.Fields("TempsHebergement")
Else
 txtTempsHebergement.Text = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsRepas")) Then
 txtTempsRepas.Text = rstProjSoum.Fields("TempsRepas")
 Else
 txtTempsRepas.Text = "0"
 End If
 Else
 If Not IsNull(rstProjSoum.Fields("TotalHebergement")) Then
 lblPrixHebergement.Caption = rstProjSoum.Fields("TotalHebergement")
5  End If

5  If Not IsNull(rstProjSoum.Fields("TotalRepas")) Then
5  lblPrixRepas.Caption = rstProjSoum.Fields("TotalRepas")
5  End If
5  End If

5  If Not IsNull(rstProjSoum.Fields("TempsTransport")) Then
5  txtTempsDeplacement.Text = rstProjSoum.Fields("TempsTransport")
5  Else
60 txtTempsDeplacement.Text = "0"
  End If

  If Not IsNull(rstProjSoum.Fields("TempsUniteMobile")) Then
  txtTempsUniteMobile.Text = rstProjSoum.Fields("TempsUniteMobile")
  Else
  txtTempsUniteMobile.Text = "0"
  End If

  If Not IsNull(rstProjSoum.Fields("PrixEmballage")) Then
  txtPrixEmballage.Text = rstProjSoum.Fields("PrixEmballage")
  Else
  txtPrixEmballage.Text = "0"
  End If
6  Else
6  If m_eType = TYPE_SOUMISSION Then
6  m_bSansTemps = FrmProjSoumElec.m_bSansTemps

6  txtTempsDessinSoum.Text = FrmProjSoumElec.m_sTempsDessin
6  lblTempsFabricationSoum.Caption = FrmProjSoumElec.m_sTempsFabrication
6  txtTempsAssemblageSoum.Text = FrmProjSoumElec.m_sTempsAssemblage
6  txtTempsProgInterfaceSoum.Text = FrmProjSoumElec.m_sTempsProgInterface
6  txtTempsProgAutomateSoum.Text = FrmProjSoumElec.m_sTempsProgAutomate
70 txtTempsProgRobotSoum.Text = FrmProjSoumElec.m_sTempsProgRobot
  txtTempsVisionSoum.Text = FrmProjSoumElec.m_sTempsVision
  txtTempsTestSoum.Text = FrmProjSoumElec.m_sTempsTest
  txtTempsInstallationSoum.Text = FrmProjSoumElec.m_sTempsInstallation
  txtTempsMiseServiceSoum.Text = FrmProjSoumElec.m_sTempsMiseService
  txtTempsFormationSoum.Text = FrmProjSoumElec.m_sTempsFormation
  txtTempsGestionSoum.Text = FrmProjSoumElec.m_sTempsGestion
  txtTempsShippingSoum.Text = FrmProjSoumElec.m_sTempsShipping
  Else
  If Not rstProjSoum.EOF And Not IsNull(rstProjSoum.Fields("IDSoumission")) Then
  If rstProjSoum.Fields("IDSoumission") <> "" Then
  Set rstSoum = New ADODB.Recordset

   Call rstSoum.Open("SELECT * FROM GrbSoumissionElec WHERE IDSoumission = '" & rstProjSoum.Fields("IDSoumission") & "'", g_connData, adOpenDynamic, adLockOptimistic)

   If Not rstSoum.EOF Then
7  m_bSansTemps = FrmProjSoumElec.m_bSansTemps

7  If Not IsNull(rstSoum.Fields("TempsDessin")) Then
7  txtTempsDessinSoum.Text = rstSoum.Fields("TempsDessin")
7  Else
7  txtTempsDessinSoum.Text = "0"
7  End If

80 If m_eType = TYPE_PROJET Then
  lblTempsFabricationSoum.Caption = CalculerTempsFabrication
  Else
  If Not IsNull(rstSoum.Fields("TempsFabrication")) Then
  lblTempsFabricationSoum.Caption = rstSoum.Fields("TempsFabrication")
  Else
  lblTempsFabricationSoum.Caption = "0"
  End If
  End If

  If Not IsNull(rstSoum.Fields("TempsAssemblage")) Then
  txtTempsAssemblageSoum.Text = rstSoum.Fields("TempsAssemblage")
  Else
   txtTempsAssemblageSoum.Text = "0"
   End If

   If Not IsNull(rstSoum.Fields("TempsProgInterface")) Then
   txtTempsProgInterfaceSoum.Text = rstSoum.Fields("TempsProgInterface")
8  Else
8  txtTempsProgInterfaceSoum.Text = "0"
8  End If

8  If Not IsNull(rstSoum.Fields("TempsProgAutomate")) Then
90 txtTempsProgAutomateSoum.Text = rstSoum.Fields("TempsProgAutomate")
  Else
  txtTempsProgAutomateSoum.Text = "0"
  End If

  If Not IsNull(rstSoum.Fields("TempsProgRobot")) Then
  txtTempsProgRobotSoum.Text = rstSoum.Fields("TempsProgRobot")
  Else
  txtTempsProgRobotSoum.Text = "0"
  End If

  If Not IsNull(rstSoum.Fields("TempsVision")) Then
  txtTempsVisionSoum.Text = rstSoum.Fields("TempsVision")
  Else
 txtTempsVisionSoum.Text = "0"
   End If

 If Not IsNull(rstSoum.Fields("TempsTest")) Then
   txtTempsTestSoum.Text = rstSoum.Fields("TempsTest")
 Else
   txtTempsTestSoum.Text = "0"
 End If

9  If Not IsNull(rstSoum.Fields("TempsInstallation")) Then
 txtTempsInstallationSoum.Text = rstSoum.Fields("TempsInstallation")
 Else
 txtTempsInstallationSoum.Text = "0"
 End If

 If Not IsNull(rstSoum.Fields("TempsMiseService")) Then
 txtTempsMiseServiceSoum.Text = rstSoum.Fields("TempsMiseService")
 Else
 txtTempsMiseServiceSoum.Text = "0"
 End If

 If Not IsNull(rstSoum.Fields("TempsFormation")) Then
 txtTempsFormationSoum.Text = rstSoum.Fields("TempsFormation")
 Else
10  txtTempsFormationSoum.Text = "0"
10  End If

10  If Not IsNull(rstSoum.Fields("TempsGestion")) Then
10  txtTempsGestionSoum.Text = rstSoum.Fields("TempsGestion")
10  Else
10  txtTempsGestionSoum.Text = "0"
10  End If

10  If Not IsNull(rstSoum.Fields("TempsShipping")) Then
1 txtTempsShippingSoum.Text = rstSoum.Fields("TempsShipping")
1 Else
1 txtTempsShippingSoum.Text = "0"
1 End If
1 Else
1 m_bSansTemps = False

1 txtTempsDessinSoum.Text = "0"
1 lblTempsFabricationSoum.Caption = CalculerTempsFabrication
1 txtTempsAssemblageSoum.Text = "0"
1 txtTempsProgInterfaceSoum.Text = "0"
1 txtTempsProgAutomateSoum.Text = "0"
1 txtTempsProgRobotSoum.Text = "0"
1 txtTempsVisionSoum.Text = "0"
1 txtTempsTestSoum.Text = "0"
 txtTempsInstallationSoum.Text = "0"
1 txtTempsMiseServiceSoum.Text = "0"
 txtTempsFormationSoum.Text = "0"
1 txtTempsGestionSoum.Text = "0"
 txtTempsShippingSoum.Text = "0"
11  End If

 Call rstSoum.Close
 Set rstSoum = Nothing
1 Else
1 m_bSansTemps = False

1 txtTempsDessinSoum.Text = "0"
1 lblTempsFabricationSoum.Caption = CalculerTempsFabrication
1 txtTempsAssemblageSoum.Text = "0"
1 txtTempsProgInterfaceSoum.Text = "0"
1 txtTempsProgAutomateSoum.Text = "0"
1 txtTempsProgRobotSoum.Text = "0"
1 txtTempsVisionSoum.Text = "0"
1 txtTempsTestSoum.Text = "0"
1 txtTempsInstallationSoum.Text = "0"
1 txtTempsMiseServiceSoum.Text = "0"
1 txtTempsFormationSoum.Text = "0"
1 txtTempsGestionSoum.Text = "0"
1 txtTempsShippingSoum.Text = "0"
1 End If
1 Else
1 m_bSansTemps = False

1 txtTempsDessinSoum.Text = "0"
1 lblTempsFabricationSoum.Caption = CalculerTempsFabrication
1 txtTempsAssemblageSoum.Text = "0"
1 txtTempsProgInterfaceSoum.Text = "0"
1 txtTempsProgAutomateSoum.Text = "0"
1 txtTempsProgRobotSoum.Text = "0"
1 txtTempsVisionSoum.Text = "0"
1 txtTempsTestSoum.Text = "0"
1 txtTempsInstallationSoum.Text = "0"
1 txtTempsMiseServiceSoum.Text = "0"
1 txtTempsFormationSoum.Text = "0"
1 txtTempsGestionSoum.Text = "0"
1 txtTempsShippingSoum.Text = "0"
1 End If
13  End If

1 If m_bSansTemps = True Then
1 Croix1.Visible = True
1 Croix2.Visible = True
1Else
1 Croix1.Visible = False
1 Croix2.Visible = False
14End If

14 txtNbrePersonne.Text = FrmProjSoumElec.m_sNbrePersonne
14 txtTempsHebergement.Text = FrmProjSoumElec.m_sTempsHebergement
14 txtTempsRepas.Text = FrmProjSoumElec.m_sTempsRepas
14 txtTempsDeplacement.Text = FrmProjSoumElec.m_sTempsTransport
14 txtTempsUniteMobile.Text = FrmProjSoumElec.m_sTempsUniteMobile
14 txtPrixEmballage.Text = FrmProjSoumElec.m_sPrixEmballage
14 End If

14 If m_eType = TYPE_PROJET Then
14 Call AfficherTempsReels

14 Call CalculerTotalArgent
146End If

14  Call rstProjSoum.Close
14  Set rstProjSoum = Nothing

14  Exit Sub

Oups:

14  wOups "frmProjSoumElecTemps", "AfficherEnregistrement", Err, Err.number, Err.Description
End Sub

Private Sub AfficherTempsReels()
 
 On Error GoTo Oups

 Dim rstPunch As ADODB.Recordset
 Dim sDateDebut As String
 Dim sDateFin As String
 Dim sTotal As String
 Dim sFilterNoProjet As String

 sDateDebut = "TIMESERIAL(Left(GrbPunch.HeureDébut,2),RIGHT(GrbPunch.HeureDébut,2),0)"

 sDateFin = "TIMESERIAL(Left(GrbPunch.HeureFin,2),RIGHT(GrbPunch.HeureFin,2),0)"

 sTotal = "(SUM(" & sDateFin & " - " & sDateDebut & ")* 24) As Total"

 If Right$(m_sNoProjSoum, 2) = "99" Then
 sFilterNoProjet = "LEFT(NoProjet, 6) = '" & Left$(m_sNoProjSoum, 6) & "'"
  Else
  sFilterNoProjet = "NoProjet = '" & m_sNoProjSoum & "'"
  End If

  Set rstPunch = New ADODB.Recordset

  Call rstPunch.Open("SELECT Type, " & sTotal & " FROM GrbPunch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null GROUP BY Type", g_connData, adOpenDynamic, adLockOptimistic)

  lblTempsDessinReel.Caption = "0"
  lblTempsFabricationReel.Caption = "0"
  lblTempsAssemblageReel.Caption = "0"
10 lblTempsProgInterfaceReel.Caption = "0"
lblTempsProgAutomateReel.Caption = "0"
lblTempsProgRobotReel.Caption = "0"
lblTempsVisionReel.Caption = "0"
lblTempsTestReel.Caption = "0"
lblTempsInstallationReel.Caption = "0"
lblTempsMiseServiceReel.Caption = "0"
lblTempsFormationReel.Caption = "0"
lblTempsGestionReel.Caption = "0"
lblTempsShippingReel.Caption = "0"
 lblTempsPrototypeReel.Caption = "0"

Do While Not rstPunch.EOF
 If Not IsNull(rstPunch.Fields("Total")) Then
 Select Case rstPunch.Fields("Type")
 Case "Dessin": lblTempsDessinReel.Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Fabrication": lblTempsFabricationReel.Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Assemblage": lblTempsAssemblageReel.Caption = Round(rstPunch.Fields("Total"), 2)
 Case "ProgInterface": lblTempsProgInterfaceReel.Caption = Round(rstPunch.Fields("Total"), 2)
 Case "ProgAutomate": lblTempsProgAutomateReel.Caption = Round(rstPunch.Fields("Total"), 2)
 Case "ProgRobot": lblTempsProgRobotReel.Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Vision": lblTempsVisionReel.Caption = Round(rstPunch.Fields("Total"), 2)
1  Case "Test": lblTempsTestReel.Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Installation": lblTempsInstallationReel.Caption = Round(rstPunch.Fields("Total"), 2)
 Case "MiseService": lblTempsMiseServiceReel.Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Formation": lblTempsFormationReel.Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Gestion": lblTempsGestionReel.Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Shipping": lblTempsShippingReel.Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Prototypage-Dévelloppement expérimental": lblTempsPrototypeReel.Caption = Round(rstPunch.Fields("Total"), 2)
 End Select
 End If

 Call rstPunch.MoveNext
Loop

Call rstPunch.Close

 'Ouverture des enregistrements avec comme filtre, le numéro du projet
Call rstPunch.Open("SELECT " & sTotal & " FROM GrbPunch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null", g_connData, adOpenDynamic, adLockOptimistic)

If Not IsNull(rstPunch.Fields("Total")) Then
lblTotalTempsRHReel.Caption = Round(rstPunch.Fields("Total"), 2)
Else
lblTotalTempsRHReel.Caption = "0"
End If

2  Call rstPunch.Close
Set rstPunch = Nothing

2  Exit Sub

Oups:

wOups "frmProjSoumElecTemps", "AfficherTempsReels", Err, Err.number, Err.Description
End Sub

Private Sub CalculerTotalArgent()

 On Error GoTo Oups

 If IsNumeric(lblTempsDessinReel.Caption) Then
 lblPrixDessin.Caption = Round(Replace(lblTempsDessinReel.Caption * m_sTauxDessin, ".", ","), 2)
 Else
 lblPrixDessin.Caption = 0
 End If

 If IsNumeric(lblTempsFabricationReel.Caption) Then
 lblPrixFabrication.Caption = Round(Replace(lblTempsFabricationReel.Caption * m_sTauxFabrication, ".", ","), 2)
 Else
 lblPrixFabrication.Caption = 0
 End If

  If IsNumeric(lblTempsAssemblageReel.Caption) Then
  lblPrixAssemblage.Caption = Round(Replace(lblTempsAssemblageReel.Caption * m_sTauxAssemblage, ".", ","), 2)
  Else
  lblPrixAssemblage.Caption = 0
  End If

  If IsNumeric(lblTempsProgInterfaceReel.Caption) Then
  lblPrixProgInterface.Caption = Round(Replace(lblTempsProgInterfaceReel.Caption * m_sTauxProgInterface, ".", ","), 2)
  Else
lblPrixProgInterface.Caption = 0
End If

If IsNumeric(lblTempsProgAutomateReel.Caption) Then
 lblPrixProgAutomate.Caption = Round(Replace(lblTempsProgAutomateReel.Caption * m_sTauxProgAutomate, ".", ","), 2)
Else
 lblPrixProgAutomate.Caption = 0
End If

If IsNumeric(lblTempsProgRobotReel.Caption) Then
 lblPrixProgRobot.Caption = Round(Replace(lblTempsProgRobotReel.Caption * m_sTauxProgRobot, ".", ","), 2)
Else
 lblPrixProgRobot.Caption = 0
End If

1  If IsNumeric(lblTempsVisionReel.Caption) Then
 lblPrixVision.Caption = Round(Replace(lblTempsVisionReel.Caption * m_sTauxVision, ".", ","), 2)
 Else
 lblPrixVision.Caption = 0
 End If

If IsNumeric(lblTempsTestReel.Caption) Then
 lblPrixTest.Caption = Round(Replace(lblTempsTestReel.Caption * m_sTauxTest, ".", ","), 2)
1  Else
 lblPrixTest.Caption = 0
 End If

If IsNumeric(lblTempsInstallationReel.Caption) Then
 lblPrixInstallation.Caption = Round(Replace(lblTempsInstallationReel.Caption * m_sTauxInstallation, ".", ","), 2)
Else
 lblPrixInstallation.Caption = 0
End If

If IsNumeric(lblTempsMiseServiceReel.Caption) Then
 lblPrixMiseService.Caption = Round(Replace(lblTempsMiseServiceReel.Caption * m_sTauxMiseService, ".", ","), 2)
Else
 lblPrixMiseService.Caption = 0
End If

2  If IsNumeric(lblTempsFormationReel.Caption) Then
 lblPrixFormation.Caption = Round(Replace(lblTempsFormationReel.Caption * m_sTauxFormation, ".", ","), 2)
2  Else
 lblPrixFormation.Caption = 0
2  End If

If IsNumeric(lblTempsGestionReel.Caption) Then
lblPrixGestion.Caption = Round(Replace(lblTempsGestionReel.Caption * m_sTauxGestion, ".", ","), 2)
Else
lblPrixGestion.Caption = 0
End If

If IsNumeric(lblTempsShippingReel.Caption) Then
 lblPrixShipping.Caption = Round(Replace(lblTempsShippingReel.Caption * m_sTauxShipping, ".", ","), 2)
Else
 lblPrixShipping.Caption = 0
End If

 If IsNumeric(lblTempsPrototypeReel.Caption) Then
3 lblPrixPrototype.Caption = Round(Replace(lblTempsPrototypeReel.Caption * m_sTauxGestion, ".", ","), 2)
33 Else
33 lblPrixPrototype.Caption = 0
334 End If






Call CalculerTotal

Exit Sub

Oups:

wOups "frmProjSoumElecTemps", "CalculerTotalArgent", Err, Err.number, Err.Description
End Sub

Private Sub BarrerChamps(ByVal bLocked As Boolean)

 On Error GoTo Oups

 txtTempsDessinSoum.Locked = bLocked
 txtTempsAssemblageSoum.Locked = bLocked
 txtTempsProgInterfaceSoum.Locked = bLocked
 txtTempsProgAutomateSoum.Locked = bLocked
 txtTempsProgRobotSoum.Locked = bLocked
 txtTempsVisionSoum.Locked = bLocked
 txtTempsTestSoum.Locked = bLocked
 txtTempsInstallationSoum.Locked = bLocked
 txtTempsMiseServiceSoum.Locked = bLocked
 txtTempsFormationSoum.Locked = bLocked
  txtTempsGestionSoum.Locked = bLocked
  txtTempsShippingSoum.Locked = bLocked

  txtNbrePersonne.Locked = bLocked
  txtTempsHebergement.Locked = bLocked
  txtTempsRepas.Locked = bLocked
  txtTempsDeplacement.Locked = bLocked
  txtTempsUniteMobile.Locked = bLocked
 
  txtPrixEmballage.Locked = bLocked

10 Exit Sub

Oups:

wOups "frmProjSoumElecTemps", "BarrerChamps", Err, Err.number, Err.Description
End Sub

Private Sub cmdDetail_Click()

 On Error GoTo Oups

 If m_eType = TYPE_PROJET Then
 Call frmDetailTemps.Afficher(m_sNoProjSoum, ELECTRIQUE, True)
 Else
 Call frmDetailTemps.Afficher(m_sNoProjSoum, ELECTRIQUE, False)
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "cmdDetail_Click", Err, Err.number, Err.Description
End Sub

Private Sub Cmdfermer_Click()

 On Error GoTo Oups

 If m_eMode = MODE_AJOUT_MODIF Then
 Call EnregistrerTemps
 End If

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "cmdFermer_Click", Err, Err.number, Err.Description
End Sub

Private Sub EnregistrerTemps()

 On Error GoTo Oups

 If m_eType = TYPE_SOUMISSION Then
 If Trim$(txtTempsDessinSoum.Text) <> vbNullString And IsNumeric(txtTempsDessinSoum.Text) Then
 FrmProjSoumElec.m_sTempsDessin = txtTempsDessinSoum.Text
 Else
 FrmProjSoumElec.m_sTempsDessin = "0"
 End If
 
 If Trim$(lblTempsFabricationSoum.Caption) <> vbNullString Then
 FrmProjSoumElec.m_sTempsFabrication = lblTempsFabricationSoum.Caption
 Else
 FrmProjSoumElec.m_sTempsFabrication = "0"
  End If
 
  If Trim$(txtTempsAssemblageSoum.Text) <> vbNullString And IsNumeric(txtTempsAssemblageSoum.Text) Then
  FrmProjSoumElec.m_sTempsAssemblage = txtTempsAssemblageSoum.Text
  Else
  FrmProjSoumElec.m_sTempsAssemblage = "0"
  End If
 
  If Trim$(txtTempsProgInterfaceSoum.Text) <> vbNullString And IsNumeric(txtTempsProgInterfaceSoum.Text) Then
  FrmProjSoumElec.m_sTempsProgInterface = txtTempsProgInterfaceSoum.Text
Else
FrmProjSoumElec.m_sTempsProgInterface = "0"
 End If
 
 If Trim$(txtTempsProgAutomateSoum.Text) <> vbNullString And IsNumeric(txtTempsProgAutomateSoum.Text) Then
 FrmProjSoumElec.m_sTempsProgAutomate = txtTempsProgAutomateSoum.Text
 Else
 FrmProjSoumElec.m_sTempsProgAutomate = "0"
 End If
 
 If Trim$(txtTempsProgRobotSoum.Text) <> vbNullString And IsNumeric(txtTempsProgRobotSoum.Text) Then
 FrmProjSoumElec.m_sTempsProgRobot = txtTempsProgRobotSoum.Text
 Else
 FrmProjSoumElec.m_sTempsProgRobot = "0"
End If
 
 If Trim$(txtTempsVisionSoum.Text) <> vbNullString And IsNumeric(txtTempsVisionSoum.Text) Then
 FrmProjSoumElec.m_sTempsVision = txtTempsVisionSoum.Text
 Else
 FrmProjSoumElec.m_sTempsVision = "0"
 End If
 
 If Trim$(txtTempsTestSoum.Text) <> vbNullString And IsNumeric(txtTempsTestSoum.Text) Then
1  FrmProjSoumElec.m_sTempsTest = txtTempsTestSoum.Text
 Else
 FrmProjSoumElec.m_sTempsTest = "0"
 End If
 
 If Trim$(txtTempsInstallationSoum.Text) <> vbNullString And IsNumeric(txtTempsInstallationSoum.Text) Then
 FrmProjSoumElec.m_sTempsInstallation = txtTempsInstallationSoum.Text
 Else
 FrmProjSoumElec.m_sTempsInstallation = "0"
 End If
 
 If Trim$(txtTempsMiseServiceSoum.Text) <> vbNullString And IsNumeric(txtTempsMiseServiceSoum.Text) Then
 FrmProjSoumElec.m_sTempsMiseService = txtTempsMiseServiceSoum.Text
 Else
 FrmProjSoumElec.m_sTempsMiseService = "0"
End If
 
 If Trim$(txtTempsFormationSoum.Text) <> vbNullString And IsNumeric(txtTempsFormationSoum.Text) Then
 FrmProjSoumElec.m_sTempsFormation = txtTempsFormationSoum.Text
 Else
 FrmProjSoumElec.m_sTempsFormation = "0"
 End If
 
If Trim$(txtTempsGestionSoum.Text) <> vbNullString And IsNumeric(txtTempsGestionSoum.Text) Then
 FrmProjSoumElec.m_sTempsGestion = txtTempsGestionSoum.Text
Else
FrmProjSoumElec.m_sTempsGestion = "0"
 End If

 If Trim$(txtTempsShippingSoum.Text) <> vbNullString And IsNumeric(txtTempsShippingSoum.Text) Then
 FrmProjSoumElec.m_sTempsShipping = txtTempsShippingSoum.Text
 Else
 FrmProjSoumElec.m_sTempsShipping = "0"
 End If
End If

 
If m_bSansTemps = True Then
 FrmProjSoumElec.m_bSansTemps = True
 FrmProjSoumElec.tmrTemps.Enabled = True
3  Else
 FrmProjSoumElec.m_bSansTemps = False
FrmProjSoumElec.tmrTemps.Enabled = False
 FrmProjSoumElec.lblPasTemps.Visible = False
3  End If

If Trim$(txtNbrePersonne.Text) <> vbNullString And IsNumeric(txtNbrePersonne.Text) Then
 FrmProjSoumElec.m_sNbrePersonne = txtNbrePersonne.Text
 Else
FrmProjSoumElec.m_sNbrePersonne = "0"
End If
 
4 If Trim$(txtTempsHebergement.Text) <> vbNullString And IsNumeric(txtTempsHebergement.Text) Then
4 FrmProjSoumElec.m_sTempsHebergement = txtTempsHebergement.Text
4 Else
4 FrmProjSoumElec.m_sTempsHebergement = "0"
4 End If
 
4 If Trim$(txtTempsRepas.Text) <> vbNullString And IsNumeric(txtTempsRepas.Text) Then
4 FrmProjSoumElec.m_sTempsRepas = txtTempsRepas.Text
4 Else
4 FrmProjSoumElec.m_sTempsRepas = "0"
4 End If
 
4  If Trim$(txtTempsDeplacement.Text) <> vbNullString And IsNumeric(txtTempsDeplacement.Text) Then
4  FrmProjSoumElec.m_sTempsTransport = txtTempsDeplacement.Text
4  Else
4  FrmProjSoumElec.m_sTempsTransport = "0"
4  End If
 
4  If Trim$(txtTempsUniteMobile.Text) <> vbNullString And IsNumeric(txtTempsUniteMobile.Text) Then
4  FrmProjSoumElec.m_sTempsUniteMobile = txtTempsUniteMobile.Text
4  Else
50 FrmProjSoumElec.m_sTempsUniteMobile = "0"
50 End If
 
 If Trim$(txtPrixEmballage.Text) <> vbNullString And IsNumeric(txtPrixEmballage.Text) Then
 FrmProjSoumElec.m_sPrixEmballage = txtPrixEmballage.Text
 Else
 FrmProjSoumElec.m_sPrixEmballage = "0"
 End If
 
 FrmProjSoumElec.m_sTauxHebergement1 = m_sHebergement1
 FrmProjSoumElec.m_sTauxHebergement2 = m_sHebergement2
 FrmProjSoumElec.m_sTauxRepas = m_sRepas
 FrmProjSoumElec.m_sTauxTransport = m_sStandard
 FrmProjSoumElec.m_sTauxUniteMobile = m_sUniteMobile

5  Exit Sub

Oups:

5  wOups "frmProjSoumElecTemps", "EnregistrerTemps", Err, Err.number, Err.Description
End Sub

Private Sub InitialiserVariablesConfig()

 On Error GoTo Oups

 'Initialise les variables à partir de la table Config (Pour avoir le taux
 'horaire le plus récent)
 Dim rstConfig As ADODB.Recordset
 
 Set rstConfig = New ADODB.Recordset
 
 Call rstConfig.Open("SELECT TauxDessinElec, TauxFabrication, TauxAssemblageElec, TauxProgInterface, TauxProgAutomate, TauxProgRobot, TauxVision, TauxTestElec, TauxInstallationElec, TauxMiseService, TauxFormationElec, TauxGestionProjetsElec, TauxShippingElec, Repas, Hebergement1, Hebergement2, Standard, UniteMobile FROM GrbConfig", g_connData, adOpenDynamic, adLockOptimistic)
 
 If Not IsNull(rstConfig.Fields("TauxDessinElec")) Then
 m_sTauxDessin = rstConfig.Fields("TauxDessinElec")
 Else
 m_sTauxDessin = "0"
 End If

 If Not IsNull(rstConfig.Fields("TauxFabrication")) Then
 m_sTauxFabrication = rstConfig.Fields("TauxFabrication")
  Else
  m_sTauxFabrication = "0"
  End If

  If Not IsNull(rstConfig.Fields("TauxAssemblageElec")) Then
  m_sTauxAssemblage = rstConfig.Fields("TauxAssemblageElec")
  Else
  m_sTauxAssemblage = "0"
  End If

10 If Not IsNull(rstConfig.Fields("TauxProgInterface")) Then
1 m_sTauxProgInterface = rstConfig.Fields("TauxProgInterface")
Else
 m_sTauxProgInterface = "0"
End If

If Not IsNull(rstConfig.Fields("TauxProgAutomate")) Then
 m_sTauxProgAutomate = rstConfig.Fields("TauxProgAutomate")
Else
 m_sTauxProgAutomate = "0"
End If

If Not IsNull(rstConfig.Fields("TauxProgRobot")) Then
 m_sTauxProgRobot = rstConfig.Fields("TauxProgRobot")
1  Else
 m_sTauxProgRobot = "0"
 End If

If Not IsNull(rstConfig.Fields("TauxVision")) Then
 m_sTauxVision = rstConfig.Fields("TauxVision")
Else
 m_sTauxVision = "0"
1  End If

 If Not IsNull(rstConfig.Fields("TauxTestElec")) Then
 m_sTauxTest = rstConfig.Fields("TauxTestElec")
Else
 m_sTauxTest = "0"
End If

If Not IsNull(rstConfig.Fields("TauxInstallationElec")) Then
 m_sTauxInstallation = rstConfig.Fields("TauxInstallationElec")
Else
 m_sTauxInstallation = "0"
End If

If Not IsNull(rstConfig.Fields("TauxMiseService")) Then
 m_sTauxMiseService = rstConfig.Fields("TauxMiseService")
2  Else
 m_sTauxMiseService = "0"
2  End If

If Not IsNull(rstConfig.Fields("TauxFormationElec")) Then
m_sTauxFormation = rstConfig.Fields("TauxFormationElec")
Else
m_sTauxFormation = "0"
End If

30 If Not IsNull(rstConfig.Fields("TauxGestionProjetsElec")) Then
3 m_sTauxGestion = rstConfig.Fields("TauxGestionProjetsElec")
Else
 m_sTauxGestion = "0"
End If

If Not IsNull(rstConfig.Fields("TauxShippingElec")) Then
 m_sTauxShipping = rstConfig.Fields("TauxShippingElec")
Else
 m_sTauxShipping = "0"
End If
 
m_sRepas = rstConfig.Fields("Repas")
m_sHebergement1 = rstConfig.Fields("Hebergement1")
3  m_sHebergement2 = rstConfig.Fields("Hebergement2")
m_sStandard = rstConfig.Fields("Standard")
3  m_sUniteMobile = rstConfig.Fields("UniteMobile")
 
Call rstConfig.Close
3  Set rstConfig = Nothing

Exit Sub

Oups:

3  wOups "frmProjSoumElecTemps", "InitialiserVariablesConfig", Err, Err.number, Err.Description
End Sub

Private Sub InitialiserVariablesProjSoum()

 On Error GoTo Oups

 m_sTauxDessin = FrmProjSoumElec.m_sTauxDessin
 m_sTauxFabrication = FrmProjSoumElec.m_sTauxFabrication
 m_sTauxAssemblage = FrmProjSoumElec.m_sTauxAssemblage
 m_sTauxProgInterface = FrmProjSoumElec.m_sTauxProgInterface
 m_sTauxProgAutomate = FrmProjSoumElec.m_sTauxProgAutomate
 m_sTauxProgRobot = FrmProjSoumElec.m_sTauxProgRobot
 m_sTauxVision = FrmProjSoumElec.m_sTauxVision
 m_sTauxTest = FrmProjSoumElec.m_sTauxTest
 m_sTauxInstallation = FrmProjSoumElec.m_sTauxInstallation
 m_sTauxMiseService = FrmProjSoumElec.m_sTauxMiseService
  m_sTauxFormation = FrmProjSoumElec.m_sTauxFormation
  m_sTauxGestion = FrmProjSoumElec.m_sTauxGestion
  m_sTauxShipping = FrmProjSoumElec.m_sTauxShipping

  m_sRepas = FrmProjSoumElec.m_sTauxRepas
  m_sHebergement1 = FrmProjSoumElec.m_sTauxHebergement1
  m_sHebergement2 = FrmProjSoumElec.m_sTauxHebergement2
  m_sStandard = FrmProjSoumElec.m_sTauxTransport
  m_sUniteMobile = FrmProjSoumElec.m_sTauxUniteMobile
 
10 Exit Sub

Oups:

wOups "frmProjSoumElecTemps", "InitialiserVariablesProjSoum", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()
 
 On Error GoTo Oups
 
 If FrmProjSoumElec.m_bDroitPrix = False Then
 Me.width = 3765
 Cmdfermer.Left = 2280
 Else
 Me.width = 11115
 Cmdfermer.Left = 9480
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub txtNbrePersonne_Change()

 On Error GoTo Oups

 If txtTempsHebergement.Text <> vbNullString Then
 If IsNumeric(txtNbrePersonne.Text) Then
 Call CalculerHebergement
 Call CalculerRepas

 Call CalculerTotal
 End If
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtNbrePersonne_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtPrixEmballage_KeyPress(KeyAscii As Integer)

 On Error GoTo Oups

 If KeyAscii = 4 Then  'Si c'est le "."
 KeyAscii = 44 'Remplace par la ","
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "lblPrixEmballage_KeyPress", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsHebergement_Change()

 On Error GoTo Oups

 If IsNumeric(txtTempsHebergement.Text) Then
 If txtNbrePersonne.Text <> vbNullString Then
 Call CalculerHebergement
 End If
 End If
 
 Call CalculerTotal

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsHebergement_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsHebergement_LostFocus()

 On Error GoTo Oups

 txtTempsHebergement.Text = Replace(txtTempsHebergement.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsHebergement_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsRepas_Change()

 On Error GoTo Oups

 If IsNumeric(txtTempsRepas.Text) Then
 If txtNbrePersonne.Text <> vbNullString Then
 Call CalculerRepas
 End If
 End If
 
 Call CalculerTotal

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsRepas_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsRepas_LostFocus()

 On Error GoTo Oups

 txtTempsRepas.Text = Replace(txtTempsRepas.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsRepas_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsDeplacement_Change()

 On Error GoTo Oups

 If IsNumeric(txtTempsDeplacement.Text) Then
 lblPrixDeplacement.Caption = Round(Replace(txtTempsDeplacement.Text * m_sStandard, ".", ","), 2)
 Else
 lblPrixDeplacement.Caption = 0
 End If
 
 Call CalculerTotal

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsDeplacement_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsDeplacement_LostFocus()

 On Error GoTo Oups

 txtTempsDeplacement.Text = Replace(txtTempsDeplacement.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsDeplacement_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsUniteMobile_Change()

 On Error GoTo Oups

 If IsNumeric(txtTempsUniteMobile.Text) Then
 lblPrixUniteMobile.Caption = Round(Replace(txtTempsUniteMobile.Text * m_sUniteMobile, ".", ","), 2)
 Else
 lblPrixUniteMobile.Caption = 0
 End If
 
 Call CalculerTotal

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsUniteMobile_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsUniteMobile_LostFocus()

 On Error GoTo Oups

 txtTempsUniteMobile.Text = Replace(txtTempsUniteMobile.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsUniteMobile_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtPrixEmballage_Change()

 On Error GoTo Oups
 
 Call CalculerTotal

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "lblPrixEmballage_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtPrixEmballage_LostFocus()

 On Error GoTo Oups

 If IsNumeric(txtPrixEmballage.Text) Then
 txtPrixEmballage.Text = Round(Replace(txtPrixEmballage.Text, ".", ","), 2)
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "lblPrixEmballage_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsDessinSoum_Change()

 On Error GoTo Oups

 If m_eType = TYPE_SOUMISSION Then
 If IsNumeric(txtTempsDessinSoum.Text) Then
 lblPrixDessin.Caption = Round(Replace(txtTempsDessinSoum.Text * m_sTauxDessin, ".", ","), 2)
 Else
 lblPrixDessin.Caption = 0
 End If
 
 Call CalculerTotal
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsDessinSoum_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsDessinSoum_LostFocus()

 On Error GoTo Oups

 txtTempsDessinSoum.Text = Replace(txtTempsDessinSoum.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsDessinSoum_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub lblTempsFabricationSoum_Change()

 On Error GoTo Oups

 If m_eType = TYPE_SOUMISSION Then
 If m_bSansTemps = False Then
 If IsNumeric(lblTempsFabricationSoum.Caption) Then
 lblPrixFabrication.Caption = Round(Replace(lblTempsFabricationSoum.Caption * m_sTauxFabrication, ".", ","), 2)
 Else
 lblPrixFabrication.Caption = "0"
 End If
 Else
 lblPrixFabrication.Caption = "0"
 End If
 
  Call CalculerTotal
  End If

  Exit Sub

Oups:

  wOups "frmProjSoumElecTemps", "txtTempsMécanique_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsAssemblageSoum_Change()

 On Error GoTo Oups

 If m_eType = TYPE_SOUMISSION Then
 If IsNumeric(txtTempsAssemblageSoum.Text) Then
 lblPrixAssemblage.Caption = Round(Replace(txtTempsAssemblageSoum.Text * m_sTauxAssemblage, ".", ","), 2)
 Else
 lblPrixAssemblage.Caption = 0
 End If
 
 Call CalculerTotal
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsAssemblageSoum_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsAssemblageSoum_LostFocus()

 On Error GoTo Oups

 txtTempsAssemblageSoum.Text = Replace(txtTempsAssemblageSoum.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsAssemblageSoum_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsProgInterfaceSoum_Change()

 On Error GoTo Oups

 If m_eType = TYPE_SOUMISSION Then
 If IsNumeric(txtTempsProgInterfaceSoum.Text) Then
 lblPrixProgInterface.Caption = Round(Replace(txtTempsProgInterfaceSoum.Text * m_sTauxProgInterface, ".", ","), 2)
 Else
 lblPrixProgInterface.Caption = 0
 End If
 
 Call CalculerTotal
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsProgInterfaceSoum_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsProgInterfaceSoum_LostFocus()

 On Error GoTo Oups

 txtTempsProgInterfaceSoum.Text = Replace(txtTempsProgInterfaceSoum.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsProgInterfaceSoum_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsProgAutomateSoum_Change()

 On Error GoTo Oups

 If m_eType = TYPE_SOUMISSION Then
 If IsNumeric(txtTempsProgAutomateSoum.Text) Then
 lblPrixProgAutomate.Caption = Round(Replace(txtTempsProgAutomateSoum.Text * m_sTauxProgAutomate, ".", ","), 2)
 Else
 lblPrixProgAutomate.Caption = 0
 End If
 
 Call CalculerTotal
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsProgAutomate_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsProgAutomateSoum_LostFocus()

 On Error GoTo Oups

 txtTempsProgAutomateSoum.Text = Replace(txtTempsProgAutomateSoum.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsProgAutomate_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsProgRobotSoum_Change()

 On Error GoTo Oups

 If m_eType = TYPE_SOUMISSION Then
 If IsNumeric(txtTempsProgRobotSoum.Text) Then
 lblPrixProgRobot.Caption = Round(Replace(txtTempsProgRobotSoum.Text * m_sTauxProgRobot, ".", ","), 2)
 Else
 lblPrixProgRobot.Caption = 0
 End If
 
 Call CalculerTotal
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsProgRobotSoum_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsProgRobotSoum_LostFocus()

 On Error GoTo Oups

 txtTempsProgRobotSoum.Text = Replace(txtTempsProgRobotSoum.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsProgRobotSoum_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsVisionSoum_Change()

 On Error GoTo Oups

 If m_eType = TYPE_SOUMISSION Then
 If IsNumeric(txtTempsVisionSoum.Text) Then
 lblPrixVision.Caption = Round(Replace(txtTempsVisionSoum.Text * m_sTauxVision, ".", ","), 2)
 Else
 lblPrixVision.Caption = 0
 End If
 
 Call CalculerTotal
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsVisionSoum_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsVisionSoum_LostFocus()

 On Error GoTo Oups

 txtTempsVisionSoum.Text = Replace(txtTempsVisionSoum.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsVisionSoum_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsTestSoum_Change()

 On Error GoTo Oups

 If m_eType = TYPE_SOUMISSION Then
 If IsNumeric(txtTempsTestSoum.Text) Then
 lblPrixTest.Caption = Round(Replace(txtTempsTestSoum.Text * m_sTauxTest, ".", ","), 2)
 Else
 lblPrixTest.Caption = 0
 End If
 
 Call CalculerTotal
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsTestSoum_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsTestSoum_LostFocus()

 On Error GoTo Oups

 txtTempsTestSoum.Text = Replace(txtTempsTestSoum.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsTestSoum_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsInstallationSoum_Change()

 On Error GoTo Oups

 If m_eType = TYPE_SOUMISSION Then
 If IsNumeric(txtTempsInstallationSoum.Text) Then
 lblPrixInstallation.Caption = Round(Replace(txtTempsInstallationSoum.Text * m_sTauxInstallation, ".", ","), 2)
 Else
 lblPrixInstallation.Caption = 0
 End If
 
 Call CalculerTotal
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsInstallationSoum_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsInstallationSoum_LostFocus()

 On Error GoTo Oups

 txtTempsInstallationSoum.Text = Replace(txtTempsInstallationSoum.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsInstallationSoum_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsMiseServiceSoum_Change()

 On Error GoTo Oups

 If m_eType = TYPE_SOUMISSION Then
 If IsNumeric(txtTempsMiseServiceSoum.Text) Then
 lblPrixMiseService.Caption = Round(Replace(txtTempsMiseServiceSoum.Text * m_sTauxMiseService, ".", ","), 2)
 Else
 lblPrixMiseService.Caption = 0
 End If
 
 Call CalculerTotal
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsMiseServiceSoum_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsMiseServiceSoum_LostFocus()

 On Error GoTo Oups

 txtTempsMiseServiceSoum.Text = Replace(txtTempsMiseServiceSoum.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsMiseServiceSoum_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsFormationSoum_Change()

 On Error GoTo Oups

 If m_eType = TYPE_SOUMISSION Then
 If IsNumeric(txtTempsFormationSoum.Text) Then
 lblPrixFormation.Caption = Round(Replace(txtTempsFormationSoum.Text * m_sTauxFormation, ".", ","), 2)
 Else
 lblPrixFormation.Caption = 0
 End If
 
 Call CalculerTotal
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsFormationSoum_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsFormationSoum_LostFocus()

 On Error GoTo Oups

 txtTempsFormationSoum.Text = Replace(txtTempsFormationSoum.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsFormationSoum_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsGestionSoum_Change()

 On Error GoTo Oups

 If m_eType = TYPE_SOUMISSION Then
 If IsNumeric(txtTempsGestionSoum.Text) Then
 lblPrixGestion.Caption = Round(Replace(txtTempsGestionSoum.Text * m_sTauxGestion, ".", ","), 2)
 Else
 lblPrixGestion.Caption = 0
 End If
 
 Call CalculerTotal
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsGestionSoum_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsGestionSoum_LostFocus()

 On Error GoTo Oups

 txtTempsGestionSoum.Text = Replace(txtTempsGestionSoum.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsGestionSoum_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsShippingSoum_Change()

 On Error GoTo Oups

 If m_eType = TYPE_SOUMISSION Then
 If IsNumeric(txtTempsShippingSoum.Text) Then
 lblPrixShipping.Caption = Round(Replace(txtTempsShippingSoum.Text * m_sTauxShipping, ".", ","), 2)
 Else
 lblPrixShipping.Caption = 0
 End If
 
 Call CalculerTotal
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsShippingSoum_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsShippingSoum_LostFocus()

 On Error GoTo Oups

 txtTempsShippingSoum.Text = Replace(txtTempsShippingSoum.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumElecTemps", "txtTempsShippingSoum_LostFocus", Err, Err.number, Err.Description
End Sub


Private Sub CalculerHebergement()

 On Error GoTo Oups

 Dim dblNbreDeux As Double
 Dim dblHebergement As Double
 Dim iReste As Integer
 Dim dblNbrePers As Double
 Dim dblNbreJours As Double
 
 If IsNumeric(txtNbrePersonne.Text) Then
 dblNbrePers = CDbl(txtNbrePersonne.Text)
 Else
 dblNbrePers = 0
 End If
 
  If IsNumeric(txtTempsHebergement.Text) Then
  dblNbreJours = CDbl(txtTempsHebergement.Text)
  Else
  dblNbreJours = 0
  End If
 
  dblNbreDeux = Int(dblNbrePers / 2)
 
  iReste = CInt(dblNbrePers) - (dblNbreDeux * 2)
 
  dblHebergement = dblNbreJours * ((dblNbreDeux * CDbl(m_sHebergement2)) + (iReste * CDbl(m_sHebergement1)))
 
10 lblPrixHebergement.Caption = Round(Replace(dblHebergement, ".", ","), 2)

Exit Sub

Oups:

wOups "frmProjSoumElecTemps", "CalculerHebergement", Err, Err.number, Err.Description
End Sub

Private Sub CalculerRepas()

 On Error GoTo Oups

 Dim dblNbrePers As Double
 Dim dblRepas As Double
 Dim dblNbreJours As Double

 If IsNumeric(txtNbrePersonne.Text) Then
 dblNbrePers = CDbl(txtNbrePersonne.Text)
 Else
 dblNbrePers = 0
 End If
 
 If IsNumeric(txtTempsRepas.Text) Then
 dblNbreJours = CDbl(txtTempsRepas.Text)
  Else
  dblNbreJours = 0
  End If
 
  dblRepas = dblNbreJours * dblNbrePers * CDbl(m_sRepas)
 
  lblPrixRepas.Caption = Round(Replace(dblRepas, ".", ","), 2)

  Exit Sub

Oups:

  wOups "frmProjSoumElecTemps", "CalculerRepas", Err, Err.number, Err.Description
End Sub

Private Sub CalculerTotal()

 On Error GoTo Oups

 Dim dblTotal As Double
 Dim dblPrixEmballage As Double
 Dim dblTotalArgentRH As Double
 Dim dblPrixDessin As Double
 Dim dblPrixFabrication As Double
 Dim dblPrixAssemblage As Double
 Dim dblPrixProgInterface As Double
 Dim dblPrixProgAutomate As Double
 Dim dblPrixProgRobot As Double
 Dim dblPrixVision As Double
  Dim dblPrixTest As Double
  Dim dblPrixInstallation As Double
  Dim dblPrixMiseService As Double
  Dim dblPrixFormation As Double
  Dim dblPrixGestion As Double
  Dim dblPrixShipping As Double
 Dim dblPrixPrototype As Double
  Dim dblPrixHebergement As Double
  Dim dblPrixRepas As Double
10 Dim dblPrixDeplacement As Double
Dim dblPrixUniteMobile As Double
 
 'Prix de dessin
If IsNumeric(lblPrixDessin.Caption) Then
 dblPrixDessin = CDbl(lblPrixDessin.Caption)
Else
 dblPrixDessin = 0
End If

 'Prix de Fabrication
If IsNumeric(lblPrixFabrication.Caption) Then
 dblPrixFabrication = CDbl(lblPrixFabrication.Caption)
Else
 dblPrixFabrication = 0
End If

 'Prix de Assemblage
1  If IsNumeric(lblPrixAssemblage.Caption) Then
 dblPrixAssemblage = CDbl(lblPrixAssemblage.Caption)
 Else
 dblPrixAssemblage = 0
 End If

 'Prix de ProgInterface
If IsNumeric(lblPrixProgInterface.Caption) Then
 dblPrixProgInterface = CDbl(lblPrixProgInterface.Caption)
1  Else
 dblPrixProgInterface = 0
 End If

 'Prix de ProgAutomate
If IsNumeric(lblPrixProgAutomate.Caption) Then
 dblPrixProgAutomate = CDbl(lblPrixProgAutomate.Caption)
Else
 dblPrixProgAutomate = 0
End If

 'Prix de ProgRobot
If IsNumeric(lblPrixProgRobot.Caption) Then
 dblPrixProgRobot = CDbl(lblPrixProgRobot.Caption)
Else
 dblPrixProgRobot = 0
End If

 'Prix de vision
2  If IsNumeric(lblPrixVision.Caption) Then
 dblPrixVision = CDbl(lblPrixVision.Caption)
2  Else
 dblPrixVision = 0
2  End If

 'Prix de test
If IsNumeric(lblPrixTest.Caption) Then
dblPrixTest = CDbl(lblPrixTest.Caption)
Else
dblPrixTest = 0
End If

 'Prix de Installation
If IsNumeric(lblPrixInstallation.Caption) Then
 dblPrixInstallation = CDbl(lblPrixInstallation.Caption)
Else
 dblPrixInstallation = 0
End If

 'Prix de MiseService
If IsNumeric(lblPrixMiseService.Caption) Then
 dblPrixMiseService = CDbl(lblPrixMiseService.Caption)
Else
 dblPrixMiseService = 0
End If

 'Prix de formation
3  If IsNumeric(lblPrixFormation.Caption) Then
 dblPrixFormation = CDbl(lblPrixFormation.Caption)
3  Else
 dblPrixFormation = 0
3  End If

 'Prix de Gestion
If IsNumeric(lblPrixGestion.Caption) Then
 dblPrixGestion = CDbl(lblPrixGestion.Caption)
 Else
dblPrixGestion = 0
End If

 'Prix de Shipping
4 If IsNumeric(lblPrixShipping.Caption) Then
4 dblPrixShipping = CDbl(lblPrixShipping.Caption)
4 Else
4 dblPrixShipping = 0
4 End If


 'Prix de Prototype
43 If IsNumeric(lblPrixPrototype.Caption) Then
4 dblPrixPrototype = CDbl(lblPrixPrototype.Caption)
43 Else
434 dblPrixPrototype = 0
4 End If



 'Prix d'hébergement
4 If IsNumeric(lblPrixHebergement.Caption) Then
4 dblPrixHebergement = CDbl(lblPrixHebergement.Caption)
4 Else
4 dblPrixHebergement = 0
4 End If

 'Prix des repas
4  If IsNumeric(lblPrixRepas.Caption) Then
4  dblPrixRepas = CDbl(lblPrixRepas.Caption)
4  Else
4  dblPrixRepas = 0
4  End If
 
 'Prix du déplacement
4  If IsNumeric(lblPrixDeplacement.Caption) Then
4  dblPrixDeplacement = CDbl(lblPrixDeplacement.Caption)
4  Else
50 dblPrixDeplacement = 0
50 End If

 'Prix de l'unité mobile
 If IsNumeric(lblPrixUniteMobile.Caption) Then
 dblPrixUniteMobile = CDbl(lblPrixUniteMobile.Caption)
 Else
 dblPrixUniteMobile = 0
 End If
 
 'Prix de transport et emballage
 If IsNumeric(txtPrixEmballage.Text) Then
 dblPrixEmballage = CDbl(txtPrixEmballage.Text)
 Else
 dblPrixEmballage = 0
 End If

5  dblTotalArgentRH = dblPrixDessin + _
 dblPrixFabrication + _
 dblPrixAssemblage + _
 dblPrixProgInterface + _
 dblPrixProgAutomate + _
 dblPrixProgRobot + _
 dblPrixVision + _
 dblPrixTest + _
 dblPrixInstallation + _
 dblPrixMiseService + _
 dblPrixFormation + _
 dblPrixGestion + _
 dblPrixShipping + _
 dblPrixPrototype
 
5  lblTotalPrixRH.Caption = Conversion(CStr(dblTotalArgentRH), MODE_DECIMAL)

5  dblTotal = dblTotalArgentRH + _
 dblPrixHebergement + _
 dblPrixRepas + _
 dblPrixDeplacement + _
 dblPrixUniteMobile + _
 dblPrixEmballage

5  lblTotal.Caption = Conversion(CStr(dblTotal), MODE_DECIMAL)

5  Call CalculerTotalTemps

5  Exit Sub

Oups:

5  wOups "frmProjSoumElecTemps", "CalculerTotal", Err, Err.number, Err.Description
End Sub

Private Sub CalculerTotalTemps()

 
 On Error GoTo Oups

 Dim dblTempsDessin As Double
 Dim dblTempsFabrication As Double
 Dim dblTempsAssemblage As Double
 Dim dblTempsProgInterface As Double
 Dim dblTempsProgAutomate As Double
 Dim dblTempsProgRobot As Double
 Dim dblTempsVision As Double
 Dim dblTempsTest As Double
 Dim dblTempsInstallation As Double
 Dim dblTempsMiseService As Double
  Dim dblTempsFormation As Double
  Dim dblTempsGestion As Double
  Dim dblTempsShipping As Double
 Dim dblTempsPrototype As Double
  Dim dblTotalTemps As Double

 'SOUMISSION
  If IsNumeric(txtTempsDessinSoum.Text) Then
  dblTempsDessin = CDbl(txtTempsDessinSoum.Text)
  Else
  dblTempsDessin = 0
10 End If

If m_bSansTemps = False Then
 If IsNumeric(lblTempsFabricationSoum.Caption) Then
 dblTempsFabrication = CDbl(lblTempsFabricationSoum.Caption)
 Else
 dblTempsFabrication = 0
 End If
Else
 dblTempsFabrication = 0
End If

If IsNumeric(txtTempsAssemblageSoum.Text) Then
 dblTempsAssemblage = CDbl(txtTempsAssemblageSoum.Text)
1  Else
 dblTempsAssemblage = 0
 End If

If IsNumeric(txtTempsProgInterfaceSoum.Text) Then
 dblTempsProgInterface = CDbl(txtTempsProgInterfaceSoum.Text)
Else
 dblTempsProgInterface = 0
1  End If

 If IsNumeric(txtTempsProgAutomateSoum.Text) Then
 dblTempsProgAutomate = CDbl(txtTempsProgAutomateSoum.Text)
Else
 dblTempsProgAutomate = 0
End If

If IsNumeric(txtTempsProgRobotSoum.Text) Then
 dblTempsProgRobot = CDbl(txtTempsProgRobotSoum.Text)
Else
 dblTempsProgRobot = 0
End If

If IsNumeric(txtTempsVisionSoum.Text) Then
 dblTempsVision = CDbl(txtTempsVisionSoum.Text)
2  Else
 dblTempsVision = 0
2  End If

If IsNumeric(txtTempsTestSoum.Text) Then
dblTempsTest = CDbl(txtTempsTestSoum.Text)
Else
dblTempsTest = 0
End If

30 If IsNumeric(txtTempsInstallationSoum.Text) Then
3 dblTempsInstallation = CDbl(txtTempsInstallationSoum.Text)
Else
 dblTempsInstallation = 0
End If

If IsNumeric(txtTempsMiseServiceSoum.Text) Then
 dblTempsMiseService = CDbl(txtTempsMiseServiceSoum.Text)
Else
 dblTempsMiseService = 0
End If

If IsNumeric(txtTempsFormationSoum.Text) Then
 dblTempsFormation = CDbl(txtTempsFormationSoum.Text)
3  Else
 dblTempsFormation = 0
3  End If

If IsNumeric(txtTempsGestionSoum.Text) Then
dblTempsGestion = CDbl(txtTempsGestionSoum.Text)
Else
 dblTempsGestion = 0
 End If

40 If IsNumeric(txtTempsShippingSoum.Text) Then
4 dblTempsShipping = CDbl(txtTempsShippingSoum.Text)
4 Else
4 dblTempsShipping = 0
4 End If


 If IsNumeric(txtTempsprototypeSoum.Text) Then
4 dblTempsPrototype = CDbl(txtTempsprototypeSoum.Text)
42 Else
42 dblTempsPrototype = 0
424 End If

4 dblTotalTemps = dblTempsDessin + _
 dblTempsFabrication + _
 dblTempsAssemblage + _
 dblTempsProgInterface + _
 dblTempsProgAutomate + _
 dblTempsProgRobot + _
 dblTempsVision + _
 dblTempsTest + _
 dblTempsInstallation + _
 dblTempsMiseService + _
 dblTempsFormation + _
 dblTempsGestion + _
 dblTempsShipping + _
 dblTempsPrototype

4 lblTotalTempsRHSoum.Caption = Conversion(dblTotalTemps, MODE_DECIMAL)

4 Exit Sub

Oups:

4 wOups "frmProjSoumElecTemps", "CalculerTotalTemps", Err, Err.number, Err.Description
End Sub

Private Sub lblTempsFabricationSoum_Click()
 'Active ou désactive le temps des pièces
 
 On Error GoTo Oups
 
 If m_eMode = MODE_AJOUT_MODIF Then
 If m_bSansTemps = True Then
 Croix1.Visible = False
 Croix2.Visible = False

 m_bSansTemps = False
 Else
 Croix1.Visible = True
 Croix2.Visible = True

 m_bSansTemps = True
 End If

  Call lblTempsFabricationSoum_Change
  End If

  Call CalculerTotal

  Exit Sub

Oups:

  wOups "frmProjSoumElecTemps", "lblTempsMécanique_Click", Err, Err.number, Err.Description
End Sub

Private Function CalculerTempsFabrication() As String

 On Error GoTo Oups

 Dim dblTempsFab As Double
 Dim iCompteur As Integer

 'Pour chaque élément du listView
 For iCompteur = 1 To FrmProjSoumElec.lvwSoumission.ListItems.count
 If Trim$(FrmProjSoumElec.lvwSoumission.ListItems(iCompteur).SubItems(9)) <> vbNullString Then
 'On additionne le temps
 dblTempsFab = dblTempsFab + CDbl(Replace(Trim$(FrmProjSoumElec.lvwSoumission.ListItems(iCompteur).SubItems(9)), ".", ","))
 End If
 Next
 
 CalculerTempsFabrication = Replace(dblTempsFab / 10, ".", ",")

 Exit Function

Oups:

 wOups "frmProjSoumElecTemps", "CalculerTempsFabrication", Err, Err.number, Err.Description
End Function

