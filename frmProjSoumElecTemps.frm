VERSION 5.00
Begin VB.Form frmProjSoumElecTemps 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Temps"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11025
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmProjSoumElecTemps.frx":0000
   ScaleHeight     =   8100
   ScaleWidth      =   11025
   StartUpPosition =   2  'CenterScreen
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
      BackColor       =   &H00000000&
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
      ForeColor       =   &H00FFFFFF&
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
      Caption         =   "Détails"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   7560
      Width           =   735
   End
   Begin VB.Label Label52 
      BackStyle       =   0  'Transparent
      Caption         =   "Total de la ressource humaine :"
      ForeColor       =   &H00FFFFFF&
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
      BackStyle       =   0  'Transparent
      Caption         =   "Hrs"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   100
      Top             =   7080
      Width           =   255
   End
   Begin VB.Label lblDollarRH 
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5640
      TabIndex        =   99
      Top             =   7080
      Width           =   135
   End
   Begin VB.Label lblTotalTempsRHSoum 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2640
      TabIndex        =   98
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label lblTotalTempsRHReel 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3480
      TabIndex        =   97
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label lblTotalPrixRH 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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

Private Enum enumType
  TYPE_PROJET = 0
  TYPE_SOUMISSION = 1
End Enum

Private Enum enumMode
  MODE_AJOUT_MODIF = 0
  MODE_INACTIF = 1
End Enum

Private m_sTauxDessin        As String
Private m_sTauxFabrication   As String
Private m_sTauxAssemblage    As String
Private m_sTauxProgInterface As String
Private m_sTauxProgAutomate  As String
Private m_sTauxProgRobot     As String
Private m_sTauxVision        As String
Private m_sTauxTest          As String
Private m_sTauxInstallation  As String
Private m_sTauxMiseService   As String
Private m_sTauxFormation     As String
Private m_sTauxGestion       As String
Private m_sTauxShipping      As String

Private m_sRepas             As String
Private m_sHebergement1      As String
Private m_sHebergement2      As String
Private m_sStandard          As String
Private m_sUniteMobile       As String

Private m_sNoProjSoum        As String

Private m_eType              As enumType

Private m_eMode              As enumMode
 
Private m_bNouveauTaux       As Boolean 'Pour savoir si les nouveaux taux doivent être prit

Private m_bSansTemps         As Boolean

Public Sub Afficher(ByVal sNoProjSoum As String, ByVal iType As Integer, ByVal iMode As Integer, ByVal bNouveauTaux As Boolean)

5       On Error GoTo AfficherErreur
  
10      m_eType = iType
    
15      m_eMode = iMode
    
20      m_sNoProjSoum = sNoProjSoum
  
25      m_bNouveauTaux = bNouveauTaux
  
30      If bNouveauTaux = True Then
35        Call InitialiserVariablesConfig
40      Else
45        Call InitialiserVariablesProjSoum
50      End If
    
55      If m_eMode = MODE_AJOUT_MODIF Then
60        Call BarrerChamps(False)
65      Else
70        Call BarrerChamps(True)
75      End If
     
80      Call AfficherEnregistrement
  
85      Call Me.Show(vbModal)

90      Exit Sub

AfficherErreur:

95      woups "frmProjSoumElecTemps", "Afficher", Err, Erl
End Sub

Private Sub AfficherEnregistrement()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset
15      Dim rstSoum     As ADODB.Recordset
20      Dim sChamps     As String
25      Dim sTable      As String
  
30      If m_eType = TYPE_PROJET Then
35        sChamps = "IDProjet"
40        sTable = "GRB_ProjetElec"
45      Else
50        sChamps = "IDSoumission"
55        sTable = "GRB_SoumissionElec"
60      End If
  
65      Set rstProjSoum = New ADODB.Recordset
  
70      Call rstProjSoum.Open("SELECT * FROM " & sTable & " WHERE " & sChamps & " = '" & m_sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)
  
75      If Not rstProjSoum.EOF And FrmProjSoumElec.m_bTempsDejaOuvert = False And m_eMode = MODE_INACTIF Then
80        m_bSansTemps = rstProjSoum.Fields("SansTemps")

85        If Not IsNull(rstProjSoum.Fields("TempsDessin")) Then
90          txtTempsDessinSoum.Text = rstProjSoum.Fields("TempsDessin")
95        Else
100         txtTempsDessinSoum.Text = "0"
105       End If

110       If m_eType = TYPE_SOUMISSION Then
115         If Not IsNull(rstProjSoum.Fields("TempsFabrication")) Then
120           lblTempsFabricationSoum.Caption = rstProjSoum.Fields("TempsFabrication")
125         Else
130           lblTempsFabricationSoum.Caption = "0"
135         End If
140       Else
145         lblTempsFabricationSoum.Caption = CalculerTempsFabrication
150       End If

155       If Not IsNull(rstProjSoum.Fields("TempsAssemblage")) Then
160         txtTempsAssemblageSoum.Text = rstProjSoum.Fields("TempsAssemblage")
165       Else
170         txtTempsAssemblageSoum.Text = "0"
175       End If

180       If Not IsNull(rstProjSoum.Fields("TempsProgInterface")) Then
185         txtTempsProgInterfaceSoum.Text = rstProjSoum.Fields("TempsProgInterface")
190       Else
195         txtTempsProgInterfaceSoum.Text = "0"
200       End If

205       If Not IsNull(rstProjSoum.Fields("TempsProgAutomate")) Then
210         txtTempsProgAutomateSoum.Text = rstProjSoum.Fields("TempsProgAutomate")
215       Else
220         txtTempsProgAutomateSoum.Text = "0"
225       End If

230       If Not IsNull(rstProjSoum.Fields("TempsProgRobot")) Then
235         txtTempsProgRobotSoum.Text = rstProjSoum.Fields("TempsProgRobot")
240       Else
245         txtTempsProgRobotSoum.Text = "0"
250       End If

255       If Not IsNull(rstProjSoum.Fields("TempsVision")) Then
260         txtTempsVisionSoum.Text = rstProjSoum.Fields("TempsVision")
265       Else
270         txtTempsVisionSoum.Text = "0"
275       End If

280       If Not IsNull(rstProjSoum.Fields("TempsTest")) Then
285         txtTempsTestSoum.Text = rstProjSoum.Fields("TempsTest")
290       Else
295         txtTempsTestSoum.Text = "0"
300       End If

305       If Not IsNull(rstProjSoum.Fields("TempsInstallation")) Then
310         txtTempsInstallationSoum.Text = rstProjSoum.Fields("TempsInstallation")
315       Else
320         txtTempsInstallationSoum.Text = "0"
325       End If

330       If Not IsNull(rstProjSoum.Fields("TempsMiseService")) Then
335         txtTempsMiseServiceSoum.Text = rstProjSoum.Fields("TempsMiseService")
340       Else
345         txtTempsMiseServiceSoum.Text = "0"
350       End If

355       If Not IsNull(rstProjSoum.Fields("TempsFormation")) Then
360         txtTempsFormationSoum.Text = rstProjSoum.Fields("TempsFormation")
365       Else
370         txtTempsFormationSoum.Text = "0"
375       End If

380       If Not IsNull(rstProjSoum.Fields("TempsGestion")) Then
385         txtTempsGestionSoum.Text = rstProjSoum.Fields("TempsGestion")
390       Else
395         txtTempsGestionSoum.Text = "0"
400       End If

405       If Not IsNull(rstProjSoum.Fields("TempsShipping")) Then
410         txtTempsShippingSoum.Text = rstProjSoum.Fields("TempsShipping")
415       Else
420         txtTempsShippingSoum.Text = "0"
425       End If
            txtTempsprototypeSoum.Text = "0"

430       If m_bSansTemps = True Then
435         Croix1.Visible = True
440         Croix2.Visible = True
445       Else
450         Croix1.Visible = False
455         Croix2.Visible = False
460       End If

465       If Right$(m_sNoProjSoum, 2) <> "99" Then
470         If Not IsNull(rstProjSoum.Fields("NbrePersonne")) Then
475           txtNbrePersonne.Text = rstProjSoum.Fields("NbrePersonne")
480         Else
485           txtNbrePersonne.Text = "0"
490         End If

495         If Not IsNull(rstProjSoum.Fields("TempsHebergement")) Then
500           txtTempsHebergement.Text = rstProjSoum.Fields("TempsHebergement")
505         Else
510           txtTempsHebergement.Text = "0"
515         End If

520         If Not IsNull(rstProjSoum.Fields("TempsRepas")) Then
525           txtTempsRepas.Text = rstProjSoum.Fields("TempsRepas")
530         Else
535           txtTempsRepas.Text = "0"
540         End If
545       Else
550         If Not IsNull(rstProjSoum.Fields("TotalHebergement")) Then
555           lblPrixHebergement.Caption = rstProjSoum.Fields("TotalHebergement")
560         End If

565         If Not IsNull(rstProjSoum.Fields("TotalRepas")) Then
570           lblPrixRepas.Caption = rstProjSoum.Fields("TotalRepas")
575         End If
580       End If

585       If Not IsNull(rstProjSoum.Fields("TempsTransport")) Then
590         txtTempsDeplacement.Text = rstProjSoum.Fields("TempsTransport")
595       Else
600         txtTempsDeplacement.Text = "0"
605       End If

610       If Not IsNull(rstProjSoum.Fields("TempsUniteMobile")) Then
615         txtTempsUniteMobile.Text = rstProjSoum.Fields("TempsUniteMobile")
620       Else
625         txtTempsUniteMobile.Text = "0"
630       End If

635       If Not IsNull(rstProjSoum.Fields("PrixEmballage")) Then
640         txtPrixEmballage.Text = rstProjSoum.Fields("PrixEmballage")
645       Else
650         txtPrixEmballage.Text = "0"
655       End If
660     Else
665       If m_eType = TYPE_SOUMISSION Then
670         m_bSansTemps = FrmProjSoumElec.m_bSansTemps

675         txtTempsDessinSoum.Text = FrmProjSoumElec.m_sTempsDessin
680         lblTempsFabricationSoum.Caption = FrmProjSoumElec.m_sTempsFabrication
685         txtTempsAssemblageSoum.Text = FrmProjSoumElec.m_sTempsAssemblage
690         txtTempsProgInterfaceSoum.Text = FrmProjSoumElec.m_sTempsProgInterface
695         txtTempsProgAutomateSoum.Text = FrmProjSoumElec.m_sTempsProgAutomate
700         txtTempsProgRobotSoum.Text = FrmProjSoumElec.m_sTempsProgRobot
705         txtTempsVisionSoum.Text = FrmProjSoumElec.m_sTempsVision
710         txtTempsTestSoum.Text = FrmProjSoumElec.m_sTempsTest
715         txtTempsInstallationSoum.Text = FrmProjSoumElec.m_sTempsInstallation
720         txtTempsMiseServiceSoum.Text = FrmProjSoumElec.m_sTempsMiseService
725         txtTempsFormationSoum.Text = FrmProjSoumElec.m_sTempsFormation
730         txtTempsGestionSoum.Text = FrmProjSoumElec.m_sTempsGestion
735         txtTempsShippingSoum.Text = FrmProjSoumElec.m_sTempsShipping
740       Else
745         If Not rstProjSoum.EOF And Not IsNull(rstProjSoum.Fields("IDSoumission")) Then
750           If rstProjSoum.Fields("IDSoumission") <> "" Then
755             Set rstSoum = New ADODB.Recordset

760             Call rstSoum.Open("SELECT * FROM GRB_SoumissionElec WHERE IDSoumission = '" & rstProjSoum.Fields("IDSoumission") & "'", g_connData, adOpenDynamic, adLockOptimistic)

765             If Not rstSoum.EOF Then
770               m_bSansTemps = FrmProjSoumElec.m_bSansTemps

775               If Not IsNull(rstSoum.Fields("TempsDessin")) Then
780                 txtTempsDessinSoum.Text = rstSoum.Fields("TempsDessin")
785               Else
790                 txtTempsDessinSoum.Text = "0"
795               End If

800               If m_eType = TYPE_PROJET Then
805                 lblTempsFabricationSoum.Caption = CalculerTempsFabrication
810               Else
815                 If Not IsNull(rstSoum.Fields("TempsFabrication")) Then
820                   lblTempsFabricationSoum.Caption = rstSoum.Fields("TempsFabrication")
825                 Else
830                   lblTempsFabricationSoum.Caption = "0"
835                 End If
840               End If

845               If Not IsNull(rstSoum.Fields("TempsAssemblage")) Then
850                 txtTempsAssemblageSoum.Text = rstSoum.Fields("TempsAssemblage")
855               Else
860                 txtTempsAssemblageSoum.Text = "0"
865               End If

870               If Not IsNull(rstSoum.Fields("TempsProgInterface")) Then
875                 txtTempsProgInterfaceSoum.Text = rstSoum.Fields("TempsProgInterface")
880               Else
885                 txtTempsProgInterfaceSoum.Text = "0"
890               End If

895               If Not IsNull(rstSoum.Fields("TempsProgAutomate")) Then
900                 txtTempsProgAutomateSoum.Text = rstSoum.Fields("TempsProgAutomate")
905               Else
910                 txtTempsProgAutomateSoum.Text = "0"
915               End If

920               If Not IsNull(rstSoum.Fields("TempsProgRobot")) Then
925                 txtTempsProgRobotSoum.Text = rstSoum.Fields("TempsProgRobot")
930               Else
935                 txtTempsProgRobotSoum.Text = "0"
940               End If

945               If Not IsNull(rstSoum.Fields("TempsVision")) Then
950                 txtTempsVisionSoum.Text = rstSoum.Fields("TempsVision")
955               Else
960                 txtTempsVisionSoum.Text = "0"
965               End If

970               If Not IsNull(rstSoum.Fields("TempsTest")) Then
975                 txtTempsTestSoum.Text = rstSoum.Fields("TempsTest")
980               Else
985                 txtTempsTestSoum.Text = "0"
990               End If

995               If Not IsNull(rstSoum.Fields("TempsInstallation")) Then
1000                txtTempsInstallationSoum.Text = rstSoum.Fields("TempsInstallation")
1005              Else
1010                txtTempsInstallationSoum.Text = "0"
1015              End If

1020              If Not IsNull(rstSoum.Fields("TempsMiseService")) Then
1025                txtTempsMiseServiceSoum.Text = rstSoum.Fields("TempsMiseService")
1030              Else
1035                txtTempsMiseServiceSoum.Text = "0"
1040              End If

1045              If Not IsNull(rstSoum.Fields("TempsFormation")) Then
1050                txtTempsFormationSoum.Text = rstSoum.Fields("TempsFormation")
1055              Else
1060                txtTempsFormationSoum.Text = "0"
1065              End If

1070              If Not IsNull(rstSoum.Fields("TempsGestion")) Then
1075                txtTempsGestionSoum.Text = rstSoum.Fields("TempsGestion")
1080              Else
1085                txtTempsGestionSoum.Text = "0"
1090              End If

1095              If Not IsNull(rstSoum.Fields("TempsShipping")) Then
1100                txtTempsShippingSoum.Text = rstSoum.Fields("TempsShipping")
1105              Else
1110                txtTempsShippingSoum.Text = "0"
1115              End If
1120            Else
1125              m_bSansTemps = False

1130              txtTempsDessinSoum.Text = "0"
1135              lblTempsFabricationSoum.Caption = CalculerTempsFabrication
1140              txtTempsAssemblageSoum.Text = "0"
1145              txtTempsProgInterfaceSoum.Text = "0"
1150              txtTempsProgAutomateSoum.Text = "0"
1155              txtTempsProgRobotSoum.Text = "0"
1160              txtTempsVisionSoum.Text = "0"
1165              txtTempsTestSoum.Text = "0"
1170              txtTempsInstallationSoum.Text = "0"
1175              txtTempsMiseServiceSoum.Text = "0"
1180              txtTempsFormationSoum.Text = "0"
1185              txtTempsGestionSoum.Text = "0"
1190              txtTempsShippingSoum.Text = "0"
1195            End If

1200            Call rstSoum.Close
1205            Set rstSoum = Nothing
1210          Else
1215            m_bSansTemps = False

1220            txtTempsDessinSoum.Text = "0"
1225            lblTempsFabricationSoum.Caption = CalculerTempsFabrication
1230            txtTempsAssemblageSoum.Text = "0"
1235            txtTempsProgInterfaceSoum.Text = "0"
1240            txtTempsProgAutomateSoum.Text = "0"
1245            txtTempsProgRobotSoum.Text = "0"
1250            txtTempsVisionSoum.Text = "0"
1255            txtTempsTestSoum.Text = "0"
1260            txtTempsInstallationSoum.Text = "0"
1265            txtTempsMiseServiceSoum.Text = "0"
1270            txtTempsFormationSoum.Text = "0"
1275            txtTempsGestionSoum.Text = "0"
1280            txtTempsShippingSoum.Text = "0"
1285          End If
1290        Else
1295          m_bSansTemps = False

1300          txtTempsDessinSoum.Text = "0"
1305          lblTempsFabricationSoum.Caption = CalculerTempsFabrication
1310          txtTempsAssemblageSoum.Text = "0"
1315          txtTempsProgInterfaceSoum.Text = "0"
1320          txtTempsProgAutomateSoum.Text = "0"
1325          txtTempsProgRobotSoum.Text = "0"
1330          txtTempsVisionSoum.Text = "0"
1335          txtTempsTestSoum.Text = "0"
1340          txtTempsInstallationSoum.Text = "0"
1345          txtTempsMiseServiceSoum.Text = "0"
1350          txtTempsFormationSoum.Text = "0"
1355          txtTempsGestionSoum.Text = "0"
1360          txtTempsShippingSoum.Text = "0"
1365        End If
1370      End If

1375      If m_bSansTemps = True Then
1380        Croix1.Visible = True
1385        Croix2.Visible = True
1390      Else
1395        Croix1.Visible = False
1400        Croix2.Visible = False
1405      End If

1410      txtNbrePersonne.Text = FrmProjSoumElec.m_sNbrePersonne
1415      txtTempsHebergement.Text = FrmProjSoumElec.m_sTempsHebergement
1420      txtTempsRepas.Text = FrmProjSoumElec.m_sTempsRepas
1425      txtTempsDeplacement.Text = FrmProjSoumElec.m_sTempsTransport
1430      txtTempsUniteMobile.Text = FrmProjSoumElec.m_sTempsUniteMobile
1435      txtPrixEmballage.Text = FrmProjSoumElec.m_sPrixEmballage
1440    End If

1445    If m_eType = TYPE_PROJET Then
1450      Call AfficherTempsReels

1455      Call CalculerTotalArgent
1460    End If

1465    Call rstProjSoum.Close
1470    Set rstProjSoum = Nothing

1475    Exit Sub

AfficherErreur:

1480    woups "frmProjSoumElecTemps", "AfficherEnregistrement", Err, Erl
End Sub

Private Sub AfficherTempsReels()
  
5       On Error GoTo AfficherErreur

10      Dim rstPunch        As ADODB.Recordset
15      Dim sDateDebut      As String
20      Dim sDateFin        As String
25      Dim sTotal          As String
30      Dim sFilterNoProjet As String

35      sDateDebut = "TIMESERIAL(Left(GRB_Punch.HeureDébut,2),RIGHT(GRB_Punch.HeureDébut,2),0)"

40      sDateFin = "TIMESERIAL(Left(GRB_Punch.HeureFin,2),RIGHT(GRB_Punch.HeureFin,2),0)"

45      sTotal = "(SUM(" & sDateFin & " - " & sDateDebut & ")* 24) As Total"

50      If Right$(m_sNoProjSoum, 2) = "99" Then
55        sFilterNoProjet = "LEFT(NoProjet, 6) = '" & Left$(m_sNoProjSoum, 6) & "'"
60      Else
65        sFilterNoProjet = "NoProjet = '" & m_sNoProjSoum & "'"
70      End If

75      Set rstPunch = New ADODB.Recordset

80      Call rstPunch.Open("SELECT Type, " & sTotal & " FROM GRB_Punch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null GROUP BY Type", g_connData, adOpenDynamic, adLockOptimistic)

85      lblTempsDessinReel.Caption = "0"
90      lblTempsFabricationReel.Caption = "0"
95      lblTempsAssemblageReel.Caption = "0"
100     lblTempsProgInterfaceReel.Caption = "0"
105     lblTempsProgAutomateReel.Caption = "0"
110     lblTempsProgRobotReel.Caption = "0"
115     lblTempsVisionReel.Caption = "0"
120     lblTempsTestReel.Caption = "0"
125     lblTempsInstallationReel.Caption = "0"
130     lblTempsMiseServiceReel.Caption = "0"
135     lblTempsFormationReel.Caption = "0"
140     lblTempsGestionReel.Caption = "0"
145     lblTempsShippingReel.Caption = "0"
        lblTempsPrototypeReel.Caption = "0"

150     Do While Not rstPunch.EOF
155       If Not IsNull(rstPunch.Fields("Total")) Then
160         Select Case rstPunch.Fields("Type")
              Case "Dessin":        lblTempsDessinReel.Caption = Round(rstPunch.Fields("Total"), 2)
165           Case "Fabrication":   lblTempsFabricationReel.Caption = Round(rstPunch.Fields("Total"), 2)
170           Case "Assemblage":    lblTempsAssemblageReel.Caption = Round(rstPunch.Fields("Total"), 2)
175           Case "ProgInterface": lblTempsProgInterfaceReel.Caption = Round(rstPunch.Fields("Total"), 2)
180           Case "ProgAutomate":  lblTempsProgAutomateReel.Caption = Round(rstPunch.Fields("Total"), 2)
185           Case "ProgRobot":     lblTempsProgRobotReel.Caption = Round(rstPunch.Fields("Total"), 2)
190           Case "Vision":        lblTempsVisionReel.Caption = Round(rstPunch.Fields("Total"), 2)
195           Case "Test":          lblTempsTestReel.Caption = Round(rstPunch.Fields("Total"), 2)
200           Case "Installation":  lblTempsInstallationReel.Caption = Round(rstPunch.Fields("Total"), 2)
205           Case "MiseService":   lblTempsMiseServiceReel.Caption = Round(rstPunch.Fields("Total"), 2)
210           Case "Formation":     lblTempsFormationReel.Caption = Round(rstPunch.Fields("Total"), 2)
215           Case "Gestion":       lblTempsGestionReel.Caption = Round(rstPunch.Fields("Total"), 2)
220           Case "Shipping":      lblTempsShippingReel.Caption = Round(rstPunch.Fields("Total"), 2)
              Case "Prototypage-Dévelloppement expérimental":      lblTempsPrototypeReel.Caption = Round(rstPunch.Fields("Total"), 2)
225         End Select
230       End If

235       Call rstPunch.MoveNext
240     Loop

245     Call rstPunch.Close

        'Ouverture des enregistrements avec comme filtre, le numéro du projet
250     Call rstPunch.Open("SELECT " & sTotal & " FROM GRB_Punch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null", g_connData, adOpenDynamic, adLockOptimistic)

255     If Not IsNull(rstPunch.Fields("Total")) Then
260       lblTotalTempsRHReel.Caption = Round(rstPunch.Fields("Total"), 2)
265     Else
270       lblTotalTempsRHReel.Caption = "0"
275     End If

280     Call rstPunch.Close
285     Set rstPunch = Nothing

290     Exit Sub

AfficherErreur:

295     woups "frmProjSoumElecTemps", "AfficherTempsReels", Err, Erl
End Sub

Private Sub CalculerTotalArgent()

5       On Error GoTo AfficherErreur

10      If IsNumeric(lblTempsDessinReel.Caption) Then
15        lblPrixDessin.Caption = Round(Replace(lblTempsDessinReel.Caption * m_sTauxDessin, ".", ","), 2)
20      Else
25        lblPrixDessin.Caption = 0
30      End If

35      If IsNumeric(lblTempsFabricationReel.Caption) Then
40        lblPrixFabrication.Caption = Round(Replace(lblTempsFabricationReel.Caption * m_sTauxFabrication, ".", ","), 2)
45      Else
50        lblPrixFabrication.Caption = 0
55      End If

60      If IsNumeric(lblTempsAssemblageReel.Caption) Then
65        lblPrixAssemblage.Caption = Round(Replace(lblTempsAssemblageReel.Caption * m_sTauxAssemblage, ".", ","), 2)
70      Else
75        lblPrixAssemblage.Caption = 0
80      End If

85      If IsNumeric(lblTempsProgInterfaceReel.Caption) Then
90        lblPrixProgInterface.Caption = Round(Replace(lblTempsProgInterfaceReel.Caption * m_sTauxProgInterface, ".", ","), 2)
95      Else
100       lblPrixProgInterface.Caption = 0
105     End If

110     If IsNumeric(lblTempsProgAutomateReel.Caption) Then
115       lblPrixProgAutomate.Caption = Round(Replace(lblTempsProgAutomateReel.Caption * m_sTauxProgAutomate, ".", ","), 2)
120     Else
125       lblPrixProgAutomate.Caption = 0
130     End If

135     If IsNumeric(lblTempsProgRobotReel.Caption) Then
140       lblPrixProgRobot.Caption = Round(Replace(lblTempsProgRobotReel.Caption * m_sTauxProgRobot, ".", ","), 2)
145     Else
150       lblPrixProgRobot.Caption = 0
155     End If

160     If IsNumeric(lblTempsVisionReel.Caption) Then
165       lblPrixVision.Caption = Round(Replace(lblTempsVisionReel.Caption * m_sTauxVision, ".", ","), 2)
170     Else
175       lblPrixVision.Caption = 0
180     End If

185     If IsNumeric(lblTempsTestReel.Caption) Then
190       lblPrixTest.Caption = Round(Replace(lblTempsTestReel.Caption * m_sTauxTest, ".", ","), 2)
195     Else
200       lblPrixTest.Caption = 0
205     End If

210     If IsNumeric(lblTempsInstallationReel.Caption) Then
215       lblPrixInstallation.Caption = Round(Replace(lblTempsInstallationReel.Caption * m_sTauxInstallation, ".", ","), 2)
220     Else
225       lblPrixInstallation.Caption = 0
230     End If

235     If IsNumeric(lblTempsMiseServiceReel.Caption) Then
240       lblPrixMiseService.Caption = Round(Replace(lblTempsMiseServiceReel.Caption * m_sTauxMiseService, ".", ","), 2)
245     Else
250       lblPrixMiseService.Caption = 0
255     End If

260     If IsNumeric(lblTempsFormationReel.Caption) Then
265       lblPrixFormation.Caption = Round(Replace(lblTempsFormationReel.Caption * m_sTauxFormation, ".", ","), 2)
270     Else
275       lblPrixFormation.Caption = 0
280     End If

285     If IsNumeric(lblTempsGestionReel.Caption) Then
290       lblPrixGestion.Caption = Round(Replace(lblTempsGestionReel.Caption * m_sTauxGestion, ".", ","), 2)
295     Else
300       lblPrixGestion.Caption = 0
305     End If

310     If IsNumeric(lblTempsShippingReel.Caption) Then
315       lblPrixShipping.Caption = Round(Replace(lblTempsShippingReel.Caption * m_sTauxShipping, ".", ","), 2)
320     Else
325       lblPrixShipping.Caption = 0
330     End If

        If IsNumeric(lblTempsPrototypeReel.Caption) Then
331       lblPrixPrototype.Caption = Round(Replace(lblTempsPrototypeReel.Caption * m_sTauxGestion, ".", ","), 2)
332     Else
333       lblPrixPrototype.Caption = 0
334     End If






335     Call CalculerTotal

340     Exit Sub

AfficherErreur:

345     woups "frmProjSoumElecTemps", "CalculerTotalArgent", Err, Erl
End Sub

Private Sub BarrerChamps(ByVal bLocked As Boolean)

5       On Error GoTo AfficherErreur

10      txtTempsDessinSoum.Locked = bLocked
15      txtTempsAssemblageSoum.Locked = bLocked
20      txtTempsProgInterfaceSoum.Locked = bLocked
25      txtTempsProgAutomateSoum.Locked = bLocked
30      txtTempsProgRobotSoum.Locked = bLocked
35      txtTempsVisionSoum.Locked = bLocked
40      txtTempsTestSoum.Locked = bLocked
45      txtTempsInstallationSoum.Locked = bLocked
50      txtTempsMiseServiceSoum.Locked = bLocked
55      txtTempsFormationSoum.Locked = bLocked
60      txtTempsGestionSoum.Locked = bLocked
65      txtTempsShippingSoum.Locked = bLocked

70      txtNbrePersonne.Locked = bLocked
75      txtTempsHebergement.Locked = bLocked
80      txtTempsRepas.Locked = bLocked
85      txtTempsDeplacement.Locked = bLocked
90      txtTempsUniteMobile.Locked = bLocked
  
95      txtPrixEmballage.Locked = bLocked

100     Exit Sub

AfficherErreur:

105     woups "frmProjSoumElecTemps", "BarrerChamps", Err, Erl
End Sub

Private Sub cmdDetail_Click()

5       On Error GoTo AfficherErreur

10      If m_eType = TYPE_PROJET Then
15        Call frmDetailTemps.Afficher(m_sNoProjSoum, ELECTRIQUE, True)
20      Else
25        Call frmDetailTemps.Afficher(m_sNoProjSoum, ELECTRIQUE, False)
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmProjSoumElecTemps", "cmdDetail_Click", Err, Erl
End Sub

Private Sub Cmdfermer_Click()

5       On Error GoTo AfficherErreur

10      If m_eMode = MODE_AJOUT_MODIF Then
15        Call EnregistrerTemps
20      End If

25      Call Unload(Me)

30      Exit Sub

AfficherErreur:

35      woups "frmProjSoumElecTemps", "cmdFermer_Click", Err, Erl
End Sub

Private Sub EnregistrerTemps()

5       On Error GoTo AfficherErreur

10      If m_eType = TYPE_SOUMISSION Then
15        If Trim$(txtTempsDessinSoum.Text) <> vbNullString And IsNumeric(txtTempsDessinSoum.Text) Then
20          FrmProjSoumElec.m_sTempsDessin = txtTempsDessinSoum.Text
25        Else
30          FrmProjSoumElec.m_sTempsDessin = "0"
35        End If
        
40        If Trim$(lblTempsFabricationSoum.Caption) <> vbNullString Then
45          FrmProjSoumElec.m_sTempsFabrication = lblTempsFabricationSoum.Caption
50        Else
55          FrmProjSoumElec.m_sTempsFabrication = "0"
60        End If
      
65        If Trim$(txtTempsAssemblageSoum.Text) <> vbNullString And IsNumeric(txtTempsAssemblageSoum.Text) Then
70          FrmProjSoumElec.m_sTempsAssemblage = txtTempsAssemblageSoum.Text
75        Else
80          FrmProjSoumElec.m_sTempsAssemblage = "0"
85        End If
      
90        If Trim$(txtTempsProgInterfaceSoum.Text) <> vbNullString And IsNumeric(txtTempsProgInterfaceSoum.Text) Then
95          FrmProjSoumElec.m_sTempsProgInterface = txtTempsProgInterfaceSoum.Text
100       Else
105         FrmProjSoumElec.m_sTempsProgInterface = "0"
110       End If
      
115       If Trim$(txtTempsProgAutomateSoum.Text) <> vbNullString And IsNumeric(txtTempsProgAutomateSoum.Text) Then
120         FrmProjSoumElec.m_sTempsProgAutomate = txtTempsProgAutomateSoum.Text
125       Else
130         FrmProjSoumElec.m_sTempsProgAutomate = "0"
135       End If
      
140       If Trim$(txtTempsProgRobotSoum.Text) <> vbNullString And IsNumeric(txtTempsProgRobotSoum.Text) Then
145         FrmProjSoumElec.m_sTempsProgRobot = txtTempsProgRobotSoum.Text
150       Else
155         FrmProjSoumElec.m_sTempsProgRobot = "0"
160       End If
      
165       If Trim$(txtTempsVisionSoum.Text) <> vbNullString And IsNumeric(txtTempsVisionSoum.Text) Then
170         FrmProjSoumElec.m_sTempsVision = txtTempsVisionSoum.Text
175       Else
180         FrmProjSoumElec.m_sTempsVision = "0"
185       End If
      
190       If Trim$(txtTempsTestSoum.Text) <> vbNullString And IsNumeric(txtTempsTestSoum.Text) Then
195         FrmProjSoumElec.m_sTempsTest = txtTempsTestSoum.Text
200       Else
205         FrmProjSoumElec.m_sTempsTest = "0"
210       End If
      
215       If Trim$(txtTempsInstallationSoum.Text) <> vbNullString And IsNumeric(txtTempsInstallationSoum.Text) Then
220         FrmProjSoumElec.m_sTempsInstallation = txtTempsInstallationSoum.Text
225       Else
230         FrmProjSoumElec.m_sTempsInstallation = "0"
235       End If
      
240       If Trim$(txtTempsMiseServiceSoum.Text) <> vbNullString And IsNumeric(txtTempsMiseServiceSoum.Text) Then
245         FrmProjSoumElec.m_sTempsMiseService = txtTempsMiseServiceSoum.Text
250       Else
255         FrmProjSoumElec.m_sTempsMiseService = "0"
260       End If
  
265       If Trim$(txtTempsFormationSoum.Text) <> vbNullString And IsNumeric(txtTempsFormationSoum.Text) Then
270         FrmProjSoumElec.m_sTempsFormation = txtTempsFormationSoum.Text
275       Else
280         FrmProjSoumElec.m_sTempsFormation = "0"
285       End If
      
290       If Trim$(txtTempsGestionSoum.Text) <> vbNullString And IsNumeric(txtTempsGestionSoum.Text) Then
295         FrmProjSoumElec.m_sTempsGestion = txtTempsGestionSoum.Text
300       Else
305         FrmProjSoumElec.m_sTempsGestion = "0"
310       End If

315       If Trim$(txtTempsShippingSoum.Text) <> vbNullString And IsNumeric(txtTempsShippingSoum.Text) Then
320         FrmProjSoumElec.m_sTempsShipping = txtTempsShippingSoum.Text
325       Else
330         FrmProjSoumElec.m_sTempsShipping = "0"
335       End If
340     End If

        
345     If m_bSansTemps = True Then
350       FrmProjSoumElec.m_bSansTemps = True
355       FrmProjSoumElec.tmrTemps.Enabled = True
360     Else
365       FrmProjSoumElec.m_bSansTemps = False
370       FrmProjSoumElec.tmrTemps.Enabled = False
375       FrmProjSoumElec.lblPasTemps.Visible = False
380     End If

385     If Trim$(txtNbrePersonne.Text) <> vbNullString And IsNumeric(txtNbrePersonne.Text) Then
390       FrmProjSoumElec.m_sNbrePersonne = txtNbrePersonne.Text
395     Else
400       FrmProjSoumElec.m_sNbrePersonne = "0"
405     End If
  
410     If Trim$(txtTempsHebergement.Text) <> vbNullString And IsNumeric(txtTempsHebergement.Text) Then
415       FrmProjSoumElec.m_sTempsHebergement = txtTempsHebergement.Text
420     Else
425       FrmProjSoumElec.m_sTempsHebergement = "0"
430     End If
       
435     If Trim$(txtTempsRepas.Text) <> vbNullString And IsNumeric(txtTempsRepas.Text) Then
440       FrmProjSoumElec.m_sTempsRepas = txtTempsRepas.Text
445     Else
450       FrmProjSoumElec.m_sTempsRepas = "0"
455     End If
    
460     If Trim$(txtTempsDeplacement.Text) <> vbNullString And IsNumeric(txtTempsDeplacement.Text) Then
465       FrmProjSoumElec.m_sTempsTransport = txtTempsDeplacement.Text
470     Else
475       FrmProjSoumElec.m_sTempsTransport = "0"
480     End If
    
485     If Trim$(txtTempsUniteMobile.Text) <> vbNullString And IsNumeric(txtTempsUniteMobile.Text) Then
490       FrmProjSoumElec.m_sTempsUniteMobile = txtTempsUniteMobile.Text
495     Else
500       FrmProjSoumElec.m_sTempsUniteMobile = "0"
505     End If
    
510     If Trim$(txtPrixEmballage.Text) <> vbNullString And IsNumeric(txtPrixEmballage.Text) Then
515       FrmProjSoumElec.m_sPrixEmballage = txtPrixEmballage.Text
520     Else
525       FrmProjSoumElec.m_sPrixEmballage = "0"
530     End If
    
535     FrmProjSoumElec.m_sTauxHebergement1 = m_sHebergement1
540     FrmProjSoumElec.m_sTauxHebergement2 = m_sHebergement2
545     FrmProjSoumElec.m_sTauxRepas = m_sRepas
550     FrmProjSoumElec.m_sTauxTransport = m_sStandard
555     FrmProjSoumElec.m_sTauxUniteMobile = m_sUniteMobile

560     Exit Sub

AfficherErreur:

565     woups "frmProjSoumElecTemps", "EnregistrerTemps", Err, Erl
End Sub

Private Sub InitialiserVariablesConfig()

5       On Error GoTo AfficherErreur

        'Initialise les variables à partir de la table Config (Pour avoir le taux
        'horaire le plus récent)
10      Dim rstConfig As ADODB.Recordset
  
15      Set rstConfig = New ADODB.Recordset
  
20      Call rstConfig.Open("SELECT TauxDessinElec, TauxFabrication, TauxAssemblageElec, TauxProgInterface, TauxProgAutomate, TauxProgRobot, TauxVision, TauxTestElec, TauxInstallationElec, TauxMiseService, TauxFormationElec, TauxGestionProjetsElec, TauxShippingElec, Repas, Hebergement1, Hebergement2, Standard, UniteMobile FROM GRB_Config", g_connData, adOpenDynamic, adLockOptimistic)
    
25      If Not IsNull(rstConfig.Fields("TauxDessinElec")) Then
30        m_sTauxDessin = rstConfig.Fields("TauxDessinElec")
35      Else
40        m_sTauxDessin = "0"
45      End If

50      If Not IsNull(rstConfig.Fields("TauxFabrication")) Then
55        m_sTauxFabrication = rstConfig.Fields("TauxFabrication")
60      Else
65        m_sTauxFabrication = "0"
70      End If

75      If Not IsNull(rstConfig.Fields("TauxAssemblageElec")) Then
80        m_sTauxAssemblage = rstConfig.Fields("TauxAssemblageElec")
85      Else
90        m_sTauxAssemblage = "0"
95      End If

100     If Not IsNull(rstConfig.Fields("TauxProgInterface")) Then
105       m_sTauxProgInterface = rstConfig.Fields("TauxProgInterface")
110     Else
115       m_sTauxProgInterface = "0"
120     End If

125     If Not IsNull(rstConfig.Fields("TauxProgAutomate")) Then
130       m_sTauxProgAutomate = rstConfig.Fields("TauxProgAutomate")
135     Else
140       m_sTauxProgAutomate = "0"
145     End If

150     If Not IsNull(rstConfig.Fields("TauxProgRobot")) Then
155       m_sTauxProgRobot = rstConfig.Fields("TauxProgRobot")
160     Else
165       m_sTauxProgRobot = "0"
170     End If

175     If Not IsNull(rstConfig.Fields("TauxVision")) Then
180       m_sTauxVision = rstConfig.Fields("TauxVision")
185     Else
190       m_sTauxVision = "0"
195     End If

200     If Not IsNull(rstConfig.Fields("TauxTestElec")) Then
205       m_sTauxTest = rstConfig.Fields("TauxTestElec")
210     Else
215       m_sTauxTest = "0"
220     End If

225     If Not IsNull(rstConfig.Fields("TauxInstallationElec")) Then
230       m_sTauxInstallation = rstConfig.Fields("TauxInstallationElec")
235     Else
240       m_sTauxInstallation = "0"
245     End If

250     If Not IsNull(rstConfig.Fields("TauxMiseService")) Then
255       m_sTauxMiseService = rstConfig.Fields("TauxMiseService")
260     Else
265       m_sTauxMiseService = "0"
270     End If

275     If Not IsNull(rstConfig.Fields("TauxFormationElec")) Then
280       m_sTauxFormation = rstConfig.Fields("TauxFormationElec")
285     Else
290       m_sTauxFormation = "0"
295     End If

300     If Not IsNull(rstConfig.Fields("TauxGestionProjetsElec")) Then
305       m_sTauxGestion = rstConfig.Fields("TauxGestionProjetsElec")
310     Else
315       m_sTauxGestion = "0"
320     End If

325     If Not IsNull(rstConfig.Fields("TauxShippingElec")) Then
330       m_sTauxShipping = rstConfig.Fields("TauxShippingElec")
335     Else
340       m_sTauxShipping = "0"
345     End If
    
350     m_sRepas = rstConfig.Fields("Repas")
355     m_sHebergement1 = rstConfig.Fields("Hebergement1")
360     m_sHebergement2 = rstConfig.Fields("Hebergement2")
365     m_sStandard = rstConfig.Fields("Standard")
370     m_sUniteMobile = rstConfig.Fields("UniteMobile")
    
375     Call rstConfig.Close
380     Set rstConfig = Nothing

385     Exit Sub

AfficherErreur:

390     woups "frmProjSoumElecTemps", "InitialiserVariablesConfig", Err, Erl
End Sub

Private Sub InitialiserVariablesProjSoum()

5       On Error GoTo AfficherErreur

10      m_sTauxDessin = FrmProjSoumElec.m_sTauxDessin
15      m_sTauxFabrication = FrmProjSoumElec.m_sTauxFabrication
20      m_sTauxAssemblage = FrmProjSoumElec.m_sTauxAssemblage
25      m_sTauxProgInterface = FrmProjSoumElec.m_sTauxProgInterface
30      m_sTauxProgAutomate = FrmProjSoumElec.m_sTauxProgAutomate
35      m_sTauxProgRobot = FrmProjSoumElec.m_sTauxProgRobot
40      m_sTauxVision = FrmProjSoumElec.m_sTauxVision
45      m_sTauxTest = FrmProjSoumElec.m_sTauxTest
50      m_sTauxInstallation = FrmProjSoumElec.m_sTauxInstallation
55      m_sTauxMiseService = FrmProjSoumElec.m_sTauxMiseService
60      m_sTauxFormation = FrmProjSoumElec.m_sTauxFormation
65      m_sTauxGestion = FrmProjSoumElec.m_sTauxGestion
70      m_sTauxShipping = FrmProjSoumElec.m_sTauxShipping

75      m_sRepas = FrmProjSoumElec.m_sTauxRepas
80      m_sHebergement1 = FrmProjSoumElec.m_sTauxHebergement1
85      m_sHebergement2 = FrmProjSoumElec.m_sTauxHebergement2
90      m_sStandard = FrmProjSoumElec.m_sTauxTransport
95      m_sUniteMobile = FrmProjSoumElec.m_sTauxUniteMobile
  
100     Exit Sub

AfficherErreur:

105     woups "frmProjSoumElecTemps", "InitialiserVariablesProjSoum", Err, Erl
End Sub

Private Sub Form_Load()
  
5       On Error GoTo AfficherErreur
  
10      If FrmProjSoumElec.m_bDroitPrix = False Then
15        Me.width = 3765
20        Cmdfermer.Left = 2280
25      Else
30        Me.width = 11115
35        Cmdfermer.Left = 9480
40      End If

45      Exit Sub

AfficherErreur:

50      woups "frmProjSoumElecTemps", "Form_Load", Err, Erl
End Sub

Private Sub txtNbrePersonne_Change()

5       On Error GoTo AfficherErreur

10      If txtTempsHebergement.Text <> vbNullString Then
15        If IsNumeric(txtNbrePersonne.Text) Then
20          Call CalculerHebergement
25          Call CalculerRepas

30          Call CalculerTotal
35        End If
40      End If

45      Exit Sub

AfficherErreur:

50      woups "frmProjSoumElecTemps", "txtNbrePersonne_Change", Err, Erl
End Sub

Private Sub txtPrixEmballage_KeyPress(KeyAscii As Integer)

5       On Error GoTo AfficherErreur

10      If KeyAscii = 46 Then 'Si c'est le "."
15        KeyAscii = 44 'Remplace par la ","
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmProjSoumElecTemps", "lblPrixEmballage_KeyPress", Err, Erl
End Sub

Private Sub txtTempsHebergement_Change()

5       On Error GoTo AfficherErreur

10      If IsNumeric(txtTempsHebergement.Text) Then
15        If txtNbrePersonne.Text <> vbNullString Then
20          Call CalculerHebergement
25        End If
30      End If
  
35      Call CalculerTotal

40      Exit Sub

AfficherErreur:

45      woups "frmProjSoumElecTemps", "txtTempsHebergement_Change", Err, Erl
End Sub

Private Sub txtTempsHebergement_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsHebergement.Text = Replace(txtTempsHebergement.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElecTemps", "txtTempsHebergement_LostFocus", Err, Erl
End Sub

Private Sub txtTempsRepas_Change()

5       On Error GoTo AfficherErreur

10      If IsNumeric(txtTempsRepas.Text) Then
15        If txtNbrePersonne.Text <> vbNullString Then
20          Call CalculerRepas
25        End If
30      End If
  
35      Call CalculerTotal

40      Exit Sub

AfficherErreur:

45      woups "frmProjSoumElecTemps", "txtTempsRepas_Change", Err, Erl
End Sub

Private Sub txtTempsRepas_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsRepas.Text = Replace(txtTempsRepas.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElecTemps", "txtTempsRepas_LostFocus", Err, Erl
End Sub

Private Sub txtTempsDeplacement_Change()

5       On Error GoTo AfficherErreur

10      If IsNumeric(txtTempsDeplacement.Text) Then
15        lblPrixDeplacement.Caption = Round(Replace(txtTempsDeplacement.Text * m_sStandard, ".", ","), 2)
20      Else
25        lblPrixDeplacement.Caption = 0
30      End If
  
35      Call CalculerTotal

40      Exit Sub

AfficherErreur:

45      woups "frmProjSoumElecTemps", "txtTempsDeplacement_Change", Err, Erl
End Sub

Private Sub txtTempsDeplacement_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsDeplacement.Text = Replace(txtTempsDeplacement.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElecTemps", "txtTempsDeplacement_LostFocus", Err, Erl
End Sub

Private Sub txtTempsUniteMobile_Change()

5       On Error GoTo AfficherErreur

10      If IsNumeric(txtTempsUniteMobile.Text) Then
15        lblPrixUniteMobile.Caption = Round(Replace(txtTempsUniteMobile.Text * m_sUniteMobile, ".", ","), 2)
20      Else
25        lblPrixUniteMobile.Caption = 0
30      End If
  
35      Call CalculerTotal

40      Exit Sub

AfficherErreur:

45      woups "frmProjSoumElecTemps", "txtTempsUniteMobile_Change", Err, Erl
End Sub

Private Sub txtTempsUniteMobile_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsUniteMobile.Text = Replace(txtTempsUniteMobile.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElecTemps", "txtTempsUniteMobile_LostFocus", Err, Erl
End Sub

Private Sub txtPrixEmballage_Change()

5       On Error GoTo AfficherErreur
        
10      Call CalculerTotal

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElecTemps", "lblPrixEmballage_Change", Err, Erl
End Sub

Private Sub txtPrixEmballage_LostFocus()

5       On Error GoTo AfficherErreur

10      If IsNumeric(txtPrixEmballage.Text) Then
15        txtPrixEmballage.Text = Round(Replace(txtPrixEmballage.Text, ".", ","), 2)
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmProjSoumElecTemps", "lblPrixEmballage_LostFocus", Err, Erl
End Sub

Private Sub txtTempsDessinSoum_Change()

5       On Error GoTo AfficherErreur

10      If m_eType = TYPE_SOUMISSION Then
15        If IsNumeric(txtTempsDessinSoum.Text) Then
20          lblPrixDessin.Caption = Round(Replace(txtTempsDessinSoum.Text * m_sTauxDessin, ".", ","), 2)
25        Else
30          lblPrixDessin.Caption = 0
35        End If
  
40        Call CalculerTotal
45      End If

50      Exit Sub

AfficherErreur:

55      woups "frmProjSoumElecTemps", "txtTempsDessinSoum_Change", Err, Erl
End Sub

Private Sub txtTempsDessinSoum_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsDessinSoum.Text = Replace(txtTempsDessinSoum.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElecTemps", "txtTempsDessinSoum_LostFocus", Err, Erl
End Sub

Private Sub lblTempsFabricationSoum_Change()

5       On Error GoTo AfficherErreur

10      If m_eType = TYPE_SOUMISSION Then
15        If m_bSansTemps = False Then
20          If IsNumeric(lblTempsFabricationSoum.Caption) Then
25            lblPrixFabrication.Caption = Round(Replace(lblTempsFabricationSoum.Caption * m_sTauxFabrication, ".", ","), 2)
30          Else
35            lblPrixFabrication.Caption = "0"
40          End If
45        Else
50          lblPrixFabrication.Caption = "0"
55        End If
  
60        Call CalculerTotal
65      End If

70      Exit Sub

AfficherErreur:

75      woups "frmProjSoumElecTemps", "txtTempsMécanique_Change", Err, Erl
End Sub

Private Sub txtTempsAssemblageSoum_Change()

5       On Error GoTo AfficherErreur

10      If m_eType = TYPE_SOUMISSION Then
15        If IsNumeric(txtTempsAssemblageSoum.Text) Then
20          lblPrixAssemblage.Caption = Round(Replace(txtTempsAssemblageSoum.Text * m_sTauxAssemblage, ".", ","), 2)
25        Else
30          lblPrixAssemblage.Caption = 0
35        End If
  
40        Call CalculerTotal
45      End If

50      Exit Sub

AfficherErreur:

55      woups "frmProjSoumElecTemps", "txtTempsAssemblageSoum_Change", Err, Erl
End Sub

Private Sub txtTempsAssemblageSoum_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsAssemblageSoum.Text = Replace(txtTempsAssemblageSoum.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElecTemps", "txtTempsAssemblageSoum_LostFocus", Err, Erl
End Sub

Private Sub txtTempsProgInterfaceSoum_Change()

5       On Error GoTo AfficherErreur

10      If m_eType = TYPE_SOUMISSION Then
15        If IsNumeric(txtTempsProgInterfaceSoum.Text) Then
20          lblPrixProgInterface.Caption = Round(Replace(txtTempsProgInterfaceSoum.Text * m_sTauxProgInterface, ".", ","), 2)
25        Else
30          lblPrixProgInterface.Caption = 0
35        End If
  
40        Call CalculerTotal
45      End If

50      Exit Sub

AfficherErreur:

55      woups "frmProjSoumElecTemps", "txtTempsProgInterfaceSoum_Change", Err, Erl
End Sub

Private Sub txtTempsProgInterfaceSoum_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsProgInterfaceSoum.Text = Replace(txtTempsProgInterfaceSoum.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElecTemps", "txtTempsProgInterfaceSoum_LostFocus", Err, Erl
End Sub

Private Sub txtTempsProgAutomateSoum_Change()

5       On Error GoTo AfficherErreur

10      If m_eType = TYPE_SOUMISSION Then
15        If IsNumeric(txtTempsProgAutomateSoum.Text) Then
20          lblPrixProgAutomate.Caption = Round(Replace(txtTempsProgAutomateSoum.Text * m_sTauxProgAutomate, ".", ","), 2)
25        Else
30          lblPrixProgAutomate.Caption = 0
35        End If
  
40        Call CalculerTotal
45      End If

50      Exit Sub

AfficherErreur:

55      woups "frmProjSoumElecTemps", "txtTempsProgAutomate_Change", Err, Erl
End Sub

Private Sub txtTempsProgAutomateSoum_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsProgAutomateSoum.Text = Replace(txtTempsProgAutomateSoum.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElecTemps", "txtTempsProgAutomate_LostFocus", Err, Erl
End Sub

Private Sub txtTempsProgRobotSoum_Change()

5       On Error GoTo AfficherErreur

10      If m_eType = TYPE_SOUMISSION Then
15        If IsNumeric(txtTempsProgRobotSoum.Text) Then
20          lblPrixProgRobot.Caption = Round(Replace(txtTempsProgRobotSoum.Text * m_sTauxProgRobot, ".", ","), 2)
25        Else
30          lblPrixProgRobot.Caption = 0
35        End If
  
40        Call CalculerTotal
45      End If

50      Exit Sub

AfficherErreur:

55      woups "frmProjSoumElecTemps", "txtTempsProgRobotSoum_Change", Err, Erl
End Sub

Private Sub txtTempsProgRobotSoum_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsProgRobotSoum.Text = Replace(txtTempsProgRobotSoum.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElecTemps", "txtTempsProgRobotSoum_LostFocus", Err, Erl
End Sub

Private Sub txtTempsVisionSoum_Change()

5       On Error GoTo AfficherErreur

10      If m_eType = TYPE_SOUMISSION Then
15        If IsNumeric(txtTempsVisionSoum.Text) Then
20          lblPrixVision.Caption = Round(Replace(txtTempsVisionSoum.Text * m_sTauxVision, ".", ","), 2)
25        Else
30          lblPrixVision.Caption = 0
35        End If
  
40        Call CalculerTotal
45      End If

50      Exit Sub

AfficherErreur:

55      woups "frmProjSoumElecTemps", "txtTempsVisionSoum_Change", Err, Erl
End Sub

Private Sub txtTempsVisionSoum_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsVisionSoum.Text = Replace(txtTempsVisionSoum.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElecTemps", "txtTempsVisionSoum_LostFocus", Err, Erl
End Sub

Private Sub txtTempsTestSoum_Change()

5       On Error GoTo AfficherErreur

10      If m_eType = TYPE_SOUMISSION Then
15        If IsNumeric(txtTempsTestSoum.Text) Then
20          lblPrixTest.Caption = Round(Replace(txtTempsTestSoum.Text * m_sTauxTest, ".", ","), 2)
25        Else
30          lblPrixTest.Caption = 0
35        End If
  
40        Call CalculerTotal
45      End If

50      Exit Sub

AfficherErreur:

55      woups "frmProjSoumElecTemps", "txtTempsTestSoum_Change", Err, Erl
End Sub

Private Sub txtTempsTestSoum_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsTestSoum.Text = Replace(txtTempsTestSoum.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElecTemps", "txtTempsTestSoum_LostFocus", Err, Erl
End Sub

Private Sub txtTempsInstallationSoum_Change()

5       On Error GoTo AfficherErreur

10      If m_eType = TYPE_SOUMISSION Then
15        If IsNumeric(txtTempsInstallationSoum.Text) Then
20          lblPrixInstallation.Caption = Round(Replace(txtTempsInstallationSoum.Text * m_sTauxInstallation, ".", ","), 2)
25        Else
30          lblPrixInstallation.Caption = 0
35        End If
  
40        Call CalculerTotal
45      End If

50      Exit Sub

AfficherErreur:

55      woups "frmProjSoumElecTemps", "txtTempsInstallationSoum_Change", Err, Erl
End Sub

Private Sub txtTempsInstallationSoum_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsInstallationSoum.Text = Replace(txtTempsInstallationSoum.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElecTemps", "txtTempsInstallationSoum_LostFocus", Err, Erl
End Sub

Private Sub txtTempsMiseServiceSoum_Change()

5       On Error GoTo AfficherErreur

10      If m_eType = TYPE_SOUMISSION Then
15        If IsNumeric(txtTempsMiseServiceSoum.Text) Then
20          lblPrixMiseService.Caption = Round(Replace(txtTempsMiseServiceSoum.Text * m_sTauxMiseService, ".", ","), 2)
25        Else
30          lblPrixMiseService.Caption = 0
35        End If
  
40        Call CalculerTotal
45      End If

50      Exit Sub

AfficherErreur:

55      woups "frmProjSoumElecTemps", "txtTempsMiseServiceSoum_Change", Err, Erl
End Sub

Private Sub txtTempsMiseServiceSoum_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsMiseServiceSoum.Text = Replace(txtTempsMiseServiceSoum.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElecTemps", "txtTempsMiseServiceSoum_LostFocus", Err, Erl
End Sub

Private Sub txtTempsFormationSoum_Change()

5       On Error GoTo AfficherErreur

10      If m_eType = TYPE_SOUMISSION Then
15        If IsNumeric(txtTempsFormationSoum.Text) Then
20          lblPrixFormation.Caption = Round(Replace(txtTempsFormationSoum.Text * m_sTauxFormation, ".", ","), 2)
25        Else
30          lblPrixFormation.Caption = 0
35        End If
  
40        Call CalculerTotal
45      End If

50      Exit Sub

AfficherErreur:

55      woups "frmProjSoumElecTemps", "txtTempsFormationSoum_Change", Err, Erl
End Sub

Private Sub txtTempsFormationSoum_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsFormationSoum.Text = Replace(txtTempsFormationSoum.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElecTemps", "txtTempsFormationSoum_LostFocus", Err, Erl
End Sub

Private Sub txtTempsGestionSoum_Change()

5       On Error GoTo AfficherErreur

10      If m_eType = TYPE_SOUMISSION Then
15        If IsNumeric(txtTempsGestionSoum.Text) Then
20          lblPrixGestion.Caption = Round(Replace(txtTempsGestionSoum.Text * m_sTauxGestion, ".", ","), 2)
25        Else
30          lblPrixGestion.Caption = 0
35        End If
  
40        Call CalculerTotal
45      End If

50      Exit Sub

AfficherErreur:

55      woups "frmProjSoumElecTemps", "txtTempsGestionSoum_Change", Err, Erl
End Sub

Private Sub txtTempsGestionSoum_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsGestionSoum.Text = Replace(txtTempsGestionSoum.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElecTemps", "txtTempsGestionSoum_LostFocus", Err, Erl
End Sub

Private Sub txtTempsShippingSoum_Change()

5       On Error GoTo AfficherErreur

10      If m_eType = TYPE_SOUMISSION Then
15        If IsNumeric(txtTempsShippingSoum.Text) Then
20          lblPrixShipping.Caption = Round(Replace(txtTempsShippingSoum.Text * m_sTauxShipping, ".", ","), 2)
25        Else
30          lblPrixShipping.Caption = 0
35        End If
  
40        Call CalculerTotal
45      End If

50      Exit Sub

AfficherErreur:

55      woups "frmProjSoumElecTemps", "txtTempsShippingSoum_Change", Err, Erl
End Sub

Private Sub txtTempsShippingSoum_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsShippingSoum.Text = Replace(txtTempsShippingSoum.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElecTemps", "txtTempsShippingSoum_LostFocus", Err, Erl
End Sub


Private Sub CalculerHebergement()

5       On Error GoTo AfficherErreur

10      Dim dblNbreDeux    As Double
15      Dim dblHebergement As Double
20      Dim iReste         As Integer
25      Dim dblNbrePers    As Double
30      Dim dblNbreJours   As Double
  
35      If IsNumeric(txtNbrePersonne.Text) Then
40        dblNbrePers = CDbl(txtNbrePersonne.Text)
45      Else
50        dblNbrePers = 0
55      End If
  
60      If IsNumeric(txtTempsHebergement.Text) Then
65        dblNbreJours = CDbl(txtTempsHebergement.Text)
70      Else
75        dblNbreJours = 0
80      End If
  
85      dblNbreDeux = Int(dblNbrePers / 2)
  
90      iReste = CInt(dblNbrePers) - (dblNbreDeux * 2)
  
95      dblHebergement = dblNbreJours * ((dblNbreDeux * CDbl(m_sHebergement2)) + (iReste * CDbl(m_sHebergement1)))
  
100     lblPrixHebergement.Caption = Round(Replace(dblHebergement, ".", ","), 2)

105     Exit Sub

AfficherErreur:

110     woups "frmProjSoumElecTemps", "CalculerHebergement", Err, Erl
End Sub

Private Sub CalculerRepas()

5       On Error GoTo AfficherErreur

10      Dim dblNbrePers  As Double
15      Dim dblRepas     As Double
20      Dim dblNbreJours As Double

25      If IsNumeric(txtNbrePersonne.Text) Then
30        dblNbrePers = CDbl(txtNbrePersonne.Text)
35      Else
40        dblNbrePers = 0
45      End If
  
50      If IsNumeric(txtTempsRepas.Text) Then
55        dblNbreJours = CDbl(txtTempsRepas.Text)
60      Else
65        dblNbreJours = 0
70      End If
  
75      dblRepas = dblNbreJours * dblNbrePers * CDbl(m_sRepas)
  
80      lblPrixRepas.Caption = Round(Replace(dblRepas, ".", ","), 2)

85      Exit Sub

AfficherErreur:

90     woups "frmProjSoumElecTemps", "CalculerRepas", Err, Erl
End Sub

Private Sub CalculerTotal()

5       On Error GoTo AfficherErreur

10      Dim dblTotal             As Double
15      Dim dblPrixEmballage     As Double
20      Dim dblTotalArgentRH     As Double
25      Dim dblPrixDessin        As Double
30      Dim dblPrixFabrication   As Double
35      Dim dblPrixAssemblage    As Double
40      Dim dblPrixProgInterface As Double
45      Dim dblPrixProgAutomate  As Double
50      Dim dblPrixProgRobot     As Double
55      Dim dblPrixVision        As Double
60      Dim dblPrixTest          As Double
65      Dim dblPrixInstallation  As Double
70      Dim dblPrixMiseService   As Double
75      Dim dblPrixFormation     As Double
80      Dim dblPrixGestion       As Double
85      Dim dblPrixShipping      As Double
        Dim dblPrixPrototype      As Double
90      Dim dblPrixHebergement   As Double
95      Dim dblPrixRepas         As Double
100     Dim dblPrixDeplacement   As Double
105     Dim dblPrixUniteMobile   As Double
   
        'Prix de dessin
110     If IsNumeric(lblPrixDessin.Caption) Then
115       dblPrixDessin = CDbl(lblPrixDessin.Caption)
120     Else
125       dblPrixDessin = 0
130     End If

        'Prix de Fabrication
135     If IsNumeric(lblPrixFabrication.Caption) Then
140       dblPrixFabrication = CDbl(lblPrixFabrication.Caption)
145     Else
150       dblPrixFabrication = 0
155     End If

        'Prix de Assemblage
160     If IsNumeric(lblPrixAssemblage.Caption) Then
165       dblPrixAssemblage = CDbl(lblPrixAssemblage.Caption)
170     Else
175       dblPrixAssemblage = 0
180     End If

        'Prix de ProgInterface
185     If IsNumeric(lblPrixProgInterface.Caption) Then
190       dblPrixProgInterface = CDbl(lblPrixProgInterface.Caption)
195     Else
200       dblPrixProgInterface = 0
205     End If

        'Prix de ProgAutomate
210     If IsNumeric(lblPrixProgAutomate.Caption) Then
215       dblPrixProgAutomate = CDbl(lblPrixProgAutomate.Caption)
220     Else
225       dblPrixProgAutomate = 0
230     End If

        'Prix de ProgRobot
235     If IsNumeric(lblPrixProgRobot.Caption) Then
240       dblPrixProgRobot = CDbl(lblPrixProgRobot.Caption)
245     Else
250       dblPrixProgRobot = 0
255     End If

        'Prix de vision
260     If IsNumeric(lblPrixVision.Caption) Then
265       dblPrixVision = CDbl(lblPrixVision.Caption)
270     Else
275       dblPrixVision = 0
280     End If

        'Prix de test
285     If IsNumeric(lblPrixTest.Caption) Then
290       dblPrixTest = CDbl(lblPrixTest.Caption)
295     Else
300       dblPrixTest = 0
305     End If

        'Prix de Installation
310     If IsNumeric(lblPrixInstallation.Caption) Then
315       dblPrixInstallation = CDbl(lblPrixInstallation.Caption)
320     Else
325       dblPrixInstallation = 0
330     End If

        'Prix de MiseService
335     If IsNumeric(lblPrixMiseService.Caption) Then
340       dblPrixMiseService = CDbl(lblPrixMiseService.Caption)
345     Else
350       dblPrixMiseService = 0
355     End If

        'Prix de formation
360     If IsNumeric(lblPrixFormation.Caption) Then
365       dblPrixFormation = CDbl(lblPrixFormation.Caption)
370     Else
375       dblPrixFormation = 0
380     End If

        'Prix de Gestion
385     If IsNumeric(lblPrixGestion.Caption) Then
390       dblPrixGestion = CDbl(lblPrixGestion.Caption)
395     Else
400       dblPrixGestion = 0
405     End If

        'Prix de Shipping
410     If IsNumeric(lblPrixShipping.Caption) Then
415       dblPrixShipping = CDbl(lblPrixShipping.Caption)
420     Else
425       dblPrixShipping = 0
430     End If


        'Prix de Prototype
431     If IsNumeric(lblPrixPrototype.Caption) Then
432       dblPrixPrototype = CDbl(lblPrixPrototype.Caption)
433     Else
434       dblPrixPrototype = 0
435     End If



        'Prix d'hébergement
436     If IsNumeric(lblPrixHebergement.Caption) Then
440       dblPrixHebergement = CDbl(lblPrixHebergement.Caption)
445     Else
450       dblPrixHebergement = 0
455     End If

        'Prix des repas
460     If IsNumeric(lblPrixRepas.Caption) Then
465       dblPrixRepas = CDbl(lblPrixRepas.Caption)
470     Else
475       dblPrixRepas = 0
480     End If
  
        'Prix du déplacement
485     If IsNumeric(lblPrixDeplacement.Caption) Then
490       dblPrixDeplacement = CDbl(lblPrixDeplacement.Caption)
495     Else
500       dblPrixDeplacement = 0
505     End If

        'Prix de l'unité mobile
510     If IsNumeric(lblPrixUniteMobile.Caption) Then
515       dblPrixUniteMobile = CDbl(lblPrixUniteMobile.Caption)
520     Else
525       dblPrixUniteMobile = 0
530     End If
   
        'Prix de transport et emballage
535     If IsNumeric(txtPrixEmballage.Text) Then
540       dblPrixEmballage = CDbl(txtPrixEmballage.Text)
545     Else
550       dblPrixEmballage = 0
555     End If

560     dblTotalArgentRH = dblPrixDessin + _
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
                           
565     lblTotalPrixRH.Caption = Conversion(CStr(dblTotalArgentRH), MODE_DECIMAL)

570     dblTotal = dblTotalArgentRH + _
                   dblPrixHebergement + _
                   dblPrixRepas + _
                   dblPrixDeplacement + _
                   dblPrixUniteMobile + _
                   dblPrixEmballage

575     lblTotal.Caption = Conversion(CStr(dblTotal), MODE_DECIMAL)

580     Call CalculerTotalTemps

585     Exit Sub

AfficherErreur:

590     woups "frmProjSoumElecTemps", "CalculerTotal", Err, Erl
End Sub

Private Sub CalculerTotalTemps()

  
5       On Error GoTo AfficherErreur

10      Dim dblTempsDessin        As Double
15      Dim dblTempsFabrication   As Double
20      Dim dblTempsAssemblage    As Double
25      Dim dblTempsProgInterface As Double
30      Dim dblTempsProgAutomate  As Double
35      Dim dblTempsProgRobot     As Double
40      Dim dblTempsVision        As Double
45      Dim dblTempsTest          As Double
50      Dim dblTempsInstallation  As Double
55      Dim dblTempsMiseService   As Double
60      Dim dblTempsFormation     As Double
65      Dim dblTempsGestion       As Double
70      Dim dblTempsShipping      As Double
        Dim dblTempsPrototype     As Double
75      Dim dblTotalTemps         As Double

        'SOUMISSION
80      If IsNumeric(txtTempsDessinSoum.Text) Then
85        dblTempsDessin = CDbl(txtTempsDessinSoum.Text)
90      Else
95        dblTempsDessin = 0
100     End If

105     If m_bSansTemps = False Then
110       If IsNumeric(lblTempsFabricationSoum.Caption) Then
115         dblTempsFabrication = CDbl(lblTempsFabricationSoum.Caption)
120       Else
125         dblTempsFabrication = 0
130       End If
135     Else
140       dblTempsFabrication = 0
145     End If

150     If IsNumeric(txtTempsAssemblageSoum.Text) Then
155       dblTempsAssemblage = CDbl(txtTempsAssemblageSoum.Text)
160     Else
165       dblTempsAssemblage = 0
170     End If

175     If IsNumeric(txtTempsProgInterfaceSoum.Text) Then
180       dblTempsProgInterface = CDbl(txtTempsProgInterfaceSoum.Text)
185     Else
190       dblTempsProgInterface = 0
195     End If

200     If IsNumeric(txtTempsProgAutomateSoum.Text) Then
205       dblTempsProgAutomate = CDbl(txtTempsProgAutomateSoum.Text)
210     Else
215       dblTempsProgAutomate = 0
220     End If

225     If IsNumeric(txtTempsProgRobotSoum.Text) Then
230       dblTempsProgRobot = CDbl(txtTempsProgRobotSoum.Text)
235     Else
240       dblTempsProgRobot = 0
245     End If

250     If IsNumeric(txtTempsVisionSoum.Text) Then
255       dblTempsVision = CDbl(txtTempsVisionSoum.Text)
260     Else
265       dblTempsVision = 0
270     End If

275     If IsNumeric(txtTempsTestSoum.Text) Then
280       dblTempsTest = CDbl(txtTempsTestSoum.Text)
285     Else
290       dblTempsTest = 0
295     End If

300     If IsNumeric(txtTempsInstallationSoum.Text) Then
305       dblTempsInstallation = CDbl(txtTempsInstallationSoum.Text)
310     Else
315       dblTempsInstallation = 0
320     End If

325     If IsNumeric(txtTempsMiseServiceSoum.Text) Then
330       dblTempsMiseService = CDbl(txtTempsMiseServiceSoum.Text)
335     Else
340       dblTempsMiseService = 0
345     End If

350     If IsNumeric(txtTempsFormationSoum.Text) Then
355       dblTempsFormation = CDbl(txtTempsFormationSoum.Text)
360     Else
365       dblTempsFormation = 0
370     End If

375     If IsNumeric(txtTempsGestionSoum.Text) Then
380       dblTempsGestion = CDbl(txtTempsGestionSoum.Text)
385     Else
390       dblTempsGestion = 0
395     End If

400     If IsNumeric(txtTempsShippingSoum.Text) Then
405       dblTempsShipping = CDbl(txtTempsShippingSoum.Text)
410     Else
415       dblTempsShipping = 0
420     End If


        If IsNumeric(txtTempsprototypeSoum.Text) Then
421       dblTempsPrototype = CDbl(txtTempsprototypeSoum.Text)
422     Else
423       dblTempsPrototype = 0
424     End If

425     dblTotalTemps = dblTempsDessin + _
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

430     lblTotalTempsRHSoum.Caption = Conversion(dblTotalTemps, MODE_DECIMAL)

435     Exit Sub

AfficherErreur:

440     woups "frmProjSoumElecTemps", "CalculerTotalTemps", Err, Erl
End Sub

Private Sub lblTempsFabricationSoum_Click()
        'Active ou désactive le temps des pièces
  
5       On Error GoTo AfficherErreur
        
10      If m_eMode = MODE_AJOUT_MODIF Then
15        If m_bSansTemps = True Then
20          Croix1.Visible = False
25          Croix2.Visible = False

30          m_bSansTemps = False
35        Else
40          Croix1.Visible = True
45          Croix2.Visible = True

50          m_bSansTemps = True
55        End If

60        Call lblTempsFabricationSoum_Change
65      End If

70      Call CalculerTotal

75      Exit Sub

AfficherErreur:

80      woups "frmProjSoumElecTemps", "lblTempsMécanique_Click", Err, Erl
End Sub

Private Function CalculerTempsFabrication() As String

5       On Error GoTo AfficherErreur

10      Dim dblTempsFab As Double
15      Dim iCompteur   As Integer

        'Pour chaque élément du listView
20      For iCompteur = 1 To FrmProjSoumElec.lvwSoumission.ListItems.count
25        If Trim$(FrmProjSoumElec.lvwSoumission.ListItems(iCompteur).SubItems(9)) <> vbNullString Then
            'On additionne le temps
30          dblTempsFab = dblTempsFab + CDbl(Replace(Trim$(FrmProjSoumElec.lvwSoumission.ListItems(iCompteur).SubItems(9)), ".", ","))
35        End If
40      Next
        
45      CalculerTempsFabrication = Replace(dblTempsFab / 10, ".", ",")

50      Exit Function

AfficherErreur:

55      woups "frmProjSoumElecTemps", "CalculerTempsFabrication", Err, Erl
End Function

