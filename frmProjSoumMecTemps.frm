VERSION 5.00
Begin VB.Form frmProjSoumMecTemps 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Temps"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmProjSoumMecTemps.frx":0000
   ScaleHeight     =   9180
   ScaleWidth      =   13155
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
      Height          =   5415
      Left            =   120
      TabIndex        =   41
      Top             =   840
      Width           =   7935
      Begin VB.TextBox txtTempsPrototypeConc 
         Height          =   285
         Left            =   4680
         TabIndex        =   141
         Top             =   4800
         Width           =   735
      End
      Begin VB.TextBox txtTempsPrototypeSoum 
         Height          =   285
         Left            =   3000
         TabIndex        =   140
         Top             =   4800
         Width           =   735
      End
      Begin VB.TextBox txtTempsPrototypeProj 
         Height          =   285
         Left            =   3840
         TabIndex        =   139
         Top             =   4800
         Width           =   735
      End
      Begin VB.TextBox txtTempsShippingProj 
         Height          =   285
         Left            =   3840
         TabIndex        =   133
         Top             =   4440
         Width           =   735
      End
      Begin VB.TextBox txtTempsShippingSoum 
         Height          =   285
         Left            =   3000
         TabIndex        =   132
         Top             =   4440
         Width           =   735
      End
      Begin VB.TextBox txtTempsShippingConc 
         Height          =   285
         Left            =   4680
         TabIndex        =   131
         Top             =   4440
         Width           =   735
      End
      Begin VB.TextBox txtTempsGestionConc 
         Height          =   285
         Left            =   4680
         TabIndex        =   29
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox txtTempsGestionSoum 
         Height          =   285
         Left            =   3000
         TabIndex        =   9
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox txtTempsGestionProj 
         Height          =   285
         Left            =   3840
         TabIndex        =   19
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox txtTempsInstallationConc 
         Height          =   285
         Left            =   4680
         TabIndex        =   27
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox txtTempsFormationConc 
         Height          =   285
         Left            =   4680
         TabIndex        =   28
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox txtTempsDessinConc 
         Height          =   285
         Left            =   4680
         TabIndex        =   20
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtTempsTestConc 
         Height          =   285
         Left            =   4680
         TabIndex        =   26
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtTempsPeintureConc 
         Height          =   285
         Left            =   4680
         TabIndex        =   25
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtTempsAssemblageConc 
         Height          =   285
         Left            =   4680
         TabIndex        =   24
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtTempsSoudureConc 
         Height          =   285
         Left            =   4680
         TabIndex        =   23
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtTempsCoupeConc 
         Height          =   285
         Left            =   4680
         TabIndex        =   21
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtTempsMachinageConc 
         Height          =   285
         Left            =   4680
         TabIndex        =   22
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtTempsMachinageProj 
         Height          =   285
         Left            =   3840
         TabIndex        =   12
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtTempsCoupeProj 
         Height          =   285
         Left            =   3840
         TabIndex        =   11
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtTempsSoudureProj 
         Height          =   285
         Left            =   3840
         TabIndex        =   13
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtTempsAssemblageProj 
         Height          =   285
         Left            =   3840
         TabIndex        =   14
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtTempsPeintureProj 
         Height          =   285
         Left            =   3840
         TabIndex        =   15
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtTempsTestProj 
         Height          =   285
         Left            =   3840
         TabIndex        =   16
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtTempsDessinProj 
         Height          =   285
         Left            =   3840
         TabIndex        =   10
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtTempsFormationProj 
         Height          =   285
         Left            =   3840
         TabIndex        =   18
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox txtTempsInstallationProj 
         Height          =   285
         Left            =   3840
         TabIndex        =   17
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox txtTempsInstallationSoum 
         Height          =   285
         Left            =   3000
         TabIndex        =   7
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox txtTempsFormationSoum 
         Height          =   285
         Left            =   3000
         TabIndex        =   8
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox txtTempsDessinSoum 
         Height          =   285
         Left            =   3000
         TabIndex        =   0
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtTempsTestSoum 
         Height          =   285
         Left            =   3000
         TabIndex        =   6
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtTempsPeintureSoum 
         Height          =   285
         Left            =   3000
         TabIndex        =   5
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtTempsAssemblageSoum 
         Height          =   285
         Left            =   3000
         TabIndex        =   4
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtTempsSoudureSoum 
         Height          =   285
         Left            =   3000
         TabIndex        =   3
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtTempsCoupeSoum 
         Height          =   285
         Left            =   3000
         TabIndex        =   1
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtTempsMachinageSoum 
         Height          =   285
         Left            =   3000
         TabIndex        =   2
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblPrixPrototype 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   6840
         TabIndex        =   146
         Top             =   4800
         Width           =   735
      End
      Begin VB.Label lblTempsPrototypeReel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5520
         TabIndex        =   145
         Top             =   4800
         Width           =   735
      End
      Begin VB.Label Label60 
         BackStyle       =   0  'Transparent
         Caption         =   "Prototypage-Développement :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   144
         Top             =   4800
         Width           =   2055
      End
      Begin VB.Label Label51 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6360
         TabIndex        =   143
         Top             =   4800
         Width           =   255
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7680
         TabIndex        =   142
         Top             =   4800
         Width           =   135
      End
      Begin VB.Label Label59 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7680
         TabIndex        =   138
         Top             =   4440
         Width           =   135
      End
      Begin VB.Label Label55 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6360
         TabIndex        =   137
         Top             =   4440
         Width           =   255
      End
      Begin VB.Label Label53 
         BackStyle       =   0  'Transparent
         Caption         =   "Expédition :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   136
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label lblTempsShippingReel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5520
         TabIndex        =   135
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label lblPrixShipping 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   6840
         TabIndex        =   134
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label lblPrixInstallation 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   6840
         TabIndex        =   121
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label lblPrixFormation 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   6840
         TabIndex        =   120
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label lblPrixDessin 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   6840
         TabIndex        =   119
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblPrixGestion 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   6840
         TabIndex        =   118
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label lblPrixTest 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   6840
         TabIndex        =   117
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label lblPrixPeinture 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   6840
         TabIndex        =   116
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label lblPrixAssemblage 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   6840
         TabIndex        =   115
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label lblPrixSoudure 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   6840
         TabIndex        =   114
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label lblPrixCoupe 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   6840
         TabIndex        =   113
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblPrixMachinage 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   6840
         TabIndex        =   112
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblTempsGestionReel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5520
         TabIndex        =   111
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label67 
         BackStyle       =   0  'Transparent
         Caption         =   "Gestion de projets :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   110
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label Label66 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6360
         TabIndex        =   109
         Top             =   4080
         Width           =   255
      End
      Begin VB.Label Label65 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7680
         TabIndex        =   108
         Top             =   4080
         Width           =   135
      End
      Begin VB.Label lblTempsInstallationReel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5520
         TabIndex        =   107
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label lblTempsFormationReel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5520
         TabIndex        =   106
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label lblTempsDessinReel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5520
         TabIndex        =   105
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblTempsTestReel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5520
         TabIndex        =   104
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label lblTempsPeintureReel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5520
         TabIndex        =   103
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label lblTempsAssemblageReel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5520
         TabIndex        =   102
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label lblTempsSoudureReel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5520
         TabIndex        =   101
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label lblTempsCoupeReel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5520
         TabIndex        =   100
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblTempsMachinageReel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5520
         TabIndex        =   99
         Top             =   1560
         Width           =   735
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
         Left            =   5400
         TabIndex        =   98
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Conception"
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
         Left            =   4440
         TabIndex        =   97
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label56 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Projet"
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
         Left            =   3720
         TabIndex        =   96
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label54 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Soumission"
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
         Left            =   2760
         TabIndex        =   95
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Caption         =   "Installation :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6360
         TabIndex        =   69
         Top             =   3360
         Width           =   255
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7680
         TabIndex        =   70
         Top             =   3360
         Width           =   135
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "Formation du personnel :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6360
         TabIndex        =   66
         Top             =   3720
         Width           =   255
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7680
         TabIndex        =   67
         Top             =   3720
         Width           =   135
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7680
         TabIndex        =   64
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7680
         TabIndex        =   61
         Top             =   3000
         Width           =   135
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7680
         TabIndex        =   58
         Top             =   2640
         Width           =   135
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7680
         TabIndex        =   55
         Top             =   2280
         Width           =   135
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7680
         TabIndex        =   52
         Top             =   1920
         Width           =   135
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7680
         TabIndex        =   49
         Top             =   1200
         Width           =   135
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7680
         TabIndex        =   46
         Top             =   1560
         Width           =   135
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6360
         TabIndex        =   63
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6360
         TabIndex        =   60
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6360
         TabIndex        =   57
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6360
         TabIndex        =   54
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6360
         TabIndex        =   51
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6360
         TabIndex        =   48
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6360
         TabIndex        =   45
         Top             =   1560
         Width           =   255
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
         Left            =   6840
         TabIndex        =   43
         Top             =   120
         Width           =   735
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
         Left            =   3000
         TabIndex        =   42
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Conception et dessins :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Tests finaux :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Assemblage des systèmes :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Peinture et finition :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Coupe, soudure et meulage :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Coupe et préparation (sauf soudage) :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Machinage :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1560
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "Détails"
      Height          =   375
      Left            =   5640
      TabIndex        =   32
      Top             =   7080
      Width           =   735
   End
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   11760
      TabIndex        =   40
      Top             =   6360
      Width           =   1215
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
      Left            =   8160
      TabIndex        =   86
      Top             =   2880
      Width           =   4815
      Begin VB.TextBox txtPrixEmballage 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3720
         TabIndex        =   38
         Text            =   "0"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4560
         TabIndex        =   89
         Top             =   480
         Width           =   135
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
         TabIndex        =   87
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Frais de transport / emballage :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   88
         Top             =   480
         Width           =   2295
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
      Left            =   8160
      TabIndex        =   71
      Top             =   840
      Width           =   4815
      Begin VB.TextBox txtNbrePersonne 
         Height          =   285
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   33
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtTempsDeplacement 
         Height          =   285
         Left            =   2400
         TabIndex        =   36
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtTempsHebergement 
         Height          =   285
         Left            =   2400
         TabIndex        =   34
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtTempsRepas 
         Height          =   285
         Left            =   2400
         TabIndex        =   35
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtTempsUniteMobile 
         Height          =   285
         Left            =   2400
         TabIndex        =   37
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblPrixUniteMobile 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   3720
         TabIndex        =   130
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblPrixDeplacement 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   3720
         TabIndex        =   129
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblPrixRepas 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   3720
         TabIndex        =   128
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblPrixHebergement 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   3720
         TabIndex        =   127
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "pers."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   74
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "Hébergement :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "Repas :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Transport / déplacement : "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   "Transport de l'unité mobile :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   1560
         Width           =   2055
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
         TabIndex        =   72
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Jours"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   75
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Jours"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   78
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Km"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   81
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Km"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   84
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4560
         TabIndex        =   76
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4560
         TabIndex        =   79
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4560
         TabIndex        =   82
         Top             =   1200
         Width           =   135
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4560
         TabIndex        =   85
         Top             =   1560
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdUnlock 
      Height          =   615
      Left            =   10800
      Picture         =   "frmProjSoumMecTemps.frx":2F0D
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7920
      Width           =   735
   End
   Begin VB.CommandButton cmdLock 
      Height          =   615
      Left            =   10800
      Picture         =   "frmProjSoumMecTemps.frx":334F
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7920
      Width           =   735
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   285
      Left            =   11880
      TabIndex        =   39
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label lblTotalPrixRH 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   285
      Left            =   6960
      TabIndex        =   126
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label lblTotalTempsRHProj 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3960
      TabIndex        =   125
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label lblTotalTempsRHConc 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4800
      TabIndex        =   124
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label lblTotalTempsRHReel 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5640
      TabIndex        =   123
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label lblTotalTempsRHSoum 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3120
      TabIndex        =   122
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label Label52 
      BackStyle       =   0  'Transparent
      Caption         =   "Total de la ressource humaine :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   92
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label lblDollarRH 
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7800
      TabIndex        =   94
      Top             =   6600
      Width           =   135
   End
   Begin VB.Label Label50 
      BackStyle       =   0  'Transparent
      Caption         =   "Hrs"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6480
      TabIndex        =   93
      Top             =   6600
      Width           =   255
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12720
      TabIndex        =   91
      Top             =   3960
      Width           =   135
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11040
      TabIndex        =   90
      Top             =   3960
      Width           =   615
   End
End
Attribute VB_Name = "frmProjSoumMecTemps"
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

Private m_sTauxDessin             As String
Private m_sTauxCoupe              As String
Private m_sTauxMachinage          As String
Private m_sTauxSoudure            As String
Private m_sTauxAssemblage         As String
Private m_sTauxPeinture           As String
Private m_sTauxTest               As String
Private m_sTauxInstallation       As String
Private m_sTauxFormation          As String
Private m_sTauxGestion            As String
Private m_sTauxShipping           As String
Private m_sTauxPrototype           As String

Private m_sRepas                  As String
Private m_sHebergement1           As String
Private m_sHebergement2           As String
Private m_sStandard               As String
Private m_sUniteMobile            As String

Private m_sTempsDessinAvant       As String
Private m_sTempsCoupeAvant        As String
Private m_sTempsMachinageAvant    As String
Private m_sTempsSoudureAvant      As String
Private m_sTempsAssemblageAvant   As String
Private m_sTempsPeintureAvant     As String
Private m_sTempsTestAvant         As String
Private m_sTempsInstallationAvant As String
Private m_sTempsFormationAvant    As String
Private m_sTempsGestionAvant      As String
Private m_sTempsShippingAvant     As String
Private m_sTempsPrototypeAvant     As String
Private m_sTempsTotalRHAvant      As String

Private m_sNoProjet               As String
Private m_sNoSoumission           As String

Private m_eType                   As enumType

Private m_eMode                   As enumMode
 
Private m_bNouveauTaux            As Boolean 'Pour savoir si les nouveaux taux doivent être pris
Private m_bLocked                 As Boolean 'Pour savoir si la section projet est barrée ou non

Public Sub Afficher(ByVal sNoProjet As String, ByVal sNoSoumission As String, ByVal iType As Integer, ByVal iMode As Integer, ByVal bNouveauTaux As Boolean)

5       On Error GoTo AfficherErreur
  
10      m_eType = iType
    
15      m_eMode = iMode
    
20      m_sNoProjet = sNoProjet
25      m_sNoSoumission = sNoSoumission
  
30      m_bNouveauTaux = bNouveauTaux
  
35      If bNouveauTaux = True Then
40        Call InitialiserVariablesConfig
45      Else
50        Call InitialiserVariablesProjSoum
55      End If
     
60      Call AfficherEnregistrement
  
65      Call RemplirValeursAvant

70      If m_eMode = MODE_AJOUT_MODIF Then
75        Call BarrerChamps(False)
80      Else
85        Call BarrerChamps(True)
90      End If
    
95      Call Me.Show(vbModal)

100     Exit Sub

AfficherErreur:

105     woups "frmProjSoumMecTemps", "Afficher", Err, Erl
End Sub

Private Sub RemplirValeursAvant()
        
5       On Error GoTo AfficherErreur

10      m_sTempsDessinAvant = txtTempsDessinProj.Text
15      m_sTempsCoupeAvant = txtTempsCoupeProj.Text
20      m_sTempsMachinageAvant = txtTempsMachinageProj.Text
25      m_sTempsSoudureAvant = txtTempsSoudureProj.Text
30      m_sTempsAssemblageAvant = txtTempsAssemblageProj.Text
35      m_sTempsPeintureAvant = txtTempsPeintureProj.Text
40      m_sTempsTestAvant = txtTempsTestProj.Text
45      m_sTempsInstallationAvant = txtTempsInstallationProj.Text
50      m_sTempsFormationAvant = txtTempsFormationProj.Text
55      m_sTempsGestionAvant = txtTempsGestionProj.Text
60      m_sTempsShippingAvant = txtTempsShippingProj.Text
61      m_sTempsPrototypeAvant = txtTempsPrototypeProj.Text

65      Exit Sub

AfficherErreur:

70      woups "frmProjSoumMecTemps", "RemplirValeursVant", Err, Erl
End Sub

Private Sub AfficherEnregistrement()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset
15      Dim rstSoum     As ADODB.Recordset
20      Dim rstPunch    As ADODB.Recordset
25      Dim sChamps     As String
30      Dim sTable      As String
    
35      If m_eType = TYPE_PROJET Then
40        sChamps = "IDProjet"
45        sTable = "GRB_ProjetMec"
50      Else
55        sChamps = "IDSoumission"
60        sTable = "GRB_SoumissionMec"
65      End If
  
70      Set rstProjSoum = New ADODB.Recordset
  
75      If m_eType = TYPE_PROJET Then
80        Call rstProjSoum.Open("SELECT * FROM " & sTable & " WHERE " & sChamps & " = '" & m_sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
85      Else
90        Call rstProjSoum.Open("SELECT * FROM " & sTable & " WHERE " & sChamps & " = '" & m_sNoSoumission & "'", g_connData, adOpenDynamic, adLockOptimistic)
95      End If
  
100     If Not rstProjSoum.EOF And FrmProjSoumMec.m_bTempsDejaOuvert = False And m_eMode = MODE_INACTIF Then
105       If m_eType = TYPE_SOUMISSION Then
110         If Not IsNull(rstProjSoum.Fields("TempsDessin")) Then
115           txtTempsDessinSoum.Text = rstProjSoum.Fields("TempsDessin")
120         Else
125           txtTempsDessinSoum.Text = "0"
130         End If

135         If Not IsNull(rstProjSoum.Fields("TempsCoupe")) Then
140           txtTempsCoupeSoum.Text = rstProjSoum.Fields("TempsCoupe")
145         Else
150           txtTempsCoupeSoum.Text = "0"
155         End If

160         If Not IsNull(rstProjSoum.Fields("TempsMachinage")) Then
165           txtTempsMachinageSoum.Text = rstProjSoum.Fields("TempsMachinage")
170         Else
175           txtTempsMachinageSoum.Text = "0"
180         End If

185         If Not IsNull(rstProjSoum.Fields("TempsSoudure")) Then
190           txtTempsSoudureSoum.Text = rstProjSoum.Fields("TempsSoudure")
195         Else
200           txtTempsSoudureSoum.Text = "0"
205         End If

210         If Not IsNull(rstProjSoum.Fields("TempsAssemblage")) Then
215           txtTempsAssemblageSoum.Text = rstProjSoum.Fields("TempsAssemblage")
220         Else
225           txtTempsAssemblageSoum.Text = "0"
230         End If

235         If Not IsNull(rstProjSoum.Fields("TempsPeinture")) Then
240           txtTempsPeintureSoum.Text = rstProjSoum.Fields("TempsPeinture")
245         Else
250           txtTempsPeintureSoum.Text = "0"
255         End If

260         If Not IsNull(rstProjSoum.Fields("TempsTest")) Then
265           txtTempsTestSoum.Text = rstProjSoum.Fields("TempsTest")
270         Else
275           txtTempsTestSoum.Text = "0"
280         End If

285         If Not IsNull(rstProjSoum.Fields("TempsInstallation")) Then
290           txtTempsInstallationSoum.Text = rstProjSoum.Fields("TempsInstallation")
295         Else
300           txtTempsInstallationSoum.Text = "0"
305         End If

310         If Not IsNull(rstProjSoum.Fields("TempsFormation")) Then
315           txtTempsFormationSoum.Text = rstProjSoum.Fields("TempsFormation")
320         Else
325           txtTempsFormationSoum.Text = "0"
330         End If

335         If Not IsNull(rstProjSoum.Fields("TempsGestion")) Then
340           txtTempsGestionSoum.Text = rstProjSoum.Fields("TempsGestion")
345         Else
350           txtTempsGestionSoum.Text = "0"
355         End If

360         If Not IsNull(rstProjSoum.Fields("TempsShipping")) Then
365           txtTempsShippingSoum.Text = rstProjSoum.Fields("TempsShipping")
370         Else
375           txtTempsShippingSoum.Text = "0"
380         End If
            txtTempsprototypeSoum.Text = "0"

385       Else
390         If m_sNoSoumission <> "" Then
395           Set rstSoum = New ADODB.Recordset

400           Call rstSoum.Open("SELECT * FROM GRB_SoumissionMec WHERE IDSoumission = '" & m_sNoSoumission & "'", g_connData, adOpenDynamic, adLockOptimistic)

405           If Not rstSoum.EOF Then
410             If Not IsNull(rstSoum.Fields("TempsDessin")) Then
415               txtTempsDessinSoum.Text = rstSoum.Fields("TempsDessin")
420             Else
425               txtTempsDessinSoum.Text = "0"
430             End If

435             If Not IsNull(rstSoum.Fields("TempsCoupe")) Then
440               txtTempsCoupeSoum.Text = rstSoum.Fields("TempsCoupe")
445             Else
450               txtTempsCoupeSoum.Text = "0"
455             End If

460             If Not IsNull(rstSoum.Fields("TempsMachinage")) Then
465               txtTempsMachinageSoum.Text = rstSoum.Fields("TempsMachinage")
470             Else
475               txtTempsMachinageSoum.Text = "0"
480             End If

485             If Not IsNull(rstSoum.Fields("TempsSoudure")) Then
490               txtTempsSoudureSoum.Text = rstSoum.Fields("TempsSoudure")
495             Else
500               txtTempsSoudureSoum.Text = "0"
505             End If

510             If Not IsNull(rstSoum.Fields("TempsAssemblage")) Then
515               txtTempsAssemblageSoum.Text = rstSoum.Fields("TempsAssemblage")
520             Else
525               txtTempsAssemblageSoum.Text = "0"
530             End If

535             If Not IsNull(rstSoum.Fields("TempsPeinture")) Then
540               txtTempsPeintureSoum.Text = rstSoum.Fields("TempsPeinture")
545             Else
550               txtTempsPeintureSoum.Text = "0"
555             End If
  
560             If Not IsNull(rstSoum.Fields("TempsTest")) Then
565               txtTempsTestSoum.Text = rstSoum.Fields("TempsTest")
570             Else
575               txtTempsTestSoum.Text = "0"
580             End If

585             If Not IsNull(rstSoum.Fields("TempsInstallation")) Then
590               txtTempsInstallationSoum.Text = rstSoum.Fields("TempsInstallation")
595             Else
600               txtTempsInstallationSoum.Text = "0"
605             End If

610             If Not IsNull(rstSoum.Fields("TempsFormation")) Then
615               txtTempsFormationSoum.Text = rstSoum.Fields("TempsFormation")
620             Else
625               txtTempsFormationSoum.Text = "0"
630             End If

635             If Not IsNull(rstSoum.Fields("TempsGestion")) Then
640               txtTempsGestionSoum.Text = rstSoum.Fields("TempsGestion")
645             Else
650               txtTempsGestionSoum.Text = "0"
655             End If

660             If Not IsNull(rstSoum.Fields("TempsShipping")) Then
665               txtTempsShippingSoum.Text = rstSoum.Fields("TempsShipping")
670             Else
675               txtTempsShippingSoum.Text = "0"
680             End If
                txtTempsprototypeSoum.Text = "0"

685           Else
690             txtTempsDessinSoum.Text = 0
695             txtTempsCoupeSoum.Text = 0
700             txtTempsMachinageSoum.Text = 0
705             txtTempsSoudureSoum.Text = 0
710             txtTempsAssemblageSoum.Text = 0
715             txtTempsPeintureSoum.Text = 0
720             txtTempsTestSoum.Text = 0
725             txtTempsInstallationSoum.Text = 0
730             txtTempsFormationSoum.Text = 0
735             txtTempsGestionSoum.Text = 0
740             txtTempsShippingSoum.Text = 0
                txtTempsprototypeSoum.Text = 0

745           End If

750           Call rstSoum.Close
755           Set rstSoum = Nothing
760         Else
765           txtTempsDessinSoum.Text = 0
770           txtTempsCoupeSoum.Text = 0
775           txtTempsMachinageSoum.Text = 0
780           txtTempsSoudureSoum.Text = 0
785           txtTempsAssemblageSoum.Text = 0
790           txtTempsPeintureSoum.Text = 0
795           txtTempsTestSoum.Text = 0
800           txtTempsInstallationSoum.Text = 0
805           txtTempsFormationSoum.Text = 0
810           txtTempsGestionSoum.Text = 0
815           txtTempsShippingSoum.Text = 0
                txtTempsprototypeSoum.Text = 0
820         End If

825         m_bLocked = rstProjSoum.Fields("TempsProjBarré")

830         If Not IsNull(rstProjSoum.Fields("TempsDessinProj")) Then
835           txtTempsDessinProj.Text = rstProjSoum.Fields("TempsDessinProj")
840         Else
845           txtTempsDessinProj.Text = "0"
850         End If

855         If Not IsNull(rstProjSoum.Fields("TempsCoupeProj")) Then
860           txtTempsCoupeProj.Text = rstProjSoum.Fields("TempsCoupeProj")
865         Else
870           txtTempsCoupeProj.Text = "0"
875         End If

880         If Not IsNull(rstProjSoum.Fields("TempsMachinageProj")) Then
885           txtTempsMachinageProj.Text = rstProjSoum.Fields("TempsMachinageProj")
890         Else
895           txtTempsMachinageProj.Text = "0"
900         End If

905         If Not IsNull(rstProjSoum.Fields("TempsSoudureProj")) Then
910           txtTempsSoudureProj.Text = rstProjSoum.Fields("TempsSoudureProj")
915         Else
920           txtTempsSoudureProj.Text = "0"
925         End If

930         If Not IsNull(rstProjSoum.Fields("TempsAssemblageProj")) Then
935           txtTempsAssemblageProj.Text = rstProjSoum.Fields("TempsAssemblageProj")
940         Else
945           txtTempsAssemblageProj.Text = "0"
950         End If

955         If Not IsNull(rstProjSoum.Fields("TempsPeintureProj")) Then
960           txtTempsPeintureProj.Text = rstProjSoum.Fields("TempsPeintureProj")
965         Else
970           txtTempsPeintureProj.Text = "0"
975         End If

980         If Not IsNull(rstProjSoum.Fields("TempsTestProj")) Then
985           txtTempsTestProj.Text = rstProjSoum.Fields("TempsTestProj")
990         Else
995           txtTempsTestProj.Text = "0"
1000        End If

1005        If Not IsNull(rstProjSoum.Fields("TempsInstallationProj")) Then
1010          txtTempsInstallationProj.Text = rstProjSoum.Fields("TempsInstallationProj")
1015        Else
1020          txtTempsInstallationProj.Text = "0"
1025        End If

1030        If Not IsNull(rstProjSoum.Fields("TempsFormationProj")) Then
1035          txtTempsFormationProj.Text = rstProjSoum.Fields("TempsFormationProj")
1040        Else
1045          txtTempsFormationProj.Text = "0"
1050        End If

1055        If Not IsNull(rstProjSoum.Fields("TempsGestionProj")) Then
1060          txtTempsGestionProj.Text = rstProjSoum.Fields("TempsGestionProj")
1065        Else
1070          txtTempsGestionProj.Text = "0"
1075        End If

1080        If Not IsNull(rstProjSoum.Fields("TempsShippingProj")) Then
1085          txtTempsShippingProj.Text = rstProjSoum.Fields("TempsShippingProj")
1090        Else
1095          txtTempsShippingProj.Text = "0"
1100        End If

            txtTempsPrototypeProj.Text = "0"

1105        If m_bLocked = False Then
1110          txtTempsDessinConc.Text = vbNullString
1115          txtTempsCoupeConc.Text = vbNullString
1120          txtTempsMachinageConc.Text = vbNullString
1125          txtTempsSoudureConc.Text = vbNullString
1130          txtTempsAssemblageConc.Text = vbNullString
1135          txtTempsPeintureConc.Text = vbNullString
1140          txtTempsTestConc.Text = vbNullString
1145          txtTempsInstallationConc.Text = vbNullString
1150          txtTempsFormationConc.Text = vbNullString
1155          txtTempsGestionConc.Text = vbNullString
1160          txtTempsShippingConc.Text = vbNullString
                txtTempsPrototypeConc.Text = vbNullString
1165        Else
1170          If Not IsNull(rstProjSoum.Fields("TempsDessinConc")) Then
1175            txtTempsDessinConc.Text = rstProjSoum.Fields("TempsDessinConc")
1180          Else
1185            txtTempsDessinConc.Text = "0"
1190          End If

1195          If Not IsNull(rstProjSoum.Fields("TempsCoupeConc")) Then
1200            txtTempsCoupeConc.Text = rstProjSoum.Fields("TempsCoupeConc")
1205          Else
1210            txtTempsCoupeConc.Text = "0"
1215          End If

1220          If Not IsNull(rstProjSoum.Fields("TempsMachinageConc")) Then
1225            txtTempsMachinageConc.Text = rstProjSoum.Fields("TempsMachinageConc")
1230          Else
1235            txtTempsMachinageConc.Text = "0"
1240          End If

1245          If Not IsNull(rstProjSoum.Fields("TempsSoudureConc")) Then
1250            txtTempsSoudureConc.Text = rstProjSoum.Fields("TempsSoudureConc")
1255          Else
1260            txtTempsSoudureConc.Text = "0"
1265          End If

1270          If Not IsNull(rstProjSoum.Fields("TempsAssemblageConc")) Then
1275            txtTempsAssemblageConc.Text = rstProjSoum.Fields("TempsAssemblageConc")
1280          Else
1285            txtTempsAssemblageConc.Text = "0"
1290          End If

1295          If Not IsNull(rstProjSoum.Fields("TempsPeintureConc")) Then
1300            txtTempsPeintureConc.Text = rstProjSoum.Fields("TempsPeintureConc")
1305          Else
1310            txtTempsPeintureConc.Text = "0"
1315          End If

1320          If Not IsNull(rstProjSoum.Fields("TempsTestConc")) Then
1325            txtTempsTestConc.Text = rstProjSoum.Fields("TempsTestConc")
1330          Else
1335            txtTempsTestConc.Text = "0"
1340          End If
  
1345          If Not IsNull(rstProjSoum.Fields("TempsInstallationConc")) Then
1350            txtTempsInstallationConc.Text = rstProjSoum.Fields("TempsInstallationConc")
1355          Else
1360            txtTempsInstallationConc.Text = "0"
1365          End If

1370          If Not IsNull(rstProjSoum.Fields("TempsFormationConc")) Then
1375            txtTempsFormationConc.Text = rstProjSoum.Fields("TempsFormationConc")
1380          Else
1385            txtTempsFormationConc.Text = "0"
1390          End If

1395          If Not IsNull(rstProjSoum.Fields("TempsGestionConc")) Then
1400            txtTempsGestionConc.Text = rstProjSoum.Fields("TempsGestionConc")
1405          Else
1410            txtTempsGestionConc.Text = "0"
1415          End If

1420          If Not IsNull(rstProjSoum.Fields("TempsShippingConc")) Then
1425            txtTempsShippingConc.Text = rstProjSoum.Fields("TempsShippingConc")
1430          Else
1435            txtTempsShippingConc.Text = "0"
1440          End If
                txtTempsPrototypeConc.Text = "0"

1445        End If
1450      End If

1455      If m_eType = TYPE_PROJET Then
1460        Call AfficherTempsReels

1465        Call CalculerTotalArgent
1470      End If

1475      If Not IsNull(rstProjSoum.Fields("NbrePersonne")) Then
1480        txtNbrePersonne.Text = rstProjSoum.Fields("NbrePersonne")
1485      Else
1490        txtNbrePersonne.Text = "0"
1495      End If

1500      If Not IsNull(rstProjSoum.Fields("TempsHebergement")) Then
1505        txtTempsHebergement.Text = rstProjSoum.Fields("TempsHebergement")
1510      Else
1515        txtTempsHebergement.Text = "0"
1520      End If

1525      If Not IsNull(rstProjSoum.Fields("TempsRepas")) Then
1530        txtTempsRepas.Text = rstProjSoum.Fields("TempsRepas")
1535      Else
1540        txtTempsRepas.Text = "0"
1545      End If

1550      If Not IsNull(rstProjSoum.Fields("TempsTransport")) Then
1555        txtTempsDeplacement.Text = rstProjSoum.Fields("TempsTransport")
1560      Else
1565        txtTempsDeplacement.Text = "0"
1570      End If

1575      If Not IsNull(rstProjSoum.Fields("TempsUniteMobile")) Then
1580        txtTempsUniteMobile.Text = rstProjSoum.Fields("TempsUniteMobile")
1585      Else
1590        txtTempsUniteMobile.Text = "0"
1595      End If

1600      txtPrixEmballage.Text = rstProjSoum.Fields("PrixEmballage")
1605    Else
1610      If m_eType = TYPE_SOUMISSION Then
1615        txtTempsDessinSoum.Text = FrmProjSoumMec.m_sTempsDessin
1620        txtTempsCoupeSoum.Text = FrmProjSoumMec.m_sTempsCoupe
1625        txtTempsMachinageSoum.Text = FrmProjSoumMec.m_sTempsMachinage
1630        txtTempsSoudureSoum.Text = FrmProjSoumMec.m_sTempsSoudure
1635        txtTempsAssemblageSoum.Text = FrmProjSoumMec.m_sTempsAssemblage
1640        txtTempsPeintureSoum.Text = FrmProjSoumMec.m_sTempsPeinture
1645        txtTempsTestSoum.Text = FrmProjSoumMec.m_sTempsTest
1650        txtTempsInstallationSoum.Text = FrmProjSoumMec.m_sTempsInstallation
1655        txtTempsFormationSoum.Text = FrmProjSoumMec.m_sTempsFormation
1660        txtTempsGestionSoum.Text = FrmProjSoumMec.m_sTempsGestion
1665        txtTempsShippingSoum.Text = FrmProjSoumMec.m_sTempsShipping
1670      Else
1675        If m_sNoSoumission <> "" Then
1680          Set rstSoum = New ADODB.Recordset

1685          Call rstSoum.Open("SELECT * FROM GRB_SoumissionMec WHERE IDSoumission = '" & m_sNoSoumission & "'", g_connData, adOpenDynamic, adLockOptimistic)

1690          If Not rstSoum.EOF Then
1695            If Not IsNull(rstSoum.Fields("TempsDessin")) Then
1700              txtTempsDessinSoum.Text = rstSoum.Fields("TempsDessin")
1705            Else
1710              txtTempsDessinSoum.Text = "0"
1715            End If

1720            If Not IsNull(rstSoum.Fields("TempsCoupe")) Then
1725              txtTempsCoupeSoum.Text = rstSoum.Fields("TempsCoupe")
1730            Else
1735              txtTempsCoupeSoum.Text = "0"
1740            End If

1745            If Not IsNull(rstSoum.Fields("TempsMachinage")) Then
1750              txtTempsMachinageSoum.Text = rstSoum.Fields("TempsMachinage")
1755            Else
1760              txtTempsMachinageSoum.Text = "0"
1765            End If

1770            If Not IsNull(rstSoum.Fields("TempsSoudure")) Then
1775              txtTempsSoudureSoum.Text = rstSoum.Fields("TempsSoudure")
1780            Else
1785              txtTempsSoudureSoum.Text = "0"
1790            End If

1795            If Not IsNull(rstSoum.Fields("TempsAssemblage")) Then
1800              txtTempsAssemblageSoum.Text = rstSoum.Fields("TempsAssemblage")
1805            Else
1810              txtTempsAssemblageSoum.Text = "0"
1815            End If

1820            If Not IsNull(rstSoum.Fields("TempsPeinture")) Then
1825              txtTempsPeintureSoum.Text = rstSoum.Fields("TempsPeinture")
1830            Else
1835              txtTempsPeintureSoum.Text = "0"
1840            End If

1845            If Not IsNull(rstSoum.Fields("TempsTest")) Then
1850              txtTempsTestSoum.Text = rstSoum.Fields("TempsTest")
1855            Else
1860              txtTempsTestSoum.Text = "0"
1865            End If
  
1870            If Not IsNull(rstSoum.Fields("TempsInstallation")) Then
1875              txtTempsInstallationSoum.Text = rstSoum.Fields("TempsInstallation")
1880            Else
1885              txtTempsInstallationSoum.Text = "0"
1890            End If

1895            If Not IsNull(rstSoum.Fields("TempsFormation")) Then
1900              txtTempsFormationSoum.Text = rstSoum.Fields("TempsFormation")
1905            Else
1910              txtTempsFormationSoum.Text = "0"
1915            End If

1920            If Not IsNull(rstSoum.Fields("TempsGestion")) Then
1925              txtTempsGestionSoum.Text = rstSoum.Fields("TempsGestion")
1930            Else
1935              txtTempsGestionSoum.Text = "0"
1940            End If

1945            If Not IsNull(rstSoum.Fields("TempsShipping")) Then
1950              txtTempsShippingSoum.Text = rstSoum.Fields("TempsShipping")
1955            Else
1960              txtTempsShippingSoum.Text = "0"
1965            End If
                txtTempsprototypeSoum.Text = "0"
1970          Else
1975            txtTempsDessinSoum.Text = 0
1980            txtTempsCoupeSoum.Text = 0
1985            txtTempsMachinageSoum.Text = 0
1990            txtTempsSoudureSoum.Text = 0
1995            txtTempsAssemblageSoum.Text = 0
2000            txtTempsPeintureSoum.Text = 0
2005            txtTempsTestSoum.Text = 0
2010            txtTempsInstallationSoum.Text = 0
2015            txtTempsFormationSoum.Text = 0
2020            txtTempsGestionSoum.Text = 0
2025            txtTempsShippingSoum.Text = 0
                txtTempsprototypeSoum.Text = 0
2030          End If

2035          Call rstSoum.Close
2040          Set rstSoum = Nothing
2045        Else
2050          txtTempsDessinSoum.Text = 0
2055          txtTempsCoupeSoum.Text = 0
2060          txtTempsMachinageSoum.Text = 0
2065          txtTempsSoudureSoum.Text = 0
2070          txtTempsAssemblageSoum.Text = 0
2075          txtTempsPeintureSoum.Text = 0
2080          txtTempsTestSoum.Text = 0
2085          txtTempsInstallationSoum.Text = 0
2090          txtTempsFormationSoum.Text = 0
2095          txtTempsGestionSoum.Text = 0
2100          txtTempsShippingSoum.Text = 0
                txtTempsprototypeSoum.Text = 0
2105        End If

2110        m_bLocked = FrmProjSoumMec.m_bTempsProjLock

2115        If m_bLocked = True Then
2120          If Not IsNull(rstProjSoum.Fields("TempsDessinProj")) Then
2125            txtTempsDessinProj.Text = rstProjSoum.Fields("TempsDessinProj")
2130          Else
2135            txtTempsDessinProj.Text = "0"
2140          End If

2145          If Not IsNull(rstProjSoum.Fields("TempsCoupeProj")) Then
2150            txtTempsCoupeProj.Text = rstProjSoum.Fields("TempsCoupeProj")
2155          Else
2160            txtTempsCoupeProj.Text = "0"
2165          End If

2170          If Not IsNull(rstProjSoum.Fields("TempsMachinageProj")) Then
2175            txtTempsMachinageProj.Text = rstProjSoum.Fields("TempsMachinageProj")
2180          Else
2185            txtTempsMachinageProj.Text = "0"
2190          End If

2195          If Not IsNull(rstProjSoum.Fields("TempsSoudureProj")) Then
2200            txtTempsSoudureProj.Text = rstProjSoum.Fields("TempsSoudureProj")
2205          Else
2210            txtTempsSoudureProj.Text = "0"
2215          End If

2220          If Not IsNull(rstProjSoum.Fields("TempsAssemblageProj")) Then
2225            txtTempsAssemblageProj.Text = rstProjSoum.Fields("TempsAssemblageProj")
2230          Else
2235            txtTempsAssemblageProj.Text = "0"
2240          End If

2245          If Not IsNull(rstProjSoum.Fields("TempsPeintureProj")) Then
2250            txtTempsPeintureProj.Text = rstProjSoum.Fields("TempsPeintureProj")
2255          Else
2260            txtTempsPeintureProj.Text = "0"
2265          End If

2270          If Not IsNull(rstProjSoum.Fields("TempsTestProj")) Then
2275            txtTempsTestProj.Text = rstProjSoum.Fields("TempsTestProj")
2280          Else
2285            txtTempsTestProj.Text = "0"
2290          End If

2295          If Not IsNull(rstProjSoum.Fields("TempsInstallationProj")) Then
2300            txtTempsInstallationProj.Text = rstProjSoum.Fields("TempsInstallationProj")
2305          Else
2310            txtTempsInstallationProj.Text = "0"
2315          End If

2320          If Not IsNull(rstProjSoum.Fields("TempsFormationProj")) Then
2325            txtTempsFormationProj.Text = rstProjSoum.Fields("TempsFormationProj")
2330          Else
2335            txtTempsFormationProj.Text = "0"
2340          End If

2345          If Not IsNull(rstProjSoum.Fields("TempsGestionProj")) Then
2350            txtTempsGestionProj.Text = rstProjSoum.Fields("TempsGestionProj")
2355          Else
2360            txtTempsGestionProj.Text = "0"
2365          End If

2370          If Not IsNull(rstProjSoum.Fields("TempsShippingProj")) Then
2375            txtTempsShippingProj.Text = rstProjSoum.Fields("TempsShippingProj")
2380          Else
2385            txtTempsShippingProj.Text = "0"
2390          End If

                txtTempsPrototypeProj.Text = "0"

2395          txtTempsDessinConc.Text = FrmProjSoumMec.m_sTempsDessinConc
2400          txtTempsCoupeConc.Text = FrmProjSoumMec.m_sTempsCoupeConc
2405          txtTempsMachinageConc.Text = FrmProjSoumMec.m_sTempsMachinageConc
2410          txtTempsSoudureConc.Text = FrmProjSoumMec.m_sTempsSoudureConc
2415          txtTempsAssemblageConc.Text = FrmProjSoumMec.m_sTempsAssemblageConc
2420          txtTempsPeintureConc.Text = FrmProjSoumMec.m_sTempsPeintureConc
2425          txtTempsTestConc.Text = FrmProjSoumMec.m_sTempsTestConc
2430          txtTempsInstallationConc.Text = FrmProjSoumMec.m_sTempsInstallationConc
2435          txtTempsFormationConc.Text = FrmProjSoumMec.m_sTempsFormationConc
2440          txtTempsGestionConc.Text = FrmProjSoumMec.m_sTempsGestionConc
2445          txtTempsShippingConc.Text = FrmProjSoumMec.m_sTempsShippingConc
                txtTempsPrototypeConc.Text = FrmProjSoumMec.m_sTempsPrototypeConc
2450        Else
2455          txtTempsDessinProj.Text = FrmProjSoumMec.m_sTempsDessinProj
2460          txtTempsCoupeProj.Text = FrmProjSoumMec.m_sTempsCoupeProj
2465          txtTempsMachinageProj.Text = FrmProjSoumMec.m_sTempsMachinageProj
2470          txtTempsSoudureProj.Text = FrmProjSoumMec.m_sTempsSoudureProj
2475          txtTempsAssemblageProj.Text = FrmProjSoumMec.m_sTempsAssemblageProj
2480          txtTempsPeintureProj.Text = FrmProjSoumMec.m_sTempsPeintureProj
2485          txtTempsTestProj.Text = FrmProjSoumMec.m_sTempsTestProj
2490          txtTempsInstallationProj.Text = FrmProjSoumMec.m_sTempsInstallationProj
2495          txtTempsFormationProj.Text = FrmProjSoumMec.m_sTempsFormationProj
2500          txtTempsGestionProj.Text = FrmProjSoumMec.m_sTempsGestionProj
2505          txtTempsShippingProj.Text = FrmProjSoumMec.m_sTempsShippingProj
              txtTempsPrototypeProj.Text = FrmProjSoumMec.m_sTempsPrototypeProj
2510        End If
2515      End If

2520      If m_eType = TYPE_PROJET Then
2525        Call AfficherTempsReels

2530        Call CalculerTotalArgent
2535      End If

2540      txtNbrePersonne.Text = FrmProjSoumMec.m_sNbrePersonne
2545      txtTempsHebergement.Text = FrmProjSoumMec.m_sTempsHebergement
2550      txtTempsRepas.Text = FrmProjSoumMec.m_sTempsRepas
2555      txtTempsDeplacement.Text = FrmProjSoumMec.m_sTempsTransport
2560      txtTempsUniteMobile.Text = FrmProjSoumMec.m_sTempsUniteMobile
2565      txtPrixEmballage.Text = FrmProjSoumMec.m_sPrixEmballage
2570    End If
    
2575    Call rstProjSoum.Close
2580    Set rstProjSoum = Nothing

2585    Exit Sub

AfficherErreur:

2590    woups "frmProjSoumMecTemps", "AfficherEnregistrement", Err, Erl
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

50      If Right$(m_sNoProjet, 2) = "99" Then
55        sFilterNoProjet = "LEFT(NoProjet, 6) = '" & Left$(m_sNoProjet, 6) & "'"
60      Else
65        sFilterNoProjet = "NoProjet = '" & m_sNoProjet & "'"
70      End If

75      Set rstPunch = New ADODB.Recordset

80      Call rstPunch.Open("SELECT Type, " & sTotal & " FROM GRB_Punch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null GROUP BY Type", g_connData, adOpenDynamic, adLockOptimistic)

85      lblTempsDessinReel.Caption = "0"
90      lblTempsCoupeReel.Caption = "0"
95      lblTempsMachinageReel.Caption = "0"
100     lblTempsSoudureReel.Caption = "0"
105     lblTempsAssemblageReel.Caption = "0"
110     lblTempsPeintureReel.Caption = "0"
115     lblTempsTestReel.Caption = "0"
120     lblTempsInstallationReel.Caption = "0"
125     lblTempsFormationReel.Caption = "0"
130     lblTempsGestionReel.Caption = "0"
135     lblTempsShippingReel.Caption = "0"
        lblTempsPrototypeReel.Caption = "0"

140     Do While Not rstPunch.EOF
145       If Not IsNull(rstPunch.Fields("Total")) Then
150         Select Case rstPunch.Fields("Type")
              Case "Dessin":       lblTempsDessinReel.Caption = Round(rstPunch.Fields("Total"), 2)
155           Case "Coupe":        lblTempsCoupeReel.Caption = Round(rstPunch.Fields("Total"), 2)
160           Case "Machinage":    lblTempsMachinageReel.Caption = Round(rstPunch.Fields("Total"), 2)
165           Case "Soudure":      lblTempsSoudureReel.Caption = Round(rstPunch.Fields("Total"), 2)
170           Case "Assemblage":   lblTempsAssemblageReel.Caption = Round(rstPunch.Fields("Total"), 2)
175           Case "Peinture":     lblTempsPeintureReel.Caption = Round(rstPunch.Fields("Total"), 2)
180           Case "Test":         lblTempsTestReel.Caption = Round(rstPunch.Fields("Total"), 2)
185           Case "Installation": lblTempsInstallationReel.Caption = Round(rstPunch.Fields("Total"), 2)
190           Case "Formation":    lblTempsFormationReel.Caption = Round(rstPunch.Fields("Total"), 2)
195           Case "Gestion":      lblTempsGestionReel.Caption = Round(rstPunch.Fields("Total"), 2)
200           Case "Shipping":     lblTempsShippingReel.Caption = Round(rstPunch.Fields("Total"), 2)
              Case "Prototypage-Dévelloppement expérimental":     lblTempsPrototypeReel.Caption = Round(rstPunch.Fields("Total"), 2)
205         End Select
210       End If

215       Call rstPunch.MoveNext
220     Loop

225     Call rstPunch.Close

        'Ouverture des enregistrements avec comme filtre, le numéro du projet
230     Call rstPunch.Open("SELECT " & sTotal & " FROM GRB_Punch WHERE " & sFilterNoProjet & " AND HeureFin is not null AND HeureDébut is not null", g_connData, adOpenDynamic, adLockOptimistic)

235     If Not IsNull(rstPunch.Fields("Total")) Then
240       lblTotalTempsRHReel.Caption = Round(rstPunch.Fields("Total"), 2)
245     Else
250       lblTotalTempsRHReel.Caption = "0"
255     End If

260     Call rstPunch.Close
265     Set rstPunch = Nothing

270     Exit Sub

AfficherErreur:

275     woups "frmProjSoumMecTemps", "AfficherTempsReels", Err, Erl
End Sub

Private Sub CalculerTotalArgent()
  
5       On Error GoTo AfficherErreur

10      If IsNumeric(lblTempsDessinReel.Caption) Then
15        lblPrixDessin.Caption = Round(Replace(lblTempsDessinReel.Caption * m_sTauxDessin, ".", ","), 2)
20      Else
25        lblPrixDessin.Caption = 0
30      End If

35      If IsNumeric(lblTempsCoupeReel.Caption) Then
40        lblPrixCoupe.Caption = Round(Replace(lblTempsCoupeReel.Caption * m_sTauxCoupe, ".", ","), 2)
45      Else
50        lblPrixCoupe.Caption = 0
55      End If

60      If IsNumeric(lblTempsMachinageReel.Caption) Then
65        lblPrixMachinage.Caption = Round(Replace(lblTempsMachinageReel.Caption * m_sTauxMachinage, ".", ","), 2)
70      Else
75        lblPrixMachinage.Caption = 0
80      End If

85      If IsNumeric(lblTempsSoudureReel.Caption) Then
90        lblPrixSoudure.Caption = Round(Replace(lblTempsSoudureReel.Caption * m_sTauxSoudure, ".", ","), 2)
95      Else
100       lblPrixSoudure.Caption = 0
105     End If

110     If IsNumeric(lblTempsAssemblageReel.Caption) Then
115       lblPrixAssemblage.Caption = Round(Replace(lblTempsAssemblageReel.Caption * m_sTauxAssemblage, ".", ","), 2)
120     Else
125       lblPrixAssemblage.Caption = 0
130     End If

135     If IsNumeric(lblTempsPeintureReel.Caption) Then
140       lblPrixPeinture.Caption = Round(Replace(lblTempsPeintureReel.Caption * m_sTauxPeinture, ".", ","), 2)
145     Else
150       lblPrixPeinture.Caption = 0
155     End If

160     If IsNumeric(lblTempsTestReel.Caption) Then
165       lblPrixTest.Caption = Round(Replace(lblTempsTestReel.Caption * m_sTauxTest, ".", ","), 2)
170     Else
175       lblPrixTest.Caption = 0
180     End If

185     If IsNumeric(lblTempsInstallationReel.Caption) Then
190       lblPrixInstallation.Caption = Round(Replace(lblTempsInstallationReel.Caption * m_sTauxInstallation, ".", ","), 2)
195     Else
200       lblPrixInstallation.Caption = 0
205     End If

210     If IsNumeric(lblTempsFormationReel.Caption) Then
215       lblPrixFormation.Caption = Round(Replace(lblTempsFormationReel.Caption * m_sTauxFormation, ".", ","), 2)
220     Else
225       lblPrixFormation.Caption = 0
230     End If

235     If IsNumeric(lblTempsGestionReel.Caption) Then
240       lblPrixGestion.Caption = Round(Replace(lblTempsGestionReel.Caption * m_sTauxGestion, ".", ","), 2)
245     Else
250       lblPrixGestion.Caption = 0
255     End If

260     If IsNumeric(lblTempsShippingReel.Caption) Then
265       lblPrixShipping.Caption = Round(Replace(lblTempsShippingReel.Caption * m_sTauxShipping, ".", ","), 2)
270     Else
275       lblPrixShipping.Caption = 0
280     End If

        If IsNumeric(lblTempsPrototypeReel.Caption) Then
281       lblPrixPrototype.Caption = Round(Replace(lblTempsPrototypeReel.Caption * m_sTauxGestion, ".", ","), 2)
282     Else
283       lblPrixPrototype.Caption = 0
284     End If


285     Call CalculerTotal

290     Exit Sub

AfficherErreur:

295     woups "frmProjSoumMecTemps", "CalculerTotalArgent", Err, Erl
End Sub

Private Sub BarrerChamps(ByVal bLocked As Boolean)

5       On Error GoTo AfficherErreur

10      If m_eType = TYPE_SOUMISSION Then
15        txtTempsDessinSoum.Enabled = True
20        txtTempsCoupeSoum.Enabled = True
25        txtTempsMachinageSoum.Enabled = True
30        txtTempsSoudureSoum.Enabled = True
35        txtTempsAssemblageSoum.Enabled = True
40        txtTempsPeintureSoum.Enabled = True
45        txtTempsTestSoum.Enabled = True
50        txtTempsInstallationSoum.Enabled = True
55        txtTempsFormationSoum.Enabled = True
60        txtTempsGestionSoum.Enabled = True
65        txtTempsShippingSoum.Enabled = True
          txtTempsprototypeSoum.Enabled = True
  

70        txtTempsDessinSoum.Locked = bLocked
75        txtTempsCoupeSoum.Locked = bLocked
80        txtTempsMachinageSoum.Locked = bLocked
85        txtTempsSoudureSoum.Locked = bLocked
90        txtTempsAssemblageSoum.Locked = bLocked
95        txtTempsPeintureSoum.Locked = bLocked
100       txtTempsTestSoum.Locked = bLocked
105       txtTempsInstallationSoum.Locked = bLocked
110       txtTempsFormationSoum.Locked = bLocked
115       txtTempsGestionSoum.Locked = bLocked
120       txtTempsShippingSoum.Locked = bLocked
          txtTempsprototypeSoum.Locked = bLocked



125       txtTempsDessinProj.Enabled = False
130       txtTempsCoupeProj.Enabled = False
135       txtTempsMachinageProj.Enabled = False
140       txtTempsSoudureProj.Enabled = False
145       txtTempsAssemblageProj.Enabled = False
150       txtTempsPeintureProj.Enabled = False
155       txtTempsTestProj.Enabled = False
160       txtTempsInstallationProj.Enabled = False
165       txtTempsFormationProj.Enabled = False
170       txtTempsGestionProj.Enabled = False
175       txtTempsShippingProj.Enabled = False
          txtTempsPrototypeProj.Enabled = False

180       txtTempsDessinConc.Enabled = False
185       txtTempsCoupeConc.Enabled = False
190       txtTempsMachinageConc.Enabled = False
195       txtTempsSoudureConc.Enabled = False
200       txtTempsAssemblageConc.Enabled = False
205       txtTempsPeintureConc.Enabled = False
210       txtTempsTestConc.Enabled = False
215       txtTempsInstallationConc.Enabled = False
220       txtTempsFormationConc.Enabled = False
225       txtTempsGestionConc.Enabled = False
230       txtTempsShippingConc.Enabled = False
          txtTempsPrototypeConc.Enabled = False


235       cmdLock.Visible = False
240       cmdUnlock.Visible = False
245     Else
250       If m_bLocked = False Then
255         txtTempsDessinProj.Enabled = True
260         txtTempsCoupeProj.Enabled = True
265         txtTempsMachinageProj.Enabled = True
270         txtTempsSoudureProj.Enabled = True
275         txtTempsAssemblageProj.Enabled = True
280         txtTempsPeintureProj.Enabled = True
285         txtTempsTestProj.Enabled = True
290         txtTempsInstallationProj.Enabled = True
295         txtTempsFormationProj.Enabled = True
300         txtTempsGestionProj.Enabled = True
305         txtTempsShippingProj.Enabled = True
            txtTempsPrototypeProj.Enabled = True


310         txtTempsDessinProj.Locked = bLocked
315         txtTempsCoupeProj.Locked = bLocked
320         txtTempsMachinageProj.Locked = bLocked
325         txtTempsSoudureProj.Locked = bLocked
330         txtTempsAssemblageProj.Locked = bLocked
335         txtTempsPeintureProj.Locked = bLocked
340         txtTempsTestProj.Locked = bLocked
345         txtTempsInstallationProj.Locked = bLocked
350         txtTempsFormationProj.Locked = bLocked
355         txtTempsGestionProj.Locked = bLocked
360         txtTempsShippingProj.Locked = bLocked
361         txtTempsPrototypeProj.Locked = bLocked


365         txtTempsDessinSoum.Enabled = False
370         txtTempsCoupeSoum.Enabled = False
375         txtTempsMachinageSoum.Enabled = False
380         txtTempsSoudureSoum.Enabled = False
385         txtTempsAssemblageSoum.Enabled = False
390         txtTempsPeintureSoum.Enabled = False
395         txtTempsTestSoum.Enabled = False
400         txtTempsInstallationSoum.Enabled = False
405         txtTempsFormationSoum.Enabled = False
410         txtTempsGestionSoum.Enabled = False
415         txtTempsShippingSoum.Enabled = False
416         txtTempsprototypeSoum.Enabled = False



420         txtTempsDessinConc.Enabled = False
425         txtTempsCoupeConc.Enabled = False
430         txtTempsMachinageConc.Enabled = False
435         txtTempsSoudureConc.Enabled = False
440         txtTempsAssemblageConc.Enabled = False
445         txtTempsPeintureConc.Enabled = False
450         txtTempsTestConc.Enabled = False
455         txtTempsInstallationConc.Enabled = False
460         txtTempsFormationConc.Enabled = False
465         txtTempsGestionConc.Enabled = False
470         txtTempsShippingConc.Enabled = False
471         txtTempsPrototypeConc.Enabled = False

475         If m_eMode = MODE_AJOUT_MODIF Then
480           If g_bVerrouillageTempsProjet = True Then
485             cmdLock.Visible = True
490           Else
495             cmdLock.Visible = False
500           End If

505           cmdUnlock.Visible = False
510         Else
515           cmdLock.Visible = False
520           cmdUnlock.Visible = False
525         End If
530       Else
535         txtTempsDessinConc.Enabled = True
540         txtTempsCoupeConc.Enabled = True
545         txtTempsMachinageConc.Enabled = True
550         txtTempsSoudureConc.Enabled = True
555         txtTempsAssemblageConc.Enabled = True
560         txtTempsPeintureConc.Enabled = True
565         txtTempsTestConc.Enabled = True
570         txtTempsInstallationConc.Enabled = True
575         txtTempsFormationConc.Enabled = True
580         txtTempsGestionConc.Enabled = True
585         txtTempsShippingConc.Enabled = True
586         txtTempsPrototypeConc.Enabled = True


590         txtTempsDessinConc.Locked = bLocked
595         txtTempsCoupeConc.Locked = bLocked
600         txtTempsMachinageConc.Locked = bLocked
605         txtTempsSoudureConc.Locked = bLocked
610         txtTempsAssemblageConc.Locked = bLocked
615         txtTempsPeintureConc.Locked = bLocked
620         txtTempsTestConc.Locked = bLocked
625         txtTempsInstallationConc.Locked = bLocked
630         txtTempsFormationConc.Locked = bLocked
635         txtTempsGestionConc.Locked = bLocked
640         txtTempsShippingConc.Locked = bLocked
641         txtTempsPrototypeConc.Locked = bLocked

645         txtTempsDessinSoum.Enabled = False
650         txtTempsCoupeSoum.Enabled = False
655         txtTempsMachinageSoum.Enabled = False
660         txtTempsSoudureSoum.Enabled = False
665         txtTempsAssemblageSoum.Enabled = False
670         txtTempsPeintureSoum.Enabled = False
675         txtTempsTestSoum.Enabled = False
680         txtTempsInstallationSoum.Enabled = False
685         txtTempsFormationSoum.Enabled = False
690         txtTempsGestionSoum.Enabled = False
695         txtTempsShippingSoum.Enabled = False
696         txtTempsprototypeSoum.Enabled = False


700         txtTempsDessinProj.Enabled = False
705         txtTempsCoupeProj.Enabled = False
710         txtTempsMachinageProj.Enabled = False
715         txtTempsSoudureProj.Enabled = False
720         txtTempsAssemblageProj.Enabled = False
725         txtTempsPeintureProj.Enabled = False
730         txtTempsTestProj.Enabled = False
735         txtTempsInstallationProj.Enabled = False
740         txtTempsFormationProj.Enabled = False
745         txtTempsGestionProj.Enabled = False
750         txtTempsShippingProj.Enabled = False
751         txtTempsPrototypeProj.Enabled = False


755         If m_eMode = MODE_AJOUT_MODIF Then
760           If g_bDeverrouillageTempsProjet = True Then
765             cmdUnlock.Visible = True
770           Else
775             cmdUnlock.Visible = False
780           End If

785           cmdLock.Visible = False
790         Else
795           cmdLock.Visible = False
800           cmdUnlock.Visible = False
805         End If
810       End If
815     End If
  
820     txtNbrePersonne.Locked = bLocked
825     txtTempsHebergement.Locked = bLocked
830     txtTempsRepas.Locked = bLocked
835     txtTempsDeplacement.Locked = bLocked
840     txtTempsUniteMobile.Locked = bLocked
  
845     txtPrixEmballage.Locked = bLocked

850     Exit Sub

AfficherErreur:

855     woups "frmProjSoumMecTemps", "BarrerChamps", Err, Erl
End Sub

Private Sub cmdDetail_Click()

5       On Error GoTo AfficherErreur

10      If m_eType = TYPE_PROJET Then
15        Call frmDetailTemps.Afficher(m_sNoProjet, MECANIQUE, True)
20      Else
25        Call frmDetailTemps.Afficher(m_sNoSoumission, MECANIQUE, False)
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmProjSoumMecTemps", "cmdDetail_Click", Err, Erl
End Sub

Private Sub Cmdfermer_Click()

5       On Error GoTo AfficherErreur

10      If m_eMode = MODE_AJOUT_MODIF Then
15        Call EnregistrerTemps
20      End If

25      Call Unload(Me)

30      Exit Sub

AfficherErreur:

35      woups "frmProjSoumMecTemps", "cmdFermer_Click", Err, Erl
End Sub

Private Sub EnregistrerTemps()

5       On Error GoTo AfficherErreur

10      If m_eType = TYPE_SOUMISSION Then
15        If Trim$(txtTempsDessinSoum.Text) <> vbNullString And IsNumeric(txtTempsDessinSoum.Text) Then
20          FrmProjSoumMec.m_sTempsDessin = txtTempsDessinSoum.Text
25        Else
30          FrmProjSoumMec.m_sTempsDessin = "0"
35        End If

40        If Trim$(txtTempsCoupeSoum.Text) <> vbNullString And IsNumeric(txtTempsCoupeSoum.Text) Then
45          FrmProjSoumMec.m_sTempsCoupe = txtTempsCoupeSoum.Text
50        Else
55          FrmProjSoumMec.m_sTempsCoupe = "0"
60        End If

65        If Trim$(txtTempsMachinageSoum.Text) <> vbNullString And IsNumeric(txtTempsMachinageSoum.Text) Then
70          FrmProjSoumMec.m_sTempsMachinage = txtTempsMachinageSoum.Text
75        Else
80          FrmProjSoumMec.m_sTempsMachinage = "0"
85        End If
          
90        If Trim$(txtTempsSoudureSoum.Text) <> vbNullString And IsNumeric(txtTempsSoudureSoum.Text) Then
95          FrmProjSoumMec.m_sTempsSoudure = txtTempsSoudureSoum.Text
100       Else
105         FrmProjSoumMec.m_sTempsSoudure = "0"
110       End If
    
115       If Trim$(txtTempsAssemblageSoum.Text) <> vbNullString And IsNumeric(txtTempsAssemblageSoum.Text) Then
120         FrmProjSoumMec.m_sTempsAssemblage = txtTempsAssemblageSoum.Text
125       Else
130         FrmProjSoumMec.m_sTempsAssemblage = "0"
135       End If
        
140       If Trim$(txtTempsPeintureSoum.Text) <> vbNullString And IsNumeric(txtTempsPeintureSoum.Text) Then
145         FrmProjSoumMec.m_sTempsPeinture = txtTempsPeintureSoum.Text
150       Else
155         FrmProjSoumMec.m_sTempsPeinture = "0"
160       End If
    
165       If Trim$(txtTempsTestSoum.Text) <> vbNullString And IsNumeric(txtTempsTestSoum.Text) Then
170         FrmProjSoumMec.m_sTempsTest = txtTempsTestSoum.Text
175       Else
180         FrmProjSoumMec.m_sTempsTest = "0"
185       End If

190       If Trim$(txtTempsInstallationSoum.Text) <> vbNullString And IsNumeric(txtTempsInstallationSoum.Text) Then
195         FrmProjSoumMec.m_sTempsInstallation = txtTempsInstallationSoum.Text
200       Else
205         FrmProjSoumMec.m_sTempsInstallation = "0"
210       End If

215       If Trim$(txtTempsFormationSoum.Text) <> vbNullString And IsNumeric(txtTempsFormationSoum.Text) Then
220         FrmProjSoumMec.m_sTempsFormation = txtTempsFormationSoum.Text
225       Else
230         FrmProjSoumMec.m_sTempsFormation = "0"
235       End If

240       If Trim$(txtTempsGestionSoum.Text) <> vbNullString And IsNumeric(txtTempsGestionSoum.Text) Then
245         FrmProjSoumMec.m_sTempsGestion = txtTempsGestionSoum.Text
250       Else
255         FrmProjSoumMec.m_sTempsGestion = "0"
260       End If

265       If Trim$(txtTempsShippingSoum.Text) <> vbNullString And IsNumeric(txtTempsShippingSoum.Text) Then
270         FrmProjSoumMec.m_sTempsShipping = txtTempsShippingSoum.Text
275       Else
280         FrmProjSoumMec.m_sTempsShipping = "0"
285       End If
290     Else
295       FrmProjSoumMec.m_bTempsProjLock = m_bLocked

300       If m_bLocked = False Then
305         If Trim$(txtTempsDessinProj.Text) <> vbNullString And IsNumeric(txtTempsDessinProj.Text) Then
310           FrmProjSoumMec.m_sTempsDessinProj = txtTempsDessinProj.Text
315         Else
320           FrmProjSoumMec.m_sTempsDessinProj = "0"
325         End If

330         If Trim$(txtTempsCoupeProj.Text) <> vbNullString And IsNumeric(txtTempsMachinageProj.Text) Then
335           FrmProjSoumMec.m_sTempsCoupeProj = txtTempsCoupeProj.Text
340         Else
345           FrmProjSoumMec.m_sTempsCoupeProj = "0"
350         End If

355         If Trim$(txtTempsMachinageProj.Text) <> vbNullString And IsNumeric(txtTempsMachinageProj.Text) Then
360           FrmProjSoumMec.m_sTempsMachinageProj = txtTempsMachinageProj.Text
365         Else
370           FrmProjSoumMec.m_sTempsMachinageProj = "0"
375         End If

380         If Trim$(txtTempsSoudureProj.Text) <> vbNullString And IsNumeric(txtTempsSoudureProj.Text) Then
385           FrmProjSoumMec.m_sTempsSoudureProj = txtTempsSoudureProj.Text
390         Else
395           FrmProjSoumMec.m_sTempsSoudureProj = "0"
400         End If

405         If Trim$(txtTempsAssemblageProj.Text) <> vbNullString And IsNumeric(txtTempsAssemblageProj.Text) Then
410           FrmProjSoumMec.m_sTempsAssemblageProj = txtTempsAssemblageProj.Text
415         Else
420           FrmProjSoumMec.m_sTempsAssemblageProj = "0"
425         End If

430         If Trim$(txtTempsPeintureProj.Text) <> vbNullString And IsNumeric(txtTempsPeintureProj.Text) Then
435           FrmProjSoumMec.m_sTempsPeintureProj = txtTempsPeintureProj.Text
440         Else
445           FrmProjSoumMec.m_sTempsPeintureProj = "0"
450         End If

455         If Trim$(txtTempsTestProj.Text) <> vbNullString And IsNumeric(txtTempsTestProj.Text) Then
460           FrmProjSoumMec.m_sTempsTestProj = txtTempsTestProj.Text
465         Else
470           FrmProjSoumMec.m_sTempsTestProj = "0"
475         End If

480         If Trim$(txtTempsInstallationProj.Text) <> vbNullString And IsNumeric(txtTempsInstallationProj.Text) Then
485           FrmProjSoumMec.m_sTempsInstallationProj = txtTempsInstallationProj.Text
490         Else
495           FrmProjSoumMec.m_sTempsInstallationProj = "0"
500         End If

505         If Trim$(txtTempsFormationProj.Text) <> vbNullString And IsNumeric(txtTempsFormationProj.Text) Then
510           FrmProjSoumMec.m_sTempsFormationProj = txtTempsFormationProj.Text
515         Else
520           FrmProjSoumMec.m_sTempsFormationProj = "0"
525         End If

530         If Trim$(txtTempsGestionProj.Text) <> vbNullString And IsNumeric(txtTempsGestionProj.Text) Then
535           FrmProjSoumMec.m_sTempsGestionProj = txtTempsGestionProj.Text
540         Else
545           FrmProjSoumMec.m_sTempsGestionProj = "0"
550         End If

555         If Trim$(txtTempsShippingProj.Text) <> vbNullString And IsNumeric(txtTempsShippingProj.Text) Then
560           FrmProjSoumMec.m_sTempsShippingProj = txtTempsShippingProj.Text
565         Else
570           FrmProjSoumMec.m_sTempsShippingProj = "0"
575         End If
580       Else
585         If Trim$(txtTempsDessinConc.Text) <> vbNullString And IsNumeric(txtTempsDessinConc.Text) Then
590           FrmProjSoumMec.m_sTempsDessinConc = txtTempsDessinConc.Text
595         Else
600           FrmProjSoumMec.m_sTempsDessinConc = "0"
605         End If

610         If Trim$(txtTempsCoupeConc.Text) <> vbNullString And IsNumeric(txtTempsCoupeConc.Text) Then
615           FrmProjSoumMec.m_sTempsCoupeConc = txtTempsCoupeConc.Text
620         Else
625           FrmProjSoumMec.m_sTempsCoupeConc = "0"
630         End If

635         If Trim$(txtTempsMachinageConc.Text) <> vbNullString And IsNumeric(txtTempsMachinageConc.Text) Then
640           FrmProjSoumMec.m_sTempsMachinageConc = txtTempsMachinageConc.Text
645         Else
650           FrmProjSoumMec.m_sTempsMachinageConc = "0"
655         End If

660         If Trim$(txtTempsSoudureConc.Text) <> vbNullString And IsNumeric(txtTempsSoudureConc.Text) Then
665           FrmProjSoumMec.m_sTempsSoudureConc = txtTempsSoudureConc.Text
670         Else
675           FrmProjSoumMec.m_sTempsSoudureConc = "0"
680         End If

685         If Trim$(txtTempsAssemblageConc.Text) <> vbNullString And IsNumeric(txtTempsAssemblageConc.Text) Then
690           FrmProjSoumMec.m_sTempsAssemblageConc = txtTempsAssemblageConc.Text
695         Else
700           FrmProjSoumMec.m_sTempsAssemblageConc = "0"
705         End If

710         If Trim$(txtTempsPeintureConc.Text) <> vbNullString And IsNumeric(txtTempsPeintureConc.Text) Then
715           FrmProjSoumMec.m_sTempsPeintureConc = txtTempsPeintureConc.Text
720         Else
725           FrmProjSoumMec.m_sTempsPeintureConc = "0"
730         End If

735         If Trim$(txtTempsTestConc.Text) <> vbNullString And IsNumeric(txtTempsTestConc.Text) Then
740           FrmProjSoumMec.m_sTempsTestConc = txtTempsTestConc.Text
745         Else
750           FrmProjSoumMec.m_sTempsTestConc = "0"
755         End If

760         If Trim$(txtTempsInstallationConc.Text) <> vbNullString And IsNumeric(txtTempsInstallationConc.Text) Then
765           FrmProjSoumMec.m_sTempsInstallationConc = txtTempsInstallationConc.Text
770         Else
775           FrmProjSoumMec.m_sTempsInstallationConc = "0"
780         End If

785         If Trim$(txtTempsFormationConc.Text) <> vbNullString And IsNumeric(txtTempsFormationConc.Text) Then
790           FrmProjSoumMec.m_sTempsFormationConc = txtTempsFormationConc.Text
795         Else
800           FrmProjSoumMec.m_sTempsFormationConc = "0"
805         End If

810         If Trim$(txtTempsGestionConc.Text) <> vbNullString And IsNumeric(txtTempsGestionConc.Text) Then
815           FrmProjSoumMec.m_sTempsGestionConc = txtTempsGestionConc.Text
820         Else
825           FrmProjSoumMec.m_sTempsGestionConc = "0"
830         End If

835         If Trim$(txtTempsShippingConc.Text) <> vbNullString And IsNumeric(txtTempsShippingConc.Text) Then
840           FrmProjSoumMec.m_sTempsShippingConc = txtTempsShippingConc.Text
845         Else
850           FrmProjSoumMec.m_sTempsShippingConc = "0"
855         End If
860       End If
865     End If
    
870     If Trim$(txtNbrePersonne.Text) <> vbNullString And IsNumeric(txtNbrePersonne.Text) Then
875       FrmProjSoumMec.m_sNbrePersonne = txtNbrePersonne.Text
880     Else
885       FrmProjSoumMec.m_sNbrePersonne = "0"
890     End If
  
895     If Trim$(txtTempsHebergement.Text) <> vbNullString And IsNumeric(txtTempsHebergement.Text) Then
900       FrmProjSoumMec.m_sTempsHebergement = txtTempsHebergement.Text
905     Else
910       FrmProjSoumMec.m_sTempsHebergement = "0"
915     End If
       
920     If Trim$(txtTempsRepas.Text) <> vbNullString And IsNumeric(txtTempsRepas.Text) Then
925       FrmProjSoumMec.m_sTempsRepas = txtTempsRepas.Text
930     Else
935       FrmProjSoumMec.m_sTempsRepas = "0"
940     End If
    
945     If Trim$(txtTempsDeplacement.Text) <> vbNullString And IsNumeric(txtTempsDeplacement.Text) Then
950       FrmProjSoumMec.m_sTempsTransport = txtTempsDeplacement.Text
955     Else
960       FrmProjSoumMec.m_sTempsTransport = "0"
965     End If
    
970     If Trim$(txtTempsUniteMobile.Text) <> vbNullString And IsNumeric(txtTempsUniteMobile.Text) Then
975       FrmProjSoumMec.m_sTempsUniteMobile = txtTempsUniteMobile.Text
980     Else
985       FrmProjSoumMec.m_sTempsUniteMobile = "0"
990     End If
    
995     If Trim$(txtPrixEmballage.Text) <> vbNullString And IsNumeric(txtPrixEmballage.Text) Then
1000      FrmProjSoumMec.m_sPrixEmballage = txtPrixEmballage.Text
1005    Else
1010      FrmProjSoumMec.m_sPrixEmballage = "0"
1015    End If
    
1020    FrmProjSoumMec.m_sTauxHebergement1 = m_sHebergement1
1025    FrmProjSoumMec.m_sTauxHebergement2 = m_sHebergement2
1030    FrmProjSoumMec.m_sTauxRepas = m_sRepas
1035    FrmProjSoumMec.m_sTauxTransport = m_sStandard
1040    FrmProjSoumMec.m_sTauxUniteMobile = m_sUniteMobile

1045    Exit Sub

AfficherErreur:

1050    woups "frmProjSoumMecTemps", "EnregistrerTemps", Err, Erl
End Sub

Private Sub cmdLock_Click()
        
5       On Error GoTo AfficherErreur

10      If m_sTempsDessinAvant <> txtTempsDessinProj.Text Or _
           m_sTempsCoupeAvant <> txtTempsCoupeProj.Text Or _
           m_sTempsMachinageAvant <> txtTempsMachinageProj.Text Or _
           m_sTempsSoudureAvant <> txtTempsSoudureProj.Text Or _
           m_sTempsAssemblageAvant <> txtTempsAssemblageProj.Text Or _
           m_sTempsPeintureAvant <> txtTempsPeintureProj.Text Or _
           m_sTempsTestAvant <> txtTempsTestProj.Text Or _
           m_sTempsInstallationAvant <> txtTempsInstallationProj.Text Or _
           m_sTempsFormationAvant <> txtTempsFormationProj.Text Or _
           m_sTempsGestionAvant <> txtTempsGestionProj.Text Or _
           m_sTempsShippingAvant <> txtTempsShippingProj.Text Then
15        Call MsgBox("Veuillez enregistrer le projet en premier sinon vous allez perdre les informations qui ont été modifiées dans le temps projets!", vbOKOnly, "Erreur")
20      Else
25        m_bLocked = True
          
30        Call BarrerChamps(False)
35      End If

40      Exit Sub

AfficherErreur:

45      woups "frmProjSoumMecTemps", "cmdLock_Click", Err, Erl
End Sub

Private Sub cmdUnlock_Click()
        
5       On Error GoTo AfficherErreur

10      m_bLocked = False
        
15      Call BarrerChamps(False)

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumMecTemps", "cmdUnlock_Click", Err, Erl
End Sub

Private Sub InitialiserVariablesConfig()

5       On Error GoTo AfficherErreur

        'Initialise les variables à partir de la table Config (Pour avoir le taux
        'horaire le plus récent)
10      Dim rstConfig As ADODB.Recordset
  
15      Set rstConfig = New ADODB.Recordset
  
20      Call rstConfig.Open("SELECT TauxDessinMec, TauxCoupe, TauxMachinage, TauxSoudure, TauxAssemblageMec, TauxPeinture, TauxTestMec, TauxInstallationMec, TauxFormationMec, TauxGestionProjetsMec, TauxShippingMec, Repas, Hebergement1, Hebergement2, Standard, UniteMobile FROM GRB_Config", g_connData, adOpenDynamic, adLockOptimistic)
    
25      If Not IsNull(rstConfig.Fields("TauxDessinMec")) Then
30        m_sTauxDessin = rstConfig.Fields("TauxDessinMec")
35      Else
40        m_sTauxDessin = "0"
45      End If

50      If Not IsNull(rstConfig.Fields("TauxCoupe")) Then
55        m_sTauxCoupe = rstConfig.Fields("TauxCoupe")
60      Else
65        m_sTauxCoupe = "0"
70      End If

75      If Not IsNull(rstConfig.Fields("TauxMachinage")) Then
80        m_sTauxMachinage = rstConfig.Fields("TauxMachinage")
85      Else
90        m_sTauxMachinage = "0"
95      End If

100     If Not IsNull(rstConfig.Fields("TauxSoudure")) Then
105       m_sTauxSoudure = rstConfig.Fields("TauxSoudure")
110     Else
115       m_sTauxSoudure = "0"
120     End If

125     If Not IsNull(rstConfig.Fields("TauxAssemblageMec")) Then
130       m_sTauxAssemblage = rstConfig.Fields("TauxAssemblageMec")
135     Else
140       m_sTauxAssemblage = "0"
145     End If

150     If Not IsNull(rstConfig.Fields("TauxPeinture")) Then
155       m_sTauxPeinture = rstConfig.Fields("TauxPeinture")
160     Else
165       m_sTauxPeinture = "0"
170     End If

175     If Not IsNull(rstConfig.Fields("TauxTestMec")) Then
180       m_sTauxTest = rstConfig.Fields("TauxTestMec")
185     Else
190       m_sTauxTest = "0"
195     End If

200     If Not IsNull(rstConfig.Fields("TauxInstallationMec")) Then
205       m_sTauxInstallation = rstConfig.Fields("TauxInstallationMec")
210     Else
215       m_sTauxInstallation = "0"
220     End If

225     If Not IsNull(rstConfig.Fields("TauxFormationMec")) Then
230       m_sTauxFormation = rstConfig.Fields("TauxFormationMec")
235     Else
240       m_sTauxFormation = "0"
245     End If

250     If Not IsNull(rstConfig.Fields("TauxGestionProjetsMec")) Then
255       m_sTauxGestion = rstConfig.Fields("TauxGestionProjetsMec")
260     Else
265       m_sTauxGestion = "0"
270     End If

275     If Not IsNull(rstConfig.Fields("TauxShippingMec")) Then
280       m_sTauxShipping = rstConfig.Fields("TauxShippingMec")
285     Else
290       m_sTauxShipping = "0"
295     End If

        If Not IsNull(rstConfig.Fields("TauxShippingMec")) Then
296       m_sTauxPrototype = rstConfig.Fields("TauxShippingMec")
297     Else
298       m_sTauxPrototype = "0"
299     End If




300     m_sRepas = rstConfig.Fields("Repas")
305     m_sHebergement1 = rstConfig.Fields("Hebergement1")
310     m_sHebergement2 = rstConfig.Fields("Hebergement2")
315     m_sStandard = rstConfig.Fields("Standard")
320     m_sUniteMobile = rstConfig.Fields("UniteMobile")
    
325     Call rstConfig.Close
330     Set rstConfig = Nothing

335     Exit Sub

AfficherErreur:

340     woups "frmProjSoumMecTemps", "InitialiserVariablesConfig", Err, Erl
End Sub

Private Sub InitialiserVariablesProjSoum()

5       On Error GoTo AfficherErreur

10      m_sTauxDessin = FrmProjSoumMec.m_sTauxDessin
15      m_sTauxCoupe = FrmProjSoumMec.m_sTauxCoupe
20      m_sTauxMachinage = FrmProjSoumMec.m_sTauxMachinage
25      m_sTauxSoudure = FrmProjSoumMec.m_sTauxSoudure
30      m_sTauxAssemblage = FrmProjSoumMec.m_sTauxAssemblage
35      m_sTauxPeinture = FrmProjSoumMec.m_sTauxPeinture
40      m_sTauxTest = FrmProjSoumMec.m_sTauxTest
45      m_sTauxInstallation = FrmProjSoumMec.m_sTauxInstallation
50      m_sTauxFormation = FrmProjSoumMec.m_sTauxFormation
55      m_sTauxGestion = FrmProjSoumMec.m_sTauxGestion
60      m_sTauxShipping = FrmProjSoumMec.m_sTauxShipping
        m_sTauxPrototype = FrmProjSoumMec.m_sTauxShipping

65      m_sRepas = FrmProjSoumMec.m_sTauxRepas
70      m_sHebergement1 = FrmProjSoumMec.m_sTauxHebergement1
75      m_sHebergement2 = FrmProjSoumMec.m_sTauxHebergement2
80      m_sStandard = FrmProjSoumMec.m_sTauxTransport
85      m_sUniteMobile = FrmProjSoumMec.m_sTauxUniteMobile

90      Exit Sub

AfficherErreur:

95      woups "frmProjSoumMecTemps", "InitialiserVariablesProjSoum", Err, Erl
End Sub

Private Sub Form_Load()
  
5       On Error GoTo AfficherErreur
  
10      If FrmProjSoumMec.m_bDroitPrix = False Then
15        fraRessourcesHumaines.width = 4150
20        fraFraisSubsistences.width = 4150
    
25        fraFraisSubsistences.Left = 4390
    
30        fraManutention.Visible = False
35        lblTotalPrixRH.Visible = False
    
40        Cmdfermer.Left = 7320
    
45        Cmdfermer.Top = 4200
    
50        Me.width = 8800
55        Me.Height = 7485
60      End If

65      Exit Sub

AfficherErreur:

70      woups "frmProjSoumMecTemps", "Form_Load", Err, Erl
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

50      woups "frmProjSoumMecTemps", "txtNbrePersonne_Change", Err, Erl
End Sub

Private Sub txtPrixEmballage_KeyPress(KeyAscii As Integer)

5       On Error GoTo AfficherErreur

10      If KeyAscii = 46 Then 'Si c'est le "."
15        KeyAscii = 44 'Remplace par la ","
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmProjSoumMecTemps", "lblPrixEmballage_KeyPress", Err, Erl
End Sub

Private Sub txtTempsAssemblageConc_Change()
  
5       On Error GoTo AfficherErreur

10      Call CalculerTotalTemps

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsAssemblageConc_Change", Err, Erl
End Sub

Private Sub txtTempsAssemblageProj_Change()
  
5       On Error GoTo AfficherErreur

10      Call CalculerTotalTemps

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsAssemblageProj_Change", Err, Erl
End Sub

Private Sub txtTempsCoupeConc_Change()
  
5       On Error GoTo AfficherErreur

10      Call CalculerTotalTemps

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsCoupeConc_Change", Err, Erl
End Sub

Private Sub txtTempsCoupeProj_Change()
  
5       On Error GoTo AfficherErreur

10      Call CalculerTotalTemps

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsCoupeProj_Change", Err, Erl
End Sub

Private Sub txtTempsDessinConc_Change()
  
5       On Error GoTo AfficherErreur

10      Call CalculerTotalTemps

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsDessinConc_Change", Err, Erl
End Sub

Private Sub txtTempsDessinProj_Change()

5       On Error GoTo AfficherErreur

10      Call CalculerTotalTemps

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsDessinProj_Change", Err, Erl
End Sub

Private Sub txtTempsFormationConc_Change()
  
5       On Error GoTo AfficherErreur

10      Call CalculerTotalTemps

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsFormationConc_Change", Err, Erl
End Sub

Private Sub txtTempsFormationProj_Change()
  
5       On Error GoTo AfficherErreur

10      Call CalculerTotalTemps

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsFormationProj_Change", Err, Erl
End Sub

Private Sub txtTempsGestionConc_Change()
  
5       On Error GoTo AfficherErreur

10      Call CalculerTotalTemps

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsGestionConc_Change", Err, Erl
End Sub

Private Sub txtTempsGestionProj_Change()
  
5       On Error GoTo AfficherErreur

10      Call CalculerTotalTemps

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsGestionProj_Change", Err, Erl
End Sub

Private Sub txtTempsShippingConc_Change()
  
5       On Error GoTo AfficherErreur

10      Call CalculerTotalTemps

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsShippingConc_Change", Err, Erl
End Sub

Private Sub txtTempsShippingProj_Change()
  
5       On Error GoTo AfficherErreur

10      Call CalculerTotalTemps

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsShippingProj_Change", Err, Erl
End Sub

Private Sub txtTempsInstallationConc_Change()
  
5       On Error GoTo AfficherErreur

10      Call CalculerTotalTemps

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsInstallationConc_Change", Err, Erl
End Sub

Private Sub txtTempsInstallationProj_Change()
  
5       On Error GoTo AfficherErreur

10      Call CalculerTotalTemps

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsInstallationProj_Change", Err, Erl
End Sub

Private Sub txtTempsMachinageConc_Change()
  
5       On Error GoTo AfficherErreur

10      Call CalculerTotalTemps

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsMachinageConc_Change", Err, Erl
End Sub

Private Sub txtTempsMachinageProj_Change()
  
5       On Error GoTo AfficherErreur

10      Call CalculerTotalTemps

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsMachinageProj_Change", Err, Erl
End Sub

Private Sub txtTempsMachinageSoum_Change()

5       On Error GoTo AfficherErreur

10      If m_eType = TYPE_SOUMISSION Then
15        If IsNumeric(txtTempsMachinageSoum.Text) Then
20          lblPrixMachinage.Caption = Round(Replace(txtTempsMachinageSoum.Text * m_sTauxMachinage, ".", ","), 2)
25        Else
30          lblPrixMachinage.Caption = 0
35        End If
  
40        Call CalculerTotal
45      End If

50      Exit Sub

AfficherErreur:

55      woups "frmProjSoumMecTemps", "txtTempsMachinageSoum_Change", Err, Erl
End Sub

Private Sub txtTempsMachinageSoum_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsMachinageSoum.Text = Replace(txtTempsMachinageSoum.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsMachinageSoum_LostFocus", Err, Erl
End Sub

Private Sub txtTempsMachinageProj_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsMachinageProj.Text = Replace(txtTempsMachinageProj.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsMachinageProj_LostFocus", Err, Erl
End Sub

Private Sub txtTempsMachinageConc_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsMachinageConc.Text = Replace(txtTempsMachinageConc.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsMachinageConc_LostFocus", Err, Erl
End Sub

Private Sub txtTempsCoupeSoum_Change()

5       On Error GoTo AfficherErreur

10      If m_eType = TYPE_SOUMISSION Then
15        If IsNumeric(txtTempsCoupeSoum.Text) Then
20          lblPrixCoupe.Caption = Round(Replace(txtTempsCoupeSoum.Text * m_sTauxCoupe, ".", ","), 2)
25        Else
30          lblPrixCoupe.Caption = 0
35        End If

40        Call CalculerTotal
45      End If

50      Exit Sub

AfficherErreur:

55      woups "frmProjSoumMecTemps", "txtTempsCoupeSoum_Change", Err, Erl
End Sub

Private Sub txtTempsCoupeSoum_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsCoupeSoum.Text = Replace(txtTempsCoupeSoum.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsCoupeSoum_LostFocus", Err, Erl
End Sub

Private Sub txtTempsCoupeProj_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsCoupeProj.Text = Replace(txtTempsCoupeProj.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsCoupeProj_LostFocus", Err, Erl
End Sub

Private Sub txtTempsCoupeConc_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsCoupeConc.Text = Replace(txtTempsCoupeConc.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsCoupeConc_LostFocus", Err, Erl
End Sub

Private Sub txtTempsPeintureConc_Change()
        
5       On Error GoTo AfficherErreur

10      Call CalculerTotalTemps

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsPeintureConc_Change", Err, Erl
End Sub

Private Sub txtTempsPeintureProj_Change()
  
5       On Error GoTo AfficherErreur

10      Call CalculerTotalTemps

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsPeintureProj_Change", Err, Erl
End Sub

Private Sub txtTempsSoudureConc_Change()
  
5       On Error GoTo AfficherErreur

10      Call CalculerTotalTemps

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsSoudureConc_Change", Err, Erl
End Sub

Private Sub txtTempsSoudureProj_Change()
  
5       On Error GoTo AfficherErreur

10      Call CalculerTotalTemps

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsSoudureProj_Change", Err, Erl
End Sub

Private Sub txtTempsSoudureSoum_Change()

5       On Error GoTo AfficherErreur

10      If m_eType = TYPE_SOUMISSION Then
15        If IsNumeric(txtTempsSoudureSoum.Text) Then
20          lblPrixSoudure.Caption = Round(Replace(txtTempsSoudureSoum.Text * m_sTauxSoudure, ".", ","), 2)
25        Else
30          lblPrixSoudure.Caption = 0
35        End If
  
40        Call CalculerTotal
45      End If

50      Exit Sub

AfficherErreur:

55      woups "frmProjSoumMecTemps", "txtTempsSoudureSoum_Change", Err, Erl
End Sub

Private Sub txtTempsSoudureSoum_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsSoudureSoum.Text = Replace(txtTempsSoudureSoum.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsSoudureSoum_LostFocus", Err, Erl
End Sub

Private Sub txtTempsSoudureProj_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsSoudureProj.Text = Replace(txtTempsSoudureProj.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsSoudureProj_LostFocus", Err, Erl
End Sub

Private Sub txtTempsSoudureConc_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsSoudureConc.Text = Replace(txtTempsSoudureConc.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsSoudureConc_LostFocus", Err, Erl
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

55      woups "frmProjSoumMecTemps", "txtTempsAssemblageSoum_Change", Err, Erl
End Sub

Private Sub txtTempsAssemblageSoum_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsAssemblageSoum.Text = Replace(txtTempsAssemblageSoum.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsAssemblageSoum_LostFocus", Err, Erl
End Sub

Private Sub txtTempsAssemblageProj_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsAssemblageProj.Text = Replace(txtTempsAssemblageProj.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsAssemblageProj_LostFocus", Err, Erl
End Sub

Private Sub txtTempsAssemblageConc_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsAssemblageConc.Text = Replace(txtTempsAssemblageConc.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsAssemblageConc_LostFocus", Err, Erl
End Sub

Private Sub txtTempsPeintureSoum_Change()

5       On Error GoTo AfficherErreur

10      If m_eType = TYPE_SOUMISSION Then
15        If IsNumeric(txtTempsPeintureSoum.Text) Then
20          lblPrixPeinture.Caption = Round(Replace(txtTempsPeintureSoum.Text * m_sTauxPeinture, ".", ","), 2)
25        Else
30          lblPrixPeinture.Caption = 0
35        End If
  
40        Call CalculerTotal
45      End If

50      Exit Sub

AfficherErreur:

55      woups "frmProjSoumMecTemps", "txtTempsPeintureSoum_Change", Err, Erl
End Sub

Private Sub txtTempsPeintureSoum_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsPeintureSoum.Text = Replace(txtTempsPeintureSoum.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsPeintureSoum_LostFocus", Err, Erl
End Sub

Private Sub txtTempsPeintureProj_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsPeintureProj.Text = Replace(txtTempsPeintureProj.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsPeintureProj_LostFocus", Err, Erl
End Sub

Private Sub txtTempsPeintureConc_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsPeintureConc.Text = Replace(txtTempsPeintureConc.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsPeintureConc_LostFocus", Err, Erl
End Sub

Private Sub txtTempsTestConc_Change()
  
5       On Error GoTo AfficherErreur

10      Call CalculerTotalTemps

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsTestConc_Change", Err, Erl
End Sub

Private Sub txtTempsTestProj_Change()
  
5       On Error GoTo AfficherErreur

10      Call CalculerTotalTemps

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsTestProj_Change", Err, Erl
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

55      woups "frmProjSoumMecTemps", "txtTempsTestSoum_Change", Err, Erl
End Sub

Private Sub txtTempsTestSoum_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsTestSoum.Text = Replace(txtTempsTestSoum.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsTestSoum_LostFocus", Err, Erl
End Sub

Private Sub txtTempsTestProj_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsTestProj.Text = Replace(txtTempsTestProj.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsTestProj_LostFocus", Err, Erl
End Sub

Private Sub txtTempsTestConc_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsTestConc.Text = Replace(txtTempsTestConc.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsTestConc_LostFocus", Err, Erl
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

55      woups "frmProjSoumMecTemps", "txtTempsInstallationSoum_Change", Err, Erl
End Sub

Private Sub txtTempsInstallationSoum_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsInstallationSoum.Text = Replace(txtTempsInstallationSoum.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsInstallationSoum_LostFocus", Err, Erl
End Sub

Private Sub txtTempsInstallationProj_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsInstallationProj.Text = Replace(txtTempsInstallationProj.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsInstallationProj_LostFocus", Err, Erl
End Sub

Private Sub txtTempsInstallationConc_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsInstallationConc.Text = Replace(txtTempsInstallationConc.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsInstallationConc_LostFocus", Err, Erl
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

55      woups "frmProjSoumMecTemps", "txtTempsDessinSoum_Change", Err, Erl
End Sub

Private Sub txtTempsDessinSoum_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsDessinSoum.Text = Replace(txtTempsDessinSoum.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsDessinSoum_LostFocus", Err, Erl
End Sub

Private Sub txtTempsDessinProj_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsDessinProj.Text = Replace(txtTempsDessinProj.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsDessinProj_LostFocus", Err, Erl
End Sub

Private Sub txtTempsDessinConc_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsDessinConc.Text = Replace(txtTempsDessinConc.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsDessinConc_LostFocus", Err, Erl
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

55      woups "frmProjSoumMecTemps", "txtTempsFormationSoum_Change", Err, Erl
End Sub

Private Sub txtTempsFormationSoum_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsFormationSoum.Text = Replace(txtTempsFormationSoum.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsFormationSoum_LostFocus", Err, Erl
End Sub

Private Sub txtTempsFormationProj_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsFormationProj.Text = Replace(txtTempsFormationProj.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsFormationProj_LostFocus", Err, Erl
End Sub

Private Sub txtTempsFormationConc_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsFormationConc.Text = Replace(txtTempsFormationConc.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsFormationConc_LostFocus", Err, Erl
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

55      woups "frmProjSoumMecTemps", "txtTempsGestionSoum_Change", Err, Erl
End Sub

Private Sub txtTempsGestionSoum_LostFocus()

5       On Error GoTo AfficherErreur
 
10      txtTempsGestionSoum.Text = Replace(txtTempsGestionSoum.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsGestionSoum_LostFocus", Err, Erl
End Sub

Private Sub txtTempsGestionProj_LostFocus()

5       On Error GoTo AfficherErreur
 
10      txtTempsGestionProj.Text = Replace(txtTempsGestionProj.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsGestionProj_LostFocus", Err, Erl
End Sub

Private Sub txtTempsGestionConc_LostFocus()

5       On Error GoTo AfficherErreur
 
10      txtTempsGestionConc.Text = Replace(txtTempsGestionConc.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsGestionConc_LostFocus", Err, Erl
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

55      woups "frmProjSoumMecTemps", "txtTempsShippingSoum_Change", Err, Erl
End Sub

Private Sub txtTempsShippingSoum_LostFocus()

5       On Error GoTo AfficherErreur
 
10      txtTempsShippingSoum.Text = Replace(txtTempsShippingSoum.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsShippingSoum_LostFocus", Err, Erl
End Sub

Private Sub txtTempsShippingProj_LostFocus()

5       On Error GoTo AfficherErreur
 
10      txtTempsShippingProj.Text = Replace(txtTempsShippingProj.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsShippingProj_LostFocus", Err, Erl
End Sub

Private Sub txtTempsShippingConc_LostFocus()

5       On Error GoTo AfficherErreur
 
10      txtTempsShippingConc.Text = Replace(txtTempsShippingConc.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsShippingConc_LostFocus", Err, Erl
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

45      woups "frmProjSoumMecTemps", "txtTempsHebergement_Change", Err, Erl
End Sub

Private Sub txtTempsHebergement_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsHebergement.Text = Replace(txtTempsHebergement.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsHebergement_LostFocus", Err, Erl
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

45      woups "frmProjSoumMecTemps", "txtTempsRepas_Change", Err, Erl
End Sub

Private Sub txtTempsRepas_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsRepas.Text = Replace(txtTempsRepas.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsRepas_LostFocus", Err, Erl
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

45      woups "frmProjSoumMecTemps", "txtTempsDeplacement_Change", Err, Erl
End Sub

Private Sub txtTempsDeplacement_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsDeplacement.Text = Replace(txtTempsDeplacement.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsDeplacement_LostFocus", Err, Erl
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

45      woups "frmProjSoumMecTemps", "txtTempsUniteMobile_Change", Err, Erl
End Sub

Private Sub txtTempsUniteMobile_LostFocus()

5       On Error GoTo AfficherErreur

10      txtTempsUniteMobile.Text = Replace(txtTempsUniteMobile.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "txtTempsUniteMobile_LostFocus", Err, Erl
End Sub

Private Sub txtPrixEmballage_Change()

5       On Error GoTo AfficherErreur
        
10      Call CalculerTotal

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMecTemps", "lblPrixEmballage_Change", Err, Erl
End Sub

Private Sub txtPrixEmballage_LostFocus()

5       On Error GoTo AfficherErreur

10      If IsNumeric(txtPrixEmballage.Text) Then
15        txtPrixEmballage.Text = Round(Replace(txtPrixEmballage.Text, ".", ","), 2)
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmProjSoumMecTemps", "lblPrixEmballage_LostFocus", Err, Erl
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

110     woups "frmProjSoumMecTemps", "CalculerHebergement", Err, Erl
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

90     woups "frmProjSoumMecTemps", "CalculerRepas", Err, Erl
End Sub

Private Sub CalculerTotal()

5       On Error GoTo AfficherErreur

10      Dim dblTotal            As Double
15      Dim dblPrixEmballage    As Double
20      Dim dblTotalArgentRH    As Double
25      Dim dblPrixDessin       As Double
30      Dim dblPrixCoupe        As Double
35      Dim dblPrixMachinage    As Double
40      Dim dblPrixSoudure      As Double
45      Dim dblPrixAssemblage   As Double
50      Dim dblPrixPeinture     As Double
55      Dim dblPrixTest         As Double
60      Dim dblPrixInstallation As Double
65      Dim dblPrixFormation    As Double
70      Dim dblPrixGestion      As Double
75      Dim dblPrixShipping     As Double
80      Dim dblPrixHebergement  As Double
85      Dim dblPrixRepas        As Double
90      Dim dblPrixDeplacement  As Double
95      Dim dblPrixUniteMobile  As Double
  
        'Prix de dessin
100     If IsNumeric(lblPrixDessin.Caption) Then
105       dblPrixDessin = CDbl(lblPrixDessin.Caption)
110     Else
115       dblPrixDessin = 0
120     End If
  
        'Prix de coupe et préparation
125     If IsNumeric(lblPrixCoupe.Caption) Then
130       dblPrixCoupe = CDbl(lblPrixCoupe.Caption)
135     Else
140       dblPrixCoupe = 0
145     End If
  
        'Prix de machinage
150     If IsNumeric(lblPrixMachinage.Caption) Then
155       dblPrixMachinage = CDbl(lblPrixMachinage.Caption)
160     Else
165       dblPrixMachinage = 0
170     End If
        
        'Prix de soudure et meulage
175     If IsNumeric(lblPrixSoudure.Caption) Then
180       dblPrixSoudure = CDbl(lblPrixSoudure.Caption)
185     Else
190       dblPrixSoudure = 0
195     End If

        'Prix d'assemblage du système
200     If IsNumeric(lblPrixAssemblage.Caption) Then
205       dblPrixAssemblage = CDbl(lblPrixAssemblage.Caption)
210     Else
215       dblPrixAssemblage = 0
220     End If

        'Prix de peinture et finition
225     If IsNumeric(lblPrixPeinture.Caption) Then
230       dblPrixPeinture = CDbl(lblPrixPeinture.Caption)
235     Else
240       dblPrixPeinture = 0
245     End If

        'Prix de tests finaux
250     If IsNumeric(lblPrixTest.Caption) Then
255       dblPrixTest = CDbl(lblPrixTest.Caption)
260     Else
265       dblPrixTest = 0
270     End If

        'Prix d'Installation
275     If IsNumeric(lblPrixInstallation.Caption) Then
280       dblPrixInstallation = CDbl(lblPrixInstallation.Caption)
285     Else
290       dblPrixInstallation = 0
295     End If

        'Prix de formation
300     If IsNumeric(lblPrixFormation.Caption) Then
305       dblPrixFormation = CDbl(lblPrixFormation.Caption)
310     Else
315       dblPrixFormation = 0
320     End If

        'Prix de gestion du projet
325     If IsNumeric(lblPrixGestion.Caption) Then
330       dblPrixGestion = CDbl(lblPrixGestion.Caption)
335     Else
340       dblPrixGestion = 0
345     End If

        'Prix de shipping
350     If IsNumeric(lblPrixShipping.Caption) Then
355       dblPrixShipping = CDbl(lblPrixShipping.Caption)
360     Else
365       dblPrixShipping = 0
370     End If


        'Prix de dévelloppement prototypage
371     If IsNumeric(lblPrixPrototype.Caption) Then
372       'dblPrixPrototype = CDbl(lblPrixPrototype.Caption)
373     Else
374       'dblPrixPrototype = 0
        End If



        'Prix d'hébergement
375     If IsNumeric(lblPrixHebergement.Caption) Then
380       dblPrixHebergement = CDbl(lblPrixHebergement.Caption)
385     Else
390       dblPrixHebergement = 0
395     End If

        'Prix des repas
400     If IsNumeric(lblPrixRepas.Caption) Then
405       dblPrixRepas = CDbl(lblPrixRepas.Caption)
410     Else
415       dblPrixRepas = 0
420     End If
  
        'Prix du déplacement
425     If IsNumeric(lblPrixDeplacement.Caption) Then
430       dblPrixDeplacement = CDbl(lblPrixDeplacement.Caption)
435     Else
440       dblPrixDeplacement = 0
445     End If

        'Prix de l'unité mobile
450     If IsNumeric(lblPrixUniteMobile.Caption) Then
455       dblPrixUniteMobile = CDbl(lblPrixUniteMobile.Caption)
460     Else
465       dblPrixUniteMobile = 0
470     End If
   
        'Prix de transport et emballage
475     If IsNumeric(txtPrixEmballage.Text) Then
480       dblPrixEmballage = CDbl(txtPrixEmballage.Text)
485     Else
490       dblPrixEmballage = 0
495     End If
                          
500     dblTotalArgentRH = dblPrixDessin + _
                           dblPrixCoupe + _
                           dblPrixMachinage + _
                           dblPrixSoudure + _
                           dblPrixAssemblage + _
                           dblPrixPeinture + _
                           dblPrixTest + _
                           dblPrixInstallation + _
                           dblPrixFormation + _
                           dblPrixGestion + _
                           dblPrixShipping
                           'dblPrixPrototype

505     lblTotalPrixRH.Caption = Conversion(CStr(dblTotalArgentRH), MODE_DECIMAL)

  
510     dblTotal = dblTotalArgentRH + _
                   dblPrixHebergement + _
                   dblPrixRepas + _
                   dblPrixDeplacement + _
                   dblPrixUniteMobile + _
                   dblPrixEmballage
             
515     lblTotal.Caption = Conversion(CStr(dblTotal), MODE_DECIMAL)

520     Call CalculerTotalTemps

525     Exit Sub

AfficherErreur:

530     woups "frmProjSoumMecTemps", "CalculerTotal", Err, Erl
End Sub

Private Sub CalculerTotalTemps()

  
5       On Error GoTo AfficherErreur

10      Dim dblTempsDessin       As Double
15      Dim dblTempsCoupe        As Double
20      Dim dblTempsMachinage    As Double
25      Dim dblTempsSoudure      As Double
30      Dim dblTempsAssemblage   As Double
35      Dim dblTempsPeinture     As Double
40      Dim dblTempsTest         As Double
45      Dim dblTempsInstallation As Double
50      Dim dblTempsFormation    As Double
55      Dim dblTempsGestion      As Double
60      Dim dblTempsShipping     As Double
65      Dim dblTotalTemps        As Double

        'SOUMISSION
70      If IsNumeric(txtTempsDessinSoum.Text) Then
75        dblTempsDessin = CDbl(txtTempsDessinSoum.Text)
80      Else
85        dblTempsDessin = 0
90      End If
        
95      If IsNumeric(txtTempsCoupeSoum.Text) Then
100       dblTempsCoupe = CDbl(txtTempsCoupeSoum.Text)
105     Else
110       dblTempsCoupe = 0
115     End If

120     If IsNumeric(txtTempsMachinageSoum.Text) Then
125       dblTempsMachinage = CDbl(txtTempsMachinageSoum.Text)
130     Else
135       dblTempsMachinage = 0
140     End If

145     If IsNumeric(txtTempsSoudureSoum.Text) Then
150       dblTempsSoudure = CDbl(txtTempsSoudureSoum.Text)
155     Else
160       dblTempsSoudure = 0
165     End If

170     If IsNumeric(txtTempsAssemblageSoum.Text) Then
175       dblTempsAssemblage = CDbl(txtTempsAssemblageSoum.Text)
180     Else
185       dblTempsAssemblage = 0
190     End If

195     If IsNumeric(txtTempsPeintureSoum.Text) Then
200       dblTempsPeinture = CDbl(txtTempsPeintureSoum.Text)
205     Else
210       dblTempsPeinture = 0
215     End If

220     If IsNumeric(txtTempsTestSoum.Text) Then
225       dblTempsTest = CDbl(txtTempsTestSoum.Text)
230     Else
235       dblTempsTest = 0
240     End If

245     If IsNumeric(txtTempsInstallationSoum.Text) Then
250       dblTempsInstallation = CDbl(txtTempsInstallationSoum.Text)
255     Else
260       dblTempsInstallation = 0
265     End If

270     If IsNumeric(txtTempsFormationSoum.Text) Then
275       dblTempsFormation = CDbl(txtTempsFormationSoum.Text)
280     Else
285       dblTempsFormation = 0
290     End If

295     If IsNumeric(txtTempsGestionSoum.Text) Then
300       dblTempsGestion = CDbl(txtTempsGestionSoum.Text)
305     Else
310       dblTempsGestion = 0
315     End If

320     If IsNumeric(txtTempsShippingSoum.Text) Then
325       dblTempsShipping = CDbl(txtTempsShippingSoum.Text)
330     Else
335       dblTempsShipping = 0
340     End If

345     dblTotalTemps = dblTempsDessin + _
                        dblTempsCoupe + _
                        dblTempsMachinage + _
                        dblTempsSoudure + _
                        dblTempsAssemblage + _
                        dblTempsPeinture + _
                        dblTempsTest + _
                        dblTempsInstallation + _
                        dblTempsFormation + _
                        dblTempsGestion + _
                        dblTempsShipping

350     lblTotalTempsRHSoum.Caption = Conversion(dblTotalTemps, MODE_DECIMAL)

        'PROJET
355     If m_eType = TYPE_PROJET Then
360       If IsNumeric(txtTempsDessinProj.Text) Then
365         dblTempsDessin = CDbl(txtTempsDessinProj.Text)
370       Else
375         dblTempsDessin = 0
380       End If

385       If IsNumeric(txtTempsCoupeProj.Text) Then
390         dblTempsCoupe = CDbl(txtTempsCoupeProj.Text)
395       Else
400         dblTempsCoupe = 0
405       End If

410       If IsNumeric(txtTempsMachinageProj.Text) Then
415         dblTempsMachinage = CDbl(txtTempsMachinageProj.Text)
420       Else
425         dblTempsMachinage = 0
430       End If

435       If IsNumeric(txtTempsSoudureProj.Text) Then
440         dblTempsSoudure = CDbl(txtTempsSoudureProj.Text)
445       Else
450         dblTempsSoudure = 0
455       End If

460       If IsNumeric(txtTempsAssemblageProj.Text) Then
465         dblTempsAssemblage = CDbl(txtTempsAssemblageProj.Text)
470       Else
475         dblTempsAssemblage = 0
480       End If

485       If IsNumeric(txtTempsPeintureProj.Text) Then
490         dblTempsPeinture = CDbl(txtTempsPeintureProj.Text)
495       Else
500         dblTempsPeinture = 0
505       End If

510       If IsNumeric(txtTempsTestProj.Text) Then
515         dblTempsTest = CDbl(txtTempsTestProj.Text)
520       Else
525         dblTempsTest = 0
530       End If

535       If IsNumeric(txtTempsInstallationProj.Text) Then
540         dblTempsInstallation = CDbl(txtTempsInstallationProj.Text)
545       Else
550         dblTempsInstallation = 0
555       End If

560       If IsNumeric(txtTempsFormationProj.Text) Then
565         dblTempsFormation = CDbl(txtTempsFormationProj.Text)
570       Else
575         dblTempsFormation = 0
580       End If

585       If IsNumeric(txtTempsGestionProj.Text) Then
590         dblTempsGestion = CDbl(txtTempsGestionProj.Text)
595       Else
600         dblTempsGestion = 0
605       End If

610       If IsNumeric(txtTempsShippingProj.Text) Then
615         dblTempsShipping = CDbl(txtTempsShippingProj.Text)
620       Else
625         dblTempsShipping = 0
630       End If


631       If IsNumeric(txtTempsPrototypeProj.Text) Then
632       '  dblTempsPrototype = CDbl(txtTempsPrototypeProj.Text)
633       Else
634       '  dblTempsPrototype = 0
          End If



635       dblTotalTemps = dblTempsDessin + _
                          dblTempsCoupe + _
                          dblTempsMachinage + _
                          dblTempsSoudure + _
                          dblTempsAssemblage + _
                          dblTempsPeinture + _
                          dblTempsTest + _
                          dblTempsInstallation + _
                          dblTempsFormation + _
                          dblTempsGestion + _
                          dblTempsShipping

                          'dblTempsPrototype

640       lblTotalTempsRHProj.Caption = Conversion(dblTotalTemps, MODE_DECIMAL)
645     End If

        'CONCEPTION
650     If m_eType = TYPE_PROJET And m_bLocked = True Then
655       If IsNumeric(txtTempsDessinConc.Text) Then
660         dblTempsDessin = CDbl(txtTempsDessinConc.Text)
665       Else
670         dblTempsDessin = 0
675       End If

680       If IsNumeric(txtTempsCoupeConc.Text) Then
685         dblTempsCoupe = CDbl(txtTempsCoupeConc.Text)
690       Else
695         dblTempsCoupe = 0
700       End If

705       If IsNumeric(txtTempsMachinageConc.Text) Then
710         dblTempsMachinage = CDbl(txtTempsMachinageConc.Text)
715       Else
720         dblTempsMachinage = 0
725       End If

730       If IsNumeric(txtTempsSoudureConc.Text) Then
735         dblTempsSoudure = CDbl(txtTempsSoudureConc.Text)
740       Else
745         dblTempsSoudure = 0
750       End If

755       If IsNumeric(txtTempsAssemblageConc.Text) Then
760         dblTempsAssemblage = CDbl(txtTempsAssemblageConc.Text)
765       Else
770         dblTempsAssemblage = 0
775       End If

780       If IsNumeric(txtTempsPeintureConc.Text) Then
785         dblTempsPeinture = CDbl(txtTempsPeintureConc.Text)
790       Else
795         dblTempsPeinture = 0
800       End If

805       If IsNumeric(txtTempsTestConc.Text) Then
810         dblTempsTest = CDbl(txtTempsTestConc.Text)
815       Else
820         dblTempsTest = 0
825       End If

830       If IsNumeric(txtTempsInstallationConc.Text) Then
835         dblTempsInstallation = CDbl(txtTempsInstallationConc.Text)
840       Else
845         dblTempsInstallation = 0
850       End If

855       If IsNumeric(txtTempsFormationConc.Text) Then
860         dblTempsFormation = CDbl(txtTempsFormationConc.Text)
865       Else
870         dblTempsFormation = 0
875       End If

880       If IsNumeric(txtTempsGestionConc.Text) Then
885         dblTempsGestion = CDbl(txtTempsGestionConc.Text)
890       Else
895         dblTempsGestion = 0
900       End If

905       If IsNumeric(txtTempsShippingConc.Text) Then
910         dblTempsShipping = CDbl(txtTempsShippingConc.Text)
915       Else
920         dblTempsShipping = 0
925       End If

930       dblTotalTemps = dblTempsDessin + _
                          dblTempsCoupe + _
                          dblTempsMachinage + _
                          dblTempsSoudure + _
                          dblTempsAssemblage + _
                          dblTempsPeinture + _
                          dblTempsTest + _
                          dblTempsInstallation + _
                          dblTempsFormation + _
                          dblTempsGestion + _
                          dblTempsShipping
                        
935       lblTotalTempsRHConc.Caption = Conversion(dblTotalTemps, MODE_DECIMAL)
940     End If

945     Exit Sub

AfficherErreur:

950     woups "frmProjSoumMecTemps", "CalculerTotalTemps", Err, Erl
End Sub
