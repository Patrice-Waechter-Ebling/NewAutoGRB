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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   13155
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
      Picture         =   "frmProjSoumMecTemps.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7920
      Width           =   735
   End
   Begin VB.CommandButton cmdLock 
      Height          =   615
      Left            =   10800
      Picture         =   "frmProjSoumMecTemps.frx":0442
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

Private m_sTauxDessin As String
Private m_sTauxCoupe As String
Private m_sTauxMachinage As String
Private m_sTauxSoudure As String
Private m_sTauxAssemblage As String
Private m_sTauxPeinture As String
Private m_sTauxTest As String
Private m_sTauxInstallation As String
Private m_sTauxFormation As String
Private m_sTauxGestion As String
Private m_sTauxShipping As String
Private m_sTauxPrototype As String

Private m_sRepas As String
Private m_sHebergement As String
Private m_sStandard As String
Private m_sUniteMobile As String

Private m_sTempsDessinAvant As String
Private m_sTempsCoupeAvant As String
Private m_sTempsMachinageAvant As String
Private m_sTempsSoudureAvant As String
Private m_sTempsAssemblageAvant As String
Private m_sTempsPeintureAvant As String
Private m_sTempsTestAvant As String
Private m_sTempsInstallationAvant As String
Private m_sTempsFormationAvant As String
Private m_sTempsGestionAvant As String
Private m_sTempsShippingAvant As String
Private m_sTempsPrototypeAvant As String
Private m_sTempsTotalRHAvant As String

Private m_sNoProjet As String
Private m_sNoSoumission As String

Private m_eType As enumType

Private m_eMode As enumMode
 
Private m_bNouveauTaux As Boolean 'Pour savoir si les nouveaux taux doivent être pris
Private m_bLocked As Boolean 'Pour savoir si la section projet est barrée ou non

Public Sub Afficher(ByVal sNoProjet As String, ByVal sNoSoumission As String, ByVal iType As Integer, ByVal iMode As Integer, ByVal bNouveauTaux As Boolean)

 On Error GoTo Oups
 
 m_eType = iType
 
 m_eMode = iMode
 
 m_sNoProjet = sNoProjet
 m_sNoSoumission = sNoSoumission
 
 m_bNouveauTaux = bNouveauTaux
 
 If bNouveauTaux = True Then
 Call InitialiserVariablesConfig
 Else
 Call InitialiserVariablesProjSoum
 End If
 
  Call AfficherEnregistrement
 
  Call RemplirValeursAvant

  If m_eMode = MODE_AJOUT_MODIF Then
  Call BarrerChamps(False)
  Else
  Call BarrerChamps(True)
  End If
 
  Call Me.Show(vbModal)

10 Exit Sub

Oups:

wOups "frmProjSoumMecTemps", "Afficher", Err, Err.number, Err.Description
End Sub

Private Sub RemplirValeursAvant()
 
 On Error GoTo Oups

 m_sTempsDessinAvant = txtTempsDessinProj.Text
 m_sTempsCoupeAvant = txtTempsCoupeProj.Text
 m_sTempsMachinageAvant = txtTempsMachinageProj.Text
 m_sTempsSoudureAvant = txtTempsSoudureProj.Text
 m_sTempsAssemblageAvant = txtTempsAssemblageProj.Text
 m_sTempsPeintureAvant = txtTempsPeintureProj.Text
 m_sTempsTestAvant = txtTempsTestProj.Text
 m_sTempsInstallationAvant = txtTempsInstallationProj.Text
 m_sTempsFormationAvant = txtTempsFormationProj.Text
 m_sTempsGestionAvant = txtTempsGestionProj.Text
  m_sTempsShippingAvant = txtTempsShippingProj.Text
  m_sTempsPrototypeAvant = txtTempsPrototypeProj.Text

  Exit Sub

Oups:

  wOups "frmProjSoumMecTemps", "RemplirValeursVant", Err, Err.number, Err.Description
End Sub

Private Sub AfficherEnregistrement()

 On Error GoTo Oups

 Dim rstProjSoum As ADODB.Recordset
 Dim rstSoum As ADODB.Recordset
 Dim rstPunch As ADODB.Recordset
 Dim sChamps As String
 Dim sTable As String
 
 If m_eType = TYPE_PROJET Then
 sChamps = "IDProjet"
 sTable = "GrbProjetMec"
 Else
 sChamps = "IDSoumission"
  sTable = "GrbSoumissionMec"
  End If
 
  Set rstProjSoum = New ADODB.Recordset
 
  If m_eType = TYPE_PROJET Then
  Call rstProjSoum.Open("SELECT * FROM " & sTable & " WHERE " & sChamps & " = '" & m_sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
  Else
  Call rstProjSoum.Open("SELECT * FROM " & sTable & " WHERE " & sChamps & " = '" & m_sNoSoumission & "'", g_connData, adOpenDynamic, adLockOptimistic)
  End If
 
10 If Not rstProjSoum.EOF And FrmProjSoumMec.m_bTempsDejaOuvert = False And m_eMode = MODE_INACTIF Then
1 If m_eType = TYPE_SOUMISSION Then
 If Not IsNull(rstProjSoum.Fields("TempsDessin")) Then
 txtTempsDessinSoum.Text = rstProjSoum.Fields("TempsDessin")
 Else
 txtTempsDessinSoum.Text = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsCoupe")) Then
 txtTempsCoupeSoum.Text = rstProjSoum.Fields("TempsCoupe")
 Else
 txtTempsCoupeSoum.Text = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsMachinage")) Then
 txtTempsMachinageSoum.Text = rstProjSoum.Fields("TempsMachinage")
 Else
 txtTempsMachinageSoum.Text = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsSoudure")) Then
 txtTempsSoudureSoum.Text = rstProjSoum.Fields("TempsSoudure")
1  Else
 txtTempsSoudureSoum.Text = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsAssemblage")) Then
 txtTempsAssemblageSoum.Text = rstProjSoum.Fields("TempsAssemblage")
 Else
 txtTempsAssemblageSoum.Text = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsPeinture")) Then
 txtTempsPeintureSoum.Text = rstProjSoum.Fields("TempsPeinture")
 Else
 txtTempsPeintureSoum.Text = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsTest")) Then
 txtTempsTestSoum.Text = rstProjSoum.Fields("TempsTest")
 Else
 txtTempsTestSoum.Text = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsInstallation")) Then
 txtTempsInstallationSoum.Text = rstProjSoum.Fields("TempsInstallation")
 Else
 txtTempsInstallationSoum.Text = "0"
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

 If Not IsNull(rstProjSoum.Fields("TempsShipping")) Then
 txtTempsShippingSoum.Text = rstProjSoum.Fields("TempsShipping")
 Else
 txtTempsShippingSoum.Text = "0"
 End If
 txtTempsprototypeSoum.Text = "0"

 Else
 If m_sNoSoumission <> "" Then
 Set rstSoum = New ADODB.Recordset

 Call rstSoum.Open("SELECT * FROM GrbSoumissionMec WHERE IDSoumission = '" & m_sNoSoumission & "'", g_connData, adOpenDynamic, adLockOptimistic)

4 If Not rstSoum.EOF Then
4 If Not IsNull(rstSoum.Fields("TempsDessin")) Then
4 txtTempsDessinSoum.Text = rstSoum.Fields("TempsDessin")
4 Else
4 txtTempsDessinSoum.Text = "0"
4 End If

4 If Not IsNull(rstSoum.Fields("TempsCoupe")) Then
4 txtTempsCoupeSoum.Text = rstSoum.Fields("TempsCoupe")
4 Else
4 txtTempsCoupeSoum.Text = "0"
4 End If

4  If Not IsNull(rstSoum.Fields("TempsMachinage")) Then
4  txtTempsMachinageSoum.Text = rstSoum.Fields("TempsMachinage")
4  Else
4  txtTempsMachinageSoum.Text = "0"
4  End If

4  If Not IsNull(rstSoum.Fields("TempsSoudure")) Then
4  txtTempsSoudureSoum.Text = rstSoum.Fields("TempsSoudure")
4  Else
50 txtTempsSoudureSoum.Text = "0"
 End If

 If Not IsNull(rstSoum.Fields("TempsAssemblage")) Then
 txtTempsAssemblageSoum.Text = rstSoum.Fields("TempsAssemblage")
 Else
 txtTempsAssemblageSoum.Text = "0"
 End If

 If Not IsNull(rstSoum.Fields("TempsPeinture")) Then
 txtTempsPeintureSoum.Text = rstSoum.Fields("TempsPeinture")
 Else
 txtTempsPeintureSoum.Text = "0"
 End If
 
5  If Not IsNull(rstSoum.Fields("TempsTest")) Then
5  txtTempsTestSoum.Text = rstSoum.Fields("TempsTest")
5  Else
5  txtTempsTestSoum.Text = "0"
5  End If

5  If Not IsNull(rstSoum.Fields("TempsInstallation")) Then
5  txtTempsInstallationSoum.Text = rstSoum.Fields("TempsInstallation")
5  Else
60 txtTempsInstallationSoum.Text = "0"
  End If

  If Not IsNull(rstSoum.Fields("TempsFormation")) Then
  txtTempsFormationSoum.Text = rstSoum.Fields("TempsFormation")
  Else
  txtTempsFormationSoum.Text = "0"
  End If

  If Not IsNull(rstSoum.Fields("TempsGestion")) Then
  txtTempsGestionSoum.Text = rstSoum.Fields("TempsGestion")
  Else
  txtTempsGestionSoum.Text = "0"
  End If

6  If Not IsNull(rstSoum.Fields("TempsShipping")) Then
6  txtTempsShippingSoum.Text = rstSoum.Fields("TempsShipping")
6  Else
6  txtTempsShippingSoum.Text = "0"
6  End If
 txtTempsprototypeSoum.Text = "0"

6  Else
6  txtTempsDessinSoum.Text = 0
6  txtTempsCoupeSoum.Text = 0
70 txtTempsMachinageSoum.Text = 0
  txtTempsSoudureSoum.Text = 0
  txtTempsAssemblageSoum.Text = 0
  txtTempsPeintureSoum.Text = 0
  txtTempsTestSoum.Text = 0
  txtTempsInstallationSoum.Text = 0
  txtTempsFormationSoum.Text = 0
  txtTempsGestionSoum.Text = 0
  txtTempsShippingSoum.Text = 0
 txtTempsprototypeSoum.Text = 0

  End If

  Call rstSoum.Close
  Set rstSoum = Nothing
   Else
   txtTempsDessinSoum.Text = 0
7  txtTempsCoupeSoum.Text = 0
7  txtTempsMachinageSoum.Text = 0
7  txtTempsSoudureSoum.Text = 0
7  txtTempsAssemblageSoum.Text = 0
7  txtTempsPeintureSoum.Text = 0
7  txtTempsTestSoum.Text = 0
80 txtTempsInstallationSoum.Text = 0
  txtTempsFormationSoum.Text = 0
  txtTempsGestionSoum.Text = 0
  txtTempsShippingSoum.Text = 0
 txtTempsprototypeSoum.Text = 0
  End If

  m_bLocked = rstProjSoum.Fields("TempsProjBarré")

  If Not IsNull(rstProjSoum.Fields("TempsDessinProj")) Then
  txtTempsDessinProj.Text = rstProjSoum.Fields("TempsDessinProj")
  Else
  txtTempsDessinProj.Text = "0"
  End If

  If Not IsNull(rstProjSoum.Fields("TempsCoupeProj")) Then
   txtTempsCoupeProj.Text = rstProjSoum.Fields("TempsCoupeProj")
   Else
   txtTempsCoupeProj.Text = "0"
   End If

8  If Not IsNull(rstProjSoum.Fields("TempsMachinageProj")) Then
8  txtTempsMachinageProj.Text = rstProjSoum.Fields("TempsMachinageProj")
8  Else
8  txtTempsMachinageProj.Text = "0"
90 End If

  If Not IsNull(rstProjSoum.Fields("TempsSoudureProj")) Then
  txtTempsSoudureProj.Text = rstProjSoum.Fields("TempsSoudureProj")
  Else
  txtTempsSoudureProj.Text = "0"
  End If

  If Not IsNull(rstProjSoum.Fields("TempsAssemblageProj")) Then
  txtTempsAssemblageProj.Text = rstProjSoum.Fields("TempsAssemblageProj")
  Else
  txtTempsAssemblageProj.Text = "0"
  End If

  If Not IsNull(rstProjSoum.Fields("TempsPeintureProj")) Then
 txtTempsPeintureProj.Text = rstProjSoum.Fields("TempsPeintureProj")
   Else
 txtTempsPeintureProj.Text = "0"
   End If

 If Not IsNull(rstProjSoum.Fields("TempsTestProj")) Then
   txtTempsTestProj.Text = rstProjSoum.Fields("TempsTestProj")
 Else
9  txtTempsTestProj.Text = "0"
 End If

10 If Not IsNull(rstProjSoum.Fields("TempsInstallationProj")) Then
 txtTempsInstallationProj.Text = rstProjSoum.Fields("TempsInstallationProj")
1 Else
 txtTempsInstallationProj.Text = "0"
1 End If

 If Not IsNull(rstProjSoum.Fields("TempsFormationProj")) Then
 txtTempsFormationProj.Text = rstProjSoum.Fields("TempsFormationProj")
 Else
 txtTempsFormationProj.Text = "0"
 End If

1 If Not IsNull(rstProjSoum.Fields("TempsGestionProj")) Then
10  txtTempsGestionProj.Text = rstProjSoum.Fields("TempsGestionProj")
10  Else
10  txtTempsGestionProj.Text = "0"
10  End If

10  If Not IsNull(rstProjSoum.Fields("TempsShippingProj")) Then
10  txtTempsShippingProj.Text = rstProjSoum.Fields("TempsShippingProj")
10  Else
10  txtTempsShippingProj.Text = "0"
1 End If

 txtTempsPrototypeProj.Text = "0"

11 If m_bLocked = False Then
1 txtTempsDessinConc.Text = vbNullString
1 txtTempsCoupeConc.Text = vbNullString
1 txtTempsMachinageConc.Text = vbNullString
1 txtTempsSoudureConc.Text = vbNullString
1 txtTempsAssemblageConc.Text = vbNullString
1 txtTempsPeintureConc.Text = vbNullString
1 txtTempsTestConc.Text = vbNullString
1 txtTempsInstallationConc.Text = vbNullString
1 txtTempsFormationConc.Text = vbNullString
1 txtTempsGestionConc.Text = vbNullString
1 txtTempsShippingConc.Text = vbNullString
 txtTempsPrototypeConc.Text = vbNullString
1 Else
 If Not IsNull(rstProjSoum.Fields("TempsDessinConc")) Then
1 txtTempsDessinConc.Text = rstProjSoum.Fields("TempsDessinConc")
 Else
1 txtTempsDessinConc.Text = "0"
 End If

11  If Not IsNull(rstProjSoum.Fields("TempsCoupeConc")) Then
 txtTempsCoupeConc.Text = rstProjSoum.Fields("TempsCoupeConc")
 Else
1 txtTempsCoupeConc.Text = "0"
1 End If

1 If Not IsNull(rstProjSoum.Fields("TempsMachinageConc")) Then
1 txtTempsMachinageConc.Text = rstProjSoum.Fields("TempsMachinageConc")
1 Else
1 txtTempsMachinageConc.Text = "0"
1 End If

1 If Not IsNull(rstProjSoum.Fields("TempsSoudureConc")) Then
1 txtTempsSoudureConc.Text = rstProjSoum.Fields("TempsSoudureConc")
1 Else
1 txtTempsSoudureConc.Text = "0"
1 End If

1 If Not IsNull(rstProjSoum.Fields("TempsAssemblageConc")) Then
1 txtTempsAssemblageConc.Text = rstProjSoum.Fields("TempsAssemblageConc")
1 Else
1 txtTempsAssemblageConc.Text = "0"
1 End If

1 If Not IsNull(rstProjSoum.Fields("TempsPeintureConc")) Then
1 txtTempsPeintureConc.Text = rstProjSoum.Fields("TempsPeintureConc")
1 Else
1 txtTempsPeintureConc.Text = "0"
1 End If

1 If Not IsNull(rstProjSoum.Fields("TempsTestConc")) Then
1 txtTempsTestConc.Text = rstProjSoum.Fields("TempsTestConc")
1 Else
1 txtTempsTestConc.Text = "0"
1 End If
 
1 If Not IsNull(rstProjSoum.Fields("TempsInstallationConc")) Then
1 txtTempsInstallationConc.Text = rstProjSoum.Fields("TempsInstallationConc")
1 Else
1 txtTempsInstallationConc.Text = "0"
1 End If

1 If Not IsNull(rstProjSoum.Fields("TempsFormationConc")) Then
1 txtTempsFormationConc.Text = rstProjSoum.Fields("TempsFormationConc")
1 Else
1 txtTempsFormationConc.Text = "0"
1 End If

1 If Not IsNull(rstProjSoum.Fields("TempsGestionConc")) Then
1 txtTempsGestionConc.Text = rstProjSoum.Fields("TempsGestionConc")
14 Else
14 txtTempsGestionConc.Text = "0"
14 End If

14 If Not IsNull(rstProjSoum.Fields("TempsShippingConc")) Then
14 txtTempsShippingConc.Text = rstProjSoum.Fields("TempsShippingConc")
14 Else
14 txtTempsShippingConc.Text = "0"
14 End If
 txtTempsPrototypeConc.Text = "0"

14 End If
14 End If

14 If m_eType = TYPE_PROJET Then
14  Call AfficherTempsReels

14  Call CalculerTotalArgent
14  End If

14  If Not IsNull(rstProjSoum.Fields("NbrePersonne")) Then
14  txtNbrePersonne.Text = rstProjSoum.Fields("NbrePersonne")
14  Else
14  txtNbrePersonne.Text = "0"
14  End If

150 If Not IsNull(rstProjSoum.Fields("TempsHebergement")) Then
15 txtTempsHebergement.Text = rstProjSoum.Fields("TempsHebergement")
 Else
 txtTempsHebergement.Text = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsRepas")) Then
 txtTempsRepas.Text = rstProjSoum.Fields("TempsRepas")
 Else
 txtTempsRepas.Text = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsTransport")) Then
 txtTempsDeplacement.Text = rstProjSoum.Fields("TempsTransport")
15  Else
15  txtTempsDeplacement.Text = "0"
15  End If

15  If Not IsNull(rstProjSoum.Fields("TempsUniteMobile")) Then
15  txtTempsUniteMobile.Text = rstProjSoum.Fields("TempsUniteMobile")
15  Else
15  txtTempsUniteMobile.Text = "0"
15  End If

160 txtPrixEmballage.Text = rstProjSoum.Fields("PrixEmballage")
160 Else
 If m_eType = TYPE_SOUMISSION Then
 txtTempsDessinSoum.Text = FrmProjSoumMec.m_sTempsDessin
 txtTempsCoupeSoum.Text = FrmProjSoumMec.m_sTempsCoupe
 txtTempsMachinageSoum.Text = FrmProjSoumMec.m_sTempsMachinage
 txtTempsSoudureSoum.Text = FrmProjSoumMec.m_sTempsSoudure
 txtTempsAssemblageSoum.Text = FrmProjSoumMec.m_sTempsAssemblage
 txtTempsPeintureSoum.Text = FrmProjSoumMec.m_sTempsPeinture
 txtTempsTestSoum.Text = FrmProjSoumMec.m_sTempsTest
 txtTempsInstallationSoum.Text = FrmProjSoumMec.m_sTempsInstallation
 txtTempsFormationSoum.Text = FrmProjSoumMec.m_sTempsFormation
16  txtTempsGestionSoum.Text = FrmProjSoumMec.m_sTempsGestion
16  txtTempsShippingSoum.Text = FrmProjSoumMec.m_sTempsShipping
16  Else
16  If m_sNoSoumission <> "" Then
16  Set rstSoum = New ADODB.Recordset

16  Call rstSoum.Open("SELECT * FROM GrbSoumissionMec WHERE IDSoumission = '" & m_sNoSoumission & "'", g_connData, adOpenDynamic, adLockOptimistic)

16  If Not rstSoum.EOF Then
16  If Not IsNull(rstSoum.Fields("TempsDessin")) Then
170 txtTempsDessinSoum.Text = rstSoum.Fields("TempsDessin")
 Else
 txtTempsDessinSoum.Text = "0"
 End If

 If Not IsNull(rstSoum.Fields("TempsCoupe")) Then
 txtTempsCoupeSoum.Text = rstSoum.Fields("TempsCoupe")
 Else
 txtTempsCoupeSoum.Text = "0"
 End If

 If Not IsNull(rstSoum.Fields("TempsMachinage")) Then
 txtTempsMachinageSoum.Text = rstSoum.Fields("TempsMachinage")
 Else
1   txtTempsMachinageSoum.Text = "0"
1   End If

17  If Not IsNull(rstSoum.Fields("TempsSoudure")) Then
17  txtTempsSoudureSoum.Text = rstSoum.Fields("TempsSoudure")
17  Else
17  txtTempsSoudureSoum.Text = "0"
17  End If

17  If Not IsNull(rstSoum.Fields("TempsAssemblage")) Then
180 txtTempsAssemblageSoum.Text = rstSoum.Fields("TempsAssemblage")
 Else
 txtTempsAssemblageSoum.Text = "0"
 End If

 If Not IsNull(rstSoum.Fields("TempsPeinture")) Then
 txtTempsPeintureSoum.Text = rstSoum.Fields("TempsPeinture")
 Else
 txtTempsPeintureSoum.Text = "0"
 End If

 If Not IsNull(rstSoum.Fields("TempsTest")) Then
 txtTempsTestSoum.Text = rstSoum.Fields("TempsTest")
 Else
1   txtTempsTestSoum.Text = "0"
1   End If
 
1   If Not IsNull(rstSoum.Fields("TempsInstallation")) Then
1   txtTempsInstallationSoum.Text = rstSoum.Fields("TempsInstallation")
18  Else
18  txtTempsInstallationSoum.Text = "0"
18  End If

18  If Not IsNull(rstSoum.Fields("TempsFormation")) Then
190 txtTempsFormationSoum.Text = rstSoum.Fields("TempsFormation")
 Else
1  txtTempsFormationSoum.Text = "0"
1  End If

1  If Not IsNull(rstSoum.Fields("TempsGestion")) Then
1  txtTempsGestionSoum.Text = rstSoum.Fields("TempsGestion")
1  Else
1  txtTempsGestionSoum.Text = "0"
1  End If

1  If Not IsNull(rstSoum.Fields("TempsShipping")) Then
1  txtTempsShippingSoum.Text = rstSoum.Fields("TempsShipping")
1  Else
 txtTempsShippingSoum.Text = "0"
1   End If
 txtTempsprototypeSoum.Text = "0"
 Else
1   txtTempsDessinSoum.Text = 0
 txtTempsCoupeSoum.Text = 0
1   txtTempsMachinageSoum.Text = 0
 txtTempsSoudureSoum.Text = 0
19  txtTempsAssemblageSoum.Text = 0
200 txtTempsPeintureSoum.Text = 0
 txtTempsTestSoum.Text = 0
 txtTempsInstallationSoum.Text = 0
 txtTempsFormationSoum.Text = 0
 txtTempsGestionSoum.Text = 0
 txtTempsShippingSoum.Text = 0
 txtTempsprototypeSoum.Text = 0
 End If

 Call rstSoum.Close
 Set rstSoum = Nothing
 Else
 txtTempsDessinSoum.Text = 0
 txtTempsCoupeSoum.Text = 0
20  txtTempsMachinageSoum.Text = 0
20  txtTempsSoudureSoum.Text = 0
20  txtTempsAssemblageSoum.Text = 0
20  txtTempsPeintureSoum.Text = 0
20  txtTempsTestSoum.Text = 0
20  txtTempsInstallationSoum.Text = 0
20  txtTempsFormationSoum.Text = 0
20  txtTempsGestionSoum.Text = 0
txtTempsShippingSoum.Text = 0
 txtTempsprototypeSoum.Text = 0
21 End If

m_bLocked = FrmProjSoumMec.m_bTempsProjLock

If m_bLocked = True Then
 If Not IsNull(rstProjSoum.Fields("TempsDessinProj")) Then
 txtTempsDessinProj.Text = rstProjSoum.Fields("TempsDessinProj")
 Else
 txtTempsDessinProj.Text = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsCoupeProj")) Then
 txtTempsCoupeProj.Text = rstProjSoum.Fields("TempsCoupeProj")
 Else
 txtTempsCoupeProj.Text = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsMachinageProj")) Then
 txtTempsMachinageProj.Text = rstProjSoum.Fields("TempsMachinageProj")
 Else
 txtTempsMachinageProj.Text = "0"
 End If

21  If Not IsNull(rstProjSoum.Fields("TempsSoudureProj")) Then
 txtTempsSoudureProj.Text = rstProjSoum.Fields("TempsSoudureProj")
 Else
2 txtTempsSoudureProj.Text = "0"
2 End If

2 If Not IsNull(rstProjSoum.Fields("TempsAssemblageProj")) Then
2 txtTempsAssemblageProj.Text = rstProjSoum.Fields("TempsAssemblageProj")
2 Else
2 txtTempsAssemblageProj.Text = "0"
2 End If

2 If Not IsNull(rstProjSoum.Fields("TempsPeintureProj")) Then
2 txtTempsPeintureProj.Text = rstProjSoum.Fields("TempsPeintureProj")
2 Else
2 txtTempsPeintureProj.Text = "0"
2 End If

2 If Not IsNull(rstProjSoum.Fields("TempsTestProj")) Then
2 txtTempsTestProj.Text = rstProjSoum.Fields("TempsTestProj")
2 Else
2 txtTempsTestProj.Text = "0"
2 End If

2 If Not IsNull(rstProjSoum.Fields("TempsInstallationProj")) Then
2 txtTempsInstallationProj.Text = rstProjSoum.Fields("TempsInstallationProj")
2 Else
2 txtTempsInstallationProj.Text = "0"
2 End If

2 If Not IsNull(rstProjSoum.Fields("TempsFormationProj")) Then
2 txtTempsFormationProj.Text = rstProjSoum.Fields("TempsFormationProj")
2 Else
2 txtTempsFormationProj.Text = "0"
2 End If

2 If Not IsNull(rstProjSoum.Fields("TempsGestionProj")) Then
2 txtTempsGestionProj.Text = rstProjSoum.Fields("TempsGestionProj")
2 Else
2 txtTempsGestionProj.Text = "0"
2 End If

2 If Not IsNull(rstProjSoum.Fields("TempsShippingProj")) Then
2 txtTempsShippingProj.Text = rstProjSoum.Fields("TempsShippingProj")
2 Else
2 txtTempsShippingProj.Text = "0"
2 End If

 txtTempsPrototypeProj.Text = "0"

2 txtTempsDessinConc.Text = FrmProjSoumMec.m_sTempsDessinConc
2 txtTempsCoupeConc.Text = FrmProjSoumMec.m_sTempsCoupeConc
24 txtTempsMachinageConc.Text = FrmProjSoumMec.m_sTempsMachinageConc
24 txtTempsSoudureConc.Text = FrmProjSoumMec.m_sTempsSoudureConc
24 txtTempsAssemblageConc.Text = FrmProjSoumMec.m_sTempsAssemblageConc
24 txtTempsPeintureConc.Text = FrmProjSoumMec.m_sTempsPeintureConc
24 txtTempsTestConc.Text = FrmProjSoumMec.m_sTempsTestConc
24 txtTempsInstallationConc.Text = FrmProjSoumMec.m_sTempsInstallationConc
24 txtTempsFormationConc.Text = FrmProjSoumMec.m_sTempsFormationConc
24 txtTempsGestionConc.Text = FrmProjSoumMec.m_sTempsGestionConc
24 txtTempsShippingConc.Text = FrmProjSoumMec.m_sTempsShippingConc
 txtTempsPrototypeConc.Text = FrmProjSoumMec.m_sTempsPrototypeConc
24 Else
24 txtTempsDessinProj.Text = FrmProjSoumMec.m_sTempsDessinProj
24  txtTempsCoupeProj.Text = FrmProjSoumMec.m_sTempsCoupeProj
24  txtTempsMachinageProj.Text = FrmProjSoumMec.m_sTempsMachinageProj
24  txtTempsSoudureProj.Text = FrmProjSoumMec.m_sTempsSoudureProj
24  txtTempsAssemblageProj.Text = FrmProjSoumMec.m_sTempsAssemblageProj
24  txtTempsPeintureProj.Text = FrmProjSoumMec.m_sTempsPeintureProj
24  txtTempsTestProj.Text = FrmProjSoumMec.m_sTempsTestProj
24  txtTempsInstallationProj.Text = FrmProjSoumMec.m_sTempsInstallationProj
24  txtTempsFormationProj.Text = FrmProjSoumMec.m_sTempsFormationProj
250 txtTempsGestionProj.Text = FrmProjSoumMec.m_sTempsGestionProj
2 txtTempsShippingProj.Text = FrmProjSoumMec.m_sTempsShippingProj
 txtTempsPrototypeProj.Text = FrmProjSoumMec.m_sTempsPrototypeProj
 End If
2 End If

 If m_eType = TYPE_PROJET Then
 Call AfficherTempsReels

 Call CalculerTotalArgent
 End If

 txtNbrePersonne.Text = FrmProjSoumMec.m_sNbrePersonne
 txtTempsHebergement.Text = FrmProjSoumMec.m_sTempsHebergement
 txtTempsRepas.Text = FrmProjSoumMec.m_sTempsRepas
 txtTempsDeplacement.Text = FrmProjSoumMec.m_sTempsTransport
25  txtTempsUniteMobile.Text = FrmProjSoumMec.m_sTempsUniteMobile
25  txtPrixEmballage.Text = FrmProjSoumMec.m_sPrixEmballage
257End If
 
25  Call rstProjSoum.Close
25  Set rstProjSoum = Nothing

25  Exit Sub

Oups:

25  wOups "frmProjSoumMecTemps", "AfficherEnregistrement", Err, Err.number, Err.Description
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

 If Right$(m_sNoProjet, 2) = "99" Then
 sFilterNoProjet = "LEFT(NoProjet, 6) = '" & Left$(m_sNoProjet, 6) & "'"
  Else
  sFilterNoProjet = "NoProjet = '" & m_sNoProjet & "'"
  End If

  Set rstPunch = New ADODB.Recordset

  Call rstPunch.Open("SELECT Type, " & sTotal & " FROM GrbPunch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null GROUP BY Type", g_connData, adOpenDynamic, adLockOptimistic)

  lblTempsDessinReel.Caption = "0"
  lblTempsCoupeReel.Caption = "0"
  lblTempsMachinageReel.Caption = "0"
10 lblTempsSoudureReel.Caption = "0"
lblTempsAssemblageReel.Caption = "0"
lblTempsPeintureReel.Caption = "0"
lblTempsTestReel.Caption = "0"
lblTempsInstallationReel.Caption = "0"
lblTempsFormationReel.Caption = "0"
lblTempsGestionReel.Caption = "0"
lblTempsShippingReel.Caption = "0"
 lblTempsPrototypeReel.Caption = "0"

Do While Not rstPunch.EOF
 If Not IsNull(rstPunch.Fields("Total")) Then
 Select Case rstPunch.Fields("Type")
 Case "Dessin": lblTempsDessinReel.Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Coupe": lblTempsCoupeReel.Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Machinage": lblTempsMachinageReel.Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Soudure": lblTempsSoudureReel.Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Assemblage": lblTempsAssemblageReel.Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Peinture": lblTempsPeintureReel.Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Test": lblTempsTestReel.Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Installation": lblTempsInstallationReel.Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Formation": lblTempsFormationReel.Caption = Round(rstPunch.Fields("Total"), 2)
1  Case "Gestion": lblTempsGestionReel.Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Shipping": lblTempsShippingReel.Caption = Round(rstPunch.Fields("Total"), 2)
 Case "Prototypage-Dévelloppement expérimental": lblTempsPrototypeReel.Caption = Round(rstPunch.Fields("Total"), 2)
 End Select
 End If

 Call rstPunch.MoveNext
Loop

Call rstPunch.Close

 'Ouverture des enregistrements avec comme filtre, le numéro du projet
Call rstPunch.Open("SELECT " & sTotal & " FROM GrbPunch WHERE " & sFilterNoProjet & " AND HeureFin is not null AND HeureDébut is not null", g_connData, adOpenDynamic, adLockOptimistic)

If Not IsNull(rstPunch.Fields("Total")) Then
 lblTotalTempsRHReel.Caption = Round(rstPunch.Fields("Total"), 2)
Else
 lblTotalTempsRHReel.Caption = "0"
End If

2  Call rstPunch.Close
Set rstPunch = Nothing

2  Exit Sub

Oups:

wOups "frmProjSoumMecTemps", "AfficherTempsReels", Err, Err.number, Err.Description
End Sub

Private Sub CalculerTotalArgent()
 
 On Error GoTo Oups

 If IsNumeric(lblTempsDessinReel.Caption) Then
 lblPrixDessin.Caption = Round(Replace(lblTempsDessinReel.Caption * m_sTauxDessin, ".", ","), 2)
 Else
 lblPrixDessin.Caption = 0
 End If

 If IsNumeric(lblTempsCoupeReel.Caption) Then
 lblPrixCoupe.Caption = Round(Replace(lblTempsCoupeReel.Caption * m_sTauxCoupe, ".", ","), 2)
 Else
 lblPrixCoupe.Caption = 0
 End If

  If IsNumeric(lblTempsMachinageReel.Caption) Then
  lblPrixMachinage.Caption = Round(Replace(lblTempsMachinageReel.Caption * m_sTauxMachinage, ".", ","), 2)
  Else
  lblPrixMachinage.Caption = 0
  End If

  If IsNumeric(lblTempsSoudureReel.Caption) Then
  lblPrixSoudure.Caption = Round(Replace(lblTempsSoudureReel.Caption * m_sTauxSoudure, ".", ","), 2)
  Else
lblPrixSoudure.Caption = 0
End If

If IsNumeric(lblTempsAssemblageReel.Caption) Then
 lblPrixAssemblage.Caption = Round(Replace(lblTempsAssemblageReel.Caption * m_sTauxAssemblage, ".", ","), 2)
Else
 lblPrixAssemblage.Caption = 0
End If

If IsNumeric(lblTempsPeintureReel.Caption) Then
 lblPrixPeinture.Caption = Round(Replace(lblTempsPeintureReel.Caption * m_sTauxPeinture, ".", ","), 2)
Else
 lblPrixPeinture.Caption = 0
End If

1  If IsNumeric(lblTempsTestReel.Caption) Then
 lblPrixTest.Caption = Round(Replace(lblTempsTestReel.Caption * m_sTauxTest, ".", ","), 2)
 Else
 lblPrixTest.Caption = 0
 End If

If IsNumeric(lblTempsInstallationReel.Caption) Then
 lblPrixInstallation.Caption = Round(Replace(lblTempsInstallationReel.Caption * m_sTauxInstallation, ".", ","), 2)
1  Else
 lblPrixInstallation.Caption = 0
 End If

If IsNumeric(lblTempsFormationReel.Caption) Then
 lblPrixFormation.Caption = Round(Replace(lblTempsFormationReel.Caption * m_sTauxFormation, ".", ","), 2)
Else
 lblPrixFormation.Caption = 0
End If

If IsNumeric(lblTempsGestionReel.Caption) Then
 lblPrixGestion.Caption = Round(Replace(lblTempsGestionReel.Caption * m_sTauxGestion, ".", ","), 2)
Else
 lblPrixGestion.Caption = 0
End If

2  If IsNumeric(lblTempsShippingReel.Caption) Then
 lblPrixShipping.Caption = Round(Replace(lblTempsShippingReel.Caption * m_sTauxShipping, ".", ","), 2)
2  Else
 lblPrixShipping.Caption = 0
2  End If

 If IsNumeric(lblTempsPrototypeReel.Caption) Then
 lblPrixPrototype.Caption = Round(Replace(lblTempsPrototypeReel.Caption * m_sTauxGestion, ".", ","), 2)
2  Else
2  lblPrixPrototype.Caption = 0
284 End If


Call CalculerTotal

2  Exit Sub

Oups:

wOups "frmProjSoumMecTemps", "CalculerTotalArgent", Err, Err.number, Err.Description
End Sub

Private Sub BarrerChamps(ByVal bLocked As Boolean)

 On Error GoTo Oups

 If m_eType = TYPE_SOUMISSION Then
 txtTempsDessinSoum.Enabled = True
 txtTempsCoupeSoum.Enabled = True
 txtTempsMachinageSoum.Enabled = True
 txtTempsSoudureSoum.Enabled = True
 txtTempsAssemblageSoum.Enabled = True
 txtTempsPeintureSoum.Enabled = True
 txtTempsTestSoum.Enabled = True
 txtTempsInstallationSoum.Enabled = True
 txtTempsFormationSoum.Enabled = True
  txtTempsGestionSoum.Enabled = True
  txtTempsShippingSoum.Enabled = True
 txtTempsprototypeSoum.Enabled = True
 

  txtTempsDessinSoum.Locked = bLocked
  txtTempsCoupeSoum.Locked = bLocked
  txtTempsMachinageSoum.Locked = bLocked
  txtTempsSoudureSoum.Locked = bLocked
  txtTempsAssemblageSoum.Locked = bLocked
  txtTempsPeintureSoum.Locked = bLocked
txtTempsTestSoum.Locked = bLocked
1 txtTempsInstallationSoum.Locked = bLocked
 txtTempsFormationSoum.Locked = bLocked
 txtTempsGestionSoum.Locked = bLocked
 txtTempsShippingSoum.Locked = bLocked
 txtTempsprototypeSoum.Locked = bLocked



 txtTempsDessinProj.Enabled = False
 txtTempsCoupeProj.Enabled = False
 txtTempsMachinageProj.Enabled = False
 txtTempsSoudureProj.Enabled = False
 txtTempsAssemblageProj.Enabled = False
 txtTempsPeintureProj.Enabled = False
 txtTempsTestProj.Enabled = False
txtTempsInstallationProj.Enabled = False
 txtTempsFormationProj.Enabled = False
 txtTempsGestionProj.Enabled = False
 txtTempsShippingProj.Enabled = False
 txtTempsPrototypeProj.Enabled = False

 txtTempsDessinConc.Enabled = False
 txtTempsCoupeConc.Enabled = False
 txtTempsMachinageConc.Enabled = False
1  txtTempsSoudureConc.Enabled = False
 txtTempsAssemblageConc.Enabled = False
 txtTempsPeintureConc.Enabled = False
 txtTempsTestConc.Enabled = False
 txtTempsInstallationConc.Enabled = False
 txtTempsFormationConc.Enabled = False
 txtTempsGestionConc.Enabled = False
 txtTempsShippingConc.Enabled = False
 txtTempsPrototypeConc.Enabled = False


 cmdLock.Visible = False
 cmdUnlock.Visible = False
Else
 If m_bLocked = False Then
 txtTempsDessinProj.Enabled = True
 txtTempsCoupeProj.Enabled = True
 txtTempsMachinageProj.Enabled = True
 txtTempsSoudureProj.Enabled = True
 txtTempsAssemblageProj.Enabled = True
 txtTempsPeintureProj.Enabled = True
 txtTempsTestProj.Enabled = True
 txtTempsInstallationProj.Enabled = True
 txtTempsFormationProj.Enabled = True
 txtTempsGestionProj.Enabled = True
txtTempsShippingProj.Enabled = True
 txtTempsPrototypeProj.Enabled = True


 txtTempsDessinProj.Locked = bLocked
 txtTempsCoupeProj.Locked = bLocked
 txtTempsMachinageProj.Locked = bLocked
 txtTempsSoudureProj.Locked = bLocked
 txtTempsAssemblageProj.Locked = bLocked
 txtTempsPeintureProj.Locked = bLocked
 txtTempsTestProj.Locked = bLocked
 txtTempsInstallationProj.Locked = bLocked
 txtTempsFormationProj.Locked = bLocked
 txtTempsGestionProj.Locked = bLocked
 txtTempsShippingProj.Locked = bLocked
 txtTempsPrototypeProj.Locked = bLocked


 txtTempsDessinSoum.Enabled = False
 txtTempsCoupeSoum.Enabled = False
 txtTempsMachinageSoum.Enabled = False
 txtTempsSoudureSoum.Enabled = False
 txtTempsAssemblageSoum.Enabled = False
 txtTempsPeintureSoum.Enabled = False
 txtTempsTestSoum.Enabled = False
 txtTempsInstallationSoum.Enabled = False
4 txtTempsFormationSoum.Enabled = False
4 txtTempsGestionSoum.Enabled = False
4 txtTempsShippingSoum.Enabled = False
4 txtTempsprototypeSoum.Enabled = False



4 txtTempsDessinConc.Enabled = False
4 txtTempsCoupeConc.Enabled = False
4 txtTempsMachinageConc.Enabled = False
4 txtTempsSoudureConc.Enabled = False
4 txtTempsAssemblageConc.Enabled = False
4 txtTempsPeintureConc.Enabled = False
4 txtTempsTestConc.Enabled = False
4 txtTempsInstallationConc.Enabled = False
4  txtTempsFormationConc.Enabled = False
4  txtTempsGestionConc.Enabled = False
4  txtTempsShippingConc.Enabled = False
4  txtTempsPrototypeConc.Enabled = False

4  If m_eMode = MODE_AJOUT_MODIF Then
4  If g_bVerrouillageTempsProjet = True Then
4  cmdLock.Visible = True
4  Else
4  cmdLock.Visible = False
50 End If

 cmdUnlock.Visible = False
 Else
 cmdLock.Visible = False
 cmdUnlock.Visible = False
 End If
 Else
 txtTempsDessinConc.Enabled = True
 txtTempsCoupeConc.Enabled = True
 txtTempsMachinageConc.Enabled = True
 txtTempsSoudureConc.Enabled = True
 txtTempsAssemblageConc.Enabled = True
5  txtTempsPeintureConc.Enabled = True
5  txtTempsTestConc.Enabled = True
5  txtTempsInstallationConc.Enabled = True
5  txtTempsFormationConc.Enabled = True
5  txtTempsGestionConc.Enabled = True
5  txtTempsShippingConc.Enabled = True
5   txtTempsPrototypeConc.Enabled = True


5  txtTempsDessinConc.Locked = bLocked
5  txtTempsCoupeConc.Locked = bLocked
60 txtTempsMachinageConc.Locked = bLocked
  txtTempsSoudureConc.Locked = bLocked
  txtTempsAssemblageConc.Locked = bLocked
  txtTempsPeintureConc.Locked = bLocked
  txtTempsTestConc.Locked = bLocked
  txtTempsInstallationConc.Locked = bLocked
  txtTempsFormationConc.Locked = bLocked
  txtTempsGestionConc.Locked = bLocked
  txtTempsShippingConc.Locked = bLocked
64 txtTempsPrototypeConc.Locked = bLocked

  txtTempsDessinSoum.Enabled = False
  txtTempsCoupeSoum.Enabled = False
  txtTempsMachinageSoum.Enabled = False
6  txtTempsSoudureSoum.Enabled = False
6  txtTempsAssemblageSoum.Enabled = False
6  txtTempsPeintureSoum.Enabled = False
6  txtTempsTestSoum.Enabled = False
6  txtTempsInstallationSoum.Enabled = False
6  txtTempsFormationSoum.Enabled = False
6  txtTempsGestionSoum.Enabled = False
6  txtTempsShippingSoum.Enabled = False
6   txtTempsprototypeSoum.Enabled = False


70 txtTempsDessinProj.Enabled = False
  txtTempsCoupeProj.Enabled = False
  txtTempsMachinageProj.Enabled = False
  txtTempsSoudureProj.Enabled = False
  txtTempsAssemblageProj.Enabled = False
  txtTempsPeintureProj.Enabled = False
  txtTempsTestProj.Enabled = False
  txtTempsInstallationProj.Enabled = False
  txtTempsFormationProj.Enabled = False
  txtTempsGestionProj.Enabled = False
  txtTempsShippingProj.Enabled = False
75 txtTempsPrototypeProj.Enabled = False


  If m_eMode = MODE_AJOUT_MODIF Then
   If g_bDeverrouillageTempsProjet = True Then
   cmdUnlock.Visible = True
7  Else
7  cmdUnlock.Visible = False
7  End If

7  cmdLock.Visible = False
7  Else
7  cmdLock.Visible = False
80 cmdUnlock.Visible = False
  End If
  End If
  End If
 
  txtNbrePersonne.Locked = bLocked
  txtTempsHebergement.Locked = bLocked
  txtTempsRepas.Locked = bLocked
  txtTempsDeplacement.Locked = bLocked
  txtTempsUniteMobile.Locked = bLocked
 
  txtPrixEmballage.Locked = bLocked

  Exit Sub

Oups:

  wOups "frmProjSoumMecTemps", "BarrerChamps", Err, Err.number, Err.Description
End Sub

Private Sub cmdDetail_Click()

 On Error GoTo Oups

 If m_eType = TYPE_PROJET Then
 Call frmDetailTemps.Afficher(m_sNoProjet, MECANIQUE, True)
 Else
 Call frmDetailTemps.Afficher(m_sNoSoumission, MECANIQUE, False)
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "cmdDetail_Click", Err, Err.number, Err.Description
End Sub

Private Sub Cmdfermer_Click()

 On Error GoTo Oups

 If m_eMode = MODE_AJOUT_MODIF Then
 Call EnregistrerTemps
 End If

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "cmdFermer_Click", Err, Err.number, Err.Description
End Sub

Private Sub EnregistrerTemps()

 On Error GoTo Oups

 If m_eType = TYPE_SOUMISSION Then
 If Trim$(txtTempsDessinSoum.Text) <> vbNullString And IsNumeric(txtTempsDessinSoum.Text) Then
 FrmProjSoumMec.m_sTempsDessin = txtTempsDessinSoum.Text
 Else
 FrmProjSoumMec.m_sTempsDessin = "0"
 End If

 If Trim$(txtTempsCoupeSoum.Text) <> vbNullString And IsNumeric(txtTempsCoupeSoum.Text) Then
 FrmProjSoumMec.m_sTempsCoupe = txtTempsCoupeSoum.Text
 Else
 FrmProjSoumMec.m_sTempsCoupe = "0"
  End If

  If Trim$(txtTempsMachinageSoum.Text) <> vbNullString And IsNumeric(txtTempsMachinageSoum.Text) Then
  FrmProjSoumMec.m_sTempsMachinage = txtTempsMachinageSoum.Text
  Else
  FrmProjSoumMec.m_sTempsMachinage = "0"
  End If
 
  If Trim$(txtTempsSoudureSoum.Text) <> vbNullString And IsNumeric(txtTempsSoudureSoum.Text) Then
  FrmProjSoumMec.m_sTempsSoudure = txtTempsSoudureSoum.Text
Else
FrmProjSoumMec.m_sTempsSoudure = "0"
 End If
 
 If Trim$(txtTempsAssemblageSoum.Text) <> vbNullString And IsNumeric(txtTempsAssemblageSoum.Text) Then
 FrmProjSoumMec.m_sTempsAssemblage = txtTempsAssemblageSoum.Text
 Else
 FrmProjSoumMec.m_sTempsAssemblage = "0"
 End If
 
 If Trim$(txtTempsPeintureSoum.Text) <> vbNullString And IsNumeric(txtTempsPeintureSoum.Text) Then
 FrmProjSoumMec.m_sTempsPeinture = txtTempsPeintureSoum.Text
 Else
 FrmProjSoumMec.m_sTempsPeinture = "0"
End If
 
 If Trim$(txtTempsTestSoum.Text) <> vbNullString And IsNumeric(txtTempsTestSoum.Text) Then
 FrmProjSoumMec.m_sTempsTest = txtTempsTestSoum.Text
 Else
 FrmProjSoumMec.m_sTempsTest = "0"
 End If

 If Trim$(txtTempsInstallationSoum.Text) <> vbNullString And IsNumeric(txtTempsInstallationSoum.Text) Then
1  FrmProjSoumMec.m_sTempsInstallation = txtTempsInstallationSoum.Text
 Else
 FrmProjSoumMec.m_sTempsInstallation = "0"
 End If

 If Trim$(txtTempsFormationSoum.Text) <> vbNullString And IsNumeric(txtTempsFormationSoum.Text) Then
 FrmProjSoumMec.m_sTempsFormation = txtTempsFormationSoum.Text
 Else
 FrmProjSoumMec.m_sTempsFormation = "0"
 End If

 If Trim$(txtTempsGestionSoum.Text) <> vbNullString And IsNumeric(txtTempsGestionSoum.Text) Then
 FrmProjSoumMec.m_sTempsGestion = txtTempsGestionSoum.Text
 Else
 FrmProjSoumMec.m_sTempsGestion = "0"
End If

 If Trim$(txtTempsShippingSoum.Text) <> vbNullString And IsNumeric(txtTempsShippingSoum.Text) Then
 FrmProjSoumMec.m_sTempsShipping = txtTempsShippingSoum.Text
 Else
 FrmProjSoumMec.m_sTempsShipping = "0"
 End If
2  Else
 FrmProjSoumMec.m_bTempsProjLock = m_bLocked

If m_bLocked = False Then
If Trim$(txtTempsDessinProj.Text) <> vbNullString And IsNumeric(txtTempsDessinProj.Text) Then
 FrmProjSoumMec.m_sTempsDessinProj = txtTempsDessinProj.Text
 Else
 FrmProjSoumMec.m_sTempsDessinProj = "0"
 End If

 If Trim$(txtTempsCoupeProj.Text) <> vbNullString And IsNumeric(txtTempsMachinageProj.Text) Then
 FrmProjSoumMec.m_sTempsCoupeProj = txtTempsCoupeProj.Text
 Else
 FrmProjSoumMec.m_sTempsCoupeProj = "0"
 End If

 If Trim$(txtTempsMachinageProj.Text) <> vbNullString And IsNumeric(txtTempsMachinageProj.Text) Then
 FrmProjSoumMec.m_sTempsMachinageProj = txtTempsMachinageProj.Text
 Else
 FrmProjSoumMec.m_sTempsMachinageProj = "0"
 End If

 If Trim$(txtTempsSoudureProj.Text) <> vbNullString And IsNumeric(txtTempsSoudureProj.Text) Then
 FrmProjSoumMec.m_sTempsSoudureProj = txtTempsSoudureProj.Text
 Else
 FrmProjSoumMec.m_sTempsSoudureProj = "0"
 End If

4 If Trim$(txtTempsAssemblageProj.Text) <> vbNullString And IsNumeric(txtTempsAssemblageProj.Text) Then
4 FrmProjSoumMec.m_sTempsAssemblageProj = txtTempsAssemblageProj.Text
4 Else
4 FrmProjSoumMec.m_sTempsAssemblageProj = "0"
4 End If

4 If Trim$(txtTempsPeintureProj.Text) <> vbNullString And IsNumeric(txtTempsPeintureProj.Text) Then
4 FrmProjSoumMec.m_sTempsPeintureProj = txtTempsPeintureProj.Text
4 Else
4 FrmProjSoumMec.m_sTempsPeintureProj = "0"
4 End If

4 If Trim$(txtTempsTestProj.Text) <> vbNullString And IsNumeric(txtTempsTestProj.Text) Then
4  FrmProjSoumMec.m_sTempsTestProj = txtTempsTestProj.Text
4  Else
4  FrmProjSoumMec.m_sTempsTestProj = "0"
4  End If

4  If Trim$(txtTempsInstallationProj.Text) <> vbNullString And IsNumeric(txtTempsInstallationProj.Text) Then
4  FrmProjSoumMec.m_sTempsInstallationProj = txtTempsInstallationProj.Text
4  Else
4  FrmProjSoumMec.m_sTempsInstallationProj = "0"
50 End If

If Trim$(txtTempsFormationProj.Text) <> vbNullString And IsNumeric(txtTempsFormationProj.Text) Then
 FrmProjSoumMec.m_sTempsFormationProj = txtTempsFormationProj.Text
 Else
 FrmProjSoumMec.m_sTempsFormationProj = "0"
 End If

 If Trim$(txtTempsGestionProj.Text) <> vbNullString And IsNumeric(txtTempsGestionProj.Text) Then
 FrmProjSoumMec.m_sTempsGestionProj = txtTempsGestionProj.Text
 Else
 FrmProjSoumMec.m_sTempsGestionProj = "0"
 End If

 If Trim$(txtTempsShippingProj.Text) <> vbNullString And IsNumeric(txtTempsShippingProj.Text) Then
5  FrmProjSoumMec.m_sTempsShippingProj = txtTempsShippingProj.Text
5  Else
5  FrmProjSoumMec.m_sTempsShippingProj = "0"
5  End If
5  Else
5  If Trim$(txtTempsDessinConc.Text) <> vbNullString And IsNumeric(txtTempsDessinConc.Text) Then
5  FrmProjSoumMec.m_sTempsDessinConc = txtTempsDessinConc.Text
5  Else
60 FrmProjSoumMec.m_sTempsDessinConc = "0"
  End If

  If Trim$(txtTempsCoupeConc.Text) <> vbNullString And IsNumeric(txtTempsCoupeConc.Text) Then
  FrmProjSoumMec.m_sTempsCoupeConc = txtTempsCoupeConc.Text
  Else
  FrmProjSoumMec.m_sTempsCoupeConc = "0"
  End If

  If Trim$(txtTempsMachinageConc.Text) <> vbNullString And IsNumeric(txtTempsMachinageConc.Text) Then
  FrmProjSoumMec.m_sTempsMachinageConc = txtTempsMachinageConc.Text
  Else
  FrmProjSoumMec.m_sTempsMachinageConc = "0"
  End If

6  If Trim$(txtTempsSoudureConc.Text) <> vbNullString And IsNumeric(txtTempsSoudureConc.Text) Then
6  FrmProjSoumMec.m_sTempsSoudureConc = txtTempsSoudureConc.Text
6  Else
6  FrmProjSoumMec.m_sTempsSoudureConc = "0"
6  End If

6  If Trim$(txtTempsAssemblageConc.Text) <> vbNullString And IsNumeric(txtTempsAssemblageConc.Text) Then
6  FrmProjSoumMec.m_sTempsAssemblageConc = txtTempsAssemblageConc.Text
6  Else
70 FrmProjSoumMec.m_sTempsAssemblageConc = "0"
  End If

  If Trim$(txtTempsPeintureConc.Text) <> vbNullString And IsNumeric(txtTempsPeintureConc.Text) Then
  FrmProjSoumMec.m_sTempsPeintureConc = txtTempsPeintureConc.Text
  Else
  FrmProjSoumMec.m_sTempsPeintureConc = "0"
  End If

  If Trim$(txtTempsTestConc.Text) <> vbNullString And IsNumeric(txtTempsTestConc.Text) Then
  FrmProjSoumMec.m_sTempsTestConc = txtTempsTestConc.Text
  Else
  FrmProjSoumMec.m_sTempsTestConc = "0"
  End If

   If Trim$(txtTempsInstallationConc.Text) <> vbNullString And IsNumeric(txtTempsInstallationConc.Text) Then
   FrmProjSoumMec.m_sTempsInstallationConc = txtTempsInstallationConc.Text
7  Else
7  FrmProjSoumMec.m_sTempsInstallationConc = "0"
7  End If

7  If Trim$(txtTempsFormationConc.Text) <> vbNullString And IsNumeric(txtTempsFormationConc.Text) Then
7  FrmProjSoumMec.m_sTempsFormationConc = txtTempsFormationConc.Text
7  Else
80 FrmProjSoumMec.m_sTempsFormationConc = "0"
  End If

  If Trim$(txtTempsGestionConc.Text) <> vbNullString And IsNumeric(txtTempsGestionConc.Text) Then
  FrmProjSoumMec.m_sTempsGestionConc = txtTempsGestionConc.Text
  Else
  FrmProjSoumMec.m_sTempsGestionConc = "0"
  End If

  If Trim$(txtTempsShippingConc.Text) <> vbNullString And IsNumeric(txtTempsShippingConc.Text) Then
  FrmProjSoumMec.m_sTempsShippingConc = txtTempsShippingConc.Text
  Else
  FrmProjSoumMec.m_sTempsShippingConc = "0"
  End If
   End If
   End If
 
   If Trim$(txtNbrePersonne.Text) <> vbNullString And IsNumeric(txtNbrePersonne.Text) Then
   FrmProjSoumMec.m_sNbrePersonne = txtNbrePersonne.Text
8  Else
8  FrmProjSoumMec.m_sNbrePersonne = "0"
8  End If
 
8  If Trim$(txtTempsHebergement.Text) <> vbNullString And IsNumeric(txtTempsHebergement.Text) Then
90 FrmProjSoumMec.m_sTempsHebergement = txtTempsHebergement.Text
90 Else
  FrmProjSoumMec.m_sTempsHebergement = "0"
  End If
 
  If Trim$(txtTempsRepas.Text) <> vbNullString And IsNumeric(txtTempsRepas.Text) Then
  FrmProjSoumMec.m_sTempsRepas = txtTempsRepas.Text
  Else
  FrmProjSoumMec.m_sTempsRepas = "0"
  End If
 
  If Trim$(txtTempsDeplacement.Text) <> vbNullString And IsNumeric(txtTempsDeplacement.Text) Then
  FrmProjSoumMec.m_sTempsTransport = txtTempsDeplacement.Text
  Else
 FrmProjSoumMec.m_sTempsTransport = "0"
   End If
 
 If Trim$(txtTempsUniteMobile.Text) <> vbNullString And IsNumeric(txtTempsUniteMobile.Text) Then
   FrmProjSoumMec.m_sTempsUniteMobile = txtTempsUniteMobile.Text
 Else
   FrmProjSoumMec.m_sTempsUniteMobile = "0"
 End If
 
9  If Trim$(txtPrixEmballage.Text) <> vbNullString And IsNumeric(txtPrixEmballage.Text) Then
 FrmProjSoumMec.m_sPrixEmballage = txtPrixEmballage.Text
100 Else
1 FrmProjSoumMec.m_sPrixEmballage = "0"
10 End If
 
FrmProjSoumMec.m_sTauxHebergement1 = m_sHebergement1
10 FrmProjSoumMec.m_sTauxHebergement2 = m_sHebergement2
FrmProjSoumMec.m_sTauxRepas = m_sRepas
10 FrmProjSoumMec.m_sTauxTransport = m_sStandard
FrmProjSoumMec.m_sTauxUniteMobile = m_sUniteMobile

10 Exit Sub

Oups:

wOups "frmProjSoumMecTemps", "EnregistrerTemps", Err, Err.number, Err.Description
End Sub

Private Sub cmdLock_Click()
 
 On Error GoTo Oups

 If m_sTempsDessinAvant <> txtTempsDessinProj.Text Or _
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
 Call MsgBox("Veuillez enregistrer le projet en premier sinon vous allez perdre les informations qui ont été modifiées dans le temps projets!", vbOKOnly, "Erreur")
 Else
 m_bLocked = True
 
 Call BarrerChamps(False)
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "cmdLock_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdUnlock_Click()
 
 On Error GoTo Oups

 m_bLocked = False
 
 Call BarrerChamps(False)

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "cmdUnlock_Click", Err, Err.number, Err.Description
End Sub

Private Sub InitialiserVariablesConfig()

 On Error GoTo Oups

 'Initialise les variables à partir de la table Config (Pour avoir le taux
 'horaire le plus récent)
 Dim rstConfig As ADODB.Recordset
 
 Set rstConfig = New ADODB.Recordset
 
 Call rstConfig.Open("SELECT TauxDessinMec, TauxCoupe, TauxMachinage, TauxSoudure, TauxAssemblageMec, TauxPeinture, TauxTestMec, TauxInstallationMec, TauxFormationMec, TauxGestionProjetsMec, TauxShippingMec, Repas, Hebergement1, Hebergement2, Standard, UniteMobile FROM GrbConfig", g_connData, adOpenDynamic, adLockOptimistic)
 
 If Not IsNull(rstConfig.Fields("TauxDessinMec")) Then
 m_sTauxDessin = rstConfig.Fields("TauxDessinMec")
 Else
 m_sTauxDessin = "0"
 End If

 If Not IsNull(rstConfig.Fields("TauxCoupe")) Then
 m_sTauxCoupe = rstConfig.Fields("TauxCoupe")
  Else
  m_sTauxCoupe = "0"
  End If

  If Not IsNull(rstConfig.Fields("TauxMachinage")) Then
  m_sTauxMachinage = rstConfig.Fields("TauxMachinage")
  Else
  m_sTauxMachinage = "0"
  End If

10 If Not IsNull(rstConfig.Fields("TauxSoudure")) Then
1 m_sTauxSoudure = rstConfig.Fields("TauxSoudure")
Else
 m_sTauxSoudure = "0"
End If

If Not IsNull(rstConfig.Fields("TauxAssemblageMec")) Then
 m_sTauxAssemblage = rstConfig.Fields("TauxAssemblageMec")
Else
 m_sTauxAssemblage = "0"
End If

If Not IsNull(rstConfig.Fields("TauxPeinture")) Then
 m_sTauxPeinture = rstConfig.Fields("TauxPeinture")
1  Else
 m_sTauxPeinture = "0"
 End If

If Not IsNull(rstConfig.Fields("TauxTestMec")) Then
 m_sTauxTest = rstConfig.Fields("TauxTestMec")
Else
 m_sTauxTest = "0"
1  End If

 If Not IsNull(rstConfig.Fields("TauxInstallationMec")) Then
 m_sTauxInstallation = rstConfig.Fields("TauxInstallationMec")
Else
 m_sTauxInstallation = "0"
End If

If Not IsNull(rstConfig.Fields("TauxFormationMec")) Then
 m_sTauxFormation = rstConfig.Fields("TauxFormationMec")
Else
 m_sTauxFormation = "0"
End If

If Not IsNull(rstConfig.Fields("TauxGestionProjetsMec")) Then
 m_sTauxGestion = rstConfig.Fields("TauxGestionProjetsMec")
2  Else
 m_sTauxGestion = "0"
2  End If

If Not IsNull(rstConfig.Fields("TauxShippingMec")) Then
m_sTauxShipping = rstConfig.Fields("TauxShippingMec")
Else
m_sTauxShipping = "0"
End If

 If Not IsNull(rstConfig.Fields("TauxShippingMec")) Then
2   m_sTauxPrototype = rstConfig.Fields("TauxShippingMec")
2   Else
2   m_sTauxPrototype = "0"
29  End If




30 m_sRepas = rstConfig.Fields("Repas")
m_sHebergement1 = rstConfig.Fields("Hebergement1")
m_sHebergement2 = rstConfig.Fields("Hebergement2")
m_sStandard = rstConfig.Fields("Standard")
m_sUniteMobile = rstConfig.Fields("UniteMobile")
 
Call rstConfig.Close
Set rstConfig = Nothing

Exit Sub

Oups:

wOups "frmProjSoumMecTemps", "InitialiserVariablesConfig", Err, Err.number, Err.Description
End Sub

Private Sub InitialiserVariablesProjSoum()

 On Error GoTo Oups

 m_sTauxDessin = FrmProjSoumMec.m_sTauxDessin
 m_sTauxCoupe = FrmProjSoumMec.m_sTauxCoupe
 m_sTauxMachinage = FrmProjSoumMec.m_sTauxMachinage
 m_sTauxSoudure = FrmProjSoumMec.m_sTauxSoudure
 m_sTauxAssemblage = FrmProjSoumMec.m_sTauxAssemblage
 m_sTauxPeinture = FrmProjSoumMec.m_sTauxPeinture
 m_sTauxTest = FrmProjSoumMec.m_sTauxTest
 m_sTauxInstallation = FrmProjSoumMec.m_sTauxInstallation
 m_sTauxFormation = FrmProjSoumMec.m_sTauxFormation
 m_sTauxGestion = FrmProjSoumMec.m_sTauxGestion
  m_sTauxShipping = FrmProjSoumMec.m_sTauxShipping
 m_sTauxPrototype = FrmProjSoumMec.m_sTauxShipping

  m_sRepas = FrmProjSoumMec.m_sTauxRepas
  m_sHebergement1 = FrmProjSoumMec.m_sTauxHebergement1
  m_sHebergement2 = FrmProjSoumMec.m_sTauxHebergement2
  m_sStandard = FrmProjSoumMec.m_sTauxTransport
  m_sUniteMobile = FrmProjSoumMec.m_sTauxUniteMobile

  Exit Sub

Oups:

  wOups "frmProjSoumMecTemps", "InitialiserVariablesProjSoum", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()
 
 On Error GoTo Oups
 
 If FrmProjSoumMec.m_bDroitPrix = False Then
 fraRessourcesHumaines.width = 4150
 fraFraisSubsistences.width = 4150
 
 fraFraisSubsistences.Left = 4390
 
 fraManutention.Visible = False
 lblTotalPrixRH.Visible = False
 
 Cmdfermer.Left = 7320
 
 Cmdfermer.Top = 4200
 
 Me.width = 8800
 Me.Height = 7485
  End If

  Exit Sub

Oups:

  wOups "frmProjSoumMecTemps", "Form_Load", Err, Err.number, Err.Description
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

 wOups "frmProjSoumMecTemps", "txtNbrePersonne_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtPrixEmballage_KeyPress(KeyAscii As Integer)

 On Error GoTo Oups

 If KeyAscii = 4 Then  'Si c'est le "."
 KeyAscii = 44 'Remplace par la ","
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "lblPrixEmballage_KeyPress", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsAssemblageConc_Change()
 
 On Error GoTo Oups

 Call CalculerTotalTemps

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsAssemblageConc_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsAssemblageProj_Change()
 
 On Error GoTo Oups

 Call CalculerTotalTemps

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsAssemblageProj_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsCoupeConc_Change()
 
 On Error GoTo Oups

 Call CalculerTotalTemps

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsCoupeConc_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsCoupeProj_Change()
 
 On Error GoTo Oups

 Call CalculerTotalTemps

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsCoupeProj_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsDessinConc_Change()
 
 On Error GoTo Oups

 Call CalculerTotalTemps

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsDessinConc_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsDessinProj_Change()

 On Error GoTo Oups

 Call CalculerTotalTemps

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsDessinProj_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsFormationConc_Change()
 
 On Error GoTo Oups

 Call CalculerTotalTemps

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsFormationConc_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsFormationProj_Change()
 
 On Error GoTo Oups

 Call CalculerTotalTemps

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsFormationProj_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsGestionConc_Change()
 
 On Error GoTo Oups

 Call CalculerTotalTemps

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsGestionConc_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsGestionProj_Change()
 
 On Error GoTo Oups

 Call CalculerTotalTemps

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsGestionProj_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsShippingConc_Change()
 
 On Error GoTo Oups

 Call CalculerTotalTemps

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsShippingConc_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsShippingProj_Change()
 
 On Error GoTo Oups

 Call CalculerTotalTemps

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsShippingProj_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsInstallationConc_Change()
 
 On Error GoTo Oups

 Call CalculerTotalTemps

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsInstallationConc_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsInstallationProj_Change()
 
 On Error GoTo Oups

 Call CalculerTotalTemps

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsInstallationProj_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsMachinageConc_Change()
 
 On Error GoTo Oups

 Call CalculerTotalTemps

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsMachinageConc_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsMachinageProj_Change()
 
 On Error GoTo Oups

 Call CalculerTotalTemps

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsMachinageProj_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsMachinageSoum_Change()

 On Error GoTo Oups

 If m_eType = TYPE_SOUMISSION Then
 If IsNumeric(txtTempsMachinageSoum.Text) Then
 lblPrixMachinage.Caption = Round(Replace(txtTempsMachinageSoum.Text * m_sTauxMachinage, ".", ","), 2)
 Else
 lblPrixMachinage.Caption = 0
 End If
 
 Call CalculerTotal
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsMachinageSoum_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsMachinageSoum_LostFocus()

 On Error GoTo Oups

 txtTempsMachinageSoum.Text = Replace(txtTempsMachinageSoum.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsMachinageSoum_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsMachinageProj_LostFocus()

 On Error GoTo Oups

 txtTempsMachinageProj.Text = Replace(txtTempsMachinageProj.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsMachinageProj_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsMachinageConc_LostFocus()

 On Error GoTo Oups

 txtTempsMachinageConc.Text = Replace(txtTempsMachinageConc.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsMachinageConc_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsCoupeSoum_Change()

 On Error GoTo Oups

 If m_eType = TYPE_SOUMISSION Then
 If IsNumeric(txtTempsCoupeSoum.Text) Then
 lblPrixCoupe.Caption = Round(Replace(txtTempsCoupeSoum.Text * m_sTauxCoupe, ".", ","), 2)
 Else
 lblPrixCoupe.Caption = 0
 End If

 Call CalculerTotal
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsCoupeSoum_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsCoupeSoum_LostFocus()

 On Error GoTo Oups

 txtTempsCoupeSoum.Text = Replace(txtTempsCoupeSoum.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsCoupeSoum_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsCoupeProj_LostFocus()

 On Error GoTo Oups

 txtTempsCoupeProj.Text = Replace(txtTempsCoupeProj.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsCoupeProj_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsCoupeConc_LostFocus()

 On Error GoTo Oups

 txtTempsCoupeConc.Text = Replace(txtTempsCoupeConc.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsCoupeConc_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsPeintureConc_Change()
 
 On Error GoTo Oups

 Call CalculerTotalTemps

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsPeintureConc_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsPeintureProj_Change()
 
 On Error GoTo Oups

 Call CalculerTotalTemps

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsPeintureProj_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsSoudureConc_Change()
 
 On Error GoTo Oups

 Call CalculerTotalTemps

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsSoudureConc_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsSoudureProj_Change()
 
 On Error GoTo Oups

 Call CalculerTotalTemps

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsSoudureProj_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsSoudureSoum_Change()

 On Error GoTo Oups

 If m_eType = TYPE_SOUMISSION Then
 If IsNumeric(txtTempsSoudureSoum.Text) Then
 lblPrixSoudure.Caption = Round(Replace(txtTempsSoudureSoum.Text * m_sTauxSoudure, ".", ","), 2)
 Else
 lblPrixSoudure.Caption = 0
 End If
 
 Call CalculerTotal
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsSoudureSoum_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsSoudureSoum_LostFocus()

 On Error GoTo Oups

 txtTempsSoudureSoum.Text = Replace(txtTempsSoudureSoum.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsSoudureSoum_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsSoudureProj_LostFocus()

 On Error GoTo Oups

 txtTempsSoudureProj.Text = Replace(txtTempsSoudureProj.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsSoudureProj_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsSoudureConc_LostFocus()

 On Error GoTo Oups

 txtTempsSoudureConc.Text = Replace(txtTempsSoudureConc.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsSoudureConc_LostFocus", Err, Err.number, Err.Description
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

 wOups "frmProjSoumMecTemps", "txtTempsAssemblageSoum_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsAssemblageSoum_LostFocus()

 On Error GoTo Oups

 txtTempsAssemblageSoum.Text = Replace(txtTempsAssemblageSoum.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsAssemblageSoum_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsAssemblageProj_LostFocus()

 On Error GoTo Oups

 txtTempsAssemblageProj.Text = Replace(txtTempsAssemblageProj.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsAssemblageProj_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsAssemblageConc_LostFocus()

 On Error GoTo Oups

 txtTempsAssemblageConc.Text = Replace(txtTempsAssemblageConc.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsAssemblageConc_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsPeintureSoum_Change()

 On Error GoTo Oups

 If m_eType = TYPE_SOUMISSION Then
 If IsNumeric(txtTempsPeintureSoum.Text) Then
 lblPrixPeinture.Caption = Round(Replace(txtTempsPeintureSoum.Text * m_sTauxPeinture, ".", ","), 2)
 Else
 lblPrixPeinture.Caption = 0
 End If
 
 Call CalculerTotal
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsPeintureSoum_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsPeintureSoum_LostFocus()

 On Error GoTo Oups

 txtTempsPeintureSoum.Text = Replace(txtTempsPeintureSoum.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsPeintureSoum_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsPeintureProj_LostFocus()

 On Error GoTo Oups

 txtTempsPeintureProj.Text = Replace(txtTempsPeintureProj.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsPeintureProj_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsPeintureConc_LostFocus()

 On Error GoTo Oups

 txtTempsPeintureConc.Text = Replace(txtTempsPeintureConc.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsPeintureConc_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsTestConc_Change()
 
 On Error GoTo Oups

 Call CalculerTotalTemps

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsTestConc_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsTestProj_Change()
 
 On Error GoTo Oups

 Call CalculerTotalTemps

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsTestProj_Change", Err, Err.number, Err.Description
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

 wOups "frmProjSoumMecTemps", "txtTempsTestSoum_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsTestSoum_LostFocus()

 On Error GoTo Oups

 txtTempsTestSoum.Text = Replace(txtTempsTestSoum.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsTestSoum_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsTestProj_LostFocus()

 On Error GoTo Oups

 txtTempsTestProj.Text = Replace(txtTempsTestProj.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsTestProj_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsTestConc_LostFocus()

 On Error GoTo Oups

 txtTempsTestConc.Text = Replace(txtTempsTestConc.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsTestConc_LostFocus", Err, Err.number, Err.Description
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

 wOups "frmProjSoumMecTemps", "txtTempsInstallationSoum_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsInstallationSoum_LostFocus()

 On Error GoTo Oups

 txtTempsInstallationSoum.Text = Replace(txtTempsInstallationSoum.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsInstallationSoum_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsInstallationProj_LostFocus()

 On Error GoTo Oups

 txtTempsInstallationProj.Text = Replace(txtTempsInstallationProj.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsInstallationProj_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsInstallationConc_LostFocus()

 On Error GoTo Oups

 txtTempsInstallationConc.Text = Replace(txtTempsInstallationConc.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsInstallationConc_LostFocus", Err, Err.number, Err.Description
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

 wOups "frmProjSoumMecTemps", "txtTempsDessinSoum_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsDessinSoum_LostFocus()

 On Error GoTo Oups

 txtTempsDessinSoum.Text = Replace(txtTempsDessinSoum.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsDessinSoum_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsDessinProj_LostFocus()

 On Error GoTo Oups

 txtTempsDessinProj.Text = Replace(txtTempsDessinProj.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsDessinProj_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsDessinConc_LostFocus()

 On Error GoTo Oups

 txtTempsDessinConc.Text = Replace(txtTempsDessinConc.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsDessinConc_LostFocus", Err, Err.number, Err.Description
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

 wOups "frmProjSoumMecTemps", "txtTempsFormationSoum_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsFormationSoum_LostFocus()

 On Error GoTo Oups

 txtTempsFormationSoum.Text = Replace(txtTempsFormationSoum.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsFormationSoum_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsFormationProj_LostFocus()

 On Error GoTo Oups

 txtTempsFormationProj.Text = Replace(txtTempsFormationProj.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsFormationProj_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsFormationConc_LostFocus()

 On Error GoTo Oups

 txtTempsFormationConc.Text = Replace(txtTempsFormationConc.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsFormationConc_LostFocus", Err, Err.number, Err.Description
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

 wOups "frmProjSoumMecTemps", "txtTempsGestionSoum_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsGestionSoum_LostFocus()

 On Error GoTo Oups
 
 txtTempsGestionSoum.Text = Replace(txtTempsGestionSoum.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsGestionSoum_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsGestionProj_LostFocus()

 On Error GoTo Oups
 
 txtTempsGestionProj.Text = Replace(txtTempsGestionProj.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsGestionProj_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsGestionConc_LostFocus()

 On Error GoTo Oups
 
 txtTempsGestionConc.Text = Replace(txtTempsGestionConc.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsGestionConc_LostFocus", Err, Err.number, Err.Description
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

 wOups "frmProjSoumMecTemps", "txtTempsShippingSoum_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsShippingSoum_LostFocus()

 On Error GoTo Oups
 
 txtTempsShippingSoum.Text = Replace(txtTempsShippingSoum.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsShippingSoum_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsShippingProj_LostFocus()

 On Error GoTo Oups
 
 txtTempsShippingProj.Text = Replace(txtTempsShippingProj.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsShippingProj_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsShippingConc_LostFocus()

 On Error GoTo Oups
 
 txtTempsShippingConc.Text = Replace(txtTempsShippingConc.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsShippingConc_LostFocus", Err, Err.number, Err.Description
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

 wOups "frmProjSoumMecTemps", "txtTempsHebergement_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsHebergement_LostFocus()

 On Error GoTo Oups

 txtTempsHebergement.Text = Replace(txtTempsHebergement.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsHebergement_LostFocus", Err, Err.number, Err.Description
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

 wOups "frmProjSoumMecTemps", "txtTempsRepas_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsRepas_LostFocus()

 On Error GoTo Oups

 txtTempsRepas.Text = Replace(txtTempsRepas.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsRepas_LostFocus", Err, Err.number, Err.Description
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

 wOups "frmProjSoumMecTemps", "txtTempsDeplacement_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsDeplacement_LostFocus()

 On Error GoTo Oups

 txtTempsDeplacement.Text = Replace(txtTempsDeplacement.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsDeplacement_LostFocus", Err, Err.number, Err.Description
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

 wOups "frmProjSoumMecTemps", "txtTempsUniteMobile_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtTempsUniteMobile_LostFocus()

 On Error GoTo Oups

 txtTempsUniteMobile.Text = Replace(txtTempsUniteMobile.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "txtTempsUniteMobile_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtPrixEmballage_Change()

 On Error GoTo Oups
 
 Call CalculerTotal

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "lblPrixEmballage_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtPrixEmballage_LostFocus()

 On Error GoTo Oups

 If IsNumeric(txtPrixEmballage.Text) Then
 txtPrixEmballage.Text = Round(Replace(txtPrixEmballage.Text, ".", ","), 2)
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "lblPrixEmballage_LostFocus", Err, Err.number, Err.Description
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

wOups "frmProjSoumMecTemps", "CalculerHebergement", Err, Err.number, Err.Description
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

  wOups "frmProjSoumMecTemps", "CalculerRepas", Err, Err.number, Err.Description
End Sub

Private Sub CalculerTotal()

 On Error GoTo Oups

 Dim dblTotal As Double
 Dim dblPrixEmballage As Double
 Dim dblTotalArgentRH As Double
 Dim dblPrixDessin As Double
 Dim dblPrixCoupe As Double
 Dim dblPrixMachinage As Double
 Dim dblPrixSoudure As Double
 Dim dblPrixAssemblage As Double
 Dim dblPrixPeinture As Double
 Dim dblPrixTest As Double
  Dim dblPrixInstallation As Double
  Dim dblPrixFormation As Double
  Dim dblPrixGestion As Double
  Dim dblPrixShipping As Double
  Dim dblPrixHebergement As Double
  Dim dblPrixRepas As Double
  Dim dblPrixDeplacement As Double
  Dim dblPrixUniteMobile As Double
 
 'Prix de dessin
10 If IsNumeric(lblPrixDessin.Caption) Then
1 dblPrixDessin = CDbl(lblPrixDessin.Caption)
Else
 dblPrixDessin = 0
End If
 
 'Prix de coupe et préparation
If IsNumeric(lblPrixCoupe.Caption) Then
 dblPrixCoupe = CDbl(lblPrixCoupe.Caption)
Else
 dblPrixCoupe = 0
End If
 
 'Prix de machinage
If IsNumeric(lblPrixMachinage.Caption) Then
 dblPrixMachinage = CDbl(lblPrixMachinage.Caption)
1  Else
 dblPrixMachinage = 0
 End If
 
 'Prix de soudure et meulage
If IsNumeric(lblPrixSoudure.Caption) Then
 dblPrixSoudure = CDbl(lblPrixSoudure.Caption)
Else
 dblPrixSoudure = 0
1  End If

 'Prix d'assemblage du système
 If IsNumeric(lblPrixAssemblage.Caption) Then
 dblPrixAssemblage = CDbl(lblPrixAssemblage.Caption)
Else
 dblPrixAssemblage = 0
End If

 'Prix de peinture et finition
If IsNumeric(lblPrixPeinture.Caption) Then
 dblPrixPeinture = CDbl(lblPrixPeinture.Caption)
Else
 dblPrixPeinture = 0
End If

 'Prix de tests finaux
If IsNumeric(lblPrixTest.Caption) Then
 dblPrixTest = CDbl(lblPrixTest.Caption)
2  Else
 dblPrixTest = 0
2  End If

 'Prix d'Installation
If IsNumeric(lblPrixInstallation.Caption) Then
dblPrixInstallation = CDbl(lblPrixInstallation.Caption)
Else
dblPrixInstallation = 0
End If

 'Prix de formation
30 If IsNumeric(lblPrixFormation.Caption) Then
3 dblPrixFormation = CDbl(lblPrixFormation.Caption)
Else
 dblPrixFormation = 0
End If

 'Prix de gestion du projet
If IsNumeric(lblPrixGestion.Caption) Then
 dblPrixGestion = CDbl(lblPrixGestion.Caption)
Else
 dblPrixGestion = 0
End If

 'Prix de shipping
If IsNumeric(lblPrixShipping.Caption) Then
 dblPrixShipping = CDbl(lblPrixShipping.Caption)
3  Else
 dblPrixShipping = 0
3  End If


 'Prix de dévelloppement prototypage
3  If IsNumeric(lblPrixPrototype.Caption) Then
 'dblPrixPrototype = CDbl(lblPrixPrototype.Caption)
3  Else
374 'dblPrixPrototype = 0
 End If



 'Prix d'hébergement
If IsNumeric(lblPrixHebergement.Caption) Then
dblPrixHebergement = CDbl(lblPrixHebergement.Caption)
Else
 dblPrixHebergement = 0
 End If

 'Prix des repas
40 If IsNumeric(lblPrixRepas.Caption) Then
4 dblPrixRepas = CDbl(lblPrixRepas.Caption)
4 Else
4 dblPrixRepas = 0
4 End If
 
 'Prix du déplacement
4 If IsNumeric(lblPrixDeplacement.Caption) Then
4 dblPrixDeplacement = CDbl(lblPrixDeplacement.Caption)
4 Else
4 dblPrixDeplacement = 0
4 End If

 'Prix de l'unité mobile
4 If IsNumeric(lblPrixUniteMobile.Caption) Then
4 dblPrixUniteMobile = CDbl(lblPrixUniteMobile.Caption)
4  Else
4  dblPrixUniteMobile = 0
4  End If
 
 'Prix de transport et emballage
4  If IsNumeric(txtPrixEmballage.Text) Then
4  dblPrixEmballage = CDbl(txtPrixEmballage.Text)
4  Else
4  dblPrixEmballage = 0
4  End If
 
50 dblTotalArgentRH = dblPrixDessin + _
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

50 lblTotalPrixRH.Caption = Conversion(CStr(dblTotalArgentRH), MODE_DECIMAL)

 
 dblTotal = dblTotalArgentRH + _
 dblPrixHebergement + _
 dblPrixRepas + _
 dblPrixDeplacement + _
 dblPrixUniteMobile + _
 dblPrixEmballage
 
 lblTotal.Caption = Conversion(CStr(dblTotal), MODE_DECIMAL)

 Call CalculerTotalTemps

 Exit Sub

Oups:

 wOups "frmProjSoumMecTemps", "CalculerTotal", Err, Err.number, Err.Description
End Sub

Private Sub CalculerTotalTemps()

 
 On Error GoTo Oups

 Dim dblTempsDessin As Double
 Dim dblTempsCoupe As Double
 Dim dblTempsMachinage As Double
 Dim dblTempsSoudure As Double
 Dim dblTempsAssemblage As Double
 Dim dblTempsPeinture As Double
 Dim dblTempsTest As Double
 Dim dblTempsInstallation As Double
 Dim dblTempsFormation As Double
 Dim dblTempsGestion As Double
  Dim dblTempsShipping As Double
  Dim dblTotalTemps As Double

 'SOUMISSION
  If IsNumeric(txtTempsDessinSoum.Text) Then
  dblTempsDessin = CDbl(txtTempsDessinSoum.Text)
  Else
  dblTempsDessin = 0
  End If
 
  If IsNumeric(txtTempsCoupeSoum.Text) Then
dblTempsCoupe = CDbl(txtTempsCoupeSoum.Text)
Else
 dblTempsCoupe = 0
End If

If IsNumeric(txtTempsMachinageSoum.Text) Then
 dblTempsMachinage = CDbl(txtTempsMachinageSoum.Text)
Else
 dblTempsMachinage = 0
End If

If IsNumeric(txtTempsSoudureSoum.Text) Then
 dblTempsSoudure = CDbl(txtTempsSoudureSoum.Text)
Else
dblTempsSoudure = 0
End If

 If IsNumeric(txtTempsAssemblageSoum.Text) Then
 dblTempsAssemblage = CDbl(txtTempsAssemblageSoum.Text)
 Else
 dblTempsAssemblage = 0
 End If

1  If IsNumeric(txtTempsPeintureSoum.Text) Then
 dblTempsPeinture = CDbl(txtTempsPeintureSoum.Text)
 Else
 dblTempsPeinture = 0
End If

If IsNumeric(txtTempsTestSoum.Text) Then
 dblTempsTest = CDbl(txtTempsTestSoum.Text)
Else
 dblTempsTest = 0
End If

If IsNumeric(txtTempsInstallationSoum.Text) Then
 dblTempsInstallation = CDbl(txtTempsInstallationSoum.Text)
Else
dblTempsInstallation = 0
End If

2  If IsNumeric(txtTempsFormationSoum.Text) Then
 dblTempsFormation = CDbl(txtTempsFormationSoum.Text)
2  Else
 dblTempsFormation = 0
2  End If

If IsNumeric(txtTempsGestionSoum.Text) Then
dblTempsGestion = CDbl(txtTempsGestionSoum.Text)
Else
 dblTempsGestion = 0
End If

If IsNumeric(txtTempsShippingSoum.Text) Then
 dblTempsShipping = CDbl(txtTempsShippingSoum.Text)
Else
 dblTempsShipping = 0
End If

dblTotalTemps = dblTempsDessin + _
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

lblTotalTempsRHSoum.Caption = Conversion(dblTotalTemps, MODE_DECIMAL)

 'PROJET
If m_eType = TYPE_PROJET Then
If IsNumeric(txtTempsDessinProj.Text) Then
 dblTempsDessin = CDbl(txtTempsDessinProj.Text)
Else
 dblTempsDessin = 0
End If

 If IsNumeric(txtTempsCoupeProj.Text) Then
 dblTempsCoupe = CDbl(txtTempsCoupeProj.Text)
 Else
 dblTempsCoupe = 0
4 End If

4 If IsNumeric(txtTempsMachinageProj.Text) Then
4 dblTempsMachinage = CDbl(txtTempsMachinageProj.Text)
4 Else
4 dblTempsMachinage = 0
4 End If

4 If IsNumeric(txtTempsSoudureProj.Text) Then
4 dblTempsSoudure = CDbl(txtTempsSoudureProj.Text)
4 Else
4 dblTempsSoudure = 0
4 End If

4  If IsNumeric(txtTempsAssemblageProj.Text) Then
4  dblTempsAssemblage = CDbl(txtTempsAssemblageProj.Text)
4  Else
4  dblTempsAssemblage = 0
4  End If

4  If IsNumeric(txtTempsPeintureProj.Text) Then
4  dblTempsPeinture = CDbl(txtTempsPeintureProj.Text)
4  Else
50 dblTempsPeinture = 0
5 End If

 If IsNumeric(txtTempsTestProj.Text) Then
 dblTempsTest = CDbl(txtTempsTestProj.Text)
 Else
 dblTempsTest = 0
 End If

 If IsNumeric(txtTempsInstallationProj.Text) Then
 dblTempsInstallation = CDbl(txtTempsInstallationProj.Text)
 Else
 dblTempsInstallation = 0
 End If

5  If IsNumeric(txtTempsFormationProj.Text) Then
5  dblTempsFormation = CDbl(txtTempsFormationProj.Text)
5  Else
5  dblTempsFormation = 0
5  End If

5  If IsNumeric(txtTempsGestionProj.Text) Then
5  dblTempsGestion = CDbl(txtTempsGestionProj.Text)
5  Else
60 dblTempsGestion = 0
  End If

  If IsNumeric(txtTempsShippingProj.Text) Then
  dblTempsShipping = CDbl(txtTempsShippingProj.Text)
  Else
  dblTempsShipping = 0
  End If


  If IsNumeric(txtTempsPrototypeProj.Text) Then
  ' dblTempsPrototype = CDbl(txtTempsPrototypeProj.Text)
63 Else
634 ' dblTempsPrototype = 0
 End If



  dblTotalTemps = dblTempsDessin + _
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

  lblTotalTempsRHProj.Caption = Conversion(dblTotalTemps, MODE_DECIMAL)
  End If

 'CONCEPTION
  If m_eType = TYPE_PROJET And m_bLocked = True Then
  If IsNumeric(txtTempsDessinConc.Text) Then
6  dblTempsDessin = CDbl(txtTempsDessinConc.Text)
6  Else
6  dblTempsDessin = 0
6  End If

6  If IsNumeric(txtTempsCoupeConc.Text) Then
6  dblTempsCoupe = CDbl(txtTempsCoupeConc.Text)
6  Else
6  dblTempsCoupe = 0
70 End If

  If IsNumeric(txtTempsMachinageConc.Text) Then
  dblTempsMachinage = CDbl(txtTempsMachinageConc.Text)
  Else
  dblTempsMachinage = 0
  End If

  If IsNumeric(txtTempsSoudureConc.Text) Then
  dblTempsSoudure = CDbl(txtTempsSoudureConc.Text)
  Else
  dblTempsSoudure = 0
  End If

  If IsNumeric(txtTempsAssemblageConc.Text) Then
   dblTempsAssemblage = CDbl(txtTempsAssemblageConc.Text)
   Else
7  dblTempsAssemblage = 0
7  End If

7  If IsNumeric(txtTempsPeintureConc.Text) Then
7  dblTempsPeinture = CDbl(txtTempsPeintureConc.Text)
7  Else
7  dblTempsPeinture = 0
80 End If

  If IsNumeric(txtTempsTestConc.Text) Then
  dblTempsTest = CDbl(txtTempsTestConc.Text)
  Else
  dblTempsTest = 0
  End If

  If IsNumeric(txtTempsInstallationConc.Text) Then
  dblTempsInstallation = CDbl(txtTempsInstallationConc.Text)
  Else
  dblTempsInstallation = 0
  End If

  If IsNumeric(txtTempsFormationConc.Text) Then
   dblTempsFormation = CDbl(txtTempsFormationConc.Text)
   Else
   dblTempsFormation = 0
   End If

8  If IsNumeric(txtTempsGestionConc.Text) Then
8  dblTempsGestion = CDbl(txtTempsGestionConc.Text)
8  Else
8  dblTempsGestion = 0
90 End If

  If IsNumeric(txtTempsShippingConc.Text) Then
  dblTempsShipping = CDbl(txtTempsShippingConc.Text)
  Else
  dblTempsShipping = 0
  End If

  dblTotalTemps = dblTempsDessin + _
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
 
  lblTotalTempsRHConc.Caption = Conversion(dblTotalTemps, MODE_DECIMAL)
  End If

  Exit Sub

Oups:

  wOups "frmProjSoumMecTemps", "CalculerTotalTemps", Err, Err.number, Err.Description
End Sub
