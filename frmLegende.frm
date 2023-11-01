VERSION 5.00
Begin VB.Form frmLegende 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Légende"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8490
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   7440
      TabIndex        =   28
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CTRL-C / CTRL-V"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Permet de copier une pièce (CTRL-C) et de la coller à un autre endroit (CTRL-V)."
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5040
      TabIndex        =   9
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Permet d'entrer un nom d'ID (Projets / Soumissions électriques)."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5040
      TabIndex        =   20
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   19
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Permet de mettre une pièce non chargeable."
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5040
      TabIndex        =   17
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   16
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Permet de mettre une date de facturation."
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5040
      TabIndex        =   15
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Entrée"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   14
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Permet d'effacer une pièce dans le projet ou la soumission."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5040
      TabIndex        =   12
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Suppr/Delete"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   11
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Permet de continuer la recherche au prochain enregistrement dans la liste des pièces du projet/soumission."
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   5040
      TabIndex        =   6
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CTRL-N"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Permet de faire une recherche par no. de pièce dans la liste des pièces du catalogue ou du projet/soumission."
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   5040
      TabIndex        =   2
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CTRL-F"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   3840
      X2              =   3840
      Y1              =   120
      Y2              =   6120
   End
   Begin VB.Shape shpBrun 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   255
      Left            =   240
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Cette pièce est en attente d'être retournée."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   13
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "La pièce est non-chargeable."
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   27
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NC"
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
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Cette pièce provient d'un extra non-chargeable."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   22
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Shape shpRose 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   255
      Left            =   240
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape shpBleu 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   255
      Left            =   240
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Cette pièce provient d'un extra chargeable."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   23
      Top             =   4680
      Width           =   2895
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F-AAAA-MM-JJ"
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
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "La pièce a été facturée à la date AAAA-MM-JJ."
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   25
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "La commande de cette pièce a été annulée."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   21
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Shape shpVertForet 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   255
      Left            =   240
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape shpRed 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   240
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "La pièce a été retournée au fournisseur."
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   18
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Shape shpGris 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   240
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "La pièce a été reçue en entier."
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "La soumission n'est pas complète parce qu'il manque des prix."
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   720
      TabIndex        =   7
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "La pièce a besoin d'un prix."
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "La pièce a été commandée."
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "La pièce est à quoter."
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Shape shpJaune 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   240
      Top             =   1680
      Width           =   255
   End
   Begin VB.Shape shpMagenta 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   240
      Top             =   1200
      Width           =   255
   End
   Begin VB.Shape shpOrange 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   240
      Top             =   720
      Width           =   255
   End
   Begin VB.Shape shpVert 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   240
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "frmLegende"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmdfermer_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmLegende", "cmdFermer_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur
  
10      shpOrange.BackColor = COLOR_ORANGE
15      shpOrange.BorderColor = COLOR_ORANGE
  
20      shpVert.BackColor = COLOR_VERT
25      shpVert.BorderColor = COLOR_VERT
  
30      shpMagenta.BackColor = COLOR_MAGENTA
35      shpMagenta.BorderColor = COLOR_MAGENTA
  
40      shpJaune.BackColor = COLOR_JAUNE
45      shpJaune.BorderColor = COLOR_JAUNE
  
50      shpGris.BackColor = COLOR_GRIS
55      shpGris.BorderColor = COLOR_GRIS
  
60      shpRed.BackColor = COLOR_ROUGE
65      shpRed.BorderColor = COLOR_ROUGE
  
70      shpVertForet.BackColor = COLOR_VERT_FORET
75      shpVertForet.BorderColor = COLOR_VERT_FORET

80      shpBleu.BackColor = COLOR_BLEU
85      shpBleu.BorderColor = COLOR_BLEU

90      shpRose.BackColor = COLOR_ROSE
95      shpRose.BorderColor = COLOR_ROSE

100     shpBrun.BackColor = COLOR_BRUN
105     shpBrun.BorderColor = COLOR_BRUN

110     Exit Sub

AfficherErreur:

115     woups "frmLegende", "Form_Load", Err, Erl
End Sub
