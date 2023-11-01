VERSION 5.00
Begin VB.Form frmPhotoProjSoum 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Photos"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPhotoProjSoum.frx":0000
   ScaleHeight     =   5730
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrécédent 
      Caption         =   "Précédente"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdSuivant 
      Caption         =   "Suivante"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   5280
      Width           =   1095
   End
   Begin VB.FileListBox filPhotos 
      Height          =   285
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image imgProjSoum 
      Height          =   3975
      Left            =   120
      Stretch         =   -1  'True
      Top             =   720
      Width           =   7215
   End
End
Attribute VB_Name = "frmPhotoProjSoum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enumDeplacement
  I_AUCUN = 0
  I_PRECEDENT = 1
  I_SUIVANT = 2
End Enum

Private m_iIndexPhoto As Integer

Public Sub Afficher(ByVal sRepertoire As String)

5      On Error GoTo AfficherErreur

10     Dim fso As FileSystemObject

15     Set fso = CreateObject("Scripting.FileSystemObject")

20     If fso.FolderExists(sRepertoire) = True Then
25       filPhotos.Path = sRepertoire

30       m_iIndexPhoto = 0

35       Call AfficherPhoto(I_AUCUN)

40       Call Me.Show(vbModal)
45     Else
50       Call MsgBox("Accès refusé!")
55     End If

60     Exit Sub

AfficherErreur:

65     woups "frmPhotoProjSoum", "Afficher", Err, Erl
End Sub

Private Sub AfficherPhoto(ByVal eDeplacement As enumDeplacement)

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer
15      Dim sFile     As String

20      Select Case eDeplacement
          Case I_AUCUN:
25          For iCompteur = m_iIndexPhoto To filPhotos.ListCount - 1
30            sFile = filPhotos.LIST(iCompteur)

35            If UCase(Right$(sFile, 3)) = "JPG" Or UCase(Right$(sFile, 3)) = "BMP" Then
40              imgProjSoum.Picture = LoadPicture(filPhotos.Path & "\" & sFile)

45              m_iIndexPhoto = iCompteur

50              Exit For
55            Else
60              Call MsgBox("Le fichier " & sFile & " n'est pas un fichier valide!", vbOKOnly, "Erreur")
65            End If
70          Next

          Case I_SUIVANT:
75          For iCompteur = m_iIndexPhoto + 1 To filPhotos.ListCount - 1
80            sFile = filPhotos.LIST(iCompteur)

85            If UCase(Right$(sFile, 3)) = "JPG" Or UCase(Right$(sFile, 3)) = "BMP" Then
90              imgProjSoum.Picture = LoadPicture(filPhotos.Path & "\" & sFile)

95              m_iIndexPhoto = iCompteur

100             Exit For
105           Else
110             Call MsgBox("Le fichier " & sFile & " n'est pas un fichier valide!", vbOKOnly, "Erreur")
115           End If
120         Next

          Case I_PRECEDENT:
125         If m_iIndexPhoto > 0 Then
130           For iCompteur = m_iIndexPhoto - 1 To 0 Step -1
135             sFile = filPhotos.LIST(iCompteur)

140             If UCase(Right$(sFile, 3)) = "JPG" Or UCase(Right$(sFile, 3)) = "BMP" Then
145               imgProjSoum.Picture = LoadPicture(filPhotos.Path & "\" & sFile)

150               m_iIndexPhoto = iCompteur

155               Exit For
160             Else
165               Call MsgBox("Le fichier " & sFile & " n'est pas un fichier valide!", vbOKOnly, "Erreur")
170             End If
175           Next
180         End If
185     End Select

190     If m_iIndexPhoto = filPhotos.ListCount - 1 Then
195       cmdSuivant.Enabled = False
200     Else
205       cmdSuivant.Enabled = True
210     End If

215     If m_iIndexPhoto = 0 Then
220       cmdPrécédent.Enabled = False
225     Else
230       cmdPrécédent.Enabled = True
235     End If

240     Exit Sub

AfficherErreur:

245     woups "frmPhotoProjSoum", "AfficherPhoto", Err, Erl
End Sub

Private Sub cmdPrécédent_Click()

5       On Error GoTo AfficherErreur

10      Call AfficherPhoto(I_PRECEDENT)

15      Exit Sub

AfficherErreur:

20      woups "frmPhotoProjSoum", "cmdPrécédent_Click", Err, Erl
End Sub

Private Sub cmdSuivant_Click()

5       On Error GoTo AfficherErreur

10      Call AfficherPhoto(I_SUIVANT)

15      Exit Sub

AfficherErreur:

20      woups "frmPhotoProjSoum", "cmdSuivant_Click", Err, Erl
End Sub

Private Sub Cmdfermer_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmPhotoProjSoum", "cmdFermer_Click", Err, Erl
End Sub
