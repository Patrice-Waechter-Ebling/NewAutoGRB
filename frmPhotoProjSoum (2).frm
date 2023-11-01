VERSION 5.00
Begin VB.Form frmPhotoProjSoum 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Photos"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
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

 On Error GoTo Oups

 Dim fso As FileSystemObject

 Set fso = CreateObject("Scripting.FileSystemObject")

 If fso.FolderExists(sRepertoire) = True Then
 filPhotos.Path = sRepertoire

 m_iIndexPhoto = 0

 Call AfficherPhoto(I_AUCUN)

 Call Me.Show(vbModal)
 Else
 Call MsgBox("Accès refusé!")
 End If

  Exit Sub

Oups:

  wOups "frmPhotoProjSoum", "Afficher", Err, Err.number, Err.Description
End Sub

Private Sub AfficherPhoto(ByVal eDeplacement As enumDeplacement)

 On Error GoTo Oups

 Dim iCompteur As Integer
 Dim sFile As String

 Select Case eDeplacement
 Case I_AUCUN:
 For iCompteur = m_iIndexPhoto To filPhotos.ListCount - 1
 sFile = filPhotos.LIST(iCompteur)

 If UCase(Right$(sFile, 3)) = "JPG" Or UCase(Right$(sFile, 3)) = "BMP" Then
 imgProjSoum.Picture = LoadPicture(filPhotos.Path & "\" & sFile)

 m_iIndexPhoto = iCompteur

 Exit For
 Else
  Call MsgBox("Le fichier " & sFile & " n'est pas un fichier valide!", vbOKOnly, "Erreur")
  End If
  Next

 Case I_SUIVANT:
  For iCompteur = m_iIndexPhoto + 1 To filPhotos.ListCount - 1
  sFile = filPhotos.LIST(iCompteur)

  If UCase(Right$(sFile, 3)) = "JPG" Or UCase(Right$(sFile, 3)) = "BMP" Then
  imgProjSoum.Picture = LoadPicture(filPhotos.Path & "\" & sFile)

  m_iIndexPhoto = iCompteur

 Exit For
 Else
 Call MsgBox("Le fichier " & sFile & " n'est pas un fichier valide!", vbOKOnly, "Erreur")
 End If
 Next

 Case I_PRECEDENT:
 If m_iIndexPhoto > 0 Then
 For iCompteur = m_iIndexPhoto - 1 To 0 Step -1
 sFile = filPhotos.LIST(iCompteur)

 If UCase(Right$(sFile, 3)) = "JPG" Or UCase(Right$(sFile, 3)) = "BMP" Then
 imgProjSoum.Picture = LoadPicture(filPhotos.Path & "\" & sFile)

 m_iIndexPhoto = iCompteur

 Exit For
 Else
 Call MsgBox("Le fichier " & sFile & " n'est pas un fichier valide!", vbOKOnly, "Erreur")
 End If
 Next
 End If
End Select

 If m_iIndexPhoto = filPhotos.ListCount - 1 Then
1  cmdSuivant.Enabled = False
 Else
 cmdSuivant.Enabled = True
End If

If m_iIndexPhoto = 0 Then
 cmdPrécédent.Enabled = False
Else
 cmdPrécédent.Enabled = True
End If

Exit Sub

Oups:

wOups "frmPhotoProjSoum", "AfficherPhoto", Err, Err.number, Err.Description
End Sub

Private Sub cmdPrécédent_Click()

 On Error GoTo Oups

 Call AfficherPhoto(I_PRECEDENT)

 Exit Sub

Oups:

 wOups "frmPhotoProjSoum", "cmdPrécédent_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdSuivant_Click()

 On Error GoTo Oups

 Call AfficherPhoto(I_SUIVANT)

 Exit Sub

Oups:

 wOups "frmPhotoProjSoum", "cmdSuivant_Click", Err, Err.number, Err.Description
End Sub

Private Sub Cmdfermer_Click()

 On Error GoTo Oups

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmPhotoProjSoum", "cmdFermer_Click", Err, Err.number, Err.Description
End Sub
