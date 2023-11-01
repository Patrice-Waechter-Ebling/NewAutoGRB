VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRetourMarchandise 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retour de marchandise"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmRetourMarchandise.frx":0000
   ScaleHeight     =   7680
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.MonthView mvwRetour 
      Height          =   2370
      Left            =   9120
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   0
      Appearance      =   1
      StartOfWeek     =   90243073
      CurrentDate     =   38170
   End
   Begin VB.TextBox txtDateRetour 
      Height          =   285
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdDate 
      Caption         =   "..."
      Height          =   285
      Left            =   11400
      TabIndex        =   5
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton cmdImprimer 
      Caption         =   "Imprimer"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   10560
      TabIndex        =   8
      Top             =   7200
      Width           =   1215
   End
   Begin VB.ComboBox cmbNoProjet 
      Height          =   315
      Left            =   3360
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin MSComctlLib.ListView lvwProjet 
      Height          =   6255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   11033
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Qté"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "No. Item"
         Object.Width           =   3228
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   6720
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Manufacturier"
         Object.Width           =   2037
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Distributeur"
         Object.Width           =   1773
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "No. Retour"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtNoProjet 
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date de retour :"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8400
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmRetourMarchandise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Index des colonnes de lvwSoumission
Private Const I_COL_SOUM_QUANTITE  As Integer = 0
Private Const I_COL_SOUM_PIECE     As Integer = 1
Private Const I_COL_SOUM_DESCR     As Integer = 2
Private Const I_COL_SOUM_MANUFACT  As Integer = 3
Private Const I_COL_SOUM_DISTRIB   As Integer = 4
Private Const I_COL_SOUM_NO_RETOUR As Integer = 5
Private Const I_COL_SOUM_DATE      As Integer = 6

Private Enum enumTypeRetour
  PROJET = 0
  ACHAT = 1
End Enum

Private m_sUserID          As String

'Pour l'impression
Public m_bAnnuleImpression As Boolean
Public m_eTypeImpression   As enumImpressionRetour

Private m_eTypeRetour      As enumTypeRetour

Public Sub Afficher(ByVal sNoProjet As String, ByVal sUserID As String)

5       On Error GoTo AfficherErreur

10      m_eTypeRetour = PROJET

15      m_sUserID = sUserID

20      Call RemplirComboProjetElec
25      Call RemplirComboProjetMec

30      Call NouveauRetour(sNoProjet)

35      Call Me.Show(vbModal)

40      Exit Sub

AfficherErreur:

45      woups "frmRetourMarchandise", "Afficher", Err, Erl
End Sub

Public Sub AfficherAchat(ByVal sNoAchat As String, ByVal iIndexAchat As Integer, ByVal sUserID As String)

5       On Error GoTo AfficherErreur

10      m_eTypeRetour = ACHAT

15      m_sUserID = sUserID

20      Call RemplirComboAchats

25      Call NouveauRetourAchat(sNoAchat, iIndexAchat)

30      Call Me.Show(vbModal)

35      Exit Sub

AfficherErreur:

40      woups "frmRetourMarchandise", "AfficherAchat", Err, Erl
End Sub

Private Sub cmbNoProjet_Click()

5       On Error GoTo AfficherErreur
  
10      Screen.MousePointer = vbHourglass
  
15      txtnoprojet.Text = cmbNoProjet.Text
  
20      If m_eTypeRetour = ACHAT Then
          'Rempli les valeurs de l'achat sélectionné
25        Call RemplirListViewAchat
30      Else
          'Rempli les valeurs du projet sélectionné
35        Call RemplirListViewProjet
40      End If

45      Screen.MousePointer = vbDefault

50      Exit Sub

AfficherErreur:

55      woups "frmRetourMarchandise", "cmbNoProjet_Click", Err, Erl
End Sub

Private Sub cmdImprimer_Click()

5       On Error GoTo AfficherErreur

10      Dim iCompteur     As Integer
15      Dim bChecked      As Boolean
20      Dim rstProjet     As ADODB.Recordset
25      Dim bRetourPermis As Boolean

30      If cmbNoProjet.ListIndex > -1 Then
35        Set rstProjet = New ADODB.Recordset

40        If m_eTypeRetour = PROJET Then
45          If cmbNoProjet.ItemData(cmbNoProjet.ListIndex) = 0 Then
50            Call rstProjet.Open("SELECT Modification, Par FROM GRB_ProjetElec WHERE IDProjet = '" & Right$(txtnoprojet.Text, Len(txtnoprojet.Text) - 1) & "'", g_connData, adOpenDynamic, adLockOptimistic)
55          Else
60            Call rstProjet.Open("SELECT Modification, Par FROM GRB_ProjetMec WHERE IDProjet = '" & Right$(txtnoprojet.Text, Len(txtnoprojet.Text) - 1) & "'", g_connData, adOpenDynamic, adLockOptimistic)
65          End If

70          If rstProjet.Fields("Modification") = False Then
75            bRetourPermis = True
80          Else
85            bRetourPermis = False
90          End If

95          Call rstProjet.Close
100         Set rstProjet = Nothing
105       Else
110         bRetourPermis = True
115       End If

120       If bRetourPermis = True Then
125         For iCompteur = 1 To lvwProjet.ListItems.count
130           If lvwProjet.ListItems(iCompteur).Checked = True Then
135             bChecked = True
 
140             Exit For
145           End If
150         Next

155         If bChecked = True Then
160           Call OuvrirForm(frmChoixImpressionRetourMarchandise, True)

165           If m_bAnnuleImpression = False Then
170             If m_eTypeImpression = MODE_DEMANDE_RETOUR Then
175               Call ImprimerDemandeRetour
180             Else
185               Call ImprimerRetour
190             End If
195           End If
200         End If
205       Else
210         Call MsgBox("Ce projet est en modification par " & rstProjet.Fields("Par") & "!", vbOKOnly, "Erreur")
215       End If
220     Else
225       Call MsgBox("Vous devez choisir un numéro de retour!", vbOKOnly, "Erreur")
230     End If

235     Exit Sub

AfficherErreur:

240     woups "frmRetourMarchandise", "cmdImprimer_Click", Err, Erl
End Sub

Private Sub ImprimerRetour()

5       On Error GoTo AfficherErreur

10      Dim collPiece   As Collection
15      Dim collNoLigne As Collection
20      Dim iCompteur   As Integer
25      Dim sNoBC       As String

30      sNoBC = InputBox("Quel est le numéro du retour?", , txtnoprojet.Text)

35      If sNoBC <> vbNullString Then
40        Set collPiece = New Collection
45        Set collNoLigne = New Collection

50        For iCompteur = 1 To lvwProjet.ListItems.count
55          If lvwProjet.ListItems(iCompteur).Checked = True Then
60            Call collPiece.Add(lvwProjet.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE))
65            Call collNoLigne.Add(lvwProjet.ListItems(iCompteur).Tag)
70          End If
75        Next

80        If m_eTypeRetour = ACHAT Then
85          Call frmBonCommande.AfficherFormRetourMarchandiseAchat(Replace(Trim$(Left$(txtnoprojet.Text, InStrRev(txtnoprojet.Text, "-") - 1)), "R", ""), CInt(Right$(txtnoprojet.Text, 3)), sNoBC, collPiece, collNoLigne, m_sUserID, MODE_RETOUR)
90        Else
95          Call frmBonCommande.AfficherFormRetourMarchandiseProjet(txtnoprojet.Text, sNoBC, collPiece, collNoLigne, m_sUserID, MODE_RETOUR)
100       End If
105     End If

110     Exit Sub

AfficherErreur:

115     woups "frmRetourMarchandise", "ImprimerRetour", Err, Erl
End Sub

Private Sub ImprimerDemandeRetour()

5       On Error GoTo AfficherErreur

10      Dim collPiece   As Collection
15      Dim collNoLigne As Collection
20      Dim iCompteur   As Integer
25      Dim sNoBC       As String

30      sNoBC = InputBox("Quel est le numéro de la demande de retour?", , txtnoprojet.Text)

35      If sNoBC <> vbNullString Then
40        Set collPiece = New Collection
45        Set collNoLigne = New Collection

50        For iCompteur = 1 To lvwProjet.ListItems.count
55          If lvwProjet.ListItems(iCompteur).Checked = True Then
60            Call collPiece.Add(lvwProjet.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE))
65            Call collNoLigne.Add(lvwProjet.ListItems(iCompteur).Tag)
70          End If
75        Next

80        If m_eTypeRetour = ACHAT Then
85          Call frmBonCommande.AfficherFormRetourMarchandiseAchat(Trim$(Replace(Left$(txtnoprojet.Text, InStrRev(txtnoprojet.Text, "-") - 1), "R", "")), CInt(Right$(txtnoprojet.Text, 3)), sNoBC, collPiece, collNoLigne, m_sUserID, MODE_DEMANDE_RETOUR)
90        Else
95          Call frmBonCommande.AfficherFormRetourMarchandiseProjet(txtnoprojet.Text, sNoBC, collPiece, collNoLigne, m_sUserID, MODE_DEMANDE_RETOUR)
100       End If
105     End If

110     Exit Sub

AfficherErreur:

115     woups "frmRetourMarchandise", "ImprimerDemandeRetour", Err, Erl
End Sub

Private Sub NouveauRetour(ByVal sNoProjet As String)

5       On Error GoTo AfficherErreur

10      Dim rstProjet As ADODB.Recordset
15      Dim iCompteur As Integer
20      Dim bExiste   As Boolean
25      Dim eType     As enumCatalogue

30      bExiste = False

35      If ComboContient(cmbNoProjet, "R" & sNoProjet) = False Then
40        Set rstProjet = New ADODB.Recordset

45        Call rstProjet.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

50        If Not rstProjet.EOF Then
55          bExiste = True

60          eType = ELECTRIQUE
65        End If

70        Call rstProjet.Close

75        If bExiste = False Then
80          Call rstProjet.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

85          If Not rstProjet.EOF Then
90            bExiste = True

95            eType = MECANIQUE
100         End If

105         Call rstProjet.Close
110       End If

115       If bExiste = True Then
120         Call cmbNoProjet.AddItem("R" & sNoProjet)

125         cmbNoProjet.ItemData(cmbNoProjet.newIndex) = eType

130         cmbNoProjet.ListIndex = -1
135       Else
140         Call MsgBox("Projet inexistant!", vbOKOnly, "Erreur")
145       End If

150       Set rstProjet = Nothing
155     End If

160     For iCompteur = 0 To cmbNoProjet.ListCount - 1
165       If cmbNoProjet.LIST(iCompteur) = "R" & sNoProjet Then
170         cmbNoProjet.ListIndex = iCompteur

175         Exit Sub
180       End If
185     Next

190     Exit Sub

AfficherErreur:

195     woups "frmRetourMarchandise", "NouveauRetour", Err, Erl
End Sub

Private Sub NouveauRetourAchat(ByVal sNoAchat As String, ByVal iIndexAchat As Integer)

5       On Error GoTo AfficherErreur

10      Dim rstAchat  As ADODB.Recordset
15      Dim iCompteur As Integer
20      Dim bExiste   As Boolean
25      Dim eType     As enumCatalogue

30      bExiste = False

35      If ComboContient(cmbNoProjet, "R" & sNoAchat & "-" & Right$("000" & iIndexAchat, 3)) = False Then
40        Set rstAchat = New ADODB.Recordset

45        Call rstAchat.Open("SELECT * FROM GRB_Achat WHERE IDAchat = '" & sNoAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenDynamic, adLockOptimistic)

50        If Not rstAchat.EOF Then
55          bExiste = True

60          If rstAchat.Fields("Type") = "M" Then
65            eType = MECANIQUE
70          Else
75            eType = ELECTRIQUE
80          End If
85        End If

90        Call rstAchat.Close
95        Set rstAchat = Nothing

100       If bExiste = True Then
105         Call cmbNoProjet.AddItem("R" & sNoAchat & "-" & Right$("000" & iIndexAchat, 3))

110         cmbNoProjet.ItemData(cmbNoProjet.newIndex) = eType
115       Else
120         Call MsgBox("Projet inexistant!", vbOKOnly, "Erreur")
125       End If
130     End If

135     For iCompteur = 0 To cmbNoProjet.ListCount - 1
140       If cmbNoProjet.LIST(iCompteur) = "R" & sNoAchat & "-" & Right$("000" & iIndexAchat, 3) Then
145         cmbNoProjet.ListIndex = iCompteur

150         Exit For
155       End If
160     Next


165     Exit Sub

AfficherErreur:

170     woups "frmRetourMarchandise", "NouveauRetourAchat", Err, Erl
End Sub

Private Sub Cmdfermer_Click()

5       On Error GoTo AfficherErreur
 
10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmRetourMarchandise", "cmdFermer_Click", Err, Erl
End Sub

Private Sub RemplirComboProjetElec()

5       On Error GoTo AfficherErreur

        'Rempli le combo des projets
10      Dim rstProjet As ADODB.Recordset
  
15      Set rstProjet = New ADODB.Recordset
  
        'Ouvre le recordset selon le type
20      Call rstProjet.Open("SELECT DISTINCT GRB_ProjetElec.IDProjet FROM GRB_ProjetElec INNER JOIN GRB_Projet_Pieces ON GRB_ProjetElec.IDProjet = GRB_Projet_Pieces.IDProjet WHERE Retour = True ORDER BY GRB_ProjetElec.IDProjet", g_connData, adOpenDynamic, adLockOptimistic)
    
        'Tant que ce n'est pas la fin des enregistrements
25      Do While Not rstProjet.EOF
30        Call cmbNoProjet.AddItem("R" & rstProjet.Fields("IDProjet"))

35        cmbNoProjet.ItemData(cmbNoProjet.newIndex) = 0

40        Call rstProjet.MoveNext
45      Loop

50      Call rstProjet.Close
55      Set rstProjet = Nothing

60      Exit Sub

AfficherErreur:

65      woups "frmRetourMarchandise", "RemplirComboProjetElec", Err, Erl
End Sub

Private Sub RemplirComboProjetMec()

5       On Error GoTo AfficherErreur

        'Rempli le combo des projets
10      Dim rstProjet As ADODB.Recordset
    
15      Set rstProjet = New ADODB.Recordset
    
        'Ouvre le recordset selon le type
20      Call rstProjet.Open("SELECT DISTINCT GRB_ProjetMec.IDProjet FROM GRB_ProjetMec INNER JOIN GRB_Projet_Pieces ON GRB_ProjetMec.IDProjet = GRB_Projet_Pieces.IDProjet WHERE Retour = True  ORDER BY GRB_ProjetMec.IDProjet", g_connData, adOpenDynamic, adLockOptimistic)
        
        'Tant que ce n'est pas la fin des enregistrements
25      Do While Not rstProjet.EOF
30        Call cmbNoProjet.AddItem("R" & rstProjet.Fields("IDProjet"))

35        cmbNoProjet.ItemData(cmbNoProjet.newIndex) = 1

40        Call rstProjet.MoveNext
45      Loop

50      Call rstProjet.Close
55      Set rstProjet = Nothing

60      Exit Sub

AfficherErreur:

65      woups "frmRetourMarchandise", "RemplirComboProjetMec", Err, Erl
End Sub

Private Sub RemplirComboAchats()

5       On Error GoTo AfficherErreur

        'Rempli le combo des projets
10      Dim rstAchat As ADODB.Recordset
  
15      Set rstAchat = New ADODB.Recordset
  
        'Ouvre le recordset selon le type
20      Call rstAchat.Open("SELECT DISTINCT GRB_Achat.IDAchat, GRB_Achat.IndexAchat, GRB_Achat.Type FROM GRB_Achat INNER JOIN GRB_Achat_Pieces ON GRB_Achat.IDAchat = GRB_Achat_Pieces.IDAchat AND GRB_Achat.IndexAchat = GRB_Achat_Pieces.IndexAchat WHERE GRB_Achat_Pieces.Retour = True ORDER BY GRB_Achat.IDAchat, GRB_Achat.IndexAchat", g_connData, adOpenDynamic, adLockOptimistic)
    
        'Tant que ce n'est pas la fin des enregistrements
25      Do While Not rstAchat.EOF
30        Call cmbNoProjet.AddItem("R" & rstAchat.Fields("IDAchat") & "-" & Right$("000" & rstAchat.Fields("IndexAchat"), 3))

35        If rstAchat.Fields("Type") = "E" Then
40          cmbNoProjet.ItemData(cmbNoProjet.newIndex) = 0
45        Else
50          cmbNoProjet.ItemData(cmbNoProjet.newIndex) = 1
55        End If
          
60        Call rstAchat.MoveNext
65      Loop

70      Call rstAchat.Close
75      Set rstAchat = Nothing

80      Exit Sub

AfficherErreur:

85      woups "frmRetourMarchandise", "RemplirComboAchats", Err, Erl
End Sub

Private Sub RemplirListViewProjet()

5       On Error GoTo AfficherErreur

        'Remplis les pièces de la soumission avec la BD
10      Dim rstProjet     As ADODB.Recordset
15      Dim rstFRS        As ADODB.Recordset
20      Dim itmProjet     As ListItem
25      Dim lColor        As Long
  
30      If cmbNoProjet.ListIndex <> -1 Then
35        Call lvwProjet.ListItems.Clear
  
40        Set rstProjet = New ADODB.Recordset
45        Set rstFRS = New ADODB.Recordset
  
50        Call rstProjet.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & Right$(txtnoprojet.Text, Len(txtnoprojet.Text) - 1) & "' AND Left$(Qté,1) = '-' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

55        Do While Not rstProjet.EOF
60          Set itmProjet = lvwProjet.ListItems.Add

65          itmProjet.Checked = False
   
70          If rstProjet.Fields("Retour") = True Then
75            lColor = COLOR_ROUGE
80          Else
85            If rstProjet.Fields("Commandé") = True Then
90              lColor = COLOR_ORANGE     'COLOR_ORANGE
95            Else
100             If rstProjet.Fields("Recu") = True Then
125               lColor = COLOR_GRIS 'Gris
150             Else
155               If rstProjet.Fields("MatérielInutile") = True Then
160                 lColor = COLOR_BRUN
165               Else
170                 lColor = COLOR_NOIR
175               End If
180             End If
185           End If
190         End If

            'No Ligne
195         itmProjet.Tag = rstProjet.Fields("NuméroLigne")
   
            'Quantité
200         If Not IsNull(rstProjet.Fields("Qté")) Then
205           itmProjet.Text = rstProjet.Fields("Qté")
210         Else
215           itmProjet.Text = vbNullString
220         End If

225         itmProjet.ForeColor = lColor
    
            'Numéro d'item
230         If Not IsNull(rstProjet.Fields("NumItem")) Then
235           itmProjet.SubItems(I_COL_SOUM_PIECE) = rstProjet.Fields("NumItem")
240         Else
245           itmProjet.SubItems(I_COL_SOUM_PIECE) = vbNullString
250         End If
      
255         itmProjet.ListSubItems(I_COL_SOUM_PIECE).ForeColor = lColor
    
            'On met le nom de la sous-section dans le tag du numéro d'item
260         itmProjet.ListSubItems(I_COL_SOUM_PIECE).Tag = rstProjet.Fields("SousSection")
    
            'Description en francais
265         If Not IsNull(rstProjet.Fields("Desc_FR")) Then
270           itmProjet.SubItems(I_COL_SOUM_DESCR) = rstProjet.Fields("Desc_FR")
275         Else
280           itmProjet.SubItems(I_COL_SOUM_DESCR) = vbNullString
285         End If
    
290         itmProjet.ListSubItems(I_COL_SOUM_DESCR).ForeColor = lColor
   
            'On met la description en anglais dans le tag de la description en francais
295         If Not IsNull(rstProjet.Fields("DESC_EN")) Then
300           itmProjet.ListSubItems(I_COL_SOUM_DESCR).Tag = rstProjet.Fields("Desc_EN")
305         Else
310           itmProjet.ListSubItems(I_COL_SOUM_DESCR).Tag = vbNullString
315         End If
   
            'Fabricant
320         If Not IsNull(rstProjet.Fields("Manufact")) Then
325           itmProjet.SubItems(I_COL_SOUM_MANUFACT) = rstProjet.Fields("Manufact")
330         Else
335           itmProjet.SubItems(I_COL_SOUM_MANUFACT) = vbNullString
340         End If
    
345         itmProjet.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = lColor
          
            'Fournisseur
350         If Not IsNull(rstProjet.Fields("IDFRS")) And rstProjet.Fields("IDFRS") > 0 Then
355           If itmProjet.SubItems(I_COL_SOUM_PIECE) <> "Texte" Then
360             Call rstFRS.Open("SELECT NomFournisseur FROM GRB_Fournisseur WHERE IDFRS = " & rstProjet.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
    
                'On affiche le nom dans la colonne
365             itmProjet.SubItems(I_COL_SOUM_DISTRIB) = rstFRS.Fields("NomFournisseur")
          
370             itmProjet.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = lColor
         
                'On affiche l'Id dans le tag
375             itmProjet.ListSubItems(I_COL_SOUM_DISTRIB).Tag = rstProjet.Fields("IDFRS")
        
380             Call rstFRS.Close
385           End If
390         Else
395           itmProjet.SubItems(I_COL_SOUM_DISTRIB) = vbNullString
400         End If

405         If Not IsNull(rstProjet.Fields("NoRetour")) Then
410           itmProjet.SubItems(I_COL_SOUM_NO_RETOUR) = rstProjet.Fields("NoRetour")
415         Else
420           itmProjet.SubItems(I_COL_SOUM_NO_RETOUR) = vbNullString
425         End If

430         itmProjet.ListSubItems(I_COL_SOUM_NO_RETOUR).ForeColor = lColor

435         If Not IsNull(rstProjet.Fields("DateRetour")) Then
440           itmProjet.SubItems(I_COL_SOUM_DATE) = rstProjet.Fields("DateRetour")
445         Else
450           itmProjet.SubItems(I_COL_SOUM_DATE) = vbNullString
455         End If

460         itmProjet.ListSubItems(I_COL_SOUM_DATE).ForeColor = lColor
             
465         Call rstProjet.MoveNext
470       Loop
  
475       Call rstProjet.Close
480       Set rstProjet = Nothing

485       Set rstFRS = Nothing
490     End If
 
495     Exit Sub

AfficherErreur:

500     woups "frmRetourMarchandise", "RemplirListViewProjet", Err, Erl
End Sub

Private Sub RemplirListViewAchat()

5       On Error GoTo AfficherErreur

        'Remplis les pièces de la soumission avec la BD
10      Dim rstAchat As ADODB.Recordset
15      Dim rstFRS   As ADODB.Recordset
20      Dim itmAchat As ListItem
25      Dim lColor   As Long
  
30      If cmbNoProjet.ListIndex <> -1 Then
35        Call lvwProjet.ListItems.Clear
  
40        Set rstAchat = New ADODB.Recordset
45        Set rstFRS = New ADODB.Recordset
  
50        Call rstAchat.Open("SELECT * FROM GRB_Achat_Pieces WHERE IDAchat = '" & Replace(Trim$(Left$(txtnoprojet.Text, InStrRev(txtnoprojet.Text, "-") - 1)), "R", "") & "' AND IndexAchat = " & CInt(Right$(txtnoprojet.Text, 3)) & " AND Left$(Qté,1) = '-' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

55        Do While Not rstAchat.EOF
60          Set itmAchat = lvwProjet.ListItems.Add

65          itmAchat.Checked = False
   
70          If rstAchat.Fields("Retour") = True Then
75            lColor = COLOR_ROUGE
80          Else
85            If rstAchat.Fields("Commandé") = True Then
90              lColor = COLOR_ORANGE     'COLOR_ORANGE
95            End If
100         End If

            'No Ligne
105          itmAchat.Tag = rstAchat.Fields("NuméroLigne")
   
            'Quantité
110         If Not IsNull(rstAchat.Fields("Qté")) Then
115           itmAchat.Text = rstAchat.Fields("Qté")
120         Else
125           itmAchat.Text = vbNullString
130         End If

135         itmAchat.ForeColor = lColor
    
            'Numéro d'item
140         If Not IsNull(rstAchat.Fields("PIECE")) Then
145           itmAchat.SubItems(I_COL_SOUM_PIECE) = rstAchat.Fields("PIECE")
150         Else
155           itmAchat.SubItems(I_COL_SOUM_PIECE) = vbNullString
160         End If
      
165         itmAchat.ListSubItems(I_COL_SOUM_PIECE).ForeColor = lColor
        
            'Description en francais
170         If Not IsNull(rstAchat.Fields("DESC_FR")) Then
175           itmAchat.SubItems(I_COL_SOUM_DESCR) = rstAchat.Fields("DESC_FR")
180         Else
185           itmAchat.SubItems(I_COL_SOUM_DESCR) = vbNullString
190         End If
    
195         itmAchat.ListSubItems(I_COL_SOUM_DESCR).ForeColor = lColor
   
            'On met la description en anglais dans le tag de la description en francais
200         If Not IsNull(rstAchat.Fields("DESC_EN")) Then
205           itmAchat.ListSubItems(I_COL_SOUM_DESCR).Tag = rstAchat.Fields("DESC_EN")
210         Else
215           itmAchat.ListSubItems(I_COL_SOUM_DESCR).Tag = vbNullString
220         End If
    
            'Fabricant
225         If Not IsNull(rstAchat.Fields("Manufact")) Then
230           itmAchat.SubItems(I_COL_SOUM_MANUFACT) = rstAchat.Fields("Manufact")
235         Else
240           itmAchat.SubItems(I_COL_SOUM_MANUFACT) = vbNullString
245         End If
    
250         itmAchat.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = lColor
      
            'Fournisseur
255         If Not IsNull(rstAchat.Fields("IDFRS")) And rstAchat.Fields("IDFRS") > 0 Then
260           Call rstFRS.Open("SELECT NomFournisseur FROM GRB_Fournisseur WHERE IDFRS = " & rstAchat.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
    
              'On affiche le nom dans la colonne
265           itmAchat.SubItems(I_COL_SOUM_DISTRIB) = rstFRS.Fields("NomFournisseur")
          
270           itmAchat.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = lColor
         
              'On affiche l'Id dans le tag
275           itmAchat.ListSubItems(I_COL_SOUM_DISTRIB).Tag = rstAchat.Fields("IDFRS")
        
280           Call rstFRS.Close
285         Else
290           itmAchat.SubItems(I_COL_SOUM_DISTRIB) = vbNullString
295         End If

300         If Not IsNull(rstAchat.Fields("NoRetour")) Then
305           itmAchat.SubItems(I_COL_SOUM_NO_RETOUR) = rstAchat.Fields("NoRetour")
310         Else
315           itmAchat.SubItems(I_COL_SOUM_NO_RETOUR) = vbNullString
320         End If

325         itmAchat.ListSubItems(I_COL_SOUM_NO_RETOUR).ForeColor = lColor

330         If Not IsNull(rstAchat.Fields("DateRetour")) Then
335           itmAchat.SubItems(I_COL_SOUM_DATE) = rstAchat.Fields("DateRetour")
340         Else
345           itmAchat.SubItems(I_COL_SOUM_DATE) = vbNullString
350         End If

355         itmAchat.ListSubItems(I_COL_SOUM_DATE).ForeColor = lColor
             
360         Call rstAchat.MoveNext
365       Loop
  
370       Call rstAchat.Close
375       Set rstAchat = Nothing

380       Set rstFRS = Nothing
385     End If
 
390     Exit Sub

AfficherErreur:

395     woups "frmRetourMarchandise", "RemplirListViewAchat", Err, Erl
End Sub

Public Sub Retour()

5       On Error GoTo AfficherErreur

10      Dim rstBC           As ADODB.Recordset
15      Dim rstBCPiece      As ADODB.Recordset
20      Dim rstPiece        As ADODB.Recordset
25      Dim rstModif        As ADODB.Recordset
30      Dim rstInventaire   As ADODB.Recordset
35      Dim rstInvModif     As ADODB.Recordset
40      Dim rstEmploye      As ADODB.Recordset
45      Dim sWhere          As String
50      Dim sWherePiece     As String
55      Dim sWhereNoLigne   As String
60      Dim bPremier        As Boolean
65      Dim bPremierNoLigne As Boolean
70      Dim bRetourFait     As Boolean
75      Dim sPiece          As String
80      Dim sNoLigne        As String
85      Dim sNoRetour       As String

90      sNoRetour = DR_Commande.Sections("Section2").Controls("lblNoSoum").Caption
          
95      Set rstBC = New ADODB.Recordset
100     Set rstBCPiece = New ADODB.Recordset
105     Set rstPiece = New ADODB.Recordset
          
110     Call rstBC.Open("SELECT * FROM GRB_BonsCommandes WHERE NoBonCommande = '" & sNoRetour & "'", g_connData, adOpenDynamic, adLockOptimistic)
         
        'Pour chaque enregistrement
115     Call rstBCPiece.Open("SELECT NoItem, NuméroLigne FROM GRB_BonsCommandes_Pieces WHERE NoBonCommande = '" & sNoRetour & "'", g_connData, adOpenDynamic, adLockOptimistic)
                
        'Tant que ce n'est pas la fin des enregistrements
120     If m_eTypeRetour = ACHAT Then
125       sWhere = "(IDAchat = '" & Replace(Trim$(Left$(txtnoprojet.Text, InStrRev(txtnoprojet.Text, "-") - 1)), "R", "") & "' AND IndexAchat = " & Int(Right$(txtnoprojet.Text, 3)) & ")"

130       sWherePiece = "PIECE In ("
135       sWhereNoLigne = "NuméroLigne In ("

140       bPremier = True

145       Do While Not rstBCPiece.EOF
150         If Not IsNull(rstBCPiece.Fields("NoItem")) Then
155           sNoLigne = rstBCPiece.Fields("NuméroLigne")

160           If bPremier = True Then
165             If InStr(1, sNoLigne, ",") = 0 Then
170               sWherePiece = sWherePiece & "'" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
175               sWhereNoLigne = sWhereNoLigne & rstBCPiece.Fields("NuméroLigne")
180             Else
185               bPremierNoLigne = True

190               Do While InStr(1, sNoLigne, ",") > 0
195                 If bPremierNoLigne = True Then
200                   sWherePiece = sWherePiece & "'" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
205                   sWhereNoLigne = sWhereNoLigne & Left$(sNoLigne, InStr(1, sNoLigne, ",") - 1)

210                   bPremierNoLigne = False
215                 Else
220                   sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
225                   sWhereNoLigne = sWhereNoLigne & ", " & Left$(sNoLigne, InStr(1, sNoLigne, ",") - 1)
230                 End If

235                 sNoLigne = Right$(sNoLigne, Len(sNoLigne) - (InStr(1, sNoLigne, ",") + 1))
240               Loop

245               If Trim$(sNoLigne) <> "" Then
250                 sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
255                 sWhereNoLigne = sWhereNoLigne & ", " & sNoLigne
260               End If
265             End If

270             bPremier = False
275           Else
280             If InStr(1, sNoLigne, ",") = 0 Then
285               sWherePiece = sWherePiece & ", '" & rstBCPiece.Fields("NoItem") & "'"
290               sWhereNoLigne = sWhereNoLigne & ", " & rstBCPiece.Fields("NuméroLigne")
295             Else
300               Do While InStr(1, sNoLigne, ",") > 0
305                 sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
310                 sWhereNoLigne = sWhereNoLigne & ", " & Left$(sNoLigne, InStr(1, sNoLigne, ",") - 1)

315                 sNoLigne = Right$(sNoLigne, Len(sNoLigne) - (InStr(1, sNoLigne, ",") + 1))
320               Loop

325               If Trim$(sNoLigne) <> "" Then
330                 sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
335                 sWhereNoLigne = sWhereNoLigne & ", " & sNoLigne
340               End If
345             End If
350           End If
355         End If

360         Call rstBCPiece.MoveNext
365       Loop

370       sWherePiece = sWherePiece & ")"
375       sWhereNoLigne = sWhereNoLigne & ")"

380       sWhere = sWhere & " AND " & sWherePiece & " AND " & sWhereNoLigne

385       Call rstPiece.Open("SELECT * FROM GRB_Achat_Pieces WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
390     Else
395       sWhere = "(IDProjet = '" & Right$(txtnoprojet.Text, Len(txtnoprojet.Text) - 1) & "')"

400       sWherePiece = "NumItem In ("
405       sWhereNoLigne = "NuméroLigne In ("

410       bPremier = True

415       Do While Not rstBCPiece.EOF
420         If Not IsNull(rstBCPiece.Fields("NoItem")) Then
425           sNoLigne = rstBCPiece.Fields("NuméroLigne")

430           If bPremier = True Then
435             If InStr(1, sNoLigne, ",") = 0 Then
440               sWherePiece = sWherePiece & "'" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
445               sWhereNoLigne = sWhereNoLigne & rstBCPiece.Fields("NuméroLigne")
450             Else
455               bPremierNoLigne = True

460               Do While InStr(1, sNoLigne, ",") > 0
465                 If bPremierNoLigne = True Then
470                   sWherePiece = sWherePiece & "'" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
475                   sWhereNoLigne = sWhereNoLigne & Left$(sNoLigne, InStr(1, sNoLigne, ",") - 1)

480                   bPremierNoLigne = False
485                 Else
490                   sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
495                   sWhereNoLigne = sWhereNoLigne & ", " & Left$(sNoLigne, InStr(1, sNoLigne, ",") - 1)
500                 End If

505                 sNoLigne = Right$(sNoLigne, Len(sNoLigne) - (InStr(1, sNoLigne, ",") + 1))
510               Loop

515               If Trim$(sNoLigne) <> "" Then
520                 sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
525                 sWhereNoLigne = sWhereNoLigne & ", " & sNoLigne
530               End If
535             End If

540             bPremier = False
545           Else
550             If InStr(1, sNoLigne, ",") = 0 Then
555               sWherePiece = sWherePiece & ", '" & rstBCPiece.Fields("NoItem") & "'"
560               sWhereNoLigne = sWhereNoLigne & ", " & rstBCPiece.Fields("NuméroLigne")
565             Else
570               Do While InStr(1, sNoLigne, ",") > 0
575                 sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
580                 sWhereNoLigne = sWhereNoLigne & ", " & Left$(sNoLigne, InStr(1, sNoLigne, ",") - 1)

585                 sNoLigne = Right$(sNoLigne, Len(sNoLigne) - (InStr(1, sNoLigne, ",") + 1))
590               Loop

595               If Trim$(sNoLigne) <> "" Then
600                 sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
605                 sWhereNoLigne = sWhereNoLigne & ", " & sNoLigne
610               End If
615             End If
620           End If
625         End If

630         Call rstBCPiece.MoveNext
635       Loop

640       sWherePiece = sWherePiece & ")"
645       sWhereNoLigne = sWhereNoLigne & ")"

650       sWhere = sWhere & " AND " & sWherePiece & " AND " & sWhereNoLigne

655       Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
660     End If

665     Call rstBCPiece.Close
670     Set rstBCPiece = Nothing

675     Set rstInventaire = New ADODB.Recordset
680     Set rstInvModif = New ADODB.Recordset

685     Do While Not rstPiece.EOF
690       Call rstBC.MoveFirst

695       Do While Not rstBC.EOF
700         If rstBC.Fields("NoFournisseur") = rstPiece.Fields("IDFRS") Then
705           If rstPiece.Fields("Retour") = True Then
710             bRetourFait = True
715           Else
720             bRetourFait = False
725           End If

730           rstPiece.Fields("DateRetour") = txtDateRetour.Text

735           rstPiece.Fields("Retour") = True
740           rstPiece.Fields("NoRetour") = rstBC.Fields("NoBonCommande")

745           If m_eTypeRetour = PROJET Then
750             rstPiece.Fields("MatérielInutile") = False
755           End If

760           Call rstPiece.Update

765           If bRetourFait = False Then
770             If rstPiece.Fields("IDFRS") = 717 Then
775               If m_eTypeRetour = ACHAT Then
780                 sPiece = rstPiece.Fields("PIECE")
785               Else
790                 sPiece = rstPiece.Fields("NumItem")
795               End If

800               If MsgBox("Voulez vous modifier l'inventaire pour la pièce " & sPiece & " ?", vbYesNo) = vbYes Then
805                 If cmbNoProjet.ItemData(cmbNoProjet.ListIndex) = 0 Then
810                   Call rstInventaire.Open("SELECT * FROM GRB_InventaireElec WHERE NoItem = '" & Replace(sPiece, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
815                 Else
820                   Call rstInventaire.Open("SELECT * FROM GRB_InventaireMec WHERE NoItem = '" & Replace(sPiece, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
825                 End If

830                 If rstInventaire.EOF Then
835                   Call rstInventaire.AddNew

840                   rstInventaire.Fields("NoItem") = sPiece
845                   rstInventaire.Fields("Description") = rstPiece.Fields("Desc_FR")
850                   rstInventaire.Fields("Manufacturier") = rstPiece.Fields("Manufact")

855                   Call frmChoixQteBoite.Afficher(rstPiece.Fields("NumItem"))

860                   rstInventaire.Fields("CommandeParBoite") = g_bQteBoite
865                   rstInventaire.Fields("QteBoite") = g_sQteBoite

870                   rstInventaire.Fields("QuantitéStock") = 0
875                   rstInventaire.Fields("Commentaires") = ""

880                   If cmbNoProjet.ItemData(cmbNoProjet.ListIndex) = 0 Then
885                     Call frmChoixLocalisation.Afficher(ELECTRIQUE, rstPiece.Fields("NumItem"))
890                   Else
895                     Call frmChoixLocalisation.Afficher(MECANIQUE, rstPiece.Fields("NumItem"))
900                   End If

905                   rstInventaire.Fields("Localisation") = g_sLocalisation
910                   rstInventaire.Fields("Minimum") = False
915                   rstInventaire.Fields("QuantitéMinimum") = ""
920                   rstInventaire.Fields("Commande") = ""
925                 End If

930                 If rstInventaire.Fields("CommandeParBoite") = True Then
935                   rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) + (CDbl(Replace(rstPiece.Fields("Qté"), "-", "")) * CDbl(rstInventaire.Fields("QteBoite"))), ".", ",")
940                 Else
945                   rstInventaire.Fields("QuantitéStock") = Replace(CDbl(rstInventaire.Fields("QuantitéStock")) + CDbl(Replace(rstPiece.Fields("Qté"), "-", "")), ".", ",")
950                 End If

955                 If rstPiece.Fields("Prix_List") = "" Then
960                   rstInventaire.Fields("Prix Liste") = " "
965                 Else
970                   rstInventaire.Fields("Prix Liste") = rstPiece.Fields("Prix_List")
975                 End If

980                 rstInventaire.Fields("Escompte") = rstPiece.Fields("Escompte")
985                 rstInventaire.Fields("Prix net") = rstPiece.Fields("Prix_Net")

990                 Call rstInventaire.Update

995                 Call rstInventaire.Close

1000                If cmbNoProjet.ItemData(cmbNoProjet.ListIndex) = 0 Then
1005                  Call rstInvModif.Open("SELECT * FROM GRB_InventaireElecModif", g_connData, adOpenDynamic, adLockOptimistic)
1010                Else
1015                  Call rstInvModif.Open("SELECT * FROM GRB_InventaireMecModif", g_connData, adOpenDynamic, adLockOptimistic)
1020                End If

1025                Call rstInvModif.AddNew

1030                rstInvModif.Fields("Date") = txtDateRetour.Text
1035                rstInvModif.Fields("IDProjet") = txtnoprojet.Text
1040                rstInvModif.Fields("NoItem") = sPiece

1045                rstInvModif.Fields("Quantité") = Replace(rstPiece.Fields("Qté"), "-", "")

1050                rstInvModif.Fields("User") = g_sInitiale

1055                Call rstInvModif.Update

1060                Call rstInvModif.Close
1065              End If
1070            End If
1075          End If

1080          Exit Do
1085        End If
               
1090        Call rstBC.MoveNext
1095      Loop

1100      Call rstPiece.MoveNext
1105    Loop

1110    Set rstInventaire = Nothing
1115    Set rstInvModif = Nothing
  
1120    Call rstPiece.Close
1125    Set rstPiece = Nothing
     
1130    Call rstBC.Close
1135    Set rstBC = Nothing

1140    If m_eTypeRetour = ACHAT Then
1145      Call RemplirListViewAchat
1150    Else
1155      Call RemplirListViewProjet
       
          'Ajout aux modifs
1160      Set rstModif = New ADODB.Recordset
1165      Set rstEmploye = New ADODB.Recordset
         
1170      Call rstModif.Open("SELECT * FROM GRB_Projet_Modif", g_connData, adOpenDynamic, adLockOptimistic)
      
1175      Call rstModif.AddNew

1180      Call rstEmploye.Open("SELECT noEmploye FROM GRB_Employés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)
       
1185      rstModif.Fields("Type") = "E"
1190      rstModif.Fields("IDProjet") = Right$(txtnoprojet.Text, Len(txtnoprojet.Text) - 1)
1195      rstModif.Fields("noEmployé") = rstEmploye.Fields("noEmploye")
1200      rstModif.Fields("Date") = ConvertDate(Date)
1205      rstModif.Fields("Heure") = Time
1210      rstModif.Fields("TypeModif") = "RETOUR"

1215      Call rstEmploye.Close
1220      Set rstEmploye = Nothing
      
1225      Call rstModif.Update
    
1230      Call rstModif.Close
1235      Set rstModif = Nothing
1240    End If
  
1245    Exit Sub

AfficherErreur:

1250    woups "frmRetourMarchandise", "Retour", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      txtDateRetour.Text = ConvertDate(Date)

15      Exit Sub

AfficherErreur:

20      woups "frmRetourMarchandise", "Form_Load", Err, Erl
End Sub

Private Sub lvwProjet_ItemCheck(ByVal Item As MSComctlLib.ListItem)

5       On Error GoTo AfficherErreur

10      If Item.Text <> vbNullString Then
15        If Item.Text > 0 Then
20          Item.Checked = False
25        End If
30      Else
35        Item.Checked = False
40      End If

45      Exit Sub

AfficherErreur:

50      woups "frmRetourMarchandise", "lvwProjet_ItemCheck", Err, Erl
End Sub

Private Sub mvwRetour_DateClick(ByVal DateClicked As Date)

5       On Error GoTo AfficherErreur

10      txtDateRetour.Text = ConvertDate(DateClicked)

        'Enlever le calendrier
15      mvwRetour.Visible = False

20      Exit Sub

AfficherErreur:

25      woups "frmRetourMarchandise", "mvwRetour_DateClick", Err, Erl
End Sub

Private Sub mvwRetour_LostFocus()

5       On Error GoTo AfficherErreur

10      mvwRetour.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmRetourMarchandise", "mvwRetour_LostFocus", Err, Erl
End Sub

Private Sub cmdDate_Click()

5       On Error GoTo AfficherErreur

        'Ouverture du calendrier
10      If txtDateRetour.Text <> vbNullString Then
15        mvwRetour.Value = txtDateRetour.Text
20      Else
25        mvwRetour.Value = Date
35      End If

40      mvwRetour.Visible = True

45      Call mvwRetour.SetFocus

50      Exit Sub

AfficherErreur:

55      woups "frmRetourMarchandise", "cmdDate_Click", Err, Erl
End Sub
