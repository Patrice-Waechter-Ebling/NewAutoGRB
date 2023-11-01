Attribute VB_Name = "noyau"
Option Explicit
Public Const MAX_PATH = 260
Public Const MAX_PRIORITY = 99
Public Const Titre As String = "NewAutoGRB"
Public Const RGN_OR = 2
Public Const LVM_FIRST As Long = &H1000
Public Const LVM_GETITEMRECT As Long = (LVM_FIRST + 14)
Public Const LVM_GETITEMPOSITION As Long = (LVM_FIRST + 16)
Public Const LVM_GETCOLUMNORDERARRAY As Long = (LVM_FIRST + 59)
Public Const LVM_GETCOLUMNWIDTH As Long = (LVM_FIRST + 29)
Public Const LVM_GETITEM As Long = (LVM_FIRST + 75)
Public Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 45)
Public Const LVIF_TEXT As Long = &H1
Public Const HWND_TOP As Long = 0
Public Const WM_GETFONT = &H31
Public Const WM_SETFONT = &H30
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const ULW_COLORKEY = &H1
Public Const ULW_ALPHA = &H2
Public Const ULW_OPAQUE = &H4
Public Const AC_SRC_OVER = &H0
Public Const AC_SRC_ALPHA = &H1
Public Const AC_SRC_NO_PREMULT_ALPHA = &H1
Public Const AC_SRC_NO_ALPHA = &H2
Public Const AC_DST_NO_PREMULT_ALPHA = &H10
Public Const AC_DST_NO_ALPHA = &H20
Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2
Public Const VER_NT_WORKSTATION As Byte = 1
Public Const VER_NT_DOMAIN_CONTROLLER As Byte = 2
Public Const VER_NT_SERVER As Byte = 3
Public Const VER_SUITE_PERSONAL As Integer = &H200
Public Const cbCSDVersion As Long = 128 * 2
Public Const STATUS_SUCCESS As Long = 0
Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const BIF_DONTGOBELOWDOMAIN = &H2
Public Const BIF_STATUSTEXT = &H4
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_EDITBOX = &H10
Public Const BIF_VALIDATE = &H20
Public Const BIF_NEWDIALOGSTYLE = &H40
Public Const BIF_BROWSEINCLUDEURLS = &H80
Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const BIF_BROWSEFORPRINTER = &H2000
Public Const BIF_BROWSEINCLUDEFILES = &H4000
Public Const BIF_SHAREABLE = &H8000
Public Const MAXDWORD = &HFFFF
Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const S_GROUPE_DEFAUT  As String = "Par défaut"
Public Const S_GROUPE_ADMIN   As String = "Administrateur"
Public Const S_PASS_DEFAUT    As String = "passe"
Public Const S_CHEMIN_DEFAUT  As String = "Data.mdb" 'chemin pour la db lors de la conception
Public Const COLOR_ORANGE     As Long = &H80FF&
Public Const COLOR_VERT       As Long = vbGreen
Public Const COLOR_MAGENTA    As Long = vbMagenta
Public Const COLOR_JAUNE      As Long = vbYellow
Public Const COLOR_GRIS       As Long = &H808080
Public Const COLOR_ROUGE      As Long = vbRed
Public Const COLOR_VERT_FORET As Long = &H4000&
Public Const COLOR_BLEU       As Long = vbBlue
Public Const COLOR_ROSE       As Long = &H8080FF
Public Const COLOR_BRUN       As Long = &H404080
Public Const COLOR_NOIR       As Long = vbBlack
Public Const COLOR_CYAN       As Long = vbCyan
Public Const SW_SHOWNORMAL = 1
Public Const LOCALE_SDECIMAL = &HE


Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSecurityDescriptorOwner Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pOwner As Any, ByVal lpbOwnerDefaulted As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer
Public Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetVersion Lib "kernel32" () As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Declare Function CreateDirectoryEx Lib "kernel32" Alias "CreateDirectoryExA" (ByVal lpTemplateDirectory As String, ByVal lpNewDirectory As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Public Declare Function FindEnvironmentString Lib "shell32.dll" Alias "FindEnvironmentStringA" (ByVal szEnvVar As String) As String
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendNotifyMessage Lib "user32" Alias "SendNotifyMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendDlgItemMessage Lib "user32" Alias "SendDlgItemMessageA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias " Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Public Declare Function WriteProfileSection Lib "kernel32" Alias "WriteProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, ByRef bWow64Process As Boolean) As Long
Public Declare Function RtlGetVersion Lib "ntdll" (lpVersionInformation As RTL_OSVERSIONINFOEXW) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Public Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Enum ListPaths
   PathsAndFilenames = 1
   FilenamesOnly = 2
   PathsOnly = 3
End Enum
Public Enum enumCatalogue
  ELECTRIQUE = 0
  MECANIQUE = 1
End Enum
Public Enum enumConvert
  MODE_PAS_FORMAT = 0
  MODE_DECIMAL = 1
  MODE_ARGENT = 2
  MODE_POURCENT = 3
End Enum

Public Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type
Public Type SHITEMID
    cb As Long
    abID() As Byte
End Type
Public Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type
Public Type ITEMIDLIST
    mkid As SHITEMID
End Type
Public Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long

End Type
Public Type RTL_OSVERSIONINFOEXW 'pour les windows >= 7
   dwOSVersionInfoSize As Long
   dwMajorVersion      As Long
   dwMinorVersion      As Long
   dwBuildNumber       As Long
   dwPlatformId        As Long
   szCSDVersion        As String * cbCSDVersion ' 128*2 =&H100 &HFF+ parité
   wServicePackMajor   As Integer
   wServicePackMinor   As Integer
   wSuiteMask          As Integer
   wProductType        As Byte
   wReserved           As Byte
End Type
Public Type ACL
        AclRevision As Byte
        Sbz1 As Byte
        AclSize As Integer
        AceCount As Integer
        Sbz2 As Integer
End Type
Public Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type
Public Type SECURITY_DESCRIPTOR
        Revision As Byte
        Sbz1 As Byte
        Control As Long
        Owner As Long
        Group As Long
        Sacl As ACL
        Dacl As ACL
End Type
Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type
Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
Public Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type
Public Type SIZE
    cx As Long
    cy As Long
End Type
Public Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type
Public Type RGB
    Bleu As Byte
    Vert As Byte
    Rouge As Byte
End Type

Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Type LVITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type

Public g_connData As New ADODB.Connection 'SQL
Public g_connMDB As ADODB.Connection      'Access
Public IdEmploye As Integer
Dim rs As New ADODB.Recordset
Dim fld As ADODB.Field
Dim recCount As Long
Dim i As Integer, X As Integer
Dim li As ListItem
Dim TheOS As OSVERSIONINFO
Dim OFName As OPENFILENAME
Dim ListSelected As ListPaths
Dim colInPaths As Collection
Dim colOutpaths As Collection
Dim sInputPath As String
Dim sOutputPath As String
Dim sInputPath2 As String
Dim sOutputPath2 As String
Dim lTotalProcess As Long
Dim SearchPath As String, FindStr As String
Dim FileSize As Long
Dim NumFiles As Integer, NumDirs As Integer
Dim AppStringName As String
Dim cTempCollection As Collection

'tables
Public Employes As New ListitemsView
Public Famille As ListitemsView

'AutoGRB
Public CheminBD                         As String  'Contient le chemin de la base de donnée
Public g_sUserID                        As String  'Contient le UserID de l'employé connecté
Public g_sEmploye                       As String  'Contient le nom de l'employé connecté
Public g_sInitiale                      As String  'Contient les initiales de l'employé connecté
Public g_bBonPasswd                     As Boolean 'Pour savoir si le mot de passe entré est bon
Public g_iNoGroupe                      As Integer 'No du groupe de l'employé connecté
Public BdMaintenant                     As Boolean 'Savoir si vielle bd ou récente
Public g_bAffichageClients              As Boolean
Public g_bAffichageFournisseurs         As Boolean
Public g_bAffichageContacts             As Boolean
Public g_bAffichageContactsVendeurs     As Boolean
Public g_bAffichageRapports             As Boolean
Public g_bAffichageEmployes             As Boolean
Public g_bAffichageCedule               As Boolean
Public g_bAffichageConfiguration        As Boolean
Public g_bAffichagePunch                As Boolean
Public g_bAffichageOutils               As Boolean
Public g_bAffichageInventaireMec        As Boolean
Public g_bAffichageCatalogueMec         As Boolean
Public g_bAffichageSoumissionsMec       As Boolean
Public g_bAffichageProjetsMec           As Boolean
Public g_bAffichageInventaireElec       As Boolean
Public g_bAffichageCatalogueElec        As Boolean
Public g_bAffichageSoumissionsElec      As Boolean
Public g_bAffichageProjetsElec          As Boolean
Public g_bAffichageAchats               As Boolean
Public g_bModificationClients           As Boolean
Public g_bModificationFournisseurs      As Boolean
Public g_bModificationContacts          As Boolean
Public g_bModificationEmployes          As Boolean
Public g_bModificationGroupes           As Boolean
Public g_bModificationFeuillesTemps     As Boolean
Public g_bModificationOutils            As Boolean
Public g_bModificationFacturation       As Boolean
Public g_bModificationBC                As Boolean
Public g_bModificationPunchEmployes     As Boolean
Public g_bModificationInventaireMec     As Boolean
Public g_bModificationCatalogueMec      As Boolean
Public g_bModificationSoumissionsMec    As Boolean
Public g_bModificationProjetsMec        As Boolean
Public g_bModificationInventaireElec    As Boolean
Public g_bModificationCatalogueElec     As Boolean
Public g_bModificationSoumissionsElec   As Boolean
Public g_bModificationProjetsElec       As Boolean
Public g_bSuppressionProjets            As Boolean
Public g_bModificationReception         As Boolean
Public g_bModificationRetourMarchandise As Boolean
Public g_bModificationListeDistribution As Boolean
Public g_bPunchSemaineAnterieure        As Boolean
Public g_bVerrouillageTempsProjet       As Boolean
Public g_bDeverrouillageTempsProjet     As Boolean
Public g_admin                          As Boolean
Public g_bCeduleOuverte                 As Boolean
Public g_sLocalisation                  As String
Public g_sQteBoite                      As String
Public g_bQteBoite                      As Boolean
Private m_iCmpApp                       As Integer 'doit rester private pas de protected

'Public g_objAchatElec                   As frmAchat
'Public g_objAchatMec                    As frmAchat
'----------------------------
'Fonctions du noyau supérieur
'----------------------------
Public Function StripNulls(OriginalStr As String) As String
   If (InStr(OriginalStr, Chr(0)) > 0) Then
       OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
   End If
   StripNulls = OriginalStr
End Function
Public Function TestOS()
    If IsHost64Bit() = True Then
        Conteneur.StatusBar1.Panels(6).Text = GetActualWindowsVersion + " 64bits"
    Else
        Conteneur.StatusBar1.Panels(6).Text = GetActualWindowsVersion
    End If
End Function
Public Function TestLocalInfo()
     Dim lSize   As Long
     Dim lLCID   As Long
     Dim sBuffer As String
     lLCID = GetUserDefaultLCID
     lSize = GetLocaleInfo(lLCID, LOCALE_SDECIMAL, StrPtr(sBuffer), lSize)
     sBuffer = Space$(lSize)
     lSize = GetLocaleInfo(lLCID, LOCALE_SDECIMAL, sBuffer, lSize)
     sBuffer = Trim$(Replace(sBuffer, Chr(0), ""))
     If sBuffer = "." Then
        MsgBox "Vos paramètres régionaux sont incorrects!" & vbCrLf & "Vous devez avoir la virgule (,) comme symbole de décimal!" & vbCrLf & "Des erreurs vont se produire si vous utilisez des formulaires contenants des montants d'argent!", vbOKOnly + vbInformation, Titre
    End If
End Function
Public Sub woups ByVal sSourceName As String, ByVal sMethode As String, ByVal Erreur As ErrObject, ByVal iNoLigne As Integer, Optional ByVal o_sParams As String)
    Dim rstErreur    As ADODB.Recordset
    Dim datNow       As Date
    Dim lNoErr       As Long
    Dim sDescription As String
    Dim sSource      As String
    datNow = Now
    lNoErr = Erreur.number
    sDescription = Erreur.Description
    sSource = Erreur.Source
    MsgBox "Une erreur est survenue!" + vbCrLf + vbCrLf + "Erreur numéro " + lNoErr + vbCrLf + sMethode + "@" + sSourceName + vbCrLf + "Description : " + sDescription, vbOKOnly + vbCritical, Titre
    Set rstErreur = New ADODB.Recordset
    rstErreur.Open "SELECT * FROM GRBErreurs", g_connData, adOpenDynamic, adLockOptimistic
    rstErreur.AddNew
    If g_sEmploye <> vbNullString Then
        rstErreur.Fields("Qui") = g_sEmploye
    End If
    rstErreur.Fields("Date") = ConvertDate(datNow)
    rstErreur.Fields("Heure") = Right$("0" & Hour(datNow), 2) & ":" & Right$("0" & Minute(datNow), 2) & ":" & Right$("0" & Second(datNow), 2)
    rstErreur.Fields("Form") = sSourceName
    rstErreur.Fields("Methode") = sMethode
    rstErreur.Fields("NoLigne") = iNoLigne
    rstErreur.Fields("NoErreur") = lNoErr
    rstErreur.Fields("Description") = sDescription
    rstErreur.Fields("Source") = sSource
    If Not IsMissing(o_sParams) Then
        rstErreur.Fields("Params") = o_sParams
    End If
    rstErreur.Update
    rstErreur.Close
    Set rstErreur = Nothing
End Sub
Public Function ComboContient(ByVal cmbSource As ComboBox, ByVal sRecherche As String) As Boolean
    On Error GoTo Oups
    Dim iCompteur As Integer
    ComboContient = False
    For iCompteur = 0 To cmbSource.ListCount - 1
        If UCase(Trim$(cmbSource.LIST(iCompteur))) = UCase(Trim$(sRecherche)) Then
            ComboContient = True
            Exit For
        End If
    Next
    Exit Function
Oups:
    woups "Noyau Secondaire", "ComboContient", Err, Err.number, Err.Description
End Function
Public Function IsDisplayHD() As Boolean
    Dim rc As RECT
    GetClientRect GetDesktopWindow(), rc
    If rc.Right > 1900 Then
        If rc.Bottom > 999 Then
            IsDisplayHD = True
        Else
            IsDisplayHD = False
        End If
        IsDisplayHD = False
    End If
End Function
Public Function ConvertDate(ByVal datDate As Date) As String
    On Error GoTo Oups
    ConvertDate = Year(datDate) & "-" & Right$("0" & Month(datDate), 2) & "-" & Right$("0" & Day(datDate), 2)
    Exit Function
Oups:
    woups "Noyau Secondaire", "ConvertDate", Err, Err.number, Err.Description
End Function
Public Function UneSeuleInstance() As Boolean
    If EnumWindows(AddressOf EnumWindowProc, &H0) > 1 Then
        MsgBox "Le programme est déjà ouvert!", vbCritical, Titre
        UneSeuleInstance = False
        Exit Function
    Else
    UneSeuleInstance = True
    End If
End Function
Public Function EnumWindowProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    On Error GoTo Oups
    Dim sTitle As String * 80
    If GetParent(hwnd) = 0& And IsWindowVisible(hwnd) Then
        GetWindowText Conteneur.hwnd, sTitle, Len(sTitle)
        If InStr(1, sTitle, App.Title) > 0 Then
            m_iCmpApp = m_iCmpApp + 1
        End If
    End If
    If m_iCmpApp > 1 Then
        EnumWindowProc = 0
    Else
        EnumWindowProc = 1
    End If
    Exit Function
Oups:
       woups "Noyau Secondaire", "EnumWindowProc", Err, Err.number, Err.Description
End Function

Public Function GetWindowsVersion(ByRef IsWin2000 As Boolean) As String
    Dim strCSDVersion As String
    TheOS.dwOSVersionInfoSize = Len(TheOS)
    GetVersionEx TheOS
    Select Case TheOS.dwPlatformId
    Case VER_PLATFORM_WIN32_WINDOWS
        If TheOS.dwMinorVersion >= 10 Then
            GetWindowsVersion = "Windows 98 version: "
        Else
            GetWindowsVersion = "Windows 95 version: "
        End If
    Case VER_PLATFORM_WIN32_NT
        GetWindowsVersion = "Windows NT version: "
    End Select
    If InStr(TheOS.szCSDVersion, Chr(0)) <> 0 Then
        strCSDVersion = ": " & Left(TheOS.szCSDVersion, InStr(TheOS.szCSDVersion, Chr(0)) - 1)
    Else
        strCSDVersion = ""
    End If
    GetWindowsVersion = GetWindowsVersion & TheOS.dwMajorVersion & "." & _
        TheOS.dwMinorVersion & " (Build " & TheOS.dwBuildNumber & strCSDVersion & ")"
    If TheOS.dwMajorVersion = 5 Then IsWin2000 = True Else IsWin2000 = False
    If TheOS.dwMajorVersion = 6 Then
        IsWin2000 = False
        If TheOS.dwMinorVersion = 2 Then GetWindowsVersion = "Windows 10"
        End If
End Function
Public Function TrouverDossierSpecial(pID As Integer) As String


End Function
Public Function ChargerFichier() As String
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = Conteneur.hwnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = "Base de données Acces (*.mdb)" + Chr$(0) + "*.mdb" + Chr$(0) + "Fichier de données Info (*.inf)" + Chr$(0) + "*.inf" + Chr$(0) + "Tous les fichiers (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrInitialDir = App.Path
    OFName.lpstrTitle = Titre + "- Charger un fichier"
    OFName.flags = 0
    If GetOpenFileName(OFName) Then
      ChargerFichier = Trim$(OFName.lpstrFile)
    Else
        ChargerFichier = "requete invalide"
    End If
End Function
Public Function SauverFichier() As String
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = Conteneur.hwnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = "Base de données Acces (*.mdb)" + Chr$(0) + "*.mdb" + Chr$(0) + "Fichier de données Info (*.inf)" + Chr$(0) + "*.inf" + Chr$(0) + "Tous les fichiers (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrInitialDir = App.Path
    OFName.lpstrTitle = Titre + "- Enregistrer un fichier"
    OFName.flags = 0
    If GetSaveFileName(OFName) Then
      SauverFichier = Trim$(OFName.lpstrFile)
    Else
        SauverFichier = "requete invalide"
    End If
End Function
Public Function NewGetTheOSVersion() As String
    TheOS.dwOSVersionInfoSize = Len(TheOS)
    If GetVersionEx(TheOS) = 1 Then
        Select Case TheOS.dwPlatformId
            Case VER_PLATFORM_WIN32s
                NewGetTheOSVersion = "Win32s sur Windows 3.1x"
            Case VER_PLATFORM_WIN32_NT
                NewGetTheOSVersion = "Windows NT"
                
                Select Case TheOS.dwMajorVersion
                    Case 3
                        NewGetTheOSVersion = "Windows NT 3.5"
                    Case 4
                        NewGetTheOSVersion = "Windows NT 4.0"
                    Case 5
                        Select Case TheOS.dwMinorVersion
                            Case 0
                                NewGetTheOSVersion = "Windows 2000"
                            Case 1
                                NewGetTheOSVersion = "Windows XP"
                            Case 2
                                NewGetTheOSVersion = "Windows Server 2003"
                        End Select
                    Case 6
                        Select Case TheOS.dwMinorVersion
                            Case 0
                                NewGetTheOSVersion = "Windows Vista/Server 2008"
                            Case 1
                                NewGetTheOSVersion = "Windows 7/Server 2008 R2"
                            Case 2
                                NewGetTheOSVersion = "Windows 8/Server 2012"
                            Case 3
                                NewGetTheOSVersion = "Windows 8.1/Server 2012 R2"
                        End Select
                End Select
        
            Case VER_PLATFORM_WIN32_WINDOWS:
                Select Case TheOS.dwMinorVersion
                    Case 0
                        NewGetTheOSVersion = "Windows 95"
                    Case 90
                        NewGetTheOSVersion = "Windows Me"
                    Case Else
                        NewGetTheOSVersion = "Windows 98"
                End Select
        End Select
    Else
        NewGetTheOSVersion = TheOS.dwMajorVersion & "." & TheOS.dwMinorVersion & "." & TheOS.dwBuildNumber
    End If
End Function
Public Function IsLayeredWindow(ByVal hwnd As Long) As Boolean
    Dim WinInfo As Long
    WinInfo = GetWindowLong(hwnd, GWL_EXSTYLE)
    If (WinInfo And WS_EX_LAYERED) = WS_EX_LAYERED Then
        IsLayeredWindow = True
    Else
        IsLayeredWindow = False
    End If
End Function
Public Sub SetLayeredWindow(ByVal hwnd As Long, ByVal bIsLayered As Boolean)
    Dim WinInfo As Long
    WinInfo = GetWindowLong(hwnd, GWL_EXSTYLE)
    If bIsLayered = True Then
        WinInfo = WinInfo Or WS_EX_LAYERED
    Else
        WinInfo = WinInfo And Not WS_EX_LAYERED
    End If
    SetWindowLong hwnd, GWL_EXSTYLE, WinInfo
End Sub
Sub LoadListViewFromRecordset(LV As ListView, rs As ADODB.Recordset, Optional MaxRecords As Long)
 On Error Resume Next
   Dim fld As ADODB.Field, alignment As Integer
    Dim recCount As Long, i As Long, fldName As String
    Dim li As ListItem
    LV.ListItems.Clear
    LV.ColumnHeaders.Clear
    For Each fld In rs.Fields
        Select Case fld.Type
            Case adBoolean, adCurrency, adDate, adDecimal, adDouble
                alignment = lvwColumnRight
            Case adInteger, adNumeric, adSingle, adSmallInt, adVarNumeric
                alignment = lvwColumnRight
            Case adBSTR, adChar, adVarChar, adVariant
                alignment = lvwColumnLeft
            Case Else
                alignment = 0
        End Select
        If alignment <> -1 Then
            If LV.ColumnHeaders.count = 0 Then alignment = lvwColumnLeft
            LV.ColumnHeaders.Add , , fld.Name, fld.DefinedSize * 200, alignment
        End If
    Next
    If LV.ColumnHeaders.count = 0 Then Exit Sub
    rs.MoveFirst
    Do Until rs.EOF
        recCount = recCount + 1
        fldName = LV.ColumnHeaders(1).Text
        Set li = LV.ListItems.Add(, , rs.Fields(fldName) & "")
        For i = 2 To LV.ColumnHeaders.count
            fldName = LV.ColumnHeaders(i)
            li.ListSubItems.Add , , rs.Fields(fldName) & ""
        Next
        If recCount = MaxRecords Then Exit Do
        rs.MoveNext
    Loop
End Sub
Sub LoadComboFromRecordset(LV As ComboBox, rs As ADODB.Recordset, Optional MaxRecords As Long)
On Error Resume Next
    Dim fld As ADODB.Field, alignment As Integer
    Dim recCount As Long, i As Long, fldName As String
    LV.Clear
    rs.MoveFirst
    Do Until rs.EOF
        recCount = recCount + 1
        LV.AddItem rs.Fields(1)
        If recCount = MaxRecords Then Exit Do
        rs.MoveNext
    Loop
End Sub
Sub ListViewAdjustColumnWidth(LV As ListView, Optional AccountForHeaders As Boolean)
#If USE_API Then
    Dim col As Integer, lParam As Long
    If AccountForHeaders Then
        lParam = LVSCW_AUTOSIZE_USEHEADER
    Else
        lParam = LVSCW_AUTOSIZE
    End If
    For col = 1 To LV.ColumnHeaders.count
        SendMessage LV.hwnd, LVM_SETCOLUMNWIDTH, col, lParam
    Next
#Else
    Dim row As Long, col As Long
    Dim width As Single, maxWidth As Single
    Dim saveFont As StdFont, saveScaleMode As Integer
    Dim cellText As String
    If LV.ListItems.count = 0 Then Exit Sub
    Set saveFont = LV.Parent.Font
    Set LV.Parent.Font = LV.Font
    saveScaleMode = LV.Parent.ScaleMode
    LV.Parent.ScaleMode = vbTwips
    For col = 1 To LV.ColumnHeaders.count
        maxWidth = 0
        If AccountForHeaders Then
            maxWidth = LV.Parent.TextWidth(LV.ColumnHeaders(col).Text) + 200
        End If
        For row = 1 To LV.ListItems.count
            If col = 1 Then
                cellText = LV.ListItems(row).Text
            Else
                cellText = LV.ListItems(row).ListSubItems(col - 1).Text
            End If
            width = LV.Parent.TextWidth(cellText) + 200
            If width > maxWidth Then maxWidth = width
        Next
        LV.ColumnHeaders(col).width = maxWidth
    Next
    Set LV.Parent.Font = saveFont
    LV.Parent.ScaleMode = saveScaleMode
#End If
End Sub
Sub ListViewSortOnNonStringField(LV As ListView, ByVal ColumnIndex As Integer, Optional SortOrder As ListSortOrderConstants, Optional IsDateValue As Boolean)
    Dim li As ListItem, number As Double, newIndex As Integer
    Dim minValue As Double
    LV.Visible = False
    LV.Sorted = False
    LV.ColumnHeaders.Add , , "dummy column", 1000
    newIndex = LV.ColumnHeaders.count - 1
    For Each li In LV.ListItems
        If IsDateValue Then
            number = DateValue(li.ListSubItems(ColumnIndex - 1))
        Else
            number = CDbl(li.ListSubItems(ColumnIndex - 1))
        End If
        If number < minValue Then minValue = number
        li.ListSubItems.Add , , Format$(number, "000000000000000.000")
    Next
    If minValue < 0 Then
        For Each li In LV.ListItems
            number = CDbl(li.ListSubItems(newIndex)) - minValue
            li.ListSubItems(newIndex).Text = Format$(number, "000000000000000.000")
        Next
    End If
        LV.SortKey = newIndex
    LV.SortOrder = SortOrder
    LV.Sorted = True
    LV.ColumnHeaders.Remove newIndex + 1
    For Each li In LV.ListItems
        li.ListSubItems.Remove newIndex
    Next
    LV.Visible = True
End Sub
Public Function AquerirEmployes()
    Employes.Table = "GrbEmploye"
End Function
Public Function AquerirFamille()
    Famille.Table = "GrbFamille"
End Function
Private Sub ActiverBoutons(ByVal bEnabled As Boolean)
    On Error GoTo Oups
    Conteneur.MenuPrincipal.Buttons(10).Enabled = bEnabled
    Conteneur.MenuPrincipal.Buttons(12).Enabled = bEnabled
    Conteneur.MenuPrincipal.Buttons(7).Enabled = bEnabled
    Conteneur.MenuPrincipal.Buttons(1).Enabled = bEnabled
    Conteneur.MenuPrincipal.Buttons(13).Enabled = bEnabled
    Conteneur.MenuPrincipal.Buttons(3).Enabled = bEnabled
    Conteneur.MenuPrincipal.Buttons(4).Enabled = bEnabled
    Conteneur.MenuPrincipal.Buttons(2).Enabled = bEnabled
    Conteneur.MenuPrincipal.Buttons(11).Enabled = bEnabled
    Conteneur.MenuPrincipal.Buttons(13).Enabled = bEnabled
    Conteneur.MenuPrincipal.Buttons(6).Enabled = bEnabled
    Conteneur.MenuPrincipal.Buttons(8).Enabled = bEnabled
    Exit Sub
Oups:
    woups "frmDispatch", "ActiverBoutons", Err, Err.number, Err.Description
End Sub
Private Function ActiverBoutonsGroupe()
    On Error GoTo Oups
    Conteneur.MenuPrincipal.Buttons(1).Enabled = g_bAffichageClients
    Conteneur.MenuPrincipal.Buttons(2).Enabled = g_bAffichageFournisseurs
    Conteneur.MenuPrincipal.Buttons(3).Enabled = g_bAffichageContacts
    Conteneur.MenuPrincipal.Buttons(5).Enabled = g_bAffichageContactsVendeurs
    Conteneur.MenuPrincipal.Buttons(8).Enabled = g_bAffichageRapports
    Conteneur.MenuPrincipal.Buttons(4).Enabled = g_bAffichageEmployes
    Conteneur.MenuPrincipal.Buttons(7).Enabled = g_bAffichageCedule
    Conteneur.MenuPrincipal.Buttons(14).Enabled = g_bAffichageConfiguration
    Conteneur.MenuPrincipal.Buttons(9).Enabled = g_bModificationListeDistribution
    If g_bAffichagePunch = True Or g_bModificationFeuillesTemps = True Or g_bModificationFacturation = True Then
        Conteneur.MenuPrincipal.Buttons(6).Enabled = True
    Else
        Conteneur.MenuPrincipal.Buttons(6).Enabled = False
    End If
    If g_bAffichageSoumissionsMec = True Or g_bAffichageSoumissionsElec = True Or g_bAffichageProjetsMec = True Or g_bAffichageProjetsElec = True Then
        Conteneur.MenuPrincipal.Buttons(13).Enabled = True
    Else
        Conteneur.MenuPrincipal.Buttons(5).Enabled = False
    End If
    If g_bAffichageCatalogueMec = True Or g_bAffichageCatalogueElec = True Then
        Conteneur.MenuPrincipal.Buttons(12).Enabled = True
    Else
        Conteneur.MenuPrincipal.Buttons(12).Enabled = False
    End If
    If g_bAffichageInventaireMec = True Or g_bAffichageInventaireElec = True Or g_bAffichageOutils = True Then
        Conteneur.MenuPrincipal.Buttons(11).Enabled = True
    Else
        Conteneur.MenuPrincipal.Buttons(11).Enabled = False
    End If
    Exit Function
Oups:
    woups "frmDispatch", "ActiverBoutonsGroupe", Err, Err.number, Err.Description
End Function

Public Function TesterVersion() As Boolean
    Dim sVersion  As String
    Dim rstConfig As ADODB.Recordset
    ActiverBoutonsGroupe
    g_sEmploye = Conteneur.StatusBar1.Panels(2).Text
    sVersion = App.Major & "." & Right$("0" & App.Minor, 2) & "." & Right$("0" & App.Revision, 4)
    Conteneur.lblVersion.Caption = "Version " & sVersion
    Set rstConfig = New ADODB.Recordset
    rstConfig.Open "SELECT DerniereVersion FROM GrbConfig", g_connData, adOpenDynamic, adLockOptimistic 'on veut l'info depuis le serveur SQL
    If Not IsNull(rstConfig.Fields("DerniereVersion")) Then
        If rstConfig.Fields("DerniereVersion") <> "" Then
            Conteneur.lblDerniereVersion.Caption = "Dernière Version : " & rstConfig.Fields("DerniereVersion")
        Else
            Conteneur.lblDerniereVersion.Caption = ""
        End If
    Else
       Conteneur.lblDerniereVersion.Caption = ""
    End If
    rstConfig.Close
    Set rstConfig = Nothing
    If Trim$(Replace(Conteneur.lblDerniereVersion.Caption, "Dernière Version : ", "")) = Trim$(Replace(Conteneur.lblVersion.Caption, "Version", "")) Then
        Conteneur.lblVersion.ForeColor = vbGreen
    Else
        Conteneur.lblVersion.ForeColor = vbRed
    End If
End Function
Public Function Login(username As String, password As String)
    Dim count As Integer
    TestOS
    g_sEmploye = Employes.ListView1.ListItems(frmLogin.Combo1.ListIndex + 1).SubItems(3)
    Conteneur.StatusBar1.Panels(2).Text = g_sEmploye
    Conteneur.StatusBar1.Panels(3).Text = Employes.ListView1.ListItems(frmLogin.Combo1.ListIndex + 1).SubItems(1)
    IdEmploye = CInt(Employes.ListView1.ListItems(frmLogin.Combo1.ListIndex + 1).Text)
    Conteneur.StatusBar1.Panels(4).Text = IdEmploye
    Conteneur.StatusBar1.Panels(5).Text = Employes.ListView1.ListItems(frmLogin.Combo1.ListIndex + 1).SubItems(9)
    Conteneur.Caption = Titre + " Solution GRB inc. (" & g_sEmploye & ")"
    frmVide.Show
    'utilisateur actif dans le menu fast connect
    If (Employes.ListView1.ListItems(frmLogin.Combo1.ListIndex).SubItems(8) = True) Then
        
        'Conteneur.Menu.Visible = True
    Else
        'Conteneur.Menu.Visible = False
    End If
    If username = Employes.ListView1.ListItems(frmLogin.Combo1.ListIndex).SubItems(3) Then
        If password = Employes.ListView1.ListItems(frmLogin.Combo1.ListIndex).SubItems(2) Then
            MsgBox "Bienvenue " + Employes.ListView1.ListItems(frmLogin.Combo1.ListIndex).SubItems(3), , Titre
        Else
            MsgBox "Identification invalide " + vbCrLf + password + " = " + Employes.ListView1.ListItems(frmLogin.Combo1.ListIndex + 1).SubItems(1), vbCritical + vbOKOnly, Titre
            count = count + 1
            If count = 3 Then
                MsgBox "3 échecs de connexion" + vbCrLf + "Abandon du programme", vbCritical, Titre
            End If
        End If
    Else
    Debug.Print Employes.ListView1.ListItems(frmLogin.Combo1.ListIndex + 1).SubItems(3) + "=" + Employes.ListView1.ListItems(frmLogin.Combo1.ListIndex + 1).SubItems(2) + "=" + Employes.ListView1.ListItems(frmLogin.Combo1.ListIndex).SubItems(1)
    End If
End Function
Sub Main()
    g_connData.Open "Driver={SQL Server};Server=Intranet;Database=GRB2023;uid=sa;pwd=Password01$"
    If IsDisplayHD = False Then
     If MsgBox("Ce programme est concu pour un affichage full HD" + vbCrLf + "Des déformations visuelles peuvent se produire" + vbCrLf + vbCrLf + "Voulez poursuivre néanoins?", vbQuestion + vbYesNo + vbDefaultButton2, Titre) = vbNo Then End
    End If
    Dim Employes As New ListitemsView
    frmLogin.Show
End Sub
Public Function IsHost64Bit() As Boolean
    Dim handle As Long
    Dim is64Bit As Boolean
    is64Bit = False
    handle = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")
    If handle <> 0 Then
        IsWow64Process GetCurrentProcess(), is64Bit
    End If
    IsHost64Bit = is64Bit
End Function

Public Function GetActualWindowsVersion() As String
   Dim ver As RTL_OSVERSIONINFOEXW
   ver.dwOSVersionInfoSize = Len(ver)
   If (RtlGetVersion(ver) <> STATUS_SUCCESS) Then
      GetActualWindowsVersion = "Failed to retrieve Windows version"
   End If
   Debug.Assert ver.dwPlatformId = VER_PLATFORM_WIN32_NT
   GetActualWindowsVersion = GetWinVerName(ver) & " " & GetWinSPVerNumber(ver) & " (v" & GetWinVerNumber(ver) & ")"
End Function
Public Function IsWinServerVersion(ByRef ver As RTL_OSVERSIONINFOEXW) As Boolean
   Debug.Assert ver.wProductType = VER_NT_WORKSTATION Or ver.wProductType = VER_NT_DOMAIN_CONTROLLER Or ver.wProductType = VER_NT_SERVER
   IsWinServerVersion = (ver.wProductType <> VER_NT_WORKSTATION)
End Function
Public Function GetWinVerNumber(ByRef ver As RTL_OSVERSIONINFOEXW) As String
   Debug.Assert ver.dwPlatformId = VER_PLATFORM_WIN32_NT
   GetWinVerNumber = ver.dwMajorVersion & "." & ver.dwMinorVersion & "." & ver.dwBuildNumber
End Function
Public Function GetWinSPVerNumber(ByRef ver As RTL_OSVERSIONINFOEXW) As String
   Debug.Assert ver.dwPlatformId = VER_PLATFORM_WIN32_NT
   If (ver.wServicePackMajor > 0) Then
      If (ver.wServicePackMinor > 0) Then
         GetWinSPVerNumber = "SP" & CStr(ver.wServicePackMajor) & "." & CStr(ver.wServicePackMinor)
         Exit Function
      Else
         GetWinSPVerNumber = "SP" & CStr(ver.wServicePackMajor)
         Exit Function
      End If
   End If
End Function
Private Function GetWinVerName(ByRef ver As RTL_OSVERSIONINFOEXW) As String
   Debug.Assert ver.dwPlatformId = VER_PLATFORM_WIN32_NT
   Select Case ver.dwMajorVersion
      Case 3
         If IsWinServerVersion(ver) Then
            GetWinVerName = "Windows NT 3.5 Server"
            Exit Function
         Else
            GetWinVerName = "Windows NT 3.5 Workstation"
            Exit Function
         End If
      Case 4
         If IsWinServerVersion(ver) Then
            GetWinVerName = "Windows NT 4.0 Server"
            Exit Function
         Else
            GetWinVerName = "Windows NT 4.0 Workstation"
            Exit Function
         End If
      Case 5
         Select Case ver.dwMinorVersion
            Case 0
               If IsWinServerVersion(ver) Then
                  GetWinVerName = "Windows 2000 Server"
                  Exit Function
               Else
                  GetWinVerName = "Windows 2000 Workstation"
                  Exit Function
               End If
            Case 1
               If (ver.wSuiteMask And VER_SUITE_PERSONAL) Then
                  GetWinVerName = "Windows XP Home Edition"
                  Exit Function
               Else
                  GetWinVerName = "Windows XP Professional"
                  Exit Function
               End If
            Case 2
               If IsWinServerVersion(ver) Then
                  GetWinVerName = "Windows Server 2003"
                  Exit Function
               Else
                  GetWinVerName = "Windows XP 64-bit Edition"
                  Exit Function
               End If
            Case Else
               Debug.Assert False
         End Select
      Case 6
         Select Case ver.dwMinorVersion
            Case 0
               If IsWinServerVersion(ver) Then
                  GetWinVerName = "Windows Server 2008"
                  Exit Function
               Else
                  GetWinVerName = "Windows Vista"
                  Exit Function
               End If
            Case 1
               If IsWinServerVersion(ver) Then
                  GetWinVerName = "Windows Server 2008 R2"
                  Exit Function
               Else
                  GetWinVerName = "Windows 7"
                  Exit Function
               End If
            Case 2
               If IsWinServerVersion(ver) Then
                  GetWinVerName = "Windows Server 2012"
                  Exit Function
               Else
                  GetWinVerName = "Windows 8"
                  Exit Function
               End If
            Case 3
               If IsWinServerVersion(ver) Then
                  GetWinVerName = "Windows Server 2012 R2"
                  Exit Function
               Else
                  GetWinVerName = "Windows 8.1"
                  Exit Function
               End If
            Case Else
               Debug.Assert False
         End Select
      Case 10
         If IsWinServerVersion(ver) Then
            GetWinVerName = "Windows Server 2016"
            Exit Function
         Else
            GetWinVerName = "Windows 10"
            Exit Function
         End If
      Case Else
         Debug.Assert False
   End Select
   GetWinVerName = "ID non listé"
End Function
Public Function EcrireTableDansFichier(Cible As String, LV As ListView) As Boolean
    On Error GoTo Oups
    Open "Cible" For Output As #1
'    Print #1, Texte
    Close 1
    EcrireTableDansFichier = True
Oups:
    EcrireTableDansFichier = False
End Function
Private Sub EcrireDansFichier(Cible As String, Texte As String)
    Open "Cible" For Output As #1
    Print #1, Texte
    Close 1
End Sub
Public Function ExplorerDossier() As String
  Dim BI As BROWSEINFO
  Dim nFolder As Long
  Dim IDL As ITEMIDLIST
  Dim pidl As Long
  Dim sPath As String, txtDisplayName As String
  Dim SHFI As SHFILEINFO
  With BI
    .hOwner = Conteneur.hwnd
    .pszDisplayName = String$(MAX_PATH, 0)
    .lpszTitle = "Explorer l'arborescence"
    'limitations d'exploration celui ci est full paquetage
    .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or BIF_STATUSTEXT Or BIF_RETURNFSANCESTORS Or BIF_EDITBOX Or BIF_VALIDATE Or BIF_NEWDIALOGSTYLE Or BIF_BROWSEINCLUDEURLS Or BIF_BROWSEFORCOMPUTER Or BIF_BROWSEFORPRINTER Or BIF_BROWSEINCLUDEFILES Or BIF_SHAREABLE
  End With
  ExplorerDossier = ""
  txtDisplayName = ""
  pidl = SHBrowseForFolder(BI)
  If pidl = 0 Then Exit Function
  sPath = String$(MAX_PATH, 0)
  SHGetPathFromIDList ByVal pidl, ByVal sPath
  ExplorerDossier = Left(sPath, InStr(sPath, vbNullChar) - 1)
  txtDisplayName = Left$(BI.pszDisplayName, InStr(BI.pszDisplayName, vbNullChar) - 1)
  CoTaskMemFree pidl
End Function
Public Function OuvrirConnectionMDB(FichierDB As String) As Boolean
      Dim sdsn As String
      If g_connMDB Is Nothing Then
      Set g_connMDB = New ADODB.Connection 'Access 2000/XP
        sdsn = "Provider=Microsoft.Jet.OLEDB.4.0;User ID = Admin;Data Source=" & FichierDB & ";Persist Security Info=False"
        g_connMDB.Open sdsn
        OuvrirConnectionMDB = True
    Else
        MsgBox "La base de donnée est introuvable à l'adresse:" & vbCrLf & FichierDB, vbOKOnly, Titre
        OuvrirConnectionMDB = False
    End If
End Function
Public Sub FermerConnectionMDB()
    If Not g_connMDB Is Nothing Then
        g_connMDB.Close
        Set g_connMDB = Nothing
    End If
End Sub
Public Function GetFirstDay(ByVal datDate As Date) As Date
    On Error GoTo Oups
    Dim iNoJour As Integer
    iNoJour = Weekday(datDate)
    Do While iNoJour > 1
        iNoJour = iNoJour - 1
        datDate = datDate - TimeSerial(24, 0, 0)
    Loop
    GetFirstDay = datDate
    Exit Function
Oups:
    woups "Noyau Secondaire", "GetFirstDay", Err, Err.number, Err.Description
End Function
Public Function GetLastDay(ByVal datDate As Date) As Date
    On Error GoTo Oups
    Dim iNoJour As Integer
    iNoJour = Weekday(datDate)
    Do While iNoJour < 7
        iNoJour = iNoJour + 1
        datDate = datDate + TimeSerial(24, 0, 0)
    Loop
    GetLastDay = datDate
    Exit Function
Oups:
    woups "Noyau Secondaire", "GetLastDay", Err, Err.number, Err.Description
End Function
Public Function GetDateTexte(ByVal datDate As Date) As String
    On Error GoTo Oups
    Dim sMonth As String
    sMonth = MonthName(Month(datDate))
    GetDateTexte = Day(datDate) & " " & sMonth & " " & Year(datDate)
    Exit Function
Oups:
    woups "Noyau Secondaire", "GetDateTexte", Err, Err.number, Err.Description
End Function
Public Function ValiderFormatNumeroProjSoum(ByVal sNoProjSoum As String) As Boolean
    On Error GoTo Oups
    Dim bNoValide  As Boolean
    Dim sErreurMsg As String
    bNoValide = True
    If Len(sNoProjSoum) <> 9 Then
        bNoValide = False
        sErreurMsg = "Le numéro doit contenir 9 caractères!"
    End If
    If bNoValide = True Then
        If UCase(Left$(sNoProjSoum, 1)) <> "M" And UCase(Left$(sNoProjSoum, 1)) <> "E" Then
            bNoValide = False
            sErreurMsg = "Le numéro doit commencé par : " & vbCrLf & vbCrLf & "      E pour les soumissions et projets électriques" & vbCrLf & "      M pour les soumissions et projets mécaniques"
        End If
    End If
    If bNoValide = True Then
        If Not IsNumeric(Mid$(sNoProjSoum, 2, 5)) Then
            bNoValide = False
            sErreurMsg = "Format invalide !"
        End If
    End If
    If bNoValide = True Then
        If Not IsNumeric(Right$(sNoProjSoum, 2)) Then
            bNoValide = False
            sErreurMsg = "Format invalide !"
        End If
    End If
    If bNoValide = True Then
        If Mid$(sNoProjSoum, 7, 1) <> "-" Then
            bNoValide = False
            sErreurMsg = "Format invalide !"
        End If
    End If
    If bNoValide = True Then
        If Mid$(sNoProjSoum, 3, 1) = 0 Then
            bNoValide = False
            sErreurMsg = "Le 3e caractère ne peut pas être '0' !"
        End If
    End If
    If bNoValide = True Then
        If Right$(sNoProjSoum, 2) = "99" Or Right$(sNoProjSoum, 2) = "00" Then
            bNoValide = False
            sErreurMsg = "L'extension doit être comprise entre 01 et 98"
        End If
    End If
    If bNoValide = False Then
        MsgBox sErreurMsg, vbOKOnly + vbExclamation, Titre
    End If
    ValiderFormatNumeroProjSoum = bNoValide
    Exit Function
Oups:
    woups "Noyau Secondaire", "ValiderFormatNumeroProjSoum", Err, Err.number, Err.Description
End Function


