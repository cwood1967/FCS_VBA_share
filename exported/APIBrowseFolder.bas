Attribute VB_Name = "APIBrowseFolder"
Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const BIF_DONTGOBELOWDOMAIN As Long = &H2
Private Const BIF_RETURNFSANCESTORS As Long = &H8
Private Const BIF_BROWSEFORCOMPUTER As Long = &H1000
Private Const BIF_BROWSEFORPRINTER As Long = &H2000
Private Const BIF_BROWSEINCLUDEFILES As Long = &H4000
Private Const MAX_PATH As Long = 260

Type BrowseInfo
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszINSTRUCTIONS As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Declare Function SHGetPathFromIDListA Lib "shell32.dll" ( _
    ByVal pidl As Long, _
    ByVal pszBuffer As String) As Long

Declare Function SHBrowseForFolderA Lib "shell32.dll" ( _
    lpBrowseInfo As BrowseInfo) As Long

Public bUpdateEnable As Boolean

Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustomFilter As Long
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

Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_OVERWRITEPROMPT = &H2

Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (lpofn _
    As OPENFILENAME) As Long

Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (lpofn _
    As OPENFILENAME) As Long
    
Function BrowseFolder(Optional Caption As String = "") As String

    Dim BrowseInfo As BrowseInfo
    Dim FolderName As String
    Dim ID As Long
    Dim res As Long

    With BrowseInfo
        .hOwner = 0
        .pidlRoot = 0
        .pszDisplayName = String$(MAX_PATH, vbNullChar)
        .lpszINSTRUCTIONS = Caption
        .ulFlags = BIF_RETURNONLYFSDIRS
        .lpfn = 0
    End With

    FolderName = String(MAX_PATH, vbNullChar)

    ID = SHBrowseForFolderA(BrowseInfo)

    If ID Then
        res = SHGetPathFromIDListA(ID, FolderName)
        If res Then
            BrowseFolder = VBA.Left(FolderName, InStr(FolderName, vbNullChar) - 1)
        End If
        End If

End Function



