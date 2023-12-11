Attribute VB_Name = "Module1"
'---------------------------------------------------------------------------------------
' Module    : Module1
' Author    : beededea
' Date      : 27/04/2023
' Purpose   : Module for declaring any public and private constants, APIs and types used by the functions therein.
'---------------------------------------------------------------------------------------

Option Explicit

'------------------------------------------------------ STARTS
'constants used to choose a font via the system dialog window
Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40
Private Const GHND As Long = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Private Const LF_FACESIZE As Integer = 32
Private Const CF_INITTOLOGFONTSTRUCT  As Long = &H40&
Private Const CF_SCREENFONTS As Long = &H1

'type declaration used to choose a font via the system dialog window
Private Type FormFontInfo
  Name As String
  Weight As Integer
  Height As Integer
  UnderLine As Boolean
  Italic As Boolean
  Color As Long
End Type

Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type FONTSTRUC
  lStructSize As Long
  hwnd As Long
  hdc As Long
  lpLogFont As Long
  iPointSize As Long
  Flags As Long
  rgbColors As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
  hInstance As Long
  lpszStyle As String
  nFontType As Integer
  MISSING_ALIGNMENT As Integer
  nSizeMin As Long
  nSizeMax As Long
End Type

Private Type ChooseColorStruct
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
'APIs used to choose a font via the system dialog window
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As FONTSTRUC) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'------------------------------------------------------ ENDS



'------------------------------------------------------ STARTS
' API and enums for acquiring the special folder paths
Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long

Public Enum FolderEnum ' has to be public
    feCDBurnArea = 59 ' \Docs & Settings\User\Local Settings\Application Data\Microsoft\CD Burning
    feCommonAppData = 35 ' \Docs & Settings\All Users\Application Data
    feCommonAdminTools = 47 ' \Docs & Settings\All Users\Start Menu\Programs\Administrative Tools
    feCommonDesktop = 25 ' \Docs & Settings\All Users\Desktop
    feCommonDocs = 46 ' \Docs & Settings\All Users\Documents
    feCommonPics = 54 ' \Docs & Settings\All Users\Documents\Pictures
    feCommonMusic = 53 ' \Docs & Settings\All Users\Documents\Music
    feCommonStartMenu = 22 ' \Docs & Settings\All Users\Start Menu
    feCommonStartMenuPrograms = 23 ' \Docs & Settings\All Users\Start Menu\Programs
    feCommonTemplates = 45 ' \Docs & Settings\All Users\Templates
    feCommonVideos = 55 ' \Docs & Settings\All Users\Documents\My Videos
    feLocalAppData = 28 ' \Docs & Settings\User\Local Settings\Application Data
    feLocalCDBurning = 59 ' \Docs & Settings\User\Local Settings\Application Data\Microsoft\CD Burning
    feLocalHistory = 34 ' \Docs & Settings\User\Local Settings\History
    feLocalTempInternetFiles = 32 ' \Docs & Settings\User\Local Settings\Temporary Internet Files
    feProgramFiles = 38 ' \Program Files
    feProgramFilesCommon = 43 ' \Program Files\Common Files
    'feRecycleBin = 10 ' ???
    feUser = 40 ' \Docs & Settings\User
    feUserAdminTools = 48 ' \Docs & Settings\User\Start Menu\Programs\Administrative Tools
    feUserAppData = 26 ' \Docs & Settings\User\Application Data
    feUserCache = 32 ' \Docs & Settings\User\Local Settings\Temporary Internet Files
    feUserCookies = 33 ' \Docs & Settings\User\Cookies
    feUserDesktop = 16 ' \Docs & Settings\User\Desktop
    feUserDocs = 5 ' \Docs & Settings\User\My Documents
    feUserFavorites = 6 ' \Docs & Settings\User\Favorites
    feUserMusic = 13 ' \Docs & Settings\User\My Documents\My Music
    feUserNetHood = 19 ' \Docs & Settings\User\NetHood
    feUserPics = 39 ' \Docs & Settings\User\My Documents\My Pictures
    feUserPrintHood = 27 ' \Docs & Settings\User\PrintHood
    feUserRecent = 8 ' \Docs & Settings\User\Recent
    feUserSendTo = 9 ' \Docs & Settings\User\SendTo
    feUserStartMenu = 11 ' \Docs & Settings\User\Start Menu
    feUserStartMenuPrograms = 2 ' \Docs & Settings\User\Start Menu\Programs
    feUserStartup = 7 ' \Docs & Settings\User\Start Menu\Programs\Startup
    feUserTemplates = 21 ' \Docs & Settings\User\Templates
    feUserVideos = 14  ' \Docs & Settings\User\My Documents\My Videos
    feWindows = 36 ' \Windows
    feWindowFonts = 20 ' \Windows\Fonts
    feWindowsResources = 56 ' \Windows\Resources
    feWindowsSystem = 37 ' \Windows\System32
End Enum
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
' APIs for useful functions START
Public Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
' APIs for useful functions END
'------------------------------------------------------ ENDS

'------------------------------------------------------ STARTS
' Constants and APIs for playing sounds
Public Const SND_ASYNC As Long = &H1             '  play asynchronously
Public Const SND_FILENAME  As Long = &H20000     '  name is a file name

Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
'API Functions to read/write information from INI File
Private Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any _
    , ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long _
    , ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any _
    , ByVal lpString As Any, ByVal lpFileName As String) As Long
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
'constants and APIs defined for querying the registry
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Public Const HKEY_CURRENT_USER As Long = &H80000001
Private Const REG_SZ  As Long = 1                          ' Unicode nul terminated string

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Any, ByVal cbData As Long) As Long
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
' Enums defined for opening a common dialog box to select files without OCX dependencies
Private Enum FileOpenConstants
    'ShowOpen, ShowSave constants.
    cdlOFNAllowMultiselect = &H200&
    cdlOFNCreatePrompt = &H2000&
    cdlOFNExplorer = &H80000
    cdlOFNExtensionDifferent = &H400&
    cdlOFNFileMustExist = &H1000&
    cdlOFNHideReadOnly = &H4&
    cdlOFNLongNames = &H200000
    cdlOFNNoChangeDir = &H8&
    cdlOFNNoDereferenceLinks = &H100000
    cdlOFNNoLongNames = &H40000
    cdlOFNNoReadOnlyReturn = &H8000&
    cdlOFNNoValidate = &H100&
    cdlOFNOverwritePrompt = &H2&
    cdlOFNPathMustExist = &H800&
    cdlOFNReadOnly = &H1&
    cdlOFNShareAware = &H4000&
End Enum

' Types defined for opening a common dialog box to select files without OCX dependencies
Private Type OPENFILENAME
    lStructSize As Long    'The size of this struct (Use the Len function)
    hwndOwner As Long       'The hWnd of the owner window. The dialog will be modal to this window
    hInstance As Long            'The instance of the calling thread. You can use the App.hInstance here.
    lpstrFilter As String        'Use this to filter what files are showen in the dialog. Separate each filter with Chr$(0). The string also has to end with a Chr(0).
    lpstrCustomFilter As String  'The pattern the user has choosed is saved here if you pass a non empty string. I never use this one
    nMaxCustFilter As Long       'The maximum saved custom filters. Since I never use the lpstrCustomFilter I always pass 0 to this.
    nFilterIndex As Long         'What filter (of lpstrFilter) is showed when the user opens the dialog.
    lpstrFile As String          'The path and name of the file the user has chosed. This must be at least MAX_PATH (260) character long.
    nMaxFile As Long             'The length of lpstrFile + 1
    lpstrFileTitle As String     'The name of the file. Should be MAX_PATH character long
    nMaxFileTitle As Long        'The length of lpstrFileTitle + 1
    lpstrInitialDir As String    'The path to the initial path :) If you pass an empty string the initial path is the current path.
    lpstrTitle As String         'The caption of the dialog.
    Flags As FileOpenConstants                'Flags. See the values in MSDN Library (you can look at the flags property of the common dialog control)
    nFileOffset As Integer       'Points to the what character in lpstrFile where the actual filename begins (zero based)
    nFileExtension As Integer    'Same as nFileOffset except that it points to the file extention.
    lpstrDefExt As String        'Can contain the extention Windows should add to a file if the user doesn't provide one (used with the GetSaveFileName API function)
    lCustData As Long            'Only used if you provide a Hook procedure (Making a Hook procedure is pretty messy in VB.
    lpfnHook As Long             'Pointer to the hook procedure.
    lpTemplateName As String     'A string that contains a dialog template resource name. Only used with the hook procedure.
End Type

Private Type BROWSEINFO
    hwndOwner As Long
    pidlRoot As Long 'LPCITEMIDLIST
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long  'BFFCALLBACK
    lParam As Long
    iImage As Long
End Type

' vars defined for opening a common dialog box to select files without OCX dependencies
Private x_OpenFilename As OPENFILENAME

' APIs declared for opening a common dialog box to select files without OCX dependencies
Private Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (lpofn As OPENFILENAME) As Long
Private Declare Function SHBrowseForFolderA Lib "Shell32.dll" (bInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDListA Lib "Shell32.dll" (ByVal pidl As Long, ByVal szPath As String) As Long
Private Declare Function CoTaskMemFree Lib "ole32.dll" (lp As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'------------------------------------------------------ ENDS

' APIs and variables for querying processes START
Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * 260
End Type

Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Const TH32CS_SNAPPROCESS As Long = 2&
Private uProcess   As PROCESSENTRY32
Private hSnapshot As Long

Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32.dll" Alias "Process32First" (ByVal hSnapshot As Long, ByRef uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32.dll" Alias "Process32Next" (ByVal hSnapshot As Long, ByRef uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32.dll" (ByVal lFlags As Long, ByRef lProcessID As Long) As Long ' Alias "CreateToolhelp32Snapshot"
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
' APIs for querying processes END

' APIs and variables for querying running processes' paths START
Private Const PROCESS_QUERY_INFORMATION As Long = &H400
Private Const PROCESS_VM_READ As Long = (&H10)
Private Const API_NULL As Long = 0

Private Declare Function GetProcessImageFileName Lib "psapi.dll" Alias "GetProcessImageFileNameA" (ByVal hProcess As Long, ByVal lpImageFileName As String, ByVal nSize As Long) As Long
' APIs and variables for querying running processes' paths ENDS

Private lstDevices(1, 25) As String
Private lstDevicesListCount As Integer
Public sAllDrives As String

'Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByVal lpRect As RECT) As Long
'Private Declare Function GetClientRect Lib "user32.dll" (ByVal hwnd As Long, ByVal lpRect As RECT) As Long
'
'Public Type RECT
'  Left As Long
'  Top As Long
'  Right As Long ' This is +1 (right - left = width)
'  Bottom As Long ' This is +1 (bottom - top = height)
'End Type

'------------------------------------------------------ STARTS
' APIs, constants and types defined for determining the OS version
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
  OSVSize         As Long
  dwVerMajor      As Long
  dwVerMinor      As Long
  dwBuildNumber   As Long
  PlatformID      As Long
  szCSDVersion    As String * 128
End Type

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
'------------------------------------------------------ ENDS

'------------------------------------------------------ STARTS
' stored vars read from settings.ini
'
' general
Public PzEStartup As String
Public PzEGaugeFunctions As String
'Public PzEAnimationInterval As String
'Public 'PzEWidgetSkew As String

' config
Public PzEEnableTooltips As String
Public PzEEnableBalloonTooltips As String
Public PzEShowTaskbar As String

Public PzEGaugeSize As String
Public PzEScrollWheelDirection As String

' position
Public PzEAspectHidden As String
Public PzEWidgetPosition As String
Public PzEWidgetLandscape As String
Public PzEWidgetPortrait As String
Public PzELandscapeFormHoffset As String
Public PzELandscapeFormVoffset As String
Public PzEPortraitHoffset As String
Public PzEPortraitYoffset As String
Public PzEvLocationPercPrefValue As String
Public PzEhLocationPercPrefValue As String

' sounds
Public PzEEnableSounds  As String

' development
Public PzEDebug As String
Public PzEDblClickCommand As String
Public PzEOpenFile As String
Public PzEDefaultEditor As String
       
' font
Public PzEPrefsFont  As String
Public PzEPrefsFontSize As String
Public PzEPrefsFontItalics  As String
Public PzEPrefsFontColour  As String

' window
Public PzEWindowLevel As String
Public PzEPreventDragging As String
Public PzEOpacity  As String
Public PzEWidgetHidden  As String
Public PzEHidingTime  As String
Public PzEIgnoreMouse  As String
Public PzEFirstTimeRun  As String

' General storage variables declared
Public PzESettingsDir As String
Public PzESettingsFile As String

Public PzETrinketsDir      As String
Public PzETrinketsFile      As String

Public PzEMaximiseFormX As String
Public PzEMaximiseFormY As String
Public PzELastSelectedTab As String
Public PzESkinTheme As String
Public PzEUnhide As String

' vars stored for positioning the prefs form
Public PzEFormXPosTwips As String
Public PzEFormYPosTwips As String
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
' General variables declared
Public toolSettingsFile  As String
Public classicThemeCapable As Boolean
Public storeThemeColour As Long
Public windowsVer As String

' vars to obtain correct screen width (to correct VB6 bug)
Public screenWidthTwips As Long
Public screenHeightTwips As Long
Public screenHeightPixels As Long
Public screenWidthPixels As Long
Public oldScreenHeightPixels As Long
Public oldScreenWidthPixels As Long

' key presses
Public CTRL_1 As Boolean
Public SHIFT_1 As Boolean

' other globals
Public debugflg As Integer
Public minutesToHide As Integer
Public aspectRatio As String
  
Public oldPzESettingsModificationTime  As Date

Public Const visibleAreaWidth As Long = 648 ' this is the width of the rightmost visible point of the widget - ie. the surround
'------------------------------------------------------ ENDS

'------------------------------------------------------ STARTS
Private Const OF_EXIST         As Long = &H4000
Private Const OFS_MAXPATHNAME  As Long = 128
Private Const HFILE_ERROR      As Long = -1
 
Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type
     
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, _
                            lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

'------------------------------------------------------ ENDS
                            
     
'---------------------------------------------------------------------------------------
' Procedure : fFExists
' Author    : RobDog888 https://www.vbforums.com/member.php?17511-RobDog888
' Date      : 19/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function fFExists(ByVal Fname As String) As Boolean
 
    Dim lRetVal As Long
    Dim OfSt As OFSTRUCT
    
    On Error GoTo fFExists_Error
    
    lRetVal = OpenFile(Fname, OfSt, OF_EXIST)
    If lRetVal <> HFILE_ERROR Then
        fFExists = True
    Else
        fFExists = False
    End If

   On Error GoTo 0
   Exit Function

fFExists_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fFExists of Module Module1"
    
End Function



'---------------------------------------------------------------------------------------
' Procedure : fDirExists
' Author    : zeezee https://www.vbforums.com/member.php?90054-zeezee
' Date      : 19/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function fDirExists(ByVal pstrFolder As String) As Boolean
   On Error GoTo fDirExists_Error

    fDirExists = (PathFileExists(pstrFolder) = 1)
    If fDirExists Then fDirExists = (PathIsDirectory(pstrFolder) <> 0)

   On Error GoTo 0
   Exit Function

fDirExists_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fDirExists of Module Module1"
End Function
''---------------------------------------------------------------------------------------
'' Procedure : fFExists
'' Author    : beededea
'' Date      : 17/10/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Function fFExists(ByRef OrigFile As String) As Boolean
'    Dim FS As Object
'    On Error GoTo fFExists_Error
'   'If debugflg = 1  Then Debug.Print "%fFExists"
'
'    Set FS = CreateObject("Scripting.FileSystemObject")
'    fFExists = FS.FileExists(OrigFile)
'
'   On Error GoTo 0
'   Exit Function
'
'fFExists_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fFExists of module module1"
'End Function


'---------------------------------------------------------------------------------------
' Procedure : fDirExists
' Author    : beededea
' Date      : 17/10/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Public Function fDirExists(ByRef OrigFile As String) As Boolean
'    Dim FS As Object
'    On Error GoTo fDirExists_Error
'   '''If debugflg = 1  Then msgBox "%fDirExists"
'
'    Set FS = CreateObject("Scripting.FileSystemObject")
'    fDirExists = FS.FolderExists(OrigFile)
'
'   On Error GoTo 0
'   Exit Function
'
'fDirExists_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fDirExists of module module1"
'End Function




'---------------------------------------------------------------------------------------
' Procedure : fExtractSuffix
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : extract the suffix from a filename
'---------------------------------------------------------------------------------------
'
Public Function fExtractSuffix(ByVal strPath As String) As String

    
    Dim stringBits() As String ' string array
    Dim upperBit As Integer: upperBit = 0
    
    On Error GoTo fExtractSuffix_Error
    '''If debugflg = 1  Then DebugPrint "%" & "fnExtractSuffix"
   
    If strPath = vbNullString Then
        fExtractSuffix = vbNullString
        Exit Function
    End If
        
    If InStr(strPath, ".") <> 0 Then
        stringBits = Split(strPath, ".")
        upperBit = UBound(stringBits)
        fExtractSuffix = stringBits(upperBit)
    Else
        fExtractSuffix = strPath
    End If

   On Error GoTo 0
   Exit Function

fExtractSuffix_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fExtractSuffix of module module1"
End Function
'---------------------------------------------------------------------------------------
' Procedure : fExtractSuffixWithDot
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : extract the suffix from a filename
'---------------------------------------------------------------------------------------
'
Public Function fExtractSuffixWithDot(ByVal strPath As String) As String
    
    Dim stringBits() As String ' string array
    Dim upperBit As Integer:    upperBit = 0
    
    On Error GoTo fExtractSuffixWithDot_Error
    '''If debugflg = 1  Then DebugPrint "%" & "fExtractSuffixWithDot"
   
    If strPath = vbNullString Then
        fExtractSuffixWithDot = vbNullString
        Exit Function
    End If
        
    If InStr(strPath, ".") <> 0 Then
        stringBits = Split(strPath, ".")
        upperBit = UBound(stringBits)
        fExtractSuffixWithDot = "." & stringBits(upperBit)
    Else
        fExtractSuffixWithDot = vbNullString
    End If

   On Error GoTo 0
   Exit Function

fExtractSuffixWithDot_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fExtractSuffixWithDot of module module1"
End Function

'---------------------------------------------------------------------------------------
' Procedure : fExtractFileNameNoSuffix
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : extract the filename without a suffix
'---------------------------------------------------------------------------------------
'
Public Function fExtractFileNameNoSuffix(ByVal strPath As String) As String
    
    Dim stringBits() As String ' string array
    Dim lowerBit As Integer:    lowerBit = 0
    
    On Error GoTo fExtractFileNameNoSuffix_Error
    '''If debugflg = 1  Then DebugPrint "%" & "fnExtractFileNameNoSuffix"
   
    If strPath = vbNullString Then
        fExtractFileNameNoSuffix = vbNullString
        Exit Function
    End If
        
    If InStr(strPath, ".") <> 0 Then
        stringBits = Split(strPath, ".")
        lowerBit = LBound(stringBits)
        fExtractFileNameNoSuffix = stringBits(lowerBit)
    Else
        fExtractFileNameNoSuffix = strPath
    End If

   On Error GoTo 0
   Exit Function

fExtractFileNameNoSuffix_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fExtractFileNameNoSuffix of module module1"
End Function
'
'---------------------------------------------------------------------------------------
' Procedure : checkLicenceState
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : check the state of the licence
'---------------------------------------------------------------------------------------
'
Public Sub checkLicenceState()
    Dim slicence As String: slicence = "0"
    On Error GoTo checkLicenceState_Error
    ''If debugflg = 1  Then DebugPrint "%" & "checkLicenceState"
    
    ' read the tool's own settings file
    If fFExists(PzESettingsFile) Then ' does the tool's own settings.ini exist?
        slicence = fGetINISetting("Software\PzKill", "Licence", PzESettingsFile)
        ' if the licence state is not already accepted then display the licence form
        If slicence = "0" Then
            Call LoadFileToTB(frmLicence.txtLicenceTextBox, App.Path & "\Resources\txt\licence.txt", False)
            
            frmLicence.show vbModal ' show the licence screen in VB modal mode (ie. on its own)
            ' on the licence box change the state fo the licence acceptance
        End If
    End If

   On Error GoTo 0
   Exit Sub

checkLicenceState_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkLicenceState of Form common"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : LoadFileToTB
' Author    : beededea
' Date      : 26/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub LoadFileToTB(ByVal TxtBox As Object, ByVal FilePath As String, Optional ByVal Append As Boolean = False)
    'PURPOSE: Loads file specified by FilePath into textcontrol
    '(e.g., Text Box, Rich Text Box) specified by TxtBox
    
    'If Append = true, then loaded text is appended to existing
    ' contents else existing contents are overwritten
    
    'Returns: True if Successful, false otherwise
    
    Dim iFile As Integer: iFile = 0
    Dim s As String: s = vbNullString
    
    On Error GoTo LoadFileToTB_Error

   ''If debugflg = 1  Then msgbox "%" & LoadFileToTB

    If Dir$(FilePath) = vbNullString Then Exit Sub
    
    On Error GoTo ErrorHandler:
    s = TxtBox.Text
    
    iFile = FreeFile
    Open FilePath For Input As #iFile
    s = Input(LOF(iFile), #iFile)
    If Append Then
        TxtBox.Text = TxtBox.Text & s
    Else
        TxtBox.Text = s
    End If
    
    'LoadFileToTB = True
    
ErrorHandler:
    If iFile > 0 Then Close #iFile

   On Error GoTo 0
   Exit Sub

LoadFileToTB_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure LoadFileToTB of Form common"

End Sub


'
'---------------------------------------------------------------------------------------
' Procedure : fGetINISetting
' Author    : beededea
' Date      : 05/07/2019
' Purpose   : Get the INI Setting from the File
'---------------------------------------------------------------------------------------
'
Public Function fGetINISetting(ByVal sHeading As String, ByVal sKey As String, ByRef sINIFileName As String) As String
   On Error GoTo fGetINISetting_Error
    Const cparmLen As Integer = 500 ' maximum no of characters allowed in the returned string
    Dim sReturn As String * cparmLen ' not going to initialise this with a 500 char string
    Dim sDefault As String * cparmLen
    Dim lLength As Long: lLength = 0

    lLength = GetPrivateProfileString(sHeading, sKey, sDefault, sReturn, cparmLen, sINIFileName)
    fGetINISetting = Mid$(sReturn, 1, lLength)

   On Error GoTo 0
   Exit Function

fGetINISetting_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fGetINISetting of module module1"
End Function

'
'---------------------------------------------------------------------------------------
' Procedure : sPutINISetting
' Author    : beededea
' Date      : 05/07/2019
' Purpose   : Save INI Setting in the File
'---------------------------------------------------------------------------------------
'
Public Sub sPutINISetting(ByVal sHeading As String, ByVal sKey As String, ByVal sSetting As String, ByRef sINIFileName As String)

   On Error GoTo sPutINISetting_Error

    Dim unusedReturnValue As Long: unusedReturnValue = 0
    
    unusedReturnValue = WritePrivateProfileString(sHeading, sKey, sSetting, sINIFileName)

   On Error GoTo 0
   Exit Sub

sPutINISetting_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sPutINISetting of module module1"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : savestring
' Author    : beededea
' Date      : 05/07/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub savestring(ByRef hKey As Long, ByRef strPath As String, ByRef strvalue As String, ByRef strData As String)

    Dim keyhand As Long: keyhand = 0
    Dim unusedReturnValue As Long: unusedReturnValue = 0
    
    On Error GoTo savestring_Error

    unusedReturnValue = RegCreateKey(hKey, strPath, keyhand)
    unusedReturnValue = RegSetValueEx(keyhand, strvalue, 0, REG_SZ, ByVal strData, Len(strData))
    unusedReturnValue = RegCloseKey(keyhand)

   On Error GoTo 0
   Exit Sub

savestring_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure savestring of module module1"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : fSpecialFolder
' Author    : si_the_geek vbforums
' Date      : 17/10/2019
' Purpose   : Returns the path to the specified special folder (AppData etc)
'---------------------------------------------------------------------------------------
'
Public Function fSpecialFolder(ByVal pfe As FolderEnum) As String
    Const MAX_PATH As Integer = 260
    Dim strPath As String: strPath = vbNullString
    Dim strBuffer As String: strBuffer = vbNullString
    
   On Error GoTo fSpecialFolder_Error

    strBuffer = Space$(MAX_PATH)
    If SHGetFolderPath(0, pfe, 0, 0, strBuffer) = 0 Then strPath = Left$(strBuffer, InStr(strBuffer, vbNullChar) - 1)
    If Right$(strPath, 1) = "\" Then strPath = Left$(strPath, Len(strPath) - 1)
    fSpecialFolder = strPath

   On Error GoTo 0
   Exit Function

fSpecialFolder_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fSpecialFolder of Module Module1"
End Function

'---------------------------------------------------------------------------------------
' Procedure : addTargetfile
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : open a dialogbox to select a file as the target, normally a binary
'---------------------------------------------------------------------------------------
'
Public Sub addTargetFile(ByVal fieldValue As String, ByRef retFileName As String)
    Dim FilePath As String: FilePath = vbNullString
    Dim dialogInitDir As String: dialogInitDir = vbNullString
    Dim retfileTitle As String: retfileTitle = vbNullString
    Const x_MaxBuffer As Integer = 256
    
    ''If debugflg = 1  Then Debug.Print "%" & "addTargetfile"
    
    On Error Resume Next
    
    ' set the default folder to the existing reference
    If Not fieldValue = vbNullString Then
        If fFExists(fieldValue) Then
            ' extract the folder name from the string
            FilePath = fGetDirectory(fieldValue)
            ' set the default folder to the existing reference
            dialogInitDir = FilePath 'start dir, might be "C:\" or so also
        ElseIf fDirExists(fieldValue) Then ' this caters for the entry being just a folder name
            ' set the default folder to the existing reference
            dialogInitDir = fieldValue 'start dir, might be "C:\" or so also
        Else
            dialogInitDir = App.Path 'start dir, might be "C:\" or so also
        End If
    End If
    
  With x_OpenFilename
'    .hwndOwner = Me.hWnd
    .hInstance = App.hInstance
    .lpstrTitle = "Select a File Target"
    .lpstrInitialDir = dialogInitDir
    
    .lpstrFilter = "Text Files" & vbNullChar & "*.txt" & vbNullChar & "All Files" & vbNullChar & "*.*" & vbNullChar & vbNullChar
    .nFilterIndex = 2
    
    .lpstrFile = String(x_MaxBuffer, 0)
    .nMaxFile = x_MaxBuffer - 1
    .lpstrFileTitle = .lpstrFile
    .nMaxFileTitle = x_MaxBuffer - 1
    .lStructSize = Len(x_OpenFilename)
  End With
  

  Call obtainOpenFileName(retFileName, retfileTitle) ' retfile will be buffered to 256 bytes

   On Error GoTo 0
   
   Exit Sub

'addTargetfile_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addTargetfile of module module1.bas"
 
End Sub
'---------------------------------------------------------------------------------------
' Procedure : fGetDirectory
' Author    : beededea
' Date      : 11/07/2019
' Purpose   : get the folder or directory path as a string not including the last backslash
'---------------------------------------------------------------------------------------
'
Public Function fGetDirectory(ByRef Path As String) As String

   On Error GoTo fGetDirectory_Error
   ''If debugflg = 1  Then DebugPrint "%" & "fnGetDirectory"

    If InStrRev(Path, "\") = 0 Then
        fGetDirectory = vbNullString
        Exit Function
    End If
    fGetDirectory = Left$(Path, InStrRev(Path, "\") - 1)

   On Error GoTo 0
   Exit Function

fGetDirectory_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fGetDirectory of module module1"
End Function

'---------------------------------------------------------------------------------------
' Procedure : obtainOpenFileName
' Author    : beededea
' Date      : 02/09/2019
' Purpose   : using GetOpenFileName API rturns file name and title, the filename will be buffered to 256 bytes
'---------------------------------------------------------------------------------------
'
Public Sub obtainOpenFileName(ByRef retFileName As String, ByRef retfileTitle As String)
   On Error GoTo obtainOpenFileName_Error
   ''If debugflg = 1  Then Debug.Print "%obtainOpenFileName"

  If GetOpenFileName(x_OpenFilename) <> 0 Then
'    If x_OpenFilename.lpstrFile = "*.*" Then
'        'txtTarget.Text = savLblTarget
'    Else
        retfileTitle = x_OpenFilename.lpstrFileTitle
        retFileName = x_OpenFilename.lpstrFile
'    End If
  'Else
    'The CANCEL button was pressed
    'MsgBox "Cancel"
  End If

   On Error GoTo 0
   Exit Sub

obtainOpenFileName_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure obtainOpenFileName of module module1.bas"
End Sub





'
'---------------------------------------------------------------------------------------
' Procedure : GetWindowsVersion
' Author    :
' Date      : 28/05/2023
' Purpose   : Returns the version of Windows that the user is running
'---------------------------------------------------------------------------------------
'
Public Function GetWindowsVersion() As String
    Dim osv As OSVERSIONINFO
    
    On Error GoTo GetWindowsVersion_Error

    osv.OSVSize = Len(osv)

    If GetVersionEx(osv) = 1 Then
        Select Case osv.PlatformID
            Case VER_PLATFORM_WIN32s
                GetWindowsVersion = "Win32s on Windows 3.1"
            Case VER_PLATFORM_WIN32_NT
                GetWindowsVersion = "Windows NT"
                
                Select Case osv.dwVerMajor
                    Case 3
                        GetWindowsVersion = "Windows NT 3.5"
                    Case 4
                        GetWindowsVersion = "Windows NT 4.0"
                    Case 5
                        Select Case osv.dwVerMinor
                            Case 0
                                GetWindowsVersion = "Windows 2000"
                            Case 1
                                GetWindowsVersion = "Windows XP"
                            Case 2
                                GetWindowsVersion = "Windows Server 2003"
                        End Select
                    Case 6
                        Select Case osv.dwVerMinor
                            Case 0
                                GetWindowsVersion = "Windows Vista"
                            Case 1
                                GetWindowsVersion = "Windows 7"
                            Case 2
                                GetWindowsVersion = "Windows 8"
                            Case 3
                                GetWindowsVersion = "Windows 8.1"
                            Case 10
                                GetWindowsVersion = "Windows 10"
                        End Select
                End Select
        
            Case VER_PLATFORM_WIN32_WINDOWS:
                Select Case osv.dwVerMinor
                    Case 0
                        GetWindowsVersion = "Windows 95"
                    Case 90
                        GetWindowsVersion = "Windows Me"
                    Case Else
                        GetWindowsVersion = "Windows 98"
                End Select
        End Select
    Else
        GetWindowsVersion = "Unable to identify your version of Windows."
    End If

   On Error GoTo 0
   Exit Function

GetWindowsVersion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetWindowsVersion of Module Module1"
End Function




'----------------------------------------
'Name: TestWinVer
'Description: Tests the multiplicity of Windows versions and returns some values
'----------------------------------------
Public Function fTestClassicThemeCapable() As Boolean

    '=================================
    '2000 / XP / NT / 7 / 8 / 10
    '=================================
    On Error GoTo fTestClassicThemeCapable_Error

    Dim ProgramFilesDir As String: ProgramFilesDir = vbNullString
    Dim strString As String: strString = vbNullString
    'Dim shortWindowsVer As String: shortWindowsVer = vbNullString
    
    fTestClassicThemeCapable = False
    windowsVer = vbNullString
    
    ' other variable assignments
    strString = fGetstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName")
    windowsVer = strString

    ' note that when running in compatibility mode the o/s will respond with "Windows XP"
    ' The IDE runs in compatibility mode so it may report the wrong working folder

    'Get the value of "ProgramFiles", or "ProgramFilesDir"
        
    windowsVer = GetWindowsVersion
    
    Select Case windowsVer
    Case "Windows NT 4.0"
        fTestClassicThemeCapable = True
        strString = fGetstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
    Case "Windows 2000"
        fTestClassicThemeCapable = True
        strString = fGetstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
    Case "Windows XP"
        fTestClassicThemeCapable = True
        strString = fGetstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "ProgramFilesDir")
    Case "Windows Server 2003"
        fTestClassicThemeCapable = True
        strString = fGetstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
    Case "Windows Vista"
        fTestClassicThemeCapable = True
        strString = fGetstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
    Case "Windows 7"
        fTestClassicThemeCapable = True
        strString = fGetstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
    Case Else ' windows 8/10/11+
        fTestClassicThemeCapable = False
        strString = fGetstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "ProgramFilesDir")
    End Select

    ProgramFilesDir = strString
    If ProgramFilesDir = vbNullString Then ProgramFilesDir = "c:\program files (x86)" ' 64bit systems
    If Not fDirExists(ProgramFilesDir) Then
        ProgramFilesDir = "c:\program files" ' 32 bit systems
    End If
   
    On Error GoTo 0: Exit Function

fTestClassicThemeCapable_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fTestClassicThemeCapable of module module1"

End Function




'---------------------------------------------------------------------------------------
' Procedure : fGetstring
' Author    : beededea
' Date      : 05/07/2019
' Purpose   : get a string from the registry
'---------------------------------------------------------------------------------------
'
Public Function fGetstring(ByRef hKey As Long, ByRef strPath As String, ByRef strvalue As String) As String

    Dim keyhand As Long: keyhand = 0
    Dim lResult As Long: lResult = 0
    Dim strBuf As String: strBuf = vbNullString
    Dim lDataBufSize As Long: lDataBufSize = 0
    Dim intZeroPos As Integer: intZeroPos = 0
    Dim unusedReturnValue As Integer: unusedReturnValue = 0

    Dim lValueType As Variant

    On Error GoTo fGetstring_Error

    unusedReturnValue = RegOpenKey(hKey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strvalue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then
        strBuf = String$(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strvalue, 0&, 0&, ByVal strBuf, lDataBufSize)
        Dim ERROR_SUCCESS As Variant
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                fGetstring = Left$(strBuf, intZeroPos - 1)
            Else
                fGetstring = strBuf
            End If
        End If
    End If

   On Error GoTo 0
   Exit Function

fGetstring_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fGetstring of module module1"
End Function



' select a font for the fnt form
'---------------------------------------------------------------------------------------
' Procedure : changeFont
' Author    : beededea
' Date      : 02/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub changeFont(ByVal frm As Form, ByVal fntNow As Boolean, ByRef fntFont As String, ByRef fntSize As Integer, ByRef fntWeight As Integer, ByRef fntStyle As Boolean, ByRef fntColour As Long, ByRef fntItalics As Boolean, ByRef fntUnderline As Boolean, ByRef fntFontResult As Boolean)
    
   On Error GoTo changeFont_Error

    fntWeight = 0
    fntStyle = False
    'fntColour = 0
    'fntBold = False
    'fntUnderline = False
    fntFontResult = False
    
    'If debugflg = 1  Then Debug.Print "%mnuFont_Click"

    displayFontSelector fntFont, fntSize, fntWeight, fntStyle, fntColour, fntItalics, fntUnderline, fntFontResult
    If fntFontResult = False Then Exit Sub
'
'    If fntWeight > 700 Then
'        'fntBold = True
'    Else
'        'fntBold = False
'    End If
    
    If fntFont <> vbNullString And fntNow = True Then
        Call changeFormFont(frm, fntFont, Val(fntSize), fntWeight, fntStyle, fntItalics, fntColour)
    End If
    
   On Error GoTo 0
   Exit Sub

changeFont_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure changeFont of Module Module1"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : displayFontSelector
' Author    : beededea
' Date      : 29/02/2020
' Purpose   : select a font for the supplied form
'---------------------------------------------------------------------------------------
'
Private Sub displayFontSelector(ByRef currFont As String, ByRef currSize As Integer, ByRef currWeight As Integer, ByVal currStyle As Boolean, ByRef currColour As Long, ByRef currItalics As Boolean, ByRef currUnderline As Boolean, ByRef fontResult As Boolean)

    Dim thisFont As FormFontInfo

    On Error GoTo displayFontSelector_Error

    With thisFont
      .Color = currColour
      .Height = currSize
      .Weight = currWeight
      '400     Font is normal.
      '700     Font is bold.
      .Italic = currItalics
      .UnderLine = currUnderline
      .Name = currFont
    End With
    
    fontResult = fDialogFont(thisFont)
    If fontResult = False Then Exit Sub
    
    ' some fonts have naming problems and the result is an empty font name field on the font selector
    If thisFont.Name = vbNullString Then thisFont.Name = "times new roman"
    If thisFont.Name = vbNullString Then Exit Sub
    
    With thisFont
        currFont = .Name
        currSize = .Height
        currWeight = .Weight
        currItalics = .Italic
        currUnderline = .UnderLine
        currColour = .Color
        'ctl = .Name & " - Size:" & .Height
    End With

   On Error GoTo 0
   Exit Sub

displayFontSelector_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure displayFontSelector of module module1"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : changeFormFont
' Author    : beededea
' Date      : 12/07/2019
' Purpose   : change the font throughout the whole form
'---------------------------------------------------------------------------------------
'
Public Sub changeFormFont(ByVal formName As Object, ByVal suppliedFont As String, ByVal suppliedSize As Integer, ByVal suppliedWeight As Integer, ByVal suppliedStyle As Boolean, ByVal suppliedItalics As Boolean, ByVal suppliedColour As Long)
    On Error GoTo changeFormFont_Error
        
    Dim ctrl As Control
      
    ' loop through all the controls and identify the labels and text boxes
    For Each ctrl In formName.Controls
        If (TypeOf ctrl Is CommandButton) Or (TypeOf ctrl Is TextBox) Or (TypeOf ctrl Is FileListBox) Or (TypeOf ctrl Is Label) Or (TypeOf ctrl Is ComboBox) Or (TypeOf ctrl Is CheckBox) Or (TypeOf ctrl Is OptionButton) Or (TypeOf ctrl Is Frame) Or (TypeOf ctrl Is ListBox) Then
            If suppliedFont <> vbNullString Then ctrl.Font.Name = suppliedFont
            If suppliedSize > 0 Then ctrl.Font.Size = suppliedSize
            ctrl.Font.Italic = suppliedItalics
            
            Select Case True
                Case (TypeOf ctrl Is CommandButton)
                    ' stupif fecking VB6 will not let you change the font of the forecolour on a button!
                    'Ctrl.ForeColor = suppliedColour
                    ' do nothing
                Case Else
                    ctrl.ForeColor = suppliedColour
            End Select
        End If
    Next
     
   On Error GoTo 0
   Exit Sub

changeFormFont_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure changeFormFont of module module1"
    
End Sub



'---------------------------------------------------------------------------------------
' Procedure : fDialogFont
' Author    : beededea
' Date      : 21/08/2020
' Purpose   : display the default windows dialog box that allows the user to select a font
'---------------------------------------------------------------------------------------
'
Public Function fDialogFont(ByRef f As FormFontInfo) As Boolean
      
    Dim logFnt As LOGFONT
    Dim ftStruc As FONTSTRUC
    Dim lLogFontAddress As Long: lLogFontAddress = 0
    Dim lMemHandle As Long: lMemHandle = 0
    Dim hWndAccessApp As Long: hWndAccessApp = 0
    
    Const LOGPIXELSY As Integer = 90        '  Logical pixels/inch in Y

    On Error GoTo fDialogFont_Error
    
    logFnt.lfWeight = f.Weight
    logFnt.lfItalic = f.Italic * -1
    logFnt.lfUnderline = f.UnderLine * -1
    logFnt.lfHeight = -fMulDiv(CLng(f.Height), GetDeviceCaps(GetDC(hWndAccessApp), LOGPIXELSY), 72)
    f.Name = "Centurion Light SF"
    Call StringToByte(f.Name, logFnt.lfFaceName()) ' HERE
    ftStruc.rgbColors = f.Color
    ftStruc.lStructSize = Len(ftStruc)
    
    lMemHandle = GlobalAlloc(GHND, Len(logFnt))
    If lMemHandle = 0 Then
      fDialogFont = False
      Exit Function
    End If

    lLogFontAddress = GlobalLock(lMemHandle)
    If lLogFontAddress = 0 Then
      fDialogFont = False
      Exit Function
    End If
    
    CopyMemory ByVal lLogFontAddress, logFnt, Len(logFnt)
    ftStruc.lpLogFont = lLogFontAddress
    'ftStruc.flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT
    ftStruc.Flags = CF_SCREENFONTS Or CF_INITTOLOGFONTSTRUCT
    If ChooseFont(ftStruc) = 1 Then
      CopyMemory logFnt, ByVal lLogFontAddress, Len(logFnt)
      f.Weight = logFnt.lfWeight
      f.Italic = CBool(logFnt.lfItalic)
      f.UnderLine = CBool(logFnt.lfUnderline)
      f.Name = fByteToString(logFnt.lfFaceName())
      f.Height = CLng(ftStruc.iPointSize / 10)
      f.Color = ftStruc.rgbColors
      fDialogFont = True
    Else
      fDialogFont = False
    End If

   On Error GoTo 0
   Exit Function

fDialogFont_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fDialogFont of Module module1"
End Function


'---------------------------------------------------------------------------------------
' Procedure : fMulDiv
' Author    :
' Date      : 21/08/2020
' Purpose   :  fMulDiv function multiplies two 32-bit values and then divides the 64-bit result by a third 32-bit value.
'---------------------------------------------------------------------------------------
'
Private Function fMulDiv(ByVal In1 As Long, ByVal In2 As Long, ByVal In3 As Long) As Long
        
    Dim lngTemp As Long: lngTemp = 0
    On Error GoTo fMulDiv_Error
    
    On Error GoTo fMulDiv_err
    If In3 <> 0 Then
        lngTemp = In1 * In2
        lngTemp = lngTemp / In3
    Else
        lngTemp = -1
    End If

    fMulDiv = lngTemp
    Exit Function
fMulDiv_err:
    lngTemp = -1
    Resume fMulDiv_err

   On Error GoTo 0
   Exit Function

fMulDiv_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fMulDiv of Module module1"
End Function



'---------------------------------------------------------------------------------------
' Procedure : StringToByte
' Author    :
' Date      : 21/08/2020
' Purpose   : convert a provided string to a byte array
'---------------------------------------------------------------------------------------
'
Private Sub StringToByte(ByVal InString As String, ByRef ByteArray() As Byte)
    
    Dim intLbound As Integer: intLbound = 0
    Dim intUbound As Integer: intUbound = 0
    Dim intLen As Integer: intLen = 0
    Dim intX As Integer: intX = 0
    
    On Error GoTo StringToByte_Error

    intLbound = LBound(ByteArray)
    intUbound = UBound(ByteArray)
    intLen = Len(InString)
    If intLen > intUbound - intLbound Then intLen = intUbound - intLbound
    For intX = 1 To intLen
        ByteArray(intX - 1 + intLbound) = Asc(Mid(InString, intX, 1))
    Next

   On Error GoTo 0
   Exit Sub

StringToByte_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure StringToByte of Module module1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fByteToString
' Author    :
' Date      : 21/08/2020
' Purpose   : convert a byte array provided to a string
'---------------------------------------------------------------------------------------
'
Private Function fByteToString(ByRef aBytes() As Byte) As String
      
    Dim dwBytePoint As Long: dwBytePoint = 0
    Dim dwByteVal As Long: dwByteVal = 0
    Dim szOut As String: szOut = vbNullString
    
    On Error GoTo fByteToString_Error

    dwBytePoint = LBound(aBytes)
    While dwBytePoint <= UBound(aBytes) ' whileing and wending my way through the bytearrays >sigh<
      dwByteVal = aBytes(dwBytePoint)
      If dwByteVal = 0 Then
        fByteToString = szOut
        Exit Function
      Else
        szOut = szOut & Chr$(dwByteVal)
      End If
      dwBytePoint = dwBytePoint + 1
    Wend
    fByteToString = szOut

   On Error GoTo 0
   Exit Function

fByteToString_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fByteToString of Module module1"
End Function

'---------------------------------------------------------------------------------------
' Procedure : aboutClickEvent
' Author    : beededea
' Date      : 02/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub aboutClickEvent()

    ' The RC forms are measured in pixels so the positioning needs to turn the twips into pixels
    On Error GoTo aboutClickEvent_Error
   
    fMain.aboutForm.Top = (screenHeightPixels / 2) - (fMain.aboutForm.Height / 2)
    fMain.aboutForm.Left = (screenWidthPixels / 2) - (fMain.aboutForm.Width / 2)
     
    fMain.aboutForm.Load
    fMain.aboutForm.show
    
    'aboutWidget.opacity = 0
    aboutWidget.show = True
    aboutWidget.Widget.Refresh
    
     If (fMain.aboutForm.WindowState = 1) Then
         fMain.aboutForm.WindowState = 0
     End If

   On Error GoTo 0
   Exit Sub

aboutClickEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure aboutClickEvent of Module Module1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuCoffee_ClickEvent
' Author    : beededea
' Date      : 20/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub mnuCoffee_ClickEvent()
    Dim answer As VbMsgBoxResult: answer = vbNo
    On Error GoTo mnuCoffee_ClickEvent_Error
    
    answer = MsgBox(" Help support the creation of more widgets like this, DO send us a coffee! This button opens a browser window and connects to the Kofi donate page for this widget). Will you be kind and proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(menuForm.hwnd, "Open", "https://www.ko-fi.com/yereverluvinunclebert", vbNullString, App.Path, 1)
    End If

   On Error GoTo 0
   Exit Sub

mnuCoffee_ClickEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuCoffee_ClickEvent of Module Module1"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuSupport_ClickEvent
' Author    : beededea
' Date      : 20/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub mnuSupport_ClickEvent()

    Dim answer As VbMsgBoxResult: answer = vbNo
    On Error GoTo mnuSupport_ClickEvent_Error
    
    answer = MsgBox("Visiting the support page - this button opens a browser window and connects to our Github issues page where you can send us a support query. Proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(menuForm.hwnd, "Open", "https://github.com/yereverluvinunclebert/Panzer-Kill-gauge-VB6/issues", vbNullString, App.Path, 1)
    End If

   On Error GoTo 0
   Exit Sub

mnuSupport_ClickEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuSupport_ClickEvent of Module Module1"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuLicence_ClickEvent
' Author    : beededea
' Date      : 20/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub mnuLicence_ClickEvent()

   On Error GoTo mnuLicence_ClickEvent_Error

    Call LoadFileToTB(frmLicence.txtLicenceTextBox, App.Path & "\Resources\txt\licence.txt", False)
    frmLicence.show

   On Error GoTo 0
   Exit Sub

mnuLicence_ClickEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLicence_ClickEvent of Module Module1"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : setMainTooltips
' Author    : beededea
' Date      : 15/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub setMainTooltips()
   On Error GoTo setMainTooltips_Error

    If PzEEnableTooltips = "1" Then
        'BodyWidget.Widget.FontName = PzEPrefsFont ' does not apply to the tooltip
        BodyWidget.Widget.ToolTip = "Use CTRL+mouse scrollwheel up/down to resize."
    Else
        BodyWidget.Widget.ToolTip = ""
    End If
    
    Call ChangeToolTipWidgetDefaultSettings(Cairo.ToolTipWidget.Widget)

   On Error GoTo 0
   Exit Sub

setMainTooltips_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setMainTooltips of Module Module1"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : ChangeToolTipWidgetDefaultSettings
' Author    : beededea
' Date      : 20/06/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub ChangeToolTipWidgetDefaultSettings(My_Widget As cWidgetBase)

   On Error GoTo ChangeToolTipWidgetDefaultSettings_Error

    With My_Widget

    .FontName = PzEPrefsFont
    .FontSize = PzEPrefsFontSize

    End With

   On Error GoTo 0
   Exit Sub

ChangeToolTipWidgetDefaultSettings_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ChangeToolTipWidgetDefaultSettings of Module Module1"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : savePosition
' Author    : beededea
' Date      : 04/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub savePosition()

   On Error GoTo savePosition_Error

    PzEMaximiseFormX = Str$(fMain.BodyForm.Left) ' saving in pixels
    PzEMaximiseFormY = Str$(fMain.BodyForm.Top)
   
    PzEGaugeSize = LTrim$(Str$(BodyWidget.Zoom * 100))
    
    sPutINISetting "Software\PzKill", "maximiseFormX", PzEMaximiseFormX, PzESettingsFile
    sPutINISetting "Software\PzKill", "maximiseFormY", PzEMaximiseFormY, PzESettingsFile
    sPutINISetting "Software\PzKill", "gaugeSize", PzEGaugeSize, PzESettingsFile
    
   On Error GoTo 0
   Exit Sub

savePosition_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure savePosition of Module Module1"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : makeVisibleFormElements
' Author    : beededea
' Date      : 01/03/2023
' Purpose   : ' adjust Form Position on startup placing form onto Correct Monitor when placed off screen due to
'               monitor/resolution changes.
'---------------------------------------------------------------------------------------
'
Public Sub makeVisibleFormElements()
    
    On Error GoTo makeVisibleFormElements_Error

    'NOTE that when you position a widget you are positioning the form it is drawn upon.
    
'    PzEMaximiseFormX = fGetINISetting("Software\PzKill", "maximiseFormX", PzESettingsFile)
'    PzEMaximiseFormY = fGetINISetting("Software\PzKill", "maximiseFormY", PzESettingsFile)
    
    fMain.BodyForm.Left = Val(PzEMaximiseFormX)
    fMain.BodyForm.Top = Val(PzEMaximiseFormY)

    ' The RC forms are measured in pixels, do remember that...

    fMain.BodyForm.show

    On Error GoTo 0
    Exit Sub

makeVisibleFormElements_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure makeVisibleFormElements of Module Module1"
            Resume Next
          End If
    End With
        
End Sub



'
''---------------------------------------------------------------------------------------
'' Procedure : fTwipsPerPixelX
'' Author    : Elroy from Vbforums
'' Date      : 23/01/2022
'' Purpose   : This works even on Tablet PC.  The problem is: when the tablet screen is rotated, the "Screen" object of VB doesn't pick it up.
''---------------------------------------------------------------------------------------
''
'Public Function fTwipsPerPixelX() As Single
'    Dim hdc As Long: hdc = 0
'    Dim lPixelsPerInch As Long: lPixelsPerInch = 0
'
'    Const LOGPIXELSX As Integer = 88        '  Logical pixels/inch in X
'    Const POINTS_PER_INCH As Long = 72 ' A point is defined as 1/72 inches.
'    Const TWIPS_PER_POINT As Long = 20 ' Also, by definition.
'    '
'    On Error GoTo fTwipsPerPixelX_Error
'
'    ' 23/01/2022 .01 monitorModule.bas DAEB added if then else if you can't get device context
'    hdc = GetDC(0)
'    If hdc <> 0 Then
'        lPixelsPerInch = GetDeviceCaps(hdc, LOGPIXELSX)
'        ReleaseDC 0, hdc
'        fTwipsPerPixelX = TWIPS_PER_POINT * (POINTS_PER_INCH / lPixelsPerInch) ' Cancel units to see it.
'    Else
'        fTwipsPerPixelX = Screen.TwipsPerPixelX
'    End If
'
'   On Error GoTo 0
'   Exit Function
'
'fTwipsPerPixelX_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fTwipsPerPixelX of Module Module1"
'End Function
'
''---------------------------------------------------------------------------------------
'' Procedure : fTwipsPerPixelY
'' Author    : Elroy from Vbforums
'' Date      : 23/01/2022
'' Purpose   : This works even on Tablet PC.  The problem is: when the tablet screen is rotated, the "Screen" object of VB doesn't pick it up.
''---------------------------------------------------------------------------------------
''
'Public Function fTwipsPerPixelY() As Single
'    Dim hdc As Long: hdc = 0
'    Dim lPixelsPerInch As Long: lPixelsPerInch = 0
'
'    Const LOGPIXELSY As Integer = 90         '  Logical pixels/inch in Y
'    Const POINTS_PER_INCH As Long = 72 ' A point is defined as 1/72 inches.
'    Const TWIPS_PER_POINT As Long = 20 ' Also, by definition.
'
'   On Error GoTo fTwipsPerPixelY_Error
'
'    ' 23/01/2022 .01 monitorModule.bas DAEB added if then else if you can't get device context
'    hdc = GetDC(0)
'    If hdc <> 0 Then
'        lPixelsPerInch = GetDeviceCaps(hdc, LOGPIXELSY)
'        ReleaseDC 0, hdc
'        fTwipsPerPixelY = TWIPS_PER_POINT * (POINTS_PER_INCH / lPixelsPerInch) ' Cancel units to see it.
'    Else
'        fTwipsPerPixelY = Screen.TwipsPerPixelY
'    End If
'
'   On Error GoTo 0
'   Exit Function
'
'fTwipsPerPixelY_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fTwipsPerPixelY of Module Module1"
'
'End Function


'---------------------------------------------------------------------------------------
' Procedure : getkeypress
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : getting a keypress from the keyboard
    '36 home
    '40 is down
    '38 is up
    '37 is left
    '39 is right
    '33  Page up
    '34  Page down
    '35  End
    'ctrl 116
    'Shift 16
    'f5 18
'---------------------------------------------------------------------------------------
'
Public Sub getKeyPress(ByVal KeyCode As Integer, ByVal Shift As Integer)

    On Error GoTo getkeypress_Error

    If CTRL_1 Or SHIFT_1 Then
            CTRL_1 = False
            SHIFT_1 = False
    End If
    
    If Shift Then
        SHIFT_1 = True
    End If

    Select Case KeyCode
        Case vbKeyControl
            CTRL_1 = True
        Case vbKeyShift
            SHIFT_1 = True
        Case 116
            Call reloadWidget 'f5 refresh button as per all browsers
    End Select
 
    On Error GoTo 0
   Exit Sub

getkeypress_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getkeypress of Module module1"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : determineScreenDimensions
' Author    : beededea
' Date      : 18/09/2020
' Purpose   : VB6 has a bug - it should return width in twips on my screen but often returns a faulty value when a game runs full screen, changing the resolution
'             the screen width determination is incorrect, the API call below resolves this.
'---------------------------------------------------------------------------------------
'
Public Sub determineScreenDimensions()

   On Error GoTo determineScreenDimensions_Error
   
    'If debugflg = 1 Then msgbox "% sub determineScreenDimensions"

    ' only calling TwipsPerPixelX/Y functions once on startup
    screenTwipsPerPixelX = fTwipsPerPixelX
    screenTwipsPerPixelY = fTwipsPerPixelY
    
    screenHeightPixels = GetDeviceCaps(menuForm.hdc, VERTRES) ' we use the name of any form that we don't mind being loaded at this point
    screenWidthPixels = GetDeviceCaps(menuForm.hdc, HORZRES)

    screenHeightTwips = screenHeightPixels * screenTwipsPerPixelY
    screenWidthTwips = screenWidthPixels * screenTwipsPerPixelX
    
    oldScreenHeightPixels = screenHeightPixels ' will be used to check for orientation changes
    oldScreenWidthPixels = screenWidthPixels
    
   On Error GoTo 0
   Exit Sub

determineScreenDimensions_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & " in procedure determineScreenDimensions of Form dock"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mainScreen
' Author    : beededea
' Date      : 04/05/2023
' Purpose   : Function to move the main_window onto the main screen - called on startup and by timer
'---------------------------------------------------------------------------------------
'
Public Sub mainScreen()
   On Error GoTo mainScreen_Error

    ' check for aspect ratio and determine whether it is in portrait or landscape mode
    If screenWidthPixels > screenHeightPixels Then
        aspectRatio = "landscape"
    Else
        aspectRatio = "portrait"
    End If
    
    ' check if the widget has a lock for the screen type.
    If aspectRatio = "landscape" Then
        If PzEWidgetLandscape = "1" Then
            If PzELandscapeFormHoffset <> vbNullString Then
                fMain.BodyForm.Left = Val(PzELandscapeFormHoffset)
                fMain.BodyForm.Top = Val(PzELandscapeFormVoffset)
            End If
        End If
        If PzEAspectHidden = "2" Then
            'Print "Hiding the widget for landscape mode"
            BodyWidget.opacity = 0
        End If
    End If
    
    ' check if the widget has a lock for the screen type.
    If aspectRatio = "portrait" Then
        If PzEWidgetPortrait = "1" Then
            fMain.BodyForm.Left = Val(PzEPortraitHoffset)
            fMain.BodyForm.Top = Val(PzEPortraitYoffset)
        End If
        If PzEAspectHidden = "1" Then
            'Print "Hiding the widget for portrait mode"
            BodyWidget.opacity = 0
        End If
    End If

    ' calculate the on screen widget position
    If fMain.BodyForm.Left < 0 Then
        fMain.BodyForm.Left = 10
    End If
    If fMain.BodyForm.Top < 0 Then
        fMain.BodyForm.Top = 0
    End If
    If fMain.BodyForm.Left > screenWidthPixels - 50 Then
        fMain.BodyForm.Left = screenWidthPixels - 150
    End If
    If fMain.BodyForm.Top > screenHeightPixels - 50 Then
        fMain.BodyForm.Top = screenHeightPixels - 150
    End If

    ' calculate the current hlocation in % of the screen
    ' store the current hlocation in % of the screen
    If PzEWidgetPosition = "1" Then
        PzEhLocationPercPrefValue = Str$(fMain.BodyForm.Left / screenWidthPixels * 100)
        PzEvLocationPercPrefValue = Str$(fMain.BodyForm.Top / screenHeightPixels * 100)
    End If

   On Error GoTo 0
   Exit Sub

mainScreen_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mainScreen of Module Module1"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : BodyForm_Unload
' Author    : beededea
' Date      : 18/08/2022
' Purpose   : the standard form unload routine
'---------------------------------------------------------------------------------------
'
Public Sub BodyForm_Unload() ' name follows VB6 standard naming convention
    
    On Error GoTo Form_Unload_Error
    
    Call savePosition
    
    Call unloadAllForms(True)

    On Error GoTo 0
    Exit Sub

Form_Unload_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Class Module cfMain"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : unloadAllForms
' Author    : beededea
' Date      : 28/06/2023
' Purpose   : unload all VB6 and RC5 forms
'---------------------------------------------------------------------------------------
'
Public Sub unloadAllForms(ByVal endItAll As Boolean)
    
   On Error GoTo unloadAllForms_Error

    'unload the RC5 widgets on the RC5 forms first
    
    aboutWidget.Widgets.RemoveAll
    BodyWidget.Widgets.RemoveAll
    
    ' unload the native VB6 and RC5 forms
    
    Unload panzerPrefs
    Unload helpForm
    Unload frmLicence
    Unload frmTimer
    Unload menuForm

    fMain.aboutForm.Unload  ' RC5's own method for killing forms
    fMain.BodyForm.Unload
    
    ' remove all variable references to each form in turn
    
    Set panzerPrefs = Nothing
    Set helpForm = Nothing
    Set fMain.aboutForm = Nothing
    Set fMain.BodyForm = Nothing
    Set frmLicence = Nothing
    Set frmTimer = Nothing
    Set menuForm = Nothing
    
    If endItAll = True Then End

   On Error GoTo 0
   Exit Sub

unloadAllForms_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure unloadAllForms of Module Module1"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : reloadWidget
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub reloadWidget()
    
    On Error GoTo reloadWidget_Error
    
    Call unloadAllForms(False) ' unload forms but do not END
    
    ' this will call the routines as called by sub main() and initialise the program and RELOAD the RC5 forms.
    Call mainRoutine(True) ' sets the restart flag to avoid repriming the Rc5 message pump.

    On Error GoTo 0
    Exit Sub

reloadWidget_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure reloadWidget of Module Module1"
            Resume Next
          End If
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : makeProgramPreferencesAvailable
' Author    : beededea
' Date      : 01/05/2023
' Purpose   : open the prefs
'---------------------------------------------------------------------------------------
'
Public Sub makeProgramPreferencesAvailable()
    On Error GoTo makeProgramPreferencesAvailable_Error
    
    If panzerPrefs.IsVisible = False Then
    
        If panzerPrefs.WindowState = vbMinimized Then
            panzerPrefs.WindowState = vbNormal
            Call readPrefsPosition
        End If
        
        ' set the current position of the utility according to previously stored positions

        If panzerPrefs.WindowState = vbNormal Then
        
            Call readPrefsPosition
            
            If panzerPrefs.Left = 0 Then
                If ((fMain.BodyForm.Left + fMain.BodyForm.Width) * screenTwipsPerPixelX) + 200 + panzerPrefs.Width > screenWidthTwips Then
                    panzerPrefs.Left = (fMain.BodyForm.Left * screenTwipsPerPixelX) - (panzerPrefs.Width + 200)
                End If
            End If
            
            If panzerPrefs.Left < 0 Then panzerPrefs.Left = 0
            If panzerPrefs.Top < 0 Then panzerPrefs.Top = 0
            
            panzerPrefs.show  ' show it again
            panzerPrefs.SetFocus
        End If
    End If
    

   On Error GoTo 0
   Exit Sub

makeProgramPreferencesAvailable_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure makeProgramPreferencesAvailable of Form menuForm"
End Sub
    

'---------------------------------------------------------------------------------------
' Procedure : readPrefsPosition
' Author    : beededea
' Date      : 28/05/2023
' Purpose   : read the form X/Y params from the toolSettings.ini
'---------------------------------------------------------------------------------------
'
Public Sub readPrefsPosition()
            
   On Error GoTo readPrefsPosition_Error

    PzEFormXPosTwips = fGetINISetting("Software\PzKill", "formXPos", PzESettingsFile)
    PzEFormYPosTwips = fGetINISetting("Software\PzKill", "formYPos", PzESettingsFile)

    ' if a current location not stored then position to the middle of the screen
    If PzEFormXPosTwips <> "" Then
        panzerPrefs.Left = Val(PzEFormXPosTwips)
    Else
        panzerPrefs.Left = screenWidthTwips / 2 - panzerPrefs.Width / 2
    End If

    If PzEFormYPosTwips <> "" Then
        panzerPrefs.Top = Val(PzEFormYPosTwips)
    Else
        panzerPrefs.Top = Screen.Height / 2 - panzerPrefs.Height / 2
    End If

   On Error GoTo 0
   Exit Sub

readPrefsPosition_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readPrefsPosition of Module Module1"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : writePrefsPosition
' Author    : beededea
' Date      : 28/05/2023
' Purpose   : save the current X and y position of this form to allow repositioning when restarting
'---------------------------------------------------------------------------------------
'
Public Sub writePrefsPosition()
        
   On Error GoTo writePrefsPosition_Error

    If panzerPrefs.WindowState = vbNormal Then ' when vbMinimised the value = -48000  !
        PzEFormXPosTwips = Str$(panzerPrefs.Left)
        PzEFormYPosTwips = Str$(panzerPrefs.Top)
        
        ' now write those params to the toolSettings.ini
        sPutINISetting "Software\PzKill", "formXPos", PzEFormXPosTwips, PzESettingsFile
        sPutINISetting "Software\PzKill", "formYPos", PzEFormYPosTwips, PzESettingsFile
    End If
    
    On Error GoTo 0
   Exit Sub

writePrefsPosition_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writePrefsPosition of Form panzerPrefs"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : settingsTimer_Timer
' Author    : beededea
' Date      : 03/03/2023
' Purpose   : Checking the date/time of the settings.ini file meaning that another tool has edited the settings
'---------------------------------------------------------------------------------------
' this has to be in a shared module and not in the prefs form as it will be running in the normal context woithout prefs showing.

Private Sub settingsTimer_Timer()

    Dim timeDifferenceInSecs As Long: timeDifferenceInSecs = 0 ' max 86 years as a LONG in secs
    Dim settingsModificationTime As Date: settingsModificationTime = #1/1/2000 12:00:00 PM#
    
    On Error GoTo settingsTimer_Timer_Error

    If Not fFExists(PzESettingsFile) Then
        MsgBox ("%Err-I-ErrorNumber 13 - FCW was unable to access the dock settings ini file. " & vbCrLf & PzESettingsFile)
        Exit Sub
    End If
    
    ' check the settings.ini file date/time
    settingsModificationTime = FileDateTime(PzESettingsFile)
    timeDifferenceInSecs = Int(DateDiff("s", oldPzESettingsModificationTime, settingsModificationTime))

    ' if the settings.ini has been modified then reload the map
    If timeDifferenceInSecs > 1 Then

        oldPzESettingsModificationTime = settingsModificationTime
        
    End If
    
    On Error GoTo 0
    Exit Sub

settingsTimer_Timer_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure settingsTimer_Timer of Form module1.bas"
            Resume Next
          End If
    End With
End Sub




'---------------------------------------------------------------------------------------
' Procedure : checkAndKill
' Author    : beededea
' Date      : 21/09/2019
' Purpose   : Find and kill any given process name
'           : This routine is an analog of checkAndKillPutWindowBehind. It is more or less identical and you should keep them in synch.
'             This version does NOT have calls to routines that require additional API calls
'             I could have used compile time references (#) to bypass these but it seemed more appropriate to create
'             separate copy for DockSettings and Enhance Icon Settings to run that it would not share with the other utilities.
'---------------------------------------------------------------------------------------
'
Public Function checkAndKill(ByRef NameProcess As String, ByVal checkForFolder As Boolean, ByVal confirmEachProcessKill As Boolean) As Boolean

    ' variables declared
    Dim AppCount As Integer: AppCount = 0
    Dim RProcessFound As Long: RProcessFound = 0
    Dim SzExename As String: SzExename = vbNullString
    Dim MyProcess As Long: MyProcess = 0
    Dim i As Integer: i = 0
    Dim binaryName As String: binaryName = vbNullString
    Dim folderName As String: folderName = vbNullString
    Dim procId As Long: procId = 0
    Dim runningProcessFolder As String: runningProcessFolder = vbNullString
    Dim processToKill As Long: processToKill = 0
    Dim ExitCode As Long: ExitCode = 0
    
    On Error GoTo checkAndKill_Error
    'If debugflg = 1 Then debugLog "%checkAndKill"

    checkAndKill = False
    MyProcess = GetCurrentProcessId()
    
    If NameProcess <> vbNullString Then
          AppCount = 0
          
          binaryName = getFileNameFromPath(NameProcess)
          If binaryName = vbNullString Then Exit Function ' catchall to prevent closure of unknown processes if the name is malformed
          
          folderName = getFolderNameFromPath(NameProcess)
          
          uProcess.dwSize = Len(uProcess)
          hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)

          'hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
          RProcessFound = ProcessFirst(hSnapshot, uProcess)
          Do
            i = InStr(1, uProcess.szexeFile, Chr(0))
            SzExename = LCase$(Left$(uProcess.szexeFile, i - 1))

            If Right$(SzExename, Len(binaryName)) = LCase$(binaryName) Then

                    AppCount = AppCount + 1
                    processToKill = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
                    If uProcess.th32ProcessID = MyProcess Then
                       'MsgBox "hmmm" & MyProcess ' we never want to kill our own process...
                    Else
                        If checkForFolder = True Then ' only check the process actual run folder when killing an app from the dock
                            procId = uProcess.th32ProcessID ' actual PID
                            runningProcessFolder = getFolderNameFromPath(getExePathFromPID(procId))
                            If LCase$(runningProcessFolder) = LCase$(folderName) Then
                                ' checkAndKill = TerminateProcess(processToKill, ExitCode)
                                ' Call CloseHandle(processToKill)
                                checkAndKill = confirmEachKill(binaryName, procId, processToKill, confirmEachProcessKill, ExitCode)
                            End If
                        Else ' just go ahead and kill whatever process I say must go
                            ' checkAndKill = TerminateProcess(processToKill, ExitCode)
                            ' Call CloseHandle(processToKill)
                            checkAndKill = confirmEachKill(binaryName, procId, processToKill, confirmEachProcessKill, ExitCode)
                        End If
                    End If
            End If
            RProcessFound = ProcessNext(hSnapshot, uProcess)
            
          Loop While RProcessFound
          Call CloseHandle(hSnapshot)
    End If


   On Error GoTo 0
   Exit Function

checkAndKill_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkAndKill of Module Common"

End Function



'
'---------------------------------------------------------------------------------------
' Procedure : getFileNameFromPath
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : A function to getFileNameFromPath
'
'---------------------------------------------------------------------------------------
'
Public Function getFileNameFromPath(ByRef strFullPath As String) As String
   On Error GoTo getFileNameFromPath_Error
      
   ' returns the remainder of the path from the final backslash which can be a file or a folder
   If Not fFExists(strFullPath) Then ' tests to see if a file or a folder of the same name in the same location
        getFileNameFromPath = strFullPath   ' if a file does not exist then what remains must be a folder or just a binary name without a folder
        Exit Function               ' if a file does exist get its name below
   End If
   getFileNameFromPath = Right$(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))

   On Error GoTo 0
   Exit Function

getFileNameFromPath_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getFileNameFromPath of Module Common"
End Function

'---------------------------------------------------------------------------------------
' Procedure : getFolderNameFromPath
' Author    : beededea
' Date      : 11/07/2019
' Purpose   : get the folder or directory path as a string not including the last backslash
'---------------------------------------------------------------------------------------
'
Public Function getFolderNameFromPath(ByRef Path As String) As String

   On Error GoTo getFolderNameFromPath_Error
   'If debugflg = 1 Then debugLog "%" & "getFolderNameFromPath"

    If InStrRev(Path, "\") = 0 Then
        getFolderNameFromPath = vbNullString
        Exit Function
    End If
    getFolderNameFromPath = Left$(Path, InStrRev(Path, "\") - 1)

   On Error GoTo 0
   Exit Function

getFolderNameFromPath_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getFolderNameFromPath of Module Common"
End Function


'---------------------------------------------------------------------------------------
' Procedure : getExePathFromPID
' Author    : beededea
' Date      : 25/08/2020
' Purpose   : getting the full path of a running process is not as easy as you'd expect
'---------------------------------------------------------------------------------------
'
Public Function getExePathFromPID(ByVal idProc As Long) As String
    Dim sBuf As String:  sBuf = vbNullString
    Dim sChar As Long: sChar = 0
    Dim useloop As Integer: useloop = 0
    Dim hProcess As Long: hProcess = 0
    
    On Error GoTo getExePathFromPID_Error

    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, idProc)
    If hProcess Then
        sBuf = String$(260, vbNullChar)
        sChar = GetProcessImageFileName(hProcess, sBuf, 260)
        If sChar Then
            sBuf = NoNulls(sBuf)
            ' this loop replaces the internal windows volume name with the legacy naming convention, ie. C:\, D:\ &c
            For useloop = 1 To lstDevicesListCount
                If InStr(1, sBuf, lstDevices(1, useloop)) > 0 Then
                    sBuf = Replace(sBuf, lstDevices(1, useloop), Chr$(lstDevices(0, useloop)) & ":")
                    Exit For
                End If
            Next useloop
            getExePathFromPID = sBuf
        End If
        CloseHandle hProcess
    End If

   On Error GoTo 0
   Exit Function

getExePathFromPID_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getExePathFromPID of Module common"
End Function
'---------------------------------------------------------------------------------------
' Procedure : confirmEachKill
' Author    : beededea
' Date      : 20/12/2022
' Purpose   : This routine is an analog of confirmEachKillPutWindowBehind. It is more or less identical and you should keep them in synch.
'             This version does NOT have calls to routines that require additional API calls
'             I could have used compile time references (#) to bypass these but it seemed more appropriate to create
'             separate copy for DockSettings and Enhance Icon Settings to run that it would not share with the other utilities.
'---------------------------------------------------------------------------------------
'
Public Function confirmEachKill(ByVal binaryName As String, ByVal procId As Long, ByVal processToKill As String, ByVal confirmEachProcessKill As Boolean, ByRef ExitCode As Long) As Boolean
    Dim goAheadAndKill As Boolean: goAheadAndKill = False
    Dim rmessage As String: rmessage = ""
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim a As Long

    On Error GoTo confirmEachKill_Error

    If confirmEachProcessKill = True Then
        rmessage = "A matching process has been found. Kill this application? - " & binaryName & " with process ID " & procId
        'nswer = MsgBox(rmessage, vbYesNo)
        'answer = msgBoxA(rmessage, vbYesNo, "Killing this application", True, "confirmEachKill")
        answer = MsgBox(rmessage, vbExclamation + vbYesNo)

        If answer = vbNo Then
            goAheadAndKill = False
        Else
            goAheadAndKill = True
        End If
    Else
        goAheadAndKill = True
    End If
    
    If goAheadAndKill = True Then
        confirmEachKill = TerminateProcess(processToKill, ExitCode)
        Call CloseHandle(processToKill)
    End If

    On Error GoTo 0
    Exit Function

confirmEachKill_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure confirmEachKill of Module common"
            Resume Next
          End If
    End With
End Function


'---------------------------------------------------------------------------------------
' Procedure : NoNulls
' Author    : beededea
' Date      : 25/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function NoNulls(ByVal Strng As String) As String
    Dim i As Integer: i = 0
    On Error GoTo NoNulls_Error

    If Len(Strng) > 0 Then
        i = InStr(Strng, vbNullChar)
        Select Case i
            Case 0
                NoNulls = Strng
            Case 1
                NoNulls = vbNullString
            Case Else
                NoNulls = Left$(Strng, i - 1)
        End Select
    End If

   On Error GoTo 0
   Exit Function

NoNulls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure NoNulls of Module common"
End Function
