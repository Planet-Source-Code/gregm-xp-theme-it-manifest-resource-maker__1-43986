Attribute VB_Name = "basFunctions"
Option Explicit

Public Const MAX_PATH As Long = 260

Public Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersion As Long
    dwFileVersionMS As Long
    dwFileVersionLS As Long
    dwProductVersionMS As Long
    dwProductVersionLS As Long
    dwFileFlagsMask As Long
    dwFileFlags As Long
    dwFileOS As Long
    dwFileType As Long
    dwFileSubtype As Long
    dwFileDateMS As Long
    dwFileDateLS As Long
End Type

Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" _
    (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
              
Type RECT
        left As Long
        tOp As Long
        Right As Long
        Bottom As Long
End Type

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
    szCSDVersion As String * 128
End Type

Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias _
              "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" _
              (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, _
              lpData As Any) As Long
              
Type DLLVERSIONINFO
   cbSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformID As Long
End Type

Declare Function DllGetCtl32Version Lib "comctl32.dll" Alias "DllGetVersion" (pdvi As DLLVERSIONINFO) As Long

Public Declare Function IsAppThemed Lib "uxtheme.dll" () As Long
Public Declare Function IsThemeActive Lib "uxtheme.dll" () As Long

Private Const WM_USER = &H400

' SetWIndowPos
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2

Enum OperatingSystems
    WindowsAny
    WindowsNT351
    WindowsNT4
    Windows95
    Windows98
    Windows2000
    WindowsNT5
End Enum

Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nBufferLength As Long) As Long
Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As Long) As Long
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (OSInfo As OSVERSIONINFO) As Long

' Version of Windows running for compatibility
Declare Function GetOSVersion Lib "kernel32" Alias "GetVersion" () As Long
Declare Function GetMilliseconds Lib "kernel32" Alias "GetTickCount" () As Long

Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" _
    (ByVal lpszShortPath As String, _
    ByVal lpszLongPath As String, _
    ByVal cchBuffer As Long) As Long
    
Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" ( _
    ByVal lpszLongPath As String, _
    ByVal lpszShortPath As String, _
    ByVal cchBuffer As Long) As Long

Declare Function OpenIcon Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd&, ByVal hWndInsertAfter&, ByVal X&, ByVal Y&, ByVal cx&, ByVal cy&, ByVal wFlags&) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd&, ByVal lpClassName$, ByVal nMaxCount&) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long

Private Const ICCC_LISTVIEW_CLASSES As Long = &H1 'listview, header
Private Const ICCC_TREEVIEW_CLASSES As Long = &H2 'treeview, tooltips
Private Const ICCC_BAR_CLASSES As Long = &H4      'toolbar, statusbar, trackbar, tooltips
Private Const ICCC_TAB_CLASSES As Long = &H8      'tab, tooltips
Private Const ICCC_UPDOWN_CLASS As Long = &H10    'updown
Private Const ICCC_PROGRESS_CLASS As Long = &H20  'progress
Private Const ICCC_HOTKEY_CLASS As Long = &H40    'hotkey
Private Const ICCC_ANIMATE_CLASS As Long = &H80   'animate
Private Const ICCC_WIN95_CLASSES As Long = &HFF   'everything else
Private Const ICCC_DATE_CLASSES As Long = &H100   'month picker, date picker, time picker, updown
Private Const ICCC_USEREX_CLASSES As Long = &H200 'comboex
Private Const ICCC_COOL_CLASSES As Long = &H400   'rebar (coolbar) control

'WIN32_IE >= 0x0400
Private Const ICCC_INTERNET_CLASSES As Long = &H800
Private Const ICCC_PAGESCROLLER_CLASS As Long = 1000 'page scroller
Private Const ICCC_NATIVEFNTCTL_CLASS As Long = 2000 'native font control

'WIN32_WINNT >= 0x501
Private Const ICCC_STANDARD_CLASSES As Long = 4000
Private Const ICCC_LINK_CLASS As Long = 8000

Private Const ICCC_ALL_CLASSES As Long = ICCC_ANIMATE_CLASS Or ICCC_BAR_CLASSES _
        Or ICCC_COOL_CLASSES Or ICCC_DATE_CLASSES Or ICCC_HOTKEY_CLASS _
        Or ICCC_LISTVIEW_CLASSES Or ICCC_PROGRESS_CLASS Or ICCC_TAB_CLASSES _
        Or ICCC_TREEVIEW_CLASSES Or ICCC_UPDOWN_CLASS _
        Or ICCC_USEREX_CLASSES Or ICCC_WIN95_CLASSES

Enum INITCommonControlsClasses
    ICC_ANIMATE_CLASS& = ICCC_ANIMATE_CLASS
    ICC_BAR_CLASSES& = ICCC_BAR_CLASSES
    ICC_COOL_CLASSES& = ICCC_COOL_CLASSES
    ICC_DATE_CLASSES& = ICCC_DATE_CLASSES
    ICC_HOTKEY_CLASS& = ICCC_HOTKEY_CLASS
    ICC_INTERNET_CLASSES& = ICCC_INTERNET_CLASSES
    ICC_LINK_CLASS& = ICCC_LINK_CLASS
    ICC_LISTVIEW_CLASSES& = ICCC_LISTVIEW_CLASSES
    ICC_NATIVEFNTCTL_CLASS& = ICCC_NATIVEFNTCTL_CLASS
    ICC_PAGESCROLLER_CLASS& = ICCC_PAGESCROLLER_CLASS
    ICC_PROGRESS_CLASS& = ICCC_PROGRESS_CLASS
    ICC_STANDARD_CLASSES& = ICCC_STANDARD_CLASSES
    ICC_TAB_CLASSES& = ICCC_TAB_CLASSES
    ICC_TREEVIEW_CLASSES& = ICCC_TREEVIEW_CLASSES
    ICC_UPDOWN_CLASS& = ICCC_UPDOWN_CLASS
    ICC_USEREX_CLASSES& = ICCC_USEREX_CLASSES
    ICC_WIN95_CLASSES& = ICCC_WIN95_CLASSES
    ICC_ALL_CLASSES& = ICCC_ALL_CLASSES
        
End Enum

Private Type tagINITCOMMONCONTROLSEX
   dwSize As Long
   dwICC As Long
End Type

Public Declare Sub InitCommonControls Lib "comctl32.dll" ()

Public Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean
   
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
    
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As Any, ByVal lpFile As String, ByVal lpParameters As Any, ByVal lpDirectory As Any, ByVal nShowCmd As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpexitcode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" _
        (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, _
         ByVal dwProcId As Long) As Long
         
' changes the position and dimensions of the specified window relative
' to the upper-left corner of the screen or it's parent window's client area.
Declare Function MoveWindow Lib "user32" ( _
    ByVal hWnd As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal bRepaint As Long) As Long

'  variables for ofnHookCallback
Private OFN_PROC_BtnOkText As String
Private OFN_PROC_BtnCancelText As String
Private OFN_PROC_dlgCenter As Boolean
Private OFN_PROC_dlgDrivesLabel As String
Private OFN_PROC_FiltersLabel As String
Private OFN_PROC_FileSpecLabel As String
Private OFN_PROC_CheckboxLabel As String

Public Enum OpenFileNameFlags
    ' Specifies that the File Name list box allows multiple selections. If you also
    ' set the OFN_EXPLORER flag, the dialog box uses the Explorer-style user
    ' interface; otherwise, it uses the old-style user interface.
    OFN_ALLOWMULTISELECT = &H200

    ' If the user specifies a file that does not exist, this flag causes the
    ' dialog box to prompt the user for permission to create the file.
    OFN_CREATEPROMPT = &H2000

    ' Indicates that the sTemplateName member is a pointer to the
    ' name of a dialog template resource in the module identified by
    ' the hInstance member.
    OFN_ENABLETEMPLATE = &H40

    ' Indicates that the hInstance member identifies a data block that
    ' contains a preloaded dialog box template. The system ignores
    ' the sTemplateName if this flag is specified.
    OFN_ENABLETEMPLATEHANDLE = &H80

    ' Indicates that any customizations made to the Open or
    ' Save As dialog box use the new Explorer-style customization methods.
    OFN_EXPLORER = &H80000

    ' Specifies that the user typed a file name extension that differs
    ' from the extension specified by sDefFileExt.
    OFN_EXTENSIONDIFFERENT = &H400

    ' Specifies that the user can type only names of existing files
    ' in the File Name entry field.
    OFN_FILEMUSTEXIST = &H1000

    ' Hides the Read Only check box.
    OFN_HIDEREADONLY = &H4

    ' For old-style dialog boxes, this flag causes the
    ' dialog box to use long file names.
    OFN_LONGNAMES = &H200000

    ' Restores the current directory to its original value if the
    ' user changed the directory while searching for files.
    OFN_NOCHANGEDIR = &H8

    ' Directs the dialog box to return the path and file name of the
    ' selected shortcut (.LNK) file. If this value is not specified, the dialog box
    ' returns the path and file name of the file referenced by the shortcut.
    OFN_NODEREFERENCELINKS = &H100000

    ' For old-style dialog boxes, this flag causes the dialog box to use
    ' short file names (8.3 format).
    OFN_NOLONGNAMES = &H40000

    ' Hides and disables the Network button.
    OFN_NONETWORKBUTTON = &H20000

    ' Specifies that the returned file does not have the Read Only check
    ' box selected and is not in a write-protected directory.
    OFN_NOREADONLYRETURN = &H8000&

    ' Specifies that a test file is not created before the dialog box is closed.
    OFN_NOTESTFILECREATE = &H10000

    ' Specifies that the common dialog boxes allow invalid
    ' characters in the returned file name.
    OFN_NOVALIDATE = &H100

    ' Causes the Save As dialog box to generate a message
    ' box if the selected file already exists.
    OFN_OVERWRITEPROMPT = &H2

    ' Specifies that the user can type only valid paths and file names.
    OFN_PATHMUSTEXIST = &H800

    ' Causes the Read Only check box to be selected initially when
    ' the dialog box is created.
    OFN_READONLY = &H1

    ' Causes the dialog box to display the Help button. The hwndOwner
    ' member must specify the window to receive the HELPMSGSTRING
    ' registered messages that the dialog box sends when the user
    ' clicks the Help button.
    OFN_SHOWHELP = &H10

    ' Enables the hook function specified in the fnHook member.
    OFN_ENABLEHOOK = &H20

    ' Enables the Explorer-style dialog box to be resized using either
    ' the mouse or the keyboard.
    OFN_ENABLESIZING = &H800000

    ' Specifies that if a call to the OpenFile function fails because of a
    ' network sharing violation, the error is ignored and the dialog box
    ' returns the selected file name.
    OFN_SHAREAWARE = &H4000
    '-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-
    '-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-
    'Exclusive to Windows 2000/NT5
    '-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-
    ' Prevents the system from adding a link to the selected
    ' file in the file system directory that contains the user's most recently
    ' used documents.
    OFN_DONTADDTORECENT = &H2000000

    ' Causes the dialog box to send CDN_INCLUDEITEM notification messages
    ' to your ofnHookProc hook procedure when the user opens a folder.
    OFN_ENABLEINCLUDENOTIFY = &H400000

    ' Forces the showing of files with attributes marked exclusively as system or hidden.
    OFN_FORCESHOWHIDDEN = &H10000000

    ' InitFlagsEx value::
    ' Hides the places bar containing icons for commonly-used folders, such
    ' as Favorites and Desktop. This flag is set to the InitFlagsEx member
    ' of OPENFILENAME structure
    OFN_EX_NOPLACESBAR = &H1

End Enum

'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-
'
Private Type OPENFILENAME
    ' Specifies the length, in bytes, of the structure.
    nStructSize       As Long

    ' Handle to the window that owns the dialog box.
    hwndOwner         As Long

    ' hInstance is a handle to a memory object containing a dialog box template.
    hInstance         As Long

    ' Pointer to a buffer containing pairs of null-terminated filter strings.
    sFilter           As String

    ' Pointer to a static buffer that contains a pair of null-terminated filter strings
    ' for preserving the filter pattern chosen
    sCustomFilter     As String

    ' Specifies the size of the buffer identified by sCustomFilter.
    ' (Bytes) for the ANSI; (characters) for the Unicode version.
    nMaxCustFilter    As Long

    ' Specifies the index of the currently selected filter in
    ' the File Types control.
    nFilterIndex      As Long

    ' Pointer to a buffer that contains a file name used
    ' to initialize the File Name edit control.
    sFile             As String

    ' Specifies the size of the buffer pointed to by sFile.
    ' (Bytes) for the ANSI; (characters) for the Unicode version.
    nMaxFile          As Long

    ' Pointer to a buffer that receives the file name and extension
    ' (without path information) of the selected file.
    sFileTitle        As String

    ' Specifies the size of the buffer pointed to by sFileTitle.
    ' (Bytes) for the ANSI; (characters) for the Unicode version.
    nMaxTitle         As Long

    ' Pointer to a null terminated string that can
    ' specify the initial directory.
    sInitDir       As String

    ' Pointer to a custom string to be placed in the title bar of the dialog box.
    sDialogTitle      As String

    ' OFN bit flags you can use to initialize the dialog box.
    InitFlags             As OpenFileNameFlags

    ' Specifies the zero-based offset from the beginning of
    ' the path to the file name in the string pointed to by sFile.
    ' (Bytes) for the ANSI; (characters) for the Unicode version.
    nFileOffset       As Integer

    ' Specifies the zero-based offset from the beginning of
    ' the path to the file name extension in the string pointed to by sFile.
    ' (Bytes) for the ANSI; (characters) for the Unicode version.
    nFileExtension    As Integer

    ' Pointer to a buffer that contains the default extension.
    sDefFileExt       As String

    ' Specifies application-defined data that the system passes to
    ' the hook procedure identified by the fnHook member.
    nCustData         As Long

    ' Pointer to a hook procedure if OFN_ENABLEHOOK flag is set.
    fnHook            As Long

    ' Pointer to a null-terminated string that names a dialog template
    ' resource in the module identified by the hInstance member.
    sTemplateName     As String

'' NEW Windows 2000 structure members
    pvReserved        As Long
    dwReserved        As Long
    InitFlagsEx           As Long ' Additional bit flag member
End Type

' User defined structure containing information about a notification message.
Private Type LVNOTIFY_short
    hwndFrom As Long ' not used
    idFrom As Long ' not used
    code As Long ' notification code for the common dialog
End Type

Private Enum CommonDialogMessages
    CDM_FIRST = (WM_USER + 100)
    CDM_LAST = (WM_USER + 200)

    ' retrieves the file name (not including the path) of the currently selected file
    CDM_GETSPEC = CDM_FIRST + &H0
    
    ' retrieves the path and file name of the selected file
    CDM_GETFILEPATH = CDM_FIRST + &H1
    
    ' retrieves the path of the currently open folder or directory
    CDM_GETFOLDERPATH = CDM_FIRST + &H2
    
    ' retrieves a pointer to the list of item identifiers corresponding
    ' to the folder currently opened in the dialog
    CDM_GETFOLDERIDLIST = CDM_FIRST + &H3
    
    ' Sets the text for a specified control in the dialog box
    CDM_SETCONTROLTEXT = CDM_FIRST + &H4
    
    ' hides the specified control in the dialog box
    CDM_HIDECONTROL = CDM_FIRST + &H5
    
    ' sets the default file name extension
    CDM_SETDEFEXT = CDM_FIRST + &H6
End Enum

' Notifications when Open or Save dialog status changes
Private Enum CommonDialogNotificationMessages
   CDN_FIRST = -601
   CDN_LAST = -699
   
   ' message sent to the ofnHookProc when the all controls in
   ' the dialog box have been positioned.
   CDN_INITDONE = CDN_FIRST - &H0
   
   ' The user selected a new item from the file list.
   CDN_SELCHANGE = CDN_FIRST - &H1
   
   ' A new folder or directory was opened.
   CDN_FOLDERCHANGE = CDN_FIRST - &H2
   
   ' A sharing violation was encountered
   ' on the file about to be returned.
   CDN_SHAREVIOLATION = CDN_FIRST - &H3
   
   ' The help button was clicked
   CDN_HELP = CDN_FIRST - &H4
   
   ' User clicked OK, selected filename is about to be returned
   CDN_FILEOK = CDN_FIRST - &H5
   
   ' The user selected a new file type from the list of file types
   CDN_TYPECHANGE = CDN_FIRST - &H6
   
   ' Sent for each item added to the list. Win2000 ONLY, Ignored if the
   ' OFN_ENABLEINCLUDENOTIFY is not set.
   CDN_INCLUDEITEM = CDN_FIRST - &H7 '
End Enum

' IDs for specified controls in an Explorer-style Open or Save common dialog
Private Enum CommonDialogControlIDs
    CDC_IDOK = &H1 ' Open/Save/OK Button
    CDC_IDCANCEL = &H2 ' Cancel Button
    CDC_IDCONTENTS = &H460   ' Listbox Contents
    CDC_IDCONTENTSLABEL = &H440    ' Listbox Label
    CDC_IDFILETYPELABEL = &H441   ' File/Objects of Type
    CDC_IDFILEEDITLABEL = &H442   ' File/Object Name
    CDC_IDDRIVESLABEL = &H443     ' Look In/Save to
    CDC_IDREADONLY = &H410    ' ReadOnly Checkbox
    CDC_IDFILETYPE = &H470   ' File Types combo
    CDC_IDDRIVES = &H471     '  Drives combo
    CDC_IDFILEEDIT = &H480    ' File Editbox Control
    CDC_IDHELP = &H40E ' Help Command Button
End Enum

' function that fills a block of memory with zeros
Private Declare Sub ZeroMem Lib "kernel32" Alias "RtlZeroMemory" ( _
    Destination As Any, ByVal length As Long)

' creates an Open common dialog box that lets the user specify the drive, directory, and the
' name of a file or set of files to open
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" ( _
    pOpenfilename As OPENFILENAME) As Long

' creates a Save common dialog box that lets the user specify the
' drive, directory, and name of a file to save.
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" ( _
    pOpenfilename As OPENFILENAME) As Long




Public Sub Wait(ByVal Milliseconds As Long)
Dim MaxTimer As Long

    On Local Error GoTo ERR_Wait
    
    MaxTimer = (GetMilliseconds + Milliseconds)
    Do: DoEvents: Loop While (MaxTimer > GetMilliseconds)

Done:
Err.Clear
Exit Sub

ERR_Wait:
Resume Done
End Sub

Public Function ShellExec(ByVal szFileName As String, _
                                    Optional ByVal WinState As VbAppWinStyle = vbNormalFocus, _
                                    Optional ByVal ExitCodeProcessDelay As Long, _
                                    Optional ByVal ResolveAssociation As Boolean) As Boolean

Dim lShell As Boolean, hShell As Long
Dim hProc As Long, lProc As Long, SafeDelay As Long
Dim lpszShortPath As String, sShortFileName As String
Dim ResumeError As Boolean
Const SE_ERR_NOASSOC As Long = 31
Const PROCESS_QUERY_INFORMATION = &H400
Const STILL_ACTIVE = &H103

     On Error GoTo ShellErr
    
    If ExistFile(szFileName) Then
        ' Get short path name
        lpszShortPath = String(MAX_PATH, Chr(0))
        sShortFileName = left$(lpszShortPath, GetShortPathName(szFileName, lpszShortPath, MAX_PATH))
        
    End If
    
    ResumeError = True
    hShell = Shell(szFileName, WinState)
    ResumeError = False
    
ExitCodeProcess:

    If ExitCodeProcessDelay <> 0 Then
        SafeDelay = GetMilliseconds
        hProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, hShell)
        Do
            GetExitCodeProcess hProc&, lProc&
        Loop While lProc& = STILL_ACTIVE And GetMilliseconds() < (SafeDelay + ExitCodeProcessDelay)
    
        If hShell <> 0 Then lShell = True
        ShellExec = lShell
    Else
        ShellExec = True
    End If
    
ShellErr_Exit:
Err.Clear
Exit Function

ShellErr:
If ResumeError Then
    ResumeError = False
    hShell = 0
    Err.Clear
    
    If Len(sShortFileName) <> 0 Then
        ' Instance handle over 32 is returned, else error number is returned.
        hShell = ShellExecute(GetDesktopWindow(), "Open", sShortFileName, vbNullString, vbNullString, WinState)
    End If
    Select Case hShell
        Case SE_ERR_NOASSOC
            If ResolveAssociation Then
                 On Error Resume Next
                Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & szFileName, vbNormalFocus)
                ShellExec = True
            Else
                ShellExec = False
            End If
        Case Is <= 32
            ShellExec = False
        Case Else
            GoTo ExitCodeProcess
    End Select
    GoTo ShellErr_Exit
Else
    ShellExec = False
    Resume ShellErr_Exit
End If
End Function
'CAUTION: This function is called by the system to initialize the GetOpenFileName
' or GetSaveFileName dialog box. Attempting to set breakpoints or adding other
'debugging code to this routine may cause unexpected problems.
Private Function OFNHookProc(ByVal hWnd As Long, _
                            ByVal uMsg As Long, _
                            ByVal wParam As Long, _
                            ByVal lParam As Long) As Long
                                     
Static hWndParent As Long
Dim blnRetVal As Boolean
Const WM_NOTIFY = &H4E
Const WM_NCDESTROY = &H82
Const WM_INITDIALOG = &H110

 On Error Resume Next

    Select Case uMsg
        ' On initialization, set aspects of the
        ' dialog that are not obtainable through
        ' manipulating the OPENFILENAME structure members.
        Case WM_INITDIALOG
        
            ' Obtain the handle to the parent dialog
            hWndParent = GetParent(hWnd)
             
            ' Center the dialog box
            If OFN_PROC_dlgCenter Then
                blnRetVal = OFNCenterDialog(hWndParent)
            End If
    
        Case WM_NOTIFY
        
            Dim LVN As LVNOTIFY_short
            Call ZeroMem(LVN, Len(LVN))
            Call CopyMemory(LVN, ByVal lParam, Len(LVN))
            
            Select Case LVN.code
                Case CDN_INITDONE
                
                    'Debug.Print "CDN_INITDONE"
                    
                    If Len(OFN_PROC_FileSpecLabel) <> 0 Then
                        ' change caption to the file edit box label
                        Call SendMessage(hWndParent, CDM_SETCONTROLTEXT, _
                        CDC_IDFILEEDITLABEL, ByVal OFN_PROC_FileSpecLabel)
                    End If
                    
                    If Len(OFN_PROC_FiltersLabel) <> 0 Then
                        ' change caption to file type label
                        Call SendMessage(hWndParent, CDM_SETCONTROLTEXT, _
                        CDC_IDFILETYPELABEL, ByVal OFN_PROC_FiltersLabel)
                    End If
                    
                    If Len(OFN_PROC_dlgDrivesLabel) <> 0 Then
                        ' change caption to drives label
                        Call SendMessage(hWndParent, CDM_SETCONTROLTEXT, _
                        CDC_IDDRIVESLABEL, ByVal OFN_PROC_dlgDrivesLabel)
                    End If
                    
                    If Len(OFN_PROC_BtnOkText) <> 0 Then
                        ' change text to OK Button
                        Call SendMessage(hWndParent, CDM_SETCONTROLTEXT, _
                        CDC_IDOK, ByVal OFN_PROC_BtnOkText)
                    End If
                    
                    If Len(OFN_PROC_BtnCancelText) <> 0 Then
                        ' change text to Cancel Button
                        Call SendMessage(hWndParent, CDM_SETCONTROLTEXT, _
                        CDC_IDCANCEL, ByVal OFN_PROC_BtnCancelText)
                    End If
                    
                    If Len(OFN_PROC_CheckboxLabel) <> 0 Then
                        ' change caption to read-only checkbox label
                        Call SendMessage(hWndParent, CDM_SETCONTROLTEXT, _
                        CDC_IDREADONLY, ByVal OFN_PROC_CheckboxLabel)
                    End If
                
                Case CDN_INCLUDEITEM: 'Debug.Print "CDN_INCLUDEITEM"
                    ' message sent for each item added to list if
                    ' OFN_ENABLEINCLUDENOTIFY bit flag is set
                Case CDN_FOLDERCHANGE: 'Debug.Print "CDN_FOLDERCHANGE"
                Case CDN_SELCHANGE: 'Debug.Print "CDN_SELCHANGE"
                Case CDN_FILEOK:: 'Debug.Print "CDN_FILEOK"
                    ' Clicked OK:
                Case CDN_SHAREVIOLATION: 'Debug.Print "CDN_SHAREVIOLATION"
                    ' sharing violation occurred with selected file
                Case CDN_HELP: 'Debug.Print "CDN_HELP"
                    ' Help clicked
                Case CDN_TYPECHANGE: 'Debug.Print "CDN_TYPECHANGE"
                    ' File type changed
                Case Else: 'Debug.Print "ELSE"
            End Select
            CopyMemory ByVal lParam, 0&, 0&
        
        Case WM_NCDESTROY
            ' message sent indicating dialog is closing
        
    End Select

    OFNHookProc = Abs(blnRetVal)

End Function

Private Function GetProc(ByVal pfn As Long) As Long
  
  'Dummy procedure that receives and returns
  'the return value of the AddressOf operator.
 
  'Obtain and set the address of the callback
  'This workaround is needed as you can't assign
  'AddressOf directly to a member of a user-
  'defined type, but you can assign it to another
  'long and use that (as returned here)
 
  GetProc = pfn

End Function

Public Function GetOpenSaveFileName(InitFile As String, _
                                        Optional InitFolder As String, _
                                        Optional DialogTitle As String, _
                                        Optional GetSaveFile As Boolean, _
                                        Optional Filter As String = "All (*.*)| *.*", _
                                        Optional DefaultExt As String, _
                                        Optional FilterIndex As Long = 1, _
                                        Optional FileExistPrompt As Boolean, _
                                        Optional ReadOnly As Boolean, _
                                        Optional AllowMultiSelect As Boolean, _
                                        Optional NoCurDirChange As Boolean, _
                                        Optional HideReadOnly As Boolean = True, _
                                        Optional hOwner As Long, _
                                        Optional ofnFlags As OpenFileNameFlags, _
                                        Optional DlgCentered As Boolean = True, _
                                        Optional EnableResizing As Boolean, _
                                        Optional HidePlacesBar As Boolean, _
                                        Optional DlgDrivesLabel, _
                                        Optional DlgFileSpecLabel, _
                                        Optional DlgFiltersLabel, _
                                        Optional DlgCheckboxLabel, _
                                        Optional DlgOkBtnText, _
                                        Optional DlgCancelBtnText) As Boolean

Dim OFN As OPENFILENAME, OFNew As OPENFILENAME
Dim buff As String, ofnResult As Long, EnableHook As Boolean
Dim flags As OpenFileNameFlags, nMax As Long, PathSelected As Boolean
Dim ThisFolder As String, ThisFile As String, msg As String, gosFiles As String

REOPEN:
PathSelected = False
OFN = OFNew
ThisFolder = "": ThisFile = ""
flags = 0: nMax = 0: ofnResult = 0

Select Case True
    Case DlgCentered, (Not IsMissing(DlgCancelBtnText)), (Not IsMissing(DlgOkBtnText)), _
        (Not IsMissing(DlgDrivesLabel)), (Not IsMissing(DlgFiltersLabel)), _
        (Not IsMissing(DlgFileSpecLabel)), (Not IsMissing(DlgCheckboxLabel))
            EnableHook = True
    Case Else: EnableHook = False
End Select

If EnableHook Then
    ' Set the custom text strings for ofnHookProc
    OFN_PROC_BtnCancelText = IIf(IsMissing(DlgCancelBtnText), "Cancel", CStr(DlgCancelBtnText))
    OFN_PROC_FileSpecLabel = IIf(IsMissing(DlgFileSpecLabel), "File name:", CStr(DlgFileSpecLabel))
    OFN_PROC_CheckboxLabel = IIf(IsMissing(DlgCheckboxLabel), "Open as read-only", CStr(DlgCheckboxLabel))
    OFN_PROC_dlgCenter = DlgCentered
End If

    With OFN
        ' Since the last 3 members of the OPENFILENAME structure require
        ' Win2000, we must adjust the size of the structure if current OS is Win95/98.
        .nStructSize = Len(OFN) - (IIf(CBool(OS(Windows2000)), 0, 12))

        ' Handle from any window can own the dialog box.
        ' can be null if hOwner is not specified.
        If hOwner <> 0 Then .hwndOwner = hOwner

        ' To make Windows-style filter, replace | and : with nulls
         ' Filter spec is separated by NullChars
        Dim char As String, i As Integer
        For i = 1 To Len(Filter)
            char = Mid$(Filter, i, 1)
            buff = IIf((InStr("|:", char) <> 0), (buff & vbNullChar), buff & char)
        Next
        .sFilter = buff & vbNullChar & vbNullChar
        .nFilterIndex = FilterIndex
        .sDefFileExt = DefaultExt & vbNullChar & vbNullChar
        
        ' text to be displayed in the title bar of the common dialog box
        .sDialogTitle = DialogTitle
        
        ' This buffer is filled with the selected (File+Ext) when dialog closes
        .sFileTitle = Space$(MAX_PATH) & vbNullChar & vbNullChar
        .nMaxTitle = Len(.sFileTitle)

        ' Adjust buffer size to accommodate multiselected files
        nMax = IIf(CBool(ofnFlags And (-AllowMultiSelect * OFN_ALLOWMULTISELECT)), 8192, MAX_PATH)
        
        '
        ' Win95/98: If .sInitDir contains a path, that path is the initial directory.
        '                 Otherwise, .sFile containing a path is used as initial directory.
        ' Win2000:  If .sFile contains a path, that path is used as the initial directory.
        '                 Otherwise, .sInitDir is used as initial directory.
        '
        ' prepare InitFolder
        If Len(InitFolder) > 3 Then InitFolder = QualifyPath(InitFolder, False)
        ' If InitFile parameter is empty, first character must be null before
        ' padding the .sFile buffer.
        .sInitDir = InitFolder & vbNullChar & vbNullChar
        buff = IIf((Len(InitFile) <> 0), InitFile, vbNullChar)
        .sFile = buff & Space$(nMax - Len(buff)) & vbNullChar & vbNullChar
        .nMaxFile = Len(.sFile)

        ' New for Win2000: this bit flag if set, hides the places bar
        ' normally located on left side of dialog box
        If CBool(OS(Windows2000)) Then .InitFlagsEx = (-(HidePlacesBar * OFN_EX_NOPLACESBAR))
        
        ' Enable Callback hook for common dialog box to handle event messages.
        ' Although This hookproc is written exclusively for use with the accompanied demo
        ' program, It's fairly easy customizing this function for your own use.
        If EnableHook Then .fnHook = GetProc(AddressOf OFNHookProc)

        ' Some common bit flags to be switched on/off.
        flags = ofnFlags Or _
            OFN_EXPLORER Or OFN_LONGNAMES Or _
            OFN_NOTESTFILECREATE Or _
            (-(NoCurDirChange * OFN_NOCHANGEDIR)) Or _
            (-(EnableResizing * OFN_ENABLESIZING)) Or _
            (-(EnableHook * OFN_ENABLEHOOK)) Or _
            (-(HideReadOnly * OFN_HIDEREADONLY)) Or _
            (-((CBool(OS(Windows2000)) And EnableHook) * OFN_ENABLEINCLUDENOTIFY))
        
        ' Determine additional bit flags and settings based on the type of dialog displayed.
        If GetSaveFile Then
            OFN_PROC_FiltersLabel = IIf(IsMissing(DlgFiltersLabel), "Save as type:", CStr(DlgFiltersLabel))
            OFN_PROC_BtnOkText = IIf(IsMissing(DlgOkBtnText), "Save", CStr(DlgOkBtnText))
            OFN_PROC_dlgDrivesLabel = IIf(IsMissing(DlgDrivesLabel), "Save to:", CStr(DlgDrivesLabel))
            .InitFlags = flags Or (-(FileExistPrompt * OFN_OVERWRITEPROMPT))
           ofnResult = GetSaveFileName(OFN)
        Else
            OFN_PROC_FiltersLabel = IIf(IsMissing(DlgFiltersLabel), "Files of type:", CStr(DlgFiltersLabel))
            OFN_PROC_BtnOkText = IIf(IsMissing(DlgOkBtnText), "Open", CStr(DlgOkBtnText))
            OFN_PROC_dlgDrivesLabel = IIf(IsMissing(DlgDrivesLabel), "Look in:", CStr(DlgDrivesLabel))
            .InitFlags = flags Or (-(ReadOnly * OFN_READONLY)) Or _
                (-(CBool(InitFile = "path") * OFN_NOVALIDATE)) Or _
                (-(AllowMultiSelect * OFN_ALLOWMULTISELECT)) Or _
                (-(FileExistPrompt * OFN_FILEMUSTEXIST))
            ofnResult = GetOpenFileName(OFN)
        End If

        Select Case ofnResult

            Case 1: ' Success:
                
                GetOpenSaveFileName = True
                
                ' Strip nullchars from end of buffer
                buff = RTrim$(left$(.sFile, Len(.sFile) - 2))
                
                If Not HideReadOnly Then
                    ReadOnly = CBool((OFN.InitFlags And OFN_READONLY) = OFN_READONLY)
                End If
                
                FilterIndex = .nFilterIndex
                
                ' If multiple files were selected, prepare buffer for return of files names
                If AllowMultiSelect And Mid$(buff, .nFileOffset, 1) = vbNullChar Then
                
                        ' Strip and Return folderpath contained in the buffer
                        ThisFolder = QualifyPath(CropString(buff, , True), True)
                    
                        ' Return buffer containing all selected files
                        gosFiles = buff
                        
                        If FileExistPrompt Then
                            While Len(gosFiles) <> 0
                                ThisFile = CropString(gosFiles, , True)
                                If Not ExistFile(ThisFolder & ThisFile) Then
                                    msg = "'" & ThisFolder & ThisFile & "'" & vbCrLf & vbCrLf
                                    msg = msg & "File Not Found." & vbCrLf
                                    msg = msg & "Please varify the correct file name was given."
                                    MsgBox msg, vbExclamation Or vbMsgBoxSetForeground, "File Error!"
                                    GoTo REOPEN
                                End If
                            Wend
                        End If
                        InitFolder = ThisFolder
                        InitFile = buff
                        
                Else
                        
                        ' Return selected file name from the sFileTitle buffer
                        ThisFile = Trim$(CropString(.sFileTitle))

                        ' Return folderpath from buffer
                        ThisFolder = left$(buff, .nFileOffset - 1)
                        
                        If GetSaveFile Then
                            InitFolder = QualifyPath(ThisFolder, True)
                            InitFile = ThisFile
                            Exit Function
                        Else
                            PathSelected = (InitFile = "path" And Len(ThisFile) = 0)
                            ThisFolder = QualifyPath(ThisFolder, (Not PathSelected))
                        End If
                        
                        If PathSelected Then
                            If FileExistPrompt Then
                                If Not ExistFile(ThisFolder) Then
                                    msg = "'" & ThisFolder & "'" & vbCrLf & vbCrLf
                                    msg = msg & "Path Not Found." & vbCrLf
                                    msg = msg & "Please varify the correct path was given."
                                    MsgBox msg, vbExclamation Or vbMsgBoxSetForeground, "Path Error!"
                                    GoTo REOPEN
                                End If
                            End If
                            ThisFile = ""
                        Else
                            If FileExistPrompt Then
                                If Not ExistFile(ThisFolder & ThisFile) Then
                                    msg = "'" & ThisFolder & ThisFile & "'" & vbCrLf & vbCrLf
                                    msg = msg & "File Not Found." & vbCrLf
                                    msg = msg & "Please varify the correct file name was given."
                                    MsgBox msg, vbExclamation Or vbMsgBoxSetForeground, "File Error!"
                                    GoTo REOPEN
                                End If
                            End If
                        End If
                        
                        InitFolder = QualifyPath(ThisFolder, True)
                        InitFile = ThisFile

                End If
                
                Exit Function

            Case Else:
                ' Canceled:

        End Select

        GetOpenSaveFileName = False

   End With

   Exit Function
   End Function

Private Function OFNCenterDialog(hWnd As Long) As Boolean
Dim rc As RECT
Dim newLeft As Long
Dim newTop As Long
Dim dlgWidth As Long
Dim dlgHeight As Long
Dim scrWidth As Long
Dim scrHeight As Long

   If hWnd <> 0 Then
        
        'Position the dialog in the center of
        'the screen. First get the current dialog size.
        Call GetWindowRect(hWnd, rc)
        
        '(To show the calculations involved, I've
        'used variables instead of creating a
        'one-line MoveWindow call.)
        dlgWidth = rc.Right - rc.left
        dlgHeight = rc.Bottom - rc.tOp
        
        scrWidth = Screen.Width \ Screen.TwipsPerPixelX
        scrHeight = Screen.Height \ Screen.TwipsPerPixelY
        
        newLeft = (scrWidth - dlgWidth) \ 2
        newTop = (scrHeight - dlgHeight) \ 2
        
        '..and set the new dialog position.
        Call MoveWindow(hWnd, newLeft, newTop, dlgWidth, dlgHeight, True)
        
        OFNCenterDialog = True
        
End If

End Function

Public Function IsXPRunning() As Boolean

    On Error Resume Next
    
    'Declare structure.
    Dim osVer As OSVERSIONINFO
    
    'Set size of structure.
    osVer.dwOSVersionInfoSize = Len(osVer)
    
    'Fill structure with data.
    GetVersionEx osVer
    
    'Evaluate return. If greater than or equal to 5.1 then running
    'WindowsXP or newer.
    If osVer.dwMajorVersion + osVer.dwMinorVersion / 10 >= 5.1 Then
        IsXPRunning = True
    End If
    
End Function

Public Function IsAppThemedEX() As Boolean

    On Error GoTo NoXPThem
    
    If IsXPRunning Then
        'Declare structure.
        Dim dllVer As DLLVERSIONINFO
        
        'Set size of structure.
        dllVer.cbSize = Len(dllVer)
        
        'Fill structure with data.
        DllGetCtl32Version dllVer
    
        IsAppThemedEX = (IsAppThemed And IsThemeActive)
        
    End If

ThemeChecked:
Err.Clear
Exit Function

NoXPThem:
IsAppThemedEX = False
Resume ThemeChecked

End Function
Public Sub InitializeComctl32(dwICC As INITCommonControlsClasses)

    Dim uccex As tagINITCOMMONCONTROLSEX
    
    With uccex
        .dwICC = dwICC
        .dwSize = Len(uccex)
    End With
    
    On Error Resume Next        '// Avoid "entry point blah not found" error.
    
    Call InitCommonControlsEx(uccex)
    
    '// if we're running anything earlier than ComCtl 4.71, this
    '// will fail, so we retreat to the standard initialization proc.
    If (Err.Number <> 0) Then
        Err.Clear
        On Error Resume Next
        Call InitCommonControls
    End If
    Err.Clear
   
End Sub

Public Function vbGetFileVersion(ByVal sFileName As String, vffi As VS_FIXEDFILEINFO) As Boolean

Dim pData As Long             ' pointer to version info data
Dim nDataLen As Long          ' length of info pointed at by pData
Dim buffer() As Byte          ' buffer for version info resource

    On Error GoTo NoVersion
    
            If sFileName <> "" Then
                ' First, get the size of the version info resource.  If this function fails, then Text1
                ' identifies a file that isn't a 32-bit executable/DLL/etc.
                nDataLen = GetFileVersionInfoSize(sFileName, pData)
                If nDataLen <> 0 Then
                    ' Make the buffer large enough to hold the version info resource.
                    ReDim buffer(0 To nDataLen - 1) As Byte
                    ' Get the version information resource.
                    Call GetFileVersionInfo(sFileName, 0, nDataLen, buffer(0))
    
                    ' Get a pointer to a structure that holds a bunch of data.
                    Call VerQueryValue(buffer(0), "\", pData, nDataLen)
                    ' Copy that structure into the one we can access.
                    CopyMemory vffi, ByVal pData, nDataLen
                    vbGetFileVersion = True
                End If
           
            End If

Exit Function
ExitVer:

NoVersion:
vbGetFileVersion = False
Resume ExitVer

End Function

Public Function vbGetPathName(ByVal szFileName As String, Optional ByVal ReturnLongPath As Boolean) As String

        Dim PathExists As Boolean, lpszPathName As String
        
        On Error Resume Next
        
        PathExists = ExistFile(szFileName)
    
        If Not PathExists Then Open szFileName For Output As #1: Close #1
        
        ' Get short path name
        lpszPathName = String(MAX_PATH, Chr(0))
        If ReturnLongPath Then
            vbGetPathName = left$(lpszPathName, GetLongPathName(szFileName, lpszPathName, MAX_PATH))
        Else
            vbGetPathName = left$(lpszPathName, GetShortPathName(szFileName, lpszPathName, MAX_PATH))
        End If
        
        If Not PathExists Then Kill szFileName
        
End Function

Public Sub KillFile(ByVal sFile As String)
    On Error Resume Next
   Kill sFile
   Exit Sub
End Sub

' PartPos: Returns file extension, including directory, file, and extension positions
Public Function PartPos(sFull As String, iFilePart As Long, _
                      iExtPart As Long) As Boolean


    Dim iDrv As Long, i As Long, cMax As Long
    Dim iDirPart As Long
    
     On Error GoTo ERR_PartPos
    
    cMax = Len(sFull)

    iDrv = Asc(UCase$(left$(sFull, 1)))

    ' If in format d:\path\name.ext, return 3
    If iDrv <= 90 Then                          ' Less than Z
        If iDrv >= 65 Then                      ' Greater than A
            If Mid$(sFull, 2, 1) = ":" Then     ' Second character is :
                If Mid$(sFull, 3, 1) = "\" Then ' Third character is \
                    iDirPart = 3
                End If
            End If
        End If
    Else

        ' If in format \\machine\share\path\name.ext, return position of \path
        ' First and second character must be \
        If Mid$(sFull, 1, 2) <> "\\" Then
            PartPos = False
            GoTo ERR_PartPos
        End If

        Dim fFirst As Boolean
        i = 3
        Do
            If Mid$(sFull, i, 1) = "\" Then
                If fFirst Then
                    iDirPart = i
                    Exit Do
                Else
                    fFirst = True
                End If
            End If
            i = i + 1
        Loop Until i = cMax
    End If

    ' Start from end and find extension
    iExtPart = cMax + 1       ' Assume no extension
    fFirst = False
    Dim sChar As String
    For i = cMax To iDirPart Step -1
        sChar = Mid$(sFull, i, 1)
        If Not fFirst Then
            If sChar = "." Then
                iExtPart = i
                fFirst = True
            End If
        End If
        If sChar = "\" Then
            iFilePart = i + 1
            Exit For
        End If
    Next
    
    PartPos = True
    
PartPos_Continue:
Err.Clear
 Exit Function

ERR_PartPos:
 iFilePart = 0
 iExtPart = 0
 GoTo PartPos_Continue

End Function

Public Function GetFullPath(FullPath As String, _
                     Optional FileDir, _
                     Optional FileBase, _
                     Optional FileExt, _
                     Optional Qualify As Boolean = True) As String


    Dim C As Long, p As Long, sRet As String
    Dim lFileBase As Long, lFileExt As Long
    Dim sFileName As String
    
    If FullPath = vbNullString Then
        GoTo ERR_GetFullPath
    End If
    
     On Error GoTo ERR_GetFullPath
    
    sFileName = Replace(FullPath, "/", "\")
    
    ' Get the path size, then create string of that size
    sRet = String(MAX_PATH, 0)
    If SUCCESSAPI(GetFullPathName(sFileName, MAX_PATH, sRet, p), C) Then 'success
        sRet = left$(sRet, C)

        ' Get the directory, file, and extension parts
        PartPos sRet, lFileBase, lFileExt
        If Not IsMissing(FileDir) Then FileDir = QualifyPath(left$(sRet, lFileBase - 1), Qualify)
        
        If Not IsMissing(FileBase) And lFileBase <= Len(sRet) Then
            FileBase = Mid$(sRet, lFileBase, lFileExt - lFileBase)
        End If
        
        If Not IsMissing(FileExt) And lFileExt <= Len(sRet) Then
            FileExt = Mid$(sRet, lFileExt)
        End If
        
        GetFullPath = sRet
    Else
        GoTo ERR_GetFullPath
    End If
    
GetFullPath_Continue:
Err.Clear
 Exit Function

ERR_GetFullPath:
 GetFullPath = sFileName
 FileDir = vbNullString
 FileBase = vbNullString
 FileExt = vbNullString
GoTo GetFullPath_Continue

End Function

Public Function vbGetWindowClass(ByVal gwcHwnd&) As String
    Dim tempText As String
    tempText = String(260, 0)
    vbGetWindowClass = left(tempText, GetClassName(gwcHwnd, tempText, 260))
End Function

Public Function vbAppPrevInstance(Frm As Form, Optional ByVal BringToFront As Boolean) As Boolean

Dim AppTitle As String, AppHandle As Long

     On Error Resume Next
    
    AppTitle = App.Title
    App.Title = "[#$#%!~*?'2;&]"
    Frm.Caption = "[#$#%!~*?'2;&]"

    AppHandle = FindWindow(vbGetWindowClass(Frm.hWnd), AppTitle)
    If App.PrevInstance Or AppHandle <> 0 Then
        vbAppPrevInstance = True
        
        If BringToFront Then
            AppActivate AppTitle
            If AppHandle <> 0 Then
                OpenIcon AppHandle
                SetForegroundWindow AppHandle
                StayOnTop AppHandle, True
                StayOnTop AppHandle, False
            End If
        End If
        
    Else
        App.Title = AppTitle
        Frm.Caption = AppTitle
        vbAppPrevInstance = False
    End If

End Function

Public Sub StayOnTop(hWnd As Long, ByVal StayOT As Boolean)

Const flags = SWP_NOMOVE Or SWP_NOSIZE

    Select Case StayOT
      Case True
            SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, flags
      Case False
            SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags
      End Select
    
End Sub

Public Function QualifyPath(ByVal sPath As String, Optional ByVal WithBackSlash As Boolean) As String
Dim fp As String

    If Len(sPath) = 0 Then Exit Function
    
    ' QualifyPath the root directory path
    If Right(sPath, 1) = ":" Then
        fp = sPath & "\"
    Else
        fp = sPath
    End If
    
    Select Case WithBackSlash
        Case True   ' Make sure path ends with a backslash
            If Right$(fp, 1) <> "\" Then
                QualifyPath = fp & "\"
            Else
                QualifyPath = fp
            End If
        Case False  ' Make sure folder path doesn't end with a backslash
            If Len(fp) > 3 And Right$(fp, 1) = "\" Then
                QualifyPath = left$(fp, (Len(fp) - 1))
            Else
                QualifyPath = fp
            End If
    End Select
    
End Function

Public Function SUCCESSAPI(ByVal dwIn As Long, Optional apiResult As Long) As Boolean
  If (dwIn <> 0) Then
    SUCCESSAPI = True
  End If
  apiResult = dwIn
End Function

Public Function OS(Optional RequiredOS As OperatingSystems) As OperatingSystems
    ' distinguish operating system
Dim typOSInfo As OSVERSIONINFO
Static CurrOS As OperatingSystems

 On Error GoTo Err_OS

' Only needed once per instance
If CurrOS <> 0 Then GoTo OS_EXIT

' according to the specs of GetVersionExA
'       Before calling the GetVersionEx function,
'       set the dwOSVersionInfoSize member of the OSVERSIONINFO
'       data structure to sizeof(OSVERSIONINFO).
typOSInfo.dwOSVersionInfoSize = Len(typOSInfo)

' check if OS info is valid
If GetVersionEx(typOSInfo) = 0 Then
    CurrOS = 0    ' unexpected error
    GoTo OS_EXIT
End If

' determine OS_name
With typOSInfo
   If .dwPlatformID = 1 And _
                .dwMajorVersion = 4 And _
                .dwMinorVersion < 10 Then ' Windows 95
          CurrOS = Windows95
   ElseIf .dwPlatformID = 1 And _
                .dwMajorVersion = 4 And _
                .dwMinorVersion = 10 Then ' Windows 98
          CurrOS = Windows98
   ElseIf .dwPlatformID = 2 And _
                .dwMajorVersion = 3 And _
                .dwMinorVersion = 51 Then ' Windows NT 3.51
          CurrOS = WindowsNT351
   ElseIf .dwPlatformID = 2 And _
                .dwMajorVersion = 4 And _
                .dwMinorVersion >= 0 Then ' Windows NT 4.0
          CurrOS = WindowsNT4
   ElseIf .dwPlatformID = 2 And _
            .dwMajorVersion = 5 And _
            .dwMinorVersion >= 0 Then
                If ((GetOSVersion() And &H80000000) = 0) Then
                    ' Windows 2000
                    CurrOS = Windows2000
                Else
                    ' Windows NT 5.0
                    CurrOS = WindowsNT5
                End If
   Else
          CurrOS = 0   ' Unknown
   End If
End With

OS_EXIT:
OS = IIf((RequiredOS = 0 Or CurrOS = RequiredOS), CurrOS, 0)
Err.Clear
Exit Function

Err_OS:
CurrOS = 0    ' unexpected error
Resume OS_EXIT
End Function

Public Function CropString(buffer As String, Optional ByVal delimiter As String = vbNullChar, Optional EnumNext As Boolean) As String

Dim intPos As Long, buff As String

    If Len(buffer) = 0 Then Exit Function
    
     On Error GoTo ERR_CropString
    
    buff = buffer & IIf(Right$(buffer, 1) <> delimiter, delimiter, "")
    
    intPos = InStr(1, buff, delimiter)
    CropString = left$(buff, IIf(intPos > 0, intPos - 1, 1))
    
    If EnumNext Then
    
        If Len(Replace(Mid$(buff, intPos), delimiter, "")) <> 0 Then
            ' Return buffer beginning at next character following delimiter
            buffer = Mid$(buff, (intPos + 1))
        Else
            ' We're finished if buffer contains only delimiters
            buffer = vbNullString
        End If
        
    End If
    

CropString_Continue:
Err.Clear
 Exit Function

ERR_CropString:
buffer = vbNullString
Resume CropString_Continue

End Function


' Test file existence with error trapping
Public Function ExistFile(ByVal sSpec As String, Optional RtrnAttr As VbFileAttribute) As Boolean

Dim sFile As String, bExists As Boolean, lAttr As Long

    If Len(sSpec) = 0 Then GoTo ExistFile_Continue
    
    sFile = sSpec
    
     On Error GoTo ERR_ExistFile
    
    If InStr("/\", Right$(sFile, 1)) <> 0 Then sFile = left$(sFile, (Len(sFile) - 1))

    Call FileLen(sFile)
    If (Err.Number = 0) Then bExists = Len(Dir$(sFile, vbDirectory _
                                                            Or vbHidden _
                                                            Or vbNormal _
                                                            Or vbSystem _
                                                            Or vbReadOnly)) > 0
        
    ExistFile = bExists
    
    If bExists Then
        lAttr = RtrnAttr
        
        lAttr = GetAttr(sFile)
        RtrnAttr = lAttr
    End If
    
ExistFile_Continue:
Err.Clear
 Exit Function
 
ERR_ExistFile:

 Resume Next

End Function

Public Function HiWord(lDWord As Long) As Integer
  HiWord = (lDWord And &HFFFF0000) \ &H10000
End Function

Public Function LoWord(lDWord As Long) As Integer
  If lDWord And &H8000& Then
    LoWord = lDWord Or &HFFFF0000
  Else
    LoWord = lDWord And &HFFFF&
  End If
End Function

Public Function vbGetTempPath(Optional ByVal bQualify As Boolean = True, Optional ByVal bUserTemp As Boolean) As String
    Dim sRetPath As String, tmpPath As String
    
    On Error Resume Next
    
    sRetPath = String$(MAX_PATH, 0)
    
    If bUserTemp Then
        tmpPath = left$(sRetPath, GetTempPath(MAX_PATH, sRetPath))
    Else
        tmpPath = QualifyPath(left$(sRetPath, GetWindowsDirectory(sRetPath, MAX_PATH)), True) & "Temp\"
    End If
    
    vbGetTempPath = QualifyPath(tmpPath, bQualify)
    
End Function


