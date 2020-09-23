VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XP Theme-It Manifest Maker"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6195
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picInfoFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   690
      ScaleHeight     =   795
      ScaleWidth      =   5115
      TabIndex        =   20
      Top             =   3540
      Width           =   5115
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1380
         TabIndex        =   23
         ToolTipText     =   "Name in the manifest identifying the executable."
         Top             =   60
         Width           =   3630
      End
      Begin VB.TextBox txtDesc 
         Height          =   285
         Left            =   1380
         TabIndex        =   21
         ToolTipText     =   "Description of the manifest."
         Top             =   410
         Width           =   3630
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   195
         Left            =   855
         TabIndex        =   24
         Top             =   105
         Width           =   465
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         Height          =   195
         Left            =   465
         TabIndex        =   22
         Top             =   450
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H80000005&
      Caption         =   "&Close"
      Height          =   405
      Left            =   2760
      TabIndex        =   26
      ToolTipText     =   "Close program"
      Top             =   4470
      UseMaskColor    =   -1  'True
      Width           =   1425
   End
   Begin VB.CommandButton cmdTheme 
      BackColor       =   &H80000005&
      Caption         =   "&Apply Theme"
      Height          =   405
      Left            =   4290
      TabIndex        =   27
      Top             =   4470
      UseMaskColor    =   -1  'True
      Width           =   1425
   End
   Begin VB.PictureBox picTabFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4605
      Index           =   0
      Left            =   210
      ScaleHeight     =   4605
      ScaleWidth      =   5745
      TabIndex        =   12
      Top             =   630
      Width           =   5745
      Begin VB.CheckBox optCustom 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Manifest Custom Information                                                                  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   200
         TabIndex        =   17
         ToolTipText     =   "Use custom info for new manifest files."
         Top             =   2640
         Width           =   5475
      End
      Begin VB.CheckBox chkCtxMnu 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Enable manifest creation from explorer context menu."
         Height          =   255
         Left            =   200
         TabIndex        =   15
         ToolTipText     =   "Enables you to right click any executable from an explorer window and create a manifest for that file automatically."
         Top             =   2280
         Width           =   4815
      End
      Begin VB.TextBox tbxApp 
         Height          =   315
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1620
         Width           =   3525
      End
      Begin VB.CommandButton cmdBrowse 
         BackColor       =   &H80000005&
         Caption         =   "Load Executable..."
         Height          =   375
         Left            =   3810
         TabIndex        =   13
         Top             =   1590
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"Main.frx":2CFA
         Height          =   1245
         Left            =   210
         TabIndex        =   28
         Top             =   60
         Width           =   5175
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Themed!"
         ForeColor       =   &H00D54600&
         Height          =   195
         Left            =   1080
         TabIndex        =   19
         Top             =   1380
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Name:"
         Height          =   195
         Left            =   210
         TabIndex        =   16
         Top             =   1380
         Width           =   750
      End
   End
   Begin VB.PictureBox picTabFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4605
      Index           =   1
      Left            =   210
      ScaleHeight     =   4605
      ScaleWidth      =   5745
      TabIndex        =   1
      Top             =   630
      Visible         =   0   'False
      Width           =   5745
      Begin ComctlLib.ListView lvManifest 
         Height          =   1695
         Left            =   210
         TabIndex        =   2
         Top             =   1170
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   2990
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Application Manifest"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Location"
            Object.Width           =   8819
         EndProperty
      End
      Begin VB.CommandButton cmdResource 
         BackColor       =   &H80000005&
         Caption         =   "Compile to Resource..."
         Height          =   405
         Left            =   210
         TabIndex        =   18
         ToolTipText     =   "Compile the selected manifest to a new resource file."
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   $"Main.frx":2E80
         Height          =   1065
         Left            =   210
         TabIndex        =   29
         Top             =   60
         Width           =   5385
      End
   End
   Begin VB.PictureBox picTabFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4605
      Index           =   2
      Left            =   210
      Picture         =   "Main.frx":2FCF
      ScaleHeight     =   4605
      ScaleWidth      =   5745
      TabIndex        =   3
      Top             =   630
      Visible         =   0   'False
      Width           =   5745
      Begin VB.CommandButton cmdSysInfo 
         BackColor       =   &H80000005&
         Caption         =   "&System Info..."
         Height          =   345
         Left            =   3960
         TabIndex        =   4
         Top             =   180
         Width           =   1515
      End
      Begin VB.Label lblDisclaimer 
         BackStyle       =   0  'Transparent
         Height          =   1005
         Left            =   360
         TabIndex        =   7
         Top             =   3480
         Width           =   5115
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   390
         Picture         =   "Main.frx":44A7
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblCopywrite 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "written by Greg S. Miller"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3660
         TabIndex        =   25
         Top             =   1170
         Width           =   1725
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email your comments:"
         Height          =   195
         Left            =   1110
         TabIndex        =   11
         Top             =   3000
         Width           =   1590
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "xp.theme-it.com@earthlink.net"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2730
         MouseIcon       =   "Main.frx":6171
         MousePointer    =   99  'Custom
         TabIndex        =   10
         ToolTipText     =   "Contact Greg - (Right Click copies to Clipboard)"
         Top             =   2985
         Width           =   2340
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://home.earthlink.net/~xp.theme-it.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1770
         MouseIcon       =   "Main.frx":647B
         MousePointer    =   99  'Custom
         TabIndex        =   9
         ToolTipText     =   "Visit Homepage - (Right Click copies to Clipboard)"
         Top             =   2655
         Width           =   3210
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Home Page:"
         Height          =   195
         Left            =   840
         TabIndex        =   8
         Top             =   2670
         Width           =   990
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAuthor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "XP Theme-It Manifest Maker"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   315
         Left            =   1290
         TabIndex        =   6
         Top             =   780
         Width           =   4080
      End
      Begin VB.Label lblLicense 
         BackStyle       =   0  'Transparent
         Height          =   825
         Left            =   360
         TabIndex        =   5
         Top             =   1560
         Width           =   5115
         WordWrap        =   -1  'True
      End
   End
   Begin ComctlLib.TabStrip tabManifest 
      Height          =   5235
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   9234
      TabWidthStyle   =   2
      TabFixedWidth   =   2999
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Application"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Manifest/Resource"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "About"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "XP Theme-It Manifest Maker"

Option Explicit

Private WithEvents cReg As CRegistry
Attribute cReg.VB_VarHelpID = -1
Private AppExePath As String
Private SysTempPath As String
Private m_DefName As String
Private m_DefDesc As String

Private tbxPath As String
Private tbxFile As String
Private LClk As Boolean

' Application keys in registry
Const RegOnCreate As String = "Software\XPTheme-It\Settings\OnCreate"
Const RegManifests As String = "Software\XPTheme-It\Manifests\"
Const RegLayers As String = "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers"
Const sHome As String = "home.earthlink.net/~xp.theme-it.com/"
Const sEmail As String = "xp.theme-it.com@earthlink.net"

Private Function CreateManifest(sExecutable As String, _
            Optional ByVal sName As String, _
            Optional ByVal sDesc As String)

    Dim sVer As String, sManifest As String
    Dim iRet As Long
    Dim vffi As VS_FIXEDFILEINFO

    If ExistFile(sExecutable) Then
    
        sManifest = sExecutable & ".manifest"
        
        'prompt to overwrite if manifest already exists.
        If ExistFile(sManifest) Then
            Select Case MsgBox("A manifest already exists for the selected executable." & vbCrLf & vbCrLf & "Do you want to backup the existing manifest before saving?", vbQuestion Or vbYesNoCancel Or vbMsgBoxSetForeground)
                Case vbYes
                    Name sManifest As FileBackup(sManifest)
                    iRet = adSaveCreateNotExist 'backup manifest
                Case vbNo:
                    iRet = adSaveCreateOverWrite 'overwrite manifest
                Case Else:
                    CreateManifest = False
                    Exit Function
            End Select
        Else
            iRet = True
        End If

        On Error Resume Next
        
        ' Determine the exe version number to list in the manifest
        If vbGetFileVersion(sExecutable, vffi) Then
            sVer = Trim$(Str$(HiWord(vffi.dwFileVersionMS))) & "." & _
            Trim$(Str$(LoWord(vffi.dwFileVersionMS))) & "." & _
            Trim$(Str$(HiWord(vffi.dwFileVersionLS))) & "." & _
            Trim$(Str$(LoWord(vffi.dwFileVersionLS)))
        Else
            sVer = "1.0.0.0"
        End If
        
        ' Extract the name of the exe file
        If Len(sName) = 0 Or Not CBool(optCustom.Value) Then sName = Dir$(sExecutable)
        If Len(sDesc) = 0 Or Not CBool(optCustom.Value) Then
            sDesc = FileDescription(sExecutable)
            If Len(sDesc) = 0 Then sDesc = "XP.Visual.Style.Manifest"
        End If

        ' To properly link to version 6 of the common controls you need to
        ' be able to provide an application Manifest to the executable. I found
        ' that using a Stream is a safe way here to save the xml contents to
        ' file, while maintaining the UTF-8 CharSet encoding.
        'NOTE:  Microsoft ActiveX Data Objects library reference is required.
        Dim stm As New StreamData.Stream
        Dim iPad As Integer
        stm.Open
        stm.Type = adTypeText
        stm.Position = 0
        stm.Charset = "UTF-8"
        
        stm.WriteText "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf
        stm.WriteText "<assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" manifestVersion=""1.0"">" & vbCrLf
        stm.WriteText "<!--Manifest Created With " & Trim$(App.Title) & " " & CStr(App.Major) & "." & CStr(App.Minor) & "!-->" & vbCrLf
        stm.WriteText "<assemblyIdentity" & vbCrLf
        stm.WriteText "    version=""" & Trim$(sVer) & """" & vbCrLf
        stm.WriteText "    processorArchitecture=""X86""" & vbCrLf
        stm.WriteText "    name=""" & sName & """" & vbCrLf
        stm.WriteText "    type=""win32""" & vbCrLf
        stm.WriteText "/>" & vbCrLf
        stm.WriteText "<description>" & sDesc & "</description>" & vbCrLf
        stm.WriteText "<dependency>" & vbCrLf
        stm.WriteText "    <dependentAssembly>" & vbCrLf
        stm.WriteText "        <assemblyIdentity" & vbCrLf
        stm.WriteText "            type=""win32""" & vbCrLf
        stm.WriteText "            name=""Microsoft.Windows.Common-Controls""" & vbCrLf
        stm.WriteText "            version=""6.0.0.0""" & vbCrLf
        stm.WriteText "            processorArchitecture=""X86""" & vbCrLf
        stm.WriteText "            publicKeyToken=""6595b64144ccf1df""" & vbCrLf
        stm.WriteText "            language=""*""" & vbCrLf
        stm.WriteText "        />" & vbCrLf
        stm.WriteText "    </dependentAssembly>" & vbCrLf
        stm.WriteText "</dependency>" & vbCrLf
        stm.WriteText "</assembly>"
        
        ' just to be sure we got it all
        stm.Flush
        
        ' The manifest must be saved with an even number of bytes, multiple of 4.
        ' To assure we have the correct bytes, we may need to pad the end of
        ' the file with spaces.
        iPad = (4 - (stm.Size Mod 4))
        If iPad < 4 Then stm.WriteText Space$(iPad)
        
        ' Save the new manifest file
        If (Err = 0) Then
            ' looks good!
            stm.SaveToFile sManifest, adSaveCreateOverWrite
            CreateManifest = iRet
        End If
        
        ' clean up
        stm.Close
        Set stm = Nothing
    
    End If

End Function

Private Function FileBackup(ByVal sFile As String) As String

Dim BckIncr As Integer, sBckup As String

    On Error Resume Next
    BckIncr = 1
    
    sBckup = (sFile & ".BAK")
    Do While ExistFile(sBckup)
        BckIncr = BckIncr + 1
        sBckup = (sFile & Format$(CStr(BckIncr), "00") & ".BAK")
    Loop
    FileBackup = sBckup
    
End Function
Private Function FileDescription(Optional ByVal sFile As String) As String

Dim tempFile As String, i As Integer, FileLength As Long
Dim pos As Long, pnStart As Long, NextText As String
Dim StartPos As Long, EndPos As Long
Dim nStage As String, lSide As String, rSide As String
Dim StringFileInfo As String

On Error Resume Next

If Not ExistFile(sFile) Then Exit Function

Open sFile For Binary As #1
    tempFile = Space(LOF(1))
    Get #1, , tempFile
Close #1

pos = InStr(tempFile, (Chr(1) & Replace(" S t r i n g F i l e I n f o", " ", Chr(0))))
If pos <> 0 Then
    pnStart = InStr(pos, tempFile, (Chr(1) & Replace(" F i l e D e s c r i p t i o n", " ", Chr(0))))
    NextText = (Chr(1) & Replace(" F i l e V e r s i o n", " ", Chr(0)))
    FileLength = 34
Else
    pos = InStr(tempFile, "StringFileInfo")
    If pos = 0 Then pos = 1
    pnStart = InStr(pos, tempFile, "FileDescription")
    NextText = "FileVersion"
    FileLength = 16
End If

If pnStart > 0 Then
    StartPos = pnStart + FileLength
    EndPos = InStr(StartPos, tempFile, String(3, Chr(0)))
    
    If InStr(Mid(tempFile, StartPos, EndPos - StartPos), NextText) <> 0 Then
        For i = 1 To 255
            If CInt(Asc(Mid(tempFile, StartPos + i, 1))) <= 31 Then
                EndPos = StartPos + i
                Exit For
            End If
        Next i
    End If
    
    nStage = Mid(tempFile, StartPos, EndPos - StartPos)
    If InStr(nStage, Chr(0)) <> 0 Then
        Do Until InStr(nStage, Chr(0)) = 0
            lSide = left(nStage, InStr(nStage, Chr(0)) - 1)
            rSide = Right(nStage, (Len(nStage) - Len(lSide) - Len(Chr(0))))
            nStage = lSide & rSide
        Loop
    End If
    
    FileDescription = Trim$(nStage)

End If

End Function
Private Sub chkCtxMnu_Click()

' Set the context menu commands to the registry.
If chkCtxMnu.Value = vbChecked Then
    cReg.SetValueData HKEY_CLASSES_ROOT, "exefile\shell\XPManifest", "", ValString, "Create Manifest for this executable"
    cReg.SetValueData HKEY_CLASSES_ROOT, "exefile\shell\XPManifest\Command", "", ValString, """" & AppExePath & """ /Create ""%1"""
Else
    ' remove
    cReg.DeleteKey HKEY_CLASSES_ROOT, "exefile\shell\XPManifest"
End If
End Sub

Private Sub cmdBrowse_Click()

Dim InitDir As String, FileName As String

    ' Select an executable for the new manifest.
    If GetOpenSaveFileName(FileName, InitDir, "Choose Any executable", , _
        "Program Executables |*.exe;*.dll;*.ocx;*.msc;*.cpl;*.msi|All Files (*.*)|*.*", "exe", _
        , True, , , , , Me.hWnd, , True) Then
            
            On Error GoTo CancelErr
            
            tbxPath = InitDir
            tbxFile = FileName
            txtName.Text = tbxFile
            
            ' save the path
            tbxApp.Text = InitDir & FileName
            If ExistFile(tbxApp.Text & ".manifest") Then
                tbxApp.ToolTipText = "Manifest created: " & FileDateTime(tbxApp.Text & ".manifest")
                lblStatus.Caption = "Themed!"
            Else
                tbxApp.ToolTipText = vbNullString
                lblStatus.Caption = vbNullString
            End If
            
    End If

ErrContinue:
Err.Clear
Exit Sub

CancelErr: Resume ErrContinue

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdResource_Click()

Dim Manifest As String, ManifestName As String
Dim DirRes As String, ResName As String
Dim ManifestPath As String, bRet As Boolean

' Get the current path to the selected manifest file.
ManifestPath = QualifyPath(lvManifest.SelectedItem.SubItems(1), True)
ManifestName = lvManifest.SelectedItem.Text
ManifestName = left$(ManifestName, Len(ManifestName) - 9)

Manifest = Replace((vbGetPathName(ManifestPath & lvManifest.SelectedItem.Text)), "\", "\\")

' Use the manifest name as default for the new resource file.
ResName = (ManifestName & ".res")

' Select a file name and location for the new resource file.
If GetOpenSaveFileName(ResName, DirRes, "Save Resource As", True, _
        "Resource Files |*.res|All Files (*.*)|*.*", "res", , False, , , , True, Me.hWnd, _
        OFN_NODEREFERENCELINKS, True, , , , "Resource name:") Then
        
        On Error Resume Next
        
        Screen.MousePointer = 11
        picTabFrame(0).Enabled = False
        
        ResName = CropString(ResName)
        ResName = left$(ResName, Len(ResName) - 4)

        ' Prompt to overwrite if resource file already exists.
        If ExistFile(DirRes & ResName & ".res") Then
            If vbNo = MsgBox((DirRes & ResName & ".res") & " already exists." & vbCrLf & vbCrLf & "Do you want to replace it?", _
            vbQuestion Or vbMsgBoxSetForeground Or vbYesNo, "Save Resource As") Then
                GoSub Continue
            End If
        End If

        ' Delete old
        KillFile (DirRes & ResName & ".res")
        KillFile (DirRes & ResName & ".rc")
        KillFile SysTempPath & "~re.res"
        KillFile SysTempPath & "~rc.rc"

        ' Store the path of the manifest file into a temporary resource script.
        Open SysTempPath & "~rc.rc" For Output Access Write As #1
        
        ' CREATEPROCESS_MANIFEST_RESOURCE_ID (1),  RT_MANIFEST (24), PATH
        Print #1, ("1 24 """ & Manifest & """")
        Close #1

        Me.Refresh
        
        ' Compile the resource file containing the selected manifest.
        bRet = ShellExec("rc.exe /l409 /c1252 /r /fo """ _
        & SysTempPath & "~re.res" & """ """ & SysTempPath & "~rc.rc" & """", vbHide, 2000)

        If bRet Then
            ' Copy the temp resource file to the selected path.
            Name SysTempPath & "~re.res" As (DirRes & ResName & ".res")
            MsgBox "Resource file containing manifest was successfully compiled to " _
                & vbCrLf & vbCrLf & DirRes & ResName & ".res", vbInformation Or vbMsgBoxSetForeground
        Else
            MsgBox "There was a problem compiling the resource manifest." _
                & vbCrLf & vbCrLf & "Make sure you have the Microsoft Resource Compiler installed on your system, and try again.", vbCritical Or vbMsgBoxSetForeground
        End If

        ' Clean up
        KillFile SysTempPath & "~re.res"
        KillFile SysTempPath & "~rc.rc"

Continue:
Err.Clear

        Screen.MousePointer = 0
        picTabFrame(0).Enabled = True
        Exit Sub
                
End If


End Sub

Private Sub cmdSysInfo_Click()

    Dim SysInfoPath As String
    Const RegKeySharedTools = "SOFTWARE\Microsoft\Shared Tools\MSInfo"

    On Error GoTo ERR_StartSysInfo

    ' Try To Get System Info Program Path From Registry...
    SysInfoPath = cReg.GetValueData(HKEY_LOCAL_MACHINE, RegKeySharedTools, "Path", ValString, , vbNullString)

    If SysInfoPath = vbNullString Then GoTo ERR_StartSysInfo

    'If successfull then open system info
    Call Shell(SysInfoPath, vbNormalFocus)


Exit Sub
ERR_StartSysInfo:
Resume Next

    
End Sub

Private Sub cmdTheme_Click()

    Dim iRet As Long, ManifestName As String
    
    ' Action to perform
    On Error Resume Next
    
    Select Case cmdTheme.Caption
        Case "&Apply Theme"
            If tbxApp.Text = "" Or tbxPath = "" Then
                MsgBox "Executable not loaded", vbExclamation
                
            ElseIf CBool(CreateManifest(tbxPath & tbxFile, Trim$(txtName.Text), Trim$(txtDesc.Text))) Then
                    'save to manifest settings
                    cReg.SetValueData HKEY_LOCAL_MACHINE, RegManifests & tbxFile, "Location", ValString, tbxPath
                    cReg.SetValueData HKEY_CURRENT_USER, RegLayers, (tbxFile & ".manifest"), ValString, "MANIFEST.xpsp1.(6595b64144ccf1df)"
                
                    'Enumerate and reload saved manifests to listview
                    lvManifest.ListItems.Clear
                    cReg.EnumSubKeys HKEY_LOCAL_MACHINE, RegManifests, 1

                    'update manifest status
                    tbxApp.ToolTipText = "Manifest created: " & FileDateTime(tbxPath & tbxFile & ".manifest")
                    
                    ' ok, not so fast!
                    Wait 300
                    
                    lblStatus.Caption = "Themed!"
                    
                    MsgBox "Application manifest for '" & tbxPath & tbxFile & " was created successfully.", vbInformation Or vbMsgBoxSetForeground
            End If
            
        Case "&Restore"
            If lvManifest.SelectedItem.Selected = False Then
                MsgBox "Please Select Application To Restore"
            ElseIf CBool(lvManifest.ListItems.Count) Then
                
                'Delete application manifest
                KillFile (QualifyPath(lvManifest.SelectedItem.SubItems(1), True) & lvManifest.SelectedItem.Text)

                ' Manifest name
                ManifestName = left(lvManifest.SelectedItem.Text, Len(lvManifest.SelectedItem.Text) - 9)
                
                'Remove from listview
                lvManifest.ListItems.Remove (lvManifest.SelectedItem.Index)
                                
                'Delete from registry
                cReg.DeleteKey HKEY_LOCAL_MACHINE, RegManifests & ManifestName
                cReg.DeleteValueName HKEY_CURRENT_USER, RegLayers, lvManifest.SelectedItem.Text
                
                'Update status
                lblStatus.Caption = vbNullString
                tbxApp.ToolTipText = vbNullString

            End If
    End Select

End Sub

Private Sub cReg_Subkeys(ByVal SubKeyRoot As String, ByVal SubKeyName As String)

Dim fPath As String
Dim tmpReg As New CRegistry

' Any manifest previously created by Theme-It should be saved to the
' registry.
fPath = tmpReg.GetValueData(HKEY_LOCAL_MACHINE, RegManifests & SubKeyName, "Location", ValString, , vbNullString)

If Len(fPath) <> 0 Then
    fPath = QualifyPath(fPath, True)
    ' Varify each manifest file in registry exists on the system before adding them to the listview
    If ExistFile(fPath & SubKeyName) And ExistFile(fPath & SubKeyName & ".manifest") Then
        With lvManifest.ListItems.Add(, , SubKeyName & ".manifest")
            .SubItems(1) = fPath
        End With
    Else
        'Delete  from registry
        cReg.DeleteKey HKEY_LOCAL_MACHINE, RegManifests & SubKeyName
        cReg.DeleteValueName HKEY_CURRENT_USER, RegLayers, SubKeyName & ".manifest"
    End If
End If
Set tmpReg = Nothing

End Sub

Private Sub Form_Initialize()

    ' You most likely will need to call InitCommonControls before any visual
    ' elements are displayed. In some cases, this may not be neccesary,
    ' because your application may already link to ComCtl32.dll. For example,
    ' your project may have one or more of the Windows Common Controls
    ' on a form module. However, to be certain, it's best to call
    ' InitCommonControls prior to displaying any forms.
    InitializeComctl32 ICC_USEREX_CLASSES
    
End Sub


Private Sub Form_Load()

Dim cmd As String, cmName As String, cmExt As String, bNoPrevInst As Boolean
Dim cmPath As String, cmCommand As String, cmFile As String

On Error Resume Next

' Only allow one instance of the program to run.
bNoPrevInst = (Not vbAppPrevInstance(Me, True))

If bNoPrevInst Then
    
    Set cReg = New CRegistry

    cmd = Trim$(Replace(Command$, """", ""))
    If Len(cmd) <> 0 Then
        '''this code is here for the context menus
        
        'extract the flags from the command-line arguments, that
        If left$(LCase$(cmd), 7) = "/create" Then
            
            'file executable from command-line arguments
            cmFile = Trim$(Mid$(cmd, 8))
            cmName = Dir$(cmFile)
            
            ' Create the new manifest for the selected executable.
            If CreateManifest(cmFile) Then
                'save the new manifest to the registry.
                cReg.SetValueData HKEY_LOCAL_MACHINE, RegManifests & cmName, "Location", ValString, left$(cmFile, InStr(cmFile, cmName) - 1)
                cReg.SetValueData HKEY_CURRENT_USER, RegLayers, cmName & ".manifest", ValString, "MANIFEST.xpsp1.(6595b64144ccf1df)"
            
                MsgBox "Application manifest for '" & cmName & " was created successfully.", vbExclamation Or vbMsgBoxSetForeground
            
            End If
            
        ElseIf left$(LCase$(cmd), 8) = "/restore" Then

            'file executable from command-line arguments
            cmFile = Trim$(Mid$(cmd, 9))
            cmName = Dir$(cmFile) ' name

            'remove manifest settings in the registry that belong to the selected executable.
            cReg.DeleteKey HKEY_LOCAL_MACHINE, RegManifests & cmName
            cReg.DeleteValueName HKEY_CURRENT_USER, RegLayers, (cmName & ".manifest")
            
            'delete the manifest
            KillFile (cmFile & ".manifest")
        
        End If
    
        End
        
    End If


    ' No command-line arguments were passed from the shell.
    
    AppExePath = QualifyPath(App.Path, True) & App.EXEName & ".exe"
    SysTempPath = vbGetTempPath() 'windows temp dir
    
    ' window display settings
    lblStatus.Caption = vbNullString
    lblLicense.Caption = App.FileDescription
    lblDisclaimer.Caption = App.Comments
    picTabFrame(2).ZOrder 0
    picTabFrame(1).ZOrder 0
    picTabFrame(0).ZOrder 0
    picInfoFrame.ZOrder 0
    cmdExit.ZOrder 0
    cmdTheme.ZOrder 0
    
    '
    If Not IsAppThemedEX Then
        lvManifest.Appearance = 1
        tbxApp.Appearance = 1
        picTabFrame(0).Appearance = 1
        picTabFrame(1).Appearance = 1
        picTabFrame(2).Appearance = 1
        chkCtxMnu.Appearance = 1
        optCustom.Appearance = 1
        txtDesc.Appearance = 1
        txtName.Appearance = 1
        picInfoFrame.Appearance = 1
        optCustom.BackColor = vbButtonFace
        txtDesc.BackColor = vbButtonFace
        txtName.BackColor = vbButtonFace
        cmdBrowse.BackColor = vbButtonFace
        cmdTheme.BackColor = vbButtonFace
        cmdExit.BackColor = vbButtonFace
        picInfoFrame.BackColor = vbButtonFace
    End If

    'Enumerate and list saved manifests to listview
    cReg.EnumSubKeys HKEY_LOCAL_MACHINE, RegManifests, 1

    ' Begin with the first tab
    tabManifest_Click
    
    ' get app settings
    chkCtxMnu.Value = Abs(cReg.KeyExist(HKEY_CLASSES_ROOT, "exefile\shell\XPManifest\Command"))
    optCustom.Value = cReg.GetValueData(HKEY_LOCAL_MACHINE, RegOnCreate, "Customize", ValLong, , vbChecked)
        
Else
    End
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set cReg = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Label2_Click()

On Error Resume Next

If LClk Then
    Call ShellExecute(0&, 0&, "http://" & sHome, 0&, 0&, vbNormalFocus)
Else
      Clipboard.Clear
      Clipboard.SetText "http://" & sHome
End If

Exit Sub

End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
LClk = Button <> 2
If LClk Then Label2.ForeColor = &HFF&
End Sub


Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 2 Then Label2.ForeColor = vbBlue
End Sub

Private Sub Label3_Click()

 On Error Resume Next

If LClk Then
      Call ShellExecute(0&, 0&, "mailto:" & sEmail, 0&, 0&, vbNormalFocus)
Else
      Clipboard.Clear
      Clipboard.SetText sEmail
End If

Exit Sub

End Sub


Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
LClk = Button <> 2
If LClk Then Label3.ForeColor = &HFF&
End Sub


Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 2 Then Label3.ForeColor = vbBlue
End Sub



Private Sub lvManifest_AfterLabelEdit(Cancel As Integer, NewString As String)
    Cancel = True
End Sub

Private Sub lvManifest_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub lvManifest_Click()

Dim ManifestPath As String, NextTxtLine As String, pos As Integer
Dim bManifest As Boolean, bAppName As Boolean, bAppDesc As Boolean

On Error Resume Next

    ' Open selected manifest in listview, and retrieve the info
    If tabManifest.SelectedItem.Index = 2 Then
        ' Restore
        ManifestPath = QualifyPath(lvManifest.SelectedItem.SubItems(1), True) & lvManifest.SelectedItem.Text
        
        If ExistFile(ManifestPath) Then
        
            Open ManifestPath For Input Access Read As #1 ' Open text file.
            Do While Not EOF(1)
                Line Input #1, NextTxtLine
                
                If (InStr(LCase$(NextTxtLine), "manifest created with xp theme-it") <> 0) _
                And (Not bManifest) Then
                    bManifest = True
                ElseIf bManifest Then
                    pos = InStr(LCase$(NextTxtLine), "name=")
                    If pos <> 0 And (Not bAppName) Then
                        bAppName = True
                        txtName.Text = Replace(Trim(Mid$(NextTxtLine, pos + 5)), """", "")
                    ElseIf bAppName And (Not bAppDesc) Then
                        If InStr(LCase$(NextTxtLine), "<description>") <> 0 Then
                            bAppDesc = True
                            txtDesc.Text = Trim$(Replace(Replace(NextTxtLine, "<description>", ""), "</description>", ""))
                            Exit Do
                        End If
                    End If
                End If
NextItem:
            
            Loop
            Close #1
        
        End If

End If

End Sub

Private Sub optCustom_Click()

    ' Custom info will be used for new manifest
    txtDesc.Enabled = CBool(optCustom.Value)
    txtName.Enabled = CBool(optCustom.Value)
    
    ' save to registry
    cReg.SetValueData HKEY_LOCAL_MACHINE, RegOnCreate, "Customize", ValLong, Abs(CBool(optCustom.Value))
    
End Sub

Private Sub tabManifest_Click()

picTabFrame(0).Visible = False
picTabFrame(1).Visible = False
picTabFrame(2).Visible = False

txtDesc.Locked = True
txtName.Locked = True

On Error Resume Next

Select Case tabManifest.SelectedItem.Index
    Case 1 ' Application
    
        txtDesc.Text = m_DefDesc
        txtName.Text = m_DefName
        txtDesc.Locked = False
        txtName.Locked = False

        txtDesc.Enabled = CBool(optCustom.Value)
        txtName.Enabled = CBool(optCustom.Value)

        cmdTheme.ToolTipText = "Create a manifest for the selected executable."
        cmdTheme.Caption = "&Apply Theme"
        cmdTheme.Visible = True
        cmdExit.Visible = True
        picTabFrame(0).Visible = True
        picInfoFrame.Visible = True
        
    Case 2 ' Restore
    
        cmdTheme.ToolTipText = "Delete the selected manifest file."
        cmdTheme.Caption = "&Restore"
        cmdTheme.Visible = True
        cmdExit.Visible = True
        
        picTabFrame(1).Visible = True
        picInfoFrame.Visible = True
        txtDesc.Enabled = True
        txtName.Enabled = True
        
        If lvManifest.ListItems.Count <> 0 Then
            lvManifest.SelectedItem.Selected = True
            lvManifest.SetFocus
        End If
        
        ' Update seleted item
        lvManifest_Click
        
    Case 3 ' About
    
        cmdExit.Visible = False
        cmdTheme.Visible = False
        picInfoFrame.Visible = False
        picTabFrame(2).Visible = True
End Select

End Sub








Private Sub txtDesc_Change()
    ' store the manifest description
    If Not txtDesc.Locked Then m_DefDesc = txtDesc.Text
End Sub

Private Sub txtName_Change()
    ' store the manifest name
    If Not txtName.Locked Then m_DefName = txtName.Text
End Sub


