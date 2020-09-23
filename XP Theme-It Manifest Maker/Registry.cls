VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-
'   Windows 95/98/2000/NT Workstation/NT Server.
'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-
'   CREGISTRY.CLS - Fully Functional, Updated by Greg Miller 09/02/01
'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-'-
Option Explicit

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
    szCSDVersion As String * 128
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (OSInfo As OSVERSIONINFO) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegConnectRegistry Lib "advapi32" Alias "RegConnectRegistryA" (ByVal szMachineName As String, ByVal hKey As Long, lphKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal szSubkey As String, ByVal lpReserved As Long, ByVal szClass As String, ByVal dwOptions As Long, ByVal dwDesiredAccess As Long, lpSecurityAttributes As SecurityAttributes, lphResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal szSubkey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal szValueName As String) As Long
Private Declare Function RegEnumKey Lib "advapi32" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal szSubkey As String, ByVal lpcbSubkey As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal szValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpdwType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal szSubkey As String, ByVal dwOptions As Long, ByVal dwDesiredAccess As Long, lphKey As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpszClass As String, lpcbClass As Long, ByVal lReserved As Long, lpSubkeys As Long, lpcbSubkey As Long, lpcbClass As Long, lpValues As Long, lpcbValueName As Long, lpcbValue As Long, lpSecurityDescriptor As Long, lpftLastWritetime As Currency) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal szValueName As String, ByVal lpReserved As Long, lpValuetype As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExByte Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal szValueName As String, ByVal lpReserved As Long, lpValuetype As Long, lpbData As Byte, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal szValueName As String, ByVal lpReserved As Long, lpValuetype As Long, lpdwData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal szValueName As String, ByVal lpReserved As Long, lpValuetype As Long, ByVal lpszData As String, lpcbData As Long) As Long
Private Declare Function RegSetValueExByte Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal szValueName As String, ByVal lpReserved As Long, ByVal dwValuetype As Long, bData As Byte, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal szValueName As String, ByVal lpReserved As Long, ByVal dwValuetype As Long, dwData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal szValueName As String, ByVal lpReserved As Long, ByVal dwValuetype As Long, ByVal szData As String, ByVal cbData As Long) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal szSource As String, ByVal szDestination As String, ByVal cbSize As Long) As Long

' Security Mask constants
Private Const Reg_KeyQueryValue = &H1 ' Permission to query subkey data.
Private Const Reg_KeySetValue = &H2 ' Permission to set subkey data.
Private Const Reg_KeyCreateSubkey = &H4 ' Permission to create subkeys.
Private Const Reg_KeyEnumSubkey = &H8 ' Permission to enumerate subkeys.
Private Const Reg_KeyNotify = &H10 ' Permission for change notification.
Private Const Reg_KeyCreateLink = &H20 ' Permission to create a symbolic link.
Private Const Reg_Synchronize = &H100000
Private Const Reg_Standard_Rights_All = &H1F0000

Private Const Reg_KeyAllAccess = ((Reg_Standard_Rights_All Or Reg_KeyQueryValue Or _
   Reg_KeySetValue Or Reg_KeyCreateSubkey Or Reg_KeyEnumSubkey Or Reg_KeyNotify Or _
   Reg_KeyCreateLink) And (Not Reg_Synchronize))

Private Const Reg_OptionNonVolatile = &H0 ' The information is stored in a file and is preserved when the system is restarted.
Private Const Reg_OptionVolatile = &H1 ' The information is stored in memory and is not preserved when the system is restarted. Ignored in Windows 95/98.
Private Const Reg_OptionCreateLink = &H2 ' The created key is a symbolic link.
Private Const Reg_OptionBackupRestore = &H4 ' Ignore the SamDesired parameter and attempts to open the key with the access required to backup or restore the key. Ignored in Windows 95/98.
Private Const Reg_OptionOpenLink = &H8 ' Open symbolic link.
 
Private Const Reg_KeyCreated = &H1 ' The key did not exist and was created.
Private Const Reg_KeyOpened = &H2 ' The key existed and was simply opened without being changed.

Private Const Reg_RestoreWholeHiveVolatile = &H1 ' Instead of restoring the given key, a new volatile set of registry information is created. Only valid for HKEY_USERS or HKEY_LOCAL_MACHINE.
Private Const Reg_RestoreRefreshHive = &H2 ' Unwind changes to last flush.
Private Const Reg_RestoreNoLazyFlush = &H4 ' Never lazy flush this hive.

Private Const Reg_NotifyName = &H1 ' A subkey is added or deleted.
Private Const Reg_NotifyAttributes = &H2 ' Changes to the attributes of a keys, such as security descriptor information.
Private Const Reg_NotifyLastSet = &H4  ' Changes to a value of the key. Include adding or deleting values, or changing an existing value.
Private Const Reg_NotifySecurity = &H8 ' Changes to the security descriptor of the key.

Public Enum RegEnumAbortFlags
    Abort_SubKeys
    Abort_ValueNames
    Abort_KeyDelete
End Enum

Private m_AbortEnum As Boolean
Private m_Abort_ValueNames As Boolean
Private m_Abort_KeyDelete As Boolean

Private Type arrSubKeys
    sKeyName As String
    sSubKey As String
End Type

Private m_EnumKeys() As arrSubKeys
Private m_DeleteKeys() As arrSubKeys

Private Type SecurityAttributes
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type

Enum RegKey
  HKEY_CLASSES_ROOT = &H80000000
  HKEY_CURRENT_USER = &H80000001
  HKEY_LOCAL_MACHINE = &H80000002
  HKEY_USERS = &H80000003
  HKEY_PERFORMANCE_DATA = &H80000004
  HKEY_CURENT_CONFIG = &H80000005
  HKEY_DYN_DATA = &H80000006
End Enum

Enum RegValue
  ValString = &H1 ' Null-terminated string.
  ValStringEx = &H2 ' Null-terminated string with unexpanded references to environment variables.
  ValBinary = &H3 ' Binary data.
  ValLong = &H4 ' 32-bit number.
  ValBoolean = &H5 ' Set as Long, Return as Boolean
  ValMultiString = &H7 ' Array of null-terminated strings, terminated by two null characters.
End Enum

' Event raised with each key found
Event Subkeys(ByVal SubKeyRoot As String, ByVal SubKeyName As String)
Event ValueNames(ByVal SubKeyRoot As String, ByVal ValueName As String)
Event KeyDelete(ByVal KeyName As String, Cancel As Boolean)




Function CreateKey(ByVal PredefinedKey As RegKey, ByVal KeyName As String, Optional RegResult As Long) As Boolean
  Dim hKey As Long, dwDisposition As Long, SA As SecurityAttributes
    
    On Error GoTo ERR_CreateKey
    
    If RegOk(RegConnectRegistry(vbNullString, PredefinedKey, hKey), RegResult) Then
        If RegOk(RegCreateKeyEx(hKey, KeyName, 0, "", Reg_OptionNonVolatile, Reg_KeyCreateSubkey, SA, hKey, dwDisposition), RegResult) Then
            CreateKey = True
        End If
    End If

CreateKey_Continue:
If hKey <> 0 Then RegCloseKey hKey
Exit Function
    
ERR_CreateKey:
Resume CreateKey_Continue
End Function

Private Function Is9XRunning() As Boolean

Dim typOSInfo As OSVERSIONINFO
Static CurrOS As Integer

    On Error GoTo Err_OS
    
    Select Case CurrOS
        Case 1 ' Windows 9X is NOT running
            Is9XRunning = False
            GoTo OS_EXIT
        Case 2 ' Windows 9X is running
            Is9XRunning = True
            GoTo OS_EXIT
        Case Else: ' Determine which version of windows is running

            ' according to the specs of GetVersionExA
            '       Before calling the GetVersionEx function,
            '       set the dwOSVersionInfoSize member of the OSVERSIONINFO
            '       data structure to sizeof(OSVERSIONINFO).
            typOSInfo.dwOSVersionInfoSize = Len(typOSInfo)

            ' check if OS info is valid
            If GetVersionEx(typOSInfo) <> 0 Then
                With typOSInfo
                    If .dwPlatformID = 1 And .dwMajorVersion = 4 And .dwMinorVersion <= 10 Then
                         ' Windows 95/98 is running
                           CurrOS = 2
                           GoTo OS_EXIT
                    End If
                End With
            End If
            CurrOS = 1
    
    End Select


OS_EXIT:
Err.Clear
Exit Function


Err_OS:
CurrOS = 0
Resume OS_EXIT

End Function

Public Function KeyExist(ByVal PredefinedKey As RegKey, ByVal KeyName As String) As Boolean
  Dim hKey As Long
  
    If RegOk(RegConnectRegistry(vbNullString, PredefinedKey, hKey)) Then
        If RegOk(RegOpenKeyEx(hKey, KeyName, 0, Reg_KeyQueryValue, hKey)) Then
            KeyExist = True
        End If
    End If
    
KeyExist_Continue:
If hKey <> 0 Then RegCloseKey hKey
Exit Function

ERR_KeyExist:
KeyExist = False
Resume KeyExist_Continue

End Function

Public Function GetKeyInfo(ByVal PredefinedKey As RegKey, ByVal KeyName As String) As Variant
  Dim hKey As Long, Class As String, RegResult As Long, Values As Long
  Dim SecurityDescriptor As Long, LastWriteTime As Currency, Subkeys As Long
  
  On Error GoTo ERR_GetKeyInfo
  
    If RegOk(RegConnectRegistry(vbNullString, PredefinedKey, hKey), RegResult) Then
        RegResult = RegOpenKeyEx(hKey, KeyName, 0, Reg_KeyQueryValue, hKey)
    End If

    If RegResult = 0 Then
        Class = String(260, Chr(0))
        RegQueryInfoKey hKey, Class, 0, 0, Subkeys, 0, 0, Values, 0, 0, SecurityDescriptor, LastWriteTime
        
        GetKeyInfo = Array(left$((Class & vbNullChar), _
            Abs(InStr(Class & vbNullChar, vbNullChar) - 1)), _
                Subkeys, Values, SecurityDescriptor, LastWriteTime)
    End If

GetKeyInfo_Continue:
RegCloseKey hKey
Exit Function

ERR_GetKeyInfo:
Resume GetKeyInfo_Continue
End Function

Public Function DeleteValueName(ByVal PredefinedKey As RegKey, ByVal KeyName As String, ByVal ValueName As String, Optional RegResult As Long) As Boolean
  Dim hKey As Long
  
  On Error GoTo ERR_DeleteValueName
  
    If RegOk(RegConnectRegistry(vbNullString, PredefinedKey, hKey), RegResult) Then
        If RegOk(RegOpenKeyEx(hKey, KeyName, 0, Reg_KeySetValue, hKey), RegResult) Then
            If RegOk(RegDeleteValue(hKey, ValueName), RegResult) Then DeleteValueName = True
        End If
    End If

DeleteValueName_Continue:
If hKey <> 0 Then RegCloseKey hKey
Exit Function
    
ERR_DeleteValueName:
Resume DeleteValueName_Continue

End Function

Public Sub RegEnumAbort(AbortFlag As RegEnumAbortFlags)
    
    Select Case AbortFlag
        Case Abort_SubKeys: m_AbortEnum = True
        Case Abort_ValueNames: m_Abort_ValueNames = True
        Case Abort_KeyDelete: m_Abort_KeyDelete = True
    End Select

End Sub

' Usage SetValueData
' ---------------------------------------------------------
' 1)    Dim bData() As Byte
'        bData = "String" & vbNullChar
'        Call SetValueData(...Valuetype:=ValBinary, ValueName:="Binary", Value:=bData)
'
' 2)    Call SetValueData(...Valuetype:=ValStringEx, ValueName:="ExpandedString", Value:="%TEMP%")
'
' 3)    Dim vData As Variant
'        vData = Array("String1", "String2", "String3")
'        Call SetValueData(...Valuetype:=ValMultiString, ValueName:="MultipleString", Value:=vData)

Public Function SetValueData(ByVal PredefinedKey As RegKey, ByVal KeyName As String, ByVal ValueName As String, ByVal ValueType As RegValue, ByVal Value As Variant) As Boolean
  Dim hKey As Long, dwDisposition As Long, cbData As Long, SA As SecurityAttributes
  
    If Not KeyExist(PredefinedKey, KeyName) Then
        CreateKey PredefinedKey, KeyName
    End If
    
    On Error GoTo ERR_SetValueData
  
    If RegOk(RegConnectRegistry(vbNullString, PredefinedKey, hKey)) Then
        If RegOk(RegCreateKeyEx(hKey, KeyName, 0, "", Reg_OptionNonVolatile, Reg_KeyAllAccess, SA, hKey, dwDisposition)) Then
        
            Select Case ValueType
            
                Case ValLong, ValBoolean
                    If Len(CStr(Value)) = 0 Then GoSub SetValueData_Continue
                    RegSetValueExLong hKey, ValueName, 0, ValLong, CLng(Value), 4

                Case ValMultiString
                    Dim vData As Variant, sData As String
                    If IsArray(Value) Then
                        For Each vData In Value
                            If vData <> "" And vData <> vbNullChar And Not IsNull(vData) Then
                              sData = sData & vData & vbNullChar
                            End If
                        Next vData
                    End If
                    sData = sData & vbNullChar
                    cbData = Len(sData)
                    RegSetValueExString hKey, ValueName, 0, ValMultiString, sData, cbData
                      
                Case ValBinary
                    If Len(CStr(Value)) = 0 Then GoSub SetValueData_Continue
                    Dim bData() As Byte
                    bData = Value
                        
                    cbData = UBound(bData) - LBound(bData) - 1
                    RegSetValueExByte hKey, ValueName, 0, ValBinary, bData(0), cbData
                                          
                Case ValString  '(Default)
                    cbData = CLng(Len(Value) + Abs(CBool(Len(Value))))
                    RegSetValueExString hKey, ValueName, 0, ValString, CStr(Value), cbData
                    
            End Select

            SetValueData = True
        
        End If
    End If
    
    
SetValueData_Continue:
If hKey <> 0 Then RegCloseKey hKey
Exit Function

ERR_SetValueData:
SetValueData = False
Resume SetValueData_Continue
End Function

Public Function GetValueData(ByVal PredefinedKey As RegKey, _
                                            ByVal KeyName As String, _
                                            Optional ByVal ValueName As String, _
                                            Optional ValueType As RegValue = ValLong, _
                                            Optional ByVal ShowHexAndBinary As Boolean, _
                                            Optional ErrDefault As Variant) As Variant
  
Dim hKey As Long, lpValuetype As RegValue, cbData As Long
Dim ValueData As Variant

    On Error GoTo ERR_GetValueData
    
    If RegConnectRegistry(vbNullString, PredefinedKey, hKey) <> 0 Then
        GoTo ERR_GetValueData
    ElseIf RegOpenKeyEx(hKey, KeyName, 0, Reg_KeyAllAccess, hKey) <> 0 Then
        GoTo ERR_GetValueData
    ElseIf RegQueryValueEx(hKey, ValueName, 0, lpValuetype, 0, cbData) <> 0 Then
        GoTo ERR_GetValueData
    End If
            
    Select Case lpValuetype
        
        Case ValString, ValStringEx, ValMultiString
            Dim sData As String, s As String, C As Long
            
            sData = String(cbData, 0)
            If RegQueryValueExString(hKey, ValueName, 0, lpValuetype, sData, cbData) <> 0 _
                Then GoTo ERR_GetValueData
                
            If lpValuetype = ValString Then

                Dim sValue As String
                
                sValue = left$((sData & vbNullChar), Abs(InStr(sData & vbNullChar, vbNullChar) - 1))
                
                GetValueData = IIf(sValue <> "", sValue, CStr(ErrDefault))

            ElseIf lpValuetype = ValStringEx Then
                s = vbEmpty
                C = ExpandEnvironmentStrings(sData, s, C)
                s = String(C - 1, 0)
                C = ExpandEnvironmentStrings(sData, s, C)
                GetValueData = left$((s & vbNullChar), Abs(InStr(s & vbNullChar, vbNullChar) - 1))
            Else
                GetValueData = sData
            End If
    
        Case ValBinary
            Dim i As Long, A As String, b As String
            Dim bData() As Byte
            
            ReDim bData(cbData)
            
            If RegQueryValueExByte(hKey, ValueName, 0, lpValuetype, bData(0), cbData) <> 0 _
                Then GoTo ERR_GetValueData
            
            ' Adjust Variant(Byte) to CByte
            ValueData = bData
            ReDim Preserve bData(0 To (Len(ValueData) * 2) + 1) As Byte
            ValueData = bData
            
            For i = LBound(ValueData) To (UBound(ValueData) - 2)
                b = b & "&H"
                If (i < &H10) Then b = b & "0"
                b = b & Hex$(ValueData(i)) & " "
            Next i

            If ShowHexAndBinary Then
                GetValueData = Replace(b, " &H0", "")
            ElseIf ValueType = ValBoolean Then
                GetValueData = CBool(ValueData)
            ElseIf ValueType = ValLong Then
                GetValueData = CLng(Val(Replace(b, " &H0", "")))
            Else
                GetValueData = ValueData
            End If
            
        Case ValLong
            Dim dwData As Long
            
            If RegQueryValueExLong(hKey, ValueName, 0, lpValuetype, dwData, cbData) <> 0 _
                Then GoTo ERR_GetValueData
            
            If ShowHexAndBinary Then
                GetValueData = "&H" & Hex(dwData)
            ElseIf ValueType = ValBoolean Then
                GetValueData = CBool(dwData)
            Else
                GetValueData = dwData
            End If
            
        Case Else ' Unknown, read as string.
            sData = String(cbData, 0)
            If RegQueryValueExString(hKey, ValueName, 0, lpValuetype, sData, cbData) <> 0 Then
                GoTo ERR_GetValueData
            End If
            
            GetValueData = left(sData, cbData)
    
    End Select
    
    ValueType = lpValuetype
        

GetValueData_Continue:
If hKey <> 0 Then RegCloseKey hKey
Exit Function
    
ERR_GetValueData:
If IsMissing(ErrDefault) Then
    ' (Null) Value Not Set
    Select Case ValueType
        Case ValBoolean: GetValueData = False
        Case ValString, ValStringEx, ValMultiString
            GetValueData = vbNullString
        Case Else
            If ValueType = ValBinary And ShowHexAndBinary Then
                GetValueData = vbNullString
            Else
                GetValueData = 0
            End If
    End Select

Else
    Select Case TypeName(ErrDefault)
        Case "String": GetValueData = CStr(ErrDefault)
        Case "Integer": GetValueData = CInt(ErrDefault)
        Case "Long": GetValueData = CLng(ErrDefault)
        Case "Boolean": GetValueData = CBool(ErrDefault)
        Case "Byte": GetValueData = CByte(ErrDefault)
        Case "Date": GetValueData = CDate(ErrDefault)
        Case "Single": GetValueData = CSng(ErrDefault)
        Case "Double": GetValueData = CDbl(ErrDefault)
        Case "Currency": GetValueData = CCur(ErrDefault)
        Case Else: GetValueData = ErrDefault
    End Select
End If
If Err.Number <> 0 Then
    Resume GetValueData_Continue
Else
    GoTo GetValueData_Continue
End If
End Function

'   Windows 95/98: The DeleteKey function deletes a key, (Win95/98 ONLY: All Subkeys Included).
'   Windows NT/XP: This function deletes the specified key and all subkeys.
' CLASS MODULE UPDATED: 1/13/2003
Public Function DeleteKey(ByVal PredefinedKey As RegKey, ByVal KeyName As String, Optional RegResult As Long) As Boolean
    Dim SubKey_Num As Integer, X As Integer
    Dim SubKey As String, Cancel As Boolean
    Dim length As Long, hKey As Long
    Dim SubKeyCnt As Integer
    Static PredefinedHKey2 As RegKey
        
    On Error GoTo ERR_DeleteKey
    
    If PredefinedKey <> 0 Then
        ' First level search, reset Abort flag
        
        ' Windows versions 95/98 don't require empty
        ' subkey deletions.
        If Is9XRunning Then GoSub Root_Key_Delete
        
        m_Abort_KeyDelete = False
        PredefinedHKey2 = PredefinedKey
    End If
    
    If RegOk(RegConnectRegistry(vbNullString, PredefinedHKey2, hKey), RegResult) Then
        If RegOk(RegOpenKeyEx(hKey, KeyName, 0, Reg_KeyEnumSubkey, hKey), RegResult) Then
        
            SubKey_Num = 0
            Do While RegResult = 0 And (Not m_Abort_KeyDelete)
            
                length = 256
                SubKey = Space$(length)
                
                If RegOk(RegEnumKey(hKey, SubKey_Num, SubKey, length), RegResult) Then
                    SubKey = left$((SubKey & vbNullChar), Abs(InStr(SubKey & vbNullChar, vbNullChar) - 1))
                    
                    If m_Abort_KeyDelete Then
                        GoSub DeleteKey_Continue
                    Else
                        DeleteKey 0, KeyName & "\" & SubKey
                    End If
                    SubKey_Num = SubKey_Num + 1

                    '''' Descending order ''''
                    ' Add subkeys to delete Array
                    SubKeyCnt = pEnumCount(2)
                    ReDim Preserve m_DeleteKeys(0 To SubKeyCnt) As arrSubKeys
                    m_DeleteKeys(SubKeyCnt).sKeyName = KeyName
                    m_DeleteKeys(SubKeyCnt).sSubKey = SubKey

                End If
                
            Loop

        End If
    End If

If hKey <> 0 Then
    RegCloseKey hKey: hKey = 0: SubKey = vbNullString
    
    If PredefinedKey <> 0 Then
        ' Loop through each subkey scheduled for delete
        For X = 0 To (pEnumCount(2) - 1)
            If RegOk(RegConnectRegistry(vbNullString, PredefinedKey, hKey)) Then
                SubKey = (m_DeleteKeys(X).sKeyName & "\" & m_DeleteKeys(X).sSubKey)

                RaiseEvent KeyDelete(SubKey, Cancel)
                If Not Cancel Then
                    If Not RegOk(RegDeleteKey(hKey, SubKey)) Then
                        RegCloseKey hKey
                        GoTo ERR_DeleteKey
                    End If
                End If
                RegCloseKey hKey: hKey = 0
            Else
                GoTo ERR_DeleteKey
            End If
        Next X
        
' Delete the empty top key
Root_Key_Delete:
        Cancel = False
        If RegOk(RegConnectRegistry(vbNullString, PredefinedKey, hKey)) Then
            RaiseEvent KeyDelete(KeyName, Cancel)
            If Not Cancel Then
                If Not RegOk(RegDeleteKey(hKey, KeyName)) Then
                    RegCloseKey hKey
                    DeleteKey = (Not m_Abort_KeyDelete)
                    GoTo ERR_DeleteKey
                End If
            End If
        End If
        
        PredefinedHKey2 = 0
        Erase m_DeleteKeys
    End If
    
ElseIf RegResult = 2 Then
    DeleteKey = True
End If

DeleteKey_Continue:
If hKey <> 0 Then RegCloseKey hKey
Err.Clear
Exit Function

ERR_DeleteKey:
Erase m_DeleteKeys
DeleteKey = False
If Err.Number <> 0 Then Resume DeleteKey_Continue
Exit Function

End Function

Public Function EnumSubKeys(ByVal PredefinedKey As RegKey, ByVal KeyName As String, Optional ByVal MaxLevel As Integer, Optional RegResult As Long) As Boolean
    Dim SubKey_Num As Integer, X As Integer
    Dim SubKey As String
    Dim length As Long, hKey As Long
    Dim Level As Integer, SubKeyCnt As Integer
    Static PredefinedHKey1 As RegKey
        
    On Error GoTo ERR_EnumSubKeys
    
    If PredefinedKey <> 0 Then
        ' First level search, reset Abort flag
        m_AbortEnum = False
        PredefinedHKey1 = PredefinedKey
    End If
    
    If RegOk(RegConnectRegistry(vbNullString, PredefinedHKey1, hKey), RegResult) Then
        If RegOk(RegOpenKeyEx(hKey, KeyName, 0, Reg_KeyEnumSubkey, hKey), RegResult) Then
        
            SubKey_Num = 0
            Do While RegResult = 0 And (Not m_AbortEnum)
            
                length = 256
                SubKey = Space$(length)
                
                If RegOk(RegEnumKey(hKey, SubKey_Num, SubKey, length), RegResult) Then
                    SubKey = left$((SubKey & vbNullChar), Abs(InStr(SubKey & vbNullChar, vbNullChar) - 1))
                    
                    ' Subkey found for KeyName
                    If Not m_AbortEnum Then

                        '''' Ascending order to add first
                        ' Fill the array with each subkey found
                        SubKeyCnt = pEnumCount(1)
                        ReDim Preserve m_EnumKeys(0 To SubKeyCnt) As arrSubKeys
                        m_EnumKeys(SubKeyCnt).sKeyName = KeyName
                        m_EnumKeys(SubKeyCnt).sSubKey = SubKey

                        If MaxLevel > 0 Then
                            If MaxLevel > 1 Then
                                EnumSubKeys 0, KeyName & "\" & SubKey, MaxLevel - 1
                            End If
                        Else
                            EnumSubKeys 0, KeyName & "\" & SubKey
                        End If
                    End If
                    SubKey_Num = SubKey_Num + 1

                End If
                
            Loop
            
            EnumSubKeys = (Not m_AbortEnum)

        End If
    End If

If hKey <> 0 Then
    RegCloseKey hKey
    If PredefinedKey <> 0 Then
        For X = 0 To (pEnumCount(1) - 1)
            ' Raise event for every subkey found
            RaiseEvent Subkeys(m_EnumKeys(X).sKeyName, m_EnumKeys(X).sSubKey)
        Next X
        Erase m_EnumKeys
        PredefinedHKey1 = 0
    End If
End If

EnumSubKeys_Continue:
Err.Clear
Exit Function

ERR_EnumSubKeys:
EnumSubKeys = False
Resume EnumSubKeys_Continue

End Function
Private Function pEnumCount(ByVal nID As Integer) As Long

 On Error GoTo ERR_pEnumCount

Dim tmp_Cnt As Long

If nID = 1 Then
    tmp_Cnt = UBound(m_EnumKeys)
ElseIf nID = 2 Then
    tmp_Cnt = UBound(m_DeleteKeys)
Else
    GoTo pEnumCount_Exit
End If

pEnumCount = (tmp_Cnt + 1)

pEnumCount_Exit:

Err.Clear
Exit Function

ERR_pEnumCount:
pEnumCount = 0
Resume pEnumCount_Exit

End Function
Private Function RegOk(dwIn As Long, Optional RegResult As Long) As Boolean
  If (dwIn = 0) Then
    RegOk = True
  End If
  RegResult = dwIn
End Function
Public Sub EnumValueNames(ByVal PredefinedKey As RegKey, ByVal KeyName As String, Optional RegResult As Long)
  Dim hKey As Long, sBuffer As String, cBuffer As Long
  Dim dwIndex As Long, dwType As Long

    On Error GoTo ERR_EnumValueNames
    
    If RegOk(RegConnectRegistry(vbNullString, PredefinedKey, hKey), RegResult) Then
        RegResult = RegOpenKeyEx(hKey, KeyName, 0, Reg_KeyAllAccess, hKey)
    End If
    
    Do While RegResult = 0 And (Not m_Abort_ValueNames)
        cBuffer = 256: sBuffer = Space(cBuffer)
        If RegOk(RegEnumValue(hKey, dwIndex, sBuffer, cBuffer, 0, dwType, ByVal 0, 0), RegResult) Then
            RaiseEvent ValueNames(KeyName, Mid(sBuffer, 1, cBuffer))
            dwIndex = dwIndex + 1
        End If
    Loop
    
EnumValueNames_Continue:
If hKey <> 0 Then RegCloseKey hKey
m_Abort_ValueNames = False
Exit Sub

ERR_EnumValueNames:
Resume EnumValueNames_Continue

End Sub



