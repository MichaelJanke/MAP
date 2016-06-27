Attribute VB_Name = "basCheckAccessRights"
Option Explicit
'   ********************************************
'   *          2000 Sergey Merzlikin          *
'   *  source: http://www.freevbcode.com/ShowCode.asp?ID=4318
'   ********************************************

' Desired access rights constants
Public Const MAXIMUM_ALLOWED As Long = &H2000000
Public Const DELETE As Long = &H10000
Public Const READ_CONTROL As Long = &H20000
Public Const WRITE_DAC As Long = &H40000
Public Const WRITE_OWNER As Long = &H80000
Public Const SYNCHRONIZE As Long = &H100000
Public Const STANDARD_RIGHTS_READ As Long = READ_CONTROL
Public Const STANDARD_RIGHTS_WRITE As Long = READ_CONTROL
Public Const STANDARD_RIGHTS_EXECUTE As Long = READ_CONTROL
Public Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Public Const FILE_READ_DATA As Long = &H1 ' file & pipe
Public Const FILE_LIST_DIRECTORY As Long = &H1 ' directory
Public Const FILE_ADD_FILE As Long = &H2 ' directory
Public Const FILE_WRITE_DATA As Long = &H2 ' file & pipe
Public Const FILE_CREATE_PIPE_INSTANCE As Long = &H4 ' named pipe
Public Const FILE_ADD_SUBDIRECTORY As Long = &H4 ' directory
Public Const FILE_APPEND_DATA As Long = &H4 ' file
Public Const FILE_READ_EA As Long = &H8 ' file & directory
Public Const FILE_READ_PROPERTIES As Long = FILE_READ_EA
Public Const FILE_WRITE_EA As Long = &H10 ' file & directory
Public Const FILE_WRITE_PROPERTIES As Long = FILE_WRITE_EA
Public Const FILE_EXECUTE As Long = &H20 ' file
Public Const FILE_TRAVERSE As Long = &H20 ' directory
Public Const FILE_DELETE_CHILD As Long = &H40 ' directory
Public Const FILE_READ_ATTRIBUTES As Long = &H80 ' all
Public Const FILE_WRITE_ATTRIBUTES As Long = &H100 ' all
Public Const FILE_GENERIC_READ As Long = (STANDARD_RIGHTS_READ _
  Or FILE_READ_DATA Or FILE_READ_ATTRIBUTES _
  Or FILE_READ_EA Or SYNCHRONIZE)
Public Const FILE_GENERIC_WRITE As Long = (STANDARD_RIGHTS_WRITE _
  Or FILE_WRITE_DATA Or FILE_WRITE_ATTRIBUTES _
  Or FILE_WRITE_EA Or FILE_APPEND_DATA Or SYNCHRONIZE)
Public Const FILE_GENERIC_EXECUTE As Long = (STANDARD_RIGHTS_EXECUTE _
  Or FILE_READ_ATTRIBUTES Or FILE_EXECUTE Or SYNCHRONIZE)
Public Const FILE_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED _
  Or SYNCHRONIZE Or &H1FF&)
Public Const GENERIC_READ As Long = &H80000000
Public Const GENERIC_WRITE As Long = &H40000000
Public Const GENERIC_EXECUTE As Long = &H20000000
Public Const GENERIC_ALL As Long = &H10000000

' Types, constants and functions
' to work with access rights
Public Const OWNER_SECURITY_INFORMATION As Long = &H1
Public Const GROUP_SECURITY_INFORMATION As Long = &H2
Public Const DACL_SECURITY_INFORMATION As Long = &H4
Public Const TOKEN_QUERY As Long = 8
Public Const SecurityImpersonation As Integer = 3
Public Const ANYSIZE_ARRAY = 1
Public Type GENERIC_MAPPING
  GenericRead As Long
  GenericWrite As Long
  GenericExecute As Long
  GenericAll As Long
End Type
Public Type LUID
  LowPart As Long
  HighPart As Long
End Type
Public Type LUID_AND_ATTRIBUTES
  pLuid As LUID
  Attributes As Long
End Type
Public Type PRIVILEGE_SET
  PrivilegeCount As Long
  Control As Long
  Privilege(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type
Public Declare Function GetFileSecurity Lib "advapi32.dll" _
  Alias "GetFileSecurityA" (ByVal lpFileName As String, _
  ByVal RequestedInformation As Long, pSecurityDescriptor As Byte, _
  ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Public Declare Function AccessCheck Lib "advapi32.dll" _
  (pSecurityDescriptor As Byte, ByVal ClientToken As Long, _
  ByVal DesiredAccess As Long, GenericMapping As GENERIC_MAPPING, _
  PrivilegeSet As PRIVILEGE_SET, PrivilegeSetLength As Long, _
  GrantedAccess As Long, Status As Long) As Long
Public Declare Function ImpersonateSelf Lib "advapi32.dll" _
  (ByVal ImpersonationLevel As Integer) As Long
Public Declare Function RevertToSelf Lib "advapi32.dll" () As Long
Public Declare Sub MapGenericMask Lib "advapi32.dll" (AccessMask As Long, _
  GenericMapping As GENERIC_MAPPING)
Public Declare Function OpenThreadToken Lib "advapi32.dll" _
  (ByVal ThreadHandle As Long, ByVal DesiredAccess As Long, _
  ByVal OpenAsSelf As Long, TokenHandle As Long) As Long
Public Declare Function GetCurrentThread Lib "kernel32" () As Long
Public Declare Function CloseHandle Lib "kernel32" _
  (ByVal hObject As Long) As Long

' Types, constants and functions for OS version detection
Public Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type
Public Const VER_PLATFORM_WIN32_NT As Long = 2
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
  (lpVersionInformation As OSVERSIONINFO) As Long

' Constant and function for detection of support
' of access rights by file system
Public Const FS_PERSISTENT_ACLS As Long = &H8
Public Declare Function GetVolumeInformation Lib "kernel32" _
  Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, _
  ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, _
  lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, _
  lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, _
  ByVal nFileSystemNameSize As Long) As Long


' *-----------------------------------------------------------------------*
' CheckFileAccess function checks access rights to given file.
' DesiredAccess - bitmask of desired access rights.
' The function returns bitmask, which contains those bits of desired bitmask,
' which correspond with existing access rights.
Public Function CheckFileAccess(Filename As String, _
    ByVal DesiredAccess As Long) As Long
  Dim r As Long, SecDesc() As Byte, SDSize As Long, hToken As Long
  Dim PrivSet As PRIVILEGE_SET, GenMap As GENERIC_MAPPING
  Dim Volume As String, FSFlags As Long
  Dim strStatus As String
  On Error GoTo err_CheckFileAccess
  ' Checking OS type
  strStatus = "Check OS Type"
  If Not IsNT() Then
    ' Rights not supported. Returning -1.
'''    MsgBox "CheckFileAccess:Not Win32"
    CheckFileAccess = -1:       Exit Function
  End If
  ' Checking access rights support by file system
  strStatus = "Check acces rights support by file system"
  If Left$(Filename, 2) = "\\" Then
    ' Path in UNC format. Extracting share name from it
    r = InStr(3, Filename, "\")
    If r = 0 Then
      Volume = Filename & "\"
    Else
      Volume = Left$(Filename, r)
    End If
  ElseIf Mid$(Filename, 2, 2) = ":\" Then
    ' Path begins with drive letter
    Volume = Left$(Filename, 3)
    'Else
    ' If path not set, we are leaving Volume blank.
    ' It retutns information about current drive.
  End If
  ' Getting information about drive
  strStatus = "Get information about drive"
  GetVolumeInformation Volume, vbNullString, 0, ByVal 0&, _
    ByVal 0&, FSFlags, vbNullString, 0
  If (FSFlags And FS_PERSISTENT_ACLS) = 0 Then
    ' Rights not supported. Returning -1.
'''    MsgBox "CheckFileAccess:Rights not Supported"
    CheckFileAccess = -1:       Exit Function
  End If
  ' Determination of buffer size
  strStatus = "Determination of buffer size"
  GetFileSecurity Filename, OWNER_SECURITY_INFORMATION _
    Or GROUP_SECURITY_INFORMATION _
    Or DACL_SECURITY_INFORMATION, 0, 0, SDSize
  If Err.LastDllError <> 122 Then
    ' Rights not supported. Returning -1.
'''    MsgBox "CheckFileAccess:Rights not Supported_2"
    CheckFileAccess = -1:       Exit Function
  End If
  If SDSize = 0 Then Exit Function
  ' Buffer allocation
  strStatus = "Buffer allocation"
  ReDim SecDesc(1 To SDSize)
  ' Once more call of function
  ' to obtain Security Descriptor
  strStatus = "GetFileSecurity"
  If GetFileSecurity(Filename, OWNER_SECURITY_INFORMATION _
    Or GROUP_SECURITY_INFORMATION _
    Or DACL_SECURITY_INFORMATION, _
    SecDesc(1), SDSize, SDSize) = 0 Then
    ' Error. We must return no access rights.
'''    MsgBox "CheckFileAccess:no access rights"
    CheckFileAccess = -1:       Exit Function
  End If
  ' Adding Impersonation Token for thread
  strStatus = "Impersonate"
  ImpersonateSelf SecurityImpersonation
  ' Opening of Token of current thread
  strStatus = "Open token"
  OpenThreadToken GetCurrentThread(), TOKEN_QUERY, 0, hToken
  If hToken <> 0 Then
'''    MsgBox "CheckFileAccess:hToken <> 0  -> Fill GenericMask"
    ' Filling GenericMask type
    strStatus = "Fill generic mask"
    GenMap.GenericRead = FILE_GENERIC_READ
    GenMap.GenericWrite = FILE_GENERIC_WRITE
    GenMap.GenericExecute = FILE_GENERIC_EXECUTE
    GenMap.GenericAll = FILE_ALL_ACCESS
    ' Conversion of generic rights
    ' to specific file access rights
    strStatus = "MapGenericMask"
    MapGenericMask DesiredAccess, GenMap
    If Err.LastDllError <> 0 Then
'''      MsgBox "CheckFileAccess:MapGenericMask failed with " & Err.LastDllError
      CheckFileAccess = -1:            Exit Function
    End If
    ' Checking access
    strStatus = "AccessCheck"
    AccessCheck SecDesc(1), hToken, DesiredAccess, GenMap, PrivSet, Len(PrivSet), CheckFileAccess, r
    If Err.LastDllError <> 0 Then
'''      MsgBox "CheckFileAccess:AccessCheck failed with " & Err.LastDllError
      CheckFileAccess = -1:            Exit Function
    End If
    strStatus = "CloseHandle"
    CloseHandle hToken
  Else
'''    MsgBox "CheckFileAccess:hToken = 0  -> No Fill GenericMask"
  End If
  ' Deleting Impersonation Token
  RevertToSelf
''' MsgBox "CheckFileAccess:return:" & CheckFileAccess
  Exit Function
err_CheckFileAccess:
    MsgBox "ERROR /" & strStatus & "/Err:" & Err.Number & " Desc:" & Err.Description, , "CheckFileAccess"
End Function

' *-----------------------------------------------------------------------*
' IsNT() function returns True, if the program works
' in Windows NT or Windows 2000 operating system, and False
' otherwise.
Private Function IsNT() As Boolean
  Dim OSVer As OSVERSIONINFO
  OSVer.dwOSVersionInfoSize = Len(OSVer)
  GetVersionEx OSVer
  IsNT = (OSVer.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

' *-----------------------------------------------------------------------*
Public Function CanWrite() As Boolean
    Dim strFolder As String
    strFolder = App.Path
    
    Dim AccessRead As Boolean, AccessWrite As Boolean
    Dim AccessMask As Long
''    MsgBox "CanWrite:Folder:" & strFolder
    AccessMask = CheckFileAccess(strFolder, MAXIMUM_ALLOWED)
''    MsgBox "CanWrite:AccessMask:" & AccessMask
    If AccessMask = -1 Then
        CanWrite = True         ' Funktion hat Fehler zurückgegeben - Keine Aussage, dann lieber nicht blockieren
    Else
        AccessRead = (AccessMask _
                     And FILE_GENERIC_READ) = FILE_GENERIC_READ
        AccessWrite = (AccessMask _
                     And FILE_GENERIC_WRITE) = FILE_GENERIC_WRITE
        CanWrite = AccessWrite
    End If
    CanWrite = True
    
   
End Function


