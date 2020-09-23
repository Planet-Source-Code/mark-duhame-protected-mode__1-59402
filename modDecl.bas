Attribute VB_Name = "modDecl"
Option Explicit

Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function VirtualLock Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long) As Long
Public Declare Function VirtualUnlock Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal Classname As String, ByVal WindowName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function ReadProcessMem Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub ArrayDescriptor Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc() As Any, ByVal ByteLen As Long)
Public Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Public Declare Function GetCurrentThread Lib "kernel32" () As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" _
                         (ByVal lpLibFileName As String) As Long
'Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
                          ByVal lpProcName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal L As Long)
    
Public Type Trustee
     pMultipleTrustee As Long
     MultipleTrusteeOperation As Long
     TrusteeForm As Long
     TrusteeType As Long
     ptstrName As String
End Type

'MULTIPLE_TRUSTEE_OPERATION
Public Const NO_MULTIPLE_TRUSTEE = 0
'Public Const TRUSTEE_IS_IMPERSONATE = 1

'TRUSTEE_FORM
'Public Const TRUSTEE_IS_SID = 0
Public Const TRUSTEE_IS_NAME = 1
'Public Const TRUSTEE_BAD_FORM = 2
 
'TRUSTEE_TYPE
'Public Const TRUSTEE_IS_UNKNOWN = 0
Public Const TRUSTEE_IS_USER = 1
'Public Const TRUSTEE_IS_GROUP = 2
'Public Const TRUSTEE_IS_DOMAIN = 3
'Public Const TRUSTEE_IS_ALIAS = 4
'Public Const TRUSTEE_IS_WELL_KNOWN_GROUP = 5
'Public Const TRUSTEE_IS_DELETED = 6
'Public Const TRUSTEE_IS_INVALID = 7

Public Type EXPLICIT_ACCESS
     grfAccessPermissions As Long
     grfAccessMode  As Long
     grfInheritance As Long
     Trustee As Trustee
End Type

'ACCESS_MODE
'Public Const NOT_USED_ACCESS = 0
Public Const GRANT_ACCESS = 1
'Public Const SET_ACCESS = 2
'Public Const DENY_ACCESS = 3
'Public Const REVOKE_ACCESS = 4
'Public Const SET_AUDIT_SUCCESS = 5
'Public Const SET_AUDIT_FAILURE = 6
Public Const PROCESS_ALL_ACCESS As Long = &H1F0FFF

'Inheritance flags
Public Const NO_INHERITANCE = 0
'Public Const SUB_OBJECTS_ONLY_INHERIT = 1
'Public Const SUB_CONTAINERS_ONLY_INHERIT = 2
'Public Const SUB_CONTAINERS_AND_OBJECTS_INHERIT = 3
'Public Const INHERIT_NO_PROPAGATE = 4
'Public Const INHERIT_ONLY = 8
'Public Const INHERITED_ACCESS_ENTRY = &H10
'Public Const INHERITED_PARENT = &H10000000
'Public Const INHERITED_GRANDPARENT = &H20000000

'SE_OBJECT_TYPE
'Public Const SE_UNKNOWN_OBJECT_TYPE = 0
'Public Const SE_FILE_OBJECT = 1
'Public Const SE_SERVICE = 2
'Public Const SE_PRINTER = 3
'Public Const SE_REGISTRY_KEY = 4
'Public Const SE_LMSHARE = 5
'Public Const SE_KERNEL_OBJECT = 6
'Public Const SE_WINDOW_OBJECT = 7
'Public Const SE_DS_OBJECT = 8
'Public Const SE_DS_OBJECT_ALL = 9
'Public Const SE_PROVIDER_DEFINED_OBJECT = 10
'Public Const OWNER_SECURITY_INFORMATION = &H1
'Public Const GROUP_SECURITY_INFORMATION = &H2
'Public Const DACL_SECURITY_INFORMATION = &H4
'Public Const SACL_SECURITY_INFORMATION = &H8

'Public Const ALL_SECURITY_INFORMATION = OWNER_SECURITY_INFORMATION Or GROUP_SECURITY_INFORMATION Or DACL_SECURITY_INFORMATION Or SACL_SECURITY_INFORMATION

'Public Const SECTION_QUERY = &H1
Public Const SECTION_MAP_WRITE = &H2
'Public Const SECTION_MAP_READ = &H4
'Public Const SECTION_MAP_EXECUTE = &H8
'Public Const SECTION_EXTEND_SIZE = &H10

'Public Const THREAD_BASE_PRIORITY_IDLE = -15
Public Const THREAD_BASE_PRIORITY_LOWRT = 15
'Public Const THREAD_BASE_PRIORITY_MIN = -2
'Public Const THREAD_BASE_PRIORITY_MAX = 2
'Public Const THREAD_PRIORITY_LOWEST = THREAD_BASE_PRIORITY_MIN
'Public Const THREAD_PRIORITY_HIGHEST = THREAD_BASE_PRIORITY_MAX
'Public Const THREAD_PRIORITY_BELOW_NORMAL = (THREAD_PRIORITY_LOWEST + 1)
'Public Const THREAD_PRIORITY_ABOVE_NORMAL = (THREAD_PRIORITY_HIGHEST - 1)
'Public Const THREAD_PRIORITY_IDLE = THREAD_BASE_PRIORITY_IDLE
'Public Const THREAD_PRIORITY_NORMAL = 0
Public Const THREAD_PRIORITY_TIME_CRITICAL = THREAD_BASE_PRIORITY_LOWRT
Public ASMDis As New DisAsm
'Public Const HIGH_PRIORITY_CLASS = &H80
'Public Const IDLE_PRIORITY_CLASS = &H40
'Public Const NORMAL_PRIORITY_CLASS = &H20
'Public Const REALTIME_PRIORITY_CLASS = &H100
'Public Const PC_WRITEABLE = &H20000
'Public Const PC_USER = &H40000
'Public Const PC_PRESENT = &H80000000
'Public Const PC_STATIC = &H20000000
'Public Const PC_DIRTY = &H8000000
'Public Const PageModifyPermissions = &H1000D
Public DataArr() As Byte 'Used for Ram Array

'Function Name ReadMem
'
'
'Purpose: Read Ramdom Access Memory
'
'Parameters: Offset as Long = Address to read
'          : max as Integer = Number of bytes to read
'
'Return String Containing memory read
Public Function ReadMem(Offset As Long, ByVal max As Integer) As String
    Dim hwnd As Long
    Dim ProcessID As Long
    Dim ProcessHandle As Long
    
    'Try to find the window that was passed in the variable WindowName to this function.
    hwnd = FindWindow(vbNullString, "")
    
    If hwnd = 0 Then
        'This is executed if the window cannot be found.
        'You can add or write own code here to customize your program.
        ReadMem = ""
        Exit Function
    End If
    
    'Get the window's process ID.
    GetWindowThreadProcessId hwnd, ProcessID
    
    'Get a process handle.
    ProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)
    
    If ProcessHandle = 0 Then
        'This is executed if a process handle cannot be found.
        'You can add or write your own code here to customize your program.
        ReadMem = ""
        Exit Function
    End If
    ReDim DataArr(max)
    'Read a BYTE value from the specified memory offset.
    ReadProcessMem ProcessHandle, Offset, DataArr(0), max, 0&
    
    'Return the found memory value.
    ReadMem = StrConv(DataArr, vbUnicode)
    
    'It is important to close the current process handle.
    CloseHandle ProcessHandle
           
End Function
