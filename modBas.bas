Attribute VB_Name = "modBas"
Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal Classname As String, ByVal WindowName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function ReadProcessMem Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, lpOverlapped As Any) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Enum HKEYs
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

Private Const PROCESS_ALL_ACCESS As Long = &H1F0FFF
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL = &HFFFF

Private Const SYNCHRONIZE = &H100000
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const ERROR_SUCCESS = 0
Private Const REG_SZ = 1
Private Const ERROR_NO_MORE_ITEMS = 259&

Public Type DDB
    next As Long
    ver As Integer
    no As Integer
    Mver As String
    minorver As String
    flags As Integer
    iname() As Byte
    Init As Long
    ControlProc As Long
    v86_proc As Long
    Pm_Proc As Long
    v86() As Byte
    PM() As Byte
    Data As Long
    service_size As Long
    win32_ptr As Long
End Type

Public Type DIOC_REGISTERS
    reg_EBX As Long
    reg_EDX As Long
    reg_ECX As Long
    reg_EAX As Long
    reg_EDI As Long
    reg_ESI As Long
    reg_Flags As Long
End Type

Public Type outBuf
    send1 As Long ' varptr
    send2 As Long ' varptr
    nobytes As Integer
    stack(8) As Byte
End Type

Public out1(34) As Byte

'Holds Static list of VxD's as stored in the registry
Public VxDs() As String
'Hold dynamic list of VxD's - either loaded in system.ini file [386Enh] section
'or loaded unloaded as needed
'only avail on win 9x
Public DynVxDs() As String
Public Value() As Byte      'Holds Memory read

Public Function ReadByte(Offset As Long, WindowName As String, ByVal max As Long) As String
    Dim hwnd As Long
    Dim ProcessID As Long
    Dim ProcessHandle As Long
    
    'Try to find the window that was passed in the variable WindowName to this function.
    hwnd = FindWindow(vbNullString, WindowName)
    
    If hwnd = 0 Then
        
        'This is executed if the window cannot be found.
        'You can add or write own code here to customize your program.
        
        MsgBox "Could not find process window!", vbCritical, "Read error"
        
        Exit Function
    
    End If
    
    'Get the window's process ID.
    GetWindowThreadProcessId hwnd, ProcessID
    
    'Get a process handle.
    ProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)
    
    If ProcessHandle = 0 Then
        
        'This is executed if a process handle cannot be found.
        'You can add or write your own code here to customize your program.
        
        MsgBox "Could not get a process handle!", vbCritical, "Read error"
        
        Exit Function
    
    End If
    ReDim Value(max)
    'Read a BYTE value from the specified memory offset.
    ReadProcessMem ProcessHandle, Offset, Value(0), max, 0&
    
    'Return the found memory value.
    ReadByte = StrConv(Value, vbUnicode)
    'ReadByte = Value
    
    'It is important to close the current process handle.
    CloseHandle ProcessHandle
           
End Function

Public Function CallASM(addL As Long) As Long
    Dim tmpStr As String
    Dim hStr As String
    Dim callLocation As Long
    Dim x As Long
    Dim asmCode(16) As Byte
    Dim cnt As Integer
    
    'pointer in mem to our machine code
    'needed for CallWindowProc
    x = VarPtr(asmCode(0))
    
    asmCode(0) = &H55   ' push ebp
    
    asmCode(1) = &H8B   ' mov ebp, esp
    asmCode(2) = &HEC
    
    asmCode(3) = &H83   ' sub esp, 10
    asmCode(4) = &HEC
    asmCode(5) = &H10
    
    asmCode(6) = &HE8   ' call to service call
    'these next 4 must be calculated from
    'actual address of VxDCall - our machine
    'code location
    
    callLocation = (addL - x) - 11
    'callLocation = Not callLocation
    
    hStr = Hex(callLocation)
    hStr = String(8 - Len(hStr), "0") & hStr
    tmpStr = ""
    For cnt = 8 To 1 Step -2
        tmpStr = tmpStr + Mid(hStr, cnt - 1, 2)
    Next
    
    hStr = Left(tmpStr, 2)
    asmCode(7) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 3, 2)
    asmCode(8) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 5, 2)
    asmCode(9) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 7, 2)
    asmCode(10) = CLng("&H" & hStr)
    
    
    asmCode(11) = &H8B  ' mov esp, ebp
    asmCode(12) = &HE5
    
    asmCode(13) = &H5D  ' pop ebp
    
    asmCode(14) = &HC2  ' ret
    asmCode(15) = &H10
    asmCode(16) = &H0   ' just in case stack filled bith garbage
    
    CallASM = CallWindowProc(x, 0, 0, 0, 0)
    Erase asmCode
    
End Function

Public Function GetVxDStatic() As Long
    Dim VxDGroups As String
    Dim cnt As Integer
    Const BUFFER_SIZE As Long = 255
    Dim hKey As Long
    Dim sName As String
    Dim sData As String
    Dim ret As Long
    Dim pos As Integer
    Dim counter As Integer
    
    hKey = RegOpen(&H80000002, "System\CurrentControlSet\Services\VxD")
    cnt = 0
    If hKey <> 0 Then

        sName = Space(BUFFER_SIZE)
        'Enumerate the keys
        While RegEnumKeyEx(hKey, cnt, sName, ret, ByVal 0&, vbNullString, ByVal 0&, ByVal 0&) <> ERROR_NO_MORE_ITEMS
            'Show the enumerated key
            ReDim Preserve VxDs(cnt)
            pos = InStr(sName, Chr(0))
            If pos > 0 Then
                sName = Left(sName, pos - 1)
            End If
            VxDs(cnt) = sName
                
            cnt = cnt + 1
            sName = Space(BUFFER_SIZE)
            ret = BUFFER_SIZE
        Wend
        'close the registry key
        RegCloseKey hKey
        
    End If
    
    cnt = cnt - 1
    VxDGroups = RegGetKeyValue(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\VxD\VMM", "VxDGroups", cnt)
    sName = ""
    sData = ""
    counter = 0
   
End Function

Public Function RegOpen(hKey As Long, SubKey As String) As Long
    Dim lngRet As Long
    Dim lngResult As Long

    lngRet = RegOpenKeyEx(hKey, SubKey, 0, KEY_ALL_ACCESS, lngResult)
    RegOpen = lngResult

End Function

Public Function RegGetKeyValue(hKey As Long, SubKey As String, ValueName As String, vxdcnt As Integer, Optional Default As String = "")
    Dim lngRet As Long
    Dim lngResult As Long
    Dim sData As String
    
    lngRet = RegOpenKeyEx(hKey, SubKey, 0, KEY_ALL_ACCESS, lngResult)
    If lngRet = ERROR_SUCCESS Then
        sData = String(vxdcnt * 8, vbNullChar)
        
        lngRet = RegQueryValueEx(lngResult, ValueName, 0, REG_SZ, ByVal sData, Len(sData))
        
        If Not lngRet = ERROR_SUCCESS Then RegGetKeyValue = Default: Exit Function
        RegGetKeyValue = sData
        RegCloseKey lngResult
    Else
        RegGetKeyValue = Default
    End If
    
End Function

Public Function CallASMEBX(addL As Long) As Long
    Dim tmpStr As String
    Dim hStr As String
    Dim x As Long
    Dim asmCode(18) As Byte
    Dim cnt As Integer
    Dim callLocation As Long
    
    asmCode(0) = &H55
    asmCode(1) = &H8B
    asmCode(2) = &HEC
    asmCode(3) = &H83
    asmCode(4) = &HEC
    asmCode(5) = &H10
    asmCode(6) = &HE8
    
    'pointer in mem to our machine code
    x = VarPtr(asmCode(0))
    'these next 4 must be calculated from
    'actual address of VxDCall - our machine
    'code location
    
    callLocation = (addL - x) - 11
    'callLocation = Not callLocation
    
    hStr = Hex(callLocation)
    hStr = String(8 - Len(hStr), "0") & hStr
    tmpStr = ""
    For cnt = 8 To 1 Step -2
        tmpStr = tmpStr + Mid(hStr, cnt - 1, 2)
    Next
    
    hStr = Left(tmpStr, 2)
    asmCode(7) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 3, 2)
    asmCode(8) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 5, 2)
    asmCode(9) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 7, 2)
    asmCode(10) = CLng("&H" & hStr)
    
    asmCode(11) = &H8B
    asmCode(12) = &HC3
    asmCode(13) = &H8B
    asmCode(14) = &HE5
    asmCode(15) = &H5D
    asmCode(16) = &HC2
    asmCode(17) = &H10
    asmCode(18) = &H0    ' just in case stack filled bith garbage
    
    CallASMEBX = CallWindowProc(x, 0, 0, 0, 0)
    Erase asmCode
    
End Function

Public Function CallASMECX(addL As Long) As Long
    Dim tmpStr As String
    Dim hStr As String
    Dim x As Long
    Dim asmCode(18) As Byte
    Dim cnt As Integer
    Dim callLocation As Long
    
    asmCode(0) = &H55
    asmCode(1) = &H8B
    asmCode(2) = &HEC
    asmCode(3) = &H83
    asmCode(4) = &HEC
    asmCode(5) = &H10
    asmCode(6) = &HE8
    
    'pointer in mem to our machine code
    x = VarPtr(asmCode(0))
    'these next 4 must be calculated from
    'actual address of VxDCall - our machine
    'code location
    
    callLocation = (addL - x) - 11
    'callLocation = Not callLocation
    
    hStr = Hex(callLocation)
    hStr = String(8 - Len(hStr), "0") & hStr
    tmpStr = ""
    For cnt = 8 To 1 Step -2
        tmpStr = tmpStr + Mid(hStr, cnt - 1, 2)
    Next
    
    hStr = Left(tmpStr, 2)
    asmCode(7) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 3, 2)
    asmCode(8) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 5, 2)
    asmCode(9) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 7, 2)
    asmCode(10) = CLng("&H" & hStr)
    
    asmCode(11) = &H8B
    asmCode(12) = &HC1
    asmCode(13) = &H8B
    asmCode(14) = &HE5
    asmCode(15) = &H5D
    asmCode(16) = &HC2
    asmCode(17) = &H10
    asmCode(18) = &H0   ' just in case stack filled bith garbage
    
    CallASMECX = CallWindowProc(x, 0, 0, 0, 0)
    Erase asmCode
    
End Function

Public Function CallASMESI(addL As Long) As Long
    Dim tmpStr As String
    Dim hStr As String
    Dim x As Long
    Dim asmCode(18) As Byte
    Dim cnt As Integer
    Dim callLocation As Long
    
    asmCode(0) = &H55
    asmCode(1) = &H8B
    asmCode(2) = &HEC
    asmCode(3) = &H83
    asmCode(4) = &HEC
    asmCode(5) = &H10
    asmCode(6) = &HE8
    
    'pointer in mem to our machine code
    x = VarPtr(asmCode(0))
    'these next 4 must be calculated from
    'actual address of VxDCall - our machine
    'code location
    
    callLocation = (addL - x) - 11
    'callLocation = Not callLocation
    
    hStr = Hex(callLocation)
    hStr = String(8 - Len(hStr), "0") & hStr
    tmpStr = ""
    For cnt = 8 To 1 Step -2
        tmpStr = tmpStr + Mid(hStr, cnt - 1, 2)
    Next
    
    hStr = Left(tmpStr, 2)
    asmCode(7) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 3, 2)
    asmCode(8) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 5, 2)
    asmCode(9) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 7, 2)
    asmCode(10) = CLng("&H" & hStr)
    
    asmCode(11) = &H8B
    asmCode(12) = &HC6
    asmCode(13) = &H8B
    asmCode(14) = &HE5
    asmCode(15) = &H5D
    asmCode(16) = &HC2
    asmCode(17) = &H10
    asmCode(18) = &H0   ' just in case stack filled bith garbage
    
    CallASMESI = CallWindowProc(x, 0, 0, 0, 0)
    Erase asmCode
    
End Function

Public Function CallASMCopy(addL As Long, toLocation As Long, ByVal fromLocation As Long) As Long
    Dim tmpStr As String
    Dim hStr As String
    Dim x As Long
    Dim asmCode(46) As Byte
    Dim cnt As Integer
    Dim callLocation As Long
    
    asmCode(0) = &H55 ' Push ebp
    asmCode(1) = &H8B
    asmCode(2) = &HEC ' Mov ebp, esp
    
    asmCode(3) = &H83
    asmCode(4) = &HEC
    asmCode(5) = &H20 ' Sub esp, 20h
    
    asmCode(6) = &H50 ' push eax
    asmCode(7) = &H51 ' push ecx
    asmCode(8) = &H56 ' push esi
    asmCode(9) = &H57 ' push edi
    
    asmCode(10) = &HB8
    asmCode(11) = &H10  '# of bytes to copy
    asmCode(12) = &H0
    asmCode(13) = &H0
    asmCode(14) = &H0 ' mov eax, 00000010
    
    asmCode(15) = &H50 'Push eax

    asmCode(16) = &HB8
    hStr = Hex(fromLocation)
    hStr = String(8 - Len(hStr), "0") & hStr
    tmpStr = ""
    For cnt = 8 To 1 Step -2
        tmpStr = tmpStr + Mid(hStr, cnt - 1, 2)
    Next
    hStr = Left(tmpStr, 2)
    asmCode(17) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 3, 2)
    asmCode(18) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 5, 2)
    asmCode(19) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 7, 2)
    asmCode(20) = CLng("&H" & hStr)
    
    asmCode(21) = &H50 'Push eax
    
    asmCode(22) = &HB8
    hStr = Hex(toLocation)
    hStr = String(8 - Len(hStr), "0") & hStr
    tmpStr = ""
    For cnt = 8 To 1 Step -2
        tmpStr = tmpStr + Mid(hStr, cnt - 1, 2)
    Next
    hStr = Left(tmpStr, 2)
    asmCode(23) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 3, 2)
    asmCode(24) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 5, 2)
    asmCode(25) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 7, 2)
    asmCode(26) = CLng("&H" & hStr)
    
    asmCode(27) = &H50 ' Push eax
    asmCode(28) = &HE8
    'pointer in mem to our machine code
    x = VarPtr(asmCode(0))
    'these next 4 must be calculated from
    'actual address of VxDCall - our machine
    'code location
    
    callLocation = (addL - x) - 33
    'callLocation = Not callLocation
    
    hStr = Hex(callLocation)
    hStr = String(8 - Len(hStr), "0") & hStr
    tmpStr = ""
    For cnt = 8 To 1 Step -2
        tmpStr = tmpStr + Mid(hStr, cnt - 1, 2)
    Next
    
    hStr = Left(tmpStr, 2)
    asmCode(29) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 3, 2)
    asmCode(30) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 5, 2)
    asmCode(31) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 7, 2)
    asmCode(32) = CLng("&H" & hStr)
    
    asmCode(33) = &H58 'pop eax
    asmCode(34) = &H58 'pop eax
    asmCode(35) = &H58 'pop eax
    
    asmCode(36) = &H5F 'pop edi
    asmCode(37) = &H5E 'pop esi
    asmCode(38) = &H59 'pop ecx
    asmCode(39) = &H58 'pop eax
    
    asmCode(40) = &H8B
    asmCode(41) = &HE5 'mov esp, ebp
    asmCode(42) = &H5D 'pop ebp
    asmCode(43) = &HC2
    asmCode(44) = &H10 ' ret 10
    asmCode(45) = &H0
    asmCode(46) = &H0  ' just in case stack filled bith garbage
    
    CallASMCopy = CallWindowProc(x, 0, 0, 0, 0)
    Erase asmCode
    
End Function

