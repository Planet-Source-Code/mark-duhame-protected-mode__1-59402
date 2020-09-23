Attribute VB_Name = "modAsmCalls"
Option Explicit

'
'
'Sub to retrieve CPU information
'
'add = pointer byte array to store CPU info
'
'
Public Sub ASMGetCpu(ByRef add() As Byte, ByVal base As Long, ByVal limit As Integer)
    Dim tmpStr As String
    Dim hStr As String
    Dim x As Long
    Dim asmCode(70) As Byte
    Dim CallCode(31) As Byte
    Dim cnt As Integer
    Dim callLocation As Long
    Dim copyBytes(8) As Byte
    Dim origBytes(8) As Byte
    Dim codeCall As Long
    
    x = base + limit - 8
    hStr = Hex(base)
    hStr = Hex(x)
    hStr = "0" & Right(hStr, 3)
    cnt = limit - 8
    hStr = Hex(cnt)
    If Len(hStr) < 4 Then
        hStr = String(4 - Len(hStr), "0") & hStr
    End If
    
    'callcode will call machine we place
    'at new last entry GDT location
    CallCode(0) = &H55 ' Push ebp
    CallCode(1) = &H8B
    CallCode(2) = &HEC ' Mov ebp, esp
    
    CallCode(3) = &H83
    CallCode(4) = &HEC
    CallCode(5) = &H20 ' Sub esp, 20h
    
    CallCode(6) = &H50 ' push eax
    CallCode(7) = &H51 ' push ecx
    CallCode(8) = &H56 ' push esi
    CallCode(9) = &H57 ' push edi
    CallCode(10) = &H52 ' push edx
    CallCode(11) = &H53 ' push ebx
    CallCode(12) = &H9A ' call last GDT entry
    CallCode(13) = &H0
    CallCode(14) = &H0
    CallCode(15) = &H0
    CallCode(16) = &H0
    CallCode(17) = CLng("&H" & Right(hStr, 2))
    CallCode(18) = CLng("&H" & Left(hStr, 2))
    CallCode(19) = &H5B 'pop ebx
    CallCode(20) = &H5A 'pop edx
    CallCode(21) = &H5F 'pop edi
    CallCode(22) = &H5E 'pop esi
    CallCode(23) = &H59 'pop ecx
    CallCode(24) = &H58 'pop eax
    
    CallCode(25) = &H8B
    CallCode(26) = &HE5 'mov esp, ebp
    CallCode(27) = &H5D 'pop ebp
    CallCode(28) = &HC2
    CallCode(29) = &H10 ' ret 10
    CallCode(30) = &H0
    CallCode(31) = &H0  ' just in case stack filled bith garbage
    
    asmCode(0) = &HFA   ' cli
    asmCode(1) = &HFC   ' cld
    asmCode(2) = &HBF ' mov edi with address we passed in add
    callLocation = VarPtr(add(0))
    hStr = Hex(callLocation)
    hStr = String(8 - Len(hStr), "0") & hStr
    tmpStr = ""
    For cnt = 8 To 1 Step -2
        tmpStr = tmpStr + Mid(hStr, cnt - 1, 2)
    Next
    
    hStr = Left(tmpStr, 2)
    asmCode(3) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 3, 2)
    asmCode(4) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 5, 2)
    asmCode(5) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 7, 2)
    asmCode(6) = CLng("&H" & hStr)
        
    asmCode(7) = &HF   ' mov eax, cr0
    asmCode(8) = &H20
    asmCode(9) = &HC0
    
    asmCode(10) = &HAB  ' stosd
    
    asmCode(11) = &HF   ' mov eax, cr2
    asmCode(12) = &H20
    asmCode(13) = &HD0
    
    asmCode(14) = &HAB  ' stosd
    
    asmCode(15) = &HF   ' mov eax, cr3
    asmCode(16) = &H20
    asmCode(17) = &HD8
    
    asmCode(18) = &HAB  ' stosd
    
    asmCode(19) = &HF   ' mov eax, cr4
    asmCode(20) = &H20
    asmCode(21) = &HE0
    
    asmCode(22) = &HAB  ' stosd
    
    asmCode(23) = &HF   ' mov eax, dr0
    asmCode(24) = &H21
    asmCode(25) = &HC0
    
    asmCode(26) = &HAB  ' stosd
    
    asmCode(27) = &HF   ' mov eax, dr1
    asmCode(28) = &H21
    asmCode(29) = &HC8
    
    asmCode(30) = &HAB  ' stosd
    
    asmCode(31) = &HF   ' mov eax, dr2
    asmCode(32) = &H21
    asmCode(33) = &HD0
    
    asmCode(34) = &HAB  ' stosd
    
    asmCode(35) = &HF   ' mov eax, dr3
    asmCode(36) = &H21
    asmCode(37) = &HD8
    
    asmCode(38) = &HAB  ' stosd
    
    asmCode(39) = &HF   ' mov eax, dr6
    asmCode(40) = &H21
    asmCode(41) = &HF0
    
    asmCode(42) = &HAB  ' stosd
    
    asmCode(43) = &HF   ' mov eax, dr7
    asmCode(44) = &H21
    asmCode(45) = &HF8
    
    asmCode(46) = &HAB  ' stosd
    
    asmCode(47) = &H33  ' xor eax, eax
    asmCode(48) = &HC0
    
    asmCode(49) = &HB3   ' mov bl, 8
    asmCode(50) = &H8
        
    asmCode(51) = &HF   ' mov eax, CPU
    asmCode(52) = &HA2
        
    asmCode(53) = &H3C  ' cmp al, 0
    asmCode(54) = &H0
    
    asmCode(55) = &H75  ' jnz + 2
    asmCode(56) = &H2
    
    asmCode(57) = &HF   ' mov eax, CPU
    asmCode(58) = &HA2
    
    asmCode(59) = &H8B  ' mov eax, ebx
    asmCode(60) = &HC3
    
    asmCode(61) = &HAB  ' stosd
   
    asmCode(62) = &H8B  ' mov eax, edx
    asmCode(63) = &HC2
    
    asmCode(64) = &HAB  ' stosd
    
    asmCode(65) = &H8B  'mov eax, ecx
    asmCode(66) = &HC1
    
    asmCode(67) = &HAB  ' stosd
    
    asmCode(68) = &HFB  ' sti
    asmCode(69) = &HCB  ' retf
    asmCode(70) = &H0  ' just in case stack filled bith garbage

    x = VarPtr(asmCode(0)) - 8
    tmpStr = Hex(x + 8)
    If Len(tmpStr) < 8 Then
        tmpStr = String(8 - Len(tmpStr), "0") & tmpStr
    End If

    hStr = Right(tmpStr, 2) & Mid(tmpStr, 5, 2)
    hStr = hStr & "280000EC" & Mid(tmpStr, 3, 2) & Left(tmpStr, 2)
    
    copyBytes(0) = CLng("&H" & Mid(hStr, 1, 2))
    copyBytes(1) = CLng("&H" & Mid(hStr, 3, 2))
    copyBytes(2) = CLng("&H" & Mid(hStr, 5, 2))
    copyBytes(3) = CLng("&H" & Mid(hStr, 7, 2))
    copyBytes(4) = CLng("&H" & Mid(hStr, 9, 2))
    copyBytes(5) = CLng("&H" & Mid(hStr, 11, 2))
    copyBytes(6) = CLng("&H" & Mid(hStr, 13, 2))
    copyBytes(7) = CLng("&H" & Mid(hStr, 15, 2))

    base = base + limit - 8
    callLocation = VarPtr(origBytes(0))
    'copy GDT to origbytes array
    CopyMemory ByVal callLocation, ByVal base, 8
    callLocation = VarPtr(copyBytes(0))
    'copy our routine to the GDT
    CopyMemory ByVal base, ByVal callLocation, 8
    cnt = limit - 9
    VirtualLock x + 7, 88
    Sleep 0
    codeCall = VarPtr(CallCode(0))
    'pointer in mem to our machine code execute it
    CallWindowProc codeCall, 0, 0, 0, 0
    VirtualUnlock x + 7, 88
    callLocation = VarPtr(origBytes(0))
    'copy original GDT bytes back
    CopyMemory ByVal base, ByVal callLocation, 8
    Erase asmCode
    Erase CallCode
    Erase copyBytes
    Erase origBytes
    
End Sub

'
'
'Funtion to retrieve Vxd List Location
'
'add = pointer byte array to store CPU info
'
'Returns: ret = addrres of ddblist location
'
Public Function ASMGetDDBListLocation(ByRef add() As Byte, ByVal base As Long, ByVal limit As Integer) As Long
    Dim tmpStr As String
    Dim hStr As String
    Dim x As Long
    Dim asmCode(10) As Byte
    Dim CallCode(30) As Byte
    Dim cnt As Integer
    Dim callLocation As Long
    Dim copyBytes(8) As Byte
    Dim origBytes(8) As Byte
    Dim codeCall As Long
    Dim ret As Long
    
    hStr = Hex(base)
    hStr = Hex(x)
    hStr = "0" & Right(hStr, 3)
    cnt = limit - 8
    hStr = Hex(cnt)
    If Len(hStr) < 4 Then
        hStr = String(4 - Len(hStr), "0") & hStr
    End If
    
    'callcode will call machine we place
    'at new last entry GDT location
    CallCode(0) = &H55 ' Push ebp
    CallCode(1) = &H8B
    CallCode(2) = &HEC ' Mov ebp, esp
    
    CallCode(3) = &H83
    CallCode(4) = &HEC
    CallCode(5) = &H20 ' Sub esp, 20h
    
    CallCode(6) = &H90 ' nop
    CallCode(7) = &H51 ' push ecx
    CallCode(8) = &H56 ' push esi
    CallCode(9) = &H57 ' push edi
    CallCode(10) = &H52 ' push edx

    CallCode(11) = &H9A      ' call last GDT entry
    CallCode(12) = &H0
    CallCode(13) = &H0
    CallCode(14) = &H0
    CallCode(15) = &H0
    CallCode(16) = CLng("&H" & Right(hStr, 2))
    CallCode(17) = CLng("&H" & Left(hStr, 2))
    CallCode(18) = &H90
    CallCode(19) = &H5A 'pop edx
    CallCode(20) = &H5F 'pop edi
    CallCode(21) = &H5E 'pop esi
    CallCode(22) = &H59 'pop ecx
    CallCode(23) = &H90 'nop
    
    CallCode(24) = &H8B
    CallCode(25) = &HE5 'mov esp, ebp
    CallCode(26) = &H5D 'pop ebp
    CallCode(27) = &HC2
    CallCode(28) = &H10 ' ret 10
    CallCode(29) = &H0
    CallCode(30) = &H0  ' just in case stack filled bith garbage
    
    asmCode(0) = &HFA   ' cli
    asmCode(1) = &HFC   ' cld
    
    asmCode(2) = &HCD   ' Int 20 to vmm service 013E
    asmCode(3) = &H20   ' which is GetDDBListLocation
    asmCode(4) = &H3E
    asmCode(5) = &H1
    asmCode(6) = &H1
    asmCode(7) = &H0
    
    asmCode(8) = &HFB   ' sti
    asmCode(9) = &HCB   ' retf
    
    asmCode(10) = &H0  ' just in case stack filled bith garbage

    x = VarPtr(asmCode(0)) - 8
    tmpStr = Hex(x + 8)
    If Len(tmpStr) < 8 Then
        tmpStr = String(8 - Len(tmpStr), "0") & tmpStr
    End If

    hStr = Right(tmpStr, 2) & Mid(tmpStr, 5, 2)
    hStr = hStr & "280000EC" & Mid(tmpStr, 3, 2) & Left(tmpStr, 2)
    
    copyBytes(0) = CLng("&H" & Mid(hStr, 1, 2))
    copyBytes(1) = CLng("&H" & Mid(hStr, 3, 2))
    copyBytes(2) = CLng("&H" & Mid(hStr, 5, 2))
    copyBytes(3) = CLng("&H" & Mid(hStr, 7, 2))
    copyBytes(4) = CLng("&H" & Mid(hStr, 9, 2))
    copyBytes(5) = CLng("&H" & Mid(hStr, 11, 2))
    copyBytes(6) = CLng("&H" & Mid(hStr, 13, 2))
    copyBytes(7) = CLng("&H" & Mid(hStr, 15, 2))

    base = base + limit - 8
    callLocation = VarPtr(origBytes(0))
    'copy GDT to origbytes array
    CopyMemory ByVal callLocation, ByVal base, 8
    callLocation = VarPtr(copyBytes(0))
    'copy our routine to the GDT
    CopyMemory ByVal base, ByVal callLocation, 8
    cnt = limit - 9
    VirtualLock x + 7, 88
    Sleep 0
    codeCall = VarPtr(CallCode(0))
    'pointer in mem to our machine code execute it
    ret = CallWindowProc(codeCall, 0, 0, 0, 0)
    VirtualUnlock x + 7, 88
    callLocation = VarPtr(origBytes(0))
    'copy original GDT bytes back
    CopyMemory ByVal base, ByVal callLocation, 8
    
    ASMGetDDBListLocation = ret
    Erase asmCode
    Erase CallCode
    Erase copyBytes
    Erase origBytes
    
End Function
