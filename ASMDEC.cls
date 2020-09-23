VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DisAsm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*****W32 OPCODE DISASSEMBLER WRITTEN BY VANJA FUCKAR EMAIL:INGA@VIP.HR
'PENTIUM 3+PENTIUM 4 OPCODE INCLUDED!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim Term As Boolean
Private VBASE As Long 'VIRTUAL BASE ADDRESS

Public Property Let BaseAddress(ByVal Address As Long)
    
    VBASE = Address
    
End Property

Public Function DisAssemble(Data() As Byte, ByVal Start As Long, ByRef Forward As Byte, Optional ByVal HexAdr As Byte = 1, Optional ByVal BinDump As Byte) As String
    Dim I As Integer
    Dim tmp As Integer
    Dim tmpStr As String
    Dim newStr As String
    Dim XRG As Byte
    Dim ZRG As Byte
    Dim OPT16 As Byte '16-OPCODE
    Dim EXTRAADD As Byte
    Dim EXTRA As Byte
    Dim CL As Byte 'Command Length
    Dim DL As Byte 'Data Length
    Dim ISWP As Byte
    Dim COMMD As String
    Dim X1 As String
    Dim X2 As String
    Dim X3 As String
    Dim ret As Byte
    'FLOAT************
    Dim FWR As Byte
    Dim DB2 As Byte
    Dim DB3 As Byte
    Dim CHK As String
    '*****************
    Dim CX3 As Byte
    Dim MCs As Byte
    
    On Error GoTo EendX:

BEGGIN:

    MCs = Data(Start)
    
    If Start < UBound(Data) Then
        DB2 = Data(Start + 1) 'SEGMENTS 0-3f,40-7f,80-bf,c0-ff
        FWR = Int((DB2 And &H3F) / 8) 'OCTAL SEGMENTS 0-7
    End If
    
    Select Case MCs
    
        'MATH ***********************************************+DOBRO
        'Case Is = &H0, &H8, &H10, &H18, &H20, &H28, &H30, &H38
        'COMMD = MATHOFFSET(Int(MCs / 8))
        'GoTo InTTS2
        
        'MATH ***********************************************+DOBRO
        Case &H0  ', &H8, &H10, &H18, &H20, &H28, &H30, &H38
            If ISWP = 0 Then
                If Check0(Data, Start, DisAssemble, Forward) = 1 Then
                    GoTo OutOfF
                End If
            End If
    
            COMMD = MATHOFFSET(Int(MCs / 8))
            GoTo InTTS2
    
        Case Is = &H8, &H10, &H18, &H20, &H28, &H30, &H38
            COMMD = MATHOFFSET(Int(MCs / 8))
            GoTo InTTS2
        '**********************************************************
    
    
        Case Is = &H1, &H9, &H11, &H19, &H21, &H29, &H31, &H39
            COMMD = MATHOFFSET(Int(MCs / 8))
            GoTo InTTS4
    
        Case Is = &H2, &HA, &H12, &H1A, &H22, &H2A, &H32, &H3A
            COMMD = MATHOFFSET(Int(MCs / 8))
            GoTo InTTSR2
    
        Case Is = &H3, &HB, &H13, &H1B, &H23, &H2B, &H33, &H3B
            COMMD = MATHOFFSET(Int(MCs / 8))
            GoTo InTTSR4
    
        Case Is = &H4, &HC, &H14, &H1C, &H24, &H2C, &H34, &H3C
            CL = 2
            DisAssemble = MATHOFFSET(Int(MCs / 8)) & " AL," & ByteToStr(DB2)
    
        Case Is = &H5, &HD, &H15, &H1D, &H25, &H2D, &H35, &H3D
            If ISWP = 1 Then
                CL = 3
                DisAssemble = MATHOFFSET(Int(MCs / 8)) & " AX," & WordToStr(GetWordFromList(Data, Start + 1))
            Else
                CL = 5
                DisAssemble = MATHOFFSET(Int(MCs / 8)) & " EAX," & DwordToStr(GetDWordFromList(Data, Start + 1))
            End If
    
        Case &H6
            CL = 1
            DisAssemble = "PUSH ES"
    
        Case &H7
            CL = 1
            DisAssemble = "POP ES"
    
        Case &HE
            CL = 1
            DisAssemble = "PUSH CS"
    
        Case &H16
            CL = 1
            DisAssemble = "PUSH SS"
    
        Case &H17
            CL = 1
            DisAssemble = "POP SS"
    
        Case &H1E
            CL = 1
            DisAssemble = "PUSH DS"
    
        Case &H1F
            CL = 1
            DisAssemble = "POP DS"
    
        '26:ES ,2E:CS ,36:SS ,3E:DS ,64:FS ,65:GS ,66:?? (SEGMENTS)
        Case &H26
            Start = Start + 1
            EXTRA = EXTRA + 1
            GoTo BEGGIN
    
        Case &H2E
            Start = Start + 1
            EXTRA = EXTRA + 1
            GoTo BEGGIN
    
        Case &H36
            Start = Start + 1
            EXTRA = EXTRA + 1
            GoTo BEGGIN
    
        Case &H3E
            Start = Start + 1
            EXTRA = EXTRA + 1
            GoTo BEGGIN
    
        Case &H64
            Start = Start + 1
            EXTRA = EXTRA + 1
            GoTo BEGGIN
    
        Case &H65
            Start = Start + 1
            EXTRA = EXTRA + 1
            GoTo BEGGIN
    
        Case &H66
            ISWP = 1 '16 BYTES!
            Start = Start + 1
            EXTRA = EXTRA + 1
            GoTo BEGGIN
    
        Case &H27
            CL = 1
            DisAssemble = "DAA"
    
        Case &H2F
            CL = 1
            DisAssemble = "DAS"
    
        Case &H37
            CL = 1
            DisAssemble = "AAA"
    
        Case &H3F
            CL = 1
            DisAssemble = "AAS"
    
        Case &H40 To &H47
            CL = 1
            If ISWP = 1 Then
                DisAssemble = "INC " & REGOFFSET2(MCs And 7)
            Else
                DisAssemble = "INC " & REGOFFSET4(MCs And 7)
            End If
    
        Case &H48 To &H4F
            CL = 1
            If ISWP = 1 Then
                DisAssemble = "DEC " & REGOFFSET2(MCs And 7)
            Else
                DisAssemble = "DEC " & REGOFFSET4(MCs And 7)
            End If
    
        Case &H50
            DisAssemble = "Push EAX"
            CL = 1
        
        Case &H51
            DisAssemble = "Push ECX"
            CL = 1
        
        Case &H52
            DisAssemble = "Push EDX"
            CL = 1
        
        Case &H53
            DisAssemble = "Push EBX"
            CL = 1
        
        Case &H54
            DisAssemble = "Push ESP"
            CL = 1
            
        Case &H55
            DisAssemble = "Push EBP"
            CL = 1
        
        Case &H56
            DisAssemble = "Push ESI"
            CL = 1
        
        Case &H57
            DisAssemble = "Push EDI"
            CL = 1
    
        Case &H58
            DisAssemble = "Pop EAX"
            CL = 1
        
        Case &H59
            DisAssemble = "Pop ECX"
            CL = 1
        
        Case &H5A
            DisAssemble = "Pop EDX"
            CL = 1
        
        Case &H5B
            DisAssemble = "Pop EBX"
            CL = 1
        
        Case &H5C
            DisAssemble = "Pop ESP"
            CL = 1
        
        Case &H5D
            DisAssemble = "Pop EBP"
            CL = 1
        
        Case &H5E
            DisAssemble = "Pop ESI"
            CL = 1
        
        Case &H5F
            DisAssemble = "Pop EDI"
            CL = 1
    
        Case &H60
            CL = 1
            If ISWP = 1 Then
                DisAssemble = "PUSHA"
            Else
                DisAssemble = "PUSHAD"
            End If
    
        Case &H61
            CL = 1
            If ISWP = 1 Then
                DisAssemble = "POPA"
            Else
                DisAssemble = "POPAD"
            End If
    
        Case &H62 '+
            If DB2 >= &HC0 Then GoTo EendX
            COMMD = "BOUND"
            If ISWP = 1 Then
                ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
                X3 = REGOFFSET2(XRG)
                RJoin DPR, DisAssemble, COMMD, X1, X2, X3, ret
            Else
                ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
                X3 = REGOFFSET4(XRG)
                RJoin QPR, DisAssemble, COMMD, X1, X2, X3, ret
            End If
    
        Case &H63
            COMMD = "ARPL"
            ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
            X3 = REGOFFSET2(XRG)
            LJoin WPL, DisAssemble, COMMD, X1, X2, X3, ret
    
        '********RESERVED!!!!!!!!!!!!!!!16 BYTES!
        Case &H67
            Start = Start + 1
            EXTRA = EXTRA + 1
            OPT16 = 1
            GoTo BEGGIN
    
        Case &H68 '+
            If ISWP = 1 Then
                CL = 3
                DisAssemble = "PUSH " & WordToStr(GetWordFromList(Data, Start + 1))
            Else
                CL = 5
                DisAssemble = "PUSH " & DwordToStr(GetDWordFromList(Data, Start + 1))
            End If
    
        Case &H69 '+
            COMMD = "IMUL"
            If ISWP = 1 Then
                ret = PerformCMD3(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16, ISWP)
                TJoin WPR, DisAssemble, COMMD, REGOFFSET2(FWR), X1, X2, X3, ret
            Else
                ret = PerformCMD3(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16, ISWP)
                TJoin DPR, DisAssemble, COMMD, REGOFFSET4(FWR), X1, X2, X3, ret
            End If
    
        Case &H6A '+
            CL = 2
            If DB2 > &H7F Then
                If ISWP = 1 Then
                    DisAssemble = "PUSH " & WordToStr(DB2 Or &HFF00)
                Else
                    DisAssemble = "PUSH " & DwordToStr(DB2 Or &HFFFFFF00)
                End If
            Else
                DisAssemble = "PUSH " & Hex(DB2)
            End If
    
        Case &H6B '+
            COMMD = "IMUL"
            If ISWP = 1 Then
                ret = PerformCMD2(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, ZRG, OPT16)
                'CX3 = CByte(Val("&H" & X3))
                CX3 = CByte("&H" & X3)
                If CX3 > &H7F Then
                    X3 = WordToStr(CX3 Or &HFF00)
                End If
                TJoin WPR, DisAssemble, COMMD, REGOFFSET2(FWR), X1, X2, X3, ret
            Else
                ret = PerformCMD2(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, ZRG, OPT16)
                'CX3 = CByte(Val("&H" & X3))
                CX3 = CByte("&H" & X3)
                If CX3 > &H7F Then
                    X3 = DwordToStr(CX3 Or &HFFFFFF00)
                End If
                TJoin DPR, DisAssemble, COMMD, REGOFFSET4(FWR), X1, X2, X3, ret
            End If
    
        Case &H6C
            DisAssemble = "INS BYTE PTR [EDI],DX"
            CL = 1
    
        Case &H6D '+
            If ISWP = 1 Then
                DisAssemble = "INS WORD PTR [EDI],DX"
            Else
                DisAssemble = "INS DWORD PTR [EDI],DX"
            End If
            CL = 1
    
        Case &H6E
            DisAssemble = "OUTS DX,BYTE PTR [ESI]"
            CL = 1
    
        Case &H6F '+
            If ISWP = 1 Then
                DisAssemble = "OUTS DX,WORD PTR [ESI]"
            Else
                DisAssemble = "OUTS DX,DWORD PTR [ESI]"
            End If
            CL = 1
    
        Case &H70 To &H7F
            DisAssemble = CalcShortJump(DB2, JXX(MCs And &HF), Start + 1, VBASE + Start)
            CL = 2
    
        Case &HF
            '************SC
            DB3 = Data(Start + 2)
            FWR = Int((DB3 And &H3F) / 8)
            Select Case DB2
    
                Case 0
                    If FWR = 7 Then GoTo EendX
                    ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
                    COMMD = MMI(XRG)
                    MJoin WPL, DisAssemble, COMMD, X1, X2, ret
    
                Case 1
                    If DB3 >= &HC0 Then
                        If Not (FWR = 4) And Not (FWR > 5) Then
                            GoTo EendX
                        ElseIf FWR = 7 Then
                            DisAssemble = "SWAPGS"
                            CL = 2
                        Else
                            ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
                            COMMD = MMI2(XRG)
                            MJoin NOL, DisAssemble, COMMD, X1, X2, ret
                        End If
                    Else
                        If FWR = 5 Then GoTo EendX
                            ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
                            COMMD = MMI2(XRG)
                            Select Case FWR
                                Case 0 To 3
                                    MJoin FPL, DisAssemble, COMMD, X1, X2, ret
                                Case Is = 4, 6
                                    MJoin WPL, DisAssemble, COMMD, X1, X2, ret
                                Case Else
                                    MJoin NOL, DisAssemble, COMMD, X1, X2, ret
                            End Select
                        End If
    
                Case 2
                    COMMD = "LAR"
                    GoTo InIML
    
                Case 3
                    COMMD = "LSL"
                    GoTo InIML
    
                Case 5
                    DisAssemble = "SYSCALL"
                    CL = 1
    
                Case 6
                    DisAssemble = "CLTS"
                    CL = 1
    
                Case 7
                    DisAssemble = "SYSRET"
                    CL = 1
    
                Case 8
                    DisAssemble = "INVD"
                    CL = 1
    
                Case 9
                    DisAssemble = "WBINVD"
                    CL = 1
    
                Case &HB
                    DisAssemble = "UD2"
                    CL = 1
    
                Case &HD 'PREFETCH
                    If DB3 >= &HC0 Then GoTo EendX
                    If FWR = 0 Then
                        COMMD = "PREFETCH"
                    ElseIf FWR = 1 Then
                        COMMD = "PREFETCHW"
                    Else
                        GoTo EendX
                    End If
                    ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
                    X3 = REGOFFSET4(XRG)
                    MJoin NOL, DisAssemble, COMMD, X1, X2, ret
    
                Case &HE
                    DisAssemble = "FEMMS"
                    CL = 1
    
                Case &H6C
                    If ISWP = 1 Then
                        COMMD = "PUNPCKLQDQ": GoTo InXMMX
                    Else
                        GoTo EendX
                    End If
    
                Case &H6D
                    If ISWP = 1 Then
                        COMMD = "PUNPCKHQDQ": GoTo InXMMX
                    Else
                        GoTo EendX
                    End If
    
                Case &H12
                    If DB3 >= &HC0 Then
                        COMMD = "MOVHLPS"
                    Else
                        If ISWP = 1 Then
                            COMMD = "MOVLPD"
                        Else
                            COMMD = "MOVLPS"
                        End If
                    End If
                    ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
                    X3 = XMM(XRG)
                    RJoin QPR, DisAssemble, COMMD, X1, X2, X3, ret
    
                Case &H13
                    If DB3 >= &HC0 Then GoTo EendX
                    If ISWP = 1 Then
                        COMMD = "MOVLPD"
                    Else
                        COMMD = "MOVLPS"
                    End If
                    ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
                    X3 = XMM(XRG)
                    LJoin QPL, DisAssemble, COMMD, X1, X2, X3, ret
    
                Case &H14
                    If ISWP = 1 Then
                        COMMD = "UNPCKLPD"
                    Else
                        COMMD = "UNPCKLPS"
                    End If
                    ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
                    X3 = XMM(XRG)
                    RJoin XMMPR, DisAssemble, COMMD, X1, X2, X3, ret
    
    
                Case &H15
                    If ISWP = 1 Then
                        COMMD = "UNPCKHPD"
                    Else
                        COMMD = "UNPCKHPS"
                    End If
                    ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
                    X3 = XMM(XRG)
                    RJoin XMMPR, DisAssemble, COMMD, X1, X2, X3, ret
    
                Case &H16
                    If ISWP = 1 Then
                        COMMD = "MOVHPD"
                    Else
                        COMMD = "MOVHPS"
                    End If
                    ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
                    X3 = XMM(XRG)
                    RJoin QPR, DisAssemble, COMMD, X1, X2, X3, ret
    
                Case &H17
                    If DB3 >= &HC0 Then GoTo EendX
                    If ISWP = 1 Then
                        COMMD = "MOVHPD"
                    Else
                        COMMD = "MOVHPS"
                    End If
                    ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
                    X3 = XMM(XRG)
                    LJoin QPL, DisAssemble, COMMD, X1, X2, X3, ret
    
                Case &H18 'PREFETCH A-2
                    If DB3 >= &HC0 Then GoTo EendX
                    If FWR = 0 Then
                        COMMD = "PREFETCHNTA"
                        GoTo In18
                    ElseIf FWR = 1 Then
                        COMMD = "PREFETCHT0"
                        GoTo In18
                    ElseIf FWR = 2 Then
                        COMMD = "PREFETCHT1"
                        GoTo In18
                    ElseIf FWR = 3 Then
                        COMMD = "PREFETCHT2"
In18:
    ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
    X3 = REGOFFSET4(XRG)
    MJoin NOL, DisAssemble, COMMD, X1, X2, ret
                    Else
                        GoTo EendX
                    End If
    
                Case &H28
                    If ISWP = 1 Then
                        COMMD = "MOVAPD"
                    Else
                        COMMD = "MOVAPS"
                    End If
                    ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
                    X3 = XMM(XRG)
                    RJoin XMMPR, DisAssemble, COMMD, X1, X2, X3, ret
    
                Case &H29
                    If ISWP = 1 Then
                        COMMD = "MOVAPD"
                    Else
                        COMMD = "MOVAPS"
                    End If
                    ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
                    X3 = XMM(XRG)
                    LJoin XMMPL, DisAssemble, COMMD, X1, X2, X3, ret
    
                Case &H2A
                    If ISWP = 1 Then
                        COMMD = "CVTPI2PD"
                    Else
                        COMMD = "CVTPI2PS"
                    End If
                    ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, MMX, XRG, OPT16)
                    X3 = XMM(XRG)
                    RJoin MMPR, DisAssemble, COMMD, X1, X2, X3, ret
    
                Case &H2B
                    If DB3 >= &HC0 Then GoTo EendX
                    If ISWP = 1 Then
                        COMMD = "MOVNTPD"
                    Else
                        COMMD = "MOVNTPS"
                    End If
                    ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
                    X3 = XMM(XRG)
                    LJoin XMMPL, DisAssemble, COMMD, X1, X2, X3, ret
    
                Case &H2C
                    If ISWP = 1 Then
                        COMMD = "CVTTPD2PI"
                    Else
                        COMMD = "CVTTPS2PI"
                    End If
                    GoTo In2D
    
                Case &H2E
                    If ISWP = 1 Then
                        COMMD = "UCOMISD"
                    Else
                        COMMD = "UCOMISS"
                    End If
                    GoTo In2F
    
                Case &H2F
                    If ISWP = 1 Then
                        COMMD = "COMISD"
                    Else
                        COMMD = "COMISS"
                    End If
In2F:
    If ISWP = 1 Then
        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
        X3 = XMM(XRG)
        RJoin MMPR, DisAssemble, COMMD, X1, X2, X3, ret
    Else
        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
        X3 = XMM(XRG)
        RJoin DPR, DisAssemble, COMMD, X1, X2, X3, ret
    End If
    
                Case &H30
                    DisAssemble = "WRMSR"
                    CL = 1
    
                Case &H31
                    DisAssemble = "RDTSC"
                    CL = 1
    
                Case &H32
                    DisAssemble = "RDMSR"
                    CL = 1
    
                Case &H33
                    DisAssemble = "RDPMC"
                    CL = 1
    
                Case &H34
                    DisAssemble = "SYSENTER"
                    CL = 1
    
                Case &H35
                    DisAssemble = "SYSEXIT"
                    CL = 1
    
                Case &H40
                    COMMD = "CMOVO"
                    GoTo InIML
    
                Case &H41
                    COMMD = "CMOVNO"
                    GoTo InIML
    
                Case &H42
                    COMMD = "CMOVB"
                    GoTo InIML
    
                Case &H43
                    COMMD = "CMOVAE"
                    GoTo InIML
    
                Case &H44
                    COMMD = "CMOVE"
                    GoTo InIML
    
                Case &H45
                    COMMD = "CMOVNE"
                    GoTo InIML
    
                Case &H46
                    COMMD = "CMOVBE"
                    GoTo InIML
    
                Case &H47
                    COMMD = "CMOVA"
                    GoTo InIML
    
                Case &H48
                    COMMD = "CMOVS"
                    GoTo InIML
    
                Case &H49
                    COMMD = "CMOVNS"
                    GoTo InIML
    
                Case &H4A
                    COMMD = "CMOVP"
                    GoTo InIML
    
                Case &H4B
                    COMMD = "CMOVNP"
                    GoTo InIML
    
                Case &H4C
                    COMMD = "CMOVL"
                    GoTo InIML
    
                Case &H4D
                    COMMD = "CMOVGE"
                    GoTo InIML
    
                Case &H4E
                    COMMD = "CMOVLE"
                    GoTo InIML
    
                Case &H4F
                    COMMD = "CMOVG"
                    GoTo InIML
    
                Case &H51
                    If ISWP = 1 Then
                        COMMD = "SQRTPD"
                    Else
                        COMMD = "SQRTPS"
                    End If
                    GoTo InSQ
    
                Case &H52
                    If ISWP = 1 Then GoTo EendX
                    COMMD = "RSQRTPS"
                    GoTo InSQ
    
                Case &H53
                    If ISWP = 1 Then GoTo EendX
                    COMMD = "RCPPS"
                    GoTo InSQ
    
                Case &H54
                    If ISWP = 1 Then
                        COMMD = "ANDPD"
                    Else
                        COMMD = "ANDPS"
                    End If
                    GoTo InSQ
    
                Case &H55
                    If ISWP = 1 Then
                        COMMD = "ANDNPD"
                    Else
                        COMMD = "ANDNPS"
                    End If
                    GoTo InSQ
    
                Case &H56
                    If ISWP = 1 Then
                        COMMD = "ORPD"
                    Else
                        COMMD = "ORPS"
                    End If
                    GoTo InSQ
    
                Case &H57
                    If ISWP = 1 Then
                        COMMD = "XORPD"
                    Else
                        COMMD = "XORPS"
                    End If
                    GoTo InSQ
    
                Case &H58
                If ISWP = 1 Then
                    COMMD = "ADDPD"
                Else
                    COMMD = "ADDPS"
                End If
                GoTo InSQ
    
                Case &H59
                    If ISWP = 1 Then
                        COMMD = "MULPD"
                    Else
                        COMMD = "MULPS"
                    End If
                    GoTo InSQ
    
                Case &H5A
                    If ISWP = 1 Then
                        COMMD = "CVTPD2PS"
                    Else
                        COMMD = "CVTPS2PD"
                    End If
                    GoTo InSQ
    
                Case &H5B
                    If ISWP = 1 Then
                        COMMD = "CVTPS2DQ"
                    Else
                        COMMD = "CVTDQ2PD"
                    End If
                    GoTo InSQ
    
                Case &H5C
                    If ISWP = 1 Then
                        COMMD = "SUBPD"
                    Else
                        COMMD = "SUBPS"
                    End If
                    GoTo InSQ
    
                Case &H5D
                    If ISWP = 1 Then
                        COMMD = "MINPD"
                    Else
                        COMMD = "MINPS"
                    End If
                    GoTo InSQ
    
                Case &H5E
                    If ISWP = 1 Then
                        COMMD = "DIVPD"
                    Else
                        COMMD = "DIVPS"
                    End If
                    GoTo InSQ
    
                Case &H5F
                    If ISWP = 1 Then
                        COMMD = "MAXPD"
                    Else
                        COMMD = "MAXPS"
                    End If
    
InSQ:
    ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
    X3 = XMM(XRG)
    RJoin XMMPR, DisAssemble, COMMD, X1, X2, X3, ret
    
                Case &H60
                    COMMD = "PUNPCKLBW"
                    GoTo InPU
    
                Case &H61
                    COMMD = "PUNPCKLWD"
                    GoTo InPU
    
                Case &H62
                    COMMD = "PUNPCKLDQ"
InPU:
    If ISWP = 1 Then
        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
        X3 = XMM(XRG)
        RJoin DPR, DisAssemble, COMMD, X1, X2, X3, ret
    Else
        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, MMX, XRG, OPT16)
        X3 = MMX(XRG)
        RJoin DPR, DisAssemble, COMMD, X1, X2, X3, ret
    End If
    
                Case &H63
                    COMMD = "PACKSSWB"
                    GoTo InPA
    
                Case &H64
                    COMMD = "PCMPGTB"
                    GoTo InPA
    
                Case &H65
                    COMMD = "PCMPGTW"
                    GoTo InPA
    
                Case &H66
                    COMMD = "PCMPGTD"
                    GoTo InPA
    
                Case &H67
                    COMMD = "PACKUSWB"
                    GoTo InPA
    
                Case &H68
                    COMMD = "PUNPCKHBW"
                    GoTo InPA
    
                Case &H69
                    COMMD = "PUNPCKHWD"
                    GoTo InPA
    
                Case &H6A
                    COMMD = "PUNPCKHDQ"
                    GoTo InPA
    
                Case &H6B
                    COMMD = "PACKSSDW"
                    GoTo InPA
    
                Case &H6E
                    COMMD = "MOVD"
                    If ISWP = 1 Then
                        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
                        X3 = XMM(XRG)
                        RJoin DPR, DisAssemble, COMMD, X1, X2, X3, ret
                    Else
InMD:
    ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
    X3 = MMX(XRG)
    RJoin DPR, DisAssemble, COMMD, X1, X2, X3, ret
                    End If
    
                Case &H6F
                    If ISWP = 1 Then
                        COMMD = "MOVDQA"
                    Else
                        COMMD = "MOVQ"
                    End If
    
InPA:
    If ISWP = 1 Then
        GoTo InXMMX
    Else
        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, MMX, XRG, OPT16)
        X3 = MMX(XRG)
        RJoin MMPR, DisAssemble, COMMD, X1, X2, X3, ret
    End If
    
    
                Case &H70 '***BYTE OPCODE
                    If ISWP = 1 Then
                        COMMD = "PSHUFD"
                        ret = PerformCMD2(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, ZRG, OPT16)
                        TJoin XMMPR, DisAssemble, COMMD, XMM(FWR), X1, X2, X3, ret
                    Else
                        COMMD = "PSHUFW"
                        ret = PerformCMD2(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, MMX, XRG, ZRG, OPT16)
                        TJoin MMPR, DisAssemble, COMMD, MMX(FWR), X1, X2, X3, ret
                    End If
    
                Case &H74
                    COMMD = "PCMPEQB"
                    If ISWP = 1 Then
                        GoTo InXMMX
                    Else
                        GoTo InPE
                    End If
    
                Case &H75
                    COMMD = "PCMPEQW"
                    If ISWP = 1 Then
                        GoTo InXMMX
                    Else
                        GoTo InPE
                    End If
    
                Case &H76
                    COMMD = "PCMPEQD"
                    If ISWP = 1 Then GoTo InXMMX
InPE:
    ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, MMX, XRG, OPT16)
    X3 = MMX(XRG)
    RJoin MMPR, DisAssemble, COMMD, X1, X2, X3, ret
    
                Case &H77
                    DisAssemble = "EMMS"
                    CL = 1
    
                Case &H7E
                    COMMD = "MOVD"
                    If ISWP = 1 Then
                        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
                        X3 = XMM(XRG)
                        LJoin DPL, DisAssemble, COMMD, X1, X2, X3, ret
                    Else
                        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
                        X3 = MMX(XRG)
                        LJoin DPL, DisAssemble, COMMD, X1, X2, X3, ret
                    End If
    
                Case &H7F
                    If ISWP = 1 Then
                        COMMD = "MOVDQA"
                        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
                        X3 = XMM(XRG)
                        LJoin XMMPL, DisAssemble, COMMD, X1, X2, X3, ret
                    Else
                        COMMD = "MOVQ"
                        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, MMX, XRG, OPT16)
                        X3 = MMX(XRG)
                        LJoin MMPL, DisAssemble, COMMD, X1, X2, X3, ret
                    End If
    
                Case &H80 To &H8F
                    DisAssemble = CalcLongJump(Data, JXX(DB2 And &HF), Start + 1, VBASE + Start)
                    CL = 5
    
                Case &H90 To &H9F
                    COMMD = SETXX(DB2 And &HF)
                    ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, OPT16)
                    MJoin BPL, DisAssemble, COMMD, X1, X2, ret
    
                Case &HA0
                    DisAssemble = "PUSH FS"
                    CL = 1
    
                Case &HA1
                    DisAssemble = "POP FS"
                    CL = 1
    
                Case &HA2
                    DisAssemble = "CPUID"
                    CL = 1
    
                Case &HA3
                    COMMD = "BT"
                    If ISWP = 1 Then
                        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
                        X3 = REGOFFSET2(XRG)
                        LJoin WPL, DisAssemble, COMMD, X1, X2, X3, ret
                    Else
                        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
                        X3 = REGOFFSET4(XRG)
                        LJoin DPL, DisAssemble, COMMD, X1, X2, X3, ret
                    End If
    
                Case &HA4
                    COMMD = "SHLD"
                    GoTo InSHD
    
                Case &HA5
                    COMMD = "SHLD"
                    GoTo InSHD2
    
                Case &HA6
                    COMMD = "XBTS"
                    GoTo InIML
    
                Case &HA7
                    COMMD = "IBTS"
                    GoTo InBTS
    
                Case &HA8
                    DisAssemble = "PUSH GS"
                    CL = 1
    
                Case &HA9
                    DisAssemble = "POP GS"
                    CL = 1
    
                Case &HAA
                    DisAssemble = "RSM"
                    CL = 1
    
                Case &HAB
                    COMMD = "BTS"
                    GoTo InBTS
    
                Case &HAC
                    COMMD = "SHRD"
InSHD:
    If ISWP = 1 Then
        ret = PerformCMD2(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, ZRG, OPT16)
        TJoin2 WPL, DisAssemble, COMMD, REGOFFSET2(FWR), X1, X2, X3, ret
    Else
        ret = PerformCMD2(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, ZRG, OPT16)
        TJoin2 DPL, DisAssemble, COMMD, REGOFFSET4(FWR), X1, X2, X3, ret
    End If
    
                Case &HAD
                    COMMD = "SHRD"
InSHD2:
    X3 = "CL"
    If ISWP = 1 Then
        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
        TJoin2 WPL, DisAssemble, COMMD, REGOFFSET2(FWR), X1, X2, X3, ret
    Else
        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
        TJoin2 DPL, DisAssemble, COMMD, REGOFFSET4(FWR), X1, X2, X3, ret
    End If
    
                Case &HAE
                    If DB3 >= &HC0 Then
                        Select Case FWR
                            Case 5
                                DisAssemble = "LFENCE"
                                CL = 2
                            Case 6
                                DisAssemble = "MFENCE"
                                CL = 2
                            Case 7
                                DisAssemble = "SFENCE"
                                CL = 2
                            Case Else
                                GoTo EendX
                        End Select
                    Else
                        Select Case FWR
                            Case Is = 4, 5, 6
                                GoTo EendX
                            Case Else
                                ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
                                COMMD = CFLX(XRG)
                                MJoin NOL, DisAssemble, COMMD, X1, X2, ret
                        End Select
                    End If
    
                Case &HAF
                    COMMD = "IMUL"
                    GoTo InIML
    
                Case &HB0
                    COMMD = "CMPXCHG"
                    GoTo InLA
    
                Case &HB1
                    COMMD = "CMPXCHG"
                    GoTo InBTS
    
                Case &HB2
                    If DB3 >= &HC0 Then GoTo EendX
                    COMMD = "LSS"
                    GoTo InLS
    
                Case &HB3
                    COMMD = "BTR"
                    GoTo InBTS
    
                Case &HB4
                    If DB3 >= &HC0 Then GoTo EendX
                    COMMD = "LFS"
                    GoTo InLS
    
                Case &HB5
                    If DB3 >= &HC0 Then GoTo EendX
                    COMMD = "LGS"
InLS:
    If ISWP = 1 Then
        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
        X3 = REGOFFSET2(XRG)
        RJoin DPR, DisAssemble, COMMD, X1, X2, X3, ret
    Else
        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
        X3 = REGOFFSET4(XRG)
        RJoin FPR, DisAssemble, COMMD, X1, X2, X3, ret
    End If
    
                Case &HB6
                    COMMD = "MOVZX"
                    GoTo InBSX
    
                Case &HB7
                    COMMD = "MOVZX"
                    GoTo InWSX
    
                Case &HB8
                    COMMD = "JMPE"
                    DisAssemble = CalcLongJump(Data, COMMD, Start + 1, VBASE + Start)
                    CL = 5
    
                Case &HBB
                    COMMD = "BTC"
                    GoTo InBTS
    
                Case &HBC
                    COMMD = "BSF"
                    GoTo InIML
    
                Case &HBD
                    COMMD = "BSR"
InIML:
    If ISWP = 1 Then
        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
        X3 = REGOFFSET2(XRG)
        RJoin WPR, DisAssemble, COMMD, X1, X2, X3, ret
    Else
        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
        X3 = REGOFFSET4(XRG)
        RJoin DPR, DisAssemble, COMMD, X1, X2, X3, ret
    End If
    
                Case &HBE
                    COMMD = "MOVSX"
InBSX:
    If ISWP = 1 Then
        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
        X3 = REGOFFSET2(XRG)
        RJoin BPR, DisAssemble, COMMD, X1, X2, X3, ret
    Else
        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, OPT16)
        X3 = REGOFFSET4(XRG)
        RJoin BPR, DisAssemble, COMMD, X1, X2, X3, ret
    End If
    
                Case &HBF
                    COMMD = "MOVSX"
InWSX:
    ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
    X3 = REGOFFSET4(XRG)
    RJoin WPR, DisAssemble, COMMD, X1, X2, X3, ret
    
                Case &HC0
                    COMMD = "XADD"
InLA:
    ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, OPT16)
    X3 = REGOFFSET1(XRG)
    LJoin BPL, DisAssemble, COMMD, X1, X2, X3, ret
    
                Case &HC1
                    COMMD = "XADD"
                    GoTo InBTS
    
                Case &HC2 'CMPS!
                    ret = PerformCMD2(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, ZRG, OPT16)
                    If (DB3 Mod &H8&) = 4 Then
                        If Data(Start + 3) <= 7 Then
                            COMMD = CMPS(Data(Start + 3))
                        Else
                            COMMD = "CMPSS"
                        End If
                    ElseIf DB3 <= 7 Then
                        COMMD = CMPS(DB3)
                    Else
                        COMMD = "CMPSS"
                    End If
                    TJoin XMMPR, DisAssemble, COMMD, XMM(FWR), X1, X2, X3, ret
    
                Case &HC3
                    If DB3 >= &HC0 Then GoTo EendX
                    COMMD = "MOVNTI"
InBTS:
    If ISWP = 1 Then
        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
        X3 = REGOFFSET2(XRG)
        LJoin WPL, DisAssemble, COMMD, X1, X2, X3, ret
    Else
        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
        X3 = REGOFFSET4(XRG)
        LJoin DPL, DisAssemble, COMMD, X1, X2, X3, ret
    End If
    
                Case &HC4
                    COMMD = "PINSRW"
                    If ISWP = 1 Then
                        ret = PerformCMD2(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, ZRG, OPT16)
                        TJoin WPR, DisAssemble, COMMD, XMM(FWR), X1, X2, X3, ret
                    Else
                        ret = PerformCMD2(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, ZRG, OPT16)
                        TJoin WPR, DisAssemble, COMMD, MMX(FWR), X1, X2, X3, ret
                    End If
    
                Case &HC5
                    COMMD = "PEXTRW"
                    If ISWP = 1 Then
                        ret = PerformCMD2(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, ZRG, OPT16)
                        TJoin XMMPR, DisAssemble, COMMD, REGOFFSET4(FWR), X1, X2, X3, ret
                    Else
                        ret = PerformCMD2(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, MMX, XRG, ZRG, OPT16)
                        TJoin MMPR, DisAssemble, COMMD, REGOFFSET4(FWR), X1, X2, X3, ret
                    End If
    
                Case &HC6
                    If ISWP = 1 Then
                        COMMD = "SHUFPD"
                    Else
                        COMMD = "SHUFPS"
                    End If
                    ret = PerformCMD2(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, ZRG, OPT16)
                    TJoin XMMPR, DisAssemble, COMMD, XMM(FWR), X1, X2, X3, ret
    
                Case &HC7
                    If (DB3 >= 8 And DB3 <= 15) Or (DB3 >= &H48 And DB3 <= &H4F) Or (DB3 >= &H88 And DB3 <= &H8F) Then
                        COMMD = "CMPXCHG8B"
                        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
                        MJoin QPL, DisAssemble, COMMD, X1, X2, ret
                    Else
                        GoTo EendX
                    End If
    
                Case &HC8 To &HCF
                    DisAssemble = "BSWAP " & REGOFFSET4(DB2 And 7)
                    CL = 1
    
                Case &HD1
                    COMMD = "PSRLW"
                    GoTo InPSR
    
                Case &HD2
                    COMMD = "PSRLD"
                    GoTo InPSR
    
                Case &HD3
                    COMMD = "PSRLQ"
                    GoTo InPSR
    
                Case &HD4
                    COMMD = "PADDQ"
                    GoTo InPSR
    
                Case &HD5
                    COMMD = "PMULLW"
                    GoTo InPSR
    
                Case &HD7
                    COMMD = "PMOVMSKB"
                    If ISWP = 1 Then
                        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
                        X3 = REGOFFSET4(XRG)
                        RJoin XMMPR, DisAssemble, COMMD, X1, X2, X3, ret
                    Else
                        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, MMX, XRG, OPT16)
                        X3 = REGOFFSET4(XRG)
                        RJoin MMPR, DisAssemble, COMMD, X1, X2, X3, ret
                    End If
    
                Case &HD8
                    COMMD = "PSUBUSB"
                    GoTo InPSR
    
                Case &HD9
                    COMMD = "PSUBUSW"
                    GoTo InPSR
    
                Case &HDA
                    COMMD = "PMINUB"
                    GoTo InPSR
    
                Case &HDB
                    COMMD = "PAND"
                    GoTo InPSR
    
                Case &HDC
                    COMMD = "PADDUSB"
                    GoTo InPSR
    
                Case &HDD
                    COMMD = "PADDUSW"
                    GoTo InPSR
    
                Case &HDE
                    COMMD = "PMAXUB"
                    GoTo InPSR
    
                Case &HDF
                    COMMD = "PANDN"
                    GoTo InPSR
    
                Case &HE0
                    COMMD = "PAVGB"
                    GoTo InPSR
    
                Case &HE1
                    COMMD = "PSRAW"
                    GoTo InPSR
    
                Case &HE2
                    COMMD = "PSRAD"
                    GoTo InPSR
    
                Case &HE3
                    COMMD = "PAVGW"
                    GoTo InPSR
    
                Case &HE4
                    COMMD = "PMULHUW"
                    GoTo InPSR
    
                Case &HE5
                    COMMD = "PMULHW"
                    GoTo InPSR
    
                Case &HE7
                    If DB3 >= &HC0 Then GoTo EendX
                    If ISWP = 1 Then
                        COMMD = "MOVNTDQ"
                        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
                        X3 = XMM(XRG)
                        LJoin XMMPL, DisAssemble, COMMD, X1, X2, X3, ret
                    Else
                        COMMD = "MOVNTQ"
                        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, MMX, XRG, OPT16)
                        X3 = MMX(XRG)
                        LJoin MMPL, DisAssemble, COMMD, X1, X2, X3, ret
                    End If
    
                Case &HE8
                    COMMD = "PSUBSB"
                    GoTo InPSR
    
                Case &HE9
                    COMMD = "PSUBSW"
                    GoTo InPSR
    
                Case &HEA
                    COMMD = "PMINSW"
                    GoTo InPSR
    
                Case &HEB
                    COMMD = "POR"
                    GoTo InPSR
    
                Case &HEC
                    COMMD = "PADDSB"
                    GoTo InPSR
    
                Case &HED
                    COMMD = "PADDSW"
                    GoTo InPSR
    
                Case &HEE
                    COMMD = "PMAXSW"
                    GoTo InPSR
    
                Case &HEF
                    COMMD = "PXOR"
                    GoTo InPSR
    
                Case &HF1
                    COMMD = "PSLLW"
                    GoTo InPSR
    
                Case &HF2
                    COMMD = "PSLLD"
                    GoTo InPSR
    
                Case &HF3
                    COMMD = "PSLLQ"
                    GoTo InPSR
    
                Case &HF4
                    COMMD = "PMULUDQ"
                    GoTo InPSR
    
                Case &HF5
                    COMMD = "PMADDWD"
                    GoTo InPSR
    
                Case &HF6
                    COMMD = "PSADBW"
                    GoTo InPSR
    
                Case &HF8
                    COMMD = "PSUBB"
                    GoTo InPSR
    
                Case &HF9
                    COMMD = "PSUBW"
                    GoTo InPSR
    
                Case &HFA
                    COMMD = "PSUBD"
                    GoTo InPSR
    
                Case &HFB
                    COMMD = "PSUBQ"
                    GoTo InPSR
    
                Case &HFC
                    COMMD = "PADDB"
                    GoTo InPSR
    
                Case &HFD
                    COMMD = "PADDW"
                    GoTo InPSR
    
                Case &HFE
                    COMMD = "PADDD"
InPSR:
    If ISWP = 1 Then
        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
        X3 = XMM(XRG)
        RJoin XMMPR, DisAssemble, COMMD, X1, X2, X3, ret
    Else
        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, MMX, XRG, OPT16)
        X3 = MMX(XRG)
        RJoin MMPR, DisAssemble, COMMD, X1, X2, X3, ret
    End If
    
                Case &H10
                    If ISWP = 1 Then
                        COMMD = "MOVUPD"
                    Else
                        COMMD = "MOVUPS"
                    End If
InXMMX:
    ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
    X3 = XMM(XRG)
    RJoin XMMPR, DisAssemble, COMMD, X1, X2, X3, ret
    
                Case &H11
                    If ISWP = 1 Then
                        COMMD = "MOVUPD"
                    Else
                        COMMD = "MOVUPS"
                    End If
InXMMX2:
    ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
    X3 = XMM(XRG)
    LJoin XMMPR, DisAssemble, COMMD, X1, X2, X3, ret
    
                Case &H2D
                If ISWP = 1 Then
                    COMMD = "CVTPD2PI"
                Else
                    COMMD = "CVTPS2PI"
                End If
In2D:
    If ISWP = 1 Then
        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
        X3 = MMX(XRG)
        RJoin XMMPR, DisAssemble, COMMD, X1, X2, X3, ret
    Else
        ret = PerformCMD(Data, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
        X3 = MMX(XRG)
        RJoin MMPR, DisAssemble, COMMD, X1, X2, X3, ret
    End If
    
                Case Else
                    GoTo EendX
    
            End Select
            CL = CL + 1
    
    
        Case &H88
            COMMD = "MOV"
            GoTo InTTS2
    
        Case &H89
            COMMD = "MOV"
            GoTo InTTS4
    
        Case &H8A
            COMMD = "MOV"
            GoTo InTTSR2
    
        Case &H8B
            COMMD = "MOV"
            GoTo InTTSR4
    
        Case &H80
            ret = PerformCMD2(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, ZRG, OPT16)
            COMMD = MATHOFFSET(ZRG)
            LJoin BPL, DisAssemble, COMMD, X1, X2, X3, ret
    
        Case &H81 '+
            If ISWP = 1 Then
                ret = PerformCMD3(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16, ISWP)
                COMMD = MATHOFFSET(XRG)
                LJoin WPL, DisAssemble, COMMD, X1, X2, X3, ret
            Else
                ret = PerformCMD3(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16, ISWP)
                COMMD = MATHOFFSET(XRG)
                LJoin DPL, DisAssemble, COMMD, X1, X2, X3, ret
            End If
    
        Case &H83 '+
            If ISWP = 1 Then
                ret = PerformCMD2(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, ZRG, OPT16)
                COMMD = MATHOFFSET(ZRG)
                'CX3 = CByte(Val("&H" & X3))
                CX3 = CByte("&H" & X3)
                If CX3 > &H7F Then X3 = WordToStr(CX3 Or &HFF00)
                LJoin WPL, DisAssemble, COMMD, X1, X2, X3, ret
            Else
                ret = PerformCMD2(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, ZRG, OPT16)
                COMMD = MATHOFFSET(ZRG)
                'CX3 = CByte(Val("&H" & X3))
                CX3 = CByte("&H" & X3)
                If CX3 > &H7F Then X3 = DwordToStr(CX3 Or &HFFFFFF00)
                LJoin DPL, DisAssemble, COMMD, X1, X2, X3, ret
            End If
    
        Case &H84
            COMMD = "TEST"
InTTS2:
    ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, OPT16)
    X3 = REGOFFSET1(XRG)
    LJoin BPL, DisAssemble, COMMD, X1, X2, X3, ret
    
        Case &H85
            COMMD = "TEST"
InTTS4:
    If ISWP = 1 Then
        ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
        X3 = REGOFFSET2(XRG)
        LJoin WPL, DisAssemble, COMMD, X1, X2, X3, ret
    Else
        ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
        X3 = REGOFFSET4(XRG)
        LJoin DPL, DisAssemble, COMMD, X1, X2, X3, ret
    End If
    
        Case &H86
            COMMD = "XCHG"
InTTSR2:
    ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, OPT16)
    X3 = REGOFFSET1(XRG)
    RJoin BPR, DisAssemble, COMMD, X1, X2, X3, ret
    
        Case &H87
            COMMD = "XCHG"
InTTSR4:
    If ISWP = 1 Then
        ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
        X3 = REGOFFSET2(XRG)
        RJoin WPR, DisAssemble, COMMD, X1, X2, X3, ret
    Else
        ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
        X3 = REGOFFSET4(XRG)
        RJoin DPR, DisAssemble, COMMD, X1, X2, X3, ret
    End If
    
        Case &H8C
            COMMD = "MOV"
            ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
            X3 = SEGOFFSET(XRG)
            LJoin WPL, DisAssemble, COMMD, X1, X2, X3, ret
    
        Case &H8E
            COMMD = "MOV"
            ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
            X3 = SEGOFFSET(XRG)
            RJoin WPR, DisAssemble, COMMD, X1, X2, X3, ret
    
        Case &H8D '+
            If DB2 >= &HC0 Then
                GoTo EendX
            Else
                COMMD = "LEA"
                If ISWP = 1 Then
                    ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
                    X3 = REGOFFSET2(XRG)
                    RJoin NOR, DisAssemble, COMMD, X1, X2, X3, ret
                Else
                    ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
                    X3 = REGOFFSET4(XRG)
                    RJoin NOR, DisAssemble, COMMD, X1, X2, X3, ret
                End If
            End If
    
        Case &H8F '+
            COMMD = "POP"
            ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
            If ISWP = 1 Then
                MJoin WPL, DisAssemble, COMMD, X1, X2, ret
            Else
                MJoin DPL, DisAssemble, COMMD, X1, X2, ret
            End If
    
        Case &H90
            DisAssemble = "NOP"
            CL = 1
    
        Case &H91 To &H97
            If ISWP = 1 Then
                DisAssemble = "XCHG AX," & REGOFFSET2(MCs And 7)
            Else
                DisAssemble = "XCHG EAX," & REGOFFSET4(MCs And 7)
            End If
            CL = 1
    
        Case &H98
            If ISWP = 1 Then
                DisAssemble = "CBW"
            Else
                DisAssemble = "CWDE"
            End If
            CL = 1
    
        Case &H99
            If ISWP = 1 Then
                DisAssemble = "CWD"
            Else
                DisAssemble = "CDQ"
            End If
            CL = 1
    
        Case &H9A
            If ISWP = 1 Then
                DisAssemble = "CALL " & WordToStr(GetWordFromList(Data, Start + 3)) & ":" & _
                WordToStr(GetWordFromList(Data, Start + 1))
                CL = 5
            Else
                DisAssemble = "CALL " & WordToStr(GetWordFromList(Data, Start + 5)) & ":" & _
                DwordToStr(GetDWordFromList(Data, Start + 1))
                CL = 7
            End If
    
        Case &H9B
            DisAssemble = "WAIT"
            CL = 1
    
        Case &H9C
            If ISWP = 1 Then
                DisAssemble = "PUSHF"
            Else
                DisAssemble = "PUSHFD"
            End If
            CL = 1
    
        Case &H9D
            If ISWP = 1 Then
                DisAssemble = "POPF"
            Else
                DisAssemble = "POPFD"
            End If
            CL = 1
    
        Case &H9E
            DisAssemble = "SAHF"
            CL = 1
    
        Case &H9F
            DisAssemble = "LAHF"
            CL = 1
    
        Case &HA0
            DisAssemble = "MOV AL,BYTE PTR [" & DwordToStr(GetDWordFromList(Data, Start + 1)) & "]"
            CL = 5
    
        Case &HA1
            If ISWP = 1 Then
                DisAssemble = "MOV AX,WORD PTR [" & DwordToStr(GetDWordFromList(Data, Start + 1)) & "]"
            Else
                DisAssemble = "MOV EAX,DWORD PTR [" & DwordToStr(GetDWordFromList(Data, Start + 1)) & "]"
            End If
            CL = 5
    
        Case &HA2
            DisAssemble = "MOV BYTE PTR [" & DwordToStr(GetDWordFromList(Data, Start + 1)) & "],AL"
            CL = 5
    
        Case &HA3
            If ISWP = 1 Then
                DisAssemble = "MOV WORD PTR [" & DwordToStr(GetDWordFromList(Data, Start + 1)) & "],AX"
            Else
                DisAssemble = "MOV DWORD PTR [" & DwordToStr(GetDWordFromList(Data, Start + 1)) & "],EAX"
            End If
            CL = 5
    
        Case &HA4
            DisAssemble = "MOVS BYTE PTR [EDI],BYTE PTR [ESI]"
            CL = 1
    
        Case &HA5
            If ISWP = 1 Then
                DisAssemble = "MOVS WORD PTR [EDI],WORD PTR [ESI]"
            Else
                DisAssemble = "MOVS DWORD PTR [EDI],DWORD PTR [ESI]"
            End If
            CL = 1
    
        Case &HA6
            DisAssemble = "CMPS BYTE PTR [ESI],BYTE PTR [EDI]"
            CL = 1
    
        Case &HA7
            If ISWP = 1 Then
                DisAssemble = "CMPS WORD PTR [ESI],WORD PTR [EDI]"
            Else
                DisAssemble = "CMPS DWORD PTR [ESI],DWORD PTR [EDI]"
            End If
            CL = 1
    
        Case &HA8
            DisAssemble = "TEST AL," & ByteToStr(DB2)
            CL = 2
    
        Case &HA9
            If ISWP = 1 Then
                DisAssemble = "TEST AX," & WordToStr(GetWordFromList(Data, Start + 1))
                CL = 3
            Else
                DisAssemble = "TEST EAX," & DwordToStr(GetDWordFromList(Data, Start + 1))
                CL = 5
            End If
    
        Case &HAA
            DisAssemble = "STOS BYTE PTR [EDI]"
            CL = 1
    
        Case &HAB
            If ISWP = 1 Then
                DisAssemble = "STOS WORD PTR [EDI]"
            Else
                DisAssemble = "STOS DWORD PTR [EDI]"
            End If
            CL = 1
    
        Case &HAC
            DisAssemble = "LODS BYTE PTR [ESI]"
            CL = 1
    
        Case &HAD
            If ISWP = 1 Then
                DisAssemble = "LODS WORD PTR [ESI]"
            Else
                DisAssemble = "LODS DWORD PTR [ESI]"
            End If
            CL = 1
    
        Case &HAE
            DisAssemble = "SCAS BYTE PTR [EDI]"
            CL = 1
    
        Case &HAF
            If ISWP = 1 Then
                DisAssemble = "SCAS WORD PTR [EDI]"
            Else
                DisAssemble = "SCAS DWORD PTR [EDI]"
            End If
            CL = 1
    
        Case &HB0 To &HB7
            DisAssemble = "MOV " & REGOFFSET1(MCs And 7) & "," & ByteToStr(DB2)
            CL = 2
    
        Case &HB8 To &HBF
            If ISWP = 1 Then
                DisAssemble = "MOV " & REGOFFSET2(MCs And 7) & "," & WordToStr(GetWordFromList(Data, Start + 1))
                CL = 3
            Else
                DisAssemble = "MOV " & REGOFFSET4(MCs And 7) & "," & DwordToStr(GetDWordFromList(Data, Start + 1))
                CL = 5
            End If
    
        Case &HC0
            ret = PerformCMD2(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, ZRG, OPT16)
            COMMD = BITOFFSET(ZRG)
            LJoin BPL, DisAssemble, COMMD, X1, X2, X3, ret
    
        Case &HC1
            If ISWP = 1 Then
                ret = PerformCMD2(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, ZRG, OPT16)
                COMMD = BITOFFSET(ZRG)
                LJoin WPL, DisAssemble, COMMD, X1, X2, X3, ret
            Else
                ret = PerformCMD2(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, ZRG, OPT16)
                COMMD = BITOFFSET(ZRG)
                LJoin DPL, DisAssemble, COMMD, X1, X2, X3, ret
            End If
    
        Case &HC2
            DisAssemble = "RET " & WordToStr(GetWordFromList(Data, Start + 1))
            CL = 3
    
        Case &HC3
            DisAssemble = "RET"
            CL = 1
    
        Case &HC4
            COMMD = "LES"
            GoTo InLDX
    
        Case &HC5
            COMMD = "LDS"
InLDX:
    If DB2 >= &HC0 Then GoTo EendX
    If ISWP = 1 Then
        ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
        X3 = REGOFFSET2(XRG)
        RJoin DPR, DisAssemble, COMMD, X1, X2, X3, ret
    Else
        ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
        X3 = REGOFFSET4(XRG)
        RJoin FPR, DisAssemble, COMMD, X1, X2, X3, ret
    End If
    
        Case &HC6
            COMMD = "MOV"
            ret = PerformCMD2(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, ZRG, OPT16)
            LJoin BPL, DisAssemble, COMMD, X1, X2, X3, ret
    
        Case &HC7 '+
            COMMD = "MOV"
            If ISWP = 1 Then
                ret = PerformCMD3(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16, ISWP)
                LJoin WPL, DisAssemble, COMMD, X1, X2, X3, ret
            Else
                ret = PerformCMD3(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16, ISWP)
                LJoin DPL, DisAssemble, COMMD, X1, X2, X3, ret
            End If
    
        Case &HC8
            DisAssemble = "ENTER " & WordToStr(GetWordFromList(Data, Start + 1)) & "," & _
            ByteToStr(Data(Start + 3))
            CL = 4
    
        Case &HC9
            DisAssemble = "LEAVE"
            CL = 1
    
        Case &HCA
            DisAssemble = "RETF " & WordToStr(GetWordFromList(Data, Start + 1))
            CL = 3
    
        Case &HCB
            DisAssemble = "RETF"
            CL = 1
    
        Case &HCC
            DisAssemble = "INT 3"
            CL = 1
    
        Case &HCD
            DisAssemble = "INT " & ByteToStr(DB2)
            CL = 2
    
        Case &HCE
            DisAssemble = "INTO"
            CL = 1
    
        Case &HCF
            DisAssemble = "IRETD"
            CL = 1
    
        Case &HD0
            ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, OPT16)
            COMMD = BITOFFSET(XRG)
            X3 = "1"
            LJoin BPL, DisAssemble, COMMD, X1, X2, X3, ret
    
        Case &HD1 '+
            If ISWP = 1 Then
                ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
                COMMD = BITOFFSET(XRG)
                X3 = "1"
                LJoin WPL, DisAssemble, COMMD, X1, X2, X3, ret
            Else
                ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
                COMMD = BITOFFSET(XRG)
                X3 = "1"
                LJoin DPL, DisAssemble, COMMD, X1, X2, X3, ret
            End If
    
        Case &HD2
            ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, OPT16)
            COMMD = BITOFFSET(XRG)
            X3 = "CL"
            LJoin BPL, DisAssemble, COMMD, X1, X2, X3, ret
    
        Case &HD3 '+
            If ISWP = 1 Then
                ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
                COMMD = BITOFFSET(XRG)
                X3 = "CL"
                LJoin WPL, DisAssemble, COMMD, X1, X2, X3, ret
            Else
                ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
                COMMD = BITOFFSET(XRG)
                X3 = "CL"
                LJoin DPL, DisAssemble, COMMD, X1, X2, X3, ret
            End If
    
        Case &HD4
            DisAssemble = "AAMB " & ByteToStr(DB2)
            CL = 2
    
        Case &HD5
            DisAssemble = "AADB " & ByteToStr(DB2)
            CL = 2
    
        Case &HD7
            DisAssemble = "XLAT BYTE PTR [EBX]"
            CL = 1
    
        Case &HFE
            If FWR > 1 Then GoTo EendX
            ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, OPT16)
            COMMD = IDCJP(XRG)
            MJoin BPL, DisAssemble, COMMD, X1, X2, ret
    
        Case &HFF '+
            If FWR = 7 Then GoTo EendX
            If DB2 >= &HC0 Then If FWR = 3 Or FWR = 5 Then GoTo EendX
            If ISWP = 1 Then
                ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
                COMMD = IDCJP(XRG)
                If FWR = 3 Or FWR = 5 Then
                    MJoin DPL, DisAssemble, COMMD, X1, X2, ret
                Else
                    MJoin WPL, DisAssemble, COMMD, X1, X2, ret
                End If
            Else
                ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
                COMMD = IDCJP(XRG)
                If FWR = 3 Or FWR = 5 Then
                    MJoin FPL, DisAssemble, COMMD, X1, X2, ret
                Else
                    MJoin DPL, DisAssemble, COMMD, X1, X2, ret
                End If
            End If
    
        'FLOAT*
        Case &HD8
            ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, FLOATSTACK, XRG, OPT16)
            If DB2 >= &HC0 Then X1 = "ST(0)," & X1
            COMMD = FLOATOFFSET(XRG)
            MJoin DPL, DisAssemble, COMMD, X1, X2, ret
    
        Case &HDA
            If DB2 >= &HC0 Then
                CL = 2
                CHK = FLOATOP5(DB2)
                If Len(CHK) = 0 Then GoTo EendX
                DisAssemble = CHK
            Else
                ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, FLOATSTACK, XRG, OPT16)
                COMMD = INTFLOATOFFSET(XRG)
                MJoin DPL, DisAssemble, COMMD, X1, X2, ret
            End If
    
        Case &HD9
            If DB2 >= &HC0 Then
                CL = 2
                CHK = FLOATOP3(DB2)
                If Len(CHK) = 0 Then GoTo EendX
                DisAssemble = CHK
            Else
                If FWR = 1 Then GoTo EendX
                ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, FLOATSTACK, XRG, OPT16)
                COMMD = FLOATOP1(XRG)
                Select Case FWR
                    Case Is = 0, 2, 3
                        MJoin DPL, DisAssemble, COMMD, X1, X2, ret
                    Case Is = 4, 6
                        MJoin NOL, DisAssemble, COMMD, X1, X2, ret
                    Case Is = 5, 7
                        MJoin WPL, DisAssemble, COMMD, X1, X2, ret
                End Select
            End If
    
        Case &HDB
            If DB2 >= &HC0 Then
                CL = 2
                CHK = FLOATOP4(DB2)
                If Len(CHK) = 0 Then GoTo EendX
                DisAssemble = CHK
            Else
                If FWR = 1 Or FWR = 4 Or FWR = 6 Then GoTo EendX
                ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, FLOATSTACK, XRG, OPT16)
                COMMD = FLOATOP2(XRG)
                Select Case FWR
                    Case Is = 0, 2, 3
                        MJoin DPL, DisAssemble, COMMD, X1, X2, ret
                    Case Is = 5, 7
                        MJoin TBPL, DisAssemble, COMMD, X1, X2, ret
                End Select
            End If
    
        Case &HDC
            ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, FLOATSTACK, XRG, OPT16)
            If DB2 >= &HC0 Then
                X1 = X1 & ",ST(0)"
                If FWR = 2 Then
                    COMMD = FLOATOFFSET(XRG) & "2"
                ElseIf FWR = 3 Then
                    COMMD = FLOATOFFSET(XRG) & "3"
                ElseIf FWR = 4 Or FWR = 6 Then
                    COMMD = FLOATOFFSET(XRG + 1)
                ElseIf FWR = 5 Or FWR = 7 Then
                    COMMD = FLOATOFFSET(XRG - 1)
                Else
                    COMMD = FLOATOFFSET(XRG)
                End If
            Else
                COMMD = FLOATOFFSET(XRG)
                End If
            MJoin QPL, DisAssemble, COMMD, X1, X2, ret
    
        Case &HDD
            ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, FLOATSTACK, XRG, OPT16)
            If DB2 >= &HC0 Then
                If FWR > 5 Then GoTo EendX
                COMMD = FLOATOP7(XRG)
                MJoin QPL, DisAssemble, COMMD, X1, X2, ret
            Else
                If FWR = 1 Or FWR = 5 Then GoTo EendX
                COMMD = FLOATOP6(XRG)
                Select Case FWR
                    Case Is = 0, 2, 3
                        MJoin QPL, DisAssemble, COMMD, X1, X2, ret
                    Case Is = 4, 6
                        MJoin NOL, DisAssemble, COMMD, X1, X2, ret
                    Case 7
                        MJoin WPL, DisAssemble, COMMD, X1, X2, ret
                End Select
            End If
    
        Case &HDE
            If DB2 >= &HC0 Then
                If DB2 = &HD9 Then
                    DisAssemble = "FCOMPP"
                    CL = 2
                ElseIf FWR = 3 Then
                    GoTo EendX
                Else
                    ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, FLOATSTACK, XRG, OPT16)
                    COMMD = FLOATOP8(XRG)
                    If FWR <> 2 Then X1 = X1 & ",ST(0)"
                    MJoin WPL, DisAssemble, COMMD, X1, X2, ret
                End If
            Else
                ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, FLOATSTACK, XRG, OPT16)
                COMMD = INTFLOATOFFSET(XRG)
                MJoin WPL, DisAssemble, COMMD, X1, X2, ret
            End If
    
        Case &HDF
            If DB2 >= &HC0 Then
                If DB2 = &HE0 Then
                    DisAssemble = "FNSTSW AX"
                    CL = 2
                ElseIf FWR = 4 Or FWR = 5 Or FWR = 7 Then
                    GoTo EendX
                Else
                    ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, FLOATSTACK, XRG, OPT16)
                    COMMD = FLOATOP10(XRG)
                    If FWR = 6 Then COMMD = COMMD & " ST(0),"
                    MJoin WPL, DisAssemble, COMMD, X1, X2, ret
                End If
            Else
                Select Case FWR
                    Case Is = 0, 2, 3
                        ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, FLOATSTACK, XRG, OPT16)
                        COMMD = FLOATOP9(XRG)
                        MJoin WPL, DisAssemble, COMMD, X1, X2, ret
                    Case Is = 4, 6
                        ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, FLOATSTACK, XRG, OPT16)
                        COMMD = FLOATOP9(XRG)
                        MJoin TBPL, DisAssemble, COMMD, X1, X2, ret
                    Case Is = 5, 7
                        ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, FLOATSTACK, XRG, OPT16)
                        COMMD = FLOATOP9(XRG)
                        MJoin QPL, DisAssemble, COMMD, X1, X2, ret
                    Case 1
                    GoTo EendX
                End Select
            End If
    
        Case &HE0 To &HE3
            DisAssemble = CalcShortJump(DB2, LOOPX(MCs And &H3), Start + 1, VBASE + Start)
            CL = 2
    
        Case &HE4
            DisAssemble = "IN AL," & ByteToStr(DB2)
            CL = 2
    
        Case &HE5
            If ISWP = 1 Then
                DisAssemble = "IN AX," & ByteToStr(DB2)
            Else
                DisAssemble = "IN EAX," & ByteToStr(DB2)
            End If
            CL = 2
    
        Case &HE6
            DisAssemble = "OUT " & ByteToStr(DB2) & ",AL"
            CL = 2
    
        Case &HE7
            If ISWP = 1 Then
                DisAssemble = "OUT " & ByteToStr(DB2) & ",AX"
            Else
                DisAssemble = "OUT " & ByteToStr(DB2) & ",EAX"
            End If
            CL = 2
    
        Case &HE8
            DisAssemble = CalcLongJump(Data, "CALL", Start, VBASE + Start - 1)
            CL = 5
    
        Case &HE9
            DisAssemble = CalcLongJump(Data, "JMP", Start, VBASE + Start - 1)
            CL = 5
    
        Case &HEA '+
            If ISWP = 1 Then
                DisAssemble = "JMP " & WordToStr(GetWordFromList(Data, Start + 3)) & ":" & _
                WordToStr(GetWordFromList(Data, Start + 1))
                CL = 5
            Else
                DisAssemble = "JMP " & WordToStr(GetWordFromList(Data, Start + 5)) & ":" & _
                DwordToStr(GetDWordFromList(Data, Start + 1))
                CL = 7
            End If
    
        Case &HEB
            DisAssemble = CalcShortJump(DB2, "JMP", Start + 1, VBASE + Start)
            CL = 2
    
        Case &HEC
            DisAssemble = "IN AL,DX"
            CL = 1
    
        Case &HED
            If ISWP = 1 Then
                DisAssemble = "IN AX,DX"
            Else
                DisAssemble = "IN EAX,DX"
            End If
            CL = 1
    
        Case &HEE
            DisAssemble = "OUT DX,AL"
            CL = 1
    
        Case &HEF
            If ISWP = 1 Then
                DisAssemble = "OUT DX,AX"
            Else
                DisAssemble = "OUT DX,EAX"
            End If
            CL = 1
    
        Case &HF0
            Start = Start + 1
            EXTRA = EXTRA + 1
            EXTRAADD = 1
            GoTo BEGGIN
    
        Case &HF2
            Start = Start + 1
            EXTRA = EXTRA + 1
            EXTRAADD = 2
            GoTo BEGGIN
    
        Case &HF3
            Start = Start + 1
            EXTRA = EXTRA + 1
            EXTRAADD = 3
            GoTo BEGGIN
    
        Case &HF4
            DisAssemble = "HLT"
            CL = 1
    
        Case &HF5
            DisAssemble = "CMC"
            CL = 1
    
        Case &HF6
            Select Case FWR
                Case Is = 0, 1
                    ret = PerformCMD2(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, ZRG, OPT16)
                    COMMD = "TEST"
                    LJoin BPL, DisAssemble, COMMD, X1, X2, X3, ret
                Case Is = 2, 3, 5, 7
                    ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, OPT16)
                    COMMD = MATH2(XRG)
                    MJoin BPL, DisAssemble, COMMD, X1, X2, ret
                Case Is = 4, 6
                    ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, OPT16)
                    COMMD = MATH2(XRG)
                    X3 = "AL"
                    RJoin BPR, DisAssemble, COMMD, X1, X2, X3, ret
            End Select
    
        Case &HF7 '+
            Select Case FWR
                Case Is = 0, 1
                    If ISWP = 1 Then
                        ret = PerformCMD3(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16, ISWP)
                        COMMD = "TEST"
                        LJoin WPL, DisAssemble, COMMD, X1, X2, X3, ret
                    Else
                        ret = PerformCMD3(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16, ISWP)
                        COMMD = "TEST"
                        LJoin DPL, DisAssemble, COMMD, X1, X2, X3, ret
                    End If
    
                Case Is = 2, 3, 5
                    If ISWP = 1 Then
                        ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
                        COMMD = MATH2(XRG)
                        MJoin WPL, DisAssemble, COMMD, X1, X2, ret
                    Else
                        ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
                        COMMD = MATH2(XRG)
                        MJoin DPL, DisAssemble, COMMD, X1, X2, ret
                    End If
    
                Case Is = 4, 6, 7
                    If ISWP = 1 Then
                        ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
                        COMMD = MATH2(XRG)
                        X3 = "AX"
                        RJoin WPR, DisAssemble, COMMD, X1, X2, X3, ret
                    Else
                        ret = PerformCMD(Data, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
                        COMMD = MATH2(XRG)
                        X3 = "EAX"
                        RJoin DPR, DisAssemble, COMMD, X1, X2, X3, ret
                    End If
    
            End Select
    
        Case &HF8
            DisAssemble = "CLC"
            CL = 1
    
        Case &HF9
            DisAssemble = "STC"
            CL = 1
    
        Case &HFA
            DisAssemble = "CLI"
            CL = 1
    
        Case &HFB
            DisAssemble = "STI"
            CL = 1
    
        Case &HFC
            DisAssemble = "CLD"
            CL = 1
    
        Case &HFD
            DisAssemble = "STD"
            CL = 1
    
        Case Else
            GoTo EendX
    
    End Select
    
    Forward = CL + DL + EXTRA
    
    If EXTRAADD = 1 Then
        DisAssemble = "LOCK " & DisAssemble
    ElseIf EXTRAADD = 2 Then
        DisAssemble = "REPNE " & DisAssemble
    ElseIf EXTRAADD = 3 Then
        DisAssemble = "REP " & DisAssemble
    End If
    
    If BinDump = 1 Then
        RWDump Data, Start, VBASE, Forward, DisAssemble
    End If
OutOfF:
    If HexAdr = 1 Then
        tmpStr = ""
        newStr = ""
        For I = 0 To Forward - 1
            tmp = Data(Start + I)
            tmpStr = Hex(tmp)
            If Len(tmpStr) = 1 Then
                tmpStr = "0" & tmpStr
            End If
            newStr = newStr & tmpStr & " "
        Next I
        tmpStr = newStr & Space(15 - Len(newStr))
        newStr = tmpStr
        DisAssemble = Hex(VBASE + Start - EXTRA) & vbTab & newStr & vbTab & DisAssemble
    
    End If
    

    Exit Function
    
EendX:
    On Error GoTo 0
    DisAssemble = "db " & BByteToStr(Data(Start - EXTRA)): Forward = 1
    If HexAdr = 1 Then
        DisAssemble = Hex(VBASE + Start - EXTRA) & " " & DisAssemble
    End If

End Function

Friend Function PerformCMD2(Data() As Byte, ByRef Start As Long, ByRef COMMD As String, ByRef X1 As String, ByRef X2 As String, ByRef X3 As String, ByRef CL As Byte, ByRef DL As Byte, R1() As String, R2() As String, R3() As String, ByRef XRG As Byte, ByRef ZRG As Byte, ByRef OPT16 As Byte) As Byte
    Dim XD As Byte 'Data
    Dim XOO As Byte 'OCTAL OFFSET
    Dim YD As Byte 'Data2
    Dim YOO As Byte 'OCTAL OFFSET
    Dim YRG As Byte 'REGS OFFSET
    Dim MPO As Byte 'MULTIPLY
    
    XD = Data(Start + 1)
    XOO = XD Mod 8&
    XRG = Int((XD And &H3F) / 8)
    If OPT16 = 1 Then GoTo OP16X

    If XD >= 0 And XD <= &H3F Then
        If XOO = 4 Then
            YD = Data(Start + 2)
            YOO = YD Mod 8&
            YRG = Int(YD / 8)
            ZRG = Int(XD - 5) / 8
            If (YRG Mod 8&) = 4 Then
                If YOO = 5 Then
                    CL = 3
                    DL = 5
                    X1 = DwordToStr(GetDWordFromList(Data, Start + CL))
                Else
                    CL = 3
                    DL = 1
                    X1 = R1(YOO)
                End If
            ElseIf YOO = 5 Then
                CL = 3
                DL = 5
                MPO = 2 ^ Int(YD / &H40)
                If MPO < 2 Then
                    X1 = R2(YRG Mod 8&)
                Else
                    X1 = R2(YRG Mod 8&) & "*" & CStr(MPO)
                End If
                X2 = SDwordToStr(GetDWordFromList(Data, Start + CL))
            Else
                MPO = 2 ^ Int(YD / &H40)
                If MPO < 2 Then
                    X2 = "+" & R2(YRG Mod 8&)
                Else
                    X2 = "+" & R2(YRG Mod 8&) & "*" & CStr(MPO)
                End If
                CL = 3
                DL = 1
                X1 = R1(YOO)
            End If
        ElseIf XOO = 5 Then
            CL = 2
            DL = 5
            ZRG = Int(XD - 5) / 8
            X1 = DwordToStr(GetDWordFromList(Data, Start + CL))
        Else
            CL = 2
            DL = 1
            ZRG = XRG
            X1 = R1(XOO)
        End If
        X3 = ByteToStr(Data(Start + CL + DL - 1))
    ElseIf XD >= &H40 And XD <= &H7F Then
        ZRG = XRG
        If XOO = 4 Then
            CL = 3
            DL = 2
            YD = Data(Start + 2)
            YOO = YD Mod 8&
            YRG = Int(YD / 8)
            If (YRG Mod 8&) = 4 Then
                X1 = R1(YOO)
                X2 = SByteToStr(Data(Start + CL))
            Else
                X1 = R1(YOO)
                MPO = 2 ^ Int(YD / &H40)
                If MPO < 2 Then
                    X2 = "+" & R2(YRG Mod 8&) & SByteToStr(Data(Start + CL))
                Else
                    X2 = "+" & R2(YRG Mod 8&) & "*" & CStr(MPO) & SByteToStr(Data(Start + CL))
                End If
            End If
        Else
            CL = 2
            DL = 2
            X1 = R1(XOO)
            X2 = SByteToStr(Data(Start + CL))
        End If
        X3 = ByteToStr(Data(Start + CL + DL - 1))
    ElseIf XD >= &H80 And XD <= &HBF Then
        If XOO = 4 Then
            CL = 3
            DL = 5
            YD = Data(Start + 2)
            YOO = YD Mod 8
            YRG = Int(YD / 8)
            If (YRG Mod 8&) = 4 Then
                X2 = SDwordToStr(GetDWordFromList(Data, Start + CL))
                X1 = R1(YOO)
            Else
                X1 = R1(YOO)
                MPO = 2 ^ Int(YD / &H40)
                If MPO < 2 Then
                    X2 = "+" & R2(YRG Mod 8&) & SDwordToStr(GetDWordFromList(Data, Start + CL))
                Else
                    X2 = "+" & R2(YRG Mod 8&) & "*" & CStr(MPO) & SDwordToStr(GetDWordFromList(Data, Start + CL))
                End If
            End If
        Else
            CL = 2
            DL = 5
            X1 = R1(XOO)
            X2 = SDwordToStr(GetDWordFromList(Data, Start + CL))
        End If
    
        X3 = ByteToStr(Data(Start + CL + DL - 1))
        ZRG = XRG
    ElseIf XD >= &HC0 Then
        CL = 2
        DL = 1
        X3 = ByteToStr(Data(Start + CL))
        X1 = R3(XOO)
        PerformCMD2 = 1
        ZRG = XRG
    End If

Exit Function

OP16X:
    ZRG = XRG
    If XD >= 0 And XD <= &H3F Then
        If XOO = 6 Then
            CL = 2
            DL = 3
            X1 = WordToStr(GetWordFromList(Data, Start + CL))
            X3 = ByteToStr(Data(Start + CL + 2))
        Else
            X1 = REG16O1(XOO)
            CL = 2
            DL = 1
            X3 = ByteToStr(Data(Start + CL))
        End If
    ElseIf XD >= &H40 And XD <= &H7F Then
        CL = 2
        DL = 2
        X2 = SByteToStr(Data(Start + CL))
        X1 = REG16O1(XOO)
        X3 = ByteToStr(Data(Start + CL + 1))
    ElseIf XD >= &H80 And XD <= &HBF Then
        CL = 2
        DL = 3
        X2 = "+" & WordToStr(GetWordFromList(Data, Start + CL))
        X1 = REG16O1(XOO)
        X3 = ByteToStr(Data(Start + CL + 2))
    ElseIf XD >= &HC0 Then
        CL = 2
        DL = 1
        X3 = ByteToStr(Data(Start + CL))
        X1 = R3(XOO)
        PerformCMD2 = 1
    End If
    
End Function

Friend Function PerformCMD3(Data() As Byte, ByRef Start As Long, ByRef COMMD As String, ByRef X1 As String, ByRef X2 As String, ByRef X3 As String, ByRef CL As Byte, ByRef DL As Byte, R1() As String, R2() As String, R3() As String, ByRef XRG As Byte, ByRef OPT16 As Byte, Optional ByRef ISWP As Byte) As Byte
    Dim XD As Byte 'Data
    Dim XOO As Byte 'OCTALNI OFFSET
    Dim YD As Byte 'Data2
    Dim YOO As Byte 'OCTALNI OFFSET
    Dim YRG As Byte 'REGS OFFSET
    Dim MPO As Byte 'MULTIPLY

    XD = Data(Start + 1)
    XOO = XD Mod 8&
    XRG = Int((XD And &H3F) / 8)
    If OPT16 = 1 Then GoTo OPT16X
    If XD >= 0 And XD <= &H3F Then
        If XOO = 5 Then
            CL = 1
            X2 = DwordToStr(GetDWordFromList(Data, Start + CL + 1))
    
            If ISWP = 1 Then
                X3 = WordToStr(GetWordFromList(Data, Start + CL + 5))
                DL = 7
            Else
                X3 = DwordToStr(GetDWordFromList(Data, Start + CL + 5))
                DL = 9
            End If
    
        ElseIf XOO = 4 Then
            YD = Data(Start + 2)
            YOO = YD Mod 8&
            YRG = Int(YD / 8)
            If (YRG Mod 8&) = 4 Then
                If YOO = 5 Then
                    CL = 2
                    X1 = DwordToStr(GetDWordFromList(Data, Start + CL + 1))
                    If ISWP = 1 Then
                        X3 = WordToStr(GetWordFromList(Data, Start + CL + 5))
                        DL = 7
                    Else
                        X3 = DwordToStr(GetDWordFromList(Data, Start + CL + 5))
                        DL = 9
                    End If
                Else
                    CL = 2
                    X1 = R1(YOO)
                    If ISWP = 1 Then
                        X3 = WordToStr(GetWordFromList(Data, Start + CL + 1))
                        DL = 3
                    Else
                        X3 = DwordToStr(GetDWordFromList(Data, Start + CL + 1))
                        DL = 5
                    End If
                End If
            ElseIf YOO = 5 Then
                MPO = 2 ^ Int(YD / &H40)
                If MPO < 2 Then
                    X1 = R2(YRG Mod 8&)
                Else
                    X1 = R2(YRG Mod 8&) & "*" & CStr(MPO)
                End If
                CL = 2
                X2 = SDwordToStr(GetDWordFromList(Data, Start + CL + 1))
                If ISWP = 1 Then
                    X3 = WordToStr(GetWordFromList(Data, Start + CL + 5))
                    DL = 7
                Else
                    X3 = DwordToStr(GetDWordFromList(Data, Start + CL + 5))
                    DL = 9
                End If
            Else
                CL = 2
                MPO = 2 ^ Int(YD / &H40)
                If MPO < 2 Then
                    X2 = "+" & R2(YRG Mod 8&)
                Else
                    X2 = "+" & R2(YRG Mod 8&) & "*" & CStr(MPO)
                End If
                X1 = R1(YOO)
                If ISWP = 1 Then
                    X3 = WordToStr(GetWordFromList(Data, Start + CL + 1))
                    DL = 3
                Else
                    X3 = DwordToStr(GetDWordFromList(Data, Start + CL + 1))
                    DL = 5
                End If
            End If
        Else
            CL = 1
            X1 = R1(XOO)
            If ISWP = 1 Then
                X3 = WordToStr(GetWordFromList(Data, Start + CL + 1))
                DL = 3
            Else
                X3 = DwordToStr(GetDWordFromList(Data, Start + CL + 1))
                DL = 5
            End If
        End If
    ElseIf XD >= &H40 And XD <= &H7F Then
        If XOO = 4 Then
            YD = Data(Start + 2)
            YOO = YD Mod 8&
            YRG = Int(YD / 8)
            CL = 3
            If (YRG Mod 8&) = 4 Then
                X1 = R1(YOO) & SByteToStr(Data(Start + CL))
                If ISWP = 1 Then
                    X3 = WordToStr(GetWordFromList(Data, Start + CL + 1))
                    DL = 3
                Else
                    X3 = DwordToStr(GetDWordFromList(Data, Start + CL + 1))
                    DL = 5
                End If
            Else
                X1 = R1(YOO)
                MPO = 2 ^ Int(YD / &H40)
                If MPO < 2 Then
                    X2 = "+" & R2(YRG Mod 8&) & SByteToStr(Data(Start + CL))
                Else
                    X2 = "+" & R2(YRG Mod 8&) & "*" & CStr(MPO) & SByteToStr(Data(Start + CL))
                End If
                If ISWP = 1 Then
                    X3 = WordToStr(GetWordFromList(Data, Start + CL + 1))
                    DL = 3
                Else
                    X3 = DwordToStr(GetDWordFromList(Data, Start + CL + 1))
                    DL = 5
                End If
            End If
        Else
            CL = 2
            If ISWP = 1 Then
                X3 = WordToStr(GetWordFromList(Data, Start + CL + 1))
                DL = 3
            Else
                X3 = DwordToStr(GetDWordFromList(Data, Start + CL + 1))
                DL = 5
            End If
            X2 = SByteToStr(Data(Start + CL))
            X1 = R1(XOO)
        End If
    ElseIf XD >= &H80 And XD <= &HBF Then
        If XOO = 4 Then
            CL = 2
            YD = Data(Start + 2)
            YOO = YD Mod 8&
            YRG = Int(YD / 8)
            If (YRG Mod 8&) = 4 Then
                X1 = R1(YOO)
                X2 = SDwordToStr(GetDWordFromList(Data, Start + CL + 1))
                If ISWP = 1 Then
                    X3 = WordToStr(GetWordFromList(Data, Start + CL + 5))
                    DL = 7
                Else
                    X3 = DwordToStr(GetDWordFromList(Data, Start + CL + 5))
                    DL = 9
                End If
            Else
                X1 = R1(YOO)
                MPO = 2 ^ Int(YD / &H40)
                If MPO < 2 Then
                    X2 = "+" & R2(YRG Mod 8&) & SDwordToStr(GetDWordFromList(Data, Start + CL + 1))
                Else
                    X2 = "+" & R2(YRG Mod 8&) & "*" & CStr(MPO) & SDwordToStr(GetDWordFromList(Data, Start + CL + 1))
                End If
                If ISWP = 1 Then
                    X3 = WordToStr(GetWordFromList(Data, Start + CL + 5))
                    DL = 7
                Else
                    X3 = DwordToStr(GetDWordFromList(Data, Start + CL + 5))
                    DL = 9
                End If
            End If
        Else
            CL = 1
            If ISWP = 1 Then
                X3 = WordToStr(GetWordFromList(Data, Start + CL + 5))
                DL = 7
            Else
                X3 = DwordToStr(GetDWordFromList(Data, Start + CL + 5))
                DL = 9
            End If
            X2 = SDwordToStr(GetDWordFromList(Data, Start + CL + 1))
            X1 = R1(XOO)
        End If
    ElseIf XD >= &HC0 Then
        PerformCMD3 = 1
        CL = 1
        X1 = R3(XOO)
        If ISWP = 1 Then
            X3 = WordToStr(GetWordFromList(Data, Start + CL + 1))
            DL = 3
        Else
            X3 = DwordToStr(GetDWordFromList(Data, Start + CL + 1))
            DL = 5
        End If
    End If

    Exit Function
OPT16X:
    If XD >= 0 And XD <= &H3F Then
        If XOO = 6 Then
            CL = 2
            X1 = WordToStr(GetWordFromList(Data, Start + CL))
            If ISWP = 1 Then
                X3 = WordToStr(GetWordFromList(Data, Start + CL + 2))
                DL = 4
            Else
                X3 = DwordToStr(GetDWordFromList(Data, Start + CL + 2))
                DL = 6
            End If
        Else
            CL = 2
            X1 = REG16O1(XOO)
            If ISWP = 1 Then
                X3 = WordToStr(GetWordFromList(Data, Start + CL))
                DL = 2
            Else
                X3 = DwordToStr(GetDWordFromList(Data, Start + CL))
                DL = 4
            End If
        End If
    ElseIf XD >= &H40 And XD <= &H7F Then
        CL = 2
        X1 = REG16O1(XOO)
        X2 = SByteToStr(Data(Start + CL))
        If ISWP = 1 Then
            X3 = WordToStr(GetWordFromList(Data, Start + CL + 1))
            DL = 3
        Else
            X3 = DwordToStr(GetDWordFromList(Data, Start + CL + 1))
            DL = 5
        End If
    ElseIf XD >= &H80 And XD <= &HBF Then
        CL = 2
        X1 = REG16O1(XOO)
        X2 = "+" & WordToStr(GetWordFromList(Data, Start + CL))
        If ISWP = 1 Then
            X3 = WordToStr(GetWordFromList(Data, Start + CL + 2))
            DL = 4
        Else
            X3 = DwordToStr(GetDWordFromList(Data, Start + CL + 2))
            DL = 6
        End If
    ElseIf XD >= &HC0 Then
        PerformCMD3 = 1
        CL = 1
        X1 = R3(XOO)
        If ISWP = 1 Then
            X3 = WordToStr(GetWordFromList(Data, Start + CL + 1))
            DL = 3
        Else
            X3 = DwordToStr(GetDWordFromList(Data, Start + CL + 1))
            DL = 5
        End If
    End If
    
End Function

Friend Function PerformCMD(Data() As Byte, ByRef Start As Long, ByRef COMMD As String, ByRef X1 As String, ByRef X2 As String, ByRef X3 As String, ByRef CL As Byte, ByRef DL As Byte, R1() As String, R2() As String, R3() As String, ByRef XRG As Byte, ByRef OPT16 As Byte) As Byte
    Dim XD As Byte 'Data
    Dim XOO As Byte 'OCTAL OFFSET
    Dim YD As Byte 'Data2
    Dim YOO As Byte 'OCTAL OFFSET
    Dim YRG As Byte 'REGS OFFSET
    Dim MPO As Byte 'MULTIPLY
    
    XD = Data(Start + 1)
    XOO = XD Mod 8&
    XRG = Int((XD And &H3F) / 8)
    If OPT16 = 1 Then GoTo OPT16X
    If XD >= 0 And XD <= &H3F Then
        If XOO = 5 Then
            CL = 2
            DL = 4
            X1 = DwordToStr(GetDWordFromList(Data, Start + CL))
        ElseIf XOO = 4 Then
            YD = Data(Start + 2)
            YOO = YD Mod 8&
            YRG = Int(YD / 8)
            If (YRG Mod 8&) = 4 Then
                If YOO = 5 Then
                    CL = 3
                    DL = 4
                    X1 = DwordToStr(GetDWordFromList(Data, Start + CL))
                Else
                    CL = 3
                    DL = 0
                    X1 = R1(YOO)
                End If
            ElseIf YOO = 5 Then
                MPO = 2 ^ Int(YD / &H40)
                If MPO < 2 Then
                    X1 = R2(YRG Mod 8&)
                Else
                    X1 = R2(YRG Mod 8&) & "*" & CStr(MPO)
                End If
                CL = 3
                DL = 4
                X2 = SDwordToStr(GetDWordFromList(Data, Start + CL))
            Else
                MPO = 2 ^ Int(YD / &H40)
                If MPO < 2 Then
                    X2 = "+" & R2(YRG Mod 8&)
                Else
                    X2 = "+" & R2(YRG Mod 8&) & "*" & CStr(MPO)
                End If
                X1 = R1(YOO)
                CL = 3
                DL = 0
            End If
        Else
            CL = 2
            DL = 0
            X1 = R1(XOO)
        End If
    ElseIf XD >= &H40 And XD <= &H7F Then
        If XOO = 4 Then
            CL = 3
            DL = 1
            YD = Data(Start + 2)
            YOO = YD Mod 8&
            YRG = Int(YD / 8)
            If (YRG Mod 8&) = 4 Then
                X1 = R1(YOO) & SByteToStr(Data(Start + CL))
            Else
                X1 = R1(YOO)
                MPO = 2 ^ Int(YD / &H40)
                If MPO < 2 Then
                    X2 = "+" & R2(YRG Mod 8&) & SByteToStr(Data(Start + CL))
                Else
                    X2 = "+" & R2(YRG Mod 8&) & "*" & CStr(MPO) & SByteToStr(Data(Start + CL))
                End If
            End If
        Else
            CL = 2
            DL = 1
            X2 = SByteToStr(Data(Start + CL))
            X1 = R1(XOO)
        End If
    ElseIf XD >= &H80 And XD <= &HBF Then
        If XOO = 4 Then
            CL = 3
            DL = 4
            YD = Data(Start + 2)
            YOO = YD Mod 8
            YRG = Int(YD / 8)
            If (YRG Mod 8&) = 4 Then
                X2 = SDwordToStr(GetDWordFromList(Data, Start + CL))
                X1 = R1(YOO)
            Else
                X1 = R1(YOO)
                MPO = 2 ^ Int(YD / &H40)
                If MPO < 2 Then
                    X2 = "+" & R2(YRG Mod 8&) & SDwordToStr(GetDWordFromList(Data, Start + CL))
                Else
                    X2 = "+" & R2(YRG Mod 8&) & "*" & CStr(MPO) & SDwordToStr(GetDWordFromList(Data, Start + CL))
                End If
            End If
        Else
            CL = 2
            DL = 4
            X1 = R1(XOO)
            X2 = SDwordToStr(GetDWordFromList(Data, Start + CL))
        End If
    ElseIf XD >= &HC0 Then
        PerformCMD = 1
        CL = 2
        X1 = R3(XOO)
    End If

    Exit Function
    
OPT16X:
    If XD >= 0 And XD <= &H3F Then
        If XOO = 6 Then
            CL = 2
            DL = 2
            X1 = WordToStr(GetWordFromList(Data, Start + CL))
        Else
            X1 = REG16O1(XOO)
            CL = 1
            DL = 1
        End If
    ElseIf XD >= &H40 And XD <= &H7F Then
        CL = 2
        DL = 1
        X1 = REG16O1(XOO)
        X2 = SByteToStr(Data(Start + CL))
    ElseIf XD >= &H80 And XD <= &HBF Then
        CL = 2
        DL = 2
        X2 = "+" & WordToStr(GetWordFromList(Data, Start + CL))
        X1 = REG16O1(XOO)
    ElseIf XD >= &HC0 Then
        PerformCMD = 1
        CL = 2
        X1 = R3(XOO)
    End If
    
End Function

Private Sub Class_Initialize()
    Static Init_ As Byte

    Term = False
    If Init_ = 1 Then Exit Sub
    Init_ = 1

End Sub
