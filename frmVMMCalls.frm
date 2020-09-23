VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVMMCalls 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DBL Click Mouse to Perform Function"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   Icon            =   "frmVMMCalls.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lstVMMCalls 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "VMM Call Functions"
      Top             =   240
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   11245
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Function Name"
         Object.Width           =   6350
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Service ID"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmVMMCalls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ReadVmmCalls()
    Dim vmmStr As String
    Dim tmpStr As String
    Dim sMax As Long
    Dim cnt As Long
    Dim pos As Integer
    Dim newStr As String
    Dim counter As Integer
    Dim itmX As ListItem
    Dim reg As outBuf
    Dim vxdHand As Long
    
    lstVMMCalls.ListItems.Clear
    sMax = &H192 * &H4
    counter = 0
    vmmStr = ReadByte(VMMServiceCalls, "", sMax)
    For cnt = 1 To sMax Step 4
        tmpStr = Mid(vmmStr, cnt, 4)
        newStr = ""
        pos = Asc(Mid(tmpStr, 4, 1))
        If pos < 16 Then
            newStr = newStr & "0" & Hex(pos)
        Else
            newStr = newStr & Hex(pos)
        End If
        pos = Asc(Mid(tmpStr, 3, 1))
        If pos < 16 Then
            newStr = newStr & "0" & Hex(pos)
        Else
            newStr = newStr & Hex(pos)
        End If
        pos = Asc(Mid(tmpStr, 2, 1))
        If pos < 16 Then
            newStr = newStr & "0" & Hex(pos)
        Else
            newStr = newStr & Hex(pos)
        End If
        pos = Asc(Mid(tmpStr, 1, 1))
        If pos < 16 Then
            newStr = newStr & "0" & Hex(pos)
        Else
            newStr = newStr & Hex(pos)
        End If
        Set itmX = lstVMMCalls.ListItems.add(, , newStr)
        itmX.SubItems(1) = SCall(counter)
        itmX.SubItems(2) = Hex(counter)
        counter = counter + 1
    Next cnt
    'Me.Hide
    If VMMHandle = 0 Then
        lstVMMCalls.ListItems.Item(4).Selected = True
        lstVMMCalls_DblClick
        VMMHandle = Not VMMHandle + 1
        'Comment these lines out if ou don't wish
        'to the the blue screen "inside of vvxd
        '
        newStr = "Sent from vvxd.vxd" & Chr(0) & "Inside of vvxd" & Chr(0)
        For cnt = 1 To Len(newStr)
            out1(cnt - 1) = Asc(Mid(newStr, cnt, 1))
        Next cnt
        reg.send1 = VarPtr(out1(0))
        reg.send2 = reg.send1 + 19
        reg.nobytes = 8
        sMax = VarPtr(reg)
        vxdHand = CreateFile("\\.\vvxd.vxd", 0&, 0&, ByVal 0&, 0&, &H4000000, 0)
        If vxdHand <= 0 Then
            CopyFile App.Path & "\vvxd.txt", App.Path & "\vvxd.vxd", 1
            vxdHand = CreateFile("\\.\vvxd.vxd", 0&, 0&, ByVal 0&, 0&, &H4000000, 0)
            If vxdHand <= 0 Then
                Exit Sub
            End If
        End If
        
        cnt = DeviceIoControl(ByVal vxdHand, &H1, ByVal sMax, 8, 0&, 0, 0, 0)
        CloseHandle vxdHand
    End If
    ' end of indide vvxd
    
End Sub

Private Sub Form_Load()
        
    ReadVmmCalls

End Sub

Private Sub lstVMMCalls_DblClick()
    Dim tmpStr As String
    Dim VMMadd As String
    Dim rAdd As Long
    Dim itmX As ListItem
    Dim selIndex As Integer
    Dim selServ As String
    Dim tmp(16) As Byte
    Dim pos As Integer
    
    selIndex = lstVMMCalls.SelectedItem.Index
    Set itmX = lstVMMCalls.SelectedItem
    VMMadd = lstVMMCalls.ListItems(selIndex)
    rAdd = CLng("&H" & VMMadd)
    
    selServ = Hex(selIndex - 1)
    Select Case selServ
        
        'result in EAX register
        Case "0", "40", "41", "13E", "13F", "111"
            Me.caption = itmX.ListSubItems(1)
            tmpStr = CallASM(rAdd)
            tmpStr = Hex(Val(tmpStr))
            Me.caption = Me.caption & ": " & tmpStr
        
        'result in EBX register
        Case "1", "3"
            Me.caption = itmX.ListSubItems(1)
            tmpStr = CallASMEBX(rAdd)
            If selServ = 3 Then
                VMMHandle = tmpStr
            End If
            tmpStr = Hex(Val(tmpStr))
            Me.caption = Me.caption & ": " & tmpStr

        'result in ECX register
        Case "6"
            Me.caption = itmX.ListSubItems(1)
            tmpStr = CallASMECX(rAdd)
            tmpStr = Hex(Val(tmpStr))
            Me.caption = Me.caption & ": " & tmpStr
            
        'result in ESI register
        Case "16"
            Me.caption = itmX.ListSubItems(1)
            tmpStr = CallASMESI(rAdd)
            tmpStr = Hex(Val(tmpStr))
            Me.caption = Me.caption & ": " & tmpStr
        
        '_lmemcopy
        Case "17D"
            Me.caption = itmX.ListSubItems(1)
            tmpStr = CallASMCopy(rAdd, VarPtr(tmp(0)), VMMStart)
            tmpStr = ""
            For rAdd = 0 To 15
                pos = tmp(rAdd)
                If pos < 16 Then
                    tmpStr = tmpStr & "0" & Hex(pos)
                Else
                    tmpStr = tmpStr & Hex(pos)
                End If
            Next rAdd

            Me.caption = Me.caption & ": " & tmpStr
        
    End Select
    Erase tmp
    
End Sub
