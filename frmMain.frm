VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10485
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   10485
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkStatus 
      Caption         =   "Messages ON"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   35
      Top             =   6120
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.Timer timerStatus 
      Interval        =   10000
      Left            =   9960
      Top             =   7200
   End
   Begin VB.PictureBox picStatus 
      AutoRedraw      =   -1  'True
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   2475
      TabIndex        =   32
      Top             =   6600
      Width           =   2535
      Begin VB.TextBox txtStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   1815
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   34
         Top             =   0
         Width           =   2415
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   375
      Left            =   -1800
      TabIndex        =   31
      Top             =   6360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   250
   End
   Begin VB.PictureBox picCPU 
      AutoRedraw      =   -1  'True
      Height          =   5595
      Left            =   120
      ScaleHeight     =   5535
      ScaleWidth      =   2355
      TabIndex        =   13
      Top             =   400
      Width           =   2415
      Begin VB.CommandButton cmdOther 
         Caption         =   "REFRESH"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Jump to selected ram block"
         Top             =   5160
         Width           =   885
      End
      Begin VB.Label Label2 
         Caption         =   "Flags: "
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblFlags 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   600
         TabIndex        =   27
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lblCpu 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   10
         Left            =   0
         TabIndex        =   25
         Top             =   480
         Width           =   2295
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   2400
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   2400
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   2400
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label lblCpu 
         Caption         =   "DR7: "
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   24
         Top             =   4800
         Width           =   2295
      End
      Begin VB.Label lblCpu 
         Caption         =   "DR6: "
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   23
         Top             =   4440
         Width           =   2295
      End
      Begin VB.Label lblCpu 
         Caption         =   "DR3: "
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   22
         Top             =   4080
         Width           =   2295
      End
      Begin VB.Label lblCpu 
         Caption         =   "DR2: "
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   21
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label lblCpu 
         Caption         =   "DR1: "
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   20
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label lblCpu 
         Caption         =   "DR0: "
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   19
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Label lblCpu 
         Caption         =   "CR4: "
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   18
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label lblCpu 
         Caption         =   "CR3: "
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   17
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label lblCpu 
         Caption         =   "CR2: "
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   16
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label lblCpu 
         Caption         =   "CR0: "
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   15
         Top             =   1000
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "CPU Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   2520
      TabIndex        =   0
      Top             =   -120
      Width           =   7935
      Begin VB.CommandButton cmdCalls 
         Cancel          =   -1  'True
         Caption         =   "VxD CALLS"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6200
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Disassemble current ram block."
         Top             =   6240
         Width           =   1125
      End
      Begin VB.CommandButton cmdOther 
         Caption         =   "DISASSEMBLE"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Disassemble current ram block."
         Top             =   6240
         Width           =   1125
      End
      Begin VB.TextBox txtJump 
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Top             =   6240
         Width           =   1815
      End
      Begin VB.CommandButton cmdOther 
         Caption         =   "JMP"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3015
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Jump to selected ram block"
         Top             =   6240
         Width           =   405
      End
      Begin VB.CommandButton cmdOther 
         Height          =   285
         Index           =   1
         Left            =   675
         Picture         =   "frmMain.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Display next ram block"
         Top             =   6240
         Width           =   405
      End
      Begin VB.CommandButton cmdOther 
         Height          =   285
         Index           =   0
         Left            =   240
         Picture         =   "frmMain.frx":05D4
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Display previous ram block"
         Top             =   6240
         Width           =   405
      End
      Begin MSComctlLib.ListView lstGDT 
         Height          =   5655
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   9975
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Off."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Base"
            Object.Width           =   1799
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Limit"
            Object.Width           =   1799
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DPL"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "A1"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "A2"
            Object.Width           =   1058
         EndProperty
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   3135
         Left            =   7320
         TabIndex        =   2
         ToolTipText     =   "Select front or back of busines card"
         Top             =   480
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   5530
         _Version        =   393216
         TabOrientation  =   3
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   529
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "GDT"
         TabPicture(0)   =   "frmMain.frx":0766
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).ControlCount=   0
         TabCaption(1)   =   "IDT"
         TabPicture(1)   =   "frmMain.frx":0782
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "LDT"
         TabPicture(2)   =   "frmMain.frx":079E
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         TabCaption(3)   =   "VxD's"
         TabPicture(3)   =   "frmMain.frx":07BA
         Tab(3).ControlEnabled=   0   'False
         Tab(3).ControlCount=   0
         TabCaption(4)   =   "RAM"
         TabPicture(4)   =   "frmMain.frx":07D6
         Tab(4).ControlEnabled=   0   'False
         Tab(4).ControlCount=   0
      End
      Begin MSComctlLib.ListView lstIDT 
         Height          =   5655
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   9975
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Off."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Base"
            Object.Width           =   1799
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Limit"
            Object.Width           =   1799
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DPL"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "A1"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "A2"
            Object.Width           =   1058
         EndProperty
      End
      Begin MSComctlLib.ListView lstLDT 
         Height          =   5655
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   9975
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Off."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Base"
            Object.Width           =   1799
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Limit"
            Object.Width           =   1799
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DPL"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "A1"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "A2"
            Object.Width           =   1058
         EndProperty
      End
      Begin MSComctlLib.ListView lstRam 
         Height          =   5655
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   9975
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   18
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Address"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "00"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "01"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "02"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "03"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "04"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "05"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "06"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "07"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "08"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "09"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "0A"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "0B"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "0C"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "0D"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "0E"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "0F"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "ASCII Representation"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView lstAll 
         Height          =   5655
         Left            =   2640
         TabIndex        =   28
         ToolTipText     =   "Static VxD's as retrieved from VMM "
         Top             =   480
         Width           =   4690
         _ExtentX        =   8281
         _ExtentY        =   9975
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483641
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Address"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Length"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Seg"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DDB"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstStaticVxD 
         Height          =   5655
         Left            =   240
         TabIndex        =   29
         ToolTipText     =   "Static VxD's as listed in the registry"
         Top             =   480
         Width           =   2260
         _ExtentX        =   3995
         _ExtentY        =   9975
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483641
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Static VxD's"
            Object.Width           =   3810
         EndProperty
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   195
         Width           =   7095
      End
   End
   Begin RichTextLib.RichTextBox rt1 
      Height          =   1455
      Left            =   2760
      TabIndex        =   12
      Top             =   6600
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   2566
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":07F2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
Option Explicit

Dim GDTBase(5) As Byte
Dim GDTL As Long
Dim IDTBase(5) As Byte
Dim sendKer(3) As Byte
Dim GDTLimit As Integer
Dim IDTLimit As Integer
Dim copyBytes() As Byte
Dim LDTAddress As String
Dim LDTLimit As String
Dim GDTCapStr As String
Dim IDTCapStr As String
Dim LDTCapStr As String
Dim RAMCapStr As String

Private VxdName As String

Dim RamRead As Long
Const rLimit = &H1000
Const Loadlib = "kernel32.dll"

Private Sub chkStatus_Click()

    If chkStatus.Value = 0 Then
        timerStatus.Enabled = False
        txtStatus.Text = ""
    Else
        timerStatus.Enabled = True
    End If
    
End Sub

Private Sub cmdOther_Click(Index As Integer)
    Dim tmp As Long

    Select Case Index
        Case 0 'back
            Screen.MousePointer = vbHourglass
            If RamRead > &HFFF Then
                RamRead = RamRead - &H1000
                ReadRam RamRead
            ElseIf RamRead < 0 Then
                RamRead = RamRead - &H1000
                ReadRam RamRead
            End If
            
        Case 1 'forward
            Screen.MousePointer = vbHourglass
            RamRead = RamRead + &H1000
            ReadRam RamRead
            
        Case 2 ' jump
            Screen.MousePointer = vbHourglass
            If IsNumeric(txtJump.Text) Then
                RamRead = CLng(txtJump.Text)
                ReadRam RamRead
            Else
                'try hex
                tmp = CLng("&H" & txtJump.Text)
                RamRead = tmp
                ReadRam RamRead
            End If
        Case 9 ' Disaassemble code
            Screen.MousePointer = vbHourglass
            DisAssemble
        Case 10 ' Refresh CPUInfo
            GetCpuInfo
    End Select
    Screen.MousePointer = vbDefault
            
End Sub

Private Sub Form_Load()
    Dim thisAccess As EXPLICIT_ACCESS
    Dim ststName As String
    Dim hThread As Long
    Dim newStr As String
    Dim tmpStr As String
    Dim BaseStr As String
    Dim LimitStr As String
    Dim cnt As Integer
    Dim ret As Long
    Dim oldWidth As Long
    Dim oldHeight As Long
    
    oldWidth = Me.Width
    oldHeight = Me.Height
    Me.Width = Me.Width / 4
    Me.Height = Me.Height / 10 - 40
    Me.Show
    Me.Refresh
    pBar.Left = 0
    pBar.Top = 40
    pBar.Height = Me.Height - 420
    pBar.Width = Me.Width - 120
    
    ASMDis.BaseAddress = 0
    
    With thisAccess
        .grfAccessMode = SECTION_MAP_WRITE
        .grfAccessPermissions = GRANT_ACCESS
        .grfInheritance = NO_INHERITANCE
        .Trustee.MultipleTrusteeOperation = 0
        .Trustee.pMultipleTrustee = NO_MULTIPLE_TRUSTEE
        .Trustee.TrusteeForm = TRUSTEE_IS_NAME
        .Trustee.TrusteeType = TRUSTEE_IS_USER
        .Trustee.ptstrName = ststName
    End With

    'retrieve the current thread and process
    hThread = GetCurrentThread

    'set the new thread priority to "lowest"
    ret = SetThreadPriority(hThread, THREAD_PRIORITY_TIME_CRITICAL)
    If ret = 0 Then
        MsgBox "Unable to set thread, exiting"
        Unload Me
    End If
    ret = RetrieveGDT
    tmpStr = ""
    For cnt = 1 To 0 Step -1
        tmpStr = tmpStr & Hex(GDTBase(cnt))
    Next cnt
    GDTLimit = CLng("&H" & tmpStr)
    GDTLimit = GDTLimit + 1
    LimitStr = Hex(GDTLimit)
    tmpStr = ""
    For cnt = 5 To 2 Step -1
        newStr = ""
        newStr = Hex(GDTBase(cnt))
        If Len(newStr) = 1 Then
            newStr = "0" & newStr
        End If
        tmpStr = tmpStr & newStr
    Next cnt
    BaseStr = tmpStr
    GDTL = CLng("&H" & BaseStr)
    GDTCapStr = "GDT Base: " & BaseStr & "   Limit: " & LimitStr
    ProcessDT lstGDT, BaseStr, LimitStr, GDTCapStr
    pBar.Value = pBar.Value + 25
    tmpStr = ""
    For cnt = 1 To 0 Step -1
        tmpStr = tmpStr & Hex(IDTBase(cnt))
    Next cnt
    IDTLimit = CLng("&H" & tmpStr)
    IDTLimit = IDTLimit + 1
    LimitStr = Hex(IDTLimit)
    tmpStr = ""
    For cnt = 5 To 2 Step -1
        newStr = ""
        newStr = Hex(IDTBase(cnt))
        If Len(newStr) = 1 Then
            newStr = "0" & newStr
        End If
        tmpStr = tmpStr & newStr
    Next cnt
    BaseStr = tmpStr
    IDTCapStr = "IDT Base: " & BaseStr & "   Limit: " & LimitStr
    ProcessDT lstIDT, BaseStr, LimitStr, IDTCapStr
    pBar.Value = pBar.Value + 25
    For cnt = 1 To lstIDT.ListItems.count
        newStr = Left(lstIDT.ListItems.Item(cnt).SubItems(2), 2)
        tmpStr = lstIDT.ListItems.Item(cnt).SubItems(3)
        tmpStr = newStr & Right(tmpStr, 6)
        lstIDT.ListItems.Item(cnt).SubItems(3) = tmpStr
        tmpStr = lstIDT.ListItems.Item(cnt).SubItems(2)
        newStr = "00"
        tmpStr = newStr & Right(tmpStr, 6)
        lstIDT.ListItems.Item(cnt).SubItems(2) = tmpStr
    Next
    BaseStr = LDTAddress
    LimitStr = LDTLimit
    LDTCapStr = "LDT Base: " & BaseStr & "   Limit: " & LimitStr
    ProcessDT lstLDT, BaseStr, LimitStr, LDTCapStr
    pBar.Value = pBar.Value + 25
    ReadRam 0
    pBar.Value = pBar.Value + 25
    lblTitle.caption = GDTCapStr
    SSTab2_Click (0)
    GetCpuInfo
    pBar.Value = pBar.Value = pBar.Value + 5
    ServiceCalls
    pBar.Value = pBar.max
    GetDDBLstLoc
    pBar.Visible = False
    Me.Height = oldHeight
    Me.Width = oldWidth
    
End Sub

'
'This sub retrieves GDT and IDT via machine code
'with processor instructions SGDT and SIDT
'
'and CallWindowProc to execute the machine code
'The LDT is listed as a GDT we simply point to
'that table when it is located in sub ProcessDT
'
'
Private Function RetrieveGDT() As Long
    Dim tmpStr As String
    Dim hStr As String
    Dim callLocationGDT As Long
    Dim callLocationIDT As Long
    Dim callLocationKer As Long
    Dim x As Long
    Dim asmCode(81) As Byte
    Dim cnt As Integer
    Dim counter As Integer
    Dim lib As Long
    Dim add As Long
            
    'pointer in mem to our machine code
    'needed for CallWindowProc
    x = VarPtr(asmCode(0))
    lib = LoadLibrary(ByVal Loadlib)
    add = lib + 5076
    FreeLibrary lib
   
    hStr = Hex(add)
    For cnt = 7 To 1 Step -2
        sendKer(counter) = CLng("&H" & Mid(hStr, cnt, 2))
        counter = counter + 1
    Next cnt
    callLocationKer = VarPtr(sendKer(0))
    
    asmCode(0) = &H55   ' push ebp
    asmCode(1) = &H8B   ' mov ebp, esp
    asmCode(2) = &HEC
    
    asmCode(3) = &H83   ' sub esp, 10
    asmCode(4) = &HEC
    asmCode(5) = &H10
    
    asmCode(6) = &HF    ' SGDT Store Global descriptor table
    asmCode(7) = &H1    '
    asmCode(8) = &H5    '
    
    callLocationGDT = VarPtr(GDTBase(0))
    
    hStr = Hex(callLocationGDT)
    hStr = String(8 - Len(hStr), "0") & hStr
    tmpStr = ""
    For cnt = 8 To 1 Step -2
        tmpStr = tmpStr + Mid(hStr, cnt - 1, 2)
    Next
    
    hStr = Left(tmpStr, 2)
    asmCode(9) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 3, 2)
    asmCode(10) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 5, 2)
    asmCode(11) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 7, 2)
    asmCode(12) = CLng("&H" & hStr)
    
    asmCode(13) = &HF    ' SIDT  Store Global interrupt table
    asmCode(14) = &H1    '
    asmCode(15) = &HD
    
    callLocationIDT = VarPtr(IDTBase(0))
    
    hStr = Hex(callLocationIDT)
    hStr = String(8 - Len(hStr), "0") & hStr
    tmpStr = ""
    For cnt = 8 To 1 Step -2
        tmpStr = tmpStr + Mid(hStr, cnt - 1, 2)
    Next
    
    hStr = Left(tmpStr, 2)
    asmCode(16) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 3, 2)
    asmCode(17) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 5, 2)
    asmCode(18) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 7, 2)
    asmCode(19) = CLng("&H" & hStr)
       
        
    asmCode(20) = &HA1                ' mov eax GDT Table just received
    asmCode(21) = asmCode(9) + 2      ' from above 'Base
    asmCode(22) = asmCode(10)
    asmCode(23) = asmCode(11)
    asmCode(24) = asmCode(12)
    
    asmCode(25) = &H8B                ' mov ecx, eax
    asmCode(26) = &HC8
    
    asmCode(27) = &H81                ' and ecx with &H00000FFF
    asmCode(28) = &HE1
    asmCode(29) = &HFF
    asmCode(30) = &HF
    asmCode(31) = &H0
    asmCode(32) = &H0
    
    asmCode(33) = &H51                ' push ecx
    
    asmCode(34) = &H66
    asmCode(35) = &H3                 ' add GDT limit obtained from above
    asmCode(36) = &HD                 ' to ecx
    asmCode(37) = asmCode(9)
    asmCode(38) = asmCode(10)
    asmCode(39) = asmCode(11)
    asmCode(40) = asmCode(12)
    
    asmCode(41) = &H81          ' add &H00000FFF to ecx
    asmCode(42) = &HC1
    asmCode(43) = &HFF
    asmCode(44) = &HF
    asmCode(45) = &H0
    asmCode(46) = &H0
    
    asmCode(47) = &HC1          ' shr ecx, &H0C
    asmCode(48) = &HE9
    asmCode(49) = &HC
    
    asmCode(50) = &HC1          ' shr eax, &H0C
    asmCode(51) = &HE8
    asmCode(52) = &HC
    
    asmCode(53) = &H50          ' push eax
    
    asmCode(54) = &H68          ' push 20060000
    asmCode(55) = &H0           ' STATIC, USER, and WRITEABLE
    asmCode(56) = &H0
    asmCode(57) = &H6
    asmCode(58) = &H20
    
    asmCode(59) = &H6A          ' push &HFF
    asmCode(60) = &HFF
    
    asmCode(61) = &H51          ' push ecx
    
    asmCode(62) = &H50          ' push eax
    
    asmCode(63) = &H68          ' push PAGE_MODIFY_PERMISSIONS
    asmCode(64) = &HD
    asmCode(65) = &H0
    asmCode(66) = &H1
    asmCode(67) = &H0
    
    asmCode(68) = &HFF          ' Call KERNEL_ORD0001
    asmCode(69) = &H15          ' to enter ring0
    
    hStr = Hex(callLocationKer)
    hStr = String(8 - Len(hStr), "0") & hStr
    tmpStr = ""
    For cnt = 8 To 1 Step -2
        tmpStr = tmpStr + Mid(hStr, cnt - 1, 2)
    Next
    
    hStr = Left(tmpStr, 2)
    asmCode(70) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 3, 2)
    asmCode(71) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 5, 2)
    asmCode(72) = CLng("&H" & hStr)
    hStr = Mid(tmpStr, 7, 2)
    asmCode(73) = CLng("&H" & hStr)
    
    asmCode(74) = &H5A          ' pop edx
    asmCode(75) = &H59          'pop ecx
    
    asmCode(76) = &H8B          ' mov esp, ebp
    asmCode(77) = &HE5
    asmCode(78) = &H5D          ' pop ebp
    
    asmCode(79) = &HC2  ' ret
    asmCode(80) = &H10
    asmCode(81) = &H0   ' just in case stack filled bith garbage
    tmpStr = ""
    
    'execute the machine code and store the tables
    CallWindowProc x, 0, 0, 0, 0
    Erase asmCode
    
End Function

'This sub merely parses the GDT SDT LDT pointed to
'by previous call RetrieveGDT Sub
'and populates the approproiate listview
'
Private Sub ProcessDT(lv As ListView, strBase As String, strLimit As String, caption As String)
    Dim typeStr As String
    Dim at1Str As String
    Dim at2Str As String
    Dim BaseStr As String
    Dim LimitStr As String
    Dim dplStr As String
    Dim newStr As String
    Dim cnt As Integer
    Dim gBase As Long
    Dim gLimit As Integer
    Dim counter As Integer
    Dim x As Long
    Dim tmpBase As String
    Dim tmpLimit As String
    Dim tmpA As Integer
    Dim tmpD As Integer
    
    lblTitle.caption = caption
    
    gBase = CLng("&H" & strBase)
    gLimit = CLng("&H" & strLimit)
    ReDim copyBytes(gLimit)
    x = VarPtr(copyBytes(0))
    MoveMemory ByVal x, ByVal gBase, gLimit
    counter = 1
    newStr = ""
    For cnt = 0 To gLimit - 1 Step 8
        tmpLimit = ""
        tmpBase = ""
        newStr = Hex(copyBytes(cnt + 1))
        If Len(newStr) = 1 Then
            newStr = "0" & newStr
        End If
        tmpLimit = tmpLimit & newStr
        newStr = Hex(copyBytes(cnt))
        If Len(newStr) = 1 Then
            newStr = "0" & newStr
        End If
        tmpLimit = tmpLimit + newStr
        
        newStr = Hex(copyBytes(cnt + 3))
        If Len(newStr) = 1 Then
            newStr = "0" & newStr
        End If
        tmpBase = tmpBase & newStr
        newStr = Hex(copyBytes(cnt + 2))
        If Len(newStr) = 1 Then
            newStr = "0" & newStr
        End If
        tmpBase = tmpBase & newStr
        newStr = Hex(cnt)
        Select Case tmpLimit
            Case "0000"
                lv.ListItems.add , , String(4 - Len(newStr), "0") & newStr
                lv.ListItems.Item(counter).SubItems(1) = "Reserved"
                lv.ListItems.Item(counter).SubItems(2) = String(8 - Len(tmpBase), "0") & tmpBase
                lv.ListItems.Item(counter).SubItems(3) = String(8 - Len(tmpLimit), "0") & tmpLimit
                lv.ListItems.Item(counter).SubItems(4) = 0
                lv.ListItems.Item(counter).SubItems(5) = "NP"

            Case Else
                BaseStr = tmpBase
                LimitStr = tmpLimit
                tmpA = copyBytes(cnt + 5)
                tmpD = Int(tmpA / 32) And 3
                x = tmpA And 128
                If tmpD = 0 Then
                    dplStr = "0"
                    at1Str = "P"
                Else
                    at1Str = "NP"
                     
                End If
                If tmpA And 16 Then
                    tmpD = copyBytes(cnt + 6)
                    tmpA = tmpA And 8
                    tmpD = tmpD And &H40
                    tmpA = tmpA / 8
                    tmpD = tmpD / 32
                    tmpA = tmpA Or tmpD
                    tmpA = tmpA Or &H10
                End If
                
                tmpA = tmpA And &H1F
                at1Str = "P"
                Select Case tmpA
                    Case 0
                        typeStr = "Reserved"
                        BaseStr = tmpBase
                        LimitStr = tmpLimit
                        at1Str = "NP"
                    
                    Case 1      ' TSS16
                        BaseStr = BaseCalc(cnt, LimitStr)
                        LimitStr = String(8 - Len(LimitStr), "0") & LimitStr
                        BaseStr = String(8 - Len(BaseStr), "0") & BaseStr
                        typeStr = "TSS16"
                        at2Str = ""
                    
                    Case 2
                        BaseStr = BaseCalc(cnt, LimitStr)
                        LimitStr = String(8 - Len(LimitStr), "0") & LimitStr
                        BaseStr = String(8 - Len(BaseStr), "0") & BaseStr
                        typeStr = "LDT"
                        at2Str = ""
                        LDTAddress = BaseStr
                        LDTLimit = LimitStr
                        
                    Case 3      ' TSS32
                        BaseStr = BaseCalc(cnt, LimitStr)
                        LimitStr = String(8 - Len(LimitStr), "0") & LimitStr
                        BaseStr = String(8 - Len(BaseStr), "0") & BaseStr
                        typeStr = "TSS32"
                        at2Str = ""
                                            
                    Case &H4   ' CallG16
                        BaseStr = BaseCalc(cnt, LimitStr)
                        LimitStr = String(8 - Len(LimitStr), "0") & LimitStr
                        BaseStr = String(8 - Len(BaseStr), "0") & BaseStr
                        at2Str = ""
                        typeStr = "CallG16"
                    
                    Case &H5   ' TaskG
                        BaseStr = BaseCalc(cnt, LimitStr)
                        LimitStr = String(8 - Len(LimitStr), "0") & LimitStr
                        BaseStr = String(8 - Len(BaseStr), "0") & BaseStr
                        at2Str = ""
                        typeStr = "TaskG"
                        
                    Case &H6   ' IntG16
                        BaseStr = BaseCalc(cnt, LimitStr)
                        LimitStr = String(8 - Len(LimitStr), "0") & LimitStr
                        BaseStr = String(8 - Len(BaseStr), "0") & BaseStr
                        at2Str = ""
                        typeStr = "IntG16"
                        
                    Case &H7   ' TrapG16
                        BaseStr = BaseCalc(cnt, LimitStr)
                        LimitStr = String(8 - Len(LimitStr), "0") & LimitStr
                        BaseStr = String(8 - Len(BaseStr), "0") & BaseStr
                        at2Str = ""
                        typeStr = "TrapG16"
                        
                    Case &H8, &HA, &HD   ' Reserved
                        BaseStr = BaseCalc(cnt, LimitStr)
                        LimitStr = String(8 - Len(LimitStr), "0") & LimitStr
                        BaseStr = String(8 - Len(BaseStr), "0") & BaseStr
                        at2Str = ""
                        typeStr = "Reserved"
                        
                    Case &H9   ' RSS32
                        BaseStr = BaseCalc(cnt, LimitStr)
                        LimitStr = String(8 - Len(LimitStr), "0") & LimitStr
                        BaseStr = String(8 - Len(BaseStr), "0") & BaseStr
                        at2Str = ""
                        typeStr = "RSS32"
                        
                    Case &HB   ' TSS32
                        BaseStr = BaseCalc(cnt, LimitStr)
                        LimitStr = String(8 - Len(LimitStr), "0") & LimitStr
                        BaseStr = String(8 - Len(BaseStr), "0") & BaseStr
                        at2Str = "A"
                        typeStr = "TSS32"
                        
                    Case &HC  ' CallG32
                        BaseStr = BaseCalc(cnt, LimitStr)
                        LimitStr = String(8 - Len(LimitStr), "0") & LimitStr
                        BaseStr = String(8 - Len(BaseStr), "0") & BaseStr
                        at2Str = ""
                        typeStr = "CallG32"
                        
                    Case &HE   ' IntG32
                        BaseStr = BaseCalc(cnt, LimitStr)
                        LimitStr = String(8 - Len(LimitStr), "0") & LimitStr
                        BaseStr = String(8 - Len(BaseStr), "0") & BaseStr
                        at2Str = ""
                        typeStr = "IntG32"
                        
                    Case &HF   ' TrapG32
                        BaseStr = BaseCalc(cnt, LimitStr)
                        LimitStr = String(8 - Len(LimitStr), "0") & LimitStr
                        BaseStr = String(8 - Len(BaseStr), "0") & BaseStr
                        at2Str = ""
                        typeStr = "TrapG32"
                        
                    Case &H10  ' Data 16
                        BaseStr = BaseCalc(cnt, LimitStr)
                        LimitStr = String(8 - Len(LimitStr), "0") & LimitStr
                        BaseStr = String(8 - Len(BaseStr), "0") & BaseStr
                        typeStr = "Data16"
                        tmpD = copyBytes(cnt + 5)
                        If tmpD And 2 = 0 Then
                            at2Str = "R0"
                        ElseIf tmpD And 4 = 0 Then
                            at2Str = "RW"
                        Else
                            at2Str = ""
                        End If
                        
                    Case &H11   ' Code16
                        BaseStr = BaseCalc(cnt, LimitStr)
                        LimitStr = String(8 - Len(LimitStr), "0") & LimitStr
                        BaseStr = String(8 - Len(BaseStr), "0") & BaseStr
                        tmpD = copyBytes(cnt + 5)
                        If tmpD And 2 = 0 Then
                            at2Str = "E0"
                        ElseIf tmpD And 4 = 0 Then
                            at2Str = "RE"
                        Else
                            at2Str = ""
                        End If
                        typeStr = "Code16"
                    
                    Case &H12  ' Data32
                        BaseStr = BaseCalc(cnt, LimitStr)
                        LimitStr = String(8 - Len(LimitStr), "0") & LimitStr
                        BaseStr = String(8 - Len(BaseStr), "0") & BaseStr
                        typeStr = "Data32"
                        tmpD = copyBytes(cnt + 5)
                        If tmpD And 2 = 0 Then
                            at2Str = "R0"
                        ElseIf tmpD And 4 = 0 Then
                            at2Str = "RW"
                        Else
                            at2Str = ""
                        End If
                        
                    Case &H13   ' Code32
                        BaseStr = BaseCalc(cnt, LimitStr)
                        LimitStr = String(8 - Len(LimitStr), "0") & LimitStr
                        BaseStr = String(8 - Len(BaseStr), "0") & BaseStr
                        typeStr = "Code32"
                        tmpD = copyBytes(cnt + 5)
                        If tmpD And 2 = 0 Then
                            at2Str = "E0"
                        ElseIf tmpD And 4 = 0 Then
                            at2Str = "RE"
                        Else
                            at2Str = "C"
                        End If
                        
                End Select
                
                lv.ListItems.add , , String(4 - Len(newStr), "0") & newStr
                lv.ListItems.Item(counter).SubItems(1) = typeStr
                If Len(BaseStr) < 8 Then
                    lv.ListItems.Item(counter).SubItems(2) = String(8 - Len(tmpBase), "0") & BaseStr
                Else
                    lv.ListItems.Item(counter).SubItems(2) = BaseStr
                End If
                If Len(LimitStr) < 8 Then
                    lv.ListItems.Item(counter).SubItems(3) = String(8 - Len(tmpLimit), "0") & LimitStr
                Else
                    lv.ListItems.Item(counter).SubItems(3) = LimitStr
                End If
                lv.ListItems.Item(counter).SubItems(4) = dplStr
                lv.ListItems.Item(counter).SubItems(5) = at1Str
                
                lv.ListItems.Item(counter).SubItems(6) = at2Str
        End Select
        counter = counter + 1
        pBar.Value = pBar.Value + 0.025
    Next cnt
    
End Sub

'
'This routine is part of the parsing of GDT LDT IDT
'and merely changes the Base and limits respectively
'
Private Function BaseCalc(cnt As Integer, ByRef limit As String) As String
    Dim low As Integer
    Dim high As Integer
    Dim tmpD As Integer
    Dim tmpC As Long
    Dim newStr As String
    Dim tmpStr As String
    Dim tmpCalc As String
    
    high = copyBytes(cnt + 7)
    low = copyBytes(cnt + 4)
    newStr = Hex(high)
    tmpStr = String(4 - Len(newStr), "0") & newStr
    newStr = Hex(low)
    newStr = String(2 - Len(newStr), "0") & newStr
    tmpStr = String(8 - (Len(newStr) + Len(tmpStr)), "0") & tmpStr & newStr
    tmpStr = Right(tmpStr, 4)
    high = copyBytes(cnt + 3)
    tmpCalc = tmpStr
    
    If high < 16 Then
        tmpStr = "0" & Hex(high)
    Else
        tmpStr = Hex(high)
    End If
    
    low = copyBytes(cnt + 2)
    If low < 16 Then
        newStr = "0" & Hex(low)
    Else
        newStr = Hex(low)
    End If
    tmpStr = tmpStr & newStr
    tmpCalc = tmpCalc & tmpStr
    
    tmpC = copyBytes(cnt + 6)
    tmpD = tmpC
    tmpC = tmpC And &HF
    low = copyBytes(cnt)
    high = copyBytes(cnt + 1)
    newStr = Hex(low)
    tmpStr = String(2 - Len(newStr), "0") & newStr
    newStr = Hex(high)
    newStr = String(2 - Len(newStr), "0") & newStr
    tmpStr = String(4 - (Len(newStr) + Len(tmpStr)), "0") & newStr & tmpStr
    
    If tmpD And 128 <> 0 Then
        tmpC = CLng("&H" & tmpStr)
    End If
    limit = tmpStr
    BaseCalc = tmpCalc
    
End Function

Private Sub SSTab2_Click(PreviousTab As Integer)
    
    lstGDT.Visible = False
    lstIDT.Visible = False
    lstLDT.Visible = False
    lstRam.Visible = False
    lstAll.Visible = False
    lstStaticVxD.Visible = False
    cmdOther(0).Visible = False
    cmdOther(1).Visible = False
    cmdOther(2).Visible = False
    cmdOther(9).Visible = False
    cmdCalls.Visible = False
    txtJump.Visible = False
    rt1.Visible = False
    lblTitle.caption = ""
    
    Select Case SSTab2.Tab
        
        Case 0  ' GDT
            lstGDT.Visible = True
            lblTitle.caption = GDTCapStr
        
        Case 1  ' IDT
            lstIDT.Visible = True
            lblTitle.caption = IDTCapStr
        
        Case 2  ' LDT
            lstLDT.Visible = True
            lblTitle = LDTCapStr
        
        Case 3  ' VxD's
            lstAll.Visible = True
            lstStaticVxD.Visible = True
            cmdCalls.Visible = True
            
        Case 4  ' RAM
            lstRam.Visible = True
            cmdOther(0).Visible = True
            cmdOther(1).Visible = True
            cmdOther(2).Visible = True
            cmdOther(9).Visible = True
            rt1.Visible = True
            txtJump.Visible = True
            lblTitle.caption = RAMCapStr
            
    End Select

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim f As Form
    Dim c As Control
    Dim o As Object
    Dim a As Variant
    
    'destroy any left over objects
    For Each f In Forms
        
        For Each c In f
            Set c = Nothing
        Next c
        
        For Each o In f
            Set o = Nothing
        Next o
        
        For Each a In f
            Set a = Nothing
        Next a
        
        If f.hwnd <> Me.hwnd Then
            Unload f
            Set f = Nothing
        End If
    Next f
    
    Set ASMDis = Nothing
    Erase copyBytes
    Erase GDTBase
    Erase sendKer
    Erase IDTBase
    Erase DataArr
    Erase SCall
    Erase StatusStr
    Erase VxDs
    Erase DynVxDs
    Erase Value
    
    rt1 = ""
    
    Unload Me
    
End Sub

'Cut n dry this routine reads RAM
'pointed to by variable add
'
'the const rlimit ensures
'the read is always 4096 bytes
'
'the routine also populates the
'listview with the parsed info
'
Private Sub ReadRam(add As Long)
    Dim readStr As String
    Dim cnt As Integer
    Dim tmpStr As String
    Dim hold As Integer
    Dim addStr As String
    Dim counter As Integer
    Dim asciiStr As String
    
    lstRam.ListItems.Clear
    counter = 1
    readStr = ReadMem(add, rLimit)
    'to check for read error here
    Me.AutoRedraw = False
    tmpStr = ""
    For cnt = 1 To rLimit - 1 Step 16
        hold = Asc(Mid(readStr, cnt, 1))
        If hold < 16 Then
            tmpStr = "0"
        End If
        tmpStr = tmpStr & Hex(hold) '1st Column
        addStr = Hex(add + cnt - 1)
        If Len(addStr) < 8 Then
            addStr = String(8 - Len(addStr), "0") & addStr
        End If
        lstRam.ListItems.add , , addStr
        lstRam.ListItems.Item(counter).SubItems(1) = tmpStr
        If Val(hold) = 0 Then
            asciiStr = "."
        Else
            asciiStr = Chr(hold)
        End If
        tmpStr = ""
        hold = Asc(Mid(readStr, cnt + 1, 1))
        If hold < 16 Then
            tmpStr = "0"
        End If
        tmpStr = tmpStr & Hex(hold) '2nd Column
        lstRam.ListItems.Item(counter).SubItems(2) = tmpStr
        If Val(hold) = 0 Then
            asciiStr = asciiStr & "."
        Else
            asciiStr = asciiStr & Chr(hold)
        End If
        tmpStr = ""
        hold = Asc(Mid(readStr, cnt + 2, 1))
        If hold < 16 Then
            tmpStr = "0"
        End If
        tmpStr = tmpStr & Hex(hold) '3nd Column
        lstRam.ListItems.Item(counter).SubItems(3) = tmpStr
        If hold = 0 Then
            asciiStr = asciiStr & "."
        Else
            asciiStr = asciiStr & Chr(hold)
        End If
        tmpStr = ""
        hold = Asc(Mid(readStr, cnt + 3, 1))
        If hold < 16 Then
            tmpStr = "0"
        End If
        tmpStr = tmpStr & Hex(hold) '4th Column
        lstRam.ListItems.Item(counter).SubItems(4) = tmpStr
        If hold = 0 Then
            asciiStr = asciiStr & "."
        Else
            asciiStr = asciiStr & Chr(hold)
        End If
        tmpStr = ""
        hold = Asc(Mid(readStr, cnt + 4, 1))
        If hold < 16 Then
            tmpStr = "0"
        End If
        tmpStr = tmpStr & Hex(hold) '5th Column
        lstRam.ListItems.Item(counter).SubItems(5) = tmpStr
        If hold = 0 Then
            asciiStr = asciiStr & "."
        Else
            asciiStr = asciiStr & Chr(hold)
        End If
        tmpStr = ""
        hold = Asc(Mid(readStr, cnt + 5, 1))
        If hold < 16 Then
            tmpStr = "0"
        End If
        tmpStr = tmpStr & Hex(hold) '6th Column
        lstRam.ListItems.Item(counter).SubItems(6) = tmpStr
        If hold = 0 Then
            asciiStr = asciiStr & "."
        Else
            asciiStr = asciiStr & Chr(hold)
        End If
        tmpStr = ""
        hold = Asc(Mid(readStr, cnt + 6, 1))
        If hold < 16 Then
            tmpStr = "0"
        End If
        tmpStr = tmpStr & Hex(hold) '7th Column
        lstRam.ListItems.Item(counter).SubItems(7) = tmpStr
        If hold = 0 Then
            asciiStr = asciiStr & "."
        Else
            asciiStr = asciiStr & Chr(hold)
        End If
        tmpStr = ""
        hold = Asc(Mid(readStr, cnt + 7, 1))
        If hold < 16 Then
            tmpStr = "0"
        End If
        tmpStr = tmpStr & Hex(hold) '8th Column
        lstRam.ListItems.Item(counter).SubItems(8) = tmpStr
        If hold = 0 Then
            asciiStr = asciiStr & "."
        Else
            asciiStr = asciiStr & Chr(hold)
        End If
        tmpStr = ""
        hold = Asc(Mid(readStr, cnt + 8, 1))
        If hold < 16 Then
            tmpStr = "0"
        End If
        tmpStr = tmpStr & Hex(hold) '9th Column
        lstRam.ListItems.Item(counter).SubItems(9) = tmpStr
        If hold = 0 Then
            asciiStr = asciiStr & "."
        Else
            asciiStr = asciiStr & Chr(hold)
        End If
        tmpStr = ""
        hold = Asc(Mid(readStr, cnt + 9, 1))
        If hold < 16 Then
            tmpStr = "0"
        End If
        tmpStr = tmpStr & Hex(hold) '10th Column
        lstRam.ListItems.Item(counter).SubItems(10) = tmpStr
        If hold = 0 Then
            asciiStr = asciiStr & "."
        Else
            asciiStr = asciiStr & Chr(hold)
        End If
        tmpStr = ""
        hold = Asc(Mid(readStr, cnt + 10, 1))
        If hold < 16 Then
            tmpStr = "0"
        End If
        tmpStr = tmpStr & Hex(hold) '11th Column
        lstRam.ListItems.Item(counter).SubItems(11) = tmpStr
        If hold = 0 Then
            asciiStr = asciiStr & "."
        Else
            asciiStr = asciiStr & Chr(hold)
        End If
        tmpStr = ""
        hold = Asc(Mid(readStr, cnt + 11, 1))
        If hold < 16 Then
            tmpStr = "0"
        End If
        tmpStr = tmpStr & Hex(hold) '12thColumn
        lstRam.ListItems.Item(counter).SubItems(12) = tmpStr
        If hold = 0 Then
            asciiStr = asciiStr & "."
        Else
            asciiStr = asciiStr & Chr(hold)
        End If
        tmpStr = ""
        hold = Asc(Mid(readStr, cnt + 12, 1))
        If hold < 16 Then
            tmpStr = "0"
        End If
        tmpStr = tmpStr & Hex(hold) '13th Column
        lstRam.ListItems.Item(counter).SubItems(13) = tmpStr
        If hold = 0 Then
            asciiStr = asciiStr & "."
        Else
            asciiStr = asciiStr & Chr(hold)
        End If
        tmpStr = ""
        hold = Asc(Mid(readStr, cnt + 13, 1))
        If hold < 16 Then
            tmpStr = "0"
        End If
        tmpStr = tmpStr & Hex(hold) '14th Column
        lstRam.ListItems.Item(counter).SubItems(14) = tmpStr
        If hold = 0 Then
            asciiStr = asciiStr & "."
        Else
            asciiStr = asciiStr & Chr(hold)
        End If
        tmpStr = ""
        hold = Asc(Mid(readStr, cnt + 14, 1))
        If hold < 16 Then
            tmpStr = "0"
        End If
        tmpStr = tmpStr & Hex(hold) '15th Column
        lstRam.ListItems.Item(counter).SubItems(15) = tmpStr
        If hold = 0 Then
            asciiStr = asciiStr & "."
        Else
            asciiStr = asciiStr & Chr(hold)
        End If
        tmpStr = ""
        hold = Asc(Mid(readStr, cnt + 15, 1))
        If hold < 16 Then
            tmpStr = "0"
        End If
        tmpStr = tmpStr & Hex(hold) '16th Column
        lstRam.ListItems.Item(counter).SubItems(16) = tmpStr
        If hold = 0 Then
            asciiStr = asciiStr & "."
        Else
            asciiStr = asciiStr & Chr(hold)
        End If
        tmpStr = ""
        lstRam.ListItems.Item(counter).SubItems(17) = asciiStr
        asciiStr = ""
        counter = counter + 1
    Next cnt
    
    addStr = Hex(add)
    tmpStr = Hex(add + rLimit)
    If Len(addStr) < 8 Then
        addStr = String(8 - Len(addStr), "0") & addStr
    End If
    If Len(tmpStr) < 8 Then
        tmpStr = String(8 - Len(tmpStr), "0") & tmpStr
    End If
    lblTitle.caption = "Ram Memory from " & addStr & "  to " & tmpStr
    RAMCapStr = lblTitle.caption
    Me.AutoRedraw = True
    txtJump.Text = addStr
    
End Sub

'
'
'Displays pseudo mnemonic disassembly of the
'present RAM page
'the class module asmdec.cls and moduleasm.bas
''*****W32 OPCODE DISASSEMBLER WRITTEN BY VANJA FUCKAR EMAIL:INGA@VIP.HR
'
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=41582&lngWId=1
'
'so credits to him accordingly
'
'
Private Sub DisAssemble()
    Dim ARD As Long
    Dim ptr As Long
    Dim Forward As Byte 'Next Instruction!
    Dim cnt As Long
    Dim u As Long
    Dim DATAS() As String
    
    ArrayDescriptor ARD, DataArr, 4
    
    If ARD = 0 Then
        MsgBox "RAM is not loaded yet!", vbCritical, "Error"
        Exit Sub
    End If
    ptr = RamRead
    ASMDis.BaseAddress = ptr
    ReDim DATAS(rLimit)
    cnt = 0
    Do
        DATAS(u) = ASMDis.DisAssemble(DataArr, cnt, Forward, 1, 0) & vbCrLf
        u = u + 1
        cnt = cnt + Forward
    Loop While cnt < rLimit
    ReDim Preserve DATAS(u - 1)
    rt1.Locked = True
    rt1 = Join(DATAS, "")
    Erase DATAS
    Screen.MousePointer = vbDefault

End Sub

Private Sub GetCpuInfo()
    Dim cpuArr(&HFF) As Byte
    Dim cnt As Integer
    Dim tmpStr As String
    Dim hold As Integer
    Dim counter As Integer
    Dim flags As Long
    Dim tmpFlags As String
    
    ASMGetCpu cpuArr(), GDTL, GDTLimit
    counter = 0
    tmpStr = ""
    tmpFlags = ""
    For cnt = 0 To 37 Step 4
        hold = cpuArr(cnt + 3)
        If hold < 16 Then
            tmpStr = "0"
        End If
        tmpStr = tmpStr & Hex(hold)
         
        hold = cpuArr(cnt + 2)
        If hold < 16 Then
            tmpStr = tmpStr & "0"
        End If
        tmpStr = tmpStr & Hex(hold)
        
        hold = cpuArr(cnt + 1)
        If hold < 16 Then
            tmpStr = tmpStr & "0"
        End If
        tmpStr = tmpStr & Hex(hold)
        
        hold = cpuArr(cnt)
        If hold < 16 Then
            tmpStr = tmpStr & "0"
        End If
        tmpStr = tmpStr & Hex(hold)
        Select Case counter
            Case 0
                lblCpu(counter).caption = "CR0: " & tmpStr
                flags = CLng("&H" & tmpStr)
                flags = flags And &H7F
                If flags And 1 Then
                    tmpFlags = tmpFlags & "PE "
                End If
                If flags And 2 Then
                    tmpFlags = tmpFlags & "MP "
                End If
                If flags And 4 Then
                    tmpFlags = tmpFlags & "EM "
                End If
                If flags And 8 Then
                    tmpFlags = tmpFlags And "TS "
                End If
                If flags And 16 Then
                    tmpFlags = tmpFlags & "R "
                End If
                lblFlags.caption = tmpFlags
            Case 1
                lblCpu(counter).caption = "CR2: " & tmpStr
            Case 2
                lblCpu(counter).caption = "CR3: " & tmpStr
            Case 3
                lblCpu(counter).caption = "CR4: " & tmpStr
            Case 4
                lblCpu(counter).caption = "DR0: " & tmpStr
            Case 5
                lblCpu(counter).caption = "DR1: " & tmpStr
            Case 6
                lblCpu(counter).caption = "DR2: " & tmpStr
            Case 7
                lblCpu(counter).caption = "DR3: " & tmpStr
            Case 8
                lblCpu(counter).caption = "DR6: " & tmpStr
            Case 9
                lblCpu(counter).caption = "DR7: " & tmpStr
        End Select
        tmpStr = ""
        counter = counter + 1
    Next cnt
    tmpStr = ""
    counter = counter * 4
    For cnt = counter To &HFF
        If cpuArr(cnt) = 0 Then
            Exit For
        End If
        hold = cpuArr(cnt)
        If hold < 16 Then
            tmpStr = tmpStr & "0"
        End If
        tmpStr = tmpStr & Chr(hold)
    
    Next cnt
    lblCpu(10).caption = tmpStr
    Erase cpuArr
    
End Sub

Private Sub GetDDBLstLoc()
    Dim cpuArr(&HFF) As Byte
    Dim vxdListLocation As Long
    Dim parseStr As String
    Dim intLoop As Long
    Dim counter As Integer
    Dim nextByte As Long
    Dim setUp As Boolean
    Dim Seg As Integer
    Dim sMax As Long
    Dim tmpStr As String
    
    AddToStaticList
    vxdListLocation = ASMGetDDBListLocation(cpuArr(), GDTL, GDTLimit)
        
    sMax = &H7FFF
    tmpStr = ReadByte(vxdListLocation, "", sMax)
    'parse it out
    counter = Asc(Mid(tmpStr, 5))
    nextByte = 14
    setUp = True
    intLoop = 1
    Seg = 1
    For intLoop = intLoop To sMax
        parseStr = Mid(tmpStr, intLoop, nextByte)
        ReadAndUpdate parseStr, setUp, Seg
        setUp = False
        intLoop = intLoop + nextByte - 1
        nextByte = 9
        counter = counter - 1
        Seg = Seg + 1
        If counter <= 0 Then
            setUp = True
            nextByte = 14
            Seg = 1
            counter = Asc(Mid(tmpStr, intLoop + 5, 1))
            If counter = 0 Then
                Exit For
            End If
        End If
    Next intLoop
    
    frmVMMCalls.ReadVmmCalls
    
    Erase cpuArr
    
End Sub

Private Sub AddToStaticList()
    Dim cnt As Integer
    Dim itmX As ListItem
        
    GetVxDStatic
    For cnt = 1 To UBound(VxDs)
        Set itmX = lstStaticVxD.ListItems.add(, , VxDs(cnt))
    Next cnt
    
End Sub

Private Sub ReadAndUpdate(tmpVxD As String, setVxd As Boolean, Seg As Integer)
    Dim pos As Integer
    Dim newStr As String
    Dim tmpStr As String
    Dim rAdd As Long
    Dim itmX As ListItem
    Dim addStr As String
    Dim VMMadd As Long
    Dim vmsa As String
    Dim vmPos As Integer
    
    newStr = ""

    pos = Asc(Mid(tmpVxD, 4, 1))
    If pos < 16 Then
        newStr = newStr & "0" & Hex(pos)
    Else
        newStr = newStr & Hex(pos)
    End If
    pos = Asc(Mid(tmpVxD, 3, 1))
    If pos < 16 Then
        newStr = newStr & "0" & Hex(pos)
    Else
        newStr = newStr & Hex(pos)
    End If
    pos = Asc(Mid(tmpVxD, 2, 1))
    If pos < 16 Then
        newStr = newStr & "0" & Hex(pos)
    Else
        newStr = newStr & Hex(pos)
    End If
    pos = Asc(Mid(tmpVxD, 1, 1))
    If pos < 16 Then
        newStr = newStr & "0" & Hex(pos)
    Else
        newStr = newStr & Hex(pos)
    End If

    'add DDB
    If setVxd Then
        rAdd = CLng("&H" & newStr)
        VMMadd = rAdd + &H30
        addStr = newStr
        'GET NAME
        tmpStr = ReadByte(rAdd, "", 24)
        newStr = Mid(tmpStr, 13, 8)
        VxdName = Trim(newStr)
        Set itmX = lstAll.ListItems.add(, , VxdName)
        itmX.SubItems(4) = addStr
    Else
        Set itmX = lstAll.ListItems.add(, , VxdName)
        'Address
        itmX.SubItems(1) = newStr
    End If
    
    newStr = ""
    If setVxd Then
        pos = Asc(Mid(tmpVxD, 9, 1))
        If pos < 16 Then
            newStr = newStr & "0" & Hex(pos)
        Else
            newStr = newStr & Hex(pos)
        End If
        pos = Asc(Mid(tmpVxD, 8, 1))
        If pos < 16 Then
            newStr = newStr & "0" & Hex(pos)
        Else
            newStr = newStr & Hex(pos)
        End If
        pos = Asc(Mid(tmpVxD, 7, 1))
        If pos < 16 Then
            newStr = newStr & "0" & Hex(pos)
        Else
            newStr = newStr & Hex(pos)
        End If
            pos = Asc(Mid(tmpVxD, 6, 1))
        If pos < 16 Then
            newStr = newStr & "0" & Hex(pos)
        Else
            newStr = newStr & Hex(pos)
        End If
        'Address
        itmX.SubItems(1) = newStr
        If Seg = 1 And VMMStart = 0 Then
            VMMStart = CLng("&H" & newStr)
            tmpStr = ReadByte(VMMadd, "", 4)
            vmsa = ""
            vmPos = Asc(Mid(tmpStr, 4, 1))
            If vmPos < 16 Then
                vmsa = "0"
            End If
            vmsa = vmsa & Hex(vmPos)
            vmPos = Asc(Mid(tmpStr, 3, 1))
            If vmPos < 16 Then
                vmsa = vmsa & "0"
            End If
            vmsa = vmsa & Hex(vmPos)
            vmPos = Asc(Mid(tmpStr, 2, 1))
            If vmPos < 16 Then
                vmsa = vmsa & "0"
            End If
            vmsa = vmsa & Hex(vmPos)
            vmPos = Asc(Mid(tmpStr, 1, 1))
            If vmPos < 16 Then
                vmsa = vmsa & "0"
            End If
            vmsa = vmsa & Hex(vmPos)
            VMMadd = CLng("&H" & vmsa)
'            tmpStr = ReadByte(VMMadd, "", 4)
'            vmsa = ""
'            vmPos = Asc(Mid(tmpStr, 4, 1))
'            If vmPos < 16 Then
'                vmsa = "0"
'            End If
'            vmsa = vmsa & Hex(vmPos)
'            vmPos = Asc(Mid(tmpStr, 3, 1))
'            If vmPos < 16 Then
'                vmsa = vmsa & "0"
'            End If
'            vmsa = vmsa & Hex(vmPos)
'            vmPos = Asc(Mid(tmpStr, 2, 1))
'            If vmPos < 16 Then
'                vmsa = vmsa & "0"
'            End If
'            vmsa = vmsa & Hex(vmPos)
'            vmPos = Asc(Mid(tmpStr, 1, 1))
'            If vmPos < 16 Then
'                vmsa = vmsa & "0"
'            End If
'            vmsa = vmsa & Hex(vmPos)
            VMMadd = CLng("&H" & vmsa)
            VMMServiceCalls = VMMadd
        End If
    Else
        pos = Asc(Mid(tmpVxD, 6, 1))
        If pos < 16 Then
            newStr = newStr & "0" & Hex(pos)
        Else
            newStr = newStr & Hex(pos)
        End If
            pos = Asc(Mid(tmpVxD, 5, 1))
        If pos < 16 Then
            newStr = newStr & "0" & Hex(pos)
        Else
            newStr = newStr & Hex(pos)
        End If
        'Length
        itmX.SubItems(2) = newStr
    End If
    newStr = ""
    If setVxd Then
        pos = Asc(Mid(tmpVxD, 12, 1))
        If pos < 16 Then
            newStr = newStr & "0" & Hex(pos)
        Else
            newStr = newStr & Hex(pos)
        End If
        pos = Asc(Mid(tmpVxD, 11, 1))
        If pos < 16 Then
            newStr = newStr & "0" & Hex(pos)
        Else
            newStr = newStr & Hex(pos)
        End If
        pos = Asc(Mid(tmpVxD, 10, 1))
        If pos < 16 Then
            newStr = newStr & "0" & Hex(pos)
        Else
            newStr = newStr & Hex(pos)
        End If
        'Length
        itmX.SubItems(2) = newStr
    
        newStr = ""
        pos = Asc(Mid(tmpVxD, 13, 1))
        If pos < 16 Then
            newStr = newStr & "0" & Hex(pos)
        Else
            newStr = newStr & Hex(pos)
        End If
        pos = Asc(Mid(tmpVxD, 14, 1))
        If pos < 16 Then
            newStr = newStr & "0" & Hex(pos)
        Else
            newStr = newStr & Hex(pos)
        End If
    End If
        
    'Seg
    itmX.SubItems(3) = Hex(Seg)
    
End Sub

Private Sub cmdCalls_Click()

    frmVMMCalls.Show vbModal
    
End Sub

Private Sub timerStatus_Timer()
    Dim cnt As Integer
    Dim col As Long

    Randomize
    cnt = Int((Rnd * 15) + 1)
    txtStatus.Text = StatusStr(cnt)
    col = CLng((Rnd * 8) + 1)
    Select Case col
        Case 0
            col = vbBlack
        Case 1
            col = vbWhite
        Case 2
            col = vbYellow
        Case 3
            col = vbRed
        Case 4
            col = vbCyan
        Case 5
            col = vbBlue
        Case 6
            col = vbMagenta
        Case Else
            col = vbGreen
    End Select
    txtStatus.ForeColor = vbBlue 'col
    
End Sub
