VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{7FDF243A-2E06-4F93-989D-6C9CC526FFC5}#10.0#0"; "HoverButton.ocx"
Begin VB.Form frmMain2 
   Caption         =   "infiniti"
   ClientHeight    =   8325
   ClientLeft      =   2415
   ClientTop       =   2655
   ClientWidth     =   13425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00400000&
   Icon            =   "frmMain2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   13425
   WindowState     =   2  'Maximized
   Begin VB.Frame FraTerminate 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2FFFF&
      Caption         =   "Employees who are due for termination"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   8160
      TabIndex        =   52
      Top             =   7680
      Visible         =   0   'False
      Width           =   8010
      Begin MSComctlLib.ListView LvwTerminate 
         Height          =   5715
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   10081
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgTree"
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Height          =   270
         Left            =   165
         TabIndex        =   6
         Top             =   4485
         Width           =   3270
      End
   End
   Begin VB.Frame FraBirthDays 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2FFFF&
      Caption         =   "Birthdays Due This Month"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   7920
      TabIndex        =   49
      Top             =   7560
      Visible         =   0   'False
      Width           =   7125
      Begin MSComctlLib.ListView LvwBirthdays 
         Height          =   5595
         Left            =   120
         TabIndex        =   50
         Top             =   480
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   9869
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgTree"
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "EmpCode"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Employee's Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date Of Birth"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "BirthDay Date"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Height          =   270
         Left            =   165
         TabIndex        =   51
         Top             =   4485
         Width           =   3270
      End
   End
   Begin VB.CommandButton cmdProb 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   10800
      TabIndex        =   48
      Top             =   8040
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imgPD 
      Left            =   4800
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":0442
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraLog 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8400
      Left            =   13200
      TabIndex        =   17
      Top             =   840
      Width           =   12135
      Begin VB.Frame fraLogo 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   2370
         Left            =   3765
         TabIndex        =   18
         Top             =   3060
         Width           =   4860
         Begin VB.Label lblLogo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Personnel Director"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   240
            Left            =   1365
            TabIndex        =   19
            Top             =   1515
            Width           =   1815
         End
         Begin VB.Label lblPro 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Infiniti Systems Ltd"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   165
            Index           =   0
            Left            =   1290
            TabIndex        =   20
            Top             =   1815
            Width           =   3000
         End
         Begin VB.Image Image3 
            Height          =   1605
            Left            =   1320
            Picture         =   "frmMain2.frx":0894
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1890
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            BorderColor     =   &H00808080&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   1965
            Left            =   90
            Shape           =   4  'Rounded Rectangle
            Top             =   120
            Width           =   4470
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00FF8080&
            FillColor       =   &H00FFFFFF&
            Height          =   1305
            Left            =   1320
            Shape           =   3  'Circle
            Top             =   300
            Width           =   1935
         End
      End
      Begin MSComDlg.CommonDialog cdl 
         Left            =   1320
         Top             =   3120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame fraEmploymentExpiry 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2FFFF&
      Caption         =   "Employees Whose Employment is due to Expire"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5280
      Left            =   840
      TabIndex        =   42
      Top             =   9120
      Visible         =   0   'False
      Width           =   8730
      Begin VB.CommandButton cmdOkExpiry 
         BackColor       =   &H80000000&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7865
         TabIndex        =   43
         Top             =   4575
         Width           =   700
      End
      Begin MSComctlLib.ListView lstExpiry 
         Height          =   4020
         Left            =   -3360
         TabIndex        =   44
         Top             =   1800
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   7091
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgTree"
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label lblExpiry 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label5"
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   4680
         Width           =   465
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Height          =   270
         Left            =   165
         TabIndex        =   45
         Top             =   4485
         Width           =   3270
      End
   End
   Begin VB.Frame fraProb 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2FFFF&
      Caption         =   "Employees Due for Confirmation from Probation"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   7680
      TabIndex        =   34
      Top             =   7440
      Visible         =   0   'False
      Width           =   9810
      Begin MSComctlLib.ListView lvwProb 
         Height          =   5595
         Left            =   120
         TabIndex        =   35
         Top             =   480
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   9869
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgTree"
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblProb 
         BackStyle       =   0  'Transparent
         Height          =   270
         Left            =   165
         TabIndex        =   36
         Top             =   4485
         Width           =   3270
      End
   End
   Begin VB.CommandButton cmdShowPrompts 
      Caption         =   "Show Prompts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10680
      TabIndex        =   33
      Top             =   7320
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Frame fraContracts 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2FFFF&
      Caption         =   "Employees About to Retire"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   7440
      TabIndex        =   30
      Top             =   7200
      Visible         =   0   'False
      Width           =   8730
      Begin MSComctlLib.ListView lvwContracts 
         Height          =   5595
         Left            =   240
         TabIndex        =   31
         Top             =   480
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   9869
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgTree"
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label lblContracts 
         BackStyle       =   0  'Transparent
         Height          =   270
         Left            =   165
         TabIndex        =   32
         Top             =   4485
         Width           =   3270
      End
   End
   Begin VB.Frame fraVisa 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2FFFF&
      Caption         =   "Employees whose Contracts are about ot expire"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   7200
      TabIndex        =   27
      Top             =   6720
      Visible         =   0   'False
      Width           =   8730
      Begin MSComctlLib.ListView lvwVisa 
         Height          =   5595
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   9869
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgTree"
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label lblVisa 
         BackStyle       =   0  'Transparent
         Height          =   270
         Left            =   165
         TabIndex        =   29
         Top             =   4485
         Width           =   3270
      End
   End
   Begin MSComctlLib.ImageList imgTree 
      Left            =   5400
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   27
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":21DF96
            Key             =   "A"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":21E3E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":21E83A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":21EB54
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":21EE6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":21F2C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":21F5DA
            Key             =   "B"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":21FA2C
            Key             =   "C"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":21FE7E
            Key             =   "D"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":2202D0
            Key             =   "E"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":220722
            Key             =   "F"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":220B74
            Key             =   "G"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":220FC6
            Key             =   "H"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":221418
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":22186A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":221CBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":22210E
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":222560
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":2229B2
            Key             =   "I"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":222E04
            Key             =   "v"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":223256
            Key             =   "P"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":2236A8
            Key             =   "Z"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":223AFA
            Key             =   "J"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":223F4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":22405E
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":224170
            Key             =   "O"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":22448A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwMain 
      Height          =   6930
      Left            =   0
      TabIndex        =   4
      Top             =   1150
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   12224
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   176
      LabelEdit       =   1
      Style           =   5
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "imgTree"
      BorderStyle     =   1
      Appearance      =   0
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
   Begin VB.Frame fra1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13665
      Begin VB.CommandButton Command1 
         Height          =   495
         Left            =   12240
         Picture         =   "frmMain2.frx":2248DC
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   120
         Width           =   615
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5295
         Top             =   690
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   22
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain2.frx":224D1E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain2.frx":224EB8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain2.frx":22530A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain2.frx":22575C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain2.frx":225BAE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin HoverButton.Button cmdDetails 
         Height          =   570
         Left            =   600
         TabIndex        =   53
         Top             =   120
         Visible         =   0   'False
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1005
         BackColor       =   8388608
         HoverBackColor  =   8388608
         Border          =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontOverX {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   8421504
         HilightColor    =   8388608
         ShadowColor     =   8388608
         HoverHilightColor=   8388608
         HoverShadowColor=   8388608
         ForeColor       =   16777215
         HoverForeColor  =   8454016
         Caption         =   "Calc"
         CaptionDown     =   "Calc"
         CaptionOver     =   "Calc"
         ShowFocusRect   =   0   'False
         Sink            =   -1  'True
         Picture         =   "frmMain2.frx":226000
         Style           =   2
         PictureLocation =   1
         ButtonStyleX    =   0
         State           =   0
         IconHeight      =   0
         IconWidth       =   0
      End
      Begin HoverButton.Button cmdFind 
         Height          =   570
         Left            =   0
         TabIndex        =   54
         Top             =   120
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1005
         BackColor       =   8388608
         HoverBackColor  =   8388608
         Border          =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontOverX {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   8421504
         HilightColor    =   8388608
         ShadowColor     =   8388608
         HoverHilightColor=   8388608
         HoverShadowColor=   8388608
         ForeColor       =   16777215
         HoverForeColor  =   8454016
         Caption         =   "Find"
         CaptionDown     =   "Find"
         CaptionOver     =   "Find"
         ShowFocusRect   =   0   'False
         Sink            =   -1  'True
         Picture         =   "frmMain2.frx":226112
         Style           =   2
         PictureLocation =   1
         ButtonStyleX    =   0
         State           =   0
         IconHeight      =   0
         IconWidth       =   0
      End
      Begin VB.Label txtDetails 
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   2400
         TabIndex        =   16
         Top             =   120
         Width           =   5370
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   1440
         Picture         =   "frmMain2.frx":22BF84
         Top             =   120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image cmdCurrUsers 
         Height          =   480
         Left            =   1440
         Picture         =   "frmMain2.frx":22CDC6
         Top             =   120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   525
         Left            =   12870
         Picture         =   "frmMain2.frx":22DC08
         Stretch         =   -1  'True
         Top             =   135
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "Personnel Director"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   10170
         TabIndex        =   9
         Top             =   330
         Width           =   1575
      End
   End
   Begin MSComctlLib.ImageList imgMain 
      Left            =   6000
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   22
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":44B30A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":44B4A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":44B8F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":44BD48
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":44C19A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   640
      Width           =   1995
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Task "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   200
         TabIndex        =   3
         Top             =   195
         Width           =   525
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5160
      TabIndex        =   8
      Top             =   645
      Width           =   8145
      Begin VB.Frame fracmd 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   5010
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   3060
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   5.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1680
            Picture         =   "frmMain2.frx":44C5EC
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Cancel Process"
            Top             =   120
            Width           =   405
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   5.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   420
            Picture         =   "frmMain2.frx":44C6EE
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Save Record"
            Top             =   120
            Width           =   405
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   5.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1260
            Picture         =   "frmMain2.frx":44C7F0
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Delete Record"
            Top             =   120
            Width           =   405
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   5.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   840
            Picture         =   "frmMain2.frx":44CCE2
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Edit Record"
            Top             =   120
            Width           =   405
         End
         Begin VB.CommandButton cmdNew 
            Caption         =   "New"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   5.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   120
            Picture         =   "frmMain2.frx":44CDE4
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Add New record"
            Top             =   120
            Width           =   405
         End
      End
      Begin MSComctlLib.Toolbar tlbPD 
         Height          =   345
         Left            =   2280
         TabIndex        =   64
         Top             =   120
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         ButtonWidth     =   4128
         ButtonHeight    =   556
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgPD"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Personnel Director Prompts"
               Key             =   "PDP"
               ImageIndex      =   1
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   7
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "EDCP"
                     Text            =   "Employees due for confirmation from probation"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "TRM"
                     Text            =   "Employees due for termination"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "EDCE"
                     Text            =   "Employees whose contracts are due to expire"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "EDTR"
                     Text            =   "Employees who are due to Retire"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "BDY"
                     Text            =   "Birthdays due this month"
                  EndProperty
                  BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Sep"
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PS"
                     Text            =   "PROMPTS SETUP"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label lblEmpCount 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1530
         TabIndex        =   26
         Top             =   135
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lblEmpC 
         BackColor       =   &H80000000&
         Caption         =   "Employee Count:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   45
         TabIndex        =   25
         Top             =   135
         Visible         =   0   'False
         Width           =   1485
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   640
      Width           =   3270
      Begin VB.Label lblECount 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1800
         TabIndex        =   24
         Top             =   200
         Width           =   105
      End
      Begin VB.Label lblEmpList 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee List"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   200
         Width           =   1365
      End
   End
   Begin VB.Frame frmCoDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "Company Details"
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   5520
      TabIndex        =   55
      Top             =   2880
      Visible         =   0   'False
      Width           =   6135
      Begin VB.Label lblTotalCategories 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4560
         TabIndex        =   63
         Top             =   1680
         Width           =   105
      End
      Begin VB.Label lblTotalOUs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4560
         TabIndex        =   62
         Top             =   1320
         Width           =   105
      End
      Begin VB.Label lblTotalEmployees 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4560
         TabIndex        =   61
         Top             =   960
         Width           =   105
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Categories"
         Height          =   195
         Left            =   2640
         TabIndex        =   60
         Top             =   1680
         Width           =   1185
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Organization Units"
         Height          =   195
         Left            =   2640
         TabIndex        =   59
         Top             =   1320
         Width           =   1725
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Employees"
         Height          =   195
         Left            =   2640
         TabIndex        =   58
         Top             =   960
         Width           =   1170
      End
      Begin VB.Label lblPro 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Infiniti Systems Ltd "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   360
         TabIndex        =   57
         Top             =   2400
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1470
         Left            =   480
         Picture         =   "frmMain2.frx":44CEE6
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1425
      End
      Begin VB.Label lblcompanyname 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Infiniti  System Ltd"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   360
         TabIndex        =   56
         Top             =   2160
         Width           =   1815
      End
   End
   Begin VB.Frame fraEmployees 
      BackColor       =   &H00F2FFFF&
      BorderStyle     =   0  'None
      Height          =   6975
      Left            =   2040
      TabIndex        =   21
      Top             =   1150
      Visible         =   0   'False
      Width           =   2550
      Begin VB.CheckBox chkIncChildren 
         Caption         =   "Incl. Children"
         Height          =   195
         Left            =   990
         TabIndex        =   47
         Top             =   0
         Width           =   1275
      End
      Begin VB.ComboBox cboStructure 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmMain2.frx":66A5E8
         Left            =   0
         List            =   "frmMain2.frx":66A5FB
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   240
         Width           =   2475
      End
      Begin VB.ComboBox cboTerms 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmMain2.frx":66A635
         Left            =   0
         List            =   "frmMain2.frx":66A648
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   1080
         Width           =   2475
      End
      Begin VB.ComboBox cboCat 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1680
         Width           =   2475
      End
      Begin MSComctlLib.ListView lvwEmp 
         Height          =   4875
         Left            =   0
         TabIndex        =   22
         Top             =   1800
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   8599
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgTree"
         ForeColor       =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Org .Unit:"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   1305
      End
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Category:"
         Height          =   225
         Index           =   1
         Left            =   0
         TabIndex        =   39
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Terms:"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   38
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLog 
         Caption         =   "Log-In"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAdjustPay 
         Caption         =   "Adjust Pay"
      End
      Begin VB.Menu mnuG 
         Caption         =   "Groups"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuU 
         Caption         =   "Users"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMA 
         Caption         =   "Module Access"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEA 
         Caption         =   "Employee Access"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuShowPrompts 
         Caption         =   "Show Prompts"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileAdministrator 
         Caption         =   "Administrator"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileAdministratorUserAccounts 
         Caption         =   "User Accounts"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuphotos 
         Caption         =   "Employees Photos Seup"
      End
      Begin VB.Menu mnuUpdateEmpDept 
         Caption         =   "Update employee departments"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLogout 
         Caption         =   "Log Out"
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangePassword 
         Caption         =   "Change Password"
      End
      Begin VB.Menu mnuSepChangePassword 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuSSReports 
      Caption         =   "Reports"
      Begin VB.Menu mnuGeneral 
         Caption         =   "List"
         Begin VB.Menu mnuEmpTerms 
            Caption         =   "Employment Terms"
         End
         Begin VB.Menu mnuPosition 
            Caption         =   "Job Positions"
         End
         Begin VB.Menu mnuHeadCountNationalityList 
            Caption         =   "Nationality List"
         End
         Begin VB.Menu mnuReligion 
            Caption         =   "Religions"
         End
         Begin VB.Menu mnuHeadCountTribesList 
            Caption         =   "Tribe List"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuOrganizationAssets 
            Caption         =   "Organization Assets"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCompanyAssets 
            Caption         =   "Company assets"
            Shortcut        =   ^A
         End
         Begin VB.Menu mnuDivisions 
            Caption         =   "Department"
         End
         Begin VB.Menu mnuHeadcountCategGradesList 
            Caption         =   "Grades"
         End
         Begin VB.Menu mnuBioDataTypes 
            Caption         =   "Bio-data types"
         End
         Begin VB.Menu mnuContactTypes 
            Caption         =   "Contact types"
         End
         Begin VB.Menu mnuDefDetTypes 
            Caption         =   "Defined details types"
         End
         Begin VB.Menu mnurptBanks 
            Caption         =   "Banks"
         End
         Begin VB.Menu mnurptBranches 
            Caption         =   "Bank Branches"
         End
         Begin VB.Menu mnurptSectors 
            Caption         =   "Sectors"
         End
         Begin VB.Menu mnuprojects 
            Caption         =   "Projects"
         End
         Begin VB.Menu mnuProjectFunding 
            Caption         =   "Project Funding"
         End
         Begin VB.Menu mnuemployeeprojects 
            Caption         =   "Employee Export List"
         End
         Begin VB.Menu mnurptLocations 
            Caption         =   "Locations"
         End
         Begin VB.Menu mnuPrompts 
            Caption         =   "Prompts"
            Begin VB.Menu mnuBirthDayPrompts 
               Caption         =   "Birthdays for this month"
            End
            Begin VB.Menu mnuContractsAlmost 
               Caption         =   "Contracts about to end"
            End
            Begin VB.Menu mnuTerminationPrompt 
               Caption         =   "Employees Due for Termination"
            End
            Begin VB.Menu mnuDueProbationEnds 
               Caption         =   "Employees due from probation"
            End
            Begin VB.Menu mnuRetirementPrompt 
               Caption         =   "Employees Due For Retirement"
            End
         End
      End
      Begin VB.Menu mnuemployee 
         Caption         =   "Employee"
         Begin VB.Menu mnuHDEmpListByDept 
            Caption         =   "Age Profile"
         End
         Begin VB.Menu mnudptHierarchy 
            Caption         =   "Department Hierarchy"
         End
         Begin VB.Menu mnuEmployeeGenDet 
            Caption         =   "General details"
         End
         Begin VB.Menu mnuCheckList 
            Caption         =   "Pre-Employment Check List"
         End
         Begin VB.Menu mnuEmpContactDetails 
            Caption         =   "Contact details"
            Visible         =   0   'False
            Begin VB.Menu mnuCGrowing 
               Caption         =   "Growing Format"
            End
         End
         Begin VB.Menu mnuContactList 
            Caption         =   "Contacts"
         End
         Begin VB.Menu mnuEmployeeDefDetails 
            Caption         =   "Defined details"
            Visible         =   0   'False
            Begin VB.Menu mnuDGrowing 
               Caption         =   "Growing format"
            End
         End
         Begin VB.Menu mnuDlistReport 
            Caption         =   "Defined Details"
         End
         Begin VB.Menu mnuEmployeeBioData 
            Caption         =   "Bio-data main"
            Visible         =   0   'False
            Begin VB.Menu mnuGrow 
               Caption         =   "Growing Format"
            End
         End
         Begin VB.Menu mnuBlist 
            Caption         =   "Bio Data"
         End
         Begin VB.Menu mnuEmpNextOfKin 
            Caption         =   "Next of kin"
         End
         Begin VB.Menu mnuPhyAddresses 
            Caption         =   "Physical Addresses"
         End
         Begin VB.Menu mnuEmployeeFamily 
            Caption         =   "Family"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuEmployeeReferees 
            Caption         =   "Referees"
         End
         Begin VB.Menu mnuEmpContractDetails 
            Caption         =   "Contract details"
         End
         Begin VB.Menu mnuIssuedAssets 
            Caption         =   "Issued assets"
         End
         Begin VB.Menu mnuEmployeeExpatriateVisaDetails 
            Caption         =   "Expatriate visa details"
         End
         Begin VB.Menu mnuEmployeeEmpHist 
            Caption         =   "Employment history"
         End
         Begin VB.Menu mnuEmpEducHistory 
            Caption         =   "Education history"
         End
         Begin VB.Menu mnuEmpProfQualification 
            Caption         =   "Professional qualification"
         End
         Begin VB.Menu mnuEmpAwards 
            Caption         =   "Awards"
         End
         Begin VB.Menu mnuEmpBankDetails 
            Caption         =   "Bank details"
         End
         Begin VB.Menu mnuEmpProjects 
            Caption         =   "Employee Projects"
         End
         Begin VB.Menu mnuSepEmpHistory 
            Caption         =   "-"
         End
         Begin VB.Menu mnuReengage 
            Caption         =   "Reengagement"
            Begin VB.Menu mnuReengaded 
               Caption         =   "Reengaged Employees"
            End
            Begin VB.Menu mnuhistory 
               Caption         =   "Reengagement History"
            End
         End
         Begin VB.Menu mnuArchived 
            Caption         =   "Archived Employee "
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDisengaged 
            Caption         =   "Disengaged"
            Begin VB.Menu mnuDisengagedCount 
               Caption         =   "Count"
            End
            Begin VB.Menu MnuDisList 
               Caption         =   "List"
            End
         End
         Begin VB.Menu mnuEmpDefDetByType 
            Caption         =   "By type"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnuEmpDefDetByEmp 
            Caption         =   "By employee"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSepAwards 
            Caption         =   "-"
         End
         Begin VB.Menu mnuHeadcountCategGradesEmpList 
            Caption         =   "Employees by Grade"
         End
         Begin VB.Menu mnuChartEmpsDpt 
            Caption         =   "Employees Per Department"
         End
         Begin VB.Menu mnuEmployeeHcount 
            Caption         =   "Head Count"
            Begin VB.Menu mnuHAnalysis 
               Caption         =   "Head count Analysis"
            End
            Begin VB.Menu mnuHDEmpCountDirect 
               Caption         =   "Head by Designitions"
            End
         End
         Begin VB.Menu mnuEmpBirthdays 
            Caption         =   "Birthdays"
            Visible         =   0   'False
            Begin VB.Menu mnuBirthdayByEmp 
               Caption         =   "By employee"
            End
            Begin VB.Menu mnuBirthdayByDept 
               Caption         =   "By department"
            End
         End
         Begin VB.Menu mnuEmpMedicalRept 
            Caption         =   "Medical report"
         End
         Begin VB.Menu mnuEmpPhysicallyChallenged 
            Caption         =   "Physically challenged"
         End
         Begin VB.Menu mnuexemployeeinfo 
            Caption         =   "Ex Employee Info"
         End
         Begin VB.Menu mnuempperdatasumsheet 
            Caption         =   "Employee Personal Data Summary Sheet"
         End
         Begin VB.Menu mnuepd_beneficiaries 
            Caption         =   "Employees Personal Details & Beneficiaries"
         End
         Begin VB.Menu pifform 
            Caption         =   "Personel Information Form"
         End
         Begin VB.Menu mnufiltered 
            Caption         =   "Filter Report By:"
            Begin VB.Menu mnureportperfilter 
               Caption         =   "Report Per Filter"
            End
            Begin VB.Menu mnudrilldown 
               Caption         =   "Drill Down"
            End
         End
      End
      Begin VB.Menu mnuHeadCount 
         Caption         =   "Headcount"
         Visible         =   0   'False
         Begin VB.Menu mnuHeadCountDesignation 
            Caption         =   "Designation"
            Begin VB.Menu mnuHDEmpListDirect 
               Caption         =   "List"
            End
            Begin VB.Menu mnuSepListCount 
               Caption         =   "-"
            End
         End
         Begin VB.Menu mnuHeadCountDesignationList 
            Caption         =   "List"
            Begin VB.Menu mnuHDByEmpTerms 
               Caption         =   "Employment terms"
            End
            Begin VB.Menu mnuHeadCountDesignationEmpCount 
               Caption         =   "Count"
               Begin VB.Menu mnuHDCountByEmpDept 
                  Caption         =   "By department"
               End
               Begin VB.Menu mnuHDCountByEmpTerms 
                  Caption         =   "By employment terms"
               End
            End
            Begin VB.Menu mnuHeadcountCategGradesEmp 
               Caption         =   "Employee"
               Begin VB.Menu mnuHeadcountCategGradesEmpCount 
                  Caption         =   "Count"
               End
            End
            Begin VB.Menu mnuHeadcountCategGradesByDept 
               Caption         =   "By department"
               Begin VB.Menu mnuHeadcountCategGradesByDeptCount 
                  Caption         =   "Count"
               End
               Begin VB.Menu mnuHeadcountCategGradesByDeptList 
                  Caption         =   "List"
               End
            End
         End
         Begin VB.Menu mnuHeadCountNationality 
            Caption         =   "Nationality"
            Begin VB.Menu mnuNationalityEmployee 
               Caption         =   "Employee"
               Begin VB.Menu mnuNationalityEmpList 
                  Caption         =   "List"
               End
               Begin VB.Menu mnuHeadCountNationalityCount 
                  Caption         =   "Count"
               End
            End
            Begin VB.Menu mnuHeadCountNationalityByDept 
               Caption         =   "By department"
               Begin VB.Menu mnuHeadCountNationalityByDeptList 
                  Caption         =   "List"
               End
               Begin VB.Menu mnuHeadCountNationalityByDeptCount 
                  Caption         =   "Count"
               End
            End
         End
         Begin VB.Menu mnuHeadCountTribes 
            Caption         =   "Tribes"
            Visible         =   0   'False
            Begin VB.Menu mnuHeadCountTribesEmployee 
               Caption         =   "Employee"
               Begin VB.Menu mnuHeadCountTribesEmployeeList 
                  Caption         =   "List"
               End
               Begin VB.Menu mnuHeadCountTribesEmployeeCount 
                  Caption         =   "Count"
               End
            End
            Begin VB.Menu mnuHeadCountTribesDept 
               Caption         =   "By department"
               Begin VB.Menu mnuHeadCountTribesDeptList 
                  Caption         =   "List"
               End
               Begin VB.Menu mnuHeadCountTribesDeptCount 
                  Caption         =   "Count"
               End
            End
         End
         Begin VB.Menu mnuHeadCountGender 
            Caption         =   "Gender"
            Begin VB.Menu mnuHeadCountGenderEmployee 
               Caption         =   "Employee"
               Visible         =   0   'False
               Begin VB.Menu mnuHeadCountGenderEmployeeList 
                  Caption         =   "List"
               End
               Begin VB.Menu mnuHeadCountGenderEmployeeCount 
                  Caption         =   "Count"
               End
            End
            Begin VB.Menu mnuHeadCountGenderDept 
               Caption         =   "By department"
               Begin VB.Menu mnuHeadCountGenderDeptList 
                  Caption         =   "List"
               End
               Begin VB.Menu mnuHeadCountGenderDeptCount 
                  Caption         =   "Count"
               End
            End
         End
      End
      Begin VB.Menu mnuSalaryInfo 
         Caption         =   "Salary information"
         Visible         =   0   'False
         Begin VB.Menu mnuSalaryInfoSalaryDetails 
            Caption         =   "Salary details"
         End
         Begin VB.Menu mnuSalaryInfoByDept 
            Caption         =   "Details by department"
         End
         Begin VB.Menu mnuSepSalIncrement 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSalaryInfoIncrement 
            Caption         =   "Salary increment"
            Enabled         =   0   'False
            Visible         =   0   'False
            Begin VB.Menu mnuSalaryInfoIncrementList 
               Caption         =   "List"
            End
            Begin VB.Menu mnuSalaryInfoIncrementListByDept 
               Caption         =   "List by department"
            End
         End
      End
      Begin VB.Menu mnuBonus 
         Caption         =   "Bonus"
         Visible         =   0   'False
         Begin VB.Menu mnuBonusAnnualList 
            Caption         =   "Annual list"
         End
         Begin VB.Menu mnuBonusAnnualListNoAwards 
            Caption         =   "Annual (No awards)"
         End
         Begin VB.Menu mnuSepBonus 
            Caption         =   "-"
         End
         Begin VB.Menu mnuBonusAnnualListOther 
            Caption         =   "Other bonus list"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuLongService 
         Caption         =   "Long service"
         Visible         =   0   'False
         Begin VB.Menu mnuLongServiceList 
            Caption         =   "List"
         End
         Begin VB.Menu mnuLongServiceByDept 
            Caption         =   "List by department"
         End
         Begin VB.Menu mnuSepLongService 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLongServiceGroupingByYears 
            Caption         =   "Grouping by years"
         End
      End
      Begin VB.Menu mnuRetirement 
         Caption         =   "Retirement"
         Visible         =   0   'False
         Begin VB.Menu mnuRetirementSchedule 
            Caption         =   "Schedule"
         End
         Begin VB.Menu mnuSepRetByDept 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRetirementByDept 
            Caption         =   "By department"
         End
         Begin VB.Menu mnuRetirementByDeptWithoutSal 
            Caption         =   "By department without salary"
         End
         Begin VB.Menu mnuSepByDepWithoutSal 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRetirementActualDates 
            Caption         =   "Actual dates"
         End
         Begin VB.Menu mnuRetirementEarly 
            Caption         =   "Early retirees"
         End
         Begin VB.Menu mnuRetirementByYearGrouping 
            Caption         =   "By year grouping"
         End
      End
      Begin VB.Menu mnuTurnOver 
         Caption         =   "Turn over"
         Visible         =   0   'False
         Begin VB.Menu mnuTurnOverHires 
            Caption         =   "Hires"
         End
         Begin VB.Menu mnuTurnOverHiresYearToDate 
            Caption         =   "Hires year-to-date"
         End
         Begin VB.Menu mnuSepHires 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTurnOverLeft 
            Caption         =   "Left"
         End
         Begin VB.Menu mnuTurnOverLeftYearToDate 
            Caption         =   "Left year-to-date"
         End
         Begin VB.Menu mnuSepTransfers 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTurnOverTransfers 
            Caption         =   "Transfers"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuMissingMandatory 
         Caption         =   "Missing mandatory data"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpAboutSoftware 
         Caption         =   "About Software"
      End
   End
End
Attribute VB_Name = "frmMain2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MyNodes As Node
Dim MyData() As Variant
Dim i As Long
Dim ASearch As Boolean
Dim MStruc As String
Dim ProbPromptDays As Prompts

Private ChangedByCode As Boolean

''Private EmpCats As HRCORE.EmployeeCategories
''Private empTerms As HRCORE.EmploymentTerms
''Private OUs As HRCORE.OrganizationUnits

Private selEmpCat As HRCORE.EmployeeCategory
Private selEmpTerm As HRCORE.EmploymentTerm
Private selOU As HRCORE.OrganizationUnit
'Private FilteredEmpList As HRCORE.Employees
Private FilterOURecursively As Boolean  'flags whether to filter OUs Recursively: depends on chkIncChildren.Value
''Private WithEvents background As NetFX20Wrapper.BackgroundWorkerWrapper

Private Declare Function GetComputerName Lib "kernel32" Alias _
    "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

''Private Sub background_ProgressChanged(ByVal sender As Variant, ByVal e As NetFX20Wrapper.ProgressChangedEventArgsWrapper)
    ''ProgressBar1 = e.ProgressPercentage
''End Sub

'Private Sub background_RunWorkerCompleted(ByVal sender As Variant, ByVal e As NetFX20Wrapper.RunWorkerCompletedEventArgsWrapper)
'    If e.Error.Number <> 0 Then
'        MsgBox "Error in background process: " & e.Error.Description
'    Else
'        If e.Cancelled Then
'            MsgBox "Background process cancelled."
'        Else
'            MsgBox "Completed processing of: " & e.GetResult
'        End If
'    End If
'
'    Command2.Enabled = False
'    ProgressBar1 = 0
'End Sub


'THIS IS NECESSARY FOR PASSING PARAMETERFIEL OBJECT VALUES

Private Sub LogIn()

    On Error GoTo ErrorTrap

    Call CloseMyWindows

    fraEmployees.Visible = True
    frmMain2.fraLog.Visible = True
    'frmLog.Show vbModal

   ' GoTo SkipSecurity
   
    'HRMSEC LOGIN USING NEW SECURITY
    
    Set currUser = gUser.LoginUser("PDR")
    If Not (currUser Is Nothing) Then
        frmMain2.Caption = "Infiniti HRMIS - PDR [Current User:\" & currUser.FullNames & "]" & " vs " & App.FileDescription & ""
    Else
        'end the program
        
        End
    End If
    
SkipSecurity:
    frmMain2.fraLog.Visible = False
    
    getphotosetup
    Call InitializeHRCOREObjects

    Call LoadCbo    'loads the Employee Terms, Employee Categories, and STypes

    Call LoadEmployeeList     'loads the employees into the listview
    
    
  ''  StartBackground background, "Loading Employee banks accounts"
'------------

       If (objEmployeeBankAccounts Is Nothing) Then
             Set objEmployeeBankAccounts = New EmployeeBankAccounts2
             objEmployeeBankAccounts.GetActiveEmployeeBankAccounts
             Set myEmployeeBankAccounts = objEmployeeBankAccounts
            '' Set pEmployeeBankAccounts = objEmployeeBankAccounts
       End If
'----------------------
Dim ff As Integer
ff = 0
''----------------

    Exit Sub

ErrorTrap:
    MsgBox err.Description, vbExclamation
    
    GoTo SkipSecurity
End Sub
Private Sub Reload_objects()

    On Error GoTo ErrorTrap
MousePointer = vbHourglass

    
    
      If (objEmployeeBankAccounts Is Nothing) Then
             Set objEmployeeBankAccounts = New EmployeeBankAccounts2
             objEmployeeBankAccounts.GetActiveEmployeeBankAccounts
             Set myEmployeeBankAccounts = objEmployeeBankAccounts
             ''Set pEmployeeBankAccounts = objEmployeeBankAccounts
      End If
    Call InitializeHRCOREObjects

    Call LoadCbo    'loads the Employee Terms, Employee Categories, and STypes

    Call LoadEmployeeList     'loads the employees into the listview
    
    
'------------

       If (objEmployeeBankAccounts Is Nothing) Then
             Set objEmployeeBankAccounts = New EmployeeBankAccounts2
             objEmployeeBankAccounts.GetActiveEmployeeBankAccounts
             Set myEmployeeBankAccounts = objEmployeeBankAccounts
            '' Set pEmployeeBankAccounts = objEmployeeBankAccounts
       End If
'----------------------

MousePointer = vbDefault

    Exit Sub

ErrorTrap:
    MsgBox err.Description, vbExclamation
    
MousePointer = vbDefault
End Sub


Private Sub cboCat_Click()
    If ChangedByCode Then Exit Sub
    'call procedure to filter employees
    FilterEmployees
End Sub

Private Sub FilterEmployees()
    Dim TheCategory As HRCORE.EmployeeCategory
    Dim TheEmpTerm As HRCORE.EmploymentTerm
    Dim theOU As HRCORE.OrganizationUnit
    Dim ResultList As HRCORE.Employees

    On Error GoTo ErrorHandler

    'check the OU
    If UCase(cboStructure.Text) <> "(ALL ORGANIZATION UNITS)" Then
        Set theOU = OUs.FindOrganizationUnit(CLng(cboStructure.ItemData(cboStructure.ListIndex)))
    Else
        Set theOU = Nothing
    End If

    'check the Terms
    If UCase(cboTerms.Text) <> "(ALL TERMS)" Then
        Set TheEmpTerm = empTerms.FindEmploymentTerm(CLng(cboTerms.ItemData(cboTerms.ListIndex)))
    Else
        Set TheEmpTerm = Nothing
    End If

    'check the Categories
    If UCase(cboCat.Text) <> "(ALL CATEGORIES)" Then
        Set TheCategory = EmpCats.FindEmployeeCategory(CLng(cboCat.ItemData(cboCat.ListIndex)))
    Else
        Set TheCategory = Nothing
    End If

    Set ResultList = AllEmployees.FilterEmployeesEx(clientemplist, AllEmployees, theOU, TheEmpTerm, TheCategory, FilterOURecursively)
    
    LoadEmployeeListFiltered ResultList

    Exit Sub

ErrorHandler:
    MsgBox "An error has occurred" & vbNewLine & err.Description, vbInformation, TITLES
    Set ResultList = Nothing
    LoadEmployeeListFiltered ResultList
End Sub

Private Sub cboStructure_Click()
    If ChangedByCode Then Exit Sub
    FilterEmployees
End Sub

Public Sub cboTerms_Click()
    If ChangedByCode Then Exit Sub
    FilterEmployees
End Sub

Private Sub chkIncChildren_Click()
    If chkIncChildren.value = vbChecked Then
        FilterOURecursively = True
    Else
        FilterOURecursively = False
    End If
    
    'perform the filter
    FilterEmployees
End Sub

Public Sub cmdCancel_Click()
    On Error GoTo errHandler
    If TheLoadedForm.Name = "frmSysCheck" Or TheLoadedForm.Name = "frmPost" Or TheLoadedForm.Name = "frmBackUp" Or TheLoadedForm.Name = "frmPrompt" Then Exit Sub
    
    If TheLoadedForm.cmdCancel.Enabled = True Then
        If MsgBox("Do you want to save changes?", vbQuestion + vbYesNo, TITLES) = vbYes Then '
            TheLoadedForm.cmdSave_Click
            Exit Sub
        End If
        'If TheLoadedForm.Name = "frmCheck" Then frmMain2.cmdNew.Enabled = False
        cmdSave.Enabled = False
        cmdNew.Enabled = True
        cmdEdit.Enabled = True
        cmdCancel.Enabled = False
        cmdDelete.Enabled = True
        TheLoadedForm.cmdCancel_Click
    Else
        MsgBox "You cannot cancel the process.", vbInformation
    End If
    Exit Sub
errHandler:
    MsgBox err.Description, vbInformation
End Sub

Private Function ExemptForThisForm(TheForm As Form) As Boolean

    'This procedure checks whether to exempt the form from global menu click
    ExemptForThisForm = False

    If UCase(TheForm.Name) = "FRMCOMPANYDETAILS" Or TheForm.Name = UCase("frmOUEmployees") Then
        ExemptForThisForm = True
    End If
End Function

Public Sub cmdDelete_Click()
    On Error Resume Next
    If TheLoadedForm.Name = "frmEmpLApp" Or TheLoadedForm.Name = "frmSysCheck" Or TheLoadedForm.Name = "frmPost" Or TheLoadedForm.Name = "frmBackUp" Or TheLoadedForm.Name = "frmPrompt" Or TheLoadedForm.Name = "frmGenOpt" Then Exit Sub
    
        If ExemptForThisForm(TheLoadedForm) = True Then Exit Sub
        
        TheLoadedForm.cmdDelete.Enabled = True
        TheLoadedForm.cmdDelete.value = True
End Sub

Private Sub cmdDetails_Click()
    Shell ("Calc.exe")
End Sub

Public Sub cmdEdit_Click()
    On Error GoTo errHandler
    If TheLoadedForm.Name = "frmEmpLApp" Or TheLoadedForm.Name = "frmSysCheck" Or TheLoadedForm.Name = "frmPost" Or TheLoadedForm.Name = "frmBackUp" Or TheLoadedForm.Name = "frmPrompt" Or TheLoadedForm.Name = "frmEmpDivisions" Or TheLoadedForm.Name = "frmDivApp" Then Exit Sub
    
    cmdEdit.Enabled = False
    cmdNew.Enabled = False
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    cmdDelete.Enabled = False
    
    'call the edit buttom of the loaded form
    TheLoadedForm.cmdEdit.Enabled = True
    If TheLoadedForm.cmdEdit.Enabled = True Then
        TheLoadedForm.cmdEdit.value = True
    Else
        MsgBox "You cannot edit the record.", vbInformation
        cmdEdit.Enabled = True
        cmdNew.Enabled = True
        cmdSave.Enabled = False
        cmdCancel.Enabled = False
        cmdDelete.Enabled = True
        Exit Sub
    End If
    Exit Sub
errHandler:
End Sub

'Private Sub cmdEOk_Click()
'    fraEPrompt.Visible = False
'    If fraUPrompt.Visible = True Then
'        cmdUPrompt.SetFocus
'    End If
'End Sub

Private Sub cmdFind_Click()
    If ASearch = False Then
        MsgBox "Please ensure that the 'General Details' Window is open.", vbInformation
        Exit Sub
    End If

    Me.MousePointer = vbHourglass

    frmSearch.Show vbModal

    If Not Sel = "" Then
        With rsGlob
            If .RecordCount > 0 Then
                .MoveFirst
                .Find "EmpCode like '" & Sel & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    txtDetails.Caption = "Code: " & rsGlob!EmpCode & "     " & "Name: " & !SurName & "" & " " & !OtherNames & "" & " " & vbCrLf & _
                        "" & "ID No:" & " " & !IdNo & "" & "     " & "Date Employed:" & " " & !DEmployed & "" & "     " & "Gender:" & " " & !Gender & ""

                    If TheLoadedForm.Name = "frmEmpLeaves" Or TheLoadedForm.Name = "frmLeaveApp" Or TheLoadedForm.Name = "frmEmpLApp" Then
                        If TheLoadedForm.Name = "frmLeaveApp" Then
                            TheLoadedForm.DisplayEmp
                        End If

                        TheLoadedForm.DisplayRecords

                    End If

                End If
            End If
        End With

    End If

    Me.MousePointer = 0

End Sub

Public Sub cmdNew_Click()
    On Error Resume Next
    cmdNew.Enabled = False
    cmdEdit.Enabled = False
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    cmdDelete.Enabled = False
    If ExemptForThisForm(TheLoadedForm) = True Then Exit Sub
    
    Dim xx As String
    xx = TheLoadedForm.Name
    
    TheLoadedForm.cmdNew.Enabled = True
    If TheLoadedForm.cmdNew.Enabled = True Then
        Call TheLoadedForm.cmdNew_Click
    Else
        MsgBox "You can not add a new record.", vbInformation
        cmdNew.Enabled = True
        cmdEdit.Enabled = True
        cmdSave.Enabled = False
        cmdCancel.Enabled = False
        cmdDelete.Enabled = True
        Exit Sub
    End If
  
End Sub

Public Sub RestoreCommandButtonState()
    'This method will be called from the Loaded Forms
    cmdNew.Enabled = True
    cmdEdit.Enabled = True
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    cmdDelete.Enabled = True
End Sub

Private Sub cmdProb_Click()
    If fraProb.Visible = True Then
        fraProb.Visible = False
    ElseIf fraContracts.Visible = True Then
        fraContracts.Visible = False
    ElseIf fraVisa.Visible = True Then
        fraVisa.Visible = False
    ElseIf FraBirthDays.Visible = True Then
        FraBirthDays.Visible = False
    End If
    fraEmployees.Visible = True
    FraTerminate.Visible = False
End Sub

Public Sub cmdSave_Click()
    On Error GoTo errHandler
    
    If TheLoadedForm.Name = "frmEmpLApp" Or TheLoadedForm.Name = "frmSysCheck" Or TheLoadedForm.Name = "frmPost" Or TheLoadedForm.Name = "frmBackUp" Or TheLoadedForm.Name = "frmPrompt" Or TheLoadedForm.Name = "frmEmpLApp" Or TheLoadedForm.Name = "frmLeaveApp" Or TheLoadedForm.Name = "frmEmpDivisions" Or TheLoadedForm.Name = "frmDivApp" Then Exit Sub

        cmdSave.Enabled = False
        cmdCancel.Enabled = False
        cmdNew.Enabled = True
        cmdEdit.Enabled = True
        cmdDelete.Enabled = True

    If TheLoadedForm.cmdSave.Enabled = True Then
        TheLoadedForm.cmdSave_Click
    Else
        MsgBox "You cannot save a record.", vbInformation
        cmdSave.Enabled = True
        cmdCancel.Enabled = True
        cmdNew.Enabled = False
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
    End If

    Exit Sub
errHandler:
End Sub

Private Sub Command1_Click()
Reload_objects
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    Dim company As New CompanyDetails
    Dim MDOB As Date
    Dim MDEmp As Date
    Dim GenDetNode As Node  'will hold node for General Details
  
'    Set background = New NetFX20Wrapper.BackgroundWorkerWrapper
'    Dim e As New NetFX20Wrapper.RunWorkerCompletedEventArgsWrapper
  
  
    Set Image1.Picture = companyDetail.logo
'    Set EmpCats = New HRCORE.EmployeeCategories
'    Set empTerms = New HRCORE.EmploymentTerms
    'Set OUs = New HRCORE.OrganizationUnits

    FLoading = False
    oSmart.FReset Me
    FLoading = True
    
    ASearch = False
    OminisDB = False
    OData = True
    
    'Default Colour
    MyColor = &H8000000F '&HF2FFFF
    CConnect.CColor Me, MyColor

    'Exceptions to the color scheme
    fra1.BackColor = &H800000
    txtDetails.BackColor = &H800000
    txtDetails.ForeColor = &HFFFFFF
    Label1.BackColor = &H800000
    Label1.ForeColor = &HFFFFFF
    Frame2.BackColor = &H8000000F
    Frame3.BackColor = &H8000000F
    Frame5.BackColor = &H8000000F

    Call InitTree   'initializes items on the treeview menu

    Call LoadVar    'loads the General Options: RS1
    
    ''''''''''''''''''''''''
 
    '
    ''''''''''''''''''''''''
    Set rsGlob = CConnect.setGlobalRecordset
    Set rsGlob2 = rsGlob
    
    With lvwEmp
        .ColumnHeaders.add , , "Code", 300
        .ColumnHeaders.add , , "Names", 2500
        .View = lvwReport
    End With

    Set rsGenOpt = CConnect.GetRecordSet("SELECT * FROM GeneralOpt")
     
     'Set Currency RecordSet
    Set Rcurrency = CConnect.GetRecordSet(curSql)
    
    Call InitGrid   ' Initializes Listviews that will be used For Prompts
    
    'login
    

   

    Call mnuLog_Click
    
    ''''''''''''''''''''''''
   
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  'Load the Banks and Populate them
   
    
    'load the bank branches but don't populate
''    pBankBranches.GetActiveBankBranches
    
    'load employee bank accounts but don't populate
  ''  pEmployeeBankAccounts.GetActiveEmployeeBankAccounts
  
  
   If (objEmployeeBankAccounts Is Nothing) Then
             Set objEmployeeBankAccounts = New EmployeeBankAccounts2
             objEmployeeBankAccounts.GetActiveEmployeeBankAccounts
             Set myEmployeeBankAccounts = objEmployeeBankAccounts
             ''Set pEmployeeBankAccounts = objEmployeeBankAccounts
  End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




    'Get all the prompts delays from DB
    GetPromptDays
    
    If ProbPromptDays.Item(1).EnablePrompts = 1 Then 'Prompts Are Enabled
        If Not (ProbPrompt(1)) Then
            If Not (EmployeeTerminatePrompt(1)) Then
                If Not (RetirementPrompt(1)) Then
                    If Not (ContractExpirePrompt(1)) Then
                        If Not (GetBirthDays(1)) Then
                            'This means there are actually no prompts on
                            'ANY of the three prompts
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    'Company Snapshot
    lblcompanyname.Caption = companyDetail.CompanyName
    lblTotalEmployees.Caption = AllEmployees.count
    lblTotalOUs.Caption = OUs.count
    lblTotalCategories.Caption = EmpCats.count
    'Me.Caption = App.EXEName & App.FileDescription
    
    'SECURITY ENFORCEMENT ON EMPLOYEE REMUNERATION DISPLAY
    If currUser.CheckRight("ViewEmployeeRemuneration") = secNone Then
        'EMPLOYEES NOT ALLOWED TO VIEW RENUMERATION INFO
        mnuEmployeeGenDet.Enabled = False
    End If
    Dim k As Long
    k = getAccessiblePayrollids(currUser)
    Exit Sub
    
    
ErrorHandler:
    MsgBox "An error has occurred" & vbNewLine & err.Description, vbInformation, TITLES
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitializeHRCOREObjects()
    Set outs = New HRCORE.OrganizationUnitTypes
    Set OUnits = New HRCORE.OrganizationUnits
    
    Set OUs = New HRCORE.OrganizationUnits
    
    Set pBanks = New Banks
    
    Set pBankBranches = New BankBranches
    Set myEmployeeBankAccounts = New EmployeeBankAccounts2
    
    Set AllEmployeesPhotos = New EmployeesPhotos
    
    Set EmpCats = New HRCORE.EmployeeCategories
    Set empTerms = New HRCORE.EmploymentTerms
    Set empNationalities = New HRCORE.Nationalities
    Set empTribes = New HRCORE.Tribes
    Set empReligions = New HRCORE.Religions
    Set empPositions = New HRCORE.JobPositions
    Set empCurrencies = New HRCORE.Currencies
    Set empCountries = New HRCORE.Countries
    Set empLocations = New HRCORE.Locations
    Set empProjects = New HRCORE.Programmes
    Set empStaffCategories = New HRCORE.CSSSCategories
    Set objProgrammes = New HRCORE.Programmes
    Set objProgrammeFundings = New HRCORE.ProgrammeFundings
    Set objEmployeeProgrammeFundings = New HRCORE.EmployeeProgrammes
        'Set emps = New HRCORE.Employees
    
    
    company.LoadCompanyDetails
    objProgrammes.GetActiveProgrammes
    objProgrammeFundings.GetActiveProgrammeFundings
    objEmployeeProgrammeFundings.GetActiveEmployeeProgrammes
    
    ''code below takes too long. check y
        
    empLocations.GetActiveLocations
    EmpCats.GetActiveEmployeeCategories
   '' AllEmployeesPhotos.GetAccessibleEmployeesPhotosByUser currUser.UserID
    pBanks.GetAllBanks
    pBankBranches.GetAllBankBranches
''****************************************
''set accessrights
If Not (UCase(currUser.UserName) = "INFINITI") Then
Dim rs As New ADODB.Recordset
        Set rs = CConnect.GetRecordSet("exec sp_groupAccessRightsByUserID  " & currUser.UserID & "")
        If Not (rs.EOF) Then
            gAccessRightTypeId = rs!AccessRightTypeID
            gAccessRightName = rs!accessrighttypename
           
            gAccessRightClassIds = ""
            Dim i2 As Integer
            i2 = 0
              ReDim gAccessRighClassIdsArray(rs.RecordCount)
             While Not rs.EOF
             gAccessRightClassIds = gAccessRightClassIds & rs!AccessibleClassID & ","
              gAccessRighClassIdsArray(i2) = rs!AccessibleClassID
            rs.MoveNext
            i2 = i2 + 1
            Wend
'            Dim i3 As Integer
'            i3 = 0
'             ReDim gAccessRighClassIdsArray(i2)
'            While i3 <= i2
'             gAccessRighClassIdsArray(i3) = rs!AccessibleClassID
'            i3 = i3 + 1
'            Wend

        gAccessRightClassIds = Mid(gAccessRightClassIds, 1, Len(gAccessRightClassIds) - 1)
        End If
Else
  gAccessRightClassIds = ""
End If
''***************************************

    
    'emps.GetAllEmployees
    
    'PopulateEmployees emps
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim fm As Form

    If MsgBox("Are you sure you want to close the application?", vbYesNo + vbQuestion, TITLES) = vbNo Then
        Cancel = 1
    Else
        For Each fm In Forms
            Unload fm
        Next fm
        On Error GoTo 10

        Set currUser = currUser.LogOffUser()
    End If
Exit Sub
10:
    err.Clear
End Sub

Private Sub Form_Resize()

'    If Me.WindowState <> vbMinimized Then
'        Me.WindowState = vbMaximized
'    End If

    If FLoading = True Then
        oSmart.FResize Me
        FLoading = False
        'Get the heights and top of tvwMain tree view and lvwEmp listview
        tvwMainheight = (CLng(Screen.Height - (2 * tvwMain.Top))) + 80
        
        tvwMain.Move tvwMain.Left, tvwMain.Top, tvwMain.Width, tvwMainheight
        fraEmployees.Move fraEmployees.Left, fraEmployees.Top, fraEmployees.Width, tvwMain.Height
        lvwEmp.Move lvwEmp.Left, lvwEmp.Top, lvwEmp.Width, (tvwMain.Height - lvwEmp.Top)
        cboCat.Move cboCat.Left, cboCat.Top + 50, lvwEmp.Width
        cboTerms.Width = lvwEmp.Width
        cboStructure.Width = lvwEmp.Width
        
        tvwMainheight = tvwMainheight + 70
        
        tvwMainTop = tvwMain.Top
        lvwEmpHeight = lvwEmp.Height
    End If

    If oSmart.wRatio > 1.1 Then
        fraLogo.Move ((Me.Width / 2) - (fraLogo.Width / 2)) + 800
    Else
        fraLogo.Move ((Me.Width / 2) - (fraLogo.Width / 2))
        
    End If
End Sub


Public Sub CloseMyWindows()
    Dim fm As Form

    For Each fm In Forms
        If Not fm.Name = "frmMain2" Then
            Unload fm
        End If
    Next fm

End Sub

Public Sub ClosePrompts()
    'hide all the frames that hold the prompts
'    fraBDay.Visible = False
'    fraRet.Visible = False
'    fraCasuals.Visible = False
'    fraVisa.Visible = False
'    fraContracts.Visible = False
'    fraProb.Visible = False
'    fraEmploymentExpiry.Visible = False
End Sub

Public Sub InitTree()
    tvwMain.Nodes.Clear
'
    Set MyNodes = tvwMain.Nodes.add(, , "L", "Personnel Director", "B")
    Set MyNodes = tvwMain.Nodes.add("L", , "S", "SET-UP", "J")
    Set MyNodes = tvwMain.Nodes.add("S", tvwChild, "S80", "Company Details", "Z")
    Set MyNodes = tvwMain.Nodes.add("S", tvwChild, "S6", "Organization Unit Types", "Z")
    Set MyNodes = tvwMain.Nodes.add("S", tvwChild, "S1", "Organization Units", "Z")
    
    Set MyNodes = tvwMain.Nodes.add("S", tvwChild, "currencies", "Currencies", "Z")
    Set MyNodes = tvwMain.Nodes.add("S", tvwChild, "banks", "Banks", "Z")
    Set MyNodes = tvwMain.Nodes.add("S", tvwChild, "staffcategories", "Staff Categories", "Z")
    Set MyNodes = tvwMain.Nodes.add("S", tvwChild, "S15", "Job Positions", "Z")
    Set MyNodes = tvwMain.Nodes.add("S", tvwChild, "S11", "Employee Grades", "Z")
    Set MyNodes = tvwMain.Nodes.add("S", tvwChild, "sectors", "Sectors and Projects", "Z")
    Set MyNodes = tvwMain.Nodes.add("S", tvwChild, "S22", "Employment Terms", "Z")
    Set MyNodes = tvwMain.Nodes.add("S", tvwChild, "jdsetup", "JD Setup", "Z")
    Set MyNodes = tvwMain.Nodes.add("S", tvwChild, "locationssetup", "Locations Setup", "Z")
    
    'Set MyNodes = tvwMain.Nodes.Add("S", tvwChild, "gradetitles", "Grade Titles", "Z")
'    'Set MyNodes = tvwMain.Nodes.Add("S", tvwChild, "S25", "Positions Requirements", "Z")
'    'Set MyNodes = tvwMain.Nodes.Add("S", tvwChild, "S24", "Positions Qualifications", "Z")
'
    'Set MyNodes = tvwMain.Nodes.add("S", tvwChild, "S2", "Contacts Types", "Z")
    Set MyNodes = tvwMain.Nodes.add("S", tvwChild, "S17", "Contract Types", "Z")
    Set MyNodes = tvwMain.Nodes.add("S", tvwChild, "S27", "Casual Types", "Z")
    Set MyNodes = tvwMain.Nodes.add("S", tvwChild, "S4", "Defined Details Types", "Z")
    Set MyNodes = tvwMain.Nodes.add("S", tvwChild, "S3", "Bio-Data Types", "Z")
    
    Set MyNodes = tvwMain.Nodes.add("S", tvwChild, "S41", "Education Courses", "Z")
    Set MyNodes = tvwMain.Nodes.add("S", tvwChild, "S40", "Awards Types", "Z")
    
    'Set MyNodes = tvwMain.Nodes.Add("S", tvwChild, "KS42", "Education Types", "Z")
    
    Set MyNodes = tvwMain.Nodes.add("S", tvwChild, "S12", "Nationalities", "Z")
    Set MyNodes = tvwMain.Nodes.add("S", tvwChild, "S13", "Ethnicity", "Z")
    Set MyNodes = tvwMain.Nodes.add("S", tvwChild, "S18", "Religion", "Z")
'    Set MyNodes = tvwMain.Nodes.Add("S", tvwChild, "S19", "Bank Details", "Z")
'    Set MyNodes = tvwMain.Nodes.Add("S", tvwChild, "S20", "Bank Branch SetUp", "Z")
    'Set MyNodes = tvwMain.Nodes.Add("S", tvwChild, "S14", "Annual Bonus Rates", "Z")
    Set MyNodes = tvwMain.Nodes.add("S", tvwChild, "DisengagementReasons", "Disengagement Reasons", "Z")

'    'Set MyNodes = tvwMain.Nodes.Add("S", tvwChild, "S21", "System Security", "Z")   'Caters for all systtem Security needs
    Set MyNodes = tvwMain.Nodes.add("S", tvwChild, "S23", "Company Assets", "Z")
    Set MyNodes = tvwMain.Nodes.add("S", tvwChild, "SDDF", "Disengagement Dateformat", "Z")
'    Set MyNodes = tvwMain.Nodes.Add("S", tvwChild, "S10", "General Options", "Z")
    
    MyNodes.EnsureVisible
'
'    'Set MyNodes = tvwMain.Nodes.Add("L", , "U", "UTILITIES", "J")
'    'Set MyNodes = tvwMain.Nodes.Add("U", tvwChild, "U3", "Global Posting", "Z")
'    'Set MyNodes = tvwMain.Nodes.Add("U", tvwChild, "U2", "Import Details", "Z")
'
'    'MyNodes.EnsureVisible
'
    Set MyNodes = tvwMain.Nodes.add("L", , "E", "EMPLOYEE", "J")
    Set MyNodes = tvwMain.Nodes.add("E", tvwChild, "E1", "General Details", "Z")
    Set MyNodes = tvwMain.Nodes.add("E", tvwChild, "S5", "Department Employees", "Z")
    Set MyNodes = tvwMain.Nodes.add("E", tvwChild, "E2", "Contact Details", "Z")
    Set MyNodes = tvwMain.Nodes.add("E", tvwChild, "E9", "Defined Details", "Z")
    Set MyNodes = tvwMain.Nodes.add("E", tvwChild, "E3", "Bio-Data", "Z")
'    Set MyNodes = tvwMain.Nodes.Add("E", tvwChild, "E4", "Next of Kin", "Z")
'    Set MyNodes = tvwMain.Nodes.Add("E", tvwChild, "E12", "Family Members", "Z")
    Set MyNodes = tvwMain.Nodes.add("E", tvwChild, "E6", "Contract Details", "Z")
    Set MyNodes = tvwMain.Nodes.add("E", tvwChild, "E18", "Casuals Details", "Z")
    Set MyNodes = tvwMain.Nodes.add("E", tvwChild, "E19", "Assets Issue", "Z")
    Set MyNodes = tvwMain.Nodes.add("E", tvwChild, "E13", "Expatriates Visa Details", "Z")
    Set MyNodes = tvwMain.Nodes.add("E", tvwChild, "E15", "Pre-Employment Checklist", "Z")
    Set MyNodes = tvwMain.Nodes.add("E", tvwChild, "E7", "Employment History", "Z")
    Set MyNodes = tvwMain.Nodes.add("E", tvwChild, "E5", "Referees", "Z")
    Set MyNodes = tvwMain.Nodes.add("E", tvwChild, "E11", "Awards", "Z")
    Set MyNodes = tvwMain.Nodes.add("E", tvwChild, "E8", "Education History", "Z")
    Set MyNodes = tvwMain.Nodes.add("E", tvwChild, "E10", "Professional Qualification", "Z")
    
    
'    Set MyNodes = tvwMain.Nodes.Add("E", tvwChild, "E16", "Employee Archives", "Z")
    
    Set MyNodes = tvwMain.Nodes.add("E", tvwChild, "EmployeeBankAccounts", "Employee Bank Accounts", "Z")
    Set MyNodes = tvwMain.Nodes.add("E", tvwChild, "Disengagement", "Disengage Employee", "Z")
    Set MyNodes = tvwMain.Nodes.add("E", tvwChild, "Reengagement", "Employee Archives", "Z")
    
    MyNodes.EnsureVisible
'
End Sub

Private Sub lvwEmp_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwEmp
        If ColumnHeader.Index = 1 Then
            .ColumnHeaders(1).Width = 1000
        Else
            .ColumnHeaders(1).Width = 500
        End If
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
        End If
    End With

End Sub

Private Sub lvwEmp_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo ErrorHandler

    If new_Record = True Or EmployeeIsInEditMode = True Then
        If MsgBox("You were in the middle of data entry." & vbCrLf & "Please be informed that switching to another" & vbCrLf & "employee will lead to loss of data already entered." & vbCrLf & "Do you still wish to move to the specified operation?", vbYesNo + vbExclamation, "Confirm data loss") = vbNo Then Exit Sub
        new_Record = False
        EmployeeIsInEditMode = False
    End If
        
    'search for selected employee
    Set SelectedEmployee = Nothing

    Set SelectedEmployee = AllEmployees.FindEmployee(CLng(Item.Tag))
    
     If Not (SelectedEmployee Is Nothing) Then
         txtDetails.Caption = "EmpCode: " & SelectedEmployee.EmpCode & " | Name: " & SelectedEmployee.SurName & "" & " " & SelectedEmployee.OtherNames & "" & " " & vbCrLf & _
             "" & "ID No:" & " " & SelectedEmployee.IdNo & " | Date Employed:" & " " & Format(SelectedEmployee.DateOfEmployment, "dd-MMM-yyyy")

         If TheLoadedForm Is Nothing Then Exit Sub

         If TheLoadedForm.Name = "frmEmployee" Or TheLoadedForm.Name = "frmCheck" Or TheLoadedForm.Name = "frmJProg" Or TheLoadedForm.Name = "frmDisEngagement" Or TheLoadedForm.Name = "frmContacts" Or TheLoadedForm.Name = "frmContract" Or TheLoadedForm.Name = "frmBio" Or TheLoadedForm.Name = "frmKin" Or TheLoadedForm.Name = "frmEdu" Or TheLoadedForm.Name = "frmProf" Or TheLoadedForm.Name = "frmEmployment" Or TheLoadedForm.Name = "frmDDetails" Or TheLoadedForm.Name = "frmRef" Or TheLoadedForm.Name = "frmFamily" Or TheLoadedForm.Name = "frmAwards" Or TheLoadedForm.Name = "frmVisa" Or TheLoadedForm.Name = "frmCasuals" Or TheLoadedForm.Name = "frmEmployeeBankAccounts" Or TheLoadedForm.Name = "frmAssetIssue" Then
             TheLoadedForm.DisplayRecords
         End If

         If FViewOnly = False Then
             If TheLoadedForm.Name <> "frmEmployee" Then
                 If SelectedEmployee.IsDisengaged = True Then
                     Call DisableCmd
                 Else
                     Call Decla.EnableCmd
                 End If
             End If
         End If
     End If

    If Not (TheLoadedForm Is Nothing) Then      'code added on 31.10.2006
        If TheLoadedForm.Name = "frmEmployee" Then Call frmEmployee.SwitchEmp
    End If

    Exit Sub
ErrorHandler:
    MsgBox "An error has occurred while retrieving employee info." & vbNewLine & err.Description, vbInformation, TITLES
End Sub
''*************** added by kalya

Private Sub lvwEmp_ItemClick_thisEmployee(selectedemp As Employee)
    On Error GoTo ErrorHandler

    If new_Record = True Or EmployeeIsInEditMode = True Then
        If MsgBox("You were in the middle of data entry." & vbCrLf & "Please be informed that switching to another" & vbCrLf & "employee will lead to loss of data already entered." & vbCrLf & "Do you still wish to move to the specified operation?", vbYesNo + vbExclamation, "Confirm data loss") = vbNo Then Exit Sub
        new_Record = False
        EmployeeIsInEditMode = False
    End If
    
    'search for selected employee
    Set SelectedEmployee = Nothing

    ''Set SelectedEmployee = AllEmployees.FindEmployee(CLng(Item.Tag))
    Set SelectedEmployee = selectedemp
    
     If Not (SelectedEmployee Is Nothing) Then
         txtDetails.Caption = "EmpCode: " & SelectedEmployee.EmpCode & " | Name: " & SelectedEmployee.SurName & "" & " " & SelectedEmployee.OtherNames & "" & " " & vbCrLf & _
             "" & "ID No:" & " " & SelectedEmployee.IdNo & " | Date Employed:" & " " & Format(SelectedEmployee.DateOfEmployment, "dd-MMM-yyyy")

         If TheLoadedForm Is Nothing Then Exit Sub

         If TheLoadedForm.Name = "frmEmployee" Or TheLoadedForm.Name = "frmCheck" Or TheLoadedForm.Name = "frmJProg" Or TheLoadedForm.Name = "frmDisEngagement" Or TheLoadedForm.Name = "frmContacts" Or TheLoadedForm.Name = "frmContract" Or TheLoadedForm.Name = "frmBio" Or TheLoadedForm.Name = "frmKin" Or TheLoadedForm.Name = "frmEdu" Or TheLoadedForm.Name = "frmProf" Or TheLoadedForm.Name = "frmEmployment" Or TheLoadedForm.Name = "frmDDetails" Or TheLoadedForm.Name = "frmRef" Or TheLoadedForm.Name = "frmFamily" Or TheLoadedForm.Name = "frmAwards" Or TheLoadedForm.Name = "frmVisa" Or TheLoadedForm.Name = "frmCasuals" Or TheLoadedForm.Name = "frmEmployeeBankAccounts" Or TheLoadedForm.Name = "frmAssetIssue" Then
             TheLoadedForm.DisplayRecords
         End If

         If FViewOnly = False Then
             If TheLoadedForm.Name <> "frmEmployee" Then
                 If SelectedEmployee.IsDisengaged = True Then
                     Call DisableCmd
                 Else
                     Call Decla.EnableCmd
                 End If
             End If
         End If
     End If

    If Not (TheLoadedForm Is Nothing) Then      'code added on 31.10.2006
        If TheLoadedForm.Name = "frmEmployee" Then Call frmEmployee.SwitchEmp
    End If

    Exit Sub
ErrorHandler:
    MsgBox "An error has occurred while retrieving employee info." & vbNewLine & err.Description, vbInformation, TITLES
End Sub

''end added by kalya
Private Sub mnuAdjustPay_Click()
    'this right is based on the ability to view employee payment details
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("GeneralDetails") = secNone Then
            MsgBox "You Don't have right to access this feature. Please liaise with security admin"
            Exit Sub
        End If
    End If
    
    'Adjust Payments
    frmGlobalPost.Show 1, frmMain2
End Sub

Private Sub mnuArchived_Click()
    Set r = crtAchivedEmployees
    ShowReport r
End Sub

Private Sub mnuBioDataTypes_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("GeneralBioData") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
   
    Set r = crtBioDataTypes
    ShowReport r
End Sub

Private Sub mnuBirthdayByDept_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("BirthdayReports") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    'VIEW BIRTH DAY REPORTS BY DEPARTMENTS

    Set r = crtEmployeeBirthdayByDepartment
    RFilter = "BirthDay"
    frmRange.Show
End Sub

Private Sub mnuBirthdayByEmp_Click()
    
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("BirthdayReports") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    ' VIEW BIRTHDAY REPORT GROUPED BY EMPLOYEE
    Set r = crtEmployeeBirthdayByEmployee
    RFilter = "BirthDay"
    frmRange.Show
End Sub



Private Sub mnuBirthDayPrompts_Click()
    ShowReport crptEmpBirthdaysThisMonth
End Sub

Private Sub mnuBlist_Click()
     If Not (currUser Is Nothing) Then
        If currUser.CheckRight("EmployeeBiodata") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    Set r = crtEmployeeBioData
    RFilter = "General"
    frmRange.Show
  ''  ShowReport R
'    ReportSchemaName = "Employee"
'    ReportType = "Normal"
'    frmHeadcountFilter.Show
End Sub

Private Sub mnuBonusAnnualList_Click()
    
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("BonusReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    Set r = crtBonusList
    RFilter = "Bonus"
    frmRange.Show
End Sub

Private Sub mnuBonusAnnualListNoAwards_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("BonusReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    
    Set r = crtNoBonus
    RFilter = "Bonus"
    frmRange.Show
End Sub

Private Sub mnuCGrowing_Click()
    Set r = crtEmployeeContactsReport
    ReportSchemaName = "EmployeeContacts"
    ReportType = "Normal"
    frmHeadcountFilter.Show
End Sub

Private Sub mnuChangePassword_Click()
    currUser.ChangePassword
End Sub

Private Sub mnuChartEmpsDpt_Click()
    On Error GoTo errHandler
    
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("Pre_EmploymentChecklistReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
    End If
    
    Set r = crptPieChart
  
   ShowReport r
    
    Exit Sub
errHandler:
    MsgBox err.Description, vbExclamation, "Report Error"
End Sub

Private Sub mnuCheckList_Click()
    
     If Not (currUser Is Nothing) Then
        If currUser.CheckRight("Pre_EmploymentChecklistReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    Set r = crtCheck
    
    ''-------inserted by kalya
    
    RFilter = "General"
    myreport = "crtCheck"
    frmRange.Show
    
    ''------end inserted by kalya
    
   '' ShowReport R
'    ChangeHeight = True 'enable the pre-employment check filters
'
'    ReportSchemaName = "Employee"
'    ReportType = "Normal"
'    frmHeadcountFilter.Show

End Sub

Private Sub mnuCompanyAssets_Click()
    
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("CompanyAssetsReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    ShowReport crtCompanyAssets
End Sub

Private Sub mnuContactList_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("Contact") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
    End If
   
    Set r = crtEmployeeContacts
    RFilter = "General"
    frmRange.Show
    ''ShowReport R
'    ReportSchemaName = "Employee"
'    ReportType = "Normal"
'    frmHeadcountFilter.Show
End Sub

Private Sub mnuContactTypes_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("GeneralContactTypes") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    
    ShowReport crtContactTypes
End Sub



Private Sub mnuContractsAlmost_Click()
    ShowReport crptEmpsContractsDue
End Sub

Private Sub mnuDefDetTypes_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("GeneralDefineDetails") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    ShowReport crtDefinedTypes
End Sub

Private Sub mnuDGrowing_Click()
    Set r = crtEmplyeeDefineDetailReport
    ReportSchemaName = "EmployeeDefineDetails"
    ReportType = "Normal"
    frmHeadcountFilter.Show
End Sub

Private Sub mnuDisengagedCount_Click()
    ReportType = "Disengaged"
    ReportSchemaName = "Employee"
    Set r = crtDisengagedHeadCount
    RFilter = "Disangagement Date"
    frmdisangagecriteria.Show
    ''ShowReport R
'    frmHeadcountFilter.Show
End Sub

Private Sub MnuDisList_Click()
     Set r = crtDisengaged
     frmdisangagecriteria.Show vbModal
     RFilter = "Disangagement Date"
    ''ShowReport R
'     ReportSchemaName = "Employee"
'     ReportType = "Disengaged"
'    frmHeadcountFilter.Show
End Sub

Private Sub mnuDivisions_Click()
    Set r = crtDivisions
    ShowReport r
End Sub



Private Sub mnuDlistReport_Click()
    Set r = crtEmplyeeDefineDetailReport
    RFilter = "General"
    frmRange.Show
    ''ShowReport R
End Sub

Private Sub mnudptHierarchy_Click()


    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("AwardReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
    End If
    
    
    
    'VIEW REPORT OF AWARDS AWARDED TO EMPLOYEE
     Set r = crtDptHierarchy
     frmRange2.Show
     ''ShowReport R
End Sub

Private Sub mnudrilldown_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("FilteredReportdrill") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    
frmreportfilter_drill.Show
End Sub

Private Sub mnuDueProbationEnds_Click()
    ShowReport crptEmpsDueFromProbation
End Sub

Private Sub mnuEmpAwards_Click()
    
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("AwardReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    'VIEW REPORT OF AWARDS AWARDED TO EMPLOYEE
    Set r = crtEmployeeAwards
    RFilter = "General"
    frmRange.Show
   '' ShowReport R
'    ReportSchemaName = "Employee"
'    ReportType = "Normal"
'    frmHeadcountFilter.Show
End Sub

Private Sub mnuEmpBankDetails_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("BankDetails") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    
    Set r = crtEmployeeBanks
'    RFilter = "General"
'    frmRange.Show
     ShowReport r
''    ReportSchemaName = "vwEmployeeBankAccounts"
''    ReportType = "Normal"
''    frmHeadcountFilter.Show
End Sub


Private Sub mnuEmpContractDetails_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("ContractDetails") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
   
    Set r = crtEmployeeContracts
    RFilter = "General"
    frmRange.Show
    ''ShowReport R
'    ReportSchemaName = "Employee"
'    ReportType = "Normal"
'    frmHeadcountFilter.Show
'    RFilter = "Standard"
'    frmRange.Show
End Sub

Private Sub mnuEmpDefDetByEmp_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("DefineDetails") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
   
    Set r = crtDefinedDetailsByEmployee
    RFilter = "Standard"
    frmRange.Show
End Sub

Private Sub mnuEmpDefDetByType_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("DefineDetails") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
   
    Set r = crtDefinedDetByType
    RFilter = "Standard"
    frmRange.Show
End Sub


Private Sub mnuEmpEducHistory_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("EducationReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    Set r = crtEmployeeEducationHistory
    RFilter = "General"
    frmRange.Show
    ''ShowReport R
'    ReportSchemaName = "Employee"
'    ReportType = "Normal"
'    frmHeadcountFilter.Show
'    RFilter = "Standard"
'    frmRange.Show
End Sub



Private Sub mnuEmployeeEmpHist_Click()
     If Not (currUser Is Nothing) Then
        If currUser.CheckRight("EmployementHistory") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
   
    Set r = crtEmploymentHistory
    RFilter = "General"
    ''myreport = "chkEmploymentHistory"
    
    frmRange.Show
    ''ShowReport R
'    ReportSchemaName = "Employee"
'    ReportType = "Normal"
'    frmHeadcountFilter.Show
'    RFilter = "Standard"
'    frmRange.Show
End Sub
'
Private Sub mnuEmployeeExpatriateVisaDetails_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("ExpertriateReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    
    Set r = crtEmployeeVisa
    RFilter = "General"
    frmRange.Show
    ''ShowReport R
'    ReportSchemaName = "Employee"
'    ReportType = "Normal"
'    frmHeadcountFilter.Show
End Sub

Private Sub mnuEmployeeFamily_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("familyMember") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    
    Set r = crtEmployeeFamily
    RFilter = "General"
    frmRange.Show
    ''ShowReport R
'    ReportSchemaName = "Employee"
'    ReportType = "Normal"
'    frmHeadcountFilter.Show
'    RFilter = "Standard"
'    frmRange.Show
End Sub

Private Sub mnuEmployeeGenDet_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("GeneralDetails") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
    End If
    
    Set r = crtEmployeeGenDet
    'ShowReport R
    RFilter = "General"
    
    frmRange.Show
End Sub

Private Sub mnuemployeeprojects_Click()
'    Set R = crptEmployeeProgrammes
'    ShowReport R
    FrmEmployeeFilter.Show 1
End Sub

Private Sub mnuEmployeeReferees_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("EmplyeeReferees") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    
    Set r = crtEmployeeReferees
    RFilter = "General"
    frmRange.Show
    ''ShowReport R
'    ReportSchemaName = "Employee"
'    ReportType = "Normal"
'    frmHeadcountFilter.Show
End Sub


Private Sub mnuEmpMedicalRept_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("MedicalReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    Set r = crtEmployeeMedical
    
    
   '' ShowReport R
   ''RFilter = "Standard"
   RFilter = "General"
   frmRange.Show
End Sub

Private Sub mnuEmpNextOfKin_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("HeadCountReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    
    Set r = crtEmployeeNextKin
    RFilter = "NextOFKin"
    frmRange.Show
End Sub

Private Sub mnuempperdatasumsheet_Click()
  If Not (currUser Is Nothing) Then
        If currUser.CheckRight("EMPLOYEEPERSONALDATASUM") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
    End If
    
    
    
    'VIEW REPORT OF AWARDS AWARDED TO EMPLOYEE
     Set r = crptemployeepersonaldatasummary
     r.reportTitle = "EMPLOYEES PERSONAL DATA SUMMARY SHEET"
     r.DisplayProgressDialog = True
    
     ''frmRange2.Show
     ShowReport r
End Sub

Private Sub mnuEmpPhysicallyChallenged_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("MedicalReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    
    Set r = crtDisabledEmployees
   '' ShowReport R
    RFilter = "General"
    frmRange.Show
End Sub


Private Sub mnuEmpProfQualification_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("Professionalqualification") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
   
    Set r = crtEmployeeProfessionalQual
    RFilter = "General"
    frmRange.Show
    'ShowReport R
'    ReportSchemaName = "Employee"
'    ReportType = "Normal"
'    frmHeadcountFilter.Show
'    RFilter = "Standard"
'    frmRange.Show
End Sub


Private Sub mnuEmpProjects_Click()
    Set r = crptEmployeeProgrammes
    RFilter = "General"
    frmRange.Show
    ''ShowReport R
End Sub

Private Sub mnuEmpTerms_Click()
    Set r = ctrEmpTerms
    ShowReport r
End Sub

Private Sub mnuepd_beneficiaries_Click()
  If Not (currUser Is Nothing) Then
        If currUser.CheckRight("EMP_P_D_B") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
    End If
    
    
    
    'VIEW REPORT OF AWARDS AWARDED TO EMPLOYEE
     Set r = crptemployeespersonaldetails_beneficiaries
     r.reportTitle = companyDetail.CompanyName & "- Welfare"
     r.ReportComments = "EMPLOYEE'S PERSONAL DETAILS & BENEFICIARIES"
     ''frmRange2.Show
     ShowReport r
End Sub

Private Sub mnuexemployeeinfo_Click()

    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("EXEMPLOYEEINFO") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
    End If
    
    
    
    'VIEW REPORT OF AWARDS AWARDED TO EMPLOYEE
     Set r = crtexemployeeinfo
     r.reportTitle = "EX EMPLOYEE INFORMATION"
     ''frmexempinfofilter.Show
     ''frmRange2.Showl
     ShowReport r
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFileAdministrator_Click()
    Dim hrmod As HRMSEC.CHRModule
    Set hrmod = New HRMSEC.CHRModule
    hrmod.DisplayEmployeeAccess
End Sub

Private Sub mnuFileAdministratorUserAccounts_Click()
    Dim hrmod As HRMSEC.CHRModule
    Set hrmod = New HRMSEC.CHRModule
    hrmod.DisplayUsers
End Sub

Private Sub mnuG_Click()
    Dim hrmod As HRMSEC.CHRModule
    Set hrmod = New HRMSEC.CHRModule
    hrmod.DisplayGroups
End Sub

'
'Private Sub mnuGen_Click()
'On Error GoTo errHandler
'        Set a = New Application
'        myfile = App.Path & "\Leave Reports\GeneralOptions.rpt"
'        Set R = a.OpenReport(myfile)
'
'      R.ReadRecords
'
'      With frmReports.CRViewer1
'          .ReportSource = R
'          .ViewReport
'      End With
'
'      frmReports.Show vbModal
'     Me.MousePointer = 0
'     Exit Sub
'
'errHandler:
'If Err.Description = "File not found." Then
'    cdl.DialogTitle = "Select the report to show"
'    cdl.InitDir = App.Path & "/Leave Reports"
'    cdl.Filter = "Reports {* .rpt|* .rpt"
'    cdl.ShowOpen
'    myfile = cdl.FileName
'    If Not myfile = "" Then
'        Resume
'    Else
'        Me.MousePointer = 0
'    End If
'Else
'    MsgBox Err.Description, vbInformation
'    Me.MousePointer = 0
'End If
'End Sub
'
'
Private Sub mnuFamily_Click()
    Me.MousePointer = vbHourglass
    Call CloseMyWindows
    ReportType = "Family"
    frmRange.Show , Me
End Sub






Private Sub mnuHcount_Click()
    frmHeadcountFilter.Show
End Sub

Private Sub mnuGrow_Click()
    Set r = crtBioDataReport
    ReportSchemaName = "EmployeeBioData"
    ReportType = "Normal"
    frmHeadcountFilter.Show
End Sub

Private Sub mnuHAnalysis_Click()
'    ReportSchemaName = "Employee"
    Set r = crtHeadCountAnalysist
    RFilter = "General"
    frmRange.Show
    ''ShowReport R
'    ReportHeading = True
'    ReportType = "Normal"
'    ReportSchemaName = "Employee"
'    frmHeadcountFilter.Show
End Sub

Private Sub mnuHDByEmpTerms_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("HeadCountReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    
    Set r = crtHeadCountCasualList
    RFilter = "Standard"
    frmRange.Show
End Sub

Private Sub mnuHDCountByEmpDept_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("HeadCountReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
   
    Set r = crtEmpCountByDeptSumm
    RFilter = "Standard"
    frmRange.Show
End Sub

Private Sub mnuHDCountByEmpTerms_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("HeadCountReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    
    Set r = crtHeadCountByDeptCount
    RFilter = "Standard"
    frmRange.Show
End Sub


Private Sub mnuHDEmpCountDirect_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("HeadCountReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    ReportHeading = True
    Set r = crtHeadCountByPosition
    RFilter = "General"
    frmRange.Show
    ''ShowReport R
'    ReportSchemaName = "Employee"
'    ReportType = "Normal"
'    frmHeadcountFilter.Show
End Sub


Private Sub mnuHDEmpListByDept_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("HeadCountReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    Set r = crtHeadCountByDeptSumm
    RFilter = "General"
    frmRange.Show
End Sub

Private Sub mnuHDEmpListDirect_Click()
    
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("HeadCountReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    TheReport = "crtPositionList"
    'If (isModuleRegistered_Report(TheReport, "Headcount designation " & mnuHDEmpListDirect.Caption) = False) Or (CheckRights(TheReport) = False) Then MsgBox "You have insufficient rights to access" & vbCrLf & "the selected report.", vbExclamation + vbOKOnly, "Limited priviledges": Me.MousePointer = vbNormal: Exit Sub

    Set r = crtPositionList
    r.reportTitle = "Desigination List Report"
    RFilter = "Standard"
    frmRange.Show
End Sub



Private Sub mnuHeadcountCategGradesByDeptCount_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("HeadCountReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    
    Set r = crtGradeEmployeeCountByDept
    RFilter = "Standard"
    frmRange.Show
End Sub


Private Sub mnuHeadcountCategGradesByDeptList_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("HeadCountReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    
    Set r = crtGradeEmployeeByDept
    RFilter = "Standard"
    frmRange.Show 1
End Sub

Private Sub mnuHeadcountCategGradesEmpCount_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("HeadCountReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    
    Set r = crtGradesEmployeesCount
    RFilter = "Standard"
    frmRange.Show
End Sub

Private Sub mnuHeadcountCategGradesEmpList_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("HeadCountReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    Set r = crtGradesEmployees
    RFilter = "General"
    frmRange.Show
    ''ShowReport R
End Sub

Private Sub mnuHeadcountCategGradesList_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("HeadCountReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    ''Set r = crtGradesList
    Set r = crptEmployeeCategories

    ShowReport r
End Sub

Private Sub mnuHeadCountGenderDeptCount_Click()
    Set r = crtGenderCountByDept
    
    ReportHeading = True
    RFilter = "Standard"
    frmRange.Show
End Sub

Private Sub mnuHeadCountGenderDeptList_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("HeadCountReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
   
    Set r = crtGenderListByDept
    RFilter = "Standard"
    frmRange.Show
End Sub

Private Sub mnuHeadCountGenderEmployeeCount_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("HeadCountReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    ReportHeading = True
    
    Set r = crtGenderCount
    RFilter = "Standard"
    frmRange.Show
End Sub

Private Sub mnuHeadCountGenderEmployeeList_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("HeadCountReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    Set r = crtGender
    RFilter = "Standard"
    frmRange.Show
End Sub

Private Sub mnuHeadCountNationalityByDeptCount_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("HeadCountReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    ReportHeading = True
    
    Set r = crtNationalityByDeptSumm
    RFilter = "Standard"
    frmRange.Show
End Sub

Private Sub mnuHeadCountNationalityByDeptList_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("HeadCountReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
   
    Set r = crtNationalityByDept
    RFilter = "Standard"
    frmRange.Show
End Sub

Private Sub mnuHeadCountNationalityCount_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("HeadCountReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    ReportHeading = True
    Set r = crtNationalityListSumm
    RFilter = "Standard"
    frmRange.Show
End Sub


Private Sub mnuHeadCountNationalityList_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("HeadCountReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
    End If
    
    ShowReport crtNationalities
End Sub

Private Sub mnuHeadCountTribesDeptCount_Click()
'    TheReport = "crtEmployeeTribesByDeptCount"
'    If (isModuleRegistered_Report(TheReport, "Headcount tribes by department " & mnuHeadCountTribesDeptCount.Caption) = False) Or (CheckRights(TheReport) = False) Then MsgBox "You have insufficient rights to access" & vbCrLf & "the selected report.", vbExclamation + vbOKOnly, "Limited priviledges": Me.MousePointer = vbNormal: Exit Sub
    ReportHeading = True
    Set r = crtEmployeeTribesByDeptCount
    RFilter = "Standard"
    frmRange.Show
End Sub

Private Sub mnuHeadCountTribesDeptList_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("HeadCountReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
   
    Set r = crtEmployeeTribesListByDept
    RFilter = "Standard"
    frmRange.Show
End Sub

Private Sub mnuHeadCountTribesEmployeeCount_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("HeadCountReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    Set r = crtEmployeeTribesSumm
    ReportHeading = True
    
    RFilter = "Standard"
    frmRange.Show
End Sub

Private Sub mnuHeadCountTribesEmployeeList_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("HeadCountReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    Set r = crtEmployeeTribes
    RFilter = "Standard"
    frmRange.Show
End Sub

Private Sub mnuHeadCountTribesList_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("HeadCountReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
   
    ShowReport crtTribes
End Sub

Private Sub mnuHelpAboutSoftware_Click()
DisplayModuleVersionInfo companyDetail.CompanyName, CDate(App.FileDescription)

End Sub
'Public Sub DisplayModuleVersionInfo(ByVal strCompanyName As String, ByVal dtFileCreationDate As Date)
'    Dim Buffer As String * 512
'    Dim BufferSize As Long
'    On Error GoTo ErrorHandler
'
'    With frmModuleVersion
'        .lblModuleName.Caption = ": " & "Personel Director"
'        .lblFileCreationDate.Caption = ": " & Format$(dtFileCreationDate, "dddd dd mmmm,yyyy")
'        .lblModuleVersion.Caption = ": " & App.Major & "." & App.Minor & "." & App.Revision
'        .lblEMailAddress.Caption = "support@infiniti.co.ke"
'        .lblEMailAddress.ForeColor = vbBlue
'        'GETTING THE MACHINE NAME USING GetComputerName API FUNCTION
'        BufferSize = Len(Buffer)
'        If GetComputerName(Buffer, BufferSize) Then
'            .lblMachineName.Caption = Left$(Buffer, BufferSize)
'        Else
'            .lblMachineName.Caption = vbNullString
'        End If
'        .lblcompanyname.Caption = strCompanyName
'        .Show vbModal
'    End With
'    Exit Sub
'
'ErrorHandler:
'    MsgBox "An Error has occurred while attempting to display the software module version" & vbCrLf & err.Description, vbExclamation
'End Sub


Private Sub mnuhistory_Click()
    Set r = crtReengagementHistory
    RFilter = "re-angagement"
    frmdisangagecriteria.Show
    ''frmRange.Show
    ''ShowReport R
End Sub

Private Sub mnuIssuedAssets_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("IssuedAssets") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
  
    Set r = crtAssetIssue
    RFilter = "General"
    frmRange.Show
    ''ShowReport R
'    ReportSchemaName = "Employee"
'    ReportType = "Normal"
'    frmHeadcountFilter.Show
End Sub

Private Sub mnuLogout_Click()
    'Me.Hide
    If Not (currUser Is Nothing) Then
            If Not (currUser Is Nothing) Then
            Set currUser = currUser.LogOffUser
            End If
    End If
    Call LogIn
    'Me.Show
    Call EnableCmd
End Sub


Private Sub mnuMA_Click()
    Dim hrmod As HRMSEC.CHRModule
    Set hrmod = New HRMSEC.CHRModule
    hrmod.DisplayModuleAccess
End Sub

Private Sub mnuLongServiceByDept_Click()
    
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("LongService") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
   
    Set r = crtLongServicePerDept
    r.EnableParameterPrompting = False
    RFilter = "LongService"
    frmRange.Show
End Sub

Private Sub mnuLongServiceGroupingByYears_Click()
    
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("LongService") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
   
    Set r = crtLongServiceGrouping
    r.EnableParameterPrompting = False
    RFilter = "LongService"
    frmRange.Show
End Sub

Private Sub mnuLongServiceList_Click()
    
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("LongService") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    Set r = crtLongServiceList
    
    RFilter = "LongService"
    frmRange.Show
End Sub

Private Sub mnuMissingMandatory_Click()

    Set r = crtMissingMandatory
    RFilter = "Mandatory"
    frmRange.Show
End Sub
'
Private Sub mnuNationalityEmpList_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("HeadCountReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
   
    Set r = crtNationalityList
    RFilter = "Standard"
    frmRange.Show
End Sub

Private Sub mnuOrganizationAssets_Click()
    Set r = crtOrganizationAssets
    ShowReport r
End Sub

Private Sub mnuphotos_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("setupphoto") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
    End If
    

   frmphotosetup.Show vbModal
   
End Sub

Private Sub mnuPhyAddresses_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("GeneralDetails") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
    End If
    
    Set r = crptPhysicalAddresses
'    RFilter = "PhyAddresses"
    RFilter = "General"
    frmRange.Show
   '' ShowReport R
End Sub

Private Sub mnuPosition_Click()
    Set r = crtPositions
    ShowReport r
End Sub

Private Sub mnuProjectFunding_Click()
    Set r = crptProgrammeFunding
    ShowReport r
End Sub

Private Sub mnuprojects_Click()
    Set r = crptProgrammes
    ShowReport r
End Sub

Private Sub mnuReengaded_Click()
    'Display disengaged Employees
    Set r = crtReengagedEmployees
    RFilter = "rengagementhist"
    frmdisangagecriteria.Show
'    frmHeadcountFilter.Show
    ''ShowReport R
End Sub

Private Sub mnuReligion_Click()
    Set r = crtReligions
    ShowReport r
End Sub

Private Sub mnureportperfilter_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("FilteredReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    
frmreportfilter.Show
End Sub

Private Sub mnuRetirementActualDates_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("RetirementReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
    End If
   
    Set r = crtRetirementByDeptActual
    RFilter = "Retirement"
    frmRange.Show
End Sub

Private Sub mnuRetirementByDept_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("RetirementReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    Set r = crtRetirementByDept
    RFilter = "Retirement"
    frmRange.Show
End Sub


Private Sub mnuRetirementByDeptWithoutSal_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("RetirementReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
   
    Set r = crtRetirementByDeptSalaries
    RFilter = "Retirement"
    frmRange.Show
End Sub



Private Sub mnuRetirementByYearGrouping_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("RetirementReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
   
    Set r = crtRetirementByYears
    RFilter = "Retirement"
    frmRange.Show
End Sub

Private Sub mnuRetirementEarly_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("RetirementReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    
    Set r = crtEarlyRetirees
    RFilter = "Retirement"
    frmRange.Show
End Sub

Private Sub mnuRetirementPrompt_Click()
    ShowReport crptEmpsDueForRetire
End Sub

Private Sub mnuRetirementSchedule_Click()
    
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("RetirementReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    Set r = crtRetirement
    RFilter = "Retirement"
    frmRange.Show
End Sub

Private Sub mnurptBanks_Click()
   
    ' crtBanksReport
    ShowReport crtBanksReport
End Sub

Private Sub mnurptBranches_Click()
    ShowReport crtBankBranches
End Sub

Private Sub mnurptLocations_Click()
    'EVALUATE WHETHER THE EMPLOYEE HAS VIEW PRIVILEDGES TO ACCESS THIS REPORT
    ShowReport crptLocations
End Sub

Private Sub mnurptSectors_Click()
    'EVALUATE WHETHER THE EMPLOYEE HAS VIEW PRIVILEDGES TO ACCESS THIS REPORT
    ReDim objParamField(1 To 1)
    objParamField(1).Name = "@SectorID"
    objParamField(1).value = "NULL"
    ShowReport crptSectors, , True
End Sub

Private Sub mnuSalaryInfoByDept_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("SalaryDetailReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
    End If

    Set r = crtSalaryDetailsByDept
    RFilter = "Standard"
    frmRange.Show
End Sub

Private Sub mnuSalaryInfoIncrementList_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("SalaryDetailReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
    End If
    
    Set r = crtSalaryIncrementList
    RFilter = "Standard"
    frmRange.Show
End Sub

'Private Sub mnuSalaryInfoIncrementListByDept_Click()
'TheReport = "crtSalaryIncrementByDept"
'If (isModuleRegistered_Report(TheReport, "Salary increment " & mnuSalaryInfoIncrementListByDept.Caption) = False) Or (CheckRights(TheReport) = False) Then MsgBox "You have insufficient rights to access" & vbCrLf & "the selected report.", vbExclamation + vbOKOnly, "Limited priviledges": Me.MousePointer = vbNormal: Exit Sub
'
'Set R = crtSalaryIncrementByDept
'RFilter = "Standard"
'frmRange.Show
'End Sub

Private Sub mnuSalaryInfoSalaryDetails_Click()
    
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("SalaryDetailReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    Set r = crtSalaryDetails
    r.reportTitle = "Salary Details Report"
    RFilter = "Standard"
    frmRange.Show
End Sub

Private Sub mnuTurnOverHires_Click()
    
     If Not (currUser Is Nothing) Then
        If currUser.CheckRight("TurnverReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
   
    Set r = crtTurnOverHires
    RFilter = "TurnOver"
    frmRange.Show
End Sub

Private Sub mnuTurnOverHiresYearToDate_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("TurnverReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
   
    Set r = crtTurnOverHiresToDate
    RFilter = "Hires"
    frmRange.Show
End Sub

Private Sub mnuTurnOverLeft_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("TurnverReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
   
    Set r = crtTurnOverLeft
    RFilter = "TurnOver"
    frmRange.Show
End Sub

Private Sub mnuTurnOverLeftYearToDate_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("TurnverReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
   
    Set r = crtTurnOverLeftToDate
    RFilter = "TurnOver"
    frmRange.Show
End Sub

Private Sub mnuU_Click()
    Dim hrmod As HRMSEC.CHRModule
    Set hrmod = New HRMSEC.CHRModule
    hrmod.DisplayUsers
End Sub

Private Sub pifform_Click()
  If Not (currUser Is Nothing) Then
        If currUser.CheckRight("PIFORM") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
    End If
    
    
    
    'VIEW REPORT OF AWARDS AWARDED TO EMPLOYEE
     Set r = crptPersonelInformationForm
     r.reportTitle = "PERSONEL INFORMATION FORM"
     r.ReportComments = "EMPLOYEE'S PERSONAL INFORMATION & BENEFICIARIES"
     ''frmRange2.Show
     ShowReport r
End Sub

Private Sub tlbPD_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    On Error GoTo ErrorHandler
    Select Case ButtonMenu.Key
'    Case "EJN"
'        MsgBox "There are no job expiry Notices", vbOKOnly + vbInformation, "Job Notice Prompts"
    Case "EDCP"
        ProbPrompt (0)
        fraProb.Left = 3000
        fraProb.Top = 1800
    Case "TRM"
        EmployeeTerminatePrompt (0)
        FraTerminate.Left = 3500
        FraTerminate.Top = 1800
    Case "EDCE"
        ContractExpirePrompt (0)
        fraVisa.Left = 3000
        fraVisa.Top = 1800
    Case "EDTR"
        RetirementPrompt (0)
        fraContracts.Left = 3000
        fraContracts.Top = 1800
    Case "BDY"
        GetBirthDays (0)
        FraBirthDays.Left = 3000
        FraBirthDays.Top = 1800
    Case "PS"
        frmPromptsSetup.Show 1
    End Select
    
Exit Sub
ErrorHandler:
    MsgBox err.Description, vbCritical + vbOKOnly, "Error"
End Sub

'''Private Sub tvwMain_NodeClick(ByVal Node As MSComctlLib.Node)
'''    On Error GoTo errHandler
'''
'''    If fraLog.Visible = True Then
'''        Exit Sub
'''    End If
'''
'''    Dim i As Object
'''
'''    frmCoDetails.Visible = False
'''
'''    '==HIDE PROMPTS WHEN NAVIGATING
'''    fraProb.Visible = False
'''    fraEmployees.Visible = False
'''    FraTerminate.Visible = False
'''    FraBirthDays.Visible = False
'''    '==END HIDE PROMPTS
'''
'''    fraProb.Visible = False
'''    fracmd.Visible = False
'''    fraContracts.Visible = False
'''    fraVisa.Visible = False
'''
'''    'make sure cmdNew is enabled
'''    Me.cmdNew.Enabled = True
'''
'''    Dim mm As String
'''
'''    For Each i In Me
'''        If TypeOf i Is CommandButton Then
'''            If i.Name = "cmdCancel" Or i.Name = "cmdSave" Then
'''                i.Enabled = False
'''            Else
'''                i.Enabled = True
'''            End If
'''        End If
'''    Next i
'''
'''    Call Disabblepromt
'''    Call CloseMyWindows
'''    Call ClosePrompts
'''
'''    'added on January 2007
''''    rsglob variable store all employees who are not terminated
'''
'''    If rsGlob Is Nothing Then
'''        Set rsGlob = CConnect.GetRecordSet("SELECT * FROM pvwrsglob")
'''    End If
'''
'''
'''    'if it happens that this listview is not visible, all procedures which are depending on it will  fail
'''    ' badly!!
'''
'''    If lvwEmp.Visible = False Then lvwEmp.Visible = True
'''
'''    Select Case Node.Key
'''        Case "sectors"
'''           txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmSectors
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''
'''
'''        Case "staffcategories"
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmCSSCategories
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''
'''         Case "gradetitles"
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmGradeTitles
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''
'''        Case "jdsetup"
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmJDCategories
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''
'''        Case "locationssetup"
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmLocations
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''
'''        Case "currencies"
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmCurrencies
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''
'''        Case "banks"
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmBanks
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''
'''        Case "Reengagement"
'''            fraEmployees.Visible = False
'''            lvwEmp.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''            'load the form
'''            DisplayTheForm frmReengageMent
'''
'''            lblEmpList.Visible = True
'''            lblECount.Visible = True
'''            ASearch = True
'''
'''        Case "Disengagement"
'''            fraEmployees.Visible = True
'''            lvwEmp.Visible = True
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmDisEngagement
'''
'''            lblEmpList.Visible = True
'''            lblECount.Visible = True
'''            ASearch = True
'''
'''        Case "E1"
'''
'''        '****************************************
'''
'''            If (objEmployeeBankAccounts Is Nothing) Then
'''             Set objEmployeeBankAccounts = New EmployeeBankAccounts2
'''             objEmployeeBankAccounts.GetActiveEmployeeBankAccounts
'''             Set myEmployeeBankAccounts = objEmployeeBankAccounts
'''             ''Set pEmployeeBankAccounts = objEmployeeBankAccounts
'''            End If
'''
'''        '***************************************
'''
'''
'''            fraEmployees.Visible = True
'''            lvwEmp.Visible = True
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmEmployee
'''
'''            lblEmpList.Visible = True
'''            lblECount.Visible = True
'''            ASearch = True
'''
'''        Case "E15"
'''            fraEmployees.Visible = True
'''            lvwEmp.SetFocus
'''            Me.MousePointer = vbHourglass
'''            DisplayTheForm frmCheck
'''            lblEmpList.Visible = True
'''            lblECount.Visible = True
'''            ASearch = True
'''
'''        Case "E18"
'''            fraEmployees.Visible = True
'''            lvwEmp.SetFocus
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmCasuals
'''            lblEmpList.Visible = True
'''            lblECount.Visible = True
'''            ASearch = True
''''            cboTerms.Text = "Casual"
'''            Call cboTerms_Click
'''
'''        Case "E19"
'''            fraEmployees.Visible = True
'''            lvwEmp.SetFocus
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmAssetIssue
'''
'''            lblEmpList.Visible = True
'''            lblECount.Visible = True
'''            ASearch = True
'''
'''        Case "EmployeeBankAccounts"
'''            fraEmployees.Visible = True
'''
'''            If (objEmployeeBankAccounts Is Nothing) Then
'''             Set objEmployeeBankAccounts = New EmployeeBankAccounts2
'''             objEmployeeBankAccounts.GetActiveEmployeeBankAccounts
'''             Set myEmployeeBankAccounts = objEmployeeBankAccounts
'''            '' Set pEmployeeBankAccounts = objEmployeeBankAccounts
'''            End If
'''
'''            lvwEmp.SetFocus
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmEmployeeBankAccounts
'''
'''            lblEmpList.Visible = True
'''            lblECount.Visible = True
'''            ASearch = True
'''
'''        Case "E2"
'''            fraEmployees.Visible = True
'''            lvwEmp.SetFocus
'''            Me.MousePointer = vbHourglass
'''
'''
'''            DisplayTheForm frmContacts
'''
'''            lblEmpList.Visible = True
'''            lblECount.Visible = True
'''            ASearch = True
'''
'''        Case "E3"
'''            fraEmployees.Visible = True
'''            lvwEmp.SetFocus
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmBio
'''
'''            lblEmpList.Visible = True
'''            lblECount.Visible = True
'''            ASearch = True
'''
'''        Case "E4"
'''            fraEmployees.Visible = True
'''            lvwEmp.SetFocus
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmKin
'''
'''            lblEmpList.Visible = True
'''            lblECount.Visible = True
'''            ASearch = True
'''
'''        Case "E12"
'''            fraEmployees.Visible = True
'''            lvwEmp.SetFocus
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmFamily
'''
'''            lblEmpList.Visible = True
'''            lblECount.Visible = True
'''            ASearch = True
'''
'''        Case "E5"
'''            fraEmployees.Visible = True
'''            lvwEmp.SetFocus
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmRef
'''
'''            lblEmpList.Visible = True
'''            lblECount.Visible = True
'''            ASearch = True
'''
'''
'''        Case "E6"
'''            fraEmployees.Visible = True
'''            lvwEmp.SetFocus
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmContract
'''
'''            lblEmpList.Visible = True
'''            lblECount.Visible = True
'''            ASearch = True
'''
'''        Case "E13"
'''            fraEmployees.Visible = True
'''            lvwEmp.SetFocus
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmVisa
'''
'''            lblEmpList.Visible = True
'''            lblECount.Visible = True
'''            ASearch = True
'''
'''        Case "E7"
'''            fraEmployees.Visible = True
'''            lvwEmp.SetFocus
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmEmployment
'''
'''            lblEmpList.Visible = True
'''            lblECount.Visible = True
'''            ASearch = True
'''
'''        Case "E8"
'''            fraEmployees.Visible = True
'''            lvwEmp.SetFocus
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmEdu
'''
'''            lblEmpList.Visible = True
'''            ASearch = True
'''            lblECount.Visible = True
''''
'''        Case "E9"
'''            fraEmployees.Visible = True
'''            lvwEmp.SetFocus
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmDDetails
'''
'''            lblEmpList.Visible = True
'''            lblECount.Visible = True
'''            ASearch = True
'''
''''            Professional Qualifications
'''        Case "E10"
'''            fraEmployees.Visible = True
'''            lvwEmp.SetFocus
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmProf
'''
'''            lblEmpList.Visible = True
'''            lblECount.Visible = True
'''            ASearch = True
''''
'''''           Employee Awards
'''
'''        Case "E11"
'''            fraEmployees.Visible = True
'''            lvwEmp.SetFocus
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmAwards
'''
'''            lblEmpList.Visible = True
'''            lblECount.Visible = True
'''            ASearch = True
'''
'''        Case "S80"
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmCompanyDetails
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''
'''        Case "DisengagementReasons"
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmDisengagementReasons
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''
'''        Case "S1"
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''
'''            DisplayTheForm frmOUs
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''
'''        Case "S2"
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmCTypes
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''
'''        Case "S3"
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmBTypes
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''
'''        Case "S4"
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmDTypes
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''
'''        Case "S12"
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmNationalities
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''
'''        Case "S18"
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmReligions
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''
'''        Case "S19"
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmBankSetUp
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''
'''        Case "S20"
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''
'''            DisplayTheForm frmBankBranchSetUp
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''
'''        Case "S40"
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmEducationAwardsSetup
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''
'''        Case "S41"
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmEducationCourses
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''
'''            Case "KS42"
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmEducationTypes
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
''''KS42
''''        Case "S21"
''''            txtDetails.Caption = ""
''''            fraEmployees.Visible = False
''''            Me.MousePointer = vbHourglass
''''
''''            '++Check For User Rights ++
''''            'check_User_Rights frmSecurity
''''
''''            lblEmpList.Visible = False
''''            lblECount.Visible = False
'''        Case "S23"
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmCompanyAssets
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''
''''        Case "S24"
''''            txtDetails.Caption = ""
''''            fraEmployees.Visible = False
''''            Me.MousePointer = vbHourglass
''''
''''            '++Check For User Rights ++
''''            'check_User_Rights frmPositionsQualifications
''''
''''            DisplayTheForm frmPositionsQualifications
''''
''''            lblEmpList.Visible = False
''''            lblECount.Visible = False
''''
''''         Case "S25"
''''            txtDetails.Caption = ""
''''            fraEmployees.Visible = False
''''            Me.MousePointer = vbHourglass
''''
''''            '++Check For User Rights ++
''''            'check_User_Rights frmPositionRequirements
''''
''''            DisplayTheForm frmPositionRequirements
''''
''''            lblEmpList.Visible = False
''''            lblECount.Visible = False
'''        Case "S14"
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmBonus
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''
'''        Case "S13"
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmTribes
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
''''
'''        Case "S5"
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''            '++Check For User Rights ++
'''            'check_User_Rights frmCEStructure
'''
'''            DisplayTheForm frmOUEmployees
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''            lblEmpC.Visible = True
'''            lblEmpCount.Visible = True
''''
'''        Case "S6"
'''            'display the form for setting up the OU Types
'''            fraEmployees.Visible = False
'''            'Disable the New Button
'''            Me.cmdNew.Enabled = False
'''            DisplayTheForm frmOUTypes
'''
'''
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''
'''        Case "S15"
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = True
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmPositions
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''
''''        Case "S16"
''''            txtDetails.Caption = ""
''''            fraEmployees.Visible = True
''''            Me.MousePointer = vbHourglass
''''
''''            '++Check For User Rights ++
''''            'check_User_Rights frmPApproval
''''
''''            DisplayTheForm frmPApproval
''''
''''            lblEmpList.Visible = False
''''            lblECount.Visible = False
''''
''''
''''        Case "S8"
''''            txtDetails.Caption = ""
''''            fraEmployees.Visible = True
''''            Me.MousePointer = vbHourglass
''''
''''            '++Check For User Rights ++
''''            'check_User_Rights frmUGroups
''''
''''            DisplayTheForm frmUGroups
''''
''''
''''            lblEmpList.Visible = False
''''            lblECount.Visible = False
''''
''''        Case "S9"
''''            txtDetails.Caption = ""
''''            fraEmployees.Visible = False
''''            Me.MousePointer = vbHourglass
''''
''''            '++Check For User Rights ++
''''            'check_User_Rights frmSystUsers
''''
''''            DisplayTheForm frmSystUsers
''''
''''            lblEmpList.Visible = False
''''            lblECount.Visible = False
''''
'''        Case "S10"
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmGenOpt
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''
'''        Case "S11"
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''
'''            'DisplayTheForm frmEmpCategories
'''            DisplayTheForm frmGradeTitles
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''
'''        Case "S17"
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmContractTypes
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''        Case "SDDF"
'''
'''                DisplayTheForm frmdisengagementdateformat
'''        Case "S22"
'''
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''            '++Check For User Rights ++
'''            'check_User_Rights frmEmploymentTerms
'''
'''            DisplayTheForm frmEmploymentTerms
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''        Case "S27"
'''
'''            txtDetails.Caption = ""
'''            fraEmployees.Visible = False
'''            Me.MousePointer = vbHourglass
'''
'''            DisplayTheForm frmCasualTypes
'''
'''
'''            lblEmpList.Visible = False
'''            lblECount.Visible = False
'''
'''        'setup node
'''        Case "S"
'''            txtDetails.Caption = ""
'''            frmCoDetails.Visible = True
'''
''''        'Utilities node
''''        Case "U"
''''            txtDetails.Caption = ""
'''
'''        'this is the personnel director node
'''        Case "L"
'''            txtDetails.Caption = ""
'''            frmCoDetails.Visible = True
'''
'''        'employee node
'''        Case "E"
'''            txtDetails.Caption = ""
'''            frmCoDetails.Visible = True
'''
'''    End Select
'''
'''    Me.MousePointer = 0
'''    Call Disabblepromt
'''    Exit Sub
'''
'''errHandler:
'''    MsgBox err.Description, vbExclamation
'''    Call CloseMyWindows
'''End Sub
Private Sub tvwMain_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error GoTo errHandler
    
    If fraLog.Visible = True Then
        Exit Sub
    End If
    
    Dim i As Object
    
    frmCoDetails.Visible = False

    '==HIDE PROMPTS WHEN NAVIGATING
    fraProb.Visible = False
    fraEmployees.Visible = False
    FraTerminate.Visible = False
    FraBirthDays.Visible = False
    '==END HIDE PROMPTS
    
    fraProb.Visible = False
    fracmd.Visible = False
    fraContracts.Visible = False
    fraVisa.Visible = False
    
    'make sure cmdNew is enabled
    Me.cmdNew.Enabled = True

    Dim mm As String

    For Each i In Me
        If TypeOf i Is CommandButton Then
            If i.Name = "cmdCancel" Or i.Name = "cmdSave" Then
                i.Enabled = False
            Else
                i.Enabled = True
            End If
        End If
    Next i
    
'    Call Disabblepromt
    Call CloseMyWindows
    Call ClosePrompts
    
    'added on January 2007
'    rsglob variable store all employees whof are not terminated

    If rsGlob Is Nothing Then
        Set rsGlob = CConnect.GetRecordSet("SELECT * FROM pvwrsglob")
    End If
    
    'if it happens that this listview is not visible, all procedures which are depending on it will  fail
    ' badly!!
    
    If lvwEmp.Visible = False Then lvwEmp.Visible = True
    
    Select Case Node.Key
        Case "sectors"
           txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass

            DisplayTheForm frmSectors

            lblEmpList.Visible = False
            lblECount.Visible = False
               
        
        Case "staffcategories"
            txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass

            DisplayTheForm frmCSSCategories

            lblEmpList.Visible = False
            lblECount.Visible = False
        
         Case "gradetitles"
            txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass

            DisplayTheForm frmGradeTitles

            lblEmpList.Visible = False
            lblECount.Visible = False
        
        Case "jdsetup"
            txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass

            DisplayTheForm frmJDCategories

            lblEmpList.Visible = False
            lblECount.Visible = False
        
        Case "locationssetup"
            txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass

            DisplayTheForm frmLocations

            lblEmpList.Visible = False
            lblECount.Visible = False
        
        Case "currencies"
            txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass

            DisplayTheForm frmCurrencies

            lblEmpList.Visible = False
            lblECount.Visible = False
        
        Case "banks"
            txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass

            DisplayTheForm frmBanks
            lblEmpList.Visible = False
            lblECount.Visible = False
        
        Case "Reengagement"
            fraEmployees.Visible = False
            lvwEmp.Visible = False
            Me.MousePointer = vbHourglass
            
            'load the form
            DisplayTheForm frmReengageMent

            lblEmpList.Visible = True
            lblECount.Visible = True
            ASearch = True
            
        Case "Disengagement"
            fraEmployees.Visible = True
            lvwEmp.Visible = True
            Me.MousePointer = vbHourglass

            DisplayTheForm frmDisEngagement

            lblEmpList.Visible = True
            lblECount.Visible = True
            ASearch = True
            
        Case "E1"
        
        '****************************************
        
            If (objEmployeeBankAccounts Is Nothing) Then
             Set objEmployeeBankAccounts = New EmployeeBankAccounts2
             objEmployeeBankAccounts.GetActiveEmployeeBankAccounts
             Set myEmployeeBankAccounts = objEmployeeBankAccounts
             ''Set pEmployeeBankAccounts = objEmployeeBankAccounts
            End If
        
        '***************************************
        
        
            fraEmployees.Visible = True
            lvwEmp.Visible = True
            Me.MousePointer = vbHourglass

            DisplayTheForm frmEmployee

            lblEmpList.Visible = True
            lblECount.Visible = True
            ASearch = True

        Case "E15"
            fraEmployees.Visible = True
            lvwEmp.SetFocus
            Me.MousePointer = vbHourglass
            DisplayTheForm frmCheck
            lblEmpList.Visible = True
            lblECount.Visible = True
            ASearch = True

        Case "E18"
            fraEmployees.Visible = True
            lvwEmp.SetFocus
            Me.MousePointer = vbHourglass
            
            DisplayTheForm frmCasuals
            lblEmpList.Visible = True
            lblECount.Visible = True
            ASearch = True
'            cboTerms.Text = "Casual"
            Call cboTerms_Click
            
        Case "E19"
            fraEmployees.Visible = True
            lvwEmp.SetFocus
            Me.MousePointer = vbHourglass

            DisplayTheForm frmAssetIssue

            lblEmpList.Visible = True
            lblECount.Visible = True
            ASearch = True
        
        Case "EmployeeBankAccounts"
            fraEmployees.Visible = True
            
            If (objEmployeeBankAccounts Is Nothing) Then
             Set objEmployeeBankAccounts = New EmployeeBankAccounts2
             objEmployeeBankAccounts.GetActiveEmployeeBankAccounts
             Set myEmployeeBankAccounts = objEmployeeBankAccounts
            '' Set pEmployeeBankAccounts = objEmployeeBankAccounts
            End If
            
            lvwEmp.SetFocus
            Me.MousePointer = vbHourglass

            DisplayTheForm frmEmployeeBankAccounts

            lblEmpList.Visible = True
            lblECount.Visible = True
            ASearch = True
            
        Case "E2"
            fraEmployees.Visible = True
            lvwEmp.SetFocus
            Me.MousePointer = vbHourglass


            DisplayTheForm frmContacts

            lblEmpList.Visible = True
            lblECount.Visible = True
            ASearch = True

        Case "E3"
            fraEmployees.Visible = True
            lvwEmp.SetFocus
            Me.MousePointer = vbHourglass

            DisplayTheForm frmBio

            lblEmpList.Visible = True
            lblECount.Visible = True
            ASearch = True

        Case "E4"
            fraEmployees.Visible = True
            lvwEmp.SetFocus
            Me.MousePointer = vbHourglass

            DisplayTheForm frmKin

            lblEmpList.Visible = True
            lblECount.Visible = True
            ASearch = True

        Case "E12"
            fraEmployees.Visible = True
            lvwEmp.SetFocus
            Me.MousePointer = vbHourglass

            DisplayTheForm frmFamily

            lblEmpList.Visible = True
            lblECount.Visible = True
            ASearch = True

        Case "E5"
            fraEmployees.Visible = True
            lvwEmp.SetFocus
            Me.MousePointer = vbHourglass

            DisplayTheForm frmRef

            lblEmpList.Visible = True
            lblECount.Visible = True
            ASearch = True


        Case "E6"
            fraEmployees.Visible = True
            lvwEmp.SetFocus
            Me.MousePointer = vbHourglass
            
            DisplayTheForm frmContract

            lblEmpList.Visible = True
            lblECount.Visible = True
            ASearch = True

        Case "E13"
            fraEmployees.Visible = True
            lvwEmp.SetFocus
            Me.MousePointer = vbHourglass

            DisplayTheForm frmVisa

            lblEmpList.Visible = True
            lblECount.Visible = True
            ASearch = True

        Case "E7"
            fraEmployees.Visible = True
            lvwEmp.SetFocus
            Me.MousePointer = vbHourglass

            DisplayTheForm frmEmployment

            lblEmpList.Visible = True
            lblECount.Visible = True
            ASearch = True

        Case "E8"
            fraEmployees.Visible = True
            lvwEmp.SetFocus
            Me.MousePointer = vbHourglass

            DisplayTheForm frmEdu

            lblEmpList.Visible = True
            ASearch = True
            lblECount.Visible = True
'
        Case "E9"
            fraEmployees.Visible = True
            lvwEmp.SetFocus
            Me.MousePointer = vbHourglass

            DisplayTheForm frmDDetails

            lblEmpList.Visible = True
            lblECount.Visible = True
            ASearch = True

'            Professional Qualifications
        Case "E10"
            fraEmployees.Visible = True
            lvwEmp.SetFocus
            Me.MousePointer = vbHourglass

            DisplayTheForm frmProf

            lblEmpList.Visible = True
            lblECount.Visible = True
            ASearch = True
'
''           Employee Awards
    
        Case "E11"
            fraEmployees.Visible = True
            lvwEmp.SetFocus
            Me.MousePointer = vbHourglass

            DisplayTheForm frmAwards

            lblEmpList.Visible = True
            lblECount.Visible = True
            ASearch = True

        Case "S80"
            txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass

            DisplayTheForm frmCompanyDetails

            lblEmpList.Visible = False
            lblECount.Visible = False
        
        Case "DisengagementReasons"
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass

            DisplayTheForm frmDisengagementReasons

            lblEmpList.Visible = False
            lblECount.Visible = False
            
        Case "S1"
            txtDetails.Caption = ""
            fraEmployees.Visible = False
            lblEmpList.Visible = False
            lblECount.Visible = False

            DisplayTheForm frmOUs

            lblEmpList.Visible = False
            lblECount.Visible = False

        Case "S2"
            txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass

            DisplayTheForm frmCTypes
            lblEmpList.Visible = False
            lblECount.Visible = False

        Case "S3"
            txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass

            DisplayTheForm frmBTypes

            lblEmpList.Visible = False
            lblECount.Visible = False

        Case "S4"
            txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass

            DisplayTheForm frmDTypes

            lblEmpList.Visible = False
            lblECount.Visible = False

        Case "S12"
            txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass

            DisplayTheForm frmNationalities

            lblEmpList.Visible = False
            lblECount.Visible = False

        Case "S18"
            txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass

            DisplayTheForm frmReligions

            lblEmpList.Visible = False
            lblECount.Visible = False

        Case "S19"
            txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass
            
            DisplayTheForm frmBankSetUp

            lblEmpList.Visible = False
            lblECount.Visible = False

        Case "S20"
            txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass


            DisplayTheForm frmBankBranchSetUp

            lblEmpList.Visible = False
            lblECount.Visible = False
        
        Case "S40"
            txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass

            DisplayTheForm frmEducationAwardsSetup

            lblEmpList.Visible = False
            lblECount.Visible = False
            
        Case "S41"
            txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass

            DisplayTheForm frmEducationCourses

            lblEmpList.Visible = False
            lblECount.Visible = False
            
            Case "KS42"
            txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass

            DisplayTheForm frmEducationTypes

            lblEmpList.Visible = False
            lblECount.Visible = False
'KS42
'        Case "S21"
'            txtDetails.Caption = ""
'            fraEmployees.Visible = False
'            Me.MousePointer = vbHourglass
'
'            '++Check For User Rights ++
'            'check_User_Rights frmSecurity
'
'            lblEmpList.Visible = False
'            lblECount.Visible = False
        Case "S23"
            txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass

            DisplayTheForm frmCompanyAssets

            lblEmpList.Visible = False
            lblECount.Visible = False
            
'        Case "S24"
'            txtDetails.Caption = ""
'            fraEmployees.Visible = False
'            Me.MousePointer = vbHourglass
'
'            '++Check For User Rights ++
'            'check_User_Rights frmPositionsQualifications
'
'            DisplayTheForm frmPositionsQualifications
'
'            lblEmpList.Visible = False
'            lblECount.Visible = False
'
'         Case "S25"
'            txtDetails.Caption = ""
'            fraEmployees.Visible = False
'            Me.MousePointer = vbHourglass
'
'            '++Check For User Rights ++
'            'check_User_Rights frmPositionRequirements
'
'            DisplayTheForm frmPositionRequirements
'
'            lblEmpList.Visible = False
'            lblECount.Visible = False
        Case "S14"
            txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass

            DisplayTheForm frmBonus

            lblEmpList.Visible = False
            lblECount.Visible = False

        Case "S13"
            txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass

            DisplayTheForm frmTribes

            lblEmpList.Visible = False
            lblECount.Visible = False
'
        Case "S5"
            txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass

            '++Check For User Rights ++
            'check_User_Rights frmCEStructure

            DisplayTheForm frmOUEmployees

            lblEmpList.Visible = False
            lblECount.Visible = False
            lblEmpC.Visible = True
            lblEmpCount.Visible = True
'
        Case "S6"
            'display the form for setting up the OU Types
            fraEmployees.Visible = False
            'Disable the New Button
            Me.cmdNew.Enabled = False
            DisplayTheForm frmOUTypes


            txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass
            lblEmpList.Visible = False
            lblECount.Visible = False

        Case "S15"
            txtDetails.Caption = ""
            fraEmployees.Visible = True
            Me.MousePointer = vbHourglass
           
            DisplayTheForm frmPositions

            lblEmpList.Visible = False
            lblECount.Visible = False

'        Case "S16"
'            txtDetails.Caption = ""
'            fraEmployees.Visible = True
'            Me.MousePointer = vbHourglass
'
'            '++Check For User Rights ++
'            'check_User_Rights frmPApproval
'
'            DisplayTheForm frmPApproval
'
'            lblEmpList.Visible = False
'            lblECount.Visible = False
'
'
'        Case "S8"
'            txtDetails.Caption = ""
'            fraEmployees.Visible = True
'            Me.MousePointer = vbHourglass
'
'            '++Check For User Rights ++
'            'check_User_Rights frmUGroups
'
'            DisplayTheForm frmUGroups
'
'
'            lblEmpList.Visible = False
'            lblECount.Visible = False
'
'        Case "S9"
'            txtDetails.Caption = ""
'            fraEmployees.Visible = False
'            Me.MousePointer = vbHourglass
'
'            '++Check For User Rights ++
'            'check_User_Rights frmSystUsers
'
'            DisplayTheForm frmSystUsers
'
'            lblEmpList.Visible = False
'            lblECount.Visible = False
'
        Case "S10"
            txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass

            DisplayTheForm frmGenOpt

            lblEmpList.Visible = False
            lblECount.Visible = False

        Case "S11"
            txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass

         
            'DisplayTheForm frmEmpCategories
            DisplayTheForm frmGradeTitles
            
            lblEmpList.Visible = False
            lblECount.Visible = False

        Case "S17"
            txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass

            DisplayTheForm frmContractTypes

            lblEmpList.Visible = False
            lblECount.Visible = False
        Case "SDDF"
            
                DisplayTheForm frmdisengagementdateformat
        Case "S22"

            txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass

            '++Check For User Rights ++
            'check_User_Rights frmEmploymentTerms

            DisplayTheForm frmEmploymentTerms

            lblEmpList.Visible = False
            lblECount.Visible = False
        Case "S27"

            txtDetails.Caption = ""
            fraEmployees.Visible = False
            Me.MousePointer = vbHourglass

            DisplayTheForm frmCasualTypes


            lblEmpList.Visible = False
            lblECount.Visible = False
        
        'setup node
        Case "S"
            txtDetails.Caption = ""
            frmCoDetails.Visible = True
        
'        'Utilities node
'        Case "U"
'            txtDetails.Caption = ""
        
        'this is the personnel director node
        Case "L"
            txtDetails.Caption = ""
            frmCoDetails.Visible = True
        
        'employee node
        Case "E"
            txtDetails.Caption = ""
            frmCoDetails.Visible = True
            
    End Select

    Me.MousePointer = 0
'    Call Disabblepromt
    Exit Sub
    
errHandler:
    MsgBox err.Description, vbExclamation
    Call CloseMyWindows
End Sub

Private Sub LoadVar()
    On Error GoTo ErrorHandler
    
    Set rs1 = CConnect.GetRecordSet("SELECT * FROM GeneralOpt WHERE subsystem = '" & SubSystem & "'")
    With rs1
        If .RecordCount > 0 Then
            If !PSave = "Yes" Then PromptSave = True
            If !VSal = "Yes" Then ViewSal = True
            DSource = !DSource & ""
            AppGroup = !AppGroup & ""
            AChange = !SRes & ""
            FRetire = !FRet & ""
            MRetire = !MRet & ""
            Dis = !Dis & ""
            CConnect.DisBase = !DisBase & ""
            Recruit = !Recruit & ""
            CConnect.RecBase = !RecBase & ""

'            If !ADBase = "Yes" Then
'                Dfmt = "dd/mm/yyyy"
'            Else
'                Dfmt = "yyyymmdd"
'            End If

            If IsNull(!DBase) Then
                MsgBox "Emloyee Data source has not been specified therefor the local database will be used.", vbExclamation
                DSource = "Local"
            Else
                CConnect.EDBase = !DBase & ""
            End If

            If !AuditTrail = "Yes" Then
                AuditTrail = True
            Else
                AuditTrail = False
            End If

            EmpGroup = !DGNo & ""
            DPass = !DPass & ""
            IDiv = !IDiv & ""
            CPass = !CPass & ""

        End If
    End With
    Set rs1 = Nothing
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while configuring the General Details" & vbNewLine & err.Description, vbInformation, TITLES

End Sub

Public Sub LoadEmployeeList(Optional ByVal SkipSelection As Boolean)

    Dim i As Long
    Dim ItemX As ListItem

    On Error GoTo ErrorTrap

    lvwEmp.ListItems.Clear
    lblECount.Caption = "0"
    'clear the Details
    txtDetails.Caption = ""

    'try getting accessible employees: Pass the userID
   '' AllEmployees.GetAccessibleEmployeesByUser currUser.UserID
    AllEmployees.GetAccessibleEmployeesByUserID currUser.UserID
    
    '1\2\2007: have modified the listview to display only the employees who are engaged
    
    For i = 1 To AllEmployees.count
        If Not (AllEmployees.Item(i).IsDisengaged) Then 'IF CLAUSE ADDED BY JOHN TO TRAP OBJECT PRE-LOADS
            Set ItemX = Me.lvwEmp.ListItems.add(, , AllEmployees.Item(i).EmpCode)
            ItemX.SubItems(1) = AllEmployees.Item(i).SurName & ", " & AllEmployees.Item(i).OtherNames
            ItemX.ForeColor = AllEmployees.Item(i).Category.CategoryColorCode
            ItemX.ListSubItems(1).ForeColor = AllEmployees.Item(i).Category.CategoryColorCode
            ItemX.Tag = AllEmployees.Item(i).EmployeeID
        End If
    Next i
    lblECount.Caption = lvwEmp.ListItems.count
    If Not IsMissing(SkipSelection) And SkipSelection = True Then GoTo SkipTheSelection

    If Not (TheLoadedForm Is Nothing) Then
        'ensure textboxes are cleared if it is employees
        If TheLoadedForm.Name = "frmEmployee" Then 'Or TheLoadedForm.Name = "frmCheck" Or TheLoadedForm.Name = "frmJProg" Or TheLoadedForm.Name = "frmContacts" Or TheLoadedForm.Name = "frmContract" Or TheLoadedForm.Name = "frmBio" Or TheLoadedForm.Name = "frmKin" Or TheLoadedForm.Name = "frmEdu" Or TheLoadedForm.Name = "frmProf" Or TheLoadedForm.Name = "frmEmployment" Or TheLoadedForm.Name = "frmDDetails" Or TheLoadedForm.Name = "frmRef" Or TheLoadedForm.Name = "frmFamily" Or TheLoadedForm.Name = "frmAwards" Or TheLoadedForm.Name = "frmVisa" Or TheLoadedForm.Name = "frmCasuals" Or TheLoadedForm.Name = "frmEmployeeBanks" Or TheLoadedForm.Name = "frmAssetIssue" Then
            TheLoadedForm.Cleartxt
        End If
    

    'then check whether to populate new data
        If TheLoadedForm.Name = "frmReengageMent" Then GoTo SkipTheSelection:
            If Me.lvwEmp.ListItems.count > 0 Then
             Me.lvwEmp.ListItems(1).Selected = True
             Call lvwEmp_ItemClick(Me.lvwEmp.ListItems(1))
            End If
End If
SkipTheSelection:
    Exit Sub
ErrorTrap:
    MsgBox err.Description, vbExclamation, TITLES
End Sub

''********************* added by kalya on 050909

Public Sub LoadEmployeeListwithemployee(defaultemp As Employee, Optional ByVal SkipSelection As Boolean)

    Dim i As Long
    Dim ItemX As ListItem

    On Error GoTo ErrorTrap
Dim empid As Long
empid = defaultemp.EmployeeID
    lvwEmp.ListItems.Clear
    lblECount.Caption = "0"
    'clear the Details
    txtDetails.Caption = ""

    'try getting accessible employees: Pass the userID
    AllEmployees.GetAccessibleEmployeesByUser currUser.UserID
    
    '1\2\2007: have modified the listview to display only the employees who are engaged
    
    For i = 1 To AllEmployees.count
        If Not (AllEmployees.Item(i).IsDisengaged) Then 'IF CLAUSE ADDED BY JOHN TO TRAP OBJECT PRE-LOADS
            Set ItemX = Me.lvwEmp.ListItems.add(, , AllEmployees.Item(i).EmpCode)
            ItemX.SubItems(1) = AllEmployees.Item(i).SurName & ", " & AllEmployees.Item(i).OtherNames
            ItemX.ForeColor = AllEmployees.Item(i).Category.CategoryColorCode
            ItemX.ListSubItems(1).ForeColor = AllEmployees.Item(i).Category.CategoryColorCode
            ItemX.Tag = AllEmployees.Item(i).EmployeeID
        End If
    Next i
    lblECount.Caption = lvwEmp.ListItems.count
    If Not IsMissing(SkipSelection) And SkipSelection = True Then GoTo SkipTheSelection

    If Not (TheLoadedForm Is Nothing) Then
        'ensure textboxes are cleared if it is employees
        If TheLoadedForm.Name = "frmEmployee" Then 'Or TheLoadedForm.Name = "frmCheck" Or TheLoadedForm.Name = "frmJProg" Or TheLoadedForm.Name = "frmContacts" Or TheLoadedForm.Name = "frmContract" Or TheLoadedForm.Name = "frmBio" Or TheLoadedForm.Name = "frmKin" Or TheLoadedForm.Name = "frmEdu" Or TheLoadedForm.Name = "frmProf" Or TheLoadedForm.Name = "frmEmployment" Or TheLoadedForm.Name = "frmDDetails" Or TheLoadedForm.Name = "frmRef" Or TheLoadedForm.Name = "frmFamily" Or TheLoadedForm.Name = "frmAwards" Or TheLoadedForm.Name = "frmVisa" Or TheLoadedForm.Name = "frmCasuals" Or TheLoadedForm.Name = "frmEmployeeBanks" Or TheLoadedForm.Name = "frmAssetIssue" Then
            TheLoadedForm.Cleartxt
        End If
    

    'then check whether to populate new data
        If TheLoadedForm.Name = "frmReengageMent" Then GoTo SkipTheSelection:
            If Me.lvwEmp.ListItems.count > 0 Then
             Me.lvwEmp.ListItems(1).Selected = True
             Call lvwEmp_ItemClick_thisEmployee(defaultemp)
            End If
End If
SkipTheSelection:
    Exit Sub
ErrorTrap:
    MsgBox err.Description, vbExclamation, TITLES
End Sub


''end added by kalya

Public Sub LoadEmployeeListFiltered(ByVal TheFilteredList As HRCORE.Employees)

    Dim i As Long
    Dim ItemX As ListItem

    On Error GoTo ErrorTrap

    lvwEmp.ListItems.Clear

    txtDetails.Caption = ""

    If TheFilteredList Is Nothing Then Exit Sub

    For i = 1 To TheFilteredList.count
        If TheFilteredList.Item(i).IsDisengaged = False Then
            Set ItemX = Me.lvwEmp.ListItems.add(, , TheFilteredList.Item(i).EmpCode)
            ItemX.SubItems(1) = TheFilteredList.Item(i).SurName & ", " & TheFilteredList.Item(i).OtherNames
            ItemX.ForeColor = TheFilteredList.Item(i).Category.CategoryColorCode
            ItemX.ListSubItems(1).ForeColor = TheFilteredList.Item(i).Category.CategoryColorCode
            ItemX.Tag = TheFilteredList.Item(i).EmployeeID
        End If
    Next i

    'ensure textboxes are cleared if it is employees
    If Not TheLoadedForm Is Nothing Then
        If TheLoadedForm.Name = "frmEmployee" Or TheLoadedForm.Name = "frmCheck" Or TheLoadedForm.Name = "frmJProg" Or TheLoadedForm.Name = "frmContacts" Or TheLoadedForm.Name = "frmContract" Or TheLoadedForm.Name = "frmBio" Or TheLoadedForm.Name = "frmKin" Or TheLoadedForm.Name = "frmEdu" Or TheLoadedForm.Name = "frmProf" Or TheLoadedForm.Name = "frmEmployment" Or TheLoadedForm.Name = "frmDDetails" Or TheLoadedForm.Name = "frmRef" Or TheLoadedForm.Name = "frmFamily" Or TheLoadedForm.Name = "frmAwards" Or TheLoadedForm.Name = "frmVisa" Or TheLoadedForm.Name = "frmCasuals" Or TheLoadedForm.Name = "frmEmployeeBanks" Or TheLoadedForm.Name = "frmAssetIssue" Then
            ClearText
        End If
    End If
    'then check whether to populate new data
    If Me.lvwEmp.ListItems.count > 0 Then
        Me.lvwEmp.ListItems(1).Selected = True
        Call lvwEmp_ItemClick(Me.lvwEmp.ListItems(1))
    End If
    Exit Sub
ErrorTrap:
    MsgBox err.Description, vbExclamation, TITLES
    
End Sub

'
Public Sub DisableCmd()
    Dim i As Object
    For Each i In Me
        If TypeOf i Is CommandButton Then
            i.Enabled = False
        End If
    Next i


End Sub

Public Sub EnableCmd()
    Dim i As Object
    For Each i In Me
        If TypeOf i Is CommandButton Then
            If i.Name <> cmdShowPrompts Then
                i.Enabled = True
            End If
        End If
    Next i

End Sub

Public Sub InitGrid()
    With lvwProb
        .ColumnHeaders.Clear
        .ColumnHeaders.add , , "Employee Code.", 1500
        .ColumnHeaders.add , , "Name", 2450
        .ColumnHeaders.add , , "Date Employed", 1500
        .ColumnHeaders.add , , "Probation Type", 1500
        .ColumnHeaders.add , , "Probation Start Date", 1800
        .ColumnHeaders.add , , "Confirmation Date", 1500
        .ColumnHeaders.add , , "Days remaining(Days)", 2000
        .View = lvwReport
    End With
    
    With lvwContracts
        .ColumnHeaders.Clear
        .ColumnHeaders.add , , "Employee Code", 1500
        .ColumnHeaders.add , , "Employee's Name", 2450
        .ColumnHeaders.add , , "Date Of Birth", 1500
        .ColumnHeaders.add , , "Date of Employment", 2000
        .ColumnHeaders.add , , "Retirement Date", 1500
        .ColumnHeaders.add , , "Days To Retirement", 1800
        .View = lvwReport
    
    End With
    With lvwVisa
        .ColumnHeaders.Clear
        .ColumnHeaders.add , , "Employee Code", 1300
        .ColumnHeaders.add , , "Name", 2450
        .ColumnHeaders.add , , "Date of Employment", 1800
        .ColumnHeaders.add , , "Contract Expiry Date", 1800
        .ColumnHeaders.add , , "Time Remaining(Days)", 2000
        .View = lvwReport
    
    End With

End Sub

Public Sub LoadCbo()
    LoadEmployeeCategories
    LoadEmploymentTerms
    LoadOrganizationUnits
End Sub

Private Sub DisplayTheForm(onForm As Form)
    'This procedure done by Oscar overrides the Check_User_Rights procedure

    onForm.Show , Me
    fracmd.Visible = True
    'frmMain2.Caption = Me.Caption & " - " & onForm.Caption
    Set TheLoadedForm = Nothing
    Set TheLoadedForm = onForm
    Me.MousePointer = 0
End Sub

Public Sub PositionTheFormWithEmpList(TheForm As Form)
    On Error GoTo ErrorHandler

    oSmart.FReset TheForm

    If oSmart.hRatio > 1.1 Then
        With frmMain2
            TheForm.Move .tvwMain.Width + .lvwEmp.Width + (.tvwMain.Width / 36) * 2, .Frame2.Top + (.Frame2.Height * 2.5) '(.Height / 5.52) '- 155
        End With
    Else
         With frmMain2
            TheForm.Move .tvwMain.Width + .lvwEmp.Width + (.tvwMain.Width / 36#) * 2, .Frame2.Top + (.Frame2.Height * 2.5) '(.Height / 5.52)
        End With

    End If

    'MsgBox App.FileDescription

    CConnect.CColor TheForm, MyColor
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred while Positioning the window" & vbNewLine & err.Description, vbInformation, TITLES

End Sub

Public Sub PositionTheFormWithoutEmpList(TheForm As Form)
    On Error GoTo ErrorHandler

    oSmart.FReset TheForm

    If oSmart.hRatio > 1.1 Then
        With frmMain2
            TheForm.Move .tvwMain.Width + .tvwMain.Width / 32.75, .Frame2.Top + (.Frame2.Height * 2.5) '(.Height / 5.52)
            .lvwEmp.Visible = False
        End With
    Else
         With frmMain2
            TheForm.Move .tvwMain.Width + .tvwMain.Width / 32.75, .Frame2.Top + (.Frame2.Height * 2.5) '(.Height / 5.52)
        End With

    End If

    CConnect.CColor TheForm, MyColor

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred while Positioning the window", vbInformation, TITLES
End Sub

Public Sub LoadEmployeeCategories()
    Dim i As Long

    ChangedByCode = True
    cboCat.Clear

   '' EmpCats.GetAllEmployeeCategories

    Me.cboCat.AddItem "(All Categories)"
    For i = 1 To EmpCats.count
        Me.cboCat.AddItem EmpCats.Item(i).CategoryName
        Me.cboCat.ItemData(Me.cboCat.NewIndex) = EmpCats.Item(i).CategoryID
    Next i
    cboCat.ListIndex = 0

    ChangedByCode = False
End Sub

Public Sub LoadEmploymentTerms()
    Dim i As Long

    ChangedByCode = True
    cboTerms.Clear

    empTerms.GetAllEmploymentTerms
    cboTerms.AddItem "(All Terms)"
    For i = 1 To empTerms.count
        cboTerms.AddItem empTerms.Item(i).EmpTermName
        cboTerms.ItemData(cboTerms.NewIndex) = empTerms.Item(i).EmpTermID
    Next i

    cboTerms.ListIndex = 0

    ChangedByCode = False
End Sub

Public Sub LoadOrganizationUnits()
    Dim i As Long

    ChangedByCode = True
    Me.cboStructure.Clear

    OUs.GetAllOrganizationUnits

    cboStructure.AddItem "(All Organization Units)"
    For i = 1 To OUs.count
        cboStructure.AddItem OUs.Item(i).OrganizationUnitName
        cboStructure.ItemData(cboStructure.NewIndex) = OUs.Item(i).OrganizationUnitID
    Next i

    cboStructure.ListIndex = 0
    ChangedByCode = False
End Sub

Private Sub ClearText()

    Dim i As Long
    Dim cont As Control
    For Each cont In TheLoadedForm
        If TypeOf cont Is TextBox Then
            cont.Text = ""
        End If
    Next cont
End Sub

Private Function ProbPrompt(ByVal OnClick As Integer) As Boolean
    On Error GoTo ErrorHandler
    Dim ProbDiff As Long
    Dim PRS1 As ADODB.Recordset, ItemX As ListItem
    Dim X As Integer
    Dim PType As String
    
    ProbPrompt = False 'Initiliaze to FALSE on start
    
    If fraContracts.Visible = True Then
        fraContracts.Visible = False
    End If
    
    Set PRS1 = New ADODB.Recordset
    Set PRS1 = con.Execute("pdspSelectAllDueProbationEnds")
        
    If Not (PRS1.BOF Or PRS1.EOF) Then
        PRS1.MoveFirst 'Force recordset to start @ the begining
        
        lvwProb.ListItems.Clear
        Do Until PRS1.EOF
            'StartDate = Format(pRS.Fields("ProbationDate").value, "dd/mm/yyyy")
           '' ProbDiff = DateDiff("d", Format(Now, "dd/mm/yyyy"), PRS1!ConfirmationDate)
             ProbDiff = PRS1!daysremaining
'            If ((ProbDiff <= ProbPromptDays.Item(1).Probation) And (ProbDiff > 0)) Then
                Set ItemX = lvwProb.ListItems.add(, , PRS1.Fields("empcode").value)
                ItemX.SubItems(1) = PRS1!EmployeeName
                'itemX.SubItems(2) = pRS.Fields("IDNo").value
                ItemX.SubItems(2) = Format(PRS1.Fields("DateOfEmployment").value, "dd-MMM-yyyy")
                ItemX.SubItems(3) = PRS1!ProbationType
                ItemX.SubItems(4) = Format(PRS1!ProbationStartDate, "dd-MMM-yyyy")
                ItemX.SubItems(5) = Format(PRS1!ConfirmationDate, "dd-MMM-yyyy")
                ItemX.SubItems(6) = ProbDiff
            
            PRS1.MoveNext
        Loop
    End If
    
    'Dislay Frame for probation
    If (lvwProb.ListItems.count > 0) Then
        fraProb.Visible = True
        fraProb.Left = 2500
        fraProb.Top = 1800
        fraEmployees.Visible = False
        FraTerminate.Visible = False
        ProbPrompt = True
    Else
     'No Employees are on probation
        If (OnClick = 0) Then
            MsgBox "There are no employees due for confirmation from probation", _
            vbOKOnly + vbInformation, "Probation prompts"
        End If
    End If
    
    Exit Function

ErrorHandler:
    MsgBox "An error has occured: " & err.Description, vbExclamation, "Personell Director: Error"
End Function

Private Function RetirementPrompt(ByVal OnClick As Integer) As Boolean
    On Error GoTo ErrorHandler
    
    Dim prs As New ADODB.Recordset, ItemX As ListItem
    Dim X As Integer, AgeDiff As Long
    Dim TimeRemaining As Long
     
    RetirementPrompt = False 'Set as FALSE on start
    Set prs = CConnect.GetRecordSet("pdrspSelectAllDueRetirements")
    
    If Not (prs.BOF Or prs.EOF) Then
        prs.MoveFirst
        lvwContracts.ListItems.Clear
        
        Do Until prs.EOF
            Set ItemX = lvwContracts.ListItems.add(, , prs.Fields("empcode").value)
            ItemX.SubItems(1) = prs!EmployeeName
            ItemX.SubItems(2) = Format(prs!DateOfBirth, "DD-MMM-YYYY")
            ItemX.SubItems(3) = Format(prs!DateOfEmployment, "DD-MMM-YYYY")
            ItemX.SubItems(4) = Format(prs!RetirementDate, "DD-MMM-YYYY")
            ItemX.SubItems(5) = DateDiff("d", Now(), prs!RetirementDate)
            prs.MoveNext
        Loop
        
        'Display Due Retirees
        fraContracts.Visible = True
        fraVisa.Visible = False
        fraProb.Visible = False
        fraContracts.Left = 2500
        fraContracts.Top = 1800
        fraEmployees.Visible = False
        FraTerminate.Visible = False
        RetirementPrompt = True
    Else
        'No Employees are about to retire
        If (OnClick = 0) Then
            MsgBox "There are no employees about to retire", _
            vbOKOnly + vbInformation, "Retirement prompts"
        End If
    End If
        
    Exit Function
ErrorHandler:
    MsgBox "An error has occured " & err.Description, vbInformation, "Personel Director: Error"
End Function

Private Function EmployeeTerminatePrompt(ByVal OnClick As Integer) As Boolean
    On Error GoTo ErrorHandler
    Dim prs As New ADODB.Recordset
    Dim TimeRemaining As Long, my As Long
    Dim ItemX As ListItem
    
    EmployeeTerminatePrompt = False 'Set as FALSE on Start
        
    With LvwTerminate
        .ColumnHeaders.Clear
        .ColumnHeaders.add , , "Emp. CODE", 1400
        .ColumnHeaders.add , , "Employee Name", 2500
        .ColumnHeaders.add , , "Employment Date", 1800
        .ColumnHeaders.add , , "Date of Birth", 1800
        .ColumnHeaders.add , , "Employment Validity", 1800
        .ColumnHeaders.add , , "Days To Go", 1400
    End With
    
    Set prs = CConnect.GetRecordSet("pdspGetEmployeeTerminationsDue")
    
    If Not (prs.BOF Or prs.EOF) Then
        prs.MoveFirst 'Force recordset to start @ the begining
        
        'Display Employees Due for Termination
        LvwTerminate.ListItems.Clear
        
        Do Until prs.EOF
            TimeRemaining = DateDiff("d", prs!EmploymentValidThro, Format(Now, "dd/mm/yyyy"))
            'If ((TimeRemaining <= ProbPromptDays.Item(1).Contract) And (TimeRemaining > 0)) Then
                Set ItemX = LvwTerminate.ListItems.add(, , prs.Fields("empcode").value)
                ItemX.SubItems(1) = prs!EmployeeName
                ItemX.SubItems(2) = Format(prs!DateOfEmployment, "DD-MMM-YYYY")
                ItemX.SubItems(3) = Format(prs!DateOfBirth, "DD-MMM-YYYY")
                ItemX.SubItems(4) = Format(prs!EmploymentValidThro, "DD-MMM-YYYY")
                ItemX.SubItems(5) = prs!daysremaining
            'Else
            'End If
            prs.MoveNext
        Loop
        
        If LvwTerminate.ListItems.count > 0 Then
            fraProb.Visible = False
            fraContracts.Visible = False
            fraEmployees.Visible = False
            FraBirthDays.Visible = False
            FraTerminate.Visible = True
            FraTerminate.Left = 3000
            FraTerminate.Top = 1800
            fraVisa.Visible = False
            fraEmployees.Visible = False
            EmployeeTerminatePrompt = True 'If there are prompts, then TRUE
        End If
    Else
        'No Employees due for termination
        If (OnClick = 0) Then
            MsgBox "There are no employees due for termination"
        End If
    End If
    
    Exit Function
ErrorHandler:
      MsgBox "An error has occured " & err.Description, vbInformation, "Personel Director: Error"
End Function

Private Function ContractExpirePrompt(ByVal OnClick As Integer) As Boolean
    On Error GoTo ErrorHandler
    Dim prs As New ADODB.Recordset
    Dim TimeRemaining As Long, my As Long
    Dim ItemX As ListItem
    
    ContractExpirePrompt = False 'Set as FALSE on Start
    
    If fraProb.Visible = True Then
        fraProb.Visible = False
    ElseIf fraContracts.Visible = True Then
        fraContracts.Visible = False
    End If
        
'     mySQL = "Select EmpCode, (Surname + ' ' + OtherNames) as EmployeeName, " & _
'    "DateOfEmployment,EmploymentValidThro as ExpiryDate " & _
'     " from employees where EmpTermID=3"
    
    mySQL = "pdspSelectAllDueContractExpiries"
    Set prs = CConnect.GetRecordSet(mySQL)
    
    If Not (prs.BOF Or prs.EOF) Then
        prs.MoveFirst 'Force recordset to start @ the begining
        'Dislay Frame for Retirees
        lvwVisa.ListItems.Clear
        Do Until prs.EOF
            ''TimeRemaining = DateDiff("d", prs!ExpiryDate, Format(Now, "dd/mm/yyyy"))
            TimeRemaining = prs!daysremaining
            If ((TimeRemaining <= ProbPromptDays.Item(1).Contract) And (TimeRemaining > 0)) Then
                Set ItemX = lvwVisa.ListItems.add(, , prs.Fields("empcode").value)
                ItemX.SubItems(1) = prs!EmployeeName
                ItemX.SubItems(2) = Format(prs!DateOfEmployment, "DD-MMM-YYYY")
                ''ItemX.SubItems(3) = Format(prs!ExpiryDate, "DD-MMM-YYYY")
                ItemX.SubItems(3) = Format(CDate(prs!EmploymentValidThro), "DD-MMM-YYYY")
                
                ItemX.SubItems(4) = TimeRemaining
            Else
            End If
            prs.MoveNext
        Loop
        If lvwVisa.ListItems.count > 0 Then
            fraVisa.Visible = True
            fraVisa.Left = 2500
            fraVisa.Top = 1800
            fraEmployees.Visible = False
            FraTerminate.Visible = False
            ContractExpirePrompt = True 'If there are prompts, then TRUE
        End If
    Else
        'No Employees whose contracts are about to expire
        If (OnClick = 0) Then
            MsgBox "There are no employees whose contracts are about to expire"
        End If
    End If
    
    Exit Function
ErrorHandler:
      MsgBox "An error has occured " & err.Description, vbInformation, "Personel Director: Error"
End Function

Private Sub GetPromptDays()
    On Error GoTo errHandler
        
    Set ProbPromptDays = New Prompts
    ProbPromptDays.GetAllPrompts
        
    Exit Sub
errHandler:
    MsgBox "An error has occured " & err.Description & err.Number, vbInformation, "Personel Director: Getting days Error"
End Sub

Function GetBirthDays(ByVal OnClick As Integer) As Boolean
    On Error GoTo errHandler
    
    Dim mySQL As String
    Dim RsT As New ADODB.Recordset
    Dim ItemX As ListItem
    
    fraContracts.Visible = False
    fraProb.Visible = False
    fraVisa.Visible = False
    fraEmployees.Visible = False
    FraTerminate.Visible = False
    
    mySQL = "pdspGetBirthdaysForThisMonth"
    
    Set RsT = con.Execute(mySQL)
    If Not (RsT.EOF Or RsT.BOF) Then
        With LvwBirthdays
            .ListItems.Clear
            RsT.MoveFirst
            Do Until RsT.EOF
                Set ItemX = .ListItems.add(, , RsT!EmpCode)
                ItemX.SubItems(1) = RsT!SurName & " " & RsT!OtherNames
                ItemX.SubItems(2) = Format(RsT!DateOfBirth, "dd-MMM-yyyy")
                ItemX.SubItems(3) = Format(Format(RsT!DateOfBirth, "dd-MMM"), "dddd, MMMM dd")
                RsT.MoveNext
            Loop
        End With
        
        'Show Frame With Data
        FraBirthDays.Left = 3000
        FraBirthDays.Top = 1800
        FraBirthDays.Visible = True
    Else
        If (OnClick = 0) Then
           MsgBox "No birthdays for this month", vbInformation, "PDR"
        End If
    End If
    
    Set RsT = Nothing
    
    Exit Function
    
errHandler:
    MsgBox "An error has occured " & err.Description & err.Number, vbInformation, "Personel Director: Getting days Error"
End Function

Private Sub mnuLog_Click()
    Call LogIn
    Call EnableCmd
End Sub
