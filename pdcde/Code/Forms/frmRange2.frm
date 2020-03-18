VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmRange2 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Range"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRange2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   8610
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Paypoint"
      Height          =   2415
      Left            =   4440
      TabIndex        =   56
      Top             =   2640
      Width           =   3735
      Begin MSComctlLib.ListView lvwpaypoint 
         Height          =   2055
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   3625
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Paypoint"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cost Center"
      Height          =   2415
      Left            =   4440
      TabIndex        =   54
      Top             =   120
      Width           =   3735
      Begin MSComctlLib.ListView lvwcostcenter 
         Height          =   2055
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   3625
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cost Center"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSComctlLib.ListView lvwEmp 
      Height          =   255
      Left            =   4680
      TabIndex        =   52
      Top             =   840
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame fraOUs 
      Appearance      =   0  'Flat
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   120
      TabIndex        =   50
      Top             =   0
      Width           =   4185
      Begin MSComctlLib.ListView LvwDpts 
         Height          =   4455
         Left            =   240
         TabIndex        =   51
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   7858
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Departments"
            Object.Width           =   5821
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2FFFF&
      Caption         =   "Filters"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4860
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   3450
      Begin VB.Frame fraEmployee 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FFFF&
         Caption         =   "Employee"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1200
         Left            =   105
         TabIndex        =   2
         Top             =   225
         Width           =   3825
         Begin VB.TextBox txtFromE 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1665
            TabIndex        =   6
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtToE 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            TabIndex        =   5
            Top             =   720
            Width           =   1695
         End
         Begin VB.CommandButton cmdFromE 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3360
            Picture         =   "frmRange2.frx":1762
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            Width           =   345
         End
         Begin VB.CommandButton cmdToE 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3360
            Picture         =   "frmRange2.frx":1864
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   720
            Width           =   345
         End
         Begin VB.Label lblERange 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Range"
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1500
         End
      End
      Begin VB.Frame fraCategory 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FFFF&
         Caption         =   "Employee Category"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1080
         Left            =   105
         TabIndex        =   38
         Top             =   2520
         Visible         =   0   'False
         Width           =   3825
         Begin VB.CommandButton cmdCatT 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3360
            Picture         =   "frmRange2.frx":1966
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   645
            Width           =   345
         End
         Begin VB.CommandButton cmdCatF 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3360
            Picture         =   "frmRange2.frx":1A68
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   270
            Width           =   345
         End
         Begin VB.TextBox txtCatT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1665
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   645
            Width           =   1695
         End
         Begin VB.TextBox txtCatF 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1665
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   270
            Width           =   1695
         End
         Begin VB.Label Label9 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Category Range"
            Height          =   315
            Left            =   135
            TabIndex        =   43
            Top             =   285
            Width           =   1500
         End
      End
      Begin VB.Frame fraBonus 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FFFF&
         Caption         =   "Bonus"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   105
         TabIndex        =   26
         Top             =   5700
         Width           =   3825
         Begin VB.ComboBox cboBonus 
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   315
            ItemData        =   "frmRange2.frx":1B6A
            Left            =   1680
            List            =   "frmRange2.frx":1B6C
            TabIndex        =   27
            Top             =   210
            Width           =   2040
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Range"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   840
            TabIndex        =   28
            Top             =   225
            Width           =   465
         End
      End
      Begin VB.Frame fraLongService 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FFFF&
         Caption         =   "Long Service"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1380
         Left            =   105
         TabIndex        =   21
         Top             =   4270
         Width           =   3825
         Begin MSComCtl2.DTPicker dtpLongService 
            Height          =   375
            Left            =   960
            TabIndex        =   49
            Top             =   240
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "MMMM yyyy"
            Format          =   63963139
            CurrentDate     =   38740
         End
         Begin VB.ComboBox cboLST 
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   315
            ItemData        =   "frmRange2.frx":1B6E
            Left            =   2520
            List            =   "frmRange2.frx":1B8A
            TabIndex        =   47
            Top             =   960
            Width           =   630
         End
         Begin VB.ComboBox cboLS 
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   315
            ItemData        =   "frmRange2.frx":1BAA
            Left            =   1440
            List            =   "frmRange2.frx":1BC6
            TabIndex        =   24
            Top             =   945
            Width           =   750
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2280
            TabIndex        =   48
            Top             =   960
            Width           =   180
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Long Service Years"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   105
            TabIndex        =   25
            Top             =   720
            Width           =   1365
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "As at"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   960
            TabIndex        =   22
            Top             =   960
            Width           =   360
         End
      End
      Begin VB.Frame fraTermination 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FFFF&
         Caption         =   "Termination"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   105
         TabIndex        =   14
         Top             =   3600
         Width           =   3825
         Begin VB.ComboBox cboTermReasons 
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   315
            ItemData        =   "frmRange2.frx":1BE6
            Left            =   1680
            List            =   "frmRange2.frx":1BF9
            TabIndex        =   15
            Top             =   210
            Width           =   2040
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reasons"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   825
            TabIndex        =   16
            Top             =   225
            Width           =   615
         End
      End
      Begin VB.Frame fraDates 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FFFF&
         Caption         =   "Dates"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   945
         Left            =   105
         TabIndex        =   11
         Top             =   2640
         Width           =   3825
         Begin VB.CheckBox chkDates 
            BackColor       =   &H00F2FFFF&
            Caption         =   "Do not Use Date"
            Height          =   375
            Left            =   120
            TabIndex        =   37
            Top             =   210
            Width           =   1110
         End
         Begin MSComCtl2.DTPicker dtpTo 
            Height          =   300
            Left            =   2160
            TabIndex        =   12
            Top             =   540
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   63963139
            CurrentDate     =   38475
         End
         Begin MSComCtl2.DTPicker dtpFrom 
            Height          =   315
            Left            =   2160
            TabIndex        =   13
            Top             =   180
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   63963139
            CurrentDate     =   38475
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   315
            Left            =   1440
            TabIndex        =   18
            Top             =   525
            Width           =   555
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "From"
            Height          =   315
            Left            =   1365
            TabIndex        =   17
            Top             =   180
            Width           =   555
         End
      End
      Begin VB.Frame fraDiv 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FFFF&
         Caption         =   "Division"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   105
         TabIndex        =   8
         Top             =   1800
         Visible         =   0   'False
         Width           =   3825
         Begin VB.TextBox txtDiv 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1650
            TabIndex        =   10
            Top             =   195
            Width           =   1710
         End
         Begin VB.CommandButton cmdDiv 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3360
            Picture         =   "frmRange2.frx":1C39
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   195
            Width           =   360
         End
      End
   End
   Begin VB.Frame fraCheck 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2FFFF&
      Caption         =   "Missing Mandatory Fields"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1650
      Left            =   360
      TabIndex        =   19
      Top             =   3120
      Visible         =   0   'False
      Width           =   3825
      Begin VB.CheckBox chkDesig 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Designation"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   33
         Top             =   1950
         Width           =   2055
      End
      Begin VB.CheckBox chkBasic 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Basic Pay"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   32
         Top             =   1605
         Width           =   2055
      End
      Begin VB.CheckBox chkCert 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Certificate No"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   31
         Top             =   1260
         Width           =   2055
      End
      Begin VB.CheckBox chkPIN 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "P.I.N No"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   30
         Top             =   930
         Width           =   2055
      End
      Begin VB.CheckBox chkNSSF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "N.S.S.F No"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   29
         Top             =   585
         Width           =   2055
      End
      Begin VB.CheckBox chkNHIF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "N.H.I.F No"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   20
         Top             =   270
         Width           =   2055
      End
   End
   Begin VB.Frame fraMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2FFFF&
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   360
      TabIndex        =   34
      Top             =   2415
      Visible         =   0   'False
      Width           =   3825
      Begin VB.ComboBox cboBMonth 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmRange2.frx":1D3B
         Left            =   1995
         List            =   "frmRange2.frx":1D66
         TabIndex        =   35
         Top             =   180
         Width           =   1725
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   810
         TabIndex        =   36
         Top             =   180
         Width           =   450
      End
   End
   Begin VB.Frame fraSelCat 
      BackColor       =   &H00F2FFFF&
      Caption         =   "Employee Categories"
      Height          =   4335
      Left            =   240
      TabIndex        =   44
      Top             =   555
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CommandButton cmdCancelC 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   2640
         TabIndex        =   46
         Top             =   4080
         Width           =   1215
      End
      Begin MSComctlLib.ListView lsvCat 
         Height          =   4095
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   7223
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   5160
      Width           =   1500
   End
   Begin MSComDlg.CommonDialog Cdl 
      Left            =   660
      Top             =   4395
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblEmpCount 
      Height          =   375
      Left            =   2040
      TabIndex        =   53
      Top             =   4560
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmRange2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim From As Boolean
Dim myTo As Boolean
Dim myLevelCode As String
Dim MyLevel As Integer
Dim myEmpAccess As String
Dim ReportPeriod As Integer
Dim ReportYear As Long
Dim rsR1 As Recordset
Dim rsR2 As Recordset
Dim rsR3 As Recordset
Dim rsR4 As Recordset
Dim ParentN As Boolean
Dim catFrom As Boolean

Dim rServer As String, rCatalog As String, rConnection As String, rUId As String, rPass As String

Public Sub disabletxt()
Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
            i.Locked = True
        End If
    Next i
End Sub

Private Sub cboDiv_Click()
Set rsR1 = CConnect.GetRecordSet("SELECT TLevel FROM DivisionTypes WHERE TypeName='" & cboDiv.Text & "'")
    
With rsR1
    If .RecordCount > 0 Then
        .MoveFirst
        MyLevel = !TLevel
    End If
End With

Set rsR1 = Nothing

End Sub

Private Sub cboDiv_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cboPeriod_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub chkDetailed_Click()
If ReportType = "TransAnalysis" And chkDetailed.value = 1 Then
    cboPeriod.Enabled = False
Else
    cboPeriod.Enabled = True
End If
End Sub

Private Sub chkNet_Click()
If chkNet.value = 0 Then
    chkSigned.value = 1
End If
End Sub

Private Sub chkSigned_Click()
If chkSigned.value = 0 Then
    chkNet.value = 1
End If
End Sub

Private Sub cboBonus_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboLS_KeyPress(KeyAscii As Integer)
    'KeyAscii = 0
End Sub

Private Sub cboLS_LostFocus()
    cboLST.Text = cboLS.Text
End Sub

Private Sub cboTermReasons_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboTerms_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
    fraBrowse.Visible = False
End Sub

Private Sub cmdCancelC_Click()
    fraSelCat.Visible = False
    Frame1.Enabled = True
End Sub

Private Sub cmdCatF_Click()
    catFrom = True
    fraSelCat.Visible = True
    Frame1.Enabled = False
End Sub

Private Sub cmdCatT_Click()
    catFrom = False
    fraSelCat.Visible = True
    Frame1.Enabled = False
End Sub

Private Sub cmdDiv_Click()
    Call Hidelvw
    trwStruc.Visible = True
End Sub

Private Sub cmdFromE_Click()
    Sel = txtFromE.Text & ""
    popupText = "RFrom2"
    frmPopUp.Show vbModal
  
End Sub

Private Sub cmdOFrom_Click()
    From = True
    Call Hidelvw
If ReportType = "TransList" Or ReportType = "TransDiv" Or ReportType = "TransAnalysis" Then
    lvwTrans.Visible = True
ElseIf ReportType = "BankList" Or ReportType = "BBranchList" Then
    lvwBanks.Visible = True
ElseIf ReportType = "CurList" Or ReportType = "CurDenomList" Then
    lvwCur.Visible = True
ElseIf ReportType = "PayRateList" Then
    lvwPayrate.Visible = True
ElseIf ReportType = "CompList" Or ReportType = "CompPayslip" Or ReportType = "P10Sum" Then
    lvwComp.Visible = True
ElseIf ReportType = "HseList" Then
    lvwHse.Visible = True
ElseIf ReportType = "InsList" Then
    lvwIns.Visible = True
End If
End Sub

Private Sub cmdOTo_Click()
 myTo = True
    Hidelvw
If ReportType = "TransList" Or ReportType = "TransDiv" Or ReportType = "TransAnalysis" Then
    lvwTrans.Visible = True
ElseIf ReportType = "BankList" Or ReportType = "BBranchList" Then
    lvwBanks.Visible = True
ElseIf ReportType = "CurList" Or ReportType = "CurDenomList" Then
    lvwCur.Visible = True
ElseIf ReportType = "PayRateList" Then
    lvwPayrate.Visible = True
ElseIf ReportType = "CompList" Or ReportType = "CompPayslip" Or ReportType = "P10Sum" Then
    lvwComp.Visible = True
ElseIf ReportType = "HseList" Then
    lvwHse.Visible = True
ElseIf ReportType = "InsList" Then
    lvwIns.Visible = True
End If

End Sub

Private Sub cmdprint_Click()

On Error GoTo ErrHandler
    Dim myfile As String
    Dim ss As String
    Dim MyL As Long
    Dim MyNo1 As Long
    Dim MyNo2 As Long
    Dim conProps As CRAXDDRT.ConnectionProperties
    Set r = crtDptHierarchy
    MyL = Len(myLevelCode)
    
   
    Me.MousePointer = vbHourglass
    
'    If R.HasSavedData = True Then
'        R.DiscardSavedData
'    End If
    ''pdspEmployeesDptsByHierarchy.
    mySQL = ""
    
    
    
    If RFilter <> "Audit Trail" Then
        If Not txtFromE.Text = "" Then
            mySQL = "{E.EMPCODE} in ['" & txtFromE.Text & "' to '" & txtToE.Text & "']"
        End If
        
        ''
        ''--------------
        
        Dim tSQL As String
           tSQL = ""
        For k = 2 To LvwDpts.ListItems.count
            If (LvwDpts.ListItems.Item(k).Checked = True) Then
                tSQL = tSQL & "'" & Trim(Replace(LvwDpts.ListItems.Item(k).Text, "'", "''")) & "',"
            End If
        Next k
        If (Trim(tSQL) = "") Then
            MsgBox "You must have at least one department selected", vbExclamation, "Report Error"
            Exit Sub
        Else
            tSQL = Mid$(tSQL, 1, Len(tSQL) - 1)
        End If
        ''-----------
        
       If (tSQL <> "") Then
       If (mySQL <> "") Then
       mySQL = mySQL & " AND {E.ORGANIZATIONUNITNAME}  in [" & tSQL & "]"
       Else
       mySQL = "{E.ORGANIZATIONUNITNAME}  in [" & tSQL & "]"
       End If
       End If
       ''----------
       tSQL = ""
       For k = 2 To lvwcostcenter.ListItems.count
            If (lvwcostcenter.ListItems.Item(k).Checked = True) Then
                tSQL = tSQL & "" & lvwcostcenter.ListItems(k).Tag & ","
            End If
        Next k
        If (Trim(tSQL) = "") Then
            MsgBox "You must have at least one Costcenter selected", vbExclamation, "Report Error"
            Exit Sub
        Else
            tSQL = Mid$(tSQL, 1, Len(tSQL) - 1)
        End If
        ''---------
        
        If (Trim(tSQL) <> "") Then
        If (mySQL <> "") Then
        mySQL = mySQL & " AND {tblcostcenter.costcenter_id}  in [" & tSQL & "]"
        Else
        mySQL = "{tblcostcenter.costcenter_id}  in [" & tSQL & "]"
        End If
        End If
       
          tSQL = ""
          For k = 2 To lvwpaypoint.ListItems.count
            If (lvwpaypoint.ListItems.Item(k).Checked = True) Then
                tSQL = tSQL & "" & lvwpaypoint.ListItems(k).Tag & ","
            End If
          Next k
          If (Trim(tSQL) = "") Then
            MsgBox "You must have at least one Paypoint selected", vbExclamation, "Report Error"
            Exit Sub
          Else
            tSQL = Mid$(tSQL, 1, Len(tSQL) - 1)
          End If
          
          If (tSQL <> "") Then
          If (mySQL <> "") Then
          mySQL = mySQL & " AND {prlEmployeePayrollInfo.payrollid}  in [" & tSQL & "]"
          Else
          mySQL = "{prlEmployeePayrollInfo.payrollid}  in [" & tSQL & "]"
          End If
          End If
          
         
        
       ShowReport r
       Unload Me
       Exit Sub
  End If
       
ErrHandler:
 
        MsgBox err.Description, vbInformation
         Me.MousePointer = 0
    'End If

End Sub

Private Sub cmdSelect_Click()
On Error GoTo Hell
If lvwEmp.Visible = True Then
    If lvwEmp.ListItems.count < 1 Then
        fraBrowse.Visible = False
        Exit Sub
    End If
    If From = True Then
        txtFromE.Text = lvwEmp.SelectedItem
    Else
        txtToE.Text = lvwEmp.SelectedItem
    End If

ElseIf trwStruc.Visible = True Then

Dim i As Long
    If trwStruc.Nodes.count < 2 Then
        Exit Sub
    End If
    
    ParentN = False
    If trwStruc.Nodes.count > 0 Then
    Set rs3 = CConnect.GetRecordSet("SELECT * FROM CStructure WHERE PCode = '" & myLevelCode & "'")

    With rs3
        If .RecordCount > 0 Then
            'MsgBox "You cannot select this level. You can only select the lowest level in your company's structure.", vbInformation

            'Exit Sub
            ParentN = True
        End If
    End With

    Set rs3 = Nothing

    txtDiv.Text = trwStruc.SelectedItem.Key

End If
End If
fraBrowse.Visible = False
From = False
myTo = False
Exit Sub
Hell:
MsgBox err.Description, vbExclamation
End Sub

Private Sub cmdToE_Click()
    Sel = txtToE.Text & ""
    popupText = "RTo2"
    frmPopUp.Show vbModal
    
End Sub



Private Sub dtpFrom_CloseUp()
If DateDiff("d", dtpFrom.value, dtpTo.value) < 0 Then
    dtpFrom.value = dtpTo.value
End If
End Sub

Private Sub dtpTo_CloseUp()
If DateDiff("d", dtpFrom.value, dtpTo.value) < 0 Then
    dtpFrom.value = dtpTo.value
End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    On Error GoTo Hell
    
    frmMain2.txtDetails.Caption = ""

    CConnect.CColor Me, MyColor

'    Call InitRConnection

    Call InitGrid
    

    Set rs6 = Nothing
    
    Call LoadList
    ''Call LoadCbo
    ''Call myStructure
    
    If RFilter = "Retirement" Or RFilter = "Termination" Then
        fraDates.Enabled = True
    '    fraDates.BackColor = &HFFC0C0
    '    chkDates.BackColor = &HFFC0C0
        chkDates.value = 1
    End If
    
    If RFilter = "Hires" Or ReportType = "Salary Increment 1.rpt" Then
        fraDates.Enabled = True
    '    fraDates.BackColor = &HFFC0C0
    '    chkDates.BackColor = &HFFC0C0
        chkDates.value = 1
    End If

    If RFilter = "BirthDay" Then
        fraMonth.Visible = True
    '    fraMonth.BackColor = &HFFC0C0
    End If
    
    If RFilter = "Bonus" Then
        fraBonus.Enabled = True
    '    fraBonus.BackColor = &HFFC0C0
        fraCategory.Visible = True
        fraMonth.Enabled = False
    End If
    
    If RFilter = "LongService" Then
        fraLongService.Enabled = True
    '    fraLongService.BackColor = &HFFC0C0
    End If
    
    If RFilter = "Mandatory" Then
        fraCheck.Visible = True
    '    fraCheck.BackColor = &HFFC0C0
    End If

    If RFilter = "Termination" Then
        fraTermination.Enabled = True
    '    fraTermination.BackColor = &HFFC0C0
    End If

    dtpTo.value = Date
    dtpFrom.value = Date
    PopulateDpts
    PopulatePaypoints
    PopulateCostCenters
Exit Sub
Hell:
MsgBox err.Description, vbExclamation

End Sub

Private Sub PopulateCostCenters()
    On Error GoTo ErrorHandler
    Dim rsPP As New ADODB.Recordset, ItemD As ListItem
    
    Set rsPP = CConnect.GetRecordSet("exec spget_costcenters '" & companyDetail.CompanyName & "'")
    If rsPP.RecordCount > 0 Then

        Do Until rsPP.EOF
        
            With lvwcostcenter
                Set ItemD = .ListItems.add(, , rsPP!Description)
                ItemD.Tag = rsPP!costcenter_ID
                ItemD.Checked = True
            End With
            rsPP.MoveNext
        Loop
    End If
    
    Set ItemD = lvwcostcenter.ListItems.add(1, , "(Select All)")
    ItemD.Tag = "S_A"
    lvwcostcenter.ListItems.Item(1).Checked = True
    ItemD.Bold = True
        
    Exit Sub
ErrorHandler:
    MsgBox "Error Loading Paypoints" & vbNewLine & err.Description, vbExclamation, "PDR Error"
End Sub


Private Sub PopulatePaypoints()
    On Error GoTo ErrorHandler
    Dim rsPP As New ADODB.Recordset, ItemD As ListItem
    
    Set rsPP = CConnect.GetRecordSet("Select * From tblpaypoint Order by paypoint_id")
    If rsPP.RecordCount > 0 Then

        Do Until rsPP.EOF
        
            With lvwpaypoint
                Set ItemD = .ListItems.add(, , rsPP!Paypoint_Name)
                ItemD.Tag = rsPP!Paypoint_ID
                ItemD.Checked = True
            End With
            rsPP.MoveNext
        Loop
    End If
    
    Set ItemD = lvwpaypoint.ListItems.add(1, , "(Select All)")
    ItemD.Tag = "S_A"
    lvwpaypoint.ListItems.Item(1).Checked = True
    ItemD.Bold = True
    
    
        
    Exit Sub
ErrorHandler:
    MsgBox "Error Loading Paypoints" & vbNewLine & err.Description, vbExclamation, "PDR Error"
End Sub

Private Sub Form_Resize()
'oSmart.FResize Me
End Sub

Private Sub lvwBanks_DblClick()
    Call cmdSelect_Click
End Sub

Private Sub lvwComp_DblClick()
Call cmdSelect_Click
End Sub

Private Sub lvwCur_DblClick()
    Call cmdSelect_Click
End Sub

Private Sub lvwDiv_DblClick()
    Call cmdSelect_Click
End Sub

Private Sub lsvCat_DblClick()
    If catFrom = True Then
        txtCatF.Text = lsvCat.SelectedItem.Text
        txtCatT.Text = lsvCat.SelectedItem.Text
    Else
        txtCatT.Text = lsvCat.SelectedItem.Text
    End If
    Call cmdCancelC_Click
End Sub

Private Sub lvpaypoint_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub lvpaypoint_ItemCheck(ByVal Item As MSComctlLib.ListItem)
  
End Sub

Private Sub lvwcostcenter_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ErrorHandler
    Dim n As Integer, State As Boolean
    
    State = False
    If lvwcostcenter.ListItems.Item(1).Checked = True Then State = True
    
    If Item.Tag = "S_A" Then
        'Uncheck All Departments
        For n = 2 To lvwcostcenter.ListItems.count
            lvwcostcenter.ListItems.Item(n).Checked = State
        Next n
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error Selecting cost centers:" & vbNewLine & err.Description, vbExclamation, "PDR Error"

End Sub

Private Sub LvwDpts_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo ErrorHandler
    Dim n As Integer, State As Boolean
    
    State = False
    If LvwDpts.ListItems.Item(1).Checked = True Then State = True
    
    If Item.Tag = "S-A" Then
        'Uncheck All Departments
        For n = 2 To LvwDpts.ListItems.count
            LvwDpts.ListItems.Item(n).Checked = State
        Next n
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error Selecting Departments:" & vbNewLine & err.Description, vbExclamation, "PDR Error"
End Sub

Private Sub lvwEmp_DblClick()
    Call cmdSelect_Click
End Sub

Private Sub lvwHse_DblClick()
    Call cmdSelect_Click
End Sub

Private Sub lvwins_DblClick()
    Call cmdSelect_Click
End Sub

Private Sub lvwPayrate_DblClick()
    Call cmdSelect_Click
End Sub

Private Sub lvwTrans_DblClick()
    Call cmdSelect_Click
End Sub

Private Sub optAll_Click()
txtThersh.Visible = False
chkSigned.Visible = False
chkNet.Visible = False
End Sub

Private Sub optBanks_Click()
chkDetailed.Visible = False
fraPeriod.Visible = False
End Sub

Private Sub SSTab1_DblClick()

End Sub

Private Sub trwStruc_Click()
    myLevelCode = trwStruc.SelectedItem.Key
    
End Sub

Private Sub trwStruc_DblClick()
    Call cmdSelect_Click
End Sub

Private Sub txtChequeNo_KeyPress(KeyAscii As Integer)
    If Len(txtChequeNo.Text) > 48 Then
        Beep
        MsgBox "You can't enter more than 50 characters", vbInformation
    End If
    
    Select Case KeyAscii
      Case Asc("A") To Asc("Z")
      Case Asc("a") To Asc("z")
      Case Asc("0") To Asc("9")
      Case Asc("/")
      Case Asc("-")
      Case Is = 8
      Case Else
          Beep
          KeyAscii = 0
    End Select
End Sub


Private Sub lvwpaypoint_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ErrorHandler
    Dim n As Integer, State As Boolean
    
    State = False
    If lvwpaypoint.ListItems.Item(1).Checked = True Then State = True
    
    If Item.Tag = "S_A" Then
        'Uncheck All Departments
        For n = 2 To lvwpaypoint.ListItems.count
            lvwpaypoint.ListItems.Item(n).Checked = State
        Next n
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error Selecting Paypoints:" & vbNewLine & err.Description, vbExclamation, "PDR Error"

End Sub

Private Sub txtDiv_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtSig_KeyPress(KeyAscii As Integer)
    If Len(txtSig.Text) > 48 Then
        Beep
        MsgBox "You can't enter more than 50 characters", vbInformation
    End If
    
    Select Case KeyAscii
      Case Asc("A") To Asc("Z")
      Case Asc("a") To Asc("z")
      Case Asc("0") To Asc("9")
      Case Asc("/")
      Case Asc("-")
      Case Asc(" ")
      Case Is = 8
      Case Else
          Beep
          KeyAscii = 0
    End Select
End Sub

Private Sub UpDown1_DownClick()
    If txtYear.Text = 1900 Then
        txtYear.Text = 2200
    Else
        txtYear.Text = Val(txtYear.Text) - 1
    End If

End Sub

Private Sub UpDown1_UpClick()
    If txtYear.Text = 2200 Then
        txtYear.Text = 1900
    Else
        txtYear.Text = Val(txtYear.Text) + 1
    End If
End Sub
Private Sub txtYear_Change()
    If Val(txtYear.Text) > 2200 Then
        txtYear.Text = 2200
    ElseIf Val(txtYear.Text) < 1900 Then
        txtYear.Text = 1900
    End If
End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
        Case Is = 8
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub
Private Sub optBankT_Click()
    txtThersh.Visible = False
    chkSigned.Visible = False
    chkNet.Visible = False
End Sub

Private Sub optBen_Click()
    chkDetailed.Visible = False
    fraPeriod.Visible = True
End Sub

Private Sub optCash_Click()
    txtThersh.Visible = False
    chkSigned.Visible = True
    chkNet.Visible = True
End Sub

Private Sub optCheque_Click()
    txtThersh.Visible = False
    chkSigned.Visible = True
    chkNet.Visible = True
End Sub

Private Sub optKin_Click()
    chkDetailed.Visible = False
    fraPeriod.Visible = False
End Sub

Private Sub optOther_Click()
    txtThersh.Visible = False
    chkSigned.Visible = False
    chkNet.Visible = False
End Sub

Private Sub optThers_Click()
    txtThersh.Visible = True
    chkSigned.Visible = False
    chkNet.Visible = False
End Sub

Private Sub OptTrans_Click()
    chkDetailed.Visible = False
    fraPeriod.Visible = True
End Sub

Private Sub txtFromE_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtOFrom_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtOTo_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtThersh_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case Is = 8
    Case Else
        Beep
        KeyAscii = 0
End Select
End Sub

Private Sub txtToE_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Public Sub InitGrid()

   Dim li As ListItem
    With lvwEmp
        .ColumnHeaders.Clear
        .ColumnHeaders.add , , "Code", 1300
        .ColumnHeaders.add , , "Name", 2800
        
        .View = lvwReport
    End With
End Sub

Public Sub LoadList()
    Dim i As Long
    


    AllEmployees.GetAccessibleEmployeesByUser currUser.UserID
    'Set rst1 = cConnect.GetPayData("Select * from Employee order by EmpCode")
    txtFromE.Text = AllEmployees.Item(1).EmpCode
    For i = 1 To AllEmployees.count
        If Not (AllEmployees.Item(i).IsDisengaged) Then 'IF CLAUSE ADDED BY JOHN TO TRAP OBJECT PRE-LOADS
            Set li = Me.lvwEmp.ListItems.add(, , AllEmployees.Item(i).EmpCode)
            li.SubItems(1) = AllEmployees.Item(i).SurName & "" & AllEmployees.Item(i).OtherNames
            li.Tag = AllEmployees.Item(i).EmployeeID
        End If
    Next i
    lblEmpCount.Caption = AllEmployees.count
   txtToE.Text = AllEmployees.Item(AllEmployees.count).EmpCode
End Sub

Public Sub LoadCbo()
If ReportType = "DivList" Then
    cboCompany.Clear
    cboDiv.Visible = True
    cboCompany.Visible = True
    lblOrange2.Visible = True
    lblORange.Caption = "Company"
    Set rs1 = CConnect.GetRecordSet("SELECT RLCode,LCode FROM Company ORDER BY LCode")
        With rs1
            If .RecordCount > 0 Then
                .MoveFirst
                While Not .EOF
                    cboCompany.AddItem !LCode & ""
                    .MoveNext
                Wend
                .MoveFirst
                cboCompany.Text = !LCode & ""
            End If
        End With
    
    Set rs1 = Nothing

    cboDiv.Visible = True
End If
    
    Set rs3 = CConnect.GetRecordSet("SELECT * FROM EmpTerms ORDER BY code")
    
    cboTerms.Clear
    lvwTerms.ListItems.Clear

    Dim li As ListItem
    With rs3
        If .RecordCount > 0 Then
            .MoveFirst
            cboTerms.Text = "All Records"
            cboTerms.AddItem "All Records"
            Do While Not .EOF
                cboTerms.AddItem !Description & ""
                Set li = lvwTerms.ListItems.add(, , !Description)
                li.Tag = !Code
                li.Checked = True
                .MoveNext
            Loop
        End If
    End With
    Set rs3 = Nothing
    
    'get payroll types
    
    Set rs3 = CConnect.GetRecordSet("SELECT * FROM tblpayroll ORDER BY payroll_id ")
    lvwpayrolltype.ListItems.Clear

   
    With rs3
        If .RecordCount > 0 Then
            .MoveFirst
        
            
            Do While Not .EOF
            
                Set li = lvwpayrolltype.ListItems.add(, , !Payroll_name)
                li.Tag = !payroll_id
                li.Checked = True
                .MoveNext
            Loop
        End If
    End With

End Sub

Public Sub Hidelvw()
    fraBrowse.Visible = True
    lvwEmp.Visible = False
    trwStruc.Visible = False
    lblECount.Visible = False
    lblEmpCount.Visible = False
End Sub

Public Sub AllocateNet()
    Dim myNet As Currency
    Dim myLabelNo As Integer
    Dim myAmt As Currency
    Dim myCount As Integer
    'delete from coinage 1st
    'Set rs2 = cconnect.GetRecordSet("SELECT Employee.EmpCode, Employee.CurCode,Payslips.Net" & _
    '    " FROM (Employee LEFT JOIN EmpPayment ON Employee.EmpCode = EmpPayment.EmpCode) LEFT JOIN Payslips ON Employee.EmpCode = Payslips.EmpCode" & _

    Set rs2 = CConnect.GetRecordSet("SELECT Employee.EmpCode, EmpPayment.CurCode, Payslips.Net " & _
         " FROM EmpPayment INNER JOIN (Employee LEFT JOIN Payslips ON Employee.EmpCode = Payslips.EmpCode) ON EmpPayment.EmpCode = Employee.EmpCode" & _
         " WHERE Payslips.MMonth=" & ReportPeriod & " AND Payslips.YYear=" & ReportYear & " AND Payslips.Net>0 AND EmpPayment.Cash='Yes'")
        With rs2
            If .RecordCount > 0 Then
                .MoveFirst
                myCount = 0
                frmProgress.Show , Me
                frmProgress.ProgressBar1.Max = .RecordCount
                While Not .EOF
                    frmProgress.Refresh
                    frmProgress.ProgressBar1.value = myCount
                    frmProgress.lblProgress.Caption = rs2!EmpCode & ""
                    myLabelNo = 1
                    myNet = 0
                    If Not IsNull(!Net) Then myNet = !Net
                    Set rs1 = CConnect.GetRecordSet("SELECT * FROM CurDenom WHERE CurCode='" & rs2!CurCode & "' ORDER BY MyValue DESC;")
                        With rs1
                            If .RecordCount > 0 Then
                                .MoveFirst
                                CConnect.ExecuteSql ("DELETE FROM CoinageAnalysis WHERE EmpCode='" & rs2!EmpCode & "'")
                                CConnect.ExecuteSql ("SET QUOTED IDENTIFIER ON INSERT INTO CoinageAnalysis (Empcode)VALUES (""" & rs2!EmpCode & """)")
                                While Not .EOF And myLabelNo <= 15 And myNet > 0
                                    If myNet >= !myValue And !myValue > 0 Then
                                        myAmt = Int(myNet / !myValue)
                                        CConnect.ExecuteSql ("UPDATE CoinageAnalysis SET Label" & myLabelNo & "=" & myAmt & ",NetPay=" & rs2!Net & " WHERE Empcode='" & rs2!EmpCode & "'")
                                        myNet = myNet - (myAmt * !myValue)
                                    End If
                                    myLabelNo = myLabelNo + 1
                                    .MoveNext
                                Wend
                            End If
                        End With
                    
                    Set rs1 = Nothing
                    .MoveNext
                    myCount = myCount + 1
                Wend
                Unload frmProgress
            End If
        End With
        
    Set rs2 = Nothing

End Sub
Public Sub myStructure()
    On Error GoTo Hell
    trwStruc.Nodes.Clear
    Set rs5 = CConnect.GetRecordSet("SELECT * FROM STypes WHERE SMain= 1 ORDER BY Code")
    Set MyNodes = trwStruc.Nodes.add(, , "O", rs5!Description & "")
    
    'Set rs = cconnect.GetRecordSet("SELECT * FROM CStructure WHERE SCode = '" & rs5!Code & "' ORDER BY MyLevel")
    Set rs1 = CConnect.GetRecordSet("SELECT STypes.Code, STypes.SMain, CStructure.PCode, CStructure.MyLevel, CStructure.LCode, CStructure.Description AS Division " & _
                " FROM CStructure INNER JOIN STypes ON CStructure.SCode = STypes.Code " & _
                " WHERE (((STypes.SMain)= 1))ORDER BY MyLevel")
    With rs1
        If .RecordCount > 0 Then
            .MoveFirst
            CNode = !LCode & ""
            
            Do While Not .EOF
                If !MyLevel = 0 Then
                     Set MyNodes = trwStruc.Nodes.add("O", tvwChild, !LCode & "", !Division & "")
                    MyNodes.EnsureVisible
                Else
                    Set MyNodes = trwStruc.Nodes.add(!PCode & "", tvwChild, !LCode & "", !Division & "")
                    MyNodes.EnsureVisible
                End If
                
                .MoveNext
            Loop
            .MoveFirst
            
            
        End If
    End With
    
    Set rs1 = Nothing
    Set rs5 = Nothing
Exit Sub
Hell:

End Sub

Public Sub SemiStructure(Full As Boolean)
    trwStruc.Nodes.Clear
    Dim myMain As Integer
    
    If Full = False Then
        Set MyNodes = trwStruc.Nodes.add(, , "OL" & "", "No Divisions")
        Exit Sub
    End If
    
    Set rs1 = CConnect.GetRecordSet("SELECT TLevel FROM CompDivisions WHERE LevelCode='" & myLevelCode & "'")
        With rs1
            If .RecordCount > 0 Then
                .MoveFirst
                If Not IsNull(!TLevel) Then myMain = !TLevel
            End If
        End With
        
    Set rs1 = Nothing
    
    Set rs = CConnect.GetRecordSet("SELECT * FROM CompDivisions WHERE LevelCode Like '" & myLevelCode & "%' ORDER BY TLevel")
    
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            CNode = !LevelCode & ""
    
            Do While Not .EOF
                If !TLevel = myMain Then
                    Set MyNodes = trwStruc.Nodes.add(, , !LevelCode & "", !Name & "")
                    MyNodes.EnsureVisible
                Else
                    Set MyNodes = trwStruc.Nodes.add(Mid(!LevelCode, 1, (Len(!LevelCode) - Len(!DivCode))), tvwChild, !LevelCode & "", !Name & "")
                    MyNodes.EnsureVisible
                End If
    
                .MoveNext
            Loop
            .MoveFirst
    
        End If
    End With
End Sub

Public Sub InitRConnection()
    Dim rec_t As New ADODB.Recordset
    Set rec_t = CConnect.GetRecordSet("SELECT servername, connectionname, dcatalog, userid, passwd FROM GeneralOpt WHERE subsystem = '" & SubSystem & "'")
    With rec_t
        If rec_t.EOF = False Then
            rConnection = Trim(!connectionName & "")
            rServer = Trim(!ServerName & "")
            rCatalog = Trim(!dcatalog & "")
            rUId = Trim(!UserID & "")
            rPass = Trim(!passwd & "")
        End If
    End With
End Sub

Private Sub PopulateDpts()
    On Error GoTo ErrorHandler
    Dim rsOUs As New ADODB.Recordset, ItemD As ListItem
    
    Set rsOUs = CConnect.GetRecordSet("Select * From OrganizationUnits Order by OrganizationUnitName")
    If rsOUs.RecordCount > 0 Then
        Do Until rsOUs.EOF
            With LvwDpts
                Set ItemD = .ListItems.add(, , rsOUs!OrganizationUnitName)
                ItemD.Tag = rsOUs!OrganizationUnitID
                ItemD.Checked = True
            End With
            rsOUs.MoveNext
        Loop
    End If
    
    Set ItemD = LvwDpts.ListItems.add(1, , "(Select All)")
    ItemD.Tag = "S-A"
    LvwDpts.ListItems.Item(1).Checked = True
    ItemD.Bold = True
        
    Exit Sub
ErrorHandler:
    MsgBox "Error Loading Departments" & vbNewLine & err.Description, vbExclamation, "PDR Error"
End Sub


Public Sub ShowReport(rpt As CRAXDDRT.Report, Optional formula As String, Optional blnAlterParamValue As Boolean)
    Dim conProps As CRAXDDRT.ConnectionProperties
    
    On Error GoTo ErrHandler
    'force crystal to use the basic syntax report
    
    frmMain2.MousePointer = vbHourglass
    
    If rpt.HasSavedData = True Then
        rpt.DiscardSavedData
    End If
    
    
    ' Loop through all database tables and set the correct server & database
        Dim tbl As CRAXDDRT.DatabaseTable
        Dim tbls As CRAXDDRT.DatabaseTables
        
        Set tbls = rpt.Database.Tables
        For Each tbl In tbls
            
            On Error Resume Next
            Set conProps = tbl.ConnectionProperties
            conProps.DeleteAll
            If tbl.DllName <> "crdb_ado.dll" Then
                tbl.DllName = "crdb_ado.dll"
            End If
                    
            conProps.add "Data Source", ConParams.DataSource
            conProps.add "Initial Catalog", ConParams.InitialCatalog
            conProps.add "Provider", "SQLOLEDB.1"
            'conProps.Add "Integrated Security", "true"
            conProps.add "User ID", ConParams.UserID
            conProps.add "Password", ConParams.Password
            conProps.add "Persist Security Info", "false"
            tbl.Location = tbl.Name
        Next tbl
        
        For Each tbl In tbls
        If (tbl.Name <> "tblCostCenter" Or tbl.Name = "prlEmployeePayrollInfo") Then
        tbl.Name = "E"
        End If
        
        
        Next tbl
        
         rpt.FormulaSyntax = crCrystalSyntaxFormula
         rpt.RecordSelectionFormula = mySQL
        'formula
    With rpt
    '        DEALING WITH ALTERED PARAM FIELD OBJECT VALUES
        If blnAlterParamValue = True Then
            For lngLoopVariable = LBound(objParamField) To UBound(objParamField)
                For lngloopvariable2 = 1 To .ParameterFields.count
                    If .ParameterFields.Item(lngloopvariable2).ParameterFieldName = objParamField(lngLoopVariable).Name Then
                        .ParameterFields.Item(lngloopvariable2).SetCurrentValue (objParamField(lngLoopVariable).value)
                        Exit For
                    End If
                    
                Next
            Next
        End If
        .EnableParameterPrompting = False
    End With
    rpt.PaperSize = crPaperA4
    
    If rpt.PaperOrientation = crLandscape Then
        rpt.BottomMargin = 192
        rpt.RightMargin = 720
        rpt.LeftMargin = 58
        rpt.TopMargin = 192
    ElseIf rpt.PaperOrientation = crPortrait Then
        rpt.BottomMargin = 300
        rpt.RightMargin = 338
        rpt.LeftMargin = 300
        rpt.TopMargin = 281
    End If
        With frmReports.CRViewer1
            .DisplayGroupTree = False
            .EnableAnimationCtrl = False
            .ReportSource = rpt
            .ViewReport
        End With
        
        formula = ""
        frmReports.Show
'        frmReports.Show vbModal
        frmMain2.MousePointer = vbNormal
    Exit Sub
ErrHandler:
    MsgBox err.Description, vbInformation
    frmMain2.MousePointer = vbNormal
End Sub
