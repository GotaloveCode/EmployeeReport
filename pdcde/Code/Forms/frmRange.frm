VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmRange 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Range"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   330
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
   HasDC           =   0   'False
   Icon            =   "frmRange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   13425
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lvwEmp 
      Height          =   4215
      Left            =   7080
      TabIndex        =   58
      Top             =   2160
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7435
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
      Enabled         =   0   'False
      NumItems        =   0
   End
   Begin VB.Frame fraEmployee 
      Appearance      =   0  'Flat
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
      Height          =   6120
      Left            =   6960
      TabIndex        =   60
      Top             =   360
      Width           =   6225
      Begin VB.CommandButton cmdfind2 
         Enabled         =   0   'False
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
         Left            =   5640
         Picture         =   "frmRange.frx":1762
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   1440
         Width           =   345
      End
      Begin VB.TextBox txtvalue 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4200
         TabIndex        =   73
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ComboBox cboCrieria 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   315
         ItemData        =   "frmRange.frx":1864
         Left            =   2640
         List            =   "frmRange.frx":186E
         TabIndex        =   71
         Top             =   1410
         Width           =   1455
      End
      Begin VB.ComboBox cboField 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   315
         ItemData        =   "frmRange.frx":187B
         Left            =   855
         List            =   "frmRange.frx":187D
         TabIndex        =   69
         Top             =   1410
         Width           =   1695
      End
      Begin VB.OptionButton optlist 
         Caption         =   "List"
         Height          =   195
         Left            =   120
         TabIndex        =   68
         Top             =   1440
         Width           =   1455
      End
      Begin VB.OptionButton optrange 
         Caption         =   "Range"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox txtFromE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1680
         TabIndex        =   64
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtToE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3960
         TabIndex        =   63
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
         Picture         =   "frmRange.frx":187F
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   720
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
         Left            =   5640
         Picture         =   "frmRange.frx":1981
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   720
         Width           =   345
      End
      Begin VB.Label lblvalue 
         Caption         =   "Value"
         Height          =   255
         Left            =   4200
         TabIndex        =   74
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Criteria"
         Height          =   255
         Left            =   2640
         TabIndex        =   72
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Search Field"
         Height          =   255
         Left            =   840
         TabIndex        =   70
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   5760
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Employment Terms"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   5400
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblERange 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Range"
         Height          =   315
         Left            =   240
         TabIndex        =   65
         Top             =   720
         Width           =   1260
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Employment date Filter"
      Height          =   735
      Left            =   480
      TabIndex        =   53
      Top             =   6720
      Width           =   5655
      Begin MSComCtl2.DTPicker dtto 
         Height          =   255
         Left            =   4200
         TabIndex        =   57
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   63963137
         CurrentDate     =   39905
      End
      Begin MSComCtl2.DTPicker dtfrom 
         Height          =   255
         Left            =   2640
         TabIndex        =   56
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   63963137
         CurrentDate     =   39905
      End
      Begin VB.OptionButton optbetween 
         Caption         =   "Between"
         Height          =   255
         Left            =   1320
         TabIndex        =   55
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optalldates 
         Caption         =   "All"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
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
      Height          =   6105
      Left            =   340
      TabIndex        =   44
      Top             =   360
      Width           =   3825
      Begin VB.CheckBox chkundefineddepts 
         Caption         =   "Include with Undefined depts"
         Height          =   255
         Left            =   240
         TabIndex        =   80
         Top             =   360
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin MSComctlLib.ListView LvwDpts 
         Height          =   5175
         Left            =   120
         TabIndex        =   45
         Top             =   840
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   9128
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
      Height          =   3090
      Left            =   360
      TabIndex        =   13
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   14
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
      TabIndex        =   28
      Top             =   2415
      Visible         =   0   'False
      Width           =   3825
      Begin VB.ComboBox cboBMonth 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmRange.frx":1A83
         Left            =   1995
         List            =   "frmRange.frx":1AAE
         TabIndex        =   29
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
         TabIndex        =   30
         Top             =   180
         Width           =   450
      End
   End
   Begin VB.Frame fraSelCat 
      Caption         =   "Employee Categories"
      Height          =   4815
      Left            =   240
      TabIndex        =   38
      Top             =   555
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CommandButton cmdCancelC 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   2640
         TabIndex        =   40
         Top             =   4440
         Width           =   1215
      End
      Begin MSComctlLib.ListView lsvCat 
         Height          =   4095
         Left            =   120
         TabIndex        =   39
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
      Left            =   11400
      TabIndex        =   0
      Top             =   7080
      Width           =   1020
   End
   Begin MSComDlg.CommonDialog Cdl 
      Left            =   660
      Top             =   4395
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   " Filters"
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
      Height          =   7740
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   13170
      Begin VB.CommandButton cmduncheckall 
         Caption         =   "Uncheck All"
         Height          =   255
         Left            =   6840
         TabIndex        =   79
         Top             =   6720
         Width           =   1575
      End
      Begin VB.Frame Frame6 
         Caption         =   "Employee Projects"
         Height          =   3975
         Left            =   13200
         TabIndex        =   76
         Top             =   360
         Width           =   3375
         Begin MSComctlLib.ListView lvwfundcodes 
            Height          =   1695
            Left            =   120
            TabIndex        =   78
            Top             =   2040
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   2990
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
            NumItems        =   0
         End
         Begin MSComctlLib.ListView lvwlocations 
            Height          =   1455
            Left            =   120
            TabIndex        =   77
            Top             =   240
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   2566
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
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Pypoint"
         Height          =   2055
         Left            =   3960
         TabIndex        =   51
         Top             =   4440
         Visible         =   0   'False
         Width           =   2535
         Begin MSComctlLib.ListView lvwPaypoint 
            Height          =   1695
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Visible         =   0   'False
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   2990
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
               Text            =   "Paypoints"
               Object.Width           =   5821
            EndProperty
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Payroll Type"
         Height          =   1935
         Left            =   3960
         TabIndex        =   49
         Top             =   2520
         Visible         =   0   'False
         Width           =   2655
         Begin MSComctlLib.ListView lvwpayrolltype 
            Height          =   1575
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Visible         =   0   'False
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   2778
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
               Text            =   "Payroll Type"
               Object.Width           =   5821
            EndProperty
         End
      End
      Begin VB.ComboBox cboTerms 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmRange.frx":1B16
         Left            =   1080
         List            =   "frmRange.frx":1B2C
         TabIndex        =   47
         Top             =   840
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Frame Frame2 
         Caption         =   "Employment Terms"
         Height          =   6015
         Left            =   3960
         TabIndex        =   46
         Top             =   360
         Width           =   2655
         Begin MSComctlLib.ListView lvwTerms 
            Height          =   5775
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   10186
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
               Text            =   "Employment terms"
               Object.Width           =   5821
            EndProperty
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
         TabIndex        =   32
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
            Picture         =   "frmRange.frx":1B70
            Style           =   1  'Graphical
            TabIndex        =   36
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
            Picture         =   "frmRange.frx":1C72
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   270
            Width           =   345
         End
         Begin VB.TextBox txtCatT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1665
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   645
            Width           =   1695
         End
         Begin VB.TextBox txtCatF 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1665
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   270
            Width           =   1695
         End
         Begin VB.Label Label9 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Category Range"
            Height          =   315
            Left            =   135
            TabIndex        =   37
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
         Height          =   1485
         Left            =   105
         TabIndex        =   20
         Top             =   4980
         Width           =   3825
         Begin VB.ComboBox cboBonus 
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   315
            ItemData        =   "frmRange.frx":1D74
            Left            =   1680
            List            =   "frmRange.frx":1D76
            TabIndex        =   21
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
            TabIndex        =   22
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
         TabIndex        =   15
         Top             =   4270
         Width           =   3825
         Begin MSComCtl2.DTPicker dtpLongService 
            Height          =   375
            Left            =   960
            TabIndex        =   43
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
            ItemData        =   "frmRange.frx":1D78
            Left            =   2520
            List            =   "frmRange.frx":1D94
            TabIndex        =   41
            Top             =   960
            Width           =   630
         End
         Begin VB.ComboBox cboLS 
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   315
            ItemData        =   "frmRange.frx":1DB4
            Left            =   1440
            List            =   "frmRange.frx":1DD0
            TabIndex        =   18
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
            TabIndex        =   42
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
            TabIndex        =   19
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
            TabIndex        =   17
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
            TabIndex        =   16
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
         TabIndex        =   8
         Top             =   3600
         Width           =   3825
         Begin VB.ComboBox cboTermReasons 
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   315
            ItemData        =   "frmRange.frx":1DF0
            Left            =   1680
            List            =   "frmRange.frx":1E03
            TabIndex        =   9
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
            TabIndex        =   10
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
         TabIndex        =   5
         Top             =   2640
         Width           =   3825
         Begin VB.CheckBox chkDates 
            BackColor       =   &H00F2FFFF&
            Caption         =   "Do not Use Date"
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   210
            Width           =   1110
         End
         Begin MSComCtl2.DTPicker dtpTo 
            Height          =   300
            Left            =   2160
            TabIndex        =   6
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
            TabIndex        =   7
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
            TabIndex        =   12
            Top             =   525
            Width           =   555
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "From"
            Height          =   315
            Left            =   1365
            TabIndex        =   11
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
         TabIndex        =   2
         Top             =   1800
         Visible         =   0   'False
         Width           =   3825
         Begin VB.TextBox txtDiv 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1650
            TabIndex        =   4
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
            Picture         =   "frmRange.frx":1E43
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   195
            Width           =   360
         End
      End
   End
   Begin VB.Label lblEmpCount 
      Caption         =   "Label8"
      Height          =   255
      Left            =   6000
      TabIndex        =   59
      Top             =   6480
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmRange"
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
Dim deptsexists As Boolean
Dim payrollsexists As Boolean
Dim deptscount As Integer
Dim payrollscount As Integer
Dim canAccessEmployees As Boolean

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

Private Sub cmdfind2_Click()
find2
End Sub

Private Sub cmdFromE_Click()
    Sel = txtFromE.Text & ""
    popupText = "RFrom"
    frmPopUp.Show vbModal
  gMinEmployeeID = gEmployeeID
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
    
    
    If (canAccessEmployees = False) Then
    MsgBox "The Loged In User Has no priviledge to Access the infomation of any Employee", vbInformation
    Unload Me
    Exit Sub
    End If
    
    If deptsexists = True Then
        If deptscount = 0 Then
        MsgBox ("The Currently Logged In User Has no priviledge to access employees in any Department")
        Exit Sub
        End If
    End If
    
    If payrollsexists = True Then
    If payrollscount = 0 Then
    MsgBox ("The Currently Logged In User Has no priviledge to access employees in any Payrolltype")
    End If
    End If
    
    Dim myfile As String
    Dim ss As String
    Dim MyL As Long
    Dim MyNo1 As Long
    Dim MyNo2 As Long
    Dim conProps As CRAXDDRT.ConnectionProperties
    
    MyL = Len(myLevelCode)
    
   
    Me.MousePointer = vbHourglass
    
    If lvwEmp.ListItems.count = 0 Then
    MsgBox ("Employee List is Empty.")
    Exit Sub
    End If
    
    
    If r.HasSavedData = True Then
        r.DiscardSavedData
    End If
    
    mySQL = ""
    
    Dim strEmp As String
    
    strEmp = ""
    
    
    
    
    
    
    
    
    
    '''*********************************
            If Not lvwTerms.ListItems.count = 0 Then
        Dim strTermID As String
        strTermID = ""
       
        k2 = 1
        While k2 <= lvwTerms.ListItems.count
          If lvwTerms.ListItems(k2).Checked = True Then
            
             strTermID = strTermID & lvwTerms.ListItems(k2).Tag & ","
            
          End If
        k2 = k2 + 1
        Wend
        If (strTermID <> "") Then
        strTermID = Mid(strTermID, 1, Len(strTermID) - 1)
        Else
        MsgBox ("Employement Term was not selected")
        Exit Sub
        End If
        End If
        ''
        
        ''''
        
        If Not lvwpayrolltype.ListItems.count = 0 Then
        Dim strPayrollid As String
        strPayrollid = ""
      
        k2 = 1
        While k2 <= lvwpayrolltype.ListItems.count
          If lvwpayrolltype.ListItems(k2).Checked = True Then
            
             strPayrollid = strPayrollid & lvwpayrolltype.ListItems(k2).Tag & ","
           
          End If
        k2 = k2 + 1
        Wend
        If (strPayrollid <> "") Then
        strPayrollid = Mid(strPayrollid, 1, Len(strPayrollid) - 1)
        Else
        If payrollsexists = False Then
        strPayrollid = 0
        Else
        MsgBox ("No Payroll Type was selected")
        End If
        Exit Sub
        End If
        End If
        
        ''''
        Dim strEmpDate As String
        strEmpDate = ""
        If (optbetween.value = True) Then
        strEmpDate = "between " & Format(dtfrom.value, "YYYY-MM-DD") & " and " & Format(dtto.value, "YYYY-MM-DD") & ""
        strEmpDate = " In DateTime " & Format(dtfrom.value, "(yyyy, mm, dd, hh, nn, ss)") & " to DateTime " & Format(dtto.value, "(yyyy, mm, dd, hh, nn, ss)") & ""
        End If
        
            ''''
        
'        If Not lvwpaypoint.ListItems.count = 0 Then
'        Dim strPaypoint As String
'        strPaypoint = ""
'
'        k2 = 1
''        While k2 <= lvwpaypoint.ListItems.count
''         If k2 > 1 Then
''          If lvwpaypoint.ListItems(k2).Checked = True Then
''
''             strPaypoint = strPaypoint & lvwpaypoint.ListItems(k2).Tag & ","
''
''          End If
''          End If
''        k2 = k2 + 1
''        Wend
''         If (strPaypoint <> "") Then
''         strPaypoint = Mid(strPaypoint, 1, Len(strPaypoint) - 1)
''         Else
''         MsgBox ("No Paypoint was selected")
''         Exit Sub
''         End If
'        End If
        
        ''''
        
        
        
        
'        If cboTerms.Text <> "" And cboTerms.Text <> "All Records" Then
'            If mySQL <> "" Then
'                'mySQL = mySQL & " " & "AND {Employee.Terms} = '" & cboTerms.Text & "'"
'                mySQL = mySQL & " " & "AND {Employee.TermsID} in [" & strTermID & "]"
'            Else
'               ' mySQL = "{Employee.TermsID} = '" & cboTerms.Text & "'"
'                mySQL = "{Employee.TermsID} in [" & strTermID & "]"
'            End If
'
''            If cboTerms.Text <> "Disengaged" Then
''                mySQL = mySQL & " AND {Employee.Term} = false"
''            End If
'            Debug.Print mySQL
'
'        Else
'            'mySQL = mySQL & " AND {Employee.Term} = false"
'            mySQL = mySQL & " AND {Employee.TermsID} in [" & strTermID & "]"
'        End If
        
'        If (strPayrollid <> "") Then
'          If mySQL <> "" Then
'          mySQL = mySQL & " " & "AND {Employee.Payroll_ID} in [" & strPayrollid & "]"
'          Else
'          mySQL = "{Employee.Payroll_ID} in [" & strPayrollid & "]"
'          End If
'        End If
    
'        If (strPaypoint <> "") Then
'          If mySQL <> "" Then
'          mySQL = mySQL & " " & "AND {Employee.Paypoint_ID} in [" & strPaypoint & "]"
'          Else
'          mySQL = "{Employee.Paypoint_ID} in [" & strPaypoint & "]"
'          End If
'        End If
    
        If (strEmpDate <> "") Then
          If mySQL <> "" Then
          mySQL = mySQL & " " & "AND ({Employee.DEmployed} " & strEmpDate & ")"
          Else
          mySQL = "({Employee.DEmployed} " & strEmpDate & ")"
          End If
        End If
        
        
        '''dept
        
        Dim depts As String
           depts = ""
        For k = 2 To LvwDpts.ListItems.count
            If (LvwDpts.ListItems.Item(k).Checked = True) Then
                depts = depts & "'" & Trim(Replace(LvwDpts.ListItems.Item(k).Text, "'", "''")) & "',"
            End If
        Next k
        If (Trim(depts) = "") Then
            If deptsexists = True Then
            MsgBox "You must have at least one department selected", vbExclamation, "Report Error"
            Exit Sub
            Else
            depts = "Un Defined"
            End If
        Else
            depts = Mid$(depts, 1, Len(depts) - 1)
        End If
        
        
'        If (strPaypoint <> "") Then
'          If mySQL <> "" Then
'            'If chkundefineddepts.value = 0 Then
'            mySQL = mySQL & " " & "AND {Employee.Paypoint_ID} in [" & strPaypoint & "]"
''            Else
''
''            End If
'          Else
'            'If chkundefineddepts.value = 0 Then
'            mySQL = "{Employee.Paypoint_ID} in [" & strPaypoint & "]"
''            Else
''
''            End If
'          End If
'        End If
    
        If (Trim(depts) <> "") Then
          If mySQL <> "" Then
          
            mySQL = mySQL & " " & "AND ({Employee.organizationname} in [" & depts & "])"
         
          Else
           
            mySQL = "({Employee.organizationname} in [" & depts & "])"
          
          End If
        End If
  
    ''**********************************
    
    If RFilter <> "Audit Trail" Then
    
    If (optrange.value = True) Then
       If Not txtFromE.Text = "" Then
            strEmp = "(" & gMinEmployeeID & " and " & gMinEmployeeID & ")"
            
            
'                   If Not (gAccessRightClassIds = "") Then '' user is not infiniti
'                   If UCase(gAccessRightName) = "PAYROLL TYPES" Then
'                        If Trim(strEmp) <> "" Then
'
'                                strEmp = strEmp & " and {Employee.payroll_id} in [" & gAccessRightClassIds & "]"
'                        Else
'                                strEmp = "{employee.payroll_id} in [" & gAccessRightClassIds & "]"
'                        End If
'                    End If
'                    End If

             If Trim(mySQL) = "" Then
             mySQL = "{Employee.Employee_ID} >= " & gMinEmployeeID & " and  {Employee.Employee_ID} <=" & gMaxEmployeeID & ""
             Else
             mySQL = mySQL & " and " & "{Employee.Employee_ID} >= " & gMinEmployeeID & " and  {Employee.Employee_ID} <=" & gMaxEmployeeID & ""
             End If
         
        End If
    Else
    
    
        If Not lvwEmp.ListItems.count = 0 Then
    
        strEmp = ""
       '' Dim k2 As Integer
        k2 = 1
         n = lvwEmp.ListItems.count
                If n > 1000 Then n = 1000
                        While k2 <= n
                          If lvwEmp.ListItems(k2).Checked = True Then
                            
                             strEmp = strEmp & "'" & lvwEmp.ListItems(k2).Text & "',"
                            
                          End If
                        k2 = k2 + 1
                        Wend
                        If (Trim(strEmp) <> "") Then
                        
                        
                           
                        
                        
                        strEmp = Mid(strEmp, 1, Len(strEmp) - 1)
                        mySQL = "{Employee.EmpCode} in ['" & strEmp & " ']"
                        
                       
                        Else
                        MsgBox ("No Employee was selected")
                        Exit Sub
                        End If
                End If
    


    
    
    
    
        End If
        
        
        ''
        
        
        If optlist.value = True Then
        Dim i2 As Long
        Dim chk As Boolean
        chk = False
        For i2 = 1 To lvwEmp.ListItems.count
        
        If (lvwEmp.ListItems(i2).Checked) Then
        chk = True
        End If
        
        Next i2
        If (chk = False) Then
        MsgBox ("You chose the list option of employees filter but didn't check any Employee")
        Exit Sub
        End If
        End If
        
''''''''''''''''''''''**************

'''''''''''''''''''''*****************
    
    
      
    
        If ParentN = True Then
            If txtDiv.Text <> "" Then
                If mySQL <> "" Then
                    mySQL = mySQL & " " & "AND {MyDivisions.PCode} Like '" & myLevelCode & "'"
            
                Else
                    mySQL = "{MyDivisions.PCode} Like '" & myLevelCode & "'"
                End If
            End If
        Else
            If txtDiv.Text <> "" Then
                If mySQL <> "" Then
                    mySQL = mySQL & " " & "AND {SEmp.LCode} Like '" & myLevelCode & "'"
            
                Else
                    mySQL = "{SEmp.LCode} Like '" & myLevelCode & "'"
                End If
            End If
        End If
        
        If chkDates.value = 0 Then
            If RFilter = "Hires" Then
                If mySQL <> "" Then
                    mySQL = mySQL & " " & "AND {Employee.DEmployed} In DateTime " & Format(dtpFrom.value, "(yyyy, mm, dd, hh, nn, ss)") & " to DateTime " & Format(dtpTo.value, "(yyyy, mm, dd, hh, nn, ss)") & ""
                Else
                    mySQL = "{Employee.DEmployed} In DateTime " & Format(dtpFrom.value, "(yyyy, mm, dd, hh, nn, ss)") & " to DateTime " & Format(dtpTo.value, "(yyyy, mm, dd, hh, nn, ss)") & ""
                End If
            
            End If
            
        End If
    
        If chkDates.value = 0 Then
            If RFilter = "Retirement" Or RFilter = "Termination" Then
                If mySQL <> "" Then
                    mySQL = mySQL & " " & "AND {@Retire} In DateTime " & Format(dtpFrom.value, "(yyyy, mm, dd, hh, nn, ss)") & " to DateTime " & Format(dtpTo.value, "(yyyy, mm, dd, hh, nn, ss)") & ""
                Else
                    mySQL = "{@Retire} In DateTime " & Format(dtpFrom.value, "(yyyy, mm, dd, hh, nn, ss)") & " to DateTime " & Format(dtpTo.value, "(yyyy, mm, dd, hh, nn, ss)") & ""
                End If
            
            End If
        End If
    
        If RFilter = "Termination" Then
            If cboTermReasons.Text <> "" Then
                If mySQL <> "" Then
                    mySQL = mySQL & " " & "AND {Employee.TermReasons} = '" & cboTermReasons.Text & "'"
                Else
                    mySQL = "{Employee.TermReasons} = '" & cboTermReasons.Text & "'"
                End If
            End If
            
            ''''------------------------------------------------
            ''If cboTerms.Text <> "" And cboTerms.Text <> "All Records" Then
            If (strTermID <> "") Then
               If mySQL <> "" Then
                'mySQL = mySQL & " " & "AND {Employee.Terms} = '" & cboTerms.Text & "'"
                mySQL = mySQL & " " & "AND {Employee.TermsID} in [" & strTermID & "]"
               Else
               ' mySQL = "{Employee.TermsID} = '" & cboTerms.Text & "'"
                mySQL = "{Employee.TermsID} in [" & strTermID & "]"
               End If
            
'               If cboTerms.Text <> "Disengaged" Then
'                mySQL = mySQL & " AND {Employee.Term} = false"
'               End If
               Debug.Print mySQL
        
'              Else
'               'mySQL = mySQL & " AND {Employee.Term} = false"
'                mySQL = mySQL & " AND {Employee.TermsID} in [" & strTermID & "]"
'               End If
             End If
'        If (strPayrollid <> "") Then
'          If mySQL <> "" Then
'          mySQL = mySQL & " " & "AND {Employee.Payroll_ID} in [" & strPayrollid & "]"
'          Else
'          mySQL = "{Employee.Payroll_ID} in [" & strPayrollid & "]"
'          End If
'        End If
    
'        If (strPaypoint <> "") Then
'          If mySQL <> "" Then
'          mySQL = mySQL & " " & "AND {Employee.Paypoint_ID} in [" & strPaypoint & "]"
'          Else
'          mySQL = "{Employee.Paypoint_ID} in [" & strPaypoint & "]"
'          End If
'        End If
        
         If (depts <> "") Then
         If (mySQL <> "") Then
         mySQL = mySQL & " AND {Employee.ORGANIZATIONUNITNAME}  in [" & tSQL & "]"
         Else
         mySQL = "{Employee.ORGANIZATIONUNITNAME}  in [" & tSQL & "]"
         End If
         End If

            ''''------------------------------------------------
        
        End If
        
        If RFilter = "BirthDay" Then
            If cboBMonth.Text <> "" Then
                If mySQL <> "" Then
                    mySQL = mySQL & " " & "AND {@Month} = '" & cboBMonth.Text & "'"
                Else
                    mySQL = "{@Month} = '" & cboBMonth.Text & "'"
                End If
            
            End If
            
            
            
            '''''-------------------------------------
            
            ''If cboTerms.Text <> "" And cboTerms.Text <> "All Records" Then
            If (strTermID <> "") Then
            If mySQL <> "" Then
                'mySQL = mySQL & " " & "AND {Employee.Terms} = '" & cboTerms.Text & "'"
                mySQL = mySQL & " " & "AND {Employee.TermsID} in [" & strTermID & "]"
            Else
               ' mySQL = "{Employee.TermsID} = '" & cboTerms.Text & "'"
                mySQL = "{Employee.TermsID} in [" & strTermID & "]"
            End If
            
'            If cboTerms.Text <> "Disengaged" Then
'                mySQL = mySQL & " AND {Employee.Term} = false"
'            End If
            Debug.Print mySQL
        
'            Else
'            'mySQL = mySQL & " AND {Employee.Term} = false"
'            mySQL = mySQL & " AND {Employee.TermsID} in [" & strTermID & "]"
'            End If
            End If
'        If (strPayrollid <> "") Then
'          If mySQL <> "" Then
'          mySQL = mySQL & " " & "AND {Employee.Payroll_ID} in [" & strPayrollid & "]"
'          Else
'          mySQL = "{Employee.Payroll_ID} in [" & strPayrollid & "]"
'          End If
'        End If
    
'        If (strPaypoint <> "") Then
'          If mySQL <> "" Then
'          mySQL = mySQL & " " & "AND {Employee.Paypoint_ID} in [" & strPaypoint & "]"
'          Else
'          mySQL = "{Employee.Paypoint_ID} in [" & strPaypoint & "]"
'          End If
'        End If

       If (depts <> "") Then
         If (mySQL <> "") Then
         mySQL = mySQL & " AND {Employee.ORGANIZATIONUNITNAME}  in [" & tSQL & "]"
         Else
         mySQL = "{Employee.ORGANIZATIONUNITNAME}  in [" & tSQL & "]"
         End If
       End If

            
            '''''-------------------------------------
            
        End If
        
        If RFilter = "Bonus" Then
            If cboBonus.Text <> "" Then
                If mySQL <> "" Then
                    mySQL = mySQL & " " & "AND {Employee.BPerc} = " & Left(cboBonus.Text, InStr(cboBonus.Text, "%") - 1) & ""
                Else
                    mySQL = "{Employee.BPerc} = " & Left(cboBonus.Text, Mid(cboBonus.Text, "%")) & ""
                End If
            End If
            If txtCatF.Text <> "" Then
                mySQL = mySQL & " AND {Employee.ECategory} in '" & txtCatF.Text & "' to '" & txtCatT.Text & "'"
            End If
            
            ''''''''''--------------------------------
            ''If cboTerms.Text <> "" And cboTerms.Text <> "All Records" Then
            If (strTermID <> "") Then
               If mySQL <> "" Then
                'mySQL = mySQL & " " & "AND {Employee.Terms} = '" & cboTerms.Text & "'"
                mySQL = mySQL & " " & "AND {Employee.TermsID} in [" & strTermID & "]"
               Else
               ' mySQL = "{Employee.TermsID} = '" & cboTerms.Text & "'"
                mySQL = "{Employee.TermsID} in [" & strTermID & "]"
               End If
            
'            If cboTerms.Text <> "Disengaged" Then
'                mySQL = mySQL & " AND {Employee.Term} = false"
'            End If
              Debug.Print mySQL
        
'              Else
'            'mySQL = mySQL & " AND {Employee.Term} = false"
'              mySQL = mySQL & " AND {Employee.TermsID} in [" & strTermID & "]"
'              End If
        End If
        
'        If (strPayrollid <> "") Then
'          If mySQL <> "" Then
'          mySQL = mySQL & " " & "AND {Employee.Payroll_ID} in [" & strPayrollid & "]"
'          Else
'          mySQL = "{Employee.Payroll_ID} in [" & strPayrollid & "]"
'          End If
'        End If
    
'        If (strPaypoint <> "") Then
'          If mySQL <> "" Then
'          mySQL = mySQL & " " & "AND {Employee.Paypoint_ID} in [" & strPaypoint & "]"
'          Else
'          mySQL = "{Employee.Paypoint_ID} in [" & strPaypoint & "]"
'          End If
'        End If

        If (depts <> "") Then
         If (mySQL <> "") Then
         mySQL = mySQL & " AND {Employee.ORGANIZATIONUNITNAME}  in [" & tSQL & "]"
         Else
         mySQL = "{Employee.ORGANIZATIONUNITNAME}  in [" & tSQL & "]"
         End If
        End If


            
            '''''''''--------------------------------
        End If
        
        Dim MyName As String
        
        If RFilter = "LongService" Then
            
            r.ParameterFields(1).ClearCurrentValueAndRange
            r.ParameterFields(1).AddCurrentValue (dtpLongService.value)
            r.EnableParameterPrompting = False

            If cboLS.Text <> "" Then
                If IsNumeric(cboLS.Text) And IsNumeric(cboLST.Text) Then
                    If mySQL <> "" Then
                        mySQL = mySQL & " AND {@Yrs} >= " & cboLS.Text & " AND  {@Yrs} <= " & cboLST.Text
                    Else
                        mySQL = "{@Yrs} >= " & cboLS.Text & " AND  {@Yrs} <= " & cboLST.Text
                    End If
                Else
                    MsgBox "Enter a numeric value for the report range.", vbInformation
                    Exit Sub
                End If
                
                If cboLS.Text <> cboLST.Text Then
                    r.ReportComments = "" & cboLS.Text & " to " & cboLST.Text & " Years Service in " & Format(dtpLongService.value, "MMM yyyy")
                Else
                    r.ReportComments = "" & cboLS.Text & " Years Service in " & Format(dtpLongService.value, "MMM yyyy")
                End If
                
            End If
        
                  ''''''''''--------------------------------
            ''If cboTerms.Text <> "" And cboTerms.Text <> "All Records" Then
            If (strTermID <> "") Then
              If mySQL <> "" Then
                'mySQL = mySQL & " " & "AND {Employee.Terms} = '" & cboTerms.Text & "'"
                mySQL = mySQL & " " & "AND {Employee.TermsID} in [" & strTermID & "]"
              Else
               ' mySQL = "{Employee.TermsID} = '" & cboTerms.Text & "'"
                mySQL = "{Employee.TermsID} in [" & strTermID & "]"
              End If
            
'              If cboTerms.Text <> "Disengaged" Then
'                mySQL = mySQL & " AND {Employee.Term} = false"
'               End If
               Debug.Print mySQL
        
'             Else
'               'mySQL = mySQL & " AND {Employee.Term} = false"
'                mySQL = mySQL & " AND {Employee.TermsID} in [" & strTermID & "]"
'             End If
              End If
              
              
             If (depts <> "") Then
             If (mySQL <> "") Then
             mySQL = mySQL & " AND {Employee.ORGANIZATIONUNITNAME}  in [" & tSQL & "]"
             Else
             mySQL = "{Employee.ORGANIZATIONUNITNAME}  in [" & tSQL & "]"
            End If
            End If
            
            
'        If (strPayrollid <> "") Then
'          If mySQL <> "" Then
'          mySQL = mySQL & " " & "AND {Employee.Payroll_ID} in [" & strPayrollid & "]"
'          Else
'          mySQL = "{Employee.Payroll_ID} in [" & strPayrollid & "]"
'          End If
'        End If
    
'        If (strPaypoint <> "") Then
'          If mySQL <> "" Then
'          mySQL = mySQL & " " & "AND {Employee.Paypoint_ID} in [" & strPaypoint & "]"
'          Else
'          mySQL = "{Employee.Paypoint_ID} in [" & strPaypoint & "]"
'          End If
'        End If

            
            '''''''''--------------------------------
        
        
        End If
    End If
    
    Dim Mand As String
    
    If RFilter = "Mandatory" Then
        If chkNHIF.value = 1 Then
            Mand = "ISNULL({Employee.NHIFNo}) OR {Employee.NHIFNo} = ''"
        End If
        
        If chkNSSF.value = 1 Then
            If Mand <> "" Then
                Mand = Mand & " " & "OR ISNULL({Employee.NSSFNo}) OR {Employee.NSSFNo} = ''"
            Else
                Mand = "ISNULL({Employee.NSSFNo}) OR {Employee.NSSFNo} = ''"
            End If
        End If
        
        If chkPIN.value = 1 Then
            If Mand <> "" Then
                Mand = Mand & " " & "OR ISNULL({Employee.PINNo}) OR {Employee.PINNo} = ''"
            Else
                Mand = "ISNULL({Employee.PINNo}) OR {Employee.PINNo} = ''"
            End If
        End If
        
        If chkCert.value = 1 Then
            If Mand <> "" Then
                Mand = Mand & " " & "OR ISNULL({Employee.CertNo}) OR {Employee.CertNo} = ''"
            Else
                Mand = "ISNULL({Employee.CertNo}) OR {Employee.CertNo} = ''"
            End If
        End If
        
        If chkBasic.value = 1 Then
            If Mand <> "" Then
                Mand = Mand & " " & "OR {Employee.BasicPay} = 0"
            Else
                Mand = "{Employee.BasicPay} = 0"
            End If
        End If
        
        If chkDesig.value = 1 Then
            If Mand <> "" Then
                Mand = Mand & " " & "OR ISNULL({Employee.Desig}) OR {Employee.Desig} = ''"
            Else
                Mand = "ISNULL({Employee.Desig}) OR {Employee.Desig} = ''"
            End If
'            If Mand <> "" Then
'                Mand = Mand & " " & "OR {Employee.Desig} OR {Employee.Desig} = ''"
'            Else
'                Mand = "{Employee.Desig} OR {Employee.Desig} = ''"
'            End If
        End If
          
        If mySQL <> "" And Mand <> "" Then
            mySQL = mySQL & " AND " & Mand
        Else
            mySQL = Mand
        End If
        
                ''''''''''--------------------------------
           '' If cboTerms.Text <> "" And cboTerms.Text <> "All Records" Then
            If (strTermID <> "") Then
              If mySQL <> "" Then
                'mySQL = mySQL & " " & "AND {Employee.Terms} = '" & cboTerms.Text & "'"
                mySQL = mySQL & " " & "AND {Employee.TermsID} in [" & strTermID & "]"
              Else
               ' mySQL = "{Employee.TermsID} = '" & cboTerms.Text & "'"
                mySQL = "{Employee.TermsID} in [" & strTermID & "]"
              End If
            
'              If cboTerms.Text <> "Disengaged" Then
'                mySQL = mySQL & " AND {Employee.Term} = false"
'              End If
               Debug.Print mySQL
        
'              Else
'              'mySQL = mySQL & " AND {Employee.Term} = false"
'              mySQL = mySQL & " AND {Employee.TermsID} in [" & strTermID & "]"
'              End If
            End If
        
        
'        If (strPayrollid <> "") Then
'          If mySQL <> "" Then
'          mySQL = mySQL & " " & "AND {Employee.Payroll_ID} in [" & strPayrollid & "]"
'          Else
'          mySQL = "{Employee.Payroll_ID} in [" & strPayrollid & "]"
'          End If
'        End If
'
'        If (strPaypoint <> "") Then
'          If mySQL <> "" Then
'          mySQL = mySQL & " " & "AND {Employee.Paypoint_ID} in [" & strPaypoint & "]"
'          Else
'          mySQL = "{Employee.Paypoint_ID} in [" & strPaypoint & "]"
'          End If
'        End If

            
            '''''''''--------------------------------
        
        
    End If
    
    '==========General Details Report=========
    If RFilter = "General" Then
     tSQL = ""
        mySQL = ""
        If (optrange.value = True) Then
         ''mySQL = "({Employee.EmpCode} in ['" & txtFromE.Text & "' to '" & txtToE.Text & "'])"
          mySQL = "{Employee.Employee_ID} >= " & gMinEmployeeID & " and  {Employee.Employee_ID} <=" & gMaxEmployeeID & ""
          
          
'                                If Not (gAccessRightClassIds = "") Then '' user is not infiniti
'                                   If UCase(gAccessRightName) = "PAYROLL TYPES" Then
'                                        If Trim(mySQL) <> "" Then
'
'                                                mySQL = mySQL & " and {Employee.payroll_id} in [" & gAccessRightClassIds & "]"
'                                        Else
'                                                mySQL = "{employee.payroll_id} in [" & gAccessRightClassIds & "]"
'                                        End If
'                                    End If
'                                 End If

        Else
          If (Trim(strEmp) <> "") Then
          mySQL = "{Employee.EmpCode} in [" & strEmp & " ]"
          Else
          MsgBox ("No Employee was selected")
          Exit Sub
          End If
        End If
     
        
        If Not txtFromE.Text = "" Then
           '' mySQL = "{Employee.EmpCode} in '" & txtFromE.Text & "' to '" & txtToE.Text & "' AND {Employee.Term} = FALSE "
            Dim L As Integer, LSQL As String
            
            n = lvwEmp.ListItems.count
            If n > 1000 Then n = 1000
           
            'LSQL = Mid$(LSQL, 1, Len(tSQL) - 1)
            If (mySQL <> "") Then
           
                       mySQL = mySQL & "  AND {Employee.Term} = FALSE and {Employee.Disengaged} = FALSE"
            Else
                       mySQL = "{Employee.Term} = FALSE and {Employee.Disengaged} = FALSE"
            End If
        End If
        
         If (strEmpDate <> "") Then
          If mySQL <> "" Then
          mySQL = mySQL & " " & "AND  {Employee.DEmployed} " & strEmpDate & ""
          Else
          mySQL = "{Employee.DEmployed} " & strEmpDate & ""
          End If
        End If
        
        
        ''----------inserted by kalya
          '' If cboTerms.Text <> "" And cboTerms.Text <> "All Records" Then
           If (strTermID <> "") Then
             If mySQL <> "" Then
                'mySQL = mySQL & " " & "AND {Employee.Terms} = '" & cboTerms.Text & "'"
                mySQL = mySQL & " " & "AND {Employee.TermsID} in [" & strTermID & "]"
             Else
               ' mySQL = "{Employee.TermsID} = '" & cboTerms.Text & "'"
                mySQL = "{Employee.TermsID} in [" & strTermID & "]"
             End If
            
'            If cboTerms.Text <> "Disengaged" Then
'                mySQL = mySQL & " AND {Employee.Term} = false"
'            End If
              Debug.Print mySQL
           End If
'           Else
'            'mySQL = mySQL & " AND {Employee.Term} = false"
'            mySQL = mySQL & " AND {Employee.TermsID} in [" & strTermID & "]"
'           End If
        
'          If (strPayrollid <> "") Then
'           If mySQL <> "" Then
'           mySQL = mySQL & " " & "AND {Employee.Payroll_ID} in [" & strPayrollid & "]"
'           Else
'           mySQL = "{Employee.Payroll_ID} in [" & strPayrollid & "]"
'           End If
'          End If
'
'
'           If (strPaypoint <> "") Then
'            If mySQL <> "" Then
'             mySQL = mySQL & " " & "AND {Employee.Paypoint_ID} in [" & strPaypoint & "]"
'            Else
'            mySQL = "{Employee.Paypoint_ID} in [" & strPaypoint & "]"
'            End If
'           End If


        If (depts <> "") Then
         If (mySQL <> "") Then
         mySQL = mySQL & " AND {Employee.ORGANIZATIONUNITNAME}  in [" & depts & "]"
         Else
         mySQL = "{Employee.ORGANIZATIONUNITNAME}  in [" & depts & "]"
         End If
         End If
          
          
'          If (strEmpDate <> "") Then
'            If mySQL <> "" Then
'             mySQL = mySQL & " " & "AND ({Employee.DEmployed} between " & strEmpDate & ")"
'            Else
'             mySQL = "({Employee.DEmployed} between " & strEmpDate & ")"
'            End If
'          End If
'           Dim k As Integer, tSQL As String
'        For k = 2 To LvwDpts.ListItems.Count
'            If (LvwDpts.ListItems.Item(k).Checked = True) Then
'                tSQL = tSQL & "'" & Trim(Replace(LvwDpts.ListItems.Item(k).Text, "'", "''")) & "',"
'            End If
'        Next k
'
'        If (Trim(tSQL) = "") Then
'            MsgBox "You must have at least one department selected", vbExclamation, "Report Error"
'            Exit Sub
'        Else
'            tSQL = Mid$(tSQL, 1, Len(tSQL) - 1)
'            If mySQL <> "" Then
'            mySQL = mySQL & " AND {Employee.OrganizationUnitName} in [" & tSQL & "]"
'            Else
'             mySQL = "{Employee.OrganizationUnitName} in [" & tSQL & "]"
'            End If
'        End If
        '''-----------end inserted by kalya
        
    End If
    '==========End Of General Details Report==
    
    '==========Next of Kin Report=========
    If RFilter = "NextOFKin" Then
        mySQL = ""
        
        
        If (optrange.value = True) Then
         ''mySQL = "({Employee.EmpCode} in ['" & txtFromE.Text & "' to '" & txtToE.Text & "'])"
          mySQL = "{Employee.Employee_ID} >= " & gMinEmployeeID & " and  {Employee.Employee_ID} <=" & gMaxEmployeeID & "  AND {Employee.Disengaged} = FALSE"

        Else
          If (Trim(strEmp) <> "") Then
          mySQL = "({pdvwNextOfKin.EmpCode} in '" & txtFromE.Text & "' to '" & txtToE.Text & "' AND {Employee.Disengaged} = FALSE)"
          mySQL = "{Employee.EmpCode} in [" & strEmp & " ]"
          Else
          MsgBox ("No Employee was selected")
          Exit Sub
          End If
        End If
     
        
'        If Not txtFromE.Text = "" Then
'            mySQL = "({pdvwNextOfKin.EmpCode} in '" & txtFromE.Text & "' to '" & txtToE.Text & "' AND {pdvwNextOfKin.IsDisengaged} = FALSE)"
'        End If
    End If
    '==========End Of Next of Kin Report==
    
    '==========CHECK FILTER BY DEPARTMENT=================================
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
            
            ''--------edited by kalya
            
            If RFilter = "NextOFKin" Then
'            If myreport = "crtCheck" Then
'             mySQL = mySQL & " AND ({Employee.OrganizationUnitName} in [" & tSQL & "])"
'            Else
'            mySQL = mySQL & " AND ({pdvwNextOfKin.Department} in [" & tSQL & "])"
'            End If
                
            Else
            
'            If myreport = "crtCheck" Or myreport = "chkEmploymentHistory" Then
'            mySQL = mySQL & " AND {Employee.OrganizationUnitName} in [" & tSQL & "]"
'            Else
'                mySQL = mySQL & " AND {MyDivisions.Department} in [" & tSQL & "]"
'            End If
'            If (tSQL <> "") Then
'            mySQL = mySQL & " AND {Employee.OrganizationUnitName} in [" & tSQL & "]"
'            End If
            End If
            
            
            ''--------end edited by kalya
            
        End If
        
        ''''
    '==========END CHECK FILTER BY DEPARTMENT=============================
    
    
     '==========CHECK FILTER BY EMPLOYEES ACCESSED=================================
'    Dim L As Integer, LSQL As String
'
'    If RFilter = "General" Then
'        For L = 1 To lvwEmp.ListItems.Count
'                LSQL = LSQL & "'" & Trim(Replace(lvwEmp.ListItems.Item(L).Tag, "'", "''")) & "',"
'        Next L
'        'LSQL = Mid$(LSQL, 1, Len(tSQL) - 1)
'        mySQL = mySQL & " AND {Employee.Employee_ID} in [" & LSQL & "]"
'    End If
    
    '==========END CHECK FILTER BY EMPLOYEES ACCESSED=============================
    
    '    ########################################### Server not opened Work Around ###############
      'Loop through all database tables and set the correct server & database
        Dim tbl As CRAXDDRT.DatabaseTable
        Dim tbls As CRAXDDRT.DatabaseTables
    '    ShowReport R
    '    Exit Sub
        Set tbls = r.Database.Tables

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
            conProps.add "User ID", ConParams.UserID
            conProps.add "Password", ConParams.Password
            conProps.add "Persist Security Info", "false"
            tbl.Location = tbl.Name
        Next tbl
    '    R.PrinterSetup 0
    '    ########################################### Server not opened Work Around - END ###############

    '    Dim rsCompany As New ADODB.Recordset
    '    Dim rsCompName As New ADODB.Recordset
    '    Set rsCompany = CConnect.GetRecordSet("select * from STYPES where smain=1")
    '    If rsCompany.RecordCount > 0 Then
    '        Set rsCompName = CConnect.GetRecordSet("select * from GeneralOpt")
    '        If rsCompName.RecordCount > 0 Then
    '            If UCase(Trim(rsCompany!Description & "")) = UCase(Trim(rsCompName!cName & "")) Then
    '                R.ReportTitle = UCase(Trim(rsCompany!Description & ""))
    '            Else
    '                R.ReportTitle = UCase(Trim(rsCompName!cName & "")) & " - " & UCase(Trim(rsCompany!Description & ""))
    '            End If
    '        Else
    '            R.ReportTitle = UCase(Trim(rsCompany!Description & ""))
    '        End If
    '    Else
    '        R.ReportTitle = "TEST COMPANY"
    '    End If

    r.FormulaSyntax = crCrystalSyntaxFormula
    r.RecordSelectionFormula = mySQL
    r.PaperSize = crPaperA4
    
    If r.PaperOrientation = crLandscape Then
        r.BottomMargin = 192
        r.RightMargin = 720
        r.LeftMargin = 58
        r.TopMargin = 192
    ElseIf r.PaperOrientation = crPortrait Then
        r.BottomMargin = 720
        r.RightMargin = 192
        r.LeftMargin = 192
        r.TopMargin = 360
    End If

        
    With frmReports.CRViewer1
        .DisplayGroupTree = False
        .EnableExportButton = True
        .EnableSelectExpertButton = True
        .EnablePrintButton = True
        .DisplayToolbar = True
        .DisplayTabs = True
        .ReportSource = r
        .ViewReport
        
    End With
    
    frmReports.Show vbModal
    Me.MousePointer = 0
    
    ReportType = ""
    RFilter = ""
   
    mySQL = ""
    ''R.RecordSelectionFormula = mySQL
    'frmMain2.fraEmployees.Visible = True
    Set r = Nothing
    Set a = Nothing
    Unload Me
    
    Exit Sub
    
ErrHandler:
    'If Err.Description = "File not found." Then
    '    frmMain2.Cdl.DialogTitle = "Select the report to show"
    '    frmMain2.Cdl.InitDir = App.Path & "/Reports"
    '    frmMain2.Cdl.Filter = "Reports {* .rpt|* .rpt"
    '    frmMain2.Cdl.ShowOpen
    '    myfile = frmMain2.Cdl.FileName
    '    If Not myfile = "" Then
    '        Resume
    '    Else
    '        Me.MousePointer = 0
    '    End If
    'Else
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
    popupText = "RTo"
    frmPopUp.Show vbModal
    gMaxEmployeeID = gEmployeeID
End Sub



Private Sub cmduncheckall_Click()
Dim k As Integer
Dim bSTATE As Boolean

If (lvwEmp.ListItems.count <> 0) Then
        If UCase(cmduncheckall.Caption) = "UNCHECK ALL" Then
                bSTATE = False
                For k = 1 To lvwEmp.ListItems.count
                lvwEmp.ListItems(k).Checked = bSTATE
                Next k
                cmduncheckall.Caption = "Check All"
        Else
                bSTATE = True
                For k = 1 To lvwEmp.ListItems.count
                lvwEmp.ListItems(k).Checked = bSTATE
                Next k
                cmduncheckall.Caption = "Uncheck All"
        End If
End If
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

Private Sub find2()
    Dim Find As Long
    Dim li As ListItem
    Dim Field As String
    Dim t As Integer
    'Set Cnn = New Connection
    Set RsT = New Recordset
    Field = cboField.Text
   ''lvwEmp.ListItems.Clear
    ''lblCount.Caption = 0
    
    
    If Not cboField.Text = "" Then
        If Not cboCrieria.Text = "" Then
            
            If Not cboCrieria.Text = "Between" And Not cboCrieria.Text = "Like" Then
                
    
                Set rst1 = CConnect.GetRecordSet("Select * from Employee where " & cboField.Text & "" & cboCrieria.Text & "'" & txtValue.Text & "'")
    
                
                With rst1
                    If .RecordCount > 0 Then
                        ''lblCount.Caption = .RecordCount
                        .MoveFirst
                        
                        If (lvwEmp.ListItems.count <> 0) Then
                        
                        Do While Not .EOF
                        
                            For t = 1 To lvwEmp.ListItems.count
                              If (lvwEmp.ListItems(t).Text = !EmpCode) Then
                              lvwEmp.ListItems(t).Selected = True
                              lvwEmp.ListItems(t).Checked = True
                              End If
                            Next t
                        .MoveNext
                        Loop
                        
                        Else
                        
                        Do While Not .EOF
                        
                        
                            Set li = lvwEmp.ListItems.add(, , !EmpCode & "")
                            li.ListSubItems.add , , !SurName & ""
                            li.ListSubItems.add , , !OtherNames & ""
                            li.Selected = True
                            li.Checked = True
                            .MoveNext
                        Loop
                        
                        End If
                                                
                        
'
                        
                    Else
                    MsgBox ("No such Record Exists")
                    End If
                End With
                
                Set rst1 = Nothing
                
            ElseIf cboCrieria.Text = "Like" Then
            
                'Set rst1 = cConnect.GetPayData("Select * from Employee order by EmpCode")
                sql = "Select * from Employee where " & cboField.Text & "  like '" & txtValue.Text & "%'  order by EmpCode "
                Set rst1 = CConnect.GetRecordSet(sql)
               
                With rst1
                    If .RecordCount > 0 Then
                        ''lblCount.Caption = .RecordCount
                        .MoveFirst
                        
                        If (lvwEmp.ListItems.count <> 0) Then
                        
                        Do While Not .EOF
                        
                            For t = 1 To lvwEmp.ListItems.count
                              If (lvwEmp.ListItems(t).Text = !EmpCode) Then
                              lvwEmp.ListItems(t).Selected = True
                              lvwEmp.ListItems(t).Checked = True
                              End If
                            Next t
                        .MoveNext
                        Loop
                        
                        Else
                        
                        Do While Not .EOF
                        
                        
                            Set li = lvwEmp.ListItems.add(, , !EmpCode & "")
                            li.ListSubItems.add , , !SurName & ""
                            li.ListSubItems.add , , !OtherNames & ""
                            li.Selected = True
                            li.Checked = True
                            .MoveNext
                        Loop
                        
                        End If
                                                
                        
'
                        
                    Else
                    MsgBox ("No such Record Exists")
                    End If
                End With

                
                
                Set rst1 = Nothing
                
                
            Else
                If cboField.Text = "Amount" Then
    '                Set rst1 = cConnect.GetPayData("select * from Employee where " & cboField.Text & " " & cboCrieria & " " & txtFrom.Text & " And " & txtTo.Text & "")
    '                Set rst1 = CConnect.GetRecordSet("select * from Employee where " & cboField.Text & " " & cboCrieria & " " & txtFrom.Text & " And " & txtTo.Text & "")
                    Set rst1 = CConnect.GetRecordSet("SELECT * FROM Employee as e LEFT JOIN ECategory as ec ON " & _
                            "e.ECategory = ec.code WHERE e.Term <> 1 AND " & cboField.Text & " " & cboCrieria & " " & txtFrom.Text & " And " & txtTo.Text & "' ORDER BY e.EmpCode")
                    
                Else
    '                Set rst1 = cConnect.GetPayData("select * from Employee where " & cboField.Text & " " & cboCrieria & " '" & txtFrom.Text & "' And '" & txtTo.Text & "'")
                    Set rst1 = CConnect.GetRecordSet("SELECT * FROM Employee as e LEFT JOIN ECategory as ec ON " & _
                            "e.ECategory = ec.code WHERE e.Term <> 1 AND " & cboField.Text & " " & cboCrieria & " '" & txtFrom.Text & "' And '" & txtTo.Text & "' ORDER BY e.EmpCode")
                    Set rst1 = CConnect.GetRecordSet("select * from Employee where " & cboField.Text & " " & cboCrieria & " '" & txtFrom.Text & "' And '" & txtTo.Text & "'")
                    
                End If
                
                With rst1
                    If .RecordCount > 0 Then
                        ''lblCount.Caption = .RecordCount
                        .MoveFirst
                        
                        If (lvwEmp.ListItems.count <> 0) Then
                        
                        Do While Not .EOF
                          
                            For t = 1 To lvwEmp.ListItems.count
                              If (lvwEmp.ListItems(t).Text = !EmpCode) Then
                              lvwEmp.ListItems(t).Selected = True
                              lvwEmp.ListItems(t).Checked = True
                              End If
                            Next t
                        .MoveNext
                        Loop
                        
                        Else
                        
                        Do While Not .EOF
                        
                        
                            Set li = lvwEmp.ListItems.add(, , !EmpCode & "")
                            li.ListSubItems.add , , !SurName & ""
                            li.ListSubItems.add , , !OtherNames & ""
                            li.Selected = True
                            li.Checked = True
                            .MoveNext
                        Loop
                        
                        End If
                                                
                        
'
                        
                    Else
                    MsgBox ("No such Record Exists")
                    End If
                End With

                
                
                Set rst1 = Nothing
                
            End If
            
           '' cmdSelect.SetFocus
        Else
            MsgBox "Select the search criteria.", vbExclamation
        End If
    Else
        MsgBox "Select the search field.", vbExclamation
    End If

End Sub


Private Sub load_locations()

End Sub


Private Sub initializeBooleans()
Dim RsT As New ADODB.Recordset
Set RsT = CConnect.GetRecordSet("exec sp_initializeBools")
If Not (RsT.EOF) Then
payrollscount = RsT!pcount
deptscount = RsT!dcount
payrollsexists = True
deptsexists = True
Else
payrollsexists = False
deptsexists = False
payrollscount = 0
deptscount = 0
End If
End Sub
Private Sub Form_Load()
    Dim i As Integer
    On Error GoTo Hell
    
     With cboField
        .AddItem "EmpCode"
        .AddItem "Surname"
        .AddItem "OtherNames"
        .AddItem "IDNo"
    End With
    
    initializeBooleans
    
    lvwLocations.ColumnHeaders.add , , "Location", 2350
    Dim li As ListItem
    Set li = lvwLocations.ListItems.add(, , "Select All")
    li.Tag = "S_A"
    load_locations
    
    lvwFundCodes.ColumnHeaders.add , , "Fund Codes", 2350
   
    Set li = lvwFundCodes.ListItems.add(, , "Select All")
    li.Tag = "S_A"
   '' load_fundcodes
    
    frmMain2.txtDetails.Caption = ""

    CConnect.CColor Me, MyColor

'    Call InitRConnection

    Call InitGrid
    
    Set rs6 = CConnect.GetRecordSet("SELECT * FROM Bonus ORDER BY BonusID")

    With rs6
        If .RecordCount > 0 Then
            .MoveFirst
            cboBonus.AddItem ""
            Do While Not .EOF
                cboBonus.AddItem Format(!Perc & "", "##0.00") & "%    " & !BonusID & "" & " - " & !Comments & "" & " Months"
                           
                .MoveNext
            Loop
        End If
    End With

    Set rs6 = Nothing
    canAccessEmployees = False
    Call LoadList
    Call LoadCbo
    Call myStructure
    
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
    
Exit Sub
Hell:
MsgBox err.Description, vbExclamation

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

Private Sub lvwEmp_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Item.Selected = True
End Sub

Private Sub lvwpaypoint_ItemCheck(ByVal Item As MSComctlLib.ListItem)
  On Error GoTo ErrorHandler
    Dim n As Integer, State As Boolean
    
    State = False
    If lvwpaypoint.ListItems.Item(1).Checked = True Then State = True
    
    If Item.Tag = "S-A" Then
        'Uncheck All Departments
        For n = 2 To lvwpaypoint.ListItems.count
            lvwpaypoint.ListItems.Item(n).Checked = State
        Next n
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error Selecting Paypoints:" & vbNewLine & err.Description, vbExclamation, "PDR Error"

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


Private Sub optalldates_Click()
If (optalldates.value = True) Then
dtfrom.Visible = False
dtto.Visible = False
Else
dtfrom.Visible = True
dtto.Visible = True
End If
End Sub

Private Sub optbetween_Click()
If (optbetween.value = True) Then
dtfrom.Visible = True
dtto.Visible = True
Else
dtfrom.Visible = False
dtto.Visible = False
End If
End Sub

Private Sub optlist_Click()
If (optlist.value = True) Then
txtFromE.Enabled = False
cmdFromE.Enabled = False
txtToE.Enabled = False
cmdToE.Enabled = False
lvwEmp.Enabled = True

cboField.Enabled = True
cboCrieria.Enabled = True
txtValue.Enabled = True
cmdfind2.Enabled = True
Else
lvwEmp.Enabled = False
txtFromE.Enabled = True
cmdFromE.Enabled = True
txtToE.Enabled = True
cmdToE.Enabled = True

cboField.Enabled = False
cboCrieria.Enabled = False
txtValue.Enabled = False
cmdfind2.Enabled = False
End If
End Sub

Private Sub optrange_Click()
If (optrange.value = True) Then
txtFromE.Enabled = True
cmdFromE.Enabled = True
txtToE.Enabled = True
cmdToE.Enabled = True

cboField.Enabled = False
cboCrieria.Enabled = False
txtValue.Enabled = False
cmdfind2.Enabled = False
lvwEmp.Enabled = False
Else
lvwEmp.Enabled = True
txtFromE.Enabled = False
cmdFromE.Enabled = False
txtToE.Enabled = False
cmdToE.Enabled = False

cboField.Enabled = True
cboCrieria.Enabled = True
txtValue.Enabled = True
cmdfind2.Enabled = True
End If
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
    Dim tmprs As New ADODB.Recordset
    Dim li As ListItem
    With lvwEmp
        .ColumnHeaders.Clear
        .ColumnHeaders.add , , "Code", 1300
        .ColumnHeaders.add , , "Name", 2800
        
        .View = lvwReport
    End With
    
    With lsvCat
        .ColumnHeaders.Clear
        .ColumnHeaders.add , , "Code", 1000
        .ColumnHeaders.add , , "Description", .Width - 1000
        
        .View = lvwReport
        
        Set tmprs = CConnect.GetRecordSet("SELECT * FROM ECategory ORDER BY seq ASC")
        While tmprs.EOF = False
            Set li = .ListItems.add(, , tmprs!Code & "")
            li.ListSubItems.add , , tmprs!Comments & ""
            tmprs.MoveNext
        Wend
    End With
End Sub

Public Sub LoadList()
    Dim i As Long
    
'    With rsGlob2
'        If .RecordCount > 0 Then
'            .MoveFirst
'            txtFromE.Text = !empcode & ""
'            While Not .EOF
'                Set LI = lvwEmp.ListItems.Add(, , !empcode & "")
'                LI.ListSubItems.Add , , !SurName & " " & !OtherNames & ""
'                LI.Tag = !employee_id
'                .MoveNext
'            Wend
'            .MoveLast
'            txtToE.Text = !empcode & ""
'            .MoveFirst
'            lblEmpCount.Caption = .RecordCount
'        End If
'    End With

''  MODIFIED BY KEVIN ON 12-02-2009
   On error2 GoTo err
    AllEmployees.GetAccessibleEmployeesByUser currUser.UserID
    'Set rst1 = cConnect.GetPayData("Select * from Employee order by EmpCode")
     If (AllEmployees.count = 0) Then
     MsgBox ("User Cannot Access Any Employee")
     canAccessEmployees = False
     Exit Sub
     Else
     canAccessEmployees = True
     End If
   
      txtFromE.Text = AllEmployees.Item(1).EmpCode
      gMinEmployeeID = AllEmployees.Item(1).EmployeeID
    For i = 1 To AllEmployees.count
        If Not (AllEmployees.Item(i).IsDisengaged) Then 'IF CLAUSE ADDED BY JOHN TO TRAP OBJECT PRE-LOADS
            Set li = Me.lvwEmp.ListItems.add(, , AllEmployees.Item(i).EmpCode)
            li.SubItems(1) = AllEmployees.Item(i).SurName & "  " & AllEmployees.Item(i).OtherNames
            li.Tag = AllEmployees.Item(i).EmployeeID
            li.Checked = True
        End If
    Next i
    lblEmpCount.Caption = AllEmployees.count
   txtToE.Text = AllEmployees.Item(AllEmployees.count).EmpCode
   gMaxEmployeeID = AllEmployees.Item(AllEmployees.count).EmployeeID
   Exit Sub
err:
   MsgBox (err.Description)
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
                Set li = lvwTerms.ListItems.add(, , "Un defined")
                li.Tag = 0
                li.Checked = True
        End If
    End With
    Set rs3 = Nothing
    
    'get payroll types
    
   '' Set rs3 = CConnect.GetRecordSet("SELECT * FROM tblpayroll ORDER BY payroll_id ")
    Set rs3 = CConnect.GetRecordSet("exec get_AccessiblePayrollsByUseID " & currUser.UserID)
    lvwpayrolltype.ListItems.Clear

   
    With rs3
        If .RecordCount > 0 Then
        
        
        
        
        
            .MoveFirst
        
            
            
            
            Do While Not .EOF
            
            ''**********************
'             If Not (gAccessRightClassIds = "") Then '' user is not infiniti
'                If UCase(gAccessRightName) = "PAYROLL TYPES" Then
            ''*********************
                Set li = lvwpayrolltype.ListItems.add(, , !Payroll_name)
                li.Tag = !payroll_id
                li.Checked = True
                .MoveNext
            Loop
                Set li = lvwpayrolltype.ListItems.add(, , "Un Defined")
                li.Tag = 0
                li.Checked = True
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
    ItemD.Tag = "S-A"
    lvwpaypoint.ListItems.Item(1).Checked = True
    ItemD.Bold = True
    
    Set ItemD = lvwpaypoint.ListItems.add(, , "Un Defined")
    ItemD.Tag = 0
    lvwpaypoint.ListItems.Item(1).Checked = True
    ItemD.Bold = True
        
    Exit Sub
ErrorHandler:
    MsgBox "Error Loading Paypoints" & vbNewLine & err.Description, vbExclamation, "PDR Error"
End Sub


Private Sub PopulateDpts()
    On Error GoTo ErrorHandler
    Dim rsOUs As New ADODB.Recordset, ItemD As ListItem
    
    ''Set rsOUs = CConnect.GetRecordSet("Select * From OrganizationUnits Order by OrganizationUnitName")
    Set rsOUs = CConnect.GetRecordSet("exec sp_getAccessibleOUNITS " & currUser.UserID & "")
    If rsOUs.RecordCount > 0 Then
'        Do Until rsOUs.EOF
'            With LvwDpts
'                Set ItemD = .ListItems.add(, , rsOUs!OrganizationUnitName)
'                ItemD.Tag = rsOUs!OrganizationUnitID
                
                ''************************
                            Do Until rsOUs.EOF
                               
                                With LvwDpts
                                    Set ItemD = .ListItems.add(, , rsOUs!OrganizationUnitName)
                                    ItemD.Tag = rsOUs!OrganizationUnitID
                                    ItemD.Checked = True
                                    rsOUs.MoveNext
                                End With
                            Loop
                ''*********************
                
                
'                If Not (gAccessRightName = "") Then
'                    If (UCase(Trim(gAccessRightName)) = "ORGANIZATION UNITS") Then
'                        i = UBound(gAccessRighClassIdsArray) - LBound(gAccessRighClassIdsArray)
'                        If (i > 0) Then
'
'                        Dim i2 As Integer
'                        Dim Found As Boolean
'
'                        ''***************************
'                            Do Until rsOUs.EOF
'                                    With LvwDpts
''                                    Set ItemD = .ListItems.add(, , rsOUs!OrganizationUnitName)
''                                    ItemD.Tag = rsOUs!OrganizationUnitID
'
'
'                                    i2 = 0
'                                    Found = False
'                                    For i2 = 0 To i - 1
'                                    Dim id As Integer
'                                        id = gAccessRighClassIdsArray(i2)
'                                        If gAccessRighClassIdsArray(i2) = rsOUs!OrganizationUnitID Then
'                                        Found = True
'                                        Exit For
'                                        End If
'                                    Next i2
'
'
'                                        If Found = True Then
'                                            Set ItemD = .ListItems.add(, , rsOUs!OrganizationUnitName)
'                                            ItemD.Tag = rsOUs!OrganizationUnitID
'                                            ItemD.Checked = True
'
'                                        End If
'                                     End With
'                             rsOUs.MoveNext
'                             Loop
'                       ''******************
'                        End If
'                    Else
'                            Do Until rsOUs.EOF
'                                With LvwDpts
'                                    Set ItemD = .ListItems.add(, , rsOUs!OrganizationUnitName)
'                                    ItemD.Tag = rsOUs!OrganizationUnitID
'                                    ItemD.Checked = True
'                                    rsOUs.MoveNext
'                                End With
'                            Loop
'                    End If
'                Else
'                            Do Until rsOUs.EOF
'                                With LvwDpts
'                                    Set ItemD = .ListItems.add(, , rsOUs!OrganizationUnitName)
'                                    ItemD.Tag = rsOUs!OrganizationUnitID
'                                    ItemD.Checked = True
'                                    rsOUs.MoveNext
'                                End With
'                            Loop
'                End If
''                ItemD.Checked = True
''            End With
''            rsOUs.MoveNext
''        Loop


    Set ItemD = LvwDpts.ListItems.add(1, , "(Select All)")
    ItemD.Tag = "S-A"
    LvwDpts.ListItems.Item(1).Checked = True
    ItemD.Bold = True
    
        Set ItemD = LvwDpts.ListItems.add(, , "Un Defined")
        ItemD.Tag = 0
        ItemD.Checked = True
 
    
    
    End If
    
'    Set ItemD = LvwDpts.ListItems.add(1, , "(Select All)")
'    ItemD.Tag = "S-A"
'    LvwDpts.ListItems.Item(1).Checked = True
'    ItemD.Bold = True
        
    Exit Sub
ErrorHandler:
    MsgBox "Error Loading Departments" & vbNewLine & err.Description, vbExclamation, "PDR Error"
End Sub
