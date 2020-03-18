VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEmployeeOld 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   Caption         =   "Employee Details"
   ClientHeight    =   7335
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraProgress 
      Caption         =   "Updating records, please wait ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   540
      TabIndex        =   64
      Top             =   3330
      Visible         =   0   'False
      Width           =   5535
      Begin MSComctlLib.ProgressBar prgSave 
         Height          =   195
         Left            =   120
         TabIndex        =   65
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   344
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame FraEdit 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7350
      Left            =   0
      TabIndex        =   55
      Top             =   -75
      Width           =   7440
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   120
         TabIndex        =   117
         Top             =   6750
         Width           =   7215
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   5.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6180
            Picture         =   "frmEmployeeOld.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Save Record"
            Top             =   120
            Width           =   495
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   5.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6670
            Picture         =   "frmEmployeeOld.frx":0102
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Cancel Process"
            Top             =   120
            Width           =   495
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "EMPLOYEE'S GENERAL DETAILS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1840
         Left            =   120
         TabIndex        =   100
         Top             =   120
         Width           =   7215
         Begin VB.TextBox txtAlien 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3930
            TabIndex        =   6
            Top             =   847
            Width           =   1575
         End
         Begin VB.TextBox txtPassport 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1125
            TabIndex        =   2
            Top             =   850
            Width           =   1635
         End
         Begin VB.TextBox txtONames 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1125
            TabIndex        =   1
            Top             =   530
            Width           =   1635
         End
         Begin VB.TextBox txtSurname 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3930
            TabIndex        =   4
            Top             =   210
            Width           =   1575
         End
         Begin VB.TextBox txtEmpCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1125
            TabIndex        =   0
            Top             =   210
            Width           =   1635
         End
         Begin VB.TextBox txtIDNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3930
            TabIndex        =   5
            Top             =   525
            Width           =   1575
         End
         Begin VB.ComboBox cboGender 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            ItemData        =   "frmEmployeeOld.frx":0204
            Left            =   1125
            List            =   "frmEmployeeOld.frx":0211
            TabIndex        =   3
            Text            =   "Unspecified"
            Top             =   1170
            Width           =   1635
         End
         Begin MSComCtl2.DTPicker dtpDOB 
            Height          =   285
            Left            =   3930
            TabIndex        =   7
            Top             =   1170
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
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
            OLEDropMode     =   1
            CustomFormat    =   "dd, MMM, yyyy"
            Format          =   20709379
            CurrentDate     =   37845
            MinDate         =   -36522
         End
         Begin MSComctlLib.ImageList imgEmpTool 
            Left            =   6000
            Top             =   360
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   4
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEmployeeOld.frx":0230
                  Key             =   "Search"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEmployeeOld.frx":0342
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEmployeeOld.frx":0454
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEmployeeOld.frx":0566
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton cmdPNew 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5640
            Picture         =   "frmEmployeeOld.frx":0AA8
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   1080
            Visible         =   0   'False
            Width           =   320
         End
         Begin VB.CommandButton cmdPDelete 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6000
            Picture         =   "frmEmployeeOld.frx":0BAA
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   1080
            Visible         =   0   'False
            Width           =   320
         End
         Begin VB.Image Picture1 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   1335
            Left            =   5670
            Stretch         =   -1  'True
            Top             =   120
            Width           =   1425
         End
         Begin VB.Image imgDeletePic 
            Appearance      =   0  'Flat
            Height          =   180
            Left            =   6900
            MouseIcon       =   "frmEmployeeOld.frx":109C
            MousePointer    =   99  'Custom
            Picture         =   "frmEmployeeOld.frx":14DE
            Stretch         =   -1  'True
            ToolTipText     =   "Click this icon to DELETE employee photo"
            Top             =   1530
            Width           =   195
         End
         Begin VB.Image imgLoadPic 
            Appearance      =   0  'Flat
            Height          =   180
            Left            =   5670
            MouseIcon       =   "frmEmployeeOld.frx":1920
            MousePointer    =   99  'Custom
            Picture         =   "frmEmployeeOld.frx":1D62
            Stretch         =   -1  'True
            ToolTipText     =   "Click this icon to ADD a new photo"
            Top             =   1530
            Width           =   195
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "*ID No.:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2820
            TabIndex        =   108
            Top             =   555
            Width           =   615
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "*Date Of Birth:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2820
            TabIndex        =   107
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "*Gender:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   106
            Top             =   1185
            Width           =   675
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "*Staff No.:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   105
            Top             =   255
            Width           =   810
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "*Surname:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2820
            TabIndex        =   104
            Top             =   240
            Width           =   780
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "*Other Names:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   103
            Top             =   570
            Width           =   1095
         End
         Begin VB.Label Label43 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "*Passport No.:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   102
            Top             =   870
            Width           =   1080
         End
         Begin VB.Label Label44 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "*Alien Card No.:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2820
            TabIndex        =   101
            Top             =   885
            Width           =   1185
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   2835
         Left            =   120
         TabIndex        =   66
         Top             =   3960
         Width           =   7215
         Begin TabDlg.SSTab SSTab1 
            Height          =   2745
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   4842
            _Version        =   393216
            TabHeight       =   441
            BackColor       =   -2147483648
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "EMPLOYMENT DETAILS"
            TabPicture(0)   =   "frmEmployeeOld.frx":1EAC
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Frame4"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "EMPLOYEE PAYMENT DETAILS"
            TabPicture(1)   =   "frmEmployeeOld.frx":1EC8
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "fraSalary"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "DISENGAGEMENT DETAILS"
            TabPicture(2)   =   "frmEmployeeOld.frx":1EE4
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Frame6"
            Tab(2).ControlCount=   1
            Begin VB.Frame Frame6 
               Height          =   2175
               Left            =   -74880
               TabIndex        =   94
               Top             =   360
               Width           =   6975
               Begin VB.Frame Frame12 
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1935
                  Left            =   120
                  TabIndex        =   98
                  Top             =   120
                  Width           =   3015
                  Begin VB.CheckBox chkTerm 
                     Appearance      =   0  'Flat
                     Caption         =   "Disengaged"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00000000&
                     Height          =   255
                     Left            =   75
                     TabIndex        =   41
                     Top             =   720
                     Width           =   1185
                  End
                  Begin VB.CheckBox chkReEngage 
                     Appearance      =   0  'Flat
                     Caption         =   "Cannot be re-engaged"
                     ForeColor       =   &H80000008&
                     Height          =   255
                     Left            =   75
                     TabIndex        =   43
                     Top             =   1560
                     Visible         =   0   'False
                     Width           =   1935
                  End
                  Begin VB.ComboBox cboTermReasons 
                     Appearance      =   0  'Flat
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00000000&
                     Height          =   315
                     ItemData        =   "frmEmployeeOld.frx":1F00
                     Left            =   75
                     List            =   "frmEmployeeOld.frx":1F1C
                     TabIndex        =   42
                     Top             =   1080
                     Width           =   2820
                  End
                  Begin MSComCtl2.DTPicker dtpTerm 
                     Height          =   330
                     Left            =   1680
                     TabIndex        =   40
                     Top             =   360
                     Width           =   1215
                     _ExtentX        =   2143
                     _ExtentY        =   582
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
                     CustomFormat    =   "yyyy-MM-dd"
                     Format          =   20709379
                     CurrentDate     =   37845
                  End
                  Begin VB.Label Label9 
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     Caption         =   "Disengagement Date:"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   195
                     Left            =   120
                     TabIndex        =   99
                     Top             =   360
                     Width           =   1560
                  End
               End
               Begin VB.Frame fraTerm 
                  Caption         =   "RETIREMENT DETAILS"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1935
                  Left            =   3360
                  TabIndex        =   95
                  Top             =   120
                  Width           =   3495
                  Begin VB.CheckBox chkTermTrain 
                     Caption         =   "Retirement Training"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   270
                     Left            =   135
                     TabIndex        =   44
                     Top             =   330
                     Width           =   1950
                  End
                  Begin VB.TextBox txtAdvisor 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFFFF&
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00000000&
                     Height          =   285
                     Left            =   1200
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   46
                     Top             =   1170
                     Width           =   1920
                  End
                  Begin VB.CheckBox chkAchieved 
                     Caption         =   "Retirement Training Achieved"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   300
                     Left            =   135
                     TabIndex        =   47
                     Top             =   1560
                     Width           =   2490
                  End
                  Begin MSComCtl2.DTPicker dtpTerminalDate 
                     Height          =   330
                     Left            =   1200
                     TabIndex        =   45
                     Top             =   780
                     Width           =   1920
                     _ExtentX        =   3387
                     _ExtentY        =   582
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
                     CustomFormat    =   "yyyy-MM-dd"
                     Format          =   20709379
                     CurrentDate     =   37845
                  End
                  Begin VB.Label Label29 
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     Caption         =   "Training Date:"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   195
                     Left            =   120
                     TabIndex        =   97
                     Top             =   840
                     Width           =   1020
                  End
                  Begin VB.Label Label8 
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     Caption         =   "Advisor"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   195
                     Left            =   135
                     TabIndex        =   96
                     Top             =   1170
                     Width           =   540
                  End
               End
            End
            Begin VB.Frame fraSalary 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   2085
               Left            =   -74880
               TabIndex        =   78
               Top             =   360
               Width           =   6855
               Begin VB.Frame fraStateNumbers 
                  BorderStyle     =   0  'None
                  Height          =   1935
                  Left            =   120
                  TabIndex        =   79
                  Top             =   120
                  Width           =   6615
                  Begin VB.TextBox txtHAllow 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFFFF&
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00000000&
                     Height          =   285
                     Left            =   1095
                     TabIndex        =   37
                     Top             =   1320
                     Width           =   1410
                  End
                  Begin VB.TextBox txtBasicPay 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFFFF&
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00000000&
                     Height          =   285
                     Left            =   3885
                     TabIndex        =   50
                     Top             =   960
                     Width           =   1410
                  End
                  Begin VB.TextBox txtCert 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFFFF&
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00000000&
                     Height          =   285
                     Left            =   1095
                     TabIndex        =   36
                     Top             =   950
                     Width           =   1410
                  End
                  Begin VB.Frame fraSal 
                     BorderStyle     =   0  'None
                     Height          =   495
                     Left            =   5880
                     TabIndex        =   85
                     Top             =   1560
                     Width           =   5055
                     Begin VB.TextBox txtLAllow 
                        Alignment       =   1  'Right Justify
                        Appearance      =   0  'Flat
                        BackColor       =   &H00FFFFFF&
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00000000&
                        Height          =   285
                        Left            =   9690
                        TabIndex        =   89
                        Top             =   210
                        Visible         =   0   'False
                        Width           =   1110
                     End
                     Begin VB.TextBox txtOAllow 
                        Alignment       =   1  'Right Justify
                        Appearance      =   0  'Flat
                        BackColor       =   &H00FFFFFF&
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00000000&
                        Height          =   285
                        Left            =   1935
                        TabIndex        =   88
                        Top             =   90
                        Visible         =   0   'False
                        Width           =   765
                     End
                     Begin VB.TextBox txtGrossPay 
                        Alignment       =   1  'Right Justify
                        Appearance      =   0  'Flat
                        BackColor       =   &H00FFFFFF&
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00000000&
                        Height          =   285
                        Left            =   8445
                        TabIndex        =   87
                        Top             =   210
                        Visible         =   0   'False
                        Width           =   1155
                     End
                     Begin VB.TextBox txtTAllow 
                        Alignment       =   1  'Right Justify
                        Appearance      =   0  'Flat
                        BackColor       =   &H00FFFFFF&
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00000000&
                        Height          =   285
                        Left            =   9615
                        TabIndex        =   86
                        Top             =   210
                        Visible         =   0   'False
                        Width           =   1140
                     End
                     Begin VB.Label Label23 
                        Appearance      =   0  'Flat
                        AutoSize        =   -1  'True
                        Caption         =   "Leave Allow."
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H80000008&
                        Height          =   195
                        Left            =   9690
                        TabIndex        =   93
                        Top             =   0
                        Visible         =   0   'False
                        Width           =   915
                     End
                     Begin VB.Label Label25 
                        Appearance      =   0  'Flat
                        AutoSize        =   -1  'True
                        Caption         =   "Other Allow."
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H80000008&
                        Height          =   195
                        Left            =   1935
                        TabIndex        =   92
                        Top             =   120
                        Visible         =   0   'False
                        Width           =   780
                     End
                     Begin VB.Label Label26 
                        Appearance      =   0  'Flat
                        AutoSize        =   -1  'True
                        Caption         =   "Gross Pay"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H80000008&
                        Height          =   195
                        Left            =   8445
                        TabIndex        =   91
                        Top             =   0
                        Visible         =   0   'False
                        Width           =   720
                     End
                     Begin VB.Label Label27 
                        Appearance      =   0  'Flat
                        AutoSize        =   -1  'True
                        Caption         =   "Transport Allow."
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H80000008&
                        Height          =   195
                        Left            =   8295
                        TabIndex        =   90
                        Top             =   0
                        Visible         =   0   'False
                        Width           =   1185
                     End
                  End
                  Begin VB.TextBox txtKRAFileNO 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFFFF&
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00000000&
                     Height          =   285
                     Left            =   1095
                     TabIndex        =   34
                     Top             =   210
                     Width           =   1410
                  End
                  Begin VB.TextBox txtPin 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFFFF&
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00000000&
                     Height          =   285
                     Left            =   1095
                     TabIndex        =   35
                     Top             =   580
                     Width           =   1410
                  End
                  Begin VB.TextBox txtNssf 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFFFF&
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00000000&
                     Height          =   285
                     Left            =   3885
                     TabIndex        =   38
                     Top             =   210
                     Width           =   1410
                  End
                  Begin VB.TextBox txtNhif 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFFFF&
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00000000&
                     Height          =   285
                     Left            =   3885
                     TabIndex        =   39
                     Top             =   570
                     Width           =   1410
                  End
                  Begin VB.Label Label28 
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     Caption         =   "Basic Pay:"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   195
                     Left            =   3105
                     TabIndex        =   126
                     Top             =   990
                     Width           =   735
                  End
                  Begin VB.Label Label24 
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     Caption         =   "House Allowance:"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   195
                     Left            =   0
                     TabIndex        =   125
                     Top             =   1350
                     Width           =   1275
                  End
                  Begin VB.Label Label39 
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     Caption         =   "KRA File No.:"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   195
                     Left            =   0
                     TabIndex        =   84
                     Top             =   240
                     Width           =   945
                  End
                  Begin VB.Label Label17 
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     Caption         =   "PIN No:"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   195
                     Left            =   0
                     TabIndex        =   83
                     Top             =   600
                     Width           =   555
                  End
                  Begin VB.Label Label16 
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     Caption         =   "NSSF No.:"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   195
                     Left            =   3105
                     TabIndex        =   82
                     Top             =   240
                     Width           =   735
                  End
                  Begin VB.Label Label15 
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     Caption         =   "NHIF No.:"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   195
                     Left            =   3105
                     TabIndex        =   81
                     Top             =   600
                     Width           =   720
                  End
                  Begin VB.Label Label32 
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     Caption         =   "Good Conduct No.:"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   195
                     Left            =   0
                     TabIndex        =   80
                     Top             =   960
                     Width           =   1380
                  End
               End
            End
            Begin VB.Frame Frame4 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2340
               Left            =   120
               TabIndex        =   67
               Top             =   270
               Width           =   6975
               Begin VB.TextBox txtDisabilityDet 
                  Height          =   285
                  Left            =   4860
                  TabIndex        =   135
                  Top             =   1980
                  Width           =   1905
               End
               Begin VB.CheckBox chkOnProbation 
                  Caption         =   "On Probation"
                  Height          =   195
                  Left            =   4950
                  TabIndex        =   133
                  Top             =   540
                  Width           =   1275
               End
               Begin VB.TextBox txtProb 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   315
                  Left            =   4965
                  TabIndex        =   25
                  Top             =   840
                  Width           =   270
               End
               Begin VB.ComboBox cboProbType 
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   315
                  ItemData        =   "frmEmployeeOld.frx":1F76
                  Left            =   5250
                  List            =   "frmEmployeeOld.frx":1F83
                  TabIndex        =   26
                  Top             =   840
                  Width           =   1560
               End
               Begin VB.ComboBox txtDesig 
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   315
                  Left            =   1350
                  TabIndex        =   24
                  Top             =   522
                  Width           =   1845
               End
               Begin VB.CheckBox chkDisabled 
                  Appearance      =   0  'Flat
                  Caption         =   "Physically Challenged"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Left            =   1350
                  TabIndex        =   29
                  Top             =   2070
                  Width           =   1845
               End
               Begin VB.ComboBox cboTerms 
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   285
                  ItemData        =   "frmEmployeeOld.frx":1FA1
                  Left            =   1350
                  List            =   "frmEmployeeOld.frx":1FB7
                  TabIndex        =   31
                  Top             =   924
                  Width           =   1890
               End
               Begin VB.ComboBox cboType 
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   285
                  ItemData        =   "frmEmployeeOld.frx":1FFB
                  Left            =   1350
                  List            =   "frmEmployeeOld.frx":2005
                  TabIndex        =   28
                  Top             =   1296
                  Width           =   1890
               End
               Begin VB.ComboBox cboCat 
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   315
                  ItemData        =   "frmEmployeeOld.frx":2021
                  Left            =   1350
                  List            =   "frmEmployeeOld.frx":2023
                  TabIndex        =   32
                  Top             =   1668
                  Width           =   1890
               End
               Begin MSComCtl2.DTPicker dtpDEmployed 
                  Height          =   315
                  Left            =   1365
                  TabIndex        =   23
                  Top             =   120
                  Width           =   1845
                  _ExtentX        =   3254
                  _ExtentY        =   556
                  _Version        =   393216
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  CustomFormat    =   "dd, MMM, yyyy"
                  Format          =   20709379
                  CurrentDate     =   37845
               End
               Begin MSComCtl2.DTPicker dtpValidThrough 
                  Height          =   315
                  Left            =   4950
                  TabIndex        =   27
                  Top             =   135
                  Width           =   1845
                  _ExtentX        =   3254
                  _ExtentY        =   556
                  _Version        =   393216
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  CustomFormat    =   "dd, MMM, yyyy"
                  Format          =   20709379
                  CurrentDate     =   37845
               End
               Begin MSComCtl2.DTPicker dtpPSDate 
                  Height          =   285
                  Left            =   4965
                  TabIndex        =   30
                  Top             =   1200
                  Visible         =   0   'False
                  Width           =   1860
                  _ExtentX        =   3281
                  _ExtentY        =   503
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
                  CustomFormat    =   "dd, MMM, yyyy"
                  Format          =   20709379
                  CurrentDate     =   37845
               End
               Begin MSComCtl2.DTPicker dtpCDate 
                  Height          =   300
                  Left            =   4950
                  TabIndex        =   33
                  Top             =   1560
                  Visible         =   0   'False
                  Width           =   1860
                  _ExtentX        =   3281
                  _ExtentY        =   529
                  _Version        =   393216
                  Enabled         =   0   'False
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  CustomFormat    =   "dd, MMM, yyyy"
                  Format          =   20709379
                  CurrentDate     =   37845
               End
               Begin VB.Label Label13 
                  Caption         =   "Disability Details:"
                  Height          =   195
                  Left            =   3600
                  TabIndex        =   134
                  Top             =   2025
                  Width           =   1275
               End
               Begin VB.Label lblCDate 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Confirmation Date:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Left            =   3600
                  TabIndex        =   77
                  Top             =   1620
                  Width           =   1365
               End
               Begin VB.Label Label35 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Prob. Period (M):"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Left            =   3600
                  TabIndex        =   76
                  Top             =   900
                  Width           =   1230
               End
               Begin VB.Label lblSDate 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Start Date:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Left            =   3600
                  TabIndex        =   75
                  Top             =   1245
                  Width           =   810
               End
               Begin VB.Label Label12 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Designation:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Left            =   120
                  TabIndex        =   70
                  Top             =   570
                  Width           =   900
               End
               Begin VB.Label Label18 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "*Emp. Terms:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Left            =   120
                  TabIndex        =   73
                  Top             =   975
                  Width           =   990
               End
               Begin VB.Label Label19 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "*Emp. Type:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Left            =   90
                  TabIndex        =   72
                  Top             =   1305
                  Width           =   915
               End
               Begin VB.Label Label10 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "*Emp. Grade:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Left            =   90
                  TabIndex        =   71
                  Top             =   1680
                  Width           =   990
               End
               Begin VB.Label Label6 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "*Date Employed:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Left            =   120
                  TabIndex        =   69
                  Top             =   180
                  Width           =   1230
               End
               Begin VB.Label Label41 
                  Appearance      =   0  'Flat
                  Caption         =   "Emp. Valid Thro:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Left            =   3600
                  TabIndex        =   68
                  Top             =   195
                  Width           =   1290
               End
            End
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   120
         TabIndex        =   74
         Top             =   1920
         Width           =   7215
         Begin TabDlg.SSTab SSTab2 
            Height          =   1875
            Left            =   0
            TabIndex        =   8
            Top             =   120
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   3307
            _Version        =   393216
            TabHeight       =   441
            BackColor       =   -2147483648
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "OTHER DETAILS"
            TabPicture(0)   =   "frmEmployeeOld.frx":2025
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Frame8"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "BANK DETAILS"
            TabPicture(1)   =   "frmEmployeeOld.frx":2041
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame9"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "DEPARTMENT INFO"
            TabPicture(2)   =   "frmEmployeeOld.frx":205D
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Label45"
            Tab(2).Control(1)=   "cboOU"
            Tab(2).Control(2)=   "txtOUInfo"
            Tab(2).Control(3)=   "chkHasOUV"
            Tab(2).Control(4)=   "fraEOUV"
            Tab(2).ControlCount=   5
            Begin VB.Frame fraEOUV 
               Caption         =   "Select the Organization Units"
               Enabled         =   0   'False
               Height          =   1215
               Left            =   -72000
               TabIndex        =   131
               Top             =   600
               Width           =   4095
               Begin MSComctlLib.ListView lvwOU 
                  Height          =   855
                  Left            =   120
                  TabIndex        =   132
                  Top             =   240
                  Width           =   3855
                  _ExtentX        =   6800
                  _ExtentY        =   1508
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   0   'False
                  Checkboxes      =   -1  'True
                  FullRowSelect   =   -1  'True
                  GridLines       =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  Appearance      =   1
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "OU Name"
                     Object.Width           =   3528
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Text            =   "Family Tree"
                     Object.Width           =   5292
                  EndProperty
               End
            End
            Begin VB.CheckBox chkHasOUV 
               Caption         =   "Employee is Visible in Other Organization Units"
               Height          =   195
               Left            =   -72000
               TabIndex        =   130
               ToolTipText     =   "Similar Organization Units MUST have same CODE and NAME"
               Top             =   360
               Width           =   3975
            End
            Begin VB.TextBox txtOUInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000018&
               Height          =   735
               Left            =   -74880
               MultiLine       =   -1  'True
               ScrollBars      =   3  'Both
               TabIndex        =   128
               Top             =   960
               Width           =   2775
            End
            Begin VB.ComboBox cboOU 
               Height          =   315
               Left            =   -74880
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   127
               Top             =   600
               Width           =   2535
            End
            Begin VB.Frame Frame9 
               Height          =   1400
               Left            =   -74880
               TabIndex        =   118
               Top             =   300
               Width           =   6975
               Begin VB.TextBox txtBankBranchName 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   1320
                  Locked          =   -1  'True
                  TabIndex        =   19
                  Top             =   650
                  Width           =   2445
               End
               Begin VB.TextBox txtBankName 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   1320
                  Locked          =   -1  'True
                  TabIndex        =   17
                  Top             =   240
                  Width           =   2445
               End
               Begin VB.TextBox txtAccountNO 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   21
                  Top             =   1020
                  Width           =   2445
               End
               Begin VB.CommandButton cmdSchBankBranch 
                  Height          =   285
                  Left            =   3840
                  Picture         =   "frmEmployeeOld.frx":2079
                  Style           =   1  'Graphical
                  TabIndex        =   20
                  Top             =   650
                  Width           =   315
               End
               Begin VB.CommandButton cmdSchBank 
                  Height          =   285
                  Left            =   3840
                  Picture         =   "frmEmployeeOld.frx":2403
                  Style           =   1  'Graphical
                  TabIndex        =   18
                  Top             =   240
                  Width           =   315
               End
               Begin VB.TextBox txtBankCode 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   1800
                  TabIndex        =   119
                  Text            =   "BankCode"
                  Top             =   240
                  Width           =   795
               End
               Begin VB.TextBox txtBankBranch 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   1800
                  TabIndex        =   120
                  Text            =   "BankBranch"
                  Top             =   650
                  Width           =   915
               End
               Begin VB.Label Label38 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Account No.:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Left            =   120
                  TabIndex        =   123
                  Top             =   1020
                  Width           =   945
               End
               Begin VB.Label Label37 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Bank Branch Name:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Left            =   120
                  TabIndex        =   122
                  Top             =   675
                  Width           =   1395
               End
               Begin VB.Label Label36 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Bank Name:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Left            =   120
                  TabIndex        =   121
                  Top             =   240
                  Width           =   855
               End
            End
            Begin VB.Frame Frame8 
               Height          =   1540
               Left            =   120
               TabIndex        =   109
               Top             =   260
               Width           =   6975
               Begin VB.ComboBox cboMarritalStat 
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   315
                  ItemData        =   "frmEmployeeOld.frx":278D
                  Left            =   1080
                  List            =   "frmEmployeeOld.frx":27A0
                  Sorted          =   -1  'True
                  TabIndex        =   12
                  Top             =   1080
                  Width           =   1860
               End
               Begin VB.ComboBox cboNationality 
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   315
                  ItemData        =   "frmEmployeeOld.frx":27D5
                  Left            =   1080
                  List            =   "frmEmployeeOld.frx":27D7
                  Sorted          =   -1  'True
                  TabIndex        =   9
                  Top             =   180
                  Width           =   1860
               End
               Begin VB.TextBox txtEmail 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   1080
                  TabIndex        =   11
                  Top             =   840
                  Width           =   1860
               End
               Begin VB.TextBox txtTel 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   1080
                  TabIndex        =   10
                  Top             =   515
                  Width           =   1860
               End
               Begin VB.ComboBox CboReligion 
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   285
                  ItemData        =   "frmEmployeeOld.frx":27D9
                  Left            =   4395
                  List            =   "frmEmployeeOld.frx":27DB
                  TabIndex        =   15
                  Top             =   850
                  Width           =   2490
               End
               Begin VB.ComboBox cboTribe 
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   315
                  ItemData        =   "frmEmployeeOld.frx":27DD
                  Left            =   4395
                  List            =   "frmEmployeeOld.frx":27DF
                  TabIndex        =   16
                  Top             =   1155
                  Width           =   2490
               End
               Begin VB.TextBox txtPhysicalAddress 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   315
                  Left            =   4395
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   13
                  Top             =   180
                  Width           =   2490
               End
               Begin VB.TextBox txtHAddress 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   315
                  Left            =   4395
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   14
                  Top             =   515
                  Width           =   2490
               End
               Begin VB.Label Label30 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Tribe:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Left            =   3120
                  TabIndex        =   124
                  Top             =   1200
                  Width           =   420
               End
               Begin VB.Label Label42 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Physical Address:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Left            =   3120
                  TabIndex        =   116
                  Top             =   180
                  Width           =   1260
               End
               Begin VB.Label Label20 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Telephone:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Left            =   120
                  TabIndex        =   115
                  Top             =   500
                  Width           =   810
               End
               Begin VB.Label Label11 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Email:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Left            =   120
                  TabIndex        =   114
                  Top             =   885
                  Width           =   420
               End
               Begin VB.Label Label22 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Home Address:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Left            =   3120
                  TabIndex        =   113
                  Top             =   480
                  Width           =   1095
               End
               Begin VB.Label Label40 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Marital status:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Left            =   120
                  TabIndex        =   112
                  Top             =   1170
                  Width           =   1035
               End
               Begin VB.Label Label31 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "*Nationality:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Left            =   120
                  TabIndex        =   111
                  Top             =   180
                  Width           =   915
               End
               Begin VB.Label Label14 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Religion:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Left            =   3120
                  TabIndex        =   110
                  Top             =   840
                  Width           =   615
               End
            End
            Begin VB.Label Label45 
               Caption         =   "Organization Unit:"
               Height          =   255
               Left            =   -74880
               TabIndex        =   129
               Top             =   360
               Width           =   1335
            End
         End
      End
      Begin VB.TextBox txtCBonus 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5745
         TabIndex        =   54
         Top             =   5295
         Visible         =   0   'False
         Width           =   1320
      End
      Begin MSComDlg.CommonDialog Cdl 
         Left            =   6720
         Top             =   5760
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtRBonus 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2985
         TabIndex        =   53
         Top             =   5280
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Line Line2 
         BorderWidth     =   5
         X1              =   0
         X2              =   7080
         Y1              =   -120
         Y2              =   -120
      End
      Begin VB.Label Label34 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Bonus"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5145
         TabIndex        =   63
         Top             =   5280
         Width           =   435
      End
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Responsibility Bonus"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1470
         TabIndex        =   62
         Top             =   5280
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Frame FraList 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   7020
      Left            =   0
      TabIndex        =   56
      Top             =   -90
      Width           =   7440
      Begin MSComctlLib.ListView lvwEmp 
         Height          =   6720
         Left            =   0
         TabIndex        =   57
         Top             =   90
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   11853
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgTree"
         ForeColor       =   0
         BackColor       =   14737632
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
      Begin VB.CommandButton cmdNew 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   135
         Picture         =   "frmEmployeeOld.frx":27E1
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Add New record"
         Top             =   6360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdEdit 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   615
         Picture         =   "frmEmployeeOld.frx":28E3
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Edit Record"
         Top             =   6360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdDelete 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1095
         Picture         =   "frmEmployeeOld.frx":29E5
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Delete Record"
         Top             =   6360
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSComctlLib.ImageList imgTree 
         Left            =   2175
         Top             =   405
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   23
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeOld.frx":2ED7
               Key             =   "A"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeOld.frx":3329
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeOld.frx":3643
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeOld.frx":3A95
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeOld.frx":3EE7
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeOld.frx":4339
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeOld.frx":4653
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeOld.frx":496D
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeOld.frx":4DBF
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeOld.frx":50D9
               Key             =   "B"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeOld.frx":552B
               Key             =   "C"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeOld.frx":597D
               Key             =   "D"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeOld.frx":5DCF
               Key             =   "E"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeOld.frx":6221
               Key             =   "F"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeOld.frx":6673
               Key             =   "G"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeOld.frx":6AC5
               Key             =   "H"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeOld.frx":6F17
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeOld.frx":7369
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeOld.frx":77BB
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeOld.frx":7C0D
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeOld.frx":805F
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeOld.frx":84B1
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeOld.frx":8903
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "CLOSE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8535
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   6420
         Visible         =   0   'False
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmEmployeeOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'======== HRCORE DECLARATIONS =========
Private outs As HRCORE.OrganizationUnitTypes
Private OUnits As HRCORE.OrganizationUnits
Private selOU As HRCORE.OrganizationUnit
Private company As New HRCORE.CompanyDetails
Private Emp As HRCORE.Employee
Private emps As HRCORE.Employees

Private empCats As HRCORE.EmployeeCategories
Private empTerms As HRCORE.EmploymentTerms
Private empTribes As HRCORE.Tribes
Private empNationalities As HRCORE.Nationalities
Private empReligions As HRCORE.Religions


'====== END OF HRCORE DECLARATIONS =========


Dim NewAcct As Boolean
Dim LastSecID As Long
Dim GenerateID As Boolean
Dim EnterDOB As Boolean
Dim EnterDEmp As Boolean
Dim MStruc As String        'holds the Main Structure from STypes

Public Sub ClearMyTexts()
dtpCDate.Enabled = False
FraList.Visible = False
Cleartxt
dtpDOB.Value = DateAdd("m", -220, Date)
'txtDOB.Text = Format(dtpDOB.Value, "yyyy-mm-dd")
dtpDEmployed.Value = Date
dtpCDate.Value = Date
EnterDOB = False
EnterDEmp = False

'Set Default Values
cboNationality.Text = "Kenyan"
CboReligion.Text = "Christian"
txtKRAFileNO.Text = "0"
cboType.ListIndex = 0
With txtBankCode
    .Text = ""
    .Tag = ""
End With

With txtBankBranch
    .Text = ""
    .Tag = ""
End With

With txtBankBranchName
    .Text = ""
    .Tag = ""
End With
dtpSPension.Value = Date
dtpTerm.Value = Date
cmdSave.Enabled = True
cmdCancel.Enabled = True
Set Picture1 = Nothing
End Sub
Public Sub SwitchEmp()
'    If PSave = False Then
'        If MsgBox("Do you want to save changes?", vbQuestion + vbYesNo) = vbYes Then
'            PSave = True
'            Call cmdSave_Click
'            Exit Sub
'        End If
'    Else
'        PSave = True
'    End If
    
    Call disabletxt
    FraList.Visible = True
    EnableCmd
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    SaveNew = False
    
    If DSource = "Local" Then
        With frmMain2
            .cmdNew.Enabled = True
            .cmdDelete.Enabled = True
            .cmdEdit.Enabled = True
            .cmdSave.Enabled = False
            .cmdCancel.Enabled = False
        End With
        
    Else
        With frmMain2
            .cmdEdit4.Enabled = True
            .cmdSave4.Enabled = False
            .cmdCancel4.Enabled = False
        End With
    
    End If
    
    Call DisplayRecords
    
End Sub
'Public Sub SwitchEmp()
'    If PSave = False Then
'        If MsgBox("Do you want to save changes?", vbQuestion + vbYesNo) = vbYes Then
'            PSave = True
'            Call cmdSave_Click
'            Exit Sub
'        End If
'    Else
'        PSave = True
'    End If
'
'    Call disabletxt
'    FraList.Visible = True
'    EnableCmd
'    cmdSave.Enabled = False
'    cmdCancel.Enabled = False
'    SaveNew = False
'
'    If DSource = "Local" Then
'        With frmMain2
'            .cmdNew.Enabled = True
'            .cmdDelete.Enabled = True
'            .cmdEdit.Enabled = True
'            .cmdSave.Enabled = False
'            .cmdCancel.Enabled = False
'        End With
'
'    Else
'        With frmMain2
'            .cmdEdit4.Enabled = True
'            .cmdSave4.Enabled = False
'            .cmdCancel4.Enabled = False
'        End With
'
'    End If
'
'    Call DisplayRecords
'
'End Sub


Private Sub cboCat_Click()
Dim rsCheckCat As New ADODB.Recordset
Set rsCheckCat = CConnect.GetRecordSet("select * from ecategory where code='" & cboCat.Text & "'")
If rsCheckCat.RecordCount > 0 Then
    cboCat.Tag = Trim(rsCheckCat!ecategory_id & "")
End If
End Sub

Private Sub cboCat_KeyPress(KeyAscii As Integer)
    Dim i As Integer
'''    KeyAscii = 0
    For i = 0 To cboCat.ListCount - 1
        If cboCat.Text Like cboCat.List(i) Then
            cboCat.TopIndex = i
        End If
    Next i
    cboCat.Text = ""
    KeyAscii = 0
End Sub

Private Sub cboCCode_Click()
    If Not cboCCode.Text = "" Then
        Set rs3 = CConnect.GetRecordSet("SELECT * FROM MyDivisions WHERE Code = '" & cboCCode.Text & "'")
        
        With rs3
            If .RecordCount > 0 Then
                cboCName.Text = !Description & ""
                cboCCode.Tag = !cstructure_id & ""
            End If
        End With
        
        Set rs3 = Nothing
    
    End If
End Sub

Private Sub cboCCode_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
    cboCCode.Text = ""
    cboCName.Text = ""
End Sub

Private Sub cboCName_Click()
    If Not cboCName.Text = "" Then
        Set rs3 = CConnect.GetRecordSet("SELECT * FROM MyDivisions WHERE Description = '" & cboCName.Text & "'")
        
        With rs3
            If .RecordCount > 0 Then
                cboCCode.Text = !code & ""
                cboCCode.Tag = !cstructure_id & ""
            End If
        End With
        
        Set rs3 = Nothing
    
    End If

End Sub

Private Sub cboCName_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    For i = 0 To cboCName.ListCount - 1
        If cboCName.Text Like cboCName.List(i) Then
            cboCName.TopIndex = i
        End If
    Next i
    cboCName.Text = ""
    KeyAscii = 0
End Sub

Private Sub cboCurCode_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboCurName_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub cboDesignation_Change()

End Sub

Private Sub cboDesignation_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboGender_KeyPress(KeyAscii As Integer)
    cboGender.Text = ""
    KeyAscii = 0
End Sub

Private Sub cboMarritalStat_KeyPress(KeyAscii As Integer)
Dim i As Integer
cboMarritalStat.Text = UCase(Chr(KeyAscii))
For i = 0 To cboMarritalStat.ListCount - 1
    If cboMarritalStat.Text Like UCase(Mid(cboMarritalStat.List(i), 1, 1)) Then
        cboMarritalStat.TopIndex = i
        Exit For
    End If
Next i
KeyAscii = 0
'cboMarritalStat.Text = ""
End Sub

Private Sub cboNationality_KeyPress(KeyAscii As Integer)
Dim i As Integer
cboNationality.Text = UCase(Chr(KeyAscii))
For i = 0 To cboNationality.ListCount - 1
    If cboNationality.Text Like UCase(Mid(cboNationality.List(i), 1, 1)) Then
        cboNationality.TopIndex = i
        Exit For
    End If
Next i
KeyAscii = 0
cboNationality.Text = ""
End Sub

Private Sub cboProbType_Click()
    If cboProbType.Text = "Appointment" Then
        'lblSDate.Visible = False
        'dtpPSDate.Visible = False
        
        lblSDate.Visible = True
        dtpPSDate.Visible = True
        dtpPSDate.Enabled = False
        
        lblCDate.Visible = True
        dtpCDate.Visible = True
        
        dtpCDate.Enabled = False
        
        txtProbationReason.Visible = True
        lblProbReason.Visible = True
        dtpCDate.Value = DateAdd("m", Val(txtProb.Text), dtpDEmployed.Value)
        
    ElseIf cboProbType.Text = "Promotion" Then
        lblSDate.Visible = True
        dtpPSDate.Visible = True
        
        dtpPSDate.Value = Date
        lblCDate.Visible = True
        dtpCDate.Visible = True
        
        dtpCDate.Enabled = True
        dtpPSDate.Enabled = True
        
        txtProbationReason.Visible = True
        lblProbReason.Visible = True
        dtpCDate.Value = DateAdd("m", Val(txtProb.Text), dtpPSDate.Value)
        
    Else
'        lblSDate.Visible = False
'        dtpPSDate.Visible = False
        lblSDate.Visible = True
        dtpPSDate.Visible = True
        dtpPSDate.Enabled = False
        
'        lblCDate.Visible = False
'        dtpCDate.Visible = False
        
        lblCDate.Visible = True
        dtpCDate.Visible = True
        dtpCDate.Enabled = False

        txtProbationReason.Visible = True
        lblProbReason.Visible = True
        txtProbationReason.Enabled = False
        txtProb.Text = ""
    End If
End Sub

Private Sub cboProbType_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
    'cboProbType.Text = ""
End Sub

Private Sub CboReligion_KeyPress(KeyAscii As Integer)
    Dim i As Integer
'''    KeyAscii = 0
    For i = 0 To CboReligion.ListCount - 1
        If CboReligion.Text Like CboReligion.List(i) Then
            CboReligion.ListIndex = i
        End If
    Next i
    CboReligion.Text = ""
    KeyAscii = 0
End Sub

Private Sub cboTermReasons_Click()
    If chkTerm.Value = 1 Then
        If cboTermReasons.Text = "Retirement" Then
            fraTerm.Visible = True
            fraTerm.Enabled = True
        Else
            'fraTerm.Visible = False
            fraTerm.Enabled = False
        End If
        
        If cboTermReasons.Text = "Death" Then chkReEngage.Value = 1: chkReEngage.Enabled = False Else chkReEngage.Enabled = True
    Else
        fraTerm.Enabled = False
    End If
End Sub

Private Sub cboTermReasons_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboTerms_Click()
On Error GoTo Hell
Dim rsMatch As New ADODB.Recordset
strDatePart = ""
strValue = 0
strCode = ""
strName = ""
cboTerms.Tag = ""
dtpValidThrough.Enabled = True
Set rsMatch = CConnect.GetRecordSet("SELECT * FROM EmpTerms WHERE code=" & cboTerms.ItemData(cboTerms.ListIndex))
If rsMatch.RecordCount > 0 Then
    If Trim(rsMatch!matchtocasual & "") = True Then
        With frmSelCasualTypes
            .Show vbModal, frmMain2
            .left = (Screen.Width - .Width) / 2
            .top = (Screen.Height - .Height) / 2
        End With
        Select Case strDatePart
        Case "d"
            dtpValidThrough.Value = DateAdd("d", strValue, dtpDEmployed.Value)
        Case "w"
            dtpValidThrough.Value = DateAdd("d", strValue * 7, dtpDEmployed.Value)
        Case "m"
            dtpValidThrough.Value = DateAdd("m", strValue, dtpDEmployed.Value)
        Case "y"
            dtpValidThrough.Value = DateAdd("m", strValue * 12, dtpDEmployed.Value)
        End Select
        dtpValidThrough.Enabled = False
        cboTerms.Tag = "cas"
    Else
        If Trim(rsMatch!matchtocontract & "") = True Then
            With frmSelContractTypes
                .Show vbModal, frmMain2
                .left = (Screen.Width - .Width) / 2
                .top = (Screen.Height - .Height) / 2
            End With
            Select Case strDatePart
            Case "d"
                dtpValidThrough.Value = DateAdd("d", strValue, dtpDEmployed.Value)
            Case "w"
                dtpValidThrough.Value = DateAdd("d", strValue * 7, dtpDEmployed.Value)
            Case "m"
                dtpValidThrough.Value = DateAdd("m", strValue, dtpDEmployed.Value)
            Case "y"
                dtpValidThrough.Value = DateAdd("m", strValue * 12, dtpDEmployed.Value)
            End Select
            dtpValidThrough.Enabled = False
            cboTerms.Tag = "cont"
        Else
            If Trim(rsMatch!MatchToPermanent & "") = True Then
                cboTerms.Tag = "perm"
                Dim rsAddDate As New ADODB.Recordset
                Set rsAddDate = CConnect.GetRecordSet("select * from GeneralOpt where subsystem='" & SubSystem & "'")
                If rsAddDate.RecordCount > 0 Then
                    dtpValidThrough.Value = DateAdd("m", IIf(cboGender.Text = "Male", IIf(IsNumeric(Trim(rsAddDate!MRet & "")) = True, Trim(rsAddDate!MRet & ""), 1) * 12, IIf(IsNumeric(Trim(rsAddDate!FRet & "")) = True, Trim(rsAddDate!FRet & ""), 1) * 12), dtpDOB.Value)
                End If
                dtpValidThrough.Enabled = False
            End If
        End If
    End If
End If
Exit Sub
Hell:
End Sub

Private Sub cboTerms_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
    cboTerms.Text = ""
End Sub

Private Sub cboTribe_KeyPress(KeyAscii As Integer)
Dim i As Integer
For i = 0 To cboTribe.ListCount - 1
    If cboTribe.Text Like cboTribe.List(i) Then
        cboTribe.TopIndex = i
    End If
Next i
cboTribe.Text = ""
End Sub

Private Sub cboType_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
    cboType.Text = ""
End Sub

Private Sub chkTerm_Click()
    If chkTerm.Value = 1 Then
        dtpTerm.Value = Date
        chkReEngage.Visible = True
        cboTermReasons.Locked = False
    Else
        fraTerm.Enabled = False '.Visible = False
        chkReEngage.Visible = False
        cboTermReasons.Locked = True
    End If
    
End Sub

'
Public Sub cmdCancel_Click()
    If PSave = False Then
        If MsgBox("Do you want to save changes?", vbQuestion + vbYesNo) = vbYes Then
            PSave = True
            Call cmdSave_Click
            Exit Sub
        End If
    Else
        PSave = True
    End If
    
    Call disabletxt
    FraList.Visible = True
    EnableCmd
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    SaveNew = False
    
    If DSource = "Local" Then
        With frmMain2
            .cmdNew.Enabled = True
            .cmdDelete.Enabled = True
            .cmdEdit.Enabled = True
            .cmdSave.Enabled = False
            .cmdCancel.Enabled = False
        End With
        
    Else
        With frmMain2
            .cmdEdit4.Enabled = True
            .cmdSave4.Enabled = False
            .cmdCancel4.Enabled = False
        End With
    
    End If
    
    Call DisplayRecords
    
End Sub

Public Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub cmdDelete_Click()
Dim Emp As String
    
'    Omnis_ActionTag = "D" 'Deletes a new record in the Omnis database 'monte++
    
    If DSource <> "Local" Then
        MsgBox "You are not allowed to delete employees since this records are from another module.", vbInformation
'        Call cmdCancel_Click
        Exit Sub
    End If
    
    If rsGlob.RecordCount > 0 Then
        resp = MsgBox("Are you sure you want to delete employee - " & rsGlob!empcode & "?", vbQuestion + vbYesNo)
        
        If resp = vbNo Then Exit Sub
        Emp = rsGlob!employee_id
        
        Employees_TextFile  '++Passes a "D" to the Omnis text file for deletion 'monte+++
        Action = "DELETED EMPLOYEE; EMPLOYEE CODE: " & rsGlob!empcode
        CConnect.ExecuteSql ("DELETE FROM Employee Where employee_id = '" & Emp & "'")
        CConnect.ExecuteSql ("DELETE FROM SEmp Where employee_id = '" & Emp & "'")
            
'        ' if AuditTrail = True Then cConnect.ExecuteSql ("INSERT INTO AuditTrail (UserId, DTime, Trans, TDesc, MySection)VALUES('" & CurrentUser & "','" & Date & " " & Time & "','Deleting Employee','" & frmMain2.lvwEmp.SelectedItem & "','Employee')")
        
        rs.Requery
        
        Set rs5 = CConnect.GetRecordSet("SELECT * FROM Security WHERE UID = '" & CurrentUser & "' AND subsystem ='" & SubSystem & "'")

        With rs5
            If Not .EOF And Not .BOF Then
                Set rsGlob = Nothing
                
                If Not IsNull(!terms) And Not IsNull(!LCode) Then
                    Set rsGlob = CConnect.GetRecordSet("SELECT e.*, c.Code, c.Description FROM (Employee as e" & _
                            " LEFT JOIN SEmp as s ON e.employee_id = s.employee_id) LEFT JOIN CStructure as c ON s.LCode = c.LCode" & _
                            "LEFT JOIN ECategory as  ec ON e.ECategory = ec.code WHERE s.LCode like '" & !LCode & "%" & _
                            "' AND e.Terms = '" & !terms & "' AND ec.seq >= '" & maxCatAccess & "' ORDER BY e.EmpCode")
                                    
                ElseIf Not IsNull(!terms) And IsNull(!LCode) Then
                    Set rsGlob = CConnect.GetRecordSet("SELECT e.*, c.Code, c.Description FROM (Employee as e LEFT JOIN " & _
                            "SEmp as s ON e.employee_id = s.employee_id) LEFT JOIN CStructure as c ON s.LCode = c.LCode" & _
                            "LEFT JOIN ECategory as ec ON e.ECategory = ec.code" & _
                            " WHERE e.Terms = '" & !terms & "' AND ec.seq >= '" & maxCatAccess & "' ORDER BY e.EmpCode")
                            
                            
                ElseIf IsNull(!terms) And Not IsNull(!LCode) Then
                    Set rsGlob = CConnect.GetRecordSet("SELECT e.*, c.Code, c.Description FROM (Employee as e LEFT JOIN " & _
                            "SEmp as s ON e.employee_id = s.employee_id) LEFT JOIN CStructure as c ON s.LCode = c.LCode" & _
                            "LEFT JOIN ECategory as  ec ON e.ECategory = ec.code" & _
                            " WHERE s.LCode like '" & !LCode & "%" & "' AND ec.seq >= '" & maxCatAccess & "' " & _
                            "ORDER BY e.EmpCode")
                            
                Else
                    Set rsGlob = CConnect.GetRecordSet("SELECT e.*, c.Code, c.Description FROM (Employee as e " & _
                            "LEFT JOIN SEmp as s ON e.employee_id = s.employee_id) LEFT JOIN CStructure as c ON s.LCode" & _
                            "= c.LCode LEFT JOIN ECategory as  ec ON e.ECategory = ec.code WHERE " & _
                            "ec.seq >= '" & maxCatAccess & "' ORDER BY e.EmpCode")
        
                End If
            
        
            End If
        End With
                    
        Set rs5 = Nothing
        
        
        Call LoadList
        frmMain2.LoadMyList
        Call DisplayRecords
        
        Call frmMain2.cboTerms_Click

        
        'frmMain2.lblECount.Caption = rsGlob.RecordCount
        
     
                       
    Else
        MsgBox "No records to be deleted.", vbInformation
    End If

End Sub

Public Sub cmdEdit_Click()
    On Error GoTo ErrorTrap
    Omnis_ActionTag = "E" 'Edits a record in the Omnis database 'monte++
    
    If Not (SelectedEmployee Is Nothing) Then
        'Call DisplayRecords
        FraList.Visible = False
        enabletxt
        'txtBankBranch.Locked = True
        'txtBankName.Locked = True
        'txtBankBranchName.Locked = True
        dtpCDate.Enabled = False
        'DisableCmd
        cmdSave.Enabled = True
        cmdCancel.Enabled = True
        'txtEmpCode.Locked = True
        txtSurname.SetFocus
        SaveNew = False
    Else
        MsgBox "Select the Employee to Edit", vbInformation, TITLES
        PSave = True
        Call cmdCancel_Click
        PSave = False
    End If
   
    Exit Sub
    
ErrorTrap:
    MsgBox Err.Description, vbExclamation, TITLES
End Sub



Public Sub cmdNew_Click()
    On Error Resume Next
    Omnis_ActionTag = "I" 'Inserts a new record in the Omnis database 'monte++
    new_Record = True
    If DSource <> "Local" Then
        MsgBox "You are not allowed to add new employees since this records are from another module.", vbInformation
        PSave = True
        Call cmdCancel_Click
        PSave = False
        Exit Sub
    End If
    
    dtpCDate.Enabled = False
    FraList.Visible = False
    enabletxt
    Cleartxt
    cboGender.Text = ""
    'dtpDOB.Value = DateAdd("m", -220, Date)
    'txtDOB.Text = "" 'Format(dtpDOB.Value, "yyyy-mm-dd")
    dtpDEmployed.Value = Date
    dtpCDate.Value = Date
    EnterDOB = False
    EnterDEmp = False
    
    'Set Default Values
    cboNationality.Text = ""
    CboReligion.Text = ""
    txtKRAFileNO.Text = "0"
    'cboType.ListIndex = 0
    cboType.Text = ""
    With txtBankCode
        .Text = ""
        .Tag = ""
    End With
    
    With txtBankBranch
        .Text = ""
        .Tag = ""
    End With
    
    With txtBankBranchName
        .Text = ""
        .Tag = ""
    End With
    dtpSPension.Value = Date
    txtEmpCode.SetFocus
    Call GenID
    dtpTerm.Value = Date
'    DisableCmd
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    SaveNew = True
    Set Picture1 = Nothing
    
End Sub

Private Sub cmdPDelete_Click()
Dim resp As String
Dim picturepath As String
Dim MovePic As FileSystemObject

resp = MsgBox("Are you sure you want to delete picture?", vbQuestion + vbYesNo)
If resp = vbNo Then
    Exit Sub
End If


Set MovePic = New FileSystemObject

On Error Resume Next
'picturepath = App.Path & "\Pic\Gen.jpg"
'Picture1.Picture = LoadPicture(App.Path & "\Pic\Gen.jpg")
'MovePic.CopyFile picturepath, App.Path & "\Photos\" & txtRegNo.Text & ".jpg"
'MovePic.DeleteFile
MovePic.DeleteFile (App.Path & "\Photos\" & txtEmpCode.Text & ".jpg")
Set Picture1 = Nothing
Picture1.Picture = LoadPicture(App.Path & "\Pic\Gen.jpg")
Action = "DELETED EMPLOYEE PHOTO; EMPLOYEE CODE: " & rsGlob!empcode & "; PHOTO PATH: " & App.Path & "\Photos\" & txtEmpCode.Text & ".jpg"
CConnect.ExecuteSql "SELECT * FROM EMPLOYEE"
End Sub

Private Sub cmdPNew_Click()

Dim picturepath As String
Dim MovePic As FileSystemObject


If Len(Trim(txtEmpCode.Text)) <= 0 Then
    MsgBox "Enter Employee Code first", vbInformation, "Picture"
    Exit Sub
Else
    Set MovePic = New FileSystemObject
    
    With Cdl
        .Filter = "Pictures {*.bmp;*.gif;*.jpeg;*.ico;*.jpg;*.ICO;*.JPEG;*.JPG;*.BMP;*.GIF|*.bmp;*.gif;*.jpeg;*.ico;*.jpg;*.ICO;*.JPEG;*.JPG;*.BMP;*.GIF"
        .ShowOpen
    End With
    

    picturepath = Cdl.FileName
    
End If

If Len(Trim(picturepath)) > 0 Then
    If Cdl.FileName Like "*.ico" Or _
        Cdl.FileName Like "*.jpeg" Or _
        Cdl.FileName Like "*.jpg" Or _
        Cdl.FileName Like "*.bmp" Or _
        Cdl.FileName Like "*.gif" Or _
        Cdl.FileName Like "*.ICO" Or _
        Cdl.FileName Like "*.JPEG" Or _
        Cdl.FileName Like "*.JPG" Or _
        Cdl.FileName Like "*.BMP" Or _
        Cdl.FileName Like "*.gif" Then
        
        On Error Resume Next
        Picture1.Picture = LoadPicture(picturepath)
        MovePic.CopyFile picturepath, App.Path & "\Photos\" & txtEmpCode.Text & ".jpg", True
        Action = "SPECIFIED EMPLOYEE PHOTO; EMPLOYEE CODE: " & rsGlob!empcode & "; PHOTO PATH: " & App.Path & "\Photos\" & txtEmpCode.Text & ".jpg"
        CConnect.ExecuteSql "SELECT * FROM EMPLOYEE"
    Else
        MsgBox "Unsupported file format", vbExclamation, "Picture"
        On Error Resume Next
        Picture1.Picture = LoadPicture(App.Path & "\Pic\Gen.jpg")
    End If
  
End If
End Sub

Public Sub cmdSave_Click()
Dim Term As Integer
Dim Pension As String
Dim i, countForbidden As Integer
Dim Unsolicited As String
Dim PromptDate As Date
Dim rec_emp As New ADODB.Recordset, emp_id As Integer

On Error GoTo errHandler

    'validation of entries
    If ((txtAlien.Text = "") And (txtPassport.Text = "") And (txtIDNo.Text = "")) Then MsgBox "Please ensure you've entered any of:" & vbCrLf & "Alien, Passport and ID Number", vbOKOnly + vbInformation, "Missing Details": Exit Sub
    
    If ((CheckForNumbers(txtSurname.Text) > 0) Or (CheckForNumbers(txtONames.Text) > 0)) Then
        MsgBox "Employee's name cannot contain numbers," & vbCrLf & "please change to continue.", vbOKOnly + vbInformation, "Forbidden characters!"
        Exit Sub
    End If
    
    If txtTel.Text <> "" And (CheckForNumbers(txtTel.Text) = 0) Then
        MsgBox "Telephone numbers must at least contain few numeric entries," & vbCrLf & "please change to continue.", vbOKOnly + vbInformation, "Forbidden characters!"
        txtTel.SetFocus
        Exit Sub
    End If
    
    If CheckForNumbers(txtEmpCode.Text) = 0 Then
        MsgBox "Employee's code must contain at least few numeric entries," & vbCrLf & "please change to continue.", vbOKOnly + vbInformation, "Forbidden characters!"
        txtEmpCode.SetFocus
        Exit Sub
    End If
    
    If txtAccountNo.Text <> "" And CheckForNumbers(txtAccountNo.Text) = 0 Then
        MsgBox "Account number must contain at least few numeric entries," & vbCrLf & "please change to continue.", vbOKOnly + vbInformation, "Forbidden characters!"
        txtAccountNo.SetFocus
        Exit Sub
    End If
    
    If (((txtBankName.Text = "") Or (txtBankBranchName.Text = "")) And (CheckForNumbers(txtAccountNo.Text) > 0)) Then MsgBox "Please specify the bank and the corresponding bank where the" & vbCrLf & "specified account belongs.", vbOKOnly + vbInformation, "Incomplete bank details": Exit Sub
    
    If CheckForNumbers(txtIDNo.Text) = 0 And CheckForNumbers(txtAlien.Text) = 0 And CheckForNumbers(txtPassport.Text) = 0 Then
        MsgBox "Employee's ID, Alien card & Passport number must contain at least few numeric entries," & vbCrLf & "please change to continue.", vbOKOnly + vbInformation, "Forbidden characters!"
        txtIDNo.SetFocus
        Exit Sub
    End If
    
    If cboCat.Text = "" Then MsgBox "Please specify the employee grade to proceed.", vbInformation + vbOKOnly, "Missing mandatory information": cboCat.SetFocus: Exit Sub
    
    If cboType.Text = "" Then MsgBox "Please specify the employee type to proceed.", vbInformation + vbOKOnly, "Missing mandatory information": cboType.SetFocus: Exit Sub
    
    If cboTerms.Text = "" Then MsgBox "Please specify the employee's employment terms to proceed.", vbInformation + vbOKOnly, "Missing mandatory information": cboTerms.SetFocus: Exit Sub
    
    If Not IsDate(dtpDEmployed.Value) Then MsgBox "You've not specified the date of employment," & vbCrLf & "please do so to proceed.", vbOKOnly + vbInformation, "Missing mandatory information": Exit Sub
    
    'end validation of entries
    
    'If txtDOB.Text = "" Then MsgBox "The date of birth is a mandatory field," & vbCrLf & " please specify to proceed.", vbOKOnly + vbExclamation, "Missing date of Birth": Exit Sub
        
    If (DateDiff("m", dtpDOB.Value, Date) / 12) < 18 Then MsgBox "The selected individual is below the stipulated" & vbCrLf & "employment age. Please check on the date of birth.", vbOKOnly + vbExclamation, "Under age": Exit Sub
    
    'what is Unsolicited supposed to mean
    Unsolicited = 0
    
    If cboCCode.Text = "" Then MsgBox "Please specify the employee's department.", vbOKOnly + vbInformation, "Missing department": cboCCode.SetFocus: Exit Sub
        
    If chkTerm.Value = 1 Then
        If MsgBox("This action will initiate the employee termination process!", vbInformation + vbOKCancel) = vbCancel Then
            Exit Sub
        End If
'        Do you wish to validate employee data?
        If MsgBox("Do you wish to validate employee data?", vbInformation + vbYesNo) = vbYes Then
            If validateEmployeeData = False Then Exit Sub
        End If
    Else
        If validateEmployeeData = False Then
            Exit Sub
        End If
    End If

    Term = chkTerm.Value
    Pension = chkPension.Value

    If SaveNew = True Then
        Set rs1 = CConnect.GetRecordSet("SELECT * FROM Employee WHERE EmpCode = '" & txtEmpCode.Text & "'")

        With rs1
            If .RecordCount > 0 Then
                MsgBox "The employee code exists. Enter another code.", vbInformation
                txtEmpCode.Text = ""
                txtEmpCode.SetFocus
                Set rs1 = Nothing
                Call CancelMain
                Exit Sub
            End If
        End With
        Set rs1 = Nothing

        If txtDesig.ListCount = 0 Then
            chkUnsolicited.Value = 1
            Unsolicited = 1
        End If
    End If
    
    Dim rsCheckSal As New ADODB.Recordset
    Set rsCheckSal = CConnect.GetRecordSet("SELECT * FROM ECategory WHERE ECategory_ID=" & IIf(IsNumeric(cboCat.Tag) = True, cboCat.Tag, -1) & " AND (LowestSalary<=" & IIf(IsNumeric(txtBasicPay.Text) = True, txtBasicPay.Text, -1) & " AND HighestSalary>=" & IIf(IsNumeric(txtBasicPay.Text) = True, txtBasicPay.Text, -1) & ")")
    If rsCheckSal.RecordCount = 0 Then
        If MsgBox("You are about to register an employee whose grade does not match" & vbNewLine & "the  provided basic pay. Do you wish to continue?", vbYesNo + vbQuestion, "Salary, Grade Discrepancy!") = vbNo Then Exit Sub
    End If
    
    If PromptSave = True Then
        If MsgBox("Please confirm your request to save this record.", vbQuestion + vbYesNo) = vbNo Then
            Call CancelMain
            Exit Sub
        End If
    End If

'        Initialise Progrssbar
    Dim pCount As Integer
    fraProgress.Visible = True
    FraEdit.Enabled = False
    fraProgress.Caption = "Verifying input!"
    prgSave.Value = pCount
'       Initialise the employee by entering a blank record if this is a new entry - then update the data
    If SaveNew = True Then
        CConnect.ExecuteSql ("INSERT INTO employee (empcode) VALUES ('')")

    '       Obtain the new employee_id
        Set rec_emp = CConnect.GetRecordSet("SELECT TOP 1 employee_id FROM employee ORDER BY employee_id DESC")
        If rec_emp.EOF = False Then 'Should always be true
            emp_id = rec_emp!employee_id
        End If
    Else
        emp_id = rsGlob!employee_id
    End If

    pCount = 5
    prgSave.Value = pCount
    fraProgress.Caption = "Updating salary changes!"
    'Call cboTerms_Click
'            Record new salary
    Call updateSalaryChanges(CDbl(txtBasicPay.Text), CDbl(txtHAllow.Text), CDbl(txtTAllow.Text), CDbl(txtOAllow.Text), CDbl(txtLAllow.Text), emp_id)

    mySQL = "UPDATE employee SET empcode = '" & txtEmpCode.Text & "', surname = '" & cQ(txtSurname.Text) & _
            "', othernames = '" & cQ(txtONames.Text) & "', idno = '" & cQ(txtIDNo.Text) & "',Passport = '" & cQ(txtPassport.Text) & "',AlienNo = '" & cQ(txtAlien.Text) & "', dob = '" & _
            SQLDate(dtpDOB.Value) & "', demployed = '" & SQLDate(dtpDEmployed.Value) & _
            "', basicpay = " & IIf(txtBasicPay.Text <> "", CDbl(txtBasicPay.Text), "0") & ", cstructure_id = " & cboOU.ItemData(cboOU.ListIndex) & ", dcode = '" & cboCCode.Text & _
            "', type = '" & cboType.Text & "', pinno = '" & txtPin.Text & _
            "', nssfno = '" & txtNssf.Text & "', nhifno = '" & txtNhif.Text & "', term = " & Term & _
            ", dleft = '" & SQLDate(dtpTerm.Value) & "', ECategory = '" & cboCat.Text & _
            "', htel = '" & cQ(txtTel.Text) & "', Email = '" & cQ(txtEMail.Text) & "', haddress = '" & _
            txtHAddress.Text & "', desig = '" & cQ(txtDesig.Text) & "', gender = '" & cboGender.Text & _
            "', lallow = " & CDbl(txtLAllow.Text) & ", hallow = " & CDbl(txtHAllow.Text) & _
            ", oallow = " & CDbl(txtOAllow.Text) & ", tallow = " & CDbl(txtTAllow.Text) & _
            ", termtrain = '" & chkTermTrain.Value & "', termdate = '" & SQLDate(dtpTerminalDate.Value) & _
            "',advisor = '" & txtAdvisor.Text & "', achieved = '" & chkAchieved.Value & "', certno = '" & _
            txtCert.Text & "', nationality ='" & cboNationality.Text & "', tribe = '" & cboTribe.Text & _
            "', pension = " & chkPension.Value & ", disabled = " & chkDisabled.Value & ",payroll = " & _
            chkPayroll.Value & ", termreasons = '" & cboTermReasons.Text & "', unsolicited = " & _
            Unsolicited & ", rbonus = " & CDbl(txtRBonus.Text) & ", cbonus = " & _
            CDbl(txtCBonus.Text) & ", prob = " & txtProb.Text & ", cdate = '" & SQLDate(PromptDate) & _
            "', spension = '" & SQLDate(dtpSPension.Value) & "', probtype = '" & cboProbType.Text & "', psdate = '" & _
            SQLDate(dtpPSDate.Value) & "', Religion = '" & CboReligion.Text & "', bankcode = '" & txtBankCode.Text & _
            "',bankname = '" & txtBankName.Text & "', bankbranch = '" & txtBankBranchName.Tag & _
            "', bankbranchname = '" & txtBankBranchName.Text & "', AccountNo = '" & cQ(txtAccountNo.Text) & _
            "', KRAFileNO = '" & txtKRAFileNO.Text & "',Marital_Status='" & cboMarritalStat.Text & "',EmploymentValidThro='" & SQLDate(dtpValidThrough.Value) & "',PhysicalAddress='" & Replace(txtPhysicalAddress.Text, "'", "''") & "',ECategory_id=" & cboCat.Tag & ",ProbationReason='" & Replace(txtProbationReason.Text, "'", "''") & "' WHERE employee_id = " & emp_id
            
    pCount = 15
    prgSave.Value = pCount
    fraProgress.Caption = "Saving employee data!"
    Dim mySQLBank As String
    If chkTerm.Value = 0 Then
        If ((txtBankBranchName.Text <> "") And (txtBankBranchName.Text <> "") And (txtBankBranchName.Tag <> "")) Then
            mySQLBank = "INSERT INTO employeeBanks(employee_id,branchID,AccNumber,AccType,MainAcct) VALUES('" & emp_id & "','" & txtBankBranchName.Tag & "','" & txtAccountNo.Text & "','Savings',1)"
            CConnect.ExecuteSql "DELETE FROM EmployeeBanks WHERE employee_id='" & emp_id & "' AND branchID='" & txtBankBranchName.Tag & "'"
            
            Action = "SET THE EMPLOYEE MAIN ACCOUNT; EMPLOYEE CODE: " & txtEmpCode.Text & "; BANK: " & txtBankName.Text & "; BRANCH: " & txtBankBranchName.Text & "; ACCOUNT NUMBER: " & txtAccountNo.Text
            
            CConnect.ExecuteSql "UPDATE EmployeeBanks SET MainAcct=0 WHERE employee_id='" & emp_id & "'" ' AND branchID='" & txtBankBranchName.Tag & "'"
            CConnect.ExecuteSql (mySQLBank)
        End If
        Action = "ADDED DETAILS FOR EMPLOYEE; EMPLOYEE CODE: " & txtEmpCode.Text
        CConnect.ExecuteSql (mySQL)
    Else
        CConnect.ExecuteSql ("UPDATE employee SET Term = 1,CanReEngage=" & chkReEngage.Value & " WHERE employee_id = " & emp_id)
        CConnect.ExecuteSql ("UPDATE employee SET Termdate ='" & SQLDate(dtpTerm.Value) & "',dleft='" & SQLDate(dtpTerm.Value) & "' WHERE employee_id = " & emp_id)
    End If
    
    If cboTerms.ListIndex > -1 Then
        mySQL = "UPDATE  employee SET termsID=" & cboTerms.ItemData(cboTerms.ListIndex) & ", terms = '" & cboTerms.Text & "' WHERE employee_id=" & emp_id
        CConnect.ExecuteSql mySQL
    End If
    'Capture casual, contract details
    If strCode <> "" Then
        Dim rsSaveTypes As New ADODB.Recordset
        Select Case cboTerms.Tag
        Case "cont"
            Set rsSaveTypes = CConnect.GetRecordSet("SELECT * FROM Contracts WHERE employee_id='" & emp_id & "' AND code='" & strCode & "'")
            If rsSaveTypes.RecordCount = 0 Then
                CConnect.ExecuteSql "UPDATE contracts SET IsActive=0 WHERE employee_id='" & emp_id & "'"
                CConnect.ExecuteSql "INSERT INTO Contracts(empcode,code,description,cfrom,cto,employee_id,IsActive) VALUES('" & txtEmpCode.Text & "','" & strCode & "','" & strName & "','" & SQLDate(dtpDEmployed.Value) & "','" & SQLDate(dtpValidThrough.Value) & "','" & emp_id & "',1)"
            Else
                CConnect.ExecuteSql "UPDATE contracts SET IsActive=0 WHERE employee_id='" & emp_id & "'"
                CConnect.ExecuteSql "UPDATE contracts SET cto='" & SQLDate(dtpValidThrough.Value) & "',cfrom='" & SQLDate(dtpDEmployed.Value) & "',IsActive=1"
            End If
        Case "cas"
            Set rsSaveTypes = CConnect.GetRecordSet("SELECT * FROM casuals WHERE employee_id='" & emp_id & "' AND code='" & strCode & "'")
            If rsSaveTypes.RecordCount = 0 Then
                CConnect.ExecuteSql "UPDATE casuals SET IsActive=0 WHERE employee_id='" & emp_id & "'"
                CConnect.ExecuteSql "INSERT INTO casuals(empcode,code,description,cfrom,cto,employee_id,IsActive) VALUES('" & txtEmpCode.Text & "','" & strCode & "','" & strName & "','" & SQLDate(dtpDEmployed.Value) & "','" & SQLDate(dtpValidThrough.Value) & "','" & emp_id & "',1)"
            Else
                CConnect.ExecuteSql "UPDATE casuals SET IsActive=0 WHERE employee_id='" & emp_id & "'"
                CConnect.ExecuteSql "UPDATE casuals SET cto='" & SQLDate(dtpValidThrough.Value) & "',cfrom='" & SQLDate(dtpDEmployed.Value) & "',IsActive=1"
            End If
        End Select
    End If
    'End Capture casual, contract details
    
    Set rec_emp = CConnect.GetRecordSet("SELECT * FROM SEmp WHERE employee_id = '" & emp_id & "'")

    If cboCCode.Text <> "" Then
        Set rs5 = CConnect.GetRecordSet("SELECT * FROM MyDivisions WHERE Code = '" & cboCCode.Text & "'")
        With rs5
            If .RecordCount > 0 Then
               .MoveFirst
               'Set rec_emp = CConnect.GetRecordSet("SELECT * FROM SEmp WHERE employee_id = '" & emp_id & "'")
                If rec_emp.RecordCount < 1 Then
                    CConnect.ExecuteSql ("INSERT INTO SEmp (SCode, LCode, employee_id, cstructure_id) VALUES('" & !scode & "','" & !RLCode & "'," & emp_id & "," & !cstructure_id & ")")
                Else
                    CConnect.ExecuteSql ("UPDATE SEmp SET SCode = '" & !scode & "', LCode = '" & !RLCode & "', cstructure_id = " & !cstructure_id & " WHERE employee_id = " & emp_id)
                End If
            End If
        End With
    Else
        MsgBox "Since you've not provided the employee's department," & vbCrLf & "you will be expected to edit the record later for the" & vbCrLf & "correct operation of the system." & vbCrLf & "This is because the system considers your records partially complete.", vbOKOnly + vbInformation, "Missing department"
    End If
    
    Set rec_emp = CConnect.GetRecordSet("SELECT * FROM SEmp WHERE employee_id = '" & emp_id & "'")
    
    pCount = pCount + 10
    prgSave.Value = pCount
    fraProgress.Caption = "Updating Job Progression!"

'           Record job progression
    rec_emp.Requery
    If rec_emp.RecordCount > 0 Then
        Call updateJobProgression(emp_id, rec_emp!LCode & "", cQ(txtDesig.Text), cboCat.Text, cboTerms.Text, CDbl(txtBasicPay.Text), CDbl(txtHAllow.Text), CDbl(txtLAllow.Text), CDbl(txtTAllow.Text), CDbl(txtOAllow.Text))
    End If
    
    If GenerateID = True Then
        CConnect.ExecuteSql ("UPDATE GeneralOpt SET LastSecID = " & LastSecID & "")
        GenerateID = False
    End If

    pCount = pCount + 30 '65
    prgSave.Value = pCount
    fraProgress.Caption = "Reloading employee data!"

    Set rsGlob2 = CConnect.deptFilter("SELECT e.*, c.Code, c.Description FROM (Employee as e LEFT JOIN SEmp " & _
        "as s ON e.employee_id = s.employee_id) LEFT JOIN CStructure as c ON s.LCode = c.LCode LEFT JOIN ECategory " & _
        "as ec ON e.ECategory = ec.code WHERE e.Term <> 1 AND (((s.SCode)='" & MStruc & "')) OR (((s.SCode)" & _
        " Is Null)) AND ec.seq >= '" & maxCatAccess & "' ORDER BY e.EmpCode")

    rs.Requery
    rsGlob.Requery
    Set rsGlob = CConnect.setGlobalRecordset

    If SaveNew = True Then
        Call LoadList
        frmMain2.LoadMyList
''''    End If
'''''
'''''        Call frmMain2.cboTerms_Click
'''''
''''        If SaveNew = True Then
''''
        Call Cleartxt
        Call GenID
    Else
        With rsGlob
            If .RecordCount > 0 Then
                .MoveFirst
                .Find "employee_id like " & emp_id, , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    PSave = True
                    Call cmdCancel_Click
                    PSave = False
                End If
            End If
        End With
        frmMain2.LoadMyList
    End If

    'Employees_TextFile  '++Caters for the Omnis text file 'Monte++

    frmMain2.lblECount.Caption = rsGlob.RecordCount

    pCount = pCount + 10 '75
    prgSave.Value = pCount
    fraProgress.Caption = "Verifying terminated employees!"

''''''' Move Terminated employee to history
'//Commented by Juma because it makes re-engagement of terminated employee impossible due to loss of link
'    If SaveNew = False Then
'
        'If chkTerm.Value = 1 And chkPayroll.Value = 0 Then
        If chkTerm.Value = 1 Then
            Dim ctran As New CTransfer
            'ctran.Transfer_Employee emp_id, , True
            frmMain2.lvwEmp.ListItems.Remove frmMain2.lvwEmp.SelectedItem.Index
        End If
'    End If
'//End of Juma's comment
    pCount = pCount + 10 '85
    prgSave.Value = pCount
    fraProgress.Caption = "Reloading employee data!"

    pCount = 100
    prgSave.Value = pCount
    fraProgress.Visible = False
    FraEdit.Enabled = True

    MsgBox "Employee Saved successfully.", vbInformation
    new_Record = False
Exit Sub
errHandler:
    'MsgBox Err.Description, vbInformation
    fraProgress.Visible = False
    FraEdit.Enabled = True
End Sub

Private Sub cmdSchBank_Click()
    frmSelBanks.Show vbModal
    SelectedBank = strName
    txtBankCode = strName
    txtBankCode.Tag = strName
    If strName <> "" Then
    Dim rsBanks As Recordset
        Set rsBanks = New Recordset
        Set rsBanks = CConnect.GetRecordSet("SELECT * FROM tblBank where bank_id = " & strName)
        If rsBanks.EOF = False Then
            If Not IsNull(rsBanks!bank_Name) Then txtBankName = rsBanks!bank_Name
        End If
        Set rsBanks = Nothing
    End If
End Sub

Private Sub cmdSchBank_GotFocus()
txtBankName.SetFocus
Call cmdSchBank_KeyPress(9)
End Sub

Private Sub cmdSchBank_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Then cmdSchBank_Click
End Sub

Private Sub cmdSchBankBranch_Click()
    SelectedBank = ""
    SelectedBank = txtBankCode.Tag
    frmSelBankBranches.Show vbModal
    txtBankBranch = strName

    If strName <> "" Then
    Dim rsBanks As Recordset
    Set rsBanks = New Recordset
        Set rsBanks = CConnect.GetRecordSet("SELECT * FROM tblBankBranch where bank_id=" & SelectedBank) 'BankBranch_Code = " & strName)
        If Not IsNull(rsBanks!BANKBRANCH_NAME) Then txtBankBranchName = strBranchName: txtBankBranchName.Tag = strBranchID
    Set rsBanks = Nothing
    End If
End Sub

Private Sub cmdSchBankBranch_GotFocus()
txtBankBranchName.SetFocus
Call cmdSchBankBranch_KeyPress(9)
End Sub

Private Sub cmdSchBankBranch_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Then cmdSchBankBranch_Click
End Sub

Private Sub dtpDEmployed_CloseUp()
    EnterDEmp = True
    Call CheckDEmployed
    'txtDEmployed.Text = Format(dtpDEmployed.Value, "yyyy-mm-dd")
    Select Case strDatePart
    Case "d"
        dtpValidThrough.Value = DateAdd("d", strValue, dtpDEmployed.Value)
    Case "w"
        dtpValidThrough.Value = DateAdd("d", strValue * 7, dtpDEmployed.Value)
    Case "m"
        dtpValidThrough.Value = DateAdd("m", strValue, dtpDEmployed.Value)
    Case "y"
        dtpValidThrough.Value = DateAdd("m", strValue * 12, dtpDEmployed.Value)
    End Select
End Sub

Private Sub dtpDEmployed_KeyPress(KeyAscii As Integer)
    EnterDEmp = True
    Call CheckDEmployed
End Sub

Private Sub dtpDOB_CloseUp()
    EnterDOB = True
    Call CheckDOB
    Dim rsAddDate As New ADODB.Recordset
    Set rsAddDate = CConnect.GetRecordSet("select * from GeneralOpt where subsystem='" & SubSystem & "'")
    If rsAddDate.RecordCount > 0 Then
        dtpValidThrough.Value = DateAdd("m", IIf(cboGender.Text = "Male", IIf(IsNumeric(Trim(rsAddDate!MRet & "")) = True, Trim(rsAddDate!MRet & ""), 1) * 12, IIf(IsNumeric(Trim(rsAddDate!FRet & "")) = True, Trim(rsAddDate!FRet & ""), 1) * 12), dtpDOB.Value)
    End If
End Sub

Private Sub dtpDOB_KeyPress(KeyAscii As Integer)
    EnterDOB = True
    Call CheckDOB
End Sub

Private Sub dtpPSDate_Change()
    dtpCDate.Value = DateAdd("m", Val(txtProb.Text), dtpPSDate.Value)
End Sub

Private Sub dtpValidThrough_CloseUp()
'txtValidThrough.Text = Format(dtpValidThrough.Value, "yyyy-mm-dd")
'If ((IsDate(txtValidThrough.Text) = True) And (IsDate(txtDEmployed.Text) = True)) Then
If DateDiff("d", dtpDEmployed.Value, dtpValidThrough.Value) < 0 Then MsgBox "The validity period of employment cannot be earlier than date of employment.", vbExclamation + vbOKOnly, "Wrong date"
'End If
End Sub

Private Sub Form_Load()

    On Error GoTo errorHandler
    Call InitializeHRCOREObjects
    Call disabletxt
    Call InitGrid       'the employee listview
    Call LoadEmployees
    Call DisplayRecords
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    If ViewSal = True Then
        fraSal.Visible = True
    Else
        fraSal.Visible = False
    End If

    'position the form
    frmMain2.PositionTheFormWithEmpList Me
    
    Exit Sub
    
errorHandler:
    MsgBox Err.Description, vbExclamation, TITLES
End Sub

Function Load_Individual_Emp(sEmpCode As String)
Call Cleartxt

If rsGlob.EOF = True And rsGlob.BOF = True Then Exit Function
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "EmpCode = '" & sEmpCode & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
            
                frmMain2.txtDetails.Caption = "Code: " & rsGlob!empcode & "     " & "Name: " & !SurName & "" & " " & !OtherNames & "" & " " & vbCrLf & _
                "" & "ID No:" & " " & !IdNo & "" & "     " & "Date Employed:" & " " & !DEmployed & "" & "     " & "Gender:" & " " & !Gender & ""
                
                txtEmpCode.Text = !empcode & ""
                txtSurname.Text = !SurName & ""
                txtONames.Text = !OtherNames & ""
                txtIDNo.Text = !IdNo & ""
                cboGender.Text = !Gender & ""
                If Not IsNull(!DOB) Then dtpDOB.Value = !DOB Else dtpDOB.Value = Date
                'txtDOB.Text = !DOB & ""
                If Not IsNull(!DEmployed) Then dtpDEmployed.Value = !DEmployed Else dtpDOB.Value = Date
                'txtDEmployed.Text = !DEmployed & ""
                cboCCode.Text = !code & ""
                cboCName.Text = !Description & ""
                cboTerms.Text = !terms & ""
                cboType.Text = !Type & ""
                cboTermReasons.Text = !TermReasons & ""
                txtPin.Text = !PinNo & ""
                txtNssf.Text = !NssfNo & ""
                txtNhif.Text = !NhifNo & ""
                cboCat.Text = !ECategory & ""
                txtTel.Text = !HTel & ""
                txtHAddress.Text = !HAddress & ""
                txtDesig.Text = !Desig & ""
                txtEMail.Text = !EMail & ""
                txtCert.Text = !CertNo & ""
                cboNationality.Text = !Nationality & ""
                cboTribe.Text = !Tribe & ""
                txtBasicPay.Text = Format(!BasicPay, Cfmt)
                txtLAllow.Text = Format(!lallow & "", Cfmt)
                txtHAllow.Text = Format(!hallow & "", Cfmt)
                txtOAllow.Text = Format(!oallow & "", Cfmt)
                txtTAllow.Text = Format(!tallow & "", Cfmt)
                txtRBonus.Text = Format(!RBonus & "", Cfmt)
                txtCBonus.Text = Format(!CBonus & "", Cfmt)
                txtGrossPay.Text = Format(!tallow + !hallow + !oallow + !BasicPay, Cfmt)
                chkDisabled.Value = !Disabled & ""
                chkPayroll.Value = !Payroll & ""
                chkUnsolicited.Value = !Unsolicited & ""
                txtProb.Text = !Prob & ""
                If !SPension <> "" Then dtpSPension.Value = !SPension
                
                cboProbType.Text = !ProbType & ""
                lblSDate.Visible = False
                dtpPSDate.Visible = False
                lblCDate.Visible = False
                dtpCDate.Value = False
                
                If !ProbType = "Appointment" Then
                    lblSDate.Visible = False
                    dtpPSDate.Visible = False
                    lblCDate.Visible = True
                    dtpCDate.Visible = True
                    txtProbationReason.Visible = True
                    lblProbReason.Visible = True
                    txtProbationReason.Text = !probationReason & ""
                    If !CDate <> "" Then
                        dtpCDate.Value = !CDate & ""
                    Else
                        If Not IsNull(!DEmployed) Then
                            dtpCDate.Value = !DEmployed & ""
                        Else
                            dtpCDate.Value = Date
                        End If
                    End If
                
                ElseIf !ProbType = "Promotion" Then
                    lblSDate.Visible = True
                    dtpPSDate.Visible = True
                    lblCDate.Visible = True
                    dtpCDate.Visible = True
                    txtProbationReason.Visible = True
                    lblProbReason.Visible = True
                    
                    If !PSDate <> "" Then
                        dtpPSDate.Value = !PSDate
                    Else
                        dtpPSDate.Value = Date
                    End If
                    
                    If !CDate <> "" Then
                        dtpCDate.Value = !CDate & ""
                    Else
                        dtpCDate.Value = Date
                    End If
                End If
                
                If !Term = True Then
                    chkTerm.Value = 1
                    If Not !dleft = "" Then dtpTerm.Value = !dleft
                    If cboTermReasons = "Retirement" Then
                        fraTerm.Visible = True
                        
                        chkTermTrain.Value = !TermTrain & ""
                        dtpTerminalDate.Value = Format(!termdate & "", Dfmt)
                        txtAdvisor.Text = !Advisor & ""
                        chkAchieved.Value = !Achieved & ""
                    End If
                Else
                    chkTerm.Value = 0
                    fraTerm.Visible = False
                End If
                
                If !Pension = True Then
                    chkPension.Value = 1
                Else
                    chkPension.Value = 0
                    
                End If
                    
                
                Set Picture1 = Nothing
        
                On Error Resume Next 'this handler is specific to the photos only
                Picture1.Picture = LoadPicture(App.Path & "\Photos\" & txtEmpCode.Text & ".jpg")
                
                If Picture1.Picture = 0 Then
                    On Error Resume Next
                    Picture1.Picture = LoadPicture(App.Path & "\Pic\Gen.jpg")
                End If
            End If
        End If
    End With
End Function


Public Sub DisplayRecords()

On Error GoTo errorHandler
Call Cleartxt
If Not (SelectedEmployee Is Nothing) Then
    With SelectedEmployee
        txtEmpCode.Text = .empcode
        txtSurname.Text = .SurName
        txtONames.Text = .OtherNames
        txtIDNo.Text = .IdNo
        cboGender.Text = .GenderStr
        txtPhysicalAddress.Text = .physicaladdress
        dtpDOB.Value = .DateOfBirth
        dtpDEmployed.Value = .DateOfEmployment
        cboOU.Text = .OrganizationUnit.OrganizationUnitName
        cboTerms.Text = .EmploymentTerm.EmpTermName
        cboType.Text = .EmployeeTypeStr
        txtPin.Text = .PinNo
        txtNssf.Text = .NssfNo
        txtNhif.Text = .NhifNo
        cboCat.Text = .category.CategoryName
        cboCat.Tag = .category.CategoryID
        txtTel.Text = .HomeTelephone
        txtHAddress.Text = .HomeAddress
        txtEMail.Text = .EMailAddress
        txtDesig.Text = .position.positionName  'Actually a combo box
        txtCert.Text = .GoodConductCertNo
        cboNationality.Text = .Nationality
        cboTribe.Text = .Tribe
        txtBasicPay.Text = .BasicPay
        txtHAllow.Text = .HouseAllowance
        If .IsPhysicallyDisabled Then
            chkDisabled.Value = vbChecked
            txtDisabilityDet.Text = .DisabilityDetails
        Else
            chkDisabled = vbUnchecked
            txtDisabilityDet.Text = ""
        End If
        If .IsOnProbation Then
            chkOnProbation.Value = vbChecked
        Else
            chkOnProbation.Value = vbUnchecked
        End If
        txtProb.Text = .ProbationPeriod
        dtpValidThrough.Value = .EmploymentValidThrough
        txtPassport.Text = .passportNo
        txtAlien.Text = .alienNo
        CboReligion.Text = .Religion
        txtKRAFileNO.Text = .KRAFileNO
        cboMarritalStat.Text = .MaritalStatusStr
        cboProbType.Text = .ProbationTypeStr

        lblSDate.Visible = False
        dtpPSDate.Visible = False
        lblCDate.Visible = False
        dtpCDate.Visible = False
                                
        If .ProbationType = Appointment Then
            lblSDate.Visible = False
            dtpPSDate.Visible = False
            lblCDate.Visible = True
            dtpCDate.Visible = True
           
            dtpCDate.Value = .ConfirmationDate
        
        ElseIf .ProbationType = Promotion Then
            lblSDate.Visible = True
            dtpPSDate.Visible = True
            lblCDate.Visible = True
            dtpCDate.Visible = True
            'txtProbationReason.Text = Trim(!probationReason & "")
            'txtProbationReason.Visible = True
            'lblProbReason.Visible = True
            
            dtpPSDate.Value = .ProbationStartDate
            dtpCDate = .ConfirmationDate
        End If
        
        If .IsDisengaged Then
            chkTerm.Value = vbChecked
            dtpTerm.Value = .DateOfDisengagement
            cboTermReasons.Text = .DisengagementReason
            If cboTermReasons = "Retirement" Then
                fraTerm.Visible = True
                chkTermTrain.Value = setCheckBoxes(.IsTrainedOnRetirement)
                dtpTerminalDate.Value = Format(.TrainingDate, Dfmt)
                txtAdvisor.Text = .TrainingAdvisor
                chkAchieved.Value = setCheckBoxes(.TrainingAchieved)
            End If
        Else
            chkTerm.Value = vbUnchecked
            fraTerm.Visible = False
        End If
                
        'chkPension.Value = setCheckBoxes(!Pension & "")
        
        Set Picture1 = Nothing

        On Error Resume Next 'this handler is specific to the photos only
        Picture1.Picture = LoadPicture(App.Path & "\Photos\" & txtEmpCode.Text & ".jpg")
        If Picture1.Picture = 0 Then
            On Error Resume Next
            Picture1.Picture = LoadPicture(App.Path & "\Pic\Gen.jpg")
        End If
        
    End With
End If
    fraTerm.Visible = True
    fraTerm.Enabled = True
    Exit Sub
errorHandler:
    MsgBox Err.Description, vbExclamation
End Sub

Function setCheckBoxes(b As String) As Integer
    If b = "True" Then
        setCheckBoxes = 1
    Else
        setCheckBoxes = 0
    End If
End Function

Public Sub Cleartxt()
Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Then
            i.Text = ""
        ElseIf TypeOf i Is ComboBox Then
            i.ListIndex = -1
        End If
    Next i
End Sub

Public Sub EnableCmd()
Dim i As Object
    For Each i In Me
        If TypeOf i Is CommandButton Then
            i.Enabled = True
        End If
    Next i

    
End Sub

Public Sub DisableCmd()
Dim i As Object
    For Each i In Me
        If TypeOf i Is CommandButton Then
            i.Enabled = False
        End If
    Next i

    
End Sub

Public Sub disabletxt()
Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
            i.Locked = True
        End If
    Next i
    
    For Each i In Me
        If TypeOf i Is DTPicker Or TypeOf i Is CheckBox Then
            i.Enabled = False
        End If
    Next i
End Sub

Public Sub enabletxt()
Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Or TypeOf i Is ComboBox Or TypeOf i Is MaskEdBox Then
            i.Locked = False
        End If
    Next i
    
        
    For Each i In Me
        If TypeOf i Is DTPicker Or TypeOf i Is CheckBox Then
            i.Enabled = True
        End If
    Next i
    
End Sub

Public Sub InitGrid()
With lvwEmp
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "Employee Code", 1700
    .ColumnHeaders.Add , , "Surname", 2000
    .ColumnHeaders.Add , , "Other Names", 3500
    .ColumnHeaders.Add , , "ID No", 2500
    .ColumnHeaders.Add , , "Gender"
    .ColumnHeaders.Add , , "Date of Birth", 2000
    .ColumnHeaders.Add , , "Date Employed", 2000
'    .ColumnHeaders.Add , , "Tel No", 2500
'    .ColumnHeaders.Add , , "Home Address", 4500
'    .ColumnHeaders.Add , , "E-Mail", 3000
'    .ColumnHeaders.Add , , "Previous Employer", 4000
    .ColumnHeaders.Add , , "Division Code", 1700
    .ColumnHeaders.Add , , "Division Name", 4000
    .ColumnHeaders.Add , , "Terms"
    .ColumnHeaders.Add , , "Employee Type", 2500
    .ColumnHeaders.Add , , "PIN No"
    .ColumnHeaders.Add , , "N.S.S.F No"
    .ColumnHeaders.Add , , "N.H.I.F No"
    .ColumnHeaders.Add , , "KRA File No"

    
    .View = lvwReport
    
End With

End Sub

Public Sub LoadEmployees()

    Dim i As Integer
    Dim j As Long
    Dim itemX As ListItem
    Dim Emp As HRCORE.Employee
    
    On Error GoTo errorHandler
    
    lvwEmp.ListItems.Clear

    For j = 1 To AllEmployees.Count
        Set Emp = AllEmployees.Item(j)
        Set itemX = Me.lvwEmp.ListItems.Add(, , Emp.empcode, , i)
        itemX.SubItems(1) = Emp.SurName
        itemX.SubItems(2) = Emp.OtherNames
        itemX.SubItems(3) = Emp.IdNo
        itemX.SubItems(4) = Emp.DateOfBirth
        itemX.SubItems(5) = Emp.DateOfEmployment
        itemX.SubItems(6) = Emp.OrganizationUnit.OrganizationUnitCode
        itemX.SubItems(7) = Emp.OrganizationUnit.OrganizationUnitName
        itemX.SubItems(8) = Emp.EmploymentTerm.EmpTermName
        itemX.SubItems(9) = Emp.EmployeeType
        itemX.SubItems(10) = Emp.PinNo
        itemX.SubItems(11) = Emp.NssfNo
        itemX.Tag = Emp.EmployeeID
    Next j
    
    Exit Sub
errorHandler:
    MsgBox "An error has occurred while populating employees" & vbNewLine & Err.Description, vbInformation, TITLES
    
End Sub

Public Sub LoadCbo()
    'By Oscar Handled on Form_Load by LoadOUs
    
'    cboCCode.Clear
'    cboCName.Clear
'
'    If frmMain2.cboStructure.Tag = "" Then
'        Set rs3 = CConnect.GetRecordSet("SELECT * FROM MyDivisions ORDER BY Code")
'    Else
'        Set rs3 = CConnect.GetRecordSet("SELECT * FROM MyDivisions WHERE SCode = '" & frmMain2.cboStructure.Tag & "' ORDER BY Code")
'    End If

'With rs3
'    If .RecordCount > 0 Then
'        .MoveFirst
'        Do While Not .EOF
'            cboCCode.AddItem (!code & "")
'            cboCName.AddItem (!Description & "")
'
'            .MoveNext
'        Loop
'    End If
'End With
'
'Set rs3 = Nothing

'Set rs3 = cConnect.GetRecordSet("SELECT NormalEmp, AgricEmp, GroupIdent FROM GeneralOpt")
'
'With rs3
'    If .RecordCount > 0 Then
'        If Not IsNull(!NormalEmp) Then cboType.AddItem (!NormalEmp & "")
'        If Not IsNull(!AgricEmp) Then cboType.AddItem (!AgricEmp & "")
'        If Not IsNull(!GroupIdent) Then cboType.AddItem (!GroupIdent & "")
'
'    End If
'End With
'
'Set rs3 = Nothing

End Sub

Private Sub Form_Resize()
oSmart.FResize Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs = Nothing
    frmMain2.Caption = "Personnel Director " & App.FileDescription
    
End Sub





Private Sub Image1_Click()

End Sub

Private Sub imgDeletePic_Click()
Call cmdPDelete_Click
End Sub

Private Sub imgLoadPic_Click()
Call cmdPNew_Click
End Sub

Private Sub lvwEmp_DblClick()
If frmMain2.cmdEdit.Enabled = True And frmMain2.fracmd.Visible = True Then
    Call frmMain2.cmdEdit_Click
End If
    
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
If SaveNew = False Then
    Select Case ButtonMenu.Key
        Case "Sum"
            FraList.Visible = True
        Case "Det"
            FraList.Visible = False
            disabletxt
            
    End Select
End If
    
End Sub

Private Sub txtBankBranchName_Change()
If frmSelBankBranches.lvwDetails.ListItems.Count > 0 Then txtBankBranchName.Tag = frmSelBankBranches.lvwDetails.SelectedItem.Tag
End Sub



Private Sub txtBankBranchName_KeyPress(KeyAscii As Integer)
With txtBankBranchName
    .Text = ""
    .Tag = ""
End With
KeyAscii = 0
End Sub

Private Sub txtBankName_KeyPress(KeyAscii As Integer)
With txtBankName
    .Text = ""
    .Tag = ""
End With
txtBankCode.Text = ""
KeyAscii = 0
End Sub



Private Sub txtBasicPay_KeyPress(KeyAscii As Integer)
If Len(Trim(txtBasicPay.Text)) > 20 Then
    Beep
    MsgBox "Can't enter more than 20 characters", vbExclamation
    KeyAscii = 8
End If
Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub

Private Sub txtBasicPay_LostFocus()
On Error Resume Next
Dim rsCheckSal As New ADODB.Recordset
Set rsCheckSal = CConnect.GetRecordSet("select * from ecategory where ((lowestsalary<=" & FormatNumber(txtBasicPay.Text, 0, vbFalse, vbFalse, vbFalse) & ") and (highestsalary>=" & FormatNumber(txtBasicPay.Text, 0, vbFalse, vbFalse, vbFalse) & "))")
If rsCheckSal.RecordCount > 0 Then
    cboCat.Text = rsCheckSal!code & ""
    cboCat.Tag = Trim(rsCheckSal!ecategory_id & "")
Else
    cboCat.Text = ""
    cboCat.Tag = ""
End If
txtBasicPay.Text = Format(txtBasicPay.Text, Cfmt)
End Sub


Private Sub txtCBonus_KeyPress(KeyAscii As Integer)
If Len(Trim(txtCBonus.Text)) > 9 Then
    Beep
    MsgBox "Can't enter more than 9 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub

Private Sub txtCBonus_LostFocus()
txtCBonus.Text = Format(txtCBonus.Text, Cfmt)
End Sub

Private Sub txtDEmployed_KeyPress(KeyAscii As Integer)
'    KeyAscii = 0
'txtDEmployed.Text = ""
KeyAscii = 0
End Sub


Private Sub txtDesig_KeyPress(KeyAscii As Integer)
If Len(Trim(txtDesig.Text)) > 199 Then
    Beep
    MsgBox "Can't enter more than 200 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("A") To Asc("Z")
  Case Asc("a") To Asc("z")
  Case Asc(" ")
  Case Asc("/")
  Case Asc("\")
  Case Asc("?")
  Case Asc(":")
  Case Asc(";")
  Case Asc(",")
  Case Asc("-")
  Case Asc("(")
  Case Asc(")")
  Case Asc("&")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
txtDesig.Text = ""
KeyAscii = 0
End Sub

Private Sub txtDOB_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtDOB_LostFocus()
'    On Error GoTo errHandler
'    If IsDate(txtDOB.Text) Then
'        dtpDOB = CDate(txtDOB.Text)
'    Else
'        MsgBox "The date entered is invalid.", vbInformation
'        txtDOB.SetFocus
'    End If
'    Exit Sub
'errHandler:
'    MsgBox Err.Description, vbInformation
'    txtDOB.SetFocus
End Sub

Private Sub txtEmpCode_Change()
    txtEmpCode.Text = UCase(txtEmpCode.Text)
    txtEmpCode.SelStart = Len(txtEmpCode.Text)
End Sub

Private Sub txtEmpCode_KeyPress(KeyAscii As Integer)
If Len(Trim(txtEmpCode.Text)) > 20 Then
    Beep
    MsgBox "Can't enter more than 20 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("A") To Asc("Z")
  Case Asc("a") To Asc("z")
  Case Asc("/")
  Case Asc("\")
  Case Asc("?")
  Case Asc(":")
  Case Asc(";")
  Case Asc(",")
  Case Asc("-")
  Case Asc("(")
  Case Asc(")")
  Case Asc("&")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub




Private Sub txtGrossPay_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub txtGrossPay_LostFocus()
txtGrossPay.Text = Format(txtGrossPay.Text, Cfmt)
End Sub

Private Sub txtHAddress_KeyPress(KeyAscii As Integer)
If Len(Trim(txtHAddress.Text)) > 200 Then
    Beep
    MsgBox "Can't enter more than 200 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("A") To Asc("Z")
  Case Asc("a") To Asc("z")
  Case Asc(" ")
  Case Asc("/")
  Case Asc("\")
  Case Asc("?")
  Case Asc(":")
  Case Asc(";")
  Case Asc(",")
  Case Asc("-")
  Case Asc("(")
  Case Asc(")")
  Case Asc("&")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select

End Sub

Private Sub txtHAllow_KeyPress(KeyAscii As Integer)
If Len(Trim(txtHAllow.Text)) > 9 Then
    Beep
    MsgBox "Can't enter more than 9 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select

End Sub

Private Sub txtHAllow_LostFocus()
txtHAllow.Text = Format(txtHAllow.Text, Cfmt)
End Sub

Private Sub txtIDNo_Change()
    txtIDNo.Text = UCase(txtIDNo.Text)
    txtIDNo.SelStart = Len(txtIDNo.Text)
End Sub

Private Sub txtIDNo_KeyPress(KeyAscii As Integer)
If Len(Trim(txtIDNo.Text)) > 50 Then
    Beep
    MsgBox "Can't enter more than 50 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("A") To Asc("Z")
  Case Asc("a") To Asc("z")
  Case Asc(" ")
  Case Asc("-")
  Case Asc("/")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub

Private Sub txtLAllow_KeyPress(KeyAscii As Integer)
If Len(Trim(txtLAllow.Text)) > 9 Then
    Beep
    MsgBox "Can't enter more than 9 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select

End Sub



Private Sub txtLAllow_LostFocus()
txtLAllow.Text = Format(txtLAllow.Text, Cfmt)
End Sub

Private Sub txtNhif_Change()
    txtNhif.Text = UCase(txtNhif.Text)
    txtNhif.SelStart = Len(txtNhif.Text)
End Sub

Private Sub txtNhif_KeyPress(KeyAscii As Integer)
If Len(Trim(txtNhif.Text)) > 20 Then
    Beep
    MsgBox "Can't enter more than 20 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("A") To Asc("Z")
  Case Asc("a") To Asc("z")
  Case Asc(" ")
  Case Asc("-")
  Case Asc("/")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub

Private Sub txtNssf_Change()
    txtNssf.Text = UCase(txtNssf.Text)
    txtNssf.SelStart = Len(txtNssf.Text)
End Sub

Private Sub txtNssf_KeyPress(KeyAscii As Integer)
If Len(Trim(txtNssf.Text)) > 20 Then
    Beep
    MsgBox "Can't enter more than 20 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("A") To Asc("Z")
  Case Asc("a") To Asc("z")
  Case Asc(" ")
  Case Asc("-")
  Case Asc("/")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub



Private Sub txtOAllow_KeyPress(KeyAscii As Integer)
If Len(Trim(txtOAllow.Text)) > 9 Then
    Beep
    MsgBox "Can't enter more than 9 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select

End Sub

Private Sub txtOAllow_LostFocus()
txtOAllow.Text = Format(txtOAllow.Text, Cfmt)
End Sub

Private Sub txtONames_KeyPress(KeyAscii As Integer)
If Len(Trim(txtONames.Text)) > 100 Then
    Beep
    MsgBox "Can't enter more than 100 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("A") To Asc("Z")
  Case Asc("a") To Asc("z")
  Case Asc(" ")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub



Private Sub txtPin_Change()
    txtPin.Text = UCase(txtPin.Text)
    txtPin.SelStart = Len(txtPin.Text)
End Sub

Private Sub txtPin_KeyPress(KeyAscii As Integer)
If Len(Trim(txtPin.Text)) > 20 Then
    Beep
    MsgBox "Can't enter more than 20 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("A") To Asc("Z")
  Case Asc("a") To Asc("z")
  Case Asc(" ")
  Case Asc("-")
  Case Asc("/")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub


Private Sub txtProb_Change()
    If cboProbType.Text = "Appointment" Then
        dtpCDate.Value = DateAdd("m", Val(txtProb.Text), dtpDEmployed.Value)
    ElseIf cboProbType.Text = "Promotion" Then
        dtpCDate.Value = DateAdd("m", Val(txtProb.Text), dtpPSDate.Value)
        
    End If
    
End Sub

Private Sub txtProb_KeyPress(KeyAscii As Integer)
If Len(Trim(txtProb.Text)) > 9 Then
    Beep
    MsgBox "Can't enter more than 9 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub

Private Sub txtRBonus_KeyPress(KeyAscii As Integer)
If Len(Trim(txtRBonus.Text)) > 9 Then
    Beep
    MsgBox "Can't enter more than 9 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub

Private Sub txtRBonus_LostFocus()
txtRBonus.Text = Format(txtRBonus.Text, Cfmt)
End Sub

Private Sub txtSurName_KeyPress(KeyAscii As Integer)
If Len(Trim(txtSurname.Text)) > 200 Then
    Beep
    MsgBox "Can't enter more than 200 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("A") To Asc("Z")
  Case Asc("a") To Asc("z")
  Case Asc(" ")
  Case Asc(".")
  Case Asc("/")
  Case Asc("-")
  Case Asc("_")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub

Private Sub GenID()
Dim NewID As String
LastSecID = 0
GenerateID = False

Set rs1 = CConnect.GetRecordSet("SELECT GenID, IDInitials, StartFrom, LastSecID FROM GeneralOpt")

With rs1
    If .RecordCount > 0 Then
        If Not IsNull(!GenID) Then
            If !GenID = "Yes" Then
                GenerateID = True
                If IsNull(!IDInitials) Then
                    If Not IsNull(!LastSecID) Then
                        NewID = !LastSecID + 1
                        LastSecID = !LastSecID + 1
                    Else
                        NewID = 0
                        LastSecID = 0
                    End If
                    
                Else
                    If Not IsNull(!LastSecID) Then
                        NewID = !IDInitials & "" & !LastSecID + 1
                        LastSecID = !LastSecID + 1
                    Else
                        NewID = !IDInitials & "" & 0
                        LastSecID = 0
                    End If
                    
                End If
                
                txtEmpCode.Text = NewID
                txtEmpCode.Locked = True
                txtSurname.SetFocus
            Else
                txtEmpCode.Text = ""
                txtEmpCode.Locked = False
'                txtEmpCode.SetFocus
            End If
        End If
    End If
End With

Set rs1 = Nothing

End Sub

Private Sub CheckDOB()
Dim ddif As Long

   
    If DateDiff("d", dtpDOB.Value, Date) < 0 Then
        MsgBox "Date of birth cannot be in the future. Enter the correct date.", vbInformation
        dtpDOB.Value = Date
        'txtDOB.Text = Date
        dtpDOB.SetFocus
        Exit Sub
    End If
    
    'If txtDEmployed.Text <> "" Then
    If DateDiff("d", dtpDEmployed.Value, dtpDOB.Value) > 0 Then
        MsgBox "Date birth cannot be greater than date employed. Enter correct dates.", vbInformation
        dtpDOB.Value = Date
        'txtDOB.Text = Date
        dtpDOB.SetFocus
        Exit Sub
    End If
    'End If
        
        
End Sub

Private Sub CheckDEmployed()
Dim ddf As Long
    
'    If DateDiff("d", dtpDEmployed.Value, Date) < 0 Then
'        MsgBox "Date of employment cannot be in the future. Enter the correct date.", vbInformation
'        dtpDEmployed.Value = Date
'        txtDEmployed.Text = Date
'        dtpDEmployed.SetFocus
'        Exit Sub
'    End If
    
'    If txtDOB.Text <> "" Then
'        dtpDOB.Value = txtDOB.Text
        If DateDiff("d", dtpDEmployed.Value, dtpDOB.Value) > 0 Then
            MsgBox "Date birth cannot be greater than date employed. Enter correct dates.", vbInformation
            dtpDEmployed.Value = Date
            'txtDEmployed.Text = Date
            dtpDEmployed.SetFocus
            Exit Sub
        End If
        
'    End If
    
'    If DateDiff("d", dtpDOB.Value, dtpDEmployed.Value) < 3600 Then
'        MsgBox "The employee could not have been employed when less than 10 years old. Enter correct dates.", vbInformation
'        dtpDEmployed.Value = Date
'        dtpDEmployed.SetFocus
'        Exit Sub
'    End If
     
End Sub

Private Sub CColor()
Dim i As Object

Me.BackColor = MyColor

For Each i In Me
    If TypeOf i Is Label Or TypeOf i Is Form Or TypeOf i Is Frame Then
        i.BackColor = MyColor
    End If
Next i

End Sub

Public Sub CancelMain()
If DSource = "Local" Then
    With frmMain2
        .cmdNew.Enabled = False
        .cmdDelete.Enabled = False
        .cmdEdit.Enabled = False
        .cmdSave.Enabled = True
        .cmdCancel.Enabled = True
    End With
Else
    With frmMain2
        .cmdEdit4.Enabled = False
        .cmdSave4.Enabled = True
        .cmdCancel4.Enabled = True
    End With

End If
End Sub


Private Sub txtTAllow_KeyPress(KeyAscii As Integer)
If Len(Trim(txtTAllow.Text)) > 9 Then
    Beep
    MsgBox "Can't enter more than 9 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select

End Sub

Private Sub txtTAllow_LostFocus()
txtTAllow.Text = Format(txtTAllow.Text, Cfmt)
End Sub

Private Sub txtTel_KeyPress(KeyAscii As Integer)
If Len(Trim(txtTel.Text)) > 49 Then
    Beep
    MsgBox "Can't enter more than 50 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("A") To Asc("Z")
  Case Asc("a") To Asc("z")
  Case Asc(" ")
  Case Asc("/")
  Case Asc("\")
  Case Asc("?")
  Case Asc(":")
  Case Asc(";")
  Case Asc(",")
  Case Asc("-")
  Case Asc("(")
  Case Asc(")")
  Case Asc("&")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select

End Sub

Public Function validateEmployeeData() As Boolean
    validateEmployeeData = False
    If txtEmpCode.Text = "" Then
        MsgBox "Enter Employee code", vbInformation
        txtEmpCode.SetFocus
        Exit Function
    End If
    
    If txtSurname.Text = "" Then
        MsgBox "Enter Employee's SurName", vbInformation
        txtSurname.SetFocus
        Exit Function
    End If
    
    If cboGender.Text = "" Then
        MsgBox "Enter Employee's gender", vbInformation
        cboGender.SetFocus
        Exit Function
    End If
    
    If cboTerms.Text = "" Then
        MsgBox "Enter Employee terms of employment", vbInformation
        cboTerms.SetFocus
        Exit Function
    End If
    
    If cboType.Text = "" Then
        MsgBox "Enter Employee type", vbInformation
        cboType.SetFocus
        Exit Function
    End If
    
    If cboCat.Text = "" Then
        MsgBox "Enter Employee category.", vbInformation
        cboCat.SetFocus
        Exit Function
    End If
    
    '+++++++++++++++++++++++++++++++++++++++++++++++++++
'    If CboReligion.Text = "" Then
'        MsgBox "Enter Employee Religion.", vbInformation
'        CboReligion.SetFocus
'        Exit Function
'    End If
    
'    '// Let the bank Details be optional
'    If txtBankName.Text = "" Or txtBankBranchName.Text = "" Or txtAccountNO.Text = "" Then
'        If MsgBox("Bank details are missing. Do you want to continue anyway?", vbInformation + vbYesNo, "Bank Details") = vbNo Then
'            If txtBankCode.Text = "" Then
'                txtBankCode.SetFocus
'            ElseIf txtBankBranch.Text = "" Then
'                MsgBox "Enter Bank Branch. Use the search button to select the bank.", vbInformation
'                txtBankBranch.SetFocus
'            ElseIf txtAccountNO.Text = "" Then
'                MsgBox "Enter Employee AccountNO.", vbInformation
'                txtAccountNO.SetFocus
'            End If
'            Exit Function
'        End If
'    End If
'
    '+++++++++++++++++++++++++++++++++++++++++++++++++++
            
    If txtNhif.Text = "" Or txtPin.Text = "" Or txtNssf.Text = "" Or txtKRAFileNO.Text = "" Then
        If MsgBox("Your statutory numbers information is incomplete. Do you wish to continue?", vbInformation + vbYesNo) = vbNo Then
            If txtNhif.Text = "" Then
                txtNhif.SetFocus
            ElseIf txtPin.Text = "" Then
                txtPin.SetFocus
            ElseIf txtNssf.Text = "" Then
                txtNssf.SetFocus
            ElseIf txtKRAFileNO.Text = "" Then
                txtKRAFileNO.SetFocus
            End If
            Exit Function
        Else
            If txtNhif.Text = "" Then txtNhif.Text = "0"
            If txtPin.Text = "" Then txtPin.Text = "0"
            If txtNssf.Text = "" Then txtNssf.Text = "0"
            If txtKRAFileNO.Text = "" Then txtKRAFileNO.Text = "0"
        End If
    End If
            
    If txtCert.Text = "" Then
        If MsgBox("Certificate of good conduct No. missing. Are you sure you want to continue?", vbInformation + vbYesNo) = vbNo Then
            txtCert.SetFocus
            Exit Function
        End If
    End If
    
    If txtBasicPay.Text = "" Then
        txtBasicPay.Text = 0
    End If
        
'    If txtLAllow.Text = "" Then
'        txtLAllow.Text = 0
'     End If
    
    If txtHAllow.Text = "" Then
        txtHAllow.Text = 0
    End If
    
'    If txtOAllow.Text = "" Then
'        txtOAllow.Text = 0
'    End If
    
'    If txtTAllow.Text = "" Then
'        txtTAllow.Text = 0
'    End If
    
'    If txtRBonus.Text = "" Then
'        txtRBonus.Text = 0
'    End If
'
'    If txtCBonus.Text = "" Then
'        txtCBonus.Text = 0
'    End If
    
    If txtProb.Text = "" Then
        txtProb.Text = 0
    End If
    
'    If txtDEmployed.Text <> "" Then
'        dtpDEmployed.Value = txtDEmployed.Text
'    End If
    
'    If txtDOB.Text <> "" Then
'        dtpDOB.Value = txtDOB.Text
'    End If
    
    If cboProbType = "Promotion" Then
        PromptDate = DateAdd("m", Val(txtProb.Text), dtpPSDate.Value)
    Else
        PromptDate = DateAdd("m", Val(txtProb.Text), dtpDEmployed.Value)
    End If
    validateEmployeeData = True
End Function

Private Sub txtValidThrough_KeyPress(KeyAscii As Integer)
'txtValidThrough.Text = ""
KeyAscii = 0
End Sub




'=========== HRCORE CODE ===========
Private Sub cboOU_Click()
    Dim theOU As HRCORE.OrganizationUnit
    Dim colReps As New HRCORE.OrganizationUnits
    txtOUInfo.Text = ""
    If cboOU.ListIndex >= 0 Then
        Set theOU = OUnits.FindOrganizationUnit(cboOU.ItemData(cboOU.ListIndex))
        If Not (theOU Is Nothing) Then
            Me.txtOUInfo.Text = "OU Type: " & theOU.OUType.OUTypeName & vbNewLine
            If (theOU.ParentOU Is Nothing) Or (theOU.ParentOU.OrganizationUnitID <= 0) Then
                Me.txtOUInfo.Text = Me.txtOUInfo.Text & "Parent OU: " & company.CompanyName
            Else
                Me.txtOUInfo.Text = Me.txtOUInfo.Text & "Parent OU: " & theOU.ParentOU.OrganizationUnitName
            End If
            Me.txtOUInfo.Text = Me.txtOUInfo.Text & vbNewLine & "Hierarchy: " & LCase(OUnits.GetOUFamilyTree(theOU))
            'check whether the selected ou has replicas
            Set colReps = OUnits.FindOUReplicas(theOU)
            If colReps.Count > 0 Then
                'uncheck and enable the check box
                Me.chkHasOUV.Value = vbUnchecked
                Me.chkHasOUV.Enabled = True
                Me.lvwOU.ListItems.Clear
                'populate the listview with data
                LoadOUReplicas colReps
            Else
                'uncheck and disable the checkbox
                Me.chkHasOUV.Value = vbUnchecked
                Me.chkHasOUV.Enabled = False
                Me.lvwOU.ListItems.Clear
            End If
        End If
    End If
End Sub

Private Sub LoadOUReplicas(ByVal TheReplicas As OrganizationUnits)
    Dim itemX As ListItem
    Dim i As Long
    Dim par As OrganizationUnit
        
    For i = 1 To TheReplicas.Count
        Set par = TheReplicas.Item(i).ParentOU
        'display data of the parents
        Set itemX = lvwOU.ListItems.Add(, , par.OrganizationUnitName)
        itemX.SubItems(1) = LCase(OUnits.GetOUFamilyTree(par))
        'store the ID of the Replica
        itemX.Tag = TheReplicas.Item(i).OrganizationUnitID
    Next i
End Sub


Private Sub chkHasOUV_Click()
    If chkHasOUV.Value = vbChecked Then
        Me.fraEOUV.Enabled = True
    Else
        Me.fraEOUV.Enabled = False
    End If
End Sub

Private Sub InitializeHRCOREObjects()
    Set outs = New HRCORE.OrganizationUnitTypes
    Set OUnits = New HRCORE.OrganizationUnits
    Set empCats = New HRCORE.EmployeeCategories
    Set empTerms = New HRCORE.EmploymentTerms
    Set empNationalities = New HRCORE.Nationalities
    Set empTribes = New HRCORE.Tribes
    Set empReligions = New HRCORE.Religions
    
    
    
    'Set emps = New HRCORE.Employees
    company.LoadCompanyDetails
    
    LoadOUTypes
    LoadOrganizationUnits
    LoadEmployeeCategories
    LoadEmploymentTerms
    LoadNationalities
    LoadTribes
    LoadReligions
    'emps.GetAllEmployees
    
    'PopulateEmployees emps
End Sub

Private Sub LoadOUTypes()
    Dim myOUT As HRCORE.OrganizationUnitType
    outs.GetAllOUTypes
          
End Sub


Private Sub LoadOrganizationUnits()
    Dim myOU As HRCORE.OrganizationUnit
   
    Dim i As Long
    
    OUnits.getallorganizationunits
    cboOU.Clear
    For i = 1 To OUnits.Count
        Set myOU = OUnits.Item(i)
        
        'Force the OU Type details to be loaded
        Set myOU.OUType = outs.FindOUType(myOU.OUType.OUTypeID)
        
        'Force ParentOU to be loaded
        Set myOU.ParentOU = OUnits.FindOrganizationUnit(myOU.ParentOU.OrganizationUnitID)
                        
        cboOU.AddItem myOU.OrganizationUnitName
        cboOU.ItemData(cboOU.NewIndex) = myOU.OrganizationUnitID
    Next i
        
End Sub


Private Sub LoadEmployeeCategories()
    Dim i As Long
    
    On Error GoTo errorHandler
    
    cboCat.Clear
    empCats.GetAllEmployeeCategories
    For i = 1 To empCats.Count
        cboCat.AddItem empCats.Item(i).CategoryName
        cboCat.ItemData(cboCat.NewIndex) = empCats.Item(i).CategoryID
    Next i
    
    Exit Sub
    
errorHandler:
    MsgBox "An error has occurred while Populating Employee Categories" & vbNewLine & Err.Description, vbInformation, TITLES
End Sub

Private Sub LoadEmploymentTerms()
    Dim i As Long
    
    On Error GoTo errorHandler
    
    cboTerms.Clear
    empTerms.GetAllEmploymentTerms
    
    For i = 1 To empTerms.Count
        cboTerms.AddItem empTerms.Item(i).EmpTermName
        cboTerms.ItemData(cboTerms.NewIndex) = empTerms.Item(i).EmpTermID
    Next i
    
    Exit Sub
    
errorHandler:
    MsgBox "An error has occurred while populating Employment Terms" & vbNewLine & Err.Description, vbInformation, TITLES
    
End Sub

Private Sub LoadNationalities()
    Dim i As Long
    
    On Error GoTo errorHandler
    
    cboNationality.Clear
    
    empNationalities.GetAllNationalities
    
    For i = 1 To empNationalities.Count
        cboNationality.AddItem empNationalities.Item(i).Nationality
        cboNationality.ItemData(cboNationality.NewIndex) = empNationalities.Item(i).NationalityID
    Next i
    
    Exit Sub
    
errorHandler:
    MsgBox "An error occurred while populating the Nationalities" & vbNewLine & Err.Description, vbInformation, TITLES
    
End Sub

Private Sub LoadTribes()
    Dim i As Long
    
    On Error GoTo errorHandler
    
    cboTribe.Clear
    
    empTribes.GetAllTribes
    For i = 1 To empTribes.Count
        cboTribe.AddItem empTribes.Item(i).Tribe
        cboTribe.ItemData(cboTribe.NewIndex) = empTribes.Item(i).TribeID
    Next i
    
    Exit Sub
    
errorHandler:
    MsgBox "An error has occurred while populating Tribes" & vbNewLine & Err.Description, vbInformation, TITLES
    
End Sub

Private Sub LoadReligions()
    Dim i As Long
    
    On Error GoTo errorHandler
    
    CboReligion.Clear
    
    empReligions.GetAllReligions
    
    For i = 1 To empReligions.Count
        CboReligion.AddItem empReligions.Item(i).Religion
        CboReligion.ItemData(CboReligion.NewIndex) = empReligions.Item(i).ReligionID
    Next i
    
    Exit Sub
    
errorHandler:
    MsgBox "An error has occurred while populating Religions" & vbNewLine & Err.Description, vbInformation, TITLES
    
End Sub

Private Sub InsertEmployee()
    Dim myEmp As HRCORE.Employee
    Dim Myeouv As HRCORE.EmployeeOUVisibility
    Dim retval As Long
    Dim i As Long
    
    Set myEmp = New HRCORE.Employee
    myEmp.empcode = Me.txtEmpCode.Text
    If Me.chkHasOUV.Value = vbChecked Then
        myEmp.HasOUVisibility = True
    Else
        myEmp.HasOUVisibility = False
    End If
    
    myEmp.SurName = Me.txtSurname.Text
    myEmp.OtherNames = Me.txtOtherNames.Text
    myEmp.VisibleInTheseOUs.Clear
    If cboOU.ListIndex >= 0 Then
        Set myEmp.OrganizationUnit = OUnits.FindOrganizationUnit(CLng(Me.cboOU.ItemData(cboOU.ListIndex)))
        
        If chkHasOUV.Value = vbChecked Then
            myEmp.HasOUVisibility = True
            If lvwOU.ListItems.Count > 0 Then
                For i = 1 To lvwOU.ListItems.Count
                    If lvwOU.ListItems(i).Checked = True Then
                        Set Myeouv = New HRCORE.EmployeeOUVisibility
                        Set Myeouv.Employee = myEmp
                        Set Myeouv.OrganizationUnit = OUnits.FindOrganizationUnit(CLng(lvwOU.ListItems(i).Tag))
                        myEmp.VisibleInTheseOUs.Add Myeouv
                    End If
                Next i
            End If
        Else
            myEmp.HasOUVisibility = False
        End If
    Else
        Set myEmp.OrganizationUnit = Nothing
    End If
    
    retval = myEmp.InsertNew()
End Sub

'========== END OF HRCORE CODE ===========

