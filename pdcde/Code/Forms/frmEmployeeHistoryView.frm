VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEmployeeHistoryView 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   Caption         =   "Employee History Details"
   ClientHeight    =   7260
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7260
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraEdit 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7260
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7440
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   120
         TabIndex        =   100
         Top             =   1920
         Width           =   7215
         Begin TabDlg.SSTab SSTab2 
            Height          =   1875
            Left            =   0
            TabIndex        =   101
            Top             =   120
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   3307
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
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
            TabPicture(0)   =   "frmEmployeeHistoryView.frx":0000
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Frame8"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "BANK DETAILS"
            TabPicture(1)   =   "frmEmployeeHistoryView.frx":001C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame9"
            Tab(1).ControlCount=   1
            Begin VB.Frame Frame8 
               Height          =   1540
               Left            =   120
               TabIndex        =   113
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
                  ItemData        =   "frmEmployeeHistoryView.frx":0038
                  Left            =   960
                  List            =   "frmEmployeeHistoryView.frx":004B
                  TabIndex        =   121
                  Top             =   1155
                  Width           =   1530
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
                  Left            =   3840
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   120
                  Top             =   515
                  Width           =   1530
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
                  Left            =   3840
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   119
                  Top             =   180
                  Width           =   1530
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
                  ItemData        =   "frmEmployeeHistoryView.frx":0080
                  Left            =   3840
                  List            =   "frmEmployeeHistoryView.frx":0082
                  TabIndex        =   118
                  Top             =   1155
                  Width           =   1530
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
                  ItemData        =   "frmEmployeeHistoryView.frx":0084
                  Left            =   3840
                  List            =   "frmEmployeeHistoryView.frx":0086
                  TabIndex        =   117
                  Top             =   850
                  Width           =   1530
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
                  Left            =   960
                  TabIndex        =   116
                  Top             =   515
                  Width           =   1530
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
                  Left            =   960
                  TabIndex        =   115
                  Top             =   840
                  Width           =   1530
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
                  ItemData        =   "frmEmployeeHistoryView.frx":0088
                  Left            =   960
                  List            =   "frmEmployeeHistoryView.frx":008A
                  Sorted          =   -1  'True
                  TabIndex        =   114
                  Top             =   180
                  Width           =   1530
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
                  Left            =   2930
                  TabIndex        =   129
                  Top             =   840
                  Width           =   615
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
                  TabIndex        =   128
                  Top             =   180
                  Width           =   915
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
                  TabIndex        =   127
                  Top             =   1170
                  Width           =   1035
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
                  Left            =   2930
                  TabIndex        =   126
                  Top             =   480
                  Width           =   1095
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
                  TabIndex        =   125
                  Top             =   885
                  Width           =   420
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
                  TabIndex        =   124
                  Top             =   500
                  Width           =   810
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
                  Left            =   2930
                  TabIndex        =   123
                  Top             =   180
                  Width           =   1260
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
                  Left            =   2930
                  TabIndex        =   122
                  Top             =   1200
                  Width           =   420
               End
            End
            Begin VB.Frame Frame9 
               Height          =   1400
               Left            =   -74880
               TabIndex        =   102
               Top             =   300
               Width           =   6975
               Begin VB.CommandButton cmdSchBank 
                  Height          =   285
                  Left            =   4080
                  Picture         =   "frmEmployeeHistoryView.frx":008C
                  Style           =   1  'Graphical
                  TabIndex        =   107
                  Top             =   240
                  Width           =   315
               End
               Begin VB.CommandButton cmdSchBankBranch 
                  Height          =   285
                  Left            =   4080
                  Picture         =   "frmEmployeeHistoryView.frx":0416
                  Style           =   1  'Graphical
                  TabIndex        =   106
                  Top             =   650
                  Width           =   315
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
                  Left            =   1440
                  TabIndex        =   105
                  Top             =   1020
                  Width           =   2565
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
                  Left            =   1440
                  Locked          =   -1  'True
                  TabIndex        =   104
                  Top             =   240
                  Width           =   2565
               End
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
                  Left            =   1440
                  Locked          =   -1  'True
                  TabIndex        =   103
                  Top             =   650
                  Width           =   2565
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
                  TabIndex        =   108
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
                  TabIndex        =   109
                  Text            =   "BankBranch"
                  Top             =   650
                  Width           =   915
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
                  TabIndex        =   112
                  Top             =   240
                  Width           =   855
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
                  TabIndex        =   111
                  Top             =   675
                  Width           =   1395
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
                  TabIndex        =   110
                  Top             =   1020
                  Width           =   945
               End
            End
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   120
         TabIndex        =   29
         Top             =   3960
         Width           =   7215
         Begin TabDlg.SSTab SSTab1 
            Height          =   2655
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   4683
            _Version        =   393216
            Tab             =   2
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
            TabPicture(0)   =   "frmEmployeeHistoryView.frx":07A0
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Frame4"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "EMPLOYEE PAYMENT DETAILS"
            TabPicture(1)   =   "frmEmployeeHistoryView.frx":07BC
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "fraSalary"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "DISENGAGEMENT DETAILS"
            TabPicture(2)   =   "frmEmployeeHistoryView.frx":07D8
            Tab(2).ControlEnabled=   -1  'True
            Tab(2).Control(0)=   "Frame6"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).ControlCount=   1
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
               Height          =   2160
               Left            =   -74880
               TabIndex        =   70
               Top             =   360
               Width           =   6975
               Begin VB.ComboBox cboCCode 
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
                  Left            =   1065
                  Style           =   1  'Simple Combo
                  TabIndex        =   80
                  Top             =   465
                  Width           =   405
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
                  ItemData        =   "frmEmployeeHistoryView.frx":07F4
                  Left            =   5895
                  List            =   "frmEmployeeHistoryView.frx":07F6
                  TabIndex        =   79
                  Top             =   480
                  Width           =   990
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
                  ItemData        =   "frmEmployeeHistoryView.frx":07F8
                  Left            =   3885
                  List            =   "frmEmployeeHistoryView.frx":0802
                  TabIndex        =   78
                  Top             =   480
                  Width           =   940
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
                  ItemData        =   "frmEmployeeHistoryView.frx":081E
                  Left            =   5895
                  List            =   "frmEmployeeHistoryView.frx":0834
                  TabIndex        =   77
                  Top             =   120
                  Width           =   990
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
                  Height          =   435
                  Left            =   3600
                  TabIndex        =   76
                  Top             =   720
                  Width           =   1935
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
                  Left            =   1065
                  TabIndex        =   75
                  Top             =   825
                  Width           =   1965
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
                  ItemData        =   "frmEmployeeHistoryView.frx":0878
                  Left            =   1350
                  List            =   "frmEmployeeHistoryView.frx":0885
                  TabIndex        =   74
                  Top             =   1200
                  Width           =   1680
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
                  Left            =   1065
                  TabIndex        =   73
                  Top             =   1200
                  Width           =   270
               End
               Begin VB.TextBox txtProbationReason 
                  Height          =   315
                  Left            =   1065
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   72
                  Top             =   1560
                  Visible         =   0   'False
                  Width           =   1965
               End
               Begin VB.ComboBox cboCName 
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
                  Left            =   1425
                  TabIndex        =   71
                  Top             =   465
                  Width           =   1605
               End
               Begin MSComCtl2.DTPicker dtpDEmployed 
                  Height          =   315
                  Left            =   1065
                  TabIndex        =   81
                  Top             =   120
                  Width           =   1125
                  _ExtentX        =   1984
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
                  Format          =   56295427
                  CurrentDate     =   37845
               End
               Begin MSComCtl2.DTPicker dtpValidThrough 
                  Height          =   285
                  Left            =   3885
                  TabIndex        =   82
                  Top             =   135
                  Width           =   940
                  _ExtentX        =   1667
                  _ExtentY        =   503
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
                  Format          =   56295427
                  CurrentDate     =   37845
               End
               Begin MSComCtl2.DTPicker dtpPSDate 
                  Height          =   300
                  Left            =   3885
                  TabIndex        =   85
                  Top             =   1200
                  Visible         =   0   'False
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   529
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
                  Format          =   56295427
                  CurrentDate     =   37845
               End
               Begin MSComCtl2.DTPicker dtpCDate 
                  Height          =   300
                  Left            =   5895
                  TabIndex        =   86
                  Top             =   1200
                  Visible         =   0   'False
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   529
                  _Version        =   393216
                  Enabled         =   0   'False
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
                  Format          =   56295427
                  CurrentDate     =   37845
               End
               Begin MSComCtl2.DTPicker dtpSPension 
                  Height          =   300
                  Left            =   5880
                  TabIndex        =   87
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   1035
                  _ExtentX        =   1826
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
                  CustomFormat    =   "yyyy-MM-dd"
                  Format          =   56295427
                  CurrentDate     =   37845
               End
               Begin VB.CheckBox chkUnsolicited 
                  Appearance      =   0  'Flat
                  Caption         =   "Unsolicited"
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
                  Height          =   210
                  Left            =   5880
                  TabIndex        =   84
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   960
               End
               Begin VB.CheckBox chkPension 
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  Caption         =   "Pensionable"
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
                  Height          =   195
                  Left            =   5895
                  TabIndex        =   83
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   1080
               End
               Begin VB.Label Label41 
                  Appearance      =   0  'Flat
                  Caption         =   "Employment valid through:"
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
                  Left            =   2475
                  TabIndex        =   99
                  Top             =   120
                  Width           =   1920
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
                  TabIndex        =   98
                  Top             =   120
                  Width           =   1230
               End
               Begin VB.Label Label21 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Department Name:"
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
                  Left            =   1515
                  TabIndex        =   97
                  Top             =   480
                  Width           =   1365
               End
               Begin VB.Label Label13 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "*Department:"
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
                  TabIndex        =   96
                  Top             =   480
                  Width           =   1005
               End
               Begin VB.Label Label10 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "*Grades:"
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
                  Left            =   4920
                  TabIndex        =   95
                  Top             =   480
                  Width           =   660
               End
               Begin VB.Label Label19 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "*Type:"
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
                  Left            =   3315
                  TabIndex        =   94
                  Top             =   480
                  Width           =   510
               End
               Begin VB.Label Label18 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "*Terms:"
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
                  Left            =   4920
                  TabIndex        =   93
                  Top             =   120
                  Width           =   585
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
                  TabIndex        =   92
                  Top             =   840
                  Width           =   900
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
                  Left            =   3315
                  TabIndex        =   91
                  Top             =   1200
                  Width           =   810
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
                  Left            =   120
                  TabIndex        =   90
                  Top             =   1200
                  Width           =   1230
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
                  Left            =   4920
                  TabIndex        =   89
                  Top             =   1200
                  Width           =   1365
               End
               Begin VB.Label lblProbReason 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Probation Reason:"
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
                  TabIndex        =   88
                  Top             =   1600
                  Visible         =   0   'False
                  Width           =   1335
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
               TabIndex        =   45
               Top             =   360
               Width           =   6855
               Begin VB.Frame fraStateNumbers 
                  BorderStyle     =   0  'None
                  Height          =   1935
                  Left            =   120
                  TabIndex        =   46
                  Top             =   120
                  Width           =   6615
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
                     Left            =   3750
                     TabIndex        =   62
                     Top             =   570
                     Width           =   1530
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
                     Left            =   3750
                     TabIndex        =   61
                     Top             =   210
                     Width           =   1530
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
                     Left            =   970
                     TabIndex        =   60
                     Top             =   580
                     Width           =   1530
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
                     Left            =   970
                     TabIndex        =   59
                     Top             =   210
                     Width           =   1530
                  End
                  Begin VB.Frame fraSal 
                     BorderStyle     =   0  'None
                     Height          =   495
                     Left            =   5880
                     TabIndex        =   50
                     Top             =   1560
                     Width           =   5055
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
                        TabIndex        =   54
                        Top             =   210
                        Visible         =   0   'False
                        Width           =   1140
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
                        TabIndex        =   53
                        Top             =   210
                        Visible         =   0   'False
                        Width           =   1155
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
                        TabIndex        =   52
                        Top             =   90
                        Visible         =   0   'False
                        Width           =   765
                     End
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
                        TabIndex        =   51
                        Top             =   210
                        Visible         =   0   'False
                        Width           =   1110
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
                        TabIndex        =   58
                        Top             =   0
                        Visible         =   0   'False
                        Width           =   1185
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
                        TabIndex        =   57
                        Top             =   0
                        Visible         =   0   'False
                        Width           =   720
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
                        TabIndex        =   56
                        Top             =   120
                        Visible         =   0   'False
                        Width           =   780
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
                        TabIndex        =   55
                        Top             =   0
                        Visible         =   0   'False
                        Width           =   915
                     End
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
                     Left            =   970
                     TabIndex        =   49
                     Top             =   950
                     Width           =   1530
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
                     Left            =   3750
                     TabIndex        =   48
                     Top             =   960
                     Width           =   1530
                  End
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
                     Left            =   970
                     TabIndex        =   47
                     Top             =   1320
                     Width           =   1530
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
                     TabIndex        =   69
                     Top             =   960
                     Width           =   1380
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
                     Left            =   3225
                     TabIndex        =   68
                     Top             =   600
                     Width           =   720
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
                     Left            =   3225
                     TabIndex        =   67
                     Top             =   240
                     Width           =   735
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
                     TabIndex        =   66
                     Top             =   600
                     Width           =   555
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
                     TabIndex        =   65
                     Top             =   240
                     Width           =   945
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
                     TabIndex        =   64
                     Top             =   1350
                     Width           =   1275
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
                     Left            =   3225
                     TabIndex        =   63
                     Top             =   990
                     Width           =   735
                  End
               End
            End
            Begin VB.Frame Frame6 
               Height          =   2175
               Left            =   120
               TabIndex        =   31
               Top             =   360
               Width           =   6975
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
                  TabIndex        =   38
                  Top             =   120
                  Width           =   3495
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
                     TabIndex        =   41
                     Top             =   1560
                     Width           =   2490
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
                     TabIndex        =   40
                     Top             =   1170
                     Width           =   1920
                  End
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
                     TabIndex        =   39
                     Top             =   330
                     Width           =   1950
                  End
                  Begin MSComCtl2.DTPicker dtpTerminalDate 
                     Height          =   330
                     Left            =   1200
                     TabIndex        =   42
                     Top             =   780
                     Width           =   1320
                     _ExtentX        =   2328
                     _ExtentY        =   582
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
                     Format          =   56295427
                     CurrentDate     =   37845
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
                     TabIndex        =   44
                     Top             =   1170
                     Width           =   540
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
                     TabIndex        =   43
                     Top             =   840
                     Width           =   1020
                  End
               End
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
                  TabIndex        =   32
                  Top             =   120
                  Width           =   3015
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
                     ItemData        =   "frmEmployeeHistoryView.frx":08A3
                     Left            =   75
                     List            =   "frmEmployeeHistoryView.frx":08BF
                     TabIndex        =   35
                     Top             =   1080
                     Width           =   2820
                  End
                  Begin VB.CheckBox chkReEngage 
                     Appearance      =   0  'Flat
                     Caption         =   "Cannot be re-engaged"
                     ForeColor       =   &H80000008&
                     Height          =   255
                     Left            =   75
                     TabIndex        =   34
                     Top             =   1560
                     Visible         =   0   'False
                     Width           =   1935
                  End
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
                     TabIndex        =   33
                     Top             =   720
                     Width           =   1185
                  End
                  Begin MSComCtl2.DTPicker dtpTerm 
                     Height          =   330
                     Left            =   1680
                     TabIndex        =   36
                     Top             =   360
                     Width           =   1215
                     _ExtentX        =   2143
                     _ExtentY        =   582
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
                     Format          =   56295427
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
                     TabIndex        =   37
                     Top             =   360
                     Width           =   1560
                  End
               End
            End
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
         TabIndex        =   10
         Top             =   120
         Width           =   7215
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
            ItemData        =   "frmEmployeeHistoryView.frx":0919
            Left            =   910
            List            =   "frmEmployeeHistoryView.frx":0926
            TabIndex        =   19
            Text            =   "Unspecified"
            Top             =   1170
            Width           =   1485
         End
         Begin VB.CommandButton cmdPNew 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4680
            Picture         =   "frmEmployeeHistoryView.frx":0945
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   1200
            Visible         =   0   'False
            Width           =   320
         End
         Begin VB.CommandButton cmdPDelete 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5040
            Picture         =   "frmEmployeeHistoryView.frx":0A47
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   1200
            Visible         =   0   'False
            Width           =   320
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
            Left            =   3480
            TabIndex        =   16
            Top             =   525
            Width           =   1485
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
            Left            =   910
            TabIndex        =   15
            Top             =   210
            Width           =   1485
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
            Left            =   3480
            TabIndex        =   14
            Top             =   210
            Width           =   1485
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
            Left            =   910
            TabIndex        =   13
            Top             =   530
            Width           =   1485
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
            Left            =   910
            TabIndex        =   12
            Top             =   850
            Width           =   1485
         End
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
            Left            =   3480
            TabIndex        =   11
            Top             =   847
            Width           =   1485
         End
         Begin MSComCtl2.DTPicker dtpDOB 
            Height          =   315
            Left            =   3480
            TabIndex        =   20
            Top             =   1170
            Width           =   1485
            _ExtentX        =   2619
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
            OLEDropMode     =   1
            CustomFormat    =   "dd, MMM, yyyy"
            Format          =   56295427
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
                  Picture         =   "frmEmployeeHistoryView.frx":0F39
                  Key             =   "Search"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEmployeeHistoryView.frx":104B
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEmployeeHistoryView.frx":115D
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEmployeeHistoryView.frx":126F
                  Key             =   ""
               EndProperty
            EndProperty
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
            Left            =   2640
            TabIndex        =   28
            Top             =   885
            Width           =   1185
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
            TabIndex        =   27
            Top             =   870
            Width           =   1080
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
            TabIndex        =   26
            Top             =   570
            Width           =   1095
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
            Left            =   2640
            TabIndex        =   25
            Top             =   240
            Width           =   780
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
            TabIndex        =   24
            Top             =   255
            Width           =   810
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
            TabIndex        =   23
            Top             =   1185
            Width           =   675
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
            Left            =   2640
            TabIndex        =   22
            Top             =   1200
            Width           =   1095
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
            Left            =   2640
            TabIndex        =   21
            Top             =   555
            Width           =   615
         End
         Begin VB.Image Picture1 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   1335
            Left            =   5400
            Stretch         =   -1  'True
            Top             =   120
            Width           =   1425
         End
         Begin VB.Image imgLoadPic 
            Appearance      =   0  'Flat
            Height          =   180
            Left            =   5400
            MouseIcon       =   "frmEmployeeHistoryView.frx":17B1
            MousePointer    =   99  'Custom
            Picture         =   "frmEmployeeHistoryView.frx":1BF3
            Stretch         =   -1  'True
            ToolTipText     =   "Click this icon to ADD a new photo"
            Top             =   1530
            Visible         =   0   'False
            Width           =   200
         End
         Begin VB.Image imgDeletePic 
            Appearance      =   0  'Flat
            Height          =   180
            Left            =   6625
            MouseIcon       =   "frmEmployeeHistoryView.frx":1D3D
            MousePointer    =   99  'Custom
            Picture         =   "frmEmployeeHistoryView.frx":217F
            Stretch         =   -1  'True
            ToolTipText     =   "Click this icon to DELETE employee photo"
            Top             =   1530
            Visible         =   0   'False
            Width           =   200
         End
      End
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   6600
         Width           =   7215
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
            Picture         =   "frmEmployeeHistoryView.frx":25C1
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Cancel Process"
            Top             =   120
            Width           =   495
         End
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
            Picture         =   "frmEmployeeHistoryView.frx":26C3
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Save Record"
            Top             =   120
            Width           =   495
         End
      End
      Begin MSComDlg.CommonDialog Cdl 
         Left            =   6720
         Top             =   5760
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CheckBox chkPayroll 
         Appearance      =   0  'Flat
         Caption         =   "Appear in Payroll"
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
         Height          =   390
         Left            =   120
         TabIndex        =   131
         Top             =   6600
         Visible         =   0   'False
         Width           =   1365
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
         TabIndex        =   132
         Top             =   5280
         Visible         =   0   'False
         Width           =   1035
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
         TabIndex        =   130
         Top             =   5295
         Visible         =   0   'False
         Width           =   1320
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
         TabIndex        =   134
         Top             =   5280
         Visible         =   0   'False
         Width           =   1455
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
         TabIndex        =   133
         Top             =   5280
         Width           =   435
      End
      Begin VB.Line Line2 
         BorderWidth     =   5
         X1              =   0
         X2              =   7080
         Y1              =   -120
         Y2              =   -120
      End
   End
   Begin VB.Frame FraList 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   7020
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   7440
      Begin MSComctlLib.ListView lvwEmp 
         Height          =   6720
         Left            =   0
         TabIndex        =   1
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
         Picture         =   "frmEmployeeHistoryView.frx":27C5
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmEmployeeHistoryView.frx":28C7
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "frmEmployeeHistoryView.frx":29C9
         Style           =   1  'Graphical
         TabIndex        =   3
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
               Picture         =   "frmEmployeeHistoryView.frx":2EBB
               Key             =   "A"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeHistoryView.frx":330D
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeHistoryView.frx":3627
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeHistoryView.frx":3A79
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeHistoryView.frx":3ECB
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeHistoryView.frx":431D
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeHistoryView.frx":4637
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeHistoryView.frx":4951
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeHistoryView.frx":4DA3
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeHistoryView.frx":50BD
               Key             =   "B"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeHistoryView.frx":550F
               Key             =   "C"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeHistoryView.frx":5961
               Key             =   "D"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeHistoryView.frx":5DB3
               Key             =   "E"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeHistoryView.frx":6205
               Key             =   "F"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeHistoryView.frx":6657
               Key             =   "G"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeHistoryView.frx":6AA9
               Key             =   "H"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeHistoryView.frx":6EFB
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeHistoryView.frx":734D
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeHistoryView.frx":779F
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeHistoryView.frx":7BF1
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeHistoryView.frx":8043
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeHistoryView.frx":8495
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployeeHistoryView.frx":88E7
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
         TabIndex        =   2
         Top             =   6420
         Visible         =   0   'False
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmEmployeeHistoryView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NewAcct As Boolean
Dim LastSecID As Long
Dim GenerateID As Boolean
Dim EnterDOB As Boolean
Dim EnterDEmp As Boolean
Dim MStruc As String

Private Sub cboCat_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboCCode_Click()
    If Not cboCCode.Text = "" Then
        Set rs3 = CConnect.GetRecordSet("SELECT * FROM MyDivisions WHERE Code = '" & cboCCode.Text & "'")
        
        With rs3
            If .RecordCount > 0 Then
                cboCName.Text = !Description & ""
              
            End If
        End With
        
        Set rs3 = Nothing
    
    End If
End Sub

Private Sub cboCCode_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboCName_Click()
    If Not cboCName.Text = "" Then
        Set rs3 = CConnect.GetRecordSet("SELECT * FROM MyDivisions WHERE Description = '" & cboCName.Text & "'")
        
        With rs3
            If .RecordCount > 0 Then
                cboCCode.Text = !code & ""
                              
            End If
        End With
        
        Set rs3 = Nothing
    
    End If

End Sub

Private Sub cboCName_KeyPress(KeyAscii As Integer)
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
    KeyAscii = 0
End Sub



Private Sub cboNationality_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

    


Private Sub cboProbType_Click()
    If cboProbType.Text = "Appointment" Then
        lblSDate.Visible = False
        dtpPSDate.Visible = False
        lblCDate.Visible = True
        dtpCDate.Visible = True
        dtpCDate.Value = DateAdd("m", Val(txtProb.Text), dtpDEmployed.Value)
        
    ElseIf cboProbType.Text = "Promotion" Then
        lblSDate.Visible = True
        dtpPSDate.Visible = True
        lblCDate.Visible = True
        dtpCDate.Visible = True
        dtpCDate.Value = DateAdd("m", Val(txtProb.Text), dtpPSDate.Value)
        
    Else
        lblSDate.Visible = False
        dtpPSDate.Visible = False
        lblCDate.Visible = False
        dtpCDate.Visible = False
        
    End If
                
    
End Sub

Private Sub cboProbType_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboTermReasons_Click()
    If chkTerm.Value = 1 Then
        If cboTermReasons.Text = "Retirement" Then
            fraTerm.Visible = True
        Else
            fraTerm.Visible = False
        End If
    End If

End Sub

Private Sub cboTermReasons_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboTerms_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboTribe_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboType_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub chkTerm_Click()
    If chkTerm.Value = 1 Then
        dtpTerm.Value = Date
    Else
        fraTerm.Visible = False
    End If
    
End Sub

Public Sub cmdCancel_Click()
Unload Me
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
''    Call DisplayRecords
'
End Sub

Public Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub cmdDelete_Click()
Dim Emp As String
    
    Omnis_ActionTag = "D" 'Deletes a new record in the Omnis database 'monte++
    
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
        
        CConnect.ExecuteSql ("DELETE FROM Employee Where employee_id = '" & Emp & "'")
        CConnect.ExecuteSql ("DELETE FROM SEmp Where employee_id = '" & Emp & "'")
            
'        ' if AuditTrail = True Then cConnect.ExecuteSql ("INSERT INTO AuditTrail (UserId, DTime, Trans, TDesc, MySection)VALUES('" & CurrentUser & "','" & Date & " " & Time & "','Deleting Employee','" & frmMain2.lvwEmp.SelectedItem & "','Employee')")
        
        rs.Requery
        
        Set rs5 = CConnect.GetRecordSet("SELECT * FROM Security WHERE UID = '" & CurrentUser & "' AND subsystem = '" & SubSystem)

        With rs5
            If Not .EOF And Not .BOF Then
                Set rsGlob = Nothing
                
                If Not IsNull(!terms) And Not IsNull(!LCode) Then
                    Set rsGlob = CConnect.GetRecordSet("SELECT e.*, c.Code, c.Description FROM (Employee as e " & _
                            "LEFT JOIN SEmp as s ON e.employee_id = s.employee_id) LEFT JOIN CStructure as c ON s.LCode" & _
                            "= c.LCode LEFT JOIN ECategory as  ec ON e.ECategory = ec.code " & _
                            " WHERE SEmp.LCode like '" & !LCode & "%" & "' AND Employee.Terms = '" & !terms & "' AND " & _
                            "ec.seq >= '" & maxCatAccess & "' ORDER BY e.EmpCode")
                                    
                ElseIf Not IsNull(!terms) And IsNull(!LCode) Then
                    Set rsGlob = CConnect.GetRecordSet("SELECT e.*, c.Code, c.Description FROM (Employee as e " & _
                            "LEFT JOIN SEmp as s ON e.employee_id = s.employee_id) LEFT JOIN CStructure as c ON s.LCode" & _
                            "= c.LCode LEFT JOIN ECategory as  ec ON e.ECategory = ec.code " & _
                            " WHERE Employee.Terms = '" & !terms & "' AND ec.seq >= '" & maxCatAccess & "' ORDER BY e.EmpCode")
                            
                ElseIf IsNull(!terms) And Not IsNull(!LCode) Then
                    Set rsGlob = CConnect.GetRecordSet("SELECT e.*, c.Code, c.Description FROM (Employee as e " & _
                            "LEFT JOIN SEmp as s ON e.employee_id = s.employee_id) LEFT JOIN CStructure as c ON s.LCode" & _
                            "= c.LCode LEFT JOIN ECategory as  ec ON e.ECategory = ec.code " & _
                            " WHERE SEmp.LCode like '" & !LCode & "%" & "'  AND ec.seq >= '" & maxCatAccess & "' ORDER BY e.EmpCode")
                            
                Else
                    Set rsGlob = CConnect.GetRecordSet("SELECT e.*, c.Code, c.Description FROM (Employee as e " & _
                            "LEFT JOIN SEmp as s ON e.employee_id = s.employee_id) LEFT JOIN CStructure as c ON s.LCode" & _
                            "= c.LCode LEFT JOIN ECategory as  ec ON e.ECategory = ec.code " & _
                            " WHERE ec.seq >= '" & maxCatAccess & "' ORDER BY e.EmpCode")
        
                End If
            
            End If
        End With
                    
        Set rs5 = Nothing
        
        
        Call LoadList
        frmMain2.LoadMyList
        Call DisplayRecords(rsGlob!empcode)
        Call frmMain2.cboTerms_Click

        
        'frmMain2.lblECount.Caption = rsGlob.RecordCount
        
     
                       
    Else
        MsgBox "No records to be deleted.", vbInformation
    End If

End Sub

Public Sub cmdEdit_Click()
    On Error GoTo ErrorTrap
    Omnis_ActionTag = "E" 'Edits a record in the Omnis database 'monte++
    
    If lvwEmp.ListItems.Count > 0 Then
        With rs
            If .RecordCount > 0 Then
                .MoveFirst
                .Find "employee_id like '" & lvwEmp.SelectedItem.Tag & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    Call DisplayRecords(rs!empcode)
                    FraList.Visible = False
                    enabletxt
                    dtpCDate.Enabled = False
                    'DisableCmd
                    cmdSave.Enabled = True
                    cmdCancel.Enabled = True
                    txtEmpCode.Locked = True
                    txtSurName.SetFocus
                    SaveNew = False
                Else
                    PSave = True
                    Call cmdCancel_Click
                    PSave = False
                End If
            Else
                PSave = True
                Call cmdCancel_Click
                PSave = False
            End If
        End With
        
    Else
        MsgBox "No records tobe edited", vbInformation
        PSave = True
        Call cmdCancel_Click
        PSave = False
    End If
    Exit Sub
ErrorTrap:
    MsgBox Err.Description, vbExclamation
End Sub



Public Sub cmdNew_Click()
    
    Omnis_ActionTag = "I" 'Inserts a new record in the Omnis database 'monte++
    
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
    
    dtpDOB.Value = Date
    dtpDEmployed.Value = Date
    EnterDOB = False
    EnterDEmp = False
    
    txtEmpCode.SetFocus
    Call GenID
    dtpTerm.Value = Date
    DisableCmd
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

End Sub

Private Sub cmdPNew_Click()

Dim picturepath As String
Dim MovePic As FileSystemObject


If Len(Trim(txtEmpCode.Text)) <= 0 Then
    MsgBox "Enter Employee Code first", vbInformation, "Picture"
    Exit Sub
Else
    Set MovePic = New FileSystemObject
    
    With cdl
        .Filter = "Pictures {*.bmp;*.gif;*.jpeg;*.ico;*.jpg;*.ICO;*.JPEG;*.JPG;*.BMP;*.GIF|*.bmp;*.gif;*.jpeg;*.ico;*.jpg;*.ICO;*.JPEG;*.JPG;*.BMP;*.GIF"
        .ShowOpen
    End With
    

    picturepath = cdl.FileName
    
End If

If Len(Trim(picturepath)) > 0 Then
    If cdl.FileName Like "*.ico" Or _
        cdl.FileName Like "*.jpeg" Or _
        cdl.FileName Like "*.jpg" Or _
        cdl.FileName Like "*.bmp" Or _
        cdl.FileName Like "*.gif" Or _
        cdl.FileName Like "*.ICO" Or _
        cdl.FileName Like "*.JPEG" Or _
        cdl.FileName Like "*.JPG" Or _
        cdl.FileName Like "*.BMP" Or _
        cdl.FileName Like "*.gif" Then
        
        On Error Resume Next
        Picture1.Picture = LoadPicture(picturepath)
        MovePic.CopyFile picturepath, App.Path & "\Photos\" & txtEmpCode.Text & ".jpg", True
        
    Else
        MsgBox "Unsupported file format", vbExclamation, "Picture"
        On Error Resume Next
        Picture1.Picture = LoadPicture(App.Path & "\Pic\Gen.jpg")
    End If
  
End If
End Sub


Public Sub cmdSave_Click()
Unload Me
'save
End Sub
Public Sub save()
Dim Term As Integer
Dim Pension As String
Dim Unsolicited As String
Dim PromptDate As Date


Unsolicited = 0
If txtBasicPay.Text = "" Then
    txtBasicPay.Text = 0
End If

If txtEmpCode.Text = "" Then
    MsgBox "Enter Employee code", vbInformation
    txtEmpCode.SetFocus
    Call CancelMain
    Exit Sub
End If

If txtSurName.Text = "" Then
    MsgBox "Enter Employee's SurName", vbInformation
    txtSurName.SetFocus
    Call CancelMain
    Exit Sub
End If

If cboGender.Text = "" Then
    MsgBox "Enter Employee's gender", vbInformation
    cboGender.SetFocus
    Call CancelMain
    Exit Sub
End If

If cboTerms.Text = "" Then
    MsgBox "Enter Employee terms of employment", vbInformation
    cboTerms.SetFocus
    Call CancelMain
    Exit Sub
End If

If cboType.Text = "" Then
    MsgBox "Enter Employee type", vbInformation
    cboType.SetFocus
    Call CancelMain
    Exit Sub
End If

If cboCat.Text = "" Then
    MsgBox "Enter Employee category.", vbInformation
    cboCat.SetFocus
    Call CancelMain
    Exit Sub
End If

'+++++++++++++++++++++++++++++++++++++++++++++++++++
If CboReligion.Text = "" Then
    MsgBox "Enter Employee Religion.", vbInformation
    CboReligion.SetFocus
    Call CancelMain
    Exit Sub
End If

'// Let the bank Details be optional
If txtBankCode.Text = "" Then
    If MsgBox("Bank details are missing. Do you want to continue anyway?", vbInformation + vbYesNo, "Bank Details") = vbNo Then
        txtBankCode.SetFocus
        Call CancelMain
        Exit Sub
    Else
        GoTo AFTERBANKDETAILS
    End If
End If

If txtBankBranch.Text = "" Then
    MsgBox "Enter Bank Branch.", vbInformation
    txtBankBranch.SetFocus
    Call CancelMain
    Exit Sub
End If

If txtAccountNO.Text = "" Then
    MsgBox "Enter Employee AccountNO.", vbInformation
    txtAccountNO.SetFocus
    Call CancelMain
    Exit Sub
End If
AFTERBANKDETAILS:
If txtKRAFileNO.Text = "" Then
    MsgBox "Enter Employee KRA FileNO.", vbInformation
    txtKRAFileNO.SetFocus
    Call CancelMain
    Exit Sub
End If
'+++++++++++++++++++++++++++++++++++++++++++++++++++

If chkPension.Value = 1 Then
    Pension = "Yes"
Else
    Pension = "No"
End If

If chkTerm.Value = 1 Then
    If MsgBox("Make sure all the employee's records are up to date before terminating him. Do you wish to continue?", vbInformation + vbYesNo) = vbNo Then
        Call CancelMain
        Exit Sub
    End If
    Term = 1
Else
    Term = 0
End If

If txtNhif.Text = "" Then
    If MsgBox("N.H.I.F No missing. Are you sure you want to continue?", vbInformation + vbYesNo) = vbNo Then
        txtNhif.SetFocus
        Exit Sub
    End If
End If
        
If txtPin.Text = "" Then
    If MsgBox("P.I.N No missing. Are you sure you want to continue?", vbInformation + vbYesNo) = vbNo Then
        txtPin.SetFocus
        Exit Sub
    End If
End If

If txtNssf.Text = "" Then
    If MsgBox("N.S.S.F No missing. Are you sure you want to continue?", vbInformation + vbYesNo) = vbNo Then
        txtNssf.SetFocus
        Exit Sub
    End If
End If

If txtCert.Text = "" Then
    If MsgBox("Certificate of good conduct No. missing. Are you sure you want to continue?", vbInformation + vbYesNo) = vbNo Then
        txtCert.SetFocus
        Exit Sub
    End If
End If

    

If txtLAllow.Text = "" Then
    txtLAllow.Text = 0
End If

If txtHAllow.Text = "" Then
    txtHAllow.Text = 0
End If

If txtOAllow.Text = "" Then
    txtOAllow.Text = 0
End If

If txtTAllow.Text = "" Then
    txtTAllow.Text = 0
End If

If txtRBonus.Text = "" Then
    txtRBonus.Text = 0
End If

If txtCBonus.Text = "" Then
    txtCBonus.Text = 0
End If

If txtProb.Text = "" Then
    txtProb.Text = 0
End If

'If txtDEmployed.Text <> "" Then
'    dtpDEmployed.Value = txtDEmployed.Text
'End If

If cboProbType = "Promotion" Then
    PromptDate = DateAdd("m", Val(txtProb.Text), dtpPSDate.Value)
Else
    PromptDate = DateAdd("m", Val(txtProb.Text), dtpDEmployed.Value)
End If
 
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
        
        If PromptSave = True Then
            If MsgBox("Are you sure you want to save the record.", vbQuestion + vbYesNo) = vbNo Then
                Call CancelMain
                Exit Sub
            End If
        End If
        
        CConnect.ExecuteSql ("DELETE FROM Employee WHERE EmpCode = '" & txtEmpCode.Text & "'")
                
        'If txtDOB.Text <> "" And txtDEmployed.Text <> "" Then
            mySQL = "INSERT INTO Employee (EmpCode, Surname, OtherNames, IDNo, DOB, DEmployed," & _
                    " BasicPay,DCode, Terms, Type, PinNo, NssfNo, NhifNo, Term, dleft, ECategory, HTel, EMail, HAddress, Desig, Gender, LAllow, HAllow, OAllow, TAllow, TermTrain, TermDate, Advisor, Achieved, CertNo, Nationality, Tribe, Pension, Disabled, Payroll, TermReasons, Unsolicited, RBonus, CBonus, Prob, CDate, SPension, ProbType, PSDate," & _
                    " Religion,BankCode,BankName,BankBranch,BankBranchName,AccountNO,KRAFileNO)" & _
                    " VALUES('" & txtEmpCode.Text & "', '" & txtSurName.Text & "', '" & txtONames.Text & "'," & _
                    " '" & txtIDNo.Text & "', '" & Format(dtpDOB.Value, Dfmt) & "', '" & Format(dtpDEmployed.Value, Dfmt) & "'," & _
                    " " & Format(txtBasicPay.Text, Nfmt) & ",'" & cboCCode.Text & "'," & _
                    " '" & cboTerms.Text & "', '" & cboType.Text & "', '" & txtPin.Text & "', '" & txtNssf.Text & "'," & _
                    " '" & txtNhif.Text & "','" & Term & "','" & Format(dtpTerm.Value, Dfmt) & "','" & cboCat.Text & "','" & txtTel.Text & "','" & txtEMail.Text & "','" & txtHAddress.Text & "','" & txtDesig.Text & "','" & cboGender.Text & "'," & Format(txtLAllow.Text, Nfmt) & "," & Format(txtHAllow.Text, Nfmt) & "," & Format(txtOAllow.Text, Nfmt) & "," & Format(txtTAllow.Text, Nfmt) & ",'" & chkTermTrain.Value & "','" & Format(dtpTerminalDate.Value, Dfmt) & "','" & txtAdvisor.Text & "','" & chkAchieved.Value & "','" & txtCert.Text & "','" & cboNationality.Text & "','" & cboTribe.Text & "','" & Pension & "','" & chkDisabled.Value & "','" & chkPayroll.Value & "','" & cboTermReasons.Text & "','" & Unsolicited & "'," & Format(txtRBonus.Text, Nfmt) & "," & Format(txtCBonus.Text, Nfmt) & "," & txtProb.Text & ",'" & PromptDate & "','" & dtpSPension.Value & "','" & cboProbType.Text & "','" & dtpPSDate.Value & "'," & _
                    " '" & CboReligion.Text & "','" & txtBankCode.Text & "','" & txtBankName.Text & "','" & txtBankBranch.Text & "','" & txtBankBranchName.Text & "','" & txtAccountNO.Text & "','" & txtKRAFileNO.Text & "')"
        
'        ElseIf txtDOB.Text <> "" And txtDEmployed.Text = "" Then
'            mySQL = "INSERT INTO Employee (EmpCode, Surname, OtherNames, IDNo, DOB," & _
'                    " BasicPay,DCode, Terms, Type, PinNo, NssfNo, NhifNo, Term, dleft, ECategory, HTel, EMail, HAddress, Desig, Gender, LAllow, HAllow, OAllow, TAllow, TermTrain, TermDate, Advisor, Achieved, CertNo, Nationality, Tribe, Pension, Disabled, Payroll, TermReasons, Unsolicited, RBonus, CBonus, Prob, CDate, SPension, ProbType, PSDate," & _
'                    " Religion,BankCode,BankName,BankBranch,BankBranchName,AccountNO,KRAFileNO)" & _
'                    " VALUES('" & txtEmpCode.Text & "', '" & txtSurname.Text & "', '" & txtONames.Text & "'," & _
'                    " '" & txtIDNo.Text & "', '" & Format(dtpDOB.Value, Dfmt) & "'," & _
'                    " " & Format(txtBasicPay.Text, Nfmt) & ",'" & cboCCode.Text & "'," & _
'                    " '" & cboTerms.Text & "', '" & cboType.Text & "', '" & txtPin.Text & "', '" & txtNssf.Text & "'," & _
'                    " '" & txtNhif.Text & "','" & Term & "','" & Format(dtpTerm.Value, Dfmt) & "','" & cboCat.Text & "','" & txtTel.Text & "','" & txtEmail.Text & "','" & txtHAddress.Text & "','" & txtDesig.Text & "','" & cboGender.Text & "'," & Format(txtLAllow.Text, Nfmt) & "," & Format(txtHAllow.Text, Nfmt) & "," & Format(txtOAllow.Text, Nfmt) & "," & Format(txtTAllow.Text, Nfmt) & ",'" & chkTermTrain.Value & "','" & Format(dtpTerminalDate.Value, Dfmt) & "','" & txtAdvisor.Text & "','" & chkAchieved.Value & "','" & txtCert.Text & "','" & cboNationality.Text & "','" & cboTribe.Text & "','" & Pension & "','" & chkDisabled.Value & "','" & chkPayroll.Value & "','" & cboTermReasons.Text & "','" & Unsolicited & "'," & Format(txtRBonus.Text, Nfmt) & "," & Format(txtCBonus.Text, Nfmt) & "," & txtProb.Text & ",'" & PromptDate & "','" & dtpSPension.Value & "','" & cboProbType.Text & "','" & dtpPSDate.Value & "'," & _
'                    " '" & CboReligion.Text & "','" & txtBankCode.Text & "','" & txtBankName.Text & "','" & txtBankBranch.Text & "','" & txtBankBranchName.Text & "','" & txtAccountNO.Text & "','" & txtKRAFileNO.Text & "')"
'
'        ElseIf txtDOB.Text = "" And txtDEmployed.Text <> "" Then
'            mySQL = "INSERT INTO Employee (EmpCode, Surname, OtherNames, IDNo, DEmployed," & _
'                " BasicPay,DCode, Terms, Type, PinNo, NssfNo, NhifNo, Term, dleft, ECategory, HTel, EMail, HAddress, Desig, Gender, LAllow, HAllow, OAllow, TAllow, TermTrain, TermDate, Advisor, Achieved, CertNo, Nationality, Tribe, Pension, Disabled, Payroll, TermReasons, Unsolicited, RBonus, CBonus, Prob, CDate, SPension, ProbType, PSDate," & _
'                " Religion,BankCode,BankName,BankBranch,BankBranchName,AccountNO,KRAFileNO)" & _
'                " VALUES('" & txtEmpCode.Text & "', '" & txtSurname.Text & "', '" & txtONames.Text & "'," & _
'                " '" & txtIDNo.Text & "', '" & Format(dtpDEmployed.Value, Dfmt) & "'," & _
'                " " & Format(txtBasicPay.Text, Nfmt) & ",'" & cboCCode.Text & "'," & _
'                " '" & cboTerms.Text & "', '" & cboType.Text & "', '" & txtPin.Text & "', '" & txtNssf.Text & "'," & _
'                " '" & txtNhif.Text & "','" & Term & "','" & Format(dtpTerm.Value, Dfmt) & "','" & cboCat.Text & "','" & txtTel.Text & "','" & txtEmail.Text & "','" & txtHAddress.Text & "','" & txtDesig.Text & "','" & cboGender.Text & "'," & Format(txtLAllow.Text, Nfmt) & "," & Format(txtHAllow.Text, Nfmt) & "," & Format(txtOAllow.Text, Nfmt) & "," & Format(txtTAllow.Text, Nfmt) & ",'" & chkTermTrain.Value & "','" & Format(dtpTerminalDate.Value, Dfmt) & "','" & txtAdvisor.Text & "','" & chkAchieved.Value & "','" & txtCert.Text & "','" & cboNationality.Text & "','" & cboTribe.Text & "','" & Pension & "','" & chkDisabled.Value & "','" & chkPayroll.Value & "','" & cboTermReasons.Text & "','" & Unsolicited & "'," & Format(txtRBonus.Text, Nfmt) & "," & Format(txtCBonus.Text, Nfmt) & "," & txtProb.Text & ",'" & PromptDate & "','" & dtpSPension.Value & "','" & cboProbType.Text & "','" & dtpPSDate.Value & "'," & _
'                " '" & CboReligion.Text & "','" & txtBankCode.Text & "','" & txtBankName.Text & "','" & txtBankBranch.Text & "','" & txtBankBranchName.Text & "','" & txtAccountNO.Text & "','" & txtKRAFileNO.Text & "')"
'
'        Else
'            mySQL = "INSERT INTO Employee (EmpCode, Surname, OtherNames, IDNo," & _
'                " BasicPay,DCode, Terms, Type, PinNo, NssfNo, NhifNo, Term, dleft, ECategory, HTel, EMail, HAddress, Desig, Gender, LAllow, HAllow, OAllow, TAllow, TermTrain, TermDate, Advisor, Achieved, CertNo, Nationality, Tribe, Pension, Disabled, Payroll, TermReasons, Unsolicited, RBonus, CBonus, Prob, CDate, SPension, ProbType, PSDate," & _
'                " Religion,BankCode,BankName,BankBranch,BankBranchName,AccountNO,KRAFileNO)" & _
'                " VALUES('" & txtEmpCode.Text & "', '" & txtSurname.Text & "', '" & txtONames.Text & "'," & _
'                " '" & txtIDNo.Text & "'," & _
'                " " & Format(txtBasicPay.Text, Nfmt) & ",'" & cboCCode.Text & "'," & _
'                " '" & cboTerms.Text & "', '" & cboType.Text & "', '" & txtPin.Text & "', '" & txtNssf.Text & "'," & _
'                " '" & txtNhif.Text & "','" & Term & "','" & Format(dtpTerm.Value, Dfmt) & "','" & cboCat.Text & "','" & txtTel.Text & "','" & txtEmail.Text & "','" & txtHAddress.Text & "','" & txtDesig.Text & "','" & cboGender.Text & "'," & Format(txtLAllow.Text, Nfmt) & "," & Format(txtHAllow.Text, Nfmt) & "," & Format(txtOAllow.Text, Nfmt) & "," & Format(txtTAllow.Text, Nfmt) & ",'" & chkTermTrain.Value & "','" & Format(dtpTerminalDate.Value, Dfmt) & "','" & txtAdvisor.Text & "','" & chkAchieved.Value & "','" & txtCert.Text & "','" & cboNationality.Text & "','" & cboTribe.Text & "','" & Pension & "','" & chkDisabled.Value & "','" & chkPayroll.Value & "','" & cboTermReasons.Text & "','" & Unsolicited & "'," & Format(txtRBonus.Text, Nfmt) & "," & Format(txtCBonus.Text, Nfmt) & "," & txtProb.Text & ",'" & PromptDate & "','" & dtpSPension.Value & "','" & cboProbType.Text & "','" & dtpPSDate.Value & "'," & _
'                " '" & CboReligion.Text & "','" & txtBankCode.Text & "','" & txtBankName.Text & "','" & txtBankBranch.Text & "','" & txtBankBranchName.Text & "','" & txtAccountNO.Text & "','" & txtKRAFileNO.Text & "')"
'
'        End If


        CConnect.ExecuteSql (mySQL)
        
        CConnect.ExecuteSql ("DELETE FROM SEmp WHERE EmpCode = '" & txtEmpCode.Text & "' AND SCode = '" & MStruc & "'")
        
        If cboCCode.Text <> "" Then
            Set rs5 = CConnect.GetRecordSet("SELECT * FROM MyDivisions WHERE Code = '" & cboCCode.Text & "'")
            
            With rs5
                If .RecordCount > 0 Then
                   .MoveFirst
                    CConnect.ExecuteSql ("INSERT INTO SEmp (SCode, LCode, EmpCode) VALUES('" & MStruc & "','" & !RLCode & "','" & txtEmpCode.Text & "')")
                
                End If
            End With
        End If
        
                    
        If GenerateID = True Then
            CConnect.ExecuteSql ("UPDATE GeneralOpt SET LastSecID = " & LastSecID & "")
            GenerateID = False
        End If

    Set rsGlob2 = CConnect.GetRecordSet("SELECT e.*, c.Code, c.Description FROM (Employee as e " & _
            "LEFT JOIN SEmp as s ON e.employee_id = s.employee_id) LEFT JOIN CStructure as c ON s.LCode" & _
            "= c.LCode LEFT JOIN ECategory as  ec ON e.ECategory = ec.code " & _
            " WHERE e.Term <> 1 AND (((s.SCode)='" & MStruc & "')) OR (((s.SCode) Is Null))" & _
            " ec.seq >= '" & maxCatAccess & "' ORDER BY e.EmpCode")
                
rs.Requery

Set rs5 = CConnect.GetRecordSet("SELECT * FROM Security WHERE UID = '" & CurrentUser & "' AND subsystem = '" & SubSystem & "'")

With rs5
    If Not .EOF And Not .BOF Then
        Set rsGlob = Nothing
        
        If Not IsNull(!terms) And Not IsNull(!LCode) Then
            Set rsGlob = CConnect.GetRecordSet("SELECT e.*, c.Code, c.Description FROM (Employee as e " & _
                    "LEFT JOIN SEmp as s ON e.employee_id = s.employee_id) LEFT JOIN CStructure as c ON s.LCode" & _
                    "= c.LCode LEFT JOIN ECategory as  ec ON e.ECategory = ec.code " & _
                    " WHERE s.LCode like '" & !LCode & "%" & "' AND e.Terms = '" & !terms & "' AND " & _
                    " ec.seq >= '" & maxCatAccess & "' ORDER BY e.EmpCode")
                            
        ElseIf Not IsNull(!terms) And IsNull(!LCode) Then
            Set rsGlob = CConnect.GetRecordSet("SELECT e.*, c.Code, c.Description FROM (Employee as e " & _
                    "LEFT JOIN SEmp as s ON e.employee_id = s.employee_id) LEFT JOIN CStructure as c ON s.LCode" & _
                    "= c.LCode LEFT JOIN ECategory as  ec ON e.ECategory = ec.code " & _
                    " WHERE e.Terms = '" & !terms & "' AND  ec.seq >= '" & maxCatAccess & "' ORDER BY e.EmpCode")
                    
        ElseIf IsNull(!terms) And Not IsNull(!LCode) Then
            Set rsGlob = CConnect.GetRecordSet("SELECT e.*, c.Code, c.Description FROM (Employee as e " & _
                    "LEFT JOIN SEmp as s ON e.employee_id = s.employee_id) LEFT JOIN CStructure as c ON s.LCode" & _
                    "= c.LCode LEFT JOIN ECategory as  ec ON e.ECategory = ec.code " & _
                    " WHERE s.LCode like '" & !LCode & "%" & "' AND ec.seq >= '" & maxCatAccess & "' ORDER BY e.EmpCode")
                    
        Else
            Set rsGlob = CConnect.GetRecordSet("SELECT e.*, c.Code, c.Description FROM (Employee as e " & _
                    "LEFT JOIN SEmp as s ON e.employee_id = s.employee_id) LEFT JOIN CStructure as c ON s.LCode" & _
                    "= c.LCode LEFT JOIN ECategory as  ec ON e.ECategory = ec.code " & _
                    " WHERE ec.seq >= '" & maxCatAccess & "' ORDER BY e.EmpCode")

        End If
    
    End If
End With
            
Set rs5 = Nothing


If SaveNew = True Then
    Call LoadList
    frmMain2.LoadMyList
End If

Call frmMain2.cboTerms_Click


If SaveNew = True Then
Dim Posts As Long

    If txtDesig.Text <> "" Then
'''        Set rs5 = CConnect.GetRecordSet("SELECT * FROM Positions WHERE PositionName = '" & txtDesig.Text & "'")
'''
'''        With rs5
'''            If .RecordCount > 0 Then
'''                Posts = !Posts & ""
'''                If !Posts > 1 Then
'''                    CConnect.ExecuteSql ("UPDATE Positions SET Posts = " & rs5!Posts - 1 & " WHERE Code = '" & txtDesig.Text & "'")
'''                Else
'''                    CConnect.ExecuteSql ("UPDATE Positions SET Posts = 0, Approved = 'No' WHERE Code = '" & txtDesig.Text & "'")
'''
'''                End If
'''
'''            End If
'''        End With
'''
'''        Set rs5 = Nothing
    End If
    
        
    
    Call Cleartxt
    txtEmpCode.SetFocus
    Call GenID
Else
    With rsGlob
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "EmpCode like '" & txtEmpCode.Text & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                PSave = True
                Call cmdCancel_Click
                PSave = False
            End If
        End If
    End With
    
End If

Employees_TextFile  '++Caters for the Omnis text file 'Monte++

frmMain2.lblECount.Caption = rsGlob.RecordCount

If cboTermReasons.Text <> "NONE" Then

Dim ctran As New CTransfer
ctran.Transfer_Employee txtEmpCode, , True
    rsGlob.Requery
    rsGlob2.Requery
    
    Call frmMain2.LoadMyList
    MsgBox "Employee Archived successfully.", vbInformation
End If
Unload Me
End Sub




Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdSchBank_Click()
    frmSelBanks.Show vbModal
    SelectedBank = strName
    txtBankCode = strName
    
    If strName <> "" Then
    Dim rsBanks As Recordset
    Set rsBanks = New Recordset
        Set rsBanks = CConnect.GetRecordSet("select * from tblBank where bank_id='" & strName & "'")
        If Not IsNull(rsBanks!bank_Name) Then txtBankName = rsBanks!bank_Name
    Set rsBanks = Nothing
    End If
End Sub

Private Sub cmdSchBankBranch_Click()
    frmSelBankBranches.Show vbModal
    txtBankBranch = strName
    
    If strName <> "" Then
    Dim rsBanks As Recordset
    Set rsBanks = New Recordset
        Set rsBanks = CConnect.GetRecordSet("select * from tblBankBranch where bankBranch_Code='" & strName & "'")
        If Not IsNull(rsBanks!BANKBRANCH_NAME) Then txtBankBranchName = rsBanks!BANKBRANCH_NAME
    Set rsBanks = Nothing
    End If
    
End Sub

Private Sub dtpDEmployed_CloseUp()
    EnterDEmp = True
    Call CheckDEmployed
   ' txtDEmployed.Text = dtpDEmployed.Value
End Sub

Private Sub dtpDEmployed_KeyPress(KeyAscii As Integer)
    EnterDEmp = True
    Call CheckDEmployed
End Sub

Private Sub dtpDOB_CloseUp()
    EnterDOB = True
    Call CheckDOB
    'txtDOB.Text = dtpDOB.Value
    
End Sub

Private Sub dtpDOB_KeyPress(KeyAscii As Integer)
    EnterDOB = True
    Call CheckDOB
End Sub

Private Sub dtpPSDate_Change()
    dtpCDate.Value = DateAdd("m", Val(txtProb.Text), dtpPSDate.Value)
End Sub

Private Sub Form_Load()
'On Error GoTo hell

Decla.Security Me
oSmart.FReset Me

If oSmart.hRatio > 1.1 Then
    With frmMain2
        Me.Move .tvwMain.Width + .lvwEmp.Width + (.tvwMain.Width / 36) * 2, (.Height / 5.52) '- 155
    End With
Else
     With frmMain2
        Me.Move .tvwMain.Width + .lvwEmp.Width + (.tvwMain.Width / 36#) * 2, .Height / 5.55
    End With
    
End If


CConnect.CColor Me, MyColor
'''
''''frmMain2.txtDetails.Caption = ""
'''Call disabletxt
'''
'''Call InitGrid
'''CConnect.CCon
'''
'''Call Loadcbo
'''
'''Set rs5 = CConnect.GetRecordSet("SELECT * FROM STypes WHERE SMain = 1")
'''
'''With rs5
'''    If .RecordCount > 0 Then
'''        .MoveFirst
'''        MStruc = !Code & ""
'''    End If
'''End With
'''
'''Set rs5 = Nothing
'''
'''cboCat.Clear
'''
'''Set rs5 = CConnect.GetRecordSet("SELECT * FROM ECategory ORDER BY Code")
'''
'''With rs5
'''    If .RecordCount > 0 Then
'''        .MoveFirst
'''        Do While Not .EOF
'''            cboCat.AddItem !Code & ""
'''
'''            .MoveNext
'''        Loop
'''    End If
'''End With
'''
'''Set rs5 = Nothing
'''
'''Set rs5 = CConnect.GetRecordSet("SELECT * FROM Positions WHERE Approved = 'Yes' ORDER BY Code")
'''
'''With rs5
'''    If .RecordCount > 0 Then
'''        .MoveFirst
'''        Do While Not .EOF
'''            txtDesig.AddItem !Code & ""
'''
'''            .MoveNext
'''        Loop
'''    End If
'''End With
'''
'''Set rs5 = Nothing
'''
'''Set rs5 = CConnect.GetRecordSet("SELECT * FROM Nationality ORDER BY Code")
'''
'''With rs5
'''    If .RecordCount > 0 Then
'''        .MoveFirst
'''        Do While Not .EOF
'''            cboNationality.AddItem !Code & ""
'''
'''            .MoveNext
'''        Loop
'''    End If
'''End With
'''
'''Set rs5 = Nothing
'''
'''Set rs5 = CConnect.GetRecordSet("SELECT * FROM Tribes ORDER BY Code")
'''With rs5
'''    If .RecordCount > 0 Then
'''        .MoveFirst
'''        Do While Not .EOF
'''            cboTribe.AddItem !Code & ""
'''            .MoveNext
'''        Loop
'''    End If
'''End With
'''Set rs5 = Nothing
'''
'''Set rs5 = CConnect.GetRecordSet("SELECT * FROM Religion ORDER BY Code")
'''With rs5
'''    If .RecordCount > 0 Then
'''        .MoveFirst
'''        Do While Not .EOF
'''            CboReligion.AddItem !Code & ""
'''            .MoveNext
'''        Loop
'''    End If
'''End With
'''Set rs5 = Nothing
'''
'''Set Rs = CConnect.GetRecordSet("SELECT Employee.*, CStructure.Code, CStructure.Description" & _
'''            " FROM (Employee LEFT JOIN SEmp ON Employee.EmpCode = SEmp.EmpCode) LEFT JOIN CStructure ON SEmp.LCode = CStructure.LCode" & _
'''            " WHERE (((SEmp.SCode)='" & MStruc & "')) OR (((SEmp.SCode) Is Null))" & _
'''            " ORDER BY Employee.EmpCode")
'''
'''
'''Call LoadList
''''Call DisplayRecords
'''cmdSave.Enabled = False
'''cmdCancel.Enabled = False
'''If ViewSal = True Then
'''    fraSal.Visible = True
'''Else
'''    fraSal.Visible = False
'''End If
'''
Exit Sub
Hell:
MsgBox Err.Description, vbExclamation, "Employee Record"
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
                txtSurName.Text = !SurName & ""
                txtONames.Text = !OtherNames & ""
                txtIDNo.Text = !IdNo & ""
                cboGender.Text = !Gender & ""
                If Not IsNull(!DOB) Then dtpDOB.Value = !DOB Else dtpDOB.Value = Date
                If Not IsNull(!DEmployed) Then dtpDEmployed.Value = !DEmployed Else dtpDOB.Value = Date
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

Public Sub DisplayRecords(empcode As String)
Dim RsT As New ADODB.Recordset
Dim inc As Integer
    On Error GoTo ErrorTrap
    Set RsT = CConnect.GetRecordSet("SELECT * FROM Employee_History")
    Call Cleartxt

    If RsT.RecordCount = 0 Then
ResetRsT:
        'Set RsT = CConnect.GetRecordSet("SELECT * FROM Employee WHERE Term = 1")
        
        Set RsT = CConnect.deptFilter("SELECT e.*, c.Code, c.Description FROM (Employee as e LEFT JOIN SEmp " & _
        "as s ON e.employee_id = s.employee_id) LEFT JOIN CStructure as c ON s.LCode = c.LCode LEFT JOIN ECategory " & _
        "as ec ON e.ECategory = ec.code WHERE e.Term = 1 AND ec.seq >= '" & maxCatAccess & "' ORDER BY e.EmpCode")
        
        If RsT.RecordCount = 0 Then Exit Sub
    End If
    
    With RsT
        If .RecordCount > 0 Then
            .MoveFirst
            '.Find "EmpCode = '" & rsGlob!empCode & "'", , adSearchForward, adBookmarkFirst
            .Filter = "employee_id=" & empcode
            If .RecordCount = 0 Then
                If inc < 2 Then
                    inc = 2
                    GoTo ResetRsT
                Else
                    Exit Sub
                End If
            End If
            If Not .EOF Then
                txtEmpCode.Text = !empcode & ""
                txtSurName.Text = !SurName & ""
                txtONames.Text = !OtherNames & ""
                txtIDNo.Text = !IdNo & ""
                cboGender.Text = !Gender & ""
                If Not IsNull(!DOB) Then dtpDOB.Value = !DOB Else dtpDOB.Value = Date
                'txtDOB.Text = !DOB & ""
                If Not IsNull(!DEmployed) Then dtpDEmployed.Value = !DEmployed Else dtpDOB.Value = Date
               ' txtDEmployed.Text = !DEmployed & ""
'                cboCCode.Text = !Code & ""
                'cboCName.Text = !Description & ""
                
                '+++++Added 20-07-2006 by Juma, C. O.
                
                cboCCode.Text = !code & ""
                cboCCode.Tag = !cstructure_id & ""
                cboCName.Text = !Description & ""
                
                '+++++End of Addition (20-07-2006) by Juma, C. O.
                
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
                If IsNull(!Disabled) Then
                    chkDisabled.Value = 0
                Else
                    chkDisabled.Value = returnChecked(!Disabled)
                End If
                If IsNull(!Unsolicited) Then
                    chkUnsolicited.Value = 0
                Else
                    chkUnsolicited.Value = returnChecked(!Unsolicited)
                End If
                If IsNull(!Payroll) Then
                    chkPayroll.Value = 0
                Else
                    chkPayroll.Value = returnChecked(!Payroll)
                End If
                
                txtProb.Text = (!Prob & "")
                
                '+++++++++++++++++++++++++++++++++++++++
                CboReligion.Text = !Religion & ""
'                txtBankCode.Text = !BankCode & ""
'                txtBankName.Text = !BankName & ""
'                txtBankBranch.Text = !BankBranch & ""
'                txtBankBranchName.Text = !BankBranchName & ""
'                txtAccountNO.Text = !AccountNO & ""
'                txtKRAFileNO.Text = !KRAFileNO & ""
                
                dtpValidThrough.Value = Format(Trim(!employmentvalidthro & ""), Dfmt)
                txtPassport.Text = Trim(!passport & "")
                txtAlien.Text = Trim(!alienNo & "")
                'dtpValidThrough.Value = Format(Trim(!employmentvalidthro & ""), Dfmt)
                '+++++++++++++++++++++++++++++++++++++++
                Dim rsMM As New ADODB.Recordset
                Dim rsMX As New ADODB.Recordset
                Set rsMM = CConnect.GetRecordSet("select * from EmployeeBanks where employee_id='" & rsGlob!employee_id & "' and mainacct=1")
                If rsMM.RecordCount > 0 Then
                    Set rsMX = CConnect.GetRecordSet("select * from tblBankBranch where bankbranch_ID=" & rsMM!branchID & "")
                    If rsMX.RecordCount > 0 Then
                        Dim rsBankCode As New ADODB.Recordset
                        Set rsBankCode = CConnect.GetRecordSet("select * from tblBank where bank_id=" & rsMX!bank_id & "")
                        If rsBankCode.RecordCount > 0 Then txtBankCode.Text = rsBankCode!bank_code & "": txtBankCode.Tag = rsBankCode!bank_code & "": txtBankName.Text = rsBankCode!bank_Name & "" Else txtBankCode.Text = "": txtBankName.Text = ""
                        
                        txtBankBranch.Text = rsMX!BankBranch_Code & ""
                        txtBankBranchName.Text = rsMX!BANKBRANCH_NAME & "" '!BankBranchName & ""
                        txtBankBranchName.Tag = Trim(rsMM!branchID & "")
                        txtAccountNO.Text = Trim(rsMM!accnumber & "") '!AccountNO & ""
                    End If
                End If
                txtKRAFileNO.Text = IIf(!KRAFileNO & "" = "", "0", !KRAFileNO)
                cboMarritalStat.Text = Trim(!marital_status & "")
                '+++++++++++++++++++++++++++++++++++++++
                
                If !SPension <> "" Then dtpSPension.Value = Format(!SPension, Dfmt)
                
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
                    txtProbationReason.Text = Trim(!probationReason & "")
                    txtProbationReason.Visible = True
                    lblProbReason.Visible = True
                    
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
                    txtProbationReason.Text = Trim(!probationReason & "")
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
                Else
                    txtProbationReason.Visible = False
                    lblProbReason.Visible = False
                End If
                
                If !Term = True Then
                    chkTerm.Value = 1
                    If Not Trim(!dleft & "") = "" Then dtpTerm.Value = !dleft
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
    Exit Sub
ErrorTrap:
    If Err.Number = 91 Then
        RsT.Requery
        Exit Sub
    End If
    
    MsgBox Err.Description, vbExclamation
    Err.Clear
End Sub



Public Sub Cleartxt()
Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
            i.Text = ""
        End If
    Next i

    
End Sub

Public Sub EnableCmd()
'Dim i As Object
'    For Each i In Me
'        If TypeOf i Is CommandButton Then
'            i.Enabled = True
'        End If
'    Next i
'
'
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

Private Sub enabletxt()
Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
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
    .ColumnHeaders.Add , , "L.A.S.C No"

    
    .View = lvwReport
    
End With



End Sub

Public Sub LoadList()
Dim i As Integer

lvwEmp.ListItems.Clear

With rs
    If .RecordCount > 0 Then
        i = 5
        Do While Not .EOF
            Set LI = lvwEmp.ListItems.Add(, , !empcode & "", , i)
            LI.ListSubItems.Add , , !SurName & ""
            LI.ListSubItems.Add , , !OtherNames & ""
            LI.ListSubItems.Add , , !IdNo & ""
            LI.ListSubItems.Add , , !Gender & ""
            LI.ListSubItems.Add , , !DOB & ""
            LI.ListSubItems.Add , , !DEmployed & ""
            LI.ListSubItems.Add , , !code & ""
            LI.ListSubItems.Add , , !Description & ""
            LI.ListSubItems.Add , , !terms & ""
            LI.ListSubItems.Add , , !Type & ""
            LI.ListSubItems.Add , , !PinNo & ""
            LI.ListSubItems.Add , , !NssfNo & ""
            LI.ListSubItems.Add , , !NhifNo & ""
            LI.ListSubItems.Add , , !LascNo & ""

            
            If i = 5 Then
                i = 6
            Else
                i = 5
            End If
            
            .MoveNext
        Loop
        
        .MoveFirst
    End If
End With
End Sub

Public Sub LoadCbo()
cboCCode.Clear
cboCName.Clear


Set rs3 = CConnect.GetRecordSet("SELECT * FROM MyDivisions ORDER BY Code")

With rs3
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
            cboCCode.AddItem (!code & "")
            cboCName.AddItem (!Description & "")
            
            .MoveNext
        Loop
    End If
End With

Set rs3 = Nothing

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

Private Sub txtBasicPay_KeyPress(KeyAscii As Integer)
If Len(Trim(txtBasicPay.Text)) > 9 Then
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

Private Sub txtBasicPay_LostFocus()
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

Private Sub txtDEmployed_KeyPress(KeyAscii As Integer)
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

End Sub

Private Sub txtDOB_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtEMail_KeyPress(KeyAscii As Integer)
If Len(Trim(txtEMail.Text)) > 50 Then
    Beep
    MsgBox "Can't enter more than 50 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("A") To Asc("Z")
  Case Asc("a") To Asc("z")
  Case Asc(" ")
  Case Asc("@")
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

Private Sub txtSurName_KeyPress(KeyAscii As Integer)
If Len(Trim(txtSurName.Text)) > 200 Then
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
                txtSurName.SetFocus
            Else
                txtEmpCode.Text = ""
                txtEmpCode.Locked = False
                txtEmpCode.SetFocus
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
    If DateDiff("d", dtpDEmployed.Value, dtpDOB.Value) > 0 Then
        MsgBox "Date birth cannot be greater than date employed. Enter correct dates.", vbInformation
        dtpDEmployed.Value = Date
'        txtDEmployed.Text = Date
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

Function returnChecked(bool As Boolean) As Integer
    If bool = True Then
        returnChecked = 1
    ElseIf bool = False Then
        returnChecked = 0
    End If
End Function
