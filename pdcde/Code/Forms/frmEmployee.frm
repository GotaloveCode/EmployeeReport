VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmEmployee 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Details"
   ClientHeight    =   8205
   ClientLeft      =   1515
   ClientTop       =   1065
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   7995
   Begin VB.Frame FraEdit 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8100
      Left            =   0
      TabIndex        =   62
      Top             =   0
      Width           =   7850
      Begin TabDlg.SSTab sstEmployees 
         Height          =   6945
         Left            =   120
         TabIndex        =   71
         Top             =   120
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   12250
         _Version        =   393216
         Style           =   1
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   520
         WordWrap        =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "General Details"
         TabPicture(0)   =   "frmEmployee.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblDisengaged"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame5"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Employment Details"
         TabPicture(1)   =   "frmEmployee.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame4"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Department Info"
         TabPicture(2)   =   "frmEmployee.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label45"
         Tab(2).Control(1)=   "chkHasOUV"
         Tab(2).Control(2)=   "txtOUInfo"
         Tab(2).Control(3)=   "cboOU"
         Tab(2).Control(4)=   "cmdSearch"
         Tab(2).Control(5)=   "cmdSearchEOUV"
         Tab(2).Control(6)=   "fraEOUV"
         Tab(2).Control(7)=   "fraOtherDetails"
         Tab(2).ControlCount=   8
         TabCaption(3)   =   "Job Description"
         TabPicture(3)   =   "frmEmployee.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame6"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Next Of Kin info"
         TabPicture(4)   =   "frmEmployee.frx":0070
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "fraNextOfKin"
         Tab(4).Control(1)=   "fraExistingNextOfKin"
         Tab(4).Control(2)=   "Frame2"
         Tab(4).ControlCount=   3
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   615
            Left            =   -74880
            TabIndex        =   167
            Top             =   6240
            Width           =   7095
            Begin VB.CommandButton cmdNoKAdd 
               Caption         =   "Add"
               Enabled         =   0   'False
               Height          =   495
               Left            =   0
               TabIndex        =   170
               Top             =   0
               Width           =   1215
            End
            Begin VB.CommandButton cmdNOKEdit 
               Caption         =   "Edit"
               Enabled         =   0   'False
               Height          =   495
               Left            =   1440
               TabIndex        =   169
               Top             =   0
               Width           =   1215
            End
            Begin VB.CommandButton cmdNoKDelete 
               Caption         =   "Remove"
               Enabled         =   0   'False
               Height          =   495
               Left            =   2880
               TabIndex        =   168
               Top             =   0
               Width           =   1335
            End
         End
         Begin VB.Frame Frame6 
            BorderStyle     =   0  'None
            Caption         =   "Frame6"
            Height          =   6495
            Left            =   -74880
            TabIndex        =   160
            Top             =   360
            Width           =   7215
            Begin VB.CommandButton cmdDeleteJDValue 
               Caption         =   "Remove"
               Height          =   375
               Left            =   6120
               TabIndex        =   166
               Top             =   1800
               Width           =   975
            End
            Begin VB.CommandButton cmdAddJDValue 
               Caption         =   "Add"
               Height          =   375
               Left            =   3720
               TabIndex        =   165
               Top             =   1800
               Width           =   975
            End
            Begin VB.CommandButton cmdEditJDValue 
               Caption         =   "Edit"
               Height          =   375
               Left            =   4920
               TabIndex        =   164
               Top             =   1800
               Width           =   975
            End
            Begin VB.TextBox txtEmpJDValue 
               Height          =   1815
               Left            =   3720
               Locked          =   -1  'True
               MaxLength       =   5000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   161
               Top             =   240
               Width           =   3375
            End
            Begin MSComctlLib.ListView lvwJDValues 
               Height          =   4215
               Left            =   3720
               TabIndex        =   162
               Top             =   2280
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   7435
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   0   'False
               HideSelection   =   0   'False
               HideColumnHeaders=   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "No."
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "JD Value"
                  Object.Width           =   10583
               EndProperty
            End
            Begin MSComctlLib.TreeView tvwJDFields 
               Height          =   6255
               Left            =   120
               TabIndex        =   163
               Top             =   240
               Width           =   3495
               _ExtentX        =   6165
               _ExtentY        =   11033
               _Version        =   393217
               HideSelection   =   0   'False
               Indentation     =   441
               LabelEdit       =   1
               Style           =   7
               Appearance      =   1
            End
         End
         Begin VB.Frame fraOtherDetails 
            Caption         =   "Other Details:"
            Height          =   3015
            Left            =   -74880
            TabIndex        =   143
            Top             =   3840
            Width           =   7095
            Begin VB.Frame Frame1 
               Caption         =   " Employee Project Allocation "
               Height          =   2295
               Left            =   0
               TabIndex        =   150
               Top             =   720
               Width           =   7095
               Begin VB.CommandButton cmdDeleteEmployeeProgrammeFunding 
                  Caption         =   "Delete"
                  Height          =   375
                  Left            =   2040
                  TabIndex        =   157
                  Top             =   1800
                  Width           =   1455
               End
               Begin VB.TextBox txtEmployeeProgrammeFundingPercentage 
                  Height          =   315
                  Left            =   1080
                  TabIndex        =   155
                  Top             =   1320
                  Width           =   2415
               End
               Begin VB.CommandButton cmdEnterEmployeeProgrammeFunding 
                  Caption         =   "Enter"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   156
                  Top             =   1800
                  Width           =   1455
               End
               Begin VB.ComboBox cboProgrammes 
                  Height          =   315
                  Left            =   1080
                  Style           =   2  'Dropdown List
                  TabIndex        =   153
                  Top             =   360
                  Width           =   2415
               End
               Begin VB.ComboBox cboFundCodes 
                  Height          =   315
                  Left            =   1080
                  Style           =   2  'Dropdown List
                  TabIndex        =   154
                  Top             =   840
                  Width           =   2415
               End
               Begin MSComctlLib.ListView lvwEmployeeProgrammeFunding 
                  Height          =   2055
                  Left            =   3720
                  TabIndex        =   158
                  Top             =   120
                  Width           =   3255
                  _ExtentX        =   5741
                  _ExtentY        =   3625
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   -1  'True
                  HideSelection   =   0   'False
                  FullRowSelect   =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  Appearance      =   1
                  NumItems        =   3
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "Project"
                     Object.Width           =   2540
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Text            =   "Fund Code"
                     Object.Width           =   2540
                  EndProperty
                  BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   2
                     Text            =   "Percentage"
                     Object.Width           =   2540
                  EndProperty
               End
               Begin VB.Label Label8 
                  Caption         =   "Percentage:"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   159
                  Top             =   1320
                  Width           =   855
               End
               Begin VB.Label Label57 
                  Caption         =   "Project:"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   152
                  Top             =   360
                  Width           =   615
               End
               Begin VB.Label Label58 
                  Caption         =   "Fund Code:"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   151
                  Top             =   840
                  Width           =   855
               End
            End
            Begin VB.ComboBox cboLocation 
               Height          =   315
               Left            =   4320
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   147
               Top             =   240
               Width           =   2655
            End
            Begin VB.ComboBox cboCountry 
               Height          =   315
               Left            =   1080
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   146
               Top             =   240
               Width           =   2295
            End
            Begin VB.Label Label29 
               Caption         =   "Country:"
               Height          =   255
               Left            =   360
               TabIndex        =   145
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label9 
               Caption         =   "Location:"
               Height          =   255
               Left            =   3600
               TabIndex        =   144
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame fraExistingNextOfKin 
            Caption         =   "Existing Next Of Kin: "
            Height          =   2175
            Left            =   -74880
            TabIndex        =   130
            Top             =   3960
            Width           =   7095
            Begin MSComctlLib.ListView lvwNextOfKins 
               Height          =   1815
               Left            =   120
               TabIndex        =   60
               Top             =   240
               Width           =   6855
               _ExtentX        =   12091
               _ExtentY        =   3201
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   0   'False
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   6
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Surname"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Other Names"
                  Object.Width           =   5292
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Relationship"
                  Object.Width           =   5292
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Guardian Full Names"
                  Object.Width           =   3528
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Guardian IDNo"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Guardian Relationship"
                  Object.Width           =   3528
               EndProperty
            End
         End
         Begin VB.Frame fraNextOfKin 
            Enabled         =   0   'False
            Height          =   3375
            Left            =   -74880
            TabIndex        =   129
            Top             =   480
            Width           =   7095
            Begin VB.CommandButton cmdGuardian 
               Caption         =   "Guadian..."
               Height          =   495
               Left            =   5880
               TabIndex        =   59
               Top             =   2760
               Width           =   1095
            End
            Begin VB.TextBox txtNoKBenefitPercent 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3480
               Locked          =   -1  'True
               TabIndex        =   58
               Top             =   3000
               Width           =   1455
            End
            Begin VB.CheckBox chkNoKBeneficiary 
               Caption         =   "Beneficiary"
               Height          =   195
               Left            =   1200
               TabIndex        =   57
               Top             =   3000
               Width           =   1215
            End
            Begin VB.CheckBox chkNoKDependant 
               Caption         =   "Dependant (Medical Cover)"
               Height          =   195
               Left            =   3480
               TabIndex        =   56
               Top             =   2640
               Width           =   2295
            End
            Begin VB.CheckBox chkNoKEmergency 
               Caption         =   "Emergency Contact"
               Height          =   195
               Left            =   1200
               TabIndex        =   55
               Top             =   2640
               Width           =   1935
            End
            Begin VB.TextBox txtNoKEMail 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   4200
               TabIndex        =   53
               Top             =   1878
               Width           =   2775
            End
            Begin VB.TextBox txtNoKTelephone 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1200
               TabIndex        =   52
               Top             =   1878
               Width           =   1695
            End
            Begin VB.TextBox txtNoKAddress 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1200
               TabIndex        =   51
               Top             =   1476
               Width           =   5775
            End
            Begin VB.TextBox txtNoKOccupation 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1200
               TabIndex        =   54
               Top             =   2280
               Width           =   5775
            End
            Begin VB.TextBox txtNoKRelationship 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   4200
               TabIndex        =   50
               Top             =   1074
               Width           =   2775
            End
            Begin VB.TextBox txtNoKIDNo 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1200
               TabIndex        =   49
               Top             =   1074
               Width           =   1695
            End
            Begin VB.ComboBox cboNoKGender 
               Height          =   315
               ItemData        =   "frmEmployee.frx":008C
               Left            =   4200
               List            =   "frmEmployee.frx":0099
               Style           =   2  'Dropdown List
               TabIndex        =   48
               Top             =   645
               Width           =   2775
            End
            Begin MSComCtl2.DTPicker dtpNoKDOB 
               Height          =   375
               Left            =   1200
               TabIndex        =   47
               Top             =   612
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   661
               _Version        =   393216
               CustomFormat    =   "dd-MMM-yyyy"
               Format          =   62717955
               CurrentDate     =   39228
            End
            Begin VB.TextBox txtNoKOtherNames 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   4200
               TabIndex        =   46
               Top             =   240
               Width           =   2775
            End
            Begin VB.TextBox txtNoKSurname 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1200
               TabIndex        =   45
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label Label55 
               Caption         =   "Benefit (%):"
               Height          =   255
               Left            =   2520
               TabIndex        =   141
               Top             =   3000
               Width           =   855
            End
            Begin VB.Label Label54 
               Alignment       =   1  'Right Justify
               Caption         =   "E-Mail:"
               Height          =   255
               Left            =   3000
               TabIndex        =   140
               Top             =   1890
               Width           =   1095
            End
            Begin VB.Label Label53 
               Caption         =   "Telephone:"
               Height          =   255
               Left            =   120
               TabIndex        =   139
               Top             =   1893
               Width           =   855
            End
            Begin VB.Label Label52 
               Caption         =   "Address:"
               Height          =   255
               Left            =   120
               TabIndex        =   138
               Top             =   1491
               Width           =   735
            End
            Begin VB.Label Label51 
               Caption         =   "Occupation:"
               Height          =   255
               Left            =   120
               TabIndex        =   137
               Top             =   2295
               Width           =   975
            End
            Begin VB.Label Label50 
               Alignment       =   1  'Right Justify
               Caption         =   "Relationship:"
               Height          =   255
               Left            =   3000
               TabIndex        =   136
               Top             =   1095
               Width           =   1095
            End
            Begin VB.Label Label49 
               Caption         =   "ID Number:"
               Height          =   255
               Left            =   120
               TabIndex        =   135
               Top             =   1089
               Width           =   855
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               Caption         =   "Gender:"
               Height          =   255
               Left            =   3000
               TabIndex        =   134
               Top             =   675
               Width           =   1095
            End
            Begin VB.Label Label47 
               Caption         =   "Date Of Birth:"
               Height          =   255
               Left            =   120
               TabIndex        =   133
               Top             =   672
               Width           =   1095
            End
            Begin VB.Label Label46 
               Alignment       =   1  'Right Justify
               Caption         =   "Other Names:"
               Height          =   255
               Left            =   3000
               TabIndex        =   132
               Top             =   255
               Width           =   1095
            End
            Begin VB.Label Label34 
               Caption         =   "Surname:"
               Height          =   255
               Left            =   120
               TabIndex        =   131
               Top             =   255
               Width           =   735
            End
         End
         Begin VB.Frame fraEOUV 
            Caption         =   "Selected Extra Organization Units"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   -74880
            TabIndex        =   99
            Top             =   2400
            Width           =   7095
            Begin MSComctlLib.ListView lvwOU 
               Height          =   1095
               Left            =   120
               TabIndex        =   44
               Top             =   240
               Width           =   6735
               _ExtentX        =   11880
               _ExtentY        =   1931
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   0   'False
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
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
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "OU Name"
                  Object.Width           =   3528
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Family Tree"
                  Object.Width           =   8819
               EndProperty
            End
         End
         Begin VB.CommandButton cmdSearchEOUV 
            Enabled         =   0   'False
            Height          =   465
            Left            =   -70680
            Picture         =   "frmEmployee.frx":00B8
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Select Other OUs where Employee is Visible"
            Top             =   1920
            Width           =   555
         End
         Begin VB.CommandButton cmdSearch 
            Height          =   465
            Left            =   -68340
            Picture         =   "frmEmployee.frx":0523
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Select OU for the Employee"
            Top             =   690
            Width           =   465
         End
         Begin VB.ComboBox cboOU 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -74820
            Locked          =   -1  'True
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   690
            Width           =   6405
         End
         Begin VB.TextBox txtOUInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Height          =   825
            Left            =   -74820
            MultiLine       =   -1  'True
            ScrollBars      =   1  'Horizontal
            TabIndex        =   100
            Top             =   1050
            Width           =   6375
         End
         Begin VB.CheckBox chkHasOUV 
            Caption         =   "Employee is Visible in Other Organization Units"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74820
            TabIndex        =   42
            ToolTipText     =   "Similar Organization Units MUST have same CODE and NAME"
            Top             =   2040
            Width           =   3795
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
            Height          =   6375
            Left            =   -74820
            TabIndex        =   82
            Top             =   360
            Width           =   7155
            Begin VB.ComboBox cboStaffCategory 
               Height          =   315
               Left            =   1440
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   21
               Top             =   2610
               Width           =   1935
            End
            Begin VB.ComboBox cboCurrency 
               Height          =   315
               Left            =   5040
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   23
               Top             =   240
               Width           =   1815
            End
            Begin VB.Frame fraBankInfo 
               Caption         =   "Employee's Main Bank Account: "
               Height          =   1395
               Left            =   0
               TabIndex        =   102
               Top             =   4800
               Width           =   7155
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
                  Left            =   4888
                  Locked          =   -1  'True
                  TabIndex        =   37
                  Top             =   300
                  Width           =   1937
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
                  TabIndex        =   36
                  Top             =   300
                  Width           =   1935
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
                  Locked          =   -1  'True
                  TabIndex        =   39
                  Top             =   990
                  Width           =   5415
               End
               Begin VB.TextBox txtAccountName 
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
                  Height          =   285
                  Left            =   1440
                  Locked          =   -1  'True
                  TabIndex        =   38
                  Top             =   645
                  Width           =   5385
               End
               Begin VB.Label Label38 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Account No."
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
                  Top             =   1035
                  Width           =   1365
               End
               Begin VB.Label Label37 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Branch Name"
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
                  Left            =   3660
                  TabIndex        =   105
                  Top             =   345
                  Width           =   945
               End
               Begin VB.Label Label36 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "Bank Name"
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
                  TabIndex        =   104
                  Top             =   345
                  Width           =   795
               End
               Begin VB.Label Label23 
                  Caption         =   "Account Name"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   120
                  TabIndex        =   103
                  Top             =   690
                  Width           =   1665
               End
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
               Height          =   315
               Left            =   5055
               TabIndex        =   29
               Top             =   2700
               Width           =   1755
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
               Height          =   315
               Left            =   5055
               TabIndex        =   28
               Top             =   2250
               Width           =   1755
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
               Height          =   315
               Left            =   5055
               TabIndex        =   26
               Top             =   1470
               Width           =   1755
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
               Height          =   315
               Left            =   5055
               TabIndex        =   27
               Top             =   1860
               Width           =   1755
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
               Height          =   315
               Left            =   5055
               TabIndex        =   30
               Top             =   3105
               Width           =   1755
            End
            Begin VB.TextBox txtBasicPay 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Left            =   5055
               TabIndex        =   24
               Top             =   630
               Width           =   1755
            End
            Begin VB.TextBox txtHAllow 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Left            =   5055
               TabIndex        =   25
               Top             =   1050
               Width           =   1755
            End
            Begin VB.CheckBox chkOnProbation 
               Caption         =   "On Probation"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1440
               TabIndex        =   31
               Top             =   3600
               Width           =   1275
            End
            Begin VB.TextBox txtProb 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Left            =   5055
               TabIndex        =   33
               Top             =   3900
               Width           =   1755
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
               ItemData        =   "frmEmployee.frx":098E
               Left            =   1455
               List            =   "frmEmployee.frx":099B
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   32
               Top             =   3900
               Width           =   1935
            End
            Begin VB.ComboBox cboDesig 
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
               Left            =   1455
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   1170
               Width           =   1935
            End
            Begin VB.ComboBox cboTerms 
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
               ItemData        =   "frmEmployee.frx":09BD
               Left            =   1455
               List            =   "frmEmployee.frx":09D3
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   19
               Top             =   1650
               Width           =   1935
            End
            Begin VB.ComboBox cboType 
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
               ItemData        =   "frmEmployee.frx":0A17
               Left            =   1455
               List            =   "frmEmployee.frx":0A24
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   2130
               Width           =   1935
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
               ItemData        =   "frmEmployee.frx":0A4D
               Left            =   1455
               List            =   "frmEmployee.frx":0A4F
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   22
               Top             =   3090
               Width           =   1935
            End
            Begin MSComCtl2.DTPicker dtpDEmployed 
               Height          =   315
               Left            =   1455
               TabIndex        =   16
               Top             =   210
               Width           =   1935
               _ExtentX        =   3413
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
               CustomFormat    =   "dd, MMM, yyyy"
               Format          =   62717955
               CurrentDate     =   37845
            End
            Begin MSComCtl2.DTPicker dtpValidThrough 
               Height          =   315
               Left            =   1455
               TabIndex        =   17
               Top             =   690
               Width           =   1935
               _ExtentX        =   3413
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
               CustomFormat    =   "dd, MMM, yyyy"
               Format          =   62717955
               CurrentDate     =   37845
            End
            Begin MSComCtl2.DTPicker dtpPSDate 
               Height          =   315
               Left            =   1440
               TabIndex        =   34
               Top             =   4320
               Visible         =   0   'False
               Width           =   1935
               _ExtentX        =   3413
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
               CustomFormat    =   "dd, MMM, yyyy"
               Format          =   62717955
               CurrentDate     =   37845
            End
            Begin MSComCtl2.DTPicker dtpCDate 
               Height          =   315
               Left            =   5055
               TabIndex        =   35
               Top             =   4305
               Visible         =   0   'False
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
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
               Format          =   62717955
               CurrentDate     =   37845
            End
            Begin VB.Frame fraDetails 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               Height          =   840
               Left            =   120
               TabIndex        =   116
               Top             =   5160
               Visible         =   0   'False
               Width           =   6660
               Begin VB.CheckBox chkMain 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "&Mark this as the employee's main account"
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
                  Height          =   255
                  Left            =   120
                  TabIndex        =   123
                  Top             =   1680
                  Width           =   3495
               End
               Begin VB.CommandButton cmdSchBank 
                  Height          =   315
                  Left            =   5640
                  Picture         =   "frmEmployee.frx":0A51
                  Style           =   1  'Graphical
                  TabIndex        =   122
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   315
               End
               Begin VB.CommandButton cmdSchBankBranch 
                  Height          =   315
                  Left            =   5640
                  Picture         =   "frmEmployee.frx":0DDB
                  Style           =   1  'Graphical
                  TabIndex        =   121
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   315
               End
               Begin VB.TextBox txtBranchName 
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
                  Height          =   330
                  Left            =   2400
                  Locked          =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   120
                  Top             =   600
                  Width           =   3255
               End
               Begin VB.TextBox Text1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E0E0E0&
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
                  Left            =   2400
                  Locked          =   -1  'True
                  TabIndex        =   119
                  Top             =   240
                  Width           =   3255
               End
               Begin VB.TextBox txtAccountType 
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
                  Height          =   330
                  Left            =   2400
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   118
                  Top             =   1320
                  Width           =   3255
               End
               Begin VB.TextBox txtAccountNumber 
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
                  Height          =   330
                  Left            =   2400
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   117
                  Top             =   960
                  Width           =   3270
               End
               Begin VB.Label Label25 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "Branch name"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   120
                  TabIndex        =   127
                  Top             =   600
                  Width           =   930
               End
               Begin VB.Label Label26 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "Bank name"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   120
                  TabIndex        =   126
                  Top             =   240
                  Width           =   780
               End
               Begin VB.Label Label27 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "Account type"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   120
                  TabIndex        =   125
                  Top             =   1320
                  Width           =   960
               End
               Begin VB.Label Label33 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "Account number"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   135
                  TabIndex        =   124
                  Top             =   960
                  Width           =   1170
               End
            End
            Begin VB.Label lblCurrency 
               Alignment       =   1  'Right Justify
               Caption         =   "."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4320
               TabIndex        =   149
               Top             =   660
               Width           =   615
            End
            Begin VB.Label Label59 
               Caption         =   "Staff Category:"
               Height          =   255
               Left            =   75
               TabIndex        =   148
               Top             =   2640
               Width           =   1215
            End
            Begin VB.Label Label56 
               Caption         =   "Currency:"
               Height          =   255
               Left            =   3600
               TabIndex        =   142
               Top             =   270
               Width           =   1095
            End
            Begin VB.Label Label35 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Prob. Period (M)"
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
               TabIndex        =   115
               Top             =   3960
               Width           =   1170
            End
            Begin VB.Label lblCDate 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Confirmation Date"
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
               TabIndex        =   114
               Top             =   4365
               Width           =   1305
            End
            Begin VB.Label Label28 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Basic Pay"
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
               TabIndex        =   113
               Top             =   690
               Width           =   675
            End
            Begin VB.Label Label24 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "House Allowance"
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
               TabIndex        =   112
               Top             =   1110
               Width           =   1215
            End
            Begin VB.Label Label39 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "KRA File No."
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
               TabIndex        =   111
               Top             =   1920
               Width           =   885
            End
            Begin VB.Label Label17 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "PIN No."
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
               TabIndex        =   110
               Top             =   1530
               Width           =   555
            End
            Begin VB.Label Label16 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "NSSF No."
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
               TabIndex        =   109
               Top             =   2310
               Width           =   675
            End
            Begin VB.Label Label15 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "NHIF No."
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
               TabIndex        =   108
               Top             =   2760
               Width           =   660
            End
            Begin VB.Label Label32 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Good Conduct No."
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
               TabIndex        =   107
               Top             =   3165
               Width           =   1320
            End
            Begin VB.Label Label21 
               Caption         =   "Probation Type"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   75
               TabIndex        =   98
               Top             =   3915
               Width           =   1185
            End
            Begin VB.Label lblSDate 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Start Date"
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
               Left            =   75
               TabIndex        =   89
               Top             =   4365
               Width           =   750
            End
            Begin VB.Label Label12 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Designation"
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
               Left            =   75
               TabIndex        =   88
               Top             =   1230
               Width           =   840
            End
            Begin VB.Label Label18 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Emp. Terms"
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
               Left            =   75
               TabIndex        =   87
               Top             =   1710
               Width           =   840
            End
            Begin VB.Label Label19 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Emp. Type"
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
               Left            =   75
               TabIndex        =   86
               Top             =   2190
               Width           =   765
            End
            Begin VB.Label Label10 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Emp. Grade"
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
               Left            =   75
               TabIndex        =   85
               Top             =   3150
               Width           =   840
            End
            Begin VB.Label Label6 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Date Employed"
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
               Left            =   75
               TabIndex        =   84
               Top             =   270
               Width           =   1080
            End
            Begin VB.Label Label41 
               Appearance      =   0  'Flat
               Caption         =   "Emp. Valid Thro"
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
               Left            =   75
               TabIndex        =   83
               Top             =   750
               Width           =   1290
            End
         End
         Begin VB.Frame Frame5 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5835
            Left            =   90
            TabIndex        =   72
            Top             =   960
            Width           =   7125
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
               ItemData        =   "frmEmployee.frx":1165
               Left            =   4860
               List            =   "frmEmployee.frx":1167
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   172
               Top             =   3049
               Width           =   2175
            End
            Begin VB.TextBox txtExternalRefNo 
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
               Left            =   1485
               TabIndex        =   171
               Top             =   1416
               Width           =   2085
            End
            Begin VB.TextBox txtDisabilityDet 
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
               Height          =   285
               Left            =   1485
               TabIndex        =   11
               Top             =   4290
               Width           =   5505
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
               Left            =   1485
               TabIndex        =   10
               Top             =   3945
               Width           =   1845
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
               Height          =   285
               Left            =   1485
               MultiLine       =   -1  'True
               TabIndex        =   13
               Top             =   5055
               Width           =   5505
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
               Height          =   285
               Left            =   1485
               MultiLine       =   -1  'True
               TabIndex        =   12
               Top             =   4650
               Width           =   5505
            End
            Begin VB.ComboBox CboReligion 
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
               ItemData        =   "frmEmployee.frx":1169
               Left            =   4860
               List            =   "frmEmployee.frx":116B
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   3480
               Width           =   2175
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
               Left            =   1485
               TabIndex        =   14
               Top             =   5460
               Width           =   2040
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
               Left            =   4410
               TabIndex        =   15
               Top             =   5460
               Width           =   2580
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
               ItemData        =   "frmEmployee.frx":116D
               Left            =   1485
               List            =   "frmEmployee.frx":116F
               Locked          =   -1  'True
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   3480
               Width           =   2085
            End
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
               ItemData        =   "frmEmployee.frx":1171
               Left            =   4860
               List            =   "frmEmployee.frx":1187
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   2637
               Width           =   2175
            End
            Begin VB.CommandButton cmdPDelete 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6090
               Picture         =   "frmEmployee.frx":11C7
               Style           =   1  'Graphical
               TabIndex        =   74
               Top             =   1080
               Visible         =   0   'False
               Width           =   320
            End
            Begin VB.CommandButton cmdPNew 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   5730
               Picture         =   "frmEmployee.frx":16B9
               Style           =   1  'Graphical
               TabIndex        =   73
               Top             =   1080
               Visible         =   0   'False
               Width           =   320
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
               ItemData        =   "frmEmployee.frx":17BB
               Left            =   4860
               List            =   "frmEmployee.frx":17C8
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   2225
               Width           =   2175
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
               Left            =   1485
               TabIndex        =   3
               Top             =   1828
               Width           =   2085
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
               Left            =   1485
               TabIndex        =   0
               Top             =   180
               Width           =   2085
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
               Left            =   1485
               TabIndex        =   1
               Top             =   592
               Width           =   2085
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
               Left            =   1485
               TabIndex        =   2
               Top             =   1004
               Width           =   2085
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
               Left            =   1485
               TabIndex        =   4
               Top             =   2240
               Width           =   2085
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
               Left            =   1485
               TabIndex        =   5
               Top             =   2652
               Width           =   2085
            End
            Begin MSComctlLib.ImageList imgEmpTool 
               Left            =   5760
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
                     Picture         =   "frmEmployee.frx":17E7
                     Key             =   "Search"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmEmployee.frx":18F9
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmEmployee.frx":1A0B
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmEmployee.frx":1B1D
                     Key             =   ""
                  EndProperty
               EndProperty
            End
            Begin MSComCtl2.DTPicker dtpDOB 
               Height          =   285
               Left            =   1485
               TabIndex        =   173
               Top             =   3064
               Width           =   2085
               _ExtentX        =   3678
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
               Format          =   62717955
               CurrentDate     =   37845
               MinDate         =   -36522
            End
            Begin VB.Label Label60 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "External Ref No"
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
               TabIndex        =   176
               Top             =   1455
               Width           =   1140
            End
            Begin VB.Label Label5 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Date Of Birth"
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
               TabIndex        =   175
               Top             =   3109
               Width           =   945
            End
            Begin VB.Label Label30 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Ethnicity"
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
               Left            =   3750
               TabIndex        =   174
               Top             =   3105
               Width           =   615
            End
            Begin VB.Label Label13 
               Caption         =   "Disability Details"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   90
               TabIndex        =   97
               Top             =   4335
               Width           =   1275
            End
            Begin VB.Label Label14 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Religion"
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
               Left            =   3750
               TabIndex        =   96
               Top             =   3540
               Width           =   555
            End
            Begin VB.Label Label31 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Nationality"
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
               TabIndex        =   95
               Top             =   3540
               Width           =   765
            End
            Begin VB.Label Label40 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Marital status"
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
               Left            =   3750
               TabIndex        =   94
               Top             =   2697
               Width           =   975
            End
            Begin VB.Label Label22 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Home Address"
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
               TabIndex        =   93
               Top             =   5100
               Width           =   1035
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Email"
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
               Left            =   3870
               TabIndex        =   92
               Top             =   5505
               Width           =   360
            End
            Begin VB.Label Label20 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Telephone"
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
               TabIndex        =   91
               Top             =   5505
               Width           =   750
            End
            Begin VB.Label Label42 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Physical Address"
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
               TabIndex        =   90
               Top             =   4710
               Width           =   1200
            End
            Begin VB.Label Label44 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Alien Card No."
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
               TabIndex        =   81
               Top             =   2697
               Width           =   1035
            End
            Begin VB.Label Label43 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Passport No."
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
               TabIndex        =   80
               Top             =   2285
               Width           =   930
            End
            Begin VB.Label Label3 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Other Names"
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
               TabIndex        =   79
               Top             =   1049
               Width           =   945
            End
            Begin VB.Label Label2 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Surname"
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
               TabIndex        =   78
               Top             =   637
               Width           =   630
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Staff No."
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
               TabIndex        =   77
               Top             =   225
               Width           =   660
            End
            Begin VB.Label Label7 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Gender"
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
               Left            =   3750
               TabIndex        =   76
               Top             =   2285
               Width           =   525
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "ID No."
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
               TabIndex        =   75
               Top             =   1873
               Width           =   465
            End
            Begin VB.Image imgLoadPic 
               Appearance      =   0  'Flat
               Height          =   180
               Left            =   4320
               MouseIcon       =   "frmEmployee.frx":205F
               MousePointer    =   99  'Custom
               Picture         =   "frmEmployee.frx":24A1
               Stretch         =   -1  'True
               ToolTipText     =   "Click this icon to ADD a new photo"
               Top             =   450
               Width           =   195
            End
            Begin VB.Image imgDeletePic 
               Appearance      =   0  'Flat
               Height          =   180
               Left            =   4320
               MouseIcon       =   "frmEmployee.frx":25EB
               MousePointer    =   99  'Custom
               Picture         =   "frmEmployee.frx":2A2D
               Stretch         =   -1  'True
               ToolTipText     =   "Click this icon to DELETE employee photo"
               Top             =   1080
               Width           =   195
            End
            Begin VB.Image Picture1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1815
               Left            =   4920
               Stretch         =   -1  'True
               Top             =   180
               Width           =   2115
            End
         End
         Begin VB.Label lblDisengaged 
            BackColor       =   &H00FFFFFF&
            Height          =   735
            Left            =   120
            TabIndex        =   128
            Top             =   360
            Width           =   7095
         End
         Begin VB.Label Label45 
            Caption         =   "Organization Unit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -74820
            TabIndex        =   101
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   120
         TabIndex        =   70
         Top             =   7200
         Visible         =   0   'False
         Width           =   7650
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
            Left            =   6600
            Picture         =   "frmEmployee.frx":2E6F
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Save Record"
            Top             =   120
            Visible         =   0   'False
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
            Left            =   7035
            Picture         =   "frmEmployee.frx":2F71
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Cancel Process"
            Top             =   120
            Visible         =   0   'False
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
      TabIndex        =   64
      Top             =   -90
      Width           =   7440
      Begin MSComctlLib.ListView lvwEmp 
         Height          =   6720
         Left            =   0
         TabIndex        =   65
         Top             =   120
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
         Picture         =   "frmEmployee.frx":3073
         Style           =   1  'Graphical
         TabIndex        =   69
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
         Picture         =   "frmEmployee.frx":3175
         Style           =   1  'Graphical
         TabIndex        =   68
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
         Picture         =   "frmEmployee.frx":3277
         Style           =   1  'Graphical
         TabIndex        =   67
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
               Picture         =   "frmEmployee.frx":3769
               Key             =   "A"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployee.frx":3BBB
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployee.frx":3ED5
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployee.frx":4327
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployee.frx":4779
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployee.frx":4BCB
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployee.frx":4EE5
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployee.frx":51FF
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployee.frx":5651
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployee.frx":596B
               Key             =   "B"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployee.frx":5DBD
               Key             =   "C"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployee.frx":620F
               Key             =   "D"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployee.frx":6661
               Key             =   "E"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployee.frx":6AB3
               Key             =   "F"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployee.frx":6F05
               Key             =   "G"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployee.frx":7357
               Key             =   "H"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployee.frx":77A9
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployee.frx":7BFB
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployee.frx":804D
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployee.frx":849F
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployee.frx":88F1
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployee.frx":8D43
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmployee.frx":9195
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
         TabIndex        =   66
         Top             =   6420
         Visible         =   0   'False
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'======== HRCORE DECLARATIONS =========
Private empphoto As EmployeePhoto

Public Sub ClearMyTexts()
    dtpCDate.Enabled = False
    FraList.Visible = False
    Cleartxt
    dtpDOB.value = Date
    'dtpDOB.Value = DateAdd("m", -220, Date)
    'txtDOB.Text = Format(dtpDOB.Value, "yyyy-mm-dd")
    dtpDEmployed.value = Date
    dtpCDate.value = Date
    EnterDOB = False
    EnterDEmp = False
    
    'Set Default Values
    cboNationality.Text = "Kenyan"
    CboReligion.Text = "Christian"
    cboType.ListIndex = 0

   
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    Set Picture1 = Nothing
End Sub
Public Sub SwitchEmp()
    'By Oscar: This Needs to be checked
    'Comment writen on: 2007.06.07
    
    Call disabletxt
    FraList.Visible = True
    EnableCmd
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    
    'By Oscar: Then for the above two lines to work, disable the frmMain Save and Cancel Commands
    frmMain2.cmdCancel.Enabled = False
    frmMain2.cmdEdit.Enabled = True
    frmMain2.cmdNew.Enabled = True
    frmMain2.cmdSave.Enabled = False
    
    SaveNew = False
    
    If DSource = "Local" Then
        With frmMain2
            .cmdNew.Enabled = True
            .cmdDelete.Enabled = True
            .cmdEdit.Enabled = True
            .cmdSave.Enabled = False
            .cmdCancel.Enabled = False
        End With
    End If
    'The  sub bellow was commented on 19:2:2007 because it was repeating what was done
    'Call DisplayRecords
End Sub


Private Sub cboCat_Click()
'        'search for the Parent i.e. CSSS Category
'    Dim i As Long
'
'    On Error GoTo errorHandler
'    Set selCSSSCategory = Nothing
'    ParentChangedFromCode = True
'    cboStaffCategory.ListIndex = -1
'    ParentChangedFromCode = False
'
'    Set selEmpCategory = EmpCats.FindEmployeeCategoryByName(Trim(cboCat.Text))
'    If Not (selEmpCategory Is Nothing) Then
'        If Not (selEmpCategory.CSSSCategory Is Nothing) Then
'            Set selCSSSCategory = empStaffCategories.FindCSSSCategoryByID(selEmpCategory.CSSSCategory.CSSSCategoryID)
'            If Not (selCSSSCategory Is Nothing) Then
'                For i = 0 To cboStaffCategory.ListCount - 1
'                    If cboStaffCategory.ItemData(i) = selCSSSCategory.CSSSCategoryID Then
'                       ParentChangedFromCode = True
'                        cboStaffCategory.ListIndex = i
'                        ParentChangedFromCode = False
'                        Exit For
'                    End If
'                Next i
'            End If
'        End If
'    End If
'
'    Exit Sub
'
'errorHandler:
End Sub

Private Sub cboCountry_Click()
    On Error GoTo ErrorHandler
    Set selCountry = Nothing
    If cboCountry.ListIndex > -1 Then
        Set selCountry = empCountries.FindCountryByID(cboCountry.ItemData(cboCountry.ListIndex))
    End If
    
    LoadLocationsOfCountry selCountry
    
    Exit Sub
    
ErrorHandler:
End Sub

Private Sub cboCurrency_Click()
    Dim i As Long
    
    On Error GoTo ErrorHandler
    Set selCurrency = Nothing
    'clear the currency label
    lblCurrency.Caption = ""
    If cboCurrency.ListIndex > -1 Then
        Set selCurrency = empCurrencies.FindCurrencyByID(cboCurrency.ItemData(cboCurrency.ListIndex))
        If Not (selCurrency Is Nothing) Then
            lblCurrency.Caption = UCase(selCurrency.CurrencyCode)
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    
    
End Sub

Private Sub cboProbType_Click()
    If cboProbType.Text = "Appointment" Then
        lblSDate.Visible = True
        dtpPSDate.Visible = True
        'dtpPSDate.Enabled = False
        
        lblCDate.Visible = True
        dtpCDate.Visible = True
        
        'dtpCDate.Enabled = False
        
        dtpCDate.value = DateAdd("m", Val(txtProb.Text), dtpDEmployed.value)
        
    ElseIf cboProbType.Text = "Promotion" Then
        lblSDate.Visible = True
        dtpPSDate.Visible = True
        
        dtpPSDate.value = Date
        lblCDate.Visible = True
        dtpCDate.Visible = True
        
        dtpCDate.Enabled = True
        dtpPSDate.Enabled = True
        
        dtpCDate.value = DateAdd("m", Val(txtProb.Text), dtpPSDate.value)
        
    Else
'        lblSDate.Visible = False
'        dtpPSDate.Visible = False
        lblSDate.Visible = True
        dtpPSDate.Visible = True
        'dtpPSDate.Enabled = False
        
'        lblCDate.Visible = False
'        dtpCDate.Visible = False
        
        lblCDate.Visible = True
        dtpCDate.Visible = True
        'dtpCDate.Enabled = False
        
        txtProb.Text = ""
    End If
End Sub

Private Sub cboProgrammes_Click()
    If Me.cboProgrammes.ListIndex >= 0 Then
        LoadFundCodesForSpecificProgramme (Me.cboProgrammes.ItemData(Me.cboProgrammes.ListIndex))
    End If
End Sub

Private Sub cboStaffCategory_Click()
    On Error GoTo ErrorHandler
    
    Set selCSSSCategory = Nothing
    If cboStaffCategory.ListIndex > -1 Then
        Set selCSSSCategory = empStaffCategories.FindCSSSCategoryByID(cboStaffCategory.ItemData(cboStaffCategory.ListIndex))
    End If
    
    LoadGradesOfCSSSCategory selCSSSCategory
    
    Exit Sub
    
ErrorHandler:

    MsgBox "An error has occurred while processing the selected Staff Category" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub chkDisabled_Click()
    If chkDisabled.value = vbChecked Then
        Me.txtDisabilityDet.Locked = False
    Else
        Me.txtDisabilityDet.Locked = True
    End If
End Sub

Private Sub chkNoKBeneficiary_Click()
    On Error GoTo ErrorHandler
    
    If chkNoKBeneficiary.value = vbChecked Then
        Me.txtNoKBenefitPercent.Locked = False
    Else
        Me.txtNoKBenefitPercent.Locked = True
    End If
    
    Exit Sub
    
ErrorHandler:
    
End Sub

Private Sub chkOnProbation_Click()
    If chkOnProbation.value = vbChecked Then
        Me.cboProbType.Enabled = True
        Me.txtProb.Enabled = True
        Me.dtpCDate.Enabled = True
        Me.dtpPSDate.Enabled = True
    Else
        Me.cboProbType.Enabled = False
        Me.txtProb.Enabled = False
        Me.dtpCDate.Enabled = False
        Me.dtpPSDate.Enabled = False
    End If
End Sub


Private Sub cmdAddJDValue_Click()
    Select Case LCase(cmdAddJDValue.Caption)
        Case "add"
            If selJDField Is Nothing Then
                MsgBox "First Select a JD Field to set Value(s) for", vbInformation, TITLES
                Exit Sub
            End If
            cmdAddJDValue.Caption = "Update"
            cmdEditJDValue.Caption = "Cancel"
            cmdDeleteJDValue.Enabled = False
            Me.lvwJDValues.Enabled = False
            Me.txtEmpJDValue.Locked = False
            Me.tvwJDFields.Enabled = False
            Me.txtEmpJDValue.Text = ""
            Me.txtEmpJDValue.SetFocus
            'DisableEnableTabs True
            
        Case "update"
            If AddEmployeeJDValue() = False Then Exit Sub
            cmdAddJDValue.Caption = "Add"
            cmdEditJDValue.Caption = "Edit"
            cmdDeleteJDValue.Enabled = True
            Me.lvwJDValues.Enabled = True
            Me.txtEmpJDValue.Locked = True
            Me.tvwJDFields.Enabled = True
            'DisableEnableTabs False
            
            'repopulate the  JD Values listview
            PopulateEmployeeJDs TempEmpJDs
            
            'repopulate
    End Select
End Sub

Private Function AddEmployeeJDValue() As Boolean
    Dim newEmpJD As HRCORE.EmployeeJD
    
    On Error GoTo ErrorHandler
    
    Set newEmpJD = New HRCORE.EmployeeJD
    
    If Len(Trim(Me.txtEmpJDValue.Text)) > 0 Then
        newEmpJD.FieldValue = Trim(Me.txtEmpJDValue.Text)
    Else
        MsgBox "The Employee JD Value is Required", vbExclamation, TITLES
        Me.txtEmpJDValue.SetFocus
        Exit Function
    End If
    
    Set newEmpJD.JDCategory = selJDField
    
    If Not (TempEmpJDs Is Nothing) Then
        TempEmpJDs.add newEmpJD
    Else
        Set TempEmpJDs = New HRCORE.EmployeeJDs
        
        TempEmpJDs.add newEmpJD
    End If
    
    AddEmployeeJDValue = True
    
    Exit Function
    
ErrorHandler:
    AddEmployeeJDValue = False
    
End Function

Private Function UpdateEmployeeJDValue() As Boolean
    On Error GoTo ErrorHandler
    If Not (selEmpJD Is Nothing) Then
        Set selEmpJD.JDCategory = selJDField
        
        If Len(Trim(Me.txtEmpJDValue.Text)) > 0 Then
            selEmpJD.FieldValue = Trim(Me.txtEmpJDValue.Text)
        Else
            MsgBox "The Employee JD Value is Required", vbExclamation, TITLES
            Me.txtEmpJDValue.SetFocus
            Exit Function
        End If
        
        selEmpJD.Modified = True
        UpdateEmployeeJDValue = True
        
    Else
        MsgBox "There is no JD Field Selected", vbInformation, TITLES
        UpdateEmployeeJDValue = False
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error has occurred while updating the Employee JD Value" & vbNewLine & err.Description, vbExclamation, TITLES
    UpdateEmployeeJDValue = False
    
End Function


Private Sub PopulateEmployeeJDs(ByVal TheEmpJDs As HRCORE.EmployeeJDs)
    Dim ItemX As ListItem
    Dim i As Long
    Dim ParentJDField As HRCORE.JDCategory
    
    On Error GoTo ErrorHandler
    
    Me.lvwJDValues.ListItems.Clear
    Me.txtEmpJDValue.Text = vbNullString
    'NB: Load only those that are not marked as Deleted
    Dim prev As String
    prev = "no"
    If Not (TheEmpJDs Is Nothing) Then
        For i = 1 To TheEmpJDs.count
           '' If TheEmpJDs.Item(i).Deleted = False Then
                If Not (TheEmpJDs.Item(i).JDCategory Is Nothing) Then
                    If Found(TheEmpJDs.Item(i).FieldValue, lvwJDValues) = False Then
                        Set ParentJDField = pJDFields.FindJDCategoryByID(TheEmpJDs.Item(i).JDCategory.JDCategoryID)
                        TheEmpJDs.Item(i).FieldNumber = ParentJDField.FieldNumber & "." & i
                        Set ItemX = Me.lvwJDValues.ListItems.add(, , TheEmpJDs.Item(i).FieldNumber)
                        ItemX.SubItems(1) = TheEmpJDs.Item(i).FieldValue
                        
                        ItemX.Tag = TheEmpJDs.Item(i).EmployeeJDID
                        prev = TheEmpJDs.Item(i).FieldValue
                    End If
                End If
            ''End If
        Next i
        If Me.lvwJDValues.ListItems.count > 0 Then
            lvwJDValues_ItemClick Me.lvwJDValues.ListItems.Item(1)
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while populating the Employee JD Values" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub
Private Function Found(Item As String, lvw As ListView) As Boolean
Found = False
If lvw.ListItems.count > 0 Then
Dim it As ListItem
For Each it In lvw.ListItems
If (it.SubItems(1) = Item) Then
Found = True
Exit Function
End If
Next it
Else
Found = False
End If
End Function


Private Sub DisableEnableTabs(ByVal DisableTabs As Boolean)
        
    On Error GoTo ErrorHandler
    
    If DisableTabs = True Then
        Me.sstEmployees.TabEnabled(0) = False
        Me.sstEmployees.TabEnabled(1) = False
        Me.sstEmployees.TabEnabled(2) = False
        Me.sstEmployees.TabEnabled(4) = False
    Else
        Me.sstEmployees.TabEnabled(0) = True
        Me.sstEmployees.TabEnabled(1) = True
        Me.sstEmployees.TabEnabled(2) = True
        Me.sstEmployees.TabEnabled(4) = True
    End If
    
    Exit Sub
    
ErrorHandler:
End Sub

Public Sub cmdCancel_Click()
    On Error GoTo ErrorHandler
    
    FraList.Visible = True
    EnableCmd
    Call disabletxt
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
    End If
    
    'Flag that the Employee Is Not In Edit Mode
    EmployeeIsInEditMode = False
    new_Record = False
    
    Call DisplayRecords
    
    Exit Sub
ErrorHandler:
    
End Sub

Public Sub cmdclose_Click()
    Unload Me
End Sub


Private Sub cmdDeleteEmployeeProgrammeFunding_Click()
    On Error GoTo ErrorHandler
    
    If Not (EmployeeIsInEditMode Or new_Record) Then
        MsgBox "For this command buttons to be effective you must either be adding a new employee or editing an existing employee record", vbExclamation, TITLES
        GoTo Finish
    End If
    If Me.lvwEmployeeProgrammeFunding.SelectedItem Is Nothing Then
        MsgBox "Please first select an employee project allocation entry that you wish to delete", vbExclamation, TITLES
    Else
        If MsgBox("Please confirm your decision to delete the selected employee project allocation entry", vbExclamation + vbYesNo, TITLES) = vbYes Then
        Dim tg As Integer
        tg = Me.lvwEmployeeProgrammeFunding.ListItems(lvwEmployeeProgrammeFunding.SelectedItem.Index).Tag
            Me.lvwEmployeeProgrammeFunding.ListItems.remove Me.lvwEmployeeProgrammeFunding.SelectedItem.Index
            Dim pr As Integer
            
            For pr = 1 To objEmployeeProgrammeFundings.count
                If objEmployeeProgrammeFundings.Item(pr).Employee.EmployeeID = SelectedEmployee.EmployeeID Then
                    ''If (objEmployeeProgrammeFundings.Item(pr).EmployeeProgrammeID = Me.lvwEmployeeProgrammeFunding.ListItems(lvwEmployeeProgrammeFunding.SelectedItem.Index).Tag) Then
                    If (objEmployeeProgrammeFundings.Item(pr).EmployeeProgrammeID = tg) Then
                    objEmployeeProgrammeFundings.remove (pr)
                    Dim empp As New EmployeeProgramme
                    empp.EmployeeProgrammeID = tg
                    
                    ''objEmployeeProgrammeFundings.Remove (pr)
                    empp.Delete
                    Exit For
                    End If
                End If
            Next pr
            
        End If
    End If
    If Me.lvwEmployeeProgrammeFunding.ListItems.count > 0 Then
        lvwEmployeeProgrammeFunding_ItemClick Me.lvwEmployeeProgrammeFunding.ListItems.Item(1)
    Else
        Me.txtEmployeeProgrammeFundingPercentage.Text = vbNullString
        If Me.cboProgrammes.ListCount > 0 Then
            Me.cboProgrammes.ListIndex = 0
        End If
    End If
Finish:
    Exit Sub
    
ErrorHandler:
    MsgBox "An Error has occurred while attempting to delete an employee project allocation entry" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub cmdDeleteJDValue_Click()
    Dim resp As Long
    On Error GoTo ErrorHandler
    
    Select Case LCase(cmdDeleteJDValue.Caption)
        Case "remove"
            If selEmpJD Is Nothing Then
                MsgBox "You have not selected the JD Value to Remove", vbExclamation, TITLES
                Exit Sub
            End If
            
            resp = MsgBox("Are you sure you want to remove the selected JD Value?", vbQuestion + vbYesNo, TITLES)
            If resp = vbYes Then
                selEmpJD.Deleted = True
                
                Set selEmpJD = Nothing
                PopulateEmployeeJDs TempEmpJDs
            End If
        
        Case "cancel"
            cmdAddJDValue.Enabled = True
            cmdEditJDValue.Caption = "Edit"
            cmdDeleteJDValue.Caption = "Remove"
            Me.lvwJDValues.Enabled = True
            Me.tvwJDFields.Enabled = True
            Me.txtEmpJDValue.Locked = True
        
    End Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while performing the requested operation" & vbNewLine & err.Description, vbExclamation, TITLES
        
End Sub

Public Sub cmdEdit_Click()
    
    If Not currUser Is Nothing Then
        If currUser.CheckRight("GenerelDetails") <> secModify Then
            MsgBox "You dont have right to add new record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    On Error GoTo ErrorTrap
    'Omnis_ActionTag = "E" 'Edits a record in the Omnis database 'monte++
    
    If Not (SelectedEmployee Is Nothing) Then
        
        
        FraList.Visible = False
        enabletxt
        dtpCDate.Enabled = False
        'DisableCmd
        cmdSave.Enabled = True
        cmdCancel.Enabled = True
        txtEmpCode.SetFocus
        SaveNew = False
        'make sure bank Info is Disabled
        fraBankInfo.Enabled = False
        
        If SelectedEmployee.EmploymentTerm.IsPermanent Then
            Me.dtpValidThrough.Enabled = False
        Else
            Me.dtpValidThrough.Enabled = True
        End If
        
        'Flag that Employee Is In Edit Mode
        EmployeeIsInEditMode = True
        
    Else
        MsgBox "Select the Employee to Edit", vbInformation, TITLES
        'PSave = True    'prompt save
        Call cmdCancel_Click
        'PSave = False
    End If
   
    Exit Sub
    
ErrorTrap:
    MsgBox err.Description, vbExclamation, TITLES
End Sub



Private Sub cmdEditJDValue_Click()
    Select Case LCase(cmdEditJDValue.Caption)
        Case "edit"
            If selEmpJD Is Nothing Then
                MsgBox "Select the Employee JD Value to Edit", vbInformation, TITLES
                Exit Sub
            End If
            
            If selEmpJD.Deleted = True Then
             MsgBox "Item cannot be Edited here.", vbOKOnly + vbCritical
             Exit Sub
            End If
            cmdAddJDValue.Enabled = False
            cmdEditJDValue.Caption = "Update"
            cmdDeleteJDValue.Caption = "Cancel"
            Me.lvwJDValues.Enabled = False
            Me.tvwJDFields.Enabled = False
            Me.txtEmpJDValue.Locked = False
            Me.txtEmpJDValue.SetFocus
            
        Case "update"
            If UpdateEmployeeJDValue() = False Then Exit Sub
            cmdAddJDValue.Enabled = True
            cmdEditJDValue.Caption = "Edit"
            cmdDeleteJDValue.Caption = "Remove"
            Me.lvwJDValues.Enabled = True
            Me.tvwJDFields.Enabled = True
            Me.txtEmpJDValue.Locked = True
            
            're-load
            ''Tie_to_category_values (selJDField.JDCategoryID)
            ''PopulateEmployeeJDs TempEmpJDs
            PopulateEmployeeJDs FilteredEmpJDs
        Case "cancel"
            cmdAddJDValue.Caption = "Add"
            cmdEditJDValue.Caption = "Edit"
            cmdDeleteJDValue.Enabled = True
            Me.lvwJDValues.Enabled = True
            Me.txtEmpJDValue.Locked = True
            Me.tvwJDFields.Enabled = True
            
    End Select
    
End Sub

Private Sub cmdEnterEmployeeProgrammeFunding_Click()
    Dim myListItem As ListItem
    Dim myListItem2 As ListItem
    Dim lngLoopVariable As Long
    Dim blnExistInList As Boolean
    Dim sngTotalPercentage As Single
    On Error GoTo ErrorHandler
    
    If Not (EmployeeIsInEditMode Or new_Record) Then
        MsgBox "For this command buttons to be effective you must either be adding a new employee or editing an existing employee record", vbExclamation, TITLES
        GoTo Finish
    End If
    'VALIDATE USER INPUT
    If ValidateEmployeeProgrammeInput = False Then Exit Sub
    For lngLoopVariable = 1 To Me.lvwEmployeeProgrammeFunding.ListItems.count
        Set myListItem = Me.lvwEmployeeProgrammeFunding.ListItems.Item(lngLoopVariable)
        If UCase(myListItem.Text) = UCase(Me.cboProgrammes.Text) Then
            blnExistInList = True
            Set myListItem2 = Me.lvwEmployeeProgrammeFunding.ListItems.Item(lngLoopVariable)
            If UCase(myListItem.SubItems(1)) = UCase(Me.cboFundCodes.Text) Then
                sngTotalPercentage = sngTotalPercentage + CSng(Me.txtEmployeeProgrammeFundingPercentage.Text)
            Else
                sngTotalPercentage = sngTotalPercentage + CSng(Me.lvwEmployeeProgrammeFunding.ListItems.Item(lngLoopVariable).SubItems(2))
            End If
        Else
            sngTotalPercentage = sngTotalPercentage + CSng(Me.lvwEmployeeProgrammeFunding.ListItems.Item(lngLoopVariable).SubItems(2))
        End If
    Next
    If Not blnExistInList Then
        sngTotalPercentage = sngTotalPercentage + CSng(Me.txtEmployeeProgrammeFundingPercentage.Text)
    Else
        Set myListItem = myListItem2
    End If
    If sngTotalPercentage > 100 Then
        MsgBox "The entry you are attempting to update will cause the total employee project allocation to exceed 100" & vbCrLf & "This is not allowed", vbExclamation, TITLES
        GoTo Finish
    End If
    If myListItem2 Is Nothing Then
        Set myListItem = Me.lvwEmployeeProgrammeFunding.ListItems.add(, , Me.cboProgrammes.Text)
    End If
    myListItem.SubItems(1) = Me.cboFundCodes.Text
    myListItem.SubItems(2) = CSng(Me.txtEmployeeProgrammeFundingPercentage.Text)
Finish:
    Exit Sub
               
ErrorHandler:
    MsgBox "An Error has occurred while attempting to enter an employee project allocation entry" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Function ValidateEmployeeProgrammeInput() As Boolean
    On Error GoTo ErrorHandler
    
    ValidateEmployeeProgrammeInput = False
    If Me.cboProgrammes.Text = vbNullString Then
        MsgBox "The employee project input is invalid", vbExclamation, TITLES
    Else
        If Me.cboFundCodes.Text = vbNullString Then
            MsgBox "The employee fundcode input is invalid", vbExclamation, TITLES
        Else
            If Me.txtEmployeeProgrammeFundingPercentage.Text = vbNullString Or Not IsNumeric(Me.txtEmployeeProgrammeFundingPercentage.Text) Or Val(Me.txtEmployeeProgrammeFundingPercentage.Text) = 0 Then
                MsgBox "The employee project fundcode percentage input is invalid", vbExclamation, TITLES
            Else
                ValidateEmployeeProgrammeInput = True
            End If
        End If
    End If
    Exit Function
    
ErrorHandler:
    MsgBox "An Error has occurred while attempting to validate the user input for employee project allocation" & vbNewLine & err.Description, vbExclamation, TITLES
        
End Function

Private Sub cmdGuardian_Click()
    On Error GoTo ErrorHandler
    
    'This will set the temporary next of kin for data transfer
    'the data will then be picked at the point of validation
    
    frmNextOfKinGuardian.Show , Me
    
    Exit Sub
    
ErrorHandler:
End Sub

Public Sub cmdNew_Click()
    On Error Resume Next
    
    If Not currUser Is Nothing Then
        If currUser.CheckRight("GenerelDetails") <> secModify Then
            MsgBox "You dont have right to add new record. Please liaise with the security admin"
            Exit Sub
        End If
    End If

    new_Record = True   'flag that a new record is being inserted

    'dtpCDate.Enabled = False
    FraList.Visible = False
    
    'enable textboxes
    enabletxt
    
    'clear controls
    Cleartxt
    Me.cmdSearch.Enabled = True
    Me.cmdSearch.Visible = True
    
    cboGender.Text = "Unspecified"
    cboMarritalStat.Text = "Unspecified"
    cboProbType.Text = "None"
    
    dtpDEmployed.value = Date
    dtpCDate.value = Date
    dtpPSDate.value = Date
    'Default Birthdate
    dtpDOB.value = Now - (18 * 365)
    
    EnterDOB = False        'check this out
    EnterDEmp = False       'check this out
    
    'Set Default Values
    cboNationality.Text = "(Unspecified)"
    CboReligion.Text = "(Unspecified)"
    cboTribe.Text = "(Unspecified)"
    cboDesig.Text = "(Unspecified)"
    cboTerms.Text = "(Unspecified)"
    cboCat.Text = "(Unspecified)"
    Me.dtpValidThrough.value = Date
    cboType.Text = "Unspecified"
'    DisableCmd
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    
    'flags that a new employee is being added
    SaveNew = True
    
    Set Picture1 = Nothing
    
    'Disable Bank Info fields
    fraBankInfo.Enabled = False
    
    'disable Disengagement info
    'Me.fraDisEngagementInfo.Enabled = False
    
    'clear the OUVisibilities Listview
    Me.lvwOU.ListItems.Clear
    'CLEAR THE EMPLOYEE PROGRAMMES AND JD VALUES LIST VIEWS
    Me.lvwEmployeeProgrammeFunding.ListItems.Clear
    Me.lvwJDValues.ListItems.Clear
    'disable the search for eouv
    Me.cmdSearchEOUV.Enabled = False
    
    'FOR THE PURPOSE OF DISPLAYING VALUES IN THE LIST VIEWS
    Set TempNextOfKins = New HRCORE.NextOfKins
    Set TempEmpJDs = New HRCORE.EmployeeJDs
    txtEmpCode.SetFocus
    
    'CLEAR NEXT OF KIN INFORMATION
    frmEmployee.ClearControlsNextOfKin
    frmEmployee.lvwNextOfKins.ListItems.Clear
    
    Set SelectedEmployee = New Employee
    'Call Enablecmds
End Sub

Private Sub Enablecmds()
    cmdSearch.Visible = True
    cmdSearch.Enabled = True
    cmdSearchEOUV.Enabled = True
    cmdSearchEOUV.Visible = True
End Sub

Private Sub LockBankInfoFields()
    Me.txtBankBranchName.Locked = True
    Me.txtBankName.Locked = True
    Me.txtAccountName.Locked = True
    Me.txtAccountNO.Locked = True
End Sub


Private Sub cmdNoKAdd_Click()
    Select Case LCase(cmdNoKAdd.Caption)
        Case "add"
            cmdNoKAdd.Caption = "Update"
            cmdNOKEdit.Caption = "Cancel"
            cmdNoKDelete.Enabled = False
            fraNextOfKin.Enabled = True
            fraExistingNextOfKin.Enabled = False
            ClearControlsNextOfKin
            txtNoKSurname.SetFocus
            
        Case "update"
            If AddNextOfKin() = False Then Exit Sub
            cmdNoKAdd.Caption = "Add"
            cmdNOKEdit.Caption = "Edit"
            cmdNoKDelete.Enabled = True
            fraNextOfKin.Enabled = False
            fraExistingNextOfKin.Enabled = True
            PopulateNextOfKins TempNextOfKins
    End Select
    
End Sub

Public Sub ClearControlsNextOfKin()
    On Error GoTo ErrorHandler
    
    Me.txtNoKAddress.Text = ""
    Me.txtNoKBenefitPercent.Text = ""
    Me.txtNoKEMail.Text = ""
    Me.txtNoKIDNo.Text = ""
    Me.txtNoKOccupation.Text = ""
    Me.txtNoKOtherNames.Text = ""
    Me.txtNoKRelationship.Text = ""
    Me.txtNoKSurname.Text = ""
    Me.txtNoKTelephone.Text = ""
    Me.cboNoKGender.ListIndex = -1
    Me.chkNoKBeneficiary.value = vbUnchecked
    Me.chkNoKDependant.value = vbUnchecked
    Me.chkNoKEmergency.value = vbUnchecked
    Me.dtpNoKDOB.value = Date
'    Me.lvwNextOfKins.ListItems.clear
    frmNextOfKinGuardian.txtFullNames.Text = vbNullString
    frmNextOfKinGuardian.txtIDNo.Text = vbNullString
    frmNextOfKinGuardian.txtRelationship.Text = vbNullString
    frmNextOfKinGuardian.chkDisAssociate.value = vbUnchecked
    
    Exit Sub
    
ErrorHandler:
    
End Sub

Private Function AddNextOfKin() As Boolean
    Dim newNoK As HRCORE.NextOfKin
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    
    Set newNoK = New HRCORE.NextOfKin
    
    'check if any Guardian Information is available
    If Not (TempNextOfKin Is Nothing) Then
        newNoK.GuardianFullNames = TempNextOfKin.GuardianFullNames
        newNoK.GuardianIDNo = TempNextOfKin.GuardianIDNo
        newNoK.GuardianRelationship = TempNextOfKin.GuardianRelationship
    End If
    
    If Len(Trim(txtNoKSurname.Text)) > 0 Then
        newNoK.SurName = Trim(txtNoKSurname.Text)
    Else
        MsgBox "The Surname is Required", vbExclamation, TITLES
        Me.txtNoKSurname.SetFocus
        Exit Function
    End If
    
    If Len(Trim(txtNoKRelationship.Text)) > 0 Then
        newNoK.Relationship = Trim(txtNoKRelationship.Text)
    Else
        MsgBox "The Relationship is Required", vbExclamation, TITLES
        Me.txtNoKRelationship.SetFocus
        Exit Function
    End If
    
    If chkNoKBeneficiary.value = vbChecked Or chkNoKDependant.value = vbChecked Or chkNoKEmergency.value = vbChecked Then
        If chkNoKBeneficiary.value = vbChecked Then
            newNoK.IsBeneficiary = True
        Else
            newNoK.IsBeneficiary = False
        End If
        
        If chkNoKDependant.value = vbChecked Then
            newNoK.IsDependant = True
        Else
            newNoK.IsDependant = False
        End If
        
        If chkNoKEmergency.value = vbChecked Then
            newNoK.IsEmergencyContact = True
        Else
            newNoK.IsEmergencyContact = False
        End If
    Else
        MsgBox "The Person you have entered should ATLEAST be a Dependant, Beneficiary or Emergency Contact", vbExclamation, TITLES
        chkNoKEmergency.SetFocus
        Exit Function
    End If
    
    If dtpNoKDOB.value > Date Then
        MsgBox "The Next of kin could not have been born later than today's date", vbExclamation
        Exit Function
    Else
        'If Abs(Year(dtpNoKDOB.value) - Year(Date)) < 18 Then
        If Abs(Year(dtpNoKDOB.value) - Year(Date)) < 18 And chkNoKBeneficiary.value = 1 Then
            If newNoK.GuardianFullNames = "" Or newNoK.GuardianIDNo = "" Then
                MsgBox "The Next of Kin is a Beneficiary, who is below 18 years" & vbNewLine & _
                "You will have to give details about a Guardian of the Beneficiary", vbExclamation, TITLES
                cmdGuardian.Enabled = True
                cmdGuardian.SetFocus
                Exit Function
            End If
        Else
            'UNSET GUARDIAN INFO COZ ITS NOT REQUIRED FOR THOSE ABOVE 18
            If Not (selNextOfKin Is Nothing) Then
                selNextOfKin.GuardianFullNames = ""
                selNextOfKin.GuardianIDNo = ""
                selNextOfKin.GuardianRelationship = ""
            End If
        
        End If
        
        If Len(Trim(Me.txtNoKIDNo.Text)) <= 0 And Abs(Year(dtpNoKDOB.value) - Year(Date)) > 18 Then
            MsgBox "The ID Number is Required for any Beneficiary above 18 Years", vbExclamation, TITLES
            txtNoKIDNo.SetFocus
            Exit Function
        End If
    End If
            
    'UNSET GUARDIAN INFO FOR UNDER 18, BUT NOT BENEFICIARY
    If (chkNoKBeneficiary.value <> vbChecked) And (Abs(Year(dtpNoKDOB.value) - Year(Date)) < 18) Then
        If Not (selNextOfKin Is Nothing) Then
            selNextOfKin.GuardianFullNames = ""
            selNextOfKin.GuardianIDNo = ""
            selNextOfKin.GuardianRelationship = ""
        End If
    End If
    
    If chkNoKBeneficiary.value = vbChecked Then
        If Len(Trim(Me.txtNoKBenefitPercent.Text)) > 0 Then
            If Not IsNumeric(Trim(Me.txtNoKBenefitPercent.Text)) Then
                MsgBox "The Percentage Benefit allocated to this Beneficiary is not Valid", vbExclamation, TITLES
                txtNoKBenefitPercent.SetFocus
                Exit Function
                
            Else
                If CSng(Trim(Me.txtNoKBenefitPercent.Text)) > 100 Or CSng(Trim(Me.txtNoKBenefitPercent.Text)) <= 0 Then
                    MsgBox "The Percentage Benefit for the Beneficiary should be between 0 - 100%", vbExclamation, TITLES
                    txtNoKBenefitPercent.SetFocus
                    Exit Function
                End If
            End If
        Else
            MsgBox "The Percentage Benefit is Required for the Beneficiary", vbExclamation, TITLES
            txtNoKBenefitPercent.SetFocus
            Exit Function
        End If
    End If
    
    If chkNoKEmergency.value = vbChecked Then
        If Trim(txtNoKTelephone.Text) = "" Then
            MsgBox "The Telephone Number is required for an Emergency Contact", vbExclamation, TITLES
            txtNoKTelephone.SetFocus
            Exit Function
        End If
        
        If Trim(txtNoKAddress.Text) = "" Then
            MsgBox "The Address for the Emergency Contact is Required", vbExclamation, TITLES
            txtNoKAddress.SetFocus
            Exit Function
        End If
        
    End If
    
    newNoK.EMail = Trim(txtNoKEMail.Text)
    newNoK.IdNo = Trim(txtNoKIDNo.Text)
    newNoK.Occupation = Trim(txtNoKOccupation.Text)
    newNoK.OtherNames = Trim(txtNoKOtherNames.Text)
    newNoK.PostalAddress = Trim(txtNoKAddress.Text)
    newNoK.Relationship = Trim(txtNoKRelationship.Text)
    newNoK.Telephone = Trim(txtNoKTelephone.Text)
    newNoK.DateOfBirth = CDate(dtpNoKDOB.value)
    If cboNoKGender.ListIndex > -1 Then
        newNoK.GenderStr = cboNoKGender.List(cboNoKGender.ListIndex)
    End If
    If Me.chkNoKBeneficiary.value = 1 Then
        newNoK.BenefitPercent = CSng(Me.txtNoKBenefitPercent.Text)
    End If
    'Flag that it is a new entity
    newNoK.IsNewEntity = True
    
    If TempNextOfKins Is Nothing Then
        Set TempNextOfKins = New HRCORE.NextOfKins
    End If
    
    newNoK.NextOfKinID = TempNextOfKins.GetNewNextOfKinID()
    
    If newNoK.IsBeneficiary And TempNextOfKins.CollectionHasEmergencyContact Then
        MsgBox "Please Note that the Employee already has another Emergency Contact", vbInformation, TITLES
    End If
    
    TempNextOfKins.add newNoK
    
    AddNextOfKin = True
    Exit Function
    
ErrorHandler:
    MsgBox "An error has occurred while adding the next of Kin" & vbNewLine & err.Description, vbExclamation, TITLES
    AddNextOfKin = False
    
End Function


Private Function UpdateNextOfKin() As Boolean
    On Error GoTo ErrorHandler
    
     'check if any Guardian Information is available
    If Not (TempNextOfKin Is Nothing) Then
        selNextOfKin.GuardianFullNames = TempNextOfKin.GuardianFullNames
        selNextOfKin.GuardianIDNo = TempNextOfKin.GuardianIDNo
        selNextOfKin.GuardianRelationship = TempNextOfKin.GuardianRelationship
    Else
       Set selNextOfKin = New HRCORE.NextOfKin
    End If
    
    If Len(Trim(txtNoKSurname.Text)) > 0 Then
        selNextOfKin.SurName = Trim(txtNoKSurname.Text)
    Else
        MsgBox "The Surname is Required", vbExclamation, TITLES
        Me.txtNoKSurname.SetFocus
        Exit Function
    End If
    
    If Len(Trim(txtNoKRelationship.Text)) > 0 Then
        selNextOfKin.Relationship = Trim(txtNoKRelationship.Text)
    Else
        MsgBox "The Relationship is Required", vbExclamation, TITLES
        Me.txtNoKRelationship.SetFocus
        Exit Function
    End If
    
    If chkNoKBeneficiary.value = vbChecked Or chkNoKDependant.value = vbChecked Or chkNoKEmergency.value = vbChecked Then
        If chkNoKBeneficiary.value = vbChecked Then
            selNextOfKin.IsBeneficiary = True
        Else
            selNextOfKin.IsBeneficiary = False
        End If
        
        If chkNoKDependant.value = vbChecked Then
            selNextOfKin.IsDependant = True
        Else
            selNextOfKin.IsDependant = False
        End If
        
        If chkNoKEmergency.value = vbChecked Then
            selNextOfKin.IsEmergencyContact = True
        Else
            selNextOfKin.IsEmergencyContact = False
        End If
    Else
        MsgBox "The Person you have entered should ATLEAST be a Dependant, Beneficiary or Emergency Contact", vbExclamation, TITLES
        chkNoKEmergency.SetFocus
        Exit Function
    End If
    
    'EDITED TO DISABLE MISASSIGNMENT OF NEXT OF KIN INFO TO
    If dtpNoKDOB.value > Date Then
        MsgBox "The Next of kin could not have been born later than today's date", vbExclamation
        Exit Function
    End If
    
    If Len(Trim(Me.txtNoKIDNo.Text)) <= 0 And Abs(Year(dtpNoKDOB.value) - Year(Date)) > 18 Then
            MsgBox "The ID Number is Required for any Beneficiary above 18 Years", vbExclamation, TITLES
            txtNoKIDNo.SetFocus
            Exit Function
    End If
    
    If chkNoKBeneficiary.value = vbChecked Then
        'If Abs(Year(dtpNoKDOB.value) - Year(Date)) < 18 Then
        If Abs(Year(dtpNoKDOB.value) - Year(Date)) < 18 And chkNoKBeneficiary.value = 1 Then
            If selNextOfKin.GuardianFullNames = "" Or selNextOfKin.GuardianIDNo = "" Then
                MsgBox "The Next of Kin is a Beneficiary, who is below 18 years" & vbNewLine & _
                "You will have to give details about a Guardian of the Beneficiary", vbExclamation, TITLES
                cmdGuardian.Enabled = True
                cmdGuardian.SetFocus
                Exit Function
            End If
        Else
            'UNSET GUARDIAN INFO COZ ITS NOT REQUIRED
            selNextOfKin.GuardianFullNames = ""
            selNextOfKin.GuardianIDNo = ""
            selNextOfKin.GuardianRelationship = ""
        End If
        
'        If Len(Trim(Me.txtNoKIDNo.Text)) <= 0 Then
'            MsgBox "The ID Number is Required for any Beneficiary above 18 Years", vbExclamation, TITLES
'            txtNoKIDNo.SetFocus
'            Exit Function
'        End If
        
        If Len(Trim(Me.txtNoKBenefitPercent.Text)) > 0 Then
            If Not IsNumeric(Trim(Me.txtNoKBenefitPercent.Text)) Then
                MsgBox "The Percentage Benefit allocated to this Beneficiary is not Valid", vbExclamation, TITLES
                txtNoKBenefitPercent.SetFocus
                Exit Function
                
            Else
                If CSng(Trim(Me.txtNoKBenefitPercent.Text)) > 100 Or CSng(Trim(Me.txtNoKBenefitPercent.Text)) <= 0 Then
                    MsgBox "The Percentage Benefit for the Beneficiary should be between 0 - 100%", vbExclamation, TITLES
                    txtNoKBenefitPercent.SetFocus
                    Exit Function
                End If
            End If
        Else
            MsgBox "The Percentage Benefit is Required for the Beneficiary", vbExclamation, TITLES
            txtNoKBenefitPercent.SetFocus
            Exit Function
        End If
    End If
    
    If chkNoKEmergency.value = vbChecked Then
        If Trim(txtNoKTelephone.Text) = "" Then
            MsgBox "The Telephone Number is required for an Emergency Contact", vbExclamation, TITLES
            txtNoKTelephone.SetFocus
            Exit Function
        End If
        
        If Trim(txtNoKAddress.Text) = "" Then
            MsgBox "The Address for the Emergency Contact is Required", vbExclamation, TITLES
            txtNoKAddress.SetFocus
            Exit Function
        End If
    End If
    
    selNextOfKin.EMail = Trim(txtNoKEMail.Text)
    selNextOfKin.IdNo = Trim(txtNoKIDNo.Text)
    selNextOfKin.Occupation = Trim(txtNoKOccupation.Text)
    selNextOfKin.OtherNames = Trim(txtNoKOtherNames.Text)
    selNextOfKin.PostalAddress = Trim(txtNoKAddress.Text)
    selNextOfKin.Relationship = Trim(txtNoKRelationship.Text)
    selNextOfKin.Telephone = Trim(txtNoKTelephone.Text)
    selNextOfKin.DateOfBirth = CDate(dtpNoKDOB.value)
    If Me.chkNoKBeneficiary.value = 1 Then
        selNextOfKin.BenefitPercent = CSng(Me.txtNoKBenefitPercent.Text)
    End If
    If cboNoKGender.ListIndex > -1 Then
        selNextOfKin.GenderStr = cboNoKGender.List(cboNoKGender.ListIndex)
    End If
    
    'Indicate that it was modified
    selNextOfKin.Modified = True
    
    If TempNextOfKins.CollectionHasEmergencyContactEx(selNextOfKin.NextOfKinID) Then
        MsgBox "Please Note that the Employee already has other Emergency Contacts", vbInformation, TITLES
    End If
    
    UpdateNextOfKin = True
    Exit Function
    
ErrorHandler:
    MsgBox "An error has occurred while Updating the next of Kin" & vbNewLine & err.Description, vbExclamation, TITLES
    UpdateNextOfKin = False
End Function


Private Sub PopulateNextOfKins(ByVal TheNextOfKins As HRCORE.NextOfKins)
    Dim i As Long
    Dim ItemX As ListItem
    
    On Error GoTo ErrorHandler
    ClearControlsNextOfKin
    Me.lvwNextOfKins.ListItems.Clear
    
    If Not (TheNextOfKins Is Nothing) Then
        For i = 1 To TheNextOfKins.count
            Set ItemX = Me.lvwNextOfKins.ListItems.add(, , TheNextOfKins.Item(i).SurName)
            ItemX.SubItems(1) = TheNextOfKins.Item(i).OtherNames
            ItemX.SubItems(2) = TheNextOfKins.Item(i).Relationship
            
            'EDITED BY JOHN TO TRAP MISASSIGNMENT OF GUARDIANS FOR NEXT OF KIN WHO ARE ABOVE 18
'            If (DateDiff("y", CDate(TheNextOfKins.Item(i).DateOfBirth), CDate(Format(Now, "M/d/yyyy"))) < 18) Then
            If (Now - (18 * 365)) < TheNextOfKins.Item(i).DateOfBirth Then
                ItemX.SubItems(3) = TheNextOfKins.Item(i).GuardianFullNames
                ItemX.SubItems(4) = TheNextOfKins.Item(i).GuardianIDNo
                ItemX.SubItems(5) = TheNextOfKins.Item(i).GuardianRelationship
            Else
                ItemX.SubItems(3) = ""
                ItemX.SubItems(4) = ""
                ItemX.SubItems(5) = ""
                
                'UNSET GUARDIAN INFO COZ IT'S NOT REQUIRED HERE
'                TheNextOfKins.Item(i).GuardianFullNames = ""
'                TheNextOfKins.Item(i).GuardianIDNo = ""
'                TheNextOfKins.Item(i).GuardianRelationship = ""
            End If
                        
            ItemX.Tag = TheNextOfKins.Item(i).NextOfKinID
        Next i
        
        If Me.lvwNextOfKins.ListItems.count > 0 Then
            lvwNextOfKins_ItemClick Me.lvwNextOfKins.ListItems.Item(1)
        End If
    End If
    Exit Sub
    
ErrorHandler:
    
End Sub


Private Sub cmdNoKDelete_Click()
    Dim resp As Long
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    
    Select Case LCase(cmdNoKDelete.Caption)
        Case "remove"
            If selNextOfKin Is Nothing Then
                MsgBox "Select the Next of Kin that you want to Remove", vbExclamation, TITLES
            Else
                resp = MsgBox("Are you sure you want to Remove the selected Next of Kin", vbQuestion + vbYesNo, TITLES)
                If resp = vbYes Then
                    'Flag for deletion
                    selNextOfKin.Deleted = True
                    Me.lvwNextOfKins.ListItems.remove Me.lvwNextOfKins.SelectedItem.Index
'                    TempNextOfKins.Remove Me.lvwNextOfKins.SelectedItem.Index
                    'repopulate
'                    PopulateNextOfKins TempNextOfKins
                End If
            End If
            
        Case "cancel"
            cmdNOKEdit.Caption = "Edit"
            cmdNoKDelete.Caption = "Remove"
            cmdNoKAdd.Enabled = True
            fraNextOfKin.Enabled = False
            fraExistingNextOfKin.Enabled = True
'            PopulateNextOfKins TempNextOfKins
    End Select
    If Me.lvwNextOfKins.ListItems.count > 0 Then
        lvwNextOfKins_ItemClick Me.lvwNextOfKins.ListItems(1)
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while performing the requested operation" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub cmdNOKEdit_Click()
    Select Case LCase(cmdNOKEdit.Caption)
        Case "edit"
            cmdNOKEdit.Caption = "Update"
            cmdNoKDelete.Caption = "Cancel"
            cmdNoKAdd.Enabled = False
            fraNextOfKin.Enabled = True
            fraExistingNextOfKin.Enabled = False
            txtNoKSurname.SetFocus
            
        Case "update"
            If UpdateNextOfKin() = False Then Exit Sub
            cmdNOKEdit.Caption = "Edit"
            cmdNoKDelete.Caption = "Remove"
            cmdNoKAdd.Enabled = True
            fraNextOfKin.Enabled = False
            fraExistingNextOfKin.Enabled = True
            PopulateNextOfKins TempNextOfKins
            
        Case "cancel"
            cmdNoKAdd.Caption = "Add"
            cmdNOKEdit.Caption = "Edit"
            cmdNoKDelete.Enabled = True
            fraNextOfKin.Enabled = False
            fraExistingNextOfKin.Enabled = True
            PopulateNextOfKins TempNextOfKins
    End Select
    
End Sub

Private Sub cmdPNew_Click()
    
    Dim picturepath, currentname As String
'    Dim MovePic As FileSystemObject
    If Len(Trim(txtEmpCode.Text)) <= 0 Then
        MsgBox "Enter Employee Code first", vbInformation, "Picture"
        Exit Sub
    Else
'        Set MovePic = New FileSystemObject
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
            Picture1.Stretch = True
            Picture1.Picture = LoadPicture(picturepath)
            
        Else
            MsgBox "Unsupported file format", vbExclamation, "Picture"
            On Error Resume Next
            Picture1.Stretch = True
            Picture1.Picture = LoadPicture(App.Path & "\Pic\Gen.jpg")
        End If
    
    End If
End Sub

Public Sub cmdSave_Click()
    
    Dim TheEmployee As HRCORE.Employee
    Dim myinternalEmployeeProgramme As HRCORE.EmployeeProgramme
    Dim lngLoopVariable As Long
    Dim retVal As Long
    Dim blnEditedEmployeeProject As Boolean
    Dim DuplEmp As HRCORE.Employee
        
    On Error GoTo ErrorHandler
    
    If Not currUser Is Nothing Then
    
    Dim xx As String
    xx = "GenerelDetails"
        If currUser.CheckRight("GenerelDetails") <> secModify Then
            MsgBox "You dont have right to add new record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    Set TheEmployee = ValidateEmployee()
    
    Dim dialogResult As String
    Dim msg As String
    If Not TheEmployee.Category Is Nothing Then
        If IsNumeric(TheEmployee.Category.HighestSalary) Then
                If (txtBasicPay.Text > TheEmployee.Category.HighestSalary) Then
                
                
                msg = "This Employee's Basic Pay has exceeded The Maximum Amount that an employee can earn in  Group - " & TheEmployee.Category.CategoryName & "."
                msg = msg & vbNewLine & " The Range that an employee can Earn In this Group is " & TheEmployee.Category.LowestSalary & " To " & TheEmployee.Category.HighestSalary
                msg = msg & vbNewLine & ". Would You like to Proceed?"
                
               dialogResult = MsgBox(msg, vbYesNo + vbInformation, "Basic pay Ranges")
               If dialogResult = 7 Then
               Exit Sub
               End If
                  
                End If
        
        End If
        
                If IsNumeric(TheEmployee.Category.LowestSalary) Then
                If (txtBasicPay.Text < TheEmployee.Category.LowestSalary) Then
                
                
                msg = "This Employee's Basic Pay is below The Minimum Amount that an employee can earn in Group - " & TheEmployee.Category.CategoryName & "."
                msg = msg & vbNewLine & " The Range that an employee can Earn In this Group is " & TheEmployee.Category.LowestSalary & " To " & TheEmployee.Category.HighestSalary
                msg = msg & vbNewLine & ". Would Yoy like to Proceed?"
                
               dialogResult = MsgBox(msg, vbYesNo + vbInformation, "Basic pay Ranges")
               If dialogResult = 7 Then
               Exit Sub
               End If
                  
                End If
        
        End If
        
    End If
    
    '--Flag This As The New Employee
    Set NewEmpInfo = New Employee
    Set NewEmpInfo = TheEmployee
    '--End AuditTrail Flag
    
    If Not (TheEmployee Is Nothing) Then
        If Not SelectedEmployee.EmployeeProjects Is Nothing Then Set TheEmployee.EmployeeProjects = SelectedEmployee.EmployeeProjects
        If Not TheEmployee.EmployeeProjects Is Nothing Then
            For Each myinternalEmployeeProgramme In TheEmployee.EmployeeProjects
                myinternalEmployeeProgramme.Deleted = True
            Next
        End If
        For retVal = 1 To Me.lvwEmployeeProgrammeFunding.ListItems.count
            lngLoopVariable = 1
            Do Until lngLoopVariable = objProgrammeFundings.count + 1
                If UCase(objProgrammeFundings.Item(lngLoopVariable).Programme.ProgrammeName) = UCase(Me.lvwEmployeeProgrammeFunding.ListItems.Item(retVal).Text) And _
                    UCase(objProgrammeFundings.Item(lngLoopVariable).FundCode.FundCodeName) = UCase(Me.lvwEmployeeProgrammeFunding.ListItems.Item(retVal).SubItems(1)) Then
                    Set myinternalEmployeeProgramme = TheEmployee.EmployeeProjects.FindEmployeeProgrammeByProgrammeFundingID(objProgrammeFundings.Item(lngLoopVariable).ProgrammeFundingID)
                    If myinternalEmployeeProgramme Is Nothing Then
                        Set myinternalEmployeeProgramme = New HRCORE.EmployeeProgramme
                        blnEditedEmployeeProject = False
                    Else
                        blnEditedEmployeeProject = True
                    End If
                    Set myinternalEmployeeProgramme.ProgrammeFunding = objProgrammeFundings.Item(lngLoopVariable)
                    Exit Do
                End If
                lngLoopVariable = lngLoopVariable + 1
            Loop
            myinternalEmployeeProgramme.EmployeeProgrammePercentage = Me.lvwEmployeeProgrammeFunding.ListItems(retVal).SubItems(2)
            myinternalEmployeeProgramme.DateJoined = Date
            myinternalEmployeeProgramme.Deleted = False
            If Not blnEditedEmployeeProject Then TheEmployee.EmployeeProjects.add myinternalEmployeeProgramme
        Next
        
 
        
        
        If SaveNew = True Then
            retVal = TheEmployee.InsertNew()
            ''save photo
            
            'Do AuditTrail
            
            If retVal = 0 Then
                new_Record = False  'this will allow items to be selected in the Employee Listview
                'update the collection i.e. refresh
                If photoisactive Then
                    If Picture1.Picture <> 0 Then
                         ''SavePicture Picture1.Picture, App.Path & "\PHOTOS\" & CompanyId & "-" & Me.txtempcode & ".jpg"
                         Dim empphoto As New EmployeePhoto
                         empphoto.EmployeeID = TheEmployee.EmployeeID
                         empphoto.Photo = Me.Picture1.Picture
                         empphoto.UpdateEmployeePhoto
                         ''AllEmployeesPhotos.UpdateEmployeesphotosInCollection empphoto
                    End If
                End If
              ''  Call frmMain2.LoadEmployeeList
                 Call frmMain2.LoadEmployeeList
                'FraEdit.Enabled = True
                MsgBox "The new employee has been added successfully", vbInformation, TITLES
                
                'Flag that the Employee is not in Edit Mode
                EmployeeIsInEditMode = False
                
                'Clear Entries on PD - Next Of Kin
                
            ElseIf retVal = -20000 Or TheEmployee.EmployeeID = -20000 Then
                'refresh the Employee List but dont select an employee i.e. SkipTheSelection=True
                Call frmMain2.LoadEmployeeList(True)
                ''Call frmMain2.LoadEmployeeListwithemployee(TheEmployee)
                Set DuplEmp = AllEmployees.FindEmployeeByCode(TheEmployee.EmpCode)
                If Not (DuplEmp Is Nothing) Then
                    MsgBox "Another Employee Exists with the Staff Number: " & TheEmployee.EmpCode & " i.e. " & vbNewLine & _
                    UCase(DuplEmp.SurName & ", " & DuplEmp.OtherNames), vbInformation, TITLES
                Else
                    MsgBox "Another Employee Exists with the Staff Number: " & TheEmployee.EmpCode, vbInformation, TITLES
                End If
                txtEmpCode.SetFocus
                Call CancelMain
                
            Else
                MsgBox "The New Employee was not added", vbInformation, TITLES
                Call CancelMain
            End If
        Else    'now do Update operation
            
            If SelectedEmployee Is Nothing Then
                MsgBox "The system was not able to configure the Update Information", vbInformation, TITLES
                Exit Sub
            End If
            
            'pass on the employeeID of the selected employee
            TheEmployee.EmployeeID = SelectedEmployee.EmployeeID
            'retVal = TheEmployee.Update()
            retVal = TheEmployee.Update(currUser.UserName)
            If retVal = 0 Then
                'save the employee photo to a file. this will later be
                'integrated in the database
                
                
                If photoisactive Then
                    If Picture1.Picture <> 0 Then
                         ''SavePicture Picture1.Picture, App.Path & "\PHOTOS\" & CompanyId & "-" & Me.txtempcode & ".jpg"
                         Set empphoto = New EmployeePhoto
                         empphoto.EmployeeID = TheEmployee.EmployeeID
                         empphoto.Photo = Me.Picture1.Picture
                         empphoto.UpdateEmployeePhoto
                        '' AllEmployeesPhotos.UpdateEmployeesphotosInCollection empphoto
                    Else
                         Set empphoto = New EmployeePhoto
                         empphoto.EmployeeID = TheEmployee.EmployeeID
                         empphoto.Photo = Me.Picture1.Picture
                         empphoto.UpdateEmployeePhoto
                         ''AllEmployeesPhotos.UpdateEmployeesphotosInCollection empphoto
                    End If
                End If
                'Flag that the Employee is not in edit mode
                EmployeeIsInEditMode = False
                
                
                'update the collection i.e. refresh
                Call frmMain2.LoadEmployeeList
                            
                'FraEdit.Enabled = True
                MsgBox "The Employee data has been updated successfully", vbInformation, TITLES
                
                
                
            ElseIf retVal = -20000 Or TheEmployee.EmployeeID = -20000 Then
                'refresh the Employee List but dont select an employee i.e. SkipTheSelection=True
                Call frmMain2.LoadEmployeeList(True)
                Set DuplEmp = AllEmployees.FindEmployeeByCode(TheEmployee.EmpCode)
                
                If Not (DuplEmp Is Nothing) Then
                    MsgBox "Another Employee Exists with the Staff Number: " & TheEmployee.EmpCode & " i.e. " & vbNewLine & _
                    UCase(DuplEmp.SurName & ", " & DuplEmp.OtherNames), vbInformation, TITLES
                Else
                    MsgBox "Another Employee Exists with the Staff Number: " & TheEmployee.EmpCode, vbInformation, TITLES
                End If
                
                txtEmpCode.SetFocus
                Call CancelMain
                
            Else
                MsgBox "The Employee data was not updated", vbInformation, TITLES
                'Call CancelMain
            End If
        End If
    Else
        'MsgBox "The New Employee data was not Validated successfully", vbInformation, TITLES
        Call CancelMain     'make buttons remain in the Saving Mode
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while saving the new employee data" & vbNewLine & err.Description, vbInformation, TITLES
       
End Sub

Private Function ValidateEmployee() As HRCORE.Employee

    Dim NewEmp As HRCORE.Employee
    Dim DuplEmp As HRCORE.Employee   'for checking duplicate employees
    Dim j As Long
    Dim ItemX As ListItem
    Dim EmpOUV As HRCORE.EmployeeOUVisibility
        
    On Error GoTo errHandler
    
    Set NewEmp = New HRCORE.Employee
    
    'validation of entries
    If Len(Trim(txtEmpCode.Text)) <= 0 Then
        MsgBox "The Staff Number is required", vbInformation, TITLES
        txtEmpCode.SetFocus
        Exit Function
    Else
        'find out whether there are duplicates
        If SaveNew = True Then
            Set DuplEmp = AllEmployees.FindEmployeeByCode(Trim(txtEmpCode.Text))
        Else
            Set DuplEmp = AllEmployees.FindEmployeeByCodeExclusive(Trim(txtEmpCode.Text), SelectedEmployee.EmployeeID)
        End If
        
        If DuplEmp Is Nothing Then
            NewEmp.EmpCode = Trim(txtEmpCode.Text)
        Else
            MsgBox "There is another Employee with the Staff Number: " & Trim(txtEmpCode.Text) & " i.e." & vbNewLine & _
            UCase(DuplEmp.SurName & ", " & DuplEmp.OtherNames), vbInformation, TITLES
            txtEmpCode.SetFocus
            Exit Function
        End If
    End If
    
    If Len(Trim(txtSurname.Text)) <= 0 Then
        MsgBox "The Surname is required," & vbCrLf, vbOKOnly + vbInformation, TITLES
        txtSurname.SetFocus
        Exit Function
    Else
        NewEmp.SurName = Trim(Me.txtSurname.Text)
    End If
    
    If Len(Trim(Me.txtONames.Text)) <= 0 Then
        MsgBox "At Least one other name is required", vbInformation, TITLES
        txtONames.SetFocus
        Exit Function
    Else
        NewEmp.OtherNames = Trim(Me.txtONames.Text)
    End If
    
    NewEmp.ExternalRefNo = Trim(Me.txtExternalRefNo.Text)
    
    If (Trim((txtAlien.Text = "")) And (Trim(txtPassport.Text = "")) And (Trim(txtIDNo.Text = ""))) Then
        MsgBox "Please ensure you've entered any of:" & vbCrLf & "Alien, Passport and ID Number", vbOKOnly + vbInformation, TITLES
        txtIDNo.SetFocus
        Exit Function
    Else
        NewEmp.AlienNo = Trim(Me.txtAlien.Text)
        NewEmp.IdNo = Trim(Me.txtIDNo.Text)
        NewEmp.PassportNo = Trim(Me.txtPassport.Text)
    End If
    
    If CDate(dtpDOB.value) > CDate(Date) Then
        MsgBox "The Date of Birth cannot be greator than the current date", vbInformation, TITLES
        dtpDOB.SetFocus
        Exit Function
    Else
        NewEmp.DateOfBirth = dtpDOB.value
    End If
    
    If UCase(cboNationality.Text) = "(UNSPECIFIED)" Then
        Set NewEmp.Nationality = Nothing
    Else
        Set NewEmp.Nationality = empNationalities.FindNationality(cboNationality.ItemData(cboNationality.ListIndex))
    End If
    
    If cboGender.ListIndex > -1 Then
        NewEmp.Gender = cboGender.ItemData(cboGender.ListIndex)
    End If
    If cboMarritalStat.ListIndex > -1 Then
        NewEmp.MaritalStatus = cboMarritalStat.ItemData(cboMarritalStat.ListIndex)
    End If
    
    If UCase(cboTribe.Text) = "(UNSPECIFIED)" Then
        Set NewEmp.Tribe = Nothing
    Else
        Set NewEmp.Tribe = empTribes.FindTribe(cboTribe.ItemData(cboTribe.ListIndex))
    End If
    
    If UCase(CboReligion.Text) = "(UNSPECIFIED)" Then
        Set NewEmp.Religion = Nothing
    Else
        Set NewEmp.Religion = empReligions.FindReligion(CboReligion.ItemData(CboReligion.ListIndex))
    End If
    
    If chkDisabled.value = vbChecked Then
        If Len(Trim(txtDisabilityDet.Text)) <= 0 Then
            MsgBox "Enter Details about the Disability", vbInformation, TITLES
            txtDisabilityDet.SetFocus
            Exit Function
        Else
            NewEmp.IsPhysicallyDisabled = True
            NewEmp.DisabilityDetails = Trim(txtDisabilityDet.Text)
        End If
    Else
        NewEmp.IsPhysicallyDisabled = False
        NewEmp.DisabilityDetails = ""
    End If
    
    NewEmp.PhysicalAddress = Trim(txtPhysicalAddress.Text)
    NewEmp.HomeAddress = Trim(txtHAddress.Text)
    NewEmp.HomeTelephone = Trim(txtTel.Text)
    NewEmp.EMailAddress = Trim(txtEmail.Text)
    
    
    
    If CDate(dtpDEmployed.value) > CDate(Date) Then
        MsgBox "Date of Employment cannot be later than the current date", vbInformation, TITLES
        dtpDEmployed.SetFocus
        Exit Function
    Else
        NewEmp.DateOfEmployment = dtpDEmployed.value
    End If
    If CDate(dtpValidThrough.value) < CDate(dtpDEmployed.value) Then
        If SelectedEmployee.EmploymentTerm.IsPermanent Then
        Else
            MsgBox "The Validity Date cannot be earlier than the Employment date", vbInformation, TITLES
            dtpValidThrough.SetFocus
            Exit Function
        End If
    Else
        NewEmp.EmploymentValidThrough = dtpValidThrough.value
    End If
    
    If cboDesig.ListIndex > -1 Then
        If UCase(cboDesig.Text) = "(UNSPECIFIED)" Then
            Set NewEmp.position = Nothing
        Else
            Set NewEmp.position = empPositions.FindJobPosition(cboDesig.ItemData(cboDesig.ListIndex))
        End If
    End If
    
    If UCase(cboTerms.Text) = "(UNSPECIFIED)" Then
        Set NewEmp.EmploymentTerm = Nothing
    Else
        Set NewEmp.EmploymentTerm = empTerms.FindEmploymentTerm(cboTerms.ItemData(cboTerms.ListIndex))
    End If
    
    NewEmp.EmployeeType = cboType.ItemData(cboType.ListIndex)
    If UCase(cboCat.Text) = "(UNSPECIFIED)" Or UCase(cboCat.Text) = vbNullString Then
        Set NewEmp.Category = Nothing
    Else
        Set NewEmp.Category = EmpCats.FindEmployeeCategory(cboCat.ItemData(cboCat.ListIndex))
    End If
    
    If IsNumeric(Trim(txtBasicPay.Text)) Then
        NewEmp.BasicPay = CSng(Trim(txtBasicPay.Text))
    Else
        If Len(Trim(txtBasicPay.Text)) = 0 Then
            NewEmp.BasicPay = 0
        Else
            MsgBox "Enter a numeric value for basic pay", vbInformation, TITLES
            txtBasicPay.SetFocus
            Exit Function
        End If
    End If
    
    If IsNumeric(Trim(txtHAllow.Text)) Then
        NewEmp.HouseAllowance = CSng(Trim(txtHAllow.Text))
    Else
        If Len(Trim(txtHAllow.Text)) = 0 Then
            NewEmp.HouseAllowance = 0
        Else
            MsgBox "Enter a numeric value for House Allowance", vbInformation, TITLES
            txtHAllow.SetFocus
            Exit Function
        End If
    End If
    
    NewEmp.PinNo = Trim(txtPin.Text)
    NewEmp.KRAFileNo = Trim(txtKRAFileNO.Text)
    NewEmp.NssfNo = Trim(txtNssf.Text)
    NewEmp.NhifNo = Trim(txtNhif.Text)
    NewEmp.GoodConductCertNo = Trim(txtCert.Text)
    If chkOnProbation.value = vbChecked Then
        NewEmp.IsOnProbation = True
        NewEmp.ProbationType = cboProbType.ItemData(cboProbType.ListIndex)
        If NewEmp.ProbationType <> None Then
            If IsNumeric(Trim(txtProb.Text)) Then
                If CLng(Trim(txtProb.Text)) < 0 Then
                    MsgBox "Enter a numeric Probation Period value greator than zero", vbInformation, TITLES
                    txtProb.SetFocus
                    Exit Function
                Else
                    NewEmp.ProbationPeriod = CLng(Trim(txtProb.Text))
                End If
            Else
                MsgBox "Enter a numeric value for probation period", vbInformation, TITLES
            End If
            If CDate(dtpPSDate.value) > CDate(dtpCDate.value) Then
                MsgBox "The Confirmation Date should be later than the Start Date", vbInformation, TITLES
                dtpCDate.SetFocus
                Exit Function
            Else
                NewEmp.ProbationStartDate = dtpPSDate.value
                NewEmp.ConfirmationDate = dtpCDate.value
            End If
        Else
            NewEmp.ProbationType = None
        End If
    Else
        NewEmp.IsOnProbation = False
        NewEmp.ProbationType = None
        NewEmp.ProbationPeriod = 0
    End If
    
    If UCase(cboOU.Text) = "(UNSPECIFIED)" Then
        Set NewEmp.OrganizationUnit = Nothing
    Else
        If cboOU.ListIndex > -1 Then
            Set NewEmp.OrganizationUnit = OUnits.FindOrganizationUnit(cboOU.ItemData(cboOU.ListIndex))
        Else
            Set NewEmp.OrganizationUnit = Nothing
        End If
    End If
    
    If chkHasOUV.value = vbChecked Then
        NewEmp.HasOUVisibility = True
        
        'first clear the ouvisibilities
        NewEmp.VisibleInTheseOUs.Clear
        
        'now add fresh
        For Each ItemX In Me.lvwOU.ListItems
            Set EmpOUV = New HRCORE.EmployeeOUVisibility
            Set EmpOUV.Employee = NewEmp
            Set EmpOUV.OrganizationUnit = OUnits.FindOrganizationUnit(CLng(ItemX.Tag))
            NewEmp.VisibleInTheseOUs.add EmpOUV
        Next ItemX
    Else
        NewEmp.HasOUVisibility = False
        NewEmp.VisibleInTheseOUs.Clear
    End If
    
    'validate Location
    If UCase(cboLocation.Text) = "(UNSPECIFIED)" Or UCase(cboLocation.Text) = "UNSPECIFIED" Or cboLocation.ListIndex <= -1 Then
        Set NewEmp.Location = Nothing
    Else
        Set NewEmp.Location = empLocations.FindLocationByID(cboLocation.ItemData(cboLocation.ListIndex))
    End If
    
    'validate currency
    If UCase(cboCurrency.Text) = "(UNSPECIFIED)" Or UCase(cboCurrency.Text) = "UNSPECIFIED" Or cboCurrency.ListIndex <= -1 Then
        Set NewEmp.CurrencyType = Nothing
    Else
        Set NewEmp.CurrencyType = empCurrencies.FindCurrencyByID(cboCurrency.ItemData(cboCurrency.ListIndex))
        
    End If
    
    'set the JDs
    Set NewEmp.JobDescriptionValues = TempEmpJDs
    
    'set the Next Of Kins
    Set NewEmp.NextOfKins = TempNextOfKins
    
    Set ValidateEmployee = NewEmp
    
    
    Exit Function
    
errHandler:
    MsgBox "An error has occurred while Validating Employee Information" & vbNewLine & _
        err.Description, vbInformation, TITLES
    
    FraEdit.Enabled = True
    Set ValidateEmployee = Nothing
End Function

Private Sub cmdSearch_Click()
    Dim theOU As HRCORE.OrganizationUnit
    
    
    On Error GoTo ErrorHandler
    
    Set theOU = OUnits.SelectSingleOrganizationUnit()
    If Not (theOU Is Nothing) Then
        cboOU.Text = theOU.OrganizationUnitName
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "A slight error has occurred" & vbNewLine & err.Description, vbInformation, TITLES
End Sub

Private Sub cmdSearchEOUV_Click()
    Dim VisibleOUs As HRCORE.OrganizationUnits
    Dim i As Long
    Dim ItemX As ListItem
    Dim theOU As HRCORE.OrganizationUnit
    Dim ExistingVisibleOUs As HRCORE.OrganizationUnits
    
    On Error GoTo ErrorHandler
    
    'First get the Existing OU Visibilities
    'instantiate the collection
    Set ExistingVisibleOUs = New HRCORE.OrganizationUnits
    
        
    'now add items to the collection
    'otherwise its count will be zero, which is still OK
    For Each ItemX In Me.lvwOU.ListItems
        Set theOU = OUnits.FindOrganizationUnit(CLng(ItemX.Tag))
        If Not (theOU Is Nothing) Then
            ExistingVisibleOUs.add theOU
        End If
    Next ItemX
    
    'now call the method and pass the existing ous so that they are selected
    Set VisibleOUs = OUnits.SelectOrganizationUnitsWhereEmpIsVisible(selOU, ExistingVisibleOUs)
    
    'now clear the listview
    Me.lvwOU.ListItems.Clear
    
    'update the visible OUs
    If Not (VisibleOUs Is Nothing) Then
        For i = 1 To VisibleOUs.count
            Set theOU = VisibleOUs.Item(i)
            Set ItemX = Me.lvwOU.ListItems.add(, , theOU.OrganizationUnitName)
            ItemX.SubItems(1) = OUnits.GetOUFamilyTree(theOU)
            ItemX.Tag = theOU.OrganizationUnitID
        Next i
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while populating the Visible OUs" & vbNewLine & err.Description, vbInformation, TITLES
            
End Sub

Private Sub dtpDEmployed_CloseUp()
    EnterDEmp = True
    Call CheckDEmployed
    'txtDEmployed.Text = Format(dtpDEmployed.Value, "yyyy-mm-dd")
    Select Case strDatePart
    Case "d"
        dtpValidThrough.value = DateAdd("d", strValue, dtpDEmployed.value)
    Case "w"
        dtpValidThrough.value = DateAdd("d", strValue * 7, dtpDEmployed.value)
    Case "m"
        dtpValidThrough.value = DateAdd("m", strValue, dtpDEmployed.value)
    Case "y"
        dtpValidThrough.value = DateAdd("m", strValue * 12, dtpDEmployed.value)
    End Select
End Sub

Private Sub dtpDEmployed_KeyPress(KeyAscii As Integer)
'    EnterDEmp = True
'    Call CheckDEmployed
End Sub

Private Sub dtpDOB_CloseUp()
'    EnterDOB = True
'    Call CheckDOB
'    Dim rsAddDate As New ADODB.Recordset
'    Set rsAddDate = CConnect.GetRecordSet("select * from GeneralOpt where subsystem='" & SubSystem & "'")
'    If rsAddDate.RecordCount > 0 Then
'        dtpValidThrough.Value = DateAdd("m", IIf(cboGender.Text = "Male", IIf(IsNumeric(Trim(rsAddDate!MRet & "")) = True, Trim(rsAddDate!MRet & ""), 1) * 12, IIf(IsNumeric(Trim(rsAddDate!FRet & "")) = True, Trim(rsAddDate!FRet & ""), 1) * 12), dtpDOB.Value)
'    End If
End Sub

Private Sub dtpDOB_KeyPress(KeyAscii As Integer)
    EnterDOB = True
    Call CheckDOB
End Sub

Private Sub dtpPSDate_Change()
    dtpCDate.value = DateAdd("m", Val(txtProb.Text), dtpPSDate.value)
End Sub


Private Sub dtpValidThrough_CloseUp()
'txtValidThrough.Text = Format(dtpValidThrough.Value, "yyyy-mm-dd")
'If ((IsDate(txtValidThrough.Text) = True) And (IsDate(txtDEmployed.Text) = True)) Then
If DateDiff("d", dtpDEmployed.value, dtpValidThrough.value) < 0 Then MsgBox "The validity period of employment cannot be earlier than date of employment.", vbExclamation + vbOKOnly, "Wrong date"
'End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    Call InitializeHRCOREObjects
    
    'load All the JD Fields i.e. JDCategories
    Set pJDFields = New HRCORE.JDCategories
    LoadJDFields
    
    'this will temporarily hold employee jds
    Set TempEmpJDs = New HRCORE.EmployeeJDs
    
    Call disabletxt
'    Call InitGrid       'the employee listview, not necessary
    Call SetUpEnumCombos
    Call DisplayRecords
    cmdSave.Enabled = False
    cmdCancel.Enabled = False

    'position the form
    frmMain2.PositionTheFormWithEmpList Me
'    oSmart.FReset Me
    
    'select the First Tab
    sstEmployees.Tab = 0
    'LOADING
    
    'SET DEFAULT CURRENCY: WHILE LOADING THE EMPLOYEE FORM, NAVIGATE THE CURRENCIES COMBO TO THE DEFAULT CURRENCY BY DEAFULT
    Dim i As Long
    
    If empCurrencies.count > 0 Then
        For i = 1 To empCurrencies.count
            If (empCurrencies.Item(i).IsBaseCurrency) Then cboCurrency.Text = empCurrencies.Item(i).CurrencyCode & " (" & empCurrencies.Item(i).CurrencySymbol & ")"
            Exit For
        Next i
    End If
            
    Exit Sub
    
ErrorHandler:
    MsgBox err.Description, vbExclamation, TITLES
End Sub

'
Private Sub InitializeHRCOREObjects()
    LoadOUTypes
    LoadOrganizationUnits
    LoadEmployeeCategories
    LoadEmploymentTerms
    LoadNationalities
    LoadTribes
    LoadReligions
    LoadJobPositions
    LoadCurrencies
    LoadCountries
    LoadCSSSCategories
    LoadLocations
    LoadProgrammes
End Sub
Private Sub SetUpEnumCombos()
    'gender
    cboGender.Clear
    cboGender.AddItem "Unspecified"
    cboGender.ItemData(cboGender.NewIndex) = 0
    cboGender.AddItem "Male"
    cboGender.ItemData(cboGender.NewIndex) = 1
    cboGender.AddItem "Female"
    cboGender.ItemData(cboGender.NewIndex) = 2
    
    'Marital Status
    cboMarritalStat.Clear
    cboMarritalStat.AddItem "Unspecified"
    cboMarritalStat.ItemData(cboMarritalStat.NewIndex) = 0
    cboMarritalStat.AddItem "Single"
    cboMarritalStat.ItemData(cboMarritalStat.NewIndex) = 1
    cboMarritalStat.AddItem "Married"
    cboMarritalStat.ItemData(cboMarritalStat.NewIndex) = 2
    cboMarritalStat.AddItem "Divorced"
    cboMarritalStat.ItemData(cboMarritalStat.NewIndex) = 3
    cboMarritalStat.AddItem "Separated"
    cboMarritalStat.ItemData(cboMarritalStat.NewIndex) = 4
    cboMarritalStat.AddItem "Widowed"
    cboMarritalStat.ItemData(cboMarritalStat.NewIndex) = 5
    
    'Employment Type
    cboType.Clear
    cboType.AddItem "Unspecified"
    cboType.ItemData(cboType.NewIndex) = 0
    cboType.AddItem "Ordinary"
    cboType.ItemData(cboType.NewIndex) = 1
    cboType.AddItem "Agricultural"
    cboType.ItemData(cboType.NewIndex) = 2
    
    'Probation type
    cboProbType.Clear
    cboProbType.AddItem "None"
    cboProbType.ItemData(cboProbType.NewIndex) = 0
    cboProbType.AddItem "Appointment"
    cboProbType.ItemData(cboProbType.NewIndex) = 1
    cboProbType.AddItem "Promotion"
    cboProbType.ItemData(cboProbType.NewIndex) = 2
    
    
End Sub

Public Sub DisplayRecords()
    Dim ItemX As ListItem
    Dim i As Long
    Dim VisibleOU As HRCORE.OrganizationUnit
    
    Dim iCounter As Long
    Dim jCounter As Long
    On Error GoTo ErrorHandler
    'first clear records
    
    Call Cleartxt
   Set Picture1 = Nothing
    'clear the EOUV ListView
    Me.lvwOU.ListItems.Clear
    
    If Not (SelectedEmployee Is Nothing) Then
        '--Flag For AuditTrail
        Set OldEmpInfo = New Employee
        Set OldEmpInfo = SelectedEmployee
        '--End Flag
        
        If SelectedEmployee.EmploymentTerm.IsPermanent Then
            dtpValidThrough.Enabled = True
        Else
            dtpValidThrough.Enabled = False
        End If
        
        With SelectedEmployee
            txtEmpCode.Text = .EmpCode
            txtSurname.Text = .SurName
            txtONames.Text = .OtherNames
            txtExternalRefNo.Text = .ExternalRefNo
            txtIDNo.Text = .IdNo
            
            If .GenderStr <> "" Then
                cboGender.Text = .GenderStr
            End If
            
            txtPhysicalAddress.Text = .PhysicalAddress
            dtpDOB.value = .DateOfBirth
            dtpDEmployed.value = .DateOfEmployment
            If .OrganizationUnit.OrganizationUnitName <> "" Then
                cboOU.Text = .OrganizationUnit.OrganizationUnitName
            Else
                cboOU.Text = "(Unspecified)"
            End If
            If .HasOUVisibility = True Then
                chkHasOUV.value = vbChecked
                
                'load the EOUVs
                For i = 1 To .VisibleInTheseOUs.count
                    Set VisibleOU = OUnits.FindOrganizationUnit(.VisibleInTheseOUs.Item(i).OrganizationUnit.OrganizationUnitID)
                    If Not (VisibleOU Is Nothing) Then
                        Set ItemX = Me.lvwOU.ListItems.add(, , VisibleOU.OrganizationUnitName)
                        ItemX.SubItems(1) = VisibleOU.FamilyTree
                        ItemX.Tag = VisibleOU.OrganizationUnitID
                    End If
                Next i
                
            Else
                chkHasOUV.value = vbUnchecked
            End If
            If .EmploymentTerm.EmpTermName <> "" Then
                cboTerms.Text = .EmploymentTerm.EmpTermName
            Else
                cboTerms.Text = "(Unspecified)"
            End If
            If .EmployeeTypeStr <> "" Then
                cboType.Text = .EmployeeTypeStr
            Else
                cboType.ListIndex = 0
            End If
            
            txtPin.Text = .PinNo
            txtNssf.Text = .NssfNo
            txtNhif.Text = .NhifNo
            
            'to get the Grades right, first load their parent CSSS Category
            
            If Not (.Category Is Nothing) Then
                If Not (.Category.CSSSCategory Is Nothing) Then
                    For iCounter = 0 To cboStaffCategory.ListCount - 1
                        If cboStaffCategory.ItemData(iCounter) = .Category.CSSSCategory.CSSSCategoryID Then
                            cboStaffCategory.ListIndex = iCounter
                            Exit For
                        End If
                    Next iCounter
                    
                    For jCounter = 0 To cboCat.ListCount - 1
                        If cboCat.ItemData(jCounter) = .Category.CategoryID Then
                            cboCat.ListIndex = jCounter
                            Exit For
                        End If
                    Next jCounter
                End If
            Else
                cboCat.Text = "(Unspecified)"
            End If
            txtTel.Text = .HomeTelephone
            txtHAddress.Text = .HomeAddress
            txtEmail.Text = .EMailAddress
            
            If .position.PositionName <> "" Then
                cboDesig.Text = .position.PositionName  'Actually a combo box
            Else
                cboDesig.Text = "(Unspecified)"
            End If
            txtCert.Text = .GoodConductCertNo
            If .Nationality.Nationality <> "" Then
                cboNationality.Text = .Nationality.Nationality
            Else
                cboNationality.Text = "(Unspecified)"
            End If
            If .Tribe.Tribe <> "" Then
                cboTribe.Text = .Tribe.Tribe
            Else
                cboTribe.Text = "(Unspecified)"
            End If
            txtBasicPay.Text = Format(.BasicPay, "#,###,##0.00")
            txtHAllow.Text = Format(.HouseAllowance, "#,###,##0.00")
            
            'SECURITY ENFORCEMENT ON EMPLOYEE REMUNERATION DISPLAY
            If currUser.CheckRight("ViewEmployeeRemuneration") = secNone Then
                txtBasicPay.PasswordChar = "*"
                txtHAllow.PasswordChar = "*"
                'IF YOU CANT SEE THEN YOU SHOULD NOT BE ABLE TO EDIT
                txtBasicPay.Visible = False
                txtHAllow.Visible = False
            End If

            If .IsPhysicallyDisabled Then
                chkDisabled.value = vbChecked
                txtDisabilityDet.Text = .DisabilityDetails
            Else
                chkDisabled = vbUnchecked
                txtDisabilityDet.Text = ""
            End If
            If .IsOnProbation Then
                chkOnProbation.value = vbChecked
            Else
                chkOnProbation.value = vbUnchecked
            End If
            txtProb.Text = .ProbationPeriod
            
           If DateDiff("d", .DateOfEmployment, .EmploymentValidThrough) < 1 Then
                dtpValidThrough.value = Now()
            Else
                dtpValidThrough.value = .EmploymentValidThrough
            End If
            
            txtPassport.Text = .PassportNo
            txtAlien.Text = .AlienNo
            If .Religion.Religion <> "" Then
                CboReligion.Text = .Religion.Religion
            Else
                CboReligion.Text = "(Unspecified)"
            End If
            txtKRAFileNO.Text = .KRAFileNo
            If .MaritalStatusStr <> "" Then
                cboMarritalStat.Text = .MaritalStatusStr
            Else
                cboMarritalStat.ListIndex = 0
            End If
            If .ProbationTypeStr <> "" Then
                cboProbType.Text = .ProbationTypeStr
            Else
                cboProbType.ListIndex = -1
            End If
    
            lblSDate.Visible = False
            dtpPSDate.Visible = False
            lblCDate.Visible = False
            dtpCDate.Visible = False
                                    
            If .ProbationType = Appointment Then
                lblSDate.Visible = False
                dtpPSDate.Visible = False
                lblCDate.Visible = True
                dtpCDate.Visible = True
               
                dtpCDate.value = .ConfirmationDate
            
            ElseIf .ProbationType = Promotion Then
                lblSDate.Visible = True
                dtpPSDate.Visible = True
                lblCDate.Visible = True
                dtpCDate.Visible = True
                
                dtpPSDate.value = .ProbationStartDate
                dtpCDate = .ConfirmationDate
            End If
            
'            If .IsDisengaged Then
'                cboTermReasons.Text = .disengagementReason
'                If cboTermReasons = "Retirement" Then
'                    fraTerm.Visible = True
'
'                    dtpTerminalDate.Value = Format(.TrainingDate, Dfmt)
'                    txtAdvisor.Text = .TrainingAdvisor
'                    chkAchieved.Value = setCheckBoxes(.TrainingAchieved)
'                End If
'            Else
                
'            End If
                    
            'chkPension.Value = setCheckBoxes(!Pension & "")
            
            Set Picture1 = Nothing
    
            On Error Resume Next 'this handler is specific to the photos only
            ''Picture1.Picture = LoadPicture(App.Path & "\PHOTOS\" & CompanyId & "-" & SelectedEmployee.EmpCode & ".jpg")
            If photoisactive Then
                Set AllEmployeesPhotos = New EmployeesPhotos
               
                Set empphoto = AllEmployeesPhotos.getEmployeePhoto(SelectedEmployee.EmployeeID)
                
    '            If Not AllEmployeesPhotos Is Nothing Then
    '            Set empphoto = AllEmployeesPhotos.FindEmployeePhoto(SelectedEmployee.EmployeeID)
                If Not empphoto Is Nothing Then
                Picture1.Picture = empphoto.Photo
                End If
            
            End If
            ''End If
            
'            If Picture1.Picture = 0 Then
'                On Error Resume Next
'                Picture1.Picture = LoadPicture(App.Path & "\Pic\Gen.jpg")
'            End If
            
            'populate Locations and Currencies
            'In order to populate Locations, you first need to select a country
            If Not (.Location Is Nothing) Then
                If Not (.Location.Country Is Nothing) Then
                    For iCounter = 0 To cboCountry.ListCount - 1
                        If cboCountry.ItemData(iCounter) = .Location.Country.CountryID Then
                            cboCountry.ListIndex = iCounter 'this will trigger population of locations of this country
                            Exit For
                        End If
                    Next iCounter
                                        
                    'now the Location is available in the combo box
                    For jCounter = 0 To cboLocation.ListCount - 1
                        If cboLocation.ItemData(jCounter) = .Location.LocationID Then
                            cboLocation.ListIndex = jCounter
                            Exit For
                        End If
                    Next jCounter
                End If
            Else
                cboLocation.Text = "(Unspecified)"
            End If
            'DISPLAYING THE EMPLOYEE PROGRAMME FUNDING
            objEmployeeProgrammeFundings.GetActiveEmployeeProgrammes
            LoadEmployeeProgrammeFundings SelectedEmployee.EmployeeID
            
            If Not (.CurrencyType Is Nothing) Then
                For iCounter = 0 To cboCurrency.ListCount - 1
                    If cboCurrency.ItemData(iCounter) = .CurrencyType.CurrencyID Then
                        cboCurrency.ListIndex = iCounter
                        Exit For
                    End If
                Next iCounter
            End If
            
        End With
        Call NumberReengaged(SelectedEmployee)
        Call GetBanks
        
        'now populate the employee JDs: Two Steps
        '1. First load the Collection into the Temporary Buffer
        Set TempEmpJDs = SelectedEmployee.JobDescriptionValues
        
        '2. Now populate the Listview and work with the local Buffer: Loads where Deleted=False
        Dim k
        k = Tie_to_category_values(selJDField.JDCategoryID, TempEmpJDs)
        Call PopulateEmployeeJDs(TempEmpJDs)
        'NB: The Synchronization will take place when Inserting an employee or Updating the employee
        
        'now populate the next of kins
        Set TempNextOfKins = SelectedEmployee.NextOfKins
        
        Call PopulateNextOfKins(TempNextOfKins)
        'Force the Command Buttons in JD and Next Of Kin to be disabled
        EnableDisableJDCommands False
        
    End If
'    fraTerm.Visible = True
'    fraTerm.Enabled = True
    Exit Sub
ErrorHandler:
    MsgBox "An error has occurred while Displaying Employee Info" & vbNewLine & err.Description, vbInformation, TITLES
End Sub

Private Sub EnableDisableJDCommands(ByVal EnableThem As Boolean)
    On Error GoTo ErrorHandler
    
    'ENSURING THAT THE JD AND NOK RECORD MANIPULATION COMMAND BUTTONS ARE ENABLED
'    With Me
'        .cmdAddJDValue.Visible = True
'        .cmdEditJDValue.Visible = True
'        .cmdDeleteJDValue.Visible = True
'        .cmdNoKAdd.Visible = True
'        .cmdNoKDelete.Visible = True
'        .cmdNOKEdit.Visible = True
'    End With
    If EnableThem = True Then
        Me.cmdAddJDValue.Enabled = True
        Me.cmdEditJDValue.Enabled = True
        Me.cmdDeleteJDValue.Enabled = True
        
        'those for next of kin
        Me.cmdNoKAdd.Enabled = True
        Me.cmdNoKDelete.Enabled = True
        Me.cmdNOKEdit.Enabled = True
        
    Else
        Me.cmdAddJDValue.Enabled = False
        Me.cmdEditJDValue.Enabled = False
        Me.cmdDeleteJDValue.Enabled = False
        
        'those for next of kin
        Me.cmdNoKAdd.Enabled = False
        Me.cmdNoKDelete.Enabled = False
        Me.cmdNOKEdit.Enabled = False
        
    End If
    
    Exit Sub
    
ErrorHandler:
    
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
        ElseIf TypeOf i Is CheckBox Then
            i.value = vbUnchecked
        End If
    Next i
    
    ClearControlsNextOfKin
    'clear the listview data
    Me.lvwOU.ListItems.Clear
    Me.lvwEmployeeProgrammeFunding.ListItems.Clear
End Sub

Public Sub EnableCmd()
    Dim i As Object
    
    On Error GoTo ErrorHandler
    For Each i In Me
        If TypeOf i Is CommandButton Then
            i.Enabled = True
        End If
    Next i
    
    'Force the JD and Next Of Kin Command Buttons to be Disabled
    'By Oscar: 2007.06.07
    EnableDisableJDCommands False
    
    Exit Sub
    
ErrorHandler:
    
End Sub

Public Sub DisableCmd()
    Dim i As Object
    
    On Error Resume Next
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
    
    'Disable Command buttons for JD and Next Of Kin
    EnableDisableJDCommands False
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
    
    'enable the command buttons for JD and Next of Kin
    EnableDisableJDCommands True
    
End Sub

Public Sub InitGrid()
'    With lvwEmp
'        .ColumnHeaders.Clear
'        .ColumnHeaders.Add , , "Employee Code", 1700
'        .ColumnHeaders.Add , , "Surname", 2000
'        .ColumnHeaders.Add , , "Other Names", 3500
'        .ColumnHeaders.Add , , "ID No", 2500
'        .ColumnHeaders.Add , , "Gender"
'        .ColumnHeaders.Add , , "Date of Birth", 2000
'        .ColumnHeaders.Add , , "Date Employed", 2000
'    '    .ColumnHeaders.Add , , "Tel No", 2500
'    '    .ColumnHeaders.Add , , "Home Address", 4500
'    '    .ColumnHeaders.Add , , "E-Mail", 3000
'    '    .ColumnHeaders.Add , , "Previous Employer", 4000
'        .ColumnHeaders.Add , , "Division Code", 1700
'        .ColumnHeaders.Add , , "Division Name", 4000
'        .ColumnHeaders.Add , , "Terms"
'        .ColumnHeaders.Add , , "Employee Type", 2500
'        .ColumnHeaders.Add , , "PIN No"
'        .ColumnHeaders.Add , , "N.S.S.F No"
'        .ColumnHeaders.Add , , "N.H.I.F No"
'        .ColumnHeaders.Add , , "KRA File No"
'
'
'        .View = lvwReport
'
'    End With

End Sub

Public Sub LoadEmployees()

'    Dim i As Integer
'    Dim j As Long
'    Dim itemX As ListItem
'    Dim Emp As  HRCORE.Employee
'
'    On Error GoTo errorHandler
'
'    lvwEmp.ListItems.Clear
'
'    For j = 1 To AllEmployees.Count
'        Set Emp = AllEmployees.Item(j)
'        Set itemX = Me.lvwEmp.ListItems.Add(, , Emp.empcode, , i)
'        itemX.SubItems(1) = Emp.SurName
'        itemX.SubItems(2) = Emp.OtherNames
'        itemX.SubItems(3) = Emp.IdNo
'        itemX.SubItems(4) = Emp.DateOfBirth
'        itemX.SubItems(5) = Emp.DateOfEmployment
'        itemX.SubItems(6) = Emp.OrganizationUnit.OrganizationUnitCode
'        itemX.SubItems(7) = Emp.OrganizationUnit.OrganizationUnitname
'        itemX.SubItems(8) = Emp.EmploymentTerm.EmpTermName
'        itemX.SubItems(9) = Emp.EmployeeType
'        itemX.SubItems(10) = Emp.PinNo
'        itemX.SubItems(11) = Emp.NssfNo
'        itemX.Tag = Emp.EmployeeID
'    Next j
'
'    Exit Sub
'errorHandler:
'    MsgBox "An error has occurred while populating employees" & vbNewLine & Err.Description, vbInformation, TITLES
'
End Sub

Private Sub Form_Resize()
    oSmart.FResize Me
    'Me.FraEdit.Move FraEdit.Left, FraEdit.Top, FraEdit.Width, tvwMainheight - 220
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs = Nothing
    If Not (currUser Is Nothing) Then
    frmMain2.Caption = "Infiniti HRMIS - PDR [Current User:\" & currUser.FullNames & "]"
    End If
    
End Sub

Private Sub imgDeletePic_Click()
    On Error Resume Next
    Dim Picpath As String
    If Me.txtEmpCode.Locked Then Exit Sub
   If Not SelectedEmployee Is Nothing Then
        If Picture1.Picture <> 0 Then
            If MsgBox("Are you sure that you want to delete the employee photo?", vbQuestion + vbYesNo) = vbYes Then
'                Picpath = App.Path & "\PHOTOS\" & CompanyId & "-" & SelectedEmployee.EmpCode & ".jpg"
'                Kill Picpath
                
                Dim CMD As ADODB.Command
                Set CMD = New ADODB.Command
    
                CMD.ActiveConnection = con
                CMD.CommandText = "pdrspDeleteEmployeesPhoto"
                CMD.CommandType = adCmdStoredProc
                
                CMD.Parameters.Append CMD.CreateParameter(, adInteger, adParamInput, , SelectedEmployee.EmployeeID) 'EmployeeID
                CMD.Execute
                
                Set Picture1 = Nothing
                
                
               '' Picture1.Picture = LoadPicture()
            End If
        End If
    End If
    
    
    
    
    
    
End Sub

Private Sub imgLoadPic_Click()
    If txtEmpCode.Locked Then Exit Sub
    Call cmdPNew_Click
   
End Sub

Private Sub lvwEmp_DblClick()
If frmMain2.cmdEdit.Enabled = True And frmMain2.fracmd.Visible = True Then
    Call frmMain2.cmdEdit_Click
End If
    
End Sub

Private Sub lvwEmployeeProgrammeFunding_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim lngLoopVariable As Long
    On Error GoTo ErrorHandler
    
    For lngLoopVariable = 0 To Me.cboProgrammes.ListCount - 1
        If UCase(Me.cboProgrammes.List(lngLoopVariable)) = UCase(Item.Text) Then
            Me.cboProgrammes.ListIndex = lngLoopVariable
            Exit For
        End If
    Next
    For lngLoopVariable = 0 To Me.cboFundCodes.ListCount - 1
        If UCase(Me.cboFundCodes.List(lngLoopVariable)) = UCase(Item.SubItems(1)) Then
            Me.cboFundCodes.ListIndex = lngLoopVariable
            Exit For
        End If
    Next
    Me.txtEmployeeProgrammeFundingPercentage.Text = Item.SubItems(2)
    Exit Sub
    
ErrorHandler:
    MsgBox "An Error has occurred while attempting to display the employee programe details" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub lvwJDValues_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo ErrorHandler
    
    Me.txtEmpJDValue.Text = ""
    
    Set selEmpJD = Nothing
    
    If IsNumeric(Item.Tag) Then
        Set selEmpJD = TempEmpJDs.FindEmployeeJDByID(CLng(Item.Tag))
        If Not (selEmpJD Is Nothing) Then
            Me.txtEmpJDValue.Text = selEmpJD.FieldValue
        Else
        Set selEmpJD = New EmployeeJD
        selEmpJD.FieldValue = Item.SubItems(1)
        selEmpJD.Deleted = True
        Me.txtEmpJDValue.Text = selEmpJD.FieldValue
        End If
    End If
    
    
    
    Exit Sub
    
ErrorHandler:
    
End Sub

Private Sub lvwNextOfKins_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo ErrorHandler
    
    Set selNextOfKin = Nothing
    If IsNumeric(Item.Tag) Then
        Set selNextOfKin = TempNextOfKins.FindNextOfKinByID(CLng(Item.Tag))
    End If
    Item.Selected = True
    SetNextOfKinFields selNextOfKin
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while processing the Selected Next of Kin" & vbNewLine & err.Description, vbExclamation, TITLES
    
    
End Sub

Private Sub SetNextOfKinFields(ByVal TheNextOfKin As HRCORE.NextOfKin)
    Dim i As Long
    
    On Error GoTo ErrorHandler
    ClearNextOfKinFields
    If Not (TheNextOfKin Is Nothing) Then
        Me.txtNoKAddress.Text = TheNextOfKin.PostalAddress
        Me.txtNoKBenefitPercent.Text = TheNextOfKin.BenefitPercent
        Me.txtNoKEMail.Text = TheNextOfKin.EMail
        Me.txtNoKIDNo.Text = TheNextOfKin.IdNo
        Me.txtNoKOccupation.Text = TheNextOfKin.Occupation
        Me.txtNoKOtherNames.Text = TheNextOfKin.OtherNames
        Me.txtNoKSurname.Text = TheNextOfKin.SurName
        Me.txtNoKRelationship.Text = TheNextOfKin.Relationship
        Me.txtNoKTelephone.Text = TheNextOfKin.Telephone
        Me.dtpNoKDOB.value = TheNextOfKin.DateOfBirth
        
        If TheNextOfKin.IsBeneficiary Then
            Me.chkNoKBeneficiary.value = vbChecked
        Else
            Me.chkNoKBeneficiary.value = vbUnchecked
        End If
        
        If TheNextOfKin.IsDependant Then
            Me.chkNoKDependant.value = vbChecked
        Else
            Me.chkNoKDependant.value = vbUnchecked
        End If
        
        If TheNextOfKin.IsEmergencyContact Then
            Me.chkNoKEmergency.value = vbChecked
        Else
            Me.chkNoKEmergency.value = vbUnchecked
        End If
        
        
        'first deselect all
        Me.cboNoKGender.ListIndex = -1
        
        'then search for a particular gender
        For i = 0 To Me.cboNoKGender.ListCount - 1
            If UCase(Me.cboNoKGender.List(i)) = UCase(TheNextOfKin.GenderStr) Then
                cboNoKGender.ListIndex = i
                Exit For
            End If
        Next i
        With frmNextOfKinGuardian
            .txtFullNames = TheNextOfKin.GuardianFullNames
            .txtIDNo = TheNextOfKin.GuardianIDNo
            .txtRelationship = TheNextOfKin.GuardianRelationship
        End With
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while Displaying details of the selected Next Of Kin" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub ClearNextOfKinFields()
    On Error GoTo ErrorHandler
    Me.txtNoKAddress.Text = ""
    Me.txtNoKBenefitPercent.Text = ""
    Me.txtNoKEMail.Text = ""
    Me.txtNoKIDNo.Text = ""
    Me.txtNoKOccupation.Text = ""
    Me.txtNoKOtherNames.Text = ""
    Me.txtNoKSurname.Text = ""
    Me.txtNoKRelationship.Text = ""
    Me.txtNoKTelephone.Text = ""
    
    Exit Sub
ErrorHandler:
End Sub

Private Sub tvwJDFields_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error GoTo ErrorHandler
    
    Set selJDField = Nothing
    If IsNumeric(Node.Tag) Then
        Set selJDField = pJDFields.FindJDCategoryByID(CLng(Node.Tag))
    End If
    
    If Not (selJDField Is Nothing) Then
        'get the empjds of this jdfield
        Set FilteredEmpJDs = TempEmpJDs.GetEmployeeJDSByJDCategoryID(selJDField.JDCategoryID)
      Dim r
       r = Tie_to_category_values(selJDField.JDCategoryID, FilteredEmpJDs)
        
        PopulateEmployeeJDs FilteredEmpJDs
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while Retrieving the Employee JDs of the Selected Field" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub
Private Function Tie_to_category_values(cid As Integer, tempjds As HRCORE.EmployeeJDs)
On Error GoTo err
Dim sq As String
Dim rs As New Recordset
Dim jd As EmployeeJD
If Not SelectedEmployee.position Is Nothing Then
    sq = "select fieldvalue from JobPositionJDs where JDCategoryID=" & cid & " and positionid=" & SelectedEmployee.position.PositionID & ""
    Set rs = CConnect.GetRecordSet(sq)
    
    If Not rs.EOF Then
        While Not rs.EOF
            Set jd = New EmployeeJD
            Set jd.Employee = New Employee
            jd.Employee.EmployeeID = SelectedEmployee.EmployeeID
            jd.Deleted = True
            Set jd.JDCategory = New JDCategory
            jd.JDCategory.JDCategoryID = cid
            jd.FieldValue = rs!FieldValue
            jd.EmployeeJDID = 0
            tempjds.add jd
            rs.MoveNext
        Wend
    
    End If
End If
err:
Exit Function
End Function
Private Sub txtBasicPay_LostFocus()
    On Error Resume Next
    txtBasicPay.Text = Format(txtBasicPay.Text & "", Cfmt)
End Sub

Private Sub txtEmpCode_Change()
    txtEmpCode.Text = UCase(txtEmpCode.Text)
    txtEmpCode.SelStart = Len(txtEmpCode.Text)
End Sub


Private Sub txtHAllow_LostFocus()
    On Error Resume Next
    txtHAllow.Text = Format(txtHAllow.Text & "", Cfmt)
End Sub

Private Sub txtIDNo_Change()
    txtIDNo.Text = UCase(txtIDNo.Text)
    txtIDNo.SelStart = Len(txtIDNo.Text)
End Sub

Private Sub txtNhif_Change()
    txtNhif.Text = UCase(txtNhif.Text)
    txtNhif.SelStart = Len(txtNhif.Text)
End Sub



Private Sub txtNssf_Change()
    txtNssf.Text = UCase(txtNssf.Text)
    txtNssf.SelStart = Len(txtNssf.Text)
End Sub




Private Sub txtPin_Change()
    txtPin.Text = UCase(txtPin.Text)
    txtPin.SelStart = Len(txtPin.Text)
End Sub


Private Sub txtProb_Change()
    If cboProbType.Text = "Appointment" Then
        dtpCDate.value = DateAdd("m", Val(txtProb.Text), dtpDEmployed.value)
    ElseIf cboProbType.Text = "Promotion" Then
        dtpCDate.value = DateAdd("m", Val(txtProb.Text), dtpPSDate.value)
        
    End If
    
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

   
    If DateDiff("d", dtpDOB.value, Date) < 0 Then
        MsgBox "Date of birth cannot be in the future. Enter the correct date.", vbInformation
        dtpDOB.value = Date
        'txtDOB.Text = Date
        dtpDOB.SetFocus
        Exit Sub
    End If
    
    'If txtDEmployed.Text <> "" Then
    If DateDiff("d", dtpDEmployed.value, dtpDOB.value) > 0 Then
        MsgBox "Date birth cannot be greater than date employed. Enter correct dates.", vbInformation
        dtpDOB.value = Date
        'txtDOB.Text = Date
        dtpDOB.SetFocus
        Exit Sub
    End If
    'End If
        
        
End Sub

Private Sub CheckDEmployed()
    Dim ddf As Long
        
    If DateDiff("d", dtpDEmployed.value, dtpDOB.value) > 0 Then
        MsgBox "Date birth cannot be greater than date employed. Enter correct dates.", vbInformation
        dtpDEmployed.value = Date
        'txtDEmployed.Text = Date
        dtpDEmployed.SetFocus
        Exit Sub
    End If
     
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
    If DSource = "Local" Or (TheLoadedForm.Name = "frmEmployee") Then
        With frmMain2
            .cmdNew.Enabled = False
            .cmdDelete.Enabled = False
            .cmdEdit.Enabled = False
            .cmdSave.Enabled = True
            .cmdCancel.Enabled = True
        End With
    End If
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
    
    
    If txtProb.Text = "" Then
        txtProb.Text = 0
    End If
    
    If cboProbType = "Promotion" Then
        PromptDate = DateAdd("m", Val(txtProb.Text), dtpPSDate.value)
    Else
        PromptDate = DateAdd("m", Val(txtProb.Text), dtpDEmployed.value)
    End If
    validateEmployeeData = True
End Function


'===========  HRCORE CODE ===========
Private Sub cboOU_Click()
    Dim theOU As New HRCORE.OrganizationUnit
    Dim colReps As New HRCORE.OrganizationUnits
    txtOUInfo.Text = ""
    If cboOU.ListIndex >= 0 Then
        Set theOU = OUnits.FindOrganizationUnit(cboOU.ItemData(cboOU.ListIndex))
        If Not (theOU Is Nothing) Then
            Set selOU = theOU
            
            Me.txtOUInfo.Text = "OU Type: " & theOU.OUType.OUTypeName & vbNewLine
            
            'If (theOU.ParentOU Is Nothing) Or (theOU.ParentOU.OrganizationUnitID <= 0) Then
                'Me.txtOUInfo.Text = Me.txtOUInfo.Text & "Parent OU: " & company.CompanyName
            'Else
                'Me.txtOUInfo.Text = Me.txtOUInfo.Text & "Parent OU: " & theOU.ParentOU.OrganizationUnitName
            'End If
            Me.txtOUInfo.Text = Me.txtOUInfo.Text & "Hierarchy: " & LCase(OUnits.GetOUFamilyTree(theOU))
            
            'make sure chkOUVisibility is enabled
            chkHasOUV.Enabled = True
            cmdSearchEOUV.Enabled = True
        End If
    End If
End Sub

Private Sub LoadOUReplicas(ByVal TheReplicas As OrganizationUnits)
    Dim ItemX As ListItem
    Dim i As Long
    Dim par As OrganizationUnit
        
    For i = 1 To TheReplicas.count
        Set par = TheReplicas.Item(i).ParentOU
        'display data of the parents
        Set ItemX = lvwOU.ListItems.add(, , par.OrganizationUnitName)
        ItemX.SubItems(1) = LCase(OUnits.GetOUFamilyTree(par))
        'store the ID of the Replica
        ItemX.Tag = TheReplicas.Item(i).OrganizationUnitID
    Next i
End Sub


Private Sub chkHasOUV_Click()
    If chkHasOUV.value = vbChecked Then
       
        Me.cmdSearchEOUV.Enabled = True
        Me.cmdSearchEOUV.Visible = True
    Else
        Me.cmdSearchEOUV.Enabled = False
        Me.cmdSearchEOUV.Visible = True
    End If
End Sub


Private Sub LoadOUTypes()
    Dim myOUT As HRCORE.OrganizationUnitType
   '' If (outs Is Nothing) Then
    outs.GetAllOUTypes
   '' End If
End Sub

Private Sub LoadEmployeesPhotos()
   AllEmployeesPhotos.GetAccessibleEmployeesPhotosByUser currUser.UserID
   
End Sub


Private Sub LoadOrganizationUnits()
    Dim myOU As HRCORE.OrganizationUnit
   
    Dim i As Long
   '' If (OUnits Is Nothing) Then
    OUnits.GetAllOrganizationUnits
   '' End If
    cboOU.Clear
    
    'Add Unspecified
    cboOU.AddItem "(Unspecified)"
    For i = 1 To OUnits.count
        Set myOU = OUnits.Item(i)
        
        'Force the OU Type details to be loaded
        If Not (myOU.OUType Is Nothing) Then
            Set myOU.OUType = outs.FindOUType(myOU.OUType.OUTypeID)
        End If
        
        'Force ParentOU to be loaded
        If Not (myOU.ParentOU Is Nothing) Then
            Set myOU.ParentOU = OUnits.FindOrganizationUnit(myOU.ParentOU.OrganizationUnitID)
        End If
        
        cboOU.AddItem myOU.OrganizationUnitName
        cboOU.ItemData(cboOU.NewIndex) = myOU.OrganizationUnitID
    Next i
        
End Sub


Private Sub LoadEmployeeCategories()
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    cboCat.Clear
    
    'add unspecified
    cboCat.AddItem "(Unspecified)"
   '' Set EmpCats = New  HRCORE.EmployeeCategories
    ''EmpCats.GetAllEmployeeCategories
    For i = 1 To EmpCats.count
        cboCat.AddItem EmpCats.Item(i).CategoryName
        cboCat.ItemData(cboCat.NewIndex) = EmpCats.Item(i).CategoryID
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while Populating Employee Categories" & vbNewLine & err.Description, vbInformation, TITLES
End Sub

Private Sub LoadEmploymentTerms()
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    cboTerms.Clear
   '' empTerms.GetAllEmploymentTerms
    
    'add unspecified
    cboTerms.AddItem "(Unspecified)"
    For i = 1 To empTerms.count
        cboTerms.AddItem empTerms.Item(i).EmpTermName
        cboTerms.ItemData(cboTerms.NewIndex) = empTerms.Item(i).EmpTermID
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while populating Employment Terms" & vbNewLine & err.Description, vbInformation, TITLES
    
End Sub

Private Sub LoadNationalities()
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    cboNationality.Clear
    
   '' If (empNationalities Is Nothing) Then
    empNationalities.GetAllNationalities
   '' End If
    'add unspecified
    cboNationality.AddItem "(Unspecified)"
    
    For i = 1 To empNationalities.count
        cboNationality.AddItem empNationalities.Item(i).Nationality
        cboNationality.ItemData(cboNationality.NewIndex) = empNationalities.Item(i).NationalityID
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while populating the Nationalities" & vbNewLine & err.Description, vbInformation, TITLES
    
End Sub

Private Sub LoadTribes()
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    cboTribe.Clear
    ''If (empTribes Is Nothing) Then
    empTribes.GetAllTribes
   '' End If
    'add Unspecified
    cboTribe.AddItem "(Unspecified)"
    
    For i = 1 To empTribes.count
        cboTribe.AddItem empTribes.Item(i).Tribe
        cboTribe.ItemData(cboTribe.NewIndex) = empTribes.Item(i).TribeID
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while populating Tribes" & vbNewLine & err.Description, vbInformation, TITLES
    
End Sub

Private Sub LoadReligions()
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    CboReligion.Clear
   '' If (empReligions Is Nothing) Then
    empReligions.GetAllReligions
   '' End If
    'add unspecified
    CboReligion.AddItem "(Unspecified)"
    
    For i = 1 To empReligions.count
        CboReligion.AddItem empReligions.Item(i).Religion
        CboReligion.ItemData(CboReligion.NewIndex) = empReligions.Item(i).ReligionID
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while populating Religions" & vbNewLine & err.Description, vbInformation, TITLES
    
End Sub


Private Sub LoadJobPositions()
    Dim i As Long
    Dim jpos As HRCORE.JobPosition
    
    On Error GoTo ErrorHandler
    
    cboDesig.Clear
   '' If (empPositions Is Nothing) Then
    empPositions.GetAllJobPositions
   '' End If
    'add unspecified
    cboDesig.AddItem "(Unspecified)"
    For i = 1 To empPositions.count
        cboDesig.AddItem empPositions.Item(i).PositionName
        cboDesig.ItemData(cboDesig.NewIndex) = empPositions.Item(i).PositionID
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while populating Job Positions" & vbNewLine & err.Description, vbInformation, TITLES
    
End Sub

Private Sub LoadCurrencies()
    Dim i As Long
    
    On Error GoTo ErrorHandler
    Me.cboCurrency.Clear
   '' If (empCurrencies Is Nothing) Then
    empCurrencies.GetActiveCurrencies
  ''  End If
    cboCurrency.AddItem "(Unspecified)"
    For i = 1 To empCurrencies.count
        Me.cboCurrency.AddItem empCurrencies.Item(i).CurrencyCode & " (" & empCurrencies.Item(i).CurrencySymbol & ")"
        Me.cboCurrency.ItemData(Me.cboCurrency.NewIndex) = empCurrencies.Item(i).CurrencyID
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while Loading the Employee Currencies" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub LoadCountries()
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    cboCountry.Clear
    ''If (empCountries Is Nothing) Then
    empCountries.GetActiveCountries
   '' End If
    cboCountry.AddItem "(Unspecified)"
    For i = 1 To empCountries.count
        cboCountry.AddItem empCountries.Item(i).CountryName
        cboCountry.ItemData(cboCountry.NewIndex) = empCountries.Item(i).CountryID
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while populating the countries" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub LoadCSSSCategories()
    Dim i As Long
    On Error GoTo ErrorHandler
    
    cboStaffCategory.Clear
    ''If (empStaffCategories Is Nothing) Then
    empStaffCategories.GetActiveCSSSCategories
   '' End If
    cboStaffCategory.AddItem "(Unspecified)"
    For i = 1 To empStaffCategories.count
        cboStaffCategory.AddItem empStaffCategories.Item(i).CSSSCategoryName
        cboStaffCategory.ItemData(cboStaffCategory.NewIndex) = empStaffCategories.Item(i).CSSSCategoryID
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while populating the Staff Categories" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub LoadLocationsOfCountry(ByVal TheCountry As HRCORE.Country)
    Dim i As Long
    Dim FilteredLocations As HRCORE.Locations
    
    On Error GoTo ErrorHandler
    
    cboLocation.Clear
    cboLocation.AddItem "(Unspecified)"
    
    If TheCountry Is Nothing Then
        Exit Sub
    End If
    
    Set FilteredLocations = empLocations.GetLocationsByCountryID(TheCountry.CountryID)
    If FilteredLocations Is Nothing Then
        Exit Sub
    End If
    
    For i = 1 To FilteredLocations.count
        cboLocation.AddItem FilteredLocations.Item(i).LocationName
        cboLocation.ItemData(cboLocation.NewIndex) = FilteredLocations.Item(i).LocationID
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while processing Locations in the selected Country" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub LoadGradesOfCSSSCategory(ByVal TheCSSSCategory As HRCORE.CSSSCategory)
    Dim FilteredEmpCats As HRCORE.EmployeeCategories
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    cboCat.Clear
    cboCat.AddItem "(Unspecified)"
    
    If TheCSSSCategory Is Nothing Then
   
        Exit Sub
    End If
    
    ''****************************
    Set EmpCats = New HRCORE.EmployeeCategories
    EmpCats.GetActiveEmployeeCategories
    ''*****************
    
    Set FilteredEmpCats = EmpCats.GetEmployeeCategoriesByCSSSCategoryID(TheCSSSCategory.CSSSCategoryID)
    If FilteredEmpCats Is Nothing Then
        Exit Sub
    End If
    
    For i = 1 To FilteredEmpCats.count
        cboCat.AddItem FilteredEmpCats.Item(i).CategoryName
        cboCat.ItemData(cboCat.NewIndex) = FilteredEmpCats.Item(i).CategoryID
    Next i
    
    For i = 0 To cboStaffCategory.ListCount - 1
   
            If cboStaffCategory.ItemData(i) = TheCSSSCategory.CSSSCategoryID Then
            cboStaffCategory.ListIndex = i
            Exit For
            End If
    Next i
 
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while populating Employee Grades" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub LoadLocations()
   Dim i As Long
   
   On Error GoTo ErrorHandler
    cboLocation.Clear
    
    cboLocation.AddItem "(Unspecified)"
    ''If (empLocations Is Nothing) Then
    empLocations.GetActiveLocations
   '' End If
    For i = 1 To empLocations.count
        cboLocation.AddItem empLocations.Item(i).LocationName
        cboLocation.ItemData(cboLocation.NewIndex) = empLocations.Item(i).LocationID
    Next i
    
    Exit Sub
    
ErrorHandler:
 
End Sub

Private Sub LoadProgrammes()
    Dim lngLoopVariable As Long
    On Error GoTo ErrorHandler
    
    Me.cboProgrammes.Clear
    For lngLoopVariable = 1 To objProgrammes.count
        Me.cboProgrammes.AddItem objProgrammes.Item(lngLoopVariable).ProgrammeName
        Me.cboProgrammes.ItemData(Me.cboProgrammes.NewIndex) = objProgrammes.Item(lngLoopVariable).ProgrammeID
    Next
    If Me.cboProgrammes.ListCount > 0 Then
        Me.cboProgrammes.ListIndex = 0
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "An Error has occurred while attempting to load all the programmes on the employee general details" & vbNewLine, vbExclamation, TITLES
End Sub

Private Sub LoadFundCodesForSpecificProgramme(ByVal TheProgrammeID As Long)
    Dim lngLoopVariable As Long
    On Error GoTo ErrorHandler
    
    Me.cboFundCodes.Clear
    For lngLoopVariable = 1 To objProgrammeFundings.count
        If objProgrammeFundings.Item(lngLoopVariable).Programme.ProgrammeID = TheProgrammeID Then
            Me.cboFundCodes.AddItem objProgrammeFundings.Item(lngLoopVariable).FundCode.FundCodeName
            Me.cboFundCodes.ItemData(Me.cboFundCodes.NewIndex) = objProgrammeFundings.Item(lngLoopVariable).FundCode.FundCodeID
            If blnDisplayEmployeeProgrammeFundingInfo Then
                If Me.cboFundCodes.ItemData(Me.cboFundCodes.NewIndex) = SelectedEmployeeProgrammeFunding.ProgrammeFunding.FundCode.FundCodeID Then
                    Me.cboFundCodes.ListIndex = Me.cboFundCodes.NewIndex
                    blnDisplayEmployeeProgrammeFundingInfo = Not blnDisplayEmployeeProgrammeFundingInfo
                End If
            End If
        End If
    Next
    Exit Sub
    
ErrorHandler:
    MsgBox "An Error has occurred while attempting to load all the fund codes on the employee general details" & vbNewLine, vbExclamation, TITLES
End Sub

Private Sub LoadEmployeeProgrammeFundings(ByVal TheEmpID As Long)
    Dim myinternalEmployeeProgrammeFunding As EmployeeProgramme
    Dim myListItem As ListItem
    Dim lngLoopVariable As Long
    On Error GoTo ErrorHandler
   
    'CLEARING THE LIST VIEW CONTROL
    Me.lvwEmployeeProgrammeFunding.ListItems.Clear
    Me.txtEmployeeProgrammeFundingPercentage.Text = vbNullString
    'NOW LOADING EMPLOYEE PROGRAMMES FOR THE SELECTED EMPLOYEE
    For lngLoopVariable = 1 To objEmployeeProgrammeFundings.count
        Set myinternalEmployeeProgrammeFunding = objEmployeeProgrammeFundings.Item(lngLoopVariable)
        If myinternalEmployeeProgrammeFunding.Employee.EmployeeID = TheEmpID And Not myinternalEmployeeProgrammeFunding.Deleted Then
            ''Set myListItem = Me.lvwEmployeeProgrammeFunding.ListItems.Add(, , myinternalEmployeeProgrammeFunding.ProgrammeFunding.Programme.ProgrammeID)
            Set myListItem = Me.lvwEmployeeProgrammeFunding.ListItems.add(, , myinternalEmployeeProgrammeFunding.ProgrammeFunding.Programme.ProgrammeName)
            
            myListItem.SubItems(1) = myinternalEmployeeProgrammeFunding.ProgrammeFunding.FundCode.FundCodeName
            myListItem.SubItems(2) = myinternalEmployeeProgrammeFunding.EmployeeProgrammePercentage
            myListItem.Tag = myinternalEmployeeProgrammeFunding.EmployeeProgrammeID
        End If
    Next
    If Me.lvwEmployeeProgrammeFunding.ListItems.count > 0 Then
        lvwEmployeeProgrammeFunding_ItemClick Me.lvwEmployeeProgrammeFunding.ListItems.Item(1)
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "An Error has occurred while attempting to load all the employee programmes" & vbCrLf & err.Description, vbExclamation, TITLES
End Sub

'Private Sub InsertEmployee()
'    Dim myEmp As  HRCORE.Employee
'    Dim Myeouv As  HRCORE.EmployeeOUVisibility
'    Dim retval As Long
'    Dim i As Long
'
'    Set myEmp = New  HRCORE.Employee
'    myEmp.empcode = Me.txtEmpCode.Text
'    If Me.chkHasOUV.Value = vbChecked Then
'        myEmp.HasOUVisibility = True
'    Else
'        myEmp.HasOUVisibility = False
'    End If
'
'    myEmp.SurName = Me.txtSurname.Text
'    myEmp.OtherNames = Me.txtOtherNames.Text
'    myEmp.VisibleInTheseOUs.Clear
'    If cboOU.ListIndex >= 0 Then
'        Set myEmp.OrganizationUnit = OUnits.FindOrganizationUnit(CLng(Me.cboOU.ItemData(cboOU.ListIndex)))
'
'        If chkHasOUV.Value = vbChecked Then
'            myEmp.HasOUVisibility = True
'            If lvwOU.ListItems.Count > 0 Then
'                For i = 1 To lvwOU.ListItems.Count
'                    If lvwOU.ListItems(i).Checked = True Then
'                        Set Myeouv = New  HRCORE.EmployeeOUVisibility
'                        Set Myeouv.Employee = myEmp
'                        Set Myeouv.OrganizationUnit = OUnits.FindOrganizationUnit(CLng(lvwOU.ListItems(i).Tag))
'                        myEmp.VisibleInTheseOUs.Add Myeouv
'                    End If
'                Next i
'            End If
'        Else
'            myEmp.HasOUVisibility = False
'        End If
'    Else
'        Set myEmp.OrganizationUnit = Nothing
'    End If
'
'    retval = myEmp.InsertNew()
'End Sub
'
'========== END OF  HRCORE CODE ===========

Private Sub GetBanks()
    Dim myinternalEmployeeBankAccounts As EmployeeBankAccounts2
    Dim lngLoopVariable As Long
    Dim blnMainAccountPresent As Boolean
    On Error GoTo ErrorHandler
    
    blnMainAccountPresent = False
    Set myinternalEmployeeBankAccounts = objEmployeeBankAccounts.GetEmployeeBankAccountsOfEmployeeID(SelectedEmployee.EmployeeID)
    If Not myinternalEmployeeBankAccounts Is Nothing And myinternalEmployeeBankAccounts.count > 0 Then
        For lngLoopVariable = 1 To myinternalEmployeeBankAccounts.count
            If myinternalEmployeeBankAccounts.Item(lngLoopVariable).IsMainAccount = True Then
                blnMainAccountPresent = True
                GoTo DisplayTheValues
            End If
        Next
        If Not blnMainAccountPresent Then
            lngLoopVariable = 1
        End If
DisplayTheValues:
        Me.txtAccountName.Text = myinternalEmployeeBankAccounts.Item(lngLoopVariable).AccountName
        Me.txtAccountNO = myinternalEmployeeBankAccounts.Item(lngLoopVariable).AccountNumber
        Me.txtBankBranchName.Text = myinternalEmployeeBankAccounts.Item(lngLoopVariable).bankbranch.BranchName
        Me.txtBankName.Text = myinternalEmployeeBankAccounts.Item(lngLoopVariable).bankbranch.Bank.BankName
    
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "An Error has occurred while attempting to display the employee bank account details" & vbCrLf & err.Description, vbExclamation, TITLES
'    Dim rs As ADODB.Recordset
'    'Get the employee main account
'
'        mySQL = "SELECT EmployeeBanks.AccNumber,EmployeeBanks.AccType,tblBankBranch.BANKBRANCH_NAME,tblBank.BANK_NAME" _
'    & " FROM  EmployeeBanks LEFT OUTER JOIN tblBankBranch ON EmployeeBanks.BranchID = tblBankBranch.BANKBRANCH_ID LEFT OUTER JOIN" _
'    & " tblBank ON tblBankBranch.BANK_ID = tblBank.BANK_ID WHERE EmployeeBanks.MainAcct=1 AND EmployeeBanks.Employee_ID=" & SelectedEmployee.EmployeeID
'    Set rs = CConnect.GetRecordSet(mySQL)
'    If Not (rs Is Nothing) And rs.RecordCount > 0 Then
'        Me.txtAccountName.Text = rs!AccType
'        Me.txtAccountNO = rs!AccNumber
'        Me.txtBankBranchName.Text = rs!BANKBRANCH_NAME
'        Me.txtBankName.Text = rs!Bank_Name
'    Else
'        'get any other bank
'         mySQL = "SELECT EmployeeBanks.AccNumber,EmployeeBanks.AccType,tblBankBranch.BANKBRANCH_NAME,tblBank.BANK_NAME" _
'    & " FROM  EmployeeBanks LEFT OUTER JOIN tblBankBranch ON EmployeeBanks.BranchID = tblBankBranch.BANKBRANCH_ID LEFT OUTER JOIN" _
'    & " tblBank ON tblBankBranch.BANK_ID = tblBank.BANK_ID WHERE  EmployeeBanks.Employee_ID=" & SelectedEmployee.EmployeeID
'        Set rs = CConnect.GetRecordSet(mySQL)
'        If Not (rs Is Nothing) And rs.RecordCount > 0 Then
'            rs.MoveFirst
'            Me.txtAccountName.Text = rs!AccType
'            Me.txtAccountNO = rs!AccNumber
'            Me.txtBankBranchName.Text = rs!BANKBRANCH_NAME
'            Me.txtBankName.Text = rs!Bank_Name
'        Else
'            'Give the user an opportunoty to emter new bank
'            fraDetails.Visible = True
'        End If
'    End If
'    Set rs = Nothing
End Sub

Private Sub NumberReengaged(emp As HRCORE.Employee)
    On Error GoTo errHandler
    Dim i, total As Long
    Dim empid As Long
    Dim Reengaged As New ReengagedEmployees
    Reengaged.GetallReengagedEmployees
    
    total = 0
    empid = emp.EmployeeID
    
    For i = 1 To Reengaged.count
        If Reengaged.Item(i).Employee.EmployeeID = empid Then
            total = total + 1
        End If
    Next i
    If total > 0 Then
        lblDisengaged.Caption = emp.SurName & " " & emp.OtherNames & vbNewLine & "Has been reengaged  " & total & IIf(total > 1, " times", " time") & vbNewLine & "in " & UCase(companyDetail.CompanyName)
        'lblDisengaged.BackColor=QBColor(
    Else
        lblDisengaged.Caption = ""
        
    End If
    Exit Sub
    
errHandler:
    MsgBox "An error has occur : " & err.Description
End Sub

Private Sub LoadJDFields()
    Dim myJD As HRCORE.JDCategory
    Dim myNode As Node
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
       
    'clear lists
    Me.tvwJDFields.Nodes.Clear
        
    pJDFields.GetActiveJDCategories
    
    'first load the Header
    Set myNode = tvwJDFields.Nodes.add(, , "JDFIELDS", "ALL JD FIELDS")
    myNode.Tag = "JDFIELDS"
    myNode.Bold = True
    
    'now get the Top Level JD Fields
    Set TopLevelJDFields = pJDFields.GetTopLevelJDCategories()
    If Not (TopLevelJDFields Is Nothing) Then
        For i = 1 To TopLevelJDFields.count
            Set myJD = TopLevelJDFields.Item(i)
            
            'set numbering
            myJD.FieldNumber = i
                      
            'add the JD
            Set myNode = Me.tvwJDFields.Nodes.add(, , "JD:" & myJD.JDCategoryID, myJD.FieldNumber & ". " & myJD.CategoryName)
            myNode.Tag = myJD.JDCategoryID
            myNode.EnsureVisible
            myNode.Bold = True
            
            'now recursively add the children
            AddChildJDFieldsRecursively myJD
        Next i
         
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "An error has occurred while populating the Sectors" & _
    vbNewLine & err.Description, vbInformation, TITLES
End Sub

Private Sub AddChildJDFieldsRecursively(ByVal TheJD As HRCORE.JDCategory)
    
    'this is a recursive function that populates child JD Fields
    Dim ChildNode As Node
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    If Not (TheJD Is Nothing) Then
        For i = 1 To TheJD.Children.count
            
            'set the field numbering
            TheJD.Children.Item(i).FieldNumber = TheJD.FieldNumber & "." & i
            
            Set ChildNode = Me.tvwJDFields.Nodes.add("JD:" & TheJD.JDCategoryID, tvwChild, "JD:" & TheJD.Children.Item(i).JDCategoryID, TheJD.Children.Item(i).FieldNumber & ". " & TheJD.Children.Item(i).CategoryName)
            ChildNode.Tag = TheJD.Children.Item(i).JDCategoryID
            ChildNode.EnsureVisible
            ChildNode.Bold = True
            'recursively load the children
            AddChildJDFieldsRecursively TheJD.Children.Item(i)
        Next i
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred" & vbNewLine & err.Description, vbInformation, TITLES
        
End Sub

Private Sub LogEmployeeChangesToAuditTrail()
    On Error GoTo errHandler
        
    'Check EmpCode Changes
    If ((OldEmpInfo.EmpCode <> NewEmpInfo.EmpCode) And (SaveNew = False)) Then
        Action = "Changed Employee Code (Staff No.) From " & OldEmpInfo.EmpCode & " to " & NewEmpInfo.EmpCode
        currUser.AuditTrail Update, Action
    End If
     
    Exit Sub
errHandler:
    Debug.Print err.Description
End Sub


