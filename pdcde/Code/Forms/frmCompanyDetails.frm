VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmCompanyDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company Details"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   8760
   Begin VB.Frame Frame1 
      Height          =   6405
      Left            =   180
      TabIndex        =   0
      Top             =   0
      Width           =   8505
      Begin VB.Frame Frame2 
         Height          =   5835
         Left            =   0
         TabIndex        =   14
         Top             =   480
         Width           =   8505
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5580
            TabIndex        =   12
            Top             =   5280
            Width           =   1215
         End
         Begin MSComDlg.CommonDialog cdlgLogo 
            Left            =   2160
            Top             =   5280
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3960
            TabIndex        =   11
            Top             =   5280
            Width           =   1215
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "Close"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7200
            TabIndex        =   13
            Top             =   5280
            Width           =   1215
         End
         Begin VB.TextBox txtCompanyName 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1410
            TabIndex        =   1
            Text            =   "DEFAULT COMPANY"
            Top             =   180
            Width           =   6765
         End
         Begin TabDlg.SSTab sstCompanyInfo 
            Height          =   4635
            Left            =   120
            TabIndex        =   2
            Top             =   600
            Width           =   8325
            _ExtentX        =   14684
            _ExtentY        =   8176
            _Version        =   393216
            Tab             =   2
            TabHeight       =   520
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "General Info"
            TabPicture(0)   =   "frmCompanyDetails.frx":0000
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "fraLogo"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "chkIsNGO"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "txtSMTPServer"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "txtWebURL"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "txtEMail"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "txtFax"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "txtTelephone1"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "txtPostalAddress"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "txtPhysicalAddress"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "txtTelephone2"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "txtTelephone3"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "Label8"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "Label7"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "Label6"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "Label5"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "Label4"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "Label3"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).Control(17)=   "Label2"
            Tab(0).Control(17).Enabled=   0   'False
            Tab(0).Control(18)=   "Label12"
            Tab(0).Control(18).Enabled=   0   'False
            Tab(0).Control(19)=   "Label13"
            Tab(0).Control(19).Enabled=   0   'False
            Tab(0).ControlCount=   20
            TabCaption(1)   =   "Other Info"
            TabPicture(1)   =   "frmCompanyDetails.frx":001C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "txtHelbNumber"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "txtDaysPerMonth"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "txtNHIFNo"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "txtNSSFNo"
            Tab(1).Control(3).Enabled=   0   'False
            Tab(1).Control(4)=   "txtPINNo"
            Tab(1).Control(4).Enabled=   0   'False
            Tab(1).Control(5)=   "txtHoursPerDay"
            Tab(1).Control(5).Enabled=   0   'False
            Tab(1).Control(6)=   "txtHoursPerMonth"
            Tab(1).Control(6).Enabled=   0   'False
            Tab(1).Control(7)=   "Label22"
            Tab(1).Control(7).Enabled=   0   'False
            Tab(1).Control(8)=   "Label21"
            Tab(1).Control(8).Enabled=   0   'False
            Tab(1).Control(9)=   "Label11"
            Tab(1).Control(9).Enabled=   0   'False
            Tab(1).Control(10)=   "Label10"
            Tab(1).Control(10).Enabled=   0   'False
            Tab(1).Control(11)=   "Label9"
            Tab(1).Control(11).Enabled=   0   'False
            Tab(1).Control(12)=   "Label14"
            Tab(1).Control(12).Enabled=   0   'False
            Tab(1).Control(13)=   "Label15"
            Tab(1).Control(13).Enabled=   0   'False
            Tab(1).ControlCount=   14
            TabCaption(2)   =   "Company Banks"
            TabPicture(2)   =   "frmCompanyDetails.frx":0038
            Tab(2).ControlEnabled=   -1  'True
            Tab(2).Control(0)=   "cmdRemoveAccount"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).Control(1)=   "cmdAddAccount"
            Tab(2).Control(1).Enabled=   0   'False
            Tab(2).Control(2)=   "Frame3"
            Tab(2).Control(2).Enabled=   0   'False
            Tab(2).Control(3)=   "cmdEditAccount"
            Tab(2).Control(3).Enabled=   0   'False
            Tab(2).Control(4)=   "fraCompanyBanks"
            Tab(2).Control(4).Enabled=   0   'False
            Tab(2).ControlCount=   5
            Begin VB.Frame fraLogo 
               Height          =   3435
               Left            =   -69720
               TabIndex        =   57
               Top             =   480
               Width           =   2835
               Begin VB.CommandButton cmdLoadLogo 
                  Caption         =   "Logo"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   120
                  TabIndex        =   59
                  Top             =   210
                  Width           =   1245
               End
               Begin VB.CommandButton cmdRemoveLogo 
                  Caption         =   "Remove"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   1440
                  TabIndex        =   58
                  Top             =   210
                  Width           =   1125
               End
               Begin VB.Image imgLogo 
                  Height          =   2535
                  Left            =   120
                  Top             =   720
                  Width           =   2595
               End
            End
            Begin VB.Frame fraCompanyBanks 
               Height          =   2175
               Left            =   120
               TabIndex        =   55
               Top             =   1680
               Width           =   7935
               Begin MSComctlLib.ListView lvwCompanyAccounts 
                  Height          =   1785
                  Left            =   120
                  TabIndex        =   56
                  Top             =   240
                  Width           =   7635
                  _ExtentX        =   13467
                  _ExtentY        =   3149
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   0   'False
                  HideSelection   =   0   'False
                  AllowReorder    =   -1  'True
                  FullRowSelect   =   -1  'True
                  GridLines       =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  Appearance      =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  NumItems        =   5
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "Bank Name"
                     Object.Width           =   2540
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Text            =   "Branch Name"
                     Object.Width           =   2540
                  EndProperty
                  BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   2
                     Text            =   "Account Number"
                     Object.Width           =   2540
                  EndProperty
                  BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   3
                     Text            =   "Account Type"
                     Object.Width           =   2540
                  EndProperty
                  BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   4
                     Text            =   "Default Account"
                     Object.Width           =   2540
                  EndProperty
               End
            End
            Begin VB.CommandButton cmdEditAccount 
               Caption         =   "Edit"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   240
               TabIndex        =   54
               Top             =   3960
               Width           =   1305
            End
            Begin VB.Frame Frame3 
               BorderStyle     =   0  'None
               Caption         =   "Frame3"
               Height          =   1215
               Left            =   120
               TabIndex        =   44
               Top             =   360
               Width           =   8055
               Begin VB.ComboBox cboAccountTypes 
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
                  ItemData        =   "frmCompanyDetails.frx":0054
                  Left            =   5220
                  List            =   "frmCompanyDetails.frx":0064
                  Style           =   2  'Dropdown List
                  TabIndex        =   49
                  Top             =   555
                  Width           =   2625
               End
               Begin VB.TextBox txtAccountNo 
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
                  Left            =   1110
                  TabIndex        =   48
                  Top             =   570
                  Width           =   2625
               End
               Begin VB.ComboBox cboBankBranches 
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
                  Left            =   5220
                  Style           =   2  'Dropdown List
                  TabIndex        =   47
                  Top             =   120
                  Width           =   2625
               End
               Begin VB.ComboBox cboBanks 
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
                  Left            =   1110
                  Style           =   2  'Dropdown List
                  TabIndex        =   46
                  Top             =   120
                  Width           =   2625
               End
               Begin VB.CheckBox chkDefault 
                  Caption         =   "Default Bank"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   45
                  Top             =   960
                  Width           =   2415
               End
               Begin VB.Label Label19 
                  Caption         =   "Account Type"
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
                  Left            =   4140
                  TabIndex        =   53
                  Top             =   615
                  Width           =   1065
               End
               Begin VB.Label Label18 
                  Caption         =   "Account No"
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
                  TabIndex        =   52
                  Top             =   615
                  Width           =   1005
               End
               Begin VB.Label Label17 
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
                  Height          =   195
                  Left            =   4140
                  TabIndex        =   51
                  Top             =   180
                  Width           =   945
               End
               Begin VB.Label Label16 
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
                  Height          =   195
                  Left            =   120
                  TabIndex        =   50
                  Top             =   180
                  Width           =   945
               End
            End
            Begin VB.CheckBox chkIsNGO 
               Caption         =   "This is a Non - Governmental Organization"
               Height          =   195
               Left            =   -73410
               TabIndex        =   34
               Top             =   4080
               Width           =   3495
            End
            Begin VB.TextBox txtSMTPServer 
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
               Left            =   -73410
               TabIndex        =   33
               Top             =   3210
               Width           =   3375
            End
            Begin VB.TextBox txtWebURL 
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
               Left            =   -73410
               TabIndex        =   32
               Top             =   3600
               Width           =   3375
            End
            Begin VB.TextBox txtEMail 
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
               Left            =   -73410
               TabIndex        =   31
               Top             =   2820
               Width           =   3375
            End
            Begin VB.TextBox txtFax 
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
               Left            =   -73410
               TabIndex        =   30
               Top             =   2430
               Width           =   3375
            End
            Begin VB.TextBox txtTelephone1 
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
               Left            =   -73410
               TabIndex        =   29
               Top             =   1320
               Width           =   3375
            End
            Begin VB.TextBox txtPostalAddress 
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
               Left            =   -73410
               TabIndex        =   28
               Top             =   930
               Width           =   3375
            End
            Begin VB.TextBox txtPhysicalAddress 
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
               Left            =   -73410
               TabIndex        =   27
               Top             =   540
               Width           =   3375
            End
            Begin VB.TextBox txtTelephone2 
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
               Left            =   -73410
               TabIndex        =   26
               Top             =   1710
               Width           =   3375
            End
            Begin VB.TextBox txtTelephone3 
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
               Left            =   -73410
               TabIndex        =   25
               Top             =   2070
               Width           =   3375
            End
            Begin VB.TextBox txtHelbNumber 
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
               Left            =   -73110
               TabIndex        =   23
               Top             =   2040
               Width           =   3135
            End
            Begin VB.TextBox txtDaysPerMonth 
               Alignment       =   1  'Right Justify
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
               Left            =   -73110
               TabIndex        =   8
               Top             =   3420
               Width           =   465
            End
            Begin VB.TextBox txtNHIFNo 
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
               Left            =   -73110
               TabIndex        =   5
               Top             =   1680
               Width           =   3135
            End
            Begin VB.TextBox txtNSSFNo 
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
               Left            =   -73110
               TabIndex        =   4
               Top             =   1200
               Width           =   3135
            End
            Begin VB.TextBox txtPINNo 
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
               Height          =   285
               Left            =   -73110
               TabIndex        =   3
               Top             =   720
               Width           =   3135
            End
            Begin VB.TextBox txtHoursPerDay 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
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
               Height          =   285
               Left            =   -73110
               MaxLength       =   4
               TabIndex        =   6
               Top             =   2520
               Width           =   465
            End
            Begin VB.TextBox txtHoursPerMonth 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
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
               Height          =   285
               Left            =   -73110
               MaxLength       =   4
               TabIndex        =   7
               Top             =   2970
               Width           =   465
            End
            Begin VB.CommandButton cmdAddAccount 
               Caption         =   "Add"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   1680
               TabIndex        =   9
               Top             =   3960
               Width           =   1305
            End
            Begin VB.CommandButton cmdRemoveAccount 
               Caption         =   "Remove"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   3120
               TabIndex        =   10
               Top             =   3960
               Width           =   1305
            End
            Begin VB.Label Label8 
               Caption         =   "SMTP Server:"
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
               Left            =   -74850
               TabIndex        =   43
               Top             =   3225
               Width           =   1335
            End
            Begin VB.Label Label7 
               Caption         =   "Web URL (http://):"
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
               Left            =   -74850
               TabIndex        =   42
               Top             =   3615
               Width           =   1455
            End
            Begin VB.Label Label6 
               Caption         =   "E-Mail:"
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
               Left            =   -74850
               TabIndex        =   41
               Top             =   2835
               Width           =   1335
            End
            Begin VB.Label Label5 
               Caption         =   "Fax:"
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
               Left            =   -74850
               TabIndex        =   40
               Top             =   2445
               Width           =   1335
            End
            Begin VB.Label Label4 
               Caption         =   "Telephone #1:"
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
               Left            =   -74850
               TabIndex        =   39
               Top             =   1335
               Width           =   1335
            End
            Begin VB.Label Label3 
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
               Height          =   255
               Left            =   -74850
               TabIndex        =   38
               Top             =   555
               Width           =   1335
            End
            Begin VB.Label Label2 
               Caption         =   "Postal Address:"
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
               Left            =   -74850
               TabIndex        =   37
               Top             =   945
               Width           =   1335
            End
            Begin VB.Label Label12 
               Caption         =   "Telephone #2:"
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
               Left            =   -74880
               TabIndex        =   36
               Top             =   1755
               Width           =   1095
            End
            Begin VB.Label Label13 
               Caption         =   "Telephone #3:"
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
               Left            =   -74880
               TabIndex        =   35
               Top             =   2115
               Width           =   1275
            End
            Begin VB.Label Label22 
               Caption         =   "H.E.L.B #"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   -74820
               TabIndex        =   24
               Top             =   2055
               Width           =   1215
            End
            Begin VB.Label Label21 
               Caption         =   "Working Days /Month"
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
               TabIndex        =   22
               Top             =   3420
               Width           =   1635
            End
            Begin VB.Label Label11 
               Caption         =   "N.H.I.F #"
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
               TabIndex        =   19
               Top             =   1695
               Width           =   1575
            End
            Begin VB.Label Label10 
               Caption         =   "N.S.I.F Number"
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
               TabIndex        =   18
               Top             =   1215
               Width           =   1215
            End
            Begin VB.Label Label9 
               Caption         =   "P.I.N Number"
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
               TabIndex        =   17
               Top             =   735
               Width           =   1215
            End
            Begin VB.Label Label14 
               Caption         =   "Working Hours / Day"
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
               Left            =   -74820
               TabIndex        =   16
               Top             =   2520
               Width           =   1635
            End
            Begin VB.Label Label15 
               Caption         =   "Working Hours/ Month"
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
               TabIndex        =   15
               Top             =   2970
               Width           =   1725
            End
         End
         Begin VB.Label Label1 
            Caption         =   "Company Name"
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
            Left            =   90
            TabIndex        =   20
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Label Label20 
         Caption         =   "COMPANY DETAILS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   180
         Width           =   5595
      End
   End
End
Attribute VB_Name = "frmCompanyDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private CompanyDet As HRCORE.CompanyDetails
Private myBanks As New Banks
Private myBank As Bank
Private SelCompanyBank As CompanyBank
Private myCompanybanks As CompanyBanks
Private NewCompanyBanks As CompanyBanks
Private accnob4 As String




Private Sub cboBanks_Click()
    If Me.cboBanks.Text <> "" Then Call loadbankbranches
End Sub

Private Sub cmdAddAccount_Click()
    Dim myListItem As ListItem
    
    'Add company bank
    
    If Not currUser Is Nothing Then
        If currUser.CheckRight("CompanyDetails") <> secModify Then
            MsgBox "You don't have right to edit or add record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
   
    Select Case UCase(Me.cmdAddAccount.Caption)
        Case "ADD"
            Me.Frame3.Enabled = True
            Me.cmdAddAccount.Caption = "Update"
            Me.cmdEditAccount.Caption = "Cancel"
            Me.fraCompanyBanks.Enabled = False
            Me.cboBanks.SetFocus
            accnob4 = ""
        Case "UPDATE"
            If Me.txtAccountNO.Text <> "" Then
                Dim myCBank As New CompanyBank
                myCBank.AccountNo = Me.txtAccountNO
                myCBank.AccountType = Me.cboAccountTypes.Text
                myCBank.bankbranch.BankBranchID = Me.cboBankBranches.ItemData(ListIndex)
                myCBank.bankbranch.BranchName = Me.cboBankBranches.Text
                myCBank.bankbranch.Bank.bankid = cboBanks.ItemData(ListIndex)
                myCBank.bankbranch.Bank.BankName = cboBanks.Text
                
                 'VALIDATE TO ENSURE THAT THERE IS NO OTHER DEFAULT ACCOUNT IF THE NEW ENTRY IS A DEFAULT ACC
                 If chkDefault.value = 1 Then
                 
                 If validateBank("new") = False Then
                 Exit Sub
                 End If
                 
                 Dim n As Integer
                 Set myCompanybanks = New CompanyBanks
                 myCompanybanks.getAllCompanyBanks
                 For n = 1 To myCompanybanks.count
                 If myCompanybanks.Item(n).MainAccount = True Then
                 response = MsgBox("Sorry! Making the newly added bank as Default is not allowed. The Company already has its main account with " & myCompanybanks.Item(n).bankbranch.Bank.BankName & " BRANCH: " & myCompanybanks.Item(n).bankbranch.BranchName, vbOKOnly + vbCritical, "MAIN BANK")
                 Exit Sub
                 End If
                 Next n
                 
                 End If
                
                myCompanybanks.add myCBank
                NewCompanyBanks.add myCBank
                
                Call InsertCompanyBanks(myCBank)
                
                Set myListItem = Me.lvwCompanyAccounts.ListItems.add(, , Me.cboBanks.Text)
                myListItem.SubItems(1) = Me.cboBankBranches.Text
                myListItem.SubItems(2) = Me.txtAccountNO.Text
                myListItem.SubItems(3) = Me.cboAccountTypes.Text
                myListItem.SubItems(4) = IIf(Me.chkDefault.value = 1, "True", "False")
                        
                'refresh the listview
                
        '        Call PopulateCompanyBanks
                Me.txtAccountNO.Text = ""
                Me.cboAccountTypes.ListIndex = -1
                Me.cboBankBranches.ListIndex = -1
                GoTo Reset
            Else
                MsgBox "please Enter the account number"
            End If
        Case "CANCEL"
Reset:
            LoadTheCompanyBanks
            Me.cmdAddAccount.Caption = "Add"
            Me.cmdEditAccount.Caption = "Edit"
            Me.Frame3.Enabled = False
            Me.fraCompanyBanks.Enabled = True
        End Select
End Sub
Private Function validateBank(accnob4 As String) As Boolean
                 If chkDefault.value = 1 Then
                 Dim n As Integer
                 Set myCompanybanks = New CompanyBanks
                 myCompanybanks.getAllCompanyBanks
                 validateBank = True
                 For n = 1 To myCompanybanks.count
                 
                 If accnob4 <> "new" Then
                   If myCompanybanks.Item(n).MainAccount = True And myCompanybanks.Item(n).AccountNo <> accnob4 Then
                   response = MsgBox("Sorry! Making the newly added bank as Default is not allowed. The Company already has its main account with " & myCompanybanks.Item(n).bankbranch.Bank.BankName & " BRANCH: " & myCompanybanks.Item(n).bankbranch.BranchName, vbOKOnly + vbCritical, "MAIN BANK")
                   validateBank = False
                   Exit Function
                   End If
                 Else
                   If myCompanybanks.Item(n).MainAccount = True Then
                   response = MsgBox("Sorry! Making this Bank account as Default is not allowed. The Company already has its main account with " & myCompanybanks.Item(n).bankbranch.Bank.BankName & " BRANCH: " & myCompanybanks.Item(n).bankbranch.BranchName, vbOKOnly + vbCritical, "MAIN BANK")
                   validateBank = False
                   Exit Function
                   End If
                 End If
                 
                 Next n
                End If
                validateBank = True
End Function

Private Sub cmdCancel_Click()
    LoadCompanyDetails False
    cmdEdit.Caption = "Edit"
    cmdCancel.Enabled = False
    DisableControls
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdEditAccount_Click()
    On Error GoTo ErrorHandler
    
    Select Case UCase(Me.cmdEditAccount.Caption)
        Case "EDIT"
          If lvwCompanyAccounts.ListItems.count = 0 Then
          MsgBox ("select an item to be edited")
          Exit Sub
          End If
      
            Me.Frame3.Enabled = True
            Me.cmdAddAccount.Caption = "Cancel"
            Me.cmdEditAccount.Caption = "Update"
            Me.fraCompanyBanks.Enabled = False
            Me.cboBanks.SetFocus
            
'            Me.cboAccountTypes.Enabled = True
'            Me.cboBankBranches.Enabled = True
'            Me.cboBankBranches.Enabled = True
'            Me.txtAccountNO.Enabled = True
'            Me.cboBanks.Enabled = True

        Case "UPDATE"
            If Me.txtAccountNO.Text <> "" Then
            
            ''validate default banks
            If validateBank(accnob4) = False Then
            Exit Sub
            End If
                UpdateCompanyBanks
                GoTo Reset
            Else
                MsgBox "please Enter the account number"
            End If
        Case "CANCEL"
Reset:
            LoadTheCompanyBanks
            Me.cmdAddAccount.Caption = "Add"
            Me.cmdEditAccount.Caption = "Edit"
            Me.Frame3.Enabled = False
            Me.fraCompanyBanks.Enabled = True
        End Select
        
        Exit Sub
        
ErrorHandler:
        MsgBox "An Error has occured:" & vbNewLine & err.Description, vbExclamation, "Personnel Director"
End Sub

Private Sub cmdLoadLogo_Click()
    With cdlgLogo
        .DialogTitle = "Select Company Logo"
        .Filter = "Bitmap (*.bmp)|*.bmp|GIF (*.gif)|*.gif|Picture Files (*.jpg;*.gif;*.bmp)|*.jpg;*.gif;*.bmp"
        .InitDir = App.Path
        .ShowOpen
        Me.imgLogo.Stretch = True
        Me.imgLogo.Picture = LoadPicture(.FileName)
        frmMain2.Image1.Picture = LoadPicture(.FileName)
    End With
End Sub

Private Sub cmdRemoveAccount_Click()
     If Not currUser Is Nothing Then
        If currUser.CheckRight("CompanyDetails") <> secModify Then
            MsgBox "You don't have right to delete the record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    If Not (Me.lvwCompanyAccounts.SelectedItem Is Nothing) Then
        Call Delete
    End If
End Sub

Private Sub cmdRemoveLogo_Click()
   Me.imgLogo.Picture = LoadPicture()
End Sub

Private Sub cmdEdit_Click()
    If Not currUser Is Nothing Then
        If currUser.CheckRight("CompanyDetails") <> secModify Then
            MsgBox "You don't have right to edit or add record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    Select Case UCase(cmdEdit.Caption)
        Case "EDIT"
            'enable controls
            EnableControls
            cmdEdit.Caption = "Update"
            cmdCancel.Enabled = True
            sstCompanyInfo.Tab = 0
            txtCompanyName.SetFocus
        Case "UPDATE"
            Call UpdateCompanyDetails
            cmdEdit.Caption = "Edit"
            cmdCancel.Enabled = False
            DisableControls
    End Select
End Sub

Private Sub EnableControls()
    Dim ctl As Control
    
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Then
            ctl.Locked = False
        End If
    Next ctl
    
End Sub

Private Sub DisableControls()
    Dim ctl As Control
    
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Then
            ctl.Locked = True
        End If
    Next ctl
    
End Sub

Private Sub Form_Load()

    Set CompanyDet = New HRCORE.CompanyDetails
    Set myCompanybanks = New CompanyBanks 'used to populate listview
    Set SelCompanyBank = New CompanyBank
    Set NewCompanyBanks = New CompanyBanks 'use to add new Record
    frmMain2.PositionTheFormWithoutEmpList Me
    Frame2.Visible = True
    
    'disable the controls
    DisableControls
    
    'load the company details
    LoadCompanyDetails True
    'load banks
    Call DisplayBanks
    'get the company details
    LoadTheCompanyBanks
'    GetCompanyBanks
'    Call PopulateCompanyBanks
    sstCompanyInfo.Tab = 0
    Me.cmdRemoveAccount.Enabled = True
    Call CheckOverLap
End Sub

Private Sub LoadCompanyDetails(Refresh As Boolean)
    If Refresh Then
        CompanyDet.LoadCompanyDetails
    End If
    With CompanyDet
        Me.txtAccountNO.Text = ""
        Me.txtCompanyName.Text = .CompanyName
        Me.txtEmail.Text = .EMailAddress
        Me.txtFax.Text = .Fax
        Me.txtHoursPerDay.Text = .HoursPerDay
        Me.txtHoursPerMonth.Text = .HoursPerMonth
        Me.txtDaysPerMonth.Text = .WorkingDaysPerMonth
        Me.txtNHIFNo.Text = .NHIFNumber
        Me.txtNSSFNo.Text = .NSSFNumber
        Me.txtPhysicalAddress.Text = .PhysicalAddress
        Me.txtPINNo.Text = .PINNumber
        Me.txtPostalAddress.Text = .PostalAddress
        Me.txtSMTPServer.Text = .SMTPServer
        Me.txtTelephone1 = .Telephone1
        Me.txtTelephone2.Text = .Telephone2
        Me.txtTelephone3.Text = .Telephone3
        Me.txtWebURL.Text = .WebURL
        Me.imgLogo.Stretch = True
        Me.txtHelbNumber = .HELBNumber
        If .IsNGO = True Then
            Me.chkIsNGO.value = vbChecked
        Else
            Me.chkIsNGO.value = vbUnchecked
        End If
        
        Set Me.imgLogo.Picture = .logo
        
            Dim rs As New ADODB.Recordset
            Set rs = CConnect.GetRecordSet("SELECT HoursPerMonth,HoursPerDay FROM Companydetails where companyname='" & Me.txtCompanyName.Text & "'")
            If Not rs Is Nothing Then
            If Not rs.EOF Then
          
            Me.txtHoursPerMonth.Text = IIf(IsNull(rs!HoursPerMonth), vbNullString, rs!HoursPerMonth)
          
            Me.txtHoursPerDay.Text = IIf(IsNull(rs!HoursPerDay), vbNullString, rs!HoursPerDay)
            
            End If
            End If
    End With
    
End Sub


Private Sub UpdateCompanyDetails()
    Dim newCompany As New HRCORE.CompanyDetails
    
    On Error GoTo ErrorHandler
    Me.MousePointer = vbHourglass
    
    Dim hrsperday As Double
    Dim hrspermonth As Double
    hrsperday = 0
    hrspermonth = 0
    With newCompany
        .CompanyName = Trim(txtCompanyName.Text)
        .ApplicationEMailAddress = "hr@"
        .EMailAddress = Trim(txtEmail.Text)
        .Fax = Trim(txtFax.Text)
        .Telephone1 = Trim(txtTelephone1.Text)
        .Telephone2 = Trim(txtTelephone2.Text)
        .Telephone3 = Trim(txtTelephone3.Text)
        
        If IsNumeric(Trim(txtHoursPerDay.Text)) Then
        hrsperday = Trim(txtHoursPerDay.Text)
            .HoursPerDay = Trim(txtHoursPerDay.Text)
        Else
            If Len(Trim(txtHoursPerDay.Text)) > 0 Then
                MsgBox "Enter Numeric data for Hours per Day", vbInformation, TITLES
                Me.txtHoursPerDay.SetFocus
                Exit Sub
            Else
                .HoursPerDay = 0
            End If
        End If
        
        If IsNumeric(Trim(txtHoursPerMonth.Text)) Then
        hrspermonth = Trim(txtHoursPerMonth.Text)
            .HoursPerMonth = Trim(txtHoursPerMonth.Text)
        Else
            If Len(Trim(txtHoursPerMonth.Text)) > 0 Then
                MsgBox "Enter Numeric data for Hours per Day", vbInformation, TITLES
                Me.txtHoursPerMonth.SetFocus
                Exit Sub
            Else
                .HoursPerMonth = 0
            End If
        End If
        
        If IsNumeric(Trim(Me.txtDaysPerMonth)) Then
            .WorkingDaysPerMonth = Trim(txtDaysPerMonth.Text)
        Else
            If Len(Trim(txtDaysPerMonth.Text)) > 0 Then
                MsgBox "Enter Numeric data for Days per Month", vbInformation, TITLES
                Me.txtDaysPerMonth.SetFocus
                Exit Sub
            Else
                .WorkingDaysPerMonth = 0
            End If
        End If

        Set .logo = Me.imgLogo.Picture
        .NHIFNumber = Trim(Me.txtNHIFNo.Text)
        .NSSFNumber = Trim(Me.txtNSSFNo.Text)
        .PhysicalAddress = Trim(Me.txtPhysicalAddress.Text)
        .PINNumber = Trim(Me.txtPINNo.Text)
        .PostalAddress = Trim(Me.txtPostalAddress.Text)
        .SMTPServer = Trim(Me.txtSMTPServer.Text)
        .WebURL = Trim(Me.txtWebURL.Text)
        
        If Len(Me.txtHelbNumber) <= 0 Then
            .HELBNumber = 0
        Else
            .HELBNumber = CLng(Me.txtHelbNumber)
        End If
        
        If Me.chkIsNGO.value = vbChecked Then
            .IsNGO = True
        Else
            .IsNGO = False
        End If
        
        retVal = .UpdateCompanyDetails()
        Dim sql As String
        sql = "update companydetails set HoursPerMonth=" & hrspermonth & ",HoursPerDay=" & hrsperday & " where companyname='" & .CompanyName & "'"
        CConnect.ExecuteSql (sql)
        If retVal = 0 Then
            ''insert the companybanks
             
            MsgBox "The Company Details have been updated successfully", vbInformation, TITLES
        End If
        
    End With
    
    Me.MousePointer = vbDefault
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred" & vbNewLine & err.Description, vbInformation, TITLES
    Me.MousePointer = vbDefault
    
End Sub

Private Sub Form_Resize()
    oSmart.FResize Me
End Sub

Private Sub DisplayBanks()
    On Error GoTo ErrHandler
    Dim rs As ADODB.Recordset
    Dim myBank As Bank
    myBanks.Clear
    Set rs = CConnect.GetRecordSet("SELECT * FROM tblBank ORDER BY bank_id")
    
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst
        Do Until rs.EOF
        Set myBank = New Bank
            myBank.BankCode = rs!Bank_Code
            myBank.bankid = rs!bank_id
            myBank.BankName = rs!Bank_Name
           cboBanks.AddItem rs!Bank_Name
           cboBanks.ItemData(cboBanks.NewIndex) = rs!bank_id
            rs.MoveNext
           myBanks.add myBank
        Loop
    End If
    Set rs = Nothing
    Exit Sub
    
ErrHandler:
    MsgBox "An errr has occured when deleting the company bank"
End Sub

Private Sub loadbankbranches()
    Dim rs As ADODB.Recordset
    Me.cboBankBranches.Clear 'clear the listview
    
    mySQL = "SELECT * FROM tblBankBranch where Bank_Id=" & cboBanks.ItemData(cboBanks.ListIndex)
    Set rs = CConnect.GetRecordSet(mySQL)
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst
        Do Until rs.EOF
            cboBankBranches.AddItem rs!BANKBRANCH_NAME
            cboBankBranches.ItemData(Me.cboBankBranches.NewIndex) = rs!Bankbranch_id
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
End Sub

Private Sub LoadTheCompanyBanks()
    'THIS SUB PROC LOADS ALL THE COMPANY BANKS AND POPULATES THEM IN THE LIST VIEW CONTROL
    Dim rs As ADODB.Recordset
    Dim myListItem As ListItem
    Dim lngLoopVariable As Long
    On Error GoTo ErrorHandler
   
    
    
   
    'CLEARING THE LIST VIEW CONROL
    Me.lvwCompanyAccounts.ListItems.Clear
    Set rs = CConnect.GetRecordSet("SELECT * FROM vwCompanyBanks")
    If Not rs Is Nothing Then
        If Not (rs.BOF And rs.EOF) Then
            rs.MoveFirst
            Do Until rs.EOF
             
                Set myListItem = Me.lvwCompanyAccounts.ListItems.add(, , IIf(IsNull(rs!BankName), vbNullString, rs!BankName))
                myListItem.SubItems(1) = IIf(IsNull(rs!BranchName), vbNullString, rs!BranchName)
                myListItem.SubItems(2) = IIf(IsNull(rs!AccountNo), vbNullString, rs!AccountNo)
                myListItem.SubItems(3) = IIf(IsNull(rs!AccountType), vbNullString, rs!AccountType)
                myListItem.SubItems(4) = IIf(IsNull(rs!MainAccount), "False", IIf(rs!MainAccount = True, "True", "False"))
                myListItem.Tag = rs!CompanyBank_ID
                
                rs.MoveNext
            Loop
        End If
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "An Error has occurred while attempting to load the company banks" & vbCrLf & err.Description, vbExclamation, TITLES
End Sub

Private Sub GetCompanyBanks()
    On Error GoTo ErrHandler
    Dim rs As ADODB.Recordset
    Dim MycompanyBank As CompanyBank
    mySQL = "SELECT  *  FROM  vwCompanybanks"
    Set rs = CConnect.GetRecordSet(mySQL)
    If rs Is Nothing Then Exit Sub
    myCompanybanks.Clear
    
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst
        Do Until rs.EOF
            Set MycompanyBank = New CompanyBank
            MycompanyBank.AccountNo = rs!AccountNo
            If Not IsNull(rs!AccountType) Then MycompanyBank.AccountType = rs!AccountType
            If Not IsNull(rs!MainAccount) Then MycompanyBank.MainAccount = rs!MainAccount
            MycompanyBank.bankbranch.BankBranchID = rs!BranchID
            MycompanyBank.bankbranch.BranchCode = rs!BranchCode
            MycompanyBank.bankbranch.BranchName = rs!BranchName
            MycompanyBank.bankbranch.Bank.BankCode = rs!BankCode
            MycompanyBank.bankbranch.Bank.BankName = rs!BankName
            If Not IsNull(rs!bankid) Then MycompanyBank.bankbranch.Bank.bankid = rs!bankid
            myCompanybanks.add MycompanyBank
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
     Exit Sub
    
ErrHandler:
    MsgBox "An error has occured when populating the company banks"
End Sub

Private Sub PopulateCompanyBanks()
    On Error GoTo ErrHandler
    Dim ItemX As ListItem
    Dim i As Long
    
    Call GetCompanyBanks
    
    lvwCompanyAccounts.ListItems.Clear
    For i = 1 To myCompanybanks.count
        Set ItemX = Me.lvwCompanyAccounts.ListItems.add(, , myCompanybanks.Item(i).bankbranch.Bank.BankName)
        ItemX.SubItems(1) = myCompanybanks.Item(i).bankbranch.BranchName
        ItemX.SubItems(2) = myCompanybanks.Item(i).AccountNo
        ItemX.SubItems(3) = myCompanybanks.Item(i).AccountType
        ItemX.Tag = myCompanybanks.Item(i).AccountNo
    Next i
    Exit Sub
    
ErrHandler:
    MsgBox "An error has occured when populating the company bank"
End Sub

Private Sub lvwCompanyAccounts_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim lngLoopVariable As Long
    On Error GoTo ErrorHandler
    
    For lngLoopVariable = 0 To Me.cboBanks.ListCount - 1
        If UCase(Me.cboBanks.List(lngLoopVariable)) = UCase(Item.Text) Then
            Me.cboBanks.ListIndex = lngLoopVariable
            Exit For
        End If
    Next
    For lngLoopVariable = 0 To Me.cboBankBranches.ListCount - 1
        If UCase(Me.cboBankBranches.List(lngLoopVariable)) = UCase(Item.SubItems(1)) Then
            Me.cboBankBranches.ListIndex = lngLoopVariable
            Exit For
        End If
    Next
    Me.txtAccountNO.Text = Item.SubItems(2)
    Me.txtAccountNO.Tag = Item.Tag
    Me.cboAccountTypes.Text = Item.SubItems(3)
    Me.chkDefault.value = IIf(UCase(Item.SubItems(4)) = "TRUE", 1, 0)
    accnob4 = Me.txtAccountNO.Text
        
'    Set SelCompanyBank = myCompanybanks.FindCompanyBank(Item.SubItems(2))
'    If Not (SelCompanyBank Is Nothing) Then
'        cmdRemoveAccount.Enabled = True
'        For lngLoopVariable = 0 To Me.cboBanks.ListCount - 1
'            If Me.cboBanks.ItemData(lngLoopVariable) = SelCompanyBank.bankbranch.Bank.BankID Then
'                Me.cboBanks.ListIndex = lngLoopVariable
'                Exit For
'            End If
'        Next
'        For lngLoopVariable = 0 To Me.cboBankBranches.ListCount - 1
'            If Me.cboBankBranches.ItemData(lngLoopVariable) = SelCompanyBank.bankbranch.BranchID Then
'                Me.cboBankBranches.ListIndex = lngLoopVariable
'                Exit For
'            End If
'        Next
'        Me.txtAccountNO.Text = SelCompanyBank.AccountNo
'        Me.cboAccountTypes.Text = SelCompanyBank.AccountType
'        Me.chkDefault.value = IIf(UCase(Item.SubItems(4)) = "TRUE", 1, 0)
'    End If
    Exit Sub
ErrorHandler:
    MsgBox "An Error has occurred while attemting to display the selected company bank" & vbCrLf & err.Description, vbExclamation, TITLES
End Sub

Private Sub sstCompanyInfo_Click(PreviousTab As Integer)
    If sstCompanyInfo.Tab = 2 Then
        fraLogo.Visible = False
    Else
        fraLogo.Visible = True
    End If
End Sub

Private Sub txtAccountNo_Change()

'    Dim repeated, i As Integer
'    repeated = 0
'    For i = 1 To lvwCompanyAccounts.ListItems.count
'        If lvwCompanyAccounts.ListItems.Item(i).SubItems(1) = cboBankBranches.Text And lvwCompanyAccounts.ListItems.Item(i).SubItems(2) = Me.txtAccountNO Then
'             repeated = repeated + 1
'        End If
'    Next
'    If repeated >= 1 Then
'
'        MsgBox "A Bank with That account number Already exist"
'        txtAccountNO.SetFocus
'    End If
End Sub

Private Sub InsertCompanyBanks(NewBank As CompanyBank)
    Dim defaultBank As Integer
    
    defaultBank = 0
    If chkDefault.value = vbChecked Then defaultBank = 1

    mySQL = "Insert into tblCompanyBank(BankID,BankBranch_ID,Account_Number,AccountType,Default_Bank) VALUES(" & NewBank.bankbranch.Bank.bankid & "," & NewBank.bankbranch.BankBranchID & ",'" & NewBank.AccountNo & "' , '" & NewBank.AccountType & "'," & defaultBank & ")"
    CConnect.ExecuteSql (mySQL)

End Sub

Private Sub UpdateCompanyBanks()
    mySQL = "Update tblCompanyBank SET BankID=" & Me.cboBanks.ItemData(Me.cboBanks.ListIndex) & ",BankBranch_ID=" & Me.cboBankBranches.ItemData(Me.cboBankBranches.ListIndex) & ",Account_Number='" & Me.txtAccountNO.Text & "',AccountType='" & Me.cboAccountTypes.Text & "',Default_Bank=" & Me.chkDefault.value & " WHERE CompanyBank_ID=" & Me.txtAccountNO.Tag
    CConnect.ExecuteSql (mySQL)
End Sub
Private Sub Delete()
    On Error GoTo ErrHandler
    mySQL = "Delete From tblCompanyBank where CompanyBank_ID=" & Me.lvwCompanyAccounts.SelectedItem.Tag
    CConnect.ExecuteSql (mySQL)
    LoadTheCompanyBanks
'    Call GetCompanyBanks
'    Call PopulateCompanyBanks
    Exit Sub
    
ErrHandler:
    MsgBox "An errr has occured when deleting the company bank"
End Sub

Private Sub CheckOverLap()
    Me.fraLogo.Left = Me.txtPhysicalAddress.Left + Me.txtPhysicalAddress.Width + 500
    Me.fraLogo.Top = 480
'    Me.cmdLoadLogo.Left = 200
'    Me.cmdLoadLogo.Top = Me.fraLogo.Width - 50
'    Me.cmdRemoveLogo.Left = 1630
'    Me.cmdRemoveLogo.Top = Me.cmdLoadLogo.Top
    Me.imgLogo.Left = 200
End Sub

Private Sub txtHelbNumber_Change()
    If Not IsNumeric(txtHelbNumber.Text) Then
        MsgBox "Please enter numeric value"
        txtHelbNumber.Text = ""
    End If
End Sub
