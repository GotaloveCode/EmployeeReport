VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmKin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Next of Kin"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   Icon            =   "frmKin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imgEmpTool 
      Left            =   5430
      Top             =   195
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
            Picture         =   "frmKin.frx":0442
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKin.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKin.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKin.frx":0778
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   7900
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   7800
      Begin VB.Frame fraDetails 
         Appearance      =   0  'Flat
         Caption         =   "Next of Kin"
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
         Height          =   5685
         Left            =   165
         TabIndex        =   28
         Top             =   743
         Visible         =   0   'False
         Width           =   7035
         Begin VB.CheckBox chkSetAsEmergency 
            Appearance      =   0  'Flat
            Caption         =   "&Set this record as my Emergency contact"
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
            Height          =   345
            Left            =   120
            TabIndex        =   47
            Top             =   4680
            Width           =   3360
         End
         Begin VB.TextBox txtAlien 
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
            Height          =   285
            Left            =   1080
            TabIndex        =   5
            Top             =   1155
            Width           =   945
         End
         Begin VB.TextBox txtPassport 
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
            Height          =   285
            Left            =   120
            TabIndex        =   4
            Top             =   1155
            Width           =   825
         End
         Begin VB.TextBox txtONames 
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
            Height          =   285
            Left            =   2550
            TabIndex        =   2
            Top             =   615
            Width           =   3120
         End
         Begin VB.TextBox txtIDNo 
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
            Height          =   285
            Left            =   5775
            TabIndex        =   3
            Top             =   615
            Width           =   1065
         End
         Begin MSMask.MaskEdBox txtBYear 
            Height          =   285
            Left            =   3780
            TabIndex        =   7
            Top             =   1155
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtMNo 
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
            Height          =   285
            Left            =   3915
            TabIndex        =   11
            Top             =   1740
            Width           =   1800
         End
         Begin VB.TextBox txtEMail 
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
            Height          =   300
            Left            =   120
            TabIndex        =   12
            Top             =   2310
            Width           =   6750
         End
         Begin VB.TextBox txtOccupation 
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
            Height          =   285
            Left            =   5355
            TabIndex        =   8
            Top             =   1155
            Width           =   1515
         End
         Begin VB.CheckBox chkSigned 
            Appearance      =   0  'Flat
            Caption         =   "Signed"
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
            Height          =   345
            Left            =   5430
            TabIndex        =   14
            Top             =   2820
            Width           =   1320
         End
         Begin VB.TextBox txtAddress 
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
            Height          =   840
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   2880
            Width           =   5175
         End
         Begin VB.TextBox txtOTelNo 
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
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Top             =   1740
            Width           =   1755
         End
         Begin VB.TextBox txtHTelNo 
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
            Height          =   285
            Left            =   1995
            TabIndex        =   10
            Top             =   1740
            Width           =   1800
         End
         Begin VB.ComboBox cboRel 
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
            ItemData        =   "frmKin.frx":0CBA
            Left            =   2160
            List            =   "frmKin.frx":0CBC
            Style           =   1  'Simple Combo
            TabIndex        =   6
            Top             =   1140
            Width           =   1485
         End
         Begin VB.CommandButton cmdCancel 
            Cancel          =   -1  'True
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
            Left            =   6375
            Picture         =   "frmKin.frx":0CBE
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Cancel Process"
            Top             =   5040
            Width           =   495
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            Default         =   -1  'True
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
            Left            =   5895
            Picture         =   "frmKin.frx":0DC0
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Save Record"
            Top             =   5040
            Width           =   495
         End
         Begin VB.TextBox txtComments 
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
            Height          =   705
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   3960
            Width           =   6735
         End
         Begin VB.TextBox txtCode 
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
            Height          =   285
            Left            =   120
            TabIndex        =   0
            Top             =   615
            Width           =   825
         End
         Begin VB.TextBox txtSurName 
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
            Height          =   285
            Left            =   1065
            TabIndex        =   1
            Top             =   615
            Width           =   1380
         End
         Begin MSComCtl2.DTPicker dtpSDate 
            Height          =   315
            Left            =   5430
            TabIndex        =   15
            Top             =   3420
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
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
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   63766531
            CurrentDate     =   37972
         End
         Begin MSComCtl2.DTPicker dtpDOB 
            Height          =   285
            Left            =   3960
            TabIndex        =   44
            Top             =   1155
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   63766529
            CurrentDate     =   38763
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Alien No."
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
            Left            =   1080
            TabIndex        =   46
            Top             =   915
            Width           =   645
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Passport No"
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
            TabIndex        =   45
            Top             =   915
            Width           =   870
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Mobile No"
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
            Left            =   3900
            TabIndex        =   43
            Top             =   1515
            Width           =   690
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "E-Mail"
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
            TabIndex        =   42
            Top             =   2055
            Width           =   420
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Occupation"
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
            Left            =   5355
            TabIndex        =   41
            Top             =   915
            Width           =   810
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "ID No"
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
            Left            =   5775
            TabIndex        =   40
            Top             =   390
            Width           =   405
         End
         Begin VB.Label lblSDate 
            AutoSize        =   -1  'True
            Caption         =   "Signed Date"
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
            Left            =   5415
            TabIndex        =   39
            Top             =   3195
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Address"
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
            TabIndex        =   38
            Top             =   2655
            Width           =   585
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Office Tel No"
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
            TabIndex        =   37
            Top             =   1515
            Width           =   930
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Home Tel No"
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
            Left            =   1980
            TabIndex        =   36
            Top             =   1515
            Width           =   900
         End
         Begin VB.Label Label3 
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
            Height          =   195
            Left            =   3780
            TabIndex        =   35
            Top             =   915
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Relationship"
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
            Left            =   2190
            TabIndex        =   34
            Top             =   915
            Width           =   870
         End
         Begin VB.Label Label2 
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
            Height          =   195
            Left            =   2550
            TabIndex        =   33
            Top             =   390
            Width           =   945
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Comments"
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
            TabIndex        =   32
            Top             =   3735
            Width           =   750
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Code"
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
            TabIndex        =   30
            Top             =   390
            Width           =   375
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "SurName"
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
            Left            =   1080
            TabIndex        =   29
            Top             =   390
            Width           =   645
         End
      End
      Begin MSComDlg.CommonDialog Cdl 
         Left            =   5610
         Top             =   4560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Move to the First employee"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
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
         Left            =   6270
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   5400
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   20
         ToolTipText     =   "Move to the Previous employee"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   21
         ToolTipText     =   "Move to the Next employee"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
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
         Left            =   2775
         Picture         =   "frmKin.frx":0EC2
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Add New record"
         Top             =   5400
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
         Left            =   3255
         Picture         =   "frmKin.frx":0FC4
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Edit Record"
         Top             =   5400
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
         Left            =   3735
         Picture         =   "frmKin.frx":10C6
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Delete Record"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   22
         ToolTipText     =   "Move to the Last employee"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSComctlLib.ListView lvwDetails 
         Height          =   7800
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   7800
         _ExtentX        =   13758
         _ExtentY        =   13758
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
   End
End
Attribute VB_Name = "frmKin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cboRel_KeyPress(KeyAscii As Integer)
'    KeyAscii = 0
End Sub

Private Sub chkSigned_Click()
    If chkSigned.value = 1 Then
        lblSDate.Visible = True
        dtpSDate.Visible = True
    Else
        lblSDate.Visible = False
        dtpSDate.Visible = False
        
    End If
    
End Sub

Public Sub cmdCancel_Click()
    If PromptSave = True Then
        If MsgBox("Close this window?", vbYesNo + vbQuestion, "Confirm Close") = vbNo Then Exit Sub
    End If
    fraDetails.Visible = False
    With frmMain2
        .cmdNew.Enabled = True
        .cmdEdit.Enabled = True
        .cmdDelete.Enabled = True
        .cmdCancel.Enabled = False
        .cmdSave.Enabled = False
    End With
    Call EnableCmd
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Public Sub cmdDelete_Click()
Dim resp As String
On Error GoTo ErrHandler
If SelectedEmployee Is Nothing Then Exit Sub

If lvwDetails.ListItems.count > 0 Then
    resp = MsgBox("Are you sure you want to delete  " & lvwDetails.SelectedItem.SubItems(1) & " from the records?", vbQuestion + vbYesNo)
    If resp = vbNo Then
        Exit Sub
    End If
    Action = "DELETED EMPLOYEE NEXT OF KIN; EMPLOYEE CODE: " & SelectedEmployee.EmpCode & "; KIN CODE: " & lvwDetails.SelectedItem
    CConnect.ExecuteSql ("DELETE FROM Kin WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & lvwDetails.SelectedItem & "'")
    
    rs2.Requery
    Call DisplayRecords
Else
    MsgBox "You have to select the record you would like to delete.", vbInformation
            
End If
    Exit Sub
ErrHandler:
End Sub

Private Sub cmdDone_Click()
    PSave = True
    Call cmdCancel_Click
    PSave = False
End Sub

Public Sub cmdEdit_Click()
    On Error GoTo ErrHandler
    
    If SelectedEmployee Is Nothing Then Exit Sub
    
    If lvwDetails.ListItems.count < 1 Then
        MsgBox "You have to select the kin you would like to edit.", vbInformation
        PSave = True
        Call cmdCancel_Click
        PSave = False
        Exit Sub
    End If
    
    Set rs3 = CConnect.GetRecordSet("SELECT * FROM Kin WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & lvwDetails.SelectedItem & "'")
    With rs3
        If .RecordCount > 0 Then
            txtCode.Text = !Code & ""
            txtCode.Tag = txtCode.Text
            txtSurname.Text = !SurName & ""
            txtONames.Text = !OtherNames & ""
            txtIDNo.Text = !IdNo & ""
            cboRel.Text = !Relation & ""
            If Not IsNull(!DOB) And Trim(!DOB & "") <> "" Then txtBYear.Text = !DOB & ""
            txtHTelNo.Text = !HTelNo & ""
            txtOTelNo.Text = !OffTelNo & ""
            txtMNo.Text = !MNo & ""
            txtEmail.Text = !EMail & ""
            txtAddress.Text = !Address & ""
            txtComments.Text = !Comments & ""
            txtOccupation.Text = !Occupation & ""
            txtPassport.Text = !PassportNo & ""
            txtAlien.Text = !AlienNo & ""
            chkSetAsEmergency.value = IIf(Trim(!EMERGENCYCONTACT & "") = True, 1, 0)
            If !Signed = "Yes" Then
                chkSigned.value = 1
                If Not IsNull(!SDate) Then dtpSDate.value = !SDate & ""
            Else
                chkSigned.value = 0
            End If
            
            SaveNew = False
        Else
            MsgBox "Record not found.", vbInformation
            Set rs3 = Nothing
            PSave = True
            Call cmdCancel_Click
            PSave = False
            Exit Sub
        End If
    End With
        
    Set rs3 = Nothing
    
    Call DisableCmd
    
    fraDetails.Visible = True
    
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    SaveNew = False
    
    txtCode.Locked = False
    txtSurname.SetFocus
    Exit Sub
ErrHandler:
End Sub

Private Sub cmdFirst_Click()

With rsGlob
    If .RecordCount > 0 Then
        If .BOF <> True Then
            .MoveFirst
            If .BOF = True Then
                .MoveFirst
                Call DisplayRecords
            Else
                Call DisplayRecords
            End If
            
            Call FirstDisb
            
        End If
    End If
End With


End Sub

Private Sub cmdLast_Click()
With rsGlob
    If .RecordCount > 0 Then
        If .EOF <> True Then
            .MoveLast
            If .EOF = True Then
                .MoveLast
                Call DisplayRecords
            Else
                Call DisplayRecords
            End If
            
            Call LastDisb
            
        End If
    End If
End With

End Sub

Public Sub cmdNew_Click()
    Call DisableCmd
    Call Cleartxt
    
    fraDetails.Visible = True
    cmdCancel.Enabled = True
    SaveNew = True
    cmdSave.Enabled = True
    chkSigned.value = 0
    txtCode.Text = loadKinCode
    txtCode.Locked = False
    txtSurname.SetFocus
    dtpSDate.value = Date
End Sub

Private Sub cmdNext_Click()
    
With rsGlob
    If .RecordCount > 0 Then
        If .EOF <> True Then
            .MoveNext
            If .EOF = True Then
                .MoveLast
                Call DisplayRecords
            Else
                Call DisplayRecords
            End If

            Call LastDisb

        End If
    End If
End With


End Sub

Private Sub cmdPrevious_Click()

With rsGlob
    If .RecordCount > 0 Then
        If .BOF <> True Then
            .MovePrevious
            If .BOF = True Then
                .MoveFirst
                Call DisplayRecords
            Else
                Call DisplayRecords
            End If
            
            Call FirstDisb
            
        End If
    End If
End With


End Sub

Public Sub cmdSave_Click()
    
    
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("NextOfKin") <> secModify Then
            MsgBox "You Don't have right to add or edit record. Please liaise with security admin"
            Exit Sub
        End If
   End If
   
   If SelectedEmployee Is Nothing Then Exit Sub
    If txtCode.Text = "" Then
        MsgBox "Enter the code.", vbInformation
        txtCode.SetFocus
        Exit Sub
    End If
    
    If CheckForNumbers(txtSurname.Text) > 0 Then
        MsgBox "The kin's surname cannot contain numbers.", vbInformation
        txtSurname.SetFocus
        Exit Sub
    End If
    
    If CheckForNumbers(txtONames.Text) > 0 Then
        MsgBox "The kin's othernames cannot contain numbers.", vbInformation
        txtONames.SetFocus
        Exit Sub
    End If
    
    If CheckForNumbers(txtIDNo.Text) = 0 And CheckForNumbers(txtAlien.Text) = 0 And CheckForNumbers(txtPassport.Text) = 0 Then
        MsgBox "Please enter the kins ID, passport or alien card number." & vbCrLf & "This must at least contain some numeric values.", vbInformation
        txtIDNo.SetFocus
        Exit Sub
    End If
    
    If cboRel.Text = "" Then
        MsgBox "You have to enter the next of kins relationship to the employee.", vbInformation
        cboRel.SetFocus
        Exit Sub
    End If
    
    If txtBYear.Text = "" Then
        MsgBox "You have to enter the year of birth.", vbInformation
        Exit Sub
    End If
    
    If CheckForNumbers(txtHTelNo.Text) = 0 And CheckForNumbers(txtMNo.Text) = 0 And CheckForNumbers(txtOTelNo.Text) = 0 Then
        MsgBox "Please provide at least one set of the kin's telephone number." & "This must contain at least a numeric value.", vbOKOnly + vbInformation, "Missing Telephone"
        txtOTelNo.SetFocus
        Exit Sub
    End If
    
    If txtEmail.Text = "" And txtAddress.Text = "" Then
        MsgBox "Please provide at least one: E-mail or Address.", vbOKOnly + vbInformation, "Missing address"
        txtEmail.SetFocus
        Exit Sub
    End If
    If SaveNew = True Then
        
        Set rs4 = CConnect.GetRecordSet("SELECT * FROM Kin WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & txtCode.Text & "'")
        With rs4
            If .RecordCount > 0 Then
                MsgBox "Kin's code already exists. Enter another one.", vbInformation
                txtCode.Text = ""
                txtCode.SetFocus
                Set rs4 = Nothing
                Exit Sub
            End If
        End With
        Set rs4 = Nothing
    End If
       
    If PromptSave = True Then
        If MsgBox("Are you sure you want to save the record.", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    If chkSetAsEmergency.value = 1 Then
        CConnect.ExecuteSql "UPDATE family SET EmergencyContact=0 WHERE employee_id='" & SelectedEmployee.EmployeeID & "'"
        CConnect.ExecuteSql "UPDATE kin SET EmergencyContact=0 WHERE employee_id='" & SelectedEmployee.EmployeeID & "'"
    End If
    
    CConnect.ExecuteSql ("DELETE FROM Kin WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & txtCode.Tag & "'")
    
    If chkSigned.value = 1 Then
        mySQL = "INSERT INTO Kin (employee_id, Code, SurName, OtherNames, IDNo, Relation, DOB, Occupation, HTelNo, OffTelNo, MNo, EMail, Address, Signed, SDate,Comments,passportNo,alienNo,EmergencyContact)" & _
                        " VALUES('" & SelectedEmployee.EmployeeID & "','" & txtCode.Text & "','" & txtSurname.Text & "','" & txtONames.Text & "','" & txtIDNo.Text & "','" & cboRel.Text & "','" & txtBYear.Text & "'," & _
                        "'" & txtOccupation.Text & "','" & txtHTelNo.Text & "','" & txtOTelNo.Text & "','" & txtMNo.Text & "','" & txtEmail.Text & "','" & txtAddress.Text & "','Yes','" & Format(dtpSDate.value, Dfmt) & "','" & txtComments.Text & "','" & Replace(txtPassport.Text, "'", "''") & "','" & Replace(txtAlien.Text, "'", "''") & "'," & chkSetAsEmergency.value & ")"
    Else
        mySQL = "INSERT INTO Kin (employee_id, Code, SurName, OtherNames, IDNo, Relation, DOB, Occupation, HTelNo, OffTelNo, MNo, EMail, Address, Signed, Comments,passportNo,alienNo,EmergencyContact)" & _
                        " VALUES('" & SelectedEmployee.EmployeeID & "','" & txtCode.Text & "','" & txtSurname.Text & "','" & txtONames.Text & "','" & txtIDNo.Text & "','" & cboRel.Text & "','" & txtBYear.Text & "'," & _
                        "'" & txtOccupation.Text & "','" & txtHTelNo.Text & "','" & txtOTelNo.Text & "','" & txtMNo.Text & "','" & txtEmail.Text & "','" & txtAddress.Text & "','No','" & txtComments.Text & "','" & Replace(txtPassport.Text, "'", "''") & "','" & Replace(txtAlien.Text, "'", "''") & "'," & chkSetAsEmergency.value & ")"
    End If
    Action = "ADDED EMPLOYEE NEXT OF KIN; EMPLOYEE CODE: " & SelectedEmployee.EmpCode & "; KIN CODE: " & txtCode.Text & "; IDENTITY CARD NUMBER: " & txtIDNo.Text
    CConnect.ExecuteSql (mySQL)
    txtCode.Tag = ""
    rs2.Requery
    
    If SaveNew = False Then
        PSave = True
        Call DisplayRecords
        Call cmdCancel_Click
        PSave = False
    Else
        rs2.Requery
        Call DisplayRecords
        txtCode.Text = loadKinCode
        txtSurname.SetFocus
        SaveNew = True
        
    End If
        
End Sub


Private Sub dtpDOB_CloseUp()
    txtBYear.Text = Format(dtpDOB, "dd/MM/yyyy")
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler
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
    
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    
    Call InitGrid
    
    Set rs2 = CConnect.GetRecordSet("SELECT * FROM Kin ORDER BY Code")
    
    With rsGlob
        If .RecordCount < 1 Then
            Call DisableCmd
            Exit Sub
        End If
    End With
    
    Call DisplayRecords
    
    cmdFirst.Enabled = False
    cmdPrevious.Enabled = False
    Exit Sub
ErrHandler:
    MsgBox "An error has occured in the " & Me.Name & "  " & err.Description
End Sub

Private Sub Form_Resize()
    oSmart.FResize Me
    
    Me.Height = tvwMainheight - 150
    Frame1.Move Frame1.Left, 0, Frame1.Width, tvwMainheight - 150
    lvwDetails.Move lvwDetails.Left, 0, lvwDetails.Width, tvwMainheight - 150
    
End Sub

Private Sub InitGrid()
    With lvwDetails
        .ColumnHeaders.add , , "Code", 0
        .ColumnHeaders.add , , "SurName", 1700
        .ColumnHeaders.add , , "Other Names", 2500
        .ColumnHeaders.add , , "ID No", 1700
        .ColumnHeaders.add , , "Passport No", 1700
        .ColumnHeaders.add , , "Alien No", 1700
        .ColumnHeaders.add , , "Relationship", 1700
        .ColumnHeaders.add , , "Year Of Birth"
        .ColumnHeaders.add , , "Occupation"
        .ColumnHeaders.add , , "Home Tel No"
        .ColumnHeaders.add , , "Office Tel No"
        .ColumnHeaders.add , , "Cell No"
        .ColumnHeaders.add , , "E-Mail", 1250
        .ColumnHeaders.add , , "Address", 3000
        .ColumnHeaders.add , , "Signed"
        .ColumnHeaders.add , , "Sign Date"
        .ColumnHeaders.add , , "Comments", 3500
        .ColumnHeaders.add , , "Is Emergency Contact", 1750
        .View = lvwReport
    End With
End Sub

Public Sub DisplayRecords()
    If SelectedEmployee Is Nothing Then Exit Sub
    
    lvwDetails.ListItems.Clear
    Call Cleartxt
    
    With rsGlob
        If Not .EOF And Not .BOF Then
                      
            With rs2
                If .RecordCount > 0 Then
                    
                  .Filter = "employee_id like '" & SelectedEmployee.EmployeeID & "'"
                                                                            
                    If .RecordCount > 0 Then
                        .MoveFirst
                        Do While Not .EOF
                            Set li = lvwDetails.ListItems.add(, , !Code & "", , 5)
                            li.ListSubItems.add , , !SurName & ""
                            li.ListSubItems.add , , !OtherNames & ""
                            li.ListSubItems.add , , !IdNo & ""
                            li.ListSubItems.add , , !PassportNo & ""
                            li.ListSubItems.add , , !AlienNo & ""
                            li.ListSubItems.add , , !Relation & ""
                            li.ListSubItems.add , , !DOB & ""
                            li.ListSubItems.add , , !Occupation & ""
                            li.ListSubItems.add , , !HTelNo & ""
                            li.ListSubItems.add , , !OffTelNo & ""
                            li.ListSubItems.add , , !MNo & ""
                            li.ListSubItems.add , , !EMail & ""
                            li.ListSubItems.add , , !Address & ""
                            li.ListSubItems.add , , !Signed & ""
                            li.ListSubItems.add , , !SDate & ""
                            li.ListSubItems.add , , !Comments & ""
                            li.ListSubItems.add , , IIf(Trim(!EMERGENCYCONTACT & "") = True, "Yes", "No")
                            li.Tag = !ID
                            
                            .MoveNext
                        Loop
                    End If
                    .Filter = adFilterNone
                End If
            End With
            
        End If
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set rs2 = Nothing
    Set rs5 = Nothing
'    Set Cnn = Nothing
    

    frmMain2.Caption = "Infiniti HRMIS - PDR [Current User:\" & currUser.FullNames & "]"
    
End Sub

Private Sub fraList_DragDrop(Source As Control, X As Single, y As Single)

End Sub

Private Sub lvwDetails_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwDetails
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvwDetails_DblClick()
    If frmMain2.cmdEdit.Enabled = True Then
        Call frmMain2.cmdEdit_Click
    End If
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
If Len(Trim(txtAddress.Text)) > 198 Then
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
  Case Asc("@")
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

Private Sub txtCode_Change()
    txtCode.Text = UCase(txtCode.Text)
    txtCode.SelStart = Len(txtCode.Text)
End Sub


Private Sub LastDisb()
With rsGlob
    If Not .EOF Then
        .MoveNext
        If .EOF Then
            cmdLast.Enabled = False
            cmdNext.Enabled = False
            cmdPrevious.Enabled = True
            cmdFirst.Enabled = True
            cmdPrevious.SetFocus
        End If
        .MovePrevious
    End If
    
    cmdPrevious.Enabled = True
    cmdFirst.Enabled = True
End With
End Sub


Private Sub FirstDisb()
With rsGlob
    If Not .BOF Then
        .MovePrevious
        If .BOF Then
            cmdLast.Enabled = True
            cmdNext.Enabled = True
            cmdPrevious.Enabled = False
            cmdFirst.Enabled = False
            cmdNext.SetFocus
        End If
        .MoveNext
    End If
    
    cmdLast.Enabled = True
    cmdNext.Enabled = True
End With
End Sub


Private Sub Cleartxt()
Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
            i.Text = ""
        End If
    Next i
    
    txtBYear.Text = "__/__/____"
    'lvwDetails.ListItems.clear
    
End Sub

Public Sub disabletxt()
Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
            i.Locked = True
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

Public Sub EnableCmd()
Dim i As Object
    For Each i In Me
        If TypeOf i Is CommandButton Then
            i.Enabled = True
        End If
    Next i
    
End Sub



Public Sub FirstLastDisb()
cmdLast.Enabled = True
cmdNext.Enabled = True
cmdPrevious.Enabled = True
cmdFirst.Enabled = True
cmdNext.SetFocus
            
With rsGlob
    If Not .BOF = True Then
        .MovePrevious
        If .BOF = True Then
            cmdLast.Enabled = True
            cmdNext.Enabled = True
            cmdPrevious.Enabled = False
            cmdFirst.Enabled = False
            cmdNext.SetFocus
        End If
        .MoveNext
    Else
        cmdLast.Enabled = True
        cmdNext.Enabled = True
        cmdPrevious.Enabled = False
        cmdFirst.Enabled = False
        cmdNext.SetFocus
    End If
    
    If Not .EOF = True Then
        .MoveNext
        If .EOF = True Then
            cmdLast.Enabled = False
            cmdNext.Enabled = False
            cmdPrevious.Enabled = True
            cmdFirst.Enabled = True
            cmdPrevious.SetFocus
        End If
        .MovePrevious
    Else
        cmdLast.Enabled = False
        cmdNext.Enabled = False
        cmdPrevious.Enabled = True
        cmdFirst.Enabled = True
        cmdPrevious.SetFocus
    End If
End With

End Sub




Private Sub txtCode_KeyPress(KeyAscii As Integer)
If Len(Trim(txtCode.Text)) > 19 Then
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

Private Sub txtComments_KeyPress(KeyAscii As Integer)
If Len(Trim(txtComments.Text)) > 198 Then
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

Private Sub txtEMail_KeyPress(KeyAscii As Integer)
If Len(Trim(txtEmail.Text)) > 49 Then
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

Private Sub txtHTelNo_KeyPress(KeyAscii As Integer)
If Len(Trim(txtHTelNo.Text)) > 29 Then
    Beep
    MsgBox "Can't enter more than 30 characters", vbExclamation
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

Private Sub txtIDNo_KeyPress(KeyAscii As Integer)
If Len(Trim(txtIDNo.Text)) > 19 Then
    Beep
    MsgBox "Can't enter more than 20 characters", vbExclamation
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

Private Sub txtMNo_KeyPress(KeyAscii As Integer)
If Len(Trim(txtMNo.Text)) > 19 Then
    Beep
    MsgBox "Can't enter more than 20 characters", vbExclamation
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

Private Sub txtOccupation_KeyPress(KeyAscii As Integer)
If Len(Trim(txtOccupation.Text)) > 99 Then
    Beep
    MsgBox "Can't enter more than 100 characters", vbExclamation
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

Private Sub txtONames_KeyPress(KeyAscii As Integer)
If Len(Trim(txtONames.Text)) > 99 Then
    Beep
    MsgBox "Can't enter more than 100 characters", vbExclamation
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

Private Sub txtOTelNo_KeyPress(KeyAscii As Integer)
If Len(Trim(txtOTelNo.Text)) > 49 Then
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

Private Sub txtSurName_KeyPress(KeyAscii As Integer)
If Len(Trim(txtSurname.Text)) > 99 Then
    Beep
    MsgBox "Can't enter more than 100 characters", vbExclamation
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

Private Function loadKinCode() As String
    Set rs5 = CConnect.GetRecordSet("SELECT MAX(id) FROM Kin")
    If rs5.EOF = False Then
        If rs5.RecordCount > 0 And Not IsNull(rs5.Fields(0)) Then
            loadKinCode = "KN" & CStr(rs5.Fields(0) + 1)
        Else
            loadKinCode = "KN01"
        End If
    Else
        loadKinCode = "KN01"
    End If
    Set rs5 = Nothing
End Function
