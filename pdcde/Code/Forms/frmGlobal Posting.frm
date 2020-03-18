VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmGlobalPosting 
   Appearance      =   0  'Flat
   BackColor       =   &H00F2FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Global Posting"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9975
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   1680
      TabIndex        =   62
      Top             =   6360
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   375
         Left            =   240
         TabIndex        =   66
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   3120
         TabIndex        =   65
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2160
         TabIndex        =   64
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   375
         Left            =   1200
         TabIndex        =   63
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraGLobOpts 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2FFFF&
      ForeColor       =   &H80000008&
      Height          =   6525
      Left            =   0
      TabIndex        =   18
      Top             =   -90
      Width           =   9960
      Begin MSComCtl2.DTPicker txtPeriod 
         Height          =   375
         Left            =   120
         TabIndex        =   67
         Top             =   1800
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   49348611
         CurrentDate     =   38730
      End
      Begin VB.ComboBox cboType 
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
         ItemData        =   "frmGlobal Posting.frx":0000
         Left            =   120
         List            =   "frmGlobal Posting.frx":000D
         TabIndex        =   59
         Top             =   1110
         Width           =   2175
      End
      Begin VB.OptionButton optOther 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FFFF&
         Caption         =   "Change Other Allowance"
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
         Height          =   225
         Left            =   135
         TabIndex        =   53
         Top             =   1140
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.OptionButton optTrans 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FFFF&
         Caption         =   "Change Transport Allowance"
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
         Height          =   225
         Left            =   135
         TabIndex        =   52
         Top             =   1095
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.OptionButton optLeave 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FFFF&
         Caption         =   "Change Leave Allowance"
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
         Height          =   225
         Left            =   135
         TabIndex        =   51
         Top             =   1170
         Visible         =   0   'False
         Width           =   2550
      End
      Begin VB.OptionButton optHse 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FFFF&
         Caption         =   "Change House Allowance"
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
         Height          =   225
         Left            =   135
         TabIndex        =   50
         Top             =   540
         Width           =   2595
      End
      Begin VB.Frame fraBPay 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FFFF&
         Caption         =   "Changes"
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
         Height          =   2040
         Left            =   2805
         TabIndex        =   43
         Top             =   135
         Width           =   2865
         Begin VB.OptionButton optDPerc 
            Appearance      =   0  'Flat
            BackColor       =   &H00F2FFFF&
            Caption         =   "Decrease by a percentage"
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
            Height          =   225
            Left            =   165
            TabIndex        =   48
            Top             =   1155
            Width           =   2565
         End
         Begin VB.OptionButton optDAmt 
            Appearance      =   0  'Flat
            BackColor       =   &H00F2FFFF&
            Caption         =   "Decrease by an amount"
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
            Height          =   225
            Left            =   180
            TabIndex        =   47
            Top             =   540
            Width           =   2655
         End
         Begin VB.TextBox txtPayAmt 
            Alignment       =   1  'Right Justify
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
            Left            =   945
            TabIndex        =   46
            Top             =   1575
            Width           =   1350
         End
         Begin VB.OptionButton optIAmt 
            Appearance      =   0  'Flat
            BackColor       =   &H00F2FFFF&
            Caption         =   "Increase by an amount"
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
            Height          =   225
            Left            =   180
            TabIndex        =   45
            Top             =   210
            Value           =   -1  'True
            Width           =   2505
         End
         Begin VB.OptionButton optIPerc 
            Appearance      =   0  'Flat
            BackColor       =   &H00F2FFFF&
            Caption         =   "Increase by a percentage"
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
            Height          =   225
            Left            =   165
            TabIndex        =   44
            Top             =   855
            Width           =   2565
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BackColor       =   &H00F2FFFF&
            Caption         =   "Value"
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
            Left            =   240
            TabIndex        =   49
            Top             =   1590
            Width           =   735
         End
      End
      Begin VB.OptionButton optPay 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FFFF&
         Caption         =   "Change Basic Pay"
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
         Height          =   225
         Left            =   135
         TabIndex        =   42
         Top             =   210
         Value           =   -1  'True
         Width           =   1845
      End
      Begin VB.Frame fraGlobal 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FFFF&
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
         Height          =   6240
         Left            =   1410
         TabIndex        =   38
         Top             =   6300
         Visible         =   0   'False
         Width           =   4530
         Begin VB.CommandButton CmdSelect 
            Caption         =   "Select"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3090
            TabIndex        =   41
            Top             =   5490
            Width           =   1335
         End
         Begin VB.CommandButton cmdDone 
            Caption         =   "Done"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3090
            TabIndex        =   39
            Top             =   5805
            Width           =   1335
         End
         Begin MSComctlLib.TreeView trwStruc2 
            Height          =   5205
            Left            =   105
            TabIndex        =   40
            Top             =   240
            Visible         =   0   'False
            Width           =   4305
            _ExtentX        =   7594
            _ExtentY        =   9181
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   529
            LabelEdit       =   1
            Style           =   7
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
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FFFF&
         Caption         =   "Period"
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
         Left            =   150
         TabIndex        =   60
         Top             =   1575
         Width           =   1665
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00F2FFFF&
         Caption         =   "Type of Change"
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
         Left            =   165
         TabIndex        =   58
         Top             =   855
         Width           =   1155
      End
   End
   Begin VB.Frame fraEmpsel 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2FFFF&
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
      Height          =   7020
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   9960
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear Selection"
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
         Left            =   75
         TabIndex        =   61
         Top             =   6555
         Width           =   1335
      End
      Begin MSComctlLib.ListView lvwEmp2 
         Height          =   5850
         Left            =   5520
         TabIndex        =   2
         Top             =   90
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   10319
         View            =   3
         LabelEdit       =   1
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
      Begin VB.Frame fraCategory 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FFFF&
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
         Height          =   1455
         Left            =   5655
         TabIndex        =   54
         Top             =   1830
         Width           =   3825
         Begin VB.ComboBox cboEmpCat 
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
            Left            =   990
            TabIndex        =   57
            Top             =   210
            Width           =   2715
         End
         Begin VB.CommandButton cmdCatOk 
            Caption         =   "Ok"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2460
            TabIndex        =   55
            Top             =   975
            Width           =   1245
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00F2FFFF&
            Caption         =   "Category:"
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
            TabIndex        =   56
            Top             =   195
            Width           =   735
         End
      End
      Begin MSComctlLib.ListView lvwEmp 
         Height          =   5850
         Left            =   0
         TabIndex        =   1
         Top             =   90
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   10319
         View            =   3
         LabelEdit       =   1
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
      Begin MSComctlLib.TreeView trwStruc 
         Height          =   5850
         Left            =   0
         TabIndex        =   37
         Top             =   90
         Visible         =   0   'False
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   10319
         _Version        =   393217
         Indentation     =   529
         LabelEdit       =   1
         Style           =   7
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
      Begin VB.Frame fraSpecialSel 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FFFF&
         ForeColor       =   &H80000008&
         Height          =   5940
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   4395
         Begin VB.ComboBox cboPayName2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1230
            TabIndex        =   33
            Top             =   2490
            Width           =   3075
         End
         Begin VB.ComboBox cboPayCode 
            Enabled         =   0   'False
            Height          =   315
            Left            =   90
            TabIndex        =   32
            Top             =   2490
            Width           =   1125
         End
         Begin VB.CheckBox chkPayRate 
            Appearance      =   0  'Flat
            BackColor       =   &H00F2FFFF&
            Caption         =   "Payment Rate"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   120
            TabIndex        =   31
            Top             =   2010
            Width           =   1545
         End
         Begin MSComCtl2.DTPicker dtp1 
            Height          =   345
            Left            =   150
            TabIndex        =   28
            Top             =   1500
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   49348609
            CurrentDate     =   38030
            MinDate         =   2
         End
         Begin VB.CheckBox chkDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00F2FFFF&
            Caption         =   "Date of Employment between"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   150
            TabIndex        =   27
            Top             =   1170
            Width           =   3015
         End
         Begin VB.TextBox txtB2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2370
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   660
            Width           =   1905
         End
         Begin VB.TextBox txtB1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   150
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   660
            Width           =   1785
         End
         Begin VB.CheckBox chkBasic 
            Appearance      =   0  'Flat
            BackColor       =   &H00F2FFFF&
            Caption         =   "Basic Pay between"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   150
            TabIndex        =   23
            Top             =   330
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker dtp2 
            Height          =   345
            Left            =   2340
            TabIndex        =   30
            Top             =   1500
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   49348609
            CurrentDate     =   38030
            MinDate         =   2
         End
         Begin VB.Label Label20 
            Appearance      =   0  'Flat
            BackColor       =   &H00F2FFFF&
            Caption         =   "Code"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   90
            TabIndex        =   35
            Top             =   2280
            Width           =   885
         End
         Begin VB.Label Label19 
            Appearance      =   0  'Flat
            BackColor       =   &H00F2FFFF&
            Caption         =   "Name"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1230
            TabIndex        =   34
            Top             =   2280
            Width           =   885
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H00F2FFFF&
            Caption         =   "And"
            Height          =   195
            Left            =   1950
            TabIndex        =   29
            Top             =   1560
            Width           =   330
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00F2FFFF&
            Caption         =   "And"
            Height          =   195
            Left            =   1980
            TabIndex        =   26
            Top             =   690
            Width           =   330
         End
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Finish"
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
         Height          =   375
         Left            =   8880
         TabIndex        =   20
         Top             =   6570
         Width           =   1000
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
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
         Left            =   7815
         TabIndex        =   19
         Top             =   6570
         Width           =   1080
      End
      Begin VB.TextBox txtCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4410
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   90
         Width           =   1095
      End
      Begin VB.Frame fraopts 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FFFF&
         Caption         =   "Employee selection"
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
         Height          =   585
         Left            =   0
         TabIndex        =   13
         Top             =   5940
         Width           =   9960
         Begin VB.OptionButton optCategory 
            Appearance      =   0  'Flat
            BackColor       =   &H00F2FFFF&
            Caption         =   "Employee Category"
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
            Height          =   225
            Left            =   7830
            TabIndex        =   21
            Top             =   240
            Width           =   1905
         End
         Begin VB.OptionButton optCdiv 
            Appearance      =   0  'Flat
            BackColor       =   &H00F2FFFF&
            Caption         =   "Select by company division"
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
            Height          =   225
            Left            =   4770
            TabIndex        =   16
            Top             =   240
            Width           =   2835
         End
         Begin VB.OptionButton optRandom 
            Appearance      =   0  'Flat
            BackColor       =   &H00F2FFFF&
            Caption         =   "Select from list"
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
            Height          =   225
            Left            =   90
            TabIndex        =   15
            Top             =   240
            Value           =   -1  'True
            Width           =   1845
         End
         Begin VB.OptionButton optRange 
            Appearance      =   0  'Flat
            BackColor       =   &H00F2FFFF&
            Caption         =   "Select by Range"
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
            Height          =   225
            Left            =   2370
            TabIndex        =   14
            Top             =   240
            Width           =   1905
         End
      End
      Begin MSComctlLib.ImageList imgTree 
         Left            =   4590
         Top             =   4020
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   26
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlobal Posting.frx":002C
               Key             =   "A"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlobal Posting.frx":047E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlobal Posting.frx":08D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlobal Posting.frx":0BEA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlobal Posting.frx":0F04
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlobal Posting.frx":1356
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlobal Posting.frx":1670
               Key             =   "B"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlobal Posting.frx":1AC2
               Key             =   "C"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlobal Posting.frx":1F14
               Key             =   "D"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlobal Posting.frx":2366
               Key             =   "E"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlobal Posting.frx":27B8
               Key             =   "F"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlobal Posting.frx":2C0A
               Key             =   "G"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlobal Posting.frx":305C
               Key             =   "H"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlobal Posting.frx":34AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlobal Posting.frx":3900
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlobal Posting.frx":3D52
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlobal Posting.frx":41A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlobal Posting.frx":45F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlobal Posting.frx":4A48
               Key             =   "I"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlobal Posting.frx":4E9A
               Key             =   "v"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlobal Posting.frx":52EC
               Key             =   "P"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlobal Posting.frx":573E
               Key             =   "Z"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlobal Posting.frx":5B90
               Key             =   "J"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlobal Posting.frx":5FE2
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlobal Posting.frx":60F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlobal Posting.frx":6206
               Key             =   "O"
            EndProperty
         EndProperty
      End
      Begin VB.Frame frarange 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FFFF&
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
         Height          =   1455
         Left            =   5655
         TabIndex        =   7
         Top             =   1830
         Width           =   2670
         Begin VB.CommandButton cmdOk 
            Caption         =   "Ok"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1590
            TabIndex        =   36
            Top             =   1020
            Width           =   975
         End
         Begin VB.TextBox txtTo 
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
            Height          =   315
            Left            =   960
            TabIndex        =   12
            Top             =   585
            Width           =   1605
         End
         Begin VB.TextBox txtFrom 
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
            Height          =   315
            Left            =   960
            TabIndex        =   11
            Top             =   195
            Width           =   1605
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00F2FFFF&
            Caption         =   "To:"
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
            Left            =   150
            TabIndex        =   9
            Top             =   645
            Width           =   240
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00F2FFFF&
            Caption         =   "From:"
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
            TabIndex        =   8
            Top             =   195
            Width           =   420
         End
      End
      Begin VB.Frame fracmds 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FFFF&
         BorderStyle     =   0  'None
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
         Height          =   2385
         Left            =   4290
         TabIndex        =   3
         Top             =   1350
         Width           =   1275
         Begin VB.CommandButton cmdRemove 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   150
            TabIndex        =   6
            Top             =   645
            Width           =   1000
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   " >"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   150
            TabIndex        =   5
            Top             =   315
            Width           =   1000
         End
         Begin VB.CommandButton cmdRemoveall 
            Caption         =   "<<"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   150
            TabIndex        =   4
            Top             =   1860
            Width           =   1000
         End
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00F2FFFF&
         Caption         =   "Select the range from the list"
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
         Left            =   5640
         TabIndex        =   10
         Top             =   1620
         Width           =   2085
      End
   End
End
Attribute VB_Name = "frmGlobalPosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare variables
'******************************************************************
Dim CEmp As String
Dim AppCode As String
Dim empcode As String
Dim LevelCode As String
Dim Rangeto As Boolean
Dim SelEmp As Long
Dim Fixed As String
Dim TRef As Long
Dim IsFormula As Boolean
Dim mySCode As String
Dim MyRow As Long
Dim CLoaded As Boolean
Dim myBPay As Currency

Private Sub cboEmpCat_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub chkBasic_Click()
    If chkBasic.Value = 1 Then
        txtB1.Locked = False
        txtB2.Locked = False
    Else
        txtB1.Text = ""
        txtB2.Text = ""
        txtB1.Locked = True
        txtB2.Locked = True
    End If
End Sub

Private Sub chkDate_Click()
    If chkDate.Value = 1 Then
        dtp1.Enabled = True
        dtp2.Enabled = True
    Else
        dtp1.Enabled = False
        dtp2.Enabled = False
    End If
End Sub


Private Sub cmdAdd_Click()
'On Error Resume Next
Dim mySeq As Integer
Dim myLevelCode As String
On Error GoTo errHandler
If lvwEmp.Visible = True Then
Dim Date1, Date2
    Set rs1 = CConnect.GetRecordSet("SELECT * FROM GPosting ORDER BY mySeq")
        With rs1
            If .RecordCount > 0 Then
                .MoveLast
                mySeq = !mySeq + 1
            Else
                mySeq = 1
            End If
        End With
    
    Set rs1 = Nothing
    
    Set rs1 = CConnect.GetRecordSet("SELECT * FROM GPosting WHERE employee_id='" & lvwEmp.SelectedItem.Tag & "'")
        With rs1
            If .RecordCount < 1 Then
                CConnect.ExecuteSql ("INSERT INTO GPosting (Employee_id, mySeq) VALUES ('" & lvwEmp.SelectedItem.Tag & "'," & mySeq & ")")
                Call LoadList2
            End If
        End With
    Set rs1 = Nothing
ElseIf trwStruc.Visible = True Then
    'cConnect.ExecuteSql ("DELETE FROM GPosting")
    
    Set rs1 = CConnect.GetRecordSet("SELECT * FROM SEmp WHERE LCode Like '" & LevelCode & "%'")
        
    With rs1
        If .RecordCount > 0 Then
            .MoveFirst
            Set rs6 = CConnect.GetRecordSet("SELECT * FROM GPosting")
'                MySeq = 1
            While Not .EOF
                If rs6.RecordCount > 0 Then
                    rs6.Filter = "employee_id like '" & !employee_id & "'"
                    If rs6.RecordCount < 1 Then
                        CConnect.ExecuteSql ("INSERT INTO GPosting (employee_id, mySeq) VALUES ('" & !employee_id & "'," & mySeq & ")")
                    End If
                    rs6.Filter = adFilterNone
                Else
                    CConnect.ExecuteSql ("INSERT INTO GPosting (employee_id, mySeq) VALUES ('" & !employee_id & "'," & mySeq & ")")
                End If
                .MoveNext
                mySeq = mySeq + 1
                rs6.Requery
            Wend
            Set rs6 = Nothing
        End If
    End With
        
    Set rs1 = Nothing
    
    Call LoadList2
ElseIf fraSpecialSel.Visible = True Then
    If chkBasic.Value = 1 Then
        If txtB1.Text = "" Or txtB2.Text = "" Then
            MsgBox "Enter the Basic Pay range", vbInformation
            Exit Sub
        End If
    End If
    If chkPayRate.Value = 1 Then
        If cboPayCode.Text = "" Then
            MsgBox "Enter the Payment rate Code", vbInformation
            cboPayCode.SetFocus
            Exit Sub
        End If
    End If
    
    Date1 = Year(dtp1.Value) & Format(Month(dtp1.Value), "00") & Format(Day(dtp1.Value), "00")
    
    Date2 = Year(dtp2.Value) & Format(Month(dtp2.Value), "00") & Format(Day(dtp2.Value), "00")
    
    If chkBasic.Value = 1 And chkDate.Value = 0 And chkPayRate.Value = 0 Then
        mySQL = "SELECT GPosting.employee_id FROM GPosting LEFT JOIN Employee ON GPosting.employee_id=Employee.employee_id WHERE BasicPay BETWEEN " & Format(txtB1.Text, Nfmt) & " AND " & Format(txtB2.Text, Nfmt) & " "
    ElseIf chkBasic.Value = 1 And chkDate.Value = 1 And chkPayRate.Value = 0 Then
'        If AccessDB = True Then
            mySQL = "SELECT GPosting.employee_id FROM GPosting LEFT JOIN Employee ON GPosting.employee_id=Employee.employee_id WHERE (BasicPay BETWEEN " & Format(txtB1.Text, Nfmt) & " AND " & Format(txtB2.Text, Nfmt) & ") AND (((Employee.DEmployed) Between " & "#" & dtp1.Value & "#" & " And " & "#" & dtp2.Value & "#" & ")) "
'        Else
'            mysql = "SELECT GPosting.employee_id FROM GPosting LEFT JOIN Employee ON GPosting.employee_id=Employee.employee_id WHERE BasicPay BETWEEN " & Format(txtB1.Text, Nfmt) & " AND " & Format(txtB2.Text, Nfmt) & " AND (((Employee.DEmployed) Between '" & Date1 & "' And '" & Date2 & "')) "
'        End If
    ElseIf chkBasic.Value = 1 And chkDate.Value = 1 And chkPayRate.Value = 1 Then
'        If AccessDB = True Then
            mySQL = "SELECT GPosting.employee_id FROM GPosting LEFT JOIN Employee ON GPosting.employee_id=Employee.employee_id WHERE ((BasicPay BETWEEN " & Format(txtB1.Text, Nfmt) & " AND " & Format(txtB2.Text, Nfmt) & ") AND (((Employee.DEmployed) Between " & "#" & dtp1.Value & "#" & " And " & "#" & dtp2.Value & "#" & ")) AND RateCode='" & cboPayCode.Text & "') "
'        Else
'            mysql = "SELECT GPosting.employee_id FROM GPosting LEFT JOIN Employee ON GPosting.employee_id=Employee.employee_id WHERE ((BasicPay BETWEEN " & Format(txtB1.Text, Nfmt) & " AND " & Format(txtB2.Text, Nfmt) & ") AND (((Employee.DEmployed) Between '" & Date1 & "' And '" & Date2 & "')) AND RateCode='" & cboPayCode.Text & "') "
'        End If
    ElseIf chkBasic.Value = 0 And chkDate.Value = 1 And chkPayRate.Value = 0 Then
'        If AccessDB = True Then
            mySQL = "SELECT GPosting.employee_id FROM GPosting LEFT JOIN Employee ON GPosting.employee_id=Employee.employee_id WHERE (((Employee.DEmployed) Between " & "#" & dtp1.Value & "#" & " And " & "#" & dtp2.Value & "#" & ")) "
'        Else
'            mysql = "SELECT GPosting.employee_id FROM GPosting LEFT JOIN Employee ON GPosting.employee_id=Employee.employee_id WHERE (((Employee.DEmployed) Between '" & Date1 & "' And '" & Date2 & "')) "
'        End If
    ElseIf chkBasic.Value = 0 And chkDate.Value = 0 And chkPayRate.Value = 1 Then
        mySQL = "SELECT GPosting.employee_id FROM GPosting LEFT JOIN Employee ON GPosting.employee_id=Employee.employee_id WHERE RateCode='" & cboPayCode.Text & "' "
    ElseIf chkBasic.Value = 0 And chkDate.Value = 1 And chkPayRate.Value = 1 Then
'        If AccessDB = True Then
            mySQL = "SELECT GPosting.employee_id FROM GPosting LEFT JOIN Employee ON GPosting.employee_id=Employee.employee_id WHERE (((Employee.DEmployed) Between " & "#" & dtp1.Value & "#" & " And " & "#" & dtp2.Value & "#" & ") AND RateCode='" & cboPayCode.Text & "') "
'        Else
'            mysql = "SELECT GPosting.employee_id FROM GPosting LEFT JOIN Employee ON GPosting.employee_id=Employee.employee_id WHERE (((Employee.DEmployed) Between '" & Date1 & "' And '" & Date2 & "') AND RateCode='" & cboPayCode.Text & "') "
'        End If
    ElseIf chkBasic.Value = 1 And chkDate.Value = 0 And chkPayRate.Value = 1 Then
        mySQL = "SELECT GPosting.employee_id FROM GPosting LEFT JOIN EmpPayment ON GPosting.employee_id=EmpPayment.employee_id WHERE ((BasicPay BETWEEN " & Format(txtB1.Text, Nfmt) & " AND " & Format(txtB2.Text, Nfmt) & ") AND PRCode='" & cboPayCode.Text & "') "
    Else
        MsgBox "Select a criteria to use", vbInformation
        Exit Sub
    End If
    Set rs1 = CConnect.GetRecordSet(mySQL)
        With rs1
            If .RecordCount > 0 Then
                .MoveFirst
                While Not .EOF
                    CConnect.ExecuteSql ("INSERT INTO GPosting (employee_id,mySeq)VALUES ('" & rs1!empcode & "',1.1)")
                    .MoveNext
                Wend
            End If
        End With
    Set rs1 = Nothing
    'cConnect.ExecuteSql ("DELETE FROM GPosting WHERE mySeq<>1.1")
    CConnect.ExecuteSql ("UPDATE GPosting SET mySeq=1 WHERE mySeq=1.1")
    Call LoadList2
End If
cmdRemoveall.Enabled = True
txtCount.Text = SelEmp
Exit Sub
errHandler:
MsgBox Err.Description, vbInformation
End Sub


Private Sub cmdDiv_Click()
    fraGlobal.Visible = True
    lvwUTrans.Visible = False
    trwStruc2.Visible = True
    lvwBBranch.Visible = False
    lvwBanks.Visible = False
    
End Sub



Private Sub cmdCatOk_Click()
    If cboEmpCat.Text <> "" Then
        frmMain2.Filter_TermCat_ByRights "ECategory", cboEmpCat.Text
        CConnect.ExecuteSql ("DELETE FROM GPosting")
        Set rs6 = CConnect.GetRecordSet("SELECT * FROM GPosting")
        With rsGlob2
            If .RecordCount > 0 Then
                .Filter = adFilterNone
                .Filter = "ECategory like '" & cboEmpCat.Text & "'"
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                        With rs6
                            If .RecordCount > 0 Then
                                .Filter = "employee_id like '" & rsGlob2!employee_id & "'"
                                If .RecordCount < 1 Then
                                    CConnect.ExecuteSql ("INSERT INTO GPosting (employee_id) VALUES ('" & rsGlob2!employee_id & "')")
                                End If
                                .Filter = adFilterNone
                            Else
                                CConnect.ExecuteSql ("INSERT INTO GPosting (employee_id) VALUES ('" & rsGlob2!employee_id & "')")
                            End If
                        End With
                        
                        .MoveNext
                    Loop
                    
                End If
                .Filter = adFilterNone
                
                Call LoadList2
                txtCount.Text = lvwEmp2.ListItems.Count
                lvwEmp2.Visible = True
                optRandom.Value = True
                Me.MousePointer = vbDefault
                                
            End If
        End With
        
        Set rs6 = Nothing
    
    End If
    
End Sub

Private Sub cmdClear_Click()
    CConnect.ExecuteSql ("DELETE FROM GPosting")
    lvwEmp2.ListItems.Clear
    
End Sub

Public Sub cmdDelete_Click()

End Sub

Private Sub cmdDone_Click()
    fraGlobal.Visible = False
End Sub

Public Sub cmdEdit_Click()

End Sub

Public Sub cmdNew_Click()

End Sub

Private Sub cmdNext_Click()

If cmdNext.Caption = "Next" Then
    If cboType.Text = "" Then
        MsgBox "You must select increment type.", vbInformation
        cboType.SetFocus
        Exit Sub
    End If
    
'''    If txtPeriod.Text = "" Then
'''        MsgBox "You must specify period.", vbInformation
'''        txtPeriod.SetFocus
'''        Exit Sub
'''    End If
    
    If txtPayAmt.Text = "" Then
        MsgBox "You must enter value", vbInformation
        txtPayAmt.SetFocus
        Exit Sub
    End If
    
    If txtPayAmt.Text = 0 Then
        MsgBox "Value cannot be zero.", vbInformation
        txtPayAmt.SetFocus
        Exit Sub
    End If
    
    If Val(txtPayAmt.Text) < 0 And (optIPerc.Value = True Or optDPerc.Value = True) Then
        MsgBox "Value cannot be less than zero.", vbInformation
        txtPayAmt.SetFocus
        Exit Sub
    End If
    
    cmdSave.Enabled = True
    fraGLobOpts.Visible = False
    fraEmpsel.Visible = True
    cmdSave.Visible = True
    cmdNext.Caption = "Back"
Else
    cmdSave.Enabled = False
    fraGLobOpts.Visible = True
    cmdNext.Caption = "Next"
   
End If


End Sub

Private Sub CmdOk_Click()
    If txtFrom.Text = "" Then
        MsgBox "Select the employee range", vbInformation
        txtFrom.SetFocus
        Exit Sub
    End If
    
    If txtTo.Text = "" Then
        MsgBox "Select the employee range", vbInformation
        txtTo.SetFocus
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    
    With rsGlob2
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "EmpCode='" & txtFrom.Text & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                'cConnect.ExecuteSql ("DELETE FROM GPosting")
                Set rs6 = CConnect.GetRecordSet("SELECT * FROM GPosting")
                
                With rs6
                    If .RecordCount > 0 Then
                        .Filter = "employee_id like '" & rsGlob2!employee_id & "'"
                        If .RecordCount < 1 Then
                            myCount = 1
                            Do While Not rsGlob2.EOF And rsGlob2!empcode <> txtTo.Text
                                CConnect.ExecuteSql ("INSERT INTO GPosting (employee_id, mySeq) VALUES ('" & rsGlob2!employee_id & "'," & myCount & ")")
                                myCount = myCount + 1
                                rsGlob2.MoveNext
                            Loop
                            
                            CConnect.ExecuteSql ("INSERT INTO GPosting (employee_id, mySeq) VALUES ('" & rsGlob2!employee_id & "'," & myCount & ")")
                        End If
                        .Filter = adFilterNone
                    Else
                        myCount = 1
                        Do While Not rsGlob2.EOF And rsGlob2!empcode <> txtTo.Text
                            CConnect.ExecuteSql ("INSERT INTO GPosting (EmpCode, mySeq) VALUES ('" & rsGlob2!empcode & "'," & myCount & ")")
                            myCount = myCount + 1
                            rsGlob2.MoveNext
                        Loop
                        
                        CConnect.ExecuteSql ("INSERT INTO GPosting (EmpCode, mySeq) VALUES ('" & rsGlob2!empcode & "'," & myCount & ")")

                    End If
                End With
                
                Set rs6 = Nothing
            Else
                MsgBox "No records were found in the range specified.", vbInformation
                Exit Sub
            End If
            
        End If
    End With
        
    Call LoadList2
    lvwEmp2.Visible = True
    optRandom.Value = True
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdRemove_Click()
If lvwEmp2.ListItems.Count < 1 Then Exit Sub
        CConnect.ExecuteSql ("DELETE FROM GPosting WHERE employee_id='" & rsGlob!employee_id & "'")
    
    Call ReAssignSeq
    Call LoadList2
    txtCount.Text = SelEmp
End Sub

Private Sub cmdRemoveall_Click()
If MsgBox("Are you sure you want to remove all the selected employees?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

CConnect.ExecuteSql ("DELETE FROM GPosting")
Call LoadList2
cmdRemoveall.Enabled = False
txtCount.Text = SelEmp
End Sub

Public Sub cmdSave_Click()
Dim NewAmount As Currency

Me.MousePointer = vbHourglass
        
'    Get Records of all selected employees
Set rs1 = CConnect.GetRecordSet("SELECT Employee.*, CStructure.Code, CStructure.Description" & _
                                " FROM (Employee LEFT JOIN SEmp ON Employee.Employee_id = SEmp.Employee_id) LEFT JOIN CStructure ON SEmp.LCode = CStructure.LCode" & _
                                " WHERE (((Employee.Employee_id) In (SELECT Employee_id FROM GPosting)))" & _
                                " ORDER BY Employee.EmpCode")

'    Get previous salary changes for selected employees
Set rs2 = CConnect.GetRecordSet("SELECT * From pdSalaryChange " & _
                    " WHERE (((pdSalaryChange.Employee_id) In (SELECT Employee_id FROM GPosting)))")

'    Get records of job progressions for the selected period for the selected employees
Set rs5 = CConnect.GetRecordSet("SELECT * From JProg " & _
                    " WHERE (((JProg.Employee_id) In (SELECT Employee_id FROM GPosting))) AND month(JProg.Cdate) = '" & Month(txtPeriod) & "' AND year(JProg.cdate) = " & Year(txtPeriod))

'    Get records of employees selected
Set rs3 = CConnect.GetRecordSet("SELECT  e.* From gposting as j LEFT JOIN employee as e ON j.employee_id = e.employee_id ")

With rs1
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
            
            With rs3
                If .RecordCount > 0 Then
                    .Filter = "employee_id like '" & rs1!employee_id & "'"
                    If .RecordCount > 0 Then
                        If optPay.Value = True Then
                            If optIAmt = True Then
                                NewAmount = !BasicPay + Format(txtPayAmt.Text, Nfmt)
                            ElseIf optDAmt = True Then
                                NewAmount = !BasicPay - Format(txtPayAmt.Text, Nfmt)
                            ElseIf optIPerc = True Then
                                NewAmount = !BasicPay + (!BasicPay * Format(txtPayAmt.Text, Nfmt) / 100)
                            ElseIf optDPerc = True Then
                                NewAmount = !BasicPay - (!BasicPay * Format(txtPayAmt.Text, Nfmt) / 100)
                            End If
                            
                            If NewAmount <> Int(NewAmount) Then
                                NewAmount = Int(NewAmount) + 1
                            End If
                                            
                            Call updateSalaryChanges(CDbl(NewAmount), !hallow, !tallow, !oallow, !lallow, !employee_id, cboType.Text)
                            
                            Action = "CHANGED EMPLOYEE BASIC SALARY; EMPLOYEE CODE: " & rs3!empcode & "; NEW BASIC PAY: " & NewAmount & "; INCREMENT TYPE: " & cboType.Text
                            
                            CConnect.ExecuteSql ("UPDATE Employee SET BasicPay = " & NewAmount & " WHERE employee_id = '" & rs3!employee_id & "'")
                            
                        ElseIf optTrans.Value = True Then
                            If optIAmt = True Then
                                NewAmount = !tallow + Format(txtPayAmt.Text, Nfmt)
                            ElseIf optDAmt = True Then
                                NewAmount = !tallow - Format(txtPayAmt.Text, Nfmt)
                            ElseIf optIPerc = True Then
                                NewAmount = !tallow + (!tallow * Format(txtPayAmt.Text, Nfmt) / 100)
                            ElseIf optDPerc = True Then
                                NewAmount = !tallow - (!tallow * Format(txtPayAmt.Text, Nfmt) / 100)
                            End If
                                            
                            If NewAmount <> Int(NewAmount) Then
                                NewAmount = Int(NewAmount) + 1
                            End If
                                            
                            Call updateSalaryChanges(!BasicPay, !hallow, CDbl(NewAmount), !oallow, !lallow, !employee_id)
                            Action = "CHANGED EMPLOYEE TRANSPORT ALLOWANCE; EMPLOYEE CODE: " & rs3!empcode & "; NEW TRANSPORT ALLOWANCE: " & NewAmount & "; INCREMENT TYPE: " & cboType.Text
                            CConnect.ExecuteSql ("UPDATE Employee SET TAllow = " & NewAmount & " WHERE employee_id = '" & rs3!employee_id & "'")
                        
                        ElseIf optHse.Value = True Then
                            If optIAmt = True Then
                                NewAmount = !hallow + Format(txtPayAmt.Text, Nfmt)
                            ElseIf optDAmt = True Then
                                NewAmount = !hallow - Format(txtPayAmt.Text, Nfmt)
                            ElseIf optIPerc = True Then
                                NewAmount = !hallow + (!hallow * Format(txtPayAmt.Text, Nfmt) / 100)
                            ElseIf optDPerc = True Then
                                NewAmount = !hallow - (!hallow * Format(txtPayAmt.Text, Nfmt) / 100)
                            End If
                                            
                            If NewAmount <> Int(NewAmount) Then
                                NewAmount = Int(NewAmount) + 1
                            End If
                                            
                            Call updateSalaryChanges(!BasicPay, CDbl(NewAmount), !tallow, !oallow, !lallow, !employee_id)
                            Action = "CHANGED EMPLOYEE HOUSE ALLOWANCE; EMPLOYEE CODE: " & rs3!empcode & "; NEW HOUSE ALLOWANCE: " & NewAmount & "; INCREMENT TYPE: " & cboType.Text
                            CConnect.ExecuteSql ("UPDATE Employee SET HAllow = " & NewAmount & " WHERE employee_id = '" & rs3!employee_id & "'")
                                            
                        ElseIf optOther.Value = True Then
                            If optIAmt = True Then
                                NewAmount = !oallow + Format(txtPayAmt.Text, Nfmt)
                            ElseIf optDAmt = True Then
                                NewAmount = !oallow - Format(txtPayAmt.Text, Nfmt)
                            ElseIf optIPerc = True Then
                                NewAmount = !oallow + (!oallow * Format(txtPayAmt.Text, Nfmt) / 100)
                            ElseIf optDPerc = True Then
                                NewAmount = !oallow - (!oallow * Format(txtPayAmt.Text, Nfmt) / 100)
                            End If
                                            
                            If NewAmount <> Int(NewAmount) Then
                                NewAmount = Int(NewAmount) + 1
                            End If
                                            
                            Call updateSalaryChanges(!BasicPay, !hallow, !tallow, CDbl(NewAmount), !lallow, !employee_id)
                            Action = "CHANGED EMPLOYEE OTHER ALLOWANCES; EMPLOYEE CODE: " & rs3!empcode & "; NEW OTHER ALLOWANCES: " & NewAmount & "; INCREMENT TYPE: " & cboType.Text
                            
                            CConnect.ExecuteSql ("UPDATE Employee SET OAllow = " & NewAmount & " WHERE employee_id = '" & rs3!employee_id & "'")
                            
                        ElseIf optLeave.Value = True Then
                            If optIAmt = True Then
                                NewAmount = !lallow + Format(txtPayAmt.Text, Nfmt)
                            ElseIf optDAmt = True Then
                                NewAmount = !lallow - Format(txtPayAmt.Text, Nfmt)
                            ElseIf optIPerc = True Then
                                NewAmount = !lallow + (!lallow * Format(txtPayAmt.Text, Nfmt) / 100)
                            ElseIf optDPerc = True Then
                                NewAmount = !lallow - (!lallow * Format(txtPayAmt.Text, Nfmt) / 100)
                            End If
                                            
                            If NewAmount <> Int(NewAmount) Then
                                NewAmount = Int(NewAmount) + 1
                            End If
                                            
                            Call updateSalaryChanges(!BasicPay, !hallow, !tallow, !oallow, CDbl(NewAmount), !employee_id)
                            Action = "CHANGED EMPLOYEE LEAVE ALLOWANCE; EMPLOYEE CODE: " & rs3!empcode & "; NEW LEAVE ALLOWANCE: " & NewAmount & "; INCREMENT TYPE: " & cboType.Text
                            
                            CConnect.ExecuteSql ("UPDATE Employee SET LAllow = " & NewAmount & " WHERE employee_id = '" & rs3!employee_id & "'")
                            
                        End If
                        
                    End If
                    .Filter = adFilterNone
                End If
            End With
                    
            .MoveNext
        Loop
    Else
        Me.MousePointer = 0
        MsgBox "No records selected.", vbExclamation
        Set rs1 = Nothing
        
        Exit Sub
    End If
End With


Set rs1 = Nothing

txtPayAmt.Text = 0

Me.MousePointer = 0

MsgBox "Process completed successfully.", vbInformation

End Sub

Private Sub cmdSelect_Click()
'    If optDiv.Value = True Then
'        If lvwDiv.ListItems.Count < 1 Then
'            fraGlobal.Visible = False
'            Exit Sub
'        End If
'        Set rs1 = cConnect.GetRecordSet("SELECT * FROM CompDivisions WHERE DivUnderCode='" & lvwDiv.SelectedItem & "'")
'            With rs1
'                If .RecordCount > 0 Then
'                    Call LoadDiv
'                Else
'                    Set rs1 = Nothing
'
'                    Set rs1 = cConnect.GetRecordSet("SELECT DivCode, Name, LevelCode FROM CompDivisions WHERE DivCode='" & lvwDiv.SelectedItem & "'")
'                    With rs1
'                        If .RecordCount > 0 Then
'                            If Not IsNull(!LevelCode) Then LevelCode = !LevelCode
'                            txtDiv.Text = !DivCode & ""
'                            lblDName.Caption = !Name & ""
'                            fraGlobal.Visible = False
'                        End If
'                    End With
'                End If
'            End With
'
'        Set rs1 = Nothing
    If optDiv.Value = True Then
        Set rs1 = CConnect.GetRecordSet("SELECT Code, Description, LCode,SCode FROM CStructure WHERE LCode='" & trwStruc2.SelectedItem.Key & "'")
            With rs1
                If .RecordCount > 0 Then
                    If Not IsNull(!LCode) Then LevelCode = !LCode
                    txtDiv.Text = !code & ""
                    mySCode = !scode & ""
                    lblDName.Caption = !Description & ""
                    fraGlobal.Visible = False
                End If
            End With
        Set rs1 = Nothing
'    ElseIf optTrans.Value = True Then
'        If lvwUTrans.ListItems.Count < 1 Then
'            fraGlobal.Visible = False
'            Exit Sub
'        End If
'       Set rs1 = CConnect.GetRecordSet("SELECT RefCode,  Description, Calculated FROM UserTrans WHERE RefCode='" & lvwUTrans.SelectedItem & "'")
'        With rs1
'            If .RecordCount > 0 Then
'                lblTName.Caption = !Description & ""
'                txtTrans.Text = !RefCode & ""
'                If !Calculated = "Yes" Then
'                    IsFormula = True
'                    txtAmount.Text = 0
'                    txtAmount.Locked = True
'                Else
'                    IsFormula = False
'                End If
'            End If
'        End With
'
'       Set rs1 = Nothing
'       fraGlobal.Visible = False
'    ElseIf optBank.Value = True Then
'        If lvwBanks.Visible = True Then
'            If lvwBanks.ListItems.Count < 1 Then
'                fraGlobal.Visible = False
'                Exit Sub
'            End If
'            Call LoadBBranches
'            lvwBanks.Visible = False
'            lvwBBranch.Visible = True
'        Else
'            If lvwBBranch.ListItems.Count < 1 Then
'                fraGlobal.Visible = False
'                Exit Sub
'            End If
'            Set rs4 = CConnect.GetRecordSet("SELECT BranchCode,Name FROM BBranch WHERE BranchCode='" & lvwBBranch.SelectedItem & "'")
'                With rs4
'                    If .RecordCount > 0 Then
'                        .MoveFirst
'                        txtBCode.Text = !BranchCode & ""
'                        txtBName.Text = !Name & ""
'                    End If
'                End With
'            Set rs4 = Nothing
'            fraGlobal.Visible = False
'        End If
    End If
    
End Sub

Private Sub cmdTrans_Click()
    fraGlobal.Visible = True
    lvwUTrans.Visible = True
    lvwDiv.Visible = False
    lvwBBranch.Visible = False
    lvwBanks.Visible = False
End Sub


Private Sub Command1_Click()

End Sub

Private Sub dtp2_Change()
If dtp2.Value < dtp1.Value Then dtp2.Value = dtp1.Value
End Sub
Private Sub dtp1_Change()
    If dtp1.Value > dtp2.Value Then dtp1.Value = dtp2.Value
End Sub




Private Sub Form_Load()
frmMain2.txtDetails.Caption = ""
Decla.Security Me

oSmart.FReset Me

frmMain2.txtDetails.Caption = ""

If oSmart.hRatio > 1.1 Then
    With frmMain2
        Me.Move .tvwMain.Width + .tvwMain.Width / 32.75, (.Height / 5.52) ' - 155
    End With
Else
     With frmMain2
        Me.Move .tvwMain.Width + .tvwMain.Width / 32.75, .Height / 5.55
    End With
    
End If

CConnect.CColor Me, MyColor



CConnect.ExecuteSql ("DELETE FROM GPosting")

Set rs1 = CConnect.GetRecordSet("SELECT * FROM ECategory")

With rs1
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
            cboEmpCat.AddItem !code & ""
            .MoveNext
        Loop
    End If
End With

Set rs1 = Nothing
    

Call InitGrid
Call LoadList

Call LoadDiv


txtCount.Text = 0


With rsGlob
    If .RecordCount < 1 Then
        Call DisableCmd
    End If
End With

Trial2 = True
 txtPeriod = Now

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

    
End Sub

Public Sub InitGrid()
With lvwEmp
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "Code", 1000
    .ColumnHeaders.Add , , "Name", 3000
    .ColumnHeaders.Add , , "Designation", 2000
    .ColumnHeaders.Add , , "Department", 2000
    
    .View = lvwReport
End With

With lvwEmp2
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "Code"
    .ColumnHeaders.Add , , "Name", 4000
    .ColumnHeaders.Add , , "Designation", 2000
    .ColumnHeaders.Add , , "Department", 2000
    
    .View = lvwReport
End With

End Sub

Public Sub LoadList()
Dim i As Integer
i = 0
lvwEmp.ListItems.Clear

With rsGlob
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
            If i = 0 Then
                Set LI = lvwEmp.ListItems.Add(, , !empcode & "", , 2)
                LI.Tag = !employee_id
                i = 1
            Else
                Set LI = lvwEmp.ListItems.Add(, , !empcode & "", , 3)
                LI.Tag = !employee_id
                i = 0
            End If
            LI.ListSubItems.Add , , !SurName & " " & !OtherNames & ""
            LI.ListSubItems.Add , , !Desig & ""
            LI.ListSubItems.Add , , !Description & ""
            
            .MoveNext
        Loop
        .MoveFirst
    End If
End With

End Sub

Private Sub Form_Resize()
oSmart.FResize Me
End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmMain2.Caption = "Personnel Director " & App.FileDescription
    
End Sub


Private Sub fraTrans_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub lvwBanks_DblClick()
    Call cmdSelect_Click
End Sub


Private Sub lvwBBranch_DblClick()
    Call cmdSelect_Click
End Sub

Private Sub lvwDiv_DblClick()
Call cmdSelect_Click
End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub lvwEmp_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwEmp
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
        End If
    End With

End Sub

Private Sub lvwEmp_DblClick()
If optRandom.Value = True Then
    If cmdAdd.Enabled = True Then cmdAdd_Click
Else
    If Rangeto = False Then
        txtFrom.Text = lvwEmp.SelectedItem
        If txtTo.Enabled = True Then txtTo.SetFocus
    Else
        If txtTo.Enabled = True Then txtTo.Text = lvwEmp.SelectedItem
    End If
End If
End Sub

Private Sub lvwEmp_ItemClick(ByVal Item As MSComctlLib.ListItem)
If lvwEmp.ListItems.Count > 0 Then
With rsGlob
    If .RecordCount > 0 Then
        .MoveFirst
        .Find "employee_id like '" & lvwEmp.SelectedItem.Tag & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            frmMain2.txtDetails.Caption = "Code: " & rsGlob!empcode & "     " & "Name: " & !SurName & "" & " " & !OtherNames & "" & " " & vbCrLf & _
                "" & "ID No:" & " " & !IdNo & "" & "     " & "Date Employed:" & " " & !DEmployed & "" & "     " & "Gender:" & " " & !Gender & ""
        Else
            frmMain2.txtDetails.Caption = ""
        End If
    End If
End With
End If
End Sub

Private Sub lvwEmp2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwEmp2
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvwEmp2_DblClick()
    Call cmdRemove_Click
End Sub

Public Sub ReAssignSeq()
Dim i As Integer
Set rs3 = CConnect.GetRecordSet("SELECT * FROM GPosting ORDER BY mySeq")
        With rs3
            If .RecordCount > 0 Then
                .MoveFirst
                i = 1
                While Not .EOF
                   !mySeq = i
                   .MoveNext
                   i = i + 1
                Wend
                
            End If
        End With
    
Set rs3 = Nothing
End Sub


Public Sub LoadList2()
Dim i As Integer

On Error GoTo errHandler
lvwEmp2.ListItems.Clear
Set rs2 = CConnect.GetRecordSet("SELECT * FROM GPosting ORDER BY employee_id")

With rs2
    SelEmp = .RecordCount
    lvwEmp2.Refresh
    If .RecordCount > 0 Then
        .MoveFirst
        i = 5
        
        Do While Not .EOF
            With rsGlob
                If .RecordCount > 0 Then
                    .MoveFirst
                    .Find "employee_id = '" & rs2!employee_id & "'", , adSearchForward, adBookmarkFirst
                    If .EOF = False Then
                        Set LI = lvwEmp2.ListItems.Add(, , !empcode & "", , i)
                        LI.ListSubItems.Add , , !SurName & " " & !OtherNames & ""
                        LI.ListSubItems.Add , , !Desig & ""
                        LI.ListSubItems.Add , , !Description & ""
                    Else
                        CConnect.ExecuteSql ("DELETE FROM GPosting WHERE employee_id='" & rs2!employee_id & "'")
                    End If
                End If
            End With
            
            .MoveNext
        Loop
    End If
End With

Set rs2 = Nothing
Exit Sub
errHandler:
MsgBox Err.Description, vbInformation

End Sub

Private Sub lvwEmp2_ItemClick(ByVal Item As MSComctlLib.ListItem)
If lvwEmp2.ListItems.Count > 0 Then
With rsGlob
    If .RecordCount > 0 Then
        .MoveFirst
        .Find "EmpCode like '" & lvwEmp2.SelectedItem & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            frmMain2.txtDetails.Caption = "Code: " & rsGlob!empcode & "     " & "Name: " & !SurName & "" & " " & !OtherNames & "" & " " & vbCrLf & _
                "" & "ID No:" & " " & !IdNo & "" & "     " & "Date Employed:" & " " & !DEmployed & "" & "     " & "Gender:" & " " & !Gender & ""
        Else
            frmMain2.txtDetails.Caption = ""
        End If
    End If
End With
End If
End Sub

Private Sub lvwUTrans_DblClick()
    Call cmdSelect_Click
End Sub

Public Sub LoadBBranches()
'
'lvwBBranch.ListItems.Clear
'Set rs1 = CConnect.GetRecordSet("SELECT * FROM BBranch WHERE BranchCode Like'" & lvwBanks.SelectedItem & "%'")
'    With rs1
'        If .RecordCount > 0 Then
'            .MoveFirst
'            While Not .EOF
'                Set LI = lvwBBranch.ListItems.Add(, , !BranchCode)
'                    LI.ListSubItems.Add , , !Name & "", , !Name & ""
'                    LI.ListSubItems.Add , , !Address & "", , !Address & ""
'                    .MoveNext
'            Wend
'        End If
'    End With
'
'Set rs1 = Nothing

End Sub


Public Sub cmdCancel_Click()

End Sub

Private Sub optCategory_Click()
lvwEmp.Visible = True
lvwEmp2.Visible = False
fracmds.Visible = False
fraCategory.Visible = True
frarange.Visible = False
Rangeto = False
cboEmpCat.Text = ""
txtCount.Text = SelEmp

End Sub

Private Sub optCdiv_Click()
lvwEmp.Visible = False
trwStruc.Visible = True
lvwEmp2.Visible = True
fracmds.Visible = True
fraSpecialSel.Visible = False
End Sub



Private Sub optRandom_Click()
lvwEmp.Visible = True
lvwEmp2.Visible = True
fracmds.Visible = True
End Sub

Private Sub optRange_Click()
lvwEmp.Visible = True
lvwEmp2.Visible = False
fracmds.Visible = False
fraSpecialSel.Visible = False
Rangeto = False
fraCategory.Visible = False
frarange.Visible = True
If txtFrom.Enabled = True Then txtFrom.SetFocus
txtCount.Text = SelEmp
End Sub




Private Sub optSpecial_Click()

End Sub

Private Sub trwStruc_Click()
    frmMain2.cboCat.Text = "All Records"
    frmMain2.cboTerms = "All Records"
    frmMain2.Filter_TermCat_ByRights "ECategory", "All Records"
    LevelCode = trwStruc.SelectedItem.Key
End Sub

Private Sub txtAmount_LostFocus()
    txtAmount.Text = Format(txtAmount.Text, Cfmt)
End Sub

Private Sub txtB1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
        Case Asc(".")
        Case Is = 8
        
        Case Else
        Beep
        KeyAscii = 0
  End Select
End Sub

Private Sub txtB2_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case Asc(".")
    Case Is = 8
    
    Case Else
    Beep
    KeyAscii = 0
  End Select
End Sub

Private Sub txtB2_LostFocus()
    If Val(Format(txtB2.Text, Nfmt)) < Val(Format(txtB1.Text, Nfmt)) Then
        txtB2.Text = txtB1.Text
    Else
        txtB2.Text = Format(txtB2.Text, Cfmt)
    End If
End Sub
Private Sub txtB1_LostFocus()
    If Val(Format(txtB1.Text, Nfmt)) > Val(Format(txtB1.Text, Nfmt)) Then
        txtB1.Text = txtB2.Text
    Else
        txtB1.Text = Format(txtB1.Text, Cfmt)
    End If
End Sub

Private Sub txtDiv_LostFocus()
    If txtDiv.Text = "" Then Exit Sub
    Set rs1 = CConnect.GetRecordSet("SELECT Name FROM CompDivisions WHERE DivCode='" & txtDiv.Text & "'")
        With rs1
            If .RecordCount > 0 Then
                lblDName.Caption = !Name & ""
            Else
                txtDiv.Text = ""
                Call cmdDiv_Click
                txtDiv.SetFocus
            End If
        End With
    
    Set rs1 = Nothing
End Sub

Private Sub txtFrom_GotFocus()
    Rangeto = False
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub txtMsg_KeyPress(KeyAscii As Integer)
If Len(Trim(txtMsg.Text)) > 248 Then
      Beep
      MsgBox "Can't enter more than 250 characters", vbExclamation
      KeyAscii = 8
  End If
  
  Select Case KeyAscii
    Case Asc("A") To Asc("Z")
    Case Asc("a") To Asc("z")
    Case Asc("0") To Asc("9")
    Case Asc(".")
    Case Asc(" ")
    Case Asc("/")
    Case Asc("(")
    Case Asc(")")
    Case Asc("-")
    Case Is = 8
    
    Case Else
    Beep
    KeyAscii = 0
  End Select
End Sub

Private Sub txtPayAmt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
        Case Asc(".")
        Case Asc("-")
        Case Is = 8
        Case Else
            Beep
            KeyAscii = 0
    End Select
    
End Sub

Private Sub txtPayAmt_LostFocus()
    txtPayAmt.Text = Format(txtPayAmt.Text, Nfmt)
End Sub

Private Sub txtTo_Change()
With rsGlob
    If .RecordCount > 0 Then
        .MoveFirst
        .Find "EmpCode='" & txtFrom.Text & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Find "Empcode='" & txtTo.Text & "'", 1, adSearchBackward, adBookmarkCurrent
            If .BOF = False Then
                MsgBox "Invalid Range.", vbExclamation
                txtTo.Text = ""
                txtTo.SetFocus
            End If
        Else
            MsgBox "starting records not found"
            txtFrom.Text = ""
            txtFrom.SetFocus
        End If
    End If
End With
End Sub

Private Sub txtTo_GotFocus()
    Rangeto = True
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtTrans_LostFocus()
If txtTrans.Text = "" Then Exit Sub
Set rs1 = CConnect.GetRecordSet("SELECT Description FROM UserTrans WHERE TransCode='" & txtTrans.Text & "'")
    With rs1
        If .RecordCount > 0 Then
            lblTName.Caption = !Description & ""
        Else
            Call cmdTrans_Click
            txtTrans.SetFocus
        End If
    End With

Set rs1 = Nothing
End Sub

Private Sub SecureActivities()
Dim myDone As Boolean
Dim myForm As String
Set rsMySec = CConnect.GetRecordSet("SELECT * FROM GroupModules WHERE GNo='" & CGroup & "' ")
    With rsMySec
        If .RecordCount > 0 Then
            .MoveFirst
                .Find "Name='frmEmpDiv'", , adSearchForward, adBookmarkFirst
                If .EOF = False Then
                    If !MyRight <> "Modify" Then optDiv.Enabled = False
                End If
                
                .Find "Name='frmPostbyTrans'", , adSearchForward, adBookmarkFirst
                If .EOF = False Then
                    If !MyRight <> "Modify" Then optTrans.Enabled = False
                End If
                
                .Find "Name='frmEmpBanks'", , adSearchForward, adBookmarkFirst
                If .EOF = False Then
                    If !MyRight <> "Modify" Then optBank.Enabled = False
                End If
                
                .Find "Name='frmEmployee'", , adSearchForward, adBookmarkFirst
                If .EOF = False Then
                    If !MyRight <> "Modify" Then
                        optPay.Enabled = False
                        optPayRate.Enabled = False
                        optCompany.Enabled = False
                        optCurrency.Enabled = False
                    End If
                End If
                
                .Find "Name='frmPayslip'", , adSearchForward, adBookmarkFirst
                If .EOF = False Then
                    If !MyRight <> "Modify" Then
                        optRecalc.Enabled = False
                        optMsg.Enabled = False
                    End If
                End If
        End If
    End With

Set rsMySec = Nothing
End Sub


Public Sub myStructure()
Dim mm, TheNode As String
trwStruc.Nodes.Clear
On Error GoTo Hell
Set rs5 = CConnect.GetRecordSet("SELECT * FROM STypes WHERE SMain = 1")

With rs5
    If .RecordCount > 0 Then
        .MoveFirst
                    
        
        Set MyNodes = trwStruc.Nodes.Add(, , "O", rs5!Description & "")
        
        Set rs = CConnect.GetRecordSet("SELECT * FROM CStructure WHERE SCode = '" & rs5!code & "' ORDER BY MyLevel, Code")
        
        With rs
            If .RecordCount > 0 Then
                .MoveFirst
                CNode = !LCode & ""
                
                Do While Not .EOF
                    If !MyLevel = 0 Then
                        Set MyNodes = trwStruc.Nodes.Add("O", tvwChild, !LCode, !Description & "")
                        TheNode = Trim(!LCode & "")
                        MyNodes.EnsureVisible
                    Else
                        Set MyNodes = trwStruc.Nodes.Add(!PCode & "", tvwChild, Trim(!LCode & ""), !Description & "")
                        'Set MyNodes = trwStruc.Nodes.Add(TheNode, tvwChild, Trim(!LCode & ""), !Description & "")
                        MyNodes.EnsureVisible
                    End If
                    
                    .MoveNext
                Loop
                .MoveFirst
                
                
            End If
        End With
    
    End If
End With
Exit Sub
Hell:
End Sub

Public Sub SemiStructure()
trwStruc.Nodes.Clear
trwStruc2.Nodes.Clear
Dim myMain As Integer

Set rs1 = CConnect.GetRecordSet("SELECT MyLevel FROM CStructure WHERE LCode='" & LCode & "'")
    With rs1
        If .RecordCount > 0 Then
            .MoveFirst
            If Not IsNull(!MyLevel) Then myMain = !MyLevel
        End If
    End With
    
Set rs1 = Nothing

Set rs = CConnect.GetRecordSet("SELECT * FROM CStructure WHERE LCode Like '" & myLevelCode & "%' ORDER BY myLevel")

With rs
    If .RecordCount > 0 Then
        .MoveFirst

        Do While Not .EOF
            If !MyLevel = myMain Then
                Set MyNodes = trwStruc.Nodes.Add(, , !LCode & "", !Description & "")
                MyNodes.EnsureVisible
                Set MyNodes = trwStruc2.Nodes.Add(, , !LCode & "", !Description & "")
                MyNodes.EnsureVisible
            Else
                Set MyNodes = trwStruc.Nodes.Add(Mid(!LCode, 1, (Len(!LCode) - Len(!code))), tvwChild, !LCode & "", !Description & "")
                MyNodes.EnsureVisible
                Set MyNodes = trwStruc2.Nodes.Add(Mid(!LCode, 1, (Len(!LCode) - Len(!code))), tvwChild, !LCode & "", !Description & "")
                MyNodes.EnsureVisible
            End If

            .MoveNext
        Loop
        .MoveFirst

    End If
End With
End Sub
Private Sub LoadDiv()
    Call myStructure
End Sub


