VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmUGroups 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Groups Rights"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10140
   Icon            =   "frmUGroups.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraEditReportRights 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2475
      Left            =   11280
      TabIndex        =   24
      Top             =   6360
      Visible         =   0   'False
      Width           =   4575
      Begin VB.OptionButton optNoneRptRight 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3375
         TabIndex        =   30
         Top             =   1320
         Width           =   765
      End
      Begin VB.OptionButton optViewRptRight 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1740
         TabIndex        =   29
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton optModifyRptRight 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Modify"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   135
         TabIndex        =   28
         Top             =   1320
         Width           =   1365
      End
      Begin VB.TextBox txtReport 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   570
         Width           =   4305
      End
      Begin VB.CommandButton cmdSaveReportRights 
         Default         =   -1  'True
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
         Left            =   3345
         Picture         =   "frmUGroups.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Save Record"
         Top             =   1860
         Width           =   495
      End
      Begin VB.CommandButton cmdCancelReportRights 
         Cancel          =   -1  'True
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
         Left            =   3825
         Picture         =   "frmUGroups.frx":0F44
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Cancel Process"
         Top             =   1860
         Width           =   495
      End
      Begin VB.Label lblModule 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Module"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   31
         Top             =   315
         Width           =   1785
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10050
      Begin VB.Frame fraEditRights 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2475
         Left            =   4440
         TabIndex        =   32
         Top             =   2160
         Visible         =   0   'False
         Width           =   4575
         Begin VB.CommandButton cmdCanc2 
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
            Left            =   3945
            Picture         =   "frmUGroups.frx":1046
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Cancel Process"
            Top             =   1860
            Width           =   495
         End
         Begin VB.CommandButton cmdSave2 
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
            Left            =   3465
            Picture         =   "frmUGroups.frx":1148
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Save Record"
            Top             =   1860
            Width           =   495
         End
         Begin VB.TextBox txtModule 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   135
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   570
            Width           =   4065
         End
         Begin VB.OptionButton optModify 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "Modify"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   135
            TabIndex        =   35
            Top             =   1200
            Width           =   1365
         End
         Begin VB.OptionButton optView 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "View"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   1620
            TabIndex        =   34
            Top             =   1200
            Width           =   1215
         End
         Begin VB.OptionButton optNone 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "None"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   3615
            TabIndex        =   33
            Top             =   1200
            Width           =   765
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Module"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   135
            TabIndex        =   39
            Top             =   315
            Width           =   1785
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   6480
         TabIndex        =   21
         Top             =   0
         Width           =   1035
         Begin VB.Label lblReports 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Caption         =   "REPORTS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   15
            TabIndex        =   22
            Top             =   150
            Width           =   990
         End
      End
      Begin VB.Frame fraReportRights 
         BackColor       =   &H00C0E0FF&
         Height          =   6735
         Left            =   3600
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   6495
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000007&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   23
            Text            =   "To View or Change User Right on report, double click on the report item"
            Top             =   6360
            Width           =   6375
         End
         Begin MSComctlLib.ImageList TreeImg 
            Left            =   840
            Top             =   5760
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   16
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmUGroups.frx":124A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmUGroups.frx":1F24
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmUGroups.frx":2BFE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmUGroups.frx":38D8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmUGroups.frx":45B2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmUGroups.frx":528C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmUGroups.frx":5F66
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmUGroups.frx":6C40
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmUGroups.frx":791A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmUGroups.frx":85F4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmUGroups.frx":92CE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmUGroups.frx":9FA8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmUGroups.frx":A2C2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmUGroups.frx":AF9C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmUGroups.frx":BC76
                  Key             =   ""
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmUGroups.frx":C950
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.TreeView tvwReports 
            Height          =   6135
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   6300
            _ExtentX        =   11113
            _ExtentY        =   10821
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   706
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            FullRowSelect   =   -1  'True
            HotTracking     =   -1  'True
            ImageList       =   "TreeImg"
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
      Begin MSComctlLib.ImageList imgTree 
         Left            =   2640
         Top             =   3240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   25
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":CC6A
               Key             =   "A"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":D0BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":D3D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":D828
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":DC7A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":E0CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":E3E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":E700
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":EB52
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":EE6C
               Key             =   "B"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":F2BE
               Key             =   "C"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":F710
               Key             =   "D"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":FB62
               Key             =   "E"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":FFB4
               Key             =   "F"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":10406
               Key             =   "G"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":10858
               Key             =   "H"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":10CAA
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":110FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":1154E
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":119A0
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":11DF2
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":12244
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":12696
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":12AE8
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":12F3A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvwGroups 
         Height          =   7035
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   12409
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   4455
         TabIndex        =   5
         Top             =   0
         Width           =   1035
         Begin VB.Label lblUtilities 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Caption         =   "UTILITIES"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   15
            TabIndex        =   12
            Top             =   150
            Width           =   990
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   5475
         TabIndex        =   4
         Top             =   0
         Width           =   1035
         Begin VB.Label lblSetup 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Caption         =   "SET-UP"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   45
            TabIndex        =   13
            Top             =   150
            Width           =   945
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   3435
         TabIndex        =   6
         Top             =   0
         Width           =   1035
         Begin VB.Label lblEmployee 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "EMPLOYEE"
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
            Height          =   210
            Left            =   60
            TabIndex        =   11
            Top             =   150
            Width           =   900
         End
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
         Height          =   375
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSComDlg.CommonDialog Cdl 
         Left            =   3000
         Top             =   3240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame fraSetup 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   6855
         Left            =   3480
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   6495
         Begin MSComctlLib.ListView lvwSetUp 
            Height          =   6450
            Left            =   120
            TabIndex        =   10
            Top             =   345
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   11377
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483642
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
         Begin VB.Label Label2 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Set-Up"
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
            Left            =   165
            TabIndex        =   16
            Top             =   135
            Width           =   1050
         End
      End
      Begin VB.Frame fraUtilities 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   6855
         Left            =   3480
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   6495
         Begin MSComctlLib.ListView lvwUtilities 
            Height          =   6450
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   11377
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
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
         Begin VB.Label Label1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Utilities"
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
            Left            =   150
            TabIndex        =   15
            Top             =   120
            Width           =   1050
         End
      End
      Begin VB.Frame fraEmployee 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   6855
         Left            =   3480
         TabIndex        =   2
         Top             =   360
         Width           =   6495
         Begin MSComctlLib.ListView lvwEmployees 
            Height          =   6450
            Left            =   120
            TabIndex        =   3
            Top             =   345
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   11377
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
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
         Begin VB.Label Label4 
            BackColor       =   &H00C0E0FF&
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
            Height          =   195
            Left            =   165
            TabIndex        =   17
            Top             =   135
            Width           =   1050
         End
      End
      Begin MSComctlLib.ImageList imgEmpTool 
         Left            =   3000
         Top             =   4320
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
               Picture         =   "frmUGroups.frx":1338C
               Key             =   "Search"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":1349E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":135B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUGroups.frx":136C2
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   30
         Left            =   2670
         TabIndex        =   18
         Top             =   2520
         Width           =   30
      End
   End
End
Attribute VB_Name = "frmUGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'For Rpt Rights
Public MyNodes As Node
Dim CNode As String
Dim rsMStruc As Recordset
Dim myReport As String
'For Rpt Rights

Dim rs As Recordset
Dim RsT As Recordset
Dim SelectedGroup As String
Dim SelectedModule As String
Dim SelectedFamily As String

'added by juma to support modifications

Public Sub InitReports_Modif()
Dim mm As String
Dim CNode As String
tvwReports.Nodes.Clear

Set MyNodes = tvwReports.Nodes.Add(, , "Q", "Reports", 14)
Set rsMStruc = CConnect.GetRecordSet("SELECT * FROM tblModules WHERE subsystem = '" & SubSystem & "' and family='REPORTS'")

With rsMStruc
    If .RecordCount > 0 Then
        .MoveFirst
        'CNode = !linkcode & ""
        Do While Not .EOF
            Set MyNodes = tvwReports.Nodes.Add("Q", tvwChild, , !Description & "", 14)
            MyNodes.Tag = Trim(!Name & "")
            MyNodes.EnsureVisible
            
            .MoveNext
        Loop
               
        .MoveFirst
    End If
End With

tvwReports.Refresh
End Sub
'end of addition

'++++++++++For Rpt Rights 01.08.05+++++++++++++++

Public Sub InitReports()
Dim mm As String
Dim CNode As String
tvwReports.Nodes.Clear

Set MyNodes = tvwReports.Nodes.Add(, , "Q", "Reports", 14)
Set rsMStruc = CConnect.GetRecordSet("SELECT * FROM SReports WHERE subsystem = '" & SubSystem & "' ORDER BY MyLevel, Code")

With rsMStruc
    If .RecordCount > 0 Then
        .MoveFirst
        CNode = !LinkCode & ""
        Do While Not .EOF
            If !MyLevel = 0 Then
                If !ObjectID = "None" Then
                    Set MyNodes = tvwReports.Nodes.Add(, , !LinkCode, !Description & "", 14)
                Else
                    Set MyNodes = tvwReports.Nodes.Add(, , !LinkCode, !Description & "", 1)
                End If
                MyNodes.EnsureVisible
            Else
                If !ObjectID = "None" Then
                    Set MyNodes = tvwReports.Nodes.Add(!PreviousCode & "", tvwChild, !LinkCode & "", !Description & "", 14)
                Else
                    Set MyNodes = tvwReports.Nodes.Add(!PreviousCode & "", tvwChild, !LinkCode & "", !Description & "", 1)
                End If
                    
                MyNodes.EnsureVisible
            End If
            
            .MoveNext
        Loop
               
        .MoveFirst
    End If
End With

tvwReports.Refresh

End Sub

Private Sub lblReports_Click()

SelectedFamily = "REPORTS"

If SelectedGroup <> "" Then

    fraReportRights.Visible = True
    fraEmployee.Visible = False
    fraUtilities.Visible = False
    fraSetup.Visible = False
    lblEmployee.ForeColor = vbBlue
    lblUtilities.ForeColor = vbBlack
    lblSetup.ForeColor = vbBlack

End If
End Sub

Private Sub lvwEmployees_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwEmployees
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvwGroups_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwGroups
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvwSetUp_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwSetUp
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvwUtilities_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwUtilities
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub tvwReports_DblClick()

'If Get_System_Report_Field("ObjectID") = "None" Then Exit Sub
'
'GroupRight = Get_Group_right(CGroup, CNode)
'SelectedModule = CNode
'
'    Select Case GroupRight
'        Case "VIEW"
'            optView.Value = True
'            optModify.Enabled = False
'            optNone.Value = False
'            txtModule = tvwReports.SelectedItem
'
''        Case "MODIFY"
''            optView.Value = False
''            optModify.Value = True
''            optNone.Value = False
'
'        Case "NONE"
'            optView.Value = False
'            optModify.Enabled = False
'            optNone.Value = True
'            txtModule = tvwReports.SelectedItem
'    End Select
'
'    fraEditRights.Visible = True

'added by Juma
GroupRight = Get_Group_right(CGroup, tvwReports.SelectedItem.Tag)
SelectedModule = tvwReports.SelectedItem.Tag
    
    Select Case GroupRight
        Case "VIEW"
            optView.Value = True
            optModify.Enabled = False
            optNone.Value = False
            txtModule = tvwReports.SelectedItem
            
        Case "NONE"
            optView.Value = False
            optModify.Enabled = False
            optNone.Value = True
            txtModule = tvwReports.SelectedItem
    End Select

    fraEditRights.Visible = True
'end of addition
End Sub

Function Get_System_Report_Field(schField As String) As String
On Error GoTo Hell

Set rs = CConnect.GetRecordSet("SELECT * FROM SReports where LinkCode='" & CNode & "'")
    Get_System_Report_Field = rs(schField)
Set rs = Nothing

Exit Function
Hell: MsgBox Err.Description, vbCritical, "Get_System_Report_Field"
End Function

Private Sub tvwReports_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Hell
If Button = vbRightButton Then
    'Me.PopupMenu mnuReport
End If
Exit Sub
Hell:
    MsgBox Err.Description, vbExclamation, "User Groups"
End Sub

Private Sub tvwReports_NodeClick(ByVal Node As MSComctlLib.Node)
    If tvwReports.Nodes.Count > 0 Then
        CNode = tvwReports.SelectedItem.Key
    End If
End Sub

'++++++++++For Rpt Rights 01.08.05+++++++++++++++


Private Sub cmdCanc2_Click()
    txtModule.Text = ""
    fraEditRights.Visible = False
End Sub

Private Sub cmdSave2_Click()

If optModify.Value = True Then GroupRight = "MODIFY"
If optView.Value = True Then GroupRight = "VIEW"
If optNone.Value = True Then GroupRight = "NONE"

UpDate_Group_Rights SelectedGroup, SelectedModule

MsgBox "Group Right Assigned Successifuly", vbInformation, "Group Right Assignement"

fraEditRights.Visible = False

Load_Group_Rights SelectedGroup, SelectedFamily

End Sub

Function UpDate_Group_Rights(sGroupID As String, sModuleID As String)
On Error GoTo Hell

Set rs = CConnect.GetRecordSet("Select * From tblAssignedRights where subsystem = '" & SubSystem & "' and GROUP_ID= ('" & sGroupID & "') and MODULE_ID = ('" & sModuleID & "')")

With rs
    If .EOF = False Then
    
        strQ = "Update tblAssignedRights set Assigned_Rights='" & GroupRight & "' where  subsystem = '" & SubSystem & "' and GROUP_ID= ('" & sGroupID & "') and MODULE_ID = ('" & sModuleID & "') "
        Action = "ADDED ACCESS RIGHTS; RIGHT: " & GroupRight & "; MODULE: " & !MODULE_NAME & "; GROUP ID: " & sGroupID
        CConnect.ExecuteSql strQ
        
    Else
    
        .AddNew
            !GROUP_ID = sGroupID
            !MODULE_ID = sModuleID
            !MODULE_NAME = txtModule
            !ASSIGNED_RIGHTS = GroupRight
            !DoneBy = CurrentUser
            !DateDone = Date
            !SubSystem = "HRBase"
        .Update
    
    End If
End With

Set rs = Nothing
   
Exit Function
Hell: MsgBox Err.Description, vbCritical, "UpDate Group Rights"
End Function



Private Sub Form_Load()
Decla.Security Me

With frmMain2
    Me.Move .tvwMain.Width + .tvwMain.Width / 32.75, (.Height / 5.52) ''- 155
End With

CConnect.CColor Me, MyColor

Make_Group_Columns

Load_User_Groups

'Call InitReports
Call InitReports_Modif

End Sub

Function Load_User_Groups()
Set rs = CConnect.GetRecordSet("Select * From tblUserGroup WHERE subsystem = '" & SubSystem & "'")
    With rs
        Do While Not .EOF
            Set LI = lvwGroups.ListItems.Add(, , !GROUP_CODE & "")
                LI.ListSubItems.Add , , !GROUP_NAME & ""
            .MoveNext
        Loop
    End With
    rs.Close
Set rs = Nothing
End Function

Private Sub lblEmployee_Click()

SelectedFamily = "EMPLOYEE"

If SelectedGroup <> "" Then

    Load_Group_Rights SelectedGroup, "EMPLOYEE"

    fraEmployee.Visible = True
    fraReportRights.Visible = False
    fraUtilities.Visible = False
    fraSetup.Visible = False
    lblEmployee.ForeColor = vbBlue
    lblUtilities.ForeColor = vbBlack
    lblSetup.ForeColor = vbBlack

End If
    
End Sub

Private Sub lblSetup_Click()
   
SelectedFamily = "SET-UP"
   
If SelectedGroup <> "" Then

    Load_Group_Rights SelectedGroup, "SET-UP"
    
    fraEmployee.Visible = False
    fraReportRights.Visible = False
    fraUtilities.Visible = False
    fraSetup.Visible = True
    lblEmployee.ForeColor = vbBlack
    lblUtilities.ForeColor = vbBlack
    lblSetup.ForeColor = vbBlue
    
End If

End Sub

Private Sub lblUtilities_Click()
   
SelectedFamily = "UTILITIES"
   
If SelectedGroup <> "" Then

    Load_Group_Rights SelectedGroup, "UTILITIES"
    
    fraEmployee.Visible = False
    fraReportRights.Visible = False
    fraUtilities.Visible = True
    fraSetup.Visible = False
    lblEmployee.ForeColor = vbBlack
    lblUtilities.ForeColor = vbBlue
    lblSetup.ForeColor = vbBlack
    
End If
    
End Sub

Private Sub Make_Group_Columns()
With lvwGroups
    .ColumnHeaders.Clear
    .ListItems.Clear
    .ColumnHeaders.Add , , "Group Code", 1000
    .ColumnHeaders.Add , , "Name", 2000
    .View = lvwReport
End With
End Sub

Private Sub Make_Rights_Columns()
    
    With lvwEmployees
        .ColumnHeaders.Clear
        .ListItems.Clear
        .ColumnHeaders.Add , , "N", 0
        .ColumnHeaders.Add , , "Description", 3000
        .ColumnHeaders.Add , , "Rights", , vbCenter
        .View = lvwReport
    End With
    
    With lvwUtilities
        .ColumnHeaders.Clear
        .ListItems.Clear
        .ColumnHeaders.Add , , "N", 0
        .ColumnHeaders.Add , , "Description", 3000
        .ColumnHeaders.Add , , "Rights", , vbCenter
        .View = lvwReport
    End With
    
    With lvwSetUp
        .ColumnHeaders.Clear
        .ListItems.Clear
        .ColumnHeaders.Add , , "N", 0
        .ColumnHeaders.Add , , "Description", 3000
        .ColumnHeaders.Add , , "Rights", , vbCenter
        .View = lvwReport
    End With
    
End Sub

Function Load_Group_Rights(sGroupID As String, sFamily As String)
On Error GoTo Hell

Make_Rights_Columns

Set rs = CConnect.GetRecordSet("Select * From tblMODULES where (FAMILY = '" & sFamily & "') and subsystem = '" & SubSystem & "'")

Select Case sFamily

    Case "EMPLOYEE"
        
        With rs
            Do While Not .EOF
                Set LI = lvwEmployees.ListItems.Add(, , !Name & "")
                    LI.ListSubItems.Add , , !Description & ""
                    GroupRight = Get_Group_right(sGroupID, !Name)
                    LI.ListSubItems.Add , , GroupRight & ""
                .MoveNext
            Loop
        End With
        rs.Close
        Set rs = Nothing
        
    Case "UTILITIES"
    
        With rs
            Do While Not .EOF
                Set LI = lvwUtilities.ListItems.Add(, , !Name & "")
                    LI.ListSubItems.Add , , !Description & ""
                    GroupRight = Get_Group_right(sGroupID, !Name)
                    LI.ListSubItems.Add , , GroupRight & ""
                .MoveNext
            Loop
        End With
        rs.Close
        Set rs = Nothing
    
    Case "SET-UP"
    
        With rs
            Do While Not .EOF
                Set LI = lvwSetUp.ListItems.Add(, , !Name & "")
                    LI.ListSubItems.Add , , !Description & ""
                    GroupRight = Get_Group_right(sGroupID, !Name)
                    LI.ListSubItems.Add , , GroupRight & ""
                .MoveNext
            Loop
        End With
        rs.Close
        Set rs = Nothing


End Select

Exit Function
Hell: MsgBox Err.Description, vbCritical, "Load Group Rights"
End Function

Private Sub lvwEmployees_DblClick()
If lvwEmployees.ListItems.Count > 0 Then
    SelectedModule = lvwEmployees.SelectedItem
    txtModule = lvwEmployees.SelectedItem.ListSubItems(1)
    
    Select Case lvwEmployees.SelectedItem.ListSubItems(2)
        Case "VIEW"
            optView.Value = True
            optModify.Value = False
            optNone.Value = False
            
        Case "MODIFY"
            optView.Value = False
            optModify.Value = True
            optNone.Value = False
        
        Case "NONE"
            optView.Value = False
            optModify.Value = False
            optNone.Value = True
    End Select

    fraEditRights.Visible = True
    
End If
End Sub

Private Sub lvwGroups_Click()
    On Error GoTo errHandler
    SelectedGroup = lvwGroups.SelectedItem
    Exit Sub
errHandler:
End Sub

Private Sub lvwSetUp_DblClick()
If lvwSetUp.ListItems.Count > 0 Then
    SelectedModule = lvwSetUp.SelectedItem
    txtModule = lvwSetUp.SelectedItem.ListSubItems(1)
    
    Select Case lvwSetUp.SelectedItem.ListSubItems(2)
        Case "VIEW"
            optView.Value = True
            optModify.Value = False
            optNone.Value = False
            
        Case "MODIFY"
            optView.Value = False
            optModify.Value = True
            optNone.Value = False
        
        Case "NONE"
            optView.Value = False
            optModify.Value = False
            optNone.Value = True
    End Select

    fraEditRights.Visible = True
    
End If
End Sub

Private Sub lvwUtilities_DblClick()
If lvwUtilities.ListItems.Count > 0 Then
    SelectedModule = lvwUtilities.SelectedItem
    txtModule = lvwUtilities.SelectedItem.ListSubItems(1)
    
    Select Case lvwUtilities.SelectedItem.ListSubItems(2)
        Case "VIEW"
            optView.Value = True
            optModify.Value = False
            optNone.Value = False
            
        Case "MODIFY"
            optView.Value = False
            optModify.Value = True
            optNone.Value = False
        
        Case "NONE"
            optView.Value = False
            optModify.Value = False
            optNone.Value = True
    End Select

    fraEditRights.Visible = True
    
End If
End Sub

