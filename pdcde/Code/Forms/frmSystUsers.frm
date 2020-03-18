VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmSystUsers 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   Caption         =   "System Users"
   ClientHeight    =   7335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9975
   Icon            =   "frmSystUsers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraStruc 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Employees"
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
      Height          =   6000
      Left            =   3750
      TabIndex        =   33
      Top             =   345
      Visible         =   0   'False
      Width           =   4800
      Begin VB.CommandButton cmdSCancel 
         Caption         =   "CANCEL"
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
         Left            =   3345
         TabIndex        =   36
         Top             =   5580
         Width           =   1335
      End
      Begin VB.CommandButton cmdSSelect 
         Caption         =   "SELECT"
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
         Left            =   3345
         TabIndex        =   34
         Top             =   5280
         Width           =   1335
      End
      Begin MSComctlLib.TreeView trwStruc 
         Height          =   4860
         Left            =   135
         TabIndex        =   35
         Top             =   345
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   8573
         _Version        =   393217
         Indentation     =   706
         LabelEdit       =   1
         LineStyle       =   1
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
   Begin VB.Frame fraEmpList 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Employees"
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
      Height          =   6000
      Left            =   3765
      TabIndex        =   20
      Top             =   345
      Visible         =   0   'False
      Width           =   4800
      Begin VB.CommandButton cmdECancel 
         Caption         =   "CANCEL"
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
         Left            =   3345
         TabIndex        =   23
         Top             =   5580
         Width           =   1335
      End
      Begin VB.CommandButton cmdESelect 
         Caption         =   "SELECT"
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
         Left            =   3345
         TabIndex        =   21
         Top             =   5280
         Width           =   1335
      End
      Begin MSComctlLib.ListView lvwEmp 
         Height          =   4770
         Left            =   120
         TabIndex        =   22
         Top             =   405
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   8414
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
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
      Begin VB.Label lblTitle2 
         BackColor       =   &H00800000&
         Caption         =   " Employees List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   7380
      End
   End
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Users"
      ForeColor       =   &H80000008&
      Height          =   3645
      Left            =   1275
      TabIndex        =   5
      Top             =   1350
      Visible         =   0   'False
      Width           =   7305
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
         ItemData        =   "frmSystUsers.frx":0442
         Left            =   3270
         List            =   "frmSystUsers.frx":045B
         TabIndex        =   31
         Top             =   1800
         Width           =   2610
      End
      Begin VB.CommandButton cmdStruc 
         Height          =   300
         Left            =   2790
         Picture         =   "frmSystUsers.frx":04A1
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1800
         Width           =   300
      End
      Begin VB.TextBox txtDAccess 
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
         Left            =   135
         TabIndex        =   28
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox txtDesc 
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
         Left            =   1980
         TabIndex        =   26
         Top             =   615
         Width           =   5145
      End
      Begin VB.CommandButton cmdFrom 
         Height          =   300
         Left            =   1560
         Picture         =   "frmSystUsers.frx":05A3
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   615
         Width           =   300
      End
      Begin VB.CheckBox chkExcc 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Exclusive rights"
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
         Height          =   375
         Left            =   4200
         TabIndex        =   12
         Top             =   2295
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.TextBox txtUID 
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
         TabIndex        =   0
         Top             =   615
         Width           =   1440
      End
      Begin VB.ComboBox cboGNo 
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
         Left            =   135
         TabIndex        =   1
         Top             =   1215
         Width           =   1365
      End
      Begin VB.TextBox txtPass 
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
         IMEMode         =   3  'DISABLE
         Left            =   150
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   2385
         Width           =   1665
      End
      Begin VB.TextBox txtCon 
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
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   2385
         Width           =   1815
      End
      Begin VB.ComboBox cboDesc 
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
         Left            =   1590
         TabIndex        =   2
         Top             =   1215
         Width           =   3720
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
         Left            =   6630
         Picture         =   "frmSystUsers.frx":06A5
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cancel Process"
         Top             =   2955
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
         Left            =   6150
         Picture         =   "frmSystUsers.frx":07A7
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Save Record"
         Top             =   2955
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
         Left            =   1140
         Picture         =   "frmSystUsers.frx":08A9
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Delete Record"
         Top             =   2895
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
         Left            =   660
         Picture         =   "frmSystUsers.frx":09AB
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Edit Record"
         Top             =   2895
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
         Left            =   180
         Picture         =   "frmSystUsers.frx":0AAD
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Add New record"
         Top             =   2895
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "CLOSE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1935
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2910
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSComDlg.CommonDialog Cdl 
         Left            =   3450
         Top             =   2385
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList imgTree 
         Left            =   1950
         Top             =   3885
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
               Picture         =   "frmSystUsers.frx":0BAF
               Key             =   "A"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystUsers.frx":1001
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystUsers.frx":131B
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystUsers.frx":176D
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystUsers.frx":1BBF
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystUsers.frx":2011
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystUsers.frx":232B
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystUsers.frx":2645
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystUsers.frx":2A97
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystUsers.frx":2DB1
               Key             =   "B"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystUsers.frx":3203
               Key             =   "C"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystUsers.frx":3655
               Key             =   "D"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystUsers.frx":3AA7
               Key             =   "E"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystUsers.frx":3EF9
               Key             =   "F"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystUsers.frx":434B
               Key             =   "G"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystUsers.frx":479D
               Key             =   "H"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystUsers.frx":4BEF
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystUsers.frx":5041
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystUsers.frx":5493
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystUsers.frx":58E5
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystUsers.frx":5D37
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystUsers.frx":6189
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystUsers.frx":65DB
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystUsers.frx":6A2D
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystUsers.frx":6E7F
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystUsers.frx":72D1
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Employment Terms Access"
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
         Left            =   3270
         TabIndex        =   32
         Top             =   1575
         Width           =   1890
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Division Access"
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
         Left            =   135
         TabIndex        =   29
         Top             =   1575
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Description"
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
         Left            =   1980
         TabIndex        =   27
         Top             =   390
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Group ID."
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
         Left            =   135
         TabIndex        =   17
         Top             =   975
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "User ID."
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
         Left            =   120
         TabIndex        =   16
         Top             =   390
         Width           =   1335
      End
      Begin VB.Label lblPass 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Password"
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
         Left            =   150
         TabIndex        =   15
         Top             =   2145
         Width           =   1560
      End
      Begin VB.Label lblCon 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Confirm"
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
         Left            =   1920
         TabIndex        =   14
         Top             =   2145
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Description"
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
         Left            =   1590
         TabIndex        =   13
         Top             =   975
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   7020
      Left            =   15
      TabIndex        =   18
      Top             =   -90
      Width           =   9930
      Begin MSComctlLib.ListView lvwUsers 
         Height          =   6930
         Left            =   0
         TabIndex        =   19
         Top             =   90
         Width           =   9930
         _ExtentX        =   17515
         _ExtentY        =   12224
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
   End
End
Attribute VB_Name = "frmSystUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LI As ListItem
Dim Excc As String
Dim EncryptPass As String
Dim DCode As String
Dim MyNodes As Node
Dim CNode As String
Dim PNode As String
Dim oldCode As String


Private Sub cboDesc_Click()
If Not cboDesc.Text = "" Then
    With rs1
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "Description like '" & cboDesc.Text & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                cboGNo.Text = !GNo & ""
            End If
        End If
    End With
End If

End Sub



Private Sub cboDesc_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboGNo_Click()
If Not cboGNo.Text = "" Then
    With rs1
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "GNo like '" & cboGNo.Text & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                cboDesc.Text = !Description & ""
            End If
        End If
    End With
End If

End Sub

Private Sub cboGNo_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Public Sub cmdCancel_Click()
    If PSave = False Then
        If MsgBox("Do you want to save changes?", vbQuestion + vbYesNo) = vbYes Then  '
            Call cmdSave_Click
            Exit Sub
        End If
    End If

txtUID.Locked = True
txtPass.Visible = False
txtCon.Visible = False
lblPass.Visible = False
lblCon.Visible = False
'chkExcc.Visible = False

EnableCmd
cmdCancel.Enabled = False
cmdSave.Enabled = False
SaveNew = False

With frmMain2
    .cmdNew.Enabled = True
    .cmdEdit.Enabled = True
    .cmdDelete.Enabled = True
    .cmdCancel.Enabled = False
    .cmdSave.Enabled = False
End With
    
Call DisplayRecords

'Call LoadList
fraDetails.Visible = False

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub cmdDelete_Click()
Dim resp As String
If lvwUsers.ListItems.Count < 1 Then
    MsgBox "No records.", vbExclamation
    Exit Sub
End If

resp = MsgBox("Are you sure you want to delete  " & lvwUsers.SelectedItem & "   from the records?", vbQuestion + vbYesNo)
If resp = vbNo Then
    Exit Sub
End If
Action = "DELETED A SYSTEM USER; USERNAME: " & lvwUsers.SelectedItem

CConnect.ExecuteSql ("DELETE From Security WHERE UID = '" & lvwUsers.SelectedItem & "' AND subsystem = '" & SubSystem & "'")

'' if AuditTrail = True Then cConnect.ExecuteSql ("INSERT INTO AuditTrail (UserId, DTime, Trans, TDesc, MySection)VALUES('" & CurrentUser & "','" & Date & " " & Time & "','Deleting User','" & lvwUsers.SelectedItem & "','Set-up')")

rs2.Requery

Call DisplayRecords
Call LoadList
    
End Sub

Private Sub cmdECancel_Click()
    fraEmpList.Visible = False
    
End Sub

Public Sub cmdEdit_Click()
If txtUID.Text = "" Then
    MsgBox "No records.", vbExclamation
    Exit Sub
End If


With rs2
    If .RecordCount > 0 Then
        .MoveFirst
        .Find "UID like '" & CurrentUser & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            Excc = rs2!Excc
            Call EEncryptPassword
            Excc = EncryptPass
            
        
        End If
        
    Else
        If ByPass = True Then
            chkExcc.Visible = True
        Else
            chkExcc.Visible = False
        End If
        
    End If
    
End With


With rs2
    If .RecordCount > 0 Then
        .MoveFirst
        .Find "UID like '" & txtUID.Text & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            
            Excc = !Pass & ""
            Call EEncryptPassword
            Excc = EncryptPass
            
            txtPass.Text = Excc
            txtCon.Text = Excc
            
            Excc = rs2!Excc
            Call EEncryptPassword
            Excc = EncryptPass

        
        End If
   
    Else
        If ByPass = True Then
            chkExcc.Visible = True
        Else
            chkExcc.Visible = False
        End If
        
    End If
    
End With

        
Call DisableCmd

txtUID.Locked = True
txtPass.Visible = True
txtCon.Visible = True
lblPass.Visible = True
lblCon.Visible = True

cmdSave.Enabled = True
cmdCancel.Enabled = True
fraDetails.Visible = True

SaveNew = False
cmdFrom.Enabled = False


End Sub




Private Sub cmdESelect_Click()
    If lvwEmp.ListItems.Count > 0 Then
        With rsGlob2
            If .RecordCount > 0 Then
                .MoveFirst
                .Find "employee_id like '" & lvwEmp.SelectedItem.Tag & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    fraEmpList.Visible = False
                    txtUID.Text = !empcode & ""
                    txtDesc.Text = !SurName & "" & " " & !OtherNames & ""
                Else
                    MsgBox "Record not found", vbInformation
                End If
            Else
                MsgBox "Record not found", vbInformation
            End If
        End With
        
    Else
        MsgBox "No records to be selected.", vbInformation
    End If
    
End Sub

Private Sub cmdFrom_Click()
    fraEmpList.Visible = True
End Sub

Public Sub cmdNew_Click()
Call DisableCmd

Call Cleartxt

txtUID.Locked = False
txtPass.Visible = True
txtCon.Visible = True
lblPass.Visible = True
lblCon.Visible = True

cmdSave.Enabled = True
cmdCancel.Enabled = True

With rs2
    If .RecordCount > 0 Then
        .MoveFirst
        .Find "UID like '" & CurrentUser & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            Excc = CConnect.Crypt(rs2!Excc)
            If !Exc = True And Excc = 0 Then
                MsgBox "The system has detected suspicious actions. User: " & rs2!UName & " will therefor not be allowed to use the system again.", vbCritical
                
                CConnect.ExecuteSql ("DELETE From Security WHERE UID = '" & rs!UName & "' AND subsystem = '" & SubSystem & "'")
                rs2.Requery
                
'                            If Not .EOF Then
'                                .Delete
'                                .Update
'                            End If
         
            End If
                    
            If !Exc = True And Excc = 1 Then
                chkExcc.Visible = True
            Else
                chkExcc.Visible = False
               
            End If
        
        End If
  
    End If
    
End With


With rs2
    If .RecordCount < 1 Then
        If ByPass = True Then
            chkExcc.Visible = True
        Else
            chkExcc.Visible = False
        End If
        
    End If
End With


SaveNew = True
fraDetails.Visible = True
txtUID.SetFocus
cmdFrom.Enabled = True

End Sub



Public Sub cmdSave_Click()
Dim EDate As Date
If txtUID.Text = "" Then
    MsgBox "You have to enter the user ID.", vbExclamation
    Exit Sub
End If

If cboGNo.Text = "" Then
    MsgBox "You have to enter the group description.", vbExclamation
    Exit Sub
End If

If Len(Trim(txtPass.Text)) < 4 Then
    MsgBox "You have to enter the password of not less than four characters.", vbExclamation
    txtPass.SetFocus
    Exit Sub
End If

If Not txtPass.Text = txtCon.Text Then
    MsgBox "The passwords do not match. Re-enter the correct password."
    txtPass.SetFocus
    Exit Sub
End If

If chkExcc.Value = 1 Then
    Excc = 1
Else
    Excc = 0
End If

If txtDAccess.Text = "" Then
    DCode = ""
End If

With rs2
    If SaveNew = True Then
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "UID like '" & txtUID.Text & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                MsgBox "User ID already exists. Enter another one.", vbExclamation
                txtUID.Text = ""
                txtUID.SetFocus
                Exit Sub
            End If
        End If
    
    End If
    
    If PromptSave = True Then
        If MsgBox("Are you sure you want to save the record.", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
            
    EDate = DateAdd("d", rsGenOpt!PDays, Date)
    CConnect.ExecuteSql ("DELETE From Security WHERE UID = '" & txtUID.Text & "' AND subsystem = '" & SubSystem & "'")
    Action = "ADDED A SYSTEM USER; USERNAME: " & txtUID & "; DESCRIPTION: " & txtDesc.Text & "; USER GROUP: " & cboGNo.Text
    CConnect.ExecuteSql ("INSERT INTO Security (UID, Description, GNo, Pass, Exc, Excc, EDate, LCode, Terms, DName, subsystem)" & _
                " VALUES('" & txtUID.Text & "','" & txtDesc.Text & "','" & cboGNo.Text & "','" & CConnect.Crypt(txtPass.Text) & "'," & chkExcc.Value & ",'" & CConnect.Crypt(chkExcc.Value) & "','" & EDate & "','" & DCode & "','" & cboTerms.Text & "','" & txtDAccess.Text & "','" & SubSystem & "')")
    
End With
  
   
txtUID.Locked = True
txtPass.Visible = False
txtCon.Visible = False
lblPass.Visible = False
lblCon.Visible = False
'chkExcc.Visible = False

EnableCmd
cmdCancel.Enabled = False
cmdSave.Enabled = False

With frmMain2
    .cmdNew.Enabled = True
    .cmdEdit.Enabled = True
    .cmdDelete.Enabled = True
    .cmdCancel.Enabled = False
    .cmdSave.Enabled = False
End With
    
rs1.Requery
rs2.Requery
'Call DisplayRecords
Call LoadList
SaveNew = False

fraDetails.Visible = False
End Sub

Private Sub Command3_Click()

End Sub

Private Sub cmdSCancel_Click()
    fraStruc.Visible = False
End Sub

Private Sub cmdSSelect_Click()
    If trwStruc.Nodes.Count > 0 Then
        txtDAccess.Text = trwStruc.SelectedItem.Text & ""
        DCode = trwStruc.SelectedItem.Key & ""
        fraStruc.Visible = False
    End If
End Sub

Private Sub cmdStruc_Click()
    fraStruc.Visible = True
End Sub

Private Sub Form_Load()
Decla.Security Me
oSmart.FReset Me
'With frmMain2
'    Me.Move .tvwMain.Width + .tvwMain.Width / 32.75, Screen.Height * MyR
'End With

If oSmart.hRatio > 1.1 Then
    With frmMain2
        Me.Move .tvwMain.Width + .tvwMain.Width / 32.75, (.Height / 5.52) '- 155
    End With
Else
     With frmMain2
        Me.Move .tvwMain.Width + .tvwMain.Width / 32.75, .Height / 5.55
    End With
    
End If

CConnect.CColor Me, MyColor

Dim i As Object

Call InitGrid
'cnnPayroll.Open "Leave"


'Set rs = cConnect.GetSecurity("SELECT * FROM CUser")
Set rs1 = CConnect.GetRecordSet("SELECT * FROM Groups ORDER BY GNo")
Set rs2 = CConnect.GetRecordSet("SELECT * From Security WHERE subsystem = '" & SubSystem & "' ORDER BY GNo")


With rs2
    If .RecordCount > 0 Then
        .MoveFirst
        .Find "UID like '" & CurrentUser & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            Excc = CConnect.Crypt(rs2!Excc)
'                        Excc = rs2!Excc
'                        Call EEncryptPassword
'                        Excc = EncryptPass
            
'            If !Exc = True And Excc = 0 Then
'                'wrong = True
'                MsgBox "The system has detected suspicious actions. User: " & rs2!UName & " will therefor not be allowed to use the system again.", vbCritical
'
'                If Not .EOF Then
'                    .Delete
'                    .Update
'                End If
'
'            End If
                   
'                        Excc = CConnect.Crypt
            
'            If !Exc = True And Excc = 1 Then
'                chkExcc.Visible = True
'            Else
'                chkExcc.Visible = False
'                Set rs2 = Nothing
'
'                Set rs2 = cConnect.GetRecordSet("Select * From Security where Exc = false AND subsystem = '" & subsystem & "' order by UID")
'            End If
        
        End If
        

 
    End If
    
End With


cmdCancel.Enabled = False
cmdSave.Enabled = False

With rs1
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
            cboGNo.AddItem (!GNo & "")
            cboDesc.AddItem (!Description & "")
            
            .MoveNext
        Loop
        .MoveFirst
    End If
End With

Call DisplayRecords

With rsGlob2
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
            Set LI = lvwEmp.ListItems.Add(, , !empcode & "")
            LI.ListSubItems.Add , , !SurName & ""
            LI.ListSubItems.Add , , !OtherNames & ""
            LI.Tag = !employee_id
            
            .MoveNext
        Loop
        .MoveFirst
    End If
End With

 
txtUID.Locked = True
txtPass.Visible = False
txtCon.Visible = False
lblPass.Visible = False
lblCon.Visible = False

For Each i In Me
    If TypeOf i Is ComboBox Then
        i.Locked = True
    End If
Next i

'If CSecurity.SetupModify = False Then
'    Call DisableCmd
'    cmdNext.Enabled = True
'    cmdLast.Enabled = True
'    cmdClose.Enabled = True
'End If


With rs2
    If .RecordCount < 1 Then
        cmdNew.Enabled = True
     
    End If
End With
        
Call LoadList

Call myStructure

End Sub
Private Sub Form_Resize()
    oSmart.FResize Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmMain2.Caption = "Personnel Director " & App.FileDescription
    
    Set rs = Nothing
    Set rs1 = Nothing
    Set rs2 = Nothing

    
End Sub


Private Sub DisplayRecords()
Cleartxt

    With rs2
        If Not .EOF And Not .BOF Then
                        
            txtUID.Text = !UID & ""
            txtDesc.Text = !Description & ""
            cboGNo.Text = !GNo & ""
        
            If Not !GNo = "" Then
                With rs1
                    If .RecordCount > 0 Then
                        .MoveFirst
                        .Find "GNo like '" & rs2!GNo & "'", , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            cboDesc.Text = !Description & ""
                        End If
                    End If
                End With
            End If
            
            frmMain2.txtDetails.Caption = ""
            frmMain2.txtDetails.Caption = "User Name:" & " " & !UID & "" & "     " & "Group Code: " & " " & !GNo & "" & "" & "     " & vbCrLf & _
                        "" & "Description: " & " " & cboDesc.Text & "" & ""
                        
            txtPass.Text = CConnect.Crypt(!Pass & "")
            txtCon.Text = CConnect.Crypt(!Pass & "")
            txtDAccess.Text = !DName & ""
            cboTerms.Text = !terms & ""
            DCode = !LCode & ""
            
            Excc = CConnect.Crypt(rs2!Excc)
'            Excc = rs2!Excc
'            Call EEncryptPassword
'            Excc = EncryptPass
            
'            If !Exc = True And Excc = 0 Then
'                MsgBox "The system has detected suspicious actions. User: " & rs2!UName & " will therefor not be allowed to use the system again.", vbCritical
'
'                If Not .EOF Then
'                    .Delete
'                    .Update
'                End If
'
'            End If
                    
'            If !Exc = True And Excc = 1 Then
'                If chkExcc.Visible = True Then
'                    chkExcc.Value = 1
'                End If
'
'            Else
'                If chkExcc.Visible = True Then
'                    chkExcc.Value = 0
'                End If
'
'            End If
            
         
              
        End If
    End With
    
End Sub




Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Me.MousePointer = vbHourglass

Dim record As String
    Select Case Button.Key
        Case "Find"
          
    End Select
    
Me.MousePointer = 0

End Sub








Private Sub Cleartxt()
Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
            i.Text = ""
        End If
    Next i
    
    For Each i In Me
        If TypeOf i Is OptionButton Then
            i.Value = False
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

Public Sub DisableCmd()
Dim i As Object
    For Each i In Me
        If TypeOf i Is CommandButton Then
            i.Enabled = False
        End If
    Next i
    
    For Each i In Me
        If TypeOf i Is ComboBox Then
            i.Locked = False
        End If
    Next i
    
    cmdESelect.Enabled = True
    cmdECancel.Enabled = True
    cmdFrom.Enabled = True
    cmdStruc.Enabled = True
    cmdSSelect.Enabled = True
    cmdSCancel.Enabled = True
    
End Sub

Public Sub EnableCmd()
Dim i As Object
    For Each i In Me
        If TypeOf i Is CommandButton Then
            i.Enabled = True
        End If
    Next i
    
    For Each i In Me
        If TypeOf i Is ComboBox Then
            i.Locked = True
        End If
    Next i
    
End Sub






Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error GoTo errHandler
Dim ss As String
Dim myfile As String
    Me.MousePointer = vbHourglass
    Set a = New Application
    myfile = App.Path & "\Leave Reports\System Users.rpt"
    Set R = a.OpenReport(myfile)
    
'          If Not frmRange.txtFrom.Text = "" And Not frmRange.txtTo.Text = "" Then
'              If Not frmRange.cboPCode = "" Then
'                  SQL = "{Employee.PayrollNo} in '" & frmRange.txtFrom.Text & "' to '" & frmRange.txtTo.Text & "' and {Payslip.YYear} in " & frmRange.txtFYear.Text & " to " & frmRange.txtTYear.Text & " and {Payslip.MMonth} in " & FMonth & " to " & TMonth & " and {EmpPayPoint.PPCode} = '" & frmRange.cboPCode.Text & "'"
'              Else
'                  SQL = "{Employee.PayrollNo} in '" & frmRange.txtFrom.Text & "' to '" & frmRange.txtTo.Text & "' and {Payslip.YYear} in " & frmRange.txtFYear.Text & " to " & frmRange.txtTYear.Text & " and {Payslip.MMonth} in " & FMonth & " to " & TMonth & ""
'              End If
'
'              R.RecordSelectionFormula = SQL
'          Else
'              If Not frmRange.cboPCode = "" Then
'                  SQL = "{Payslip.YYear} in " & frmRange.txtFYear.Text & " to " & frmRange.txtTYear.Text & " and {Payslip.MMonth} in " & FMonth & " to " & TMonth & " and {EmpPayPoint.PPCode} = '" & frmRange.cboPCode.Text & "'"
'              Else
'                  SQL = "{Payslip.YYear} in " & frmRange.txtFYear.Text & " to " & frmRange.txtTYear.Text & " and {Payslip.MMonth} in " & FMonth & " to " & TMonth & ""
'              End If
'
'              R.RecordSelectionFormula = SQL
'          End If

  'SQL = "{Payslip.YYear} in " & frmRange.txtFYear.Text & " to " & frmRange.txtTYear.Text & " and {Payslip.MMonth} in " & FMonth & " to " & TMonth & " and {EmpPayPoint.PPCode} = '" & frmRange.cboPCode.Text & "'"
  'R.RecordSelectionFormula = MySql
  
  R.ReadRecords

  With frmReports.CRViewer1
      .ReportSource = R
      .ViewReport
  End With

  frmReports.Show vbModal
    Me.MousePointer = 0
Exit Sub

errHandler:
If Err.Description = "File not found." Then
    Cdl.DialogTitle = "Select the report to show"
    Cdl.InitDir = App.Path & "/Leave Reports"
    Cdl.Filter = "Reports {* .rpt|* .rpt"
    Cdl.ShowOpen
    myfile = Cdl.FileName
    If Not myfile = "" Then
        Resume
    Else
        Me.MousePointer = 0
    End If
Else
    MsgBox Err.Description, vbInformation
    Me.MousePointer = 0
End If
End Sub





Private Sub lvwEmp_DblClick()
    Call cmdESelect_Click
End Sub

Private Sub lvwUsers_DblClick()
If lvwUsers.ListItems.Count > 0 Then
    txtUID.Text = lvwUsers.SelectedItem & ""
    If frmMain2.cmdEdit.Enabled = True Then
        Call frmMain2.cmdEdit_Click
    End If
End If

End Sub

Private Sub lvwUsers_ItemClick(ByVal Item As MSComctlLib.ListItem)
If lvwUsers.ListItems.Count > 0 Then
    With rs2
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "UID like '" & lvwUsers.SelectedItem & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                Call DisplayRecords
            End If
        End If
    End With
End If
            
            
End Sub


Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub trwStruc_DblClick()
    Call cmdSSelect_Click
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
If Len(Trim(txtDesc.Text)) > 200 Then
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

Private Sub txtUID_Change()
    txtUID.Text = UCase(txtUID.Text)
    txtUID.SelStart = Len(txtUID.Text)
End Sub

Private Sub txtUID_KeyPress(KeyAscii As Integer)
    If Len(Trim(txtUID.Text)) > 10 Then
        Beep
        MsgBox "Can't enter more than 10 characters", vbExclamation
        KeyAscii = 8
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

Function EEncryptPassword()
Dim Pwd As Variant
Dim Temp As String, PwdChr As Long
Dim EncryptKey As Long
Pwd = Excc
EncryptKey = Int(Sqr(Len(Pwd) * 81)) + 23

For PwdChr = 1 To Len(Pwd)
    Temp = Temp + Chr(Asc(Mid(Pwd, PwdChr, 1)) Xor EncryptKey)
Next PwdChr

EncryptPass = Temp

End Function

Function EncryptPassword()
Dim Pwd As Variant
Dim Temp As String, PwdChr As Long
Dim EncryptKey As Long
Pwd = txtPass.Text
EncryptKey = Int(Sqr(Len(Pwd) * 81)) + 23

For PwdChr = 1 To Len(Pwd)
    Temp = Temp + Chr(Asc(Mid(Pwd, PwdChr, 1)) Xor EncryptKey)
Next PwdChr

EncryptPass = Temp

End Function



Private Sub InitGrid()
With lvwUsers
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "User ID"
    .ColumnHeaders.Add , , "User Description", 2000
    .ColumnHeaders.Add , , "Group No"
    .ColumnHeaders.Add , , "Group Description", 3500
    .ColumnHeaders.Add , , "Division Access", 3000
    .ColumnHeaders.Add , , "Employee Terms Access", 3000
    
    .View = lvwReport
End With

With lvwEmp
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "Emp. Code", 1200
    .ColumnHeaders.Add , , "SurName"
    .ColumnHeaders.Add , , "Other Names", 4000
    
    .View = lvwReport
End With

End Sub

Public Sub LoadList()
lvwUsers.ListItems.Clear

With rs2
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
            Set LI = lvwUsers.ListItems.Add(, , !UID & "", , 26)
            LI.ListSubItems.Add , , !Description & ""
            LI.ListSubItems.Add , , !GNo & ""
            
            With rs1
                If .RecordCount > 0 Then
                    '.Filter = "GNo like '" & rs2!GNo & "" & "'"
                    .MoveFirst
                    .Find "GNo like '" & rs2!GNo & "'", , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
'                        If .RecordCount > 0 Then
                        LI.ListSubItems.Add , , !Description & ""
                    Else
                        LI.ListSubItems.Add , , ""
                    End If
                    
'                    .Filter = adFilterNone
                Else
                    LI.ListSubItems.Add , , ""
                End If
            End With
            
            LI.ListSubItems.Add , , !DName & ""
            LI.ListSubItems.Add , , !terms & ""
            
                    
            .MoveNext
        Loop
    End If
End With

End Sub

Public Sub myStructure()
Dim mm As String
trwStruc.Nodes.Clear

Set MyNodes = trwStruc.Nodes.Add(, , "O", "Company Structure")

Set rs6 = CConnect.GetRecordSet("SELECT * FROM CStructure WHERE SCode = '01' ORDER BY MyLevel, Code")

With rs6
    If .RecordCount > 0 Then
        .MoveFirst
        CNode = !LCode & ""
        
        Do While Not .EOF
            If !MyLevel = 0 Then
                Set MyNodes = trwStruc.Nodes.Add("O", tvwChild, !LCode, !code & ",  " & !Description & "")
                MyNodes.EnsureVisible
            Else
                Set MyNodes = trwStruc.Nodes.Add(!PCode & "", tvwChild, !LCode & "", !code & ", " & !Description & "")
                MyNodes.EnsureVisible
            End If
            
            .MoveNext
        Loop
        .MoveFirst
        
        
    End If
End With


    
    
End Sub
