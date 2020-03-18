VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmCStructure 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   Caption         =   "Structures"
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9975
   Icon            =   "frmCStructure.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   7020
      Left            =   15
      TabIndex        =   17
      Top             =   -90
      Width           =   9930
      Begin VB.Frame fraDetails 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Division"
         ForeColor       =   &H80000008&
         Height          =   5205
         Left            =   3300
         TabIndex        =   26
         Top             =   840
         Visible         =   0   'False
         Width           =   5970
         Begin VB.TextBox txtPerc 
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2040
            TabIndex        =   13
            Top             =   4470
            Width           =   600
         End
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
            Height          =   495
            Left            =   4815
            TabIndex        =   14
            Top             =   4575
            Width           =   1020
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
            Height          =   660
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   2580
            Width           =   5700
         End
         Begin VB.TextBox txtNssf 
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
            Left            =   2190
            TabIndex        =   9
            Top             =   2100
            Width           =   1770
         End
         Begin VB.TextBox txtPinNo 
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
            Left            =   2190
            TabIndex        =   7
            Top             =   1590
            Width           =   1770
         End
         Begin VB.TextBox txtNhif 
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
            Left            =   4080
            TabIndex        =   8
            Top             =   1590
            Width           =   1740
         End
         Begin VB.TextBox txtLasc 
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
            Left            =   4080
            TabIndex        =   10
            Top             =   2100
            Width           =   1740
         End
         Begin VB.TextBox txtFax 
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
            Left            =   135
            TabIndex        =   6
            Top             =   1590
            Width           =   1935
         End
         Begin VB.TextBox txtTelNo 
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
            Left            =   135
            TabIndex        =   4
            Top             =   1065
            Width           =   1935
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
            Height          =   285
            Left            =   2175
            TabIndex        =   5
            Top             =   1065
            Width           =   3645
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
            Height          =   285
            Left            =   1485
            TabIndex        =   3
            Top             =   555
            Width           =   4335
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
            TabIndex        =   2
            Top             =   555
            Width           =   1230
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
            Height          =   765
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   3600
            Width           =   5715
         End
         Begin VB.CommandButton cmdSave 
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
            Left            =   4830
            Picture         =   "frmCStructure.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Save Record"
            Top             =   4575
            Width           =   510
         End
         Begin VB.CommandButton cmdCancel 
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
            Left            =   5340
            Picture         =   "frmCStructure.frx":0544
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Cancel Process"
            Top             =   4575
            Width           =   495
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            Caption         =   "Minimum % of Employees"
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
            Top             =   4425
            Width           =   1815
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
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
            TabIndex        =   37
            Top             =   2355
            Width           =   585
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            Caption         =   "N.S.S.F No"
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
            TabIndex        =   36
            Top             =   1875
            Width           =   795
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            Caption         =   "P.I.N No"
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
            TabIndex        =   35
            Top             =   1365
            Width           =   615
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            Caption         =   "N.H.I.F No"
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
            Left            =   4080
            TabIndex        =   34
            Top             =   1350
            Width           =   780
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            Caption         =   "L.A.S.C No"
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
            Left            =   4080
            TabIndex        =   33
            Top             =   1875
            Width           =   795
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            Caption         =   "Fax"
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
            Top             =   1380
            Width           =   270
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            Caption         =   "Tel No"
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
            TabIndex        =   31
            Top             =   840
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
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
            Left            =   2160
            TabIndex        =   30
            Top             =   840
            Width           =   420
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   1485
            TabIndex        =   29
            Top             =   315
            Width           =   795
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
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
            TabIndex        =   28
            Top             =   315
            Width           =   375
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
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
            TabIndex        =   27
            Top             =   3375
            Width           =   750
         End
      End
      Begin MSComctlLib.ImageList imgTree 
         Left            =   3960
         Top             =   1395
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   27
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCStructure.frx":0646
               Key             =   "A"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCStructure.frx":0A98
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCStructure.frx":0EEA
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCStructure.frx":1204
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCStructure.frx":1656
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCStructure.frx":1AA8
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCStructure.frx":1EFA
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCStructure.frx":2214
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCStructure.frx":252E
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCStructure.frx":2980
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCStructure.frx":2C9A
               Key             =   "B"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCStructure.frx":30EC
               Key             =   "C"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCStructure.frx":353E
               Key             =   "D"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCStructure.frx":3990
               Key             =   "E"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCStructure.frx":3DE2
               Key             =   "F"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCStructure.frx":4234
               Key             =   "G"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCStructure.frx":4686
               Key             =   "H"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCStructure.frx":4AD8
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCStructure.frx":4F2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCStructure.frx":537C
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCStructure.frx":57CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCStructure.frx":5C20
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCStructure.frx":6072
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCStructure.frx":64C4
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCStructure.frx":6916
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCStructure.frx":6D68
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCStructure.frx":71BA
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView trwStruc 
         Height          =   6900
         Left            =   2480
         TabIndex        =   1
         Top             =   120
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   12171
         _Version        =   393217
         HideSelection   =   0   'False
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
      Begin MSComDlg.CommonDialog Cdl 
         Left            =   8100
         Top             =   6285
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ListView lvwSTypes 
         Height          =   6930
         Left            =   0
         TabIndex        =   0
         Top             =   90
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   12224
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
      Left            =   1485
      TabIndex        =   25
      ToolTipText     =   "Move to the Last employee"
      Top             =   5145
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
      Left            =   3660
      Picture         =   "frmCStructure.frx":760C
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Delete Record"
      Top             =   5145
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
      Left            =   3180
      Picture         =   "frmCStructure.frx":7AFE
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Edit Record"
      Top             =   5145
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
      Left            =   2700
      Picture         =   "frmCStructure.frx":7C00
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Add New record"
      Top             =   5145
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
      Left            =   1005
      TabIndex        =   21
      ToolTipText     =   "Move to the Next employee"
      Top             =   5145
      Visible         =   0   'False
      Width           =   495
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
      Left            =   525
      TabIndex        =   20
      ToolTipText     =   "Move to the Previous employee"
      Top             =   5145
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
      Left            =   6195
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5145
      Visible         =   0   'False
      Width           =   1050
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
      Left            =   45
      TabIndex        =   18
      ToolTipText     =   "Move to the First employee"
      Top             =   5145
      Visible         =   0   'False
      Width           =   495
   End
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
            Picture         =   "frmCStructure.frx":7D02
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCStructure.frx":7E14
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCStructure.frx":7F26
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCStructure.frx":8038
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCStructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim MyNodes As Node
Dim CNode As String
Dim PNode As String
Dim oldCode As String

Public Sub cmdCancel_Click()
'    If PSave = False Then
'        If MsgBox("Do you want to save changes?", vbQuestion + vbYesNo) = vbYes Then  '
'            Call cmdSave_Click
'            Exit Sub
'        End If
'    End If
'
'    fraDetails.Visible = False
'
'    Call EnableCmd
'    cmdCancel.Enabled = False
'    cmdSave.Enabled = False
'    SaveNew = False
'
'    With frmMain2
'        .cmdNew.Enabled = True
'        .cmdEdit.Enabled = True
'        .cmdDelete.Enabled = True
'        .cmdCancel.Enabled = False
'        .cmdSave.Enabled = False
'    End With
'
'    Call disabletxt
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

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub cmdDelete_Click()
On Error Resume Next
Dim resp As String

If trwStruc.Nodes.Count > 0 Then
    If Not trwStruc.SelectedItem = "" Then
        resp = MsgBox("This will delete " & trwStruc.SelectedItem & ", its branches and subsequently remove employee from it's division. Do you wish to continue?", vbQuestion + vbYesNo)
        If resp = vbNo Then
            Exit Sub
        End If
            
        DisplayRecords
        Omnis_ActionTag = "D"
        Departments_TextFile
                
        CConnect.ExecuteSql ("DELETE FROM CStructure WHERE LCode like '" & trwStruc.SelectedItem.Key & "%" & "' AND SCode = '" & lvwSTypes.SelectedItem & "'")
        Action = "DELETED COMPANY STRUCTURE; DEPARTMENT/SECTION CODE: " & trwStruc.SelectedItem.Key & "; STRUCTURE CODE: " & lvwSTypes.SelectedItem
        CConnect.ExecuteSql ("DELETE FROM CStructure WHERE LCode = '" & trwStruc.SelectedItem.Key & "' AND SCode = '" & lvwSTypes.SelectedItem & "'")
        Action = "DELETED EMPLOYEES IN COMPANY STRUCTURE; DEPARTMENT/SECTION CODE: " & trwStruc.SelectedItem.Key & "; STRUCTURE CODE: " & lvwSTypes.SelectedItem
        CConnect.ExecuteSql ("DELETE FROM SEmp WHERE LCode like '" & trwStruc.SelectedItem.Key & "%" & "' AND SCode = '" & lvwSTypes.SelectedItem & "'")
        
        Call myStructure
   
   End If
End If

End Sub

Private Sub cmdDone_Click()
    PSave = True
    Call cmdCancel_Click
    PSave = False
End Sub

Public Sub cmdEdit_Click()

Omnis_ActionTag = "E"

SaveNew = False
If trwStruc.Nodes.Count > 0 Then
    If Not CNode = "" Then
        Call DisplayRecords
    Else
        MsgBox "You have to select the item you want to edit.", vbInformation
        PSave = True
        Call cmdCancel_Click
        PSave = False
        Exit Sub
    End If
    
Else
    PSave = True
    Call cmdCancel_Click
    PSave = False
    Exit Sub
End If

Call DisableCmd

fraDetails.Visible = True

cmdSave.Enabled = True
cmdCancel.Enabled = True
CmdOk.Visible = False


Call EnableCmd
txtCode.Locked = False
txtDesc.SetFocus
SaveNew = False
Call enabletxt

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

Omnis_ActionTag = "I"   '++Inserts a new record to the omnis text file  'monte++

SaveNew = True
Call DisableCmd
Call Cleartxt

If CNode = "" Then
    MsgBox "You have to select the node with you will add to.", vbInformation
    PSave = True
    Call cmdCancel_Click
    PSave = False
    Exit Sub
End If

fraDetails.Visible = True
cmdCancel.Enabled = True
SaveNew = True
cmdSave.Enabled = True
CmdOk.Visible = False
txtCode.SetFocus
Call enabletxt

PNode = CNode

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

Private Sub CmdOk_Click()
    fraDetails.Visible = False
    
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
Dim LCode As String
Dim PCode As String
Dim MyLevel As Integer

If txtCode.Text = "" Then
    MsgBox "Enter the code.", vbExclamation
    txtCode.SetFocus
    Exit Sub
End If

If txtDesc.Text = "" Then
    MsgBox "Enter the description.", vbExclamation
    txtDesc.SetFocus
    Exit Sub
End If

If txtPerc.Text = "" Then
    txtPerc.Text = 0
End If

    If SaveNew = True Then
        Set rs4 = CConnect.GetRecordSet("SELECT * FROM CStructure WHERE Code = '" & txtCode.Text & "'")
                
        With rs4
            If .RecordCount > 0 Then
                MsgBox "Code already exists. Enter another one.", vbInformation
                txtCode.Text = ""
                txtCode.SetFocus
                Set rs4 = Nothing
                Exit Sub
            End If
        End With
        
        Set rs4 = Nothing
       
        If PromptSave = True Then
            If MsgBox("Are you sure you want to save the record.", vbQuestion + vbYesNo) = vbNo Then
                PSave = True
                Call cmdCancel_Click
                PSave = False
                Exit Sub
            End If
        End If
    
        With rs
            If .RecordCount > 0 Then
                .MoveFirst
                .Find "LCode like '" & PNode & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    LCode = !LCode & "" & txtCode.Text
                    MyLevel = !MyLevel + 1
                    PCode = !LCode & ""
                Else
                    LCode = "Q" & txtCode.Text
                    MyLevel = 0
                End If
            Else
                LCode = "Q" & txtCode.Text
                MyLevel = 0
            End If
        End With
    
        Action = "REGISTERED COMPANY STRUCTURE; DEPARTMENT/SECTION CODE: " & LCode & "; STRUCTURE CODE: " & lvwSTypes.SelectedItem
        
        CConnect.ExecuteSql ("DELETE FROM CStructure WHERE LCode = '" & LCode & "' AND SCode = '" & lvwSTypes.SelectedItem & "'")
        
        mySQL = "INSERT INTO CStructure (SCode, LCode, Pcode, MyLevel, Code, Description, MPerc, Comments)" & _
                            " VALUES('" & lvwSTypes.SelectedItem & "','" & LCode & "','" & PCode & "'," & MyLevel & ",'" & txtCode.Text & "','" & txtDesc.Text & "'," & _
                            " " & txtPerc.Text & ",'" & txtComments.Text & "')"
        
        CConnect.ExecuteSql (mySQL)

        Departments_TextFile    '++Writes a record with an 'I' to the Omnis text file 'monte++
        
        rs.Requery
    
    Else
        If PromptSave = True Then
            If MsgBox("Are you sure you want to save the record.", vbQuestion + vbYesNo) = vbNo Then
                PSave = True
                Call cmdCancel_Click
                PSave = False
                Exit Sub
            End If
        End If
        
        CConnect.ExecuteSql ("UPDATE CStructure SET CStructure.Code = '" & txtCode.Text & "', CStructure.Description = '" & txtDesc.Text & "'," & _
                        " CStructure.Address = '" & txtAddress.Text & "', CStructure.TelNo = '" & txtTelNo.Text & "', CStructure.Fax = '" & txtFax.Text & "', CStructure.Email = '" & txtEmail.Text & "'," & _
                        " CStructure.PinNo = '" & txtPinNo.Text & "', CStructure.NhifNo = '" & txtNhif.Text & "', CStructure.NssfNo = '" & txtNssf.Text & "', CStructure.LascNo = '" & txtLasc.Text & "'," & _
                        " CStructure.MPerc = " & txtPerc.Text & ", CStructure.Comments = '" & txtComments.Text & "'" & _
                        " WHERE LCode = '" & CNode & "'")
                       
        Departments_TextFile    '++Writes a record with an 'E' to the Omnis text file 'monte++
        
        rs.Requery
    End If
    
        
    Call myStructure
    
    
    If SaveNew = False Then
        PSave = True
        Call cmdCancel_Click
        PSave = False
        Exit Sub
    Else
        Call Cleartxt
        txtCode.SetFocus

    End If
    
    
End Sub




Private Sub Form_Load()
Decla.Security Me
oSmart.FReset Me

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

cmdCancel.Enabled = False
cmdSave.Enabled = False

Call InitGrid
Call CConnect.CCon


Set rs5 = CConnect.GetRecordSet("SELECT * FROM STypes ORDER BY Code")

With rs5
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
            Set LI = lvwSTypes.ListItems.Add(, , !code & "", , 2)
            LI.ListSubItems.Add , , !Description & ""
            
            .MoveNext
        Loop
    End If
End With


cmdFirst.Enabled = False
cmdPrevious.Enabled = False

With rs5
    If .RecordCount > 0 Then
        .MoveFirst
        Call myStructure
        
    End If
End With

End Sub

Private Sub Form_Resize()
oSmart.FResize Me

End Sub

Private Sub InitGrid()
    With lvwSTypes
        .ColumnHeaders.Add , , "Code", 300
        .ColumnHeaders.Add , , "Description", 3000
                
        .View = lvwReport
    End With
    

End Sub

Public Sub DisplayRecords()
Call Cleartxt

If trwStruc.Nodes.Count > 0 Then
    If Not CNode = "" Then
        With rs
            If .RecordCount > 0 Then
                .MoveFirst
                .Find "LCode like '" & CNode & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    txtCode.Text = !code & ""
                    txtDesc.Text = !Description & ""
                    txtTelNo.Text = !TelNo & ""
                    txtFax.Text = !Fax & ""
                    txtAddress.Text = !Address & ""
                    txtEmail.Text = !EMail & ""
                    txtPinNo.Text = !PinNo & ""
                    txtNhif.Text = !NhifNo & ""
                    txtNssf.Text = !NssfNo & ""
                    txtLasc.Text = !LascNo & ""
                    txtComments.Text = !Comments & ""
                    txtPerc.Text = !MPerc & ""
                
                End If
            End If
        End With
  
    End If

End If

End Sub



Private Sub Form_Unload(Cancel As Integer)

    frmMain2.Caption = "Personnel Director " & App.FileDescription
    
End Sub

Private Sub fraList_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub lvwDetails_DblClick()
    If frmMain2.cmdEdit.Enabled = True Then
        Call frmMain2.cmdEdit_Click
    End If
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Search"
            Me.MousePointer = vbHourglass
    
            frmSearch.Show vbModal
            
            If Not Sel = "" Then
                With rsGlob
                    If .RecordCount > 0 Then
                        .MoveFirst
                        .Find "EmpCode like '" & Sel & "'", , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            Call DisplayRecords
                            Call FirstLastDisb
                        End If
                    End If
                End With
                
            End If
      
            Me.MousePointer = 0
    End Select
End Sub







Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error GoTo errHandler
Dim myfile As String
Dim ss As String
Select Case ButtonMenu.Key
    Case "EmpLeaves"
        Me.MousePointer = vbHourglass
        Set a = New Application
        Set R = a.OpenReport(App.Path & "\Leave Reports\Employee Leaves.rpt")
        
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
    Case "LeaveEmp"
        Me.MousePointer = vbHourglass
        Set a = New Application
        Set R = a.OpenReport(App.Path & "\Leave Reports\Leaves Employee.rpt")
        
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

End Select
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

Private Sub txtADays_Change()
'If Val(txtADays.Text) > Val(txtDays.Text) Then
'    txtADays.Text = txtDays.Text
'    txtADays.SelStart = Len(txtADays.Text)
'End If
End Sub

Private Sub txtADays_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
        Case Asc(".")
        Case Asc("-")
        Case Is = 8
        Case Else
        Beep
        KeyAscii = 8
        
    End Select
End Sub

Private Sub txtDays_Keypress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub lvwSTypes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ColumnHeader.Index = 1 Then
        lvwSTypes.ColumnHeaders(1).Width = 1000
    Else
        lvwSTypes.ColumnHeaders(1).Width = 500
    End If

    With lvwSTypes
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvwSTypes_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvwSTypes.ListItems.Count > 0 Then
        With rs5
            If .RecordCount > 0 Then
                .MoveFirst
                .Find "Code like '" & lvwSTypes.SelectedItem & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    Call myStructure
                End If
            End If
        End With
        
    End If
    
    fraDetails.Visible = False
    CNode = ""
    
End Sub

Private Sub Text4_Change()

End Sub



Private Sub trwStruc_DblClick()
Dim i As Long
    If trwStruc.Nodes.Count < 2 Then
        Exit Sub
    End If
    
    Set rs3 = CConnect.GetRecordSet("SELECT Max(MyLevel) as MaxL  from CStructure ")
    
    With rs3
        If .RecordCount > 0 Then
            For i = 0 To !MaxL + 1
                MyNodes.EnsureVisible
            Next i
            
        End If
    End With
        
    Call disabletxt
    fraDetails.Visible = True
    CmdOk.Visible = True
    Call DisplayRecords
    

End Sub

Private Sub trwStruc_NodeClick(ByVal Node As MSComctlLib.Node)
    CNode = trwStruc.SelectedItem.Key & ""
    
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
If Len(Trim(txtAddress.Text)) > 99 Then
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

    
End Sub

Private Sub enabletxt()
Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
            i.Locked = False
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

Public Sub myStructure()
Dim mm As String
On Error GoTo errHandler
trwStruc.Nodes.Clear

Set MyNodes = trwStruc.Nodes.Add(, , "O", rs5!Description & "")

Set rs = CConnect.GetRecordSet("SELECT * FROM CStructure WHERE SCode = '" & rs5!code & "' ORDER BY MyLevel, Code")

With rs
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
errHandler:
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

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
If Len(Trim(txtDesc.Text)) > 198 Then
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

Private Sub txtFax_KeyPress(KeyAscii As Integer)
If Len(Trim(txtFax.Text)) > 49 Then
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

Private Sub txtNhif_KeyPress(KeyAscii As Integer)
If Len(Trim(txtNhif.Text)) > 19 Then
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

Private Sub txtNssf_KeyPress(KeyAscii As Integer)
If Len(Trim(txtNssf.Text)) > 19 Then
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



Private Sub txtPerc_KeyPress(KeyAscii As Integer)
If Val(txtPerc.Text) > 100 Then
    Beep
    MsgBox "Can't enter more than 100 %", vbExclamation
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

Private Sub txtPinNo_KeyPress(KeyAscii As Integer)
If Len(Trim(txtPinNo.Text)) > 19 Then
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

Private Sub txtTelNo_KeyPress(KeyAscii As Integer)
If Len(Trim(txtTelNo.Text)) > 29 Then
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
