VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmCEStructure 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   Caption         =   "Division Employee"
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9975
   Icon            =   "frmCEStructure.frx":0000
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
      TabIndex        =   0
      Top             =   -90
      Width           =   9930
      Begin MSComctlLib.ImageList imgTree 
         Left            =   7080
         Top             =   1485
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
               Picture         =   "frmCEStructure.frx":0442
               Key             =   "A"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCEStructure.frx":0894
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCEStructure.frx":0CE6
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCEStructure.frx":1000
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCEStructure.frx":1452
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCEStructure.frx":18A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCEStructure.frx":1CF6
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCEStructure.frx":2010
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCEStructure.frx":232A
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCEStructure.frx":277C
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCEStructure.frx":2A96
               Key             =   "B"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCEStructure.frx":2EE8
               Key             =   "C"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCEStructure.frx":333A
               Key             =   "D"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCEStructure.frx":378C
               Key             =   "E"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCEStructure.frx":3BDE
               Key             =   "F"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCEStructure.frx":4030
               Key             =   "G"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCEStructure.frx":4482
               Key             =   "H"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCEStructure.frx":48D4
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCEStructure.frx":4D26
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCEStructure.frx":5178
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCEStructure.frx":55CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCEStructure.frx":5A1C
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCEStructure.frx":5E6E
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCEStructure.frx":62C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCEStructure.frx":6712
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCEStructure.frx":6B64
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCEStructure.frx":6FB6
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Frame Fralist 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Select Employee"
         ForeColor       =   &H80000008&
         Height          =   6255
         Left            =   3960
         TabIndex        =   12
         Top             =   375
         Visible         =   0   'False
         Width           =   5490
         Begin VB.ComboBox cboDept 
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
            Left            =   945
            TabIndex        =   24
            Text            =   "Department"
            Top             =   5790
            Visible         =   0   'False
            Width           =   2970
         End
         Begin VB.CheckBox chkHOD 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "H.O.D"
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
            TabIndex        =   23
            Top             =   5775
            Width           =   1065
         End
         Begin VB.Frame fraRange 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
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
            Height          =   885
            Left            =   1125
            TabIndex        =   18
            Top             =   5340
            Visible         =   0   'False
            Width           =   1785
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
               Height          =   285
               Left            =   570
               TabIndex        =   20
               Top             =   480
               Width           =   1170
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
               Height          =   285
               Left            =   570
               TabIndex        =   19
               Top             =   90
               Width           =   1170
            End
            Begin VB.Label Label2 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00C0E0FF&
               Caption         =   "To"
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
               TabIndex        =   22
               Top             =   450
               Width           =   180
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00C0E0FF&
               Caption         =   "From"
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
               TabIndex        =   21
               Top             =   60
               Width           =   360
            End
         End
         Begin VB.CheckBox chkRange 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "Range"
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
            TabIndex        =   17
            Top             =   5400
            Width           =   1065
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
            Left            =   210
            Picture         =   "frmCEStructure.frx":7408
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Cancel Process"
            Top             =   4800
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton cmdCanc 
            Caption         =   "DONE"
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
            Left            =   3990
            TabIndex        =   15
            Top             =   5760
            Width           =   1365
         End
         Begin VB.CommandButton cmdSelect 
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
            Height          =   375
            Left            =   3990
            TabIndex        =   13
            Top             =   5400
            Width           =   1365
         End
         Begin MSComctlLib.ListView lvwEmp 
            Height          =   4710
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   5235
            _ExtentX        =   9234
            _ExtentY        =   8308
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
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
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   1380
            Top             =   4200
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label lblECount 
            Caption         =   "Employee Count:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   105
            TabIndex        =   26
            Top             =   315
            Width           =   1485
         End
         Begin VB.Label lblEmpCount 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1590
            TabIndex        =   25
            Top             =   315
            Width           =   645
         End
      End
      Begin MSComctlLib.ListView lvwDetails 
         Height          =   6930
         Left            =   3210
         TabIndex        =   11
         Top             =   90
         Width           =   6720
         _ExtentX        =   11853
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
      Begin MSComctlLib.TreeView trwStruc 
         Height          =   5520
         Left            =   0
         TabIndex        =   9
         Top             =   1500
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   9737
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   529
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
         Height          =   1380
         Left            =   0
         TabIndex        =   10
         Top             =   90
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   2434
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
      TabIndex        =   8
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
      Picture         =   "frmCEStructure.frx":750A
      Style           =   1  'Graphical
      TabIndex        =   7
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
      Picture         =   "frmCEStructure.frx":79FC
      Style           =   1  'Graphical
      TabIndex        =   6
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
      Picture         =   "frmCEStructure.frx":7AFE
      Style           =   1  'Graphical
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   1
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
            Picture         =   "frmCEStructure.frx":7C00
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCEStructure.frx":7D12
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCEStructure.frx":7E24
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCEStructure.frx":7F36
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCEStructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim MyNodes As Node
Dim CNode As String
Dim CSt_ID As Integer
Dim Rangeto As Boolean
Dim dept_sect As String

Private Sub cboDept_Click()
    If cboDept.Text = "" Then
        cboDept.Text = "Department"
    End If

End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
    
End Sub

Private Sub chkHOD_Click()
    chkRange.Value = 0
    If chkHOD.Value = 1 Then
        cboDept.Visible = True
    Else
        cboDept.Visible = False
    End If
    
    
End Sub

Private Sub chkRange_Click()
    If chkRange.Value = 1 Then
        frarange.Visible = True
        chkHOD.Value = 0
    Else
        frarange.Visible = False
    End If
    
End Sub

Public Sub cmdCancel_Click()
    FraList.Visible = False
    
    Call EnableCmd
    cmdCancel.Enabled = False
    SaveNew = False
    
    With frmMain2
        .cmdNew3.Enabled = True
        .cmdDelete3.Enabled = True
        .cmdCancel3.Enabled = True
        .cmdSave_Click
    End With
    
    Call disabletxt
    
End Sub

Private Sub cmdCanc_Click()
    
    FraList.Visible = False
    PSave = True
    Call cmdCancel_Click
    PSave = False
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub cmdDelete_Click()
Dim resp As String

If trwStruc.Nodes.Count < 2 Then
    Exit Sub
End If

Set rs3 = CConnect.GetRecordSet("SELECT * FROM CStructure WHERE SCode = '" & lvwSTypes.SelectedItem & "' AND PCode = '" & CNode & "'")

With rs3
    If .RecordCount > 0 Then
        MsgBox "You cannot delete employees at this level. You can only delete employees at the lowest level.", vbInformation
        cmdCanc_Click
        Exit Sub
    End If
End With

Set rs3 = Nothing

    
If lvwDetails.ListItems.Count > 0 Then
    If trwStruc.Nodes.Count > 0 Then
        If Not trwStruc.SelectedItem = "" Then
            resp = MsgBox("Are you sure you want to delete employee: " & lvwDetails.SelectedItem & " from this division?", vbQuestion + vbYesNo)
            If resp = vbNo Then
                Exit Sub
            End If
                
            Action = "DETACHED EMPLOYEE FROM DEPARTMENT; EMPLOYEE CODE: " & lvwDetails.SelectedItem & "; DEPARTMENT NAME: " & trwStruc.SelectedItem.Text
            
            CConnect.ExecuteSql ("DELETE FROM SEmp WHERE LCode like '" & trwStruc.SelectedItem.Key & "%" & "' AND SCode = '" & lvwSTypes.SelectedItem & "' AND EmpCode = '" & lvwDetails.SelectedItem & "'")
            rs1.Requery
                  
            Call LoadDList
        End If
   End If
   
End If

End Sub

Private Sub cmdDone_Click()
    PSave = True
    Call cmdCancel_Click
    PSave = False
End Sub

Public Sub cmdNew_Click()
Dim i As Long
Call DisableCmd
Call Cleartxt
If CNode = "" Then
    MsgBox "You have to select the node with you will add the employee to.", vbInformation
    PSave = True
    Call cmdCancel_Click
    PSave = False
    Exit Sub
End If
CSt_ID = trwStruc.SelectedItem.Tag
If trwStruc.Nodes.Count < 2 Then
    Exit Sub
End If
Set rs3 = CConnect.GetRecordSet("SELECT * FROM CStructure WHERE SCode = '" & lvwSTypes.SelectedItem & "' AND PCode = '" & CNode & "'")
With rs3
    If .RecordCount > 0 Then
        MsgBox "You cannot add employees at this level. You can only add employees at the lowest level.", vbInformation
        cmdCanc_Click
        Exit Sub
          
    End If
End With
Set rs3 = Nothing
FraList.Visible = True
cmdCancel.Enabled = True
cmdSelect.Enabled = True
cmdCanc.Enabled = True
SaveNew = True

Call enabletxt

End Sub

Private Sub cmdSelect_Click()
Dim HOD As String
Dim DCode As String


HOD = "No"
DCode = ""

If cboDept.Text <> "Department" And cboDept.Text <> "" Then
    With rs4
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "Code like '" & left(cboDept.Text, InStr(cboDept.Text, ":") - 1) & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                DCode = !RLCode & ""
            
            End If
        End If
    End With
End If
            

If lvwEmp.ListItems.Count > 0 Then
   If chkRange.Value = 1 Then
    If txtFrom.Text = "" Or txtTo.Text = "" Then
        MsgBox "Enter a valid Range", vbInformation
        Exit Sub
    End If
        If MsgBox("This will replace any existing divisions.Continue ", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        With rsGlob
            If .RecordCount > 0 Then
            .MoveFirst
            .Find "EmpCode='" & txtFrom.Text & "'", , adSearchForward, adBookmarkFirst
            If .EOF = True Then Exit Sub
                Me.MousePointer = vbHourglass
                While !empcode <> txtTo.Text
                    
                    Action = "ADDED EMPLOYEE TO DEPARTMENT; EMPLOYEE CODE: " & rsGlob!empcode & "; DEPARTMENT NAME: " & trwStruc.SelectedItem.Text
                    
                    CConnect.ExecuteSql ("DELETE FROM SEmp WHERE SCode = '" & lvwSTypes.SelectedItem & "' AND employee_id = '" & rsGlob!employee_id & "'")
                    CConnect.ExecuteSql ("INSERT INTO SEmp (SCode, LCode, EmpCode, employee_id, HOD,cstructure_id) VALUES('" & lvwSTypes.SelectedItem & "','" & CNode & "" & "','" & rsGlob!empcode & "', '" & rsGlob!employee_id & "', " & chkHOD.Value & "," & CSt_ID & ")")
                    .MoveNext
                Wend
                
                CConnect.ExecuteSql ("DELETE FROM SEmp WHERE SCode = '" & lvwSTypes.SelectedItem & "' AND employee_id = '" & rsGlob!employee_id & "'")
                CConnect.ExecuteSql ("INSERT INTO SEmp (SCode, LCode, EmpCode, employee_id, HOD) VALUES('" & lvwSTypes.SelectedItem & "','" & CNode & "" & "','" & rsGlob!empcode & "', '" & rsGlob!employee_id & "', " & chkHOD.Value & "," & CSt_ID & ")")
                CConnect.ExecuteSql "UPDATE Employee SET cstructure_id=" & CSt_ID & " WHERE employee_id='" & rsGlob!employee_id & "'"
                Me.MousePointer = vbDefault
            End If
        End With
        
'        ' if AuditTrail = True Then cConnect.ExecuteSql ("INSERT INTO AuditTrail (UserId, DTime, Trans, TDesc, MySection)VALUES('" & CurrentUser & "','" & Date & " " & Time & "','Inserting Employee to a division','" & lvwSTypes.SelectedItem & "-" & CNode & "-" & txtFrom.Text & "to" & txtTo.Text & "','Set-up')")
        
        rs1.Requery
        
        Call LoadDList
        
        MsgBox "Employee Successfully selected.", vbInformation
   Else
        
        With rs1
            If .RecordCount > 0 Then
                .Filter = "LCode like '" & CNode & "'"
                If .RecordCount > 0 Then
                    .MoveFirst
                    .Find "employee_id like '" & lvwEmp.SelectedItem.Tag & "'", , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                        resp = MsgBox("This employee already exists in one of the divisions. Do you wish to continue?", vbQuestion + vbYesNo)
                        If resp = vbNo Then
                            Exit Sub
                        End If
                    End If
                End If
                .Filter = adFilterNone
            End If
            
            If PromptSave = True Then
                If MsgBox("Are you sure you want to save the record.", vbQuestion + vbYesNo) = vbNo Then
    
                    Exit Sub
                End If
            End If
                
            Action = "ADDED EMPLOYEE TO DEPARTMENT; EMPLOYEE CODE: " & rsGlob!empcode & "; DEPARTMENT NAME: " & trwStruc.SelectedItem.Text
            
            CConnect.ExecuteSql ("DELETE FROM SEmp WHERE SCode = '" & lvwSTypes.SelectedItem & "' AND employee_id = '" & lvwEmp.SelectedItem.Tag & "'")
            CConnect.ExecuteSql ("INSERT INTO SEmp (SCode, LCode, EmpCode, employee_id, HOD, DCode) VALUES('" & lvwSTypes.SelectedItem & "','" & CNode & "" & "','" & lvwEmp.SelectedItem & "'," & lvwEmp.SelectedItem.Tag & ", " & chkHOD.Value & ", '" & DCode & "')")
            
            rs1.Requery
            
            Call LoadDList
            
            MsgBox "Employee Successfully selected.", vbInformation
        End With
    End If
  
Else
    MsgBox "No records to be selected."
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

frmMain2.lblEmpCount.Caption = 0

Call InitGrid

Set rs5 = CConnect.GetRecordSet("SELECT * FROM STypes ORDER BY Code")
            
With rs5
    If .RecordCount > 0 Then
        .MoveFirst
        
        Set rs1 = CConnect.deptFilter("SELECT * FROM pVwRsGlob WHERE SCode='" & rs5!code & "' AND Term <> 1 ORDER BY EmpCode")
                        
        Call LoadSEmp
        
        Do While Not .EOF
            Set LI = lvwSTypes.ListItems.Add(, , !code & "", , 2)
            LI.ListSubItems.Add , , !Description & ""
            
            .MoveNext
        Loop
        
    End If
End With

With rsGlob
    If .RecordCount < 1 Then
        Call DisableCmd
        Exit Sub
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

Call LoadList

Call LoadCbo

End Sub

Private Sub Form_Resize()
oSmart.FResize Me


End Sub

Private Sub LoadDList()
Dim i As Integer
    Dim rC As Integer
    lvwDetails.ListItems.Clear
    frmMain2.lblEmpCount.Caption = 0

If CNode = "O" Then
    Call LoadSEmp
End If

With rs1
    If .RecordCount > 0 Then
        dept_sect = Mid(dept_sect, rC + 1)
        While dept_sect <> ""
            rC = InStr(dept_sect, vbTab)
            CNode = left(dept_sect, rC - 1)
            If .Filter <> 0 Then
                .Filter = .Filter & " or LCode = '" & CNode & "'"
            Else
                .Filter = "LCode = '" & CNode & "'"
            End If
            dept_sect = Mid(dept_sect, rC + 1)
        Wend
        If .RecordCount > 0 Then
            .MoveFirst
            i = 0
            Do While Not .EOF
                If i = 1 Then
                    Set LI = lvwDetails.ListItems.Add(, , !empcode & "", , 6)
                    i = 0
                Else
                    Set LI = lvwDetails.ListItems.Add(, , !empcode & "", , 7)
                    i = 1
                End If
                                
                LI.ListSubItems.Add , , !SurName & ""
                LI.ListSubItems.Add , , !OtherNames & ""
                LI.ListSubItems.Add , , IIf(Trim(!HOD & "") = True, "Yes", "")


                .MoveNext
            Loop
            .MoveFirst
            
            frmMain2.lblEmpCount.Caption = .RecordCount
        End If
        .Filter = adFilterNone
    End If
End With

End Sub

Private Sub LoadList()
lvwEmp.ListItems.Clear

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
        
        lblEmpCount.Caption = .RecordCount
    End If
End With


End Sub

Private Sub InitGrid()
    With lvwSTypes
        .ColumnHeaders.Add , , "Code", 300
        .ColumnHeaders.Add , , "Description", 2850
                
        .View = lvwReport
    End With
    
    With lvwDetails
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Employee Code"
        .ColumnHeaders.Add , , "SurName", 1500
        .ColumnHeaders.Add , , "Other Names", 2000
'        .ColumnHeaders.Add , , "ID No"
        .ColumnHeaders.Add , , "H.O.D", 800, vbCenter
         .ColumnHeaders.Add , , "Department", 2500
        .View = lvwReport
    End With
    
    With lvwEmp
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Employee Code"
        .ColumnHeaders.Add , , "SurName"
        .ColumnHeaders.Add , , "Other Names.", 2500
        
        
        .View = lvwReport
    End With
 
End Sub


Private Sub Form_Unload(Cancel As Integer)
    frmMain2.Caption = "Personnel Director " & App.FileDescription
    
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
If lvwEmp.ListItems.Count < 1 Then Exit Sub
    If chkRange.Value = 1 Then
        If Rangeto = False Then
            txtFrom.Text = lvwEmp.SelectedItem
            txtTo.SetFocus
        Else
            txtTo.Text = lvwEmp.SelectedItem
            txtFrom.SetFocus
        End If
    Else
        Call cmdSelect_Click
    End If
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
        
        Set rs1 = CConnect.deptFilter("SELECT S.HOD, e.*,c.*" & _
                        " FROM SEmp as s LEFT JOIN Employee as e ON s.employee_id = e.employee_id LEFT JOIN cstructure as c ON e.cstructure_id = c.cstructure_id" & _
                        " WHERE s.SCode = '" & lvwSTypes.SelectedItem & "' AND e.Term <> 1 ORDER BY e.EmpCode")
        
        Call LoadSEmp
    End If
    
    
    CNode = ""
    
End Sub

Private Sub Text4_Change()

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
    On Error GoTo errHandler
    trwStruc.Nodes.Clear
    
    Set MyNodes = trwStruc.Nodes.Add(, , "O", rs5!Description & "")
    MyNodes.Tag = rs5!code & ""
    
    Set rs = CConnect.GetRecordSet("SELECT * FROM CStructure WHERE SCode = '" & rs5!code & "' ORDER BY MyLevel, Code")
    
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            CNode = !LCode & ""
            
            Do While Not .EOF
                If !MyLevel = 0 Then
                    Set MyNodes = trwStruc.Nodes.Add("O", tvwChild, !LCode & "", !Description & "")
                    MyNodes.EnsureVisible
    
                Else
                    Set MyNodes = trwStruc.Nodes.Add(!PCode & "", tvwChild, !LCode & "", !Description & "")
                    MyNodes.EnsureVisible
    
                End If
                MyNodes.Tag = !cstructure_id & ""
                .MoveNext
            Loop
            .MoveFirst
            
            
        End If
    End With
    Exit Sub
errHandler:
End Sub

Public Sub LoadSEmp()
lvwDetails.ListItems.Clear
frmMain2.lblEmpCount.Caption = 0

Dim i As Integer

With rs1
    If .RecordCount > 0 Then
        .MoveFirst
        i = 0
        Do While Not .EOF
            If i = 1 Then
                Set LI = lvwDetails.ListItems.Add(, , !empcode & "", , 6)
                i = 0
            Else
                Set LI = lvwDetails.ListItems.Add(, , !empcode & "", , 7)
                i = 1
            End If
            LI.Tag = !employee_id & ""
            LI.ListSubItems.Add , , !SurName & ""
            LI.ListSubItems.Add , , !OtherNames & ""
            LI.ListSubItems.Add , , !HOD & ""
            
                        
            .MoveNext
        Loop
        .MoveFirst
        
        frmMain2.lblEmpCount.Caption = .RecordCount
    End If
End With
    
End Sub

Private Sub trwStruc_NodeClick(ByVal Node As MSComctlLib.Node)
    dept_sect = ""
    CNode = trwStruc.SelectedItem.Key & ""
    CSt_ID = trwStruc.SelectedItem.Tag
    If isDepartmentOrSection(trwStruc.SelectedItem.Key) = True Then
        dept_sect = getChildNodes(trwStruc.SelectedItem.Key)
    End If
    Call LoadDList
    
End Sub

Private Sub txtFrom_GotFocus()
    Rangeto = False
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
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

Public Sub LoadCbo()
cboDept.Clear

Set rs4 = CConnect.GetRecordSet("SELECT * FROM MyDivisions ORDER BY Code")

With rs4
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
            cboDept.AddItem (!code & ":" & !Description & "")
             
            .MoveNext
        Loop
        .MoveFirst
    End If
End With

'Set rs4 = Nothing

End Sub
