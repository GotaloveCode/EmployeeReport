VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGradeSteps 
   BorderStyle     =   0  'None
   Caption         =   "Currencies"
   ClientHeight    =   6615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7275
   Icon            =   "frmGradeSteps.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab sstCurrencies 
      Height          =   5715
      Left            =   150
      TabIndex        =   1
      Top             =   675
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   10081
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Currencies"
      TabPicture(0)   =   "frmGradeSteps.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdNewC"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdEditC"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdDeleteC"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdDenominations"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdClose"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Denominations"
      TabPicture(1)   =   "frmGradeSteps.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdBack"
      Tab(1).Control(1)=   "cmdDelete"
      Tab(1).Control(2)=   "cmdEdit"
      Tab(1).Control(3)=   "cmdNew"
      Tab(1).Control(4)=   "txtParentCurrency"
      Tab(1).Control(5)=   "Frame3"
      Tab(1).Control(6)=   "fraDenominations"
      Tab(1).Control(7)=   "Label5"
      Tab(1).ControlCount=   8
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   465
         Left            =   5700
         TabIndex        =   29
         Top             =   5100
         Width           =   1215
      End
      Begin VB.CommandButton cmdDenominations 
         Caption         =   "Denominations..."
         Height          =   465
         Left            =   3825
         TabIndex        =   28
         Top             =   5100
         Width           =   1440
      End
      Begin VB.CommandButton cmdDeleteC 
         Caption         =   "Delete"
         Height          =   465
         Left            =   2475
         TabIndex        =   27
         Top             =   5100
         Width           =   1140
      End
      Begin VB.CommandButton cmdEditC 
         Caption         =   "Edit"
         Height          =   465
         Left            =   1425
         TabIndex        =   26
         Top             =   5100
         Width           =   990
      End
      Begin VB.CommandButton cmdNewC 
         Caption         =   "New"
         Height          =   465
         Left            =   300
         TabIndex        =   25
         Top             =   5100
         Width           =   990
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back To Currencies"
         Height          =   465
         Left            =   -69750
         TabIndex        =   24
         Top             =   5100
         Width           =   1590
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   465
         Left            =   -72150
         TabIndex        =   23
         Top             =   5100
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   465
         Left            =   -73575
         TabIndex        =   22
         Top             =   5100
         Width           =   1215
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   465
         Left            =   -74850
         TabIndex        =   21
         Top             =   5100
         Width           =   1140
      End
      Begin VB.TextBox txtParentCurrency 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   -73950
         TabIndex        =   20
         Top             =   525
         Width           =   5490
      End
      Begin VB.Frame Frame3 
         Height          =   1140
         Left            =   -74850
         TabIndex        =   14
         Top             =   1050
         Width           =   6765
         Begin VB.TextBox txtValue 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1275
            TabIndex        =   18
            Top             =   675
            Width           =   2490
         End
         Begin VB.TextBox txtDenomination 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1275
            TabIndex        =   16
            Top             =   300
            Width           =   5040
         End
         Begin VB.Label Label7 
            Caption         =   "Value:"
            Height          =   165
            Left            =   150
            TabIndex        =   17
            Top             =   750
            Width           =   915
         End
         Begin VB.Label Label6 
            Caption         =   "Denomination:"
            Height          =   165
            Left            =   75
            TabIndex        =   15
            Top             =   300
            Width           =   1065
         End
      End
      Begin VB.Frame fraDenominations 
         Caption         =   "Denominations"
         Height          =   2640
         Left            =   -74850
         TabIndex        =   12
         Top             =   2325
         Width           =   6765
         Begin MSComctlLib.ListView lvwDenominations 
            Height          =   2115
            Left            =   150
            TabIndex        =   13
            Top             =   300
            Width           =   6465
            _ExtentX        =   11404
            _ExtentY        =   3731
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Denomination"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Value"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Currency"
               Object.Width           =   5292
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Existing Currencies:"
         Height          =   3090
         Left            =   150
         TabIndex        =   3
         Top             =   1800
         Width           =   6765
         Begin MSComctlLib.ListView lvwCurrencies 
            Height          =   2565
            Left            =   150
            TabIndex        =   11
            Top             =   375
            Width           =   6465
            _ExtentX        =   11404
            _ExtentY        =   4524
            View            =   3
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Currency Name"
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Symbol"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Conversion Rate"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1215
         Left            =   150
         TabIndex        =   2
         Top             =   450
         Width           =   6765
         Begin VB.TextBox txtConversionRate 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1425
            TabIndex        =   10
            Top             =   675
            Width           =   915
         End
         Begin VB.TextBox txtSymbol 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5775
            TabIndex        =   9
            Top             =   225
            Width           =   840
         End
         Begin VB.TextBox txtCurrencyName 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1425
            TabIndex        =   8
            Top             =   225
            Width           =   3465
         End
         Begin VB.CheckBox chkIsBaseCurrency 
            Caption         =   "This is the Base Currency"
            Height          =   240
            Left            =   2775
            TabIndex        =   7
            Top             =   712
            Width           =   2190
         End
         Begin VB.Label Label4 
            Caption         =   "Conversion Rate:"
            Height          =   240
            Left            =   75
            TabIndex        =   6
            Top             =   712
            Width           =   1290
         End
         Begin VB.Label Label3 
            Caption         =   "Symbol:"
            Height          =   240
            Left            =   5100
            TabIndex        =   5
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Currency Name:"
            Height          =   240
            Left            =   75
            TabIndex        =   4
            Top             =   247
            Width           =   1215
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Currency:"
         Height          =   240
         Left            =   -74775
         TabIndex        =   19
         Top             =   600
         Width           =   840
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Currencies"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   2895
   End
End
Attribute VB_Name = "frmGradeSteps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBack_Click()
    Me.sstCurrencies.Tab = 0
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDenominations_Click()
    Me.sstCurrencies.Tab = 1
End Sub

Private Sub Form_Load()
    frmMain2.PositionTheFormWithoutEmpList Me
End Sub

