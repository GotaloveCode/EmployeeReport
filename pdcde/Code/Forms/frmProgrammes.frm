VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmProgrammes 
   BorderStyle     =   0  'None
   Caption         =   "Programmes"
   ClientHeight    =   6540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvwProgrammes 
      Height          =   2535
      Left            =   120
      TabIndex        =   12
      Top             =   3195
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   4471
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Prog. Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Programme Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Donor Body"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Fund Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Duration (M)"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Start Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "End Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   645
      Width           =   7470
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   330
         Left            =   5850
         TabIndex        =   22
         Top             =   1590
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   82575363
         CurrentDate     =   39210
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   330
         Left            =   3225
         TabIndex        =   20
         Top             =   1590
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   82575363
         CurrentDate     =   39210
      End
      Begin VB.TextBox txtDuration 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   18
         Top             =   1575
         Width           =   915
      End
      Begin VB.TextBox txtDonor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   16
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox txtPCode 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   5640
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtSectorName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   3735
      End
      Begin VB.TextBox txtDetails 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   5490
         TabIndex        =   6
         Top             =   1065
         Width           =   1740
      End
      Begin VB.TextBox txtParentSector 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   675
         Width           =   6015
      End
      Begin VB.Label Label9 
         Caption         =   "End Date:"
         Height          =   255
         Left            =   5025
         TabIndex        =   21
         Top             =   1590
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Start Date:"
         Height          =   180
         Left            =   2400
         TabIndex        =   19
         Top             =   1620
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Duration (M):"
         Height          =   255
         Left            =   150
         TabIndex        =   17
         Top             =   1590
         Width           =   1020
      End
      Begin VB.Label Label6 
         Caption         =   "Donor Body:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1095
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Code:"
         Height          =   255
         Left            =   5160
         TabIndex        =   13
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Prog. Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Fund Code:"
         Height          =   180
         Left            =   4560
         TabIndex        =   9
         Top             =   1125
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Prog. Sector:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   690
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   5910
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   5910
      Width           =   1335
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   5910
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   5910
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Programmes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   165
      Width           =   3135
   End
End
Attribute VB_Name = "frmProgrammes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private selProg As HRCORE.Programme
Private pProgs As HRCORE.Programmes

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    frmMain2.PositionTheFormWithoutEmpList Me
    
    Set pProgs = New HRCORE.Programmes
    
    
    pProgs.GetActiveProgrammes
    
    PopulateProgrammes pProgs
    
    Exit Sub
    
ErrorHandler:
    
End Sub


Private Sub PopulateProgrammes(ByVal TheProgs As HRCORE.Programmes)
    Dim i As Long
    Dim itemX As ListItem
    
    On Error GoTo ErrorHandler
    lvwProgrammes.ListItems.clear
    
    If Not (TheProgs Is Nothing) Then
        For i = 1 To TheProgs.count
            Set itemX = lvwProgrammes.ListItems.Add(, , TheProgs.Item(i).ProgrammeCode)
            itemX.SubItems(1) = TheProgs.Item(i).ProgrammeName
            itemX.SubItems(2) = TheProgs.Item(i).Donor
            itemX.SubItems(3) = TheProgs.Item(i).FundCode
            itemX.SubItems(4) = TheProgs.Item(i).ExpectedDuration
            itemX.SubItems(5) = Format(TheProgs.Item(i).StartDate, "dd-MMM-yyyy")
            itemX.SubItems(6) = Format(TheProgs.Item(i).EndDate, "dd-MMM-yyyy")
        Next i
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while Populating the Programmes" & vbNewLine & Err.Description, vbExclamation, TITLES
End Sub
