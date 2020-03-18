VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmReengageMent 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   12360
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraReengage 
      Appearance      =   0  'Flat
      Caption         =   "Disengaged employees in archive"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7860
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12135
      Begin VB.CommandButton cmdcancel 
         Caption         =   "Exit"
         Height          =   375
         Left            =   9600
         TabIndex        =   42
         Top             =   7320
         Width           =   1095
      End
      Begin VB.CommandButton cmdEngage 
         Caption         =   "Engage"
         Height          =   375
         Left            =   3240
         TabIndex        =   41
         Top             =   7320
         Width           =   1095
      End
      Begin VB.Frame fradetails 
         Height          =   7455
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   3015
         Begin VB.CommandButton cmdView 
            Caption         =   "View Reengagement History"
            Height          =   375
            Left            =   120
            TabIndex        =   49
            Top             =   6960
            Visible         =   0   'False
            Width           =   2775
         End
         Begin MSComctlLib.ListView lvwEmployees 
            Height          =   3735
            Left            =   120
            TabIndex        =   31
            Top             =   120
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   6588
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Employee Code"
               Object.Width           =   2284
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Employee Names"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Disengagement Reason"
               Object.Width           =   2822
            EndProperty
         End
         Begin MSComctlLib.ListView lvwEmpHistory 
            Height          =   2655
            Left            =   120
            TabIndex        =   48
            Top             =   4200
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   4683
            View            =   3
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Employee Code"
               Object.Width           =   2284
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Employee Names"
               Object.Width           =   5292
            EndProperty
         End
         Begin VB.Label Label13 
            Caption         =   "Reengaged Employees"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   3960
            Width           =   2775
         End
      End
      Begin VB.Frame fraEmployee 
         Height          =   6975
         Left            =   3240
         TabIndex        =   1
         Top             =   240
         Width           =   8775
         Begin VB.TextBox txtDisEngagementReferenceNumber 
            Height          =   285
            Left            =   1800
            TabIndex        =   51
            Top             =   6480
            Width           =   2295
         End
         Begin VB.CheckBox chkStaffNumber 
            Caption         =   "Change staff Number"
            Height          =   255
            Left            =   4440
            TabIndex        =   47
            Top             =   6495
            Width           =   2655
         End
         Begin VB.Frame fraEmp 
            Height          =   3255
            Left            =   4320
            TabIndex        =   32
            Top             =   2280
            Visible         =   0   'False
            Width           =   4335
            Begin VB.ComboBox cboDepartment 
               Height          =   315
               Left            =   1920
               Style           =   2  'Dropdown List
               TabIndex        =   44
               Top             =   1722
               Width           =   2295
            End
            Begin VB.CheckBox chkForce 
               Caption         =   "Force Reengagement"
               Height          =   375
               Left            =   120
               TabIndex        =   43
               Top             =   2760
               Width           =   2775
            End
            Begin MSComCtl2.DTPicker dtReengaged 
               Height          =   315
               Left            =   1920
               TabIndex        =   40
               Top             =   2220
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   556
               _Version        =   393216
               Format          =   63963137
               CurrentDate     =   39151
            End
            Begin VB.ComboBox cboGrade 
               Height          =   315
               Left            =   1920
               Style           =   2  'Dropdown List
               TabIndex        =   38
               Top             =   1218
               Width           =   2295
            End
            Begin VB.ComboBox cbEmpTerm 
               Height          =   315
               Left            =   1920
               Style           =   2  'Dropdown List
               TabIndex        =   36
               Top             =   714
               Width           =   2295
            End
            Begin VB.ComboBox cboPosition 
               Height          =   315
               Left            =   1920
               Style           =   2  'Dropdown List
               TabIndex        =   34
               Top             =   210
               Width           =   2295
            End
            Begin VB.Label Label16 
               Caption         =   "Department"
               Height          =   255
               Left            =   120
               TabIndex        =   45
               Top             =   1752
               Width           =   1815
            End
            Begin VB.Label Label12 
               Caption         =   "Date of Reengagement"
               Height          =   255
               Left            =   120
               TabIndex        =   39
               Top             =   2256
               Width           =   1815
            End
            Begin VB.Label Label11 
               Caption         =   "Grade"
               Height          =   255
               Left            =   120
               TabIndex        =   37
               Top             =   1248
               Width           =   1815
            End
            Begin VB.Label Label10 
               Caption         =   "Employment Term"
               Height          =   255
               Left            =   120
               TabIndex        =   35
               Top             =   744
               Width           =   1815
            End
            Begin VB.Label Label9 
               Caption         =   "Designition"
               Height          =   255
               Left            =   120
               TabIndex        =   33
               Top             =   240
               Width           =   1815
            End
         End
         Begin VB.TextBox txtdate 
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   6000
            Width           =   2295
         End
         Begin VB.TextBox txtReason 
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   5442
            Width           =   2295
         End
         Begin VB.TextBox txtIDNo 
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   1518
            Width           =   2295
         End
         Begin VB.TextBox txtSurname 
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   646
            Width           =   2295
         End
         Begin VB.TextBox txtONames 
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   1082
            Width           =   2295
         End
         Begin VB.TextBox txtPassport 
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   1954
            Width           =   2295
         End
         Begin VB.TextBox txtAlien 
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   2390
            Width           =   2295
         End
         Begin VB.TextBox txtEmpcode 
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox txtGender 
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   3698
            Width           =   2295
         End
         Begin VB.TextBox txtMstatus 
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   4134
            Width           =   2295
         End
         Begin VB.TextBox txtTribe 
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   4570
            Width           =   2295
         End
         Begin VB.TextBox txtReligion 
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   5006
            Width           =   2295
         End
         Begin VB.TextBox txtNationality 
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   3262
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker dtpDOB 
            Height          =   285
            Left            =   1830
            TabIndex        =   13
            Top             =   2820
            Width           =   2295
            _ExtentX        =   4048
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
            Format          =   63963139
            CurrentDate     =   37845
            MinDate         =   -36522
         End
         Begin VB.Label Label15 
            Caption         =   "Reference Number"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   6495
            Width           =   1935
         End
         Begin VB.Label lblTimes 
            Height          =   855
            Left            =   4320
            TabIndex        =   46
            Top             =   5640
            Width           =   4335
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            Caption         =   "Date Disengaged"
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
            Height          =   270
            Left            =   120
            TabIndex        =   29
            Top             =   6007
            Width           =   1365
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            Caption         =   "Disengagement reason"
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
            TabIndex        =   27
            Top             =   5457
            Width           =   1650
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
            Left            =   120
            TabIndex        =   25
            Top             =   5051
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
            Left            =   120
            TabIndex        =   24
            Top             =   3307
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
            Left            =   120
            TabIndex        =   23
            Top             =   4179
            Width           =   975
         End
         Begin VB.Label Label30 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Tribe"
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
            Top             =   4615
            Width           =   360
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
            Left            =   120
            TabIndex        =   21
            Top             =   2435
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
            Left            =   120
            TabIndex        =   20
            Top             =   1999
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
            Left            =   120
            TabIndex        =   19
            Top             =   1127
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
            Left            =   120
            TabIndex        =   18
            Top             =   691
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
            Left            =   120
            TabIndex        =   17
            Top             =   285
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
            Left            =   120
            TabIndex        =   16
            Top             =   3743
            Width           =   525
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
            Left            =   120
            TabIndex        =   15
            Top             =   2865
            Width           =   945
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
            Left            =   120
            TabIndex        =   14
            Top             =   1563
            Width           =   465
         End
         Begin VB.Image Picture1 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   1935
            Left            =   6720
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1875
         End
      End
   End
End
Attribute VB_Name = "frmReengageMent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private empid As Long
Private emp As HRCORE.Employees
Private Terms As HRCORE.EmploymentTerms
Private Positions As HRCORE.JobPositions
Private Grades As HRCORE.EmployeeCategories
Private Department As HRCORE.OrganizationUnits
Private Reengaged As ReengagedEmployees
Private SelEmployee As HRCORE.Employee
Private RSHisrory As ADODB.Recordset

Private Sub LoadCombos()
    Dim i As Long
    On Error GoTo ErrHandler
    Set Positions = New HRCORE.JobPositions
    Set Grades = New HRCORE.EmployeeCategories
    Set Department = New HRCORE.OrganizationUnits
    Set Terms = New HRCORE.EmploymentTerms
    
    Positions.GetAllJobPositions
    Grades.GetActiveEmployeeCategories
    Department.GetAllOrganizationUnits
    Terms.GetAllEmploymentTerms
    
    For i = 1 To Terms.count
        cbEmpTerm.AddItem Terms.Item(i).EmpTermName
        cbEmpTerm.ItemData(cbEmpTerm.NewIndex) = Terms.Item(i).EmpTermID
    Next i
    
    For i = 1 To Department.count
        cboDepartment.AddItem Department.Item(i).OrganizationUnitName
        cboDepartment.ItemData(cboDepartment.NewIndex) = Department.Item(i).OrganizationUnitID
    Next i
    
    For i = 1 To Grades.count
        cboGrade.AddItem Grades.Item(i).CategoryName
        cboGrade.ItemData(cboGrade.NewIndex) = Grades.Item(i).CategoryID
    Next i
    
    For i = 1 To Positions.count
        cboPosition.AddItem Positions.Item(i).PositionName
        cboPosition.ItemData(cboPosition.NewIndex) = Positions.Item(i).PositionID
    Next i
    
    Exit Sub
ErrHandler:
    MsgBox "The following error has occured when loading data: " & err.Description
End Sub

Private Sub chkStaffNumber_Click()
    If chkStaffNumber.value = vbChecked Then
        txtEmpCode.Locked = False
    ElseIf chkStaffNumber.value = vbUnchecked Then
        txtEmpCode.Locked = True
        txtEmpCode.Text = txtEmpCode.Tag
    End If
End Sub

Private Sub cmdEngage_Click()
    
    If Not currUser Is Nothing Then
        If currUser.CheckRight("Archives") <> secModify Then
            MsgBox "You don't have right to modify the record. Please liaise with the security admin"
            Exit Sub
        End If
     End If
     
     If SelEmployee Is Nothing Then
        MsgBox "Select employee to be reengaged"
        Exit Sub
     End If
     
     Select Case cmdEngage.Caption
        Case "Engage"
            fraDetails.Enabled = False
            fraEmp.Visible = True
            cmdEngage.Caption = "Save"
            cmdCancel.Caption = "Cancel"
            
        Case "Save"
            'Reengage employee
            If Me.chkStaffNumber.value = vbChecked And Me.txtEmpCode.Tag = Me.txtEmpCode.Text Then
                MsgBox "Sorry you must change the Staff number to continue", vbInformation, "Error"
                Exit Sub
            End If
            
            If chkForce.value = vbUnchecked And SelEmployee.CanBeReengaged = False Then
                MsgBox SelEmployee.SurName & "  " & SelEmployee.OtherNames & vbNewLine & "by default can not be re engage. To reengage " & IIf(SelEmployee.GenderStr = "Male", "him ", "her ") & " ensure the 'Force Reengagement' Checkbox is checked"
                Exit Sub
            End If
            
            If ReengageEmployee Then
            
                fraDetails.Enabled = True
                fraEmp.Visible = False
                cmdEngage.Caption = "Engage"
                cmdCancel.Caption = "Exit"
                
                Set SelEmployee = Nothing
                
                'refresh the listview
                Call LoadEmployees
                
                loadReengagementHistory
                Call Cleartxt
                
            'clear picture
            Me.Picture1.Picture = LoadPicture()
            End If
            
    End Select
End Sub

Private Sub cmdView_Click()
    Set r = crtReengagementHistory
    r.FormulaSyntax = crCrystalSyntaxFormula
    r.RecordSelectionFormula = "{vwReengagementHistory.EmployeeID}=" & lvwEmpHistory.SelectedItem.Tag & ""
    ShowReport r
    cmdview.Visible = False
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler
    Set emp = New HRCORE.Employees
    'load disengaged Employees
    Call LoadEmployees
    
    'position the form in the main from
    
    frmMain2.PositionTheFormWithoutEmpList Me
    
    'hind all, the command buttons in the main window
    Call HideMainWindowButtons
    'load infrmation
    LoadCombos
    
    Set Reengaged = New ReengagedEmployees
    Reengaged.GetallReengagedEmployees
    
    dtReengaged.value = Now
    loadReengagementHistory
    Exit Sub
ErrHandler:
    MsgBox "An error has occur: " & err.Description
End Sub

Private Sub loadReengagementHistory()
    On Error Resume Next
    Dim myqsl As String
    Dim ItemX As ListItem
    Me.lvwEmpHistory.ListItems.Clear
    mySQL = "Select * from VWArchivedEmployees"
    Set RSHisrory = CConnect.GetRecordSet(mySQL)
    If RSHisrory Is Nothing Then Exit Sub
    If (RSHisrory.EOF Or RSHisrory.BOF) Then Exit Sub
    
    RSHisrory.MoveFirst
    Do Until RSHisrory.EOF
        If RSHisrory!disengaged = 0 Then
            Set ItemX = Me.lvwEmpHistory.ListItems.add(, , RSHisrory!EmpCode)
            ItemX.SubItems(1) = RSHisrory!OtherNames & "  " & RSHisrory!SurName
            ItemX.Tag = RSHisrory!Employee_ID
        End If
        RSHisrory.MoveNext
    Loop
        
    
End Sub

Private Sub LoadEmployees()
    On Error GoTo ErrHandler
    Dim i As Long
        
    Dim ItemX As ListItem
    Me.lvwEmployees.ListItems.Clear
    ''emp.GetAllEmployees
    Set emp = AllEmployees
 'Display all the disengaged employees
   For i = 1 To emp.count
        If emp.Item(i).IsDisengaged = True Then
            Set ItemX = Me.lvwEmployees.ListItems.add(, , emp.Item(i).EmpCode)
            ItemX.SubItems(1) = emp.Item(i).SurName & ", " & emp.Item(i).OtherNames
            ItemX.SubItems(2) = emp.Item(i).disengagementReason
            ItemX.Tag = emp.Item(i).EmployeeID
        End If
    Next i
    If lvwEmployees.ListItems.count > 0 Then
        fraReengage.Caption = lvwEmployees.ListItems.count & " Disengaed employees"
    End If
    
    Exit Sub
ErrHandler:
    MsgBox "An error has occured when loading details of archived emplyees"
End Sub

Private Sub Form_Resize()
    fraReengage.Move fraReengage.Left, 0, fraReengage.Width, tvwMainheight - 200
End Sub

'Private Sub FindEmployee()
'    On Error GoTo errhandler
'    Dim X As ListItem
'    Set X = Me.lvwEmployees.FindItem(Me.txtSearch.Text, lvwText)
'    If Not (X Is Nothing) Then
'        X.EnsureVisible
'    End If
'    Exit Sub
'errhandler:
'
'End Sub

Private Sub Form_Unload(Cancel As Integer)
     'refresh the employee general listview
    Call frmMain2.LoadEmployeeList
    Call UnHideMainWindowButtons
End Sub

Private Sub lvwEmpHistory_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwEmpHistory
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvwEmpHistory_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmdview.Visible = True
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

Private Sub lvwEmployees_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    On errr GoTo errhanheler
    For i = 1 To lvwEmployees.ListItems.count
        lvwEmployees.ListItems.Item(i).Checked = False

    Next i
    
    lvwEmployees.SelectedItem.Checked = True

    empid = CLng(Item.Tag)
    Set SelEmployee = emp.FindEmployee(empid)
    Call DisplayRecords
    
    'Display the number of time employee has been reengaged
    Call NumberReengaged(SelEmployee)
    
    Exit Sub
errhanheler:
    MsgBox "An error has occur when displaying employee details"
End Sub

Private Sub lvwEmployees_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On errr GoTo errhanheler
    For i = 1 To lvwEmployees.ListItems.count
        lvwEmployees.ListItems.Item(i).Checked = False
    Next i
    
    lvwEmployees.SelectedItem.Checked = True

    empid = CLng(Item.Tag)
    Set SelEmployee = emp.FindEmployee(empid)
        
    Call DisplayRecords
    
    'Display the number of time employee has been reengaged
    Call NumberReengaged(SelEmployee)
    Exit Sub
errhanheler:
    MsgBox "An error has occur when displaying employee details"
End Sub

Private Sub cmdCancel_Click()
    Select Case cmdCancel.Caption
        Case "Exit"
            Unload Me
            frmMain2.cmdShowPrompts.Visible = False
        Case "Cancel"
            fraDetails.Enabled = True
            fraEmp.Visible = False
            cmdEngage.Caption = "Engage"
            cmdCancel.Caption = "Exit"
            Call Cleartxt
            'clear picture
            Me.Picture1.Picture = LoadPicture()
            lblTimes.Caption = ""
            lvwEmployees.SelectedItem.Checked = False
            Set SelEmployee = Nothing
            
    End Select
End Sub

'Private Sub txtSearch_Change()
'    If Me.txtSearch.Text <> "" Then Call FindEmployee
'End Sub

'Private Sub txtSearch_KeyPress(KeyAscii As Integer)
'    If Me.txtSearch.Text <> "" And KeyAscii = 13 Then
'        Call FindEmployee
'    End If
'End Sub

Private Function ReengageEmployee() As Boolean
    On Error GoTo ErrHandler
    
    Dim ReengageEmp As ReengagedEmployee
    Dim retVal As Boolean
    
    ReengageEmployee = False
        
    Set ReengageEmp = New ReengagedEmployee
    Set ReengageEmp.Employee = SelEmployee
        
    'These are the required parameters before re-engaging a member of staff: 12/11/2008, John
    ReengageEmp.NewStaffCode = Me.txtEmpCode.Text
    ReengageEmp.ReengagedDate = dtReengaged.value
    
    If Trim(cboDepartment.Text) <> "" Then ReengageEmp.Department.OrganizationUnitID = cboDepartment.ItemData(cboDepartment.ListIndex)
    If Trim(cbEmpTerm.Text) <> "" Then ReengageEmp.EmpTerm.EmpTermID = cbEmpTerm.ItemData(cbEmpTerm.ListIndex)
    If Trim(cboGrade.Text) <> "" Then ReengageEmp.Grade.CategoryID = cboGrade.ItemData(cboGrade.ListIndex)
    If Trim(cboPosition.Text <> "") Then ReengageEmp.position.PositionID = cboPosition.ItemData(cboPosition.ListIndex)
'        If Trim(cboGrade.Text <> "") Then
'        Dim empcat As New HRCORE.EmployeeCategory
'        If EmpCats Is Nothing Then
'        EmpCats.GetAllEmployeeCategories
'        End If
'        Set empcat = EmpCats.FindEmployeeCategoryByID(cboGrade.ItemData(cboGrade.ListIndex))
'        If Not empcat Is Nothing Then
'
'        End If
'        End If
    
    retVal = ReengageEmp.Insert
        
    If retVal Then
        ReengageEmployee = True
        currUser.AuditTrail Update, ("Has Re-engaged Employee Code: " & SelEmployee.EmpCode & "  Name: " & SelEmployee.SurName & "  " & SelEmployee.OtherNames)
    End If
        
    Exit Function
    
ErrHandler:
    MsgBox "An error has occured when re-engaging an employee"
End Function

Private Sub HideMainWindowButtons()
    Dim i As Object
    For Each i In frmMain2
        If TypeOf i Is CommandButton Then
            i.Visible = False
        End If
    Next i

End Sub

Private Sub UnHideMainWindowButtons()
    On Error Resume Next
    Dim i As Object
    For Each i In frmMain2
        If TypeOf i Is CommandButton Then
            i.Visible = True
        End If
    Next i
End Sub

Private Sub DisplayRecords()
    
    On Error GoTo ErrorHandler
    'first clear records
    
    Call Cleartxt
    
    If Not (SelEmployee Is Nothing) Then
        With SelEmployee
            Me.txtReason.Text = .disengagementReason
            Me.txtdate.Text = Format(.DateOfDisengagement, "dd-MM-yyyy")
            txtEmpCode.Text = .EmpCode
            txtEmpCode.Tag = .EmpCode
            txtSurname.Text = .SurName
            txtONames.Text = .OtherNames
            txtIDNo.Text = .IdNo
            txtGender.Text = .GenderStr
            dtpDOB.value = .DateOfBirth
            If .position.PositionName <> "" Then
                cboPosition.Text = .position.PositionName
            Else    'DEFAULT
                If cboPosition.ListCount > 0 Then cboPosition.ListIndex = 0
            End If
            
            If .OrganizationUnit.OrganizationUnitName <> "" Then
                cboDepartment.Text = .OrganizationUnit.OrganizationUnitName
            Else    'DEFAULT
                If cboDepartment.ListCount > 0 Then cboDepartment.ListIndex = 0
            End If
            
            If .EmploymentTerm.EmpTermName <> "" Then
                cbEmpTerm.Text = .EmploymentTerm.EmpTermName
            Else
                If cbEmpTerm.ListCount > 0 Then cbEmpTerm.ListIndex = 0
            End If
            If .Category.CategoryName <> "" Then
                cboGrade.Text = .Category.CategoryName
            Else
                If cboGrade.ListCount > 0 Then cboGrade.ListIndex = 0
            End If
            If .Nationality.Nationality <> "" Then
                txtNationality.Text = .Nationality.Nationality
            Else
                txtNationality.Text = "(Unspecified)"
            End If
            If .Tribe.Tribe <> "" Then
                txtTribe.Text = .Tribe.Tribe
            Else
                txtTribe.Text = "(Unspecified)"
            End If
            txtPassport.Text = .PassportNo
            txtAlien.Text = .AlienNo
            If .Religion.Religion <> "" Then
                txtReligion.Text = .Religion.Religion
            Else
                txtReligion.Text = "(Unspecified)"
            End If
            If SelEmployee.CanBeReengaged = False Then
                chkForce.value = vbUnchecked
            Else
                chkForce.value = vbGrayed
            End If
            
            txtMstatus.Text = .MaritalStatusStr
            txtDisEngagementReferenceNumber.Text = .DisEngagementReferenceNumber
            Set Picture1 = Nothing
    
            On Error Resume Next 'this handler is specific to the photos only
            Picture1.Picture = LoadPicture(App.Path & "\Photos\" & CompanyId & "-" & txtEmpCode.Text & ".jpg")
            If Picture1.Picture = 0 Then
                On Error Resume Next
                Picture1.Picture = LoadPicture(App.Path & "\Pic\Gen.jpg")
            End If
            
        End With
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "An error has occurred while Displaying Employee Info" & vbNewLine & err.Description, vbInformation, TITLES
End Sub

Private Sub Cleartxt()
    Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Then
            i.Text = ""
        End If
    Next i
    chkForce.value = vbUnchecked
    lblTimes.Caption = ""
End Sub

Private Sub NumberReengaged(emp As HRCORE.Employee)
    On Error GoTo ErrHandler
    Dim i, total As Long
    Dim empid As Long
    
    total = 0
    empid = emp.EmployeeID
    
    For i = 1 To Reengaged.count
        If Reengaged.Item(i).Employee.EmployeeID = empid Then
            total = total + 1
        End If
    Next i
    If total > 0 Then
        lblTimes.Caption = emp.SurName & " " & emp.OtherNames & vbNewLine & "Has been reengaged  " & total & IIf(total > 1, " times", " time") & vbNewLine & "in " & UCase(companyDetail.CompanyName)
    Else
        lblTimes.Caption = ""
    End If
    
    Exit Sub
ErrHandler:
    MsgBox "An error has occur : " & err.Description
End Sub
    

