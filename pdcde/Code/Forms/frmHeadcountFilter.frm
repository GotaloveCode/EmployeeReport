VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmHeadcountFilter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Report Filter"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7950
   HelpContextID   =   1940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkExclude 
      Caption         =   "Override Filters"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   6120
      Width           =   2895
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Employees"
      TabPicture(0)   =   "frmHeadcountFilter.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Other Filters"
      TabPicture(1)   =   "frmHeadcountFilter.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   5295
         Left            =   -74880
         TabIndex        =   23
         Top             =   360
         Width           =   7455
         Begin VB.Frame fraChecklist 
            Height          =   2295
            Left            =   1440
            TabIndex        =   24
            Top             =   960
            Visible         =   0   'False
            Width           =   3495
            Begin VB.CheckBox chkNHIF 
               Caption         =   "Has NHIF Number"
               Height          =   195
               Left            =   120
               TabIndex        =   31
               Top             =   240
               Width           =   2655
            End
            Begin VB.CheckBox chkNSSF 
               Caption         =   "Has NSSF Number"
               Height          =   195
               Left            =   120
               TabIndex        =   30
               Top             =   480
               Width           =   2655
            End
            Begin VB.CheckBox chkPIN 
               Caption         =   "Has PIN Number"
               Height          =   195
               Left            =   120
               TabIndex        =   29
               Top             =   720
               Width           =   2655
            End
            Begin VB.CheckBox chkAppForm 
               Caption         =   "Has filled Application Form"
               Height          =   195
               Left            =   120
               TabIndex        =   28
               Top             =   960
               Width           =   2655
            End
            Begin VB.CheckBox chkCV 
               Caption         =   "Has brought copy of ID"
               Height          =   195
               Left            =   120
               TabIndex        =   27
               Top             =   1225
               Width           =   2655
            End
            Begin VB.CheckBox chkID 
               Caption         =   "Has brought copy of CV"
               Height          =   195
               Left            =   120
               TabIndex        =   26
               Top             =   1520
               Width           =   2655
            End
            Begin VB.CheckBox chkHandBook 
               Caption         =   "Has been given HandBook"
               Height          =   195
               Left            =   120
               TabIndex        =   25
               Top             =   1760
               Width           =   2655
            End
         End
         Begin MSComctlLib.ListView lvwEmployees 
            Height          =   4815
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   8493
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Employee Code"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Employee Names"
               Object.Width           =   5106
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5295
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   7455
         Begin VB.Frame Frame2 
            Caption         =   "Date Range"
            Height          =   1695
            Left            =   3720
            TabIndex        =   16
            Top             =   2400
            Width           =   3495
            Begin VB.ComboBox cboCriteria 
               Height          =   315
               Left            =   1080
               Style           =   2  'Dropdown List
               TabIndex        =   21
               Top             =   240
               Width           =   2295
            End
            Begin MSComCtl2.DTPicker dtFrom 
               Height          =   375
               Left            =   1080
               TabIndex        =   17
               Top             =   720
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   661
               _Version        =   393216
               CustomFormat    =   "dd-MMM-yyyy"
               Format          =   61341699
               CurrentDate     =   39197
            End
            Begin MSComCtl2.DTPicker dtTo 
               Height          =   375
               Left            =   1080
               TabIndex        =   18
               Top             =   1200
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   661
               _Version        =   393216
               CustomFormat    =   "dd-MMM-yyyy"
               Format          =   61341699
               CurrentDate     =   39197
            End
            Begin VB.Label Label1 
               Caption         =   "Filter Criteria"
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   240
               Width           =   975
            End
            Begin VB.Label lblTo 
               Caption         =   "End Date:"
               Height          =   255
               Left            =   120
               TabIndex        =   20
               Top             =   1200
               Width           =   855
            End
            Begin VB.Label lblFrom 
               Caption         =   "Start Date:"
               Height          =   255
               Left            =   120
               TabIndex        =   19
               Top             =   720
               Width           =   855
            End
         End
         Begin VB.ComboBox cboGender 
            Height          =   315
            ItemData        =   "frmHeadcountFilter.frx":0038
            Left            =   4800
            List            =   "frmHeadcountFilter.frx":0045
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   4320
            Width           =   2415
         End
         Begin VB.ComboBox cboMarital 
            Height          =   315
            ItemData        =   "frmHeadcountFilter.frx":0064
            Left            =   4800
            List            =   "frmHeadcountFilter.frx":007A
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   4800
            Width           =   2415
         End
         Begin MSComctlLib.ListView lvwPositioins 
            Height          =   2655
            Left            =   120
            TabIndex        =   7
            Top             =   2520
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   4683
            View            =   3
            LabelEdit       =   1
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
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Position"
               Object.Width           =   12091
            EndProperty
         End
         Begin MSComctlLib.ListView lvwEmpTerms 
            Height          =   1815
            Left            =   3720
            TabIndex        =   8
            Top             =   360
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   3201
            View            =   3
            LabelEdit       =   1
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
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Term of Employment"
               Object.Width           =   12303
            EndProperty
         End
         Begin MSComctlLib.ListView lvwDepartments 
            Height          =   1815
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   3201
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Department name"
               Object.Width           =   12091
            EndProperty
         End
         Begin VB.Label lblDepartments 
            Caption         =   "Department(s)"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label lblPositions 
            Caption         =   "Designation(s)"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   2280
            Width           =   1935
         End
         Begin VB.Label lblEmpTerms 
            Caption         =   "EmploymentTerm(s)"
            Height          =   255
            Left            =   3720
            TabIndex        =   12
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label lblGender 
            Caption         =   "Gender"
            Height          =   255
            Left            =   3720
            TabIndex        =   11
            Top             =   4350
            Width           =   855
         End
         Begin VB.Label lblMarital 
            Caption         =   "Marital Status"
            Height          =   255
            Left            =   3720
            TabIndex        =   10
            Top             =   4830
            Width           =   1215
         End
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   6120
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6360
      TabIndex        =   1
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   4800
      TabIndex        =   0
      Top             =   6720
      Width           =   1335
   End
End
Attribute VB_Name = "frmHeadcountFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private OUs As HRCORE.OrganizationUnits
Private empTerms As HRCORE.EmploymentTerms
Private jpos As HRCORE.JobPositions
Private Department, EmploymentTerms, Positions, mySQL, Employees, DOB, DOE, DOD As String
Private emp As HRCORE.Employees
Private Const lvwEmpIntHeight As Integer = 7455
Private Const lvwEmpComHeight As Integer = 5000

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdOK_Click()
'    Dim Report As CRAXDDRT.Report
    Dim mySQL As String

    On Error GoTo ErrHandler
    mySQL = ""
     mySQL = getFilter(ReportType, ReportSchemaName)

    R.FormulaSyntax = 0 ' Use crCrystalSyntaxFormula value
    ' SET the record and group selection formula
    
    R.RecordSelectionFormula = mySQL
    
    Employees = ""
    Department = ""
    Positions = ""
    EmploymentTerms = ""
    ReportType = ""
    Me.lblFrom.Tag = ""
    Me.lblTo.Tag = ""
    ChangeHeight = False
    ReportHeading = False
    Me.Hide
    ShowReport R 'display the report
    
    Unload Me
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description
End Sub

Private Function getFilter(Report As String, SchemaName As String) As String
    Dim sqlFilter As String
        sqlFilter = ""
        
    Select Case Report
        Case "Disengaged"
            sqlFilter = "{" & SchemaName & ".Disengaged}=true "
            If Me.chkExclude.value = vbUnchecked Then
                If Me.cboCriteria.ListIndex <> -1 And (Me.lblFrom.Tag <> "" Or Me.lblTo.Tag <> "") Then
                    Select Case Me.cboCriteria.Text
                        Case DOB
                             sqlFilter = sqlFilter & " AND {" & SchemaName & ".DOB} In DateTime " & Format(dtFrom.value, "(yyyy, mm, dd, hh, nn, ss)") & " to DateTime " & Format(dtTo.value, "(yyyy, mm, dd, hh, nn, ss)") & ""
                        Case DOE
                            sqlFilter = sqlFilter & " AND {" & SchemaName & ".DEmployed} In DateTime " & Format(dtFrom.value, "(yyyy, mm, dd, hh, nn, ss)") & " to DateTime " & Format(dtTo.value, "(yyyy, mm, dd, hh, nn, ss)") & ""
                        Case DOD
                            sqlFilter = sqlFilter & " AND {" & SchemaName & ".DLeft} In DateTime " & Format(dtFrom.value, "(yyyy, mm, dd, hh, nn, ss)") & " to DateTime " & Format(dtTo.value, "(yyyy, mm, dd, hh, nn, ss)") & ""
                    End Select
                End If
            End If
            
        Case "Normal"
            sqlFilter = "{" & SchemaName & ".Disengaged}=False "
            If Me.chkExclude.value = vbUnchecked Then
                If Me.cboCriteria.ListIndex <> -1 And (Me.lblFrom.Tag <> "" Or Me.lblTo.Tag <> "") Then
                    Select Case Me.cboCriteria.Text
                        Case DOB
                            sqlFilter = sqlFilter & " AND {" & SchemaName & ".DOB} In DateTime " & Format(dtFrom.value, "(yyyy, mm, dd, hh, nn, ss)") & " to DateTime " & Format(dtTo.value, "(yyyy, mm, dd, hh, nn, ss)") & ""
                        Case DOE
                            sqlFilter = sqlFilter & " AND {" & SchemaName & ".DEmployed} In DateTime " & Format(dtFrom.value, "(yyyy, mm, dd, hh, nn, ss)") & " to DateTime " & Format(dtTo.value, "(yyyy, mm, dd, hh, nn, ss)") & ""
                    End Select
                End If
            End If
    End Select
    
    If Me.chkExclude.value = vbUnchecked Then
        If Employees <> "" Then sqlFilter = sqlFilter & " AND {" & SchemaName & ".Employee_ID} in [ " & Employees & "]"
        
        If Department <> "" Then sqlFilter = sqlFilter & " AND {" & SchemaName & ".OrganizationUnitName} in [ " & Department & "]"
        
        If Positions <> "" Then sqlFilter = sqlFilter & " AND {" & SchemaName & ".Desig} in [ " & Positions & " ]"
        
        If EmploymentTerms <> "" Then sqlFilter = sqlFilter & " AND  {" & SchemaName & ".Terms} in [ " & EmploymentTerms & "]"
            
        If Me.cboGender.ListIndex <> -1 Then sqlFilter = sqlFilter & " AND {" & SchemaName & ".Gender}= '" & Me.cboGender.Text & "'"
    
        If Me.cboMarital.ListIndex <> -1 Then sqlFilter = sqlFilter & " AND {" & SchemaName & ".Marital_Status}= '" & Me.cboMarital.Text & "'"
        
        '--------------------Pre-employment check list filters------------------------------------------------------------------------------------------
        If ChangeHeight = True Then
        
            If Me.chkAppForm.value = vbChecked Then sqlFilter = sqlFilter & " AND {" & SchemaName & ".ApplicationForm}=true"
            
            If Me.chkCV.value = vbChecked Then sqlFilter = sqlFilter & " AND {" & SchemaName & ".CVCopy}=true"
            
            If Me.chkHandBook.value = vbChecked Then sqlFilter = sqlFilter & " AND {" & SchemaName & ".EmployeeHandBook}=true"
            
            If Me.chkID.value = vbChecked Then sqlFilter = sqlFilter & " AND {" & SchemaName & ".IDCardCopy}=true"
            
            If Me.chkNHIF.value = vbChecked Then sqlFilter = sqlFilter & " AND {" & SchemaName & ".NHIFCopy}=true"
            
            If Me.chkNSSF.value = vbChecked Then sqlFilter = sqlFilter & " AND {" & SchemaName & ".NSSFCopy}=true"
            
            If Me.chkPIN.value = vbChecked Then sqlFilter = sqlFilter & " AND {" & SchemaName & ".PINCopy}=true"
            
          End If
          
    End If
    
    getFilter = sqlFilter
End Function

Private Sub dtFrom_Change()
    Me.Text1.Text = Year(dtFrom.value)
    Me.lblFrom.Tag = dtFrom.value
End Sub



Private Sub dtTo_Change()
    Me.Text1.Text = Year(dtTo.value)
    Me.lblTo.Tag = dtTo.value
End Sub

Private Sub Form_Load()
    frmMain2.MousePointer = 11
    Me.Top = 0
    Me.Left = (frmMain2.Width - Me.Width) / 2
    
    Set OUs = New HRCORE.OrganizationUnits
    Set empTerms = New HRCORE.EmploymentTerms
    Set jpos = New HRCORE.JobPositions
    Set emp = New HRCORE.Employees
    emp.GetAllEmployees
    OUs.GetAllOrganizationUnits
    empTerms.GetAllEmploymentTerms
    jpos.GetAllJobPositions
    
    Call LoadEmployees
    Call PopulateOU
    Call PopulateEmpTerms
    Call PopulateJpos
    Call LoadDateFilterCriteria
    frmMain2.MousePointer = 0
    Me.cboGender.ListIndex = -1
    Me.cboMarital.ListIndex = -1
    Me.cboCriteria.ListIndex = -1
    Me.SSTab1.TabsPerRow = 2
    
    If ChangeHeight = True Then
        'Me.lvwEmployees.Height = lvwEmpComHeight
        fraChecklist.Visible = True
    Else
        'Me.lvwEmployees.Height = lvwEmpIntHeight
        fraChecklist.Visible = False
    End If
    
    If ReportHeading = True Then
        Me.SSTab1.TabsPerRow = 1
        Me.SSTab1.TabVisible(0) = False
    Else
'        Me.SSTab1.TabsPerRow = 2
'        Me.SSTab1.TabVisible(0) = true
    End If
End Sub

Private Sub LoadDateFilterCriteria()

    DOB = "Date of Birth"
    DOE = "Date of Employment"
    DOD = "Date of Disengagement"
    Me.cboCriteria.clear
    
    Select Case ReportType
        Case "Disengaged"
            Me.cboCriteria.AddItem DOB
            Me.cboCriteria.AddItem DOE
            Me.cboCriteria.AddItem DOD
        Case "Normal"
            Me.cboCriteria.AddItem DOB
            Me.cboCriteria.AddItem DOE
    End Select
    
End Sub

Private Sub PopulateOU()
    Dim itemX As ListItem
    Dim i As Long
    
    For i = 1 To OUs.count
        Set itemX = Me.lvwDepartments.ListItems.Add(, , OUs.Item(i).OrganizationUnitName)
        itemX.Tag = OUs.Item(i).OrganizationUnitID
    Next i
    
End Sub

Private Sub PopulateEmpTerms()
    Dim itemX As ListItem
    Dim i As Long
    
    For i = 1 To empTerms.count
        Set itemX = Me.lvwEmpTerms.ListItems.Add(, , empTerms.Item(i).EmpTermName)
        itemX.Tag = empTerms.Item(i).EmpTermID
        
    Next i
    
End Sub

Private Sub PopulateJpos()
    Dim itemX As ListItem
    Dim i As Long
    For i = 1 To jpos.count
        Set itemX = Me.lvwPositioins.ListItems.Add(, , jpos.Item(i).PositionName)
        itemX.Tag = jpos.Item(i).PositionID
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set OUs = Nothing
    Set empTerms = Nothing
    Set jpos = Nothing
End Sub


Private Sub lvwDepartments_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim selItem As HRCORE.OrganizationUnit
    Dim i, total As Long
    On Error GoTo ErrHandler
    Department = ""
    
    For i = 1 To lvwDepartments.ListItems.count
        With lvwDepartments.ListItems.Item(i)
            If .Checked Then
                If Department = "" Then
                    Department = "'" & .Text & "'"
                Else
                    Department = Department & ",'" & .Text & "'"
                End If
            End If
        End With
    Next i
    
    'If Department <> "" Then Department = Department & " ]"
    Text1.Text = Department
    Exit Sub
ErrHandler:
    MsgBox Err.Description
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
    Dim i As Long
    On Error GoTo ErrHandler
    Employees = ""
    
    For i = 1 To lvwEmployees.ListItems.count
        With lvwEmployees.ListItems.Item(i)
            If .Checked Then
                If Employees = "" Then
                    Employees = .Tag
                Else
                    Employees = Employees & "," & .Tag
                End If
            End If
        End With
    Next i
    Text1.Text = Employees
    
    Exit Sub
ErrHandler:
    'MsgBox Err.Description, vbInformation, "error"
End Sub

Private Sub lvwEmpTerms_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim selItem As HRCORE.EmploymentTerm
    Dim i, total As Long
    On Error GoTo ErrHandler
    Dim Terms As String
    
    EmploymentTerms = ""
    For i = 1 To lvwEmpTerms.ListItems.count
        With lvwEmpTerms.ListItems.Item(i)
            If .Checked Then
                If EmploymentTerms = "" Then
                    EmploymentTerms = "'" & .Text & "'"
                Else
                    EmploymentTerms = EmploymentTerms & ",'" & .Text & "'"
                End If
            End If
        End With
    Next i
    Text1.Text = EmploymentTerms
    Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub


Private Sub lvwPositioins_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim selItem As HRCORE.JobPosition
    Dim i As Long
    On Error GoTo ErrHandler
    Positions = ""
    For i = 1 To lvwPositioins.ListItems.count
        With lvwPositioins.ListItems.Item(i)
            If .Checked Then
               If Positions = "" Then
                    Positions = "'" & .Text & "'"
                Else
                    Positions = Positions & ",'" & .Text & "'"
                End If
            End If
        End With
    Next i
    Text1.Text = Positions
    Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub LoadEmployees()
    Dim itemX As ListItem
    Dim i As Integer
    If emp Is Nothing Then Exit Sub
    lvwEmployees.ListItems.clear
    
    For i = 1 To emp.count
        If ReportType = "Normal" Then
            Set itemX = lvwEmployees.ListItems.Add(, , emp.Item(i).empcode)
            itemX.SubItems(1) = emp.Item(i).SurName & "  " & emp.Item(i).OtherNames
            itemX.Tag = emp.Item(i).EmployeeID
        Else
            If emp.Item(i).IsDisengaged = True Then
                Set itemX = lvwEmployees.ListItems.Add(, , emp.Item(i).empcode)
                itemX.SubItems(1) = emp.Item(i).SurName & "  " & emp.Item(i).OtherNames
                itemX.Tag = emp.Item(i).EmployeeID
            End If
        End If
    Next
End Sub
