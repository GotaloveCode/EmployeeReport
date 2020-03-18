VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOUEmployees 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employees"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   11115
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4410
      Top             =   2970
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOUEmployees.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOUEmployees.frx":0452
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraEmployees 
      Caption         =   "Employees"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7910
      Left            =   3120
      TabIndex        =   1
      Top             =   90
      Width           =   7935
      Begin MSComctlLib.ListView lvwEmployees 
         Height          =   7500
         Left            =   90
         TabIndex        =   3
         ToolTipText     =   "Double Click an Employee to View more Details"
         Top             =   240
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   13229
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Emp Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Surname"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Other Names"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Gender"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Marital Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Designation"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Emp. Terms"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame fraOUnits 
      Caption         =   "Organization Units"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7910
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   2985
      Begin VB.CheckBox chkInclEOUV 
         Caption         =   "Include Visible Employees"
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
         Left            =   90
         TabIndex        =   5
         Top             =   630
         Width           =   2355
      End
      Begin VB.CheckBox chkInclChildren 
         Caption         =   "Include Children Of Selected OU"
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
         Left            =   90
         TabIndex        =   4
         ToolTipText     =   "e.g. Include Sub-Departments under the Selected Department"
         Top             =   270
         Width           =   2625
      End
      Begin MSComctlLib.TreeView tvwOUnits 
         Height          =   6800
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   11986
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   441
         LabelEdit       =   1
         Style           =   6
         FullRowSelect   =   -1  'True
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
      End
   End
End
Attribute VB_Name = "frmOUEmployees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private company As HRCORE.CompanyDetails
Private outs As HRCORE.OrganizationUnitTypes
Private TopLevelOUnits As HRCORE.OrganizationUnits
Private OUnits As HRCORE.OrganizationUnits
Private SelEmps As HRCORE.Employees
Private IncludeChildren As Boolean
Private IncludeVisibleEmployees As Boolean
Private CompanyNodeSelected As Boolean  'indicates that Company is selected
Private ClickedOU As HRCORE.OrganizationUnit
Private SelEmployee As HRCORE.Employee      'will hold the selected employee
Private myInternalPeriod As Period
Private vperiod As Period


Private Sub chkInclChildren_Click()
    If chkInclChildren.value = vbChecked Then
        IncludeChildren = True
    Else
        IncludeChildren = False
    End If
End Sub

Private Sub chkInclEOUV_Click()
    If chkInclEOUV.value = vbChecked Then
        IncludeVisibleEmployees = True
    Else
        IncludeVisibleEmployees = False
    End If
End Sub

Private Sub Form_Activate()
    If tvwOUnits.Nodes.count > 0 Then
        tvwOUnits.Nodes(1).Selected = True
        Call tvwOUnits_NodeClick(tvwOUnits.Nodes(1))
    End If
End Sub

Private Sub Form_Load()

    On Error GoTo ErrorHandler
    
    'Set company = New HRCORE.CompanyDetails
    Set outs = New HRCORE.OrganizationUnitTypes
    Set OUnits = New HRCORE.OrganizationUnits
    Set myInternalPeriod = New Period
    
    myInternalPeriod.GetAllPeriods
    myInternalPeriod.GetOpenPeriod
    Set vperiod = New Period
    Set vperiod = myInternalPeriod.GetOpenPeriod
    'company.LoadCompanyDetails
    
    'load the OUTypes
    outs.GetAllOUTypes
    
    'now load the organization units
    LoadOrganizationUnits
    
    frmMain2.PositionTheFormWithoutEmpList Me
    
    'reset the flags to False
    IncludeVisibleEmployees = False
    IncludeChildren = False
    CompanyNodeSelected = False
    
    Exit Sub
    
ErrorHandler:
    MsgBox "A slight error has occurred" & vbNewLine & err.Description, vbInformation, HR_TITLE
End Sub

Private Sub LoadOrganizationUnits()
    Dim myOU As OrganizationUnit
    Dim myNode As Node
    Dim i As Long
    
    On Error GoTo ErrorHandler
    'clear lists
    tvwOUnits.Nodes.Clear
    
    'first load the Company Name
    Set myNode = Me.tvwOUnits.Nodes.add(, , "HRCOMPANY", companyDetail.CompanyName)
    myNode.Tag = 0
    myNode.Bold = True
    'get all the organization units
    OUnits.GetAllOrganizationUnits
    
    'now get the topmost OUs
    Set TopLevelOUnits = OUnits.GetOrganizationUnitsOfTopmostLevel()
    If Not (TopLevelOUnits Is Nothing) Then
        For i = 1 To TopLevelOUnits.count
            Set myOU = TopLevelOUnits.Item(i)
            'add the OU
            Set myNode = Me.tvwOUnits.Nodes.add("HRCOMPANY", tvwChild, "OU:" & myOU.OrganizationUnitID, myOU.OrganizationUnitName)
            myNode.Tag = myOU.OrganizationUnitID
            myNode.EnsureVisible
            
            'now recursively add the children
            AddChildOUsRecursively myOU
        Next i
                
    End If
    Exit Sub
ErrorHandler:
    MsgBox "An error has occurred while populating the Organization Units" & _
    vbNewLine & err.Description, vbInformation, HR_TITLE
End Sub


Private Sub AddChildOUsRecursively(ByVal theOU As HRCORE.OrganizationUnit)
    'this is a recursive function that populates child ous
    Dim ChildNode As Node
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    If Not (theOU Is Nothing) Then
        For i = 1 To theOU.Children.count
            Set ChildNode = tvwOUnits.Nodes.add("OU:" & theOU.OrganizationUnitID, tvwChild, "OU:" & theOU.Children.Item(i).OrganizationUnitID, theOU.Children.Item(i).OrganizationUnitName)
            ChildNode.Tag = theOU.Children.Item(i).OrganizationUnitID
            ChildNode.EnsureVisible
            'recursively load the children
            AddChildOUsRecursively theOU.Children.Item(i)
        Next i
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred" & vbNewLine & err.Description, vbInformation, HR_TITLE
        
End Sub

Private Sub Form_Resize()
    oSmart.FResize Me
    lvwEmployees.Move lvwEmployees.Left, lvwEmployees.Top, lvwEmployees.Width, (tvwMainheight - 500)
    tvwOUnits.Move tvwOUnits.Left, tvwOUnits.Top, tvwOUnits.Width, (lvwEmployees.Height - (tvwOUnits.Top - 200))
    
End Sub


Private Sub lvwEmployees_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static PrevSortKey As Long
    Static PrevSortOrder As ListSortOrderConstants
    Dim colX As ColumnHeader
    
    On Error GoTo ErrorHandler
    
    With lvwEmployees
        .SortKey = ColumnHeader.Index - 1
        If .SortKey = PrevSortKey Then
            If PrevSortOrder = lvwAscending Then
                .SortOrder = lvwDescending
                'ColumnHeader.Icon = 2
            Else
                .SortOrder = lvwAscending
                'ColumnHeader.Icon = 1
            End If
        Else
            .SortOrder = lvwAscending
            'ColumnHeader.Icon = 1
        End If
        .Sorted = True
        
        PrevSortKey = .SortKey
        PrevSortOrder = .SortOrder
    End With
    
    Exit Sub
    
ErrorHandler:
End Sub

Private Sub lvwEmployees_DblClick()
    Dim i As Long
    Dim ItemX As ListItem
    
    On Error GoTo ErrorHandler
    
    If Not (SelEmployee Is Nothing) Then
        Me.MousePointer = vbHourglass
        
        Set SelectedEmployee = SelEmployee
        frmEmployee.Show , frmMain2
        
        'call method to display records
        frmEmployee.DisplayRecords
        
        frmMain2.fracmd.Visible = True
        frmMain2.fraEmployees.Visible = True
        Set TheLoadedForm = frmEmployee
        
        'highlight the selected employee
        For Each ItemX In frmMain2.lvwEmp.ListItems
            If IsNumeric(ItemX.Tag) Then
                If CLng(ItemX.Tag) = SelectedEmployee.EmployeeID Then
                    ItemX.Selected = True
                    ItemX.EnsureVisible
                End If
            End If
        Next ItemX
        Me.MousePointer = vbDefault
        
        'unload the current form
        Unload Me
        
    End If
       
       
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred" & vbNewLine & err.Description, vbInformation, TITLES
End Sub

Private Sub lvwEmployees_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo ErrorHandler
    
    frmMain2.txtDetails.Caption = ""
    Set SelEmployee = Nothing
    
    If Not (AllEmployees Is Nothing) Then
        Set SelEmployee = AllEmployees.FindEmployee(CLng(Item.Tag))
        If Not (SelEmployee Is Nothing) Then
            'display the Employee Info
            frmMain2.txtDetails.Caption = "EmpCode: " & SelEmployee.EmpCode & " | Name: " & SelEmployee.SurName & "" & " " & SelEmployee.OtherNames & "" & " " & vbCrLf & _
             "" & "ID No:" & " " & SelEmployee.IdNo & " | Date Employed:" & " " & Format(SelEmployee.DateOfEmployment, "dd-MMM-yyyy")
        Else
            frmMain2.txtDetails.Caption = ""
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred" & vbNewLine & err.Description, vbInformation, TITLES
End Sub

Private Sub lvwEmployees_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim ItemX As ListItem
    Dim strInfo As String
    Dim CurrEmp As HRCORE.Employee
    
    On Error GoTo ErrorHandler
    
    Set ItemX = lvwEmployees.HitTest(X, y)
    strInfo = ""
    If Not (ItemX Is Nothing) Then
        If IsNumeric(ItemX.Tag) Then
            Set CurrEmp = AllEmployees.FindEmployee(CLng(ItemX.Tag))
            If Not (CurrEmp Is Nothing) Then
                strInfo = UCase(CurrEmp.SurName) & "'s Actual Department is: " & UCase(CurrEmp.OrganizationUnit.OrganizationUnitName)
                lvwEmployees.ToolTipText = strInfo
            End If
        End If
    Else
        lvwEmployees.ToolTipText = "Double Click on an Employee to view all the Details"
    End If
    
    Exit Sub
    
ErrorHandler:
End Sub

Private Sub tvwOUnits_NodeClick(ByVal Node As MSComctlLib.Node)
    
    On Error GoTo ErrorHandler
    Me.MousePointer = vbHourglass
    
    'clear the employees list
    Me.lvwEmployees.ListItems.Clear
    
    'destroy the selected employee object
    Set SelEmployee = Nothing
    
    'clear the details display
    frmMain2.txtDetails.Caption = ""
    
    If Node.Key = "HRCOMPANY" Then
        'display all the employees
        'indicate that companynode was selected
        CompanyNodeSelected = True
        DisplayEmployees AllEmployees
        CompanyNodeSelected = False
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    Set ClickedOU = OUnits.FindOrganizationUnit(CLng(Node.Tag))
    If Not (ClickedOU Is Nothing) Then
        Set SelEmps = AllEmployees.FilterEmployeesByOU(AllEmployees, ClickedOU, IncludeChildren, IncludeVisibleEmployees)
        DisplayEmployees SelEmps
    End If
    
    Me.MousePointer = vbDefault
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred" & vbNewLine & err.Description, vbInformation, TITLES
    Me.MousePointer = vbDefault
    
End Sub

Private Sub DisplayEmployees(ByVal TheEmployees As HRCORE.Employees)
    Dim i As Long
    Dim MyEmp As HRCORE.Employee
    Dim ItemX As ListItem
    Dim SubItemX As ListSubItem
    
    Dim IsVisibleEmp As Boolean
    
    On Error GoTo ErrorHandler
    'first clear
    Me.lvwEmployees.ListItems.Clear
    
    If TheEmployees Is Nothing Then Exit Sub
        
    For i = 1 To TheEmployees.count
        Set MyEmp = TheEmployees.Item(i)
        If Not (vperiod Is Nothing) Then
        If MyEmp.IsDisengaged = True And MyEmp.DateOfDisengagement < vperiod.OpenDate Then
        GoTo k
        End If
        End If
        
        
        Set ItemX = Me.lvwEmployees.ListItems.add(, , MyEmp.EmpCode)
        ItemX.SubItems(1) = MyEmp.SurName
        ItemX.SubItems(2) = MyEmp.OtherNames
        ItemX.SubItems(3) = MyEmp.GenderStr
        ItemX.SubItems(4) = MyEmp.MaritalStatusStr
        ItemX.SubItems(5) = MyEmp.position.PositionName
        ItemX.SubItems(6) = MyEmp.EmploymentTerm.EmpTermName
        ItemX.Tag = MyEmp.EmployeeID
        If Not CompanyNodeSelected Then
        'if Company node was selected, no employee should be highlighted
            If IsVisibleEmployee(MyEmp) Then
                ItemX.ForeColor = vbBlue
                For Each SubItemX In ItemX.ListSubItems
                    SubItemX.ForeColor = vbBlue
                Next SubItemX
            End If
        End If
k:
    Next i
    
    'display the Employee Count
    frmMain2.lblEmpCount.Caption = lvwEmployees.ListItems.count
    
    'highlight the first employee if s/he exists
    If lvwEmployees.ListItems.count > 0 Then
        lvwEmployees.ListItems(1).Selected = True
        Call lvwEmployees_ItemClick(lvwEmployees.ListItems(1))
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred--" & vbNewLine & err.Description, vbInformation, TITLES
End Sub

Private Function IsVisibleEmployee(ByVal TheEmployee As HRCORE.Employee) As Boolean
    Dim OUs As HRCORE.OrganizationUnits
    Dim i As Long
    Dim EmployeeExistsInOneOfTheOUs As Boolean
    
    On Error GoTo ErrorHandler
    
    If Not (ClickedOU Is Nothing) Then
        If IncludeChildren Then
            
            'first get the children
            Set OUs = OUnits.GetChildrenRecursively(ClickedOU)
            
            'then add the clicked ou to the collection
            If Not (OUs Is Nothing) Then
                OUs.add ClickedOU
            Else
                Set OUs = New HRCORE.OrganizationUnits
                OUs.add ClickedOU
            End If
            
            For i = 1 To OUs.count
                If TheEmployee.OrganizationUnit.OrganizationUnitID = OUs.Item(i).OrganizationUnitID Then
                    EmployeeExistsInOneOfTheOUs = True
                    Exit For
                End If
            Next i
            
            If Not EmployeeExistsInOneOfTheOUs Then
                IsVisibleEmployee = True
            Else
                IsVisibleEmployee = False
            End If
        Else
            If TheEmployee.OrganizationUnit.OrganizationUnitID <> ClickedOU.OrganizationUnitID Then
                IsVisibleEmployee = True
            Else
                IsVisibleEmployee = False
            End If
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error has occurred" & vbNewLine & err.Description, vbInformation, TITLES
End Function
