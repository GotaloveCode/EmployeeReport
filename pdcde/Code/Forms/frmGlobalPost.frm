VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGlobalPost 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Pay Adjustments"
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlbox 
      Left            =   4560
      Top             =   8880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8535
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   15055
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Uniform Increament"
      TabPicture(0)   =   "frmGlobalPost.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "pBar"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(3)=   "cmdPost"
      Tab(0).Control(4)=   "cmdClose"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Importation"
      TabPicture(1)   =   "frmGlobalPost.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "txtimportfile"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdbrowse"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdimport"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdclose2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame4"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.Frame Frame4 
         Caption         =   " "
         Height          =   495
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   6735
         Begin VB.CheckBox chkhseall 
            Caption         =   "IMPORT HSE ALLOWANCE"
            Height          =   195
            Left            =   2880
            TabIndex        =   40
            Top             =   240
            Value           =   1  'Checked
            Width           =   2535
         End
         Begin VB.CheckBox chkbasic 
            Caption         =   "IMPORT BASIC PAY"
            Height          =   195
            Left            =   240
            TabIndex        =   39
            Top             =   240
            Value           =   1  'Checked
            Width           =   2295
         End
      End
      Begin VB.CommandButton cmdclose2 
         Caption         =   "Close"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   8040
         Width           =   1695
      End
      Begin VB.CommandButton cmdimport 
         Caption         =   "Import File"
         Height          =   255
         Left            =   2400
         TabIndex        =   32
         Top             =   8040
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         Caption         =   " "
         Height          =   6855
         Left            =   120
         TabIndex        =   31
         Top             =   1080
         Width           =   6735
         Begin VB.Frame fraImportLog 
            Caption         =   "Import Log:"
            Height          =   2295
            Left            =   120
            TabIndex        =   36
            Top             =   4200
            Width           =   6495
            Begin VB.TextBox txtLog 
               Height          =   1935
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   3  'Both
               TabIndex        =   37
               Top             =   240
               Width           =   6135
            End
         End
         Begin MSComctlLib.ProgressBar prgProgress 
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   6480
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ListView lvwimportedfile 
            Height          =   4095
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   7223
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "EMPCODE"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "NAMES"
               Object.Width           =   4304
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "BASIC PAY"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "HOUSE ALLOWANCE"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.CommandButton cmdbrowse 
         Caption         =   "Get the file"
         Height          =   255
         Left            =   5880
         TabIndex        =   30
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtimportfile 
         Height          =   285
         Left            =   120
         TabIndex        =   29
         Text            =   " "
         Top             =   840
         Width           =   5655
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   255
         Left            =   -70800
         TabIndex        =   28
         Top             =   8160
         Width           =   2285
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "Post Adjustments"
         Height          =   255
         Left            =   -74880
         TabIndex        =   27
         Top             =   8160
         Width           =   2285
      End
      Begin VB.Frame Frame2 
         Caption         =   "Post Increment/Decrement"
         Height          =   1815
         Left            =   -74880
         TabIndex        =   18
         Top             =   5760
         Width           =   6855
         Begin VB.ComboBox cboAdjustmentType 
            Height          =   315
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   360
            Width           =   4575
         End
         Begin VB.ComboBox cboAdjustment 
            Height          =   315
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   840
            Width           =   4575
         End
         Begin VB.TextBox txtValue 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2160
            TabIndex        =   20
            Top             =   1320
            Width           =   2285
         End
         Begin VB.ComboBox cboPay 
            Height          =   315
            Left            =   4560
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1350
            Width           =   2175
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Adjustment"
            Height          =   195
            Left            =   500
            TabIndex        =   25
            Top             =   900
            Width           =   825
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Adjustment Type"
            Height          =   195
            Left            =   500
            TabIndex        =   24
            Top             =   420
            Width           =   1230
         End
         Begin VB.Label lblVaue 
            AutoSize        =   -1  'True
            Caption         =   "Value"
            Height          =   195
            Left            =   500
            TabIndex        =   23
            Top             =   1410
            Width           =   390
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Select Employee(s)"
         Height          =   5295
         Left            =   -74880
         TabIndex        =   1
         Top             =   360
         Width           =   6855
         Begin VB.ComboBox CboPayScaleFilter 
            Height          =   315
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1140
            Width           =   2175
         End
         Begin VB.OptionButton optPayScalesFilter 
            Appearance      =   0  'Flat
            Caption         =   "Pay Scales filter"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4320
            TabIndex        =   9
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox chkSelect 
            Appearance      =   0  'Flat
            Caption         =   "Select All"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   1920
            Width           =   2055
         End
         Begin VB.ComboBox cboFilter 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   720
            Width           =   4935
         End
         Begin VB.OptionButton optDepartments 
            Appearance      =   0  'Flat
            Caption         =   "Organization Unit"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2640
            TabIndex        =   6
            Top             =   330
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton optCategory 
            Appearance      =   0  'Flat
            Caption         =   "Employee Category"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   840
            TabIndex        =   5
            Top             =   330
            Width           =   1815
         End
         Begin VB.TextBox txtUpperScale 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   4080
            TabIndex        =   4
            Top             =   1560
            Width           =   1320
         End
         Begin VB.TextBox txtLowerScale 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2160
            TabIndex        =   3
            Top             =   1560
            Width           =   1320
         End
         Begin VB.CommandButton CmdGo 
            Caption         =   "&GO"
            Height          =   375
            Left            =   5520
            TabIndex        =   2
            Top             =   1560
            Width           =   495
         End
         Begin MSComctlLib.ListView lvwEmployee 
            Height          =   2775
            Left            =   120
            TabIndex        =   11
            Top             =   2160
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   4895
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Staff No."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Employee"
               Object.Width           =   5644
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Basic Pay"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "House Allowance"
               Object.Width           =   2469
            EndProperty
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Between"
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
            Left            =   1080
            TabIndex        =   17
            Top             =   1650
            Width           =   735
         End
         Begin VB.Label lblSelected 
            AutoSize        =   -1  'True
            Caption         =   "Selected: 0"
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
            Left            =   1440
            TabIndex        =   16
            Top             =   5040
            Width           =   930
         End
         Begin VB.Label lblLoaded 
            AutoSize        =   -1  'True
            Caption         =   "Loaded:"
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
            Left            =   240
            TabIndex        =   15
            Top             =   5040
            Width           =   660
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Filter By"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   585
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "and"
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
            Left            =   3600
            TabIndex        =   13
            Top             =   1650
            Width           =   315
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Where"
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
            Left            =   1080
            TabIndex        =   12
            Top             =   1200
            Width           =   555
         End
      End
      Begin MSComctlLib.ProgressBar pBar 
         Height          =   255
         Left            =   -74880
         TabIndex        =   26
         Top             =   7800
         Visible         =   0   'False
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
   End
End
Attribute VB_Name = "frmGlobalPost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const ErrTitle = "An error has occured while posting pay adjustments" & vbNewLine
Private allPostEmps As New HRCORE.Employees
Private EmpCats As New HRCORE.EmployeeCategories
Private OUs As New HRCORE.OrganizationUnits
Dim i As Long, ItemX As ListItem
Dim eLoad As Long, eSeld As Long
Dim pSQL As String, Filter As String
Dim pRST As New adodb.Recordset

Private Sub cboFilter_Click()
    On Error GoTo errHandler
    
    lvwEmployee.ListItems.Clear
    eLoad = 0
    eSeld = 0
    
'    allPostEmps.GetAllEmployees
        
    If (optCategory.value = True) Then
        'Filter by Category
        For i = 1 To allPostEmps.count
            If (allPostEmps.Item(i).Category.CategoryID = cboFilter.ItemData(cboFilter.ListIndex) And _
            allPostEmps.Item(i).IsDisengaged = False) Then
                With lvwEmployee
                    Set ItemX = .ListItems.add(, , allPostEmps.Item(i).EmpCode)
                    ItemX.SubItems(1) = allPostEmps.Item(i).SurName & " " & allPostEmps.Item(i).OtherNames
                    ItemX.SubItems(2) = allPostEmps.Item(i).BasicPay
                    ItemX.SubItems(3) = allPostEmps.Item(i).HouseAllowance
                    ItemX.Tag = allPostEmps.Item(i).EmployeeID
                    eLoad = eLoad + 1
                End With
            End If
        Next i
        lblLoaded = "Loaded: " & eLoad
        lblSelected = "Selected: 0"
        
    ElseIf (optDepartments.value = True) Then
        'Filter By Departments
        If (cboFilter.Text = "(All Organization Units)") Then
            'All Employees
            For i = 1 To allPostEmps.count
                If (allPostEmps.Item(i).IsDisengaged = False) Then
                    With lvwEmployee
                        Set ItemX = .ListItems.add(, , allPostEmps.Item(i).EmpCode)
                        ItemX.SubItems(1) = allPostEmps.Item(i).SurName & " " & allPostEmps.Item(i).OtherNames
                        ItemX.SubItems(2) = allPostEmps.Item(i).BasicPay
                        ItemX.SubItems(3) = allPostEmps.Item(i).HouseAllowance
                        ItemX.Tag = allPostEmps.Item(i).EmployeeID
                        eLoad = eLoad + 1
                    End With
                End If
            Next i
            lblLoaded = "Loaded: " & eLoad
            lblSelected = "Selected: 0"
        Else
            'Filtered
            For i = 1 To allPostEmps.count
                If (allPostEmps.Item(i).OrganizationUnit.OrganizationUnitID = cboFilter.ItemData(cboFilter.ListIndex) And _
                allPostEmps.Item(i).IsDisengaged = False) Then
                    With lvwEmployee
                        Set ItemX = .ListItems.add(, , allPostEmps.Item(i).EmpCode)
                        ItemX.SubItems(1) = allPostEmps.Item(i).SurName & " " & allPostEmps.Item(i).OtherNames
                        ItemX.SubItems(2) = allPostEmps.Item(i).BasicPay
                        ItemX.SubItems(3) = allPostEmps.Item(i).HouseAllowance
                        ItemX.Tag = allPostEmps.Item(i).EmployeeID
                        eLoad = eLoad + 1
                    End With
                End If
            Next i
            lblLoaded = "Loaded: " & eLoad
            lblSelected = "Selected: 0"
        End If
    End If
    
    Exit Sub
errHandler:
    MsgBox ErrTitle & err.Description, vbExclamation, "PDR"
End Sub

Private Sub chkSelect_Click()
    If (chkSelect.value = vbChecked) Then
        chkSelect.Caption = "Uncheck All"
        eSeld = 0
        For i = 1 To lvwEmployee.ListItems.count
            lvwEmployee.ListItems.Item(i).Checked = True
            eSeld = eSeld + 1
        Next i
        lblSelected = "Selected: " & eSeld
    Else
        chkSelect.Caption = "Select All"
        For i = 1 To lvwEmployee.ListItems.count
            lvwEmployee.ListItems.Item(i).Checked = False
        Next i
        lblSelected.Caption = "Selected: 0"
    End If
End Sub

Private Sub cmdbrowse_Click()
On Error GoTo err
 cdlbox.ShowOpen
 txtimportfile.Text = cdlbox.FileName
 Exit Sub
err:
 MsgBox ("The following error occured: " & err.Description)
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdclose2_Click()
Unload Me
End Sub

Private Sub CmdGo_Click()
    On Error GoTo errHandler
    Dim tEmployee As HRCORE.Employee
    
    'Validation
    If (Trim(txtLowerScale) = "") Then
        MsgBox "Kindly provide the lower scale for the filter"
        txtLowerScale.SetFocus
        Exit Sub
    End If
    
    If Not (IsNumeric(Trim(txtLowerScale))) Then
        MsgBox "The value entered for the lower scale is invalid", vbExclamation, "Error"
        txtLowerScale.SetFocus
        txtLowerScale.SelLength = Len(txtLowerScale.Text)
        txtLowerScale.SelStart = 0
        Exit Sub
    End If

    If (Trim(txtUpperScale) = "") Then
        MsgBox "Kindly provide the upper scale for the filter"
        txtUpperScale.SetFocus
        Exit Sub
    End If
    
    If Not (IsNumeric(Trim(txtLowerScale))) Then
        MsgBox "The value entered for the upper scale is invalid", vbExclamation, "Error"
        txtUpperScale.SetFocus
        txtUpperScale.SelLength = Len(txtUpperScale.Text)
        txtUpperScale.SelStart = 0
        Exit Sub
    End If
    
    'No Errors,Proceed
    lvwEmployee.ListItems.Clear
    
    pSQL = "Select * From Employees Where " & CboPayScaleFilter.Text & " between " & CDbl(Trim(txtLowerScale)) & " and " & CDbl(Trim(txtUpperScale))
    Set pRST = CConnect.GetRecordSet(pSQL)
    
    If Not (pRST.EOF Or pRST.BOF) Then
        eLoad = 0
        eSeld = 0
        Do Until pRST.EOF
            With lvwEmployee
                
                Set ItemX = .ListItems.add(, , pRST!EmpCode)
                ItemX.SubItems(1) = pRST!SurName & " " & pRST!OtherNames
                ItemX.SubItems(2) = pRST!BasicPay
                ItemX.SubItems(3) = pRST!HouseAllowance
                ItemX.Tag = pRST!EmployeeID
                eLoad = eLoad + 1
            End With
            pRST.MoveNext
        Loop
    End If
   
    lblLoaded = "Loaded: " & eLoad
    lblSelected = "Selected: 0"
    
    Exit Sub
errHandler:
    MsgBox "An error has occured while filtering employees in that scale", vbExclamation, "Filter Error"
End Sub

Private Sub cmdimport_Click()

On Error GoTo err

    If chkBasic.value = vbChecked And chkhseall.value = vbChecked Then

        If MsgBox("Both Basic pay and House allowance will be Imported. Are You sure?", vbYesNo + vbInformation) = vbNo Then
        Exit Sub
        End If
    ElseIf chkBasic.value = vbChecked And chkhseall.value = vbUnchecked Then
        If MsgBox("Only Basic pay  will be Imported. Are You sure?", vbYesNo + vbInformation) = vbNo Then
        Exit Sub
        End If
    ElseIf chkBasic.value = vbUnchecked And chkhseall.value = vbChecked Then
        If MsgBox("Only House allowance  will be Imported. Are You sure?", vbYesNo + vbInformation) = vbNo Then
        Exit Sub
        End If
    ElseIf chkBasic.value = vbUnchecked And chkhseall.value = vbUnchecked Then
       MsgBox ("Neither Basic Pay nor House allowance can be imported")
       Exit Sub
    End If
    
    Dim conn As adodb.Connection
    Dim rsdata As adodb.Recordset
    Set conn = New adodb.Connection
    Dim CMD As adodb.Command
    Dim SelEmployee As HRCORE.Employee
    Dim intMissingEmpInDBRecords As Long
    lvwimportedfile.ListItems.Clear
    Dim li As ListItem
    Dim lngLoopVariable As Long
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtimportfile.Text & ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"""
    conn.Open
     If conn.State = adStateClosed Then
        MsgBox "The connection to the excel import file could not be established", vbExclamation
        GoTo Finish
    End If
    Set rsdata = New adodb.Recordset
    rsdata.Open "SELECT * FROM [EmployeeBasicpays$]", conn, adOpenKeyset, adLockOptimistic
    
        If rsdata.Fields.count < 3 Then
    
        MsgBox "The number of fields fall short of the expected fields" & vbCrLf & "Please confirm the expected fields on the expected fields display on the import wizard form", vbExclamation
        GoTo Finish
 
    End If
      With prgProgress
        .Min = 0
        .Max = rsdata.RecordCount
    End With
    
    
    
        Do Until rsdata.EOF
        
        'DISPLAYING THE PROGRESS OF THE IMPORTATION PROCESS
        ''fraProgress.Caption = "Importing : Record " & lngLoopVariable & " of " & rsdata.RecordCount & " Records"
       
        'ENSURING THAT ALL THE MANDATORY FIELDS ARE PRESENT
        If IsNull(rsdata.Fields(0)) Or IsNull(rsdata.Fields(1)) Or IsNull(rsdata.Fields(2)) Then
            ''txtLog.Text = txtLog.Text & Time & ":    MISSING DETAIL : One of the Mandatory Fields is Missing in Record No." & lngLoopVariable & vbNewLine
           '' intMissingMandatoryFieldRecords = intMissingMandatoryFieldRecords + 1
            GoTo NextRecord
        End If
 
 
 
        'INSTANTIATING THE COMMAND OBJECT
        Set CMD = New adodb.Command
       
            Set CMD = New adodb.Command
            With CMD
                .ActiveConnection = con
                .CommandType = adCmdStoredProc
                
                
                If chkBasic.value = vbChecked And chkhseall.value = vbChecked Then
                .CommandText = "prlspupdateBasicHse"
                ElseIf chkBasic.value = vbChecked And chkhseall.value = vbUnchecked Then
                   .CommandText = "prlspupdateBasic"
                ElseIf chkBasic.value = vbUnchecked And chkhseall.value = vbChecked Then
                   .CommandText = "prlspupdatehse"
                End If
                
                 
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 8, rsdata.Fields(0))
                .Parameters.Append .CreateParameter(, adCurrency, adParamInput, , CCur(IIf(IsNull(rsdata.Fields(1)), 0, rsdata.Fields(1))))
                .Parameters.Append .CreateParameter(, adCurrency, adParamInput, , CCur(IIf(IsNull(rsdata.Fields(2)), 0, rsdata.Fields(2))))
                 
                .Execute
            End With
             
     
        
    
NextRecord:
  lngLoopVariable = lngLoopVariable + 1
  
        Set SelEmployee = AllEmployees.FindEmployeeByCode(rsdata.Fields(0))
        If Not SelEmployee Is Nothing Then
        Set li = lvwimportedfile.ListItems.add(, , rsdata.Fields(0))
        li.SubItems(1) = SelEmployee.SurName & "  " & SelEmployee.OtherNames
        li.SubItems(2) = rsdata.Fields(1)
        li.SubItems(3) = rsdata.Fields(2)
        Else
                    'LOGGING
            txtLog.Text = txtLog.Text & Time & ":    MISSING DETAIL : Employee Code does not exist For Record No." & lngLoopVariable & vbNewLine
            intMissingEmpInDBRecords = intMissingEmpInDBRecords + 1
        Set li = lvwimportedfile.ListItems.add(, , rsdata.Fields(0))
        li.SubItems(1) = ""
        li.SubItems(2) = rsdata.Fields(1)
        li.SubItems(3) = rsdata.Fields(2)
        End If
      
        rsdata.MoveNext
        prgProgress.value = lngLoopVariable
    Loop
    If intMissingEmpInDBRecords > 0 Then
    MsgBox ("Importation Complete. " & intMissingEmpInDBRecords & " Skipped.")
    Else
    MsgBox ("Importation Complete")
    End If
    Exit Sub
Finish:
    Exit Sub
err:
    MsgBox ("The following error occured: " & err.Description)
End Sub

Private Sub cmdPost_Click()
    On Error GoTo errHandler
    'Post Adjustments
    
    'Validate entries
    If (cboAdjustment.Text = "") Then
        MsgBox "Select adjustment required", vbExclamation, "PDR"
        cboAdjustment.SetFocus
        Exit Sub
    End If
    
    If (cboAdjustmentType.Text = "") Then
        MsgBox "Select adjustment type required", vbExclamation, "PDR"
        cboAdjustmentType.SetFocus
        Exit Sub
    End If
    
    If (Trim(txtValue.Text) = "") Then
        MsgBox "Enter adjustment value", vbExclamation, "PDR Error"
        txtValue.SetFocus
        Exit Sub
    End If
    
    If Not (IsNumeric(Trim(txtValue))) Then
        MsgBox "The value provided is invalid in this context, provide a valid number", vbExclamation, "PDR"
        txtValue.SetFocus
        Exit Sub
    End If
    
    If eSeld = 0 Then
        MsgBox "No employees have been selected for adjustment", vbExclamation, "Error"
        Exit Sub
    End If
    
    'Actual Adjustments Now
    'Reset Controls
    cmdClose.Enabled = False
    cmdPost.Enabled = False
    pBar.Max = eSeld
    pBar.value = 0
    pBar.Visible = True
    
    If (cboAdjustmentType.Text = "Percentage") Then
        'Adjust by even percentage
        For i = 1 To lvwEmployee.ListItems.count
            If (lvwEmployee.ListItems.Item(i).Checked = True) Then
                pSQL = "Update Employees Set " & cboPay.Text & " = " & cboPay & "+(" & cboPay.Text & "*" & GetAdjustment(cboAdjustment) & "*0.01) Where EmployeeID = " & lvwEmployee.ListItems.Item(i).Tag
                'Set pRST =
                CConnect.ExecuteSql (pSQL)
                currUser.AuditTrail Update, (cboAdjustment & "d :" & lvwEmployee.ListItems.Item(i).Text & ", " & lvwEmployee.ListItems.Item(i).SubItems(1) & " " & cboPay & " from " & lvwEmployee.ListItems.Item(i).SubItems(2) & " to " & lvwEmployee.ListItems.Item(i).SubItems(2) + (lvwEmployee.ListItems.Item(i).SubItems(2) * GetAdjustment(cboAdjustment) * 0.01) & " by " & txtValue & "%")
                pBar = pBar + 1
            End If
        Next i
        MsgBox "Records updated successfully", vbInformation, "Adjustment"
    Else
        'Add even amount
        For i = 1 To lvwEmployee.ListItems.count
            If (lvwEmployee.ListItems.Item(i).Checked = True) Then
                pSQL = "Update Employees Set " & cboPay & " = " & cboPay & "+" & GetAdjustment(cboAdjustment) & " Where EmployeeID = " & lvwEmployee.ListItems.Item(i).Tag
                'Set pRST =
                CConnect.ExecuteSql (pSQL)
                currUser.AuditTrail Update, (cboAdjustment & "d :" & lvwEmployee.ListItems.Item(i).Text & ", " & lvwEmployee.ListItems.Item(i).SubItems(1) & " " & cboPay & " from " & lvwEmployee.ListItems.Item(i).SubItems(2) & " to " & lvwEmployee.ListItems.Item(i).SubItems(2) + (GetAdjustment(cboAdjustment)) & " by " & txtValue)
                pBar = pBar + 1
            End If
        Next i
        MsgBox "Records updated successfully", vbInformation, "Adjustment"
    End If
    
    'Revert Controls
    pBar.Visible = False
    cmdClose.Enabled = True
    cmdPost.Enabled = True
    chkSelect.value = vbUnchecked
    txtValue.Text = ""
    Filter = cboFilter
        
    'Reload
    cboFilter.Text = Filter
    cboFilter.SetFocus
    
    Exit Sub
errHandler:
    MsgBox ErrTitle & err.Description, vbExclamation, "PDR"
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    
    eSeld = 0: eLoad = 0
    
    'Load Ajustments
    With cboAdjustment
        .AddItem "Increase"
        .AddItem "Decrease"
        .Text = "Increase"
    End With
    
    'Load Adjust Types
    With cboAdjustmentType
        .AddItem "Percentage"
        .AddItem "Amount"
        .Text = "Percentage"
    End With
    
    LoadFilterCriteria
    
    'Load Pays
    With cboPay
        .AddItem "BasicPay"
        .AddItem "HouseAllowance"
        .Text = "BasicPay"
    End With
    
    'Load Pay Scale Filters
    With CboPayScaleFilter
        .AddItem "BasicPay"
        .AddItem "HouseAllowance"
        .Text = "BasicPay"
    End With
    
    allPostEmps.GetAccessibleEmployeesByUser currUser.UserID
    
    Exit Sub
    
errHandler:
    MsgBox ErrTitle & err.Description, vbExclamation, "PDR"
End Sub

Private Sub LoadFilterCriteria()

    If (optCategory.value = True) Then
        'Filter by Category
        cboFilter.Clear
        EmpCats.GetAllEmployeeCategories
        Me.cboFilter.AddItem "(All Categories)"
        For i = 1 To EmpCats.count
            Me.cboFilter.AddItem EmpCats.Item(i).CategoryName
            Me.cboFilter.ItemData(Me.cboFilter.NewIndex) = EmpCats.Item(i).CategoryID
        Next i
        cboFilter.ListIndex = 0
        
    ElseIf (optDepartments.value = True) Then
        'Filter By Departments
        Me.cboFilter.Clear
        OUs.GetAllOrganizationUnits
        cboFilter.AddItem "(All Organization Units)"
        For i = 1 To OUs.count
            cboFilter.AddItem OUs.Item(i).OrganizationUnitName
            cboFilter.ItemData(cboFilter.NewIndex) = OUs.Item(i).OrganizationUnitID
        Next i
        cboFilter.ListIndex = 0

    End If
End Sub

Private Sub lvwEmployee_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If (Item.Checked = True) Then
        eSeld = eSeld + 1
    Else
        eSeld = eSeld - 1
    End If
    lblSelected = "Selected: " & eSeld
End Sub

Private Sub optCategory_Click()
    cboFilter.Enabled = True
    LoadFilterCriteria
    chkSelect.value = vbUnchecked
    DisablePayScaleFilters
End Sub

Private Sub optDepartments_Click()
    cboFilter.Enabled = True
    LoadFilterCriteria
    chkSelect.value = vbUnchecked
    DisablePayScaleFilters
End Sub

Private Sub optPayScalesFilter_Click()
    cboFilter.Enabled = False
    txtLowerScale.Enabled = True
    txtUpperScale.Enabled = True
    CboPayScaleFilter.Enabled = True
    CmdGo.Enabled = True
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
    'If (InStr(1, "1234567890.", Chr(KeyAscii)) = True) Then
    If Not (IsNumeric(Chr(KeyAscii))) Then
        If Not (Chr(KeyAscii) = vbBack Or Chr(KeyAscii) = ".") Then
            KeyAscii = 0
        End If
    End If
End Sub

Function GetAdjustment(ByVal ADJ As String) As Double
    Select Case ADJ
        Case "Increase"
            GetAdjustment = CDbl(Trim(txtValue))
        Case "Decrease"
            GetAdjustment = CDbl(Trim(txtValue) * -1)
    End Select
End Function

Private Sub DisablePayScaleFilters()
    'Disable the others
    txtLowerScale.Text = ""
    txtLowerScale.Enabled = False
    txtUpperScale.Text = ""
    txtUpperScale.Enabled = False
    CboPayScaleFilter.Enabled = False
    CmdGo.Enabled = False
End Sub
