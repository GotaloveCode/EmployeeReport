VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmreportfilter 
   Caption         =   " "
   ClientHeight    =   5775
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   4905
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtreporttitle 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   18
      Text            =   "frmreportfilter.frx":0000
      Top             =   0
      Width           =   4575
   End
   Begin VB.CommandButton cmdview 
      Caption         =   "View"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   " "
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4575
      Begin VB.Frame Frame2 
         Caption         =   " "
         Height          =   3255
         Left            =   4680
         TabIndex        =   14
         Top             =   1080
         Visible         =   0   'False
         Width           =   4215
         Begin MSComctlLib.ListView lvwbankbranches 
            Height          =   2415
            Left            =   120
            TabIndex        =   17
            Top             =   720
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   4260
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Bank Branch"
               Object.Width           =   6068
            EndProperty
         End
         Begin VB.ComboBox cbobank 
            Height          =   315
            Left            =   120
            TabIndex        =   15
            Text            =   " "
            Top             =   360
            Width           =   3975
         End
         Begin VB.Label Label3 
            Caption         =   "Bank"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   120
            Width           =   2655
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Employment date Filter"
         Height          =   2535
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Visible         =   0   'False
         Width           =   4215
         Begin VB.OptionButton optafter 
            Caption         =   "After"
            Height          =   375
            Left            =   3000
            TabIndex        =   13
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optbefore 
            Caption         =   "Before"
            Height          =   315
            Left            =   1920
            TabIndex        =   12
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optalldates 
            Caption         =   "All"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optbetween 
            Caption         =   "Between"
            Height          =   255
            Left            =   840
            TabIndex        =   10
            Top             =   360
            Width           =   975
         End
         Begin MSComCtl2.DTPicker dtto 
            Height          =   255
            Left            =   1680
            TabIndex        =   8
            Top             =   840
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   108199937
            CurrentDate     =   39905
         End
         Begin MSComCtl2.DTPicker dtfrom 
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   108199937
            CurrentDate     =   39905
         End
      End
      Begin MSComctlLib.ListView lvwoptions 
         Height          =   3255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   5741
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "DESCRIPTION"
            Object.Width           =   6068
         EndProperty
      End
      Begin VB.ComboBox cbofilterby 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Text            =   " "
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "Options:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Filter By:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmreportfilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim reportTitle As String

Private Sub cbobank_Click()
loadbankbranches
End Sub
Private Sub loadbankbranches()
On Error GoTo ErrorHandler
    Dim rsOUs As New ADODB.Recordset, ItemD As ListItem
    lvwbankbranches.ListItems.Clear
    
    Set rsOUs = CConnect.GetRecordSet("select bankbranchid,branchname from bankbranches where bankid=" & cbobank.ItemData(cbobank.ListIndex) & "")
    ''If rsOUs.RecordCount > 0 Then
    If Not rsOUs.EOF Then
    Set ItemD = lvwbankbranches.ListItems.add(1, , "(Select All)")
    ItemD.Tag = "S-A"
    lvwbankbranches.ListItems.Item(1).Checked = True
    ItemD.Bold = True
    
        Do Until rsOUs.EOF
            With lvwbankbranches
                Set ItemD = .ListItems.add(, , rsOUs!BranchName)
                
                ItemD.Tag = rsOUs!BankBranchID
                
                ItemD.Checked = True
            End With
            rsOUs.MoveNext
        Loop
    End If
    

        
    Exit Sub
ErrorHandler:
    MsgBox "Error Loading Departments" & vbNewLine & err.Description, vbExclamation, "PDR Error"

End Sub
Private Sub cbofilterby_Click()
load_listview
End Sub
Private Sub load_listview()
lvwoptions.Visible = True
Frame5.Visible = False
Frame2.Visible = False
lvwoptions.ListItems.Clear
Dim li As ListItem
Select Case cbofilterby.Text
Case Is = "Employment_Date"
lvwoptions.Visible = False
Frame5.Visible = True
Case Is = "Marital_Status"
Set ItemD = lvwoptions.ListItems.add(1, , "(Select All)")
ItemD.Tag = "S-A"
lvwoptions.ListItems.Item(1).Checked = True
ItemD.Bold = True
Set li = lvwoptions.ListItems.add(, , "Single")
li.Tag = 1
li.Checked = True
Set li = lvwoptions.ListItems.add(, , "Married")
li.Tag = 2
li.Checked = True
Set li = lvwoptions.ListItems.add(, , "Divorced")
li.Tag = 3
li.Checked = True
Set li = lvwoptions.ListItems.add(, , "Separated")
li.Tag = 4
li.Checked = True
Set li = lvwoptions.ListItems.add(, , "Widowed")
li.Tag = 5
li.Checked = True
Set li = lvwoptions.ListItems.add(, , "Un Defined")
li.Tag = 0
li.Checked = True
Case Is = "Sex"
Set ItemD = lvwoptions.ListItems.add(1, , "(Select All)")
ItemD.Tag = "S-A"
lvwoptions.ListItems.Item(1).Checked = True
ItemD.Bold = True
Set li = lvwoptions.ListItems.add(, , "Male")
li.Tag = 1
li.Checked = True
Set li = lvwoptions.ListItems.add(, , "Female")
li.Tag = 2
li.Checked = True
Set li = lvwoptions.ListItems.add(, , "Un Defined")
li.Tag = 0
li.Checked = True
Case Is = "Department"
PopulateDpts
Case Is = "Designation"
loaddesignation
Case Is = "Employment_Term"
loadTerms
Case Is = "Nationality"
LoadNationalities
Case Is = "Blood_Group"
load_ddetails
Case Is = "Company_Code"
load_ddetails
Case Is = "Bureau_of_Statistic"
load_ddetails
Case Is = "Insurance_Cover"
load_ddetails
Case Is = "Present_Medical_Condition"
load_ddetails
Case Is = "Section"
load_ddetails
Case Is = "Reports_To"
load_ddetails
Case Is = "Bank_Name"
LoadBanks
Case Is = "Branch_Name"
Frame5.Visible = False
lvwoptions.Visible = False
Frame2.Left = Frame5.Left
Frame2.Width = Frame5.Width
Frame2.Visible = True
LoadBanks
Case Else

End Select
 
End Sub
Private Sub load_jds()
Dim rs As New Recordset, ItemD As ListItem

Set rs = CConnect.GetRecordSet("exec pdrspGetDistinctDdetails '" & cbofilterby.Text & "' ")
If Not rs.EOF Then
Set ItemD = lvwoptions.ListItems.add(1, , "(Select All)")
    ItemD.Tag = "S-A"
    lvwoptions.ListItems.Item(1).Checked = True
    ItemD.Bold = True
    
        Do Until rs.EOF
            With lvwoptions
              
                Set ItemD = .ListItems.add(, , rs!Detail)
              
                ItemD.Tag = rs!Code
                 
                ItemD.Checked = True
            End With
            rs.MoveNext
        Loop
End If

End Sub

Private Sub load_ddetails()
Dim rs As New Recordset, ItemD As ListItem

Set rs = CConnect.GetRecordSet("exec pdrspGetDistinctDdetails '" & cbofilterby.Text & "' ")
If Not rs.EOF Then
Set ItemD = lvwoptions.ListItems.add(1, , "(Select All)")
    ItemD.Tag = "S-A"
    lvwoptions.ListItems.Item(1).Checked = True
    ItemD.Bold = True
    
        Do Until rs.EOF
            With lvwoptions
              
                Set ItemD = .ListItems.add(, , rs!Detail)
              
                ItemD.Tag = rs!Code
                 
                ItemD.Checked = True
            End With
            rs.MoveNext
        Loop
End If

End Sub
Private Sub LoadBanks()

On Error GoTo ErrorHandler
    Dim rsOUs As New ADODB.Recordset, ItemD As ListItem
    cbobank.Clear
    Set rsOUs = CConnect.GetRecordSet("select bankid,bankname from banks")
    ''If rsOUs.RecordCount > 0 Then
    If Not rsOUs.EOF Then
    Set ItemD = lvwoptions.ListItems.add(1, , "(Select All)")
    ItemD.Tag = "S-A"
    lvwoptions.ListItems.Item(1).Checked = True
    ItemD.Bold = True
    
        Do Until rsOUs.EOF
            With lvwoptions
                Set ItemD = .ListItems.add(, , rsOUs!BankName)
                cbobank.AddItem (rsOUs!BankName)
                ItemD.Tag = rsOUs!bankid
                cbobank.ItemData(cbobank.NewIndex) = rsOUs!bankid
                ItemD.Checked = True
            End With
            rsOUs.MoveNext
        Loop
    End If
    

        
    Exit Sub
ErrorHandler:
    MsgBox "Error Loading Departments" & vbNewLine & err.Description, vbExclamation, "PDR Error"

End Sub
Private Sub LoadNationalities()
 On Error GoTo ErrorHandler
    Dim rsOUs As New ADODB.Recordset, ItemD As ListItem
    
    Set rsOUs = CConnect.GetRecordSet("select * from nationalities")
    ''If rsOUs.RecordCount > 0 Then
    If Not rsOUs.EOF Then
    Set ItemD = lvwoptions.ListItems.add(1, , "(Select All)")
    ItemD.Tag = "S-A"
    lvwoptions.ListItems.Item(1).Checked = True
    ItemD.Bold = True
    
        Do Until rsOUs.EOF
            With lvwoptions
                Set ItemD = .ListItems.add(, , rsOUs!Nationality)
                ItemD.Tag = rsOUs!NationalityID
                ItemD.Checked = True
            End With
            rsOUs.MoveNext
        Loop
    End If
    

        
    Exit Sub
ErrorHandler:
    MsgBox "Error Loading Departments" & vbNewLine & err.Description, vbExclamation, "PDR Error"

End Sub
Private Sub cmdClose_Click()
Unload Me
End Sub
Private Sub loaddesignation()
 On Error GoTo ErrorHandler
    Dim rsOUs As New ADODB.Recordset, ItemD As ListItem
    
    Set rsOUs = CConnect.GetRecordSet("select positionid,positionname from jobpositions")
    If rsOUs.RecordCount > 0 Then
        Do Until rsOUs.EOF
            With lvwoptions
                Set ItemD = .ListItems.add(, , rsOUs!PositionName)
                ItemD.Tag = rsOUs!PositionID
                ItemD.Checked = True
            End With
            rsOUs.MoveNext
        Loop
    End If
    
    Set ItemD = lvwoptions.ListItems.add(1, , "(Select All)")
    ItemD.Tag = "S-A"
    lvwoptions.ListItems.Item(1).Checked = True
    ItemD.Bold = True
        
    Exit Sub
ErrorHandler:
    MsgBox "Error Loading Departments" & vbNewLine & err.Description, vbExclamation, "PDR Error"
End Sub
Private Sub PopulateDpts()
    On Error GoTo ErrorHandler
    Dim rsOUs As New ADODB.Recordset, ItemD As ListItem
    
    Set rsOUs = CConnect.GetRecordSet("Select organizationunitid,organizationunitname From OrganizationUnits Order by OrganizationUnitName")
    If rsOUs.RecordCount > 0 Then
        Do Until rsOUs.EOF
            With lvwoptions
                Set ItemD = .ListItems.add(, , rsOUs!OrganizationUnitName)
                ItemD.Tag = rsOUs!OrganizationUnitID
                ItemD.Checked = True
            End With
            rsOUs.MoveNext
        Loop
    End If
    
    Set ItemD = lvwoptions.ListItems.add(1, , "(Select All)")
    ItemD.Tag = "S-A"
    lvwoptions.ListItems.Item(1).Checked = True
    ItemD.Bold = True
        
    Exit Sub
ErrorHandler:
    MsgBox "Error Loading Departments" & vbNewLine & err.Description, vbExclamation, "PDR Error"
End Sub
Private Sub loadTerms()
Dim rs3 As New Recordset
 Set rs3 = CConnect.GetRecordSet("SELECT code,Description FROM EmpTerms ORDER BY code")
    
     
     

    Dim li As ListItem
    With rs3
        ''If .RecordCount > 0 Then
        If Not .EOF Then
            .MoveFirst
            
               Set ItemD = lvwoptions.ListItems.add(1, , "(Select All)")
                ItemD.Tag = "S-A"
                lvwoptions.ListItems.Item(1).Checked = True
                ItemD.Bold = True
        
           
            Do While Not .EOF
                
                Set li = lvwoptions.ListItems.add(, , !Description)
                li.Tag = !Code
                li.Checked = True
                .MoveNext
            Loop
                Set li = lvwoptions.ListItems.add(, , "Un defined")
                li.Tag = 0
                li.Checked = True
        End If
    End With
    Set rs3 = Nothing

End Sub

Private Sub cmdView_Click()

    'VIEW REPORT OF AWARDS AWARDED TO EMPLOYEE
    Dim ids As String
    Dim what As Integer
    
    If cbofilterby.Text = "Branch_Name" Then
    ids = CreateTheIdsListFromListView(lvwbankbranches, True)
    ElseIf cbofilterby.Text = "Employment_Date" Then
    ids = "ALL"
    Else
    ids = CreateTheIdsListFromListView(lvwoptions, True)
    End If
    
    If Trim(ids) = "" Then
    MsgBox "Invalid Entries.", vbOKOnly + vbInformation
    Exit Sub
    End If
    
   Exit Sub
    
    If Trim(cbofilterby.Text) = "" Then
    MsgBox "Invalid Entries.", vbOKOnly + vbInformation
    Exit Sub
    End If
     reportTitle = txtreporttitle.Text
    
     If cbofilterby.Text = "Bank_Name" Or cbofilterby.Text = "Branch_Name" Then
     Set r = crptEmployeesBanks
     ElseIf cbofilterby.Text = "Employment_Date" Then
            If optalldates.value = True Then
            what = 1
            ElseIf optbetween.value = True Then
            what = 2
            ElseIf optbefore.value = True Then
            what = 3
            ElseIf optafter.value = True Then
            what = 4
            Else
            what = 0
            End If
    
            Set r = crptEmployeesEmpdate
     Else
     Set r = crptfilteredReport
     End If
     r.reportTitle = reportTitle
    '' R.ReportComments = "EMPLOYEE'S PERSONAL INFORMATION & BENEFICIARIES"
     ''frmRange2.Show
      
     
     Dim rs As New Recordset
     Dim str As String
   
     
     If cbofilterby.Text = "Employment_Date" Then
        ReDim objParamField(1 To 3)
        objParamField(1).Name = "@what"
        objParamField(1).value = what
        objParamField(2).Name = "@from"
        objParamField(2).value = CDate(Format((dtfrom.value), "dd-MM-yyyy"))
        objParamField(3).Name = "@to"
        objParamField(3).value = CDate(Format((dtto.value), "dd-MM-yyyy"))
     Else
         ReDim objParamField(1 To 2)
        objParamField(1).Name = "@filterby"
        objParamField(1).value = cbofilterby.Text
        objParamField(2).Name = "@ids"
        objParamField(2).value = ids
    End If
        ShowReport r, , True
End Sub
Public Function CreateTheIdsListFromListView(ByVal ctrlListView As ListView, ByVal blnChecked As Boolean) As String
'THIS FUNCTION LOOPS THRO THE PRESENTED EMPLOYEES AND CONCATENATES THE EMPLOYEE IDs OF ALL THE SELECTED
    Dim strCodeIdsList As String
    Dim lngLoopVariable As Long
    On Error GoTo ErrorHandler
    
    For lngLoopVariable = 2 To ctrlListView.ListItems.count
    
      If ctrlListView.ListItems(lngLoopVariable).Checked = blnChecked Then
       If Not strCodeIdsList = vbNullString Then
       
            strCodeIdsList = strCodeIdsList & ","
        
       End If
       
 
       If cbofilterby.Text = "Reports_To" Or cbofilterby.Text = "Present_Medical_Condition" Or cbofilterby.Text = "Insurance_Cover" Or cbofilterby.Text = "Company_Code" Or cbofilterby.Text = "Section" Or cbofilterby.Text = "Blood_Group" Or cbofilterby.Text = "Bureau_of_Statistic" Then
            strCodeIdsList = strCodeIdsList & "'" & ctrlListView.ListItems(lngLoopVariable).Text & "'"
       Else
            strCodeIdsList = strCodeIdsList & ctrlListView.ListItems(lngLoopVariable).Tag
       End If
       
'        If blnChecked Then
'            If ctrlListView.ListItems(lngLoopVariable).Checked = True Then GoTo EnterCodeIDInList
'        Else
'EnterCodeIDInList:
'            If Not strCodeIdsList = vbNullString Then strCodeIdsList = strCodeIdsList & ","
'            strCodeIdsList = strCodeIdsList & ctrlListView.ListItems(lngLoopVariable).Tag
        End If
    Next
Finish:
    CreateTheIdsListFromListView = strCodeIdsList
    Exit Function
    
ErrorHandler:
    MsgBox "An Error has occurred while attempting to create the employee list from the selected employees", vbExclamation
End Function


Private Sub Form_Load()

load_cbo
End Sub

Private Function togglechecks(ByVal Item As MSComctlLib.ListItem, lvw As ListView) As Integer
    On Error GoTo ErrorHandler
    Dim n As Integer, State As Boolean
    
    State = False
    If Item.Tag = "S-A" Then
    If Item.Checked = True Then
    State = True
    Else
    State = False
    End If
    End If
     
    
    If Item.Tag = "S-A" Then
        'Uncheck All Departments
        For n = 2 To lvw.ListItems.count
            lvw.ListItems.Item(n).Checked = State
        Next n
    End If
    
    Exit Function
ErrorHandler:
    MsgBox "Error Selecting Departments:" & vbNewLine & err.Description, vbExclamation, "PDR Error"

End Function
Private Sub load_cbo()
cbofilterby.AddItem ("Employment_Date")
cbofilterby.AddItem ("Marital_Status")
cbofilterby.AddItem ("Sex")
cbofilterby.AddItem ("Department")
cbofilterby.AddItem ("Designation")
cbofilterby.AddItem ("Employment_Term")
cbofilterby.AddItem ("Nationality")
cbofilterby.AddItem ("Blood_Group")
cbofilterby.AddItem ("Company_Code")
cbofilterby.AddItem ("Bureau_of_Statistic")
cbofilterby.AddItem ("Insurance_Cover")
cbofilterby.AddItem ("Present_Medical_Condition")
cbofilterby.AddItem ("Section")
cbofilterby.AddItem ("Reports_To")
cbofilterby.AddItem ("Bank_Name")
cbofilterby.AddItem ("Branch_Name")
 
End Sub

Private Sub lvwbankbranches_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim k
k = togglechecks(Item, lvwbankbranches)
End Sub

Private Sub lvwoptions_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim k As Integer
k = togglechecks(Item, lvwoptions)
End Sub

 

Private Sub optafter_Click()
If (optafter.value = True) Then
dtfrom.Visible = True
dtto.Visible = False
Else
dtfrom.Visible = False
dtto.Visible = False
End If
End Sub

Private Sub optalldates_Click()
If (optalldates.value = True) Then
dtfrom.Visible = False
dtto.Visible = False
Else
dtfrom.Visible = True
dtto.Visible = True
End If
End Sub

Private Sub optbefore_Click()
If (optbefore.value = True) Then
dtfrom.Visible = True
dtto.Visible = False
Else
dtfrom.Visible = False
dtto.Visible = False
End If
End Sub

Private Sub optbetween_Click()
If (optbetween.value = True) Then
dtfrom.Visible = True
dtto.Visible = True
Else
dtfrom.Visible = False
dtto.Visible = False
End If
End Sub
