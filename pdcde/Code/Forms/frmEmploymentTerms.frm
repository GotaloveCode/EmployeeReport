VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEmploymentTerms 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employment terms"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
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
      Left            =   4830
      TabIndex        =   19
      Top             =   6570
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame fraExisting 
      Caption         =   "Existing Employment Terms"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   90
      TabIndex        =   16
      Top             =   3420
      Width           =   6675
      Begin MSComctlLib.ListView lvwEmpTerms 
         Height          =   2715
         Left            =   90
         TabIndex        =   9
         Top             =   270
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   4789
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Term Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Term Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Matched To:"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Details"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.Frame fraDetails 
      Caption         =   "Details"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2625
      Left            =   120
      TabIndex        =   12
      Top             =   540
      Width           =   6675
      Begin VB.TextBox txtTermDetails 
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
         Left            =   1080
         TabIndex        =   3
         Top             =   720
         Width           =   5415
      End
      Begin VB.Frame fraMappedTo 
         Caption         =   "Match Employment Term To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   120
         TabIndex        =   15
         Top             =   1170
         Width           =   6375
         Begin VB.OptionButton optOther 
            Caption         =   "Other Employment Term"
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
            Left            =   180
            TabIndex        =   8
            Top             =   900
            Width           =   2715
         End
         Begin VB.OptionButton optExpatriate 
            Caption         =   "Expatriate Terms"
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
            Left            =   3510
            TabIndex        =   7
            Top             =   630
            Width           =   1545
         End
         Begin VB.OptionButton optContract 
            Caption         =   "Contract Terms"
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
            Left            =   3510
            TabIndex        =   5
            Top             =   315
            Width           =   1635
         End
         Begin VB.OptionButton optCasual 
            Caption         =   "Casual Terms"
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
            Left            =   180
            TabIndex        =   6
            Top             =   630
            Width           =   1365
         End
         Begin VB.OptionButton optPermanent 
            Caption         =   "Permanent Terms"
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
            Left            =   180
            TabIndex        =   4
            Top             =   360
            Width           =   1635
         End
      End
      Begin VB.TextBox txtTermName 
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
         Left            =   3060
         TabIndex        =   2
         Top             =   270
         Width           =   3435
      End
      Begin VB.TextBox txtTermCode 
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
         Left            =   1080
         TabIndex        =   1
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Details"
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
         TabIndex        =   18
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Term Name"
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
         Left            =   2070
         TabIndex        =   14
         Top             =   315
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Term Code"
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
         TabIndex        =   13
         Top             =   315
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdDelete 
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
      Left            =   6375
      Picture         =   "frmEmploymentTerms.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Delete Record"
      Top             =   6570
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdEdit 
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
      Left            =   5895
      Picture         =   "frmEmploymentTerms.frx":04F2
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Edit Record"
      Top             =   6570
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdNew 
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
      Left            =   5415
      Picture         =   "frmEmploymentTerms.frx":05F4
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Add New record"
      Top             =   6570
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Employment Terms"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   17
      Top             =   90
      Width           =   3165
   End
End
Attribute VB_Name = "frmEmploymentTerms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private empTerms As HRCORE.EmploymentTerms
Private selEmpTerm As HRCORE.EmploymentTerm
Private blnEditMode As Boolean

Public Sub cmdDelete_Click()
    Dim resp As Long
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    If Not currUser Is Nothing Then
        If currUser.CheckRight("SetEmpTerms") <> secModify Then
            MsgBox "You dont have right to delete record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    If selEmpTerm Is Nothing Then
        MsgBox "You have to select the Employment Term to delete", vbInformation, TITLES
        Exit Sub
    End If
    resp = MsgBox("Are you sure you want to delete the Selected Employment Term i.e." & UCase(selEmpTerm.EmpTermName), vbYesNo + vbQuestion, TITLES)
    If resp = vbYes Then
        retVal = selEmpTerm.Delete()
        If retVal = 0 Then
            MsgBox "The Employment term has been deleted", vbInformation, TITLES
            LoadEmploymentTerms
            'reload employees
            AllEmployees.GetAccessibleEmployeesByUser (currUser.UserID)
        Else
            MsgBox "The Employment Term could not be deleted", vbInformation, TITLES
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while deleting the Employment term" & _
    vbNewLine & err.Description, vbInformation, TITLES
End Sub

Public Sub cmdEdit_Click()
    If Not currUser Is Nothing Then
        If currUser.CheckRight("SetEmpTerms") <> secModify Then
            MsgBox "You dont have right to edit the record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    If selEmpTerm Is Nothing Then
        MsgBox "You have to select the Employment term to edit", vbInformation, TITLES
        Exit Sub
    End If
    Me.fraDetails.Enabled = True
    Me.fraExisting.Enabled = False
    Me.txtTermCode.SetFocus
    blnEditMode = True
End Sub

Public Sub cmdNew_Click()
    If Not currUser Is Nothing Then
        If currUser.CheckRight("SetEmpTerms") <> secModify Then
            MsgBox "You dont have right to add new record. Please liaise with the security admin"
            Exit Sub
        End If
        
    End If
    ClearText
    Me.fraDetails.Enabled = True
    Me.fraExisting.Enabled = False
    Me.txtTermCode.SetFocus
End Sub

Public Sub cmdSave_Click()
        
    If Not currUser Is Nothing Then
        If currUser.CheckRight("SetEmpTerms") <> secModify Then
            MsgBox "You dont have right to add new record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    If Update = True Then
        'reload employment terms
        LoadEmploymentTerms
        
        'reload them in frmMain2
        Call frmMain2.LoadEmploymentTerms
        
        If blnEditMode = True Then
            'reload employees
            Call frmMain2.LoadEmployeeList
            MsgBox "The Employment Term has been Updated successfully", vbInformation, TITLES

        Else
            MsgBox "The New Employment Term has been created", vbInformation, TITLES
        End If
       
        
        Me.fraDetails.Enabled = False
        Me.fraExisting.Enabled = True
    Else
        If blnEditMode = True Then
            MsgBox "The Employment term could not be Updated", vbInformation, TITLES
        Else
            MsgBox "The Employment Term could not be created", vbInformation, TITLES
        End If
    End If
   
End Sub

Private Function Update() As Boolean
    Dim NewEmpTerm As HRCORE.EmploymentTerm
    Dim retVal As Long
    Dim CheckEmpTerm As HRCORE.EmploymentTerm
    
    On Error GoTo ErrorHandler
    Set NewEmpTerm = New HRCORE.EmploymentTerm
    With NewEmpTerm
        If blnEditMode = True Then
            .EmpTermID = selEmpTerm.EmpTermID
        End If
        
        If Len(Trim(Me.txtTermCode.Text)) > 0 Then
            .EmpTermCode = Trim(Me.txtTermCode.Text)
            If blnEditMode = True Then
                 If Not (empTerms.FindEmploymentTermByCodeExclusive(.EmpTermCode, .EmpTermID) Is Nothing) Then
                    MsgBox "Another Employment Term exists with the supplied Code", vbInformation, TITLES
                    Me.txtTermCode.SetFocus
                    Exit Function
                End If
            Else
                If Not (empTerms.FindEmploymentTermByCode(.EmpTermCode) Is Nothing) Then
                    MsgBox "Another Employment Term exists with the supplied Code", vbInformation, TITLES
                    Me.txtTermCode.SetFocus
                    Exit Function
                End If
            End If
        Else
            MsgBox "Enter the Unique code to identify the Employment Term", vbInformation, TITLES
            Me.txtTermCode.SetFocus
            Exit Function
        End If
        
        If Len(Trim(Me.txtTermName.Text)) > 0 Then
            .EmpTermName = Trim(Me.txtTermName.Text)
            
            If blnEditMode = True Then
                If Not empTerms.FindEmploymentTermByNameExclusive(.EmpTermName, .EmpTermID) Is Nothing Then
                    MsgBox "Another Employment term exists with the supplied Name", vbInformation, TITLES
                    Me.txtTermName.SetFocus
                    Exit Function
                End If
            Else
                If Not (empTerms.FindEmploymentTermByName(.EmpTermName) Is Nothing) Then
                    MsgBox "Another Employment term exists with the supplied Name", vbInformation, TITLES
                    Me.txtTermName.SetFocus
                    Exit Function
                End If
            End If
        Else
            MsgBox "Enter the Name to identify the Employment Term", vbInformation, TITLES
            Me.txtTermName.SetFocus
            Exit Function
        End If
        
        .EmpTermDetails = Trim(Me.txtTermDetails.Text)
        'first set the default mapping
        .MappedToEmpTerm = Other
        
        'then find the actual mapping
        If Me.optCasual.value = True Then .MappedToEmpTerm = Casual
        If Me.optPermanent.value = True Then .MappedToEmpTerm = Permanent
        If Me.optContract.value = True Then .MappedToEmpTerm = Contract
        If Me.optExpatriate.value = True Then .MappedToEmpTerm = Expatriate
        
        If blnEditMode = True Then
            retVal = .Update()
            blnEditMode = False
        Else
            retVal = .InsertNew()
        End If
    End With
    
    If retVal = 0 Then
        Update = True
    Else
        Update = False
    End If
    
    Exit Function
        
ErrorHandler:
    MsgBox "An error has occurred while creating the new Employment Term" & _
    vbNewLine & err.Description, vbInformation, TITLES
    Update = False
End Function

    
Private Sub Form_Load()
        
    On Error GoTo ErrorHandler
    oSmart.FReset Me
    
    Set empTerms = New HRCORE.EmploymentTerms
    LoadEmploymentTerms
    
    frmMain2.PositionTheFormWithoutEmpList Me
    
    'force the selected emp term to nothing
    Set selEmpTerm = Nothing
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred" & vbNewLine & err.Description, vbInformation, TITLES
End Sub

Private Sub Form_Resize()
    oSmart.FResize Me
End Sub

Private Sub LoadEmploymentTerms()
    Dim i As Long
    Dim ItemX As ListItem
    
    On Error GoTo ErrorHandler
    Me.lvwEmpTerms.ListItems.Clear
    
    empTerms.GetAllEmploymentTerms
    
    For i = 1 To empTerms.count
        Set ItemX = lvwEmpTerms.ListItems.add(, , empTerms.Item(i).EmpTermCode)
        ItemX.SubItems(1) = empTerms.Item(i).EmpTermName
        ItemX.SubItems(2) = empTerms.Item(i).MappedToEmpTermStr
        ItemX.SubItems(3) = empTerms.Item(i).EmpTermDetails
        ItemX.Tag = empTerms.Item(i).EmpTermID
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while retrieving Employment Terms" & _
    vbNewLine & err.Description, vbInformation, TITLES
    
End Sub

Public Sub ClearText()

    Me.txtTermCode.Text = ""
    Me.txtTermDetails.Text = ""
    Me.txtTermName.Text = ""
    Me.optOther.value = True
    
End Sub

Private Sub SetFields(ByVal SelectedEmpTerm As HRCORE.EmploymentTerm)
    On Error GoTo ErrorHandler
    
    If Not (SelectedEmpTerm Is Nothing) Then
        With SelectedEmpTerm
            Me.txtTermCode.Text = .EmpTermCode
            Me.txtTermDetails.Text = .EmpTermDetails
            Me.txtTermName.Text = .EmpTermName
            
            Me.optCasual.value = .IsCasual
            Me.optContract.value = .IsContract
            Me.optExpatriate.value = .IsExpatriate
            Me.optPermanent.value = .IsPermanent
        End With
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while populating Emplyment term details" & _
    vbNewLine & err.Description, vbInformation, TITLES
    
End Sub

Private Sub lvwEmpTerms_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo ErrorHandler
    
    Set selEmpTerm = Nothing
    ClearText
    
    Set selEmpTerm = empTerms.FindEmploymentTerm(CLng(Item.Tag))
    If Not (selEmpTerm Is Nothing) Then
        SetFields selEmpTerm
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while selecting an Employment term" & _
    vbNewLine & err.Description, vbInformation, TITLES
End Sub
