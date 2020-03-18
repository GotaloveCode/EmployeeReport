VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEmployeeBankAccounts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Banks"
   ClientHeight    =   9510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9510
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imgEmpTool 
      Left            =   3840
      Top             =   600
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
            Picture         =   "frmEmployeeBankAccounts.frx":0000
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeeBankAccounts.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeeBankAccounts.frx":0224
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeeBankAccounts.frx":0336
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      Caption         =   "Employee Bank Account:"
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
      Height          =   5040
      Left            =   645
      TabIndex        =   18
      Top             =   840
      Visible         =   0   'False
      Width           =   6060
      Begin VB.TextBox TxtSwiftCode 
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
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   3300
         Width           =   5790
      End
      Begin VB.TextBox txtAccountName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Width           =   5775
      End
      Begin VB.ComboBox cboBranchName 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1320
         Width           =   5775
      End
      Begin VB.ComboBox cboBankName 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   5775
      End
      Begin VB.CheckBox chkIsMainAccount 
         Appearance      =   0  'Flat
         Caption         =   "&Mark this as the employee's main account"
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
         TabIndex        =   7
         Top             =   4440
         Width           =   3495
      End
      Begin VB.TextBox txtAccountType 
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
         Height          =   330
         Left            =   135
         TabIndex        =   6
         Top             =   3930
         Width           =   5790
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
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
         Left            =   4965
         Picture         =   "frmEmployeeBankAccounts.frx":0878
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Save Record"
         Top             =   4440
         Width           =   495
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
         Left            =   5445
         Picture         =   "frmEmployeeBankAccounts.frx":097A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Cancel Process"
         Top             =   4440
         Width           =   495
      End
      Begin VB.TextBox txtAccountNumber 
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
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   2685
         Width           =   5790
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Bank Account Swift Code"
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
         TabIndex        =   24
         Top             =   3045
         Width           =   1800
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Branch name"
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
         TabIndex        =   22
         Top             =   1080
         Width           =   930
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Bank name"
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
         TabIndex        =   21
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Account type"
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
         Left            =   135
         TabIndex        =   20
         Top             =   3675
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Account number"
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
         Left            =   135
         TabIndex        =   19
         Top             =   2475
         Width           =   1170
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
      Left            =   1560
      TabIndex        =   17
      ToolTipText     =   "Move to the Last employee"
      Top             =   5910
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
      Left            =   3735
      Picture         =   "frmEmployeeBankAccounts.frx":0A7C
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Delete Record"
      Top             =   5910
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
      Left            =   3255
      Picture         =   "frmEmployeeBankAccounts.frx":0F6E
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Edit Record"
      Top             =   5910
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
      Left            =   2775
      Picture         =   "frmEmployeeBankAccounts.frx":1070
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Add New record"
      Top             =   5910
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
      Left            =   1080
      TabIndex        =   13
      ToolTipText     =   "Move to the Next employee"
      Top             =   5910
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
      Left            =   600
      TabIndex        =   12
      ToolTipText     =   "Move to the Previous employee"
      Top             =   5910
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "CLOSE"
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
      Left            =   6270
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5910
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
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Move to the First employee"
      Top             =   5910
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComDlg.CommonDialog Cdl 
      Left            =   5610
      Top             =   5070
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvwDetails 
      Height          =   9360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   16510
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgTree"
      ForeColor       =   0
      BackColor       =   16777215
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Account Number"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Account Name"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bank Name"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Branch Name"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Is Main Account"
         Object.Width           =   2646
      EndProperty
   End
End
Attribute VB_Name = "frmEmployeeBankAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pBanks As Banks
Private pBankBranches As BankBranches
Private pEmployeeBankAccounts As EmployeeBankAccounts2

Private selBank As Bank
Private selBankBranch As bankbranch
Private selEmployeeBankAccount As EmployeeBankAccount2
Private empBankAccounts As EmployeeBankAccounts2

Private blnNewEntry As Boolean
Private ChangedFromCode As Boolean

Private myInternalPeriod As Period
Private openPeriod As Period

Private ThisMonth As Integer
Private ThisYear As Integer
Private Sub cboBankName_Click()
    On Error GoTo ErrorHandler
    
    If Not ChangedFromCode Then
        Set selBank = Nothing
        If Me.cboBankName.ListIndex > -1 Then
            Set selBank = pBanks.FindBankByID(Me.cboBankName.ItemData(Me.cboBankName.ListIndex))
            LoadBankBranchesOfbank selBank, True
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while processing the selected bank" & vbNewLine & err.Description, vbExclamation, TITLES
    
End Sub

Private Sub cboBranchName_Click()
    On Error GoTo ErrorHandler
    
    Set selBankBranch = Nothing
    If Me.cboBranchName.ListIndex > -1 Then
        Set selBankBranch = pBankBranches.FindBankBranchByID(Me.cboBranchName.ItemData(Me.cboBranchName.ListIndex))
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while processing the selected Bank Branch", vbExclamation, TITLES
    
End Sub

Public Sub cmdCancel_Click()
    frmMain2.RestoreCommandButtonState
    Me.fraDetails.Visible = False
    Me.cmdCancel.Enabled = False
    Me.cmdEdit.Enabled = True
    Me.cmdNew.Enabled = True
    Me.cmdSave.Enabled = False
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo ErrorHandler
    
    If SelectedEmployee Is Nothing Then
        MsgBox "First select an Employee", vbInformation, TITLES
        frmMain2.RestoreCommandButtonState
        Exit Sub
    Else
        If Not Me.lvwDetails.SelectedItem Is Nothing Then
            If MsgBox("Please confirm your decision to delete the selected employee bank account record", vbYesNo + vbExclamation, TITLES) = vbYes Then
                
                ''***************************************
                ''see if the employee is assigned  some amount in the bankaccount to be deleted
                If HasSomeAmountAssignedToIT(selEmployeeBankAccount.EmployeeBankAccountID, openPeriod.PeriodYear, monthno(openPeriod.PeriodMonth)) = True Then
                MsgBox "The Bank Account you are trying to delete has some amount Assigned to It." & vbCrLf & "Please Deallocate the Amount from the Payroll system then resume the Act", vbInformation
                Exit Sub
                End If
                ''***************************************
            
                
                selEmployeeBankAccount.Delete
                MsgBox "The employee bank account record has been deleted successfully", vbExclamation
                objEmployeeBankAccounts.RemoveByID selEmployeeBankAccount.EmployeeBankAccountID
                con.Execute ("Update EmployeeBankAccounts Set IsMainAccount=0 Where Deleted=1")
                DisplayRecords
            End If
        End If
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "An Error has occurred while attempting to delete an employee bank account record" & vbCrLf & err.Description, vbExclamation, TITLES
End Sub

Private Function monthno(mn As String) As Integer
Select Case UCase(mn)
Case UCase("January")
monthno = 1
Case UCase("February")
monthno = 2
Case UCase("March")
monthno = 3
Case UCase("April")
monthno = 4
Case UCase("May")
monthno = 5
Case UCase("June")
monthno = 6
Case UCase("July")
monthno = 7
Case UCase("August")
monthno = 8
Case UCase("September")
monthno = 9
Case UCase("October")
monthno = 10
Case UCase("November")
monthno = 11
Case UCase("December")
monthno = 12
Case Else
monthno = 12
End Select
End Function

Private Sub cmdEdit_Click()
    Dim lngLoopVariable As Long
    On Error GoTo ErrorHandler
    
    If SelectedEmployee Is Nothing Then
        MsgBox "First select an Employee", vbInformation, TITLES
        frmMain2.RestoreCommandButtonState
        Exit Sub
    Else
        If Not Me.lvwDetails.SelectedItem Is Nothing Then
            Me.fraDetails.Visible = True
            For lngLoopVariable = 0 To Me.cboBankName.ListCount - 1
                If Me.cboBankName.ItemData(lngLoopVariable) = selEmployeeBankAccount.bankbranch.Bank.bankid Then
                    Me.cboBankName.ListIndex = lngLoopVariable
                    Exit For
                End If
            Next
            For lngLoopVariable = 0 To Me.cboBranchName.ListCount - 1
                If Me.cboBranchName.ItemData(lngLoopVariable) = selEmployeeBankAccount.bankbranch.BankBranchID Then
                    Me.cboBranchName.ListIndex = lngLoopVariable
                    Exit For
                End If
            Next
            Me.txtAccountName.Text = selEmployeeBankAccount.AccountName
            Me.txtAccountNumber.Text = selEmployeeBankAccount.AccountNumber
            Me.txtAccountType.Text = selEmployeeBankAccount.AccountType
            Me.TxtSwiftCode = selEmployeeBankAccount.SwiftCode
            Me.chkIsMainAccount.value = IIf(selEmployeeBankAccount.IsMainAccount = True, 1, 0)

            Me.cmdCancel.Enabled = True
            Me.cmdSave.Enabled = True
            Me.cmdEdit.Enabled = False
            blnNewEntry = False
        End If
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "An Error has occurred while attempting to edit an employee bank account record" & vbCrLf & err.Description, vbExclamation, TITLES
End Sub

Public Sub cmdNew_Click()
    If SelectedEmployee Is Nothing Then
        MsgBox "First select an Employee", vbInformation, TITLES
        frmMain2.RestoreCommandButtonState
        Exit Sub
    Else
        Me.fraDetails.Visible = True
        Me.txtAccountName.Text = vbNullString
        Me.txtAccountNumber.Text = vbNullString
        Me.txtAccountType.Text = vbNullString
        Me.chkIsMainAccount.value = 0
        Me.cmdCancel.Enabled = True
        Me.cmdSave.Enabled = True
        Me.cmdEdit.Enabled = False
        blnNewEntry = True
    End If
End Sub

Public Sub cmdSave_Click()
    Dim myinternalEmployeeBankAccount As EmployeeBankAccount2
    On Error GoTo ErrorHandler
    
    If ValidateUserInput Then
        Set myinternalEmployeeBankAccount = New EmployeeBankAccount2
        With myinternalEmployeeBankAccount
            .AccountName = CStr(Me.txtAccountName.Text)
            .AccountNumber = Trim(Me.txtAccountNumber.Text)
            .AccountType = CStr(Me.txtAccountType.Text)
            .SwiftCode = CStr(Trim(Me.TxtSwiftCode))
            .IsMainAccount = CBool(Me.chkIsMainAccount.value)
            Set .bankbranch = selBankBranch
            Set .Employee = SelectedEmployee
            .EmployeeID = SelectedEmployee.EmployeeID
            .Deleted = False
            If blnNewEntry = True Then
            
            
            
             'check to see if the added bank is checked as main account
            Dim newbankismain As Boolean
            newbankismain = False
             Dim respo As Integer
            If myinternalEmployeeBankAccount.IsMainAccount Then
            newbankismain = True
            Dim i As Integer
            Dim k As Integer
            Dim Found As Boolean
            i = 1
            Found = False
            k = objEmployeeBankAccounts.count
            While (i <= k) And Found = False
            ''If (objEmployeeBankAccounts.Item(i).IsMainAccount = True And objEmployeeBankAccounts.Item(i).Employee.Employeeid = SelectedEmployee.Employeeid) Then
            If (objEmployeeBankAccounts.Item(i).IsMainAccount = True And objEmployeeBankAccounts.Item(i).EmployeeID = SelectedEmployee.EmployeeID) Then
            Found = True
            GoTo tt
            End If
            i = i + 1
            Wend
tt:
'            If Found Then
'                respo = MsgBox("Sorry !. The Employee already has an account with " & objEmployeeBankAccounts.Item(i).BankBranch.Bank.BankName & "   marked as his/her main account", vbCritical)
'                Set myinternalEmployeeBankAccount = Nothing
'            Exit Sub
'            End If
            
            
            
                If Found Then
                respo = MsgBox("The Employee already has an account with " & objEmployeeBankAccounts.Item(i).bankbranch.Bank.BankName & " marked as his/her main account. Make the Currennt Account as main account?", vbYesNo + vbCritical)
                
                If respo = vbYes Then
                objEmployeeBankAccounts.Item(i).IsMainAccount = False
                Else
                .IsMainAccount = False
                End If
'                    respo = MsgBox("Sorry !. The Employee already has an account with " & objEmployeeBankAccounts.Item(i).BankBranch.Bank.BankName & "   marked as his/her main account", vbCritical)
'                    Set myinternalEmployeeBankAccount = Nothing
                'Exit Sub
                End If
            Else ''the new entry is not marked as a main account
                newbankismain = False
                Dim iB As Integer
                Dim kB As Integer
                Dim FoundB As Boolean
                iB = 1
                Found = False
                kB = objEmployeeBankAccounts.count
                While (iB <= kB) And Found = False
                If (objEmployeeBankAccounts.Item(iB).IsMainAccount = True And objEmployeeBankAccounts.Item(iB).Employee.EmployeeID = SelectedEmployee.EmployeeID) Then
                FoundB = True
                GoTo ttB
                End If
                iB = iB + 1
                Wend
                
ttB:
             If FoundB = False Then
             .IsMainAccount = True
             End If
            End If
            
            'end checking of existsing banks as main accounts
            
            
                .InsertNew
            Else
            
                .EmployeeBankAccountID = selEmployeeBankAccount.EmployeeBankAccountID
                
                 
                
                            'check to see if the added bank is checked as main account
            
            newbankismain = False
            If myinternalEmployeeBankAccount.IsMainAccount Then
            newbankismain = True
        
           
            i = 1
            Found = False
            k = objEmployeeBankAccounts.count
            While (i <= k) And Found = False
           '' If (objEmployeeBankAccounts.Item(i).IsMainAccount = True And objEmployeeBankAccounts.Item(i).Employee.Employeeid = SelectedEmployee.Employeeid) Then
            If (objEmployeeBankAccounts.Item(i).IsMainAccount = True And objEmployeeBankAccounts.Item(i).EmployeeID = SelectedEmployee.EmployeeID) Then
            If Not (myinternalEmployeeBankAccount.EmployeeBankAccountID = objEmployeeBankAccounts.Item(i).EmployeeBankAccountID) Then
            Found = True
            End If
            GoTo yy
            End If
            i = i + 1
            Wend
yy:
'            If Found Then
'
'
'
'            respo = MsgBox("Sorry !. The Employee already has an account with " & objEmployeeBankAccounts.Item(i).BankBranch.Bank.BankName & "   marked as his/her main account", vbCritical)
'            Set myinternalEmployeeBankAccount = Nothing
'            Exit Sub
'            End If
            
            
            
                If Found Then
                respo = MsgBox("The Employee already has an account with " & objEmployeeBankAccounts.Item(i).bankbranch.Bank.BankName & " marked as his/her main account. Make the Currennt Account as main account?", vbYesNo + vbCritical)
                
                If respo = vbYes Then
                objEmployeeBankAccounts.Item(i).IsMainAccount = False
                Else
                .IsMainAccount = False
                End If
                ''Set myinternalEmployeeBankAccount = Nothing
            
            
            
            
                End If
            Else '' the edited bankaccount is marked as not ain account
            newbankismain = False
        
          
       
            i = 1
            Found = False
            k = objEmployeeBankAccounts.count
            While (i <= k) And Found = False
           ''If (objEmployeeBankAccounts.Item(i).IsMainAccount = True And objEmployeeBankAccounts.Item(i).Employee.Employeeid = SelectedEmployee.Employeeid) Then
             If (objEmployeeBankAccounts.Item(i).IsMainAccount = True And objEmployeeBankAccounts.Item(i).EmployeeID = SelectedEmployee.EmployeeID) Then
            If Not (myinternalEmployeeBankAccount.EmployeeBankAccountID = objEmployeeBankAccounts.Item(i).EmployeeBankAccountID) Then
            Found = True
            End If
            GoTo yyB
            End If
            i = i + 1
            Wend
yyB:
            If Found = False Then
            .IsMainAccount = True
            End If
            End If
            'end checking of existsing banks as main accounts
                
                
                .Update
                'REMOVING THE REDUNDANT RECORD
                objEmployeeBankAccounts.RemoveByID selEmployeeBankAccount.EmployeeBankAccountID
            End If
            

            
            
            
            objEmployeeBankAccounts.add myinternalEmployeeBankAccount
            MsgBox "The new employee bank account has been saved successfully", vbExclamation
            'NOW REPOPULATING THE EMPLOYEE BANK ACCOUNTS
            DisplayRecords
        End With
        Me.fraDetails.Visible = False
        Me.cmdCancel.Enabled = False
        Me.cmdEdit.Enabled = True
        Me.cmdNew.Enabled = True
        frmMain2.RestoreCommandButtonState
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "An Error has occurred while attempting to save an employee bank account entry" & vbCrLf & err.Description, vbExclamation, TITLES
End Sub


Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    'position the form
    frmMain2.PositionTheFormWithEmpList Me
'    oSmart.FReset Me
    
    Set pBanks = New Banks
    Set pBankBranches = New BankBranches
    ''Set objEmployeeBankAccounts = New EmployeeBankAccounts2
    
    
    Set myInternalPeriod = New Period
    
    myInternalPeriod.GetAllPeriods
    Set openPeriod = New Period
    Set openPeriod = myInternalPeriod.GetOpenPeriod
    'Load the Banks and Populate them
    LoadBanks
    
    'load the bank branches but don't populate
  ''  pBankBranches.GetActiveBankBranches
    
    'load employee bank accounts but don't populate
   '' objEmployeeBankAccounts.GetActiveEmployeeBankAccounts
    
    DisplayRecords
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while opening the window" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Function ValidateUserInput() As Boolean
    'RETS TRUE UPON SUCCESSFUL VALIDATION
    On Error GoTo ErrorHandler
    
    If Me.cboBankName.Text = vbNullString Then
        MsgBox "Please select the employee bank", vbExclamation
        Me.cboBankName.SetFocus
    Else
        If Me.cboBranchName.Text = vbNullString Then
            MsgBox "Please select the employee bank branch", vbExclamation
            Me.cboBranchName.SetFocus
        Else
            If IsNumeric(Me.txtAccountName.Text) Then
                MsgBox "The bank account name detail is not valid", vbExclamation
                Me.txtAccountName.SetFocus
            Else
                If Me.txtAccountNumber.Text = vbNullString Then 'Or Not IsNumeric(Me.txtAccountNumber.Text) Then
                    MsgBox "The bank account number detail is not valid", vbExclamation
                    Me.txtAccountNumber.SetFocus
                Else
                    If IsNumeric(Me.txtAccountType.Text) Then
                        MsgBox "The bank account type detail is not valid", vbExclamation
                        Me.txtAccountType.SetFocus
                    Else
                        ValidateUserInput = True
                    End If
                End If
            End If
        End If
    End If
    Exit Function
    
ErrorHandler:
    MsgBox "An Error has occurred while attempting to validate the employee bank account user input" & vbCrLf & err.Description, vbExclamation, TITLES
End Function
Private Sub LoadBanks()
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    pBanks.GetActiveBanks
    
    For i = 1 To pBanks.count
        Me.cboBankName.AddItem "[" & pBanks.Item(i).BankCode & "] " & pBanks.Item(i).BankName
        Me.cboBankName.ItemData(cboBankName.NewIndex) = pBanks.Item(i).bankid
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while Populating the Banks" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub LoadBankBranchesOfbank(ByVal TheBank As Bank, ByVal Refresh As Boolean)
    Dim i As Long
    Dim TheBranches As BankBranches
    
    On Error GoTo ErrorHandler
    
    Me.cboBranchName.Clear
    
    If TheBank Is Nothing Then
        Exit Sub
    End If
    
    If Refresh = True Then
        pBankBranches.GetActiveBankBranches
    End If
    
    Set TheBranches = pBankBranches.GetBranchesOfBankID(TheBank.bankid)
    If Not (TheBranches Is Nothing) Then
        For i = 1 To TheBranches.count
            Me.cboBranchName.AddItem "[" & TheBranches.Item(i).BranchCode & "] " & TheBranches.Item(i).BranchName
            Me.cboBranchName.ItemData(Me.cboBranchName.NewIndex) = TheBranches.Item(i).BankBranchID
        Next i
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while populating the Branches of the selected Bank" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub


Private Function InsertNew() As Boolean
    Dim newEmpAccount As EmployeeBankAccount2
    
    On Error GoTo ErrorHandler
    
    Set newEmpAccount = New EmployeeBankAccount2
    
    Exit Function
    
ErrorHandler:
    
    
End Function

Public Sub DisplayRecords()
    'THIS METHOD DISPLAYS THE EMPLOYEE BANK ACCOUNT DETAILS IN THE LIST VIEW CONTROL
    Dim myListItem As ListItem
    Dim lngLoopVariable As Long
    On Error GoTo ErrorHandler
    
    'CLEARING THE LIST VIEW CONTROL
    Me.lvwDetails.ListItems.Clear
    If Not SelectedEmployee Is Nothing Then
        ''Set empBankAccounts = objEmployeeBankAccounts.GetEmployeeBankAccountsOfEmployeeID(SelectedEmployee.EmployeeID)
        Set empBankAccounts = objEmployeeBankAccounts.GetEmployeeBankAccountsOfEmployeeID(SelectedEmployee.EmployeeID)
        If Not empBankAccounts Is Nothing Then
            For lngLoopVariable = 1 To empBankAccounts.count
                Set myListItem = Me.lvwDetails.ListItems.add(, , empBankAccounts.Item(lngLoopVariable).AccountNumber)
                myListItem.SubItems(1) = empBankAccounts.Item(lngLoopVariable).AccountName
                myListItem.SubItems(2) = empBankAccounts.Item(lngLoopVariable).bankbranch.Bank.BankName
                myListItem.SubItems(3) = empBankAccounts.Item(lngLoopVariable).bankbranch.BranchName
                myListItem.SubItems(4) = empBankAccounts.Item(lngLoopVariable).IsMainAccount
                myListItem.Tag = empBankAccounts.Item(lngLoopVariable).EmployeeBankAccountID
            Next
            If Me.lvwDetails.ListItems.count > 0 Then
                lvwDetails_ItemClick Me.lvwDetails.ListItems.Item(1)
            End If
        End If
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "An Error has occurred while attempting to display the employee bank account records on the display" & vbCrLf & err.Description, vbExclamation, TITLES
    
End Sub

Private Sub lvwDetails_DblClick()
    Me.cmdEdit.value = True
End Sub

Private Sub lvwDetails_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim lngLoopVariable As Long
    On Error GoTo ErrorHandler
    
    If Not Me.lvwDetails.SelectedItem Is Nothing Then
        Set selEmployeeBankAccount = objEmployeeBankAccounts.FindEmployeeBankAccountByID(Me.lvwDetails.SelectedItem.Tag)
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "An Error has occurred while attempting to process the itemclick event of the employee bank accounts list view control" & vbCrLf & err.Description, vbExclamation, TITLES
End Sub
