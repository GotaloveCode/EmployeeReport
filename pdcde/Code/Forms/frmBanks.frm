VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmBanks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Banks"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   Icon            =   "frmBanks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab sstBanks 
      Height          =   6300
      Left            =   30
      TabIndex        =   16
      Top             =   75
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   11113
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Banks"
      TabPicture(0)   =   "frmBanks.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraBank"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraExistingBanks"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdNewBank"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdEditBank"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdDeleteBank"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdBranches"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdClose"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Bank Branches"
      TabPicture(1)   =   "frmBanks.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(1)=   "fraExistingBranches"
      Tab(1).Control(2)=   "fraBranch"
      Tab(1).Control(3)=   "txtParentBank"
      Tab(1).Control(4)=   "cmdNew"
      Tab(1).Control(5)=   "cmdEdit"
      Tab(1).Control(6)=   "cmdDelete"
      Tab(1).Control(7)=   "cmdBack"
      Tab(1).ControlCount=   8
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   465
         Left            =   5700
         TabIndex        =   7
         Top             =   5700
         Width           =   1215
      End
      Begin VB.CommandButton cmdBranches 
         Caption         =   "Branches..."
         Height          =   465
         Left            =   3825
         TabIndex        =   6
         Top             =   5700
         Width           =   1440
      End
      Begin VB.CommandButton cmdDeleteBank 
         Caption         =   "Delete"
         Height          =   465
         Left            =   2475
         TabIndex        =   5
         Top             =   5700
         Width           =   1140
      End
      Begin VB.CommandButton cmdEditBank 
         Caption         =   "Edit"
         Height          =   465
         Left            =   1425
         TabIndex        =   4
         Top             =   5700
         Width           =   990
      End
      Begin VB.CommandButton cmdNewBank 
         Caption         =   "New"
         Height          =   465
         Left            =   300
         TabIndex        =   3
         Top             =   5700
         Width           =   990
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back To Banks"
         Height          =   465
         Left            =   -69750
         TabIndex        =   15
         Top             =   5820
         Width           =   1590
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   465
         Left            =   -72150
         TabIndex        =   14
         Top             =   5820
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   465
         Left            =   -73575
         TabIndex        =   13
         Top             =   5820
         Width           =   1215
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   465
         Left            =   -74850
         TabIndex        =   12
         Top             =   5820
         Width           =   1140
      End
      Begin VB.TextBox txtParentBank 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74280
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   525
         Width           =   5490
      End
      Begin VB.Frame fraBranch 
         Height          =   1140
         Left            =   -74850
         TabIndex        =   21
         Top             =   1050
         Width           =   6765
         Begin VB.TextBox txtBranchName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1275
            TabIndex        =   10
            Top             =   675
            Width           =   5370
         End
         Begin VB.TextBox txtBranchCode 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1275
            TabIndex        =   9
            Top             =   300
            Width           =   1800
         End
         Begin VB.Label Label7 
            Caption         =   "Branch Name:"
            Height          =   165
            Left            =   150
            TabIndex        =   23
            Top             =   750
            Width           =   1035
         End
         Begin VB.Label Label6 
            Caption         =   "Branch Code:"
            Height          =   165
            Left            =   195
            TabIndex        =   22
            Top             =   375
            Width           =   1065
         End
      End
      Begin VB.Frame fraExistingBranches 
         Caption         =   "Bank Branches:"
         Height          =   3360
         Left            =   -74850
         TabIndex        =   20
         Top             =   2325
         Width           =   6765
         Begin MSComctlLib.ListView lvwBankBranches 
            Height          =   2955
            Left            =   150
            TabIndex        =   11
            Top             =   300
            Width           =   6465
            _ExtentX        =   11404
            _ExtentY        =   5212
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Branch Code"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Branch Name"
               Object.Width           =   7056
            EndProperty
         End
      End
      Begin VB.Frame fraExistingBanks 
         Caption         =   "Existing Banks:"
         Height          =   3810
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   6765
         Begin MSComctlLib.ListView lvwBanks 
            Height          =   3285
            Left            =   150
            TabIndex        =   2
            Top             =   375
            Width           =   6465
            _ExtentX        =   11404
            _ExtentY        =   5794
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Bank Code"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Bank Name"
               Object.Width           =   7056
            EndProperty
         End
      End
      Begin VB.Frame fraBank 
         Height          =   1215
         Left            =   150
         TabIndex        =   17
         Top             =   450
         Width           =   6765
         Begin VB.TextBox txtBankName 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            TabIndex        =   1
            Top             =   690
            Width           =   5520
         End
         Begin VB.TextBox txtBankCode 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1065
            TabIndex        =   0
            Top             =   225
            Width           =   2745
         End
         Begin VB.Label Label8 
            Caption         =   "Bank Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   705
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Bank Code:"
            Height          =   240
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Bank:"
         Height          =   240
         Left            =   -74775
         TabIndex        =   24
         Top             =   600
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmBanks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Private pBanks As HRCORE.Banks
Private selBank As Bank
''Private pBankBranches As BankBranches
Private selBankBranch As bankbranch
Private myEmployees As HRCORE.Employees
''Private myEmployeeBankAccounts As HRCORE.EmployeeBankAccounts2

Private Sub cmdBack_Click()
    Me.sstBanks.TabEnabled(0) = True
    sstBanks.TabVisible(1) = False
    sstBanks.Tab = 0
End Sub

Private Sub cmdclose_Click()
Unload frmEmployeesOnBank
Unload Me
End Sub

Private Sub cmdDeleteBank_Click()
    Dim retVal As Long
    Dim resp As Long
    
    Select Case LCase(cmdDeleteBank.Caption)
        Case "delete"
        
        If Not currUser Is Nothing Then
        If currUser.CheckRight("BankDetails") <> secModify Then
            MsgBox "You dont have right to Delete the record. Please liaise with the security admin"
            Exit Sub
        End If
        
       
        End If
        
        
        Dim bankCanBeDeleted As Boolean
        bankCanBeDeleted = True
            If selBank Is Nothing Then
                MsgBox "Select the Bank you want to delete", vbExclamation, TITLES
            Else
                resp = MsgBox("Are yo sure you want to delete the selected Bank?", vbYesNo + vbQuestion, TITLES)
                If resp = vbYes Then
                
                '---------confirm that there are no employees currently linked to the bank to be deleted: Added by Kalya
                Dim i As Long
                Dim respo As Integer
                For i = 1 To myEmployeeBankAccounts.count
                If myEmployeeBankAccounts.Item(i).bankbranch.Bank.bankid = selBank.bankid Then
                respo = MsgBox("This bank cannot be deleted. Some Employees Are currently Attached to the bank. Want to See the List of Employees attached to this bank?", vbYesNo + vbCritical)
                
                If respo = 6 Then
                Me.MousePointer = vbHourglass
                 myReportPrinter
                Me.MousePointer = 1
                End If
                bankCanBeDeleted = False
                Exit For
                End If
                Next i
                
                '----------end confirmation
                
                If bankCanBeDeleted = False Then
                Exit Sub
                End If
                    retVal = selBank.Delete()
                If retVal = 0 Then
                Dim g As Integer
                For g = 1 To pBanks.count
                If pBanks.Item(g).bankid = selBank.bankid Then
                pBanks.remove (g)
                End If
                Next g
                End If
                    txtBankCode.Text = ""
                    txtBankName.Text = ""
                    LoadBanks
                End If
            End If
        Case "cancel"
            cmdNewBank.Enabled = True
            cmdEditBank.Caption = "Edit"
            cmdDeleteBank.Caption = "Delete"
            fraBank.Enabled = False
            cmdBranches.Enabled = True
            LoadBanks
    End Select
End Sub

Private Sub myReportPrinter()


               ''''''''''''''
                   If Not (currUser Is Nothing) Then
                   If currUser.CheckRight("AwardReport") = secNone Then
                   MsgBox "You Don't have right to view the report. Please liaise with security admin"
                   Exit Sub
                   End If
                   End If
                 'VIEW REPORT OF Employees On Report AWARDED TO EMPLOYEE
                   Set r = crtEmployeesOnBank
                  r.reportTitle = "LIST OF EMPLOYEES ATTACHED TO " & txtBankName.Text
                  r.ReportComments = "GROUPED BY: DEPARTMENTS"
                  
                  
                  
                  mySQL = "{EB.BankName} = '" & txtBankName.Text & "'"
                  printReport r
               
End Sub

Private Sub cmdDelete_Click()
    Dim retVal As Long
    Dim resp As Long
    
    Select Case LCase(cmdDelete.Caption)
        Case "delete"
            If selBankBranch Is Nothing Then
                MsgBox "Select the Bank Branch you want to delete", vbExclamation, TITLES
            Else
                resp = MsgBox("Are yo sure you want to delete the selected BankBranch?", vbYesNo + vbQuestion, TITLES)
                If resp = vbYes Then
                    retVal = selBankBranch.Delete()
                    LoadBranchesOfBank selBank, True
                End If
            End If
            
        Case "cancel"
            cmdNew.Enabled = True
            cmdEdit.Caption = "Edit"
            cmdDelete.Caption = "Delete"
            cmdBack.Enabled = True
            fraBranch.Enabled = False
            LoadBranchesOfBank selBank, False
    End Select
End Sub

Private Sub cmdBranches_Click()
         If txtBankCode.Text = "" Then
              MsgBox ("Choose the bank to view its branches")
         Exit Sub
         End If
    On Error GoTo ErrorHandler
    
    'clear
    ClearControlsBranch
    
    If selBank Is Nothing Then
        MsgBox "Select the Currency to set BankBranches for", vbInformation, TITLES
    Else
        Me.txtParentBank.Text = "(" & selBank.BankCode & ") " & selBank.BankName
        LoadBranchesOfBank selBank, False
        sstBanks.TabVisible(1) = True
        sstBanks.TabEnabled(0) = False
        sstBanks.Tab = 1
    End If
    
    Exit Sub
    
ErrorHandler:
End Sub

Private Sub cmdEditBank_Click()
    Select Case LCase(cmdEditBank.Caption)
        Case "edit"
            cmdNewBank.Enabled = False
            cmdEditBank.Caption = "Update"
            cmdDeleteBank.Caption = "Cancel"
            cmdBranches.Enabled = False
            fraBank.Enabled = True
            
        Case "update"
        If validateupdatebankcode = False Then Exit Sub
       
            If Update() = False Then Exit Sub
            cmdNewBank.Enabled = True
            cmdEditBank.Caption = "Edit"
            cmdDeleteBank.Caption = "Delete"
            cmdBranches.Enabled = True
            fraBank.Enabled = False
            LoadBanks
            
        Case "cancel"   'cancels a new operation
            cmdNewBank.Caption = "New"
            cmdEditBank.Caption = "Edit"
            cmdDeleteBank.Enabled = True
            fraBank.Enabled = False
            cmdBranches.Enabled = True
            
    End Select
End Sub

Private Sub cmdEdit_Click()
    Select Case LCase(cmdEdit.Caption)
        Case "edit"
            cmdNew.Enabled = False
            cmdEdit.Caption = "Update"
            cmdDelete.Caption = "Cancel"
            cmdBack.Enabled = False
            fraBranch.Enabled = True
            
        Case "update"
            If UpdateBranch() = False Then Exit Sub
            cmdNew.Enabled = True
            cmdEdit.Caption = "Edit"
            cmdDelete.Caption = "Delete"
            fraBranch.Enabled = False
            cmdBack.Enabled = True
            LoadBranchesOfBank selBank, True
            
        Case "cancel"   'cancels a new operation
            cmdNew.Caption = "New"
            cmdEdit.Caption = "Edit"
            cmdDelete.Enabled = True
            cmdBack.Enabled = True
            fraBranch.Enabled = False
            
    End Select
End Sub

Private Sub cmdNew_Click()
     Select Case LCase(cmdNew.Caption)
        Case "new"
            cmdNew.Caption = "Update"
            cmdEdit.Caption = "Cancel"
            cmdDelete.Enabled = False
            ClearControlsBranch
            fraBranch.Enabled = True
            
        Case "update"
            If InsertNewBranch() = False Then Exit Sub
            cmdNew.Caption = "New"
            cmdEdit.Caption = "Edit"
            cmdDelete.Enabled = True
            fraBranch.Enabled = False
            LoadBranchesOfBank selBank, True
    End Select
End Sub

Private Sub cmdNewBank_Click()
    Select Case LCase(cmdNewBank.Caption)
    
    
        Case "new"
        
        If Not currUser Is Nothing Then
        If currUser.CheckRight("BankDetails") <> secModify Then
            MsgBox "You dont have right to Add the record. Please liaise with the security admin"
            Exit Sub
        End If
        
       
        End If
        
            cmdNewBank.Caption = "Update"
            cmdEditBank.Caption = "Cancel"
            cmdDeleteBank.Enabled = False
            cmdBranches.Enabled = False
            ClearControls
            fraBank.Enabled = True
            
        Case "update"
        If validatebankcode = False Then Exit Sub
            If InsertNew() = False Then Exit Sub
            cmdNewBank.Caption = "New"
            cmdEditBank.Caption = "Edit"
            cmdDeleteBank.Enabled = True
            cmdBranches.Enabled = True
            fraBank.Enabled = False
            LoadBanks
    End Select
    
End Sub
Private Function validateupdatebankcode() As Boolean

If Not pBanks Is Nothing Then
Dim i As Integer
i = 1
Dim k As Integer
k = pBanks.count
While (i <= k)
    If UCase(pBanks.Item(i).BankCode) = UCase(txtBankCode.Text) And UCase(pBanks.Item(i).BankName) <> UCase(txtBankName.Text) Then
    MsgBox "The Bank Code already exists. ", vbCritical
         validateupdatebankcode = False
    Exit Function
    End If
i = i + 1
Wend
validateupdatebankcode = True
Else
validateupdatebankcode = True
End If

End Function
Private Function validatebankcode() As Boolean

If Not pBanks Is Nothing Then
Dim i As Integer
i = 1
Dim k As Integer
k = pBanks.count
While (i <= k)
    If UCase(pBanks.Item(i).BankCode) = UCase(txtBankCode.Text) Then
    MsgBox "The Bank Code already exists. ", vbCritical
         validatebankcode = False
    Exit Function
    End If
i = i + 1
Wend
validatebankcode = True
Else
validatebankcode = True
End If

End Function


Private Function Update() As Boolean
    
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    
    If Trim(Me.txtBankCode.Text) = "" Then
        MsgBox "Enter the Bank Code", vbExclamation, TITLES
        Me.txtBankCode.SetFocus
        Exit Function
    Else
        selBank.BankCode = Trim(Me.txtBankCode.Text)
    End If
    
    If Trim(Me.txtBankName.Text) = "" Then
        MsgBox "Enter the Bank Name", vbExclamation, TITLES
        Me.txtBankName.SetFocus
        Exit Function
    Else
        selBank.BankName = Trim(Me.txtBankName.Text)
    End If
   
    retVal = selBank.Update()
    If retVal = 0 Then
        MsgBox "The Bank has been UpdateBranch successfully", vbInformation, TITLES
        Update = True
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred while updating the Bank" & vbNewLine & err.Description, vbExclamation, TITLES
    Update = False
End Function

Private Function InsertNewBranch() As Boolean
    ''Dim newBranch As HRCORE.BankBranch
    Dim newBranch As bankbranch
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    
    Set newBranch = New bankbranch
    If Trim(Me.txtBranchCode.Text) = "" Then
        MsgBox "Enter the Branch Code", vbExclamation, TITLES
        Me.txtBranchCode.SetFocus
        Exit Function
    Else
        newBranch.BranchCode = Trim(Me.txtBranchCode.Text)
    End If
    
    If Trim(Me.txtBranchName.Text) = "" Then
        MsgBox "Enter the Branch Name", vbExclamation, TITLES
        Me.txtBranchName.SetFocus
        Exit Function
    Else
        newBranch.BranchName = Trim(Me.txtBranchName.Text)
    End If
    
 
    Set newBranch.Bank = selBank
    
    retVal = newBranch.InsertNew()
    If retVal = 0 Then
        MsgBox "The new Bank Branch has been added successfully", vbInformation, TITLES
        InsertNewBranch = True
            Dim ItemX As ListItem
        
        
           Set ItemX = lvwbankbranches.ListItems.add(, , newBranch.BranchCode)
            ItemX.SubItems(1) = newBranch.BranchName
            ItemX.Tag = newBranch.BankBranchID
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred while creating a new Bank Branch" & vbNewLine & err.Description, vbExclamation, TITLES
    InsertNewBranch = False
    
End Function

Private Function UpdateBranch() As Boolean
    
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
       
    If Trim(Me.txtBranchCode.Text) = "" Then
        MsgBox "Enter the Branch Code", vbExclamation, TITLES
        Me.txtBranchCode.SetFocus
        Exit Function
    Else
        selBankBranch.BranchCode = Trim(Me.txtBranchCode.Text)
    End If
    
    If Trim(Me.txtBranchName.Text) = "" Then
        MsgBox "Enter the Branch Name", vbExclamation, TITLES
        Me.txtBranchName.SetFocus
        Exit Function
    Else
        selBankBranch.BranchName = Trim(Me.txtBranchName.Text)
    End If
    
    Set selBankBranch.Bank = selBank
    
    retVal = selBankBranch.Update()
    If retVal = 0 Then
        MsgBox "The Bank Branch has been Updated successfully", vbInformation, TITLES
        UpdateBranch = True
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred while Updating the Bank Branch" & vbNewLine & err.Description, vbExclamation, TITLES
    UpdateBranch = False
    
End Function


Private Function InsertNew() As Boolean
    Dim NewBank As Bank
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    Set NewBank = New Bank
    
    If Trim(Me.txtBankCode.Text) = "" Then
        MsgBox "Enter the Bank Code", vbExclamation, TITLES
        Me.txtBankCode.SetFocus
        Exit Function
    Else
        NewBank.BankCode = Trim(Me.txtBankCode.Text)
    End If
    
    If Trim(Me.txtBankName.Text) = "" Then
        MsgBox "Enter the name of the Bank", vbExclamation, TITLES
        Me.txtBankCode.SetFocus
        Exit Function
    Else
        NewBank.BankName = Trim(Me.txtBankName.Text)
    End If
    
    
    retVal = NewBank.InsertNew()
    If retVal = 0 Then
        MsgBox "The New Bank has been added successfully", vbInformation, TITLES
        InsertNew = True
        pBanks.add NewBank
'            Dim itemx As ListItem
'            Set itemx = Me.lvwBanks.ListItems.add(, , newbank.BankCode)
'            itemx.SubItems(1) = newbank.BankName
'            itemx.Tag = newbank.BankID
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred while creating a new Bank" & vbNewLine & err.Description, vbExclamation, TITLES
    InsertNew = False
End Function


Private Sub ClearControls()
    Me.txtBankCode.Text = ""
    Me.txtBankName.Text = ""
End Sub

Private Sub Form_Load()
    frmMain2.PositionTheFormWithoutEmpList Me
    
'    Set pBanks = New HRCORE.Banks
'    Set pBankBranches = New BankBranches
'    pBankBranches.GetActiveBankBranches
    
    LoadBanks
  
    sstBanks.TabVisible(1) = False
End Sub

Private Sub LoadBranchesOfBank(ByVal TheBank As Bank, ByVal Refresh As Boolean)
    Dim TheBranches As BankBranches
    
    On Error GoTo ErrorHandler
    If Refresh = True Then
        pBankBranches.GetActiveBankBranches
    End If
    
    If Not TheBank Is Nothing Then
        Set TheBranches = pBankBranches.GetBranchesOfBankID(TheBank.bankid)
    End If
    
    PopulateBankBranches TheBranches
        
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while loading the Bank Branches" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub PopulateBankBranches(ByVal TheBranches As BankBranches)
    Dim i As Long
    Dim ItemX As ListItem
    
    On Error GoTo ErrorHandler
    
    Me.lvwbankbranches.ListItems.Clear
    
    If Not TheBranches Is Nothing Then
        For i = 1 To TheBranches.count
            Set ItemX = lvwbankbranches.ListItems.add(, , TheBranches.Item(i).BranchCode)
            ItemX.SubItems(1) = TheBranches.Item(i).BranchName
            ItemX.Tag = TheBranches.Item(i).BankBranchID
        Next i
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while populating Bank Branches" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub ClearControlsBranch()
    Me.txtBranchCode.Text = ""
    Me.txtBranchName.Text = ""
    ''lvwBankBranches.ListItems.Clear
    
End Sub


Private Sub SetFieldsBranch(ByVal TheBranch As bankbranch)
    ClearControlsBranch
    If Not (TheBranch Is Nothing) Then
        Me.txtBranchCode.Text = TheBranch.BranchCode
        Me.txtBranchName.Text = TheBranch.BranchName
    End If
End Sub

Private Sub LoadBanks()
    Set pBanks = New Banks
    pBanks.GetActiveBanks
    Set pBankBranches = New BankBranches
    pBankBranches.GetActiveBankBranches
    PopulateBanks pBanks
End Sub


Private Sub PopulateBanks(ByVal TheBanks As Banks)
    Dim i As Long
    Dim ItemX As ListItem
    
    On Error GoTo ErrorHandler
    Me.lvwBanks.ListItems.Clear
    If Not (TheBanks Is Nothing) Then
        For i = 1 To TheBanks.count
            Set ItemX = Me.lvwBanks.ListItems.add(, , TheBanks.Item(i).BankCode)
            ItemX.SubItems(1) = TheBanks.Item(i).BankName
            ItemX.Tag = TheBanks.Item(i).bankid
        Next i
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "An error has occurred while populating existing Banks" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub


Private Sub SetFields(ByVal TheBank As Bank)
    On Error GoTo ErrorHandler
    ClearControls
    If Not (TheBank Is Nothing) Then
        Me.txtBankCode.Text = TheBank.BankCode
        Me.txtBankName.Text = TheBank.BankName
    End If
    
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while populating the Bank details" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub lvwBanks_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set selBank = Nothing
    If IsNumeric(Item.Tag) Then
        Set selBank = pBanks.FindBankByID(CLng(Item.Tag))
    End If
    SetFields selBank
End Sub

Private Sub lvwBankBranches_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set selBankBranch = Nothing
    If IsNumeric(Item.Tag) Then
        Set selBankBranch = pBankBranches.FindBankBranchByID(CLng(Item.Tag))
    End If
    
    SetFieldsBranch selBankBranch
    
End Sub


''-----------added by kalya
Private Sub printReport(rpt As CRAXDDRT.Report)


   Dim conProps As CRAXDDRT.ConnectionProperties
    
    On Error GoTo ErrHandler
    'force crystal to use the basic syntax report
    
    frmMain2.MousePointer = vbHourglass
    
    If rpt.HasSavedData = True Then
        rpt.DiscardSavedData
    End If
    
    
    ' Loop through all database tables and set the correct server & database
        Dim tbl As CRAXDDRT.DatabaseTable
        Dim tbls As CRAXDDRT.DatabaseTables
        
        Set tbls = rpt.Database.Tables
        For Each tbl In tbls
            
            On Error Resume Next
            Set conProps = tbl.ConnectionProperties
            conProps.DeleteAll
            If tbl.DllName <> "crdb_ado.dll" Then
                tbl.DllName = "crdb_ado.dll"
            End If
              tbl.Name = "EB"
            conProps.add "Data Source", ConParams.DataSource
            conProps.add "Initial Catalog", ConParams.InitialCatalog
            conProps.add "Provider", "SQLOLEDB.1"
            'conProps.Add "Integrated Security", "true"
            conProps.add "User ID", ConParams.UserID
            conProps.add "Password", ConParams.Password
            conProps.add "Persist Security Info", "false"
            tbl.Location = tbl.Name
        Next tbl
        
        rpt.FormulaSyntax = crCrystalSyntaxFormula
        rpt.RecordSelectionFormula = mySQL
        
    With rpt
    '        DEALING WITH ALTERED PARAM FIELD OBJECT VALUES
        If blnAlterParamValue = True Then
            For lngLoopVariable = LBound(objParamField) To UBound(objParamField)
                For lngloopvariable2 = 1 To .ParameterFields.count
                    If .ParameterFields.Item(lngloopvariable2).ParameterFieldName = objParamField(lngLoopVariable).Name Then
                        .ParameterFields.Item(lngloopvariable2).SetCurrentValue (objParamField(lngLoopVariable).value)
                        Exit For
                    End If
                    
                Next
            Next
        End If
        .EnableParameterPrompting = False
    End With
    rpt.PaperSize = crPaperA4
    
    If rpt.PaperOrientation = crLandscape Then
        rpt.BottomMargin = 192
        rpt.RightMargin = 720
        rpt.LeftMargin = 58
        rpt.TopMargin = 192
    ElseIf rpt.PaperOrientation = crPortrait Then
        rpt.BottomMargin = 300
        rpt.RightMargin = 338
        rpt.LeftMargin = 300
        rpt.TopMargin = 281
    End If
        With frmReports.CRViewer1
            .DisplayGroupTree = False
            .EnableAnimationCtrl = False
            .ReportSource = rpt
            .ViewReport

        End With
       '' rpt.PrintOut False, 1, True, 1, 1
        
        
        
        formula = ""
     frmReports.Show vbModal
    Me.MousePointer = 0
      
       frmMain2.MousePointer = vbNormal
        Unload Me
    Exit Sub
ErrHandler:
    MsgBox err.Description, vbInformation
    frmMain2.MousePointer = vbNormal
End Sub

