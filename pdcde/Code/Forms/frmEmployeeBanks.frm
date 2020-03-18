VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmEmployeeBanks 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   Caption         =   "Employee Banks"
   ClientHeight    =   7815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
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
            Picture         =   "frmEmployeeBanks.frx":0000
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeeBanks.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeeBanks.frx":0224
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeeBanks.frx":0336
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Employee banks"
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
      Height          =   3480
      Left            =   645
      TabIndex        =   16
      Top             =   1440
      Visible         =   0   'False
      Width           =   6060
      Begin VB.CheckBox chkMain 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         Top             =   3000
         Width           =   3495
      End
      Begin VB.CommandButton cmdSchBank 
         Height          =   315
         Left            =   5640
         Picture         =   "frmEmployeeBanks.frx":0878
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   600
         Width           =   315
      End
      Begin VB.CommandButton cmdSchBankBranch 
         Height          =   315
         Left            =   5640
         Picture         =   "frmEmployeeBanks.frx":0C02
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1190
         Width           =   315
      End
      Begin VB.TextBox txtBranchName 
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
         Left            =   120
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1170
         Width           =   5415
      End
      Begin VB.TextBox txtBankName 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   5415
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
         Height          =   450
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   2370
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
         Picture         =   "frmEmployeeBanks.frx":0F8C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Save Record"
         Top             =   2880
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
         Picture         =   "frmEmployeeBanks.frx":108E
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Cancel Process"
         Top             =   2865
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
         Height          =   330
         Left            =   135
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1725
         Width           =   5790
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
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
         Top             =   960
         Width           =   930
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
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
         BackColor       =   &H00C0E0FF&
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
         Top             =   2115
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
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
         Top             =   1515
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
      TabIndex        =   15
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
      Picture         =   "frmEmployeeBanks.frx":1190
      Style           =   1  'Graphical
      TabIndex        =   14
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
      Picture         =   "frmEmployeeBanks.frx":1682
      Style           =   1  'Graphical
      TabIndex        =   13
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
      Picture         =   "frmEmployeeBanks.frx":1784
      Style           =   1  'Graphical
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      Height          =   7800
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   13758
      LabelEdit       =   1
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
      NumItems        =   0
   End
End
Attribute VB_Name = "frmEmployeeBanks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Update_Emp_MainAcct()
    If SelectedEmployee Is Nothing Then Exit Sub
     
    rs1.Requery
    With rs1
        .Filter = "mainacct=1 and employee_id='" & SelectedEmployee.EmployeeID & "'"
        If .recordcount > 0 Then
        Dim cmd As New ADODB.Command
        cmd.ActiveConnection = ObjConnection
        cmd.CommandText = "UPDATE employee SET bankcode='" & Trim(!Bank_Code) & "',bankname='" & Trim(!Bank_Name) & "',bankbranch='" & Trim(!BankBranch_Code) & "',bankbranchname='" & Trim(!BANKBRANCH_NAME) & "',accountno='" & Trim(!AccNumber) & "' WHERE employee_id='" & SelectedEmployee.EmployeeID & "'"
        cmd.Execute
        ''cConnect.ExecuteSql "UPDATE employee SET bankcode='" & Trim(!Bank_Code) & "',bankname='" & Trim(!Bank_Name) & "',bankbranch='" & Trim(!BankBranch_Code) & "',bankbranchname='" & Trim(!BANKBRANCH_NAME) & "',accountno='" & Trim(!AccNumber) & "' WHERE employee_id='" & SelectedEmployee.EmployeeID & "'"
        End If
    End With

End Sub
Public Sub ClearText()
    With txtBankName
        .Text = ""
        .Tag = ""
    End With
    With txtBranchName
        .Text = ""
        .Tag = ""
    End With
    With txtAccountNumber
        .Text = ""
        .Tag = ""
    End With
    With txtAccountType
        .Text = ""
        .Tag = ""
    End With
    chkMain.value = 0
End Sub

Public Sub DisplayRecords()
    If Not SelectedEmployee Is Nothing Then
        Set rs1 = cConnect.GetRecordset("SELECT EmployeeBanks.EmployeeBankID,EmployeeBanks.Employee_ID,EmployeeBanks.BranchID,EmployeeBanks.AccNumber,EmployeeBanks.AccType," _
        & " EmployeeBanks.MainAcct ,tblBankBranch.BANKBRANCH_NAME,tblBankBranch.BANKBRANCH_CODE,tblBank.BANK_NAME,tblBank.BANK_CODE" _
        & " FROM  EmployeeBanks LEFT OUTER JOIN tblBankBranch ON EmployeeBanks.BranchID = tblBankBranch.BANKBRANCH_ID LEFT OUTER JOIN" _
        & " tblBank ON tblBankBranch.BANK_ID = tblBank.BANK_ID ORDER BY tblBank.BANK_NAME ASC")
        
        lvwDetails.ListItems.Clear
        Call ClearText

        With rs1
            If .recordcount > 0 Then
                .Filter = "employee_id like '" & SelectedEmployee.EmployeeID & "'"
            End If
        End With

        With rs1
            If .recordcount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    Set li = lvwDetails.ListItems.Add(, , !Bank_Name & "", , 5)
                    li.SubItems(1) = !BranchId & ""
                    li.SubItems(2) = !BANKBRANCH_NAME & ""
                    li.SubItems(3) = Trim(!AccNumber & "")
                    li.SubItems(4) = (!AccType & "")
                    If !MainAcct = True Then
                        li.SubItems(5) = "True" 'IIf(!mainacct = True, True, "")
                    Else
                        li.SubItems(5) = ""
                    End If
                   
                    li.Tag = !EmployeeBankId
                    .MoveNext
                Loop
                .MoveFirst
            End If
        End With
            'rs2.Filter = adFilterNone
    End If
   ' End With

End Sub

Public Sub InitGrid()
With lvwDetails
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "Bank name", .Width - ((2 * .Width / 5) + .Width / 6 + .Width / 7)
    .ColumnHeaders.Add , , "Branchcode", 0
    .ColumnHeaders.Add , , "Branch name", .Width / 5
    .ColumnHeaders.Add , , "Account number", .Width / 6
    .ColumnHeaders.Add , , "Account Type", .Width / 5
    .ColumnHeaders.Add , , "Main Account", .Width / 7
    .View = lvwReport
End With
End Sub

Public Sub cmdCancel_Click()
'If PromptSave = True Then
'    If MsgBox("Close this window?", vbYesNo + vbQuestion, "Confirm Close") = vbNo Then Exit Sub
'End If
'fraDetails.Visible = False
'With frmMain2
'    .cmdNew.Enabled = True
'    .cmdEdit.Enabled = True
'    .cmdDelete.Enabled = True
'    .cmdCancel.Enabled = False
'    .cmdSave.Enabled = False
'End With
Call EnableCmd
End Sub

Public Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdDelete_Click()
    If Not currUser Is Nothing Then
        If currUser.CheckRight("EmployeeBanks") <> secModify Then
            MsgBox "You dont have right to delete the record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    If Not SelectedEmployee Is Nothing Then
    
        If PromptSave = True Then If MsgBox("Are you sure you want to delete the selected record?", vbYesNo + vbQuestion, "Confirm delete") = vbNo Then Exit Sub
        
        If lvwDetails.ListItems.Count > 0 Then
            Action = "DELETED EMPLOYEE BANKS; EMPLOYEE CODE: " & SelectedEmployee.empcode & "; BRANCH: " & lvwDetails.SelectedItem.SubItems(2) & "; BANK NAME: " & lvwDetails.SelectedItem.Text & "; ACCOUNT NUMBER: " & lvwDetails.SelectedItem.SubItems(3)
            
            ''cConnect.ExecuteSql "Delete from Employeebanks Where EmployeeBankID=" & lvwDetails.SelectedItem.Tag
            ObjConnection.ExecuteSql "Delete from Employeebanks Where EmployeeBankID=" & lvwDetails.SelectedItem.Tag
            
            lvwDetails.ListItems.Remove lvwDetails.SelectedItem.index
        End If
    End If
End Sub

Public Sub cmdEdit_Click()
    
     If Not currUser Is Nothing Then
        If currUser.CheckRight("EmployeeBanks") <> secModify Then
            MsgBox "You dont have right to modify the record. Please liaise with the security admin"
            Exit Sub
        End If
    End If

    Dim RsT As New ADODB.Recordset
    ''Set RsT = cConnect.GetRecordset("select * from tblBank where bank_name like '" & lvwDetails.SelectedItem.Text & "'")
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = ObjConnection
    cmd.CommandText = "select * from tblBank where bank_name like '" & lvwDetails.SelectedItem.Text & "'"
    Set RsT = cmd.Execute
     
    If RsT.recordcount > 0 Then txtBankName.Tag = Trim(RsT!bank_id & "")
    txtBankName.Text = lvwDetails.SelectedItem.Text
    
    With txtBranchName
        .Text = lvwDetails.SelectedItem.ListSubItems(2)
        .Tag = lvwDetails.SelectedItem.ListSubItems(1) 'store the branch code
    End With
    txtAccountNumber.Tag = lvwDetails.SelectedItem.Tag 'store the ID od the table to be edited
    txtAccountNumber.Text = lvwDetails.SelectedItem.ListSubItems(3)
    txtAccountType.Text = lvwDetails.SelectedItem.ListSubItems(4)
    If lvwDetails.SelectedItem.ListSubItems(5) = "True" Then
        chkMain.value = 1
    Else
        chkMain.value = 0
    End If
    fraDetails.Visible = True
    cmdCancel.Enabled = True
    cmdSave.Enabled = True
End Sub

Public Sub cmdNew_Click()
    'check category access rights
    If Not currUser Is Nothing Then
        If currUser.CheckRight("EmployeeBanks") <> secModify Then
            MsgBox "You dont have right to add new record. Please liaise with the security admin"
            Exit Sub
        End If
    End If

    Call DisableCmd
    
    txtBankName.Text = ""
    txtBranchName.Text = ""
    txtAccountType.Text = ""
    txtAccountNumber.Text = ""
    fraDetails.Visible = True
    cmdCancel.Enabled = True
    SaveNew = True
    cmdSave.Enabled = True
    txtBankName.SetFocus
    
End Sub

Private Sub cmdSave_Click()
    Dim RsT As New ADODB.Recordset
    Dim cmd As ADODB.Command
    If txtBankName.Text = "" Then
        MsgBox "Enter the bank name.", vbExclamation
        txtBankName.SetFocus
        Exit Sub
    End If
    
    If txtBranchName.Text = "" Then
        MsgBox "Enter the bank branch.", vbExclamation
        txtBranchName.SetFocus
        Exit Sub
    End If
    Dim myMain ' As Boolean
    
    If PromptSave = True Then
        If MsgBox("Are you sure you want to save the record.", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
          
    If chkMain.value = 1 Then
        myMain = 1
       '' cConnect.ExecuteSql "UPDATE EmployeeBanks SET MainAcct=0 WHERE employee_id='" & SelectedEmployee.EmployeeID & "'"
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = ObjConnection
        cmd.CommandText = "UPDATE EmployeeBanks SET MainAcct=0 WHERE employee_id='" & SelectedEmployee.EmployeeID & "'"
        cmd.Execute
    Else
        myMain = 0
    End If
    
    'check for update or insert
    If Me.txtAccountNumber.Tag = "" Then
        mysql = "INSERT INTO EmployeeBanks (employee_id, branchID, AccNumber, AccType,MainAcct)" & _
                            " VALUES('" & SelectedEmployee.EmployeeID & "','" & txtBranchName.Tag & "'," & _
                            "'" & txtAccountNumber.Text & "','" & txtAccountType.Text & "'," & myMain & ")"
        Action = "REGISTERED EMPLOYEE BANKS; EMPLOYEE CODE: " & SelectedEmployee.empcode & "; BRANCH: " & txtBranchName.Text & "; BANK NAME: " & txtBankName.Text & "; IS THE MAIN ACCOUNT: " & IIf(chkMain.value = 1, "Yes", "No")
        
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = ObjConnection
        cmd.CommandText = mysl
        cmd.Execute
        ''cConnect.ExecuteSql (mysql)
    Else
         mysql = "UPDATE EmployeeBanks SET employee_id=" & SelectedEmployee.EmployeeID & ", branchID=" & txtBranchName.Tag & ",AccNumber='" & txtAccountNumber.Text & "',AccType='" & txtAccountType.Text & "',MainAcct=" & myMain & " WHERE EmployeeBankID= " & txtAccountNumber.Tag
         Action = "UDATED EMPLOYEE BANKS; EMPLOYEE CODE: " & SelectedEmployee.empcode & "; BRANCH: " & txtBranchName.Text & "; BANK NAME: " & txtBankName.Text & "; IS THE MAIN ACCOUNT: " & IIf(chkMain.value = 1, "Yes", "No")
         ''cConnect.ExecuteSql (mysql)
         
         Set cmd = New ADODB.Command
         cmd.ActiveConnection = ObjConnection
         cmd.CommandText = mysql
         cmd.Execute
    End If
    
   Me.txtAccountNumber.Tag = ""
   
    rs1.Requery
    
    With rs1
        If .recordcount > 0 Then
            If Not .EOF Then
                .MoveNext
                If Not .EOF Then
                    Call DisplayRecords
                    'Call DispDetails
                    'txtContact.SetFocus
                    Call Decla.DisableCmd
                Else
                    Call DisplayRecords
                    PSave = True
                    Call cmdCancel_Click
                    PSave = False
                End If
            Else
                Call DisplayRecords
                PSave = True
                Call cmdCancel_Click
                PSave = False
            End If
        Else
            Call DisplayRecords
            PSave = True
            Call cmdCancel_Click
            PSave = False
        End If
    End With
    'the main purpose of this procedure not yet known
    ' Call Update_Emp_MainAcct
   '************************************************************
   
'    With frmMain2
'        .cmdCancel.Enabled = False
'        .cmdNew.Enabled = True
'        .cmdEdit.Enabled = True
'        .cmdDelete.Enabled = True
'        .cmdSave.Enabled = False
'    End With
End Sub

Private Sub cmdSchBank_Click()
    With txtBranchName
        .Tag = ""
        .Text = ""
    End With
    frmSelBanks.show vbModal
    txtBankName.Tag = strName
    txtBankName.Text = strBankName
End Sub

Private Sub cmdSchBankBranch_Click()
    SelectedBank = txtBankName.Tag
    frmSelBankBranches.show vbModal
    txtBranchName.Text = strBranchName
    txtBranchName.Tag = strBranchID
End Sub

Private Sub Form_Load()
     
'    oSmart.FReset Me
'
'    If oSmart.hRatio > 1.1 Then
'        With frmMain2
'            Me.Move .tvwMain.Width + .lvwEmp.Width + (.tvwMain.Width / 36) * 2, (.Height / 5.52) '- 155
'        End With
'    Else
'         With frmMain2
'            Me.Move .tvwMain.Width + .lvwEmp.Width + (.tvwMain.Width / 36#) * 2, .Height / 5.55
'        End With
'
'    End If
    
    ''cConnect.CColor Me, MyColor
    
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    
    Call InitGrid
    
    Call DisplayRecords
    
    cmdFirst.Enabled = False
    cmdPrevious.Enabled = False

End Sub

Private Sub Form_Resize()
'    oSmart.FResize Me
'
'    Me.Height = tvwMainheight - 150
'    'Frame1.Move Frame1.Left, 0, Frame1.Width, tvwMainheight
'    lvwDetails.Move lvwDetails.Left, 0, lvwDetails.Width, tvwMainheight - 150
End Sub

Private Sub Form_Unload(Cancel As Integer)
''frmMain2.Caption = "Personnel Director " & App.FileDescription
End Sub
Private Sub lvwDetails_DblClick()
    Call cmdEdit_Click
End Sub
