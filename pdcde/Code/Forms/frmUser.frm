VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmUser 
   Caption         =   "User Setup Master"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5190
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   5190
   Begin VB.Frame fraDept 
      Caption         =   "Organisation Structure"
      Height          =   3615
      Left            =   120
      TabIndex        =   38
      Top             =   360
      Visible         =   0   'False
      Width           =   4935
      Begin MSComctlLib.TreeView trvOrg 
         Height          =   2775
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   4895
         _Version        =   393217
         HideSelection   =   0   'False
         Style           =   7
         Appearance      =   1
      End
      Begin VB.CommandButton cmdCloseD 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   40
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton cmdSelectD 
         Caption         =   "Select"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   39
         Top             =   3120
         Width           =   975
      End
   End
   Begin VB.Frame fraCat 
      Caption         =   "Employee Categories"
      Height          =   3615
      Left            =   120
      TabIndex        =   31
      Top             =   360
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   34
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton cmdCloseC 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   33
         Top             =   3120
         Width           =   975
      End
      Begin MSComctlLib.ListView lsvCat 
         Height          =   2775
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   7223
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "User Listing"
      TabPicture(0)   =   "frmUser.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblCompanyID"
      Tab(0).Control(1)=   "lblEmployeeID"
      Tab(0).Control(2)=   "LvAllUsers"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "User Details"
      TabPicture(1)   =   "frmUser.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Accounts listing"
      TabPicture(2)   =   "frmUser.frx":047A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "fraAccList"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "chkUnlock"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdSaveUnlock"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "chkLockAccount"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.CheckBox chkLockAccount 
         Caption         =   "&Lock selected account"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   3720
         Width           =   2175
      End
      Begin VB.CommandButton cmdSaveUnlock 
         Caption         =   "&Update [ Alt + U ]"
         Height          =   375
         Left            =   3360
         TabIndex        =   45
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CheckBox chkUnlock 
         Caption         =   "&Unlock selected account"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Frame fraAccList 
         Caption         =   "Account list:"
         Height          =   2895
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   4935
         Begin MSComctlLib.ListView lstAccList 
            Height          =   2535
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   4471
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   1
         Top             =   420
         Width           =   4935
         Begin VB.TextBox txtDept 
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
            IMEMode         =   3  'DISABLE
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   2400
            Width           =   2415
         End
         Begin VB.CommandButton cmdDept 
            Height          =   375
            Left            =   4080
            Picture         =   "frmUser.frx":0496
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   2400
            Width           =   375
         End
         Begin VB.CommandButton cmdCat 
            Height          =   375
            Left            =   4080
            Picture         =   "frmUser.frx":0598
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   2040
            Width           =   375
         End
         Begin VB.TextBox txtCat 
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
            IMEMode         =   3  'DISABLE
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   2040
            Width           =   2415
         End
         Begin VB.TextBox txtDepartCode 
            Height          =   285
            Left            =   5880
            TabIndex        =   26
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtEmpCode 
            Height          =   285
            Left            =   3480
            TabIndex        =   25
            Top             =   600
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtGroupCode 
            Height          =   285
            Left            =   3480
            TabIndex        =   24
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cmdUserGroup 
            Height          =   375
            Left            =   4440
            Picture         =   "frmUser.frx":069A
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtEmpName 
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
            Left            =   1560
            TabIndex        =   22
            Top             =   600
            Width           =   2895
         End
         Begin VB.TextBox txtDeptName 
            Height          =   285
            Left            =   7320
            TabIndex        =   21
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox txtGroupName 
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
            Left            =   1560
            TabIndex        =   20
            Top             =   240
            Width           =   2895
         End
         Begin VB.CommandButton cmdEmp 
            Height          =   375
            Left            =   4440
            Picture         =   "frmUser.frx":079C
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtConfirmPassword 
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
            IMEMode         =   3  'DISABLE
            Left            =   1560
            PasswordChar    =   "#"
            TabIndex        =   4
            Top             =   1680
            Width           =   2895
         End
         Begin VB.TextBox txtPassword 
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
            IMEMode         =   3  'DISABLE
            Left            =   1560
            PasswordChar    =   "#"
            TabIndex        =   3
            Top             =   1320
            Width           =   2895
         End
         Begin VB.TextBox txtUsername 
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
            Left            =   1560
            TabIndex        =   2
            Top             =   960
            Width           =   2895
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Department Access"
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
            Index           =   8
            Left            =   120
            TabIndex        =   37
            Top             =   2415
            Width           =   1395
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Category Access"
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
            Index           =   7
            Left            =   120
            TabIndex        =   28
            Top             =   2050
            Width           =   1215
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DepartCode"
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
            Index           =   6
            Left            =   5880
            TabIndex        =   27
            Top             =   720
            Width           =   870
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Group"
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
            Index           =   5
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   810
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Confirm Password"
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
            Index           =   4
            Left            =   120
            TabIndex        =   9
            Top             =   1680
            Width           =   1290
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
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
            Index           =   3
            Left            =   120
            TabIndex        =   8
            Top             =   1320
            Width           =   690
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Username"
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
            Index           =   2
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   720
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Full Name"
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
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   600
            Width           =   690
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Department"
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
            Index           =   0
            Left            =   7320
            TabIndex        =   5
            Top             =   720
            Width           =   1230
         End
      End
      Begin MSComctlLib.ListView LvAllUsers 
         Height          =   3120
         Left            =   -74880
         TabIndex        =   11
         Top             =   675
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   5503
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Group"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Employee"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Username"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   -74880
         TabIndex        =   10
         Top             =   3300
         Width           =   4935
         Begin VB.CommandButton CmdCancel 
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            TabIndex        =   19
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton CmdDelete 
            Caption         =   "Delete"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            TabIndex        =   18
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton CmdSave 
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
            Height          =   375
            Left            =   1920
            TabIndex        =   17
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton CmdNew 
            Caption         =   "New"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   960
            TabIndex        =   16
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Label lblEmployeeID 
         Height          =   375
         Left            =   -72360
         TabIndex        =   14
         Top             =   1740
         Width           =   975
      End
      Begin VB.Label lblCompanyID 
         Height          =   135
         Left            =   -72480
         TabIndex        =   13
         Top             =   900
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub LoadGrid()
With lstAccList
    .ColumnHeaders.Add , , "Username", .Width / 4
    .ColumnHeaders.Add , , "Full names", .Width - 3 * .Width / 4
    .ColumnHeaders.Add , , "Account status", .Width / 4
    .ColumnHeaders.Add , , "Reason", .Width / 4
End With
End Sub
Private Sub Make_Columns_AllUsers()
With LvAllUsers
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "UserName", 1200
    .ColumnHeaders.Add , , "Group ID", 1300
    .ColumnHeaders.Add , , "Employee Name", 2500
    .ColumnHeaders.Add , , "Max. Category Access", 2500
    .View = lvwReport
End With
End Sub

Private Sub Load_All_Users()
Dim rs As ADODB.Recordset
Dim listIt As ListItem
Make_Columns_AllUsers
Dim frozen As String
Set rs = CConnect.GetRecordSet("Select * From SECURITY WHERE subsystem = '" & SubSystem & "'")
With rs
    LvAllUsers.ListItems.Clear
        Do While Not .EOF
            Set LI = LvAllUsers.ListItems.Add(, , !UID & "")
                LI.ListSubItems.Add , , !GNo & ""
                LI.ListSubItems.Add , , !empname & ""
                LI.ListSubItems.Add , , Trim(!categoryAccess & "")
            .MoveNext
        Loop
    rs.Requery
    With lstAccList
        .ListItems.Clear
        Do While Not rs.EOF
            Set listIt = lstAccList.ListItems.Add(, , rs!UID & "")
            listIt.ListSubItems.Add , , rs!empname & ""
            If rs!frozen & "" = True Then frozen = "FROZEN" Else frozen = "ACTIVE"
            listIt.ListSubItems.Add , , frozen
            listIt.ListSubItems.Add , , rs!Reason_Frozen & ""
            listIt.Tag = rs!id & ""
            rs.MoveNext
        Loop
    End With
End With

rs.Close
Set rs = Nothing
End Sub

Private Sub CmdBrowse_Click()

End Sub

Public Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCat_Click()
    fraCat.Visible = True
End Sub

Private Sub cmdCloseC_Click()
    fraCat.Visible = False
End Sub

Private Sub cmdCloseD_Click()
    fraDept.Visible = False
End Sub

Public Sub cmdDelete_Click()
    '// If test box is empty then exit sub
    If Trim$(txtUsername) = "" Then
        MsgBox "Please specify the user you wish to delete.", vbInformation
        Exit Sub
    ElseIf Trim$(txtUsername) = CurrentUser Then
        MsgBox "You cannot delete the current user.", vbInformation
        Exit Sub
    End If
    If Trim$(txtGroupCode) = "Infiniti" Or Trim$(txtUsername) = "ADMIN" Then Exit Sub
    '// Confrim with the user before
    If MsgBox("Are you sure you want to delete " & txtUsername & "?", vbQuestion + vbYesNo) = vbYes Then
        CConnect.ExecuteSql ("DELETE FROM SECURITY where GNO='" & txtGroupCode & "' AND subsystem = '" & SubSystem & "' AND UID = '" & Trim$(txtUsername) & "'")
        Clear_Text
    End If
    Call Load_All_Users
    SSTab1.Tab = 0
End Sub

Private Sub cmdDept_Click()
    fraDept.Visible = True
End Sub

Public Sub cmdEdit_Click()

End Sub

Private Sub cmdEmp_Click()

On Error GoTo ErrorTrap
Dim rs As ADODB.Recordset
If txtGroupCode.Text <> "" Then

    strName = ""
    
    frmSelEmployees.Show vbModal
    
    txtEmpCode = strName
    If strName <> "" Then
        Set rs = CConnect.GetRecordSet("Select * From Employee where EmpCode='" & strName & "'")
        txtEmpName = rs!SurName & " " & rs!OtherNames
    End If
    
End If

Exit Sub
ErrorTrap:
MsgBox Err.Description, vbExclamation, "User Groups"

End Sub


Public Sub cmdNew_Click()
    Clear_Text
End Sub

Public Sub cmdSave_Click()
    Save_User_Details
    Load_All_Users
End Sub

Private Sub cmdSaveUnlock_Click()
If chkLockAccount.Value = 1 Then CConnect.ExecuteSql "UPDATE security SET frozen=1 WHERE id=" & lstAccList.SelectedItem.Tag
If chkUnlock.Value = 1 Then CConnect.ExecuteSql "UPDATE security SET frozen=0 WHERE id=" & lstAccList.SelectedItem.Tag
Load_All_Users
End Sub

Private Sub cmdSelect_Click()
    On Error GoTo errHandler
    txtCat.Text = lsvCat.SelectedItem.Text
    fraCat.Visible = False
    Exit Sub
errHandler:
End Sub

Private Sub cmdSelectD_Click()
On Error GoTo errHandler
    'txtDept.Text = trvOrg.SelectedItem.Key
    'fraCat.Visible = False
    trvOrg_DblClick
Exit Sub
errHandler:
End Sub

Private Sub cmdUserGroup_Click()
On Error GoTo ErrorTrap
Dim rs As ADODB.Recordset

    strName = ""
    frmSelUserGroup.Show vbModal
    If strName <> "" Then
        txtGroupCode = strName
        Set rs = CConnect.GetRecordSet("Select * From tblUserGroup where subsystem = '" & SubSystem & "' AND GROUP_CODE='" & strName & "'")
        txtGroupName = rs!GROUP_NAME
    End If
Exit Sub
ErrorTrap:
MsgBox Err.Description, vbExclamation, "User Groups"
End Sub

Private Sub Form_Load()
    Dim LI As ListItem
    With frmMain2
        Me.Move .tvwMain.Width + .lvwEmp.Width + (.tvwMain.Width / 36) * 2, (.Height / 5.52) + 155
    End With
    
    CConnect.CColor Me, MyColor
    Call LoadGrid
    Load_All_Users
    
    Set rs5 = CConnect.GetRecordSet("SELECT * FROM ECategory ORDER BY seq")
    lsvCat.ListItems.Clear
    While rs5.EOF = False
        Set LI = lsvCat.ListItems.Add(, , rs5!code & "")
        LI.ListSubItems.Add , , rs5!Comments & ""
        rs5.MoveNext
    Wend
    
    Call myStructure
End Sub

Function Load_Selected_User(sUserName As String)
Dim rs As ADODB.Recordset

Set rs = CConnect.GetRecordSet("Select * From SECURITY where UID='" & sUserName & "' AND subsystem = '" & SubSystem & "'")
    If Not rs.EOF Then '// If the record exist
        txtGroupCode = rs!GNo
        txtGroupName = rs!Description & ""
        txtEmpCode = rs!empcode & ""
        txtEmpName = rs!empname & ""
    '''    txtDepartCode = Rs!DeptCode & ""
        txtDept = rs!DeptCode & ""
        txtUsername = rs!UID
        txtPassword = frmLog.EEncryptPassword(rs!Pass) '// Load a decrytpted password into the textbox
        txtConfirmPassword = frmLog.EEncryptPassword(rs!Pass)  '// Load a decrypted version
        txtCat.Text = Trim(rs!categoryAccess & "")
    End If
Set rs = Nothing

End Function

Function Save_User_Details()
On Error GoTo ErrorTrap

Dim rs As ADODB.Recordset
Dim rsExpiry As New Recordset
Dim ExpiryDays As Integer

    If Trim$(txtUsername) = "" Then
        MsgBox "Please specify the username.", vbInformation
        Exit Function
    End If

    If Trim$(txtGroupCode) = "Infiniti" Or UCase(Trim$(txtUsername)) = "ADMIN" Then Exit Function

    CConnect.ExecuteSql ("DELETE From SECURITY where UID='" & txtUsername & "' AND subsystem = '" & SubSystem & "'")
    
    Set rsExpiry = CConnect.GetRecordSet("Select * From tblPasswordRule")
    
    If Not rsExpiry.EOF Then
        ExpiryDays = rsExpiry!Change_After
    Else
        ExpiryDays = 30 '// Default period in case non is set
    End If
    
    Set rs = CConnect.GetRecordSet("Select * From SECURITY WHERE subsystem = '" & SubSystem & "'")
    rs.AddNew
        rs!GNo = IIf(txtGroupName = "ADMIN", "ADMIN", txtGroupCode) '// If admin give full rights
        rs!Description = txtGroupName
        rs!empcode = txtEmpCode
        rs!empname = txtEmpName
        If txtCat.Text <> "" Then
            rs!categoryAccess = Trim(txtCat.Text)
        End If
        rs!DeptCode = txtDept.Text
        rs!SubSystem = SubSystem
        rs!UID = txtUsername
        rs!EDate = DateAdd("d", ExpiryDays, Date)  '// End date is the date from now plus the number of expiry period
        If txtConfirmPassword = "" Then MsgBox "Type Your Password in the 'Confirm Box'", vbExclamation, "Password Check": Exit Function
        If txtConfirmPassword = txtPassword Then
            rs!Pass = frmLog.EncryptPassword(txtPassword)
        Else
            MsgBox "Your Passwords do not match", vbExclamation, "Password Check": Exit Function
        End If
    '    rs!DateCreated = Date
    rs.Update
    Set rs = Nothing
    
    MsgBox "New user successfully added.", vbInformation, "Add new User."

Exit Function
ErrorTrap:
MsgBox Err.Description, vbExclamation
End Function

Private Sub Clear_Text()
    txtGroupCode = ""
    txtGroupName = ""
    txtEmpCode = ""
    txtEmpName = ""
    txtDeptName = ""
    txtUsername = ""
    txtPassword = ""
    txtConfirmPassword = ""
    txtCat.Text = ""
End Sub

Private Sub lstAccList_ItemClick(ByVal Item As MSComctlLib.ListItem)
If lstAccList.SelectedItem.ListSubItems(2).Text = "FROZEN" Then chkUnlock.Enabled = True: chkLockAccount.Enabled = False: chkLockAccount.Value = 0 Else chkUnlock.Enabled = False: chkLockAccount.Enabled = True: chkUnlock.Value = 0
End Sub

Private Sub lsvCat_DblClick()
    Call cmdSelect_Click
End Sub

Private Sub LvAllUsers_Click()
    Load_Selected_User LvAllUsers.SelectedItem
    SSTab1.Tab = 1
End Sub



Private Sub trvOrg_DblClick()
    txtDept.Text = trvOrg.SelectedItem.Tag
    Call cmdCloseD_Click
End Sub

Private Sub txtConfirmPassword_Change()
'    txtConfirmPassword.Text = UCase(txtConfirmPassword.Text)
    txtConfirmPassword.SelStart = Len(txtConfirmPassword.Text)
End Sub

Private Sub txtEmpName_Change()
    txtEmpName.Text = UCase(txtEmpName.Text)
    txtEmpName.SelStart = Len(txtEmpName.Text)
End Sub

Private Sub txtGroupName_Change()
    txtGroupName.Text = UCase(txtGroupName.Text)
    txtGroupName.SelStart = Len(txtGroupName.Text)
End Sub

Private Sub txtPassword_Change()
'    txtPassword.Text = UCase(txtPassword.Text)
    txtPassword.SelStart = Len(txtPassword.Text)
End Sub

Private Sub txtUsername_Change()
    txtUsername.Text = UCase(txtUsername.Text)
    txtUsername.SelStart = Len(txtUsername.Text)
End Sub

Public Sub myStructure()
    Dim rec_o As New ADODB.Recordset, rec_o1 As New ADODB.Recordset
    Dim MyNodes As Node
    trvOrg.Nodes.Clear
    
    On Error GoTo errHandler
    Set rec_o = CConnect.GetRecordSet("SELECT cName FROM generalopt WHERE subsystem = '" & SubSystem & "'")
    Set MyNodes = trvOrg.Nodes.Add(, , "All", rec_o!cName & "")
    MyNodes.Tag = "All"
    Set rec_o = CConnect.GetRecordSet("SELECT code, Description FROM stypes")
    While rec_o.EOF = False
        Set MyNodes = trvOrg.Nodes.Add("All", tvwChild, "XS" & rec_o!code, rec_o!Description & "")
        MyNodes.Tag = rec_o!code & ""
        Set rec_o1 = CConnect.GetRecordSet("SELECT * FROM CStructure WHERE SCode = '" & rec_o!code & "' ORDER BY MyLevel, Code")
    
        With rec_o1
            If .RecordCount > 0 Then
                .MoveFirst
                
                Do While Not .EOF
                    If !MyLevel = 0 Then
                        Set MyNodes = trvOrg.Nodes.Add("XS" & rec_o!code, tvwChild, !LCode, !code & ",  " & !Description & "")
                        MyNodes.EnsureVisible
                    Else
                        Set MyNodes = trvOrg.Nodes.Add(!PCode & "", tvwChild, !LCode & "", !code & ", " & !Description & "")
'                        MyNodes.EnsureVisible
                    End If
                    MyNodes.Tag = !LCode & ""
                    .MoveNext
                Loop
                .MoveFirst
            End If
        End With
        rec_o.MoveNext
    Wend
    Exit Sub
errHandler:
End Sub

