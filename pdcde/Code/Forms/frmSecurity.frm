VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSecurity 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "System Security Settings"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   9315
   Icon            =   "frmSecurity.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   16
      Top             =   6240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3120
      TabIndex        =   15
      Top             =   6240
      Visible         =   0   'False
      Width           =   855
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
      Left            =   8760
      Picture         =   "frmSecurity.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Delete Record"
      Top             =   6240
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
      Left            =   8760
      Picture         =   "frmSecurity.frx":0934
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Edit Record"
      Top             =   6240
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
      Left            =   8280
      Picture         =   "frmSecurity.frx":0A36
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Add New record"
      Top             =   6240
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.Frame Frame3 
         Height          =   5895
         Left            =   3720
         TabIndex        =   3
         Top             =   120
         Width           =   5415
         Begin MSComctlLib.ListView lsvLog 
            Height          =   5655
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   9975
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
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
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Key             =   "ID"
               Text            =   "ID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Key             =   "UserID"
               Text            =   "UserID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Key             =   "Username"
               Text            =   "Username"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Key             =   "Date"
               Text            =   "Date"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Key             =   "Time"
               Text            =   "Time"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Key             =   "EventType"
               Text            =   "Event Type"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Key             =   "EventDescription"
               Text            =   "Event Description"
               Object.Width           =   8819
            EndProperty
         End
         Begin MSComctlLib.ListView lsvGroups 
            Height          =   5655
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   9975
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Group Code"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Group Description"
               Object.Width           =   7937
            EndProperty
         End
         Begin MSComctlLib.ListView lsvSettings 
            Height          =   5655
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   9975
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Type"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Policy Setting"
               Object.Width           =   7937
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Setting"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView lsvUsers 
            Height          =   5655
            Left            =   120
            TabIndex        =   10
            Top             =   120
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   9975
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Username"
               Object.Width           =   7937
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Height          =   5895
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   3495
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   195
            Left            =   1200
            TabIndex        =   14
            Top             =   6960
            Width           =   1095
         End
         Begin TabDlg.SSTab stbTree 
            Height          =   5655
            Left            =   120
            TabIndex        =   4
            Top             =   120
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   9975
            _Version        =   393216
            Style           =   1
            Tabs            =   1
            TabsPerRow      =   1
            TabHeight       =   520
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Security Policy Settings"
            TabPicture(0)   =   "frmSecurity.frx":0B38
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lblCompanyID"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "trTree"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            Begin MSComctlLib.TreeView trTree 
               Height          =   5175
               Left            =   120
               TabIndex        =   5
               Top             =   360
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   9128
               _Version        =   393217
               HideSelection   =   0   'False
               LabelEdit       =   1
               LineStyle       =   1
               Style           =   7
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
            Begin VB.Label lblCompanyID 
               Height          =   135
               Left            =   2160
               TabIndex        =   6
               Top             =   360
               Width           =   135
            End
         End
      End
      Begin MSComDlg.CommonDialog CdlgMain 
         Left            =   3480
         Top             =   5640
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " Security Details"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   120
         Width           =   6255
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Enabled         =   0   'False
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear Event Log"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public GrpName As String

Private Sub prLoadTree()
With trTree.Nodes
        .Add , , "Security", "System Security Policy"
        .Add "Security", tvwChild, "PASSWORD", "Password Policies"
        .Add "Security", tvwChild, "ACCOUNT", "Account Policies"
        .Add "Security", tvwChild, "USERGROUPS", "User Group(s)"
        .Add "Security", tvwChild, "USERS", "System User(s)"
        .Add "Security", tvwChild, "GROUPRIGHTS", "Group Right(s)"
        '.Add "Security", tvwChild, "USERRIGHTS", "User Right(s)"
'        .Add "Security", tvwChild, "ACTIVEUSER", "Active User(s)"
'        .Add "Security", tvwChild, "SYSTEMLOG", "System Log"
End With
End Sub
Private Sub prCheckNode(ndNode As Node)
Dim rs As ADODB.Recordset
Dim Lst As ListSubItems

On Error GoTo errHandler
lsvSettings.ListItems.Clear '//clear the list item from the details
With ndNode
Select Case .Key
    Case "PASSWORD"
        Set rs = CConnect.GetRecordSet("Select * From tblPasswordRule")
        lsvGroups.Visible = False
        lsvSettings.Visible = True
        lsvUsers.Visible = False
        lsvLog.Visible = False
        cmdNew.Enabled = False: cmdDelete.Enabled = False
        '//set the details for the password policy
            Set Lst = lsvSettings.ListItems.Add(, , "")
                Lst.Add , , "Length"
                Lst.Add , , "Minimum password length"
                Lst.Add , , IIf(rs.RecordCount < 1, "0", rs!Minimum_Length & "")
            Set Lst = lsvSettings.ListItems.Add(, , "")
                Lst.Add , , "Age"
                Lst.Add , , "Minimum password age"
                Lst.Add , , IIf(rs.RecordCount < 1, "0", rs!Change_After)
            Set Lst = lsvSettings.ListItems.Add(, , "")
                Lst.Add , , "History"
                Lst.Add , , "Minimum passwords remembered"
                Lst.Add , , IIf(rs.RecordCount < 1, "0", rs!Password_History & "")
    Case "ACCOUNT"
        Set rs = CConnect.GetRecordSet("Select * From tblPasswordRule")
        lsvGroups.Visible = False
        lsvSettings.Visible = True
        lsvUsers.Visible = False
        lsvLog.Visible = False
        cmdNew.Enabled = False: cmdDelete.Enabled = False
            Set Lst = lsvSettings.ListItems.Add(, , "")
                Lst.Add , , "After"
                Lst.Add , , "Maximum tries before account lockout"
                Lst.Add , , IIf(rs.RecordCount < 1, "0", rs!Lockout_Threshold & "")
            Set Lst = lsvSettings.ListItems.Add(, , "")
                Lst.Add , , "Lockout"
                Lst.Add , , "Account lockout threshold"
                Lst.Add , , IIf(rs.RecordCount < 1, "0", rs!Lock_OutTime & "")
    
    Case "USERGROUPS"
        lsvGroups.Visible = True
        lsvSettings.Visible = False
        lsvUsers.Visible = False
        lsvLog.Visible = False
        cmdNew.Enabled = True: cmdDelete.Enabled = True
        prLoadGroups
        
    Case "USERS"
        lsvGroups.Visible = Not True
        lsvSettings.Visible = False
        lsvUsers.Visible = Not False
        lsvLog.Visible = False
        cmdNew.Enabled = False: cmdDelete.Enabled = False
        prLoadUsers
        
    Case "GROUPRIGHTS"
        'frmAssignUserRights.Show vbModal
        cmdNew.Enabled = False: cmdDelete.Enabled = False
        frmUGroups.Show vbModal
        
    Case "ACTIVEUSER"
'        lsvGroups.Visible = Not True
'        lsvSettings.Visible = False
'        lsvUsers.Visible = Not False
'        lsvLog.Visible = False
'        prLoadActiveUsers
        
    Case "SYSTEMLOG"
'        lsvGroups.Visible = Not True
'        lsvSettings.Visible = False
'        lsvUsers.Visible = Not True
'        lsvLog.Visible = True
'        prLoadSystemLog
        
End Select
End With
Exit Sub
errHandler:
End Sub

Private Sub prLoadSystemLog()
Dim rs As ADODB.Recordset
Dim Lst As ListSubItems

Set rs = CConnect.GetRecordSet("Select U.Username,L.* From tblUser U,tblUserLog L Where U.User_ID=L.User_ID")
With rs
    lsvLog.ListItems.Clear
        Do While Not .EOF
                Set Lst = lsvLog.ListItems.Add(, , !Log_ID)
                    Lst.Add , , !User_ID
                    Lst.Add , , !UserName
                    Lst.Add , , Format(!Log_Date, "dd-MMM-yyyy")
                    Lst.Add , , Format(!Log_Time, "hh:mm:ss")
                    Lst.Add , , UCase(!Log_Type)
                    Lst.Add , , LCase(!Log_Description)
            .MoveNext
        Loop
End With
rs.Close
Set rs = Nothing
End Sub
Private Sub prLoadGroups()
Dim rs As ADODB.Recordset

With lsvGroups
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "GROUP CODE", 1500
    .ColumnHeaders.Add , , "GROUP NAME", 4000
    .View = lvwReport
End With

Set rs = CConnect.GetRecordSet("Select * From tblUserGroup WHERE Subsystem = '" & SubSystem & "'")
With rs
    lsvGroups.ListItems.Clear
        Do While Not .EOF
            Set LI = lsvGroups.ListItems.Add(, , !GROUP_CODE & "")
                LI.ListSubItems.Add , , !GROUP_NAME & ""
            .MoveNext
        Loop
End With
Set rs = Nothing
End Sub

Private Sub Make_Columns_AllUsers()
With lsvUsers
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "UserName", 1200
    .ColumnHeaders.Add , , "Group ID", 1300
    .ColumnHeaders.Add , , "Employee Name", 2500
    .ColumnHeaders.Add , , "Max. Category Access", 2500
    .View = lvwReport
End With
End Sub

Private Sub prLoadUsers()
Dim rs As ADODB.Recordset

Make_Columns_AllUsers

Set rs = CConnect.GetRecordSet("Select * From SECURITY WHERE subsystem= '" & SubSystem & "'")
With rs
    lsvUsers.ListItems.Clear
        Do While Not .EOF
            Set LI = lsvUsers.ListItems.Add(, , !UID & "")
                LI.ListSubItems.Add , , !GNo & ""
                LI.ListSubItems.Add , , !empname & ""
                LI.ListSubItems.Add , , Trim(!categoryAccess & "")
            .MoveNext
        Loop
End With

rs.Close
Set rs = Nothing
End Sub
Private Sub prLoadActiveUsers()
Dim rs As ADODB.Recordset
Dim Lst As ListSubItems


Set rs = CConnect.GetRecordSet("Select tblUser.* From tblUser,tblUserGroup Where Status Like 'LOGGED IN' and tblUser.Group_ID=tblUserGroup.Group_ID and tblUserGroup.Company_ID=" & Val(lblCompanyID.Caption))
With rs
    lsvUsers.ListItems.Clear
    Do While Not .EOF
            Set Lst = lsvUsers.ListItems.Add(, , !employee_id)
                Lst.Add , , !User_ID
                Lst.Add , , !UserName
        .MoveNext
    Loop
End With
rs.Close
Set rs = Nothing
End Sub

Public Sub cmdCancel_Click()

End Sub

Public Sub cmdDelete_Click()
Dim LstCount As Integer, DelRec As Integer
DelRec = 0
If lsvGroups.ListItems.Count > 0 Then

    If MsgBox("Are you sure you want to delete the records?", vbQuestion + vbYesNo) = vbYes Then
        
        For LstCount = 1 To lsvGroups.ListItems.Count
            If lsvGroups.ListItems(LstCount).Checked = True Then
                If UCase(lsvGroups.ListItems(LstCount)) = "ADMIN" Then
                    MsgBox "System user groups cannot be deleted!", vbInformation, "User Group Management"
                Else
                    strQ = ("DELETE FROM tblUserGroup WHERE GROUP_CODE = '" & lsvGroups.ListItems(LstCount) & "'")
                    Action = "DELETED USER GROUP; GROUP_CODE: " & lsvGroups.ListItems(LstCount)
                    CConnect.ExecuteSql strQ
                    DelRec = DelRec + 1
                End If
            End If
        Next LstCount
    
        If DelRec > 0 Then
            MsgBox DelRec & " User Groups have been deleted!", vbInformation, "User Group Management"
        Else
            MsgBox "No User Groups were deleted!", vbInformation, "User Group Management"
        End If
        prLoadGroups
        
    End If
     
End If
End Sub

Public Sub cmdEdit_Click()

End Sub

Public Sub cmdNew_Click()

    strName = ""
        frmSelAddGroups.Show vbModal
        
    If strName <> "" Then
    
        Set rs = CConnect.GetRecordSet("Select * From tblUserGroup where GROUP_CODE='" & strName & "' AND subsystem = '" & SubSystem & "'")
        
        If rs.EOF = False Then
            MsgBox "User Group already exists", vbCritical, "User Group Management"
            Exit Sub
        End If
        
        Set rs2 = CConnect.GetRecordSet("Select * From cstructure where CODE='" & strName & "'")
        
        If rs2.EOF = False Then GrpName = rs2!Description & ""
        
        rs.AddNew
            rs!GROUP_CODE = strName
            rs!GROUP_NAME = GrpName
            rs!SubSystem = SubSystem
        rs.Update
        
        MsgBox "User Group " & GrpName & " Added!", vbInformation, "User Group Management"
        
    End If
    
    prLoadGroups

End Sub


Public Sub cmdSave_Click()

End Sub

Private Sub Form_Load()

Decla.Security Me

With frmMain2
    Me.Move .tvwMain.Width + .tvwMain.Width / 32.75, (.Height / 5.52) '- 155
End With

'''If oSmart.hRatio > 1.1 Then
'''    With frmMain2
'''        Me.Move .tvwMain.Width + .tvwMain.Width / 32.75, (.Height / 5.52)'- 155
'''    End With
'''Else
'''     With frmMain2
'''        Me.Move .tvwMain.Width + .tvwMain.Width / 32.75, .Height / 5.55
'''    End With
'''
'''End If

CConnect.CColor Me, MyColor

    prLoadTree
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain2.Caption = "Personnel Director " & App.FileDescription
End Sub

Private Sub lsvGroups_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lsvGroups
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lsvGroups_DblClick()
'''    frmUserGroup.lblCompanyID.Caption = lblCompanyID.Caption
'''    frmUserGroup.Show vbModal
End Sub

Private Sub lsvLog_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
With lsvLog
    .SortKey = ColumnHeader.Index - 1
    .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    .Sorted = True
End With
End Sub

Private Sub lsvSettings_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lsvSettings
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lsvSettings_DblClick()
If lsvSettings.ListItems.Count < 1 Then Exit Sub
Select Case UCase(lsvSettings.SelectedItem.ListSubItems(1).Text)
    Case "LENGTH"
        frmPasswordLength.lblCompanyID.Caption = lblCompanyID.Caption
        frmPasswordLength.Show vbModal
    Case "AGE"
        frmPasswordAge.lblCompanyID.Caption = lblCompanyID.Caption
        frmPasswordAge.Show vbModal
    Case "HISTORY"
        frmPasswordHistory.lblCompanyID.Caption = lblCompanyID.Caption
        frmPasswordHistory.Show vbModal
    Case "AFTER"
        frmAccountLockout.lblCompanyID.Caption = lblCompanyID.Caption
        frmAccountLockout.Show vbModal
    Case "LOCKOUT"
'        frmAccountLockTime.lblCompanyID.Caption = lblCompanyID.Caption
'        frmAccountLockTime.Show vbModal
End Select
End Sub

Private Sub lsvUsers_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lsvUsers
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lsvUsers_DblClick()
    frmUser.Show 'vbModal
End Sub

Private Sub mnuClear_Click()
Dim IntC As Integer
Dim iFile As Integer
Dim sPath As String
Dim sText As String

With CdlgMain
    .ShowSave
    '//check the file name
    If .FileName = "" Then Exit Sub
        sPath = .FileName & ".evt"
End With
If Len(Dir(sPath)) <> 0 Then
    If MsgBox("The Selected File Name Already Exists " & vbCrLf & "Do  You Wish To Overwrite It?", vbExclamation + vbYesNo, "File Exists") = vbYes Then
        FileCopy sPath, "C:\" & "EventLog" & Day(Date) & "F" & Month(Date) & "F" & Year(Date) & ".evt"
        
        Kill sPath
    End If
End If
    iFile = FreeFile
'//open the file
Open sPath For Append As #iFile

'//loop thru the listview
For IntC = 1 To lsvLog.ListItems.Count
    '//write the line
    Print #iFile, lsvLog.ListItems(IntC).Text & "," & lsvLog.ListItems(IntC).ListSubItems(1).Text & "," & lsvLog.ListItems(IntC).ListSubItems(2).Text & "," & lsvLog.ListItems(IntC).ListSubItems(3).Text & "," & lsvLog.ListItems(IntC).ListSubItems(4).Text & "," & lsvLog.ListItems(IntC).ListSubItems(5).Text & "," & lsvLog.ListItems(IntC).ListSubItems(6).Text
Next
lsvLog.ListItems.Clear
Close #iFile
SetAttr sPath, vbHidden
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub trTree_DblClick()
    prCheckNode trTree.SelectedItem
End Sub

Public Sub disabletxt()
Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
            i.Locked = True
        End If
    Next i

    
End Sub

