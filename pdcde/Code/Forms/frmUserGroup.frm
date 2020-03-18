VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmUserGroup 
   Caption         =   "User Group(s)"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   10398
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "User Group(s)"
      TabPicture(0)   =   "frmUserGroup.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCompanyID"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lsvUsers"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraMain"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame fraMain 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   720
         TabIndex        =   2
         Top             =   2040
         Visible         =   0   'False
         Width           =   3255
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
            Left            =   2400
            TabIndex        =   11
            Top             =   1080
            Width           =   735
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
            Left            =   1320
            TabIndex        =   10
            Top             =   1080
            Width           =   735
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
            Left            =   120
            TabIndex        =   9
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            TabIndex        =   5
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtCode 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            TabIndex        =   4
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblCaption 
            BackStyle       =   0  'Transparent
            Caption         =   "User Group Name:*"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblCaption 
            BackStyle       =   0  'Transparent
            Caption         =   "User Group Code:*"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackColor       =   &H00800000&
            Caption         =   "  Group Details"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   3255
         End
      End
      Begin MSComctlLib.ListView lsvUsers 
         Height          =   5415
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   9551
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Group Description"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label lblCompanyID 
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmUserGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Add_New As Boolean

Private Sub prLoadGroups()
'Dim rst As ADODB.Recordset
'Dim Lst As ListSubItems
'
'Set rst = cConnect.GetRecordSet("Select * From tblUserGroup Where Company_ID=" & Val(lblCompanyID.Caption))
'With rst
'    lsvUsers.ListItems.Clear
'        Do While Not .EOF
'                Set Lst = lsvUsers.ListItems.Add(, , !GROUP_ID)
'                    Lst.Add , , !GROUP_CODE
'                    Lst.Add , , !GROUP_NAME
'            .MoveNext
'        Loop
'End With
'rst.Close
'Set rst = Nothing
End Sub
Private Sub prInsertValues()
'//check if the user has entered all the required details
'//check if the code has been entered
If txtCode.Text = "" Then
    MsgBox "Please Ensure That The Group  Code Has Been Entered", vbInformation, "Group Code Missing"
    txtCode.SetFocus
    Exit Sub
End If
'//check if the name has been entered
If txtName.Text = "" Then
    MsgBox "Please Ensure That The Group Name Has Been Entered", vbInformation, "Group Name Missing"
    txtName.SetFocus
    Exit Sub
End If
If Add_New = True Or lsvUsers.ListItems.Count < 1 Then
    CConnect.GetRecordSet ("Insert Into tblUserGroup(Company_ID,Group_Code,Group_Name) Values(2,'" & txtCode.Text & "','" & txtName.Text & "')")
Else
    CConnect.GetRecordSet ("Update tblUserGroup Set Group_Code='" & txtCode.Text & "',Group_Name='" & txtName.Text & "' Where Group_ID=" & lsvUsers.SelectedItem.Text)
End If
    MsgBox "Database Update Complete", vbInformation, "Update Complete"
    Add_New = False
    fraMain.Visible = False
    prLoadGroups
End Sub

Private Sub cmdCancel_Click()
    fraMain.Visible = False
End Sub

Private Sub cmdNew_Click()
    Add_New = True
    txtCode.Text = ""
    txtName.Text = ""
End Sub

Private Sub cmdSave_Click()
    prInsertValues
End Sub

Private Sub Form_Load()
CConnect.CColor Me, MyColor
lblCompanyID.Caption = 2
End Sub

Private Sub lblCompanyID_Change()
If lblCompanyID.Caption = "" Then Exit Sub
    prLoadGroups
End Sub

Private Sub lsvUsers_DblClick()
If lsvUsers.ListItems.Count > 0 Then '//user has selected an item
    txtCode.Text = lsvUsers.SelectedItem.ListSubItems(1).Text
    txtName.Text = lsvUsers.SelectedItem.ListSubItems(2).Text
End If
    fraMain.Visible = True
End Sub
