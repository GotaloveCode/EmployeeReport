VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSelAddGroups 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search Engine-Add User Groups"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   80
      TabIndex        =   10
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4480
      TabIndex        =   9
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox txtUserGroup 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1800
      TabIndex        =   8
      Top             =   480
      Width           =   3735
   End
   Begin VB.OptionButton chkCustom 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Custom User Group"
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   120
      Width           =   2655
   End
   Begin VB.OptionButton chkDept 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Dept User Group "
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   2295
   End
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   4155
      Left            =   0
      TabIndex        =   0
      Top             =   -80
      Visible         =   0   'False
      Width           =   5655
      Begin VB.OptionButton OptName 
         BackColor       =   &H00C0E0FF&
         Caption         =   "By Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton OptCode 
         BackColor       =   &H00C0E0FF&
         Caption         =   "By Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtSearch 
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
         Left            =   3960
         TabIndex        =   1
         Top             =   1080
         Width           =   1560
      End
      Begin MSComctlLib.ListView lvwDetails 
         Height          =   2130
         Left            =   60
         TabIndex        =   2
         Top             =   1440
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   3757
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgTree"
         ForeColor       =   16711680
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
         NumItems        =   0
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "User Group Name"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "Search Field"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   2640
         TabIndex        =   3
         Top             =   1080
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frmSelAddGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCont As Recordset
Dim X   '++This is the search field++

Private Sub InitGrid()
    With lvwDetails
        .ColumnHeaders.Add , , "GROUP CODE", 1500
        .ColumnHeaders.Add , , "GROUP NAME", 3000
        .View = lvwReport
    End With
End Sub

Public Sub DisplayRecords()
    lvwDetails.ListItems.Clear
    With rsCont
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Set LI = lvwDetails.ListItems.Add(, , !code & "", , 5)
                LI.ListSubItems.Add , , !Description & ""
                .MoveNext
            Loop
        End If
    End With
End Sub

Private Sub chkCustom_Click()
txtUserGroup.Enabled = chkCustom.Value
fraDetails.Enabled = Not chkCustom.Value

End Sub

Private Sub chkDept_Click()
txtUserGroup.Enabled = chkCustom.Value
fraDetails.Enabled = Not chkCustom.Value

End Sub

Private Sub cmdCancel_Click()
    strName = ""
    Unload Me
End Sub

Private Sub cmdSelect_Click()
On Error GoTo 10
    If chkDept.Value = True Then
    strName = lvwDetails.SelectedItem
    Else
        If txtUserGroup = "" Then
        MsgBox "Please Input the Custom User Group", vbInformation, "User Group"
        Exit Sub
        End If
    strName = txtUserGroup
    frmSecurity.GrpName = txtUserGroup
    End If
    Unload Me
Exit Sub
10: MsgBox Err.Description
End Sub

Private Sub Form_Load()
    On Error GoTo Hell
    
    CConnect.CColor Me, MyColor
    
    OptName.Value = True
    
    Set rsCont = New Recordset
    fraDetails.Visible = True
    
    Set rsCont = CConnect.GetRecordSet("select * from cstructure ")
    Call InitGrid
    Call DisplayRecords
    Set rsCont = Nothing
    
    OptName.Value = True
    
    Exit Sub
Hell: MsgBox Err.Description, vbCritical, "Search Records"
End Sub

Private Sub lvwDetails_Click()
    cmdSelect.Enabled = True
End Sub

Private Sub lvwDetails_DblClick()
    Call cmdSelect_Click
End Sub

Private Sub OptCode_Click()
    OptName.Value = False
End Sub

Private Sub OptName_Click()
    OptCode.Value = False
End Sub

Private Sub txtSearch_Change()
On Error GoTo Hell
Set rsCont = New Recordset

    X = txtSearch & "%"
    
    If OptCode.Value = True Then
        Set rsCont = CConnect.GetRecordSet("select * from cstructure where  CODE Like '" & X & "'" & " order by CODE")
    Else
        Set rsCont = CConnect.GetRecordSet("select * from cstructure where  DESCRIPTION Like '" & X & "'" & " order by DESCRIPTION")
    End If
    
    Call InitGrid
    Call DisplayRecords
    
Exit Sub
Hell: MsgBox Err.Description, vbCritical, "Search Records"
End Sub

Private Sub txtUserGroup_Change()
If txtUserGroup <> "" Then
cmdSelect.Enabled = True
Else
cmdSelect.Enabled = False
End If
End Sub
