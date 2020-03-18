VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSelUserGroup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search Engine-All User Groups"
   ClientHeight    =   3090
   ClientLeft      =   2775
   ClientTop       =   2385
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   3075
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
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
         Left            =   4200
         TabIndex        =   7
         Top             =   2640
         Width           =   1095
      End
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
         Left            =   120
         TabIndex        =   6
         Top             =   2640
         Width           =   1095
      End
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
         Left            =   1560
         TabIndex        =   5
         Top             =   240
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
         Left            =   120
         TabIndex        =   4
         Top             =   240
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
         Height          =   315
         Left            =   4080
         TabIndex        =   1
         Top             =   240
         Width           =   1200
      End
      Begin MSComctlLib.ListView lvwDetails 
         Height          =   2010
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   3545
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
         Left            =   2760
         TabIndex        =   3
         Top             =   240
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frmSelUserGroup"
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
                Set LI = lvwDetails.ListItems.Add(, , !GROUP_CODE & "", , 5)
                LI.ListSubItems.Add , , !GROUP_NAME & ""
                .MoveNext
            Loop
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    strName = ""
    Unload Me
End Sub

Private Sub cmdSelect_Click()
On Error GoTo 10
    strName = lvwDetails.SelectedItem
    Unload Me
Exit Sub
10: MsgBox Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo Hell
OptName.Value = True

CConnect.CColor Me, MyColor

Set rsCont = New Recordset
fraDetails.Visible = True

Set rsCont = CConnect.GetRecordSet("select * from tblUserGroup Where Subsystem = '" & SubSystem & "'")
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
        Set rsCont = CConnect.GetRecordSet("select * from tblUserGroup where  GROUP_CODE Like '" & X & "'" & " order by GROUP_CODE")
    Else
        Set rsCont = CConnect.GetRecordSet("select * from tblUserGroup where  GROUP_NAME Like '" & X & "'" & " order by GROUP_NAME")
    End If
    
    Call InitGrid
    Call DisplayRecords
    
Exit Sub
Hell: MsgBox Err.Description, vbCritical, "Search Records"
End Sub
