VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmAccounts 
   Caption         =   "User Accounts"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8640
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdActivate 
      Caption         =   "Activate"
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmdFreeze 
      Caption         =   "Freeze"
      Height          =   375
      Left            =   7200
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
   End
   Begin MSComctlLib.ListView lvwDetails 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   6376
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
Attribute VB_Name = "frmAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private RsT As New ADODB.Recordset
Private Const Reason_Frozen As String = "Frozen by Administration"
Private X As Integer

Private Sub cmdActivate_Click()
For X = 1 To lvwDetails.ListItems.Count
    If lvwDetails.ListItems(X).Checked = True Then
    Activate_Account lvwDetails.ListItems(X).Text
    End If
Next
Load_lvwDetails
End Sub

Private Sub cmdFreeze_Click()
For X = 1 To lvwDetails.ListItems.Count
    If lvwDetails.ListItems(X).Checked = True Then
    Freeze_Account lvwDetails.ListItems(X).Text
    End If
Next
Load_lvwDetails
End Sub

Private Sub Form_Load()
CConnect.CColor Me, MyColor
Set RsT = CConnect.GetRecordSet("SELECT SECURITY.UID as UserID,SECURITY.FROZEN,SECURITY.REASON_FROZEN,SECURITY.FROZEN_COUNT,TBLUSERGROUP.GROUP_NAME as UserGroup FROM SECURITY INNER JOIN TBLUSERGROUP ON SECURITY.GNO=TBLUSERGROUP.GROUP_CODE WHERE security.subsystem = '" & SubSystem & "' AND (tblUserGroup.SubSystem = '" & SubSystem & "')")
Load_lvwDetails
End Sub

Public Sub Load_lvwDetails()
Dim LI As ListItem
On Error GoTo errHandler
RsT.Requery
With lvwDetails
    .ColumnHeaders.Clear
    .ListItems.Clear
    
    .ColumnHeaders.Add , , "User ID", 1000
    .ColumnHeaders.Add , , "User Group", 2500
    .ColumnHeaders.Add , , "Frozen", 1000
    .ColumnHeaders.Add , , "Reason Frozen", 2655
    .ColumnHeaders.Add , , "Frozen Count", 1500
    
    '*****************************************'
    If RsT.RecordCount = 0 Then Exit Sub
    RsT.MoveFirst
    For X = 0 To RsT.RecordCount - 1
    Set LI = .ListItems.Add(, , RsT!UserID & "")
    LI.ListSubItems.Add , , RsT!UserGroup & ""
    LI.ListSubItems.Add , , RsT!frozen & ""
    LI.ListSubItems.Add , , RsT!Reason_Frozen & ""
    LI.ListSubItems.Add , , RsT!frozen_count + 0
    RsT.MoveNext
    Next
    
    RsT.MoveFirst
End With
errHandler:
End Sub


Public Sub Activate_Account(UserID As String)
Dim rsa As New ADODB.Recordset
Set rsa = CConnect.GetRecordSet("SELECT * FROM Security WHERE UID='" & UserID & "' AND subsystem = '" & SubSystem & "'")
If rsa.RecordCount = 0 Then Exit Sub
rsa!frozen = False
rsa!frozen_count = rsa!frozen_count + 1
rsa!Reason_Frozen = ""
rsa.Update
End Sub

Public Sub Freeze_Account(UserID As String)
Dim rsa As New ADODB.Recordset
Set rsa = CConnect.GetRecordSet("SELECT * FROM Security WHERE UID='" & UserID & "' AND subsystem = '" & SubSystem & "'")
If rsa.RecordCount = 0 Then Exit Sub
rsa!frozen = True
rsa!Reason_Frozen = Reason_Frozen
rsa.Update
End Sub

Private Sub lvwDetails_DblClick()
'if lvwdetails.SelectedItem=nothing
End Sub
