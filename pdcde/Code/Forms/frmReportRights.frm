VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmReportRights 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rights On Reports"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReportRights.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   6255
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraEditRights 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2475
      Left            =   720
      TabIndex        =   16
      Top             =   2280
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton cmdCanc2 
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
         Left            =   3825
         Picture         =   "frmReportRights.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Cancel Process"
         Top             =   1860
         Width           =   495
      End
      Begin VB.CommandButton cmdSave2 
         Caption         =   "Save"
         Default         =   -1  'True
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
         Left            =   3345
         Picture         =   "frmReportRights.frx":0DCC
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Save Record"
         Top             =   1860
         Width           =   495
      End
      Begin VB.TextBox txtModule 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   570
         Width           =   4305
      End
      Begin VB.OptionButton optModify 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Modify"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   135
         TabIndex        =   19
         Top             =   1320
         Width           =   1365
      End
      Begin VB.OptionButton optView 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1740
         TabIndex        =   18
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton optNone 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3375
         TabIndex        =   17
         Top             =   1320
         Width           =   765
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Module"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   23
         Top             =   315
         Width           =   1785
      End
   End
   Begin VB.Frame fraReports 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2280
      Left            =   6810
      TabIndex        =   8
      Top             =   1980
      Visible         =   0   'False
      Width           =   5970
      Begin VB.ComboBox cboRFilter 
         Height          =   315
         ItemData        =   "frmReportRights.frx":0ECE
         Left            =   135
         List            =   "frmReportRights.frx":0EF0
         TabIndex        =   14
         Top             =   1710
         Width           =   3390
      End
      Begin VB.CommandButton cmdCancel 
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
         Left            =   5325
         Picture         =   "frmReportRights.frx":0F57
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancel Process"
         Top             =   1605
         Width           =   495
      End
      Begin VB.CommandButton cmdSave 
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
         Left            =   4830
         Picture         =   "frmReportRights.frx":1059
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Save Record"
         Top             =   1605
         Width           =   510
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   135
         TabIndex        =   0
         Top             =   555
         Width           =   1230
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1485
         TabIndex        =   1
         Top             =   555
         Width           =   4335
      End
      Begin VB.TextBox txtReport 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   135
         TabIndex        =   2
         Top             =   1125
         Width           =   3645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Report Filter"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   1455
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   315
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Left            =   1485
         TabIndex        =   11
         Top             =   315
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Report Name"
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   885
         Width           =   945
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00800000&
         Caption         =   " Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   6435
      End
   End
   Begin MSComctlLib.ImageList imgTree 
      Left            =   735
      Top             =   1830
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportRights.frx":115B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportRights.frx":1E35
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportRights.frx":2B0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportRights.frx":37E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportRights.frx":44C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportRights.frx":519D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportRights.frx":5E77
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportRights.frx":6B51
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportRights.frx":782B
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportRights.frx":8505
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportRights.frx":91DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportRights.frx":9EB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportRights.frx":A1D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportRights.frx":AEAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportRights.frx":BB87
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportRights.frx":C861
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRCancel 
      Caption         =   "Cancel"
      Height          =   390
      Left            =   3870
      TabIndex        =   7
      Top             =   8175
      Width           =   1170
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   390
      Left            =   5130
      TabIndex        =   6
      Top             =   8190
      Width           =   1230
   End
   Begin MSComctlLib.TreeView tvwReports 
      Height          =   7575
      Left            =   -45
      TabIndex        =   5
      Top             =   60
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   13361
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   706
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "imgTree"
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
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Caption         =   "To View or Change User Right on report, double click on the report item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   0
      TabIndex        =   15
      Top             =   7680
      Width           =   6435
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Report"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmReportRights"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MyNodes As Node
Dim CNode As String
Dim rsMStruc As Recordset
Dim myReport As String

Public Sub InitReports()
Dim mm As String
Dim CNode As String
tvwReports.Nodes.Clear

Set MyNodes = tvwReports.Nodes.Add(, , "Q", "Reports", 14)
Set rsMStruc = CConnect.GetRecordSet("SELECT * FROM SReports ORDER BY MyLevel, Code")

With rsMStruc
    If .RecordCount > 0 Then
        .MoveFirst
        CNode = !LinkCode & ""
        Do While Not .EOF
            If !MyLevel = 0 Then
                If !ObjectID = "None" Then
                    Set MyNodes = tvwReports.Nodes.Add(, , !LinkCode, !Description & "", 14)
                Else
                    Set MyNodes = tvwReports.Nodes.Add(, , !LinkCode, !Description & "", 1)
                End If
                MyNodes.EnsureVisible
            Else
                If !ObjectID = "None" Then
                    Set MyNodes = tvwReports.Nodes.Add(!PreviousCode & "", tvwChild, !LinkCode & "", !Description & "", 14)
                Else
                    Set MyNodes = tvwReports.Nodes.Add(!PreviousCode & "", tvwChild, !LinkCode & "", !Description & "", 1)
                End If
                    
                MyNodes.EnsureVisible
            End If
            
            .MoveNext
        Loop
               
        .MoveFirst
    End If
End With

'tvwReports.Nodes("Q005").Expanded = False
'tvwReports.Nodes("Q001").Expanded = False
'tvwReports.Nodes("Q010001").Expanded = False
'tvwReports.Nodes("Q010005").Expanded = False
'tvwReports.Nodes("Q010010").Expanded = False
'tvwReports.Nodes("Q010015").Expanded = False
'tvwReports.Nodes("Q010020").Expanded = False

tvwReports.Refresh

End Sub



Private Sub cboRFilter_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
'''    fraReports.Visible = False
End Sub

Private Sub cmdRCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
'''Dim LCode As String
'''Dim PCode As String
'''Dim MyLevel As Integer
'''Dim rsM1 As Recordset
'''
''''Prompt if required fields are not entered
'''If CNode = "" Then
'''    MsgBox "You must select the node you want to add to.", vbInformation
'''    Exit Sub
'''End If
'''
'''If txtCode.Text = "" Then
'''    MsgBox "Enter the code.", vbExclamation
'''    txtCode.SetFocus
'''    Exit Sub
'''End If
'''
'''If txtDescription.Text = "" Then
'''    MsgBox "Enter the Description.", vbExclamation
'''    txtDescription.SetFocus
'''    Exit Sub
'''End If
'''
'''If txtReport.Text = "" Then
'''    MsgBox "Enter the report name.", vbExclamation
'''    txtReport.SetFocus
'''    Exit Sub
'''End If
'''
'''If cboRFilter.Text = "" Then
'''    MsgBox "Select report filter.", vbExclamation
'''    cboRFilter.SetFocus
'''    Exit Sub
'''End If
'''
'''If SaveNew = True Then
'''    Set rsM1 = cConnect.GetRecordSet("SELECT * FROM SReports WHERE LinkCode = '" & CNode & "" & txtCode.Text & "'")
'''
'''    With rsM1
'''        If .RecordCount > 0 Then
'''            MsgBox "Code already exists. Enter another one.", vbInformation
'''            txtCode.Text = ""
'''            txtCode.SetFocus
'''            Exit Sub
'''        End If
'''    End With
'''
'''    Set rsM1 = Nothing
'''
'''
'''
'''    Set rsM1 = cConnect.GetRecordSet("SELECT * FROM SReports ORDER BY MyLevel, Code")
'''
'''    With rsM1
'''        If .RecordCount > 0 Then
'''            .MoveFirst
'''            .Find "LinkCode like '" & CNode & "'", , adSearchForward, adBookmarkFirst
'''            If Not .EOF Then
'''                LCode = !LinkCode & "" & txtCode.Text
'''                MyLevel = !MyLevel + 1
'''                PCode = !LinkCode & ""
'''            Else
'''                LCode = "Q" & txtCode.Text
'''                MyLevel = 0
'''            End If
'''        Else
'''            LCode = "Q" & txtCode.Text
'''            MyLevel = 0
'''        End If
'''    End With
'''
'''    Set rsM1 = Nothing
'''
'''
'''    cConnect.ExecuteSql ("DELETE FROM SReports WHERE LinkCode = '" & LCode & "'")
'''
'''    cConnect.ExecuteSql ("INSERT INTO SReports (LinkCode, PreviousCode, MyLevel, Code, Description, ObjectID, RFilter)" & _
'''                " VALUES('" & LCode & "','" & PCode & "'," & MyLevel & ",'" & txtCode.Text & "','" & txtDescription.Text & "','" & txtReport.Text & "','" & cboRFilter.Text & "')")
'''
'''    Call InitReports
'''    Cleattxt
'''    txtCode.SetFocus
'''Else
'''    cConnect.ExecuteSql ("UPDATE SReports SET Description = '" & txtDescription.Text & "', ObjectID = '" & txtReport.Text & "', RFilter = '" & cboRFilter.Text & "' WHERE LinkCode = '" & CNode & "'")
'''    Call InitReports
'''    fraReports.Visible = False
'''End If
   
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdSave2_Click()
If optModify.Value = True Then GroupRight = "MODIFY"
If optView.Value = True Then GroupRight = "VIEW"
If optNone.Value = True Then GroupRight = "NONE"

UpDate_Group_Rights CNode, GroupRight

MsgBox "Group Right Assigned Successifuly", vbInformation, "Group Right Assignement"

fraEditRights.Visible = False
End Sub

Function UpDate_Group_Rights(sReport As String, sAssignedRight As String)
On Error GoTo Hell

Set rs = CConnect.GetRecordSet("Select * From tblAssignedRights where GROUP_ID= ('" & sGroupID & "') and MODULE_ID = ('" & sModuleID & "')")

With rs
    If .EOF = False Then
        strQ = "Update tblAssignedRights set Assigned_Rights='" & GroupRight & "' where GROUP_ID= ('" & sGroupID & "') and MODULE_ID = ('" & sModuleID & "') "
        Action = "ASSIGNED REPORT RIGHTS; RIGHT: " & GroupRight & "; REPORT NAME: " & !MODULE_NAME & "; GROUP ID: " & sGroupID
        CConnect.ExecuteSql strQ
    End If
End With

Set rs = Nothing
   
Exit Function
Hell:
End Function

Private Sub cmdView_Click()
On Error GoTo errHandler
    With rsMStruc
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "LinkCode like '" & CNode & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                If !ObjectID <> "None" Then
                    ReportType = !ObjectID & ""
                    RFilter = !RFilter & ""
                    If RFilter = "None" Then
                        Me.MousePointer = vbHourglass
                        Set a = New Application
                        Set R = a.OpenReport(App.Path & "\HR Base Reports\" & "" & ReportType)
                         
                        R.ReadRecords
                    
                        With frmReports.CRViewer1
                            .ReportSource = R
                            .ViewReport
                        End With
                        
                        frmReports.Show vbModal
                        Me.MousePointer = 0
    
                    Else
                        frmRange.Show vbModal, Me
                    End If
                    
                End If
            End If
            
        End If
    End With

Exit Sub
errHandler:
    MsgBox Err.Description, vbInformation
    Me.MousePointer = 0

End Sub

Private Sub Form_Load()
    
    Call InitReports

End Sub

Private Sub mnuAdd_Click()
    If CNode = "" Then
        MsgBox "You must select the node you want to add to.", vbInformation
        Exit Sub
    End If
    
    Call Cleattxt
    SaveNew = True
    fraReports.Visible = True
    txtCode.Locked = False
    txtCode.SetFocus
    
End Sub

Private Sub mnuDelete_Click()
    If CNode = "" Then
        MsgBox "You must select the node you want to delete.", vbInformation
        Exit Sub
    End If
    
    CConnect.ExecuteSql ("DELETE FROM SReports WHERE LinkCode = '" & CNode & "'")
    Call InitReports
End Sub

Private Sub mnuEdit_Click()
    If CNode = "" Then
        MsgBox "You must select the node you want to edit.", vbInformation
        Exit Sub
    End If
    
    With rsMStruc
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "LinkCode like '" & CNode & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                fraReports.Visible = True
                txtCode.Text = !code & ""
                txtDescription.Text = !Description & ""
                txtReport.Text = !ObjectID & ""
                cboRFilter.Text = !RFilter & ""
                txtCode.Locked = True
                SaveNew = False
            End If
        End If
    End With
    
        
End Sub

Private Sub tvwReports_DblClick()
    Call cmdView_Click
End Sub

Private Sub tvwReports_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    Me.PopupMenu mnuReport
End If
End Sub

Private Sub tvwReports_NodeClick(ByVal Node As MSComctlLib.Node)
    If tvwReports.Nodes.Count > 0 Then
        CNode = tvwReports.SelectedItem.Key
    End If
End Sub

Public Sub Cleattxt()
txtCode.Text = ""
txtDescription.Text = ""
txtReport.Text = ""
End Sub
