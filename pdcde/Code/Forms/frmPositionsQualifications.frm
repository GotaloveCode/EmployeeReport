VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmPositionsQualifications 
   BorderStyle     =   0  'None
   Caption         =   "Positions Qualifications"
   ClientHeight    =   6960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraHold 
      Caption         =   "Positions Minimum Qualifications"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9975
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
         Left            =   2520
         Picture         =   "frmPositionsQualifications.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Add New record"
         Top             =   5040
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
         Left            =   3000
         Picture         =   "frmPositionsQualifications.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Edit Record"
         Top             =   5040
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
         Left            =   3480
         Picture         =   "frmPositionsQualifications.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Delete Record"
         Top             =   5040
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame fraDetails 
         Height          =   2415
         Left            =   2685
         TabIndex        =   3
         Top             =   1800
         Visible         =   0   'False
         Width           =   4695
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
            Left            =   4095
            Picture         =   "frmPositionsQualifications.frx":06F6
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Cancel Process"
            Top             =   1800
            Width           =   495
         End
         Begin VB.CommandButton cmdSave 
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
            Left            =   3600
            Picture         =   "frmPositionsQualifications.frx":07F8
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Save Record"
            Top             =   1800
            Width           =   510
         End
         Begin VB.TextBox txtQualificationParameterVal 
            Height          =   285
            Left            =   120
            TabIndex        =   5
            Top             =   1440
            Width           =   4455
         End
         Begin VB.ComboBox cboParameter 
            Height          =   315
            Left            =   120
            TabIndex        =   4
            Top             =   840
            Width           =   4455
         End
         Begin VB.Label lblPosition 
            AutoSize        =   -1  'True
            Caption         =   "Qualification parameter"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1980
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Qualification parameter value"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   1200
            Width           =   2055
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Qualification parameter"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   600
            Width           =   1620
         End
      End
      Begin MSComctlLib.ListView lvwDetails 
         Height          =   2850
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   9810
         _ExtentX        =   17304
         _ExtentY        =   5027
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgTree"
         ForeColor       =   0
         BackColor       =   -2147483643
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
      Begin MSComctlLib.ListView lvwQualification 
         Height          =   3570
         Left            =   120
         TabIndex        =   2
         Top             =   3240
         Width           =   9810
         _ExtentX        =   17304
         _ExtentY        =   6297
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgTree"
         ForeColor       =   0
         BackColor       =   -2147483643
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
End
Attribute VB_Name = "frmPositionsQualifications"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rsShowRecords As New ADODB.Recordset
Public Sub LoadCbo()
Dim rsLoadCbo As New ADODB.Recordset
Set rsLoadCbo = CConnect.GetRecordSet("select * from positionrequirements")
With rsLoadCbo
    If .RecordCount > 0 Then
        While .EOF = False
            cboParameter.AddItem !positionrequirementsdescription & ""
            .MoveNext
        Wend
    End If
End With
End Sub

Private Sub cboParameter_Click()
Dim rsLD As New ADODB.Recordset
Set rsLD = CConnect.GetRecordSet("select * from positionrequirements where positionrequirementsdescription like '" & cboParameter.Text & "'")
If rsLD.RecordCount > 0 Then
    cboParameter.Tag = rsLD!positionrequirementsid & ""
End If
End Sub

Private Sub cboParameter_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Public Sub cmdCancel_Click()
'fraDetails.Visible = False
'With frmMain2
'    .cmdSave.Enabled = False
'    .cmdNew.Enabled = True
'    .cmdEdit.Enabled = True
'    .cmdCancel.Enabled = False
'    .cmdDelete.Enabled = False
'End With
If PromptSave = True Then
    If MsgBox("Close this window?", vbYesNo + vbQuestion, "Confirm Close") = vbNo Then Exit Sub
End If
fraDetails.Visible = False
With frmMain2
    .cmdNew.Enabled = True
    .cmdEdit.Enabled = True
    .cmdDelete.Enabled = True
    .cmdCancel.Enabled = False
    .cmdSave.Enabled = False
End With
Call EnableCmd
End Sub

Public Sub cmdDelete_Click()
If MsgBox("Are you sure you want to delete " & UCase(lvwQualification.SelectedItem.Text) & vbCrLf & " which is attached to " & UCase(lvwDetails.SelectedItem.ListSubItems(1).Text) & " position?", vbYesNo + vbQuestion, "Confirm delete") = vbNo Then Exit Sub
Action = "DELETED A POSITION QUALIFICATION; POSITION: " & lvwDetails.SelectedItem.ListSubItems(1).Text & "; REQUIREMENT: " & lvwQualification.SelectedItem.Text & "; VALUE: " & lvwQualification.SelectedItem.ListSubItems(1).Text
CConnect.ExecuteSql ("delete from positionrequirementsvalue where positionrequirementsID=" & lvwQualification.SelectedItem.Tag & " and positionID=" & lvwDetails.SelectedItem.Tag & "")
'lvwQualification.ListItems.Remove lvwQualification.SelectedItem.Index
Call DisplayRecordsQualifications
End Sub

Public Sub cmdEdit_Click()
If lvwQualification.ListItems.Count = 0 Then MsgBox "There are no records to edit.", vbOKOnly + vbInformation, "No records": Exit Sub
fraDetails.Visible = True
With cboParameter
    .Text = lvwQualification.SelectedItem.Text
    .Tag = lvwQualification.SelectedItem.Tag
    .Locked = True
End With
lblPosition.Caption = "Position: " & UCase(lvwDetails.SelectedItem.ListSubItems(1).Text)
txtQualificationParameterVal.Text = lvwQualification.SelectedItem.ListSubItems(1).Text
End Sub

Public Sub cmdNew_Click()
If lvwDetails.ListItems.Count = 0 Then MsgBox "There are no positions defined.", vbOKOnly + vbInformation, "No Positions": Exit Sub
lblPosition.Caption = "Position: " & UCase(lvwDetails.SelectedItem.ListSubItems(1).Text)
fraDetails.Visible = True

With cboParameter
    .Text = ""
    .Tag = ""
    .Locked = False
End With
txtQualificationParameterVal.Text = ""
End Sub

Public Sub cmdSave_Click()
Dim rsLoad As New ADODB.Recordset
If txtQualificationParameterVal.Text = "" Then MsgBox "Please supply the qualifications parameters.", vbOKOnly + vbInformation, "Missing parameters": cboParameter.SetFocus: Exit Sub
If cboParameter.Text = "" Then MsgBox "Please supply a value corresponding to the selected parameter.", vbOKOnly + vbInformation, "Missing value": txtQualificationParameterVal.SetFocus: Exit Sub
If lvwDetails.ListItems.Count = 0 Then Exit Sub

Set rsLoad = CConnect.GetRecordSet("select * from positionrequirementsvalue where positionrequirementsID=" & cboParameter.Tag & " and positionID=" & lvwDetails.SelectedItem.Tag & "")
If rsLoad.RecordCount > 0 Then
    If MsgBox("The specified parameter already had an entry made corresponding" & vbCrLf & "to it on the selected position. Selecting YES will update the records." & vbCrLf & "Do you wish to proceed with update?", vbYesNo + vbQuestion, "Confirm update") = vbNo Then Exit Sub
    Action = "UPDATED A POSITION QUALIFICATION; POSITION: " & lvwDetails.SelectedItem.ListSubItems(1).Text & "; REQUIREMENT: " & cboParameter.Text & "; VALUE: " & txtQualificationParameterVal.Text
    CConnect.ExecuteSql ("update positionrequirementsvalue set positionrequirementsvalue='" & Replace(txtQualificationParameterVal.Text, "'", "''") & "' where positionrequirementsID=" & cboParameter.Tag & " and positionID=" & lvwDetails.SelectedItem.Tag & "")
Else
    If MsgBox("Are you sure you want to save the records?", vbYesNo + vbQuestion, "Confirm saving!") = vbNo Then Exit Sub
    Action = "ADDED A POSITION QUALIFICATION; POSITION: " & lvwDetails.SelectedItem.ListSubItems(1).Text & "; REQUIREMENT: " & cboParameter.Text & "; VALUE: " & txtQualificationParameterVal.Text
    CConnect.ExecuteSql "INSERT INTO positionrequirementsvalue (positionrequirementsID,positionID,positionrequirementsvalue) VALUES(" & cboParameter.Tag & "," & lvwDetails.SelectedItem.Tag & ",'" & Replace(txtQualificationParameterVal.Text, "'", "''") & "')"
End If
With cboParameter
    .Text = ""
    .Tag = ""
    .Locked = False
End With

txtQualificationParameterVal.Text = ""
Call DisplayRecordsQualifications
End Sub

Private Sub Form_Load()
Decla.Security Me
    oSmart.FReset Me
    
    If oSmart.hRatio > 1.1 Then
        With frmMain2
            Me.Move .tvwMain.Width + .tvwMain.Width / 32.75, (.Height / 5.52) ' - 155
            .lvwEmp.Visible = False
        End With
    Else
         With frmMain2
            Me.Move .tvwMain.Width + .tvwMain.Width / 32.75, .Height / 5.55
        End With
        
    End If
    
    CConnect.CColor Me, MyColor
    LoadCbo
   
    Call InitGrid
    
    Set rs2 = CConnect.GetRecordSet("SELECT * FROM Positions ORDER BY Code")
    
    Call DisplayRecords
    
End Sub

Private Sub Form_Resize()
    oSmart.FResize Me
End Sub

Private Sub InitGrid()
    With lvwDetails
        .ColumnHeaders.Add , , "Code", 0
        .ColumnHeaders.Add , , "Position", 2 * .Width / 7
        .ColumnHeaders.Add , , "Approved", 2 * .Width / 7, vbCenter
        .ColumnHeaders.Add , , "Comments", 2 * .Width / 7
        .ColumnHeaders.Add , , "Daily Casual Rate", .Width / 7
                   
        .View = lvwReport
    End With
    
    With lvwQualification
        .ColumnHeaders.Add , , "Position Parameter", .Width / 4
        .ColumnHeaders.Add , , "Qualification", 3 * .Width / 4
        
        .View = lvwReport
    End With
End Sub
Public Sub DisplayRecordsQualifications()
On Error GoTo Hell
    lvwQualification.ListItems.Clear
    With rsShowRecords
        .Requery
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Set LI = lvwQualification.ListItems.Add(, , !positionrequirementsdescription & "", , 5)
                LI.ListSubItems.Add , , !positionrequirementsvalue & ""
                LI.Tag = Trim(!positionrequirementsid & "")
                .MoveNext
            Loop
        End If
    End With
 Exit Sub
Hell:
End Sub
Public Sub DisplayRecords()
On Error GoTo Hell
    lvwDetails.ListItems.Clear
    With rs2
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Set LI = lvwDetails.ListItems.Add(, , !code & "", , 5)
                LI.ListSubItems.Add , , !positionName & ""
                LI.ListSubItems.Add , , !approved & ""
                LI.ListSubItems.Add , , !Comments & ""
                LI.ListSubItems.Add , , !dailyrate & ""
                LI.Tag = Trim(!id & "")
                .MoveNext
            Loop
        End If
    End With
 Exit Sub
Hell:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs2 = Nothing
    Set rs5 = Nothing
'    Set Cnn = Nothing
    
    frmMain2.lvwEmp.Visible = True
    frmMain2.Caption = "Personnel Director " & App.FileDescription
End Sub

Private Sub lvwDetails_ItemClick(ByVal Item As MSComctlLib.ListItem)
Set rsShowRecords = CConnect.GetRecordSet("select pr.positionrequirementsdescription,prv.positionrequirementsvalue,prv.positionrequirementsID from positionrequirements pr inner join positionrequirementsvalue prv on pr.positionrequirementsID=prv.positionrequirementsID where prv.positionid=" & lvwDetails.SelectedItem.Tag & "")
Call DisplayRecordsQualifications
End Sub
Private Sub lvwQualification_DblClick()
Call cmdEdit_Click
With frmMain2
    .cmdSave.Enabled = True
    .cmdNew.Enabled = False
    .cmdEdit.Enabled = False
    .cmdCancel.Enabled = True
    .cmdDelete.Enabled = False
End With
End Sub

Private Sub lvwQualification_ItemClick(ByVal Item As MSComctlLib.ListItem)
With frmMain2
    .cmdDelete.Enabled = True
End With
End Sub
