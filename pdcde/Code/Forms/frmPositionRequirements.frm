VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmPositionRequirements 
   BorderStyle     =   0  'None
   Caption         =   "Position Requirements"
   ClientHeight    =   7155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   7020
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9930
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
         TabIndex        =   16
         ToolTipText     =   "Move to the Last employee"
         Top             =   5400
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
         Picture         =   "frmPositionRequirements.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Delete Record"
         Top             =   5400
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
         Picture         =   "frmPositionRequirements.frx":04F2
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Edit Record"
         Top             =   5400
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
         Picture         =   "frmPositionRequirements.frx":05F4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Add New record"
         Top             =   5400
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
         TabIndex        =   12
         ToolTipText     =   "Move to the Next employee"
         Top             =   5400
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
         TabIndex        =   11
         ToolTipText     =   "Move to the Previous employee"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "CLOSE"
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
         Left            =   6270
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   5400
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
         TabIndex        =   9
         ToolTipText     =   "Move to the First employee"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame fraDetails 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Position Requirements Set-up"
         ForeColor       =   &H80000008&
         Height          =   3330
         Left            =   2025
         TabIndex        =   1
         Top             =   1695
         Visible         =   0   'False
         Width           =   6015
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
            Left            =   5370
            Picture         =   "frmPositionRequirements.frx":06F6
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Cancel Process"
            Top             =   2700
            Width           =   495
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
            Left            =   4875
            Picture         =   "frmPositionRequirements.frx":07F8
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Save Record"
            Top             =   2700
            Width           =   510
         End
         Begin VB.TextBox txtComments 
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
            Height          =   1305
            Left            =   135
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   1320
            Width           =   5715
         End
         Begin VB.TextBox txtPositionParameter 
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
            Left            =   120
            TabIndex        =   2
            Top             =   720
            Width           =   5535
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Comments"
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
            TabIndex        =   7
            Top             =   1080
            Width           =   750
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Requirement parameter"
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
            TabIndex        =   6
            Top             =   480
            Width           =   1710
         End
      End
      Begin MSComctlLib.ListView lvwDetails 
         Height          =   6330
         Left            =   50
         TabIndex        =   8
         Top             =   600
         Width           =   9810
         _ExtentX        =   17304
         _ExtentY        =   11165
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
      Begin MSComDlg.CommonDialog Cdl 
         Left            =   5610
         Top             =   4560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Positions Requirements"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   3405
      End
   End
End
Attribute VB_Name = "frmPositionRequirements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub cmdCancel_Click()
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
If MsgBox("Are you sure you want to delete " & lvwDetails.SelectedItem.ListSubItems(1).Text & "?", vbYesNo + vbQuestion, "Confirm delete") = vbNo Then Exit Sub
Action = "DELETED A POSITION REQUIREMENT; REQUIREMENT DESCRIPTION: " & lvwDetails.SelectedItem.ListSubItems(1).Text
CConnect.ExecuteSql "DELETE FROM positionrequirements WHERE positionrequirementsID=" & lvwDetails.SelectedItem.Tag
CConnect.ExecuteSql "DELETE FROM positionrequirementsvalue WHERE positionrequirementsID=" & lvwDetails.SelectedItem.Tag
Call DisplayRecords
End Sub

Public Sub cmdEdit_Click()
If lvwDetails.ListItems.Count = 0 Then MsgBox "There are no records to edit.", vbOKOnly + vbInformation, "No records": Exit Sub
With txtPositionParameter
    .Text = lvwDetails.SelectedItem.ListSubItems(1).Text
    .Tag = lvwDetails.SelectedItem.ListSubItems(1).Text
End With

txtComments.Text = lvwDetails.SelectedItem.ListSubItems(2).Text
txtPositionParameter.Locked = False
fraDetails.Visible = True
End Sub

Public Sub cmdNew_Click()
fraDetails.Visible = True
With txtPositionParameter
    .Locked = False
    .Tag = ""
    .Text = ""
End With
End Sub

Public Sub cmdSave_Click()
Dim rs As New ADODB.Recordset

If txtPositionParameter.Text = "" Then MsgBox "Please specify a parameter to proceed.", vbOKOnly + vbInformation, "Missing parameter": Exit Sub

Set rs = CConnect.GetRecordSet("select * from positionrequirements where positionrequirementsdescription like '" & txtPositionParameter.Tag & "'")
If rs.RecordCount > 0 Then
    If MsgBox("The specified parameter had already been registered." & vbCrLf & "Do you wish to modify it?", vbYesNo + vbQuestion, "Requirements parameters") = vbNo Then Exit Sub
    Action = "UPDATED POSITIONS REQUIREMENTS; REQUIREMENT NAME: " & txtPositionParameter & "; COMMENTS: " & txtComments.Text
    CConnect.ExecuteSql "UPDATE positionrequirements SET positionrequirementsdescription='" & Replace(txtPositionParameter.Text, "'", "''") & "',positioNrequirementscomments='" & Replace(txtComments.Text, "'", "''") & "' where positionrequirementsdescription='" & Replace(txtPositionParameter.Tag, "'", "''") & "'"
Else
    If MsgBox("Are you sure you want to save the record.", vbYesNo + vbQuestion, "Confirm save") = vbNo Then Exit Sub
    
    Action = "ADDED POSITIONS REQUIREMENTS; REQUIREMENT NAME: " & txtPositionParameter & "; COMMENTS: " & txtComments.Text
    CConnect.ExecuteSql "INSERT INTO positionrequirements(positionrequirementsdescription,positionrequirementscomments) VALUES('" & Replace(txtPositionParameter.Text, "'", "''") & "','" & Replace(txtComments.Text, "'", "''") & "')"
End If

With txtPositionParameter
    .Text = ""
    .Tag = ""
    .SetFocus
End With

txtComments.Text = ""
Call DisplayRecords
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
    
    Call InitGrid
    
    Set rs2 = CConnect.GetRecordSet("SELECT * FROM PositionRequirements")
    
    Call DisplayRecords
    
End Sub

Private Sub Form_Resize()
    oSmart.FResize Me
End Sub

Private Sub InitGrid()
With lvwDetails
    .ColumnHeaders.Add , , "S/No.", .Width / 6
    .ColumnHeaders.Add , , "Requirements Description", .Width / 3
    .ColumnHeaders.Add , , "Comments", .Width / 2
    .View = lvwReport
End With
End Sub
Public Sub DisplayRecords()
    lvwDetails.ListItems.Clear
    i = 1
    With rs2
        .Requery
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Set LI = lvwDetails.ListItems.Add(, , i & ".", , 5)
                LI.ListSubItems.Add , , !positionrequirementsdescription & ""
                LI.ListSubItems.Add , , !positionrequirementsComments & ""
                LI.Tag = Trim(!positionrequirementsid & "")
                i = i + 1
                .MoveNext
            Loop
        End If
    End With
End Sub

Private Sub lvwDetails_DblClick()
Call cmdEdit_Click
End Sub

Private Sub lvwDetails_ItemClick(ByVal Item As MSComctlLib.ListItem)
frmMain2.cmdDelete.Enabled = True
End Sub
