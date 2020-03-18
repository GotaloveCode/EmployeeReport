VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEducationCourses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11160
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      Caption         =   "Education Courses Setup"
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   2280
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox txtCourseName 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1455
         TabIndex        =   15
         Top             =   600
         Width           =   4455
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   135
         TabIndex        =   12
         Top             =   600
         Width           =   1110
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
         Picture         =   "frmEducationCourses.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Save Record"
         Top             =   1095
         Width           =   510
      End
      Begin VB.CommandButton cmdCancel 
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
         Picture         =   "frmEducationCourses.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Cancel Process"
         Top             =   1095
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Course Name"
         Height          =   195
         Left            =   1455
         TabIndex        =   14
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Course Code"
         Height          =   195
         Left            =   135
         TabIndex        =   13
         Top             =   360
         Width           =   930
      End
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
      Left            =   1800
      TabIndex        =   8
      ToolTipText     =   "Move to the First employee"
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "CLOSE"
      Height          =   495
      Left            =   7950
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6600
      Visible         =   0   'False
      Width           =   1050
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
      Left            =   2280
      TabIndex        =   6
      ToolTipText     =   "Move to the Previous employee"
      Top             =   6600
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
      Left            =   2760
      TabIndex        =   5
      ToolTipText     =   "Move to the Next employee"
      Top             =   6600
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
      Left            =   4455
      Picture         =   "frmEducationCourses.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Add New record"
      Top             =   6600
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
      Left            =   4935
      Picture         =   "frmEducationCourses.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Edit Record"
      Top             =   6600
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
      Left            =   5400
      Picture         =   "frmEducationCourses.frx":0408
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Delete Record"
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
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
      Left            =   3240
      TabIndex        =   1
      ToolTipText     =   "Move to the Last employee"
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComctlLib.ListView lvwDetails 
      Height          =   7800
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   11130
      _ExtentX        =   19632
      _ExtentY        =   13758
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
      Left            =   7290
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmEducationCourses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--Modification to allow for setting up of education courses and their respective awards
Option Explicit

Private MyEdCourse As HRCORE.EducationCourse
Private MyEducationCourses As HRCORE.EducationCourses

Private Sub cmdCancel_Click()
    On Error GoTo ErrHandler
    
    frmMain2.RestoreCommandButtonState
    Me.fraDetails.Visible = False
    Me.cmdCancel.Enabled = False
    Me.cmdEdit.Enabled = True
    Me.cmdNew.Enabled = True
    Me.cmdSave.Enabled = False
    
    Exit Sub
ErrHandler:
    MsgBox "An error has occured: " & vbNewLine & err.Description, vbExclamation, "PDR Error"
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Public Sub cmdDelete_Click()
    On Error Resume Next
    Dim resp As String
    
    Dim rs As ADODB.Recordset
    
    If lvwDetails.ListItems.count > 0 Then
        'Deleting a record
        '---------added by kalya on 16th march,2009
        resp = MsgBox("Are you sure you want to delete course  " & lvwDetails.SelectedItem.SubItems(1) & vbNewLine & "from employee education history records?", vbQuestion + vbYesNo)
        If resp = vbNo Then
            Exit Sub
        End If
        
        Action = "DELETED A COURSE; COURSE CODE: " & lvwDetails.SelectedItem.Text & "; COURSE: " & lvwDetails.SelectedItem.SubItems(1)
        CConnect.ExecuteSql ("DELETE FROM pdreducationcourses WHERE coursecode = '" & lvwDetails.SelectedItem.Text & "' AND Coursename = '" & lvwDetails.SelectedItem.SubItems(1) & "'")
        rs2.Requery
        Call DisplayRecords
        
        '--------end deleting
    Else
        MsgBox "You have to select the course you would like to delete.", vbInformation
    End If
End Sub

Private Sub cmdDone_Click()
    PSave = True
    Call cmdCancel_Click
    PSave = False
End Sub

Public Sub cmdEdit_Click()
    On Error GoTo ErrHandler
    
    If lvwDetails.ListItems.count < 1 Then
        MsgBox "You have to select the course details you would like to edit.", vbInformation
        PSave = True
        'Call cmdCancel_Click
        PSave = False
        Exit Sub
    End If
                
    Call DisableCmd
    
    Set MyEdCourse = New EducationCourse
    Set MyEdCourse = MyEducationCourses.GetByEducationCourseID(lvwDetails.SelectedItem.Tag)
    With MyEdCourse
        txtCode.Text = .CourseCode
        txtCourseName.Text = .CourseName
    End With
    
    fraDetails.Visible = True
    
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    SaveNew = False
    txtCode.Locked = False
    txtCourseName.Locked = False
    
    Exit Sub
    
ErrHandler:
    MsgBox err.Description, vbInformation
End Sub

'Private Sub cmdFirst_Click()
'
'With rsGlob
'    If .RecordCount > 0 Then
'        If .BOF <> True Then
'            .MoveFirst
'            If .BOF = True Then
'                .MoveFirst
'                Call DisplayRecords
'            Else
'                Call DisplayRecords
'            End If
'
'            Call FirstDisb
'
'        End If
'    End If
'End With
'
'
'End Sub
'
'Private Sub cmdLast_Click()
'With rsGlob
'    If .RecordCount > 0 Then
'        If .EOF <> True Then
'            .MoveLast
'            If .EOF = True Then
'                .MoveLast
'                Call DisplayRecords
'            Else
'                Call DisplayRecords
'            End If
'
'            Call LastDisb
'
'        End If
'    End If
'End With
'
'End Sub

Public Sub cmdNew_Click()
    Call DisableCmd
        
    txtCode.Text = ""
    txtCourseName.Text = ""
    
    fraDetails.Visible = True
    cmdCancel.Enabled = True
    SaveNew = True
    cmdSave.Enabled = True
    txtCode.Locked = False
    txtCode.SetFocus
End Sub

Public Sub cmdSave_Click()
    On Error GoTo ErrHandler
    
    If txtCode.Text = "" Then
        MsgBox "Enter the education course code.", vbExclamation
        txtCode.SetFocus
        Exit Sub
    End If
    
    If txtCourseName.Text = "" Then
        MsgBox "Enter the course name", vbExclamation, "Save Error"
        Exit Sub
    End If
    
    'Assign Values to Education Course
    Set MyEdCourse = New HRCORE.EducationCourse
    With MyEdCourse
        .CourseCode = Trim(Replace(txtCode.Text, "'", "''"))
        .CourseName = Trim(Replace(txtCourseName.Text, "'", "''"))
    End With
    
    If SaveNew = True Then
        'Saved Record
        MyEdCourse.ModifyEducationCourse (0)    'SAVE
        'Update the AuditTrail
        Action = "Added an education course: " & txtCode & " - " & txtCourseName
        currUser.AuditTrail Add_New, Action
    Else
        'Update
        MyEdCourse.EducationCourseID = lvwDetails.SelectedItem.Tag
        MyEdCourse.ModifyEducationCourse (1)
        'Update the AuditTrail
        Action = "Modified an education course: " & txtCode & " - " & txtCourseName & " TO: - " & txtCode & " - " & txtCourseName
        currUser.AuditTrail Add_New, Action
    End If
      ''
       frmMain2.cmdNew.Enabled = True
       frmMain2.cmdSave.Enabled = False
       frmMain2.cmdCancel = False
      ''
    If SaveNew = False Then
        DisplayRecords
    Else
        Call DisplayRecords
    End If
    
    'Unset Controls
    txtCode.Text = ""
    txtCourseName.Text = ""
    Set MyEdCourse = Nothing
    fraDetails.Visible = False
    
    Exit Sub
ErrHandler:
    MsgBox "An error has occured while updating the education courses" & vbNewLine & err.Description, vbInformation, "PDR Error"
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler
    oSmart.FReset Me
    
    If oSmart.hRatio > 1.1 Then
        With frmMain2
            Me.Move .tvwMain.Width + .tvwMain.Width / 32.75, (.Height / 5.52) ' - 155
        End With
    Else
         With frmMain2
            Me.Move .tvwMain.Width + .tvwMain.Width / 32.75, .Height / 5.55
        End With
        
    End If
    
    CConnect.CColor Me, MyColor
    
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    
    Call InitGrid
    
    DisplayRecords
    
    cmdFirst.Enabled = False
    cmdPrevious.Enabled = False
    
    Exit Sub
    
ErrHandler:
    MsgBox err.Description, vbInformation
End Sub

Private Sub Form_Resize()
    oSmart.FResize Me
    
    Me.Height = tvwMainheight - 150
    'Frame1.Move Frame1.Left, 0, Frame1.Width, tvwMainheight - 150
    lvwDetails.Move lvwDetails.Left, 0, lvwDetails.Width, tvwMainheight - 150
    
End Sub

Private Sub InitGrid()
    With lvwDetails
        .ColumnHeaders.add , , "Course Code", .Width / 7
        .ColumnHeaders.add , , "Course Name", 3 * .Width / 7
        .View = lvwReport
    End With
End Sub

Public Sub DisplayRecords()
    On Error GoTo ErrHandler
    
    lvwDetails.ListItems.Clear
    'Call Cleartxt
    
    Set MyEdCourse = New EducationCourse
    
    'Load Education Courses
    Set MyEducationCourses = New EducationCourses
    MyEducationCourses.GetAllEducationCourses
    
    Dim n As Integer
    
    For n = 1 To MyEducationCourses.count
        Set li = lvwDetails.ListItems.add(, , MyEducationCourses.Item(n).CourseCode, , 5)
        li.ListSubItems.add , , MyEducationCourses.Item(n).CourseName
        li.Tag = Trim(MyEducationCourses.Item(n).EducationCourseID)
    Next n
    
    Exit Sub
    
ErrHandler:
    MsgBox err.Description, vbExclamation, "Error"
End Sub

Private Sub lvwDetails_DblClick()
    Me.cmdEdit_Click
End Sub
