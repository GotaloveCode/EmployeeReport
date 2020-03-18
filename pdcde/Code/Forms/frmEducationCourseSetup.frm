VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmEducationCourseSetup 
   BorderStyle     =   0  'None
   Caption         =   "Education Courses Setup"
   ClientHeight    =   7920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Education Courses Setup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   2280
      TabIndex        =   9
      Top             =   1560
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox txtCode 
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
         Picture         =   "frmEducationCourseSetup.frx":0000
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
         Picture         =   "frmEducationCourseSetup.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Cancel Process"
         Top             =   1095
         Width           =   495
      End
      Begin VB.TextBox txtCourseName 
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
         Left            =   1455
         TabIndex        =   14
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "Course Name"
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
         Left            =   1455
         TabIndex        =   15
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "Course Code"
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
      Left            =   1680
      TabIndex        =   8
      ToolTipText     =   "Move to the First employee"
      Top             =   5880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7830
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5880
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
      Left            =   2160
      TabIndex        =   6
      ToolTipText     =   "Move to the Previous employee"
      Top             =   5880
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
      Left            =   2640
      TabIndex        =   5
      ToolTipText     =   "Move to the Next employee"
      Top             =   5880
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
      Left            =   4335
      Picture         =   "frmEducationCourseSetup.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Add New record"
      Top             =   5880
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
      Left            =   4815
      Picture         =   "frmEducationCourseSetup.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Edit Record"
      Top             =   5880
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
      Left            =   5280
      Picture         =   "frmEducationCourseSetup.frx":0408
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Delete Record"
      Top             =   5880
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
      Left            =   3120
      TabIndex        =   1
      ToolTipText     =   "Move to the Last employee"
      Top             =   5880
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComctlLib.ListView lvwDetails 
      Height          =   7800
      Left            =   0
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
      Left            =   7170
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmEducationCourseSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MyEdCourse As EducationCourse
Dim MyEducationCourses As EducationCourses
Dim n As Integer

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub cmdDelete_Click()
    On Error Resume Next
    Dim resp As String
    
    Dim rs As ADODB.Recordset
    
    If lvwDetails.ListItems.Count > 0 Then
        resp = MsgBox("This will delete  " & lvwDetails.SelectedItem & " and the corresponding employee BankDetails from the records . Do you wish to continue?", vbQuestion + vbYesNo)
        If resp = vbNo Then
            Exit Sub
        End If
        'get all the branches of this bank
        mySQL = "select distinct bankbranch_id as branchID from tblbankbranch  where bank_id = '" & lvwDetails.SelectedItem.Tag & "'"
        
        Set rs = CConnect.GetRecordSet(mySQL)
        
        Action = "DELETED BANK; BANK CODE: " & lvwDetails.SelectedItem & "; BANK NAME: " & lvwDetails.SelectedItem.ListSubItems(1) & "; COMMENTS: " & lvwDetails.SelectedItem.ListSubItems(2)
        
        CConnect.ExecuteSql ("DELETE FROM tblBank WHERE bank_id = '" & lvwDetails.SelectedItem.Tag & "'")
        
        Action = ""
        If Not rs Is Nothing Then
            If rs.RecordCount > 0 Then
                CConnect.ExecuteSql ("DELETE FROM tblBankBranch WHERE bank_id = '" & lvwDetails.SelectedItem.Tag & "'")
                rs.MoveFirst
                
                Do Until rs.EOF
                    CConnect.ExecuteSql ("DELETE FROM employeebanks WHERE branchID = " & rs!BranchID)
                    
                    CConnect.ExecuteSql ("DELETE FROM tblCompanybank WHERE Bankbranch_ID =" & rs!BranchID)
                    rs.MoveNext
                Loop
            End If
        End If
        rs2.Requery
        
        Call DisplayRecords
        
        Set rs = Nothing
    Else
        MsgBox "You have to select the BankDetails  you would like to delete.", vbInformation
    End If
End Sub

Private Sub cmdDone_Click()
    PSave = True
    
    PSave = False
End Sub

Public Sub cmdEdit_Click()

    On Error GoTo ErrHandler
    If lvwDetails.ListItems.Count < 1 Then
        MsgBox "You have to select the course details you would like to edit.", vbInformation
        PSave = True
        
        PSave = False
        Exit Sub
    End If
    
    If (MyEducationCourses.GetByEducationCourseID(lvwDetails.SelectedItem.Tag) Is Nothing) Then
        MsgBox "Error loading recorsd", vbExclamation, "Edit Error"
        Exit Sub
    Else
        'Edit Record
        Set MyEdCourse = MyEducationCourses.GetByEducationCourseID(lvwDetails.SelectedItem.Tag)
        txtCode.Text = MyEdCourse.CourseCode
        txtCourseName.Text = MyEdCourse.CourseName
        SaveNew = False
    End If
       
    
    Call DisableCmd
    
    fraDetails.Visible = True
    
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    SaveNew = False
    txtCode.Locked = False
    
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbInformation
End Sub

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
    
    If Not (MyEducationCourses.GetByEducationCourseName(txtCourseName) Is Nothing) Then
        MsgBox "The Code Selected has already been used for: " & MyEducationCourses.GetByEducationCourseName(txtCourseName).CourseName, vbExclamation, "Duplicate Record Avoided"
        Exit Sub
    End If
            
    'thus far means validation is good
    With MyEdCourse
        .CourseCode = Trim(txtCode)
        .CourseName = Trim(txtCourseName)
    End With
    
    If SaveNew = True Then
        'Saved Record
        MyEdCourse.ModifyEducationCourse (0) '0 Is Add
        'audit trail
        Action = "Added an education course: " & txtCode & " - " & txtCourseName & " Award Name: " & txtCourseName
    End If
        
    If SaveNew = False Then
        MyEdCourse.ModifyEducationCourse (1)
    End If
        
    If SaveNew = False Then
        PSave = True
        
        PSave = False
    Else
        Call DisplayRecords
        txtCode.SetFocus
    End If
    
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbInformation
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
    MsgBox Err.Description, vbInformation
End Sub

Private Sub Form_Resize()
    oSmart.FResize Me
    
    Me.Height = tvwMainheight - 150
    'Frame1.Move Frame1.Left, 0, Frame1.Width, tvwMainheight - 150
    lvwDetails.Move lvwDetails.Left, 0, lvwDetails.Width, tvwMainheight - 150
    
End Sub

Private Sub InitGrid()
    With lvwDetails
        .ColumnHeaders.Add , , "Course Code", .Width / 7
        .ColumnHeaders.Add , , "Course Name", 3 * .Width / 7
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
    
    For n = 1 To MyEducationCourses.Count
        Set LI = lvwDetails.ListItems.Add(, , MyEducationCourses.Item(n).CourseCode, , 5)
        LI.ListSubItems.Add , , MyEducationCourses.Item(n).CourseName
        LI.Tag = Trim(MyEducationCourses.Item(n).EducationCourseID)
    Next n
    
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub


