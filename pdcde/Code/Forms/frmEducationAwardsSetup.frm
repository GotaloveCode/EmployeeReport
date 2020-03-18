VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEducationAwardsSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
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
      Left            =   2640
      TabIndex        =   17
      ToolTipText     =   "Move to the Last employee"
      Top             =   6240
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
      Left            =   4800
      Picture         =   "frmEducationAwardsSetup.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Delete Record"
      Top             =   6240
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
      Left            =   4335
      Picture         =   "frmEducationAwardsSetup.frx":04F2
      Style           =   1  'Graphical
      TabIndex        =   15
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
      Left            =   3855
      Picture         =   "frmEducationAwardsSetup.frx":05F4
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Add New record"
      Top             =   6240
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
      Left            =   2160
      TabIndex        =   13
      ToolTipText     =   "Move to the Next employee"
      Top             =   6240
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
      Left            =   1680
      TabIndex        =   12
      ToolTipText     =   "Move to the Previous employee"
      Top             =   6240
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
      Left            =   7350
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6240
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
      Left            =   1200
      TabIndex        =   10
      ToolTipText     =   "Move to the First employee"
      Top             =   6240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      Caption         =   "Education Awards Setup"
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
      Height          =   2655
      Left            =   1680
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox txtAwardCode 
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
         TabIndex        =   2
         Top             =   1440
         Width           =   1110
      End
      Begin VB.ComboBox cboCourseName 
         Height          =   315
         Left            =   1455
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   4455
      End
      Begin VB.TextBox txtAwardName 
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
         TabIndex        =   3
         Top             =   1440
         Width           =   4455
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
         Picture         =   "frmEducationAwardsSetup.frx":06F6
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Cancel Process"
         Top             =   1935
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
         Picture         =   "frmEducationAwardsSetup.frx":07F8
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Save Record"
         Top             =   1935
         Width           =   510
      End
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
         TabIndex        =   0
         Top             =   600
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Award Code"
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
         TabIndex        =   19
         Top             =   1200
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Award Name"
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
         Left            =   1440
         TabIndex        =   18
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   9
         Top             =   360
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   8
         Top             =   360
         Width           =   960
      End
   End
   Begin MSComDlg.CommonDialog Cdl 
      Left            =   6690
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvwDetails 
      Height          =   7800
      Left            =   0
      TabIndex        =   4
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
End
Attribute VB_Name = "frmEducationAwardsSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MyEdCourse As EducationCourse
Dim MyEducationCourses As EducationCourses
Dim myEdAward As EducationCourseAward
Dim MyEducationAwards As EducationCourseAwards


Private Sub cboCourseName_Click()
    On Error Resume Next
    
    Set MyEdCourse = New EducationCourse
    Set MyEdCourse = MyEducationCourses.GetByEducationCourseID(cboCourseName.ItemData(cboCourseName.ListIndex))
    txtCode.Text = MyEdCourse.CourseCode
    
End Sub

Private Sub cmdCancel_Click()
    fraDetails.Visible = False
    EnableCmd
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Public Sub cmdDelete_Click()

''-------------

    Dim resp As String
    On Error GoTo ErrHandler
    If Not currUser Is Nothing Then
        If currUser.CheckRight("SetUpAwardTypes") <> secModify Then
            MsgBox "You dont have right to delete the record. Please liaise with the security admin"
            Exit Sub
        End If
        
        Dim rt As String
        rt = currUser.CheckRight("SetUpAwardTypes")
    End If
    
    
     
    If lvwDetails.ListItems.count > 0 Then
        resp = MsgBox("Are you sure you want to delete  " & lvwDetails.SelectedItem & " from the records?", vbQuestion + vbYesNo)
        If resp = vbNo Then
            Exit Sub
        End If
        Action = "DELETED AWARD TYPE. AWARD CODE: " & lvwDetails.SelectedItem
        CConnect.ExecuteSql ("UPDATE pdrEducationAwards set deleted=1 WHERE educationawardid = '" & lvwDetails.SelectedItem.Tag & "'")
       
        Call DisplayRecords
    Else
        MsgBox "You have to select the award type you would like to delete.", vbInformation
                
    End If
    Exit Sub
ErrHandler:
MsgBox (err.Description)

''--------------




'    On Error Resume Next
'    Dim resp As String
'
'    ''place deleting code here
    
    Call DisplayRecords
    
End Sub

Private Sub cmdDone_Click()
'    PSave = True
'    Call cmdCancel_Click
'    PSave = False
End Sub

Public Sub cmdEdit_Click()
    On Error GoTo ErrHandler
    
    If lvwDetails.ListItems.count < 1 Then
        MsgBox "You have to select the course details you would like to edit.", vbInformation
        PSave = True
        PSave = False
        Exit Sub
    End If
    
    Set myEdAward = New EducationCourseAward
    
    Set myEdAward = MyEducationAwards.GetByEducationAwardID(lvwDetails.SelectedItem.Tag)
    
    Call DisableCmd
    
    fraDetails.Visible = True
    txtAwardCode.Text = myEdAward.AwardCode
    txtAwardName.Text = myEdAward.AwardName
    cboCourseName.Text = "[" & myEdAward.EducationCourseOBJ.CourseCode & "]  " & myEdAward.EducationCourseOBJ.CourseName
        
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    SaveNew = False
    txtCode.Locked = False
    
    Exit Sub
    
ErrHandler:
    MsgBox "An error has occured:" & vbNewLine & err.Description, vbInformation
End Sub

Public Sub cmdNew_Click()


  If Not currUser Is Nothing Then
        If currUser.CheckRight("SetUpAwardTypes") <> secModify Then
            MsgBox "You dont have right to delete the record. Please liaise with the security admin"
            Exit Sub
        End If
        
        Dim rt As String
        rt = currUser.CheckRight("SetUpAwardTypes")
    End If
    Call DisableCmd
        
    txtCode.Text = ""
    txtAwardName.Text = ""
    fraDetails.Visible = True
    cmdCancel.Enabled = True
    SaveNew = True
    cmdSave.Enabled = True
    txtCode.Locked = False
    txtCode.SetFocus
End Sub

Public Sub cmdSave_Click()
    On Error GoTo ErrHandler
    
    If txtAwardCode.Text = "" Then
        MsgBox "Enter the education award code.", vbExclamation
        txtAwardCode.SetFocus
        Exit Sub
    End If
    
    If txtAwardName.Text = "" Then
        MsgBox "Enter the education award name.", vbExclamation
        txtAwardName.SetFocus
        Exit Sub
    End If
    
    Set myEdAward = New EducationCourseAward
           
    With myEdAward
        .AwardCode = Trim(txtAwardCode.Text)
        .AwardName = Trim(txtAwardName.Text)
        .EducationCourseID = cboCourseName.ItemData(cboCourseName.ListIndex)
    End With
    
    If SaveNew = True Then
        'Saved Record
        If (myEdAward.ModifyEducationAward(0) = True) Then
            'Saved OK, Update Audit Trail
            Action = "Added an education award: " & txtCode & " - " & " Award Name: " & txtAwardName
            currUser.AuditTrail Add_New, Action
        Else
            Exit Sub
        End If
    Else
        'Modify Record
        myEdAward.EducationAwardID = lvwDetails.SelectedItem.Tag
        If (myEdAward.ModifyEducationAward(1)) Then
            'Updated OK, Update Audit Trail
            Action = "Updated an education award: edited to " & UCase(txtAwardCode) & " - " & UCase(txtAwardName)
            currUser.AuditTrail Update, Action
        Else
            Exit Sub
        End If
    End If
      
    If SaveNew = False Then
        'EDIT MODE
        DisplayRecords
        cmdCancel_Click
        EnableCmd
    Else
        Call DisplayRecords
        cmdCancel_Click
        EnableCmd
    End If
    
    Exit Sub
ErrHandler:
    MsgBox "An error has occured: " & vbNewLine & err.Description, vbInformation
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
    
    
      'Load Education Courses
    Set MyEducationCourses = New EducationCourses
    MyEducationCourses.GetAllEducationCourses
Dim n As Integer
    cboCourseName.Clear
    For n = 1 To MyEducationCourses.count
        cboCourseName.AddItem "[" & MyEducationCourses.Item(n).CourseCode & "]  " & MyEducationCourses.Item(n).CourseName
        cboCourseName.ItemData(cboCourseName.NewIndex) = MyEducationCourses.Item(n).EducationCourseID
    Next n
    '------------------
    Call InitGrid
    
    DisplayRecords
    
    cmdFirst.Enabled = False
    cmdPrevious.Enabled = False
    
    Exit Sub
    
ErrHandler:
    MsgBox "An error has occured while loading awards" & vbNewLine & err.Description, vbInformation
End Sub

Private Sub Form_Resize()
    oSmart.FResize Me
    
    Me.Height = tvwMainheight - 150
    'Frame1.Move Frame1.Left, 0, Frame1.Width, tvwMainheight - 150
    lvwDetails.Move lvwDetails.Left, 0, lvwDetails.Width, tvwMainheight - 150
    
End Sub

Private Sub InitGrid()
    With lvwDetails
        .ColumnHeaders.add , , "Award Code", .Width / 7
        .ColumnHeaders.add , , "Award Name", 3 * .Width / 7
        .ColumnHeaders.add , , "Course", 3 * .Width / 7
        .View = lvwReport
    End With
End Sub

Public Sub DisplayRecords()
    On Error GoTo ErrHandler
    Dim n As Integer
    
    lvwDetails.ListItems.Clear
    
    'Initialize Objects

    Set MyEducationAwards = New EducationCourseAwards
    
  
    
    'Load The Awards
    MyEducationAwards.GetAllEducationCourseAwards
    
    For n = 1 To MyEducationAwards.count
        Set li = lvwDetails.ListItems.add(, , MyEducationAwards.Item(n).AwardCode, , 5)
        li.ListSubItems.add , , MyEducationAwards.Item(n).AwardName
        li.ListSubItems.add , , "[" & MyEducationAwards.Item(n).EducationCourseOBJ.CourseCode & "] " & MyEducationAwards.Item(n).EducationCourseOBJ.CourseName
        li.Tag = Trim(MyEducationAwards.Item(n).EducationAwardID)
    Next n
    
    Exit Sub
    
ErrHandler:
    MsgBox "An error has occured while loading course awards" & vbNewLine & err.Description, vbExclamation, "Error"
End Sub

Private Sub lvwDetails_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmdEdit_Click
End Sub
