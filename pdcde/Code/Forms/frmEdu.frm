VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmEdu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Educational History"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   Icon            =   "frmEdu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imgEmpTool 
      Left            =   5430
      Top             =   195
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdu.frx":0442
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdu.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdu.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdu.frx":0778
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7860
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   7800
      Begin MSComDlg.CommonDialog Cdl 
         Left            =   5610
         Top             =   4560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
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
         Left            =   6270
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   5400
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
         Left            =   600
         TabIndex        =   10
         ToolTipText     =   "Move to the Previous employee"
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
         TabIndex        =   11
         ToolTipText     =   "Move to the Next employee"
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
         Left            =   2760
         Picture         =   "frmEdu.frx":0CBA
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Add New record"
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
         Picture         =   "frmEdu.frx":0DBC
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Edit Record"
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
         Picture         =   "frmEdu.frx":0EBE
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Delete Record"
         Top             =   5400
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
         Left            =   1560
         TabIndex        =   12
         ToolTipText     =   "Move to the Last employee"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame fraDetails 
         Appearance      =   0  'Flat
         Caption         =   "Education History"
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
         Height          =   4800
         Left            =   480
         TabIndex        =   18
         Top             =   795
         Visible         =   0   'False
         Width           =   6450
         Begin VB.ComboBox cboAward 
            Height          =   315
            Left            =   150
            TabIndex        =   28
            Text            =   "cboAward"
            Top             =   2400
            Width           =   6135
         End
         Begin VB.ComboBox cboCourse 
            Height          =   315
            Left            =   1560
            TabIndex        =   27
            Text            =   "cboCourse"
            Top             =   600
            Width           =   4695
         End
         Begin VB.TextBox txtInstitution 
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
            Height          =   300
            Left            =   150
            TabIndex        =   2
            Top             =   1185
            Width           =   6135
         End
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
            Left            =   5790
            Picture         =   "frmEdu.frx":13B0
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Cancel Process"
            Top             =   4095
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
            Left            =   5280
            Picture         =   "frmEdu.frx":14B2
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Save Record"
            Top             =   4095
            Width           =   495
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
            Height          =   975
            Left            =   150
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   2955
            Width           =   6135
         End
         Begin VB.TextBox txtCode 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Height          =   300
            Left            =   150
            TabIndex        =   1
            Top             =   600
            Width           =   1260
         End
         Begin VB.TextBox txtLevel 
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
            Left            =   3000
            TabIndex        =   5
            Top             =   1800
            Width           =   3285
         End
         Begin MSComCtl2.DTPicker dtpFrom 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   330
            Left            =   150
            TabIndex        =   3
            Top             =   1785
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   62980099
            CurrentDate     =   37673
            MinDate         =   21916
         End
         Begin MSComCtl2.DTPicker dtpTo 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   330
            Left            =   1425
            TabIndex        =   4
            Top             =   1792
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   62980099
            CurrentDate     =   37673
            MinDate         =   21916
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Institution"
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
            Left            =   150
            TabIndex        =   26
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Education"
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
            Left            =   1470
            TabIndex        =   25
            Top             =   375
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "To"
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
            TabIndex        =   24
            Top             =   1560
            Width           =   180
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
            Left            =   150
            TabIndex        =   23
            Top             =   2715
            Width           =   750
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Award"
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
            Left            =   150
            TabIndex        =   22
            Top             =   2160
            Width           =   465
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Code"
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
            Left            =   150
            TabIndex        =   21
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Level"
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
            Left            =   3000
            TabIndex        =   20
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "From"
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
            Left            =   150
            TabIndex        =   19
            Top             =   1560
            Width           =   360
         End
      End
      Begin MSComctlLib.ListView lvwDetails 
         Height          =   7800
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   7800
         _ExtentX        =   13758
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
End
Attribute VB_Name = "frmEdu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Myeducations As Educations
Dim MyEdAwards As EducationCourseAwards
Dim myEdAward As EducationCourseAward
Dim myInternalEduaction As Education
Dim myEducationHistories As EducationHistories
Dim myinternalEducationHistories As EducationHistories
Dim myinternaleducationhistory As EducationHistory
Dim MyEducationCourses As HRCORE.EducationCourses
Dim MyEducationCourse As HRCORE.EducationCourse
Dim n As Integer

Private Sub cboCourse_Change()
cboCourse_Click
End Sub

Private Sub cboCourse_Click()


cboAward.Clear
'For n = 1 To cboAward.ListCount
'cboAward.ListIndex = n
'If Not cboAward.Text = "fromedu" Then
'cboAward.RemoveItem (n)
'End If
'Next n


    For n = 1 To MyEdAwards.count
    If MyEdAwards.Item(n).EducationCourseOBJ.CourseName = cboCourse.Text Then
 
        cboAward.AddItem MyEdAwards.Item(n).AwardName
         
    End If
       Dim ss As String
    
    Next n
End Sub

Public Sub cmdCancel_Click()
'    If PSave = False Then
'        If MsgBox("Do you want to save changes?", vbQuestion + vbYesNo) = vbYes Then  '
'            Call cmdSave_Click
'            Exit Sub
'        End If
'    End If
'
'    Call DisplayRecords
'    fraDetails.Visible = False
'
'    Call EnableCmd
'    cmdCancel.Enabled = False
'    cmdSave.Enabled = False
'
'
'    With frmMain2
'        .cmdNew.Enabled = True
'        .cmdEdit.Enabled = True
'        .cmdDelete.Enabled = True
'        .cmdCancel.Enabled = False
'        .cmdSave.Enabled = False
'    End With
'
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

Private Sub cmdclose_Click()
    Unload Me
End Sub

Public Sub cmdDelete_Click()

    Dim resp As String
    On Error GoTo ErrHandler
    
    If Not currUser Is Nothing Then
        If currUser.CheckRight("EducationHistory") <> secModify Then
            MsgBox "You dont have right to delete record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
     If SelectedEmployee Is Nothing Then Exit Sub
     
    If lvwDetails.ListItems.count > 0 Then
        resp = MsgBox("Are you sure you want to delete course  " & lvwDetails.SelectedItem.SubItems(1) & vbNewLine & "from employee education history records?", vbQuestion + vbYesNo)
        If resp = vbNo Then
            Exit Sub
        End If
        Action = "DELETED EMPLOYEE'S EDUCATIONAL HISTORY; EMPLOYEE CODE: " & SelectedEmployee.EmpCode & "; EDUCATION CODE: " & lvwDetails.SelectedItem
        CConnect.ExecuteSql ("DELETE FROM edu WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & lvwDetails.SelectedItem.Tag & "'")
        
        '-------------------
        Set myEducationHistories = New EducationHistories
        myEducationHistories.GetAllEducationsHistories
        '------------
        Call DisplayRecords
    Else
        MsgBox "You have to select the eductation you would like to delete.", vbInformation
                
    End If
    Exit Sub
ErrHandler:
End Sub

Private Sub cmdDone_Click()
    PSave = True
    Call cmdCancel_Click
    PSave = False
End Sub

Public Sub cmdEdit_Click()
    On Error Resume Next
    If Not currUser Is Nothing Then
        If currUser.CheckRight("EducationHistory") <> secModify Then
            MsgBox "You dont have right to modify the record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    If SelectedEmployee Is Nothing Then Exit Sub
    
    If lvwDetails.ListItems.count < 1 Then
        MsgBox "You have to select the education you would like to edit.", vbInformation
        PSave = True
        Call cmdCancel_Click
        PSave = False
        Exit Sub
    End If
    Set rs3 = CConnect.GetRecordSet("SELECT EmpCode,Code,(select coursename from pdrEducationCourses where EducationCourseID=edu.education_id) as course,CFrom,CTo,ELevel,Award,Comments,education_id,Institution,education_id FROM Edu WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & lvwDetails.SelectedItem & "'")
    With rs3
        If .RecordCount > 0 Then
            txtCode.Text = !Code & ""
            txtCode.Tag = txtCode.Text
            
            txtLevel.Text = !ELevel & ""
            If Not IsNull(!cFrom) Then dtpFrom.value = !cFrom & ""
            If Not IsNull(!cTo) Then dtpTo.value = !cTo & ""
            txtComments.Text = !Comments & ""
            cboAward.Text = !Award & ""
            cboCourse.Text = !Course & ""
            cboCourse.Tag = IIf(IsNull(rs!education_id), 0, rs!education_id)
            txtInstitution.Text = Trim(!Institution & "")
            SaveNew = False
        Else
            MsgBox "Record not found.", vbInformation
            Set rs3 = Nothing
            PSave = True
            Call cmdCancel_Click
            PSave = False
            Exit Sub
        End If
    End With
    
    Set rs3 = Nothing
    
    Call DisableCmd
    
    fraDetails.Visible = True
    
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    SaveNew = False
    
    txtCode.Locked = False
    cboCourse.SetFocus
    
End Sub

Private Sub cmdFirst_Click()

With rsGlob
    If .RecordCount > 0 Then
        If .BOF <> True Then
            .MoveFirst
            If .BOF = True Then
                .MoveFirst
                Call DisplayRecords
            Else
                Call DisplayRecords
            End If
            
            Call FirstDisb
            
        End If
    End If
End With


End Sub

Private Sub cmdLast_Click()
With rsGlob
    If .RecordCount > 0 Then
        If .EOF <> True Then
            .MoveLast
            If .EOF = True Then
                .MoveLast
                Call DisplayRecords
            Else
                Call DisplayRecords
            End If
            
            Call LastDisb
            
        End If
    End If
End With

End Sub

Public Sub cmdNew_Click()
    If Not currUser Is Nothing Then
        If currUser.CheckRight("EducationHistory") <> secModify Then
            MsgBox "You dont have right to add new record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    Call DisableCmd
    
    txtCode.Text = loadACode
    txtLevel.Text = ""
    cboAward.Text = ""
    txtComments.Text = ""
    cboCourse.Text = ""
    
    fraDetails.Visible = True
    cmdCancel.Enabled = True
    SaveNew = True
    cmdSave.Enabled = True
    txtCode.Locked = False
    cboCourse.SetFocus

End Sub

Private Sub cmdNext_Click()
        
    With rsGlob
        If .RecordCount > 0 Then
            If .EOF <> True Then
                .MoveNext
                If .EOF = True Then
                    .MoveLast
                    Call DisplayRecords
                Else
                    Call DisplayRecords
                End If
                Call LastDisb
            End If
        End If
    End With
End Sub

Private Sub cmdPrevious_Click()

With rsGlob
    If .RecordCount > 0 Then
        If .BOF <> True Then
            .MovePrevious
            If .BOF = True Then
                .MoveFirst
                Call DisplayRecords
            Else
                Call DisplayRecords
            End If
            
            Call FirstDisb
            
        End If
    End If
End With


End Sub

Public Sub cmdSave_Click()
If txtCode.Text = "" Then
    MsgBox "Enter the eductation code.", vbExclamation
    txtCode.SetFocus
    Exit Sub
End If

If cboCourse.Text = "" Then
    MsgBox "Enter Education.", vbExclamation
    cboCourse.SetFocus
    Exit Sub
End If

If txtInstitution.Text = "" Then
    MsgBox "Enter the institution.", vbExclamation
    txtInstitution.SetFocus
    Exit Sub
End If

If dtpFrom.value > dtpTo.value Then
    MsgBox "Enter the valid start and end dates.", vbInformation
    dtpFrom.SetFocus
    Exit Sub
End If

    If SaveNew = True Then
        
        
        If SelectedEmployee Is Nothing Then
        MsgBox ("Please Select An Employee first")
        Exit Sub
        End If
'        ss = "SELECT * FROM pdrEducationHistory"
'        ss = ss & " WHERE employee_id = '" & SelectedEmployee.EmployeeID & "'"
'        ss = ss & " AND education_Code = '" & txtCode.Text & "'"
        Set rs4 = CConnect.GetRecordSet("SELECT * FROM edu WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & txtCode.Text & "'")
        
        If Not rs4 Is Nothing Then
        
        With rs4
            If .RecordCount > 0 Then
                MsgBox "eductation code already exists. Enter another one.", vbInformation
                txtCode.Text = ""
                txtCode.SetFocus
                Set rs4 = Nothing
                Exit Sub
            End If
        End With
        
        End If
        Set rs4 = Nothing
    
        
    
       
    If PromptSave = True Then
        If MsgBox("Are you sure you want to save the record.", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
        '----------retrieve the course id associated to the course selected by user. Added by kalya on 16th march,2009
    Dim Educationid As Long
    Educationid = 0
    Set MyEducationCourse = New EducationCourse
Set MyEducationCourse = MyEducationCourses.GetByEducationCourseName(cboCourse.Text)
    If Not MyEducationCourse Is Nothing Then
    Educationid = MyEducationCourse.EducationCourseID
    End If
    '------------end selecting the course id
    
    CConnect.ExecuteSql ("DELETE FROM edu WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & txtCode.Tag & "'")
    
    mySQL = "exec spinsertEducationhistory " & SelectedEmployee.EmployeeID & ",'" & txtCode.Text & "','" & Educationid & "','" & txtLevel.Text & "'," & _
"'" & Format(dtpFrom.value, "yyyy-MM-dd") & "','" & Format(dtpTo.value, "yyyy-MM-dd") & "','" & cboAward.Text & "','" & txtComments.Text & "','" & txtInstitution.Text & "'"
  
    Action = "REGISTERED EMPLOYEE'S EDUCATIONAL HISTORY; EMPLOYEE CODE: " & SelectedEmployee.EmpCode & "; EDUCATION CODE: " & txtCode.Text & "; COURSE: " & cboCourse.Text & "; LEVEL: " & txtLevel.Text & "; FROM: " & Format(dtpFrom.value, "dd-MMM-yyyy") & "; TO: " & Format(dtpTo.value, "dd-MMM-yyyy") & "; AWARD: " & cboAward.Text & "; INSTITUTION: " & txtInstitution.Text & "; COMMENTS: " & txtComments.Text
    Else
    'place code to update the record
        Dim EdID As Long
    EdID = 0
    Set MyEducationCourse = New EducationCourse
    Set MyEducationCourse = MyEducationCourses.GetByEducationCourseName(cboCourse.Text)
    If Not MyEducationCourse Is Nothing Then
    EdID = MyEducationCourse.EducationCourseID
    End If
    
    
    mySQL = "exec spupdateEducationhistory " & EdID & ", '" & txtLevel.Text & "', '" & Format(dtpFrom.value, "YYYY-MM-dd") & "', '" & Format(dtpTo.value, "yyy-mm-dd") & "', '" & cboAward.Text & "', '" & txtComments.Text & "', '" & txtInstitution.Text & "'," & SelectedEmployee.EmployeeID & ",'" & txtCode.Text & "'"
    Action = "UPDATED EMPLOYEE'S EDUCATION HISTORY; EMPLOYEE CODE:" & SelectedEmployee.EmpCode & "; EDUCATION CODE:" & txtCode.Text & " COURSE: " & cboCourse.Text & "; LEVEL: " & txtLevel.Text & "; FROM: " & Format(dtpFrom.value, "dd-MMM-yyyy") & "; TO: " & Format(dtpTo.value, "dd-MMM-yyyy") & "; AWARD: " & cboAward.Text & "; INSTITUTION: " & txtInstitution.Text & "; COMMENTS: " & txtComments.Text
    
    End If
    CConnect.ExecuteSql (mySQL)
    
    '-------------------
        Set myEducationHistories = New EducationHistories
        myEducationHistories.GetAllEducationsHistories
        '------------
    txtCode.Tag = ""
 
    If SaveNew = False Then
        PSave = True
        Call DisplayRecords
        Call cmdCancel_Click
        PSave = False
    Else
        rs2.Requery
        Call DisplayRecords
        txtCode.Text = loadACode
        cboCourse.SetFocus
        SaveNew = False
    End If
    
End Sub


Private Sub Form_Load()
    On Error GoTo ErrHandler
    oSmart.FReset Me
If oSmart.hRatio > 1.1 Then
    With frmMain2
        Me.Move .tvwMain.Width + .lvwEmp.Width + (.tvwMain.Width / 36) * 2, (.Height / 5.52) ' - 155
    End With
Else
     With frmMain2
        Me.Move .tvwMain.Width + .lvwEmp.Width + (.tvwMain.Width / 36#) * 2, .Height / 5.55
    End With
    
End If

CConnect.CColor Me, MyColor

'---------------------Load Education Type

'Set MyEducationS = New Educations
'MyEducationS.GetAllEducations

Set MyEducationCourses = New EducationCourses
MyEducationCourses.GetAllEducationCourses

Dim i As Integer
i = 1
cboCourse.Clear
'For i = 1 To MyEducationS.Count
'Set myInternalEduaction = MyEducationS.Item(i)
'cboCourse.AddItem myInternalEduaction.educationName
'Next i
For i = 1 To MyEducationCourses.count
cboCourse.AddItem MyEducationCourses.Item(i).CourseName
cboCourse.ItemData(cboCourse.NewIndex) = MyEducationCourses.Item(i).EducationCourseID
''cboCourse.Tag = MyEducationCourses.Item(i).EducationCourseID
Next i
    'Education Courses
'    Set rs = con.Execute("Select Distinct Course From Edu")
'    If Not (rs.EOF Or rs.BOF) Then
'        rs.MoveFirst
'        Do Until rs.EOF
'            cboCourse.AddItem rs!Course
'            rs.MoveNext
'        Loop
'    End If
    
    'Education Awards
    Set rs = con.Execute("Select Distinct Award From Edu")
    If Not (rs.EOF Or rs.BOF) Then
        rs.MoveFirst
        Do Until rs.EOF
           If Not IsNull(rs!Award) Then
            cboAward.AddItem rs!Award
            cboAward.Tag = "fromedu"
            End If
            rs.MoveNext
        Loop
    End If
    
    'Get Objects
    Set MyEdAwards = New EducationCourseAwards
    MyEdAwards.GetAllEducationCourseAwards
    
'---------------------


cmdCancel.Enabled = False
cmdSave.Enabled = False

Call InitGrid
'Call 'CConnect.CCon

Set rs2 = CConnect.GetRecordSet("SELECT * FROM Edu ORDER BY Code")

With rsGlob
    If .RecordCount < 1 Then
        Call DisableCmd
        Exit Sub
    End If
End With

'------------

Set myEducationHistories = New EducationHistories
myEducationHistories.GetAllEducationsHistories

'-------------

Call DisplayRecords

cmdFirst.Enabled = False
cmdPrevious.Enabled = False

Exit Sub

ErrHandler:
    MsgBox "An error has occured " & err.Description
End Sub

Private Sub Form_Resize()
    oSmart.FResize Me
    
    Me.Height = tvwMainheight - 150
    Frame1.Move Frame1.Left, 0, Frame1.Width, tvwMainheight - 150
    lvwDetails.Move lvwDetails.Left, 0, lvwDetails.Width, tvwMainheight - 150

End Sub

Private Sub InitGrid()
    With lvwDetails
        .ColumnHeaders.add , , "Code", 0
        .ColumnHeaders.add , , "Education", .Width / 6
        .ColumnHeaders.add , , "Institution", 46 * .Width / 360
        .ColumnHeaders.add , , "From", .Width / 8
        .ColumnHeaders.add , , "To", .Width / 8
        .ColumnHeaders.add , , "Level", .Width / 5
        .ColumnHeaders.add , , "Award", 46 * .Width / 360
        .ColumnHeaders.add , , "Comments", 46 * .Width / 360

        .View = lvwReport
    End With
    

End Sub

Public Sub DisplayRecords()
        On Error GoTo ErrHandler
    
    lvwDetails.ListItems.Clear
    'Call Cleartxt
    If SelectedEmployee Is Nothing Then
    Exit Sub
    End If
    Set myinternalEducationHistories = New EducationHistories
    Set myinternalEducationHistories = myEducationHistories.GetByEmployeeID(SelectedEmployee.EmployeeID)
    'Load Education Courses
    Dim Educationid As Long
    Dim streducation As String
    If myinternalEducationHistories Is Nothing Then
    Exit Sub
    End If
    Dim n As Integer
    
    For n = 1 To myinternalEducationHistories.count
        Set li = lvwDetails.ListItems.add(, , myinternalEducationHistories.Item(n).EducationCode, , 5)
        streducation = myinternalEducationHistories.Item(n).educationName
        Educationid = myinternalEducationHistories.Item(n).Educationid
        Set MyEducationCourse = MyEducationCourses.GetByEducationCourseID(Educationid)
        If Not MyEducationCourse Is Nothing Then
        streducation = MyEducationCourse.CourseName
        End If
        li.ListSubItems.add , , streducation
        li.ListSubItems.add , , myinternalEducationHistories.Item(n).Institution
        li.ListSubItems.add , , myinternalEducationHistories.Item(n).From
        li.ListSubItems.add , , myinternalEducationHistories.Item(n).dto
        li.ListSubItems.add , , myinternalEducationHistories.Item(n).educationlevel
        
        li.ListSubItems.add , , myinternalEducationHistories.Item(n).Award
        
        li.ListSubItems.add , , myinternalEducationHistories.Item(n).Comments
        
        li.Tag = myinternalEducationHistories.Item(n).EducationCode
    Next n
    
    Exit Sub
    
ErrHandler:
    MsgBox err.Description, vbExclamation, "Error"
End Sub



Private Sub Form_Unload(Cancel As Integer)

    Set rs2 = Nothing
    Set rs5 = Nothing
'    Set Cnn = Nothing
    

    frmMain2.Caption = "Infiniti HRMIS - PDR [Current User:\" & currUser.FullNames & "]"
    
End Sub

Private Sub fraList_DragDrop(Source As Control, X As Single, y As Single)

End Sub

Private Sub lvwDetails_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwDetails
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvwDetails_DblClick()
    If frmMain2.cmdEdit.Enabled = True Then
        Call frmMain2.cmdEdit_Click
    End If
End Sub

Private Sub cboAward_KeyPress(KeyAscii As Integer)
If Len(Trim(cboAward.Text)) > 49 Then
    Beep
    MsgBox "Can't enter more than 50 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("A") To Asc("Z")
  Case Asc("a") To Asc("z")
  Case Asc(" ")
  Case Asc("/")
  Case Asc("\")
  Case Asc("?")
  Case Asc(":")
  Case Asc(";")
  Case Asc(",")
  Case Asc("-")
  Case Asc("(")
  Case Asc(")")
  Case Asc("&")
  Case Asc(".")
  Case Asc("+")
  Case Asc("_")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub

Private Sub txtCode_Change()
    txtCode.Text = UCase(txtCode.Text)
    txtCode.SelStart = Len(txtCode.Text)
End Sub

Private Sub LastDisb()
With rsGlob
    If Not .EOF Then
        .MoveNext
        If .EOF Then
            cmdLast.Enabled = False
            cmdNext.Enabled = False
            cmdPrevious.Enabled = True
            cmdFirst.Enabled = True
            cmdPrevious.SetFocus
        End If
        .MovePrevious
    End If
    
    cmdPrevious.Enabled = True
    cmdFirst.Enabled = True
End With
End Sub

Private Sub FirstDisb()
With rsGlob
    If Not .BOF Then
        .MovePrevious
        If .BOF Then
            cmdLast.Enabled = True
            cmdNext.Enabled = True
            cmdPrevious.Enabled = False
            cmdFirst.Enabled = False
            cmdNext.SetFocus
        End If
        .MoveNext
    End If
    
    cmdLast.Enabled = True
    cmdNext.Enabled = True
End With
End Sub


Private Sub Cleartxt()
Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
            i.Text = ""
        End If
    Next i
    
    dtpFrom.value = Date
    dtpTo.value = Date
    
    lvwDetails.ListItems.Clear
    
End Sub

Public Sub disabletxt()
Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
            i.Locked = True
        End If
    Next i

    
End Sub

Public Sub DisableCmd()
Dim i As Object
    For Each i In Me
        If TypeOf i Is CommandButton Then
            i.Enabled = False
        End If
    Next i
End Sub

Public Sub EnableCmd()
Dim i As Object
    For Each i In Me
        If TypeOf i Is CommandButton Then
            i.Enabled = True
        End If
    Next i
    
End Sub

Public Sub FirstLastDisb()
    cmdLast.Enabled = True
    cmdNext.Enabled = True
    cmdPrevious.Enabled = True
    cmdFirst.Enabled = True
    cmdNext.SetFocus
            
With rsGlob
    If Not .BOF = True Then
        .MovePrevious
        If .BOF = True Then
            cmdLast.Enabled = True
            cmdNext.Enabled = True
            cmdPrevious.Enabled = False
            cmdFirst.Enabled = False
            cmdNext.SetFocus
        End If
        .MoveNext
    Else
        cmdLast.Enabled = True
        cmdNext.Enabled = True
        cmdPrevious.Enabled = False
        cmdFirst.Enabled = False
        cmdNext.SetFocus
    End If
    
    If Not .EOF = True Then
        .MoveNext
        If .EOF = True Then
            cmdLast.Enabled = False
            cmdNext.Enabled = False
            cmdPrevious.Enabled = True
            cmdFirst.Enabled = True
            cmdPrevious.SetFocus
        End If
        .MovePrevious
    Else
        cmdLast.Enabled = False
        cmdNext.Enabled = False
        cmdPrevious.Enabled = True
        cmdFirst.Enabled = True
        cmdPrevious.SetFocus
    End If
End With

End Sub




Private Sub txtCode_KeyPress(KeyAscii As Integer)
If Len(Trim(txtCode.Text)) > 19 Then
    Beep
    MsgBox "Can't enter more than 20 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("A") To Asc("Z")
  Case Asc("a") To Asc("z")
  Case Asc("/")
  Case Asc("\")
  Case Asc("?")
  Case Asc(":")
  Case Asc(";")
  Case Asc(",")
  Case Asc("-")
  Case Asc("(")
  Case Asc(")")
  Case Asc("&")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub

Private Sub txtComments_KeyPress(KeyAscii As Integer)
If Len(Trim(txtComments.Text)) > 198 Then
    Beep
    MsgBox "Can't enter more than 200 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("A") To Asc("Z")
  Case Asc("a") To Asc("z")
  Case Asc(" ")
  Case Asc("/")
  Case Asc("\")
  Case Asc("?")
  Case Asc(":")
  Case Asc(";")
  Case Asc(",")
  Case Asc("-")
  Case Asc("(")
  Case Asc(")")
  Case Asc("&")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub

Private Sub cboCourse_KeyPress(KeyAscii As Integer)
If Len(Trim(cboCourse.Text)) > 198 Then
    Beep
    MsgBox "Can't enter more than 200 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("A") To Asc("Z")
  Case Asc("a") To Asc("z")
  Case Asc(" ")
  Case Asc("/")
  Case Asc("\")
  Case Asc("?")
  Case Asc(":")
  Case Asc(";")
  Case Asc(",")
  Case Asc("-")
  Case Asc("(")
  Case Asc(")")
  Case Asc("&")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub

Private Sub txtLevel_KeyPress(KeyAscii As Integer)
If Len(Trim(txtLevel.Text)) > 49 Then
    Beep
    MsgBox "Can't enter more than 50 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("A") To Asc("Z")
  Case Asc("a") To Asc("z")
  Case Asc(" ")
  Case Asc("/")
  Case Asc("\")
  Case Asc("?")
  Case Asc(":")
  Case Asc(";")
  Case Asc(",")
  Case Asc("-")
  Case Asc("(")
  Case Asc(")")
  Case Asc("&")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub

Private Function loadACode() As String
    Set rs5 = CConnect.GetRecordSet("SELECT MAX(id) FROM Edu")
    If rs5.EOF = False Then
        If rs5.RecordCount > 0 And Not IsNull(rs5.Fields(0)) Then
            loadACode = "ED" & CStr(rs5.Fields(0) + 1)
        Else
            loadACode = "ED1"
        End If
    Else
        loadACode = "ED1"
    End If
    Set rs5 = Nothing
End Function

