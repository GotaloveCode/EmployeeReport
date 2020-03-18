VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmProf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Professional Qualifications"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7845
   Icon            =   "frmProf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   7845
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
            Picture         =   "frmProf.frx":0442
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProf.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProf.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProf.frx":0778
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7900
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   7800
      Begin MSComDlg.CommonDialog Cdl 
         Left            =   4905
         Top             =   5340
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
         TabIndex        =   10
         ToolTipText     =   "Move to the First employee"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "CLOSE"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   17
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
         TabIndex        =   11
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
         TabIndex        =   12
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
         Left            =   2775
         Picture         =   "frmProf.frx":0CBA
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Picture         =   "frmProf.frx":0DBC
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Picture         =   "frmProf.frx":0EBE
         Style           =   1  'Graphical
         TabIndex        =   16
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
         TabIndex        =   13
         ToolTipText     =   "Move to the Last employee"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame fraDetails 
         Appearance      =   0  'Flat
         Caption         =   "Profesional Qualifications"
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
         Height          =   4785
         Left            =   600
         TabIndex        =   19
         Top             =   630
         Visible         =   0   'False
         Width           =   6285
         Begin VB.ComboBox txtAward 
            Height          =   315
            Left            =   1200
            TabIndex        =   30
            Text            =   "Combo1"
            Top             =   2640
            Width           =   4935
         End
         Begin VB.ComboBox cbocourses 
            Height          =   315
            Left            =   1440
            TabIndex        =   28
            Top             =   600
            Width           =   4695
         End
         Begin VB.TextBox txtInstitution 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   3
            Top             =   1440
            Width           =   4695
         End
         Begin VB.TextBox txtCourse 
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
            Left            =   1440
            TabIndex        =   2
            Top             =   1080
            Width           =   4650
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
            Left            =   5640
            Picture         =   "frmProf.frx":13B0
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Cancel Process"
            Top             =   4200
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
            Left            =   5145
            Picture         =   "frmProf.frx":14B2
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Save Record"
            Top             =   4200
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
            Height          =   660
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   3480
            Width           =   5985
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
            Height          =   285
            Left            =   1440
            TabIndex        =   1
            Top             =   240
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
            Height          =   285
            Left            =   3150
            TabIndex        =   6
            Top             =   2160
            Width           =   2925
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
            Left            =   120
            TabIndex        =   4
            Top             =   2160
            Width           =   1455
            _ExtentX        =   2566
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
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   63963139
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
            Left            =   1680
            TabIndex        =   5
            Top             =   2160
            Width           =   1440
            _ExtentX        =   2540
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
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   63963139
            CurrentDate     =   37673
            MinDate         =   21916
         End
         Begin VB.Label Label8 
            Caption         =   "Program"
            Height          =   375
            Left            =   240
            TabIndex        =   29
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Institution"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Left            =   1800
            TabIndex        =   26
            Top             =   1920
            Width           =   180
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
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
            Left            =   240
            TabIndex        =   25
            Top             =   1800
            Width           =   360
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Course"
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
            Left            =   240
            TabIndex        =   24
            Top             =   1080
            Width           =   510
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
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
            TabIndex        =   23
            Top             =   3240
            Width           =   750
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
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
            Left            =   120
            TabIndex        =   22
            Top             =   2640
            Width           =   465
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
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
            Left            =   240
            TabIndex        =   21
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
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
            Left            =   3360
            TabIndex        =   20
            Top             =   1920
            Width           =   375
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
Attribute VB_Name = "frmProf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim MyEducationCourse As HRCORE.EducationCourse
Dim MyEducationCourses As HRCORE.EducationCourses
Dim myemployeeeducationcourse As EmployeeEducationCourse
Dim myemployeeeducationcourses As EmployeeEducationCourses
Dim myinternalemployeeeducationcourses As EmployeeEducationCourses
Dim myinternalemployeeeeducationcourse As EmployeeEducationCourse
Dim MyEdAwards As EducationCourseAwards

Private Sub cbocourses_Click()
Dim n As Integer
 txtAward.Clear
 For n = 1 To MyEdAwards.count
    If MyEdAwards.Item(n).EducationCourseOBJ.CourseName = cbocourses.Text Then
 
        txtAward.AddItem MyEdAwards.Item(n).AwardName
        txtAward.Tag = MyEdAwards.Item(n).EducationAwardID
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
            If currUser.CheckRight("Qualification") <> secModify Then
                MsgBox "You dont have right to delete record. Please liaise with the security admin"
                Exit Sub
            End If
        End If
     If SelectedEmployee Is Nothing Then Exit Sub
     
    If lvwDetails.ListItems.count > 0 Then
        resp = MsgBox("Are you sure you want to delete  " & lvwDetails.SelectedItem & " from the records?", vbQuestion + vbYesNo)
        If resp = vbNo Then
            Exit Sub
        End If
       Action = "DELETED EMPLOYEE PROFESSIONAL QUALIFICATION; EMPLOYEE CODE: " & SelectedEmployee.EmpCode & "; QUALIFICATION CODE: " & lvwDetails.SelectedItem
        
        CConnect.ExecuteSql ("DELETE FROM Prof WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & lvwDetails.SelectedItem & "'")
        rs2.Requery
        Set myemployeeeducationcourses = New EmployeeEducationCourses
        myemployeeeducationcourses.getAllEmployeeEducationCourses
        Call DisplayRecords
    Else
        MsgBox "You have to select the professional qualification you would like to delete.", vbInformation
                
    End If
    Exit Sub
ErrHandler:
End Sub


Public Sub cmdEdit_Click()
    On Error Resume Next
    
    If Not currUser Is Nothing Then
            If currUser.CheckRight("Qualification") <> secModify Then
                MsgBox "You dont have right to modify the record. Please liaise with the security admin"
                Exit Sub
            End If
    End If
    
    If SelectedEmployee Is Nothing Then Exit Sub
    
    If lvwDetails.ListItems.count < 1 Then
        MsgBox "You have to select the professional qualification  you would like to edit.", vbInformation
        PSave = True
        Call cmdCancel_Click
        PSave = False
        Exit Sub
    End If
        
    Set rs3 = CConnect.GetRecordSet("SELECT * FROM Prof WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & lvwDetails.SelectedItem & "'")
    With rs3
        If .RecordCount > 0 Then
            txtCode.Text = !Code & ""
            txtCode.Tag = txtCode.Text
            txtLevel.Text = !ELevel & ""
            If Not IsNull(!cFrom) Then dtpFrom.value = !cFrom & ""
            If Not IsNull(!cTo) Then dtpTo.value = !cTo & ""
            txtComments.Text = !Comments & ""
            txtAward.Text = !Award & ""
            txtCourse.Text = !Course & ""
            txtInstitution.Text = !Institution & ""
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
    txtCode.SetFocus

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
        If currUser.CheckRight("Qualification") <> secModify Then
            MsgBox "You dont have right to add new record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    Call DisableCmd
    
    txtCode.Text = loadACode
    txtCode.Locked = False
    txtLevel.Text = ""
    txtAward.Text = ""
    txtComments.Text = ""
    txtCourse.Text = ""
    dtpFrom.value = Date
    dtpTo.value = Date
    'dtpFrom.Value = Date
    'dtpTo.Value = Date
    fraDetails.Visible = True
    cmdCancel.Enabled = True
    SaveNew = True
    cmdSave.Enabled = True
    txtCourse.SetFocus
    
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
    If SelectedEmployee Is Nothing Then Exit Sub
    
    If txtCode.Text = "" Then
        MsgBox "Enter the professional qualification  code.", vbExclamation
        txtCode.SetFocus
        Exit Sub
    End If
    
    If txtCourse.Text = "" Then
        MsgBox "Enter the course.", vbExclamation
        txtCourse.SetFocus
        Exit Sub
    End If
    
    If Trim(txtInstitution.Text) = "" Then
        MsgBox "Enter the Institution", vbExclamation
        txtInstitution.SetFocus
        Exit Sub
    End If
    
    If dtpFrom.value > dtpTo.value Then
        MsgBox "Enter the valid start and end dates.", vbInformation
        dtpFrom.SetFocus
        Exit Sub
    End If

    If SaveNew = True Then
        
        Set rs4 = CConnect.GetRecordSet("SELECT * FROM Prof WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & txtCode.Text & "'")
        With rs4
            If .RecordCount > 0 Then
                MsgBox "professional qualification  code already exists. Enter another one.", vbInformation
                txtCode.Text = ""
                txtCode.SetFocus
                Set rs4 = Nothing
                Exit Sub
            End If
        End With
        Set rs4 = Nothing
    End If
       
    If PromptSave = True Then
        If MsgBox("Are you sure you want to save the record.", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    CConnect.ExecuteSql ("DELETE FROM Prof WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & txtCode.Tag & "'")
    
    Dim profid As Long
    profid = 0
    Set MyEducationCourse = New HRCORE.EducationCourse
    Set MyEducationCourse = MyEducationCourses.GetByEducationCourseName(cbocourses.Text)
    If Not MyEducationCourse Is Nothing Then
    profid = MyEducationCourse.EducationCourseID
    End If
    mySQL = "INSERT INTO Prof (employee_id, Code,profid,Course, ELevel, CFrom, CTo, Award, Comments,institution)" & _
                        " VALUES(" & SelectedEmployee.EmployeeID & ",'" & txtCode.Text & "'," & profid & ",'" & txtCourse.Text & "','" & txtLevel.Text & "'," & _
                        "'" & Format(dtpFrom.value, "yyyy-MM-dd") & "','" & Format(dtpTo.value, "yyyy-MM-dd") & "','" & txtAward.Text & "','" & txtComments.Text & "','" & Trim(txtInstitution.Text) & "')"
    
    Action = "ADDED EMPLOYEE PROFESSIONAL QUALIFICATIONS; EMPLOYEE CODE: " _
    & SelectedEmployee.EmpCode & "; CODE: " & txtCode.Text _
    & "; COURSE: " & txtCourse.Text & "; EDUCATIONAL LEVEL: " _
    & txtLevel.Text & "; FROM: " _
    & Format(dtpFrom.value, "dd-MMM-yyyy") _
    & "; TO: " & Format(dtpTo.value, "dd-MMM-yyyy") _
    & "; AWARDS: " & txtAward.Text & "; COMMENTS: " _
    & txtComments.Text & "; INSTITUTION: " & txtInstitution.Text
    
    CConnect.ExecuteSql (mySQL)
    txtCode.Tag = ""
    rs2.Requery
    
    Set myemployeeeducationcourses = New EmployeeEducationCourses
    myemployeeeducationcourses.getAllEmployeeEducationCourses
        
    If SaveNew = False Then
        PSave = True
        Call DisplayRecords
        Call cmdCancel_Click
        PSave = False
    Else
        rs2.Requery
        Call DisplayRecords
        txtCode.Text = loadACode
        txtCourse.SetFocus
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
    
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    Call InitGrid
    'Call CConnect.CCon
    Set rs2 = CConnect.GetRecordSet("SELECT * FROM Prof ORDER BY CFrom")
    With rsGlob
        If .RecordCount < 1 Then
            Call DisableCmd
            Exit Sub
        End If
    End With

'-------------------
Set MyEducationCourses = New EducationCourses
MyEducationCourses.GetAllEducationCourses
Set MyEdAwards = New EducationCourseAwards
MyEdAwards.GetAllEducationCourseAwards

If Not MyEducationCourses Is Nothing Then
Dim k As Integer
cbocourses.Clear
For k = 1 To MyEducationCourses.count
cbocourses.AddItem MyEducationCourses.Item(k).CourseName
Next k
End If
Set myemployeeeducationcourses = New EmployeeEducationCourses
myemployeeeducationcourses.getAllEmployeeEducationCourses
'-------------
Call DisplayRecords

cmdFirst.Enabled = False
cmdPrevious.Enabled = False
    Exit Sub
ErrHandler:
    MsgBox "An errr has occured ERROR DESCRIPTION  " & err.Description
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
        .ColumnHeaders.add , , "Programme(Course)", 4000
        .ColumnHeaders.add , , "From", 1400
        .ColumnHeaders.add , , "To", 1400
        .ColumnHeaders.add , , "Level", 3000
        .ColumnHeaders.add , , "Award", 3500
        .ColumnHeaders.add , , "Comments", 3500
        .ColumnHeaders.add , , "Institution", 3500
                
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
    Set myinternalemployeeeducationcourses = New EmployeeEducationCourses
    Set myinternalemployeeeducationcourses = myemployeeeducationcourses.GetByEmployeeID(SelectedEmployee.EmployeeID)
    'Load Education Courses
    Dim courseid As Long
    Dim strcourse As String
    If myinternalemployeeeducationcourses Is Nothing Then
    Exit Sub
    End If
    Dim n As Integer
    
    For n = 1 To myinternalemployeeeducationcourses.count
        Set li = lvwDetails.ListItems.add(, , myinternalemployeeeducationcourses.Item(n).Code, , 5)
        strcourse = myinternalemployeeeducationcourses.Item(n).Course
        courseid = myinternalemployeeeducationcourses.Item(n).courseid
        
        Set MyEducationCourse = MyEducationCourses.GetByEducationCourseID(courseid)
        
        If Not MyEducationCourse Is Nothing Then
        strcourse = MyEducationCourse.CourseName & " ( " & myinternalemployeeeducationcourses.Item(n).Course & " )"
        End If
        li.ListSubItems.add , , strcourse
        li.ListSubItems.add , , myinternalemployeeeducationcourses.Item(n).cFrom
        li.ListSubItems.add , , myinternalemployeeeducationcourses.Item(n).cTo
        li.ListSubItems.add , , myinternalemployeeeducationcourses.Item(n).Level
        li.ListSubItems.add , , myinternalemployeeeducationcourses.Item(n).Award
        li.ListSubItems.add , , myinternalemployeeeducationcourses.Item(n).Comments
        li.ListSubItems.add , , myinternalemployeeeducationcourses.Item(n).Institution
        li.Tag = myinternalemployeeeducationcourses.Item(n).Code
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

Private Sub txtAward_KeyPress(KeyAscii As Integer)
If Len(Trim(txtAward.Text)) > 49 Then
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
  Case Asc("+")
  Case Asc("_")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub

Private Sub txtCourse_KeyPress(KeyAscii As Integer)
If Len(Trim(txtCourse.Text)) > 198 Then
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
  Case Asc("+")
  Case Asc("_")
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
  Case Asc("+")
  Case Asc("_")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub

Private Function loadACode() As String
    Set rs5 = CConnect.GetRecordSet("SELECT MAX(id) FROM Prof")
    If rs5.EOF = False Then
        If rs5.RecordCount > 0 And Not IsNull(rs5.Fields(0)) Then
            loadACode = "PR" & CStr(rs5.Fields(0) + 1)
        Else
            loadACode = "PR1"
        End If
    Else
        loadACode = "PR1"
    End If
    Set rs5 = Nothing
End Function

