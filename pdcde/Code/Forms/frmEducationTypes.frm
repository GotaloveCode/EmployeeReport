VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEducationTypes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   10785
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
      Left            =   2610
      TabIndex        =   14
      ToolTipText     =   "Move to the Last employee"
      Top             =   5610
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
      Left            =   4770
      Picture         =   "frmEducationTypes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Delete Record"
      Top             =   5610
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
      Left            =   4305
      Picture         =   "frmEducationTypes.frx":04F2
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Edit Record"
      Top             =   5610
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
      Left            =   3825
      Picture         =   "frmEducationTypes.frx":05F4
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Add New record"
      Top             =   5610
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
      Left            =   2130
      TabIndex        =   10
      ToolTipText     =   "Move to the Next employee"
      Top             =   5610
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
      Left            =   1650
      TabIndex        =   9
      ToolTipText     =   "Move to the Previous employee"
      Top             =   5610
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "CLOSE"
      Height          =   495
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5610
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
      Left            =   1170
      TabIndex        =   7
      ToolTipText     =   "Move to the First employee"
      Top             =   5610
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      Caption         =   "Education Types"
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   1650
      TabIndex        =   0
      Top             =   450
      Visible         =   0   'False
      Width           =   6015
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
         Picture         =   "frmEducationTypes.frx":06F6
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancel Process"
         Top             =   1095
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
         Picture         =   "frmEducationTypes.frx":07F8
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Save Record"
         Top             =   1095
         Width           =   510
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   135
         TabIndex        =   2
         Top             =   600
         Width           =   1110
      End
      Begin VB.TextBox txtEducationName 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1455
         TabIndex        =   1
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         Height          =   195
         Left            =   135
         TabIndex        =   6
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Left            =   1455
         TabIndex        =   5
         Top             =   360
         Width           =   420
      End
   End
   Begin MSComctlLib.ListView lvwDetails 
      Height          =   7320
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   12912
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
Attribute VB_Name = "frmEducationTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Myeducation As Education
Private Myeducations As Educations

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
        resp = MsgBox("Are you sure you want to delete Education Type  " & lvwDetails.SelectedItem.SubItems(1) & vbNewLine & "from employee education history records?", vbQuestion + vbYesNo)
        If resp = vbNo Then
            Exit Sub
        End If
        
        Action = "DELETED A EDUCATION; EDUUCATION CODE: " & lvwDetails.SelectedItem.Text & "; EDUCATION NAME: " & lvwDetails.SelectedItem.SubItems(1)
        CConnect.ExecuteSql ("DELETE FROM pdreducation WHERE EDUCATIONecode = '" & lvwDetails.SelectedItem.Text & "' AND EDUCATIONNAME = '" & lvwDetails.SelectedItem.SubItems(1) & "'")
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
        MsgBox "You have to select the Education details you would like to edit.", vbInformation
        PSave = True
        'Call cmdCancel_Click
        PSave = False
        Exit Sub
    End If
                
    Call DisableCmd
    
    Set Myeducation = New Education
    Set Myeducation = Myeducations.GetByEducationID(lvwDetails.SelectedItem.Tag)
    With Myeducation
        txtCode.Text = .EducationCode
        txtEducationName.Text = .educationName
    End With
    
    fraDetails.Visible = True
    
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    SaveNew = False
    txtCode.Locked = False
    txtEducationName.Locked = False
    
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
    txtEducationName.Text = ""
    
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
        MsgBox "Enter the education type code.", vbExclamation
        txtCode.SetFocus
        Exit Sub
    End If
    
    If txtEducationName.Text = "" Then
        MsgBox "Enter the education name", vbExclamation, "Save Error"
        Exit Sub
    End If
    
    'Assign Values to Education Course
    Set Myeducation = New Education
    With Myeducation
        .EducationCode = Trim(Replace(txtCode.Text, "'", "''"))
        .educationName = Trim(Replace(txtEducationName.Text, "'", "''"))
    End With
    
    If SaveNew = True Then
        'Saved Record
        Myeducation.ModifyEducation (0)   'SAVE
        'Update the AuditTrail
        Action = "Added an education Type: " & txtCode & " - " & txtEducationName
        currUser.AuditTrail Add_New, Action
    Else
        'Update
        Myeducation.Educationid = lvwDetails.SelectedItem.Tag
        Myeducation.ModifyEducation (1)
        'Update the AuditTrail
        Action = "Modified an education Type: " & txtCode & " - " & txtEducationName & " TO: - " & txtCode & " - " & txtEducationName
        currUser.AuditTrail Add_New, Action
    End If
       frmMain2.cmdNew.Enabled = True
       frmMain2.cmdSave.Enabled = False
       frmMain2.cmdCancel = False
    If SaveNew = False Then
        DisplayRecords
    Else
        Call DisplayRecords
    End If
    
    'Unset Controls
    txtCode.Text = ""
    txtEducationName.Text = ""
    Set Myeducation = Nothing
    fraDetails.Visible = False
    
    Exit Sub
ErrHandler:
    MsgBox "An error has occured while updating the education Types" & vbNewLine & err.Description, vbInformation, "PDR Error"
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
        .ColumnHeaders.add , , "Education Code", .Width / 7
        .ColumnHeaders.add , , "Education Name", 3 * .Width / 7
        .View = lvwReport
    End With
End Sub

Public Sub DisplayRecords()
    On Error GoTo ErrHandler
    
    lvwDetails.ListItems.Clear
    'Call Cleartxt
    
    Set Myeducation = New Education
    
    'Load Education Courses
    Set Myeducations = New Educations
    Myeducations.GetAllEducations
    
    Dim n As Integer
    
    For n = 1 To Myeducations.count
        Set li = lvwDetails.ListItems.add(, , Myeducations.Item(n).EducationCode, , 5)
        li.ListSubItems.add , , Myeducations.Item(n).educationName
        li.Tag = Trim(Myeducations.Item(n).Educationid)
    Next n
    
    Exit Sub
    
ErrHandler:
    MsgBox err.Description, vbExclamation, "Error"
End Sub

Private Sub lvwDetails_DblClick()
    Me.cmdEdit_Click
End Sub

