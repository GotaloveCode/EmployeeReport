VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmReligionSetUp 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   Caption         =   "Religions"
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9975
   Icon            =   "frmReligionSetUp.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Religion Set Up"
      ForeColor       =   &H80000008&
      Height          =   3030
      Left            =   2175
      TabIndex        =   14
      Top             =   1020
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
         TabIndex        =   0
         Top             =   645
         Width           =   2670
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
         Height          =   825
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   1215
         Width           =   5715
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
         Left            =   4875
         Picture         =   "frmReligionSetUp.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Save Record"
         Top             =   2415
         Width           =   510
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
         Left            =   5370
         Picture         =   "frmReligionSetUp.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cancel Process"
         Top             =   2415
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "Religion"
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
         TabIndex        =   16
         Top             =   405
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
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
         Left            =   135
         TabIndex        =   15
         Top             =   990
         Width           =   750
      End
   End
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
            Picture         =   "frmReligionSetUp.frx":0646
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReligionSetUp.frx":0758
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReligionSetUp.frx":086A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReligionSetUp.frx":097C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   7020
      Left            =   15
      TabIndex        =   12
      Top             =   -90
      Width           =   9930
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
         TabIndex        =   4
         ToolTipText     =   "Move to the First employee"
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
         TabIndex        =   11
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
         TabIndex        =   5
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
         TabIndex        =   6
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
         Picture         =   "frmReligionSetUp.frx":0EBE
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "frmReligionSetUp.frx":0FC0
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Picture         =   "frmReligionSetUp.frx":10C2
         Style           =   1  'Graphical
         TabIndex        =   10
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
         TabIndex        =   7
         ToolTipText     =   "Move to the Last employee"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSComctlLib.ListView lvwDetails 
         Height          =   6930
         Left            =   0
         TabIndex        =   13
         Top             =   120
         Width           =   9930
         _ExtentX        =   17515
         _ExtentY        =   12224
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
Attribute VB_Name = "frmReligionSetUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

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
'    SaveNew = False
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

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub cmdDelete_Click()
Dim resp As String


If lvwDetails.ListItems.Count > 0 Then
    resp = MsgBox("This will delete  " & lvwDetails.SelectedItem & " and the corresponding employee Religion from the records. Do you wish to continue?", vbQuestion + vbYesNo)
    If resp = vbNo Then
        Exit Sub
    End If
      
    Action = "DELETED A RELIGION; CODE: " & lvwDetails.SelectedItem & "; DESCRIPTION: " & lvwDetails.SelectedItem.ListSubItems(1)
    
    CConnect.ExecuteSql ("DELETE FROM Religion WHERE Code = '" & lvwDetails.SelectedItem & "'")
     
   
    rs2.Requery
    
    Call DisplayRecords
        
Else
    MsgBox "You have to select the Religion  you would like to delete.", vbInformation
            
End If
    
    
End Sub

Private Sub cmdDone_Click()
    PSave = True
    Call cmdCancel_Click
    PSave = False
End Sub

Public Sub cmdEdit_Click()


If lvwDetails.ListItems.Count < 1 Then
    MsgBox "You have to select the Religion you would like to edit.", vbInformation
    PSave = True
    Call cmdCancel_Click
    PSave = False
    Exit Sub
End If


Set rs3 = CConnect.GetRecordSet("SELECT * FROM Religion WHERE Code = '" & lvwDetails.SelectedItem & "'")


With rs3
    If .RecordCount > 0 Then
        txtCode.Text = Trim(!code & "")
        txtCode.Tag = Trim(!code & "")
        txtComments.Text = !Comments & ""

        
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
'txtCode.Locked = True


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
Call DisableCmd
txtCode.Text = ""
txtComments.Text = ""
fraDetails.Visible = True
cmdCancel.Enabled = True
SaveNew = True
cmdSave.Enabled = True
txtCode.Locked = False
txtCode.SetFocus

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
    MsgBox "Enter the Religion type code.", vbExclamation
    txtCode.SetFocus
    Exit Sub
End If



    If SaveNew = True Then
        
        Set rs4 = CConnect.GetRecordSet("SELECT * FROM Religion WHERE Code = '" & txtCode.Text & "'")
        
        
        With rs4
            If .RecordCount > 0 Then
                MsgBox "Religion code already exists. Enter another one.", vbInformation
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
            
    
    CConnect.ExecuteSql ("DELETE FROM Religion WHERE Code = '" & txtCode.Tag & "'")
    
    mySQL = "INSERT INTO Religion (Code, Comments)" & _
                        " VALUES('" & Replace(txtCode.Text, "'", "''") & "','" & Replace(txtComments.Text, "'", "''") & "')"
    
    Action = "ADDED A RELIGION; CODE: " & txtCode.Text & "; DESCRIPTION: " & txtComments.Text
    
    With txtCode
        .Text = ""
        .Tag = ""
    End With
    
    CConnect.ExecuteSql (mySQL)
    

    rs2.Requery
    
    If SaveNew = False Then
        PSave = True
        Call cmdCancel_Click
        PSave = False
    Else
        rs2.Requery
        Call DisplayRecords
        txtCode.SetFocus
        
        
    End If
    
    
End Sub


Private Sub Form_Load()
Decla.Security Me
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
'Call 'CConnect.CCon


Set rs2 = CConnect.GetRecordSet("SELECT * FROM Religion ORDER BY Code")

Call DisplayRecords

cmdFirst.Enabled = False
cmdPrevious.Enabled = False

End Sub

Private Sub Form_Resize()
oSmart.FResize Me


End Sub

Private Sub InitGrid()
    With lvwDetails
        .ColumnHeaders.Add , , "Religion", 2500
        .ColumnHeaders.Add , , "Comments", 10000
   
                
        .View = lvwReport
    End With
    

End Sub

Public Sub DisplayRecords()
lvwDetails.ListItems.Clear
Call Cleartxt

With rs2
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
            Set LI = lvwDetails.ListItems.Add(, , !code & "", , 5)
            LI.ListSubItems.Add , , !Comments & ""
                  
            .MoveNext
        Loop
     
    End If
End With
 


End Sub



Private Sub Form_Unload(Cancel As Integer)

    Set rs2 = Nothing
    Set rs5 = Nothing
'    Set Cnn = Nothing
    

    frmMain2.Caption = "Personnel Director " & App.FileDescription
    
End Sub

Private Sub fraList_DragDrop(Source As Control, X As Single, Y As Single)

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



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Search"
            Me.MousePointer = vbHourglass
    
            frmSearch.Show vbModal
            
            If Not Sel = "" Then
                With rsGlob
                    If .RecordCount > 0 Then
                        .MoveFirst
                        .Find "employee_id like '" & Sel & "'", , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            Call DisplayRecords
                            Call FirstLastDisb
                        End If
                    End If
                End With
                
            End If
      
            Me.MousePointer = 0
    End Select
End Sub







Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error GoTo errHandler
Dim myfile As String
Dim ss As String
Select Case ButtonMenu.Key
    Case "EmpLeaves"
        Me.MousePointer = vbHourglass
        Set a = New Application
        Set R = a.OpenReport(App.Path & "\Leave Reports\Employee Leaves.rpt")
        
'          If Not frmRange.txtFrom.Text = "" And Not frmRange.txtTo.Text = "" Then
'              If Not frmRange.cboPCode = "" Then
'                  SQL = "{Employee.PayrollNo} in '" & frmRange.txtFrom.Text & "' to '" & frmRange.txtTo.Text & "' and {Payslip.YYear} in " & frmRange.txtFYear.Text & " to " & frmRange.txtTYear.Text & " and {Payslip.MMonth} in " & FMonth & " to " & TMonth & " and {EmpPayPoint.PPCode} = '" & frmRange.cboPCode.Text & "'"
'              Else
'                  SQL = "{Employee.PayrollNo} in '" & frmRange.txtFrom.Text & "' to '" & frmRange.txtTo.Text & "' and {Payslip.YYear} in " & frmRange.txtFYear.Text & " to " & frmRange.txtTYear.Text & " and {Payslip.MMonth} in " & FMonth & " to " & TMonth & ""
'              End If
'
'              R.RecordSelectionFormula = SQL
'          Else
'              If Not frmRange.cboPCode = "" Then
'                  SQL = "{Payslip.YYear} in " & frmRange.txtFYear.Text & " to " & frmRange.txtTYear.Text & " and {Payslip.MMonth} in " & FMonth & " to " & TMonth & " and {EmpPayPoint.PPCode} = '" & frmRange.cboPCode.Text & "'"
'              Else
'                  SQL = "{Payslip.YYear} in " & frmRange.txtFYear.Text & " to " & frmRange.txtTYear.Text & " and {Payslip.MMonth} in " & FMonth & " to " & TMonth & ""
'              End If
'
'              R.RecordSelectionFormula = SQL
'          End If
    
      'SQL = "{Payslip.YYear} in " & frmRange.txtFYear.Text & " to " & frmRange.txtTYear.Text & " and {Payslip.MMonth} in " & FMonth & " to " & TMonth & " and {EmpPayPoint.PPCode} = '" & frmRange.cboPCode.Text & "'"
      'R.RecordSelectionFormula = MySql
      
      R.ReadRecords

      With frmReports.CRViewer1
          .ReportSource = R
          .ViewReport
      End With

      frmReports.Show vbModal
        Me.MousePointer = 0
    Case "LeaveEmp"
        Me.MousePointer = vbHourglass
        Set a = New Application
        Set R = a.OpenReport(App.Path & "\Leave Reports\Leaves Employee.rpt")
        
'          If Not frmRange.txtFrom.Text = "" And Not frmRange.txtTo.Text = "" Then
'              If Not frmRange.cboPCode = "" Then
'                  SQL = "{Employee.PayrollNo} in '" & frmRange.txtFrom.Text & "' to '" & frmRange.txtTo.Text & "' and {Payslip.YYear} in " & frmRange.txtFYear.Text & " to " & frmRange.txtTYear.Text & " and {Payslip.MMonth} in " & FMonth & " to " & TMonth & " and {EmpPayPoint.PPCode} = '" & frmRange.cboPCode.Text & "'"
'              Else
'                  SQL = "{Employee.PayrollNo} in '" & frmRange.txtFrom.Text & "' to '" & frmRange.txtTo.Text & "' and {Payslip.YYear} in " & frmRange.txtFYear.Text & " to " & frmRange.txtTYear.Text & " and {Payslip.MMonth} in " & FMonth & " to " & TMonth & ""
'              End If
'
'              R.RecordSelectionFormula = SQL
'          Else
'              If Not frmRange.cboPCode = "" Then
'                  SQL = "{Payslip.YYear} in " & frmRange.txtFYear.Text & " to " & frmRange.txtTYear.Text & " and {Payslip.MMonth} in " & FMonth & " to " & TMonth & " and {EmpPayPoint.PPCode} = '" & frmRange.cboPCode.Text & "'"
'              Else
'                  SQL = "{Payslip.YYear} in " & frmRange.txtFYear.Text & " to " & frmRange.txtTYear.Text & " and {Payslip.MMonth} in " & FMonth & " to " & TMonth & ""
'              End If
'
'              R.RecordSelectionFormula = SQL
'          End If
    
      'SQL = "{Payslip.YYear} in " & frmRange.txtFYear.Text & " to " & frmRange.txtTYear.Text & " and {Payslip.MMonth} in " & FMonth & " to " & TMonth & " and {EmpPayPoint.PPCode} = '" & frmRange.cboPCode.Text & "'"
      'R.RecordSelectionFormula = MySql
      
      R.ReadRecords

      With frmReports.CRViewer1
          .ReportSource = R
          .ViewReport
      End With

      frmReports.Show vbModal
        Me.MousePointer = 0

End Select
Exit Sub

errHandler:
If Err.Description = "File not found." Then
    Cdl.DialogTitle = "Select the report to show"
    Cdl.InitDir = App.Path & "/Leave Reports"
    Cdl.Filter = "Reports {* .rpt|* .rpt"
    Cdl.ShowOpen
    myfile = Cdl.FileName
    If Not myfile = "" Then
        Resume
    Else
        Me.MousePointer = 0
    End If
Else
    MsgBox Err.Description, vbInformation
    Me.MousePointer = 0
End If
End Sub

Private Sub txtADays_Change()
'If Val(txtADays.Text) > Val(txtDays.Text) Then
'    txtADays.Text = txtDays.Text
'    txtADays.SelStart = Len(txtADays.Text)
'End If
End Sub

Private Sub txtADays_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
        Case Asc(".")
        Case Asc("-")
        Case Is = 8
        Case Else
        Beep
        KeyAscii = 8
        
    End Select
End Sub

Private Sub txtDays_Keypress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtCode_Change()
'    txtCode.Text = UCase(txtCode.Text)
'    txtCode.SelStart = Len(txtCode.Text)
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




'Private Sub txtCode_KeyPress(KeyAscii As Integer)
''If Len(Trim(txtCode.Text)) > 19 Then
''    Beep
''    MsgBox "Can't enter more than 20 characters", vbExclamation
''    KeyAscii = 8
''End If
''
''Select Case KeyAscii
''  Case Asc("0") To Asc("9")
''  Case Asc("A") To Asc("Z")
''  Case Asc("a") To Asc("z")
''  Case Asc("/")
''  Case Asc("\")
''  Case Asc("?")
''  Case Asc(":")
''  Case Asc(";")
''  Case Asc(",")
''  Case Asc("-")
''  Case Asc("(")
''  Case Asc(")")
''  Case Asc("&")
''  Case Asc(".")
''  Case Is = 8
''  Case Else
''      Beep
''      KeyAscii = 0
''End Select
'End Sub

Private Sub txtComments_KeyPress(KeyAscii As Integer)
'If Len(Trim(txtComments.Text)) > 198 Then
'    Beep
'    MsgBox "Can't enter more than 200 characters", vbExclamation
'    KeyAscii = 8
'End If
'
'Select Case KeyAscii
'  Case Asc("0") To Asc("9")
'  Case Asc("A") To Asc("Z")
'  Case Asc("a") To Asc("z")
'  Case Asc(" ")
'  Case Asc("/")
'  Case Asc("\")
'  Case Asc("?")
'  Case Asc(":")
'  Case Asc(";")
'  Case Asc(",")
'  Case Asc("-")
'  Case Asc("(")
'  Case Asc(")")
'  Case Asc("&")
'  Case Asc(".")
'  Case Is = 8
'  Case Else
'      Beep
'      KeyAscii = 0
'End Select
End Sub


