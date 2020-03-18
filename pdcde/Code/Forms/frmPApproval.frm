VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmPApproval 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   Caption         =   "Nationalities"
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9975
   Icon            =   "frmPApproval.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraEmployee 
      Caption         =   "Employee List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   4005
      TabIndex        =   13
      Top             =   180
      Visible         =   0   'False
      Width           =   5565
      Begin VB.CommandButton cmdSCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4200
         TabIndex        =   16
         Top             =   6030
         Width           =   1215
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4200
         TabIndex        =   15
         Top             =   5700
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvwEmployee 
         Height          =   4890
         Left            =   120
         TabIndex        =   14
         Top             =   630
         Width           =   5310
         _ExtentX        =   9366
         _ExtentY        =   8625
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
      Begin VB.Label lblEmpCount 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1590
         TabIndex        =   18
         Top             =   270
         Width           =   645
      End
      Begin VB.Label lblECount 
         Caption         =   "Employee Count:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   105
         TabIndex        =   17
         Top             =   270
         Width           =   1485
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
            Picture         =   "frmPApproval.frx":0442
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPApproval.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPApproval.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPApproval.frx":0778
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
      TabIndex        =   8
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
         TabIndex        =   0
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
         TabIndex        =   7
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
         TabIndex        =   1
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
         TabIndex        =   2
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
         Picture         =   "frmPApproval.frx":0CBA
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "frmPApproval.frx":0DBC
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmPApproval.frx":0EBE
         Style           =   1  'Graphical
         TabIndex        =   6
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
         TabIndex        =   3
         ToolTipText     =   "Move to the Last employee"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSComctlLib.ListView lvwDetails 
         Height          =   6930
         Left            =   0
         TabIndex        =   9
         Top             =   90
         Width           =   3630
         _ExtentX        =   6403
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
      Begin MSComctlLib.ListView lvwAppovees 
         Height          =   6930
         Left            =   3675
         TabIndex        =   10
         Top             =   90
         Width           =   6255
         _ExtentX        =   11033
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
   Begin VB.CommandButton cmdSave 
      Default         =   -1  'True
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
      Left            =   4230
      Picture         =   "frmPApproval.frx":13B0
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Save Record"
      Top             =   5310
      Width           =   510
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
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
      Left            =   4725
      Picture         =   "frmPApproval.frx":14B2
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Cancel Process"
      Top             =   5310
      Width           =   495
   End
End
Attribute VB_Name = "frmPApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Sub cmdCancel_Click()
    fraEmployee.Visible = False
    
    Call EnableCmd
    cmdCancel.Enabled = False

    SaveNew = False
    
    With frmMain2
        .cmdNew3.Enabled = True
        .cmdDelete3.Enabled = True
        .cmdCancel3.Enabled = False
     
    End With
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub cmdDelete_Click()
Dim resp As String


If lvwAppovees.ListItems.Count > 0 Then
    resp = MsgBox("This will delete  " & lvwAppovees.SelectedItem & " and the corresponding employee Positions from the records. Do you wish to continue?", vbQuestion + vbYesNo)
    If resp = vbNo Then
        Exit Sub
    End If
      
    
    CConnect.ExecuteSql ("DELETE FROM PApproval WHERE PositionCode = '" & lvwDetails.SelectedItem & "' AND EmployeeCOde = '" & lvwAppovees.SelectedItem & "'")
     
   
    rs5.Requery
    
    Call LoadApprovees
        
Else
    MsgBox "You have to select the approvee you would like to delete.", vbInformation
            
End If
    
    
End Sub

Private Sub cmdDone_Click()
    PSave = True
    Call cmdCancel_Click
    PSave = False
End Sub

Public Sub cmdEdit_Click()



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
    fraEmployee.Visible = True

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

End Sub


Private Sub cmdSCancel_Click()
fraEmployee.Visible = False

PSave = True
Call cmdCancel_Click
PSave = False
    
End Sub

Private Sub cmdSelect_Click()
If lvwDetails.ListItems.Count > 0 And lvwEmployee.ListItems.Count > 0 Then
    With rs5
        If .RecordCount > 0 Then
            .Filter = "PositionCode like '" & lvwDetails.SelectedItem & "' AND EmployeeCode = '" & lvwEmployee.SelectedItem & "'"
            If .RecordCount > 0 Then
                MsgBox "Record exists", vbInformation
                .Filter = adFilterNone
                Exit Sub
            Else
                CConnect.ExecuteSql ("INSERT INTO PApproval (PositionCode, EmployeeCode, Approved) VALUES ('" & lvwDetails.SelectedItem & "','" & lvwEmployee.SelectedItem & "','No')")
                rs5.Requery
                Call LoadApprovees
            End If
        Else
            CConnect.ExecuteSql ("INSERT INTO PApproval (PositionCode, EmployeeCode, Approved) VALUES ('" & lvwDetails.SelectedItem & "','" & lvwEmployee.SelectedItem & "','No')")
            rs5.Requery
            Call LoadApprovees

        End If
    End With
    
Else
    MsgBox "No records.", vbInformation
End If
        

            
    
End Sub

Private Sub Form_Load()
Decla.Security Me
oSmart.FReset Me

If oSmart.hRatio > 1.1 Then
    With frmMain2
        Me.Move .tvwMain.Width + .tvwMain.Width / 32.75, (.Height / 5.52) '- 155
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

Set rs5 = CConnect.GetRecordSet("SELECT PApproval.PositionCode, PApproval.EmployeeCode, PApproval.Approved, Employee.SurName, Employee.OtherNames, Employee.Desig FROM PApproval INNER JOIN Employee ON PApproval.EmployeeCode = Employee.EmpCode")

Set rs2 = CConnect.GetRecordSet("SELECT * FROM Positions ORDER BY Code")

With rsGlob2
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
            Set LI = lvwEmployee.ListItems.Add(, , !empcode & "")
            LI.ListSubItems.Add , , !SurName & "" & " " & !OtherNames & ""
            LI.ListSubItems.Add , , !Desig & ""
        
            .MoveNext
        Loop
        
        lblEmpCount.Caption = .RecordCount
    End If
End With


Call DisplayRecords

cmdFirst.Enabled = False
cmdPrevious.Enabled = False

End Sub

Private Sub Form_Resize()
oSmart.FResize Me


End Sub

Private Sub InitGrid()
    With lvwDetails
        .ColumnHeaders.Add , , "Position", 2500
        .ColumnHeaders.Add , , "Vacant Posts", , vbRightJustify
        .ColumnHeaders.Add , , "Comments", 10000
                   
        .View = lvwReport
    End With
    
    With lvwEmployee
        .ColumnHeaders.Add , , "Code"
        .ColumnHeaders.Add , , "Names", 3000
        .ColumnHeaders.Add , , "Designation", 2500
                   
        .View = lvwReport
    End With
    
    With lvwAppovees
        .ColumnHeaders.Add , , "Code"
        .ColumnHeaders.Add , , "Names", 3000
        .ColumnHeaders.Add , , "Designation", 2500
        .ColumnHeaders.Add , , "Approved", , vbCenter
                   
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
            LI.ListSubItems.Add , , !Posts & ""
            LI.ListSubItems.Add , , !Comments & ""
                  
            .MoveNext
        Loop
     
    End If
End With


Call LoadApprovees


End Sub



Private Sub Form_Unload(Cancel As Integer)

    Set rs2 = Nothing
    Set rs5 = Nothing
    frmMain2.Caption = "Personnel Director " & App.FileDescription
    
End Sub

Private Sub fraList_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub lvwAppovees_DblClick()
    With rs5
        If .RecordCount > 0 Then
            .Filter = "PositionCode = '" & lvwDetails.SelectedItem & "' AND EmployeeCode = '" & lvwAppovees.SelectedItem & "'"
            If .RecordCount > 0 Then
                If !approved = "Yes" Then
                    CConnect.ExecuteSql ("UPDATE PApproval SET Approved = 'No' WHERE PositionCode = '" & lvwDetails.SelectedItem & "' AND EmployeeCode = '" & lvwAppovees.SelectedItem & "'")
                Else
                    CConnect.ExecuteSql ("UPDATE PApproval SET Approved = 'Yes' WHERE PositionCode = '" & lvwDetails.SelectedItem & "' AND EmployeeCode = '" & lvwAppovees.SelectedItem & "'")
                End If
                
                rs5.Requery
                Call LoadApprovees
            End If
            .Filter = adFilterNone
        End If
    End With
    
    With rs5
        If .RecordCount > 0 Then
            .Filter = "PositionCode = '" & lvwDetails.SelectedItem & "' AND Approved = 1"
            If .RecordCount = lvwAppovees.ListItems.Count Then
                CConnect.ExecuteSql ("UPDATE Positions SET Approved = 1 WHERE Code = '" & lvwDetails.SelectedItem & "'")
                rs2.Requery
                Call DisplayRecords
            End If
            .Filter = adFilterNone
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
                        .Find "EmpCode like '" & Sel & "'", , adSearchForward, adBookmarkFirst
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



Private Sub lvwDetails_ItemClick(ByVal Item As MSComctlLib.ListItem)
Call LoadApprovees

End Sub

Public Sub disabletxt()
Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
            i.Locked = True
        End If
    Next i

    
End Sub

Public Sub LoadApprovees()
lvwAppovees.ListItems.Clear

If lvwDetails.ListItems.Count > 0 Then
    With rs5
        If .RecordCount > 0 Then
            .Filter = "PositionCode like '" & lvwDetails.SelectedItem & "'"
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    Set LI = lvwAppovees.ListItems.Add(, , !EmployeeCode & "", , 1)
                    LI.ListSubItems.Add , , !SurName & "" & " " & !OtherNames & ""
                    LI.ListSubItems.Add , , !Desig & ""
                    LI.ListSubItems.Add , , !approved & ""
                    
                    .MoveNext
                Loop
            End If
            .Filter = adFilterNone
        End If
    End With
     
End If

End Sub

Private Sub lvwEmployee_DblClick()
    Call cmdSelect_Click
    
End Sub
