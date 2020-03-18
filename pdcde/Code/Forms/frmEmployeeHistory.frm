VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmEmployeeHistory 
   BackColor       =   &H00F2FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Employee History"
   ClientHeight    =   7305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   6000
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame fraArchive 
      BackColor       =   &H00F2FFFF&
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   10215
      Begin MSComctlLib.ListView lvwDetails 
         Height          =   5775
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   10186
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Frame fraCom 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FFFF&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   6000
         Width           =   9975
         Begin VB.CommandButton cmdExit 
            Caption         =   "&Exit"
            Height          =   435
            Left            =   8520
            TabIndex        =   2
            Top             =   200
            Width           =   1335
         End
         Begin VB.CommandButton cmdEngage 
            Caption         =   "&Engage"
            Height          =   435
            Left            =   7200
            TabIndex        =   3
            Top             =   200
            Width           =   1335
         End
         Begin VB.CommandButton cmdView 
            Caption         =   "&View"
            Height          =   435
            Left            =   5880
            TabIndex        =   4
            Top             =   200
            Width           =   1335
         End
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Archive"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "frmEmployeeHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ctran As New CTransfer

Public Sub cmdCancel_Click()

End Sub

Public Sub cmdDelete_Click()

End Sub

Public Sub cmdEdit_Click()

End Sub

Private Sub cmdEngage_Click()
Dim RsT As New ADODB.Recordset
Dim rsMove_It As ADODB.Recordset
Dim myReEngage As New CTransfer
Dim sKey As String
If lvwDetails.ListItems.Count = 0 Then
    MsgBox "There are no employees to engage.", vbInformation
    cmdExit_Click
    Exit Sub
End If
wasThere = False
Set RsT = CConnect.GetRecordSet("SELECT * FROM Employee Where EmpCode = '" & lvwDetails.SelectedItem.ListSubItems(1).Text & "' and term <> 1")
If RsT.RecordCount <> 0 Then
confirm:
    sKey = InputBox("Sorry but the Employee Code is already in use" & vbCrLf & "Please input a new Employee Code", "Employee Code in Use")
    If sKey = "" Then
        If MsgBox("Do you wish to end the engage process?", vbQuestion + vbYesNo) = vbNo Then
            GoTo confirm
        Else
            Exit Sub
        End If
    End If
    Set RsT = Nothing
    Set RsT = CConnect.GetRecordSet("SELECT * FROM Employee Where EmpCode = '" & sKey & "'")
    If RsT.RecordCount <> 0 Then
        GoTo confirm
    Else
        Set rsMove_It = CConnect.GetRecordSet("select CanReEngage from employee_history where empcode='" & lvwDetails.SelectedItem.ListSubItems(1).Text & "'")
        If rsMove_It.RecordCount > 0 Then
            If rsMove_It!CanReEngage = False Then
                ctran.Move_Employee sKey, "employee_history", "employee"
                myReEngage.Re_Engage_Employee lvwDetails.SelectedItem.ListSubItems(1)
            Else
                MsgBox "The selected employee is among the" & vbCrLf & "list of employees who cannot be re-engaged!", vbOKOnly + vbExclamation, "Process aborted"
                Exit Sub
            End If
        Else
            ctran.Move_Employee sKey, "employee_history", "employee"
            myReEngage.Re_Engage_Employee lvwDetails.SelectedItem.ListSubItems(1)
        End If
    
'        Set rsMove_It = CConnect.GetRecordSet("select CanReEngage from employee_history where empcode='" & lvwDetails.SelectedItem.ListSubItems(1).Text & "'")
'        If rsMove_It.RecordCount > 0 Then
'            If rsMove_It!CanReEngage = True Then ctran.Move_Employee sKey, "employee_history", "employee" Else MsgBox "The selected employee is among the" & vbCrLf & "list of employees who cannot be re-engaged!", vbOKOnly + vbExclamation, "Process aborted": Exit Sub
'        Else
'            ctran.Move_Employee sKey, "employee_history", "employee"
'        End If
    End If
ElseIf MsgBox("Do you wish to change the employee code?", vbQuestion + vbYesNo) = vbYes Then
change:
    sKey = InputBox("Input the new Employee Code", "New Employee Code")
    If sKey = "" Then
        If MsgBox("Do you wish to end the change employee code process?", vbQuestion + vbYesNo) = vbNo Then
            GoTo change
        Else
            Exit Sub
        End If
    End If
    Set RsT = Nothing
    Set RsT = CConnect.GetRecordSet("SELECT * FROM Employee Where EmpCode = '" & sKey & "'")
    If RsT.RecordCount <> 0 Then
        GoTo change
    Else
        wasThere = True
        Set rsMove_It = CConnect.GetRecordSet("select CanReEngage from employee where empcode='" & lvwDetails.SelectedItem.ListSubItems(1).Text & "'")
        If rsMove_It.RecordCount > 0 Then
            If rsMove_It!CanReEngage = False Then
                ctran.Move_Employee_With_Changed_Code sKey, lvwDetails.SelectedItem.ListSubItems(1), "employee_history", "employee"
                'myReEngage.Re_Engage_Employee lvwDetails.SelectedItem.ListSubItems(1)
            Else
                MsgBox "The selected employee is among the" & vbCrLf & "list of employees who cannot be re-engaged!", vbOKOnly + vbExclamation, "Process aborted"
                Exit Sub
            End If
        Else
            ctran.Move_Employee_With_Changed_Code sKey, lvwDetails.SelectedItem.ListSubItems(1), "employee_history", "employee"
            'myReEngage.Re_Engage_Employee lvwDetails.SelectedItem.ListSubItems(1)
        End If
    End If
Else
    Set rsMove_It = CConnect.GetRecordSet("select CanReEngage from employee where empcode='" & lvwDetails.SelectedItem.ListSubItems(1).Text & "'")
    
    If rsMove_It.RecordCount > 0 Then
        If rsMove_It!CanReEngage & "" = False Then
            ctran.Move_Employee lvwDetails.SelectedItem.ListSubItems(1), "employee_history", "employee"
            'myReEngage.Re_Engage_Employee lvwDetails.SelectedItem.ListSubItems(1)
        Else
            MsgBox "The selected employee is among the" & vbCrLf & "list of employees who cannot be re-engaged!", vbOKOnly + vbExclamation, "Process aborted"
            Exit Sub
        End If
    Else
        ctran.Move_Employee sKey, "employee_history", "employee"
        'myReEngage.Re_Engage_Employee lvwDetails.SelectedItem.ListSubItems(1)
    End If

'    If rsMove_It.RecordCount > 0 Then
'        If rsMove_It!CanReEngage = True Then ctran.Move_Employee lvwDetails.SelectedItem, "employee_history", "employee" Else MsgBox "The selected employee is among the" & vbCrLf & "list of employees who cannot be re-engaged!", vbOKOnly + vbExclamation, "Process aborted": Exit Sub
'    Else
'        ctran.Move_Employee lvwDetails.SelectedItem, "employee_history"
'    End If
End If
Load_lvwDetails
frmMain2.LoadMyList
End Sub

Public Sub cmdExit_Click()
    Unload Me
End Sub

Public Sub cmdNew_Click()

End Sub

Public Sub cmdSave_Click()

End Sub

Private Sub cmdView_Click()
    If lvwDetails.SelectedItem Is Nothing Then
        MsgBox "There are no Archived Records to View", vbInformation, "Employee History"
        Exit Sub
    Else
        Load frmEmployeeHistoryView
        frmEmployeeHistoryView.DisplayRecords lvwDetails.SelectedItem.Text
        Me.Hide
        frmEmployeeHistoryView.Show 1, frmMain2
        Me.Show
    End If
End Sub

Private Sub Form_Load()
'    Decla.Security Me
'    oSmart.FReset Me
    
    
    If oSmart.hRatio > 1.1 Then
        With frmMain2
            Me.Move .tvwMain.Width + .lvwEmp.Width + (.tvwMain.Width / 36) * 2, (.Height / 5.52) + 155
        End With
    Else
         With frmMain2
            Me.Move .tvwMain.Width + .lvwEmp.Width + (.tvwMain.Width / 36#) * 2, .Height / 5.55
        End With
        
    End If
    
    Load_lvwDetails
    If lvwDetails.ListItems.Count <> 0 Then
        lvwDetails.ListItems(1).Selected = True
    End If
    CConnect.CColor Me, MyColor
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    frmMain2.lvwEmp.Visible = True
    frmMain2.lvwEmp.Visible = True
    frmMain2.Caption = "Personnel Director " & App.FileDescription
    
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

Private Sub lvwDetails_ItemClick(ByVal Item As MSComctlLib.ListItem)
    frmEmployeeHistoryView.DisplayRecords lvwDetails.SelectedItem.Text
End Sub

Private Sub Load_lvwDetails()
Dim RsT As New ADODB.Recordset
Dim LI As ListItem
    With lvwDetails
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .ColumnHeaders.Clear
        .ListItems.Clear
    
        .ColumnHeaders.Add , , "EmpId", 0
        .ColumnHeaders.Add , , "Staff No", .Width / 6
        .ColumnHeaders.Add , , "Name", 19 * .Width / 120
        .ColumnHeaders.Add , , "Date Employed", 19 * .Width / 120
        .ColumnHeaders.Add , , "Termination Date", 19 * .Width / 120
        .ColumnHeaders.Add , , "Service Years", 19 * .Width / 120
        .ColumnHeaders.Add , , "Reasons", .Width / 5
    End With

    'Fully terminated employees'
    '//Commented by Juma as it causes loss of link in disengagement
'    Set RsT = CConnect.GetRecordSet("SELECT Employee_id,EmpCode,Surname,Othernames, termdate FROM Employee_History")
'    If RsT.RecordCount > 0 Then
'        With RsT
'            .MoveFirst
'            While .EOF = False
'                Set LI = lvwDetails.ListItems.Add(, "A" & !employee_id, !employee_id & "")
'                LI.ListSubItems.Add , , !empcode & ""
'                LI.ListSubItems.Add , , !SurName & " " & !OtherNames & ""
'                LI.ListSubItems.Add , , !termdate & ""
'                .MoveNext
'            Wend
'            .MoveFirst
'        End With
'    End If
    '//End of comment
    
    'Employees in the termination process'
    Set RsT = CConnect.GetRecordSet("SELECT Employee_id,EmpCode,Surname,Othernames, termdate,dleft,demployed,termReasons FROM Employee WHERE Term = 1")
    If RsT.RecordCount > 0 Then
        With RsT
            '.MoveFirst
            While .EOF = False
                Set LI = lvwDetails.ListItems.Add(, , !employee_id & "") '(, "A" & !employee_id, !employee_id & "")
                LI.ListSubItems.Add , , !empcode & ""
                LI.ListSubItems.Add , , !SurName & " " & !OtherNames & ""
                LI.ListSubItems.Add , , Format(Trim(!DEmployed & ""), "dd, MMM, yyyy")
                LI.ListSubItems.Add , , Format(Trim(!dleft & ""), "dd, MMM, yyyy")
                
                'Calculate service years
                LI.ListSubItems.Add , , IIf(DateDiff("m", IIf(IsDate(Trim(!DEmployed & "")) = True, Trim(!DEmployed & ""), Date), IIf(IsDate(Trim(!dleft & "")) = True, Trim(!dleft & ""), Date)) > 0, Round((DateDiff("m", IIf(IsDate(Trim(!DEmployed & "")) = True, Trim(!DEmployed & ""), Date), IIf(IsDate(Trim(!dleft & "")) = True, Trim(!dleft & ""), Date))) / 12, 2), 0)
                LI.ListSubItems.Add , , Trim(!TermReasons & "")
                .MoveNext
            Wend
            .MoveFirst
        End With
    End If
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
        cmdEngage.Enabled = False
End Sub

