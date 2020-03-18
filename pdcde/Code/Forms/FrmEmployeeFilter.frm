VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmEmployeeFilter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Filter Employees"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5970
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
   ScaleHeight     =   6570
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ChkShowDisEngaged 
      Appearance      =   0  'Flat
      Caption         =   "Show disengaged employees"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   0
      Width           =   3495
   End
   Begin VB.CheckBox ChkAll 
      Appearance      =   0  'Flat
      Caption         =   "Select All Employees"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   1935
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2025
      TabIndex        =   2
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton CmdReport 
      Caption         =   "&Generate Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   105
      TabIndex        =   1
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   45
      TabIndex        =   0
      Top             =   360
      Width           =   5655
      Begin MSComctlLib.ListView LvwEmps 
         Height          =   4935
         Left            =   105
         TabIndex        =   3
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   8705
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "EmpCode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   7011
         EndProperty
      End
      Begin VB.Label LblSelected 
         AutoSize        =   -1  'True
         Caption         =   "Selected:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   120
         TabIndex        =   6
         Top             =   5280
         Width           =   675
      End
   End
End
Attribute VB_Name = "FrmEmployeeFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Emps As HRCORE.Employees
Dim i As Integer, NoChecked As Integer

Private Sub ChkAll_Click()
    On Error GoTo ErrorHandler
    
    If (Me.ChkAll.value = vbChecked) Then
        Me.ChkAll.Caption = "Unselect all employees"
        lblSelected.Caption = "Selected: " & Me.LvwEmps.ListItems.count & " OF " & LvwEmps.ListItems.count
        NoChecked = Me.LvwEmps.ListItems.count
    Else
        Me.ChkAll.Caption = "Select all employees"
        lblSelected.Caption = "Selected: 0 OF " & LvwEmps.ListItems.count
        NoChecked = 0
    End If
    
    For i = 1 To Me.LvwEmps.ListItems.count
        Me.LvwEmps.ListItems.Item(i).Checked = Me.ChkAll.value
    Next i
    
    Exit Sub
ErrorHandler:
    MsgBox "Error Has Occured:" & vbCrLf & err.Description, vbExclamation, "PDR"
End Sub

Private Sub ChkShowDisEngaged_Click()
    On Error GoTo ErrorHandler
    
    If (Me.ChkShowDisEngaged.value = vbChecked) Then
        Me.ChkShowDisEngaged.Caption = "Do not show disengaged employees"
        LoadEmployeesInformationIncludingDisengaged
    Else
        Me.ChkShowDisEngaged.Caption = "Show disengaged employees"
        LoadEmployeesInformation
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error Has Occured:" & vbCrLf & err.Description, vbExclamation, "PDR"
End Sub

Private Sub CmdReport_Click()
    On Error GoTo ErrorHandler
    
    If (CheckItemsOnList) = 0 Then
        MsgBox "No employee(s) selected for viewing on the report", vbExclamation, "PDR"
        Exit Sub
    End If
    
    'Create Formula String: A list of Codes
    Set rs = CConnect.GetRecordSet("DELETE FROM pdrtmpEmpCodesForExport")
    
    For i = 1 To Me.LvwEmps.ListItems.count
        If LvwEmps.ListItems.Item(i).Checked = True Then
            Set rs = CConnect.GetRecordSet("INSERT INTO pdrtmpEmpCodesForExport(EmpCode) VALUES ('" & LvwEmps.ListItems.Item(i).Text & "')")
        End If
    Next i
    
    Unload Me 'GoTo The Report Now
    
    Set R = CrtImportToPD
    ShowReport R, , True
    
    Exit Sub
ErrorHandler:
    MsgBox "Error Has Occured:" & vbCrLf & err.Description, vbExclamation, "PDR"
End Sub

Private Sub LoadEmployeesInformation()
    On Error GoTo ErrorHandler
    Dim ItemX As ListItem
    
    Set Emps = New HRCORE.Employees
    Emps.GetAccessibleEmployeesByUser currUser.UserID
    
    LvwEmps.ListItems.Clear
    
    With LvwEmps
        For i = 1 To Emps.count
            If Not (Emps.Item(i).IsDisengaged) Then
                Set ItemX = .ListItems.add(, , Emps.Item(i).EmpCode)
                ItemX.SubItems(1) = Emps.Item(i).SurName & " ," & UCase(Emps.Item(i).OtherNames)
            End If
        Next i
    End With
    
    lblSelected.Caption = "Selected: 0 OF " & LvwEmps.ListItems.count
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An Error Has Occured:" & vbCrLf & err.Description, vbExclamation, "PDR"
End Sub

Private Sub Form_Load()
    'Initialize and Load Employees
    LoadEmployeesInformation
End Sub

Private Sub LoadEmployeesInformationIncludingDisengaged()
    On Error GoTo ErrorHandler
    Dim ItemX As ListItem
    
    Set Emps = New HRCORE.Employees
    ''Emps.GetAllEmployees
    Emps.GetAccessibleEmployeesByUser currUser.UserID
    LvwEmps.ListItems.Clear
    
    With LvwEmps
        For i = 1 To Emps.count
            Set ItemX = .ListItems.add(, , Emps.Item(i).EmpCode)
            .ListItems.Item(i).SubItems(1) = Emps.Item(i).SurName & " ," & UCase(Emps.Item(i).OtherNames)
        Next i
    End With
    
    lblSelected.Caption = "Selected: 0 OF " & LvwEmps.ListItems.count
    ChkAll.value = vbUnchecked
    
    Exit Sub
ErrorHandler:
    MsgBox "An Error Has Occured:" & vbCrLf & err.Description, vbExclamation, "PDR"
End Sub

Private Sub LvwEmps_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo ErrorHandler
          
    lblSelected.Caption = "Selected:" & CheckItemsOnList & " OF " & Me.LvwEmps.ListItems.count
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An Error Has Occured:" & vbCrLf & err.Description, vbExclamation, "PDR"
End Sub

Private Function CheckItemsOnList() As Long
    On Error GoTo ErrorHandler
    
    For i = 1 To Me.LvwEmps.ListItems.count Step 1
        If Me.LvwEmps.ListItems.Item(i).Checked = True Then CheckItemsOnList = CheckItemsOnList + 1
    Next i
    
    Exit Function

ErrorHandler:
    MsgBox "An error has occured: " & err.Description, vbExclamation, "PDR"
End Function
