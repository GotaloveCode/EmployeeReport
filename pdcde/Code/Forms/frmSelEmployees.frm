VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSelEmployees 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search Engine- Employee Records"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   3795
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   7
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   3360
         Width           =   1095
      End
      Begin VB.OptionButton OptName 
         BackColor       =   &H00C0E0FF&
         Caption         =   "By Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton OptCode 
         BackColor       =   &H00C0E0FF&
         Caption         =   "By Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtSearch 
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
         Left            =   4800
         TabIndex        =   1
         Top             =   240
         Width           =   1200
      End
      Begin MSComctlLib.ListView lvwDetails 
         Height          =   2730
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   4815
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgTree"
         ForeColor       =   16711680
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "Search Field"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   3480
         TabIndex        =   3
         Top             =   240
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frmSelEmployees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCont As Recordset
Dim X   '++This is the search field++

Function Load_DeptEmp_ByStructure(sDeptCode As String)
'''On Error GoTo hell
'''Dim rsQ As Recordset, EmpCount As Integer, sEmpCode As String
'''
'''InitGrid
'''
'''Set rsQ = cConnect.GetRecordSet("SELECT * from Employee where (DCode='" & sDeptCode & "') and (Term <> 'YES') ")
'''
'''If rsQ.EOF = True Then Exit Function
'''
'''sEmpCode = rsQ!EmpCode
'''
'''Set rsQ = cConnect.GetRecordSet("SELECT SEmp.*, Employee.SurName, Employee.OtherNames, Employee.Term, Employee.IDNo, cstructure.Description" & _
'''        " FROM (SEmp LEFT JOIN Employee ON SEmp.EmpCode = Employee.EmpCode) LEFT JOIN cstructure ON (SEmp.LCode = cstructure.LCode) AND (SEmp.SCode = cstructure.SCode)" & _
'''        " WHERE (((SEmp.SCode)='01')) AND (Employee.Term <> 'Yes') and (SEmp.LCode ='" & sDeptCode & "') ORDER BY SEmp.EmpCode")
'''
'''lvwDetails.ListItems.Clear
'''
'''With rsQ
'''
'''    For EmpCount = 1 To rsQ.RecordCount
'''        Set Li = lvwDetails.ListItems.Add(, , !EmpCode & "", , 5)
'''        Li.ListSubItems.Add , , !SurName & ""
'''        Li.ListSubItems.Add , , !OtherNames & ""
'''        .MoveNext
'''    Next EmpCount
'''
'''End With
'''
'''Set rsQ = Nothing

Exit Function
Hell: MsgBox Err.Description, vbCritical, "Personnel Director"
End Function


Private Sub InitGrid()
    With lvwDetails
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Emp Code", 1500
        .ColumnHeaders.Add , , "SurName", 2000
        .ColumnHeaders.Add , , "Other Names", .Width - 3500
        .View = lvwReport
    End With
End Sub

Public Sub DisplayRecords()
lvwDetails.ListItems.Clear
    With rsCont
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Set LI = lvwDetails.ListItems.Add(, , !empcode & "", , 5)
                LI.ListSubItems.Add , , !SurName & ""
                LI.ListSubItems.Add , , !OtherNames & ""
                .MoveNext
            Loop
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    strName = ""
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    On Error GoTo 10
    strName = lvwDetails.SelectedItem
    Unload Me
Exit Sub
10: MsgBox Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo Hell

    CConnect.CColor Me, MyColor
    
    Set rsCont = New Recordset
    fraDetails.Visible = True
'Set rsCont = CConnect.GetRecordSet("select * from Employee order by SurName")
    Set rsCont = CConnect.deptFilter("SELECT * FROM pVwRsGlob WHERE seq >= '" & maxCatAccess & "' AND Term <> 1 ORDER BY EmpCode")
    Call InitGrid
    Call DisplayRecords
    Set rsCont = Nothing
    
    Exit Sub
Hell: MsgBox Err.Description, vbCritical, "Search Records"
End Sub


Private Sub lvwDetails_Click()
    cmdSelect.Enabled = True
End Sub

Private Sub lvwDetails_DblClick()
    Call cmdSelect_Click
End Sub

Private Sub OptCode_Click()
    OptName.Value = False
End Sub

Private Sub OptName_Click()
    OptCode.Value = False
End Sub

Private Sub txtSearch_Change()
On Error GoTo Hell
Set rsCont = New Recordset

    X = txtSearch & "%"
    
    If OptCode.Value = True Then
'        Set rsCont = CConnect.GetRecordSet("select * from Employee where  EmpCode Like '" & X & "'" & " order by EmpCode")
        Set rsCont = CConnect.deptFilter("SELECT * FROM pVwRsGlob WHERE seq >= '" & maxCatAccess & "' AND Term <> 1 AND EmpCode Like '" & X & "' ORDER BY EmpCode")
    Else
'        Set rsCont = CConnect.GetRecordSet("select * from Employee where  SurName Like '" & X & "'" & " order by SurName")
        Set rsCont = CConnect.deptFilter("SELECT * FROM pVwRsGlob WHERE seq >= '" & maxCatAccess & "' AND Term <> 1 AND SurName Like '" & X & "' ORDER BY EmpCode")
    End If
    
    Call InitGrid
    Call DisplayRecords
    
Exit Sub
Hell: MsgBox Err.Description, vbCritical, "Search Records"
End Sub
