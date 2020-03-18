VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReengagedEmployees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ColReengagedEmployees

Private Sub Class_Initialize()
    Set ColReengagedEmployees = New Collection
End Sub

Private Sub Class_Terminate()
    Set ColReengagedEmployees = Nothing
End Sub

Public Sub Add(ByVal vnew As ReengagedEmployee)
    ColReengagedEmployees.Add vnew
End Sub

Public Function Item(ByVal vnew As Variant) As ReengagedEmployee
   Set Item = ColReengagedEmployees.Item(vnew)
End Function

Public Sub Remove(ByVal vnew As Variant)
    ColReengagedEmployees.Remove vnew
End Sub

Public Function count() As Long
    count = ColReengagedEmployees.count
End Function

Public Sub clear()
    Do Until ColReengagedEmployees.count <= 0
        ColReengagedEmployees.Remove ColReengagedEmployees.count
    Loop
End Sub

Public Function FindEmployee(ByVal vnew As Long) As ReengagedEmployee
    Dim i As Long
    Dim found As Boolean
    found = False
    For i = 1 To Me.count
        If Me.Item(i).Employee.EmployeeID = vnew Then
            Set FindEmployee = Me.Item(i)
            found = True
            Exit For
        End If
    Next i
    
    If Not found Then Set FindEmployee = Nothing
    
End Function

Public Sub GetallReengagedEmployees()
On Error GoTo Errhandler

    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim emp As ReengagedEmployee
    Set cmd = New ADODB.Command
    cmd.CommandText = "spGetAllReengagedEmployees"
    cmd.CommandType = adCmdStoredProc
    cmd.ActiveConnection = con
    Set rs = cmd.Execute
    Me.clear
    If Not (rs Is Nothing) Then
        If Not (rs.BOF And rs.EOF) Then
            rs.MoveFirst
            Do Until rs.EOF
                
                Set emp = New ReengagedEmployee
                emp.Employee.EmployeeID = rs!EmployeeID
                emp.Employee.empcode = rs!OldStaffCode
                emp.NewStaffCode = rs!NewStaffCode
                emp.Employee.DateOfDisengagement = rs!DateTerminated
                emp.Employee.DisengagementReason = rs!TerminationReason
                emp.Employee.DateOfEmployment = rs!DateEmployed
                emp.ReengagedDate = rs!DateReengaged
                emp.Employee.OrganizationUnit.OrganizationUnitID = rs!OldDepartment
                emp.Department.OrganizationUnitID = rs!NewDepartment
                emp.Employee.position.PositionID = rs!OldPosition
                emp.position.PositionID = rs!NewPosition
                emp.Employee.category.CategoryID = rs!OldGrade
                emp.Grade.CategoryID = rs!newGrade
                emp.EmpTerm.EmpTermID = rs!NewEmpTerm
                emp.Employee.EmploymentTerm.EmpTermID = rs!OldEmpTerm
                Me.Add emp
                
                rs.MoveNext
            Loop
        End If
     End If
     
     Exit Sub
Errhandler:
     MsgBox "An error has occur " & Err.Description
End Sub
