VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReengagedEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private objDepartment As HRCORE.OrganizationUnit
Private objGrade As HRCORE.EmployeeCategory
Private objPosition As HRCORE.JobPosition
Private objEmployee As HRCORE.Employee
Private dtReengagedDate As Date
Private strNewStaffCode As String
Private objEmpTerm As HRCORE.EmploymentTerm

Public Property Get EmpTerm() As HRCORE.EmploymentTerm
    Set EmpTerm = objEmpTerm
End Property

Public Property Set EmpTerm(ByVal data As HRCORE.EmploymentTerm)
    Set objEmpTerm = data
End Property

Public Property Get NewStaffCode() As String
     NewStaffCode = strNewStaffCode
End Property

Public Property Let NewStaffCode(ByVal data As String)
    strNewStaffCode = data
End Property

Public Property Get ReengagedDate() As Date
     ReengagedDate = dtReengagedDate
End Property

Public Property Let ReengagedDate(ByVal data As Date)
    dtReengagedDate = data
End Property

Public Property Get position() As HRCORE.JobPosition
    Set position = objPosition
End Property

Public Property Set position(ByVal data As HRCORE.JobPosition)
    Set objPosition = data
End Property

Public Property Get Grade() As HRCORE.EmployeeCategory
    Set Grade = objGrade
End Property

Public Property Set Grade(ByVal data As HRCORE.EmployeeCategory)
    Set objGrade = data
End Property

Public Property Get Department() As HRCORE.OrganizationUnit
    Set Department = objDepartment
End Property

Public Property Set Department(ByVal data As HRCORE.OrganizationUnit)
    Set objDepartment = data
End Property

Public Property Get Employee() As HRCORE.Employee
    Set Employee = objEmployee
End Property

Public Property Set Employee(ByVal data As HRCORE.Employee)
    Set objEmployee = data
End Property

Public Function Insert() As Boolean
    On Error GoTo ErrHandler
    Dim CMD As ADODB.Command
    Set CMD = New ADODB.Command
    CMD.ActiveConnection = con
    CMD.CommandText = "spInsertReEngagedEmployeeInfo"
    CMD.CommandType = adCmdStoredProc
    CMD.Parameters.Append CMD.CreateParameter(, adInteger, adParamInput, , Me.Employee.EmployeeID)
    CMD.Parameters.Append CMD.CreateParameter(, adVarChar, adParamInput, 20, Me.Employee.empcode)
    CMD.Parameters.Append CMD.CreateParameter(, adVarChar, adParamInput, 20, Me.NewStaffCode)
    CMD.Parameters.Append CMD.CreateParameter(, adInteger, adParamInput, , Me.Employee.position.PositionID)
    CMD.Parameters.Append CMD.CreateParameter(, adInteger, adParamInput, , Me.position.PositionID)
    CMD.Parameters.Append CMD.CreateParameter(, adInteger, adParamInput, , Me.Employee.category.CategoryID)
    CMD.Parameters.Append CMD.CreateParameter(, adInteger, adParamInput, , Me.Grade.CategoryID)
    CMD.Parameters.Append CMD.CreateParameter(, adInteger, adParamInput, , Me.Employee.OrganizationUnit.OrganizationUnitID)
    CMD.Parameters.Append CMD.CreateParameter(, adInteger, adParamInput, , Me.Department.OrganizationUnitID)
    CMD.Parameters.Append CMD.CreateParameter(, adDate, adParamInput, , Me.Employee.DateOfDisengagement)
    CMD.Parameters.Append CMD.CreateParameter(, adDate, adParamInput, , Me.ReengagedDate)
    CMD.Parameters.Append CMD.CreateParameter(, adVarChar, adParamInput, 50, Me.Employee.disengagementReason)
    CMD.Parameters.Append CMD.CreateParameter(, adDate, adParamInput, , Me.Employee.DateOfEmployment)
    CMD.Parameters.Append CMD.CreateParameter(, adInteger, adParamInput, , Me.Employee.EmploymentTerm.EmpTermID)
    CMD.Parameters.Append CMD.CreateParameter(, adInteger, adParamInput, , Me.EmpTerm.EmpTermID)
    
    CMD.Execute
    Insert = True
    Exit Function
ErrHandler:
    Insert = False
    MsgBox "An error has occured " & Err.Description
End Function

Private Sub Class_Initialize()
    Set objDepartment = New HRCORE.OrganizationUnit
    Set objGrade = New HRCORE.EmployeeCategory
    Set objPosition = New HRCORE.JobPosition
    Set objEmployee = New HRCORE.Employee
    Set objEmpTerm = New HRCORE.EmploymentTerm
End Sub

Private Sub Class_Terminate()
    Set objDepartment = Nothing
    Set objGrade = Nothing
    Set objPosition = Nothing
    Set objEmployee = Nothing
    Set objEmpTerm = Nothing
End Sub
