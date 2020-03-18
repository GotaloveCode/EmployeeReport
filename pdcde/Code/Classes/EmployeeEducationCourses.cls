VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmployeeEducationCourses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private vCol As Collection

Private Sub Class_Initialize()
Set vCol = New Collection
End Sub

Public Sub Class_Terminate()
'Set vCol = Nothing
End Sub
Public Sub add(Item As EmployeeEducationCourse)
vCol.add Item
End Sub
Public Sub remove(Index As Long)
vCol.remove (Index)
End Sub
Public Function count() As Long
count = vCol.count
End Function
Public Function Item(ByVal Index As Long) As EmployeeEducationCourse
    Set Item = vCol.Item(Index)
End Function

Public Sub Clear()
    Dim n As Long
    For n = 1 To mCol.count
        mCol.remove n
    Next n
End Sub
Public Function getAllEmployeeEducationCourses()
On Error GoTo ErrHandler
    Dim CMD As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim myemployeeeducationcourse As EmployeeEducationCourse
    
    Set CMD = New ADODB.Command
        
    CMD.ActiveConnection = con
    CMD.CommandText = "spSelectAllEmployeeEducationCourses"
    CMD.CommandType = adCmdStoredProc
    
    Set rs = CMD.Execute()
    
    If Not (rs.EOF Or rs.BOF) Then
        rs.MoveFirst
        Do Until rs.EOF
            Set myemployeeeducationcourse = New EmployeeEducationCourse
            
            With myemployeeeducationcourse
                If Not IsNull(rs!Award) Then .Award = rs!Award
                If Not IsNull(rs!cFrom) Then .cFrom = rs!cFrom
                If Not IsNull(rs!cFrom) Then .Code = rs!Code
                If Not IsNull(rs!Code) Then .Comments = rs!Comments
                If Not IsNull(rs!Comments) Then .Course = rs!Course
                If Not IsNull(rs!profid) Then .courseid = rs!profid
                If Not IsNull(rs!cTo) Then .cTo = rs!cTo
           
                 If Not IsNull(rs!EmpCode) Then .EmpCode = rs!EmpCode
                 If Not IsNull(rs!Employee_ID) Then .EmployeeID = rs!Employee_ID
                 If Not IsNull(rs!Institution) Then .Institution = rs!Institution
                 If Not IsNull(rs!Institution) Then .Institution = rs!Institution
                 If Not IsNull(rs!ELevel) Then .Level = rs!ELevel
            End With
            
      Me.add myemployeeeducationcourse
            rs.MoveNext
        Loop
    End If
    
    Set myemployeeeducationcourse = Nothing
    
    Exit Function
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
End Function
Public Function GetByEmployeeID(ByVal empid As Long) As EmployeeEducationCourses
    On Error GoTo ErrHandler
    Dim Found As Boolean, EmpCourse As EmployeeEducationCourse, n As Long
    
    Dim empCourses As EmployeeEducationCourses
    Set EmpCourse = New EmployeeEducationCourse
    Set empCourses = New EmployeeEducationCourses
    
    For n = 1 To Me.count
        If (Me.Item(n).EmployeeID = empid) Then
            Set EmpCourse = Me.Item(n)
            
            empCourses.add EmpCourse
        End If
        Dim emppid As Long
        emppid = Me.Item(n).EmployeeID
    Next n
    
    If EmpCourse Is Nothing Then
    GetByEmployeeID = Nothing
    End If
        Set GetByEmployeeID = empCourses
    
    
    Exit Function
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
End Function


Public Function GetByEmployeeName(ByVal Empname As String) As EmployeeEducationCourses
    On Error GoTo ErrHandler
    Dim Found As Boolean, EmpCourse As EmployeeEducationCourse, n As Long
    
    Dim empCourses As EmployeeEducationCourses
    Set EmpCourse = New EmployeeEducationCourse
    Set empCourses = New EmployeeEducationCourses
    
    For n = 1 To Me.count
        If (Me.Item(n).EmployeeID = Empname) Then
            Set EmpCourse = Me.Item(n)
            
            empCourses.add EmpCourse
        End If
    Next n
    
    
        Set GetByEmployeeeID = empCourses
    
    
    Exit Function
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
End Function
