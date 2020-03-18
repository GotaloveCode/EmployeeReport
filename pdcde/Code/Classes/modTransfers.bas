Attribute VB_Name = "modTransfers"
Option Explicit


Public Sub Archive_Employee(empcode As String)

End Sub

Public Sub Transfer_Employee(oldEmpCode As String, newEmpCode As String, Optional ByVal Archive As Boolean = False)
If Archive = False Then

Else

End If
End Sub




Public Sub Update_Employee(oldEmpCode As String, newEmpCode As String, Table As String)
CConnect.ExecuteSql "UPDATE " & Table & " SET EmpCode= '" & newEmpCode & " WHERE EmpCode = '" & oldEmpCode & "'"
End Sub
