VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim rsNextID As New ADODB.Recordset
Dim ArchiveCode As String
Private Sub Calculate_Off_Hols_Leave(empid As Integer)
Dim Hol As Integer
Dim Off As Integer
Dim m As Integer
Dim Hol_Claimed As Integer
Dim Off_Claimed As Integer
Dim Holidays As Integer
Dim OffDays As Integer
Dim RsT As New ADODB.Recordset
Dim rsEmpOff As New ADODB.Recordset
Dim rsEmpHol As New ADODB.Recordset
Dim rsArchive As New ADODB.Recordset
Dim rsApplication As New ADODB.Recordset
Dim X As Integer
Hol = 0
Off = 0

'//Calculate the due Off days

Set rsEmpOff = CConnect.GetRecordSet("select * from lvwOff_Hol where employee_id='" & empid & "' and description='Off day'")
If rsEmpOff.RecordCount > 0 Then Set rsEmpOff = CConnect.GetRecordSet("select sum(days) as OffEarned from lvwOff_Hol where employee_id='" & empid & "' and description='Off day'"): Off = Trim(rsEmpOff!OffEarned & "")

Set rsEmpOff = CConnect.GetRecordSet("select * from lvwOff_Hol_Claimed where employee_id='" & empid & "' and description='Off day'")
If rsEmpOff.RecordCount > 0 Then Set rsEmpOff = CConnect.GetRecordSet("select sum(days) as Off_Claimed from lvwOff_Hol_Claimed where employee_id='" & empid & "' and description='Off day'"): Off_Claimed = Trim(rsEmpOff!Off_Claimed & "")
OffDays = Val(Off - Off_Claimed)

'//End Calculating the due Off days

'//Calculate the due Holidays

Set rsEmpHol = CConnect.GetRecordSet("select * from lvwOff_Hol where employee_id='" & empid & "' and description='Holiday'")
If rsEmpHol.RecordCount > 0 Then Set rsEmpHol = CConnect.GetRecordSet("select sum(days) as HolEarned from lvwOff_Hol where employee_id='" & empid & "' and description='Holiday'"): Hol = Trim(rsEmpHol!HolEarned & "")

Set rsEmpHol = CConnect.GetRecordSet("select * from lvwOff_Hol_Claimed where employee_id='" & empid & "' and description='Holiday'")
If rsEmpHol.RecordCount > 0 Then Set rsEmpHol = CConnect.GetRecordSet("select sum(days) as Hol_Claimed from lvwOff_Hol_Claimed where employee_id='" & empid & "' and description='Holiday'"): Hol_Claimed = Trim(rsEmpHol!Hol_Claimed & "")
Holidays = Val(Hol - Hol_Claimed)

'//End Calculating the due Holidays

'//Remove existing record for off and holidays

CConnect.ExecuteSql "DELETE FROM EmployeeDues  WHERE employee_id='" & empid & "'"

'//Add the off and holiday dues

If OffDays > 0 Then CConnect.ExecuteSql "INSERT INTO EmployeeDues(Employee_id,Type,Description,Days) VALUES('" & empid & "','Off days','Off days'," & OffDays & ")"
If Holidays > 0 Then CConnect.ExecuteSql "INSERT INTO EmployeeDues(Employee_id,Type,Description,Days) VALUES('" & empid & "','Holidays','Holidays'," & Holidays & ")"

'//Retrieve leave balances and add to dues table

Set RsT = CConnect.GetRecordSet("select el.*,LT.name from lvwEmpLeaveDays el left outer join lvwLeaveTypes LT on el.lvcode=lt.code where el.employee_id='" & empid & "'")
If RsT.RecordCount > 0 Then
    While RsT.EOF = False
        CConnect.ExecuteSql "INSERT INTO EmployeeDues(Employee_id,Type,Description,Days) VALUES('" & empid & "','Leave','" & RsT!Name & "'," & RsT!days & ")"
        RsT.MoveNext
    Wend
End If

'//Archive leave application records

Set rsApplication = CConnect.GetRecordSet("SELECT * FROM lvwApplication Where employee_id = " & empid)
Set rsArchive = CConnect.GetRecordSet("SELECT * FROM lvwLeaveApplication_History")
With rsArchive
    .Filter = "employee_id = " & empid
    If .RecordCount > 0 Then
        .Delete
    End If
    .Requery
    .AddNew
    If rsApplication.RecordCount > 0 Then
        For X = 0 To rsApplication.Fields.Count - 1
            Debug.Print rsArchive.Fields(X).Name, rsApplication.Fields(X).Name
            rsArchive.Fields(X) = rsApplication.Fields(X)
        Next
        .Update
        rsArchive.Close
    End If
End With
CConnect.ExecuteSql "DELETE FROM lvwApplication WHERE employee_id = " & empid

'//Archive leave accrual summary

Set rsApplication = CConnect.GetRecordSet("SELECT * FROM lvwLeavePosted Where employee_id = " & empid)
Set rsArchive = CConnect.GetRecordSet("SELECT * FROM lvwLeavePosted_History")
With rsArchive
    .Filter = "employee_id = " & empid
    If .RecordCount > 0 Then
        .Delete
    End If
    .Requery
    .AddNew
    If rsApplication.RecordCount > 0 Then
        For X = 0 To rsApplication.Fields.Count - 1
            Debug.Print rsArchive.Fields(X).Name, rsApplication.Fields(X).Name
            rsArchive.Fields(X) = rsApplication.Fields(X)
        Next
        .Update
        rsArchive.Close
    End If
    
End With
CConnect.ExecuteSql "DELETE FROM lvwLeavePosted WHERE employee_id = " & empid

'//Archive Holidays and offs claim

Set rsApplication = CConnect.GetRecordSet("SELECT * FROM lvwOff_Hol_Claimed Where employee_id = " & empid)
Set rsArchive = CConnect.GetRecordSet("SELECT * FROM lvwOff_Hol_Claimed_History")
With rsArchive
    .Filter = "employee_id = " & empid
    If .RecordCount > 0 Then
        .Delete
    End If
    .Requery
    .AddNew
    If rsApplication.RecordCount > 0 Then
        For X = 0 To rsApplication.Fields.Count - 1
            Debug.Print rsArchive.Fields(X).Name, rsApplication.Fields(X).Name
            rsArchive.Fields(X) = rsApplication.Fields(X)
        Next
        .Update
        rsArchive.Close
    End If
End With
CConnect.ExecuteSql "DELETE FROM lvwOff_Hol_Claimed WHERE employee_id = " & empid

'//Archive Public Holidays Earnings

Set rsApplication = CConnect.GetRecordSet("SELECT * FROM lvwEarnedPublicHolidays Where employee_id = " & empid)
Set rsArchive = CConnect.GetRecordSet("SELECT * FROM lvwEarnedPublicHolidays_History")
With rsArchive
    .Filter = "employee_id = " & empid
    If .RecordCount > 0 Then
        .Delete
    End If
    .Requery
    .AddNew
    If rsApplication.RecordCount > 0 Then
        For X = 0 To rsApplication.Fields.Count - 1
            Debug.Print rsArchive.Fields(X).Name, rsApplication.Fields(X).Name
            rsArchive.Fields(X) = rsApplication.Fields(X)
        Next
        .Update
        rsArchive.Close
    End If
End With
CConnect.ExecuteSql "DELETE FROM lvwEarnedPublicHolidays WHERE employee_id = " & empid

'//Archive off day earnings

Set rsApplication = CConnect.GetRecordSet("SELECT * FROM lvwOffdays Where employee_id = " & empid)
Set rsArchive = CConnect.GetRecordSet("SELECT * FROM lvwOffdays_History")
With rsArchive
    .Filter = "employee_id = " & empid
    If .RecordCount > 0 Then
        .Delete
    End If
    .Requery
    .AddNew
    If rsApplication.RecordCount > 0 Then
        For X = 0 To rsApplication.Fields.Count - 1
            Debug.Print rsArchive.Fields(X).Name, rsApplication.Fields(X).Name
            rsArchive.Fields(X) = rsApplication.Fields(X)
        Next
        .Update
        rsArchive.Close
    End If
End With
CConnect.ExecuteSql "DELETE FROM lvwOffdays WHERE employee_id = " & empid

'//Archive allowances taken

Set rsApplication = CConnect.GetRecordSet("SELECT * FROM lvwAllowancesTaken Where employee_id = " & empid)
Set rsArchive = CConnect.GetRecordSet("SELECT * FROM lvwAllowancesTaken_History")
With rsArchive
    .Filter = "employee_id = " & empid
    If .RecordCount > 0 Then
        .Delete
    End If
    .Requery
    .AddNew
    If rsApplication.RecordCount > 0 Then
        For X = 0 To rsApplication.Fields.Count - 1
            Debug.Print rsArchive.Fields(X).Name, rsApplication.Fields(X).Name
            rsArchive.Fields(X) = rsApplication.Fields(X)
        Next
        .Update
        rsArchive.Close
    End If
End With
CConnect.ExecuteSql "DELETE FROM lvwAllowancesTaken WHERE employee_id = " & empid
End Sub
Private Sub Archive_Employee(empid As Integer)
Dim X As Integer
Dim RsT As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset

    On Error GoTo errHandler
    Set rsEmp = CConnect.GetRecordSet("SELECT * FROM Employee Where employee_id = " & empid)
    Set RsT = CConnect.GetRecordSet("SELECT * FROM Employee_History")
    With RsT
        .Filter = "employee_id = " & empid
        If .RecordCount > 0 Then
            .Delete
        End If
        .Requery
        .AddNew
        For X = 0 To rsEmp.Fields.Count - 1
            Debug.Print RsT.Fields(X).Name, rsEmp.Fields(X).Name
            On Error Resume Next
            RsT.Fields(X) = rsEmp.Fields(X)
        Next
        !CanReEngage = frmEmployee.chkReEngage.Value
        .Update
        RsT.Close
        'CConnect.ExecuteSql "UPDATE Employee_History SET CanReEngage=" & frmEmployee.chkReEngage.Value & " WHERE employee_id='" & empid & "'"
    End With
    CConnect.ExecuteSql "DELETE FROM Employee WHERE employee_id = " & empid
Exit Sub
errHandler:
    MsgBox Err.Description, vbInformation
End Sub

'Re-Engage Employee
Public Sub Re_Engage_Employee(empcode As String)
Dim X As Integer
Dim RsT As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset

    On Error GoTo errHandler
    Set rsEmp = CConnect.GetRecordSet("SELECT * FROM Employee_History Where empcode = '" & empcode & "'")
    Set RsT = CConnect.GetRecordSet("SELECT * FROM Employee")
    With RsT
        .Requery
        .AddNew
        For X = 0 To rsEmp.Fields.Count - 1
            Debug.Print RsT.Fields(X).Name, rsEmp.Fields(X).Name
            On Error Resume Next
            If ((LCase(rsEmp.Fields(X).Name) <> "employee_id") Or (LCase(rsEmp.Fields(X).Name) <> "empcode")) Then rsEmp.Fields(X) = RsT.Fields(X)
        Next
        RsT!empcode = empcode
        .Update
        RsT.Close
        'CConnect.ExecuteSql "UPDATE Employee_History SET CanReEngage=" & frmEmployee.chkReEngage & " WHERE employee_id='" & empid & "'"
    End With
    CConnect.ExecuteSql "UPDATE Employee SET term=0 WHERE empcode = '" & empcode & "'"
    CConnect.ExecuteSql "DELETE FROM Employee_History WHERE EmpCode='" & empcode & "'"
    With frmEmployeeHistory.lvwDetails
        .ListItems.Remove .SelectedItem.Index
    End With
    
Exit Sub
errHandler:
    MsgBox Err.Description, vbInformation
End Sub
'End Re-Engage Employee

Public Sub Transfer_Employee(ByVal empid As Integer, Optional ByVal newEmpCode As String, Optional ByVal Archive As Boolean = False)
    If Archive = True Then
        Archive_Employee empid
        Calculate_Off_Hols_Leave empid
    End If
End Sub
'
Public Sub Move_Employee_With_Changed_Code(Optional ByVal empcode As String, Optional oldCode As String, Optional ByVal tblFrom As String = "Employee", Optional ByVal tblTo As String = "Employee_History")
    Dim rec_e As New ADODB.Recordset, rec_h As New ADODB.Recordset
    On Error GoTo errHandler
    If Trim(empcode) <> "" Then
        CConnect.ExecuteSql ("UPDATE employee SET Term = 0,empcode='" & empcode & "' WHERE empCODE = '" & oldCode & "'")
        'frmEmployeeHistory.lvwDetails.ListItems.Remove frmEmployeeHistory.lvwDetails.SelectedItem.Index
    End If
    Exit Sub
errHandler:
    'MsgBox Err.Description, vbInformation
End Sub
Public Sub Move_Employee(Optional ByVal empcode As String, Optional ByVal tblFrom As String = "Employee", Optional ByVal tblTo As String = "Employee_History")
    Dim rec_e As New ADODB.Recordset, rec_h As New ADODB.Recordset
    On Error GoTo errHandler
    If Trim(empcode) <> "" Then
        CConnect.ExecuteSql ("UPDATE employee SET Term = 0 WHERE empCODE = '" & empcode & "'")
        'frmEmployeeHistory.lvwDetails.ListItems.Remove frmEmployeeHistory.lvwDetails.SelectedItem.Index
    End If
    Exit Sub
errHandler:
    'MsgBox Err.Description, vbInformation
End Sub

Private Sub Update_Employee(oldEmpCode As String, newEmpCode As String, Table As String)
CConnect.ExecuteSql "UPDATE " & Table & " SET EmpCode= '" & newEmpCode & "' WHERE EmpCode = '" & oldEmpCode & "'"
End Sub
