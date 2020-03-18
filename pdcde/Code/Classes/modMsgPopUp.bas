Attribute VB_Name = "modMsgPopUp"
Option Explicit
Public m_iActiveAlertWindows As Long
Private m_iBackColor As Long


Function Get_PopUp_Color() '++Colour Choice ++
m_iBackColor = RGB(160, 195, 255)

'    Select Case Index
'    Case 0 '/* Yellow
'        m_iBackColor = &HC0FFFF
'    Case 1 '/* Blue
'        m_iBackColor = RGB(160, 195, 255)
'    Case 2 '/* Red
'        m_iBackColor = RGB(255, 200, 200)
'    Case 3 '/* Green
'        m_iBackColor = RGB(200, 255, 200)
'    End Select
End Function

Function Auto_Close_PopUp()
Dim AlertWindow As frmAlertWindow
Dim SMessage As String
  
    Set AlertWindow = New frmAlertWindow
    SMessage = "Current Users" & vbCrLf & Get_Current_Users
    Get_PopUp_Color
        
    AlertWindow.DisplayMessage SMessage, 4, _
        1, True, 0, m_iBackColor, 1
End Function

Function Get_Current_Users() As String
Dim sUser As String, j As Integer, rsAllUsers As Recordset, Default_time, Last_Login
Default_time = "00:00:00": Last_Login = DateAdd("d", -1, Now)
Last_Login = Format(Last_Login, "dd/mm/yyyy") & " " & Default_time
Set rsAllUsers = CConnect.GetRecordSet("select UID from SECURITY WHERE subsystem = '" & SubSystem & "' and status='LOGGED IN'")
For j = 1 To rsAllUsers.RecordCount
    If j = 1 Then sUser = rsAllUsers!UID
    If j > 1 Then sUser = sUser & vbCrLf & rsAllUsers!UID
    rsAllUsers.MoveNext
Next j
Get_Current_Users = sUser
End Function

Function Load_EmpDefined_details(sEmpCode As String)
On Error GoTo Hell
Dim rsDefDetails As Recordset

Set rsDefDetails = CConnect.GetRecordSet("SELECT * FROM tblEmpDefinedDetails where EmpCode = '" & sEmpCode & "' ORDER BY Detail_Description")

    If rsDefDetails.EOF = True Then
        Dim rsQ As Recordset, DetailTypes As Integer
        Set rsQ = CConnect.GetRecordSet("SELECT * FROM DTypes order by Description")
        For DetailTypes = 1 To rsQ.RecordCount
            If Not IsNull(rsQ!code) Then
                strQ = "insert into tblEmpDefinedDetails (EmpCode,Detail_Code,Detail_Description) values ('" & sEmpCode & "','" & rsQ!code & "','" & rsQ!Description & "')"
                CConnect.ExecuteSql (strQ)
            End If
        rsQ.MoveNext
        Next DetailTypes
        Set rsQ = Nothing
        
        Set rsDefDetails = CConnect.GetRecordSet("SELECT * FROM tblEmpDefinedDetails where EmpCode = '" & sEmpCode & "' ORDER BY Detail_Description")

    End If
    
'++ Display the records
With frmDDetails.lvwDetails
.ListItems.Clear
    While rsDefDetails.EOF = False
        Set LI = .ListItems.Add(, , rsDefDetails!Detail_Code & "", , 5)
        LI.ListSubItems.Add , , rsDefDetails!Detail_Description & ""
        LI.ListSubItems.Add , , rsDefDetails!Details & ""
        LI.ListSubItems.Add , , rsDefDetails!Comments & ""
        rsDefDetails.MoveNext
    Wend
End With
Set rsDefDetails = Nothing
'++ Display the records
    
Exit Function
Hell: MsgBox Err.Description, vbCritical, "Employee Defined Details"
End Function

