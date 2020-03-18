Attribute VB_Name = "modTextWriter"
Option Explicit
Public Omnis_ActionTag As String
Dim MyFileDatabase As New FileSystemObject, strData As TextStream, sData As String, j As Integer '++monte caters for the text files++

Public Sub Create_Omnis_Folder()
On Error GoTo Hell

If MyFileDatabase.FolderExists(App.Path & "\Omnis TxtFiles\") Then
    '++monte++
Else
    MyFileDatabase.CreateFolder (App.Path & "\Omnis TxtFiles\")
End If

Exit Sub
Hell: MsgBox Err.Description, vbCritical, "Text File Writer"
End Sub

Public Sub Employees_TextFile()
On Error GoTo Hell
    
    Get_TextFile "Employees"
'    Write_Employee_Details_To_TextFile Omnis_ActionTag

Exit Sub
Hell: MsgBox Err.Description, vbCritical, "Text File Writer"
End Sub

Public Sub Departments_TextFile()
On Error GoTo Hell
    
    Get_TextFile "Departments"
    Write_Department_Details_To_TextFile Omnis_ActionTag

Exit Sub
Hell: MsgBox Err.Description, vbCritical, "Text File Writer"
End Sub

Public Sub Company_TextFile()
On Error GoTo Hell
    
    Get_TextFile "Company"
    Write_Company_Details_To_TextFile Omnis_ActionTag

Exit Sub
Hell: MsgBox Err.Description, vbCritical, "Text File Writer"
End Sub

Function Get_TextFile(sTxtFile)
On Error GoTo Hell
   
Create_Omnis_Folder
   
If MyFileDatabase.FileExists(App.Path & "\Omnis TxtFiles\" & sTxtFile & ".txt") Then
    '++monte++
Else
    MyFileDatabase.CreateTextFile (App.Path & "\Omnis TxtFiles\" & sTxtFile & ".txt")
End If
    
Exit Function
Hell: MsgBox Err.Description, vbCritical, "Text File Writer"
End Function

Function Write_Employee_Details_To_TextFile(sTag As String)
On Error GoTo Hell
    
    Set strData = MyFileDatabase.OpenTextFile(App.Path & "\Omnis TxtFiles" & "\Employees.txt", ForAppending, False, TristateMixed)
       
    Set rs = CConnect.GetRecordSet("SELECT Employee.*, CStructure.Code, CStructure.Description" & _
            " FROM (Employee LEFT JOIN SEmp ON Employee.employee_id = SEmp.employee_id) LEFT JOIN CStructure ON SEmp.LCode = CStructure.LCode" & _
            " WHERE Employee.EmpCode='" & frmEmployee.txtEmpCode & "'" & _
            " ORDER BY Employee.EmpCode")
    
    strData.WriteLine sTag & Chr(9) & rs!empcode & Chr(9) & rs!SurName & "" & Chr(9) & rs!OtherNames & "" & Chr(9) & rs!DCode & "" & Chr(9) & rs!BasicPay & "" & Chr(9) & rs!ECategory & ""

    Set rs = Nothing
    Set strData = Nothing

Exit Function
Hell: MsgBox Err.Description, vbCritical, "Text File Writer"
End Function

Function Write_Department_Details_To_TextFile(sTag As String)
On Error GoTo Hell
    
Set strData = MyFileDatabase.OpenTextFile(App.Path & "\Omnis TxtFiles" & "\Departments.txt", ForAppending, False, TristateMixed)
   
    With frmCStructure
        If Not IsNumeric(.txtCode) Then
            strData.WriteLine sTag & Chr(9) & .txtCode & Chr(9) & .txtDesc & ""
        End If
    End With
    
Set strData = Nothing

Exit Function
Hell: MsgBox Err.Description, vbCritical, "Text File Writer"
End Function

Function Write_Company_Details_To_TextFile(sTag As String)
On Error GoTo Hell
    
    Set strData = MyFileDatabase.OpenTextFile(App.Path & "\Omnis TxtFiles" & "\Company.txt", ForAppending, False, TristateMixed)
        strData.WriteLine sTag & Chr(9) & frmGenOpt.txtCName & Chr(9) & Format(Date, "mm/yyyy")
    Set strData = Nothing

Exit Function
Hell: MsgBox Err.Description, vbCritical, "Text File Writer"
End Function
