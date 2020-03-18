VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmployeesPhotos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'local variable to hold collection
Private mCol As Collection
 
Private photoPath  As String
Private stdphoto As StdPicture
Public Sub Clear()
    Do Until mCol.count <= 0
        mCol.remove mCol.count
    Loop
End Sub

 
Private Sub PopulateEmployeesphotosIntoCollection(ByVal rs As ADODB.Recordset)
    Dim i As Long
    Dim empp As EmployeePhoto
    Dim strmphoto As ADODB.Stream
    On Error GoTo ErrorHandler
    
    Me.Clear 'First remove all others
    
    If Not (rs Is Nothing) Then
        If Not (rs.BOF Or rs.EOF) Then
            'rs.MoveFirst
            Do Until rs.EOF
                Set empp = New EmployeePhoto
                With empp
                    .EmployeeID = rs!EmployeeID
                        If Not IsNull(rs!Photo) Then
                             Set strmphoto = New ADODB.Stream
                             strmphoto.Type = adTypeBinary
                             strmphoto.Open
                             strmphoto.Write rs.Fields("photo").value
                             If strmphoto.Size > 0 Then
                                 strmphoto.SaveToFile photoPath, adSaveCreateOverWrite
                                 'read the logo
                            
                                 .Photo = LoadPicture(photoPath, , vbLPColor)
                             End If
                         End If
                     
                End With
                Me.add empp
                rs.MoveNext
            Loop
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while Populating Employee photos" & vbNewLine & err.Description, vbInformation, HR_TITLE
End Sub

Public Sub UpdateEmployeesphotosInCollection(empp As EmployeePhoto)
    Dim i As Long
   
    Dim strmphoto As ADODB.Stream
    On Error GoTo ErrorHandler
    
    
    
    If (Me.count > 0) Then
    Dim i2 As Integer
      For i2 = 1 To Me.count
               
                With empp
                    If .EmployeeID = Me.Item(i2).EmployeeID Then
                    Me.remove (i2)
                    Me.add empp
                    Exit For
                    End If
                     
                End With
                
      Next i2
        
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while inserting Employee photo into collection" & vbNewLine & err.Description, vbInformation, HR_TITLE
End Sub

Public Sub GetAllEmployeesPhotos()
    Dim CodeBlock As String
    Dim CMD As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim empp As EmployeePhoto
    
    On Error GoTo ErrorHandler
    
    Me.Clear
       
    
    
 
    Set CMD = New ADODB.Command
    CMD.ActiveConnection = con
    CMD.CommandText = "spGetAllEmployeesPhotos"
    CMD.CommandType = adCmdStoredProc
    
    Set rs = CMD.Execute()
    
    If Not (rs Is Nothing) Then
        If Not (rs.BOF Or rs.EOF) Then
            rs.MoveFirst
            Do Until rs.EOF
                Set empp = New EmployeePhoto
                With empp
                    .EmployeeID = rs!EmployeeID
            If Not IsNull(rs!Photo) Then
                Set strmphoto = New ADODB.Stream
                strmphoto.Type = adTypeBinary
                strmphoto.Open
                strmphoto.Write rs.Fields("photo").value
                If strmphoto.Size > 0 Then
                    strmphoto.SaveToFile photoPath, adSaveCreateOverWrite
                    'read the logo
                     .Photo = LoadPicture(photoPath, , vbLPColor)
                End If
            End If
            End With
                Me.add empp
                rs.MoveNext
            Loop
        End If
    End If
               
 Exit Sub
 
ErrorHandler:
 MsgBox "An error has occurred while retrieving Employee Information" & vbNewLine & "CODE BLOCK: " & CodeBlock & vbNewLine & err.Description, vbInformation, HR_TITLE
    
End Sub
Public Function getEmployeePhoto(empid As Integer) As EmployeePhoto
Dim RsT As Recordset
Set getEmployeePhoto = Nothing
Dim empp As EmployeePhoto
Set RsT = con.Execute("select employeeid,photo from EmployeesPhotos where employeeid=" & empid & "")

            If Not RsT.EOF Then
                Set empp = New EmployeePhoto
                With empp
                    .EmployeeID = RsT!EmployeeID
                        If Not IsNull(RsT!Photo) Then
                             Set strmphoto = New ADODB.Stream
                             strmphoto.Type = adTypeBinary
                             strmphoto.Open
                             strmphoto.Write RsT.Fields("photo").value
                             If strmphoto.Size > 0 Then
                                 strmphoto.SaveToFile photoPath, adSaveCreateOverWrite
                                 'read the logo
                                 .Photo = LoadPicture(photoPath, , vbLPColor)
                             End If
                         End If
                     
                End With
                Set getEmployeePhoto = empp
            
           End If
End Function
Public Sub GetAccessibleEmployeesPhotosByUser(ByVal TheUserID As Long)
    Dim CodeBlock As String
    Dim CMD As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim GroupID As Long
    Dim sQL As String
    
    On Error GoTo ErrorHandler
    
    Me.Clear
       
   
    sQL = "Select dbo.hrsfnGetGroupIDOfUser(" & TheUserID & ") AS GroupID"
 
    Set CMD = New ADODB.Command
    CMD.ActiveConnection = con
    CMD.CommandText = sQL
    CMD.CommandType = adCmdText
    
    Set rs = CMD.Execute()
    GroupID = IIf(IsNull(rs!GroupID), 0, rs!GroupID)
    
    GetAccessibleEmployeesphotosByGroup GroupID
    
 Exit Sub
 
ErrorHandler:
    MsgBox "An error has occurred while retrieving Employee photos" & vbNewLine & ": " & err.Description, vbInformation, HR_TITLE
End Sub

Public Sub GetAccessibleEmployeesphotosByGroup(ByVal TheGroupID As Long)
    Dim CodeBlock As String
    Dim CMD As ADODB.Command
    Dim emp As Employee
    Dim AccessRightTypeID As Long, CmdText As String
    
    On Error GoTo ErrorHandler
    
    Set CMD = New ADODB.Command
    Set RsT = New ADODB.Recordset
    
   
    CodeBlock = "Get The Employees Photos"
    
    'Get AccessRightType
    Set RsT = con.Execute("SELECT AccessRightTypeID FROM hrsSecurityOptions")
        
    If Not (RsT.EOF Or RsT.BOF) Then
        AccessRightTypeID = RsT!AccessRightTypeID
    End If
    
    Select Case AccessRightTypeID
        Case 1
            'Employee Categories
            CmdText = "Select employeeid,photo From hrsfnGetAccessibleEmployeesPhotosByCategories(" & TheGroupID & ")"
        Case 2
            'Employee Grades
            CmdText = "Select employeeid,photo From hrsfnGetAccessibleEmployeesPhotosByGrades(" & TheGroupID & ")"
        Case 3
            'Payroll Types
            CmdText = "Select employeeid,photo From hrsfnGetAccessibleEmployeesPhotosByPayrollType(" & TheGroupID & ")"
        Case 4
            'Organization Units
            CmdText = "Select employeeid,photo From hrsfnGetAccessibleEmployeesPhotosByOU(" & TheGroupID & ")"
    End Select
    
    Set RsT = con.Execute(CmdText)
    
    If (TheGroupID = 0) Then
        'Infiniti Super Account
        Set RsT = con.Execute("Select employeeid,photo From Employees")
    End If
    
    If Not (RsT.EOF Or RsT.BOF) Then
        'populate the data into the collection
        PopulateEmployeesphotosIntoCollection RsT
    End If
               
Exit Sub
ErrorHandler:
    MsgBox "An error has occurred while retrieving Employee Information" & vbNewLine & "CODE BLOCK: " & CodeBlock & vbNewLine & err.Description, vbInformation, HR_TITLE
End Sub

Public Function FindEmployeePhoto(ByVal TheEmployeeID As Long) As EmployeePhoto
    Dim i As Long
    Dim Found As Boolean
    
    On Error GoTo ErrorHandler
    
    For i = 1 To mCol.count
        If mCol.Item(i).EmployeeID = TheEmployeeID Then
            Set FindEmployeePhoto = mCol.Item(i)
            Found = True
            Exit For
        End If
    Next i
    
    If Not Found Then
        Set FindEmployeePhoto = Nothing
    End If
    
    Exit Function
ErrorHandler:
    Set FindEmployeePhoto = Nothing
    MsgBox (err.Description)
End Function

 
 
 
 
Public Sub add(ByVal TheEmployeephoto As EmployeePhoto)
    mCol.add TheEmployeephoto
End Sub

 
Public Property Get Item(vntIndexKey As Variant) As EmployeePhoto
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get count() As Long
    count = mCol.count
End Property

Public Sub remove(vntIndexKey As Variant)

    mCol.remove vntIndexKey
End Sub

 

Private Sub Class_Initialize()
    On Error Resume Next
    Set stdphoto = New StdPicture
    photoPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "EmployeePhoto.bin"
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
 
    
End Sub

 
 



