VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmployeePhoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private vEmployeeID As Long
Private vPhoto As StdPicture
Private photoPath As String
Public Property Get EmployeeID() As Long
    EmployeeID = vEmployeeID
End Property
    
Public Property Let EmployeeID(ByVal vNew As Long)
    vEmployeeID = vNew
End Property

Public Property Get Photo() As StdPicture
  Set Photo = vPhoto
End Property
    
Public Property Let Photo(ByVal vNew As StdPicture)
   Set vPhoto = vNew
End Property

Private Sub SaveEmployeePhoto()
    Dim rsEmployee As adodb.Recordset
    Dim lngOffset As Long
    Dim lngSize As Long
    Dim sPath As String
    Dim nHandle As Integer
    Dim lngSubChunks As Long
    Dim Chunk() As Byte
    Dim nFragmentOffset As Long
    Dim i As Long
    Dim lngChunks As Long
    
    On Error GoTo ErrorHandler
    Set rsCompany = New adodb.Recordset
        
    rsEmployee.Open "SELECT * FROM employees where employeeid=" & Me.EmployeeID & "", con, adOpenDynamic, adLockOptimistic
    If Not (rsEmployee.BOF Or rsEmployee.EOF) Then
        'first save the picture
        SavePicture Me.Photo, photoPath
        
        nHandle = FreeFile
        Open photoPath For Binary Access Read As nHandle
        
        lngSize = LOF(nHandle)
        If nHandle = 0 Then
            Close nHandle
        End If
            
        lngChunks = lngSize \ ChunkSize
        nFragmentOffset = lngSize Mod ChunkSize
               
        ReDim Chunk(nFragmentOffset)
        Get nHandle, , Chunk()
        rsEmployee("Logo").AppendChunk Chunk()
        ReDim Chunk(ChunkSize)
        lngOffset = nFragmentOffset
        For i = 1 To lngChunks
            Get nHandle, , Chunk()
            rsEmployee("Logo").AppendChunk Chunk()
            lngOffset = lngOffset + ChunkSize
            DoEvents
        Next
       
        rsEmployee.Update
      
    End If
   Exit Sub

ErrorHandler:
    MsgBox "An error has occurred while saving the Photo" & vbNewLine & err.Description
    
End Sub
Private Function ProcessPhotoStream() As adodb.Stream
    Dim stmPhoto As adodb.Stream
    Dim lngOffset As Long
    Dim lngSize As Long
    Dim sPath As String
    Dim nHandle As Integer
    Dim lngSubChunks As Long
    Dim Chunk() As Byte
    Dim nFragmentOffset As Long
    Dim i As Long
    Dim lngChunks As Long
    
    On Error GoTo ErrorHandler
    Set stmPhoto = New adodb.Stream
    stmPhoto.Type = adTypeBinary
    stmPhoto.Open
    
    
    'first save the picture
    SavePicture Me.Photo, photoPath
    stmPhoto.LoadFromFile photoPath
    Set ProcessPhotoStream = stmPhoto
   Exit Function

ErrorHandler:
    MsgBox "An error has occurred while saving the Photo" & vbNewLine & err.Description
    Set ProcessPhotoStream = Nothing
End Function

Public Function UpdateEmployeePhoto() As Long
    Dim CMD As adodb.Command
    Dim prmPhoto As adodb.Parameter
    Dim strmphoto As adodb.Stream
    
    On Error GoTo ErrorHandler
    
    Set CMD = New adodb.Command
    
    CMD.ActiveConnection = con
    CMD.CommandText = "spUpdateEmployeePhoto"
    CMD.CommandType = adCmdStoredProc
    
    CMD.Parameters.Append CMD.CreateParameter(, adInteger, adParamInputOutput, , Me.EmployeeID) 'EmployeeID
    If Me.Photo.Type = 0 Then
        CMD.Parameters.Append CMD.CreateParameter(, adLongVarBinary, adParamInput, Len(vbNullString) + 1, vbNullString)
    Else
        Set strmphoto = ProcessPhotoStream()
        CMD.Parameters.Append CMD.CreateParameter(, adLongVarBinary, adParamInput, strmphoto.Size + 1, strmphoto.Read)
    End If
 
    CMD.Execute
    
    'Call SaveCompanyLogo
    UpdateEmployeePhoto = 0
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred while updating Employee Photo " & vbNewLine & err.Description, vbInformation, HR_TITLE
    UpdateEmployeePhoto = err.Number
End Function



Private Sub Class_Initialize()
    Set vPhoto = New StdPicture
    photoPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "EmployeePhoto.bin"
End Sub
