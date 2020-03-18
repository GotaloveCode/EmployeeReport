VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EducationHistories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mCol As Collection
Public errormessage As String

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Public Sub add(ByVal Item As EducationHistory)
mCol.add Item
End Sub

Public Sub remove(ByVal Index As Long)
    mCol.remove (Index)
End Sub

Public Function count() As Long
    count = mCol.count
End Function

Public Function Item(ByVal Index As Long) As EducationHistory
    Set Item = mCol.Item(Index)
End Function

Public Sub Clear()
    Dim n As Long
    
    For n = 1 To mCol.count
        mCol.remove n
    Next n
End Sub
Public Function getErrorMessage() As String
getErrorMessage = errormessage
End Function

Public Sub GetAllEducationsHistories()
   '' On Error GoTo errHandler
      On Error Resume Next
    Dim CMD As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim myeducationhist As EducationHistory
    
    
    Set CMD = New ADODB.Command
        
    CMD.ActiveConnection = con
    CMD.CommandText = "spSelectAllEducationHistories"
    CMD.CommandType = adCmdStoredProc
    
    Set rs = CMD.Execute()
    
    If Not (rs.EOF Or rs.BOF) Then
        rs.MoveFirst
        Do Until rs.EOF
        Set myeducationhist = New EducationHistory
            
            With myeducationhist
                
              If Not IsNull(rs!Award) Then .Award = rs!Award
                '.Award = "to be retrieved"
              If Not IsNull(rs!Comments) Then .Comments = rs!Comments
              If Not IsNull(rs!cTo) Then .dto = CDate(rs!cTo)
              If Not IsNull(rs!Code) Then .EducationCode = rs!Code
              If Not IsNull(rs!education_id) Then .Educationid = rs!education_id
                ''
                
                .educationName = "to be retrieved"
                .educationName = rs!Course
                ''
               If Not IsNull(rs!ELevel) Then .educationlevel = rs!ELevel
               If Not IsNull(rs!Employee_ID) Then .EmployeeID = rs!Employee_ID
               If Not IsNull(rs!cFrom) Then .From = rs!cFrom
               If Not IsNull(rs!Institution) Then .Institution = rs!Institution
                
                
            End With
            
            Me.add myeducationhist
            rs.MoveNext
        Loop
    End If
    
    Set myeducationhist = Nothing
    
    Exit Sub
'errHandler:
'    MsgBox err.Description, vbExclamation, "Error"
End Sub

Public Function GetByEmployeeID(ByVal empid As Long) As EducationHistories
    On Error GoTo errHandler
    Dim Found As Boolean, n As Long
    Dim edhists As EducationHistories
    Dim edhist As EducationHistory
    
    Set edhist = New EducationHistory
    Set edhists = New EducationHistories
    
    For n = 1 To Me.count
        
            If Me.Item(n).EmployeeID = empid Then
            Set edhist = Me.Item(n)
            edhists.add edhist
            End If
        
    Next n
    Dim gg As Integer
'    gg = edhists.Item(3).
'    Dim strr As String
'    strr = edhists.Item(3).Institution
   Set GetByEmployeeID = edhists
    
    Exit Function
errHandler:
    MsgBox err.Description, vbExclamation, "Error"
End Function

Public Function GetByEmployeeName(ByVal educationName As String) As EducationHistories
  On Error GoTo errHandler
    Dim Found As Boolean, n As Long
    Dim edhists As EducationHistories
    Dim edhist As EducationHistory
    
    Set edhist = New EducationHistory
    Set edhists = New EducationHistories
    
    For n = 1 To Me.count
        
            If Me.Item(i).educationName = educationName Then
            Set edhist = Me.Item(i)
            edhists.add (edhist)
            End If
        
    Next n
    
   Set GetByEmployeeName = edhists
    
    Exit Function
errHandler:
    MsgBox err.Description, vbExclamation, "Error"
End Function




