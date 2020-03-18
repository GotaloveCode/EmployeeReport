VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Educations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCol As Collection
Public errormessage As String

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Public Sub add(ByVal vnew As Education)
    mCol.add vnew
End Sub

Public Sub remove(ByVal Index As Long)
    mCol.remove (Index)
End Sub

Public Function count() As Long
    count = mCol.count
End Function

Public Function Item(ByVal Index As Long) As Education
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

Public Sub GetAllEducations()
    On Error GoTo ErrHandler
    Dim CMD As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim MyEducation As Education
    
    Set CMD = New ADODB.Command
        
    CMD.ActiveConnection = con
    CMD.CommandText = "spSelectAllEducations"
    CMD.CommandType = adCmdStoredProc
    
    Set rs = CMD.Execute()
    
    If Not (rs.EOF Or rs.BOF) Then
        rs.MoveFirst
        Do Until rs.EOF
            Set MyEducation = New Education
            
            With MyEducation
               If Not IsNull(rs!educationid) Then .educationid = rs!educationid
               If Not IsNull(rs!EducationCode) Then .EducationCode = rs!EducationCode
               If Not IsNull(rs!educationName) Then .educationName = rs!educationName
            End With
            
            Me.add MyEducation
            rs.MoveNext
        Loop
    End If
    
    Set MyEducation = Nothing
    
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Function GetByEducationID(ByVal EdID As Long) As Education
    On Error GoTo ErrHandler
    Dim Found As Boolean, Ed As Education, n As Long
    
    Set Ed = New Education
    
    For n = 1 To Me.count
        If (Me.Item(n).educationid = EdID) Then
            Set Ed = Me.Item(n)
            Found = True
            Exit For
        End If
    Next n
    
    If Found Then
        Set GetByEducationID = Ed
    Else
        Set GetByEducationID = Nothing
    End If
    
    Exit Function
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
End Function

Public Function GetByEducationName(ByVal educationName As String) As Education
    On Error GoTo ErrHandler
    Dim Found As Boolean, Ed As Education, n As Long
    
    Set Ed = New Education
    
    For n = 1 To Me.count
        If (Me.Item(n).educationName = educationName) Then
            Set Ed = Me.Item(n)
            Found = True
            Exit For
        End If
    Next n
    
    If Found Then
        Set GetByEducationName = Ed
    Else
        Set Ed = Nothing
    End If
    
    Dim edd As Integer
    edd = Ed.educationid
    Exit Function
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
End Function


