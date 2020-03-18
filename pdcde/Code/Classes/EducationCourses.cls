VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EducationCourses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCol As Collection

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Public Sub Add(ByVal vNew As EducationCourse)
    mCol.Add vNew
End Sub

Public Sub Remove(ByVal Index As Long)
    mCol.Remove (Index)
End Sub

Public Function Count() As Long
    Count = mCol.Count
End Function

Public Function Item(ByVal Index As Long) As EducationCourse
    Set Item = mCol.Item(Index)
End Function

Public Sub Clear()
    Dim n As Long
    
    For n = 1 To mCol.Count
        mCol.Remove n
    Next n
End Sub

Public Sub GetAllEducationCourses()
    On Error GoTo ErrHandler
    Dim CMD As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim MyEducationCourse As EducationCourse
    
    Set CMD = New ADODB.Command
        
    CMD.ActiveConnection = con
    CMD.CommandText = "spSelectAllEducationCourses"
    CMD.CommandType = adCmdStoredProc
    
    Set rs = CMD.Execute()
    
    If Not (rs.EOF Or rs.BOF) Then
        rs.MoveFirst
        Do Until rs.EOF
            Set MyEducationCourse = New EducationCourse
            
            With MyEducationCourse
                .EducationCourseID = rs!EducationCourseID
                .CourseCode = rs!CourseCode
                .CourseName = rs!CourseName
            End With
            
            Me.Add MyEducationCourse
            rs.MoveNext
        Loop
    End If
    
    Set MyEducationCourse = Nothing
    
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Function GetByEducationCourseID(ByVal EdCourseID As Long) As EducationCourse
    On Error GoTo ErrHandler
    Dim Found As Boolean, EdCourse As EducationCourse, n As Long
    
    Set EdCourse = New EducationCourse
    
    For n = 1 To Me.Count
        If (Me.Item(n).EducationCourseID = EdCourseID) Then
            Set EdCourse = Me.Item(n)
            Found = True
            Exit For
        End If
    Next n
    
    If Found Then
        Set GetByEducationCourseID = EdCourse
    Else
        Set EdCourse = Nothing
    End If
    
    Exit Function
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
End Function

Public Function GetByEducationCourseName(ByVal CourseName As String) As EducationCourse
    On Error GoTo ErrHandler
    Dim Found As Boolean, EdCourse As EducationCourse, n As Long
    
    Set EdCourse = New EducationCourse
    
    For n = 1 To Me.Count
        If (Me.Item(n).CourseName = CourseName) Then
            Set EdCourse = Me.Item(n)
            Found = True
            Exit For
        End If
    Next n
    
    If Found Then
        Set GetByEducationCourseName = EdCourse
    Else
        Set EdCourse = Nothing
    End If
    
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
End Function

