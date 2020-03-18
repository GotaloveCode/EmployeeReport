VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EducationCourseAwards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCol As Collection

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Public Sub Add(ByVal vNew As EducationCourseAward)
    mCol.Add vNew
End Sub

Public Sub Remove(ByVal Index As Long)
    mCol.Remove (Index)
End Sub

Public Function Count() As Long
    Count = mCol.Count
End Function

Public Function Item(ByVal Index As Long) As EducationCourseAward
    Set Item = mCol.Item(Index)
End Function

Public Sub Clear()
    Dim n As Long
    
    For n = 1 To mCol.Count
        mCol.Remove n
    Next n
End Sub

Public Sub GetAllEducationCourseAwards()
    On Error GoTo ErrHandler
    Dim CMD As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim MyEdAward As EducationCourseAward
    Dim MyEdCourses As New EducationCourses
    
    MyEdCourses.GetAllEducationCourses
    
    Set CMD = New ADODB.Command
        
    CMD.ActiveConnection = con
    CMD.CommandText = "spSelectAllEducationAwards"
    CMD.CommandType = adCmdStoredProc
    
    Set rs = CMD.Execute()
    
    If Not (rs.EOF Or rs.BOF) Then
        rs.MoveFirst
        Do Until rs.EOF
            Set MyEdAward = New EducationCourseAward
            
            With MyEdAward
                .AwardCode = rs!AwardCode
                .AwardName = rs!AwardName
                .EducationAwardID = rs!EducationAwardID
                .EducationCourseID = rs!EducationCourseID
                Set .EducationCourseOBJ = MyEdCourses.GetByEducationCourseID(rs!EducationCourseID)
            End With
            
            Me.Add MyEdAward
            rs.MoveNext
        Loop
    End If
    
    Set MyEdAward = Nothing
    
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Function GetByEducationAwardID(ByVal EdAwardID As Long) As EducationCourseAward
    On Error GoTo ErrHandler
    Dim Found As Boolean, EdAward As EducationCourseAward, n As Long
    
    Set EdAward = New EducationCourseAward
    
    For n = 1 To Me.Count
        If (Me.Item(n).EducationAwardID = EdAwardID) Then
            Set EdAward = Me.Item(n)
            Found = True
            Exit For
        End If
    Next n
    
    If Found Then
        Set GetByEducationAwardID = EdAward
    Else
        Set EdAward = Nothing
    End If
    
    Exit Function
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
End Function

Public Function GetByEducationAwardName(ByVal AwardName As String) As EducationCourseAward
    On Error GoTo ErrHandler
    Dim Found As Boolean, EdAward As EducationCourseAward, n As Long
    
    Set EdAward = New EducationCourseAward
    
    For n = 1 To Me.Count
        If (Me.Item(n).AwardName = AwardName) Then
            Set EdAward = Me.Item(n)
            Found = True
            Exit For
        End If
    Next n
    
    If Found Then
        Set GetByEducationAwardName = EdAward
    Else
        Set EdAward = Nothing
    End If
    
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
End Function
