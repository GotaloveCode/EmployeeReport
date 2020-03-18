VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EducationCourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private lngEducationCourseID As Long
Private strCourseCode As String
Private strCourseName As String

Public Property Let EducationCourseID(ByVal vNew As Long)
    lngEducationCourseID = vNew
End Property

Public Property Get EducationCourseID() As Long
    EducationCourseID = lngEducationCourseID
End Property

Public Property Let CourseCode(ByVal vNew As String)
    strCourseCode = vNew
End Property

Public Property Get CourseCode() As String
    CourseCode = strCourseCode
End Property

Public Property Let CourseName(ByVal vNew As String)
    strCourseName = vNew
End Property

Public Property Get CourseName() As String
    CourseName = strCourseName
End Property

Public Function ModifyEducationCourse(ByVal ModifyID As Integer) As Boolean
    On Error GoTo ErrHandler
    
    Dim CMD As ADODB.Command
    
    Set CMD = New ADODB.Command
    
    CMD.ActiveConnection = con
    CMD.CommandText = "spModifyEducationCourse"
    CMD.CommandType = adCmdStoredProc
    
    CMD.Parameters.Append CMD.CreateParameter(, adInteger, adParamInput, , ModifyID)
    If (Me.EducationCourseID = 0) Then
        CMD.Parameters.Append CMD.CreateParameter(, adBigInt, adParamInput, , 0)
    Else
        CMD.Parameters.Append CMD.CreateParameter(, adBigInt, adParamInput, , Me.EducationCourseID)
    End If
    CMD.Parameters.Append CMD.CreateParameter(, adVarChar, adParamInput, 30, Me.CourseCode)
    CMD.Parameters.Append CMD.CreateParameter(, adVarChar, adParamInput, 50, Me.CourseName)
    
    CMD.Execute
        
    ModifyEducationCourse = True
    
    Exit Function
ErrHandler:
    MsgBox "An error has occured:" & vbNewLine & Err.Description, vbExclamation, "PDR Error"
End Function
