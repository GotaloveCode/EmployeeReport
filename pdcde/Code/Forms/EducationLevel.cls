VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EducationLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private lngEducationID As Long
Private streducationCode As String
Private streducationName As String

Public Property Let EducationID(ByVal vNew As Long)
    lngEducationID = vNew
End Property

Public Property Get EducatioID() As Long
    EducationID = lngEducationID
End Property

Public Property Let EducationCode(ByVal vNew As String)
    streducationCode = vNew
End Property

Public Property Get EducationCode() As String
    EducationCode = streducationCode
End Property

Public Property Let educationName(ByVal vNew As String)
    streducationName = vNew
End Property

Public Property Get educationName() As String
    educationName = streducationName
End Property

Public Function ModifyEducation(ByVal ModifyID As Integer) As Boolean
    On Error GoTo ErrHandler
    
    Dim CMD As ADODB.Command
    
    Set CMD = New ADODB.Command
    
    CMD.ActiveConnection = con
    CMD.CommandText = "spModifyEducation" 'implement stored procedure
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

