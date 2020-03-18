VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EducationCourseAward"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private lngEducationAwardID As Long
Private lngEducationCourseID As Long
Private objEducationCourse As New EducationCourse
Private strAwardName As String
Private strAwardCode As String

Public Property Let EducationAwardID(ByVal nVal As Long)
    lngEducationAwardID = nVal
End Property

Public Property Get EducationAwardID() As Long
    EducationAwardID = lngEducationAwardID
End Property

Public Property Let EducationCourseID(ByVal nVal As Long)
    lngEducationCourseID = nVal
End Property

Public Property Get EducationCourseID() As Long
    EducationCourseID = lngEducationCourseID
End Property

Public Property Set EducationCourseOBJ(ByVal nVal As EducationCourse)
    Set objEducationCourse = nVal
End Property

Public Property Get EducationCourseOBJ() As EducationCourse
    Set EducationCourseOBJ = objEducationCourse
End Property

Public Property Let AwardName(ByVal nVal As String)
    strAwardName = nVal
End Property

Public Property Get AwardName() As String
    AwardName = strAwardName
End Property

Public Property Let AwardCode(ByVal nVal As String)
    strAwardCode = nVal
End Property

Public Property Get AwardCode() As String
    AwardCode = strAwardCode
End Property

Public Function ModifyEducationAward(ByVal ModifyID As Integer) As Boolean
    On Error GoTo ErrHandler
    
    Dim CMD As ADODB.Command
    
    Set CMD = New ADODB.Command
    
    CMD.ActiveConnection = con
    CMD.CommandText = "spModifyEducationAward"
    CMD.CommandType = adCmdStoredProc
    
    CMD.Parameters.Append CMD.CreateParameter(, adInteger, adParamInput, , ModifyID)
    If (Me.EducationAwardID = 0) Then
        CMD.Parameters.Append CMD.CreateParameter(, adBigInt, adParamInput, , 0)
    Else
        CMD.Parameters.Append CMD.CreateParameter(, adBigInt, adParamInput, , Me.EducationAwardID)
    End If
    CMD.Parameters.Append CMD.CreateParameter(, adVarChar, adParamInput, 30, Me.AwardCode)
    CMD.Parameters.Append CMD.CreateParameter(, adVarChar, adParamInput, 50, Me.AwardName)
    CMD.Parameters.Append CMD.CreateParameter(, adVarChar, adParamInput, 50, Me.EducationCourseID)
    
    CMD.Execute

    ModifyEducationAward = True
    
    Exit Function
ErrHandler:
    MsgBox "An error has occured:" & vbNewLine & Err.Description, vbExclamation, "PDR Error"
End Function

