VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Education"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private lngEducationID As Long
Private streducationCode As String
Private streducationName As String

Public Property Let Educationid(ByVal vnew As Long)
    lngEducationID = vnew
End Property

Public Property Get Educationid() As Long
    Educationid = lngEducationID
End Property

Public Property Let EducationCode(ByVal vnew As String)
    streducationCode = vnew
End Property

Public Property Get EducationCode() As String
    EducationCode = streducationCode
End Property

Public Property Let educationName(ByVal vnew As String)
    streducationName = vnew
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
    If (Me.Educationid = 0) Then
        CMD.Parameters.Append CMD.CreateParameter(, adBigInt, adParamInput, , 0)
    Else
        CMD.Parameters.Append CMD.CreateParameter(, adBigInt, adParamInput, , Me.Educationid)
    End If
    CMD.Parameters.Append CMD.CreateParameter(, adVarChar, adParamInput, 30, Me.EducationCode)
    CMD.Parameters.Append CMD.CreateParameter(, adVarChar, adParamInput, 50, Me.educationName)
    
    CMD.Execute
        
    ModifyEducation = True
    
    Exit Function
ErrHandler:
    MsgBox "An error has occured:" & vbNewLine & err.Description, vbExclamation, "PDR Error"
End Function

