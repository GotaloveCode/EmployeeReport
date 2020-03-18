VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "disengagementReason"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private strReason As String
Private strCode As String
Private lngReasonID As Long
Private lnDeleted As Integer

Public Property Get Reason() As String
    Reason = strReason
End Property



Public Property Let Reason(ByVal data As String)
    strReason = data
End Property

Public Property Get Deleted() As Integer
Deleted = lnDeleted
End Property
Public Property Let Deleted(ByVal vnew As Integer)
lnDeleted = vnew
End Property

Public Property Get Code() As String
    Code = strCode
End Property

Public Property Let Code(ByVal data As String)
    strCode = data
End Property

Public Property Get ReasonID() As Long
    ReasonID = lngReasonID
End Property

Public Property Let ReasonID(ByVal data As Long)
    lngReasonID = data
End Property

Public Function Insert() As Boolean
    On Error GoTo ErrHandler
    Dim i As HRCORE.CompanyDetails
    
    Dim CMD As ADODB.Command
    Set CMD = New ADODB.Command
    CMD.CommandText = "spinsertDisengagementReasons"
    CMD.CommandType = adCmdStoredProc
    CMD.ActiveConnection = con
    Insert = False
    CMD.Parameters.Append CMD.CreateParameter(, adVarChar, adParamInput, 20, Me.Code)
    CMD.Parameters.Append CMD.CreateParameter(, adVarChar, adParamInput, 500, Me.Reason)
    CMD.Parameters.Append CMD.CreateParameter(, adInteger, adParamOutput)
    CMD.Execute
    
    Insert = True
    
    Exit Function
ErrHandler:
    Insert = False
End Function


Public Function Update() As Boolean
    On Error GoTo ErrHandler
    Dim CMD As ADODB.Command
    Set CMD = New ADODB.Command
    CMD.CommandText = "spUpdateDisengagementReasons"
    CMD.CommandType = adCmdStoredProc
    CMD.ActiveConnection = con
    Update = False
    CMD.Parameters.Append CMD.CreateParameter(, adVarChar, adParamInput, 20, Me.Code)
    CMD.Parameters.Append CMD.CreateParameter(, adVarChar, adParamInput, 500, Me.Reason)
    CMD.Parameters.Append CMD.CreateParameter(, adInteger, adParamInput, , Me.ReasonID)
    CMD.Execute
    
    Update = True
    
    Exit Function
ErrHandler:
    Update = False
End Function

