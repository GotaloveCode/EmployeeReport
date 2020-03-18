VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Prompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private intProbation As Integer
Private intRetirement As Integer
Private intContract As Integer
Private intTermination
Private intMaleRetirementAge As Integer
Private intFemaleRetirementAge As Integer
Private intEnablePrompts As Integer
Private intProjectEnd As Integer

Public Property Let ProjectEnd(ByVal newVal As Integer)
    intProjectEnd = newVal
End Property

Public Property Get ProjectEnd() As Integer
    ProjectEnd = intProjectEnd
End Property

Public Property Let EnablePrompts(ByVal newVal As Integer)
    intEnablePrompts = newVal
End Property

Public Property Get EnablePrompts() As Integer
    EnablePrompts = intEnablePrompts
End Property

Public Property Let Termination(ByVal newVal As Integer)
    intTermination = newVal
End Property

Public Property Get Termination() As Integer
    Termination = intTermination
End Property

Public Property Let Probation(ByVal newVal As Integer)
    intProbation = newVal
End Property

Public Property Get Probation() As Integer
    Probation = intProbation
End Property

Public Property Let Retirement(ByVal newVal As Integer)
    intRetirement = newVal
End Property

Public Property Get Retirement() As Integer
    Retirement = intRetirement
End Property

Public Property Let Contract(ByVal newVal As Integer)
    intContract = newVal
End Property

Public Property Get Contract() As Integer
    Contract = intContract
End Property

Public Property Let MaleRetirementAge(ByVal newVal As Integer)
    intMaleRetirementAge = newVal
End Property

Public Property Get MaleRetirementAge() As Integer
    MaleRetirementAge = intMaleRetirementAge
End Property

Public Property Let FemaleRetirementAge(ByVal newVal As Integer)
    intFemaleRetirementAge = newVal
End Property

Public Property Get FemaleRetirementAge() As Integer
    FemaleRetirementAge = intFemaleRetirementAge
End Property

Public Function Update() As Boolean
    On Error GoTo ErrorHandler
           
     Dim CMD As New ADODB.Command
     
    CMD.ActiveConnection = con
    CMD.CommandType = adCmdStoredProc
    CMD.CommandText = "pdrspUpdatePromptSetup"
    
    CMD.Parameters.Append CMD.CreateParameter(, adInteger, adParamInput, , Me.Probation)
    CMD.Parameters.Append CMD.CreateParameter(, adInteger, adParamInput, , Me.Retirement)
    CMD.Parameters.Append CMD.CreateParameter(, adInteger, adParamInput, , Me.Contract)
    CMD.Parameters.Append CMD.CreateParameter(, adInteger, adParamInput, , Me.Termination)
    CMD.Parameters.Append CMD.CreateParameter(, adInteger, adParamInput, , Me.MaleRetirementAge)
    CMD.Parameters.Append CMD.CreateParameter(, adInteger, adParamInput, , Me.FemaleRetirementAge)
    CMD.Parameters.Append CMD.CreateParameter(, adInteger, adParamInput, , Me.ProjectEnd)
    CMD.Parameters.Append CMD.CreateParameter(, adInteger, adParamInput, , Me.EnablePrompts)
    CMD.Execute
    
    Update = True
    
    Exit Function
ErrorHandler:
    MsgBox "An error has occured:" & vbCrLf & Err.Description, vbExclamation, "PDR: Error"
End Function

