VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Prompts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim mCol As Collection

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

Public Sub Add(ByVal ThePrompt As Prompt)
    mCol.Add ThePrompt
End Sub

Public Sub Remove(ByVal Index As Long)
    mCol.Remove Index
End Sub

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Property Get Item(ByVal Index As Long) As Prompt
    On Error Resume Next
    Set Item = mCol.Item(Index)
End Property

Public Sub Clear()
    Do While Me.Count > 0
        Me.Remove Me.Count
    Loop
End Sub

Public Sub GetAllPrompts()
    On Error GoTo ErrorHandler
        
    Dim CMD As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim MyPrompt As Prompt
    
    Set CMD = New ADODB.Command
    Set rs = New ADODB.Recordset
    
    CMD.ActiveConnection = con
    CMD.CommandType = adCmdStoredProc
    CMD.CommandText = "pdrspSelectAllPrompts"
    
    Set rs = CMD.Execute()
    
    If Not (rs.EOF Or rs.BOF) Then
        rs.MoveFirst
        Do Until rs.EOF
            Set MyPrompt = New Prompt
            
            MyPrompt.Probation = rs!Probation
            MyPrompt.Contract = rs!Contracted
            MyPrompt.Retirement = rs!Retirement
            MyPrompt.Termination = rs!Termination
            MyPrompt.MaleRetirementAge = rs!MaleRetirementAge
            MyPrompt.FemaleRetirementAge = rs!FemaleRetirementAge
            MyPrompt.ProjectEnd = IIf(IsNull(rs!ProjectEnd), 0, rs!ProjectEnd)
            If CBool(rs!EnablePrompts) Then
                MyPrompt.EnablePrompts = 1
            Else
                MyPrompt.EnablePrompts = 0
            End If
            
            Me.Add MyPrompt
            rs.MoveNext
        Loop
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox ERR_INTRO & err.Description, vbExclamation, APP_TITLE
    Exit Sub
End Sub


