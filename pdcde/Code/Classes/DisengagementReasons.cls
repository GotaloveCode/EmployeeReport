VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DisengagementReasons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ColDisReasons

Private Sub Class_Initialize()
    Set ColDisReasons = New Collection
End Sub

Private Sub Class_Terminate()
    Set ColDisReasons = Nothing
End Sub

Public Sub Add(ByVal vnew As disengagementReason)
    ColDisReasons.Add vnew
End Sub

Public Function Item(ByVal vnew As Variant) As disengagementReason
   Set Item = ColDisReasons.Item(vnew)
End Function

Public Sub Remove(ByVal vnew As Variant)
    ColDisReasons.Remove vnew
End Sub

Public Function Count() As Long
    Count = ColDisReasons.Count
End Function

Public Sub Clear()
    Do Until ColDisReasons.Count <= 0
        ColDisReasons.Remove ColDisReasons.Count
    Loop
End Sub

Public Function FindReason(ByVal vnew As Long) As disengagementReason
    Dim i As Long
    Dim found As Boolean
    found = False
    
    For i = 1 To Me.Count
        If Me.Item(i).ReasonID = vnew Then
            Set FindReason = Me.Item(i)
            found = True
            Exit For
        End If
    Next i
    
    If Not found Then Set FindReason = Nothing
    
End Function

Public Function FindTheReason(ByVal vnew As String) As Boolean
    Dim i As Long
    Dim found As Boolean
    found = False
    
    For i = 1 To Me.Count
        If UCase(Me.Item(i).Reason) = UCase(vnew) Then
            found = True
            Exit For
        End If
    Next i
    
    FindTheReason = found
    
End Function

Public Sub GetallDisengagementReasons()
On Error GoTo ErrHandler

    Dim CMD As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim DisReason As disengagementReason
    Set CMD = New ADODB.Command
    CMD.CommandText = "spGetDisengagementReasons"
    CMD.CommandType = adCmdStoredProc
    CMD.ActiveConnection = con
    Set rs = CMD.Execute
    Me.Clear
    
    If Not (rs Is Nothing) Then
        If Not (rs.BOF And rs.EOF) Then
            rs.MoveFirst
            Do Until rs.EOF
                Set DisReason = New disengagementReason
                If Not IsNull(rs!Code) Then DisReason.Code = rs!Code
                If Not IsNull(rs!Reason) Then DisReason.Reason = rs!Reason
                If Not IsNull(rs!Deleted) Then DisReason.Deleted = rs!Deleted
                DisReason.ReasonID = rs!ReasonID
                Me.Add DisReason
                
                rs.MoveNext
            Loop
        End If
     End If
     
     Exit Sub
ErrHandler:
     MsgBox "An error has occur " & err.Description
End Sub

