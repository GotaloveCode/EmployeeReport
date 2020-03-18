VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Banks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'====Collection Class: BANKS  for Objects Bank===

'Private Collection Variable
Private mCol As Collection


Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    Set mCol = New Collection
    
    Exit Sub
ErrorHandler:
    MsgBox "An error has occurred" & vbNewLine & err.Description, vbExclamation, APP_TITLE
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    
    Set mCol = Nothing

    'Destroy the Parent Collection Objects that were used by this Collection

    Exit Sub
ErrorHandler:
    MsgBox "An error has occurred" & vbNewLine & err.Description, vbExclamation, APP_TITLE
End Sub

Public Sub Add(ByVal NewBank As Bank)
    On Error GoTo ErrorHandler
    mCol.Add NewBank
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while adding a new Bank" & vbNewLine & _
    err.Description, vbExclamation, APP_TITLE
End Sub

'Function To get Bank object at a given index
Public Function Item(varIndex As Variant) As Bank
    On Error GoTo ErrorHandler
    Set Item = mCol.Item(varIndex)
    
    Exit Function
ErrorHandler:
    MsgBox "An error occurred while Retrieving Bank Information" & vbNewLine & _
        err.Description, vbExclamation, APP_TITLE
    Set Item = Nothing
End Function

'Function To get the Count of  Objects in the Collection
Public Function Count() As Long
    On Error GoTo ErrorHandler
    Count = mCol.Count
    
    Exit Function
ErrorHandler:
    MsgBox "An error has occurred while retrieving the Count of Bank objects" & vbNewLine & _
        err.Description, vbExclamation, APP_TITLE
    Count = 0
End Function

'Method To Remove An Item
Public Sub Remove(varIndex As Variant)
    On Error GoTo ErrorHandler
    mCol.Remove varIndex
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred While Trying To Remove Bank From the Collection" & vbNewLine & _
        err.Description, vbExclamation, APP_TITLE
End Sub

Public Sub Clear()
    On Error GoTo ErrorHandler
    Do Until mCol.Count <= 0
        mCol.Remove mCol.Count
    Loop
    
    Exit Sub
ErrorHandler:
    MsgBox "An error occurred while Clearing Items From the collection" & vbNewLine & _
    err.Description, vbExclamation, APP_TITLE
End Sub

'Method To GetAllBanks Including Deleted i.e. Deleted=1
Public Sub GetAllBanks()
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim objBank As Bank
    
    On Error GoTo ErrorHandler
    
    'First populate Required Collections
    
    'Clear Collection
    Me.Clear
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "spSelectAllBanks"
    Set rs = cmd.Execute()
    
    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveFirst
        Do Until rs.EOF
            Set objBank = New Bank
            With objBank
                If Not IsNull(rs!BankID) Then .BankID = rs!BankID
                If Not IsNull(rs!BankCode) Then .BankCode = rs!BankCode
                If Not IsNull(rs!BankName) Then .BankName = rs!BankName
                If Not IsNull(rs!Comments) Then .Comments = rs!Comments
                If Not IsNull(rs!Deleted) Then .Deleted = rs!Deleted
            End With
            Me.Add objBank
            rs.MoveNext
        Loop
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An Error has occurred while Populating Bank Objects" & vbNewLine & _
        err.Description, vbExclamation, APP_TITLE
End Sub

'Method To GetActiveBanks Excluding Deleted i.e. Deleted=0
Public Sub GetActiveBanks()
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim objBank As Bank
    
    On Error GoTo ErrorHandler
    
    'First populate Required Collections
    
    'Clear Collection
    Me.Clear
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "spSelectActiveBanks"
    Set rs = cmd.Execute()
    
    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveFirst
        Do Until rs.EOF
            Set objBank = New Bank
            With objBank
                If Not IsNull(rs!BankID) Then .BankID = rs!BankID
                If Not IsNull(rs!BankCode) Then .BankCode = rs!BankCode
                If Not IsNull(rs!BankName) Then .BankName = rs!BankName
                If Not IsNull(rs!Comments) Then .Comments = rs!Comments
                If Not IsNull(rs!Deleted) Then .Deleted = rs!Deleted
            End With
            Me.Add objBank
            rs.MoveNext
        Loop
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An Error has occurred while Populating Bank Objects" & vbNewLine & _
        err.Description, vbExclamation, APP_TITLE
End Sub

Public Function FindBankByID(ByVal TheBankID As Long) As Bank
    Dim Found As Boolean
    Dim i As Long
    Dim TheBank As Bank
    
    On Error GoTo ErrorHandler
    For i = 1 To Me.Count
        If Me.Item(i).BankID = TheBankID Then
            Set TheBank = Me.Item(i)
            Found = True
        End If
    Next i
    
    If Found Then
        Set FindBankByID = TheBank
    Else
        Set FindBankByID = Nothing
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred while Searching for Bank" & vbNewLine & _
        err.Description, vbExclamation, APP_TITLE
    Set FindBankByID = Nothing
End Function






'-------------- End Of Collection Definition for: BANKS ------








