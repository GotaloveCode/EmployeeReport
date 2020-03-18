VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BankBranches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'====Collection Class: BANKBRANCHES  for Objects BankBranch===

'Private Collection Variable
Private mCol As Collection
Private colBanks As Banks


Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    Set mCol = New Collection
    Set colBanks = New Banks
    
    Exit Sub
ErrorHandler:
    MsgBox "An error has occurred" & vbNewLine & err.Description, vbExclamation, APP_TITLE
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    
    Set mCol = Nothing

    'Destroy the Parent Collection Objects that were used by this Collection
    Set colBanks = Nothing

    Exit Sub
ErrorHandler:
    MsgBox "An error has occurred" & vbNewLine & err.Description, vbExclamation, APP_TITLE
End Sub

Public Sub add(ByVal NewBankBranch As BankBranch)
    On Error GoTo ErrorHandler
    mCol.add NewBankBranch
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while adding a new BankBranch" & vbNewLine & _
    err.Description, vbExclamation, APP_TITLE
End Sub

'Function To get BankBranch object at a given index
Public Function Item(varIndex As Variant) As BankBranch
    On Error GoTo ErrorHandler
    Set Item = mCol.Item(varIndex)
    
    Exit Function
ErrorHandler:
    MsgBox "An error occurred while Retrieving BankBranch Information" & vbNewLine & _
        err.Description, vbExclamation, APP_TITLE
    Set Item = Nothing
End Function

'Function To get the Count of  Objects in the Collection
Public Function count() As Long
    On Error GoTo ErrorHandler
    count = mCol.count
    
    Exit Function
ErrorHandler:
    MsgBox "An error has occurred while retrieving the Count of BankBranch objects" & vbNewLine & _
        err.Description, vbExclamation, APP_TITLE
    count = 0
End Function

'Method To Remove An Item
Public Sub remove(varIndex As Variant)
    On Error GoTo ErrorHandler
    mCol.remove varIndex
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred While Trying To Remove BankBranch From the Collection" & vbNewLine & _
        err.Description, vbExclamation, APP_TITLE
End Sub

Public Sub Clear()
    On Error GoTo ErrorHandler
    Do Until mCol.count <= 0
        mCol.remove mCol.count
    Loop
    
    Exit Sub
ErrorHandler:
    MsgBox "An error occurred while Clearing Items From the collection" & vbNewLine & _
    err.Description, vbExclamation, APP_TITLE
End Sub

'Method To GetAllBankBranches Including Deleted i.e. Deleted=1
Public Sub GetAllBankBranches()
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim objBankBranch As BankBranch
    
    On Error GoTo ErrorHandler
    
    'First populate Required Collections
    colBanks.GetAllBanks
    
    'Clear Collection
    Me.Clear
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "spSelectAllBankBranches"
    Set rs = cmd.Execute()
    
    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveFirst
        Do Until rs.EOF
            Set objBankBranch = New BankBranch
            With objBankBranch
                If Not IsNull(rs!BankBranchID) Then .BankBranchID = rs!BankBranchID
                If Not IsNull(rs!BankID) Then
                    Set .Bank = colBanks.FindBankByID(rs!BankID)
                Else
                    Set .Bank = Nothing
                End If
                If Not IsNull(rs!BranchCode) Then .BranchCode = rs!BranchCode
                If Not IsNull(rs!BranchName) Then .BranchName = rs!BranchName
                If Not IsNull(rs!Comments) Then .Comments = rs!Comments
                If Not IsNull(rs!Deleted) Then .Deleted = rs!Deleted
            End With
            Me.add objBankBranch
            rs.MoveNext
        Loop
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An Error has occurred while Populating BankBranch Objects" & vbNewLine & _
        err.Description, vbExclamation, APP_TITLE
End Sub

'Method To GetActiveBankBranches Excluding Deleted i.e. Deleted=0
Public Sub GetActiveBankBranches()
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim objBankBranch As BankBranch
    
    On Error GoTo ErrorHandler
    
    'First populate Required Collections
    colBanks.GetAllBanks
    
    'Clear Collection
    Me.Clear
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "spSelectActiveBankBranches"
    Set rs = cmd.Execute()
    
    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveFirst
        Do Until rs.EOF
            Set objBankBranch = New BankBranch
            With objBankBranch
                If Not IsNull(rs!BankBranchID) Then .BankBranchID = rs!BankBranchID
                If Not IsNull(rs!BankID) Then
                    Set .Bank = colBanks.FindBankByID(rs!BankID)
                Else
                    Set .Bank = Nothing
                End If
                If Not IsNull(rs!BranchCode) Then .BranchCode = rs!BranchCode
                If Not IsNull(rs!BranchName) Then .BranchName = rs!BranchName
                If Not IsNull(rs!Comments) Then .Comments = rs!Comments
                If Not IsNull(rs!Deleted) Then .Deleted = rs!Deleted
            End With
            Me.add objBankBranch
            rs.MoveNext
        Loop
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An Error has occurred while Populating BankBranch Objects" & vbNewLine & _
        err.Description, vbExclamation, APP_TITLE
End Sub

Public Function FindBankBranchByID(ByVal TheBankBranchID As Long) As BankBranch
    Dim Found As Boolean
    Dim i As Long
    Dim TheBankBranch As BankBranch
    
    On Error GoTo ErrorHandler
    For i = 1 To Me.count
        If Me.Item(i).BankBranchID = TheBankBranchID Then
            Set TheBankBranch = Me.Item(i)
            Found = True
            Exit For
        End If
    Next i
    
    If Found Then
        Set FindBankBranchByID = TheBankBranch
    Else
        Set FindBankBranchByID = Nothing
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred while Searching for BankBranch" & vbNewLine & _
        err.Description, vbExclamation, APP_TITLE
    Set FindBankBranchByID = Nothing
End Function


Public Function GetBranchesOfBankID(ByVal TheBankID As Long) As BankBranches
    Dim i As Long
    Dim colBranches As New BankBranches
    
    On Error GoTo ErrorHandler
    
    For i = 1 To Me.count
        If Not (Me.Item(i).Bank Is Nothing) Then
            If Me.Item(i).Bank.BankID = TheBankID Then
                colBranches.add Me.Item(i)
            End If
        End If
    Next i
    
    Set GetBranchesOfBankID = colBranches
    
    Exit Function
    
ErrorHandler:
    Set GetBranchesOfBankID = Nothing
End Function



'-------------- End Of Collection Definition for: BANKBRANCHES ------








