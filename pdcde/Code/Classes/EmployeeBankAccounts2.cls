VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmployeeBankAccounts2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'====Collection Class: EMPLOYEEBANKACCOUNTS  for Objects EmployeeBankAccount===

'Private Collection Variable
Private mCol As Collection
''Private colEmployees As Employees
Private colBankBranches As BankBranches


Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    Set mCol = New Collection
    Set colEmployees = New Employees
    Set colBankBranches = New BankBranches
    
    Exit Sub
ErrorHandler:
    MsgBox "An error has occurred" & vbNewLine & err.Description, vbExclamation, APP_TITLE
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    
    Set mCol = Nothing

    'Destroy the Parent Collection Objects that were used by this Collection
    Set colEmployees = Nothing
    Set colBankBranches = Nothing

    Exit Sub
ErrorHandler:
    MsgBox "An error has occurred" & vbNewLine & err.Description, vbExclamation, APP_TITLE
End Sub

Public Sub add(ByVal NewEmployeeBankAccount As EmployeeBankAccount2)
    On Error GoTo ErrorHandler
    mCol.add NewEmployeeBankAccount
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while adding a new EmployeeBankAccount" & vbNewLine & _
    err.Description, vbExclamation, APP_TITLE
End Sub

'Function To get EmployeeBankAccount object at a given index
Public Function Item(varIndex As Variant) As EmployeeBankAccount2
    On Error GoTo ErrorHandler
    Set Item = mCol.Item(varIndex)
    
    Exit Function
ErrorHandler:
    MsgBox "An error occurred while Retrieving EmployeeBankAccount Information" & vbNewLine & _
        err.Description, vbExclamation, APP_TITLE
    Set Item = Nothing
End Function

'Function To get the Count of  Objects in the Collection
Public Function count() As Long
    On Error GoTo ErrorHandler
    count = mCol.count
    
    Exit Function
ErrorHandler:
    MsgBox "An error has occurred while retrieving the Count of EmployeeBankAccount objects" & vbNewLine & _
        err.Description, vbExclamation, APP_TITLE
    count = 0
End Function

'Method To Remove An Item
Public Sub remove(varIndex As Variant)
    On Error GoTo ErrorHandler
    mCol.remove varIndex
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred While Trying To Remove EmployeeBankAccount From the Collection" & vbNewLine & _
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

'Method To GetAllEmployeeBankAccounts Including Deleted i.e. Deleted=1
Public Sub GetAllEmployeeBankAccounts()
    Dim CMD As adodb.Command
    Dim rs As adodb.Recordset
    Dim objEmployeeBankAccount As EmployeeBankAccount2
    
    On Error GoTo ErrorHandler
    
    'First populate Required Collections
'    colEmployees.GetAllEmployees
'    colBankBranches.GetAllBankBranches
    
    'Clear Collection
    Me.Clear
    
    Set CMD = New adodb.Command
    CMD.ActiveConnection = con
    CMD.CommandType = adCmdStoredProc
    CMD.CommandText = "spSelectAllEmployeeBankAccounts"
    Set rs = CMD.Execute()
    
    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveFirst
        Do Until rs.EOF
            Set objEmployeeBankAccount = New EmployeeBankAccount2
            With objEmployeeBankAccount
                If Not IsNull(rs!EmployeeBankAccountID) Then .EmployeeBankAccountID = rs!EmployeeBankAccountID
                If Not IsNull(rs!Employeeid) Then
                    Set .Employee = colEmployees.FindEmployeeByID(rs!Employeeid)
                Else
                    Set .Employee = Nothing
                End If
                .Employeeid = rs!Employeeid
                If Not IsNull(rs!BankBranchID) Then
                    Set .bankbranch = colBankBranches.FindBankBranchByID(rs!BankBranchID)
                Else
                    Set .bankbranch = Nothing
                End If
                If Not IsNull(rs!AccountName) Then .AccountName = rs!AccountName
                If Not IsNull(rs!AccountNumber) Then .AccountNumber = rs!AccountNumber
                If Not IsNull(rs!AccountType) Then .AccountType = rs!AccountType
                If Not IsNull(rs!IsMainAccount) Then .IsMainAccount = rs!IsMainAccount
                If Not IsNull(rs!SwiftCode) Then .SwiftCode = rs!SwiftCode
                If Not IsNull(rs!Deleted) Then .Deleted = rs!Deleted
            End With
            Me.add objEmployeeBankAccount
            rs.MoveNext
        Loop
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An Error has occurred while Populating EmployeeBankAccount Objects" & vbNewLine & _
        err.Description, vbExclamation, APP_TITLE
End Sub

'Method To GetActiveEmployeeBankAccounts Excluding Deleted i.e. Deleted=0
Public Sub GetActiveEmployeeBankAccounts()
    Dim CMD As adodb.Command
    Dim rs As adodb.Recordset
    Dim objEmployeeBankAccount As EmployeeBankAccount2
    
    On Error GoTo ErrorHandler
    
    'First populate Required Collections
    ''colEmployees.GetAllEmployees
    Set colEmployees = AllEmployees
    
   
    ''colBankBranches.getAllBankBranches
    Set colBankBranches = pBankBranches
    
    'Clear Collection
    Me.Clear
    
    Set CMD = New adodb.Command
    CMD.ActiveConnection = con
    CMD.CommandType = adCmdStoredProc
    CMD.CommandText = "spSelectActiveEmployeeBankAccounts"
    Set rs = CMD.Execute()
    
    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveFirst
        Do Until rs.EOF
            Set objEmployeeBankAccount = New EmployeeBankAccount2
            With objEmployeeBankAccount
            
                If Not IsNull(rs!EmployeeBankAccountID) Then .EmployeeBankAccountID = rs!EmployeeBankAccountID
                If Not IsNull(rs!Employeeid) Then
                
                
                      ''Set .Employee = colEmployees.FindEmployeeByID(rs!EmployeeID)
                       Set .Employee = AllEmployees.FindEmployeeByID(rs!Employeeid)
                       
                      .Employeeid = rs!Employeeid
'                      Dim xxx As Long
'                      Dim x3 As Long
'                      For xxx = 1 To AllEmployees.count
'                      x3 = AllEmployees.Item(xxx).EmployeeID
'
'                      Next xxx
                      
                      
'                  Dim PP As String
'                  PP = .Employee.OtherNames
'                      If AllEmployees Is Nothing Then
'                      Set AllEmployees = New HRCORE.Employees
'                      AllEmployees.GetAccessibleEmployeesByUser currUser.UserID
'                      End If
'
'                      If Not AllEmployees Is Nothing Then
'                      Set .Employee = AllEmployees.FindEmployeeByID(rs!EmployeeID)
'                      End If
                Else
                    Set .Employee = Nothing
                End If
                
               
                If Not IsNull(rs!BankBranchID) Then
                
                If colBankBranches Is Nothing Then
                If Not pBankBranches Is Nothing Then
                Set colBankBranches = pBankBranches
                Else
                Set colBankBranches = New BankBranches
                colBankBranches.GetAllBankBranches
                End If
                End If
                    Set .bankbranch = colBankBranches.FindBankBranchByID(rs!BankBranchID)
                Else
                    Set .bankbranch = Nothing
                End If
                If Not IsNull(rs!AccountName) Then .AccountName = rs!AccountName
                If Not IsNull(rs!AccountNumber) Then .AccountNumber = rs!AccountNumber
                If Not IsNull(rs!AccountType) Then .AccountType = rs!AccountType
                If Not IsNull(rs!SwiftCode) Then .SwiftCode = rs!SwiftCode
                If Not IsNull(rs!IsMainAccount) Then .IsMainAccount = rs!IsMainAccount
                If Not IsNull(rs!Deleted) Then .Deleted = rs!Deleted
            End With
            Me.add objEmployeeBankAccount
            rs.MoveNext
        Loop
    End If
   
    Exit Sub
    
ErrorHandler:
    MsgBox "An Error has occurred while Populating EmployeeBankAccount Objects" & vbNewLine & _
        err.Description, vbExclamation, APP_TITLE
End Sub





Public Function FindEmployeeBankAccountByID(ByVal TheEmployeeBankAccountID As Long) As EmployeeBankAccount2
    Dim Found As Boolean
    Dim i As Long
    Dim TheEmployeeBankAccount As EmployeeBankAccount2
    
    On Error GoTo ErrorHandler
    For i = 1 To Me.count
        If Me.Item(i).EmployeeBankAccountID = TheEmployeeBankAccountID Then
            Set TheEmployeeBankAccount = Me.Item(i)
            Found = True
        End If
    Next i
    
    If Found Then
        Set FindEmployeeBankAccountByID = TheEmployeeBankAccount
    Else
        Set FindEmployeeBankAccountByID = Nothing
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred while Searching for EmployeeBankAccount" & vbNewLine & _
        err.Description, vbExclamation, APP_TITLE
    Set FindEmployeeBankAccountByID = Nothing
End Function

Public Function GetEmployeeBankAccountsOfEmployeeID(ByVal TheEmployeeID As Long) As EmployeeBankAccounts2
    Dim i As Long
    Dim colAccounts As EmployeeBankAccounts2
    
    On Error GoTo ErrorHandler
    
    Set colAccounts = New EmployeeBankAccounts2
    For i = 1 To Me.count
        If Not (Me.Item(i).Employee Is Nothing) Then
           '' If Me.Item(i).Employee.Employeeid = TheEmployeeID Then
            If Me.Item(i).Employeeid = TheEmployeeID Then
                colAccounts.add Me.Item(i)
            End If
        End If
    Next i
    
    Set GetEmployeeBankAccountsOfEmployeeID = colAccounts
    
    Exit Function
    
ErrorHandler:
    Set GetEmployeeBankAccountsOfEmployeeID = Nothing
            
End Function


Public Function GetEmployeeBankAccountsOfBankID(ByVal TheBankID As Long) As EmployeeBankAccounts2
    Dim i As Long
    Dim colAccounts As EmployeeBankAccounts2
    
    On Error GoTo ErrorHandler
    
    For i = 1 To Me.count
        If Not (Me.Item(i).bankbranch Is Nothing) Then
            If Not (Me.Item(i).bankbranch.Bank Is Nothing) Then
                If Me.Item(i).bankbranch.Bank.BankID = TheBankID Then
                    colAccounts.add Me.Item(i)
                End If
            End If
        End If
    Next i
    
    Set GetEmployeeBankAccountsOfBankID = colAccounts
    
    Exit Function
    
ErrorHandler:
    Set GetEmployeeBankAccountsOfBankID = Nothing
End Function


Public Function GetEmployeeBankAccountsOfBankBranchID(ByVal TheBankBranchID As Long) As EmployeeBankAccounts2
    Dim i As Long
    Dim colAccounts As EmployeeBankAccounts2
    
    On Error GoTo ErrorHandler
    
    For i = 1 To Me.count
        If Not (Me.Item(i).bankbranch Is Nothing) Then
            If Me.Item(i).bankbranch.BankBranchID = TheBankBranchID Then
                colAccounts.add Me.Item(i)
            End If
        End If
    Next i
    
    Set GetEmployeeBankAccountsOfBankBranchID = colAccounts
    
    Exit Function
    
ErrorHandler:
    Set GetEmployeeBankAccountsOfBankBranchID = Nothing
End Function

Public Sub RemoveByID(ByVal TheEmployeeBankAccountID As Long)
    Dim i As Long
    
    On Error GoTo ErrorHandler
    For i = 1 To Me.count
        If Me.Item(i).EmployeeBankAccountID = TheEmployeeBankAccountID Then
            Me.remove i
            GoTo Finish
        End If
    Next i
Finish:
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while remove a specific employee bank account" & vbNewLine & _
        err.Description, vbExclamation, APP_TITLE
End Sub
'-------------- End Of Collection Definition for: EmployeeBankAccounts2 ------







