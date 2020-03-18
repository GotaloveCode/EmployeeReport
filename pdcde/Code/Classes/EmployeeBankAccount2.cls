VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmployeeBankAccount2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========START OF CLASS EmployeeBankAccount =============

'These are Private Variables Handling The Fields in the Classes

Private lngEmployeeBankAccountID As Long 'This is a Primary Key Column
Private objEmployee As HRCORE.Employee
Private objEmployeeid As Long
'Private lngEmployeeID As Long  'this is A Foregn key Column
Private objBankBranch As bankbranch
'Private lngBankBranchID As Long    'this is A Foregn key Column
Private strAccountName As String
Private strAccountNumber As String
Private strAccountType As String
Private blnIsMainAccount As Boolean
Private blnDeleted As Boolean
Private strSwiftCode As String

Private Sub Class_Initialize()
    'Instantiate the Object Variables
    On Error GoTo ErrorHandler
    
    Set objEmployee = New HRCORE.Employee
    Set objBankBranch = New bankbranch
    Exit Sub
ErrorHandler:
    MsgBox "An error has Ocuured" & vbNewLine & err.Description, vbExclamation, APP_TITLE
End Sub

Private Sub Class_Terminate()
    'Destroy the Object Variables
    On Error GoTo ErrorHandler
    
    Set objEmployee = Nothing
    Set objBankBranch = Nothing
    Exit Sub
ErrorHandler:
    MsgBox "An error has occurred" & vbNewLine & err.Description, vbExclamation, APP_TITLE
End Sub

' === Property Procedures For Object EmployeeBankAccount ===
    
Public Property Get EmployeeBankAccountID() As Long
    EmployeeBankAccountID = lngEmployeeBankAccountID
End Property
    
Public Property Let EmployeeBankAccountID(ByVal vNew As Long)
    lngEmployeeBankAccountID = vNew
End Property
    
    'This is an Object Variable: objEmployee
Public Property Get Employee() As HRCORE.Employee
    Set Employee = objEmployee
End Property

Public Property Set Employee(ByVal vNew As HRCORE.Employee)
    Set objEmployee = vNew
End Property

Public Property Let Employeeid(ByVal vNew As Long)
    objEmployeeid = vNew
End Property
Public Property Get Employeeid() As Long
    Employeeid = objEmployeeid
End Property


    
    
    'This is an Object Variable: objBankBranch
Public Property Get bankbranch() As bankbranch
    Set bankbranch = objBankBranch
End Property

Public Property Set bankbranch(ByVal vNew As bankbranch)
    Set objBankBranch = vNew
End Property
    
    
Public Property Get AccountName() As String
    AccountName = strAccountName
End Property
    
Public Property Let AccountName(ByVal vNew As String)
    strAccountName = vNew
End Property
    
Public Property Get AccountNumber() As String
    AccountNumber = strAccountNumber
End Property
    
Public Property Let AccountNumber(ByVal vNew As String)
    strAccountNumber = vNew
End Property

Public Property Get SwiftCode() As String
    SwiftCode = strSwiftCode
End Property
    
Public Property Let SwiftCode(ByVal vNew As String)
    strSwiftCode = vNew
End Property
    
Public Property Get AccountType() As String
    AccountType = strAccountType
End Property
    
Public Property Let AccountType(ByVal vNew As String)
    strAccountType = vNew
End Property
    
Public Property Get IsMainAccount() As Boolean
    IsMainAccount = blnIsMainAccount
End Property
    
Public Property Let IsMainAccount(ByVal vNew As Boolean)
    blnIsMainAccount = vNew
End Property
    
Public Property Get Deleted() As Boolean
    Deleted = blnDeleted
End Property
    
Public Property Let Deleted(ByVal vNew As Boolean)
    blnDeleted = vNew
End Property

'===METHODS FOR Object EmployeeBankAccount

'Method To insert a new EmployeeBankAccount object
Public Function InsertNew() As Long
    Dim CMD As adodb.Command
    
    On Error GoTo ErrorHandler
    Set CMD = New adodb.Command
    CMD.ActiveConnection = con
    CMD.CommandType = adCmdStoredProc
    CMD.CommandText = "spInsertEmployeeBankAccount"
    
    'The Primary Key Column
    CMD.Parameters.Append CMD.CreateParameter("@EmployeeBankAccountID", adInteger, adParamOutput)
    If Not (Employee Is Nothing) Then
        ''CMD.Parameters.Append CMD.CreateParameter("@EmployeeID", adInteger, adParamInput, , Employee.Employeeid)
        CMD.Parameters.Append CMD.CreateParameter("@EmployeeID", adInteger, adParamInput, , Employeeid)
    Else
        'If the FK Is Nothing Then Don't Supply a Value: A NULL Value will be Inserted by The Insert Procedure
        CMD.Parameters.Append CMD.CreateParameter("@EmployeeID", adInteger, adParamInput)
    End If
    If Not (bankbranch Is Nothing) Then
        CMD.Parameters.Append CMD.CreateParameter("@BankBranchID", adInteger, adParamInput, , bankbranch.BankBranchID)
    Else
        'If the FK Is Nothing Then Don't Supply a Value: A NULL Value will be Inserted by The Insert Procedure
        CMD.Parameters.Append CMD.CreateParameter("@BankBranchID", adInteger, adParamInput)
    End If
    CMD.Parameters.Append CMD.CreateParameter("@AccountName", adVarChar, adParamInput, 80, AccountName)
    CMD.Parameters.Append CMD.CreateParameter("@AccountNumber", adVarChar, adParamInput, 50, AccountNumber)
    CMD.Parameters.Append CMD.CreateParameter("@AccountType", adVarChar, adParamInput, 50, AccountType)
    CMD.Parameters.Append CMD.CreateParameter("SwiftCode", adVarChar, adParamInput, 50, SwiftCode)
    CMD.Parameters.Append CMD.CreateParameter("@IsMainAccount", adBoolean, adParamInput, , IsMainAccount)
    CMD.Parameters.Append CMD.CreateParameter("@Deleted", adBoolean, adParamInput, , Deleted)
    CMD.Execute
    lngEmployeeBankAccountID = CMD.Parameters(0).value
    
    InsertNew = 0
    Exit Function
ErrorHandler:
    MsgBox "An error has occurred while inserting a new EmployeeBankAccount" & vbNewLine & _
    err.Description, vbExclamation, APP_TITLE
    InsertNew = err.Number
End Function

'Method to Update the Object EmployeeBankAccount
Public Function Update() As Long
    Dim CMD As adodb.Command
    
    On Error GoTo ErrorHandler
    Set CMD = New adodb.Command
    CMD.ActiveConnection = con
    CMD.CommandType = adCmdStoredProc
    CMD.CommandText = "spUpdateEmployeeBankAccount"
    
    'The Primary Key Column
    CMD.Parameters.Append CMD.CreateParameter("@EmployeeBankAccountID", adInteger, adParamInput, , EmployeeBankAccountID)
        If Not (Employee Is Nothing) Then
       '' CMD.Parameters.Append CMD.CreateParameter("@EmployeeID", adInteger, adParamInput, , Employee.Employeeid)
        CMD.Parameters.Append CMD.CreateParameter("@EmployeeID", adInteger, adParamInput, , Employeeid)
    Else
        'If the FK Is Nothing Then Don't Supply a Value: A NULL Value will be Inserted by The Insert Procedure
        CMD.Parameters.Append CMD.CreateParameter("@EmployeeID", adInteger, adParamInput)
    End If
        If Not (bankbranch Is Nothing) Then
        CMD.Parameters.Append CMD.CreateParameter("@BankBranchID", adInteger, adParamInput, , bankbranch.BankBranchID)
    Else
        'If the FK Is Nothing Then Don't Supply a Value: A NULL Value will be Inserted by The Insert Procedure
        CMD.Parameters.Append CMD.CreateParameter("@BankBranchID", adInteger, adParamInput)
    End If
    CMD.Parameters.Append CMD.CreateParameter("@AccountName", adVarChar, adParamInput, 80, AccountName)
    CMD.Parameters.Append CMD.CreateParameter("@AccountNumber", adVarChar, adParamInput, 50, AccountNumber)
    CMD.Parameters.Append CMD.CreateParameter("@AccountType", adVarChar, adParamInput, 50, AccountType)
    CMD.Parameters.Append CMD.CreateParameter("SwiftCode", adVarChar, adParamInput, 50, SwiftCode)
    CMD.Parameters.Append CMD.CreateParameter("@IsMainAccount", adBoolean, adParamInput, , IsMainAccount)
    CMD.Parameters.Append CMD.CreateParameter("@Deleted", adBoolean, adParamInput, , Deleted)
    CMD.Execute
    
    Update = 0
    Exit Function
ErrorHandler:
    MsgBox "An error has occurred while Updating EmployeeBankAccount" & vbNewLine & _
    err.Description, vbExclamation, APP_TITLE
    Update = err.Number
End Function

'Method To Delete the Object EmployeeBankAccount
Public Function Delete() As Long
    Dim CMD As adodb.Command
    
    On Error GoTo ErrorHandler
    Set CMD = New adodb.Command
    CMD.ActiveConnection = con
    CMD.CommandType = adCmdStoredProc
    CMD.CommandText = "spDeleteEmployeeBankAccount"
            
    CMD.Parameters.Append CMD.CreateParameter("@EmployeeBankAccountID", adInteger, adParamInput, , EmployeeBankAccountID)
    CMD.Execute
    Delete = 0
    Exit Function
ErrorHandler:
    MsgBox "An error has occurred while Deleting EmployeeBankAccount" & vbNewLine & _
    err.Description, vbExclamation, APP_TITLE
    Delete = err.Number
End Function






'--------------- End Of Class Definition for Class: EmployeeBankAccount  ----------





