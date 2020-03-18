VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BankBranch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'===========START OF CLASS BankBranch =============

'These are Private Variables Handling The Fields in the Classes

Private lngBankBranchID As Long 'This is a Primary Key Column
Private objBank As Bank
'Private lngBankID As Long  'this is A Foregn key Column
Private strBranchCode As String
Private strBranchName As String
Private strComments As String
Private blnDeleted As Boolean

Private Sub Class_Initialize()
    'Instantiate the Object Variables
    On Error GoTo ErrorHandler
    
    Set objBank = New Bank
    Exit Sub
ErrorHandler:
    MsgBox "An error has Ocuured" & vbNewLine & err.Description, vbExclamation, APP_TITLE
End Sub

Private Sub Class_Terminate()
    'Destroy the Object Variables
    On Error GoTo ErrorHandler
    
    Set objBank = Nothing
    Exit Sub
ErrorHandler:
    MsgBox "An error has occurred" & vbNewLine & err.Description, vbExclamation, APP_TITLE
End Sub

' === Property Procedures For Object BankBranch ===
    
Public Property Get BankBranchID() As Long
    BankBranchID = lngBankBranchID
End Property
    
Public Property Let BankBranchID(ByVal vNew As Long)
    lngBankBranchID = vNew
End Property
    
    'This is an Object Variable: objBank
Public Property Get Bank() As Bank
    Set Bank = objBank
End Property

Public Property Set Bank(ByVal vNew As Bank)
    Set objBank = vNew
End Property
    
    
Public Property Get BranchCode() As String
    BranchCode = strBranchCode
End Property
    
Public Property Let BranchCode(ByVal vNew As String)
    strBranchCode = vNew
End Property
    
Public Property Get BranchName() As String
    BranchName = strBranchName
End Property
    
Public Property Let BranchName(ByVal vNew As String)
    strBranchName = vNew
End Property
    
Public Property Get Comments() As String
    Comments = strComments
End Property
    
Public Property Let Comments(ByVal vNew As String)
    strComments = vNew
End Property
    
Public Property Get Deleted() As Boolean
    Deleted = blnDeleted
End Property
    
Public Property Let Deleted(ByVal vNew As Boolean)
    blnDeleted = vNew
End Property

'===METHODS FOR Object BankBranch

'Method To insert a new BankBranch object
Public Function InsertNew() As Long
    Dim cmd As ADODB.Command
    
    On Error GoTo ErrorHandler
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "spInsertBankBranch"
    
    'The Primary Key Column
    cmd.Parameters.Append cmd.CreateParameter("@BankBranchID", adInteger, adParamOutput)
    If Not (Bank Is Nothing) Then
        cmd.Parameters.Append cmd.CreateParameter("@BankID", adInteger, adParamInput, , Bank.BankID)
    Else
        'If the FK Is Nothing Then Don't Supply a Value: A NULL Value will be Inserted by The Insert Procedure
        cmd.Parameters.Append cmd.CreateParameter("@BankID", adInteger, adParamInput)
    End If
    cmd.Parameters.Append cmd.CreateParameter("@BranchCode", adVarChar, adParamInput, 20, BranchCode)
    cmd.Parameters.Append cmd.CreateParameter("@BranchName", adVarChar, adParamInput, 100, BranchName)
    cmd.Parameters.Append cmd.CreateParameter("@Comments", adVarChar, adParamInput, 150, Comments)
    cmd.Parameters.Append cmd.CreateParameter("@Deleted", adBoolean, adParamInput, , Deleted)
    cmd.Execute
    lngBankBranchID = cmd.Parameters(0).value
    
    InsertNew = 0
    Exit Function
ErrorHandler:
    MsgBox "An error has occurred while inserting a new BankBranch" & vbNewLine & _
    err.Description, vbExclamation, APP_TITLE
    InsertNew = err.Number
End Function

'Method to Update the Object BankBranch
Public Function Update() As Long
    Dim cmd As ADODB.Command
    
    On Error GoTo ErrorHandler
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "spUpdateBankBranch"
    
    'The Primary Key Column
    cmd.Parameters.Append cmd.CreateParameter("@BankBranchID", adInteger, adParamInput, , BankBranchID)
        If Not (Bank Is Nothing) Then
        cmd.Parameters.Append cmd.CreateParameter("@BankID", adInteger, adParamInput, , Bank.BankID)
    Else
        'If the FK Is Nothing Then Don't Supply a Value: A NULL Value will be Inserted by The Insert Procedure
        cmd.Parameters.Append cmd.CreateParameter("@BankID", adInteger, adParamInput)
    End If
    cmd.Parameters.Append cmd.CreateParameter("@BranchCode", adVarChar, adParamInput, 20, BranchCode)
    cmd.Parameters.Append cmd.CreateParameter("@BranchName", adVarChar, adParamInput, 100, BranchName)
    cmd.Parameters.Append cmd.CreateParameter("@Comments", adVarChar, adParamInput, 150, Comments)
    cmd.Parameters.Append cmd.CreateParameter("@Deleted", adBoolean, adParamInput, , Deleted)
    cmd.Execute
    
    Update = 0
    Exit Function
ErrorHandler:
    MsgBox "An error has occurred while Updating BankBranch" & vbNewLine & _
    err.Description, vbExclamation, APP_TITLE
    Update = err.Number
End Function

'Method To Delete the Object BankBranch
Public Function Delete() As Long
    Dim cmd As ADODB.Command
    
    On Error GoTo ErrorHandler
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "spDeleteBankBranch"
            
    cmd.Parameters.Append cmd.CreateParameter("@BankBranchID", adInteger, adParamInput, , BankBranchID)
    cmd.Execute
    Delete = 0
    Exit Function
ErrorHandler:
    MsgBox "An error has occurred while Deleting BankBranch" & vbNewLine & _
    err.Description, vbExclamation, APP_TITLE
    Delete = err.Number
End Function






'--------------- End Of Class Definition for Class: BankBranch  ----------








