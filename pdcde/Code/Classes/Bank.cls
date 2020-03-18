VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'===========START OF CLASS Bank =============

'These are Private Variables Handling The Fields in the Classes

Private lngBankID As Long 'This is a Primary Key Column
Private strBankCode As String
Private strBankName As String
Private strComments As String
Private blnDeleted As Boolean

' === Property Procedures For Object Bank ===
    
Public Property Get BankID() As Long
    BankID = lngBankID
End Property
    
Public Property Let BankID(ByVal vNew As Long)
    lngBankID = vNew
End Property
    
Public Property Get BankCode() As String
    BankCode = strBankCode
End Property
    
Public Property Let BankCode(ByVal vNew As String)
    strBankCode = vNew
End Property
    
Public Property Get BankName() As String
    BankName = strBankName
End Property
    
Public Property Let BankName(ByVal vNew As String)
    strBankName = vNew
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

'===METHODS FOR Object Bank

'Method To insert a new Bank object
Public Function InsertNew() As Long
    Dim cmd As ADODB.Command
    
    On Error GoTo ErrorHandler
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "spInsertBank"
    
    'The Primary Key Column
    cmd.Parameters.Append cmd.CreateParameter("@BankID", adInteger, adParamOutput)
    cmd.Parameters.Append cmd.CreateParameter("@BankCode", adVarChar, adParamInput, 30, BankCode)
    cmd.Parameters.Append cmd.CreateParameter("@BankName", adVarChar, adParamInput, 80, BankName)
    cmd.Parameters.Append cmd.CreateParameter("@Comments", adVarChar, adParamInput, 150, Comments)
    cmd.Parameters.Append cmd.CreateParameter("@Deleted", adBoolean, adParamInput, , Deleted)
    cmd.Execute
    lngBankID = cmd.Parameters(0).value
    
    InsertNew = 0
    Exit Function
ErrorHandler:
    MsgBox "An error has occurred while inserting a new Bank" & vbNewLine & _
    err.Description, vbExclamation, APP_TITLE
    InsertNew = err.Number
End Function

'Method to Update the Object Bank
Public Function Update() As Long
    Dim cmd As ADODB.Command
    
    On Error GoTo ErrorHandler
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "spUpdateBank"
    
    'The Primary Key Column
    cmd.Parameters.Append cmd.CreateParameter("@BankID", adInteger, adParamInput, , BankID)
    cmd.Parameters.Append cmd.CreateParameter("@BankCode", adVarChar, adParamInput, 30, BankCode)
    cmd.Parameters.Append cmd.CreateParameter("@BankName", adVarChar, adParamInput, 80, BankName)
    cmd.Parameters.Append cmd.CreateParameter("@Comments", adVarChar, adParamInput, 150, Comments)
    cmd.Parameters.Append cmd.CreateParameter("@Deleted", adBoolean, adParamInput, , Deleted)
    cmd.Execute
    
    Update = 0
    Exit Function
ErrorHandler:
    MsgBox "An error has occurred while Updating Bank" & vbNewLine & _
    err.Description, vbExclamation, APP_TITLE
    Update = err.Number
End Function

'Method To Delete the Object Bank
Public Function Delete() As Long
    Dim cmd As ADODB.Command
    
    On Error GoTo ErrorHandler
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "spDeleteBank"
            
    cmd.Parameters.Append cmd.CreateParameter("@BankID", adInteger, adParamInput, , BankID)
    cmd.Execute
    Delete = 0
    Exit Function
ErrorHandler:
    MsgBox "An error has occurred while Deleting Bank" & vbNewLine & _
    err.Description, vbExclamation, APP_TITLE
    Delete = err.Number
End Function






'--------------- End Of Class Definition for Class: Bank  ----------








