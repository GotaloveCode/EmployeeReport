VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CompanyBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private objBankBranch As bankbranch
Private strAccountNo As String
Private strAccountType As String
Private blnMainAccount As Boolean
Private objBranch As bankbranch





Public Property Get MainAccount() As Boolean
   MainAccount = blnMainAccount
End Property

Public Property Let MainAccount(ByVal vnew As Boolean)
    blnMainAccount = vnew
End Property

Public Property Get AccountType() As String
   AccountType = strAccountType
End Property

Public Property Let AccountType(ByVal vnew As String)
    strAccountType = vnew
End Property

Public Property Get AccountNo() As String
   AccountNo = strAccountNo
End Property

Public Property Let AccountNo(ByVal vnew As String)
    strAccountNo = vnew
End Property


Public Property Get bankbranch() As bankbranch
    Set bankbranch = objBankBranch
End Property

Public Property Set bankbranch(ByVal vnew As bankbranch)
    Set objBankBranch = vnew
End Property


Private Sub Class_Initialize()
    Set objBankBranch = New bankbranch
End Sub

Private Sub Class_Terminate()
    Set objBankBranch = Nothing
End Sub
