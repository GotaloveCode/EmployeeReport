VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CompanyBanks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ColCompanyBank As Collection
Private tempBranch As BankBranch
Private tempBranches As BankBranches
Private tempCbank As CompanyBank

Private Sub Class_Initialize()
    Set ColCompanyBank = New Collection
    Set tempBranches = New BankBranches
  tempBranches.GetAllBankBranches
End Sub

Private Sub Class_Terminate()
    Set ColCompanyBank = Nothing
End Sub


Public Sub add(ByVal vNew As CompanyBank)
    ColCompanyBank.add vNew
    
End Sub

Public Function Item(ByVal vNew As Variant) As CompanyBank
   Set Item = ColCompanyBank.Item(vNew)
End Function

Public Sub remove(ByVal vNew As Variant)
    ColCompanyBank.remove vNew
End Sub

Public Function count() As Long
    count = ColCompanyBank.count
End Function

Public Sub Clear()
    Do Until ColCompanyBank.count <= 0
        ColCompanyBank.remove ColCompanyBank.count
    Loop
End Sub

Public Function FindCompanyBank(ByVal vNew As String) As CompanyBank
    Dim i As Long
    Dim Found As Boolean
    Found = False
    For i = 1 To Me.count
        If Me.Item(i).AccountNo = vNew Then
            Set FindCompanyBank = Me.Item(i)
            Found = True
            Exit For
        End If
    Next i
    
    If Not Found Then Set FindCompanyBank = Nothing
    
End Function


Public Function getAllCompanyBanks()
    Set rs = CConnect.GetRecordSet("SELECT * FROM tblCompanyBank")
    If Not rs Is Nothing Then
        If Not (rs.BOF And rs.EOF) Then
            rs.MoveFirst
            
            Do Until rs.EOF
            
             Set tempCbank = New CompanyBank
             With tempBranch
            
                
               tempCbank.AccountNo = IIf(IsNull(rs!Account_Number), vbNullString, rs!Account_Number)
               tempCbank.AccountType = IIf(IsNull(rs!AccountType), vbNullString, rs!AccountType)
               tempCbank.MainAccount = IIf(IsNull(rs!Default_bank), "False", IIf(rs!Default_bank = True, "True", "False"))
               
                Set tempBranch = New BankBranch
                Set tempBranch = tempBranches.FindBankBranchByID(rs!Bankbranch_id)
                If Not tempBranch Is Nothing Then
                Set tempCbank.BankBranch = tempBranch
                End If
               
             
            
               
               Me.add tempCbank
             End With
                rs.MoveNext
            Loop
        End If
    End If
End Function
