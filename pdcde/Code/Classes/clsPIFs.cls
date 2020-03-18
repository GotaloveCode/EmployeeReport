VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPIFs"
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

Public Sub add(ByVal NewPIF As clsPIF)
    On Error GoTo ErrorHandler
    mCol.add NewPIF
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while adding a new Bank" & vbNewLine & _
    err.Description, vbExclamation, APP_TITLE
End Sub

'Function To get Bank object at a given index
Public Function Item(varIndex As Variant) As clsPIF
    On Error GoTo ErrorHandler
    Set Item = mCol.Item(varIndex)
    
    Exit Function
ErrorHandler:
    MsgBox "An error occurred while Retrieving Bank Information" & vbNewLine & _
        err.Description, vbExclamation, APP_TITLE
    Set Item = Nothing
End Function

'Function To get the Count of  Objects in the Collection
Public Function count() As Long
    On Error GoTo ErrorHandler
    count = mCol.count
    
    Exit Function
ErrorHandler:
    MsgBox "An error has occurred while retrieving the Count of Bank objects" & vbNewLine & _
        err.Description, vbExclamation, APP_TITLE
    count = 0
End Function

'Method To Remove An Item
Public Sub remove(varIndex As Variant)
    On Error GoTo ErrorHandler
    mCol.remove varIndex
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred While Trying To Remove Bank From the Collection" & vbNewLine & _
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

'Method To GetAllBanks Including Deleted i.e. Deleted=1
'Public Sub GetAllPIFcategories()
'    Dim CMD As ADODB.Command
'    Dim rs As ADODB.Recordset
'    Dim objPIF As clsPIF
'
'    On Error GoTo ErrorHandler
'
'    'First populate Required Collections
'
'    'Clear Collection
'    Me.Clear
'
'    Set CMD = New ADODB.Command
'    CMD.ActiveConnection = con
'    CMD.CommandType = adCmdStoredProc
'    CMD.CommandText = "spSelectAllPIFcategories"
'    Set rs = CMD.Execute()
'
'    If Not (rs.BOF Or rs.EOF) Then
'        rs.MoveFirst
'        Do Until rs.EOF
'            Set objPIF = New clsPIF
'            With objPIF
'                If Not IsNull(rs!CategoryID) Then .ID = rs!CategoryID
'                If Not IsNull(rs!Category) Then .Category = rs!Category
'
'            End With
'            Me.add objPIF
'            rs.MoveNext
'        Loop
'    End If
'
'    Exit Sub
'
'ErrorHandler:
'    MsgBox "An Error has occurred while Populating Bank Objects" & vbNewLine & _
'        err.Description, vbExclamation, APP_TITLE
'End Sub

 
Public Function FindPIFByID(ByVal theID As Long) As clsPIF
    Dim Found As Boolean
    Dim i As Long
    Dim ThePIF As clsPIF
    
    On Error GoTo ErrorHandler
    For i = 1 To Me.count
        If Me.Item(i).ID = theID Then
            Set ThePIF = Me.Item(i)
            Found = True
        End If
    Next i
    
    If Found Then
        Set FindPIFByID = ThePIF
    Else
        Set FindPIFByID = Nothing
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred while Searching for Bank" & vbNewLine & _
        err.Description, vbExclamation, APP_TITLE
    Set FindPIFByID = Nothing
End Function






'-------------- End Of Collection Definition for: BANKS ------










