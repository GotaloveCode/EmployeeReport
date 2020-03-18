VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Period"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private lngPeriodId As Long
Private lngCompanyId As Long
Private strPeriodMonth As String
Private lngPeriodYear As Long
Private dtOpenDate As Date
Private dtCloseDate As Date
Private strStatus As String
Private colAllPeriods As Collection

Public Property Get AllPeriods() As Collection
    Set AllPeriods = colAllPeriods
End Property
Public Property Set AllPeriods(ByVal vNew As Collection)
    Set colAllPeriods = vNew
End Property
Public Property Get PeriodId() As Long
    PeriodId = lngPeriodId
End Property
Public Property Let PeriodId(ByVal vNew As Long)
    lngPeriodId = vNew
End Property

Public Property Get CompanyId() As Long
    CompanyId = lngCompanyId
End Property
Public Property Let CompanyId(ByVal vNew As Long)
    lngCompanyId = vNew
End Property

Public Property Get PeriodMonth() As String
    PeriodMonth = strPeriodMonth
End Property
Public Property Let PeriodMonth(ByVal vNew As String)
    strPeriodMonth = vNew
End Property

Public Property Get PeriodYear() As Long
    PeriodYear = lngPeriodYear
End Property
Public Property Let PeriodYear(ByVal vNew As Long)
    lngPeriodYear = vNew
End Property
Public Property Get CloseDate() As Date
    CloseDate = dtCloseDate
End Property
Public Property Let CloseDate(ByVal vNew As Date)
    dtCloseDate = vNew
End Property

Public Property Get OpenDate() As Date
    OpenDate = dtOpenDate
End Property
Public Property Let OpenDate(ByVal vNew As Date)
    dtOpenDate = vNew
End Property

Public Property Get Status() As String
    Status = strStatus
End Property
Public Property Let Status(ByVal vNew As String)
    strStatus = vNew
End Property
'Public Function OpenPeriod() As Long
'    On Error GoTo errHandler
'        OpenPeriod = 1
'        sql = ""
'        sql = "Insert Into tblPeriods(CompanyId,PeriodMonth,PeriodYear,OpenDate,CloseDate,Status)values("
'        sql = sql & Me.CompanyId & ",'" & Me.PeriodMonth & "'," & Me.PeriodYear & ",'" & DateString(Me.OpenDate) & "',"
'        sql = sql & "'" & DateString(Me.CloseDate) & "','" & Me.Status & "')"
'        CConnect.ExecuteThis (sql)
'        OpenPeriod = 0
'    Exit Function
'errHandler:
'    OpenPeriod = err.Number
'End Function
'Public Function ClosePeriod() As Long
'    On Error GoTo errHandler
'        ClosePeriod = 1
'        sql = ""
'        sql = "Update tblPeriods set CloseDate='" & DateString(Date) & "',Status='" & Me.Status & "' Where Periodid=" & Me.PeriodId & ""
'        CConnect.ExecuteThis (sql)
'        ClosePeriod = 0
'    Exit Function
'errHandler:
'    ClosePeriod = err.Number
'End Function

Public Function GetOpenPeriod() As Period
    On Error GoTo errHandler
        Dim blnfound As Boolean
        Dim i As Long
        
        blnfound = False
        For i = 1 To Me.AllPeriods.Count
            If LCase(Me.AllPeriods(i).Status) = "open" Then
                blnfound = True
                Set GetOpenPeriod = Me.AllPeriods(i)
                Exit For
            End If
        Next i
        
        If blnfound = False Then
            Set GetOpenPeriod = Nothing
        End If
    Exit Function
errHandler:
    Set GetOpenPeriod = Nothing
End Function
Public Function GetAllPeriods() As Long
    On Error GoTo errHandler
        Dim oPerd As Period
        Dim rs As New ADODB.Recordset
        
        GetAllPeriods = 1
        
        '//Clear the collection
        Do Until Me.AllPeriods.Count <= 0
            Me.AllPeriods.Remove (Me.AllPeriods.Count)
        Loop
        
        sql = ""
        sql = "Select * From tblPeriods"
        Set rs = CConnect.GetRecordSet(sql)
        
        If Not (rs Is Nothing) Then
            If Not rs.EOF Then
                rs.MoveFirst
                Do While Not rs.EOF
                    Set oPerd = New Period
                    With oPerd
                        .PeriodId = rs!PeriodId
                        .CompanyId = IIf(IsNull(rs!CompanyId), 0, rs!CompanyId)
                        .PeriodMonth = IIf(IsNull(rs!PeriodMonth), "", rs!PeriodMonth)
                        .PeriodYear = IIf(IsNull(rs!PeriodYear), 0, rs!PeriodYear)
                        .OpenDate = IIf(IsNull(rs!OpenDate), Date, rs!OpenDate)
                        .CloseDate = IIf(IsNull(rs!CloseDate), Date, rs!CloseDate)
                        .Status = IIf(IsNull(rs!Status), "", rs!Status)
                        Me.AllPeriods.Add oPerd
                    End With
                    rs.MoveNext
                Loop
            End If
        End If
        
        GetAllPeriods = 0
    Exit Function
errHandler:
    GetAllPeriods = err.Number
End Function

Private Sub Class_Initialize()
    Set colAllPeriods = New Collection
End Sub

Private Sub Class_Terminate()
    Set colAllPeriods = Nothing
End Sub
