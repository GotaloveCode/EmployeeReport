VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Cnn As adodb.Connection
Private cnnPayData As Connection
Private cnnDis As Connection
Private cnnRec As Connection
Private cnnSec As Connection
Private rst4 As Recordset
Private rst2 As Recordset
Private rst3 As Recordset
Public RDWorking As Boolean
Public HasErr As Boolean
Public LPlanner As Boolean
Public ALCode As String
Public EDBase As String
Public DisBase As String
Public RecBase As String
Public APath As String
Public ConnUnsuccessfull As Boolean
Dim sQL As String


Public Sub CConPayData()
On Error GoTo errHandler

Set cnnPayData = New Connection
cnnPayData.Open EDBase

'Set cnnPayData = New ADODB.Connection
'cnnPayData.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & APath & "\" & EDBase & ".mdb;Persist Security Info=False"
'cnnPayData.Open

Exit Sub
errHandler:
    MsgBox "Invalid database connection to the source database (Omnis Paydata). Contact your systems administrator.", vbInformation
    End

    
End Sub

Public Sub cConDis()
On Error GoTo errHandler

'Set cnnDis = New Connection
'cnnDis.Open DisBase

Set cnnDis = New adodb.Connection
cnnDis.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & APath & "\" & DisBase & ".mdb;Persist Security Info=False"
cnnDis.Open

Exit Sub
errHandler:
    MsgBox err.Description, vbExclamation
    End


End Sub

Public Sub cConRec()

On Error GoTo errHandler

'Set cnnDis = New Connection
'cnnDis.Open RecBase


Set cnnRec = New adodb.Connection
cnnRec.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & APath & "\" & RecBase & ".mdb;Persist Security Info=False"
cnnRec.Open

Exit Sub
errHandler:
    MsgBox err.Description, vbExclamation
    End

End Sub

Public Sub cConSecurity()
On Error GoTo errHandler

'Set cnnSec = New Connection
'cnnSec.Open "HLSec"

Set cnnSec = New adodb.Connection
cnnSec.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & APath & "\HLsec.mdb;Persist Security Info=False"
cnnSec.Open

Exit Sub
errHandler:
    MsgBox err.Description, vbExclamation
    End

End Sub

Public Function GetSecurity(mySQL As String) As Recordset
On Error GoTo errHandler

Dim RsT As Recordset
Set RsT = New Recordset

RsT.Open (mySQL), cnnSec, adOpenStatic, adLockOptimistic


Set GetSecurity = RsT
Set RsT = Nothing

Exit Function
errHandler:
    HasErr = True
    MsgBox err.Description
    'FClose.UnloadForm frm
    
End Function

Public Function GetPayData(mySQL As String) As Recordset
    On Error GoTo errHandler
    
    Dim RsT As Recordset
    Set RsT = New Recordset
    
    RsT.Open (mySQL), cnnPayData, adOpenStatic, adLockReadOnly
    
    
    Set GetPayData = RsT
    Set RsT = Nothing
    
    Exit Function
errHandler:
        HasErr = True
        MsgBox err.Description
        'FClose.UnloadForm frm
        
End Function

Public Function GetDisData(mySQL As String) As Recordset
    On Error GoTo errHandler
    
    Dim RsT As Recordset
    Set RsT = New Recordset
    
    RsT.Open (mySQL), cnnDis, adOpenStatic, adLockReadOnly
    
    
    Set GetDisData = RsT
    Set RsT = Nothing
    
    Exit Function
errHandler:
    HasErr = True
    MsgBox err.Description
    'FClose.UnloadForm frm
    
End Function

Public Function GetRecData(mySQL As String) As Recordset
    On Error GoTo errHandler
    
    Dim RsT As Recordset
    Set RsT = New Recordset
    
    RsT.Open (mySQL), cnnRec, adOpenStatic, adLockReadOnly
    
    
    Set GetRecData = RsT
    Set RsT = Nothing
    
    Exit Function
errHandler:
    HasErr = True
    MsgBox err.Description
    'FClose.UnloadForm frm
    
End Function

Public Sub CCon()
    Dim mHandle As Integer
    Dim i As Integer
    Dim retVal As Long
    Dim retVal2 As Long
    Dim R&
    Dim mTextLine As String
    'Dim ServerName As String, catalog As String, userID As String, passwd As String
    Dim inFileName As String
    
    On Error GoTo errHandler
    
    mHandle = FreeFile

    inFileName = App.Path & "\lscript.ini"
    
    If Trim(Dir(inFileName)) = "" Then
        MsgBox "Please create the logon script first!", vbInformation
        Open inFileName For Append As #mHandle
        Shell "notepad.exe " & inFileName, vbMaximizedFocus
        Close #mHandle
        End
    End If
         
    Set Cnn = New Connection
    Open inFileName For Input As #mHandle
   
    i = 0
    
    Do While EOF(mHandle) = False
        Line Input #mHandle, mTextLine
        If Len(Trim(mTextLine)) > 0 Then
            i = i + 1
            R = InStr(mTextLine, vbTab)
            ServerName = Left(mTextLine, R - 1)
            
            If ServerName <> "#" Then
                mTextLine = Mid(mTextLine, R + 1)
        
                R = InStr(mTextLine, vbTab)
                
                catalog = Left(mTextLine, R - 1)
                mTextLine = Mid(mTextLine, R + 1)
        
                R = InStr(mTextLine, vbTab)
                
                UserID = Left(mTextLine, R - 1)
                mTextLine = Mid(mTextLine, R + 1)
        
                R = InStr(mTextLine, vbTab)
                
                passwd = Left(mTextLine, R - 1)
                
                Cnn.ConnectionTimeout = 600
                
                'connection_string = "Provider=msdasql.1;driver={sql server};server=" & serverName & ";uid=" & UserID & ";passwd=" & passwd & ";database=" & catalog
                connection_string = "Provider=SQLOLEDB.1;Data Source=" & ServerName & ";Initial Catalog=" & catalog & ";User ID=" & UserID & ";Password=" & passwd & ";Persist Security Info=False"
                Cnn.ConnectionString = connection_string
                Cnn.Open
            End If
        End If
    Loop
    
    Close #mHandle
    
    If Cnn.State = 0 Then
        MsgBox "Your connection parameters are incorrect. Refer to your system configuration files!", vbInformation
        End
    Else
        'Set the HRCORE connection
        retVal = HRCon.SetConnection(Cnn)
        If retVal = 0 Then
            retVal2 = HRMSECCon.SetConnection(Cnn)
            If retVal2 <> 0 Then
                MsgBox "Database Connection Could not be synchronized with HR Security", vbExclamation, "Security Connection"
                End
            End If
        Else
            MsgBox "Database connection could not be set for the HR CORE Engine", vbExclamation, "HR CORE Engine"
            End
        End If
    End If

Exit Sub
errHandler:
    If err.Number = -2147467259 Then
        MsgBox "You connection to the server has been timed-out! Please try again.", vbInformation
    Else
        MsgBox "Error on Line " & i & " in your config file. Contact your system administrator!", vbInformation
        Debug.Print err.Description
    End If
    End
End Sub


Public Function GetRecordSet(mySQL As String) As Recordset
    On Error GoTo errHandler
    Dim RsT As Recordset
    Set RsT = New Recordset
    
    'Rst.Open (mysql), Cnn, adOpen, adLockOptimistic
    RsT.Source = mySQL
    Set RsT.ActiveConnection = Cnn
    RsT.CursorLocation = adUseClient
    RsT.LockType = adLockOptimistic
    RsT.Open
    Set GetRecordSet = RsT
    Set RsT = Nothing

    Exit Function
errHandler:
    HasErr = True
    MsgBox err.Description, vbExclamation

    
End Function

Public Function setGlobalRecordset() As adodb.Recordset
    Dim rec_s As New adodb.Recordset
    Dim sTemp As String
    Dim aSection As String, tCount As Integer, tFilter As String
    Dim aNode As String
    
    On Error GoTo ErrorHandler
    
    tFilter = "0"
    
    If isDepartmentOrSection(deptConstraint) = True Then
        aSection = getChildNodes(deptConstraint)
        aSection = Mid(aSection, tCount + 1)
        While aSection <> ""
            tCount = InStr(aSection, vbTab)
            aNode = Left(aSection, tCount - 1)
            If tFilter <> "0" Then
                tFilter = tFilter & " or (LCode = '" & aNode & "')"
            Else
                tFilter = "(LCode = '" & aNode & "')"
            End If
            aSection = Mid(aSection, tCount + 1)
        Wend
        If tFilter <> "0" Then tFilter = "(" & tFilter & ")"
    End If
    
    If frmMain2.cboStructure.Text <> "All Records" And frmMain2.cboStructure.Text <> "" Then 'Filter by Structure
        If tFilter <> "0" Then
            Set rec_s = deptFilter("SELECT * FROM pVwRsGlob WHERE  scode = '" & frmMain2.cboStructure.Tag & "' " & " AND " & tFilter & " ORDER BY EmpCode")
        Else
            Set rec_s = deptFilter("SELECT * FROM pVwRsGlob WHERE scode = '" & frmMain2.cboStructure.Tag & "' ORDER BY EmpCode")
        End If
    Else
        If tFilter <> "0" Then
            Set rec_s = deptFilter("SELECT * FROM pVwRsGlob WHERE '" & tFilter & "'  ORDER BY EmpCode")
        Else
            Set rec_s = deptFilter("SELECT * FROM pVwRsGlob  ORDER BY EmpCode")
        End If
    End If
    
    With frmMain2
    ' Filter by Terms and Grade
        If rec_s.Filter = 0 Then
            If (.cboTerms.Text <> "All Records" And .cboTerms.Text <> "") And (.cboCat.Text <> "All Records" And .cboCat.Text <> "") Then
                rec_s.Filter = "(ECategory_Id = '" & .cboCat.Tag & "' AND Terms = '" & .cboTerms.Text & "')"
            ElseIf (.cboTerms.Text <> "All Records" And .cboTerms.Text <> "") And (.cboCat.Text = "All Records" Or .cboCat.Text = "") Then
                rec_s.Filter = "(Terms = '" & .cboTerms.Text & "')"
            ElseIf (.cboTerms.Text = "All Records" Or .cboTerms.Text = "") And (.cboCat.Text <> "All Records" And .cboCat.Text <> "") Then
                rec_s.Filter = "(ECategory_id = '" & .cboCat.Tag & "')"
            End If
        Else
            sTemp = rec_s.Filter
            rec_s.Filter = adFilterNone
            If (.cboTerms.Text <> "All Records" And .cboTerms.Text <> "") And (.cboCat.Text <> "All Records" And .cboCat.Text <> "") Then
                rec_s.Filter = sTemp & " AND (ECategory_id = '" & .cboCat.Tag & "' AND Terms = '" & _
                        .cboTerms.Text & "')"
            ElseIf (.cboTerms.Text <> "All Records" And .cboTerms.Text <> "") And (.cboCat.Text = "All Records" Or .cboCat.Text = "") Then
                rec_s.Filter = sTemp & " And (Terms = '" & .cboTerms.Text & "')"
            ElseIf (.cboTerms.Text = "All Records" Or .cboTerms.Text = "") And (.cboCat.Text <> "All Records" And .cboCat.Text <> "") Then
                rec_s.Filter = sTemp & " AND (ECategory_id = '" & .cboCat.Tag & "')"
            ElseIf (.cboTerms.Text = "All Records" Or .cboTerms.Text = "") And (.cboCat.Text = "All Records" Or .cboCat.Text = "") Then
                rec_s.Filter = sTemp
            End If
        End If
        'Debug.Print rec_s.Filter
        
        'Debug.Print rec_s.RecordCount
    
    End With
    
    Set setGlobalRecordset = rec_s
    Set rec_s = Nothing
    
    Exit Function
    
ErrorHandler:
    Set setGlobalRecordset = Nothing
End Function

Public Function deptFilter(mySQL As String) As adodb.Recordset
    Dim rec_d As New adodb.Recordset, r_tmp As New adodb.Recordset
    Dim dSection As String, CurNode As String
    Dim rCount As Integer
    On Error GoTo errHandler
    Set rec_d = GetRecordSet(mySQL)
    
    With rec_d
        If UCase(deptConstraint) <> "ALL" And Trim(deptConstraint) <> "" Then
            If isDivision(deptConstraint) = True Then
                .Filter = "(SCode = '" & deptConstraint & "')"
            End If
        End If
    End With
    Set deptFilter = rec_d
    Set rec_d = Nothing
    Exit Function
errHandler:
    
End Function

Public Function ExecuteSec(mySQL As String)
    On Error GoTo errHandler
    
    Dim CMD As New Command
    
    With CMD
        .ActiveConnection = cnnSec
        .CommandType = adCmdText
        .CommandText = mySQL
        .Execute
    End With
    
    Exit Function
errHandler:
'    HasErr = True
'    MsgBox err.Description, vbExclamation

    
End Function

Public Sub ExecuteSql(mySQL As String)
    Dim CMD As New Command
    Dim AValues As String
    Dim AAction As String
    Dim MyAction As String
    Dim ATable As String
    Dim MyValues As String
    Dim MyDest As String
    Dim CurrentCode As String
    Dim qRec As New adodb.Recordset
    Dim tmpStr As String
    Dim tmpStr2 As String
    Dim errCount As Integer
    
    On Error GoTo errHandler
Restart:
    With CMD
        .ActiveConnection = Cnn
        .CommandType = adCmdText
        .CommandText = mySQL
        .Prepared = True
        .Execute
    End With
    'log the action
    If Action <> "" Then
        Dim rstAudit As New adodb.Recordset
        
'        Set rstAudit = CConnect.GetRecordSet("INSERT INTO New_AuditTrail (UserID, Username, DateOfAction,TimeOfAction, Action,subsystem)" & _
'                    " VALUES('" & CurrentUser & "','" & CurrentUser & "','" & SQLDate(Date) & "', '" & Now & "','" & Replace(Action, "'", "''") & "','" & SubSystem & "')")
'
   
'
'     Set rstAudit = CConnect.GetRecordSet("INSERT INTO New_AuditTrail (UserID, Username, DateOfAction,TimeOfAction, Action,subsystem)" & _
'                    " VALUES('" & CurrentUser & "','" & CurrentUser & "','" & SQLDate(Date) & "', '" & Now & "','" & Replace(Action, "'", "''") & "','" & SubSystem & "')")
'
    
    
        If Not currUser Is Nothing Then
             currUser.AuditTrail Add_New, Action
        End If
     End If
    
    Action = ""
    Exit Sub
errHandler:
        If err.Number = -2147217900 Then
            s = InStr(err.Description, "near ")
            tmpStr = Mid(err.Description, s + 5, 3)
            errCount = errCount + 1
            tmpStr2 = "'" & tmpStr
            E = InStr(1, mySQL, tmpStr)
            mySQL = Replace(mySQL, tmpStr, tmpStr2)
           
            If errCount > 5 Then
                MsgBox "Your request cannot be resolved!", vbExclamation
                Exit Sub
            Else
                GoTo Restart
            End If
        Else
            MsgBox err.Description, vbInformation
        End If
End Sub


Public Function ExecuteRec(mySQL As String)
    On Error GoTo errHandler
    
    Dim CMD As New Command
    
    With CMD
        .ActiveConnection = cnnRec
        .CommandType = adCmdText
        .CommandText = mySQL
    '    Debug.Print mysql
        .Execute
    End With
    
    Exit Function
errHandler:
    HasErr = True
    MsgBox err.Description, vbExclamation
End Function

Public Function CColor(frm As Object, MyColor As Long)
    On Error GoTo Hell
    Dim i As Object
    
    frm.BackColor = MyColor
    
    For Each i In frm
        If TypeOf i Is Form Or TypeOf i Is Frame Or TypeOf i Is OptionButton Or TypeOf i Is CheckBox Then ' Or TypeOf i Is PictureBox
            If i.Name <> "lblTitle" And i.Name <> "lblTitle2" Then
                i.BackColor = MyColor
            End If
        End If
        
        If TypeOf i Is Label Then
            i.BackStyle = 0
        End If
    Next i
    
    Exit Function
Hell:
MsgBox err.Description, vbExclamation
End Function

Public Function ResumeDate(StartDate As Date, NoofDays As Double, HDays As Double, LeaveCode As String) As Date
    Dim IncludeHols As Boolean
    Dim SatisHol As Boolean
    Dim SunisHol As Boolean
    Dim CheckforHols As Boolean
    Dim FHDays As Double
    Dim i As Double
    Dim a As Double
        
        IncludeHols = False
        SatisHol = False
        SunisHol = False
        
        If HDays > 0 Then
            FHDays = HDays / 2
        End If
            
        NoofDays = NoofDays + FHDays
        
        a = NoofDays - Int(NoofDays)
        
        Set rst4 = GetRecordSet("SELECT * FROM LeaveTypes WHERE LCode='" & LeaveCode & "'")
            With rst4
                If .RecordCount > 0 Then
                    If !ExcHols = "No" Then IncludeHols = True
                    If !ExcSat = "No" Then SatisHol = True
                    If !ExcSun = "No" Then SunisHol = True
                    If RWorking = "Yes" Then RDWorking = True
                End If
            End With
        Set rst4 = Nothing
        
        If IncludeHols = False And SatisHol = False And SunisHol = False Then
            Dim myCount As Long
            myCount = 0
            
            While myCount < NoofDays
                'if day is a Saturday ,check if saturday is a holiday
                If Weekday(StartDate, vbSunday) = 7 Then
                    If SatisHol = False Then
                        myCount = myCount - 1
                        CheckforHols = False
                    Else
                        'check if it is a public holiday
                        CheckforHols = True
                    End If
                ElseIf Weekday(StartDate, vbSunday) = 1 Then
                    If SunisHol = False Then
                        myCount = myCount - 1
                        CheckforHols = False
                    Else
                        'check if it is a public holiday
                        CheckforHols = True
                    End If
                Else
                        'check if it is a public holiday
                        CheckforHols = True
                End If
                
                If CheckforHols = True Then
                    If IncludeHols = False Then
                        If DayisHoliday(StartDate) = True Then
                            myCount = myCount - 1
                        End If
                    End If
                End If
                
                If a > 0 Then
                    i = 0.5
                    a = 0
                Else
                    i = 1
                End If
                    
                
                StartDate = DateAdd("d", i, StartDate)
                myCount = myCount + 1
            Wend
            
            ResumeDate = StartDate
        Else
            ResumeDate = DateAdd("d", NoofDays, StartDate)
        End If
        
        If RDWorking = True Then
            'ensure you resume work on a working day
            ResumeDate = NextWorkingday(ResumeDate)
        End If
End Function
Public Function NextWorkingday(Datefrom As Date) As Date
    Dim Workdate As Boolean
        Workdate = False
        While Workdate = False
            If DayisWorking(Datefrom) = False Then
                Datefrom = DateAdd("d", 1, Datefrom)
            Else
                Workdate = True
            End If
        Wend
        NextWorkingday = Datefrom
End Function

Public Function LeaveDays(StartDate As Date, EndDate As Date, LeaveCode As String) As Long
    Dim IncludeHols As Boolean
    Dim SatisHol As Boolean
    Dim SunisHol As Boolean
    Dim CheckforHols As Boolean
    
    IncludeHols = False
    SatisHol = False
    SunisHol = False
    
    Set rst4 = GetRecordSet("SELECT * FROM LeaveTypes WHERE LCode='" & LeaveCode & "'")
        With rst4
            If .RecordCount > 0 Then
                If !ExcHols = "No" Then IncludeHols = True
                If !ExcSat = "No" Then SatisHol = True
                If !ExcSun = "No" Then SunisHol = True
            End If
        End With
    Set rst4 = Nothing
    
    If IncludeHols = False Or SatisHol = False Or SunisHol = False Then
        Dim myCount As Long
        myCount = 0
        
        While StartDate < EndDate
            'if day is a Saturday ,check if saturday is a holiday
            If Weekday(StartDate, vbSunday) = 7 Then
                If SatisHol = False Then
                    myCount = myCount - 1
                    CheckforHols = False
                Else
                    'check if it is a public holiday
                    CheckforHols = True
                End If
            ElseIf Weekday(StartDate, vbSunday) = 1 Then
                If SunisHol = False Then
                    myCount = myCount - 1
                    CheckforHols = False
                Else
                    'check if it is a public holiday
                    CheckforHols = True
                End If
            Else
                    'check if it is a public holiday
                    CheckforHols = True
            End If
            
            If CheckforHols = True Then
                If IncludeHols = False Then
                    If DayisHoliday(StartDate) = True Then
                        myCount = myCount - 1
                    End If
                End If
            End If
            
            StartDate = DateAdd("d", 1, StartDate)
            myCount = myCount + 1
        Wend
        LeaveDays = myCount
    Else
        LeaveDays = DateDiff("d", StartDate, EndDate)
    End If
End Function

Public Function DayisHoliday(myDay As Date) As Boolean

    Dim myYear As Long
    myDay = Format(myDay, "dd/mm/yyyy")
    myYear = Format(myDay, "yyyy")
    DayisHoliday = False
    Set rst2 = GetRecordSet("SELECT * FROM Holidays ORDER BY HCode")
        With rst2
            If .RecordCount > 0 Then
                .MoveFirst
            While Not .EOF And DayisHoliday = False
                'check if holiday is a Sunday
                If Weekday(!DDay & "/" & myYear, vbSunday) = 1 Then
                    If myDay = !DayMon & "/" & myYear Or myDay = !DDay & "/" & myYear Then DayisHoliday = True
                ElseIf Weekday(!DDay & "/" & myYear, vbSunday) = 2 Then
                    If myDay = DateAdd("d", 1, (!DDay & "/" & myYear)) Then
                        Set rst3 = GetRecordSet("SELECT * FROM Holidays WHERE DayMon='" & !DDay & "'")
                            With rst3
                                If .RecordCount > 0 Then
                                    DayisHoliday = True
                                End If
                            End With
                        
                        Set rst3 = Nothing
                            
                    Else
                        If (!DDay & "/" & myYear) = myDay Then DayisHoliday = True
                    End If
                Else
                    If (!DDay & "/" & myYear) = myDay Then DayisHoliday = True
                End If
                .MoveNext
            Wend
            
            End If
        End With
        
    Set rst2 = Nothing
End Function

Public Function DayisWorking(DayofYear As Date) As Boolean
    Dim SatWorking As Boolean
    Dim SunWorking As Boolean
    Dim HolWorking As Boolean
    
    Set rst2 = GetRecordSet("SELECT * FROM GeneralOpt")
        With rst2
            If .RecordCount > 0 Then
                If !SatWorking = "Yes" Then SatWorking = True
                If !SunWorking = "Yes" Then SunWorking = True
                If !HolWorking = "Yes" Then HolWorking = True
            End If
        End With
        
    Set rst2 = Nothing

If SatWorking = False Or SunWorking = False Or HolWorking = False Then
    DayisWorking = True
    If Weekday(DayofYear, vbSunday) = 7 Then
        If SatWorking = False Then
            DayisWorking = False
            CheckforHols = False
        Else
            'check if it is a public holiday
            CheckforHols = True
        End If
    ElseIf Weekday(DayofYear, vbSunday) = 1 Then
        If SunWorking = False Then
            DayisWorking = False
            CheckforHols = False
        Else
            'check if it is a public holiday
            CheckforHols = True
        End If
    Else
            'check if it is a public holiday
            CheckforHols = True
    End If
    
    If CheckforHols = True Then
        If HolWorking = False Then
            If DayisHoliday(DayofYear) = True Then
               DayisWorking = False
            Else
                DayisWorking = True
            End If
        Else
            DayisWorking = True
        End If
    End If
Else
    DayisWorking = True
End If
End Function


Public Function LeaveAccrual(EmpCode As String, LCode As String, EndDate As Date, ToDate As Boolean) As Double
    Dim LDays As Long
    Dim RDays As Long
    Dim AccPeriod As Integer
    Dim AccType As String
    Dim DDue As Date
    Dim myPeriod As Integer
    Dim YDays As Integer
    Dim myYear As Long
    Dim Accrues As Boolean
    Dim DDue2 As Date
    Dim myDate As Date
    
    myDate = Format(EndDate, "dd/mm/yyyy")
    myYear = Year(Date)
    YDays = DateDiff("d", "1/1/" & myYear, "31/12/" & myYear) + 1
    Set rst2 = GetRecordSet("SELECT * FROM EmpLeaves WHERE EmpCode='" & EmpCode & "' AND LCode='" & LCode & "'")
        With rst2
            If .RecordCount > 0 Then
                LDays = !days & ""
                If Not IsNull(!Datedue) Then DDue = Format(!Datedue, "dd/mm/yyyy")
                RDays = !RDays & ""
            End If
        End With
    Set rst2 = Nothing
    
    Set rst2 = GetRecordSet("SELECT * FROM LeaveTypes WHERE LCode='" & LCode & "'")
        With rst2
            If .RecordCount > 0 Then
                .MoveFirst
                
                If !Accrued = "Yes" Then
                    AccType = !APeriodType & ""
                    If Not IsNull(!AccrualPeriod) Then
                        AccPeriod = !AccrualPeriod & ""
                    Else
                        AccPeriod = 1
                    End If
                    If AccType = "" Then
                        LeaveAccrual = RDays
                    End If
                Else
                    LeaveAccrual = RDays
                    Set rst2 = Nothing
                    Exit Function
                End If
            End If
        End With
    
    Set rst2 = Nothing
    
    If ToDate = False Then
        Set rst2 = GetRecordSet("SELECT CYearEnd FROM GeneralOpt")
            With rst2
                If .RecordCount > 0 Then
                    If Not IsNull(!CyearEnd) Then
                            DDue2 = Format(DateAdd("d", 1, !CyearEnd), "dd/mm/yyyy")
                            myYear = Year(DDue2) - 1
                            DDue2 = Format(DDue2, "dd/mm") & "/" & myYear
                    Else
                        DDue = "1/1/ & myYear"
                    End If
                End If
            End With
        
        Set srt2 = Nothing
    End If
    
    If DDue < DDue2 Then DDue = DDue2
    
    If AccType = "Years" Then
        myPeriod = DateDiff("y", DDue, myDate)
        
        If Month(myDate) >= 6 Then
            myPeriod = myPeriod + 1
        End If
        
        If myPeriod >= AccPeriod Then
            LeaveAccrual = LDays * AccPeriod + (RDays - LDays)
        End If
        
    ElseIf AccType = "Months" Then
        myPeriod = DateDiff("m", DDue, myDate)
        
        If Day(myDate) >= 15 Then
            myPeriod = myPeriod + 1
        End If
        
        If myPeriod >= AccPeriod Then
            LeaveAccrual = LDays / 12 * myPeriod + (RDays - LDays)
        End If
        
    ElseIf AccType = "Weeks" Then
        myPeriod = DateDiff("d", DDue, myDate)
        AccPeriod = 7 * AccPeriod
        
        If myPeriod >= AccPeriod Then
            LeaveAccrual = LDays / YDays * myPeriod + (RDays - LDays)
        End If
        
    ElseIf AccType = "Days" Then
        myPeriod = DateDiff("d", DDue, myDate)
        If myPeriod >= AccPeriod Then
            LeaveAccrual = LDays / YDays * myPeriod + (RDays - LDays)
        End If
        
    End If
    
    LeaveAccrual = Format(LeaveAccrual, "##############0.0")
    
End Function


Public Function CanEditApp(EmpCode As String, AppCode As String) As Boolean
Set rst2 = GetRecordSet("SELECT * FROM LeaveApproval WHERE EmpCode='" & EmpCode & "'AND AppCode='" & AppCode & "'")
    If HasErr = True Then
        HasErr = False
        Exit Function
    End If
    With rst2
        If .RecordCount > 0 Then
            CanEditApp = False
        Else
            CanEditApp = True
        End If
    End With
Set rst2 = Nothing

Set rst2 = GetRecordSet("SELECT * FROM Application WHERE EmpCode = '" & EmpCode & "' AND AppCode = '" & AppCode & "' AND Status = 'Not Approved'")

With rst2
    If .RecordCount > 0 Then
        CanEditApp = False
    End If
End With

Set rst2 = Nothing

If LPlanner = True Then
    Set rst2 = GetRecordSet("SELECT * FROM LeaveApp WHERE AppCode='" & AppCode & "'AND LCode='" & ALCode & "'")
        With rst2
            If .RecordCount > 0 Then
                CanEditApp = False
            End If
        End With
    
    Set rst2 = Nothing
End If
End Function

Public Function ClearedProbation(DEmployed As Date, SDate As Date, PPeriod As Long, PType As String) As Boolean
    If PType = "Days" Then
        If DateDiff("d", DEmployed, SDate) > PPeriod Then
            ClearedProbation = True
        Else
            ClearedProbation = False
        End If
    ElseIf PType = "Weeks" Then
        If DateDiff("d", DEmployed, SDate) > (PPeriod * 7) Then
            ClearedProbation = True
        Else
            ClearedProbation = False
        End If
    ElseIf PType = "Months" Then
        If DateDiff("m", DEmployed, SDate) > PPeriod Then
            ClearedProbation = True
        Else
            ClearedProbation = False
        End If
    ElseIf PType = "Years" Then
        If DateDiff("y", DEmployed, SDate) > PPeriod Then
            ClearedProbation = True
        Else
            ClearedProbation = False
        End If
    End If
End Function

Public Function LDateDue(DEmployed As Date, PPeriod As Long, PType As String) As Date
    If PType = "Days" Then
        LDateDue = DateAdd("d", PPeriod, DEmployed)
    ElseIf PType = "Weeks" Then
        LDateDue = DateAdd("d", (PPeriod * 7), DEmployed)
    ElseIf PType = "Months" Then
        LDateDue = DateAdd("m", PPeriod, DEmployed)
    ElseIf PType = "Years" Then
        LDateDue = DateAdd("y", PPeriod, DEmployed)
    End If
    
End Function

Public Function WorkingDays(StartDate As Date, EndDate As Date) As Long

    While StartDate <= EndDate
        If DayisWorking(StartDate) = True Then
            WorkingDays = WorkingDays + 1
        End If
            StartDate = DateAdd("d", 1, StartDate)
        
    Wend
End Function
Public Function StaffOnLeave(EmpCode As String, StartDate As Date, EndDate As Date) As Double
Dim TotalStaff As Integer
Dim Onleave As Integer
Dim DCode As String

Set rst2 = GetRecordSet("SELECT * FROM EmpDivisions WHERE EmpCode='" & EmpCode & "'")
    With rst2
        If .RecordCount > 0 Then
            DCode = !DCode & ""
        End If
    End With
Set rst2 = Nothing

Set rst2 = GetRecordSet("SELECT * FROM EmpDivisions WHERE DCode='" & DCode & "'")
              
        TotalStaff = rst2.RecordCount

Set rst2 = Nothing

'Set rst2 = GetRecordSet("SELECT Application.AppCode, EmpDivisions.DCode FROM Application LEFT JOIN EmpDivisions ON Application.EmpCode = EmpDivisions.EmpCode WHERE EmpDivisions.DCode='" & DCode & "' AND Application.Status='Approved' AND Application.ASDate BETWEEN" & "#" & StartDate & "#" & " AND " & "#" & EndDate & "#" & " OR Application.AEDate BETWEEN " & "#" & StartDate & "#" & " AND " & "#" & EndDate & "#" & "")
Set rst2 = GetRecordSet("SELECT DISTINCT Application.EmpCode" & _
        " FROM Application LEFT JOIN EmpDivisions ON Application.EmpCode = EmpDivisions.EmpCode" & _
        " WHERE (((EmpDivisions.DCode)='" & DCode & "') AND ((Application.Status)='Approved') AND (((Application.ASDate) Between " & "#" & StartDate & "#" & " And " & "#" & EndDate & "#" & ") OR ((Application.AEDate) Between " & "#" & StartDate & "#" & " And " & "#" & EndDate & "#" & ")))")
        
    Onleave = rst2.RecordCount + 1
    
Set rst2 = Nothing

If TotalStaff > 0 Then
    StaffOnLeave = Onleave / TotalStaff * 100
Else
    StaffOnLeave = 0
End If
End Function
Public Function NoofHolidays(StartDate As Date, EndDate As Date) As Long
    While StartDate <= EndDate
        If DayisHoliday(StartDate) Then
            NoofHolidays = NoofHolidays + 1
        End If
        StartDate = DateAdd("d", 1, StartDate)
    Wend
    
End Function
Public Function OnLeaveStaff(DCode As String, StartDate As Date, EndDate As Date) As Double
Dim AllStaff As Integer
Dim SOnleave As Integer

Set rst2 = GetRecordSet("SELECT * FROM EmpDivisions WHERE DCode='" & DCode & "'")
        
        AllStaff = rst2.RecordCount

Set rst2 = Nothing


Set rst2 = GetRecordSet("SELECT DISTINCT Application.EmpCode" & _
        " FROM Application LEFT JOIN EmpDivisions ON Application.EmpCode = EmpDivisions.EmpCode" & _
        " WHERE (((EmpDivisions.DCode)='" & DCode & "') AND ((Application.Status)='Approved') AND (((Application.ASDate) Between " & "#" & StartDate & "#" & " And " & "#" & EndDate & "#" & ") OR ((Application.AEDate) Between " & "#" & StartDate & "#" & " And " & "#" & EndDate & "#" & ")))")

    SOnleave = rst2.RecordCount
    
Set rst2 = Nothing

If AllStaff > 0 Then
    OnLeaveStaff = SOnleave / AllStaff * 100
Else
    OnLeaveStaff = 0
End If

End Function

Public Function LogIn(UID As String, Password As String)
Dim rst5 As Recordset
Dim rst6 As Recordset
Dim rst7 As Recordset

End Function

Public Function Crypt(Pwd As Variant) As String
'Dim Pwd As Variant
Dim Temp As String, PwdChr As Long
Dim EncryptKey As Long
'Pwd = Excc

EncryptKey = Int(Sqr(Len(Pwd) * 81)) + 23

For PwdChr = 1 To Len(Pwd)
    Temp = Temp + Chr(Asc(Mid(Pwd, PwdChr, 1)) Xor EncryptKey)
Next PwdChr

'EncryptPass = Temp
Crypt = Temp

End Function

Public Function GenID() As String
LastSecID = 0

Set rs1 = GetRecordSet("SELECT GenID, IDInitials, StartFrom, LastSecID FROM GeneralOpt")

With rs1
    If .RecordCount > 0 Then
        If Not IsNull(!GenID) Then
            If !GenID = "Yes" Then
                GenerateID = True
                If IsNull(!IDInitials) Then
                    If Not IsNull(!LastSecID) Then
                        LastSecID = !LastSecID + 1
                        GenID = LastSecID
                    Else
                        LastSecID = 1
                        GenID = 1
                    End If
                    
                Else
                    If Not IsNull(!LastSecID) Then
                        LastSecID = !LastSecID + 1
                    Else
                        LastSecID = 1
                    End If
                        GenID = !IDInitials & "" & LastSecID
                End If
                ExecuteSql ("UPDATE GeneralOpt SET LastSecid=" & LastSecID & "")
                
            End If
        End If
    End If
End With

Set rs1 = Nothing

End Function

'Public Function GetECode(empcode As String, rsGlob As Recordset) As String
'    With rsGlob
'        If .RecordCount > 0 Then
'            .MoveFirst
'            .Filter = "EmpCode like '" & empcode & "'"
'            If .RecordCount > 0 Then
'                GetECode = !ECode & ""
'            End If
'            .Filter = adFilterNone
'        End If
'    End With
'
'End Function

Function ImportEmp(inFileName As String, TabDilimited As Boolean) As Variant
On Error GoTo errHandler

    Dim mTextLine As String
    Dim arrResult() As String
    Dim mHandle
    Dim i As Integer
    Dim R&
    Dim EmpCode As String
    Dim SurName As String
    Dim ONames As String
    Dim IdNo As String
    Dim DOB As Date
    Dim Gender As String
    Dim DEmp As Date
    Dim NhifNo As String
    Dim NssfNo As String
    Dim PinNo As String
    Dim myEmpType As String
    Dim myTerms As String
    Dim myPay As Currency
    Dim ECategory As String

    If MsgBox("This will import Employee Details hence updating current records. Do you wish to continue", vbInformation + vbYesNo) = vbNo Then
        Exit Function
    End If

    mHandle = FreeFile

    Open inFileName For Input As #mHandle
    i = 0
    Do While Not EOF(mHandle)
        Line Input #mHandle, mTextLine
        If Len(Trim(mTextLine)) > 0 Then
              ReDim Preserve arrResult(i)
              i = i + 1
                   ' Place the line read into array
              arrResult(UBound(arrResult)) = mTextLine
                If TabDilimited = True Then
                    R = InStr(mTextLine, vbTab)

                    EmpCode = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    SurName = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    ONames = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    IdNo = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    If Not IsDate(Left(mTextLine, R - 1)) Then
                        DOB = Date
                    Else
                        DOB = Left(mTextLine, R - 1)
                    End If
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Gender = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    If Not IsDate(Left(mTextLine, R - 1)) Then
                        DEmp = Date
                    Else
                        DEmp = Left(mTextLine, R - 1)
                    End If
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    NhifNo = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    NssfNo = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    PinNo = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    myEmpType = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    myTerms = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    myPay = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)

                    ECategory = mTextLine
                    
                Else
                    R = InStr(mTextLine, ",")

                    EmpCode = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    SurName = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    ONames = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    IdNo = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    If Not IsDate(Left(mTextLine, R - 1)) Then
                        DOB = Date
                    Else
                        DOB = Left(mTextLine, R - 1)
                    End If
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Gender = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    If Not IsDate(Left(mTextLine, R - 1)) Then
                        DEmp = Date
                    Else
                        DEmp = Left(mTextLine, R - 1)
                    End If
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    NhifNo = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    NssfNo = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    PinNo = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    myEmpType = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    myTerms = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    myPay = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")

                    ECategory = mTextLine

                End If
                
                ExecuteSql ("DELETE FROM Employee WHERE EmpCode='" & EmpCode & "'")
                
                sQL = "INSERT INTO Employee (EmpCode, Surname, OtherNames, IDNo, DOB, DEmployed," & _
                    " BasicPay,Terms, Type, PinNo, NssfNo, NhifNo, ECategory)" & _
                    " VALUES('" & EmpCode & "','" & SurName & "','" & ONames & "'," & _
                    " '" & IdNo & "', '" & DOB & "', '" & DEmp & "'," & _
                    " " & myPay & ",'" & myTerms & "', '" & myEmpType & "', '" & PinNo & "', '" & NssfNo & "'," & _
                    " '" & NhifNo & "','" & ECategory & "')"
                            
            ExecuteSql (sQL)
            Dim rsMarryDept_Details As New adodb.Recordset
            Dim rsBring_ID As New adodb.Recordset
            Dim rsBringOtherDet As New adodb.Recordset
            
            Set rsBring_ID = CConnect.GetRecordSet("select * from employee where empcode='" & EmpCode & "'")
            If rsBring_ID.RecordCount > 0 Then
                Set rsBringOtherDet = CConnect.GetRecordSet("select * from cstructure")
                If rsBringOtherDet.RecordCount > 0 Then
                    Action = "IMPORTED EMPLOYEE DATA; STAFF NUMBER: " & EmpCode & "; UPDATED DEPARTMENT TO: " & rsBringOtherDet!Description
                    CConnect.ExecuteSql "DELETE FROM sEmp WHERE Employee_id='" & rsBring_ID!Employee_ID & "'"
                    CConnect.ExecuteSql "INSERT INTO sEmp (scode,lcode,empcode,employee_id,cstructure_id) VALUES('" & rsBringOtherDet!scode & "','" & rsBringOtherDet!LCode & "','" & EmpCode & "','" & rsBring_ID!Employee_ID & "','" & rsBringOtherDet!cstructure_id & "')"
                    CConnect.ExecuteSql "UPDATE employee SET cstructure_id='" & rsBringOtherDet!cstructure_id & "', term=0, dcode='" & rsBringOtherDet!Code & "',desig='" & rsBringOtherDet!Description & "' WHERE employee_id='" & rsBring_ID!Employee_ID & "'"
                End If
            End If
            
        End If
    Loop
    Close #mHandle

    MsgBox "Importation completed successfully", vbInformation

'    ReadText = arrResult
    Exit Function

errHandler:
    MsgBox err.Description '"Invalid import file format. Import not successful"

End Function
Function ImportCTypes(inFileName As String, TabDilimited As Boolean) As Variant
On Error GoTo errHandler

    Dim mTextLine As String
    Dim arrResult() As String
    Dim mHandle
    Dim i As Integer
    Dim R&
    Dim Code As String
    Dim Name As String
    Dim Comments As String
    

    If MsgBox("This will import Contacts Type Details hence updating current records. Do you wish to continue", vbInformation + vbYesNo) = vbNo Then
        Exit Function
    End If


    mHandle = FreeFile

    Open inFileName For Input As #mHandle
    i = 0
    Do While Not EOF(mHandle)
        Line Input #mHandle, mTextLine
        If Len(Trim(mTextLine)) > 0 Then
              ReDim Preserve arrResult(i)
              i = i + 1
                   ' Place the line read into array
              arrResult(UBound(arrResult)) = mTextLine
                If TabDilimited = True Then
                    R = InStr(mTextLine, vbTab)

                    Code = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Name = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Comments = mTextLine
                Else
                    R = InStr(mTextLine, ",")

                    Code = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Name = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Comments = mTextLine

                End If
                
                ExecuteSql ("DELETE FROM CTypes WHERE Code = '" & Code & "' ")
                          
                ExecuteSql ("INSERT INTO CTypes (Code,Description, Comments)" & _
                        " VALUES ('" & Code & "','" & Name & "','" & Comments & "')")
                            

        End If
    Loop
    Close #mHandle

    MsgBox "Importation completed successfully", vbInformation

'    ReadText = arrResult
    Exit Function

errHandler:
    MsgBox err.Description & ". Import not successful"


End Function

Function ImportBioTypes(inFileName As String, TabDilimited As Boolean) As Variant
On Error GoTo errHandler

    Dim mTextLine As String
    Dim arrResult() As String
    Dim mHandle
    Dim i As Integer
    Dim R&
    Dim Code As String
    Dim Name As String
    Dim Comments As String
    

    If MsgBox("This will import Bio Data Type Details hence updating current records. Do you wish to continue", vbInformation + vbYesNo) = vbNo Then
        Exit Function
    End If


    mHandle = FreeFile

    Open inFileName For Input As #mHandle
    i = 0
    Do While Not EOF(mHandle)
        Line Input #mHandle, mTextLine
        If Len(Trim(mTextLine)) > 0 Then
              ReDim Preserve arrResult(i)
              i = i + 1
                   ' Place the line read into array
              arrResult(UBound(arrResult)) = mTextLine
                If TabDilimited = True Then
                    R = InStr(mTextLine, vbTab)

                    Code = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Name = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Comments = mTextLine
                Else
                    R = InStr(mTextLine, ",")

                    Code = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Name = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Comments = mTextLine

                End If
                
                ExecuteSql ("DELETE FROM BTypes WHERE Code = '" & Code & "' ")
                          
                ExecuteSql ("INSERT INTO BTypes (Code,Description, Comments)" & _
                        " VALUES ('" & Code & "','" & Name & "','" & Comments & "')")
                            

        End If
    Loop
    Close #mHandle

    MsgBox "Importation completed successfully.", vbInformation

'    ReadText = arrResult
    Exit Function

errHandler:
    MsgBox "Invalid import format. Import not successful"

End Function
Function ImportDTypes(inFileName As String, TabDilimited As Boolean) As Variant
On Error GoTo errHandler

    Dim mTextLine As String
    Dim arrResult() As String
    Dim mHandle
    Dim i As Integer
    Dim R&
    Dim Code As String
    Dim Name As String
    Dim Comments As String
    

    If MsgBox("This will import Defined Details Type Details hence updating current records. Do you wish to continue", vbInformation + vbYesNo) = vbNo Then
        Exit Function
    End If


    mHandle = FreeFile

    Open inFileName For Input As #mHandle
    i = 0
    Do While Not EOF(mHandle)
        Line Input #mHandle, mTextLine
        If Len(Trim(mTextLine)) > 0 Then
              ReDim Preserve arrResult(i)
              i = i + 1
                   ' Place the line read into array
              arrResult(UBound(arrResult)) = mTextLine
                If TabDilimited = True Then
                    R = InStr(mTextLine, vbTab)

                    Code = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Name = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Comments = mTextLine
                Else
                    R = InStr(mTextLine, ",")

                    Code = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Name = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Comments = mTextLine

                End If
                
                ExecuteSql ("DELETE FROM DTypes WHERE Code = '" & Code & "' ")
                          
                ExecuteSql ("INSERT INTO DTypes (Code,Description, Comments)" & _
                        " VALUES ('" & Code & "','" & Name & "','" & Comments & "')")
                            

        End If
    Loop
    Close #mHandle

    MsgBox "Importation completed successfully", vbInformation

'    ReadText = arrResult
    Exit Function

errHandler:
    MsgBox err.Description & ". Import not successful"


End Function
Function ImportEmpCTypes(inFileName As String, TabDilimited As Boolean) As Variant
On Error GoTo errHandler

    Dim mTextLine As String
    Dim arrResult() As String
    Dim mHandle
    Dim i As Integer
    Dim R&
    Dim EmpCode As String
    Dim Code As String
    Dim Description As String
    

    If MsgBox("This will import Employee Contacts Details hence updating current records. Do you wish to continue", vbInformation + vbYesNo) = vbNo Then
        Exit Function
    End If


    mHandle = FreeFile

    Open inFileName For Input As #mHandle
    i = 0
    Do While Not EOF(mHandle)
        Line Input #mHandle, mTextLine
        If Len(Trim(mTextLine)) > 0 Then
              ReDim Preserve arrResult(i)
              i = i + 1
                   ' Place the line read into array
              arrResult(UBound(arrResult)) = mTextLine
                If TabDilimited = True Then
                    R = InStr(mTextLine, vbTab)
                    
                    EmpCode = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)

                    Code = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Name = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Comments = mTextLine
                Else
                    R = InStr(mTextLine, ",")
                    
                    EmpCode = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")

                    Code = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Name = mTextLine

                End If
                
                ExecuteSql ("DELETE FROM Contacts WHERE  employee_id =" & getEmployeeID(EmpCode) & " AND Code = '" & Code & "' ")
                          
                ExecuteSql ("INSERT INTO Contacts (employee_id,Code,Contact)" & _
                        " VALUES ('" & getEmployeeID(EmpCode) & "','" & Code & "','" & Name & "')")
                            

        End If
    Loop
    Close #mHandle

    MsgBox "Importation completed successfully", vbInformation

'    ReadText = arrResult
    Exit Function

errHandler:
    MsgBox err.Description & ". Import not successful"


End Function
Function ImportEmpBTypes(inFileName As String, TabDilimited As Boolean) As Variant
On Error GoTo errHandler

    Dim mTextLine As String
    Dim arrResult() As String
    Dim mHandle
    Dim i As Integer
    Dim R&
    Dim EmpCode As String
    Dim Code As String
    Dim Description As String
    

    If MsgBox("This will import Employee Bio data hence updating current records. Do you wish to continue", vbInformation + vbYesNo) = vbNo Then
        Exit Function
    End If


    mHandle = FreeFile

    Open inFileName For Input As #mHandle
    i = 0
    Do While Not EOF(mHandle)
        Line Input #mHandle, mTextLine
        If Len(Trim(mTextLine)) > 0 Then
              ReDim Preserve arrResult(i)
              i = i + 1
                   ' Place the line read into array
              arrResult(UBound(arrResult)) = mTextLine
                If TabDilimited = True Then
                    R = InStr(mTextLine, vbTab)
                    
                    EmpCode = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)

                    Code = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Name = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Comments = mTextLine
                Else
                    R = InStr(mTextLine, ",")
                    
                    EmpCode = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")

                    Code = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Name = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Comments = mTextLine

                End If
                
                ExecuteSql ("DELETE FROM Bio WHERE  employee_id ='" & getEmployeeID(EmpCode) & "' AND Code = '" & Code & "' ")
                          
                ExecuteSql ("INSERT INTO Bio (employee_id,Code,BioData,comments)" & _
                        " VALUES (" & getEmployeeID(EmpCode) & ",'" & Code & "','" & Name & "','" & Comments & "')")
                            

        End If
    Loop
    Close #mHandle

    MsgBox "Importation completed successfully", vbInformation

'    ReadText = arrResult
    Exit Function

errHandler:
    MsgBox err.Description & ". Import not successful"


End Function

Function ImportEmpDTypes(inFileName As String, TabDilimited As Boolean) As Variant
On Error GoTo errHandler

    Dim mTextLine As String
    Dim arrResult() As String
    Dim mHandle
    Dim i As Integer
    Dim R&
    Dim EmpCode As String
    Dim Code As String
    Dim Description As String
    Dim Comments As String
    

    If MsgBox("This will import Employee Defined Details hence updating current records. Do you wish to continue", vbInformation + vbYesNo) = vbNo Then
        Exit Function
    End If


    mHandle = FreeFile

    Open inFileName For Input As #mHandle
    i = 0
    Do While Not EOF(mHandle)
        Line Input #mHandle, mTextLine
        If Len(Trim(mTextLine)) > 0 Then
              ReDim Preserve arrResult(i)
              i = i + 1
                   ' Place the line read into array
              arrResult(UBound(arrResult)) = mTextLine
                If TabDilimited = True Then
                    R = InStr(mTextLine, vbTab)
                    
                    EmpCode = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)

                    Code = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Description = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Comments = mTextLine
                Else
                    R = InStr(mTextLine, ",")
                    
                    EmpCode = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")

                    Code = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Description = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Comments = mTextLine

                End If
                
                ExecuteSql ("DELETE FROM DDetails WHERE  employee_id='" & getEmployeeID(EmpCode) & "' AND Code = '" & Code & "' ")
                          
                ExecuteSql ("INSERT INTO DDetails (employee_id,Code,Detail,Comments)" & _
                        " VALUES (" & getEmployeeID(EmpCode) & ",'" & Code & "','" & Description & "','" & Comments & "')")
                            

        End If
    Loop
    Close #mHandle

    MsgBox "Importation completed successfully", vbInformation

'    ReadText = arrResult
    Exit Function

errHandler:
    MsgBox err.Description & ". Import not successful"


End Function

Function ImportEmpBasic(inFileName As String, TabDilimited As Boolean) As Variant
On Error GoTo errHandler

    Dim mTextLine As String
    Dim arrResult() As String
    Dim mHandle
    Dim i As Integer
    Dim R&
    Dim EmpCode As String
    Dim BasicPay As Currency
       

    If MsgBox("This will import Employee basic pay/daily rate hence updating current records. Do you wish to continue", vbInformation + vbYesNo) = vbNo Then
        Exit Function
    End If


    mHandle = FreeFile

    Open inFileName For Input As #mHandle
    i = 0
    Do While Not EOF(mHandle)
        Line Input #mHandle, mTextLine
        If Len(Trim(mTextLine)) > 0 Then
              ReDim Preserve arrResult(i)
              i = i + 1
                   ' Place the line read into array
              arrResult(UBound(arrResult)) = mTextLine
                If TabDilimited = True Then
                    R = InStr(mTextLine, vbTab)
                    
                    EmpCode = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    BasicPay = mTextLine
                    
                Else
                    R = InStr(mTextLine, ",")
                    
                    EmpCode = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    BasicPay = mTextLine

                End If
                
                ExecuteSql ("UPDATE Employee SET BasicPay = " & BasicPay & " WHERE  EmpCode = '" & EmpCode & "'")
                          
        End If
    Loop
    Close #mHandle

    MsgBox "Importation completed successfully", vbInformation

'    ReadText = arrResult
    Exit Function

errHandler:
    MsgBox err.Description & ". Import not successful"


End Function

Function ImportKins(inFileName As String, TabDilimited As Boolean) As Variant
On Error GoTo errHandler

    Dim mTextLine As String
    Dim arrResult() As String
    Dim mHandle
    Dim i As Integer
    Dim R&
    Dim EmpCode As String
    Dim Code As String
    Dim SurName As String
    Dim ONames As String
    Dim IdNo As String
    Dim DOB As Date
    Dim Relation As String
    Dim Occupation As String
    Dim Address As String
    Dim TelNo As String
    Dim OffTelNo As String
    Dim EMail As String
    Dim CellNo As String
    Dim Signed As String
    Dim SDate As Date
    Dim Comments As String

    If MsgBox("This will import Employee Next of Kin Details hence updating current records. Do you wish to continue", vbInformation + vbYesNo) = vbNo Then
        Exit Function
    End If


    mHandle = FreeFile

    Open inFileName For Input As #mHandle
    i = 0
    Do While Not EOF(mHandle)
        Line Input #mHandle, mTextLine
        If Len(Trim(mTextLine)) > 0 Then
              ReDim Preserve arrResult(i)
              i = i + 1
                   ' Place the line read into array
              arrResult(UBound(arrResult)) = mTextLine
                If TabDilimited = True Then
                    R = InStr(mTextLine, vbTab)

                    EmpCode = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Code = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    SurName = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    ONames = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    IdNo = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Relation = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    If Not IsDate(Left(mTextLine, R - 1)) Then
                        DOB = Date
                    Else
                        DOB = Left(mTextLine, R - 1)
                    End If
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Occupation = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    TelNo = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    OffTelNo = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    CellNo = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    EMail = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Address = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Signed = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    If Not IsDate(Left(mTextLine, R - 1)) Then
                        SDate = Date
                    Else
                        SDate = Left(mTextLine, R - 1)
                    End If
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)

                    Comments = mTextLine
                Else
                    R = InStr(mTextLine, ",")

                    EmpCode = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Code = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    SurName = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    ONames = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    IdNo = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Relation = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    If Not IsDate(Left(mTextLine, R - 1)) Then
                        DOB = Date
                    Else
                        DOB = Left(mTextLine, R - 1)
                    End If
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Occupation = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    TelNo = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    OffTelNo = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    CellNo = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    EMail = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Address = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Signed = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    If Not IsDate(Left(mTextLine, R - 1)) Then
                        SDate = Date
                    Else
                        SDate = Left(mTextLine, R - 1)
                    End If
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")

                    Comments = mTextLine

                End If
                
                ExecuteSql ("DELETE FROM Kin WHERE employee_id= " & getEmployeeID(EmpCode) & " AND Code='" & Code & "'")
                
                sQL = "INSERT INTO Kin (employee_id, Code, SurName, OtherNames, IDNo, Relation, DOB, Occupation, HTelNo, OffTelNo, MNo, EMail, Address, Signed, Comments)" & _
                    " VALUES(" & getEmployeeID(EmpCode) & ",'" & Code & "','" & SurName & "','" & ONames & "','" & IdNo & "','" & Relation & "','" & DOB & "'," & _
                    "'" & Occupation & "','" & TelNo & "','" & OffTelNo & "','" & CellNo & "','" & EMail & "','" & Address & "','" & Signed & "','" & Comments & "')"
                            
            ExecuteSql (sQL)

        End If
    Loop
    Close #mHandle

    MsgBox "Importation completed successfully", vbInformation

'    ReadText = arrResult
    Exit Function

errHandler:
    MsgBox err.Description & ". Import not successful"


End Function

Function ImportRefeeres(inFileName As String, TabDilimited As Boolean) As Variant
On Error GoTo errHandler

    Dim mTextLine As String
    Dim arrResult() As String
    Dim mHandle
    Dim i As Integer
    Dim R&
    Dim EmpCode As String
    Dim Code As String
    Dim Names As String
    Dim IdNo As String
    Dim Address As String
    Dim EMail As String
    Dim CellNo As String
    Dim Comments As String

    If MsgBox("This will import Employee Refeeres Details hence updating current records. Do you wish to continue", vbInformation + vbYesNo) = vbNo Then
        Exit Function
    End If


    mHandle = FreeFile

    Open inFileName For Input As #mHandle
    i = 0
    Do While Not EOF(mHandle)
        Line Input #mHandle, mTextLine
        If Len(Trim(mTextLine)) > 0 Then
              ReDim Preserve arrResult(i)
              i = i + 1
                   ' Place the line read into array
              arrResult(UBound(arrResult)) = mTextLine
                If TabDilimited = True Then
                    R = InStr(mTextLine, vbTab)

                    EmpCode = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Code = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Names = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    IdNo = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    CellNo = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    EMail = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)
                    
                    R = InStr(mTextLine, vbTab)
                    
                    Address = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)

                    Comments = mTextLine
                Else
                    R = InStr(mTextLine, ",")

                    EmpCode = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Code = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Names = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    IdNo = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Address = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    CellNo = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    EMail = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")

                    Comments = mTextLine

                End If
                
                ExecuteSql ("DELETE FROM Ref WHERE employee_id=" & getEmployeeID(EmpCode) & " AND Code='" & Code & "'")
                
                sQL = "INSERT INTO Ref (employee_id, Code, Names, IDNo,MNo,EMail, Address, Comments)" & _
                " VALUES(" & getEmployeeID(EmpCode) & ",'" & Code & "','" & Names & "','" & IdNo & "'," & _
                "'" & CellNo & "','" & EMail & "','" & Address & "','" & Comments & "')"
                            
            ExecuteSql (sQL)

        End If
    Loop
    Close #mHandle

    MsgBox "Importation completed successfully", vbInformation

'    ReadText = arrResult
    Exit Function

errHandler:
    MsgBox err.Description & ". Import not successful"


End Function
Function ImportEDetails(inFileName As String, TabDilimited As Boolean) As Variant
On Error GoTo errHandler

    Dim mTextLine As String
    Dim arrResult() As String
    Dim mHandle
    Dim i As Integer
    Dim R&
    Dim EmpCode As String
    Dim Code As String
    Dim Course As String
    Dim cFrom As Date
    Dim cTo As Date
    Dim ELevel As String
    Dim Award As String
    Dim Comments As String

    If MsgBox("This will import Employee Education Details hence updating current records. Do you wish to continue", vbInformation + vbYesNo) = vbNo Then
        Exit Function
    End If


    mHandle = FreeFile

    Open inFileName For Input As #mHandle
    i = 0
    Do While Not EOF(mHandle)
        Line Input #mHandle, mTextLine
        If Len(Trim(mTextLine)) > 0 Then
              ReDim Preserve arrResult(i)
              i = i + 1
                   ' Place the line read into array
              arrResult(UBound(arrResult)) = mTextLine
                If TabDilimited = True Then
                    R = InStr(mTextLine, vbTab)

                    EmpCode = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Code = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Course = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    If Not IsDate(Left(mTextLine, R - 1)) Then
                        cFrom = Date
                    Else
                        cFrom = Left(mTextLine, R - 1)
                    End If
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    If Not IsDate(Left(mTextLine, R - 1)) Then
                        cTo = Date
                    Else
                        cTo = Left(mTextLine, R - 1)
                    End If
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    ELevel = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Award = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)

                    Comments = mTextLine
                    
                    R = InStr(mTextLine, vbTab)
                    
                    Institution = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)

                    Comments = mTextLine
                Else
                    R = InStr(mTextLine, ",")

                    EmpCode = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Code = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Course = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    If Not IsDate(Left(mTextLine, R - 1)) Then
                        cFrom = Date
                    Else
                        cFrom = Left(mTextLine, R - 1)
                    End If
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    If Not IsDate(Left(mTextLine, R - 1)) Then
                        cTo = Date
                    Else
                        cTo = Left(mTextLine, R - 1)
                    End If
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    ELevel = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Award = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")

                    Comments = mTextLine

                End If
                
                ExecuteSql ("DELETE FROM Edu WHERE employee_id= " & getEmployeeID(EmpCode) & " AND Code='" & Code & "'")
                
                sQL = "INSERT INTO Edu (employee_id, Code, Course, ELevel, CFrom, CTo, Award, Comments)" & _
                        " VALUES(" & getEmployeeID(EmpCode) & ",'" & Code & "','" & Course & "','" & ELevel & "'," & _
                        "'" & cFrom & "','" & cTo & "','" & Award & "','" & Comments & "')"
                            
            ExecuteSql (sQL)

        End If
    Loop
    Close #mHandle

    MsgBox "Importation completed successfully", vbInformation

'    ReadText = arrResult
    Exit Function

errHandler:
    MsgBox err.Description & ". Import not successful"


End Function

Function ImportEmploy(inFileName As String, TabDilimited As Boolean) As Variant
On Error GoTo errHandler

    Dim mTextLine As String
    Dim arrResult() As String
    Dim mHandle
    Dim i As Integer
    Dim R&
    Dim EmpCode As String
    Dim Code As String
    Dim Employer As String
    Dim cFrom As Date
    Dim cTo As Date
    Dim Reasons As String
    Dim Designation As String
    Dim Super As String
    Dim Salary As Currency
    Dim Comments As String

    If MsgBox("This will import Employee Education Details hence updating current records. Do you wish to continue", vbInformation + vbYesNo) = vbNo Then
        Exit Function
    End If


    mHandle = FreeFile

    Open inFileName For Input As #mHandle
    i = 0
    Do While Not EOF(mHandle)
        Line Input #mHandle, mTextLine
        If Len(Trim(mTextLine)) > 0 Then
              ReDim Preserve arrResult(i)
              i = i + 1
                   ' Place the line read into array
              arrResult(UBound(arrResult)) = mTextLine
                If TabDilimited = True Then
                    R = InStr(mTextLine, vbTab)

                    EmpCode = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Code = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Employer = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    If Not IsDate(Left(mTextLine, R - 1)) Then
                        cFrom = Date
                    Else
                        cFrom = Left(mTextLine, R - 1)
                    End If
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    If Not IsDate(Left(mTextLine, R - 1)) Then
                        cTo = Date
                    Else
                        cTo = Left(mTextLine, R - 1)
                    End If
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Reasons = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Designation = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Super = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Salary = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)

                    Comments = mTextLine
                Else
                    R = InStr(mTextLine, ",")

                    EmpCode = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Code = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Employer = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    If Not IsDate(Left(mTextLine, R - 1)) Then
                        cFrom = Date
                    Else
                        cFrom = Left(mTextLine, R - 1)
                    End If
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    If Not IsDate(Left(mTextLine, R - 1)) Then
                        cTo = Date
                    Else
                        cTo = Left(mTextLine, R - 1)
                    End If
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Reasons = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Designation = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Super = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Salary = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    Comments = mTextLine

                End If
                
                ExecuteSql ("DELETE FROM Employment WHERE employee_id = " & getEmployeeID(EmpCode) & " AND Code='" & Code & "'")
                
                sQL = "INSERT INTO Employment (employee_id, Code, Employer, CFrom, CTo, Reasons, Desig, Super, Salary, Comments)" & _
                        " VALUES(" & getEmployeeID(EmpCode) & ",'" & Code & "','" & Employer & "'," & _
                        "'" & cFrom & "','" & cTo & "','" & Reasons & "','" & Designation & "','" & Super & "'," & Salary & ",'" & Comments & "')"
                            
                ExecuteSql (sQL)

        End If
    Loop
    Close #mHandle

    MsgBox "Importation completed successfully", vbInformation

'    ReadText = arrResult
    Exit Function

errHandler:
    MsgBox err.Description & ". Import not successful"


End Function


Function ImportPDetails(inFileName As String, TabDilimited As Boolean) As Variant
On Error GoTo errHandler

    Dim mTextLine As String
    Dim arrResult() As String
    Dim mHandle
    Dim i As Integer
    Dim R&
    Dim EmpCode As String
    Dim Code As String
    Dim Course As String
    Dim cFrom As Date
    Dim cTo As Date
    Dim ELevel As String
    Dim Award As String
    Dim Comments As String

    If MsgBox("This will import Employee Professional Details hence updating current records. Do you wish to continue", vbInformation + vbYesNo) = vbNo Then
        Exit Function
    End If


    mHandle = FreeFile

    Open inFileName For Input As #mHandle
    i = 0
    Do While Not EOF(mHandle)
        Line Input #mHandle, mTextLine
        If Len(Trim(mTextLine)) > 0 Then
              ReDim Preserve arrResult(i)
              i = i + 1
                   ' Place the line read into array
              arrResult(UBound(arrResult)) = mTextLine
                If TabDilimited = True Then
                    R = InStr(mTextLine, vbTab)

                    EmpCode = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Code = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Course = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    If Not IsDate(Left(mTextLine, R - 1)) Then
                        cFrom = Date
                    Else
                        cFrom = Left(mTextLine, R - 1)
                    End If
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    If Not IsDate(Left(mTextLine, R - 1)) Then
                        cTo = Date
                    Else
                        cTo = Left(mTextLine, R - 1)
                    End If
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    ELevel = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Award = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)

                    Comments = mTextLine
                Else
                    R = InStr(mTextLine, ",")

                    EmpCode = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Code = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Course = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    If Not IsDate(Left(mTextLine, R - 1)) Then
                        cFrom = Date
                    Else
                        cFrom = Left(mTextLine, R - 1)
                    End If
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    If Not IsDate(Left(mTextLine, R - 1)) Then
                        cTo = Date
                    Else
                        cTo = Left(mTextLine, R - 1)
                    End If
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    ELevel = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Award = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")

                    Comments = mTextLine

                End If
                
                ExecuteSql ("DELETE FROM Prof WHERE employee_id='" & getEmployeeID(EmpCode) & "'AND Code='" & Code & "'")
                
                sQL = "INSERT INTO Prof (employee_id, Code, Course, ELevel, CFrom, CTo, Award, Comments)" & _
                        " VALUES(" & getEmployeeID(EmpCode) & ",'" & Code & "','" & Course & "','" & ELevel & "'," & _
                        "'" & cFrom & "','" & cTo & "','" & Award & "','" & Comments & "')"
                            
            ExecuteSql (sQL)

        End If
    Loop
    Close #mHandle

    MsgBox "Importation completed successfully", vbInformation

'    ReadText = arrResult
    Exit Function

errHandler:
    MsgBox err.Description & ". Import not successful"


End Function

Function ImportContracts(inFileName As String, TabDilimited As Boolean) As Variant
On Error GoTo errHandler

    Dim mTextLine As String
    Dim arrResult() As String
    Dim mHandle
    Dim i As Integer
    Dim R&
    Dim EmpCode As String
    Dim Code As String
    Dim Description As String
    Dim cFrom As Date
    Dim cTo As Date
    Dim Ref As String
    Dim Comments As String

    If MsgBox("This will import Employee Contracts Details hence updating current records. Do you wish to continue", vbInformation + vbYesNo) = vbNo Then
        Exit Function
    End If


    mHandle = FreeFile

    Open inFileName For Input As #mHandle
    i = 0
    Do While Not EOF(mHandle)
        Line Input #mHandle, mTextLine
        If Len(Trim(mTextLine)) > 0 Then
              ReDim Preserve arrResult(i)
              i = i + 1
                   ' Place the line read into array
              arrResult(UBound(arrResult)) = mTextLine
                If TabDilimited = True Then
                    R = InStr(mTextLine, vbTab)

                    EmpCode = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Code = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Description = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    If Not IsDate(Left(mTextLine, R - 1)) Then
                        cFrom = Date
                    Else
                        cFrom = Left(mTextLine, R - 1)
                    End If
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    If Not IsDate(Left(mTextLine, R - 1)) Then
                        cTo = Date
                    Else
                        cTo = Left(mTextLine, R - 1)
                    End If
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)
                    
                    Ref = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, vbTab)

                    Comments = mTextLine
                Else
                    R = InStr(mTextLine, ",")

                    EmpCode = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Code = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Description = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    If Not IsDate(Left(mTextLine, R - 1)) Then
                        cFrom = Date
                    Else
                        cFrom = Left(mTextLine, R - 1)
                    End If
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    If Not IsDate(Left(mTextLine, R - 1)) Then
                        cTo = Date
                    Else
                        cTo = Left(mTextLine, R - 1)
                    End If
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")
                    
                    Ref = Left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)

                    R = InStr(mTextLine, ",")

                    Comments = mTextLine

                End If
                
                ExecuteSql ("DELETE FROM Contracts WHERE employee_id = " & getEmployeeID(EmpCode) & " AND Code='" & Code & "'")
                
                sQL = "INSERT INTO Contracts (employee_id, Code, Description, CFrom, CTo, Ref, Comments)" & _
                        " VALUES(" & getEmployeeID(EmpCode) & ",'" & Code & "','" & Description & "'," & _
                        "'" & cFrom & "','" & cTo & "','" & Ref & "','" & Comments & "')"
                            
            ExecuteSql (sQL)

        End If
    Loop
    Close #mHandle

    MsgBox "Importation completed successfully", vbInformation

    Exit Function

errHandler:
    MsgBox err.Description & ". Import not successful"


End Function

Public Function getEmployeeID(EmpCode As String) As Integer
    Dim rec_t As New adodb.Recordset
    On Error GoTo errHandler
    If EmpCode <> "" And Not IsNull(EmpCode) Then
        Set rec_t = GetRecordSet("SELECT employee_id FROM employee WHERE empcode = '" & EmpCode & "'")
        While rec_t.EOF = False
            getEmployeeID = CInt(rec_t!Employee_ID)
            rec_t.MoveNext
        Wend
    End If
    Exit Function
errHandler:
    MsgBox err.Description, vbInformation
End Function

Public Sub XMLConnection()
    Dim XMLDocument As DOMDocument
    Dim XMLParentNode As IXMLDOMNode
    Dim XMLPath As String
    Dim strConString As String
    Dim strServerName As String
    Dim strDBName As String
    Dim strUserID As String
    Dim strPassword As String
    Dim blnWinAuth As Boolean
    
    On Error GoTo ErrorHandler
    Set Cnn = New adodb.Connection
        
    XMLPath = GetSetting("HRMIS", "XML", "PATH") & "HRMISConn.xml"

    Set XMLDocument = New DOMDocument
    'LOADING THE XML FILE
    XMLDocument.Load XMLPath
    'CAUTION: WHETHER THE XML FILE IS REALLY A VALID XML FILE
    If XMLDocument.documentElement Is Nothing Then
        MsgBox "The specified XML File was not loaded successful" & vbCrLf & "Please confirm that the specified file a valid XML File"
        Exit Sub
    Else
        'CAUTION: WHETHER THE XML FILE IS REALLY A VALID HRMIS DATABASE CONFIGURATION XML FILE
        Set XMLParentNode = XMLDocument.selectSingleNode("DatabaseConnection")
        If XMLParentNode Is Nothing Then
            MsgBox "The specified XML File is not a valid HRMIS Database Configuration XML File" & vbCrLf & "Please load a valid HRMIS Database Configuration XML File", vbExclamation
            Exit Sub
        Else
            'NOW LOADING THE VALUES
            Set xmlreader = New HRMIS_XML.hrmisXML
            strServerName = GetNodeValue(XMLParentNode, "ServerName")
            strDBName = GetNodeValue(XMLParentNode, "DatabaseName")
            strUserID = xmlreader.Decode64(GetNodeValue(XMLParentNode, "UserID"))
            strPassword = xmlreader.Decode64(GetNodeValue(XMLParentNode, "Password"))
'            strUserID = GetNodeValue(XMLParentNode, "UserID")
'            strPassword = GetNodeValue(XMLParentNode, "Password")
        End If
        'TESTING THE CONNECTION
        strConString = "Provider=SQLOLEDB.1;Data Source=" & strServerName & ";Initial Catalog=" & strDBName & ";User ID=" & strUserID & ";Password=" & strPassword & ";Persist Security Info=False"
        Cnn.ConnectionString = strConString
        Cnn.Open
        If Cnn.State = adStateOpen Then
            'store the parameters for the purpose of the reports
            Set con = Cnn
            With ConParams
                .DataSource = strServerName
                .InitialCatalog = strDBName
                .UserID = strUserID
                .Password = strPassword
                .TrustedConnection = False
            End With
    
            'Set the HRCORE connection
            retVal = HRCon.SetConnection(Cnn)
            If retVal = 0 Then
                retVal2 = HRMSECCon.SetConnection(Cnn)
                If retVal2 <> 0 Then
                    MsgBox "Database Connection Could not be synchronized with HR Security", vbExclamation, "Security Connection"
                    err.Raise 430
                    'End
                End If
            Else
                MsgBox "Database connection could not be set for the HR CORE Engine", vbExclamation, "HR CORE Engine"
                End
            End If
        Else
            MsgBox "Configuration Information was not found", vbInformation, TITLES
        End If
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while checking configuration information" & vbNewLine & err.Description, vbInformation, TITLES
End Sub

Private Function GetNodeValue(ByVal XMLParentNode As IXMLDOMNode, strNodeName As String, Optional strDefautlValue As String = "") As String
    On Error GoTo ErrorHandler
    
    Dim XMLChildNode As IXMLDOMNode
    
    'SEARCHING FOR A MATCHING DESCENDANT
    Set XMLChildNode = XMLParentNode.selectSingleNode(".//" & strNodeName)
    If XMLChildNode Is Nothing Then
        GetNodeValue = strDefautlValue
    Else
        GetNodeValue = XMLChildNode.Text
    End If
    Exit Function
ErrorHandler:
    MsgBox "An Error has occurred while attempting to search for a matching node descendant in an XML File" & vbCrLf & err.Description, vbExclamation
End Function
