Attribute VB_Name = "Decla"
Option Explicit
'These variables are used to maintain the height of the screens(forms)
Public dFormat As String
Public tvwMainheight As Long
Public tvwMainTop As Long
Public lvwEmpHeight As Long
Public myreport As String
Public Rcurrency As ADODB.Recordset
Public Const curSql As String = "Select * from tblCurrency"
Public ReportType As String
Public ReportSchemaName As String
Public ChangeHeight As Boolean
Public ReportHeading As Boolean 'used to hide or unhide the employee Tab
'*********************************************************************
''''''''''''
'
Public xmlreader As HRMIS_XML.hrmisXML
Public PIFcats As clsPIFs

'************************************
Public gAccessRightTypeId As Integer
Public gAccessRightName As String
Public gAccessRightClassIds As String
'*********************************

Public CompanyId As Long
''''''''''''''''''''''''''''''''

Public outs As HRCORE.OrganizationUnitTypes
Public OUnits As HRCORE.OrganizationUnits
Public selOU As HRCORE.OrganizationUnit
Public company As New HRCORE.CompanyDetails
Public emp As HRCORE.Employee
Public Emps As HRCORE.Employees

Public OUs As HRCORE.OrganizationUnits
  



Public EmpCats As HRCORE.EmployeeCategories
Public empTerms As HRCORE.EmploymentTerms
Public empTribes As HRCORE.Tribes
Public empNationalities As HRCORE.Nationalities
Public empReligions As HRCORE.Religions
Public empPositions As HRCORE.JobPositions
Public empCurrencies As HRCORE.Currencies
Public empCountries As HRCORE.Countries
Public empLocations As HRCORE.Locations
Public empProjects As HRCORE.Programmes
Public empStaffCategories As HRCORE.CSSSCategories

Public selCSSSCategory As HRCORE.CSSSCategory
Public selCountry As HRCORE.Country
Public selEmpCurrency As HRCORE.CCurrency
Public selEmpCategory As HRCORE.EmployeeCategory
Public selLocation As HRCORE.Location
Public selCurrency As CCurrency

'Next Of Kins
'Private newNextOfKins As New HRCORE.NextOfKins
Public TempNextOfKins As HRCORE.NextOfKins
Public FilteredNextOfKins As HRCORE.NextOfKins
Public selNextOfKin As HRCORE.NextOfKin


Public pJDFields As HRCORE.JDCategories
Public TopLevelJDFields As HRCORE.JDCategories
Public selJDField As HRCORE.JDCategory
Public pEmpJDs As HRCORE.EmployeeJDs
Public FilteredEmpJDs As HRCORE.EmployeeJDs
Public selEmpJD As HRCORE.EmployeeJD
Public TempEmpJDs As HRCORE.EmployeeJDs

Public ParentChangedFromCode As Boolean
'====== END OF HRCORE DECLARATIONS =========

Public objProgrammes As HRCORE.Programmes
Public objFundCodes As HRCORE.FundCodes
Public objProgrammeFundings As HRCORE.ProgrammeFundings
Public objEmployeeProgrammeFundings As HRCORE.EmployeeProgrammes
Public SelectedEmployeeProgrammeFunding As HRCORE.EmployeeProgramme
Public TempObjEmployeeProgrammeFundings As HRCORE.EmployeeProgrammes
Public blnDisplayEmployeeProgrammeFundingInfo As Boolean

Public objEmployeeBankAccounts As EmployeeBankAccounts2
Public NewAcct As Boolean
Public LastSecID As Long
Public GenerateID As Boolean
Public EnterDOB As Boolean
Public EnterDEmp As Boolean
Public MStruc As String        'holds the Main Structure from STypes

'For Tracking Employee Changes
Public OldEmpInfo As Employee
Public NewEmpInfo As Employee



''''''''''''''''''''''''''''''''''''''''''''

Public pBanks As Banks
Public pBankBranches As BankBranches
''Public pEmployeeBankAccounts As EmployeeBankAccounts2

Public selBank As Bank
Public selBankBranch As bankbranch
Public selEmployeeBankAccount As EmployeeBankAccount2
Public empBankAccounts As EmployeeBankAccounts2

Public blnNewEntry As Boolean
Public ChangedFromCode As Boolean
Public bankindex As Integer





''''''''''''

Public gEmployeeID As Long
Public gMinEmployeeID As Long
Public gMaxEmployeeID As Long
Public companyDetail As HRCORE.CompanyDetails
Public PSave As Boolean
Public branchReports As Boolean
Public empCat As String
Public strQ As String
Public strName As String
Public strNamePart As String
Public strcode As String
Public strAccessiblePayrollids As String
Public strValue As String
Public strDatePart As String
Public strID As String
Public sql As String
Public TheReport As String
Public strBranchName As String
Public strBranchID As Integer
Public strBankName As String
Public EmployeeIsInEditMode As Boolean    'will indicate that an employee is being edited
Public new_Record As Boolean    'flag indicating that a new record is being inserted
Public rstAudit As New ADODB.Recordset
Public Action As String
Public popupText As String
Public Picpath As String
Public KeyAscii As KeyCodeConstants
Public RFilter As String
Public Recruit As String
Public APath As String
Public ServerName As String
Public catalog As String
Public passwd As String
Public UserID As String
Public Dis As String
Public AuditTrail As Boolean
Public DSource As String
Public PUnApp As Long
Public AppGroup As String
Public OData As Boolean
Public ViewSal As Boolean
'Public ECode As String
Public OminisDB As Boolean
Public PRate As Double
Public FRetire As Double
Public MRetire As Double
Public ALCode As String
Public ByPass As Boolean
Public AChange  As String
Public MyColor As Long
Public Godays As Double
Public PromptDate As Date
Public PromptSave As Boolean
'Public RType As String
Public EmpGroup As String
Public CPass As String
Public IDiv As String
Public DPass As String
Public myfile As String
Public myYear As Double
Public CloseDate As Date
Public Sel As String
Public CurrentEvent As String
Public MyR As Double
Public TheLoadedForm As Object
Public FViewOnly As Boolean
Public Machine_User_Settings As Variant
'Public a As CRAXDDRT.Application
Public r As CRAXDDRT.Report
Public a As CRAXDRT.Application
Public li As ListItem
Public Cnn As Connection
Public cnnPayData As Connection
Public rs As Recordset
Public rs1 As Recordset
Public rs2 As Recordset
Public rs3 As Recordset
Public rs4 As Recordset
Public rs5 As Recordset
Public rs6 As Recordset
Public rst1 As Recordset
Public rsGlob As Recordset
Public rsGlob2 As Recordset
Public rsGenOpt As Recordset
Public CConnect As PDR.Connect
Public oSmart As New PDR.SmartForm
Public rsMySecurity As Recordset
Public rsMySec As Recordset
Public rsMyStruc As Recordset
Public SaveNew As Boolean       'flag to indicate that a new record is being inserted
Public resp As String
Public CurrentUser As String
Public Current_Logged_User As String
Public Current_GroupID As String
Public CGroup As String
Public Const SubSystem = "HRBase"
Public Const Cfmt = "###,###,###,###,##0.00;(##0.00)"
Public Const Nfmt = "##############0"
Public Const Dfmt = "yyyy-mm-dd"
'Public Dfmt As String
Public strReport As String 'Set for sending formula to crystal reports
Public mySQL As String
Public FormIsLoading As Boolean
Public FLoading As Boolean

Public SelectedBank As String
Public GroupRight As String
Public sDeptCode As String
Public wasThere As Boolean
Public connection_string As String
'Public CConnect As New Connect
Public con As ADODB.Connection
Public myEmployeeBankAccounts As EmployeeBankAccounts2
Public Type GroupType
    SubDepartment As Boolean
    Department As Boolean
    SuperUser As Boolean
    Admin As Boolean
    SubAdmin As Boolean
End Type

Private UserGroup As GroupType
Public maxCatAccess As String
Public deptConstraint As String

'========== HRMSEC DECLARATIONS ========
Public currUser As HRMSEC.CUser
Public gUser As HRMSEC.CUser
Public HRMSECCon As HRMSEC.CConnect
Public photoisactive As Boolean
'========= END OF HRMSEC DECLARATIONS ======

'For MultiCompany
Public MultiCo As HRMSEC.MultiCompany 'MultiCompany Builder

'========== HRCORE DECLARATIONS ==========
Public HRCon As New HRCORE.CConnect
Public Const TITLES As String = "Personnel Director"
Public AllEmployees As HRCORE.Employees
Public AllEmployeesPhotos As EmployeesPhotos
Public SelectedEmployee As HRCORE.Employee
Public SelectedEmployeePhoto As EmployeePhoto


Public TempNextOfKin As HRCORE.NextOfKin    'will be used to transfer data when setting Guardian Info
'========== END OF HRCORE DECLARATIONS =========

Public Type ConnectionParameters
    DataSource As String
    InitialCatalog As String
    UserID As String
    Password As String
    TrustedConnection As Boolean
End Type

Public ConParams As ConnectionParameters
'============== HRCORE FUNCTIONS ===============

Sub Main()
    On Error GoTo ErrHandler
      
    Set CConnect = New PDR.Connect
    Set HRMSECCon = New HRMSEC.CConnect
    Set gUser = New HRMSEC.CUser

     
     ''''''''''''''''''''
    'Multi-Company Check
    Set MultiCo = New MultiCompany
    If (MultiCo.IsMultiCompany) = -1 Then Exit Sub 'Only -1 if the xml file doesnt exist or contains invalid info
    
    'Create the connection
    CConnect.APath = App.Path
    CConnect.XMLConnection
   
 
    Set companyDetail = New HRCORE.CompanyDetails
    companyDetail.LoadCompanyDetails
    Dim xxxx As String
    xxxx = companyDetail.CompanyName
    
    
    
   '' ************getting the companyid
   
   Dim RsT As New ADODB.Recordset
    sql = "SELECT * FROM sysobjects WHERE NAME =N'ChildCompanies' AND xtype='U'"
    Set RsT = CConnect.GetRecordSet(sql)
   
    'RST.RecordCount
    If Not (RsT.EOF Or RsT.BOF) Then
        sql = "Select Count(CompanyName) AS CoCount From ChildCompanies"
        Set RsT = CConnect.GetRecordSet(sql)
        
        If (RsT!CoCount > 0) Then
            'show company selector
            sql = "select top 1 childcompanyid as companyid from ChildCompanies where upper(companyname)='" & UCase(companyDetail.CompanyName) & "'"
        Else
            sql = "select top 1 companyid as companyid from CompanyDetails where upper(companyname)='" & UCase(companyDetail.CompanyName) & "'"
        End If
        Set RsT = CConnect.GetRecordSet(sql)
        If Not (RsT.EOF) Then
        CompanyId = RsT!CompanyId
        Else
        CompanyId = 0
        End If
    End If
   
   ''**********************************
    
    
    Set AllEmployees = New HRCORE.Employees
    
                
     '
     ''''''''''''''''
     
    Set pBanks = New Banks
    Set pBankBranches = New BankBranches
   '' Set pEmployeeBankAccounts = New EmployeeBankAccounts2
    
    
    
    frmMain2.Show
'    frmSplash.Show

    Exit Sub
    
ErrHandler:
    MsgBox "An error has occured while starting Personnel Director" & _
    vbNewLine & err.Description & _
    vbNewLine & "PDR will be closed", vbExclamation, "PDR Error"

End Sub

Public Sub RefreshEmployeesCol()
    AllEmployees.GetAccessibleEmployeesByUser (currUser.UserID)
End Sub
Public Sub getphotosetup()
On Error GoTo err
Dim rs As New Recordset
Set rs = con.Execute("select active from PhotosSetup")
If Not rs.EOF Then
    If rs!active = 0 Then
      photoisactive = False
    Else
     photoisactive = True
    End If
Else
  photoisactive = False
End If
Exit Sub
err:
photoisactive = False
MsgBox ("Error occured trying to get the photos setup")
End Sub

Public Sub PositionForm(TheForm As Form)
    TheForm.Top = frmMain2.tvwMain.Top
    TheForm.Left = frmMain2.tvwMain.Left + frmMain2.tvwMain.Width
End Sub

Public Function SQLDate(myDate As Date) As String
    If IsDate(myDate) = True Then
        SQLDate = Format(myDate, "yyyy") & Format(myDate, "mm") & Format(myDate, "dd")
    Else
        SQLDate = Format(Date, "yyyy") & Format(Date, "mm") & Format(Date, "dd")
    End If
End Function

Function CheckForNumbers(mystring As String) As Integer
    Dim countForbidden, i As Integer
    countForbidden = 0
    For i = 1 To Len(mystring)
        If Val(Mid(mystring, i, Len(mystring))) <> 0 Then
            countForbidden = 1
        End If
    Next i
    CheckForNumbers = countForbidden
End Function

Public Sub DisableCmd()
    With frmMain2
        .cmdNew.Enabled = False
        .cmdEdit.Enabled = False
        .cmdDelete.Enabled = False
        .cmdSave.Enabled = True
        .cmdCancel.Enabled = True
    End With
End Sub

Public Sub EnableCmd()
    With frmMain2
        .cmdNew.Enabled = True
        .cmdEdit.Enabled = True
        .cmdDelete.Enabled = True
        .cmdSave.Enabled = False
        .cmdCancel.Enabled = False
    End With
End Sub

Public Sub updateSalaryChanges(bp As Double, ha As Double, ta As Double, oa As Double, la As Double, empid As Integer, Optional incrementType As String)
    Dim rs_usc As New ADODB.Recordset
    Dim rs_usc2 As New ADODB.Recordset
    On Error GoTo ErrHandler
    Set rs_usc = CConnect.GetRecordSet("SELECT * FROM pVwRsGlob WHERE employee_id = " & empid)
    With rs_usc
        If .EOF = False Then
            If (bp <> !BasicPay) Or (ha <> !hallow) Or (ta <> !tallow) Or (oa <> !oallow) Or (la <> !lallow) Then
                If incrementType <> "" Then
                    CConnect.ExecuteSql ("INSERT INTO pdSalaryChange (employee_id, changedate, basicpay, hallow, tallow, oallow, lallow, increaseType, pBP, pHA, pTA, pLA, pOA) VALUES (" & empid & ",getdate()," & bp & "," & ha & "," & ta & "," & oa & "," & la & ",'" & incrementType & "'," & !BasicPay & "," & !hallow & "," & !tallow & "," & !lallow & "," & !oallow & ")")
                Else
                    CConnect.ExecuteSql ("INSERT INTO pdSalaryChange (employee_id, changedate, basicpay, hallow, tallow, oallow, lallow, increaseType, pBP, pHA, pTA, pLA, pOA) VALUES (" & empid & ",getdate()," & bp & "," & ha & "," & ta & "," & oa & "," & la & ",'Adjustment'," & !BasicPay & "," & !hallow & "," & !tallow & "," & !lallow & "," & !oallow & ")")
                End If
    
'                Update Job Progression
                Set rs_usc2 = CConnect.GetRecordSet("SELECT * FROM JProg WHERE employee_id = " & empid & " ORDER BY cdate desc")
                If rs_usc.EOF = False Then
                    rs_usc2!BasicPay = bp
                    rs_usc2!hallow = ha
                    rs_usc2!tallow = ta
                    rs_usc2!oallow = oa
                    rs_usc2!lallow = la
                    rs_usc2.Update
                Else
                    updateJobProgression empid, !LCode & "", !Desig & "", !ECategory & "", !Terms & "", bp, ha, la, ta, oa
                End If
            End If
        End If
    End With
    
    Exit Sub
ErrHandler:
End Sub

Public Sub updateJobProgression(empid As Integer, dept As String, position As String, Category As String, Terms As String, Optional bp As Double, Optional ha As Double, Optional la As Double, Optional ta As Double, Optional oa As Double)
    Dim rs_ujp As New ADODB.Recordset
    On Error GoTo ErrHandler
    Set rs_ujp = CConnect.GetRecordSet("SELECT TOP 1 DCode,position,Category,Terms FROM JProg WHERE employee_id = " & empid & " ORDER BY cdate DESC")
    With rs_ujp
        If .EOF = False Then
            If (dept <> !DCode) Or (position <> !position) Or (Category <> !Category) Or (Terms <> !Terms) Or IsNull(!DCode) Or IsNull(!position) Or IsNull(!Category) Or IsNull(!Terms) Then
                CConnect.ExecuteSql ("INSERT INTO JProg (Code, employee_id,position,cdate,basicpay,hallow,lallow,tallow,oallow,category,dcode,terms)" & _
                                    "VALUES('" & loadJCode & "'," & empid & ",'" & position & "',getdate()," & bp & "," & ha & "," & la & "," & ta & "," & oa & ",'" & Category & "','" & dept & "','" & Terms & "')")
            End If
        Else
            CConnect.ExecuteSql ("INSERT INTO JProg (Code, employee_id,position,cdate,basicpay,hallow,lallow,tallow,oallow,category,dcode,terms)" & _
                                "VALUES('" & loadJCode & "'," & empid & ",'" & position & "',getdate()," & bp & "," & ha & "," & la & "," & ta & "," & oa & ",'" & Category & "','" & dept & "','" & Terms & "')")
        End If
    End With
    Exit Sub
ErrHandler:
End Sub

Public Sub updateRS()
With frmMain2
    If .cboTerms.Text <> "All Records" And .cboCat.Text <> "All Records" Then
        Set rs = CConnect.GetRecordSet("SELECT e.*, c.scode, c.Code, c.Description FROM (Employee as e " & _
                "LEFT JOIN SEmp as s ON e.employee_id = s.employee_id) LEFT JOIN CStructure as c ON s.LCode" & _
                "= c.LCode LEFT JOIN ECategory as ec ON e.ECategory = ec.code " & _
                " WHERE (e.Term <> 1) and (e.ECategory = '" & .cboCat.Text & "') AND e.Terms = '" & _
                .cboTerms.Text & "' ORDER BY e.EmpCode")
    
    ElseIf .cboTerms.Text <> "All Records" And .cboCat.Text = "All Records" Then
        Set rs = CConnect.GetRecordSet("SELECT e.*, c.scode, c.Code, c.Description FROM (Employee as e " & _
                "LEFT JOIN SEmp as s ON e.employee_id = s.employee_id) LEFT JOIN CStructure as c ON s.LCode" & _
                "= c.LCode LEFT JOIN ECategory as ec ON e.ECategory = ec.code " & _
            " WHERE (e.Term <> 1) AND e.Terms = '" & .cboTerms.Text & "'  ORDER BY e.EmpCode")
    
    ElseIf .cboTerms.Text = "All Records" And .cboCat.Text <> "All Records" Then
        Set rs = CConnect.GetRecordSet("SELECT e.*, c.scode, c.Code, c.Description FROM (Employee as e " & _
                "LEFT JOIN SEmp as s ON e.employee_id = s.employee_id) LEFT JOIN CStructure as c ON s.LCode" & _
                "= c.LCode LEFT JOIN ECategory as ec ON e.ECategory = ec.code " & _
            " WHERE (e.Term <> 1) and (e.ECategory = '" & .cboCat.Text & "')  ORDER BY e.EmpCode")
    
    ElseIf .cboTerms.Text = "All Records" And .cboCat.Text = "All Records" Then
        Set rs = CConnect.GetRecordSet("SELECT e.*, c.scode, c.Code, c.Description FROM (Employee as e " & _
                "LEFT JOIN SEmp as s ON e.employee_id = s.employee_id) LEFT JOIN CStructure as c ON s.LCode" & _
                "= c.LCode LEFT JOIN ECategory as ec ON e.ECategory = ec.code " & _
            " WHERE (e.Term <> 1)  ORDER BY e.EmpCode")
            
    End If
    
    If .cboStructure.Text <> "All Records" Then
        If Not (rs Is Nothing) Then
            rs.Filter = "scode = '" & .cboStructure.Tag & "'"
        End If
    End If
End With

End Sub

Public Function loadJCode() As String 'Returns Job Progression Code
    Dim rst_temp As New ADODB.Recordset
    Set rst_temp = CConnect.GetRecordSet("SELECT MAX(id) FROM JProg")
    If rst_temp.EOF = False Then
        If rst_temp.RecordCount > 0 And Not IsNull(rst_temp.Fields(0)) Then
            loadJCode = "JP" & CStr(rst_temp.Fields(0) + 1)
        Else
            loadJCode = "JP1"
        End If
    Else
        loadJCode = "JP1"
    End If
    Set rst_temp = Nothing
End Function
''***********************************check to see if the bankaccount send has some amount assiged to it
Public Function HasSomeAmountAssignedToIT(bankaccid As Long, pyr As Long, pm As Long) As Boolean
sql = "select * from tblEmployeeBankNetPay where employeebank_id=" & bankaccid & " and Amount>0 and Period_year=" & pyr & " and period_month=" & pm & ""
Dim rs As New ADODB.Recordset
Set rs = CConnect.GetRecordSet(sql)

If Not rs.EOF Then
HasSomeAmountAssignedToIT = True
Else
HasSomeAmountAssignedToIT = False
End If

End Function

Public Function isDepartmentOrSection(Code As String) As Boolean
    Dim rst_temp As New ADODB.Recordset
    On Error GoTo ErrHandler
    isDepartmentOrSection = False
    Set rst_temp = CConnect.GetRecordSet("SELECT * FROM CStructure WHERE LCode = '" & Code & "'")
    If rst_temp.RecordCount > 0 Then isDepartmentOrSection = True
    Exit Function
ErrHandler:
End Function

Public Function isDivision(Code As String) As Boolean
    Dim rst_temp As New ADODB.Recordset
    On Error GoTo ErrHandler
    isDivision = False
    Set rst_temp = CConnect.GetRecordSet("SELECT * FROM STypes WHERE Code = '" & Code & "'")
    If rst_temp.RecordCount > 0 Then isDivision = True
    Exit Function
ErrHandler:
End Function

Public Function getChildNodes(Code As String) As String
    Dim rst_temp As New ADODB.Recordset
    On Error GoTo ErrHandler
        getChildNodes = getChildNodes & Code & vbTab
        Set rst_temp = CConnect.GetRecordSet("SELECT * FROM CStructure WHERE PCode = '" & Code & "'")
        While rst_temp.EOF = False
            getChildNodes = getChildNodes & getChildNodes(rst_temp!LCode)
            rst_temp.MoveNext
        Wend
    Exit Function
ErrHandler:
End Function

Public Function cQ(str As String) As String
    Dim s As Integer, tmpStr As String, tmpStr2 As String
    s = InStr(str, "'")
    If s > 0 Then
        cQ = Replace(str, "'", "''")
    Else
        cQ = str
    End If
End Function
Public Function getAccessiblePayrollids(user As HRMSEC.CUser) As Long
Dim rss As New ADODB.Recordset
Set rs3 = CConnect.GetRecordSet("exec get_AccessiblePayrollsByUseID " & user.UserID)


   
    With rs3
        If .RecordCount > 0 Then
            .MoveFirst
        
            
            Do While Not .EOF
            
              
                If (strAccessiblePayrollids <> "") Then
                strAccessiblePayrollids = strAccessiblePayrollids & "," & rs3!payroll_id
                Else
                strAccessiblePayrollids = rs3!payroll_id
                End If
               
                .MoveNext
            Loop
        End If
    End With

End Function

Public Sub Disabblepromt()
    frmMain2.cmdShowPrompts.Visible = False
End Sub

