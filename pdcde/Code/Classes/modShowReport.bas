Attribute VB_Name = "modShowReport"
'THIS IS NECESSARY FOR PASSING PARAMETERFIEL OBJECT VALUES
Public Type ParamFieldObj
    Name As String
    value As Variant
End Type

'THESE IS NECESSARY FOR PASSING PARAM VALUES
Public objParamField() As ParamFieldObj
Public Function DateString(dDate As Date) As String
    On Error GoTo errHandler
        Dim strYear As String
        Dim strMonth As String
        Dim strDay As String


        strYear = Year(dDate)
        strMonth = Month(dDate)
        strDay = Day(dDate)

        If Len(strMonth) < 2 Then strMonth = "0" & strMonth
        If Len(strDay) < 2 Then strDay = "0" & strDay

        DateString = strYear & strMonth & strDay

    Exit Function
errHandler:
    DateString (Date)
'    MsgBox Err.Description, vbExclamation, TITLES
End Function

Public Sub ShowReport(rpt As CRAXDDRT.Report, Optional formula As String, Optional blnAlterParamValue As Boolean)
    Dim conProps As CRAXDDRT.ConnectionProperties

    On Error GoTo errHandler
    'force crystal to use the basic syntax report

    frmMain2.MousePointer = vbHourglass

    If rpt.HasSavedData = True Then
        rpt.DiscardSavedData
    End If


    ' Loop through all database tables and set the correct server & database
        Dim tbl As CRAXDDRT.DatabaseTable
        Dim tbls As CRAXDDRT.DatabaseTables

        Set tbls = rpt.Database.Tables
        For Each tbl In tbls

            On Error Resume Next
            Set conProps = tbl.ConnectionProperties
            conProps.DeleteAll
            If tbl.DllName <> "crdb_ado.dll" Then
                tbl.DllName = "crdb_ado.dll"
            End If

            conProps.add "Data Source", ConParams.DataSource
            conProps.add "Initial Catalog", ConParams.InitialCatalog
            conProps.add "Provider", "SQLOLEDB.1"
            'conProps.Add "Integrated Security", "true"
            conProps.add "User ID", ConParams.UserID
            conProps.add "Password", ConParams.Password
            conProps.add "Persist Security Info", "false"
            tbl.Location = tbl.Name
        Next tbl

        rpt.FormulaSyntax = crCrystalSyntaxFormula
        'formula
    With rpt
    '        DEALING WITH ALTERED PARAM FIELD OBJECT VALUES
        If blnAlterParamValue = True Then
            For lngLoopVariable = LBound(objParamField) To UBound(objParamField)
                For lngloopvariable2 = 1 To .ParameterFields.count
                    If .ParameterFields.Item(lngloopvariable2).ParameterFieldName = objParamField(lngLoopVariable).Name Then
                        .ParameterFields.Item(lngloopvariable2).SetCurrentValue (objParamField(lngLoopVariable).value)
                        Exit For
                    End If

                Next
            Next
        End If
        .EnableParameterPrompting = False
    End With
    rpt.PaperSize = crPaperA4

    If rpt.PaperOrientation = crLandscape Then
        rpt.BottomMargin = 192
        rpt.RightMargin = 720
        rpt.LeftMargin = 58
        rpt.TopMargin = 192
    ElseIf rpt.PaperOrientation = crPortrait Then
        rpt.BottomMargin = 300
        rpt.RightMargin = 338
        rpt.LeftMargin = 300
        rpt.TopMargin = 281
    End If
        With frmReports.CRViewer1
            .DisplayGroupTree = False
            .EnableAnimationCtrl = False
            .ReportSource = rpt
            .ViewReport
        End With

        formula = ""
        frmReports.Show
'        frmReports.Show vbModal
        frmMain2.MousePointer = vbNormal
    Exit Sub
errHandler:
    MsgBox err.Description, vbInformation
    frmMain2.MousePointer = vbNormal
End Sub

Public Sub pbDisplayReportUsingRDC(ReportName As CRAXDRT.Report, Optional Selection As String, Optional sReportTitle As String, Optional SubTitle As String, Optional CallerFormName As Form, Optional ByVal ReportPrintRight As String, Optional blnAlterParamValue As Boolean)

'    On Error GoTo ErrHandler
    Dim MyDatabase As CRAXDRT.Database
    Dim myTable As CRAXDRT.DatabaseTable
    Dim myTables As CRAXDRT.DatabaseTables
    Dim myConprops As CRAXDRT.ConnectionProperties
    Dim ReportSelection As String
    Dim GroupSelection As String
    Dim lngLoopVariable As Long
    Dim lngloopvariable2 As Long

    Dim strPath As String
    Dim attempted As Boolean

    If IsMissing(ReportPrintRight) Then
        ReportModifyPolicy = "MISSING"
    Else
        ReportModifyPolicy = ReportPrintRight
    End If
    ReportName.DiscardSavedData

    Set myTables = ReportName.Database.Tables
    attempted = False
'     blnDllChanged = False
Display:

  On Error GoTo errHandler

    For Each myTable In myTables
        Set myConprops = myTable.ConnectionProperties
        With myConprops
            .DeleteAll
            .add "Data Source", ConParams.DataSource
            .add "Initial Catalog", ConParams.InitialCatalog
            .add "User ID", ConParams.UserID
            .add "Password", ConParams.Password
            .add "Provider", "SQLOLEDB.1"
        End With

        If Not blnDllChanged Then
            myTable.DllName = "crdb_ado.dll"
            blnDllChanged = True
        End If
ContinueAfterDLLChange:
        myTable.Location = myTable.Name
    Next myTable

    '//Allocate the connection properties to sub reports if any
    Dim objs As CRAXDRT.ReportObjects
    Dim obj As Object
    Dim rpt As CRAXDRT.Report
    Dim sec As CRAXDRT.Section
    Dim mytable2 As CRAXDRT.DatabaseTable
    Dim subrpt As CRAXDRT.SubreportObject
    Dim myconprops2 As CRAXDRT.ConnectionProperties
    Dim blnDLLChanged2 As Boolean
    Dim mycomplogo As CRAXDRT.OLEObject
    Dim i As Integer

    ReportName.FormulaSyntax = crCrystalSyntaxFormula 'VERY NECESSARY TO TELL CRpt WE USING CRYSTAL SYNTAX IN SPECIFYING FORMULAS

    For Each sec In ReportName.Sections
'        Set objs = sec.ReportObjects
        For Each obj In sec.ReportObjects
            If obj.Kind = crSubreportObject Then
                Set subrpt = obj
                Set rpt = subrpt.OpenSubreport()
                rpt.FormulaSyntax = crCrystalSyntaxFormula 'VERY NECESSARY TO TELL CRpt WE USING CRYSTAL SYNTAX IN SPECIFYING FORMULAS

                For Each mytable2 In rpt.Database.Tables
                    Set myconprops2 = mytable2.ConnectionProperties
                    With myconprops2
                        .DeleteAll
                        .add "Data Soure", pServerName
                        .add "Database", pDatabaseName
                        .add "User ID", myUserID
                        .add "Password", myPassword
                        .add "Provider", con.Provider
                    End With
                    mytable2.Location = mytable2.Name
                Next mytable2
            End If
        Next obj
    Next sec

    ReportSelection = ""
    GroupSelection = ""

    With ReportName

        'DEALING WITH ALTERED PARAM FIELD OBJECT VALUES
        If blnAlterParamValue = True Then
            For lngLoopVariable = LBound(objParamField) To UBound(objParamField)
                For lngloopvariable2 = 1 To .ParameterFields.count
                    If .ParameterFields.Item(lngloopvariable2).ParameterFieldName = objParamField(lngLoopVariable).Name Then
                        .ParameterFields.Item(lngloopvariable2).SetCurrentValue (objParamField(lngLoopVariable).value)
                        Exit For
                    End If
                Next
            Next
        End If
        .EnableParameterPrompting = False

        If sReportTitle <> "" Then
            .ReportTitle = sReportTitle
        End If

        If Not .RecordSelectionFormula = "" Then
            If Selection <> "" Then
                ReportSelection = .RecordSelectionFormula
'                Selection = ReportSelection & " and " & Selection
            Else
                Selection = .RecordSelectionFormula
            End If
        End If

        If Trim$(.GroupSelectionFormula) = "" Then
            Selection = Selection
        Else
            GroupSelection = .GroupSelectionFormula
'            Selection = GroupSelection & " and " & Selection
        End If

        Debug.Print "Record: " & ReportSelection
        Debug.Print "Group: " & GroupSelection

        .RecordSelectionFormula = Selection

        If .HasSavedData = True Then
            .DiscardSavedData
        End If

    End With

    With frmReports.CRViewer1
        .Refresh
        .EnablePrintButton = True '//Use rights
        .EnableSearchControl = True
        .EnableZoomControl = True
        .EnableHelpButton = True
        .DisplayBorder = True
        .DisplayBackgroundEdge = True
        .DisplayTabs = False
        .DisplayToolbar = True
        .EnableExportButton = True
        .ReportSource = ReportName
        .EnableGroupTree = False
        .EnableAnimationCtrl = False
        .Zoom 1
        .ViewReport
    End With

    frmReports.Show

    Exit Sub

errHandler:

    If err.Number = 401 Then '// Cannot show modal form when a on modal form is displayed

        Unload CallerFormName '// Wheresuch error exist, one has to take advantage of the situation
        Resume Next
    Else
        Debug.Print err.Number
        If err.Number = -2147191858 Then GoTo ErrH
        If err.Number = -2147191803 Then GoTo ErrH 'String is required here
        If err.Number = -2147184634 Then blnDllChanged = True: Resume ContinueAfterDLLChange

     End If
     MsgBox err.Description & err.Number & " When Displaying the Report"
ErrH:
       Resume Next
End Sub


