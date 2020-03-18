VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEmployeesOnBank 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.CommandButton cmdprint 
         Caption         =   "Print"
         Height          =   375
         Left            =   4080
         TabIndex        =   5
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton cmdclosed 
         Caption         =   "Close"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   4680
         Width           =   1335
      End
      Begin MSComctlLib.ListView lvwEmployeesOnBank 
         Height          =   3975
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   7011
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox txtbankname 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1200
         TabIndex        =   1
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblbankname 
         Alignment       =   2  'Center
         Caption         =   "BANK="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   120
         Width           =   5655
      End
   End
End
Attribute VB_Name = "frmEmployeesOnBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim myEmpsBankAccs As EmployeeBankAccounts2

Private Sub cmdclosed_Click()
Unload Me
End Sub

Private Sub cmdprint_Click()
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("AwardReport") = secNone Then
            MsgBox "You Don't have right to view the report. Please liaise with security admin"
            Exit Sub
        End If
   End If
    'VIEW REPORT OF AWARDS AWARDED TO EMPLOYEE
    Set r = crtEmployeesOnBank
    r.reportTitle = "LIST OF REPORTS ATTACHED TO " & txtBankName.Text
    mySQL = "{EB.BankName} = '" & txtBankName.Text & "'"
    printReport r
'    ReportSchemaName = "Employee"
'    ReportType = "Normal"
'    frmHeadcountFilter.Show
End Sub
Private Sub printReport(rpt As CRAXDDRT.Report)


   Dim conProps As CRAXDDRT.ConnectionProperties
    
    On Error GoTo ErrHandler
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
              tbl.Name = "EB"
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
        rpt.RecordSelectionFormula = mySQL
        
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
       '' rpt.PrintOut False, 1, True, 1, 1
        
        
        
        formula = ""
     frmReports.Show vbModal
    Me.MousePointer = 0
      
       frmMain2.MousePointer = vbNormal
        Unload Me
    Exit Sub
ErrHandler:
    MsgBox err.Description, vbInformation
    frmMain2.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
Me.Left = frmMain2.Left + frmBanks.Width
Me.Top = frmBanks.Top + 200
lvwEmployeesOnBank.ColumnHeaders.add , , "Employee Code", 1200
lvwEmployeesOnBank.ColumnHeaders.add , , "Employee Names", 3600
Set myEmpsBankAccs = New EmployeeBankAccounts2
myEmpsBankAccs.GetAllEmployeeBankAccounts
txtBankName.Text = frmBanks.txtBankName
lblbankname.Caption = lblbankname.Caption & " " & txtBankName.Text
End Sub

Private Sub Text1_Change()
displayEmployess
End Sub

Private Sub displayEmployess()
lvwEmployeesOnBank.ListItems.Clear
Dim i As Long
Dim li As ListItem
For i = 1 To myEmpsBankAccs.count
If myEmpsBankAccs.Item(i).bankbranch.Bank.BankName = txtBankName.Text Then

Set li = lvwEmployeesOnBank.ListItems.add(, , myEmpsBankAccs.Item(i).Employee.EmpCode)
li.ListSubItems.add , , myEmpsBankAccs.Item(i).Employee.SurName & myEmpsBankAccs.Item(i).Employee.OtherNames

End If

Next i
End Sub

Private Sub Form_LostFocus()
Me.Visible = False
End Sub

Private Sub txtbankname_Change()
displayEmployess
End Sub
