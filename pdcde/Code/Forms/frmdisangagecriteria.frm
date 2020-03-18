VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmdisangagecriteria 
   Caption         =   "Criteria"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5550
   LinkTopic       =   "Criteria"
   ScaleHeight     =   7830
   ScaleWidth      =   5550
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "Departments"
      Height          =   3615
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   5295
      Begin MSComctlLib.ListView LvwDpts 
         Height          =   3135
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Departments"
            Object.Width           =   5821
         EndProperty
      End
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton cmdview 
      Caption         =   "View"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   7320
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.Frame Frame3 
         Caption         =   "Reason"
         Height          =   1935
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   4935
         Begin MSComctlLib.ListView lvwreasons 
            Height          =   1455
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   2566
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
      End
      Begin VB.Frame Frame2 
         Caption         =   "Disangagement Date"
         Height          =   1095
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4935
         Begin MSComCtl2.DTPicker dtto 
            Height          =   255
            Left            =   3240
            TabIndex        =   5
            Top             =   720
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
            Format          =   84672513
            CurrentDate     =   39905
         End
         Begin MSComCtl2.DTPicker dtfrom 
            Height          =   255
            Left            =   1320
            TabIndex        =   4
            Top             =   720
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            _Version        =   393216
            Format          =   84672513
            CurrentDate     =   39905
         End
         Begin VB.OptionButton optbetween 
            Caption         =   "Between"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton optallemps 
            Caption         =   "All"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "frmdisangagecriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strReasons As String
Private strDater As String

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdView_Click()



strReasons = ""
Dim i As Integer
If (lvwreasons.ListItems.count > 0) Then
For i = 2 To lvwreasons.ListItems.count
  If (lvwreasons.ListItems(i).Checked) Then
    If (Trim(strReasons) <> "") Then
    strReasons = strReasons & ",'" & lvwreasons.ListItems(i).Text & "'"
    Else
    strReasons = "'" & lvwreasons.ListItems(i).Text & "'"
    End If
  End If
Next i
End If

If (Trim(strReasons) <> "") Then
mySQL = "({Employee.TermReasons} in [" & strReasons & "])"
Else
MsgBox ("Reason(s) for disengagement were selected. please sellect atleast 1")
Exit Sub
End If



strDater = ""

If (RFilter = "disangagement") Then
    If (optbetween.value = True) Then
      strDater = "({VWArchivedEmployees.Dleft} In DateTime " & Format(dtfrom.value, "(yyyy, mm, dd, hh, nn, ss)") & " to DateTime " & Format(dtto.value, "(yyyy, mm, dd, hh, nn, ss)") & ")"
    End If
Else

If (RFilter = "rengagementhist") Then
If (optbetween.value = True) Then
strDater = "({VWrengageemployeereprt.datereengaged} In DateTime " & Format(dtfrom.value, "(yyyy, mm, dd, hh, nn, ss)") & " to DateTime " & Format(dtto.value, "(yyyy, mm, dd, hh, nn, ss)") & ")"
End If
Else
If (optbetween.value = True) Then
strDater = "({VWreengagementhistory.datereengaged} In DateTime " & Format(dtfrom.value, "(yyyy, mm, dd, hh, nn, ss)") & " to DateTime " & Format(dtto.value, "(yyyy, mm, dd, hh, nn, ss)") & ")"
End If
End If
End If
RFilter = ""

If (strDater <> "") Then
If (mySQL <> "") Then
mySQL = mySQL & " and " & strDater
Else

End If
End If


           Dim depts As String
           depts = ""
        For k = 2 To LvwDpts.ListItems.count
            If (LvwDpts.ListItems.Item(k).Checked = True) Then
            If (Trim(depts) <> "") Then
            depts = depts & "," & Trim(LvwDpts.ListItems.Item(k).Tag)
            Else
            depts = Trim(LvwDpts.ListItems.Item(k).Tag)
            End If
                
            End If
        Next k
        If (Trim(depts) = "") Then
            MsgBox "You must have at least one department selected", vbExclamation, "Report Error"
            Exit Sub
'        Else
'            depts = Mid$(depts, 1, Len(depts) - 1)
        End If

If (Trim(depts) <> "") Then
If (Trim(mySQL) <> "") Then
mySQL = mySQL & " and ({Employee.CStructure_ID} in [" & depts & "])"
Else
mySQL = "{Employee.CStructure_ID} in [" & depts & "]"
End If
End If



ShowReport R
End Sub








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
        
       ' rpt.FormulaSyntax = crCrystalSyntaxFormula
       If (Trim(mySQL) <> "") Then
       rpt.RecordSelectionFormula = mySQL
       End If
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
        
        frmReports.Show vbModal
'        frmReports.Show vbModal
'
Unload Me
        frmMain2.MousePointer = vbNormal
    Exit Sub
errHandler:
    MsgBox err.Description, vbInformation
    frmMain2.MousePointer = vbNormal
End Sub


Private Sub Form_Load()
If (RFilter = "Disangagement Date") Then
Frame2.Caption = "Disangagement Date"
Else
Frame2.Caption = "Re-angagement Date"
End If
lvwreasons.Checkboxes = True
lvwreasons.ColumnHeaders.add , , "Reason", 2350
Dim li As ListItem
Set li = lvwreasons.ListItems.add(, , "Select All")
li.Tag = "S_A"
lvwreasons.Checkboxes = True
load_reasons
PopulateDpts
End Sub

Private Sub load_reasons()
sQL = "select reason from DisengagementReasons"
Dim rs As New ADODB.Recordset
Set rs = CConnect.GetRecordSet(sQL)
If Not rs Is Nothing Then
Dim li As ListItem
lvwreasons.ListItems.Clear
Set li = lvwreasons.ListItems.add(, , "Select All")
li.Tag = "S_A"
lvwreasons.Checkboxes = True
While rs.EOF = False

Set li = lvwreasons.ListItems.add(, , rs!Reason)
rs.MoveNext
Wend
End If
End Sub

Private Sub LvwDpts_ItemCheck(ByVal Item As MSComctlLib.ListItem)
  On Error GoTo ErrorHandler
    Dim n As Integer, State As Boolean
    
    State = False
    If LvwDpts.ListItems.Item(1).Checked = True Then State = True
    
    If Item.Tag = "S-A" Then
        'Uncheck All Departments
        For n = 2 To LvwDpts.ListItems.count
            LvwDpts.ListItems.Item(n).Checked = State
        Next n
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error Selecting Departments:" & vbNewLine & err.Description, vbExclamation, "PDR Error"
End Sub

Private Sub lvwreasons_ItemCheck(ByVal Item As MSComctlLib.ListItem)
 On Error GoTo ErrorHandler
    Dim n As Integer, State As Boolean
    
    State = False
    If lvwreasons.ListItems.Item(1).Checked = True Then State = True
    
    If Item.Tag = "S_A" Then
        'Uncheck All Departments
        For n = 2 To lvwreasons.ListItems.count
            lvwreasons.ListItems.Item(n).Checked = State
        Next n
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error Selecting reasons:" & vbNewLine & err.Description, vbExclamation, "PDR Error"

End Sub

Private Sub optallemps_Click()
If (optallemps.value = True) Then
dtfrom.Visible = False
dtto.Visible = False
Else
dtfrom.Visible = True
dtto.Visible = True
End If
End Sub

Private Sub optbetween_Click()
If (optbetween.value = True) Then
dtfrom.Visible = True
dtto.Visible = True
Else
dtfrom.Visible = False
dtto.Visible = False
End If
End Sub
Private Sub PopulateDpts()
    On Error GoTo ErrorHandler
    Dim rsOUs As New ADODB.Recordset, ItemD As ListItem
    
    Set rsOUs = CConnect.GetRecordSet("Select * From OrganizationUnits Order by OrganizationUnitName")
    If rsOUs.RecordCount > 0 Then
        Do Until rsOUs.EOF
            With LvwDpts
                Set ItemD = .ListItems.add(, , rsOUs!OrganizationUnitName)
                ItemD.Tag = rsOUs!OrganizationUnitID
                ItemD.Checked = True
            End With
            rsOUs.MoveNext
        Loop
    End If
    
    Set ItemD = LvwDpts.ListItems.add(1, , "(Select All)")
    ItemD.Tag = "S-A"
    LvwDpts.ListItems.Item(1).Checked = True
    ItemD.Bold = True
        
    Exit Sub
ErrorHandler:
    MsgBox "Error Loading Departments" & vbNewLine & err.Description, vbExclamation, "PDR Error"
End Sub
