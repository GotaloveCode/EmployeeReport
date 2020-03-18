VERSION 5.00
Begin VB.Form frmexempinfofilter 
   Caption         =   "Form1"
   ClientHeight    =   1755
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   5685
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "View"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.TextBox txtempcode 
         Height          =   405
         Left            =   2400
         TabIndex        =   3
         Top             =   840
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.OptionButton optviewindividual 
         Caption         =   "View For Individual"
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   360
         Width           =   2415
      End
      Begin VB.OptionButton optviewall 
         Caption         =   "View All"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.Label lblempcode 
         Caption         =   "Emp Code"
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmexempinfofilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If (optviewindividual.value = True And txtempcode.Text = "") Then
MsgBox ("Employee code as not been entered")
Exit Sub
End If

If (optviewall.value <> True) Then
mySQL = "{dbo.spget_exemployeeinfo.empcode}= '" & Trim(txtempcode.Text) & "'"
Else
mySQL = ""
End If
ShowReport R
Unload Me
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
            
'            If (Trim(mySQL) <> "") Then
'            tbl.Name = "E"
'            End If
                    
            conProps.add "Data Source", ConParams.DataSource
            conProps.add "Initial Catalog", ConParams.InitialCatalog
            conProps.add "Provider", "SQLOLEDB.1"
            'conProps.Add "Integrated Security", "true"
            conProps.add "User ID", ConParams.UserID
            conProps.add "Password", ConParams.Password
            conProps.add "Persist Security Info", "false"
            tbl.Location = tbl.Name
        Next tbl
        
       '' rpt.FormulaSyntax = crCrystalSyntaxFormula
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
        frmReports.Show
'        frmReports.Show vbModal
        frmMain2.MousePointer = vbNormal
    Exit Sub
errHandler:
    MsgBox err.Description, vbInformation
    frmMain2.MousePointer = vbNormal
End Sub


Private Sub optviewall_Click()
If (optviewall.value = True) Then
lblempcode.Visible = False
txtempcode.Visible = False
Else
lblempcode.Visible = True
txtempcode.Visible = True
End If

End Sub

Private Sub optviewindividual_Click()
If (optviewindividual.value = True) Then
lblempcode.Visible = True
txtempcode.Visible = True
Else
lblempcode.Visible = False
txtempcode.Visible = False
End If
End Sub
