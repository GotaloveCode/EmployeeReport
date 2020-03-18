VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSReports 
   Caption         =   "Reports"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6450
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   6450
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraReports 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2280
      Left            =   210
      TabIndex        =   8
      Top             =   1980
      Visible         =   0   'False
      Width           =   5970
      Begin VB.ComboBox cboRFilter 
         Height          =   315
         ItemData        =   "frmSReports.frx":0000
         Left            =   135
         List            =   "frmSReports.frx":0022
         TabIndex        =   14
         Top             =   1710
         Width           =   3390
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5325
         Picture         =   "frmSReports.frx":0089
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancel Process"
         Top             =   1605
         Width           =   495
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4830
         Picture         =   "frmSReports.frx":018B
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Save Record"
         Top             =   1605
         Width           =   510
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   135
         TabIndex        =   0
         Top             =   555
         Width           =   1230
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1485
         TabIndex        =   1
         Top             =   555
         Width           =   4335
      End
      Begin VB.TextBox txtReport 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   135
         TabIndex        =   2
         Top             =   1125
         Width           =   3645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Report Filter"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   1455
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   315
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Left            =   1485
         TabIndex        =   11
         Top             =   315
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Report Name"
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   885
         Width           =   945
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00800000&
         Caption         =   " Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   6435
      End
   End
   Begin MSComctlLib.ImageList imgTree 
      Left            =   735
      Top             =   1830
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSReports.frx":028D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSReports.frx":0F67
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSReports.frx":1C41
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSReports.frx":291B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSReports.frx":35F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSReports.frx":42CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSReports.frx":4FA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSReports.frx":5C83
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSReports.frx":695D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSReports.frx":7637
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSReports.frx":8311
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSReports.frx":8FEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSReports.frx":9305
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSReports.frx":9FDF
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSReports.frx":ACB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSReports.frx":B993
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRCancel 
      Caption         =   "Cancel"
      Height          =   390
      Left            =   3870
      TabIndex        =   7
      Top             =   6615
      Width           =   1170
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   390
      Left            =   5130
      TabIndex        =   6
      Top             =   6630
      Width           =   1230
   End
   Begin MSComctlLib.TreeView tvwReports 
      Height          =   6375
      Left            =   75
      TabIndex        =   5
      Top             =   60
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   11245
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   706
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "imgTree"
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Report"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmSReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MyNodes As Node
Dim CNode As String
Dim rsMStruc As Recordset
Dim myReport As String
Dim rServer As String, rCatalog As String, rConnection As String, rUId As String, rPass As String

Public Sub InitReports()
Dim mm As String
Dim CNode As String
tvwReports.Nodes.Clear

Set MyNodes = tvwReports.Nodes.Add(, , "Q", "Reports", 14)
Set rsMStruc = CConnect.GetRecordSet("SELECT * FROM SReports WHERE subsystem = '" & SubSystem & "' ORDER BY MyLevel, Code")

With rsMStruc
    If .RecordCount > 0 Then
        .MoveFirst
        CNode = !LinkCode & ""
        Do While Not .EOF
            If !MyLevel = 0 Then
                If !ObjectID = "None" Then
                    Set MyNodes = tvwReports.Nodes.Add(, , !LinkCode, !Description & "", 14)
                Else
                    Set MyNodes = tvwReports.Nodes.Add(, , !LinkCode, !Description & "", 1)
                End If
                MyNodes.EnsureVisible
            Else
                If !ObjectID = "None" Then
                    Set MyNodes = tvwReports.Nodes.Add(!PreviousCode & "", tvwChild, !LinkCode & "", !Description & "", 14)
                Else
                    Set MyNodes = tvwReports.Nodes.Add(!PreviousCode & "", tvwChild, !LinkCode & "", !Description & "", 1)
                End If
                    
                MyNodes.EnsureVisible
            End If
            
            .MoveNext
        Loop
               
        .MoveFirst
    End If
End With

'tvwReports.Nodes("Q005").Expanded = False
'tvwReports.Nodes("Q001").Expanded = False
'tvwReports.Nodes("Q010001").Expanded = False
'tvwReports.Nodes("Q010005").Expanded = False
'tvwReports.Nodes("Q010010").Expanded = False
'tvwReports.Nodes("Q010015").Expanded = False
'tvwReports.Nodes("Q010020").Expanded = False

tvwReports.Refresh

End Sub

Private Sub cboRFilter_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
    fraReports.Visible = False
End Sub

Private Sub cmdPrinterSetup_Click()
    R.PrinterSetup 0
End Sub

Private Sub cmdRCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
Dim LCode As String
Dim PCode As String
Dim MyLevel As Integer
Dim rsM1 As Recordset

'Prompt if required fields are not entered
If CNode = "" Then
    MsgBox "You must select the node you want to add to.", vbInformation
    Exit Sub
End If

If txtCode.Text = "" Then
    MsgBox "Enter the code.", vbExclamation
    txtCode.SetFocus
    Exit Sub
End If

If txtDescription.Text = "" Then
    MsgBox "Enter the Description.", vbExclamation
    txtDescription.SetFocus
    Exit Sub
End If

If txtReport.Text = "" Then
    MsgBox "Enter the report name.", vbExclamation
    txtReport.SetFocus
    Exit Sub
End If

If cboRFilter.Text = "" Then
    MsgBox "Select report filter.", vbExclamation
    cboRFilter.SetFocus
    Exit Sub
End If

If SaveNew = True Then
    Set rsM1 = CConnect.GetRecordSet("SELECT * FROM SReports WHERE LinkCode = '" & CNode & "" & txtCode.Text & "'")
            
    With rsM1
        If .RecordCount > 0 Then
            MsgBox "Code already exists. Enter another one.", vbInformation
            txtCode.Text = ""
            txtCode.SetFocus
            Exit Sub
        End If
    End With
    
    Set rsM1 = Nothing
    
    Set rsM1 = CConnect.GetRecordSet("SELECT * FROM SReports ORDER BY MyLevel, Code")
    
    With rsM1
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "LinkCode like '" & CNode & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                LCode = !LinkCode & "" & txtCode.Text
                MyLevel = !MyLevel + 1
                PCode = !LinkCode & ""
            Else
                LCode = "Q" & txtCode.Text
                MyLevel = 0
            End If
        Else
            LCode = "Q" & txtCode.Text
            MyLevel = 0
        End If
    End With
    
    Set rsM1 = Nothing
    
    
    CConnect.ExecuteSql ("DELETE FROM SReports WHERE LinkCode = '" & LCode & "'")
            
    CConnect.ExecuteSql ("INSERT INTO SReports (LinkCode, PreviousCode, MyLevel, Code, Description, ObjectID, RFilter)" & _
                " VALUES('" & LCode & "','" & PCode & "'," & MyLevel & ",'" & txtCode.Text & "','" & txtDescription.Text & "','" & txtReport.Text & "','" & cboRFilter.Text & "')")
    
    Call InitReports
    Cleartxt
    txtCode.SetFocus
Else
    CConnect.ExecuteSql ("UPDATE SReports SET Description = '" & txtDescription.Text & "', ObjectID = '" & txtReport.Text & "', RFilter = '" & cboRFilter.Text & "' WHERE LinkCode = '" & CNode & "'")
    Call InitReports
    fraReports.Visible = False
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdView_Click()

Dim conProps As CRAXDDRT.ConnectionProperties
On Error GoTo errHandler
If CNode = "" Then Exit Sub

'// Full rights to admin
If CGroup = "PROSOFT" Or CGroup = "ADMIN" Or CGroup = "SUP" Then GoTo RUKAHAPA

'++Kazia Wasee kaa hawana rights++
If Get_Group_right(CGroup, CNode) = "NONE" Then
    Me.MousePointer = 0
    MsgBox "Sorry! You Have Insufficient Rights to view this Report" & vbCrLf & "Contact System Administrator", vbExclamation, "GROUP RIGHTS": Exit Sub
    Exit Sub
End If
RUKAHAPA:
    With rsMStruc
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "LinkCode like '" & CNode & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                If !ObjectID <> "None" Then
                    Debug.Print !ObjectID
                    ReportType = !ObjectID & ""
                    RFilter = !RFilter & ""
                    If RFilter = "None" Then
                        Me.MousePointer = vbHourglass
                        Set a = New CRAXDDRT.Application
                        Set R = a.OpenReport(App.Path & "\HR Base Reports\" & "" & ReportType)
                        
                        If R.HasSavedData = True Then
                            R.DiscardSavedData
                        End If
                    
                    ' Loop through all database tables and set the correct server & database
                        Dim tbl As CRAXDDRT.DatabaseTable
                        Dim tbls As CRAXDDRT.DatabaseTables
                        
                        Set tbls = R.Database.Tables
'                        tbls(1).DllName = "crdb_odbc.dll"
                        For Each tbl In tbls
                            Set conProps = tbl.ConnectionProperties
                            conProps.DeleteAll
                            conProps.Add "connectionstring", connection_string
'                            conProps.Add "DSN", rConnection
'                            conProps.Add "Database", rCatalog
'                            conProps.Add "UID", rUId
'                            conProps.Add "PWD", rPass
                            tbl.Location = tbl.Name
                        Next tbl
                        
                        Dim rsCompany As New ADODB.Recordset
                        Dim rsCompName As New ADODB.Recordset
                        Set rsCompany = CConnect.GetRecordSet("select * from STYPES where smain=1")
                        If rsCompany.RecordCount > 0 Then
                            Set rsCompName = CConnect.GetRecordSet("select * from GeneralOpt")
                            If rsCompName.RecordCount > 0 Then
                                If UCase(Trim(rsCompany!Description & "")) = UCase(Trim(rsCompName!cName & "")) Then
                                    R.ReportTitle = UCase(Trim(rsCompany!Description & ""))
                                Else
                                    R.ReportTitle = UCase(Trim(rsCompName!cName & "")) & " - " & UCase(Trim(rsCompany!Description & ""))
                                End If
                            Else
                                R.ReportTitle = UCase(Trim(rsCompany!Description & ""))
                            End If
                        Else
                            R.ReportTitle = "TEST COMPANY"
                        End If
                        
                        With frmReports.CRViewer1
                            .DisplayGroupTree = False
                            .EnableAnimationCtrl = False
                            .ReportSource = R
                            .ViewReport
                        End With
                        
                        frmReports.Show vbModal
                        Me.MousePointer = 0
    
                    Else
                        frmRange.Show vbModal, Me
                    End If
                    
                End If
            End If
            
        End If
    End With

Exit Sub
errHandler:
    MsgBox Err.Description, vbInformation
    Me.MousePointer = 0

End Sub

Private Sub Form_Load()
    
    Call InitReports
'    Call InitRConnection
End Sub

Private Sub mnuAdd_Click()
    If CNode = "" Then
        MsgBox "You must select the node you want to add to.", vbInformation
        Exit Sub
    End If
    
    Call Cleartxt
    SaveNew = True
    fraReports.Visible = True
    txtCode.Locked = False
    txtCode.SetFocus
    
End Sub

Private Sub mnuDelete_Click()
    If CNode = "" Then
        MsgBox "You must select the node you want to delete.", vbInformation
        Exit Sub
    End If
    
    CConnect.ExecuteSql ("DELETE FROM SReports WHERE LinkCode = '" & CNode & "'")
    Call InitReports
End Sub

Private Sub mnuEdit_Click()
    If CNode = "" Then
        MsgBox "You must select the node you want to edit.", vbInformation
        Exit Sub
    End If
    
    With rsMStruc
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "LinkCode like '" & CNode & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                fraReports.Visible = True
                txtCode.Text = !code & ""
                txtDescription.Text = !Description & ""
                txtReport.Text = !ObjectID & ""
                cboRFilter.Text = !RFilter & ""
                txtCode.Locked = True
                SaveNew = False
            End If
        End If
    End With
    
        
End Sub

Private Sub tvwReports_DblClick()
    Call cmdView_Click
End Sub

Private Sub tvwReports_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    Me.PopupMenu mnuReport
End If
End Sub

Private Sub tvwReports_NodeClick(ByVal Node As MSComctlLib.Node)
    If tvwReports.Nodes.Count > 0 Then
        CNode = tvwReports.SelectedItem.Key
    End If
End Sub

Public Sub Cleartxt()
txtCode.Text = ""
txtDescription.Text = ""
txtReport.Text = ""
End Sub

Public Sub disabletxt()
Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
            i.Locked = True
        End If
    Next i

    
End Sub

Public Sub InitRConnection()
    Dim rec_t As New ADODB.Recordset
    Set rec_t = CConnect.GetRecordSet("SELECT servername, connectionname, dcatalog, userid, passwd FROM GeneralOpt WHERE subsystem = '" & SubSystem & "'")
    With rec_t
        If rec_t.EOF = False Then
            rConnection = Trim(!connectionName & "")
            rServer = Trim(!ServerName & "")
            rCatalog = Trim(!dcatalog & "")
            rUId = Trim(!UserID & "")
            rPass = Trim(!passwd & "")
        End If
    End With
End Sub
