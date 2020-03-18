VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmGenOpt 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Options"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9945
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   9945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNew 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      Picture         =   "frmGenOpt.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "Add New record"
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComctlLib.ImageList imgEmpTool 
      Left            =   5355
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGenOpt.frx":0102
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGenOpt.frx":0214
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGenOpt.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGenOpt.frx":0438
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraGenOpt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8000
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   9930
      Begin VB.CheckBox chkCapital 
         Appearance      =   0  'Flat
         Caption         =   "&In capital letters"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1320
         TabIndex        =   56
         Top             =   200
         Width           =   1575
      End
      Begin VB.Frame fraEmpPrompt 
         Height          =   1695
         Left            =   5400
         TabIndex        =   49
         Top             =   4920
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdExpiryCancel 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   2760
            TabIndex        =   54
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton cmdExpiryOk 
            Caption         =   "&Ok"
            Height          =   375
            Left            =   1560
            TabIndex        =   53
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtExpiryMonth 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   120
            TabIndex        =   51
            Top             =   1080
            Width           =   1305
         End
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Prompt for employment expiry within .... months"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   720
            Width           =   3495
         End
         Begin VB.Label lblCaption 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Define When To Prompt For Expiry"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   3840
         End
      End
      Begin VB.CommandButton cmdContractExpiryAlert 
         Caption         =   "Employment Ex&piry Alert"
         Height          =   390
         Left            =   2760
         TabIndex        =   48
         Top             =   6480
         Width           =   2535
      End
      Begin VB.CommandButton cmdBonus 
         Caption         =   "Calculate Anual Bonus"
         Height          =   390
         Left            =   120
         TabIndex        =   20
         Top             =   6480
         Width           =   2100
      End
      Begin VB.CommandButton cmdRConnections 
         Caption         =   "Report Connections"
         Height          =   390
         Left            =   120
         TabIndex        =   18
         Top             =   5520
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.Frame fraEmpTerms 
         Height          =   3495
         Left            =   5280
         TabIndex        =   45
         Top             =   720
         Visible         =   0   'False
         Width           =   4095
         Begin VB.TextBox txtTerms 
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   720
            TabIndex        =   21
            Top             =   480
            Width           =   3225
         End
         Begin VB.CommandButton cmdDeleteT 
            Caption         =   "Delete"
            Height          =   345
            Left            =   2160
            TabIndex        =   24
            Top             =   3000
            Width           =   810
         End
         Begin VB.CommandButton cmdAddT 
            Caption         =   "Add"
            Height          =   345
            Left            =   1200
            TabIndex        =   23
            Top             =   3000
            Width           =   930
         End
         Begin VB.CommandButton cmdExitT 
            Caption         =   "Exit"
            Height          =   345
            Left            =   3000
            TabIndex        =   25
            Top             =   3000
            Width           =   930
         End
         Begin MSComctlLib.ListView lsvEmpTerms 
            Height          =   2055
            Left            =   120
            TabIndex        =   22
            Top             =   840
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   3625
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Label Label16 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Terms"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   47
            Top             =   480
            Width           =   435
         End
         Begin VB.Label lblCaption 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Define Employee Terms"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   120
            TabIndex        =   46
            Top             =   120
            Width           =   2625
         End
      End
      Begin VB.CommandButton cmdAddETerms 
         Caption         =   "Add Employee Type"
         Height          =   390
         Left            =   120
         TabIndex        =   19
         Top             =   6000
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.TextBox txtCasuals 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4815
         TabIndex        =   8
         Top             =   3030
         Width           =   1395
      End
      Begin VB.TextBox txtVisa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4815
         TabIndex        =   6
         Top             =   2325
         Width           =   1395
      End
      Begin VB.TextBox txtContracts 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4815
         TabIndex        =   7
         Top             =   2685
         Width           =   1395
      End
      Begin VB.TextBox txtBirthdayDays 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4815
         TabIndex        =   5
         Top             =   1965
         Width           =   1395
      End
      Begin VB.TextBox txtRetireDays 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4815
         TabIndex        =   4
         Top             =   1605
         Width           =   1395
      End
      Begin VB.CheckBox chkVSal 
         Appearance      =   0  'Flat
         Caption         =   "Allow viewing of employee salary details"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   150
         TabIndex        =   17
         Top             =   5265
         Width           =   3945
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   "Retirement Ages"
         ForeColor       =   &H80000008&
         Height          =   840
         Left            =   3960
         TabIndex        =   37
         Top             =   3720
         Width           =   1395
         Begin VB.TextBox txtMRet 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   90
            TabIndex        =   9
            Top             =   420
            Width           =   570
         End
         Begin VB.TextBox txtFRet 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   750
            TabIndex        =   10
            Top             =   420
            Width           =   540
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Female"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   765
            TabIndex        =   39
            Top             =   210
            Width           =   510
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Male"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   90
            TabIndex        =   38
            Top             =   195
            Width           =   330
         End
      End
      Begin VB.Frame FraApp 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   3495
         TabIndex        =   34
         Top             =   4590
         Visible         =   0   'False
         Width           =   2595
         Begin VB.TextBox txtStart 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1470
            TabIndex        =   15
            Top             =   300
            Width           =   945
         End
         Begin VB.TextBox txtId 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   60
            TabIndex        =   14
            Top             =   300
            Width           =   1290
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "ID Initials"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   60
            TabIndex        =   36
            Top             =   90
            Width           =   675
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Start From"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1455
            TabIndex        =   35
            Top             =   75
            Width           =   765
         End
      End
      Begin VB.CheckBox chkAppcode 
         Appearance      =   0  'Flat
         Caption         =   "Generate Employee Code"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   150
         TabIndex        =   13
         Top             =   4800
         Width           =   2715
      End
      Begin VB.CheckBox chkSRes 
         Appearance      =   0  'Flat
         Caption         =   "1024 By 768 Screen Resolution"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   150
         TabIndex        =   12
         Top             =   4395
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.TextBox txtDGroup 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4815
         TabIndex        =   2
         Top             =   1245
         Width           =   1095
      End
      Begin VB.CommandButton cmdFrom 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5910
         Picture         =   "frmGenOpt.frx":097A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1245
         Width           =   300
      End
      Begin VB.CheckBox chkPSave 
         Appearance      =   0  'Flat
         Caption         =   "Prompt before saving any record"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   150
         TabIndex        =   11
         Top             =   3960
         Width           =   3975
      End
      Begin VB.TextBox txtCName 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   420
         Width           =   9675
      End
      Begin VB.TextBox txtDPass 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4815
         TabIndex        =   1
         Top             =   885
         Width           =   1395
      End
      Begin MSComDlg.CommonDialog Cdl 
         Left            =   3390
         Top             =   765
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame fraLMarks 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   3765
         TabIndex        =   33
         Top             =   3405
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Prompt for casuals to expire within .... months"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   44
         Top             =   3090
         Width           =   3345
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Prompt for visas to expire within .... months"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   43
         Top             =   2385
         Width           =   3180
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Prompt for contracts to expire within .... months"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   42
         Top             =   2745
         Width           =   3495
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Prompt for employees birthdays within .... months"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   41
         Top             =   2025
         Width           =   3615
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Prompt for employees to retire within .... months"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   40
         Top             =   1665
         Width           =   3525
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Company Name"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   165
         Width           =   1125
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Default Employees passwords"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   27
         Top             =   945
         Width           =   2145
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Default users groups"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   26
         Top             =   1260
         Width           =   1500
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   30
      TabIndex        =   29
      Top             =   3180
      Width           =   9855
      Begin VB.CommandButton cmdEdit 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   105
         Picture         =   "frmGenOpt.frx":0A7C
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Edit Record"
         Top             =   60
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.CommandButton cmdSave 
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8775
         Picture         =   "frmGenOpt.frx":0B7E
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Save Record"
         Top             =   60
         Width           =   525
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9285
         Picture         =   "frmGenOpt.frx":0C80
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Cancel Process"
         Top             =   60
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmGenOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub disabletxt()
Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
            i.Locked = True
        End If
    Next i

    
End Sub

Private Sub Check1_Click()

End Sub

Private Sub chkAppcode_Click()
    If chkAppcode.value = 1 Then
        FraApp.Visible = True
    Else
        FraApp.Visible = False
    End If
    
End Sub

Private Sub chkCapital_Click()
If chkCapital.value = 1 Then txtCName.Text = UCase(txtCName.Text)
End Sub

Private Sub cmdAddETerms_Click()
    Dim li As ListItem
    On Error GoTo ErrHandler
    Set rs4 = CConnect.GetRecordSet("SELECT * FROM empTerms order by Code")
    lsvEmpTerms.View = lvwReport
    lsvEmpTerms.FullRowSelect = True
    lsvEmpTerms.LabelEdit = lvwManual
    lsvEmpTerms.ListItems.Clear
    lsvEmpTerms.ColumnHeaders.Clear
    lsvEmpTerms.ColumnHeaders.add , , "Code", 700
    lsvEmpTerms.ColumnHeaders.add , , "Description", lsvEmpTerms.Width - 700
    While rs4.EOF = False
        Set li = lsvEmpTerms.ListItems.add(, , rs4!Code & "")
        li.ListSubItems.add , , rs4!Description & ""
        rs4.MoveNext
    Wend
    fraEmpTerms.Visible = True
    Exit Sub
ErrHandler:
End Sub

Private Sub cmdAddT_Click()
    On Error Resume Next
    If txtTerms.Text <> "" Then
        CConnect.ExecuteSql ("INSERT INTO EmpTerms (Description,UUser,TTime) VALUES ('" & txtTerms.Text & "','" & CurrentUser & "','" & Now & "')")
    End If
    Call cmdAddETerms_Click
End Sub

Private Sub cmdBonus_Click()
    Dim MyMonth As Double

    On Error GoTo ErrorHandler
    
    If MsgBox("This will calculate yearly bonus for the current year. Do you wish to continue?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    
    Set rs6 = CConnect.GetRecordSet("SELECT * FROM Bonus ORDER BY Floor")

    With rsGlob2
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                If !BasicPay > 0 And Not IsNull(!DEmployed) Then
                    If Year(!DEmployed) <> Year(Date) Then
                        With rs6
                            If .RecordCount > 0 Then
                                .MoveLast
                                .Filter = 0
                                .Filter = "Floor<=" & rsGlob2!BasicPay & " AND ceiling>=" & rsGlob2!BasicPay
                                If .RecordCount > 0 Then
                                    CConnect.ExecuteSql ("UPDATE employees SET Annualbonus = " & (rs6!Perc / 100 * rsGlob2!BasicPay) & ", bperc = " & rs6!Perc & " WHERE employeeid = '" & rsGlob!Employee_ID & "'")
                                End If
                            End If
                        End With
                    
    '                Else
    '                    Dim NoofDays As Integer
    '
    '                    NoofDays = DateDiff("d", !DEmployed, "31/12/" & Year(Date))
    '
    '                    If Day(!DEmployed) = 1 Then
    '
    '
    '                        MyMonth = 12 - Month(!DEmployed) + 1 'For those employed on the 1st of the month.
    '                    Else
    '                        MyMonth = (NoofDays / 30) 'MyMonth = 12 - Month(!DEmployed)
    '                    End If
    '                    With rs6
    '                        If .RecordCount > 0 Then
    '                            .MoveFirst
    '                            Do While Not .EOF
    '                                If !code < MyMonth And !Description >= MyMonth Then
    ''''                                    rsGlob2!Bonus = rs6!Perc / 100 * rsGlob2!BasicPay
    ''''                                    rsGlob2!BPerc = rs6!Perc
    ''''                                    rsGlob2.Update
    '                                    .Filter = 0
    '                                    .Filter = "Floor<=" & rsGlob2!BasicPay & " AND ceiling>=" & rsGlob2!BasicPay
    '                                    CConnect.ExecuteSql ("UPDATE employee SET bonus = " & (rs6!Perc / 100 * rsGlob2!BasicPay) & ", bperc = " & rs6!Perc & " WHERE employee_id = '" & rsGlob!employee_id & "'")
    '                                    Exit Do
    '                                End If
    '                                .MoveNext
    '                            Loop
    '
    '                        End If
    '                    End With
                    End If
                End If
            
                .MoveNext
            Loop
        End If
    End With
    
    Me.MousePointer = 0
    
    MsgBox "Process completed succesfully.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while calculating the Annual Bonus" & vbNewLine & err.Description, vbInformation, TITLES
    Me.MousePointer = vbDefault
End Sub

Public Sub cmdCancel_Click()
    fraGenOpt.Enabled = False
    fraGrp.Visible = False
    cmdSave.Enabled = False
    cmdEdit.Enabled = True
    cmdCancel.Enabled = False
    Call DisplayRecords
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdCancelDRC_Click()
    fraDRC.Visible = False
End Sub

Private Sub cmdChange_Click()
On Error GoTo ErrHandler
If txtFromE.Text = "" Then
    MsgBox "You must select employee code.", vbInformation
    txtFromE.SetFocus
    Exit Sub
End If

If txtToE.Text = "" Then
    MsgBox "You must enter new employee code.", vbInformation
    txtToE.SetFocus
    Exit Sub
End If

If MsgBox("Are you sure you to change the employee code?", vbQuestion + vbYesNo) = vbYes Then
    CConnect.ExecuteSql ("UPDATE Employee SET EmpCode = '" & txtToE.Text & "' WHERE EmpCode = '" & txtFromE.Text & "'")

    rsGlob.Requery
    rsGlob2.Requery
    
    Call frmMain2.LoadEmployeeList
    
    MsgBox "Employee code changed successfully.", vbInformation
End If
    Exit Sub
ErrHandler:
    MsgBox err.Description, vbInformation
    
End Sub

Private Sub cmdCompanyLogo_Click()
On Error GoTo Hell
If fraCompanyLogo.Visible = True Then
    fraCompanyLogo.Visible = False
    cmdCompanyLogo.Caption = "Set Company &Logo"
Else
    fraCompanyLogo.Visible = True
    cmdCompanyLogo.Caption = "Hide Set &Logo Window"
End If
picLogo.Picture = LoadPicture(Picpath)
Exit Sub
Hell:
End Sub

Private Sub cmdContractExpiryAlert_Click()
If fraEmpPrompt.Visible = False Then
    fraEmpPrompt.Visible = True
Else
    fraEmpPrompt.Visible = False
End If
End Sub

Private Sub cmdCS_Click()
    fraGrp.Visible = False
End Sub

Private Sub cmdDeleteT_Click()
    On Error GoTo ErrHandler
        CConnect.ExecuteSql ("DELETE FROM empTerms WHERE code = '" & lsvEmpTerms.SelectedItem.Text & "'")
        Call cmdAddETerms_Click
    Exit Sub
ErrHandler:
End Sub

Public Sub cmdEdit_Click()
    fraGenOpt.Enabled = True
    cmdSave.Enabled = True
    cmdEdit.Enabled = False
    cmdCancel.Enabled = True
    
    Omnis_ActionTag = "E"   '++Caters for the omnis text writer - monte++
    
End Sub

Private Sub cmdExitT_Click()
    fraEmpTerms.Visible = False
End Sub

Private Sub cmdExpiryCancel_Click()
fraEmpPrompt.Visible = False
End Sub

Private Sub cmdExpiryOk_Click()
On Error GoTo ErrHandler
If Not IsNumeric(txtExpiryMonth.Text) Then MsgBox "Please provide a numeric value.", vbOKOnly + vbExclamation, "Wrong values": fraEmpPrompt.Visible = False: Exit Sub
    CConnect.ExecuteSql "UPDATE generalopt SET ExpiryPrompt = " & txtExpiryMonth.Text & " WHERE subsystem = '" & SubSystem & "'"
    MsgBox "Employment expiry prompt settings updated.", vbInformation
    fraEmpPrompt.Visible = False
    Exit Sub
ErrHandler:
    MsgBox err.Description, vbInformation
End Sub

Private Sub cmdFrom_Click()
    fraGrp.Visible = True
    lvwGrp.Visible = True
''    cmdCS.Visible = False
End Sub

Private Sub cmdFromE_Click()
    Sel = ""
    popupText = "ChangeCode"
    frmPopUp.Show vbModal
    
End Sub

Public Sub cmdNew_Click()
fraGenOpt.Enabled = True
    cmdSave.Enabled = True
    cmdEdit.Enabled = False
    cmdCancel.Enabled = True
    
    Omnis_ActionTag = "E"
End Sub

Private Sub cmdOKDRC_Click()
    On Error GoTo ErrHandler
    CConnect.ExecuteSql "UPDATE generalopt SET servername = '" & txtServer.Text & "', connectionName ='" & txtConnection.Text & "', dcatalog = '" & txtCatalog.Text & "', userid ='" & txtUID.Text & "', passwd='" & txtPass & "' WHERE subsystem = '" & SubSystem & "'"
    MsgBox "Report connections updated.", vbInformation
    fraDRC.Visible = False
    Exit Sub
ErrHandler:
    MsgBox err.Description, vbInformation
End Sub

Private Sub cmdRConnections_Click()
    Dim rec_d As New ADODB.Recordset
    On Error GoTo ErrHandler
    Set rec_d = CConnect.GetRecordSet("SELECT * FROM generalopt WHERE subsystem = '" & SubSystem & "'")
    With rec_d
        If rec_d.EOF = False Then
            txtConnection.Text = !connectionName & ""
            txtServer.Text = !ServerName & ""
            txtCatalog.Text = !dcatalog & ""
            txtUID.Text = !UserID & ""
        End If
    End With
    fraDRC.Visible = True
    Exit Sub
ErrHandler:
MsgBox err.Description, vbInformation
End Sub

Public Sub cmdSave_Click()

'Company_TextFile    '++Create the Omnis Company Text file 'monte++

'If txtDGroup.Text = "" Then
'    MsgBox "You must select the default group", vbInformation
'    txtDGroup.SetFocus
'    Exit Sub
'End If
'If txtDPass.Text = "" Then
'    MsgBox "You must enter the default password", vbInformation
'    txtDPass.SetFocus
'    Exit Sub
'End If

'If txtDBase.Text = "" Then
'    MsgBox "You must enter the Employee data source name.", vbInformation
'    txtDBase.SetFocus
'    Exit Sub
'End If

If txtMRet.Text = "" Then
    MsgBox "You must enter the retirement age for Male employees.", vbInformation
    txtMRet.SetFocus
    Exit Sub
End If

If txtFRet.Text = "" Then
    MsgBox "You must enter the retirement age for Female employees.", vbInformation
    txtFRet.SetFocus
    Exit Sub
End If

myYear = Year(Date)

If chkAppcode.value = 1 Then
    If txtId.Text = "" Then
        MsgBox "You must enter the application code initials", vbInformation
        txtId.SetFocus
        Exit Sub
    End If
End If

'If chkDis.Value = 1 Then
'    If txtDis.Text = "" Then
'        MsgBox "You must enter the database connection to disciplinary system", vbInformation
'        txtDis.SetFocus
'        Exit Sub
'    End If
'End If

'If chkRec.Value = 1 Then
'    If txtRec.Text = "" Then
'        MsgBox "You must enter the database connection to recruitment system", vbInformation
'        txtRec.SetFocus
'        Exit Sub
'    End If
'End If

If txtRetireDays.Text = "" Then
    txtRetireDays.Text = 0
End If

If txtBirthdayDays.Text = "" Then
    txtBirthdayDays.Text = 0
End If

If txtVisa.Text = "" Then
    txtVisa.Text = 0
End If

If txtCasuals.Text = "" Then
    txtCasuals.Text = 0
End If

If txtContracts.Text = "" Then
    txtContracts.Text = 0
End If

'If txtPDays.Text = "" Then
'    MsgBox "You must enter days for password to expire.", vbInformation
'    txtPDays.SetFocus
'    Exit Sub
'End If


With rs
    If .RecordCount > 0 Then
        If PromptSave = True Then
            If MsgBox("Are you sure you want to save the records?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        
        !cName = txtCName.Text & ""
        !DPass = txtDPass.Text & ""
        !DGNo = txtDGroup.Text & ""
        !DBase = txtDBase.Text & ""
        !FRet = txtFRet.Text & ""
        !MRet = txtMRet.Text & ""
        !RDays = txtRetireDays.Text & ""
        !BDays = txtBirthdayDays.Text & ""
        !VisasP = txtVisa.Text & ""
        !CasualsP = txtCasuals.Text & ""
        !ContractsP = txtContracts.Text & ""
        !PDays = txtPDays.Text & ""
        
  
        If chkCPass.value = 1 Then
            !CPass = "Yes"
'            !DisBase = txtDis.Text & ""
        Else
            !CPass = "No"
'            !DisBase = ""
        End If
        
        If chkVSal.value = 1 Then
            !VSal = "Yes"
            ViewSal = True
        Else
            !VSal = "No"
            ViewSal = False
        End If

        
        If chkDis.value = 1 Then
            !Dis = "Yes"
            !DisBase = txtDis.Text & ""
        Else
            !Dis = "No"
            !DisBase = ""
        End If
        
        If chkRec.value = 1 Then
            !Recruit = "Yes"
            !RecBase = txtRec.Text & ""
        Else
            !Recruit = "No"
            !RecBase = ""
        End If
   
        
        If chkPSave.value = 1 Then
            !PSave = "Yes"
            PromptSave = True
        Else
            !PSave = "No"
            PromptSave = False
        End If
        If chkAppcode.value = 1 Then
            !GenID = "Yes"
            !StartFrom = txtStart.Text & ""
            !IDInitials = txtId.Text & ""
        Else
            !GenID = "No"
        End If
                
        If optODBase = True Then
            !DSource = "Omnis"
        ElseIf optPDBase = True Then
            !DSource = "Payroll"
        Else
            !DSource = "Local"
        End If
        

        .Update
    End If
End With

Set rsGenOpt = CConnect.GetRecordSet("SELECT * FROM GeneralOpt")

Call LoadVar

fraGenOpt.Enabled = False
cmdSave.Enabled = False
cmdEdit.Enabled = True
cmdCancel.Enabled = False

End Sub

Private Sub cmdSavePic_Click()
Dim RsT As New ADODB.Recordset
On Error GoTo Hell
CConnect.ExecuteSql "Delete from CompanyLogo"
CConnect.ExecuteSql "INSERT INTO CompanyLogo(Path) VALUES('" & Replace(dirDirectory.Path & "\" & flFile.FileName, "'", "''") & "')"
Picpath = Replace(dirDirectory.Path & "\" & flFile.FileName, "'", "''")
picLogo.Picture = LoadPicture(Replace(Picpath, "''", "'"))
MsgBox "Logo successfully saved.", vbOKOnly + vbInformation, "Saved"
cmdSavePic.Enabled = False
Exit Sub
Hell:
End Sub

Private Sub cmdSelect_Click()
    txtDGroup.Text = lvwGrp.SelectedItem
    fraGrp.Visible = False
End Sub

Private Sub cmdSelect1_Click()
        txtALCode.Text = LvwLeaves.SelectedItem
        fraGrp.Visible = False
End Sub

Private Sub dirDirectory_Change()
flFile = dirDirectory
End Sub

Private Sub drvDrive_Change()
On Error GoTo ErrHandler
dirDirectory = drvDrive.Drive
Exit Sub
ErrHandler:
MsgBox err.Description, vbExclamation + vbOKOnly, "Drive"
End Sub

Private Sub flFile_DblClick()
On Error GoTo ErrMsg
picLogo.Picture = LoadPicture(dirDirectory.Path & "\" & flFile.FileName)
cmdSavePic.Enabled = True
Exit Sub
ErrMsg:
MsgBox "Invalid picture.", vbExclamation + vbOKOnly, "Invalid"
cmdSavePic.Enabled = False
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    oSmart.FReset Me
    
    frmMain2.txtDetails.Caption = ""

    If oSmart.hRatio > 1.1 Then
        With frmMain2
            Me.Move .tvwMain.Width + .tvwMain.Width / 32.75, (.Height / 5.52) ' - 155
        End With
    Else
         With frmMain2
            Me.Move .tvwMain.Width + .tvwMain.Width / 32.75, .Height / 5.55
        End With
        
    End If

    CConnect.CColor Me, MyColor
    
    
    Set rs = CConnect.GetRecordSet("SELECT * FROM GeneralOpt WHERE subsystem = '" & SubSystem & "'")
    
    Call DisplayRecords

    Call InitGrid
    Call LoadList
    
    fraGenOpt.Enabled = False
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    
    Exit Sub

ErrorHandler:

End Sub


Public Sub DisplayRecords()
    With rs
        If .RecordCount > 0 Then
            txtCName.Text = !cName & ""
            txtDPass.Text = !DPass & ""
            txtDGroup.Text = !DGNo & ""
            txtDBase.Text = !DBase & ""
            txtFRet.Text = !FRet & ""
            txtMRet.Text = !MRet & ""
            txtBirthdayDays.Text = !BDays & ""
            txtRetireDays.Text = !RDays & ""
            txtVisa.Text = !VisasP & ""
            txtCasuals.Text = !CasualsP & ""
            txtContracts.Text = !ContractsP & ""
            txtPDays.Text = !PDays & ""
                                  
            If !Dis = "Yes" Then
                Dis = "Yes"
                chkDis.value = 1
                txtDis.Text = !DisBase & ""
               
            Else
                Dis = "No"
                chkDis.value = 0
            End If
                        
            If !Recruit = "Yes" Then
                Recruit = "Yes"
                chkRec.value = 1
                txtRec.Text = !RecBase & ""
               
            Else
                Recruit = "No"
                chkRec.value = 0
            End If
            
            If !CPass = "Yes" Then
                chkCPass.value = 1
            Else
                chkCPass.value = 0
            End If
            
            If !VSal = "Yes" Then
                chkVSal.value = 1
            Else
                chkVSal.value = 0
            End If
            
            If !PSave = "Yes" Then
                chkPSave.value = 1
            Else
                chkPSave.value = 0
            End If
            
            If !GenID = "Yes" Then
                chkAppcode.value = 1
                Call chkAppcode_Click
                txtStart.Text = !StartFrom & ""
                txtId.Text = !IDInitials & ""
            Else
                chkAppcode.value = 0
            End If
                      
            If !DSource = "Omnis" Then
                optODBase = True
                optLDBase = False
                optPDBase = False
            ElseIf !DSource = "Payroll" Then
                optODBase = False
                optLDBase = False
                optPDBase = True
            Else
                optODBase = False
                optLDBase = True
                optPDBase = False
            End If
            
        End If
    End With
End Sub

Private Sub Form_Resize()
    oSmart.FResize Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs = Nothing
'    Set Cnn = Nothing

    frmMain2.Caption = "Infiniti HRMIS - PDR [Current User:\" & currUser.FullNames & "]"
    
End Sub

Private Sub lsvEmpTerms_Click()
    On Error Resume Next
    txtTerms.Text = lsvEmpTerms.SelectedItem.ListSubItems(1).Text
End Sub

Private Sub lsvEmpTerms_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lsvEmpTerms
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvwGrp_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwGrp
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvwGrp_DblClick()
    Call cmdSelect_Click
End Sub

Private Sub LvwLeaves_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With LvwLeaves
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvwLeaves_DblClick()
    cmdSelect1_Click
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error GoTo ErrHandler
Dim ss As String
Dim myfile As String
    Me.MousePointer = vbHourglass
        Set a = New Application
        myfile = App.Path & "\Leave Reports\General Options.rpt"
        Set r = a.OpenReport(myfile)
             
      r.ReadRecords

      With frmReports.CRViewer1
          .ReportSource = r
          .ViewReport
      End With

      frmReports.Show vbModal
      Me.MousePointer = 0
Exit Sub

ErrHandler:
If err.Description = "File not found." Then
    Cdl.DialogTitle = "Select the report to show"
    Cdl.InitDir = App.Path & "/Leave Reports"
    Cdl.Filter = "Reports {* .rpt|* .rpt"
    Cdl.ShowOpen
    myfile = Cdl.FileName
    If Not myfile = "" Then
        Resume
    Else
        Me.MousePointer = 0
    End If
Else
    MsgBox err.Description, vbInformation
    Me.MousePointer = 0
End If
End Sub

Private Sub txtCName_Change()
If chkCapital.value = 1 Then
    txtCName.Text = UCase(txtCName.Text)
    txtCName.SelStart = Len(txtCName.Text)
End If
End Sub

Private Sub txtCName_KeyPress(KeyAscii As Integer)
    If Len(Trim(txtCName.Text)) > 199 Then
        Beep
        MsgBox "Can't enter more than 200 characters", vbExclamation
        KeyAscii = 8
    End If
    
    Select Case KeyAscii
      Case Asc("0") To Asc("9")
      Case Asc("A") To Asc("Z")
      Case Asc("a") To Asc("z")
      Case Asc(" ")
      Case Asc("/")
      Case Asc("\")
      Case Asc("?")
      Case Asc(":")
      Case Asc(";")
      Case Asc(",")
      Case Asc("-")
      Case Asc("(")
      Case Asc(")")
      Case Asc("&")
      Case Asc(".")
      Case Is = 8
      Case Else
          Beep
          KeyAscii = 0
    End Select
End Sub

Private Sub txtCYEnd_KeyPress(KeyAscii As Integer)
If Len(Trim(txtCYEnd.Text)) > 8 Then
    Beep
    MsgBox "Can't enter more than 9 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("/")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select

End Sub

Private Sub txtCYEnd_LostFocus()
   txtCYEnd.Text = Format(txtCYEnd.Text, "dd/mm")
    If Not IsDate(txtCYEnd.Text) Or txtCYEnd.Text = "29/02" Then
        MsgBox "Please enter a valid date", vbInformation
        txtCYEnd.Text = ""
        txtCYEnd.SetFocus
    End If
    
End Sub

Public Sub LoadList()
Set rs2 = CConnect.GetRecordSet("SELECT * FROM tblUserGroup WHERE subsystem = '" & SubSystem & "'")
With rs2
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
            Set li = lvwGrp.ListItems.add(, , !GROUP_CODE & "")
                li.ListSubItems.add , , !GROUP_NAME & ""
                .MoveNext
        Loop
    End If
End With
Set rs2 = Nothing


End Sub

Public Sub InitGrid()
With lvwGrp
    .ColumnHeaders.Clear
    .ColumnHeaders.add , , "ID"
    .ColumnHeaders.add , , "Group Name", 3000
    
    lvwGrp.View = lvwReport
End With

With LvwLeaves
    .ColumnHeaders.Clear
    .ColumnHeaders.add , , "Code"
    .ColumnHeaders.add , , "Leave Name", 2000
    
    LvwLeaves.View = lvwReport
End With
End Sub

Private Sub txtDis_KeyPress(KeyAscii As Integer)
    If Len(Trim(txtDis.Text)) > 19 Then
        Beep
        MsgBox "Can't enter more than 20 characters", vbExclamation
        KeyAscii = 8
    End If
    
    Select Case KeyAscii
      Case Asc("0") To Asc("9")
      Case Asc("A") To Asc("Z")
      Case Asc("a") To Asc("z")
      Case Asc(":")
      Case Asc(";")
      Case Asc(",")
      Case Asc("-")
      Case Asc("(")
      Case Asc(")")
      Case Asc(".")
      Case Is = 8
      Case Else
          Beep
          KeyAscii = 0
    End Select

End Sub

Private Sub txtFRet_KeyPress(KeyAscii As Integer)
    If Len(Trim(txtFRet.Text)) > 2 Then
        Beep
        MsgBox "Can't enter more than 2 characters", vbExclamation
        KeyAscii = 8
    End If
    
    Select Case KeyAscii
      Case Asc("0") To Asc("9")

      Case Is = 8
      Case Else
          Beep
          KeyAscii = 0
    End Select
End Sub

Private Sub txtFromE_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtId_Change()
    txtId.Text = UCase(txtId.Text)
    txtId.SelStart = Len(txtId.Text)
    
End Sub

Private Sub LoadVar()
Set rs1 = CConnect.GetRecordSet("SELECT * FROM GeneralOpt")

    With rs1
        If .RecordCount > 0 Then
            If !PSave = "Yes" Then PromptSave = True
            DSource = !DSource & ""
            
            If IsNull(!DBase) Then
                MsgBox "Emloyee Data source has not been specified therefor the local database will be used.", vbExclamation
                DSource = "Local"
            Else
                CConnect.EDBase = !DBase & ""
            End If
            
            
            If Not IsNull(!PRate) Then
                PRate = !PRate
            Else
                PRate = 0
            End If
            
'            Dis = !Dis & ""
'            CConnect.DisBase = !DisBase & ""
            
'            Recruit = !Recruit & ""
'            CConnect.RecBase = !RecBase & ""
            
            EmpGroup = !DGNo & ""
            DPass = !DPass & ""
            IDiv = !IDiv & ""
            CPass = !CPass & ""

        
         
        End If
    End With
    
Set rs1 = Nothing

End Sub






Private Sub txtDBase_KeyPress(KeyAscii As Integer)
If Len(Trim(txtDBase.Text)) > 49 Then
    Beep
    MsgBox "Can't enter more than 50 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("A") To Asc("Z")
  Case Asc("a") To Asc("z")
  Case Asc(" ")
  Case Asc("/")
  Case Asc("\")
  Case Asc("?")
  Case Asc(":")
  Case Asc(";")
  Case Asc(",")
  Case Asc("-")
  Case Asc("(")
  Case Asc(")")
  Case Asc("&")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub

Private Sub txtDGroup_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtDPass_KeyPress(KeyAscii As Integer)
If Len(Trim(txtDPass.Text)) > 49 Then
    Beep
    MsgBox "Can't enter more than 50 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("A") To Asc("Z")
  Case Asc("a") To Asc("z")
  Case Asc(" ")
  Case Asc("/")
  Case Asc("\")
  Case Asc("?")
  Case Asc(":")
  Case Asc(";")
  Case Asc(",")
  Case Asc("-")
  Case Asc("(")
  Case Asc(")")
  Case Asc("&")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub


Private Sub txtHigh_KeyPress(KeyAscii As Integer)
If Len(Trim(txtHigh.Text)) > 4 Then
    Beep
    MsgBox "Can't enter more than 5 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("-")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub

Private Sub txtLow_KeyPress(KeyAscii As Integer)
If Len(Trim(txtLow.Text)) > 4 Then
    Beep
    MsgBox "Can't enter more than 5 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("-")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub

Private Sub txtId_KeyPress(KeyAscii As Integer)
If Len(Trim(txtId.Text)) > 15 Then
    Beep
    MsgBox "Can't enter more than 15 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("A") To Asc("Z")
  Case Asc("a") To Asc("z")
  Case Asc(" ")
  Case Asc("/")
  Case Asc("\")
  Case Asc("?")
  Case Asc(":")
  Case Asc(";")
  Case Asc(",")
  Case Asc("-")
  Case Asc("(")
  Case Asc(")")
  Case Asc("&")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub

Private Sub txtMRet_KeyPress(KeyAscii As Integer)
    If Len(Trim(txtMRet.Text)) > 2 Then
        Beep
        MsgBox "Can't enter more than 2 characters", vbExclamation
        KeyAscii = 8
    End If
    
    Select Case KeyAscii
      Case Asc("0") To Asc("9")

      Case Is = 8
      Case Else
          Beep
          KeyAscii = 0
    End Select
End Sub

Private Sub txtRec_KeyPress(KeyAscii As Integer)
    If Len(Trim(txtRec.Text)) > 19 Then
        Beep
        MsgBox "Can't enter more than 20 characters", vbExclamation
        KeyAscii = 8
    End If
    
    Select Case KeyAscii
      Case Asc("0") To Asc("9")
      Case Asc("A") To Asc("Z")
      Case Asc("a") To Asc("z")
      Case Asc(":")
      Case Asc(";")
      Case Asc(",")
      Case Asc("-")
      Case Asc("(")
      Case Asc(")")
      Case Asc(".")
      Case Is = 8
      Case Else
          Beep
          KeyAscii = 0
    End Select
    
End Sub

Private Sub txtStart_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Is = 8
  Case Else
    Beep
    KeyAscii = 0
End Select
End Sub
