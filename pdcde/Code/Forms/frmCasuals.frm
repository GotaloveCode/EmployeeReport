VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmCasuals 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Casuals Details"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   Icon            =   "frmCasuals.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imgEmpTool 
      Left            =   5430
      Top             =   195
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
            Picture         =   "frmCasuals.frx":0442
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCasuals.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCasuals.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCasuals.frx":0778
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7900
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   7800
      Begin MSComDlg.CommonDialog Cdl 
         Left            =   5610
         Top             =   4560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Move to the First employee"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "CLOSE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6270
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5400
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   9
         ToolTipText     =   "Move to the Previous employee"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   10
         ToolTipText     =   "Move to the Next employee"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
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
         Left            =   2775
         Picture         =   "frmCasuals.frx":0CBA
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Add New record"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
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
         Left            =   3255
         Picture         =   "frmCasuals.frx":0DBC
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Edit Record"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdDelete 
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
         Left            =   3735
         Picture         =   "frmCasuals.frx":0EBE
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Delete Record"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   11
         ToolTipText     =   "Move to the Last employee"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame fraDetails 
         Appearance      =   0  'Flat
         Caption         =   "Casuals"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3555
         Left            =   645
         TabIndex        =   17
         Top             =   840
         Visible         =   0   'False
         Width           =   6015
         Begin VB.CheckBox chkIsActive 
            Appearance      =   0  'Flat
            Caption         =   "&Make this the active casual definition"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2640
            TabIndex        =   26
            Top             =   1680
            Width           =   3195
         End
         Begin VB.CommandButton cmdSearch 
            Height          =   315
            Left            =   1320
            Picture         =   "frmCasuals.frx":13B0
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   600
            Width           =   315
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
            Left            =   5370
            Picture         =   "frmCasuals.frx":173A
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Cancel Process"
            Top             =   2910
            Width           =   495
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            Default         =   -1  'True
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
            Left            =   4890
            Picture         =   "frmCasuals.frx":183C
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Save Record"
            Top             =   2910
            Width           =   495
         End
         Begin VB.TextBox txtComments 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   810
            Left            =   135
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   1935
            Width           =   5730
         End
         Begin VB.TextBox txtReff 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   3450
            TabIndex        =   4
            Top             =   1140
            Width           =   2370
         End
         Begin VB.TextBox txtCode 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   135
            TabIndex        =   0
            Top             =   600
            Width           =   1140
         End
         Begin VB.TextBox txtDesc 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1800
            TabIndex        =   1
            Top             =   600
            Width           =   4020
         End
         Begin MSComCtl2.DTPicker dtpFrom 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   330
            Left            =   135
            TabIndex        =   2
            Top             =   1125
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd, MMM, yyyy"
            Format          =   63045635
            CurrentDate     =   37673
            MinDate         =   21916
         End
         Begin MSComCtl2.DTPicker dtpTo 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   330
            Left            =   1800
            TabIndex        =   3
            Top             =   1125
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd, MMM, yyyy"
            Format          =   63045635
            CurrentDate     =   37673
            MinDate         =   21916
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "To"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1800
            TabIndex        =   24
            Top             =   900
            Width           =   180
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Comments"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   23
            Top             =   1590
            Width           =   750
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Ref No."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3450
            TabIndex        =   22
            Top             =   900
            Width           =   555
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Code"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1800
            TabIndex        =   19
            Top             =   360
            Width           =   795
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            Caption         =   "From"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   18
            Top             =   900
            Width           =   360
         End
      End
      Begin MSComctlLib.ListView lvwDetails 
         Height          =   7800
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   7800
         _ExtentX        =   13758
         _ExtentY        =   13758
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgTree"
         ForeColor       =   0
         BackColor       =   -2147483643
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
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmCasuals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Sub cmdCancel_Click()
'    If PSave = False Then
'        If MsgBox("Do you want to save changes?", vbQuestion + vbYesNo) = vbYes Then  '
'            Call cmdSave_Click
'            Exit Sub
'        End If
'    End If
'
'    Call DisplayRecords
'    fraDetails.Visible = False
'
    Call EnableCmd
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    SaveNew = False
'
'    With frmMain2
'        .cmdNew.Enabled = True
'        .cmdEdit.Enabled = True
'        .cmdDelete.Enabled = True
'        .cmdCancel.Enabled = False
'        .cmdSave.Enabled = False
'    End With
If PromptSave = True Then
    If MsgBox("Close this window?", vbYesNo + vbQuestion, "Confirm Close") = vbNo Then Exit Sub
End If
fraDetails.Visible = False
With frmMain2
    .cmdNew.Enabled = True
    .cmdEdit.Enabled = True
    .cmdDelete.Enabled = True
    .cmdCancel.Enabled = False
    .cmdSave.Enabled = False
End With
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Public Sub cmdDelete_Click()
    Dim resp As String
    On Error GoTo ErrHandler
    'check rights
    If Not currUser Is Nothing Then
        If currUser.CheckRight("Casual") <> secModify Then
            MsgBox "You dont have right to delete the record. Please liaise with the security admin"
            Exit Sub
        End If
     End If
    If SelectedEmployee Is Nothing Then
        MsgBox "Select Employee", vbInformation, "Inform"
        Exit Sub
        
    End If
    
    If lvwDetails.ListItems.count > 0 Then
        resp = MsgBox("Are you sure you want to delete  " & lvwDetails.SelectedItem & " from the records?", vbQuestion + vbYesNo)
        If resp = vbNo Then
            Exit Sub
        End If
        Action = "DELETED CASUAL DATA; EMPLOYEE CODE: " & SelectedEmployee.EmpCode & "; CASUAL CODE: " & lvwDetails.SelectedItem & "; CASUAL DESCRIPTION: " & lvwDetails.SelectedItem.ListSubItems(1)
        CConnect.ExecuteSql ("DELETE FROM Casuals WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & lvwDetails.SelectedItem & "'")
        rs2.Requery
        Call DisplayRecords
    Else
        MsgBox "You have to select the contract you would like to delete.", vbInformation
                
    End If
    
    Exit Sub
ErrHandler:
End Sub

Private Sub cmdDone_Click()
    PSave = True
    Call cmdCancel_Click
    PSave = False
    
End Sub

Public Sub cmdEdit_Click()
    On Error GoTo ErrHandler
    
     If Not currUser Is Nothing Then
        If currUser.CheckRight("CasualTypes") <> secModify Then
            MsgBox "You dont have right to modify record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    If SelectedEmployee Is Nothing Then
        MsgBox "Please select employee"
        Exit Sub
    End If
    
    If lvwDetails.ListItems.count < 1 Then
        MsgBox "You have to select the casual you would like to edit.", vbInformation
        PSave = True
        Call cmdCancel_Click
        PSave = False
        Exit Sub
    End If
    Set rs3 = CConnect.GetRecordSet("SELECT * FROM Casuals WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & lvwDetails.SelectedItem & "'")
    
    With rs3
        If .RecordCount > 0 Then
            txtCode.Text = !Code & ""
            txtCode.Tag = Trim(!ID & "")
            txtDesc.Locked = True
            Dim rsCont As New ADODB.Recordset
            Set rsCont = CConnect.GetRecordSet("select * from pdCasualTypes where Code='" & txtCode.Text & "'")
            If rsCont.RecordCount > 0 Then
                strDatePart = IIf(Trim(rsCont!InDays & "") = True, "d", IIf(Trim(rsCont!InWeeks & "") = True, "w", IIf(Trim(rsCont!InMonths & "") = True, "m", IIf(Trim(rsCont!InYears & "") = True, "y", "u"))))
                strValue = IIf(IsNumeric(Trim(rsCont!correspondingvalue & "")) = True, Trim(rsCont!correspondingvalue & ""), 0)
            End If
            txtDesc.Text = !Description & ""
            If Not IsNull(!cFrom) Then dtpFrom.value = !cFrom & ""
            If Not IsNull(!cTo) Then dtpTo.value = !cTo & ""
            txtComments.Text = !Comments & ""
            txtReff.Text = !Ref & ""
            Me.Tag = Trim(SelectedEmployee.EmploymentTerm.EmpTermID & "")
            strNamePart = Trim(SelectedEmployee.EmploymentTerm.EmpTermName & "")
            chkIsActive.value = IIf(Trim(!isactive & "") = True, 1, 0)
            SaveNew = False
        Else
            MsgBox "Record not found.", vbInformation
            Set rs3 = Nothing
            PSave = True
            Call cmdCancel_Click
            PSave = False
            Exit Sub
        End If
    End With
    
    Set rs3 = Nothing
    
    Call DisableCmd
    
    fraDetails.Visible = True
    
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    SaveNew = False
    txtCode.Locked = True
    txtDesc.SetFocus
    Exit Sub
ErrHandler:
End Sub


Private Sub cmdFirst_Click()

With rsGlob
    If .RecordCount > 0 Then
        If .BOF <> True Then
            .MoveFirst
            If .BOF = True Then
                .MoveFirst
                Call DisplayRecords
            Else
                Call DisplayRecords
            End If
            
            Call FirstDisb
            
        End If
    End If
End With


End Sub

Private Sub cmdLast_Click()
With rsGlob
    If .RecordCount > 0 Then
        If .EOF <> True Then
            .MoveLast
            If .EOF = True Then
                .MoveLast
                Call DisplayRecords
            Else
                Call DisplayRecords
            End If
            
            Call LastDisb
            
        End If
    End If
End With

End Sub

Public Sub cmdNew_Click()
    On Error GoTo ErrHandler
        
    Dim rsCheckEmpTerms As New ADODB.Recordset
    
     If Not currUser Is Nothing Then
        If currUser.CheckRight("Casual") <> secModify Then
            MsgBox "You dont have right to add new record. Please liaise with the security admin"
            Exit Sub
        End If
     End If
     If SelectedEmployee Is Nothing Then
        MsgBox "select employee"
        Exit Sub
     End If
     
    'Call DisableCmd
    'cmdSearch.Enabled = True
    ''txtCode.Text = loadACode
    'txtDesc.Text = ""
    'txtReff.Text = ""
    'txtComments.Text = ""
    'dtpFrom.Value = Date
    'dtpTo.Value = Date
    'fraDetails.Visible = True
    'cmdCancel.Enabled = True
    'SaveNew = True
    'cmdSave.Enabled = True
    With rsGlob
        .Requery
        
            .Filter = "employee_id=" & SelectedEmployee.EmployeeID
        
        If .RecordCount > 0 Then
            dtpFrom.value = IIf(IsDate(Trim(!DEmployed & "")) = True, Trim(!DEmployed & ""), Date)
            dtpTo.value = dtpFrom.value
            Set rsCheckEmpTerms = CConnect.GetRecordSet("SELECT * FROM Empterms WHERE matchToCasual=1 and description like '" & !Terms & "'")
            If rsCheckEmpTerms.RecordCount = 0 Then MsgBox "This employee is not casual terms." & vbNewLine & "You therefore cannot assign " & IIf(Trim(!Gender & "") = "Male", "him ", "her ") & "a casual detail.", vbOKOnly + vbInformation, "Request denied": Exit Sub
        Else
            dtpFrom.value = Date
            dtpTo.value = Date
        End If
    End With
    Call DisableCmd
    cmdSearch.Enabled = True
    'txtCode.Text = loadACode
    txtDesc.Text = ""
    txtReff.Text = ""
    txtComments.Text = ""
    dtpFrom.value = Date
    dtpTo.value = Date
    fraDetails.Visible = True
    cmdCancel.Enabled = True
    SaveNew = True
    cmdSave.Enabled = True
    
    'txtCode.Locked = True
    txtDesc.SetFocus
    'dtpFrom.Value = Date
    'dtpTo.Value = Date
     Exit Sub
ErrHandler:
        MsgBox err.Description
End Sub

Private Sub cmdNext_Click()
    
With rsGlob
    If .RecordCount > 0 Then
        If .EOF <> True Then
            .MoveNext
            If .EOF = True Then
                .MoveLast
                Call DisplayRecords
            Else
                Call DisplayRecords
            End If

            Call LastDisb

        End If
    End If
End With


End Sub

Private Sub cmdPrevious_Click()

With rsGlob
    If .RecordCount > 0 Then
        If .BOF <> True Then
            .MovePrevious
            If .BOF = True Then
                .MoveFirst
                Call DisplayRecords
            Else
                Call DisplayRecords
            End If
            
            Call FirstDisb
            
        End If
    End If
End With


End Sub

Public Sub cmdSave_Click()
    On Error GoTo ErrHandler
    If txtCode.Text = "" Then
        MsgBox "Enter the contract code.", vbExclamation
        txtCode.SetFocus
        Exit Sub
    End If
    
    If txtDesc.Text = "" Then
        MsgBox "Enter the contract description.", vbExclamation
        txtDesc.SetFocus
        Exit Sub
    End If
    
    If dtpFrom.value > dtpTo.value Then
        MsgBox "Enter the valid contract start and end dates.", vbInformation
        dtpFrom.SetFocus
        Exit Sub
    End If
    
   If SelectedEmployee Is Nothing Then
        MsgBox "Please select  employee"
        Exit Sub
   End If

    If SaveNew = True Then
        
        Set rs4 = CConnect.GetRecordSet("SELECT * FROM Casuals WHERE employee_id = " & SelectedEmployee.EmployeeID & " AND Code = '" & txtCode.Text & "'")
        
        With rs4
            If .RecordCount > 0 Then
                MsgBox "Casual code already exists. Enter another one.", vbInformation
                Exit Sub
                txtCode.Text = loadACode
                txtDesc.SetFocus
                Set rs4 = Nothing
                Exit Sub
            End If
        End With
        Set rs4 = Nothing
    End If
       
    If PromptSave = True Then
        If MsgBox("Are you sure you want to save the record.", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    If chkIsActive.value = 1 Then
        CConnect.ExecuteSql ("UPDATE Casuals SET IsActive=0 WHERE employee_id='" & SelectedEmployee.EmployeeID & "'")
    End If
    
    Action = "ADDED CASUAL DATA; EMPLOYEE CODE: " & SelectedEmployee.EmpCode & "; CASUAL CODE: " & txtCode.Text & "; CASUAL DESCRIPTION: " & txtDesc.Text & "; COMMENTS: " & txtComments.Text
    
    CConnect.ExecuteSql ("DELETE FROM Casuals WHERE employee_id = " & SelectedEmployee.EmployeeID & " AND Code = '" & txtCode.Text & "'")
    
'    mySQL = "INSERT INTO Casuals (employee_id, Code, Description, CFrom, CTo, Ref, Comments)" & _
                        " VALUES(" & frmMain2.lvwEmp.SelectedItem.Tag & ",'" & txtCode.Text & "','" & txtDesc.Text & "'," & _
                        "'" & Format(dtpFrom.Value, Dfmt) & "','" & Format(dtpTo.Value, Dfmt) & "','" & txtReff.Text & "','" & txtComments.Text & "')"
    
    mySQL = "INSERT INTO Casuals (employee_id, Code, Description, CFrom, CTo, Ref, Comments,IsActive)" & _
                        " VALUES('" & SelectedEmployee.EmployeeID & "','" & txtCode.Text & "','" & txtDesc.Text & "'," & _
                        "'" & SQLDate(dtpFrom.value) & "','" & SQLDate(dtpTo.value) & "','" & txtReff.Text & "','" & txtComments.Text & "'," & chkIsActive.value & ")"
                        
    CConnect.ExecuteSql (mySQL)
    
    If chkIsActive.value = 1 Then CConnect.ExecuteSql "UPDATE employee SET EmploymentValidThro='" & SQLDate(dtpTo.value) & "',Terms='" & strNamePart & "',TermsID='" & txtCode.Tag & "' WHERE employee_id='" & SelectedEmployee.EmployeeID & "'"
    
    rs2.Requery
    
    'refresh the employee list
    Call RefreshEmployeesCol
    
    If SaveNew = False Then
        PSave = True
        Call cmdCancel_Click
        PSave = False
    Else
        rs2.Requery
        txtCode.Text = loadACode
        txtDesc.SetFocus
    End If
    Call DisplayRecords
    Exit Sub
ErrHandler:
End Sub


Private Sub cmdSearch_Click()
If frmSelCasualTypes.lvwDetails.ListItems.count = 0 Then MsgBox "There are no registered casual types." & vbNewLine & "Please do so to proceed.", vbInformation + vbOKOnly, "Casual types": Exit Sub
frmSelCasualTypes.Show vbModal
With txtCode
    .Text = strcode
    .Tag = strID
End With

txtDesc.Text = strName

Select Case strDatePart
Case "d"
    dtpTo.value = DateAdd("d", strValue, dtpFrom.value)
Case "w"
    dtpTo.value = DateAdd("d", strValue * 7, dtpFrom.value)
Case "m"
    dtpTo.value = DateAdd("m", strValue, dtpFrom.value)
Case "y"
    dtpTo.value = DateAdd("m", strValue * 12, dtpFrom.value)
End Select
End Sub

Private Sub dtpFrom_Change()
Select Case strDatePart
Case "d"
    dtpTo.value = DateAdd("d", strValue, dtpFrom.value)
Case "w"
    dtpTo.value = DateAdd("d", strValue * 7, dtpFrom.value)
Case "m"
    dtpTo.value = DateAdd("m", strValue, dtpFrom.value)
Case "y"
    dtpTo.value = DateAdd("m", strValue * 12, dtpFrom.value)
End Select
End Sub

Private Sub dtpFrom_CloseUp()
Select Case strDatePart
Case "d"
    dtpTo.value = DateAdd("d", strValue, dtpFrom.value)
Case "w"
    dtpTo.value = DateAdd("d", strValue * 7, dtpFrom.value)
Case "m"
    dtpTo.value = DateAdd("m", strValue, dtpFrom.value)
Case "y"
    dtpTo.value = DateAdd("m", strValue * 12, dtpFrom.value)
End Select
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler
    oSmart.FReset Me
    
    
    If oSmart.hRatio > 1.1 Then
        With frmMain2
            Me.Move .tvwMain.Width + .lvwEmp.Width + (.tvwMain.Width / 36) * 2, (.Height / 5.52) ' - 155
        End With
    Else
         With frmMain2
            Me.Move .tvwMain.Width + .lvwEmp.Width + (.tvwMain.Width / 36#) * 2, .Height / 5.55
        End With
        
    End If
    
    CConnect.CColor Me, MyColor
    
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    
    Call InitGrid
    '''Call CConnect.CCon
    
    Set rs2 = Nothing
    Set rs2 = CConnect.GetRecordSet("SELECT * FROM Casuals ORDER BY Code")
    
    With rsGlob
        If .RecordCount < 1 Then
            Call DisableCmd
            Exit Sub
        End If
    End With
    
    Call DisplayRecords
    
    cmdFirst.Enabled = False
    cmdPrevious.Enabled = False
    Exit Sub
ErrHandler:
    MsgBox "an error has occured in " & Me.Name & " Error description " & err.Description
End Sub

Private Sub Form_Resize()
    oSmart.FResize Me
    
    Me.Height = tvwMainheight - 150
    Frame1.Move Frame1.Left, 0, Frame1.Width, tvwMainheight - 150
    lvwDetails.Move lvwDetails.Left, 0, lvwDetails.Width, tvwMainheight - 150
End Sub

Private Sub InitGrid()
    With lvwDetails
        .ColumnHeaders.Clear
        .ColumnHeaders.add , , "Casual Code", 0
        .ColumnHeaders.add , , "Description", 294 * .Width / 1080
        .ColumnHeaders.add , , "Date From", .Width / 8
        .ColumnHeaders.add , , "Date To", .Width / 8
        .ColumnHeaders.add , , "Ref No.", .Width / 6
        .ColumnHeaders.add , , "Active Casual Type", .Width / 9
        .ColumnHeaders.add , , "Comments", .Width / 5
        .View = lvwReport
    End With
End Sub

Public Sub DisplayRecords()
    On Error GoTo ErrHandler
    lvwDetails.ListItems.Clear
    Call Cleartxt
    If SelectedEmployee Is Nothing Then Exit Sub
    With rsGlob
        If Not .EOF And Not .BOF Then
            With rs2
                .Requery
                If .RecordCount > 0 Then
                    .Filter = "employee_id like '" & SelectedEmployee.EmployeeID & "'"
                    If .RecordCount > 0 Then
                        .MoveFirst
                        Do While Not .EOF
                            Set li = lvwDetails.ListItems.add(, , !Code & "", , 5)
                            li.ListSubItems.add , , !Description & ""
                            li.ListSubItems.add , , Format(Trim(!cFrom & ""), "dd, MMM, yyyy")
                            li.ListSubItems.add , , Format(Trim(!cTo & ""), "dd, MMM, yyyy")
                            
                            li.ListSubItems.add , , !Ref & ""
                            li.ListSubItems.add , , IIf(Trim(!isactive & "") = True, "Yes", "")
                            li.ListSubItems.add , , !Comments & ""
                                  
                            .MoveNext
                        Loop
                    End If
                    .Filter = adFilterNone
                End If
            End With
            
        End If
    End With
    Exit Sub
ErrHandler:
    MsgBox "An error has occured when displaying record ERROR DESCRIPTION:  " & err.Description
End Sub



Private Sub Form_Unload(Cancel As Integer)

    Set rs2 = Nothing
    Set rs5 = Nothing
'    Set Cnn = Nothing
    

    frmMain2.Caption = "Infiniti HRMIS - PDR [Current User:\" & currUser.FullNames & "]"
    
End Sub

Private Sub fraList_DragDrop(Source As Control, X As Single, y As Single)

End Sub


Private Sub lvwDetails_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwDetails
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvwDetails_DblClick()
    If frmMain2.cmdEdit.Enabled = True Then
        Call frmMain2.cmdEdit_Click
    End If
End Sub

Private Sub LastDisb()
With rsGlob
    If Not .EOF Then
        .MoveNext
        If .EOF Then
            cmdLast.Enabled = False
            cmdNext.Enabled = False
            cmdPrevious.Enabled = True
            cmdFirst.Enabled = True
            cmdPrevious.SetFocus
        End If
        .MovePrevious
    End If
    
    cmdPrevious.Enabled = True
    cmdFirst.Enabled = True
End With
End Sub


Private Sub FirstDisb()
With rsGlob
    If Not .BOF Then
        .MovePrevious
        If .BOF Then
            cmdLast.Enabled = True
            cmdNext.Enabled = True
            cmdPrevious.Enabled = False
            cmdFirst.Enabled = False
            cmdNext.SetFocus
        End If
        .MoveNext
    End If
    
    cmdLast.Enabled = True
    cmdNext.Enabled = True
End With
End Sub


Private Sub Cleartxt()
Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
            i.Text = ""
        End If
    Next i

    lvwDetails.ListItems.Clear
    
End Sub

Public Sub disabletxt()
Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
            i.Locked = True
        End If
    Next i

    
End Sub

Public Sub DisableCmd()
Dim i As Object
    For Each i In Me
        If TypeOf i Is CommandButton Then
            i.Enabled = False
        End If
    Next i
End Sub

Public Sub EnableCmd()
Dim i As Object
    For Each i In Me
        If TypeOf i Is CommandButton Then
            i.Enabled = True
        End If
    Next i
    
End Sub



Public Sub FirstLastDisb()
cmdLast.Enabled = True
cmdNext.Enabled = True
cmdPrevious.Enabled = True
cmdFirst.Enabled = True
cmdNext.SetFocus
            
With rsGlob
    If Not .BOF = True Then
        .MovePrevious
        If .BOF = True Then
            cmdLast.Enabled = True
            cmdNext.Enabled = True
            cmdPrevious.Enabled = False
            cmdFirst.Enabled = False
            cmdNext.SetFocus
        End If
        .MoveNext
    Else
        cmdLast.Enabled = True
        cmdNext.Enabled = True
        cmdPrevious.Enabled = False
        cmdFirst.Enabled = False
        cmdNext.SetFocus
    End If
    
    If Not .EOF = True Then
        .MoveNext
        If .EOF = True Then
            cmdLast.Enabled = False
            cmdNext.Enabled = False
            cmdPrevious.Enabled = True
            cmdFirst.Enabled = True
            cmdPrevious.SetFocus
        End If
        .MovePrevious
    Else
        cmdLast.Enabled = False
        cmdNext.Enabled = False
        cmdPrevious.Enabled = True
        cmdFirst.Enabled = True
        cmdPrevious.SetFocus
    End If
End With

End Sub




Private Sub txtCode_KeyPress(KeyAscii As Integer)
If Len(Trim(txtCode.Text)) > 19 Then
    Beep
    MsgBox "Can't enter more than 20 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("A") To Asc("Z")
  Case Asc("a") To Asc("z")
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

Private Sub txtComments_KeyPress(KeyAscii As Integer)
If Len(Trim(txtComments.Text)) > 198 Then
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

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
If Len(Trim(txtDesc.Text)) > 198 Then
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

Private Sub txtReff_KeyPress(KeyAscii As Integer)
If Len(Trim(txtReff.Text)) > 19 Then
    Beep
    MsgBox "Can't enter more than 20 characters", vbExclamation
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

Private Function loadACode() As String
    Set rs5 = CConnect.GetRecordSet("SELECT MAX(id) FROM Casuals")
    If rs5.EOF = False Then
        If rs5.RecordCount > 0 And Not IsNull(rs5.Fields(0)) Then
            loadACode = "CS" & CStr(rs5.Fields(0) + 1)
        Else
            loadACode = "CS1"
        End If
    Else
        loadACode = "CS1"
    End If
    Set rs5 = Nothing
End Function

