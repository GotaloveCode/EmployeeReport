VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmContract 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Contracts"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   Icon            =   "frmContracts.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7860
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
            Picture         =   "frmContracts.frx":0442
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContracts.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContracts.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContracts.frx":0778
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7900
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   7800
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
         TabIndex        =   9
         ToolTipText     =   "Move to the First employee"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
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
         TabIndex        =   16
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
         TabIndex        =   10
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
         TabIndex        =   11
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
         Picture         =   "frmContracts.frx":0CBA
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Picture         =   "frmContracts.frx":0DBC
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Picture         =   "frmContracts.frx":0EBE
         Style           =   1  'Graphical
         TabIndex        =   15
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
         TabIndex        =   12
         ToolTipText     =   "Move to the Last employee"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame fraDetails 
         Appearance      =   0  'Flat
         Caption         =   "Contract"
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
         Height          =   3315
         Left            =   720
         TabIndex        =   18
         Top             =   2040
         Visible         =   0   'False
         Width           =   6015
         Begin VB.CheckBox chkIsActive 
            Appearance      =   0  'Flat
            Caption         =   "&Make it the Active contract"
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
            Height          =   195
            Left            =   3450
            TabIndex        =   25
            Top             =   1560
            Width           =   2295
         End
         Begin VB.CommandButton cmdSearch 
            Height          =   315
            Left            =   1440
            Picture         =   "frmContracts.frx":13B0
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   480
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
            Picture         =   "frmContracts.frx":173A
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Cancel Process"
            Top             =   2670
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
            Left            =   4920
            Picture         =   "frmContracts.frx":183C
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Save Record"
            Top             =   2670
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
            Top             =   1815
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
            Left            =   3600
            TabIndex        =   4
            Top             =   1125
            Width           =   1650
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
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   480
            Width           =   1200
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
            Left            =   1860
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   480
            Width           =   3900
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
            Format          =   63111171
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
            Left            =   1920
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
            Format          =   63111171
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
            Left            =   1920
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
            Caption         =   "Contract Ref No."
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
            Left            =   3600
            TabIndex        =   22
            Top             =   915
            Width           =   1230
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
            TabIndex        =   21
            Top             =   240
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
            Left            =   1860
            TabIndex        =   20
            Top             =   240
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
            TabIndex        =   19
            Top             =   900
            Width           =   360
         End
      End
      Begin MSComctlLib.ListView lvwDetails 
         Height          =   7800
         Left            =   0
         TabIndex        =   26
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
      Begin MSComDlg.CommonDialog Cdl 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmContract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private ContractID As Long
Function Get_Contract_Type_Details(sField, sContCode) As Variant
    On Error GoTo Hell
    Dim rsCont As Recordset
    
    Set rsCont = CConnect.GetRecordSet("select * from pdContractTypes where Code='" & sContCode & "'")
    
        If rsCont.EOF = False Then Get_Contract_Type_Details = rsCont(sField)
    
    Set rsCont = Nothing
    
    Exit Function
Hell: MsgBox err.Description, vbCritical, "Search Records"
End Function

Public Sub cmdCancel_Click()

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
    If Not currUser Is Nothing Then
        If currUser.CheckRight("Contract") <> secModify Then
            MsgBox "You dont have right to delete the record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    If lvwDetails.ListItems.count > 0 Then
        resp = MsgBox("Are you sure you want to delete  " & lvwDetails.SelectedItem & " from the records?", vbQuestion + vbYesNo)
        If resp = vbNo Then
            Exit Sub
        End If
          
        Action = "DELETED EMPLOYEE CONTRACT; EMPLOYEE CODE: " & SelectedEmployee.EmpCode & "; CONTRACT NAME: " & lvwDetails.SelectedItem.ListSubItems(1).Text
        CConnect.ExecuteSql ("DELETE FROM Contracts WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND ID = " & lvwDetails.SelectedItem.Tag)
         
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
        If currUser.CheckRight("Contract") <> secModify Then
            MsgBox "You dont have right to edit the record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    If SelectedEmployee Is Nothing Then
        MsgBox "Please select an employee"
        Exit Sub
    End If
    
    If lvwDetails.ListItems.count < 1 Then
        MsgBox "You have to select the contract you would like to edit.", vbInformation
        PSave = True
        Call cmdCancel_Click
        PSave = False
        Exit Sub
    End If
    'rsGlob.Requery
    
    Set rs3 = CConnect.GetRecordSet("SELECT * FROM Contracts WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & lvwDetails.SelectedItem.Text & "'")
    With rs3
        If .RecordCount > 0 Then
            txtCode.Text = Trim(!Code & "")
            ' keep infor to tracking update or insert
            txtCode.Tag = Trim(!ID & "")
            Dim rsCont As New ADODB.Recordset
            Set rsCont = CConnect.GetRecordSet("select * from pdContractTypes where Code='" & txtCode.Text & "'")
            If rsCont.RecordCount > 0 Then
                strDatePart = IIf(Trim(rsCont!InDays & "") = True, "d", IIf(Trim(rsCont!InWeeks & "") = True, "w", IIf(Trim(rsCont!InMonths & "") = True, "m", IIf(Trim(rsCont!InYears & "") = True, "y", "u"))))
                strValue = IIf(IsNumeric(Trim(rsCont!correspondingvalue & "")) = True, Trim(rsCont!correspondingvalue & ""), 0)
            End If
            txtDesc.Text = !Description & ""
            If Not IsNull(!cFrom) Then dtpFrom.value = !cFrom & ""
            If Not IsNull(!cTo) Then dtpTo.value = !cTo & ""
            Me.Tag = SelectedEmployee.EmploymentTerm.EmpTermID
            strNamePart = SelectedEmployee.EmploymentTerm.EmpTermName
            txtComments.Text = !Comments & ""
            txtReff.Text = !Ref & ""
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
    cmdSearch.Enabled = False
    dtpFrom.Enabled = True
    SaveNew = False
    txtCode.Locked = True
    txtDesc.SetFocus
    Exit Sub
ErrHandler:
    MsgBox err.Description
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
    'Dim rsCheckEmpTerms As New ADODB.Recordset
    If Not currUser Is Nothing Then
        If currUser.CheckRight("Contract") <> secModify Then
            MsgBox "You dont have right to add new record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    If Not (SelectedEmployee Is Nothing) Then
        If Not (SelectedEmployee.EmploymentTerm.IsContract) Then
             MsgBox "This employee is not in contract terms bases." & vbNewLine & "Therefore you cannot assign a contract detail.", vbOKOnly + vbInformation, "Request denied": Exit Sub
        End If
        
        Call DisableCmd
    
        txtCode.Text = ""
        txtDesc.Text = ""
        txtReff.Text = ""
        txtComments.Text = ""
        dtpFrom.value = Date
        dtpTo.value = Date
        fraDetails.Visible = True
        cmdCancel.Enabled = True
        SaveNew = True
        cmdSave.Enabled = True
        txtCode.Locked = True
        txtDesc.SetFocus
        
        'fraDetails.Visible = True
        txtCode.Locked = True
        txtDesc.Locked = True
        cmdSearch.Enabled = True
        dtpFrom.Enabled = True
    End If
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

    If SelectedEmployee Is Nothing Then
        MsgBox "Please select employee To edit his\her contract details"
        Exit Sub
    End If
    
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

    If PromptSave = True Then
        If MsgBox("Are you sure you want to save the record.", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    If chkIsActive.value = 1 Then
        CConnect.ExecuteSql ("UPDATE contracts SET IsActive=0 WHERE employee_id='" & SelectedEmployee.EmployeeID & "'")
    End If
        
     'check if to update or insert
    If txtCode.Tag = "" Then
        
        Action = "ADDED EMPLOYEE CONTRACT; EMPLOYEE CODE: " & SelectedEmployee.EmpCode & "; CONTRACT CODE: " & txtCode.Text
              
        mySQL = "INSERT INTO Contracts (employee_id, Code, Description, CFrom, CTo, Ref, Comments,IsActive)" & _
                            " VALUES('" & SelectedEmployee.EmployeeID & "','" & txtCode.Text & "','" & txtDesc.Text & "'," & _
                            "'" & SQLDate(dtpFrom.value) & "','" & SQLDate(dtpTo.value) & "','" & txtReff.Text & "','" & txtComments.Text & "'," & chkIsActive.value & ")"
        
        CConnect.ExecuteSql (mySQL)
        
    Else
         mySQL = "UPDATE  Contracts SET CFrom='" & SQLDate(dtpFrom.value) & "', CTo='" & SQLDate(dtpTo.value) & "',Ref='" & txtReff.Text & "', Comments='" & txtComments.Text & "',IsActive= " & chkIsActive.value & " where ID=" & CInt(txtCode.Tag)
         Action = "ADDED EMPLOYEE CONTRACT; EMPLOYEE CODE: " & SelectedEmployee.EmpCode & "; CONTRACT CODE: " & txtCode.Text
          CConnect.ExecuteSql (mySQL)
    End If
    
    If strNamePart <> "" Then
        If chkIsActive.value = 1 Then CConnect.ExecuteSql "UPDATE employees SET EmploymentValidThro='" & SQLDate(dtpTo.value) & "' WHERE EmployeeID='" & SelectedEmployee.EmployeeID & "'"
    End If
    
    rs2.Requery
    Call DisplayRecords
    
    'refresh the employee list
    Call RefreshEmployeesCol
    'refresh the detauls of the employee
    Set SelectedEmployee = AllEmployees.FindEmployee(SelectedEmployee.EmployeeID)
    If SaveNew = False Then
        PSave = True
        Call cmdCancel_Click
        PSave = False
    Else
        txtDesc.SetFocus
        txtCode.Text = loadACode
    End If
    
'++Unload the frame & do Maujanja
    fraDetails.Visible = False
    
    Call EnableCmd
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    SaveNew = False
    
    With frmMain2
        .cmdNew.Enabled = True
        .cmdEdit.Enabled = True
        .cmdDelete.Enabled = True
        .cmdCancel.Enabled = False
        .cmdSave.Enabled = False
    End With
'++Unload the frame & do Maujanja
    
End Sub


Private Sub cmdSearch_Click()
    If frmSelContractTypes.lvwDetails.ListItems.count = 0 Then MsgBox "There are no registered contract types." & vbNewLine & "Please do so to proceed.", vbOKOnly + vbInformation, "Contract types": Exit Sub
    frmSelContractTypes.Show vbModal
    txtCode.Text = strcode
    txtCode.Tag = strID
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
    
    Set rs2 = CConnect.GetRecordSet("SELECT * FROM Contracts ORDER BY Code")

     With rsGlob
        If .RecordCount < 1 Then
            Call DisableCmd
            Exit Sub
        End If
    End With
    
    Call DisplayRecords
    
    cmdFirst.Enabled = False
    cmdPrevious.Enabled = False
    Me.txtCode.Locked = True
    Me.txtDesc.Locked = True
    Me.dtpFrom.Enabled = False
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
        .ColumnHeaders.add , , "Contract Code", 0
        .ColumnHeaders.add , , "Description", 294 * .Width / 1080
        .ColumnHeaders.add , , "Date From", .Width / 8
        .ColumnHeaders.add , , "Date To", .Width / 8
        .ColumnHeaders.add , , "Active Contract", .Width / 9
        .ColumnHeaders.add , , "Contract Ref No.", .Width / 6
        .ColumnHeaders.add , , "Comments", .Width / 5
        .View = lvwReport
    End With
End Sub

Public Sub DisplayRecords()
    On Error GoTo ErrHandler
   
    Call Cleartxt
    
    If Not SelectedEmployee Is Nothing Then
        
        With rs2
            If .RecordCount > 0 Then
                .Filter = "employee_id like '" & SelectedEmployee.EmployeeID & "'"
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                        Set li = lvwDetails.ListItems.add(, , !Code & "", , 5)
                        li.ListSubItems.add , , !Description & ""
                        li.ListSubItems.add , , Format(Trim(!cFrom & ""), "dd, MMM, yyyy")
                        li.ListSubItems.add , , Format(Trim(!cTo & ""), "dd, MMM, yyyy")
                        li.ListSubItems.add , , IIf(Trim(!isactive & "") = True, "Yes", "")
                        li.ListSubItems.add , , !Ref & ""
                        li.ListSubItems.add , , !Comments & ""
                        li.Tag = !ID
                        .MoveNext
                    Loop
                End If
                .Filter = adFilterNone
            End If
        End With
    End If
 
    Exit Sub
ErrHandler:
    MsgBox "please ensure that you have selected employee"
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set rs2 = Nothing
    Set rs5 = Nothing
    frmMain2.Caption = "Infiniti HRMIS - PDR [Current User:\" & currUser.FullNames & "]"
    
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
    Me.cmdDelete.Enabled = True
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
    Set rs5 = CConnect.GetRecordSet("SELECT MAX(id) FROM Contracts")
    If rs5.EOF = False Then
        If rs5.RecordCount > 0 And Not IsNull(rs5.Fields(0)) Then
            loadACode = "CONT" & CStr(rs5.Fields(0) + 1)
        Else
            loadACode = "CONT1"
        End If
    Else
        loadACode = "CONT1"
    End If
    Set rs5 = Nothing
End Function

