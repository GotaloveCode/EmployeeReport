VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmContractTypes 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contract Types"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7900
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11130
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
         TabIndex        =   14
         ToolTipText     =   "Move to the First employee"
         Top             =   6120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "CLOSE"
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
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   6120
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
         TabIndex        =   12
         ToolTipText     =   "Move to the Previous employee"
         Top             =   6120
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
         Top             =   6120
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
         Picture         =   "frmContractTypes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Add New record"
         Top             =   6120
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
         Picture         =   "frmContractTypes.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Edit Record"
         Top             =   6120
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
         Picture         =   "frmContractTypes.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Delete Record"
         Top             =   6120
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
         TabIndex        =   7
         ToolTipText     =   "Move to the Last employee"
         Top             =   6120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame fraDetails 
         Appearance      =   0  'Flat
         Caption         =   "Contract Types"
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   1800
         TabIndex        =   4
         Top             =   1680
         Visible         =   0   'False
         Width           =   5895
         Begin VB.Frame Frame2 
            Caption         =   "Define working days per week"
            Height          =   1935
            Left            =   60
            TabIndex        =   22
            Top             =   2200
            Width           =   5775
            Begin VB.TextBox TxtOverTimeRate 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   2400
               TabIndex        =   27
               Text            =   "0"
               Top             =   840
               Width           =   855
            End
            Begin VB.CheckBox ChkOverTime 
               Appearance      =   0  'Flat
               Caption         =   "Charge extra days worked to overtime"
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   2400
               TabIndex        =   26
               Top             =   240
               Width           =   3255
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
               Left            =   5160
               Picture         =   "frmContractTypes.frx":06F6
               Style           =   1  'Graphical
               TabIndex        =   25
               ToolTipText     =   "Cancel Process"
               Top             =   1320
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
               Left            =   4680
               Picture         =   "frmContractTypes.frx":07F8
               Style           =   1  'Graphical
               TabIndex        =   24
               ToolTipText     =   "Save Record"
               Top             =   1320
               Width           =   495
            End
            Begin MSComctlLib.ListView LvwWeekDays 
               Height          =   1575
               Left            =   120
               TabIndex        =   23
               Top             =   240
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   2778
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               Checkboxes      =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Week Day"
                  Object.Width           =   3836
               EndProperty
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Overtime Rate (Per day)"
               Height          =   195
               Left            =   2400
               TabIndex        =   28
               Top             =   600
               Width           =   1695
            End
         End
         Begin VB.Frame fraContractDetails 
            Caption         =   "Contract details declaration:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   60
            TabIndex        =   15
            Top             =   840
            Width           =   5775
            Begin VB.TextBox txtValue 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Left            =   120
               TabIndex        =   20
               Top             =   960
               Width           =   1575
            End
            Begin VB.OptionButton optInYears 
               Appearance      =   0  'Flat
               Caption         =   "In years"
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
               Height          =   255
               Left            =   3240
               TabIndex        =   19
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton optInMonths 
               Appearance      =   0  'Flat
               Caption         =   "In Months"
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
               Height          =   255
               Left            =   2160
               TabIndex        =   18
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton optInWeeks 
               Appearance      =   0  'Flat
               Caption         =   "In weeks"
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
               Height          =   255
               Left            =   1080
               TabIndex        =   17
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton optInDays 
               Appearance      =   0  'Flat
               Caption         =   "In days"
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
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Top             =   240
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Please enter a value:"
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
               Top             =   720
               Width           =   1530
            End
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
            Height          =   285
            Left            =   60
            TabIndex        =   1
            Top             =   480
            Width           =   1065
         End
         Begin VB.TextBox txtDescription 
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
            Height          =   285
            Left            =   1215
            TabIndex        =   2
            Top             =   480
            Width           =   4620
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
            Left            =   60
            TabIndex        =   6
            Top             =   255
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
            Left            =   1215
            TabIndex        =   5
            Top             =   255
            Width           =   795
         End
      End
      Begin MSComDlg.CommonDialog Cdl 
         Left            =   5520
         Top             =   6120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ListView lvwContractTypes 
         Height          =   7800
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   11130
         _ExtentX        =   19632
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
   Begin MSComctlLib.ImageList imgEmpTool 
      Left            =   5310
      Top             =   165
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
            Picture         =   "frmContractTypes.frx":08FA
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContractTypes.frx":0A0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContractTypes.frx":0B1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContractTypes.frx":0C30
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmContractTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ContractCount As Integer

Private Sub cboRel_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub ChkOverTime_Click()
    On Error GoTo ErrorHandler
    
    If (ChkOverTime.value = 1) Then
        If (NumberOfDaysChecked = 7) Then 'IF all days are selected, then there are no overtime days
            MsgBox "Not applicable because all week days as selected as working days", vbExclamation, "Error"
            ChkOverTime.value = vbUnchecked
            Exit Sub
        Else
            TxtOverTimeRate.Enabled = True
        End If
    Else
        TxtOverTimeRate.Enabled = False
    End If
    TxtOverTimeRate.Text = ""
    
    Exit Sub
ErrorHandler:
    MsgBox "An error has occured:" & vbNewLine & err.Description, vbExclamation, "PDR Error"
End Sub

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
    Call EnableCmd
    txtCode.Tag = ""
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Public Sub cmdDelete_Click()
    Dim resp As String
    Dim rs As ADODB.Recordset
    Set rs = CConnect.GetRecordSet("SELECT * FROM Contracts WHERE code = '" & lvwContractTypes.SelectedItem.Text & "'")
  
    If Not (rs Is Nothing) And rs.RecordCount = 0 Then
        If lvwContractTypes.ListItems.count > 0 Then
            resp = MsgBox("Are you sure you want to delete  " & lvwContractTypes.SelectedItem & " from the records?", vbQuestion + vbYesNo)
            If resp = vbNo Then
                Exit Sub
            End If
            
            CConnect.ExecuteSql ("DELETE FROM pdContractTypes WHERE Code = '" & lvwContractTypes.SelectedItem & "'")
               
            rs2.Requery
            
            Call DisplayRecords
        Else
            MsgBox "You have to select the record you would like to delete.", vbInformation
                    
        End If
    Else
        MsgBox " That Contract can not be deleted now " & vbNewLine & "because it is assigned to employees."
    End If
        
End Sub

Private Sub cmdDone_Click()
    Call cmdCancel_Click
End Sub

Public Sub cmdEdit_Click()
On Error GoTo 10

If lvwContractTypes.ListItems.count < 1 Then
    MsgBox "You have to select the code you would like to edit.", vbInformation
    Call cmdCancel_Click
    Exit Sub
End If

Set rs3 = CConnect.GetRecordSet("SELECT * FROM pdContractTypes WHERE ID = '" & lvwContractTypes.SelectedItem.Tag & "'")

With rs3
    If .RecordCount > 0 Then
        txtCode.Text = !Code & ""
        txtDescription.Text = !Description & ""
        optInDays.value = IIf(Trim(!InDays & "") = True, True, False)
        optInWeeks.value = IIf(Trim(!InWeeks & "") = True, True, False)
        optInMonths.value = IIf(Trim(!InMonths & "") = True, True, False)
        optInYears.value = IIf(Trim(!InYears & "") = True, True, False)
        txtValue.Text = IIf(IsNumeric(Trim(!correspondingvalue & "")) = True, Trim(!correspondingvalue & ""), 0)
        txtCode.Tag = !Code
        If (CBool(!ChargeToOverTime)) Then
            ChkOverTime.value = vbChecked
        Else
            ChkOverTime.value = vbUnchecked
        End If
        
        If Not IsNull(!OverTimeRate) Then TxtOverTimeRate.Text = CDbl(Trim(!OverTimeRate))
        
        '==SHOW WORKING DAYS==
        Dim wDays, wK As Integer, n As Integer
        wDays = Split(!ContractWorkingDays, ",")
        For wK = LBound(wDays) To UBound(wDays)
            For n = 1 To 7
                If (wDays(wK) = LvwWeekDays.ListItems.Item(n).Tag) Then
                    LvwWeekDays.ListItems.Item(n).Checked = True
                    Exit For
                End If
            Next n
        Next wK
        '==END SHOW WORKING DAYS==
        
        SaveNew = False
    Else
        MsgBox "Record not found.", vbInformation
        Set rs3 = Nothing
        Call cmdCancel_Click
        Exit Sub
    End If
End With

Set rs3 = Nothing

Call DisableCmd

fraDetails.Visible = True

cmdSave.Enabled = True
cmdCancel.Enabled = True
SaveNew = False

txtCode.Locked = False
txtDescription.SetFocus

Exit Sub
10:
MsgBox err.Description, vbCritical, "Contract Set Up"
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
On Error Resume Next

Call DisableCmd
Call Cleartxt

fraDetails.Visible = True
cmdCancel.Enabled = True
SaveNew = True
cmdSave.Enabled = True
'dtpFrom.Value = Date
'dtpTo.Value = Date
'txtDOB.Text = "__/__/____"
'txtDOB1.Text = "__/__/____"
txtCode = loadCTCode
txtCode.Locked = False

'==BY DEFAULT, IT ASSUMED THAT WORKING DAYS ARE MON-SAT
    Dim wK As Integer
    If (LvwWeekDays.ListItems.count > 0) Then
        For wK = 1 To 5
            LvwWeekDays.ListItems.Item(wK).Checked = True
        Next wK
    End If
'==END DEFAULT WEEKDAYS ALLOCATION

End Sub

Private Sub cmdNext_Click()
    
With rsGlob
    ''If .RecordCount > 0 Then
    
    If Not .EOF Then
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
    Dim rsBenefit As Recordset, TotalBenefit As Currency
    On Error GoTo ErrHandler
    
    If txtCode.Text = "" Then
        MsgBox "Enter the code.", vbInformation
        txtCode.SetFocus
        Exit Sub
    End If
    
    If txtDescription.Text = "" Then
        MsgBox "Enter the description.", vbInformation
        txtDescription.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtValue.Text) Then
        MsgBox "Please enter a corresponding value.", vbOKOnly + vbInformation, "Missing value"
        txtValue.SetFocus
        Exit Sub
    End If
     
    ' test for duplicate
    Dim ItemX As ListItem
   
    Set ItemX = Me.lvwContractTypes.FindItem(txtCode.Text, lvwText)
    If Not ItemX Is Nothing Then
        If txtCode.Tag <> txtCode.Text Then
            MsgBox "Contract type with that code already exists. Enter another one.", vbInformation
            txtCode.Text = ""
            txtCode.SetFocus
            
            Exit Sub
        End If
    End If
      
    If PromptSave = True Then
        If MsgBox("Are you sure you want to save the record.", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    'vERIFY wORKING DAYS
    If (LvwWeekDays.ListItems.count = 0) Then
        MsgBox "At least ONE working day must be selected", vbExclamation, "Error"
        Exit Sub
    End If
    If (NumberOfDaysChecked = 7) Then 'IF all days are selected, then there are no overtime days
        MsgBox "Not applicable because all week days as selected as working days", vbExclamation, "Error"
        ChkOverTime.value = vbUnchecked
        Exit Sub
    End If
        
    Dim ContractWorkingDays As String, wK As Integer
    ContractWorkingDays = ""
    For wK = 1 To 7
        With LvwWeekDays
            If .ListItems.Item(wK).Checked = True Then
                ContractWorkingDays = ContractWorkingDays & WeekdayName(wK, True, vbMonday) & ","
            End If
        End With
    Next wK
    ContractWorkingDays = Mid(ContractWorkingDays, 1, Len(ContractWorkingDays) - 1)
    
 'check if to update or insert
    If txtCode.Tag = "" Then
        Action = "REGISTERED CONTRACT TYPES; CONTRACT CODE: " & txtCode.Text & "; CONTRACT DESCRIPTION: " & txtDescription.Text
        '-CConnect.ExecuteSql ("DELETE FROM pdContractTypes WHERE Code = '" & txtCode.Text & "'")
        If (Trim(TxtOverTimeRate.Text) = "") Then TxtOverTimeRate.Text = 0
        mySQL = "INSERT INTO pdContractTypes (Code,Description,InDays,InWeeks, InMonths, InYears,CorrespondingValue,ContractWorkingDays,ChargeToOverTime,OverTimeRate)" & _
                            " VALUES('" & txtCode.Text & "','" & txtDescription.Text & "'," & IIf(optInDays.value = True, 1, 0) & "," & IIf(optInWeeks.value = True, 1, 0) & "," & IIf(optInMonths.value = True, 1, 0) & "," & IIf(optInYears.value = True, 1, 0) & _
                            "," & txtValue.Text & ",'" & ContractWorkingDays & "'," & CInt(ChkOverTime.value) & "," & CDbl(Trim(TxtOverTimeRate)) & ")"
        CConnect.ExecuteSql (mySQL)
    Else
        Action = "UPDATED CONTRACT TYPES; CONTRACT CODE: " & txtCode.Text & "; CONTRACT DESCRIPTION: " & txtDescription.Text
        TxtOverTimeRate.Text = CDbl(IIf(TxtOverTimeRate = "", 0, TxtOverTimeRate))
        mySQL = "UPDATE pdContractTypes SET Code='" & txtCode.Text & "',Description='" & txtDescription.Text & _
        "',InDays=" & IIf(optInDays.value = True, 1, 0) & ",InWeeks=" & IIf(optInWeeks.value = True, 1, 0) & ",InMonths=" & IIf(optInMonths.value = True, 1, 0) & _
        ", InYears=" & IIf(optInYears.value = True, 1, 0) & ",CorrespondingValue= " & txtValue.Text & _
        ", ContractWorkingDays = '" & ContractWorkingDays & "', ChargeToOvertime = " & CInt(ChkOverTime.value) & _
        ", OverTimeRate = " & CDbl(Trim(TxtOverTimeRate)) & " WHERE code='" & txtCode.Tag & "'" ',Cto='',CFrom=''"
        CConnect.ExecuteSql (mySQL)
    End If
    rs2.Requery
    Call DisplayRecords
    
    If SaveNew = False Then
        Call cmdCancel_Click
    Else
        'rs2.Requery
        Call DisplayRecords
        txtCode.Text = loadCTCode
        txtDescription.SetFocus
        SaveNew = True
    End If
    txtCode.Tag = ""
    MsgBox "Records updated Successfully", vbInformation, "Contract Set up"
    
    Exit Sub
ErrHandler:
    MsgBox "An error has occured when saving the contract details"
End Sub

Private Sub dtpFrom_CloseUp()
'txtDOB.Text = Format(dtpFrom.Value, "dd/mm/yyyy")
End Sub

Private Sub dtpTo_CloseUp()
'txtDOB1.Text = Format(dtpTo.Value, "dd/mm/yyyy")
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
        
    oSmart.FReset Me
    
    If oSmart.hRatio > 1.1 Then
        With frmMain2
            Me.Move .tvwMain.Width + .tvwMain.Width / 32.75, (.Height / 5.52) '- 155
        End With
    Else
         With frmMain2
            Me.Move .tvwMain.Width + .tvwMain.Width / 32.75, .Height / 5.55
        End With
        
    End If
    
    CConnect.CColor Me, MyColor
    
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    
    Call InitGrid
    
    Set rs2 = CConnect.GetRecordSet("SELECT * FROM pdContractTypes ORDER BY Code")
    
    With rsGlob
       '' If .RecordCount < 1 Then
        If .EOF Then
            Call DisableCmd
            Exit Sub
        End If
    End With
    
    Call DisplayRecords
    Call LoadWeekDaysToList
    
    cmdFirst.Enabled = False
    cmdPrevious.Enabled = False
    
    txtCode.Locked = True
        
    Exit Sub
    
ErrorHandler:
    MsgBox "A slight error has occurred" & vbNewLine & err.Description, vbInformation, TITLES
End Sub


Private Sub Form_Resize()
    oSmart.FResize Me
    
    Me.Height = tvwMainheight - 120
    Frame1.Move Frame1.Left, 0, Frame1.Width, tvwMainheight - 140
    lvwContractTypes.Move lvwContractTypes.Left, 0, lvwContractTypes.Width, tvwMainheight - 140

End Sub

Private Sub InitGrid()
    With lvwContractTypes
        .ColumnHeaders.add , , "Code", 0 '.Width / 7
        .ColumnHeaders.add , , "Description", 4 * .Width / 5
        .ColumnHeaders.add , , "Duration", .Width / 5
        .View = lvwReport
    End With
End Sub

Public Sub DisplayRecords()
    lvwContractTypes.ListItems.Clear
    Call Cleartxt
    With rs2
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Set li = lvwContractTypes.ListItems.add(, , !Code & "", , 5)
                li.ListSubItems.add , , !Description & ""
                li.ListSubItems.add , , IIf(Trim(!InDays & "") = True, Trim(!correspondingvalue & "") & " Day(s)", IIf(Trim(!InWeeks & "") = True, Trim(!correspondingvalue & "") & " Week(s)", IIf(Trim(!InMonths & "") = True, Trim(!correspondingvalue & "") & " Month(s)", IIf(Trim(!InYears & "") = True, Trim(!correspondingvalue & "") & " Year(s)", Trim(!correspondingvalue & "") & " Unit(s)"))))
                li.Tag = !ID
                .MoveNext
            Loop
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs2 = Nothing
    Set rs5 = Nothing
'    Set Cnn = Nothing
    frmMain2.Caption = "Infiniti HRMIS - PDR [Current User:\" & currUser.FullNames & "]"
End Sub

Private Sub lvwContractTypes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwContractTypes
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvwContractTypes_DblClick()
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
    Dim i As Object, wK As Integer
    
    For Each i In Me
        If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
            i.Text = ""
        End If
    Next i
    
    'lvwContractTypes.ListItems.clear
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
Case Is = 24
Case Is = 3
        Case Is = 22
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


Private Sub txtDescription_KeyPress(KeyAscii As Integer)
If Len(Trim(txtDescription.Text)) > 198 Then
    Beep
    MsgBox "Can't enter more than 200 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
Case Is = 24
Case Is = 3
        Case Is = 22
  'Case Asc("0") To Asc("9")
  Case Asc("A") To Asc("Z")
  Case Asc("a") To Asc("z")
  Case Asc("0") To Asc("9")
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

Private Sub txtDOB_Change()
On Error Resume Next
'dtpFrom.Value = CDate(txtDOB.Text)
End Sub

Private Sub txtDOB_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case Is = 24
Case Is = 3
        Case Is = 22
  Case Asc("0") To Asc("9")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub

Private Sub txtDOB1_Change()
On Error Resume Next
'dtpTo.Value = CDate(txtDOB1.Text)
End Sub

Private Sub txtDOB1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case Is = 24
Case Is = 3
        Case Is = 22
  Case Asc("0") To Asc("9")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub

Private Function loadCTCode() As String
    Set rs5 = CConnect.GetRecordSet("SELECT MAX(id) FROM pdContractTypes")
    If rs5.EOF = False Then
        If rs5.RecordCount > 0 And Not IsNull(rs5.Fields(0)) Then
            loadCTCode = "CONT0" & CStr(rs5.Fields(0) + 1)
        Else
            loadCTCode = "CONT01"
        End If
    Else
        loadCTCode = "CONT01"
    End If
    Set rs5 = Nothing
End Function

Private Sub LoadWeekDaysToList()
    On Error GoTo ErrorHandler
    Dim wK As Integer, ItemX As ListItem
    
    With LvwWeekDays
        For wK = 1 To 7
            Set ItemX = .ListItems.add(, , WeekdayName(wK, False, vbMonday))
            ItemX.Tag = WeekdayName(wK, True, vbMonday)
        Next wK
    End With
    
    Exit Sub
ErrorHandler:
    MsgBox "Error Loading Week Days:" & vbCrLf & err.Description, vbExclamation, "Error"
End Sub

Private Function NumberOfDaysChecked() As Integer
    On Error GoTo ErrorHandler
    
    Dim n As Integer
    For n = 1 To 7 Step 1
        If (LvwWeekDays.ListItems.Item(n).Checked = True) Then NumberOfDaysChecked = NumberOfDaysChecked + 1
    Next n
    
    Exit Function
ErrorHandler:
    MsgBox "Error:" & vbNewLine & err.Description, vbExclamation, "Error"
End Function
