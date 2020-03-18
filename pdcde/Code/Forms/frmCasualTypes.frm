VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCasualTypes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Casual Types"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7860
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11130
      Begin VB.Frame fraDetails 
         Appearance      =   0  'Flat
         Caption         =   "Casual Types"
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
         Height          =   2775
         Left            =   1800
         TabIndex        =   9
         Top             =   1680
         Visible         =   0   'False
         Width           =   5895
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
            TabIndex        =   20
            Top             =   480
            Width           =   4620
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
            TabIndex        =   19
            Top             =   480
            Width           =   1065
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
            Height          =   1815
            Left            =   60
            TabIndex        =   10
            Top             =   840
            Width           =   5775
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
               Picture         =   "frmCasualTypes.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   17
               ToolTipText     =   "Cancel Process"
               Top             =   1230
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
               Picture         =   "frmCasualTypes.frx":0102
               Style           =   1  'Graphical
               TabIndex        =   16
               ToolTipText     =   "Save Record"
               Top             =   1230
               Width           =   495
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
               TabIndex        =   15
               Top             =   240
               Value           =   -1  'True
               Width           =   855
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
               Left            =   120
               TabIndex        =   14
               Top             =   480
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
               Left            =   120
               TabIndex        =   13
               Top             =   720
               Width           =   1095
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
               Left            =   120
               TabIndex        =   12
               Top             =   960
               Width           =   1215
            End
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
               TabIndex        =   11
               Top             =   1440
               Width           =   1575
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
               TabIndex        =   18
               Top             =   1200
               Width           =   1530
            End
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
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
            TabIndex        =   22
            Top             =   255
            Width           =   795
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
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
            TabIndex        =   21
            Top             =   255
            Width           =   375
         End
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
         TabIndex        =   8
         ToolTipText     =   "Move to the Last employee"
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
         Picture         =   "frmCasualTypes.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Delete Record"
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
         Picture         =   "frmCasualTypes.frx":06F6
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Edit Record"
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
         Picture         =   "frmCasualTypes.frx":07F8
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Add New record"
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
         TabIndex        =   4
         ToolTipText     =   "Move to the Next employee"
         Top             =   6120
         Visible         =   0   'False
         Width           =   495
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
         TabIndex        =   3
         ToolTipText     =   "Move to the Previous employee"
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
         TabIndex        =   2
         Top             =   6120
         Visible         =   0   'False
         Width           =   1050
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
         TabIndex        =   1
         ToolTipText     =   "Move to the First employee"
         Top             =   6120
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSComDlg.CommonDialog Cdl 
         Left            =   5610
         Top             =   4560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ListView lvwContractTypes 
         Height          =   7800
         Left            =   0
         TabIndex        =   23
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
End
Attribute VB_Name = "frmCasualTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim ContractCount As Integer

Private Sub cboRel_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Public Sub cmdCancel_Click()
'    Call DisplayRecords
'    fraDetails.Visible = False
'
'    Call EnableCmd
'    cmdCancel.Enabled = False
'    cmdSave.Enabled = False
'    SaveNew = False
'
'    With frmMain2
'        .cmdNew.Enabled = True
'        .cmdEdit.Enabled = True
'        .cmdDelete.Enabled = True
'        .cmdCancel.Enabled = False
'        .cmdSave.Enabled = False
'    End With
'Refresh the listview
Call DisplayRecords
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
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Public Sub cmdDelete_Click()
Dim resp As String


If lvwContractTypes.ListItems.count > 0 Then
    resp = MsgBox("Are you sure you want to delete  " & lvwContractTypes.SelectedItem & " from the records?", vbQuestion + vbYesNo)
    If resp = vbNo Then
        Exit Sub
    End If
    
    CConnect.ExecuteSql ("DELETE FROM pdCasualTypes WHERE Code = '" & lvwContractTypes.SelectedItem & "'")
       
    rs2.Requery
    
    Call DisplayRecords
Else
    MsgBox "You have to select the record you would like to delete.", vbInformation
            
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

Set rs3 = CConnect.GetRecordSet("SELECT * FROM pdCasualTypes WHERE Code = '" & lvwContractTypes.SelectedItem & "'")

With rs3
    If .RecordCount > 0 Then
        txtCode.Text = !Code & ""
        txtDescription.Text = !Description & ""
        optInDays.value = IIf(Trim(!InDays & "") = True, True, False)
        optInWeeks.value = IIf(Trim(!InWeeks & "") = True, True, False)
        optInMonths.value = IIf(Trim(!InMonths & "") = True, True, False)
        optInYears.value = IIf(Trim(!InYears & "") = True, True, False)
        txtValue.Text = IIf(IsNumeric(Trim(!correspondingvalue & "")) = True, Trim(!correspondingvalue & ""), 0)
        
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
Dim rsBenefit As Recordset, TotalBenefit As Currency

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

If SaveNew = True Then
    Set rs4 = CConnect.GetRecordSet("SELECT * FROM pdCasualTypes WHERE Code = '" & txtCode.Text & "'")
    With rs4
        If .RecordCount > 0 Then
            MsgBox "The code already exists. Enter another one.", vbInformation
            txtCode.Text = ""
            txtCode.SetFocus
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

Action = "REGISTERED CASUAL TYPES; CASUAL CODE: " & txtCode.Text & "; CONTRACT DESCRIPTION: " & txtDescription.Text

CConnect.ExecuteSql ("DELETE FROM pdCasualTypes WHERE Code = '" & txtCode.Text & "'")

    mySQL = "INSERT INTO pdCasualTypes (Code,Description,InDays,InWeeks, InMonths, InYears,CorrespondingValue)" & _
                    " VALUES('" & txtCode.Text & "','" & txtDescription.Text & "'," & IIf(optInDays.value = True, 1, 0) & "," & IIf(optInWeeks.value = True, 1, 0) & "," & IIf(optInMonths.value = True, 1, 0) & "," & IIf(optInYears.value = True, 1, 0) & "," & txtValue.Text & ")"

CConnect.ExecuteSql (mySQL)

rs2.Requery

If SaveNew = False Then
    Call cmdCancel_Click
Else
    rs2.Requery
    Call DisplayRecords
    txtCode.Text = loadCTCode
    txtDescription.SetFocus
    SaveNew = True
End If

MsgBox "Records updated Successfully", vbInformation, "Casual Set up"

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
    
    Set rs2 = CConnect.GetRecordSet("SELECT * FROM pdCasualTypes ORDER BY Code")
    
    With rsGlob
        If .RecordCount < 1 Then
            Call DisableCmd
            Exit Sub
        End If
    End With
    
    Call DisplayRecords
    
    cmdFirst.Enabled = False
    cmdPrevious.Enabled = False
    
    txtCode.Locked = True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "A slight error has occurred" & vbNewLine & err.Description, vbInformation, TITLES
End Sub
Private Sub Form_Resize()
    oSmart.FResize Me
    
    Me.Height = tvwMainheight - 130
    Frame1.Move Frame1.Left, 0, Frame1.Width, tvwMainheight - 140
    lvwContractTypes.Move lvwContractTypes.Left, 0, lvwContractTypes.Width, tvwMainheight - 140
End Sub

Private Sub InitGrid()
    With lvwContractTypes
        .ColumnHeaders.add , , "Code", 0
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
Dim i As Object
For Each i In Me
    If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
        i.Text = ""
    End If
Next i
lvwContractTypes.ListItems.Clear
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
    Set rs5 = CConnect.GetRecordSet("SELECT MAX(id) FROM pdCasualTypes")
    If rs5.EOF = False Then
        If rs5.RecordCount > 0 And Not IsNull(rs5.Fields(0)) Then
            loadCTCode = "CS0" & CStr(rs5.Fields(0) + 1)
        Else
            loadCTCode = "CS01"
        End If
    Else
        loadCTCode = "CS01"
    End If
    Set rs5 = Nothing
End Function



