VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelContractTypes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search Engine-Contract Types"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3795
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   7
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   3360
         Width           =   1095
      End
      Begin VB.OptionButton OptName 
         Caption         =   "By Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton OptCode 
         Caption         =   "By Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtSearch 
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
         Left            =   4800
         TabIndex        =   1
         Top             =   240
         Width           =   1200
      End
      Begin MSComctlLib.ListView lvwDetails 
         Height          =   2730
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   4815
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgTree"
         ForeColor       =   16711680
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Search Field"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   3480
         TabIndex        =   3
         Top             =   240
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frmSelContractTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCont As Recordset
Dim X   '++This is the search field++

Private Sub InitGrid()
    With lvwDetails
        .ColumnHeaders.add , , "Code", 1500
        .ColumnHeaders.add , , "Description", 2000
        .ColumnHeaders.add , , "Duration", 1300
        .View = lvwReport
    End With
End Sub

Public Sub DisplayRecords()
Dim rsCheckMatch As New ADODB.Recordset
Set rsCheckMatch = CConnect.GetRecordSet("select * from EmpTerms where matchToContract=1")
IIf rsCheckMatch.RecordCount > 0, frmSelContractTypes.Tag = Trim(rsCheckMatch!Code & ""), frmSelContractTypes.Tag = ""

lvwDetails.ListItems.Clear
With rsCont
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
            Set li = lvwDetails.ListItems.add(, , !Code & "", , 5)
            li.ListSubItems.add , , !Description & ""
            li.ListSubItems.add , , IIf(IsNumeric(Trim(!correspondingvalue & "")) = True, Trim(!correspondingvalue & ""), 0) & " " & IIf(Trim(!InDays & "") = True, "Day(s)", IIf(Trim(!InWeeks & "") = True, "Week(s)", IIf(Trim(!InMonths & "") = True, "Month(s)", IIf(Trim(!InYears & "") = True, "Year(s)", "Unit(s)"))))
            li.ListSubItems(1).Tag = IIf(Trim(!InDays & "") = True, "d", IIf(Trim(!InWeeks & "") = True, "w", IIf(Trim(!InMonths & "") = True, "m", IIf(Trim(!InYears & "") = True, "y", "u"))))
            li.Tag = Trim(!correspondingvalue & "")
            li.ListSubItems(2).Tag = IIf(rsCheckMatch.RecordCount > 0, Trim(rsCheckMatch!Description & ""), "")
            .MoveNext
        Loop
    End If
End With
End Sub

Private Sub cmdCancel_Click()
    strName = ""
    Unload Me
End Sub

Private Sub cmdSelect_Click()
On Error GoTo 10
    strcode = lvwDetails.SelectedItem
    strName = lvwDetails.SelectedItem.ListSubItems(1).Text
    strValue = CLng(lvwDetails.SelectedItem.Tag)
    strDatePart = lvwDetails.SelectedItem.ListSubItems(1).Tag
    strNamePart = lvwDetails.SelectedItem.ListSubItems(2).Tag
    strID = Me.Tag
    Unload Me
Exit Sub
10: MsgBox err.Description
End Sub

Private Sub Form_Load()
On Error GoTo Hell
strNamePart = ""
CConnect.CColor Me, MyColor

Set rsCont = New Recordset
fraDetails.Visible = True
Set rsCont = CConnect.GetRecordSet("select * from pdContractTypes order by Code")
    Call InitGrid
    Call DisplayRecords
Set rsCont = Nothing

Exit Sub
Hell: MsgBox err.Description, vbCritical, "Search Records"
End Sub


Private Sub lvwDetails_Click()
    cmdSelect.Enabled = True
End Sub

Private Sub lvwDetails_DblClick()
    Call cmdSelect_Click
End Sub

Private Sub OptCode_Click()
    OptName.value = False
End Sub

Private Sub OptName_Click()
    OptCode.value = False
End Sub

Private Sub txtSearch_Change()
On Error GoTo Hell
Set rsCont = New Recordset

    X = txtSearch & "%"
    
    If OptCode.value = True Then
        Set rsCont = CConnect.GetRecordSet("select * from pdContractTypes where  Code Like '" & X & "'" & " order by Code")
    Else
        Set rsCont = CConnect.GetRecordSet("select * from pdContractTypes where  Description Like '" & X & "'" & " order by Description")
    End If
    
    Call InitGrid
    Call DisplayRecords
    
Exit Sub
Hell: MsgBox err.Description, vbCritical, "Search Records"
End Sub
