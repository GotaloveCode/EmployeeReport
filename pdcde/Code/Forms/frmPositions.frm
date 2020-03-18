VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPositions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Positions"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11160
   Icon            =   "frmPositions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imgEmpTool 
      Left            =   7920
      Top             =   1440
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
            Picture         =   "frmPositions.frx":0442
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPositions.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPositions.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPositions.frx":0778
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8100
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11130
      Begin VB.Frame fraDetails 
         Appearance      =   0  'Flat
         Caption         =   "Positions"
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
         Height          =   4170
         Left            =   1680
         TabIndex        =   17
         Top             =   1695
         Visible         =   0   'False
         Width           =   6855
         Begin VB.Frame Frame2 
            Caption         =   "Set parameters for workforce module"
            Height          =   855
            Left            =   120
            TabIndex        =   24
            Top             =   2640
            Width           =   6615
            Begin VB.CheckBox chkVisibleInOGram 
               Caption         =   "Visible In Organogram"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   26
               Top             =   360
               Width           =   2055
            End
            Begin VB.TextBox txtMaxEmpCount 
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
               Left            =   5760
               TabIndex        =   25
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Max. Employees Count in Organogram"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2760
               TabIndex        =   27
               Top             =   360
               Width           =   3015
            End
         End
         Begin VB.TextBox txtDefaultRemun 
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
            Left            =   2400
            TabIndex        =   4
            Top             =   2160
            Width           =   1215
         End
         Begin VB.TextBox txtDailyCasualRate 
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5520
            TabIndex        =   5
            Top             =   2160
            Width           =   1050
         End
         Begin VB.TextBox txtCode 
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   480
            Width           =   930
         End
         Begin VB.TextBox txtPositionName 
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
            Left            =   1320
            TabIndex        =   2
            Top             =   480
            Width           =   5295
         End
         Begin VB.TextBox txtDetails 
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
            Height          =   825
            Left            =   135
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   1200
            Width           =   6435
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
            Left            =   5595
            Picture         =   "frmPositions.frx":0CBA
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Save Record"
            Top             =   3540
            Width           =   510
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
            Left            =   6090
            Picture         =   "frmPositions.frx":0DBC
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Cancel Process"
            Top             =   3540
            Width           =   495
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Default Monthly Remuneration"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   2160
            Width           =   2295
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Daily Casual Rate"
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
            Index           =   2
            Left            =   4080
            TabIndex        =   22
            Top             =   2160
            Width           =   1260
         End
         Begin VB.Label Label1 
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
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Position Name"
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
            Left            =   1320
            TabIndex        =   19
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Position Details"
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
            Top             =   960
            Width           =   1080
         End
      End
      Begin MSComctlLib.ListView lvwDetails 
         Height          =   7800
         Left            =   0
         TabIndex        =   0
         Top             =   120
         Width           =   11130
         _ExtentX        =   19632
         _ExtentY        =   13758
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Name            =   "Verdana"
            Size            =   9.75
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
         Picture         =   "frmPositions.frx":0EBE
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
         Picture         =   "frmPositions.frx":0FC0
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
         Picture         =   "frmPositions.frx":10C2
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Positions"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   2805
      End
   End
End
Attribute VB_Name = "frmPositions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private myJobPositions As HRCORE.JobPositions
Private selJobPos As HRCORE.JobPosition
Private IsInEditMode As Boolean

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
End Sub

Public Sub cmdclose_Click()
    Unload Me
End Sub

Public Sub cmdDelete_Click()
    Dim resp As String
    Dim retVal As Long
    
    If Not (selJobPos Is Nothing) Then
        resp = MsgBox("This will delete Position " & selJobPos.PositionName & " from the records. Do you wish to continue?", vbQuestion + vbYesNo)
        If resp = vbNo Then
            Exit Sub
        End If
        
        Action = "DELETED EMPLOYEE POSITION; CODE: " & selJobPos.PositionCode & "; DESCRIPTION: " & selJobPos.PositionName
        
        retVal = selJobPos.Delete
        If retVal = 0 Then
            MsgBox "The Position has been deleted", vbInformation, TITLES
        Else
            MsgBox "The Position could not be deleted", vbInformation, TITLES
        End If
        
        'reload the data
        LoadJobPositions
    Else
        MsgBox "You have to select the Positions  you would like to delete.", vbInformation
    End If
        
End Sub


Public Sub cmdEdit_Click()
    Dim lngPosID As Long
    
    If lvwDetails.ListItems.count < 1 Then
        MsgBox "You have to select the Position you would like to edit.", vbInformation
        
        PSave = True
        Call cmdCancel_Click
        PSave = False
        Exit Sub
    End If
    
    lngPosID = CLng(lvwDetails.SelectedItem.Tag)
    
    'Set rs3 = CConnect.GetRecordSet("SELECT * FROM Positions WHERE Code = '" & Trim(lvwDetails.SelectedItem) & "'")
    Set selJobPos = myJobPositions.FindJobPosition(lngPosID)
    If Not (selJobPos Is Nothing) Then
        With selJobPos
            txtCode.Text = .PositionCode
            txtPositionName.Text = .PositionName
            txtDetails.Text = .PositionDetails
            txtDailyCasualRate.Text = .CasualDailyRate
            txtDefaultRemun.Text = .DefaultRemuneration
            txtMaxEmpCount.Text = .MaxEmpsInOrganogram
            
            If .ShowInOrganogram Then
                chkVisibleInOGram.value = vbChecked
            Else
                chkVisibleInOGram.value = vbUnchecked
            End If
            SaveNew = False
        End With
    Else
        MsgBox "Record not found.", vbInformation
        PSave = True
        Call cmdCancel_Click
        PSave = False
        Exit Sub
    End If
                
    Call DisableCmd
    
    fraDetails.Visible = True
    
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    SaveNew = False
    txtCode.Locked = True

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

Private Sub ClearFields()
    'By Oscar: To clear the textboxes
    Dim ctl As Control
    
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        ElseIf TypeOf ctl Is CheckBox Then
            ctl.value = vbUnchecked
        End If
    Next ctl
End Sub
Public Sub cmdNew_Click()
    Call DisableCmd
    'txtCode.Text = loadPCode
    ClearFields
    fraDetails.Visible = True
    cmdCancel.Enabled = True
    SaveNew = True
    cmdSave.Enabled = True
    txtCode.Locked = True
    txtPositionName.SetFocus
    

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
    Dim TheJobPos As New HRCORE.JobPosition
    Dim retVal As Long
    
    If PromptSave = True Then
        If MsgBox("Are you sure you want to save the record.", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    If Trim(txtPositionName.Text) = "" Then
        MsgBox "Enter the Position name", vbExclamation
        txtPositionName.SetFocus
        Exit Sub
    End If
    
    
    'set the fields
    With TheJobPos
        .PositionName = Trim(txtPositionName.Text)
        If IsNumeric(Trim(Me.txtDailyCasualRate.Text)) Then
            .CasualDailyRate = CSng(Trim(Me.txtDailyCasualRate.Text))
        Else
            If Len(Trim(Me.txtDailyCasualRate.Text)) > 0 Then
                MsgBox "Enter Numeric Data", vbInformation, TITLES
                Me.txtDailyCasualRate.SetFocus
                Exit Sub
            Else
                .CasualDailyRate = 0
            End If
        End If
        If IsNumeric(Trim(Me.txtDefaultRemun.Text)) Then
            .DefaultRemuneration = CSng(Trim(Me.txtDefaultRemun.Text))
        Else
            If Len(Trim(Me.txtDefaultRemun.Text)) > 0 Then
                MsgBox "Enter Numeric Data for Default Remuneration", vbExclamation, TITLES
                Me.txtDefaultRemun.SetFocus
                Exit Sub
            Else
                .DefaultRemuneration = 0
            End If
        End If
        If IsNumeric(Trim(Me.txtMaxEmpCount.Text)) Then
            .MaxEmpsInOrganogram = CLng(Trim(Me.txtMaxEmpCount.Text))
        Else
            If Len(Trim(Me.txtMaxEmpCount.Text)) > 0 Then
                MsgBox "Enter Numeric Data", vbExclamation, TITLES
                Me.txtMaxEmpCount.SetFocus
                Exit Sub
            Else
                .MaxEmpsInOrganogram = 0
            End If
        End If
        
        .PositionDetails = Trim(Me.txtDetails.Text)
        If Me.chkVisibleInOGram.value = vbChecked Then
            .ShowInOrganogram = True
        Else
            .ShowInOrganogram = False
        End If
        'or use if savenew
        'If (IsInEditMode = True) And (Not (selJobPos Is Nothing)) Then
        If SaveNew = True Then
            retVal = .InsertNew()
            If retVal = 0 Then
                MsgBox "The new Position has been added successfully"
            Else
                MsgBox "Update Failed"
            End If
        Else
            .PositionID = selJobPos.PositionID
            retVal = .Update()
            If retVal = 0 Then
                MsgBox "The Position has been updated successfully"
            Else
                MsgBox "Update Failed"
            End If
        End If
    End With
    
    'reload the data from db
    LoadJobPositions
    
    If SaveNew = False Then
        PSave = True
        Call cmdCancel_Click
        PSave = False
    Else
        ClearFields
        txtPositionName.SetFocus
    End If
End Sub


Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    frmMain2.PositionTheFormWithoutEmpList Me
    
    'instantiate the Job Positions
    Set myJobPositions = New HRCORE.JobPositions
    
    'set listview columns
    Call InitGrid
    
    'load jobpositions
    LoadJobPositions
    
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
     Me.Top = Me.Top - 200
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while loading Job Positions" & _
    err.Description, vbInformation, TITLES

End Sub


Private Sub LoadJobPositions()
    'By Oscar, Populate the Positions from collection into ListView
    Dim jpos As HRCORE.JobPosition
    Dim i As Long
    Dim ItemX As ListItem
    
    Me.lvwDetails.ListItems.Clear
    
    myJobPositions.GetAllJobPositions
    
    For i = 1 To myJobPositions.count
        Set jpos = myJobPositions.Item(i)
        Set ItemX = Me.lvwDetails.ListItems.add(, , jpos.PositionCode)
        ItemX.SubItems(1) = jpos.PositionName
        ItemX.SubItems(2) = jpos.ShowInOrganogram
        ItemX.SubItems(3) = jpos.MaxEmpsInOrganogram
        ItemX.SubItems(4) = jpos.PositionDetails
        ItemX.SubItems(5) = jpos.CasualDailyRate
        ItemX.SubItems(6) = jpos.DefaultRemuneration
        ItemX.Tag = jpos.PositionID
    Next i
        
End Sub

Private Sub Form_Resize()
    oSmart.FResize Me
    
    Me.Height = tvwMainheight
    Frame1.Move Frame1.Left, 0, Frame1.Width, tvwMainheight - 20
    lvwDetails.Move lvwDetails.Left, lvwDetails.Top, lvwDetails.Width, (tvwMainheight - lvwDetails.Top) - 20
End Sub

Private Sub InitGrid()
    With lvwDetails
        .ColumnHeaders.add , , "Position Code", .Width / 8
        .ColumnHeaders.add , , "Position Name", 2 * .Width / 7
        .ColumnHeaders.add , , "Show In O.Gram", .Width / 7
        .ColumnHeaders.add , , "Max. Emps In O.Gram", .Width / 7
        .ColumnHeaders.add , , "Position Details", 2 * .Width / 7
        .ColumnHeaders.add , , "Casual Daily Rate", .Width / 7
        .ColumnHeaders.add , , "Default Remuneration", .Width / 7
        
        .View = lvwReport
    End With
    

End Sub

Public Sub DisplayRecords()
    lvwDetails.ListItems.Clear
    Call Cleartxt
    
    With rs2
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Set li = lvwDetails.ListItems.add(, , !Code & "", , 5)
                li.ListSubItems.add , , !PositionName & ""
                li.ListSubItems.add , , !approved & ""
                li.ListSubItems.add , , !Comments & ""
                li.ListSubItems.add , , !dailyrate & ""
                .MoveNext
            Loop
        End If
    End With
 
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set rs2 = Nothing
    Set rs5 = Nothing
'    Set Cnn = Nothing
    
    frmMain2.lvwEmp.Visible = True
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

Private Sub lvwDetails_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim lngPosID As Long
    Set selJobPos = Nothing
    lngPosID = CLng(lvwDetails.SelectedItem.Tag)
    Set selJobPos = myJobPositions.FindJobPosition(lngPosID)
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

Private Function loadPCode() As String
    Set rs5 = CConnect.GetRecordSet("SELECT MAX(id) FROM Positions")
    If rs5.EOF = False Then
        If rs5.RecordCount > 0 And Not IsNull(rs5.Fields(0)) Then
            loadPCode = "P" & CStr(rs5.Fields(0) + 1)
        Else
            loadPCode = "P01"
        End If
    Else
        loadPCode = "P01"
    End If
    Set rs5 = Nothing
End Function

