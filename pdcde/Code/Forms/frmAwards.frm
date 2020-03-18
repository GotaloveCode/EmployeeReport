VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmAwards 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Awards"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   Icon            =   "frmAwards.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   7875
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
            Picture         =   "frmAwards.frx":0442
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAwards.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAwards.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAwards.frx":0778
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   7875
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   7800
      Begin MSComDlg.CommonDialog Cdl 
         Left            =   4800
         Top             =   4440
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
         Left            =   2760
         Picture         =   "frmAwards.frx":0CBA
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
         Left            =   3240
         Picture         =   "frmAwards.frx":0DBC
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
         Left            =   3720
         Picture         =   "frmAwards.frx":0EBE
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
         BackColor       =   &H00C0E0FF&
         Caption         =   "Awards"
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
         Height          =   3375
         Left            =   660
         TabIndex        =   17
         Top             =   1665
         Visible         =   0   'False
         Width           =   6135
         Begin VB.TextBox txtCourse 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            TabIndex        =   26
            Top             =   600
            Width           =   4335
         End
         Begin VB.ComboBox cboCurrency 
            Height          =   315
            Left            =   3000
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtAType 
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
            Left            =   4800
            TabIndex        =   3
            Top             =   1200
            Width           =   1245
         End
         Begin VB.TextBox txtCAward 
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
            Height          =   300
            Left            =   1680
            TabIndex        =   4
            Top             =   1200
            Width           =   1260
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
            Left            =   5505
            Picture         =   "frmAwards.frx":13B0
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Cancel Process"
            Top             =   2745
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
            Left            =   5025
            Picture         =   "frmAwards.frx":14B2
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Save Record"
            Top             =   2745
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
            Height          =   780
            Left            =   135
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   1815
            Width           =   5865
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
            Height          =   300
            Left            =   135
            TabIndex        =   1
            Top             =   600
            Width           =   1260
         End
         Begin MSComCtl2.DTPicker dtpFrom 
            Height          =   330
            Left            =   135
            TabIndex        =   2
            Top             =   1185
            Width           =   1335
            _ExtentX        =   2355
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
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   63045635
            CurrentDate     =   37673
            MinDate         =   21916
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            Caption         =   "Currency"
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
            Left            =   3000
            TabIndex        =   25
            Top             =   960
            Width           =   660
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            Caption         =   "Award Type"
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
            Left            =   4785
            TabIndex        =   23
            Top             =   960
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            Caption         =   "Cash Award"
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
            Left            =   1680
            TabIndex        =   22
            Top             =   960
            Width           =   870
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            Caption         =   "Award"
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
            Left            =   1470
            TabIndex        =   21
            Top             =   375
            Width           =   465
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
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
            TabIndex        =   20
            Top             =   1590
            Width           =   750
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
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
            TabIndex        =   19
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            Caption         =   "Date"
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
            Width           =   345
         End
      End
      Begin MSComctlLib.ListView lvwDetails 
         Height          =   7800
         Left            =   0
         TabIndex        =   0
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
         BackColor       =   16777215
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
Attribute VB_Name = "frmAwards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

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

Private Sub cmdclose_Click()
    Unload Me
End Sub

Public Sub cmdDelete_Click()
    Dim resp As String
    On Error GoTo ErrHandler
    If Not currUser Is Nothing Then
        If currUser.CheckRight("Award") <> secModify Then
            MsgBox "You dont have right to delete the record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
     If SelectedEmployee Is Nothing Then Exit Sub
     
    If lvwDetails.ListItems.count > 0 Then
        resp = MsgBox("Are you sure you want to delete  " & lvwDetails.SelectedItem & " from the records?", vbQuestion + vbYesNo)
        If resp = vbNo Then
            Exit Sub
        End If
        Action = "DELETED AWARD FROM EMPLOYEE; EMPLOYEE CODE: " & SelectedEmployee.EmpCode & "; AWARD CODE: " & lvwDetails.SelectedItem
        CConnect.ExecuteSql ("DELETE FROM Awards WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & lvwDetails.SelectedItem & "'")
        rs2.Requery
        Call DisplayRecords
    Else
        MsgBox "You have to select the award you would like to delete.", vbInformation
                
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

    If Not currUser Is Nothing Then
        If currUser.CheckRight("Award") <> secModify Then
            MsgBox "You dont have right to modify the record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    If SelectedEmployee Is Nothing Then Exit Sub
    
    If lvwDetails.ListItems.count < 1 Then
        MsgBox "You have to select the award you would like to edit.", vbInformation
        PSave = True
        Call cmdCancel_Click
        PSave = False
        Exit Sub
    End If
    Set rs3 = CConnect.GetRecordSet("SELECT * FROM Awards WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND code = '" & lvwDetails.SelectedItem & "'")
    With rs3
        If .RecordCount > 0 Then
            txtCode.Text = !Code & ""
            txtCode.Tag = txtCode.Text
            If Not IsNull(!cFrom) Then dtpFrom.value = !cFrom & ""
            txtCAward.Text = Format(!CAward & "", "#,###,##0.00")
            txtComments.Text = !Comments & ""
            txtCourse.Text = !Course & ""
            txtAType.Text = !AType & ""
            
            If Not IsNull(!CurrencyID) Then
                cboCurrency.ListIndex = getIndex(!CurrencyID)
            Else
                cboCurrency.ListIndex = -1
            End If
                       
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
    txtCode.Locked = False
    txtCourse.SetFocus
    
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
    
    If Not currUser Is Nothing Then
        If currUser.CheckRight("Award") <> secModify Then
            MsgBox "You dont have right to add new record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    Call DisableCmd
    
    With txtCode
        .Text = loadACode
        .Locked = False
    End With
    
    txtComments.Text = ""
    txtCourse.Text = ""
    dtpFrom.value = Date
    fraDetails.Visible = True
    cmdCancel.Enabled = True
    SaveNew = True
    cmdSave.Enabled = True
    'txtCode.Locked = True
    txtCourse.SetFocus
    'dtpFrom.Value = Date
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
    Dim CuRindex As Long
    If SelectedEmployee Is Nothing Then Exit Sub
    
    If txtCode.Text = "" Then
        MsgBox "Enter the award code.", vbExclamation
        txtCode.SetFocus
        Exit Sub
    End If
    
    If txtCourse.Text = "" Then
        MsgBox "Enter the course.", vbExclamation
        txtCourse.SetFocus
        Exit Sub
    End If
    
    If txtCAward.Text = "" Then
        txtCAward.Text = 0
    End If
        

    If SaveNew = True Then
        
        Set rs4 = CConnect.GetRecordSet("SELECT * FROM Awards WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & txtCode.Text & "'")
        
        With rs4
            If .RecordCount > 0 Then
                MsgBox "Award code already exists. Enter another one.", vbInformation
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
    If cboCurrency.ListCount > 0 And cboCurrency.ListIndex <> -1 Then
        CuRindex = cboCurrency.ItemData(cboCurrency.ListIndex)
    End If
    
    CConnect.ExecuteSql ("DELETE FROM Awards WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & txtCode.Tag & "'")
    
    mySQL = "INSERT INTO Awards (employee_id, Code, Course, CFrom, Comments, CAward, AType,CurrencyId)" & _
                        " VALUES('" & SelectedEmployee.EmployeeID & "','" & txtCode.Text & "','" & txtCourse.Text & "'," & _
                        "'" & SQLDate(dtpFrom.value) & "','" & txtComments.Text & "'," & Format(txtCAward.Text, Nfmt) & ",'" & txtAType.Text & "'," & CuRindex & ")"
  
    Action = "ADDED AWARD TO EMPLOYEE; EMPLOYEE CODE: " & SelectedEmployee.EmpCode & "; AWARD CODE: " & txtCode.Text & "; AWARD TYPE: " & txtAType.Text & "; CASH: " & txtCAward.Text
    
    CConnect.ExecuteSql (mySQL)
       
    rs2.Requery
    txtCode.Tag = ""
    If SaveNew = False Then
        PSave = True
        Call DisplayRecords
        Call cmdCancel_Click
        PSave = False
    Else
        
        Call DisplayRecords
        txtAType.SetFocus
        SaveNew = False
        txtCode.Text = loadACode
    End If
    
    
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
    
    Set rs2 = CConnect.GetRecordSet("SELECT * FROM Awards ORDER BY Code")
    
    If (Not rsGlob Is Nothing) Then
        With rsGlob
            If .RecordCount < 1 Then
                Call DisableCmd
                Exit Sub
            End If
        End With
    End If
    Call DisplayRecords
    
    cmdFirst.Enabled = False
    cmdPrevious.Enabled = False
    
    Rcurrency.Requery
    If Not (Rcurrency Is Nothing) Then
        If Not (Rcurrency.EOF Or Rcurrency.BOF) Then
            
            cboCurrency.Clear
            Rcurrency.MoveFirst
            Do Until Rcurrency.EOF
                cboCurrency.AddItem Rcurrency!CURRENCY_NAME
                cboCurrency.ItemData(cboCurrency.NewIndex) = Rcurrency!CURRENCY_ID
                Rcurrency.MoveNext
            Loop
        End If
    End If
       Exit Sub
ErrHandler:
    MsgBox "An error has occur: " & err.Description
End Sub

Private Sub Form_Resize()
    oSmart.FResize Me
    Me.Height = tvwMainheight - 150
    Frame1.Move Frame1.Left, 0, Frame1.Width, tvwMainheight - 150
    lvwDetails.Move lvwDetails.Left, 0, lvwDetails.Width, tvwMainheight - 150
    
End Sub

Private Sub InitGrid()
    With lvwDetails
        .ColumnHeaders.add , , "Code", 0
        .ColumnHeaders.add , , "Award", 4000
        .ColumnHeaders.add , , "Date", 1400
        .ColumnHeaders.add , , "Cash Award", , vbRightJustify
        .ColumnHeaders.add , , "Award Type"
        .ColumnHeaders.add , , "Comments", 3500
   
        .View = lvwReport
    End With

End Sub

Public Sub DisplayRecords()
    If SelectedEmployee Is Nothing Then Exit Sub
    
    lvwDetails.ListItems.Clear
    Call Cleartxt
    
    With rsGlob
        If Not .EOF And Not .BOF Then
            
            With rs2
                If .RecordCount > 0 Then
                    .Filter = "employee_id like '" & SelectedEmployee.EmployeeID & "'"
                    If .RecordCount > 0 Then
                        .MoveFirst
                        Do While Not .EOF
                            Set li = lvwDetails.ListItems.add(, , !Code & "", , 5)
                            li.ListSubItems.add , , !Course & ""
                            li.ListSubItems.add , , Format(!cFrom & "", "dd-mm-yyyy")
                            li.ListSubItems.add , , Format(!CAward & "", Cfmt)
                            li.ListSubItems.add , , !AType & ""
                            li.ListSubItems.add , , !Comments & ""
                            li.Tag = Trim(!ID & "")
                            .MoveNext
                        Loop
                    End If
                    .Filter = adFilterNone
                End If
            End With
            
        End If
    End With
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



Private Sub txtCAward_KeyPress(KeyAscii As Integer)
    If Len(Trim(txtCAward.Text)) > 9 Then
        Beep
        MsgBox "Can't enter more than 9 characters", vbExclamation
        KeyAscii = 8
    End If
    
    Select Case KeyAscii
      Case Asc("0") To Asc("9")
      Case Asc(".")
      Case Is = 8
      Case Else
          Beep
          KeyAscii = 0
    End Select

End Sub

Private Sub txtCAward_LostFocus()
    On Error Resume Next
    txtCAward.Text = Format(txtCAward & "", Cfmt)
End Sub

Private Sub txtCode_Change()
    txtCode.Text = UCase(txtCode.Text)
    txtCode.SelStart = Len(txtCode.Text)
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
            If i.Name <> "cboCurrency" Then i.Text = ""
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

Private Sub txtCourse_KeyPress(KeyAscii As Integer)
If Len(Trim(txtCourse.Text)) > 198 Then
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

Private Function loadACode() As String
    Set rs5 = CConnect.GetRecordSet("SELECT MAX(id) FROM Awards")
    If rs5.EOF = False Then
        If rs5.RecordCount > 0 And Not IsNull(rs5.Fields(0)) Then
            loadACode = "D" & CStr(rs5.Fields(0) + 1)
        Else
            loadACode = "D01"
        End If
    Else
        loadACode = "D01"
    End If
    Set rs5 = Nothing
End Function

Private Function getIndex(data As Long) As Long
    Dim i As Long
    getIndex = -1
    For i = 0 To cboCurrency.ListCount - 1
        If cboCurrency.ItemData(i) = data Then
            getIndex = i
            Exit For
        End If
    Next i
End Function
