VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBankBranchSetUp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Branches"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11130
   Icon            =   "frmBankBranchSetUp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDetails 
      Caption         =   "Bank Branch Set-up"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3030
      Left            =   3480
      TabIndex        =   5
      Top             =   1020
      Visible         =   0   'False
      Width           =   6375
      Begin VB.TextBox txtBcode 
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   1110
      End
      Begin VB.TextBox txtBname 
         Height          =   315
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txtBankName 
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
         Height          =   315
         Left            =   3180
         TabIndex        =   1
         Top             =   840
         Width           =   3015
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
         Left            =   1080
         TabIndex        =   0
         Top             =   840
         Width           =   1110
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
         Height          =   825
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1455
         Width           =   6075
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
         Left            =   5235
         Picture         =   "frmBankBranchSetUp.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Save Record"
         Top             =   2415
         Width           =   510
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5730
         Picture         =   "frmBankBranchSetUp.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancel Process"
         Top             =   2415
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Bank code"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Bank Name"
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Name"
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
         Left            =   2280
         TabIndex        =   8
         Top             =   900
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Code"
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
         TabIndex        =   7
         Top             =   900
         Width           =   915
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
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
         TabIndex        =   6
         Top             =   1230
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7665
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   11100
      Begin MSComctlLib.ListView LvBankBranches 
         Height          =   7605
         Left            =   4200
         TabIndex        =   23
         Top             =   0
         Width           =   6930
         _ExtentX        =   12224
         _ExtentY        =   13414
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
         TabIndex        =   21
         ToolTipText     =   "Move to the Last employee"
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
         Picture         =   "frmBankBranchSetUp.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Delete Record"
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
         Picture         =   "frmBankBranchSetUp.frx":0B38
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Edit Record"
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
         Picture         =   "frmBankBranchSetUp.frx":0C3A
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Add New record"
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
         TabIndex        =   17
         ToolTipText     =   "Move to the Next employee"
         Top             =   5400
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
         TabIndex        =   16
         ToolTipText     =   "Move to the Previous employee"
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
         Top             =   5400
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
      Begin MSComctlLib.ListView lvwDetails 
         Height          =   7605
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   4170
         _ExtentX        =   7355
         _ExtentY        =   13414
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
      Left            =   6390
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
            Picture         =   "frmBankBranchSetUp.frx":0D3C
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankBranchSetUp.frx":0E4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankBranchSetUp.frx":0F60
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankBranchSetUp.frx":1072
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBankBranchSetUp"
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
    txtCode.Tag = ""
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Public Sub cmdDelete_Click()
    Dim resp As String
    If LvBankBranches.ListItems.count = 0 Then Exit Sub
    If LvBankBranches.SelectedItem <> "" Then
        If lvwDetails.ListItems.count > 0 Then
            resp = MsgBox("This Action will delete  " & LvBankBranches.SelectedItem.SubItems(1) & " Bank Branch" & vbNewLine & "and employee bank acounts of this Branch. Do you wish to continue?", vbQuestion + vbYesNo)
            If resp = vbNo Then
                Exit Sub
            End If
            
            Action = "DELETED BANK BRANCH; BRANCH CODE: " & LvBankBranches.SelectedItem.Text & "  " & LvBankBranches.SelectedItem.SubItems(1)
            
            CConnect.ExecuteSql ("DELETE FROM tblBankBranch WHERE BankBranch_Code = '" & LvBankBranches.SelectedItem & "'")
            
            ' Delete the bank records of this Branch from employee details
            CConnect.ExecuteSql ("DELETE FROM employeebanks WHERE branchID =" & LvBankBranches.SelectedItem.Tag)
            Action = "DELETED EMPLOYEE BANK BRANCH; BRANCH CODE: " & LvBankBranches.SelectedItem.Text & "  " & LvBankBranches.SelectedItem.SubItems(1)
            ' Delete the bank records of this Branch from Company banks
            CConnect.ExecuteSql ("DELETE FROM tblCompanybank WHERE Bankbranch_ID =" & LvBankBranches.SelectedItem.Tag)
            Action = "DELETED COMPANY BANK BRANCH; BRANCH CODE: " & LvBankBranches.SelectedItem.Text & "  " & LvBankBranches.SelectedItem.SubItems(1)
            
            rs2.Requery
            Display_BankBranch_Records lvwDetails.SelectedItem.Tag
        Else
            MsgBox "You have to select the Bank Branch Details you would like to Delete.", vbInformation
        End If
    End If

End Sub

Private Sub cmdDone_Click()
    PSave = True
    Call cmdCancel_Click
    PSave = False
End Sub

Public Sub cmdEdit_Click()

    If lvwDetails.ListItems.count < 1 Then
        MsgBox "You have to select the Bank Branch Details you would like to edit.", vbInformation
        PSave = True
        Call cmdCancel_Click
        PSave = False
        Exit Sub
    End If
    
    Set rs3 = CConnect.GetRecordSet("SELECT * FROM tblBankBranch WHERE bankBranch_Code = '" & LvBankBranches.SelectedItem & "'")
    
    With rs3
        If .RecordCount > 0 Then
            Me.txtBname = lvwDetails.SelectedItem.SubItems(1)
            Me.txtBcode.Text = lvwDetails.SelectedItem.Text
            txtCode.Text = !BankBranch_Code & ""
            txtCode.Tag = txtCode.Text
            txtBankName.Text = !BANKBRANCH_NAME & ""
            txtComments.Text = ""
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
End Sub

Private Sub cmdFirst_Click()
With rsGlob
    If .RecordCount > 0 Then
        If .BOF <> True Then
            .MoveFirst
            If .BOF = True Then
                .MoveFirst
                Call Display_Bank_Records
            Else
                Call Display_Bank_Records
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
                Call Display_Bank_Records
            Else
                Call Display_Bank_Records
            End If
            
            Call LastDisb
            
        End If
    End If
End With
End Sub

Public Sub cmdNew_Click()
    If lvwDetails.SelectedItem.Selected = False Then
        MsgBox "please select the bank"
        Exit Sub
    End If
    
    Call DisableCmd
    txtCode.Text = ""
    txtBankName = ""
    txtComments.Text = ""
    fraDetails.Visible = True
    txtBcode.Text = lvwDetails.SelectedItem.Text
    txtBname.Text = lvwDetails.SelectedItem.SubItems(1)
    cmdCancel.Enabled = True
    SaveNew = True
    cmdSave.Enabled = True
    txtCode.Locked = False
    txtCode.SetFocus
End Sub

Private Sub cmdNext_Click()
    With rsGlob
        If .RecordCount > 0 Then
            If .EOF <> True Then
                .MoveNext
                If .EOF = True Then
                    .MoveLast
                    Call Display_Bank_Records
                Else
                    Call Display_Bank_Records
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
                Call Display_Bank_Records
            Else
                Call Display_Bank_Records
            End If
            
            Call FirstDisb
            
        End If
    End If
End With

End Sub

Public Sub cmdSave_Click()
    If txtCode.Text = "" Then
        MsgBox "Enter the Bank Branch Details code.", vbExclamation
        txtCode.SetFocus
        Exit Sub
    End If
    
    ' test for duplicate
    If Me.LvBankBranches.ListItems.count > 0 Then
         Dim ItemX As ListItem
         Set ItemX = Me.LvBankBranches.FindItem(txtCode.Text, lvwText)
         If Not ItemX Is Nothing Then
             If txtCode.Tag = "" Then
                 MsgBox "Branch code already exists. Enter another one.", vbInformation
                txtCode.Text = ""
                txtCode.SetFocus
                Exit Sub
             End If
        End If
    End If
          
    If PromptSave = True Then
        If MsgBox("Are you sure you want to save the record.", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    If SaveNew = True Then
        Action = "ADDED BANK BRANCH; BRANCH CODE: " & txtCode.Text & "; BRANCH NAME: " & txtBankName.Text & "; COMMENTS: " & txtComments.Text
        mySQL = "INSERT INTO tblBankBranch (Bank_id,bankBranch_Code,bankBranch_Name,Bank_Name,comments)" & _
                        " VALUES('" & lvwDetails.SelectedItem.Tag & "','" & txtCode.Text & "','" & txtBankName.Text & "','" & txtBname.Text & "' ,'" & Replace(txtComments.Text, "'", "''") & "')"
        CConnect.ExecuteSql (mySQL)
    Else
        Action = "UPDATED BANK BRANCH; BRANCH CODE: " & txtCode.Text & "; BRANCH NAME: " & txtBankName.Text & "; COMMENTS: " & txtComments.Text
        CConnect.ExecuteSql ("UPDATE tblBankBranch SET bankbranch_name = '" & txtBankName.Text & "',comments='" & Replace(txtComments.Text, "'", "''") & "' WHERE bankBranch_Code = '" & txtCode.Text & "'")
    End If

    rs2.Requery
    
    If SaveNew = False Then
        PSave = True
        Call cmdCancel_Click
        PSave = False
    Else
        txtCode.SetFocus
    End If
    
    Display_BankBranch_Records lvwDetails.SelectedItem.Tag
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
    Me.txtCode.Tag = ""
End Sub

Private Sub Form_Load()
 On Error GoTo Hell
 
    oSmart.FReset Me
    
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
    
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    
    Call Make_Bank_Columns
    Call Make_BankBranch_Columns
    'Call 'CConnect.CCon
    
    Set rs2 = CConnect.GetRecordSet("SELECT * FROM tblBank ORDER BY Bank_id")
    
    Call Display_Bank_Records
    
    cmdFirst.Enabled = False
    cmdPrevious.Enabled = False
    
    Exit Sub
Hell:
    MsgBox "An error has occur: " & err.Description
End Sub

Private Sub Form_Resize()
    oSmart.FResize Me
    
    Me.Height = tvwMainheight - 150
    Frame1.Move Frame1.Left, 0, Frame1.Width, tvwMainheight - 150
    LvBankBranches.Move LvBankBranches.Left, 0, LvBankBranches.Width, tvwMainheight - 150
    lvwDetails.Move lvwDetails.Left, 0, lvwDetails.Width, tvwMainheight - 150
End Sub

Private Sub Make_Bank_Columns()
    With lvwDetails
        .ColumnHeaders.add , , "Bank Code", .Width / 6
        .ColumnHeaders.add , , "Bank Name", 5 * .Width / 6
        '.ColumnHeaders.Add , , "Comments", 3500
        .View = lvwReport
    End With
End Sub

Private Sub Make_BankBranch_Columns()
    With LvBankBranches
        .ColumnHeaders.add , , "Branch Code", .Width / 7
        .ColumnHeaders.add , , "Branch Name", 3 * .Width / 7
        .ColumnHeaders.add , , "Comments", 3 * .Width / 7
        .View = lvwReport
    End With
End Sub

Public Sub Display_Bank_Records()
        lvwDetails.ListItems.Clear
        Call Cleartxt
        With rs2
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    Set li = lvwDetails.ListItems.add(, , !Bank_Code & "", , 5)
                    li.ListSubItems.add , , !Bank_Name & ""
                    li.ListSubItems.add , , !bank_Comments & ""
                    li.Tag = Trim(!bank_id & "")
                    .MoveNext
                Loop
            End If
        End With
End Sub

Function Display_BankBranch_Records(sBankCode)
    'Display the branches of the selected bank
    Dim rsBankBranches As Recordset
    
    Set rsBankBranches = CConnect.GetRecordSet("SELECT * FROM tblBankBranch where (BANK_ID='" & sBankCode & "') ORDER BY bankBranch_Name")
    
    LvBankBranches.ListItems.Clear
    'Call cleartxt
    With rsBankBranches
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Set li = LvBankBranches.ListItems.add(, , !BankBranch_Code & "", , 5)
                li.ListSubItems.add , , !BANKBRANCH_NAME & ""
                li.ListSubItems.add , , !Comments & ""
                li.Tag = !Bankbranch_id
                .MoveNext
            Loop
        End If
    End With

    Set rsBankBranches = Nothing
    If fraDetails.Visible Then
        txtBcode.Text = lvwDetails.SelectedItem.Text
        txtBname.Text = lvwDetails.SelectedItem.SubItems(1)
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set rs2 = Nothing
    Set rs5 = Nothing
'    Set Cnn = Nothing
    frmMain2.Caption = "Infiniti HRMIS - PDR [Current User:\" & currUser.FullNames & "]"
    
End Sub

Private Sub fraList_DragDrop(Source As Control, X As Single, y As Single)

End Sub

Private Sub LvBankBranches_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With LvBankBranches
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub LvBankBranches_DblClick()
    If frmMain2.cmdEdit.Enabled = True Then
        Call frmMain2.cmdEdit_Click
    End If
End Sub

Private Sub lvwDetails_Click()
    If lvwDetails.ListItems.count > 0 Then
        Display_BankBranch_Records lvwDetails.SelectedItem.Tag
    End If
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
    If lvwDetails.ListItems.count > 0 Then
        Display_BankBranch_Records lvwDetails.SelectedItem.Tag
    End If
End Sub

Private Sub txtCode_Change()
'    txtCode.Text = UCase(txtCode.Text)
'    txtCode.SelStart = Len(txtCode.Text)
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


