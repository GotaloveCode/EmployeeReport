VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Bio-Data"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   Icon            =   "frmBio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   7815
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
            Picture         =   "frmBio.frx":0442
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBio.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBio.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBio.frx":0778
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7860
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   7800
      Begin MSComDlg.CommonDialog Cdl 
         Left            =   1545
         Top             =   4080
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
         Picture         =   "frmBio.frx":0CBA
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
         Picture         =   "frmBio.frx":0DBC
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
         Picture         =   "frmBio.frx":0EBE
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
         Caption         =   "Bio-Data"
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
         Height          =   3645
         Left            =   600
         TabIndex        =   17
         Top             =   1440
         Visible         =   0   'False
         Width           =   5925
         Begin VB.ComboBox cboBGroup 
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
            ItemData        =   "frmBio.frx":13B0
            Left            =   135
            List            =   "frmBio.frx":13CC
            TabIndex        =   3
            Top             =   1155
            Visible         =   0   'False
            Width           =   2370
         End
         Begin VB.TextBox txtContact 
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
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   1155
            Width           =   5655
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
            Left            =   5310
            Picture         =   "frmBio.frx":13FA
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Cancel Process"
            Top             =   3015
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
            Left            =   4800
            Picture         =   "frmBio.frx":14FC
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Save Record"
            Top             =   3015
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
            Height          =   705
            Left            =   135
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   1830
            Width           =   5655
         End
         Begin VB.TextBox txtCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            TabIndex        =   1
            Top             =   615
            Width           =   1260
         End
         Begin VB.TextBox txtDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   1500
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   615
            Width           =   4290
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Details"
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
            TabIndex        =   21
            Top             =   945
            Width           =   480
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
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
            TabIndex        =   20
            Top             =   1575
            Width           =   750
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
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
            Left            =   120
            TabIndex        =   19
            Top             =   375
            Width           =   375
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
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
            Left            =   1500
            TabIndex        =   18
            Top             =   375
            Width           =   795
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
Attribute VB_Name = "frmBio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Sub cmdCancel_Click()
'    If PSave = False Then
'        If PromptSave = True Then
'            If MsgBox("Do you want to save changes?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
'            Call cmdSave_Click
'        Else
'            fraDetails.Visible = False
'            With frmMain2
'                .cmdNew.Enabled = True
'                .cmdEdit.Enabled = True
'                .cmdDelete.Enabled = True
'                .cmdCancel.Enabled = False
'                .cmdSave.Enabled = False
'            End With
'        End If
'        Call EnableCmd
'        Exit Sub
'    End If
'
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
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("BioData") <> secModify Then
            MsgBox "You Don't have right to delete the record. Please liaise with security admin"
            Exit Sub
        End If
   End If
    
    If Not (SelectedEmployee Is Nothing) Then
        If lvwDetails.ListItems.count > 0 Then
            resp = MsgBox("Are you sure you want to delete  " & lvwDetails.SelectedItem.SubItems(1) & " from the records?", vbQuestion + vbYesNo)
            If resp = vbNo Then
                Exit Sub
            End If
              
            Action = "DELETED BIO DATA DETAIL code " & lvwDetails.SelectedItem.Text & "  " & lvwDetails.SelectedItem.SubItems(1) & " of employee " & SelectedEmployee.EmpCode
            CConnect.ExecuteSql ("DELETE FROM Bio WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & lvwDetails.SelectedItem & "'")
            rs2.Requery
            
            Call DisplayRecords
                
        Else
            MsgBox "You have to select the BioData you would like to delete.", vbInformation
                    
        End If
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
    
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("BioData") <> secModify Then
            MsgBox "You Don't have right to edit record. Please liaise with security admin"
            Exit Sub
        End If
   End If
   
    If lvwDetails.ListItems.count < 1 Then
        MsgBox "You have to select the BioData you would like to edit.", vbInformation
        PSave = True
        Call cmdCancel_Click
        PSave = False
    
        Exit Sub
    End If
    rs1.Requery
    
    With rs1
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "Code like '" & lvwDetails.SelectedItem & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                Call DispDetails
            
            Else
                MsgBox "Record not found.", vbInformation
                Exit Sub
    
            End If
        Else
            MsgBox "Record not found.", vbInformation
            Exit Sub
        End If
    End With
  
Call DisableCmd

fraDetails.Visible = True

cmdSave.Enabled = True
cmdCancel.Enabled = True
SaveNew = False
Exit Sub
'txtContact.SetFocus
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
    
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("BioData") <> secModify Then
            MsgBox "You Don't have right to add new record. Please liaise with security admin"
            Exit Sub
        End If
   End If
    If SelectedEmployee Is Nothing Then
        MsgBox "Pease select employee to record his\her bio data"
        Exit Sub
    End If
    
    Call DisableCmd
    txtCode.Text = ""
        txtDesc.Text = ""
        txtContact.Text = ""
        txtComments.Text = ""
    
    With rs1
    If .RecordCount > 0 Then
        .MoveFirst
        Call DispDetails
            
    Else
        MsgBox "BioData types have not been defined yet.", vbInformation
        Exit Sub
    End If
    End With
       
    fraDetails.Visible = True
    cmdCancel.Enabled = True
    SaveNew = True
    cmdSave.Enabled = True
    'txtContact.SetFocus
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
    
    If Not (currUser Is Nothing) Then
        If currUser.CheckRight("BioData") <> secModify Then
            MsgBox "You Don't have right to edit record. Please liaise with security admin"
            Exit Sub
        End If
   End If
   
   If SelectedEmployee Is Nothing Then
        MsgBox "Please select employee"
        Exit Sub
   End If

    If txtCode.Text = "" Then
        MsgBox "Enter the bio data code.", vbExclamation
        txtCode.SetFocus
        Exit Sub
    End If
    
    If txtDesc.Text = "" Then
        MsgBox "Enter the bio data description.", vbExclamation
        txtDesc.SetFocus
        Exit Sub
    End If
    
    If cboBGroup.Visible = True Then
        txtContact.Text = cboBGroup.Text
    End If
    
    If PromptSave = True Then
        If MsgBox("Are you sure you want to save the record.", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
         CConnect.ExecuteSql ("delete from Bio WHERE employee_id=" & SelectedEmployee.EmployeeID & "  AND Code='" & txtCode.Text & "'")
    
        Action = "ADDED EMPLOYEE BIO-DATA; EMPLOYEE CODE: " & SelectedEmployee.EmpCode & "; BIO-DATA: " & txtContact.Text & "; COMMENTS: " & txtComments.Text
        mySQL = "INSERT INTO Bio (employee_id, Code, BioData, Comments)" & _
                            " VALUES('" & SelectedEmployee.EmployeeID & "','" & txtCode.Text & "'," & _
                            "'" & txtContact.Text & "','" & txtComments.Text & "')"
       CConnect.ExecuteSql (mySQL)
      txtCode.Tag = ""
    rs2.Requery
    With rs1
        If .RecordCount > 0 Then
            If Not .EOF Then
                .MoveNext
                If Not .EOF Then
                    Call DisplayRecords
                   Call DispDetails
                    txtContact.SetFocus
                    Call Decla.DisableCmd
                Else
                    Call DisplayRecords
                    PSave = True
                    Call cmdCancel_Click
                    PSave = False
                End If
            Else
                Call DisplayRecords
                PSave = True
                Call cmdCancel_Click
                PSave = False
            End If
        Else
            Call DisplayRecords
            PSave = True
            Call cmdCancel_Click
            PSave = False
        End If
    End With
End Sub


Private Sub Form_Load()
    On Error GoTo ErrHandler
    oSmart.FReset Me
    If oSmart.hRatio > 1.1 Then
        With frmMain2
            Me.Move .tvwMain.Width + .lvwEmp.Width + (.tvwMain.Width / 36) * 2, (.Height / 5.52) '- 155
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
    'Call 'CConnect.CCon
    
    Set rs1 = CConnect.GetRecordSet("SELECT * FROM BTypes ORDER BY Code")
    Set rs2 = CConnect.GetRecordSet("SELECT * FROM Bio ORDER BY Code")
    Set rs4 = CConnect.GetRecordSet("SELECT * FROM BTypes ORDER BY Code")
    
    
    With rsGlob
        If .RecordCount < 1 Then
            Call DisableCmd
            Exit Sub
        End If
    End With
    If frmMain2.lvwEmp.ListItems.count > 0 Then
        Call DisplayRecords
    End If
    cmdFirst.Enabled = False
    cmdPrevious.Enabled = False
    Exit Sub
ErrHandler:
    MsgBox "An error has occured " & err.Description
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
        .ColumnHeaders.add , , "Bio Data Description", .Width / 5
        .ColumnHeaders.add , , "Value", .Width / 3
        .ColumnHeaders.add , , "Comments", 7 * .Width / 15
   
        .View = lvwReport
    End With
    

End Sub

Public Sub DisplayRecords()
    On Error GoTo Hell
    lvwDetails.ListItems.Clear
    Call Cleartxt
        If Not (SelectedEmployee Is Nothing) Then
            With rsGlob
            If Not .EOF And Not .BOF Then
                With rs2
                    If .RecordCount > 0 Then
                      .Filter = "employee_id like '" & SelectedEmployee.EmployeeID & "'"
                    
                    End If
                End With
    
                With rs4
                    If .RecordCount > 0 Then
                        .MoveFirst
                        Do While Not .EOF
                            Set li = lvwDetails.ListItems.add(, , !Code & "")
                            li.ListSubItems.add , , !Description & ""
                            li.Tag = !Code
                              With rs2
                                    If .RecordCount > 0 Then
                                        .MoveFirst
                                        .Find "Code like '" & rs4!Code & "'", , adSearchForward, adBookmarkFirst
                                        If Not .EOF Then
                                            li.ListSubItems.add , , !BioData & ""
                                            li.ListSubItems.add , , !Comments & ""
                                        Else
                                            li.ListSubItems.add , , ""
                                            li.ListSubItems.add , , ""
                                         
                                        End If
                                    Else
                                        li.ListSubItems.add , , ""
                                        li.ListSubItems.add , , ""
                                       
                                    End If
    
                            End With
    
                            .MoveNext
                        Loop
                        .MoveFirst
                    End If
                End With
    
                rs2.Filter = adFilterNone
            End If
        End With
        End If
        Exit Sub
Hell:
        MsgBox "The System has encouter a problem when displaying employee data" & _
            vbNewLine & "If this persist, please consult the Support Team"
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
            i.Text = ""
        End If
    Next i
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


Public Sub DispDetails()
    Call Cleartxt
    'rs1.Requery
    With rs1
        If Not .BOF And Not .EOF Then
            If Not .EOF Then
                txtCode.Text = !Code & ""
                txtCode.Tag = txtCode.Text
                txtDesc.Text = !Description & ""
                If !Description = "Blood Group" Then
                    cboBGroup.Visible = True
                    txtContact.Visible = False
                Else
                    cboBGroup.Visible = False
                    txtContact.Visible = True
                End If
                
                Set rs3 = CConnect.GetRecordSet("SELECT * FROM Bio WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & rs1!Code & "'")
                
                With rs3
                    If .RecordCount > 0 Then
                        txtComments.Text = !Comments & ""
                        txtContact.Text = !BioData & ""
               
                    End If
                End With
                
                Set rs3 = Nothing
    
            End If
        End If
    End With

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
  Case Asc("+")
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

Private Sub txtContact_KeyPress(KeyAscii As Integer)
If Len(Trim(txtContact.Text)) > 198 Then
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
  Case Asc("+")
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
