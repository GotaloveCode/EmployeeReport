VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmAssetIssue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asset Issue"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   17
      ToolTipText     =   "Move to the First employee"
      Top             =   5910
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
      TabIndex        =   16
      Top             =   5910
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
      TabIndex        =   15
      ToolTipText     =   "Move to the Previous employee"
      Top             =   5910
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
      TabIndex        =   14
      ToolTipText     =   "Move to the Next employee"
      Top             =   5910
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
      Picture         =   "frmAssetIssue.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Add New record"
      Top             =   5910
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
      Picture         =   "frmAssetIssue.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Edit Record"
      Top             =   5910
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
      Picture         =   "frmAssetIssue.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Delete Record"
      Top             =   5910
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
      TabIndex        =   10
      ToolTipText     =   "Move to the Last employee"
      Top             =   5910
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Asset Issue"
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
      Height          =   4200
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   6180
      Begin MSComctlLib.ListView lvwAsseList 
         Height          =   2415
         Left            =   1680
         TabIndex        =   19
         Top             =   840
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   4260
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
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
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0FF&
         Height          =   1095
         Left            =   120
         TabIndex        =   20
         Top             =   2520
         Width           =   5895
         Begin VB.TextBox txtrLicenceNo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            TabIndex        =   21
            Top             =   240
            Width           =   3855
         End
         Begin MSComCtl2.DTPicker dtExpiryDate 
            Height          =   255
            Left            =   4320
            TabIndex        =   22
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            Format          =   63045633
            CurrentDate     =   39148
         End
         Begin MSComCtl2.DTPicker dtStartdate 
            Height          =   255
            Left            =   1320
            TabIndex        =   23
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            Format          =   63045633
            CurrentDate     =   39148
         End
         Begin VB.Label Label8 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Driver Licence Number"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label7 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Commence Date"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Expiry Date"
            Height          =   255
            Left            =   3360
            TabIndex        =   24
            Top             =   720
            Width           =   1095
         End
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
         Height          =   1290
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1125
         Width           =   5790
      End
      Begin VB.CommandButton cmdCancel 
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
         Left            =   5565
         Picture         =   "frmAssetIssue.frx":06F6
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Cancel Process"
         Top             =   3705
         Width           =   495
      End
      Begin VB.CommandButton cmdSave 
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
         Left            =   5085
         Picture         =   "frmAssetIssue.frx":07F8
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Save Record"
         Top             =   3705
         Width           =   495
      End
      Begin VB.TextBox txtAssetCode 
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtAssetName 
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
         Height          =   330
         Left            =   1680
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   480
         Width           =   4215
      End
      Begin VB.CommandButton cmdSchAsset 
         Height          =   315
         Left            =   1200
         Picture         =   "frmAssetIssue.frx":08FA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   315
      End
      Begin VB.Label Label1 
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
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "Asset Code"
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
         TabIndex        =   8
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "Asset Name"
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
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComctlLib.ImageList imgEmpTool 
      Left            =   3840
      Top             =   600
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
            Picture         =   "frmAssetIssue.frx":0C84
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssetIssue.frx":0D96
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssetIssue.frx":0EA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssetIssue.frx":0FBA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwDetails 
      Height          =   7800
      Left            =   0
      TabIndex        =   18
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
Attribute VB_Name = "frmAssetIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private infDetailHeight, FrLicenceTop, intCmdSaveTop As Integer
Private TypeOfasset As String

Public Sub DisplayRecords()
    On Error GoTo ErrHandler
    'CLEARING THE LIST VIEW CONTROL
    lvwDetails.ListItems.Clear
    
    If Not SelectedEmployee Is Nothing Then
        Set rs1 = CConnect.GetRecordSet("SELECT ci.employee_id,ci.id,ca.IsVehicle,ca.assetcode,ca.assetname,ca.Make,ca.Model,ci.comments FROM CompanyAssetIssue ci INNER JOIN CompanyAssets ca ON ca.assetcode=ci.assetcode WHERE ci.employee_id='" & SelectedEmployee.EmployeeID & "'")
         If rs1 Is Nothing Then Exit Sub
            If Not (rs1.EOF Or rs1.BOF) Then
                With lvwDetails
                    .ListItems.Clear
                    'rs1.Filter = "employee_id='" & rsGlob!employee_id & "'"
                    While rs1.EOF = False
                        Set li = lvwDetails.ListItems.add(, , Trim(rs1!assetCode & ""))
                        li.ListSubItems.add , , Trim(rs1!assetname & "")
                        li.ListSubItems.add , , Trim(rs1!Make & "")
                        li.ListSubItems.add , , Trim(rs1!Model & "")
                        li.ListSubItems.add , , Trim(rs1!Comments & "")
                        If rs1!IsVehicle Then
                            li.ListSubItems.add , , "Vehicle"
                        Else
                            li.ListSubItems.add , , "Others"
                        End If
                        rs1.MoveNext
                    Wend
                End With
            End If
        End If
    
    Exit Sub
ErrHandler:
    MsgBox "An error has occured when  Displaying the allocated asset " & err.Description
End Sub

Public Sub InitGrid()
    With lvwDetails
        .ColumnHeaders.Clear
        .ColumnHeaders.add , , "Asset code", 0 '.Width / 7
        .ColumnHeaders.add , , "Asset Name", .Width * (2 / 9)
        .ColumnHeaders.add , , "Make", .Width * (2 / 9)
        .ColumnHeaders.add , , "Model", .Width * (2 / 9)
        .ColumnHeaders.add , , "Comments", .Width / 3 '.Width - (.Width / 3 + .Width / 7)
        .ColumnHeaders.add , , "Asset Type", 0
        
        .View = lvwReport
    End With
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

End Sub

Public Sub cmdDelete_Click()
    If Not currUser Is Nothing Then
        If currUser.CheckRight("IssueAsset") <> secModify Then
            MsgBox "You dont have right to delete record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    If SelectedEmployee Is Nothing Then
        MsgBox "please select the employee"
        Exit Sub
    End If
    
    If lvwDetails.ListItems.count > 0 Then
        If MsgBox("Are you sure you want to detach " & UCase(lvwDetails.SelectedItem.ListSubItems(1).Text) & " from" & vbCrLf & frmMain2.lvwEmp.SelectedItem.ListSubItems(1).Text & "?", vbYesNo + vbQuestion, "Confirm delete") = vbNo Then Exit Sub
        
        Action = "DETACHED ASSET FROM EMPLOYEE; EMPLOYEE CODE: " & SelectedEmployee.EmpCode & "; ASSET CODE: " & txtAssetCode.Text & "; ASSET NAME: " & lvwDetails.SelectedItem.Text
        CConnect.ExecuteSql "DELETE FROM CompanyAssetIssue WHERE employee_id='" & SelectedEmployee.EmployeeID & "' AND assetcode='" & lvwDetails.SelectedItem.Text & "'"
        Call DisplayRecords
    Else
        MsgBox "There are no records to delete.", vbOKOnly + vbInformation, "Cannot deleted"
    End If
    
End Sub

Public Sub cmdEdit_Click()
   
    'check for user rights
     If Not currUser Is Nothing Then
        If currUser.CheckRight("IssueAsset") <> secModify Then
            MsgBox "You dont have right to edit record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    If lvwDetails.ListItems.count > 0 Then
        txtAssetCode.Text = lvwDetails.SelectedItem.Text
        txtAssetName.Text = lvwDetails.SelectedItem.ListSubItems(1).Text
        txtComments.Text = lvwDetails.SelectedItem.ListSubItems(4).Text
       Call resizeControls(lvwDetails.SelectedItem.ListSubItems(5).Text)
       fraDetails.Visible = True
        
        cmdSave.Enabled = True
        'cmdEdit.Enabled = True
        cmdCancel.Enabled = True
        txtComments.Locked = False
    
    Else
        MsgBox "Cannot edit non-existent record.", vbOKOnly + vbInformation, "Edit Failed"
    End If
End Sub

Public Sub cmdNew_Click()
    If Not currUser Is Nothing Then
        If currUser.CheckRight("IssueAsset") <> secModify Then
            MsgBox "You dont have right to add new record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
'Enable  the cntrols
    If Not SelectedEmployee Is Nothing Then
        cmdSave.Enabled = True
        cmdEdit.Enabled = True
        cmdCancel.Enabled = True
        txtComments.Locked = False
        fraDetails.Visible = True
    Else
        MsgBox "Please select the employee you want to assign the asset"
    End If
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrHandler
    
     If Not currUser Is Nothing Then
        If currUser.CheckRight("IssueAsset") <> secModify Then
            MsgBox "You dont have right to edit record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    If Me.txtAssetCode.Text = "" Or Me.txtAssetName = "" Then
        MsgBox "Select the asset "
        Exit Sub
    End If
    
    Dim rsAsset As New ADODB.Recordset
    Set rsAsset = CConnect.GetRecordSet("select * from CompanyAssetIssue where employee_id='" & SelectedEmployee.EmployeeID & "' and assetcode='" & txtAssetCode.Text & "'")
    If rsAsset.RecordCount > 0 Then
        If MsgBox("The specfied asset had already been assigned the selected employee." & vbCrLf & "Do you want to edit the records?", vbYesNo + vbQuestion, "Confirm update") = vbYes Then
            
            Action = "ISSUED ASSET TO EMPLOYEE; EMPLOYEE CODE: " & SelectedEmployee.EmpCode & "; ASSET CODE: " & txtAssetCode.Text & "; ASSET NAME: " & txtAssetName.Text
            
            CConnect.ExecuteSql "UPDATE CompanyAssetIssue SET comments='" & Replace(txtComments.Text, "'", "''") & "' WHERE employee_id='" & SelectedEmployee.EmployeeID & "' AND assetCode='" & txtAssetCode.Text & "'"
            Call DisplayRecords
            txtAssetCode.Text = ""
            txtComments.Text = ""
            txtAssetName.Text = ""
            MsgBox "Records successfully updated!", vbOKOnly + vbInformation, "Update successful"
            Exit Sub
        Else
            Exit Sub
        End If
    End If
    
    Action = "ISSUED ASSET TO EMPLOYEE; EMPLOYEE CODE: " & SelectedEmployee.EmpCode & "; ASSET CODE: " & txtAssetCode.Text & "; ASSET NAME: " & txtAssetName.Text
    If TypeOfasset = "Vehicle" Then
        CConnect.ExecuteSql "INSERT INTO CompanyAssetIssue(employee_id,assetCode,Comments,DriverLicenceNo, StartDate, Expirydate) VALUES('" & SelectedEmployee.EmployeeID & "','" & txtAssetCode.Text & "','" & Replace(txtComments.Text, "'", "''") & "','" & txtrLicenceNo.Text & "','" & SQLDate(dtStartdate.value) & "','" & SQLDate(dtExpiryDate.value) & "')"
    Else
        CConnect.ExecuteSql "INSERT INTO CompanyAssetIssue(employee_id,assetCode,Comments) VALUES('" & SelectedEmployee.EmployeeID & "','" & txtAssetCode.Text & "','" & Replace(txtComments.Text, "'", "''") & "')"
    End If
    Call DisplayRecords
    MsgBox "Asset successfully assigned.", vbOKOnly + vbInformation, "Asset issue"
    txtAssetCode.Text = ""
    txtComments.Text = ""
    txtAssetName.Text = ""
    txtrLicenceNo.Text = ""
    Exit Sub
    
ErrHandler:
    MsgBox "An error has occured when assigning an asset to the employee"
End Sub

Private Sub cmdSchAsset_Click()
    Dim rsSelAsset, IssuedAssets As New ADODB.Recordset
    Dim i As Long
    Dim issued As Boolean
    With lvwAsseList
        .ColumnHeaders.Clear
        .ColumnHeaders.add , , "CODE", 2 * .Width / 4
        .ColumnHeaders.add , , "ASSET NAME", 2 * .Width / 4
        .ColumnHeaders.add , , "Asset Type", 0
                
        .View = lvwReport
        .FullRowSelect = True
    End With
    'get the issues asset and all the assets
    Set IssuedAssets = CConnect.GetRecordSet("select * from CompanyAssetIssue order by Assetcode asc")
    Set rsSelAsset = CConnect.GetRecordSet("select * from CompanyAssets order by Assetcode asc")
    
    If rsSelAsset Is Nothing Then
        MsgBox "Please set up the assets"
        Exit Sub
    End If
    
    'If Not (IssuedAssets Is Nothing) Then
        lvwAsseList.ListItems.Clear
        If Not (rsSelAsset.BOF And rsSelAsset.EOF) Then
            rsSelAsset.MoveFirst
            Do Until rsSelAsset.EOF
                issued = False
                If Not IssuedAssets.BOF Then IssuedAssets.MoveFirst
                Do Until IssuedAssets.EOF
                    If IssuedAssets!assetCode = rsSelAsset!assetCode Then
                     issued = True
                        Exit Do
                    End If
                    IssuedAssets.MoveNext
                Loop
                
                If Not issued Then
                    Set li = lvwAsseList.ListItems.add(, , Trim(rsSelAsset!assetCode & ""))
                    li.ListSubItems.add , , Trim(rsSelAsset!assetname & "")
                    If rsSelAsset!IsVehicle Then
                        li.ListSubItems.add , , "Vehicle"
                    Else
                        li.ListSubItems.add , , "Others"
                    End If
                             
                End If
                rsSelAsset.MoveNext
            Loop
        Else
            MsgBox "There are no company assets in the database to be issued out to employees", vbExclamation
        End If
      
      If lvwAsseList.ListItems.count > 0 Then
        lvwAsseList.Visible = True
      Else
         lvwAsseList.Visible = False
      End If
        
    If lvwAsseList.ListItems.count <= 0 And Not (rsSelAsset.BOF And rsSelAsset.EOF) Then
        MsgBox "No more asset to Issue. They have all been issued"
    End If
'End If

    Set rsSelAsset = Nothing
    Set IssuedAssets = Nothing
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
    
    Call DisplayRecords
    
    cmdFirst.Enabled = False
    cmdPrevious.Enabled = False
    Me.txtAssetCode.Locked = True
    Me.txtAssetName.Locked = True
    infDetailHeight = fraDetails.Height
    FrLicenceTop = Frame4.Top
    
    intCmdSaveTop = cmdSave.Top
    Exit Sub
ErrHandler:
    MsgBox "An error has occur: " & err.Description
End Sub

Private Sub Form_Resize()
    oSmart.FResize Me
    Me.Height = tvwMainheight - 100

    lvwDetails.Move lvwDetails.Left, 0, lvwDetails.Width, tvwMainheight - 100
End Sub



Private Sub Form_Unload(Cancel As Integer)
    frmMain2.Caption = "Infiniti HRMIS - PDR [Current User:\" & currUser.FullNames & "]"
End Sub

Private Sub lvwAsseList_DblClick()
    With lvwAsseList
        txtAssetCode.Text = .SelectedItem.Text
        txtAssetName.Text = .SelectedItem.ListSubItems(1).Text
        TypeOfasset = .SelectedItem.ListSubItems(2).Text
        Call resizeControls(.SelectedItem.ListSubItems(2).Text)
    End With
    lvwAsseList.Visible = False
End Sub

Private Sub lvwDetails_DblClick()
    Call cmdEdit_Click
End Sub


Private Sub resizeControls(assetType As String)
    If assetType = "Vehicle" Then
        Frame4.Enabled = True
        Frame4.Move txtComments.Left, FrLicenceTop
        Frame4.Visible = True
        fraDetails.Move fraDetails.Left, fraDetails.Top, fraDetails.Width, infDetailHeight
        
       'move the buttons
        cmdSave.Move cmdSave.Left, intCmdSaveTop
        cmdCancel.Move cmdCancel.Left, intCmdSaveTop
    Else
        Frame4.Enabled = False
        Frame4.Move txtComments.Left, txtComments.Top
        Frame4.Visible = False
        'fraDetails.Move fraDetails.Left, fraDetails.Top, fraDetails.Width, (lvwAsseList.Top + lvwAsseList.Height)
        
        'move the buttons
        cmdSave.Move cmdSave.Left, (fraDetails.Height - cmdSave.Height)
        cmdCancel.Move cmdCancel.Left, (fraDetails.Height - cmdSave.Height)
    End If
End Sub
