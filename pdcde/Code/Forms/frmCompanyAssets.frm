VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmCompanyAssets 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company Assets"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      Caption         =   "Company Assets"
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
      Height          =   4080
      Left            =   1440
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   7140
      Begin VB.Frame Frame2 
         Height          =   1575
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   6855
         Begin VB.TextBox txtRegNo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            TabIndex        =   34
            Top             =   1185
            Width           =   5055
         End
         Begin VB.TextBox txtMake 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            TabIndex        =   33
            Top             =   705
            Width           =   1815
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
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   240
            Width           =   1815
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
            Height          =   285
            Left            =   4800
            Locked          =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtModel 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4800
            TabIndex        =   20
            Top             =   712
            Width           =   1935
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Asset code"
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
            TabIndex        =   26
            Top             =   255
            Width           =   1455
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
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
            Height          =   255
            Left            =   3720
            TabIndex        =   25
            Top             =   255
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Make"
            Height          =   255
            Left            =   960
            TabIndex        =   24
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Model Number"
            Height          =   255
            Left            =   3600
            TabIndex        =   23
            Top             =   727
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Registration Number"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   1200
            Width           =   1455
         End
      End
      Begin VB.Frame fraDriver 
         Height          =   1095
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Width           =   6855
         Begin VB.TextBox txtAgentName 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4800
            TabIndex        =   30
            Top             =   240
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker dtExpiry 
            Height          =   285
            Left            =   4800
            TabIndex        =   29
            Top             =   705
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   503
            _Version        =   393216
            Format          =   63045633
            CurrentDate     =   39148
         End
         Begin MSComCtl2.DTPicker dtStartdate 
            Height          =   285
            Left            =   1680
            TabIndex        =   27
            Top             =   720
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            Format          =   63045633
            CurrentDate     =   39148
         End
         Begin VB.TextBox txtPlicyNo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            TabIndex        =   17
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Agent Name"
            Height          =   255
            Left            =   3360
            TabIndex        =   31
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Expiry Date"
            Height          =   255
            Left            =   3720
            TabIndex        =   28
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Comence Date"
            Height          =   255
            Left            =   480
            TabIndex        =   18
            Top             =   735
            Width           =   1095
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Ins. Policy Number"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   255
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Asset Type"
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   6855
         Begin VB.OptionButton optOthers 
            Caption         =   "Others"
            Height          =   255
            Left            =   4560
            TabIndex        =   14
            Top             =   240
            Width           =   2175
         End
         Begin VB.OptionButton OptVehicle 
            Caption         =   "Vehicle"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
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
         Left            =   6525
         Picture         =   "frmCompanyAssets.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Cancel Process"
         Top             =   3585
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
         Left            =   6045
         Picture         =   "frmCompanyAssets.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Save Record"
         Top             =   3585
         Width           =   495
      End
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
      TabIndex        =   10
      ToolTipText     =   "Move to the First employee"
      Top             =   5670
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
      TabIndex        =   9
      Top             =   5670
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
      TabIndex        =   8
      ToolTipText     =   "Move to the Previous employee"
      Top             =   5670
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
      TabIndex        =   7
      ToolTipText     =   "Move to the Next employee"
      Top             =   5670
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
      Picture         =   "frmCompanyAssets.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Add New record"
      Top             =   5670
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
      Picture         =   "frmCompanyAssets.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Edit Record"
      Top             =   5670
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
      Picture         =   "frmCompanyAssets.frx":0408
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Delete Record"
      Top             =   5640
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
      TabIndex        =   3
      ToolTipText     =   "Move to the Last employee"
      Top             =   5670
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComctlLib.ImageList imgEmpTool 
      Left            =   3840
      Top             =   1080
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
            Picture         =   "frmCompanyAssets.frx":08FA
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanyAssets.frx":0A0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanyAssets.frx":0B1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompanyAssets.frx":0C30
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwDetails 
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
Attribute VB_Name = "frmCompanyAssets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub DisplayRecords()
    Dim s As ListItem
    On Error GoTo errhandle
    rs1.Requery
    With rs1
        lvwDetails.ListItems.Clear
        While .EOF = False
            Set s = lvwDetails.ListItems.add(, , Trim(!assetCode & ""))
            s.ListSubItems.add , , Trim(!assetname & "")
            s.ListSubItems.add , , Trim(!Make & "")
            s.ListSubItems.add , , Trim(!Model & "")
            s.ListSubItems.add , , Trim(!RegistrationNumber & "")
            If !IsVehicle Then
                s.ListSubItems.add , , "Vehicle"
            Else
                s.ListSubItems.add , , "Others"
            End If
            s.ListSubItems.add , , !PolicyNo & ""
            
            s.Tag = Trim(!ID & "")
            .MoveNext
        Wend
    End With
    Exit Sub
errhandle:
    MsgBox err.Description
End Sub
Public Sub InitGrid()
    'clear collumn headers
    With lvwDetails
        .ColumnHeaders.Clear
        .ColumnHeaders.add , , "Asset Code", 0 '.Width / 5
        .ColumnHeaders.add , , "Asset Name", .Width / 6
        .ColumnHeaders.add , , "Make ", .Width / 6
        .ColumnHeaders.add , , "Model", .Width / 6
        .ColumnHeaders.add , , "Registration Number", .Width / 6
        .ColumnHeaders.add , , "Asset type", .Width / 6
        .ColumnHeaders.add , , "Insurance Policy No.", .Width / 6
        
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
    On Error GoTo ErrHandler
    
    If Not currUser Is Nothing Then
        If currUser.CheckRight("CompanyAssets") <> secModify Then
            MsgBox "You dont have right to delete record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    If lvwDetails.ListItems.count = 0 Then MsgBox "There are no items to delete.", vbOKOnly + vbInformation, "Empty list": Exit Sub
    If MsgBox("Are you sure you want to delete " & lvwDetails.SelectedItem.ListSubItems(1).Text & "?", vbYesNo + vbQuestion, "Confirm delete") = vbNo Then Exit Sub
    
    Dim rs As ADODB.Recordset
    Set rs = CConnect.GetRecordSet("Select * from companyAssetissue WHERE id=" & lvwDetails.SelectedItem.Tag)
    If Not rs Is Nothing And rs.RecordCount <= 0 Then
        Action = "DELETED COMPANY ASSET; ASSET NAME: " & lvwDetails.SelectedItem.ListSubItems(1)
        CConnect.ExecuteSql "DELETE FROM CompanyAssets WHERE id='" & lvwDetails.SelectedItem.Tag & "'"
        lvwDetails.ListItems.remove lvwDetails.SelectedItem.Index
        MsgBox "Record successfully deleted!", vbOKOnly + vbInformation, "Deleted"
    Else
        MsgBox "The record can not be deleted because that asset is already issused"
    End If
    
    Set rs = Nothing
    Exit Sub
ErrHandler:
    MsgBox "An error has occured when deleting asset details " & err.Description
End Sub

Public Sub cmdEdit_Click()
    If Not currUser Is Nothing Then
        If currUser.CheckRight("CompanyAssets") <> secModify Then
            MsgBox "You dont have right to edit record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    If lvwDetails.ListItems.count = 0 Then MsgBox "There are no records to edit.", vbOKOnly + vbInformation, "No records": Exit Sub
    txtAssetCode.Text = lvwDetails.SelectedItem.Text
    txtAssetName.Text = lvwDetails.SelectedItem.ListSubItems(1).Text
    Me.txtMake.Text = lvwDetails.SelectedItem.ListSubItems(2).Text
    Me.txtModel.Text = lvwDetails.SelectedItem.ListSubItems(3).Text
    Me.txtRegNo.Text = lvwDetails.SelectedItem.ListSubItems(4).Text
    txtAssetCode.Tag = lvwDetails.SelectedItem.Tag
    txtAssetName.Tag = txtAssetCode.Text
    
    If lvwDetails.SelectedItem.ListSubItems(5).Text = "Vehicle" Then
        OptVehicle.value = True
    Else
        optOthers.value = True
    End If
    
    txtAssetCode.Locked = False
    fraDetails.Visible = True
End Sub

Public Sub cmdNew_Click()

    If Not currUser Is Nothing Then
        If currUser.CheckRight("CompanyAssets") <> secModify Then
            MsgBox "You dont have right to add new record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    Dim rsAssets As New ADODB.Recordset
    Set rsAssets = CConnect.GetRecordSet("select * from CompanyAssets")
    If rsAssets.RecordCount = 0 Then
        txtAssetCode.Text = "AST-" & Val(rsAssets.RecordCount + 1)
    Else
        Set rsAssets = CConnect.GetRecordSet("select max(id) as MyID from CompanyAssets")
        If rsAssets.RecordCount > 0 Then
            txtAssetCode.Text = "AST-" & Val(rsAssets!MyID + 1)
        End If
    End If
    
    fraDetails.Visible = True
    OptVehicle.value = False
    optOthers.value = True
    fraDriver.Enabled = False
    
    With txtAssetName
        .Text = ""
        .SetFocus
    End With
    'txtComments.Text = ""
    txtAssetCode.Locked = False
End Sub

Public Sub cmdSave_Click()
    If txtAssetName.Text = "" Then MsgBox "Please specify the asset name.", vbOKOnly + vbInformation, "Missing asset name": Exit Sub
    If optOthers.value = False And OptVehicle.value = False Then MsgBox "Please specify the type of asset.", vbOKOnly + vbInformation, "Missing asset infor": Exit Sub
    
    Dim IsVehicle As Integer
    
    If OptVehicle.value = True Then
        IsVehicle = 1
    Else
        IsVehicle = 0
    End If
    
    If txtAssetCode.Tag = "" Then
        Action = "ADDED COMPANY ASSET; ASSET CODE: " & txtAssetCode.Text & "; ASSET NAME: " & txtAssetCode.Text
        CConnect.ExecuteSql "INSERT INTO CompanyAssets(AssetCode,AssetName,make,model,Registrationnumber,PolicyNo,IsVehicle,StartDate,ExpiryDate,InsuranceName) VALUES ('" & txtAssetCode.Text & "','" & txtAssetName.Text & "','" & txtMake.Text & "','" & txtModel.Text & "','" & txtRegNo.Text & "','" & txtPlicyNo.Text & "'," & IsVehicle & ",'" & SQLDate(dtStartdate.value) & "','" & SQLDate(dtExpiry.value) & "','" & txtAgentName.Text & "')"
    Else
        Action = "UPDATED COMPANY ASSET; ASSET CODE: " & txtAssetCode.Text & "; ASSET NAME: " & txtAssetCode.Text & "; COMMENTS: " '& txtComments.Text
        CConnect.ExecuteSql "UPDATE CompanyAssets SET PolicyNo='" & txtPlicyNo.Text & "',IsVehicle=" & IsVehicle & ",StartDate='" & SQLDate(dtStartdate.value) & "',ExpiryDate='" & SQLDate(dtExpiry.value) & "',InsuranceName='" & txtAgentName.Text & "', make='" & txtMake.Text & "', model='" & txtModel.Text & "',Registrationnumber='" & txtRegNo.Text & "',AssetCode='" & txtAssetCode.Text & "',AssetName='" & txtAssetName.Text & "' WHERE id='" & txtAssetCode.Tag & "'"
        If txtAssetCode.Text <> txtAssetName.Tag Then
        mySQL = "UPDATE CompanyAssetIssue set AssetCode='" & txtAssetCode.Text & "'" & " where AssetCode='" & txtAssetName.Tag & "'"
         CConnect.ExecuteSql mySQL
        End If
    End If
    Call DisplayRecords
    txtAssetCode.Tag = ""
    txtAssetCode.Text = ""
    txtAssetName.Text = ""
    'txtComments.Text = ""
    txtModel.Text = ""
    txtMake.Text = ""
    Me.txtRegNo.Text = ""
    Call cmdNew_Click
End Sub

Private Sub Form_Load()
 
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
    txtAssetName.Locked = False
    'txtComments.Locked = False
    Call InitGrid
    
    Set rs1 = CConnect.GetRecordSet("SELECT * FROM CompanyAssets")
    Call DisplayRecords
    
    cmdFirst.Enabled = False
    cmdPrevious.Enabled = False
End Sub

Private Sub Form_Resize()
    oSmart.FResize Me
    
    Me.Height = tvwMainheight - 150
    lvwDetails.Move lvwDetails.Left, 0, lvwDetails.Width, tvwMainheight - 150
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain2.Caption = "Infiniti HRMIS - PDR [Current User:\" & currUser.FullNames & "]"
End Sub
Private Sub lvwDetails_DblClick()
    Call cmdEdit_Click
End Sub

Private Sub optOthers_Click()
     If OptVehicle.value = True Then
        fraDriver.Enabled = False
    Else
        fraDriver.Enabled = True
    End If
End Sub

Private Sub OptVehicle_Click()
    If OptVehicle.value = True Then
        fraDriver.Enabled = True
    Else
        fraDriver.Enabled = False
    End If
End Sub
