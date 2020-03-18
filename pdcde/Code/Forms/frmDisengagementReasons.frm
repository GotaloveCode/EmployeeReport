VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDisengagementReasons 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraDisenReasons 
      Height          =   7900
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11130
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
         Left            =   7215
         Picture         =   "frmDisengagementReasons.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Delete Record"
         Top             =   6750
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
         Left            =   6720
         Picture         =   "frmDisengagementReasons.frx":04F2
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Edit Record"
         Top             =   6750
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton Command1 
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
         Left            =   6240
         Picture         =   "frmDisengagementReasons.frx":05F4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Add New record"
         Top             =   6750
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame fradata 
         Caption         =   "Disengagement Reasons"
         Height          =   1935
         Left            =   1920
         TabIndex        =   2
         Top             =   1920
         Width           =   5055
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
            Left            =   3600
            Picture         =   "frmDisengagementReasons.frx":06F6
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Add New record"
            Top             =   1440
            Visible         =   0   'False
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
            Left            =   4080
            Picture         =   "frmDisengagementReasons.frx":07F8
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Save Record"
            Top             =   1440
            Width           =   495
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
            Left            =   4560
            Picture         =   "frmDisengagementReasons.frx":08FA
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Cancel Process"
            Top             =   1440
            Width           =   495
         End
         Begin VB.TextBox txtDescription 
            Height          =   375
            Left            =   1080
            TabIndex        =   6
            Top             =   960
            Width           =   3855
         End
         Begin VB.TextBox txtcode 
            Height          =   375
            Left            =   1080
            TabIndex        =   4
            Top             =   360
            Width           =   3855
         End
         Begin VB.Label Label2 
            Caption         =   "Description"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Code "
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   735
         End
      End
      Begin MSComctlLib.ListView lvwDisReasons 
         Height          =   5040
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   11130
         _ExtentX        =   19632
         _ExtentY        =   8890
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmDisengagementReasons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Reasons As DisengagementReasons
Private Action As String
Private selectedReason As disengagementReason

Private Sub cmdCancel_Click()
    fradata.Visible = False
    Me.txtCode.Text = ""
    Me.txtDescription.Text = ""
    Action = "save"
    frmMain2.cmdDelete.Enabled = True
    frmMain2.cmdNew.Enabled = True
    frmMain2.cmdCancel.Enabled = False
    frmMain2.cmdSave.Enabled = False
    
End Sub

Private Sub cmdDelete_Click()
''-------------

    Dim resp As String
    On Error GoTo ErrHandler
    If Not currUser Is Nothing Then
        If currUser.CheckRight("SetUpDisengagementReasons") <> secModify Then
            MsgBox "You dont have right to delete the record. Please liaise with the security admin"
            Exit Sub
        End If
        
     
    End If
    
    
     
    If lvwDisReasons.ListItems.count > 0 Then
        resp = MsgBox("Are you sure you want to delete  " & lvwDisReasons.SelectedItem & " from the records?", vbQuestion + vbYesNo)
        If resp = vbNo Then
            Exit Sub
        End If
        Action = "DELETED DISENGAGEMENT REASON. CODE: " & lvwDisReasons.SelectedItem
        CConnect.ExecuteSql ("UPDATE disengagementreasons set deleted=1 WHERE reasonid = '" & lvwDisReasons.SelectedItem.Tag & "'")
       
        Call DisplayRecords
    Else
        MsgBox "You have to select the Disengagement Reason  you would like to delete.", vbInformation
                
    End If
    Exit Sub
ErrHandler:
MsgBox (err.Description)

''--------------




'    On Error Resume Next
'    Dim resp As String
'
'    ''place deleting code here
    
    Call DisplayRecords
End Sub

Public Sub cmdNew_Click()



  If Not currUser Is Nothing Then
        If currUser.CheckRight("SetUpDisengagementReasons") <> secModify Then
            MsgBox "You dont have right to delete the record. Please liaise with the security admin"
            Exit Sub
        End If
        
       
    End If
    
    
     Action = "save"
     txtCode.Text = ""
     txtDescription.Text = ""
     fradata.Visible = True
     
     With frmMain2
        .fracmd.Visible = True
        .fracmd.Enabled = True
        
        .cmdNew.Enabled = True
        .cmdCancel.Enabled = False
        .cmdDelete.Enabled = False
        .cmdSave.Enabled = False
        .cmdEdit.Enabled = False
    End With
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrHandler
     Dim retVal As Boolean
    Select Case Action
        Case "save"
            If txtCode.Text = "" Or txtDescription.Text = "" Then
                MsgBox "Please fill all the data"
                Exit Sub
            Else
                If Reasons.FindTheReason(Trim(txtDescription.Text)) Then
                    MsgBox "That disengagement reason is extisting"
                    Exit Sub
                End If
                
                Dim DisReason As New disengagementReason
                DisReason.Code = txtCode.Text
                DisReason.Reason = txtDescription.Text
                retVal = DisReason.Insert
                
                If retVal Then
                    Call DisplayRecords
                    txtCode.Text = ""
                    txtDescription.Text = ""
                Else
                    MsgBox "Data was not successfully saved"
                End If
            End If
        
        Case "update"
            If txtCode.Text = "" Or txtDescription.Text = "" Then
                MsgBox "Please fill all the data"
                Exit Sub
            Else
                
                selectedReason.Code = txtCode.Text
                selectedReason.Reason = txtDescription.Text
                retVal = selectedReason.Update
                If retVal Then
                
                    Call DisplayRecords
                    Me.txtCode.Text = ""
                    Me.txtDescription.Text = ""
                    Action = "save"
                    Set selectedReason = Nothing
                Else
                    MsgBox "data has nt been update"
                End If
            End If
            
    End Select
    Exit Sub
ErrHandler:
    MsgBox "An error has occur when " & IIf(Action = "save", "Saving", "Updating") & "data"
End Sub

Private Sub Form_Load()
     On Error GoTo ErrHandler
   
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
    
    Call InitGrid
    
    Set Reasons = New DisengagementReasons
    Call DisplayRecords
    
    With frmMain2
        .fracmd.Visible = True
        .fracmd.Enabled = True
        
        .cmdNew.Enabled = True
        .cmdCancel.Enabled = False
        .cmdDelete.Enabled = False
        .cmdSave.Enabled = False
        .cmdEdit.Enabled = False
    End With
    fradata.Visible = False
    Exit Sub
ErrHandler:
    MsgBox "An error has occur. ERROR DESCRIPTION: " & err.Description
End Sub

Private Sub Form_Resize()
    oSmart.FResize Me
    FraDisenReasons.Move FraDisenReasons.Left, FraDisenReasons.Top, FraDisenReasons.Width, tvwMainheight - 140
    Me.lvwDisReasons.Height = tvwMainheight - 140
End Sub

Private Sub InitGrid()
    With lvwDisReasons
        .ColumnHeaders.add , , "Code", .Width / 3
        .ColumnHeaders.add , , "Disengagement Reason", (.Width * 2 / 3)
        .View = lvwReport
    End With
End Sub

Private Sub Cleartxt()
    Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Then
            i.Text = ""
        End If
    Next i
    lvwDisReasons.ListItems.Clear
End Sub

Public Sub DisplayRecords()
    Dim i As Long
    Dim ItemX As ListItem
    lvwDisReasons.ListItems.Clear
    Reasons.GetallDisengagementReasons
    
    If Reasons.count > 0 Then
        For i = 1 To Reasons.count
        
         If Reasons.Item(i).Deleted = 0 Then '-- record has not been deleted
            Select Case Trim(Reasons.Item(i).Reason)
                Case "Death", "Retirement"
                    Set ItemX = lvwDisReasons.ListItems.add(, , Reasons.Item(i).Code)
                    ItemX.SubItems(1) = Reasons.Item(i).Reason
                    ItemX.Tag = Reasons.Item(i).ReasonID
                Case Else
                    Set ItemX = lvwDisReasons.ListItems.add(, , Reasons.Item(i).Code)
                    ItemX.SubItems(1) = Reasons.Item(i).Reason
                    ItemX.Tag = Reasons.Item(i).ReasonID
            End Select
         End If
        
        Next i
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    With frmMain2
'        .fracmd.Visible = True
'        .fracmd.Enabled = True
'
'        .cmdNew.Enabled = True
'        .cmdCancel.Enabled = True
'        .cmdDelete.Enabled = False
'        .cmdSave.Enabled = False
'        .cmdEdit.Enabled = False
'    End With
End Sub

Private Sub lvwDisReasons_Click()
    If lvwDisReasons.ListItems.count > 0 Then
        'display details for editing
       frmMain2.cmdDelete.Enabled = False
       frmMain2.cmdNew.Enabled = False
       frmMain2.cmdCancel.Enabled = True
       frmMain2.cmdSave.Enabled = True
        Call DisplayRecord(lvwDisReasons.SelectedItem.Tag)
    End If
End Sub

Private Sub lvwDisReasons_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwDisReasons
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvwDisReasons_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'display details for editing
    Call DisplayRecord(Item.Tag)
End Sub

Private Sub DisplayRecord(data As String)

    Set selectedReason = Reasons.FindReason(CLng(data))
    Action = "update"
    txtCode.Text = selectedReason.Code
    txtDescription.Text = selectedReason.Reason
    fradata.Visible = True
End Sub


