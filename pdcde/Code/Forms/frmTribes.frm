VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTribes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ETHINICITY"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   Icon            =   "frmTribes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4320
      TabIndex        =   10
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1440
      TabIndex        =   9
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5280
      TabIndex        =   8
      Top             =   5880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3360
      TabIndex        =   7
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2400
      TabIndex        =   6
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   480
      TabIndex        =   5
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame fraExisting 
      Caption         =   "Existing Ethnic Groups"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4875
      Left            =   60
      TabIndex        =   3
      Top             =   1020
      Width           =   6045
      Begin MSComctlLib.ListView lvwTribes 
         Height          =   4515
         Left            =   90
         TabIndex        =   4
         Top             =   270
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   7964
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nationality"
            Object.Width           =   7056
         EndProperty
      End
   End
   Begin VB.Frame fraDetails 
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
      Height          =   915
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   6045
      Begin VB.TextBox txtTribe 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   990
         TabIndex        =   2
         Top             =   360
         Width           =   4875
      End
      Begin VB.Label Label1 
         Caption         =   "Ethnic"
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
         Left            =   180
         TabIndex        =   1
         Top             =   450
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmTribes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private MyTribes As HRCORE.Tribes
Private selTribe As HRCORE.Tribe
Private blnEditMode As Boolean

Public Sub cmdCancel_Click()
    If blnEditMode = True Then
        cmdNew.Enabled = True
        txtTribe.Locked = True
        fraExisting.Enabled = True
        LoadTribes False
    Else
        cmdDelete.Enabled = True
        txtTribe.Locked = True
        fraExisting.Enabled = True
        LoadTribes False
    End If
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Public Sub cmdDelete_Click()
    Dim retVal As Long
    On Error GoTo ErrorHandler
    If Not currUser Is Nothing Then
        If currUser.CheckRight("Tribe") <> secModify Then
            MsgBox "You don't have right to delete record. Please liaise with the security admin"
            Exit Sub
        End If
     End If
     
    If Not (selTribe Is Nothing) Then
        If MsgBox("Are you sure you want to delete: " & UCase(selTribe.Tribe), vbYesNo + vbInformation, TITLES) = vbYes Then
            retVal = selTribe.Delete()
            LoadTribes True
        End If
    Else
        MsgBox "Select the Tribe to Delete", vbInformation, TITLES
    End If
           
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while deleting the Tribe" & vbNewLine & err.Description, vbInformation, TITLES
End Sub

Public Sub cmdEdit_Click()
    If Not currUser Is Nothing Then
        If currUser.CheckRight("Tribe") <> secModify Then
            MsgBox "You don't have right to modify record. Please liaise with the security admin"
            Exit Sub
        End If
     End If
    If selTribe Is Nothing Then
        MsgBox "Select the Tribe to Edit", vbInformation, TITLES
        Exit Sub
    End If
    txtTribe.Locked = False
    fraExisting.Enabled = False
    cmdNew.Enabled = False
    txtTribe.SetFocus
    blnEditMode = True
        
  
End Sub

Public Sub cmdNew_Click()
    If Not currUser Is Nothing Then
        If currUser.CheckRight("Tribe") <> secModify Then
            MsgBox "You don't have right to add new record. Please liaise with the security admin"
            Exit Sub
        End If
     End If
     
    txtTribe.Locked = False
    fraExisting.Enabled = False
    cmdDelete.Enabled = False
    txtTribe.Text = ""
    txtTribe.SetFocus
    blnEditMode = False
    
End Sub

Public Sub cmdSave_Click()
    If blnEditMode = True Then
        If Not Update Then Exit Sub
        cmdNew.Enabled = True
        txtTribe.Locked = True
        fraExisting.Enabled = True
        LoadTribes True
    Else
        If Not InsertNew Then Exit Sub
        cmdDelete.Enabled = True
        txtTribe.Locked = True
        fraExisting.Enabled = True
        LoadTribes True
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    Set MyTribes = New HRCORE.Tribes
    frmMain2.PositionTheFormWithoutEmpList Me
    
    'initialize the listview
    Me.lvwTribes.ColumnHeaders.Clear
    Me.lvwTribes.ColumnHeaders.add , , "Ethnic", Me.lvwTribes.Width
    
    LoadTribes True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An slight error has occurred" & vbNewLine & err.Description, vbInformation, TITLES
End Sub

Private Sub LoadTribes(FromDatabase As Boolean)
    Dim i As Long
    Dim MyTribe As HRCORE.Tribe
    Dim ItemX As ListItem
    
    Me.lvwTribes.ListItems.Clear
    If FromDatabase Then
        MyTribes.GetAllTribes
    End If
    
    For i = 1 To MyTribes.count
        Set MyTribe = MyTribes.Item(i)
        Set ItemX = Me.lvwTribes.ListItems.add(, , MyTribe.Tribe)
        ItemX.Tag = MyTribe.TribeID
    Next i
    
End Sub

Private Sub Form_Resize()
    oSmart.FResize Me
End Sub


Private Sub lvwTribes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwTribes
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvwTribes_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set selTribe = Nothing
    Me.txtTribe.Text = ""
    If IsNumeric(Item.Tag) Then
        Set selTribe = MyTribes.FindTribe(CLng(Item.Tag))
        If Not (selTribe Is Nothing) Then
            Me.txtTribe.Text = selTribe.Tribe
        End If
    End If
End Sub

Private Function InsertNew() As Boolean
    Dim NewTribe As HRCORE.Tribe
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    Set NewTribe = New HRCORE.Tribe
    If Trim(txtTribe.Text) <> "" Then
        NewTribe.Tribe = Trim(txtTribe.Text)
        If Not (MyTribes.FindTribeByName(NewTribe.Tribe) Is Nothing) Then
            MsgBox "Another Tribe already exists with the supplied name", vbInformation, TITLES
            txtTribe.SetFocus
            Exit Function
        Else
            retVal = NewTribe.InsertNew()
        End If
    Else
        MsgBox "Enter the name of the Tribe", vbInformation, TITLES
        Me.txtTribe.SetFocus
        Exit Function
    End If
    
    InsertNew = True
    
    Exit Function
    
ErrorHandler:
    InsertNew = False
End Function


Private Function Update() As Boolean
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    
    If Trim(txtTribe.Text) <> "" Then
        selTribe.Tribe = Trim(txtTribe.Text)
        If Not (MyTribes.FindTribeByNameExclusive(selTribe.Tribe, selTribe.TribeID) Is Nothing) Then
            MsgBox "Another Tribe already exists with the supplied name", vbInformation, TITLES
            txtTribe.SetFocus
            Exit Function
        Else
            retVal = selTribe.Update()
        End If
    Else
        MsgBox "Enter the name of the Tribe", vbInformation, TITLES
        Me.txtTribe.SetFocus
        Exit Function
    End If
    
    Update = True
    
    Exit Function
    
ErrorHandler:
    Update = False
End Function

