VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNationalities 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nationalities"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   Icon            =   "frmNationality.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   6210
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
      Left            =   1680
      TabIndex        =   9
      Top             =   5880
      Visible         =   0   'False
      Width           =   855
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
      Left            =   5295
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
      Left            =   3480
      TabIndex        =   7
      Top             =   5880
      Visible         =   0   'False
      Width           =   855
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
      Left            =   2520
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
      Left            =   840
      TabIndex        =   5
      Top             =   5880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame fraExisting 
      Caption         =   "Existing Nationalities"
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
      Begin MSComctlLib.ListView lvwNationalities 
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
      Begin VB.TextBox txtNationality 
         Height          =   375
         Left            =   990
         TabIndex        =   2
         Top             =   360
         Width           =   4875
      End
      Begin VB.Label Label1 
         Caption         =   "Nationality"
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
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmNationalities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private nats As HRCORE.Nationalities
Private selNat As HRCORE.Nationality
Private blnEditMode As Boolean

Public Sub cmdCancel_Click()
    If blnEditMode = True Then
        cmdNew.Enabled = True
        txtNationality.Locked = True
        fraExisting.Enabled = True
        LoadNationalities False
    Else
        cmdDelete.Enabled = True
        txtNationality.Locked = True
        fraExisting.Enabled = True
        LoadNationalities False
    End If
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Public Sub cmdDelete_Click()
    Dim retVal As Long
    On Error GoTo ErrorHandler
    If Not currUser Is Nothing Then
        If currUser.CheckRight("Nationality") <> secModify Then
            MsgBox "You don't have right to delete record. Please liaise with the security admin"
            Exit Sub
        End If
     End If
    If Not (selNat Is Nothing) Then
        If MsgBox("Are you sure you want to delete: " & UCase(selNat.Nationality), vbYesNo + vbInformation, TITLES) = vbYes Then
            retVal = selNat.Delete()
            LoadNationalities True
        End If
    Else
        MsgBox "Select the Nationality to Delete", vbInformation, TITLES
    End If
           
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while deleting the Nationality" & vbNewLine & err.Description, vbInformation, TITLES
End Sub

Public Sub cmdEdit_Click()
    
    If Not currUser Is Nothing Then
        If currUser.CheckRight("Nationality") <> secModify Then
            MsgBox "You don't have right to modify record. Please liaise with the security admin"
            Exit Sub
        End If
     End If
    
    If selNat Is Nothing Then
        MsgBox "Select the Nationality to Edit", vbInformation, TITLES
        Exit Sub
    End If
    txtNationality.Locked = False
    fraExisting.Enabled = False
    cmdNew.Enabled = False
    txtNationality.SetFocus
    blnEditMode = True
        
  
End Sub

Public Sub cmdNew_Click()
    
    If Not currUser Is Nothing Then
        If currUser.CheckRight("Nationality") <> secModify Then
            MsgBox "You dont have right to add new record. Please liaise with the security admin"
            Exit Sub
        End If
     End If
     
    txtNationality.Locked = False
    fraExisting.Enabled = False
    cmdDelete.Enabled = False
    txtNationality.Text = ""
    txtNationality.SetFocus
    blnEditMode = False
    
End Sub

Public Sub cmdSave_Click()
    If blnEditMode = True Then
        If Not Update Then Exit Sub
        cmdNew.Enabled = True
        txtNationality.Locked = True
        fraExisting.Enabled = True
        LoadNationalities True
    Else
        If Not InsertNew Then Exit Sub
        cmdDelete.Enabled = True
        txtNationality.Locked = True
        fraExisting.Enabled = True
        LoadNationalities True
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    Set nats = New HRCORE.Nationalities
    frmMain2.PositionTheFormWithoutEmpList Me
    
    'initialize the listview
    Me.lvwNationalities.ColumnHeaders.Clear
    Me.lvwNationalities.ColumnHeaders.add , , "Nationality", Me.lvwNationalities.Width
    
    LoadNationalities True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An slight error has occurred" & vbNewLine & err.Description, vbInformation, TITLES
End Sub

Private Sub LoadNationalities(FromDatabase As Boolean)
    Dim i As Long
    Dim nat As HRCORE.Nationality
    Dim ItemX As ListItem
    
    Me.lvwNationalities.ListItems.Clear
    If FromDatabase Then
        nats.GetAllNationalities
    End If
    
    For i = 1 To nats.count
        Set nat = nats.Item(i)
        Set ItemX = Me.lvwNationalities.ListItems.add(, , nat.Nationality)
        ItemX.Tag = nat.NationalityID
    Next i
    
End Sub

Private Sub Form_Resize()
    oSmart.FResize Me
End Sub


Private Sub lvwNationalities_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwNationalities
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvwNationalities_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set selNat = Nothing
    Me.txtNationality.Text = ""
    If IsNumeric(Item.Tag) Then
        Set selNat = nats.FindNationality(CLng(Item.Tag))
        If Not (selNat Is Nothing) Then
            Me.txtNationality.Text = selNat.Nationality
        End If
    End If
End Sub

Private Function InsertNew() As Boolean
    Dim newNat As HRCORE.Nationality
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    Set newNat = New HRCORE.Nationality
    If Trim(txtNationality.Text) <> "" Then
        newNat.Nationality = Trim(txtNationality.Text)
        If Not (nats.FindNationalityByName(newNat.Nationality) Is Nothing) Then
            MsgBox "Another Nationality already exists with the supplied name", vbInformation, TITLES
            txtNationality.SetFocus
            Exit Function
        Else
            retVal = newNat.InsertNew()
        End If
    Else
        MsgBox "Enter the name of the Nationality", vbInformation, TITLES
        Me.txtNationality.SetFocus
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
    
    If Trim(txtNationality.Text) <> "" Then
        selNat.Nationality = Trim(txtNationality.Text)
        If Not (nats.FindNationalityByNameExclusive(selNat.Nationality, selNat.NationalityID) Is Nothing) Then
            MsgBox "Another nationality already exists with the supplied name", vbInformation, TITLES
            txtNationality.SetFocus
            Exit Function
        Else
            retVal = selNat.Update()
        End If
    Else
        MsgBox "Enter the name of the Nationality", vbInformation, TITLES
        Me.txtNationality.SetFocus
        Exit Function
    End If
    
    Update = True
    
    Exit Function
    
ErrorHandler:
    Update = False
End Function

