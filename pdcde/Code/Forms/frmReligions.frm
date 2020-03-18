VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReligions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Religions"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   Icon            =   "frmReligions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   6240
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
      Left            =   3300
      TabIndex        =   10
      Top             =   6000
      Visible         =   0   'False
      Width           =   855
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
      Left            =   1380
      TabIndex        =   9
      Top             =   6000
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
      Left            =   5220
      TabIndex        =   8
      Top             =   6000
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
      Left            =   4260
      TabIndex        =   7
      Top             =   6000
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
      Left            =   2340
      TabIndex        =   6
      Top             =   6000
      Visible         =   0   'False
      Width           =   855
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
      Top             =   6000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame fraExisting 
      Caption         =   "Existing Religions"
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
      Begin MSComctlLib.ListView lvwReligions 
         Height          =   4515
         Left            =   90
         TabIndex        =   4
         Top             =   240
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
      Begin VB.TextBox txtReligion 
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
         Caption         =   "Religion"
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
Attribute VB_Name = "frmReligions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Religns As HRCORE.Religions
Private selRelign As HRCORE.Religion
Private blnEditMode As Boolean

Public Sub cmdCancel_Click()
    If blnEditMode = True Then
        cmdNew.Enabled = True
        txtReligion.Locked = True
        fraExisting.Enabled = True
        LoadReligions False
    Else
        cmdDelete.Enabled = True
        txtReligion.Locked = True
        fraExisting.Enabled = True
        LoadReligions False
    End If
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Public Sub cmdDelete_Click()
    Dim retVal As Long
    On Error GoTo ErrorHandler
    
    If Not (selRelign Is Nothing) Then
        If MsgBox("Are you sure you want to delete: " & UCase(selRelign.Religion), vbYesNo + vbInformation, TITLES) = vbYes Then
            retVal = selRelign.Delete()
            LoadReligions True
        End If
    Else
        MsgBox "Select the Religion to Delete", vbInformation, TITLES
    End If
           
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while deleting the Religion" & vbNewLine & err.Description, vbInformation, TITLES
End Sub

Public Sub cmdEdit_Click()
    
    If selRelign Is Nothing Then
        MsgBox "Select the Religion to Edit", vbInformation, TITLES
        Exit Sub
    End If
    txtReligion.Locked = False
    fraExisting.Enabled = False
    cmdNew.Enabled = False
    txtReligion.SetFocus
    blnEditMode = True
        
  
End Sub

Public Sub cmdNew_Click()
    txtReligion.Locked = False
    fraExisting.Enabled = False
    cmdDelete.Enabled = False
    txtReligion.Text = ""
    txtReligion.SetFocus
    blnEditMode = False
    
End Sub

Public Sub cmdSave_Click()
    If blnEditMode = True Then
        If Not Update Then Exit Sub
        cmdNew.Enabled = True
        txtReligion.Locked = True
        fraExisting.Enabled = True
        LoadReligions True
    Else
        If Not InsertNew Then Exit Sub
        cmdDelete.Enabled = True
        txtReligion.Locked = True
        fraExisting.Enabled = True
        LoadReligions True
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    Set Religns = New HRCORE.Religions
    frmMain2.PositionTheFormWithoutEmpList Me
    
    'initialize the listview
    Me.lvwReligions.ColumnHeaders.Clear
    Me.lvwReligions.ColumnHeaders.add , , "Religion", Me.lvwReligions.Width
    
    LoadReligions True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An slight error has occurred" & vbNewLine & err.Description, vbInformation, TITLES
End Sub

Private Sub LoadReligions(FromDatabase As Boolean)
    Dim i As Long
    Dim relign As HRCORE.Religion
    Dim ItemX As ListItem
    
    Me.lvwReligions.ListItems.Clear
    If FromDatabase Then
        Religns.GetAllReligions
    End If
    
    For i = 1 To Religns.count
        Set relign = Religns.Item(i)
        Set ItemX = Me.lvwReligions.ListItems.add(, , relign.Religion)
        ItemX.Tag = relign.ReligionID
    Next i
    
End Sub

Private Sub Form_Resize()
    oSmart.FResize Me
End Sub


Private Sub lvwReligions_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwReligions
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvwReligions_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set selRelign = Nothing
    Me.txtReligion.Text = ""
    If IsNumeric(Item.Tag) Then
        Set selRelign = Religns.FindReligion(CLng(Item.Tag))
        If Not (selRelign Is Nothing) Then
            Me.txtReligion.Text = selRelign.Religion
        End If
    End If
End Sub

Private Function InsertNew() As Boolean
    Dim NewRelign As HRCORE.Religion
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    Set NewRelign = New HRCORE.Religion
    If Trim(txtReligion.Text) <> "" Then
        NewRelign.Religion = Trim(txtReligion.Text)
        If Not (Religns.FindReligionByName(NewRelign.Religion) Is Nothing) Then
            MsgBox "Another Religion already exists with the supplied name", vbInformation, TITLES
            txtReligion.SetFocus
            Exit Function
        Else
            retVal = NewRelign.InsertNew()
        End If
    Else
        MsgBox "Enter the name of the Religion", vbInformation, TITLES
        Me.txtReligion.SetFocus
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
    
    If Trim(txtReligion.Text) <> "" Then
        selRelign.Religion = Trim(txtReligion.Text)
        If Not (Religns.FindReligionByNameExclusive(selRelign.Religion, selRelign.ReligionID) Is Nothing) Then
            MsgBox "Another Religion already exists with the supplied name", vbInformation, TITLES
            txtReligion.SetFocus
            Exit Function
        Else
            retVal = selRelign.Update()
        End If
    Else
        MsgBox "Enter the name of the Religion", vbInformation, TITLES
        Me.txtReligion.SetFocus
        Exit Function
    End If
    
    Update = True
    
    Exit Function
    
ErrorHandler:
    Update = False
End Function

