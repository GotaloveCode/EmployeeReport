VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSSCategories 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Staff Categories"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraExisting 
      Caption         =   "Existing Staff Categories / Staff Levels:"
      Height          =   3540
      Left            =   150
      TabIndex        =   10
      Top             =   2130
      Width           =   7290
      Begin MSComctlLib.ListView lvwCategories 
         Height          =   3210
         Left            =   150
         TabIndex        =   2
         Top             =   225
         Width           =   7110
         _ExtentX        =   12541
         _ExtentY        =   5662
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Staff Category"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Details"
            Object.Width           =   8819
         EndProperty
      End
   End
   Begin VB.Frame fraCSS 
      Height          =   1470
      Left            =   120
      TabIndex        =   6
      Top             =   570
      Width           =   7320
      Begin VB.TextBox txtCategoryName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2790
         TabIndex        =   0
         Top             =   240
         Width           =   4320
      End
      Begin VB.TextBox txtDetails 
         Appearance      =   0  'Flat
         Height          =   585
         Left            =   2790
         TabIndex        =   1
         Top             =   675
         Width           =   4350
      End
      Begin VB.Label Label1 
         Caption         =   "Category Name / Level Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Details:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   690
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6150
      TabIndex        =   7
      Top             =   5910
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   5910
      Width           =   1335
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   405
      Left            =   1680
      TabIndex        =   4
      Top             =   5910
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   5910
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Staff Categories / Staff Levels"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   11
      Top             =   90
      Width           =   4035
   End
End
Attribute VB_Name = "frmCSSCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private selCSSSCat As HRCORE.CSSSCategory
Private pCats As HRCORE.CSSSCategories

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim resp As Long
    Dim retVal As Long
    
    Select Case LCase(cmdDelete.Caption)
        Case "delete"
            If Not (selCSSSCat Is Nothing) Then
                resp = MsgBox("Are you sure you want to delete the selected Staff Category?", vbYesNo + vbQuestion, TITLES)
                If resp = vbYes Then
                    retVal = selCSSSCat.Delete()
                    LoadCSSCategories
                End If
            Else
                MsgBox "There is no selected staff category", vbInformation, TITLES
            End If
        Case "cancel"
            cmdNew.Enabled = True
            cmdEdit.Caption = "Edit"
            cmdDelete.Caption = "Delete"
            fraCSS.Enabled = True
    End Select
End Sub

Private Sub cmdEdit_Click()
    Select Case LCase(cmdEdit.Caption)
        Case "edit"
            cmdNew.Enabled = False
            cmdEdit.Caption = "Update"
            cmdDelete.Caption = "Cancel"
            fraCSS.Enabled = True
            
        Case "cancel"
            cmdNew.Caption = "New"
            cmdEdit.Caption = "Edit"
            cmdDelete.Enabled = True
            fraCSS.Enabled = False
            
        Case "update"
            If Update() = False Then Exit Sub
            cmdNew.Enabled = True
            cmdEdit.Caption = "Edit"
            cmdDelete.Caption = "Delete"
            fraCSS.Enabled = True
            LoadCSSCategories
    End Select
    
End Sub

Private Sub cmdNew_Click()
    Select Case LCase(cmdNew.Caption)
        Case "new"
            cmdNew.Caption = "Update"
            cmdEdit.Caption = "Cancel"
            cmdDelete.Enabled = False
            ClearControls
            fraCSS.Enabled = True
            
            
        Case "update"
            If InsertNew() = False Then Exit Sub
            LoadCSSCategories
            cmdNew.Caption = "New"
            cmdEdit.Caption = "Edit"
            cmdDelete.Enabled = True
            fraCSS.Enabled = False
            
            
    End Select
End Sub

Private Sub ClearControls()
    Me.txtCategoryName.Text = ""
    Me.txtDetails.Text = ""
End Sub


Private Function InsertNew() As Boolean
    Dim newCSS As HRCORE.CSSSCategory
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    
    Set newCSS = New HRCORE.CSSSCategory
    If Trim(Me.txtCategoryName.Text) <> "" Then
        newCSS.CSSSCategoryName = Trim(txtCategoryName.Text)
    Else
        MsgBox "Enter Category Name", vbExclamation, TITLES
        Me.txtCategoryName.SetFocus
        Exit Function
    End If
    
    newCSS.Details = Trim(txtDetails.Text)
    
    retVal = newCSS.InsertNew()
    If retVal = 0 Then
        MsgBox "The new Staff Category has been added successfully", vbInformation, TITLES
        InsertNew = True
    End If

    Exit Function
    
ErrorHandler:
    MsgBox "An error has occurred while creating a new Staff Category" & vbNewLine & err.Description, vbInformation, TITLES
    InsertNew = False
End Function


Private Function Update() As Boolean
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    
    If selCSSSCat Is Nothing Then Exit Function
    If Trim(Me.txtCategoryName.Text) <> "" Then
        selCSSSCat.CSSSCategoryName = Trim(txtCategoryName.Text)
    Else
        MsgBox "Enter Category Name", vbExclamation, TITLES
        Me.txtCategoryName.SetFocus
        Exit Function
    End If
    
    selCSSSCat.Details = Trim(txtDetails.Text)
    
    retVal = selCSSSCat.Update()
    If retVal = 0 Then
        MsgBox "The Staff Category has been updated successfully", vbInformation, TITLES
        Update = True
    End If

    Exit Function
    
ErrorHandler:
    MsgBox "An error has occurred while Updating the Staff Category" & vbNewLine & err.Description, vbInformation, TITLES
    Update = False
End Function

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    frmMain2.PositionTheFormWithoutEmpList Me
    
    Set pCats = New HRCORE.CSSSCategories
    
    
    LoadCSSCategories
    
    Exit Sub
    
ErrorHandler:
    
End Sub

Private Sub LoadCSSCategories()
    On Error GoTo ErrorHandler
    
    pCats.GetActiveCSSSCategories
    
    PopulateCategories pCats
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while Populating Existing Staff Categories" & vbNewLine & err.Description, vbExclamation, TITLES
    
    
End Sub

Private Sub PopulateCategories(ByVal TheCats As HRCORE.CSSSCategories)
    Dim i As Long
    Dim ItemX As ListItem
    
    On Error GoTo ErrorHandler
    lvwCategories.ListItems.Clear
    
    If Not (TheCats Is Nothing) Then
        For i = 1 To TheCats.count
            Set ItemX = lvwCategories.ListItems.add(, , TheCats.Item(i).CSSSCategoryName)
            ItemX.SubItems(1) = TheCats.Item(i).Details
            ItemX.Tag = TheCats.Item(i).CSSSCategoryID
        Next i
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while Populating the Staff Categories" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub SetFields(ByVal TheCSS As HRCORE.CSSSCategory)
    ClearControls
    If Not (TheCSS Is Nothing) Then
        Me.txtCategoryName.Text = TheCSS.CSSSCategoryName
        Me.txtDetails.Text = TheCSS.Details
    End If
End Sub

Private Sub lvwCategories_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim theID As Long
    
    Set selCSSSCat = Nothing
    If Not IsNumeric(Item.Tag) Then Exit Sub
    theID = CLng(Item.Tag)
    Set selCSSSCat = pCats.FindCSSSCategoryByID(theID)
    SetFields selCSSSCat
End Sub
