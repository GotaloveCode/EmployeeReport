VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmJDCategories 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JD Parameters"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab sstJD 
      Height          =   6165
      Left            =   75
      TabIndex        =   10
      Top             =   525
      Width           =   8190
      _ExtentX        =   14446
      _ExtentY        =   10874
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "JD Fields"
      TabPicture(0)   =   "frmJDCategories.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdNew"
      Tab(0).Control(1)=   "cmdEdit"
      Tab(0).Control(2)=   "cmdDelete"
      Tab(0).Control(3)=   "cmdClose"
      Tab(0).Control(4)=   "fraExisting"
      Tab(0).Control(5)=   "fraJDField"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Generic JDs"
      TabPicture(1)   =   "frmJDCategories.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraJDDetails"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdBack"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back To JD Fields"
         Height          =   465
         Left            =   6300
         TabIndex        =   9
         Top             =   5475
         Width           =   1740
      End
      Begin VB.Frame fraJDDetails 
         Caption         =   "Job Description"
         Height          =   4875
         Left            =   75
         TabIndex        =   15
         Top             =   465
         Width           =   7965
         Begin VB.CommandButton cmdNewJPJD 
            Caption         =   "Add"
            Height          =   345
            Left            =   4680
            TabIndex        =   23
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton cmdEditJPJD 
            Caption         =   "Edit"
            Height          =   345
            Left            =   5820
            TabIndex        =   22
            Top             =   1680
            Width           =   1005
         End
         Begin VB.CommandButton cmdDeleteJPJD 
            Caption         =   "Remove"
            Height          =   345
            Left            =   6915
            TabIndex        =   21
            Top             =   1680
            Width           =   930
         End
         Begin VB.ComboBox cboDesignation 
            Height          =   315
            Left            =   3735
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   240
            Width           =   4065
         End
         Begin VB.TextBox txtJDValue 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   2625
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   1005
            Width           =   5265
         End
         Begin MSComctlLib.ListView lvwJDValues 
            Height          =   2715
            Left            =   2625
            TabIndex        =   16
            Top             =   2085
            Width           =   5265
            _ExtentX        =   9287
            _ExtentY        =   4789
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "No."
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "JD Value"
               Object.Width           =   14111
            EndProperty
         End
         Begin MSComctlLib.TreeView tvwJDDetails 
            Height          =   4335
            Left            =   150
            TabIndex        =   8
            Top             =   345
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   7646
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   441
            LabelEdit       =   1
            Style           =   7
            Appearance      =   1
         End
         Begin VB.Label Label3 
            Caption         =   "Designation:"
            Height          =   240
            Left            =   2760
            TabIndex        =   20
            Top             =   270
            Width           =   1110
         End
         Begin VB.Label Label5 
            Caption         =   "JD Field Value:"
            Height          =   165
            Left            =   2700
            TabIndex        =   18
            Top             =   705
            Width           =   2565
         End
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   420
         Left            =   -74925
         TabIndex        =   4
         Top             =   5475
         Width           =   1140
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   420
         Left            =   -73635
         TabIndex        =   5
         Top             =   5475
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   420
         Left            =   -72195
         TabIndex        =   6
         Top             =   5475
         Width           =   1335
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   420
         Left            =   -68325
         TabIndex        =   7
         Top             =   5475
         Width           =   1335
      End
      Begin VB.Frame fraExisting 
         Caption         =   "Existing JD Fields:"
         Height          =   3615
         Left            =   -74850
         TabIndex        =   14
         Top             =   1725
         Width           =   7890
         Begin MSComctlLib.TreeView tvwJDFields 
            Height          =   3315
            Left            =   75
            TabIndex        =   3
            Top             =   225
            Width           =   7740
            _ExtentX        =   13653
            _ExtentY        =   5847
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   353
            LabelEdit       =   1
            Style           =   7
            Appearance      =   1
         End
      End
      Begin VB.Frame fraJDField 
         Height          =   1065
         Left            =   -74850
         TabIndex        =   11
         Top             =   525
         Width           =   7890
         Begin VB.TextBox txtParentJDCategory 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   600
            Width           =   6540
         End
         Begin VB.TextBox txtJDCategoryName 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            TabIndex        =   1
            Top             =   225
            Width           =   6540
         End
         Begin VB.Label Label2 
            Caption         =   "Sub-Field Of:"
            Height          =   165
            Left            =   150
            TabIndex        =   13
            Top             =   600
            Width           =   990
         End
         Begin VB.Label Label1 
            Caption         =   "JD Field:"
            Height          =   165
            Left            =   150
            TabIndex        =   12
            Top             =   300
            Width           =   1065
         End
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Job Description Setup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   15
      Width           =   3135
   End
End
Attribute VB_Name = "frmJDCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private selJDField As HRCORE.JDCategory

Private pJDs As HRCORE.JDCategories
Private TopLevelJDFields As HRCORE.JDCategories

Private pJobPos As HRCORE.JobPositions
Private selJobPos As HRCORE.JobPosition

Private pJobPosJDs As HRCORE.JobPositionJDs
Private FilteredJobPosJDs As HRCORE.JobPositionJDs 'of a Particular JDField
Private selJobPosJDField As HRCORE.JDCategory
Private SelJobPosJDValue As HRCORE.JobPositionJD



Private Sub cboDesignation_Click()
    On Error GoTo ErrorHandler
    
    'clear the listview
    Me.lvwJDValues.ListItems.Clear
    Me.txtJDValue.Text = ""
    
    Set selJobPos = Nothing
    If cboDesignation.ListIndex > -1 Then
        Set selJobPos = pJobPos.FindJobPositionByID(cboDesignation.ItemData(cboDesignation.ListIndex))
    End If
    
    If Not (selJobPos Is Nothing) Then
        'get the JDValues for the selected job position
        Set selJobPos.JobPositionJDValues = pJobPosJDs.GetJobPositionJDsOfPosition(selJobPos.PositionID)
        
        'load the JD Values in the first parse
        LoadFilteredJobPositionJDValues selJobPos.JobPositionJDValues
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while processing the Generic JD for the selected Job Position" & vbNewLine & err.Description, vbExclamation, TITLES
    
    
End Sub

Private Sub cmdBack_Click()
    Me.sstJD.Tab = 0
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim retVal As Long
    Dim resp As Long
    
    On Error GoTo ErrorHandler
    Select Case LCase(cmdDelete.Caption)
        Case "delete"
            If selJDField Is Nothing Then
                MsgBox "Select the JD Field that you want to Delete", vbInformation, TITLES
            Else
                resp = MsgBox("Are you sure you want to delete the selected JD Field: " & vbNewLine & selJDField.CategoryName, vbQuestion + vbYesNo, TITLES)
                If resp = vbYes Then
                    retVal = selJDField.Delete()
                    Set selJDField = Nothing
                    LoadJDFields
                End If
            End If
            
        Case "cancel"
            cmdNew.Enabled = True
            cmdEdit.Caption = "Edit"
            cmdDelete.Caption = "Delete"
            fraExisting.Enabled = True
            fraJDField.Enabled = False
    End Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while performing the requested operation" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub cmdDeleteJPJD_Click()
    Dim retVal As Long
    Dim resp As Long
    
    On Error GoTo ErrorHandler
    
    Select Case LCase(cmdDeleteJPJD.Caption)
        Case "remove"
            If SelJobPosJDValue Is Nothing Then
                MsgBox "Select the Generic JD Value, from the List of Values, that you want to Delete", vbInformation, TITLES
                Exit Sub
            Else
                resp = MsgBox("Are you sure you want to delete the selected Generic JD Value?", vbYesNo + vbQuestion, TITLES)
                If resp = vbYes Then
                    retVal = SelJobPosJDValue.Delete()
                    Set SelJobPosJDValue = Nothing
                    LoadJDValuesOfJobPosition selJobPos, True
                End If
            End If
            
        Case "cancel"
            cmdNewJPJD.Enabled = True
            cmdEditJPJD.Caption = "Edit"
            cmdDeleteJPJD.Caption = "Remove"
            Me.txtJDValue.Locked = True
            Me.tvwJDDetails.Enabled = True
            Me.lvwJDValues.Enabled = True
            LoadJDValuesOfJobPosition selJobPos, False
    End Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while performing the requested operation" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub cmdEdit_Click()
    Select Case LCase(cmdEdit.Caption)
        Case "edit"
            cmdNew.Enabled = False
            cmdEdit.Caption = "Update"
            cmdDelete.Caption = "Cancel"
            fraExisting.Enabled = False
            fraJDField.Enabled = True
            
        Case "update"
            If Update() = False Then Exit Sub
            cmdNew.Enabled = True
            cmdEdit.Caption = "Edit"
            cmdDelete.Caption = "Delete"
            fraExisting.Enabled = True
            fraJDField.Enabled = False
            
        Case "cancel"
            cmdNew.Caption = "New"
            cmdEdit.Caption = "Edit"
            cmdDelete.Enabled = True
            fraExisting.Enabled = True
            fraJDField.Enabled = False
            
    End Select
End Sub

Private Sub cmdEditJPJD_Click()
    Select Case LCase(cmdEditJPJD.Caption)
        Case "edit"
            If SelJobPosJDValue Is Nothing Then
                MsgBox "You have to Select the Generic JD Value to Edit", vbInformation, TITLES
                Me.lvwJDValues.Enabled = True
                Me.lvwJDValues.SetFocus
                Exit Sub
            End If
            
            cmdNewJPJD.Enabled = False
            cmdEditJPJD.Caption = "Update"
            cmdDeleteJPJD.Caption = "Cancel"
            Me.txtJDValue.Locked = False
            Me.tvwJDDetails.Enabled = False
            Me.lvwJDValues.Enabled = False
            Me.txtJDValue.SetFocus
            
        Case "update"
            If UpdateJDValue() = False Then Exit Sub
            cmdNewJPJD.Enabled = True
            cmdEditJPJD.Caption = "Edit"
            cmdDeleteJPJD.Caption = "Remove"
            Me.txtJDValue.Locked = True
            Me.tvwJDDetails.Enabled = True
            Me.lvwJDValues.Enabled = True
            LoadJDValuesOfJobPosition selJobPos, True
            
        Case "cancel"
            cmdNewJPJD.Caption = "Add"
            cmdEditJPJD.Caption = "Edit"
            cmdDeleteJPJD.Enabled = True
            Me.txtJDValue.Locked = True
            Me.tvwJDDetails.Enabled = True
            Me.lvwJDValues.Enabled = True
            Me.cboDesignation.Locked = False
            LoadJDValuesOfJobPosition selJobPos, False
    End Select
End Sub

Private Sub cmdNew_Click()
    Select Case LCase(cmdNew.Caption)
        Case "new"
            cmdNew.Caption = "Update"
            cmdEdit.Caption = "Cancel"
            cmdDelete.Enabled = False
            fraExisting.Enabled = False
            fraJDField.Enabled = True
            Me.txtJDCategoryName.Text = ""
            Me.txtJDCategoryName.SetFocus
            
        Case "update"
            If InsertNew() = False Then Exit Sub
            cmdNew.Caption = "New"
            cmdEdit.Caption = "Edit"
            cmdDelete.Enabled = True
            fraExisting.Enabled = True
            fraJDField.Enabled = False
            
            'reload JD Fields
            LoadJDFields
    End Select
End Sub

Private Sub cmdNewJPJD_Click()
    Select Case LCase(cmdNewJPJD.Caption)
        Case "add"
            If selJobPos Is Nothing Then
                MsgBox "You have to Select the Designation first", vbInformation, TITLES
                cboDesignation.Locked = False
                cboDesignation.SetFocus
                Exit Sub
            End If
            
            If selJobPosJDField Is Nothing Then
                MsgBox "You have to select the JD Field to add Values to", vbInformation, TITLES
                Me.tvwJDDetails.Enabled = True
                Me.tvwJDDetails.SetFocus
                Exit Sub
            End If
            cmdNewJPJD.Caption = "Update"
            cmdEditJPJD.Caption = "Cancel"
            cmdDeleteJPJD.Enabled = False
            Me.txtJDValue.Text = ""
            Me.txtJDValue.Locked = False
            Me.tvwJDDetails.Enabled = False
            Me.lvwJDValues.Enabled = False
            Me.cboDesignation.Locked = True
            Me.txtJDValue.SetFocus
            
        Case "update"
            If AddJobPositionJDValue() = False Then Exit Sub
            cmdNewJPJD.Caption = "Add"
            cmdEditJPJD.Caption = "Edit"
            cmdDeleteJPJD.Enabled = True
            Me.txtJDValue.Locked = True
            Me.tvwJDDetails.Enabled = True
            Me.lvwJDValues.Enabled = True
            Me.cboDesignation.Locked = False
            LoadJDValuesOfJobPosition selJobPos, True
    End Select
            
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    frmMain2.PositionTheFormWithoutEmpList Me
    
    Set pJDs = New HRCORE.JDCategories
    Set pJobPos = New HRCORE.JobPositions
    Set pJobPosJDs = New HRCORE.JobPositionJDs
    
    'populate all the Job Position JD Values
    pJobPosJDs.GetActiveJobPositionJDs
    
    'populate the job positions
    LoadJobPositions
    
    'populate the JD Categories
    LoadJDFields
    
    Exit Sub
    
ErrorHandler:
    
End Sub

Private Sub LoadJobPositions()
    Dim i As Long
    On Error GoTo ErrorHandler
    
    cboDesignation.Clear
    
    pJobPos.GetAllJobPositions
    
    For i = 1 To pJobPos.count
        Me.cboDesignation.AddItem pJobPos.Item(i).PositionName
        Me.cboDesignation.ItemData(cboDesignation.NewIndex) = pJobPos.Item(i).PositionID
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while Populating the Job Positions" & vbNewLine & err.Description, vbExclamation, TITLES
    
End Sub

Private Sub LoadJDFields()
    Dim myJD As HRCORE.JDCategory
    Dim myNode As Node
    Dim myNode2 As Node
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
       
    'clear lists
    Me.tvwJDFields.Nodes.Clear
    'handles the other treeview
    Me.tvwJDDetails.Nodes.Clear
        
    pJDs.GetActiveJDCategories
    
    'first load the Header
    Set myNode = tvwJDFields.Nodes.add(, , "JDFIELDS", "JD Fields")
    myNode.Tag = "JDFIELDS"
    myNode.Bold = True
    
'    Set myNode2 = tvwJDDetails.Nodes.Add(, , "JDFIELDS", "JD Fields")
'    myNode2.Tag = "JDFIELDS"
'    myNode2.Bold = True
    
    'now get the Top Level JD Fields
    Set TopLevelJDFields = pJDs.GetTopLevelJDCategories()
    If Not (TopLevelJDFields Is Nothing) Then
        For i = 1 To TopLevelJDFields.count
            Set myJD = TopLevelJDFields.Item(i)
            
            'set numbering
            myJD.FieldNumber = i
                      
            'add the JD
            Set myNode = Me.tvwJDFields.Nodes.add(, , "JD:" & myJD.JDCategoryID, myJD.FieldNumber & ". " & myJD.CategoryName)
            myNode.Tag = myJD.JDCategoryID
            myNode.EnsureVisible
            
            'the second treeview
            Set myNode2 = Me.tvwJDDetails.Nodes.add(, , "JD:" & myJD.JDCategoryID, myJD.FieldNumber & ". " & myJD.CategoryName)
            myNode2.Tag = myJD.JDCategoryID
            myNode2.Bold = True
            myNode2.EnsureVisible
            
            'now recursively add the children
            AddChildJDFieldsRecursively myJD
        Next i
         
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "An error has occurred while populating the JD Fields" & _
    vbNewLine & err.Description, vbInformation, TITLES
End Sub


Private Sub AddChildJDFieldsRecursively(ByVal TheJD As HRCORE.JDCategory)
    
    'this is a recursive function that populates child JD Fields
    Dim ChildNode As Node
    Dim ChildNode2 As Node
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    If Not (TheJD Is Nothing) Then
        For i = 1 To TheJD.Children.count
            
            'set the field numbering
            TheJD.Children.Item(i).FieldNumber = TheJD.FieldNumber & "." & i
            
            Set ChildNode = Me.tvwJDFields.Nodes.add("JD:" & TheJD.JDCategoryID, tvwChild, "JD:" & TheJD.Children.Item(i).JDCategoryID, TheJD.Children.Item(i).FieldNumber & ". " & TheJD.Children.Item(i).CategoryName)
            ChildNode.Tag = TheJD.Children.Item(i).JDCategoryID
            ChildNode.EnsureVisible
            
            'Set the Second Treeview
            Set ChildNode2 = Me.tvwJDDetails.Nodes.add("JD:" & TheJD.JDCategoryID, tvwChild, "JD:" & TheJD.Children.Item(i).JDCategoryID, TheJD.Children.Item(i).FieldNumber & ". " & TheJD.Children.Item(i).CategoryName)
            ChildNode2.Tag = TheJD.Children.Item(i).JDCategoryID
            ChildNode2.Bold = True
            ChildNode2.EnsureVisible
            
            'recursively load the children
            AddChildJDFieldsRecursively TheJD.Children.Item(i)
        Next i
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred" & vbNewLine & err.Description, vbInformation, TITLES
        
End Sub



Private Sub lvwJDValues_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo ErrorHandler
    
    Set SelJobPosJDValue = Nothing
    If IsNumeric(Item.Tag) Then
        Set SelJobPosJDValue = pJobPosJDs.FindJobPositionJDByID(CLng(Item.Tag))
    End If
    
    If Not (SelJobPosJDValue Is Nothing) Then
        Me.txtJDValue.Text = SelJobPosJDValue.FieldValue
    Else
        Me.txtJDValue.Text = ""
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while processing the Selected JD Value" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub tvwJDDetails_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error GoTo ErrorHandler
    
    Set selJobPosJDField = Nothing
    'clear the listview
    Me.lvwJDValues.ListItems.Clear
    Me.txtJDValue.Text = ""
    
    If IsNumeric(Node.Tag) Then
        Set selJobPosJDField = pJDs.FindJDCategoryByID(CLng(Node.Tag))
    End If
    
    If selJobPosJDField Is Nothing Then
        Exit Sub
    End If
    
    'Check the selected position
    
    If Not (selJobPos Is Nothing) Then
        'refresh the JDValues of the selected position
        'apply first filter
        Set selJobPos.JobPositionJDValues = pJobPosJDs.GetJobPositionJDsOfPosition(selJobPos.PositionID)
        If Not (selJobPos.JobPositionJDValues Is Nothing) Then
            
            'apply second filter
            Set FilteredJobPosJDs = selJobPos.JobPositionJDValues.GetJobPositionJDsOfJDCategory(selJobPosJDField.JDCategoryID)
            
            'now populate the listview
            LoadFilteredJobPositionJDValues FilteredJobPosJDs
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while populating the Generic JD" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub tvwJDFields_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error GoTo ErrorHandler
    
    Set selJDField = Nothing
    If UCase(Node.Tag) = "JDFIELDS" Then Exit Sub
    If IsNumeric(Node.Tag) Then
        Set selJDField = pJDs.FindJDCategoryByID(CLng(Node.Tag))
    End If
    
    SetFields selJDField
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while processing the selected JD Field" & vbNewLine & err.Description, vbExclamation, TITLES
    
End Sub

Private Sub SetFields(ByVal TheJDField As HRCORE.JDCategory)
    On Error GoTo ErrorHandler
    
    If Not (TheJDField Is Nothing) Then
        Me.txtJDCategoryName.Text = TheJDField.CategoryName
        If Not (TheJDField.ParentJDCategory Is Nothing) Then
            Me.txtParentJDCategory.Text = TheJDField.ParentJDCategory.CategoryName
        Else
            Me.txtParentJDCategory.Text = ""
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while processing the selected JD Field" & vbNewLine & err.Description, vbExclamation, TITLES
    
End Sub

Private Function InsertNew() As Boolean
    Dim newJDField As HRCORE.JDCategory
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    
    Set newJDField = New HRCORE.JDCategory
    
    If Trim(Me.txtJDCategoryName.Text) <> "" Then
        newJDField.CategoryName = Trim(Me.txtJDCategoryName.Text)
    Else
        MsgBox "The Name of the JD Field is Required", vbExclamation, TITLES
        Me.txtJDCategoryName.SetFocus
        Exit Function
    End If
    
    If selJDField Is Nothing Then
        Set newJDField.ParentJDCategory = Nothing
    Else
        Set newJDField.ParentJDCategory = selJDField
    End If
    
    retVal = newJDField.InsertNew()
    If retVal <> 0 Then
        MsgBox "The JD Field was not added", vbExclamation, TITLES
        InsertNew = False
    Else
        InsertNew = True
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred while adding the new JD Field" & vbNewLine & err.Description, vbExclamation, TITLES
    InsertNew = False
End Function


Private Function Update() As Boolean
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
      
    If Trim(Me.txtJDCategoryName.Text) <> "" Then
        selJDField.CategoryName = Trim(Me.txtJDCategoryName.Text)
    Else
        MsgBox "The Name of the JD Field is Required", vbExclamation, TITLES
        Me.txtJDCategoryName.SetFocus
        Exit Function
    End If
    'the parent will not be updated
    
    retVal = selJDField.Update()
    
    If retVal <> 0 Then
        MsgBox "The JD Field was not Updated", vbExclamation, TITLES
        Update = False
    Else
        Update = True
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred while Updating the JD Field" & vbNewLine & err.Description, vbExclamation, TITLES
    Update = False
End Function


Private Sub LoadJDValuesOfJobPosition(ByVal TheJobPosition As HRCORE.JobPosition, ByVal Refresh As Boolean)
    Dim i As Long
    Dim ItemX As ListItem
    
    On Error GoTo ErrorHandler
    Me.lvwJDValues.ListItems.Clear
        
    If TheJobPosition Is Nothing Then Exit Sub
    
    If Refresh = True Then
        pJobPosJDs.GetActiveJobPositionJDs
    End If
    
    Set TheJobPosition.JobPositionJDValues = pJobPosJDs.GetJobPositionJDsOfPosition(TheJobPosition.PositionID)
    
    'load all the Values of the Selected Position
    LoadFilteredJobPositionJDValues TheJobPosition.JobPositionJDValues
    
    Exit Sub
    
ErrorHandler:
    
End Sub


Private Sub LoadFilteredJobPositionJDValues(ByVal TheFilteredJobPosJDs As HRCORE.JobPositionJDs)
    Dim i As Long
    Dim ItemX As ListItem
    
    On Error GoTo ErrorHandler
    
    Me.lvwJDValues.ListItems.Clear
    
    If TheFilteredJobPosJDs Is Nothing Then Exit Sub
    
    For i = 1 To TheFilteredJobPosJDs.count
        TheFilteredJobPosJDs.Item(i).FieldNumber = selJobPosJDField.FieldNumber & "." & i
        Set ItemX = Me.lvwJDValues.ListItems.add(, , TheFilteredJobPosJDs.Item(i).FieldNumber)
        ItemX.SubItems(1) = TheFilteredJobPosJDs.Item(i).FieldValue
        ItemX.Tag = TheFilteredJobPosJDs.Item(i).JobPositionJDID
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while populating the Generic JD for the Selected Position" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub


Private Function AddJobPositionJDValue() As Boolean
    Dim newJPJDValue As HRCORE.JobPositionJD
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    Set newJPJDValue = New HRCORE.JobPositionJD
    
    If Trim(Me.txtJDValue.Text) <> "" Then
        newJPJDValue.FieldValue = Trim(Me.txtJDValue.Text)
    Else
        MsgBox "The value for the JD Field is Required", vbExclamation, TITLES
        Me.txtJDValue.SetFocus
        Exit Function
    End If
    
    Set newJPJDValue.JDCategory = selJobPosJDField
    Set newJPJDValue.position = selJobPos
    retVal = newJPJDValue.InsertNew()
    If retVal <> 0 Then
        MsgBox "The new Value was not added", vbInformation, TITLES
        AddJobPositionJDValue = False
    Else
        AddJobPositionJDValue = True
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error has occurred while adding the new Value" & vbNewLine & err.Description, vbExclamation, tiles
    AddJobPositionJDValue = False
End Function

Private Function UpdateJDValue() As Boolean
    
    On Error GoTo ErrorHandler
       
    If Trim(Me.txtJDValue.Text) <> "" Then
        SelJobPosJDValue.FieldValue = Trim(Me.txtJDValue.Text)
    Else
        MsgBox "The value for the JD Field is Required", vbExclamation, TITLES
        Me.txtJDValue.SetFocus
        Exit Function
    End If
    
    Set SelJobPosJDValue.JDCategory = selJobPosJDField
    Set SelJobPosJDValue.position = selJobPos
    retVal = SelJobPosJDValue.Update()
    If retVal <> 0 Then
        MsgBox "The JD Value was not Updated", vbInformation, TITLES
        UpdateJDValue = False
    Else
        UpdateJDValue = True
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error has occurred while Updating the JD Value" & vbNewLine & err.Description, vbExclamation, tiles
    UpdateJDValue = False
End Function
