VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOUTypes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Organization Unit Types"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   7575
   Begin VB.Frame Frame1 
      Height          =   6045
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   7215
      Begin VB.Frame Frame2 
         Height          =   5415
         Left            =   0
         TabIndex        =   11
         Top             =   630
         Width           =   7215
         Begin VB.Frame Frame3 
            Caption         =   "OU Types"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2475
            Left            =   0
            TabIndex        =   19
            Top             =   1980
            Width           =   7185
            Begin MSComctlLib.TreeView tvwOUTypes 
               Height          =   2055
               Left            =   120
               TabIndex        =   6
               Top             =   240
               Width           =   6855
               _ExtentX        =   12091
               _ExtentY        =   3625
               _Version        =   393217
               HideSelection   =   0   'False
               LabelEdit       =   1
               Style           =   6
               FullRowSelect   =   -1  'True
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
            End
            Begin VB.Label lblInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "w"
               ForeColor       =   &H80000008&
               Height          =   225
               Left            =   360
               TabIndex        =   20
               Top             =   240
               Width           =   150
            End
         End
         Begin VB.Frame fraDetails 
            Caption         =   "Details"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   7155
            Begin VB.TextBox txtOUTCode 
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
               Height          =   285
               Left            =   1410
               TabIndex        =   1
               Top             =   270
               Width           =   1035
            End
            Begin VB.TextBox txtOUTName 
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
               Height          =   285
               Left            =   4140
               TabIndex        =   2
               Top             =   270
               Width           =   2745
            End
            Begin VB.TextBox txtOUTLevel 
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
               Height          =   285
               Left            =   1410
               TabIndex        =   3
               Top             =   735
               Width           =   1035
            End
            Begin VB.ComboBox cboParentOUTypes 
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
               Left            =   1440
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   1170
               Width           =   3195
            End
            Begin VB.TextBox txtTitleOfHeadEmp 
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
               Left            =   4140
               TabIndex        =   4
               Top             =   720
               Width           =   2745
            End
            Begin VB.Label Label1 
               Caption         =   "OU Type Code"
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
               TabIndex        =   18
               Top             =   285
               Width           =   1215
            End
            Begin VB.Label Label2 
               Caption         =   "OU Type Name"
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
               Left            =   2730
               TabIndex        =   17
               Top             =   285
               Width           =   1335
            End
            Begin VB.Label Label3 
               Caption         =   "OU Type Level"
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
               TabIndex        =   16
               Top             =   750
               Width           =   1185
            End
            Begin VB.Label Label4 
               Caption         =   "Parent OU Type"
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
               Left            =   90
               TabIndex        =   15
               Top             =   1200
               Width           =   1275
            End
            Begin VB.Label Label5 
               Caption         =   "Title Of Head Emp."
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
               Left            =   2730
               TabIndex        =   14
               Top             =   750
               Width           =   1365
            End
         End
         Begin VB.CommandButton cmdNew 
            Caption         =   "New"
            Enabled         =   0   'False
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
            Left            =   2760
            TabIndex        =   8
            Top             =   4770
            Width           =   1215
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit OU Types"
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
            Left            =   1200
            TabIndex        =   7
            Top             =   4770
            Width           =   1335
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
            Height          =   495
            Left            =   4200
            TabIndex        =   9
            Top             =   4770
            Width           =   1215
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
            Height          =   495
            Left            =   5580
            TabIndex        =   10
            Top             =   4770
            Width           =   1335
         End
         Begin MSComctlLib.ProgressBar pbrOUT 
            Height          =   135
            Left            =   0
            TabIndex        =   12
            Top             =   4560
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   238
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Label lblProgress 
            Caption         =   "l"
            Height          =   255
            Left            =   0
            TabIndex        =   21
            Top             =   4410
            Width           =   4095
         End
      End
      Begin VB.Label Label6 
         Caption         =   "Organization Unit Types"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   22
         Top             =   180
         Width           =   3435
      End
   End
End
Attribute VB_Name = "frmOUTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private out As HRCORE.OrganizationUnitType
Attribute out.VB_VarHelpID = -1
Private WithEvents outs As HRCORE.OrganizationUnitTypes
Attribute outs.VB_VarHelpID = -1
Private OldOUTs As HRCORE.OrganizationUnitTypes 'will hold old copy of data before editing
Private selOUT As OrganizationUnitType
Private NodeWasClicked As Boolean
Private InsertInProgress As Boolean
Private newOUT As HRCORE.OrganizationUnitType
Private company As New HRCORE.CompanyDetails
Private ChangedFromCode As Boolean


Private Sub cboParentOUTypes_Click()
    Dim TheSelID As Long
    
    If Not ChangedFromCode Then
        If Not (selOUT Is Nothing) Then
            If cboParentOUTypes.ListIndex <> -1 Then
                theID = cboParentOUTypes.ItemData(cboParentOUTypes.ListIndex)
                Set selOUT.ParentOUType = outs.FindOUType(theID)
                selOUT.IsModified = True
            End If
        End If
    End If
End Sub

Private Sub cmdCancel_Click()

End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim resp As Long
    Dim retVal As Long
    If Not currUser Is Nothing Then
        If currUser.CheckRight("OUType") <> secModify Then
            MsgBox "You don't have right to delete the record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    Select Case UCase(cmdDelete.Caption)
        Case "DELETE"
            If Not (selOUT Is Nothing) Then
                resp = MsgBox("Are you sure you want to delete the Organization Unit" & vbNewLine & _
                UCase(selOUT.OUTypeName), vbQuestion + vbYesNo, "Confirm Deletion")
                If resp = vbYes Then
                    retVal = selOUT.Delete()
                    
                    'upon success, reload the ou types afresh
                    If retVal = 0 Then
                        LoadOUTypes
                    End If
                End If
            Else
                MsgBox "There is no Organization Unit Type Selected", vbInformation, "Organization Unit Types"
            End If
        Case "CANCEL"
            'restore old data
            ClearFields
            RestoreOldOUTypes
            
            'restore controls
            cmdNew.Enabled = False
            cmdEdit.Caption = "Edit OU Types"
            cmdDelete.Caption = "Delete"
            fraDetails.Enabled = False
    End Select
End Sub

Private Sub cmdEdit_Click()
    
    If Not currUser Is Nothing Then
        If currUser.CheckRight("OUType") <> secModify Then
            MsgBox "You don't have right to edit or add record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    Select Case UCase(cmdEdit.Caption)
        Case "EDIT OU TYPES"
            'unlock controls
            fraDetails.Enabled = True
            cmdNew.Enabled = True
            cmdEdit.Caption = "Update"
            cmdDelete.Caption = "Cancel"
        Case "UPDATE"
            If SaveChanges() = False Then Exit Sub
            cmdNew.Enabled = False
            cmdEdit.Caption = "Edit OU Types"
            cmdDelete.Caption = "Delete"
            fraDetails.Enabled = False
    End Select
End Sub

Private Sub cmdNew_Click()
    ClearFields
    SetNewOUTypeDefaults
End Sub

Private Function SaveChanges() As Boolean
    SaveChanges = False
    If outs.ValidateKids() = True Then
        MsgBox "Validation Succeeded", vbInformation, "Prototype"
        If outs.UpdateChanges() = True Then
            'reload data from db
            LoadOUTypes
            MsgBox "The Organization Unit Types have been Updated Successfully", vbInformation, "Prototype"
            SaveChanges = True
        End If
    Else
        MsgBox "Validation Failed"
    End If
End Function

Private Sub Form_Load()
    Dim myOUT As OrganizationUnitType
    Dim myNode As Node
    'position the form
    'PositionForm Me
    
    Set out = New HRCORE.OrganizationUnitType
    Set outs = New HRCORE.OrganizationUnitTypes
    Set OldOUTs = New HRCORE.OrganizationUnitTypes
    company.LoadCompanyDetails
    
    'call procedure to load OU Types
    LoadOUTypes
    
    frmMain2.PositionTheFormWithoutEmpList Me
End Sub


Private Sub Form_Resize()
    oSmart.FResize Me
End Sub

Private Sub outs_FailedToInsert(ByVal TheNewEntity As HRCORE.OrganizationUnitType, actiontotake As HRCORE.ActionOnFailure)
    MsgBox "Failed to Insert the Entity: " & TheNewEntity.OUTypeCode & ": " & TheNewEntity.OUTypeName
End Sub

Private Sub outs_FailedToUpdate(ByVal TheModifiedEntity As HRCORE.OrganizationUnitType, actiontotake As HRCORE.ActionOnFailure)
    MsgBox "Failed to Update the Entity: " & TheModifiedEntity.OUTypeCode & ": " & TheModifiedEntity.OUTypeName
End Sub

Private Sub outs_FailedValidation(ByVal TheContainer As HRCORE.OrganizationUnitTypes, ByVal InvalidOUTypes As HRCORE.OrganizationUnitTypes)
    Dim i As Long
    Dim j As Long
    Dim invOUT As HRCORE.OrganizationUnitType
    MsgBox "The OU Types highlighted in Red have invalid Data", vbInformation, "Prototype"
    
    For i = 1 To InvalidOUTypes.count
        Set invOUT = InvalidOUTypes.Item(i)
        For j = 1 To Me.tvwOUTypes.Nodes.count
            If (Me.tvwOUTypes.Nodes(j).Tag <> "") And (Me.tvwOUTypes.Nodes(j).Tag <> "HRCOMPANY") Then
                If invOUT.OUTypeID = CLng(Me.tvwOUTypes.Nodes(j).Tag) Then
                    Me.tvwOUTypes.Nodes(j).ForeColor = vbRed
                    Me.tvwOUTypes.Nodes(j).Bold = True
                End If
            End If
        Next j
    Next i
    
End Sub

Private Sub outs_FinishedInsertingNewEntities()
    lblProgress.Caption = "Finished Inserting New Entities"
End Sub

Private Sub outs_FinishedUpdatingModifiedEntities()
    lblProgress.Caption = "Finished Updating Modified Entities"
End Sub

Private Sub outs_InsertingNewEntity(ByVal TheNewEntity As HRCORE.OrganizationUnitType, Cancel As Boolean)
    pbrOUT.value = pbrOUT.value + 1
    lblProgress.Caption = "Inserting: " & TheNewEntity.OUTypeName
'    MsgBox "The System is now inserting" & vbNewLine & _
'    TheNewEntity.OUTypeCode & ": " & TheNewEntity.OUTypeName, vbInformation, "Inserting New Entity"
    DoEvents
End Sub

Private Sub outs_SynchronizingChild(ByVal TheOldParentID As Long, ByVal TheNewParentID As Long, Cancel As Boolean)
    MsgBox "Old Parent ID = " & TheOldParentID & vbNewLine & "The New Parent ID = " & TheNewParentID
End Sub

Private Sub outs_UpdatingModifiedEntity(ByVal TheModifiedEntity As HRCORE.OrganizationUnitType, Cancel As Boolean)
    pbrOUT.value = pbrOUT.value + 1
    lblProgress.Caption = "Updating: " & TheModifiedEntity.OUTypeName
    DoEvents
End Sub

Private Sub outs_WillInsertNewEntities(ByVal TheNewEntities As HRCORE.OrganizationUnitTypes, Cancel As Boolean)
    lblProgress.Caption = "Inserting: " & TheNewEntities.count & " New Entities"
    pbrOUT.Min = 0
    pbrOUT.Max = TheNewEntities.count
    pbrOUT.value = 0
    DoEvents
End Sub

Private Sub outs_WillUpdateModifiedEntities(ByVal TheModifiedEntities As HRCORE.OrganizationUnitTypes, Cancel As Boolean)
    lblProgress.Caption = "Updating: " & TheModifiedEntities.count & " Modified Entities"
    pbrOUT.Min = 0
    pbrOUT.Max = TheModifiedEntities.count
    pbrOUT.value = 0
    DoEvents
End Sub

Private Sub tvwOUTypes_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim theID As Long
    If (Node.Tag = "HRCOMPANY") Or (Node.Tag = "") Then
        Set selOUT = Nothing 'show nothing is selected
        ClearFields
    Else
        theID = CLng(Node.Tag)
    End If
    Set selOUT = outs.FindOUType(theID)     'It returns the reference
    If Not (selOUT Is Nothing) Then
        NodeWasClicked = True
        SetFields selOUT
    End If
End Sub

Private Sub SetFields(ByVal TheOUType As OrganizationUnitType)
    Dim i As Long
    Dim Found As Boolean
    ChangedFromCode = True
    With TheOUType
        Me.txtOUTCode.Text = .OUTypeCode
        Me.txtOUTLevel.Text = .OUTypeLevel
        Me.txtOUTName.Text = .OUTypeName
        Me.txtTitleOfHeadEmp.Text = .TitleOfHeadEmployee
        If Not (.ParentOUType Is Nothing) Then
            SelectParentInCombo .ParentOUType.OUTypeID
        Else
            cboParentOUTypes.Text = "COMPANY"
        End If
    End With
    ChangedFromCode = False
End Sub

Private Sub SelectParentInCombo(ByVal TheItemID As Long)
    Dim Found As Boolean
    Dim i As Long
    
    For i = 0 To cboParentOUTypes.ListCount - 1
        If cboParentOUTypes.ItemData(i) = TheItemID Then
            cboParentOUTypes.ListIndex = i
            Found = True
        End If
    Next
        
    If Not Found Then
        cboParentOUTypes.ListIndex = -1
    End If
    
End Sub

Private Sub ClearFields()
    'to avoid automatic update of objects
    ChangedFromCode = True
    
    Me.txtOUTCode.Text = ""
    Me.txtOUTLevel.Text = ""
    Me.txtOUTName.Text = ""
    Me.txtTitleOfHeadEmp.Text = ""
    Me.cboParentOUTypes.ListIndex = -1
    
    ChangedFromCode = False
End Sub

Private Sub SetNewOUTypeDefaults()
    
    Dim newCount As Long
    Dim NewID As Long
    Dim newParent As HRCORE.OrganizationUnitType
    Dim newNode As Node

    Set newOUT = New HRCORE.OrganizationUnitType
    If Not (outs Is Nothing) Then
        newCount = outs.count + 1
        NewID = outs.GetNextKidID()
    Else
        newCount = 1
        NewID = 1
    End If
    
    With newOUT
        .OUTypeCode = "New Code " & newCount
        .OUTypeLevel = 1
        .OUTypeID = NewID
        .InsertionOrderNo = NewID
        .OUTypeName = "New OU Type " & newCount
        .TitleOfHeadEmployee = "New Head"
        .IsNewEntity = True
        If Not (selOUT Is Nothing) Then
            Set .ParentOUType = selOUT
            'SelectParentInCombo .ParentOUType.OUTypeID
            .OUTypeLevel = .ParentOUType.OUTypeLevel + 1
            Set newNode = tvwOUTypes.Nodes.add("OUT:" & .ParentOUType.OUTypeID, tvwChild, "OUT:" & .OUTypeID, newOUT.OUTypeName)
            newNode.Selected = True
            newNode.Tag = NewID
            newNode.EnsureVisible
        Else
            Set .ParentOUType = Nothing
            Set newNode = tvwOUTypes.Nodes.add("HRCOMPANY", tvwChild, "OUT:" & newOUT.OUTypeID, newOUT.OUTypeName)
            newNode.Tag = NewID
            newNode.Selected = True
            newNode.EnsureVisible
        End If
    End With
    'refresh the combobox
    RefreshParentCombo
    
    If Not (outs Is Nothing) Then
        outs.add newOUT
    Else
        Set outs = New HRCORE.OrganizationUnitTypes
        outs.add newOUT
    End If
    
        
    SetFields newOUT
    InsertInProgress = True
            
End Sub

Private Sub RefreshParentCombo(Optional ByVal RetainSelection As Boolean)
    Dim out As OrganizationUnitType
    Dim prevValue As String
    
    prevValue = cboParentOUTypes.Text
    cboParentOUTypes.Clear
    'add the company first
    cboParentOUTypes.AddItem "COMPANY"
    For Each out In outs
        cboParentOUTypes.AddItem out.OUTypeName
        cboParentOUTypes.ItemData(cboParentOUTypes.NewIndex) = out.OUTypeID
    Next
    
    If Not IsMissing(RetainSelection) Then
        If RetainSelection = True Then
            If Not (selOUT Is Nothing) Then
                If (selOUT.ParentOUType Is Nothing) Or (selOUT.ParentOUType.OUTypeID <= 0) Then
                    cboParentOUTypes.Text = "COMPANY"
                Else
                    cboParentOUTypes.Text = selOUT.ParentOUType.OUTypeName
                End If
            End If
        End If
    End If
        
    
End Sub

Private Sub txtOUTCode_Change()
    If Not ChangedFromCode Then
        If Not (selOUT Is Nothing) Then
            selOUT.OUTypeCode = txtOUTCode.Text
            selOUT.IsModified = True
            'commented coz it seems the objects pass the reference
            'outs.ChangeKidFields selOUT
        End If
    End If
End Sub

Private Sub txtOUTLevel_Change()
    If Not ChangedFromCode Then
        If Not (selOUT Is Nothing) Then
            If IsNumeric(Trim(txtOUTLevel.Text)) Then
                selOUT.OUTypeLevel = CLng(Trim(txtOUTLevel.Text))
                selOUT.IsModified = True
            End If
        End If
    End If
End Sub

Private Sub txtOUTName_Change()
    If Not ChangedFromCode Then
        If Not (selOUT Is Nothing) Then
            selOUT.OUTypeName = txtOUTName.Text
            'outs.ChangeKidFields selOUT
            Me.tvwOUTypes.SelectedItem.Text = selOUT.OUTypeName
            selOUT.IsModified = True
            'refresh the combobox to reflect the changes
            RefreshParentCombo True
        End If
    End If
End Sub

Private Sub txtTitleOfHeadEmp_Change()
    If Not ChangedFromCode Then
        If Not (selOUT Is Nothing) Then
            selOUT.TitleOfHeadEmployee = Trim(txtTitleOfHeadEmp.Text)
            selOUT.IsModified = True
        End If
    End If
End Sub

Private Sub LoadOUTypes()
    Dim myOUT As OrganizationUnitType
    Dim myNode As Node
            
    outs.GetAllOUTypes
    
    'clear old data
    OldOUTs.Clear
    
    cboParentOUTypes.Clear
    tvwOUTypes.Nodes.Clear
    
    Set myNode = tvwOUTypes.Nodes.add(, , "HRCOMPANY", "COMPANY: " & company.CompanyName)
    myNode.Bold = True
    
    cboParentOUTypes.AddItem "COMPANY"
    For Each myOUT In outs
        cboParentOUTypes.AddItem myOUT.OUTypeName
        cboParentOUTypes.ItemData(cboParentOUTypes.NewIndex) = myOUT.OUTypeID
        
        'keep copy of this initial data for restore purposes
        OldOUTs.add myOUT
        
        'force to get parent
        Set myOUT.ParentOUType = outs.FindOUType(myOUT.ParentOUType.OUTypeID)
        If (myOUT.ParentOUType Is Nothing) Or (myOUT.ParentOUType.OUTypeID <= 0) Then
            Set myNode = Me.tvwOUTypes.Nodes.add("HRCOMPANY", tvwChild, "OUT:" & myOUT.OUTypeID, myOUT.OUTypeName)
            myNode.Tag = myOUT.OUTypeID
        Else
            Set myNode = Me.tvwOUTypes.Nodes.add("OUT:" & myOUT.ParentOUType.OUTypeID, tvwChild, "OUT:" & myOUT.OUTypeID, myOUT.OUTypeName)
            myNode.Tag = myOUT.OUTypeID
        End If
        myNode.EnsureVisible
    Next
    
    'select an item
    If Me.tvwOUTypes.Nodes.count > 1 Then
        Me.tvwOUTypes.Nodes(2).Selected = True
        Set selOUT = outs.FindOUType(CLng(Me.tvwOUTypes.Nodes(2).Tag))
        SetFields selOUT
    End If
End Sub

Private Sub RestoreOldOUTypes()
    Dim myOUT As OrganizationUnitType
    Dim myNode As Node
            
    
    cboParentOUTypes.Clear
    tvwOUTypes.Nodes.Clear
    
    'clear current outs
    outs.Clear
    
    Set myNode = tvwOUTypes.Nodes.add(, , "HRCOMPANY", "COMPANY: " & company.CompanyName)
    myNode.Bold = True
    
    cboParentOUTypes.AddItem "COMPANY"
    For Each myOUT In OldOUTs
        cboParentOUTypes.AddItem myOUT.OUTypeName
        cboParentOUTypes.ItemData(cboParentOUTypes.NewIndex) = myOUT.OUTypeID
        'now copy
        outs.add myOUT
        'force to get parent
        Set myOUT.ParentOUType = OldOUTs.FindOUType(myOUT.ParentOUType.OUTypeID)
        If (myOUT.ParentOUType Is Nothing) Or (myOUT.ParentOUType.OUTypeID <= 0) Then
            Set myNode = Me.tvwOUTypes.Nodes.add("HRCOMPANY", tvwChild, "OUT:" & myOUT.OUTypeID, myOUT.OUTypeName)
            myNode.Tag = myOUT.OUTypeID
        Else
            Set myNode = Me.tvwOUTypes.Nodes.add("OUT:" & myOUT.ParentOUType.OUTypeID, tvwChild, "OUT:" & myOUT.OUTypeID, myOUT.OUTypeName)
            myNode.Tag = myOUT.OUTypeID
        End If
        myNode.EnsureVisible
    Next
    
    'select an item
    If Me.tvwOUTypes.Nodes.count > 1 Then
        Me.tvwOUTypes.Nodes(2).Selected = True
        Set selOUT = outs.FindOUType(CLng(Me.tvwOUTypes.Nodes(2).Tag))
        SetFields selOUT
    End If
End Sub

