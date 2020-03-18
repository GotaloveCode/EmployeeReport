VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOUs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Organization Units"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   8655
   Begin VB.Frame fraTop 
      Height          =   6225
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   8475
      Begin VB.Frame fraData 
         Height          =   6045
         Left            =   90
         TabIndex        =   13
         Top             =   180
         Width           =   8295
         Begin VB.CommandButton cmdChangeParent 
            Caption         =   "Change Parent"
            Height          =   495
            Left            =   240
            TabIndex        =   23
            Top             =   5400
            Width           =   1455
         End
         Begin VB.Frame fraDetails 
            Caption         =   "Organization Unit Details"
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
            Height          =   2085
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   8295
            Begin VB.TextBox txtOUCode 
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
               Left            =   1290
               TabIndex        =   1
               Top             =   360
               Width           =   2115
            End
            Begin VB.TextBox txtOUName 
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
               Left            =   5130
               TabIndex        =   2
               Top             =   360
               Width           =   3015
            End
            Begin VB.ComboBox cboParentOU 
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
               Left            =   5130
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   1140
               Width           =   3015
            End
            Begin VB.TextBox txtTelephone 
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
               Left            =   1290
               TabIndex        =   3
               Top             =   750
               Width           =   2115
            End
            Begin VB.TextBox txtEMail 
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
               Left            =   5130
               TabIndex        =   4
               Top             =   750
               Width           =   3015
            End
            Begin VB.ComboBox cboOUType 
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
               Left            =   1290
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   1140
               Width           =   2115
            End
            Begin VB.TextBox txtOUFunction 
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
               Left            =   1650
               MultiLine       =   -1  'True
               TabIndex        =   7
               Top             =   1620
               Width           =   6525
            End
            Begin VB.Label Label1 
               Caption         =   "Org. Unit Code"
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
               TabIndex        =   22
               Top             =   390
               Width           =   1335
            End
            Begin VB.Label Label2 
               Caption         =   "Org. Unit Name"
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
               Left            =   3900
               TabIndex        =   21
               Top             =   390
               Width           =   1335
            End
            Begin VB.Label Label4 
               Caption         =   "Parent Org. Unit."
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
               Left            =   3900
               TabIndex        =   20
               Top             =   1170
               Width           =   1335
            End
            Begin VB.Label Label3 
               Caption         =   "Telephone"
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
               TabIndex        =   19
               Top             =   780
               Width           =   945
            End
            Begin VB.Label Label5 
               Caption         =   "E-Mail Address"
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
               Left            =   3900
               TabIndex        =   18
               Top             =   780
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Org. Unit Type"
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
               TabIndex        =   17
               Top             =   1170
               Width           =   1335
            End
            Begin VB.Label Label7 
               Caption         =   "Org. Unit Function"
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
               TabIndex        =   16
               Top             =   1620
               Width           =   1815
            End
         End
         Begin VB.Frame fraHierarchy 
            Caption         =   "Organization Units"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3165
            Left            =   0
            TabIndex        =   14
            Top             =   2160
            Width           =   8235
            Begin MSComctlLib.TreeView tvwOUnits 
               Height          =   2775
               Left            =   120
               TabIndex        =   8
               Top             =   270
               Width           =   7905
               _ExtentX        =   13944
               _ExtentY        =   4895
               _Version        =   393217
               HideSelection   =   0   'False
               Indentation     =   353
               LabelEdit       =   1
               Style           =   7
               FullRowSelect   =   -1  'True
               Appearance      =   1
            End
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
            Height          =   495
            Left            =   3660
            TabIndex        =   10
            Top             =   5430
            Width           =   1335
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit Org. Units"
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
            Left            =   1920
            TabIndex        =   9
            Top             =   5430
            Width           =   1425
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
            Left            =   5220
            TabIndex        =   11
            Top             =   5400
            Width           =   1335
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
            Left            =   6840
            TabIndex        =   12
            Top             =   5400
            Width           =   1395
         End
      End
   End
End
Attribute VB_Name = "frmOUs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents OUnits As HRCORE.OrganizationUnits
Attribute OUnits.VB_VarHelpID = -1
Private OldOUs As HRCORE.OrganizationUnits
Private outs As HRCORE.OrganizationUnitTypes
Private selOU As HRCORE.OrganizationUnit
Private selOUT As HRCORE.OrganizationUnitType
Private company As New HRCORE.CompanyDetails
Private newOU As HRCORE.OrganizationUnit
Private ChangedFromCode As Boolean
Private TopLevelOUnits As HRCORE.OrganizationUnits

Private OUChildrenMovedSuccessfully As Boolean     'will track moving of OUs
Private RestoringOldOUnits As Boolean       'will flag when Restoring Old OUs in AddChildrenRecursively() i.e. When Cancel is clicked
Private NodeToDropTo As Node 'the node to be dropped onto

Private nodeX As Node   'will hold a node being dragged
Private InDragMode As Boolean   'flag indicating drag operation

Private Sub cmdChangeParent_Click()
    Dim newParent As HRCORE.OrganizationUnit
    Dim resp As Long
    
    'On Error GoTo errorHandler
    
    If selOU Is Nothing Then
        MsgBox "There is no Organization Unit selected", vbInformation, TITLES
        Exit Sub
    End If
    resp = MsgBox("An Organization Unit can be a Top-Level OU or can belong to another OU." & vbNewLine & _
        "Do you want the selected Organization Unit i.e. " & vbNewLine & selOU.OrganizationUnitName & vbNewLine & " to become a Top-Level OU?", vbYesNoCancel + vbQuestion, TITLES)
    If resp = vbYes Then
        Set selOU.ParentOU = Nothing
    ElseIf resp = vbCancel Then
        Exit Sub
    Else
        resp = MsgBox("Select the Organization Unit that will become the New Parent", vbInformation + vbOKCancel, TITLES)
        If resp = vbOK Then
            Set newParent = OUnits.SelectSingleOrganizationUnit()
            If Not (newParent Is Nothing) Then
                'check whether user selected same OU to be Parent
                If newParent.OrganizationUnitID = selOU.OrganizationUnitID Then
                    MsgBox "An Organization Unit cannot be a parent to itself", vbInformation, TITLES
                    Exit Sub
                Else
                    Set selOU.ParentOU = newParent
                    Set selOU.ParentOU.OUType = outs.FindOUType(selOU.ParentOU.OUType.OUTypeID)
                End If
            Else
                MsgBox "No Parent Organization Unit was Set", vbInformation, TITLES
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If
    
    'update the OUType also
    If Not (selOU.ParentOU Is Nothing) Then
        If selOU.ParentOU.OrganizationUnitID = 0 Then
            Set selOU.OUType = outs.GetFirstTopmostOUType()
            If selOU.OUType Is Nothing Then
                MsgBox "The Organization Unit Level could not be determined", vbInformation, TITLES
                Exit Sub
            End If
        Else
            If Not (selOU.ParentOU.OUType Is Nothing) Then
                If selOU.ParentOU.OUType.ChildOUTypeName <> "----" Then
                    Set selOU.OUType = outs.FindOUType(selOU.ParentOU.OUType.ChildOUTypeID)
                Else
                    MsgBox "The Lowest Organization Unit Level would be exceeded", vbInformation, TITLES
                    Exit Sub
                End If
            Else
                MsgBox "The Organization Unit Level could not be determined", vbInformation, TITLES
                Exit Sub
            End If
        End If
    Else
        Set selOU.OUType = outs.GetFirstTopmostOUType()
        If selOU.OUType Is Nothing Then
            MsgBox "The Organization Unit Level could not be determined", vbInformation, TITLES
            Exit Sub
        End If
    End If
    
    'assume child ous were not moved successfully
    OUChildrenMovedSuccessfully = True
    
    'then now try to move them recursively
    MoveChildOUsRecursively selOU
    
    'check again status of the Flag
    If OUChildrenMovedSuccessfully Then
        'update the changes
        retVal = selOU.Update()
        
        'update the children
        UpdateMovingChildOUsRecursively selOU
    End If
    
    'reload the ous
    LoadOrganizationUnitsEx
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while moving the Organization Unit" & vbNewLine & err.Description, vbExclamation, TITLES
    
End Sub

Private Sub MoveChildOUsRecursively(ByVal theOU As HRCORE.OrganizationUnit)
    
    'this is a recursive function that populates child ous
  
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    If Not (theOU Is Nothing) Then
        For i = 1 To theOU.Children.count
             If theOU.Children.Item(i).ParentOU.OUType.ChildOUTypeName <> "----" Then
                'get the OUType Object using the ID
                Set theOU.Children.Item(i).OUType = outs.FindOUType(theOU.Children.Item(i).ParentOU.OUType.ChildOUTypeID)
                
            Else
                'flag that it is not successful
                MsgBox "Cannot Move the Organization Unit: " & theOU.Children.Item(i).OrganizationUnitName & vbNewLine & _
                "This would exceed the Lowest Level of Organization Unit Hierarchy", vbInformation, TITLES
                
                OUChildrenMovedSuccessfully = False
                Exit Sub
            End If
            'recursively load the children
            MoveChildOUsRecursively theOU.Children.Item(i)
        Next i
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred" & vbNewLine & err.Description, vbInformation, TITLES
        
End Sub

Private Sub UpdateMovingChildOUsRecursively(ByVal theOU As HRCORE.OrganizationUnit)
    
    'this is a recursive function that Updates changes in child OUs due to moving the Parent
  
    Dim i As Long
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    
    If Not (theOU Is Nothing) Then
        For i = 1 To theOU.Children.count
            retVal = theOU.Children.Item(i).Update()
            
            'recursively load the children
            UpdateMovingChildOUsRecursively theOU.Children.Item(i)
        Next i
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred" & vbNewLine & err.Description, vbInformation, TITLES
        
End Sub


Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim resp As Long
    Dim retVal As Long
     
    If Not currUser Is Nothing Then
        If currUser.CheckRight("OrganizationUnits") <> secModify Then
            MsgBox "You don't have right to delete record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    Select Case UCase(cmdDelete.Caption)
        Case "DELETE"
        'If Not (selOUT Is Nothing) Then
            If Not (selOU Is Nothing) Then
                resp = MsgBox("Are you sure you want to delete the Organization Unit" & vbNewLine & _
                UCase(selOU.OUType.OUTypeName), vbQuestion + vbYesNo, "Confirm Deletion")
                If resp = vbYes Then
                    retVal = selOU.Delete()
                    
                    'upon success, reload the ou types afresh
                    If retVal = 0 Then
                        LoadOrganizationUnitsEx
                    End If
                End If
            Else
                MsgBox "There is no Organization Unit Selected", vbInformation, "Organization Units"
            End If
        Case "CANCEL"
            'restore old data
            ClearFields
            RestoreOldOrganizationUnits
            
            'restore controls
            cmdNew.Enabled = False
            cmdEdit.Caption = "Edit Org. Units"
            cmdDelete.Caption = "Delete"
            fraDetails.Enabled = False
            cmdChangeParent.Enabled = True
    End Select
End Sub

Private Sub cmdEdit_Click()
    
    If Not currUser Is Nothing Then
        If currUser.CheckRight("OrganizationUnits") <> secModify Then
            MsgBox "You don't have right to edit or add record. Please liaise with the security admin"
            Exit Sub
        End If
    End If

    Select Case UCase(cmdEdit.Caption)
        Case "EDIT ORG. UNITS"
            'unlock controls
            fraDetails.Enabled = True
            cmdNew.Enabled = False
            cmdEdit.Caption = "Update"
            cmdDelete.Caption = "Cancel"
            cmdChangeParent.Enabled = False
        Case "UPDATE"
            If SaveChanges() = False Then Exit Sub
            cmdNew.Enabled = False
            cmdEdit.Caption = "Edit Org. Units"
            cmdNew.Enabled = True
            cmdDelete.Caption = "Delete"
            fraDetails.Enabled = False
            cmdChangeParent.Enabled = True
    End Select
    
End Sub

Private Sub cmdNew_Click()

    ClearFields
    
     'unlock controls
            fraDetails.Enabled = True
            cmdNew.Enabled = True
            cmdEdit.Caption = "Update"
            cmdDelete.Caption = "Cancel"
            cmdChangeParent.Enabled = False
            
    SetNewOUDefaults
    cmdNew.Enabled = False
    cmdEdit.Enabled = True
    cmdEdit.Caption = "UPDATE"
End Sub

Private Sub Form_Load()
    Set OUnits = New HRCORE.OrganizationUnits
    Set outs = New HRCORE.OrganizationUnitTypes
    Set OldOUs = New HRCORE.OrganizationUnits
    company.LoadCompanyDetails
    
    LoadOUTypes
    'LoadOrganizationUnits 'this works perfectly
    LoadOrganizationUnitsEx 'Done on 13th May 2005 By Oscar To allow for re-organizing OUs
    
        
    frmMain2.PositionTheFormWithoutEmpList Me
    
hidemaincmds

End Sub

Private Sub hidemaincmds()
    frmMain2.cmdNew.Enabled = False
    frmMain2.cmdEdit.Enabled = False
    frmMain2.cmdSave.Enabled = False
    frmMain2.cmdCancel.Enabled = False
    frmMain2.cmdDelete.Enabled = False
End Sub

Private Sub LoadOUTypes()
    Dim myOUT As HRCORE.OrganizationUnitType
    
    outs.GetAllOUTypes
    
    For Each myOUT In outs
        cboOUType.AddItem myOUT.OUTypeName
        cboOUType.ItemData(cboOUType.NewIndex) = myOUT.OUTypeID
    Next
            
End Sub

Private Sub SelectParentInCombo(ByVal TheOUID As Long)
    Dim Found As Boolean
    Dim i As Long
    
    For i = 0 To cboParentOU.ListCount - 1
        If cboParentOU.ItemData(i) = TheOUID Then
            cboParentOU.ListIndex = i
            Found = True
        End If
    Next
        
    If Not Found Then
        cboParentOU.ListIndex = -1
    End If
    
End Sub

Private Sub SelectOUTypeInCombo(ByVal TheOUTypeID As Long)
    Dim Found As Boolean
    Dim i As Long
    
    For i = 0 To cboOUType.ListCount - 1
        If cboOUType.ItemData(i) = TheOUTypeID Then
            cboOUType.ListIndex = i
            Found = True
        End If
    Next
    
    If Not Found Then
        cboOUType.ListIndex = -1
    End If
        
End Sub


Private Sub LoadOrganizationUnits()
    Dim myOU As HRCORE.OrganizationUnit
    Dim myNode As Node
    Dim i As Long
    
    OUnits.GetAllOrganizationUnits
    
    'clear the oldous
    OldOUs.Clear
    
    'clear lists
    tvwOUnits.Nodes.Clear
    cboParentOU.Clear
    
    'first add the company
    cboParentOU.AddItem company.CompanyName
    
    Set myNode = tvwOUnits.Nodes.add(, , "HRCOMPANY", company.CompanyName)
    myNode.Tag = "HRCOMPANY"
    myNode.Bold = True
    
    For i = 1 To OUnits.count
        Set myOU = OUnits.Item(i)
        
        'Force the OU Type details to be loaded
        Set myOU.OUType = outs.FindOUType(myOU.OUType.OUTypeID)
        
        'Force ParentOU to be loaded
        Set myOU.ParentOU = OUnits.FindOrganizationUnit(myOU.ParentOU.OrganizationUnitID)
        
        'populate the collection to hold initial OUs
        OldOUs.add myOU
        
        cboParentOU.AddItem myOU.OrganizationUnitName
        cboParentOU.ItemData(cboParentOU.NewIndex) = myOU.OrganizationUnitID
        
        If (myOU.ParentOU Is Nothing) Or (myOU.ParentOU.OrganizationUnitID <= 0) Then
            Set myNode = tvwOUnits.Nodes.add("HRCOMPANY", tvwChild, "OU:" & myOU.OrganizationUnitID, myOU.OrganizationUnitName)
            myNode.Tag = myOU.OrganizationUnitID
        Else
            Set myNode = tvwOUnits.Nodes.add("OU:" & myOU.ParentOU.OrganizationUnitID, tvwChild, "OU:" & myOU.OrganizationUnitID, myOU.OrganizationUnitName)
            myNode.Tag = myOU.OrganizationUnitID
        End If
        myNode.EnsureVisible
    Next i
    
    'select an item
    If Me.tvwOUnits.Nodes.count > 1 Then
        Me.tvwOUnits.Nodes(2).Selected = True
        Set selOU = OUnits.FindOrganizationUnit(CLng(Me.tvwOUnits.Nodes(2).Tag))
        SetOUFields selOU
    End If
End Sub


Private Sub RestoreOldOrganizationUnits()
    Dim myOU As HRCORE.OrganizationUnit
    Dim myNode As Node
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    'clear the current OUs
    OUnits.Clear
    
    'clear lists
    tvwOUnits.Nodes.Clear
    cboParentOU.Clear
    
    'first add the company
    cboParentOU.AddItem company.CompanyName
    
    Set myNode = tvwOUnits.Nodes.add(, , "HRCOMPANY", company.CompanyName)
    myNode.Tag = "HRCOMPANY"
    myNode.Bold = True
    
    'now get the Top Level Organization Units
    Set TopLevelOUnits = OldOUs.GetOrganizationUnitsOfTopmostLevel()
    
    If Not (TopLevelOUnits Is Nothing) Then
        For i = 1 To TopLevelOUnits.count
            Set myOU = TopLevelOUnits.Item(i)
        
            'Force the OU Type details to be loaded
            If Not (myOU.OUType Is Nothing) Then
                Set myOU.OUType = outs.FindOUType(myOU.OUType.OUTypeID)
            End If
        
            'Force ParentOU to be loaded
            'Set myOU.ParentOU = OldOUs.FindOrganizationUnit(myOU.ParentOU.OrganizationUnitID)
        
            'populate the collection to hold the new current OUs
            OUnits.add myOU
        
            cboParentOU.AddItem myOU.OrganizationUnitName
            cboParentOU.ItemData(cboParentOU.NewIndex) = myOU.OrganizationUnitID
            
            'add the OU
            Set myNode = Me.tvwOUnits.Nodes.add(, , "OU:" & myOU.OrganizationUnitID, myOU.OrganizationUnitName)
            myNode.Tag = myOU.OrganizationUnitID
            myNode.EnsureVisible
            
            'Flag to indicate that Restore Operation is in progress
            'so that OUnits will be reloaded
            RestoringOldOUnits = True
            
            'now recursively add the children
            AddChildOUsRecursively myOU
            
            'Indicate end of restoring old ous
            RestoringOldOUnits = False
            
'            If (myOU.ParentOU Is Nothing) Or (myOU.ParentOU.OrganizationUnitID <= 0) Then
'                Set myNode = tvwOUnits.Nodes.Add("HRCOMPANY", tvwChild, "OU:" & myOU.OrganizationUnitID, myOU.OrganizationUnitName)
'                myNode.Tag = myOU.OrganizationUnitID
'            Else
'                Set myNode = tvwOUnits.Nodes.Add("OU:" & myOU.ParentOU.OrganizationUnitID, tvwChild, "OU:" & myOU.OrganizationUnitID, myOU.OrganizationUnitName)
'                myNode.Tag = myOU.OrganizationUnitID
'            End If
'            myNode.EnsureVisible
        Next i
    End If
    
       
    'select an item
    If Me.tvwOUnits.Nodes.count > 1 Then
        Me.tvwOUnits.Nodes(2).Selected = True
        Set selOU = OUnits.FindOrganizationUnit(CLng(Me.tvwOUnits.Nodes(2).Tag))
        SetOUFields selOU
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while Refreshing the Organization Units" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub Form_Resize()
    oSmart.FResize Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
frmMain2.RestoreCommandButtonState
End Sub

Private Sub tvwOUnits_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim theID As Long
    
    On Error GoTo ErrorHandler
    
    If (Node.Tag = "HRCOMPANY") Or (Node.Tag = "") Then
        Set selOU = Nothing 'show nothing is selected
        ClearFields
    Else
        theID = CLng(Node.Tag)
        Set selOU = OUnits.FindOrganizationUnit(theID)
        
        If Not (selOU Is Nothing) Then
            SetOUFields selOU
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while processing the selected Organization Unit" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub SetNewOUDefaults()
    Dim newCount As Long
    Dim NewID As Long
    Dim newParent As HRCORE.OrganizationUnit
    Dim newOUType As HRCORE.OrganizationUnitType
    Dim newNode As Node
    
    
    On Error GoTo ErrorHandler
    
    Set newOU = New HRCORE.OrganizationUnit
    If Not (OUnits Is Nothing) Then
        newCount = OUnits.count + 1
        NewID = OUnits.GetNextKidID()
    Else
        newCount = 1
        NewID = 1
    End If
    
    With newOU
        .OrganizationUnitCode = "New Code " & newCount
        .EMail = company.EMailAddress
        .OrganizationUnitFunction = "New Function"
        .OrganizationUnitName = "New OU " & newCount
        .Telephone = company.Telephone1
        .OrganizationUnitID = NewID
        .InsertionOrderNo = NewID
        If Not (selOU Is Nothing) Then
            Set selOU.OUType = outs.FindOUType(selOU.OUType.OUTypeID)
            Set .ParentOU = selOU
            If Not (.ParentOU.OUType Is Nothing) Then
                If .ParentOU.OUType.ChildOUTypeName <> "----" Then
                    .OrganizationUnitName = "New " & .ParentOU.OUType.ChildOUTypeName & " " & newCount
                    'get the OUType Object using the ID
                    Set .OUType = outs.FindOUType(.ParentOU.OUType.ChildOUTypeID)
                Else
                    MsgBox UCase(selOU.OrganizationUnitName) & vbNewLine & _
                    "Is the Lowest in the Hierarchy of Organization Structures", vbInformation, "Prototype"
                    Exit Sub
                End If
            Else
                MsgBox "No Hierarchy Information was found the Selected Parent Organization Unit" & vbNewLine & _
                "The New Organization Unit cannot be created", vbInformation, TITLES
                Exit Sub
            End If
            
            Set newNode = tvwOUnits.Nodes.add("OU:" & .ParentOU.OrganizationUnitID, tvwChild, "OU:" & .OrganizationUnitID, newOU.OrganizationUnitName)
            newNode.Selected = True
            newNode.Tag = NewID
            newNode.EnsureVisible
        Else
            'if no Node is selected, always add the new OU to the Company
            Set .ParentOU = Nothing
            Set .OUType = outs.GetFirstTopmostOUType()
            .OrganizationUnitName = "New " & .OUType.OUTypeName & " " & newCount
            Set newNode = tvwOUnits.Nodes.add("HRCOMPANY", tvwChild, "OU:" & newOU.OrganizationUnitID, newOU.OrganizationUnitName)
            newNode.Tag = NewID
            newNode.Selected = True
            newNode.EnsureVisible
        End If
        
        newOU.IsNewEntity = True
    End With
    'refresh the combobox
    RefreshParentCombo
    
    If Not (OUnits Is Nothing) Then
        OUnits.add newOU
    Else
        Set OUnits = New HRCORE.OrganizationUnits
        OUnits.add newOU
    End If
    
    ClearFields
    SetOUFields newOU
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while creating a new Organization Unit" & vbNewLine & err.Description, vbExclamation, Title
End Sub


Private Sub SetOUFields(ByVal TheOUnit As HRCORE.OrganizationUnit)
    On Error GoTo ErrorHandler
    
    ChangedFromCode = True
    With TheOUnit
        Me.txtEmail.Text = .EMail
        Me.txtOUCode.Text = .OrganizationUnitCode
        Me.txtOUName.Text = .OrganizationUnitName
        Me.txtTelephone.Text = .Telephone
        If (.ParentOU Is Nothing) Or (.ParentOU.OrganizationUnitID <= 0) Then
            cboParentOU.Text = company.CompanyName
        Else
            SelectParentInCombo .ParentOU.OrganizationUnitID
        End If
        If Not (.OUType Is Nothing) Then
            SelectOUTypeInCombo .OUType.OUTypeID
        End If
        Me.txtOUFunction.Text = .OrganizationUnitFunction
    End With
    ChangedFromCode = False
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while displaying the data" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub ClearFields()
    Dim ctrl As Control
    
    ChangedFromCode = True
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is TextBox Then
            ctrl.Text = ""
        ElseIf TypeOf ctrl Is ComboBox Then
            ctrl.ListIndex = -1
        End If
    Next
    ChangedFromCode = False
End Sub


Private Sub RefreshParentCombo(Optional ByVal RetainSelection As Boolean)
    Dim ou As HRCORE.OrganizationUnit
    Dim i As Long
    
    'first clear the combo box
    cboParentOU.Clear
    
    'add the company first
    cboParentOU.AddItem company.CompanyName
    
    For i = 1 To OUnits.count
        Set ou = OUnits.Item(i)
        cboParentOU.AddItem ou.OrganizationUnitName
        cboParentOU.ItemData(cboParentOU.NewIndex) = ou.OrganizationUnitID
    Next
    
    If Not IsMissing(RetainSelection) Then
        If RetainSelection = True Then
            If Not (selOU Is Nothing) Then
                If (selOU.ParentOU Is Nothing) Or (selOU.ParentOU.OrganizationUnitID <= 0) Then
                    cboParentOU.Text = company.CompanyName
                Else
                    cboParentOU.Text = selOU.ParentOU.OrganizationUnitName
                End If
            End If
        End If
    End If
End Sub

Private Function SaveChanges() As Boolean
    On Error GoTo ErrorHandler
    
    SaveChanges = False
    If OUnits.ValidateKids() = True Then
        'MsgBox "Validation Succeeded", vbInformation, "Prototype"
        Dim i As Integer
        Dim isd As Boolean
        Dim txt As String
        i = 1
        While i <= OUnits.count
        txt = OUnits.Item(i).OrganizationUnitCode
            If Len(txt) > 2 Then
                If UCase(Mid(txt, 1, 3)) = "NEW" Then
                     OUnits.Item(i).OrganizationUnitCode = txtOUCode.Text
                     OUnits.Item(i).OrganizationUnitName = txtOUName.Text
                     OUnits.Item(i).IsModified = True
                End If
            End If
         i = i + 1
        Wend
       
        If OUnits.UpdateChanges() Then
            'reload data from db
            LoadOrganizationUnitsEx
            MsgBox "The Organization Units have been Updated Successfully", vbInformation, "Prototype"
            SaveChanges = True
        End If
    Else
        MsgBox "Validation Failed", vbInformation, TITLES
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred while updating the Organization Units" & vbNewLine & err.Description, vbExclamation, TITLES
    SaveChanges = False
End Function



Private Sub tvwOUnits_OLEStartDrag(data As MSComctlLib.DataObject, AllowedEffects As Long)
    data.SetData selOU
    AllowedEffects = vbDropEffectMove
End Sub

Private Sub txtEMail_Change()
    If Not ChangedFromCode Then
        If Not (selOU Is Nothing) Then
            selOU.EMail = txtEmail.Text
            selOU.IsModified = True
        End If
    End If
End Sub

Private Sub txtOUCode_Change()
    If Not ChangedFromCode Then
        If Not (selOU Is Nothing) Then
            selOU.OrganizationUnitCode = txtOUCode.Text
            selOU.IsModified = True
        End If
    End If
End Sub

Private Sub txtOUFunction_Change()
    If Not ChangedFromCode Then
        If Not (selOU Is Nothing) Then
            selOU.OrganizationUnitFunction = txtOUFunction.Text
            selOU.IsModified = True
        End If
    End If
End Sub

Private Sub txtOUName_Change()
    If Not ChangedFromCode Then
        If Not (selOU Is Nothing) Then
            selOU.OrganizationUnitName = txtOUName.Text
            selOU.IsModified = True
            Me.tvwOUnits.SelectedItem.Text = selOU.OrganizationUnitName
           
            'refresh the combobox to reflect the changes
            RefreshParentCombo True
        End If
    End If
End Sub

Private Sub txtTelephone_Change()
    If Not ChangedFromCode Then
        If Not (selOU Is Nothing) Then
            selOU.Telephone = txtTelephone.Text
            selOU.IsModified = True
        End If
    End If
End Sub


Private Sub LoadOrganizationUnitsEx()
    Dim myOU As HRCORE.OrganizationUnit
    Dim myNode As Node
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    'clear the old OUs
    OldOUs.Clear
    
    'clear lists
    tvwOUnits.Nodes.Clear
    cboParentOU.Clear
    
    'get all the organization units
    OUnits.GetAllOrganizationUnits
    
     'first add the company
    cboParentOU.AddItem company.CompanyName
    
    
    Set myNode = tvwOUnits.Nodes.add(, , "HRCOMPANY", company.CompanyName)
    myNode.Tag = "HRCOMPANY"
    myNode.Bold = True
    
    'now get the Top Level Organization Units
    Set TopLevelOUnits = OUnits.GetOrganizationUnitsOfTopmostLevel()
    
    If Not (TopLevelOUnits Is Nothing) Then
        For i = 1 To TopLevelOUnits.count
            Set myOU = TopLevelOUnits.Item(i)
            
             'Force the OU Type details to be loaded
            If Not (myOU.OUType Is Nothing) Then
                Set myOU.OUType = outs.FindOUType(myOU.OUType.OUTypeID)
            End If
            
            'populate the collection to hold initial OUs
            OldOUs.add myOU
            
            cboParentOU.AddItem myOU.OrganizationUnitName
            cboParentOU.ItemData(cboParentOU.NewIndex) = myOU.OrganizationUnitID
            
            'add the OU
            Set myNode = Me.tvwOUnits.Nodes.add(, , "OU:" & myOU.OrganizationUnitID, myOU.OrganizationUnitName)
            myNode.Tag = myOU.OrganizationUnitID
            myNode.EnsureVisible
            
            'now recursively add the children
            AddChildOUsRecursively myOU
        Next i
         
    End If
    
    'The COmmented Code below was used by Oscar to fix some logical Bug
    'Date: 14th May 2007
'    For i = 1 To OldOUs.count
'        Debug.Print OldOUs.Item(i).OrganizationUnitName
'    Next i
    
    'select an item
    If Me.tvwOUnits.Nodes.count > 1 Then
        Me.tvwOUnits.Nodes(2).Selected = True
        Set selOU = OUnits.FindOrganizationUnit(CLng(Me.tvwOUnits.Nodes(2).Tag))
        SetOUFields selOU
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "An error has occurred while populating the Organization Units" & _
    vbNewLine & err.Description, vbInformation, TITLES
End Sub


Private Sub AddChildOUsRecursively(ByVal theOU As HRCORE.OrganizationUnit)
    
    'this is a recursive function that populates child ous
    Dim ChildNode As Node
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    If Not (theOU Is Nothing) Then
        For i = 1 To theOU.Children.count
            'If Restoring Old OUs, use OUnits, Otherwise use OldOus
            
            If RestoringOldOUnits = True Then
                'rebuild OUnits
                'This assumes OUnits has been Clear i.e. OUnits.Count=0
                OUnits.add theOU.Children.Item(i)
            Else
                'populate the collection to hold initial OUs
                OldOUs.add theOU.Children.Item(i)
            End If
            
            Set ChildNode = tvwOUnits.Nodes.add("OU:" & theOU.OrganizationUnitID, tvwChild, "OU:" & theOU.Children.Item(i).OrganizationUnitID, theOU.Children.Item(i).OrganizationUnitName)
            ChildNode.Tag = theOU.Children.Item(i).OrganizationUnitID
            ChildNode.EnsureVisible
            'recursively load the children
            AddChildOUsRecursively theOU.Children.Item(i)
        Next i
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred" & vbNewLine & err.Description, vbInformation, TITLES
        
End Sub


