VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSectors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sectors"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab sstSectors 
      Height          =   5940
      Left            =   30
      TabIndex        =   24
      Top             =   0
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   10478
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Fund Codes"
      TabPicture(0)   =   "frmSectors.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdCloseFundCode"
      Tab(0).Control(1)=   "cmdDeleteFundCode"
      Tab(0).Control(2)=   "cmdEditFundCode"
      Tab(0).Control(3)=   "cmdNewFundCode"
      Tab(0).Control(4)=   "fraExistingFundCodes"
      Tab(0).Control(5)=   "fraFundCode"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Sectors"
      TabPicture(1)   =   "frmSectors.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdProjects"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "cmdNew"
      Tab(1).Control(3)=   "cmdEdit"
      Tab(1).Control(4)=   "cmdDelete"
      Tab(1).Control(5)=   "cmdClose"
      Tab(1).Control(6)=   "fraSector"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Projects"
      TabPicture(2)   =   "frmSectors.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label8"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdBack"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdDeleteP"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdEditP"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdNewP"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "fraExisting"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "txtSectorP"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "fraProject"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).ControlCount=   8
      Begin VB.CommandButton cmdCloseFundCode 
         Caption         =   "Close"
         Height          =   495
         Left            =   -69000
         TabIndex        =   6
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CommandButton cmdDeleteFundCode 
         Caption         =   "Delete"
         Height          =   495
         Left            =   -71880
         TabIndex        =   5
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CommandButton cmdEditFundCode 
         Caption         =   "Edit"
         Height          =   495
         Left            =   -73320
         TabIndex        =   4
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CommandButton cmdNewFundCode 
         Caption         =   "New"
         Height          =   495
         Left            =   -74850
         TabIndex        =   3
         Top             =   5160
         Width           =   1335
      End
      Begin VB.Frame fraExistingFundCodes 
         Caption         =   "Existing Fund Codes:"
         Height          =   3165
         Left            =   -74880
         TabIndex        =   36
         Top             =   1800
         Width           =   7335
         Begin MSComctlLib.ListView lvwFundCodes 
            Height          =   2685
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   4736
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
               Text            =   "Fund Code Name"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Other Details"
               Object.Width           =   7056
            EndProperty
         End
      End
      Begin VB.Frame fraFundCode 
         Height          =   1110
         Left            =   -74880
         TabIndex        =   35
         Top             =   450
         Width           =   7335
         Begin VB.TextBox txtOtherDetails 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1515
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   645
            Width           =   5730
         End
         Begin VB.TextBox txtFundCodeName 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1515
            Locked          =   -1  'True
            TabIndex        =   0
            Top             =   240
            Width           =   3735
         End
         Begin VB.Label Label10 
            Caption         =   "Other Details:"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   660
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "Fund Code Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   255
            Width           =   1335
         End
      End
      Begin VB.Frame fraProject 
         Caption         =   "Project Details:"
         Enabled         =   0   'False
         Height          =   4140
         Left            =   3195
         TabIndex        =   31
         Top             =   1065
         Width           =   4290
         Begin VB.Frame Frame1 
            Caption         =   "Fund Codes:"
            Height          =   2940
            Left            =   75
            TabIndex        =   39
            Top             =   1200
            Width           =   4215
            Begin VB.CommandButton cmdEditProgFCode 
               Caption         =   "Edit"
               Height          =   315
               Left            =   2280
               TabIndex        =   47
               Top             =   1200
               Width           =   840
            End
            Begin VB.TextBox txtFundCodeContribution 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   46
               Top             =   750
               Width           =   2790
            End
            Begin VB.ComboBox cboFundCodes 
               Height          =   315
               Left            =   1200
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   45
               Top             =   300
               Width           =   2790
            End
            Begin VB.CommandButton cmdDeleteProgFCode 
               Caption         =   "Remove"
               Height          =   315
               Left            =   3240
               TabIndex        =   42
               Top             =   1200
               Width           =   840
            End
            Begin VB.CommandButton cmdNewProgFCode 
               Caption         =   "Add"
               Height          =   315
               Left            =   1200
               TabIndex        =   41
               Top             =   1200
               Width           =   990
            End
            Begin MSComctlLib.ListView lvwProgrammeFundCodes 
               Height          =   1215
               Left            =   75
               TabIndex        =   40
               Top             =   1650
               Width           =   3990
               _ExtentX        =   7038
               _ExtentY        =   2143
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   0   'False
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Fund Code Name"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "% Contribution"
                  Object.Width           =   2540
               EndProperty
            End
            Begin VB.Label Label12 
               Caption         =   "% Contribution:"
               Height          =   165
               Left            =   75
               TabIndex        =   44
               Top             =   750
               Width           =   1140
            End
            Begin VB.Label Label11 
               Caption         =   "Fund Codes:"
               Height          =   240
               Left            =   75
               TabIndex        =   43
               Top             =   300
               Width           =   1065
            End
         End
         Begin VB.TextBox txtProjectName 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1425
            TabIndex        =   18
            Top             =   750
            Width           =   2715
         End
         Begin VB.TextBox txtProjectNumber 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1425
            TabIndex        =   17
            Top             =   300
            Width           =   2715
         End
         Begin VB.Label Label7 
            Caption         =   "Project Name:"
            Height          =   165
            Left            =   150
            TabIndex        =   33
            Top             =   810
            Width           =   1065
         End
         Begin VB.Label Label5 
            Caption         =   "Project Number:"
            Height          =   165
            Left            =   150
            TabIndex        =   32
            Top             =   360
            Width           =   1290
         End
      End
      Begin VB.TextBox txtSectorP 
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
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   540
         Width           =   6390
      End
      Begin VB.Frame fraExisting 
         Caption         =   "Existing Projects:"
         Height          =   4140
         Left            =   120
         TabIndex        =   30
         Top             =   1065
         Width           =   3015
         Begin MSComctlLib.ListView lvwProjects 
            Height          =   3840
            Left            =   75
            TabIndex        =   19
            Top             =   225
            Width           =   2865
            _ExtentX        =   5054
            _ExtentY        =   6773
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
               Text            =   "Project Number"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Project Name"
               Object.Width           =   7056
            EndProperty
         End
      End
      Begin VB.CommandButton cmdNewP 
         Caption         =   "New"
         Height          =   465
         Left            =   195
         TabIndex        =   20
         Top             =   5340
         Width           =   1290
      End
      Begin VB.CommandButton cmdEditP 
         Caption         =   "Edit"
         Height          =   465
         Left            =   1695
         TabIndex        =   21
         Top             =   5340
         Width           =   1290
      End
      Begin VB.CommandButton cmdDeleteP 
         Caption         =   "Delete"
         Height          =   465
         Left            =   3195
         TabIndex        =   22
         Top             =   5340
         Width           =   1290
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back To Sectors"
         Height          =   465
         Left            =   5820
         TabIndex        =   23
         Top             =   5340
         Width           =   1590
      End
      Begin VB.Frame fraSector 
         Enabled         =   0   'False
         Height          =   1785
         Left            =   -74880
         TabIndex        =   26
         Top             =   420
         Width           =   7290
         Begin VB.TextBox txtParentSector 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   1290
            Width           =   5775
         End
         Begin VB.TextBox txtDetails 
            Appearance      =   0  'Flat
            Height          =   540
            Left            =   1320
            MaxLength       =   150
            MultiLine       =   -1  'True
            TabIndex        =   8
            Top             =   645
            Width           =   5775
         End
         Begin VB.TextBox txtSectorName 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1320
            TabIndex        =   7
            Top             =   225
            Width           =   5775
         End
         Begin VB.Label Label3 
            Caption         =   "Sub-Sector Of:"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   1290
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Details:"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   765
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Sector Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   285
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   495
         Left            =   -68790
         TabIndex        =   15
         Top             =   5295
         Width           =   1185
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   495
         Left            =   -72150
         TabIndex        =   13
         Top             =   5295
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit Sectors"
         Height          =   495
         Left            =   -74865
         TabIndex        =   11
         Top             =   5295
         Width           =   1215
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -73530
         TabIndex        =   12
         Top             =   5295
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "Existing Sectors:"
         Height          =   2790
         Left            =   -74880
         TabIndex        =   25
         Top             =   2370
         Width           =   7290
         Begin MSComctlLib.TreeView tvwSectors 
            Height          =   2385
            Left            =   150
            TabIndex        =   10
            Top             =   300
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   4207
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   529
            LabelEdit       =   1
            Style           =   7
            Appearance      =   1
         End
      End
      Begin VB.CommandButton cmdProjects 
         Caption         =   "Projects..."
         Height          =   495
         Left            =   -70680
         TabIndex        =   14
         Top             =   5295
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Sector:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   270
         TabIndex        =   34
         Top             =   570
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmSectors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pFundCodes As HRCORE.FundCodes
Private selFundCode As HRCORE.FundCode

Private selSector As HRCORE.Sector
Private newSector As HRCORE.Sector
Private WithEvents pSectors As HRCORE.Sectors
Attribute pSectors.VB_VarHelpID = -1
Private OldSectors As HRCORE.Sectors
Private TopLevelSectors As HRCORE.Sectors

Private RestoringOldSectors As Boolean
Private ChangedFromCode As Boolean


'========== PROGRAMMES ====
Private pProgs As HRCORE.Programmes
Private selProg As HRCORE.Programme
Private FilteredProgs As HRCORE.Programmes


'======= PROGRAMME FUNDING =======
Private pProgFundings As HRCORE.ProgrammeFundings
Private TempProgFundings As HRCORE.ProgrammeFundings
Private selProgFunding As HRCORE.ProgrammeFunding



Private Sub cmdBack_Click()
    ResetProgrammeFundingCommandButtons
    sstSectors.TabEnabled(1) = True
    sstSectors.Tab = 1
    sstSectors.TabVisible(2) = False
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdCloseFundCode_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim resp As Long
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    
    Select Case LCase(cmdDelete.Caption)
        Case "delete"
            If selSector Is Nothing Then
                MsgBox "Select the Sector that you want to delete", vbInformation, TITLES
            Else
                resp = MsgBox("Are you sure you want to delete the selected sector:" & vbNewLine & UCase(selSector.SectorName), vbQuestion + vbYesNo, TITLES)
                If resp = vbYes Then
                    retVal = selSector.Delete()
                    Set selSector = Nothing
                    
                    'reload
                    LoadSectorsEx
                End If
            End If
                
        Case "cancel"
            RestoreOldSectors
            cmdEdit.Caption = "Edit Sectors"
            cmdNew.Enabled = False
            cmdDelete.Caption = "Delete"
            cmdProjects.Enabled = True
            fraSector.Enabled = False
    End Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has ocurred while performing the requested operation" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub cmdDeleteFundCode_Click()
    Dim resp As Long
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    
    Select Case LCase(Me.cmdDeleteFundCode.Caption)
        Case "delete"
            If selFundCode Is Nothing Then
                MsgBox "select the Fund Code you want to delete", vbInformation, TITLES
                Exit Sub
            Else
                resp = MsgBox("Are you sure you want to delete the selected Fund Code?", vbQuestion + vbYesNo, TITLES)
                If resp = vbYes Then
                    retVal = selFundCode.Delete()
                    Set selFundCode = Nothing
                    LoadFundCodes
                End If
            End If
                
        Case "cancel"
            Me.cmdNewFundCode.Enabled = True
            Me.cmdEditFundCode.Caption = "Edit"
            Me.cmdDeleteFundCode.Caption = "Delete"
            LockUnlockControls True
            PopulateFundCodes pFundCodes
    End Select
    
    Exit Sub
    
ErrorHandler:
End Sub

Private Sub cmdDeleteP_Click()
    Dim myinternalProgFunding As HRCORE.ProgrammeFunding
    Dim resp As Long
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    
    ResetProgrammeFundingCommandButtons
    Select Case LCase(cmdDeleteP.Caption)
        Case "delete"
            If selProg Is Nothing Then
                MsgBox "Select the Project that you want to delete", vbInformation, TITLES
                Exit Sub
            Else
                resp = MsgBox("Are you sure you want to delete the Project: " & selProg.ProgrammeNumber, vbYesNo + vbQuestion, TITLES)
                If resp = vbYes Then
                    retVal = selProg.Delete()
                    For Each myinternalProgFunding In selProg.ProgrammeFundCodes
                        ProcessChildObject myinternalProgFunding, TempProgFundings.FindProgrammeFundingByID(myinternalProgFunding.ProgrammeFundingID)
                    Next
                    Set selProg = Nothing
                    'reload
                    LoadProgrammesOfSector selSect, True
                End If
            End If
            
        Case "cancel"
            cmdNewP.Enabled = True
            cmdEditP.Caption = "Edit"
            cmdDeleteP.Caption = "Delete"
            cmdBack.Enabled = True
            fraProject.Enabled = False
            fraExisting.Enabled = True
            'reload
            LoadProgrammesOfSector selSector, True
    End Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while performing the requested operation" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub cmdDeleteProgFCode_Click()
    On Error GoTo ErrorHandler
    
    Select Case UCase(Me.cmdDeleteProgFCode.Caption)
        Case "REMOVE"
            If selProgFunding Is Nothing Then
                MsgBox "Please select the programme funding entry you wish to delete", vbExclamation, TITLES
            Else
                If MsgBox("Please confirm your decision to delete the selected programme funding entry", vbExclamation + vbYesNo, TITLES) = vbYes Then
'                    selProgFunding.Delete
                    selProgFunding.Deleted = True
                    'REMOVE FROM THE COLLECTION
                    selProg.ProgrammeFundCodes.RemoveByID selProgFunding.ProgrammeFundingID
                    'NOW ADDING THE UPDATED PROGRAMMEFUNDING OBJECT
                    selProg.ProgrammeFundCodes.add selProgFunding
                    'REFRESH THE PROGRAMME FUNDING COLLECTION
                    PopulateProgrammeFundings selProg.ProgrammeFundCodes
                End If
            End If
        Case "CANCEL"
            Me.cmdEditProgFCode.Caption = "Edit"
            Me.cmdDeleteProgFCode.Caption = "Remove"
            Me.cmdNewProgFCode.Enabled = True
            Me.cboFundCodes.Locked = True
            Me.txtFundCodeContribution.Locked = True
            PopulateProgrammeFundings selProg.ProgrammeFundCodes
    End Select
    Exit Sub
    
ErrorHandler:
    MsgBox "An Error has occurred while attempting to process the click event of the delete command button of the programme funding", vbExclamation, TITLES
End Sub

Private Sub cmdEditFundCode_Click()
    Select Case LCase(Me.cmdEditFundCode.Caption)
        Case "edit"
            If selFundCode Is Nothing Then
                MsgBox "First Select the Fund Code to edit", vbInformation, TITLES
                Exit Sub
            End If
            
            Me.cmdNewFundCode.Enabled = False
            Me.cmdEditFundCode.Caption = "Update"
            Me.cmdDeleteFundCode.Caption = "Cancel"
            LockUnlockControls False
            Me.txtFundCodeName.SetFocus
            
        Case "update"
            If UpdateFundCode() = False Then Exit Sub
            Me.cmdNewFundCode.Enabled = True
            Me.cmdEditFundCode.Caption = "Edit"
            Me.cmdDeleteFundCode.Caption = "Delete"
            LockUnlockControls True
            LoadFundCodes
            
        Case "cancel"
            Me.cmdNewFundCode.Caption = "New"
            Me.cmdEditFundCode.Caption = "Edit"
            Me.cmdDeleteFundCode.Enabled = True
            LockUnlockControls True
            
            PopulateFundCodes pFundCodes
    End Select
    
End Sub

Private Function InsertNewFundCode() As Boolean
    Dim newFC As HRCORE.FundCode
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    
    Set newFC = New HRCORE.FundCode
    If Len(Trim(txtFundCodeName.Text)) > 0 Then
        newFC.FundCodeName = Trim(txtFundCodeName.Text)
        newFC.OtherDetails = Trim(txtOtherDetails.Text)
        retVal = newFC.InsertNew()
        
        If retVal = 0 Then
            MsgBox "The New Fund Code has been added Successfully", vbInformation, TITLES
            InsertNewFundCode = True
        End If
    Else
        MsgBox "The FundCode Name is Required", vbExclamation, TITLES
        Me.txtFundCodeName.SetFocus
        Exit Function
    End If
    
    Exit Function
    
ErrorHandler:
    InsertNewFundCode = False
        
End Function


Private Function UpdateFundCode() As Boolean
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
   
    If Len(Trim(txtFundCodeName.Text)) > 0 Then
        selFundCode.FundCodeName = Trim(txtFundCodeName.Text)
        selFundCode.OtherDetails = Trim(txtOtherDetails.Text)
        retVal = selFundCode.Update()
        
        If retVal = 0 Then
            MsgBox "The Fund Code has been Updated Successfully", vbInformation, TITLES
            UpdateFundCode = True
        End If
    Else
        MsgBox "The FundCode Name is Required", vbExclamation, TITLES
        Me.txtFundCodeName.SetFocus
        Exit Function
    End If
    
    Exit Function
    
ErrorHandler:
    UpdateFundCode = False
        
End Function

Private Sub cmdEditP_Click()
    Dim myinternalProgFunding As HRCORE.ProgrammeFunding
    On Error GoTo ErrorHandler
    
    ResetProgrammeFundingCommandButtons
    Select Case LCase(cmdEditP.Caption)
        Case "edit"
            cmdNewP.Enabled = False
            cmdEditP.Caption = "Update"
            cmdDeleteP.Caption = "Cancel"
            cmdBack.Enabled = False
            fraProject.Enabled = True
            fraExisting.Enabled = False
            txtProjectNumber.SetFocus
            
        Case "update"
            If UpdateP() = False Then Exit Sub
            'NOW SAVING THE PROJECT FUNDING DETAILS
            For Each myinternalProgFunding In selProg.ProgrammeFundCodes
                Set myinternalProgFunding.Programme = selProg
                ProcessChildObject myinternalProgFunding, TempProgFundings.FindProgrammeFundingByID(myinternalProgFunding.ProgrammeFundingID)
            Next
            cmdNewP.Enabled = True
            cmdEditP.Caption = "Edit"
            cmdDeleteP.Caption = "Delete"
            cmdBack.Enabled = True
            fraProject.Enabled = False
            fraExisting.Enabled = True
            
            'reload
            LoadProgrammesOfSector selSector, True
            
        Case "cancel"
            cmdNewP.Caption = "New"
            cmdEditP.Caption = "Edit"
            cmdDeleteP.Enabled = True
            cmdBack.Enabled = True
            fraProject.Enabled = False
            fraExisting.Enabled = True
            're-load existing Programmes
            LoadProgrammesOfSector selSector, False
    End Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while performing the requested operation" & vbNewLine & err.Description, vbExclamation, TITLES
    
End Sub

Private Sub cmdEdit_Click()
    Select Case LCase(cmdEdit.Caption)
        Case "edit sectors"
            cmdEdit.Caption = "Update"
            cmdNew.Enabled = True
            cmdDelete.Caption = "Cancel"
            cmdProjects.Enabled = False
            fraSector.Enabled = True
            
        Case "update"
            If SaveChanges() = False Then Exit Sub
            cmdEdit.Caption = "Edit Sectors"
            cmdNew.Enabled = False
            cmdDelete.Caption = "Delete"
            cmdProjects.Enabled = True
            fraSector.Enabled = False
        
    End Select
    
End Sub

Private Sub cmdEditProgFCode_Click()
    Dim myinternalProgFunding As HRCORE.ProgrammeFunding
    On Error GoTo ErrorHandler
    
    Select Case UCase(Me.cmdEditProgFCode.Caption)
        Case "UPDATE"
            'ENTERING INTO THE COLLECTION
            'BUT FIRST ENSURING THAT THERE ARE NO DUPLICATES
            If Not ValidateProgrammeFundCodeEntry(True) Then Exit Sub
            Set selProgFunding.FundCode = pFundCodes.FindFundCodeByID(Me.cboFundCodes.ItemData(Me.cboFundCodes.ListIndex))
            selProgFunding.Deleted = False
            selProgFunding.PercentageFunding = CSng(Me.txtFundCodeContribution.Text)
            Set selProgFunding.Programme = selProg
            'REMOVING THE DUPLICATE FIRST
            selProg.ProgrammeFundCodes.RemoveByID selProgFunding.ProgrammeFundingID
            'NOW ENTERING THE UPDATED PROGRAMMEFUNDING OBJECT INTO THE COLLECTION
            selProg.ProgrammeFundCodes.add selProgFunding
            PopulateProgrammeFundings selProg.ProgrammeFundCodes
            Me.cmdEditProgFCode.Caption = "Edit"
            Me.cmdDeleteProgFCode.Caption = "Remove"
            Me.cmdNewProgFCode.Enabled = True
            Me.txtFundCodeContribution.Locked = True
            Me.cboFundCodes.Locked = True
        Case "EDIT"
            Me.cmdEditProgFCode.Caption = "Update"
            Me.cmdDeleteProgFCode.Caption = "Cancel"
            Me.cmdNewProgFCode.Enabled = False
            Me.txtFundCodeContribution.Locked = False
            Me.cboFundCodes.Locked = False
        Case "CANCEL"
            Me.cmdNewProgFCode.Caption = "Add"
            Me.cmdEditProgFCode.Caption = "Edit"
            Me.cmdDeleteProgFCode.Enabled = True
            Me.txtFundCodeContribution.Locked = True
            Me.cboFundCodes.Locked = True
            PopulateProgrammeFundings selProg.ProgrammeFundCodes
    End Select
    Exit Sub
    
ErrorHandler:
    MsgBox "An Error has occurred while attempting to process the click event of the edit command button of programme funding", vbExclamation, TITLES
End Sub

Private Sub cmdNew_Click()
    SetNewSectorDefaults
End Sub

Private Sub cmdNewFundCode_Click()
    Select Case LCase(cmdNewFundCode.Caption)
        Case "new"
            Me.cmdNewFundCode.Caption = "Update"
            Me.cmdEditFundCode.Caption = "Cancel"
            Me.cmdDeleteFundCode.Enabled = False
            ClearControlsFundCode
            LockUnlockControls False
            Me.txtFundCodeName.SetFocus
            
        Case "update"
            If InsertNewFundCode() = False Then Exit Sub
            Me.cmdNewFundCode.Caption = "New"
            Me.cmdEditFundCode.Caption = "Edit"
            Me.cmdDeleteFundCode.Enabled = True
            LockUnlockControls True
            
            LoadFundCodes
    End Select
End Sub

Private Sub ClearControlsFundCode()
    On Error GoTo ErrorHandler
    
    Me.txtFundCodeName.Text = ""
    Me.txtOtherDetails.Text = ""
       
    Exit Sub
    
ErrorHandler:
End Sub

Private Sub ResetProgrammeFundingCommandButtons()
    Me.cmdNewProgFCode.Caption = "Add"
    Me.cmdDeleteProgFCode.Caption = "Remove"
    Me.cmdEditProgFCode.Caption = "Edit"
End Sub

Private Sub LockUnlockControls(ByVal LockControls As Boolean)
    On Error GoTo ErrorHandler
    
    If LockControls Then
        Me.txtFundCodeName.Locked = True
        Me.txtOtherDetails.Locked = True
    Else
        Me.txtFundCodeName.Locked = False
        Me.txtOtherDetails.Locked = False
    End If
    
    Exit Sub
    
ErrorHandler:
End Sub

Private Sub cmdNewP_Click()
    Dim myinternalProgFunding As HRCORE.ProgrammeFunding
    
    ResetProgrammeFundingCommandButtons
    Select Case LCase(cmdNewP.Caption)
        Case "new"
            cmdNewP.Caption = "Update"
            cmdEditP.Caption = "Cancel"
            cmdDeleteP.Enabled = False
            cmdBack.Enabled = False
            ClearControlsP
            fraProject.Enabled = True
            Me.txtProjectNumber.SetFocus
            Set selProg = New HRCORE.Programme
            Set selProg.ProgrammeFundCodes = New HRCORE.ProgrammeFundings
            Set TempProgFundings = New HRCORE.ProgrammeFundings
            PopulateProgrammeFundings selProg.ProgrammeFundCodes
        Case "update"
            If InsertNewP() = False Then Exit Sub
            For Each myinternalProgFunding In selProg.ProgrammeFundCodes
                Set myinternalProgFunding.Programme = selProg
                ProcessChildObject myinternalProgFunding, TempProgFundings.FindProgrammeFundingByID(myinternalProgFunding.ProgrammeFundingID)
            Next
            cmdNewP.Caption = "New"
            cmdEditP.Caption = "Edit"
            cmdDeleteP.Enabled = True
            cmdBack.Enabled = True
            fraProject.Enabled = False
            'refresh the Programmes of the Selected Sector
            LoadProgrammesOfSector selSector, True
    End Select
End Sub

Private Sub ProcessChildObject(TheUpdatedObject As ProgrammeFunding, ThePrevObject As ProgrammeFunding)
    Dim blnEdited As Boolean
    On Error GoTo ErrorHandler
    
    blnEdited = True
    'NOW COMPARING THE PROPERTY VALUES
    If ThePrevObject Is Nothing Then GoTo EnterRecordInDB
    If TheUpdatedObject.FundCode.FundCodeID <> ThePrevObject.FundCode.FundCodeID Then
    Else
        If TheUpdatedObject.PercentageFunding <> ThePrevObject.PercentageFunding Then
        Else
            If TheUpdatedObject.Programme.ProgrammeID <> ThePrevObject.Programme.ProgrammeID Then
            Else
                blnEdited = False
            End If
        End If
    End If
EnterRecordInDB:
    'UPDATING THE DATABASE WITH THE LATEST INFORMATION
    If blnEdited = True Then
        TheUpdatedObject.InsertNew
        'NOW UPDATING THE GLOBAL COLLECTION
        pProgFundings.RemoveByID TheUpdatedObject.ProgrammeFundingID
        pProgFundings.add TheUpdatedObject
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "An Error has occurred while attempting to process the programme funding record" & vbCrLf & err.Description, vbExclamation, TITLES
End Sub

Private Sub ClearControlsP()
    
    Me.txtProjectName.Text = ""
    Me.txtProjectNumber.Text = ""
End Sub

Private Sub cmdNewProgFCode_Click()
    Dim myinternalProgFunding As HRCORE.ProgrammeFunding
    On Error GoTo ErrorHandler
    
    Select Case UCase(cmdNewProgFCode.Caption)
        Case "ADD"
            'CLEARING THE USER INPUT CONTROLS
            Me.cmdNewProgFCode.Caption = "Update"
            Me.cmdEditProgFCode.Caption = "Cancel"
            Me.cmdDeleteProgFCode.Enabled = False
            Me.txtFundCodeContribution.Text = ""
            Me.cboFundCodes.SetFocus
            Me.cboFundCodes.Locked = False
            Me.txtFundCodeContribution.Locked = False
            Set myinternalProgFunding = New HRCORE.ProgrammeFunding
        Case "UPDATE"
            'VALIDATING THE USER INPUT
            If Not ValidateProgrammeFundCodeEntry(False) Then Exit Sub
            Set myinternalProgFunding = New HRCORE.ProgrammeFunding
            Set myinternalProgFunding.FundCode = pFundCodes.FindFundCodeByID(Me.cboFundCodes.ItemData(Me.cboFundCodes.ListIndex))
'            Set myinternalProgFunding = selProg.ProgrammeFundCodes.GetProgrammeFundingByFundCodeIDAndProgrammeID(Me.cboFundCodes.ItemData(Me.cboFundCodes.ListIndex), selProg.ProgrammeID)
            Set myinternalProgFunding.Programme = selProg
            myinternalProgFunding.PercentageFunding = CSng(Me.txtFundCodeContribution.Text)
            myinternalProgFunding.Deleted = False
            selProg.ProgrammeFundCodes.add myinternalProgFunding
            Me.cmdNewProgFCode.Caption = "Add"
            Me.cmdEditProgFCode.Caption = "Edit"
            Me.cmdDeleteProgFCode.Enabled = True
            Me.txtFundCodeContribution.Text = vbNullString
            PopulateProgrammeFundings selProg.ProgrammeFundCodes
    End Select
    Exit Sub
    
ErrorHandler:
    MsgBox "An Error has occurred while attempting to process the click event of the add command button of the programme funding", vbExclamation, TITLES
End Sub

Private Function ValidateProgrammeFundCodeEntry(ByVal blnEditMode As Boolean) As Boolean
    Dim myinternalProgrammeFunding As ProgrammeFunding
    Dim lngLoopVariable As Long
    Dim sngTotalPercentage As Single
    On Error GoTo ErrorHandler
    
    ValidateProgrammeFundCodeEntry = False
    If Me.txtFundCodeContribution.Text = vbNullString Or Not IsNumeric(Me.txtFundCodeContribution.Text) Then
        MsgBox "The user input for the programme fundcode contribution percentage is not valid", vbExclamation, TITLES
        GoTo Finish
    Else
        For lngLoopVariable = 1 To selProg.ProgrammeFundCodes.count
            Set myinternalProgrammeFunding = selProg.ProgrammeFundCodes.Item(lngLoopVariable)
            If myinternalProgrammeFunding.Deleted = False Then
                If myinternalProgrammeFunding.FundCode.FundCodeID = Me.cboFundCodes.ItemData(Me.cboFundCodes.ListIndex) Then
                    If Not blnEditMode Then
                        GoTo DuplicateEntry
                    Else
                        If myinternalProgrammeFunding.ProgrammeFundingID <> selProgFunding.ProgrammeFundingID Then
DuplicateEntry:
                            MsgBox "The entry you are attempting to update would create a duplicate in the database", vbExclamation, TITLES
                            GoTo Finish
                        Else
                            sngTotalPercentage = sngTotalPercentage + CSng(Me.txtFundCodeContribution.Text)
                        End If
                    End If
                Else
                    sngTotalPercentage = sngTotalPercentage + myinternalProgrammeFunding.PercentageFunding
                End If
            End If
        Next
    End If
    If Not blnEditMode Then
        sngTotalPercentage = sngTotalPercentage + CSng(Me.txtFundCodeContribution.Text)
    End If
    If sngTotalPercentage > 100 Then
        MsgBox "The sum total of the programme fundcode contribution percentage should not exceed 100", vbExclamation, TITLES
        GoTo Finish
    End If
    ValidateProgrammeFundCodeEntry = True
Finish:
    Exit Function
    
ErrorHandler:
    MsgBox "An Error has occurred while attempting to validate the user entry for programme fundcode contribution" & vbCrLf & err.Description, vbExclamation, TITLES
End Function
Private Sub cmdProjects_Click()

    On Error GoTo ErrorHandler
    
    If selSector Is Nothing Then
        MsgBox "Select the Sector for which to view or add projects to", vbInformation, TITLES
    Else
        'set the selected sector
        Me.txtSectorP.Text = selSector.SectorName
        
        'show the projects
        
        sstSectors.TabEnabled(1) = False
        sstSectors.TabVisible(2) = True
        sstSectors.Tab = 2
        
        'load the projects of the selected sector
        LoadProgrammesOfSector selSector, True
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while loading the Projects of the selected sector" & vbNewLine & err.Description, TITLES
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    frmMain2.PositionTheFormWithoutEmpList Me
    
    Set pFundCodes = New HRCORE.FundCodes
    
    LoadFundCodes
    
    Set pProgFundings = New HRCORE.ProgrammeFundings
    pProgFundings.GetActiveProgrammeFundings
    
    Set pSectors = New HRCORE.Sectors
    Set OldSectors = New HRCORE.Sectors
    
    Set pProgs = New HRCORE.Programmes
    
    LoadSectorsEx
    
    sstSectors.Tab = 0
    sstSectors.TabVisible(2) = False
    
    
    'also populate the programmes
    pProgs.GetActiveProgrammes
    
    Exit Sub
    
ErrorHandler:
    
End Sub

Private Sub LoadFundCodes()
    On Error GoTo ErrorHandler
    
    pFundCodes.GetActiveFundCodes
    
    PopulateFundCodes pFundCodes
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while populating Fund Codes" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub PopulateFundCodes(ByVal TheFundCodes As HRCORE.FundCodes)
    Dim i As Long
    Dim ItemX As ListItem
    
    On Error GoTo ErrorHandler
    
    Me.lvwFundCodes.ListItems.Clear
    'also populate the Combo Box
    cboFundCodes.Clear
    
    If TheFundCodes Is Nothing Then
        Exit Sub
    End If
    
    For i = 1 To TheFundCodes.count
        Set ItemX = lvwFundCodes.ListItems.add(, , TheFundCodes.Item(i).FundCodeName)
        ItemX.SubItems(1) = TheFundCodes.Item(i).OtherDetails
        ItemX.Tag = TheFundCodes.Item(i).FundCodeID
        
        cboFundCodes.AddItem TheFundCodes.Item(i).FundCodeName
        cboFundCodes.ItemData(cboFundCodes.NewIndex) = TheFundCodes.Item(i).FundCodeID
    Next i
    
    If lvwFundCodes.ListItems.count > 0 Then
        Me.lvwFundCodes.ListItems.Item(1).Selected = True
        If IsNumeric(Me.lvwFundCodes.ListItems.Item(1).Tag) Then
            Set selFundCode = pFundCodes.FindFundCodeByID(CLng(lvwFundCodes.ListItems.Item(1).Tag))
        End If
        SetFieldsFundCode selFundCode
        
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while populating the Fund Codes" & vbNewLine & err.Description, vbExclamation, TITLES
        
End Sub


Private Sub SetFieldsFundCode(ByVal TheFundCode As HRCORE.FundCode)
    On Error GoTo ErrorHandler
    
    ClearControlsFundCode
    
    If TheFundCode Is Nothing Then Exit Sub
    Me.txtFundCodeName.Text = TheFundCode.FundCodeName
    Me.txtOtherDetails.Text = TheFundCode.OtherDetails
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while displaying details about the selected Fund Code" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub EnableDisableProgFundCodeButtons(ByVal EnableThem As Boolean)
    On Error GoTo ErrorHandler
    
    If EnableThem Then
        Me.cmdNewProgFCode.Enabled = True
        Me.cmdEditProgFCode.Enabled = True
        Me.cmdDeleteProgFCode.Enabled = True
    Else
        Me.cmdNewProgFCode.Enabled = False
        Me.cmdEditProgFCode.Enabled = False
        Me.cmdDeleteProgFCode.Enabled = False
    End If
    
    Exit Sub
    
ErrorHandler:
End Sub

Private Sub ClearControls()
    Dim ctl As Control
    
    On Error GoTo ErrorHandler
    
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        End If
    Next ctl
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Could not clear controls" & vbNewLine & err.Description, vbExclamation, TITLES
    
End Sub

Private Sub LoadSectors()
    On Error GoTo ErrorHandler
    
    pSectors.GetActiveSectors
    
    PopulateSectors pSectors
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Could not load the existing Sectors" & vbNewLine & err.Description, vbExclamation, TITLES
    
End Sub

Private Sub PopulateSectors(ByVal TheSectors As HRCORE.Sectors)
    Dim i As Long
    Dim nodeX As Node
    
    On Error GoTo ErrorHandler
    tvwSectors.Nodes.Clear
    
    If Not (TheSectors Is Nothing) Then
        For i = 1 To TheSectors.count
            Set nodeX = tvwSectors.Nodes.add(, , "S" & TheSectors.Item(i).SectorID, TheSectors.Item(i).SectorName)
            nodeX.Tag = TheSectors.Item(i).SectorID
        Next i
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while Populating the Sectors" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub LoadSectorsEx()
    Dim mySector As HRCORE.Sector
    Dim myNode As Node
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    'clear the old Sectors
    OldSectors.Clear
    
    'clear lists
    tvwSectors.Nodes.Clear
        
    'get all the Sectors
    pSectors.GetActiveSectors
    
    'first load the Header
    Set myNode = tvwSectors.Nodes.add(, , "CSSSSECTORS", "Sectors")
    myNode.Tag = "CSSSSECTORS"
    myNode.Bold = True
    
    'now get the Top Level Sectors
    Set TopLevelSectors = pSectors.GetTopLevelSectors()
    If Not (TopLevelSectors Is Nothing) Then
        For i = 1 To TopLevelSectors.count
            Set mySector = TopLevelSectors.Item(i)
                                    
            'populate the collection to hold initial Sectors
            OldSectors.add mySector
            
           
            'add the Sector
            Set myNode = Me.tvwSectors.Nodes.add(, , "SEC:" & mySector.SectorID, mySector.SectorName)
            myNode.Tag = mySector.SectorID
            myNode.EnsureVisible
            
            'now recursively add the children
            AddChildSectorsRecursively mySector
        Next i
         
    End If
       
    'select an item
    If Me.tvwSectors.Nodes.count > 1 Then
        Me.tvwSectors.Nodes(2).Selected = True
        Set selSector = pSectors.FindSectorByID(CLng(Me.tvwSectors.Nodes(2).Tag))
        SetFields selSector
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "An error has occurred while populating the Sectors" & _
    vbNewLine & err.Description, vbInformation, TITLES
End Sub


Private Sub RestoreOldSectors()
    Dim mySector As HRCORE.Sector
    Dim myNode As Node
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    'clear the pSectors
    pSectors.Clear
    
    'clear lists
    tvwSectors.Nodes.Clear
     
        
    'now get the Top Level Sectors
    Set TopLevelSectors = OldSectors.GetTopLevelSectors()
    If Not (TopLevelSectors Is Nothing) Then
        For i = 1 To TopLevelSectors.count
            Set mySector = TopLevelSectors.Item(i)
                                    
            'repopulate the collection to hold Sectors
            pSectors.add mySector
           
            'add the OU
            Set myNode = Me.tvwSectors.Nodes.add(, , "SEC:" & mySector.SectorID, mySector.SectorName)
            myNode.Tag = mySector.SectorID
            myNode.EnsureVisible
            
            RestoringOldSectors = True
            'now recursively add the children
            AddChildSectorsRecursively mySector
            RestoringOldSectors = False
        Next i
         
    End If
       
    'select an item
    If Me.tvwSectors.Nodes.count > 1 Then
        Me.tvwSectors.Nodes(1).Selected = True
        Set selSector = pSectors.FindSectorByID(CLng(Me.tvwSectors.Nodes(1).Tag))
        SetFields selSector
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "An error has occurred while populating the Sectors" & _
    vbNewLine & err.Description, vbInformation, TITLES
End Sub



Private Sub AddChildSectorsRecursively(ByVal TheSector As HRCORE.Sector)
    
    'this is a recursive function that populates child Sectors
    Dim ChildNode As Node
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    If Not (TheSector Is Nothing) Then
        For i = 1 To TheSector.Children.count
            'If Restoring Old Sectors, use pSectors, Otherwise use OldSectors
            
            If RestoringOldSectors = True Then
                'rebuild pSectors
                'This assumes pSectors has been Clear i.e. pSectors.Count=0
                pSectors.add TheSector.Children.Item(i)
            Else
                'populate the collection to hold initial Sectors
                OldSectors.add TheSector.Children.Item(i)
            End If
            
            Set ChildNode = tvwSectors.Nodes.add("SEC:" & TheSector.SectorID, tvwChild, "SEC:" & TheSector.Children.Item(i).SectorID, TheSector.Children.Item(i).SectorName)
            ChildNode.Tag = TheSector.Children.Item(i).SectorID
            ChildNode.EnsureVisible
            'recursively load the children
            AddChildSectorsRecursively TheSector.Children.Item(i)
        Next i
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred" & vbNewLine & err.Description, vbInformation, TITLES
        
End Sub

Private Sub SetFields(ByVal TheSector As HRCORE.Sector)
    On Error GoTo ErrorHandler
    
    ChangedFromCode = True
    ClearControls
    If Not (TheSector Is Nothing) Then
        Me.txtDetails.Text = TheSector.Details
        Me.txtSectorName.Text = TheSector.SectorName
        If Not (TheSector.ParentSector Is Nothing) Then
            Me.txtParentSector.Text = TheSector.ParentSector.SectorName
        End If
    End If
    
    ChangedFromCode = False
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while processing the selected sector" & vbNewLine & err.Description, vbExclamation, TITLES
    ChangedFromCode = False
        
End Sub

Private Sub lvwFundCodes_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo ErrorHandler
    Set selFundCode = Nothing
    If IsNumeric(Item.Tag) Then
        Set selFundCode = pFundCodes.FindFundCodeByID(CLng(Item.Tag))
    End If
    
    SetFieldsFundCode selFundCode
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while processing the selected Fund Code" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub lvwProgrammeFundCodes_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim lngLoopVariable As Long
    On Error GoTo ErrorHandler
    
    For lngLoopVariable = 0 To Me.cboFundCodes.ListCount
        If UCase(Me.cboFundCodes.List(lngLoopVariable)) = UCase(Item.Text) Then
            Me.cboFundCodes.ListIndex = lngLoopVariable
        End If
    Next
    Me.txtFundCodeContribution.Text = Item.SubItems(1)
    Set selProgFunding = selProg.ProgrammeFundCodes.FindProgrammeFundingByID(Item.Tag)
    Exit Sub
    
ErrorHandler:
    MsgBox "An Error has occurred while attempting to process the click event of the list view for programme funding" & vbNewLine, vbExclamation, TITLES
End Sub

Private Sub lvwProjects_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error GoTo ErrorHandler
    
    lvwProjects.SortKey = ColumnHeader.Index - 1
    lvwProjects.SortOrder = lvwAscending
    lvwProjects.Sorted = True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Could not sort the Projects by: " & ColumnHeader.Text, vbExclamation, TITLES
End Sub

Private Sub lvwProjects_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo ErrorHandler
    
    Set selProg = Nothing
    If IsNumeric(Item.Tag) Then
        Set selProg = pProgs.FindProgrammeByID(CLng(Item.Tag))
        'NOW LOADING THE PROGRAMME FUNDINGS
        Set TempProgFundings = pProgFundings.GetProgrammeFundingsByProgrammeID(CLng(Item.Tag))
        'POPULATE THE PROGRAMME FUNDINGS
        PopulateProgrammeFundings selProg.ProgrammeFundCodes
    End If
    
    SetFieldsP selProg
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while populating information about the selected Project" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub PopulateProgrammeFundings(TheProgFundings As HRCORE.ProgrammeFundings)
    Dim myinternalFundCode As HRCORE.FundCode
    Dim myListItem As ListItem
    Dim lngLoopVariable As Long
    On Error GoTo ErrorHandler
    
    Me.txtFundCodeContribution.Text = vbNullString
    Me.lvwProgrammeFundCodes.ListItems.Clear
    If Not TheProgFundings Is Nothing Then
        For lngLoopVariable = 1 To TheProgFundings.count
            Set myinternalFundCode = pFundCodes.FindFundCodeByID(TheProgFundings.Item(lngLoopVariable).FundCode.FundCodeID)
            If TheProgFundings.Item(lngLoopVariable).Deleted = False Then
                Set myListItem = Me.lvwProgrammeFundCodes.ListItems.add(, , myinternalFundCode.FundCodeName)
                myListItem.SubItems(1) = TheProgFundings.Item(lngLoopVariable).PercentageFunding
                myListItem.Tag = TheProgFundings.Item(lngLoopVariable).ProgrammeFundingID
            End If
        Next
        If Me.lvwProgrammeFundCodes.ListItems.count > 0 Then
            lvwProgrammeFundCodes_ItemClick Me.lvwProgrammeFundCodes.ListItems.Item(1)
        End If
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "An Error has occurred while attempting to load all the programme fundings for a specific programme" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub tvwSectors_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error GoTo ErrorHandler
    
    Set selSector = Nothing
    If IsNumeric(Node.Tag) Then
        Set selSector = pSectors.FindSectorByID(CLng(Node.Tag))
    End If
    
    SetFields selSector
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while processing the selected Sector" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub


Private Sub SetNewSectorDefaults()
    Dim newCount As Long
    Dim NewID As Long
    Dim newParent As HRCORE.Sector
    Dim newNode As Node
    
    On Error GoTo ErrorHandler
    
    Set newSector = New HRCORE.Sector
    If Not (pSectors Is Nothing) Then
        newCount = pSectors.count + 1
        NewID = pSectors.GetNextKidID()
    Else
        newCount = 1
        NewID = 1
    End If
    
    With newSector
        .SectorName = "New Sector " & newCount
        .SectorID = NewID
        .InsertionOrderNo = NewID
        If Not (selSector Is Nothing) Then
            Set .ParentSector = selSector
            Set newNode = tvwSectors.Nodes.add("SEC:" & .ParentSector.SectorID, tvwChild, "SEC:" & CStr(.SectorID), newSector.SectorName)
            newNode.Selected = True
            newNode.Tag = NewID
            newNode.EnsureVisible
        Else
            'if no Node is selected, always add the new Sector to the CSSS Sector
            Set .ParentSector = Nothing
            Set newNode = tvwSectors.Nodes.add("CSSSSECTORS", tvwChild, "SEC:" & CStr(newSector.SectorID), newSector.SectorName)
            newNode.Tag = NewID
            newNode.Selected = True
            newNode.EnsureVisible
        End If
        
        newSector.IsNewEntity = True
    End With
    
    
    If Not (pSectors Is Nothing) Then
        pSectors.add newSector
    Else
        Set pSectors = New HRCORE.Sectors
        pSectors.add newSector
    End If
    
    'ClearControls
    SetFields newSector
               
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while creating a new sector" & vbNewLine & err.Description, vbExclamation, TITLES
    
End Sub

Private Sub txtDetails_Change()
    If Not ChangedFromCode Then
        If Not (selSector Is Nothing) Then
            selSector.Details = Trim(txtDetails.Text)
            selSector.IsModified = True
        End If
    End If
End Sub

Private Sub txtSectorName_Change()
    If Not ChangedFromCode Then
        If Not (selSector Is Nothing) Then
            selSector.SectorName = Trim(txtSectorName.Text)
            tvwSectors.SelectedItem.Text = Trim(txtSectorName.Text)
            selSector.IsModified = True
        End If
    End If
End Sub


Private Function SaveChanges() As Boolean
    On Error GoTo ErrorHandler
    
    SaveChanges = False
    If pSectors.ValidateKids() = True Then
        'MsgBox "Validation Succeeded", vbInformation, "Prototype"
        If pSectors.UpdateChanges() Then
            'reload data from db
            LoadSectorsEx
            MsgBox "The Sectors have been Updated Successfully", vbInformation, TITLES
            SaveChanges = True
        End If
    Else
        MsgBox "Validation Failed", vbInformation, TITLES
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred while updating the Sectors" & vbNewLine & err.Description, vbExclamation, TITLES
    SaveChanges = False
    
End Function


Private Sub LoadProgrammesOfSector(ByVal TheSector As HRCORE.Sector, ByVal Refresh As Boolean)
    Dim i As Long
    Dim ItemX As ListItem
    
    On Error GoTo ErrorHandler
    
    'First Clear the Listview
    Me.lvwProjects.ListItems.Clear
    
    If Refresh = True Then
        pProgs.GetActiveProgrammes
    End If
    
    If Not (TheSector Is Nothing) Then
        Set FilteredProgs = pProgs.GetProgrammesOfSectorID(TheSector.SectorID)
        If Not (FilteredProgs Is Nothing) Then
            For i = 1 To FilteredProgs.count
                Set ItemX = lvwProjects.ListItems.add(, , FilteredProgs.Item(i).ProgrammeNumber)
                ItemX.SubItems(1) = FilteredProgs.Item(i).ProgrammeName
                ItemX.Tag = FilteredProgs.Item(i).ProgrammeID
            Next i
        End If
        If Me.lvwProjects.ListItems.count > 0 Then
            lvwProjects_ItemClick Me.lvwProjects.ListItems.Item(1)
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while Populating the Projects under the selected sector" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Function InsertNewP() As Boolean
'    Dim newProg As HRCORE.Programme
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    
'    Set newProg = New HRCORE.Programme
    
    If Trim(Me.txtProjectNumber.Text) <> "" Then
        selProg.ProgrammeNumber = Trim(Me.txtProjectNumber.Text)
    Else
        MsgBox "The Project Number is Required", vbInformation, TITLES
        Me.txtProjectNumber.SetFocus
        Exit Function
    End If
        
'    If Trim(Me.txtFundCode.Text) <> "" Then
'        newProg.FundCode = Trim(Me.txtFundCode.Text)
'    Else
'        MsgBox "The Fund Code is Required", vbInformation, TITLES
'        Me.txtFundCode.SetFocus
'        Exit Function
'    End If
    
    selProg.ProgrammeName = Trim(Me.txtProjectName.Text)
    If Not selSector Is Nothing Then
        Set selProg.Sector = selSector
    Else
        Set selProg = Nothing
    End If
    Dim k As Long
    k = selSector.SectorID
    
    retVal = selProg.InsertNew()
    If retVal = 0 Then
        MsgBox "The New Project was added successfully", vbInformation, TITLES
        InsertNewP = True
    Else
        InsertNewP = False
    End If
    
    Exit Function
ErrorHandler:
    MsgBox "An error occurred while creating a new project" & vbNewLine & err.Description, vbExclamation, TITLES
    InsertNewP = False
End Function

Private Sub SetFieldsP(ByVal TheProg As HRCORE.Programme)
    On Error GoTo ErrorHandler
    ClearControlsP
    
    If Not (TheProg Is Nothing) Then
        Me.txtProjectNumber.Text = TheProg.ProgrammeNumber
        Me.txtProjectName.Text = TheProg.ProgrammeName
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while populating the details about the selected Project" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Function UpdateP() As Boolean
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    
    If selProg Is Nothing Then
        MsgBox "There is no Selected Project", vbInformation, TITLES
        UpdateP = False
        Exit Function
    End If
    
    If Trim(Me.txtProjectNumber.Text) <> "" Then
        selProg.ProgrammeNumber = Trim(Me.txtProjectNumber.Text)
    Else
        MsgBox "The Project Number is Required", vbExclamation, TITLES
        txtProjectNumber.SetFocus
        Exit Function
    End If
    
'    If Trim(Me.txtFundCode.Text) <> "" Then
'        selProg.FundCode = Trim(txtFundCode.Text)
'    Else
'        MsgBox "The Fund Code is Required", vbExclamation, TITLES
'        Me.txtFundCode.SetFocus
'        Exit Function
'    End If
    
    
    selProg.ProgrammeName = Trim(txtProjectName.Text)
    
    Set selProg.Sector = selSector
    
    retVal = selProg.Update()
    If retVal = 0 Then
        MsgBox "The Project Details have been updated successfully", vbInformation, TITLES
    End If
    
    UpdateP = True
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred while updating the Project" & vbNewLine & err.Description, vbExclamation, TITLES
    UpdateP = False
    
    
End Function
