VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmGradeTitles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grade Titles"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5520
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab sstGrades 
      Height          =   6540
      Left            =   150
      TabIndex        =   34
      Top             =   600
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   11536
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Employee Grades"
      TabPicture(0)   =   "frmGradeTitles.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraGrade"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cboStaffCategory"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdNew"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdEdit"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdDelete"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdClose"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdTitles"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdSteps"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Grade Titles"
      TabPicture(1)   =   "frmGradeTitles.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAddDesig"
      Tab(1).Control(1)=   "cmdRemoveDesig"
      Tab(1).Control(2)=   "cmdBackGTitle"
      Tab(1).Control(3)=   "Frame1"
      Tab(1).Control(4)=   "cboDesignations"
      Tab(1).Control(5)=   "lvwPositions"
      Tab(1).Control(6)=   "Label3"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Northwise Progression"
      TabPicture(2)   =   "frmGradeTitles.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdBackS"
      Tab(2).Control(1)=   "cmdDeleteStep"
      Tab(2).Control(2)=   "cmdEditStep"
      Tab(2).Control(3)=   "cmdNewStep"
      Tab(2).Control(4)=   "Frame6"
      Tab(2).Control(5)=   "Frame5"
      Tab(2).Control(6)=   "Frame2"
      Tab(2).ControlCount=   7
      Begin VB.CommandButton cmdSteps 
         Caption         =   "Steps..."
         Height          =   465
         Left            =   5250
         TabIndex        =   14
         Top             =   5925
         Width           =   915
      End
      Begin VB.CommandButton cmdTitles 
         Caption         =   "Titles..."
         Height          =   465
         Left            =   4125
         TabIndex        =   13
         Top             =   5925
         Width           =   990
      End
      Begin VB.CommandButton cmdBackS 
         Caption         =   "Back To Grades"
         Height          =   465
         Left            =   -68850
         TabIndex        =   33
         Top             =   5925
         Width           =   1440
      End
      Begin VB.CommandButton cmdDeleteStep 
         Caption         =   "Delete"
         Height          =   465
         Left            =   -72075
         TabIndex        =   32
         Top             =   5925
         Width           =   1290
      End
      Begin VB.CommandButton cmdEditStep 
         Caption         =   "Edit"
         Height          =   465
         Left            =   -73275
         TabIndex        =   31
         Top             =   5925
         Width           =   1065
      End
      Begin VB.CommandButton cmdNewStep 
         Caption         =   "New"
         Height          =   465
         Left            =   -74775
         TabIndex        =   30
         Top             =   5925
         Width           =   1290
      End
      Begin VB.Frame Frame6 
         Caption         =   "Existing Steps:"
         Height          =   2925
         Left            =   -74850
         TabIndex        =   53
         Top             =   2910
         Width           =   7440
         Begin MSComctlLib.ListView lvwSteps 
            Height          =   2550
            Left            =   75
            TabIndex        =   29
            Top             =   225
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   4498
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Step"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Salary"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Grade"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Staff Category"
               Object.Width           =   3528
            EndProperty
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Step Details:"
         Height          =   765
         Left            =   -74850
         TabIndex        =   52
         Top             =   1935
         Width           =   7365
         Begin VB.TextBox txtSalary 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3750
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   300
            Width           =   1965
         End
         Begin VB.TextBox txtStep 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   750
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   300
            Width           =   990
         End
         Begin VB.Label Label15 
            Caption         =   "Salary:"
            Height          =   240
            Left            =   3165
            TabIndex        =   55
            Top             =   330
            Width           =   495
         End
         Begin VB.Label Label14 
            Caption         =   "Step:"
            Height          =   240
            Left            =   75
            TabIndex        =   54
            Top             =   337
            Width           =   540
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1425
         Left            =   -74850
         TabIndex        =   49
         Top             =   375
         Width           =   7440
         Begin VB.TextBox txtHighestSalaryS 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   5040
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   960
            Width           =   2175
         End
         Begin VB.TextBox txtLowestSalaryS 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox txtCategoryNameS 
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
            Height          =   285
            Left            =   1425
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   600
            Width           =   5790
         End
         Begin VB.TextBox txtStaffCategoryS 
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
            Height          =   315
            Left            =   1425
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   225
            Width           =   5790
         End
         Begin VB.Label Label17 
            Caption         =   "Highest Salary:"
            Height          =   255
            Left            =   3840
            TabIndex        =   57
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label16 
            Caption         =   "Lowest Salary:"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label13 
            Caption         =   "Employee Grade:"
            Height          =   240
            Left            =   75
            TabIndex        =   51
            Top             =   600
            Width           =   1590
         End
         Begin VB.Label Label12 
            Caption         =   "Staff Category:"
            Height          =   240
            Left            =   75
            TabIndex        =   50
            Top             =   225
            Width           =   1665
         End
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   465
         Left            =   6375
         TabIndex        =   15
         Top             =   5925
         Width           =   1140
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   465
         Left            =   2625
         TabIndex        =   12
         Top             =   5925
         Width           =   1140
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   465
         Left            =   1350
         TabIndex        =   11
         Top             =   5925
         Width           =   1065
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   465
         Left            =   150
         TabIndex        =   10
         Top             =   5925
         Width           =   1065
      End
      Begin VB.ComboBox cboStaffCategory 
         Height          =   315
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   525
         Width           =   5865
      End
      Begin VB.Frame Frame4 
         Caption         =   "Existing Grades:"
         Height          =   2940
         Left            =   75
         TabIndex        =   43
         Top             =   2850
         Width           =   7515
         Begin MSComctlLib.ListView lvwCSSSGrades 
            Height          =   2490
            Left            =   150
            TabIndex        =   9
            Top             =   300
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   4392
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Grade"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Staff Category"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Band S'ty (%)"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Band S'ty (Amount)"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame fraGrade 
         Enabled         =   0   'False
         Height          =   1740
         Left            =   75
         TabIndex        =   39
         Top             =   975
         Width           =   7515
         Begin VB.CommandButton cmdColorCode 
            Caption         =   "Color Code..."
            Height          =   390
            Left            =   5625
            TabIndex        =   6
            Top             =   712
            Width           =   1140
         End
         Begin VB.TextBox txtHighestSalary 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4200
            TabIndex        =   5
            Top             =   750
            Width           =   1140
         End
         Begin VB.TextBox txtLowestSalary 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1575
            TabIndex        =   4
            Top             =   750
            Width           =   1140
         End
         Begin VB.TextBox txtLevel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6600
            TabIndex        =   3
            Top             =   285
            Width           =   690
         End
         Begin VB.TextBox txtCategoryName 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1575
            TabIndex        =   2
            Top             =   300
            Width           =   3765
         End
         Begin VB.TextBox txtBandSensitivityAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5625
            TabIndex        =   8
            Top             =   1200
            Width           =   1665
         End
         Begin VB.TextBox txtBandSensitivityPercent 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1575
            TabIndex        =   7
            Top             =   1200
            Width           =   1365
         End
         Begin VB.Label lblColorCode 
            Caption         =   "Color"
            Height          =   315
            Left            =   6825
            TabIndex        =   47
            Top             =   750
            Width           =   465
         End
         Begin VB.Label Label11 
            Caption         =   "Highest Salary:"
            Height          =   240
            Left            =   3000
            TabIndex        =   46
            Top             =   780
            Width           =   1140
         End
         Begin VB.Label Label10 
            Caption         =   "Lowest Salary:"
            Height          =   240
            Left            =   75
            TabIndex        =   45
            Top             =   787
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "Grade Level:"
            Height          =   165
            Left            =   5550
            TabIndex        =   44
            Top             =   360
            Width           =   990
         End
         Begin VB.Label Label8 
            Caption         =   "Band Sensitivity (Amount):"
            Height          =   240
            Left            =   3675
            TabIndex        =   42
            Top             =   1230
            Width           =   1740
         End
         Begin VB.Label Label7 
            Caption         =   "Band Sensitivity (%):"
            Height          =   240
            Left            =   75
            TabIndex        =   41
            Top             =   1237
            Width           =   1440
         End
         Begin VB.Label Label6 
            Caption         =   "Employee Grade:"
            Height          =   240
            Left            =   75
            TabIndex        =   40
            Top             =   322
            Width           =   1440
         End
      End
      Begin VB.CommandButton cmdAddDesig 
         Caption         =   "Add"
         Height          =   495
         Left            =   -74850
         TabIndex        =   20
         Top             =   5895
         Width           =   1140
      End
      Begin VB.CommandButton cmdRemoveDesig 
         Caption         =   "Remove"
         Height          =   495
         Left            =   -73500
         TabIndex        =   21
         Top             =   5895
         Width           =   1140
      End
      Begin VB.CommandButton cmdBackGTitle 
         Caption         =   "Back To Grades"
         Height          =   495
         Left            =   -69225
         TabIndex        =   22
         Top             =   5895
         Width           =   1635
      End
      Begin VB.Frame Frame1 
         Height          =   1095
         Left            =   -74850
         TabIndex        =   35
         Top             =   450
         Width           =   7320
         Begin VB.TextBox txtCategoryName2 
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
            Height          =   315
            Left            =   1425
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   600
            Width           =   5565
         End
         Begin VB.TextBox txtStaffCategory 
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
            Height          =   315
            Left            =   1425
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   225
            Width           =   5565
         End
         Begin VB.Label Label1 
            Caption         =   "Employee Grade:"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   675
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Staff Category:"
            Height          =   240
            Left            =   150
            TabIndex        =   36
            Top             =   225
            Width           =   1140
         End
      End
      Begin VB.ComboBox cboDesignations 
         Height          =   315
         Left            =   -73545
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1830
         Width           =   6015
      End
      Begin MSComctlLib.ListView lvwPositions 
         Height          =   3330
         Left            =   -74850
         TabIndex        =   19
         Top             =   2430
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   5874
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Designation"
            Object.Width           =   12347
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "Staff Category:"
         Height          =   240
         Left            =   150
         TabIndex        =   48
         Top             =   600
         Width           =   1590
      End
      Begin VB.Label Label3 
         Caption         =   "Designation:"
         Height          =   240
         Left            =   -74820
         TabIndex        =   38
         Top             =   1830
         Width           =   1065
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Employee Grades"
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
      Top             =   90
      Width           =   3135
   End
End
Attribute VB_Name = "frmGradeTitles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pCSSCats As HRCORE.CSSSCategories
Private selCSSCat As HRCORE.CSSSCategory

Private pGrades As HRCORE.EmployeeCategories
Private selGrade As HRCORE.EmployeeCategory

Private pJobPositions As HRCORE.JobPositions
Private selJPos As HRCORE.JobPosition

Private pGradeTitles As HRCORE.GradeTitles
Private selGradeTitle As HRCORE.GradeTitle
Private SelGTitles As HRCORE.GradeTitles

Private pCatSteps As HRCORE.CategorySteps
Private selCatSteps As HRCORE.CategorySteps
Private selCatStep As HRCORE.CategoryStep

Private Sub cboDesignations_Click()
    On Error GoTo ErrorHandler
    Set selJPos = Nothing
    If cboDesignations.ListIndex > -1 Then
        Set selJPos = pJobPositions.FindJobPositionByID(cboDesignations.ItemData(cboDesignations.ListIndex))
    End If
    Exit Sub
    
ErrorHandler:
End Sub

Private Sub cboStaffCategory_Click()
    Set selCSSCat = Nothing
    If cboStaffCategory.ListIndex > -1 Then
        Set selCSSCat = pCSSCats.FindCSSSCategoryByID(cboStaffCategory.ItemData(cboStaffCategory.ListIndex))
    End If
    
    LoadGradesOfCSSSCategory selCSSCat, False
End Sub

Private Sub cmdAddDesig_Click()
    Dim newGradeTitle As HRCORE.GradeTitle
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    If selGrade Is Nothing Then
        MsgBox "First select a Grade", vbInformation, TITLES
        Exit Sub
    End If
    
    If selJPos Is Nothing Then
        MsgBox "Select the Job Position to add to this Grade", vbInformation, TITLES
        Exit Sub
    End If
    
    If SelGTitles.PositionExists(selJPos, selGrade) Then
        MsgBox "The Position already exists for the selected grade", vbInformation, TITLES
    Else
        Set newGradeTitle = New HRCORE.GradeTitle
        Set newGradeTitle.Category = selGrade
        Set newGradeTitle.position = selJPos
        retVal = newGradeTitle.InsertNew()
        LoadTitlesOfGrade selGrade, True
    End If
    
    Exit Sub
    
ErrorHandler:
End Sub

Private Sub cmdBackGTitle_Click()
    sstGrades.TabEnabled(0) = True
    sstGrades.Tab = 0
    sstGrades.TabVisible(1) = False
End Sub

Private Sub cmdBackS_Click()
    sstGrades.TabEnabled(0) = True
    sstGrades.TabVisible(2) = False
    sstGrades.Tab = 0
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdColorCode_Click()
    With CommonDialog1
        .DialogTitle = "Select Grade Color"
        .ShowColor
        lblColorCode.ForeColor = .Color
    End With
End Sub

Private Sub cmdDelete_Click()
    Dim resp As Long
    Dim retVal As Long
    
    Select Case LCase(cmdDelete.Caption)
        Case "delete"
            If selGrade Is Nothing Then
                MsgBox "Select the Grade that you want to Delete", vbInformation, TITLES
                Exit Sub
            Else
                resp = MsgBox("Are you sure you want to delete the selected grade?", vbYesNo + vbQuestion, TITLES)
                If resp = vbYes Then
                    retVal = selGrade.Delete()
                    
                    'force the selected item to nothing, coz it has been deleted
                    Set selGrade = Nothing
                    LoadGradesOfCSSSCategory selCSSCat, True
                    
                End If
            End If
        Case "cancel"
            cmdNew.Enabled = True
            cmdEdit.Caption = "Edit"
            cmdDelete.Caption = "Delete"
            fraGrade.Enabled = False
            EnableDisableControlsG False
            LoadGradesOfCSSSCategory selCSSCat, False
    End Select
End Sub

Private Sub cmdDeleteStep_Click()
    Dim retVal As Long
    Dim resp As Long
    
    
    On Error GoTo ErrorHandler
    
    Select Case LCase(cmdDeleteStep.Caption)
        Case "delete"
            If selCatStep Is Nothing Then
                MsgBox "Select the Step you want to Delete", vbInformation, TITLES
                Exit Sub
            Else
                resp = MsgBox("Are you sure you want to delete the selected step: " & selCatStep.Step, vbYesNo + vbQuestion, TITLES)
                If resp = vbYes Then
                    retVal = selCatStep.Delete()
                    
                    'force selected step to be nothing
                    Set selCatStep = Nothing
                    
                    LoadStepsOfGrade selGrade, True
                End If
            End If
            
        Case "cancel"
            cmdNewStep.Enabled = True
            cmdEditStep.Caption = "Edit"
            cmdDeleteStep.Caption = "Delete"
            txtStep.Locked = True
            txtSalary.Locked = True
            cmdBackS.Enabled = True
            LoadStepsOfGrade selGrade, False
    End Select
    
    Exit Sub
    
ErrorHandler:
End Sub

Private Sub cmdEdit_Click()
    Select Case LCase(cmdEdit.Caption)
        Case "edit"
            cmdNew.Enabled = False
            cmdEdit.Caption = "Update"
            cmdDelete.Caption = "Cancel"
            fraGrade.Enabled = True
            EnableDisableControlsG True
            
        Case "update"
                    If Not IsNumeric(txtBandSensitivityPercent.Text) Then
            MsgBox "band Sensitivity % required", vbOKOnly + vbCritical
            Exit Sub
            End If
            
            If Not (IsNumeric(txtBandSensitivityAmount.Text)) Then
            MsgBox "Band sensitivity amount is required ", vbOKOnly + vbCritical
            Exit Sub
            End If
            
            If txtBandSensitivityAmount.Text <= 0 Then
            MsgBox "Less than zero Sensitivity amount are are not allowed", vbOKOnly + vbCritical
            Exit Sub
            End If
            
            If txtBandSensitivityPercent.Text <= 0 Then
            MsgBox "Invalid bamd sensitivity", vbOKOnly + vbCritical
            Exit Sub
            End If
            If UpdateG() = False Then Exit Sub
            cmdNew.Enabled = True
            cmdEdit.Caption = "Edit"
            cmdDelete.Caption = "Delete"
            fraGrade.Enabled = False
            EnableDisableControlsG False
            LoadGradesOfCSSSCategory selCSSCat, True
            
        Case "cancel"
            cmdNew.Caption = "New"
            cmdEdit.Caption = "Edit"
            cmdDelete.Enabled = True
            fraGrade.Enabled = False
            EnableDisableControlsG False
            LoadGradesOfCSSSCategory selCSSCat, True
    End Select
End Sub

Private Sub cmdEditStep_Click()
    Select Case LCase(cmdEditStep.Caption)
        Case "edit"
            If selCatStep Is Nothing Then
                MsgBox "Select the Step to Edit", vbInformation, TITLES
                Exit Sub
            End If
            cmdNewStep.Enabled = False
            cmdEditStep.Caption = "Update"
            cmdDeleteStep.Caption = "Cancel"
            txtStep.Locked = False
            txtSalary.Locked = False
            cmdBackS.Enabled = False
            
        Case "update"
            If UpdateStep() = False Then Exit Sub
            cmdNewStep.Enabled = True
            cmdEditStep.Caption = "Edit"
            cmdDeleteStep.Caption = "Delete"
            txtStep.Locked = True
            txtSalary.Locked = True
            cmdBackS.Enabled = True
            LoadStepsOfGrade selGrade, True
            
        Case "cancel"
            cmdNewStep.Caption = "New"
            cmdEditStep.Caption = "Edit"
            cmdDeleteStep.Enabled = True
            txtStep.Locked = True
            txtSalary.Locked = True
            cmdBackS.Enabled = True
            LoadStepsOfGrade selGrade, False
    End Select
End Sub

Private Sub cmdNew_Click()
    Select Case LCase(cmdNew.Caption)
        Case "new"
            If selCSSCat Is Nothing Then
                MsgBox "First select a Staff Category in which to define the New Grade", vbInformation, TITLES
                cboStaffCategory.SetFocus
                Exit Sub
            End If
            cmdNew.Caption = "Update"
            cmdEdit.Caption = "Cancel"
            cmdDelete.Enabled = False
            Me.cmdSteps.Enabled = False
            EnableDisableControlsG True
            ClearControls
            fraGrade.Enabled = True
            
        Case "update"
        
            If Not IsNumeric(txtBandSensitivityPercent.Text) Then
            MsgBox "band Sensitivity % required", vbOKOnly + vbCritical
            Exit Sub
            End If
            
            If Not (IsNumeric(txtBandSensitivityAmount.Text)) Then
            MsgBox "Band sensitivity amount is required ", vbOKOnly + vbCritical
            Exit Sub
            End If
            
            If txtBandSensitivityAmount.Text <= 0 Then
            MsgBox "Less than zero Sensitivity amount are are not allowed", vbOKOnly + vbCritical
            Exit Sub
            End If
            
            If txtBandSensitivityPercent.Text <= 0 Then
            MsgBox "Invalid bamd sensitivity", vbOKOnly + vbCritical
            Exit Sub
            End If
        
            If InsertNewG() = False Then Exit Sub

            cmdNew.Caption = "New"
            cmdEdit.Caption = "Edit"
            cmdDelete.Enabled = True
            fraGrade.Enabled = False
            EnableDisableControlsG False
            LoadGradesOfCSSSCategory selCSSCat, True
    End Select
End Sub

Private Sub cmdNewStep_Click()
    Select Case LCase(cmdNewStep.Caption)
        Case "new"
            cmdNewStep.Caption = "Update"
            cmdEditStep.Caption = "Cancel"
            cmdDeleteStep.Enabled = False
            txtStep.Text = ""
            txtSalary.Text = ""
            txtStep.Locked = False
            txtSalary.Locked = False
            cmdBackS.Enabled = False
            
            
        Case "update"
            If InsertNewStep() = False Then Exit Sub
            cmdNewStep.Caption = "New"
            cmdEditStep.Caption = "Edit"
            cmdDeleteStep.Enabled = True
            txtStep.Locked = True
            txtSalary.Locked = True
            cmdBackS.Enabled = True
            LoadStepsOfGrade selGrade, True
        
    End Select
End Sub

Private Sub cmdRemoveDesig_Click()
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    If selGradeTitle Is Nothing Then
        MsgBox "Select the Position to Remove", vbInformation, TITLES
    Else
        retVal = selGradeTitle.Delete()
        Set selGradeTitle = Nothing
        LoadTitlesOfGrade selGrade, True
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while removing the Grade Title" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub cmdSteps_Click()
  
    On Error GoTo ErrorHandler
    If selGrade Is Nothing Then
        MsgBox "First Select a Grade", vbInformation, TITLES
    Else
        Me.txtCategoryNameS.Text = selGrade.CategoryName
        Me.txtStaffCategoryS.Text = selCSSCat.CSSSCategoryName
        Me.txtLowestSalaryS.Text = selGrade.LowestSalary
        Me.txtHighestSalaryS.Text = selGrade.HighestSalary
        LoadStepsOfGrade selGrade, True
        sstGrades.TabEnabled(0) = False
        sstGrades.TabVisible(2) = True
        sstGrades.Tab = 2
    End If
    
    Exit Sub
ErrorHandler:

End Sub

Private Sub cmdTitles_Click()
    On Error GoTo ErrorHandler
    If selGrade Is Nothing Then
        MsgBox "First Select a Grade", vbInformation, TITLES
    Else
        Me.txtCategoryName2.Text = selGrade.CategoryName
        Me.txtStaffCategory.Text = selCSSCat.CSSSCategoryName
        LoadTitlesOfGrade selGrade, True
        sstGrades.TabEnabled(0) = False
        sstGrades.TabVisible(1) = True
        sstGrades.Tab = 1
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "An error has occurred" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    frmMain2.PositionTheFormWithoutEmpList Me
    
    sstGrades.TabVisible(1) = False
    sstGrades.TabVisible(2) = False
    
    Set pCSSCats = New HRCORE.CSSSCategories
    Set pGrades = New EmployeeCategories
    Set pJobPositions = New HRCORE.JobPositions
    Set pGradeTitles = New HRCORE.GradeTitles
    Set pCatSteps = New HRCORE.CategorySteps
    
    pGrades.GetActiveEmployeeCategories
    pJobPositions.GetAllJobPositions
    pGradeTitles.GetActiveGradeTitles
    pCatSteps.GetActiveCategorySteps
    
    LoadCSSSCategories
    LoadJobPositions
    
    
    Exit Sub
    
ErrorHandler:
    
End Sub

Private Sub LoadCSSSCategories()
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    pCSSCats.GetActiveCSSSCategories
    
    For i = 1 To pCSSCats.count
        cboStaffCategory.AddItem pCSSCats.Item(i).CSSSCategoryName
        cboStaffCategory.ItemData(cboStaffCategory.NewIndex) = pCSSCats.Item(i).CSSSCategoryID
    Next i
    
    If cboStaffCategory.ListCount > 0 Then cboStaffCategory.ListIndex = 0
    
    Exit Sub
    
ErrorHandler:
    MsgBox "The Staff Categories could not be loaded" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub


Private Sub LoadJobPositions()
    Dim i As Long
    
    On Error GoTo ErrorHandler
    For i = 1 To pJobPositions.count
        cboDesignations.AddItem pJobPositions.Item(i).PositionName
        cboDesignations.ItemData(cboDesignations.NewIndex) = pJobPositions.Item(i).PositionID
    Next i
    
    Exit Sub
    
ErrorHandler:
    
    
    
End Sub

Private Sub LoadGradesOfCSSSCategory(ByVal TheCSSSCat As HRCORE.CSSSCategory, ByVal Refresh As Boolean)
    Dim TheCats As HRCORE.EmployeeCategories
    
    On Error GoTo ErrorHandler
    
    If Refresh = True Then
        pGrades.GetActiveEmployeeCategories
    End If
    
    If Not (TheCSSSCat Is Nothing) Then
        Set TheCats = pGrades.GetEmployeeCategoriesByCSSSCategoryID(TheCSSSCat.CSSSCategoryID)
    End If
    PopulateGrades TheCats
    Exit Sub
ErrorHandler:
    
End Sub


Private Sub PopulateGrades(ByVal TheGrades As HRCORE.EmployeeCategories)
    Dim i As Long
    Dim ItemX As ListItem
    
    On Error GoTo ErrorHandler
    
    Me.lvwCSSSGrades.ListItems.Clear
    
    If Not (TheGrades Is Nothing) Then
        For i = 1 To TheGrades.count
            Set ItemX = Me.lvwCSSSGrades.ListItems.add(, , TheGrades.Item(i).CategoryName)
            ItemX.SubItems(1) = TheGrades.Item(i).CSSSCategory.CSSSCategoryName
            ItemX.SubItems(2) = TheGrades.Item(i).BandSensitivityPercent
            ItemX.SubItems(3) = TheGrades.Item(i).BandSensitivity
            ItemX.Tag = TheGrades.Item(i).CategoryID
        Next i
    End If
    
    Exit Sub
    
ErrorHandler:
            
End Sub


Private Sub ClearControls()
    Me.txtBandSensitivityAmount.Text = ""
    Me.txtBandSensitivityPercent.Text = ""
    Me.txtCategoryName.Text = ""
    Me.txtHighestSalary.Text = ""
    Me.txtLevel.Text = ""
    Me.txtLowestSalary.Text = ""
    
End Sub

Private Sub EnableDisableControlsG(ByVal EnableControls As Boolean)
    If EnableControls Then
        Me.txtBandSensitivityAmount.Locked = False
        Me.txtBandSensitivityPercent.Locked = False
        Me.txtCategoryName.Locked = False
        Me.txtHighestSalary.Locked = False
        Me.txtLevel.Locked = False
        Me.txtLowestSalary.Locked = False
        Me.cmdTitles.Enabled = False
        Me.cmdSteps.Enabled = False
        Me.cboStaffCategory.Locked = True
    Else
        Me.txtBandSensitivityAmount.Locked = True
        Me.txtBandSensitivityPercent.Locked = True
        Me.txtCategoryName.Locked = True
        Me.txtHighestSalary.Locked = True
        Me.txtLevel.Locked = True
        Me.txtLowestSalary.Locked = True
        Me.cmdTitles.Enabled = True
        Me.cmdSteps.Enabled = True
        Me.cboStaffCategory.Locked = False
    End If
End Sub


Private Function UpdateG() As Boolean
       
    On Error GoTo ErrorHandler
    
    With selGrade
        If Trim(Me.txtBandSensitivityPercent.Text) = "" Then
            .BandSensitivityPercent = 0#
        Else
            If IsNumeric(Trim(Me.txtBandSensitivityPercent.Text)) Then
                .BandSensitivityPercent = CSng(Trim(Me.txtBandSensitivityPercent.Text))
            Else
                MsgBox "Enter a numeric percentage value for the Band Sensitivity", vbExclamation, TITLES
                Me.txtBandSensitivityPercent.SetFocus
                Exit Function
            End If
        End If
        
        If Trim(Me.txtBandSensitivityAmount.Text) = "" Then
            .BandSensitivity = 0#
        Else
            If IsNumeric(Trim(Me.txtBandSensitivityAmount.Text)) Then
                .BandSensitivity = CSng(Trim(Me.txtBandSensitivityAmount.Text))
            Else
                MsgBox "Enter a numeric value for the Band Sensitivity", vbExclamation, TITLES
                Me.txtBandSensitivityAmount.SetFocus
                Exit Function
            End If
        End If
        
        If Trim(Me.txtCategoryName.Text) = "" Then
            MsgBox "Enter the Name of the Grade", vbExclamation, TITLES
            Me.txtCategoryName.SetFocus
            Exit Function
        Else
            .CategoryName = Trim(Me.txtCategoryName.Text)
        End If
        
        If Trim(Me.txtLevel.Text) = "" Then
            .CategoryLevel = 0
        Else
            If IsNumeric(Trim(Me.txtLevel.Text)) Then
                .CategoryLevel = CSng(Trim(Me.txtLevel.Text))
            Else
                MsgBox "Enter a numeric value for the Level of the Grade", vbExclamation, TITLES
                Me.txtLevel.SetFocus
                Exit Function
            End If
        End If
           
        If Trim(Me.txtLowestSalary.Text) = "" Then
            .LowestSalary = 0#
        Else
            If IsNumeric(Trim(Me.txtLowestSalary.Text)) Then
                .LowestSalary = CSng(Trim(Me.txtLowestSalary.Text))
            Else
                MsgBox "Enter a numeric value for the Lowest Salary", vbExclamation, TITLES
                Me.txtLowestSalary.SetFocus
                Exit Function
            End If
        End If
        
        If Trim(Me.txtHighestSalary.Text) = "" Then
            .HighestSalary = 0#
        Else
            If IsNumeric(Trim(Me.txtHighestSalary.Text)) Then
                .HighestSalary = CSng(Trim(Me.txtHighestSalary.Text))
            Else
                MsgBox "Enter a numeric value for the Highest Salary", vbExclamation, TITLES
                Me.txtHighestSalary.SetFocus
                Exit Function
            End If
        End If
        
        .CategoryColorCode = Me.lblColorCode.ForeColor
        Set .CSSSCategory = selCSSCat
        
        retVal = .Update()
        If retVal = 0 Then
        ''**********************************
        
        ''code for Auto generating steps added by kalya
        
        Dim cSteps As New HRCORE.CategorySteps
        Dim sal As Double
        If .HighestSalary > .LowestSalary Then
            If txtBandSensitivityAmount.Text > 0 Then
            sal = .LowestSalary
            Dim st As HRCORE.CategoryStep
            Dim i As Integer
            i = 1
                While sal <= .HighestSalary
                Set st = New HRCORE.CategoryStep
                st.Category = selGrade
                st.Step = i
                st.Salary = sal
                If st.Salary > .HighestSalary Then
                st.Salary = .HighestSalary
                End If
                cSteps.add st
                sal = sal + .BandSensitivity
                If sal > .HighestSalary Then
                    If st.Salary < .HighestSalary Then
                    sal = .HighestSalary
                    End If
                End If
                i = i + 1
                Wend
                
                If Not cSteps Is Nothing Then
                    If cSteps.count > 0 Then
                    Dim i2 As Integer
                    Dim ret As Integer
                    Dim st2 As HRCORE.CategoryStep
                    ''clear steps is exists
                    
                    Dim sql As String
                    sql = "delete from CategorySteps where categoryid=" & selGrade.CategoryID & ""
                    CConnect.ExecuteSql (sql)
                        For i2 = 1 To cSteps.count
                            Set st2 = cSteps.Item(i2)
                            ret = st2.InsertNew
                            If ret <> 0 Then
                            ret = st2.Update
                            End If
                        Next i2
                    End If
                End If
                
            End If
        End If

        
        
        ''***********************************
            MsgBox "The Grade has been Updated successfully", vbInformation, TITLES
            UpdateG = True
        End If
    End With
        
    Exit Function
        
ErrorHandler:
    MsgBox "an error occurred while Updating the Grade" & vbNewLine & err.Description, vbExclamation, TITLES
        UpdateG = False
End Function


Private Function InsertNewG() As Boolean
    Dim newGrade As HRCORE.EmployeeCategory
    
    On Error GoTo ErrorHandler
    Set newGrade = New HRCORE.EmployeeCategory
    
    With newGrade
        If Trim(Me.txtBandSensitivityPercent.Text) = "" Then
            .BandSensitivityPercent = 0#
        Else
            If IsNumeric(Trim(Me.txtBandSensitivityPercent.Text)) Then
                .BandSensitivityPercent = CSng(Trim(Me.txtBandSensitivityPercent.Text))
            Else
                MsgBox "Enter a numeric percentage value for the Band Sensitivity", vbExclamation, TITLES
                Me.txtBandSensitivityPercent.SetFocus
                Exit Function
            End If
        End If
        
        If Trim(Me.txtBandSensitivityAmount.Text) = "" Then
            .BandSensitivity = 0#
        Else
            If IsNumeric(Trim(Me.txtBandSensitivityAmount.Text)) Then
                .BandSensitivity = CSng(Trim(Me.txtBandSensitivityAmount.Text))
            Else
                MsgBox "Enter a numeric value for the Band Sensitivity", vbExclamation, TITLES
                Me.txtBandSensitivityAmount.SetFocus
                Exit Function
            End If
        End If
        
        If Trim(Me.txtCategoryName.Text) = "" Then
            MsgBox "Enter the Name of the Grade", vbExclamation, TITLES
            Me.txtCategoryName.SetFocus
            Exit Function
        Else
            .CategoryName = Trim(Me.txtCategoryName.Text)
        End If
        
        If Trim(Me.txtLevel.Text) = "" Then
            .CategoryLevel = 0
        Else
            If IsNumeric(Trim(Me.txtLevel.Text)) Then
                .CategoryLevel = CSng(Trim(Me.txtLevel.Text))
            Else
                MsgBox "Enter a numeric value for the Level of the Grade", vbExclamation, TITLES
                Me.txtLevel.SetFocus
                Exit Function
            End If
        End If
           
        If Trim(Me.txtLowestSalary.Text) = "" Then
            .LowestSalary = 0#
        Else
            If IsNumeric(Trim(Me.txtLowestSalary.Text)) Then
                .LowestSalary = CSng(Trim(Me.txtLowestSalary.Text))
            Else
                MsgBox "Enter a numeric value for the Lowest Salary", vbExclamation, TITLES
                Me.txtLowestSalary.SetFocus
                Exit Function
            End If
        End If
        
        If Trim(Me.txtHighestSalary.Text) = "" Then
            .HighestSalary = 0#
        Else
            If IsNumeric(Trim(Me.txtHighestSalary.Text)) Then
                .HighestSalary = CSng(Trim(Me.txtHighestSalary.Text))
            Else
                MsgBox "Enter a numeric value for the Highest Salary", vbExclamation, TITLES
                Me.txtHighestSalary.SetFocus
                Exit Function
            End If
        End If
        
        .CategoryColorCode = Me.lblColorCode.BackColor
        Set .CSSSCategory = selCSSCat
        
        retVal = .InsertNew()
        If retVal = 0 Then
        
        ''code for Auto generating steps added by kalya
        
        Dim cSteps As New HRCORE.CategorySteps
        Dim sal As Double
        If .HighestSalary > .LowestSalary Then
            If txtBandSensitivityAmount.Text > 0 Then
            sal = .LowestSalary
            Dim st As HRCORE.CategoryStep
            Dim i As Integer
            i = 1
                While sal <= .HighestSalary
                Set st = New HRCORE.CategoryStep
                st.Category = newGrade
                st.Step = i
                st.Salary = sal
                    If st.Salary > .HighestSalary Then
                     st.Salary = .HighestSalary
                    End If
                cSteps.add st
                sal = sal + .BandSensitivity
                    If sal > .HighestSalary Then
                        If st.Salary < .HighestSalary Then
                         sal = .HighestSalary
                        End If
                    End If
                i = i + 1
                Wend
                
                If Not cSteps Is Nothing Then
                    If cSteps.count > 0 Then
                    Dim i2 As Integer
                    Dim ret As Integer
                    Dim st2 As HRCORE.CategoryStep
                    Dim sql As String
                    sql = "delete from CategorySteps where categoryid=" & newGrade.CategoryID & ""
                    CConnect.ExecuteSql (sql)
                        For i2 = 1 To cSteps.count
                            Set st2 = cSteps.Item(i2)
                            ret = st2.InsertNew
                            If ret <> 0 Then
                            ret = st2.Update
                            End If
                        Next i2
                    End If
                End If
                
            End If
        End If
        
            MsgBox "The new Grade has been added successfully", vbInformation, TITLES
            InsertNewG = True
        End If
    End With
        
    Exit Function
        
ErrorHandler:
    MsgBox "an error occurred while inserting the new Grade" & vbNewLine & err.Description, vbExclamation, TITLES
        InsertNewG = False
End Function

Private Sub lvwCSSSGrades_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set selGrade = Nothing
    If IsNumeric(Item.Tag) Then
        Set selGrade = pGrades.FindEmployeeCategoryByID(CLng(Item.Tag))
    End If
    
    SetFieldsG selGrade
End Sub


Private Sub SetFieldsG(ByVal TheGrade As HRCORE.EmployeeCategory)
    On Error GoTo ErrorHandler
    ClearControls
    If Not (TheGrade Is Nothing) Then
        Me.txtBandSensitivityAmount.Text = TheGrade.BandSensitivity
        Me.txtBandSensitivityPercent.Text = TheGrade.BandSensitivityPercent
        Me.txtCategoryName.Text = TheGrade.CategoryName
        Me.txtLevel.Text = TheGrade.CategoryLevel
        Me.txtLowestSalary.Text = TheGrade.LowestSalary
        Me.txtHighestSalary.Text = TheGrade.HighestSalary
        Me.lblColorCode.BackColor = TheGrade.CategoryColorCode
    End If
    
    Exit Sub
    
ErrorHandler:
End Sub


Private Sub LoadTitlesOfGrade(ByVal TheGrade As HRCORE.EmployeeCategory, ByVal Refresh As Boolean)
    Dim i As Long
    
    Dim ItemX As ListItem
    
    On Error GoTo ErrorHandler
    
    Me.lvwPositions.ListItems.Clear
    
    If Refresh = True Then
        pGradeTitles.GetActiveGradeTitles
    End If
    
    If Not (TheGrade Is Nothing) Then
        Set SelGTitles = pGradeTitles.GetGradeTitlesOfGradeID(TheGrade.CategoryID)
        If Not (SelGTitles Is Nothing) Then
            For i = 1 To SelGTitles.count
                If Not (SelGTitles.Item(i).position Is Nothing) Then
                    Set ItemX = lvwPositions.ListItems.add(, , SelGTitles.Item(i).position.PositionName)
                    ItemX.Tag = SelGTitles.Item(i).GradeTitleID
                End If
            Next i
        End If
    End If
    
    Exit Sub
ErrorHandler:

End Sub

Private Sub lvwPositions_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo ErrorHandler
    
    Set selGradeTitle = Nothing
    If IsNumeric(Item.Tag) Then
        Set selGradeTitle = pGradeTitles.FindGradeTitleByID(CLng(Item.Tag))
    End If
    
    Exit Sub
    
ErrorHandler:
End Sub


Private Function InsertNewStep() As Boolean
    Dim newStep As HRCORE.CategoryStep
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    
    Set newStep = New HRCORE.CategoryStep
    
    If Trim(Me.txtStep.Text) = "" Then
        MsgBox "Enter the Step Value e.g. 1 or Step 1", vbInformation, TITLES
        Me.txtStep.SetFocus
        Exit Function
    Else
        newStep.Step = Trim(txtStep.Text)
    End If
    
    'check existence
    If pCatSteps.StepExists(newStep.Step, selGrade.CategoryID) = True Then
        MsgBox "A Similar Step already exists in the selected category", vbInformation, TITLES
        Me.txtStep.SetFocus
        Exit Function
    End If
    
    If Trim(Me.txtSalary.Text) = "" Then
        newStep.Salary = 0#
    Else
        If IsNumeric(Trim(Me.txtSalary.Text)) Then
            newStep.Salary = CSng(Trim(txtSalary.Text))
        Else
            MsgBox "Enter a numeric value for the salary", vbInformation, TITLES
            Me.txtSalary.SetFocus
            Exit Function
        End If
    End If
    
    If ((newStep.Salary > selGrade.HighestSalary) And (selGrade.HighestSalary <> 0)) Or ((newStep.Salary < selGrade.LowestSalary) And (selGrade.LowestSalary <> 0)) Then
        MsgBox "The Salary Value of this Step is outside the Range allowed by the Grade", vbInformation, TITLES
        Me.txtSalary.SetFocus
        Exit Function
    End If
    
    Set newStep.Category = selGrade
    
    retVal = newStep.InsertNew()
    If retVal = 0 Then
        MsgBox "The new step has been added successfully", vbInformation, TITLES
        InsertNewStep = True
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred while adding a new Step" & vbNewLine & err.Description, vbExclamation, TITLES
    InsertNewStep = False
End Function


Private Function UpdateStep() As Boolean
   
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    
    
    If Trim(Me.txtStep.Text) = "" Then
        MsgBox "Enter the Step Value e.g. 1 or Step 1", vbInformation, TITLES
        Me.txtStep.SetFocus
        Exit Function
    Else
        selCatStep.Step = Trim(txtStep.Text)
    End If
    
    'check existence
    If pCatSteps.StepExistsExclusive(selCatStep.Step, selGrade.CategoryID, selCatStep.CategoryStepID) = True Then
        MsgBox "A Similar Step already exists in the selected category", vbInformation, TITLES
        Me.txtStep.SetFocus
        Exit Function
    End If
    
    If Trim(Me.txtSalary.Text) = "" Then
        selCatStep.Salary = 0#
    Else
        If IsNumeric(Trim(Me.txtSalary.Text)) Then
            selCatStep.Salary = CSng(Trim(txtSalary.Text))
        Else
            MsgBox "Enter a numeric value for the salary", vbInformation, TITLES
            Me.txtSalary.SetFocus
            Exit Function
        End If
    End If
    
    If ((selCatStep.Salary > selGrade.HighestSalary) And (selGrade.HighestSalary <> 0)) Or ((selCatStep.Salary < selGrade.LowestSalary) And (selGrade.LowestSalary <> 0)) Then
        MsgBox "The Salary Value of this Step is outside the Range allowed by the Grade", vbInformation, TITLES
        Me.txtSalary.SetFocus
        Exit Function
    End If
    
    Set selCatStep.Category = selGrade
    
    retVal = selCatStep.Update()
    If retVal = 0 Then
        MsgBox "The step has been updated successfully", vbInformation, TITLES
        UpdateStep = True
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred while Updateting the Step" & vbNewLine & err.Description, vbExclamation, TITLES
    UpdateStep = False
End Function



Private Sub LoadStepsOfGrade(ByVal TheGrade As EmployeeCategory, ByVal Refresh As Boolean)
    Dim i As Long
    Dim ItemX As ListItem
    
    On Error GoTo ErrorHandler
    
    Me.lvwSteps.ListItems.Clear
    
    If Refresh = True Then
        pCatSteps.GetActiveCategorySteps
    End If
    
    If Not (TheGrade Is Nothing) Then
        Set selCatSteps = pCatSteps.GetCategoryStepsOfGradeID(TheGrade.CategoryID)
        If Not (selCatSteps Is Nothing) Then
            For i = 1 To selCatSteps.count
                Set ItemX = lvwSteps.ListItems.add(, , selCatSteps.Item(i).Step)
                ItemX.SubItems(1) = selCatSteps.Item(i).Salary
                ItemX.SubItems(2) = selCatSteps.Item(i).Category.CategoryName
                ItemX.SubItems(3) = selCatSteps.Item(i).Category.CSSSCategory.CSSSCategoryName
                ItemX.Tag = selCatSteps.Item(i).CategoryStepID
            Next i
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while populating the Category Steps" & vbNewLine & err.Description, vbExclamation, TITLES
                
End Sub

Private Sub lvwSteps_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo ErrorHandler
    
    Set selCatStep = Nothing
    Me.txtStep.Text = ""
    Me.txtSalary.Text = ""
    If IsNumeric(Item.Tag) Then
        Set selCatStep = pCatSteps.FindCategoryStepByID(CLng(Item.Tag))
        If Not (selCatStep Is Nothing) Then
            Me.txtStep.Text = selCatStep.Step
            Me.txtSalary.Text = selCatStep.Salary
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub txtBandSensitivityAmount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If IsNumeric(txtBandSensitivityAmount.Text) Then
        If IsNumeric(txtHighestSalary.Text) Then
            If IsNumeric(txtLowestSalary.Text) Then
            txtBandSensitivityPercent.Text = (txtBandSensitivityAmount.Text * 100) / (txtHighestSalary.Text - txtLowestSalary.Text)
            End If
        End If
    End If
End If
End Sub

Private Sub txtBandSensitivityAmount_LostFocus()

    If IsNumeric(txtBandSensitivityAmount.Text) Then
        If IsNumeric(txtHighestSalary.Text) Then
            If IsNumeric(txtLowestSalary.Text) Then
            txtBandSensitivityPercent.Text = (txtBandSensitivityAmount.Text * 100) / (txtHighestSalary.Text - txtLowestSalary.Text)
            End If
        End If
    End If

End Sub

Private Sub txtBandSensitivityPercent_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsNumeric(txtBandSensitivityPercent.Text) Then
    If IsNumeric(txtHighestSalary.Text) Then
        If IsNumeric(txtLowestSalary.Text) Then
        txtBandSensitivityAmount.Text = (txtBandSensitivityPercent.Text * (txtHighestSalary.Text - txtLowestSalary.Text)) / 100
        End If
    End If
End If
End If
End Sub

Private Sub txtBandSensitivityPercent_LostFocus()

If IsNumeric(txtBandSensitivityPercent.Text) Then
    If IsNumeric(txtHighestSalary.Text) Then
        If IsNumeric(txtLowestSalary.Text) Then
        txtBandSensitivityAmount.Text = (txtBandSensitivityPercent.Text * (txtHighestSalary.Text - txtLowestSalary.Text)) / 100
        End If
    End If
End If

End Sub
