VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEmpCategories 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Categories"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11175
   Icon            =   "frmEmpCategories.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      Caption         =   "Employee Categories"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3510
      Left            =   1065
      TabIndex        =   15
      Top             =   1050
      Visible         =   0   'False
      Width           =   6015
      Begin VB.Frame fraColorCode 
         Caption         =   "Color Code"
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
         Left            =   3780
         TabIndex        =   22
         Top             =   990
         Width           =   2085
         Begin VB.TextBox txtColorCode 
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
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   360
            Width           =   825
         End
         Begin VB.CommandButton cmdColorCode 
            Caption         =   "Set ..."
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
            Left            =   1170
            TabIndex        =   5
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame fraSalaryRange 
         Caption         =   "Salary range"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   3525
         Begin VB.TextBox txtHighestSal 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2040
            TabIndex        =   3
            Text            =   "0"
            Top             =   480
            Width           =   1245
         End
         Begin VB.TextBox txtLowestSal 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            TabIndex        =   2
            Text            =   "0"
            Top             =   480
            Width           =   1245
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Highest"
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
            Left            =   2040
            TabIndex        =   21
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Lowest"
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
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   510
         End
      End
      Begin VB.TextBox txtCategoryLevel 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4635
         TabIndex        =   1
         Top             =   645
         Width           =   1200
      End
      Begin VB.TextBox txtCategoryName 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   135
         TabIndex        =   0
         Top             =   645
         Width           =   3915
      End
      Begin VB.TextBox txtCategoryDetails 
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
         ForeColor       =   &H00000000&
         Height          =   585
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   2175
         Width           =   5715
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4860
         Picture         =   "frmEmpCategories.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Save Record"
         Top             =   2880
         Width           =   510
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5370
         Picture         =   "frmEmpCategories.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Cancel Process"
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Category Level"
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
         Left            =   4635
         TabIndex        =   18
         Top             =   405
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Category Name"
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
         Left            =   135
         TabIndex        =   17
         Top             =   405
         Width           =   1125
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Other Details"
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
         Left            =   135
         TabIndex        =   16
         Top             =   1950
         Width           =   945
      End
   End
   Begin MSComctlLib.ImageList imgEmpTool 
      Left            =   5430
      Top             =   195
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpCategories.frx":0646
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpCategories.frx":0758
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpCategories.frx":086A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpCategories.frx":097C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   7900
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   11130
      Begin MSComDlg.CommonDialog CdlColorCode 
         Left            =   5610
         Top             =   4560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "CLOSE"
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
         Left            =   6270
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   5400
         Visible         =   0   'False
         Width           =   1050
      End
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
         Left            =   2775
         Picture         =   "frmEmpCategories.frx":0EBE
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Add New record"
         Top             =   5400
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
         Left            =   3255
         Picture         =   "frmEmpCategories.frx":0FC0
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Edit Record"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
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
         Left            =   3735
         Picture         =   "frmEmpCategories.frx":10C2
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Delete Record"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSComctlLib.ListView lvwDetails 
         Height          =   7800
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   11130
         _ExtentX        =   19632
         _ExtentY        =   13758
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgTree"
         ForeColor       =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmEmpCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private EmpCats As HRCORE.EmployeeCategories
Private selEmpCat As HRCORE.EmployeeCategory
Private lngColorCode As Long
Private blnEditMode As Boolean


Public Sub cmdCancel_Click()

    If PromptSave = True Then
        If MsgBox("Close this window?", vbYesNo + vbQuestion, "Confirm Close") = vbNo Then Exit Sub
    End If
    fraDetails.Visible = False
    With frmMain2
        .cmdNew.Enabled = True
        .cmdEdit.Enabled = True
        .cmdDelete.Enabled = True
        .cmdCancel.Enabled = False
        .cmdSave.Enabled = False
    End With
    Call EnableCmd
End Sub

Private Sub cmdColorCode_Click()
    On Error GoTo ErrorHandler
    With CdlColorCode
        .DialogTitle = "Choose a Color"
        .ShowColor
        .Flags = cdlCCRGBInit
        lngColorCode = .Color
    End With
    
    Me.txtColorCode.BackColor = lngColorCode
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while setting the Color Code" & _
    err.Description, vbInformation, TITLES
End Sub

Public Sub cmdDelete_Click()
    Dim resp As String
    
    If Not currUser Is Nothing Then
        If currUser.CheckRight("EmployeeCategories") <> secModify Then
            MsgBox "You dont have right to delete the  record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    If lvwDetails.ListItems.count > 0 Then
        resp = MsgBox("This will delete  " & lvwDetails.SelectedItem & ". Do you wish to continue?", vbQuestion + vbYesNo)
        If resp = vbNo Then
            Exit Sub
        End If
          
        Action = "DELETED AN EMPLOYEE CATEGORY; CATEGORY CODE: " & lvwDetails.SelectedItem
        
        CConnect.ExecuteSql ("DELETE FROM ECategory WHERE Code = '" & lvwDetails.SelectedItem & "'")
        
    '    rs2.Requery
        
        Call frmMain2.LoadCbo
        Call LoadEmployeeCategories
        Call cmdCancel_Click
            
    Else
        MsgBox "You have to select the employee category you would like to delete.", vbInformation
                
    End If
        
    
    
End Sub

Public Sub cmdEdit_Click()
    
    On Error GoTo ErrorHandler
   If Not currUser Is Nothing Then
        If currUser.CheckRight("EmployeeCategories") <> secModify Then
            MsgBox "You dont have right to edit the record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    If selEmpCat Is Nothing Then
        MsgBox "You have to select the contact type you would like to edit.", vbInformation
        Call cmdCancel_Click
        Exit Sub
    End If
    
    'otherwise set thefields
    SetFields selEmpCat
    
    Call DisableCmd
    
    fraDetails.Visible = True
    
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    cmdEdit.Enabled = True
    SaveNew = False
    
    blnEditMode = True
    
    Me.txtCategoryName.SetFocus
    Exit Sub
    
ErrorHandler:
    
End Sub

Private Sub SetFields(ByVal TheSelectedEmpCat As EmployeeCategory)
    On Error GoTo ErrorHandler
    
    Cleartxt
    If TheSelectedEmpCat Is Nothing Then
        Exit Sub
    End If
    
    With TheSelectedEmpCat
        Me.txtCategoryDetails.Text = .CategoryDetails
        Me.txtCategoryLevel.Text = .CategoryLevel
        Me.txtCategoryName.Text = .CategoryName
        Me.txtHighestSal.Text = .HighestSalary
        Me.txtLowestSal.Text = .LowestSalary
        Me.txtColorCode.BackColor = .CategoryColorCode
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while populating the Employee Category Details" & _
    err.Description, vbInformation, TITLES
End Sub

Public Sub cmdNew_Click()
    
    If Not currUser Is Nothing Then
        If currUser.CheckRight("EmployeeCategories") <> secModify Then
            MsgBox "You dont have right to add a new record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
    
    Call DisableCmd
    Me.txtCategoryName.Text = ""
    Me.txtCategoryDetails.Text = ""
    fraDetails.Visible = True
    cmdCancel.Enabled = True
    SaveNew = True
    blnEditMode = False
    cmdSave.Enabled = True
    Me.txtCategoryName.SetFocus

End Sub


Public Sub cmdSave_Click()
    Dim NewEmpCat As HRCORE.EmployeeCategory
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    
    Set NewEmpCat = New HRCORE.EmployeeCategory
    
    With NewEmpCat
        If blnEditMode = True Then
            .CategoryID = selEmpCat.CategoryID
        End If
        
        If Trim(Me.txtCategoryName.Text) = "" Then
            MsgBox "Enter the Category Name.", vbExclamation
            Me.txtCategoryName.SetFocus
            Exit Sub
        End If
        
        .CategoryName = Trim(Me.txtCategoryName.Text)
        If blnEditMode = True Then
            If Not (EmpCats.FindEmployeeCategoryNameExclusive(.CategoryName, .CategoryID) Is Nothing) Then
                MsgBox "Another Employee Category exists with the supplied name", vbInformation, TITLES
                Me.txtCategoryName.SetFocus
                Exit Sub
            End If
        Else
            If Not (EmpCats.FindEmployeeCategoryByName(.CategoryName) Is Nothing) Then
                MsgBox "Another Employee Category exists with the supplied name", vbInformation, TITLES
                Me.txtCategoryName.SetFocus
                Exit Sub
            End If
        End If
        .CategoryColorCode = Me.txtColorCode.BackColor
        If IsNumeric(Trim(Me.txtCategoryLevel.Text)) Then
            .CategoryLevel = CInt(Trim(Me.txtCategoryLevel.Text))
        Else
            If Len(Trim(Me.txtCategoryLevel.Text)) > 0 Then
                MsgBox "Enter a numeric value for the Category Level", vbInformation, TITLES
                Me.txtCategoryLevel.SetFocus
                Exit Sub
            Else
                .CategoryLevel = 0
            End If
        End If
        
        If IsNumeric(Trim(Me.txtHighestSal.Text)) Then
            .HighestSalary = CSng(Trim(Me.txtHighestSal.Text))
        Else
            If Len(Trim(Me.txtHighestSal.Text)) > 0 Then
                MsgBox "Enter a numeric value for the Highest Salary Level", vbInformation, TITLES
                Me.txtHighestSal.SetFocus
                Exit Sub
            Else
                .HighestSalary = 0
            End If
        End If
        
        If IsNumeric(Trim(Me.txtLowestSal.Text)) Then
            .LowestSalary = CSng(Trim(Me.txtLowestSal.Text))
        Else
            If Len(Trim(Me.txtLowestSal.Text)) > 0 Then
                MsgBox "Enter a numeric value for the Lowest Salary Level", vbInformation, TITLES
                Me.txtLowestSal.SetFocus
                Exit Sub
            Else
                .LowestSalary = 0
            End If
        End If
        
        If (.HighestSalary < .LowestSalary) And (.HighestSalary <> 0) Then
            MsgBox "The Highest Salary Cannot be lower than the Lowest Salary" & vbNewLine & _
            "Enter 0 (Zero) for Limitless Highest Salary", vbInformation, TITLES
            Me.txtHighestSal.SetFocus
            Exit Sub
        End If
        .CategoryDetails = Me.txtCategoryDetails.Text
        
        If blnEditMode = True Then
            retVal = .Update()
            If retVal = 0 Then
                MsgBox "The Employee Category has been updated successfully", vbInformation, TITLES
                'reload the data
                LoadEmployeeCategories
                
                'reload empCats in frmMain2
                Call frmMain2.LoadEmployeeCategories
                
                'reload the employees
                Call frmMain2.LoadEmployeeList
                
            Else
                MsgBox "The Employee Category could not be updated", vbInformation, TITLES
                
            End If
        Else
            retVal = .InsertNew()
            If retVal = 0 Then
                MsgBox "The new Employee Category has been added successfully", vbInformation, TITLES
                'reload data
                LoadEmployeeCategories
                
                'reload empCats in frmMain2
                Call frmMain2.LoadEmployeeCategories
                
            Else
                MsgBox "The New Employee Category could not be added", vbInformation, TITLES
            End If
        End If
    End With
       
    Me.fraDetails.Visible = False
    Call cmdCancel_Click
    Exit Sub
        
ErrorHandler:
    MsgBox "An error has occurred while configuring the Employee Category" & _
    err.Description, vbInformation, TITLES
    
End Sub


Private Sub Form_Load()

    On Error GoTo ErrorHandler
    
    Set EmpCats = New HRCORE.EmployeeCategories
    
    Call InitGrid
     
    LoadEmployeeCategories
    
    frmMain2.PositionTheFormWithoutEmpList Me
    cmdCancel.Enabled = False
    cmdSave.Enabled = False

    Exit Sub
ErrorHandler:
    MsgBox "An error has occurred" & vbNewLine & err.Description, vbInformation, TITLES
End Sub

Private Sub Form_Resize()
    oSmart.FResize Me
    Me.Frame1.Move Frame1.Left, Frame1.Top, Frame1.Width, tvwMainheight - 200
    lvwDetails.Height = tvwMainheight - 210
End Sub

Private Sub InitGrid()
    With lvwDetails
        .ColumnHeaders.add , , "Category Name", .Width / 3
        .ColumnHeaders.add , , "Level", .Width / 6
        .ColumnHeaders.add , , "Lowest Salary", .Width / 4
        .ColumnHeaders.add , , "Highest Salary", .Width / 4
        .ColumnHeaders.add , , "Details", .Width / 2
                   
        .View = lvwReport
    End With
    

End Sub

Public Sub LoadEmployeeCategories()
    Dim empCat As HRCORE.EmployeeCategory
    Dim i As Long
    Dim ItemX As ListItem
    
    On Error GoTo ErrorHandler
   
    lvwDetails.ListItems.Clear
    
    Call Cleartxt
    
    EmpCats.GetAllEmployeeCategories
    
    For i = 1 To EmpCats.count
        Set empCat = EmpCats.Item(i)
        Set ItemX = Me.lvwDetails.ListItems.add(, , empCat.CategoryName)
        ItemX.SubItems(1) = empCat.CategoryLevel
        ItemX.SubItems(2) = empCat.LowestSalary
        ItemX.SubItems(3) = empCat.HighestSalary
        ItemX.SubItems(4) = empCat.CategoryDetails
        ItemX.Tag = empCat.CategoryID
    Next i

    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while Populating Employee Categories" & _
    err.Description, vbInformation, TITLES
End Sub


Private Sub Form_Unload(Cancel As Integer)
  
    frmMain2.Caption = "Infiniti HRMIS - PDR [Current User:\" & currUser.FullNames & "]"
    
End Sub

Private Sub lvwDetails_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwDetails
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvwDetails_DblClick()
    Me.cmdEdit.Enabled = True
    If frmMain2.cmdEdit.Enabled = True Then
        Call frmMain2.cmdEdit_Click
    End If
End Sub

Private Sub Cleartxt()
    Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
            i.Text = ""
        End If
    Next i

    'lvwDetails.ListItems.Clear
    
End Sub

Public Sub disabletxt()
    Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
            i.Locked = True
        End If
    Next i
End Sub

Public Sub DisableCmd()
    Dim i As Object
    For Each i In Me
        If TypeOf i Is CommandButton Then
            i.Enabled = False
        End If
    Next i
    
    'enable the cmdColorCode commandbutton
    cmdColorCode.Enabled = True
    
End Sub

Public Sub EnableCmd()
    Dim i As Object
    For Each i In Me
        If TypeOf i Is CommandButton Then
            i.Enabled = True
        End If
    Next i
    
End Sub


Private Sub lvwDetails_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo ErrorHandler
    
    Set selEmpCat = Nothing
    Set selEmpCat = EmpCats.FindEmployeeCategory(CLng(Item.Tag))
    
    Exit Sub
    
ErrorHandler:
    MsgBox "A slight error has occurred" & vbNewLine & err.Description, vbInformation, TITLES
    
End Sub

