VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmEmployment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Employment History"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   Icon            =   "frmEmployment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
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
            Picture         =   "frmEmployment.frx":0442
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployment.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployment.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployment.frx":0778
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
      TabIndex        =   22
      Top             =   0
      Width           =   7800
      Begin MSComDlg.CommonDialog Cdl 
         Left            =   4080
         Top             =   5760
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   435
         TabIndex        =   14
         ToolTipText     =   "Move to the First employee"
         Top             =   7155
         Visible         =   0   'False
         Width           =   495
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
         Left            =   6255
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   7185
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   915
         TabIndex        =   15
         ToolTipText     =   "Move to the Previous employee"
         Top             =   7155
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1395
         TabIndex        =   16
         ToolTipText     =   "Move to the Next employee"
         Top             =   7155
         Visible         =   0   'False
         Width           =   495
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
         Left            =   3090
         Picture         =   "frmEmployment.frx":0CBA
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Add New record"
         Top             =   7155
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
         Left            =   3570
         Picture         =   "frmEmployment.frx":0DBC
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Edit Record"
         Top             =   7155
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
         Left            =   4050
         Picture         =   "frmEmployment.frx":0EBE
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Delete Record"
         Top             =   7155
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1875
         TabIndex        =   17
         ToolTipText     =   "Move to the Last employee"
         Top             =   7155
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame fraDetails 
         Appearance      =   0  'Flat
         Caption         =   "Employment History"
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
         Height          =   6045
         Left            =   480
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   6375
         Begin VB.Frame Frabenefits 
            Caption         =   "Benefit details"
            Height          =   1335
            Left            =   2760
            TabIndex        =   40
            Top             =   2520
            Visible         =   0   'False
            Width           =   3495
            Begin VB.TextBox txtBenefit 
               Height          =   1095
               Left            =   0
               MultiLine       =   -1  'True
               TabIndex        =   41
               Top             =   240
               Width           =   3495
            End
         End
         Begin VB.CheckBox chkBenefit 
            Caption         =   "Benefits"
            Height          =   255
            Left            =   4440
            TabIndex        =   39
            Top             =   2230
            Width           =   1095
         End
         Begin VB.OptionButton optGross 
            Caption         =   "Gross "
            Height          =   375
            Left            =   2640
            TabIndex        =   38
            Top             =   2160
            Width           =   735
         End
         Begin VB.ComboBox cboCurrency 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   2280
            Width           =   1095
         End
         Begin VB.OptionButton optNet 
            Caption         =   "Net"
            Height          =   375
            Left            =   3480
            TabIndex        =   35
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtPhone 
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
            Height          =   300
            Left            =   150
            TabIndex        =   9
            Top             =   3420
            Width           =   2385
         End
         Begin VB.TextBox txtAddress 
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
            Height          =   450
            Left            =   135
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   3945
            Width           =   6075
         End
         Begin VB.TextBox txtSalary 
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
            Height          =   300
            Left            =   120
            TabIndex        =   8
            Top             =   2280
            Width           =   1200
         End
         Begin VB.TextBox txtSuper 
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
            Height          =   300
            Left            =   150
            TabIndex        =   7
            Top             =   2820
            Width           =   2385
         End
         Begin VB.TextBox txtEmployer 
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
            Height          =   300
            Left            =   1800
            TabIndex        =   2
            Top             =   480
            Width           =   4380
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
            Left            =   5760
            Picture         =   "frmEmployment.frx":13B0
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Cancel Process"
            Top             =   5460
            Width           =   495
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
            Left            =   5250
            Picture         =   "frmEmployment.frx":14B2
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Save Record"
            Top             =   5460
            Width           =   495
         End
         Begin VB.TextBox txtComments 
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
            Height          =   645
            Left            =   135
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   4635
            Width           =   6075
         End
         Begin VB.TextBox txtDesig 
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
            TabIndex        =   5
            Top             =   1800
            Width           =   6135
         End
         Begin VB.TextBox txtCode 
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
            Height          =   300
            Left            =   120
            TabIndex        =   1
            Top             =   480
            Width           =   1260
         End
         Begin VB.TextBox txtReasons 
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
            Height          =   315
            Left            =   2760
            TabIndex        =   6
            Top             =   1080
            Width           =   3435
         End
         Begin MSComCtl2.DTPicker dtpFrom 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   330
            Left            =   150
            TabIndex        =   3
            Top             =   1110
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   62849027
            CurrentDate     =   37673
            MinDate         =   21916
         End
         Begin MSComCtl2.DTPicker dtpTo 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   330
            Left            =   1440
            TabIndex        =   4
            Top             =   1110
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   62849027
            CurrentDate     =   37673
            MinDate         =   21916
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Currency"
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
            Left            =   1440
            TabIndex        =   37
            Top             =   2100
            Width           =   660
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            Caption         =   "From"
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
            Left            =   150
            TabIndex        =   34
            Top             =   900
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "To"
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
            Left            =   1455
            TabIndex        =   33
            Top             =   900
            Width           =   180
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Phone"
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
            Left            =   150
            TabIndex        =   32
            Top             =   3180
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Address"
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
            TabIndex        =   31
            Top             =   3705
            Width           =   585
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Salary"
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
            TabIndex        =   30
            Top             =   2100
            Width           =   450
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Supervisor"
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
            Left            =   150
            TabIndex        =   29
            Top             =   2580
            Width           =   765
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Employer"
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
            Left            =   1800
            TabIndex        =   28
            Top             =   240
            Width           =   660
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Comments"
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
            TabIndex        =   27
            Top             =   4395
            Width           =   750
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Position Held"
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
            TabIndex        =   26
            Top             =   1560
            Width           =   915
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Code"
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
            TabIndex        =   25
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Reasons for leaving"
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
            Left            =   2760
            TabIndex        =   24
            Top             =   840
            Width           =   1425
         End
      End
      Begin MSComctlLib.ListView lvwDetails 
         Height          =   7800
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   7800
         _ExtentX        =   13758
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
Attribute VB_Name = "frmEmployment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CurrID As Integer

Private Sub chkBenefit_Click()
    If chkBenefit.value = vbChecked Then
        Frabenefits.Visible = True
    Else
        Frabenefits.Visible = False
    End If
End Sub

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
    
    Call Cleartxt
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Public Sub cmdDelete_Click()
    Dim resp As String
    On Error GoTo ErrHandler
    'check rigths
    If Not currUser Is Nothing Then
        If currUser.CheckRight("EmploymentHistory") <> secModify Then
            MsgBox "You dont have right to delete the record. Please liaise with the security admin"
            Exit Sub
        End If
     End If
     
    If SelectedEmployee Is Nothing Then Exit Sub
    
    If lvwDetails.ListItems.count > 0 Then
        resp = MsgBox("Are you sure you want to delete  " & lvwDetails.SelectedItem & " from the records?", vbQuestion + vbYesNo)
        If resp = vbNo Then
            Exit Sub
        End If
             
        Action = "DELETED EMPLOYMENT HISTORY; EMPLOYEE CODE: " & SelectedEmployee.EmpCode & "; EMPLOYMENT HISTORY CODE: " & lvwDetails.SelectedItem.Text
        
        CConnect.ExecuteSql ("DELETE FROM Employment WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & lvwDetails.SelectedItem & "'")
             
        rs2.Requery
        
        Call DisplayRecords
            
    Else
        MsgBox "You have to select the Employment you would like to delete.", vbInformation
                
    End If
    Exit Sub
ErrHandler:
End Sub

Private Sub cmdDone_Click()
    PSave = True
    Call cmdCancel_Click
    PSave = False
End Sub

Public Sub cmdEdit_Click()
    On Error Resume Next
     'check rigths
    If Not currUser Is Nothing Then
        If currUser.CheckRight("EmploymentHistory") <> secModify Then
            MsgBox "You dont have right to modify the record. Please liaise with the security admin"
            Exit Sub
        End If
     End If
     
     If SelectedEmployee Is Nothing Then Exit Sub
     
    If lvwDetails.ListItems.count < 1 Then
        MsgBox "You have to select the Employment you would like to edit.", vbInformation
        PSave = True
        Call cmdCancel_Click
        PSave = False
        Exit Sub
    End If
    
    Set rs3 = CConnect.GetRecordSet("SELECT * FROM Employment WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & lvwDetails.SelectedItem & "'")
    
    With rs3
        If .RecordCount > 0 Then
            txtCode.Text = !Code & ""
            txtReasons.Text = !Reasons & ""
            If Not IsNull(!cFrom) Then dtpFrom.value = !cFrom & ""
            If Not IsNull(!cTo) Then dtpTo.value = !cTo & ""
            txtComments.Text = !Comments & ""
            txtDesig.Text = !Desig & ""
            txtEmployer.Text = !Employer & ""
            txtSalary.Text = Format(!Salary, Cfmt)
            txtSuper.Text = !Super & ""
            txtComments.Text = !Comments & ""
            txtPhone.Text = !Phone & ""
            txtAddress.Text = !Address & ""
            If Not IsNull(!CurrencyID) Then
                cboCurrency.ListIndex = getIndex(!CurrencyID)
            Else
                cboCurrency.ListIndex = -1
            End If
            If Not IsNull(!Benefits) Then
                If !Benefits <> vbNullString Then
                    Me.txtBenefit = !Benefits & ""
                    Me.chkBenefit.value = vbChecked
                    Frabenefits.Visible = True
                End If
            Else
                Frabenefits.Visible = False
                Me.txtBenefit = ""
                Me.chkBenefit.value = vbUnchecked
            End If
            
            If Not IsNull(!isGross) Then
                If !isGross Then
                    optGross.value = True
                    optNet.value = False
                Else
                    optGross.value = False
                    optNet.value = True
                End If
            Else
                optGross.value = False
                optNet.value = False
            End If
            SaveNew = False
        Else
            MsgBox "Record not found.", vbInformation
            Set rs3 = Nothing
            PSave = True
            Call cmdCancel_Click
            PSave = False
            Exit Sub
        End If
    End With

    Set rs3 = Nothing
    
    Call DisableCmd
    
    fraDetails.Visible = True
    
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    SaveNew = False
    
    txtCode.Locked = True
    txtEmployer.SetFocus
    
End Sub

Private Sub cmdFirst_Click()

With rsGlob
    If .RecordCount > 0 Then
        If .BOF <> True Then
            .MoveFirst
            If .BOF = True Then
                .MoveFirst
                Call DisplayRecords
            Else
                Call DisplayRecords
            End If
            
            Call FirstDisb
            
        End If
    End If
End With


End Sub

Private Sub cmdLast_Click()
With rsGlob
    If .RecordCount > 0 Then
        If .EOF <> True Then
            .MoveLast
            If .EOF = True Then
                .MoveLast
                Call DisplayRecords
            Else
                Call DisplayRecords
            End If
            
            Call LastDisb
            
        End If
    End If
End With

End Sub

Public Sub cmdNew_Click()
    'check rigths
    If Not currUser Is Nothing Then
        If currUser.CheckRight("EmploymentHistory") <> secModify Then
            MsgBox "You dont have right to add new the record. Please liaise with the security admin"
            Exit Sub
        End If
     End If
    Call DisableCmd
    
    txtCode.Text = loadACode
    txtReasons.Text = ""
    txtDesig.Text = ""
    txtComments.Text = ""
    txtEmployer.Text = ""
    dtpFrom.value = Date
    dtpTo.value = Date
    fraDetails.Visible = True
    cmdCancel.Enabled = True
    SaveNew = True
    cmdSave.Enabled = True
    txtCode.Locked = False
    txtEmployer.SetFocus

End Sub

Private Sub cmdNext_Click()
    
    With rsGlob
        If .RecordCount > 0 Then
            If .EOF <> True Then
                .MoveNext
                If .EOF = True Then
                    .MoveLast
                    Call DisplayRecords
                Else
                    Call DisplayRecords
                End If
                Call LastDisb
            End If
        End If
    End With
End Sub

Private Sub cmdPrevious_Click()

    With rsGlob
        If .RecordCount > 0 Then
            If .BOF <> True Then
                .MovePrevious
                If .BOF = True Then
                    .MoveFirst
                    Call DisplayRecords
                Else
                    Call DisplayRecords
                End If
                
                Call FirstDisb
                
            End If
        End If
    End With
End Sub

Public Sub cmdSave_Click()
    On Error GoTo ErrHandler
    
    Dim isGross As Integer
    
    If SelectedEmployee Is Nothing Then Exit Sub
    
    If txtCode.Text = "" Then
        MsgBox "Enter the Employment code.", vbExclamation
        txtCode.SetFocus
        Exit Sub
    End If
    
    If txtEmployer.Text = "" Then
        MsgBox "Enter the Employer.", vbExclamation
        txtEmployer.SetFocus
        Exit Sub
    End If
    
    If dtpFrom.value > dtpTo.value Then
        MsgBox "Enter the valid start and end dates.", vbInformation
        dtpFrom.SetFocus
        Exit Sub
    End If
    
    If txtSalary.Text = "" Then
        txtSalary.Text = 0
    End If

    If SaveNew = True Then
        Set rs4 = CConnect.GetRecordSet("SELECT * FROM Employment WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & txtCode.Text & "'")
        With rs4
            If .RecordCount > 0 Then
                MsgBox "Employment code already exists. Enter another one.", vbInformation
                txtCode.Text = ""
                txtCode.SetFocus
                Set rs4 = Nothing
                Exit Sub
            End If
        End With
        Set rs4 = Nothing
    End If
       
    If PromptSave = True Then
        If MsgBox("Are you sure you want to save the record.", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    If optGross.value = True Then
        isGross = 1
    Else
        isGross = 0
    End If
    
    If (cboCurrency.ListIndex <> -1) Then CurrID = cboCurrency.ItemData(cboCurrency.ListIndex)
    
    CConnect.ExecuteSql ("DELETE FROM Employment WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & txtCode.Text & "'")
    
    mySQL = "INSERT INTO Employment (employee_id, Employer, Reasons, CFrom, CTo, Desig, Super, Salary, Comments, Phone, Address,code,Isgross,CurrencyId,Benefits)" & _
                        " VALUES('" & SelectedEmployee.EmployeeID & "','" & txtEmployer.Text & "','" & txtReasons.Text & "'," & _
                        "'" & Format(dtpFrom.value, "yyyy-MM-dd") & "','" & Format(dtpTo.value, "yyyy-MM-dd") & "','" & txtDesig.Text & "','" & txtSuper.Text & "'," & CCur(txtSalary.Text) & ",'" & txtComments.Text & "','" & txtPhone.Text & "','" & txtAddress.Text & "','" & txtCode.Text & "'," & isGross & "," & CurrID & ",'" & Me.txtBenefit.Text & "')"
  
    Action = "ADDED EMPLOYMENT HISTORY; EMPLOYEE CODE: " & SelectedEmployee.EmpCode & "; EMPLOYER: " _
    & txtEmployer.Text & "; REASON FOR LEAVING: " & txtReasons.Text & "; FROM: " _
    & Format(dtpFrom.value, "dd-MMM-yyyy") & "; TO: " & Format(dtpTo.value, "dd-MMM-yyyy") _
    & "; DESIGNATION: " & txtDesig.Text & "; SUPERVISOR: " & txtSuper.Text & "; BASIC SALARY: " _
    & CCur(txtSalary.Text) & "; COMMENTS: " & txtComments.Text & "; PHONE: " & txtPhone.Text _
    & "; ADDRESS: " & txtAddress.Text & "; CODE: " & txtCode.Text
    
    CConnect.ExecuteSql (mySQL)
    
    rs2.Requery
    
    If SaveNew = False Then
        PSave = True
        Call DisplayRecords
        Call cmdCancel_Click
        PSave = False
    Else
        rs2.Requery
        Call DisplayRecords
        txtCode.Text = loadACode
        txtEmployer.SetFocus
        SaveNew = False
        
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "An error has occured:" & vbCrLf & err.Description, vbExclamation, "Error"
End Sub


Private Sub Form_Load()
    On Error GoTo ErrHandler
    'resize the form
    oSmart.FReset Me
    If oSmart.hRatio > 1.1 Then
        With frmMain2
            Me.Move .tvwMain.Width + .lvwEmp.Width + (.tvwMain.Width / 36) * 2, (.Height / 5.52) ''- 155
        End With
    Else
         With frmMain2
            Me.Move .tvwMain.Width + .lvwEmp.Width + (.tvwMain.Width / 36#) * 2, .Height / 5.55
        End With
        
    End If
    
    CConnect.CColor Me, MyColor
    
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    
    Call InitGrid
    'Call 'CConnect.CCon
    Set rs2 = CConnect.GetRecordSet("SELECT * FROM Employment ORDER BY CFrom")
    With rsGlob
        If .RecordCount < 1 Then
            Call DisableCmd
            Exit Sub
        End If
    End With
    
    Call DisplayRecords
    
    cmdFirst.Enabled = False
    cmdPrevious.Enabled = False
    
'    Rcurrency.Requery
'
'    If Not (Rcurrency Is Nothing) Then
'        If Not (Rcurrency.EOF Or Rcurrency.BOF) Then
'            cboCurrency.Clear
'            Rcurrency.MoveFirst
'            Do Until Rcurrency.EOF
'                cboCurrency.AddItem Rcurrency!CurrencyName
'                cboCurrency.ItemData(cboCurrency.NewIndex) = Rcurrency!CURRENCY_ID
'                Rcurrency.MoveNext
'            Loop
'        End If
'    End If

    'CHECK THE Currency Object Under HRCORE
    
    Dim RsT As New ADODB.Recordset
    
    Set RsT = con.Execute("SELECT * FROM CURRENCIES")
    
    If Not (RsT Is Nothing) Then
        If Not (RsT.EOF Or Rcurrency.BOF) Then
            cboCurrency.Clear
            RsT.MoveFirst
            Do Until RsT.EOF
                cboCurrency.AddItem RsT!CurrencyName
                cboCurrency.ItemData(cboCurrency.NewIndex) = RsT!CurrencyID
                If (RsT!IsBaseCurrency = 1) Then CurrID = RsT!CurrencyID
                RsT.MoveNext
            Loop
        End If
        
    End If
    
    Set RsT = Nothing
    
    Exit Sub
ErrHandler:
    MsgBox "an error has occured in " & Me.Name & " Error description " & err.Description
End Sub

Private Sub Form_Resize()
    oSmart.FResize Me
    
    Me.Height = tvwMainheight - 150
    Frame1.Move Frame1.Left, 0, Frame1.Width, tvwMainheight - 150
    lvwDetails.Move lvwDetails.Left, 0, lvwDetails.Width, tvwMainheight - 150
End Sub

Private Sub InitGrid()
    With lvwDetails
        .ColumnHeaders.add , , "Code", 0
        .ColumnHeaders.add , , "Employer", 3000
        .ColumnHeaders.add , , "From", 1200
        .ColumnHeaders.add , , "To", 1200
        .ColumnHeaders.add , , "Reasons for Leaving", 4000
'        .ColumnHeaders.Add , , "Reference Name", 3000
        .ColumnHeaders.add , , "Designation", 2500
        .ColumnHeaders.add , , "Supervisor", 3000
        .ColumnHeaders.add , , "Salary", , vbRightJustify
        .ColumnHeaders.add , , "Phone"
        .ColumnHeaders.add , , "Address"
        .ColumnHeaders.add , , "Comments", 3500
   
                
        .View = lvwReport
    End With
    

End Sub

Public Sub DisplayRecords()

    If SelectedEmployee Is Nothing Then Exit Sub
    lvwDetails.ListItems.Clear
    Call Cleartxt
        
    With rsGlob
        If Not .EOF And Not .BOF Then
                With rs2
                    If .RecordCount > 0 Then
                        .Filter = "employee_id like '" & SelectedEmployee.EmployeeID & "'"
                        If .RecordCount > 0 Then
                            .MoveFirst
                            Do While Not .EOF
                                Set li = lvwDetails.ListItems.add(, , !Code & "", , 5)
                                li.ListSubItems.add , , !Employer & ""
                                li.ListSubItems.add , , Format(!cFrom & "", "dd-MM-yyyy")
                                li.ListSubItems.add , , Format(!cTo & "", "dd-MM-yyyy")
                                li.ListSubItems.add , , !Reasons & "" '
                                li.ListSubItems.add , , !Desig & ""
                                li.ListSubItems.add , , !Super & ""
                                li.ListSubItems.add , , Format(!Salary & "", Cfmt)
                                li.ListSubItems.add , , !Phone & ""
                                li.ListSubItems.add , , !Address & ""
                                li.ListSubItems.add , , !Comments & ""
                                      
                                .MoveNext
                            Loop
                        End If
                        .Filter = adFilterNone
                    End If
                End With
            End If
       End With
End Sub



Private Sub Form_Unload(Cancel As Integer)

    Set rs2 = Nothing
    Set rs5 = Nothing
'    Set Cnn = Nothing
    

    frmMain2.Caption = "Infiniti HRMIS - PDR [Current User:\" & currUser.FullNames & "]"
    
End Sub

Private Sub fraList_DragDrop(Source As Control, X As Single, y As Single)

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
    If frmMain2.cmdEdit.Enabled = True Then
        Call frmMain2.cmdEdit_Click
    End If
End Sub


Private Sub txtCode_Change()
    txtCode.Text = UCase(txtCode.Text)
    txtCode.SelStart = Len(txtCode.Text)
End Sub


Private Sub LastDisb()
With rsGlob
    If Not .EOF Then
        .MoveNext
        If .EOF Then
            cmdLast.Enabled = False
            cmdNext.Enabled = False
            cmdPrevious.Enabled = True
            cmdFirst.Enabled = True
            cmdPrevious.SetFocus
        End If
        .MovePrevious
    End If
    
    cmdPrevious.Enabled = True
    cmdFirst.Enabled = True
End With
End Sub


Private Sub FirstDisb()
    With rsGlob
        If Not .BOF Then
            .MovePrevious
            If .BOF Then
                cmdLast.Enabled = True
                cmdNext.Enabled = True
                cmdPrevious.Enabled = False
                cmdFirst.Enabled = False
                cmdNext.SetFocus
            End If
            .MoveNext
        End If
        
        cmdLast.Enabled = True
        cmdNext.Enabled = True
    End With
End Sub


Private Sub Cleartxt()
    Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
            If i.Name <> "cboCurrency" Then i.Text = ""
        End If
    Next i
    
    dtpFrom.value = Date
    dtpTo.value = Date
    chkBenefit.value = vbUnchecked
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
End Sub

Public Sub EnableCmd()
Dim i As Object
    For Each i In Me
        If TypeOf i Is CommandButton Then
            i.Enabled = True
        End If
    Next i
    
End Sub



Public Sub FirstLastDisb()
cmdLast.Enabled = True
cmdNext.Enabled = True
cmdPrevious.Enabled = True
cmdFirst.Enabled = True
cmdNext.SetFocus
            
With rsGlob
    If Not .BOF = True Then
        .MovePrevious
        If .BOF = True Then
            cmdLast.Enabled = True
            cmdNext.Enabled = True
            cmdPrevious.Enabled = False
            cmdFirst.Enabled = False
            cmdNext.SetFocus
        End If
        .MoveNext
    Else
        cmdLast.Enabled = True
        cmdNext.Enabled = True
        cmdPrevious.Enabled = False
        cmdFirst.Enabled = False
        cmdNext.SetFocus
    End If
    
    If Not .EOF = True Then
        .MoveNext
        If .EOF = True Then
            cmdLast.Enabled = False
            cmdNext.Enabled = False
            cmdPrevious.Enabled = True
            cmdFirst.Enabled = True
            cmdPrevious.SetFocus
        End If
        .MovePrevious
    Else
        cmdLast.Enabled = False
        cmdNext.Enabled = False
        cmdPrevious.Enabled = True
        cmdFirst.Enabled = True
        cmdPrevious.SetFocus
    End If
End With

End Sub




Private Sub txtCode_KeyPress(KeyAscii As Integer)
If Len(Trim(txtCode.Text)) > 19 Then
    Beep
    MsgBox "Can't enter more than 20 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("A") To Asc("Z")
  Case Asc("a") To Asc("z")
  Case Asc("/")
  Case Asc("\")
  Case Asc("?")
  Case Asc(":")
  Case Asc(";")
  Case Asc(",")
  Case Asc("-")
  Case Asc("(")
  Case Asc(")")
  Case Asc("&")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub

Private Sub txtComments_KeyPress(KeyAscii As Integer)
If Len(Trim(txtComments.Text)) > 198 Then
    Beep
    MsgBox "Can't enter more than 200 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("A") To Asc("Z")
  Case Asc("a") To Asc("z")
  Case Asc(" ")
  Case Asc("/")
  Case Asc("\")
  Case Asc("?")
  Case Asc(":")
  Case Asc(";")
  Case Asc(",")
  Case Asc("-")
  Case Asc("(")
  Case Asc(")")
  Case Asc("&")
  Case Asc(".")
  Case Asc("+")
  Case Asc("_")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub

Private Sub txtDesig_KeyPress(KeyAscii As Integer)
If Len(Trim(txtDesig.Text)) > 49 Then
    Beep
    MsgBox "Can't enter more than 50 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("A") To Asc("Z")
  Case Asc("a") To Asc("z")
  Case Asc(" ")
  Case Asc("/")
  Case Asc("\")
  Case Asc("?")
  Case Asc(":")
  Case Asc(";")
  Case Asc(",")
  Case Asc("-")
  Case Asc("(")
  Case Asc(")")
  Case Asc("&")
  Case Asc(".")
  Case Asc("+")
  Case Asc("_")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub

Private Sub txtEmployer_KeyPress(KeyAscii As Integer)
If Len(Trim(txtEmployer.Text)) > 99 Then
    Beep
    MsgBox "Can't enter more than 100 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("A") To Asc("Z")
  Case Asc("a") To Asc("z")
  Case Asc(" ")
  Case Asc("/")
  Case Asc("\")
  Case Asc("?")
  Case Asc(":")
  Case Asc(";")
  Case Asc(",")
  Case Asc("-")
  Case Asc("(")
  Case Asc(")")
  Case Asc("&")
  Case Asc(".")
  Case Asc("+")
  Case Asc("_")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub

Private Sub txtReasons_KeyPress(KeyAscii As Integer)
If Len(Trim(txtReasons.Text)) > 99 Then
    Beep
    MsgBox "Can't enter more than 200 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("A") To Asc("Z")
  Case Asc("a") To Asc("z")
  Case Asc(" ")
  Case Asc("/")
  Case Asc("\")
  Case Asc("?")
  Case Asc(":")
  Case Asc(";")
  Case Asc(",")
  Case Asc("-")
  Case Asc("(")
  Case Asc(")")
  Case Asc("&")
  Case Asc(".")
  Case Asc("+")
  Case Asc("_")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select
End Sub





Private Sub txtSalary_KeyPress(KeyAscii As Integer)
If Len(Trim(txtSalary.Text)) > 9 Then
    Beep
    MsgBox "Can't enter more than 10 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select

End Sub

Private Sub txtSalary_LostFocus()
    On Error Resume Next
    txtSalary.Text = Format(txtSalary.Text & "", Cfmt)
End Sub

Private Sub txtSuper_KeyPress(KeyAscii As Integer)
If Len(Trim(txtSuper.Text)) > 49 Then
    Beep
    MsgBox "Can't enter more than 50 characters", vbExclamation
    KeyAscii = 8
End If

Select Case KeyAscii
  Case Asc("0") To Asc("9")
  Case Asc("A") To Asc("Z")
  Case Asc("a") To Asc("z")
  Case Asc(" ")
  Case Asc("/")
  Case Asc("\")
  Case Asc("?")
  Case Asc(":")
  Case Asc(";")
  Case Asc(",")
  Case Asc("-")
  Case Asc("(")
  Case Asc(")")
  Case Asc("&")
  Case Asc(".")
  Case Is = 8
  Case Else
      Beep
      KeyAscii = 0
End Select

End Sub

Private Function loadACode() As String
    Set rs5 = CConnect.GetRecordSet("SELECT MAX(id) FROM Employment")
    If rs5.EOF = False Then
        If rs5.RecordCount > 0 And Not IsNull(rs5.Fields(0)) Then
            loadACode = "EH" & CStr(rs5.Fields(0) + 1)
        Else
            loadACode = "EH1"
        End If
    Else
        loadACode = "EH1"
    End If
    Set rs5 = Nothing
End Function

Private Function getIndex(data As Long) As Long
    Dim i As Long
    getIndex = -1
    For i = 0 To cboCurrency.ListCount - 1
        If cboCurrency.ItemData(i) = data Then
            getIndex = i
            Exit For
        End If
    Next i
End Function

