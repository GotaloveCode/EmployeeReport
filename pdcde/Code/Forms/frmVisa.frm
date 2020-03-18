VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmVisa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Visa Details"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   Icon            =   "frmVisa.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   7830
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
            Picture         =   "frmVisa.frx":0442
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisa.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisa.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisa.frx":0778
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   7800
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   7800
      Begin MSComDlg.CommonDialog Cdl 
         Left            =   5610
         Top             =   4560
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
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Move to the First employee"
         Top             =   5400
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
         Left            =   6270
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   5400
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
         Left            =   600
         TabIndex        =   13
         ToolTipText     =   "Move to the Previous employee"
         Top             =   5400
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
         Left            =   1080
         TabIndex        =   14
         ToolTipText     =   "Move to the Next employee"
         Top             =   5400
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
         Left            =   2775
         Picture         =   "frmVisa.frx":0CBA
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Picture         =   "frmVisa.frx":0DBC
         Style           =   1  'Graphical
         TabIndex        =   17
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
         Picture         =   "frmVisa.frx":0EBE
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Delete Record"
         Top             =   5400
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
         Left            =   1560
         TabIndex        =   15
         ToolTipText     =   "Move to the Last employee"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame fraDetails 
         Appearance      =   0  'Flat
         Caption         =   "Expatriates Visa Details"
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
         Height          =   2985
         Left            =   705
         TabIndex        =   21
         Top             =   1680
         Visible         =   0   'False
         Width           =   6405
         Begin VB.TextBox txtClass 
            Height          =   285
            Left            =   3360
            TabIndex        =   34
            Top             =   600
            Width           =   1335
         End
         Begin VB.ComboBox cboCurrency 
            Height          =   315
            Left            =   3960
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   1260
            Width           =   975
         End
         Begin VB.TextBox txtTitle 
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
            Left            =   1785
            TabIndex        =   9
            Top             =   1860
            Width           =   4100
         End
         Begin VB.TextBox txtNationality 
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
            Left            =   150
            TabIndex        =   8
            Top             =   1860
            Width           =   1530
         End
         Begin VB.TextBox txtRcptNo 
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
            Left            =   5070
            TabIndex        =   7
            Top             =   1260
            Width           =   1260
         End
         Begin VB.TextBox txtYear 
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
            Left            =   5055
            TabIndex        =   3
            Top             =   600
            Width           =   1245
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
            Picture         =   "frmVisa.frx":13B0
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Cancel Process"
            Top             =   2310
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
            Left            =   4890
            Picture         =   "frmVisa.frx":14B2
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Save Record"
            Top             =   2310
            Width           =   495
         End
         Begin VB.TextBox txtKES 
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
            Left            =   2895
            TabIndex        =   6
            Top             =   1260
            Width           =   1005
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
            Height          =   285
            Left            =   135
            TabIndex        =   1
            Top             =   600
            Width           =   1530
         End
         Begin VB.TextBox txtPermit 
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
            Left            =   1770
            TabIndex        =   2
            Top             =   600
            Width           =   1515
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
            Height          =   315
            Left            =   135
            TabIndex        =   4
            Top             =   1260
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
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
            CustomFormat    =   "dd, MMM, yyyy"
            Format          =   64290819
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
            Height          =   315
            Left            =   1530
            TabIndex        =   5
            Top             =   1260
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
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
            CustomFormat    =   "dd, MMM, yyyy"
            Format          =   64290819
            CurrentDate     =   37673
            MinDate         =   21916
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
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
            Left            =   3960
            TabIndex        =   33
            Top             =   1020
            Width           =   660
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Title"
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
            Left            =   1785
            TabIndex        =   31
            Top             =   1620
            Width           =   300
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
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
            Left            =   135
            TabIndex        =   30
            Top             =   1620
            Width           =   765
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Receipt No"
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
            Left            =   5070
            TabIndex        =   29
            Top             =   1020
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Class"
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
            Left            =   3390
            TabIndex        =   28
            Top             =   375
            Width           =   375
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Years"
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
            Left            =   5055
            TabIndex        =   27
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
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
            Left            =   1530
            TabIndex        =   26
            Top             =   1020
            Width           =   180
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
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
            Left            =   2895
            TabIndex        =   25
            Top             =   1020
            Width           =   555
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "File Reference No."
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
            TabIndex        =   24
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Work Permit No"
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
            Left            =   1770
            TabIndex        =   23
            Top             =   360
            Width           =   1110
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
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
            Left            =   135
            TabIndex        =   22
            Top             =   1020
            Width           =   360
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
Attribute VB_Name = "frmVisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

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

Private Sub cmdclose_Click()
    Unload Me
End Sub

Public Sub cmdDelete_Click()
    Dim resp As String
    'On Error GoTo errHandler
    If SelectedEmployee Is Nothing Then Exit Sub
    
    If Not currUser Is Nothing Then
        If currUser.CheckRight("Visa") <> secModify Then
            MsgBox "You dont have right to delete the record. Please liaise with the security admin"
            Exit Sub
        End If
     End If
     
    If lvwDetails.ListItems.count > 0 Then
        resp = MsgBox("Are you sure you want to delete  " & lvwDetails.SelectedItem & " from the records?", vbQuestion + vbYesNo)
        If resp = vbNo Then
            Exit Sub
        End If
        
        CConnect.ExecuteSql ("DELETE FROM Visa WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & lvwDetails.SelectedItem & "'")
    
        rs2.Requery
        
        Call DisplayRecords
            
    Else
        MsgBox "You have to select the visa you would like to delete.", vbInformation
                
    End If
    'Exit Sub
    'errHandler:
End Sub

Private Sub cmdDone_Click()
    PSave = True
    Call cmdCancel_Click
    PSave = False
End Sub

Public Sub cmdEdit_Click()

    If Not currUser Is Nothing Then
        If currUser.CheckRight("Visa") <> secModify Then
            MsgBox "You dont have right to modify the record. Please liaise with the security admin"
            Exit Sub
        End If
     End If
     
    If lvwDetails.ListItems.count < 1 Then
        MsgBox "You have to select the visa you would like to edit.", vbInformation
        PSave = True
        Call cmdCancel_Click
        PSave = False
        Exit Sub
    End If
    
    Set rs3 = CConnect.GetRecordSet("SELECT * FROM Visa WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & lvwDetails.SelectedItem & "'")
    
    With rs3
        If .RecordCount > 0 Then
            txtCode.Text = !Code & ""
            txtPermit.Text = !Permit & ""
            txtClass.Text = !Class & ""
            txtYear.Text = !Years & ""
            dtpFrom.value = !cFrom & ""
            dtpTo.value = !cTo & ""
            txtKES.Text = Format(!KES & "", "#,###,##0.00")
            txtRcptNo.Text = !RcptNo & ""
            txtNationality.Text = !Nationality & ""
            txtTitle.Text = !Title & ""
            txtNationality.Locked = True
            
            If Not IsNull(!CurrencyID) Then
                cboCurrency.ListIndex = getIndex(!CurrencyID)
            Else
                cboCurrency.ListIndex = -1
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

    txtPermit.SetFocus
     
    txtClass.Locked = False
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
    Dim rsCheckEmpTerms As ADODB.Recordset
    
    If Not currUser Is Nothing Then
        If currUser.CheckRight("Visa") <> secModify Then
            MsgBox "You dont have right to add new record. Please liaise with the security admin"
            Exit Sub
        End If
     End If
     
    If Not (SelectedEmployee Is Nothing) Then
        If Not SelectedEmployee.EmploymentTerm.IsExpatriate Then
            MsgBox "Select an expatriate employee."
            Exit Sub
        End If
    Else
        Exit Sub
   End If

   '' Set rsCheckEmpTerms = CConnect.GetRecordSet("SELECT * FROM EmploymentTerms WHERE matchToExpertriate=4 and EmpTermName like'" & SelectedEmployee.EmploymentTerm.EmpTermName & "'")
    Set rsCheckEmpTerms = CConnect.GetRecordSet("exec pdrexpatriatematch '" & SelectedEmployee.EmploymentTerm.EmpTermName & "'")
    
    If Not (rsCheckEmpTerms Is Nothing) Then
        If rsCheckEmpTerms.RecordCount = 0 Then MsgBox "Please set up the expertriate details": Exit Sub
    End If
    
    dtpFrom.value = Date
    Call DisableCmd
    txtCode.Text = loadACode
    txtCode.Locked = False
    txtPermit.Text = ""
    txtClass.Text = ""
    txtYear.Text = ""
    txtKES.Text = ""
    txtRcptNo.Text = ""
    If Not (SelectedEmployee Is Nothing) Then
        txtNationality.Text = SelectedEmployee.Nationality.Nationality
    End If
    txtNationality.Locked = True
    txtTitle.Text = ""
    dtpFrom.value = Date
    dtpTo.value = Date
    fraDetails.Visible = True
    cmdCancel.Enabled = True
    SaveNew = True
    cmdSave.Enabled = True
    'txtCode.Locked = True
    txtPermit.SetFocus
    dtpFrom.value = Date
    dtpTo.value = Date
    
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
    Dim CuRindex As Long
    If SelectedEmployee Is Nothing Then
        MsgBox "Please Select employee"
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    If txtCode.Text = "" Then
        MsgBox "Enter the File No.", vbExclamation
        txtCode.SetFocus
        Exit Sub
    End If
    
    If dtpFrom.value > dtpTo.value Then
        MsgBox "Enter the valid visa start and end dates.", vbInformation
        dtpFrom.SetFocus
        Exit Sub
    End If
    
    If txtYear.Text = "" Then
        txtYear.Text = 0
    End If
    
        If SaveNew = True Then
            Set rs4 = CConnect.GetRecordSet("SELECT * FROM Visa WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & txtCode.Text & "'")
                    
            With rs4
                If .RecordCount > 0 Then
                    MsgBox "File No. already exists. Enter another one.", vbInformation
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
        CuRindex = cboCurrency.ItemData(cboCurrency.ListIndex)
        CConnect.ExecuteSql ("DELETE FROM Visa WHERE employee_id = '" & SelectedEmployee.EmployeeID & "' AND Code = '" & txtCode.Text & "'")
        
        mySQL = "INSERT INTO Visa (employee_id, Code, Permit, Class, Years, CFrom, CTo, KES, RcptNo, Nationality, Title,CurrencyId)" & _
                            " VALUES('" & SelectedEmployee.EmployeeID & "','" & txtCode.Text & "','" & txtPermit.Text & "','" & txtClass.Text & "'," & txtYear.Text & "," & _
                            "'" & Format(dtpFrom.value, Dfmt) & "','" & Format(dtpTo.value, Dfmt) & "','" & txtKES.Text & "','" & txtRcptNo.Text & "','" & txtNationality.Text & "','" & txtTitle.Text & "'," & CuRindex & " )"
        
        'MsgBox mySQL
        Action = "ADDED EMPLOYEE'S VISA DETAILS; EMPLOYEE CODE: " _
        & SelectedEmployee.EmpCode & "; CODE: " & txtCode.Text _
        & "; PERMIT: " & txtPermit.Text & "; CLASS: " _
        & txtClass.Text & "; YEAR: " & txtYear.Text & "; FROM: " _
        & Format(dtpFrom.value, "dd-MMM-yyyy") _
        & "; TO: " & Format(dtpTo.value, "dd-MMM-yyyy") _
        & "; PAY: " & txtKES.Text & "; NATIONALITY: " _
        & txtNationality.Text & "; TITLE: " & txtTitle.Text
        
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
            txtPermit.SetFocus
        End If
        
    Exit Sub
ErrHandler:
    MsgBox err.Description, vbInformation
    
End Sub


Private Sub Combo1_Change()

End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler
    
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
        
    Set rs2 = CConnect.GetRecordSet("SELECT * FROM Visa ORDER BY Code")
    
    With rsGlob
        If .RecordCount < 1 Then
            Call DisableCmd
            Exit Sub
        End If
    End With
    
    Call DisplayRecords
    
    cmdFirst.Enabled = False
    cmdPrevious.Enabled = False
    
    'On Error Resume Next
    Rcurrency.Requery
    If Not (Rcurrency Is Nothing) Then
        If Not (Rcurrency.EOF Or Rcurrency.BOF) Then
            
            cboCurrency.Clear
            Rcurrency.MoveFirst
            Do Until Rcurrency.EOF
                cboCurrency.AddItem Rcurrency!CURRENCY_NAME
                cboCurrency.ItemData(cboCurrency.NewIndex) = Rcurrency!CURRENCY_ID
                Rcurrency.MoveNext
            Loop
        End If
    End If
    
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
        .ColumnHeaders.Clear
        .ColumnHeaders.add , , "Code", 0
        .ColumnHeaders.add , , "Permit", 900
        .ColumnHeaders.add , , "Class", 900
        .ColumnHeaders.add , , "Years", 1400
        .ColumnHeaders.add , , "From", 1400
        .ColumnHeaders.add , , "To", 1400
        .ColumnHeaders.add , , "Amount", 1400
        .ColumnHeaders.add , , "Receipt No.", 1400
        .ColumnHeaders.add , , "Nationality", 1400
        .ColumnHeaders.add , , "Title", 3500
        .View = lvwReport
    End With
End Sub

Public Sub DisplayRecords()
    On Error GoTo ErrHandler
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
                            li.ListSubItems.add , , !Permit & ""
                            li.ListSubItems.add , , !Class & ""
                            li.ListSubItems.add , , !Years & ""
                            li.ListSubItems.add , , !cFrom & ""
                            li.ListSubItems.add , , !cTo & ""
                            li.ListSubItems.add , , Format(!KES & "", "#,###,##0.00")
                            li.ListSubItems.add , , !RcptNo & ""
                            li.ListSubItems.add , , !Nationality & ""
                            li.ListSubItems.add , , !Title & ""
                           .MoveNext
                        Loop
                    End If
                    .Filter = adFilterNone
                End If
            End With
        End If
    End With
    Exit Sub
ErrHandler:
    MsgBox "An error has occur when displaying records: " & err.Description
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Set rs2 = Nothing
    Set rs5 = Nothing
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
    lvwDetails.ListItems.Clear
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

Private Sub txtKES_LostFocus()
    On Error Resume Next
    txtKES.Text = Format(txtKES.Text & "", "#,###,##0.00")
End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
    If Len(Trim(txtYear.Text)) > 2 Then
        Beep
        MsgBox "Can't enter more than 3 characters", vbExclamation
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

Private Function loadACode() As String
    Set rs5 = CConnect.GetRecordSet("SELECT MAX(id) FROM Visa")
    If rs5.EOF = False Then
        If rs5.RecordCount > 0 And Not IsNull(rs5.Fields(0)) Then
            loadACode = "VS" & CStr(rs5.Fields(0) + 1)
        Else
            loadACode = "VS1"
        End If
    Else
        loadACode = "VS1"
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
