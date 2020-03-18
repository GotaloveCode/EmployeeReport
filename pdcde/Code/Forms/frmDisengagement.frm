VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmDisEngagement 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8100
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7800
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   120
         TabIndex        =   36
         Top             =   6360
         Width           =   6975
         Begin VB.CommandButton cmdCancel 
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
            Left            =   6435
            Picture         =   "frmDisengagement.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Cancel Process"
            Top             =   120
            Width           =   495
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
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
            Left            =   5940
            Picture         =   "frmDisengagement.frx":0102
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Save Record"
            Top             =   120
            Width           =   495
         End
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
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox txtReligion 
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
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox txtTribe 
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
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox txtMstatus 
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
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox txtGender 
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
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox txtEmpCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   390
         Width           =   2085
      End
      Begin VB.TextBox txtAlien 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2490
         Width           =   2085
      End
      Begin VB.TextBox txtPassport 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2070
         Width           =   2085
      End
      Begin VB.TextBox txtONames 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1230
         Width           =   2085
      End
      Begin VB.TextBox txtSurname 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   810
         Width           =   2085
      End
      Begin VB.TextBox txtIDNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1650
         Width           =   2085
      End
      Begin VB.Frame fraDisEngagementInfo 
         Height          =   2625
         Left            =   120
         TabIndex        =   1
         Top             =   3720
         Width           =   6975
         Begin VB.Frame fraTerm 
            Caption         =   "RETIREMENT DETAILS"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1035
            Left            =   120
            TabIndex        =   8
            Top             =   1560
            Width           =   6735
            Begin VB.CheckBox chkTermTrain 
               Caption         =   "Retirement Training"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   120
               TabIndex        =   9
               Top             =   420
               Width           =   1950
            End
            Begin MSComCtl2.DTPicker dtpTerminalDate 
               Height          =   330
               Left            =   3720
               TabIndex        =   10
               Top             =   390
               Width           =   1920
               _ExtentX        =   3387
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
               Format          =   62980099
               CurrentDate     =   37845
            End
            Begin VB.Label Label29 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Training Date:"
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
               Height          =   195
               Left            =   2640
               TabIndex        =   11
               Top             =   465
               Width           =   1020
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "DISENGAGEMENT DETAILS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1305
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   6735
            Begin VB.TextBox txtYY 
               Height          =   300
               Left            =   2520
               TabIndex        =   43
               Top             =   400
               Width           =   615
            End
            Begin VB.TextBox txtMM 
               Height          =   300
               Left            =   2160
               TabIndex        =   42
               Top             =   400
               Width           =   375
            End
            Begin VB.TextBox txtDD 
               Height          =   300
               Left            =   1800
               TabIndex        =   41
               Top             =   400
               Width           =   375
            End
            Begin VB.TextBox txtDisEngagementReferenceNumber 
               Height          =   315
               Left            =   3960
               TabIndex        =   39
               Top             =   840
               Width           =   2625
            End
            Begin VB.CheckBox chkReEngage 
               Appearance      =   0  'Flat
               Caption         =   "Cannot be re-engaged"
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
               Height          =   255
               Left            =   240
               TabIndex        =   4
               Top             =   840
               Visible         =   0   'False
               Width           =   1935
            End
            Begin VB.ComboBox cboTermReasons 
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
               ItemData        =   "frmDisengagement.frx":0204
               Left            =   3960
               List            =   "frmDisengagement.frx":0206
               Style           =   2  'Dropdown List
               TabIndex        =   3
               Top             =   360
               Width           =   2625
            End
            Begin MSComCtl2.DTPicker dtpTerm 
               Height          =   330
               Left            =   3480
               TabIndex        =   5
               Top             =   480
               Visible         =   0   'False
               Width           =   255
               _ExtentX        =   450
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
               Format          =   62980099
               CurrentDate     =   37845
            End
            Begin VB.Label Label10 
               Caption         =   "DD       MM      YYYY"
               Height          =   195
               Left            =   1800
               TabIndex        =   44
               Top             =   195
               Width           =   1695
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Caption         =   "Reference Number"
               Height          =   255
               Left            =   2400
               TabIndex        =   40
               Top             =   840
               Width           =   1455
            End
            Begin VB.Label lblDisDate 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Disengagement Date:"
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
               Height          =   195
               Left            =   480
               TabIndex        =   7
               Top             =   405
               Width           =   1560
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Caption         =   "Reason "
               Height          =   255
               Left            =   3240
               TabIndex        =   6
               Top             =   360
               Width           =   615
            End
         End
      End
      Begin MSComCtl2.DTPicker dtpDOB 
         Height          =   285
         Left            =   1560
         TabIndex        =   23
         Top             =   2895
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   503
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
         OLEDropMode     =   1
         CustomFormat    =   "dd, MMM, yyyy"
         Format          =   62980099
         CurrentDate     =   37845
         MinDate         =   -36522
      End
      Begin VB.Image Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1695
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   180
         Width           =   1635
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "ID No.:"
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
         Height          =   195
         Left            =   210
         TabIndex        =   35
         Top             =   1695
         Width           =   525
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Date Of Birth:"
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
         Height          =   195
         Left            =   210
         TabIndex        =   34
         Top             =   2940
         Width           =   1005
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Gender:"
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
         Height          =   195
         Left            =   3840
         TabIndex        =   33
         Top             =   2085
         Width           =   585
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Staff No.:"
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
         Height          =   195
         Left            =   210
         TabIndex        =   32
         Top             =   435
         Width           =   720
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Surname:"
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
         Height          =   195
         Left            =   210
         TabIndex        =   31
         Top             =   855
         Width           =   690
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Other Names:"
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
         Height          =   195
         Left            =   210
         TabIndex        =   30
         Top             =   1275
         Width           =   1005
      End
      Begin VB.Label Label43 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Passport No.:"
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
         Height          =   195
         Left            =   210
         TabIndex        =   29
         Top             =   2115
         Width           =   990
      End
      Begin VB.Label Label44 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Alien Card No.:"
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
         Height          =   195
         Left            =   210
         TabIndex        =   28
         Top             =   2535
         Width           =   1095
      End
      Begin VB.Label Label30 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Tribe:"
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
         Height          =   195
         Left            =   3870
         TabIndex        =   27
         Top             =   2925
         Width           =   420
      End
      Begin VB.Label Label40 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Marital status:"
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
         Height          =   195
         Left            =   3870
         TabIndex        =   26
         Top             =   2445
         Width           =   1035
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nationality:"
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
         Height          =   195
         Left            =   210
         TabIndex        =   25
         Top             =   3390
         Width           =   825
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Religion:"
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
         Height          =   195
         Left            =   3870
         TabIndex        =   24
         Top             =   3405
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog Cdl 
      Left            =   360
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmDisEngagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private prompted As Boolean
Public Sub ClearMyTexts()
    Cleartxt
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    Set Picture1 = Nothing
    Set SelectedEmployee = Nothing
End Sub


Private Sub cboTermReasons_Click()

    If cboTermReasons.Text = "Retirement" Then
        fraTerm.Visible = True
        fraTerm.Enabled = True
    Else
        'fraTerm.Visible = False
        fraTerm.Enabled = False
    End If

    If cboTermReasons.Text = "Death" Then
        chkReEngage.Enabled = False
        chkReEngage.Visible = True
    Else
        chkReEngage.Enabled = True
        chkReEngage.Visible = True
    End If

End Sub

Public Sub cmdCancel_Click()
    'clear the textboxes
    ClearMyTexts
    
    SaveNew = False
    
    Unload Me
    frmMain2.cmdShowPrompts.Visible = False
End Sub

Public Sub cmdclose_Click()
    Unload Me
End Sub

Public Sub cmdDelete_Click()

End Sub

Public Sub cmdEdit_Click()
    On Error GoTo ErrorTrap
    'Omnis_ActionTag = "E" 'Edits a record in the Omnis database 'monte++
    
    If Not (SelectedEmployee Is Nothing) Then
        'Call DisplayRecords
        'FraList.Visible = False
        enabletxt
        'dtpCDate.Enabled = False
        'DisableCmd
        cmdSave.Enabled = True
        cmdCancel.Enabled = True
        txtEmpCode.SetFocus
        SaveNew = False
        
    Else
        MsgBox "Select the Employee to Edit", vbInformation, TITLES
        'PSave = True    'prompt save
        Call cmdCancel_Click
        'PSave = False
    End If
   
    Exit Sub
    
ErrorTrap:
    MsgBox err.Description, vbExclamation, TITLES
End Sub



Public Sub cmdNew_Click()
    On Error Resume Next
    
    new_Record = True   'flag that a new record is being inserted
    'enable textboxes
    enabletxt
    
    'clear controls
    Cleartxt

'    DisableCmd
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    
    'flags that a new employee is being added
    SaveNew = True
    
    Set Picture1 = Nothing
    
    'disable Disengagement info
    Me.fraDisEngagementInfo.Enabled = False
    
    txtEmpCode.SetFocus
    
End Sub

Public Sub cmdSave_Click()
    On Error GoTo ErrHandler
        
    If Not currUser Is Nothing Then
        If currUser.CheckRight("Disengage") <> secModify Then
            MsgBox "You dont have right to edit the record. Please liaise with the security admin"
            Exit Sub
        End If
    End If
        
    If Not (SelectedEmployee Is Nothing) Then
        Call DisengageEmployee
    End If
    
    cmdCancel.Visible = False
    
    Exit Sub
    
ErrHandler:
    MsgBox "An Error Has Occured:" & vbCrLf & err.Description, vbExclamation, "PDR"
End Sub


Private Sub Form_Load()
    On Error GoTo ErrorHandler
    prompted = False
    dFormat = "DD-MM-YYYY"
    Call DisplayRecords
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    
    cmdSave.Visible = True
    cmdCancel.Visible = True
    'position the form
    frmMain2.PositionTheFormWithEmpList Me
    
    'clear the text
    ClearMyTexts
    Call HideMainWindowButtons
    
    Call LoadDisengagementReasons
    Me.dtpTerm.value = Date
    Me.dtpTerminalDate.value = Date
    
    Dim rsd As New ADODB.Recordset
    Set rsd = CConnect.GetRecordSet("exec spGetDateFormat")
    If Not rsd Is Nothing Then
        If Not rsd.EOF Then
        dFormat = rsd.Fields("format").value
        Else
        dFormat = "DD-MM-YYYY"
        End If
    Else
    dFormat = "DD-MM-YYYY"
    End If
    
    
    ''customize the date format
    Dim dlabel1 As Integer
    Dim dlabel2 As Integer
    dlabel1 = txtDD.Left
    dlabel2 = txtMM.Left
    
    If dFormat = "DD-MM-YYYY" Then
    Label10.Caption = "DD       MM      YYYY"
    txtDD.Left = dlabel1
    txtMM.Left = dlabel2
    Else
    Label10.Caption = "MM       DD      YYYY"
    txtDD.Left = dlabel2
    txtMM.Left = dlabel1
    End If
    
 
    Exit Sub

ErrorHandler:
    MsgBox err.Description, vbExclamation, TITLES
End Sub

Public Sub DisplayRecords()
    Dim ItemX As ListItem
    Dim i As Long
    Dim VisibleOU As HRCORE.OrganizationUnit
    
    On Error GoTo ErrorHandler
    'first clear records
    
    Call Cleartxt
    
    If Not (SelectedEmployee Is Nothing) Then
        With SelectedEmployee
            txtEmpCode.Text = .EmpCode
            txtSurname.Text = .SurName
            txtONames.Text = .OtherNames
            txtIDNo.Text = .IdNo
            txtGender.Text = .GenderStr
            dtpDOB.value = .DateOfBirth
            If .Nationality.Nationality <> "" Then
                txtNationality.Text = .Nationality.Nationality
            Else
                txtNationality.Text = "(Unspecified)"
            End If
            If .Tribe.Tribe <> "" Then
                txtTribe.Text = .Tribe.Tribe
            Else
                txtTribe.Text = "(Unspecified)"
            End If
            txtPassport.Text = .PassportNo
            txtAlien.Text = .AlienNo
            If .Religion.Religion <> "" Then
                txtReligion.Text = .Religion.Religion
            Else
                txtReligion.Text = "(Unspecified)"
            End If

           txtMstatus.Text = .MaritalStatusStr
            
    
            Set Picture1 = Nothing
    
            On Error Resume Next 'this handler is specific to the photos only
            Picture1.Picture = LoadPicture(App.Path & "\Photos\" & CompanyId & "-" & txtEmpCode.Text & ".jpg")
            If Picture1.Picture = 0 Then
                On Error Resume Next
                Picture1.Picture = LoadPicture(App.Path & "\Pic\Gen.jpg")
            End If
            
        End With
    End If
    fraTerm.Visible = True
    fraTerm.Enabled = True
    Exit Sub
ErrorHandler:
    MsgBox "An error has occurred while Displaying Employee Info" & vbNewLine & err.Description, vbInformation, TITLES
End Sub

Function setCheckBoxes(b As String) As Integer
    If b = "True" Then
        setCheckBoxes = 1
    Else
        setCheckBoxes = 0
    End If
End Function

Public Sub Cleartxt()
    Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Then
            i.Text = ""
        ElseIf TypeOf i Is ComboBox Then
            i.ListIndex = -1
        ElseIf TypeOf i Is CheckBox Then
            i.value = vbUnchecked
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

Public Sub DisableCmd()
    Dim i As Object
    
    On Error Resume Next
    For Each i In Me
        If TypeOf i Is CommandButton Then
            i.Enabled = False
        End If
    Next i
    
    
End Sub

Public Sub disabletxt()
Dim i As Object
On Error GoTo ErrHandler
    If TheLoadedForm.Name <> "frmDisEngagement" Then
        For Each i In Me
            
            If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
                i.Locked = True
            End If
        Next i
        
        For Each i In Me
            If TypeOf i Is DTPicker Or TypeOf i Is CheckBox Then
                i.Enabled = False
            End If
        Next i
      ElseIf TheLoadedForm.Name = "frmDisEngagement" Then
        Me.fraDisEngagementInfo.Enabled = True
        Me.Frame12.Enabled = True
        Me.Frame3.Enabled = True
        Me.Frame5.Enabled = True
        Me.chkReEngage.Visible = True
        Me.cboTermReasons.Visible = True
        Me.cmdCancel.Visible = True
        Me.cmdSave.Visible = True
        
        Me.cmdCancel.Enabled = True
        Me.cmdSave.Enabled = True
    End If
    Exit Sub
ErrHandler:
    
End Sub

Public Sub enabletxt()
Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Or TypeOf i Is ComboBox Or TypeOf i Is MaskEdBox Then
            i.Locked = False
        End If
    Next i
    
        
    For Each i In Me
        If TypeOf i Is DTPicker Or TypeOf i Is CheckBox Then
            i.Enabled = True
        End If
    Next i
    
End Sub






Private Sub Form_Resize()
    oSmart.FResize Me
    Me.Frame5.Move Me.Frame5.Left, Frame5.Top, Frame5.Width, tvwMainheight - 220
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs = Nothing
    frmMain2.Caption = "Infiniti HRMIS - PDR [Current User:\" & currUser.FullNames & "]"
    Call UnHideMainWindowButtons
End Sub

'Private Sub imgLoadPic_Click()
'Call cmdPNew_Click
'End Sub

Private Sub lvwEmp_DblClick()
    If frmMain2.cmdEdit.Enabled = True And frmMain2.fracmd.Visible = True Then
        Call frmMain2.cmdEdit_Click
    End If
    
End Sub


'Private Sub GenID()
'Dim NewID As String
'LastSecID = 0
'GenerateID = False
'
'Set rs1 = CConnect.GetRecordSet("SELECT GenID, IDInitials, StartFrom, LastSecID FROM GeneralOpt")
'
'With rs1
'    If .RecordCount > 0 Then
'        If Not IsNull(!GenID) Then
'            If !GenID = "Yes" Then
'                GenerateID = True
'                If IsNull(!IDInitials) Then
'                    If Not IsNull(!LastSecID) Then
'                        NewID = !LastSecID + 1
'                        LastSecID = !LastSecID + 1
'                    Else
'                        NewID = 0
'                        LastSecID = 0
'                    End If
'
'                Else
'                    If Not IsNull(!LastSecID) Then
'                        NewID = !IDInitials & "" & !LastSecID + 1
'                        LastSecID = !LastSecID + 1
'                    Else
'                        NewID = !IDInitials & "" & 0
'                        LastSecID = 0
'                    End If
'
'                End If
'
'                txtEmpCode.Text = NewID
'                txtEmpCode.Locked = True
'                txtSurname.SetFocus
'            Else
'                txtEmpCode.Text = ""
'                txtEmpCode.Locked = False
''                txtEmpCode.SetFocus
'            End If
'        End If
'    End If
'End With
'
'Set rs1 = Nothing
'
'End Sub


Private Sub CColor()
    Dim i As Object
    
    Me.BackColor = MyColor
    
    For Each i In Me
        If TypeOf i Is Label Or TypeOf i Is Form Or TypeOf i Is Frame Then
            i.BackColor = MyColor
        End If
    Next i

End Sub

Public Sub CancelMain()
    If DSource = "Local" Then
        With frmMain2
            .cmdNew.Enabled = False
            .cmdDelete.Enabled = False
            .cmdEdit.Enabled = False
            .cmdSave.Enabled = True
            .cmdCancel.Enabled = True
        End With
    
    End If
End Sub

Private Sub DisengageEmployee()
    On Error GoTo Hell
    
    Dim canbeengaged As Integer
    Dim retVal As Boolean
    Dim OPEN_PERIOD As String, DISENGAGEMENT_PERIOD As String, TermYes As Integer, THIS_EMP As String, m As Integer, y As Integer
    Dim rsDEL As New ADODB.Recordset
    Dim CMD As ADODB.Command
    Dim terminationDate As Date
    Dim DaysThisMonth As Long
    
    If Me.cboTermReasons.ListIndex = -1 Then
        MsgBox "Please Select the reason for disengaging the employee"
        Exit Sub
    End If
    
    If Me.txtDD.Text = "" Or Me.txtMM.Text = "" Or Me.txtYY.Text = "" Then
        MsgBox "Please Select the date of disengagement"
        txtDD.SetFocus
        Exit Sub
    End If
    If Me.cboTermReasons = "Death" Then
        canbeengaged = 0
    Else
        If Me.chkReEngage.value = vbUnchecked Then
            canbeengaged = 1
         Else
            canbeengaged = 0
         End If
    End If
    
    DaysThisMonth = No_of_days(CLng(txtMM.Text), CLng(txtYY.Text))
    
    If (CLng(txtDD.Text) > DaysThisMonth Or txtDD.Text = 0) Then
        MsgBox "The day provided is invalid"
        txtDD.SetFocus
        txtDD.SelStart = 0
        txtDD.SelLength = Len(txtDD.Text)
        Exit Sub
    End If
    
    terminationDate = CDate(txtYY.Text & "-" & txtMM.Text & "-" & txtDD.Text)
    


    'DISENGAGEMENT OPTIONS TO LINK WITH PAYROLL
    mySQL = "SELECT *,PERIOD_END =" & _
    " CASE" & _
    " WHEN PERIODMONTH='January' THEN CONVERT(datetime,CONVERT(VARCHAR(4),PERIODYEAR)+'-01-31')" & _
    " WHEN PERIODMONTH='February' THEN CONVERT(datetime,CONVERT(VARCHAR(4),PERIODYEAR)+'-02-28')" & _
    " WHEN PERIODMONTH='March' THEN CONVERT(datetime,CONVERT(VARCHAR(4),PERIODYEAR)+'-03-31')" & _
    " WHEN PERIODMONTH='April' THEN CONVERT(datetime,CONVERT(VARCHAR(4),PERIODYEAR)+'-04-30')" & _
    " WHEN PERIODMONTH='May' THEN CONVERT(datetime,CONVERT(VARCHAR(4),PERIODYEAR)+'-05-31')" & _
    " WHEN PERIODMONTH='June' THEN CONVERT(datetime,CONVERT(VARCHAR(4),PERIODYEAR)+'-06-30')" & _
    " WHEN PERIODMONTH='July' THEN CONVERT(datetime,CONVERT(VARCHAR(4),PERIODYEAR)+'-07-31')" & _
    " WHEN PERIODMONTH='August' THEN CONVERT(datetime,CONVERT(VARCHAR(4),PERIODYEAR)+'-08-31')" & _
    " WHEN PERIODMONTH='September' THEN CONVERT(datetime,CONVERT(VARCHAR(4),PERIODYEAR)+'-09-30')" & _
    " WHEN PERIODMONTH='October' THEN CONVERT(datetime,CONVERT(VARCHAR(4),PERIODYEAR)+'-10-31')" & _
    " WHEN PERIODMONTH='November' THEN CONVERT(datetime,CONVERT(VARCHAR(4),PERIODYEAR)+'-11-30')" & _
    " WHEN PERIODMONTH='December' THEN CONVERT(datetime,CONVERT(VARCHAR(4),PERIODYEAR)+'-12-31')" & _
    " END FROM tblPeriods WHERE STATUS='Open'"
    
    Set rs = CConnect.GetRecordSet(mySQL)
    If (rs.EOF Or rs.BOF) Then
        MsgBox "There is no open period in payroll" & vbNewLine & "PAYROLL link will be deactiaved", vbInformation
        GoTo ACTUAL_DISENGAGEMENT
    End If
    
    
    
    '--------code to prompt the user about employee's disangagement date.  inserted by kalya on 3-04-2008
    If (prompted = False) Then
      Dim respo As Integer
      respo = MsgBox("IMPORTANT!,Disengagement Date is so crucial. Confirm that it is the exact date. Would you like to continue?", vbYesNo + vbCritical, "Employee Disengagement")
      prompted = True
      If (respo = 6) Then
      ''--------user chose to continue
      Else
      ''-------user chose to go back
      Exit Sub
      End If
    
    End If
    
    
    
    OPEN_PERIOD = rs!PeriodMonth & " " & rs!PeriodYear
    
'    DISENGAGEMENT_PERIOD = MonthName(Month(dtpTerm.value)) & " " & Year(dtpTerm)
    DISENGAGEMENT_PERIOD = MonthName(Month(terminationDate)) & " " & Year(terminationDate)
    THIS_EMP = SelectedEmployee.EmpCode & " - " & SelectedEmployee.SurName & " " & SelectedEmployee.OtherNames
    THIS_EMP = UCase(THIS_EMP)
    
    If (OPEN_PERIOD <> DISENGAGEMENT_PERIOD) Then
'        If (Format(dtpTerm, "yyyy-mm-dd") <= Format(rs!period_end, "yyyy-mm-dd")) Then
        If (Format(terminationDate, "yyyy-mm-dd") <= Format(rs!period_end, "yyyy-mm-dd")) Then
            'This means that the disengagement has been backdated
            TermYes = MsgBox("This disengagement is backdated. The Disengagement period is " & UCase(DISENGAGEMENT_PERIOD) & _
            " while the currently open period in payroll is " & UCase(OPEN_PERIOD) & vbNewLine & "All transactions for " & THIS_EMP & " between the disengagement" & _
            " period and currently open period will be deleted." & vbNewLine & "Are you sure you want to proceed?", vbQuestion + vbYesNo, "PDR: Backdate Disengagement")
            
            'Check IF Proceed = TRUE
            If (TermYes = vbYes) Then
'                If (Year(rs!period_end) = Year(dtpTerm)) Then
'                Set CMD = New ADODB.Command
'                CMD.ActiveConnection = con
'                CMD.CommandType = adCmdStoredProc
'                CMD.CommandText = "prlspDeleteUnwantedTrans"
'                CMD.Execute
                If (Year(rs!period_end) = Year(terminationDate)) Then
                    'Same PeriodYear
'                    For m = (Month(dtpTerm) + 1) To Month(rs!period_end)
                    For m = (Month(terminationDate) + 1) To Month(rs!period_end)
                        mySQL = " begin Delete FROM tblPeriodTransactions Where Employee_ID = " & SelectedEmployee.EmployeeID & " AND Period_Month = " & m & " AND Period_Year = " & Year(dtpTerm) & " end"
                        mySQL = mySQL & " begin DELETE FROM TBLEMPLOYEETRANSACTIONS WHERE EMPLOYEE_ID = " & SelectedEmployee.EmployeeID & " AND Period_Month = " & m & " AND Period_Year = " & Year(dtpTerm) & " end"
                        mySQL = mySQL & " begin    DELETE FROM TBLEMPLOYEEPERIODDETAILS WHERE EMPLOYEE_ID = " & SelectedEmployee.EmployeeID & " AND Period_Month = " & m & " AND Period_Year = " & Year(dtpTerm) & " end"
                        mySQL = mySQL & " begin DELETE FROM TBLSTATEMENTCODES WHERE EMPLOYEE_ID = " & SelectedEmployee.EmployeeID & " AND Period_Month = " & m & " AND Period_Year = " & Year(dtpTerm) & " end"
                        mySQL = mySQL & " begin  DELETE FROM TBLEMPLOYEEBANKNETPAY WHERE EMPLOYEEBANK_ID IN (SELECT EMPLOYEEBANKID FROM EMPLOYEEBANKS INNER JOIN EMPLOYEES ON  EMPLOYEEBANKS.EMPLOYEE_ID=EMPLOYEES.EMPLOYEEID WHERE EMPLOYEES.EMPLOYEEID = " & SelectedEmployee.EmployeeID & " ) AND Period_Month = " & m & " AND Period_Year = " & Year(dtpTerm) & " end"

                   
                    Next m
                         Set rsDEL = CConnect.GetRecordSet(mySQL)
                Else
                    'Different Period Years
'                    For Y = Year(rs!period_end) To Year(dtpTerm)
                    For y = Year(rs!period_end) To Year(terminationDate)
                        If (y = Year(rs!period_end)) Then
                            'THIS IS THE PERIOD YEAR, LOOP n DELETE
                            For m = Month(rs!period_end) To 1
                                mySQL = "Delete FROM tblPeriodTransactions Where Employee_ID = " & SelectedEmployee.EmployeeID & " AND Period_Month = " & m & " AND Period_Year = " & y
                                Set rsDEL = CConnect.GetRecordSet(mySQL)
                            Next m
'                        ElseIf (Y = Year(dtpTerm)) Then
                        ElseIf (y = Year(terminationDate)) Then
                            For m = Month(rs!period_end) To 1
                                mySQL = "Delete FROM tblPeriodTransactions Where Employee_ID = " & SelectedEmployee.EmployeeID & " AND Period_Month = " & m & " AND Period_Year = " & y
                                Set rs = CConnect.GetRecordSet(mySQL)
                            Next m
                        Else
                            'YEAR BETWEEN
                            For m = 12 To 1
                                mySQL = "Delete FROM tblPeriodTransactions Where Employee_ID = " & SelectedEmployee.EmployeeID & " AND Period_Month = " & m & " AND Period_Year = " & y
                            Next m
                        End If
                    Next y
                End If
            Else
                MsgBox "Disengagement of " & THIS_EMP & " has been terminated", vbInformation, "PDR: Disengagement"
                Exit Sub
            End If
        Else
            MsgBox "The disengagement date is after the currently active period in payroll." & vbNewLine & "Disengagement of " & THIS_EMP & " cannot proceed", vbExclamation, "PDR: Disengagement"
            Exit Sub
        End If
    Else
        MsgBox THIS_EMP & " has been disengaged in the currently active period on Payroll." & vbNewLine & "You will need to recalculate his/her payslip in payroll", vbInformation, "PDR: Disengagement"
        prompted = False
    End If
    
    Set rs = Nothing
    'End of Disengagement Link to Payroll

ACTUAL_DISENGAGEMENT:
    
    mySQL = ""
    If UCase(Me.cboTermReasons) = "RETIREMENT" Then
'        mySQL = "UPDATE Employees SET IsDeleted=0,isdisengaged=1, DateOfDisengagement='" & Format(dtpTerm.value, "yyyy-mm-dd") & "', DisEngagementReason='" & cboTermReasons.Text & "',DisEngagementReferenceNumber='" & IIf(Me.txtDisEngagementReferenceNumber.Text = vbNullString, "NULL", Trim(Me.txtDisEngagementReferenceNumber.Text)) & "',IsTrainedOnRetirement=1,trainingdate='" & dtpTerminalDate.value & "', CanBeReEngaged=" & canbeengaged & " Where EmployeeID =" & SelectedEmployee.EmployeeID
        mySQL = "UPDATE Employees SET IsDeleted=0,isdisengaged=1, DateOfDisengagement='" & Format(terminationDate, "yyyy-mm-dd") & "', DisEngagementReason='" & cboTermReasons.Text & "',DisEngagementReferenceNumber='" & IIf(Me.txtDisEngagementReferenceNumber.Text = vbNullString, "NULL", Trim(Me.txtDisEngagementReferenceNumber.Text)) & "',IsTrainedOnRetirement=1,trainingdate='" & Format(dtpTerminalDate.value, "yyyy-mm-dd") & "', CanBeReEngaged=" & canbeengaged & " Where EmployeeID =" & SelectedEmployee.EmployeeID
    Else
'        mySQL = "UPDATE Employees SET IsDeleted=0,isdisengaged=1, DateOfDisengagement='" & Format(dtpTerm.value, "yyyy-mm-dd") & "', DisEngagementReason='" & cboTermReasons.Text & "',DisEngagementReferenceNumber='" & IIf(Me.txtDisEngagementReferenceNumber.Text = vbNullString, "NULL", Trim(Me.txtDisEngagementReferenceNumber.Text)) & "' ,CanBeReEngaged=" & canbeengaged & " Where EmployeeID =" & SelectedEmployee.EmployeeID
        mySQL = "UPDATE Employees SET IsDeleted=0,isdisengaged=1, DateOfDisengagement='" & Format(terminationDate, "yyyy-mm-dd") & "', DisEngagementReason='" & cboTermReasons.Text & "',DisEngagementReferenceNumber='" & IIf(Me.txtDisEngagementReferenceNumber.Text = vbNullString, "NULL", Trim(Me.txtDisEngagementReferenceNumber.Text)) & "' ,CanBeReEngaged=" & canbeengaged & " Where EmployeeID =" & SelectedEmployee.EmployeeID
    End If
    
    ' update the records
    CConnect.ExecuteSql (mySQL)
    currUser.AuditTrail Update, ("Has disengaged Employee code: " & SelectedEmployee.EmpCode & "  " & SelectedEmployee.SurName & " " & SelectedEmployee.OtherNames)
    
    '=====DISENGAGEMENT SHOULD REFLECT ON AN EMPLOYEE'S EMPLOYMENT HISTORY REFLECTING CURRENT EMPLOYER========================
        
    mySQL = "INSERT INTO Employment (employee_id, Employer, Reasons, CFrom, CTo, Desig, Super, Salary, Comments, Phone, Address,code,Isgross,CurrencyId,Benefits)" & _
    " VALUES('" & SelectedEmployee.EmployeeID & "','" & companyDetail.CompanyName & "','" & Me.cboTermReasons & "'," & _
    "'" & Format(SelectedEmployee.DateOfEmployment, "yyyy-MM-dd") & "','" & Format(terminationDate, "yyyy-MM-dd") & "','" & SelectedEmployee.position.PositionName & "',' N/A'," & CCur(SelectedEmployee.BasicPay) & ",' DISENGAGED ','" & CStr(SelectedEmployee.HomeTelephone) & "','" & CStr(SelectedEmployee.HomeAddress) & "','EH'," & 0 & ",1,'" & 0 & "')"
  
    '--AUDIT TRAIL EFFECT
    Action = "ADDED EMPLOYMENT HISTORY; EMPLOYEE CODE: " & SelectedEmployee.EmpCode & "; EMPLOYER: " _
    & companyDetail.CompanyName & "; REASON FOR LEAVING: " & Me.cboTermReasons & "; FROM: " _
    & Format(SelectedEmployee.DateOfEmployment, "dd-MMM-yyyy") & "; TO: " & Format(terminationDate, "dd-MMM-yyyy") _
    & "; DESIGNATION: " & SelectedEmployee.position.PositionName & "; SUPERVISOR: N/A; BASIC SALARY: " _
    & CCur(SelectedEmployee.BasicPay) & "; COMMENTS: DISENGAGEMENT ; PHONE: " & SelectedEmployee.HomeTelephone _
    & "; ADDRESS: " & SelectedEmployee.HomeAddress & "; CODE: EH"
    
    CConnect.ExecuteSql (mySQL)
    currUser.AuditTrail Add_New, (Replace(Action, "'", "''"))
    
    '===========================================END OF EMPLOYMENT HISTORY=======================================================
    
'    rs2.Requery
     MsgBox "The employee has been disengaged successfully", vbOKOnly
    'refresh the listview
    Call frmMain2.LoadEmployeeList
    Call ClearMyTexts
    
    Exit Sub
Hell:
    MsgBox "An Error Has Occured:" & vbNewLine & err.Description, vbExclamation, "PDR"
End Sub

Private Sub HideMainWindowButtons()
    Dim i As Object
    For Each i In frmMain2
        If TypeOf i Is CommandButton Then
            i.Visible = False
        End If
    Next i

End Sub

Private Sub UnHideMainWindowButtons()
    Dim i As Object
    For Each i In frmMain2
        If TypeOf i Is CommandButton Then
            i.Visible = True
        End If
    Next i
End Sub

Private Sub LoadDisengagementReasons()
    Dim DisReason As New DisengagementReasons
    Dim i As Long
    DisReason.GetallDisengagementReasons
    cboTermReasons.Clear
    
    For i = 1 To DisReason.count
        cboTermReasons.AddItem DisReason.Item(i).Reason
    Next i
    
End Sub



Private Sub lblDisDate_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    lblDisDate.ToolTipText = "Enter exact date the employee was terminated"
End Sub

Private Sub txtDD_Change()
    If Len(txtDD.Text) = 2 Then
        KeyAscii = 0
       '' txtMM.SetFocus
    End If
End Sub

Private Sub txtDD_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Or KeyAscii = vbKeyTab Then
    ElseIf IsNumeric(Chr(KeyAscii)) Then
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtDD_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    txtDD.ToolTipText = "Enter exact date the employee was terminated"
End Sub

Private Sub txtMM_Change()
    If Len(txtMM.Text) = 2 Then
        KeyAscii = 0
        txtYY.SetFocus
    End If
End Sub

Private Sub txtMM_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Or KeyAscii = vbKeyTab Then
    ElseIf IsNumeric(Chr(KeyAscii)) Then
    Else
        KeyAscii = 0
    End If

End Sub

Private Sub txtMM_LostFocus()
On Error GoTo err
    If txtMM.Text > 13 Or txtMM.Text = 0 Then
        MsgBox "The month number is invalid"
        txtMM.SetFocus
        txtMM.SelStart = 0
        txtMM.SelLength = Len(txtMM.Text)
    End If
    Exit Sub
err:
End Sub

Private Sub txtYY_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Or KeyAscii = vbKeyTab Then
    ElseIf IsNumeric(Chr(KeyAscii)) Then
    Else
        KeyAscii = 0
    End If
End Sub

Private Function No_of_days(lngMonth As Long, lngYear As Long) As Long
    Dim isleapyear  As Boolean
    On Error GoTo ErrHandler
    
    If lngYear Mod 4 <> 0 Then
        isleapyear = False
    Else
        isleapyear = True
    End If

        Select Case isleapyear
            Case False
                If lngMonth = 1 Then
                    No_of_days = 31
                ElseIf lngMonth = 2 Then
                    No_of_days = 28
                ElseIf lngMonth = 3 Then
                    No_of_days = 31
                ElseIf lngMonth = 4 Then
                    No_of_days = 30
                ElseIf lngMonth = 5 Then
                    No_of_days = 31
                ElseIf lngMonth = 6 Then
                    No_of_days = 30
                ElseIf lngMonth = 7 Then
                    No_of_days = 31
                ElseIf lngMonth = 8 Then
                    No_of_days = 31
                ElseIf lngMonth = 9 Then
                    No_of_days = 30
                ElseIf lngMonth = 10 Then
                    No_of_days = 31
                ElseIf lngMonth = 11 Then
                    No_of_days = 30
                ElseIf lngMonth = 12 Then
                    No_of_days = 31
                End If
            Case True
                If lngMonth = 1 Then
                    No_of_days = 31
                ElseIf lngMonth = 2 Then
                    No_of_days = 29
                ElseIf lngMonth = 3 Then
                    No_of_days = 31
                ElseIf lngMonth = 4 Then
                    No_of_days = 30
                ElseIf lngMonth = 5 Then
                    No_of_days = 31
                ElseIf lngMonth = 6 Then
                    No_of_days = 30
                ElseIf lngMonth = 7 Then
                    No_of_days = 31
                ElseIf lngMonth = 8 Then
                    No_of_days = 31
                ElseIf lngMonth = 9 Then
                    No_of_days = 30
                ElseIf lngMonth = 10 Then
                    No_of_days = 31
                ElseIf lngMonth = 11 Then
                    No_of_days = 30
                ElseIf lngMonth = 12 Then
                    No_of_days = 31
                End If
        End Select
        Exit Function
ErrHandler:
End Function

