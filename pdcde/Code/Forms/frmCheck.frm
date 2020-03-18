VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmCheck 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pre-Employment Checklist"
   ClientHeight    =   8025
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Cdl 
      Left            =   2760
      Top             =   6210
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgEmpTool 
      Left            =   2700
      Top             =   2880
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
            Picture         =   "frmCheck.frx":0000
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":0224
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":0336
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame FraEdit 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   8100
      Left            =   0
      TabIndex        =   9
      Top             =   -90
      Width           =   7800
      Begin VB.CheckBox chkAppForm 
         Appearance      =   0  'Flat
         Caption         =   "Application Form/PIF, 4 photographs"
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
         Height          =   270
         Left            =   120
         TabIndex        =   48
         Top             =   3737
         Width           =   3030
      End
      Begin VB.CheckBox chkLetters 
         Appearance      =   0  'Flat
         Caption         =   "Copies of certificates and testimonials"
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
         Height          =   270
         Left            =   120
         TabIndex        =   47
         Top             =   4114
         Width           =   3165
      End
      Begin VB.CheckBox chkMedical 
         Appearance      =   0  'Flat
         Caption         =   "Medical Reports"
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
         Height          =   270
         Left            =   120
         TabIndex        =   46
         Top             =   4491
         Width           =   2385
      End
      Begin VB.CheckBox chkHandbook 
         Appearance      =   0  'Flat
         Caption         =   "Employee hand book"
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
         Height          =   270
         Left            =   120
         TabIndex        =   45
         Top             =   4868
         Width           =   3315
      End
      Begin VB.CheckBox chkAppointment 
         Appearance      =   0  'Flat
         Caption         =   "Signed appointment letter"
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
         Height          =   270
         Left            =   120
         TabIndex        =   44
         Top             =   5245
         Width           =   2385
      End
      Begin VB.CheckBox chkJobTraining 
         Appearance      =   0  'Flat
         Caption         =   "Job Training Plan"
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
         Height          =   270
         Left            =   4080
         TabIndex        =   43
         Top             =   6000
         Width           =   2385
      End
      Begin VB.CheckBox chkJobDesc 
         Appearance      =   0  'Flat
         Caption         =   "Job Description"
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
         Height          =   270
         Left            =   4080
         TabIndex        =   42
         Top             =   5622
         Width           =   2295
      End
      Begin VB.CheckBox chkPIN 
         Appearance      =   0  'Flat
         Caption         =   "PIN Card Copy"
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
         Height          =   270
         Left            =   4080
         TabIndex        =   41
         Top             =   5245
         Width           =   1470
      End
      Begin VB.CheckBox chkNSSF 
         Appearance      =   0  'Flat
         Caption         =   "NSSF Copy"
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
         Height          =   270
         Left            =   4080
         TabIndex        =   40
         Top             =   4868
         Width           =   1185
      End
      Begin VB.CheckBox chkNHIF 
         Appearance      =   0  'Flat
         Caption         =   "NHIF Copy"
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
         Height          =   270
         Left            =   4080
         TabIndex        =   39
         Top             =   4114
         Width           =   1095
      End
      Begin VB.CheckBox chkIDCard 
         Appearance      =   0  'Flat
         Caption         =   "Copy of ID Card"
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
         Height          =   270
         Left            =   4080
         TabIndex        =   38
         Top             =   3737
         Width           =   1545
      End
      Begin VB.CheckBox chkRChecks 
         Appearance      =   0  'Flat
         Caption         =   "Three Reference Checks"
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
         Height          =   270
         Left            =   4080
         TabIndex        =   37
         Top             =   3360
         Width           =   2385
      End
      Begin VB.CheckBox chkBreakdown 
         Appearance      =   0  'Flat
         Caption         =   "Task Breakdown"
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
         Height          =   270
         Left            =   120
         TabIndex        =   36
         Top             =   6000
         Width           =   1545
      End
      Begin VB.CheckBox chkTermsAndCondition 
         Appearance      =   0  'Flat
         Caption         =   "Signed Terms && Conditions"
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
         Height          =   270
         Left            =   120
         TabIndex        =   35
         Top             =   5622
         Width           =   2385
      End
      Begin VB.CheckBox chkBankDetailsForm 
         Appearance      =   0  'Flat
         Caption         =   "Bank Details Form"
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
         Height          =   270
         Left            =   4080
         TabIndex        =   34
         Top             =   4491
         Width           =   1665
      End
      Begin VB.CheckBox chkCV 
         Appearance      =   0  'Flat
         Caption         =   " Application Letter and  CV"
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
         Height          =   270
         Left            =   120
         TabIndex        =   33
         Top             =   3360
         Width           =   2385
      End
      Begin VB.CheckBox chkOrientation 
         Appearance      =   0  'Flat
         Caption         =   "Orientation"
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
         Height          =   270
         Left            =   120
         TabIndex        =   32
         Top             =   6360
         Width           =   2385
      End
      Begin VB.TextBox txtCategory 
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
         Left            =   3330
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   1800
         Width           =   2175
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
         Left            =   3330
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1140
         Width           =   2175
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1140
         Width           =   915
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
         Left            =   1150
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1140
         Width           =   915
      End
      Begin VB.TextBox txtDesig 
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2505
         Width           =   3675
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
         Left            =   6315
         Picture         =   "frmCheck.frx":0878
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Save Record"
         Top             =   6405
         Width           =   495
      End
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
         Left            =   6810
         Picture         =   "frmCheck.frx":097A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Cancel Process"
         Top             =   6405
         Width           =   495
      End
      Begin MSComCtl2.DTPicker dtpDEmployed 
         Height          =   330
         Left            =   1680
         TabIndex        =   8
         Top             =   1800
         Width           =   1500
         _ExtentX        =   2646
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
         CustomFormat    =   "dd, MMM, yyyy"
         Format          =   63045635
         CurrentDate     =   37845
      End
      Begin MSComCtl2.DTPicker dtpDOB 
         Height          =   330
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   1470
         _ExtentX        =   2593
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
         OLEDropMode     =   1
         CustomFormat    =   "dd, MMM, yyyy"
         Format          =   63045635
         CurrentDate     =   37845
         MinDate         =   -36522
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1140
         Width           =   915
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
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   2325
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
         Left            =   1605
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   1605
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   480
         Width           =   1350
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Passport No."
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
         Left            =   2160
         TabIndex        =   29
         Top             =   900
         Width           =   930
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Alien No."
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
         Left            =   1155
         TabIndex        =   27
         Top             =   900
         Width           =   645
      End
      Begin VB.Label Label8 
         Caption         =   "Pre-Employment Check List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   2970
         Width           =   3645
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Designation"
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
         Left            =   120
         TabIndex        =   24
         Top             =   2280
         Width           =   840
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Category"
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
         Left            =   3315
         TabIndex        =   23
         Top             =   1560
         Width           =   675
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Gender"
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
         Left            =   3300
         TabIndex        =   16
         Top             =   900
         Width           =   525
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Date Employed"
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
         Left            =   1695
         TabIndex        =   15
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Date Of Birth"
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
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   945
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "ID No."
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
         Left            =   120
         TabIndex        =   13
         Top             =   900
         Width           =   465
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Other Names"
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
         Left            =   3300
         TabIndex        =   12
         Top             =   270
         Width           =   945
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Sur Name"
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
         Left            =   1605
         TabIndex        =   11
         Top             =   270
         Width           =   690
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Employee Code"
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
         Left            =   120
         TabIndex        =   10
         Top             =   270
         Width           =   1110
      End
   End
   Begin VB.Frame FraList 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   7020
      Left            =   0
      TabIndex        =   17
      Top             =   -90
      Width           =   7440
      Begin MSComctlLib.ListView lvwEmp 
         Height          =   6720
         Left            =   0
         TabIndex        =   18
         Top             =   90
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   11853
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgTree"
         ForeColor       =   0
         BackColor       =   14737632
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
         Left            =   135
         Picture         =   "frmCheck.frx":0A7C
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Add New record"
         Top             =   6360
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
         Left            =   615
         Picture         =   "frmCheck.frx":0B7E
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Edit Record"
         Top             =   6360
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
         Left            =   1095
         Picture         =   "frmCheck.frx":0C80
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Delete Record"
         Top             =   6360
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSComctlLib.ImageList imgTree 
         Left            =   2175
         Top             =   405
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   23
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCheck.frx":1172
               Key             =   "A"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCheck.frx":15C4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCheck.frx":18DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCheck.frx":1D30
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCheck.frx":2182
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCheck.frx":25D4
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCheck.frx":28EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCheck.frx":2C08
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCheck.frx":305A
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCheck.frx":3374
               Key             =   "B"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCheck.frx":37C6
               Key             =   "C"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCheck.frx":3C18
               Key             =   "D"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCheck.frx":406A
               Key             =   "E"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCheck.frx":44BC
               Key             =   "F"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCheck.frx":490E
               Key             =   "G"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCheck.frx":4D60
               Key             =   "H"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCheck.frx":51B2
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCheck.frx":5604
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCheck.frx":5A56
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCheck.frx":5EA8
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCheck.frx":62FA
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCheck.frx":674C
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCheck.frx":6B9E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "CLOSE"
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
         Left            =   8535
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   6420
         Visible         =   0   'False
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NewAcct As Boolean
Dim LastSecID As Long
Dim GenerateID As Boolean
Dim EnterDOB As Boolean
Dim EnterDEmp As Boolean
Dim MStruc As String
'Private CVCopy, AppForm, Testimonials, MReport, EHandbook, SignALetter, IDCard, NHIF, NSSF, PIN, JobDescription, JobT, Orientation, Referees, TaskBreakdown, SignedTC, BankDetailsForm As Integer


Public Sub cmdCancel_Click()
    If PromptSave = True Then
        If MsgBox("Close this window?", vbYesNo + vbQuestion, "Confirm Close") = vbNo Then Exit Sub
    End If
    Unload Me
    Call UnHideMainWindowButtons
End Sub

Public Sub cmdclose_Click()
    Unload Me
End Sub

Public Sub cmdDelete_Click()
    Dim emp As String
     If SelectedEmployee Is Nothing Then Exit Sub
     
    If DSource <> "Local" Then
        MsgBox "You are not allowed to delete employees since this records are from another module.", vbInformation
'        Call cmdCancel_Click
        Exit Sub
    End If
    
    If rsGlob.RecordCount > 0 Then
        resp = MsgBox("Are you sure you want to delete employee - " & SelectedEmployee.EmpCode & "?", vbQuestion + vbYesNo)
        
        If resp = vbNo Then Exit Sub
        emp = SelectedEmployee.EmployeeID
        
        Action = "DELETED EMPLOYEE; EMPLOYEE CODE: " & SelectedEmployee.EmpCode
        
        CConnect.ExecuteSql ("DELETE FROM Employee Where employee_id = '" & emp & "'")
        CConnect.ExecuteSql ("DELETE FROM SEmp Where employee_id = '" & emp & "'")
            
'        ' if AuditTrail = True Then cConnect.ExecuteSql ("INSERT INTO AuditTrail (UserId, DTime, Trans, TDesc, MySection)VALUES('" & CurrentUser & "','" & Date & " " & Time & "','Deleting Employee','" & frmMain2.lvwEmp.SelectedItem & "','Employee')")
        
        rs.Requery

        Set rsGlob = CConnect.GetRecordSet("SELECT e.*, c.Code, c.Description" & _
                " FROM (Employee as e LEFT JOIN SEmp as s ON e.employee_id = SEmp.employee_id) LEFT JOIN " & _
                "CStructure as c ON s.LCode = c.LCode LEFT JOIN ECategory as ec ON e.ECategory = ec.code" & _
                " WHERE (((s.SCode)='" & MStruc & "')) OR (((s.SCode) Is Null)) ORDER BY e.EmpCode")

        'Set rsGlob2 = cConnect.GetRecordSet("SELECT * FROM Employee WHERE Term <> 'Yes' or Term IS NULL ORDER BY EmpCode")
        Set rsGlob2 = CConnect.GetRecordSet("SELECT e.*, c.Code, c.Description FROM (Employee as e " & _
            "LEFT JOIN SEmp as s ON e.employee_id = s.employee_id) LEFT JOIN CStructure as c ON s.LCode" & _
            "= c.LCode LEFT JOIN ECategory as  ec ON e.ECategory = ec.code " & _
            " WHERE e.Term <> 1 AND (((s.SCode)='" & MStruc & "')) OR (((s.SCode) Is Null))" & _
            "  ORDER BY e.EmpCode")
        
        frmMain2.lblECount.Caption = rsGlob.RecordCount
        
'        Call LoadList
        frmMain2.LoadEmployeeList
        Call DisplayRecords
                       
    Else
        MsgBox "No records to be deleted.", vbInformation
    End If

End Sub

Public Sub cmdEdit_Click()
    On Error GoTo ErrHandler
    If lvwEmp.ListItems.count > 0 Then
        With rs
            If .RecordCount > 0 Then
                .MoveFirst
                .Find "employee_id like " & frmMain2.lvwEmp.SelectedItem.Tag, , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    Call DisplayRecords
                    FraList.Visible = False
                    'enabletxt
                    'DisableCmd
                    cmdSave.Enabled = True
                    cmdCancel.Enabled = True
                    txtEmpCode.Locked = True
                    txtSurname.SetFocus
                    SaveNew = False
                Else
                    PSave = True
                    Call cmdCancel_Click
                    PSave = False
                End If
            Else
                PSave = True
                Call cmdCancel_Click
                PSave = False
            End If
        End With
        
    Else
        MsgBox "No records tobe edited", vbInformation
        PSave = True
        Call cmdCancel_Click
        PSave = False
    End If
    Exit Sub
ErrHandler:
    MsgBox err.Description, vbInformation
End Sub

Public Sub cmdNew_Click()
    On Error GoTo ErrHandler
    MsgBox "On this window, please use the edit button as the records" & vbCrLf & "are only expected to be updated.", vbOKOnly + vbInformation, "Use Edit Button"
    Exit Sub
ErrHandler:
End Sub


Public Sub cmdSave_Click()
    If Not currUser Is Nothing Then
        If currUser.CheckRight("Pre_EmploymentChecklist") <> secModify Then
            MsgBox "Yu don't have rigth to add a new record. please liaise with the security admin"
            Exit Sub
        End If
    End If
    If Not (SelectedEmployee Is Nothing) Then
        Call UpdatePreCheckList
        
        ' clear the form
        Call Cleartxt
        
    Else
        MsgBox "please select the Employee"
    End If
End Sub



Private Sub Form_Load()

    frmMain2.PositionTheFormWithEmpList Me
   ' unloak the controls
    'hide all the command button of the main window
    Call HideMainWindowButtons
    Call unlockControls
End Sub


Public Sub DisplayRecords()
On Error GoTo ErrorTrap

Call Cleartxt

If Not (SelectedEmployee Is Nothing) Then
    With SelectedEmployee
        txtEmpCode.Text = .EmpCode
        txtSurname.Text = .SurName
        txtONames.Text = .OtherNames
        txtIDNo.Text = .IdNo
        txtGender.Text = .GenderStr
        dtpDOB.value = .DateOfBirth
        dtpDEmployed.value = .DateOfEmployment
        txtDesig.Text = .position.PositionName
        txtAlien.Text = .AlienNo
        txtPassport.Text = .PassportNo
        txtCategory.Text = .Category.CategoryName
    End With
    Call getExtraInformation
    End If
    
    Exit Sub
    
ErrorTrap:
MsgBox "An error has occurred" & vbNewLine & err.Description, vbExclamation, TITLES
            
End Sub

Function setCheckBoxes(b As String) As Integer
    If b = "True" Then
        setCheckBoxes = 1
    Else
        setCheckBoxes = 0
    End If
End Function


Private Sub Cleartxt()
    ' this procure will be call when user save the  record.
    Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
            i.Text = ""
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
    For Each i In Me
        If TypeOf i Is CommandButton Then
            i.Enabled = False
        End If
    Next i
End Sub

'Public Sub InitGrid()
'    With lvwEmp
'        .ColumnHeaders.Clear
'
'        .ColumnHeaders.Add , , "Employee Code", 1700
'        .ColumnHeaders.Add , , "Surname", 2000
'        .ColumnHeaders.Add , , "Other Names", 3500
'        .ColumnHeaders.Add , , "ID No", 2500
'        .ColumnHeaders.Add , , "Gender"
'        .ColumnHeaders.Add , , "Date of Birth", 2000
'        .ColumnHeaders.Add , , "Date Employed", 2000
'        .ColumnHeaders.Add , , "Division Code", 1700
'        .ColumnHeaders.Add , , "Division Name", 4000
'        .ColumnHeaders.Add , , "Terms"
'        .ColumnHeaders.Add , , "Employee Type", 2500
'        .ColumnHeaders.Add , , "PIN No"
'        .ColumnHeaders.Add , , "N.S.S.F No"
'        .ColumnHeaders.Add , , "N.H.I.F No"
'
'
'        .View = lvwReport
'
'    End With
'End Sub


Private Sub Form_Resize()
'this procedure reposition the form
    oSmart.FResize Me
    Me.FraEdit.Move FraEdit.Left, FraEdit.Top, FraEdit.Width, tvwMainheight - 130
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs = Nothing
    frmMain2.Caption = "Infiniti HRMIS - PDR [Current User:\" & currUser.FullNames & "]"
    Call UnHideMainWindowButtons
End Sub

Private Sub lvwEmp_DblClick()
    If frmMain2.cmdEdit.Enabled = True And frmMain2.fracmd.Visible = True Then
        Call frmMain2.cmdEdit_Click
    End If
    
End Sub

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

'these two procedures are provided to make the module indipedent of the mani module
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

Private Sub unlockControls()
    'unlock the controls
    Dim i As Control
    For Each i In Me
        If TypeOf i Is Frame Then
           i.Enabled = True
        ElseIf TypeOf i Is CommandButton Then
            i.Visible = True
            i.Enabled = True
        End If
    Next i
End Sub

Private Sub getExtraInformation()
    On Error Resume Next
    mySQL = "SELECT * from employees where EmployeeID=" & SelectedEmployee.EmployeeID
    Set rs1 = CConnect.GetRecordSet(mySQL)
    With rs1
        If !TaskBreakdown Then chkBreakdown.value = vbChecked
        If !CVCopy Then chkCV.value = 1
        If !ApplicationForm Then chkAppForm.value = 1
        If !Testimonials Then chkLetters.value = 1
        If !MedicalReport Then chkMedical.value = 1
        If !EmployeeHandbook Then chkHandBook.value = 1
        If !SignedAppointmentLetter Then chkAppointment.value = 1
        If !IDCardCopy Then chkIDCard.value = 1
        If !NHIFCopy Then chkNHIF.value = 1
        If !NSSFCopy Then chkNSSF.value = 1
        If !PINCopy Then chkPIN.value = 1
        If !JobDescription Then chkJobDesc.value = 1
        If !JobTraining Then chkJobTraining.value = 1
        If !RefereesCheck Then chkRChecks.value = 1
        If !Orientation Then chkOrientation.value = 1
        If !BankDetailsForm Then chkBankDetailsForm.value = 1
        If !SignedTermsAndConditions Then chkTermsAndCondition.value = 1
    End With
    
    Set rs1 = Nothing
End Sub

Private Sub UpdatePreCheckList()
    On Error GoTo ErrHandler
    mySQL = ""
    mySQL = "UPDATE Employees SET CVCopy=" & chkCV.value & ", ApplicationForm=" & chkAppForm.value & ",Testimonials=" & chkLetters.value & ",MedicalReport=" & chkMedical.value & ",EmployeeHandbook=" & chkHandBook.value & ", SignedAppointmentLetter=" & chkAppointment.value & _
    ",IDCardCopy=" & chkIDCard.value & ",NHIFCopy=" & chkNHIF.value & ",NSSFCopy=" & chkNSSF.value & ",TaskBreakdown=" & chkBreakdown.value & _
    ",PINCopy=" & chkPIN.value & ",JobDescription =" & chkJobDesc.value & ",JobTraining=" & chkJobTraining.value & ",RefereesCheck=" & chkRChecks.value & _
    ",Orientation=" & chkOrientation.value & ",BankDetailsForm=" & chkBankDetailsForm.value & ",SignedTermsAndConditions=" & chkTermsAndCondition.value & "  WHERE employeeID=" & SelectedEmployee.EmployeeID
    
    CConnect.ExecuteSql mySQL
    currUser.AuditTrail Update, ("Has Update the employee pre_employment checklist of Employee Code " & SelectedEmployee.EmpCode & " " & SelectedEmployee.SurName & " " & SelectedEmployee.OtherNames)
    Set SelectedEmployee = Nothing
    Exit Sub
ErrHandler:
    MsgBox "An error has occur when updating pre_employement check list"
End Sub
