VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmgeneralfilter 
   Caption         =   "Filter"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdview 
      Caption         =   "View"
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filters"
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin VB.Frame Frame4 
         Caption         =   "Paypoints"
         Height          =   3135
         Left            =   4680
         TabIndex        =   4
         Top             =   2760
         Width           =   2895
         Begin MSComctlLib.ListView ListView3 
            Height          =   2655
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   4683
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Payroll Types"
         Height          =   2415
         Left            =   4680
         TabIndex        =   3
         Top             =   240
         Width           =   2895
         Begin MSComctlLib.ListView ListView2 
            Height          =   2055
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   3625
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Departments"
         Height          =   5655
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4455
         Begin MSComctlLib.ListView ListView1 
            Height          =   5295
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   9340
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
   End
End
Attribute VB_Name = "frmgeneralfilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
