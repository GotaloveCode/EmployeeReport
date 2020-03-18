VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmEmployeeJDs 
   BorderStyle     =   0  'None
   Caption         =   "Employee JDs"
   ClientHeight    =   7095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Job Description Details:"
      Height          =   4035
      Left            =   90
      TabIndex        =   19
      Top             =   2190
      Width           =   7845
      Begin MSComctlLib.ListView ListView1 
         Height          =   3645
         Left            =   2790
         TabIndex        =   21
         Top             =   360
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   6429
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.TreeView tvwJDCategories 
         Height          =   3600
         Left            =   90
         TabIndex        =   20
         Top             =   360
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   6350
         _Version        =   393217
         Style           =   7
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1845
      Left            =   90
      TabIndex        =   4
      Top             =   180
      Width           =   7845
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4500
         TabIndex        =   23
         Text            =   "Text8"
         Top             =   630
         Width           =   3165
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4500
         TabIndex        =   18
         Text            =   "Text7"
         Top             =   270
         Width           =   3165
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1350
         TabIndex        =   16
         Text            =   "Text6"
         Top             =   1380
         Width           =   1995
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4500
         TabIndex        =   15
         Text            =   "Text5"
         Top             =   1380
         Width           =   3165
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4500
         TabIndex        =   12
         Text            =   "Text4"
         Top             =   1005
         Width           =   3165
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1170
         TabIndex        =   10
         Text            =   "Text3"
         Top             =   1005
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1170
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   630
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1170
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   270
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "Staff Names:"
         Height          =   195
         Left            =   3510
         TabIndex        =   22
         Top             =   675
         Width           =   915
      End
      Begin VB.Label Label7 
         Caption         =   "Supervisor:"
         Height          =   285
         Left            =   3600
         TabIndex        =   17
         Top             =   270
         Width           =   1365
      End
      Begin VB.Label Label6 
         Caption         =   "Job Grade:"
         Height          =   195
         Left            =   3600
         TabIndex        =   14
         Top             =   1425
         Width           =   915
      End
      Begin VB.Label Label5 
         Caption         =   "Date Employed:"
         Height          =   195
         Left            =   90
         TabIndex        =   13
         Top             =   1425
         Width           =   1275
      End
      Begin VB.Label Label4 
         Caption         =   "Location:"
         Height          =   195
         Left            =   3600
         TabIndex        =   11
         Top             =   1050
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Department:"
         Height          =   195
         Left            =   90
         TabIndex        =   9
         Top             =   1050
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "Staff Number:"
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   675
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Job Title:"
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   315
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   6540
      TabIndex        =   3
      Top             =   6465
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   6465
      Width           =   1335
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   6465
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   6465
      Width           =   1215
   End
End
Attribute VB_Name = "frmEmployeeJDs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private selSect As HRCORE.Sector
Private pSectors As HRCORE.Sectors

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo errorHandler
    
    frmMain2.PositionTheFormWithEmpList Me
    
    Set pSectors = New HRCORE.Sectors
    
    
    pSectors.GetActiveSectors
    
    PopulateSectors pSectors
    
    Exit Sub
    
errorHandler:
    
End Sub


Private Sub PopulateSectors(ByVal TheSectors As HRCORE.Sectors)
    Dim i As Long
    Dim nodeX As Node
    
    On Error GoTo errorHandler
    tvwSectors.Nodes.clear
    
    If Not (TheSectors Is Nothing) Then
        For i = 1 To TheSectors.count
            Set nodeX = tvwSectors.Nodes.Add(, , "S" & TheSectors.Item(i).SectorID, TheSectors.Item(i).SectorName)
            nodeX.Tag = TheSectors.Item(i).SectorID
        Next i
    End If
    
    Exit Sub
    
errorHandler:
    MsgBox "An error occurred while Populating the Sectors" & vbNewLine & Err.Description, vbExclamation, TITLES
End Sub
