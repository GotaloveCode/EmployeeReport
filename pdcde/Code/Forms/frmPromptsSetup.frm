VERSION 5.00
Begin VB.Form frmPromptsSetup 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Prompts Setup"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5280
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraProjectEnd 
      Appearance      =   0  'Flat
      Caption         =   "Project End"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   24
      Top             =   3720
      Width           =   5055
      Begin VB.TextBox txtProjectEnd 
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
         Height          =   375
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "days before a project ends"
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
         Left            =   2280
         TabIndex        =   26
         Top             =   450
         Width           =   1950
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prompt me"
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
         Left            =   360
         TabIndex        =   25
         Top             =   450
         Width           =   765
      End
   End
   Begin VB.CheckBox chkEnablePrompts 
      Appearance      =   0  'Flat
      Caption         =   "Enable/Disable Prompts"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
   Begin VB.Frame fraEmpTerm 
      Appearance      =   0  'Flat
      Caption         =   "Employment Termination"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   21
      Top             =   2640
      Width           =   5055
      Begin VB.TextBox TxtTermination 
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
         Height          =   375
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prompt me"
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
         Left            =   360
         TabIndex        =   23
         Top             =   450
         Width           =   765
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "days before employment termination"
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
         Left            =   2280
         TabIndex        =   22
         Top             =   450
         Width           =   2640
      End
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   9
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton CmdUpdate 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Frame fraRetirementPrompts 
      Appearance      =   0  'Flat
      Caption         =   "Retirement "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   113
      TabIndex        =   12
      Top             =   4800
      Width           =   5055
      Begin VB.TextBox TxtFemaleRetirementAge 
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
         Height          =   375
         Left            =   3120
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "55"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox TxtMaleRetirementAge 
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
         Height          =   375
         Left            =   3120
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "55"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox TxtRetirement 
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
         Height          =   375
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   5
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Retirement Age For Female Employees"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   1410
         Width           =   2730
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Retirement Age For Male Employees"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   930
         Width           =   2565
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "days before emloyee is due to retire"
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
         Left            =   2280
         TabIndex        =   18
         Top             =   450
         Width           =   2610
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prompt me"
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
         Left            =   240
         TabIndex        =   15
         Top             =   450
         Width           =   765
      End
   End
   Begin VB.Frame fraContractPrompts 
      Appearance      =   0  'Flat
      Caption         =   "Contract Expiry "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   113
      TabIndex        =   11
      Top             =   1560
      Width           =   5055
      Begin VB.TextBox TxtContract 
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
         Height          =   375
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "days before contract expires"
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
         Left            =   2280
         TabIndex        =   17
         Top             =   450
         Width           =   2085
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prompt me"
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
         Left            =   360
         TabIndex        =   14
         Top             =   450
         Width           =   765
      End
   End
   Begin VB.Frame fraProbPrompts 
      Appearance      =   0  'Flat
      Caption         =   "Probation "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   113
      TabIndex        =   10
      Top             =   480
      Width           =   5055
      Begin VB.TextBox TxtProbation 
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
         Height          =   375
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "days to confirmation date"
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
         Left            =   2280
         TabIndex        =   16
         Top             =   450
         Width           =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prompt me "
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
         Left            =   360
         TabIndex        =   13
         Top             =   450
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmPromptsSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------
'PROMPTS SETUP INTERFACE
'ADDED BY: JOHN KIMANI,
'FACILITATES THE MODIFICATION OF PROMPTS TIMELINES
'-----------------------------------------------------------------------------

Option Explicit
Dim sQL As String
Dim RsT As ADODB.Recordset
Dim MyPrompts As Prompts
Dim MyPrompt As Prompt

Private Function ValidateEntries() As Boolean
    On Error GoTo errHandler
    
    ValidateEntries = False
    
    'Validate
    If (TxtContract.Text) = "" Then
        MsgBox "Enter days before contract prompt", vbExclamation, "Error"
        TxtContract.SetFocus
        Exit Function
    End If
    
    If (TxtProbation.Text) = "" Then
        MsgBox "Enter days before probation prompt", vbExclamation, "Error"
        TxtProbation.SetFocus
        Exit Function
    End If
    
    If (TxtTermination.Text) = "" Then
        MsgBox "Enter days before termination prompt", vbExclamation, "Error"
        TxtTermination.SetFocus
        Exit Function
    End If
    
    If (TxtRetirement.Text) = "" Then
        MsgBox "Enter days before retirement prompt", vbExclamation, "Error"
        TxtRetirement.SetFocus
        Exit Function
    End If
    
    If (TxtMaleRetirementAge.Text) = "" Then
        MsgBox "Enter retirement age for male employees", vbExclamation, "Error"
        TxtMaleRetirementAge.SetFocus
        Exit Function
    End If
    
    If (TxtFemaleRetirementAge.Text) = "" Then
        MsgBox "Enter retirement age for female employees", vbExclamation, "Error"
        TxtFemaleRetirementAge.SetFocus
        Exit Function
    End If
    
    'Set Fields
    Set MyPrompt = New Prompt
    MyPrompt.Contract = IIf(Trim(TxtContract.Text) = "", 0, CInt(Trim(TxtContract.Text)))
    MyPrompt.Probation = IIf(Trim(TxtProbation) = "", 0, CInt(Trim(TxtProbation)))
    MyPrompt.Retirement = IIf(Trim(TxtRetirement) = "", 0, CInt(Trim(TxtRetirement)))
    MyPrompt.Termination = IIf(Trim(TxtTermination) = "", 0, CInt(Trim(TxtTermination)))
    MyPrompt.MaleRetirementAge = IIf(Trim(TxtMaleRetirementAge) = "", 0, CInt(Trim(TxtMaleRetirementAge)))
    MyPrompt.FemaleRetirementAge = IIf(Trim(TxtFemaleRetirementAge) = "", 0, CInt(Trim(TxtFemaleRetirementAge)))
    MyPrompt.ProjectEnd = IIf(Trim(txtProjectEnd.Text) = "", 0, CInt(Trim(txtProjectEnd.Text)))
    MyPrompt.EnablePrompts = chkEnablePrompts.value
    
    ValidateEntries = True
    
    Exit Function
errHandler:
    MsgBox "An Error has occured:" & vbCrLf & err.Description, vbExclamation, "Error"
End Function

Private Sub chkEnablePrompts_Click()
    If chkEnablePrompts.value = vbUnchecked Then
        fraContractPrompts.Enabled = False
        fraProbPrompts.Enabled = False
        fraRetirementPrompts.Enabled = False
        fraEmpTerm.Enabled = False
        fraProjectEnd.Enabled = False
        chkEnablePrompts.Caption = "Enable Prompts"
    Else
        fraContractPrompts.Enabled = True
        fraProbPrompts.Enabled = True
        fraRetirementPrompts.Enabled = True
        fraEmpTerm.Enabled = True
        fraProjectEnd.Enabled = True
        chkEnablePrompts.Caption = "Disable Prompts"
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdUpdate_Click()
    On Error GoTo errHandler
    
    If (ValidateEntries) Then
        'Update
        If (MyPrompt.Update) Then
            MsgBox "Update Successful!", vbInformation, "PDR"
        End If
    End If
        
    Exit Sub
errHandler:
    MsgBox "An error has occured:" & err.Description, vbExclamation, "Error"
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    Set MyPrompts = New Prompts
    MyPrompts.GetAllPrompts
    
    TxtContract.Text = MyPrompts.Item(1).Contract
    TxtProbation.Text = MyPrompts.Item(1).Probation
    TxtRetirement.Text = MyPrompts.Item(1).Retirement
    TxtMaleRetirementAge.Text = MyPrompts.Item(1).MaleRetirementAge
    TxtFemaleRetirementAge.Text = MyPrompts.Item(1).FemaleRetirementAge
    TxtTermination.Text = MyPrompts.Item(1).Termination
    txtProjectEnd.Text = MyPrompts.Item(1).ProjectEnd
    chkEnablePrompts.value = MyPrompts.Item(1).EnablePrompts
    chkEnablePrompts_Click
    
    Exit Sub
errHandler:
    MsgBox err.Description, vbExclamation, "PD: Error"
End Sub

Private Function NumericsOnly(ByVal KeyAscii As Integer) As Integer
    On Error GoTo errHandler
    
    NumericsOnly = KeyAscii
    
    If Not IsNumeric(Chr(KeyAscii)) Then
        If Not ((KeyAscii = vbKeyDelete) Or (KeyAscii = vbKeyBack)) Then
            NumericsOnly = 0
        End If
    End If
    
    Exit Function
errHandler:
    MsgBox "An error has occured:" & err.Description, vbExclamation, "Error"
End Function

Private Sub TxtFemaleRetirementAge_KeyPress(KeyAscii As Integer)
    KeyAscii = NumericsOnly(KeyAscii)
End Sub

Private Sub TxtMaleRetirementAge_KeyPress(KeyAscii As Integer)
    KeyAscii = NumericsOnly(KeyAscii)
End Sub

Private Sub TxtProbation_KeyPress(KeyAscii As Integer)
    KeyAscii = NumericsOnly(KeyAscii)
End Sub

Private Sub TxtRetirement_KeyPress(KeyAscii As Integer)
    KeyAscii = NumericsOnly(KeyAscii)
End Sub

Private Sub TxtContract_KeyPress(KeyAscii As Integer)
    KeyAscii = NumericsOnly(KeyAscii)
End Sub
