VERSION 5.00
Begin VB.Form frmNextOfKinGuardian 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Next Of Kin: Guardian Details"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6780
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkDisAssociate 
      Caption         =   "Dis-Associate Next Of Kin From Guardian"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   3375
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5520
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Frame fraGuardian 
      Caption         =   "Guardian Details:"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.TextBox txtRelationship 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Top             =   1320
         Width           =   4095
      End
      Begin VB.TextBox txtIDNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox txtFullNames 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         TabIndex        =   1
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Relationship To Next of Kin:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1335
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ID /Passport Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   855
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Guardian's Full Names:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   375
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmNextOfKinGuardian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TheNextOfKin As HRCORE.NextOfKin

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errorHandler
    If TheNextOfKin Is Nothing Then
        Set TheNextOfKin = New HRCORE.NextOfKin
    End If
    
    If chkDisAssociate.Value = vbChecked Then
        'clear the Guardian Info
        TheNextOfKin.GuardianFullNames = ""
        TheNextOfKin.GuardianIDNo = ""
        TheNextOfKin.GuardianRelationship = ""
    Else
        If Len(Trim(Me.txtFullNames.Text)) > 0 Then
            TheNextOfKin.GuardianFullNames = Trim(Me.txtFullNames.Text)
        Else
            MsgBox "Supply the Full Names of the Guardian", vbExclamation, TITLES
            Me.txtFullNames.SetFocus
            Exit Sub
        End If
        
        If Len(Trim(Me.txtIDNo.Text)) > 0 Then
            TheNextOfKin.GuardianIDNo = Trim(Me.txtIDNo.Text)
        Else
            MsgBox "Supply the ID or Passport Number of the Guardian", vbExclamation, TITLES
            Me.txtIDNo.SetFocus
            Exit Sub
        End If
        
        TheNextOfKin.GuardianRelationship = Trim(Me.txtRelationship.Text)
        
    End If
    
    're-update
    Set TempNextOfKin = TheNextOfKin
    Unload Me
    
    Exit Sub
    
errorHandler:
    MsgBox "An error has occurred while updating the Guardian Information" & vbNewLine & Err.Description, vbExclamation, TITLES
        
End Sub

Private Sub Form_Load()
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    chkDisAssociate.Value = vbUnchecked
    PopulateGuardianInfo
End Sub

Public Sub PopulateGuardianInfo()
    On Error GoTo errorHandler
    ClearControlsGuardian
    Set TheNextOfKin = TempNextOfKin
    If TheNextOfKin Is Nothing Then Exit Sub
    Me.txtFullNames.Text = TheNextOfKin.GuardianFullNames
    Me.txtIDNo.Text = TheNextOfKin.GuardianIDNo
    Me.txtRelationship = TheNextOfKin.GuardianRelationship
    
    Exit Sub
    
errorHandler:
    MsgBox "An error has occurred while populating the Guardian Information" & vbNewLine & Err.Description, vbExclamation, TITLES
End Sub

Private Sub ClearControlsGuardian()
    On Error GoTo errorHandler
    Me.txtFullNames.Text = ""
    Me.txtIDNo.Text = ""
    Me.txtRelationship.Text = ""
    
    Exit Sub
    
errorHandler:
    
End Sub
