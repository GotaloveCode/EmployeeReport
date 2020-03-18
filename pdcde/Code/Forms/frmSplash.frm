VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   5010
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   4335
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5100
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   4335
      Begin VB.Timer Timer2 
         Left            =   735
         Top             =   2895
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   2550
         Top             =   825
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Personnel Director"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1065
         Left            =   15
         TabIndex        =   2
         Top             =   3510
         Width           =   4305
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   4245
         Left            =   0
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   90
         Width           =   4335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Infiniti Systems Limited"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   795
         TabIndex        =   3
         Top             =   4590
         Width           =   2535
      End
      Begin VB.Label lblByPass 
         BackColor       =   &H00C0E0FF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   3360
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstActivation As Recordset
Dim transform As String
Dim DDays As String
Dim ByDirect As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)
    ByPass = False
    'Unload Me
End Sub

Private Sub Form_Load()
    ByPass = False
    'Frame1.BackColor = 12640480
    'Label1.BackColor = 12640480
    'Label2.BackColor = 12640480
    If App.PrevInstance = True Then
        MsgBox "There's another instance of the system already running!", vbCritical
        End
    End If
    
End Sub

Private Sub Frame1_Click()
    'Unload Me
End Sub

Private Sub Timer1_Timer()
    If ByPass = True Then Exit Sub
    Timer1.Enabled = False
    frmMain2.Visible = True
    Unload Me
End Sub

Public Function TSTransform()
    Dim Pwd As Variant
    Dim Temp As String, PwdChr As Long
    Dim EncryptKey As Long
    Pwd = DDays
    EncryptKey = Int(Sqr(Len(Pwd) * 81)) + 23
    
    For PwdChr = 1 To Len(Pwd)
        Temp = Temp + Chr(Asc(Mid(Pwd, PwdChr, 1)) Xor EncryptKey)
    Next PwdChr
    
    transform = Temp

End Function

Private Sub Timer2_Timer()
    Dim mm As Long
    
    'mm = DateDiff("d", Date, "30 / 6 / 2004")
    If DateDiff("d", Date, "30 / 05 / 2005") > 0 Then
    '    MsgBox "The system will expire in " & mm & " days. Acquire activation."
    
    Else
        Timer2.Enabled = False
        MsgBox "The system has expired. Contact the system vendor.", vbInformation
        Unload Me
        Exit Sub
    End If
    
    'If Timer2.Interval = 1000 Then
        frmMain2.Show
        Timer2.Enabled = False
    'End If

End Sub
