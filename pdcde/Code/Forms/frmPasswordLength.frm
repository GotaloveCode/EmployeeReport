VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmPasswordLength 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password length policy setting"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "0"
      Top             =   2880
      Width           =   270
   End
   Begin MSComCtl2.UpDown upBar 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Top             =   2880
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtNum"
      BuddyDispid     =   196611
      OrigLeft        =   495
      OrigTop         =   2880
      OrigRight       =   735
      OrigBottom      =   3165
      Max             =   60
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Picture         =   "frmPasswordLength.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12375
   End
   Begin VB.Label lblCompanyID 
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   -720
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "characters required in password"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   2910
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Picture         =   "frmPasswordLength.frx":4543
      Stretch         =   -1  'True
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enforce password minimum characters policy setting"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label lblMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "0 minimum characters required in password"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Reducing password characters to 0 inactivates setting"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   4455
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   4455
   End
End
Attribute VB_Name = "frmPasswordLength"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private rsPassword As New ADODB.Recordset '//used to get the database table

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOk_Click()
    prInsertValues
End Sub

Private Sub prInsertValues()
With rsPassword
    If .RecordCount < 1 Then .AddNew
        !Company_ID = 2
        !Minimum_Length = Val(txtNum.Text)
        .Update
        Unload Me
End With
End Sub

Private Sub Form_Load()
CConnect.CColor Me, MyColor
lblCompanyID.Caption = 2
'//create a connection to the database table
    Set rsPassword = CConnect.GetRecordSet("Select * From tblPasswordRule")
    If rsPassword.RecordCount > 0 Then txtNum.Text = rsPassword!Minimum_Length & ""
End Sub

Private Sub lblCompanyID_Change()
If lblCompanyID.Caption = "" Then Exit Sub
'//create a connection to the database table
    Set rsPassword = CConnect.GetRecordSet("Select * From tblPasswordRule")
    If rsPassword.RecordCount > 0 Then txtNum.Text = rsPassword!Minimum_Length & ""
End Sub
