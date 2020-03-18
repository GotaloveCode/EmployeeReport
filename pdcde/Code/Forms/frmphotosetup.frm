VERSION 5.00
Begin VB.Form frmphotosetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   2715
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   " "
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.CommandButton cmdok 
         Caption         =   "Commit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   3
         Top             =   840
         Width           =   735
      End
      Begin VB.Frame Frame1 
         Caption         =   " "
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   1815
         Begin VB.CheckBox Check1 
            Caption         =   "Active"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "frmphotosetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdok_Click()
On Error GoTo err
Dim sql As String
If Check1.value = vbChecked Then
sql = "exec spUpdatePhotoSetup  1"
Else
sql = "exec spUpdatePhotoSetup  0"
End If
con.Execute (sql)
If Check1.value = vbChecked Then
photoisactive = True
Else
photoisactive = False
End If

Exit Sub
err:
photoisactive = False
MsgBox ("The Following error Occured: " & err.Description)
End Sub

Private Sub Form_Load()
On Error GoTo err
PositionForm Me
Dim rs As New Recordset
Set rs = con.Execute("select active from PhotosSetup")
If Not rs.EOF Then
    If rs!active = 0 Then
      Check1.value = vbUnchecked
    Else
      Check1.value = vbChecked
    End If
Else
  Check1.value = vbUnchecked
End If
Exit Sub
err:
MsgBox ("Error occured trying to get the photos setup")
End Sub
