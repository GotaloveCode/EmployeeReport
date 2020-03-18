VERSION 5.00
Begin VB.Form frmdisengagementdateformat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Format for Employee Disengagement"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   1965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "SAVE"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.OptionButton optamerican 
      Caption         =   "MM-DD-YYYY"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.OptionButton optenglish 
      Caption         =   "DD-MM-YYYY"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Value           =   -1  'True
      Width           =   1455
   End
End
Attribute VB_Name = "frmdisengagementdateformat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo err
If optenglish.value = True Then
sql = "exec spSaveDateFormat 'DD-MM-YYYY'"
Else
sql = "exec spSaveDateFormat 'MM-DD-YYYY'"
End If
CConnect.ExecuteSql (sql)
MsgBox "Format Saved", vbInformation
If optenglish.value = True Then
dFormat = "DD-MM-YYYY"
Else
dFormat = "MM-DD-YYYY"
End If
Exit Sub
err:
MsgBox ("The following Error Occured " & err.Description)
dFormat = "DD-MM-YYYY"
End Sub

Private Sub Form_Load()
On Error GoTo err

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
    
    If dFormat = "MM-DD-YYYY" Then
    optamerican.value = True
    Else
    optenglish.value = True
    End If
    Exit Sub
err:
    MsgBox ("The following error occured " & err.Description)
    dFormat = "DD-MM-YYYY"
End Sub
