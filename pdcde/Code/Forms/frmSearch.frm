VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearch 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      Picture         =   "frmSearch.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Cancel"
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "SELECT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4905
      TabIndex        =   5
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   3480
      TabIndex        =   8
      Top             =   -30
      Width           =   2745
      Begin VB.CommandButton cmdFind 
         Height          =   315
         Left            =   2235
         Picture         =   "frmSearch.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   360
      End
      Begin VB.TextBox txtTo 
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.TextBox txtFrom 
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2115
      End
   End
   Begin VB.ComboBox cboCrieria 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      ItemData        =   "frmSearch.frx":0646
      Left            =   1890
      List            =   "frmSearch.frx":0650
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.ComboBox cboField 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   105
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdRef 
      Height          =   495
      Left            =   3855
      Picture         =   "frmSearch.frx":0663
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Refresh"
      Top             =   4680
      Width           =   495
   End
   Begin MSComctlLib.ListView lstSearch 
      Height          =   3210
      Left            =   105
      TabIndex        =   4
      Top             =   1350
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5662
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16711680
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
   Begin VB.Label Label3 
      Caption         =   "Criteria"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1890
      TabIndex        =   11
      Top             =   30
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Records Found"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   105
      TabIndex        =   10
      Top             =   1095
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Search Field"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   9
      Top             =   30
      Width           =   1575
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cnn As Connection
Dim RsT As Recordset
Dim rst1 As Recordset
Dim LI As ListItem
Dim recordfound As String

Private Sub cboCrieria_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboField_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
    Sel = ""
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim Find As Long
    Dim LI As ListItem
    Dim Field As String
    Set RsT = New Recordset

      
    If Not (cboField.Text = "" Or Me.txtFrom.Text = vbNullString) Then
        If Not cboCrieria.Text = "" Then
            If Me.cboField.ItemData(Me.cboField.ListIndex) = 3 Then
                Field = "OtherNames"
            Else
                Field = cboField.Text
            End If
            lstSearch.ListItems.Clear
            If Not cboCrieria.Text = "Between" And Not cboCrieria.Text = "Like" Then
                
            ElseIf cboCrieria.Text = "Like" Then
              
                With rsGlob
                    If .RecordCount > 0 Then
                        .MoveFirst
                        Dim i As Long
                        i = 1
                        Do While Not .EOF
                            If Trim$(txtFrom) = "" Then Exit Sub
                            .Find "" & Field & " " & cboCrieria.Text & " '" & txtFrom.Text & "%'", , adSearchForward
                            If Not .EOF Then
                            If (lstSearch.ListItems.count > 0) Then
                            Dim gg As String
                            If (lstSearch.ListItems(lstSearch.ListItems.count).Text = !EmpCode) Then
                            GoTo nex
                            End If
                            End If
                            
                                Set LI = lstSearch.ListItems.add(, , !EmpCode & "")
                                LI.ListSubItems.add , , !SurName & ""
                                LI.ListSubItems.add , , !OtherNames & ""
nex:
                                i = i + 1
                                .MoveNext
                            End If
                            
                        Loop
                    End If
                End With
                
                Set rst1 = Nothing
                
                
            Else
                If cboField.Text = "Amount" Then
    '                Set rst1 = cConnect.GetPayData("select * from Employee where " & cboField.Text & " " & cboCrieria & " " & txtFrom.Text & " And " & txtTo.Text & "")
    '                Set rst1 = CConnect.GetRecordSet("select * from Employee where " & cboField.Text & " " & cboCrieria & " " & txtFrom.Text & " And " & txtTo.Text & "")
                    Set rst1 = CConnect.deptFilter("SELECT * FROM pVwRsGlob WHERE seq >= '" & maxCatAccess & "' AND Term <> 1 And " & cboField.Text & " " & cboCrieria & " " & txtFrom.Text & " And " & txtTo.Text & "' ORDER BY EmpCode")
                    
                Else
    '                Set rst1 = cConnect.GetPayData("select * from Employee where " & cboField.Text & " " & cboCrieria & " '" & txtFrom.Text & "' And '" & txtTo.Text & "'")
                    Set rst1 = CConnect.deptFilter("SELECT * FROM pVwRsGlob WHERE seq >= '" & maxCatAccess & "' AND Term <> 1 And " & cboField.Text & " " & cboCrieria & " '" & txtFrom.Text & "' And '" & txtTo.Text & "' ORDER BY EmpCode")
                    
                End If
                
                With rst1
                    If .RecordCount > 0 Then
                        .MoveFirst
                        Do While Not .EOF
                            Set LI = lstSearch.ListItems.add(, , !EmpCode & "")
                            LI.ListSubItems.add , , !SurName & ""
                            LI.ListSubItems.add , , !OtherNames & ""
                            
                            .MoveNext
                        Loop
                    End If
                End With
                
                Set rst1 = Nothing
                
            End If
            
            cmdSelect.SetFocus
        Else
            MsgBox "Please specify the search criteria detail", vbExclamation
        End If
    Else
        If Me.cboCrieria.Text = vbNullString Then
            MsgBox "Please specify the search field detail", vbExclamation
        Else
             Call SRefresh
        End If
    End If
    txtFrom.SetFocus
End Sub

Private Sub cmdRef_Click()
    Call SRefresh
    
End Sub

Private Sub cmdSelect_Click()
On Error GoTo Hell
Sel = ""
    If lstSearch.ListItems.count > 0 Then
        Sel = lstSearch.SelectedItem

        frmMain2.lvwEmp.FindItem(Sel & "").Selected = True
        frmMain2.lvwEmp.SelectedItem.EnsureVisible
        Set SelectedEmployee = AllEmployees.FindEmployee(CLng(frmMain2.lvwEmp.SelectedItem.Tag))
        
        'display the record
         
         If TheLoadedForm.Name = "frmEmployee" Or TheLoadedForm.Name = "frmCheck" Or TheLoadedForm.Name = "frmJProg" Or TheLoadedForm.Name = "frmDisEngagement" Or TheLoadedForm.Name = "frmContacts" Or TheLoadedForm.Name = "frmContract" Or TheLoadedForm.Name = "frmBio" Or TheLoadedForm.Name = "frmKin" Or TheLoadedForm.Name = "frmEdu" Or TheLoadedForm.Name = "frmProf" Or TheLoadedForm.Name = "frmEmployment" Or TheLoadedForm.Name = "frmDDetails" Or TheLoadedForm.Name = "frmRef" Or TheLoadedForm.Name = "frmFamily" Or TheLoadedForm.Name = "frmAwards" Or TheLoadedForm.Name = "frmVisa" Or TheLoadedForm.Name = "frmCasuals" Or TheLoadedForm.Name = "frmEmployeeBanks" Or TheLoadedForm.Name = "frmAssetIssue" Then
             TheLoadedForm.DisplayRecords
         End If
        
        Unload Me
    Else
        MsgBox "No record selected.", vbExclamation
    End If
Exit Sub
Hell:

End Sub

Private Sub Form_Load()
CConnect.CColor Me, MyColor

    With frmSearch.lstSearch
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.add , , "Employee Code", 1500
        .ColumnHeaders.add , , "Surname", 1700
        .ColumnHeaders.add , , "Other Names", 2800
        .ColumnHeaders.add , , "ID No", 1500
        .View = lvwReport
        .GridLines = True
    End With
    
    With frmSearch.cboField
        .Clear
        .AddItem "EmpCode"
        .ItemData(.NewIndex) = 1
        .AddItem "Surname"
        .ItemData(.NewIndex) = 2
        .AddItem "Other Names"
        .ItemData(.NewIndex) = 3
        .Text = "EmpCode"
    End With
    With frmSearch.cboCrieria
        .Clear
        .AddItem "Like"
        .ItemData(.NewIndex) = 1
        .AddItem "="
        .ItemData(.NewIndex) = 2
        .Text = "Like"
    End With
    'load data in the list view
    
   Call SRefresh
        
'    cboCrieria.Text = cboCrieria.List(0)
'    cboField.Text = cboField.List(0)

    Me.Top = (Screen.Height - Height) / 2
    Me.Left = (Screen.Width - Width) / 1.4
    
End Sub

Private Sub lstSearch_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstSearch
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lstSearch_DblClick()
    Call cmdSelect_Click
End Sub

Private Sub txtFrom_Change()
    If txtFrom.Text = "" Then
        cmdFind.Enabled = False
    Else
        cmdFind.Enabled = True
        
    End If
    cmdFind_Click
End Sub

Public Sub SRefresh()
    On Error GoTo Hell
    Dim i As Long
    lstSearch.ListItems.Clear
 
    For i = 1 To AllEmployees.count
        Set LI = frmSearch.lstSearch.ListItems.add(, , AllEmployees.Item(i).EmpCode)
        LI.SubItems(1) = AllEmployees.Item(i).SurName & ""
        LI.SubItems(2) = AllEmployees.Item(i).OtherNames & ""
        LI.SubItems(3) = AllEmployees.Item(i).IdNo & ""
    Next i
    
    txtFrom.Text = ""
    txtTo.Text = ""
    
    Exit Sub
Hell:
    MsgBox "An error has occured when reloading data"
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Not txtFrom.Text = "" Then
        Call cmdFind_Click
        Exit Sub
    End If
End If

    If Len(Trim(txtFrom.Text)) > 20 Then
        Beep
        MsgBox "Can't enter more than 20 characters", vbExclamation
        KeyAscii = 8
    End If
  Select Case KeyAscii
    'Case Asc("vbBack")
    Case Asc("A") To Asc("Z")
    Case Asc("a") To Asc("z")
    Case Asc("0") To Asc("9")
    Case Asc("/")
    Case Asc("-")
    Case Asc("(")
    Case Asc(")")
    Case Asc(" ")
    Case Asc(".")
    'Case Asc("'")
    Case Is = 8
    
    Case Else
    Beep
    KeyAscii = 0
  End Select
  
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
    If Len(Trim(txtTo.Text)) > 20 Then
        Beep
        MsgBox "Can't enter more than 20 characters", vbExclamation
        KeyAscii = 8
    End If
  Select Case KeyAscii
    'Case Asc("vbBack")
    Case Asc("A") To Asc("Z")
    Case Asc("a") To Asc("z")
    Case Asc("0") To Asc("9")
    Case Asc("/")
    Case Asc("-")
    Case Asc("(")
    Case Asc(")")
    Case Asc(" ")
    Case Asc(".")
    Case Is = 8
    
    Case Else
    Beep
    KeyAscii = 0
  End Select
End Sub
