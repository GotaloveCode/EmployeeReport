VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPopUp 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee List"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   Icon            =   "frmPopUp.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
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
      Picture         =   "frmPopUp.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Cancel"
      Top             =   5610
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
      Top             =   5610
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   3480
      TabIndex        =   8
      Top             =   -30
      Width           =   2745
      Begin VB.CommandButton cmdFind 
         Height          =   315
         Left            =   2235
         Picture         =   "frmPopUp.frx":0544
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
      ItemData        =   "frmPopUp.frx":0646
      Left            =   1890
      List            =   "frmPopUp.frx":065F
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
      ItemData        =   "frmPopUp.frx":0683
      Left            =   105
      List            =   "frmPopUp.frx":0685
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdRef 
      Height          =   495
      Left            =   3855
      Picture         =   "frmPopUp.frx":0687
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Refresh"
      Top             =   5610
      Width           =   495
   End
   Begin MSComctlLib.ListView lstSearch 
      Height          =   4140
      Left            =   105
      TabIndex        =   4
      Top             =   1350
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7303
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   -2147483643
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
   Begin VB.Label lblCount 
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
      Left            =   1395
      TabIndex        =   13
      Top             =   1065
      Width           =   720
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
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
      BackStyle       =   0  'Transparent
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
      Top             =   1065
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
Attribute VB_Name = "frmPopUp"
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


Private Sub cboCrieria_Click()
  If cboCrieria.Text = "Between" Then
    txtTo.Visible = True
  Else
    txtTo.Visible = False
  End If
  
End Sub

Private Sub cboCrieria_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboField_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
    Sel = ""
    frmPopUp.Visible = False
    
End Sub

Private Sub cmdFind_Click()
    Dim Find As Long
    Dim LI As ListItem
    Dim Field As String
    'Set Cnn = New Connection
    Set RsT = New Recordset
    Field = cboField.Text
    lstSearch.ListItems.Clear
    lblCount.Caption = 0
    
    
    If Not cboField.Text = "" Then
        If Not cboCrieria.Text = "" Then
            
            If Not cboCrieria.Text = "Between" And Not cboCrieria.Text = "Like" Then
                
    
                Set rst1 = CConnect.GetRecordSet("Select * from Employee where " & cboField.Text & "" & cboCrieria.Text & "'" & txtFrom.Text & "'")
    
                
                With rst1
                    If .RecordCount > 0 Then
                        lblCount.Caption = .RecordCount
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
                
            ElseIf cboCrieria.Text = "Like" Then
            
                'Set rst1 = cConnect.GetPayData("Select * from Employee order by EmpCode")
                
                Set rst1 = CConnect.GetRecordSet("Select * from Employee order by EmpCode")
               
                With rst1
                    If .RecordCount > 0 Then
                        lblCount.Caption = .RecordCount
                        .MoveFirst
                        Do While Not .EOF
                            .Find "" & cboField.Text & " " & cboCrieria.Text & " '" & txtFrom.Text & "%'", , adSearchForward
                            If Not .EOF Then
                                Set LI = lstSearch.ListItems.add(, , !EmpCode & "")
                                LI.ListSubItems.add , , !SurName & ""
                                LI.ListSubItems.add , , !OtherNames & ""
                                
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
                    Set rst1 = CConnect.GetRecordSet("SELECT * FROM Employee as e LEFT JOIN ECategory as ec ON " & _
                            "e.ECategory = ec.code WHERE e.Term <> 1 AND " & cboField.Text & " " & cboCrieria & " " & txtFrom.Text & " And " & txtTo.Text & "' ORDER BY e.EmpCode")
                    
                Else
    '                Set rst1 = cConnect.GetPayData("select * from Employee where " & cboField.Text & " " & cboCrieria & " '" & txtFrom.Text & "' And '" & txtTo.Text & "'")
                    Set rst1 = CConnect.GetRecordSet("SELECT * FROM Employee as e LEFT JOIN ECategory as ec ON " & _
                            "e.ECategory = ec.code WHERE e.Term <> 1 AND " & cboField.Text & " " & cboCrieria & " '" & txtFrom.Text & "' And '" & txtTo.Text & "' ORDER BY e.EmpCode")
                    Set rst1 = CConnect.GetRecordSet("select * from Employee where " & cboField.Text & " " & cboCrieria & " '" & txtFrom.Text & "' And '" & txtTo.Text & "'")
                    
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
            MsgBox "Select the search criteria.", vbExclamation
        End If
    Else
        MsgBox "Select the search field.", vbExclamation
    End If

End Sub

Private Sub cmdRef_Click()
    Call SRefresh
    
End Sub

Private Sub cmdSelect_Click()

    Sel = ""
    If lstSearch.ListItems.count > 0 Then
        Sel = lstSearch.SelectedItem
        gEmployeeID = lstSearch.SelectedItem.Tag
        If popupText = "ChangeCode" Then
            'frmGenOpt.txtFromE.Text = Sel
            Me.Visible = False
            Exit Sub
        End If
        
        If popupText = "RFrom" Then
            frmRange.txtFromE.Text = Sel
        Else
            frmRange.txtToE.Text = Sel
        End If
        
        If popupText = "RFrom2" Then
            frmRange2.txtFromE.Text = Sel
        Else
            frmRange2.txtToE.Text = Sel
        End If
        
        Me.Visible = False
    Else
        MsgBox "No record selected.", vbExclamation
    End If

End Sub

Private Sub Form_Load()
Dim i As Long
CConnect.CColor Me, MyColor

    With lstSearch
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.add , , "Employee Code", 1500
        .ColumnHeaders.add , , "Surname", 1700
        .ColumnHeaders.add , , "Other Names", 2800
        .ColumnHeaders.add , , "ID No", 1500
        
        .View = lvwReport
      
    End With
    
    With cboField
        .AddItem "EmpCode"
        .AddItem "Surname"
        .AddItem "OtherNames"
        .AddItem "IDNo"
    End With
    
    AllEmployees.GetAccessibleEmployeesByUser currUser.UserID
    'Set rst1 = cConnect.GetPayData("Select * from Employee order by EmpCode")
    Set rst1 = CConnect.GetRecordSet("SELECT * FROM Employee as e LEFT JOIN ECategory as ec ON " & _
            "e.ECategory = ec.code WHERE  Term <> 1 ORDER BY e.EmpCode")
        
    For i = 1 To AllEmployees.count
    
        If Not (AllEmployees.Item(i).IsDisengaged) Then 'IF CLAUSE ADDED BY JOHN TO TRAP OBJECT PRE-LOADS
            Set LI = Me.lstSearch.ListItems.add(, , AllEmployees.Item(i).EmpCode)
            If (lstSearch.ListItems.count > 0) Then
            If lstSearch.ListItems(lstSearch.ListItems.count).Tag = AllEmployees.Item(i).EmployeeID Then
            GoTo nex
            End If
            End If
            LI.Tag = AllEmployees.Item(i).EmployeeID
            LI.SubItems(1) = AllEmployees.Item(i).SurName & ""
            LI.SubItems(2) = AllEmployees.Item(i).OtherNames & ""
            LI.SubItems(3) = AllEmployees.Item(i).IdNo & ""
nex:
        End If
    Next i
        
'    With rst1
'        If .RecordCount > 0 Then
'            .MoveFirst
'            lblCount.Caption = .RecordCount
'            Do While Not .EOF
'                Set LI = lstSearch.ListItems.Add(, , !empcode)
'                LI.SubItems(1) = !SurName & ""
'                LI.SubItems(2) = !OtherNames & ""
'                LI.SubItems(3) = !IdNo & ""
'                .MoveNext
'
'            Loop
'
'        End If
'
'    End With
    
    Set rst1 = Nothing
    cboCrieria.Text = cboCrieria.List(0)
    cboField.Text = cboField.List(0)

    
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
    
End Sub





Public Sub SRefresh()
lstSearch.ListItems.Clear
    
'    Set rst1 = CConnect.GetRecordSet("Select * from Employee order by EmpCode")
    Set rst1 = CConnect.GetRecordSet("SELECT * FROM Employee as e LEFT JOIN ECategory as ec ON " & _
            "e.ECategory = ec.code WHERE  Term <> 1 ORDER BY e.EmpCode")
    
    With rst1
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Set LI = lstSearch.ListItems.add(, , !EmpCode)
                LI.SubItems(1) = !SurName & ""
                LI.SubItems(2) = !OtherNames & ""
                LI.SubItems(3) = !IdNo & ""
                .MoveNext
                
            Loop
            
        End If
 
    End With
    
    Set rst1 = Nothing
    
    
    txtFrom.Text = ""
    txtTo.Text = ""
    cboCrieria.Text = "="
    cboField.Text = "EmpCode"
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
