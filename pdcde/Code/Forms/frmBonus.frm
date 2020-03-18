VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBonus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bonus"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11145
   Icon            =   "frmBonus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      Caption         =   "Bonus Category"
      ForeColor       =   &H80000008&
      Height          =   3030
      Left            =   1980
      TabIndex        =   16
      Top             =   1830
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox txtPerc 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2610
         TabIndex        =   3
         Top             =   885
         Width           =   885
      End
      Begin VB.TextBox txtCeiling 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1380
         TabIndex        =   2
         Top             =   885
         Width           =   1125
      End
      Begin VB.TextBox txtFloor 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   135
         TabIndex        =   1
         Top             =   885
         Width           =   1125
      End
      Begin VB.TextBox txtComments 
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
         ForeColor       =   &H00000000&
         Height          =   825
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1455
         Width           =   5715
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4875
         Picture         =   "frmBonus.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Save Record"
         Top             =   2415
         Width           =   510
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5370
         Picture         =   "frmBonus.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Cancel Process"
         Top             =   2415
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "SALARY RANGE                            % TO RECIEVE"
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
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   3480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Percentage"
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
         Left            =   2625
         TabIndex        =   20
         Top             =   645
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "To"
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
         Left            =   1350
         TabIndex        =   19
         Top             =   645
         Width           =   180
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "From"
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
         Left            =   120
         TabIndex        =   18
         Top             =   645
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Comments"
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
         Left            =   135
         TabIndex        =   17
         Top             =   1230
         Width           =   750
      End
   End
   Begin MSComctlLib.ImageList imgEmpTool 
      Left            =   5430
      Top             =   195
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBonus.frx":0646
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBonus.frx":0758
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBonus.frx":086A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBonus.frx":097C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7860
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11130
      Begin MSComDlg.CommonDialog Cdl 
         Left            =   5610
         Top             =   4560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Move to the First employee"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "CLOSE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6270
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5400
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   8
         ToolTipText     =   "Move to the Previous employee"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   9
         ToolTipText     =   "Move to the Next employee"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdNew 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2775
         Picture         =   "frmBonus.frx":0EBE
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Add New record"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdEdit 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3255
         Picture         =   "frmBonus.frx":0FC0
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Edit Record"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdDelete 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3735
         Picture         =   "frmBonus.frx":10C2
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Delete Record"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   10
         ToolTipText     =   "Move to the Last employee"
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSComctlLib.ListView lvwDetails 
         Height          =   7800
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   11130
         _ExtentX        =   19632
         _ExtentY        =   13758
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgTree"
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
   End
End
Attribute VB_Name = "frmBonus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Sub cmdCancel_Click()
    If PromptSave = True Then
        If MsgBox("Close this window?", vbYesNo + vbQuestion, "Confirm Close") = vbNo Then Exit Sub
    End If
    fraDetails.Visible = False
    With frmMain2
        .cmdNew.Enabled = True
        .cmdEdit.Enabled = True
        .cmdDelete.Enabled = True
        .cmdCancel.Enabled = False
        .cmdSave.Enabled = False
    End With
    Call EnableCmd
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Public Sub cmdDelete_Click()
Dim resp As String


If lvwDetails.ListItems.count > 0 Then
    resp = MsgBox("This will delete  " & lvwDetails.SelectedItem & "-" & lvwDetails.SelectedItem.ListSubItems(1) & " and the corresponding employee bonus from the records." & vbCrLf & "Do you wish to continue?", vbQuestion + vbYesNo)
    If resp = vbNo Then
        Exit Sub
    End If
      
    Action = "DELETED BONUS; SALARY RANGE: " & lvwDetails.SelectedItem & "-" & lvwDetails.SelectedItem.ListSubItems(1) & ";"
    
    CConnect.ExecuteSql ("DELETE FROM Bonus WHERE BonusID = " & IIf(IsNumeric(lvwDetails.SelectedItem.Tag) = True, lvwDetails.SelectedItem.Tag, 0) & "")
        
    rs2.Requery
    
    Call DisplayRecords
        
Else
    MsgBox "You have to select the bonus you would like to delete.", vbInformation
            
End If
    
    
End Sub

Private Sub cmdDone_Click()
    PSave = True
    Call cmdCancel_Click
    PSave = False
End Sub

Public Sub cmdEdit_Click()


If lvwDetails.ListItems.count < 1 Then
    MsgBox "You have to select the bonus you would like to edit.", vbInformation
    PSave = True
    Call cmdCancel_Click
    PSave = False
    Exit Sub
End If


Set rs3 = CConnect.GetRecordSet("SELECT * FROM Bonus WHERE BonusID = " & IIf(IsNumeric(lvwDetails.SelectedItem.Tag) = True, lvwDetails.SelectedItem.Tag, 0) & "")
With rs3
    If .RecordCount > 0 Then
        txtFloor.Text = !floor & ""
        txtCeiling.Text = !ceiling & ""
        txtPerc.Text = !Perc & ""
        txtComments.Text = !Comments & ""

        
        SaveNew = False
    Else
        MsgBox "Record not found.", vbInformation
        Set rs3 = Nothing
        PSave = True
        Call cmdCancel_Click
        PSave = False
        Exit Sub
    End If
End With

Set rs3 = Nothing

Call DisableCmd

fraDetails.Visible = True

cmdSave.Enabled = True
cmdCancel.Enabled = True
SaveNew = False
txtFloor.Locked = False
txtFloor.SetFocus
txtCeiling.Locked = False
txtPerc.Locked = False
End Sub

Private Sub cmdFirst_Click()

With rsGlob
    If .RecordCount > 0 Then
        If .BOF <> True Then
            .MoveFirst
            If .BOF = True Then
                .MoveFirst
                Call DisplayRecords
            Else
                Call DisplayRecords
            End If
            
            Call FirstDisb
            
        End If
    End If
End With


End Sub

Private Sub cmdLast_Click()
With rsGlob
    If .RecordCount > 0 Then
        If .EOF <> True Then
            .MoveLast
            If .EOF = True Then
                .MoveLast
                Call DisplayRecords
            Else
                Call DisplayRecords
            End If
            
            Call LastDisb
            
        End If
    End If
End With

End Sub

Public Sub cmdNew_Click()
Call DisableCmd
txtFloor.Text = ""
txtCeiling.Text = ""
txtComments.Text = ""
fraDetails.Visible = True
cmdCancel.Enabled = True
SaveNew = True
cmdSave.Enabled = True
txtFloor.Locked = False
txtFloor.SetFocus
txtCeiling.Locked = False
End Sub

Private Sub cmdNext_Click()
    
With rsGlob
    If .RecordCount > 0 Then
        If .EOF <> True Then
            .MoveNext
            If .EOF = True Then
                .MoveLast
                Call DisplayRecords
            Else
                Call DisplayRecords
            End If

            Call LastDisb

        End If
    End If
End With


End Sub

Private Sub cmdPrevious_Click()

With rsGlob
    If .RecordCount > 0 Then
        If .BOF <> True Then
            .MovePrevious
            If .BOF = True Then
                .MoveFirst
                Call DisplayRecords
            Else
                Call DisplayRecords
            End If
            
            Call FirstDisb
            
        End If
    End If
End With


End Sub

Public Sub cmdSave_Click()
    On Error GoTo Hell
    
    If IsNumeric(txtFloor.Text) = False Or IsNumeric(txtCeiling.Text) = False Then
        MsgBox "Please provide numeric values.", vbOKOnly + vbInformation, "Numeric values!"
        If IsNumeric(txtFloor.Text) = True Then
            txtFloor.SetFocus
        Else
            txtCeiling.SetFocus
        End If
        Exit Sub
    End If
    
    If IsNumeric(txtPerc.Text) = False Then
        MsgBox "Please enter the percentage increase.", vbOKOnly + vbInformation, "Numeric value"
        txtPerc.SetFocus
        Exit Sub
    End If

    If SaveNew = True Then
        Set rs4 = CConnect.GetRecordSet("SELECT * FROM Bonus WHERE (Floor <= " & txtFloor.Text & " AND ceiling>=" & txtFloor.Text & ") OR (Floor<=" & txtCeiling.Text & " AND Ceiling>=" & txtCeiling.Text & ")")
        With rs4
            If .RecordCount > 0 Then
                MsgBox "Record already exists", vbInformation
                txtFloor.Text = ""
                txtFloor.SetFocus
                Set rs4 = Nothing
                Exit Sub
            End If
        End With
        Set rs4 = Nothing
    End If
       
    If PromptSave = True Then
        If MsgBox("Are you sure you want to save the record.", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
            
    
    Action = "ADDED BONUS; FLOOR: " & txtFloor.Text & "; CEILING: " & txtCeiling.Text & "; PERCENTAGE: " & txtPerc.Text & "; COMMENTS: " & txtComments.Text
    
    CConnect.ExecuteSql ("DELETE FROM Bonus WHERE floor = " & txtFloor.Text & "")
    
    mySQL = "INSERT INTO Bonus (Floor, Ceiling, Perc, Comments)" & _
                        " VALUES(" & txtFloor.Text & "," & txtCeiling.Text & "," & txtPerc.Text & "," & _
                        "'" & txtComments.Text & "')"
'    Debug.Print MySql
    
    CConnect.ExecuteSql (mySQL)

    rs2.Requery
    If SaveNew = False Then
        PSave = True
        PSave = True
        Call cmdCancel_Click
        PSave = False
        PSave = False
    Else
        'rs2.Requery
        'Call DisplayRecords
        txtFloor.SetFocus
    End If
    Call DisplayRecords
Exit Sub
Hell:
End Sub


Private Sub Form_Load()
 
oSmart.FReset Me

If oSmart.hRatio > 1.1 Then
    With frmMain2
        Me.Move .tvwMain.Width + .tvwMain.Width / 32.75, (.Height / 5.52) ' - 155
    End With
Else
     With frmMain2
        Me.Move .tvwMain.Width + .tvwMain.Width / 32.75, .Height / 5.55
    End With
    
End If

CConnect.CColor Me, MyColor

cmdCancel.Enabled = False
cmdSave.Enabled = False

Call InitGrid
'Call 'CConnect.CCon


Set rs2 = CConnect.GetRecordSet("SELECT * FROM Bonus ORDER BY Floor")

Call DisplayRecords

cmdFirst.Enabled = False
cmdPrevious.Enabled = False

End Sub

Private Sub Form_Resize()
    oSmart.FResize Me
    Me.Height = tvwMainheight - 150
    Frame1.Move Frame1.Left, 0, Frame1.Width, tvwMainheight - 150
    lvwDetails.Move lvwDetails.Left, 0, lvwDetails.Width, tvwMainheight - 150
End Sub

Private Sub InitGrid()
    With lvwDetails
        .ColumnHeaders.add , , "From", .Width / 5
        .ColumnHeaders.add , , "To", .Width / 5
        .ColumnHeaders.add , , "Percentage", .Width / 5, vbRightJustify
        .ColumnHeaders.add , , "Comments", 2 * .Width / 5
                   
        .View = lvwReport
    End With
    

End Sub

Public Sub DisplayRecords()
lvwDetails.ListItems.Clear
Call Cleartxt
With rs2
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
            Set li = lvwDetails.ListItems.add(, , !floor & "", , 5)
            li.ListSubItems.add , , !ceiling & ""
            li.ListSubItems.add , , !Perc & ""
            li.ListSubItems.add , , !Comments & ""
            li.Tag = Trim(!BonusID & "")
            .MoveNext
        Loop
    End If
End With
End Sub



Private Sub Form_Unload(Cancel As Integer)

    Set rs2 = Nothing
    Set rs5 = Nothing
'    Set Cnn = Nothing
    

    frmMain2.Caption = "Infiniti HRMIS - PDR [Current User:\" & currUser.FullNames & "]"
    
End Sub

Private Sub fraList_DragDrop(Source As Control, X As Single, y As Single)

End Sub

Private Sub lvwDetails_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwDetails
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvwDetails_DblClick()
    If frmMain2.cmdEdit.Enabled = True Then
        Call frmMain2.cmdEdit_Click
    End If
End Sub




Private Sub LastDisb()
With rsGlob
    If Not .EOF Then
        .MoveNext
        If .EOF Then
            cmdLast.Enabled = False
            cmdNext.Enabled = False
            cmdPrevious.Enabled = True
            cmdFirst.Enabled = True
            cmdPrevious.SetFocus
        End If
        .MovePrevious
    End If
    
    cmdPrevious.Enabled = True
    cmdFirst.Enabled = True
End With
End Sub


Private Sub FirstDisb()
With rsGlob
    If Not .BOF Then
        .MovePrevious
        If .BOF Then
            cmdLast.Enabled = True
            cmdNext.Enabled = True
            cmdPrevious.Enabled = False
            cmdFirst.Enabled = False
            cmdNext.SetFocus
        End If
        .MoveNext
    End If
    
    cmdLast.Enabled = True
    cmdNext.Enabled = True
End With
End Sub


Private Sub Cleartxt()
Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
            i.Text = ""
        End If
    Next i

    lvwDetails.ListItems.Clear
    
End Sub

Public Sub disabletxt()
Dim i As Object
    For Each i In Me
        If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
            i.Locked = True
        End If
    Next i

    
End Sub

Public Sub DisableCmd()
Dim i As Object
    For Each i In Me
        If TypeOf i Is CommandButton Then
            i.Enabled = False
        End If
    Next i
End Sub

Public Sub EnableCmd()
Dim i As Object
    For Each i In Me
        If TypeOf i Is CommandButton Then
            i.Enabled = True
        End If
    Next i
    
End Sub



Public Sub FirstLastDisb()
cmdLast.Enabled = True
cmdNext.Enabled = True
cmdPrevious.Enabled = True
cmdFirst.Enabled = True
cmdNext.SetFocus
            
With rsGlob
    If Not .BOF = True Then
        .MovePrevious
        If .BOF = True Then
            cmdLast.Enabled = True
            cmdNext.Enabled = True
            cmdPrevious.Enabled = False
            cmdFirst.Enabled = False
            cmdNext.SetFocus
        End If
        .MoveNext
    Else
        cmdLast.Enabled = True
        cmdNext.Enabled = True
        cmdPrevious.Enabled = False
        cmdFirst.Enabled = False
        cmdNext.SetFocus
    End If
    
    If Not .EOF = True Then
        .MoveNext
        If .EOF = True Then
            cmdLast.Enabled = False
            cmdNext.Enabled = False
            cmdPrevious.Enabled = True
            cmdFirst.Enabled = True
            cmdPrevious.SetFocus
        End If
        .MovePrevious
    Else
        cmdLast.Enabled = False
        cmdNext.Enabled = False
        cmdPrevious.Enabled = True
        cmdFirst.Enabled = True
        cmdPrevious.SetFocus
    End If
End With
End Sub


