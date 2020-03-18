VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmEncrypt 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   Caption         =   "Data Importation"
   ClientHeight    =   7005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   1800
      TabIndex        =   26
      Top             =   5640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2760
      TabIndex        =   25
      Top             =   5640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3720
      TabIndex        =   24
      Top             =   5640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   23
      Top             =   5640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   840
      TabIndex        =   22
      Top             =   5640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame fraImportType 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3330
      Left            =   0
      TabIndex        =   6
      Top             =   1710
      Width           =   6825
      Begin VB.OptionButton optEmpHist 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Employment History"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3600
         TabIndex        =   21
         Top             =   2085
         Width           =   2445
      End
      Begin VB.OptionButton optEmpBasic 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Employee Basic/Daily Rate"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2085
         Width           =   2445
      End
      Begin VB.OptionButton optContract 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Contract Details"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3600
         TabIndex        =   18
         Top             =   1800
         Width           =   2445
      End
      Begin VB.OptionButton optProf 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Professional Details"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3600
         TabIndex        =   17
         Top             =   1500
         Width           =   2445
      End
      Begin VB.OptionButton optEdu 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Education History"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3600
         TabIndex        =   16
         Top             =   1200
         Width           =   2445
      End
      Begin VB.OptionButton optReferes 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Referees"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3600
         TabIndex        =   15
         Top             =   900
         Width           =   2445
      End
      Begin VB.OptionButton optNext 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Next of Kin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3600
         TabIndex        =   14
         Top             =   600
         Width           =   2445
      End
      Begin VB.OptionButton optEmpDDetail 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Employee Defined Details"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3600
         TabIndex        =   13
         Top             =   300
         Width           =   2715
      End
      Begin VB.OptionButton optEmpContact 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Employee Contacts"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   2445
      End
      Begin VB.OptionButton optEmpBio 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Employee Bio Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1500
         Width           =   2445
      End
      Begin VB.OptionButton optDDetail 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Defined Details Types"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   2445
      End
      Begin VB.OptionButton optContact 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Contact Types"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   900
         Width           =   2445
      End
      Begin VB.OptionButton optBio 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Bio Data Types"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   2445
      End
      Begin VB.OptionButton optEmp 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Employee Details"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   2445
      End
      Begin VB.Label lblFormat 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   765
         Left            =   135
         TabIndex        =   19
         Top             =   2445
         Width           =   6555
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   1710
      Left            =   15
      TabIndex        =   0
      Top             =   -90
      Width           =   6810
      Begin VB.CommandButton cmdFrom 
         Height          =   345
         Left            =   6285
         Picture         =   "frmEncrypt.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   435
         Width           =   375
      End
      Begin VB.CheckBox chkTab 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Tab Separated"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1620
      End
      Begin VB.TextBox txtFrom 
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
         Height          =   345
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   435
         Width           =   6165
      End
      Begin VB.CommandButton cmdEnc 
         Caption         =   "READ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5160
         TabIndex        =   1
         Top             =   1125
         Width           =   1515
      End
      Begin VB.Label SO 
         BackColor       =   &H00C0E0FF&
         Caption         =   "File to read"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   3
         Top             =   210
         Width           =   2385
      End
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   105
      Top             =   765
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmEncrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function ReadText(inFileName As String) As Variant
On Error Resume Next

    Dim mTextLine As String
    Dim arrResult() As String
    Dim mHandle
    Dim i As Integer
    Dim R&
    Dim empcode As String
    Dim LCode As String
    Dim LBal As Double
    Dim DDue As Date
    
    If MsgBox("This will import leave balances hence updating current records. Do you wish to continue", vbInformation + vbYesNo) = vbNo Then
        Exit Function
    End If

  
    mHandle = FreeFile
    
    Open inFileName For Input As #mHandle
    i = 0
    Do While Not EOF(mHandle)
        Line Input #mHandle, mTextLine
        If Len(Trim(mTextLine)) > 0 Then
              ReDim Preserve arrResult(i)
              i = i + 1
                   ' Place the line read into array
              arrResult(UBound(arrResult)) = mTextLine
                If chkTab.Value = 1 Then
                    R = InStr(mTextLine, vbTab)
                    
                    empcode = left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)
                      
                    R = InStr(mTextLine, vbTab)
                    
                    LCode = left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)
                      
                    LBal = mTextLine
                Else
                    R = InStr(mTextLine, ",")
                    
                    empcode = left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)
                      
                    R = InStr(mTextLine, ",")
                    
                    LCode = left(mTextLine, R - 1)
                    mTextLine = Mid(mTextLine, R + 1)
                      
                    LBal = mTextLine
                      
                End If
                                
                               
                With rs
                    If .RecordCount > 0 Then
                        .MoveFirst
                        .Find "LCode like '" & LCode & "'", , adSearchForward, adBookmarkFirst
                        If .RecordCount > 0 Then
                            CConnect.ExecuteSql ("DELETE FROM EmpLeaves WHERE EmpCode = '" & empcode & "' AND LCode = '" & LCode & "'")
                            
                            With rsGlob2
                                If .RecordCount > 0 Then
                                    .MoveFirst
                                    .Find "EmpCode like '" & empcode & "'", , adSearchForward, adBookmarkFirst
                                    If Not .EOF Then
                                        With rs
                                            If Not IsNull(!IgnProb) Then
                                                If !IgnProb = "No" Then
                                                    If Not IsNull(rsGlob2!DEmployed) And Not IsNull(!ProbPeriod) And Not IsNull(!ProbType) Then
                                                        'LDateDue = cConnect.LDateDue(rsGlob2!DEmployed, !ProbPeriod, !ProbType)
                                                        DDue = CConnect.LDateDue(rsGlob2!DEmployed, !ProbPeriod, !ProbType)
                                                    Else
                                                        If Not IsNull(rsGlob2!DEmployed) Then
                                                            DDue = rsGlob2!DEmployed
                                                        Else
                                                            DDue = Date
                                                        End If
                                                    End If
                                                Else
                                                    If Not IsNull(rsGlob2!DEmployed) Then
                                                        DDue = rsGlob2!DEmployed
                                                    Else
                                                        DDue = Date
                                                    End If
                                                End If
                                            End If
                                                                              
                                            CConnect.ExecuteSql ("INSERT INTO EmpLeaves (EmpCode, LCode, Name," & _
                                                    " DateDue, Days, RDays, Comments) VALUES('" & empcode & "'," & _
                                                    " '" & LCode & "','" & !Name & "" & "','" & DDue & "'," & !MaxDays & "," & LBal & ",'Imported')")
                                        End With
                                    
                                    End If
                                End If
                            End With
                            
                        End If
                    End If
                End With
                
                        
                
        End If
    Loop
    Close #mHandle
    
    MsgBox "Importation completed successfully", vbInformation
    
'    ReadText = arrResult
'    Exit Function
    
'ErrHandler:
''    DBErrMsgProc "Error occurred in GetFileStrArray"
'    MsgBox Err.Description
    
    
End Function


Public Sub cmdCancel_Click()

End Sub

Public Sub cmdDelete_Click()

End Sub

Public Sub cmdEdit_Click()

End Sub

Private Sub cmdFrom_Click()
On Error GoTo errHandler
With cdg
    .CancelError = True
    .DialogTitle = "Select the source database file"
    .Filter = "{*.txt|*.txt"
    .ShowOpen
    
    .Flags = cdlOFNHideReadOnly
    
    txtFrom.Text = .FileName
End With

Exit Sub
errHandler:
    txtFrom.Text = ""

End Sub

Function SaveText(SFileName As String, DFileName As String) As Variant
On Error GoTo errHandler
    Const mconexcludeFieldTypes = "XX8/128/204/205"
    Dim fldCount As Integer
    Dim fldWidth() As Integer
    Dim fldValue As String
    Dim fldType
    Dim mHandle
    Dim tmp
    Dim MyRec
    Dim s As String
    Dim i As Integer, j As Integer
    Dim ExFile As String
    
    
    If IsFileThere(SFileName) = False Then
        MsgBox "Source file does not exist.", vbInformation
        Exit Function
    End If

    If IsFileThere(DFileName) = False Then
        MsgBox "Destination file does not exist.", vbInformation
        Exit Function
    End If
    
    If lineCount(SFileName) < 2 Then
        MsgBox "No records to be encrypted.", vbInformation
        Exit Function
    End If
        
    mHandle = FreeFile
    MyRec = ReadText(SFileName)
    
    Open DFileName For Output As mHandle
    
    For i = 0 To UBound(MyRec)
        Print #mHandle, Crypt(MyRec(i))
        
    Next
    
    Close mHandle
    
    MsgBox "Encryption completed successfully.", vbInformation
   
Exit Function
errHandler:
    MsgBox Err.Description, vbInformation
    

End Function

Function lineCount(ByVal mInFileName As String) As Integer
    Dim mHandle As Integer
    Dim mLinesRead As Integer
    Dim strBuffer As String
    
    mLinesRead = 0
    mHandle = FreeFile

    
    Open mInFileName For Input As mHandle
   
    Do While Not EOF(mHandle)
        Input #mHandle, strBuffer
           'update count
        mLinesRead = mLinesRead + 1
    Loop
    Close mHandle
    lineCount = mLinesRead
End Function
    
Function IsFileThere(inFileSpec As String) As Boolean
    On Error Resume Next
    Dim i
    Dim mFile As String
    mFile = inFileSpec
    i = FreeFile
    Open inFileSpec For Input As i
    If Err Then
        IsFileThere = False
    Else
        Close i
        IsFileThere = True
    End If
    
End Function


Public Function Crypt(Pwd As Variant) As String
Dim Temp As String, PwdChr As Long
Dim EncryptKey As Long

EncryptKey = 20  'Int(Sqr(Len(Pwd) * 81)) + 23

For PwdChr = 1 To Len(Pwd)
    Temp = Temp + Chr(Asc(Mid(Pwd, PwdChr, 1)) + EncryptKey)
Next PwdChr

Crypt = Temp

End Function

Private Sub cmdEnc_Click()
Dim TabDil As Boolean
If chkTab.Value = 1 Then
    TabDil = True
Else
    TabDil = False
End If
If optEmp.Value = True Then
    Call CConnect.ImportEmp(txtFrom.Text, TabDil)
    rsGlob.Requery
    rsGlob2.Requery
    Call frmMain2.LoadMyList
ElseIf optBio.Value = True Then
    Call CConnect.ImportBioTypes(txtFrom.Text, TabDil)
ElseIf optContact.Value = True Then
    Call CConnect.ImportCTypes(txtFrom.Text, TabDil)
ElseIf optDDetail.Value = True Then
    Call CConnect.ImportDTypes(txtFrom.Text, TabDil)
ElseIf optEmpBio.Value = True Then
    Call CConnect.ImportEmpBTypes(txtFrom.Text, TabDil)
ElseIf optEmpBasic.Value = True Then
    Call CConnect.ImportEmpBasic(txtFrom.Text, TabDil)
ElseIf optEmpContact.Value = True Then
    Call CConnect.ImportEmpCTypes(txtFrom.Text, TabDil)
ElseIf optEmpDDetail.Value = True Then
    Call CConnect.ImportEmpDTypes(txtFrom.Text, TabDil)
ElseIf optEmpHist.Value = True Then
    Call CConnect.ImportEmploy(txtFrom.Text, TabDil)
ElseIf optNext.Value = True Then
    Call CConnect.ImportKins(txtFrom.Text, TabDil)
ElseIf optReferes.Value = True Then
    Call CConnect.ImportRefeeres(txtFrom.Text, TabDil)
ElseIf optEdu.Value = True Then
    Call CConnect.ImportEDetails(txtFrom.Text, TabDil)
ElseIf optProf.Value = True Then
    Call CConnect.ImportPDetails(txtFrom.Text, TabDil)
ElseIf optContract.Value = True Then
    Call CConnect.ImportContracts(txtFrom.Text, TabDil)
Else
    MsgBox "No Details selected", vbInformation
End If
End Sub

Private Sub cmdTo_Click()
On Error GoTo errHandler
With cdg
    .CancelError = True
    .DialogTitle = "Select the source database file"
    .Filter = "{*.txt|*.txt"
    .ShowOpen
    .Flags = cdlOFNHideReadOnly
    
    txtTo.Text = .FileName
End With

Exit Sub
errHandler:
    txtTo.Text = ""
    

End Sub

Public Sub cmdNew_Click()

End Sub

Public Sub cmdSave_Click()

End Sub

Private Sub Form_Load()
Decla.Security Me
oSmart.FReset Me

If oSmart.hRatio > 1.1 Then
    With frmMain2
        Me.Move .tvwMain.Width + .tvwMain.Width / 32.75, (.Height / 5.52) ''- 155
    End With
Else
     With frmMain2
        Me.Move .tvwMain.Width + .tvwMain.Width / 32.75, .Height / 5.55
    End With
    
End If

CConnect.CColor Me, MyColor

    
End Sub

Private Sub Form_Resize()
    oSmart.FResize Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain2.Caption = "Personnel Director " & App.FileDescription
End Sub

Private Sub optBio_Click()
    lblFormat.Caption = "Format: BioCode,Description,Comments"
End Sub

Private Sub optContact_Click()
    lblFormat.Caption = "Format: ContactCode,Description,Comments"
End Sub

Private Sub optContract_Click()
    lblFormat.Caption = "Format:EmpCode,ContractCode,Description,Start Date,End Date,Ref,Comments"
End Sub

Private Sub optDDetail_Click()
    lblFormat.Caption = "Format: DDetailCode,Description,Comments"
End Sub

Private Sub optEdu_Click()
    lblFormat.Caption = "Format:Employee Code,EduCode,Course,Start Date,End Date,Level Reached,Award,Comments,Institution"
End Sub

Private Sub optEmp_Click()
    lblFormat.Caption = "Format:Employee Code,SurName,OtherNames,IDNo,Date Of Birth,Gender,Date Employed,NhifNo,NssfNo,PinNo,Employee Type, Terms of employment, BasicPay, Employee Category"
End Sub

Private Sub optEmpBasic_Click()
    lblFormat.Caption = "Format:Employee Code,Basic/Daily Pay"
End Sub

Private Sub optEmpBio_Click()
    lblFormat.Caption = "Format: Employee Code,BioCode,BioData,Comments"
End Sub

Private Sub optEmpContact_Click()
    lblFormat.Caption = "Format: Employee Code,Contact Code,Contact"
End Sub

Private Sub optEmpDDetail_Click()
    lblFormat.Caption = "Format: Employee Code,Defined Detail Code,Detail,Comments"
End Sub


Private Sub optEmpHist_Click()
    lblFormat.Caption = "Format:Employee Code,Employment Code,Employer,Start Date,End Date,Reasons for leaving,Designation, Supervisor, Salary,Comments"
End Sub

Private Sub optNext_Click()
    lblFormat.Caption = "Format:Employee Code,KinCode,SurName,OtherNames,IDNo,Relation,Date of Birth,Occupation,Home Tel,Office Tel,Cell No,Email,Address,Signed,Sign Date,Comments"
End Sub

Private Sub optProf_Click()
    lblFormat.Caption = "Format:Employee Code,Profession Code,Course,Start Date,End Date,Level Reached,Award,Comments"
End Sub

Private Sub optReferes_Click()
    lblFormat.Caption = "Format:Employee Code,Referee Code,Names,IDNo,CellNo,Email,Address," & vbLf & _
    "Comments"
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

