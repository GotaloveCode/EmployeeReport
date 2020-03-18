VERSION 5.00
Begin VB.Form frmLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login to Personnel Director"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4860
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraChange 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1995
      Left            =   45
      TabIndex        =   12
      Top             =   -105
      Visible         =   0   'False
      Width           =   4545
      Begin VB.CommandButton cmdPChange 
         Caption         =   "CHANGE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1455
         TabIndex        =   10
         Top             =   1500
         Width           =   1380
      End
      Begin VB.CommandButton cmdPCancel 
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3030
         TabIndex        =   8
         Top             =   1515
         Width           =   1380
      End
      Begin VB.TextBox txtCPass 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1950
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   825
         Width           =   2460
      End
      Begin VB.TextBox txtNPass 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1950
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   420
         Width           =   2460
      End
      Begin VB.Label Label3 
         BackColor       =   &H00F2FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   14
         Top             =   870
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00F2FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "New Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   225
         TabIndex        =   13
         Top             =   450
         Width           =   1575
      End
   End
   Begin VB.Frame fraLog 
      BorderStyle     =   0  'None
      Height          =   1995
      Left            =   30
      TabIndex        =   7
      Top             =   0
      Width           =   4545
      Begin VB.CommandButton cmdChange 
         Caption         =   "Change Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   240
         TabIndex        =   4
         Top             =   1485
         Width           =   1455
      End
      Begin VB.TextBox txtUserName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1815
         TabIndex        =   0
         Top             =   315
         Width           =   2565
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Ok"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1830
         TabIndex        =   2
         Top             =   1485
         Width           =   1260
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3120
         TabIndex        =   3
         Top             =   1485
         Width           =   1260
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1815
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   840
         Width           =   2565
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "&User Name:"
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
         Height          =   270
         Index           =   0
         Left            =   480
         TabIndex        =   11
         Top             =   330
         Width           =   1200
      End
      Begin VB.Label lblLabels 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "&Password:"
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
         Height          =   270
         Index           =   1
         Left            =   495
         TabIndex        =   9
         Top             =   855
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As Recordset
Dim rstCUser As Recordset
Dim EncryptPass As String
Dim Excc As String
Dim Trials As Long
Dim UnfreezeIt As String
Private Const EWX_SHUTDOWN As Long = 1
Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long


Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
  
    LoginSucceeded = False
    End
 
End Sub

Private Sub cmdChange_Click()
txtPassword.Text = EncryptPassword(txtPassword.Text)
With rs
    If .RecordCount > 0 Then
        If txtUserName.Text = "" Or txtPassword.Text = "" Then
            MsgBox "Invalid Password! Check that the Caps lock is not on by mistake then try again.", vbExclamation
            Exit Sub
        End If
            
        .MoveFirst
        .Find "UID like '" & txtUserName.Text & "" & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            If IsNull(!Pass) Then
                !Pass = ""
            End If
            
            If !Pass = txtPassword.Text And !UID = txtUserName.Text Then
                txtNPass.Text = ""
                txtCPass.Text = ""
                fraChange.Visible = True
                txtNPass.SetFocus
                Exit Sub
   
                
            End If
        Else
            With rsGlob
                If .RecordCount > 0 Then
                    .MoveFirst
                    .Filter = "EmpCode like '" & txtUserName.Text & "'"
                    If .RecordCount > 0 Then
                        If txtPassword.Text = CConnect.Crypt(DPass) Then
                            txtNPass.Text = ""
                            txtCPass.Text = ""
                            fraChange.Visible = True
                            txtNPass.SetFocus
                            .Filter = adFilterNone
                            Exit Sub
                    
                        End If
                    End If
                    .Filter = adFilterNone
             
                End If
            End With
          
        End If
    Else
        With rsGlob
            If .RecordCount > 0 Then
                .MoveFirst
                .Filter = "EmpCode like '" & txtUserName.Text & "'"
                If .RecordCount > 0 Then
                    If txtPassword.Text = CConnect.Crypt(DPass) Then
                        txtNPass.Text = ""
                        txtCPass.Text = ""
                        fraChange.Visible = True
                        txtNPass.SetFocus
                        .Filter = adFilterNone
                        Exit Sub
                    End If
                End If
                .Filter = adFilterNone
       
            End If
        End With
        
    End If
End With

Beep
MsgBox "Invalid Password! Check that the Caps lock is not on by mistake then try again.", vbExclamation
         
txtUserName.SetFocus
txtPassword.Text = ""
                
End Sub

Private Sub CmdOk_Click()

Dim MyChangePass As String

On Error GoTo errHandler
If fraChange.Visible = True Then
    Call cmdPChange_Click
    Exit Sub
End If

Dim lngresult As Variant
Dim fm As Form
Dim MyPass As String

If ((Trials > 3) And (txtUserName.Text <> "ADMIN")) Then ' Get value from tblpasswordrule
    MsgBox "The System will Terminate due to wrong Credentials." & vbCrLf & "For security reasons, windows will log off the its current user." & vbCrLf & "Contact Your System Administrator", vbCritical
    rs!frozen = True
    rs!Reason_Frozen = "Too many logon Attempts"
    rs.Update
    'ExitWindowsEx 0, 0
    End
    Exit Sub
End If
txtPassword.Text = EncryptPassword(txtPassword.Text)
    
    With rs
        If .RecordCount > 0 Then
            If txtUserName.Text = "" Or txtPassword.Text = "" Then
                MsgBox "Please Enter Login Credentials.", vbExclamation
                Exit Sub
            End If
                
            .MoveFirst
            .Find "UID like '" & txtUserName.Text & "" & "'", , adSearchForward, adBookmarkFirst
            
            '********************************************************'
            'Estabish maximum access employee category'
            '********************************************************'
            If Trim(!categoryAccess) = "0" Then
                Set rs5 = CConnect.GetRecordSet("SELECT TOP 1 * FROM ECategory ORDER BY seq ASC")
                While rs5.EOF = False
                    rs5.MoveFirst
                    maxCatAccess = rs5!seq & ""
                    rs5.MoveNext
                Wend
            Else
                Set rs5 = CConnect.GetRecordSet("SELECT TOP 1 * FROM ECategory WHERE code = '" & Trim(!categoryAccess) & "' ORDER BY seq ASC")
                While rs5.EOF = False
                    rs5.MoveFirst
                    maxCatAccess = rs5!seq & ""
                    rs5.MoveNext
                Wend
            End If
            
            '********************************************************'
            'Estabish Departmental Constraints'
            '********************************************************'
            
            deptConstraint = Trim(!DeptCode & "")
            
            '********************************************************'
            'Checking for frozen accounts if found then exit procedure'
            '********************************************************'
            If !frozen = True Then
                MsgBox "This account Has been frozen" & vbCrLf & _
                       "Reason : " & !Reason_Frozen & vbCrLf & _
                       "Please Contact the Administrator for " & _
                       "Account Activation" _
                       , vbInformation, _
                       rs!UID & " Account Frozen"
                       
                txtUserName.SetFocus
                txtUserName.SelStart = 0
                txtUserName.SelLength = Len(txtUserName)
                txtPassword = ""
                UnfreezeIt = "Yes"
                Trials = 0
                Exit Sub
                       
            End If
                
            If Not .EOF Then
                If IsNull(!Pass) Then
                    !Pass = ""
                End If
                
                
                If !Pass = txtPassword.Text And !UID = txtUserName.Text Then
                    
                    If IsNull(rs!Excc) Then
                        rs!Excc = 0
                    End If

                    Excc = rs!Excc
                   ' Call EEncryptPassword
'                    Excc = EncryptPass
                   
                   '// Promt the user to change the password as the days approach the expiry period

                        If DateDiff("d", Date, !EDate) < 8 And DateDiff("d", Date, !EDate) > 0 Then
                            MsgBox "Your password will expire in " & DateDiff("d", Date, !EDate) & " Days. Please change your password.", vbInformation

                        ElseIf DateDiff("d", Date, !EDate) < 1 Or IsNull(!EDate) Then
                            MsgBox "Your password has expired. Please change.", vbInformation
                            fraChange.Visible = True
                            txtNPass.SetFocus
                            'Call cmdChange_Click
                            'txtPassword.Text = MyChangePass
                            'Call cmdChange_Click
                            Exit Sub

                        End If
                    
                    '++ Save Login credentials to the Registry ++
                        SaveSetting "Infiniti PD", "Startup", "Left", txtUserName
                    '++ Save Login credentials to the Registry ++

                    Unload Me
                    frmMain2.fraLog.Visible = False
                    
                    CurrentUser = !UID & ""
                    CGroup = !GNo & ""
                    CConnect.ExecuteSql "UPDATE SECURITY SET MACHINE='" & Environ("COMPUTERNAME") & "' , STATUS='LOGGED IN' , LOGINTIME= getdate() WHERE UID='" & CurrentUser & "' and subsystem = '" & SubSystem & "' AND GNO='" & CGroup & "'"
                                                    
                    If Not IsNull(!empcode) Then
                        If !empcode <> "" Then
                            Set rs4 = CConnect.GetRecordSet("SELECT ECategory FROM Employee WHERE empcode = '" & !empcode & "'")
                            EmpCat = rs4!ECategory & ""
                        End If
                    End If
                Else
                
                                     
                    Beep
                    
                    Trials = Trials + 1
                    
                    If Trials >= 2 Then
                        MsgBox "WARNING: You have * " & CInt(3 - Trials) & " * Attemps Only." & vbCrLf & "The 3rd Login Failure Will Terminate The System.", vbExclamation, "Checking Login Credentials"
                        Trials = Trials + 1
                    Else
                        MsgBox "Invalid Credentials! Check Caps locks then try again." & vbCrLf & "You have * " & CInt(3 - Trials) & " * Attemps Only", vbExclamation, "Checking Login Credentials"
                    End If
                    Debug.Print EncryptPassword(1958)
                    txtPassword.Text = ""
                    Exit Sub
                    
                End If
                
            Else
            
            '++The system shold go the users table and search the user name instead of using the EmpCode. Hii ni blunder man.++
            '++Monte 15.07.05++
                With rsGlob
                    If .RecordCount > 0 Then
                        .MoveFirst
                        .Filter = "EmpCode like '" & txtUserName.Text & "'"
                        If .RecordCount > 0 Then
                            If txtPassword.Text = CConnect.Crypt(DPass) Then
                                MsgBox "You have to change your password.", vbInformation
                                txtPassword.Text = CConnect.Crypt(txtPassword.Text)
                                cmdChange_Click
                                Exit Sub
                                
                                Unload Me
                                frmMain2.fraLog.Visible = False
                      
                                
                                Exit Sub
                            End If
                        End If
                        .Filter = adFilterNone
                    End If
                End With
            '++The system should go the users table and search the user name instead of using the EmpCode. Hii ni blunder man.++
            '++Monte 15.07.05++
               
                Trials = Trials + 1
                    
                Beep
                
                If Trials >= 2 Then
                    MsgBox "WARNING: You have * " & CInt(3 - Trials) & " * Attemps Only." & vbCrLf & "The 3rd Login Failure Will Terminate The System.", vbExclamation, "Checking Login Credentials"
                    Trials = Trials + 1
                Else
                    MsgBox "Invalid Credentials! Check Caps locks then try again." & vbCrLf & "You have * " & CInt(3 - Trials) & " * Attemps Only", vbExclamation, "Checking Login Credentials"
                End If
                    txtUserName.SetFocus
                    txtPassword.Text = ""
                Exit Sub
            End If
        Else

            Unload Me

            frmMain2.fraLog.Visible = False
     

        End If
    End With
    'frmMain2.mnuReports.Visible = True
Dim myDate As Date
    myDate = Format(Date, "dd/mm/yyyy")
Dim rsPath As New ADODB.Recordset
Set rsPath = CConnect.GetRecordSet("select * from companyLogo")
If rsPath.RecordCount > 0 Then
    PicPath = Trim(rsPath!Path & "")
End If
Exit Sub
errHandler:
MsgBox "Invalid Username! The system will now terminate.", vbCritical
End
End Sub

Private Sub cmdPCancel_Click()
    txtPassword.Text = CConnect.Crypt(txtPassword.Text)
    fraChange.Visible = False
    txtPassword.SetFocus
End Sub

Private Sub cmdPChange_Click()
Dim lngresult As Variant
Dim fm As Form
Dim MyPass As String
Dim ExpiryDate As Date, ExpiryPeriod As Integer
Dim rsExpiry As New Recordset

Trials = 0
If txtNPass.Text = "" Then
    MsgBox "You must enter the new password.", vbInformation
    txtNPass.SetFocus
    Exit Sub
End If

If txtCPass.Text = "" Then
    MsgBox "You must enter password confirmation.", vbInformation
    txtCPass.SetFocus
    Exit Sub
End If

If Not txtNPass.Text = txtCPass.Text Then
    MsgBox "Passwords do not match.", vbInformation
    txtNPass.SetFocus
    Exit Sub
End If

txtNPass.Text = EncryptPassword(txtNPass.Text)

    With rs
        If .RecordCount > 0 Then
        
            If txtUserName.Text = "" Or txtPassword.Text = "" Then
                MsgBox "Invalid Password! Check that the Caps lock is not on by mistake then try again.", vbExclamation
                Exit Sub
            End If
                
            .MoveFirst
            .Find "UID like '" & txtUserName.Text & "" & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                If IsNull(!Pass) Then
                    !Pass = ""
                End If
                
                If !Pass = txtPassword.Text And !UID = txtUserName.Text Then
                    '// Check on the expiry dates ++ Cornelius on 17-08-2005
                
                    Set rsExpiry = CConnect.GetRecordSet("Select *  From tblPasswordRule")
                    If Not rsExpiry.EOF Then
                        ExpiryPeriod = rsExpiry!Change_After & ""
                    Else
                        ExpiryPeriod = 30 '// A default period incase there is non set
                    End If
                    
                    ExpiryDate = DateAdd("d", ExpiryPeriod, Date)
                    
                    CConnect.ExecuteSql ("UPDATE Security SET Pass = '" & txtNPass.Text & "', Edate = '" & ExpiryDate & "' WHERE UID = '" & txtUserName.Text & "' AND subsystem = '" & SubSystem & "'")
                    If CPass = "Yes" Then
                        MsgBox "Your password has been changed.", vbInformation
                    Else
                        MsgBox "Your password has been changed.", vbInformation
                    End If
                    
                    fraChange.Visible = False
                    txtPassword.Text = ""
                    txtPassword.SetFocus
                    fraChange.Visible = False
                    rs.Requery
                    Exit Sub
                End If
            Else
                With rsGlob
                    If .RecordCount > 0 Then
                        .MoveFirst
                        .Filter = "EmpCode like '" & txtUserName.Text & "'"
                        If .RecordCount > 0 Then
                            If txtPassword.Text = CConnect.Crypt(DPass) Then
                                CConnect.ExecuteSql ("INSERT INTO Security (UID, GNo, Pass, Excc) VALUES('" & txtUserName.Text & "','" & EmpGroup & "','" & txtNPass.Text & "','" & CConnect.Crypt(0) & "')")
                                MsgBox "Your password has been changed.", vbInformation
                                fraChange.Visible = False
                                txtPassword.Text = ""
                                txtPassword.SetFocus
                                fraChange.Visible = False
                                .Filter = adFilterNone
                                rs.Requery
                                Exit Sub
                            End If
                        End If
                        .Filter = adFilterNone
                    End If
                End With
              
            End If
        Else
            With rsGlob
                If .RecordCount > 0 Then
                    .MoveFirst
                    .Filter = "EmpCode like '" & txtUserName.Text & "'"
                    If .RecordCount > 0 Then
                        If txtPassword.Text = CConnect.Crypt(DPass) Then
                            CConnect.ExecuteSql ("DELETE FROM Security WHERE UID = '" & txtUserName.Text & "'")
                            CConnect.ExecuteSql ("INSERT INTO Security (UID, GNo, Pass, Excc) VALUES('" & txtUserName.Text & "','" & EmpGroup & "','" & txtNPass.Text & "','" & CConnect.Crypt(0) & "')")
                            MsgBox "Your password has been changed.", vbInformation
                            fraChange.Visible = False
                            txtPassword.Text = ""
                            txtPassword.SetFocus
                            fraChange.Visible = False
                            .Filter = adFilterNone
                            rs.Requery
                            Exit Sub
                        End If
                    End If
                    .Filter = adFilterNone
                End If
            End With
            
        End If
    End With
    
    Beep
    MsgBox "Invalid Password! Check that the Caps lock is not on by mistake then try again.", vbExclamation
    fraChange.Visible = False
    txtUserName.SetFocus
    txtPassword.Text = ""
                    
End Sub


Private Sub Form_Load()

    CConnect.CColor Me, MyColor
    
    Machine_User_Settings = GetSetting(Appname:="Infiniti PD", Section:="Startup", _
                           Key:="Left", Default:="")
                           
    txtUserName = Trim$(Machine_User_Settings)
    
    Set rs = CConnect.GetRecordSet("SELECT * FROM Security WHERE subsystem = '" & SubSystem & "'")
    
    Trials = 0

End Sub


Public Sub shut()
Dim lngresult As String
    'shut down the computer
    lngresult = ExitWindowsEx(EWX_SHUTDOWN, 0&)
End Sub

Function EncryptPassword(Pwd As String) As String
'Dim Pwd As Variant
Dim Temp As String, PwdChr As Long 'Integer
Dim EncryptKey As Long 'Integer
'Pwd = txtPassword.Text
EncryptKey = Int(Sqr(Len(Pwd) * 81)) + 23

For PwdChr = 1 To Len(Pwd)
    Temp = Temp + Chr(Asc(Mid(Pwd, PwdChr, 1)) Xor EncryptKey)
Next PwdChr
''
''EncryptPass = Temp
''txtPassword.Text = EncryptPass
EncryptPassword = Temp
End Function

Function EEncryptPassword(Pwd As String) As String
'Dim Pwd As Variant
Dim Temp As String, PwdChr As Long
Dim EncryptKey As Long
'Pwd = Excc
EncryptKey = Int(Sqr(Len(Pwd) * 81)) + 23

For PwdChr = 1 To Len(Pwd)
    Temp = Temp + Chr(Asc(Mid(Pwd, PwdChr, 1)) Xor EncryptKey)
Next PwdChr
''
''EncryptPass = Temp
EEncryptPassword = Temp
End Function


Private Sub txtUsername_Change()
    txtUserName.Text = UCase(txtUserName.Text)
    txtUserName.SelStart = Len(txtUserName.Text)
End Sub

