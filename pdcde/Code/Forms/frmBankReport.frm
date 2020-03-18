VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBankReport 
   Caption         =   "Filter"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   4050
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraBanks 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4215
      Begin MSComctlLib.ListView lvwBaks 
         Height          =   6015
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   10610
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "BankCode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Bank Name"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmBankReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Dim itemx As ListItem
    CConnect.CColor Me, MyColor
    
    Set rs = CConnect.GetRecordSet("select * from tblBanks")
    
    If Not rs Is Nothing Then
        If Not (rs.BOF Or rs.EOF) Then
            rs.MoveFirst
            Do Until rs.EOF
                Set itemx = Me.lvwBaks.ListItems.Add(, , rs!Bank_Code)
                itemx.SubItems(1) = rs!Bank_Name
            Loop
         End If
    End If
End Sub
