VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmCurrencies 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Currencies"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   Icon            =   "frmCurrencies.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab sstCurrencies 
      Height          =   6420
      Left            =   30
      TabIndex        =   20
      Top             =   435
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   11324
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Currencies"
      TabPicture(0)   =   "frmCurrencies.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraCurrency"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdNewC"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdEditC"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdDeleteC"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdDenominations"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdClose"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Denominations"
      TabPicture(1)   =   "frmCurrencies.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdBack"
      Tab(1).Control(1)=   "cmdDelete"
      Tab(1).Control(2)=   "cmdEdit"
      Tab(1).Control(3)=   "cmdNew"
      Tab(1).Control(4)=   "txtParentCurrency"
      Tab(1).Control(5)=   "fraDenomination"
      Tab(1).Control(6)=   "fraDenominations"
      Tab(1).Control(7)=   "Label5"
      Tab(1).ControlCount=   8
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   465
         Left            =   5700
         TabIndex        =   11
         Top             =   5700
         Width           =   1215
      End
      Begin VB.CommandButton cmdDenominations 
         Caption         =   "Denominations..."
         Height          =   465
         Left            =   3825
         TabIndex        =   10
         Top             =   5700
         Width           =   1440
      End
      Begin VB.CommandButton cmdDeleteC 
         Caption         =   "Delete"
         Height          =   465
         Left            =   2475
         TabIndex        =   9
         Top             =   5700
         Width           =   1140
      End
      Begin VB.CommandButton cmdEditC 
         Caption         =   "Edit"
         Height          =   465
         Left            =   1425
         TabIndex        =   8
         Top             =   5700
         Width           =   990
      End
      Begin VB.CommandButton cmdNewC 
         Caption         =   "New"
         Height          =   465
         Left            =   300
         TabIndex        =   7
         Top             =   5700
         Width           =   990
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back To Currencies"
         Height          =   465
         Left            =   -69750
         TabIndex        =   19
         Top             =   5820
         Width           =   1590
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   465
         Left            =   -72150
         TabIndex        =   18
         Top             =   5820
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   465
         Left            =   -73575
         TabIndex        =   17
         Top             =   5820
         Width           =   1215
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   465
         Left            =   -74850
         TabIndex        =   16
         Top             =   5820
         Width           =   1140
      End
      Begin VB.TextBox txtParentCurrency 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73950
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   525
         Width           =   5490
      End
      Begin VB.Frame fraDenomination 
         Height          =   1140
         Left            =   -74850
         TabIndex        =   27
         Top             =   1050
         Width           =   6765
         Begin VB.TextBox txtValue 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1275
            TabIndex        =   14
            Top             =   675
            Width           =   2490
         End
         Begin VB.TextBox txtDenomination 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1275
            TabIndex        =   13
            Top             =   300
            Width           =   5040
         End
         Begin VB.Label Label7 
            Caption         =   "Value:"
            Height          =   165
            Left            =   150
            TabIndex        =   29
            Top             =   750
            Width           =   915
         End
         Begin VB.Label Label6 
            Caption         =   "Denomination:"
            Height          =   165
            Left            =   75
            TabIndex        =   28
            Top             =   300
            Width           =   1065
         End
      End
      Begin VB.Frame fraDenominations 
         Caption         =   "Denominations"
         Height          =   3360
         Left            =   -74850
         TabIndex        =   26
         Top             =   2325
         Width           =   6765
         Begin MSComctlLib.ListView lvwDenominations 
            Height          =   2955
            Left            =   150
            TabIndex        =   15
            Top             =   300
            Width           =   6465
            _ExtentX        =   11404
            _ExtentY        =   5212
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Denomination"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Value"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Currency"
               Object.Width           =   5292
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Existing Currencies:"
         Height          =   3810
         Left            =   150
         TabIndex        =   22
         Top             =   1800
         Width           =   6765
         Begin MSComctlLib.ListView lvwCurrencies 
            Height          =   3285
            Left            =   150
            TabIndex        =   6
            Top             =   375
            Width           =   6465
            _ExtentX        =   11404
            _ExtentY        =   5794
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Currency Name"
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Currency Code"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Symbol"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Conversion Rate"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Is Base Currency"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame fraCurrency 
         Height          =   1215
         Left            =   150
         TabIndex        =   21
         Top             =   450
         Width           =   6765
         Begin VB.TextBox txtCurrencyCode 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   3
            Top             =   690
            Width           =   840
         End
         Begin VB.TextBox txtConversionRate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3945
            TabIndex        =   4
            Top             =   675
            Width           =   915
         End
         Begin VB.TextBox txtSymbol 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5775
            TabIndex        =   2
            Top             =   225
            Width           =   840
         End
         Begin VB.TextBox txtCurrencyName 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1425
            TabIndex        =   1
            Top             =   225
            Width           =   3465
         End
         Begin VB.CheckBox chkIsBaseCurrency 
            Caption         =   "Is Base Currency"
            Height          =   195
            Left            =   5055
            TabIndex        =   5
            Top             =   735
            Width           =   1590
         End
         Begin VB.Label Label8 
            Caption         =   "Currency Code:"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   705
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Conversion Rate:"
            Height          =   240
            Left            =   2595
            TabIndex        =   25
            Top             =   705
            Width           =   1290
         End
         Begin VB.Label Label3 
            Caption         =   "Symbol:"
            Height          =   240
            Left            =   5100
            TabIndex        =   24
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Currency Name:"
            Height          =   240
            Left            =   120
            TabIndex        =   23
            Top             =   247
            Width           =   1215
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Currency:"
         Height          =   240
         Left            =   -74775
         TabIndex        =   30
         Top             =   600
         Width           =   840
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Currencies"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   75
      TabIndex        =   0
      Top             =   -45
      Width           =   1275
   End
End
Attribute VB_Name = "frmCurrencies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pCurrencies As HRCORE.Currencies
Private selCurr As HRCORE.CCurrency
Private pDenominations As HRCORE.Denominations
Private selDenom As HRCORE.Denomination

Private Sub chkIsBaseCurrency_Click()
    If chkIsBaseCurrency.value = vbChecked Then
        Me.txtConversionRate.Text = 1#
        Me.txtConversionRate.Locked = True
    Else
        Me.txtConversionRate.Locked = False
    End If
End Sub

Private Sub cmdBack_Click()
    Me.sstCurrencies.TabEnabled(0) = True
    sstCurrencies.TabVisible(1) = False
    sstCurrencies.Tab = 0
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdDeleteC_Click()
    Dim retVal As Long
    Dim resp As Long
    
    Select Case LCase(cmdDeleteC.Caption)
        Case "delete"
            If selCurr Is Nothing Then
                MsgBox "Select the Currency you want to delete", vbExclamation, TITLES
            Else
                resp = MsgBox("Are yo sure you want to delete the selected currency?", vbYesNo + vbQuestion, TITLES)
                If resp = vbYes Then
                    retVal = selCurr.Delete()
                    LoadCurrencies
                End If
            End If
        Case "cancel"
            cmdNewC.Enabled = True
            cmdEditC.Caption = "Edit"
            cmdDeleteC.Caption = "Delete"
            fraCurrency.Enabled = False
            cmdDenominations.Enabled = True
            LoadCurrencies
    End Select
End Sub

Private Sub cmdDelete_Click()
    Dim retVal As Long
    Dim resp As Long
    
    Select Case LCase(cmdDelete.Caption)
        Case "delete"
            If selDenom Is Nothing Then
                MsgBox "Select the Denomination you want to delete", vbExclamation, TITLES
            Else
                resp = MsgBox("Are yo sure you want to delete the selected Denomination?", vbYesNo + vbQuestion, TITLES)
                If resp = vbYes Then
                    retVal = selDenom.Delete()
                    LoadDenominationsOfCurrency selCurr, True
                End If
            End If
            
        Case "cancel"
            cmdNew.Enabled = True
            cmdEdit.Caption = "Edit"
            cmdDelete.Caption = "Delete"
            cmdBack.Enabled = True
            fraDenomination.Enabled = False
            LoadDenominationsOfCurrency selCurr, False
    End Select
End Sub

Private Sub cmdDenominations_Click()
    If selCurr Is Nothing Then
        MsgBox "Select the Currency to set Denominations for", vbInformation, TITLES
    Else
        Me.txtParentCurrency.Text = selCurr.CurrencyName & " (" & selCurr.CurrencySymbol & ")"
        LoadDenominationsOfCurrency selCurr, False
        sstCurrencies.TabVisible(1) = True
        sstCurrencies.TabEnabled(0) = False
        sstCurrencies.Tab = 1
    End If
End Sub

Private Sub cmdEditC_Click()
    Select Case LCase(cmdEditC.Caption)
        Case "edit"
            cmdNewC.Enabled = False
            cmdEditC.Caption = "Update"
            cmdDeleteC.Caption = "Cancel"
            cmdDenominations.Enabled = False
            fraCurrency.Enabled = True
            
        Case "update"
            If Update() = False Then Exit Sub
            cmdNewC.Enabled = True
            cmdEditC.Caption = "Edit"
            cmdDeleteC.Caption = "Delete"
            cmdDenominations.Enabled = True
            fraCurrency.Enabled = False
            LoadCurrencies
            
        Case "cancel"   'cancels a new operation
            cmdNewC.Caption = "New"
            cmdEditC.Caption = "Edit"
            cmdDeleteC.Enabled = True
            fraCurrency.Enabled = False
            cmdDenominations.Enabled = True
            
    End Select
End Sub

Private Sub cmdEdit_Click()
    Select Case LCase(cmdEdit.Caption)
        Case "edit"
            cmdNew.Enabled = False
            cmdEdit.Caption = "Update"
            cmdDelete.Caption = "Cancel"
            cmdBack.Enabled = False
            fraDenomination.Enabled = True
            
        Case "update"
            If UpdateD() = False Then Exit Sub
            cmdNew.Enabled = True
            cmdEdit.Caption = "Edit"
            cmdDelete.Caption = "Delete"
            fraDenomination.Enabled = False
            cmdBack.Enabled = True
            LoadDenominationsOfCurrency selCurr, True
            
        Case "cancel"   'cancels a new operation
            cmdNew.Caption = "New"
            cmdEdit.Caption = "Edit"
            cmdDelete.Enabled = True
            cmdBack.Enabled = True
            fraDenomination.Enabled = False
            
    End Select
End Sub

Private Sub cmdNew_Click()
     Select Case LCase(cmdNew.Caption)
        Case "new"
            cmdNew.Caption = "Update"
            cmdEdit.Caption = "Cancel"
            cmdDelete.Enabled = False
            ClearControlsD
            fraDenomination.Enabled = True
            
        Case "update"
            If InsertNewD() = False Then Exit Sub
            cmdNew.Caption = "New"
            cmdEdit.Caption = "Edit"
            cmdDelete.Enabled = True
            fraDenomination.Enabled = False
            LoadDenominationsOfCurrency selCurr, True
    End Select
End Sub

Private Sub cmdNewC_Click()
    Select Case LCase(cmdNewC.Caption)
        Case "new"
            cmdNewC.Caption = "Update"
            cmdEditC.Caption = "Cancel"
            cmdDeleteC.Enabled = False
            cmdDenominations.Enabled = False
            ClearControls
            fraCurrency.Enabled = True
            
        Case "update"
            If InsertNew() = False Then Exit Sub
            cmdNewC.Caption = "New"
            cmdEditC.Caption = "Edit"
            cmdDeleteC.Enabled = True
            cmdDenominations.Enabled = True
            fraCurrency.Enabled = False
            LoadCurrencies
    End Select
    
End Sub

Private Function Update() As Boolean
    
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    
    If Trim(Me.txtCurrencyName.Text) = "" Then
        MsgBox "Enter the name of the currency", vbExclamation, TITLES
        Me.txtCurrencyName.SetFocus
        Exit Function
    Else
        selCurr.CurrencyName = Trim(Me.txtCurrencyName.Text)
    End If
    
    If Trim(Me.txtCurrencyCode.Text) = "" Then
        MsgBox "Enter the Currency Code", vbExclamation, TITLES
        Me.txtCurrencyCode.SetFocus
        Exit Function
    Else
        selCurr.CurrencyCode = Trim(Me.txtCurrencyCode.Text)
    End If

    
    If Trim(Me.txtCurrencyName.Text) = "" Then
        selCurr.ConversionRate = 0#
        
    Else
        If Not IsNumeric(Trim(Me.txtConversionRate.Text)) Then
            MsgBox "Enter a Numeric Conversion Rate for the currency", vbExclamation, TITLES
            Me.txtConversionRate.SetFocus
            Exit Function
        Else
            selCurr.ConversionRate = CSng(Trim(Me.txtConversionRate.Text))
        End If
    End If
    
    If Trim(Me.txtSymbol.Text) = "" Then
        MsgBox "Enter the Symbol for the currency", vbExclamation, TITLES
        Me.txtSymbol.SetFocus
        Exit Function
    Else
        selCurr.CurrencySymbol = Trim(Me.txtSymbol.Text)
        
    End If
    
    If chkIsBaseCurrency.value = vbChecked Then
        selCurr.IsBaseCurrency = True
    Else
        selCurr.IsBaseCurrency = False
    End If
    
    ''''''''''''''''
    ''-------check if other is of base currency
    
    Dim i2 As Integer
    Dim respo As String
    For i2 = 1 To pCurrencies.count
        Dim cc As String
        cc = selCurr.CurrencyName
        If pCurrencies.Item(i2).IsBaseCurrency Then
            If (selCurr.IsBaseCurrency And (pCurrencies.Item(i2).CurrencyID <> selCurr.CurrencyID)) Then
                ''MsgBox ("Making More than two Currencies as main Base currency is not allowed")
                respo = MsgBox(pCurrencies.Item(i2).CurrencyName & " Is already marked as base currency. Would you like to Mark the current Currency as Base currency?", vbYesNo + vbCritical)
                
'                chkIsBaseCurrency.value = False
'                selCurr.IsBaseCurrency = False
'                Update = False
                
                If respo = vbYes Then
                pCurrencies.Item(i2).IsBaseCurrency = False
                selCurr.IsBaseCurrency = True
                Else
                pCurrencies.Item(i2).IsBaseCurrency = True
                selCurr.IsBaseCurrency = False
                chkIsBaseCurrency.value = False
                End If
                
                GoTo nex
               '' Exit Function
            End If
        End If
    Next i2
    ''''''''''''''''
    
nex:
    
    retVal = selCurr.Update()
    If retVal = 0 Then
        MsgBox "The Currency has been updated successfully", vbInformation, TITLES
        Update = True
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred while updating the currency" & vbNewLine & err.Description, vbExclamation, TITLES
    Update = False
End Function

Private Function InsertNewD() As Boolean
    Dim newDenom As HRCORE.Denomination
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    
    Set newDenom = New HRCORE.Denomination
    If Trim(Me.txtDenomination.Text) = "" Then
        MsgBox "Enter the name of the Denomination", vbExclamation, TITLES
        Me.txtDenomination.SetFocus
        Exit Function
    Else
        newDenom.DenominationName = Trim(Me.txtDenomination.Text)
    End If
    
     If Trim(Me.txtValue.Text) = "" Then
        newDenom.DenominationValue = 0#
    Else
        If Not IsNumeric(Trim(Me.txtValue.Text)) Then
            MsgBox "Enter a Numeric Value for the Denomination", vbExclamation, TITLES
            Me.txtValue.SetFocus
            Exit Function
        Else
            newDenom.DenominationValue = CSng(Trim(Me.txtValue.Text))
        End If
    End If
    
    Set newDenom.ParentCurrency = selCurr
    
    retVal = newDenom.InsertNew()
    If retVal = 0 Then
        MsgBox "The new Denomination has been added successfully", vbInformation, TITLES
        InsertNewD = True
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred while creating a new denomination" & vbNewLine & err.Description, vbExclamation, TITLES
    InsertNewD = False
    
End Function

Private Function UpdateD() As Boolean
    
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
       
    If Trim(Me.txtDenomination.Text) = "" Then
        MsgBox "Enter the name of the Denomination", vbExclamation, TITLES
        Me.txtDenomination.SetFocus
        Exit Function
    Else
        selDenom.DenominationName = Trim(Me.txtDenomination.Text)
    End If
    
     If Trim(Me.txtValue.Text) = "" Then
        selDenom.DenominationValue = 0#
    Else
        If Not IsNumeric(Trim(Me.txtValue.Text)) Then
            MsgBox "Enter a Numeric Value for the Denomination", vbExclamation, TITLES
            Me.txtValue.SetFocus
            Exit Function
        Else
            selDenom.DenominationValue = CSng(Trim(Me.txtValue.Text))
        End If
    End If
    
    Set selDenom.ParentCurrency = selCurr
    
    retVal = selDenom.Update()
    If retVal = 0 Then
        MsgBox "The Denomination has been Updated successfully", vbInformation, TITLES
        UpdateD = True
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred while Updating the Denomination" & vbNewLine & err.Description, vbExclamation, TITLES
    UpdateD = False
    
End Function


Private Function InsertNew() As Boolean
    Dim newCurr As HRCORE.CCurrency
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    Set newCurr = New HRCORE.CCurrency
    
    If Trim(Me.txtCurrencyCode.Text) = "" Then
        MsgBox "Enter the Currency Code", vbExclamation, TITLES
        Me.txtCurrencyCode.SetFocus
        Exit Function
    Else
        newCurr.CurrencyCode = Trim(Me.txtCurrencyCode.Text)
    End If
    
    If Trim(Me.txtCurrencyName.Text) = "" Then
        MsgBox "Enter the name of the currency", vbExclamation, TITLES
        Me.txtCurrencyName.SetFocus
        Exit Function
    Else
        newCurr.CurrencyName = Trim(Me.txtCurrencyName.Text)
    End If
    
    If Trim(Me.txtCurrencyName.Text) = "" Then
        newCurr.ConversionRate = 1#
        
    Else
        If Not IsNumeric(Trim(Me.txtConversionRate.Text)) Then
            MsgBox "Enter a Numeric Conversion Rate for the currency", vbExclamation, TITLES
            Me.txtConversionRate.SetFocus
            Exit Function
        Else
            newCurr.ConversionRate = CSng(Trim(Me.txtConversionRate.Text))
        End If
    End If
    
    If Trim(Me.txtSymbol.Text) = "" Then
        MsgBox "Enter the Symbol for the currency", vbExclamation, TITLES
        Me.txtSymbol.SetFocus
        Exit Function
    Else
        newCurr.CurrencySymbol = Trim(Me.txtSymbol.Text)
        
    End If
    
    If chkIsBaseCurrency.value = vbChecked Then
        newCurr.IsBaseCurrency = True
         Dim i As Integer
         
         Dim respo As String
        For i = 1 To pCurrencies.count
            If pCurrencies.Item(i).IsBaseCurrency = True Then
               '' MsgBox ("Not More than 1 base currencies are allowed")
                respo = MsgBox(pCurrencies.Item(i).CurrencyName & " Is already marked as base currency. Would you like to Mark the current Currency as Base currency?", vbYesNo + vbCritical)
                
                If respo = vbYes Then
                pCurrencies.Item(i).IsBaseCurrency = False
                newCurr.IsBaseCurrency = True
                Else
                pCurrencies.Item(i).IsBaseCurrency = True
                newCurr.IsBaseCurrency = False
                End If
'                InsertNew = False
'                Exit Function
            End If
        Next i
    Else
        newCurr.IsBaseCurrency = False
    End If
   Dim sss As String
   sss = newCurr.CurrencyCode
   sss = newCurr.CurrencyName
   sss = newCurr.CurrencySymbol
   sss = newCurr.ConversionRate
   
   Dim newVal As Double
   newVal = newCurr.ConversionRate
   sss = newCurr.IsBaseCurrency
    retVal = newCurr.InsertNew()
    
    
    If retVal = 0 Then
        MsgBox "The New Currency has been added successfully", vbInformation, TITLES
        InsertNew = True
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred while creating a new currency" & vbNewLine & err.Description, vbExclamation, TITLES
    InsertNew = False
End Function


Private Sub ClearControls()
    Me.txtConversionRate.Text = ""
    Me.txtCurrencyName.Text = ""
    Me.txtSymbol.Text = ""
    Me.chkIsBaseCurrency.value = vbUnchecked
End Sub

Private Sub Form_Load()
    frmMain2.PositionTheFormWithoutEmpList Me
    
    Set pCurrencies = New HRCORE.Currencies
    Set pDenominations = New HRCORE.Denominations
    pDenominations.GetActiveDenominations
    
    LoadCurrencies
    
    sstCurrencies.TabVisible(1) = False
End Sub

Private Sub LoadDenominationsOfCurrency(ByVal TheCurrency As HRCORE.CCurrency, ByVal Refresh As Boolean)
    Dim TheDenoms As HRCORE.Denominations
    
    On Error GoTo ErrorHandler
    If Refresh = True Then
        pDenominations.GetActiveDenominations
    End If
    
    If Not TheCurrency Is Nothing Then
        Set TheDenoms = pDenominations.GetDenominationsByCurrencyID(TheCurrency.CurrencyID)
    End If
    
    PopulateDenominations TheDenoms
        
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while loading the Denominations" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub PopulateDenominations(ByVal TheDenominations As HRCORE.Denominations)
    Dim i As Long
    Dim ItemX As ListItem
    
    On Error GoTo ErrorHandler
    
    Me.lvwDenominations.ListItems.Clear
    
    If Not TheDenominations Is Nothing Then
        For i = 1 To TheDenominations.count
            Set ItemX = lvwDenominations.ListItems.add(, , TheDenominations.Item(i).DenominationName)
            ItemX.SubItems(1) = TheDenominations.Item(i).DenominationValue
            ItemX.SubItems(2) = TheDenominations.Item(i).ParentCurrency.CurrencyName
            ItemX.Tag = TheDenominations.Item(i).DenominationID
        Next i
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while populating denominations" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub ClearControlsD()
    Me.txtDenomination.Text = ""
    Me.txtValue.Text = ""
    
End Sub


Private Sub SetFieldsD(ByVal TheDenom As HRCORE.Denomination)
    ClearControlsD
    If Not (TheDenom Is Nothing) Then
        Me.txtDenomination.Text = TheDenom.DenominationName
        Me.txtValue.Text = TheDenom.DenominationValue
    End If
End Sub

Private Sub LoadCurrencies()
    pCurrencies.GetActiveCurrencies
    pCurrencies.GetAllCurrencies
    PopulateCurrencies pCurrencies
End Sub


Private Sub PopulateCurrencies(ByVal TheCurrs As HRCORE.Currencies)
    Dim i As Long
    Dim ItemX As ListItem
    
    On Error GoTo ErrorHandler
    Me.lvwCurrencies.ListItems.Clear
    If Not (TheCurrs Is Nothing) Then
        For i = 1 To TheCurrs.count
            If Not (TheCurrs.Item(i).Deleted = True) Then
            Set ItemX = Me.lvwCurrencies.ListItems.add(, , TheCurrs.Item(i).CurrencyName)
            ItemX.SubItems(1) = TheCurrs.Item(i).CurrencyCode
            ItemX.SubItems(2) = TheCurrs.Item(i).CurrencySymbol
            ItemX.SubItems(3) = TheCurrs.Item(i).ConversionRate
            ItemX.SubItems(4) = TheCurrs.Item(i).IsBaseCurrency
            ItemX.Tag = TheCurrs.Item(i).CurrencyID
            End If
        Next i
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "An error has occurred while populating existing currencies" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub


Private Sub SetFields(ByVal TheCurr As HRCORE.CCurrency)
    On Error GoTo ErrorHandler
    ClearControls
    If Not (TheCurr Is Nothing) Then
        Me.txtConversionRate.Text = TheCurr.ConversionRate
        Me.txtCurrencyName.Text = TheCurr.CurrencyName
        Me.txtCurrencyCode.Text = TheCurr.CurrencyCode
        Me.txtSymbol.Text = TheCurr.CurrencySymbol
        If TheCurr.IsBaseCurrency Then
            Me.chkIsBaseCurrency.value = vbChecked
        Else
            Me.chkIsBaseCurrency.value = vbUnchecked
        End If
    End If
    
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while populating the Currency details" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub lvwCurrencies_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set selCurr = Nothing
    If IsNumeric(Item.Tag) Then
        Set selCurr = pCurrencies.FindCurrencyByID(CLng(Item.Tag))
    End If
    SetFields selCurr
End Sub

Private Sub lvwDenominations_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set selDenom = Nothing
    If IsNumeric(Item.Tag) Then
        Set selDenom = pDenominations.FindDenominationByID(CLng(Item.Tag))
    End If
    
    SetFieldsD selDenom
    
End Sub
