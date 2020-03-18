VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmLocations 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Locations"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   Icon            =   "frmLocations.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab sstCountries 
      Height          =   6420
      Left            =   150
      TabIndex        =   17
      Top             =   675
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   11324
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Countries"
      TabPicture(0)   =   "frmLocations.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdClose"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdLocations"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdDeleteC"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdEditC"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdNewC"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraExisting"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraCountry"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Locations"
      TabPicture(1)   =   "frmLocations.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraLocations"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraLocation"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtParentCountry"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdNew"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdEdit"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdDelete"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdBack"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   465
         Left            =   -69300
         TabIndex        =   8
         Top             =   5700
         Width           =   1215
      End
      Begin VB.CommandButton cmdLocations 
         Caption         =   "Locations..."
         Height          =   465
         Left            =   -71175
         TabIndex        =   7
         Top             =   5700
         Width           =   1440
      End
      Begin VB.CommandButton cmdDeleteC 
         Caption         =   "Delete"
         Height          =   465
         Left            =   -72525
         TabIndex        =   6
         Top             =   5700
         Width           =   1140
      End
      Begin VB.CommandButton cmdEditC 
         Caption         =   "Edit"
         Height          =   465
         Left            =   -73575
         TabIndex        =   5
         Top             =   5700
         Width           =   990
      End
      Begin VB.CommandButton cmdNewC 
         Caption         =   "New"
         Height          =   465
         Left            =   -74700
         TabIndex        =   4
         Top             =   5700
         Width           =   990
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back To Countries"
         Height          =   465
         Left            =   5250
         TabIndex        =   16
         Top             =   5820
         Width           =   1590
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   465
         Left            =   2850
         TabIndex        =   15
         Top             =   5820
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   465
         Left            =   1425
         TabIndex        =   14
         Top             =   5820
         Width           =   1215
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   465
         Left            =   150
         TabIndex        =   13
         Top             =   5820
         Width           =   1140
      End
      Begin VB.TextBox txtParentCountry 
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
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   525
         Width           =   5490
      End
      Begin VB.Frame fraLocation 
         Height          =   1140
         Left            =   150
         TabIndex        =   23
         Top             =   1050
         Width           =   6765
         Begin VB.TextBox txtLocationName 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1395
            TabIndex        =   11
            Top             =   675
            Width           =   5130
         End
         Begin VB.TextBox txtLocationCode 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1395
            TabIndex        =   10
            Top             =   300
            Width           =   2640
         End
         Begin VB.Label Label7 
            Caption         =   "Location Name:"
            Height          =   165
            Left            =   150
            TabIndex        =   25
            Top             =   750
            Width           =   1155
         End
         Begin VB.Label Label6 
            Caption         =   "Location Code:"
            Height          =   165
            Left            =   195
            TabIndex        =   24
            Top             =   300
            Width           =   1185
         End
      End
      Begin VB.Frame fraLocations 
         Caption         =   "Existing Locations:"
         Height          =   3360
         Left            =   150
         TabIndex        =   22
         Top             =   2325
         Width           =   6765
         Begin MSComctlLib.ListView lvwLocations 
            Height          =   2955
            Left            =   150
            TabIndex        =   12
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
               Text            =   "Location Code"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Location Name"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Country"
               Object.Width           =   5292
            EndProperty
         End
      End
      Begin VB.Frame fraExisting 
         Caption         =   "Existing Countries:"
         Height          =   4050
         Left            =   -74850
         TabIndex        =   19
         Top             =   1440
         Width           =   6765
         Begin MSComctlLib.ListView lvwCountries 
            Height          =   3525
            Left            =   150
            TabIndex        =   3
            Top             =   375
            Width           =   6465
            _ExtentX        =   11404
            _ExtentY        =   6218
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Country Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Country Name"
               Object.Width           =   8819
            EndProperty
         End
      End
      Begin VB.Frame fraCountry 
         Height          =   855
         Left            =   -74850
         TabIndex        =   18
         Top             =   450
         Width           =   6765
         Begin VB.TextBox txtCountryName 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   3375
            TabIndex        =   2
            Top             =   345
            Width           =   3240
         End
         Begin VB.TextBox txtCountryCode 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1185
            TabIndex        =   1
            Top             =   345
            Width           =   825
         End
         Begin VB.Label Label3 
            Caption         =   "Country Name:"
            Height          =   240
            Left            =   2280
            TabIndex        =   21
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Country Code:"
            Height          =   240
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Country:"
         Height          =   240
         Left            =   225
         TabIndex        =   26
         Top             =   600
         Width           =   840
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Locations"
      BeginProperty Font 
         Name            =   "Tahoma"
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
      Top             =   75
      Width           =   2895
   End
End
Attribute VB_Name = "frmLocations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pCountries As HRCORE.Countries
Private selCountry As HRCORE.Country
Private pLocations As HRCORE.Locations
Private selLocation As HRCORE.Location


Private Sub cmdBack_Click()
    Me.sstCountries.TabEnabled(0) = True
    sstCountries.TabVisible(1) = False
    sstCountries.Tab = 0
    
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdDeleteC_Click()
    Dim retVal As Long
    Dim resp As Long
    
    Select Case LCase(cmdDeleteC.Caption)
        Case "delete"
            If selCountry Is Nothing Then
                MsgBox "Select the Country you want to delete", vbExclamation, TITLES
            Else
                resp = MsgBox("Are you sure you want to delete the selected Country?", vbYesNo + vbQuestion, TITLES)
                If resp = vbYes Then
                    retVal = selCountry.Delete()
                Set selCountry = Nothing
                    txtCountryCode.Text = ""
                    txtCountryName.Text = ""
                    LoadCountries
                End If
            End If
        Case "cancel"
            cmdNewC.Enabled = True
            cmdEditC.Caption = "Edit"
            cmdDeleteC.Caption = "Delete"
            fraCountry.Enabled = False
            cmdLocations.Enabled = True
            txtCountryCode.Enabled = False
            txtCountryName.Enabled = False
            LoadCountries
    End Select
    
    
End Sub

Private Sub cmdDelete_Click()
    Dim retVal As Long
    Dim resp As Long
    
    Select Case LCase(cmdDelete.Caption)
        Case "delete"
            If selLocation Is Nothing Then
                MsgBox "Select the Location you want to delete", vbExclamation, TITLES
            Else
                resp = MsgBox("Are yo sure you want to delete the selected Location?", vbYesNo + vbQuestion, TITLES)
                If resp = vbYes Then
                    retVal = selLocation.Delete()
                    LoadLocationsOfCountry selCountry, True
                End If
            End If
            
        Case "cancel"
            cmdNew.Enabled = True
            cmdEdit.Caption = "Edit"
            cmdDelete.Caption = "Delete"
            cmdBack.Enabled = True
            fraLocation.Enabled = False
            LoadLocationsOfCountry selCountry, False
            txtLocationCode.Enabled = False
            txtLocationName.Enabled = False
    End Select
End Sub

Private Sub cmdLocations_Click()
lvwLocations.ListItems.Clear
    If selCountry Is Nothing Then
        MsgBox "Select the Country to set Locations for", vbInformation, TITLES
    Else
    
        Me.txtParentCountry.Text = selCountry.CountryName
        LoadLocationsOfCountry selCountry, False
        sstCountries.TabVisible(1) = True
        sstCountries.TabEnabled(0) = False
        sstCountries.Tab = 1
    End If
End Sub

Private Sub cmdEditC_Click()
    Select Case LCase(cmdEditC.Caption)
        Case "edit"
            cmdNewC.Enabled = False
            cmdEditC.Caption = "Update"
            cmdDeleteC.Caption = "Cancel"
            cmdLocations.Enabled = False
            fraCountry.Enabled = True
            txtCountryCode.Enabled = True
            txtCountryName.Enabled = True
            
        Case "update"
            If Update() = False Then Exit Sub
            cmdNewC.Enabled = True
            cmdEditC.Caption = "Edit"
            cmdDeleteC.Caption = "Delete"
            cmdLocations.Enabled = True
            fraCountry.Enabled = False
            txtCountryCode.Enabled = False
            txtCountryName.Enabled = False
            LoadCountries
            
        Case "cancel"   'cancels a new operation
            cmdNewC.Caption = "New"
            cmdEditC.Caption = "Edit"
            cmdDeleteC.Enabled = True
            fraCountry.Enabled = False
            cmdLocations.Enabled = True
            txtCountryCode.Enabled = False
            txtCountryName.Enabled = False
            
    End Select
End Sub

Private Sub cmdEdit_Click()
    Select Case LCase(cmdEdit.Caption)
        Case "edit"
            cmdNew.Enabled = False
            cmdEdit.Caption = "Update"
            cmdDelete.Caption = "Cancel"
            cmdBack.Enabled = False
            fraLocation.Enabled = True
            txtLocationCode.Enabled = True
            txtLocationName.Enabled = True
        Case "update"
            If UpdateL() = False Then Exit Sub
            cmdNew.Enabled = True
            cmdEdit.Caption = "Edit"
            cmdDelete.Caption = "Delete"
            fraLocation.Enabled = False
            cmdBack.Enabled = True
            txtLocationCode.Enabled = False
            txtLocationName.Enabled = False
            LoadLocationsOfCountry selCountry, True
            
        Case "cancel"   'cancels a new operation
            cmdNew.Caption = "New"
            cmdEdit.Caption = "Edit"
            cmdDelete.Enabled = True
            cmdBack.Enabled = True
            txtLocationCode.Enabled = False
            txtLocationName.Enabled = False
            fraLocation.Enabled = False
            
    End Select
End Sub

Private Sub cmdNew_Click()
     Select Case LCase(cmdNew.Caption)
        Case "new"
            cmdNew.Caption = "Update"
            cmdEdit.Caption = "Cancel"
            cmdDelete.Enabled = False
            ClearControlsL
              fraLocation.Enabled = True
              txtLocationCode.Enabled = True
              txtLocationCode.SetFocus
              txtLocationName.Enabled = True
        Case "update"
            If InsertNewL() = False Then Exit Sub
            cmdNew.Caption = "New"
            cmdEdit.Caption = "Edit"
            cmdLocations.Enabled = True
            txtLocationCode.Enabled = False
            txtLocationName.Enabled = False
            txtLocationCode.Text = ""
            txtLocationName.Text = ""
        
            cmdDelete.Enabled = True
            fraLocation.Enabled = False
            LoadLocationsOfCountry selCountry, True
    End Select
End Sub

Private Sub cmdNewC_Click()
    Select Case LCase(cmdNewC.Caption)
        Case "new"
            cmdNewC.Caption = "Update"
            cmdEditC.Caption = "Cancel"
            cmdDeleteC.Enabled = False
            cmdLocations.Enabled = False
            ClearControls
            fraCountry.Enabled = True
           txtCountryCode.Enabled = True
           txtCountryCode.SetFocus
           txtCountryName.Enabled = True
        Case "update"
            If InsertNew() = False Then Exit Sub
            cmdNewC.Caption = "New"
            txtLocationCode.Enabled = False
            txtLocationName.Enabled = False
            cmdEditC.Caption = "Edit"
            cmdDeleteC.Enabled = True
           
           txtCountryCode.Enabled = False
           txtCountryName.Enabled = False
            LoadCountries
    End Select
    
End Sub

Private Function Update() As Boolean
    
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    
    If Trim(Me.txtCountryName.Text) = "" Then
        MsgBox "Enter the name of the Country", vbExclamation, TITLES
        Me.txtCountryName.SetFocus
        Exit Function
    Else
        selCountry.CountryName = Trim(Me.txtCountryName.Text)
    End If
    
    If Trim(Me.txtCountryCode.Text) = "" Then
        MsgBox "Enter the Country Code", vbExclamation, TITLES
        Me.txtCountryCode.SetFocus
        Exit Function
    Else
        selCountry.CountryCode = Trim(Me.txtCountryCode.Text)
    End If
    
    
    retVal = selCountry.Update()
    If retVal = 0 Then
        MsgBox "The Country has been UpdateL successfully", vbInformation, TITLES
        Update = True
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred while updating the Country" & vbNewLine & err.Description, vbExclamation, TITLES
    Update = False
End Function

Private Function InsertNewL() As Boolean
    Dim newLocation As HRCORE.Location
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    
    Set newLocation = New HRCORE.Location
    If Trim(Me.txtLocationCode.Text) = "" Then
        MsgBox "Enter the Location Code", vbExclamation, TITLES
        Me.txtLocationCode.SetFocus
        Exit Function
    Else
        newLocation.LocationCode = Trim(Me.txtLocationCode.Text)
    End If
    
    
    If Trim(Me.txtLocationName.Text) = "" Then
        MsgBox "Enter the Location Name", vbExclamation, TITLES
        Me.txtLocationName.SetFocus
        Exit Function
    Else
        newLocation.LocationName = Trim(Me.txtLocationName.Text)
    End If
    
    Set newLocation.Country = selCountry
    
    retVal = newLocation.InsertNew()
    If retVal = 0 Then
        MsgBox "The new Location has been added successfully", vbInformation, TITLES
        InsertNewL = True
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred while creating a new Location" & vbNewLine & err.Description, vbExclamation, TITLES
    InsertNewL = False
    
End Function

Private Function UpdateL() As Boolean
    
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
       
    If Trim(Me.txtLocationCode.Text) = "" Then
        MsgBox "Enter the Location Code", vbExclamation, TITLES
        Me.txtLocationCode.SetFocus
        Exit Function
    Else
        selLocation.LocationCode = Trim(Me.txtLocationCode.Text)
    End If
    
    If Trim(Me.txtLocationName.Text) = "" Then
        MsgBox "Enter the Location Name", vbExclamation, TITLES
        Me.txtLocationName.SetFocus
        Exit Function
    Else
        selLocation.LocationName = Trim(Me.txtLocationName.Text)
    End If
    
    Set selLocation.Country = selCountry
    
    retVal = selLocation.Update()
    If retVal = 0 Then
        MsgBox "The Location has been UpdateL successfully", vbInformation, TITLES
        UpdateL = True
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred while Updating the Location" & vbNewLine & err.Description, vbExclamation, TITLES
    UpdateL = False
    
End Function


Private Function InsertNew() As Boolean
    Dim newCountry As HRCORE.Country
    Dim retVal As Long
    
    On Error GoTo ErrorHandler
    Set newCountry = New HRCORE.Country
    
    If Trim(Me.txtCountryCode.Text) = "" Then
        MsgBox "Enter the Country Code", vbExclamation, TITLES
        Me.txtCountryCode.SetFocus
        Exit Function
    Else
        newCountry.CountryCode = Trim(Me.txtCountryCode.Text)
    End If
    
    If Trim(Me.txtCountryName.Text) = "" Then
        MsgBox "Enter the name of the Country", vbExclamation, TITLES
        Me.txtCountryName.SetFocus
        Exit Function
    Else
        newCountry.CountryName = Trim(Me.txtCountryName.Text)
    End If
    
    retVal = newCountry.InsertNew()
    If retVal = 0 Then
        MsgBox "The New Country has been added successfully", vbInformation, TITLES
        InsertNew = True
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred while creating a new Country" & vbNewLine & err.Description, vbExclamation, TITLES
    InsertNew = False
End Function


Private Sub ClearControls()
    Me.txtCountryName.Text = ""
    Me.txtCountryCode.Text = ""
    
End Sub

Private Sub Form_Load()
    frmMain2.PositionTheFormWithoutEmpList Me
    
    Set pCountries = New HRCORE.Countries
    Set pLocations = New HRCORE.Locations
    pLocations.GetActiveLocations
    
    LoadCountries
    
    sstCountries.TabVisible(1) = False
End Sub

Private Sub LoadLocationsOfCountry(ByVal TheCountry As HRCORE.Country, ByVal Refresh As Boolean)
    Dim TheLocations As HRCORE.Locations
    
    On Error GoTo ErrorHandler
    If Refresh = True Then
        pLocations.GetActiveLocations
    End If
    
    If Not TheCountry Is Nothing Then
        Set TheLocations = pLocations.GetLocationsByCountryID(TheCountry.CountryID)
    End If
    
    PopulateLocations TheLocations
        
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while loading the Locations" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub PopulateLocations(ByVal TheLocations As HRCORE.Locations)
    Dim i As Long
    Dim ItemX As ListItem
    
    On Error GoTo ErrorHandler
    
    Me.lvwLocations.ListItems.Clear
    
    If Not TheLocations Is Nothing Then
        For i = 1 To TheLocations.count
        If Not (TheLocations.Item(i).Country.Deleted) Then
            Set ItemX = lvwLocations.ListItems.add(, , TheLocations.Item(i).LocationCode)
            ItemX.SubItems(1) = TheLocations.Item(i).LocationName
            ItemX.SubItems(2) = TheLocations.Item(i).Country.CountryName
            ItemX.Tag = TheLocations.Item(i).LocationID
            End If
        Next i
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while populating Locations" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub ClearControlsL()
    Me.txtLocationCode.Text = ""
    Me.txtLocationName.Text = ""
    
End Sub


Private Sub SetFieldsL(ByVal TheLocation As HRCORE.Location)
    ClearControlsL
    If Not (TheLocation Is Nothing) Then
        Me.txtLocationCode.Text = TheLocation.LocationCode
        Me.txtLocationName.Text = TheLocation.LocationName
    End If
End Sub

Private Sub LoadCountries()
    pCountries.GetActiveCountries
    PopulateCountries pCountries
End Sub


Private Sub PopulateCountries(ByVal TheCountries As HRCORE.Countries)
    Dim i As Long
    Dim ItemX As ListItem
    
    On Error GoTo ErrorHandler
    Me.lvwCountries.ListItems.Clear
    If Not (TheCountries Is Nothing) Then
        For i = 1 To TheCountries.count
            Set ItemX = Me.lvwCountries.ListItems.add(, , TheCountries.Item(i).CountryCode)
            ItemX.SubItems(1) = TheCountries.Item(i).CountryName
            ItemX.Tag = TheCountries.Item(i).CountryID
        Next i
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "An error has occurred while populating existing Countries" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub


Private Sub SetFields(ByVal TheCountry As HRCORE.Country)
    On Error GoTo ErrorHandler
    ClearControls
    If Not (TheCountry Is Nothing) Then
        Me.txtCountryCode.Text = TheCountry.CountryCode
        Me.txtCountryName.Text = TheCountry.CountryName
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error has occurred while populating the Country details" & vbNewLine & err.Description, vbExclamation, TITLES
End Sub

Private Sub lvwCountries_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set selCountry = Nothing
    If IsNumeric(Item.Tag) Then
        Set selCountry = pCountries.FindCountryByID(CLng(Item.Tag))
    End If
    SetFields selCountry
End Sub

Private Sub lvwLocations_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set selLocation = Nothing
    If IsNumeric(Item.Tag) Then
        Set selLocation = pLocations.FindLocationByID(CLng(Item.Tag))
    End If
    
    SetFieldsL selLocation
    
End Sub
