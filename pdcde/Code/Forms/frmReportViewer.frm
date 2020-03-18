VERSION 5.00
Object = "{FB992564-9055-42B5-B433-FEA84CEA93C4}#11.0#0"; "crviewer.dll"
Begin VB.Form frmReports 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reports"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11010
   Icon            =   "frmReportViewer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   11010
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CrystalActiveXReportViewerLib11Ctl.CrystalActiveXReportViewer CRViewer1 
      Height          =   8655
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11415
      _cx             =   20135
      _cy             =   15266
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   1033
   End
   Begin VB.CommandButton cmdPrinterSetup 
      Appearance      =   0  'Flat
      Default         =   -1  'True
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
      Left            =   8280
      MaskColor       =   &H8000000F&
      Picture         =   "frmReportViewer.frx":0742
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Printer Setup"
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdPrinterSetup_Click()
    On Error GoTo Hell
    R.PrinterSetup hWnd
    Exit Sub
Hell:
End Sub

Private Sub Form_Load()
    On Error GoTo Hell
'    oSmart.FReset Me

    CConnect.CColor Me, MyColor
    
    Me.Left = 0
    Me.Top = o
    Me.Height = Screen.Height
    
    Me.Width = Screen.Width
    CRViewer1.Width = Screen.Width
    CRViewer1.Height = Screen.Height - 300
    
    Exit Sub
Hell:
    MsgBox err.Description, vbExclamation
End Sub

Private Sub Form_Resize()
    On Error GoTo Hell
    
    oSmart.FResize Me
    
    With Me
        CRViewer1.Left = 0
        CRViewer1.Top = 0
        CRViewer1.Height = .Height
        CRViewer1.Width = .Width
    End With

    Exit Sub
Hell:
    MsgBox err.Description, vbExclamation
End Sub
