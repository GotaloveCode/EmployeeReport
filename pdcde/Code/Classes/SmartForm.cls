VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SmartForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private cH As Double    'Child Height
Private cW As Double    'Child Width
Private ccH As Double
Private ccW As Double
Private Rez As Boolean
Public wRatio As Double 'Width Ratio
Public hRatio As Double 'Height Ratio

Public Sub FResize(frm As Object)
    
    'Oscar Muhahala Tried to Decipher the meaning of this Procedure
    'on May 26, 2007 and made the following comments
    
    Dim pH As Double    'Parent Height
    Dim pW As Double    'Parent Width
    On Error Resume Next
    
    
    Dim rH As Double    'ratio of height
    Dim rW As Double    'ratio of width
    Dim con As Object
    
    If cH = 0 Then
        'Set the Childs Height and Width to that of Parent
        cH = pH
        cW = pW
    
        rH = 1
        rW = 1
        
        Exit Sub
    Else
        pH = cH
        pW = cW
        If frm.Name = "frmMain2" Then
            cH = frm.Height - 120
            cW = frm.Width - 120
        Else
            cH = frm.Height
            cW = frm.Width
        End If
        rH = cH / pH
        rW = cW / pW
        If frm.Name = "frmMain2" Then
            wRatio = rW
            hRatio = rH
    End If
End If



For Each con In frm

    If TypeOf con Is Object  Then
    If con.Container.Name <> "fraLogo" Then
        If TypeOf con.Container Is Frame Then
            If TypeOf con Is Frame Or TypeOf con Is ListView Or TypeOf con Is TreeView Or TypeOf con Is Image Or TypeOf con Is SSTab Then
                con.Move (con.Left * rW), (con.Top * rH), con.Width * rW, (con.Height * rH)
            Else
                If frm.Name = "frmMain2" Then
                    con.Move (con.Left * rW), (con.Top * rH), con.Width * rW, (con.Height * rH) '
                Else
                    con.Move (con.Left * rW), (con.Top * rH), con.Width * rW
                End If

            End If


        ElseIf TypeOf con.Container Is SSTab Then
            If TypeOf con Is Frame Or TypeOf con Is ListView Or TypeOf con Is TreeView Or TypeOf con Is Image Or TypeOf con Is SSTab Then
                con.Move (con.Left), (con.Top * rH), con.Width * rW, (con.Height * rH)
            Else
                If TypeOf con Is CommandButton Then
                    If frm.Name = "frmMain2" Then
                        con.Move (con.Left * rW), (con.Top * rH), con.Width * rW, (con.Height * rH) '
                    Else
                        con.Move (con.Left * rW), (con.Top * rH), con.Width * rW
                    End If

                Else
                    If frm.Name = "frmMain2" Then
                        con.Move (con.Left), (con.Top * rH), con.Width * rW, (con.Height * rH) '
                    Else
                        con.Move (con.Left), (con.Top * rH), con.Width * rW
                    End If
                End If
            End If


        Else
            If TypeOf con Is Frame Or TypeOf con Is ListView Or TypeOf con Is TreeView Or TypeOf con Is Image Or TypeOf con Is SSTab Then
                con.Move (con.Left * rW), (con.Top * rH), con.Width * rW, (con.Height * rH)
            Else
                If frm.Name = "frmMain2" Then
                    con.Move (con.Left * rW), (con.Top * rH), con.Width * rW, (con.Height * rH) '
                Else
                    con.Move (con.Left * rW), (con.Top * rH), con.Width * rW
                End If
            End If


        End If
        End If
    End If

Next con



      





Exit Sub
Hell:
MsgBox err.Description, vbExclamation
End Sub



Public Sub FReset(frm As Object)
    On Error GoTo Hell
    Dim mm As Variant
    
    Rez = True

    cH = frm.Height
    cW = frm.Width
    
    If frm.Name <> "frmMain2" Then
        frm.Height = frm.Height * hRatio
        frm.Width = frm.Width * wRatio
    Else
        wRatio = 1
        hRatio = 1
       
    End If
    
   Exit Sub
Hell:
   MsgBox err.Description, vbExclamation
End Sub

Public Sub Fboundry(frm As Object, ByVal maxHeights As Long, maxWidths As Long, minWidths As Long, minHeights As Long, Mleft As Long, MTop As Long)

MAX_HEIGHT = maxHeights / 15
MAX_WIDTH = maxWidths / 15
MIN_HEIGHT = minHeights / 15
MIN_WIDTH = minWidths / 15
MovLeft = Mleft / 15
MovTop = MTop / 15

End Sub


