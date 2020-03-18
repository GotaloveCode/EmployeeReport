Attribute VB_Name = "ModBackground"
'
'
'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'
'Private m_background As NetFX20Wrapper.BackgroundWorkerWrapper
'
'Public Sub StartBackground(background As NetFX20Wrapper.BackgroundWorkerWrapper, argument As Variant)
'    Set m_background = background
'    m_background.RunWorkerAsync AddressOf BackgroundWork, argument
'End Sub
'
'Public Sub BackgroundWork(ByRef argument As Variant, ByRef e As NetFX20Wrapper.RunWorkerCompletedEventArgsWrapper)
'    On Error GoTo eh
'
'    'Err.Raise 1, "", "Something Bad Happened"  ' Force error to test error handling
'
'    Dim i As Integer
'    For i = 1 To 20
'        Sleep 500
'        m_background.ReportProgress i * 5
'        If m_background.CancellationPending Then
'            e.Cancelled = True
'            Exit Sub
'        End If
'    Next
'
'
'
'''------------
'
''       If (objEmployeeBankAccounts Is Nothing) Then
''             Set objEmployeeBankAccounts = New EmployeeBankAccounts2
''             objEmployeeBankAccounts.GetActiveEmployeeBankAccounts
''             Set myEmployeeBankAccounts = objEmployeeBankAccounts
''             Set pEmployeeBankAccounts = objEmployeeBankAccounts
''       End If
'''----------------------
'
'
'
'    e.SetResult argument
'    Exit Sub
'
'eh:
'    e.Error.Number = err.Number
'    e.Error.Description = err.Description
'End Sub
'
'
'
