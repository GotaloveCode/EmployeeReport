Public Class Form2

    Private Sub CrystalReportViewer1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CrystalReportViewer1.Load
        Me.Left = 0
        Me.Top = 0
        Me.Height = 10000
        Me.Width = 18408

    End Sub

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CrystalReportViewer1.ReportSource = ("C:\Users\Martin\Desktop\working\EMPLOYEES.rpt")
        CrystalReportViewer1.Show()
    End Sub
End Class