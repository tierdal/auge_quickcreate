Public Class QC_mainMenu

    'EXIT APPLICATION
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        End
    End Sub

    'DEFINE VERSION
    Private Sub QC_mainMenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = Me.Text & " v2.01"
    End Sub

    'START PDF HELP DOC
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Process.Start("F:\Templates\qc_lwi.pdf")
    End Sub

    '---------------------------- Menu Display Options
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        QC_createCustomer.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        QC_createVendor.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        QC_createPOfolder.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        QC_createVendorFolder.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        QC_createQuoteFolder.Show()
    End Sub
End Class
