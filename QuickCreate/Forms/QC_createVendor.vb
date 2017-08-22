Public Class QC_createVendor

    Dim vendor_name As String
    Dim vendor_exists As Boolean

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim FolderPathMain As String

        vendor_name = TextBox1.Text
        vendor_exists = False

        If vendor_name = "" Then
            MsgBox("You need to enter a Vendor Name before continuing.")
        Else
            CheckVendorNames()
            If vendor_exists = True Then
                MsgBox("This company already exists.")
            Else
                FolderPathMain = "V:\" & vendor_name
                MkDir(FolderPathMain)
                MkDir(FolderPathMain & "\General Documents")

                MsgBox("The Vendor Folder has been created.")

                TextBox1.Text = ""
            End If
        End If
    End Sub

    Private Sub CheckVendorNames()
        If Dir("V:\" & TextBox1.Text, vbDirectory) <> "" Then
            vendor_exists = True
            Exit Sub
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        Dim KeyAsciiString As String
        KeyAsciiString = e.KeyChar

        If KeyAsciiString = "/" Or KeyAsciiString = "\" Or KeyAsciiString = ":" Or KeyAsciiString = "*" Or KeyAsciiString = "?" Or KeyAsciiString = """" Or KeyAsciiString = "<" Or KeyAsciiString = ">" Or KeyAsciiString = "|" Then
            MsgBox("This key is not allowed: " & KeyAsciiString)
            e.Handled = True
        End If
    End Sub

End Class