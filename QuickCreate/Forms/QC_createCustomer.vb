Public Class QC_createCustomer

    Dim client_name As String
    Dim client_exists As Boolean

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim FolderPathMain As String

        client_name = TextBox1.Text
        client_exists = False

        If client_name = "" Then
            MsgBox("You need to enter a Company Name before continuing.")
            Exit Sub
        Else
            CheckNames()
            If client_exists = True Then
                MsgBox("This company already exists.")
                Exit Sub
            Else
                FolderPathMain = "F:\" & client_name
                MkDir(FolderPathMain)
                MkDir(FolderPathMain & "\External Documents")
                MkDir(FolderPathMain & "\Internal Documents")
                MkDir(FolderPathMain & "\General Documents")
                MkDir(FolderPathMain & "\Quotes")

                MsgBox("The Customer Folder has been created.")

                TextBox1.Text = ""
                Exit Sub
            End If
        End If

    End Sub

    Private Sub CheckNames()
        If Dir("F:\" & TextBox1.Text, vbDirectory) <> "" Then
            client_exists = True
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