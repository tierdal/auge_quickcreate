Public Class QC_createCustomer

    Dim client_name As String
    Dim client_exists As Boolean

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim FolderPathMain As String

        client_name = TextBox1.Text
        client_exists = False

        If client_name = "" Then
            MsgBox("You need to enter a Company Name before continuing.")
        Else
            CheckNames()
            If client_exists = True Then
                MsgBox("This company already exists.")
            Else
                FolderPathMain = "F:\" & client_name
                MkDir(FolderPathMain)
                MkDir(FolderPathMain & "\External Documents")
                MkDir(FolderPathMain & "\Internal Documents")
                MkDir(FolderPathMain & "\General Documents")
                MkDir(FolderPathMain & "\Quotes")

                MsgBox("The Customer Folder has been created.")

                TextBox1.Text = ""
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
        Dim KeyAscii As String

        KeyAscii = 0

        Select Case KeyAscii
            Case 47, 92, 58, 42, 63, 34, 60, 62, 124
                MsgBox("Invalid character. The following character is not allowed: " & Chr(KeyAscii))
                KeyAscii = 0
        End Select
    End Sub
End Class