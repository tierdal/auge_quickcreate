Public Class QC_createVendorFolder

    Dim po_number As String
    Dim po_exists As Boolean

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim CurrentYear As String
        Dim FolderPathMain As String

        CurrentYear = Year(Now) & "\"

        po_number = TextBox1.Text
        po_exists = False

        If ComboBox1.Text = "" Then
            MsgBox("You need to select a Company Name before continuing.")
        Else
            If po_number = "" Then
                MsgBox("You need to enter a PO Number before continuing.")
            Else
                CheckAugePO()
                If po_exists = True Then
                    MsgBox("This PO folder already exists.")
                Else
                    On Error Resume Next
                    FolderPathMain = "V:\" & ComboBox1.Text
                    MkDir(FolderPathMain & "\" & CurrentYear)
                    MkDir(FolderPathMain & "\" & CurrentYear & "\" & po_number)

                    MsgBox("The PO Folder has been created.")

                    ComboBox1.Text = ""
                    TextBox1.Text = ""

                    Dim folderName = From dir In IO.Directory.GetDirectories("V:\")
                                     Select IO.Path.GetFileName(dir)

                    ComboBox1.Items.AddRange(folderName.ToArray)

                End If
            End If
        End If
    End Sub

    Private Sub CheckAugePO()
        Dim PO_Path As String
        Dim CurrentYear As String

        CurrentYear = Year(Now) & "\"

        PO_Path = "V:\" & ComboBox1.Text & "\" & CurrentYear & "\" & TextBox1.Text

        If Dir(PO_Path, vbDirectory) <> "" Then
            po_exists = True
            Exit Sub
        End If
    End Sub

    Private Sub QC_createVendorFolder_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim folderName = From dir In IO.Directory.GetDirectories("V:\")
                         Select IO.Path.GetFileName(dir)

        ComboBox1.Items.AddRange(folderName.ToArray)

    End Sub

End Class