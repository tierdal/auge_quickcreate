Public Class QC_createQuoteFolder

    Dim quote_number As String
    Dim quote_exists As Boolean

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim CurrentYear As String
        Dim FolderPathMain As String
        Dim FolderPathInternal As String
        Dim FolderPathExternal As String

        CurrentYear = Year(Now) & "\"

        quote_number = TextBox1.Text
        quote_exists = False

        If ComboBox1.Text = "" Then
            MsgBox("You need to select a Company Name before continuing.")
        Else
            If quote_number = "" Then
                MsgBox("You need to enter a Quote Number before continuing.")
            Else
                CheckQuote()
                If quote_exists = True Then
                    MsgBox("This Quote folder already exists.")
                Else
                    On Error Resume Next
                    FolderPathMain = "F:\" & ComboBox1.Text
                    MkDir(FolderPathMain & "\Quotes\" & CurrentYear)
                    MkDir(FolderPathMain & "\Quotes\" & CurrentYear & "\" & quote_number)
                    FolderPathExternal = FolderPathMain & "\Quotes\" & CurrentYear & "\" & quote_number & "\"
                    MkDir(FolderPathExternal & "Customer RFQ")
                    MkDir(FolderPathExternal & "Cost Analysis")
                    MkDir(FolderPathExternal & "Submitted Quotes")
                    MkDir(FolderPathExternal & "Vendor Quotes")

                    MsgBox("The Quote Folder has been created.")

                    ComboBox1.Text = ""
                    TextBox1.Text = ""

                    Dim folderName = From dir In IO.Directory.GetDirectories("F:\")
                                     Select IO.Path.GetFileName(dir)

                    ComboBox1.Items.AddRange(folderName.ToArray)

                End If
            End If
        End If
    End Sub

    Private Sub CheckQuote()
        Dim QuotePath As String
        Dim CurrentYear As String

        CurrentYear = Year(Now) & "\"

        QuotePath = "F:\" & ComboBox1.Text & "\Quotes\" & CurrentYear & "\" & TextBox1.Text

        If Dir(QuotePath, vbDirectory) <> "" Then
            quote_exists = True
            Exit Sub
        End If

    End Sub

    Private Sub QC_createQuoteFolder_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim folderName = From dir In IO.Directory.GetDirectories("F:\")
                         Select IO.Path.GetFileName(dir)

        ComboBox1.Items.AddRange(folderName.ToArray)

    End Sub

End Class