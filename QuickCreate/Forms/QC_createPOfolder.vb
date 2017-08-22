Public Class QC_createPOfolder

    Dim po_number As String
    Dim po_exists As Boolean

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim CurrentYear As String
        Dim FolderPathMain As String
        Dim FolderPathInternal As String
        Dim FolderPathExternal As String

        CurrentYear = Year(Now) & "\"

        po_number = TextBox1.Text
        po_exists = False

        If ComboBox1.Text = "" Then
            MsgBox("You need to select a Company Name before continuing.")
        Else
            If po_number = "" Then
                MsgBox("You need to enter a PO Number before continuing.")
            Else
                CheckPO()
                If po_exists = True Then
                    MsgBox("This PO folder already exists.")
                Else
                    On Error Resume Next
                    FolderPathMain = "F:\" & ComboBox1.Text
                    MkDir(FolderPathMain & "\External Documents\" & CurrentYear)
                    MkDir(FolderPathMain & "\Internal Documents\" & CurrentYear)
                    MkDir(FolderPathMain & "\External Documents\" & CurrentYear & "\" & po_number)
                    MkDir(FolderPathMain & "\Internal Documents\" & CurrentYear & "\" & po_number)
                    FolderPathExternal = FolderPathMain & "\External Documents\" & CurrentYear & "\" & po_number & "\"
                    MkDir(FolderPathExternal & "Invoices")
                    MkDir(FolderPathExternal & "Quality Documents")
                    MkDir(FolderPathExternal & "Shipping Documents")
                    MkDir(FolderPathExternal & "Customer Purchase Order")
                    FolderPathInternal = FolderPathMain & "\Internal Documents\" & CurrentYear & "\" & po_number & "\"
                    MkDir(FolderPathInternal & "Warehouse Packet")
                    MkDir(FolderPathInternal & "Other Documents")
                    MkDir(FolderPathInternal & "Production Orders")
                    MkDir(FolderPathInternal & "Purchase Orders to Vendors")
                    MkDir(FolderPathInternal & "Sales Orders")

                    MsgBox("The PO Folder has been created.")

                    ComboBox1.Text = ""
                    TextBox1.Text = ""

                    Dim folderName = From dir In IO.Directory.GetDirectories("F:\")
                                     Select IO.Path.GetFileName(dir)

                    ComboBox1.Items.AddRange(folderName.ToArray)

                End If
            End If
        End If
    End Sub

    Private Sub CheckPO()
        Dim PathExternal As String
        Dim PathInternal As String
        Dim CurrentYear As String

        CurrentYear = Year(Now) & "\"

        PathExternal = "F:\" & ComboBox1.Text & "\External Documents\" & CurrentYear & "\" & TextBox1.Text
        PathInternal = "F:\" & ComboBox1.Text & "\Internal Documents\" & CurrentYear & "\" & TextBox1.Text

        If Dir(PathExternal, vbDirectory) <> "" Then
            po_exists = True
            Exit Sub
        End If
        If Dir(PathInternal, vbDirectory) <> "" Then
            po_exists = True
            Exit Sub
        End If
    End Sub

    Private Sub QC_createPOfolder_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim folderName = From dir In IO.Directory.GetDirectories("F:\")
                         Select IO.Path.GetFileName(dir)

        ComboBox1.Items.AddRange(folderName.ToArray)

    End Sub
End Class