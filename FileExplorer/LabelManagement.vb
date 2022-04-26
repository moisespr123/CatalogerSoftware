Public Class LabelManagement

    Private Sub LabelManagement_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PopulateListView()
    End Sub
    Private Sub PopulateListView()
        Dim Labels As Dictionary(Of String, String) = Form1.SQL.GetLabels()
        ListView1.Items.Clear()
        For Each label As String In Labels.Keys
            Dim item As ListViewItem = New ListViewItem(label)
            Dim subItems As ListViewItem.ListViewSubItem() = New ListViewItem.ListViewSubItem() {New ListViewItem.ListViewSubItem(item, Labels(label))}
            item.SubItems.AddRange(subItems)
            ListView1.Items.Add(item)
        Next
        ListView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim Label As ListView.SelectedIndexCollection = ListView1.SelectedIndices
        If ListView1.SelectedIndices.Count > 0 Then
            Dim Spindle As String = InputBox("Enter the Spindle Name/Number for the selected label", , ListView1.Items(Label(0)).SubItems(1).Text)
            If Not String.IsNullOrEmpty(Spindle) Then
                Form1.SQL.UpdateSpindle(ListView1.Items(Label(0)).Text, Spindle)
                PopulateListView()
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Label As ListView.SelectedIndexCollection = ListView1.SelectedIndices
        If ListView1.SelectedIndices.Count > 0 Then
            Dim New_Label As String = InputBox("Enter the new label name", , ListView1.Items(Label(0)).Text)
            If Not String.IsNullOrEmpty(New_Label) Then
                Form1.SQL.UpdateLabel(ListView1.Items(Label(0)).Text, New_Label)
                PopulateListView()
            End If
        End If
    End Sub

    Private Sub ListView1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListView1.MouseDoubleClick
        Dim Label As ListView.SelectedIndexCollection = ListView1.SelectedIndices
        If ListView1.SelectedIndices.Count > 0 Then
            Form1.SearchFunction(Form1.SQL.GetLabelContents(ListView1.Items(Label(0)).Text), ListView1.Items(Label(0)).Text)
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim DriveLetterToUse = InputBox("Enter the Drive Letter to use to check the checksums")
        Dim Label As ListView.SelectedIndexCollection = ListView1.SelectedIndices
        If ListView1.SelectedIndices.Count > 0 Then
            Dim MD5FileName As String = ListView1.Items(ListView1.SelectedIndices(0)).Text + ".md5"
            Dim SaveDialog As New SaveFileDialog With {.FileName = MD5FileName, .Filter = "MD5 Checksum|*.md5"}
            Dim result As MsgBoxResult = SaveDialog.ShowDialog()
            If result = MsgBoxResult.Ok Then
                Dim Files As Dictionary(Of Integer, FileClass) = Form1.SQL.GetLabelContents(ListView1.Items(Label(0)).Text)
                Dim ChecksumString As String = ""
                For Each file In Files.Keys
                    Dim Path As String() = Files(file).OriginalPath.Split(":")
                    Dim NewPath As String = Path(0).Replace(Path(0), DriveLetterToUse.Chars(0)) + ":" + Path(1)
                    ChecksumString = ChecksumString + Files(file).Checksum + " *" + NewPath + Environment.NewLine
                Next
                My.Computer.FileSystem.WriteAllText(SaveDialog.FileName, ChecksumString, False, New Text.UTF8Encoding(False))
            End If
            MsgBox("Label Content Checksums saved.")
        End If
    End Sub
End Class