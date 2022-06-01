Public Class LabelManagement
    Private lvwColumnSorter As ListViewColumnSorter
    Private Sub LabelManagement_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lvwColumnSorter = New ListViewColumnSorter()
        ListView1.ListViewItemSorter = lvwColumnSorter
        PopulateListView()
    End Sub
    Private Sub PopulateListView()
        Dim Labels As Dictionary(Of Integer, LabelClass) = Form1.SQL.GetLabels()
        ListView1.Items.Clear()
        For Each label As String In Labels.Keys
            Dim item As ListViewItem = New ListViewItem(Labels(label).Name)
            Dim subItems As ListViewItem.ListViewSubItem() = New ListViewItem.ListViewSubItem() {New ListViewItem.ListViewSubItem(item, Labels(label).Spindle),
                                                                                                 New ListViewItem.ListViewSubItem(item, Labels(label).LastChecked.ToString("G")),
                                                                                                 New ListViewItem.ListViewSubItem(item, Labels(label).CheckOutcome)}
            item.SubItems.AddRange(subItems)
            ListView1.Items.Add(item)
        Next
        ListView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim Label As ListView.SelectedIndexCollection = ListView1.SelectedIndices
        If ListView1.SelectedIndices.Count > 0 Then
            Dim Spindle As String = InputBox("Enter the Spindle Name/Number for the selected label", , ListView1.Items(Label(0)).SubItems(1).Text)
            If Not String.IsNullOrWhiteSpace(Spindle) Then
                Form1.SQL.UpdateSpindle(ListView1.Items(Label(0)).Text, Spindle)
                ListView1.Items(Label(0)).SubItems(1).Text = Spindle
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Label As ListView.SelectedIndexCollection = ListView1.SelectedIndices
        If ListView1.SelectedIndices.Count > 0 Then
            Dim New_Label As String = InputBox("Enter the new label name", , ListView1.Items(Label(0)).Text)
            If Not String.IsNullOrWhiteSpace(New_Label) Then
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
        If Not String.IsNullOrWhiteSpace(DriveLetterToUse) Then
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
                        Dim NewPath As String = Path(0).Replace(Path(0), DriveLetterToUse.Chars(0).ToString().ToUpper()) + ":" + Path(1)
                        ChecksumString = ChecksumString + Files(file).Checksum + " *" + NewPath + Environment.NewLine
                    Next
                    My.Computer.FileSystem.WriteAllText(SaveDialog.FileName, ChecksumString, False, New Text.UTF8Encoding(False))
                    MsgBox("Label Content Checksums saved.")
                End If
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        MsgBox("Not implemented")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim DriveLetterToUse = InputBox("Enter the Drive Letter to use to check the checksums")
        If Not String.IsNullOrWhiteSpace(DriveLetterToUse) Then
            Dim Label As ListView.SelectedIndexCollection = ListView1.SelectedIndices
            Dim CC As New ChecksumChecker
            If ListView1.SelectedIndices.Count > 0 Then
                CC.CurrentLabel = ListView1.Items(Label(0)).Text
                Dim Files As Dictionary(Of Integer, FileClass) = Form1.SQL.GetLabelContents(CC.CurrentLabel)
                Dim ChecksumString As String = ""
                For Each file As Integer In Files.Keys
                    Dim Path As String() = Files(file).OriginalPath.Split(":")
                    Dim NewPath As String = Path(0).Replace(Path(0), DriveLetterToUse.Chars(0).ToString().ToUpper()) + ":" + Path(1)
                    ChecksumString = ChecksumString + Files(file).Checksum + " *" + NewPath + Environment.NewLine
                    Dim item As ListViewItem = New ListViewItem(NewPath)
                    item.Tag = file
                    Dim subItems As ListViewItem.ListViewSubItem() = New ListViewItem.ListViewSubItem() {New ListViewItem.ListViewSubItem(item, String.Format("{0:N2} KB", Files(file).FileSize)),
                                                                                                         New ListViewItem.ListViewSubItem(item, ""),
                                                                                                         New ListViewItem.ListViewSubItem(item, Files(file).Checksum),
                                                                                                         New ListViewItem.ListViewSubItem(item, "")}
                    item.SubItems.AddRange(subItems)
                    CC.ListView1.Items.Add(item)
                Next
                CC.Show()
            End If
        End If
    End Sub

    Private Sub ListView1_KeyUp(sender As Object, e As KeyEventArgs) Handles ListView1.KeyUp
        If e.KeyCode = Keys.F5 Then
            PopulateListView()
        End If
    End Sub

    Private Sub ListView1_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles ListView1.ColumnClick
        If e.Column = lvwColumnSorter.SortColumn Then
            If lvwColumnSorter.Order = SortOrder.Ascending Then
                lvwColumnSorter.Order = SortOrder.Descending
            Else
                lvwColumnSorter.Order = SortOrder.Ascending
            End If
        Else
            lvwColumnSorter.SortColumn = e.Column
            lvwColumnSorter.Order = SortOrder.Ascending
        End If
        ListView1.Sort()
    End Sub
End Class