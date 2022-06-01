Public Class ChecksumChecker
    Public CurrentLabel As String = ""
    Private Sub Verification(Items, UpdateCheckDate)
        Dim Counter As Integer = 0
        Dim GoodCounter As Integer = 0
        For Each file As ListViewItem In Items
            Try
                Dim hash As String = Form1.VerifyChecksum(file.SubItems(0).Text)
                If hash = file.SubItems(3).Text Then
                    ListView1.BeginInvoke(Sub()
                                              file.UseItemStyleForSubItems = False
                                              file.SubItems(2).Text = "MATCH"
                                              file.SubItems(2).BackColor = Color.LimeGreen
                                          End Sub)
                    Form1.SQL.UpdateLastCheckedFile(file.Tag, Date.Now, 1)
                    GoodCounter += 1
                Else
                    ListView1.BeginInvoke(Sub()
                                              file.UseItemStyleForSubItems = False
                                              file.SubItems(2).Text = "MISMATCH"
                                              file.SubItems(2).BackColor = Color.Red
                                          End Sub)
                    Form1.SQL.UpdateLastCheckedFile(file.Tag, Date.Now, 2)
                End If
                ListView1.BeginInvoke(Sub()
                                          file.SubItems(4).Text = hash
                                      End Sub)

            Catch ex As Exception
                ListView1.BeginInvoke(Sub()
                                          file.UseItemStyleForSubItems = False
                                          file.SubItems(2).Text = "ERROR"
                                          file.SubItems(2).BackColor = Color.Tomato
                                      End Sub)
                Form1.SQL.UpdateLastCheckedFile(file.Tag, Date.Now, 2)
            End Try
            Counter += 1
            Label1.BeginInvoke(Sub()
                                   Label1.Text = "Progress: " + Counter.ToString() + " of " + ListView1.Items.Count.ToString()
                               End Sub)
        Next
        If UpdateCheckDate Then
            Dim Outcome As Integer = 2
            If GoodCounter = Counter Then
                Outcome = 1
            End If
            Form1.SQL.UpdateLastCheckedLabel(CurrentLabel, Date.Now, Outcome)
        End If
            Button1.BeginInvoke(Sub() Button1.Enabled = True)
        MsgBox("Done")

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Items As ListViewItem() = New ListViewItem(ListView1.Items.Count - 1) {}
        ListView1.Items.CopyTo(Items, 0)
        Button1.Enabled = False
        Dim thread As New Threading.Thread(Sub() Verification(Items, CheckBox1.Checked))
        thread.Start()

    End Sub

    Private Sub ChecksumChecker_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Text = "Checksum Verification for label: " + CurrentLabel
        Label1.Text = "Progress: 0 of " + ListView1.Items.Count.ToString()
        ListView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
    End Sub
End Class