Imports System.IO
Imports System.Security.Cryptography

Public Class Form1
    Private SQL As New SQLClass()
    Private CurrentFiles As Dictionary(Of Integer, FileClass)
    Private lvwColumnSorter As ListViewColumnSorter

    Private Sub GetFolders(ByVal parent As Integer, ByVal nodeToAddTo As TreeNode)
        Dim aNode As TreeNode
        Dim folders As Dictionary(Of String, Integer) = SQL.GetFolders(parent)
        For Each folder As String In folders.Keys
            If Not nodeToAddTo.Nodes.ContainsKey(folder) Then
                aNode = New TreeNode(folder)
                aNode.Name = folder
                aNode.Tag = folders(folder)
                GetFolders(folders(folder), aNode)
                If TreeView1.InvokeRequired Then
                    TreeView1.BeginInvoke(Sub() nodeToAddTo.Nodes.Add(aNode))
                Else
                    nodeToAddTo.Nodes.Add(aNode)
                End If
            Else
                GetFolders(folders(folder), nodeToAddTo.Nodes(folder))
            End If
        Next
    End Sub
    Private Sub Initialize()
        lvwColumnSorter = New ListViewColumnSorter()
        ListView1.ListViewItemSorter = lvwColumnSorter
        Dim rootfolder As Dictionary(Of String, Integer) = SQL.GetFolders(0)
        If rootfolder.Count = 0 Then
            CreateFolder(0, String.Empty, True)
            Initialize()
        End If
        If rootfolder.Count > 0 Then
            Dim rootNode As TreeNode
            For Each key As String In rootfolder.Keys
                rootNode = New TreeNode(key)
                rootNode.Name = rootfolder(key)
                rootNode.Tag = rootfolder(key)
                GetFolders(rootfolder(key), rootNode)
                TreeView1.Nodes.Add(rootNode)
            Next
        End If
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Initialize()
    End Sub

    Private Sub treeView1_NodeMouseClick(ByVal sender As Object, ByVal e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick
        GetFiles(e.Node)
    End Sub

    Private Sub GetFiles(node As TreeNode)
        CurrentFiles = SQL.GetFiles(node.Tag)
        PopulateListView()
    End Sub

    Private Sub PopulateListView()
        ListView1.Items.Clear()
        For Each file As Integer In CurrentFiles.Keys
            Dim item As ListViewItem = New ListViewItem(CurrentFiles(file).Name)
            Dim subItems As ListViewItem.ListViewSubItem() = New ListViewItem.ListViewSubItem() {New ListViewItem.ListViewSubItem(item, CurrentFiles(file).Type), New ListViewItem.ListViewSubItem(item, CurrentFiles(file).ModifiedDate),
               New ListViewItem.ListViewSubItem(item, String.Format("{0:N2} KB", CurrentFiles(file).FileSize)), New ListViewItem.ListViewSubItem(item, CurrentFiles(file).VolumeLabel), New ListViewItem.ListViewSubItem(item, CurrentFiles(file).Checksum),
               New ListViewItem.ListViewSubItem(item, CurrentFiles(file).OriginalPath), New ListViewItem.ListViewSubItem(item, CurrentFiles(file).Comment)}
            item.SubItems.AddRange(subItems)
            ListView1.Items.Add(item)
        Next
        ListView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
    End Sub
    Private Sub Form1_DragEnter(sender As Object, e As DragEventArgs) Handles MyBase.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub
    Private Sub Form1_DragDrop(sender As Object, e As DragEventArgs) Handles MyBase.DragDrop
        Dim node As TreeNode = TreeView1.SelectedNode
        Dim thread As New System.Threading.Thread(Sub() AddFilesAndFolders(CType(e.Data.GetData(DataFormats.FileDrop), String()), node))
        thread.Start()

    End Sub
    Private Sub AddFile(path As String, parent As Integer, VolumeLabel As String)
        Dim files As Dictionary(Of String, FileClass) = SQL.GetFilesNameAsKey(parent)
        If Not files.Keys.Contains(System.IO.Path.GetFileName(path)) Then
            StatusStrip1.BeginInvoke(Sub()
                                         ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                                         StatusToolStripLabel.Text = "Adding file: " + path
                                     End Sub)
            Dim MD5Hash As MD5 = MD5.Create
            Dim MD5HashToString As String = ""
            Dim file As New FileStream(path, FileMode.Open, FileAccess.Read)
            MD5Hash.ComputeHash(file)
            For Each b In MD5Hash.Hash
                MD5HashToString += b.ToString("x2")
            Next
            file.Close()
            SQL.InsertFile(IO.Path.GetFileName(path), parent, VolumeLabel, MD5HashToString, IO.Path.GetExtension(path), IO.File.GetLastWriteTimeUtc(path), My.Computer.FileSystem.GetFileInfo(path).Length / 1024, path)
        End If
    End Sub
    Private Sub RefreshListAfterAddingFiles(node As TreeNode)
        If TreeView1.InvokeRequired Then
            TreeView1.BeginInvoke(Sub() GetFiles(TreeView1.SelectedNode))
        Else
            GetFiles(TreeView1.SelectedNode)
        End If
    End Sub
    Private Sub GetDirectoriesAndFiles(ByVal BaseFolder As DirectoryInfo, parent As Integer, VolumeLabel As String, node As TreeNode)
        For Each FI As FileInfo In BaseFolder.GetFiles()
            AddFile(FI.FullName, parent, VolumeLabel)
            RefreshListAfterAddingFiles(node)
        Next
        For Each subF As DirectoryInfo In BaseFolder.GetDirectories()
            Dim CreatedFolderId As Integer = CreateFolder(parent, subF.Name)
            GetFolders(node.Tag, node)
            GetDirectoriesAndFiles(subF, CreatedFolderId, VolumeLabel, node)
        Next
    End Sub

    Private Sub AddFilesAndFolders(filepath As String(), node As TreeNode)
        Dim VolumeLabel As String = InputBox("Enter a Volume Label")
        If String.IsNullOrWhiteSpace(VolumeLabel) Then
            Return
        End If
        StatusStrip1.BeginInvoke(Sub()
                                     ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                                     StatusToolStripLabel.Text = "Adding files"
                                 End Sub)
        For Each path In filepath
            If File.Exists(path) Then
                AddFile(path, node.Tag, VolumeLabel)
                RefreshListAfterAddingFiles(node)
            Else
                Dim CreatedFolderId As Integer = CreateFolder(node.Tag, System.IO.Path.GetFileName(path))
                GetFolders(node.Tag, node)
                GetDirectoriesAndFiles(New DirectoryInfo(path), CreatedFolderId, VolumeLabel, node)
            End If
        Next
        GetFolders(node.Tag, node)
        RefreshListAfterAddingFiles(node)
        StatusStrip1.BeginInvoke(Sub()
                                     ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                                     StatusToolStripLabel.Text = "Ready"
                                 End Sub)
    End Sub
    Private Sub NewFolderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewFolderToolStripMenuItem.Click
        CreateFolder(TreeView1.SelectedNode.Tag, String.Empty)
        GetFolders(TreeView1.SelectedNode.Tag, TreeView1.SelectedNode)
    End Sub
    Private Function CreateFolder(parent As Integer, Optional name As String = "", Optional rootFolder As Boolean = False) As Integer
        Dim Folders As Dictionary(Of String, Integer) = SQL.GetFolders(parent)
        If rootFolder Then
            While String.IsNullOrWhiteSpace(name)
                name = InputBox("Enter a name for the root folder")
            End While
        Else
            If String.IsNullOrWhiteSpace(name) Then
                name = InputBox("Enter a name for the folder")
            End If
        End If
        If Folders.Keys.Contains(name) Then
            Return Folders(name)
        Else
            Return SQL.InsertFolder(name, parent)
        End If
    End Function

    Private Function DeleteFolder(node As TreeNode)
        Dim treenodes As New List(Of Integer)
        If node.Nodes.Count > 0 Then
            For Each tree_node As TreeNode In node.Nodes
                treenodes.Add(0)
            Next
            For Each index As Integer In treenodes
                DeleteFolder(node.Nodes(index))
            Next
        End If
        SQL.DeleteFolder(node.Tag)
        node.Remove()
        Return True
    End Function

    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        If MsgBox("Do you really want to delete the selected folder?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            DeleteFolder(TreeView1.SelectedNode)
        End If
        TreeView1.Tag = 0
    End Sub

    Private Sub RenameToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RenameToolStripMenuItem.Click
        Dim NewFolderName As String = InputBox("Enter a name to rename the folder", , TreeView1.SelectedNode.Name)
        If Not String.IsNullOrEmpty(NewFolderName) Then
            SQL.RenameFolder(TreeView1.SelectedNode.Tag, NewFolderName)
            TreeView1.SelectedNode.Text = NewFolderName
            TreeView1.SelectedNode.Name = NewFolderName
        End If
    End Sub

    Private Sub DeleteToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem1.Click
        Dim FilesToDelete As ListView.SelectedIndexCollection = ListView1.SelectedIndices
        If FilesToDelete.Count > 0 Then
            If MsgBox("Do you really want to delete the selected file(s)?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                For Each File As Integer In FilesToDelete
                    SQL.DeleteFile(CurrentFiles.Keys(File))
                Next
            End If
            GetFiles(TreeView1.SelectedNode)
        End If
    End Sub
    Private Sub TreeView_DrawNode(ByVal sender As Object, ByVal e As DrawTreeNodeEventArgs) Handles TreeView1.DrawNode
        If e.Node Is Nothing Then Return
        Dim selected = (e.State And TreeNodeStates.Selected) = TreeNodeStates.Selected
        Dim unfocused = Not e.Node.TreeView.Focused
        If selected AndAlso unfocused Then
            Dim font = If(e.Node.NodeFont, e.Node.TreeView.Font)
            e.Graphics.FillRectangle(SystemBrushes.Highlight, e.Bounds)
            TextRenderer.DrawText(e.Graphics, e.Node.Text, font, e.Bounds, SystemColors.HighlightText, TextFormatFlags.GlyphOverhangPadding)
        Else
            e.DrawDefault = True
        End If
    End Sub

    Private Sub CopyVolumeLabelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyVolumeLabelToolStripMenuItem.Click
        If ListView1.SelectedIndices.Count > 0 Then
            Clipboard.SetText(CurrentFiles(CurrentFiles.Keys(ListView1.SelectedIndices(0))).VolumeLabel)
        End If
    End Sub

    Private Sub CopyLabelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyLabelToolStripMenuItem.Click
        If TreeView1.SelectedNode IsNot Nothing Then
            Clipboard.SetText(TreeView1.SelectedNode.Name)
        End If
    End Sub

    Private Sub SearchToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SearchToolStripMenuItem.Click
        Dim SearchString As String = InputBox("Search for?")
        If Not String.IsNullOrEmpty(SearchString) Then
            CurrentFiles = SQL.SearchFiles(SearchString)
        End If
        PopulateListView()
    End Sub

    Private Sub EditCommentToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EditCommentToolStripMenuItem.Click
        Dim File As ListView.SelectedIndexCollection = ListView1.SelectedIndices
        If ListView1.SelectedIndices.Count > 0 Then
            Dim Comment As String = InputBox("Enter a comment for the selected file", , CurrentFiles(CurrentFiles.Keys(File(0))).Comment)
            If Not String.IsNullOrEmpty(Comment) Then
                SQL.UpdateFileComment(CurrentFiles.Keys(File(0)), Comment)
            End If
            GetFiles(TreeView1.SelectedNode)
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
