Imports System.IO
Imports System.Security.Cryptography

Public Class Form1
    Public SQL As New SQLClass()
    Private CurrentFiles As Dictionary(Of Integer, FileClass)
    Private SearchLabel As String = ""
    Private lvwColumnSorter As ListViewColumnSorter

    Private Sub GetFolders(ByVal parent As Integer, ByVal nodeToAddTo As TreeNode)
        Dim folders As Dictionary(Of String, Integer) = SQL.GetFolders(parent)
        For Each folder As String In folders.Keys
            If Not nodeToAddTo.Nodes.ContainsKey(folder) Then
                Dim aNode As New TreeNode(folder) With {.Name = folder, .Tag = folders(folder)}
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
                rootNode = New TreeNode(key) With {.Name = rootfolder(key), .Tag = rootfolder(key)}
                GetFolders(rootfolder(key), rootNode)
                TreeView1.Nodes.Add(rootNode)
            Next
        End If
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Initialize()
    End Sub

    Private Sub TreeView1_NodeMouseClick(ByVal sender As Object, ByVal e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick
        GetFiles(e.Node)
    End Sub

    Private Sub GetFiles(node As TreeNode)
        If node.Tag IsNot "Search" Then
            CurrentFiles = SQL.GetFiles(node.Tag)
            PopulateListView()
        End If
    End Sub

    Private Sub PopulateListView()
        ListView1.Items.Clear()
        For Each file As Integer In CurrentFiles.Keys
            If Not SearchLabel = "" And FilterSearchResultsToolStripMenuItem.Checked Then
                If Not CurrentFiles(file).VolumeLabel = SearchLabel Then
                    Continue For
                End If
            End If
            Dim item As ListViewItem = New ListViewItem(CurrentFiles(file).Name)
            Dim subItems As ListViewItem.ListViewSubItem() = New ListViewItem.ListViewSubItem() {New ListViewItem.ListViewSubItem(item, CurrentFiles(file).Type), New ListViewItem.ListViewSubItem(item, CurrentFiles(file).ModifiedDate),
               New ListViewItem.ListViewSubItem(item, String.Format("{0:N2} KB", CurrentFiles(file).FileSize)), New ListViewItem.ListViewSubItem(item, CurrentFiles(file).VolumeLabel), New ListViewItem.ListViewSubItem(item, CurrentFiles(file).Checksum),
               New ListViewItem.ListViewSubItem(item, CurrentFiles(file).OriginalPath), New ListViewItem.ListViewSubItem(item, CurrentFiles(file).Comment), New ListViewItem.ListViewSubItem(item, CurrentFiles(file).Spindle)}
            item.SubItems.AddRange(subItems)
            ListView1.Items.Add(item).Tag = file
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
        Dim thread As New Threading.Thread(Sub() AddFilesAndFolders(CType(e.Data.GetData(DataFormats.FileDrop), String()), node))
        thread.Start()

    End Sub
    Private Sub AddFile(path As String, parent As Integer, VolumeLabel As String, files As Dictionary(Of String, FileClass))
        Dim Filename As String = IO.Path.GetFileName(path)
        If Not files.Keys.Contains(IO.Path.GetFileName(path)) Then
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
            SQL.InsertFile(Filename, parent, VolumeLabel, MD5HashToString, IO.Path.GetExtension(path), IO.File.GetLastWriteTimeUtc(path), My.Computer.FileSystem.GetFileInfo(path).Length / 1024, path)
        End If
    End Sub
    Private Sub RefreshListAfterAddingFiles()
        If RefreshFileListAfterOperationToolStripMenuItem.Checked Then
            If TreeView1.InvokeRequired Then
                TreeView1.BeginInvoke(Sub() GetFiles(TreeView1.SelectedNode))
            Else
                GetFiles(TreeView1.SelectedNode)
            End If
        End If
    End Sub
    Private Sub GetDirectoriesAndFiles(ByVal BaseFolder As DirectoryInfo, parent As Integer, VolumeLabel As String, node As TreeNode, files As Dictionary(Of String, FileClass))
        For Each FI As FileInfo In BaseFolder.GetFiles()
            AddFile(FI.FullName, parent, VolumeLabel, files)
            RefreshListAfterAddingFiles()
        Next
        For Each subF As DirectoryInfo In BaseFolder.GetDirectories()
            Dim CreatedFolderId As Integer = CreateFolder(parent, subF.Name)
            GetFolders(node.Tag, node)
            GetDirectoriesAndFiles(subF, CreatedFolderId, VolumeLabel, node, SQL.GetFilesNameAsKey(CreatedFolderId))
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
            Dim files As Dictionary(Of String, FileClass) = SQL.GetFilesNameAsKey(node.Tag)
            If File.Exists(path) Then
                AddFile(path, node.Tag, VolumeLabel, files)
                RefreshListAfterAddingFiles()
            Else
                Dim CreatedFolderId As Integer = CreateFolder(node.Tag, IO.Path.GetFileName(path))
                GetFolders(node.Tag, node)
                GetDirectoriesAndFiles(New DirectoryInfo(path), CreatedFolderId, VolumeLabel, node, SQL.GetFilesNameAsKey(CreatedFolderId))
            End If
        Next
        GetFolders(node.Tag, node)
        RefreshListAfterAddingFiles()
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
                    SQL.DeleteFile(ListView1.Items(File).Tag)
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
            Clipboard.SetText(CurrentFiles(ListView1.Items(ListView1.SelectedIndices(0)).Tag).VolumeLabel)
        End If
    End Sub

    Private Sub CopyLabelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyLabelToolStripMenuItem.Click
        If TreeView1.SelectedNode IsNot Nothing Then
            Clipboard.SetText(TreeView1.SelectedNode.Name)
        End If
    End Sub
    Private Sub AddNodesToSearchNode(FileParentList As List(Of Integer), nodeToAddTo As TreeNode)
        Dim Folders As New Dictionary(Of String, Integer)
        For Each Parent As Integer In FileParentList
            If Parent > 0 Then
                Dim ParentName As String = SQL.GetFolderName(Parent)
                If Not Folders.Keys().Contains(ParentName) Then
                    Folders.Add(ParentName, Parent)
                End If
            End If
        Next
        For Each FolderName As String In Folders.Keys
            If Not nodeToAddTo.Nodes.ContainsKey(FolderName) Then
                Dim aNode As New TreeNode(FolderName) With {.Name = FolderName, .Tag = Folders(FolderName)}
                If TreeView1.InvokeRequired Then
                    TreeView1.BeginInvoke(Sub() nodeToAddTo.Nodes.Add(aNode))
                Else
                    nodeToAddTo.Nodes.Add(aNode)
                    nodeToAddTo = aNode
                End If
            Else
                nodeToAddTo = nodeToAddTo.Nodes(FolderName)
            End If
        Next
    End Sub

    Private Sub SearchToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SearchToolStripMenuItem.Click
        Dim SearchString As String = InputBox("Search for?")
        If Not String.IsNullOrEmpty(SearchString) Then
            SearchFunction(SQL.SearchFiles(SearchString))
        End If
    End Sub

    Public Sub SearchFunction(SearchResults As Dictionary(Of Integer, FileClass), Optional Label As String = "")
        SearchLabel = Label
        CurrentFiles = SearchResults
        Dim FileParents As New Dictionary(Of Integer, List(Of Integer))
        For Each File In CurrentFiles
            Dim Parent As Integer = SQL.GetFileParent(File.Key)
            Dim FileParentList As New List(Of Integer) From {Parent}
            While Parent > 0
                Parent = SQL.GetFolderParent(Parent)
                FileParentList.Add(Parent)
            End While
            FileParents.Add(File.Key, FileParentList)
        Next
        If TreeView1.Nodes.ContainsKey("Search") Then
            TreeView1.Nodes.RemoveByKey("Search")
        End If
        Dim SearchNode As New TreeNode("Search Results") With {.Name = "Search", .Tag = "Search"}
        TreeView1.Nodes.Add(SearchNode)
        For Each File In CurrentFiles
            FileParents(File.Key).Reverse()
            AddNodesToSearchNode(FileParents(File.Key), SearchNode)
        Next
        If Not OnlyShowSearchTreeResultsToolStripMenuItem.Checked Then PopulateListView()
    End Sub

    Private Sub EditCommentToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EditCommentToolStripMenuItem.Click
        Dim File As ListView.SelectedIndexCollection = ListView1.SelectedIndices
        If ListView1.SelectedIndices.Count > 0 Then
            Dim Comment As String = InputBox("Enter a comment for the selected file", , CurrentFiles(ListView1.Items(File(0)).Tag).Comment)
            If Not String.IsNullOrEmpty(Comment) Then
                SQL.UpdateFileComment(ListView1.Items(File(0)).Tag, Comment)
                GetFiles(TreeView1.SelectedNode)
            End If
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

    Private Sub LabelManagementToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LabelManagementToolStripMenuItem.Click
        LabelManagement.Show()
    End Sub

    Private Sub SaveChecksumsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveChecksumsToolStripMenuItem.Click
        Dim Files As ListView.SelectedIndexCollection = ListView1.SelectedIndices
        If ListView1.SelectedIndices.Count > 0 Then
            Dim SaveDialog As New SaveFileDialog With {.FileName = "Checksums.md5", .Filter = "MD5 Checksum|*.md5"}
            Dim result As MsgBoxResult = SaveDialog.ShowDialog()
            If result = MsgBoxResult.Ok Then
                Dim ChecksumString As String = ""
                For Each file In Files
                    ChecksumString = ChecksumString + CurrentFiles(ListView1.Items(file).tag).Checksum + " *" + CurrentFiles(ListView1.Items(file).tag).Name + Environment.NewLine
                Next
                My.Computer.FileSystem.WriteAllText(SaveDialog.FileName, ChecksumString, False, New Text.UTF8Encoding(False))
            End If
        End If
    End Sub

End Class
