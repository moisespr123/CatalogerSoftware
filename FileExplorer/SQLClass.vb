Imports System.IO
Imports MySql.Data.MySqlClient
Public Class SQLClass
    Private ReadOnly MySQLString As String = ""
    Public Sub New()
        If IO.File.Exists("Config.txt") Then
            Dim ConfigFile As StreamReader = New StreamReader("Config.txt")
            Dim currentline As String
            Dim MySQLServer As String = String.Empty
            Dim MySQLUser As String = String.Empty
            Dim MySQLPassword As String = String.Empty
            Dim MySQLDatabase As String = String.Empty
            Dim Ssl As String = String.Empty
            While ConfigFile.EndOfStream = False
                currentline = ConfigFile.ReadLine
                If currentline.Contains("mysql-server") Then
                    Dim GetServer As String() = currentline.Split("=")
                    MySQLServer = GetServer(1)
                ElseIf currentline.Contains("mysql-username") Then
                    Dim GetUsername As String() = currentline.Split("=")
                    MySQLUser = GetUsername(1)
                ElseIf currentline.Contains("mysql-password") Then
                    Dim GetPassword As String() = currentline.Split("=")
                    MySQLPassword = GetPassword(1)
                ElseIf currentline.Contains("mysql-database") Then
                    Dim GetDatabase As String() = currentline.Split("=")
                    MySQLDatabase = GetDatabase(1)
                ElseIf currentline.Contains("mysql-sslmode") Then
                    Dim GetSSLMode As String() = currentline.Split("=")
                    Ssl = GetSSLMode(1)
                End If
            End While
            MySQLString = "server=" + MySQLServer + ";user=" + MySQLUser + ";database=" + MySQLDatabase + ";port=3306;password=" + MySQLPassword + ";sslmode= " + Ssl
        Else
            MsgBox("Config.txt file not found. The software will now exit")
            Form1.Close()
        End If
    End Sub
    Public Function GetFolders(parent As Integer) As Dictionary(Of String, Integer)
        Dim SQLQuery As String = "SELECT id, name FROM folders WHERE parent=@parent ORDER BY name"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@parent", parent)
        Connection.Open()
        Dim reader As MySqlDataReader = Command.ExecuteReader
        Dim folders As New Dictionary(Of String, Integer)
        If reader.HasRows Then
            While reader.Read
                folders.Add(reader("name"), reader("id"))
            End While
        End If
        Connection.Close()
        Return folders
    End Function

    Public Function GetFolderName(id As Integer) As String
        Dim SQLQuery As String = "SELECT name FROM folders WHERE id=@id "
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@id", id)
        Connection.Open()
        Dim reader As MySqlDataReader = Command.ExecuteReader
        Dim folderName As String = String.Empty
        If reader.HasRows Then
            While reader.Read
                folderName = reader("name")
            End While
        End If
        Connection.Close()
        Return folderName
    End Function
    Public Function DeleteFiles(parent As Integer) As Integer
        Dim SQLQuery As String = "DELETE FROM files WHERE parent=@parent"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@parent", parent)
        Connection.Open()
        Dim result As Integer = Command.ExecuteNonQuery()
        Connection.Close()
        Return result
    End Function
    Public Function DeleteFolder(id As Integer) As Integer
        Dim SQLQuery As String = "DELETE FROM folders WHERE id=@id"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@id", id)
        Connection.Open()
        Dim result As Integer = Command.ExecuteNonQuery()
        Connection.Close()
        DeleteFiles(id)
        Return result
    End Function
    Public Function DeleteFile(id As Integer) As Integer
        Dim SQLQuery As String = "DELETE FROM files WHERE id=@id"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@id", id)
        Connection.Open()
        Dim result As Integer = Command.ExecuteNonQuery()
        Connection.Close()
        Return result
    End Function
    Public Function RenameFolder(id As Integer, name As String) As Integer
        Dim SQLQuery As String = "UPDATE folders SET name=@name WHERE id=@id"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@id", id)
        Command.Parameters.AddWithValue("@name", name)
        Connection.Open()
        Dim result As Integer = Command.ExecuteNonQuery()
        Connection.Close()
        Return result
    End Function
    Public Function GetFiles(parent As Integer) As Dictionary(Of Integer, FileClass)
        Dim SQLQuery As String = "SELECT id, name, vol_label, checksum, type, file_size, mod_date, orig_path, comment, spindle FROM files WHERE parent=@parent ORDER BY name"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@parent", parent)
        Connection.Open()
        Dim reader As MySqlDataReader = Command.ExecuteReader
        Dim files As New Dictionary(Of Integer, FileClass)
        If reader.HasRows Then
            While reader.Read
                files.Add(reader("id"), New FileClass(reader("name"), reader("type"), reader("mod_date"), reader("file_size"), reader("vol_label"), reader("checksum"), reader("orig_path"), reader("comment").ToString(), reader("spindle").ToString()))
            End While
        End If
        Connection.Close()
        Return files
    End Function
    Public Function GetLabels() As Dictionary(Of String, String)
        Dim SQLQuery As String = "SELECT DISTINCT vol_label, spindle FROM files ORDER BY vol_label ASC"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Connection.Open()
        Dim reader As MySqlDataReader = Command.ExecuteReader
        Dim labels As New Dictionary(Of String, String)
        If reader.HasRows Then
            While reader.Read
                labels.Add(reader("vol_label"), reader("spindle").ToString())
            End While
        End If
        Connection.Close()
        Return labels
    End Function
    Public Function GetFilesNameAsKey(parent As Integer) As Dictionary(Of String, FileClass)
        Dim SQLQuery As String = "SELECT id, name, vol_label, checksum, type, file_size, mod_date, orig_path, comment, spindle FROM files WHERE parent=@parent ORDER BY name"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@parent", parent)
        Connection.Open()
        Dim reader As MySqlDataReader = Command.ExecuteReader
        Dim files As New Dictionary(Of String, FileClass)
        If reader.HasRows Then
            While reader.Read
                files.Add(reader("name"), New FileClass(reader("name"), reader("type"), reader("mod_date"), reader("file_size"), reader("vol_label"), reader("checksum"), reader("orig_path"), reader("comment").ToString(), reader("spindle").ToString()))
            End While
        End If
        Connection.Close()
        Return files
    End Function
    Public Function GetFileNameAsKey(parent As Integer, name As String) As Dictionary(Of String, FileClass)
        Dim SQLQuery As String = "SELECT id, name, vol_label, checksum, type, file_size, mod_date, orig_path, comment, spindle FROM files WHERE parent=@parent AND name=@name ORDER BY name"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@parent", parent)
        Command.Parameters.AddWithValue("@name", name)
        Connection.Open()
        Dim reader As MySqlDataReader = Command.ExecuteReader
        Dim files As New Dictionary(Of String, FileClass)
        If reader.HasRows Then
            While reader.Read
                files.Add(reader("name"), New FileClass(reader("name"), reader("type"), reader("mod_date"), reader("file_size"), reader("vol_label"), reader("checksum"), reader("orig_path"), reader("comment").ToString(), reader("spindle").ToString()))
            End While
        End If
        Connection.Close()
        Return files
    End Function
    Public Function InsertFolder(name As String, parent As Integer) As Integer
        Dim SQLQuery As String = "INSERT INTO folders (name, parent) VALUES (@name, @parent)"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@name", name)
        Command.Parameters.AddWithValue("@parent", parent)
        Connection.Open()
        Command.ExecuteNonQuery()
        Dim id As Integer = Command.LastInsertedId
        Connection.Close()
        Return id
    End Function

    Public Function InsertFile(name As String, parent As Integer, vol_label As String, checksum As String, type As String, mod_date As Date, file_size As Double, orig_path As String) As Integer
        Dim SQLQuery As String = "INSERT INTO files (name, parent, vol_label, checksum, type, mod_date, file_size, orig_path) VALUES (@name, @parent, @vol_label, @checksum, @type, @mod_date, @file_size, @orig_path)"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@name", name)
        Command.Parameters.AddWithValue("@parent", parent)
        Command.Parameters.AddWithValue("@vol_label", vol_label)
        Command.Parameters.AddWithValue("@checksum", checksum)
        Command.Parameters.AddWithValue("@type", type)
        Command.Parameters.AddWithValue("@mod_date", mod_date)
        Command.Parameters.AddWithValue("@file_size", file_size)
        Command.Parameters.AddWithValue("@orig_path", orig_path)
        Connection.Open()
        Command.ExecuteNonQuery()
        Dim id As Integer = Command.LastInsertedId
        Connection.Close()
        Return id
    End Function

    Public Function SearchFiles(searchString As String) As Dictionary(Of Integer, FileClass)
        Dim SQLQuery As String = "SELECT id, name, vol_label, checksum, type, file_size, mod_date, orig_path, comment, spindle FROM files WHERE name LIKE @searchstring OR vol_label LIKE @searchstring OR orig_path LIKE @searchstring OR comment LIKE @searchstring ORDER BY name"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@searchstring", "%" + searchString + "%")
        Connection.Open()
        Dim reader As MySqlDataReader = Command.ExecuteReader
        Dim files As New Dictionary(Of Integer, FileClass)
        If reader.HasRows Then
            While reader.Read
                files.Add(reader("id"), New FileClass(reader("name"), reader("type"), reader("mod_date"), reader("file_size"), reader("vol_label"), reader("checksum"), reader("orig_path"), reader("comment").ToString(), reader("spindle").ToString()))
            End While
        End If
        Connection.Close()
        Return files
    End Function

    Public Function GetLabelContents(Label As String) As Dictionary(Of Integer, FileClass)
        Dim SQLQuery As String = "SELECT id, name, vol_label, checksum, type, file_size, mod_date, orig_path, comment, spindle FROM files WHERE vol_label=@label"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@label", Label)
        Connection.Open()
        Dim reader As MySqlDataReader = Command.ExecuteReader
        Dim files As New Dictionary(Of Integer, FileClass)
        If reader.HasRows Then
            While reader.Read
                files.Add(reader("id"), New FileClass(reader("name"), reader("type"), reader("mod_date"), reader("file_size"), reader("vol_label"), reader("checksum"), reader("orig_path"), reader("comment").ToString(), reader("spindle").ToString()))
            End While
        End If
        Connection.Close()
        Return files
    End Function

    Public Function UpdateFileComment(id As Integer, text As String) As Integer
        Dim SQLQuery As String = "UPDATE files SET comment=@text WHERE id=@id"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@id", id)
        Command.Parameters.AddWithValue("@text", text)
        Connection.Open()
        Dim result As Integer = Command.ExecuteNonQuery()
        Connection.Close()
        Return result
    End Function
    Public Function UpdateSpindle(label As String, spindle As String) As Integer
        Dim SQLQuery As String = "UPDATE files SET spindle=@spindle WHERE vol_label=@label"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@spindle", spindle)
        Command.Parameters.AddWithValue("@label", label)
        Connection.Open()
        Dim result As Integer = Command.ExecuteNonQuery()
        Connection.Close()
        Return result
    End Function

    Public Function UpdateLabel(label As String, new_label As String) As Integer
        Dim SQLQuery As String = "UPDATE files SET vol_label=@new_label WHERE vol_label=@label"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@label", label)
        Command.Parameters.AddWithValue("@new_label", new_label)
        Connection.Open()
        Dim result As Integer = Command.ExecuteNonQuery()
        Connection.Close()
        Return result
    End Function

    Public Function GetFileParent(id As Integer) As Integer
        Dim SQLQuery As String = "SELECT parent FROM files WHERE id=@id"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@id", id)
        Connection.Open()
        Dim reader As MySqlDataReader = Command.ExecuteReader
        Dim parent As Integer = 0
        If reader.HasRows Then
            While reader.Read
                parent = reader("parent")
            End While
        End If
        Connection.Close()
        Return parent
    End Function

    Public Function GetFolderParent(id As Integer) As Integer
        Dim SQLQuery As String = "SELECT parent FROM folders WHERE id=@id"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@id", id)
        Connection.Open()
        Dim reader As MySqlDataReader = Command.ExecuteReader
        Dim parent As Integer = 0
        If reader.HasRows Then
            While reader.Read
                parent = reader("parent")
            End While
        End If
        Connection.Close()
        Return parent
    End Function
End Class
