Imports System.IO
Imports MySql.Data.MySqlClient
Public Class SQLClass
    Private ReadOnly MySQLString As String = ""
    Public Sub New()
        If File.Exists("Config.txt") Then
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
        Dim SQLQuery As String = "SELECT files.id, files.name as file_name, labels.name as vol_label, checksum, type, file_size, mod_date, orig_path, comment, spindle, last_checked FROM files 
                                  INNER JOIN labels ON labels.id = files.vol_label WHERE parent=@parent ORDER BY file_name"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@parent", parent)
        Connection.Open()
        Dim reader As MySqlDataReader = Command.ExecuteReader
        Dim files As New Dictionary(Of Integer, FileClass)
        If reader.HasRows Then
            While reader.Read
                files.Add(reader("id"), New FileClass(reader("file_name"), reader("type"), reader("mod_date"), reader("file_size"), reader("vol_label"), reader("checksum"), reader("orig_path"), reader("comment").ToString(), reader("spindle").ToString()))
            End While
        End If
        Connection.Close()
        Return files
    End Function
    Public Function GetFilesChecksum(parent As Integer) As Dictionary(Of Integer, FileClass)
        Dim SQLQuery As String = "SELECT id, name, checksum FROM files WHERE parent=@parent ORDER BY name"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@parent", parent)
        Connection.Open()
        Dim reader As MySqlDataReader = Command.ExecuteReader
        Dim files As New Dictionary(Of Integer, FileClass)
        If reader.HasRows Then
            While reader.Read
                files.Add(reader("id"), New FileClass(reader("name"), Nothing, Nothing, Nothing, Nothing, reader("checksum"), Nothing, Nothing, Nothing))
            End While
        End If
        Connection.Close()
        Return files
    End Function
    Public Function GetLabels() As Dictionary(Of Integer, LabelClass)
        Dim SQLQuery As String = "SELECT id, name, spindle, last_checked FROM labels ORDER BY name ASC"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Connection.Open()
        Dim reader As MySqlDataReader = Command.ExecuteReader
        Dim labels As New Dictionary(Of Integer, LabelClass)
        If reader.HasRows Then
            While reader.Read
                Dim LastChecked As Date = Date.MinValue
                If Not IsDBNull(reader("last_checked")) Then
                    LastChecked = reader("last_checked")
                End If
                labels.Add(reader("id"), New LabelClass(reader("name"), reader("spindle").ToString(), LastChecked))
            End While
        End If
        Connection.Close()
        Return labels
    End Function

    Public Function CheckLabelExists(Label As String) As Integer
        Dim SQLQuery As String = "SELECT id FROM labels WHERE name=@name"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@name", Label)
        Connection.Open()
        Dim reader As MySqlDataReader = Command.ExecuteReader
        Dim LabelId As Integer = -1
        If reader.HasRows Then
            reader.Read()
            LabelId = reader("id")
        End If
        Connection.Close()
        Return LabelId
    End Function

    Public Function GetFilesNameAsKey(parent As Integer) As Dictionary(Of String, FileClass)
        Dim SQLQuery As String = "SELECT files.id, files.name as file_name, labels.name as vol_label, checksum, type, file_size, mod_date, orig_path, comment, spindle, last_checked FROM files 
                                  INNER JOIN fileexplorertool.labels ON labels.id = files.vol_label WHERE parent=@parent ORDER BY file_name"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@parent", parent)
        Connection.Open()
        Dim reader As MySqlDataReader = Command.ExecuteReader
        Dim files As New Dictionary(Of String, FileClass)
        If reader.HasRows Then
            While reader.Read
                files.Add(reader("file_name"), New FileClass(reader("file_name"), reader("type"), reader("mod_date"), reader("file_size"), reader("vol_label"), reader("checksum"), reader("orig_path"), reader("comment").ToString(), reader("spindle").ToString()))
            End While
        End If
        Connection.Close()
        Return files
    End Function
    Public Function GetFileNameAsKey(parent As Integer, name As String) As Dictionary(Of String, FileClass)
        Dim SQLQuery As String = "SELECT files.id, files.name as file_name, labels.name as vol_label, checksum, type, file_size, mod_date, orig_path, comment, spindle, last_checked FROM files 
                                  INNER JOIN fileexplorertool.labels ON labels.id = files.vol_label WHERE parent=@parent AND file_name=@name ORDER BY file_name"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@parent", parent)
        Command.Parameters.AddWithValue("@name", name)
        Connection.Open()
        Dim reader As MySqlDataReader = Command.ExecuteReader
        Dim files As New Dictionary(Of String, FileClass)
        If reader.HasRows Then
            While reader.Read
                files.Add(reader("name"), New FileClass(reader("file_name"), reader("type"), reader("mod_date"), reader("file_size"), reader("vol_label"), reader("checksum"), reader("orig_path"), reader("comment").ToString(), reader("spindle").ToString()))
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

    Public Function InsertLabel(name As String) As Integer
        Dim SQLQuery As String = "INSERT INTO labels (name) VALUES (@name)"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@name", name)
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
        Dim SQLQuery As String = "SELECT files.id, files.name as file_name, labels.name as vol_label, checksum, type, file_size, mod_date, orig_path, comment, spindle, last_checked FROM files 
                                  INNER JOIN fileexplorertool.labels ON labels.id = files.vol_label 
                                  WHERE file_name LIKE @searchstring OR vol_label LIKE @searchstring OR orig_path LIKE @searchstring OR comment LIKE @searchstring ORDER BY file_name"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@searchstring", "%" + searchString + "%")
        Connection.Open()
        Dim reader As MySqlDataReader = Command.ExecuteReader
        Dim files As New Dictionary(Of Integer, FileClass)
        If reader.HasRows Then
            While reader.Read
                files.Add(reader("id"), New FileClass(reader("file_name"), reader("type"), reader("mod_date"), reader("file_size"), reader("vol_label"), reader("checksum"), reader("orig_path"), reader("comment").ToString(), reader("spindle").ToString()))
            End While
        End If
        Connection.Close()
        Return files
    End Function

    Public Function GetLabelContents(Label As String) As Dictionary(Of Integer, FileClass)
        Dim SQLQuery As String = "SELECT files.id, files.name as file_name, labels.name as vol_label, checksum, type, file_size, mod_date, orig_path, comment, spindle, last_checked FROM files 
                                  INNER JOIN fileexplorertool.labels ON labels.id = files.vol_label WHERE labels.name=@label"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@label", Label)
        Connection.Open()
        Dim reader As MySqlDataReader = Command.ExecuteReader
        Dim files As New Dictionary(Of Integer, FileClass)
        If reader.HasRows Then
            While reader.Read
                files.Add(reader("id"), New FileClass(reader("file_name"), reader("type"), reader("mod_date"), reader("file_size"), reader("vol_label"), reader("checksum"), reader("orig_path"), reader("comment").ToString(), reader("spindle").ToString()))
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
        Dim SQLQuery As String = "UPDATE labels SET spindle=@spindle WHERE name=@label"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@spindle", spindle)
        Command.Parameters.AddWithValue("@label", label)
        Connection.Open()
        Dim result As Integer = Command.ExecuteNonQuery()
        Connection.Close()
        Return result
    End Function

    Public Function UpdateLastChecked(label As String, NewDate As Date) As Integer
        Dim SQLQuery As String = "UPDATE labels SET last_checked=@checked_date WHERE name=@label"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@checked_date", NewDate)
        Command.Parameters.AddWithValue("@label", label)
        Connection.Open()
        Dim result As Integer = Command.ExecuteNonQuery()
        Connection.Close()
        Return result
    End Function

    Public Function UpdateLabel(label As String, new_label As String) As Integer
        Dim SQLQuery As String = "UPDATE labels SET name=@new_label WHERE name=@label"
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
