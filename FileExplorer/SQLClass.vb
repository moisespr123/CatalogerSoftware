﻿Imports System.IO
Imports MySql.Data.MySqlClient
Public Class SQLClass
    Private MySQLString As String = ""
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
        Dim SQLQuery As String = "SELECT id, name, vol_label, checksum, type, file_size, mod_date, orig_path FROM files WHERE parent=@parent ORDER BY name"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@parent", parent)
        Connection.Open()
        Dim reader As MySqlDataReader = Command.ExecuteReader
        Dim files As New Dictionary(Of Integer, FileClass)
        If reader.HasRows Then
            While reader.Read
                files.Add(reader("id"), New FileClass(reader("name"), reader("type"), reader("mod_date"), reader("file_size"), reader("vol_label"), reader("checksum"), reader("orig_path")))
            End While
        End If
        Connection.Close()
        Return files
    End Function
    Public Function GetFilesNameAsKey(parent As Integer) As Dictionary(Of String, FileClass)
        Dim SQLQuery As String = "SELECT id, name, vol_label, checksum, type, file_size, mod_date, orig_path FROM files WHERE parent=@parent ORDER BY name"
        Dim Connection As MySqlConnection = New MySqlConnection(MySQLString)
        Dim Command As New MySqlCommand(SQLQuery, Connection)
        Command.Parameters.AddWithValue("@parent", parent)
        Connection.Open()
        Dim reader As MySqlDataReader = Command.ExecuteReader
        Dim files As New Dictionary(Of String, FileClass)
        If reader.HasRows Then
            While reader.Read
                files.Add(reader("name"), New FileClass(reader("name"), reader("type"), reader("mod_date"), reader("file_size"), reader("vol_label"), reader("checksum"), reader("orig_path")))
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

End Class
