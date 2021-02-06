Imports System.CodeDom

Public Class FileClass
    Public Name As String = String.Empty
    Public Type As String = String.Empty
    Public ModifiedDate As Date = Nothing
    Public FileSize As Double = 0.0
    Public VolumeLabel As String = String.Empty
    Public Checksum As String = String.Empty
    Public OriginalPath As String = String.Empty
    Public Comment As String = String.Empty
    Public Spindle As String = String.Empty
    Public Sub New(file_name As String, file_type As String, file_mod_date As Date, file_size As Double, file_vol_label As String, file_checksum As String, file_orig_path As String, file_comment As String, file_spindle As String)
        Name = file_name
        Type = file_type
        ModifiedDate = file_mod_date
        FileSize = file_size
        VolumeLabel = file_vol_label
        Checksum = file_checksum
        OriginalPath = file_orig_path
        Comment = file_comment
        Spindle = file_spindle
    End Sub
End Class
