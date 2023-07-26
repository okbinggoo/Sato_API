Imports System.IO
Imports System.Text

Public Class fileOI
    'Public Shared Function readConfigFile(ByVal inputPath As String) As Hashtable
    Public Shared Function readConfigFile() As Hashtable
        'Keep data as Key,Value
        Dim path As String = AppDomain.CurrentDomain.BaseDirectory
        Dim configPath As String : configPath = path + "config\properties.conf"
        Dim hashMap As Hashtable = New Hashtable
        Using r As StreamReader = New StreamReader(configPath)
            Dim counter As Integer = 0
            ' Read first line.
            Dim line As String : line = r.ReadLine
            ' Loop over each line in file, While list is Not Nothing.
            Do While (Not line Is Nothing)
                If (line.Contains(";")) Then
                    Dim array(2) As String : array = line.Split(";")
                    hashMap.Add(array(0), array(1))
                End If
                line = r.ReadLine
            Loop
        End Using
        Return hashMap
    End Function
    Private Shared Function wrapValue(ByVal value As String, ByVal group As String, ByVal separator As String) As String
        If value.Contains(separator) Then
            If value.Contains(group) Then
                value = value.Replace(group, group + group)
            End If
            value = group & value & group
        End If
        Return value
    End Function
    Public Shared Function outputFile(ByVal dtable As DataTable, ByVal location As String, ByVal fileName As String, ByVal fileType As String, ByVal separator As String) As Boolean
        Dim result As Boolean = True
        Try
            Dim sb As New StringBuilder()
            'Dim separator As String = ","
            Dim group As String = """"
            Dim newLine As String = Environment.NewLine

            Dim pathFile As String = location & "\" & fileName & "." & fileType

            'For Each column As DataColumn In dtable.Columns
            'sb.Append(wrapValue(column.ColumnName, group, separator) & separator)
            'Next
            ' here you could add the column for the username
            'sb.Append(newLine)

            For Each row As DataRow In dtable.Rows
                For Each col As DataColumn In dtable.Columns
                    sb.Append(wrapValue(row(col).ToString(), group, separator) & separator)
                Next
                ' here you could extract the password for the username
                sb.Append(newLine)
            Next
            Using fs As New StreamWriter(pathFile)
                fs.Write(sb.ToString())
            End Using
        Catch ex As Exception
            Console.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
            result = False
        End Try
        Return result
    End Function
End Class
