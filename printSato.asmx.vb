Imports System.ComponentModel
Imports System.Web.Services
Imports System.Web.Services.Protocols


Imports System.Runtime.InteropServices
Imports System.IO

Imports Oracle.DataAccess.Client
Imports System.Web.Configuration

' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class WebService1
    Inherits System.Web.Services.WebService
    'Public location_file As String = AppDomain.CurrentDomain.BaseDirectory & "file"
    <WebMethod()>
    Public Function MainProcess(ByVal pallete As String, ByVal printer As String) As String

        Dim con_str As String = WebConfigurationManager.ConnectionStrings.Item("IMConnectionString").ToString

        Dim Conn As New OracleConnection
        Conn.ConnectionString = con_str
        Dim cmd As New OracleCommand
        Conn.Open()

        cmd.Connection = Conn
        cmd.CommandText = "Select to_char(IMFR_UD_PKG_EDP_DATE, 'YYYYMMDD') as dayy , IMFR_SD_PKG_ID as pkg from IMFR_UT_MTLLSL004 where IMFR_SD_PKG_ID = '" & pallete & "' "


        cmd.CommandType = CommandType.Text
        Dim ex_sum = cmd.ExecuteReader()
        Dim tbl_sum As DataTable = New DataTable()
        tbl_sum.Load(ex_sum)
        Conn.Close()

        Dim logistic_t As String = WebConfigurationManager.ConnectionStrings.Item("logistic_t").ToString
        Conn.ConnectionString = logistic_t

        Conn.Open()

        cmd.Connection = Conn
        cmd.CommandText = "select IP_PRINTER as ip from Printer_Db where PRINTER_NAME = '" & printer & "'"

        cmd.CommandType = CommandType.Text
        Dim print_name = cmd.ExecuteReader()
        Dim print_ip As DataTable = New DataTable()
        print_ip.Load(print_name)






        'Dim FILE_NAME As String = (AppDomain.CurrentDomain.BaseDirectory) & ipname.Replace(".", "_") & ".txt"
        'Dim printer_name As String = "\\163.50.57.110\sato_be"
        'Dim location As String = WebConfigurationManager.ConnectionStrings.Item("PATH").ToString


        Dim bin As String
        Dim line As New List(Of String)()
        Dim ipname As String = HttpContext.Current.Request.UserHostAddress

        Dim FILE_NAME As String = (AppDomain.CurrentDomain.BaseDirectory) & "\file\lable.txt"
        'Dim FILE_NAME = location_file & "\lable.txt"

        Dim FormNo As String = "\BoxID.ini"

        Dim reader = File.OpenText((AppDomain.CurrentDomain.BaseDirectory) & FormNo)
        Dim j As Integer = 0
        Dim i As Integer = 0

        While (reader.Peek() <> -1)
            line.Add(reader.ReadLine())
            line(i) = line(i).Replace("#ESC#", ChrW(CInt("&H" & "22")))

            j = j + 1
            i += 1
        End While
        Dim line1 As String() = line.ToArray()

        'For i = 0 To tbl_sum.Rows.Count - 1

        Dim Result As String = label(tbl_sum.Rows(0)("pkg"), tbl_sum.Rows(0)("dayy"), line1)
        ' Next


        RawPrinterHelper.SendFileToPrinter(print_ip.Rows(0)("ip"), FILE_NAME)
        System.IO.File.Delete(FILE_NAME)


    End Function
    Function label(ByVal pallete As String, ByVal dayy As String, ByVal ParamArray line() As String) As String
        Dim layout As String
        Dim ipname As String = HttpContext.Current.Request.UserHostAddress

        Dim FILE_NAME As String = (AppDomain.CurrentDomain.BaseDirectory) & "\file\lable.txt"
        'Dim FILE_NAME = location_file & "\lable.txt"

        Dim Chkprinter As String
        Chkprinter = "\\163.50.57.110\sato_be"

        For i = 0 To line.Length - 1
            Select Case i
                Case 0
                    My.Computer.FileSystem.WriteAllText(FILE_NAME, line(i), True, System.Text.Encoding.GetEncoding(1252))
                Case 1
                    layout = line(i).Replace("#DATA1#", "Part : ")
                    My.Computer.FileSystem.WriteAllText(FILE_NAME, layout, True, System.Text.Encoding.GetEncoding(1252))
                Case 2

                    layout = line(i).Replace("#DATA2#", "Insp .  N0# .")
                    My.Computer.FileSystem.WriteAllText(FILE_NAME, layout, True, System.Text.Encoding.GetEncoding(1252))
                Case 3
                    layout = line(i).Replace("#DATA3#", "PACKAGE ID :")
                    My.Computer.FileSystem.WriteAllText(FILE_NAME, layout, True, System.Text.Encoding.GetEncoding(1252))
                Case 4
                    layout = line(i).Replace("#DATA4#", pallete)
                    My.Computer.FileSystem.WriteAllText(FILE_NAME, layout, True, System.Text.Encoding.GetEncoding(1252))
                Case 5
                    layout = line(i).Replace("#DATA5#", pallete)
                    My.Computer.FileSystem.WriteAllText(FILE_NAME, layout, True, System.Text.Encoding.GetEncoding(1252))
                Case 6
                    layout = line(i).Replace("#DATA6#", "Registered Date: " & dayy)
                    My.Computer.FileSystem.WriteAllText(FILE_NAME, layout, True, System.Text.Encoding.GetEncoding(1252))
                Case 7
                    My.Computer.FileSystem.WriteAllText(FILE_NAME, line(i), True, System.Text.Encoding.GetEncoding(1252))
            End Select
        Next
        'objWriter.Close()
        Return layout
    End Function

    Function addBlank(ByVal Len As Integer, ByVal Text As String) As String
        Dim TempLen As Integer, TempSpec As Integer, cntSpec As Integer
        TempLen = Text.Length
        TempSpec = Len - TempLen
        Dim blankValue As String
        blankValue = ""
        For cntSpec = 1 To TempSpec
            blankValue = blankValue + " "
        Next cntSpec
        addBlank = Text & blankValue
    End Function

    Private Class FromUriAttribute
        Inherits Attribute
    End Class
End Class

Public Class RawPrinterHelper
    ' Structure and API declarions:
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Structure DOCINFOW
        <MarshalAs(UnmanagedType.LPWStr)> Public pDocName As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pOutputFile As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pDataType As String
    End Structure

    <DllImport("winspool.Drv", EntryPoint:="OpenPrinterW",
       SetLastError:=True, CharSet:=CharSet.Unicode,
       ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
    Public Shared Function OpenPrinter(ByVal src As String, ByRef hPrinter As IntPtr, ByVal pd As Integer) As Boolean
    End Function
    <DllImport("winspool.Drv", EntryPoint:="ClosePrinter",
       SetLastError:=True, CharSet:=CharSet.Unicode,
       ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
    Public Shared Function ClosePrinter(ByVal hPrinter As IntPtr) As Boolean
    End Function
    <DllImport("winspool.Drv", EntryPoint:="StartDocPrinterW",
       SetLastError:=True, CharSet:=CharSet.Unicode,
       ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
    Public Shared Function StartDocPrinter(ByVal hPrinter As IntPtr, ByVal level As Int32, ByRef pDI As DOCINFOW) As Boolean
    End Function
    <DllImport("winspool.Drv", EntryPoint:="EndDocPrinter",
       SetLastError:=True, CharSet:=CharSet.Unicode,
       ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
    Public Shared Function EndDocPrinter(ByVal hPrinter As IntPtr) As Boolean
    End Function
    <DllImport("winspool.Drv", EntryPoint:="StartPagePrinter",
       SetLastError:=True, CharSet:=CharSet.Unicode,
       ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
    Public Shared Function StartPagePrinter(ByVal hPrinter As IntPtr) As Boolean
    End Function
    <DllImport("winspool.Drv", EntryPoint:="EndPagePrinter",
       SetLastError:=True, CharSet:=CharSet.Unicode,
       ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
    Public Shared Function EndPagePrinter(ByVal hPrinter As IntPtr) As Boolean
    End Function
    <DllImport("winspool.Drv", EntryPoint:="WritePrinter",
       SetLastError:=True, CharSet:=CharSet.Unicode,
       ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
    Public Shared Function WritePrinter(ByVal hPrinter As IntPtr, ByVal pBytes As IntPtr, ByVal dwCount As Int32, ByRef dwWritten As Int32) As Boolean
    End Function

    ' SendBytesToPrinter()
    ' When the function is given a printer name and an unmanaged array of  
    ' bytes, the function sends those bytes to the print queue.
    ' Returns True on success or False on failure.
    Public Shared Function SendBytesToPrinter(ByVal szPrinterName As String, ByVal pBytes As IntPtr, ByVal dwCount As Int32) As Boolean
        Dim hPrinter As IntPtr      ' The printer handle.
        Dim dwError As Int32        ' Last error - in case there was trouble.
        Dim di As DOCINFOW          ' Describes your document (name, port, data type).
        Dim dwWritten As Int32      ' The number of bytes written by WritePrinter().
        Dim bSuccess As Boolean     ' Your success code.

        ' Set up the DOCINFO structure.
        With di
            .pDocName = "My Visual Basic .NET RAW Document"
            .pDataType = "RAW"
        End With
        ' Assume failure unless you specifically succeed.
        bSuccess = False
        If OpenPrinter(szPrinterName, hPrinter, 0) Then
            If StartDocPrinter(hPrinter, 1, di) Then
                If StartPagePrinter(hPrinter) Then
                    ' Write your printer-specific bytes to the printer.
                    bSuccess = WritePrinter(hPrinter, pBytes, dwCount, dwWritten)
                    EndPagePrinter(hPrinter)
                End If
                EndDocPrinter(hPrinter)
            End If
            ClosePrinter(hPrinter)
        End If
        ' If you did not succeed, GetLastError may give more information
        ' about why not.
        If bSuccess = False Then
            dwError = Marshal.GetLastWin32Error()
        End If
        Return bSuccess
    End Function ' SendBytesToPrinter()

    ' SendFileToPrinter()
    ' When the function is given a file name and a printer name, 
    ' the function reads the contents of the file and sends the
    ' contents to the printer.
    ' Presumes that the file contains printer-ready data.
    ' Shows how to use the SendBytesToPrinter function.
    ' Returns True on success or False on failure.
    Public Shared Function SendFileToPrinter(ByVal szPrinterName As String, ByVal szFileName As String) As Boolean
        ' Open the file.
        Dim fs As New FileStream(szFileName, FileMode.Open)
        ' Create a BinaryReader on the file.
        Dim br As New BinaryReader(fs)
        ' Dim an array of bytes large enough to hold the file's contents.
        Dim bytes(fs.Length) As Byte
        Dim bSuccess As Boolean
        ' Your unmanaged pointer.
        Dim pUnmanagedBytes As IntPtr

        ' Read the contents of the file into the array.
        bytes = br.ReadBytes(fs.Length)
        ' Allocate some unmanaged memory for those bytes.
        pUnmanagedBytes = Marshal.AllocCoTaskMem(fs.Length)
        ' Copy the managed byte array into the unmanaged array.
        Marshal.Copy(bytes, 0, pUnmanagedBytes, fs.Length)
        ' Send the unmanaged bytes to the printer.
        bSuccess = SendBytesToPrinter(szPrinterName, pUnmanagedBytes, fs.Length)
        ' Free the unmanaged memory that you allocated earlier.
        Marshal.FreeCoTaskMem(pUnmanagedBytes)
        fs.Close()
        Return bSuccess
    End Function ' SendFileToPrinter()

    ' When the function is given a string and a printer name,
    ' the function sends the string to the printer as raw bytes.
    Public Shared Function SendStringToPrinter(ByVal szPrinterName As String, ByVal szString As String) As Boolean
        Dim pBytes As IntPtr
        Dim dwCount As Int32
        ' How many characters are in the string?
        dwCount = szString.Length()
        ' Assume that the printer is expecting ANSI text, and then convert
        ' the string to ANSI text.
        pBytes = Marshal.StringToCoTaskMemAnsi(szString)
        ' Send the converted ANSI string to the printer.
        SendBytesToPrinter(szPrinterName, pBytes, dwCount)
        Marshal.FreeCoTaskMem(pBytes)
    End Function
End Class