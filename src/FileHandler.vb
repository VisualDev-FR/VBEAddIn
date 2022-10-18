Imports System.Windows.Forms
Imports Extensibility
Imports Microsoft.Vbe.Interop
Imports Microsoft.Office.Core
Imports System.IO
Imports System.Text

Module FileHandler

    Public Sub convertFileToUtf(fileName As String)

        Dim sReader As New StreamReader(fileName, Encoding.Unicode)
        Dim fileContent() As Byte = Encoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Unicode.GetBytes(sReader.ReadToEnd))
        sReader.Close()

        Dim sWriter As New StreamWriter(fileName)
        sWriter.Write(Encoding.UTF8.GetString(fileContent))
        sWriter.Close()

    End Sub

    Public Sub convertFileToUnicode(fileName As String)

        Dim sReader As New StreamReader(fileName, Encoding.Unicode)
        Dim fileContent() As Byte = Encoding.Convert(Encoding.UTF8, Encoding.Unicode, Encoding.UTF8.GetBytes(sReader.ReadToEnd))
        sReader.Close()

        Dim sWriter As New StreamWriter(fileName)
        sWriter.Write(Encoding.Default.GetString(fileContent))
        sWriter.Close()

    End Sub

    Public Function getFileEncoding(filePath As String) As Encoding

        Using sr As New StreamReader(path:=filePath, detectEncodingFromByteOrderMarks:=True)
            sr.Read()
            Return sr.CurrentEncoding
        End Using

    End Function

End Module
