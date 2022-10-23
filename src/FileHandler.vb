Imports System.Windows.Forms
Imports Extensibility
Imports Microsoft.Vbe.Interop
Imports Microsoft.Office.Core
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

Module FileHandler

    Public Function getFileVersion(fileInfo As FileInfo) As String

        Dim myShell As Shell32.Shell = New Shell32.Shell
        Dim myFolder As Shell32.Folder = myShell.NameSpace(fileInfo.Directory.FullName)
        Dim myFile As Shell32.FolderItem = myFolder.Items.Item(fileInfo.Name)

        Dim pattern As String = "[v|V]\d+[.\d+]+", details As String = ""
        For i As Integer = 0 To 500

            Dim fileDetail As String = myFolder.GetDetailsOf(myFile, i)

            If Regex.IsMatch(fileDetail, pattern) Then
                Return fileDetail
            End If

        Next

    End Function

    Public Function getFileVersion(filePath As String) As String

        Return FileVersionInfo.GetVersionInfo(filePath).FileVersion

    End Function

    Public Sub convertFileToUtf(fileName As String)

        Dim sReader As New StreamReader(fileName, Encoding.Unicode)
        Dim fileContent() As Byte = Encoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Unicode.GetBytes(sReader.ReadToEnd))
        sReader.Close()

        Dim sWriter As New StreamWriter(fileName)
        sWriter.Write(Encoding.UTF8.GetString(fileContent))
        sWriter.Close()

    End Sub

    Public Sub convertFileToUnicode(fileInfo As FileInfo)

        Call convertFileToUnicode(fileInfo.FullName)

    End Sub

    Public Sub convertFileToUnicode(fileName As String)

        Dim sReader As New StreamReader(fileName, Encoding.UTF8)
        Dim fileContent() As Byte = Encoding.Convert(Encoding.UTF8, Encoding.Default, Encoding.UTF8.GetBytes(sReader.ReadToEnd))
        sReader.Close()

        Dim sWriter As New StreamWriter(fileName, False, Encoding.Default)
        sWriter.Write(Encoding.Default.GetString(fileContent))
        sWriter.Close()

    End Sub

    Public Function getFileEncoding(filePath As String) As Encoding

        Using sr As New StreamReader(path:=filePath, detectEncodingFromByteOrderMarks:=True)
            sr.Read()
            Return sr.CurrentEncoding
        End Using

    End Function

    <DebuggerHidden>
    Public Function getVBAFolder(Optional create As Boolean = True) As DirectoryInfo

        Dim vbaPath As String = Environ("AppData") & "\VBA"

        If Directory.Exists(vbaPath) Then
            Return New DirectoryInfo(vbaPath)
        ElseIf create Then
            Return Directory.CreateDirectory(vbaPath)
        Else
            Throw New VBAFolderNotFoundException()
        End If

    End Function

    Public Function getVBProjectSourceFolder(vbProj As VBProject, Optional create As Boolean = True) As DirectoryInfo

        Dim vbProjFileName As String

        Try
            vbProjFileName = Path.Combine(New FileInfo(vbProj.FileName).Directory.FullName, "src")
        Catch ex As Exception
            Throw New VBProjectNotFoundException(vbProj)
        End Try

        If Directory.Exists(vbProjFileName) Then
            Return New DirectoryInfo(vbProjFileName)

        ElseIf create Then
            Return Directory.CreateDirectory(vbProjFileName)

        Else
            Throw New SourceFolderNotExistsException(vbProj)

        End If

    End Function

    <DebuggerHidden>
    Public Function getRequirementTextFile(vbProj As VBProject, Optional create As Boolean = True) As StreamWriter

        Dim requirementPath As String = Path.Combine(New FileInfo(vbProj.FileName).Directory.FullName, "requirement.txt")

        If File.Exists(requirementPath) Then
            Return New StreamWriter(requirementPath)
        ElseIf create Then
            Return File.CreateText(requirementPath)
        Else
            Throw New RequirementNotFoundException(vbProj)
        End If

    End Function

    <DebuggerHidden>
    Public Function getDicoFilesFromDirectory(sourceDirectory As DirectoryInfo, Optional keysWithExtensions As Boolean = True) As Dictionary(Of String, FileInfo)

        Return getDicoFilesFromDirectory(sourceDirectory.FullName, keysWithExtensions)

    End Function

    <DebuggerHidden>
    Public Function getDicoFilesFromDirectory(sourceDirectory As String, Optional keysWithExtensions As Boolean = True) As Dictionary(Of String, FileInfo)

        Dim dicoExpFiles As New Dictionary(Of String, FileInfo)
        Dim files As String() = Directory.GetFiles(sourceDirectory, "*.*", SearchOption.AllDirectories)

        For Each fileName As String In files
            Dim fileInfo As FileInfo = New FileInfo(fileName)
            Dim fileKey As String = IIf(keysWithExtensions, fileInfo.Name, Split(fileInfo.Name, ".")(0))
            If Not dicoExpFiles.ContainsKey(fileKey) Then dicoExpFiles.Add(key:=fileKey, value:=fileInfo)
        Next

        Return dicoExpFiles

    End Function

    <DebuggerHidden>
    Public Function getReportFolder(create As Boolean) As DirectoryInfo

        Dim vbaPath As String = Environ("AppData") & "\VBA"

        If Directory.Exists(vbaPath) Then
            Return New DirectoryInfo(vbaPath)
        ElseIf create Then
            Return Directory.CreateDirectory(vbaPath)
        Else
            Throw New ReportFolderNotFoundException()
        End If

    End Function

    <DebuggerHidden>
    Public Sub openMyReportFolder()
        Dim reportFolder As DirectoryInfo = getReportFolder(True)
        Shell("C:\windows\explorer.exe " & reportFolder.FullName, vbNormalFocus)
    End Sub

End Module
