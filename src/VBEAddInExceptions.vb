Imports System.Windows.Forms
Imports Extensibility
Imports Microsoft.Vbe.Interop
Imports Microsoft.Office.Core
Imports System.IO

Public Class SourceFolderNotExistsException : Inherits ApplicationException
    Sub New(vbProj As VBProject)
        message = vbProj.Name & " Source folder doesn't exists."
    End Sub
    Public Overrides ReadOnly Property message As String
End Class

Public Class VBProjectNotFoundException : Inherits ApplicationException
    Sub New(vbProj As VBProject)
        message = vbProj.Name & " File not found."
    End Sub
    Public Overrides ReadOnly Property message As String
End Class

Public Class ReportFolderNotFoundException : Inherits ApplicationException
    Sub New()
        message = "Report folder not found."
    End Sub
    Public Overrides ReadOnly Property message As String
End Class

Public Class VBAFolderNotFoundException : Inherits ApplicationException
    Sub New()
        message = "VBA folder not found."
    End Sub
    Public Overrides ReadOnly Property message As String
End Class

Public Class RequirementNotFoundException : Inherits ApplicationException
    Sub New(vbProj As VBProject)
        message = "Requirement file not found for " & vbProj.Name
    End Sub
    Public Overrides ReadOnly Property message As String
End Class
