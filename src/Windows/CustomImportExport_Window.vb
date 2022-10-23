Imports System.Windows.Forms
Imports Extensibility
Imports Microsoft.Vbe.Interop
Imports Microsoft.Office.Core
Imports System.Runtime.InteropServices

Public Class CustomImportExport_Window

    Private m_VBE As VBE
    Public Enum CodeHandlerAction
        IMPORT
        EXPORT
    End Enum

    Public Sub New(VBE_ As VBE, action As CodeHandlerAction)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        m_VBE = VBE_

    End Sub

End Class
