Imports System.Windows.Forms
Imports Extensibility
Imports Microsoft.Vbe.Interop
Imports Microsoft.Office.Core
Imports System.Runtime.InteropServices

Public Class Import_Submenu : Inherits VBE_MenuBar

    Private WithEvents m_ImportAll As CommandBarButton
    Private WithEvents m_ImportCurrent As CommandBarButton
    Private WithEvents m_ImportCustom As CommandBarButton

    Private m_codeHandler As CodeHandler

    Public Sub New(VBE_ As VBE, parentCommandBar As CommandBar, name As String)

        MyBase.New(VBE_, parentCommandBar, name)

        m_VBE = VBE_
        m_codeHandler = New CodeHandler(m_VBE)

        m_ImportAll = Me.addButton("Import all files")
        m_ImportCurrent = Me.addButton("Import current file")
        m_ImportCustom = Me.addButton("Custom import...")

    End Sub

    Private Sub m_ImportAll_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_ImportAll.Click
        m_codeHandler.importSourceCode(m_VBE.ActiveVBProject)
    End Sub

    Private Sub m_ImportCurrent_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_ImportCurrent.Click
        MessageBox.Show(
            text:="Not coded yet !",
            caption:="VBEAddIn",
            buttons:=MessageBoxButtons.OK,
            icon:=MessageBoxIcon.Information)
    End Sub

    Private Sub m_ImportCustom_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_ImportCustom.Click
        MessageBox.Show(
            text:="Not coded yet !",
            caption:="VBEAddIn",
            buttons:=MessageBoxButtons.OK,
            icon:=MessageBoxIcon.Information)
    End Sub

End Class
