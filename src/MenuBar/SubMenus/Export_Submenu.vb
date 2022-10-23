Imports System.Windows.Forms
Imports Extensibility
Imports Microsoft.Vbe.Interop
Imports Microsoft.Office.Core
Imports System.Runtime.InteropServices

Public Class Export_Submenu : Inherits VBE_MenuBar

    Private WithEvents m_exportAll As CommandBarButton
    Private WithEvents m_exportCurrent As CommandBarButton
    Private WithEvents m_exportCustom As CommandBarButton

    Private m_customExportWindow As CustomImportExport_Window
    Private m_codeHandler As CodeHandler

    Public Sub New(VBE_ As VBE, parentCommandBar As CommandBar, name As String)

        MyBase.New(VBE_, parentCommandBar, name)

        m_VBE = VBE_
        m_codeHandler = New CodeHandler(VBE_)

        m_exportAll = Me.addButton("Export all files")
        m_exportCurrent = Me.addButton("Export current file")
        m_exportCustom = Me.addButton("Custom export...")

    End Sub

    Private Sub m_exportAll_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_exportAll.Click
        m_codeHandler.exportSourceCode(m_VBE.ActiveVBProject)

        ''TODO: Implémenter la gestion d'exception ici
        'If Not isVbProjSaved(vbProj) Then
        '    MessageBox.Show(
        '        text:=vbProj.Name & " file not found",
        '        caption:="VBEAddin.importSourceCode",
        '        buttons:=MessageBoxButtons.OK,
        '        icon:=MessageBoxIcon.Error)
        '    Exit Sub
        'ElseIf Not vbProjectSourceFolderExists(vbProj) Then
        '    MessageBox.Show(
        '        text:=vbProj.Name & " source folder not found",
        '        caption:="VBEAddin.importSourceCode",
        '        buttons:=MessageBoxButtons.OK,
        '        icon:=MessageBoxIcon.Error)
        '    Exit Sub
        'End If
    End Sub

    Private Sub m_exportCurrent_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_exportCurrent.Click
        MessageBox.Show(
            text:="Not coded yet !",
            caption:="VBEAddIn",
            buttons:=MessageBoxButtons.OK,
            icon:=MessageBoxIcon.Information)
    End Sub

    Private Sub m_exportCustom_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_exportCustom.Click
        MessageBox.Show(
            text:="Not coded yet !",
            caption:="VBEAddIn",
            buttons:=MessageBoxButtons.OK,
            icon:=MessageBoxIcon.Information)
    End Sub

End Class
