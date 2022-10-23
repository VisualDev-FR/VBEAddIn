Imports System.Windows.Forms
Imports Extensibility
Imports Microsoft.Vbe.Interop
Imports Microsoft.Office.Core
Imports System.Runtime.InteropServices

Public Class ETITD_MenuBar : Inherits VBE_MenuBar

    'Tool window
    Private m_histoCheckerWindow As VBE_ToolWindow '_toolWindow1

    'ETITD mainMenu
    Private WithEvents m_test_Button As CommandBarButton
    Private WithEvents m_exportSourceCode_Button As CommandBarButton
    Private WithEvents m_importSourceCode_Button As CommandBarButton
    Private WithEvents m_displayHistoChecker_Button As CommandBarButton

    Private m_codeHandler As CodeHandler

    Public Sub New(VBE_ As VBE, AddIn_ As AddIn, parentCommandBar As CommandBar, position As Integer, name As String)

        MyBase.New(VBE_, parentCommandBar, position, name)

        m_VBE = VBE_
        m_AddIn = AddIn_
        m_codeHandler = New CodeHandler(VBE_)

        m_test_Button = addButton("TEST Button")
        m_exportSourceCode_Button = addButton("Export code")
        m_importSourceCode_Button = addButton("Import code")
        m_displayHistoChecker_Button = addButton("Histo checker")

    End Sub

    Private Sub m_test_Button_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_test_Button.Click
        Call TEST_MAIN()
    End Sub

    Private Sub m_exportSourceCode_Button_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_exportSourceCode_Button.Click
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

    Private Sub m_importSourceCode_Button_Button_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_importSourceCode_Button.Click
        m_codeHandler.importSourceCode(m_VBE.ActiveVBProject)
    End Sub

    Private Sub m_displayHistoChecker_Button_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_displayHistoChecker_Button.Click

        Dim myHistoChecker As HistoChecker_ToolWindow

        Try
            If m_histoCheckerWindow Is Nothing Then

                myHistoChecker = New HistoChecker_ToolWindow()
                m_histoCheckerWindow = New VBE_ToolWindow(m_VBE, m_AddIn, "Histo checker", "{312945A4-6B7D-4F69-82CC-ACD0879011DB}", myHistoChecker)
                myHistoChecker.Initialize(m_VBE)
            Else
                m_histoCheckerWindow.Show()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub


End Class
