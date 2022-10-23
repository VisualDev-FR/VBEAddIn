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
    Private WithEvents m_displayHistoChecker_Button As CommandBarButton
    Private WithEvents m_openReportFolder_Button As CommandBarButton

    Private m_insert_histo_anchor_subMenu As HistoAssistant_Submenu
    Private m_import_Submenu As Import_Submenu
    Private m_export_Submenu As Export_Submenu
    Private m_codeHandler As CodeHandler

    Public Sub New(VBE_ As VBE, AddIn_ As AddIn, parentCommandBar As CommandBar, name As String)

        MyBase.New(VBE_, parentCommandBar, name, 11)

        m_VBE = VBE_
        m_AddIn = AddIn_
        m_codeHandler = New CodeHandler(VBE_)

        m_test_Button = addButton("TEST Button")
        m_openReportFolder_Button = addButton("Open report folder...")
        m_export_Submenu = New Export_Submenu(VBE_, m_CommandBarPopup.CommandBar, "Export")
        m_import_Submenu = New Import_Submenu(VBE_, m_CommandBarPopup.CommandBar, "Import")
        m_displayHistoChecker_Button = addButton("Histo checker")
        m_insert_histo_anchor_subMenu = New HistoAssistant_Submenu(VBE_, m_CommandBarPopup.CommandBar, "Histo Assistant")

    End Sub

    Private Sub m_test_Button_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_test_Button.Click
        Call TEST_MAIN()
    End Sub

    Private Sub m_openReportFolder_Button_Button_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_openReportFolder_Button.Click
        Call openMyReportFolder()
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
