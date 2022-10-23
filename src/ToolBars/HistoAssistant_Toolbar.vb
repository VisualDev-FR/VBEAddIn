Imports System.Windows.Forms
Imports Extensibility
Imports Microsoft.Vbe.Interop
Imports Microsoft.Office.Core
Imports System.Runtime.InteropServices

Public Class HistoAssistant_Toolbar : Inherits VBE_ToolBar

    Private WithEvents m_BeginButton As CommandBarButton
    Private WithEvents m_endButton As CommandBarButton
    Private WithEvents m_exitButton As CommandBarButton
    Private WithEvents m_error_button As CommandBarButton
    Private WithEvents m_fatalButton As CommandBarButton

    Private m_VBE As VBE
    Private m_assistant As HistoAssistant

    Public Sub New(_VBE As VBE, name As String, position As MsoBarPosition)

        MyBase.New(_VBE, name, position)

        m_VBE = _VBE
        m_assistant = New HistoAssistant(_VBE)

        m_BeginButton = Me.addButton("BEGIN", MsoButtonStyle.msoButtonCaption)
        m_endButton = Me.addButton("END", MsoButtonStyle.msoButtonCaption)
        m_exitButton = Me.addButton("EXIT", MsoButtonStyle.msoButtonCaption)
        m_error_button = Me.addButton("ERROR", MsoButtonStyle.msoButtonCaption)
        m_fatalButton = Me.addButton("FATAL", MsoButtonStyle.msoButtonCaption)

    End Sub

    Private Sub m_begin_toolbarBtn_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_BeginButton.Click
        m_assistant.insert_BEGIN()
    End Sub

    Private Sub m_end_toolbarBtn_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_endButton.Click
        m_assistant.insert_END()
    End Sub

    Private Sub m_exit_toolbarBtn_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_exitButton.Click
        m_assistant.insert_EXIT()
    End Sub

    Private Sub m_error_toolbarBtn_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_error_button.Click
        m_assistant.insert_ERROR()
    End Sub

    Private Sub m_fatal_toolbarBtn_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_fatalButton.Click
        m_assistant.insert_FATAL()
    End Sub

End Class
