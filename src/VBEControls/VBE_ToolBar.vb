Imports System.Windows.Forms
Imports Extensibility
Imports Microsoft.Vbe.Interop
Imports Microsoft.Office.Core
Imports System.Runtime.InteropServices

Public MustInherit Class VBE_ToolBar

    Private m_ToolBar As CommandBar

    Public Sub New(_VBE As VBE, name As String, position As MsoBarPosition)

        m_ToolBar = _VBE.CommandBars.Add(name, position, System.Type.Missing, True)
        m_ToolBar.Visible = True

    End Sub

    Public Sub Show()
        m_ToolBar.Visible = True
    End Sub

    Public Sub Hide()
        m_ToolBar.Visible = False
    End Sub

    Protected Function addButton(buttonName As String, buttonStyle As MsoButtonStyle) As CommandBarButton

        Dim commandBarControl As CommandBarControl = m_ToolBar.Controls.Add(MsoControlType.msoControlButton)
        Dim commandBarButton As CommandBarButton = DirectCast(commandBarControl, CommandBarButton)

        commandBarButton.Caption = buttonName
        commandBarButton.Style = buttonStyle

        Return commandBarButton

    End Function

End Class
