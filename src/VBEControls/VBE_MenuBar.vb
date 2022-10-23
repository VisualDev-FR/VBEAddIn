Imports System.Windows.Forms
Imports Extensibility
Imports Microsoft.Vbe.Interop
Imports Microsoft.Office.Core
Imports System.Runtime.InteropServices

Public MustInherit Class VBE_MenuBar

    Protected m_VBE As VBE
    Protected m_AddIn As AddIn
    Protected m_CommandBarPopup As CommandBarPopup

    Public Sub New(VBE_ As VBE, parentCommandBar As CommandBar, position As Integer, name As String)

        m_CommandBarPopup = DirectCast(
            parentCommandBar.Controls.Add(
                MsoControlType.msoControlPopup,
                System.Type.Missing,
                System.Type.Missing,
                position,
                True),
            CommandBarPopup)

        m_CommandBarPopup.Visible = True

    End Sub

    Public Sub Hide()
        m_CommandBarPopup.Visible = False
    End Sub

    Public Sub Show()
        m_CommandBarPopup.Visible = True
    End Sub

    Protected Sub setMyName(myName As String)
        m_CommandBarPopup.Caption = myName
    End Sub

    Protected Function addButton(ByRef buttonName As String) As CommandBarButton

        Dim commandBarControl As CommandBarControl = m_CommandBarPopup.CommandBar.Controls.Add(MsoControlType.msoControlButton)
        Dim commandBarButton As CommandBarButton = DirectCast(commandBarControl, CommandBarButton)

        commandBarButton.Caption = buttonName

        Return commandBarButton

    End Function

    Protected Function addSubMenu(ByRef subMenuName As String) As CommandBarPopup

        Dim subMenuPopup As CommandBarPopup = DirectCast(m_CommandBarPopup.CommandBar.Controls.Add(
        MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, m_CommandBarPopup.CommandBar.Controls.Count + 1, True), CommandBarPopup)

        subMenuPopup.Caption = subMenuName

        Return subMenuPopup

    End Function

End Class
