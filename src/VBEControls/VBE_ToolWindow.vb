Imports System.Windows.Forms
Imports Extensibility
Imports Microsoft.Vbe.Interop
Imports Microsoft.Office.Core
Imports System.Runtime.InteropServices

Public Class VBE_ToolWindow

    Private Const PROG_ID As String = "VBEAddIn.UserControlHost"

    Private m_VBE As VBE
    Private m_AddIn As AddIn
    Private m_Window As Window

    Public Sub New(VBE_ As VBE, AddIn_ As AddIn, caption As String, windowGuid As String, ByVal toolWindowUserControl As UserControl)

        m_VBE = VBE_
        m_AddIn = AddIn_

        Dim userControlObject As Object = Nothing
        Dim userControlHost As UserControlHost

        m_Window = m_VBE.Windows.CreateToolWindow(m_AddIn, PROG_ID, caption, windowGuid, userControlObject)
        userControlHost = DirectCast(userControlObject, UserControlHost)

        m_Window.Visible = True

        userControlHost.AddUserControl(toolWindowUserControl)

    End Sub

    Public Sub Show()
        m_Window.Visible = True
    End Sub


End Class
