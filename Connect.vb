Imports System.Windows.Forms
Imports Extensibility
Imports Microsoft.Vbe.Interop
Imports Microsoft.Office.Core

Public Class Connect
    Implements Extensibility.IDTExtensibility2

    Private _VBE As VBE
    Private _AddIn As AddIn

    ' Constants for names of built-in commandbars of the VBA editor
    Private Const STANDARD_COMMANDBAR_NAME As String = "Standard"
    Private Const MENUBAR_COMMANDBAR_NAME As String = "Barre de menus"
    Private Const TOOLS_COMMANDBAR_NAME As String = "Tools"
    Private Const CODE_WINDOW_COMMANDBAR_NAME As String = "Code Window"

    ' Constants for names of commandbars created by the add-in
    Const ETITD_MENU_NAME As String = "ETITD"
    Const HISTO_CHECKER_NAME As String = "Check Histo"

    'Tool window
    Private m_histoCheckerWindow As Window '_toolWindow1

    'ETITD mainMenu
    Private m_etid_mainmenu As CommandBarPopup
    Private WithEvents m_exportSourceCode_Button As CommandBarButton
    Private WithEvents m_importSourceCode_Button As CommandBarButton
    Private WithEvents m_displayHistoChecker_Button As CommandBarButton

    'Histo anchor sub menu
    Private m_insert_histo_anchor_subMenu As CommandBarPopup
    Private WithEvents m_begin_subMenuBtn As CommandBarButton
    Private WithEvents m_end_subMenuBtn As CommandBarButton
    Private WithEvents m_exit_subMenuBtn As CommandBarButton
    Private WithEvents m_error_subMenuBtn As CommandBarButton
    Private WithEvents m_fatal_subMenuBtn As CommandBarButton

    'Histo checker ToolBar
    Private m_histo_checker_toolbar As CommandBar
    Private WithEvents m_begin_toolbarBtn As CommandBarButton
    Private WithEvents m_end_toolbarBtn As CommandBarButton
    Private WithEvents m_exit_toolbarBtn As CommandBarButton
    Private WithEvents m_error_toolbarBtn As CommandBarButton
    Private WithEvents m_fatal_toolbarBtn As CommandBarButton

    'Custom proc's
    Private m_Assistant As HistoAssistant

    'm_etid_mainmenu
    'm_insert_histo_anchor_subMenu
    'm_histo_checker_toolbar

    'm_exportSourceCode_Button
    'm_importSourceCode_Button
    'm_displayHistoChecker_Button

    'm_begin_subMenuBtn
    'm_end_subMenuBtn
    'm_exit_subMenuBtn
    'm_error_subMenuBtn
    'm_fatal_subMenuBtn
    'm_begin_toolbarBtn
    'm_end_toolbarBtn
    'm_exit_toolbarBtn
    'm_error_toolbarBtn
    'm_fatal_toolbarBtn

    Private Sub OnConnection(Application As Object, ConnectMode As Extensibility.ext_ConnectMode,
       AddInInst As Object, ByRef custom As System.Array) Implements IDTExtensibility2.OnConnection
        Try
            _VBE = DirectCast(Application, VBE)
            _AddIn = DirectCast(AddInInst, AddIn)
            Select Case ConnectMode
                Case Extensibility.ext_ConnectMode.ext_cm_Startup
                Case Extensibility.ext_ConnectMode.ext_cm_AfterStartup
                    InitializeAddIn()
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try
    End Sub

    Private Sub OnDisconnection(RemoveMode As Extensibility.ext_DisconnectMode,
       ByRef custom As System.Array) Implements IDTExtensibility2.OnDisconnection

        Try

            Select Case RemoveMode

                Case ext_DisconnectMode.ext_dm_HostShutdown, ext_DisconnectMode.ext_dm_UserClosed

                    disconnect(m_exportSourceCode_Button)
                    disconnect(m_importSourceCode_Button)
                    disconnect(m_displayHistoChecker_Button)

                    disconnect(m_begin_subMenuBtn)
                    disconnect(m_end_subMenuBtn)
                    disconnect(m_exit_subMenuBtn)
                    disconnect(m_error_subMenuBtn)
                    disconnect(m_fatal_subMenuBtn)
                    disconnect(m_begin_toolbarBtn)
                    disconnect(m_end_toolbarBtn)
                    disconnect(m_exit_toolbarBtn)
                    disconnect(m_error_toolbarBtn)
                    disconnect(m_fatal_toolbarBtn)

                    disconnect(m_histo_checker_toolbar)
                    disconnect(m_insert_histo_anchor_subMenu)
                    disconnect(m_etid_mainmenu)

            End Select

        Catch e As System.Exception
            System.Windows.Forms.MessageBox.Show(e.ToString)
        End Try

    End Sub

    Private Sub disconnect(ByRef mObject As Object)
        If Not (mObject Is Nothing) Then
            mObject.Delete()
        End If
        mObject = Nothing
    End Sub

    Private Sub OnStartupComplete(ByRef custom As System.Array) _
       Implements IDTExtensibility2.OnStartupComplete
        InitializeAddIn()
    End Sub

    Private Sub OnAddInsUpdate(ByRef custom As System.Array) Implements IDTExtensibility2.OnAddInsUpdate

    End Sub

    Private Sub OnBeginShutdown(ByRef custom As System.Array) Implements IDTExtensibility2.OnBeginShutdown

    End Sub

    Private Sub InitializeAddIn()

        m_Assistant = New HistoAssistant(_VBE)

        ' Built-in commandbars of the VBA editor
        Dim standardCommandBar As CommandBar
        Dim menuCommandBar As CommandBar
        Dim toolsCommandBar As CommandBar
        Dim codeCommandBar As CommandBar

        ' Other variables
        Dim toolsCommandBarControl As CommandBarControl
        Dim position As Integer

        Try

            'Retrieve some built-in commandbars
            standardCommandBar = _VBE.CommandBars.Item(STANDARD_COMMANDBAR_NAME)
            menuCommandBar = _VBE.CommandBars.Item(MENUBAR_COMMANDBAR_NAME)
            toolsCommandBar = _VBE.CommandBars.Item(TOOLS_COMMANDBAR_NAME)
            codeCommandBar = _VBE.CommandBars.Item(CODE_WINDOW_COMMANDBAR_NAME)

            ' ------------------------------------------------------------------------------------
            ' Create histoChecker toolbar
            ' ------------------------------------------------------------------------------------

            m_histo_checker_toolbar = _VBE.CommandBars.Add(HISTO_CHECKER_NAME, MsoBarPosition.msoBarTop, System.Type.Missing, True)
            m_histo_checker_toolbar.Visible = True

            m_begin_toolbarBtn = AddCommandBarButton(m_histo_checker_toolbar, "BEGIN")
            m_begin_toolbarBtn.Caption = "BEGIN"
            m_begin_toolbarBtn.Style = MsoButtonStyle.msoButtonCaption

            m_end_toolbarBtn = AddCommandBarButton(m_histo_checker_toolbar, "END")
            m_end_toolbarBtn.Caption = "END"
            m_end_toolbarBtn.Style = MsoButtonStyle.msoButtonCaption

            m_exit_toolbarBtn = AddCommandBarButton(m_histo_checker_toolbar, "EXIT")
            m_exit_toolbarBtn.Caption = "EXIT"
            m_exit_toolbarBtn.Style = MsoButtonStyle.msoButtonCaption

            m_error_toolbarBtn = AddCommandBarButton(m_histo_checker_toolbar, "ERROR")
            m_error_toolbarBtn.Caption = "ERROR"
            m_error_toolbarBtn.Style = MsoButtonStyle.msoButtonCaption

            m_fatal_toolbarBtn = AddCommandBarButton(m_histo_checker_toolbar, "FATAl")
            m_fatal_toolbarBtn.Caption = "FATAL"
            m_fatal_toolbarBtn.Style = MsoButtonStyle.msoButtonCaption

            ' ------------------------------------------------------------------------------------
            ' ETITD main menu
            ' ------------------------------------------------------------------------------------

            toolsCommandBarControl = DirectCast(toolsCommandBar.Parent, CommandBarControl) 'Calculate the position of a new commandbar popup to the right of the "Tools" menu
            position = 11 ' toolsCommandBarControl.Index + 1

            m_etid_mainmenu = DirectCast(menuCommandBar.Controls.Add(
            MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing,
            position, True), CommandBarPopup)

            m_etid_mainmenu.CommandBar.Name = ETITD_MENU_NAME
            m_etid_mainmenu.Caption = ETITD_MENU_NAME
            m_etid_mainmenu.Visible = True

            m_exportSourceCode_Button = AddCommandBarButton(m_etid_mainmenu.CommandBar, "Export code")
            m_importSourceCode_Button = AddCommandBarButton(m_etid_mainmenu.CommandBar, "Import code")
            m_displayHistoChecker_Button = AddCommandBarButton(m_etid_mainmenu.CommandBar, "Histo checker")

            ' ------------------------------------------------------------------------------------
            ' New submenu under the "ETITD" menu
            ' ------------------------------------------------------------------------------------

            'TODO: recabler le sous-menu
            m_insert_histo_anchor_subMenu = DirectCast(m_etid_mainmenu.CommandBar.Controls.Add(
            MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, m_etid_mainmenu.CommandBar.Controls.Count + 1, True),
            CommandBarPopup)

            m_insert_histo_anchor_subMenu.CommandBar.Name = "Insert histo anchors"
            m_insert_histo_anchor_subMenu.Caption = "Insert histo anchors"
            m_insert_histo_anchor_subMenu.Visible = True

            m_begin_subMenuBtn = AddCommandBarButton(m_insert_histo_anchor_subMenu.CommandBar, "BEGIN")
            m_end_subMenuBtn = AddCommandBarButton(m_insert_histo_anchor_subMenu.CommandBar, "END")
            m_exit_subMenuBtn = AddCommandBarButton(m_insert_histo_anchor_subMenu.CommandBar, "EXIT")
            m_error_subMenuBtn = AddCommandBarButton(m_insert_histo_anchor_subMenu.CommandBar, "ERROR")
            m_fatal_subMenuBtn = AddCommandBarButton(m_insert_histo_anchor_subMenu.CommandBar, "FATAL")

        Catch e As System.Exception
            System.Windows.Forms.MessageBox.Show(e.ToString)
        End Try


    End Sub

    Private Sub m_exportSourceCode_Button_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_exportSourceCode_Button.Click
        MessageBox.Show("Clicked " & Ctrl.Caption)
    End Sub

    Private Sub m_importSourceCode_Button_Button_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_importSourceCode_Button.Click
        MessageBox.Show("Clicked " & Ctrl.Caption)
    End Sub

    Private Sub m_begin_subMenuBtn_Button_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_begin_subMenuBtn.Click
        m_Assistant.insert_BEGIN()
    End Sub

    Private Sub m_end_subMenuBtn_Button_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_end_subMenuBtn.Click
        m_Assistant.insert_END()
    End Sub

    Private Sub m_exit_subMenuBtn_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_exit_subMenuBtn.Click
        m_Assistant.insert_EXIT()
    End Sub

    Private Sub m_error_subMenuBtn_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_error_subMenuBtn.Click
        m_Assistant.insert_ERROR()
    End Sub

    Private Sub m_fatal_subMenuBtn_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_fatal_subMenuBtn.Click
        m_Assistant.insert_FATAL()
    End Sub

    Private Sub m_begin_toolbarBtn_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_begin_toolbarBtn.Click
        m_Assistant.insert_BEGIN()
    End Sub

    Private Sub m_end_toolbarBtn_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_end_toolbarBtn.Click
        m_Assistant.insert_END()
    End Sub

    Private Sub m_exit_toolbarBtn_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_exit_toolbarBtn.Click
        m_Assistant.insert_EXIT()
    End Sub

    Private Sub m_error_toolbarBtn_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_error_toolbarBtn.Click
        m_Assistant.insert_ERROR()
    End Sub

    Private Sub m_fatal_toolbarBtn_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles m_fatal_toolbarBtn.Click
        m_Assistant.insert_FATAL()
    End Sub

    Private Function AddCommandBarButton(ByVal commandBar As CommandBar, ByRef buttonName As String) As CommandBarButton

        Dim commandBarButton As CommandBarButton
        Dim commandBarControl As CommandBarControl

        commandBarControl = commandBar.Controls.Add(MsoControlType.msoControlButton)
        commandBarButton = DirectCast(commandBarControl, CommandBarButton)

        commandBarButton.Caption = buttonName
        'commandBarButton.FaceId = 59

        Return commandBarButton

    End Function

    Private Function CreateToolWindow(ByVal toolWindowCaption As String, ByVal toolWindowGuid As String,
       ByVal toolWindowUserControl As UserControl) As Window

        Dim userControlObject As Object = Nothing
        Dim userControlHost As UserControlHost
        Dim toolWindow As Window
        Dim progId As String

        ' IMPORTANT: ensure that you use the same ProgId value used in the ProgId attribute of the UserControlHost class
        progId = "VBEAddIn.UserControlHost"

        toolWindow = _VBE.Windows.CreateToolWindow(_AddIn, progId, toolWindowCaption, toolWindowGuid, userControlObject)
        userControlHost = DirectCast(userControlObject, UserControlHost)

        toolWindow.Visible = True

        userControlHost.AddUserControl(toolWindowUserControl)

        Return toolWindow

    End Function

    Private Sub m_displayHistoChecker_Button_Click(Ctrl As Microsoft.Office.Core.CommandBarButton,
       ByRef CancelDefault As Boolean) Handles m_displayHistoChecker_Button.Click

        Dim userControlObject As Object = Nothing
        Dim userControlToolWindow1 As UserControlToolWindow1

        Try

            If m_histoCheckerWindow Is Nothing Then

                userControlToolWindow1 = New UserControlToolWindow1()

                ' TODO: Change the GUID
                m_histoCheckerWindow = CreateToolWindow("Histo checker", "{312945A4-6B7D-4F69-82CC-ACD0879011DB}", userControlToolWindow1)

                userControlToolWindow1.Initialize(_VBE)

            Else
                m_histoCheckerWindow.Visible = True
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

End Class