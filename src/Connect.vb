Imports System.Windows.Forms
Imports Extensibility
Imports Microsoft.Vbe.Interop
Imports Microsoft.Office.Core
Imports System.Runtime.InteropServices

<ComVisible(True), Guid("875B3991-9A51-48AC-A328-ABE02EB53279"), ProgId("VBEAddIn.Connect")>
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
    Private m_histoCheckerWindow As VBE_ToolWindow
    Private m_etid_mainmenu As ETITD_MenuBar
    Private m_insert_histo_anchor_subMenu As HistoAssistant_Submenu
    Private m_histo_checker_toolbar As HistoAssistant_Toolbar

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

                    'disconnect(m_exportSourceCode_Button)
                    'disconnect(m_importSourceCode_Button)
                    'disconnect(m_displayHistoChecker_Button)

                    'disconnect(m_begin_subMenuBtn)
                    'disconnect(m_end_subMenuBtn)
                    'disconnect(m_exit_subMenuBtn)
                    'disconnect(m_error_subMenuBtn)
                    'disconnect(m_fatal_subMenuBtn)
                    'disconnect(m_begin_toolbarBtn)
                    'disconnect(m_end_toolbarBtn)
                    'disconnect(m_exit_toolbarBtn)
                    'disconnect(m_error_toolbarBtn)
                    'disconnect(m_fatal_toolbarBtn)

                    'disconnect(m_histo_checker_toolbar)
                    'disconnect(m_insert_histo_anchor_subMenu)
                    'disconnect(m_etid_mainmenu)

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

            m_histo_checker_toolbar = New HistoAssistant_Toolbar(_VBE, HISTO_CHECKER_NAME, MsoBarPosition.msoBarTop)
            m_etid_mainmenu = New ETITD_MenuBar(_VBE, _AddIn, menuCommandBar, 11, ETITD_MENU_NAME)
            m_insert_histo_anchor_subMenu = New HistoAssistant_Submenu(_VBE, m_etid_mainmenu.commandBar, m_etid_mainmenu.commandBar.Controls.Count + 1, "Histo Assistant")

        Catch e As System.Exception
            System.Windows.Forms.MessageBox.Show(e.ToString)
        End Try


    End Sub

End Class