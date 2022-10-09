Imports System.Windows.Forms
Imports Extensibility
Imports Microsoft.Vbe.Interop
Imports Microsoft.Office.Core

Imports System.Runtime.InteropServices
Imports System.Drawing

Public Class Connect
    Implements Extensibility.IDTExtensibility2

    Private _VBE As VBE
    Private _AddIn As AddIn

    Private WithEvents _CommandBarButton1 As CommandBarButton
    Private _toolWindow1 As Window

    ' Buttons created by the add-in
    Private WithEvents _myStandardCommandBarButton As CommandBarButton
    Private WithEvents _myToolsCommandBarButton As CommandBarButton
    Private WithEvents _myCodeWindowCommandBarButton As CommandBarButton
    Private WithEvents _myToolBarButton As CommandBarButton
    Private WithEvents _myCommandBarPopup1Button As CommandBarButton
    Private WithEvents _myCommandBarPopup2Button As CommandBarButton

    ' CommandBars created by the add-in
    Private _myToolbar As CommandBar
    Private _myCommandBarPopup1 As CommandBarPopup
    Private _myCommandBarPopup2 As CommandBarPopup

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

        If Not _CommandBarButton1 Is Nothing Then

            _CommandBarButton1.Delete()
            _CommandBarButton1 = Nothing

        End If

        Try

            Select Case RemoveMode

                Case ext_DisconnectMode.ext_dm_HostShutdown, ext_DisconnectMode.ext_dm_UserClosed

                    ' Delete buttons on built-in commandbars
                    If Not (_myStandardCommandBarButton Is Nothing) Then
                        _myStandardCommandBarButton.Delete()
                    End If

                    If Not (_myCodeWindowCommandBarButton Is Nothing) Then
                        _myCodeWindowCommandBarButton.Delete()
                    End If

                    If Not (_myToolsCommandBarButton Is Nothing) Then
                        _myToolsCommandBarButton.Delete()
                    End If

                    ' Disconnect event handlers
                    _myToolBarButton = Nothing
                    _myCommandBarPopup1Button = Nothing
                    _myCommandBarPopup2Button = Nothing

                    ' Delete commandbars created by the add-in
                    If Not (_myToolbar Is Nothing) Then
                        _myToolbar.Delete()
                    End If

                    If Not (_myCommandBarPopup1 Is Nothing) Then
                        _myCommandBarPopup1.Delete()
                    End If

                    If Not (_myCommandBarPopup2 Is Nothing) Then
                        _myCommandBarPopup2.Delete()
                    End If

            End Select

        Catch e As System.Exception
            System.Windows.Forms.MessageBox.Show(e.ToString)
        End Try

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


        Dim standardCommandBar As CommandBar
        Dim commandBarControl As CommandBarControl

        Try

            'standardCommandBar = _VBE.CommandBars.Add("myBar", MsoBarPosition.msoBarTop, Temporary:=True)
            standardCommandBar = _VBE.CommandBars.Item("Standard")

            commandBarControl = standardCommandBar.Controls.Add(MsoControlType.msoControlButton)
            _CommandBarButton1 = DirectCast(commandBarControl, CommandBarButton)
            _CommandBarButton1.Caption = "Toolwindow 1"
            _CommandBarButton1.FaceId = 59
            _CommandBarButton1.Style = MsoButtonStyle.msoButtonIconAndCaption
            _CommandBarButton1.BeginGroup = True


        Catch ex As Exception

            MessageBox.Show(ex.ToString())

        End Try

        'TOOLBARS CREATION

        ' Constants for names of built-in commandbars of the VBA editor
        Const STANDARD_COMMANDBAR_NAME As String = "Standard"
        Const MENUBAR_COMMANDBAR_NAME As String = "Barre de menus"
        Const TOOLS_COMMANDBAR_NAME As String = "Tools"
        Const CODE_WINDOW_COMMANDBAR_NAME As String = "Code Window"

        ' Constants for names of commandbars created by the add-in
        Const MY_COMMANDBAR_POPUP1_NAME As String = "MyTemporaryCommandBarPopup1"
        Const MY_COMMANDBAR_POPUP2_NAME As String = "MyTemporaryCommandBarPopup2"

        ' Constants for captions of commandbars created by the add-in
        Const MY_COMMANDBAR_POPUP1_CAPTION As String = "My sub menu"
        Const MY_COMMANDBAR_POPUP2_CAPTION As String = "My main menu"
        Const MY_TOOLBAR_CAPTION As String = "My toolbar"

        ' Built-in commandbars of the VBA editor
        Dim menuCommandBar As CommandBar
        Dim toolsCommandBar As CommandBar
        Dim codeCommandBar As CommandBar

        ' Other variables
        Dim toolsCommandBarControl As CommandBarControl
        Dim position As Integer

        Try

            ' Retrieve some built-in commandbars
            standardCommandBar = _VBE.CommandBars.Item(STANDARD_COMMANDBAR_NAME)
            menuCommandBar = _VBE.CommandBars.Item(MENUBAR_COMMANDBAR_NAME)
            toolsCommandBar = _VBE.CommandBars.Item(TOOLS_COMMANDBAR_NAME)
            codeCommandBar = _VBE.CommandBars.Item(CODE_WINDOW_COMMANDBAR_NAME)

            ' Add a button to the built-in "Standard" toolbar
            _myStandardCommandBarButton = AddCommandBarButton(standardCommandBar)

            ' Add a button to the built-in "Tools" menu
            _myToolsCommandBarButton = AddCommandBarButton(toolsCommandBar)

            ' Add a button to the built-in "Code Window" context menu
            _myCodeWindowCommandBarButton = AddCommandBarButton(codeCommandBar)

            ' ------------------------------------------------------------------------------------
            ' New toolbar
            ' ------------------------------------------------------------------------------------

            ' Add a new toolbar 
            _myToolbar = _VBE.CommandBars.Add(MY_TOOLBAR_CAPTION, MsoBarPosition.msoBarTop, System.Type.Missing, True)

            ' Add a new button on that toolbar
            _myToolBarButton = AddCommandBarButton(_myToolbar)

            ' Make visible the toolbar
            _myToolbar.Visible = True

            ' ------------------------------------------------------------------------------------
            ' New submenu under the "Tools" menu
            ' ------------------------------------------------------------------------------------

            ' Add a new commandbar popup 
            _myCommandBarPopup1 = DirectCast(toolsCommandBar.Controls.Add(
            MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing,
            toolsCommandBar.Controls.Count + 1, True), CommandBarPopup)

            ' Change some commandbar popup properties
            _myCommandBarPopup1.CommandBar.Name = MY_COMMANDBAR_POPUP1_NAME
            _myCommandBarPopup1.Caption = MY_COMMANDBAR_POPUP1_CAPTION

            ' Add a new button on that commandbar popup
            _myCommandBarPopup1Button = AddCommandBarButton(_myCommandBarPopup1.CommandBar)

            ' Make visible the commandbar popup
            _myCommandBarPopup1.Visible = True

            ' ------------------------------------------------------------------------------------
            ' New main menu
            ' ------------------------------------------------------------------------------------

            ' Calculate the position of a new commandbar popup to the right of the "Tools" menu
            toolsCommandBarControl = DirectCast(toolsCommandBar.Parent, CommandBarControl)
            position = toolsCommandBarControl.Index + 1

            ' Add a new commandbar popup 
            _myCommandBarPopup2 = DirectCast(menuCommandBar.Controls.Add(
            MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing,
            position, True), CommandBarPopup)

            ' Change some commandbar popup properties
            _myCommandBarPopup2.CommandBar.Name = MY_COMMANDBAR_POPUP2_NAME
            _myCommandBarPopup2.Caption = MY_COMMANDBAR_POPUP2_CAPTION

            ' Add a new button on that commandbar popup
            _myCommandBarPopup2Button = AddCommandBarButton(_myCommandBarPopup2.CommandBar)

            ' Make visible the commandbar popup
            _myCommandBarPopup2.Visible = True

        Catch e As System.Exception
            System.Windows.Forms.MessageBox.Show(e.ToString)
        End Try


    End Sub

    Private Sub _myToolBarButton_Click(Ctrl As Microsoft.Office.Core.CommandBarButton,
        ByRef CancelDefault As Boolean) Handles _myToolBarButton.Click

        MessageBox.Show("Clicked " & Ctrl.Caption)

    End Sub

    Private Sub _myToolsCommandBarButton_Click(Ctrl As Microsoft.Office.Core.CommandBarButton,
       ByRef CancelDefault As Boolean) Handles _myToolsCommandBarButton.Click

        MessageBox.Show("Clicked " & Ctrl.Caption)

    End Sub

    Private Sub _myStandardCommandBarButton_Click(Ctrl As Microsoft.Office.Core.CommandBarButton,
       ByRef CancelDefault As Boolean) Handles _myStandardCommandBarButton.Click

        MessageBox.Show("Clicked " & Ctrl.Caption)

    End Sub

    Private Sub _myCodeWindowCommandBarButton_Click(Ctrl As Microsoft.Office.Core.CommandBarButton,
       ByRef CancelDefault As Boolean) Handles _myCodeWindowCommandBarButton.Click

        MessageBox.Show("Clicked " & Ctrl.Caption)

    End Sub

    Private Sub _myCommandBarPopup1Button_Click(Ctrl As Microsoft.Office.Core.CommandBarButton,
       ByRef CancelDefault As Boolean) Handles _myCommandBarPopup1Button.Click

        MessageBox.Show("Clicked " & Ctrl.Caption)

    End Sub

    Private Sub _myCommandBarPopup2Button_Click(Ctrl As Microsoft.Office.Core.CommandBarButton,
       ByRef CancelDefault As Boolean) Handles _myCommandBarPopup2Button.Click

        MessageBox.Show("Clicked " & Ctrl.Caption)

    End Sub

    Private Function AddCommandBarButton(ByVal commandBar As CommandBar) As CommandBarButton

        Dim commandBarButton As CommandBarButton
        Dim commandBarControl As CommandBarControl

        commandBarControl = commandBar.Controls.Add(MsoControlType.msoControlButton)
        commandBarButton = DirectCast(commandBarControl, CommandBarButton)

        commandBarButton.Caption = "My button"
        commandBarButton.FaceId = 59

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

    Private Sub _CommandBarButton1_Click(Ctrl As Microsoft.Office.Core.CommandBarButton,
       ByRef CancelDefault As Boolean) Handles _CommandBarButton1.Click

        Dim userControlObject As Object = Nothing
        Dim userControlToolWindow1 As UserControlToolWindow1

        Try

            If _toolWindow1 Is Nothing Then

                userControlToolWindow1 = New UserControlToolWindow1()

                ' TODO: Change the GUID
                _toolWindow1 = CreateToolWindow("My toolwindow 1", "{312945A4-6B7D-4F69-82CC-ACD0879011DB}", userControlToolWindow1)

                userControlToolWindow1.Initialize(_VBE)

            Else

                _toolWindow1.Visible = True

            End If

        Catch ex As Exception

            MessageBox.Show(ex.ToString)

        End Try

    End Sub

End Class

