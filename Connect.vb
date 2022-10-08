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

    End Sub

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