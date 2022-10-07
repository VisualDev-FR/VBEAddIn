Imports Microsoft.Office.Interop
Imports Extensibility
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports Microsoft.Vbe.Interop
Imports Microsoft.Office.Core

<ComVisible(True), Guid("875B3991-9A51-48AC-A328-ABE02EB53279"), ProgId("VBEAddIn.Connect")>
Public Class Connect
    Implements Extensibility.IDTExtensibility2

    Private _VBE As VBE
    Private _AddIn As AddIn

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
        'MessageBox.Show(_AddIn.ProgId & " loaded in VBA editor version " & _VBE.Version)

        Dim vbProj As Object
        For Each vbProj In _VBE.VBProjects

            Dim m_Window As Microsoft.Vbe.Interop.Window = _VBE.Windows.CreateToolWindow(_AddIn, "VBEAddIn.Connect", "MyWindow", "AnyString", Nothing)

        Next
    End Sub

End Class