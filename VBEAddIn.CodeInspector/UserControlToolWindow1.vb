Imports Microsoft.Vbe.Interop
Imports System.Windows.Forms

Friend Class UserControlToolWindow1

    Private _VBE As VBE

    Friend Sub Initialize(ByVal vbe As VBE)

        _VBE = vbe

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs)

        MessageBox.Show("Toolwindow shown in VBA editor version " & _VBE.Version)

    End Sub

End Class