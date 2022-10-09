Imports System.Windows.Forms
Imports Extensibility
Imports Microsoft.Vbe.Interop
Imports Microsoft.Office.Core

Public Class HistoAssistant

    Private m_VBE As VBE

    Public Sub New(_VBE As VBE)
        m_VBE = _VBE
    End Sub

    Public Sub insert_BEGIN()

        If m_VBE.ActiveCodePane Is Nothing Then Exit Sub

        Dim activeLine As Long
        activeLine = getCursorLine()

        Dim activeProcName As String
        activeProcName = m_VBE.ActiveCodePane.CodeModule.ProcOfLine(activeLine, vbext_ProcKind.vbext_pk_Proc)

        m_VBE.ActiveCodePane.CodeModule.InsertLines(activeLine, vbTab & "Call OOXOOXOOXOOXOOXOO(MODULE_NAME, """ & activeProcName & """)")

    End Sub

    Public Sub insert_END()

        If m_VBE.ActiveCodePane Is Nothing Then Exit Sub

        Dim activeLine As Long
        activeLine = getCursorLine()

        Call m_VBE.ActiveCodePane.CodeModule.InsertLines(activeLine, vbTab & "Call OOXOOXOOXOOXOOXOO(END_HISTO)")

    End Sub

    Public Sub insert_EXIT()

        If m_VBE.ActiveCodePane Is Nothing Then Exit Sub

        Dim activeLine As Long
        activeLine = getCursorLine()

        Call m_VBE.ActiveCodePane.CodeModule.InsertLines(activeLine, vbTab & "Call OOXOOXOOXOOXOOXOO(EXIT_HISTO)")

    End Sub
    Public Sub insert_ERROR()

        If m_VBE.ActiveCodePane Is Nothing Then Exit Sub

        Dim activeLine As Long
        activeLine = getCursorLine()

        Call m_VBE.ActiveCodePane.CodeModule.InsertLines(activeLine, vbTab & "Call OOXOOXOOXOOXOOXOO(ERROR_HISTO)")

    End Sub

    Public Sub insert_FATAL()

        If m_VBE.ActiveCodePane Is Nothing Then Exit Sub

        Dim activeLine As Long
        activeLine = getCursorLine()

        Call m_VBE.ActiveCodePane.CodeModule.InsertLines(activeLine, vbTab & "Call FatalError(True, True)")

    End Sub

    Private Function getCursorLine() As Long

        Dim sRow As Long, sCol As Long, eRow As Long, eCol As Long
        m_VBE.ActiveCodePane.GetSelection(sRow, sCol, eRow, eCol)

        getCursorLine = sRow

    End Function

End Class
