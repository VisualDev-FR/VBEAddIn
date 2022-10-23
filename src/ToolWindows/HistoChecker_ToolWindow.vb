Imports Microsoft.Vbe.Interop
Imports System.Windows.Forms

Friend Class HistoChecker_ToolWindow

    'MEMBERS
    Private m_AnchorsDico As Dictionary(Of String, Object)
    Private m_VBE As VBE
    Private enabledEvents As Boolean

    'MAIN FUNCTIONS
    Friend Sub Initialize(ByVal vbe As VBE)
        enabledEvents = True
        m_VBE = vbe
        Call reload()
    End Sub

    Private Sub reload()

        m_AnchorsDico = getAnchorsDico()

        Call refreshVbProjList()
        Call refreshVbCompList()
        Call refreshView()

    End Sub

    'CALCULATION FUNCTIONS
    Private Sub refreshVbProjList()

        enabledEvents = False

        Me.vbProjList.Items.Clear()

        For Each vbProjKey As String In m_AnchorsDico.Keys
            Me.vbProjList.Items.Add(vbProjKey)
        Next

        Me.vbProjList.Text = Me.vbProjList.Items(0)

        enabledEvents = True

    End Sub

    Private Sub refreshVbCompList()

        enabledEvents = False

        Me.vbCompList.Items.Clear()

        If Not m_AnchorsDico.ContainsKey(vbProjList.Text) Then Exit Sub

        For Each vbCompKey As String In m_AnchorsDico.Item(vbProjList.Text).Keys
            Me.vbCompList.Items.Add(vbCompKey)
        Next

        Me.vbCompList.SelectedIndex = 0 'Calls vbCompList_SelectedIndexChanged event

        enabledEvents = True

    End Sub

    Private Sub refreshView()

        enabledEvents = False

        Me.view.Items.Clear()

        If Not m_AnchorsDico.ContainsKey(vbProjList.Text) Then Exit Sub
        If Not TryCast(m_AnchorsDico.Item(vbProjList.Text), Dictionary(Of String, Object)).ContainsKey(vbCompList.Text) Then Exit Sub

        Dim anchors As Dictionary(Of Integer, Object) = m_AnchorsDico.Item(vbProjList.Text).Item(vbCompList.Text)

        For Each line As String In anchors.Keys

            Me.view.Items.Add(line - 1 & ":" & vbTab & anchors.Item(line).Item("prevLine"))
            Me.view.Items.Add(line & ":" & vbTab & anchors.Item(line).Item("curLine"))
            Me.view.Items.Add("")

        Next line

        enabledEvents = True

    End Sub

    Private Function getAnchorsDico() As Dictionary(Of String, Object)

        Dim appAnchors As Dictionary(Of String, Object) = New Dictionary(Of String, Object)

        Dim vbProj As VBProject
        For Each vbProj In m_VBE.VBProjects

            Dim vbProjDico As Dictionary(Of String, Object) = New Dictionary(Of String, Object)

            If vbProj.Protection <> vbext_ProjectProtection.vbext_pp_locked Then

                Dim vbComp As VBComponent
                For Each vbComp In vbProj.VBComponents

                    Dim vbCompDico As Dictionary(Of Integer, Object) = getvbCompDico(vbComp)

                    If vbCompDico.Count > 0 And Not vbProjDico.ContainsKey(vbComp.Name) Then vbProjDico.Add(key:=vbComp.Name, value:=vbCompDico)

                Next vbComp

                If vbProjDico.Count > 0 And Not appAnchors.ContainsKey(vbProj.Name) Then appAnchors.Add(key:=vbProj.Name, value:=vbProjDico)

            End If

        Next vbProj

        Return appAnchors

    End Function

    Private Function getvbCompDico(vbComp As VBComponent) As Dictionary(Of Integer, Object)

        Dim vbCompDico As Dictionary(Of Integer, Object) = New Dictionary(Of Integer, Object)

        With vbComp.CodeModule

            For i As Integer = 1 To .CountOfLines

                Dim curLine As String = .Lines(i, 1)
                Dim prevLine As String = ""

                If i > 1 Then prevLine = .Lines(i - 1, 1)

                If isEndLine(curLine) Then

                    Dim checkPrev As Boolean, checkCur As Boolean

                    checkCur = curLine.Contains("OOXOOXOOXOOXOOXOO")
                    checkPrev = prevLine.Contains("OOXOOXOOXOOXOOXOO")

                    If Not checkCur And Not checkPrev Then

                        Dim anchorDico As Dictionary(Of String, Object) = New Dictionary(Of String, Object)

                        anchorDico.Add(key:="prevLine", value:=Trim(prevLine))
                        anchorDico.Add(key:="curLine", value:=Trim(curLine))

                        If Not vbCompDico.ContainsKey(i) Then vbCompDico.Add(key:=i, value:=anchorDico)

                    End If

                End If

            Next i

        End With

        Return vbCompDico

    End Function

    'GENERIC FUNCTIONS
    Private Function getSelectedVBComp() As VBComponent

        Dim vbProjName As String, vbCompName As String

        vbProjName = Me.vbProjList.Text
        vbCompName = Me.vbCompList.Text

        For Each vbProj As VBProject In m_VBE.VBProjects

            If vbProj.Name = vbProjName Then

                For Each vbComp As VBComponent In vbProj.VBComponents

                    If vbComp.Name = vbCompName Then
                        Return vbComp
                        Exit Function

                    End If
                Next
            End If
        Next

        Return Nothing

    End Function

    Private Function isEndLine(strLine As String) As Boolean

        Dim isEnd As Boolean, exitPos As Long, commentPos As Long

        commentPos = InStr(1, strLine, "'", vbTextCompare)

        exitPos = InStr(1, strLine, "Exit Sub", vbTextCompare)
        isEnd = isEnd Or (exitPos > 0 And (commentPos <= 0 Or commentPos > exitPos)) 'Permet de filtrer les lignes de commentaire

        exitPos = InStr(1, strLine, "Exit Function", vbTextCompare)
        isEnd = isEnd Or (exitPos > 0 And (commentPos <= 0 Or commentPos > exitPos))

        exitPos = InStr(1, strLine, "Exit Property", vbTextCompare)
        isEnd = isEnd Or (exitPos > 0 And (commentPos <= 0 Or commentPos > exitPos))

        exitPos = InStr(1, strLine, "End Sub", vbTextCompare)
        isEnd = isEnd Or (exitPos > 0 And (commentPos <= 0 Or commentPos > exitPos))

        exitPos = InStr(1, strLine, "End Function", vbTextCompare)
        isEnd = isEnd Or (exitPos > 0 And (commentPos <= 0 Or commentPos > exitPos))

        exitPos = InStr(1, strLine, "End Property", vbTextCompare)
        isEnd = isEnd Or (exitPos > 0 And (commentPos <= 0 Or commentPos > exitPos))

        isEndLine = isEnd

    End Function

    'EVENTS HANDLING

    Private Sub vbProjList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles vbProjList.SelectedIndexChanged
        If Not enabledEvents Then Exit Sub
        Call refreshVbProjList()
    End Sub

    Private Sub vbCompList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles vbCompList.SelectedIndexChanged
        If Not enabledEvents Then Exit Sub
        Call refreshView()
    End Sub

    Private Sub btn_refresh_Click(sender As Object, e As EventArgs) Handles btn_refresh.Click

        enabledEvents = False

        Dim activevbProj As String = vbProjList.Text
        Dim activeVbComp As String = vbCompList.Text

        Call reload()

        vbProjList.Text = IIf(activevbProj <> "", activevbProj, vbProjList.Text)

        enabledEvents = True
        vbCompList.Text = IIf(activeVbComp <> "", activeVbComp, vbCompList.Text) 'Calls 

    End Sub

    Private Sub view_SelectedIndexChanged(sender As Object, e As EventArgs) Handles view.SelectedIndexChanged

        If Not enabledEvents Or Me.view.SelectedItem Is Nothing Then Exit Sub

        Dim strLine As String = Me.view.SelectedItem.ToString()

        If strLine.Equals("") Then Exit Sub

        Dim activeVbComp As VBComponent = getSelectedVBComp()
        Dim activeLine As Long = Val(strLine)

        activeVbComp.CodeModule.CodePane.Show()
        activeVbComp.CodeModule.CodePane.TopLine = Math.Max(activeLine - activeVbComp.CodeModule.CodePane.CountOfVisibleLines + 10, 1)
        activeVbComp.CodeModule.CodePane.SetSelection(activeLine, 1, activeLine + 1, 1)

        m_VBE.ActiveCodePane.Window.SetFocus()

    End Sub
End Class