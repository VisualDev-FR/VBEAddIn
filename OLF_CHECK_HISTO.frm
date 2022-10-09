VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OLF_CHECK_HISTO 
   Caption         =   "Histo checker"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5940
   OleObjectBlob   =   "OLF_CHECK_HISTO.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OLF_CHECK_HISTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_anchorsDico As Dictionary

Private Sub btn_refresh_Click()
    
    Application.EnableEvents = False
    
    Dim activevbProj As String, activeVbComp As String
    
    activevbProj = vbProjList.Text
    activeVbComp = vbCompList.Text
    
    Call reloadUSF

    vbProjList.Text = activevbProj
    
    Application.EnableEvents = True
    vbCompList.Text = activeVbComp
    
End Sub

Private Sub UserForm_Initialize()

    Application.VBE.ActiveWindow.SetFocus
    Call reloadUSF

End Sub

Private Sub reloadUSF()
    Set m_anchorsDico = getAnchorsDico
    Call refreshVbProjList
End Sub

Private Sub refreshVbProjList()
    
    Me.vbProjList.Clear
    
    Dim vbProjKey As Variant
    For Each vbProjKey In m_anchorsDico.Keys
        Me.vbProjList.AddItem (vbProjKey)
    Next
    
    Me.vbProjList.ListIndex = 0

End Sub

Private Sub refreshVbCompList()
    
    Me.vbCompList.Clear
    
    If Not m_anchorsDico.Exists(vbProjList.Text) Then Exit Sub
    
    Dim vbCompKey
    For Each vbCompKey In m_anchorsDico.Item(vbProjList.Text).Keys
        Me.vbCompList.AddItem (vbCompKey)
    Next
    
    Me.vbCompList.ListIndex = 0

End Sub

Private Sub refreshView()

    Me.view.Clear
    
    If Not m_anchorsDico.Exists(vbProjList.Text) Then Exit Sub
    If Not m_anchorsDico.Item(vbProjList.Text).Exists(vbCompList.Text) Then Exit Sub
    
    Dim anchors As Dictionary
    Set anchors = m_anchorsDico.Item(vbProjList.Text).Item(vbCompList.Text)
    
    Dim line As Variant
    For Each line In anchors.Keys
        
        Dim prevLine As String, curLine As String
        
        Me.view.AddItem line - 1 & ":" & vbTab & anchors.Item(line).Item("prevLine")
        Me.view.AddItem line & ":" & vbTab & anchors.Item(line).Item("curLine")
        Me.view.AddItem ""
        
    Next line

End Sub

Private Sub vbCompList_Change()
    Call refreshView
End Sub

Private Sub vbProjList_Change()
    Call refreshVbCompList
End Sub

Public Function getAnchorsDico() As Dictionary
    
    Dim appAnchors As Dictionary
    Set appAnchors = New Dictionary
    
    Dim fso As New FileSystemObject, oStream As TextStream
    Set oStream = fso.OpenTextFile("C:\Users\a872364\Desktop\anchorsReport.txt", ForWriting, True)
    
    Dim vbp As Variant, vbProj As VBProject
    For Each vbp In Application.VBE.VBProjects
        
        Set vbProj = vbp
        
        Dim vbProjDico As Dictionary
        Set vbProjDico = New Dictionary
        
        If vbProj.Protection <> vbext_pp_locked Then
        
            Dim vbc As Variant, vbComp As VBComponent
            For Each vbc In vbProj.VBComponents
                    
                Set vbComp = vbc
                
                Dim vbCompDico As Dictionary
                Set vbCompDico = getvbCompDico(vbComp, vbProj.Name & ":" & vbComp.Name & ":", oStream)
                
                If vbCompDico.count > 0 And Not vbProjDico.Exists(vbComp.Name) Then vbProjDico.Add Key:=vbComp.Name, Item:=vbCompDico
                    
            Next vbc
            
            If vbProjDico.count > 0 And Not appAnchors.Exists(vbProj.Name) Then appAnchors.Add Key:=vbProj.Name, Item:=vbProjDico
        
        End If
        
    Next vbp
    
    Set getAnchorsDico = appAnchors

End Function

Private Function getvbCompDico(vbComp As VBComponent, parent As String, Optional oStream As TextStream) As Dictionary

    Dim vbCompDico As Dictionary
    Set vbCompDico = New Dictionary

    With vbComp.CodeModule
    
        Dim i As Integer
        For i = 1 To .CountOfLines
            
            Dim curLine As String, prevLine As String
            
            curLine = .Lines(i, 1)
            
            If i > 1 Then prevLine = .Lines(i - 1, 1)

            If isEndLine(curLine) Then
            
                Dim checkPrev As Boolean, checkCur As Boolean
                
                checkCur = InStr(1, curLine, "OOXOOXOOXOOXOOXOO", vbTextCompare) > 0
                checkPrev = InStr(1, prevLine, "OOXOOXOOXOOXOOXOO", vbTextCompare) > 0
                
                If Not checkCur And Not checkPrev Then
                
                    Dim anchorDico As Dictionary
                    Set anchorDico = New Dictionary
                    
                    anchorDico.Add Key:="prevLine", Item:=Trim(prevLine)
                    anchorDico.Add Key:="curLine", Item:=Trim(curLine)

                    If Not vbCompDico.Exists(i) Then vbCompDico.Add Key:=i, Item:=anchorDico
                    
                    If Not oStream Is Nothing Then oStream.WriteLine parent & i

                End If
            
            End If
    
        Next i
    
    End With
    
    Set getvbCompDico = vbCompDico

End Function

Private Function isEndLine(strLine As String) As Boolean

    Dim isEnd As Boolean

    isEnd = InStr(1, strLine, "Exit Sub", vbTextCompare) > 0
    isEnd = isEnd Or InStr(1, strLine, "Exit Function", vbTextCompare) > 0
    isEnd = isEnd Or InStr(1, strLine, "Exit Property", vbTextCompare) > 0
    isEnd = isEnd Or InStr(1, strLine, "End Sub", vbTextCompare) > 0
    isEnd = isEnd Or InStr(1, strLine, "End Function", vbTextCompare) > 0
    isEnd = isEnd Or InStr(1, strLine, "End Property", vbTextCompare) > 0
    
    isEndLine = isEnd

End Function

Sub testPane()

    With ThisWorkbook.VBProject.VBComponents.Item(20)
    
        .CodeModule.CodePane.TopLine = 584
    
    End With

End Sub

Private Function getSelectedVBComp() As VBComponent

    Dim vbProjName As String, vbCompName As String
    
    vbProjName = Me.vbProjList.Text
    vbCompName = Me.vbCompList.Text
    
    Dim vbp, vbc As Variant
    For Each vbp In Application.VBE.VBProjects
        If vbp.Name = vbProjName Then
            For Each vbc In vbp.VBComponents
                If vbc.Name = vbCompName Then
                    Set getSelectedVBComp = vbc
                    Exit Function
                End If
            Next
        End If
    Next
    
    Set getSelectedVBComp = Nothing

End Function

Private Sub view_Click()
    
    Dim strLine As String
    strLine = Me.view.List(Me.view.ListIndex)
    
    If strLine = "" Then Exit Sub
    
    Dim activeVbComp As VBComponent, activeLine As Long
    
    Set activeVbComp = getSelectedVBComp()
    activeLine = Val(strLine)
    
    Call activeVbComp.CodeModule.CodePane.SetSelection(activeLine, 1, activeLine + 1, 1)
    activeVbComp.CodeModule.CodePane.TopLine = IIf(activeLine > 15, activeLine - 10, activeLine)
    activeVbComp.CodeModule.CodePane.Show

End Sub
