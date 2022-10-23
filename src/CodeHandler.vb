Option Explicit On

Imports System.Windows.Forms
Imports Extensibility
Imports Microsoft.Vbe.Interop
Imports Microsoft.Office.Core
Imports System.IO

Public Class CodeHandler

    Private m_VBE As VBE

    Private Const MODULE_NAME As String = "OL15_GIT"

    Private Const EXCEL_OBJECTS_FOLDER_NAME As String = "01_excel_object"
    Private Const MSFORMS_FOLDER_NAME As String = "02_userforms"
    Private Const STD_MODULES_FOLDER_NAME As String = "03_modules"
    Private Const CLASS_MODULES_FOLDER_NAME As String = "04_classes"

    Public Sub New(_VBE As VBE)
        m_VBE = _VBE
    End Sub

    Public Function printRequirements(vbProj As VBProject) As Boolean
        ' ----------------------------------------------------------------
        ' Purpose: Enregistrer automatiquement toute les références dans le fichier requirement.txt, au meme format que l'export de références MZTool
        ' Parameter wb (Workbook): Classeur dont on souhaite enregistrer les références
        ' Author: a872364
        ' Date: 30/09/2022
        ' ----------------------------------------------------------------

        'On récupère le fichier 'requirement.txt' situé à  l'emplacement du workbook à traiter / on le créer s'il n'existe pas
        Dim sWriter As StreamWriter = getRequirementTextFile(vbProj, True)

        'On vient lire toute ses références, s'il y en a une qui pointe vers un fichier excel, on enregistre son nom, son full path et sa version
        Dim ref As Reference, brokenRef As String = ""
        For Each ref In vbProj.References

            If ref.IsBroken Then
                'Si la référence est manquante, on la stocke dans un string que l'on restituera à la fin de la fonction
                brokenRef = brokenRef & ref.Name & vbCrLf
            Else

                Dim refAttr(5) As String

                refAttr(0) = ref.Name
                refAttr(1) = ref.Description
                refAttr(2) = ref.Guid
                refAttr(5) = ref.FullPath

                If InStr(1, ref.FullPath, ".xl") Then

                    Dim xlVersion As String = getFileVersion(New FileInfo(vbProj.FileName))
                    refAttr(3) = Replace(Split(xlVersion, ".")(0), "v", "", 1, 1, vbTextCompare)
                    refAttr(4) = Split(xlVersion, ".")(1)
                Else
                    'sinon on prend les attributs du vbComponent
                    refAttr(3) = ref.Major
                    refAttr(4) = ref.Minor

                End If

                sWriter.WriteLine(Strings.Join(refAttr, "|"))

            End If

        Next

        sWriter.Close()

        'On retourne true si toute les références ont été trouvés / false si au moins une référence est manquante
        If brokenRef <> "" Then
            printRequirements = False
            MessageBox.Show(text:="Les références suivantes sont introuvables et ne seront donc pas ajoutés à  requirement.txt : " & vbCrLf & vbCrLf & brokenRef, caption:=vbOKOnly Or vbExclamation) ', vbOKOnly Or vbExclamation, "Références manquante"
        Else
            printRequirements = True
        End If

        'oStream.Close

    End Function

    Public Sub exportSourceCode(vbProject As VBProject)

        Dim vbProjFileInfo As New IO.FileInfo(vbProject.FileName)
        Dim sourceFolder As DirectoryInfo = getVBProjectSourceFolder(vbProject, True)

        For Each subF As DirectoryInfo In sourceFolder.GetDirectories()
            Directory.Delete(subF.FullName, True)
        Next

        'Creation des sous-dossiers et export du code
        With sourceFolder

            Dim SheetFolder As DirectoryInfo = .CreateSubdirectory(EXCEL_OBJECTS_FOLDER_NAME)
            Call exportSourceFiles(vbProject, vbext_ComponentType.vbext_ct_Document, SheetFolder)

            Dim userfomrFolder As DirectoryInfo = .CreateSubdirectory(MSFORMS_FOLDER_NAME)
            Call exportSourceFiles(vbProject, vbext_ComponentType.vbext_ct_MSForm, userfomrFolder)

            Dim modulesFolder As DirectoryInfo = .CreateSubdirectory(STD_MODULES_FOLDER_NAME)
            Call exportSourceFiles(vbProject, vbext_ComponentType.vbext_ct_StdModule, modulesFolder)

            Dim classFolder As DirectoryInfo = .CreateSubdirectory(CLASS_MODULES_FOLDER_NAME)
            Call exportSourceFiles(vbProject, vbext_ComponentType.vbext_ct_ClassModule, classFolder)

        End With

        'Mise à  jour des références du classeur
        Call printRequirements(vbProj:=vbProject)

        'Sauvegarde le classeur et ouvre le dépot git
        'TODO: trouver un moyen de sauvegarder le vbProject
        'vbProject.SaveAs(vbProject.FileName)

        If MessageBox.Show("Open export folder ?", "Export succeded", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbYes Then
            Shell("C:\windows\explorer.exe " & vbProjFileInfo.DirectoryName, vbNormalFocus)
        End If

    End Sub

    Private Sub exportSourceFiles(vbProj As VBProject, fileType As vbext_ComponentType, destFolder As DirectoryInfo)

        Dim extension As String = ""
        For Each vbComp As VBComponent In vbProj.VBComponents

            If vbComp.Type = fileType Then

                Select Case vbComp.Type

                    Case vbext_ComponentType.vbext_ct_Document      'excel objects
                        extension = ".cls"

                    Case vbext_ComponentType.vbext_ct_MSForm        'Userform
                        extension = ".frm"

                    Case vbext_ComponentType.vbext_ct_StdModule     'Module
                        extension = ".bas"

                    Case vbext_ComponentType.vbext_ct_ClassModule   'Classe
                        extension = ".cls"

                End Select

                'On supprime toute les lignes vides pour éviter les detections de modif de fichier inutiles
                Call deleteEndBlankLines(vbComp.CodeModule) 'TODO: supprimer les lignes vides à l'export

                'Export du vbComponent vers le repertoire spécifié en entrée de fonction
                Dim expFileName As String = destFolder.FullName & "\" & vbComp.Name & extension
                vbComp.Export(expFileName)

                'Conversion de l'encodage du ficher exporté de unicode->utf-8
                Call convertFileToUtf(expFileName)

            End If
        Next

    End Sub

    Public Sub importSourceCode(vbProj As VBProject)
        ' ----------------------------------------------------------------
        ' Purpose:  Import de tout les fichiers d'un repertoire santardisé, vers le vbProject d'un classeur /
        '           suppression de tout les modules qui ne sont plus présents dans le dossier source
        ' Parameter wbToImport (Workbook):  Classeur vers lequel on va faire l'import (utilisé pour determiner
        '                                   l'emplacement des fichiers à  importer)
        ' Author: a872364
        ' Date: 03/10/2022
        ' ----------------------------------------------------------------

        'Lecture de tout les composants VBA du projets dans lequel on fait l'import
        Dim dicoVBComp As Dictionary(Of String, VBComponent) = readVBComponents(vbProj)

        'Copie de tout les fichiers vers un dossier temporaire
        Dim tempFolder As DirectoryInfo = Directory.CreateDirectory(Path.Combine(getVBAFolder().FullName, "src_") & vbProj.Name & Format(Now, "_yyyymmddhhmmss"))
        Dim sourceFolder As DirectoryInfo = getVBProjectSourceFolder(vbProj:=vbProj, create:=False)

        My.Computer.FileSystem.CopyDirectory(sourceFolder.FullName, tempFolder.FullName)

        'Lecture de tout les fichiers temporaires
        Dim dicoExpFiles As Dictionary(Of String, FileInfo) = getDicoFilesFromDirectory(tempFolder, False)
        Dim errFiles As String = ""

        For Each fileKey As String In dicoExpFiles.Keys

            Dim mFile As FileInfo = dicoExpFiles.Item(fileKey)

            'Modification de l'encodage du fichier temporaire utf8->unicode
            Call convertFileToUnicode(mFile)

            Try

                If Not dicoVBComp.ContainsKey(fileKey) Then
                    'Si le fichier n'existe pas dans le vbProject, on l'importe depuis le dossier source
                    Call vbProj.VBComponents.Import(mFile.FullName)

                Else

                    'Si le fichier existe dans le vbProject et dans le repertoire source, on supprime le vbComponent associé, puis on le ré-importe via le fichier source
                    Dim mVBComp As VBComponent = dicoVBComp.Item(fileKey)

                    'Si le composant est un excelObjet, on va se contenter de remplacer tout le code de son module
                    If mVBComp.Type = vbext_ComponentType.vbext_ct_Document Then

                        'Suppression de toute les lignes du module
                        Call mVBComp.CodeModule.DeleteLines(1, mVBComp.CodeModule.CountOfLines)

                        'Ré-écriture de toute les lignes du module à partir du fichier source
                        Using sr As New StreamReader(mFile.FullName)
                            Call mVBComp.CodeModule.AddFromString(sr.ReadToEnd)
                            sr.Close()
                        End Using

                        'Suppression des lignes de déclaration présent dans les fichiers exportés par VBA
                        Call deleteDeclarationLines(mVBComp)

                    Else
                        'Suppression du module
                        vbProj.VBComponents.Remove(mVBComp)

                        'Ré-import du module + Modification de l'item du dictionnaire des vbCOmponents
                        dicoVBComp.Item(fileKey) = vbProj.VBComponents.Import(mFile.FullName)

                    End If

                End If

            Catch ex As Exception
                errFiles = errFiles & mFile.Name & vbCrLf
            End Try

        Next

        'Suppression du dossier temporaire contenant les fichiers encodés en unicode
        tempFolder.Delete(recursive:=True)

        Dim sheetsToDelete As String = ""

        'Une fois que tout a été importé, on vient supprimer tout les modules qui sont absents du dossier source

        For Each fileKey In dicoVBComp.Keys

            Dim mVBComp As VBComponent = dicoVBComp.Item(fileKey)

            If Not dicoExpFiles.ContainsKey(fileKey) Then

                If mVBComp.Type <> vbext_ComponentType.vbext_ct_Document Then
                    'Si le composant n'est pas un excel object(feuille, classeur), on le supprime
                    vbProj.VBComponents.Remove(mVBComp)
                Else
                    'Si le composant est un excel object, on ne peut pas le supprimer via VBE, on le notifiera au developpeur plus bas
                    sheetsToDelete = sheetsToDelete & mVBComp.Name & vbCrLf
                End If

            ElseIf mVBComp.CodeModule.CountOfLines > 0 Then

                'Suppression des lignes vides en début de module (pour éviter les detections de modif inutiles)
                Call deleteStartBlankLines(mVBComp.CodeModule)

                'Suppression des lignes vides en fin de module (pour éviter les detections de modif inutiles)
                Call deleteEndBlankLines(mVBComp.CodeModule)

            End If
        Next

        If errFiles <> "" Then
            'Si au moins un fichier a rencontré une erreur, on le notifie au développeur
            MessageBox.Show(text:="Attention, les fichiers suivant ont rencontré un problème lors de l'import : " & vbCrLf & vbCrLf & errFiles,
                            caption:="VBEAddin.importSourceCode",
                            buttons:=MessageBoxButtons.OK,
                            icon:=MessageBoxIcon.Error)

        ElseIf sheetsToDelete <> "" Then
            'Si un objet excel n'existe plus dans le dossier source, on le notifie au developpeur
            MessageBox.Show(text:="Attention, les modules suivants n'existent plus et n'ont pas pu être supprimé : " & vbCrLf & vbCrLf & sheetsToDelete,
                            caption:="VBEAddin.importSourceCode",
                            buttons:=MessageBoxButtons.OK,
                            icon:=MessageBoxIcon.Warning)
        Else
            MessageBox.Show(text:="Import succeded",
                            caption:="VBEAddin.importSourceCode",
                            buttons:=MessageBoxButtons.OK,
                            icon:=MessageBoxIcon.Information)
        End If

    End Sub

    Private Function readVBComponents(vbProj As VBProject) As Dictionary(Of String, VBComponent)
        ' ----------------------------------------------------------------
        ' Purpose: Renvoyer un dictionnaire contenant tout les vbComponents d'un vbProject, avec leur noms en clé
        ' Parameter wb (Workbook): Classeur contenant le vbProject dont on veut extraire les vbComponents
        ' Return Type: Dictionary
        ' Author: a872364
        ' Date: 03/10/2022
        ' ----------------------------------------------------------------
        Dim dicoVBComp As New Dictionary(Of String, VBComponent)

        Dim vbComp As VBComponent
        For Each vbComp In vbProj.VBComponents
            dicoVBComp.Add(key:=vbComp.Name, value:=vbComp)
        Next

        Return dicoVBComp

    End Function

    Private Sub deleteStartBlankLines(mModule As CodeModule)
        ' ----------------------------------------------------------------
        ' Purpose: Supprimer toute les lignes vide présentes au début d'un module.
        ' Parameter mModule (CodeModule): L'objet codeModule (VBIDE.vbComponent) dont on souhaite supprimer les lignes vides
        ' Author: a872364
        ' Date: 04/10/2022
        ' ----------------------------------------------------------------
        If mModule.CountOfLines <= 0 Then Exit Sub

        Dim firstLineContent As String
        firstLineContent = mModule.Lines(1, 1)

        While firstLineContent = "" And mModule.CountOfLines > 1
            Call mModule.DeleteLines(1, 1)
            firstLineContent = mModule.Lines(1, 1)
        End While

    End Sub

    Private Sub deleteEndBlankLines(mModule As CodeModule)
        ' ----------------------------------------------------------------
        ' Purpose: Supprimer toute les lignes vide présentes à  la fin d'un module.
        ' Parameter mModule (CodeModule): L'objet codeModule (VBIDE.vbComponent) dont on souhaite supprimer les lignes vides
        ' Author: a872364
        ' Date: 04/10/2022
        ' ----------------------------------------------------------------
        If mModule.CountOfLines <= 0 Then Exit Sub

        Dim lastLineContent As String
        lastLineContent = mModule.Lines(mModule.CountOfLines, 1)

        While lastLineContent = "" And mModule.CountOfLines > 1
            Call mModule.DeleteLines(mModule.CountOfLines, 1)
            lastLineContent = mModule.Lines(mModule.CountOfLines, 1)
        End While

    End Sub

    Private Sub deleteDeclarationLines(vbComp As VBComponent)
        ' ----------------------------------------------------------------
        ' Purpose: Supprimer toute les lignes d'attributs en début de module (generés automatiquement à l'xport d'un fichier de code VBA)
        ' Parameter vbcomp (VBComponent): vbComponent sur lequel on veut faire un nettoyage des lignes de déclaration
        ' Author: a872364
        ' Date: 05/10/2022
        ' ----------------------------------------------------------------
        Dim mModule As CodeModule
        mModule = vbComp.CodeModule

        If mModule.CountOfLines <= 0 Then Exit Sub

        Dim i As Integer, lastLine As Integer
        For i = 1 To mModule.CountOfDeclarationLines

            If mModule.Lines(i, 1).StartsWith("Attribute VB_") Then lastLine = i

        Next

        If lastLine > 0 Then Call mModule.DeleteLines(1, lastLine)

    End Sub

End Class
