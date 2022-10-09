Attribute VB_Name = "OL15_VBE"
Option Explicit

Private Const MODULE_NAME As String = "OL15_GIT"

Private Const EXCEL_OBJECTS_FOLDER_NAME As String = "01_excel_object"
Private Const MSFORMS_FOLDER_NAME       As String = "02_userforms"
Private Const STD_MODULES_FOLDER_NAME   As String = "03_modules"
Private Const CLASS_MODULES_FOLDER_NAME As String = "04_classes"

Public Function printRequirements(wb As Workbook) As Boolean
    ' ----------------------------------------------------------------
    ' Purpose: Enregistrer automatiquement toute les références dans le fichier requirement.txt, au meme format que l'export de références MZTool
    ' Parameter wb (Workbook): Classeur dont on souhaite enregistrer les références
    ' Author: a872364
    ' Date: 30/09/2022
    ' ----------------------------------------------------------------
    Call OOXOOXOOXOOXOOXOO(MODULE_NAME, "printRequirements")

    'On récupère le fichier 'requirement.txt' situé à  l'emplacement du workbook à traiter / on le créer s'il n'existe pas
    Dim fso As New FileSystemObject, oStream As TextStream
    Set oStream = fso.OpenTextFile(wb.path & "\" & "requirement.txt", ForWriting, True)

    'On vient lire toute ses références, s'il y en a une qui pointe vers un fichier excel, on enregistre son nom, son full path et sa version
    Dim ref As Variant, brokenRef As String, versionIndex As Integer
    For Each ref In wb.VBProject.References

        If ref.IsBroken Then
            'Si la référence est manquante, on la stocke dans un string que l'on restituera à la fin de la fonction
            brokenRef = brokenRef & ref.Name & vbCrLf
        Else
            
            Dim refAttr(5) As String
            
            refAttr(0) = ref.Name
            refAttr(1) = ref.Description
            refAttr(2) = ref.GUID
            refAttr(5) = ref.FullPath
            If InStr(1, ref.FullPath, ".xl") Then
                If versionIndex = 0 Then
checkVersionIndex:
                    Call printAvailableFileDetails(ref.FullPath)
                    versionIndex = InputBox("Give the index of the file version", "Property index", 134)
                End If
                'Si la ref est un fichier excel, on vient remplir son numéro de version manuellement
                Dim xlVersion As String
                xlVersion = getFileDetail(ref.FullPath, versionIndex)
                If xlVersion = "" Then GoTo checkVersionIndex
                refAttr(3) = Replace(Split(xlVersion, ".")(0), "v", "", 1, 1, vbTextCompare)
                refAttr(4) = Split(xlVersion, ".")(1)
            Else
                'sinon on prend les attributs du vbComponent
                refAttr(3) = ref.Major
                refAttr(4) = ref.Minor
    
            End If
    
            oStream.WriteLine Join(refAttr, "|")
        
        End If

    Next
    
    'On retourne true si toute les références ont été trouvés / false si au moins une référence est manquante
    If brokenRef <> "" Then
        printRequirements = False
        MsgBox "Les références suivantes sont introuvables et ne seront donc pas ajoutés à  requirement.txt : " & vbCrLf & vbCrLf & brokenRef, vbOKOnly Or vbExclamation, "Références manquante"
    Else
        printRequirements = True
    End If
    
    oStream.Close
    
    Call OOXOOXOOXOOXOOXOO("Fin")
End Function

Public Sub exportSourceCode(workbookToCommit As Workbook)
    ' ----------------------------------------------------------------
    ' Purpose: Préparer le système de fichier pour réaliser un commit git
    '   1) Sauvegarde le code du repo actuel (s'il existe un dossier src à  l'emplacement du workbook) vers le roaming VBA
    '   2) Exporte tout le code source d'un classeur dans un système de fichiers standardisé
    '   3) Exporte toute les références externe du fichier avec leur numéro de version vers requirement.txt
    '   4) Sauvegarde le classeur
    '   5) Ouvre l'exporateur de fichier à  l'emplacement du workbook
    ' Parameter workbookToCommit (Workbook): Classeur dont on souhaite exporter le code source
    ' Parameter exportRequirements (Boolean): permet de ne pas imprimer le fichier requirement.txt
    ' Author: a872364
    ' Date: 29/09/2022
    ' ----------------------------------------------------------------
    Call OOXOOXOOXOOXOOXOO(MODULE_NAME, "exportSourceCode")

    Dim fso As New FileSystemObject
    
    Dim sourcePath As String, sourceFolder As Folder
    sourcePath = workbookToCommit.path & "\" & "src"

    If Not fso.FolderExists(sourcePath) Then
        'Création du dossier 'src'
        
        Set sourceFolder = fso.CreateFolder(sourcePath)

    Else
        'Archivage du depot actuel sur le disque utilisateur en cas de problème
        
        Set sourceFolder = fso.GetFolder(sourcePath)

        Dim subF As Variant
        For Each subF In sourceFolder.SubFolders
            subF.Delete
        Next

    End If

    'Creation des sous-dossiers et export du code
    With sourceFolder.SubFolders
    
        Dim SheetFolder As Folder
        Set SheetFolder = .Add(EXCEL_OBJECTS_FOLDER_NAME)
        Call exportSourceFiles(workbookToCommit, vbext_ct_Document, SheetFolder)
        
        Dim userfomrFolder As Folder
        Set userfomrFolder = .Add(MSFORMS_FOLDER_NAME)
        Call exportSourceFiles(workbookToCommit, vbext_ct_MSForm, userfomrFolder)
        
        Dim modulesFolder As Folder
        Set modulesFolder = .Add(STD_MODULES_FOLDER_NAME)
        Call exportSourceFiles(workbookToCommit, vbext_ct_StdModule, modulesFolder)
        
        Dim classFolder As Folder
        Set classFolder = .Add(CLASS_MODULES_FOLDER_NAME)
        Call exportSourceFiles(workbookToCommit, vbext_ct_ClassModule, classFolder)
    
    End With
    
    'Mise à  jour des références du classeur
    Call printRequirements(workbookToCommit)

    'Sauvegarde le classeur et ouvre le dépot git
    workbookToCommit.Save
    Shell "C:\windows\explorer.exe " & workbookToCommit.path, vbNormalFocus
    
    Call OOXOOXOOXOOXOOXOO("Fin")
    Exit Sub
    
theEnd:
    MsgBox "Erreur", vbOKOnly Or vbCritical, Application.Name
    Call OOXOOXOOXOOXOOXOO("Fin")
End Sub

Private Sub exportSourceFiles(wb As Workbook, fileType As VBIDE.vbext_ComponentType, destFolder As Folder)
    ' ----------------------------------------------------------------
    ' Purpose: Exporte tout les code source d'une catégorie vers un repertoire, puis encode les fichiers en utf-8
    ' Parameter wb (Workbook): Classeur dont on souhaite exporter le code
    ' Parameter fileType (vbext_ComponentType): Type de code que l'on souhaite exporter (modules, classes, userform, objets excel)
    ' Parameter destFolder (Folder): Repertoire de destination, oà¹ seront enregistrés les codes source
    ' Author: a872364
    ' Date: 03/10/2022
    ' ----------------------------------------------------------------
    Call OOXOOXOOXOOXOOXOO(MODULE_NAME, "exportSourceFiles")

    Dim vbComp As Object, extension As String
    For Each vbComp In wb.VBProject.VBComponents
    
        If vbComp.Type = fileType Then
    
            Select Case vbComp.Type
        
                Case vbext_ct_Document      'excel objects
                    extension = ".cls"

                Case vbext_ct_MSForm        'Userform
                    extension = ".frm"

                Case vbext_ct_StdModule     'Module
                    extension = ".bas"

                Case vbext_ct_ClassModule   'Classe
                    extension = ".cls"

            End Select
            
            'On supprime toute les lignes vides pour éviter les detections de modif de fichier inutiles
            Call deleteEndBlankLines(vbComp.CodeModule)
            
            'Export du vbComponent vers le repertoire spécifié en entrée de fonction
            Dim expFileName As String, fso As New FileSystemObject
            expFileName = destFolder.path & "\" & vbComp.Name & extension
            vbComp.Export (expFileName)
            
            'Conversion de l'encodage du ficher exporté de unicode->utf-8
            Call convertFileToUtf(fso.GetFile(expFileName))
            
        End If
        
    Next
    
    Call OOXOOXOOXOOXOOXOO("Fin")
End Sub

'TODO: passer cette méthode en publique une fois qu'elle sera fonctionelle
Public Sub importSourceCode(wbToImport As Workbook)
    ' ----------------------------------------------------------------
    ' Purpose:  Import de tout les fichiers d'un repertoire santardisé, vers le vbProject d'un classeur /
    '           suppression de tout les modules qui ne sont plus présents dans le dossier source
    ' Parameter wbToImport (Workbook):  Classeur vers lequel on va faire l'import (utilisé pour determiner
    '                                   l'emplacement des fichiers à  importer)
    ' Author: a872364
    ' Date: 03/10/2022
    ' ----------------------------------------------------------------
    
    'On verifie qu'on ne fait pas d'import dans l'OpenLibrary
    If InStr(1, wbToImport.Name, "OpenLibrary", vbTextCompare) > 0 Then
        MsgBox "Import de code interdit dans l'openLibrary.", vbOKOnly Or vbExclamation, "OpenLibrary.importSourceCode"
        Exit Sub
    End If
    
    'Lecture de tout les composants VBA du projets dans lequel on fait l'import
    Dim dicoVBComp As Dictionary
    Set dicoVBComp = readVBComponents(wbToImport)
    
    'Copie de tout les fichiers vers un dossier temporaire
    Dim tempFolder As Folder, fso As New FileSystemObject
    Set tempFolder = fso.CreateFolder(getVBAFolder().path & "\src_" & wbToImport.Name & Format(Now, "_yyyymmddhhmmss"))

    Dim sourceFolder As Folder
    Set sourceFolder = fso.GetFolder(wbToImport.path & "\src")
    sourceFolder.Copy (tempFolder.path)

    'Lecture de tout les fichiers temporaires
    Dim dicoExpFiles As Dictionary
    Set dicoExpFiles = readExportedFiles(tempFolder)

    Dim fileKey As Variant, errFiles As String
    For Each fileKey In dicoExpFiles.Keys
    
        On Error GoTo nextFile
        
        Dim mFile As File
        Set mFile = dicoExpFiles.Item(fileKey)

        If Not dicoVBComp.Exists(fileKey) Then
            'Si le fichier n'existe pas dans le vbProject, on l'importe depuis le dossier source
            Call wbToImport.VBProject.VBComponents.Import(mFile.path)
        
        Else
            'Si le fichier existe dans le vbProject et dans le repertoire source, on supprime le vbComponent associé, puis on le ré-importe via le fichier source
            Dim mVBComp As VBComponent
            Set mVBComp = dicoVBComp.Item(fileKey)
            
            'Modification de l'encodage du fichier utf8->unicode
            Call convertFileToUnicode(mFile)
            
            'Si le composant est un excelObjet, on va se contenter de remplacer tout le code de son module
            If mVBComp.Type = vbext_ct_Document Then
            
                'Suppression de toute les lignes du module
                Call mVBComp.CodeModule.DeleteLines(1, mVBComp.CodeModule.CountOfLines)
                
                'Ré-écriture de toute les lignes du module à partir du fichier source
                Call mVBComp.CodeModule.AddFromString(mFile.OpenAsTextStream.ReadAll)
                
                'Suppression des lignes de déclaration présent dans les fichiers exportés par VBA
                Call deleteDeclarationLines(mVBComp)
                
            Else
                'Suppression du module
                Call wbToImport.VBProject.VBComponents.Remove(mVBComp)
                
                'Ré-import du module + Modification de l'item du dictionnaire des vbCOmponents
                Set dicoVBComp.Item(fileKey) = wbToImport.VBProject.VBComponents.Import(mFile.path)
                
            End If

        End If
        
nextFile:
        
        'Si une erreur est survenue dans l'import, on le notifie au développeur
        If err.Number <> 0 Then
            errFiles = errFiles & mFile.Name & vbCrLf
            err.Clear
        End If
        
    Next

    On Error GoTo 0
    
    'Suppression du dossier temporaire contenant les fichiers encodés en unicode
    tempFolder.Delete

    Dim sheetsToDelete As String
    
    'Une fois que tout a été importé, on vient supprimer tout les modules qui sont absents du dossier source

    For Each fileKey In dicoVBComp.Keys

        Set mVBComp = dicoVBComp.Item(fileKey)

        If Not dicoExpFiles.Exists(fileKey) Then

            If mVBComp.Type <> vbext_ct_Document Then
                'Si le composant n'est pas un excel object(feuille, classeur), on le supprime
                Call wbToImport.VBProject.VBComponents.Remove(mVBComp)
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
    
    'Si au moins un fichier a rencontré une erreur, on le notifie au développeur
    If errFiles <> "" Then _
        MsgBox "Attention, les fichiers suivant ont rencontré un problème lors de l'import : " & vbCrLf & vbCrLf & sheetsToDelete, _
        vbOKOnly Or vbCritical, "OpenLibray.importSourceCode"
    
    'Si un objet excel n'existe plus dans le dossier source, on le notifie au developpeur
    If sheetsToDelete <> "" Then
        MsgBox "Attention, les modules suivants n'existent plus et n'ont pas pu être supprimé : " & vbCrLf & vbCrLf & sheetsToDelete, _
        vbOKOnly Or vbExclamation, "OpenLibray.importSourceCode"
    Else
        MsgBox "Import successfull.", vbOKOnly Or vbInformation, "OpenLibray.importSourceCode"
    End If

End Sub

Private Function readVBComponents(wb As Workbook) As Dictionary
    ' ----------------------------------------------------------------
    ' Purpose: Renvoyer un dictionnaire contenant tout les vbComponents d'un vbProject, avec leur noms en clé
    ' Parameter wb (Workbook): Classeur contenant le vbProject dont on veut extraire les vbComponents
    ' Return Type: Dictionary
    ' Author: a872364
    ' Date: 03/10/2022
    ' ----------------------------------------------------------------
    Dim dicoVBComp As Dictionary
    Set dicoVBComp = New Dictionary
    
    Dim vbc As Variant, vbComp As VBIDE.VBComponent
    For Each vbc In wb.VBProject.VBComponents
        Set vbComp = vbc
        dicoVBComp.Add Key:=vbComp.Name, Item:=vbComp
    Next
    
    Set readVBComponents = dicoVBComp

End Function

Private Function readExportedFiles(sourceFolder As Folder) As Dictionary
    ' ----------------------------------------------------------------
    ' Purpose: Crée un dictionnaire contenant tout les fichiers d'un depot git standardisé pour VBA (clé = nom de fichier, item = objet File du fichier)
    ' Parameter wb (Workbook): Classeur dont on souhaite lire les fichiers sources (sert de base pour trouver les emplacements de fichiers)
    ' Return Type: Dictionary
    ' Author: a872364
    ' Date: 03/10/2022
    ' ----------------------------------------------------------------
    Dim fso As New FileSystemObject
    
    Dim dicoExpFiles As Dictionary
    
    Call addFilesIn(dico:=dicoExpFiles, from:=sourceFolder.SubFolders(EXCEL_OBJECTS_FOLDER_NAME))
    Call addFilesIn(dico:=dicoExpFiles, from:=sourceFolder.SubFolders(MSFORMS_FOLDER_NAME))
    Call addFilesIn(dico:=dicoExpFiles, from:=sourceFolder.SubFolders(STD_MODULES_FOLDER_NAME))
    Call addFilesIn(dico:=dicoExpFiles, from:=sourceFolder.SubFolders(CLASS_MODULES_FOLDER_NAME))

    Set readExportedFiles = dicoExpFiles

End Function

Private Sub addFilesIn(ByRef dico As Dictionary, ByVal from As Folder)
    ' ----------------------------------------------------------------
    ' Purpose: Remplir un dico d'objets 'File' si leur extension est .frm, .bas ou .cls
    ' Parameter dicoFiles (Dictionary): dictionnaire passé par référence, que l'on va remplir successivement d'objets 'File'
    ' Parameter mFolder (Folder): répertoire dans lequel on vient lire les fichiers à  ajouter au dictionnaire
    ' Author: a872364
    ' Date: 03/10/2022
    ' ----------------------------------------------------------------
    If dico Is Nothing Then Set dico = New Dictionary

    Dim mFile As Variant
    For Each mFile In from.Files

        Dim fileExtension As String, fileName As String
        fileName = Split(mFile.Name, ".")(0)
        fileExtension = Split(mFile.Name, ".")(1)
        
        If fileExtension = "frm" Or fileExtension = "cls" Or fileExtension = "bas" Then
            dico.Add Key:=fileName, Item:=mFile
        End If
        
    Next

End Sub

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
    Wend
    
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
    Wend
    
End Sub

Private Sub deleteDeclarationLines(vbComp As VBComponent)
    ' ----------------------------------------------------------------
    ' Purpose: Supprimer toute les lignes d'attributs en début de module (generés automatiquement à l'xport d'un fichier de code VBA)
    ' Parameter vbcomp (VBComponent): vbComponent sur lequel on veut faire un nettoyage des lignes de déclaration
    ' Author: a872364
    ' Date: 05/10/2022
    ' ----------------------------------------------------------------
    Dim mModule As CodeModule
    Set mModule = vbComp.CodeModule
    
    If mModule.CountOfLines <= 0 Then Exit Sub

    Dim i As Integer, lastLine As Integer
    For i = 1 To mModule.CountOfDeclarationLines

        If startsWith(mModule.Lines(i, 1), "Attribute VB_") Then lastLine = i

    Next
    
    Debug.Print lastLine, vbComp.Name
    
    If lastLine > 0 Then Call mModule.DeleteLines(1, lastLine)

End Sub

Public Sub inspectCode()
    OLF_CHECK_HISTO.Show 0
End Sub
