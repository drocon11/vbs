With CreateObject("Scripting.FileSystemObject")
    For Each objFile In .GetFile(WScript.ScriptFullName).ParentFolder.Files
        If .FileExists(objFile.Path) Then
            If Now() > DateAdd("d", 7, objFile.DateCreated) Then
                strTextFile = objFile.Path & "." & .GetBaseName(WScript.ScriptName)
                If .FileExists(strTextFile) Then
                    strLine = ""
                    With .OpenTextFile(strTextFile)
                        strLine = .ReadLine
                        .Close
                    End With
                    If strLine<>"" Then
                        strDstFolder = .BuildPath(objFile.ParentFolder.Path, strLine)
                        If Not .FolderExists(strDstFolder) Then
                            .CreateFolder(strDstFolder)
                        End If
                        strDstFile = .BuildPath(strDstFolder, objFile.Name)
                        If Not .FileExists(strDstFile) Then
                            .MoveFile objFile.Path, strDstFile
                            .DeleteFile strTextFile
                        End If
                    End If
                End If
            End If
        End If
    Next
End With
