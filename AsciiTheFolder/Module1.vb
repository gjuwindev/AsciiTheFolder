﻿Module Module1

    Public startingFolderName As String

    Public veryVerbose As Boolean = False
    Public verbose As Boolean = False

    Public totalItemCount, totalFolderCount As Integer
    Public renameCandidateCount, renameNotNeeded, renameSucceded, renameFailed, renameSkipped As Integer

    Sub Main(args() As String)

        If ProcessCommandLineArguments(args) Then
            WalkRecursivelyAndRenameFilesAndFolders(startingFolderName)
            DisplayCounters()
        End If

    End Sub

    Sub WalkRecursivelyAndRenameFilesAndFolders(upperFolderName As String, Optional folderDepth As Integer = 0)

        If folderDepth = 0 Then
            ResetCounters()
            totalFolderCount += 1
            totalItemCount += 1

            Dim upperFolderFolderName As String = System.IO.Path.GetDirectoryName(upperFolderName)
            Dim folderName As String = System.IO.Path.GetFileName(upperFolderName)
            System.IO.Directory.SetCurrentDirectory(upperFolderFolderName)
            upperFolderName = CheckRenameFileOrFolder(upperFolderFolderName, folderName)  ' new upper folder name, if renamed
        End If

        ' set current directory
        System.IO.Directory.SetCurrentDirectory(upperFolderName)

        ' get list of all entries in that directory
        Dim files() As String = System.IO.Directory.GetFileSystemEntries(".", "*.*", System.IO.SearchOption.TopDirectoryOnly)
        totalItemCount += files.Count

        ' check and rename all directories
        For Each fileName In files
            Dim oldName = fileName.Substring(2)
            CheckRenameFileOrFolder(upperFolderName, oldName)
        Next

        ' get folder names (some now renamed)
        Dim folders() As String = System.IO.Directory.GetDirectories(".", "*.*", System.IO.SearchOption.AllDirectories)
        totalFolderCount += folders.Count

        ' call ourselves recursively for each folder
        For Each folder In folders
            Dim nextFolderName = upperFolderName & "\" & folder.Substring(2)
            WalkRecursivelyAndRenameFilesAndFolders(nextFolderName, folderDepth + 1)
        Next

    End Sub

    ''' <summary>
    ''' Returns new starting folder name (or old, if unsuccessful)
    ''' </summary>
    ''' <param name="folderName"></param>
    ''' <param name="oldName"></param>
    ''' <returns></returns>
    Function CheckRenameFileOrFolder(folderName As String, oldName As String) As String

        If IsPureAscii(oldName) Then
            renameNotNeeded += 1
            If verbose Then
                Console.WriteLine("IsPureASCII: " & oldName)
            End If
            Return System.IO.Path.Combine(folderName, oldName)
        Else  ' not pure ASCII, should be ASCII-fied
            renameCandidateCount += 1
            Dim newName = ToPureAscii(oldName)

            Dim tempName As String = System.IO.Path.Combine(folderName, newName)
            While System.IO.File.Exists(tempName) OrElse System.IO.Directory.Exists(tempName)
                tempName = tempName & "X"
            End While
            newName = tempName

            If veryVerbose Then  ' just echo what would be done if -V flag was not present in the command line
                renameSkipped += 1
                Console.WriteLine(renameCandidateCount & ". shold be " & oldName & " => " & newName)
                Return System.IO.Path.Combine(folderName, oldName)
            Else
                Dim errorMessage As String = RenameFileOrFolder(folderName, oldName, newName)
                If errorMessage Is Nothing Then
                    renameSucceded += 1
                    Console.WriteLine(renameCandidateCount & ". " & oldName & " => " & newName)
                    Return System.IO.Path.Combine(folderName, newName)
                Else
                    renameFailed += 1
                    Console.WriteLine(renameCandidateCount & ". FAILED " & oldName & " => " & newName)
                    Console.WriteLine("     " & errorMessage)
                    Return System.IO.Path.Combine(folderName, oldName)
                End If
            End If
        End If

    End Function

    Sub ResetCounters()

        totalItemCount = 0
        totalFolderCount = 0

        renameCandidateCount = 0
        renameNotNeeded = 0
        renameSucceded = 0
        renameFailed = 0
        renameSkipped = 0

    End Sub

    Sub DisplayCounters()

        Console.WriteLine()
        Console.WriteLine("Total item count:      " & totalItemCount)
        Console.WriteLine("Total folder count:    " & totalFolderCount)
        Console.WriteLine("Total file count:      " & totalItemCount - totalFolderCount)
        Console.WriteLine()
        Console.WriteLine("Total ASCII names:     " & renameNotNeeded)
        Console.WriteLine("Total nonASCII names:  " & renameCandidateCount)
        Console.WriteLine("Total renamed:         " & renameSucceded)
        Console.WriteLine("Total renames failed:  " & renameFailed)
        Console.WriteLine("Total renames skipped: " & renameSkipped)

    End Sub

End Module
