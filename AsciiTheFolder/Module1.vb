Module Module1

    Sub Main(args() As String)

        Dim showPureAsciiFilesToo As Boolean = False
        Dim verbose As Boolean = False

        Dim nonAsciiFileNameCount As Integer = 0

        Dim argList As New List(Of String)(args)

        If argList.Count = 2 Then
            If argList(0) = "-v" Then
                verbose = True
                Console.WriteLine("Verbose mode ON.")
                argList.RemoveAt(0)
            ElseIf argList(0) = "-V" Then
                verbose = True
                showPureAsciiFilesToo = True
                Console.WriteLine("Very verbose mode ON.")
                argList.RemoveAt(0)
            End If
        End If

        If argList.Count = 1 Then
            Dim folderName = argList(0)

            If System.IO.Directory.Exists(folderName) = False Then
                Console.WriteLine("Folder: """ & folderName & """ does not exist.")
            Else
                Console.WriteLine("Processing folder """ & folderName & """.")

                If IsPureAscii(folderName) = False Then
                    Console.WriteLine("Folder name itself contains non-ASCI characters.")
                    Dim newName = ToPureAscii(folderName)
                    If Not verbose Then
                        Console.WriteLine("Folder " & folderName & " => " & newName)
                        System.IO.Directory.Move(folderName, newName)
                        folderName = newName
                    Else
                        Console.WriteLine("Should be " & folderName & " => " & newName)
                    End If
                End If

                Dim files() As String = System.IO.Directory.GetFileSystemEntries(folderName, "*.*", System.IO.SearchOption.AllDirectories)
                Dim renamedOK, renameFailed As Integer

                System.IO.Directory.SetCurrentDirectory(folderName)
                For Each file In files
                    Dim oldName = file.Substring(folderName.Length + 1)
                    Dim newName = ToPureAscii(oldName)
                    If oldName <> newName Then
                        nonAsciiFileNameCount += 1
                        If Not verbose Then
                            Console.WriteLine(nonAsciiFileNameCount & ". " & oldName & " => " & newName)
                            If RenameFile(oldName, newName) Then
                                renamedOK += 1
                            Else
                                renameFailed += 1
                            End If
                        Else
                            Console.WriteLine(nonAsciiFileNameCount & ". shold be " & oldName & " => " & newName)
                        End If
                    ElseIf showPureAsciiFilesToo Then
                        Console.WriteLine("pureASCII: " & oldName)
                    End If
                Next

                If files.Count = 0 Then
                    Console.WriteLine("Folder """ & folderName & """ is empty.")
                Else
                    Console.WriteLine(files.Count & " filenames processed, " & renamedOK & " names renamed, " & renameFailed & " failed to rename.")
                End If

                Console.WriteLine(nonAsciiFileNameCount & " filenames contain non-ASCII characters.")
            End If
        Else
            ShowUsage()
        End If

        Console.WriteLine("Finished.")
        Console.ReadLine()

    End Sub

    Function RenameFile(oldName As String, newName As String) As Boolean

        System.IO.Directory.Move(oldName, newName)

        Return True

    End Function

    Sub ShowUsage()

        Console.WriteLine("Usage: AsciiTheFolder [-v] <folderName>")
        Console.WriteLine("Converts all filenames to pure ASCII")
        Console.WriteLine("Options:")
        Console.WriteLine("  -v  - verbose; just display all the changes that should be made")
        Console.WriteLine("  -V  - very verbose; display both changed and unchanged filenames")

    End Sub

    Function IsPureAscii(s As String) As Boolean

        For Each c In s
            Dim ch As Integer = Convert.ToInt32(c)
            If ch < 32 OrElse ch > 126 Then
                Return False
            End If
        Next

        Return True

    End Function

    Function ToPureAscii(s As String) As String

        Dim sb As New System.Text.StringBuilder
        Dim c2 As String

        For Each c In s
            Select Case c
                Case "č" : c2 = "c"
                Case "ć" : c2 = "c"
                Case "ž" : c2 = "z"
                Case "đ" : c2 = "dj"
                Case "š" : c2 = "s"

                Case "Č" : c2 = "C"
                Case "Ć" : c2 = "C"
                Case "Ž" : c2 = "Z"
                Case "Đ" : c2 = "Dj"
                Case "Š" : c2 = "S"

                Case "ö" : c2 = "oe"
                Case "ä" : c2 = "ae"
                Case "ë" : c2 = "ee"
                Case "ü" : c2 = "ue"

                Case "Ö" : c2 = "Oe"
                Case "Ä" : c2 = "Ae"
                Case "Ë" : c2 = "Ee"
                Case "Ü" : c2 = "Ue"

                Case Else : c2 = c
            End Select

            sb.Append(c2)
        Next

        Return sb.ToString

    End Function

End Module
