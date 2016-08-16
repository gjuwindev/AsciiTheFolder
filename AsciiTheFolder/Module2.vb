Module Module2

    Function RenameFileOrFolder(folderName As String, oldName As String, newName As String) As String

        ' folderName is the name of the current directory
        ' current working directory has been set in the calling level
        ' this variable serves here just for messaging purposes

        Try
            System.IO.Directory.Move(oldName, newName)
            Return Nothing
        Catch ex As Exception
            Return "FAILED TO RENAME: " & folderName & "\" & oldName & " =>>> " & newName & vbCrLf &
                    "   Exception msg: " & ex.Message
        End Try

    End Function

    ''' <summary>
    ''' Parses command line parameters
    ''' </summary>
    ''' <param name="args">Standard command line arguments</param>
    ''' <returns>True if all OK, false otherwise</returns>
    ''' <remarks>Parses arguments anf fills the following variables: startingFolderName, showAllFilenames, verbose</remarks>
    Function ProcessCommandLineArguments(args() As String) As Boolean

        Dim argList As New List(Of String)(args)

        If argList.Count = 2 Then
            If argList(0) = "-v" Then
                verbose = True
                veryVerbose = False
                'Console.WriteLine("Verbose mode ON.")
                argList.RemoveAt(0)
            ElseIf argList(0) = "-V" Then
                verbose = True
                veryVerbose = True
                'Console.WriteLine("Very verbose mode ON.")
                argList.RemoveAt(0)
            End If
        End If

        If argList.Count = 1 Then
            startingFolderName = argList(0)
            If System.IO.Directory.Exists(startingFolderName) Then
                Return True
            Else
                Console.WriteLine("Folder: """ & startingFolderName & """ does not exist.")
                Return False
            End If
        Else
            ShowUsage()
            Return False
        End If

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
