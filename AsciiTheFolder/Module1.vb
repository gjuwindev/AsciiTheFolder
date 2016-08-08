Module Module1

    Sub Main(args() As String)

        If args.Count = 1 Then
            Dim folderName = args(0)
            Dim folderNameLength = folderName.Length

            If IsPureAscii(folderName) = False Then
                Console.WriteLine("Folder name contains non-ASCI characters.")
            End If

            Dim files() As String = System.IO.Directory.GetFiles(folderName, "*.*", System.IO.SearchOption.AllDirectories)

            Console.WriteLine(folderName)
            System.IO.Directory.SetCurrentDirectory(folderName)
            For Each file In files
                Dim oldName = file.Substring(folderNameLength + 1)
                Dim newName = ToPureAscii(oldName)
                Console.WriteLine(newName)
                If oldName <> newName Then
                    System.IO.Directory.Move(oldName, newName)
                End If
            Next
        Else
            Console.WriteLine("Usage: AsciiTheFolder <folderName>")
            Console.WriteLine("Converts all filenames to pure ASCII")
        End If

        Console.ReadLine()

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
