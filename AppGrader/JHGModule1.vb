Imports System.IO
Imports SharpCompress.Archive
Imports SharpCompress.Common
' And System.Security.Cryptography for Hashes : MD5, SHA1, SHA256, ...
Imports System.Security
Imports System.Security.Cryptography
Module JHGModule1

    Public Function RemoveWhiteSpace(ByVal Str As String) As String
        ' This function removes all white space with the exception of a single space character.
        ' Including all Tab stops, Returns, and multiple spaces.
        Dim i As Integer

        Str = Str.Trim
        For i = 1 To 31
            Str = Str.Replace(Chr(i), "")
        Next

        While Str.Contains("     ")
            Str = Str.Replace("     ", " ")
        End While

        While Str.Contains("  ")
            Str = Str.Replace("  ", " ")
        End While
        Return Str
    End Function

    Public Function returnBetween(ByRef Str As String, ByVal OpenStr As String, ByVal CloseStr As String, Optional ByVal TrimToClosestr As Boolean = False) As String
        Dim s As String = Str
        Dim SearchStart As Integer = 0, SearchEnd As Integer = 0
        ' both start and end strings must be in string, otherwise a null string is returned. 
        Try
            SearchStart = Str.IndexOf(OpenStr) + OpenStr.Length + 1
            SearchEnd = Str.IndexOf(CloseStr, SearchStart - 1)
            If SearchEnd - SearchStart + 1 > 0 Then
                s = Mid(Str, SearchStart, SearchEnd - SearchStart + 1)
                If TrimToClosestr Then
                    Str = Str.Substring(SearchEnd)
                End If
            Else
                s = ""
            End If
        Catch ex As exception
            s = ""
        End Try
        Return s
    End Function

    Public Function RemoveBetween(ByVal Str As String, ByVal OpenStr As String, ByVal CloseStr As String) As String
        ' Custom function to return the substring between two specified characters/strings
        Dim i, j As Integer



        ' if openstr = "", it starts at the beginning of string.

        If OpenStr = "" Then
            If Str.Contains(CloseStr) Then Str = Str.Substring(Str.IndexOf(CloseStr))
        Else
            If CloseStr = "" Then
                Str = Str.Substring(0, OpenStr.IndexOf(OpenStr) + OpenStr.Length)
            Else
                i = Str.IndexOf(OpenStr) + OpenStr.Length
                j = Str.IndexOf(CloseStr, i)

                Str = Str.Substring(0, i) & Str.Substring(j)
            End If
        End If

        Return Str
    End Function

    Public Function ToTitleCase(ByVal source As String) As String
        Return Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(source.ToLower())
    End Function

    Public Function TrimUpTo(ByVal str As String, ByVal closestr As String) As String
        ' this removes the first part of the string
        ' if string not found, it returns a null string
        Dim i As Integer

        If closestr.Length > 0 Then
            If str.IndexOf(closestr) > -1 Then
                i = str.IndexOf(closestr) + closestr.Length
                str = str.Substring(i)
            Else
                str = ""
            End If
        End If
        Return str
    End Function

    Public Function TrimAfter(ByVal str As String, ByVal startstr As String, Optional removedelim As Boolean = False) As String
        ' this removes the portion of the string after the specified string
        ' if string not found, it returns the full string
        Dim i As Integer
        Dim x As Integer = startstr.Length

        If removedelim Then x = 0

        If startstr.Length > 0 Then
            If str.IndexOf(startstr) > -1 Then
                i = str.IndexOf(startstr) + x
                str = str.Substring(0, i)
            Else
                ' return the whole string
            End If
        End If
        Return str
    End Function


    Public Function UseSmaller(ByVal NumOne As Integer, ByVal NumTwo As Integer) As Integer
        ' Return the smaller of two numbers

        If NumOne <= NumTwo Then
            Return NumOne
        Else
            Return NumTwo
        End If
    End Function

    Public Function LeftNChar(ByVal s As String, ByVal pad As Char, ByVal n As Integer) As String
        s = s.PadLeft(n, pad)
        Return s.Substring(s.Length - n)
    End Function

    Public Sub Pause(ByVal dt As Double)
        ' a lot of wasted cycles.

        Dim CurrentTime As Date = Now()
        CurrentTime = DateAdd(DateInterval.Second, dt, CurrentTime)
        Application.DoEvents()
        Do Until Now() > CurrentTime
        Loop
    End Sub

    Public Function ReturnLastField(source As String, delim As String) As String
        Dim ss() As String
        Dim d() As String = {delim}
        Dim s As String
        ss = source.Split(d, StringSplitOptions.None)
        s = ss(ss.GetUpperBound(0))

        Return s
    End Function


    Public Function RemoveLastField(source As String, delim As String) As String
        Dim ss() As String
        Dim d() As String = {delim}
        Dim s As String
        ss = source.Split(d, StringSplitOptions.None)
        s = ss(ss.GetUpperBound(0))

        Return source.Substring(0, source.Length - s.Length - 1)
    End Function


    Function ShortenFilenamesInFolder(folder As String, sender As System.ComponentModel.BackgroundWorker) As Integer
        ' shorten all compressed files in specified folder
        ' ----------------------------------------------------------------------------------------
        Dim nProcessed As Integer = 0

        nProcessed = nProcessed + shortenFileNames(folder, ".zip", sender)
        nProcessed = nProcessed + shortenFileNames(folder, ".rar", sender)
        nProcessed = nProcessed + shortenFileNames(folder, ".7z", sender)
        nProcessed = nProcessed + shortenFileNames(folder, ".tar", sender)
        nProcessed = nProcessed + shortenFileNames(folder, ".gzip", sender)

        Return nProcessed

    End Function

    Function shortenFileNames(path As String, ext As String, sender As System.ComponentModel.BackgroundWorker) As Integer
        ' this shortens all the file names. It is based on the structure of the files downloaded from Blackboard
        ' It only shortens the files with the specified extension.
        ' It assumes the structure of the files is as follows, with underscores as delimiters

        '  Assignment Name _ Student ID _ "attempt" _ Date  Stamp _ submission filename

        ' The Path is the root location where the files can be found. The sub renames the filename to be 
        ' the student ID with the specificed extension. It also appends a counter to differentiate multiple 
        ' submissions from the same student. The first file is designated _a, the second _b, etc.
        ' ----------------------------------------------------------------------------------------------------------

        Dim shortfn As String = ""
        Dim ext1 As String = "*" & ext
        Dim lastshortfn As String = ""
        Dim cnt = 0
        Dim nProcessed As Integer = 0
        Dim i As Integer = 0
        Dim n As Integer = 0
        Dim x As Integer
        Dim percent As Decimal = 0
        Dim strEnding As String = ""

        Dim worker As System.ComponentModel.BackgroundWorker = DirectCast(sender, System.ComponentModel.BackgroundWorker)

        path = path & "\"

        n = IO.Directory.GetFiles(path, ext1, SearchOption.TopDirectoryOnly).Length

        For Each fn As String In IO.Directory.GetFiles(path, ext1, SearchOption.TopDirectoryOnly)
            i = i + 1
            ' check to see if the filename has already been shortened.
            ' see if it has a structure of   name_a.zip
            strEnding = ReturnLastField(fn, "_")
            If strEnding.Length - ext.Length > 0 Then ' > 1 Then  '  this indicates we have more than a single letter in suffix =========== I disabled this and allowed this to run for all files
                If ShortenFilename(fn, shortfn) Then
                    If shortfn.Contains(".") Then shortfn = TrimAfter(shortfn, ".", True)
                    If shortfn = lastshortfn Then
                        cnt = cnt + 1
                    Else
                        cnt = Asc("a")
                    End If
                    lastshortfn = shortfn
                    shortfn = shortfn & "_" & Chr(cnt)
                    Try
                        If AllowOverwrite And File.Exists(path & shortfn & ext) Then
                            File.Delete(path & shortfn & ext)
                        End If

                        Rename(fn, path & shortfn & ext)

                        If AllowOverwrite And File.Exists(path & shortfn) Then
                            Directory.Delete(path & shortfn)
                        End If

                        If AllowOverwrite Or Not File.Exists(path & shortfn) Then
                            ArchiveExtract(path & shortfn & ext)  ' this extracts the file, but not recursively, 

                            'find all files contained in this folder
                            For Each fn2 As String In IO.Directory.GetFiles(path & shortfn & "\", "*.zip", SearchOption.AllDirectories)
                                ' if any of them are archive files, uncompress them
                                ArchiveExtract(fn2)
                            Next
                            For Each fn2 As String In IO.Directory.GetFiles(path & shortfn & "\", "*.rar", SearchOption.AllDirectories)
                                ' if any of them are archive files, uncompress them
                                ArchiveExtract(fn2)
                            Next
                            For Each fn2 As String In IO.Directory.GetFiles(path & shortfn & "\", "*.7z", SearchOption.AllDirectories)
                                ' if any of them are archive files, uncompress them
                                ArchiveExtract(fn2)
                            Next
                            For Each fn2 As String In IO.Directory.GetFiles(path & shortfn & "\", "*.tar", SearchOption.AllDirectories)
                                ' if any of them are archive files, uncompress them
                                ArchiveExtract(fn2)
                            Next
                            For Each fn2 As String In IO.Directory.GetFiles(path & shortfn & "\", "*.gzip", SearchOption.AllDirectories)
                                ' if any of them are archive files, uncompress them
                                ArchiveExtract(fn2)
                            Next

                        End If

                        nProcessed = nProcessed + 1
                    Catch ex As Exception
                        MessageBox.Show("ShortenFilename - " & ex.Message)
                    End Try
                End If
            End If

            x = CInt(i * 100 / n)
            worker.ReportProgress(x, "Extracting Student Work")

        Next

        Return nProcessed
    End Function



    Function ShortenFilename(filename As String, ByRef ShortFilename As String) As Boolean
        ' this function determines what the shortened filename is.  
        ' It assumes the structure of the files is as follows, with underscores as delimiters

        '  Assignment Name _ Student ID _ "attempt" _ Date  Stamp _ submission filename

        ' The shortened filename is the student ID with the specificed extension. It also appends 
        ' a counter to differentiate multiple submissions from the same student. The first file 
        ' is designated _a, the second _b, etc.

        ' The function returns true if the file is a compressed file. The shortened filename is passed
        ' back with the second variable, only for compressed files.
        ' ---------------------------------------------------------------------------------------------------
        Dim i As Integer
        Dim flagZipFile As Boolean

        Dim ss() As String
        Dim ext As String

        ext = filename.Substring(filename.Length - 5)
        ext = ext.Substring(ext.IndexOf(".") + 1)

        Select Case ext.ToUpper
            Case "ZIP", "RAR", "7Z", "TAR", "GZIP"
                ss = filename.Split(CChar("_"))

                If ss.GetUpperBound(0) = 1 Then
                    If ss(1).Substring(1, 1) = "." Then   ' no need to shorten it - it already has been done.
                        ShortFilename = filename
                        flagZipFile = True
                    Else
                        ShortFilename = ss(1)    ' go ahead and shorten filename
                        flagZipFile = True
                    End If

                Else

                    For i = 2 To ss.GetUpperBound(0)
                        If ss(i) = "attempt" Then         ' This is designed to work with Blackboard files
                            ShortFilename = ss(i - 1)
                            flagZipFile = True
                            i = ss.GetUpperBound(0)
                        End If
                    Next i
                End If
            Case Else
                flagZipFile = False
        End Select


        Return flagZipFile

    End Function

    Function GetFileSource(filename As String, ByRef source As String) As Boolean
        Try
            Dim sr As New StreamReader(filename)
            source = sr.ReadToEnd
            sr.Close()
            Return True
        Catch
            source = ""
            Return False
        End Try
    End Function

    Sub ArchiveExtract(filename As String)
        ' this uses utilities from SharpCompress http://sharpcompress.codeplex.com/) to decompress the filename that is 
        ' passed to the subroutine. It places the files in a folder of the same name as the file
        ' It handles 5 different file types (zip, rar, 7z, tar, and gzip).
        ' -----------------------------------------------------------------------------------------------------------------
        Dim stream1 As Stream = File.OpenRead(filename)
        Dim archive = ArchiveFactory.Open(stream1)

        Dim ext As String
        Dim destination As String

        ext = IO.Path.GetExtension(filename)
        destination = filename.Replace(ext, "")

        Select Case ext.ToLower
            Case ".zip"
                archive.WriteToDirectory(destination, ExtractOptions.ExtractFullPath Or ExtractOptions.Overwrite)
            Case ".rar"
                archive.WriteToDirectory(destination, ExtractOptions.ExtractFullPath Or ExtractOptions.Overwrite)
            Case ".7z"
                archive.WriteToDirectory(destination, ExtractOptions.ExtractFullPath Or ExtractOptions.Overwrite)
            Case ".tar"
                archive.WriteToDirectory(destination, ExtractOptions.ExtractFullPath Or ExtractOptions.Overwrite)
            Case ".gzip"
                archive.WriteToDirectory(destination, ExtractOptions.ExtractFullPath Or ExtractOptions.Overwrite)
        End Select
    End Sub


    Sub DeleteDirectory(path As String)
        Dim ImDone As Boolean = False
        Dim cnt As Integer = 0

        Do Until ImDone Or cnt > 5  ' this was inserted to address IO errors.
            Try

                If My.Computer.FileSystem.DirectoryExists(path) Then
                    cnt += 1
                    My.Computer.FileSystem.DeleteDirectory(path, FileIO.DeleteDirectoryOption.DeleteAllContents)
                    ImDone = True
                End If

            Catch ex As InvalidExpressionException
                MessageBox.Show(ex.Message)
            End Try
        Loop
    End Sub

    ' =============================================================================================================
    ' http://stackoverflow.com/questions/3448103/how-can-i-delete-an-item-from-an-array-in-vb-net/15182002#15182002

    <System.Runtime.CompilerServices.Extension()> Public Sub RemoveAt(Of T)(ByRef a() As T, ByVal index As Integer)
        ' Move elements after "index" down 1 position.
        Array.Copy(a, index + 1, a, index, UBound(a) - index)
        ' Shorten by 1 element.
        ReDim Preserve a(UBound(a) - 1)
    End Sub

    <System.Runtime.CompilerServices.Extension()> Public Sub DropFirstElement(Of t)(ByRef a() As t)
        a.RemoveAt(0)
    End Sub

    <System.Runtime.CompilerServices.Extension()> Public Sub DropLastElement(Of T)(ByRef a() As T)
        a.RemoveAt(UBound(a))
    End Sub
    ' =============================================================================================================


    ' =============================================================================================================
    '   http://us.informatiweb.net/programmation/36--generate-hashes-md5-sha-1-and-sha-256-of-a-file.html
    ' =============================================================================================================

    ' Function to obtain the desired hash of a file
    Function hash_generator(ByVal hash_type As String, ByVal file_name As String) As String


        ' We declare the variable : hash
        Dim hash As MD5
        hash = MD5.Create

        'If hash_type.ToLower = "md5" Then
        '    ' Initializes a md5 hash object
        '    hash = MD5.Create
        '    'ElseIf hash_type.ToLower = "sha1" Then
        '    ' Initializes a SHA-1 hash object
        '    hash = SHA1.Create()
        'ElseIf hash_type.ToLower = "sha256" Then
        '    ' Initializes a SHA-256 hash object
        '    hash = SHA256.Create()
        'Else
        '    MsgBox("Unknown type of hash : " & hash_type, MsgBoxStyle.Critical)
        '    Return False
        'End If

        ' We declare a variable to be an array of bytes
        Dim hashValue() As Byte

        ' We create a FileStream for the file passed as a parameter
        Dim fileStream As FileStream = File.OpenRead(file_name)
        ' We position the cursor at the beginning of stream
        fileStream.Position = 0
        ' We calculate the hash of the file
        hashValue = hash.ComputeHash(fileStream)
        ' The array of bytes is converted into hexadecimal before it can be read easily
        Dim hash_hex = PrintByteArray(hashValue)

        ' We close the open file
        fileStream.Close()

        ' The hash is returned
        Return hash_hex

    End Function


    ' We traverse the array of bytes and converting each byte in hexadecimal
    Public Function PrintByteArray(ByVal array() As Byte) As String

        Dim hex_value As String = ""

        ' We traverse the array of bytes
        Dim i As Integer
        For i = 0 To array.Length - 1

            ' We convert each byte in hexadecimal
            hex_value += array(i).ToString("X2")

        Next i

        ' We return the string in lowercase
        Return hex_value.ToLower

    End Function

    ' md5 is a reserved name, so we named the function : md5_hash
    Function md5_hash(ByVal file_name As String) As String
        Return hash_generator("md5", file_name)
    End Function

    'Function sha_1(ByVal file_name As String) As String
    '    Return hash_generator("sha1", file_name)
    'End Function

    'Function sha_256(ByVal file_name As String) As String
    '    Return hash_generator("sha256", file_name)
    'End Function
End Module
