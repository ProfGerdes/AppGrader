Imports System.IO
Imports SharpCompress.Archive
Imports SharpCompress.Common

Module JHGModule1
    Enum EnForm
        ObjFormPrefixes = 0
        ObjButton = 1
        objLabel = 2
        ObjActiveLabel = 3
        ObjNonactiveLabel = 4
        ObjTextbox = 5
        ObjListbox = 6
        ObjCombobox = 7
        ObjRadioButton = 8
        ObjCheckbox = 9
        ObjGroupBox = 10
        ObjPanel = 11
        ObjWebBrowser = 12
        ObjWebClient = 13
        ' -----------------------------
        ObjOpenFileDialog = 14
        ObjSaveFileDialog = 15
        ' -----------------------------
        FormName = 16
        FormText = 17
        FormBackColor = 18
        FormAcceptButton = 19
        FormCancelButton = 20
        FormStartPosition = 21
    End Enum

    Enum EnSummary
        StudentID = 0
        AppTitle = 1
        AssignRoot = 2
        AssignPath = 3   ' This is specific to the student
        CompileDate = 4

        'allHaveFrm As Boolean
        'NumberOfModules As Integer
        'FlowControl As MyFlowControl
        'AppInfo As MyAppInfo

        OptionStrict = 5
        OptionExplicit = 6

        hasSLN = 7
        VBVersion = 8
        hasVBproj = 9
        hasSplashScreen = 10
        hasAboutBox = 11
        Modules = 12
        '     AppTitle =

        InfoAppTitle = 13
        InfoDescription = 14
        InfoProject = 15
        InfoCompany = 16
        InfoProduct = 17
        InfoTrademark = 18
        InfoCopyright = 19
        InfoGUID = 20
        '     BreakPoints =
        '     WatchVariables =
        CommentSub = 21
        CommentIF = 22
        CommentFor = 23
        CommentDo = 24
        CommentWhile = 25
        CommentSelect = 26


        VarBoolean = 27
        VarInteger = 28
        VarDecimal = 29
        VarDate = 30
        VarString = 31

        VarArrays = 32
        VarLists = 33
        VarStacks = 34
        VarStructures = 35

        varPrefixes = 36

        LogicFlowControl = 37
        LogicIF = 38
        LogicFor = 39
        LogicDo = 40
        LogicWhile = 41
        LogicCase = 42
        LogicElse = 43
        LogicElseIF = 44
        LogicSelect = 45
        LogicTry = 46
        LogicScreenReader = 47
        LogicScreenWriter = 48
        LogicScreenReaderClosed = 49
        LogicScreenWriterClosed = 50
        LogicSub = 51
        LogicFunction = 52
        LogicOptional = 53
        LogicByRef = 54
        LogicCStr = 55
        LogicToString = 56
        LogicToStringFormat = 57

        LogicVarPrefixes = 58
        LogicNestedIF = 59
        LogicNestedFor = 60

        '       LogicStringFormatting =
        LogicComplexConditions = 61
        LogicCaseInsensitive = 62
        LogicStringFormat = 63
        Concatination = 64
        LogicFormLoad = 65

        SystemIO = 66
        SystemNet = 67
        SystemDB = 68
    End Enum

    Enum dgvs
        ApplicationSettings = 1
        SystemVariables = 2
        LogicVariables = 3
        Splash = 4
        AdvancedVariables = 5
        FormProperties = 6
    End Enum


    Public Structure MyStatementType
        Dim isIF As Boolean
        Dim isFor As Boolean
        Dim isDo As Boolean
        Dim isNext As Boolean
        Dim iswhile As Boolean
        Dim isUntil As Boolean
        Dim isSub As Boolean
        Dim isFunction As Boolean
        Dim isEndIF As Boolean
        Dim isendSub As Boolean
        Dim isendFunction As Boolean
        Dim isEndModule As Boolean
    End Structure

    Structure myBackIndex
        Dim ptr As Integer
        Dim Name As String
        Dim dgv As Integer
    End Structure

    Structure MyItems1
        Dim ID As Integer
        Dim Name As String
        Dim dgv As Integer
    End Structure


    Public Items1 As New List(Of Assignment.MyItems)
    Public BackIndexAppSum As New List(Of myBackIndex)
    Public BackIndexAppFrm As New List(Of myBackIndex)


    Public myindex As Integer


    Public Structure MyErrorComments
        Dim topic As String
        Dim Comment As String
    End Structure

    'Public Structure MySettings
    '    Public ID As Integer
    '    Public Name As String
    '    Public Req As Boolean
    '    Public ShowVar As Boolean
    '    Public PtsPerError As Decimal
    '    Public MaxPts As Decimal
    'End Structure




    Public strStudentID As String
    Public strAssignmentSummary As String = ""
    Public EarliestPostDate As Date
    Public OutputFile As String = ""

    Public TotalLinesOfCode As Integer
    Public FileLinesOfCode As Integer
    Public TotalPossiblePts As Decimal
    Public TotalScore As String = ""
    Public SubmissionCompileTime As String = ""
    Public SubmissionCompileDate As String = ""

    Public bullet As String = Chr(149) & " "


    ' Config Settings
    Public CfgLanguage As String = "VB"
    Public cfgAssignmentTitle As String = ""

    Public CfgPath1 As String = "MyDocuments"
    Public AllowOverwrite As Boolean = False
    Public strOutputPath As String = ""     ' this is the root for the whole assignment 
    Public strStudentRoot As String = ""
    Public strStudentPath As String = ""
    Public strProjectFile As String = ""
    Public strProjectName As String = ""

      ' ==========================================================

    Public ErrorComments As New List(Of ErrComments)
    Public GuidIssues As Boolean = False
    Public CRCIssues As Boolean = False

    Public StudentReportPath As String = ""

    '  Public chkCommentAllVars As Boolean = True
    Public pbar3max As Integer = 100
    Public HideGray As String = "Hide"
    ' ===========================================================================================

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

        Try
            SearchStart = Str.IndexOf(OpenStr) + OpenStr.Length + 1
            SearchEnd = Str.IndexOf(CloseStr, SearchStart - 1)
            s = Mid(Str, SearchStart, SearchEnd - SearchStart + 1)
            If TrimToClosestr Then
                Str = Str.Substring(SearchEnd)
            End If
        Catch
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

        Dim worker As System.ComponentModel.BackgroundWorker = DirectCast(sender, System.ComponentModel.BackgroundWorker)

        path = path & "\"

        n = IO.Directory.GetFiles(path, ext1, SearchOption.TopDirectoryOnly).Length

        For Each fn As String In IO.Directory.GetFiles(path, ext1, SearchOption.TopDirectoryOnly)
            i = i + 1
            If ShortenFilename(fn, shortfn) Then
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
                        ArchiveExtract(path & shortfn & ext)
                    End If

                    nProcessed = nProcessed + 1
                Catch ex As Exception
                    MessageBox.Show("ShortenFilename - " & ex.Message)
                End Try
            End If

            x = CInt(i * 100 / n)
            worker.ReportProgress(x, "Extracting Student Work")

        Next

        Return nProcessed
    End Function



    Function ShortenFilename(filename As String, ByRef ShortFilename As String) As Boolean
        ' this function determines what the shortened filename is. It assumes a structure of 
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
                For i = 2 To ss.GetUpperBound(0)
                    If ss(i) = "attempt" Then
                        ShortFilename = ss(i - 1)
                        flagZipFile = True
                        i = ss.GetUpperBound(0)
                    End If
                Next i
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
                '     MessageBox.Show(ex.Message)
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


    Sub stuffBackIndex()

        'ApplicationSettings = 1
        'SystemVariables = 2
        'LogicVariables = 3
        'Splash = 4
        'AdvancedVariables = 5
        'FormProperties = 6
        ' --------------------------------------------------------------------------------------------------------
        '        Dim Assign As New Assignment("", "", "", "", #1/1/2015#, True, True, True, "2013 Professional", True, True, True, 2)

        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.AssignRoot, .Name = "AssignRoot", .dgv = 0})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.AssignPath, .Name = "AssignPath", .dgv = 0}) ' This is specific to the student
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.CompileDate, .Name = "CompileDate", .dgv = 0})

        'allHaveFrm As Boolean
        'NumberOfModules As Integer
        'FlowControl As MyFlowControl
        'AppInfo As MyAppInfo

        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.OptionStrict, .Name = "OptionStrict", .dgv = dgvs.ApplicationSettings})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.OptionExplicit, .Name = "OptionExplicit", .dgv = dgvs.ApplicationSettings})

        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.hasSLN, .Name = "hasSLN", .dgv = dgvs.ApplicationSettings})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.VBVersion, .Name = "VBVersion", .dgv = dgvs.ApplicationSettings})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.hasVBproj, .Name = "hasVBProj", .dgv = dgvs.ApplicationSettings})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.Modules, .Name = "Modules", .dgv = dgvs.ApplicationSettings})
        '     AppTitle =

        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.hasSplashScreen, .Name = "hasSplashScreen", .dgv = dgvs.Splash})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.hasAboutBox, .Name = "hasAboutBox", .dgv = dgvs.Splash})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.InfoAppTitle, .Name = "InfoAppTitle", .dgv = dgvs.Splash})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.InfoDescription, .Name = "InfoDescription", .dgv = dgvs.Splash})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.InfoProject, .Name = "InfoProject", .dgv = dgvs.Splash})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.InfoCompany, .Name = "InfoCompany", .dgv = dgvs.Splash})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.InfoProduct, .Name = "InfoProduct", .dgv = dgvs.Splash})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.InfoTrademark, .Name = "InfoTrademark", .dgv = dgvs.Splash})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.InfoCopyright, .Name = "InfoCopyright", .dgv = dgvs.Splash})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.InfoGUID, .Name = "InfoGUID", .dgv = dgvs.Splash})
        '     BreakPoints =
        '     WatchVariables =
        ' --------------------------------------------------------------------
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.CommentSub, .Name = "CommentSub", .dgv = dgvs.SystemVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.CommentIF, .Name = "CommentIF", .dgv = dgvs.SystemVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.CommentFor, .Name = "CommentFOR", .dgv = dgvs.SystemVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.CommentDo, .Name = "CommentDO", .dgv = dgvs.SystemVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.CommentWhile, .Name = "CommentWHILE", .dgv = dgvs.SystemVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.CommentSelect, .Name = "CommentSELECT", .dgv = dgvs.SystemVariables})

        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.VarBoolean, .Name = "VarBoolean", .dgv = dgvs.SystemVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.VarInteger, .Name = "VarInteger", .dgv = dgvs.SystemVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.VarDecimal, .Name = "VarDecimal", .dgv = dgvs.SystemVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.VarDate, .Name = "VarDate", .dgv = dgvs.SystemVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.VarString, .Name = "VarString", .dgv = dgvs.SystemVariables})

        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.VarArrays, .Name = "VarArrays", .dgv = dgvs.SystemVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.VarLists, .Name = "VarLists", .dgv = dgvs.SystemVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.VarStacks, .Name = "VarStacks", .dgv = dgvs.SystemVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.VarStructures, .Name = "VarStructures", .dgv = dgvs.SystemVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.varPrefixes, .Name = "VarPrevixes", .dgv = dgvs.SystemVariables})

        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.SystemIO, .Name = "SystemIO", .dgv = dgvs.SystemVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.SystemNet, .Name = "SystemNet", .dgv = dgvs.SystemVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.SystemDB, .Name = "SystemDB", .dgv = dgvs.SystemVariables})
        ' --------------------------------------------------------------------
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicFlowControl, .Name = "LogicFlowControl", .dgv = dgvs.LogicVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicIF, .Name = "LogicIF", .dgv = dgvs.LogicVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicFor, .Name = "LogicFOR", .dgv = dgvs.LogicVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicDo, .Name = "LogicDO", .dgv = dgvs.LogicVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicWhile, .Name = "LogicWHILE", .dgv = dgvs.LogicVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicCase, .Name = "LogicCASE", .dgv = dgvs.LogicVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicElse, .Name = "LogicELSE", .dgv = dgvs.LogicVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicElseIF, .Name = "LogicELSEIF", .dgv = dgvs.LogicVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicSelect, .Name = "LogicSELECT", .dgv = dgvs.LogicVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicTry, .Name = "LogicTRY", .dgv = dgvs.LogicVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicScreenReader, .Name = "LogicScreenReader", .dgv = dgvs.LogicVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicScreenWriter, .Name = "LogicScreenWriter", .dgv = dgvs.LogicVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicScreenReaderClosed, .Name = "LogicScreenReaderClosed", .dgv = dgvs.LogicVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicScreenWriterClosed, .Name = "LogicScreenWriterClosed", .dgv = dgvs.LogicVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicSub, .Name = "LogicSub", .dgv = dgvs.LogicVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicFunction, .Name = "LogicFunction", .dgv = dgvs.LogicVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicOptional, .Name = "LogicOptional", .dgv = dgvs.LogicVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicByRef, .Name = "LogicByRef", .dgv = dgvs.LogicVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicCStr, .Name = "LogicCStr", .dgv = dgvs.LogicVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicToString, .Name = "LogicToString", .dgv = dgvs.LogicVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicToStringFormat, .Name = "LogicToStringFormat", .dgv = dgvs.LogicVariables})

        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicVarPrefixes, .Name = "LogicVarPrefixes", .dgv = dgvs.LogicVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicNestedIF, .Name = "LogicNestedIF", .dgv = dgvs.LogicVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicNestedFor, .Name = "LogicNestedFOR", .dgv = dgvs.LogicVariables})

        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicComplexConditions, .Name = "LogicComplexConditions", .dgv = dgvs.LogicVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicCaseInsensitive, .Name = "LogicCaseInsensitive", .dgv = dgvs.LogicVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicStringFormat, .Name = "LogicStringFormat", .dgv = dgvs.LogicVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.Concatination, .Name = "LogicConcatination", .dgv = dgvs.LogicVariables})
        BackIndexAppSum.Add(New myBackIndex With {.ptr = EnSummary.LogicFormLoad, .Name = "LogicFormLoad", .dgv = dgvs.FormProperties})


        ' ===========================================================

        BackIndexAppFrm.Add(New myBackIndex With {.ptr = EnForm.ObjFormPrefixes, .Name = "ObjFormPrefixes", .dgv = dgvs.FormProperties})
        BackIndexAppFrm.Add(New myBackIndex With {.ptr = EnForm.ObjButton, .Name = "ObjButton", .dgv = dgvs.FormProperties})
        BackIndexAppFrm.Add(New myBackIndex With {.ptr = EnForm.objLabel, .Name = "ObjLabel", .dgv = dgvs.FormProperties})
        BackIndexAppFrm.Add(New myBackIndex With {.ptr = EnForm.ObjActiveLabel, .Name = "ObjActiveLabel", .dgv = dgvs.FormProperties})
        BackIndexAppFrm.Add(New myBackIndex With {.ptr = EnForm.ObjNonactiveLabel, .Name = "ObjNonactiveLabel", .dgv = dgvs.FormProperties})
        BackIndexAppFrm.Add(New myBackIndex With {.ptr = EnForm.ObjTextbox, .Name = "ObjTextbox", .dgv = dgvs.FormProperties})
        BackIndexAppFrm.Add(New myBackIndex With {.ptr = EnForm.ObjListbox, .Name = "ObjListbox", .dgv = dgvs.FormProperties})
        BackIndexAppFrm.Add(New myBackIndex With {.ptr = EnForm.ObjCombobox, .Name = "ObjCombobox", .dgv = dgvs.FormProperties})
        BackIndexAppFrm.Add(New myBackIndex With {.ptr = EnForm.ObjRadioButton, .Name = "ObjRadioButton", .dgv = dgvs.FormProperties})
        BackIndexAppFrm.Add(New myBackIndex With {.ptr = EnForm.ObjCheckbox, .Name = "ObjCheckbox", .dgv = dgvs.FormProperties})
        BackIndexAppFrm.Add(New myBackIndex With {.ptr = EnForm.ObjGroupBox, .Name = "ObjGroupBox", .dgv = dgvs.FormProperties})
        BackIndexAppFrm.Add(New myBackIndex With {.ptr = EnForm.ObjPanel, .Name = "ObjPanel", .dgv = dgvs.FormProperties})
        BackIndexAppFrm.Add(New myBackIndex With {.ptr = EnForm.ObjWebBrowser, .Name = "ObjWebBrower", .dgv = dgvs.FormProperties})
        BackIndexAppFrm.Add(New myBackIndex With {.ptr = EnForm.ObjWebClient, .Name = "ObjWebClient", .dgv = dgvs.FormProperties})
        ' -----------------------------
        BackIndexAppFrm.Add(New myBackIndex With {.ptr = EnForm.ObjOpenFileDialog, .Name = "ObjOpenFileDialog", .dgv = dgvs.FormProperties})
        BackIndexAppFrm.Add(New myBackIndex With {.ptr = EnForm.ObjSaveFileDialog, .Name = "ObjSaveFileDialog", .dgv = dgvs.FormProperties})
        ' -----------------------------
        BackIndexAppFrm.Add(New myBackIndex With {.ptr = EnForm.FormName, .Name = "FormName", .dgv = dgvs.FormProperties})
        BackIndexAppFrm.Add(New myBackIndex With {.ptr = EnForm.FormText, .Name = "FormText", .dgv = dgvs.FormProperties})
        BackIndexAppFrm.Add(New myBackIndex With {.ptr = EnForm.FormBackColor, .Name = "FormBackColor", .dgv = dgvs.FormProperties})
        BackIndexAppFrm.Add(New myBackIndex With {.ptr = EnForm.FormAcceptButton, .Name = "FromAcceptButton", .dgv = dgvs.FormProperties})
        BackIndexAppFrm.Add(New myBackIndex With {.ptr = EnForm.FormCancelButton, .Name = "FormCancelButton", .dgv = dgvs.FormProperties})
        BackIndexAppFrm.Add(New myBackIndex With {.ptr = EnForm.FormStartPosition, .Name = "FormStartPosition", .dgv = dgvs.FormProperties})

    End Sub

    Sub AppendToBackIndex(ByRef BI As List(Of myBackIndex), ID As Integer, txt As String)

        Dim B = New myBackIndex
        B.ptr = ID
        B.Name = txt
        BI.Add(B)
    End Sub

End Module
