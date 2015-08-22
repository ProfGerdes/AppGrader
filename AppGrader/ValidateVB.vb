Imports System.IO
'Imports SharpCompress.Archive
'Imports SharpCompress.Common


Module ValidateVB

    '   Public ValidateReport As String
    Public strStudentReport As String
    Public strFacReport As String



    Function InitializeStudentReport() As String
        Dim s As String = ""

        Try
            Dim sr1 As New StreamReader(Application.StartupPath & "\templates\rptStudentHeader.html")

            s = sr1.ReadToEnd
            sr1.Close()

            s = s.Replace("[title]", strStudentID & " - " & cfgAssignmentTitle & " Summary")
            ' strStudentReport = strStudentReport.Replace("[STUDENT]", strStudentID & " - " & cfgAssignmentTitle & " Summary")
            s = s.Replace("[ASSIGNMENTNAME]", frmMain.txtAssignmentName.Text)
            s = s.Replace("[STUDENT]", strStudentID)
            s = s.Replace("[VERSION]", Application.ProductVersion)

            '            strStudentReport = s & strStudentReport
        Catch ex As Exception
            MessageBox.Show("InitializeStudentReport - " & ex.Message)
        End Try

        ' -----------------------------------------------------------------------------
        ' Determine the date of the last compile
        Dim di As New IO.DirectoryInfo(Application.StartupPath)
        Dim diar1 As IO.FileInfo() = di.GetFiles("*.exe")
        Dim dra As IO.FileInfo

        SubmissionCompileTime = " - "
        SubmissionCompileDate = " - "
        'list the names of all files in the specified directory
        For Each dra In diar1
            If Not dra.FullName.ToLower.Contains("\bin\debug") Then
                If Not dra.FullName.ToString.Contains(".vshost.exe") Then
                    SubmissionCompileTime = dra.CreationTime.ToLongTimeString()
                    SubmissionCompileDate = dra.CreationTime.ToLongDateString()
                End If
            End If
        Next

        s = s.Replace("[APPTIME]", SubmissionCompileTime)
        s = s.Replace("[APPDATE]", SubmissionCompileDate)
        Return s

    End Function

    Function InitializeFacultyReport() As String
        Dim s As String = ""
        Dim sr As StreamReader
        Try
            sr = File.OpenText(Application.StartupPath & "\templates\rptFacHeader.html")
            s = sr.ReadToEnd
            sr.Close()

            s = s.Replace("[ASSIGNMENTNAME]", cfgAssignmentTitle)
            s = s.Replace("[REPORTDATE]", Today.ToString("d"))
            s = s.Replace("[VERSION]", Application.ProductVersion)
            s = s.Replace("[CONFIGFILE]", frmMain.lblConfigFile.Text)

            s &= ("<tr><th class=""header"";>Student ID</th><th class=""header"";>Filename</th><th class=""header"";>TLOC</th><th class=""header"";>Score</th></tr>" & vbCrLf)

        Catch ex As Exception
            MessageBox.Show("InitializeFacultyReport - " & ex.Message)
        End Try

        Return s
    End Function


    Public Sub CloseFacReport(src As String, fn As String, score As String)
        Dim sr As StreamReader
        Dim sw As StreamWriter

        'remove extra line breaks
        Do While (src.Contains("<br>" & vbCrLf & " <br>"))
            src = src.Replace("<br>" & vbCrLf & " <br>", "<br>")
        Loop

        src = src.Replace("[SCORE]", score)
        src = src.Replace("[CONFIGFILE]", frmMain.lblConfigFile.Text)


        sr = File.OpenText(Application.StartupPath & "\templates\rptFacFooter.html")
        src &= sr.ReadToEnd
        sr.Close()

        sw = File.CreateText((strOutputPath & fn))
        sw.Write(src)
        sw.Close()
    End Sub


    Sub GetAllFilesInBuild(path As String, ByRef filesinbuild As List(Of String))
        Dim strRoot As String
        Dim tmp As String

        Try
            filesinbuild.Clear()

            For Each filename In IO.Directory.GetFiles(path, "*.vbproj", SearchOption.AllDirectories)
                tmp = ReturnLastField(filename, "\")
                strRoot = filename.Replace(tmp, "")
                strStudentPath = strRoot

                If filename.ToLower.EndsWith(".vbproj") Then
                    Dim sr As New StreamReader(filename)
                    Dim s As String
                    Dim delim() As String = {"<Compile Include="""}
                    Dim ss() As String

                    s = sr.ReadToEnd
                    sr.Close()

                    ss = s.Split(delim, StringSplitOptions.None)

                    For i = 1 To ss.GetUpperBound(0)
                        If Not ss(i).Contains(".Designer.vb") And Not ss(i).StartsWith("My Project\") And Not ss(i).StartsWith("ApplicationEvents.vb") Then
                            ' These are the files contained in the build
                            filesinbuild.Add(strRoot & TrimAfter(ss(i), """", True))

                        End If
                    Next i
                End If
            Next filename


        Catch ex As Exception
            MessageBox.Show("GetAllFilesInBuild - " & ex.Message)
        End Try
    End Sub


    Sub DisplayFilesInReports(FilesInBuild As List(Of String))

        strStudentReport &= "<h3>Files Contained in Project</h3>"
        '        strFacReport &= "<h3>Files Contained in Project</h3>"

        strStudentReport &= "<ul>"
        '        strFacReport &= "<ul>"

        For Each filename As String In FilesInBuild
            strStudentReport &= "<li>" & filename.Replace(strStudentRoot & "\", "") & "</li>"    ' hide the local path
        Next filename

        strStudentReport &= "</ul>" & vbCrLf
        '       strFacReport &= "</ul>" & vbCrLf
    End Sub


    ' ==============================================================================================================
    ' Code Checks start here
    ' ==============================================================================================================

    Sub CheckSLNvbProj(ByRef AppInfo As AssignmentInfo)
        ' This checks for both the SLN and VBProj files. It also gets the VB Version used to create the application.
        ' This information is placed in the AppSum Structure
        ' -----------------------------------------------------------------------------------------------------------
        Dim filesource As String
        Dim hasSLNfile As Boolean
        Dim hasVBProjFile As Boolean

        ' --------------------------------------------------------------------------------------------------
        ' check to see if the submission has a sln file
        hasSLNfile = IO.Directory.GetFiles(AppInfo.AssignPath, "*.sln", SearchOption.AllDirectories).GetLength(0) > 0
        If hasSLNfile Then
            AppInfo.hasSLN.Status = vbTrue.ToString
            AppInfo.hasSLN.n += 1
            AppInfo.hasSLN.cssClass = "itemgreen"
        Else
            AppInfo.hasSLN.Status = vbFalse.ToString
            AppInfo.hasSLN.cssClass = "itemred"
        End If

        ' --------------------------------------------------------------------------------------------------
        ' Check the Version Number of Visual Basic

        If hasSLNfile Then
            For Each filename In IO.Directory.GetFiles(AppInfo.AssignPath, "*.sln", SearchOption.AllDirectories)
                If filename.Contains(".sln") Then
                    Dim sr As New StreamReader(filename)
                    filesource = sr.ReadToEnd
                    sr.Close()

                    AppInfo.VBVersion.Status = "VB Version = " & returnBetween(filesource, "# ", vbCrLf)

                End If
            Next filename
        End If
        ' --------------------------------------------------------------------------------------------------
        ' check to see if the submission has a vbProj file
        hasVBProjFile = IO.Directory.GetFiles(AppInfo.AssignPath, "*.vbProj", SearchOption.AllDirectories).GetLength(0) > 0
        If hasVBProjFile Then
            AppInfo.hasVBproj.Status = vbTrue.ToString
            AppInfo.hasVBproj.cssClass = "itemgreen"
            AppInfo.hasVBproj.n += 1
        Else
            AppInfo.hasVBproj.Status = vbFalse.ToString
            AppInfo.hasVBproj.cssClass = "itemred"
        End If

        AppInfo.hasVBproj.Status = hasVBProjFile.ToString
    End Sub



    Sub CheckAPPInfo2(AppDir As String, ByRef SSummary() As MyItems)
        Dim s As String = ""
        Dim s2 As String = ""
        Try
            Dim sr As New StreamReader(AppDir & "\" & "My Project\AssemblyInfo.vb")
            s = sr.ReadToEnd
            s2 = s
            sr.Close()

            's = s2
            'cfgAssignmentTitle = returnBetween(s, "<Assembly: AssemblyTitle(""", """)>", True)
            'SSummary.AppTitle = cfgAssignmentTitle


            s = s2
            SSummary(EnSummary.InfoAppTitle).Status = returnBetween(s, "<Assembly: AssemblyTitle(""", """)>", True)
            s = s2
            SSummary(EnSummary.InfoDescription).Status = returnBetween(s, "<Assembly: AssemblyDescription(""", """)>", True)
            s = s2
            SSummary(EnSummary.InfoCompany).Status = returnBetween(s, "<Assembly: AssemblyCompany(""", """)>", True)
            s = s2
            SSummary(EnSummary.InfoProduct).Status = returnBetween(s, "<Assembly: AssemblyProduct(""", """)>", True)
            s = s2
            SSummary(EnSummary.InfoCopyright).Status = returnBetween(s, "<Assembly: AssemblyCopyright(""", """)>", True)
            s = s2
            SSummary(EnSummary.InfoTrademark).Status = returnBetween(s, "<Assembly: AssemblyTrademark(""", """)>", True)
            s = s2
            SSummary(EnSummary.InfoGUID).Status = returnBetween(s, "<Assembly: Guid(""", """)>", True)

            ' ----------------------------------------------------------------------------------------------
            If Find_Setting("InfoAppTitle", "CheckAppInfo2").Req Then       ' **************** is this correct? *************** Info AppTile is different than AppTitle
                'If cfgAssignmentTitle = str5 Then
                '    strFacReport = (strFacReport & "???")
                'End If
            End If
            ' ----------------------------------------------------------------------------------------------
        Catch ex As Exception
            MessageBox.Show("CheckAPPInfo - Problem extracting information out of AssemblyInfo.vb. The erro encountered is: " & ex.Message)
        End Try
    End Sub


    Sub CheckForOptions2(path As String, filesinbuild As List(Of String), ByRef item As MyItems, optiontype As String)   ' ByRef HasOptionStrict As Boolean, ByRef HasOptionStrictOff As Boolean)
        ' Checks the designed directory fn for the directive Option Strict / Option Explicit

        Dim source As String = ""
        '    Dim tmp As String = ""

        With item
            .Status = "Not Set"

            ' Check each file plus Application.Designer. If it is set of off, that is a big problem and overrides any point where it is set on

            '      Dim i As Integer = 0

            'Searches directory and it's subdirectories for all files, which "*" stands for
            'Say for example you only want to search for jpeg files... then change "*" to "*.jpg"  
            Dim filenames() As String = IO.Directory.GetFiles(path, "Application.Designer.vb", IO.SearchOption.AllDirectories)

            If filenames.Length >= 1 Then
                If GetFileSource(filenames(0), source) Then

                    If source.Contains("Option " & optiontype & " Off") Then
                        .Status = "Off"
                        .cssClass = "itemred"
                        .cnt = +1
                        Exit Sub
                    ElseIf source.Contains("Option " & optiontype & " On") Then
                        .Status = "On"
                        .cssClass = "itemgreen"
                        .n += 1
                    End If
                End If
            End If

            ' Now check each file. If it is set of off, that is a big problem and overrides any point where it is set on
            For Each filename As String In filesinbuild
                If GetFileSource(filename, source) Then
                    If source.Contains("Option " & optiontype & " Off") Then
                        .Status = "Off"
                        .cssClass = "itemred"
                        .cnt = +1
                        Exit Sub
                    ElseIf source.Contains("Option " & optiontype & " On") Then
                        .Status = "On"
                        .cssClass = "itemgreen"
                        .n += 1
                        Exit For
                    End If
                End If
            Next filename
        End With
    End Sub



    Function CheckForComments2(fn As String, filesource As String, ByRef SSummary() As MyItems, ByRef AppForm() As MyItems, sender As System.ComponentModel.BackgroundWorker) As Integer

        Dim worker As System.ComponentModel.BackgroundWorker = DirectCast(sender, System.ComponentModel.BackgroundWorker)
        Dim i As Integer
        Dim ncomments As Integer
        Dim isInIf As Integer
        Dim isInFor As Integer
        ' Dim isinSub As Boolean
        ' Dim isinFunction As Boolean
        ' Dim LastlineComment As Integer
        Dim ss() As String
        Dim lineno() As Integer
        Dim n As Integer
        Dim delim() As String = {vbCrLf}
        Dim firstword As String = ""
        Dim PreviousWord As String = ""
        Dim comBad As String = ""
        '   Dim comGood As String = ""
        Dim x As Integer
        Dim srName(0) As String
        Dim swName(0) As String
        Dim s As String
        Dim doneflag As Boolean

        ss = filesource.Split(delim, StringSplitOptions.None)
        ReDim lineno(ss.GetUpperBound(0))

        n = 0
        ' -------------------------------------------------------------------------------------------------------
        ' number each line, trim each line, remove blank lines, and 
        For i = 0 To ss.GetUpperBound(0)
            ss(i) = ss(i).Trim               ' remove leading spaces

            '  This block removes blank lines
            If ss(i).Length > 0 Then   ' there is something in this line, so process it
                If i <> n Then     ' move record
                    ss(n) = ss(i)
                End If

                ' eliminate continuations
                If i > 0 AndAlso ss(i - 1).EndsWith(" _") Then
                    ss(n - 1) = ss(n - 1).Substring(ss(n - 1).Length - 2) & ss(i)
                Else
                    lineno(n) = i + 1
                    n = n + 1
                End If
            Else
                ' nothing in line, so skip it
            End If
        Next i

        ReDim Preserve ss(n)
        ReDim Preserve lineno(n)

        TotalLinesOfCode += ss.Length
        FileLinesOfCode = ss.Length
        ' ---------------------------------------------------------------------------------------------------------

        ' set default to green
        SSummary(EnSummary.LogicIF).cssClass = "itemgreen"
        SSummary(EnSummary.LogicFor).cssClass = "itemgreen"
        SSummary(EnSummary.LogicDo).cssClass = "itemgreen"
        SSummary(EnSummary.LogicSelectCase).cssClass = "itemgreen"
        SSummary(EnSummary.LogicWhile).cssClass = "itemgreen"
        SSummary(EnSummary.LogicSub).cssClass = "itemgreen"

        ' Now process each line and assess for flow control elements

        For i = 0 To ss.GetUpperBound(0)
            x = CInt(i * 100 / ss.GetUpperBound(0))
            worker.ReportProgress(x, "Line by Line")

            firstword = TrimAfter(ss(i) + " ", " ", True)

            ' in some cases, we what 2 word phrases. So grab them when needed.
            If firstword = "End" Or firstword = "Else" Or firstword = "Public" Or firstword = "Private" Then
                n = (ss(i) & "  ").IndexOf(" ", firstword.Length + 1)
                If ss(i).Length >= n Then
                    firstword = ss(i).Substring(0, n).Trim
                Else
                    firstword = ss(i).Trim
                End If
            End If

            ' need to also trim off any trailing comment. This is needed to identify single line IF statements.
            If ss(i) = Nothing Then
                ss(i) = ""
            Else
                n = (ss(i).ToUpper & " ").IndexOf(" THEN ")
                If n > -1 Then
                    n = (ss(i).ToUpper & " ").IndexOf(" '", n)
                    If n > -1 Then ss(i) = ss(i).Substring(0, n).Trim
                End If
            End If
            ' -----------------------------------------------------------------------------------
            If firstword.StartsWith("Messagebox") Or firstword.StartsWith("MsgBox") Then
                ' Move this to case below
            End If


            Select Case firstword.ToUpper
                Case "'"   ' This line is a comment
                    ncomments = ncomments + 1
                    '       LastlineComment = i   ' Note, this points to the array element, not the file line

                Case "CASE"
                    ' Looks for code the implements case insensitiveity
                    CheckForCaseInsensitivity(SSummary, ss(i), lineno(i))

                Case "CATCH"

                Case "DIM"    ' Tracks the declaration of variables
                    If ss(i).ToUpper.Contains(" AS INTEGER") Then
                        SSummary(EnSummary.VarInteger).n += 1
                        SSummary(EnSummary.VarInteger).Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf
                    End If
                    If ss(i).ToUpper.Contains(" AS DECIMAL") Or ss(i).ToUpper.Contains("DOUBLE") Then
                        SSummary(EnSummary.VarDecimal).n += 1
                        SSummary(EnSummary.VarDecimal).Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf
                    End If
                    If ss(i).ToUpper.Contains(" AS BOOLEAN") Then
                        SSummary(EnSummary.VarBoolean).n += 1
                        SSummary(EnSummary.VarBoolean).Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf
                    End If
                    If ss(i).ToUpper.Contains(" AS DATE") Then
                        SSummary(EnSummary.VarDate).n += 1
                        SSummary(EnSummary.VarDate).Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf
                    End If
                    If ss(i).ToUpper.Contains(" AS STRING") Then
                        SSummary(EnSummary.VarString).n += 1
                        SSummary(EnSummary.VarString).Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf
                    End If

                    If ss(i).ToUpper.Contains(") AS ") Then
                        SSummary(EnSummary.VarArrays).n += 1
                        SSummary(EnSummary.VarArrays).Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf
                    End If

                    If ss(i).ToUpper.Contains("LIST (OF") Then
                        SSummary(EnSummary.VarLists).n += 1
                        SSummary(EnSummary.VarLists).Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf
                    End If

                    If ss(i).ToUpper.Contains(" STRUCTURE ") Then
                        SSummary(EnSummary.VarStructures).n += 1
                        SSummary(EnSummary.VarStructures).Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf
                    End If

                    If ss(i).ToUpper.Contains(" AS STREAMREADER") Or ss(i).ToUpper.Contains(" AS NEW STREAMREADER") Then
                        SSummary(EnSummary.LogicStreamReader).n += 1
                        SSummary(EnSummary.LogicStreamReader).Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf

                        s = returnBetween(ss(i), "DIM ", "AS STREAMREADER", True).Trim

                        If srName.Length > 0 Then
                            ReDim Preserve srName(srName.Count)
                            srName(srName.GetUpperBound(0)) = s
                        End If

                    End If
                    If ss(i).ToUpper.Contains(" AS STREAMWRITER") Or ss(i).ToUpper.Contains(" AS NEW STREAMWRITER") Then
                        SSummary(EnSummary.LogicStreamWriter).n += 1
                        SSummary(EnSummary.LogicStreamWriter).Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf

                        s = returnBetween(ss(i), "DIM ", "AS STREAMWRITER", True).Trim

                        If swName.Length > 0 Then
                            ReDim Preserve swName(swName.Count)
                            swName(swName.GetUpperBound(0)) = s
                        End If

                    End If


                    ' The following tracks the opening of streamreaders and writers. Tracking closings is done above.



                Case "DO"
                    SSummary(EnSummary.LogicDo).n += 1

                    ' accept a comment before or on the same line as the DO
                    With SSummary(EnSummary.CommentDo)
                        .n += 1
                        If ss(i - 1).StartsWith("'") Or ss(i).Contains("'") Then
                            ' build the Good list
                            If .good = Nothing Then
                                .good &= "(" & lineno(i).ToString & ")"
                            Else
                                .good &= ", (" & lineno(i).ToString & ")"
                            End If
                        Else
                            ' build the bad list
                            If .bad = Nothing Then
                                .bad &= "(" & lineno(i).ToString & ")"
                            Else
                                .bad &= ", (" & lineno(i).ToString & ")"
                            End If
                            .cssClass = "itemred"
                            .cnt += 1
                        End If
                    End With

                    ' --------------------------------------------
                    ' Check for complex conditions
                    CheckForComplexConditions(SSummary, ss(i), lineno(i))

                    ' Looks for code the implements case insensitiveity
                    CheckForCaseInsensitivity(SSummary, ss(i), lineno(i))

                Case "ELSE"
                    SSummary(EnSummary.LogicElse).n += 1
                    SSummary(EnSummary.LogicElse).Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf

                Case "ELSEIF"
                    SSummary(EnSummary.LogicElseIF).n += 1
                    SSummary(EnSummary.LogicElseIF).Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf

                    ' --------------------------------------------
                    ' Check for complex conditions
                    CheckForComplexConditions(SSummary, ss(i), lineno(i))

                    ' Looks for code the implements case insensitiveity
                    CheckForCaseInsensitivity(SSummary, ss(i), lineno(i))

                Case "END FUNCTION"
                    '           isinFunction = False ' not really used. Potentially could track nested functions


                Case "END IF"
                    isInIf -= 1  ' used to track nested IF
                    If isInIf < 0 Then
                        '       Beep()
                    End If

                Case "END SELECT"

                Case "END SUB"
                    '             isinSub = False     ' not really used. Potentially could track nested subs
                    '   ChangeFormText()     ' not sure what changed here ????????????????????? jhg 
                Case "END TRY"

                Case "END WITH"


                Case "FOR"
                    SSummary(EnSummary.LogicFor).n += 1   ' counts number of For statements
                    isInFor += 1       ' tracks if we are in a For statement, to identify Nested For statements

                    ' check for nested For
                    If isInFor > 1 Then
                        With SSummary(EnSummary.LogicNestedFor)
                            .n += 1
                            If .n = 1 Then
                                .Status &= "(" & lineno(i).ToString & ")"
                            Else
                                .Status &= ", (" & lineno(i).ToString & ")"
                            End If
                        End With
                    End If

                    ' ---------------------------------------------

                    ' Look for proper commenting of FOR statement
                    ' It should be before the for statement. It also accepts it if on the same line. 
                    With SSummary(EnSummary.CommentFor)
                        .n += 1
                        If ss(i - 1).StartsWith("'") Or ss(i).Contains("'") Then

                            If .good = Nothing Then
                                .good &= "(" & lineno(i).ToString & ")"
                            Else
                                .good &= ", " & lineno(i).ToString & ")"
                            End If
                        Else
                            If .bad = Nothing Then
                                .bad &= "(" & lineno(i).ToString & ")"
                            Else
                                .bad &= ", (" & lineno(i).ToString & ")"
                            End If
                            .cssClass = "itemred"
                            .cnt += 1
                        End If
                    End With


                Case "IF"

                    SSummary(EnSummary.LogicIF).n += 1
                    isInIf += 1
                    ' Ignorore single line IF Statements. THE Is InIF counter is backed out
                    If ss(i).Trim.ToUpper.Contains(" THEN ") Then isInIf -= 1

                    ' checks for nested IF statements
                    If isInIf > 1 Then
                        With SSummary(EnSummary.LogicNestedIF)
                            .n += 1
                            If .n = 1 Then
                                .Status &= "(" & lineno(i).ToString & ")"
                            Else
                                .Status &= ", (" & lineno(i).ToString & ")"
                            End If
                        End With
                    End If

                    ' ----------------------------------------
                    ' checks for comments preceeding IF statements
                    ' Acceptable comments are either preceeding or on the same line as IF statement.
                    With SSummary(EnSummary.CommentIF)
                        .n += 1
                        If ss(i - 1).StartsWith("'") Or ss(i).Contains("'") Then
                            If .good = Nothing Then
                                .good &= " (" & lineno(i).ToString & ")"
                            Else
                                .good &= ", (" & lineno(i).ToString & ")"
                            End If
                        Else
                            If .bad = Nothing Then
                                .bad &= " (" & lineno(i).ToString & ")"
                            Else
                                .bad &= ", (" & lineno(i).ToString & ")"
                            End If
                            .cssClass = "itemred"
                            .cnt += 1   ' counts up bad comments for IF statements
                        End If
                    End With

                    ' --------------------------------------------
                    ' Check for complex conditions
                    CheckForComplexConditions(SSummary, ss(i), lineno(i))

                    ' Looks for code the implements case insensitiveity
                    CheckForCaseInsensitivity(SSummary, ss(i), lineno(i))
                    ' ===============================================

                Case "IMPORTS"
                    ' Check Imports System.x

                    If ss(i).Contains("System.IO") Then
                        SSummary(EnSummary.SystemIO).n += 1
                        SSummary(EnSummary.SystemIO).Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf
                    Else
                        SSummary(EnSummary.SystemIO).Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - Not found</p>" & vbCrLf
                    End If

                    If ss(i).Contains("System).net") Then
                        SSummary(EnSummary.SystemNet).n += 1
                        SSummary(EnSummary.SystemNet).Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf
                    Else
                        SSummary(EnSummary.SystemNet).Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - Not found</p>" & vbCrLf
                    End If

                    If ss(i).Contains("System.DB") Then
                        SSummary(EnSummary.SystemDB).n += 1
                        SSummary(EnSummary.SystemDB).Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf
                    Else
                        SSummary(EnSummary.SystemDB).Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - Not found</p>" & vbCrLf
                    End If


                Case "LOOP"
                    ' --------------------------------------------
                    ' Check for complex conditions
                    CheckForComplexConditions(SSummary, ss(i), lineno(i))

                    ' Looks for code the implements case insensitiveity
                    CheckForCaseInsensitivity(SSummary, ss(i), lineno(i))

                Case "MESSAGEBOX.SHOW", "MSGBOX"                ' This needs to be fixed - jhg
                    With SSummary(EnSummary.LogicMessageBox)
                        .n += 1

                        If (ss(i - 1).StartsWith("MSGBOX")) Then
                            .bad &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & TrimAfter(ss(i), "(", True) & "</p>" & vbCrLf
                            .cssClass = "itemred"
                        Else
                            .good &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & TrimAfter(ss(i), "(", True) & "</p>" & vbCrLf
                        End If
                    End With


                Case "NEXT"
                    isInFor -= 1

                Case "OPTION STRICT"
                    ' This is not validated in the file itself. It is checked in the application config.

                Case "PUBLIC STRUCTURE", "PRIVATE STRUCTURE"
                    SSummary(EnSummary.VarStructures).n += 1
                    SSummary(EnSummary.VarStructures).Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf

                Case "SELECT"
                    SSummary(EnSummary.LogicSelectCase).n += 1

                    ' check for comment. Accept on line before or on the same line.
                    With SSummary(EnSummary.CommentSelect)
                        If ss(i - 1).StartsWith("'") Or ss(i).Contains("'") Then
                            If .good = Nothing Then
                                .good &= "(" & lineno(i).ToString & ")"
                            Else
                                .good &= ", (" & lineno(i).ToString & ")"
                            End If
                        Else
                            If .bad = Nothing Then
                                .bad &= "(" & lineno(i).ToString & ")"
                            Else
                                .bad &= ", (" & lineno(i).ToString & ")"
                            End If
                            .cssClass = "itemred"
                            .n += 1
                        End If
                    End With

                    ' --------------------------------------------
                    ' Looks for code the implements case insensitiveity
                    CheckForCaseInsensitivity(SSummary, ss(i), lineno(i))

                Case "SUB", "PUBLIC SUB", "PRIVATE SUB", "FUNCTION", "PUBLIC FUNCTION", "PRIVATE FUNCTION"

                    SSummary(EnSummary.LogicSub).n += 1

                    ' Check for a comment in first line of sub/function. This accepts it as the previous line
                    If i < ss.GetUpperBound(0) Then
                        With SSummary(EnSummary.CommentSub)
                            .n += 1
                            If (ss(i - 1).StartsWith("'") Or ss(i + 1).StartsWith("'")) Then
                                .bad &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & TrimAfter(ss(i), "(", True) & "</p>" & vbCrLf
                                .cssClass = "itemred"
                                .cnt += 1
                            Else
                                .good &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & TrimAfter(ss(i), "(", True) & "</p>" & vbCrLf
                                .n += 1
                            End If
                        End With

                    End If

                    ' ------------------------------------------
                    ' check for optional parameters
                    With SSummary(EnSummary.LogicOptional)
                        If ss(i).Contains(" Optional ") Then
                            .n += 1
                            If .good = Nothing Then
                                .good &= bullet & " ( " & lineno(i).ToString & ") " & TrimAfter(ss(i), "(", True) & " - has an Optional Parameter"
                            Else
                                .good &= "<br> " & bullet & " ( " & lineno(i).ToString & ") " & TrimAfter(ss(i), "(", True) & " - has an Optional Parameter"
                            End If
                        End If

                        '    .Status &= TrimAfter(ss(i), "(", True) & " defined in line (" & lineno(i).ToString & ") has an Optional Parameter" & " </p>" & vbCrLf
                        '    .n += 1
                        'End If
                    End With

                    ' -------------------------------------
                    ' check for byref parameters
                    With SSummary(EnSummary.LogicByRef)
                        If ss(i).Contains("ByRef ") Then
                            .n += 1
                            If .good = Nothing Then
                                .good &= bullet & " ( " & lineno(i).ToString & ") " & TrimAfter(ss(i), "(", True) & " - has a ByRef Parameter"
                            Else
                                .good &= "<br> " & bullet & " ( " & lineno(i).ToString & ") " & TrimAfter(ss(i), "(", True) & " - has a ByRef Parameter"
                            End If
                        End If
                    End With

                Case "TRY"
                    ' just looks for the try. Should likely also look for the catch, but does not currently
                    SSummary(EnSummary.LogicTryCatch).n += 1
                    If SSummary(EnSummary.LogicTryCatch).n = 1 Then
                        SSummary(EnSummary.LogicTryCatch).Status &= "(" & lineno(i).ToString & ")"
                    Else
                        SSummary(EnSummary.LogicTryCatch).Status &= ", (" & lineno(i).ToString & ")"
                    End If

                    ' ========================================
                Case "WHILE"
                    SSummary(EnSummary.LogicWhile).n += 1

                    ' Comment While
                    With SSummary(EnSummary.CommentWhile)
                        .n += 1
                        If ss(i - 1).StartsWith("'") Then
                            If .good = Nothing Then
                                .good &= "(" & lineno(i).ToString & ")"
                            Else
                                .good &= ", (" & lineno(i).ToString & ")"
                            End If
                        Else
                            If .bad = Nothing Then
                                .bad &= "(" & lineno(i).ToString & ")"
                            Else
                                .bad &= ", (" & lineno(i).ToString & ")"
                            End If
                            .cssClass = "itemred"
                            .cnt += 1
                        End If
                    End With

                    ' --------------------------------------------
                    ' Check for complex conditions
                    CheckForComplexConditions(SSummary, ss(i), lineno(i))

                    ' Looks for code the implements case insensitiveity
                    CheckForCaseInsensitivity(SSummary, ss(i), lineno(i))


                Case "WITH"

            End Select

            doneflag = False
            For j = 1 To srName.GetUpperBound(0)
                If ss(i).ToUpper.Contains(srName(j).ToUpper.Trim & ".CLOSE") And Not doneflag Then
                    SSummary(EnSummary.LogicStreamReaderClose).n += 1
                    SSummary(EnSummary.LogicStreamReaderClose).Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & " </p>" & vbCrLf
                    doneflag = True
                End If
            Next j

            doneflag = False
            For j = 1 To swName.GetUpperBound(0)

                If ss(i).ToUpper.Contains(swName(j).ToUpper.Trim & ".CLOSE") And Not doneflag Then
                    SSummary(EnSummary.LogicStreamWriterClose).n += 1
                    SSummary(EnSummary.LogicStreamWriterClose).Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & " </p>" & vbCrLf
                    doneflag = True
                End If
            Next j

            ' =================================================================================
            ' Now we need to look at the contents of the whole line to check for other issues.

            ' Check for the FormLoad Method
            If ss(i).Contains("Handles MyBase.Load") Then
                SSummary(EnSummary.LogicFormLoad).n += 1
                SSummary(EnSummary.LogicFormLoad).good = "<span class=""hangingindent2"">" & bullet & " (" & lineno(i).ToString & ") - " & ss(i) & "</span>" & vbCrLf
            End If

            ' -----------------------------------
            ' Check for String Concatination
            With SSummary(EnSummary.LogicConcatination)
                If ss(i) <> Nothing AndAlso firstword <> "'" AndAlso ss(i).Contains(" & ") Then
                    .n += 1
                    If .n = 1 Then
                        .Status &= "(" & lineno(i).ToString & ")"
                    Else
                        .Status &= ", (" & lineno(i).ToString & ")"
                    End If
                End If
            End With

            ' ------------------------------------------------------------------------------
            ' handling string conversions and formatting
            If Not ss(i).StartsWith("'") Then   ' Ignore comments
                If ss(i).Contains("CStr(") Or ss(i).Contains(".ToString") Then
                    SSummary(EnSummary.LogicConvertToString).n += 1
                    If SSummary(EnSummary.LogicConvertToString).n = 1 Then
                        SSummary(EnSummary.LogicConvertToString).Status &= "(" & lineno(i).ToString & ")"
                    Else
                        SSummary(EnSummary.LogicConvertToString).Status &= ", (" & lineno(i).ToString & ")"
                    End If
                End If

                ' Check to see if code is using a format string with tostring
                With SSummary(EnSummary.LogicToStringFormat)
                    If ss(i).Contains(".ToString(") Then
                        .n += 1 ' This counts how many times they used the formating feature
                        If .n = 1 Then
                            .Status &= "(" & lineno(i).ToString & ")"
                        Else
                            .Status &= ", (" & lineno(i).ToString & ")"
                        End If
                    End If
                End With

                ' check the String.format approach

                If ss(i).Contains(".Format(") Then
                    With SSummary(EnSummary.LogicStringFormat)
                        .n += 1 ' This counts the number of string.format commands
                        If .n = 1 Then
                            .Status &= "(" & lineno(i).ToString & ")"
                        Else
                            .Status &= ", (" & lineno(i).ToString & ")"
                        End If
                    End With

                    ' Also try to determine if there are parameters. This is done by seeing if there is at least one comma in the parameter list
                    s = TrimUpTo(ss(i), "Format(")
                    If s.Contains(",") Then
                        SSummary(EnSummary.LogicStringFormatParameters).n += 1
                        SSummary(EnSummary.LogicStringFormatParameters).Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") " & ss(i) & "</p>" & vbCrLf
                    End If
                End If

            End If

            '
            ' ------------------------------------------------------------------------------
        Next i    ' end of processing each line
        ' ==================================================================================

        ' ============================================================
        ' NOW LOOK THROUGH ASSESSMENTS AND SET THE cssClass
        ' -------------------------------------------------------------------------------------------------------
        ' Process Comments
        ProcessComment(fn, SSummary(EnSummary.LogicSub), SSummary(EnSummary.CommentSub), "Subs / Functions", "CommentSubs")
        ProcessComment(fn, SSummary(EnSummary.LogicIF), SSummary(EnSummary.CommentIF), "IF statements", "CommentIF")
        ProcessComment(fn, SSummary(EnSummary.LogicFor), SSummary(EnSummary.CommentFor), "FOR statements", "CommentFOR")
        ProcessComment(fn, SSummary(EnSummary.LogicDo), SSummary(EnSummary.CommentDo), "DO statements", "CommentDO")
        ProcessComment(fn, SSummary(EnSummary.LogicWhile), SSummary(EnSummary.CommentWhile), "WHILE statements", "CommentWHILE")
        ProcessComment(fn, SSummary(EnSummary.LogicSelectCase), SSummary(EnSummary.CommentSelect), "SELECT CASE statements", "CommentSELECT")
        ' ------------------------------------------------------------------------------------------------------------

        If SSummary(EnSummary.LogicCStr).n > 0 Or SSummary(EnSummary.LogicToString).n > 0 Then
            SSummary(EnSummary.LogicToString).Status &= "<span class=""boldtext"">Converting to Strings </span><br>" & vbCrLf

            If SSummary(EnSummary.LogicCStr).n > 0 Then
                SSummary(EnSummary.LogicToString).Status &= "<span class=""boldtext""><br>Using CStr() to convert to string </span><br>" & vbCrLf
            End If

            If SSummary(EnSummary.LogicToString).n > 0 Then
                SSummary(EnSummary.LogicToString).Status &= "<span class=""boldtext""><br>Using .ToString to convert to string </span><br>" & vbCrLf
            End If
        End If
        ' ------------------------------------------------------------------------------------------------------------

        If SSummary(EnSummary.LogicToStringFormat).n > 0 Or SSummary(EnSummary.LogicStringFormat).n > 0 Then
            If SSummary(EnSummary.LogicToStringFormat).n > 0 Then
                SSummary(EnSummary.LogicToStringFormat).Status &= "<span class=""boldtext""><br>Included .ToString(format) command </span><br>" & vbCrLf
            End If
            If SSummary(EnSummary.LogicStringFormat).n > 0 Then
                SSummary(EnSummary.LogicStringFormat).Status &= "<span class=""boldtext""><br>Included String.Format(Template)  </span><br>" & vbCrLf
            End If
        End If
        ' ------------------------------------------------------------------------------------------------------------
        ' -------------------------------------------------------------------------------------------------------
        ' AppInfo
        ProcessReq(fn, SSummary(EnSummary.InfoAppTitle), "Application Title not modified", "Application Title modified", "InfoAppTitle")
        ProcessReq(fn, SSummary(EnSummary.InfoDescription), "Application Description not modified", "Application Description modified", "InfoDescription")
        ProcessReq(fn, SSummary(EnSummary.InfoCompany), "Application Company Info not modified", "Application Company Info modified", "InfoCompany")
        ProcessReq(fn, SSummary(EnSummary.InfoProduct), "Application Product Info not modified", "Application Product Info modified", "InfoProduct")
        ProcessReq(fn, SSummary(EnSummary.InfoTrademark), "Application Trademark not modified", "Application Trademark modified", "InfoTrademark")
        ProcessReq(fn, SSummary(EnSummary.InfoCopyright), "Application Copyright not modified", "Application Copyright modified", "InfoCopyright")

        'Compile Options
        'ProcessReq(.OptionStrict, "No System.IO found", "Imports System.IO found (Line No.)", "Include System.IO")
        'ProcessReq(.SystemNet, "No System).net found", "Imports System).net found (Line No.)", "Include System).net")

        ' Form Design


        ' Imports
        ProcessReq(fn, SSummary(EnSummary.SystemIO), "No System.IO found", "Imports System.IO found (Line No.)", "SystemIO")
        ProcessReq(fn, SSummary(EnSummary.SystemNet), "No System).net found", "Imports System).net found (Line No.)", "SystemNet")
        ProcessReq(fn, SSummary(EnSummary.SystemDB), "No System.DB found", "Imports System.DB found (Line No.)", "SystemDB")

        ' Vars
        ProcessReq(fn, SSummary(EnSummary.VarArrays), "No Arrays declared", "Data Arrays declared", "VarArrays")
        ProcessReq(fn, SSummary(EnSummary.VarLists), "No Lists declared", "Lists(of T) declared", "VarLists")
        ProcessReq(fn, SSummary(EnSummary.VarStructures), "No Structures declared", "Data Structures defined", "VarStructures")
        ProcessReq(fn, SSummary(EnSummary.VarString), "No String variables declared", "String variables declared", "VarString")
        ProcessReq(fn, SSummary(EnSummary.VarInteger), "No Integer variables declared", "Integer variables declared", "VarInteger")
        ProcessReq(fn, SSummary(EnSummary.VarDecimal), "No Decimal / Double variables declared", "Decimal / Double variables declared", "VarDecimal")
        ProcessReq(fn, SSummary(EnSummary.VarDate), "No Date variables declared", "Date variables declared", "VarDate")
        ProcessReq(fn, SSummary(EnSummary.VarBoolean), "No Boolean variables declared", "Boolean variables declared", "VarBoolean")

        ' Logic
        ProcessReq(fn, SSummary(EnSummary.LogicWhile), "No While statements found", "While statements found (Line No.)", "LogicWHILE")
        ProcessReq(fn, SSummary(EnSummary.LogicSelectCase), "No SelectCase statements found", "Select Case statements found (Line No.)", "LogicSelectCase")

        ProcessReq(fn, SSummary(EnSummary.LogicConvertToString), "No Conversion to String (.CStr or .toString) found", "Conversion to String (.CStr or .toString) found (Line No.)", "LogicConvertToString")
        ProcessReq(fn, SSummary(EnSummary.LogicStringFormat), "No String Formatting (.toString() or String.Format()) found", "String Formatting (.toString() or String.Format()) found (Line No.)", "LogicStringFormat")
        ProcessReq(fn, SSummary(EnSummary.LogicStringFormatParameters), "No Parameterized string formats (string.format(f,{0}) found", "Parameterized String Formatting (string.format(f,{0}) found (Line No.)", "LogicStringFormatParameters")

        ProcessReq(fn, SSummary(EnSummary.LogicConcatination), "No String Concatination found", "String Concatination found (Line No.)", "LogicConcatination")
        ProcessReq(fn, SSummary(EnSummary.LogicCaseInsensitive), "No Case-Insensitive comparisons found", "Case-Insensitive comparisons found (Line No.)", "LogicCaseInsensitive")
        ProcessReq(fn, SSummary(EnSummary.LogicComplexConditions), "No Complex Conditions (AND / OR / ANDALSO / ORELSE) found", "Complex Conditions (AND / OR / ANDALSO / ORELSE) found (Line No.)", "LogicComplexConditions")

        ProcessReq(fn, SSummary(EnSummary.LogicElse), "No Else statements found", "Else statements found (Line No.)", "LogicElse")
        ProcessReq(fn, SSummary(EnSummary.LogicElseIF), "No ElseIF statements found", "ElseIF statements found (Line No.)", "LogicElseIF")
        ProcessReq(fn, SSummary(EnSummary.LogicNestedIF), "No Nested IF statements found", "Nested IF statements found (Line No.)", "LogicNestedIF")
        ProcessReq(fn, SSummary(EnSummary.LogicNestedFor), "No Nested FOR statements found", "Nested FOR statements found (Line No.)", "LogicNestedFOR")
        ProcessReq(fn, SSummary(EnSummary.LogicOptional), "No Optional Parameters found", "Optional Parameters found", "LogicOptional")
        ProcessReq(fn, SSummary(EnSummary.LogicByRef), "No ByRef Parameters found", "ByRef Parameters found", "LogicByRef")
        ProcessReq(fn, SSummary(EnSummary.LogicTryCatch), "No Try ... Catch Statement Found", "Try ... Catch Statements found (Line No.)", "LogicTryCatch")
        ProcessReq(fn, SSummary(EnSummary.LogicFormLoad), "No Form Load Method Found", "Form Load Method found", "LogicFormLoad")

        ProcessReq(fn, SSummary(EnSummary.LogicStreamReader), "No StreamReaders found", "StreamReader found (Line No.)", "LogicStreamReader")
        ProcessReq(fn, SSummary(EnSummary.LogicStreamReaderClose), "No StreamReader.Close found", "StreamReader.Close found (Line No.)", "LogicStreamReaderClose")
        ProcessReq(fn, SSummary(EnSummary.LogicStreamWriter), "No StreamWriters found", "StreamWriters found (Line No.)", "LogicStreamWriter")
        ProcessReq(fn, SSummary(EnSummary.LogicStreamWriterClose), "No StreamWriter.Close found", "StreamWriter.Close found (Line No.)", "LogicStreamWriterClose")

        Return ncomments
    End Function


    Sub ProcessComment(filename As String, ByRef logictype As MyItems, ByRef commenttype As MyItems, construct As String, nm As String)
        If logictype.n = 0 Then
            logictype.Status = "<p><span class=""boldtext"">" & filename & " - No " & construct & " Found"
            commenttype.Status = "<p><span class=""boldtext"">" & filename & " - No " & construct & " Found" & "</p>" & vbCrLf
            'logictype.cssClass = "itemred"
            'commenttype.cssClass = "itemred"
        Else
            If commenttype.bad <> Nothing Then
                commenttype.Status = "<p><span class=""boldtext"">" & filename & " - " & construct & " Missing Descriptive Comments (Line No.)</span><br>" & vbCrLf & commenttype.bad & " </p> " & vbCrLf
                'logictype.cssClass = "itemred"
                'commenttype.cssClass = "itemred"
            End If
            If commenttype.good <> Nothing Then
                If logictype.bad = Nothing Then
                    'logictype.cssClass = "itemgreen"
                    'commenttype.cssClass = "itemgreen"
                End If
                commenttype.Status &= "<p><span class=""boldtext"">" & filename & " - " & construct & " Includes Descriptive Comments (Line No.)</span><br>" & vbCrLf & commenttype.good & " </p>" & vbCrLf
            End If
        End If

        ' nm = logictype.ToString.Substring(1)

        If Not Find_Setting(nm, "ProcessComment").Req Then
            logictype.cssClass = "itemclear"
            commenttype.cssClass = "itemclear"
        End If

    End Sub

    Sub ProcessReq(filename As String, ByRef type As MyItems, NotFound As String, Found As String, nm As String)
        If type.n = 0 Then
            type.Status = "<p><span class=""boldtext"">" & filename & "</span> - " & NotFound & "</p>" & vbCrLf
            '   type.cssClass = "itemred"
        Else
            type.Status = "<p> <span class=""boldtext"">" & filename & "</span> - " & Found & "<br>" & vbCrLf & type.good & "</p>" & vbCrLf
            '     type.cssClass = "itemgreen"
            '  type.n += 1
        End If

        ' nm = type.ToString.Substring(1)

        If Not Find_Setting(nm, "ProcessReq").Req Then
            type.cssClass = "itemclear"
        End If

    End Sub


    Sub CheckForComplexConditions(ByRef ssum() As MyItems, s As String, lineno As Integer)
        ' This checks a line of code for evidence of complex consitions
        With ssum(EnSummary.LogicComplexConditions)
            If s.ToUpper.Contains(" AND ") Or s.ToUpper.Contains(" OR ") Then
                .n += 1
                .Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno.ToString & ") - " & s & "</p>" & vbCrLf
                .cssClass = "itemgreen"
            End If
        End With
    End Sub

    Sub CheckForCaseInsensitivity(ByRef ssum() As MyItems, s As String, lineno As Integer)
        ' This check the line of code for evidence of case insensitivity

        With ssum(EnSummary.LogicCaseInsensitive)
            If s.ToUpper.Contains(".TOUPPER") Or s.ToUpper.Contains(".tolower") Then
                .n += 1
                .Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno.ToString & ") - " & s & "</p>" & vbCrLf
                .cssClass = "itemgreen"
            End If
        End With
    End Sub

    Function CheckForSplashScreen2(ByRef filesinbuild As List(Of String), ByRef item As MyItems, RemoveFromList As Boolean) As Boolean
        Dim s As String
        With item
            .Status = vbFalse.ToString

            For Each filename As String In filesinbuild
                Dim sr As New StreamReader(filename)
                s = sr.ReadToEnd
                sr.Close()

                If s.Contains("splash screen") Or s.Contains("SplashScreen") Then
                    ' has a splash screen
                    .Status = vbTrue.ToString
                    .n += 1
                    If RemoveFromList Then filesinbuild.Remove(filename)
                    Exit For
                End If
            Next
        End With
    End Function


    Sub CheckFormProperties2(filename As String, ByRef AppForm() As MyItems)
        Dim s As String = ""
        Dim ss As String = ""
        Dim cArray() As String
        Dim delim() As String = {"CType(CType("}
        Dim tmp As String = ""
        Dim fn As String = ""

        With AppForm


            tmp = filename.Replace(".vb", ".designer.vb")

            ' Do not process unless it has a .designer.vb file. This eliminates processing Modules and classes
            If File.Exists(tmp) Then
                fn = ReturnLastField(filename, "\")

                Dim sr As New StreamReader(tmp)
                s = sr.ReadToEnd
                sr.Close()

                If s.Contains("Me.BackColor") Then
                    If s.Contains("System.Drawing.SystemColors.") Then
                        AppForm(EnForm.FormBackColor).Status = "<p><span class=""boldtext"">" & fn & "</span> - Form color = " & TrimUpTo(ss, "System.Drawing.SystemColors.") & "</p>"
                        AppForm(EnForm.FormBackColor).cssClass = "itemgreen"
                        AppForm(EnForm.FormBackColor).n += 1
                    ElseIf s.Contains("System.Drawing.Color.FromArgb(") Then

                        ss = returnBetween(s, "Me.BackColor =", vbCrLf)
                        cArray = ss.Split(delim, StringSplitOptions.None)
                        Try
                            AppForm(EnForm.FormBackColor).Status = "<p><span class=""boldtext"">" & fn & "</span> - Form color is nongray (#" & ReturnHexEquivalent(TrimAfter(cArray(1), ",", True)) & ReturnHexEquivalent(TrimAfter(cArray(2), ",", True)) & ReturnHexEquivalent(TrimAfter(cArray(3), ",", True)) & ") </p>"
                            AppForm(EnForm.FormBackColor).cssClass = "itemgreen"
                            AppForm(EnForm.FormBackColor).n += 1
                        Catch
                            AppForm(EnForm.FormBackColor).Status = "Form color = " & ss
                        End Try
                    ElseIf s.Contains("System.Drawing.Color.") Then
                        AppForm(EnForm.FormBackColor).Status = "<p><span class=""boldtext"">" & fn & "</span> - Form color = " & returnBetween(s, "Me.BackColor = System.Drawing.Color.", vbCrLf) & "</p>"
                        AppForm(EnForm.FormBackColor).cssClass = "itemgreen"
                        AppForm(EnForm.FormBackColor).n += 1
                    End If
                Else
                    AppForm(EnForm.FormBackColor).Status = "<p> <span class=""boldtext"">" & fn & "</span> - Form Color was not changed at design time (still gray)" & "</p>"
                    AppForm(EnForm.FormBackColor).cssClass = "itemred"
                End If


                If s.Contains("Me.Text") Then
                    AppForm(EnForm.FormText).Status = "<p> <span class=""boldtext"">" & fn & "</span> - " & returnBetween(s, "Me.Text = """, """") & "</p>"
                    AppForm(EnForm.FormText).cssClass = "itemgreen"
                    AppForm(EnForm.FormText).n += 1
                Else
                    AppForm(EnForm.FormText).Status = "<p> <span class=""boldtext"">" & fn & "</span> - The form text was not changed." & "</p>"
                    AppForm(EnForm.FormText).cssClass = "itemred"

                End If

                If s.Contains("Me.StartPosition") Then
                    AppForm(EnForm.FormStartPosition).Status = "<p> <span class=""boldtext"">" & fn & "</span> - Form Startup Position set to: [" & returnBetween(s, "Me.StartPosition = System.Windows.Forms.FormStartPosition.", vbCrLf) & "] " & "</p>"
                    AppForm(EnForm.FormStartPosition).cssClass = "itemgreen"
                    AppForm(EnForm.FormStartPosition).n += 1
                Else
                    AppForm(EnForm.FormStartPosition).Status = "<p> <span class=""boldtext"">" & fn & "</span> - Form StartPosition not Modified." & "</p>"
                    AppForm(EnForm.FormStartPosition).cssClass = "itemred"
                End If


                ' accept and cancel button settings in *.resx file
                sr = New StreamReader(filename.Replace(".vb", ".resx"))
                s = sr.ReadToEnd
                sr.Close()
                If s.Contains("Me.AcceptButton") Then
                    AppForm(EnForm.FormAcceptButton).Status = "<p> <span class=""boldtext"">" & fn & "</span> - " & returnBetween(s, "Me.AcceptButton = ", vbCrLf) & "</p>"
                    AppForm(EnForm.FormAcceptButton).cssClass = "itemgreen"
                    AppForm(EnForm.FormAcceptButton).n += 1
                Else
                    AppForm(EnForm.FormAcceptButton).Status = "<p> <span class=""boldtext"">" & fn & "</span> - Accept Button Property not set at design time." & "</p>"
                    AppForm(EnForm.FormAcceptButton).cssClass = "itemred"
                End If

                If s.Contains("Me.CancelButton") Then
                    AppForm(EnForm.FormCancelButton).Status = "<p> <span class=""boldtext"">" & fn & "</span> - " & returnBetween(s, "Me.AcceptButton = ", vbCrLf) & "</p>"
                    AppForm(EnForm.FormCancelButton).cssClass = "itemgreen"
                    AppForm(EnForm.FormCancelButton).n += 1
                Else
                    AppForm(EnForm.FormCancelButton).Status = "<p> <span class=""boldtext"">" & fn & "</span> - Cancel Button Property not set at design time." & "</p>"
                    AppForm(EnForm.FormCancelButton).cssClass = "itemred"
                End If
            End If

        End With
    End Sub

    Function ReturnHexEquivalent(s As String) As String
        Dim h As String
        h = Hex(CInt(s))
        If h.Length = 1 Then h = "0" & h
        Return h
    End Function




    Sub CheckForAboutBox2(ByRef filesinbuild As List(Of String), ByRef item As MyItems, RemoveFromList As Boolean)
        Dim s As String
        With item
            .Status = vbFalse.ToString

            For Each filename As String In filesinbuild
                Dim sr As New StreamReader(filename)
                s = sr.ReadToEnd
                sr.Close()

                If s.Contains("About Box") Or s.Contains("AboutBox") Then
                    ' has a splash screen"
                    .Status = "<p> <span class=""boldtext"">" & filename & " - Is an About Box. This file will not be considered further." & "</p>"
                    .n += 1
                    If RemoveFromList Then filesinbuild.Remove(filename)
                    Exit For
                End If
            Next
        End With
    End Sub



    Sub CheckForModules2(filesinbuild As List(Of String), ByRef item As MyItems)
        Dim source As String = ""


        With item
            .Comments = ""

            For Each filename As String In filesinbuild
                If GetFileSource(filename, source) Then
                    If source.Contains("End Module") Then
                        .n += 1
                        .Comments = "Filename " & ReturnLastField(filename, "\") & " was identified as a Module"
                    End If
                End If
            Next filename
            ' ------------------------------------------------
            .Status = .n.ToString
        End With

    End Sub


    Sub CheckObjectNaming2(filename As String, ByRef AppForm() As MyItems)
        ' ----------------------------------------------------------
        ' Check the form design to see if objects on form are named properly
        ' ----------------------------------------------------------
        Dim s1 As String = ""
        Dim s2 As String = ""
        Dim s3 As String = ""
        Dim statgood As String = ""
        Dim statbad As String = ""
        Dim cnt As Integer

        Dim foundflag As Boolean

        Dim arrow As String = "- "
        Dim tmpObj As String = ""

        Dim delim() As String = {"        Me."}
        ' here are all the objects I am interested in
        Dim strObjects() As String = {"LabelActive", "LabelNonactive", "Button", "Textbox", "Listbox", "Combobox", "OpenFileDialog", "SaveFileDialog", "RadioButton", "CheckBox", "GroupBox", "WebBrowser", "WebClient", "MaskedTextBox", "PictureBox", "TabControl", "Timer"}

        ' Prefixes based on http://www.vbprogramming.org/vbbook.php?sect=se7ss1&ch=ch03&PHPSESSID=3f8ae98f461e2db49db57bd6f014e3c5
        ' Note, many sources don't refer to a RadioButton, but rather OptionButton
        Dim strObjPrefix() As String = {"lbl", "-", "btn", "txt", "lst", "cbo", "ofd", "sfd", "rbn", "chk", "grp", "wb", "wc", "msk", "pic", "tab", "tmr"}
        Dim objNbr As Integer

        Dim filesourceVB As String
        Dim filesource As String
        Dim IsActiveLabel As Boolean
        Dim tmp As String = ""
        Dim nObjSeen As Integer
        ' ================================================================================================

        ' read the source of the vb file to check to see if any of the objects not renamed are used in the source file
        Dim sr As New StreamReader(filename)

        ' Now read the source of the designer file to extract the definition of the objects.
        sr = New StreamReader(filename)
        filesourceVB = sr.ReadToEnd
        sr.Close()

        ' We now open the form design file to look at the objects
        Dim fn As String = filename.Replace(".vb", ".designer.vb")

        ' Check to see if the file exists. This avoids processing Modules and Classes.
        If File.Exists(fn) Then
            sr = New StreamReader(fn)
            filesource = sr.ReadToEnd
            sr.Close()


            fn = ReturnLastField(filename, "\")

            Dim strLayout As String = returnBetween(filesource, "Private Sub InitializeComponent()", "Me.SuspendLayout()")
            Dim strObjOnForm() As String = Nothing
            Dim strobj As String = ""
            Dim css As String = ""

            ' extract out all the objects on the form
            Try
                strObjOnForm = strLayout.Split(delim, StringSplitOptions.None)
                strObjOnForm.DropFirstElement()

            Catch ex As Exception
                MessageBox.Show("Error occured in CheckObjectNaming - " & ex.Message)
            End Try

            ' ----------------------------------------------------------------------------------------------------------------
            ' the code below is wrong. it looks at the file name, not the FormText  <<<<<< jhg
            tmp = ReturnLastField(filename, "\")

            With AppForm(EnForm.FormName)
                If tmp.ToLower.StartsWith("frm") Then
                    .Status = tmp & " starts with frm prefix"
                    .cssClass = "itemgreen"
                    .n += 1
                Else
                    .Status = tmp & " does not start with the frm prefix"
                    .cssClass = "itemred"
                    .cnt += 1
                    .n += 1
                End If
            End With
            ' ----------------------------------------------------------------------------------------------------------------
            For j As Integer = 0 To strObjects.GetUpperBound(0) ' Look through the list of object types we are interested in
                strobj = strObjects(j)
                foundflag = False ' indicates if the object on the form is in the list of interesting objects

                statgood = ""
                statbad = ""
                cnt = -1  ' indicates no instances yet
                tmpObj = strobj

                If strobj.StartsWith("Label") Then ' We are doing this to handle Active and Nonactive labels
                    strobj = "Label"
                End If



                ' Not look through all the objects on the form and process them if they match the current object in outer loop
                nobjseen = 0
                For Each obj As String In strObjOnForm   ' list of objects on the form

                    If obj.Contains("= New System.Windows.Forms.") Then  ' this is an object, so precees it.
                        Try
                            s1 = TrimUpTo(obj, "= New System.Windows.Forms.")
                            If s1.Contains("(") Then s1 = s1.Substring(0, s1.IndexOf("("))
                            ' s1 is the object on the form

                            ' compare with current object, and process if the same.
                            If s1.Trim.ToUpper = strobj.ToUpper Then
                                foundflag = True
                                nObjSeen += 1

                                s2 = TrimAfter(obj, " = ", True)   ' not sure if this is correct. I added it so s2 had an initial value.
                                '  If s2.Contains(" = ") Then s2 = obj.Substring(0, obj.IndexOf(" = "))

                                ' Extracts out any text associated with the object. This text is displayed in the output
                                If filesource.IndexOf("Me." & s2 & ".Text = ") > -1 Then
                                    s3 = returnBetween(filesource, "Me." & s2 & ".Text = ", vbCrLf)                   ' Text Value
                                    If s3.Length > 30 Then s3 = s3.Substring(0, 30) & " ..."
                                Else
                                    s3 = ""
                                End If
                                ' -----------------------------------------------------------------------------------------------
                                If cnt = -1 Then cnt = 0 ' reset it so we count those without proper prefixes. cnt = 0 means no errors

                                ' Handle Labels specially. Needed to distinguish Active and NonActive labels

                                If tmpObj = "LabelActive" Then ' only bother with this if we are processing label objects
                                    If s1 = "Label" Then    ' special check to see if the object on the form  is a label
                                        If filesourceVB.IndexOf(s2) > -1 Then   ' look in the source and see if it is referenced
                                            IsActiveLabel = True
                                        Else
                                            IsActiveLabel = False
                                        End If
                                        ' --------------------------------------------------------------------------------------------------
                                        If IsActiveLabel Then
                                            ' determine if the label name starts with lbl for active labels
                                            If s2.StartsWith(strObjPrefix(j)) Then                  ' Name of the object
                                                statgood &= " <br>" & bullet & s2 & vbCrLf
                                            Else
                                                ' this is a problem - object not renamed
                                                statbad &= " <br>" & arrow & s2 & vbCrLf
                                                cnt = cnt + 1
                                            End If
                                        End If
                                    End If
                                ElseIf tmpObj = "LabelNonactive" Then
                                    If s1 = "Label" Then statgood &= " <br>" & bullet & s2 & vbCrLf
                                Else  ' non labels
                                    If s2.StartsWith(strObjPrefix(j)) Then                  ' Name of the object
                                        statgood &= " <br>" & bullet & s2 & vbCrLf
                                    Else
                                        ' this is a problem - object not renamed
                                        statbad &= arrow & s2 & " <br>" & vbCrLf
                                        cnt = cnt + 1
                                    End If
                                End If

                            End If

                        Catch ex As Exception
                            '     Beep()
                            MessageBox.Show("Error occured in Check Object Naming - " & ex.Message)
                        End Try
                    End If

                    If s1.Trim.ToUpper = strobj.ToUpper Then
                        Exit For  ' since we found a match, no need to check the other items in the list
                    End If
                Next obj

                If Not foundflag Then statgood = "<p> <span class=""boldtext"">" & fn & "</span> - None Found </p>"

                objNbr = -1
                If cnt = -1 Then cnt = 0 ' 
                Select Case tmpObj
                    Case "Button"
                        objNbr = EnForm.ObjButton
                    Case "Textbox"
                        objNbr = EnForm.ObjTextbox
                    Case "Listbox"
                        objNbr = EnForm.ObjListbox
                    Case "Combobox"
                        objNbr = EnForm.ObjCombobox
                    Case "RadioButton"
                        objNbr = EnForm.ObjRadioButton
                    Case "CheckBox"
                        objNbr = EnForm.ObjCheckbox
                    Case "GroupBox"
                        objNbr = EnForm.ObjGroupBox
                    Case "OpenFileDialog"
                        objNbr = EnForm.ObjOpenFileDialog
                    Case "SaveFileDialog"
                        objNbr = EnForm.ObjSaveFileDialog
                    Case "WebBrowser"
                        objNbr = EnForm.ObjWebBrowser
                    Case "Label"
                        objNbr = -1
                        AppForm(EnForm.objLabel).Status = buildObjSummary(fn, strobj, cnt, nObjSeen, statgood, statbad, css)
                        AppForm(EnForm.objLabel).cssClass = css
                        AppForm(EnForm.objLabel).cnt = cnt
                        AppForm(EnForm.objLabel).n = nObjSeen
                    Case "LabelActive"
                        objNbr = -1
                        AppForm(EnForm.ObjActiveLabel).cssClass = css
                        AppForm(EnForm.ObjActiveLabel).cnt = cnt
                        AppForm(EnForm.ObjActiveLabel).n = nObjSeen
                        AppForm(EnForm.ObjActiveLabel).Status = buildObjSummary(fn, "Active Label", cnt, nObjSeen, statgood, statbad, css)
                    Case "LabelNonactive"
                        objNbr = -1
                        AppForm(EnForm.ObjNonactiveLabel).cssClass = css
                        AppForm(EnForm.ObjNonactiveLabel).cnt = cnt
                        AppForm(EnForm.ObjNonactiveLabel).n = nObjSeen
                        AppForm(EnForm.ObjNonactiveLabel).Status = buildObjSummary(fn, "Nonactive Label", cnt, nObjSeen, statgood, statbad, css)
                End Select

                If objNbr > -1 Then   ' if not a Label
                    AppForm(objNbr).Status = buildObjSummary(fn, strobj, cnt, nObjSeen, statgood, statbad, css)
                    If AppForm(objNbr).cssClass <> "itemred" Then AppForm(objNbr).cssClass = "itemgreen" ' should likely check the req status 
                    AppForm(objNbr).cnt = cnt
                    AppForm(objNbr).n = nObjSeen

                    ' Check to see if the proper prefix was used
                    If Not s2.StartsWith(strObjPrefix(j)) And strObjPrefix(j) <> "-" Then
                        AppForm(objNbr).cssClass = "itemred"
                    End If

                End If

                cnt = 0
                foundflag = False
                statgood = ""
                statbad = ""
                nObjSeen = 0
            Next j

        End If
    End Sub

    Function buildObjSummary(fn As String, strobj As String, cnt As Integer, nobjseen As Integer, statgood As String, statbad As String, ByRef css As String) As String

        ' if the count = -1, color is set to clear, count = 0, the color is set to green. IF count > 0 the color is red
        Dim txt As String

        If nobjseen = 0 Then    ' object not seen
            txt = statgood
            css = ""       ' "itemclear"

        ElseIf cnt = 0 Then    ' no problems seen
            txt = String.Format("<p> <span class=""boldtext"">" & fn & "</span> - ({0}) objects with proper prefix", nobjseen) & "<br>" & vbCrLf & statgood & "</p>"
            css = "itemgreen"
        Else
            txt = String.Format("<p> <span class=""boldtext"">" & fn & "</span> - ({0}) without proper prefix", cnt) & "<br>" & vbCrLf & "</p>"
            txt &= statbad
            css = "itemred"

            If nobjseen - cnt > 0 Then
                txt &= String.Format("<p> <span class=""boldtext"">" & fn & "</span> - ({0}) objects with proper prefix", nobjseen - cnt) & "<br>" & vbCrLf & statgood & "</p>"
            End If
        End If
        Return txt
    End Function
    ' ============================================ End of Checks ===================================================

    Sub BuildReport(key As String, ByRef SAssignment As AssignmentInfo, ByRef AppForm() As MyItems, ByRef SSummary() As MyItems, ByRef strReport As String)
        Dim dummyItem As New Assignment.MyItems
        Dim varlist() As String

        Dim errcnt As Integer
        Dim errComment As String = ""

        Dim starttable As String = "<table class=""info"">" & vbCrLf
        Dim th As String = "<tr> <th class=""req"" > Req </th> <th class=""req"" >OK </th>  <th class=""titlecol""> Item </th>   <th class=""statuscol"" > Status </th>  <th class=""ptcol"" > Possible<br> Pts. </th>   <th class=""ptcol"">  Your<br> Score</th>  <th class=""commentcol""> Comment </th> </tr>" & vbCrLf
        ' ----------------------------------------------------------------------------------
        '    Dim h2 As String = "<h2> {0} </h2>"
        Dim h3 As String = "<h3> {0} </h3>"
        ' ----------------------------------------------------------------------------------------------------
        With SSummary
            errcnt = 0

            Select Case key
                Case "Application Level"


                    strReport &= String.Format(h3, "Application Level Information") & vbCrLf & vbCrLf

                    strReport &= starttable
                    strReport &= th
                    strReport = AppendToReport(strReport, 3, "Development Environment", "", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                    strReport = AppendToReport(strReport, 2, " - SLN File", "Assignment", SAssignment.hasSLN, errcnt, errComment, "hasSLN", SAssignment.TotalScore)
                    strReport = AppendToReport(strReport, 2, " - vbProj File", "Assignment", SAssignment.hasVBproj, errcnt, errComment, "hasvbProj", SAssignment.TotalScore)
                    strReport = AppendToReport(strReport, 2, " - VB Version", "Assignment", SAssignment.VBVersion, errcnt, errComment, "hasVBVersion", SAssignment.TotalScore)
                    strReport = AppendToReport(strReport, 10, "General Help", "Assignment", SAssignment.VBVersion, errcnt, errComment, "", SAssignment.TotalScore)
                    errcnt = 0
                    errComment = ""

                    ReDim varlist(7)
                    varlist = {"hasSplashScreen", "hasAboutBox", "InfoAppTitle", "InfoDescription", "InfoCompany", "InfoProduct", "InfoTrademark", "InfoCopyright"}

                    If SomeVarDisplayed(varlist) Then

                        strReport = AppendToReport(strReport, 3, "Application Info", "", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)

                        strReport = AppendToReport(strReport, 2, " - Splash Screen", "Assignment", SAssignment.hasSplashScreen, errcnt, errComment, "hasSplashScreen", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - About Box", "Assignment", SAssignment.hasAboutBox, errcnt, errComment, "hasAboutBox", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Application Title", "Summary", SSummary(EnSummary.InfoAppTitle), errcnt, errComment, "InfoAppTitle", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Description", "Summary", SSummary(EnSummary.InfoDescription), errcnt, errComment, "InfoDescription", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Company", "Summary", SSummary(EnSummary.InfoCompany), errcnt, errComment, "InfoCompany", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Product", "Summary", SSummary(EnSummary.InfoProduct), errcnt, errComment, "InfoProduct", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Trademark", "Summary", SSummary(EnSummary.InfoTrademark), errcnt, errComment, "InfoTrademark", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Copyright", "Summary", SSummary(EnSummary.InfoCopyright), errcnt, errComment, "InfoCopyright", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 10, " Application Info", "", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        errcnt = 0
                        errComment = ""
                    End If

                    ReDim varlist(1)
                    varlist = {"OptionStrict", "OptionExplicit"}

                    If SomeVarDisplayed(varlist) Then

                        strReport = AppendToReport(strReport, 3, "Compile Options", "", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Option Strict", "Assignment", SAssignment.OptionStrict, errcnt, errComment, "OptionStrict", SAssignment.TotalScore)   ' ????????????????????????????/ jhg
                        strReport = AppendToReport(strReport, 2, " - Option Explicit", "Assignment", SAssignment.OptionExplicit, errcnt, errComment, "OptionExplicit", SAssignment.TotalScore) ' ????????????????????????????/ jhg
                        strReport = AppendToReport(strReport, 10, "Options", "", SSummary(EnSummary.InfoCopyright), errcnt, errComment, "", SAssignment.TotalScore)

                    End If
                    strReport &= "</table>" & vbCrLf & vbCrLf

                Case "Debugging"
                    ' This has not been implemented. I am not sure if it is possible, and if so how to determine if these features have been set within the student's environment.
                    strReport &= String.Format(h3, "Debugging")

                    strReport &= starttable
                    strReport &= th
                    strReport = AppendToReport(strReport, 3, "Debugging", "", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                    strReport = AppendToReport(strReport, 2, " - BreakPoints", "", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                    strReport = AppendToReport(strReport, 2, " - Watch Variables", "", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                    strReport = AppendToReport(strReport, 10, "Debugging", "", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                    errcnt = 0
                    errComment = ""

                    strReport &= "</table>" & vbCrLf & vbCrLf


                Case "New Project File"
                    '                strReport &= String.Format(h2, "Comments Related to <span class=""boldtext"">" & Var1 & "</span>", errcnt, errComment, "", SAssignment.TotalScore)

                Case "Form Objects"
                    With AppForm
                        strReport &= String.Format(h3, "Form Objects")

                        strReport &= starttable
                        strReport &= th

                        ReDim varlist(5)
                        varlist = {"ChangeFormText", "SetFormAcceptButton", "SetFormCancelButton", "ModifyStartPosition", "ChangeFormColor"}

                        If SomeVarDisplayed(varlist) Then

                            strReport = AppendToReport(strReport, 3, "Design Time Form Properties", "", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)

                            strReport = AppendToReport(strReport, 2, " - Form Text Property", "Form", AppForm(EnForm.FormText), errcnt, errComment, "ChangeFormText", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - Accept Button", "Form", AppForm(EnForm.FormAcceptButton), errcnt, errComment, "SetFormAcceptButton", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - Cancel Button", "Form", AppForm(EnForm.FormCancelButton), errcnt, errComment, "SetFormCancelButton", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - Start Position", "Form", AppForm(EnForm.FormStartPosition), errcnt, errComment, "ModifyStartPosition", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - Non-Gray Form Color", "Form", AppForm(EnForm.FormBackColor), errcnt, errComment, "ChangeFormColor", SAssignment.TotalScore)
                            '         strReport = AppendToReport(strReport, 2, " - Form Load Method", "Form", Appform(enForm.FormLoadMethod, errcnt, errComment, "UtilizeFormLoadMethod", SAssignment.TotalScore)

                            strReport = AppendToReport(strReport, 10, "", "", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                            errcnt = 0
                            errComment = ""
                        End If

                        ReDim varlist(12)
                        varlist = {"IncludeFrmInFormName", "ButtonObj", "ObjTextbox", "ActiveLabel", "NonactiveLabel", "ObjCombobox", "ObjListbox", "ObjRadioButton", "ObjCheckbox", "ObjGroupBox", "ObjOpenFileDialog", "ObjSaveFileDialog", "ObjWebBrowser"}

                        If SomeVarDisplayed(varlist) Then
                            strReport = AppendToReport(strReport, 3, "Form Object Names Incorporate Object Prefix", "", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)

                            strReport = AppendToReport(strReport, 2, " - Form (frm)", "Form", AppForm(EnForm.FormName), errcnt, errComment, "IncludeFrmInFormName", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - Buttons (btn)", "Form", AppForm(EnForm.ObjButton), errcnt, errComment, "ButtonObj", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - Textboxes (txt)", "Form", AppForm(EnForm.ObjTextbox), errcnt, errComment, "TextboxObj", SAssignment.TotalScore)
                            '  strReport = AppendToReport(strReport, 2, " - Labels (lbl)", "Form", Appform(enForm.objLabel), errcnt, errComment, "", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - Active Labels (lbl)", "Form", AppForm(EnForm.ObjActiveLabel), errcnt, errComment, "ActiveLabels", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - NonActive Labels (no prefix needed)", "Form", AppForm(EnForm.ObjNonactiveLabel), errcnt, errComment, "NonActiveLabels", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - Combobox (cbx)", "Form", AppForm(EnForm.ObjCombobox), errcnt, errComment, "ComboBoxObj", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - Listbox (lbx)", "Form", AppForm(EnForm.ObjListbox), errcnt, errComment, "ListBoxObj", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - Radiobutton (rbn)", "Form", AppForm(EnForm.ObjRadioButton), errcnt, errComment, "RadioButtonObj", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - Checkbox (cbx)", "Form", AppForm(EnForm.ObjCheckbox), errcnt, errComment, "CheckBoxObj", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - Groupbox (gbx)", "Form", AppForm(EnForm.ObjGroupBox), errcnt, errComment, "GroupBoxObj", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - OpenFileDialog (ofd)", "Form", AppForm(EnForm.ObjOpenFileDialog), errcnt, errComment, "ObjOpenFileDialog", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - SaveFileDialog (sfd)", "Form", AppForm(EnForm.ObjSaveFileDialog), errcnt, errComment, "ObjSaveFileDialog", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - WebBrowser (wb)", "Form", AppForm(EnForm.ObjWebBrowser), errcnt, errComment, "WebBrowserObj", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 10, "Form Objects", "", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                            errcnt = 0
                            errComment = ""
                        End If

                        strReport &= "</table>" & vbCrLf & vbCrLf
                    End With

                Case "Coding Standards"

                    strReport &= String.Format(h3, "Coding Standards")

                    strReport &= starttable
                    strReport &= th


                    ReDim varlist(5)
                    varlist = {"CommentSubs", "CommentIF", "CommentFOR", "CommentDO", "CommentWHILE", "CommentSELECT"}

                    If SomeVarDisplayed(varlist) Then
                        strReport = AppendToReport(strReport, 3, "Use of Comments", "", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - First Line of Sub/Function", "Summary", SSummary(EnSummary.CommentSub), errcnt, errComment, "CommentSubs", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Prior to IF", "Summary", SSummary(EnSummary.CommentIF), errcnt, errComment, "CommentIF", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Prior to For", "Summary", SSummary(EnSummary.CommentFor), errcnt, errComment, "CommentFOR", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Prior to Do", "Summary", SSummary(EnSummary.CommentDo), errcnt, errComment, "CommentDO", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Prior to While", "Summary", SSummary(EnSummary.CommentWhile), errcnt, errComment, "CommentWHILE", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Prior to Select Case", "Summary", SSummary(EnSummary.CommentSelect), errcnt, errComment, "CommentSELECT", SAssignment.TotalScore)

                        If errComment.Length > 0 Then
                            '  errComment = ""
                            dummyItem.cssClass = "itemred"
                            strReport = AppendToReport(strReport, 10, "General Comments on Comments", "", dummyItem, errcnt, errComment, "CommentGeneral", SAssignment.TotalScore)
                        Else
                            strReport = AppendToReport(strReport, 10, "Use of Comments", "", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        End If

                        errcnt = 0
                        errComment = ""
                    End If


                    ReDim varlist(2)
                    varlist = {"VarArrays", "VarLists", "VarStructures"}

                    If SomeVarDisplayed(varlist) Then

                        '     If .VarArrays.showVar OrElse .VarLists.showVar OrElse .VarStructures.showVar Then
                        strReport = AppendToReport(strReport, 3, "Data Structures", "", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Arrays", "Summary", SSummary(EnSummary.VarArrays), errcnt, errComment, "VarArrays", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Lists", "Summary", SSummary(EnSummary.VarLists), errcnt, errComment, "VarLists", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Structures", "Summary", SSummary(EnSummary.VarStructures), errcnt, errComment, "VarStructures", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 10, "Data Structures", "", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        errcnt = 0
                        errComment = ""
                    End If


                    ReDim varlist(4)
                    varlist = {"VarString", "VarInteger", "VarDecimal", "VarDate", "VarBoolean"}

                    If SomeVarDisplayed(varlist) Then

                        strReport = AppendToReport(strReport, 3, "Variable Data Types - Checking to see which data types are used ", "", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - String", "Summary", SSummary(EnSummary.VarString), errcnt, errComment, "VarString", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Integer", "Summary", SSummary(EnSummary.VarInteger), errcnt, errComment, "VarInteger", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Decimal/Double", "Summary", SSummary(EnSummary.VarDecimal), errcnt, errComment, "VarDecimal", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Date", "Summary", SSummary(EnSummary.VarDate), errcnt, errComment, "VarDate", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Boolean", "Summary", SSummary(EnSummary.VarBoolean), errcnt, errComment, "VarBoolean", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 10, "Variable Data Types", "", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        errcnt = 0
                        errComment = ""
                    End If


                    ReDim varlist(18)
                    varlist = {"LogicElse", "LogicElseIF", "LogicNestedIF", "LogicNestedFOR", "LogicConvertToString", "LogicStringFormat", "LogicStringFormatParameters", "LogicConcatination", "LogicCaseInsensitive", "LogicTryCatch", "LogicComplexConditions", "LogicStreamReader", "LogicStreamReaderClose", "LogicStreamWriter", "LogicStreamWriterClose"}

                    If SomeVarDisplayed(varlist) Then


                        strReport = AppendToReport(strReport, 3, "Program Logic - Checking to see if each Programming Control Stucture is used or not", "", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Else", "Summary", SSummary(EnSummary.LogicElse), errcnt, errComment, "LogicElse", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - ElseIF", "Summary", SSummary(EnSummary.LogicElseIF), errcnt, errComment, "LogicElseIF", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Nested IF", "Summary", SSummary(EnSummary.LogicNestedIF), errcnt, errComment, "LogicNestedIF", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Nested For/Do", "Summary", SSummary(EnSummary.LogicNestedFor), errcnt, errComment, "LogicNestedFOR", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Convert to String (cStr or .toString)", "Summary", SSummary(EnSummary.LogicConvertToString), errcnt, errComment, "LogicConvertToString", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Format String (.toString() or String.Format())", "Summary", SSummary(EnSummary.LogicStringFormat), errcnt, errComment, "LogicStringFormat", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Template Parameters", "Summary", SSummary(EnSummary.LogicStringFormatParameters), errcnt, errComment, "LogicStringFormatParameters", SAssignment.TotalScore)
                        '   strReport = AppendToReport(strReport, 2, " - ByRef Parameters", "Summary", SSummary(ensummary.LogicByRef), errcnt, errComment, "LogicByRef", SAssignment.TotalScore)
                        '   strReport = AppendToReport(strReport, 2, " - Optional Parameters", "Summary", SSummary(ensummary.LogicOptional), errcnt, errComment, "LogicOptional", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Concatenation", "Summary", SSummary(EnSummary.LogicConcatination), errcnt, errComment, "LogicConcatination", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Case Insensitve ", "Summary", SSummary(EnSummary.LogicCaseInsensitive), errcnt, errComment, "LogicCaseInsensitive", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Try ... Catch", "Summary", SSummary(EnSummary.LogicTryCatch), errcnt, errComment, "LogicTryCatch", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Complex Conditions", "Summary", SSummary(EnSummary.LogicComplexConditions), errcnt, errComment, "LogicComplexConditions", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Req Use of a StreamReader", "Summary", SSummary(EnSummary.LogicStreamReader), errcnt, errComment, "LogicStreamReader", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Req matching StreamReader.Close", "Summary", SSummary(EnSummary.LogicStreamReaderClose), errcnt, errComment, "LogicStreamReaderClose", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Req Use of a StreamWriter", "Summary", SSummary(EnSummary.LogicStreamWriter), errcnt, errComment, "LogicStreamWriter", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Req matching StreamWriter.Close", "Summary", SSummary(EnSummary.LogicStreamWriterClose), errcnt, errComment, "LogicStreamWriterClose", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 10, "Program Logic", "", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        errcnt = 0
                        errComment = ""


                    End If


                    ReDim varlist(2)
                    varlist = {"SystemIO", "SystemNet", "SystemDB"}

                    If SomeVarDisplayed(varlist) Then

                        '    If .SystemIO.showVar OrElse .SystemNet.showVar OrElse .SystemDB.showVar Then
                        strReport = AppendToReport(strReport, 3, "Imports", "", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - System.IO", "Summary", SSummary(EnSummary.SystemIO), errcnt, errComment, "SystemIO", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - System.Net", "Summary", SSummary(EnSummary.SystemNet), errcnt, errComment, "SystemNet", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - System.DB", "Summary", SSummary(EnSummary.SystemDB), errcnt, errComment, "SystemDB", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 10, "Imports", "", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        errcnt = 0
                        errComment = ""

                    End If


                    ReDim varlist(5)
                    varlist = {"LogicSub", "LogicOptional", "LogicByRef", "LogicMultipleForms", "LogicModule", "LogicFormLoad"}

                    If SomeVarDisplayed(varlist) Then

                        '    If .SystemIO.showVar OrElse .SystemNet.showVar OrElse .SystemDB.showVar Then
                        strReport = AppendToReport(strReport, 3, "Subs / Functions", "", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Subs", "Summary", SSummary(EnSummary.LogicSub), errcnt, errComment, "LogicSub", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Optional Parameters", "Summary", SSummary(EnSummary.LogicOptional), errcnt, errComment, "LogicOptional", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - ByRef Parameters", "Summary", SSummary(EnSummary.LogicByRef), errcnt, errComment, "LogicByRef", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Multiple Forms", "Summary", SSummary(EnSummary.LogicMultipleForms), errcnt, errComment, "LogicMultipleForms", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Include Module", "Summary", SSummary(EnSummary.LogicModule), errcnt, errComment, "LogicModule", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Form Load Method", "Summary", SSummary(EnSummary.LogicFormLoad), errcnt, errComment, "LogicFormLoad", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 10, "Subs / Functions", "", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        errcnt = 0
                        errComment = ""

                    End If


                    strReport &= "</table>" & vbCrLf & vbCrLf


                    strReport = strReport.Replace("[TLOC]", TotalLinesOfCode.ToString("n0") & " Lines of Code (not including comments)")

                    '  TotalScore = (SAssignment.TotalScore / TotalPossiblePts).ToString("p1")


                    '    SAssignment.strTotalScore = SAssignment.TotalScore.ToString("n1") & " deduction out of " & TotalPossiblePts.ToString("n1") & " possible points = " & (SAssignment.TotalScore / TotalPossiblePts).ToString("p1")

                    '    strReport = strReport.Replace("[SCORE]", SAssignment.strTotalScore & vbCrLf)


                Case Else
                    If key.StartsWith("Assessment Results for") Then
                        strReport &= "<h2>" & key & "</h2>" & vbCrLf
                    Else
                        MessageBox.Show("BuildReport received an unknown key = " & key)
                    End If
            End Select
        End With
    End Sub


    Function BuildSummaryDetail(Assign As AssignmentInfo, AppForm() As MyItems, AppSummary() As MyItems) As String

        Dim src As String

        AssScore = 0
        AssPossible = 0

        src = InitializeStudentReport()

        BuildReport("Application Level", Assign, AppForm, AppSummary, src)
        BuildReport("New Project File", Assign, AppForm, AppSummary, src) ' the AppForm(0) is just a placeholder
        BuildReport("Assessment Results for " & ReturnLastField("", "\"), Assign, AppForm, AppSummary, src)
        BuildReport("Form Objects", Assign, AppForm, AppSummary, src)
        BuildReport("Coding Standards", Assign, AppForm, AppSummary, src)

        Return src
    End Function



    Public Sub FindIntegratedScore(SAssignment As AssignmentInfo, AppForm() As MyItems, SSummary() As MyItems)

        Dim T As Decimal = 0
        Dim i As Integer
        Dim setting As MySettings

        Dim sw As StreamWriter

        sw = File.CreateText(Application.StartupPath & "\integratedpts.txt")

        For i = 0 To SSummary.GetUpperBound(0)
            setting = Find_Setting(EnSummaryName(i), "FindIntegratedScore")
            With SSummary(i)

                If setting.Req Then
                    If .cnt = 0 Then
                        .YourPts = Math.Min(setting.PtsPerError * Math.Max(0, .n), setting.MaxPts)
                        If .n = 0 Then
                            .YourPts = 0
                            '  .cssClass = "itemred"
                        Else
                            '  .cssClass = "itemgreen"
                        End If

                    Else
                        .YourPts = setting.MaxPts - Math.Min(setting.PtsPerError * Math.Max(0, .cnt), setting.MaxPts)
                        '  .cssClass = "itemred"
                    End If
                    sw.WriteLine(i.ToString & vbTab & setting.Name & vbTab & setting.PtsPerError & vbTab & setting.MaxPts & vbTab & .cnt & vbTab & .n & vbTab & .YourPts)
                    T += .YourPts
                Else
                    .YourPts = 0
                    '  .cssClass = "itemclear"

                End If
            End With
        Next i


        For i = 0 To AppForm.GetUpperBound(0)
            setting = Find_Setting(EnFormNames(i), "FindIntegratedScore")
            With AppForm(i)


                If setting.Req And setting.MaxPts <> 0 Then
                    If .cnt = 0 Then
                        .YourPts = Math.Min(setting.PtsPerError * Math.Max(0, .n), setting.MaxPts)
                        If .n = 0 Then
                            .YourPts = 0
                            '          .cssClass = "itemred"
                        Else
                            '         .cssClass = "itemgreen"
                        End If

                    Else
                        .YourPts = setting.MaxPts - Math.Min(setting.PtsPerError * Math.Max(0, .cnt), setting.MaxPts)
                        '        .cssClass = "itemred"
                    End If

                    sw.WriteLine(i.ToString & vbTab & setting.Name & vbTab & setting.PtsPerError & vbTab & setting.MaxPts & vbTab & .cnt & vbTab & .n & vbTab & .YourPts)

                    T += .YourPts
                Else
                    .YourPts = 0
                    '      .cssClass = "itemclear"
                End If
            End With
        Next i
        sw.Close()
    End Sub


    Function SomeVarDisplayed(a() As String) As Boolean
        ' This checks an array of strings to see if any of them have settings which are either required or show var. If so, it returns true, else it returns false.
        '      Dim Flag As Boolean
        Dim Setting As New MySettings

        '    Flag = False

        For Each s As String In a
            Setting = Find_Setting(s, "SomeVarDisplayed")
            If Setting.Name <> Nothing Then
                If Setting.Req Or Setting.ShowVar Then Return True
            Else
                Beep()
            End If
        Next s
        Return False
    End Function



    Function AppendToReport(rpt As String, Template As Integer, Title As String, ItemType As String, topic As MyItems, ByRef errcnt As Integer, ByRef err As String, Item As String, ByRef total As Decimal) As String
        Dim s As String = ""
        Dim isok As String = ""
        Dim req As String
        '     Dim errstring As String = ""
        Dim EC As New ErrComments
        Dim feedback As String
        Dim nonChk As String = ""
        Dim showvar As Boolean


        Dim Setting As New MySettings

        'If Title = " - Structures" Then
        '    Beep()
        'End If

        ' check to see if this is a header or footer line. If so it bypasses the lookup of the item
        If Item.Length > 0 Then
            Setting = Find_Setting(Item, "AppendToReport")
            If Setting.Name Is Nothing Then
                ' debug
                Dim sw As StreamWriter
                sw = File.AppendText(Application.StartupPath & "\MissingSetting.txt")
                sw.WriteLine(Title & vbTab & Item)
                sw.Close()
                feedback = ""
            Else    ' This is a viable item, so process it. 
                If Setting.Req Then
                    feedback = Setting.Feedback
                    showvar = True  ' always show required variables
                    nonChk = "ncWhite"
                Else
                    If HideGray = "Gray" Then
                        showvar = Setting.ShowVar      ' - Don't toggle on - Display only the items the user specifies.
                        feedback = Setting.Feedback
                        nonChk = "ncGray"
                    ElseIf HideGray = "ShowAll" Then
                        showvar = True  ' toggle on - This displays all checks    
                        feedback = Setting.Feedback
                        nonChk = "ncGray"
                    ElseIf HideGray = "OnlyReq" Then
                        showvar = Setting.Req  ' just display required elements   
                        feedback = Setting.Feedback
                        nonChk = "ncWhite"

                    Else
                        showvar = False
                        feedback = ""
                        topic.cssClass = "itemwhite"
                    End If
                End If
            End If
        Else
            showvar = True
            feedback = ""
            nonChk = "ncWhite"
        End If


        Dim tr1 As String = "<tr class=""{0}"">" & vbCrLf & "<td class=""req"">{1}</td><td class=""req"">{2}</td><th class=""tablebody""> {3} </td>" & vbCrLf & "  <td class=""{4}"">{5}</td>" & vbCrLf & "   <td> {6} </td>" & vbCrLf & "   <td> {7} </td>" & vbCrLf & "   <td> {8} </td>" & vbCrLf & "</tr>" & vbCrLf

        Dim tr2 As String = "<tr class=""{0}"">" & vbCrLf & "<td class=""req"">{1}</td>" & vbCrLf & "<td class=""req"">{2}</td>" & vbCrLf & "   <td class=""{3}""> {4} </td>" & vbCrLf & "   <td class=""{5}""> {6} </td>" & vbCrLf & "   <td class=""rjust""> {7} </td>" & vbCrLf & "   <td class=""rjust""> {8} </td>" & vbCrLf & " <td> {9} </td>" & vbCrLf & "</tr>" & vbCrLf

        Dim tr3 As String = "<tr>" & vbCrLf & "   <td colspan=""8"" class=""divider""> {0} </td>" & vbCrLf & "</tr>" & vbCrLf

        Dim cmt As String = "<tr>" & vbCrLf & "   <td colspan=""8"" class=""errcomment""> {0} </td> </tr>" & vbCrLf & "<tr>" & vbCrLf & "   <td colspan=""8"" class=""blankline""> </td>" & "</tr>" & vbCrLf      '       & "<tr><td  colspan=""8""></td></tr>" & vbCrLf



        ' ----------------------------------------------------------------------------------
        '       If (frmMain.rbnShowOnlyReq.Checked And topic.req) Or (Not frmMain.rbnShowOnlyReq.Checked And (Setting.ShowVar Or topic.req)) Then
        If Setting.Req Or (HideGray = "OnlyReq" And (Setting.Req Or Item.Length = 0)) Or (HideGray <> "OnlyReq" And (showvar Or Setting.Req)) Then
            With topic
                ' Calculate the grade for Summary Items
                If .cnt = -1 Then .cnt = 0
                req = ""
                isok = ""
                If Setting.Req Then
                    If Setting.Req Then req = "*"
                    If .cnt = 0 Then
                        .YourPts = Math.Min(Setting.PtsPerError * Math.Max(0, .n), Setting.MaxPts)
                        If .YourPts < Setting.MaxPts Then
                            .cssClass = "itemred"
                        Else
                            .cssClass = "itemgreen"
                        End If
                    Else
                        .YourPts = Setting.MaxPts - Math.Min(Setting.PtsPerError * Math.Max(0, .cnt), Setting.MaxPts)
                        .cssClass = "itemred"
                    End If
                    total += .YourPts

                    If .cssClass = "itemred" Then
                        isok = "&#x2717;"              ' load a X symbol
                    ElseIf .cssClass = "itemgreen" Then
                        isok = "&#x2713;"               ' Load a check symbol
                    End If
                Else
                    .YourPts = 0
                    .cssClass = "itemclear"
                End If


                ' If there is a problem, then create the error statement.
                If .cssClass = "itemred" Then
                    EC = ErrorComments.Find(Function(p) p.topic = Title)
                    If Not IsNothing(EC) Then
                        If EC.topic = Title Then
                            errcnt += 1
                            .Comments &= " [" & errcnt & "]"
                            If err.Length > 0 Then
                                err &= "<p class=""hangingindent"">[" & errcnt & "] - " & EC.Comment & " </p>" & vbCrLf
                            Else
                                err &= "<span class=""commentsummary"">Feedback</span>" & vbCrLf & "<p class=""hangingindent"">[" & errcnt & "] - " & EC.Comment & " </p>" & vbCrLf
                            End If

                        End If
                    End If


                    If feedback.Length > 0 Then
                        '          If EC.topic = Title Then
                        errcnt += 1
                        .Comments &= " [" & errcnt & "]"
                        If err.Length > 0 Then
                            err &= "<p class=""hangingindent"">[" & errcnt & "] <b>" & Title & "</b> - " & feedback & " </p>" & vbCrLf
                        Else
                            err &= "<span class=""commentsummary"">Feedback</span>" & vbCrLf & "<p class=""hangingindent"">[" & errcnt & "] <b>" & Title & "</b> - " & feedback & " </p>" & vbCrLf
                        End If
                    Else
                        ' Beep()
                    End If
                End If



                ' ---------------------------------------------------------------------------
                ' generate the HTML segment
                Select Case Template
                    Case 1
                        s = String.Format(tr1, "", req, isok, nonChk, Title, .cssClass, .Status, Setting.MaxPts, .YourPts)
                    Case 2
                        s = String.Format(tr2, "", req, isok, nonChk, Title, .cssClass, .Status, Setting.MaxPts, .YourPts, .Comments)
                    Case 3
                        s = String.Format(tr3, Title)

                    Case 10
                        If err.Length > 0 Then

                            s = String.Format(cmt, err)
                        Else
                            s = ""
                        End If

                        'errcnt = 0
                        'err = ""
                End Select


                ' Insert the HTML segment into the file
                If Setting.Req Or (HideGray = "OnlyReq" And (Setting.Req Or Item.Length = 0)) Or (HideGray <> "OnlyReq") Then
                    rpt &= s
                    AssScore += .YourPts
                    AssPossible += Setting.MaxPts
                End If
            End With
        End If

        Return rpt
    End Function


    Sub PopulateNonCheckCSS_Form2(ByRef AppForm() As MyItems)
        Dim nc As String = ""                       ' Non-checked property
        Dim c As String = "ncWhite"    ' Checked property

        ' ----------------------------------------------------------

        If HideGray = "Gray" Then
            nc = "ncGray"
        ElseIf HideGray = "Hide" Then
            nc = "ncHide"
        Else
            nc = "ncWhite"
        End If


        ' -----------------------------------------------------------

        If Find_Setting("RenameObjects", "PopulatenonCheckCSS").Req Then
            AppForm(EnForm.ObjButton).cssNonChk = c
            AppForm(EnForm.objLabel).cssNonChk = c
            AppForm(EnForm.ObjActiveLabel).cssNonChk = c
            AppForm(EnForm.ObjNonactiveLabel).cssNonChk = c
            AppForm(EnForm.ObjTextbox).cssNonChk = c
            AppForm(EnForm.ObjListbox).cssNonChk = c
            AppForm(EnForm.ObjCombobox).cssNonChk = c
            AppForm(EnForm.ObjRadioButton).cssNonChk = c
            AppForm(EnForm.ObjCheckbox).cssNonChk = c
            AppForm(EnForm.ObjGroupBox).cssNonChk = c
            AppForm(EnForm.ObjPanel).cssNonChk = c
            AppForm(EnForm.ObjWebBrowser).cssNonChk = c
            AppForm(EnForm.ObjOpenFileDialog).cssNonChk = c
            AppForm(EnForm.ObjSaveFileDialog).cssNonChk = c

        Else
            AppForm(EnForm.ObjButton).cssNonChk = nc
            AppForm(EnForm.objLabel).cssNonChk = nc
            AppForm(EnForm.ObjActiveLabel).cssNonChk = nc
            AppForm(EnForm.ObjNonactiveLabel).cssNonChk = nc
            AppForm(EnForm.ObjTextbox).cssNonChk = nc
            AppForm(EnForm.ObjListbox).cssNonChk = nc
            AppForm(EnForm.ObjCombobox).cssNonChk = nc
            AppForm(EnForm.ObjRadioButton).cssNonChk = nc
            AppForm(EnForm.ObjCheckbox).cssNonChk = nc
            AppForm(EnForm.ObjGroupBox).cssNonChk = nc
            AppForm(EnForm.ObjPanel).cssNonChk = nc
            AppForm(EnForm.ObjWebBrowser).cssNonChk = nc
            AppForm(EnForm.ObjOpenFileDialog).cssNonChk = nc
            AppForm(EnForm.ObjSaveFileDialog).cssNonChk = nc

        End If

        If Find_Setting("ObjOpenFileDialog", "PopulatenonCheckCSS").Req Then AppForm(EnForm.ObjOpenFileDialog).cssNonChk = c Else AppForm(EnForm.ObjOpenFileDialog).cssNonChk = nc
        If Find_Setting("ObjSaveFileDialog", "PopulatenonCheckCSS").Req Then AppForm(EnForm.ObjSaveFileDialog).cssNonChk = c Else AppForm(EnForm.ObjSaveFileDialog).cssNonChk = nc

        If Find_Setting("ChangeFormText", "PopulatenonCheckCSS").Req Then AppForm(EnForm.FormText).cssNonChk = c Else AppForm(EnForm.FormText).cssNonChk = nc

        If Find_Setting("ChangeFormColor", "PopulatenonCheckCSS").Req Then AppForm(EnForm.FormBackColor).cssNonChk = c Else AppForm(EnForm.FormBackColor).cssNonChk = nc
        If Find_Setting("SetFormAcceptButton", "PopulatenonCheckCSS").Req Then AppForm(EnForm.FormAcceptButton).cssNonChk = c Else AppForm(EnForm.FormAcceptButton).cssNonChk = nc
        If Find_Setting("SetFormCancelButton", "PopulatenonCheckCSS").Req Then AppForm(EnForm.FormCancelButton).cssNonChk = c Else AppForm(EnForm.FormCancelButton).cssNonChk = nc
        If Find_Setting("ModifyStartPosition", "PopulatenonCheckCSS").Req Then AppForm(EnForm.FormStartPosition).cssNonChk = c Else AppForm(EnForm.FormStartPosition).cssNonChk = nc
        If Find_Setting("LogicFormLoad", "PopulatenonCheckCSS").Req Then AppForm(EnForm.FormLoadMethod).cssNonChk = c Else AppForm(EnForm.FormLoadMethod).cssNonChk = nc

        ' ----------------------------------------------------------


    End Sub

End Module
