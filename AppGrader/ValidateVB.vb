Imports System.IO
'Imports SharpCompress.Archive
'Imports SharpCompress.Common


Module ValidateVB

    '   Public ValidateReport As String
    Public strStudentReport As String
    Public strFacReport As String



    Sub InitializeStudentReport()
        Dim s As String

        Try
            Dim sr1 As New StreamReader(Application.StartupPath & "\templates\rptStudentHeader.html")

            s = sr1.ReadToEnd
            sr1.Close()

            s = s.Replace("[title]", strStudentID & " - " & cfgAssignmentTitle & " Summary")
            ' strStudentReport = strStudentReport.Replace("[STUDENT]", strStudentID & " - " & cfgAssignmentTitle & " Summary")
            s = s.Replace("[ASSIGNMENTNAME]", frmMain.txtAssignmentName.Text)
            s = s.Replace("[STUDENT]", strStudentID)
            s = s.Replace("[VERSION]", Application.ProductVersion)

            strStudentReport = s & strStudentReport
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

        strStudentReport = strStudentReport.Replace("[APPTIME]", SubmissionCompileTime)
        strStudentReport = strStudentReport.Replace("[APPDATE]", SubmissionCompileDate)

    End Sub

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

            s &= ("<tr><th class=""header"";>Student ID</th><th class=""header"";>Filename</th><th class=""header"";>TLOC</th><th class=""header"";>Score</th></tr>" & vbCrLf)

        Catch ex As Exception
            MessageBox.Show("InitializeFacultyReport - " & ex.Message)
        End Try

        Return s
    End Function


    Public Sub CloseFacReport(path As String, fn As String)
        Dim sr As StreamReader
        Dim sw As StreamWriter

        sr = File.OpenText(Application.StartupPath & "\templates\rptFacFooter.html")
        strFacReport &= sr.ReadToEnd
        sr.Close()

        sw = File.CreateText((strOutputPath & fn))
        sw.Write(strFacReport)
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
            '           strFacReport &= "<li>" & filename & "</li>"
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



    Function CheckForComments2(filesource As String, ByRef SSummary() As MyItems, ByRef AppForm() As MyItems, sender As System.ComponentModel.BackgroundWorker) As Integer

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
                            .n += 1   ' counts up bad comments for FOR statements
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
                            .n += 1   ' counts up bad comments for IF statements
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
                        SSummary(EnSummary.SystemIO).Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf
                    Else
                        SSummary(EnSummary.SystemIO).Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - Not found</p>" & vbCrLf
                    End If

                    If ss(i).Contains("System).net") Then
                        SSummary(EnSummary.SystemNet).Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf
                    Else
                        SSummary(EnSummary.SystemNet).Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - Not found</p>" & vbCrLf
                    End If

                    If ss(i).Contains("System.DB") Then
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

                Case "MESSAGEBOX.SHOW", "MSGBOX"
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
                            If (ss(i - 1).StartsWith("'") Or ss(i + 1).StartsWith("'")) Then
                                .bad &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & TrimAfter(ss(i), "(", True) & "</p>" & vbCrLf
                                .cssClass = "itemred"
                                .n += 1
                            Else
                                .good &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & TrimAfter(ss(i), "(", True) & "</p>" & vbCrLf
                            End If
                        End With

                    End If

                    ' ------------------------------------------
                    ' check for optional parameters
                    With SSummary(EnSummary.LogicOptional)
                        If ss(i).Contains(" Optional ") Then
                            .Status = TrimAfter(ss(i), "(", True) & " defined in line (" & lineno(i).ToString & ") has an Optional Parameter" & " </p>" & vbCrLf
                            .n += 1
                        End If
                    End With

                    ' -------------------------------------
                    ' check for byref parameters
                    With SSummary(EnSummary.LogicByRef)
                        If ss(i).Contains(" ByRef ") Then
                            .Status = TrimAfter(ss(i), "(", True) & " defined in line (" & lineno(i).ToString & ") has a ByRef Parameter" & " <br />" & vbCrLf
                            .n += 1
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
                            .n += 1
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
                SSummary(EnSummary.LogicFormLoad).Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf
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
                        .cnt += 1 ' This counts how many times they used the formating feature
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
        ProcessComment(SSummary(EnSummary.LogicSub), SSummary(EnSummary.CommentSub), "Subs / Functions", "CommentSubs")
        ProcessComment(SSummary(EnSummary.LogicIF), SSummary(EnSummary.CommentIF), "IF statements", "CommentIF")
        ProcessComment(SSummary(EnSummary.LogicFor), SSummary(EnSummary.CommentFor), "FOR statements", "CommentFOR")
        ProcessComment(SSummary(EnSummary.LogicDo), SSummary(EnSummary.CommentDo), "DO statements", "CommentDO")
        ProcessComment(SSummary(EnSummary.LogicWhile), SSummary(EnSummary.CommentWhile), "WHILE statements", "CommentWHILE")
        ProcessComment(SSummary(EnSummary.LogicSelectCase), SSummary(EnSummary.CommentSelect), "SELECT CASE statements", "CommentSELECT")
        ' ------------------------------------------------------------------------------------------------------------

        If SSummary(EnSummary.LogicCStr).n > 0 Or SSummary(EnSummary.LogicToString).n > 0 Then
            SSummary(EnSummary.LogicToString).Status = "<span class=""boldtext"">Converting to Strings </span><br />" & vbCrLf

            If SSummary(EnSummary.LogicCStr).n > 0 Then
                SSummary(EnSummary.LogicToString).Status &= "<span class=""boldtext""><br />Using CStr() to convert to string </span><br /><br />" & vbCrLf
            End If

            If SSummary(EnSummary.LogicToString).n > 0 Then
                SSummary(EnSummary.LogicToString).Status &= "<span class=""boldtext""><br />Using .ToString to convert to string </span><br /><br />" & vbCrLf
            End If
        End If
        ' ------------------------------------------------------------------------------------------------------------

        If SSummary(EnSummary.LogicToStringFormat).n > 0 Or SSummary(EnSummary.LogicStringFormat).n > 0 Then
            If SSummary(EnSummary.LogicToStringFormat).n > 0 Then
                SSummary(EnSummary.LogicToStringFormat).Status &= "<span class=""boldtext""><br />Included .ToString(format) command </span><br /><br />" & vbCrLf
            End If
            If SSummary(EnSummary.LogicStringFormat).n > 0 Then
                SSummary(EnSummary.LogicStringFormat).Status &= "<span class=""boldtext""><br />Included String.Format(Template)  </span><br /><br />" & vbCrLf
            End If
        End If
        ' ------------------------------------------------------------------------------------------------------------
        ' -------------------------------------------------------------------------------------------------------
        ' AppInfo
        ProcessReq(SSummary(EnSummary.InfoAppTitle), "Application Title not modified", "Application Title modified", "InfoAppTitle")
        ProcessReq(SSummary(EnSummary.InfoDescription), "Application Description not modified", "Application Description modified", "InfoDescription")
        ProcessReq(SSummary(EnSummary.InfoCompany), "Application Company Info not modified", "Application Company Info modified", "InfoCompany")
        ProcessReq(SSummary(EnSummary.InfoProduct), "Application Product Info not modified", "Application Product Info modified", "InfoProduct")
        ProcessReq(SSummary(EnSummary.InfoTrademark), "Application Trademark not modified", "Application Trademark modified", "InfoTrademark")
        ProcessReq(SSummary(EnSummary.InfoCopyright), "Application Copyright not modified", "Application Copyright modified", "InfoCopyright")

        'Compile Options
        'ProcessReq(.OptionStrict, "No System.IO found", "Imports System.IO found (Line No.)", "Include System.IO")
        'ProcessReq(.SystemNet, "No System).net found", "Imports System).net found (Line No.)", "Include System).net")

        ' Form Design


        ' Imports
        ProcessReq(SSummary(EnSummary.SystemIO), "No System.IO found", "Imports System.IO found (Line No.)", "SystemIO")
        ProcessReq(SSummary(EnSummary.SystemNet), "No System).net found", "Imports System).net found (Line No.)", "SystemNet")
        ProcessReq(SSummary(EnSummary.SystemDB), "No System.DB found", "Imports System.DB found (Line No.)", "SystemDB")

        ' Vars
        ProcessReq(SSummary(EnSummary.VarArrays), "No Arrays declared", "Data Arrays declared (Line No.)", "VarArrays")
        ProcessReq(SSummary(EnSummary.VarLists), "No Lists declared", "Lists(of T) declared (Line No.)", "VarLists")
        ProcessReq(SSummary(EnSummary.VarStructures), "No Structures declared", "Data Structures defined (Line No.)", "VarStructures")
        ProcessReq(SSummary(EnSummary.VarString), "No String variables declared", "String variables declared (Line No.)", "VarString")
        ProcessReq(SSummary(EnSummary.VarInteger), "No Integer variables declared", "Integer variables declared (Line No.)", "VarInteger")
        ProcessReq(SSummary(EnSummary.VarDecimal), "No Decimal / Double variables declared", "Decimal / Double variables declared (Line No.)", "VarDecimal")
        ProcessReq(SSummary(EnSummary.VarDate), "No Date variables declared", "Date variables declared (Line No.)", "VarDate")
        ProcessReq(SSummary(EnSummary.VarBoolean), "No Boolean variables declared", "Boolean variables declared (Line No.)", "VarBoolean")

        ' Logic
        ProcessReq(SSummary(EnSummary.LogicWhile), "No While statements found", "While statements found (Line No.)", "LogicWHILE")
        ProcessReq(SSummary(EnSummary.LogicSelectCase), "No SelectCase statements found", "Select Case statements found (Line No.)", "LogicSelectCase")

        ProcessReq(SSummary(EnSummary.LogicConvertToString), "No Conversion to String (.CStr or .toString) found", "Conversion to String (.CStr or .toString) found (Line No.)", "LogicConvertToString")
        ProcessReq(SSummary(EnSummary.LogicStringFormat), "No String Formatting (.toString() or String.Format()) found", "String Formatting (.toString() or String.Format()) found (Line No.)", "LogicStringFormat")
        ProcessReq(SSummary(EnSummary.LogicStringFormatParameters), "No Parameterized string formats (string.format(f,{0}) found", "Parameterized String Formatting (string.format(f,{0}) found (Line No.)", "LogicStringFormatParameters")

        ProcessReq(SSummary(EnSummary.LogicConcatination), "No String Concatination found", "String Concatination found (Line No.)", "LogicConcatination")
        ProcessReq(SSummary(EnSummary.LogicCaseInsensitive), "No Case-Insensitive comparisons found", "Case-Insensitive comparisons found (Line No.)", "LogicCaseInsensitive")
        ProcessReq(SSummary(EnSummary.LogicComplexConditions), "No Complex Conditions (AND / OR / ANDALSO / ORELSE) found", "Complex Conditions (AND / OR / ANDALSO / ORELSE) found (Line No.)", "LogicComplexConditions")

        ProcessReq(SSummary(EnSummary.LogicElse), "No Else statements found", "Else statements found (Line No.)", "LogicElse")
        ProcessReq(SSummary(EnSummary.LogicElseIF), "No ElseIF statements found", "ElseIF statements found (Line No.)", "LogicElseIF")
        ProcessReq(SSummary(EnSummary.LogicNestedIF), "No Nested IF statements found", "Nested IF statements found (Line No.)", "LogicNestedIF")
        ProcessReq(SSummary(EnSummary.LogicNestedFor), "No Nested FOR statements found", "Nested FOR statements found (Line No.)", "LogicNestedFOR")
        ProcessReq(SSummary(EnSummary.LogicOptional), "No Optional Sub / Fuction Parameters found", "Optional Sub / Fuction Parameters found (Line No.)", "LogicOptional")
        ProcessReq(SSummary(EnSummary.LogicByRef), "No Sub / Function ByRef Parameters found", "Sub / Function ByRef Parameters found (Line No.)", "LogicByRef")
        ProcessReq(SSummary(EnSummary.LogicTryCatch), "No Try ... Catch Statement Found", "Try ... Catch Statements found (Line No.)", "LogicTryCatch")
        ProcessReq(SSummary(EnSummary.LogicFormLoad), "No Form Load Method Found", "Form Load Method found (Line No.)", "LogicFormLoad")

        ProcessReq(SSummary(EnSummary.LogicStreamReader), "No StreamReaders found", "StreamReader found (Line No.)", "LogicStreamReader")
        ProcessReq(SSummary(EnSummary.LogicStreamReaderClose), "No StreamReader.Close found", "StreamReader.Close found (Line No.)", "LogicStreamReaderClose")
        ProcessReq(SSummary(EnSummary.LogicStreamWriter), "No StreamWriters found", "StreamWriters found (Line No.)", "LogicStreamWriter")
        ProcessReq(SSummary(EnSummary.LogicStreamWriterClose), "No StreamWriter.Close found", "StreamWriter.Close found (Line No.)", "LogicStreamWriterClose")

        Return ncomments
    End Function


    Sub ProcessComment(ByRef logictype As MyItems, ByRef commenttype As MyItems, construct As String, nm As String)
        If logictype.n = 0 Then
            logictype.Status = "No " & construct & " Found"
            commenttype.Status = "No " & construct & " Found" & vbCrLf
            logictype.cssClass = "itemred"
            commenttype.cssClass = "itemred"
        Else
            If commenttype.bad <> Nothing Then
                commenttype.Status = "<span class=""boldtext"">" & construct & " Missing Descriptive Comments (Line No.)</span><br />" & vbCrLf & commenttype.bad & " <br /> <br />" & vbCrLf
                logictype.cssClass = "itemred"
                commenttype.cssClass = "itemred"
            End If
            If commenttype.good <> Nothing Then
                If logictype.bad = Nothing Then
                    logictype.cssClass = "itemgreen"
                    commenttype.cssClass = "itemgreen"
                End If
                commenttype.Status &= "<span class=""boldtext"">" & construct & " Includes Descriptive Comments (Line No.)</span><br />" & vbCrLf & commenttype.good & " <br />" & vbCrLf
            End If
        End If

        ' nm = logictype.ToString.Substring(1)

        If Not Find_Setting(nm, "ProcessComment").Req Then
            logictype.cssClass = "itemclear"
            commenttype.cssClass = "itemclear"
        End If

    End Sub

    Sub ProcessReq(ByRef type As MyItems, NotFound As String, Found As String, nm As String)
        If type.n = 0 Then
            type.Status = "<span class=""boldtext"">" & NotFound & "</span><br />" & vbCrLf
            type.cssClass = "itemred"
        Else
            type.Status = "<span class=""boldtext"">" & Found & "</span><br /><br />" & vbCrLf & type.Status & " <br />" & vbCrLf
            type.cssClass = "itemgreen"
            type.n += 1
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

        With AppForm


            tmp = filename.Replace(".vb", ".designer.vb")

            ' Do not process unless it has a .designer.vb file. This eliminates processing Modules and classes
            If File.Exists(tmp) Then
                Dim sr As New StreamReader(tmp)
                s = sr.ReadToEnd
                sr.Close()

                If s.Contains("Me.BackColor") Then
                    If s.Contains("System.Drawing.SystemColors.") Then
                        AppForm(EnForm.FormBackColor).Status = "<span class=""boldtext"">Form color = " & TrimUpTo(ss, "System.Drawing.SystemColors. </span>")
                        AppForm(EnForm.FormBackColor).cssClass = "itemgreen"
                        AppForm(EnForm.FormBackColor).n += 1
                    ElseIf s.Contains("System.Drawing.Color.FromArgb(") Then

                        ss = returnBetween(s, "Me.BackColor =", vbCrLf)
                        cArray = ss.Split(delim, StringSplitOptions.None)
                        Try
                            AppForm(EnForm.FormBackColor).Status = "<span class=""boldtext"">Form color is nongray (#" & ReturnHexEquivalent(TrimAfter(cArray(1), ",", True)) & ReturnHexEquivalent(TrimAfter(cArray(2), ",", True)) & ReturnHexEquivalent(TrimAfter(cArray(3), ",", True)) & ") </span>"
                            AppForm(EnForm.FormBackColor).cssClass = "itemgreen"
                            AppForm(EnForm.FormBackColor).n += 1
                        Catch
                            AppForm(EnForm.FormBackColor).Status = "Form color = " & ss
                        End Try
                    ElseIf s.Contains("System.Drawing.Color.") Then
                        AppForm(EnForm.FormBackColor).Status = "<span class=""boldtext"">Form color = " & returnBetween(s, "Me.BackColor = System.Drawing.Color.", vbCrLf) & "</span>"
                        AppForm(EnForm.FormBackColor).cssClass = "itemgreen"
                        AppForm(EnForm.FormBackColor).n += 1
                    End If
                Else
                    AppForm(EnForm.FormBackColor).Status = "<span class=""boldtext"">Form Color was not changed at design time (still gray) </span>"
                    AppForm(EnForm.FormBackColor).cssClass = "itemred"
                End If


                If s.Contains("Me.Text") Then
                    AppForm(EnForm.FormText).Status = "<span class=""boldtext"">" & returnBetween(s, "Me.Text = """, """") & "</span>"
                    AppForm(EnForm.FormText).cssClass = "itemgreen"
                    AppForm(EnForm.FormText).n += 1
                Else
                    AppForm(EnForm.FormText).Status = "<span class=""boldtext"">The form text was not changed.</span>"
                    AppForm(EnForm.FormText).cssClass = "itemred"

                End If

                If s.Contains("Me.StartPosition") Then
                    AppForm(EnForm.FormStartPosition).Status = "<span class=""boldtext"">Form Startup Position set to: [" & returnBetween(s, "Me.StartPosition = System.Windows.Forms.FormStartPosition.", vbCrLf) & "] </span>"
                    AppForm(EnForm.FormStartPosition).cssClass = "itemgreen"
                    AppForm(EnForm.FormStartPosition).n += 1
                Else
                    AppForm(EnForm.FormStartPosition).Status = "<span class=""boldtext"">Form StartPosition not Modified. </span>"
                    AppForm(EnForm.FormStartPosition).cssClass = "itemred"
                End If


                ' accept and cancel button settings in *.resx file
                sr = New StreamReader(filename.Replace(".vb", ".resx"))
                s = sr.ReadToEnd
                sr.Close()
                If s.Contains("Me.AcceptButton") Then
                    AppForm(EnForm.FormAcceptButton).Status = "<span class=""boldtext"">" & returnBetween(s, "Me.AcceptButton = ", vbCrLf) & "</span>"
                    AppForm(EnForm.FormAcceptButton).cssClass = "itemgreen"
                    AppForm(EnForm.FormAcceptButton).n += 1
                Else
                    AppForm(EnForm.FormAcceptButton).Status = "<span class=""boldtext"">Accept Button Property not set at design time.</span>"
                    AppForm(EnForm.FormAcceptButton).cssClass = "itemred"
                End If

                If s.Contains("Me.CancelButton") Then
                    AppForm(EnForm.FormCancelButton).Status = "<span class=""boldtext"">" & returnBetween(s, "Me.AcceptButton = ", vbCrLf) & "</span>"
                    AppForm(EnForm.FormCancelButton).cssClass = "itemgreen"
                    AppForm(EnForm.FormCancelButton).n += 1
                Else
                    AppForm(EnForm.FormCancelButton).Status = "<span class=""boldtext"">Cancel Button Property not set at design time.</span>"
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
                    ' has a splash screen
                    .Status = vbTrue.ToString
                    If RemoveFromList Then filesinbuild.Remove(filename)
                    Exit For
                End If
            Next
        End With
    End Sub



    Sub CheckForModules2(filesinbuild As List(Of String), ByRef item As MyItems)
        Dim source As String = ""
        Dim nmodules As Integer = 0

        With item
            .Comments = ""

            For Each filename As String In filesinbuild
                If GetFileSource(filename, source) Then
                    If source.Contains("End Module") Then
                        nmodules = nmodules + 1
                        .Comments = "Filename " & ReturnLastField(filename, "\") & " was identified as a Module"
                    End If
                End If
            Next filename
            ' ------------------------------------------------
            .Status = nmodules.ToString
            .cnt = nmodules
            If .req Then
                If .cnt > 0 Then
                    .cssClass = "itemgreen"
                Else
                    .cssClass = "itemred"
                End If
            End If

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
                    .cnt = 0
                    .n -= 1
                Else
                    .Status = tmp & " does not start with the frm prefix"
                    .cssClass = "itemred"
                    .cnt = 0
                End If
            End With
            ' ----------------------------------------------------------------------------------------------------------------
            For j As Integer = 0 To strObjects.GetUpperBound(0) ' Look through the list of object types we are interested in
                strobj = strObjects(j)
                foundflag = False ' indicates if the object on the form is in the list of interesting objects

                If strobj = "Button" Then
                    Beep()
                End If


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

                            ' compare with current object, and process if the same.
                            If s1.Trim.ToUpper = strobj.ToUpper Then
                                foundflag = True
                                nObjSeen += 1

                                s2 = obj    ' not sure if this is correct. I added it so s2 had an initial value.
                                If s2.Contains(" = ") Then s2 = obj.Substring(0, obj.IndexOf(" = "))

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
                                If s1 = "Label" Then    ' special check to see if it is active
                                    If filesourceVB.IndexOf(s2) > -1 Then
                                        IsActiveLabel = True
                                    Else
                                        IsActiveLabel = False
                                    End If
                                End If
                                ' --------------------------------------------------------------------------------------------------

                                If IsActiveLabel And tmpObj = "LabelActive" Then
                                    If Not s2.StartsWith(strObjPrefix(j)) Then                  ' Name of the object
                                        ' this is a problem - object not renamed
                                        statbad &= arrow & s2 & " <br />" & vbCrLf
                                        cnt = cnt + 1
                                    Else
                                        statgood &= bullet & s2 & " <br />" & vbCrLf
                                    End If
                                ElseIf Not IsActiveLabel And tmpObj = "LabelNonActive" Then
                                    statgood &= arrow & s2 & " <br />" & vbCrLf
                                ElseIf s1 <> "Label" Then
                                    If Not s2.StartsWith(strObjPrefix(j)) And (tmpObj <> "LabelNonactive") Then                  ' Name of the object
                                        ' this is a problem - object not renamed
                                        statbad &= arrow & s2 & " <br />" & vbCrLf
                                        cnt = cnt + 1    ' counting number of bad instances

                                    ElseIf Not IsActiveLabel Then  ' Nonactive labels need not be renamed.
                                        statgood &= bullet & s2 & " <br />" & vbCrLf
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

                If Not foundflag Then statgood = "None Found"

                objNbr = -1     ' 
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
                        AppForm(EnForm.objLabel).Status = buildObjSummary(strobj, cnt, nObjSeen, statgood, statbad, css)
                        AppForm(EnForm.objLabel).cssClass = css
                        AppForm(EnForm.objLabel).cnt = cnt
                        AppForm(EnForm.objLabel).n = cnt
                    Case "LabelActive"
                        objNbr = -1
                        AppForm(EnForm.ObjActiveLabel).cssClass = css
                        AppForm(EnForm.ObjActiveLabel).cnt = cnt
                        AppForm(EnForm.ObjActiveLabel).n += 1
                        AppForm(EnForm.ObjActiveLabel).Status = buildObjSummary("Active Label", cnt, nObjSeen, statgood, statbad, css)
                    Case "LabelNonactive"
                        objNbr = -1
                        nObjSeen += 1
                        AppForm(EnForm.ObjNonactiveLabel).cssClass = css
                        AppForm(EnForm.ObjNonactiveLabel).cnt = cnt
                        AppForm(EnForm.ObjNonactiveLabel).n -= 1
                        AppForm(EnForm.ObjNonactiveLabel).Status = buildObjSummary("Nonactive Label", cnt, nObjSeen, statgood, statbad, css)
                End Select

                If objNbr > -1 Then   ' if not a Label
                    AppForm(objNbr).Status = buildObjSummary(strobj, cnt, nObjSeen, statgood, statbad, css)
                    If AppForm(objNbr).cssClass <> "itemred" Then AppForm(objNbr).cssClass = "itemgreen" ' should likely check the req status 
                    AppForm(objNbr).cnt = cnt

                    ' Check to see if the proper prefix was used
                    If Not s2.StartsWith(strObjPrefix(j)) And strObjPrefix(j) <> "-" Then
                        AppForm(objNbr).n -= 1
                        AppForm(objNbr).cssClass = "itemred"
                    End If

                    ' AppForm(objNbr).Status &= "<br>" & bullet & "None Found" & vbCrLf
                End If

                cnt = 0
                foundflag = False
                statgood = ""
                statbad = ""
                nObjSeen = 0
            Next j

        End If

    End Sub

    Function buildObjSummary(strobj As String, cnt As Integer, nobjseen As Integer, statgood As String, statbad As String, ByRef css As String) As String

        ' if the count = -1, color is set to clear, count = 0, the color is set to green. IF count > 0 the color is red
        Dim txt As String

        If cnt = -1 Then
            txt = statgood
            css = ""       ' "itemclear"

        ElseIf cnt = 0 Then
            txt = String.Format("<span class=""boldtext"">({0}) objects with proper prefix</span>", nobjseen) & "<br />" & vbCrLf & statgood
            css = "itemgreen"
        Else
            txt = String.Format("<span class=""boldtext"">({0}) without proper prefix</span>", cnt, strobj) & "<br />" & vbCrLf
            txt &= statbad
            css = "itemred"

            If statgood.Trim.Length > 0 Then
                txt &= String.Format("<span class=""boldtext"">({0}) objects with proper prefix</span>", nobjseen) & "<br />" & vbCrLf
                txt &= statgood

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
        Dim th As String = "<tr> <th class=""req"" > Req </th> <th class=""req"" >OK </th>  <th class=""titlecol""> Item </th>   <th class=""statuscol"" > Status </th>  <th class=""ptcol"" > Possible<br /> Pts. </th>   <th class=""ptcol"">  Your<br /> Score</th>  <th class=""commentcol""> Comment </th> </tr>" & vbCrLf
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
                    strReport = AppendToReport(strReport, 3, "Development Environment", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                    strReport = AppendToReport(strReport, 2, " - SLN File", SAssignment.hasSLN, errcnt, errComment, "hasSLN", SAssignment.TotalScore)
                    strReport = AppendToReport(strReport, 2, " - vbProj File", SAssignment.hasVBproj, errcnt, errComment, "hasvbProj", SAssignment.TotalScore)
                    strReport = AppendToReport(strReport, 2, " - VB Version", SAssignment.VBVersion, errcnt, errComment, "hasVBVersion", SAssignment.TotalScore)
                    strReport = AppendToReport(strReport, 10, "General Help", SAssignment.VBVersion, errcnt, errComment, "", SAssignment.TotalScore)
                    errcnt = 0
                    errComment = ""

                    ReDim varlist(7)
                    varlist = {"hasSplashScreen", "hasAboutBox", "InfoAppTitle", "InfoDescription", "InfoCompany", "InfoProduct", "InfoTrademark", "InfoCopyright"}

                    If SomeVarDisplayed(varlist) Then

                        strReport = AppendToReport(strReport, 3, "Application Info", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)

                        strReport = AppendToReport(strReport, 2, " - Splash Screen", SAssignment.hasSplashScreen, errcnt, errComment, "hasSplashScreen", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - About Box", SAssignment.hasAboutBox, errcnt, errComment, "hasAboutBox", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Application Title", SSummary(EnSummary.InfoAppTitle), errcnt, errComment, "InfoAppTitle", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Description", SSummary(EnSummary.InfoDescription), errcnt, errComment, "InfoDescription", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Company", SSummary(EnSummary.InfoCompany), errcnt, errComment, "InfoCompany", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Product", SSummary(EnSummary.InfoProduct), errcnt, errComment, "InfoProduct", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Trademark", SSummary(EnSummary.InfoTrademark), errcnt, errComment, "InfoTrademark", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Copyright", SSummary(EnSummary.InfoCopyright), errcnt, errComment, "InfoCopyright", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 10, " Application Info", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        errcnt = 0
                        errComment = ""
                    End If

                    ReDim varlist(1)
                    varlist = {"OptionStrict", "OptionExplicit"}

                    If SomeVarDisplayed(varlist) Then

                        strReport = AppendToReport(strReport, 3, "Compile Options", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Option Strict", SAssignment.OptionStrict, errcnt, errComment, "OptionStrict", SAssignment.TotalScore)   ' ????????????????????????????/ jhg
                        strReport = AppendToReport(strReport, 2, " - Option Explicit", SAssignment.OptionExplicit, errcnt, errComment, "OptionExplicit", SAssignment.TotalScore) ' ????????????????????????????/ jhg
                        strReport = AppendToReport(strReport, 10, "Options", SSummary(EnSummary.InfoCopyright), errcnt, errComment, "", SAssignment.TotalScore)

                    End If
                    strReport &= "</table>" & vbCrLf & vbCrLf

                Case "Debugging"
                    ' This has not been implemented. I am not sure if it is possible, and if so how to determine if these features have been set within the student's environment.
                    strReport &= String.Format(h3, "Debugging")

                    strReport &= starttable
                    strReport &= th
                    strReport = AppendToReport(strReport, 3, "Debugging", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                    strReport = AppendToReport(strReport, 2, " - BreakPoints", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                    strReport = AppendToReport(strReport, 2, " - Watch Variables", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                    strReport = AppendToReport(strReport, 10, "Debugging", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
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

                            strReport = AppendToReport(strReport, 3, "Design Time Form Properties", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)

                            strReport = AppendToReport(strReport, 2, " - Form Text Property", AppForm(EnForm.FormText), errcnt, errComment, "ChangeFormText", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - Accept Button", AppForm(EnForm.FormAcceptButton), errcnt, errComment, "SetFormAcceptButton", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - Cancel Button", AppForm(EnForm.FormCancelButton), errcnt, errComment, "SetFormCancelButton", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - Start Position", AppForm(EnForm.FormStartPosition), errcnt, errComment, "ModifyStartPosition", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - Non-Gray Form Color", AppForm(EnForm.FormBackColor), errcnt, errComment, "ChangeFormColor", SAssignment.TotalScore)
                            '         strReport = AppendToReport(strReport, 2, " - Form Load Method", AppForm(enForm.FormLoadMethod, errcnt, errComment, "UtilizeFormLoadMethod", SAssignment.TotalScore)

                            strReport = AppendToReport(strReport, 10, "", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                            errcnt = 0
                            errComment = ""
                        End If

                        ReDim varlist(12)
                        varlist = {"IncludeFrmInFormName", "ButtonObj", "ObjTextbox", "ActiveLabel", "NonactiveLabel", "ObjCombobox", "ObjListbox", "ObjRadioButton", "ObjCheckbox", "ObjGroupBox", "ObjOpenFileDialog", "ObjSaveFileDialog", "ObjWebBrowser"}

                        If SomeVarDisplayed(varlist) Then
                            strReport = AppendToReport(strReport, 3, "Form Object Names Incorporate Object Prefix", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)

                            strReport = AppendToReport(strReport, 2, " - Form (frm)", AppForm(EnForm.FormName), errcnt, errComment, "IncludeFrmInFormName", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - Buttons (btn)", AppForm(EnForm.ObjButton), errcnt, errComment, "ButtonObj", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - Textboxes (txt)", AppForm(EnForm.ObjTextbox), errcnt, errComment, "TextboxObj", SAssignment.TotalScore)
                            '  strReport = AppendToReport(strReport, 2, " - Labels (lbl)", AppForm(enForm.objLabel), errcnt, errComment, "", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - Active Labels (lbl)", AppForm(EnForm.ObjActiveLabel), errcnt, errComment, "ActiveLabels", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - NonActive Labels (no prefix needed)", AppForm(EnForm.ObjNonactiveLabel), errcnt, errComment, "NonActiveLabels", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - Combobox (cbx)", AppForm(EnForm.ObjCombobox), errcnt, errComment, "ComboBoxObj", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - Listbox (lbx)", AppForm(EnForm.ObjListbox), errcnt, errComment, "ListBoxObj", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - Radiobutton (rbn)", AppForm(EnForm.ObjRadioButton), errcnt, errComment, "RadioButtonObj", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - Checkbox (cbx)", AppForm(EnForm.ObjCheckbox), errcnt, errComment, "CheckBoxObj", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - Groupbox (gbx)", AppForm(EnForm.ObjGroupBox), errcnt, errComment, "GroupBoxObj", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - OpenFileDialog (ofd)", AppForm(EnForm.ObjOpenFileDialog), errcnt, errComment, "ObjOpenFileDialog", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - SaveFileDialog (sfd)", AppForm(EnForm.ObjSaveFileDialog), errcnt, errComment, "ObjSaveFileDialog", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 2, " - WebBrowser (wb)", AppForm(EnForm.ObjWebBrowser), errcnt, errComment, "WebBrowserObj", SAssignment.TotalScore)
                            strReport = AppendToReport(strReport, 10, "Form Objects", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
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
                        strReport = AppendToReport(strReport, 3, "Use of Comments", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - First Line of Sub/Function", SSummary(EnSummary.CommentSub), errcnt, errComment, "CommentSubs", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Prior to IF", SSummary(EnSummary.CommentIF), errcnt, errComment, "CommentIF", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Prior to For", SSummary(EnSummary.CommentFor), errcnt, errComment, "CommentFOR", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Prior to Do", SSummary(EnSummary.CommentDo), errcnt, errComment, "CommentDO", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Prior to While", SSummary(EnSummary.CommentWhile), errcnt, errComment, "CommentWHILE", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Prior to Select Case", SSummary(EnSummary.CommentSelect), errcnt, errComment, "CommentSELECT", SAssignment.TotalScore)

                        If errComment.Length > 0 Then
                            '  errComment = ""
                            dummyItem.cssClass = "itemred"
                            strReport = AppendToReport(strReport, 10, "General Comments on Comments", dummyItem, errcnt, errComment, "CommentGeneral", SAssignment.TotalScore)
                        Else
                            strReport = AppendToReport(strReport, 10, "Use of Comments", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        End If

                        errcnt = 0
                        errComment = ""
                    End If


                    ReDim varlist(2)
                    varlist = {"VarArrays", "VarLists", "VarStructures"}

                    If SomeVarDisplayed(varlist) Then

                        '     If .VarArrays.showVar OrElse .VarLists.showVar OrElse .VarStructures.showVar Then
                        strReport = AppendToReport(strReport, 3, "Data Structures", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Arrays", SSummary(EnSummary.VarArrays), errcnt, errComment, "VarArrays", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Lists", SSummary(EnSummary.VarLists), errcnt, errComment, "VarLists", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Structures", SSummary(EnSummary.VarStructures), errcnt, errComment, "VarStructures", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 10, "Data Structures", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        errcnt = 0
                        errComment = ""
                    End If


                    ReDim varlist(4)
                    varlist = {"VarString", "VarInteger", "VarDecimal", "VarDate", "VarBoolean"}

                    If SomeVarDisplayed(varlist) Then

                        strReport = AppendToReport(strReport, 3, "Variable Data Types - Checking to see which data types are used ", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - String", SSummary(EnSummary.VarString), errcnt, errComment, "VarString", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Integer", SSummary(EnSummary.VarInteger), errcnt, errComment, "VarInteger", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Decimal/Double", SSummary(EnSummary.VarDecimal), errcnt, errComment, "VarDecimal", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Date", SSummary(EnSummary.VarDate), errcnt, errComment, "VarDate", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Boolean", SSummary(EnSummary.VarBoolean), errcnt, errComment, "VarBoolean", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 10, "Variable Data Types", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        errcnt = 0
                        errComment = ""
                    End If


                    ReDim varlist(18)
                    varlist = {"LogicElse", "LogicElseIF", "LogicNestedIF", "LogicNestedFOR", "LogicConvertToString", "LogicStringFormat", "LogicStringFormatParameters", "LogicConcatination", "LogicCaseInsensitive", "LogicTryCatch", "LogicComplexConditions", "LogicStreamReader", "LogicStreamReaderClose", "LogicStreamWriter", "LogicStreamWriterClose"}

                    If SomeVarDisplayed(varlist) Then


                        strReport = AppendToReport(strReport, 3, "Program Logic - Checking to see if each Programming Control Stucture is used or not", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Else", SSummary(EnSummary.LogicElse), errcnt, errComment, "LogicElse", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - ElseIF", SSummary(EnSummary.LogicElseIF), errcnt, errComment, "LogicElseIF", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Nested IF", SSummary(EnSummary.LogicNestedIF), errcnt, errComment, "LogicNestedIF", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Nested For/Do", SSummary(EnSummary.LogicNestedFor), errcnt, errComment, "LogicNestedFOR", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Convert to String (cStr or .toString)", SSummary(EnSummary.LogicConvertToString), errcnt, errComment, "LogicConvertToString", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Format String (.toString() or String.Format())", SSummary(EnSummary.LogicStringFormat), errcnt, errComment, "LogicStringFormat", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Template Parameters", SSummary(EnSummary.LogicStringFormatParameters), errcnt, errComment, "LogicStringFormatParameters", SAssignment.TotalScore)
                        '   strReport = AppendToReport(strReport, 2, " - ByRef Parameters", ssummary(ensummary.LogicByRef), errcnt, errComment, "LogicByRef", SAssignment.TotalScore)
                        '   strReport = AppendToReport(strReport, 2, " - Optional Parameters", ssummary(ensummary.LogicOptional), errcnt, errComment, "LogicOptional", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Concatenation", SSummary(EnSummary.LogicConcatination), errcnt, errComment, "LogicConcatination", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Case Insensitve ", SSummary(EnSummary.LogicCaseInsensitive), errcnt, errComment, "LogicCaseInsensitive", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Try ... Catch", SSummary(EnSummary.LogicTryCatch), errcnt, errComment, "LogicTryCatch", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Complex Conditions", SSummary(EnSummary.LogicComplexConditions), errcnt, errComment, "LogicComplexConditions", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Req Use of a StreamReader", SSummary(EnSummary.LogicStreamReader), errcnt, errComment, "LogicStreamReader", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Req matching StreamReader.Close", SSummary(EnSummary.LogicStreamReaderClose), errcnt, errComment, "LogicStreamReaderClose", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Req Use of a StreamWriter", SSummary(EnSummary.LogicStreamWriter), errcnt, errComment, "LogicStreamWriter", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Req matching StreamWriter.Close", SSummary(EnSummary.LogicStreamWriterClose), errcnt, errComment, "LogicStreamWriterClose", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 10, "Program Logic", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        errcnt = 0
                        errComment = ""


                    End If


                    ReDim varlist(2)
                    varlist = {"SystemIO", "SystemNet", "SystemDB"}

                    If SomeVarDisplayed(varlist) Then

                        '    If .SystemIO.showVar OrElse .SystemNet.showVar OrElse .SystemDB.showVar Then
                        strReport = AppendToReport(strReport, 3, "Imports", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - System.IO", SSummary(EnSummary.SystemIO), errcnt, errComment, "SystemIO", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - System.Net", SSummary(EnSummary.SystemNet), errcnt, errComment, "SystemNet", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - System.DB", SSummary(EnSummary.SystemDB), errcnt, errComment, "SystemDB", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 10, "Imports", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        errcnt = 0
                        errComment = ""

                    End If


                    ReDim varlist(5)
                    varlist = {"LogicSub", "LogicOptional", "LogicByRef", "LogicMultipleForms", "LogicModule", "LogicFormLoad"}

                    If SomeVarDisplayed(varlist) Then

                        '    If .SystemIO.showVar OrElse .SystemNet.showVar OrElse .SystemDB.showVar Then
                        strReport = AppendToReport(strReport, 3, "Subs / Functions", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Subs", SSummary(EnSummary.LogicSub), errcnt, errComment, "LogicSub", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Optional Variables", SSummary(EnSummary.LogicOptional), errcnt, errComment, "LogicOptional", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - ByRef Variables", SSummary(EnSummary.LogicByRef), errcnt, errComment, "LogicByRef", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Multiple Forms", SSummary(EnSummary.LogicMultipleForms), errcnt, errComment, "LogicMultipleForms", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Include Module", SSummary(EnSummary.LogicModule), errcnt, errComment, "LogicModule", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 2, " - Form Load Method", SSummary(EnSummary.LogicFormLoad), errcnt, errComment, "LogicFormLoad", SAssignment.TotalScore)
                        strReport = AppendToReport(strReport, 10, "Subs / Functions", dummyItem, errcnt, errComment, "", SAssignment.TotalScore)
                        errcnt = 0
                        errComment = ""

                    End If


                    strReport &= "</table>" & vbCrLf & vbCrLf


                    strReport = strReport.Replace("[TLOC]", TotalLinesOfCode.ToString("n0") & " Lines of Code (not including comments)")

                    '  TotalScore = (SAssignment.TotalScore / TotalPossiblePts).ToString("p1")
                    SAssignment.strTotalScore = SAssignment.TotalScore.ToString("n1") & " deduction out of " & TotalPossiblePts.ToString("n1") & " possible points = " & (SAssignment.TotalScore / TotalPossiblePts).ToString("p1")

                    strReport = strReport.Replace("[SCORE]", SAssignment.strTotalScore & vbCrLf)


                Case Else
                    If key.StartsWith("Assessment Results for") Then
                        strReport &= "<h2>" & key & "</h2>" & vbCrLf
                    Else
                        '  BuildSummaryDetail()
                        MessageBox.Show("BuildReport received an unknown key = " & key)
                    End If
            End Select
        End With
    End Sub


    Sub BuildSummaryDetail(Assign As AssignmentInfo, AppForm() As MyItems, AppSummary() As MyItems)

        strFacReport = InitializeFacultyReport()

        BuildReport("Application Level", Assign, AppForm, AppSummary, strFacReport)
        BuildReport("New Project File", Assign, AppForm, AppSummary, strFacReport) ' the AppForm(0) is just a placeholder
        BuildReport("Assessment Results for " & ReturnLastField("", "\"), Assign, AppForm, AppSummary, strFacReport)
        BuildReport("Form Objects", Assign, AppForm, AppSummary, strFacReport)
        BuildReport("Coding Standards", Assign, AppForm, AppSummary, strFacReport)

    End Sub



    Public Function FindIntegratedScore(SAssignment As AssignmentInfo, AppForm() As MyItems, SSummary() As MyItems) As Decimal

        Dim T As Decimal = 0
        Dim i As Integer
        Dim setting As MySettings

        For i = 0 To SSummary.GetUpperBound(0)
            setting = Find_Setting(EnSummaryName(i), "FindIntegratedScore")
            With SSummary(i)

                If setting.Req And (setting.PtsPerError = 0 Or .n = 0) Then
                    '         Beep()
                End If

                If setting.Req And setting.MaxPts <> 0 Then        '    And .cssClass = "itemred" Then
                    If EnSummaryName(i).StartsWith("Comment") Then
                        If .n >= 0 Then
                            .YourPts = Math.Min(setting.MaxPts - setting.PtsPerError * Math.Min(setting.MaxPts, .n), setting.MaxPts)  ' this awards points for having instances 
                        Else
                            .YourPts = Math.Min(setting.PtsPerError * Math.Min(setting.MaxPts, -.n) - setting.MaxPts, setting.MaxPts)  ' this awards points for having instances 
                        End If
                    Else
                        .YourPts = Math.Min(setting.PtsPerError * Math.Max(0, .n), setting.MaxPts)
                    End If
                        T += .YourPts
                    Else
                        .YourPts = 0
                    End If
            End With
        Next

        For i = 0 To AppForm.GetUpperBound(0)
            setting = Find_Setting(EnFormNames(i), "FindIntegratedScore")
            With AppForm(i)
                If setting.Req And setting.MaxPts <> 0 Then        '    And .cssClass = "itemred" Then
                    .YourPts = Math.Min(setting.PtsPerError * Math.Max(0, .n), setting.MaxPts)
                    T += .YourPts
                Else
                    .YourPts = 0
                End If
            End With
        Next


        Return T
    End Function


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



    Function AppendToReport(rpt As String, Template As Integer, Title As String, topic As MyItems, ByRef errcnt As Integer, ByRef err As String, Item As String, ByRef total As Decimal) As String
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
        If (HideGray = "OnlyReq" And (Setting.Req Or Item.Length = 0)) Or (HideGray <> "OnlyReq" And (showvar Or Setting.Req)) Then
            With topic

                ' Calculate the grade

                If Setting.Req And Setting.MaxPts <> 0 Then
                    If .n >= 0 Then
                        .YourPts = Math.Min(Setting.PtsPerError * Math.Max(0, .n), Setting.MaxPts)
                    Else
                        .YourPts = Setting.MaxPts - Math.Min(Setting.PtsPerError * Math.Max(0, -.n), Setting.MaxPts)
                    End If
                    total += .YourPts
                Else
                    .YourPts = 0
                End If


                ' check to see if item is required
                req = ""
                If Setting.Req Then
                    If Setting.Req Then req = "*" Else req = ""
                    If .cssClass = "itemred" Then
                        isok = "&#x2717;"              ' load a X symbol
                    ElseIf .cssClass = "itemgreen" Then
                        isok = "&#x2713;"               ' Load a check symbol
                    Else
                        isok = ""
                    End If
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
            End With

            ' Insert the HTML segment into the file
            If (HideGray = "OnlyReq" And (Setting.Req Or Item.Length = 0)) Or (HideGray <> "OnlyReq") Then
                rpt &= s
            End If
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



    'Sub PopulateNonCheckCSS_Summary(ByRef AppSum() As MyItems, ByRef Assign As AssignmentInfo)
    '    Dim nc As String = ""                       ' Non-checked property
    '    Dim c As String = "ncWhite"    ' Checked property
    '    Dim item As New MySettings

    '    ' ----------------------------------------------------------

    '    If HideGray = "Gray" Then
    '        nc = "ncGray"
    '    ElseIf HideGray = "Hide" Then
    '        nc = "ncHide"
    '    Else
    '        nc = "ncWhite"
    '    End If


    '    ' ----------------------------------------------------------
    '    setchecked(Find_Setting("InfoAppTitle", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.InfoAppTitle), nc, c)
    '    setchecked(Find_Setting("InfoDescription", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.InfoDescription), nc, c)
    '    setchecked(Find_Setting("InfoCompany", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.InfoCompany), nc, c)
    '    setchecked(Find_Setting("InfoProduct", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.InfoProduct), nc, c)
    '    setchecked(Find_Setting("InfoTrademark", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.InfoTrademark), nc, c)
    '    setchecked(Find_Setting("InfoCopyright", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.InfoCopyright), nc, c)
    '    setchecked(Find_Setting("", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.InfoGUID), nc, c)


    '    setchecked(Find_Setting("CommentSubs", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.CommentSub), nc, c)
    '    setchecked(Find_Setting("CommentIF", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.CommentIF), nc, c)
    '    setchecked(Find_Setting("CommentFOR", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.CommentFor), nc, c)
    '    setchecked(Find_Setting("CommentDO", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.CommentDo), nc, c)
    '    setchecked(Find_Setting("CommentWHILE", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.CommentWhile), nc, c)
    '    setchecked(Find_Setting("CommentSELECT", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.CommentSelect), nc, c)

    '    setchecked(Find_Setting("VarString", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.VarString), nc, c)
    '    setchecked(Find_Setting("VarBoolean", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.VarBoolean), nc, c)
    '    setchecked(Find_Setting("VarInteger", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.VarInteger), nc, c)
    '    setchecked(Find_Setting("VarDecimal", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.VarDecimal), nc, c)
    '    setchecked(Find_Setting("VarDate", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.VarDate), nc, c)

    '    setchecked(Find_Setting("VarArray", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.VarArrays), nc, c)
    '    setchecked(Find_Setting("VarList", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.VarLists), nc, c)
    '    setchecked(Find_Setting("VarStructure", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.VarStructures), nc, c)

    '    setchecked(Find_Setting("VariablePrefixes", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.VarPrefixes), nc, c)

    '    ' setchecked(Find_Setting("", "PopulatenonCheckCSS_summary").Req, AppSum(enSummary.LogicFlowControl), nc, c)
    '    setchecked(Find_Setting("LogicIF", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.LogicIF), nc, c)
    '    setchecked(Find_Setting("LogicFOR", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.LogicFor), nc, c)
    '    setchecked(Find_Setting("LogicDO", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.LogicDo), nc, c)
    '    setchecked(Find_Setting("LogicWHILE", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.LogicWhile), nc, c)
    '    setchecked(Find_Setting("LogicSelectCase", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.LogicSelectCase), nc, c)
    '    setchecked(Find_Setting("LogicElse", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.LogicElse), nc, c)
    '    setchecked(Find_Setting("LogicElseIF", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.LogicElseIF), nc, c)
    '    setchecked(Find_Setting("LogicTryCatch", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.LogicTryCatch), nc, c)
    '    setchecked(Find_Setting("LogicStreamReader", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.LogicStreamReader), nc, c)
    '    setchecked(Find_Setting("LogicStreamWriter", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.LogicStreamWriter), nc, c)
    '    setchecked(Find_Setting("LogicStreamReaderClose", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.LogicStreamReaderClose), nc, c)
    '    setchecked(Find_Setting("LogicStreamWriterClose", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.LogicStreamWriterClose), nc, c)
    '    setchecked(Find_Setting("LogicSub", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.LogicSub), nc, c)
    '    '  setchecked(Find_Setting("", "PopulatenonCheckCSS_summary").Req, AppSum(enSummary.LogicFunction), nc, c)
    '    setchecked(Find_Setting("LogicOptional", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.LogicOptional), nc, c)
    '    setchecked(Find_Setting("LogicByRef", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.LogicByRef), nc, c)
    '    setchecked(Find_Setting("LogicConvertToString", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.LogicCStr), nc, c)
    '    '   setchecked(Find_Setting("", "PopulatenonCheckCSS_summary").Req, AppSum(enSummary.LogicToString), nc, c)
    '    setchecked(Find_Setting("LogicStringFormat", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.LogicStringFormat), nc, c)



    '    '     setchecked(Find_Setting("Req Variable Prefixes", "PopulatenonCheckCSS_summary").Req, AppSum(enSummary.LogicVarPrefixes), nc, c)
    '    setchecked(Find_Setting("LogicNestedIF", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.LogicNestedIF), nc, c)
    '    setchecked(Find_Setting("LogicNestedFOR", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.LogicNestedFor), nc, c)

    '    '        setchecked(find_Setting("", "PopulatenonCheckCSS_summary").Req, AppSum(enSummary.LogicStringFormatting), nc, c)
    '    setchecked(Find_Setting("LogicComplexConditions", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.LogicComplexConditions), nc, c)
    '    setchecked(Find_Setting("LogicCaseInsensitive", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.LogicCaseInsensitive), nc, c)
    '    setchecked(Find_Setting("LogicStringFormatParameters", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.LogicStringFormat), nc, c)
    '    setchecked(Find_Setting("LogicConcatination", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.LogicConcatination), nc, c)

    '    '     setchecked(Find_Setting("UtilizeLogicFormLoad", "PopulatenonCheckCSS_summary").Req, AppSum(enSummary.FormLoadMethod), nc, c)

    '    setchecked(Find_Setting("SystemIO", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.SystemIO), nc, c)
    '    setchecked(Find_Setting("SystemNet", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.SystemNet), nc, c)
    '    setchecked(Find_Setting("SystemDB", "PopulatenonCheckCSS_summary").Req, AppSum(EnSummary.SystemDB), nc, c)


    '    setchecked(Find_Setting("OptionStrict", "PopulatenonCheckCSS_summary").Req, Assign.OptionStrict, nc, c)
    '    setchecked(Find_Setting("OptionExplicit", "PopulatenonCheckCSS_summary").Req, Assign.OptionExplicit, nc, c)
    '    ' ----------------------------------------------------------
    '    setchecked(Find_Setting("hasSLN", "PopulatenonCheckCSS_summary").Req, Assign.hasSLN, nc, c)
    '    setchecked(Find_Setting("hasvbProj", "PopulatenonCheckCSS_summary").Req, Assign.hasVBproj, nc, c)
    '    setchecked(Find_Setting("hasSplashScreen", "PopulatenonCheckCSS_summary").Req, Assign.hasSplashScreen, nc, c)
    '    setchecked(Find_Setting("hasAboutBox", "PopulatenonCheckCSS_summary").Req, Assign.hasAboutBox, nc, c)
    '    setchecked(Find_Setting("LogicModule", "PopulatenonCheckCSS_summary").Req, Assign.Modules, nc, c)
    '    ' ----------------------------------------------------------

    '    ' ----------------------------------------------------------
    'End Sub

    'Sub setchecked(chk As Boolean, ByRef obj As MyItems, nc As String, c As String)
    '    If Not chk Then
    '        obj.cssNonChk = nc
    '        obj.req = False
    '    Else
    '        obj.cssNonChk = c
    '        obj.req = True
    '    End If
    'End Sub

End Module
