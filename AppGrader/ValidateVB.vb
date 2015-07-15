Imports System.IO
'Imports SharpCompress.Archive
'Imports SharpCompress.Common


Module ValidateVB

    Public ValidateReport As String
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


    Public Sub CloseFacReport(path As String, Project As String)
        Dim sr As StreamReader
        Dim sw As StreamWriter

        sr = File.OpenText(Application.StartupPath & "\templates\rptFacFooter.html")
        strFacReport &= sr.ReadToEnd
        sr.Close()

        sw = File.CreateText((strOutputPath & "\FacultySummaryReport.html"))
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

    Sub CheckSLNvbProj(ByRef AppInfo As Assignment, ByRef AppSum As Assignment.AppSummary)
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
        Else
            AppInfo.hasVBproj.Status = vbFalse.ToString
            AppInfo.hasVBproj.cssClass = "itemred"
        End If

        AppInfo.hasVBproj.Status = hasVBProjFile.ToString
    End Sub



    Sub CheckAPPInfo2(AppDir As String, ByRef SSummary As Assignment.AppSummary)
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

            With SSummary
                s = s2
                .InfoAppTitle.Status = returnBetween(s, "<Assembly: AssemblyTitle(""", """)>", True)
                s = s2
                .InfoDescription.Status = returnBetween(s, "<Assembly: AssemblyDescription(""", """)>", True)
                s = s2
                .InfoCompany.Status = returnBetween(s, "<Assembly: AssemblyCompany(""", """)>", True)
                s = s2
                .InfoProduct.Status = returnBetween(s, "<Assembly: AssemblyProduct(""", """)>", True)
                s = s2
                .InfoCopyright.Status = returnBetween(s, "<Assembly: AssemblyCopyright(""", """)>", True)
                s = s2
                .InfoTrademark.Status = returnBetween(s, "<Assembly: AssemblyTrademark(""", """)>", True)
                s = s2
                .InfoGUID.Status = returnBetween(s, "<Assembly: Guid(""", """)>", True)
            End With
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


    Sub CheckForOptions2(path As String, filesinbuild As List(Of String), ByRef item As Assignment.MyItems, optiontype As String)   ' ByRef HasOptionStrict As Boolean, ByRef HasOptionStrictOff As Boolean)
        ' Checks the designed directory fn for the directive Option Strict / Option Explicit

        Dim source As String = ""
        Dim tmp As String = ""

        With item
            .Status = "Not Set"

            ' Check each file plus Application.Designer. If it is set of off, that is a big problem and overrides any point where it is set on

            Dim i As Integer = 0

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
                        Exit For
                    End If
                End If
            Next filename
        End With
    End Sub



    Function CheckForComments2(filesource As String, ByRef SSummary As Assignment.AppSummary, ByRef AppForm As Assignment.AppForm, sender As System.ComponentModel.BackgroundWorker) As Integer

        Dim worker As System.ComponentModel.BackgroundWorker = DirectCast(sender, System.ComponentModel.BackgroundWorker)
        Dim i As Integer
        Dim ncomments As Integer
        Dim isInIf As Integer
        Dim isInFor As Integer
        Dim isinSub As Boolean
        Dim isinFunction As Boolean
        Dim LastlineComment As Integer
        Dim ss() As String
        Dim lineno() As Integer
        Dim n As Integer
        Dim delim() As String = {vbCrLf}
        Dim firstword As String = ""
        Dim PreviousWord As String = ""
        Dim comBad As String = ""
        Dim comGood As String = ""
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
        With SSummary
            ' set default to green
            .LogicIF.cssClass = "itemgreen"
            .LogicFor.cssClass = "itemgreen"
            .LogicDo.cssClass = "itemgreen"
            .LogicSelectCase.cssClass = "itemgreen"
            .LogicWhile.cssClass = "itemgreen"
            .LogicSub.cssClass = "itemgreen"

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
                        LastlineComment = i   ' Note, this points to the array element, not the file line

                    Case "CASE"
                        ' Looks for code the implements case insensitiveity
                        CheckForCaseInsensitivity(SSummary, ss(i), lineno(i))

                    Case "CATCH"

                    Case "DIM"    ' Tracks the declaration of variables
                        If ss(i).ToUpper.Contains(" AS INTEGER") Then
                            .VarInteger.n += 1
                            .VarInteger.Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf
                        End If
                        If ss(i).ToUpper.Contains(" AS DECIMAL") Or ss(i).ToUpper.Contains("DOUBLE") Then
                            .VarDecimal.n += 1
                            .VarDecimal.Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf
                        End If
                        If ss(i).ToUpper.Contains(" AS BOOLEAN") Then
                            .VarBoolean.n += 1
                            .VarBoolean.Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf
                        End If
                        If ss(i).ToUpper.Contains(" AS DATE") Then
                            .VarDate.n += 1
                            .VarDate.Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf
                        End If
                        If ss(i).ToUpper.Contains(" AS STRING") Then
                            .VarString.n += 1
                            .VarString.Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf
                        End If

                        If ss(i).ToUpper.Contains(") AS ") Then
                            .VarArrays.n += 1
                            .VarArrays.Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf
                        End If

                        If ss(i).ToUpper.Contains("LIST (OF") Then
                            .VarLists.n += 1
                            .VarLists.Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf
                        End If

                        If ss(i).ToUpper.Contains(" STRUCTURE ") Then
                            .VarStructures.n += 1
                            .VarStructures.Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf
                        End If

                        If ss(i).ToUpper.Contains(" AS STREAMREADER") Or ss(i).ToUpper.Contains(" AS NEW STREAMREADER") Then
                            .LogicStreamReader.n += 1
                            .LogicStreamReader.Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf

                            s = returnBetween(ss(i), "DIM ", "AS STREAMREADER", True).Trim

                            If srName.Length > 0 Then
                                ReDim Preserve srName(srName.Count)
                                srName(srName.GetUpperBound(0)) = s
                            End If

                        End If
                        If ss(i).ToUpper.Contains(" AS STREAMWRITER") Or ss(i).ToUpper.Contains(" AS NEW STREAMWRITER") Then
                            .LogicStreamWriter.n += 1
                            .LogicStreamWriter.Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf

                            s = returnBetween(ss(i), "DIM ", "AS STREAMWRITER", True).Trim

                            If swName.Length > 0 Then
                                ReDim Preserve swName(swName.Count)
                                swName(swName.GetUpperBound(0)) = s
                            End If

                        End If


                        ' The following tracks the opening of streamreaders and writers. Tracking closings is done above.



                    Case "DO"
                        .LogicDo.n += 1

                        ' accept a comment before or on the same line as the DO
                        With .CommentDo
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
                        .LogicElse.n += 1
                        .LogicElse.Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf

                    Case "ELSEIF"
                        .LogicElseIF.n += 1
                        .LogicElseIF.Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf

                        ' --------------------------------------------
                        ' Check for complex conditions
                        CheckForComplexConditions(SSummary, ss(i), lineno(i))

                        ' Looks for code the implements case insensitiveity
                        CheckForCaseInsensitivity(SSummary, ss(i), lineno(i))

                    Case "END FUNCTION"
                        isinFunction = False ' not really used. Potentially could track nested functions


                    Case "END IF"
                        isInIf -= 1  ' used to track nested IF
                        If isInIf < 0 Then
                            '       Beep()
                        End If

                    Case "END SELECT"

                    Case "END SUB"
                        isinSub = False     ' not really used. Potentially could track nested subs
                        '   ChangeFormText()     ' not sure what changed here ????????????????????? jhg 
                    Case "END TRY"

                    Case "END WITH"


                    Case "FOR"
                        .LogicFor.n += 1   ' counts number of For statements
                        isInFor += 1       ' tracks if we are in a For statement, to identify Nested For statements

                        ' check for nested For
                        If isInFor > 1 Then
                            With .LogicNestedFOR
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
                        With .CommentFor
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
                            End If
                        End With


                    Case "IF"

                        .LogicIF.n += 1
                        isInIf += 1
                        ' Ignorore single line IF Statements. THE Is InIF counter is backed out
                        If ss(i).Trim.ToUpper.Contains(" THEN ") Then isInIf -= 1

                        ' checks for nested IF statements
                        If isInIf > 1 Then
                            With .LogicNestedIF
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
                        With .CommentIF
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
                            .SystemIO.Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf
                        Else
                            .SystemIO.Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - Not found</p>" & vbCrLf
                        End If

                        If ss(i).Contains("System.Net") Then
                            .SystemNet.Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf
                        Else
                            .SystemNet.Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - Not found</p>" & vbCrLf
                        End If

                        If ss(i).Contains("System.DB") Then
                            .SystemDB.Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf
                        Else
                            .SystemDB.Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - Not found</p>" & vbCrLf
                        End If


                    Case "LOOP"
                        ' --------------------------------------------
                        ' Check for complex conditions
                        CheckForComplexConditions(SSummary, ss(i), lineno(i))

                        ' Looks for code the implements case insensitiveity
                        CheckForCaseInsensitivity(SSummary, ss(i), lineno(i))

                    Case "MESSAGEBOX.SHOW", "MSGBOX"
                        With .LogicMessageBox
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
                        .VarStructures.n += 1
                        .VarStructures.Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf

                    Case "SELECT"
                        .LogicSelectCase.n += 1

                        ' check for comment. Accept on line before or on the same line.
                        With .CommentSelect
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
                            End If
                        End With

                        ' --------------------------------------------
                        ' Looks for code the implements case insensitiveity
                        CheckForCaseInsensitivity(SSummary, ss(i), lineno(i))

                    Case "SUB", "PUBLIC SUB", "PRIVATE SUB", "FUNCTION", "PUBLIC FUNCTION", "PRIVATE FUNCTION"

                        .LogicSub.n += 1

                        ' Check for a comment in first line of sub/function. This accepts it as the previous line
                        If i < ss.GetUpperBound(0) Then
                            With .CommentSub
                                If (ss(i - 1).StartsWith("'") Or ss(i + 1).StartsWith("'")) Then
                                    .bad &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & TrimAfter(ss(i), "(", True) & "</p>" & vbCrLf
                                    .cssClass = "itemred"
                                Else
                                    .good &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & TrimAfter(ss(i), "(", True) & "</p>" & vbCrLf
                                End If
                            End With

                        End If

                        ' ------------------------------------------
                        ' check for optional parameters
                        With .LogicOptional
                            If ss(i).Contains(" Optional ") Then
                                .Status = TrimAfter(ss(i), "(", True) & " defined in line (" & lineno(i).ToString & ") has an Optional Parameter" & " </p>" & vbCrLf
                                .n += 1
                            End If
                        End With

                        ' -------------------------------------
                        ' check for byref parameters
                        With .LogicByRef
                            If ss(i).Contains(" ByRef ") Then
                                .Status = TrimAfter(ss(i), "(", True) & " defined in line (" & lineno(i).ToString & ") has a ByRef Parameter" & " <br />" & vbCrLf
                                .n += 1
                            End If
                        End With

                    Case "TRY"
                        ' just looks for the try. Should likely also look for the catch, but does not currently
                        .LogicTryCatch.n += 1
                        If .LogicTryCatch.n = 1 Then
                            .LogicTryCatch.Status &= "(" & lineno(i).ToString & ")"
                        Else
                            .LogicTryCatch.Status &= ", (" & lineno(i).ToString & ")"
                        End If

                        ' ========================================
                    Case "WHILE"
                        .LogicWhile.n += 1

                        ' Comment While
                        With .CommentWhile
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
                        .LogicStreamReaderClose.n += 1
                        .LogicStreamReaderClose.Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & " </p>" & vbCrLf
                        doneflag = True
                    End If
                Next j

                doneflag = False
                For j = 1 To swName.GetUpperBound(0)

                    If ss(i).ToUpper.Contains(swName(j).ToUpper.Trim & ".CLOSE") And Not doneflag Then
                        .LogicStreamWriterClose.n += 1
                        .LogicStreamWriterClose.Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & " </p>" & vbCrLf
                        doneflag = True
                    End If
                Next j

                ' =================================================================================
                ' Now we need to look at the contents of the whole line to check for other issues.

                ' Check for the FormLoad Method
                If ss(i).Contains("Handles MyBase.Load") Then
                    .LogicFormLoad.n += 1
                    .LogicFormLoad.Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") - " & ss(i) & "</p>" & vbCrLf
                End If

                ' -----------------------------------
                ' Check for String Concatination
                With .LogicConcatination
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
                        .LogicConvertToString.n += 1
                        If .LogicConvertToString.n = 1 Then
                            .LogicConvertToString.Status &= "(" & lineno(i).ToString & ")"
                        Else
                            .LogicConvertToString.Status &= ", (" & lineno(i).ToString & ")"
                        End If
                    End If

                    ' Check to see if code is using a format string with tostring
                    With .LogicToStringFormat
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
                        With .LogicStringFormat
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
                            .LogicStringFormatParameters.n += 1
                            .LogicStringFormatParameters.Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno(i).ToString & ") " & ss(i) & "</p>" & vbCrLf
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
            ProcessComment(.LogicSub, .CommentSub, "Subs / Functions", "CommentSubs")
            ProcessComment(.LogicIF, .CommentIF, "IF statements", "CommentIF")
            ProcessComment(.LogicFor, .CommentFor, "FOR statements", "CommentFOR")
            ProcessComment(.LogicDo, .CommentDo, "DO statements", "CommentDO")
            ProcessComment(.LogicWhile, .CommentWhile, "WHILE statements", "CommentWHILE")
            ProcessComment(.LogicSelectCase, .CommentSelect, "SELECT CASE statements", "CommentSELECT")
            ' ------------------------------------------------------------------------------------------------------------

            If .LogicCStr.n > 0 Or .LogicToString.n > 0 Then
                .LogicToString.Status = "<span class=""boldtext"">Converting to Strings </span><br />" & vbCrLf

                If .LogicCStr.n > 0 Then
                    .LogicToString.Status &= "<span class=""boldtext""><br />Using CStr() to convert to string </span><br /><br />" & vbCrLf
                End If

                If .LogicToString.n > 0 Then
                    .LogicToString.Status &= "<span class=""boldtext""><br />Using .ToString to convert to string </span><br /><br />" & vbCrLf
                End If
            End If
            ' ------------------------------------------------------------------------------------------------------------

            If .LogicToStringFormat.n > 0 Or .LogicStringFormat.n > 0 Then
                If .LogicToStringFormat.n > 0 Then
                    .LogicToStringFormat.Status &= "<span class=""boldtext""><br />Included .ToString(format) command </span><br /><br />" & vbCrLf
                End If
                If .LogicStringFormat.n > 0 Then
                    .LogicStringFormat.Status &= "<span class=""boldtext""><br />Included String.Format(Template)  </span><br /><br />" & vbCrLf
                End If
            End If
            ' ------------------------------------------------------------------------------------------------------------
            ' -------------------------------------------------------------------------------------------------------
            ' AppInfo
            ProcessReq(.InfoAppTitle, "Application Title not modified", "Application Title modified", "InfoAppTitle")
            ProcessReq(.InfoDescription, "Application Description not modified", "Application Description modified", "InfoDescription")
            ProcessReq(.InfoCompany, "Application Company Info not modified", "Application Company Info modified", "InfoCompany")
            ProcessReq(.InfoProduct, "Application Product Info not modified", "Application Product Info modified", "InfoProduct")
            ProcessReq(.InfoTrademark, "Application Trademark not modified", "Application Trademark modified", "InfoTrademark")
            ProcessReq(.InfoCopyright, "Application Copyright not modified", "Application Copyright modified", "InfoCopyright")

            'Compile Options
            'ProcessReq(.OptionStrict, "No System.IO found", "Imports System.IO found (Line No.)", "Include System.IO")
            'ProcessReq(.SystemNet, "No System.Net found", "Imports System.Net found (Line No.)", "Include System.Net")

            ' Form Design


            ' Imports
            ProcessReq(.SystemIO, "No System.IO found", "Imports System.IO found (Line No.)", "SystemIO")
            ProcessReq(.SystemNet, "No System.Net found", "Imports System.Net found (Line No.)", "SystemNet")
            ProcessReq(.SystemDB, "No System.DB found", "Imports System.DB found (Line No.)", "SystemDB")

            ' Vars
            ProcessReq(.VarArrays, "No Arrays declared", "Data Arrays declared (Line No.)", "VarArrays")
            ProcessReq(.VarLists, "No Lists declared", "Lists(of T) declared (Line No.)", "VarLists")
            ProcessReq(.VarStructures, "No Structures declared", "Data Structures defined (Line No.)", "VarStructures")
            ProcessReq(.VarString, "No String variables declared", "String variables declared (Line No.)", "VarString")
            ProcessReq(.VarInteger, "No Integer variables declared", "Integer variables declared (Line No.)", "VarInteger")
            ProcessReq(.VarDecimal, "No Decimal / Double variables declared", "Decimal / Double variables declared (Line No.)", "VarDecimal")
            ProcessReq(.VarDate, "No Date variables declared", "Date variables declared (Line No.)", "VarDate")
            ProcessReq(.VarBoolean, "No Boolean variables declared", "Boolean variables declared (Line No.)", "VarBoolean")

            ' Logic
            ProcessReq(.LogicWhile, "No While statements found", "While statements found (Line No.)", "LogicWHILE")
            ProcessReq(.LogicSelectCase, "No SelectCase statements found", "Select Case statements found (Line No.)", "LogicSelectCase")

            ProcessReq(.LogicConvertToString, "No Conversion to String (.CStr or .toString) found", "Conversion to String (.CStr or .toString) found (Line No.)", "LogicConvertToString")
            ProcessReq(.LogicStringFormat, "No String Formatting (.toString() or String.Format()) found", "String Formatting (.toString() or String.Format()) found (Line No.)", "LogicStringFormat")
            ProcessReq(.LogicStringFormatParameters, "No Parameterized string formats (string.format(f,{0}) found", "Parameterized String Formatting (string.format(f,{0}) found (Line No.)", "LogicStringFormatParameters")

            ProcessReq(.LogicConcatination, "No String Concatination found", "String Concatination found (Line No.)", "LogicConcatination")
            ProcessReq(.LogicCaseInsensitive, "No Case-Insensitive comparisons found", "Case-Insensitive comparisons found (Line No.)", "LogicCaseInsensitive")
            ProcessReq(.LogicComplexConditions, "No Complex Conditions (AND / OR / ANDALSO / ORELSE) found", "Complex Conditions (AND / OR / ANDALSO / ORELSE) found (Line No.)", "LogicComplexConditions")

            ProcessReq(.LogicElse, "No Else statements found", "Else statements found (Line No.)", "LogicElse")
            ProcessReq(.LogicElseIF, "No ElseIF statements found", "ElseIF statements found (Line No.)", "LogicElseIF")
            ProcessReq(.LogicNestedIF, "No Nested IF statements found", "Nested IF statements found (Line No.)", "LogicNestedIF")
            ProcessReq(.LogicNestedFOR, "No Nested FOR statements found", "Nested FOR statements found (Line No.)", "LogicNestedFOR")
            ProcessReq(.LogicOptional, "No Optional Sub / Fuction Parameters found", "Optional Sub / Fuction Parameters found (Line No.)", "LogicOptional")
            ProcessReq(.LogicByRef, "No Sub / Function ByRef Parameters found", "Sub / Function ByRef Parameters found (Line No.)", "LogicByRef")
            ProcessReq(.LogicTryCatch, "No Try ... Catch Statement Found", "Try ... Catch Statements found (Line No.)", "LogicTryCatch")
            ProcessReq(.LogicFormLoad, "No Form Load Method Found", "Form Load Method found (Line No.)", "LogicFormLoad")

            ProcessReq(.LogicStreamReader, "No StreamReaders found", "StreamReader found (Line No.)", "LogicStreamReader")
            ProcessReq(.LogicStreamReaderClose, "No StreamReader.Close found", "StreamReader.Close found (Line No.)", "LogicStreamReaderClose")
            ProcessReq(.LogicStreamWriter, "No StreamWriters found", "StreamWriters found (Line No.)", "LogicStreamWriter")
            ProcessReq(.LogicStreamWriterClose, "No StreamWriter.Close found", "StreamWriter.Close found (Line No.)", "LogicStreamWriterClose")
        End With
        Return ncomments
    End Function


    Sub ProcessComment(ByRef logictype As Assignment.MyItems, ByRef commenttype As Assignment.MyItems, construct As String, nm As String)
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

    Sub ProcessReq(ByRef type As Assignment.MyItems, NotFound As String, Found As String, nm As String)
        If type.n = 0 Then
            type.Status = "<span class=""boldtext"">" & NotFound & "</span><br />" & vbCrLf
            type.cssClass = "itemred"
        Else
            type.Status = "<span class=""boldtext"">" & Found & "</span><br /><br />" & vbCrLf & type.Status & " <br />" & vbCrLf
            type.cssClass = "itemgreen"
        End If

        ' nm = type.ToString.Substring(1)

        If Not Find_Setting(nm, "ProcessReq").Req Then
            type.cssClass = "itemclear"
        End If

    End Sub


    Sub CheckForComplexConditions(ByRef ssum As Assignment.AppSummary, s As String, lineno As Integer)
        ' This checks a line of code for evidence of complex consitions
        With ssum
            With .LogicComplexConditions
                If s.ToUpper.Contains(" AND ") Or s.ToUpper.Contains(" OR ") Then
                    .n += 1
                    .Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno.ToString & ") - " & s & "</p>" & vbCrLf
                    .cssClass = "itemgreen"
                End If
            End With
        End With

    End Sub

    Sub CheckForCaseInsensitivity(ByRef ssum As Assignment.AppSummary, s As String, lineno As Integer)
        ' This check the line of code for evidence of case insensitivity
        With ssum
            With .LogicCaseInsensitive
                If s.ToUpper.Contains(".TOUPPER") Or s.ToUpper.Contains(".tolower") Then
                    .n += 1
                    .Status &= "<p class=""hangingindent2"">" & bullet & "(" & lineno.ToString & ") - " & s & "</p>" & vbCrLf
                    .cssClass = "itemgreen"
                End If
            End With
        End With
    End Sub

    Function CheckForSplashScreen2(ByRef filesinbuild As List(Of String), ByRef item As Assignment.MyItems, RemoveFromList As Boolean) As Boolean
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
                    If RemoveFromList Then filesinbuild.Remove(filename)
                    Exit For
                End If
            Next
        End With
    End Function





    Sub CheckFormProperties2(filename As String, ByRef AppForm As Assignment.AppForm)
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
                        .FormBackColor.Status = "<span class=""boldtext"">Form color = " & TrimUpTo(ss, "System.Drawing.SystemColors. </span>")
                        .FormBackColor.cssClass = "itemgreen"
                    ElseIf s.Contains("System.Drawing.Color.FromArgb(") Then

                        ss = returnBetween(s, "Me.BackColor =", vbCrLf)
                        cArray = ss.Split(delim, StringSplitOptions.None)
                        Try
                            .FormBackColor.Status = "<span class=""boldtext"">Form color is nongray (#" & ReturnHexEquivalent(TrimAfter(cArray(1), ",", True)) & ReturnHexEquivalent(TrimAfter(cArray(2), ",", True)) & ReturnHexEquivalent(TrimAfter(cArray(3), ",", True)) & ") </span>"
                            .FormBackColor.cssClass = "itemgreen"
                        Catch
                            .FormBackColor.Status = "Form color = " & ss
                        End Try
                    ElseIf s.Contains("System.Drawing.Color.") Then
                        .FormBackColor.Status = "<span class=""boldtext"">Form color = " & returnBetween(s, "Me.BackColor = System.Drawing.Color.", vbCrLf) & "</span>"
                        .FormBackColor.cssClass = "itemgreen"
                    End If
                Else
                    .FormBackColor.Status = "<span class=""boldtext"">Form Color was not changed at design time (still gray) </span>"
                    .FormBackColor.cssClass = "itemred"
                End If


                If s.Contains("Me.Text") Then
                    .FormText.Status = "<span class=""boldtext"">" & returnBetween(s, "Me.Text = """, """") & "</span>"
                    .FormText.cssClass = "itemgreen"
                Else
                    .FormText.Status = "<span class=""boldtext"">The form text was not changed.</span>"
                    .FormText.cssClass = "itemred"

                End If

                If s.Contains("Me.StartPosition") Then
                    .FormStartPosition.Status = "<span class=""boldtext"">Set to: " & returnBetween(s, "Me.StartPosition = System.Windows.Forms.FormStartPosition.", vbCrLf) & "</span>"
                    .FormStartPosition.cssClass = "itemgreen"
                Else
                    .FormStartPosition.Status = "<span class=""boldtext"">Form StartPosition not Modified. </span>"
                    .FormStartPosition.cssClass = "itemred"
                End If


                ' accept and cancel button settings in *.resx file
                sr = New StreamReader(filename.Replace(".vb", ".resx"))
                s = sr.ReadToEnd
                sr.Close()
                If s.Contains("Me.AcceptButton") Then
                    .FormAcceptButton.Status = "<span class=""boldtext"">" & returnBetween(s, "Me.AcceptButton = ", vbCrLf) & "</span>"
                    .FormAcceptButton.cssClass = "itemgreen"
                Else
                    .FormAcceptButton.Status = "<span class=""boldtext"">Accept Button Property not set at design time.</span>"
                    .FormAcceptButton.cssClass = "itemred"
                End If

                If s.Contains("Me.CancelButton") Then
                    .FormCancelButton.Status = "<span class=""boldtext"">" & returnBetween(s, "Me.AcceptButton = ", vbCrLf) & "</span>"
                    .FormCancelButton.cssClass = "itemgreen"
                Else
                    .FormCancelButton.Status = "<span class=""boldtext"">Cancel Button Property not set at design time.</span>"
                    .FormCancelButton.cssClass = "itemred"
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




    Sub CheckForAboutBox2(ByRef filesinbuild As List(Of String), ByRef item As Assignment.MyItems, RemoveFromList As Boolean)
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



    Sub CheckForfrmPrefix2(filesinbuild As List(Of String), ByRef Item As Assignment.MyItems)
        Dim fn As String


        With Item
            .Status = "All files start with frm prefix"

            For Each filename As String In filesinbuild

                filename = filename.Replace("/", "\")
                fn = ReturnLastField(filename, "\")
                ' ------------------------------------------------
                If Not fn.ToLower.StartsWith("frm") Then
                    ' .Status = vbFalse.ToString
                    .Status = filename & " does not start with frm prefix." & vbCrLf
                End If
            Next
        End With
    End Sub


    Sub CheckForModules2(filesinbuild As List(Of String), ByRef item As Assignment.MyItems)
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


    Sub CheckObjectNaming2(filename As String, ByRef AppForm As Assignment.AppForm)
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
        Dim strObjects() As String = {"Form", "LabelActive", "LabelNonactive", "Button", "Textbox", "Listbox", "Combobox", "OpenFileDialog", "SaveFileDialog", "RadioButton", "CheckBox", "GroupBox", "WebBrowser", "WebClient", "OpenFileDialog", "SaveFileDialog"}

        Dim filesourceVB As String
        Dim filesource As String
        Dim IsActiveLabel As Boolean
        Dim tmp As String = ""
        ' ------------------------------------------------------------------------------------------------------------------

        '    ' read the source of the vb file to check to see if any of the objects not renamed are used in the source file
        Dim sr As New StreamReader(filename)
        'filesourceVB = sr.ReadToEnd
        'sr.Close()

        ' Now read the source of the designer file to extract the definition of the objects.
        sr = New StreamReader(filename)
        filesourceVB = sr.ReadToEnd
        sr.Close()

        Dim fn As String = filename.Replace(".vb", ".designer.vb")

        ' Check to see if the file exists. This avoids processing Modules and Classes.
        If File.Exists(fn) Then
            sr = New StreamReader(fn)
            filesource = sr.ReadToEnd
            sr.Close()

            Dim strLayout As String = returnBetween(filesource, "Private Sub InitializeComponent()", "Me.SuspendLayout()")
            Dim strObjOnForm() As String = Nothing
            Dim css As String = ""

            Try
                strObjOnForm = strLayout.Split(delim, StringSplitOptions.None)
                strObjOnForm.DropFirstElement()

            Catch ex As Exception
                MessageBox.Show("Error occured in CheckObjectNaming - " & ex.Message)
            End Try

            ' ----------------------------------------------------------------------------------------------------------------
            ' the code below is wrong. it looks at the file name, not the FormText
            tmp = ReturnLastField(filename, "\")
            If tmp.ToLower.StartsWith("frm") Then
                AppForm.FormName.Status = tmp & " starts with frm prefix"
                AppForm.FormName.cssClass = "itemgreen"
                AppForm.FormName.cnt = 0
            Else
                AppForm.FormName.Status = tmp & " does not start with the frm prefix"
                AppForm.FormName.cssClass = "itemred"
                AppForm.FormName.cnt = 0
            End If

            ' ----------------------------------------------------------------------------------------------------------------
            For Each strObj As String In strObjects  ' list of the types of objects we are interested in
                statgood = ""
                statbad = ""
                cnt = -1  ' indicates no instances yet
                foundflag = False
                tmpObj = strObj

                If strObj.StartsWith("Label") Then ' We are doing this to handle Active and Nonactive labels
                    strObj = "Label"
                End If

                For Each obj As String In strObjOnForm   ' list of objects on the form

                    If obj.Contains("= New System.Windows.Forms.") Then
                        Try
                            s1 = TrimUpTo(obj, "= New System.Windows.Forms.")
                            If s1.Contains("(") Then s1 = s1.Substring(0, s1.IndexOf("("))

                            If s1.Trim.ToUpper = strObj.ToUpper Then                       ' type of object we are looking at
                                s2 = obj    ' not sure if this is correct. I added it so s2 had an initial value.
                                If s2.Contains(" = ") Then s2 = obj.Substring(0, obj.IndexOf(" = "))

                                If filesource.IndexOf("Me." & s2 & ".Text = ") > -1 Then
                                    s3 = returnBetween(filesource, "Me." & s2 & ".Text = ", vbCrLf)                   ' Text Value
                                Else
                                    s3 = ""
                                End If
                                ' -----------------------------------------------------------------------------------------------

                                If s1 = "Label" Then    ' check to see if it is active
                                    If filesourceVB.IndexOf(s2) > -1 Then
                                        IsActiveLabel = True
                                    Else
                                        IsActiveLabel = False
                                    End If
                                End If

                                ' --------------------------------------------------------------------------------------------------
                                If cnt = -1 Then cnt = 0 ' reset it so we count those without proper prefixes. cnt = 0 means no errors

                                foundflag = True
                                '       If Not s1.Trim = "Button" Then s3 = ""
                                If s3.Length > 30 Then s3 = s3.Substring(0, 30) & " ..."

                                If IsActiveLabel And tmpObj = "LabelActive" Then
                                    If s2.StartsWith(s1) Then                  ' Name of the object
                                        ' this is a problem - object not renamed
                                        statbad &= arrow & s2 & " [" & s3 & "] <br />" & vbCrLf
                                        cnt = cnt + 1
                                    Else
                                        statgood &= bullet & s2 & " [" & s3 & "] <br />" & vbCrLf
                                    End If
                                End If

                                If tmpObj <> "LabelActive" Then
                                    If s2.StartsWith(s1) And tmpObj <> "LabelNonactive" Then                  ' Name of the object
                                        ' this is a problem - object not renamed
                                        statbad &= arrow & s2 & " [" & s3 & "] <br />" & vbCrLf
                                        cnt = cnt + 1
                                    ElseIf Not IsActiveLabel Then  ' Nonactive labels need not be renamed.
                                        statgood &= bullet & s2 & " [" & s3 & "] <br />" & vbCrLf
                                    End If

                                End If
                            End If

                        Catch ex As Exception
                            '     Beep()
                            MessageBox.Show("Error occured in Check Object Naming - " & ex.Message)
                        End Try
                    End If
                Next obj

                If Not foundflag Then statgood = "None Found"

                Select Case tmpObj
                    Case "Button"
                        AppForm.ObjButton.Status = buildObjSummary(strObj, cnt, statgood, statbad, css)
                        AppForm.ObjButton.cssClass = css
                        AppForm.ObjButton.cnt = cnt
                    Case "Label"
                        AppForm.ObjLabel.Status = buildObjSummary(strObj, cnt, statgood, statbad, css)
                        AppForm.ObjLabel.cssClass = css
                        AppForm.ObjLabel.cnt = cnt
                    Case "LabelActive"
                        AppForm.ObjActiveLabel.Status = buildObjSummary("Active Label", cnt, statgood, statbad, css)
                        AppForm.ObjActiveLabel.cssClass = css
                        AppForm.ObjActiveLabel.cnt = cnt
                    Case "LabelNonactive"
                        AppForm.ObjNonactiveLabel.Status = buildObjSummary("Nonactive Label", cnt, statgood, statbad, css)
                        AppForm.ObjNonactiveLabel.cssClass = css
                        AppForm.ObjNonactiveLabel.cnt = cnt
                    Case "Textbox"
                        AppForm.ObjTextbox.Status = buildObjSummary(strObj, cnt, statgood, statbad, css)
                        AppForm.ObjTextbox.cssClass = css
                        AppForm.ObjTextbox.cnt = cnt
                    Case "Listbox"
                        AppForm.ObjListbox.Status = buildObjSummary(strObj, cnt, statgood, statbad, css)
                        AppForm.ObjListbox.cssClass = css
                        AppForm.ObjListbox.cnt = cnt
                    Case "Combobox"
                        AppForm.ObjCombobox.Status = buildObjSummary(strObj, cnt, statgood, statbad, css)
                        AppForm.ObjCombobox.cssClass = css
                        AppForm.ObjCombobox.cnt = cnt
                    Case "RadioButton"
                        AppForm.ObjRadioButton.Status = buildObjSummary(strObj, cnt, statgood, statbad, css)
                        AppForm.ObjRadioButton.cssClass = css
                        AppForm.ObjRadioButton.cnt = cnt
                    Case "CheckBox"
                        AppForm.ObjCheckbox.Status = buildObjSummary(strObj, cnt, statgood, statbad, css)
                        AppForm.ObjCheckbox.cssClass = css
                        AppForm.ObjCheckbox.cnt = cnt
                    Case "GroupBox"
                        AppForm.ObjGroupBox.Status = buildObjSummary(strObj, cnt, statgood, statbad, css)
                        AppForm.ObjGroupBox.cssClass = css
                        AppForm.ObjGroupBox.cnt = cnt
                    Case "OpenFileDialog"
                        AppForm.ObjOpenFileDialog.Status = buildObjSummary(strObj, cnt, statgood, statbad, css)
                        AppForm.ObjOpenFileDialog.cssClass = css
                        AppForm.ObjOpenFileDialog.cnt = cnt
                    Case "SaveFileDialog"
                        AppForm.ObjSaveFileDialog.Status = buildObjSummary(strObj, cnt, statgood, statbad, css)
                        AppForm.ObjSaveFileDialog.cssClass = css
                        AppForm.ObjSaveFileDialog.cnt = cnt
                    Case "WebBrowser"
                        AppForm.ObjWebBrowser.Status = buildObjSummary(strObj, cnt, statgood, statbad, css)
                        AppForm.ObjWebBrowser.cssClass = css
                        AppForm.ObjWebBrowser.cnt = cnt
                    Case "OpenFileDialog"
                        AppForm.ObjOpenFileDialog.Status = buildObjSummary(strObj, cnt, statgood, statbad, css)
                        AppForm.ObjOpenFileDialog.cssClass = css
                        AppForm.ObjOpenFileDialog.cnt = cnt
                    Case "SaveFileDialog"
                        AppForm.ObjSaveFileDialog.Status = buildObjSummary(strObj, cnt, statgood, statbad, css)
                        AppForm.ObjSaveFileDialog.cssClass = css
                        AppForm.ObjSaveFileDialog.cnt = cnt
                End Select

                cnt = 0
                foundflag = False
                statgood = ""
                statbad = ""

            Next strObj

        End If

    End Sub

    Function buildObjSummary(strobj As String, cnt As Integer, statgood As String, statbad As String, ByRef css As String) As String

        ' if the count = -1, color is set to clear, count = 0, the color is set to green. IF count > 0 the color is red
        Dim txt As String

        If cnt = -1 Then
            txt = statgood
            css = ""       ' "itemclear"

        ElseIf cnt = 0 Then
            txt = String.Format("<span class=""boldtext"">{0} objects with proper prefix</span>", strobj) & "<br />" & vbCrLf & statgood
            css = "itemgreen"
        Else
            txt = String.Format("<span class=""boldtext"">{0} {1} without proper prefix</span>", cnt, strobj) & "<br />" & vbCrLf
            txt &= statbad
            css = "itemred"

            If statgood.Trim.Length > 0 Then
                txt &= String.Format("<span class=""boldtext"">{0} objects with proper prefix</span>", strobj) & "<br />" & vbCrLf
                txt &= statgood

            End If
        End If
        Return txt
    End Function
    ' ============================================ End of Checks ===================================================

    Sub BuildReport(key As String, ByRef SAssignment As Assignment, ByRef AppForm As Assignment.AppForm, ByRef SSummary As Assignment.AppSummary, Var1 As String)
        Dim dummyItem As New Assignment.MyItems
        Dim varlist() As String

        Dim errcnt As Integer
        Dim errComment As String = ""

        Dim starttable As String = "<table class=""info"">" & vbCrLf
        Dim th As String = "<tr> <th class=""req"" > Req </th> <th class=""req"" >OK </th>  <th class=""titlecol""> Item </th>   <th class=""statuscol"" > Status </th>  <th class=""ptcol"" > Possible<br /> Pts. </th>   <th class=""ptcol"">  Your<br /> Score</th>  <th class=""commentcol""> Comment </th> </tr>" & vbCrLf
        ' ----------------------------------------------------------------------------------
        Dim h2 As String = "<h2> {0} </h2>"
        Dim h3 As String = "<h3> {0} </h3>"
        ' ----------------------------------------------------------------------------------------------------
        With SSummary
            errcnt = 0

            Select Case key
                Case "Application Level"


                    strStudentReport &= String.Format(h3, "Application Level Information") & vbCrLf & vbCrLf

                    strStudentReport &= starttable
                    strStudentReport &= th
                    strStudentReport = AppendToReport(strStudentReport, 3, "Development Environment", dummyItem, errcnt, errComment, "", SSummary.TotalScore)
                    strStudentReport = AppendToReport(strStudentReport, 2, " - SLN File", SAssignment.hasSLN, errcnt, errComment, "hasSLN", SSummary.TotalScore)
                    strStudentReport = AppendToReport(strStudentReport, 2, " - vbProj File", SAssignment.hasVBproj, errcnt, errComment, "hasvbProj", SSummary.TotalScore)
                    strStudentReport = AppendToReport(strStudentReport, 2, " - VB Version", SAssignment.VBVersion, errcnt, errComment, "hasVBVersion", SSummary.TotalScore)
                    strStudentReport = AppendToReport(strStudentReport, 10, "General Help", SAssignment.VBVersion, errcnt, errComment, "", SSummary.TotalScore)
                    errcnt = 0
                    errComment = ""

                    ReDim varlist(7)
                    varlist = {"hasSplashScreen", "hasAboutBox", "InfoAppTitle", "InfoDescription", "InfoCompany", "InfoProduct", "InfoTrademark", "InfoCopyright"}

                    If SomeVarDisplayed(varlist) Then

                        strStudentReport = AppendToReport(strStudentReport, 3, "Application Info", dummyItem, errcnt, errComment, "", SSummary.TotalScore)

                        strStudentReport = AppendToReport(strStudentReport, 2, " - Splash Screen", SAssignment.hasSplashScreen, errcnt, errComment, "hasSplashScreen", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - About Box", SAssignment.hasAboutBox, errcnt, errComment, "hasAboutBox", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Application Title", .InfoAppTitle, errcnt, errComment, "InfoAppTitle", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Description", .InfoDescription, errcnt, errComment, "InfoDescription", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Company", .InfoCompany, errcnt, errComment, "InfoCompany", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Product", .InfoProduct, errcnt, errComment, "InfoProduct", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Trademark", .InfoTrademark, errcnt, errComment, "InfoTrademark", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Copyright", .InfoCopyright, errcnt, errComment, "InfoCopyright", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 10, " Application Info", dummyItem, errcnt, errComment, "", SSummary.TotalScore)
                        errcnt = 0
                        errComment = ""
                    End If

                    ReDim varlist(1)
                    varlist = {"OptionStrict", "OptionExplicit"}

                    If SomeVarDisplayed(varlist) Then

                        strStudentReport = AppendToReport(strStudentReport, 3, "Compile Options", dummyItem, errcnt, errComment, "", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Option Strict", SAssignment.OptionStrict, errcnt, errComment, "OptionStrict", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Option Explicit", SAssignment.OptionExplicit, errcnt, errComment, "OptionExplicit", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 10, "Options", .InfoCopyright, errcnt, errComment, "", SSummary.TotalScore)

                    End If
                    strStudentReport &= "</table>" & vbCrLf & vbCrLf

                Case "Debugging"
                    ' This has not been implemented. I am not sure if it is possible, and if so how to determine if these features have been set within the student's environment.
                    strStudentReport &= String.Format(h3, "Debugging", errcnt, errComment, "", SSummary.TotalScore)

                    strStudentReport &= starttable
                    strStudentReport &= th
                    strStudentReport = AppendToReport(strStudentReport, 3, "Debugging", dummyItem, errcnt, errComment, "", SSummary.TotalScore)
                    strStudentReport = AppendToReport(strStudentReport, 2, " - BreakPoints", dummyItem, errcnt, errComment, "", SSummary.TotalScore)
                    strStudentReport = AppendToReport(strStudentReport, 2, " - Watch Variables", dummyItem, errcnt, errComment, "", SSummary.TotalScore)
                    strStudentReport = AppendToReport(strStudentReport, 10, "Debugging", dummyItem, errcnt, errComment, "", SSummary.TotalScore)
                    errcnt = 0
                    errComment = ""

                    strStudentReport &= "</table>" & vbCrLf & vbCrLf


                Case "New Project File"
                    '                strStudentReport &= String.Format(h2, "Comments Related to <span class=""boldtext"">" & Var1 & "</span>", errcnt, errComment, "", ssummary.totalscore)

                Case "Form Objects"
                    With AppForm
                        strStudentReport &= String.Format(h3, "Form Objects", errcnt, errComment, "", SSummary.TotalScore)

                        strStudentReport &= starttable
                        strStudentReport &= th

                        ReDim varlist(5)
                        varlist = {"ChangeFormText", "SetFormAcceptButton", "SetFormCancelButton", "ModifyStartPosition", "ChangeFormColor"}

                        If SomeVarDisplayed(varlist) Then

                            strStudentReport = AppendToReport(strStudentReport, 3, "Design Time Form Properties", dummyItem, errcnt, errComment, "", SSummary.TotalScore)

                            strStudentReport = AppendToReport(strStudentReport, 2, " - Form Text Property", .FormText, errcnt, errComment, "ChangeFormText", SSummary.TotalScore)
                            strStudentReport = AppendToReport(strStudentReport, 2, " - Accept Button", .FormAcceptButton, errcnt, errComment, "SetFormAcceptButton", SSummary.TotalScore)
                            strStudentReport = AppendToReport(strStudentReport, 2, " - Cancel Button", .FormCancelButton, errcnt, errComment, "SetFormCancelButton", SSummary.TotalScore)
                            strStudentReport = AppendToReport(strStudentReport, 2, " - Start Position", .FormStartPosition, errcnt, errComment, "ModifyStartPosition", SSummary.TotalScore)
                            strStudentReport = AppendToReport(strStudentReport, 2, " - Non-Gray Form Color", .FormBackColor, errcnt, errComment, "ChangeFormColor", SSummary.TotalScore)
                            '         strStudentReport = AppendToReport(strStudentReport, 2, " - Form Load Method", .FormLoadMethod, errcnt, errComment, "UtilizeFormLoadMethod", ssummary.totalscore)

                            strStudentReport = AppendToReport(strStudentReport, 10, "", dummyItem, errcnt, errComment, "", SSummary.TotalScore)
                            errcnt = 0
                            errComment = ""
                        End If

                        ReDim varlist(12)
                        varlist = {"IncludeFrmInFormName", "ButtonObj", "ObjTextbox", "ActiveLabel", "NonactiveLabel", "ObjCombobox", "ObjListbox", "ObjRadioButton", "ObjCheckbox", "ObjGroupBox", "ObjOpenFileDialog", "ObjSaveFileDialog", "ObjWebBrowser"}

                        If SomeVarDisplayed(varlist) Then
                            strStudentReport = AppendToReport(strStudentReport, 3, "Form Object Names Incorporate Object Prefix", dummyItem, errcnt, errComment, "", SSummary.TotalScore)

                            strStudentReport = AppendToReport(strStudentReport, 2, " - Form (frm)", .FormName, errcnt, errComment, "IncludeFrmInFormName", SSummary.TotalScore)
                            strStudentReport = AppendToReport(strStudentReport, 2, " - Buttons (btn)", .ObjButton, errcnt, errComment, "ButtonObj", SSummary.TotalScore)
                            strStudentReport = AppendToReport(strStudentReport, 2, " - Textboxes (txt)", .ObjTextbox, errcnt, errComment, "TextboxObj", SSummary.TotalScore)
                            '  strStudentReport = AppendToReport(strStudentReport, 2, " - Labels (lbl)", .objLabel, errcnt, errComment, "", ssummary.totalscore)
                            strStudentReport = AppendToReport(strStudentReport, 2, " - Active Labels (lbl)", .ObjActiveLabel, errcnt, errComment, "ActiveLabels", SSummary.TotalScore)
                            strStudentReport = AppendToReport(strStudentReport, 2, " - NonActive Labels (no prefix needed)", .ObjNonactiveLabel, errcnt, errComment, "NonActiveLabels", SSummary.TotalScore)
                            strStudentReport = AppendToReport(strStudentReport, 2, " - Combobox (cbx)", .ObjCombobox, errcnt, errComment, "ComboBoxObj", SSummary.TotalScore)
                            strStudentReport = AppendToReport(strStudentReport, 2, " - Listbox (lbx)", .ObjListbox, errcnt, errComment, "ListBoxObj", SSummary.TotalScore)
                            strStudentReport = AppendToReport(strStudentReport, 2, " - Radiobutton (rbn)", .ObjRadioButton, errcnt, errComment, "RadioButtonObj", SSummary.TotalScore)
                            strStudentReport = AppendToReport(strStudentReport, 2, " - Checkbox (cbx)", .ObjCheckbox, errcnt, errComment, "CheckBoxObj", SSummary.TotalScore)
                            strStudentReport = AppendToReport(strStudentReport, 2, " - Groupbox (gbx)", .ObjGroupBox, errcnt, errComment, "GroupBoxObj", SSummary.TotalScore)
                            strStudentReport = AppendToReport(strStudentReport, 2, " - OpenFileDialog (ofd)", .ObjOpenFileDialog, errcnt, errComment, "ObjOpenFileDialog", SSummary.TotalScore)
                            strStudentReport = AppendToReport(strStudentReport, 2, " - SaveFileDialog (sfd)", .ObjSaveFileDialog, errcnt, errComment, "ObjSaveFileDialog", SSummary.TotalScore)
                            strStudentReport = AppendToReport(strStudentReport, 2, " - WebBrowser (wb)", .ObjWebBrowser, errcnt, errComment, "WebBrowserObj", SSummary.TotalScore)
                            strStudentReport = AppendToReport(strStudentReport, 10, "Form Objects", dummyItem, errcnt, errComment, "", SSummary.TotalScore)
                            errcnt = 0
                            errComment = ""
                        End If

                        strStudentReport &= "</table>" & vbCrLf & vbCrLf
                    End With

                Case "Coding Standards"

                    strStudentReport &= String.Format(h3, "Coding Standards", errcnt, errComment, "", SSummary.TotalScore)

                    strStudentReport &= starttable
                    strStudentReport &= th


                    ReDim varlist(5)
                    varlist = {"CommentSubs", "CommentIF", "CommentFOR", "CommentDO", "CommentWHILE", "CommentSELECT"}

                    If SomeVarDisplayed(varlist) Then
                        strStudentReport = AppendToReport(strStudentReport, 3, "Use of Comments", dummyItem, errcnt, errComment, "", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - First Line of Sub/Function", .CommentSub, errcnt, errComment, "CommentSubs", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Prior to IF", .CommentIF, errcnt, errComment, "CommentIF", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Prior to For", .CommentFor, errcnt, errComment, "CommentFOR", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Prior to Do", .CommentDo, errcnt, errComment, "CommentDO", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Prior to While", .CommentWhile, errcnt, errComment, "CommentWHILE", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Prior to Select Case", .CommentSelect, errcnt, errComment, "CommentSELECT", SSummary.TotalScore)

                        If errComment.Length > 0 Then
                            '  errComment = ""
                            dummyItem.cssClass = "itemred"
                            strStudentReport = AppendToReport(strStudentReport, 10, "General Comments on Comments", dummyItem, errcnt, errComment, "CommentGeneral", SSummary.TotalScore)
                        Else
                            strStudentReport = AppendToReport(strStudentReport, 10, "Use of Comments", dummyItem, errcnt, errComment, "", SSummary.TotalScore)
                        End If

                        errcnt = 0
                        errComment = ""
                    End If


                    ReDim varlist(2)
                    varlist = {"VarArrays", "VarLists", "VarStructures"}

                    If SomeVarDisplayed(varlist) Then

                        '     If .VarArrays.showVar OrElse .VarLists.showVar OrElse .VarStructures.showVar Then
                        strStudentReport = AppendToReport(strStudentReport, 3, "Data Structures", dummyItem, errcnt, errComment, "", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Arrays", .VarArrays, errcnt, errComment, "VarArrays", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Lists", .VarLists, errcnt, errComment, "VarLists", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Structures", .VarStructures, errcnt, errComment, "VarStructures", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 10, "Data Structures", dummyItem, errcnt, errComment, "", SSummary.TotalScore)
                        errcnt = 0
                        errComment = ""
                    End If


                    ReDim varlist(4)
                    varlist = {"VarString", "VarInteger", "VarDecimal", "VarDate", "VarBoolean"}

                    If SomeVarDisplayed(varlist) Then

                        strStudentReport = AppendToReport(strStudentReport, 3, "Variable Data Types - Checking to see which data types are used ", dummyItem, errcnt, errComment, "", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - String", .VarString, errcnt, errComment, "VarString", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Integer", .VarInteger, errcnt, errComment, "VarInteger", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Decimal/Double", .VarDecimal, errcnt, errComment, "VarDecimal", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Date", .VarDate, errcnt, errComment, "VarDate", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Boolean", .VarBoolean, errcnt, errComment, "VarBoolean", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 10, "Variable Data Types", dummyItem, errcnt, errComment, "", SSummary.TotalScore)
                        errcnt = 0
                        errComment = ""
                    End If


                    ReDim varlist(18)
                    varlist = {"LogicElse", "LogicElseIF", "LogicNestedIF", "LogicNestedFOR", "LogicConvertToString", "LogicStringFormat", "LogicStringFormatParameters", "LogicConcatination", "LogicCaseInsensitive", "LogicTryCatch", "LogicComplexConditions", "LogicStreamReader", "LogicStreamReaderClose", "LogicStreamWriter", "LogicStreamWriterClose"}

                    If SomeVarDisplayed(varlist) Then


                        strStudentReport = AppendToReport(strStudentReport, 3, "Program Logic - Checking to see if each Programming Control Stucture is used or not", dummyItem, errcnt, errComment, "", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Else", .LogicElse, errcnt, errComment, "LogicElse", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - ElseIF", .LogicElseIF, errcnt, errComment, "LogicElseIF", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Nested IF", .LogicNestedIF, errcnt, errComment, "LogicNestedIF", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Nested For/Do", .LogicNestedFOR, errcnt, errComment, "LogicNestedFOR", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Convert to String (cStr or .toString)", .LogicConvertToString, errcnt, errComment, "LogicConvertToString", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Format String (.toString() or String.Format())", .LogicStringFormat, errcnt, errComment, "LogicStringFormat", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Template Parameters", .LogicStringFormatParameters, errcnt, errComment, "LogicStringFormatParameters", SSummary.TotalScore)
                        '   strStudentReport = AppendToReport(strStudentReport, 2, " - ByRef Parameters", .LogicByRef, errcnt, errComment, "LogicByRef", ssummary.totalscore)
                        '   strStudentReport = AppendToReport(strStudentReport, 2, " - Optional Parameters", .LogicOptional, errcnt, errComment, "LogicOptional", ssummary.totalscore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Concatenation", .LogicConcatination, errcnt, errComment, "LogicConcatination", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Case Insensitve ", .LogicCaseInsensitive, errcnt, errComment, "LogicCaseInsensitive", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Try ... Catch", .LogicTryCatch, errcnt, errComment, "LogicTryCatch", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Complex Conditions", .LogicComplexConditions, errcnt, errComment, "LogicComplexConditions", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Req Use of a StreamReader", .LogicStreamReader, errcnt, errComment, "LogicStreamReader", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Req matching StreamReader.Close", .LogicStreamReaderClose, errcnt, errComment, "LogicStreamReaderClose", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Req Use of a StreamWriter", .LogicStreamWriter, errcnt, errComment, "LogicStreamWriter", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Req matching StreamWriter.Close", .LogicStreamWriterClose, errcnt, errComment, "LogicStreamWriterClose", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 10, "Program Logic", dummyItem, errcnt, errComment, "", SSummary.TotalScore)
                        errcnt = 0
                        errComment = ""


                    End If


                    ReDim varlist(2)
                    varlist = {"SystemIO", "SystemNet", "SystemDB"}

                    If SomeVarDisplayed(varlist) Then

                        '    If .SystemIO.showVar OrElse .SystemNet.showVar OrElse .SystemDB.showVar Then
                        strStudentReport = AppendToReport(strStudentReport, 3, "Imports", dummyItem, errcnt, errComment, "", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - System.IO", .SystemIO, errcnt, errComment, "SystemIO", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - System.Net", .SystemNet, errcnt, errComment, "SystemNet", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - System.DB", .SystemDB, errcnt, errComment, "SystemDB", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 10, "Imports", dummyItem, errcnt, errComment, "", SSummary.TotalScore)
                        errcnt = 0
                        errComment = ""

                    End If


                    ReDim varlist(5)
                    varlist = {"LogicSub", "LogicOptional", "LogicByRef", "LogicMultipleForms", "LogicModule", "LogicFormLoad"}

                    If SomeVarDisplayed(varlist) Then

                        '    If .SystemIO.showVar OrElse .SystemNet.showVar OrElse .SystemDB.showVar Then
                        strStudentReport = AppendToReport(strStudentReport, 3, "Subs / Functions", dummyItem, errcnt, errComment, "", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Subs", .LogicSub, errcnt, errComment, "LogicSub", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Optional Variables", .LogicOptional, errcnt, errComment, "LogicOptional", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - ByRef Variables", .LogicByRef, errcnt, errComment, "LogicByRef", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Multiple Forms", .LogicMultipleForms, errcnt, errComment, "LogicMultipleForms", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Include Module", .LogicModule, errcnt, errComment, "LogicModule", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 2, " - Form Load Method", .LogicFormLoad, errcnt, errComment, "LogicFormLoad", SSummary.TotalScore)
                        strStudentReport = AppendToReport(strStudentReport, 10, "Subs / Functions", dummyItem, errcnt, errComment, "", SSummary.TotalScore)
                        errcnt = 0
                        errComment = ""

                    End If


                    strStudentReport &= "</table>" & vbCrLf & vbCrLf


                    strStudentReport = strStudentReport.Replace("[TLOC]", TotalLinesOfCode.ToString("n0") & " Lines of Code (not including comments)")

                    '  TotalScore = (SSummary.TotalScore / TotalPossiblePts).ToString("p1")
                    TotalScore = SSummary.TotalScore.ToString("n1") & " deduction out of " & TotalPossiblePts.ToString("n1") & " possible points = " & (SSummary.TotalScore / TotalPossiblePts).ToString("p1")

                    strStudentReport = strStudentReport.Replace("[SCORE]", TotalScore & vbCrLf)


                Case Else
                    If key.StartsWith("Assessment Results for") Then
                        strStudentReport &= "<h2>" & key & "</h2>" & vbCrLf
                    Else
                        MessageBox.Show("BuildReport received an unknown key = " & key)
                    End If
            End Select
        End With
    End Sub

    Sub integrateSSummary(SSummary As Assignment.AppSummary, ByRef IntSSummary As Assignment.AppSummary, isFirst As Boolean)

    End Sub

    Function SomeVarDisplayed(a() As String) As Boolean
        ' This checks an array of strings to see if any of them have settings which are either required or show var. If so, it returns true, else it returns false.
        Dim Flag As Boolean
        Dim Setting As New MySettings

        Flag = False

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



    Function AppendToReport(rpt As String, Template As Integer, Title As String, topic As Assignment.MyItems, ByRef errcnt As Integer, ByRef err As String, Item As String, ByRef total As Decimal) As String
        Dim s As String = ""
        Dim isok As String = ""
        Dim req As String
        Dim errstring As String = ""
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
                'If Setting.Req Then
                '    Beep()
                'End If

                If Setting.Req And Setting.MaxPts <> 0 And .cssClass = "itemred" Then
                    .YourPts = Setting.MaxPts - Math.Min(Setting.PtsPerError * Math.Max(1, .n), Setting.MaxPts)
                    total += .YourPts
                Else
                    .YourPts = 0
                End If


                ' check to see if item is required
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
                        s = String.Format(tr1, "", req, isok, nonChk, Title, .cssClass, .Status, Setting.MaxPts, .YourPts, .Comments)
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


    Sub PopulateNonCheckCSS_Form2(ByRef AppForm As Assignment.AppForm)
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

        With AppForm
            ' -----------------------------------------------------------

            If Find_Setting("RenameObjects", "PopulatenonCheckCSS").Req Then
                .ObjButton.cssNonChk = c
                .ObjLabel.cssNonChk = c
                .ObjActiveLabel.cssNonChk = c
                .ObjNonactiveLabel.cssNonChk = c
                .ObjTextbox.cssNonChk = c
                .ObjListbox.cssNonChk = c
                .ObjCombobox.cssNonChk = c
                .ObjRadioButton.cssNonChk = c
                .ObjCheckbox.cssNonChk = c
                .ObjGroupBox.cssNonChk = c
                .ObjPanel.cssNonChk = c
                .ObjWebBrowser.cssNonChk = c
                .ObjOpenFileDialog.cssNonChk = c
                .ObjSaveFileDialog.cssNonChk = c

            Else
                .ObjButton.cssNonChk = nc
                .ObjLabel.cssNonChk = nc
                .ObjActiveLabel.cssNonChk = nc
                .ObjNonactiveLabel.cssNonChk = nc
                .ObjTextbox.cssNonChk = nc
                .ObjListbox.cssNonChk = nc
                .ObjCombobox.cssNonChk = nc
                .ObjRadioButton.cssNonChk = nc
                .ObjCheckbox.cssNonChk = nc
                .ObjGroupBox.cssNonChk = nc
                .ObjPanel.cssNonChk = nc
                .ObjWebBrowser.cssNonChk = nc
                .ObjOpenFileDialog.cssNonChk = nc
                .ObjSaveFileDialog.cssNonChk = nc

            End If

            If Find_Setting("ObjOpenFileDialog", "PopulatenonCheckCSS").Req Then .ObjOpenFileDialog.cssNonChk = c Else .ObjOpenFileDialog.cssNonChk = nc
            If Find_Setting("ObjSaveFileDialog", "PopulatenonCheckCSS").Req Then .ObjSaveFileDialog.cssNonChk = c Else .ObjSaveFileDialog.cssNonChk = nc

            If Find_Setting("ChangeFormText", "PopulatenonCheckCSS").Req Then .FormText.cssNonChk = c Else .FormText.cssNonChk = nc

            If Find_Setting("ChangeFormColor", "PopulatenonCheckCSS").Req Then .FormBackColor.cssNonChk = c Else .FormBackColor.cssNonChk = nc
            If Find_Setting("SetFormAcceptButton", "PopulatenonCheckCSS").Req Then .FormAcceptButton.cssNonChk = c Else .FormAcceptButton.cssNonChk = nc
            If Find_Setting("SetFormCancelButton", "PopulatenonCheckCSS").Req Then .FormCancelButton.cssNonChk = c Else .FormCancelButton.cssNonChk = nc
            If Find_Setting("ModifyStartPosition", "PopulatenonCheckCSS").Req Then .FormStartPosition.cssNonChk = c Else .FormStartPosition.cssNonChk = nc
            If Find_Setting("LogicFormLoad", "PopulatenonCheckCSS").Req Then .FormLoadMethod.cssNonChk = c Else .FormLoadMethod.cssNonChk = nc

            ' ----------------------------------------------------------

        End With

    End Sub



    Sub PopulateNonCheckCSS_Summary(ByRef AppSum As Assignment.AppSummary, ByRef Assign As Assignment)
        Dim nc As String = ""                       ' Non-checked property
        Dim c As String = "ncWhite"    ' Checked property
        Dim item As New MySettings

        ' ----------------------------------------------------------

        If HideGray = "Gray" Then
            nc = "ncGray"
        ElseIf HideGray = "Hide" Then
            nc = "ncHide"
        Else
            nc = "ncWhite"
        End If

        With AppSum
            ' ----------------------------------------------------------
            setchecked(Find_Setting("InfoAppTitle", "PopulatenonCheckCSS_summary").Req, .InfoAppTitle, nc, c)
            setchecked(Find_Setting("InfoDescription", "PopulatenonCheckCSS_summary").Req, .InfoDescription, nc, c)
            setchecked(Find_Setting("InfoCompany", "PopulatenonCheckCSS_summary").Req, .InfoCompany, nc, c)
            setchecked(Find_Setting("InfoProduct", "PopulatenonCheckCSS_summary").Req, .InfoProduct, nc, c)
            setchecked(Find_Setting("InfoTrademark", "PopulatenonCheckCSS_summary").Req, .InfoTrademark, nc, c)
            setchecked(Find_Setting("InfoCopyright", "PopulatenonCheckCSS_summary").Req, .InfoCopyright, nc, c)
            setchecked(Find_Setting("", "PopulatenonCheckCSS_summary").Req, .InfoGUID, nc, c)


            setchecked(Find_Setting("CommentSubs", "PopulatenonCheckCSS_summary").Req, .CommentSub, nc, c)
            setchecked(Find_Setting("CommentIF", "PopulatenonCheckCSS_summary").Req, .CommentIF, nc, c)
            setchecked(Find_Setting("CommentFOR", "PopulatenonCheckCSS_summary").Req, .CommentFor, nc, c)
            setchecked(Find_Setting("CommentDO", "PopulatenonCheckCSS_summary").Req, .CommentDo, nc, c)
            setchecked(Find_Setting("CommentWHILE", "PopulatenonCheckCSS_summary").Req, .CommentWhile, nc, c)
            setchecked(Find_Setting("CommentSELECT", "PopulatenonCheckCSS_summary").Req, .CommentSelect, nc, c)

            setchecked(Find_Setting("VarString", "PopulatenonCheckCSS_summary").Req, .VarString, nc, c)
            setchecked(Find_Setting("VarBoolean", "PopulatenonCheckCSS_summary").Req, .VarBoolean, nc, c)
            setchecked(Find_Setting("VarInteger", "PopulatenonCheckCSS_summary").Req, .VarInteger, nc, c)
            setchecked(Find_Setting("VarDecimal", "PopulatenonCheckCSS_summary").Req, .VarDecimal, nc, c)
            setchecked(Find_Setting("VarDate", "PopulatenonCheckCSS_summary").Req, .VarDate, nc, c)

            setchecked(Find_Setting("VarArray", "PopulatenonCheckCSS_summary").Req, .VarArrays, nc, c)
            setchecked(Find_Setting("VarList", "PopulatenonCheckCSS_summary").Req, .VarLists, nc, c)
            setchecked(Find_Setting("VarStructure", "PopulatenonCheckCSS_summary").Req, .VarStructures, nc, c)

            setchecked(Find_Setting("VariablePrefixes", "PopulatenonCheckCSS_summary").Req, .varPrefixes, nc, c)

            ' setchecked(Find_Setting("", "PopulatenonCheckCSS_summary").req, .LogicFlowControl, nc, c)
            setchecked(Find_Setting("LogicIF", "PopulatenonCheckCSS_summary").Req, .LogicIF, nc, c)
            setchecked(Find_Setting("LogicFOR", "PopulatenonCheckCSS_summary").Req, .LogicFor, nc, c)
            setchecked(Find_Setting("LogicDO", "PopulatenonCheckCSS_summary").Req, .LogicDo, nc, c)
            setchecked(Find_Setting("LogicWHILE", "PopulatenonCheckCSS_summary").Req, .LogicWhile, nc, c)
            setchecked(Find_Setting("LogicSelectCase", "PopulatenonCheckCSS_summary").Req, .LogicSelectCase, nc, c)
            setchecked(Find_Setting("LogicElse", "PopulatenonCheckCSS_summary").Req, .LogicElse, nc, c)
            setchecked(Find_Setting("LogicElseIF", "PopulatenonCheckCSS_summary").Req, .LogicElseIF, nc, c)
            setchecked(Find_Setting("LogicTryCatch", "PopulatenonCheckCSS_summary").Req, .LogicTryCatch, nc, c)
            setchecked(Find_Setting("LogicStreamReader", "PopulatenonCheckCSS_summary").Req, .LogicStreamReader, nc, c)
            setchecked(Find_Setting("LogicStreamWriter", "PopulatenonCheckCSS_summary").Req, .LogicStreamWriter, nc, c)
            setchecked(Find_Setting("LogicStreamReaderClose", "PopulatenonCheckCSS_summary").Req, .LogicStreamReaderClose, nc, c)
            setchecked(Find_Setting("LogicStreamWriterClose", "PopulatenonCheckCSS_summary").Req, .LogicStreamWriterClose, nc, c)
            setchecked(Find_Setting("LogicSub", "PopulatenonCheckCSS_summary").Req, .LogicSub, nc, c)
            '  setchecked(Find_Setting("", "PopulatenonCheckCSS_summary").req, .LogicFunction, nc, c)
            setchecked(Find_Setting("LogicOptional", "PopulatenonCheckCSS_summary").Req, .LogicOptional, nc, c)
            setchecked(Find_Setting("LogicByRef", "PopulatenonCheckCSS_summary").Req, .LogicByRef, nc, c)
            setchecked(Find_Setting("LogicConvertToString", "PopulatenonCheckCSS_summary").Req, .LogicCStr, nc, c)
            '   setchecked(Find_Setting("", "PopulatenonCheckCSS_summary").req, .LogicToString, nc, c)
            setchecked(Find_Setting("LogicStringFormat", "PopulatenonCheckCSS_summary").Req, .LogicStringFormat, nc, c)



            '     setchecked(Find_Setting("Req Variable Prefixes", "PopulatenonCheckCSS_summary").req, .LogicVarPrefixes, nc, c)
            setchecked(Find_Setting("LogicNestedIF", "PopulatenonCheckCSS_summary").Req, .LogicNestedIF, nc, c)
            setchecked(Find_Setting("LogicNestedFOR", "PopulatenonCheckCSS_summary").Req, .LogicNestedFOR, nc, c)

            '        setchecked(find_Setting("", "PopulatenonCheckCSS_summary").req, .LogicStringFormatting, nc, c)
            setchecked(Find_Setting("LogicComplexConditions", "PopulatenonCheckCSS_summary").Req, .LogicComplexConditions, nc, c)
            setchecked(Find_Setting("LogicCaseInsensitive", "PopulatenonCheckCSS_summary").Req, .LogicCaseInsensitive, nc, c)
            setchecked(Find_Setting("LogicStringFormatParameters", "PopulatenonCheckCSS_summary").Req, .LogicStringFormat, nc, c)
            setchecked(Find_Setting("LogicConcatination", "PopulatenonCheckCSS_summary").Req, .LogicConcatination, nc, c)

            '     setchecked(Find_Setting("UtilizeLogicFormLoad", "PopulatenonCheckCSS_summary").Req, .FormLoadMethod, nc, c)

            setchecked(Find_Setting("SystemIO", "PopulatenonCheckCSS_summary").Req, .SystemIO, nc, c)
            setchecked(Find_Setting("SystemNet", "PopulatenonCheckCSS_summary").Req, .SystemNet, nc, c)
            setchecked(Find_Setting("SystemDB", "PopulatenonCheckCSS_summary").Req, .SystemDB, nc, c)


            setchecked(Find_Setting("OptionStrict", "PopulatenonCheckCSS_summary").Req, Assign.OptionStrict, nc, c)
            setchecked(Find_Setting("OptionExplicit", "PopulatenonCheckCSS_summary").Req, Assign.OptionExplicit, nc, c)
            ' ----------------------------------------------------------
            setchecked(Find_Setting("hasSLN", "PopulatenonCheckCSS_summary").Req, Assign.hasSLN, nc, c)
            setchecked(Find_Setting("hasvbProj", "PopulatenonCheckCSS_summary").Req, Assign.hasVBproj, nc, c)
            setchecked(Find_Setting("hasSplashScreen", "PopulatenonCheckCSS_summary").Req, Assign.hasSplashScreen, nc, c)
            setchecked(Find_Setting("hasAboutBox", "PopulatenonCheckCSS_summary").Req, Assign.hasAboutBox, nc, c)
            setchecked(Find_Setting("LogicModule", "PopulatenonCheckCSS_summary").Req, Assign.Modules, nc, c)
            ' ----------------------------------------------------------
         End With
        ' ----------------------------------------------------------
    End Sub

    Sub setchecked(chk As Boolean, ByRef obj As Assignment.MyItems, nc As String, c As String)
        If Not chk Then
            obj.cssNonChk = nc
            obj.req = False
        Else
            obj.cssNonChk = c
            obj.req = True
        End If
    End Sub

End Module
