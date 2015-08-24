Imports System.IO
Imports System.IO.Compression
' Imports Ionic.Zip
Imports SharpCompress.Reader
Imports System.Security
Imports System.Security.Cryptography

' Uses http://sharpcompress.codeplex.com/documentation for unzipping


Public Class frmMain

    Dim GenerateGradesheets As Boolean
    Dim hasWordTemplateBeenSpecified As Boolean = False
    Dim strPrefix As String = ""
    '    Dim files_Compressed(200) As String

    ' create variables to hold data
    Dim StudentAssignment As AssignmentInfo
    Dim IntegratedStudentAssignment(NSummary) As MyItems
    Dim IntegratedForm(nForm) As MyItems



    Structure crcdatum
        Dim UserID As String
        Dim Filename As String
        Dim vbMD5 As String
    End Structure

    '  Dim VBProjects(200) As MyVBProjects
    '   Dim nVBProjects As Integer
    '  Dim crcData() As crcdatum

    Private Submissions As New List(Of Submission)
    Private GUIDs As New List(Of GUIDData)


    '   Dim crcN As Integer = -1

    '  Dim WithEvents zip As New ZipUtility
    ' http://www.vbforums.com/showthread.php?t=413705&highlight=ziplib
    Public Structure myFormClass
        Dim FormName As String
        Dim Ncomments As Integer
    End Structure

    'Public Structure AssignmentInfo
    '    Dim StudentID As String
    '    Dim AssignRoot As String
    '    Dim AssignPath As String
    '    Dim vbVersion As String
    '    Dim hasSLNFile As Boolean
    '    Dim hasVBProjFile As Boolean
    '    Dim FormClass() As myFormClass
    'End Structure

    Dim AInfo As AssignmentInfo
    Dim AssignmentName As String = ""


    ' ================================================================================================
    ' ================================================================================================
    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ' #############################################################################################################
        IsFacultyVersion = True          ' if false, then it is the Student Version. It True, it is the Faculty Version
        ' #############################################################################################################

        '        AppDataDir = RemoveLastField(Application.ExecutablePath, "\")
        AppDataDir = Application.StartupPath

        ' the following was added to account for the difference in IDE and Published handling of data files
        'If Not AppDataDir.ToLower.Contains("\bin\debug") Then
        '    AppDataDir &= "\bin\Debug"
        'End If

        btnViewPlagiarism.Visible = False
        btnOutput.Visible = False
        btnAssignSummary.Visible = False
        btnDetail.Visible = False

        rbnCheckGray.Checked = True

        If File.Exists(Application.StartupPath & "\demodir.txt") Then
            Dim sr As StreamReader = File.OpenText(Application.StartupPath & "\demodir.txt")
            lblDemoDir.Text = sr.ReadLine
            sr.Close()
        End If


        ' ----------------------------------------------------------------------
        File.Delete(Application.StartupPath & "\CantFind.txt")

        Settings.LoadCfgFile(Application.StartupPath & "\templates\defaultConfig.cfg") ' this only seems to read the settings and same them to a settings.txt file.

        If File.Exists(Application.StartupPath & "\MissingSetting.txt") Then File.Delete(Application.StartupPath & "\MissingSetting.txt")

        LoadConfig()


        lblNZips.Text = "-"
        '       Me.BackgroundWorker1.RunWorkerAsync()

        If IsFacultyVersion Then
            ' use the design time defaults
        Else

            Me.BackColor = Color.LightCoral
            GroupBox1.Visible = False
            GroupBox1.BackColor = Color.LightCoral
            btnViewPlagiarism.Visible = False
            GroupBox3.Visible = False
            cbxLoadWordTemplate.Visible = False
            cbxJustUnzip.Visible = False
            cbxNoDemoFiles.Checked = True


            ConfigureAppToolStripMenuItem.Visible = False

            btnSelectFile2.Visible = True
            btnSelectFile2.Location = New Point(21, 40)

            Panel2.Location = New Point(21, 23)
            GroupBox1.Height = 170

            GroupBox2.Location = New Point(9, 200)
            GroupBox2.Height = 140
            lblConfigFile.BackColor = Color.White
            ' lblConfigFile.ForeColor = Color.LightPink
            GroupBox4.Location = New Point(9, 80)

            Panel3.Visible = False
            btnProcessApps.Location = New Point(524, 355)

            Label3.Visible = False
            lblSelectedTemplate.Visible = False
            btnBrowseTemplate.Visible = False

            Me.Height = 510
        End If

    End Sub
    ' ================================================================================================
    ' ================================================================================================
    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
        ' Set the Help text description for the FolderBrowserDialog.
        Me.FolderBrowserDialog1.Description = "Select output root directory. A new directory will be created inside the Root to hold the results."
        FolderBrowserDialog1.ShowNewFolderButton = False  ' Do not allow the user to create New files via the FolderBrowserDialog.

        FolderBrowserDialog1.ShowDialog()
        lblDir.Text = FolderBrowserDialog1.SelectedPath.ToString

        AssignmentName = ReturnLastField(lblDir.Text, "\")

        btnBrowseTemplate.Enabled = True
        btnProcessApps.Enabled = True

    End Sub


    Private Sub btnBrowseTemplate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowseTemplate.Click
        ' Let the user select a gradesheet template. It will be copied and placed in the student folder.
        ' Default location is the directory holding the compress homework file
        OpenFileDialog1.InitialDirectory = lblDir.Text

        OpenFileDialog1.Filter = "Word files |*.docx|All files |*.*"
        OpenFileDialog1.FilterIndex = 0
        OpenFileDialog1.RestoreDirectory = True

        OpenFileDialog1.ShowDialog()
        lblSelectedTemplate.Text = OpenFileDialog1.FileName.ToString
        hasWordTemplateBeenSpecified = True
        cbxLoadWordTemplate.Checked = True
    End Sub


    Private Sub btnProcessApps_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProcessApps.Click, StartProcessingToolStripMenuItem.Click
        timeStart = Now
        '  Dim path As String = lblDemoDir.Text
        '        Dim hasvbfile As Boolean

        'rbnShowAll.Enabled = False
        'rbnShowOnlyReq.Enabled = False
        'rbnCheckGray.Enabled = False

        'Dim worker As System.ComponentModel.BackgroundWorker = DirectCast(sender, System.ComponentModel.BackgroundWorker)
        'Dim x As Integer

        ' Ensure that the user entered an assignment name.
        If txtAssignmentName.Text.Trim.Length = 0 Then
            MessageBox.Show("The Assignment name is requried. Please insert appropriate text.", "Missing Assignment Name", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtAssignmentName.Focus()
            Exit Sub
        End If

        ' Ensure that the user entered an assignment name.
        If lblConfigFile.Text.Trim.Length = 0 Then
            MessageBox.Show("You need to select a Config file. It is recomended to use one that is specific to the assignment. If one is not available, use the default configuration.", "Configuration File Required", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        Else
            If rbnAppCFG.Checked Then
                LoadCfgFile(lblConfigFile.Text)
                frmConfig.LoadConfigFile(lblConfigFile.Text)
            Else
                LoadCfgFile(Application.StartupPath & "\templates\FactoryConfig.cfg")
                frmConfig.LoadConfigFile(Application.StartupPath & "\templates\defaultConfig.cfg")
            End If
        End If

        lblMessage.Text = ""

        '        worker.ReportProgress(1, "Preliminaries")


        Submissions.Clear()

        ' Load Instructor demo files
        'worker.ReportProgress(2, "Preliminaries")
        'x = 2


        GuidIssues = False
        MD5Issues = False

        If rbnBlackboardZip.Checked Then
            frmProgressBar.lblExtractStudentFiles.Visible = True
            frmProgressBar.ProgressBar1.Visible = True
        Else
            frmProgressBar.lblExtractStudentFiles.Visible = False
            frmProgressBar.ProgressBar1.Visible = False
        End If

        frmProgressBar.Show()
        btnProcessApps.Enabled = False

        Me.BackgroundWorker1.RunWorkerAsync()

    End Sub

    ' ================================================================================================
    ' DO WORK
    ' ================================================================================================

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        'This method is executed in a worker thread.
        Dim worker As System.ComponentModel.BackgroundWorker = DirectCast(sender, System.ComponentModel.BackgroundWorker)
        Dim path As String = lblDemoDir.Text
        Dim fileext As String = ""
        Dim ii As Integer
        Dim strDir() As String = Nothing

        Dim HasVBProjFile As Boolean = False
        Dim HasSLNFile As Boolean = False
        Dim hasVBFile As Boolean
        Dim strFacReportDetail As String = ""
        Dim nfiles As Integer

        '   Dim vbVersion As String = ""

        Dim fn As String

        Dim ans As Integer
        Dim sw As StreamWriter
        Dim nstudents As Integer

        Dim sr1 As StreamReader
        Dim s As String = ""
        Dim s2 As String = ""

        ' This assumes that the selected archive contains a set of compressed files with student work.\' -------------------------------------------------------------------------------------------------
        Dim n As Integer
        Dim x As Integer
        Dim filelist() As String

        '  lblMessage.Text = ""

        timeLoadInstructorFiles = Now

        ' load in CRC of the instructor demo files
        x = 0
        Try
            If Not cbxNoDemoFiles.Checked Then
                filelist = IO.Directory.GetFiles(path, "*.vb", SearchOption.AllDirectories)
                ' ----------------------------------------------------
                hasVBFile = filelist.Length > 0
                n = filelist.Length

                For Each filename In filelist
                    x += 1
                    worker.ReportProgress(CInt(x * 100 / n), "Preliminaries")

                    If filename.ToLower.EndsWith(".vb") Then
                        Dim hw As New Submission

                        With hw
                            .UserID = "Instructor Demo"
                            '                            .vbCRC = GetCRC32(filename)
                            '  .vbCRC = GetCRC32(filename) ' actually this is the MD5 hash

                            .vbCRC = md5_hash(filename)    ' this is executed
                            .Filename = filename
                        End With

                        Submissions.AddRange({hw})
                    End If
                Next filename
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error when setting up Preliminaries")
        End Try
        ninstructorfiles = n

        n = 0
        path = ""

        ' -------------------------------------------------------------------------------------------------
        ' First check to see if they have specified a template for the gradesheets, if the user specified
        ' that they want grade sheets. It it is not specified, warn the user and exit processing.
        ' -------------------------------------------------------------------------------------------------
        GenerateGradesheets = cbxLoadWordTemplate.Checked
        If cbxLoadWordTemplate.Checked Then
            If Not hasWordTemplateBeenSpecified Then
                MessageBox.Show("ProcessApps - " & "You have indicated that you want the application to generate Grade Sheets, but have not specified a Template. Please select a template or remove the checkmark on Generate Grade Sheets.", "Missing Grade Sheet Template", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If
        End If

        ' -------------------------------------------------------------------------------------------------
        ' Extract student work and place in their own folders. The folder names are shortened to avoid
        ' exceeding the maximum path length of Windows. 
        '
        ' Note, this will not overwrite existing files.
        ' -------------------------------------------------------------------------------------------------

        ' Should likely pick a specific blackboard zip file, not just a folder

        timeUnzipStart = Now

        'Dim di As New IO.DirectoryInfo(FolderBrowserDialog1.SelectedPath)
        'Dim aryFi As IO.FileInfo() = di.GetFiles("*.zip")
        If rbnBlackboardZip.Checked Then
            fn = lblTargetFile.Text


            ' before extracting the file, Check if the folder exists. If so Check if the user wants to allow overwriting of files
            strOutputPath = fn.Substring(0, fn.Length - 4)


            '        If cfgAssignmentTitle = nothing Then cfgAssignmentTitle = ReturnLastField(strOutputPath, "\")
            cfgAssignmentTitle = txtAssignmentName.Text


            If Directory.Exists(strOutputPath) Then
                ans = MessageBox.Show("The output path already exists. Do you want to allow overwriting of folders? ", "Output Folder Exists", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)

                If ans = vbYes Then
                    AllowOverwrite = True
                Else
                    AllowOverwrite = False
                End If

            End If

            If ans = vbCancel Then
                ' don't do anything
                lblMessage.Text = "Process cancelled."
            Else


                ArchiveExtract(fn)
                n = ShortenFilenamesInFolder(strOutputPath, BackgroundWorker1)
                worker.ReportProgress(-4, "The number of projects found = " & n.ToString)
            End If

        ElseIf rbnSingleProject.Checked Then
            strOutputPath = FolderBrowserDialog1.SelectedPath
            worker.ReportProgress(-4, "Processing a single Project.")
        End If

        ' -------------------------------------------------------------------------------------------------
        ' Now we are ready to process the student work. We want to do this in a Background process
        ' so we can show progress on screen.
        ' -------------------------------------------------------------------------------------------------


        ' --------------------------------------------------------------------------------------------
        ' capture the file extension of the gradesheet file so we can use it later
        If GenerateGradesheets Then
            fileext = IO.Path.GetExtension(OpenFileDialog1.FileName)
        End If

        timeProcess = Now

        ' ================================================================================================
        ' ========================================== Main Processing of Files ====================================================
        ' find all the student file directories in target folder. These were created when we uncompressed the Homework file
        ' ================================================================================================
        If rbnBlackboardZip.Checked Then
            strDir = Directory.GetDirectories(strOutputPath)
        ElseIf rbnSingleProject.Checked Then
            ReDim strDir(0)
            strDir(0) = FolderBrowserDialog1.SelectedPath
        End If

        ' If Not cbxJustUnzip.Checked Then InitializeFacultyReport()


        ' ===================================================================================
        ' Determine the total possible pts
        ' ===================================================================================
        TotalPossiblePts = 0
        ii = 0
        For Each p As MySettings In Settings.Settings
            If p.Req Then TotalPossiblePts += p.MaxPts
            If p.Req Then ii += 1
        Next


        InitializeStudentReport()
        ' ===================================================================================
        nstudents = strDir.GetUpperBound(0)

        strFacReport = InitializeFacultyReport()
        nstudentfiles = nstudents + 1

        For ii = 0 To nstudents   ' This cycles through all student folders
            ' ===================================================================================
            TotalLinesOfCode = 0


            Dim StudentAppSummary(NSummary) As MyItems
            Dim StudentAppForm(nForm) As MyItems

            '  PopulateNonCheckCSS_Summary(AppSummary)  ' this sets css to control what happens with Properties that are not being checked in the assignment (hide/show/gray/white)

            If Not cbxJustUnzip.Checked Then
                ' Determine the Student ID

                path = strDir(ii)  '  & "\"  & studentID
                StudentAssignment.StudentID = ReturnLastField(strDir(ii), "\")
                StudentAssignment.AssignRoot = RemoveLastField(strDir(ii), "\")
                StudentAssignment.AssignPath = strDir(ii)

                strStudentID = StudentAssignment.StudentID
                strStudentReport = InitializeStudentReport() & strStudentReport

                ' load file into list so we can check CRC
                Try
                    hasVBFile = IO.Directory.GetFiles(path, "*.vb", SearchOption.AllDirectories).GetLength(0) > 0
                    ' ----------------------------------------------------
                    If hasVBFile Then
                        For Each filename In IO.Directory.GetFiles(path, "*.vb", SearchOption.AllDirectories)
                            If filename.ToLower.EndsWith(".vb") Then
                                Dim hw As New Submission

                                With hw
                                    .UserID = StudentAssignment.StudentID
                                    '  .vbCRC = GetCRC32(filename)
                                    '  .vbCRC = GetCRC32(filename)     
                                    .vbCRC = md5_hash(filename)        ' actually this is the MD5 hash
                                    .Filename = filename               ' this is also executed, but later
                                End With

                                Submissions.AddRange({hw})
                            End If
                        Next filename
                    End If
                Catch ex As Exception
                    Beep()
                End Try

                ' ---------------------------------------------------------------------------------------------------

                If Not rbnBlackboardZip.Checked And rbnSingleProject.Checked Then
                    AInfo.StudentID = AssignmentName
                End If


                If Not Directory.Exists(path & "\Reports") Then
                    Directory.CreateDirectory(path & "\Reports")
                End If
                StudentReportPath = path & "\Reports\" & StudentAssignment.StudentID & "_GradeReport.html"

                sw = File.CreateText(StudentReportPath)   ' Creates a blank file
                sw.Close()

                '      End If

                ' --------------------------------------------------------------------------

                '' If Not cbxJustUnzip.Checked Then InitializeStudentReport()
                strStudentID = StudentAssignment.StudentID
                'InitializeStudentReport()


                ' check for sln and/or vbProj file. One of these is needed to run app, so it is a bad submission if neither is present.
                CheckSLNvbProj(StudentAssignment)
                HasSLNFile = CBool(StudentAssignment.hasSLN.Status)
                HasVBProjFile = CBool(StudentAssignment.hasVBproj.Status)

                ' display a message if there is no SLN or vbProj file.
                If Not HasSLNFile And Not HasVBProjFile Then
                    worker.ReportProgress(-3, StudentAssignment.StudentID & " has no SLN or vbProj file. The submission cannot be assessed.")
                Else
                    ' process the remaining part of the file, since the work can be loaded and assessed.
                    ' Proecessing the submission entails executing all the specified Logic and design testing.

                    ' ############################################################################################
                    '        ProcessVBSubmission(strDir(i), i, studentID, fileext, HasSLNFile Or HasVBProjFile, AppSummary, BackgroundWorker1)

                    '     Dim worker As System.ComponentModel.BackgroundWorker = DirectCast(sender, System.ComponentModel.BackgroundWorker)
                    Dim filesource As String
                    '  Dim nComment As Integer = 0
                    '   Dim i As Integer = 0
                    '   Dim hasFiles As Boolean = False
                    '   Dim OptionStrict As Boolean = False
                    '   Dim OptionStrictOff As Boolean = False
                    '   Dim hasOptionExplicit As Boolean = False
                    '   Dim hasOptionExplicitOff As Boolean = False


                    Dim Filesinbuild As New List(Of String)

                    Dim RemoveFromList As Boolean = True
                    Dim AppDir As String = ""
                    '    Dim var1 As String = ""
                    ' Dim x As Integer
                    Dim Appform(0) As MyItems
                    Dim NAppForm As Integer = 0
                    Dim hasForm As Boolean
                    '     Dim gradesheetext As String = "txt"
                    Dim first As Boolean

                    '   Dim sw As StreamWriter  ' This is used to write out the student report file
                    Dim sr As StreamReader
                    ' ======================================================================================================================

                    ' Create a reports Folder
                    If (Not Directory.Exists(path & "\Reports")) Then
                        Directory.CreateDirectory(path & "\Reports")
                    End If

                    ' ---------------------Copy grading template files (Word File) ----------------------------
                    ' This will only generate a grade sheet in the _a submission for the student when they submit multiple zip files.
                    ' it copies the selected tempate an renames it StudentID_Feedback.doc or docx     
                    If GenerateGradesheets And StudentAssignment.StudentID.EndsWith("_a") Then
                        If lblSelectedTemplate.Text.Contains(".doc") Then  ' Accepts both doc and docx files
                            FileCopy(lblDir.Text & "\" & lblSelectedTemplate.Text, path & "\Reports\" & StudentAssignment.StudentID & "_Feedback." & ReturnLastField(lblSelectedTemplate.Text, "."))
                        End If
                    End If
                    ' ---------------------------------------------------------------------------------------------------

                    ' ---------------------------------------------------------------------------------------------------
                    ' Build summary of student work
                    ' ---------------------------------------------------------------------------------------------------
                    ' Dim SSummary As MyAppSummary
                    'Student.StudentID = StudentAssignment.StudentID
                    'AppSummary.AssignPath = path

                    ' Determine which files are included in the student submission. Then process each one of these
                    ' ---------------------------------------------------------------------------------------------------
                    strStudentRoot = path

                    GetAllFilesInBuild(path, Filesinbuild)


                    ' =================================== Process all Files ===========================================
                    x = 0
                    ' count up the forms in the build
                    For i As Integer = 1 To Filesinbuild.Count
                        If File.Exists(Filesinbuild(i - 1).Replace(".vb", ".designer.vb")) Then
                            x = x + 1
                        End If
                    Next
                    ' ----------------------------------------------------------------------------------------------------------

                    '        worker.ReportProgress(1, "Individual Work")

                    DisplayFilesInReports(Filesinbuild)
                    '        worker.ReportProgress(2, "Individual Work")

                    ' ---------------------------------------------------------------------------------------------------
                    ' The splash screen and About Box files do not have any User coding. So check if they exist and remove from FilesInBuild
                    ' --------------------------------------- check for splash page -----------------------------------
                    CheckForSplashScreen2(Filesinbuild, StudentAssignment.hasSplashScreen, RemoveFromList)


                    With StudentAssignment.hasSplashScreen
                        .cssClass = ""     ' "itemclear"
                        If .Status = vbTrue.ToString Then
                            .Status = "Has Splashscreen"
                            If Find_Setting("hasSplashScreen", "DoWork").Req Then .cssClass = "itemgreen"
                        Else
                            .Status = "No Splashscreen"
                            If Find_Setting("hasSplashScreen", "DoWork").Req Then .cssClass = "itemred"
                        End If
                    End With

                    pbar3max = 10 + 3 * Filesinbuild.Count

                    worker.ReportProgress(3, "Individual Work")

                    ' -------------------------------------- check for about page -------------------------------------
                    CheckForAboutBox2(Filesinbuild, StudentAssignment.hasAboutBox, RemoveFromList)

                    With StudentAssignment.hasAboutBox
                        .cssClass = ""     ' "itemclear"
                        If .Status = vbTrue.ToString Then
                            .Status = "Has AboutBox"
                            If Find_Setting("hasAboutBox", "DoWork").Req Then .cssClass = "itemgreen"
                        Else
                            .Status = "No AboutBox"
                            If Find_Setting("hasAboutBox", "DoWork").Req Then .cssClass = "itemred"
                        End If
                    End With

                    worker.ReportProgress(4, "Individual Work")

                    ' -------------------------------------------------------------------------------------------------


                    ' ---------------------------------------------------------------------------------
                    ' the following checks apply across the whole application
                    ' ---------------------------------------------------------------------------------

                    ' ----------------------------------- check ApplicationInfo ----------------------------------------
                    CheckAPPInfo2(strStudentPath, StudentAppSummary)
                    worker.ReportProgress(5, "Individual Work")

                    ' ----------------------------------- check for Option Strict --------------------------------------
                    If Find_Setting("OptionStrict", "DoWork").Req Then CheckForOptions2(path, Filesinbuild, StudentAssignment.OptionStrict, "Strict")
                    worker.ReportProgress(6, "Individual Work")

                    ' ------------------------------------ check for Option Explicit -----------------------------------
                    If Find_Setting("OptionExplicit", "DoWork").Req Then CheckForOptions2(path, Filesinbuild, StudentAssignment.OptionExplicit, "Explicit")
                    worker.ReportProgress(7, "Individual Work")

                    ' ------------------------- check if the frm prefix is used with all forms -------------------------
                    worker.ReportProgress(8, "Individual Work")

                    ' ---------------------------------- check if the project has a module -----------------------------
                    If Find_Setting("LogicModule", "DoWork").Req Then CheckForModules2(Filesinbuild, StudentAssignment.Modules)
                    worker.ReportProgress(9, "Individual Work")

                    ' -------------------------------------------------------------------------------------------------
                    BuildReport("Application Level", StudentAssignment, StudentAppForm, StudentAppSummary, strStudentReport)
                    '    BuildReport("Debugging", Appform(0), ssummary, "")       ' Don't have logic to cature breakpoints, line numbers, etc.

                    worker.ReportProgress(10, "Individual Work")
                    x = 10


                    ' ===================================================================================================
                    ' now process each file in the build
                    ' =================================================================================================== 
                    first = True
                    For Each filename As String In Filesinbuild
                        AssScore = 0
                        AssPossible = 0
                        nfiles = Filesinbuild.Count

                        StudentAssignment.TotalScore = 0
                        FileLinesOfCode = 0

                        ' Clear all the data so there is no double counting
                        ClearAppArray(StudentAppSummary)
                        ClearAppArray(StudentAppForm)

                        ' display the filename being processed if working with a single app
                        If rbnSingleProject.Checked Then
                            worker.ReportProgress(-4, "Processing " & ReturnLastField(filename, "\"))
                        End If

                        ' creates a new AppForm structure for each new Form File. Avoids Classes and Modules
                        hasForm = False
                        If File.Exists(filename.Replace(".vb", ".designer.vb")) Then
                            If StudentAppForm(EnForm.FormName).Status <> Nothing Then ClearAppArray(StudentAppForm) ' don't clear it if it is already cleared.
                            hasForm = True

                            NAppForm += 1
                            ReDim Preserve Appform(NAppForm)
                        End If

                        ' Populate the structure with settings indicating how to handle non checked items.

                        PopulateNonCheckCSS_Form2(StudentAppForm)


                        BuildReport("New Project File", StudentAssignment, StudentAppForm, StudentAppSummary, strStudentReport) ' the AppForm(0) is just a placeholder

                        AppDir = IO.Path.GetDirectoryName(filename)

                        ' read in source file
                        sr = New StreamReader(filename)

                        filesource = sr.ReadToEnd
                        sr.Close()

                        ' ---- check for form  -------
                        BuildReport("Assessment Results for " & ReturnLastField(filename, "\"), StudentAssignment, StudentAppForm, StudentAppSummary, strStudentReport)

                        ' Assess the form properties & objects
                        If hasForm Then
                            If rbnSingleProject.Checked Then worker.ReportProgress(-4, "Processing " & ReturnLastField(filename, "\") & " - Form Properties && Objects")
                            ' CheckFormProperties2(filename, StudentAppForm)
                            CheckFormProperties2(filename, StudentAppForm)
                            CheckObjectNaming2(filename, StudentAppForm)
                            '    CheckFormLoad(StudentAppSummary, filesource)
                            ProcessReq(ReturnLastField(filename, "\"), StudentAppSummary(EnSummary.LogicFormLoad), "Form Load Method not found", "Form Load Method found", "LogicFormLoad")


                            BuildReport("Form Objects", StudentAssignment, StudentAppForm, StudentAppSummary, strStudentReport)
                        End If

                        x = x + 1
                        worker.ReportProgress(x, "Individual Work")
                        '        If rbnSingleProject.Checked Then worker.ReportProgress(-4, "Processing " & ReturnLastField(filename, "\") & " - Flow Control Structures")

                        x = x + 1
                        worker.ReportProgress(x, "Individual Work")
                        If rbnSingleProject.Checked Then worker.ReportProgress(-4, "Processing " & ReturnLastField(filename, "\") & " - Coding Standards")


                        ' -------------------------------- count number of comments ----------------------------------------
                        CheckForComments2(ReturnLastField(filename, "\"), filesource, StudentAppSummary, StudentAppForm, BackgroundWorker1)

                        BuildReport("Coding Standards", StudentAssignment, StudentAppForm, StudentAppSummary, strStudentReport)
                        x = x + 1
                        worker.ReportProgress(x, "Individual Work")

                        ' ---------------------------------Record summary data for Faculty Summary ------------------------

                        If first Then
                            strFacReport &= "<tr class=""newstudent""><td class=""newstudent"">" & strStudentID & "</td><td class=""newstudent"">" & ReturnLastField(filename, "\") & "</td><td class=""newstudent, tdright"">" & FileLinesOfCode.ToString("n0") & "</td><td class=""newstudent, tdcenter"">" & " - " & "</td></tr>" & vbCrLf
                        Else
                            strFacReport &= "<tr><td>" & "" & "</td><td>" & ReturnLastField(filename, "\") & "</td><td class=""tdright"">" & FileLinesOfCode.ToString("n0") & "</td><td  class=""tdcenter"">" & " - " & "</td></tr>" & vbCrLf
                        End If

                        integrateSSummary(StudentAppSummary, IntegratedStudentAssignment, filename, first)
                        integrateForm(StudentAppForm, IntegratedForm, filename, first)
                        first = False

                    Next filename      ' end of Filename loop
                    ' ---------------------------------------------------------------------------------------------------------
                    strFacReportDetail = BuildSummaryDetail(StudentAssignment, IntegratedForm, IntegratedStudentAssignment)

                    ' FindIntegratedScore(StudentAssignment, IntegratedForm, IntegratedStudentAssignment)

                    If TotalPossiblePts <> 0 Then
                        s = AssScore.ToString("n0") & " out of " & AssPossible.ToString("n0") & " = " & (AssScore / AssPossible).ToString("p1")
                    Else


                        s = AssScore.ToString("n0") & " out of " & AssPossible.ToString("n0")
                    End If

                    strFacReport &= "<tr><td>" & "" & "</td><td class=""boldtext"">" & "Overall Assessment" & "</td><td  class=""tdcenter, boldtext"">" & TotalLinesOfCode.ToString("n0") & "</td><td class=""tdcenter, boldtext"">" & s & "</td></tr>" & vbCrLf

                    Do While strStudentReport.Contains("<br>" & vbCrLf & " <br>")
                        strStudentReport = strStudentReport.Replace("<br>" & vbCrLf & " <br>", "<br>")
                    Loop

                    strStudentReport = strStudentReport.Replace("[CONFIGFILE]", lblConfigFile.Text)

                    strStudentReport = strStudentReport.Replace("[SCORE]", AssScore.ToString("n1") & " deduction out of " & AssPossible.ToString("n1") & " possible points = " & (AssScore / AssPossible).ToString("p1"))

                    sr = File.OpenText((Application.StartupPath & "\templates\rptStudentFooter.html"))
                    strStudentReport &= sr.ReadToEnd
                    sr.Close()


                    sw = File.AppendText(StudentReportPath)
                    sw.Write(strStudentReport)
                    sw.Close()
                    strStudentReport = ""
                    averageLOC += TotalLinesOfCode

                    ' ############################################################################################
                End If

            End If

            worker.ReportProgress(CInt(ii * 100 / Math.Max(1, nstudents)), "Checks")

            ' extract last compile time.
            SubmissionCompileTime = " - "
            SubmissionCompileDate = " - "
            Dim di As New IO.DirectoryInfo(path & "\")     ' lblDir.Text & "\" & strStudentID)
            Dim diar1 As IO.FileInfo() = di.GetFiles("*.exe", SearchOption.AllDirectories)
            Dim dra As IO.FileInfo

            'list the names of all files in the specified directory
            For Each dra In diar1
                If dra.DirectoryName.ToLower.Contains("\bin\debug") Then
                    If Not dra.FullName.ToLower.Contains(".vshost.exe") Then
                        SubmissionCompileTime = dra.CreationTime.ToLongTimeString()
                        SubmissionCompileDate = dra.CreationTime.ToLongDateString()
                    End If
                End If
            Next

            ' jhg Need to check how this handles multiple file submissions.
            strAssignmentSummary &= AddStudentDataToSummary(strStudentID, SubmissionCompileTime, SubmissionCompileDate, TotalLinesOfCode.ToString, TotalScore.ToString)

            ' jhg - may need to total submission info if the student submitted multiple files.

            If nfiles > 1 Then       ' no need to process an integrated application if it only has a single file
                BuildSummaryDetail(StudentAssignment, IntegratedForm, IntegratedStudentAssignment)
                s = AssScore.ToString("n1") & " out of " & AssPossible.ToString("n1") & " possible points = " & (AssScore / AssPossible).ToString("p1")
                '    s2 = "\" & strStudentID & "\Reports\" & strStudentID & "_IntegratedReport.html"
                s2 = "\Reports\" & strStudentID & "_IntegratedReport.html"
                CloseFacReport(strFacReportDetail, s2, s)
            End If
        Next ii    ' End of student loop

        ' -------------------------------------Create Faculty file. ----------------------------------------------

        '  strFacReport = strAssignmentSummary

        If IsFacultyVersion Then

            If Not cbxJustUnzip.Checked Then CloseFacReport(strFacReport, "\FacultySummaryReport.html", s)
            strFacReport = ""


            ' ############################################################################################

            ExtractGUID(strOutputPath)



            ' ----------------------- Save MD5 Data --------------------------------------
            If Not cbxJustUnzip.Checked Then
                Dim fname As String
                '   Dim sw As StreamWriter
                '   Dim lastUser As String = ""
                Dim lastMD5 As String = ""
                Dim ShortFilename As String = ""
                Dim DupFlag As Boolean
                '      Dim NoDesignerVB As Boolean = True

                fname = strOutputPath & "\" & "MD5Data-" & AssignmentName.Trim & ".txt"
                sw = File.CreateText(fname)
                sw.WriteLine("Dup" & vbTab & "MD5" & vbTab & "Filename" & vbTab & "User Name" & vbTab & "Full Filename")
                Sort_Submissions()

                DupFlag = False
                For i = 0 To Submissions.Count - 2
                    With Submissions(i)
                        If Not .Filename.ToUpper.Contains("DESIGNER.VB") And Not .Filename.ToUpper.Contains("ASSEMBLYINFO.VB") Then
                            ShortFilename = ReturnLastField(.Filename, "\")
                            If Submissions(i).vbCRC = Submissions(i + 1).vbCRC Or DupFlag = True Then
                                sw.WriteLine(("*" & vbTab & "'" & .vbCRC.PadRight(8) & vbTab & ShortFilename & vbTab & .UserID & vbTab & .Filename))
                                DupFlag = False
                            Else
                                sw.WriteLine((vbTab & "'" & .vbCRC.PadRight(8) & vbTab & ShortFilename & vbTab & .UserID & vbTab & .Filename))

                            End If
                            If Submissions(i).vbCRC = Submissions(i + 1).vbCRC Then DupFlag = True
                        End If
                    End With
                Next

                ' Print the last line
                If Submissions.Count > 2 Then
                    With Submissions(Submissions.Count - 1)
                        If Not .Filename.ToUpper.Contains("DESIGNER.VB") And Not .Filename.ToUpper.Contains("ASSEMBLYINFO.VB") Then

                            ShortFilename = ReturnLastField(.Filename, "\")

                            If Submissions(Submissions.Count - 2).vbCRC = Submissions(Submissions.Count - 1).vbCRC Then
                                sw.WriteLine(("*" & vbTab & "'" & .vbCRC.PadRight(8) & vbTab & ShortFilename & vbTab & .UserID & vbTab & .Filename))
                            Else
                                sw.WriteLine((vbTab & "'" & .vbCRC.PadRight(8) & vbTab & ShortFilename & vbTab & .UserID & vbTab & .Filename))
                            End If
                        End If
                    End With
                End If


                sw.Close()
                ' ------------------------------------Web page with dups ------------------------------------------------------

                fname = strOutputPath & "\" & "MD5Data-" & AssignmentName.Trim & ".html"
                sw = File.AppendText(fname)
                'Dim sr1 As StreamReader
                'Dim s As String

                'sw.Write(sr1.ReadToEnd)
                'sr1.Close()

                sw.WriteLine("<h2>Identical Student File MD5s</h2>" & vbCrLf)

                ' ---------------------------------------------------------------------------------------------------
                sw.WriteLine("<p><em>A MD5 is referred to as a message digest. It is a 128 bit value, represented as a 32 digit hexidimal number. The MD5 algorithm is commonly used to ensure data integrity. Any change in a data file will result in a different MD5 value. If two files have identical MD5 values, the files are identical. Note, some files that VB creates may not require modification, and therefore two separate applications will tend to have identical MD5 values for these files. Example include SplashScreens and AboutBoxes. Also, the application allows the loading of reference applications that may have been distributed to the class (such as sample applications). The application distinguishes files with duplicate MD5 values between those that do not also match instructor demo applications, and those that do. The only files checked are those ending in .vb.</em></p> <hr />" & vbCrLf)

                sr1 = File.OpenText(strOutputPath & "\" & "MD5Data-" & AssignmentName.Trim & ".txt")
                s = sr1.ReadLine


                Dim last As String = ""
                Dim ss() As String
                Dim instructordemo As Boolean
                Dim needtoclose As Boolean = False
                Dim allInstructorDemo As Boolean = False
                Dim tmp As String = ""

                ' ==================================================================================================
                Do While sr1.Peek > -1
                    s = sr1.ReadLine
                    ss = s.Split(CChar(vbTab))
                    If ss(1) = last Then
                        If ss(3) = "Instructor Demo" Then   ' this is an instructor supplied file
                            instructordemo = True
                            If needtoclose Then
                                sw.WriteLine("</ul>" & vbCrLf)
                                MD5Issues = True
                            End If
                            needtoclose = False

                            ' bypass this record
                        ElseIf instructordemo Then
                            If needtoclose Then
                                sw.WriteLine("</ul>" & vbCrLf)
                                MD5Issues = True
                            End If
                            needtoclose = False

                            ' bypass this record
                        Else
                            ' need to show this record.
                            sw.WriteLine("<li>" & ss(3) & " - " & ss(4) & "</li>")
                            needtoclose = True
                        End If
                    Else
                        last = ss(1)
                        If ss(3) = "Instructor Demo" Then   ' this is an instructor supplied file
                            instructordemo = True
                            If needtoclose Then
                                sw.WriteLine("</ul>" & vbCrLf)
                                MD5Issues = True
                            End If
                            needtoclose = False
                            ' bypass this record
                        Else
                            instructordemo = False
                            If needtoclose Then
                                sw.WriteLine("</ul>" & vbCrLf)
                                MD5Issues = True
                            End If

                            ' need to show this record.
                            If ss(0) = "*" Then
                                sw.WriteLine("<h3>" & ss(1) & " - " & ss(2) & "</h3>" & vbCrLf & "<ul>")
                                sw.WriteLine("<li>" & ss(3) & " - " & ss(4) & "</li>")
                                MD5Issues = True
                                needtoclose = True
                            End If
                        End If

                    End If


                Loop
                sr1.Close()
                ' ==================================================================================================
                sr1 = File.OpenText(strOutputPath & "\" & "MD5Data-" & AssignmentName.Trim & ".txt")
                s = sr1.ReadLine
                sw.WriteLine("<h2>Student File MD5's identical to Instructor Demo Files</h2>" & vbCrLf)
                ' ---------------------------------------------------------------------------------------------------
                sw.WriteLine("<p><em>This section lists files with identical MD5 values that match those from Instructor Demo files. Since all students would have access to these files, identical MD5 values do not necessarily indicate unauthorized collusion.</em></p> <hr />" & vbCrLf)

                ' display record the are identical to Instructor files
                Do While sr1.Peek > -1
                    s = sr1.ReadLine
                    ss = s.Split(CChar(vbTab))
                    If ss(1) = last Then
                        If ss(3) = "Instructor Demo" Then   ' this is an instructor supplied file
                            If ss(0) = "*" Then
                                instructordemo = True
                                allInstructorDemo = True
                                tmp &= ("<li>" & ss(3) & " - " & ss(4) & "</li>")
                            End If
                        ElseIf instructordemo Then
                            If ss(0) = "*" Then
                                allInstructorDemo = False
                                tmp &= ("<li>" & ss(3) & " - " & ss(4) & "</li>")
                            End If
                        Else
                            'If needtoclose Then
                            '    If Not allInstructorDemo Then
                            '        tmp &= "</ul>" & vbCrLf
                            '        sw.WriteLine(tmp)
                            '        tmp = ""
                            '    End If
                            'End If
                            'needtoclose = False

                        End If
                    Else    ' new MD5
                        last = ss(1)
                        If ss(3) = "Instructor Demo" Then   ' this is an instructor supplied file
                            instructordemo = True

                            If needtoclose And Not allInstructorDemo Then
                                tmp &= "</ul>" & vbCrLf
                                sw.WriteLine(tmp)
                                tmp = ""
                            Else
                                tmp = ""
                            End If

                            allInstructorDemo = True                        ' need to show this record.
                            If ss(0) = "*" Then
                                tmp = "<h3>" & ss(1) & " - " & ss(2) & "</h3>" & vbCrLf & "<ul>"
                                tmp &= "<li>" & ss(3) & " - " & ss(4) & "</li>"
                                needtoclose = True
                            End If
                        Else   ' student file - not based on instructor demo
                            If needtoclose And Not allInstructorDemo Then
                                tmp &= "</ul>" & vbCrLf
                                sw.WriteLine(tmp)
                                tmp = ""
                            Else
                                tmp = ""
                            End If

                            allInstructorDemo = False
                            instructordemo = False ' do not show this
                            ' bypass this record
                        End If

                    End If
                Loop

                If needtoclose And Not allInstructorDemo Then sw.WriteLine("</ul>" & vbCrLf)
                needtoclose = False

                sw.WriteLine("</body> </html>")
                sw.Close()

            End If '  If Not cbxJustUnzip.Checked
        End If ' if isFacultyVersion
        ' --------------------------------------------------------------------------------
        timeend = Now

    End Sub



    Function AddStudentDataToSummary(sn As String, ct As String, cd As String, tloc As String, sc As String) As String
        Dim tmp As String

        tmp = "<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td>"
        Return String.Format(tmp, sn, ct, cd, tloc, sc)
    End Function



    Sub LoadConfig()
        ' This is designed to process the Config file. The Config file must be in the same directory as the executable file. If found, it opens it up and processing the file line by line. Each line should have a Key Variable followed by a Property. Some Keys can only take a certain set of properties.  The code parses the Key/Property pair and sets the assocaited config value, which is defined in the JHGModule. This allows the user to modify the operation of the application.
        ' -----------------------------------------------------------------------------------------------------
        '  Dim msg As String = ""
        '  Dim msgflag As Boolean = False

        ' Check to see if the Config file exists. If not bypass the remainder of the the load process.
        If File.Exists(Application.StartupPath & "\templates\defaultConfig.cfg") Then

            '           frmConfig.LoadConfigFile(Application.StartupPath & "\templates\defaultConfig.cfg")
            '            lblConfigFile.Text = "Default Config"
        End If
    End Sub



    Function ErrorMSG(msg As String, newMsg As String) As String
        ' Used to append an error message to the Error message summary.

        Return msg & newMsg & vbCrLf
    End Function



    Private Sub btnOutput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOutput.Click, ViewOutputToolStripMenuItem.Click
        If rbnBlackboardZip.Checked Then
            ReportType = "Detail"
            frmPickStudentReport.Show()
        Else
            Dim url As New Uri("file:\\\" & StudentReportPath)

            frmOutput.WebBrowser1.Url = url
            frmOutput.Show()
        End If
    End Sub

    Private Sub btnDetail_Click(sender As Object, e As EventArgs) Handles btnDetail.Click
        If rbnBlackboardZip.Checked Then
            ReportType = "Integrated"
            frmPickStudentReport.Show()
        Else
            Dim url As New Uri("file:\\\" & StudentReportPath.Replace("GradeReport", "IntegratedReport"))

            frmOutput.WebBrowser1.Url = url
            frmOutput.Show()
        End If

    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs, Optional path As String = "", Optional txt As String = "") Handles BackgroundWorker1.ProgressChanged
        'This method is executed in the UI thread. 

        Select Case e.ProgressPercentage
            Case Is >= 0
                If TryCast(e.UserState, String) = "Extracting Student Work" Then
                    frmProgressBar.ProgressBar1.Value = e.ProgressPercentage
                    frmProgressBar.ProgressBar1.Text = "Extract Student Files"
                    Me.lblMessage.Text = TryCast(e.UserState, String)

                ElseIf TryCast(e.UserState, String) = "Individual Work" Then
                    frmProgressBar.ProgressBar3.Value = CInt(100 * e.ProgressPercentage / pbar3max)
                    '  Me.lblMessage.Text = "Processing Student: " & TryCast(e.UserState, String)
                    '       Beep()
                ElseIf TryCast(e.UserState, String) = "Preliminaries" Then
                    Me.lblMessage.Text = "Loading Instructor Demo Files"
                    frmProgressBar.ProgressBar1.Value = e.ProgressPercentage
                    frmProgressBar.ProgressBar1.Text = "Instructor Demo Files"

                ElseIf TryCast(e.UserState, String) = "Line by Line" Then
                    frmProgressBar.ProgressBar2.Value = e.ProgressPercentage

                ElseIf TryCast(e.UserState, String) = "Checks" Then
                    frmProgressBar.ProgressBar4.Value = e.ProgressPercentage

                Else
                    frmProgressBar.ProgressBar2.Value = e.ProgressPercentage
                    Me.lblMessage.Text = "Processing Student: " & TryCast(e.UserState, String)
                    '      Beep()
                End If


            Case -1
                Dim sw As StreamWriter = File.AppendText(path)
                sw.Write(txt)
                sw.Close()
            Case -2
                '               lbxBad.Items.Add(TryCast(e.UserState, String))
            Case -3
                '               lbxBad.Items.Add(TryCast(e.UserState, String))
            Case -4
                lblMessage.Text = TryCast(e.UserState, String)

                'Case -5
                '    lblMessage.Text &= TryCast(e.UserState, String)
        End Select

    End Sub

    'This method is executed in the UI thread.
    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted

        frmProgressBar.Close()

        rbnBlackboardZip.Enabled = True
        rbnSingleProject.Enabled = True

        Me.lblMessage.Text = "Processing Complete"
        If GuidIssues Then lblMessage.Text &= " *** GUID Issues Found "
        If MD5Issues Then lblMessage.Text &= " *** MD5 Issues Found "

        If GuidIssues Or MD5Issues Then btnViewPlagiarism.Visible = True

        btnOutput.Visible = True
        btnAssignSummary.Visible = True
        btnDetail.Visible = True
        ' ----------------------------------------------------------------------
        Dim dd As TimeSpan

        Dim sw As StreamWriter

        sw = File.AppendText(Application.StartupPath & "\TimeStudy.txt")
        sw.WriteLine("Today" & vbTab & Now.ToString)
        sw.WriteLine("Student Files" & vbTab & nstudentfiles.ToString("n1"))
        sw.WriteLine("Instructor Apps" & vbTab & ninstructorapps.ToString("n1"))
        sw.WriteLine("Instructor Files" & vbTab & ninstructorfiles.ToString("n1"))
        sw.WriteLine("Average LOC" & vbTab & (averageLOC / nstudentfiles).ToString("n2"))
        sw.WriteLine("")
        dd = timeLoadInstructorFiles.Subtract(timeStart)
        sw.WriteLine("Load Instructor files" & vbTab & timeLoadInstructorFiles.ToString("hh:mm:ss.fff tt") & vbTab & dd.ToString("ss\.fff"))
        dd = timeUnzipStart.Subtract(timeLoadInstructorFiles)
        sw.WriteLine("Unzip" & vbTab & timeUnzipStart.ToString("hh:mm:ss.fff tt") & vbTab & dd.ToString("ss\.fff"))
        dd = timeProcess.Subtract(timeUnzipStart)
        sw.WriteLine("Start Processing" & vbTab & timeProcess.ToString("hh:mm:ss.fff tt") & vbTab & dd.ToString("s\.fff"))
        dd = timeLoadInstructorFiles2.Subtract(timeProcess)
        sw.WriteLine("Load Instr. MD5" & vbTab & timeLoadInstructorFiles2.ToString("hh:mm:ss.fff tt") & vbTab & dd.ToString("ss\.fff"))
        dd = timeMD5.Subtract(timeLoadInstructorFiles2)
        sw.WriteLine("MD5" & vbTab & timeMD5.ToString("hh:mm:ss.fff tt") & vbTab & dd.ToString("s\.fff"))
        dd = timeend.Subtract(timeMD5)
        sw.WriteLine("end" & vbTab & timeend.ToString("hh:mm:ss.fff tt") & vbTab & dd.ToString("s\.fff"))
        sw.WriteLine("")
        dd = timeend.Subtract(timeStart)
        sw.WriteLine("Overall" & vbTab & "" & vbTab & dd.ToString("s\.fff"))
        sw.WriteLine("------------------------------------------------------------------------")
        sw.Close()
    End Sub

    Private Sub Sort_Submissions()
        Submissions = Submissions.OrderBy(Function(x) x.vbCRC).ToList
    End Sub



    Sub ExtractGUID(Path As String)
        Dim GUID As String = ""
        Dim fname As String
        Dim sw As StreamWriter
        Dim sr As StreamReader
        Dim s As String
        Dim student As String
        Dim tmp As String = ""
        Dim allInstructorDemos As Boolean
        Dim needtoprint As Boolean


        ' http://www.vbforums.com/showthread.php?721643-RESOLVED-Sorting-a-multidimensional-array

        'Dim GUIDData() As myGUIDData
        Dim i As Integer = 0
        Dim Match As Boolean
        Dim GUIDHeader As String

        GUIDs.Clear()


        timeLoadInstructorFiles2 = Now
        If Path > "" Then
            fname = lblDir.Text & "\GUIDData.txt"
            sw = File.CreateText(fname)
            ninstructorapps = 0
            ' first read in Instructor Demo Files
            If lblDemoDir.Text <> Nothing Then
                For Each myfile In Get_Files(lblDemoDir.Text, True, "AssemblyInfo.vb")
                    ' http://stackoverflow.com/questions/17007162/improve-this-function-to-get-recursively-all-files-inside-a-directory
                    '  MsgBox(file.Name)
                    sr = File.OpenText(myfile.FullName)
                    s = sr.ReadToEnd
                    sr.Close()

                    GUID = returnBetween(s, "<Assembly: Guid(""", """", True)
                    student = "Instructor Demo"

                    Dim newGUID As New GUIDData

                    'Set the GUID's properties
                    With newGUID
                        .FilePath = myfile.FullName
                        .GUID = GUID
                        .Student = student
                    End With

                    'Add it to the list
                    GUIDs.Add(newGUID)
                    ninstructorapps += 1
                Next
            End If

            ' --------------------------------------------------------------------
            timeMD5 = Now

            ' Now get student files.
            For Each myfile In Get_Files(Path, True, "AssemblyInfo.vb")
                ' http://stackoverflow.com/questions/17007162/improve-this-function-to-get-recursively-all-files-inside-a-directory
                '  MsgBox(file.Name)
                sr = File.OpenText(myfile.FullName)
                s = sr.ReadToEnd
                sr.Close()

                GUID = returnBetween(s, "<Assembly: Guid(""", """", True)
                student = returnBetween(myfile.FullName, Path & "\", "\", True)



                Dim newGUID As New GUIDData

                'Set the GUID's properties
                With newGUID
                    .FilePath = myfile.FullName
                    .GUID = GUID
                    .Student = student
                End With

                'Add it to the list
                GUIDs.Add(newGUID)

                sw.WriteLine(student & vbTab & GUID & vbTab & myfile.FullName)
            Next
            sw.Close()
            '  lblMessage.Text = "GUID file written."
            ' worker.ReportProgress(x, "Individual Work")
            ' --------------------------------------------------------------------


            'Initialize the web page
            ' ---------------------------------------------------------------------------------------------------
            fname = strOutputPath & "\" & "MD5Data-" & AssignmentName.Trim & ".html"
            sw = File.CreateText(fname)
            Dim sr1 As StreamReader = File.OpenText(Application.StartupPath & "\templates\MD5template.html")

            GUIDHeader = sr1.ReadToEnd
            sr1.Close()

            GUIDHeader = GUIDHeader.Replace("[title]", txtAssignmentName.Text)
            GUIDHeader = GUIDHeader.Replace("[DATE]", Today.ToString("d"))
            GUIDHeader = GUIDHeader.Replace("[VERSION]", Application.ProductVersion)


            sw.Write(GUIDHeader)

            sw.WriteLine("<h2>Identical Student File GUIDs</h2>" & vbCrLf)

            ' ---------------------------------------------------------------------------------------------------
            sw.WriteLine("<p><em>A GUID is 'Global Unique Indentification Number' assigned to the application when it is created. As the name implies, it is highly improbable that two applications will have the same GUID, even if created on different computers (odds are 1: 5.3×10^36). Identical GUIDS indicate that the whole application file was copied from the original. Note, identical GUIDs do not mean that the files are identical - the application could have been modified after it was copied. Similarly non-identical GUIDs do not mean that the file code has not been copied and pasted into a different application. This this check is useful to see if an application was copied, even if there was supperfical changes made to change the appearance of the form or code.</em></p> <hr />" & vbCrLf)

            ' now lets check to see if any of the GUID are the same.
            'Sort by List by UID
            GUIDs = GUIDs.OrderBy(Function(x) x.GUID).ToList

            'Display results in listbox
            Match = False
            allInstructorDemos = True
            needtoprint = False

            For i = 1 To GUIDs.Count - 1
                If GUIDs(i - 1).GUID = GUIDs(i).GUID Then
                    If Not Match Then
                        If Not allInstructorDemos Then
                            If needtoprint Then sw.WriteLine("</ul>" & vbCrLf)
                            needtoprint = False
                            sw.WriteLine(tmp)
                            GuidIssues = True
                            sw.WriteLine("</ul>" & vbCrLf)
                            tmp = ""
                        End If
                        ' Need to list the first student if there is a match
                        '                       lbxBad.Items.Add("------GUID Match -------")
                        '                      lbxBad.Items.Add(GUIDs(i - 1).Student)
                        If GUIDs(i - 1).Student = "Instructor Demo" Then allInstructorDemos = True Else allInstructorDemos = False

                        tmp = "<h3> GUID Match - " & GUIDs(i).GUID & "</h3>" & vbCrLf
                        tmp &= "<ul><li>" & GUIDs(i - 1).Student & "</li>" & vbCrLf
                        Match = True
                    End If
                    '                  lbxBad.Items.Add(GUIDs(i).Student)
                    If GUIDs(i).Student = "Instructor Demo" Then allInstructorDemos = True
                    tmp &= "<li>" & GUIDs(i).Student & "</li>" & vbCrLf
                Else
                    Match = False
                End If

            Next

            If Not allInstructorDemos Then
                If needtoprint Then
                    sw.WriteLine("</ul>" & vbCrLf)
                    GuidIssues = True
                End If

                needtoprint = False

                sw.WriteLine(tmp)
                tmp = ""
            End If

            sw.Close()

        Else
            '       MessageBox.Show("frmMain3 - " & "You must select a Directory to analyze.", "Folder Not Specified", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

#Region " Get Files "

    ' [ Get Files Function ]
    '
    ' Examples :
    '
    ' For Each file In Get_Files("C:\Windows", True, "AssemblyInfo.vb") : MsgBox(file.Name) : Next
    '

    ' Get Files {directory} {recursive} {filename}
    Private Function Get_Files(ByVal directory As String, ByVal recursive As Boolean, filename As String) As List(Of IO.FileInfo)
        Dim searchOpt As IO.SearchOption = If(recursive, IO.SearchOption.AllDirectories, IO.SearchOption.TopDirectoryOnly)
        Return IO.Directory.GetFiles(directory, filename, searchOpt).Select(Function(p) New IO.FileInfo(p)).ToList
    End Function

#End Region



    Private Sub ExitToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
        Application.Exit()
    End Sub


    Private Sub ProcessToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProcessToolStripMenuItem.Click
        frmConfig.Show()

    End Sub


    Private Sub btnSelectFile_Click(sender As Object, e As EventArgs) Handles btnSelectFile.Click, Button1.Click, btnSelectFile2.Click

        lblMessage.Text = ""
        frmPickStudentReport.Close()

        btnViewPlagiarism.Visible = False
        btnOutput.Visible = False
        btnAssignSummary.Visible = False

        'rbnBlackboardZip.Enabled = False
        'rbnSingleProject.Enabled = False
        btnProcessApps.Enabled = True

        If rbnBlackboardZip.Checked Then
            lblTarget.Text = "Target Zip File"
            OpenFileDialog1.InitialDirectory = Environment.SpecialFolder.MyDocuments.ToString
            OpenFileDialog1.Filter = "Zip files |*.zip|All files |*.*"
            OpenFileDialog1.FilterIndex = 0
            OpenFileDialog1.RestoreDirectory = True

            OpenFileDialog1.ShowDialog()
            lblTargetFile.Text = OpenFileDialog1.FileName


            AssignmentName = ReturnLastField(lblTargetFile.Text, "\")
            lblDir.Text = lblTargetFile.Text.Replace("\" & AssignmentName, "")


            AssignmentName = AssignmentName.Substring(0, AssignmentName.ToUpper.IndexOf(".ZIP"))
            txtAssignmentName.Text = AssignmentName

            ' ------------------------------------------------------------------------------------------
            ' This deletes the existing file. May want to ask the user about this ????????????????????????????????????????????????????????
            If Directory.Exists(lblDir.Text & "\" & AssignmentName) Then
                DeleteDirectory(lblDir.Text & "\" & AssignmentName)
            End If
            ' ------------------------------------------------------------------------------------------
        ElseIf rbnSingleProject.Checked Then
            lblTarget.Text = "Target Application"
            FolderBrowserDialog1.RootFolder = Environment.SpecialFolder.MyDocuments
            FolderBrowserDialog1.Description = "Select the folder than contains your Application. It normally is where your .sln file is located."

            FolderBrowserDialog1.ShowDialog()
            lblTargetFile.Text = FolderBrowserDialog1.SelectedPath


            AssignmentName = ReturnLastField(lblTargetFile.Text, "\")
            lblDir.Text = lblTargetFile.Text.Replace("\" & AssignmentName, "")

            txtAssignmentName.Text = AssignmentName
            ' ------------------------------------------------------------------------------------------
        End If

        Dim FileLocation As DirectoryInfo = New DirectoryInfo(lblDir.Text & "\")
        Dim fi As FileInfo() = FileLocation.GetFiles("*.docx")
        If fi.Length > 0 Then
            lblSelectedTemplate.Text = fi(0).ToString
            cbxLoadWordTemplate.Checked = True
            hasWordTemplateBeenSpecified = True
        Else
            cbxLoadWordTemplate.Checked = False
            hasWordTemplateBeenSpecified = False
        End If


        btnProcessApps.Enabled = True
        btnOutput.Enabled = True
    End Sub



    Public Sub LoadErrorComments()

        Dim s As String
        Dim ss() As String

        Dim newComment As New ErrComments
        Dim sr As New StreamReader(Application.StartupPath & "\templates\feedback.txt")

        Do While sr.Peek > -1

            s = sr.ReadLine

            ss = s.Split(CChar(vbTab))
            If ss.Length = 2 Then
                newcomment.topic = ss(0)
                newcomment.Comment = ss(1)

                newComment = New ErrComments
                newcomment.topic = ss(0)
                newcomment.Comment = ss(1)

                ErrorComments.Add(newComment)

            End If
        Loop
        sr.Close()
    End Sub

    Private Sub btnDemo_Click(sender As Object, e As EventArgs) Handles btnDemo.Click, IndentifyInstructorDemoDirectoryToolStripMenuItem.Click

        '  Dim hasVBFile As Boolean
        Dim path As String = lblDemoDir.Text

        '  not implemented yet, but check out http://www.vbforums.com/showthread.php?781747-RESOLVED-Setting-a-directory-that-is-remembered-when-the-application-is-opened-again-(Listbox)&highlight=folderBrowserdialog

        FolderBrowserDialog1.RootFolder = Environment.SpecialFolder.Personal
        FolderBrowserDialog1.SelectedPath = path

        If FolderBrowserDialog1.ShowDialog() <> DialogResult.Cancel Then
            lblDemoDir.Text = FolderBrowserDialog1.SelectedPath
            path = FolderBrowserDialog1.SelectedPath

            ' Save path
            'Dim sw As StreamWriter = File.CreateText(AppDataDir & "\demodir.txt")
            'sw.WriteLine(path)
            'sw.Close()
        End If


    End Sub

    Private Sub rbnSingleProject_CheckedChanged(sender As Object, e As EventArgs) Handles rbnSingleProject.CheckedChanged, rbnBlackboardZip.CheckedChanged
        btnProcessApps.Enabled = False
    End Sub


    Private Sub btnViewPlagiarism_Click(sender As Object, e As EventArgs) Handles btnViewPlagiarism.Click, PlagiarismSummaryToolStripMenuItem.Click
        If File.Exists(strOutputPath & "\" & "MD5Data-" & AssignmentName.Trim & ".html") Then
            Dim url As New Uri("file:\\\" & strOutputPath & "\" & "MD5Data-" & AssignmentName.Trim & ".html")

            frmOutput.WebBrowser1.Url = url
            '   frmOutput.Close()
            frmOutput.Show()
        End If
    End Sub


    Private Sub btnExit_Click_1(sender As Object, e As EventArgs) Handles btnExit.Click
        ' Exit application.
        Me.Close()
        Application.Exit()
    End Sub


    Private Sub HelpToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HelpToolStripMenuItem.Click
        frmAboutBox1.Show()
    End Sub

    Private Sub rbnCheckGray_CheckedChanged(sender As Object, e As EventArgs) Handles rbnCheckGray.CheckedChanged
        If rbnCheckGray.Checked Then HideGray = "Gray"
    End Sub

    Private Sub rbnShowAll_CheckedChanged(sender As Object, e As EventArgs) Handles rbnShowAll.CheckedChanged
        If rbnShowAll.Checked Then HideGray = "ShowAll"
    End Sub

    Private Sub rbnShowOnlyReq_CheckedChanged(sender As Object, e As EventArgs) Handles rbnShowOnlyReq.CheckedChanged
        If rbnShowOnlyReq.Checked Then HideGray = "OnlyReq"

    End Sub

    Private Sub btnAssignSummary_Click(sender As Object, e As EventArgs) Handles btnAssignSummary.Click
        If File.Exists(strOutputPath & "\FacultySummaryReport.html") Then
            Dim url As New Uri("file:\\\" & strOutputPath & "\FacultySummaryReport.html")

            frmOutput.WebBrowser1.Url = url
            '      frmOutput.Close()
            frmOutput.Show()
        End If
    End Sub


    Private Sub btnLoadAssessmentConfig_Click(sender As Object, e As EventArgs) Handles btnLoadAssessmentConfig.Click
        Dim filename As String

        OpenFileDialog1.Filter = "Config files |*.cfg"
        OpenFileDialog1.FilterIndex = 0
        OpenFileDialog1.RestoreDirectory = True

        OpenFileDialog1.ShowDialog()
        filename = OpenFileDialog1.FileName

        If filename.ToLower.Contains(".cfg") Then
            lblConfigFile.Text = filename
            frmConfig.LoadConfigFile(filename)
        End If

    End Sub

    Private Sub rbnDefaultCFG_CheckedChanged(sender As Object, e As EventArgs) Handles rbnDefaultCFG.CheckedChanged
        If rbnDefaultCFG.Checked Then
            lblConfigFile.Text = "Default Configuration"
        End If

    End Sub

    Private Sub rbnAppCFG_CheckedChanged(sender As Object, e As EventArgs) Handles rbnAppCFG.CheckedChanged
        If rbnAppCFG.Checked Then lblConfigFile.Text = ""
    End Sub
End Class
