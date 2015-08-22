Imports System.IO

Public Class frmConfig
    Dim previousRow As Integer
    Dim dgv As DataGridView
    Dim loadphase As Boolean
    Dim NeedsToBeSaved As Boolean = False

    Private Sub frmConfig_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' ------------------------------------------------------------------------------------------
        ' Acknowledgement
        ' The application uses the Expandable Groupbox object created by Nick Theissen, posted in VBForums.com (http://www.vbforums.com/showthread.php?576403-Expandable-GroupBox&p=3560002#post3560002)
        ' ------------------------------------------------------------------------------------------
        ' col1 - Property
        ' col2 - Criteria
        ' col3 - Active
        ' col4 - Pt per Error
        ' col5 - Max Deduct
        ' -------------------------------------------------------------------------------------------
        ' force the location and size of the gridviews

        lblMsg.Visible = False

        lblLanguage.Location = New Point(36, 37)
        cbxLanguage.Location = New Point(213, 37)
        cbxLanguage.Width = 125

        lblAppTitle.Location = New Point(36, 88)
        txtAssignmentTitle.Location = New Point(213, 88)
        txtAssignmentTitle.Width = 834
        btnSaveAsDefault.Visible = False
        btnSaveConfig.Visible = False

        StuffAlldgvs()

        ' not sure why we apparently load this twice.

        LoadCfgFile(Application.StartupPath & "\templates\defaultConfig.cfg")
        LoadConfigFile(Application.StartupPath & "\templates\defaultConfig.cfg")


    End Sub



    Public Sub StuffAlldgvs()
        loadphase = True
        ' loads in something for a default.
        Dim row As String()

        'Dev Environment
        ' -----------------------------------------------------------------------------------------------
        With dgvDevEnvironment
            row = New String() {"hasSLN", "SLN File", "", vbTrue.ToString, vbTrue.ToString, "10", "10", ""}
            .Rows.Add(row)
            row = New String() {"hasvbProj", "vbProj File", "", vbTrue.ToString, vbTrue.ToString, "10", "10", ""}
            .Rows.Add(row)
            row = New String() {"hasVBVersion", "VB Version", "Provided by the Student's Installation", vbTrue.ToString, vbTrue.ToString, "0", "0", ""}
            .Rows.Add(row)
            'row = New String() {"", "Turn On Line Numbers", "Not Implimented", vbFalse.ToString, vbFalse.ToString, "0", "0", ""}
            '.Rows.Add(row)
            'row = New String() {"", "Turn On Word Wrap", "Not Implimented", vbFalse.ToString, vbFalse.ToString, "0", "0", ""}
            '.Rows.Add(row)
        End With

        ' Application Info
        ' -----------------------------------------------------------------------------------------------
        With dgvAppInfo
            row = New String() {"hasSplashScreen", "Add Splash Screen", "", vbFalse.ToString, vbFalse.ToString, "3", "3", ""}
            .Rows.Add(row)
            row = New String() {"hasAboutBox", "Add About Box", "", vbFalse.ToString, vbFalse.ToString, "3", "3", ""}
            .Rows.Add(row)
            row = New String() {"InfoAppTitle", "Modify App Title", Application.ProductName, vbFalse.ToString, vbFalse.ToString, "3", "3", ""}
            .Rows.Add(row)
            row = New String() {"InfoDescription", "Modify Description", My.Application.Info.Description, vbFalse.ToString, vbFalse.ToString, "2", "5", ""}
            .Rows.Add(row)
            row = New String() {"InfoProduct", "Modify Product", "", vbFalse.ToString, vbFalse.ToString, "2", "2", ""}     ' ??????????????????????
            .Rows.Add(row)
            row = New String() {"InfoCompany", "Modify Company", "", vbFalse.ToString, vbFalse.ToString, "2", "2", ""}     ' ??????????????????????
            .Rows.Add(row)
            row = New String() {"InfoTrademark", "Modify Trademark", My.Application.Info.Trademark, vbFalse.ToString, vbFalse.ToString, "2", "2", ""}
            .Rows.Add(row)
            row = New String() {"InfoCopyright", "Modify Copyright", My.Application.Info.Copyright, vbFalse.ToString, vbFalse.ToString, "2", "2", ""}
            .Rows.Add(row)
        End With

        ' Compile Options
        ' -----------------------------------------------------------------------------------------------
        With dgvCompileOptions
            row = New String() {"OptionStrict", "Option Strict On", "", vbTrue.ToString, vbTrue.ToString, "10", "10", ""}
            .Rows.Add(row)
            row = New String() {"OptionExplicit", "Option Explicit On", "", vbFalse.ToString, vbFalse.ToString, "3", "3", ""}
            .Rows.Add(row)
        End With

        ' Commments
        ' -----------------------------------------------------------------------------------------------
        With dgvComments
            row = New String() {"CommentGeneral", "General Feedback on Comments", "", vbTrue.ToString, vbTrue.ToString, "0", "0", ""}
            .Rows.Add(row)
            row = New String() {"CommentSubs", "Req Comments in Subs", "", vbTrue.ToString, vbTrue.ToString, "1", "5", ""}
            .Rows.Add(row)
            row = New String() {"CommentIF", "Req Comment for IF", "", vbTrue.ToString, vbTrue.ToString, "1", "5", ""}
            .Rows.Add(row)
            row = New String() {"CommentFOR", "Req Comment for FOR", "", vbTrue.ToString, vbTrue.ToString, "1", "5", ""}
            .Rows.Add(row)
            row = New String() {"CommentDO", "Req Comment for DO", "", vbTrue.ToString, vbTrue.ToString, "1", "5", ""}
            .Rows.Add(row)
            row = New String() {"CommentWHILE", "Req Comment for While", "", vbTrue.ToString, vbTrue.ToString, "1", "3", ""}
            .Rows.Add(row)
            row = New String() {"CommentSELECT", "Req Comment for Select Case", "", vbTrue.ToString, vbTrue.ToString, "3", "3", ""}
            .Rows.Add(row)
        End With

        ' Form Design
        ' -----------------------------------------------------------------------------------------------
        With dgvFormDesign
            row = New String() {"RenameObjects", "Rename Form Names with Prefixes Identifying Object Types", "Objects referenced in the code need to include proper prefix indicating object type", vbTrue.ToString, vbTrue.ToString, "1", "5", ""}
            .Rows.Add(row)
            row = New String() {"IncludeFrmInFormName", "Include frm in Form Name", "", vbFalse.ToString, vbFalse.ToString, "2", "5", ""}
            .Rows.Add(row)
            row = New String() {"ChangeFormText", "Change Form Text", "Displayed text at top of form should be descriptive", vbTrue.ToString, vbTrue.ToString, "5", "5", ""}
            .Rows.Add(row)
            row = New String() {"ChangeFormColor", "Change Form Background Color", "", vbTrue.ToString, vbTrue.ToString, "5", "5", ""}
            .Rows.Add(row)
            row = New String() {"SetFormAcceptButton", "Set Form Accept Button", "Should be set at design time using AcceptButton Property ", vbTrue.ToString, vbTrue.ToString, "5", "5", ""}
            .Rows.Add(row)
            row = New String() {"SetFormCancelButton", "Set Form Cancel Button", "Should be set at design time using CancelButton Property ", vbTrue.ToString, vbTrue.ToString, "5", "5", ""}
            .Rows.Add(row)
            row = New String() {"ModifyStartPosition", "Modify Form Start Position", "", vbTrue.ToString, vbTrue.ToString, "5", "5", ""}
            .Rows.Add(row)
            'row = New String() {"LogicFormLoad", "Utilize Form Load Method", "", vbTrue.ToString, vbTrue.ToString, "3", "3", ""}
            '.Rows.Add(row)
        End With


        ' Form Objects
        ' -----------------------------------------------------------------------------------------------
        With dgvFormObj
            row = New String() {"ButtonObj", "Use of Button Objects", "", vbTrue.ToString, vbTrue.ToString, "1", "3", ""}
            .Rows.Add(row)
            row = New String() {"TextboxObj", "Use of Textbox Objects", "", vbFalse.ToString, vbFalse.ToString, "2", "5", ""}
            .Rows.Add(row)
            row = New String() {"ActiveLabels", "Include lbl Prefix on Active Labels", "", vbFalse.ToString, vbFalse.ToString, "2", "5", ""}
            .Rows.Add(row)
            row = New String() {"NonActiveLabels", "NonActive Labels Objects", "", vbTrue.ToString, vbTrue.ToString, "2", "2", ""}
            .Rows.Add(row)
            row = New String() {"ComboBoxObj", "Use of a ComboBox Objects", "", vbTrue.ToString, vbTrue.ToString, "2", "2", ""}
            .Rows.Add(row)
            row = New String() {"ListBoxObj", "Use of a ListBox Objects", "", vbTrue.ToString, vbTrue.ToString, "2", "2", ""}
            .Rows.Add(row)
            row = New String() {"RadioButtonObj", "Use of Radio Buttons Objects", "", vbTrue.ToString, vbTrue.ToString, "2", "2", ""}
            .Rows.Add(row)
            row = New String() {"CheckBoxObj", "Use of CheckBox Objects", "", vbTrue.ToString, vbTrue.ToString, "2", "2", ""}
            .Rows.Add(row)
            row = New String() {"GroupBoxObj", "Use of GroupBox Objects.", "", vbTrue.ToString, vbTrue.ToString, "2", "2", ""}
            .Rows.Add(row)
            row = New String() {"OpenFileDialogObj", "Use of an OpenFileDialog Objects.", "", vbTrue.ToString, vbTrue.ToString, "2", "2", ""}
            .Rows.Add(row)
            row = New String() {"SaveFileDialogObj", "Use of SaveFileDialog Objects.", "", vbTrue.ToString, vbTrue.ToString, "2", "2", ""}
            .Rows.Add(row)
            row = New String() {"WebBrowserObj", "Use of WebBrowser Objects.", "", vbTrue.ToString, vbTrue.ToString, "2", "2", ""}
            .Rows.Add(row)
        End With



        ' Imports
        ' -----------------------------------------------------------------------------------------------
        With dgvImports
            row = New String() {"SystemIO", "Imports System.IO", "", vbFalse.ToString, vbFalse.ToString, "5", "5", ""}
            .Rows.Add(row)
            row = New String() {"SystemNet", "Imports System.Net", "", vbFalse.ToString, vbFalse.ToString, "5", "5", ""}
            .Rows.Add(row)
            row = New String() {"SystemDB", "Imports System.DB", "", vbFalse.ToString, vbFalse.ToString, "5", "5", ""}
            .Rows.Add(row)
        End With

        ' Data Structures
        ' -----------------------------------------------------------------------------------------------
        With dgvDataStructures
            row = New String() {"VarArrays", "Use of Array", "", vbTrue.ToString, vbTrue.ToString, "5", "5", ""}
            .Rows.Add(row)
            row = New String() {"VarLists", "Use of List", "", vbTrue.ToString, vbTrue.ToString, "5", "5", ""}
            .Rows.Add(row)
            row = New String() {"VarStructures", "Use of Structure", "", vbFalse.ToString, vbFalse.ToString, "5", "5", ""}
            .Rows.Add(row)
        End With
        ' Variable Data Types
        ' -----------------------------------------------------------------------------------------------
        With dgvDataTypes
            row = New String() {"VarString", "Use of String", "", vbFalse.ToString, vbFalse.ToString, "5", "5", ""}
            .Rows.Add(row)
            row = New String() {"VarInteger", "Use of Integer", "", vbTrue.ToString, vbTrue.ToString, "5", "5", ""}
            .Rows.Add(row)
            row = New String() {"VarDecimal", "Use of Decimal", "", vbTrue.ToString, vbTrue.ToString, "5", "5", ""}
            .Rows.Add(row)
            row = New String() {"VarDate", "Use of Date", "", vbTrue.ToString, vbTrue.ToString, "5", "5", ""}
            .Rows.Add(row)
            row = New String() {"VarBoolean", "Use of Boolean", "", vbTrue.ToString, vbTrue.ToString, "5", "5", ""}
            .Rows.Add(row)
            row = New String() {"VariablePrefixes", "Rename Variables with Prefixes Identifying Data Types", "", vbTrue.ToString, vbTrue.ToString, "1", "5", ""}
            .Rows.Add(row)
        End With

        ' Program Logic
        ' -----------------------------------------------------------------------------------------------
        With dgvCoding
            row = New String() {"LogicIF", "Req use of IF", "", vbTrue.ToString, vbTrue.ToString, "1", "5", ""}
            .Rows.Add(row)
            row = New String() {"LogicFOR", "Req use of FOR", "", vbTrue.ToString, vbTrue.ToString, "1", "5", ""}
            .Rows.Add(row)
            row = New String() {"LogicDO", "Req use of DO", "", vbTrue.ToString, vbTrue.ToString, "1", "5", ""}
            .Rows.Add(row)
            row = New String() {"LogicWHILE", "Req use of WHILE", "", vbTrue.ToString, vbTrue.ToString, "1", "5", ""}
            .Rows.Add(row)
            row = New String() {"LogicElse", "Req use of Else", "", vbTrue.ToString, vbTrue.ToString, "1", "5", ""}
            .Rows.Add(row)
            row = New String() {"LogicElseIF", "Req use of ElseIF", "", vbTrue.ToString, vbTrue.ToString, "1", "5", ""}
            .Rows.Add(row)
            row = New String() {"LogicMessageBox", "Req use of MessageBox", "", vbTrue.ToString, vbTrue.ToString, "2", "2", ""}
            .Rows.Add(row)
            row = New String() {"LogicNestedIF", "Req use of Nested IF", "", vbTrue.ToString, vbTrue.ToString, "2", "2", ""}
            .Rows.Add(row)
            row = New String() {"LogicNestedFOR", "Req use of Nested FOR", "", vbTrue.ToString, vbTrue.ToString, "2", "2", ""}
            .Rows.Add(row)
            row = New String() {"LogicSelectCase", "Req use of Select Case", "", vbTrue.ToString, vbTrue.ToString, "1", "5", ""}
            .Rows.Add(row)
            row = New String() {"LogicConcatination", "Req String Concatination", "", vbTrue.ToString, vbTrue.ToString, "1", "5", ""}
            .Rows.Add(row)
            row = New String() {"LogicConvertToString", "Req Conversion to String (CStr or .toString)", "", vbTrue.ToString, vbTrue.ToString, "5", "5", ""}
            .Rows.Add(row)
            row = New String() {"LogicStringFormat", "Req Formatting of String (.tostring() or String.Format()", "", vbTrue.ToString, vbTrue.ToString, "5", "5", ""}
            .Rows.Add(row)
            row = New String() {"LogicStringFormatParameters", "Req Parameterized string format (string.Format(f,{0, ""})", "", vbTrue.ToString, vbTrue.ToString, "5", "5", ""}
            .Rows.Add(row)
            row = New String() {"LogicComplexConditions", "Req Complex Conditions (looks for AND / OR / ANDALSO / ORELSE)", "", vbTrue.ToString, vbTrue.ToString, "1", "5", ""}
            .Rows.Add(row)
            row = New String() {"LogicCaseInsensitive", "Req Case Insensitive", "This involves using either .toUpper or .toLower", vbTrue.ToString, vbTrue.ToString, "1", "5", ""}
            .Rows.Add(row)
            row = New String() {"LogicTryCatch", "Req use of Try ... Catch", "", vbTrue.ToString, vbTrue.ToString, "1", "5", ""}
            .Rows.Add(row)
            row = New String() {"LogicStreamReader", "Req use of ScreenReader", "", vbTrue.ToString, vbTrue.ToString, "1", "5", ""}
            .Rows.Add(row)
            row = New String() {"LogicStreamWriter", "Req use of ScreenWriter", "", vbTrue.ToString, vbTrue.ToString, "1", "5", ""}
            .Rows.Add(row)
            row = New String() {"LogicStreamReaderClose", "Req matching ScreenReader.Close", "", vbTrue.ToString, vbTrue.ToString, "1", "5", ""}
            .Rows.Add(row)
            row = New String() {"LogicStreamWriterClose", "Req matching ScreenWriter.Close", "", vbTrue.ToString, vbTrue.ToString, "1", "5", ""}
            .Rows.Add(row)
            row = New String() {"ObjOpenFileDialog", "Req use of Open File Dialog", "", vbTrue.ToString, vbTrue.ToString, "3", "3", ""}
            .Rows.Add(row)
            row = New String() {"ObjSaveFileDialog", "Req use of Save File Dialog", "", vbTrue.ToString, vbTrue.ToString, "3", "3", ""}
            .Rows.Add(row)
        End With
        ' ---------------------------------------------------------------------------------------------
        ' Subs/Functions
        ' ------------------------------------------------------------------------------------------------
        With dgvSubs
            row = New String() {"LogicSub", "Req use of User Defined Subroutine or Function", "", vbTrue.ToString, vbTrue.ToString, "1", "5", ""}
            .Rows.Add(row)
            row = New String() {"LogicOptional", "Req use of Optional Parameters in Sub / Function", "", vbTrue.ToString, vbTrue.ToString, "1", "5", ""}
            .Rows.Add(row)
            row = New String() {"LogicByRef", "Req use of ByRef Parameters in Sub / Function", "", vbTrue.ToString, vbTrue.ToString, "1", "5", ""}
            .Rows.Add(row)
            row = New String() {"LogicMultipleForms", "Include Multiple Forms", "", vbFalse.ToString, vbFalse.ToString, "5", "5", ""}
            .Rows.Add(row)
            row = New String() {"LogicModule", "Include Module", "", vbFalse.ToString, vbFalse.ToString, "5", "5", ""}
            .Rows.Add(row)
            row = New String() {"LogicFormLoad", "Include a Form LOAD Method", "", vbTrue.ToString, vbTrue.ToString, "5", "5", ""}
            .Rows.Add(row)
        End With

        loadphase = False
        ' Best Practice
        ' -----------------------------------------------------------------------------------------------
    End Sub

    Private Sub frmConfig_SizeChanged(sender As Object, e As EventArgs) Handles MyBase.SizeChanged

        TabControl1.Width = Me.Width - 40

        dgvAppInfo.Width = TabControl1.Width - 30
        dgvCompileOptions.Width = TabControl1.Width - 30
        dgvComments.Width = TabControl1.Width - 30
        dgvFormDesign.Width = TabControl1.Width - 30
    End Sub


    'Private Sub btnC_Click(sender As Object, e As EventArgs)
    '    Dim ans As Integer
    '    ans = MessageBox.Show("This information was not saved. Are you sure you want to return without saving your settings?", "Exit without Saving", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

    '    If ans = vbYes Then
    '        Me.Close()
    '        frmMain.Show()
    '    End If
    'End Sub

    Private Sub btnSetAsDefault_Click(sender As Object, e As EventArgs)

        Dim sw As New StreamWriter(Application.StartupPath & "\templates\defaultConfig.cfg")
        Dim s As String

        sw.WriteLine("AppGrader Assignment Configuration File")  ' this first line is requried. It is used to identify it as a Config File

        sw.WriteLine("CfgLanguage" & vbTab & "VB")
        sw.WriteLine("cfgAssignmentTitle" & vbTab & "")

        '  sw.WriteLine("CfgPath1" & vbTab & "MyDocuments")
        ' -------------------------------------------------------------------------------------------

        'For Each arow As DataGridViewRow In dgvAdvanced.Rows     ' this needs to be the first page
        '    s = arow.Cells(0).ToString & vbTab
        '    s = arow.Cells(1).ToString & vbTab
        '    s = arow.Cells(2).ToString & vbTab
        '    s = arow.Cells(3).ToString & vbTab
        '    s = arow.Cells(4).ToString
        '    sw.WriteLine(s)
        'Next

        For Each arow As DataGridViewRow In dgvDevEnvironment.Rows
            s = arow.Cells(0).ToString & vbTab
            s = arow.Cells(1).ToString & vbTab
            s = arow.Cells(2).ToString & vbTab
            s = arow.Cells(3).ToString & vbTab
            s = arow.Cells(4).ToString & vbTab
            s = arow.Cells(5).ToString & vbTab
            s = arow.Cells(6).ToString & vbTab
            s = arow.Cells(7).ToString & vbTab
            s = arow.Cells(8).ToString
            sw.WriteLine(s)
        Next

        For Each arow As DataGridViewRow In dgvAppInfo.Rows
            s = arow.Cells(0).ToString & vbTab
            s = arow.Cells(1).ToString & vbTab
            s = arow.Cells(2).ToString & vbTab
            s = arow.Cells(3).ToString & vbTab
            s = arow.Cells(4).ToString & vbTab
            s = arow.Cells(5).ToString & vbTab
            s = arow.Cells(6).ToString & vbTab
            s = arow.Cells(7).ToString & vbTab
            s = arow.Cells(8).ToString
            sw.WriteLine(s)
        Next

        For Each arow As DataGridViewRow In dgvCompileOptions.Rows
            s = arow.Cells(0).ToString & vbTab
            s = arow.Cells(1).ToString & vbTab
            s = arow.Cells(2).ToString & vbTab
            s = arow.Cells(3).ToString & vbTab
            s = arow.Cells(4).ToString & vbTab
            s = arow.Cells(5).ToString & vbTab
            s = arow.Cells(6).ToString & vbTab
            s = arow.Cells(7).ToString & vbTab
            s = arow.Cells(8).ToString
            sw.WriteLine(s)
        Next

        For Each arow As DataGridViewRow In dgvComments.Rows
            s = arow.Cells(0).ToString & vbTab
            s = arow.Cells(1).ToString & vbTab
            s = arow.Cells(2).ToString & vbTab
            s = arow.Cells(3).ToString & vbTab
            s = arow.Cells(4).ToString & vbTab
            s = arow.Cells(5).ToString & vbTab
            s = arow.Cells(6).ToString & vbTab
            s = arow.Cells(7).ToString & vbTab
            s = arow.Cells(8).ToString
            sw.WriteLine(s)
        Next

        For Each arow As DataGridViewRow In dgvFormDesign.Rows
            s = arow.Cells(0).ToString & vbTab
            s = arow.Cells(1).ToString & vbTab
            s = arow.Cells(2).ToString & vbTab
            s = arow.Cells(3).ToString & vbTab
            s = arow.Cells(4).ToString & vbTab
            s = arow.Cells(5).ToString & vbTab
            s = arow.Cells(6).ToString & vbTab
            s = arow.Cells(7).ToString & vbTab
            s = arow.Cells(8).ToString
            sw.WriteLine(s)
        Next

        For Each arow As DataGridViewRow In dgvFormObj.Rows
            s = arow.Cells(0).ToString & vbTab
            s = arow.Cells(1).ToString & vbTab
            s = arow.Cells(2).ToString & vbTab
            s = arow.Cells(3).ToString & vbTab
            s = arow.Cells(4).ToString & vbTab
            s = arow.Cells(5).ToString & vbTab
            s = arow.Cells(6).ToString & vbTab
            s = arow.Cells(7).ToString & vbTab
            s = arow.Cells(8).ToString
            sw.WriteLine(s)
        Next

        For Each arow As DataGridViewRow In dgvImports.Rows
            s = arow.Cells(0).ToString & vbTab
            s = arow.Cells(1).ToString & vbTab
            s = arow.Cells(2).ToString & vbTab
            s = arow.Cells(3).ToString & vbTab
            s = arow.Cells(4).ToString & vbTab
            s = arow.Cells(5).ToString & vbTab
            s = arow.Cells(6).ToString & vbTab
            s = arow.Cells(7).ToString & vbTab
            s = arow.Cells(8).ToString
            sw.WriteLine(s)
        Next

        For Each arow As DataGridViewRow In dgvDataStructures.Rows
            s = arow.Cells(0).ToString & vbTab
            s = arow.Cells(1).ToString & vbTab
            s = arow.Cells(2).ToString & vbTab
            s = arow.Cells(3).ToString & vbTab
            s = arow.Cells(4).ToString & vbTab
            s = arow.Cells(5).ToString & vbTab
            s = arow.Cells(6).ToString & vbTab
            s = arow.Cells(7).ToString & vbTab
            s = arow.Cells(8).ToString
            sw.WriteLine(s)
        Next

        For Each arow As DataGridViewRow In dgvDataTypes.Rows
            s = arow.Cells(0).ToString & vbTab
            s = arow.Cells(1).ToString & vbTab
            s = arow.Cells(2).ToString & vbTab
            s = arow.Cells(3).ToString & vbTab
            s = arow.Cells(4).ToString & vbTab
            s = arow.Cells(5).ToString & vbTab
            s = arow.Cells(6).ToString & vbTab
            s = arow.Cells(7).ToString & vbTab
            s = arow.Cells(8).ToString
            sw.WriteLine(s)
        Next

        For Each arow As DataGridViewRow In dgvCoding.Rows
            s = arow.Cells(0).ToString & vbTab
            s = arow.Cells(1).ToString & vbTab
            s = arow.Cells(2).ToString & vbTab
            s = arow.Cells(3).ToString & vbTab
            s = arow.Cells(4).ToString & vbTab
            s = arow.Cells(5).ToString & vbTab
            s = arow.Cells(6).ToString & vbTab
            s = arow.Cells(7).ToString & vbTab
            s = arow.Cells(8).ToString
            sw.WriteLine(s)
        Next

        For Each arow As DataGridViewRow In dgvSubs.Rows
            s = arow.Cells(0).ToString & vbTab
            s = arow.Cells(1).ToString & vbTab
            s = arow.Cells(2).ToString & vbTab
            s = arow.Cells(3).ToString & vbTab
            s = arow.Cells(4).ToString & vbTab
            s = arow.Cells(5).ToString & vbTab
            s = arow.Cells(6).ToString & vbTab
            s = arow.Cells(7).ToString & vbTab
            s = arow.Cells(8).ToString

            sw.WriteLine(s)
        Next

        sw.Close()

        lblMsg.Text = "Configuration Saved."
        lblMsg.Visible = True
    End Sub

    Private Sub btnLoadAssessmentFile_Click(sender As Object, e As EventArgs) Handles btnLoadAssessmentFile.Click

        OpenFileDialog1.ShowDialog()

        OpenFileDialog1.Filter = "Config|*.cfg|All Files|*.*;"
        If OpenFileDialog1.FileName <> Nothing Then
            File.Delete(Application.StartupPath & "\CantFind.txt")
            If File.Exists(Application.StartupPath & "\MissingSetting.txt") Then File.Delete(Application.StartupPath & "\MissingSetting.txt")

            LoadConfigFile(OpenFileDialog1.FileName)
            Settings.LoadCfgFile(OpenFileDialog1.FileName)
            lblAssessmentFile.Text = OpenFileDialog1.FileName
        End If


    End Sub

    Private Sub ReturnToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ReturnToolStripMenuItem1.Click, btnReturn.Click
        ' check to see if the current settings have been saved
        ' Warn user if not saved

        Dim UserResp As Integer

        If NeedsToBeSaved Then
            UserResp = MessageBox.Show("Your changes have not been saved. Do you want to exit without saving?", "Configuration not Saved", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
        End If

        If Not NeedsToBeSaved Or UserResp = vbYes Then
            frmMain.Show()
            Me.Close()
        End If


    End Sub


    Private Sub SaveNewConfigToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveNewConfigToolStripMenuItem.Click, btnSaveConfig.Click

        '    SaveFileDialog1.ShowDialog()
        If lblTargetConfig.Text > "" Then
            SaveConfigFile(lblTargetConfig.Text)
            NeedsToBeSaved = False
            lblMsg.Text = "Successfully saved Assignment Config File."
        Else
            lblMsg.Text = "Assignment Config File NOT saved. You need to specify a target filename on the first tab page."

        End If

    End Sub

    Private Sub SaveAsDefaultToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAsDefaultToolStripMenuItem.Click, btnSaveAsDefault.Click
        SaveConfigFile(Application.StartupPath & "\templates\defaultConfig.cfg")

        lblMsg.Text = "Saved default Configuration File."
    End Sub



    Sub SaveConfigFile(fn As String)

        Dim sw As New StreamWriter(fn)
        Dim s As String

        sw.WriteLine("AppGrader Assignment Configuration File")

        sw.WriteLine("CfgLanguage" & vbTab & "VB")
        sw.WriteLine("cfgAssignmentTitle" & vbTab & "")




        '  sw.WriteLine("CfgPath1" & vbTab & "MyDocuments")
        ' -------------------------------------------------------------------------------------------

        'For Each arow As DataGridViewRow In dgvAdvanced.Rows     ' this needs to be the first page
        '    s = arow.Cells(0).ToString & vbTab
        '    s = arow.Cells(1).ToString & vbTab
        '    s = arow.Cells(2).ToString & vbTab
        '    s = arow.Cells(3).ToString & vbTab
        '    s = arow.Cells(4).ToString
        '    sw.WriteLine(s)
        'Next

        For Each arow As DataGridViewRow In dgvDevEnvironment.Rows
            s = "dgvDevEnvironment" & vbTab
            s &= arow.Cells(0).Value.ToString & vbTab
            s &= arow.Cells(1).Value.ToString & vbTab
            s &= arow.Cells(2).Value.ToString & vbTab
            s &= arow.Cells(3).Value.ToString & vbTab
            s &= arow.Cells(4).Value.ToString & vbTab
            s &= arow.Cells(5).Value.ToString & vbTab
            s &= arow.Cells(6).Value.ToString & vbTab
            s &= arow.Cells(7).Value.ToString & vbTab

            '           s &= arow.Cells(8).Value.ToString
            sw.WriteLine(s)
        Next

        For Each arow As DataGridViewRow In dgvAppInfo.Rows
            s = "dgvAppInfo" & vbTab
            s &= arow.Cells(0).Value.ToString & vbTab
            s &= arow.Cells(1).Value.ToString & vbTab
            s &= arow.Cells(2).Value.ToString & vbTab
            s &= arow.Cells(3).Value.ToString & vbTab
            s &= arow.Cells(4).Value.ToString & vbTab
            s &= arow.Cells(5).Value.ToString & vbTab
            s &= arow.Cells(6).Value.ToString & vbTab
            s &= arow.Cells(7).Value.ToString & vbTab
            '           s &= arow.Cells(8).Value.ToString
            sw.WriteLine(s)
        Next

        For Each arow As DataGridViewRow In dgvCompileOptions.Rows
            s = "dgvCompileOptions" & vbTab
            s &= arow.Cells(0).Value.ToString & vbTab
            s &= arow.Cells(1).Value.ToString & vbTab
            s &= arow.Cells(2).Value.ToString & vbTab
            s &= arow.Cells(3).Value.ToString & vbTab
            s &= arow.Cells(4).Value.ToString & vbTab
            s &= arow.Cells(5).Value.ToString & vbTab
            s &= arow.Cells(6).Value.ToString & vbTab
            s &= arow.Cells(7).Value.ToString & vbTab
            '            s &= arow.Cells(8).Value.ToString
            sw.WriteLine(s)
        Next

        For Each arow As DataGridViewRow In dgvComments.Rows
            s = "dgvComments" & vbTab
            s &= arow.Cells(0).Value.ToString & vbTab
            s &= arow.Cells(1).Value.ToString & vbTab
            s &= arow.Cells(2).Value.ToString & vbTab
            s &= arow.Cells(3).Value.ToString & vbTab
            s &= arow.Cells(4).Value.ToString & vbTab
            s &= arow.Cells(5).Value.ToString & vbTab
            s &= arow.Cells(6).Value.ToString & vbTab
            s &= arow.Cells(7).Value.ToString & vbTab
            '           s &= arow.Cells(8).Value.ToString
            sw.WriteLine(s)
        Next

        For Each arow As DataGridViewRow In dgvFormDesign.Rows
            s = "dgvFormDesign" & vbTab
            s &= arow.Cells(0).Value.ToString & vbTab
            s &= arow.Cells(1).Value.ToString & vbTab
            s &= arow.Cells(2).Value.ToString & vbTab
            s &= arow.Cells(3).Value.ToString & vbTab
            s &= arow.Cells(4).Value.ToString & vbTab
            s &= arow.Cells(5).Value.ToString & vbTab
            s &= arow.Cells(6).Value.ToString & vbTab
            s &= arow.Cells(7).Value.ToString & vbTab
            '          s &= arow.Cells(8).Value.ToString
            sw.WriteLine(s)
        Next

        For Each arow As DataGridViewRow In dgvFormObj.Rows
            s = "dgvFormObj" & vbTab
            s &= arow.Cells(0).Value.ToString & vbTab
            s &= arow.Cells(1).Value.ToString & vbTab
            s &= arow.Cells(2).Value.ToString & vbTab
            s &= arow.Cells(3).Value.ToString & vbTab
            s &= arow.Cells(4).Value.ToString & vbTab
            s &= arow.Cells(5).Value.ToString & vbTab
            s &= arow.Cells(6).Value.ToString & vbTab
            s &= arow.Cells(7).Value.ToString & vbTab
            '          s &= arow.Cells(8).Value.ToString
            sw.WriteLine(s)
        Next

        For Each arow As DataGridViewRow In dgvImports.Rows
            s = "dgvImports" & vbTab
            s &= arow.Cells(0).Value.ToString & vbTab
            s &= arow.Cells(1).Value.ToString & vbTab
            s &= arow.Cells(2).Value.ToString & vbTab
            s &= arow.Cells(3).Value.ToString & vbTab
            s &= arow.Cells(4).Value.ToString & vbTab
            s &= arow.Cells(5).Value.ToString & vbTab
            s &= arow.Cells(6).Value.ToString & vbTab
            s &= arow.Cells(7).Value.ToString & vbTab
            '         s &= arow.Cells(8).Value.ToString
            sw.WriteLine(s)
        Next
        For Each arow As DataGridViewRow In dgvDataStructures.Rows
            s = "dgvDataStructures" & vbTab
            s &= arow.Cells(0).Value.ToString & vbTab
            s &= arow.Cells(1).Value.ToString & vbTab
            s &= arow.Cells(2).Value.ToString & vbTab
            s &= arow.Cells(3).Value.ToString & vbTab
            s &= arow.Cells(4).Value.ToString & vbTab
            s &= arow.Cells(5).Value.ToString & vbTab
            s &= arow.Cells(6).Value.ToString & vbTab
            s &= arow.Cells(7).Value.ToString & vbTab
            '         s &= arow.Cells(8).Value.ToString
            sw.WriteLine(s)
        Next

        For Each arow As DataGridViewRow In dgvDataTypes.Rows
            s = "dgvDataTypes" & vbTab
            s &= arow.Cells(0).Value.ToString & vbTab
            s &= arow.Cells(1).Value.ToString & vbTab
            s &= arow.Cells(2).Value.ToString & vbTab
            s &= arow.Cells(3).Value.ToString & vbTab
            s &= arow.Cells(4).Value.ToString & vbTab
            s &= arow.Cells(5).Value.ToString & vbTab
            s &= arow.Cells(6).Value.ToString & vbTab
            s &= arow.Cells(7).Value.ToString & vbTab
            '        s &= arow.Cells(8).Value.ToString
            sw.WriteLine(s)
        Next

        For Each arow As DataGridViewRow In dgvCoding.Rows
            s = "dgvCoding" & vbTab
            s &= arow.Cells(0).Value.ToString & vbTab
            s &= arow.Cells(1).Value.ToString & vbTab
            s &= arow.Cells(2).Value.ToString & vbTab
            s &= arow.Cells(3).Value.ToString & vbTab
            s &= arow.Cells(4).Value.ToString & vbTab
            s &= arow.Cells(5).Value.ToString & vbTab
            s &= arow.Cells(6).Value.ToString & vbTab
            s &= arow.Cells(7).Value.ToString & vbTab
            '      s &= arow.Cells(8).Value.ToString
            sw.WriteLine(s)
        Next

        For Each arow As DataGridViewRow In dgvSubs.Rows
            s = "dgvSubs" & vbTab
            s &= arow.Cells(0).Value.ToString & vbTab
            s &= arow.Cells(1).Value.ToString & vbTab
            s &= arow.Cells(2).Value.ToString & vbTab
            s &= arow.Cells(3).Value.ToString & vbTab
            s &= arow.Cells(4).Value.ToString & vbTab
            s &= arow.Cells(5).Value.ToString & vbTab
            s &= arow.Cells(6).Value.ToString & vbTab
            s &= arow.Cells(7).Value.ToString & vbTab
            '       s &= arow.Cells(8).Value.ToString
            sw.WriteLine(s)
        Next


        sw.Close()

        lblMsg.Text = "Configuration Saved."
        lblMsg.Visible = True
    End Sub

    Public Sub LoadConfigFile(fn As String)
        Dim s As String
        Dim ss() As String
        Dim sr As StreamReader
        Dim i As Integer
        If File.Exists(fn) Then
            sr = File.OpenText(fn)
            s = sr.ReadLine
            If s <> "AppGrader Assignment Configuration File" Then
                MessageBox.Show("The selected file (" & fn & ") is not a valid configuration file for this application. Either load a different file, or use the default settings.")
            Else
                Settings.Settings.Clear()
                Do While sr.Peek <> -1
                    s = sr.ReadLine
                    ss = s.Split(CChar(vbTab))
                    If ss.GetUpperBound(0) < 8 Then
                        ReDim Preserve ss(8)
                    End If

                    Try
                        Select Case ss(0).ToUpper                                ' Load the information into the datagridview
                            Case "CFGLANGUAGE"
                                cbxLanguage.Text = ss(1)

                            Case "CFGASSIGNMENTTITLE"
                                txtAssignmentTitle.Text = ss(1)
                                cfgAssignmentTitle = ss(1)

                            Case "DGVDEVENVIRONMENT"
                                ' Load the information into the datagridview
                                With dgvDevEnvironment
                                    For i = 0 To .RowCount - 1
                                        If .Rows(i).Cells(0).Value.ToString = ss(1) Then
                                            .Rows(i).Cells(1).Value = ss(2)
                                            .Rows(i).Cells(2).Value = ss(3)
                                            .Rows(i).Cells(3).Value = CBool(ss(4))
                                            .Rows(i).Cells(4).Value = CBool(ss(5))
                                            .Rows(i).Cells(5).Value = CInt(ss(6))
                                            .Rows(i).Cells(6).Value = CInt(ss(7))
                                            .Rows(i).Cells(7).Value = ss(8)
                                        End If
                                    Next i
                                    StuffSettings(ss)
                                End With
                            Case "DGVAPPINFO"
                                ' Load the information into the datagridview
                                With dgvAppInfo
                                    For i = 0 To .RowCount - 1
                                        If .Rows(i).Cells(0).Value.ToString = ss(1) Then
                                            .Rows(i).Cells(1).Value = ss(2)
                                            .Rows(i).Cells(2).Value = ss(3)
                                            .Rows(i).Cells(3).Value = CBool(ss(4))
                                            .Rows(i).Cells(4).Value = CBool(ss(5))
                                            .Rows(i).Cells(5).Value = CInt(ss(6))
                                            .Rows(i).Cells(6).Value = CInt(ss(7))
                                            .Rows(i).Cells(7).Value = ss(8)
                                        End If
                                    Next i
                                    StuffSettings(ss)
                                End With

                            Case "DGVCOMPILEOPTIONS"
                                ' Load the information into the datagridview
                                With dgvCompileOptions
                                    For i = 0 To .RowCount - 1
                                        If .Rows(i).Cells(0).Value.ToString = ss(1) Then
                                            .Rows(i).Cells(1).Value = ss(2)
                                            .Rows(i).Cells(2).Value = ss(3)
                                            .Rows(i).Cells(3).Value = CBool(ss(4))
                                            .Rows(i).Cells(4).Value = CBool(ss(5))
                                            .Rows(i).Cells(5).Value = CInt(ss(6))
                                            .Rows(i).Cells(6).Value = CInt(ss(7))
                                            .Rows(i).Cells(7).Value = ss(8)
                                        End If
                                    Next i
                                    StuffSettings(ss)
                                End With

                            Case "DGVCOMMENTS"
                                ' Load the information into the datagridview
                                With dgvComments
                                    For i = 0 To .RowCount - 1
                                        If .Rows(i).Cells(0).Value.ToString = ss(1) Then
                                            .Rows(i).Cells(1).Value = ss(2)
                                            .Rows(i).Cells(2).Value = ss(3)
                                            .Rows(i).Cells(3).Value = CBool(ss(4))
                                            .Rows(i).Cells(4).Value = CBool(ss(5))
                                            .Rows(i).Cells(5).Value = CInt(ss(6))
                                            .Rows(i).Cells(6).Value = CInt(ss(7))
                                            .Rows(i).Cells(7).Value = ss(8)
                                        End If
                                    Next i
                                    StuffSettings(ss)
                                End With

                            Case "DGVFORMDESIGN"
                                ' Load the information into the datagridview
                                With dgvFormDesign
                                    For i = 0 To .RowCount - 1
                                        If .Rows(i).Cells(0).Value.ToString = ss(1) Then
                                            .Rows(i).Cells(1).Value = ss(2)
                                            .Rows(i).Cells(2).Value = ss(3)
                                            .Rows(i).Cells(3).Value = CBool(ss(4))
                                            .Rows(i).Cells(4).Value = CBool(ss(5))
                                            .Rows(i).Cells(5).Value = CInt(ss(6))
                                            .Rows(i).Cells(6).Value = CInt(ss(7))
                                            .Rows(i).Cells(7).Value = ss(8)
                                        End If
                                    Next i
                                    StuffSettings(ss)
                                End With

                            Case "DGVFORMOBJ"
                                ' Load the information into the datagridview
                                With dgvFormObj
                                    For i = 0 To .RowCount - 1
                                        If .Rows(i).Cells(0).Value.ToString = ss(1) Then
                                            .Rows(i).Cells(1).Value = ss(2)
                                            .Rows(i).Cells(2).Value = ss(3)
                                            .Rows(i).Cells(3).Value = CBool(ss(4))
                                            .Rows(i).Cells(4).Value = CBool(ss(5))
                                            .Rows(i).Cells(5).Value = CInt(ss(6))
                                            .Rows(i).Cells(6).Value = CInt(ss(7))
                                            .Rows(i).Cells(7).Value = ss(8)
                                        End If
                                    Next i
                                    StuffSettings(ss)
                                End With

                            Case "DGVIMPORTS"
                                ' Load the information into the datagridview
                                With dgvImports
                                    For i = 0 To dgvImports.RowCount - 1
                                        If .Rows(i).Cells(0).Value.ToString = ss(1) Then
                                            .Rows(i).Cells(1).Value = ss(2)
                                            .Rows(i).Cells(2).Value = ss(3)
                                            .Rows(i).Cells(3).Value = CBool(ss(4))
                                            .Rows(i).Cells(4).Value = CBool(ss(5))
                                            .Rows(i).Cells(5).Value = CInt(ss(6))
                                            .Rows(i).Cells(6).Value = CInt(ss(7))
                                            .Rows(i).Cells(7).Value = ss(8)
                                        End If
                                    Next i
                                    StuffSettings(ss)
                                End With

                            Case "DGVDATASTRUCTURES"
                                ' Load the information into the datagridview
                                With dgvDataStructures
                                    For i = 0 To dgvDataStructures.RowCount - 1
                                        If .Rows(i).Cells(0).Value.ToString = ss(1) Then
                                            .Rows(i).Cells(1).Value = ss(2)
                                            .Rows(i).Cells(2).Value = ss(3)
                                            .Rows(i).Cells(3).Value = CBool(ss(4))
                                            .Rows(i).Cells(4).Value = CBool(ss(5))
                                            .Rows(i).Cells(5).Value = CInt(ss(6))
                                            .Rows(i).Cells(6).Value = CInt(ss(7))
                                            .Rows(i).Cells(7).Value = ss(8)
                                        End If
                                    Next i
                                    StuffSettings(ss)
                                End With

                            Case "DGVDATATYPES"
                                ' Load the information into the datagridview
                                With dgvDataTypes
                                    For i = 0 To dgvDataTypes.RowCount - 1
                                        If .Rows(i).Cells(0).Value.ToString = ss(1) Then
                                            .Rows(i).Cells(1).Value = ss(2)
                                            .Rows(i).Cells(2).Value = ss(3)
                                            .Rows(i).Cells(3).Value = CBool(ss(4))
                                            .Rows(i).Cells(4).Value = CBool(ss(5))
                                            .Rows(i).Cells(5).Value = CInt(ss(6))
                                            .Rows(i).Cells(6).Value = CInt(ss(7))
                                            .Rows(i).Cells(7).Value = ss(8)
                                        End If
                                    Next i
                                    StuffSettings(ss)
                                End With

                            Case "DGVCODING"
                                ' Load the information into the datagridview
                                With dgvCoding
                                    For i = 0 To dgvCoding.RowCount - 1
                                        If .Rows(i).Cells(0).Value.ToString = ss(1) Then
                                            .Rows(i).Cells(1).Value = ss(2)
                                            .Rows(i).Cells(2).Value = ss(3)
                                            .Rows(i).Cells(3).Value = CBool(ss(4))
                                            .Rows(i).Cells(4).Value = CBool(ss(5))
                                            .Rows(i).Cells(5).Value = CInt(ss(6))
                                            .Rows(i).Cells(6).Value = CInt(ss(7))
                                            .Rows(i).Cells(7).Value = ss(8)
                                        End If
                                    Next i
                                    StuffSettings(ss)
                                End With

                            Case "DGVSUBS"
                                ' Load the information into the datagridview
                                With dgvSubs
                                    For i = 0 To dgvSubs.RowCount - 1
                                        If .Rows(i).Cells(0).Value.ToString = ss(1) Then
                                            .Rows(i).Cells(1).Value = ss(2)
                                            .Rows(i).Cells(2).Value = ss(3)
                                            .Rows(i).Cells(3).Value = CBool(ss(4))
                                            .Rows(i).Cells(4).Value = CBool(ss(5))
                                            .Rows(i).Cells(5).Value = CInt(ss(6))
                                            .Rows(i).Cells(6).Value = CInt(ss(7))
                                            .Rows(i).Cells(7).Value = ss(8)
                                        End If
                                    Next i
                                    StuffSettings(ss)
                                End With

                        End Select

                    Catch Ex As Exception
                        MessageBox.Show(Ex.Message)
                    End Try
                Loop
            End If
            sr.Close()
        End If

    End Sub

    Sub StuffSettings(ss() As String)
        Dim ns As New MySettings
        Try
            ns.DVG = ss(0)
            ns.nm = ss(1)
            ns.Name = ss(1)
            ns.Text = ss(2)

            If ss(4) = "" Then
                ns.ShowVar = False
            Else
                ns.ShowVar = CBool(ss(4))
            End If
            If ss(5) = "" Then
                ns.Req = False
            Else
                ns.Req = CBool(ss(5))
            End If


            ns.PtsPerError = CDec(ss(6))
            ns.MaxPts = CDec(ss(7))
            ns.Feedback = ss(8)

            Settings.Settings.Add(ns)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        '       sw.WriteLine(ns.DVG & vbTab & ns.Name & vbTab & ns.ShowVar & vbTab & ns.Req & vbTab & ns.PtsPerError & vbTab & ns.MaxPts & vbTab & ns.Feedback)
    End Sub


    Private Sub LoadDefaultToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoadDefaultToolStripMenuItem.Click
        ' loads in the default config file. It stuffs the app settings, and also the dgvs

        lblAssessmentFile.Text = Application.StartupPath & "\templates\defaultConfig.cfg"
        LoadCfgFile(Application.StartupPath & "\templates\defaultConfig.cfg")          ' this reads the config file and stuffs the data into the settings.
        LoadConfigFile(Application.StartupPath & "\templates\defaultConfig.cfg")       ' this read in the title & language, and stuffs the dgvs.

        lblMsg.Text = "Default configuration file Loaded."
    End Sub

    Private Sub LoadConfigToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoadConfigToolStripMenuItem.Click
        ' loads a user specified config file. It stuffs the app settings, and also the dgvs
        OpenFileDialog1.Filter = "Config Files|*.cfg|All Files|*.*"
        OpenFileDialog1.ShowDialog()

        If OpenFileDialog1.FileName <> Nothing Then
            LoadCfgFile(OpenFileDialog1.FileName)
            LoadConfigFile(OpenFileDialog1.FileName)
        End If

        lblMsg.Text = OpenFileDialog1.FileName & " config file Loaded."

    End Sub

    Private Sub tbnToggle1_Click(sender As Object, e As EventArgs) Handles tbnToggle1.Click
        Dim i As Integer

        With dgvDevEnvironment
            If .Text = "Toggle All Property Show On" Then
                .Text = "Toggle All Property Show Off"
                For i = 0 To .RowCount - 1
                    .Rows(i).Cells(4).Value = False.ToString
                Next
            Else
                .Text = "Toggle All Property Show On"
                For i = 0 To .RowCount - 1
                    .Rows(i).Cells(4).Value = True.ToString
                Next
            End If
        End With
    End Sub

    Private Sub btnToggle2_Click(sender As Object, e As EventArgs) Handles btnToggle2.Click
        Dim i As Integer

        With dgvAppInfo
            If .Text = "Toggle All Property Show On" Then
                .Text = "Toggle All Property Show Off"
                For i = 0 To .RowCount - 1
                    .Rows(i).Cells(4).Value = False.ToString
                Next
            Else
                .Text = "Toggle All Property Show On"
                For i = 0 To .RowCount - 1
                    .Rows(i).Cells(4).Value = True.ToString
                Next
            End If
        End With
    End Sub

    Private Sub btnToggle3_Click(sender As Object, e As EventArgs) Handles btnToggle3.Click
        Dim i As Integer

        With dgvCompileOptions
            If .Text = "Toggle All Property Show On" Then
                .Text = "Toggle All Property Show Off"
                For i = 0 To .RowCount - 1
                    .Rows(i).Cells(4).Value = False.ToString
                Next
            Else
                .Text = "Toggle All Property Show On"
                For i = 0 To .RowCount - 1
                    .Rows(i).Cells(4).Value = True.ToString
                Next
            End If
        End With
    End Sub

    Private Sub btnToggle4_Click(sender As Object, e As EventArgs) Handles btnToggle4.Click
        Dim i As Integer

        With dgvComments
            If .Text = "Toggle All Property Show On" Then
                .Text = "Toggle All Property Show Off"
                For i = 0 To .RowCount - 1
                    .Rows(i).Cells(4).Value = False.ToString
                Next
            Else
                .Text = "Toggle All Property Show On"
                For i = 0 To .RowCount - 1
                    .Rows(i).Cells(4).Value = True.ToString
                Next
            End If
        End With
    End Sub

    Private Sub btnToggle5_Click(sender As Object, e As EventArgs) Handles btnToggle5.Click
        Dim i As Integer

        With dgvFormDesign
            If .Text = "Toggle All Property Show On" Then
                .Text = "Toggle All Property Show Off"
                For i = 0 To .RowCount - 1
                    .Rows(i).Cells(4).Value = False.ToString
                Next
            Else
                .Text = "Toggle All Property Show On"
                For i = 0 To .RowCount - 1
                    .Rows(i).Cells(4).Value = True.ToString
                Next
            End If
        End With
    End Sub

    Private Sub btnToggle6_Click(sender As Object, e As EventArgs) Handles btnToggle6.Click
        Dim i As Integer

        With dgvImports
            If .Text = "Toggle All Property Show On" Then
                .Text = "Toggle All Property Show Off"
                For i = 0 To .RowCount - 1
                    .Rows(i).Cells(4).Value = False.ToString
                Next
            Else
                .Text = "Toggle All Property Show On"
                For i = 0 To .RowCount - 1
                    .Rows(i).Cells(4).Value = True.ToString
                Next
            End If
        End With
    End Sub

    Private Sub btnToggle7_Click(sender As Object, e As EventArgs) Handles btnToggle7.Click
        Dim i As Integer

        With dgvDataStructures
            If .Text = "Toggle All Property Show On" Then
                .Text = "Toggle All Property Show Off"
                For i = 0 To .RowCount - 1
                    .Rows(i).Cells(4).Value = False.ToString
                Next
            Else
                .Text = "Toggle All Property Show On"
                For i = 0 To .RowCount - 1
                    .Rows(i).Cells(4).Value = True.ToString
                Next
            End If
        End With
    End Sub
   
    Private Sub btnToggle8_Click(sender As Object, e As EventArgs) Handles btnToggle8.Click
        Dim i As Integer

        With dgvDataTypes
            If .Text = "Toggle All Property Show On" Then
                .Text = "Toggle All Property Show Off"
                For i = 0 To .RowCount - 1
                    .Rows(i).Cells(4).Value = False.ToString
                Next
            Else
                .Text = "Toggle All Property Show On"
                For i = 0 To .RowCount - 1
                    .Rows(i).Cells(4).Value = True.ToString
                Next
            End If
        End With
    End Sub

    Private Sub btnToggle9_Click(sender As Object, e As EventArgs) Handles btnToggle9.Click
        Dim i As Integer

        With dgvCoding
            If .Text = "Toggle All Property Show On" Then
                .Text = "Toggle All Property Show Off"
                For i = 0 To .RowCount - 1
                    .Rows(i).Cells(4).Value = False.ToString
                Next
            Else
                .Text = "Toggle All Property Show On"
                For i = 0 To .RowCount - 1
                    .Rows(i).Cells(4).Value = True.ToString
                Next
            End If
        End With
    End Sub

    Private Sub btnToggle10_Click(sender As Object, e As EventArgs) Handles btnToggle10.Click
        Dim i As Integer

        With dgvSubs
            If .Text = "Toggle All Property Show On" Then
                .Text = "Toggle All Property Show Off"
                For i = 0 To .RowCount - 1
                    .Rows(i).Cells(4).Value = False.ToString
                Next
            Else
                .Text = "Toggle All Property Show On"
                For i = 0 To .RowCount - 1
                    .Rows(i).Cells(4).Value = True.ToString
                Next
            End If
        End With
    End Sub

    Private Sub dgvDevEnvironment_RowEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgvDevEnvironment.RowEnter, dgvAppInfo.RowEnter, dgvCompileOptions.RowEnter, dgvComments.RowEnter, dgvFormDesign.RowEnter, dgvFormObj.RowEnter, dgvImports.RowEnter, dgvDataStructures.RowEnter, dgvDataTypes.RowEnter, dgvCoding.RowEnter, dgvSubs.RowEnter

        If Not loadphase Then
            dgv = CType(sender, DataGridView)

            If dgv.Rows.Count <> Nothing Then
                previousRow = e.RowIndex

                txtDevEnv.Text = dgv.Rows(e.RowIndex).Cells(7).Value.ToString
                Select Case TabControl1.SelectedIndex
                    Case 1
                        txtDevEnv.Text = dgv.Rows(e.RowIndex).Cells(7).Value.ToString
                    Case 2
                        txtAppInfo.Text = dgv.Rows(e.RowIndex).Cells(7).Value.ToString
                    Case 3
                        txtCompileOpt.Text = dgv.Rows(e.RowIndex).Cells(7).Value.ToString
                    Case 4
                        txtComments.Text = dgv.Rows(e.RowIndex).Cells(7).Value.ToString
                    Case 5
                        txtFormDesign.Text = dgv.Rows(e.RowIndex).Cells(7).Value.ToString
                    Case 6
                        txtFormObj.Text = dgv.Rows(e.RowIndex).Cells(7).Value.ToString
                    Case 7
                        txtImports.Text = dgv.Rows(e.RowIndex).Cells(7).Value.ToString
                    Case 8
                        txtDataStructures.Text = dgv.Rows(e.RowIndex).Cells(7).Value.ToString
                    Case 9
                        txtDataTypes.Text = dgv.Rows(e.RowIndex).Cells(7).Value.ToString
                    Case 10
                        txtCoding.Text = dgv.Rows(e.RowIndex).Cells(7).Value.ToString
                    Case 11
                        txtSubs.Text = dgv.Rows(e.RowIndex).Cells(7).Value.ToString
                End Select

            Else
                txtDevEnv.Text = ""
            End If
        End If

    End Sub



    Sub txtDevEnv_Leave(sender As Object, e As EventArgs) Handles txtDevEnv.Leave
        dgv.Rows(previousRow).Cells(7).Value = txtDevEnv.Text
    End Sub
    Sub txtAppInfo_Leave(sender As Object, e As EventArgs) Handles txtAppInfo.Leave
        dgv.Rows(previousRow).Cells(7).Value = txtAppInfo.Text
    End Sub
    Sub txtCompileOpt_Leave(sender As Object, e As EventArgs) Handles txtCompileOpt.Leave
        dgv.Rows(previousRow).Cells(7).Value = txtCompileOpt.Text
    End Sub
    Sub txtComments_Leave(sender As Object, e As EventArgs) Handles txtComments.Leave
        dgv.Rows(previousRow).Cells(7).Value = txtComments.Text
    End Sub
    Sub txtFormDesign_Leave(sender As Object, e As EventArgs) Handles txtFormDesign.Leave
        dgv.Rows(previousRow).Cells(7).Value = txtFormDesign.Text
    End Sub
    Sub txtFormObj_Leave(sender As Object, e As EventArgs) Handles txtFormObj.Leave
        dgv.Rows(previousRow).Cells(7).Value = txtFormObj.Text
    End Sub
    Sub txtImports_Leave(sender As Object, e As EventArgs) Handles txtImports.Leave
        dgv.Rows(previousRow).Cells(7).Value = txtImports.Text
    End Sub
    Sub txtDataStructures_Leave(sender As Object, e As EventArgs) Handles txtDataStructures.Leave
        dgv.Rows(previousRow).Cells(7).Value = txtDataStructures.Text
    End Sub
    Sub txtDataTypes_Leave(sender As Object, e As EventArgs) Handles txtDataTypes.Leave
        dgv.Rows(previousRow).Cells(7).Value = txtDataTypes.Text
    End Sub
    Sub txtCoding_Leave(sender As Object, e As EventArgs) Handles txtCoding.Leave
        dgv.Rows(previousRow).Cells(7).Value = txtCoding.Text
    End Sub
    Sub txtSubs_Leave(sender As Object, e As EventArgs) Handles txtSubs.Leave
        dgv.Rows(previousRow).Cells(7).Value = txtSubs.Text
    End Sub



    Sub tabcontrol1_SelectedIntexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        Dim tbx As TextBox

        Select Case TabControl1.SelectedIndex
            Case 1
                dgv = dgvDevEnvironment
                tbx = txtDevEnv
                ' need to set the opening cell to 0 
            Case 2
                dgv = dgvAppInfo
                tbx = txtAppInfo
            Case 3
                dgv = dgvCompileOptions
                tbx = txtCompileOpt
            Case 4
                dgv = dgvComments
                tbx = txtComments
            Case 5
                dgv = dgvFormDesign
                tbx = txtFormDesign
            Case 6
                dgv = dgvFormObj
                tbx = txtFormObj
            Case 7
                dgv = dgvImports
                tbx = txtImports
            Case 8
                dgv = dgvDataStructures
                tbx = txtDataStructures
            Case 9
                dgv = dgvDataTypes
                tbx = txtDataTypes
            Case 10
                dgv = dgvCoding
                tbx = txtCoding
            Case 11
                dgv = dgvSubs
                tbx = txtSubs
        End Select

        Try
            If TabControl1.SelectedIndex <> 0 Then
                If dgv.CurrentCell.RowIndex = Nothing Then Me.dgv.CurrentCell = Me.dgv(1, 1)
                tbx.Text = dgv.Rows(dgv.CurrentCell.RowIndex).Cells(7).Value.ToString
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub

 
  
    Private Sub RestoreStandardSettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RestoreStandardSettingsToolStripMenuItem.Click
        ' This loads the factory settings. There is no 
        lblAssessmentFile.Text = Application.StartupPath & "\templates\FactoryConfig.cfg"
        LoadCfgFile(Application.StartupPath & "\templates\FactoryConfig.cfg")
        LoadConfigFile(Application.StartupPath & "\templates\FactoryConfig.cfg")
        lblMsg.Text = "Restored Factory Configuration."
    End Sub

    Private Sub dgv_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgvAppInfo.CellEndEdit, dgvCoding.CellEndEdit, dgvComments.CellEndEdit, dgvCompileOptions.CellEndEdit, dgvDataStructures.CellEndEdit, dgvDataTypes.CellEndEdit, dgvDevEnvironment.CellEndEdit, dgvFormDesign.CellEndEdit, dgvFormObj.CellEndEdit, dgvImports.CellEndEdit, dgvSubs.CellEndEdit

        ' If there are any changes in any dgv, then mark the data indicating it needs to be saved
        NeedsToBeSaved = True

        btnSaveAsDefault.Visible = True
        btnSaveConfig.Visible = True

    End Sub

    Private Sub btnSelectConfigLoc_Click(sender As Object, e As EventArgs) Handles btnSelectConfigLoc.Click
        SaveFileDialog1.ShowDialog()

        If SaveFileDialog1.FileName <> "" Then
            If SaveFileDialog1.FileName.ToLower.Contains(".cfg") Then
                lblTargetConfig.Text = SaveFileDialog1.FileName
            Else
                MessageBox.Show("The specified filename needs to have a .cfg extension. Please specify an appropriate filename.", "Improper Filename")
            End If
        End If
    End Sub
End Class