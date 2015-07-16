﻿Module Assignment

    Public Enum dgvs
        ApplicationSettings = 1
        SystemVariables = 2
        LogicVariables = 3
        Splash = 4
        AdvancedVariables = 5
        FormProperties = 6
    End Enum
    ' ======================================================================

    Public Enum EnForm
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
        FormLoadMethod = 22
    End Enum

    Public Const nForm As Integer = 22

    Public EnFormNames() As String = {"ObjFormPrefixes", "ObjButton", "objLabel", "ObjActiveLabel", "ObjNonactiveLabel", "ObjTextbox", "ObjListbox", "ObjCombobox", "ObjRadioButton", "ObjCheckbox", "ObjGroupBox", "ObjPanel", "ObjWebBrowser", "ObjWebClient", "ObjOpenFileDialog", "ObjSaveFileDialog", "FormName", "FormText", "FormBackColor", "FormAcceptButton", "FormCancelButton", "FormStartPosition", "FormLoadMethod"}

    Public Enum EnSummary
        StudentID = 0

        OptionStrict = 1
        OptionExplicit = 2

        hasSLN = 3
        hasVBproj = 4
        hasSplashScreen = 5
        hasAboutBox = 6

        InfoAppTitle = 7
        InfoDescription = 8
        InfoCompany = 9
        InfoProduct = 10
        InfoTrademark = 11
        InfoCopyright = 12
        InfoGUID = 13

        CommentGeneral = 14
        CommentSub = 15
        CommentIF = 16
        CommentFor = 17
        CommentDo = 18
        CommentWhile = 19
        CommentSelect = 20

        'dgvFormDesign
        RenameObjects = 21
        IncludeFrmInFormName = 22
        ChangeFormText = 23
        ChangeFormColor = 24
        SetFormAcceptButton = 25
        SetFormCancelButton = 26
        ModifyStartPosition = 27

        ' dgvImports
        SystemIO = 28
        SystemNet = 29
        SystemDB = 30

        VarArrays = 31
        VarLists = 32
        VarStructures = 33

        VarBoolean = 34
        VarInteger = 35
        VarDecimal = 36
        VarDate = 37
        VarString = 38
        VarPrefixes = 39

        LogicIF = 40
        LogicFor = 41
        LogicDo = 42
        LogicWhile = 43
        LogicElse = 44
        LogicElseIF = 45
        LogicMessageBox = 46
        LogicNestedIF = 47
        LogicNestedFor = 48
        LogicSelectCase = 49
        LogicConcatination = 50
        LogicConvertToString = 51
        LogicToStringFormat = 52
        LogicStringFormat = 53
        LogicStringFormatParameters = 54
        LogicComplexConditions = 55
        LogicCaseInsensitive = 56
        LogicTryCatch = 57
        LogicStreamReader = 58
        LogicStreamWriter = 59
        LogicStreamReaderClose = 60
        LogicStreamWriterClose = 61

        LogicCStr = 62
        LogicToString = 63

        LogicSub = 64
        LogicFunction = 65
        LogicOptional = 66
        LogicByRef = 67
        LogicMultipleForms = 68
        LogicModule = 69
        LogicFormLoad = 70

        LogicVarPrefixes = 71
        LogicFlowControl = 72
        TotalScore = 73
    End Enum

    Public Const NSummary As Integer = 73

    Public EnSummaryName() As String = {"StudentID", "OptionStrict", "OptionExplicit", "hasSLN", "hasVBproj", "hasSplashScreen", "hasAboutBox", "InfoAppTitle", "InfoDescription", "InfoCompany", "InfoProduct", "InfoTrademark", "InfoCopyright", "InfoGUID", "CommentGeneral", "CommentSub", "CommentIF", "CommentFor", "CommentDo", "CommentWhile", "CommentSelect", "RenameObjects", "IncludeFrmInFormName", "ChangeFormText", "ChangeFormColor", "SetFormAcceptButton", "SetFormCancelButton", "ModifyStartPosition", "SystemIO", "SystemNet", "SystemDBv", "VarArrays", "VarLists", "VarStructures", "VarBoolean", "VarInteger", "VarDecimal", "VarDate", "VarString", "VarPrefixes", "LogicIF", "LogicFor", "LogicDo", "LogicWhile", "LogicElse", "LogicElseIF", "LogicMessageBox", "LogicNestedIF", "LogicNestedFor", "LogicSelectCase", "LogicConcatination", "LogicConvertToString", "LogicToStringFormat", "LogicStringFormat", "LogicStringFormatParameters", "LogicComplexConditions", "LogicCaseInsensitive", "LogicTryCatch", "LogicStreamReader", "LogicStreamWriter", "LogicStreamReaderClose", "LogicStreamWriterClose", "LogicCStr", "LogicToString", "LogicSub", "LogicFunction", "LogicOptional", "LogicByRef", "LogicMultipleForms", "LogicModule", "LogicFormLoad", "LogicVarPrefixes", "LogicFlowControl", "TotalScore"}


    ' ========================================================================
    Public Structure MyItems
        ' each assessment item has a MyItems structure. Req, PtsperError, and PossiblePts are set by instructor to determine
        ' the grading for each assessment item per assignment.
        Dim req As Boolean
        Dim showVar As Boolean  ' if we want this, need to bring it in on the datagridview.
        Dim PtsPerError As Decimal
        Dim PossiblePts As Decimal

        Dim Status As String          ' holds the main strings associated with the item
        Dim cnt As Integer
        Dim n As Integer              ' number of instances found
        Dim cssClass As String
        Dim cssNonChk As String       ' can be either hidden, gray or white(none)
        Dim bad As String
        Dim good As String

        Dim BlockID As Integer
        Dim YourPts As Decimal
        Dim Comments As String
        Dim isBad As Boolean
        Dim ID As String
    End Structure
    ' ========================================================================

    Public Structure AssignmentInfo
        Dim StudentID As String
        Dim AppTitle As String
        Dim AssignRoot As String
        Dim AssignPath As String    ' This is specific to the student
        Dim CompileDate As String
        ' --------------------------
        Dim TotalScore As Decimal
        Dim strTotalScore As String
        ' --------------------------
        Dim OptionStrict As MyItems
        Dim OptionExplicit As MyItems

        Dim hasSLN As MyItems
        Dim VBVersion As MyItems
        Dim hasVBproj As MyItems
        Dim hasSplashScreen As MyItems
        Dim hasAboutBox As MyItems
        Dim Modules As MyItems    ' ????????????????????
    End Structure

    Structure MyItems1
        Dim ID As Integer
        Dim Name As String
        Dim dgv As Integer
    End Structure


    '    Public Items1 As New List(Of MyItems)
 

    '    Public myindex As Integer


    Public Structure MyErrorComments
        Dim topic As String
        Dim Comment As String
    End Structure


    Public strStudentID As String
    Public strAssignmentSummary As String = ""
    '    Public EarliestPostDate As Date
    '    Public OutputFile As String = ""

    Public TotalLinesOfCode As Integer
    Public FileLinesOfCode As Integer
    Public TotalPossiblePts As Decimal
    Public TotalScore As Decimal
    Public SubmissionCompileTime As String = ""
    Public SubmissionCompileDate As String = ""

    Public bullet As String = Chr(149) & " "


    ' Config Settings
    '    Public CfgLanguage As String = "VB"
    Public cfgAssignmentTitle As String = ""

    '    Public CfgPath1 As String = "MyDocuments"
    Public AllowOverwrite As Boolean = False
    Public strOutputPath As String = ""     ' this is the root for the whole assignment 
    Public strStudentRoot As String = ""
    Public strStudentPath As String = ""
    '    Public strProjectFile As String = ""
    '    Public strProjectName As String = ""

    ' ==========================================================

    Public ErrorComments As New List(Of ErrComments)
    Public GuidIssues As Boolean = False
    Public CRCIssues As Boolean = False

    Public StudentReportPath As String = ""

    '  Public chkCommentAllVars As Boolean = True
    Public pbar3max As Integer = 100
    Public HideGray As String = "Hide"
    ' ===========================================================================================

    '  Public AppSummary(80) As MyItems  ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    '  Public AppForm(80) As MyItems



    ' Application variables - there is only one of each of these per app.
    '     Dim BreakPoints As myitems
    '     Dim WatchVariables As myitems
    ' =========================================================================================================

    Public Sub ClearMyItems(a As MyItems)
        a.req = False
        a.showVar = False            ' if we want this, need to bring it in on the datagridview.
        a.PtsPerError = 0
        a.PossiblePts = 0

        a.Status = Nothing          ' holds the main strings associated with the item
        a.cnt = 0
        a.n = 0                     ' number of instances found
        a.cssClass = Nothing
        a.cssNonChk = Nothing       ' can be either hidden, gray or white(none)
        a.bad = Nothing
        a.good = Nothing

        a.BlockID = 0
        a.YourPts = 0
        a.Comments = Nothing
        a.isBad = False
        a.ID = Nothing

    End Sub

    'Public Sub ClearAssignmentInfo(a As AssignmentInfo)
    '    a.StudentID = Nothing
    '    a.AppTitle = Nothing
    '    a.AssignRoot = Nothing
    '    a.AssignPath = Nothing    ' This is specific to the student
    '    a.CompileDate = Nothing
    '    ' --------------------------
    '    a.TotalScore = 0
    '    a.strTotalScore = Nothing
    '    ' --------------------------
    '    ClearMyItems(a.OptionStrict)
    '    ClearMyItems(a.OptionExplicit)

    '    ClearMyItems(a.hasSLN)
    '    ClearMyItems(a.VBVersion)
    '    ClearMyItems(a.hasVBproj)
    '    ClearMyItems(a.hasSplashScreen)
    '    ClearMyItems(a.hasAboutBox)
    '    ClearMyItems(a.Modules)    ' ????????????????????
    'End Sub

    Public Sub ClearAppArray(a() As MyItems)
        Dim i As Integer
        For i = 0 To a.GetUpperBound(0)
            ClearMyItems(a(i))
        Next i
    End Sub

    'Public Sub NewApplication(ByRef a As AssignmentInfo)
    '    Dim nc As String = ""                       ' Non-checked property
    '    Dim c As String = "ncWhite"    ' Checked property
    '    ' ----------------------------------------------------------
    '    If HideGray = "Gray" Then
    '        nc = "ncGray"
    '    ElseIf HideGray = "Hide" Then
    '        nc = "ncHide"
    '    Else
    '        nc = "ncWhite"
    '    End If

    '    ' ----------------------------------------------------------
    '    setchecked(Find_Setting("OptionStrict", "new").Req, a.OptionStrict, nc, c)
    '    setchecked(Find_Setting("OptionExplicit", "new").Req, a.OptionExplicit, nc, c)
    '    ' ----------------------------------------------------------
    '    setchecked(Find_Setting("hasSLN", "new").Req, a.hasSLN, nc, c)
    '    setchecked(Find_Setting("hasvbProj", "new").Req, a.hasVBproj, nc, c)
    '    setchecked(Find_Setting("hasSplashScreen", "new").Req, a.hasSplashScreen, nc, c)
    '    setchecked(Find_Setting("hasAboutBox", "new").Req, a.hasAboutBox, nc, c)
    '    '       setchecked(Find_Setting("Module","new").req, Modules, nc, c)

    'End Sub


    'Public Sub StuffAppData(ByRef a As AssignmentInfo, StudID As String, ApplicationName As String, AssignmentRoot As String, AssignmentPath As String, CompDate As Date, OptStrict As Boolean, optExplicit As Boolean, has_SLN As Boolean, VB_Version As String, has_VBProj As Boolean, has_SplashScreen As Boolean, has_AboutBox As Boolean, N_Modules As Integer)

    '    Try
    '        a.StudentID = StudID
    '        a.AppTitle = ApplicationName
    '        a.AssignRoot = AssignmentRoot
    '        a.AssignPath = AssignmentPath
    '        a.CompileDate = CompDate.ToString

    '        a.OptionStrict.Status = OptStrict.ToString
    '        a.OptionExplicit.Status = optExplicit.ToString

    '        a.hasSLN.Status = has_SLN.ToString
    '        a.VBVersion.Status = VB_Version
    '        a.hasVBproj.Status = has_VBProj.ToString
    '        a.hasSplashScreen.Status = has_SplashScreen.ToString
    '        a.hasAboutBox.Status = has_AboutBox.ToString
    '        a.Modules.Status = N_Modules.ToString
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message)
    '    End Try
    'End Sub

    ' ========================================================================

    'Public Sub NewMyItems(a() As MyItems)   ' AppSummary

    '    Dim i As Integer
    '    Dim nc As String = ""                       ' Non-checked property
    '    Dim c As String = "ncWhite"    ' Checked property

    '    ' ----------------------------------------------------------

    '    If HideGray = "Gray" Then
    '        nc = "ncGray"
    '    ElseIf HideGray = "Hide" Then
    '        nc = "ncHide"
    '    Else
    '        nc = "ncWhite"
    '    End If

    '    For i = 0 To a.GetUpperBound(0)
    '        setchecked(Find_Setting(EnSummaryName(i), "NewMyItems").Req, a(i), nc, c)
    '    Next i



    '    'If Not chkMultiForm,  MultiForm.cssHideNonChk = nchide
    '    'If Not chkLineNbr,  OptionStrict.cssHideNonChk = nchide 
    '    'If Not chkWordWrap,  OptionStrict.cssHideNonChk = nchide
    '    ' ----------------------------------------------------------
    '    'setchecked(Find_Setting("hasSLN").Req, Assign.hasSLN, nc, c)
    '    'setchecked(Find_Setting("hasvbProj").Req, Assign.hasVBproj, nc, c)
    '    'setchecked(Find_Setting("hasSplashScreen").Req, Assign.hasSplashScreen, nc, c)
    '    'setchecked(Find_Setting("hasAboutBox").Req, Assign.hasAboutBox, nc, c)
    '    'setchecked(Find_Setting("Include Module").Req, Assign.Modules, nc, c)
    '    'setchecked(Find_Setting("OptionStrict").Req, Assign.OptionStrict, nc, c)
    '    'setchecked(Find_Setting("OptionExplicit").Req, Assign.OptionExplicit, nc, c)
    '    'setchecked(Find_Setting("Include a Form LOAD Method").Req, LogicFormLoad, nc, c)
    '    ' ----------------------------------------------------------
    'End Sub

    'Public Sub setchecked(chk As Boolean, ByRef obj As MyItems, nc As String, c As String)
    '    If Not chk Then
    '        obj.cssNonChk = nc
    '        obj.req = False
    '    Else
    '        obj.cssNonChk = c
    '        obj.req = True
    '    End If
    'End Sub


    Sub integrateSSummary(AppSummary() As MyItems, IntegratedStudentAssignment() As MyItems, filename As String, first As Boolean)
        Dim i As Integer
        If first Then
            For i = 0 To AppSummary.GetUpperBound(0)
                IntegratedStudentAssignment(i) = AppSummary(i)
                IntegratedStudentAssignment(i).Status &= filename & " - " & AppSummary(i).n.ToString & vbCrLf
            Next
        Else
            For i = 0 To AppSummary.GetUpperBound(0)
                IntegratedStudentAssignment(i).n += AppSummary(i).n
            Next
        End If

    End Sub

End Module