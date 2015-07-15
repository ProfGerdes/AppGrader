Public Class Assignment

    Enum dgvs
        ApplicationSettings = 1
        SystemVariables = 2
        LogicVariables = 3
        Splash = 4
        AdvancedVariables = 5
        FormProperties = 6
    End Enum
    ' ======================================================================

    ' Application variables - there is only one of each of these per app.
    Public StudentID As String
    Public AppTitle As String
    Public AssignRoot As String
    Public AssignPath As String    ' This is specific to the student
    Public CompileDate As String
    ' --------------------------

    Public OptionStrict As New MyItems
    Public OptionExplicit As New MyItems

    Public hasSLN As New MyItems
    Public VBVersion As New MyItems
    Public hasVBproj As New MyItems
    Public hasSplashScreen As New MyItems
    Public hasAboutBox As New MyItems

    Public Modules As New MyItems    ' ????????????????????
    '     Dim BreakPoints As myitems
    '     Dim WatchVariables As myitems

    Public Sub New()
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

         ' ----------------------------------------------------------
        setchecked(Find_Setting("OptionStrict", "new").Req, OptionStrict, nc, c)
        setchecked(Find_Setting("OptionExplicit", "new").Req, OptionExplicit, nc, c)
        ' ----------------------------------------------------------
        setchecked(Find_Setting("hasSLN", "new").Req, hasSLN, nc, c)
        setchecked(Find_Setting("hasvbProj", "new").Req, hasVBproj, nc, c)
        setchecked(Find_Setting("hasSplashScreen", "new").Req, hasSplashScreen, nc, c)
        setchecked(Find_Setting("hasAboutBox", "new").Req, hasAboutBox, nc, c)
        '       setchecked(Find_Setting("Module","new").req, Modules, nc, c)

    End Sub

    Public Sub setchecked(chk As Boolean, ByRef obj As MyItems, nc As String, c As String)
        If Not chk Then
            obj.cssNonChk = nc
            obj.req = False
        Else
            obj.cssNonChk = c
            obj.req = True
        End If
    End Sub


    Public Sub StuffAppData(StudID As String, ApplicationName As String, AssignmentRoot As String, AssignmentPath As String, CompDate As Date, OptStrict As Boolean, optExplicit As Boolean, has_SLN As Boolean, VB_Version As String, has_VBProj As Boolean, has_SplashScreen As Boolean, has_AboutBox As Boolean, N_Modules As Integer)

        Try
            StudentID = StudID
            AppTitle = ApplicationName
            AssignRoot = AssignmentRoot
            AssignPath = AssignmentPath
            CompileDate = CompDate.ToString


            OptionStrict.Status = OptStrict.ToString
            OptionExplicit.Status = optExplicit.ToString

            hasSLN.Status = has_SLN.ToString
            VBVersion.Status = VB_Version
            hasVBproj.Status = has_VBProj.ToString
            hasSplashScreen.Status = has_SplashScreen.ToString
            hasAboutBox.Status = has_AboutBox.ToString
            Modules.Status = N_Modules.ToString
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    ' ========================================================================
    Public Class MyItems
        ' each assessment item has a MyItems structure. Req, PtsperError, and PossiblePts are set by instructor to determine
        ' the grading for each assessment item per assignment.
        Public req As Boolean
        Public showVar As Boolean  ' if we want this, need to bring it in on the datagridview.
        Public PtsPerError As Decimal
        Public PossiblePts As Decimal

        Public Status As String          ' holds the main strings associated with the item
        Public cnt As Integer
        Public n As Integer              ' number of instances found
        Public cssClass As String
        Public cssNonChk As String       ' can be either hidden, gray or white(none)
        Public bad As String
        Public good As String

        Public BlockID As Integer
        Public YourPts As Decimal
        Public Comments As String
        Public isBad As Boolean
        Public ID As String

        Sub clear()
            req = False
            showVar = False
            PtsPerError = 0
            PossiblePts = 0

            Status = ""
            cnt = 0
            n = 0
            cssClass = ""
            cssNonChk = ""
            bad = ""
            good = ""

            BlockID = 0
            YourPts = 0
            Comments = ""
            isBad = False
            ID = ""
        End Sub

    End Class
    ' ========================================================================


    ' ========================================================================
    ' Form variables - one each of these per form (Class with form)
    Public Class AppForm
        Public ObjFormPrefixes As New MyItems

        Public ObjButton As New MyItems
        Public ObjLabel As New MyItems
        Public ObjActiveLabel As New MyItems
        Public ObjNonactiveLabel As New MyItems
        Public ObjTextbox As New MyItems
        Public ObjListbox As New MyItems
        Public ObjCombobox As New MyItems
        Public ObjRadioButton As New MyItems
        Public ObjCheckbox As New MyItems
        Public ObjGroupBox As New MyItems
        Public ObjPanel As New MyItems
        Public ObjWebBrowser As New MyItems
        ' Public ObjWebClient As New MyIte

        Public ObjOpenFileDialog As New MyItems
        Public ObjSaveFileDialog As New MyItems

        Public FormName As New MyItems
        Public FormText As New MyItems
        Public FormBackColor As New MyItems
        Public FormAcceptButton As New MyItems
        Public FormCancelButton As New MyItems
        Public FormStartPosition As New MyItems
        Public FormLoadMethod As New MyItems

        Public Sub Clear()
            ObjFormPrefixes.clear()

            ObjButton.clear()
            ObjLabel.clear()
            ObjActiveLabel.clear()
            ObjNonactiveLabel.clear()
            ObjTextbox.clear()
            ObjListbox.clear()
            ObjCombobox.clear()
            ObjRadioButton.clear()
            ObjCheckbox.clear()
            ObjGroupBox.clear()
            ObjPanel.clear()
            ObjWebBrowser.clear()
            '  ObjWebClient.clear()

            ObjOpenFileDialog.clear()
            ObjSaveFileDialog.clear()

            FormName.clear()
            FormText.clear()
            FormBackColor.clear()
            FormAcceptButton.clear()
            FormCancelButton.clear()
            FormStartPosition.clear()
            FormLoadMethod.clear()
        End Sub
    End Class
    ' ========================================================================


    ' ========================================================================
    ' Logic variables - one each of these per code file (class, module, and VB code files)

    Public Class AppSummary
        '      Public AppTitle As myitems

        Public TotalScore As Decimal


        ' dgvDevEnvironment
        Public OptionStrict As New MyItems
        Public OptionExplicit As New MyItems
        Public hasSLN As New MyItems
        Public hasvbProj As New MyItems
        Public hasSplashScreen As New MyItems
        Public hasAboutBox As New MyItems

        ' dgvAppInfo
        Public InfoAppTitle As New MyItems
        Public InfoDescription As New MyItems
        Public InfoCompany As New MyItems
        Public InfoProduct As New MyItems
        Public InfoTrademark As New MyItems
        Public InfoCopyright As New MyItems
        Public InfoGUID As New MyItems

        ' dgvCompileOptions

        ' .dgvComments
        Public CommentGeneral As New MyItems
        Public CommentSub As New MyItems
        Public CommentIF As New MyItems
        Public CommentFor As New MyItems
        Public CommentDo As New MyItems
        Public CommentWhile As New MyItems
        Public CommentSelect As New MyItems


        'dgvFormDesign
        Public RenameObjects As New MyItems
        Public IncludeFrmInFormName As New MyItems
        Public ChangeFormText As New MyItems
        Public ChangeFormColor As New MyItems
        Public SetFormAcceptButton As New MyItems
        Public SetFormCancelButton As New MyItems
        Public ModifyStartPosition As New MyItems


        ' dgvImports
        Public SystemIO As New MyItems
        Public SystemNet As New MyItems
        Public SystemDB As New MyItems

        ' DataStructures

        Public VarArrays As New MyItems
        Public VarLists As New MyItems
        Public VarStructures As New MyItems

        ' DataTypes
        Public VarBoolean As New MyItems
        Public VarInteger As New MyItems
        Public VarDecimal As New MyItems
        Public VarDate As New MyItems
        Public VarString As New MyItems
        Public VariablePrefixes As New MyItems

        ' Coding
        Public LogicIF As New MyItems
        Public LogicFor As New MyItems
        Public LogicDo As New MyItems
        Public LogicWhile As New MyItems
        Public LogicElse As New MyItems
        Public LogicElseIF As New MyItems
        Public LogicMessageBox As New MyItems     ' ??????????????? need to handle this. Add to config
        Public LogicNestedIF As New MyItems
        Public LogicNestedFOR As New MyItems
        Public LogicSelectCase As New MyItems
        Public LogicConcatination As New MyItems
        Public LogicConvertToString As New MyItems
        Public LogicToStringFormat As New MyItems
        Public LogicStringFormat As New MyItems
        Public LogicStringFormatParameters As New MyItems
        Public LogicComplexConditions As New MyItems
        Public LogicCaseInsensitive As New MyItems
        Public LogicTryCatch As New MyItems
        Public LogicStreamReader As New MyItems
        Public LogicStreamWriter As New MyItems
        Public LogicStreamReaderClose As New MyItems
        Public LogicStreamWriterClose As New MyItems

        Public LogicCStr As New MyItems
        Public LogicToString As New MyItems



        ' Subs
        Public LogicSub As New MyItems
        Public LogicFunction As New MyItems
        Public LogicOptional As New MyItems
        Public LogicByRef As New MyItems
        Public LogicMultipleForms As New MyItems
        Public LogicModule As New MyItems
        Public LogicFormLoad As New MyItems


        Public varPrefixes As New MyItems
        Public LogicVarPrefixes As New MyItems

        Public LogicFlowControl As New MyItems
 
        ' ========================================================================

        Public Sub New()   ' AppSummary

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

            setchecked(Find_Setting("InfoAppTitle", "new2").Req, InfoAppTitle, nc, c)
            setchecked(Find_Setting("InfoDescription", "new2").Req, InfoDescription, nc, c)
            setchecked(Find_Setting("InfoCompany", "new2").Req, InfoCompany, nc, c)
            setchecked(Find_Setting("InfoProduct", "new2").Req, InfoProduct, nc, c)
            setchecked(Find_Setting("InfoTrademark", "new2").Req, InfoTrademark, nc, c)
            setchecked(Find_Setting("InfoCopyright", "new2").Req, InfoCopyright, nc, c)
            setchecked(Find_Setting("", "new2").Req, InfoGUID, nc, c)


            setchecked(Find_Setting("CommentGeneral", "new2").Req, CommentGeneral, nc, c)
            setchecked(Find_Setting("CommentSubs", "new2").Req, CommentSub, nc, c)
            setchecked(Find_Setting("CommentIF", "new2").Req, CommentIF, nc, c)
            setchecked(Find_Setting("CommentFOR", "new2").Req, CommentFor, nc, c)
            setchecked(Find_Setting("CommentDO", "new2").Req, CommentDo, nc, c)
            setchecked(Find_Setting("CommentWHILE", "new2").Req, CommentWhile, nc, c)
            setchecked(Find_Setting("CommentSELECT", "new2").Req, CommentSelect, nc, c)

            setchecked(Find_Setting("VarString", "new2").Req, VarString, nc, c)
            setchecked(Find_Setting("VarBoolean", "new2").Req, VarBoolean, nc, c)
            setchecked(Find_Setting("VarInteger", "new2").Req, VarInteger, nc, c)
            setchecked(Find_Setting("VarDecimal", "new2").Req, VarDecimal, nc, c)
            setchecked(Find_Setting("VarDate", "new2").Req, VarDate, nc, c)

            setchecked(Find_Setting("VarArrays", "new2").Req, VarArrays, nc, c)
            setchecked(Find_Setting("VarLists", "new2").Req, VarLists, nc, c)
            setchecked(Find_Setting("VarStructures", "new2").Req, VarStructures, nc, c)

            setchecked(Find_Setting("VariablePrefixes", "new2").Req, varPrefixes, nc, c)

            ' setchecked(Find_Setting("","new2").req, LogicFlowControl, nc, c)
            setchecked(Find_Setting("LogicIF", "new2").Req, LogicIF, nc, c)
            setchecked(Find_Setting("LogicFOR", "new2").Req, LogicFor, nc, c)
            setchecked(Find_Setting("LogicDO", "new2").Req, LogicDo, nc, c)
            setchecked(Find_Setting("LogicWHILE", "new2").Req, LogicWhile, nc, c)
            setchecked(Find_Setting("LogicSelectCase", "new2").Req, LogicSelectCase, nc, c)
            setchecked(Find_Setting("LogicElse", "new2").Req, LogicElse, nc, c)
            setchecked(Find_Setting("LogicElseIF", "new2").Req, LogicElseIF, nc, c)
            setchecked(Find_Setting("LogicTryCatch", "new2").Req, LogicTryCatch, nc, c)
            setchecked(Find_Setting("LogicStreamReader", "new2").Req, LogicStreamReader, nc, c)
            setchecked(Find_Setting("LogicStreamWriter", "new2").Req, LogicStreamWriter, nc, c)
            setchecked(Find_Setting("LogicStreamReaderClose", "new2").Req, LogicStreamReaderClose, nc, c)
            setchecked(Find_Setting("LogicStreamWriterClose", "new2").Req, LogicStreamWriterClose, nc, c)
            setchecked(Find_Setting("LogicSub", "new2").Req, LogicSub, nc, c)
            '  setchecked(Find_Setting("","new2").req, LogicFunction, nc, c)
            setchecked(Find_Setting("LogicOptional", "new2").Req, LogicOptional, nc, c)
            setchecked(Find_Setting("LogicByRef", "new2").Req, LogicByRef, nc, c)
            setchecked(Find_Setting("LogicConvertToString", "new2").Req, LogicCStr, nc, c)
            '   setchecked(Find_Setting("","new2").req, LogicToString, nc, c)
            setchecked(Find_Setting("LogicStringFormat", "new2").Req, LogicStringFormat, nc, c)
            setchecked(Find_Setting("LogicConcatination", "new2").Req, LogicConcatination, nc, c)

            'If Not chkMultiForm,  MultiForm.cssHideNonChk = nchide
            'If Not chkLineNbr,  OptionStrict.cssHideNonChk = nchide 
            'If Not chkWordWrap,  OptionStrict.cssHideNonChk = nchide
            ' ----------------------------------------------------------
            'setchecked(Find_Setting("hasSLN").Req, Assign.hasSLN, nc, c)
            'setchecked(Find_Setting("hasvbProj").Req, Assign.hasVBproj, nc, c)
            'setchecked(Find_Setting("hasSplashScreen").Req, Assign.hasSplashScreen, nc, c)
            'setchecked(Find_Setting("hasAboutBox").Req, Assign.hasAboutBox, nc, c)
            'setchecked(Find_Setting("Include Module").Req, Assign.Modules, nc, c)
            'setchecked(Find_Setting("OptionStrict").Req, Assign.OptionStrict, nc, c)
            'setchecked(Find_Setting("OptionExplicit").Req, Assign.OptionExplicit, nc, c)
            'setchecked(Find_Setting("Include a Form LOAD Method").Req, LogicFormLoad, nc, c)
            ' ----------------------------------------------------------
        End Sub

        Public Sub setchecked(chk As Boolean, ByRef obj As MyItems, nc As String, c As String)
            If Not chk Then
                obj.cssNonChk = nc
                obj.req = False
            Else
                obj.cssNonChk = c
                obj.req = True
            End If
        End Sub

        Public Sub Clear()
            InfoAppTitle.clear()
            InfoDescription.clear()
            InfoCompany.clear()
            InfoProduct.clear()
            InfoTrademark.clear()
            InfoCopyright.clear()
            InfoGUID.clear()



            CommentGeneral.clear()
            CommentSub.clear()
            CommentIF.clear()
            CommentFor.clear()
            CommentDo.clear()
            CommentWhile.clear()
            CommentSelect.clear()

            VarBoolean.clear()
            VarInteger.clear()
            VarDecimal.clear()
            VarDate.clear()
            VarString.clear()

            VarArrays.clear()
            VarLists.clear()
            VarStructures.clear()

            varPrefixes.clear()

            LogicFlowControl.clear()
            LogicIF.clear()
            LogicFor.clear()
            LogicDo.clear()
            LogicWhile.clear()
            LogicSelectCase.clear()
            LogicElse.clear()
            LogicElseIF.clear()
            LogicTryCatch.clear()
            LogicSub.clear()
            LogicFunction.clear()
            LogicOptional.clear()
            LogicByRef.clear()
            LogicModule.clear()
            LogicMultipleForms.clear()
            LogicFormLoad.clear()
            LogicCStr.clear()
            LogicToString.clear()
            LogicToStringFormat.clear()
            LogicConvertToString.clear()
            LogicStringFormatParameters.clear()
            LogicCaseInsensitive.clear()
            LogicComplexConditions.clear()

            LogicVarPrefixes.clear()
            LogicNestedIF.clear()
            LogicNestedFor.clear()

            '       LogicStringFormatting= nothing
            LogicComplexConditions.clear()
            LogicCaseInsensitive.clear()
            LogicStringFormat.clear()
            LogicConcatination.clear()

            LogicStreamReader.clear()
            LogicStreamReaderClose.clear()
            LogicStreamWriter.clear()
            LogicStreamWriterClose.clear()

            SystemIO.clear()
            SystemNet.clear()
            SystemDB.clear()

        End Sub
    End Class     ' AppSummary

End Class
