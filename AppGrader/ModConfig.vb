Module ModConfig
    Public Sub StuffSettingsFromFile(ByRef ss() As String, item As Assignment.MyItems, bix As Integer, xdgv As Integer)

        ' This is used to read settings from a file and populate the configuration datagridviews, and the control structures for the app.

        With item  ' Populate the control structures
            .req = CBool(ss(1))
            .showVar = CBool(ss(2))
            .PtsPerError = CDec(ss(3))
            .PossiblePts = CDec(ss(4))
        End With

        ' Populate the data grid views.
        Select Case xdgv
            Case dgvs.AdvancedVariables
                Dim row As String() = New String() {BackIndexAppSum(bix).Name, "", item.req.ToString, item.showVar.ToString, item.PtsPerError.ToString, item.PossiblePts.ToString}
                frmConfig.dgvFormDesign.Rows.Add(row)

            Case dgvs.ApplicationSettings    ' not correct ??????????????????????
                'Dim row As String() = New String() {BackIndexAppSum(bix).Name, "", item.req.ToString, item.show.ToString, item.PtsPerError.ToString, item.PossiblePts.ToString}
                'frmConfig.dgvAdvanced.Rows.Add(row)

            Case dgvs.FormProperties
                Dim row As String() = New String() {BackIndexAppFrm(bix).Name, "", item.req.ToString, item.PtsPerError.ToString, item.PossiblePts.ToString}
                frmConfig.dgvImports.Rows.Add(row)

            Case dgvs.LogicVariables
                Dim row As String() = New String() {BackIndexAppSum(bix).Name, "", item.req.ToString, item.showVar.ToString, item.PtsPerError.ToString, item.PossiblePts.ToString}
                frmConfig.dgvCompileOptions.Rows.Add(row)

            Case dgvs.Splash
                Dim row As String() = New String() {BackIndexAppSum(bix).Name, "", item.req.ToString, item.showVar.ToString, item.PtsPerError.ToString, item.PossiblePts.ToString}
                frmConfig.dgvComments.Rows.Add(row)

            Case dgvs.SystemVariables
                Dim row As String() = New String() {BackIndexAppSum(bix).Name, "", item.req.ToString, item.showVar.ToString, item.PtsPerError.ToString, item.PossiblePts.ToString}
                frmConfig.dgvAppInfo.Rows.Add(row)
        End Select

    End Sub

   
End Module
