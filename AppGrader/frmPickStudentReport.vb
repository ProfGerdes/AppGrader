Imports System.IO

Public Class frmPickStudentReport

    Private myTable As New DataTable()
    Private Sub frmPickStudentReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim tmp As String

        lbxStudents.ClearSelected()

        'then in your form load, you do
        ' http://www.vbforums.com/showthread.php?645103-Add-listbox-items-with-text-and-value

        With myTable.Columns
            .Add("DisplayValue", GetType(String))
            .Add("HiddenValue", GetType(String))   '<<<< change the type of this column to what you actually need instead of integer.
        End With

        lbxStudents.DisplayMember = "DisplayValue"
        lbxStudents.ValueMember = "HiddenValue"
        lbxStudents.DataSource = myTable

        'When you want to add an item to the list box, you add it to the table like this
        '   myTable.Rows.Add("New Item", 1000)

        If ReportType = "Integrated" Then
            For Each filename In IO.Directory.GetFiles(strOutputPath, "*IntegratedReport.html", SearchOption.AllDirectories)
                tmp = ReturnLastField(filename, "\")
                myTable.Rows.Add(tmp, filename)
            Next
        Else
            For Each filename In IO.Directory.GetFiles(strOutputPath, "*GradeReport.html", SearchOption.AllDirectories)
                tmp = ReturnLastField(filename, "\")
                myTable.Rows.Add(tmp, filename)
            Next
        End If

    End Sub

    Private Sub lbxStudents_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbxStudents.SelectedIndexChanged

        Dim url As New Uri("file:\\\" & lbxStudents.SelectedValue.ToString)

        frmOutput.WebBrowser1.Url = url
        frmOutput.Show()

    End Sub

    Private Sub btnReturn_Click(sender As Object, e As EventArgs) Handles btnReturn.Click
        Me.Close()
    End Sub
End Class