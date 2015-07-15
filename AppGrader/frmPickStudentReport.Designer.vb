<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPickStudentReport
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnReturn = New System.Windows.Forms.Button()
        Me.lbxStudents = New System.Windows.Forms.ListBox()
        Me.SuspendLayout()
        '
        'btnReturn
        '
        Me.btnReturn.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnReturn.Location = New System.Drawing.Point(248, 116)
        Me.btnReturn.Name = "btnReturn"
        Me.btnReturn.Size = New System.Drawing.Size(75, 23)
        Me.btnReturn.TabIndex = 0
        Me.btnReturn.Text = "Return"
        Me.btnReturn.UseVisualStyleBackColor = True
        '
        'lbxStudents
        '
        Me.lbxStudents.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbxStudents.FormattingEnabled = True
        Me.lbxStudents.Location = New System.Drawing.Point(27, 26)
        Me.lbxStudents.Name = "lbxStudents"
        Me.lbxStudents.Size = New System.Drawing.Size(215, 108)
        Me.lbxStudents.TabIndex = 2
        '
        'frmPickStudentReport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.ClientSize = New System.Drawing.Size(331, 151)
        Me.Controls.Add(Me.lbxStudents)
        Me.Controls.Add(Me.btnReturn)
        Me.Name = "frmPickStudentReport"
        Me.Text = "Select Student Assessment Report to View"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnReturn As System.Windows.Forms.Button
    Friend WithEvents lbxStudents As System.Windows.Forms.ListBox
End Class
