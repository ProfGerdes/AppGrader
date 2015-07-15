<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.btnProcessApps = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblNZips = New System.Windows.Forms.Label()
        Me.btnOutput = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.cbxLoadWordTemplate = New System.Windows.Forms.CheckBox()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.lblMessage = New System.Windows.Forms.Label()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ConfigureAppToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProcessToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.IndentifyInstructorDemoDirectoryToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SelectTargetFileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StartProcessingToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewOutputToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PlagiarismSummaryToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator6 = New System.Windows.Forms.ToolStripSeparator()
        Me.ExitToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.toolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.btnSelectFile = New System.Windows.Forms.Button()
        Me.lblGroupWorkZip = New System.Windows.Forms.Label()
        Me.lblTarget = New System.Windows.Forms.Label()
        Me.lblTargetFile = New System.Windows.Forms.Label()
        Me.lblDir = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtAssignmentName = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnBrowseTemplate = New System.Windows.Forms.Button()
        Me.lblSelectedTemplate = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnViewPlagiarism = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.cbxJustUnzip = New System.Windows.Forms.CheckBox()
        Me.rbnSingleProject = New System.Windows.Forms.RadioButton()
        Me.rbnBlackboardZip = New System.Windows.Forms.RadioButton()
        Me.FolderBrowserDialog2 = New System.Windows.Forms.FolderBrowserDialog()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.cbxNoDemoFiles = New System.Windows.Forms.CheckBox()
        Me.btnDemo = New System.Windows.Forms.Button()
        Me.lblDemoDir = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.rbnShowOnlyReq = New System.Windows.Forms.RadioButton()
        Me.rbnShowAll = New System.Windows.Forms.RadioButton()
        Me.rbnCheckGray = New System.Windows.Forms.RadioButton()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.btnAssignSummary = New System.Windows.Forms.Button()
        Me.MenuStrip1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'FolderBrowserDialog1
        '
        Me.FolderBrowserDialog1.RootFolder = System.Environment.SpecialFolder.MyComputer
        Me.FolderBrowserDialog1.ShowNewFolderButton = False
        '
        'btnProcessApps
        '
        Me.btnProcessApps.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.btnProcessApps.Enabled = False
        Me.btnProcessApps.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnProcessApps.Location = New System.Drawing.Point(524, 576)
        Me.btnProcessApps.Margin = New System.Windows.Forms.Padding(2)
        Me.btnProcessApps.Name = "btnProcessApps"
        Me.btnProcessApps.Size = New System.Drawing.Size(128, 33)
        Me.btnProcessApps.TabIndex = 1
        Me.btnProcessApps.Text = "Start Processing"
        Me.btnProcessApps.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(23, 584)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(175, 17)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Number of Zip files found: "
        Me.Label1.Visible = False
        '
        'lblNZips
        '
        Me.lblNZips.AutoSize = True
        Me.lblNZips.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNZips.Location = New System.Drawing.Point(159, 584)
        Me.lblNZips.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblNZips.Name = "lblNZips"
        Me.lblNZips.Size = New System.Drawing.Size(59, 17)
        Me.lblNZips.TabIndex = 3
        Me.lblNZips.Text = "lblNZips"
        Me.lblNZips.Visible = False
        '
        'btnOutput
        '
        Me.btnOutput.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.btnOutput.Location = New System.Drawing.Point(9, 42)
        Me.btnOutput.Name = "btnOutput"
        Me.btnOutput.Size = New System.Drawing.Size(158, 34)
        Me.btnOutput.TabIndex = 8
        Me.btnOutput.Text = "View Assessment"
        Me.btnOutput.UseVisualStyleBackColor = False
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        Me.OpenFileDialog1.Filter = "Docx Files|*.docx|All Files|*.*"
        '
        'cbxLoadWordTemplate
        '
        Me.cbxLoadWordTemplate.AutoSize = True
        Me.cbxLoadWordTemplate.Location = New System.Drawing.Point(126, 154)
        Me.cbxLoadWordTemplate.Margin = New System.Windows.Forms.Padding(2)
        Me.cbxLoadWordTemplate.Name = "cbxLoadWordTemplate"
        Me.cbxLoadWordTemplate.Size = New System.Drawing.Size(179, 21)
        Me.cbxLoadWordTemplate.TabIndex = 14
        Me.cbxLoadWordTemplate.Text = "Generate Grade Sheets"
        Me.cbxLoadWordTemplate.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        '
        'lblMessage
        '
        Me.lblMessage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMessage.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblMessage.Location = New System.Drawing.Point(6, 6)
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.Size = New System.Drawing.Size(646, 21)
        Me.lblMessage.TabIndex = 22
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ConfigureAppToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(678, 24)
        Me.MenuStrip1.TabIndex = 25
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ConfigureAppToolStripMenuItem
        '
        Me.ConfigureAppToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ProcessToolStripMenuItem, Me.IndentifyInstructorDemoDirectoryToolStripMenuItem, Me.SelectTargetFileToolStripMenuItem, Me.StartProcessingToolStripMenuItem, Me.ViewOutputToolStripMenuItem, Me.PlagiarismSummaryToolStripMenuItem, Me.ToolStripSeparator6, Me.ExitToolStripMenuItem1})
        Me.ConfigureAppToolStripMenuItem.Name = "ConfigureAppToolStripMenuItem"
        Me.ConfigureAppToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.ConfigureAppToolStripMenuItem.Text = "File"
        '
        'ProcessToolStripMenuItem
        '
        Me.ProcessToolStripMenuItem.Name = "ProcessToolStripMenuItem"
        Me.ProcessToolStripMenuItem.Size = New System.Drawing.Size(261, 22)
        Me.ProcessToolStripMenuItem.Text = "Application Settings"
        '
        'IndentifyInstructorDemoDirectoryToolStripMenuItem
        '
        Me.IndentifyInstructorDemoDirectoryToolStripMenuItem.Name = "IndentifyInstructorDemoDirectoryToolStripMenuItem"
        Me.IndentifyInstructorDemoDirectoryToolStripMenuItem.Size = New System.Drawing.Size(261, 22)
        Me.IndentifyInstructorDemoDirectoryToolStripMenuItem.Text = "Indentify Instructor Demo Directory"
        '
        'SelectTargetFileToolStripMenuItem
        '
        Me.SelectTargetFileToolStripMenuItem.Name = "SelectTargetFileToolStripMenuItem"
        Me.SelectTargetFileToolStripMenuItem.Size = New System.Drawing.Size(261, 22)
        Me.SelectTargetFileToolStripMenuItem.Text = "Select Target File"
        '
        'StartProcessingToolStripMenuItem
        '
        Me.StartProcessingToolStripMenuItem.Name = "StartProcessingToolStripMenuItem"
        Me.StartProcessingToolStripMenuItem.Size = New System.Drawing.Size(261, 22)
        Me.StartProcessingToolStripMenuItem.Text = "Start Processing"
        '
        'ViewOutputToolStripMenuItem
        '
        Me.ViewOutputToolStripMenuItem.Name = "ViewOutputToolStripMenuItem"
        Me.ViewOutputToolStripMenuItem.Size = New System.Drawing.Size(261, 22)
        Me.ViewOutputToolStripMenuItem.Text = "View Assessment"
        '
        'PlagiarismSummaryToolStripMenuItem
        '
        Me.PlagiarismSummaryToolStripMenuItem.Name = "PlagiarismSummaryToolStripMenuItem"
        Me.PlagiarismSummaryToolStripMenuItem.Size = New System.Drawing.Size(261, 22)
        Me.PlagiarismSummaryToolStripMenuItem.Text = "Plagiarism Summary"
        '
        'ToolStripSeparator6
        '
        Me.ToolStripSeparator6.Name = "ToolStripSeparator6"
        Me.ToolStripSeparator6.Size = New System.Drawing.Size(258, 6)
        '
        'ExitToolStripMenuItem1
        '
        Me.ExitToolStripMenuItem1.Name = "ExitToolStripMenuItem1"
        Me.ExitToolStripMenuItem1.Size = New System.Drawing.Size(261, 22)
        Me.ExitToolStripMenuItem1.Text = "Exit"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.toolStripSeparator5})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(52, 20)
        Me.HelpToolStripMenuItem.Text = "&About"
        '
        'toolStripSeparator5
        '
        Me.toolStripSeparator5.Name = "toolStripSeparator5"
        Me.toolStripSeparator5.Size = New System.Drawing.Size(57, 6)
        '
        'btnSelectFile
        '
        Me.btnSelectFile.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.btnSelectFile.Location = New System.Drawing.Point(515, 282)
        Me.btnSelectFile.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSelectFile.Name = "btnSelectFile"
        Me.btnSelectFile.Size = New System.Drawing.Size(128, 29)
        Me.btnSelectFile.TabIndex = 32
        Me.btnSelectFile.Text = "Select File"
        Me.btnSelectFile.UseVisualStyleBackColor = False
        '
        'lblGroupWorkZip
        '
        Me.lblGroupWorkZip.AutoSize = True
        Me.lblGroupWorkZip.Location = New System.Drawing.Point(121, 210)
        Me.lblGroupWorkZip.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblGroupWorkZip.Name = "lblGroupWorkZip"
        Me.lblGroupWorkZip.Size = New System.Drawing.Size(0, 13)
        Me.lblGroupWorkZip.TabIndex = 33
        '
        'lblTarget
        '
        Me.lblTarget.AutoSize = True
        Me.lblTarget.Location = New System.Drawing.Point(12, 38)
        Me.lblTarget.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTarget.Name = "lblTarget"
        Me.lblTarget.Size = New System.Drawing.Size(100, 17)
        Me.lblTarget.TabIndex = 34
        Me.lblTarget.Text = "Target Zip File"
        '
        'lblTargetFile
        '
        Me.lblTargetFile.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTargetFile.Location = New System.Drawing.Point(162, 37)
        Me.lblTargetFile.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTargetFile.Name = "lblTargetFile"
        Me.lblTargetFile.Size = New System.Drawing.Size(396, 19)
        Me.lblTargetFile.TabIndex = 35
        '
        'lblDir
        '
        Me.lblDir.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblDir.Location = New System.Drawing.Point(162, 65)
        Me.lblDir.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblDir.Name = "lblDir"
        Me.lblDir.Size = New System.Drawing.Size(396, 19)
        Me.lblDir.TabIndex = 37
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 66)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(146, 17)
        Me.Label5.TabIndex = 36
        Me.Label5.Text = "Output Root Directory"
        '
        'btnBrowse
        '
        Me.btnBrowse.Location = New System.Drawing.Point(567, 61)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(71, 26)
        Me.btnBrowse.TabIndex = 38
        Me.btnBrowse.Text = "Browse"
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(12, 95)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(122, 17)
        Me.Label6.TabIndex = 39
        Me.Label6.Text = "Assignment Name"
        '
        'txtAssignmentName
        '
        Me.txtAssignmentName.Location = New System.Drawing.Point(162, 100)
        Me.txtAssignmentName.Name = "txtAssignmentName"
        Me.txtAssignmentName.Size = New System.Drawing.Size(396, 23)
        Me.txtAssignmentName.TabIndex = 40
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 124)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(111, 17)
        Me.Label3.TabIndex = 41
        Me.Label3.Text = "Grade Template"
        '
        'btnBrowseTemplate
        '
        Me.btnBrowseTemplate.Location = New System.Drawing.Point(567, 125)
        Me.btnBrowseTemplate.Name = "btnBrowseTemplate"
        Me.btnBrowseTemplate.Size = New System.Drawing.Size(71, 26)
        Me.btnBrowseTemplate.TabIndex = 43
        Me.btnBrowseTemplate.Text = "Browse"
        Me.btnBrowseTemplate.UseVisualStyleBackColor = True
        '
        'lblSelectedTemplate
        '
        Me.lblSelectedTemplate.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSelectedTemplate.Location = New System.Drawing.Point(162, 129)
        Me.lblSelectedTemplate.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblSelectedTemplate.Name = "lblSelectedTemplate"
        Me.lblSelectedTemplate.Size = New System.Drawing.Size(396, 19)
        Me.lblSelectedTemplate.TabIndex = 42
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnAssignSummary)
        Me.Panel1.Controls.Add(Me.btnExit)
        Me.Panel1.Controls.Add(Me.btnViewPlagiarism)
        Me.Panel1.Controls.Add(Me.lblMessage)
        Me.Panel1.Controls.Add(Me.btnOutput)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Panel1.Location = New System.Drawing.Point(0, 620)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(678, 80)
        Me.Panel1.TabIndex = 44
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(576, 42)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(71, 26)
        Me.btnExit.TabIndex = 39
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnViewPlagiarism
        '
        Me.btnViewPlagiarism.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.btnViewPlagiarism.Location = New System.Drawing.Point(188, 42)
        Me.btnViewPlagiarism.Name = "btnViewPlagiarism"
        Me.btnViewPlagiarism.Size = New System.Drawing.Size(154, 33)
        Me.btnViewPlagiarism.TabIndex = 23
        Me.btnViewPlagiarism.Text = "Plagiarism Summary"
        Me.btnViewPlagiarism.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.Button1.Location = New System.Drawing.Point(404, -291)
        Me.Button1.Margin = New System.Windows.Forms.Padding(2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(128, 29)
        Me.Button1.TabIndex = 32
        Me.Button1.Text = "Select File"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'cbxJustUnzip
        '
        Me.cbxJustUnzip.AutoSize = True
        Me.cbxJustUnzip.Location = New System.Drawing.Point(25, 287)
        Me.cbxJustUnzip.Margin = New System.Windows.Forms.Padding(2)
        Me.cbxJustUnzip.Name = "cbxJustUnzip"
        Me.cbxJustUnzip.Size = New System.Drawing.Size(334, 21)
        Me.cbxJustUnzip.TabIndex = 45
        Me.cbxJustUnzip.Text = "Just Unzip && shorten filenames (No Assessment)"
        Me.cbxJustUnzip.UseVisualStyleBackColor = True
        '
        'rbnSingleProject
        '
        Me.rbnSingleProject.AutoSize = True
        Me.rbnSingleProject.Checked = True
        Me.rbnSingleProject.Location = New System.Drawing.Point(7, 3)
        Me.rbnSingleProject.Name = "rbnSingleProject"
        Me.rbnSingleProject.Size = New System.Drawing.Size(113, 21)
        Me.rbnSingleProject.TabIndex = 46
        Me.rbnSingleProject.TabStop = True
        Me.rbnSingleProject.Text = "Single Project"
        Me.rbnSingleProject.UseVisualStyleBackColor = True
        '
        'rbnBlackboardZip
        '
        Me.rbnBlackboardZip.AutoSize = True
        Me.rbnBlackboardZip.Location = New System.Drawing.Point(7, 21)
        Me.rbnBlackboardZip.Name = "rbnBlackboardZip"
        Me.rbnBlackboardZip.Size = New System.Drawing.Size(147, 21)
        Me.rbnBlackboardZip.TabIndex = 47
        Me.rbnBlackboardZip.Text = "Blackboard Zip File"
        Me.rbnBlackboardZip.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.Panel3)
        Me.GroupBox1.Controls.Add(Me.Panel2)
        Me.GroupBox1.Controls.Add(Me.cbxJustUnzip)
        Me.GroupBox1.Controls.Add(Me.btnSelectFile)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(9, 41)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(657, 328)
        Me.GroupBox1.TabIndex = 48
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Select Applications to Process"
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.GroupBox3.Controls.Add(Me.cbxNoDemoFiles)
        Me.GroupBox3.Controls.Add(Me.btnDemo)
        Me.GroupBox3.Controls.Add(Me.lblDemoDir)
        Me.GroupBox3.Controls.Add(Me.Label4)
        Me.GroupBox3.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.GroupBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(21, 23)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(617, 93)
        Me.GroupBox3.TabIndex = 44
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Instructor Demo Applications"
        '
        'cbxNoDemoFiles
        '
        Me.cbxNoDemoFiles.AutoSize = True
        Me.cbxNoDemoFiles.Location = New System.Drawing.Point(12, 67)
        Me.cbxNoDemoFiles.Margin = New System.Windows.Forms.Padding(2)
        Me.cbxNoDemoFiles.Name = "cbxNoDemoFiles"
        Me.cbxNoDemoFiles.Size = New System.Drawing.Size(218, 21)
        Me.cbxNoDemoFiles.TabIndex = 46
        Me.cbxNoDemoFiles.Text = "Do not load demo applications"
        Me.cbxNoDemoFiles.UseVisualStyleBackColor = True
        '
        'btnDemo
        '
        Me.btnDemo.Location = New System.Drawing.Point(540, 32)
        Me.btnDemo.Name = "btnDemo"
        Me.btnDemo.Size = New System.Drawing.Size(71, 26)
        Me.btnDemo.TabIndex = 41
        Me.btnDemo.Text = "Browse"
        Me.btnDemo.UseVisualStyleBackColor = True
        '
        'lblDemoDir
        '
        Me.lblDemoDir.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblDemoDir.Location = New System.Drawing.Point(123, 32)
        Me.lblDemoDir.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblDemoDir.Name = "lblDemoDir"
        Me.lblDemoDir.Size = New System.Drawing.Size(414, 26)
        Me.lblDemoDir.TabIndex = 40
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 40)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(106, 17)
        Me.Label4.TabIndex = 39
        Me.Label4.Text = "Demo Directory"
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.rbnBlackboardZip)
        Me.Panel3.Controls.Add(Me.rbnSingleProject)
        Me.Panel3.Location = New System.Drawing.Point(26, 133)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(374, 47)
        Me.Panel3.TabIndex = 52
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.rbnShowOnlyReq)
        Me.Panel2.Controls.Add(Me.rbnShowAll)
        Me.Panel2.Controls.Add(Me.rbnCheckGray)
        Me.Panel2.Location = New System.Drawing.Point(25, 197)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(375, 72)
        Me.Panel2.TabIndex = 51
        '
        'rbnShowOnlyReq
        '
        Me.rbnShowOnlyReq.AutoSize = True
        Me.rbnShowOnlyReq.Location = New System.Drawing.Point(7, 3)
        Me.rbnShowOnlyReq.Name = "rbnShowOnlyReq"
        Me.rbnShowOnlyReq.Size = New System.Drawing.Size(205, 21)
        Me.rbnShowOnlyReq.TabIndex = 51
        Me.rbnShowOnlyReq.Text = "Show Only Required Checks"
        Me.rbnShowOnlyReq.UseVisualStyleBackColor = True
        '
        'rbnShowAll
        '
        Me.rbnShowAll.AutoSize = True
        Me.rbnShowAll.Location = New System.Drawing.Point(7, 39)
        Me.rbnShowAll.Name = "rbnShowAll"
        Me.rbnShowAll.Size = New System.Drawing.Size(129, 21)
        Me.rbnShowAll.TabIndex = 50
        Me.rbnShowAll.Text = "Show All Checks"
        Me.rbnShowAll.UseVisualStyleBackColor = True
        '
        'rbnCheckGray
        '
        Me.rbnCheckGray.AutoSize = True
        Me.rbnCheckGray.Location = New System.Drawing.Point(7, 21)
        Me.rbnCheckGray.Name = "rbnCheckGray"
        Me.rbnCheckGray.Size = New System.Drawing.Size(289, 21)
        Me.rbnCheckGray.TabIndex = 49
        Me.rbnCheckGray.Text = "Grayout checks not related to assignment"
        Me.rbnCheckGray.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.btnBrowseTemplate)
        Me.GroupBox2.Controls.Add(Me.lblSelectedTemplate)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.txtAssignmentName)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.btnBrowse)
        Me.GroupBox2.Controls.Add(Me.lblDir)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.lblTargetFile)
        Me.GroupBox2.Controls.Add(Me.lblTarget)
        Me.GroupBox2.Controls.Add(Me.cbxLoadWordTemplate)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(9, 384)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(657, 187)
        Me.GroupBox2.TabIndex = 49
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Assignment Detail"
        '
        'btnAssignSummary
        '
        Me.btnAssignSummary.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.btnAssignSummary.Location = New System.Drawing.Point(363, 43)
        Me.btnAssignSummary.Name = "btnAssignSummary"
        Me.btnAssignSummary.Size = New System.Drawing.Size(154, 33)
        Me.btnAssignSummary.TabIndex = 40
        Me.btnAssignSummary.Text = "Assignment Summary"
        Me.btnAssignSummary.UseVisualStyleBackColor = False
        '
        'frmMain
        '
        Me.AcceptButton = Me.btnSelectFile
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.ClientSize = New System.Drawing.Size(678, 700)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.lblGroupWorkZip)
        Me.Controls.Add(Me.lblNZips)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnProcessApps)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "AppGrader - Automated Application Assessment"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents btnProcessApps As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblNZips As System.Windows.Forms.Label
    Friend WithEvents btnOutput As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents cbxLoadWordTemplate As System.Windows.Forms.CheckBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents lblMessage As System.Windows.Forms.Label
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ConfigureAppToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ProcessToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StartProcessingToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnSelectFile As System.Windows.Forms.Button
    Friend WithEvents lblGroupWorkZip As System.Windows.Forms.Label
    Friend WithEvents lblTarget As System.Windows.Forms.Label
    Friend WithEvents lblTargetFile As System.Windows.Forms.Label
    Friend WithEvents lblDir As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtAssignmentName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnBrowseTemplate As System.Windows.Forms.Button
    Friend WithEvents lblSelectedTemplate As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents toolStripSeparator5 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents SelectTargetFileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator6 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents cbxJustUnzip As System.Windows.Forms.CheckBox
    Friend WithEvents rbnSingleProject As System.Windows.Forms.RadioButton
    Friend WithEvents rbnBlackboardZip As System.Windows.Forms.RadioButton
    Friend WithEvents FolderBrowserDialog2 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents rbnShowAll As System.Windows.Forms.RadioButton
    Friend WithEvents rbnCheckGray As System.Windows.Forms.RadioButton
    Friend WithEvents ViewOutputToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents btnDemo As System.Windows.Forms.Button
    Friend WithEvents lblDemoDir As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents IndentifyInstructorDemoDirectoryToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnViewPlagiarism As System.Windows.Forms.Button
    Friend WithEvents PlagiarismSummaryToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents rbnShowOnlyReq As System.Windows.Forms.RadioButton
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents cbxNoDemoFiles As System.Windows.Forms.CheckBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents btnAssignSummary As System.Windows.Forms.Button

End Class
