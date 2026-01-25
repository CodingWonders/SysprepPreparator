<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainForm
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
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

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
        Me.ButtonPanel = New System.Windows.Forms.Panel()
        Me.AboutBtn = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Back_Button = New System.Windows.Forms.Button()
        Me.Next_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.BannerPanel = New System.Windows.Forms.Panel()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.PageContainerPanel = New System.Windows.Forms.Panel()
        Me.WelcomePage = New System.Windows.Forms.Panel()
        Me.WelcomePage_Description = New System.Windows.Forms.Label()
        Me.WelcomePage_Header = New System.Windows.Forms.Label()
        Me.SystemCheckPanel = New System.Windows.Forms.Panel()
        Me.SysCheckPage_CCPProgressPanel = New System.Windows.Forms.Panel()
        Me.SysCheckPage_ProgressLabel = New System.Windows.Forms.Label()
        Me.SysCheckPage_CCPProgressBar = New System.Windows.Forms.ProgressBar()
        Me.SysCheckPage_CheckAgainBtn = New System.Windows.Forms.Button()
        Me.SysCheckPage_CheckDetailsGB = New System.Windows.Forms.GroupBox()
        Me.SysCheckPage_CheckDetailsTLP = New System.Windows.Forms.TableLayoutPanel()
        Me.SysCheckPage_CheckDetails_Resolution = New System.Windows.Forms.Label()
        Me.SysCheckPage_CheckDetails_DescriptionValue = New System.Windows.Forms.Label()
        Me.SysCheckPage_CheckDetails_Description = New System.Windows.Forms.Label()
        Me.SysCheckPage_CheckDetails_Title = New System.Windows.Forms.Label()
        Me.SysCheckPage_CheckDetails_ResolutionValue = New System.Windows.Forms.TextBox()
        Me.SysCheckPage_ChecksLv = New System.Windows.Forms.ListView()
        Me.SysCheckPage_CheckCH = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.SysCheckPage_CompatibleCH = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.SysCheckPage_SeverityCH = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.SysCheckPage_Description = New System.Windows.Forms.Label()
        Me.SysCheckPage_Header = New System.Windows.Forms.Label()
        Me.AdvSettingsPanel = New System.Windows.Forms.Panel()
        Me.AdvSettingsPage_SysprepPrepToolDeploySteps = New System.Windows.Forms.Label()
        Me.AdvSettingsPage_SysprepConfigPanel = New System.Windows.Forms.Panel()
        Me.AdvSettingsPage_VMModeExplanationLink = New System.Windows.Forms.LinkLabel()
        Me.AdvSettingsPage_VMMode = New System.Windows.Forms.CheckBox()
        Me.AdvSettingsPage_SysprepUnatt_Btn = New System.Windows.Forms.Button()
        Me.AdvSettingsPage_SysprepUnatt_AnswerFileText = New System.Windows.Forms.TextBox()
        Me.AdvSettingsPage_SysprepUnattLabel = New System.Windows.Forms.Label()
        Me.AdvSettingsPage_CleanupAction_Generalize = New System.Windows.Forms.CheckBox()
        Me.AdvSettingsPage_ShutdownOptionsCBox = New System.Windows.Forms.ComboBox()
        Me.AdvSettingsPage_CleanupActionCBox = New System.Windows.Forms.ComboBox()
        Me.AdvSettingsPage_ShutdownOptionLabel = New System.Windows.Forms.Label()
        Me.AdvSettingsPage_CleanupActionLabel = New System.Windows.Forms.Label()
        Me.AdvSettingsPage_ConfigSysprepSettings = New System.Windows.Forms.CheckBox()
        Me.AdvSettingsPage_Description = New System.Windows.Forms.Label()
        Me.AdvSettingsPage_Header = New System.Windows.Forms.Label()
        Me.SettingPreparationPanel = New System.Windows.Forms.Panel()
        Me.SettingPreparationPanel_TaskProgressBar = New System.Windows.Forms.ProgressBar()
        Me.SettingPreparationPanel_TaskLv = New System.Windows.Forms.ListView()
        Me.SettingPreparationPage_TaskCH = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.SettingPreparationPage_SuccessfulCH = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.SettingPreparationPage_SubProcessProgressLabel = New System.Windows.Forms.Label()
        Me.SettingPreparationPage_ProgressLabel = New System.Windows.Forms.Label()
        Me.SettingPreparationPage_Description = New System.Windows.Forms.Label()
        Me.SettingPreparationPage_Header = New System.Windows.Forms.Label()
        Me.FinishPanel = New System.Windows.Forms.Panel()
        Me.FinishPage_RestartBtn = New System.Windows.Forms.Button()
        Me.FinishPage_ResysprepBtn = New System.Windows.Forms.Button()
        Me.FinishPage_CloseBtn = New System.Windows.Forms.Button()
        Me.FinishPage_Description = New System.Windows.Forms.Label()
        Me.FinishPage_Header = New System.Windows.Forms.Label()
        Me.AdvSettingsPage_SysprepUnatt_OFD = New System.Windows.Forms.OpenFileDialog()
        Me.ButtonPanel.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.BannerPanel.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PageContainerPanel.SuspendLayout()
        Me.WelcomePage.SuspendLayout()
        Me.SystemCheckPanel.SuspendLayout()
        Me.SysCheckPage_CCPProgressPanel.SuspendLayout()
        Me.SysCheckPage_CheckDetailsGB.SuspendLayout()
        Me.SysCheckPage_CheckDetailsTLP.SuspendLayout()
        Me.AdvSettingsPanel.SuspendLayout()
        Me.AdvSettingsPage_SysprepConfigPanel.SuspendLayout()
        Me.SettingPreparationPanel.SuspendLayout()
        Me.FinishPanel.SuspendLayout()
        Me.SuspendLayout()
        '
        'ButtonPanel
        '
        Me.ButtonPanel.BackColor = System.Drawing.Color.White
        Me.ButtonPanel.Controls.Add(Me.AboutBtn)
        Me.ButtonPanel.Controls.Add(Me.TableLayoutPanel1)
        Me.ButtonPanel.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ButtonPanel.Location = New System.Drawing.Point(0, 681)
        Me.ButtonPanel.Name = "ButtonPanel"
        Me.ButtonPanel.Size = New System.Drawing.Size(1008, 48)
        Me.ButtonPanel.TabIndex = 0
        '
        'AboutBtn
        '
        Me.AboutBtn.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.AboutBtn.Location = New System.Drawing.Point(12, 13)
        Me.AboutBtn.Name = "AboutBtn"
        Me.AboutBtn.Size = New System.Drawing.Size(75, 23)
        Me.AboutBtn.TabIndex = 3
        Me.AboutBtn.Text = "About..."
        Me.AboutBtn.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 3
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.Back_Button, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Next_Button, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 2, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(780, 10)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(216, 29)
        Me.TableLayoutPanel1.TabIndex = 2
        '
        'Back_Button
        '
        Me.Back_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Back_Button.Enabled = False
        Me.Back_Button.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Back_Button.Location = New System.Drawing.Point(4, 3)
        Me.Back_Button.Name = "Back_Button"
        Me.Back_Button.Size = New System.Drawing.Size(64, 23)
        Me.Back_Button.TabIndex = 0
        Me.Back_Button.Text = "Back"
        '
        'Next_Button
        '
        Me.Next_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Next_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Next_Button.Enabled = False
        Me.Next_Button.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Next_Button.Location = New System.Drawing.Point(76, 3)
        Me.Next_Button.Name = "Next_Button"
        Me.Next_Button.Size = New System.Drawing.Size(64, 23)
        Me.Next_Button.TabIndex = 1
        Me.Next_Button.Text = "Next"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Cancel_Button.Location = New System.Drawing.Point(148, 3)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(64, 23)
        Me.Cancel_Button.TabIndex = 1
        Me.Cancel_Button.Text = "Cancel"
        '
        'BannerPanel
        '
        Me.BannerPanel.BackColor = System.Drawing.Color.Transparent
        Me.BannerPanel.Controls.Add(Me.PictureBox1)
        Me.BannerPanel.Dock = System.Windows.Forms.DockStyle.Left
        Me.BannerPanel.Location = New System.Drawing.Point(0, 0)
        Me.BannerPanel.Name = "BannerPanel"
        Me.BannerPanel.Size = New System.Drawing.Size(120, 681)
        Me.BannerPanel.TabIndex = 1
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.SysprepPreparator.My.Resources.Resources.preparator_image_full
        Me.PictureBox1.Location = New System.Drawing.Point(12, 12)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(96, 96)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'PageContainerPanel
        '
        Me.PageContainerPanel.BackColor = System.Drawing.Color.WhiteSmoke
        Me.PageContainerPanel.Controls.Add(Me.WelcomePage)
        Me.PageContainerPanel.Controls.Add(Me.SystemCheckPanel)
        Me.PageContainerPanel.Controls.Add(Me.AdvSettingsPanel)
        Me.PageContainerPanel.Controls.Add(Me.SettingPreparationPanel)
        Me.PageContainerPanel.Controls.Add(Me.FinishPanel)
        Me.PageContainerPanel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PageContainerPanel.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PageContainerPanel.Location = New System.Drawing.Point(120, 0)
        Me.PageContainerPanel.Name = "PageContainerPanel"
        Me.PageContainerPanel.Size = New System.Drawing.Size(888, 681)
        Me.PageContainerPanel.TabIndex = 2
        '
        'WelcomePage
        '
        Me.WelcomePage.Controls.Add(Me.WelcomePage_Description)
        Me.WelcomePage.Controls.Add(Me.WelcomePage_Header)
        Me.WelcomePage.Dock = System.Windows.Forms.DockStyle.Fill
        Me.WelcomePage.Location = New System.Drawing.Point(0, 0)
        Me.WelcomePage.Name = "WelcomePage"
        Me.WelcomePage.Size = New System.Drawing.Size(888, 681)
        Me.WelcomePage.TabIndex = 0
        '
        'WelcomePage_Description
        '
        Me.WelcomePage_Description.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.WelcomePage_Description.AutoEllipsis = True
        Me.WelcomePage_Description.Location = New System.Drawing.Point(17, 64)
        Me.WelcomePage_Description.Name = "WelcomePage_Description"
        Me.WelcomePage_Description.Size = New System.Drawing.Size(857, 437)
        Me.WelcomePage_Description.TabIndex = 1
        Me.WelcomePage_Description.Text = "This wizard helps you prepare this computer for image capture. To begin, click Ne" &
    "xt."
        '
        'WelcomePage_Header
        '
        Me.WelcomePage_Header.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.WelcomePage_Header.AutoEllipsis = True
        Me.WelcomePage_Header.Font = New System.Drawing.Font("Segoe UI", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.WelcomePage_Header.Location = New System.Drawing.Point(14, 12)
        Me.WelcomePage_Header.Name = "WelcomePage_Header"
        Me.WelcomePage_Header.Size = New System.Drawing.Size(860, 45)
        Me.WelcomePage_Header.TabIndex = 0
        Me.WelcomePage_Header.Text = "Welcome to the wizard"
        '
        'SystemCheckPanel
        '
        Me.SystemCheckPanel.Controls.Add(Me.SysCheckPage_CCPProgressPanel)
        Me.SystemCheckPanel.Controls.Add(Me.SysCheckPage_CheckAgainBtn)
        Me.SystemCheckPanel.Controls.Add(Me.SysCheckPage_CheckDetailsGB)
        Me.SystemCheckPanel.Controls.Add(Me.SysCheckPage_ChecksLv)
        Me.SystemCheckPanel.Controls.Add(Me.SysCheckPage_Description)
        Me.SystemCheckPanel.Controls.Add(Me.SysCheckPage_Header)
        Me.SystemCheckPanel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SystemCheckPanel.Location = New System.Drawing.Point(0, 0)
        Me.SystemCheckPanel.Name = "SystemCheckPanel"
        Me.SystemCheckPanel.Size = New System.Drawing.Size(888, 681)
        Me.SystemCheckPanel.TabIndex = 1
        Me.SystemCheckPanel.Visible = False
        '
        'SysCheckPage_CCPProgressPanel
        '
        Me.SysCheckPage_CCPProgressPanel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.SysCheckPage_CCPProgressPanel.Controls.Add(Me.SysCheckPage_ProgressLabel)
        Me.SysCheckPage_CCPProgressPanel.Controls.Add(Me.SysCheckPage_CCPProgressBar)
        Me.SysCheckPage_CCPProgressPanel.Location = New System.Drawing.Point(23, 650)
        Me.SysCheckPage_CCPProgressPanel.Name = "SysCheckPage_CCPProgressPanel"
        Me.SysCheckPage_CCPProgressPanel.Size = New System.Drawing.Size(460, 23)
        Me.SysCheckPage_CCPProgressPanel.TabIndex = 5
        Me.SysCheckPage_CCPProgressPanel.Visible = False
        '
        'SysCheckPage_ProgressLabel
        '
        Me.SysCheckPage_ProgressLabel.AutoEllipsis = True
        Me.SysCheckPage_ProgressLabel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SysCheckPage_ProgressLabel.Location = New System.Drawing.Point(0, 0)
        Me.SysCheckPage_ProgressLabel.Name = "SysCheckPage_ProgressLabel"
        Me.SysCheckPage_ProgressLabel.Size = New System.Drawing.Size(135, 23)
        Me.SysCheckPage_ProgressLabel.TabIndex = 1
        Me.SysCheckPage_ProgressLabel.Text = "Please wait..."
        Me.SysCheckPage_ProgressLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'SysCheckPage_CCPProgressBar
        '
        Me.SysCheckPage_CCPProgressBar.Dock = System.Windows.Forms.DockStyle.Right
        Me.SysCheckPage_CCPProgressBar.Location = New System.Drawing.Point(135, 0)
        Me.SysCheckPage_CCPProgressBar.MarqueeAnimationSpeed = 25
        Me.SysCheckPage_CCPProgressBar.Name = "SysCheckPage_CCPProgressBar"
        Me.SysCheckPage_CCPProgressBar.Size = New System.Drawing.Size(325, 23)
        Me.SysCheckPage_CCPProgressBar.Style = System.Windows.Forms.ProgressBarStyle.Marquee
        Me.SysCheckPage_CCPProgressBar.TabIndex = 0
        '
        'SysCheckPage_CheckAgainBtn
        '
        Me.SysCheckPage_CheckAgainBtn.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SysCheckPage_CheckAgainBtn.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.SysCheckPage_CheckAgainBtn.Location = New System.Drawing.Point(664, 650)
        Me.SysCheckPage_CheckAgainBtn.Name = "SysCheckPage_CheckAgainBtn"
        Me.SysCheckPage_CheckAgainBtn.Size = New System.Drawing.Size(207, 23)
        Me.SysCheckPage_CheckAgainBtn.TabIndex = 4
        Me.SysCheckPage_CheckAgainBtn.Text = "Check again"
        Me.SysCheckPage_CheckAgainBtn.UseVisualStyleBackColor = True
        '
        'SysCheckPage_CheckDetailsGB
        '
        Me.SysCheckPage_CheckDetailsGB.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SysCheckPage_CheckDetailsGB.Controls.Add(Me.SysCheckPage_CheckDetailsTLP)
        Me.SysCheckPage_CheckDetailsGB.Location = New System.Drawing.Point(20, 293)
        Me.SysCheckPage_CheckDetailsGB.Name = "SysCheckPage_CheckDetailsGB"
        Me.SysCheckPage_CheckDetailsGB.Size = New System.Drawing.Size(852, 350)
        Me.SysCheckPage_CheckDetailsGB.TabIndex = 3
        Me.SysCheckPage_CheckDetailsGB.TabStop = False
        Me.SysCheckPage_CheckDetailsGB.Text = "Check Details"
        '
        'SysCheckPage_CheckDetailsTLP
        '
        Me.SysCheckPage_CheckDetailsTLP.ColumnCount = 2
        Me.SysCheckPage_CheckDetailsTLP.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333!))
        Me.SysCheckPage_CheckDetailsTLP.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 66.66667!))
        Me.SysCheckPage_CheckDetailsTLP.Controls.Add(Me.SysCheckPage_CheckDetails_Resolution, 0, 2)
        Me.SysCheckPage_CheckDetailsTLP.Controls.Add(Me.SysCheckPage_CheckDetails_DescriptionValue, 1, 1)
        Me.SysCheckPage_CheckDetailsTLP.Controls.Add(Me.SysCheckPage_CheckDetails_Description, 0, 1)
        Me.SysCheckPage_CheckDetailsTLP.Controls.Add(Me.SysCheckPage_CheckDetails_Title, 0, 0)
        Me.SysCheckPage_CheckDetailsTLP.Controls.Add(Me.SysCheckPage_CheckDetails_ResolutionValue, 1, 2)
        Me.SysCheckPage_CheckDetailsTLP.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SysCheckPage_CheckDetailsTLP.Location = New System.Drawing.Point(3, 19)
        Me.SysCheckPage_CheckDetailsTLP.Name = "SysCheckPage_CheckDetailsTLP"
        Me.SysCheckPage_CheckDetailsTLP.RowCount = 3
        Me.SysCheckPage_CheckDetailsTLP.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 32.0!))
        Me.SysCheckPage_CheckDetailsTLP.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 64.0!))
        Me.SysCheckPage_CheckDetailsTLP.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.SysCheckPage_CheckDetailsTLP.Size = New System.Drawing.Size(846, 328)
        Me.SysCheckPage_CheckDetailsTLP.TabIndex = 0
        Me.SysCheckPage_CheckDetailsTLP.Visible = False
        '
        'SysCheckPage_CheckDetails_Resolution
        '
        Me.SysCheckPage_CheckDetails_Resolution.AutoSize = True
        Me.SysCheckPage_CheckDetails_Resolution.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SysCheckPage_CheckDetails_Resolution.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SysCheckPage_CheckDetails_Resolution.Location = New System.Drawing.Point(3, 96)
        Me.SysCheckPage_CheckDetails_Resolution.Name = "SysCheckPage_CheckDetails_Resolution"
        Me.SysCheckPage_CheckDetails_Resolution.Size = New System.Drawing.Size(275, 232)
        Me.SysCheckPage_CheckDetails_Resolution.TabIndex = 4
        Me.SysCheckPage_CheckDetails_Resolution.Text = "Possible Resolution:"
        Me.SysCheckPage_CheckDetails_Resolution.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'SysCheckPage_CheckDetails_DescriptionValue
        '
        Me.SysCheckPage_CheckDetails_DescriptionValue.AutoEllipsis = True
        Me.SysCheckPage_CheckDetails_DescriptionValue.AutoSize = True
        Me.SysCheckPage_CheckDetails_DescriptionValue.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SysCheckPage_CheckDetails_DescriptionValue.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SysCheckPage_CheckDetails_DescriptionValue.Location = New System.Drawing.Point(284, 32)
        Me.SysCheckPage_CheckDetails_DescriptionValue.Name = "SysCheckPage_CheckDetails_DescriptionValue"
        Me.SysCheckPage_CheckDetails_DescriptionValue.Size = New System.Drawing.Size(559, 64)
        Me.SysCheckPage_CheckDetails_DescriptionValue.TabIndex = 2
        Me.SysCheckPage_CheckDetails_DescriptionValue.Text = "Description"
        '
        'SysCheckPage_CheckDetails_Description
        '
        Me.SysCheckPage_CheckDetails_Description.AutoSize = True
        Me.SysCheckPage_CheckDetails_Description.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SysCheckPage_CheckDetails_Description.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SysCheckPage_CheckDetails_Description.Location = New System.Drawing.Point(3, 32)
        Me.SysCheckPage_CheckDetails_Description.Name = "SysCheckPage_CheckDetails_Description"
        Me.SysCheckPage_CheckDetails_Description.Size = New System.Drawing.Size(275, 64)
        Me.SysCheckPage_CheckDetails_Description.TabIndex = 1
        Me.SysCheckPage_CheckDetails_Description.Text = "Description:"
        Me.SysCheckPage_CheckDetails_Description.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'SysCheckPage_CheckDetails_Title
        '
        Me.SysCheckPage_CheckDetails_Title.AutoEllipsis = True
        Me.SysCheckPage_CheckDetails_Title.AutoSize = True
        Me.SysCheckPage_CheckDetailsTLP.SetColumnSpan(Me.SysCheckPage_CheckDetails_Title, 2)
        Me.SysCheckPage_CheckDetails_Title.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SysCheckPage_CheckDetails_Title.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SysCheckPage_CheckDetails_Title.Location = New System.Drawing.Point(8, 0)
        Me.SysCheckPage_CheckDetails_Title.Margin = New System.Windows.Forms.Padding(8, 0, 3, 0)
        Me.SysCheckPage_CheckDetails_Title.Name = "SysCheckPage_CheckDetails_Title"
        Me.SysCheckPage_CheckDetails_Title.Size = New System.Drawing.Size(835, 32)
        Me.SysCheckPage_CheckDetails_Title.TabIndex = 0
        Me.SysCheckPage_CheckDetails_Title.Text = "Check Title"
        Me.SysCheckPage_CheckDetails_Title.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'SysCheckPage_CheckDetails_ResolutionValue
        '
        Me.SysCheckPage_CheckDetails_ResolutionValue.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.SysCheckPage_CheckDetails_ResolutionValue.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SysCheckPage_CheckDetails_ResolutionValue.Location = New System.Drawing.Point(284, 96)
        Me.SysCheckPage_CheckDetails_ResolutionValue.Margin = New System.Windows.Forms.Padding(3, 0, 3, 0)
        Me.SysCheckPage_CheckDetails_ResolutionValue.Multiline = True
        Me.SysCheckPage_CheckDetails_ResolutionValue.Name = "SysCheckPage_CheckDetails_ResolutionValue"
        Me.SysCheckPage_CheckDetails_ResolutionValue.ReadOnly = True
        Me.SysCheckPage_CheckDetails_ResolutionValue.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.SysCheckPage_CheckDetails_ResolutionValue.Size = New System.Drawing.Size(559, 232)
        Me.SysCheckPage_CheckDetails_ResolutionValue.TabIndex = 5
        '
        'SysCheckPage_ChecksLv
        '
        Me.SysCheckPage_ChecksLv.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SysCheckPage_ChecksLv.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.SysCheckPage_CheckCH, Me.SysCheckPage_CompatibleCH, Me.SysCheckPage_SeverityCH})
        Me.SysCheckPage_ChecksLv.FullRowSelect = True
        Me.SysCheckPage_ChecksLv.HideSelection = False
        Me.SysCheckPage_ChecksLv.Location = New System.Drawing.Point(20, 112)
        Me.SysCheckPage_ChecksLv.MultiSelect = False
        Me.SysCheckPage_ChecksLv.Name = "SysCheckPage_ChecksLv"
        Me.SysCheckPage_ChecksLv.Size = New System.Drawing.Size(852, 175)
        Me.SysCheckPage_ChecksLv.TabIndex = 2
        Me.SysCheckPage_ChecksLv.UseCompatibleStateImageBehavior = False
        Me.SysCheckPage_ChecksLv.View = System.Windows.Forms.View.Details
        '
        'SysCheckPage_CheckCH
        '
        Me.SysCheckPage_CheckCH.Text = "Check"
        Me.SysCheckPage_CheckCH.Width = 376
        '
        'SysCheckPage_CompatibleCH
        '
        Me.SysCheckPage_CompatibleCH.Text = "Compatible?"
        Me.SysCheckPage_CompatibleCH.Width = 96
        '
        'SysCheckPage_SeverityCH
        '
        Me.SysCheckPage_SeverityCH.Text = "Severity"
        Me.SysCheckPage_SeverityCH.Width = 128
        '
        'SysCheckPage_Description
        '
        Me.SysCheckPage_Description.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SysCheckPage_Description.AutoEllipsis = True
        Me.SysCheckPage_Description.Location = New System.Drawing.Point(17, 64)
        Me.SysCheckPage_Description.Name = "SysCheckPage_Description"
        Me.SysCheckPage_Description.Size = New System.Drawing.Size(857, 212)
        Me.SysCheckPage_Description.TabIndex = 1
        Me.SysCheckPage_Description.Text = "We are performing some checks on your computer to determine if it is ready for im" &
    "age capture. This will take a couple of moments, after which you'll see results " &
    "below:"
        '
        'SysCheckPage_Header
        '
        Me.SysCheckPage_Header.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SysCheckPage_Header.AutoEllipsis = True
        Me.SysCheckPage_Header.Font = New System.Drawing.Font("Segoe UI", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SysCheckPage_Header.Location = New System.Drawing.Point(14, 12)
        Me.SysCheckPage_Header.Name = "SysCheckPage_Header"
        Me.SysCheckPage_Header.Size = New System.Drawing.Size(860, 45)
        Me.SysCheckPage_Header.TabIndex = 0
        Me.SysCheckPage_Header.Text = "System Checks"
        '
        'AdvSettingsPanel
        '
        Me.AdvSettingsPanel.Controls.Add(Me.AdvSettingsPage_SysprepPrepToolDeploySteps)
        Me.AdvSettingsPanel.Controls.Add(Me.AdvSettingsPage_SysprepConfigPanel)
        Me.AdvSettingsPanel.Controls.Add(Me.AdvSettingsPage_ConfigSysprepSettings)
        Me.AdvSettingsPanel.Controls.Add(Me.AdvSettingsPage_Description)
        Me.AdvSettingsPanel.Controls.Add(Me.AdvSettingsPage_Header)
        Me.AdvSettingsPanel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.AdvSettingsPanel.Location = New System.Drawing.Point(0, 0)
        Me.AdvSettingsPanel.Name = "AdvSettingsPanel"
        Me.AdvSettingsPanel.Size = New System.Drawing.Size(888, 681)
        Me.AdvSettingsPanel.TabIndex = 2
        Me.AdvSettingsPanel.Visible = False
        '
        'AdvSettingsPage_SysprepPrepToolDeploySteps
        '
        Me.AdvSettingsPage_SysprepPrepToolDeploySteps.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.AdvSettingsPage_SysprepPrepToolDeploySteps.AutoEllipsis = True
        Me.AdvSettingsPage_SysprepPrepToolDeploySteps.Location = New System.Drawing.Point(20, 312)
        Me.AdvSettingsPage_SysprepPrepToolDeploySteps.Name = "AdvSettingsPage_SysprepPrepToolDeploySteps"
        Me.AdvSettingsPage_SysprepPrepToolDeploySteps.Size = New System.Drawing.Size(852, 346)
        Me.AdvSettingsPage_SysprepPrepToolDeploySteps.TabIndex = 4
        Me.AdvSettingsPage_SysprepPrepToolDeploySteps.Text = resources.GetString("AdvSettingsPage_SysprepPrepToolDeploySteps.Text")
        '
        'AdvSettingsPage_SysprepConfigPanel
        '
        Me.AdvSettingsPage_SysprepConfigPanel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.AdvSettingsPage_SysprepConfigPanel.Controls.Add(Me.AdvSettingsPage_VMModeExplanationLink)
        Me.AdvSettingsPage_SysprepConfigPanel.Controls.Add(Me.AdvSettingsPage_VMMode)
        Me.AdvSettingsPage_SysprepConfigPanel.Controls.Add(Me.AdvSettingsPage_SysprepUnatt_Btn)
        Me.AdvSettingsPage_SysprepConfigPanel.Controls.Add(Me.AdvSettingsPage_SysprepUnatt_AnswerFileText)
        Me.AdvSettingsPage_SysprepConfigPanel.Controls.Add(Me.AdvSettingsPage_SysprepUnattLabel)
        Me.AdvSettingsPage_SysprepConfigPanel.Controls.Add(Me.AdvSettingsPage_CleanupAction_Generalize)
        Me.AdvSettingsPage_SysprepConfigPanel.Controls.Add(Me.AdvSettingsPage_ShutdownOptionsCBox)
        Me.AdvSettingsPage_SysprepConfigPanel.Controls.Add(Me.AdvSettingsPage_CleanupActionCBox)
        Me.AdvSettingsPage_SysprepConfigPanel.Controls.Add(Me.AdvSettingsPage_ShutdownOptionLabel)
        Me.AdvSettingsPage_SysprepConfigPanel.Controls.Add(Me.AdvSettingsPage_CleanupActionLabel)
        Me.AdvSettingsPage_SysprepConfigPanel.Enabled = False
        Me.AdvSettingsPage_SysprepConfigPanel.Location = New System.Drawing.Point(39, 149)
        Me.AdvSettingsPage_SysprepConfigPanel.Name = "AdvSettingsPage_SysprepConfigPanel"
        Me.AdvSettingsPage_SysprepConfigPanel.Size = New System.Drawing.Size(834, 157)
        Me.AdvSettingsPage_SysprepConfigPanel.TabIndex = 3
        '
        'AdvSettingsPage_VMModeExplanationLink
        '
        Me.AdvSettingsPage_VMModeExplanationLink.AutoSize = True
        Me.AdvSettingsPage_VMModeExplanationLink.LinkBehavior = System.Windows.Forms.LinkBehavior.NeverUnderline
        Me.AdvSettingsPage_VMModeExplanationLink.LinkColor = System.Drawing.Color.DodgerBlue
        Me.AdvSettingsPage_VMModeExplanationLink.Location = New System.Drawing.Point(38, 127)
        Me.AdvSettingsPage_VMModeExplanationLink.Name = "AdvSettingsPage_VMModeExplanationLink"
        Me.AdvSettingsPage_VMModeExplanationLink.Size = New System.Drawing.Size(200, 15)
        Me.AdvSettingsPage_VMModeExplanationLink.TabIndex = 7
        Me.AdvSettingsPage_VMModeExplanationLink.TabStop = True
        Me.AdvSettingsPage_VMModeExplanationLink.Text = "Learn More about Sysprep VM Mode"
        '
        'AdvSettingsPage_VMMode
        '
        Me.AdvSettingsPage_VMMode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.AdvSettingsPage_VMMode.AutoCheck = False
        Me.AdvSettingsPage_VMMode.AutoSize = True
        Me.AdvSettingsPage_VMMode.Location = New System.Drawing.Point(23, 105)
        Me.AdvSettingsPage_VMMode.Name = "AdvSettingsPage_VMMode"
        Me.AdvSettingsPage_VMMode.Size = New System.Drawing.Size(465, 19)
        Me.AdvSettingsPage_VMMode.TabIndex = 6
        Me.AdvSettingsPage_VMMode.Text = "VM Mode: Generalize the system to be deployed as a VHD on the same VM solution"
        Me.AdvSettingsPage_VMMode.UseVisualStyleBackColor = True
        '
        'AdvSettingsPage_SysprepUnatt_Btn
        '
        Me.AdvSettingsPage_SysprepUnatt_Btn.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.AdvSettingsPage_SysprepUnatt_Btn.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.AdvSettingsPage_SysprepUnatt_Btn.Location = New System.Drawing.Point(695, 72)
        Me.AdvSettingsPage_SysprepUnatt_Btn.Name = "AdvSettingsPage_SysprepUnatt_Btn"
        Me.AdvSettingsPage_SysprepUnatt_Btn.Size = New System.Drawing.Size(75, 23)
        Me.AdvSettingsPage_SysprepUnatt_Btn.TabIndex = 5
        Me.AdvSettingsPage_SysprepUnatt_Btn.Text = "Browse..."
        Me.AdvSettingsPage_SysprepUnatt_Btn.UseVisualStyleBackColor = True
        '
        'AdvSettingsPage_SysprepUnatt_AnswerFileText
        '
        Me.AdvSettingsPage_SysprepUnatt_AnswerFileText.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.AdvSettingsPage_SysprepUnatt_AnswerFileText.Location = New System.Drawing.Point(180, 72)
        Me.AdvSettingsPage_SysprepUnatt_AnswerFileText.Name = "AdvSettingsPage_SysprepUnatt_AnswerFileText"
        Me.AdvSettingsPage_SysprepUnatt_AnswerFileText.Size = New System.Drawing.Size(509, 23)
        Me.AdvSettingsPage_SysprepUnatt_AnswerFileText.TabIndex = 4
        '
        'AdvSettingsPage_SysprepUnattLabel
        '
        Me.AdvSettingsPage_SysprepUnattLabel.AutoSize = True
        Me.AdvSettingsPage_SysprepUnattLabel.Location = New System.Drawing.Point(20, 75)
        Me.AdvSettingsPage_SysprepUnattLabel.Name = "AdvSettingsPage_SysprepUnattLabel"
        Me.AdvSettingsPage_SysprepUnattLabel.Size = New System.Drawing.Size(70, 15)
        Me.AdvSettingsPage_SysprepUnattLabel.TabIndex = 3
        Me.AdvSettingsPage_SysprepUnattLabel.Text = "Answer File:"
        '
        'AdvSettingsPage_CleanupAction_Generalize
        '
        Me.AdvSettingsPage_CleanupAction_Generalize.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.AdvSettingsPage_CleanupAction_Generalize.AutoEllipsis = True
        Me.AdvSettingsPage_CleanupAction_Generalize.Checked = True
        Me.AdvSettingsPage_CleanupAction_Generalize.CheckState = System.Windows.Forms.CheckState.Checked
        Me.AdvSettingsPage_CleanupAction_Generalize.Location = New System.Drawing.Point(697, 16)
        Me.AdvSettingsPage_CleanupAction_Generalize.Name = "AdvSettingsPage_CleanupAction_Generalize"
        Me.AdvSettingsPage_CleanupAction_Generalize.Size = New System.Drawing.Size(114, 19)
        Me.AdvSettingsPage_CleanupAction_Generalize.TabIndex = 2
        Me.AdvSettingsPage_CleanupAction_Generalize.Text = "Generalize"
        Me.AdvSettingsPage_CleanupAction_Generalize.UseVisualStyleBackColor = True
        '
        'AdvSettingsPage_ShutdownOptionsCBox
        '
        Me.AdvSettingsPage_ShutdownOptionsCBox.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.AdvSettingsPage_ShutdownOptionsCBox.FormattingEnabled = True
        Me.AdvSettingsPage_ShutdownOptionsCBox.Location = New System.Drawing.Point(180, 43)
        Me.AdvSettingsPage_ShutdownOptionsCBox.Name = "AdvSettingsPage_ShutdownOptionsCBox"
        Me.AdvSettingsPage_ShutdownOptionsCBox.Size = New System.Drawing.Size(509, 23)
        Me.AdvSettingsPage_ShutdownOptionsCBox.TabIndex = 1
        '
        'AdvSettingsPage_CleanupActionCBox
        '
        Me.AdvSettingsPage_CleanupActionCBox.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.AdvSettingsPage_CleanupActionCBox.FormattingEnabled = True
        Me.AdvSettingsPage_CleanupActionCBox.Location = New System.Drawing.Point(180, 14)
        Me.AdvSettingsPage_CleanupActionCBox.Name = "AdvSettingsPage_CleanupActionCBox"
        Me.AdvSettingsPage_CleanupActionCBox.Size = New System.Drawing.Size(509, 23)
        Me.AdvSettingsPage_CleanupActionCBox.TabIndex = 1
        '
        'AdvSettingsPage_ShutdownOptionLabel
        '
        Me.AdvSettingsPage_ShutdownOptionLabel.AutoSize = True
        Me.AdvSettingsPage_ShutdownOptionLabel.Location = New System.Drawing.Point(20, 46)
        Me.AdvSettingsPage_ShutdownOptionLabel.Name = "AdvSettingsPage_ShutdownOptionLabel"
        Me.AdvSettingsPage_ShutdownOptionLabel.Size = New System.Drawing.Size(104, 15)
        Me.AdvSettingsPage_ShutdownOptionLabel.TabIndex = 0
        Me.AdvSettingsPage_ShutdownOptionLabel.Text = "Shutdown Option:"
        '
        'AdvSettingsPage_CleanupActionLabel
        '
        Me.AdvSettingsPage_CleanupActionLabel.AutoSize = True
        Me.AdvSettingsPage_CleanupActionLabel.Location = New System.Drawing.Point(20, 17)
        Me.AdvSettingsPage_CleanupActionLabel.Name = "AdvSettingsPage_CleanupActionLabel"
        Me.AdvSettingsPage_CleanupActionLabel.Size = New System.Drawing.Size(92, 15)
        Me.AdvSettingsPage_CleanupActionLabel.TabIndex = 0
        Me.AdvSettingsPage_CleanupActionLabel.Text = "Cleanup Action:"
        '
        'AdvSettingsPage_ConfigSysprepSettings
        '
        Me.AdvSettingsPage_ConfigSysprepSettings.AutoSize = True
        Me.AdvSettingsPage_ConfigSysprepSettings.Location = New System.Drawing.Point(20, 123)
        Me.AdvSettingsPage_ConfigSysprepSettings.Name = "AdvSettingsPage_ConfigSysprepSettings"
        Me.AdvSettingsPage_ConfigSysprepSettings.Size = New System.Drawing.Size(167, 19)
        Me.AdvSettingsPage_ConfigSysprepSettings.TabIndex = 2
        Me.AdvSettingsPage_ConfigSysprepSettings.Text = "Configure Sysprep settings"
        Me.AdvSettingsPage_ConfigSysprepSettings.UseVisualStyleBackColor = True
        '
        'AdvSettingsPage_Description
        '
        Me.AdvSettingsPage_Description.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.AdvSettingsPage_Description.AutoEllipsis = True
        Me.AdvSettingsPage_Description.Location = New System.Drawing.Point(17, 64)
        Me.AdvSettingsPage_Description.Name = "AdvSettingsPage_Description"
        Me.AdvSettingsPage_Description.Size = New System.Drawing.Size(857, 212)
        Me.AdvSettingsPage_Description.TabIndex = 1
        Me.AdvSettingsPage_Description.Text = "You can specify additional options for the Sysprep utility. Normally, you don't n" &
    "eed to configure advanced settings."
        '
        'AdvSettingsPage_Header
        '
        Me.AdvSettingsPage_Header.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.AdvSettingsPage_Header.AutoEllipsis = True
        Me.AdvSettingsPage_Header.Font = New System.Drawing.Font("Segoe UI", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AdvSettingsPage_Header.Location = New System.Drawing.Point(14, 12)
        Me.AdvSettingsPage_Header.Name = "AdvSettingsPage_Header"
        Me.AdvSettingsPage_Header.Size = New System.Drawing.Size(860, 45)
        Me.AdvSettingsPage_Header.TabIndex = 0
        Me.AdvSettingsPage_Header.Text = "Advanced Preparation Settings"
        '
        'SettingPreparationPanel
        '
        Me.SettingPreparationPanel.Controls.Add(Me.SettingPreparationPanel_TaskProgressBar)
        Me.SettingPreparationPanel.Controls.Add(Me.SettingPreparationPanel_TaskLv)
        Me.SettingPreparationPanel.Controls.Add(Me.SettingPreparationPage_SubProcessProgressLabel)
        Me.SettingPreparationPanel.Controls.Add(Me.SettingPreparationPage_ProgressLabel)
        Me.SettingPreparationPanel.Controls.Add(Me.SettingPreparationPage_Description)
        Me.SettingPreparationPanel.Controls.Add(Me.SettingPreparationPage_Header)
        Me.SettingPreparationPanel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SettingPreparationPanel.Location = New System.Drawing.Point(0, 0)
        Me.SettingPreparationPanel.Name = "SettingPreparationPanel"
        Me.SettingPreparationPanel.Size = New System.Drawing.Size(888, 681)
        Me.SettingPreparationPanel.TabIndex = 3
        Me.SettingPreparationPanel.Visible = False
        '
        'SettingPreparationPanel_TaskProgressBar
        '
        Me.SettingPreparationPanel_TaskProgressBar.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SettingPreparationPanel_TaskProgressBar.Location = New System.Drawing.Point(18, 187)
        Me.SettingPreparationPanel_TaskProgressBar.Name = "SettingPreparationPanel_TaskProgressBar"
        Me.SettingPreparationPanel_TaskProgressBar.Size = New System.Drawing.Size(852, 23)
        Me.SettingPreparationPanel_TaskProgressBar.TabIndex = 6
        '
        'SettingPreparationPanel_TaskLv
        '
        Me.SettingPreparationPanel_TaskLv.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SettingPreparationPanel_TaskLv.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.SettingPreparationPage_TaskCH, Me.SettingPreparationPage_SuccessfulCH})
        Me.SettingPreparationPanel_TaskLv.FullRowSelect = True
        Me.SettingPreparationPanel_TaskLv.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.SettingPreparationPanel_TaskLv.HideSelection = False
        Me.SettingPreparationPanel_TaskLv.Location = New System.Drawing.Point(18, 216)
        Me.SettingPreparationPanel_TaskLv.MultiSelect = False
        Me.SettingPreparationPanel_TaskLv.Name = "SettingPreparationPanel_TaskLv"
        Me.SettingPreparationPanel_TaskLv.Size = New System.Drawing.Size(852, 343)
        Me.SettingPreparationPanel_TaskLv.TabIndex = 5
        Me.SettingPreparationPanel_TaskLv.UseCompatibleStateImageBehavior = False
        Me.SettingPreparationPanel_TaskLv.View = System.Windows.Forms.View.Details
        '
        'SettingPreparationPage_TaskCH
        '
        Me.SettingPreparationPage_TaskCH.Text = "Task"
        Me.SettingPreparationPage_TaskCH.Width = 461
        '
        'SettingPreparationPage_SuccessfulCH
        '
        Me.SettingPreparationPage_SuccessfulCH.Text = "Successful?"
        Me.SettingPreparationPage_SuccessfulCH.Width = 128
        '
        'SettingPreparationPage_SubProcessProgressLabel
        '
        Me.SettingPreparationPage_SubProcessProgressLabel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SettingPreparationPage_SubProcessProgressLabel.AutoEllipsis = True
        Me.SettingPreparationPage_SubProcessProgressLabel.Location = New System.Drawing.Point(39, 150)
        Me.SettingPreparationPage_SubProcessProgressLabel.Name = "SettingPreparationPage_SubProcessProgressLabel"
        Me.SettingPreparationPage_SubProcessProgressLabel.Size = New System.Drawing.Size(831, 26)
        Me.SettingPreparationPage_SubProcessProgressLabel.TabIndex = 4
        '
        'SettingPreparationPage_ProgressLabel
        '
        Me.SettingPreparationPage_ProgressLabel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SettingPreparationPage_ProgressLabel.AutoEllipsis = True
        Me.SettingPreparationPage_ProgressLabel.Location = New System.Drawing.Point(17, 124)
        Me.SettingPreparationPage_ProgressLabel.Name = "SettingPreparationPage_ProgressLabel"
        Me.SettingPreparationPage_ProgressLabel.Size = New System.Drawing.Size(852, 26)
        Me.SettingPreparationPage_ProgressLabel.TabIndex = 4
        Me.SettingPreparationPage_ProgressLabel.Text = "Progress:"
        '
        'SettingPreparationPage_Description
        '
        Me.SettingPreparationPage_Description.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SettingPreparationPage_Description.AutoEllipsis = True
        Me.SettingPreparationPage_Description.Location = New System.Drawing.Point(17, 64)
        Me.SettingPreparationPage_Description.Name = "SettingPreparationPage_Description"
        Me.SettingPreparationPage_Description.Size = New System.Drawing.Size(857, 45)
        Me.SettingPreparationPage_Description.TabIndex = 1
        Me.SettingPreparationPage_Description.Text = "Please wait while we prepare your computer for Sysprep. This will take some time." &
    ""
        '
        'SettingPreparationPage_Header
        '
        Me.SettingPreparationPage_Header.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SettingPreparationPage_Header.AutoEllipsis = True
        Me.SettingPreparationPage_Header.Font = New System.Drawing.Font("Segoe UI", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SettingPreparationPage_Header.Location = New System.Drawing.Point(14, 12)
        Me.SettingPreparationPage_Header.Name = "SettingPreparationPage_Header"
        Me.SettingPreparationPage_Header.Size = New System.Drawing.Size(860, 45)
        Me.SettingPreparationPage_Header.TabIndex = 0
        Me.SettingPreparationPage_Header.Text = "Preparing your computer..."
        '
        'FinishPanel
        '
        Me.FinishPanel.Controls.Add(Me.FinishPage_RestartBtn)
        Me.FinishPanel.Controls.Add(Me.FinishPage_ResysprepBtn)
        Me.FinishPanel.Controls.Add(Me.FinishPage_CloseBtn)
        Me.FinishPanel.Controls.Add(Me.FinishPage_Description)
        Me.FinishPanel.Controls.Add(Me.FinishPage_Header)
        Me.FinishPanel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FinishPanel.Location = New System.Drawing.Point(0, 0)
        Me.FinishPanel.Name = "FinishPanel"
        Me.FinishPanel.Size = New System.Drawing.Size(888, 681)
        Me.FinishPanel.TabIndex = 4
        Me.FinishPanel.Visible = False
        '
        'FinishPage_RestartBtn
        '
        Me.FinishPage_RestartBtn.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FinishPage_RestartBtn.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.FinishPage_RestartBtn.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FinishPage_RestartBtn.Location = New System.Drawing.Point(578, 312)
        Me.FinishPage_RestartBtn.Name = "FinishPage_RestartBtn"
        Me.FinishPage_RestartBtn.Size = New System.Drawing.Size(256, 48)
        Me.FinishPage_RestartBtn.TabIndex = 2
        Me.FinishPage_RestartBtn.Text = "Restart PC"
        Me.FinishPage_RestartBtn.UseVisualStyleBackColor = True
        '
        'FinishPage_ResysprepBtn
        '
        Me.FinishPage_ResysprepBtn.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FinishPage_ResysprepBtn.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.FinishPage_ResysprepBtn.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FinishPage_ResysprepBtn.Location = New System.Drawing.Point(316, 312)
        Me.FinishPage_ResysprepBtn.Name = "FinishPage_ResysprepBtn"
        Me.FinishPage_ResysprepBtn.Size = New System.Drawing.Size(256, 48)
        Me.FinishPage_ResysprepBtn.TabIndex = 2
        Me.FinishPage_ResysprepBtn.Text = "Relaunch Sysprep"
        Me.FinishPage_ResysprepBtn.UseVisualStyleBackColor = True
        '
        'FinishPage_CloseBtn
        '
        Me.FinishPage_CloseBtn.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.FinishPage_CloseBtn.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.FinishPage_CloseBtn.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FinishPage_CloseBtn.Location = New System.Drawing.Point(54, 312)
        Me.FinishPage_CloseBtn.Name = "FinishPage_CloseBtn"
        Me.FinishPage_CloseBtn.Size = New System.Drawing.Size(256, 48)
        Me.FinishPage_CloseBtn.TabIndex = 2
        Me.FinishPage_CloseBtn.Text = "Close wizard"
        Me.FinishPage_CloseBtn.UseVisualStyleBackColor = True
        '
        'FinishPage_Description
        '
        Me.FinishPage_Description.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FinishPage_Description.AutoEllipsis = True
        Me.FinishPage_Description.Location = New System.Drawing.Point(17, 64)
        Me.FinishPage_Description.Name = "FinishPage_Description"
        Me.FinishPage_Description.Size = New System.Drawing.Size(857, 212)
        Me.FinishPage_Description.TabIndex = 1
        Me.FinishPage_Description.Text = resources.GetString("FinishPage_Description.Text")
        '
        'FinishPage_Header
        '
        Me.FinishPage_Header.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FinishPage_Header.AutoEllipsis = True
        Me.FinishPage_Header.Font = New System.Drawing.Font("Segoe UI", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FinishPage_Header.Location = New System.Drawing.Point(14, 12)
        Me.FinishPage_Header.Name = "FinishPage_Header"
        Me.FinishPage_Header.Size = New System.Drawing.Size(860, 45)
        Me.FinishPage_Header.TabIndex = 0
        Me.FinishPage_Header.Text = "Finishing Preparation..."
        '
        'AdvSettingsPage_SysprepUnatt_OFD
        '
        Me.AdvSettingsPage_SysprepUnatt_OFD.Filter = "Unattended Answer Files|*.xml"
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackgroundImage = Global.SysprepPreparator.My.Resources.Resources.wizard_background
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(1008, 729)
        Me.Controls.Add(Me.PageContainerPanel)
        Me.Controls.Add(Me.BannerPanel)
        Me.Controls.Add(Me.ButtonPanel)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1024, 768)
        Me.Name = "MainForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Sysprep Preparation Wizard"
        Me.ButtonPanel.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.BannerPanel.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PageContainerPanel.ResumeLayout(False)
        Me.WelcomePage.ResumeLayout(False)
        Me.SystemCheckPanel.ResumeLayout(False)
        Me.SysCheckPage_CCPProgressPanel.ResumeLayout(False)
        Me.SysCheckPage_CheckDetailsGB.ResumeLayout(False)
        Me.SysCheckPage_CheckDetailsTLP.ResumeLayout(False)
        Me.SysCheckPage_CheckDetailsTLP.PerformLayout()
        Me.AdvSettingsPanel.ResumeLayout(False)
        Me.AdvSettingsPanel.PerformLayout()
        Me.AdvSettingsPage_SysprepConfigPanel.ResumeLayout(False)
        Me.AdvSettingsPage_SysprepConfigPanel.PerformLayout()
        Me.SettingPreparationPanel.ResumeLayout(False)
        Me.FinishPanel.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ButtonPanel As System.Windows.Forms.Panel
    Friend WithEvents BannerPanel As System.Windows.Forms.Panel
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents PageContainerPanel As System.Windows.Forms.Panel
    Friend WithEvents WelcomePage As System.Windows.Forms.Panel
    Friend WithEvents WelcomePage_Header As System.Windows.Forms.Label
    Friend WithEvents WelcomePage_Description As System.Windows.Forms.Label
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Back_Button As System.Windows.Forms.Button
    Friend WithEvents Next_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents SystemCheckPanel As System.Windows.Forms.Panel
    Friend WithEvents SysCheckPage_Description As System.Windows.Forms.Label
    Friend WithEvents SysCheckPage_Header As System.Windows.Forms.Label
    Friend WithEvents SysCheckPage_CheckAgainBtn As System.Windows.Forms.Button
    Friend WithEvents SysCheckPage_CheckDetailsGB As System.Windows.Forms.GroupBox
    Friend WithEvents SysCheckPage_ChecksLv As System.Windows.Forms.ListView
    Friend WithEvents SysCheckPage_CheckCH As System.Windows.Forms.ColumnHeader
    Friend WithEvents SysCheckPage_SeverityCH As System.Windows.Forms.ColumnHeader
    Friend WithEvents SysCheckPage_CompatibleCH As System.Windows.Forms.ColumnHeader
    Friend WithEvents SysCheckPage_CheckDetailsTLP As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents SysCheckPage_CheckDetails_Resolution As System.Windows.Forms.Label
    Friend WithEvents SysCheckPage_CheckDetails_DescriptionValue As System.Windows.Forms.Label
    Friend WithEvents SysCheckPage_CheckDetails_Description As System.Windows.Forms.Label
    Friend WithEvents SysCheckPage_CheckDetails_Title As System.Windows.Forms.Label
    Friend WithEvents AdvSettingsPanel As System.Windows.Forms.Panel
    Friend WithEvents AdvSettingsPage_Description As System.Windows.Forms.Label
    Friend WithEvents AdvSettingsPage_Header As System.Windows.Forms.Label
    Friend WithEvents AdvSettingsPage_SysprepConfigPanel As System.Windows.Forms.Panel
    Friend WithEvents AdvSettingsPage_VMMode As System.Windows.Forms.CheckBox
    Friend WithEvents AdvSettingsPage_SysprepUnatt_Btn As System.Windows.Forms.Button
    Friend WithEvents AdvSettingsPage_SysprepUnatt_AnswerFileText As System.Windows.Forms.TextBox
    Friend WithEvents AdvSettingsPage_SysprepUnattLabel As System.Windows.Forms.Label
    Friend WithEvents AdvSettingsPage_CleanupAction_Generalize As System.Windows.Forms.CheckBox
    Friend WithEvents AdvSettingsPage_ShutdownOptionsCBox As System.Windows.Forms.ComboBox
    Friend WithEvents AdvSettingsPage_CleanupActionCBox As System.Windows.Forms.ComboBox
    Friend WithEvents AdvSettingsPage_ShutdownOptionLabel As System.Windows.Forms.Label
    Friend WithEvents AdvSettingsPage_CleanupActionLabel As System.Windows.Forms.Label
    Friend WithEvents AdvSettingsPage_ConfigSysprepSettings As System.Windows.Forms.CheckBox
    Friend WithEvents AdvSettingsPage_SysprepPrepToolDeploySteps As System.Windows.Forms.Label
    Friend WithEvents AdvSettingsPage_VMModeExplanationLink As System.Windows.Forms.LinkLabel
    Friend WithEvents AdvSettingsPage_SysprepUnatt_OFD As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SettingPreparationPanel As System.Windows.Forms.Panel
    Friend WithEvents SettingPreparationPage_ProgressLabel As System.Windows.Forms.Label
    Friend WithEvents SettingPreparationPage_Description As System.Windows.Forms.Label
    Friend WithEvents SettingPreparationPage_Header As System.Windows.Forms.Label
    Friend WithEvents SettingPreparationPanel_TaskLv As System.Windows.Forms.ListView
    Friend WithEvents SettingPreparationPage_TaskCH As System.Windows.Forms.ColumnHeader
    Friend WithEvents SettingPreparationPage_SuccessfulCH As System.Windows.Forms.ColumnHeader
    Friend WithEvents FinishPanel As System.Windows.Forms.Panel
    Friend WithEvents FinishPage_Description As System.Windows.Forms.Label
    Friend WithEvents FinishPage_Header As System.Windows.Forms.Label
    Friend WithEvents FinishPage_ResysprepBtn As System.Windows.Forms.Button
    Friend WithEvents FinishPage_RestartBtn As System.Windows.Forms.Button
    Friend WithEvents FinishPage_CloseBtn As System.Windows.Forms.Button
    Friend WithEvents SysCheckPage_CheckDetails_ResolutionValue As System.Windows.Forms.TextBox
    Friend WithEvents AboutBtn As Button
    Friend WithEvents SettingPreparationPage_SubProcessProgressLabel As Label
    Friend WithEvents SettingPreparationPanel_TaskProgressBar As ProgressBar
    Friend WithEvents SysCheckPage_CCPProgressPanel As Panel
    Friend WithEvents SysCheckPage_ProgressLabel As Label
    Friend WithEvents SysCheckPage_CCPProgressBar As ProgressBar
End Class
