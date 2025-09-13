Imports System.IO
Imports Microsoft.VisualBasic.ControlChars
Imports Microsoft.Win32
Imports SysprepPreparator.Classes
Imports SysprepPreparator.Helpers

Public Class MainForm

    ''' <summary>
    ''' The current page in the wizard
    ''' </summary>
    ''' <remarks></remarks>
    Dim CurrentWizardPage As New WizardPage()
    ''' <summary>
    ''' The pages in which the verification system should run checks
    ''' </summary>
    ''' <remarks></remarks>
    Dim VerifyInPages As New List(Of WizardPage.Page) From {
        WizardPage.Page.SysCheckPage
    }
    ''' <summary>
    ''' Determines whether the CCPs have been run to avoid running them again
    ''' </summary>
    ''' <remarks></remarks>
    Dim IsComputerChecked As Boolean = False

    ''' <summary>
    ''' The list of all the performed CCPs
    ''' </summary>
    ''' <remarks></remarks>
    Dim PerformedChecks As New List(Of CompatibilityCheckerProviderStatus)

    ''' <summary>
    ''' The current Sysprep configuration
    ''' </summary>
    ''' <remarks></remarks>
    Dim SysprepConfiguration As New SysprepConfig()

    ''' <summary>
    ''' This event is raised when a Preparation Task is about to start
    ''' </summary>
    ''' <param name="taskName">The name of the task</param>
    ''' <remarks></remarks>
    Public Event TaskStarted(taskName As String)

    ''' <summary>
    ''' This event is raised when a Preparation Task has finished and reporting will occur
    ''' </summary>
    ''' <param name="Task">The reported item featuring the task name and whether the PT succeeded</param>
    ''' <remarks></remarks>
    Public Event TaskReported(Task As Dictionary(Of String, Boolean))

    ''' <summary>
    ''' Handles the TaskStarted event
    ''' </summary>
    ''' <param name="taskName">The name of the task</param>
    ''' <remarks></remarks>
    Private Sub OnTaskStarted(taskName As String) Handles Me.TaskStarted
        SettingPreparationPage_ProgressLabel.Text = String.Format(GetValueFromLanguageData("MainForm.PreparationPanel_CurrentTaskField"), taskName)
        Refresh()
    End Sub

    ''' <summary>
    ''' Handles the TaskReported event
    ''' </summary>
    ''' <param name="Task">The reported item featuring the task name and whether the PT succeeded</param>
    ''' <remarks></remarks>
    Private Sub OnTaskReported(Task As Dictionary(Of String, Boolean)) Handles Me.TaskReported
        Dim taskName As String = Task.Keys(0)
        SettingPreparationPanel_TaskLv.Items.Add(New ListViewItem(New String() {taskName, If(Task(taskName), GetValueFromLanguageData("Common.Common_Yes"), GetValueFromLanguageData("Common.Common_No"))}))
        Refresh()
    End Sub

    ''' <summary>
    ''' A public method for the PT Helper to report task starts
    ''' </summary>
    ''' <param name="taskName">The name of the task</param>
    ''' <remarks></remarks>
    Sub ReportTaskStart(taskName As String)
        RaiseEvent TaskStarted(taskName)
    End Sub

    ''' <summary>
    ''' A public method for the PT Helper to report task status
    ''' </summary>
    ''' <param name="task">The reported item featuring the task name and whether the PT succeeded</param>
    ''' <remarks></remarks>
    Sub ReportTaskSuccess(task As Dictionary(Of String, Boolean))
        RaiseEvent TaskReported(task)
    End Sub

    Dim OriginalWindowBounds As Rectangle           ' Window bounds before full-screen
    Dim OriginalWindowState As FormWindowState      ' Window state before full-screen

    Dim currentTheme As Theme

    ''' <summary>
    ''' Changes the wizard page
    ''' </summary>
    ''' <param name="NewPage">The new page to change to</param>
    ''' <remarks>Verification in a page may be done</remarks>
    Sub ChangePage(NewPage As WizardPage.Page)
        DynaLog.LogMessage("Changing current page of the wizard...")
        DynaLog.LogMessage("New page to load: " & NewPage.ToString())
        If NewPage > CurrentWizardPage.WizardPage AndAlso VerifyInPages.Contains(CurrentWizardPage.WizardPage) Then
            If Not VerifyOptionsInPage(CurrentWizardPage.WizardPage) Then Exit Sub
        End If

        WelcomePage.Visible = (NewPage = WizardPage.Page.WelcomePage)
        SystemCheckPanel.Visible = (NewPage = WizardPage.Page.SysCheckPage)
        AdvSettingsPanel.Visible = (NewPage = WizardPage.Page.AdvSettingsPage)
        SettingPreparationPanel.Visible = (NewPage = WizardPage.Page.SettingPreparationPage)
        FinishPanel.Visible = (NewPage = WizardPage.Page.FinishPage)

        CurrentWizardPage.WizardPage = NewPage

        Select Case NewPage
            Case WizardPage.Page.SysCheckPage
                CheckComputer()
        End Select

        Next_Button.Enabled = (Not NewPage <> WizardPage.Page.FinishPage) OrElse (Not NewPage + 1 >= WizardPage.PageCount)
        Cancel_Button.Enabled = Not (NewPage = WizardPage.Page.FinishPage)
        Back_Button.Enabled = Not (NewPage = WizardPage.Page.WelcomePage) And Not (NewPage = WizardPage.Page.FinishPage)

        Next_Button.Text = If(NewPage = WizardPage.Page.FinishPage, GetValueFromLanguageData("Common.Common_Close"), GetValueFromLanguageData("Common.Common_Next"))

        ButtonPanel.Visible = Not (NewPage > WizardPage.Page.AdvSettingsPage)

        If NewPage = WizardPage.Page.SettingPreparationPage Then
            PrepareComputer()
        ElseIf NewPage = WizardPage.Page.FinishPage Then
            SysprepComputer()
        End If
    End Sub

    Sub ChangeTheme()
        ThemeHelper.LoadThemes()

        Dim ColorModeRk As RegistryKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize", False)
        Dim ColorModeVal As Integer = ColorModeRk.GetValue("AppsUseLightTheme", 1)
        ColorModeRk.Close()

        currentTheme = ThemeHelper.GetThemes().FirstOrDefault(Function(Theme) Theme.IsDark = (ColorModeVal = 0))

        If currentTheme IsNot Nothing Then
            BackColor = currentTheme.BackgroundColor
            ForeColor = currentTheme.ForegroundColor
            PageContainerPanel.BackColor = currentTheme.SectionBackgroundColor
            ButtonPanel.BackColor = currentTheme.BackgroundColor
            SysCheckPage_CheckDetailsGB.ForeColor = currentTheme.ForegroundColor
            SysCheckPage_ChecksLv.BackColor = currentTheme.SectionBackgroundColor
            SysCheckPage_ChecksLv.ForeColor = currentTheme.ForegroundColor
            SysCheckPage_CheckDetails_ResolutionValue.BackColor = currentTheme.SectionBackgroundColor
            SysCheckPage_CheckDetails_ResolutionValue.ForeColor = currentTheme.ForegroundColor
            AdvSettingsPage_CleanupActionCBox.BackColor = currentTheme.SectionBackgroundColor
            AdvSettingsPage_CleanupActionCBox.ForeColor = currentTheme.ForegroundColor
            AdvSettingsPage_ShutdownOptionsCBox.BackColor = currentTheme.SectionBackgroundColor
            AdvSettingsPage_ShutdownOptionsCBox.ForeColor = currentTheme.ForegroundColor
            AdvSettingsPage_SysprepUnatt_AnswerFileText.BackColor = currentTheme.SectionBackgroundColor
            AdvSettingsPage_SysprepUnatt_AnswerFileText.ForeColor = currentTheme.ForegroundColor
            SettingPreparationPanel_TaskLv.BackColor = currentTheme.SectionBackgroundColor
            SettingPreparationPanel_TaskLv.ForeColor = currentTheme.ForegroundColor
        End If
    End Sub

    Sub ChangeLanguage(LanguageCode As String)
        If Not File.Exists(Path.Combine(Application.StartupPath, "Languages", "lang_" & LanguageCode & ".ini")) Then
            LanguageCode = "en"
        End If
        LoadLanguageFile(Path.Combine(Application.StartupPath, "Languages", "lang_" & LanguageCode & ".ini"))

        AboutBtn.Text = GetValueFromLanguageData("Common.Common_About")
        Back_Button.Text = GetValueFromLanguageData("Common.Common_Back")
        Next_Button.Text = GetValueFromLanguageData("Common.Common_Next")
        Cancel_Button.Text = GetValueFromLanguageData("Common.Common_Cancel")
        WelcomePage_Header.Text = GetValueFromLanguageData("MainForm.WelcomePanel_Header")
        WelcomePage_Description.Text = GetValueFromLanguageData("MainForm.WelcomePanel_Description")
        SysCheckPage_Header.Text = GetValueFromLanguageData("MainForm.SystemChecksPanel_Header")
        SysCheckPage_Description.Text = GetValueFromLanguageData("MainForm.SystemChecksPanel_Description")
        SysCheckPage_CheckCH.Text = GetValueFromLanguageData("MainForm.SystemChecksPanel_CheckField")
        SysCheckPage_CompatibleCH.Text = GetValueFromLanguageData("MainForm.SystemChecksPanel_CompatibleField")
        SysCheckPage_SeverityCH.Text = GetValueFromLanguageData("MainForm.SystemChecksPanel_SeverityField")
        SysCheckPage_CheckDetailsGB.Text = GetValueFromLanguageData("MainForm.SystemChecksPanel_CheckDetailsField")
        SysCheckPage_CheckDetails_Description.Text = GetValueFromLanguageData("MainForm.SystemChecksPanel_CheckDescField")
        SysCheckPage_CheckDetails_Resolution.Text = GetValueFromLanguageData("MainForm.SystemChecksPanel_CheckResField")
        SysCheckPage_CheckAgainBtn.Text = GetValueFromLanguageData("MainForm.SystemChecksPanel_CheckAgainBtn")
        AdvSettingsPage_Header.Text = GetValueFromLanguageData("MainForm.AdvancedSettingsPanel_Header")
        AdvSettingsPage_Description.Text = GetValueFromLanguageData("MainForm.AdvancedSettingsPanel_Description")
        AdvSettingsPage_ConfigSysprepSettings.Text = GetValueFromLanguageData("MainForm.AdvancedSettingsPanel_ConfigCheck")
        AdvSettingsPage_CleanupActionLabel.Text = GetValueFromLanguageData("MainForm.AdvancedSettingsPanel_CleanupActionField")
        AdvSettingsPage_CleanupActionCBox.Items.AddRange(New String() {GetValueFromLanguageData("MainForm.AdvancedSettingsPanel_EnterOOBE_CA"), GetValueFromLanguageData("MainForm.AdvancedSettingsPanel_EnterAudit_CA")})
        AdvSettingsPage_ShutdownOptionLabel.Text = GetValueFromLanguageData("MainForm.AdvancedSettingsPanel_ShutdownOptionField")
        AdvSettingsPage_ShutdownOptionsCBox.Items.AddRange(New String() {GetValueFromLanguageData("MainForm.AdvancedSettingsPanel_RebootSO"), GetValueFromLanguageData("MainForm.AdvancedSettingsPanel_ShutdownSO"), GetValueFromLanguageData("MainForm.AdvancedSettingsPanel_QuitSO")})
        AdvSettingsPage_CleanupAction_Generalize.Text = GetValueFromLanguageData("MainForm.AdvancedSettingsPanel_GeneralizeField")
        AdvSettingsPage_SysprepUnattLabel.Text = GetValueFromLanguageData("MainForm.AdvancedSettingsPanel_AnswerFileField")
        AdvSettingsPage_SysprepUnatt_Btn.Text = GetValueFromLanguageData("Common.Common_Browse")
        AdvSettingsPage_VMMode.Text = GetValueFromLanguageData("MainForm.AdvancedSettingsPanel_VMModeField")
        AdvSettingsPage_VMModeExplanationLink.Text = GetValueFromLanguageData("MainForm.AdvancedSettingsPanel_VMModeLink")
        AdvSettingsPage_SysprepPrepToolDeploySteps.Text = GetValueFromLanguageData("MainForm.AdvancedSettingsPanel_ExplanationField")
        SettingPreparationPage_Header.Text = GetValueFromLanguageData("MainForm.PreparationPanel_Header")
        SettingPreparationPage_Description.Text = GetValueFromLanguageData("MainForm.PreparationPanel_Description")
        SettingPreparationPage_TaskCH.Text = GetValueFromLanguageData("MainForm.PreparationPanel_TaskField")
        SettingPreparationPage_SuccessfulCH.Text = GetValueFromLanguageData("MainForm.PreparationPanel_TaskSuccessField")
        FinishPage_Header.Text = GetValueFromLanguageData("MainForm.FinishPanel_Header")
        FinishPage_Description.Text = GetValueFromLanguageData("MainForm.FinishPanel_Description")
        FinishPage_CloseBtn.Text = GetValueFromLanguageData("MainForm.FinishPanel_CloseBtn")
        FinishPage_ResysprepBtn.Text = GetValueFromLanguageData("MainForm.FinishPanel_RelaunchBtn")
        FinishPage_RestartBtn.Text = GetValueFromLanguageData("MainForm.FinishPanel_RestartBtn")
    End Sub

    ''' <summary>
    ''' Verifies user-specified settings in the specified wizard page
    ''' </summary>
    ''' <param name="WizardPage">The wizard page to check on</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function VerifyOptionsInPage(WizardPage As WizardPage.Page) As Boolean
        DynaLog.LogMessage("Verifying user options before moving on to next page...")
        DynaLog.LogMessage("Page in which we need to verify user settings: " & WizardPage.ToString())
        Select Case WizardPage
            Case SysprepPreparator.WizardPage.Page.SysCheckPage
                If Not Environment.GetCommandLineArgs().Contains("/test") AndAlso PerformedChecks.Any(Function(PerformedCheck) PerformedCheck.Compatible = False) Then
                    MessageBox.Show(GetValueFromLanguageData("MainForm.IncompatibleCCPMessage"))
                    Return False
                End If
        End Select
        Return True
    End Function

    ''' <summary>
    ''' Starts the CCP Helper and reports the events logged by each of the CCPs
    ''' </summary>
    ''' <remarks></remarks>
    Sub CheckComputer()
        If IsComputerChecked Then
            Exit Sub
        End If
        SysCheckPage_ChecksLv.Items.Clear()
        Refresh()
        Cursor = Cursors.WaitCursor
        PerformedChecks = CompatibilityCheckerHelper.RunChecks()
        Cursor = Cursors.Arrow
        For Each PerformedCheck In PerformedChecks
            SysCheckPage_ChecksLv.Items.Add(New ListViewItem(New String() {PerformedCheck.StatusMessage.StatusTitle, If(PerformedCheck.Compatible, GetValueFromLanguageData("Common.Common_Yes"), GetValueFromLanguageData("Common.Common_No")), PerformedCheck.StatusMessage.SeverityToString(PerformedCheck.StatusMessage.StatusSeverity)}))
        Next
        IsComputerChecked = True
    End Sub

    ''' <summary>
    ''' Starts the PT Helper
    ''' </summary>
    ''' <remarks></remarks>
    Sub PrepareComputer()
        Refresh()
        Cursor = Cursors.WaitCursor
        PreparationTaskHelper.RunTasks()
        Cursor = Cursors.Arrow
        ButtonPanel.Visible = True
        Next_Button.PerformClick()
    End Sub

    ''' <summary>
    ''' Runs the final Sysprep process with the user-specified Sysprep settings
    ''' </summary>
    ''' <remarks></remarks>
    Sub SysprepComputer()
        Dim CmdLine As String = ParseSysprepSettings()
        If Environment.GetCommandLineArgs().Contains("/test") Then
            MsgBox("Sysprep would be launched with flags " & Quote & CmdLine & Quote, vbOKOnly + vbInformation)
        Else
            Dim sysprepProcess As New Process() With {
                .StartInfo = New ProcessStartInfo() With {
                    .FileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "system32", "sysprep", "sysprep.exe"),
                    .Arguments = CmdLine
                }
            }

            sysprepProcess.Start()
            sysprepProcess.WaitForExit()

            MsgBox("Sysprep exited with code " & sysprepProcess.ExitCode, vbOKOnly + vbInformation)
        End If
    End Sub

    ''' <summary>
    ''' Parses the current sysprep settings into a command line
    ''' </summary>
    ''' <returns>A set of command line arguments that are passed to sysprep</returns>
    ''' <remarks></remarks>
    Private Function ParseSysprepSettings() As String
        Dim sysprepCommandLine As String = ""
        If SysprepConfiguration.Generalize Then
            sysprepCommandLine &= "/generalize "
        End If
        Select Case SysprepConfiguration.Cleanup
            Case SysprepConfig.CleanupAction.EnterSystemAudit
                sysprepCommandLine &= "/audit "
            Case SysprepConfig.CleanupAction.EnterSystemOobe
                sysprepCommandLine &= "/oobe "
        End Select
        Select Case SysprepConfiguration.Shutdown
            Case SysprepConfig.ShutdownMode.Reboot
                sysprepCommandLine &= "/reboot "
            Case SysprepConfig.ShutdownMode.Shutdown
                sysprepCommandLine &= "/shutdown "
            Case SysprepConfig.ShutdownMode.Quit
                sysprepCommandLine &= "/quit "
        End Select
        If SysprepConfiguration.AnswerFile <> "" AndAlso File.Exists(SysprepConfiguration.AnswerFile) Then
            sysprepCommandLine &= String.Format("/unattend:{0} ", Quote & SysprepConfiguration.AnswerFile & Quote)
        End If
        If SysprepConfiguration.VMMode Then
            sysprepCommandLine &= "/mode:vm"
        End If
        Return sysprepCommandLine.TrimEnd(" ")
    End Function

    ''' <summary>
    ''' Initializes the default settings
    ''' </summary>
    ''' <remarks></remarks>
    Sub InitializeSettings()
        ' Initialize Sysprep Process Configuration Settings
        AdvSettingsPage_CleanupActionCBox.SelectedIndex = SysprepConfiguration.Cleanup
        AdvSettingsPage_CleanupAction_Generalize.Checked = SysprepConfiguration.Generalize
        AdvSettingsPage_ShutdownOptionsCBox.SelectedIndex = SysprepConfiguration.Shutdown
        AdvSettingsPage_SysprepUnatt_AnswerFileText.Text = SysprepConfiguration.AnswerFile
        AdvSettingsPage_VMMode.Checked = SysprepConfiguration.VMMode
    End Sub

    ''' <summary>
    ''' Handles changes in the sysprep configuration
    ''' </summary>
    ''' <remarks></remarks>
    Sub ChangeSysprepConfiguration()
        SysprepConfiguration.Cleanup = AdvSettingsPage_CleanupActionCBox.SelectedIndex
        SysprepConfiguration.Generalize = AdvSettingsPage_CleanupAction_Generalize.Checked
        SysprepConfiguration.Shutdown = AdvSettingsPage_ShutdownOptionsCBox.SelectedIndex
        SysprepConfiguration.AnswerFile = AdvSettingsPage_SysprepUnatt_AnswerFileText.Text
        SysprepConfiguration.VMMode = AdvSettingsPage_VMMode.Checked
    End Sub

    Private Sub Back_Button_Click(sender As Object, e As EventArgs) Handles Back_Button.Click
        ChangePage(CurrentWizardPage.WizardPage - 1)
    End Sub

    Private Sub Next_Button_Click(sender As Object, e As EventArgs) Handles Next_Button.Click
        If CurrentWizardPage.WizardPage = WizardPage.Page.FinishPage Then
            Close()
        Else
            ChangePage(CurrentWizardPage.WizardPage + 1)
        End If
    End Sub

    Private Sub Cancel_Button_Click(sender As Object, e As EventArgs) Handles Cancel_Button.Click
        Close()
    End Sub

    Private Sub SysCheckPage_CheckAgainBtn_Click(sender As Object, e As EventArgs) Handles SysCheckPage_CheckAgainBtn.Click
        IsComputerChecked = False
        CheckComputer()
    End Sub

    Function GetCopyrightTimespan(ByVal start As Integer, ByVal current As Integer) As String
        If current <= start Then
            Return current.ToString()
        Else
            Return start.ToString() & "-" & current.ToString()
        End If
    End Function

    Sub InitDynaLog()
        DynaLog.LogMessage("Sysprep Preparation Tool (Sysprep Preparator) - Version " & My.Application.Info.Version.ToString() & ", build timestamp: " & RetrieveLinkerTimestamp().ToString("yyMMdd-HHmm"))
        ' Display copyright/author information for every component
        DynaLog.LogMessage("Components:")
        DynaLog.LogMessage("- Program: " & My.Application.Info.Copyright.Replace("©", "(c)"))
        DynaLog.LogMessage("  Initial expansion and testing also made by Real-MullaC: https://github.com/Real-MullaC")
        DynaLog.LogMessage("- ManagedDism: (c) " & GetCopyrightTimespan(2016, 2016) & " Jeff Kluge")
        DynaLog.BeginLogging()
    End Sub

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitDynaLog()
        ChangeLanguage(My.Computer.Info.InstalledUICulture.TwoLetterISOLanguageName)
        ' Because of the DISM API, Windows 7 compatibility is out the window (no pun intended)
        If Environment.OSVersion.Version.Major = 6 And Environment.OSVersion.Version.Minor < 2 Then
            DynaLog.LogMessage("Windows 7 or an earlier version has been detected on this system. Program incompatible -- aborting any future procedures!")
            MsgBox(GetValueFromLanguageData("MainForm.Win7Message"), vbOKOnly + vbCritical, GetValueFromLanguageData("MainForm.WndTitle"))
            Environment.Exit(1)
        End If
        ' I once tested this on a computer which didn't require me to ask for admin privileges. This is a requirement of DISM. Check this
        If Not My.User.IsInRole(ApplicationServices.BuiltInRole.Administrator) Then
            DynaLog.LogMessage("This user is not part of the Administrators group/role -- aborting any future procedures!")
            MsgBox(GetValueFromLanguageData("MainForm.RunAsAdminMessage"), vbOKOnly + vbCritical, GetValueFromLanguageData("MainForm.WndTitle"))
            Environment.Exit(1)
        End If
        OriginalWindowState = WindowState
        OriginalWindowBounds = Bounds
        FormBorderStyle = FormBorderStyle.None
        WindowState = FormWindowState.Maximized
        ChangePage(WizardPage.Page.WelcomePage)
        InitializeSettings()
        AddHandler AdvSettingsPage_CleanupActionCBox.SelectedIndexChanged, AddressOf ChangeSysprepConfiguration
        AddHandler AdvSettingsPage_CleanupAction_Generalize.CheckedChanged, AddressOf ChangeSysprepConfiguration
        AddHandler AdvSettingsPage_ShutdownOptionsCBox.SelectedIndexChanged, AddressOf ChangeSysprepConfiguration
        AddHandler AdvSettingsPage_SysprepUnatt_AnswerFileText.TextChanged, AddressOf ChangeSysprepConfiguration
        AddHandler AdvSettingsPage_VMMode.CheckedChanged, AddressOf ChangeSysprepConfiguration
        ChangeTheme()
        FinishPage_CloseBtn.Enabled = Environment.GetCommandLineArgs().Contains("/test")
    End Sub

    Private Sub SysCheckPage_ChecksLv_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SysCheckPage_ChecksLv.SelectedIndexChanged
        SysCheckPage_CheckDetailsTLP.Visible = (SysCheckPage_ChecksLv.SelectedItems.Count = 1)
        If SysCheckPage_ChecksLv.SelectedItems.Count = 1 Then
            Dim focusedIdx As Integer = SysCheckPage_ChecksLv.FocusedItem.Index

            SysCheckPage_CheckDetails_Title.Text = PerformedChecks(focusedIdx).StatusMessage.StatusTitle
            SysCheckPage_CheckDetails_DescriptionValue.Text = PerformedChecks(focusedIdx).StatusMessage.StatusDescription
            SysCheckPage_CheckDetails_ResolutionValue.Text = PerformedChecks(focusedIdx).StatusMessage.StatusPossibleResolution
        End If
    End Sub

    Private Sub AdvSettingsPage_ConfigSysprepSettings_CheckedChanged(sender As Object, e As EventArgs) Handles AdvSettingsPage_ConfigSysprepSettings.CheckedChanged
        AdvSettingsPage_SysprepConfigPanel.Enabled = AdvSettingsPage_ConfigSysprepSettings.Checked
    End Sub

    Private Sub AdvSettingsPage_VMModeExplanationLink_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles AdvSettingsPage_VMModeExplanationLink.LinkClicked
        Process.Start("https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/sysprep-command-line-options?view=windows-11#modevm")
    End Sub

    Private Sub AdvSettingsPage_SysprepUnatt_Btn_Click(sender As Object, e As EventArgs) Handles AdvSettingsPage_SysprepUnatt_Btn.Click
        AdvSettingsPage_SysprepUnatt_OFD.ShowDialog()
    End Sub

    Private Sub AdvSettingsPage_SysprepUnatt_OFD_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles AdvSettingsPage_SysprepUnatt_OFD.FileOk
        AdvSettingsPage_SysprepUnatt_AnswerFileText.Text = AdvSettingsPage_SysprepUnatt_OFD.FileName
    End Sub

    Private Sub FinishPage_CloseBtn_Click(sender As Object, e As EventArgs) Handles FinishPage_CloseBtn.Click
        Close()
    End Sub

    Private Sub FinishPage_ResysprepBtn_Click(sender As Object, e As EventArgs) Handles FinishPage_ResysprepBtn.Click
        SysprepComputer()
    End Sub

    Private Sub FinishPage_RestartBtn_Click(sender As Object, e As EventArgs) Handles FinishPage_RestartBtn.Click
        Process.Start("shutdown", "-r -t 00")
    End Sub

    Private Sub MainForm_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.F11 Then
            ToggleFullScreenMode()
        End If
    End Sub

    Sub ToggleFullScreenMode()
        If FormBorderStyle = Windows.Forms.FormBorderStyle.None Then
            DynaLog.LogMessage("Exiting full-screen mode...")
            FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle
            Bounds = OriginalWindowBounds
            WindowState = OriginalWindowState
        Else
            DynaLog.LogMessage("Entering full-screen mode...")
            FormBorderStyle = Windows.Forms.FormBorderStyle.None
            OriginalWindowState = WindowState
            WindowState = FormWindowState.Normal
            OriginalWindowBounds = Bounds
            Bounds = Screen.FromControl(Me).Bounds
        End If
    End Sub

    Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        DynaLog.LogMessage("We Are Done")
        DynaLog.EndLogging()
    End Sub

    Private Sub AboutBtn_Click(sender As Object, e As EventArgs) Handles AboutBtn.Click
        MsgBox(String.Format(GetValueFromLanguageData("MainForm.AboutMessage"),
                             My.Application.Info.Version.ToString(),
                             RetrieveLinkerTimestamp().ToString("yyMMdd-HHmm"),
                             My.Application.Info.Copyright.Replace("©", "(c)"),
                             GetCopyrightTimespan(2016, 2016) & " Jeff Kluge"), vbOKOnly + vbInformation, GetValueFromLanguageData("Common.Common_About"))
    End Sub
End Class
