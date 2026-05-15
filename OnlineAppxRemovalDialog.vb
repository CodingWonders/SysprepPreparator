Imports System.IO
Imports System.Management
Imports System.Windows.Forms
Imports Microsoft.Win32

Public Class OnlineAppxRemovalDialog

    Private Appxs As New List(Of String)

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If ListView1.CheckedItems.Count = 0 Then
            MessageBox.Show("Please select AppX packages to remove.", Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Cursor = Cursors.WaitCursor

        Dim MarkedAppxPackages As New List(Of String)

        MarkedAppxPackages.AddRange(ListView1.CheckedItems.Cast(Of ListViewItem)().Select(Function(lvi) lvi.Text))
        InvokeAppxRemoval(String.Join(";", MarkedAppxPackages.ToArray()))

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub InvokeAppxRemoval(MarkedAppxPackagesFilter As String)
        Dim extAppxHelperPath As String = Path.Combine(Application.StartupPath, "Tools", "RemoveOnlineAppxPackage.ps1")

        If File.Exists(extAppxHelperPath) Then
            Dim psAppxRemovalProc As New Process() With {
                .StartInfo = New ProcessStartInfo() With {
                    .FileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "system32", "WindowsPowerShell", "v1.0", "powershell.exe"),
                    .Arguments = String.Format("-executionpolicy Bypass -noprofile -nologo -file {0}{1}{0} -appxFullNames {0}{2}{0}", ControlChars.Quote, extAppxHelperPath, MarkedAppxPackagesFilter),
                    .CreateNoWindow = True,
                    .WindowStyle = ProcessWindowStyle.Hidden
                }
            }

            psAppxRemovalProc.Start()
            psAppxRemovalProc.WaitForExit()
            If psAppxRemovalProc.ExitCode = 0 Then
                Cursor = Cursors.Arrow
                ' Invoke the restart operation in 1 minute
                DynaLog.LogMessage("Restarting the computer in 1 minute!")
                Process.Start(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "system32", "shutdown.exe"),
                              String.Format("/r /t 60 /c {0}The Sysprep Preparation Tool has scheduled a system restart in 1 minute. " &
                              "You can restart your computer now by clicking OK on the message that will appear after this one. Please save your work.{0}", ControlChars.Quote)).WaitForExit()
                If MessageBox.Show("Click OK to restart your computer now.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information) = DialogResult.OK Then
                    DynaLog.LogMessage("Restarting the computer NOW!!!")
                    Process.Start(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "system32", "shutdown.exe"), "/a")
                    Process.Start(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "system32", "shutdown.exe"), "/r /t 0").WaitForExit()
                End If
            End If
        End If
    End Sub

    Private Sub GetAppxPackages()
        Try
            Dim thirdPartyAppxsRk As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\AppModel\StagingInfo", False)
            Appxs = thirdPartyAppxsRk.GetSubKeyNames().ToList()
            thirdPartyAppxsRk.Close()

            Dim imgEditionRk As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion", False)
            Dim imgEditionId As String = imgEditionRk.GetValue("EditionID", "")
            imgEditionRk.Close()
            If {"EnterpriseS", "IoTEnterpriseS"}.Contains(imgEditionId) Then
                DynaLog.LogMessage("Windows (IoT) Enterprise LTSC detected; checking presence of Store appx...")
                Try
                    Dim userSidMOC As ManagementObjectCollection = New ManagementObjectSearcher(String.Format("SELECT SID FROM Win32_UserAccount WHERE Name = {0}{1}{0}", ControlChars.Quote, Environment.GetEnvironmentVariable("USERNAME"))).Get()

                    Dim userSid As String = WMIHelper.GetObjectValue(userSidMOC(0), "SID")
                    Dim pkgRepoRootPath As String = GetAppxPackageRepositoryRootPath()
                    Dim StoreAppDirectories As String() = Directory.EnumerateDirectories(Path.Combine(pkgRepoRootPath, "Packages"), "Microsoft.WindowsStore*").ToArray()

                    DynaLog.LogMessage("Verifying Store AppX packages for user with SID " & userSid)

                    For Each StoreAppDirectory In StoreAppDirectories
                        Dim pckgdepPath As String = Path.Combine(StoreAppDirectory, String.Format("{0}.pckgdep", userSid))
                        If File.Exists(pckgdepPath) Then
                            DynaLog.LogMessage("This appx package has a pckgdep for this user! Adding...")
                            Appxs.Add(Path.GetDirectoryName(pckgdepPath).Split("\").Last())
                        End If
                    Next
                Catch ex As Exception
                    DynaLog.LogMessage(ex.Message)
                End Try
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Function GetAppxPackageRepositoryRootPath() As String
        Dim PackageRepositoryRootPath As String = ""

        Try
            Dim AppxRootKey As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Appx", False)
            PackageRepositoryRootPath = AppxRootKey.GetValue("PackageRepositoryRoot", "")
            AppxRootKey.Close()
        Catch ex As Exception
            DynaLog.LogMessage("Could not check root path. Error message: " & ex.Message)
        End Try

        Return PackageRepositoryRootPath
    End Function

    Private Sub OnlineAppxRemovalDialog_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GetAppxPackages()

        Text = GetValueFromLanguageData("OnlineAppxRemovalDialog.WndTitle")
        Label1.Text = GetValueFromLanguageData("OnlineAppxRemovalDialog.AppxRemoval_Title")
        Label2.Text = GetValueFromLanguageData("OnlineAppxRemovalDialog.AppxRemoval_Notes")
        OK_Button.Text = GetValueFromLanguageData("Common.Common_OK")
        Cancel_Button.Text = GetValueFromLanguageData("Common.Common_Cancel")

        ListView1.Items.Clear()
        ListView1.Items.AddRange(Appxs.Select(Function(Appx) New ListViewItem(New String() {Appx})).ToArray())
    End Sub

    Private Sub ListView1_ItemChecked(sender As Object, e As ItemCheckedEventArgs) Handles ListView1.ItemChecked
        OK_Button.Enabled = ListView1.CheckedItems.Count > 0
    End Sub
End Class
