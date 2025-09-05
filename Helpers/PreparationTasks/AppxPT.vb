Imports System.IO
Imports System.Diagnostics
Imports System.Windows.Forms
Imports Microsoft.Win32

Namespace Helpers.PreparationTasks

    ''' <summary>
    ''' Preparation Task to install and provision Appx/MSIX packages
    ''' Only runs if there are Appx packages detected in HKEY_LOCAL_MACHINE\...\StagingInfo
    ''' </summary>
    Public Class AppxProvisionPT
        Inherits PreparationTask

        Private Const DefaultDestFolder As String = "C:\AppxPackages"

        ''' <summary>
        ''' Runs the Preparation Task
        ''' </summary>
        ''' <returns>Whether the operation succeeded</returns>
        Public Overrides Function RunPreparationTask() As Boolean
            Try
                If IsInTestMode Then Return True
                Dim appxPackages = GetMachineLevelAppxPackages()
                If appxPackages.Count = 0 Then
                    DynaLog.LogMessage("No Appx/MSIX packages detected. Skipping provisioning task.")
                    Return True
                End If

                DynaLog.LogMessage("Detected Appx/MSIX packages: " & String.Join(", ", appxPackages))
                ShowMessage("Detected Appx/MSIX packages: " & String.Join(", ", appxPackages) &
                            ". You may choose to provision them to avoid Sysprep issues.", "Appx Provisioning Info")

                Dim ofd As New OpenFileDialog()
                ofd.Filter = "Appx or Msix Files|*.appx;*.msix;*.msixbundle"
                ofd.Title = "Select Appx or Msix Packages to Provision (Optional)"
                ofd.Multiselect = True

                If ofd.ShowDialog() <> DialogResult.OK Then
                    DynaLog.LogMessage("User did not select any packages. Skipping provisioning.")
                    Return True
                End If

                If Not Directory.Exists(DefaultDestFolder) Then Directory.CreateDirectory(DefaultDestFolder)

                ProvisionSelectedPackages(ofd.FileNames)

                Return True
            Catch ex As Exception
                DynaLog.LogMessage("AppxProvisionPT encountered an error: " & ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Reads machine-level Appx packages from StagingInfo
        ''' </summary>
        Private Function GetMachineLevelAppxPackages() As List(Of String)
            Dim list As New List(Of String)
            Try
                Using regKey As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\AppModel\StagingInfo", False)
                    If regKey IsNot Nothing Then
                        Dim Appxs() As String = regKey.GetSubKeyNames()
                        If Appxs.Length > 0 Then list.AddRange(Appxs)
                    End If
                End Using
            Catch ex As Exception
                DynaLog.LogMessage("Error reading Appx packages from registry: " & ex.Message)
            End Try
            Return list
        End Function

        ''' <summary>
        ''' Copies selected files to DefaultDestFolder, installs user-wide, and provisions system-wide
        ''' </summary>
        Private Sub ProvisionSelectedPackages(files() As String)
            For Each selectedFile As String In files
                Dim destPath As String = Path.Combine(DefaultDestFolder, Path.GetFileName(selectedFile))

                Try
                    File.Copy(selectedFile, destPath, True)
                    DynaLog.LogMessage("Copied " & selectedFile & " to " & destPath)
                Catch ex As Exception
                    DynaLog.LogMessage("Failed to copy " & selectedFile & ": " & ex.Message)
                    Continue For
                End Try

                Try
                    DynaLog.LogMessage("Provisioning " & Path.GetFileName(selectedFile) & " system-wide...")
                    RunPowerShellCommand($"Add-AppxProvisionedPackage -Online -PackagePath '{destPath}' -SkipLicense")
                    DynaLog.LogMessage("Provisioned " & Path.GetFileName(selectedFile) & " system-wide.")
                Catch ex As Exception
                    DynaLog.LogMessage("Failed to provision system-wide: " & ex.Message)
                    Continue For
                End Try
            Next
        End Sub

        ''' <summary>
        ''' Runs a PowerShell command
        ''' </summary>
        Private Sub RunPowerShellCommand(command As String)
            Dim psi As New ProcessStartInfo()
            psi.FileName = "powershell.exe"
            psi.Arguments = "-Command " & command
            psi.UseShellExecute = True
            Dim proc As Process = Process.Start(psi)
            proc.WaitForExit()
        End Sub

    End Class

End Namespace
