Imports Microsoft.Win32

Namespace Helpers.CompatChecks

    Public Class ThirdPartyAppxCCP
        Inherits CompatibilityCheckerProvider

        Public Overrides Function PerformCompatibilityCheck() As Classes.CompatibilityCheckerProviderStatus
            Try
                Dim thirdPartyAppxsRk As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\AppModel\StagingInfo", False)
                ' AppX is just a steaming and messy pile of shit. From messy permissions to the fact that Windows breaks when they crash,
                ' to the fact that they can return File Explorer to Windows 10 by disabling a completely irrelevant feature (looking at you Recall),
                ' I just hate the entirety of the infrastructure. Yet people believe this is a REALLY good thing
                Dim Appxs() As String = thirdPartyAppxsRk.GetSubKeyNames()
                thirdPartyAppxsRk.Close()
                If Appxs.Count > 0 Then
                    Dim appxStr As String = ControlChars.CrLf & "- " & String.Join(ControlChars.CrLf & "- ", Appxs) & ControlChars.CrLf
                    Status.Compatible = True
                    Status.StatusMessage = New Classes.StatusMessage("Third-party Microsoft Store app checks",
                                                                     "Some packages that may only be registered to this user account have been detected. Sysprep may fail to run.",
                                                                     "These were the AppX packages detected: " & appxStr & "Verify the list and remove the ones you have installed. You can relaunch Sysprep as many times as you want after preparation is complete.",
                                                                     Classes.StatusMessage.StatusMessageSeverity.Warning)
                Else
                    Status.Compatible = True
                    Status.StatusMessage = New Classes.StatusMessage("Third-party Microsoft Store app checks",
                                                                     "No third-party AppX packages only registered to this user account were detected on your system.",
                                                                     Classes.StatusMessage.StatusMessageSeverity.Info)

                End If
            Catch ex As Exception
                DynaLog.LogMessage("An error occurred. Message: " & ex.Message)
                Status.Compatible = True
                Status.StatusMessage = New Classes.StatusMessage("Third-party Microsoft Store app checks",
                                                                 "No third-party AppX packages only registered to this user account could be detected on your system due to the following error: " & ex.Message,
                                                                 "Sysprep may fail. Proceed with care.",
                                                                 Classes.StatusMessage.StatusMessageSeverity.Warning)
            End Try

            Return Status
        End Function
    End Class

End Namespace
