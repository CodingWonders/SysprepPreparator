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
                    Status.StatusMessage = New Classes.StatusMessage(GetValueFromLanguageData("ThirdPartyAppxCCP.CCPTitle"),
                                                                     GetValueFromLanguageData("ThirdPartyAppxCCP.CCP_NotOK"),
                                                                     String.Format(GetValueFromLanguageData("ThirdPartyAppxCCP.CCP_NotOK_Resolution_Generic"), appxStr),
                                                                     Classes.StatusMessage.StatusMessageSeverity.Warning)
                Else
                    Status.Compatible = True
                    Status.StatusMessage = New Classes.StatusMessage(GetValueFromLanguageData("ThirdPartyAppxCCP.CCPTitle"),
                                                                     GetValueFromLanguageData("ThirdPartyAppxCCP.CCP_OK"),
                                                                     Classes.StatusMessage.StatusMessageSeverity.Info)

                End If
            Catch ex As Exception
                DynaLog.LogMessage("An error occurred. Message: " & ex.Message)
                Status.Compatible = True
                Status.StatusMessage = New Classes.StatusMessage(GetValueFromLanguageData("ThirdPartyAppxCCP.CCPTitle"),
                                                                 String.Format(GetValueFromLanguageData("ThirdPartyAppxCCP.CCP_Error"), ex.Message),
                                                                 GetValueFromLanguageData("ThirdPartyAppxCCP.CCP_Error_Resolution"),
                                                                 Classes.StatusMessage.StatusMessageSeverity.Warning)
            End Try

            Return Status
        End Function
    End Class

End Namespace
