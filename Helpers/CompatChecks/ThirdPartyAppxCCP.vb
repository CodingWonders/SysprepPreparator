Imports System.IO
Imports System.Management
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
                Dim Appxs As List(Of String) = thirdPartyAppxsRk.GetSubKeyNames().ToList()
                thirdPartyAppxsRk.Close()

                ' HACK: on (IoT) Enterprise LTSC versions of Windows, or basically any Windows edition without the Store
                ' by default, the Store appx package will not appear in StagingInfo and will throw a sysprep error. For this we need to look
                ' at whether or not we have a pckgdep file with the current user's corresponding SID. CURRENTLY WE ONLY ACCOUNT
                ' FOR THE LTSC EDITIONS OF WINDOWS FOR THIS ISSUE TO HAPPEN, ON NON-LTSC WE DON'T HAVE THIS ISSUE!!!
                ' ---
                ' Note that the Store needs to be installed with wsreset -i to do this.
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
    End Class

End Namespace
