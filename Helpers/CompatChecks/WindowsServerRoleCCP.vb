Imports SysprepPreparator.Helpers.CompatChecks
Imports Microsoft.Dism
Imports System.IO
Imports SysprepPreparator.Classes
Imports Microsoft.Win32
Imports System.Runtime.CompilerServices
Imports System.Runtime.Serialization

Namespace Helpers.CompatChecks

    Public Class WindowsServerRoleCCP
        Inherits CompatibilityCheckerProvider

        ''' <summary>
        ''' Version constant for Windows Server 2008
        ''' </summary>
        Private ReadOnly OSVER_WINSERVER2008 As Version = New Version(6, 0, 6001)

        ''' <summary>
        ''' Version constant for Windows Server 2008 R2
        ''' </summary>
        Private ReadOnly OSVER_WINSERVER2008R2 As Version = New Version(6, 1, 7600)

        ''' <summary>
        ''' Version constant for Windows Server 2012
        ''' </summary>
        Private ReadOnly OSVER_WINSERVER2012 As Version = New Version(6, 2, 9200)

        ''' <summary>
        ''' Version constant for no releases
        ''' </summary>
        Private ReadOnly OSVER_NONE As Version = New Version(9999, 0, 0, 0)

        ' Mapping dictionary for Windows Server roles and corresponding features in the system:
        ' TODO: get information from Active Directory Federation Services role
        ' TODO: get information from Network Policy Routing and Remote Access Services role (requires installing Windows Server 2008)
        ' TODO: get information from Streaming Media Services
        ' TODO: get information from UDDI Services (requires installing Windows Server 2008)
        Private RoleFeatureMappingDictionary As New Dictionary(Of String, WindowsServerRole) From {
            {"AD-Certificate", New WindowsServerRole("ADCertificateServicesRole", "Active Directory Certificate Services (AD CS)", OSVER_NONE)},
            {"AD-Domain-Services", New WindowsServerRole("DirectoryServices-DomainController", "Active Directory Domain Services (AD DS)", OSVER_NONE)},
            {"ADLDS", New WindowsServerRole("DirectoryServices-ADAM", "Active Directory Lightweight Directory Services (AD LDS)", OSVER_NONE)},
            {"ADRMS", New WindowsServerRole("RightsManagementServices-Role", "Active Directory Rights Management Services (AD RMS)", OSVER_NONE)},
            {"Application-Server", New WindowsServerRole("Application-Server", "Application Server", OSVER_WINSERVER2008)},
            {"DHCP", New WindowsServerRole("DHCPServer", "Dynamic Host Configuration Protocol (DHCP) Server", OSVER_WINSERVER2008, OSVER_WINSERVER2008)},
            {"DNS", New WindowsServerRole("DNS-Server-Full-Role", "Domain Name System (DNS) Server", OSVER_NONE, "Not Applicable")},
            {"Fax", New WindowsServerRole("FaxServiceRole", "Fax Server", OSVER_NONE)},
            {"FileAndStorage-Services", New WindowsServerRole("FileAndStorage-Services", "File and Storage Services", OSVER_WINSERVER2008R2)},
            {"Hyper-V", New WindowsServerRole("Microsoft-Hyper-V", "Hyper-V", OSVER_WINSERVER2008R2, "Not supported for a virtual network on Hyper-V™. You must delete any virtual networks before you run the Sysprep tool.")},
            {"NPAS", New WindowsServerRole("NPAS-Role", "Network Policy and Access Services", OSVER_NONE)},
            {"Print-Services", New WindowsServerRole("Printing-Server-Foundation-Features", "Printing and Document Services", OSVER_WINSERVER2008R2)},
            {"Remote-Desktop-Services", New WindowsServerRole("Remote-Desktop-Services", "Remote Desktop Services", OSVER_WINSERVER2008)},
            {"VolumeActivation", New WindowsServerRole("VolumeActivation-Full-Role", "Volume Activation Services", OSVER_NONE)},
            {"Web-Server", New WindowsServerRole("IIS-WebServerRole", "Web Server (Internet Information Services)", OSVER_WINSERVER2008, "Not supported with encrypted credentials in the Applicationhost.config file.")},
            {"WDS", New WindowsServerRole("Microsoft-Windows-Deployment-Services", "Windows Deployment Services (WDS)", OSVER_WINSERVER2012, "Not supported if Windows Deployment Services is initialized. You need to uninitialize the server by running wdsutil /uninitialize-server.")},
            {"UpdateServices", New WindowsServerRole("UpdateServices", "Windows Server Update Services (WSUS)", OSVER_NONE)}
        }

        Private Function GetOSVersion() As Version
            DynaLog.LogMessage("Preparing to get OS version from ntoskrnl...")
            If File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "System32", "ntoskrnl.exe")) Then
                DynaLog.LogMessage("System ntoskrnl exists. Getting information...")
                Dim fileInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "System32", "ntoskrnl.exe"))
                Return New Version(fileInfo.FileMajorPart, fileInfo.FileMinorPart, fileInfo.FileBuildPart, fileInfo.FilePrivatePart)
            Else
                DynaLog.LogMessage("System ntoskrnl does not exist...")
                Return New Version(0, 0, 0, 0)
            End If
        End Function

        Public Overrides Function PerformCompatibilityCheck() As Classes.CompatibilityCheckerProviderStatus
            Dim roleMessage As String = ""
            Dim CompatibleRoles As Integer = 0, IncompatibleRoles As Integer = 0

            DynaLog.LogMessage("Preparing to get compatibility with installed Windows Server roles...")

            ' The way we'll check for roles is by listing the features with DISM API. We'll also grab
            ' the current version of our OS because compatibility also depends on the OS version.
            Try
                DynaLog.LogMessage("Getting current system edition ID...")
                Dim EditionRk As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion")
                Dim EditionId As String = EditionRk.GetValue("EditionID").ToString()
                DynaLog.LogMessage("Edition ID: " & EditionId)
                EditionRk.Close()
                If Not EditionId.Contains("Server") Then
                    DynaLog.LogMessage("This is not a server install, we don't need to do any of this...")
                    ' This is not a server edition, so we skip this CCP
                    Status = New Classes.CompatibilityCheckerProviderStatus(True,
                                                                            New StatusMessage("Windows Server Role Compatibility Check",
                                                                                              "This check was skipped because this system is not running Windows Server",
                                                                                              StatusMessage.StatusMessageSeverity.Info))
                    Return Status
                End If
                DynaLog.LogMessage("This is a server install. We continue with role mappings...")
                DynaLog.LogMessage("Getting current system version...")
                Dim osVersion As Version = GetOSVersion()
                DynaLog.LogMessage("ntoskrnl version: " & osVersion.ToString())
                If osVersion.Major < 6 Then
                    DynaLog.LogMessage("System ntoskrnl major version is less than 6. Nope...")
                    Throw New Exception(String.Format("Current OS version ({0}) is not compatible with this CCP.", osVersion.ToString()))
                End If

                Dim featureInfoCollection As DismFeatureCollection = Nothing

                DynaLog.LogMessage("Initializing API...")
                DismApi.Initialize(DismLogLevel.LogErrors)
                DynaLog.LogMessage("Creating session...")
                Using session As DismSession = DismApi.OpenOnlineSession()
                    DynaLog.LogMessage("Getting features...")
                    featureInfoCollection = DismApi.GetFeatures(session)
                End Using

                If featureInfoCollection IsNot Nothing Then
                    DynaLog.LogMessage("Feature collection is not nothing. Preparing to parse features")
                    For Each Role In RoleFeatureMappingDictionary.Keys
                        DynaLog.LogMessage("Mapping feature with role " & Role & "...")

                        Dim RoleInfo As WindowsServerRole = RoleFeatureMappingDictionary(Role)
                        Dim featureFromRole As DismFeature = featureInfoCollection.FirstOrDefault(Function(feature) feature.FeatureName.Equals(RoleInfo.FeatureName,
                                                                                                                                               StringComparison.InvariantCultureIgnoreCase))
                        If featureFromRole Is Nothing Then
                            DynaLog.LogMessage("A corresponding feature does not exist.")
                            ' Then we know the feature does not exist
                            Continue For
                        End If

                        DynaLog.LogMessage("Determining state of this feature...")
                        DynaLog.LogMessage("- Determined Feature: " & featureFromRole.FeatureName)

                        If featureFromRole.State = DismPackageFeatureState.Installed Then
                            DynaLog.LogMessage("The feature is enabled. Getting information...")
                            DynaLog.LogMessage("- Baseline Version: " & RoleInfo.BaselineVersion.ToString())
                            DynaLog.LogMessage("- Maximum Version : " & RoleInfo.MaximumVersion.ToString())
                            If osVersion < RoleInfo.BaselineVersion Then
                                DynaLog.LogMessage("OS Version is lower than baseline version. The role is not supported.")
                                Select Case RoleInfo.BaselineVersion
                                    Case OSVER_NONE
                                        roleMessage &= String.Format("- {0} is not supported by Sysprep in any version of Windows Server{1}",
                                                                     RoleInfo.DisplayName, Environment.NewLine)
                                    Case Else
                                        Dim serverVersionString As String = ""
                                        Select Case RoleInfo.BaselineVersion
                                            Case OSVER_WINSERVER2008
                                                serverVersionString = "2008"
                                            Case OSVER_WINSERVER2008R2
                                                serverVersionString = "2008 R2"
                                            Case OSVER_WINSERVER2012
                                                serverVersionString = "2012"
                                            Case Else
                                                serverVersionString = String.Format("v{0}", RoleInfo.BaselineVersion.ToString())
                                        End Select
                                        roleMessage &= String.Format("- {0} is not supported by Sysprep and requires at least Windows Server {1}{2}",
                                                                     RoleInfo.DisplayName, serverVersionString, Environment.NewLine)
                                End Select
                                IncompatibleRoles += 1
                            ElseIf osVersion > RoleInfo.MaximumVersion Then
                                DynaLog.LogMessage("OS Version is greater than maximum version. The role is not supported.")
                                Dim serverVersionString As String = ""
                                Select Case RoleInfo.MaximumVersion
                                    Case OSVER_WINSERVER2008
                                        serverVersionString = "2008"
                                    Case OSVER_WINSERVER2008R2
                                        serverVersionString = "2008 R2"
                                    Case OSVER_WINSERVER2012
                                        serverVersionString = "2012"
                                    Case Else
                                        serverVersionString = String.Format("v{0}", RoleInfo.MaximumVersion.ToString())
                                End Select
                                roleMessage &= String.Format("- {0} is not supported by Sysprep and is only supported up to Windows Server {1}{2}",
                                                             RoleInfo.DisplayName, serverVersionString, Environment.NewLine)
                                IncompatibleRoles += 1
                            Else
                                DynaLog.LogMessage("OS Version is between baseline and maximum versions. The role is supported.")
                                roleMessage &= String.Format("- {0} is installed and supported. Possible caveat: {1}{2}",
                                                             RoleInfo.DisplayName, RoleInfo.CompatibilityCaveat, Environment.NewLine)
                                CompatibleRoles += 1
                            End If
                        Else
                            DynaLog.LogMessage("The feature is not enabled.")
                            roleMessage &= String.Format("- {0} is not installed.{1}", RoleInfo.DisplayName, Environment.NewLine)
                            CompatibleRoles += 1
                        End If
                    Next
                End If

            Catch ex As Exception
                DynaLog.LogMessage("An error occurred. Error message: " & ex.Message)
                Status = New Classes.CompatibilityCheckerProviderStatus(True,
                                                                        New StatusMessage("Windows Server Role Compatibility Checks",
                                                                                          "We could not check role compatibility.",
                                                                                          "The following error occurred: " & ex.Message & ". You can continue, but proceed with care.",
                                                                                          StatusMessage.StatusMessageSeverity.Warning))
            Finally
                Try
                    DynaLog.LogMessage("Shutting down API...")
                    DismApi.Shutdown()
                Catch ex As Exception

                End Try
            End Try

            Status = New Classes.CompatibilityCheckerProviderStatus((CompatibleRoles >= IncompatibleRoles),
                                                                    New StatusMessage("Windows Server Role Compatibility Check",
                                                                                      If(IncompatibleRoles > 0, "This installation contains roles that may not be supported by Sysprep", "This installation does not contain roles that may not be supported by Sysprep"),
                                                                                      String.Format("The following information was obtained from roles: " & Environment.NewLine & "{0}" & Environment.NewLine & "Verify these roles and any caveats associated to them.", roleMessage),
                                                                                      If(CompatibleRoles >= IncompatibleRoles, StatusMessage.StatusMessageSeverity.Info, StatusMessage.StatusMessageSeverity.Warning)))

            Return Status
        End Function

    End Class

End Namespace
