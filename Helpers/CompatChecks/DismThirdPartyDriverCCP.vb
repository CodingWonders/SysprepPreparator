Imports Microsoft.Dism
Imports System.IO

Namespace Helpers.CompatChecks

    ''' <summary>
    ''' The DISM Third-Party Driver Compatibility Checker Provider (CCP)
    ''' </summary>
    ''' <remarks></remarks>
    Public Class DismThirdPartyDriverCCP
        Inherits CompatibilityCheckerProvider

        ''' <summary>
        ''' Performs a compatibility check to determine whether third-party drivers are found in the installation
        ''' </summary>
        ''' <returns>An event detailing results</returns>
        ''' <remarks></remarks>
        Public Overrides Function PerformCompatibilityCheck() As Classes.CompatibilityCheckerProviderStatus
            DynaLog.LogMessage("Detecting if device contains third-party drivers...")
            Try
                DynaLog.LogMessage("Initializing API...")
                DismApi.Initialize(DismLogLevel.LogErrors)
                DynaLog.LogMessage("Creating session...")
                Using session As DismSession = DismApi.OpenOnlineSession()
                    ' We explicitly declare that we DON'T want to grab all drivers. This lets us get the third-party drivers
                    ' more easily. Getting all drivers will add all the in-box drivers to the collection. These device drivers
                    ' are extremely likely to be general drivers for all machines, not just those the sysadmin might want to target,
                    ' but all of the machines around the world that have Windows installed.
                    DynaLog.LogMessage("Getting drivers...")
                    Dim driverInfoCollection As DismDriverPackageCollection = DismApi.GetDrivers(session, False)
                    Status.Compatible = True

                    DynaLog.LogMessage("Detecting if there are any third-party drivers...")
                    If driverInfoCollection.Any(Function(driver) driver.InBox = False) Then
                        DynaLog.LogMessage("There are third party drivers. External devices have been detected.")
                        Dim Drivers() As String = driverInfoCollection.Where(Function(driver) driver.InBox = False).Select(Function(driver) String.Format("{0} ({1})", driver.PublishedName, Path.GetFileName(driver.OriginalFileName))).ToArray()
                        Dim drvStr As String = ControlChars.CrLf & "- " & String.Join(ControlChars.CrLf & "- ", Drivers) & ControlChars.CrLf
                        Status.StatusMessage = New Classes.StatusMessage(GetValueFromLanguageData("DismThirdPartyDriverCCP.CCPTitle"),
                                                                         GetValueFromLanguageData("DismThirdPartyDriverCCP.CCP_NotOK"),
                                                                         String.Format(GetValueFromLanguageData("DismThirdPartyDriverCCP.CCP_NotOK_Resolution_Generic"), drvStr),
                                                                         Classes.StatusMessage.StatusMessageSeverity.Warning)
                    Else
                        DynaLog.LogMessage("There aren't any third party drivers. External devices have not been detected.")
                        Status.StatusMessage = New Classes.StatusMessage(GetValueFromLanguageData("DismThirdPartyDriverCCP.CCPTitle"),
                                                                         GetValueFromLanguageData("DismThirdPartyDriverCCP.CCP_OK"),
                                                                         Classes.StatusMessage.StatusMessageSeverity.Info)
                    End If
                End Using
            Catch ex As Exception
                DynaLog.LogMessage("Could not get driver information. Error: " & ex.Message)
                Status.Compatible = False
                Status.StatusMessage = New Classes.StatusMessage(GetValueFromLanguageData("DismThirdPartyDriverCCP.CCPTitle"),
                                                                 GetValueFromLanguageData("DismThirdPartyDriverCCP.CCP_Error"),
                                                                 GetValueFromLanguageData("DismThirdPartyDriverCCP.CCP_Error_Resolution"),
                                                                 Classes.StatusMessage.StatusMessageSeverity.Critical)
            Finally
                DynaLog.LogMessage("Shutting down API...")
                Try
                    DismApi.Shutdown()
                Catch ex As Exception

                End Try
            End Try

            Return Status
        End Function
    End Class

End Namespace
