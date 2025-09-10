Imports SysprepPreparator.Classes
Imports SysprepPreparator.Helpers.CompatChecks

Namespace Helpers

    ''' <summary>
    ''' The Compatibility Checker Provider Helper module: it will launch all the registered CCPs
    ''' </summary>
    ''' <remarks></remarks>
    Public Module CompatibilityCheckerHelper

        ''' <summary>
        ''' The list of registered CCPs
        ''' </summary>
        ''' <remarks>To develop and register a CCP, read the documentation</remarks>
        Private CompatibilityCheckerModules As New List(Of CompatibilityCheckerProvider) From {
            New SetupStateCCP(),
            New PendingServicingOperationsCCP(),
            New DismThirdPartyDriverCCP(),
            New ActiveDirectoryDomainJoinCCP(),
            New ThirdPartyAppxCCP(),
            New WindowsServerRoleCCP()
        }

        ''' <summary>
        ''' Performs the checks of all the registered Compatibility Checker Providers (CCPs)
        ''' </summary>
        ''' <returns>All the status events logged</returns>
        ''' <remarks></remarks>
        Public Function RunChecks() As List(Of CompatibilityCheckerProviderStatus)
            DynaLog.LogMessage("Preparing to run Compatibility Checker Providers (CCPs)...")
            Dim StatusList As New List(Of CompatibilityCheckerProviderStatus)

            For Each CompatibilityCheckerModule In CompatibilityCheckerModules
                DynaLog.LogMessage("CCP to run: " & CompatibilityCheckerModule.GetType().Name)
                Dim result As CompatibilityCheckerProviderStatus = CompatibilityCheckerModule.PerformCompatibilityCheck()
                DynaLog.LogMessage(String.Format("[{0}] {1}", CompatibilityCheckerModule.GetType().Name, result.ToString()))
                StatusList.Add(result)
            Next

            Return StatusList
        End Function

    End Module

End Namespace
