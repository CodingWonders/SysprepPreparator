Imports System.DirectoryServices.ActiveDirectory

Namespace Helpers.CompatChecks

    Public Class ActiveDirectoryDomainJoinCCP
        Inherits CompatibilityCheckerProvider

        Public Overrides Function PerformCompatibilityCheck() As Classes.CompatibilityCheckerProviderStatus
            DynaLog.LogMessage("Detecting if device is part of Active Directory domain...")
            Try
                ' This, in a non-joined system, should throw an exception
                Dim currentDomain As Domain = Domain.GetComputerDomain()
                If currentDomain IsNot Nothing Then
                    DynaLog.LogMessage("Current domain info is not nothing. This device is well likely part of a domain.")
                    Status.Compatible = True
                    Status.StatusMessage = New Classes.StatusMessage(GetValueFromLanguageData("ActiveDirectoryDomainJoinCCP.CCPTitle"),
                                                                     GetValueFromLanguageData("ActiveDirectoryDomainJoinCCP.CCP_NotOK"),
                                                                     GetValueFromLanguageData("ActiveDirectoryDomainJoinCCP.CCP_NotOK_Resolution_Generic"),
                                                                     Classes.StatusMessage.StatusMessageSeverity.Info)
                End If
            Catch ex As Exception
                DynaLog.LogMessage("An error occurred. Message: " & ex.Message)
                DynaLog.LogMessage("This is expected to happen when the device is not part of a domain.")
                Status.Compatible = True
                Status.StatusMessage = New Classes.StatusMessage(GetValueFromLanguageData("ActiveDirectoryDomainJoinCCP.CCPTitle"),
                                                                 GetValueFromLanguageData("ActiveDirectoryDomainJoinCCP.CCP_OK"),
                                                                 Classes.StatusMessage.StatusMessageSeverity.Info)
            End Try
            Return Status
        End Function

    End Class

End Namespace
