Imports System.DirectoryServices.ActiveDirectory
Imports System.Management

Namespace Helpers.CompatChecks

    Public Class ActiveDirectoryDomainJoinCCP
        Inherits CompatibilityCheckerProvider

        ''' <summary>
        ''' Role of a computer in an assigned domain workgroup
        ''' </summary>
        Enum DomainRole As Integer
            ''' <summary>
            ''' Unknown domain role definition
            ''' </summary>
            Unknown = -1
            ''' <summary>
            ''' Standalone Workstation
            ''' </summary>
            StandaloneWorkstation = 0
            ''' <summary>
            ''' Member Workstation
            ''' </summary>
            MemberWorkstation = 1
            ''' <summary>
            ''' Standalone Server
            ''' </summary>
            StandaloneServer = 2
            ''' <summary>
            ''' Member Server
            ''' </summary>
            MemberServer = 3
            ''' <summary>
            ''' Backup Domain Controller
            ''' </summary>
            BackupDomainController = 4
            ''' <summary>
            ''' Primary Domain Controller
            ''' </summary>
            PrimaryDomainController = 5
        End Enum

        ''' <summary>
        ''' Gets the current domain role of the system.
        ''' </summary>
        ''' <returns>The current domain role of the system</returns>
        Private Function GetSystemDomainRole() As DomainRole
            Dim domainRoleCollection As ManagementObjectCollection = GetResultsFromManagementQuery("SELECT DomainRole FROM Win32_ComputerSystem")
            If domainRoleCollection IsNot Nothing Then
                Return GetObjectValue(domainRoleCollection(0), "DomainRole")
            End If
            Return DomainRole.Unknown
        End Function

        Public Overrides Function PerformCompatibilityCheck() As Classes.CompatibilityCheckerProviderStatus
            DynaLog.LogMessage("Detecting if device is part of Active Directory domain...")
            Try
                ' First we'll check if the device is a promoted domain controller. Both PDCs and BDCs (primaries and backups) are NOT
                ' supported by sysprep. Continuing without checking this will still return that it is part of a domain (in theory, it IS
                ' part of a domain regardless of role), but at least we don't call any ADDS functions to do this. Source:
                ' https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/get-started/virtual-dc/virtualized-domain-controllers-hyper-v#deployment-considerations
                If GetSystemDomainRole() >= DomainRole.BackupDomainController Then
                    DynaLog.LogMessage("This device is either a PDC or a BDC.")
                    Status.Compatible = True
                    Status.StatusMessage = New Classes.StatusMessage(GetValueFromLanguageData("ActiveDirectoryDomainJoinCCP.CCPTitle"),
                                                                     GetValueFromLanguageData("ActiveDirectoryDomainJoinCCP.CCP_NotOK_ADDSDC"),
                                                                     GetValueFromLanguageData("ActiveDirectoryDomainJoinCCP.CCP_NotOK_Resolution_ADDSDC"),
                                                                     Classes.StatusMessage.StatusMessageSeverity.Warning)
                Else
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
                End If
            Catch addsEx As ActiveDirectoryObjectNotFoundException
                DynaLog.LogMessage("A connection could not be established with the Domain Controller due to an error. Message: " & addsEx.Message)
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
