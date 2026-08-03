Imports BDELib.BDELib
Imports BDELib.Classes

Namespace Helpers.CompatChecks

    Public Class BitLockerCCP
        Inherits CompatibilityCheckerProvider

        Public Overrides Function PerformCompatibilityCheck() As Classes.CompatibilityCheckerProviderStatus
            Dim SystemPersistentVolumeId As String = GetPersistentVolumeIdFromSystemVolume()
            ' If we have no persistent volume ID, then we know that the volume is not encrypted
            If SystemPersistentVolumeId = "" Then
                Status.Compatible = True
                Status.StatusMessage = New Classes.StatusMessage(GetValueFromLanguageData("BitLockerCCP.CCPTitle"),
                                                                 GetValueFromLanguageData("BitLockerCCP.CCP_OK"),
                                                                 Classes.StatusMessage.StatusMessageSeverity.Info)
                Return Status
            End If

            Dim systemVolumeConversionStatus As ConversionStatus = GetVolumeConversionStatus(SystemPersistentVolumeId)
            If systemVolumeConversionStatus Is Nothing Then
                Status.Compatible = False
                Status.StatusMessage = New Classes.StatusMessage(GetValueFromLanguageData("BitLockerCCP.CCPTitle"),
                                                                 GetValueFromLanguageData("BitLockerCCP.CCP_NotOK_NoBDEInfo"),
                                                                 Classes.StatusMessage.StatusMessageSeverity.Critical)
                Return Status
            End If

            Status.Compatible = False

            Dim blStatusStr As String = GetValueFromLanguageData("BitLockerCCP.CCP_VolStat")
            Select Case systemVolumeConversionStatus.ConversionStatus
                Case VolumeConversionStatus.FullyDecrypted : blStatusStr = String.Format(blStatusStr, GetValueFromLanguageData("BitLockerCCP.CCP_VolStat_Decrypted"))       ' uhhh, what? this is not possible!
                Case VolumeConversionStatus.FullyEncrypted : blStatusStr = String.Format(blStatusStr, GetValueFromLanguageData("BitLockerCCP.CCP_VolStat_Encrypted"))
                Case VolumeConversionStatus.DecryptionInProgress : blStatusStr = String.Format(blStatusStr, GetValueFromLanguageData("BitLockerCCP.CCP_VolStat_DecryptionInProgress"))
                Case VolumeConversionStatus.EncryptionInProgress : blStatusStr = String.Format(blStatusStr, GetValueFromLanguageData("BitLockerCCP.CCP_VolStat_EncryptionInProgress"))
                Case VolumeConversionStatus.DecryptionPaused : blStatusStr = String.Format(blStatusStr, GetValueFromLanguageData("BitLockerCCP.CCP_VolStat_DecryptionPaused"))
                Case VolumeConversionStatus.EncryptionPaused : blStatusStr = String.Format(blStatusStr, GetValueFromLanguageData("BitLockerCCP.CCP_VolStat_EncryptionPaused"))
            End Select

            Dim divider As Integer = 10000

            blStatusStr &= String.Format(GetValueFromLanguageData("BitLockerCCP.CCP_PercentEncrypted"), Math.Round(systemVolumeConversionStatus.EncryptionPercentage / divider, 2))

            Status.StatusMessage = New Classes.StatusMessage(GetValueFromLanguageData("BitLockerCCP.CCPTitle"),
                                                             GetValueFromLanguageData("BitLockerCCP.CCP_NotOK_BDEInfo"),
                                                             String.Format(GetValueFromLanguageData("BitLockerCCP.CCP_NotOK_Resolution"), blStatusStr),
                                                             Classes.StatusMessage.StatusMessageSeverity.Critical)
            Return Status
        End Function

        Private Function GetPersistentVolumeIdFromSystemVolume() As String
            Dim PersistentVolumeId As String = ""

            Dim EncryptedVolumeMOC As ManagementObjectCollection = WMIHelper.GetResultsFromManagementQuery(String.Format("SELECT PersistentVolumeId FROM Win32_EncryptableVolume WHERE DriveLetter = {0}{1}{0}", Quote, WMIHelper.GetEscapedValue(Environment.GetEnvironmentVariable("SYSTEMDRIVE"))), "root\cimv2\Security\MicrosoftVolumeEncryption")
            If EncryptedVolumeMOC Is Nothing Then Return PersistentVolumeId

            PersistentVolumeId = WMIHelper.GetObjectValue(EncryptedVolumeMOC(0), "PersistentVolumeID")

            Return PersistentVolumeId
        End Function

    End Class

End Namespace
