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
                Status.StatusMessage = New Classes.StatusMessage("BitLocker Drive Encryption Status on System Volume",
                                                                 "The system volume is not encrypted with BitLocker.",
                                                                 Classes.StatusMessage.StatusMessageSeverity.Info)
                Return Status
            End If

            Dim systemVolumeConversionStatus As ConversionStatus = GetVolumeConversionStatus(SystemPersistentVolumeId)
            If systemVolumeConversionStatus Is Nothing Then
                Status.Compatible = False
                Status.StatusMessage = New Classes.StatusMessage("BitLocker Drive Encryption Status on System Volume",
                                                                 "The system volume is encrypted with BitLocker, but we could not get information about it. Sysprep will fail to validate your installation.",
                                                                 Classes.StatusMessage.StatusMessageSeverity.Critical)
                Return Status
            End If

            Status.Compatible = False

            Dim blStatusStr As String = ""
            Select Case systemVolumeConversionStatus.ConversionStatus
                Case VolumeConversionStatus.FullyDecrypted : blStatusStr &= String.Format("- Volume Status: Fully Decrypted{0}", Environment.NewLine)       ' uhhh, what? this is not possible!
                Case VolumeConversionStatus.FullyEncrypted : blStatusStr &= String.Format("- Volume Status: Fully Encrypted{0}", Environment.NewLine)
                Case VolumeConversionStatus.DecryptionInProgress : blStatusStr &= String.Format("- Volume Status: Decryption in progress{0}", Environment.NewLine)
                Case VolumeConversionStatus.EncryptionInProgress : blStatusStr &= String.Format("- Volume Status: Encryption in progress{0}", Environment.NewLine)
                Case VolumeConversionStatus.DecryptionPaused : blStatusStr &= String.Format("- Volume Status: Decryption paused{0}", Environment.NewLine)
                Case VolumeConversionStatus.EncryptionPaused : blStatusStr &= String.Format("- Volume Status: Encryption paused{0}", Environment.NewLine)
            End Select

            Dim divider As Integer = 10000

            blStatusStr &= String.Format("- % Encrypted: {0}%", Math.Round(systemVolumeConversionStatus.EncryptionPercentage / divider, 2))

            Status.StatusMessage = New Classes.StatusMessage("BitLocker Drive Encryption Status on System Volume",
                                                             "The system volume is encrypted with BitLocker. Sysprep will fail to validate your installation.",
                                                             String.Format("This is the current encryption state of the system volume:{0}{1}{0}You must decrypt this volume by using the Control Panel or manage-bde. You don't need to restart your computer. After it is fully done, refresh the checks.", Environment.NewLine, blStatusStr),
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
