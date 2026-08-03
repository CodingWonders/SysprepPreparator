Imports System.Windows.Forms
Imports BDELib.BDELib
Imports BDELib.Classes
Imports System.ComponentModel

Public Class SysvolBDEDecryptDialog

    Private Async Sub SysvolBDEDecryptDialog_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        WindowHelper.DisableCloseCapability(Handle)

        ' Change the language
        Text = GetValueFromLanguageData("SysvolBDEDecryptDialog.WndTitle")
        Label1.Text = GetValueFromLanguageData("SysvolBDEDecryptDialog.DecryptionProgress_Header")
        GroupBox1.Text = GetValueFromLanguageData("SysvolBDEDecryptDialog.DecryptionDetails")
        Label2.Text = GetValueFromLanguageData("SysvolBDEDecryptDialog.VolumeDeviceId")
        Label3.Text = GetValueFromLanguageData("SysvolBDEDecryptDialog.VolumePersistentId")
        Label4.Text = GetValueFromLanguageData("SysvolBDEDecryptDialog.VolumeConversionStatus")
        Label5.Text = GetValueFromLanguageData("SysvolBDEDecryptDialog.VolumePercentEncrypted")

        Visible = True

        Dim systemDeviceId As String = GetDeviceIdFromSystemVolume(),
            systemPersistentVolumeId As String = GetPersistentVolumeIdFromSystemVolume()

        lblDeviceID.Text = systemDeviceId
        lblPersistentVolumeID.Text = systemPersistentVolumeId

        Dim DecryptionResult As UInteger = Await Task.Run(Function()
                                                              Return StartVolumeDecryption(systemPersistentVolumeId, Sub(ConvStatus As ConversionStatus)
                                                                                                                         If ConvStatus Is Nothing Then Exit Sub
                                                                                                                         Dim blStatusStr As String = ""
                                                                                                                         Select Case ConvStatus.ConversionStatus
                                                                                                                             Case VolumeConversionStatus.FullyDecrypted : blStatusStr = GetValueFromLanguageData("BitLockerCCP.CCP_VolStat_Decrypted")
                                                                                                                             Case VolumeConversionStatus.FullyEncrypted : blStatusStr = GetValueFromLanguageData("BitLockerCCP.CCP_VolStat_Encrypted")
                                                                                                                             Case VolumeConversionStatus.DecryptionInProgress : blStatusStr = GetValueFromLanguageData("BitLockerCCP.CCP_VolStat_DecryptionInProgress")
                                                                                                                             Case VolumeConversionStatus.EncryptionInProgress : blStatusStr = GetValueFromLanguageData("BitLockerCCP.CCP_VolStat_EncryptionInProgress")
                                                                                                                             Case VolumeConversionStatus.DecryptionPaused : blStatusStr = GetValueFromLanguageData("BitLockerCCP.CCP_VolStat_DecryptionPaused")
                                                                                                                             Case VolumeConversionStatus.EncryptionPaused : blStatusStr = GetValueFromLanguageData("BitLockerCCP.CCP_VolStat_EncryptionPaused")
                                                                                                                         End Select
                                                                                                                         lblConversionStatus.Text = blStatusStr

                                                                                                                         Dim divider As Integer = 10000
                                                                                                                         lblPercentEncrypted.Text = String.Format("{0}%", Math.Round(ConvStatus.EncryptionPercentage / divider, 2))
                                                                                                                         If ConvStatus.EncryptionPercentage / divider <= pbEncrypted.Maximum Then pbEncrypted.Value = ConvStatus.EncryptionPercentage / divider
                                                                                                                     End Sub)
                                                          End Function)
        If DecryptionResult <> Constants.S_OK Then
            Dim errorMessage As String = ""
            Select Case DecryptionResult
                Case Constants.FVE_E_LOCKED_VOLUME : errorMessage = "This volume is locked."
                Case Constants.FVE_E_AUTOUNLOCK_ENABLED : errorMessage = "This volume cannot be decrypted because keys used to automatically unlock data volumes are available."
                Case Else : errorMessage = New Win32Exception(BitConverter.ToInt32(BitConverter.GetBytes(DecryptionResult), 0)).Message
            End Select
            MessageBox.Show(errorMessage, Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        Close()
    End Sub

    Private Function GetDeviceIdFromSystemVolume() As String
        Dim DeviceId As String = ""

        Dim EncryptedVolumeMOC As ManagementObjectCollection = WMIHelper.GetResultsFromManagementQuery(String.Format("SELECT DeviceID FROM Win32_EncryptableVolume WHERE DriveLetter = {0}{1}{0}", Quote, WMIHelper.GetEscapedValue(Environment.GetEnvironmentVariable("SYSTEMDRIVE"))), "root\cimv2\Security\MicrosoftVolumeEncryption")
        If EncryptedVolumeMOC Is Nothing Then Return DeviceId

        DeviceId = WMIHelper.GetObjectValue(EncryptedVolumeMOC(0), "DeviceID")

        Return DeviceId
    End Function

    Private Function GetPersistentVolumeIdFromSystemVolume() As String
        Dim PersistentVolumeId As String = ""

        Dim EncryptedVolumeMOC As ManagementObjectCollection = WMIHelper.GetResultsFromManagementQuery(String.Format("SELECT PersistentVolumeId FROM Win32_EncryptableVolume WHERE DriveLetter = {0}{1}{0}", Quote, WMIHelper.GetEscapedValue(Environment.GetEnvironmentVariable("SYSTEMDRIVE"))), "root\cimv2\Security\MicrosoftVolumeEncryption")
        If EncryptedVolumeMOC Is Nothing Then Return PersistentVolumeId

        PersistentVolumeId = WMIHelper.GetObjectValue(EncryptedVolumeMOC(0), "PersistentVolumeID")

        Return PersistentVolumeId
    End Function
End Class
