Imports Microsoft.Win32

Namespace Helpers.CompatChecks

    ''' <summary>
    ''' The Setup State Compatibility Checker Provider (CCP)
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SetupStateCCP
        Inherits CompatibilityCheckerProvider

        ''' <summary>
        ''' Image Setup State values
        ''' </summary>
        ''' <remarks>
        ''' Refer to https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/windows-setup-states?view=windows-11 for more information on
        ''' Windows Setup states
        ''' </remarks>
        Private Enum ImageState
            ''' <summary>
            ''' The image has successfully been installed. The specialize and oobeSystem configuration passes are complete.
            ''' </summary>
            ''' <remarks>
            ''' This image is not deployable to a computer that has a different hardware configuration because it is now hardware-dependent.
            ''' To deploy this image to a computer that has a different hardware configuration, you must run sysprep /generalize. This last
            ''' statement is not accounted for yet by this utility and will return an error event
            ''' </remarks>
            Complete
            ''' <summary>
            ''' This is the default state for an image in a given phase of Windows Setup that is not yet complete.
            ''' </summary>
            ''' <remarks>
            ''' If a process queries the IMAGE_STATE value and IMG_UNDEPLOYABLE is returned, the image is in one of the following states:
            ''' <list type="number">
            '''     <item>Setup is currently running and has not fully completed the phase. Once a given phase is complete, the IMAGE_STATE will be set to an appropriate completion value</item>
            '''     <item>If queried online when Setup is not running, there was a failure when completing a Setup phase. This image must be reinstalled</item>
            '''     <item>If queried offline, the image did not finish a phase and will never be deployable.</item>
            ''' </list>
            ''' </remarks>
            Undeployable
            ''' <summary>
            ''' The image has successfully completed the generalize configuration pass and will continue into oobeSystem configuration pass when Setup is initiated.
            ''' </summary>
            ''' <remarks></remarks>
            GeneralizeResealToOOBE
            ''' <summary>
            ''' The image has successfully completed the generalize configuration pass and will continue into audit mode when Setup is initiated.
            ''' </summary>
            ''' <remarks></remarks>
            GeneralizeResealToAudit
            ''' <summary>
            ''' The image has successfully completed the specialize pass and will continue into oobeSystem configuration pass when Setup is initiated.
            ''' </summary>
            ''' <remarks></remarks>
            SpecializeResealToOOBE
            ''' <summary>
            ''' The image has successfully completed the specialize configuration pass and will continue into audit mode when Setup is initiated.
            ''' </summary>
            ''' <remarks></remarks>
            SpecializeResealToAudit
            ''' <summary>
            ''' An unknown setup state
            ''' </summary>
            ''' <remarks></remarks>
            Unknown
        End Enum

        ''' <summary>
        ''' Determines whether the system is in an appropriate state for Sysprep preparation
        ''' </summary>
        ''' <returns>An event detailing results</returns>
        ''' <remarks></remarks>
        Public Overrides Function PerformCompatibilityCheck() As Classes.CompatibilityCheckerProviderStatus
            DynaLog.LogMessage("Detecting if the system is in an appropriate setup state...")

            Try
                Dim StateRK As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\State", False)
                Dim StateValue As String = StateRK.GetValue("ImageState")
                DynaLog.LogMessage("State Value in Registry: " & StateValue)
                Dim State As ImageState = StringToImageState(StateValue)
                StateRK.Close()

                DynaLog.LogMessage("State enum value: " & State)

                If {ImageState.GeneralizeResealToAudit, ImageState.SpecializeResealToAudit}.Contains(State) Then
                    DynaLog.LogMessage("Reseals to audits were detected. We are in a good setup state.")
                    Status.Compatible = True
                    Status.StatusMessage = New Classes.StatusMessage(GetValueFromLanguageData("SetupStateCCP.CCPTitle"),
                                                                            GetValueFromLanguageData("SetupStateCCP.CCP_OK"),
                                                                            Classes.StatusMessage.StatusMessageSeverity.Info)

                ElseIf State = ImageState.SpecializeResealToOOBE Then
                    DynaLog.LogMessage("Reseals to audits were not detected. We are not in a good setup state.")
                    Status.Compatible = False
                    Status.StatusMessage = New Classes.StatusMessage(GetValueFromLanguageData("SetupStateCCP.CCPTitle"),
                                                                            GetValueFromLanguageData("SetupStateCCP.CCP_NotOK"),
                                                                            GetValueFromLanguageData("SetupStateCCP.CCP_NotOK_Resolution_SpecReseal"),
                                                                            Classes.StatusMessage.StatusMessageSeverity.Warning)
                Else
                    DynaLog.LogMessage("Reseals to audits were not detected. We are not in a good setup state.")
                    Status.Compatible = False
                    Status.StatusMessage = New Classes.StatusMessage(GetValueFromLanguageData("SetupStateCCP.CCPTitle"),
                                                                            GetValueFromLanguageData("SetupStateCCP.CCP_NotOK"),
                                                                            GetValueFromLanguageData("SetupStateCCP.CCP_NotOK_Resolution_Generic"),
                                                                            Classes.StatusMessage.StatusMessageSeverity.Critical)
                End If
            Catch ex As Exception
                DynaLog.LogMessage("An error occurred. Message: " & ex.Message)
                Status.Compatible = False
                Status.StatusMessage = New Classes.StatusMessage(GetValueFromLanguageData("SetupStateCCP.CCPTitle"),
                                                                        String.Format(GetValueFromLanguageData("SetupStateCCP.CCP_Error"), ex.Message),
                                                                        GetValueFromLanguageData("CCP_Error_Resolution"),
                                                                        Classes.StatusMessage.StatusMessageSeverity.Critical)
            End Try

            Return Status
        End Function

        ''' <summary>
        ''' Converts a Setup state value string into an ImageState enumeration value
        ''' </summary>
        ''' <param name="ImageState">The image state value</param>
        ''' <returns>The enumeration value</returns>
        ''' <remarks></remarks>
        Private Function StringToImageState(ImageState As String) As ImageState
            DynaLog.LogMessage("Converting the string value to an image state enum value...")
            DynaLog.LogMessage("- State: " & ImageState)
            Select Case ImageState
                Case "IMAGE_STATE_COMPLETE"
                    Return SetupStateCCP.ImageState.Complete
                Case "IMAGE_STATE_UNDEPLOYABLE"
                    Return SetupStateCCP.ImageState.Undeployable
                Case "IMAGE_STATE_GENERALIZE_RESEAL_TO_OOBE"
                    Return SetupStateCCP.ImageState.GeneralizeResealToOOBE
                Case "IMAGE_STATE_GENERALIZE_RESEAL_TO_AUDIT"
                    Return SetupStateCCP.ImageState.GeneralizeResealToAudit
                Case "IMAGE_STATE_SPECIALIZE_RESEAL_TO_OOBE"
                    Return SetupStateCCP.ImageState.SpecializeResealToOOBE
                Case "IMAGE_STATE_SPECIALIZE_RESEAL_TO_AUDIT"
                    Return SetupStateCCP.ImageState.SpecializeResealToAudit
            End Select
            Return SetupStateCCP.ImageState.Unknown
        End Function
    End Class

End Namespace
