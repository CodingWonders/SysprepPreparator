Imports System.IO
Imports Microsoft.Dism

Namespace Helpers.PreparationTasks

    ''' <summary>
    ''' The DISM Component Store Cleanup Preparation Task (PT)
    ''' </summary>
    ''' <remarks></remarks>
    Public Class DismComponentCleanupPT
        Inherits PreparationTask

        ''' <summary>
        ''' Launches DISM
        ''' </summary>
        ''' <returns>Whether the process succeeded</returns>
        ''' <remarks>This will not launch when in test mode</remarks>
        Public Overrides Function RunPreparationTask() As PreparationTaskStatus
            If IsInTestMode Then Return PreparationTaskStatus.Skipped
            DynaLog.LogMessage("Running DISM Component Cleanup...")

            Dim PTStatus As PreparationTaskStatus = PreparationTaskStatus.Failed
            ReportSubProcessStatus(GetValueFromLanguageData("DismComponentCleanupPT_SubProcessReporting.SPR_Message1"))

            Try
                DismApi.Initialize(DismLogLevel.LogErrors)

                Dim progressCallback As DismProgressCallback = Sub(progress As DismProgress)
                                                                   ReportSubProcessStatus(String.Format("{0} ({1}%)", GetValueFromLanguageData("DismComponentCleanupPT_SubProcessReporting.SPR_Message1"), progress.Current / 10))
                                                               End Sub

                Using sysSession As DismSession = DismApi.OpenOnlineSession()
                    DismApi.CleanImage(sysSession, DismCleanImageType.Component, DismCleanImageFlags.ResetBase, progressCallback)
                    PTStatus = PreparationTaskStatus.Succeeded
                End Using
            Catch ex As Exception
                ReportSubProcessStatus(GetValueFromLanguageData("DismComponentCleanupPT_SubProcessReporting.SPR_Message1"))
                PTStatus = If(RunProcess(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "system32", "dism.exe"),
                              "/online /cleanup-image /startcomponentcleanup /resetbase") = PROC_SUCCESS, PreparationTaskStatus.Succeeded, PreparationTaskStatus.Failed)
            Finally
                Try
                    DismApi.Shutdown()
                Catch ex As Exception

                End Try
            End Try

            Return PTStatus
        End Function

    End Class

End Namespace
