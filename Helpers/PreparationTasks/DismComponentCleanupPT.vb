Imports System.IO

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
            ReportSubProcessStatus(GetValueFromLanguageData("DismComponentCleanupPT_SubProcessReporting.SPR_Message1"))
            Return If(RunProcess(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "system32", "dism.exe"),
                              "/online /cleanup-image /startcomponentcleanup /resetbase") = PROC_SUCCESS, PreparationTaskStatus.Succeeded, PreparationTaskStatus.Failed)
        End Function

    End Class

End Namespace
