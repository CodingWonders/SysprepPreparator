Imports System.IO

Namespace Helpers.PreparationTasks

    ''' <summary>
    ''' The Volume Shadow Copy / Restore Points Preparation Task (PT)
    ''' </summary>
    ''' <remarks></remarks>
    Public Class VssAdminShadowVolumeDeletePT
        Inherits PreparationTask

        ''' <summary>
        ''' Launches the VSC/RP process
        ''' </summary>
        ''' <returns>Whether the process succeeded</returns>
        ''' <remarks>This will not run when in test mode</remarks>
        Public Overrides Function RunPreparationTask() As PreparationTaskStatus
            DynaLog.LogMessage("Clearing shadow volumes/restore points...")
            If IsInTestMode Then Return PreparationTaskStatus.Skipped
            Return If(RunProcess(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "system32", "vssadmin.exe"),
                              "delete shadows /all /quiet", HideWindow:=True, Inconditional:=True) = PROC_SUCCESS, PreparationTaskStatus.Succeeded, PreparationTaskStatus.Failed)
        End Function

    End Class

End Namespace
