Imports System.IO

Namespace Helpers.PreparationTasks

    ''' <summary>
    ''' The Windows Update Cleanup Preparation Task (PT)
    ''' </summary>
    ''' <remarks></remarks>
    Public Class WindowsUpdateTempFileCleanupPT
        Inherits PreparationTask

        ''' <summary>
        ''' Clears up the Windows Update Cache
        ''' </summary>
        ''' <returns>Whether the process succeeded</returns>
        ''' <remarks></remarks>
        Public Overrides Function RunPreparationTask() As PreparationTaskStatus
            DynaLog.LogMessage("Clearing Windows Update Cache...")
            If IsInTestMode Then Return PreparationTaskStatus.Skipped
            Dim downloadPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "SoftwareDistribution\Download")
            Return If(RemoveRecursive(downloadPath), PreparationTaskStatus.Succeeded, PreparationTaskStatus.Failed)
        End Function

    End Class

End Namespace
