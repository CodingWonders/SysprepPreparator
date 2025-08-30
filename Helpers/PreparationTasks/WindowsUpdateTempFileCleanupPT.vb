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
        Public Overrides Function RunPreparationTask() As Boolean
            DynaLog.LogMessage("Clearing Windows Update Cache...")
            If IsInTestMode Then Return True
            Dim downloadPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "SoftwareDistribution\Download")
            Return RemoveRecursive(downloadPath)
        End Function

    End Class

End Namespace
