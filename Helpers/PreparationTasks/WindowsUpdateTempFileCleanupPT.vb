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
            ' TODO: Implement a more .NET-native version
            If IsInTestMode Then Return True
            Return RunProcess(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "system32", "cmd.exe"),
                              "/c del %WINDIR%\SoftwareDistribution\Download\*.* /F /S /Q") = PROC_SUCCESS
        End Function

    End Class

End Namespace
