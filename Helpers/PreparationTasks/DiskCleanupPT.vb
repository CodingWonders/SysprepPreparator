Imports System.IO

Namespace Helpers.PreparationTasks

    ''' <summary>
    ''' The Disk Cleanup Preparation Task (PT)
    ''' </summary>
    ''' <remarks></remarks>
    Public Class DiskCleanupPT
        Inherits PreparationTask

        ''' <summary>
        ''' Launches Disk Cleanup
        ''' </summary>
        ''' <returns>Whether the process succeeded</returns>
        ''' <remarks></remarks>
        Public Overrides Function RunPreparationTask() As Boolean
            DynaLog.LogMessage("Running Disk Cleanup...")
            ReportSubProcessStatus(GetValueFromLanguageData("DiskCleanupPT_SubProcessReporting.SPR_Message1"))
            Return RunProcess(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "system32", "cleanmgr.exe"),
                              "/sagerun:1") = PROC_SUCCESS
        End Function
    End Class

End Namespace
