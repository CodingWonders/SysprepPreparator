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
        Public Overrides Function RunPreparationTask() As Boolean
            If IsInTestMode Then Return True
            DynaLog.LogMessage("Running DISM Component Cleanup...")
            ReportSubProcessStatus(GetValueFromLanguageData("DismComponentCleanupPT_SubProcessReporting.SPR_Message1"))
            Return RunProcess(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "system32", "dism.exe"),
                              "/online /cleanup-image /startcomponentcleanup /resetbase") = PROC_SUCCESS
        End Function

    End Class

End Namespace
