Imports SysprepPreparator.Classes
Imports SysprepPreparator.Helpers.PreparationTasks

Namespace Helpers

    ''' <summary>
    ''' The Preparation Task Helper module: it will launch all the registered PTs
    ''' </summary>
    ''' <remarks></remarks>
    Module PreparationTaskHelper

        ''' <summary>
        ''' The list of registered PTs
        ''' </summary>
        ''' <remarks>To develop and register a PT, read the documentation</remarks>
        Private PreparationTaskModules As New Dictionary(Of String, PreparationTask) From {
            {GetValueFromLanguageData("RegisteredPTs.SysprepStopperPT"), New SysprepStopperPT()},
            {GetValueFromLanguageData("RegisteredPTs.ExplorerStopperPT"), New ExplorerStopperPT()},
            {GetValueFromLanguageData("RegisteredPTs.VssAdminShadowVolumeDeletePT"), New VssAdminShadowVolumeDeletePT()},
            {GetValueFromLanguageData("RegisteredPTs.DismComponentCleanupPT"), New DismComponentCleanupPT()},
            {GetValueFromLanguageData("RegisteredPTs.WindowsUpdateTempFileCleanupPT"), New WindowsUpdateTempFileCleanupPT()},
            {GetValueFromLanguageData("RegisteredPTs.DiskCleanupPT"), New DiskCleanupPT()},
            {GetValueFromLanguageData("RegisteredPTs.EventLogPT"), New EventLogPT()},
            {GetValueFromLanguageData("RegisteredPTs.RecycleBinCleanupPT"), New RecycleBinCleanupPT()}
        }

        ''' <summary>
        ''' Performs the tasks of all the registered Preparation Tasks (PTs)
        ''' </summary>
        ''' <returns>All the success events</returns>
        ''' <remarks></remarks>
        Public Function RunTasks(Optional ProgressStartReporter As Action(Of String) = Nothing,
                                 Optional ProgressFinishedReporter As Action(Of Dictionary(Of String, Boolean)) = Nothing) As List(Of Boolean)
            DynaLog.LogMessage("Preparing to run Preparation Tasks (PTs)...")
            Dim StatusList As New List(Of Boolean)

            For Each PreparationTaskModule In PreparationTaskModules.Keys
                If ProgressStartReporter IsNot Nothing Then ProgressStartReporter.Invoke(PreparationTaskModule)
                DynaLog.LogMessage("PT to run: " & PreparationTaskModules(PreparationTaskModule).GetType().Name)
                Dim result As Boolean = PreparationTaskModules(PreparationTaskModule).RunPreparationTask()
                DynaLog.LogMessage("PT Succeeded? " & result)
                StatusList.Add(result)
                If ProgressFinishedReporter IsNot Nothing Then ProgressFinishedReporter.Invoke(New Dictionary(Of String, Boolean) From {{PreparationTaskModule, result}})
                Threading.Thread.Sleep(100)
            Next

            Return StatusList
        End Function


    End Module

End Namespace
