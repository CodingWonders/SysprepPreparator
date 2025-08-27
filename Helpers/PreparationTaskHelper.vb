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
            {"Stop Windows Explorer", New ExplorerStopperPT()},
            {"Delete Shadow Volumes", New VssAdminShadowVolumeDeletePT()},
            {"DISM Component Store Cleanup", New DismComponentCleanupPT()},
            {"Delete Windows Update Cache Files", New WindowsUpdateTempFileCleanupPT()},
            {"Run Disk Cleanup", New DiskCleanupPT()},
            {"Export and clear event logs", New EventLogPT()},
            {"Clear Recycle Bin", New RecycleBinCleanupPT()},
            {"Generalize the image", New GeneralizePT()}
        }

        ''' <summary>
        ''' Performs the tasks of all the registered Preparation Tasks (PTs)
        ''' </summary>
        ''' <returns>All the success events</returns>
        ''' <remarks></remarks>
        Public Function RunTasks() As List(Of Boolean)
            Dim StatusList As New List(Of Boolean)

            For Each PreparationTaskModule In PreparationTaskModules.Keys
                MainForm.ReportTaskStart(PreparationTaskModule)
                Dim result As Boolean = PreparationTaskModules(PreparationTaskModule).RunPreparationTask()
                StatusList.Add(result)
                MainForm.ReportTaskSuccess(New Dictionary(Of String, Boolean) From {{PreparationTaskModule, result}})
                Threading.Thread.Sleep(100)
            Next

            Return StatusList
        End Function


    End Module

End Namespace
