Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization
Imports SysprepPreparator.Classes

Namespace Helpers.PreparationTasks

    Public Class UserProfilePT
        Inherits PreparationTask

        ''' <summary>
        ''' Prepares a user profile by clearing temporary items (such as items in the
        ''' Recents list).
        ''' </summary>
        ''' <returns></returns>
        Public Overrides Function RunPreparationTask() As PreparationTaskStatus
            If IsInTestMode Then Return PreparationTaskStatus.Skipped

            ' Remove recents in run box and file explorer
            ReportSubProcessStatus("Clearing Run items...")
            RemoveRegistryItem("HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU", "/va /f")
            ReportSubProcessStatus("Clearing MRU list in File Explorer...")
            RemoveRegistryItem("HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\TypedPaths", "/va /f")

            Return PreparationTaskStatus.Succeeded
        End Function
    End Class

End Namespace
