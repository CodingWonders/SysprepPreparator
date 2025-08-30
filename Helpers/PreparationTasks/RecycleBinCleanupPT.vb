Imports System.IO

Namespace Helpers.PreparationTasks

    ''' <summary>
    ''' The Recycle Bin Preparation Task (PT)
    ''' </summary>
    ''' <remarks></remarks>
    Public Class RecycleBinCleanupPT
        Inherits PreparationTask

        ''' <summary>
        ''' Attempts to clear the current user's Recycle Bin
        ''' </summary>
        ''' <returns>Whether the process succeeded</returns>
        ''' <remarks>This will not launch when in test mode</remarks>
        Public Overrides Function RunPreparationTask() As Boolean
            DynaLog.LogMessage("Clearing the Recycle Bin of the current user...")
            If IsInTestMode Then Return True
            Try
                For Each drive As DriveInfo In DriveInfo.GetDrives().Where(Function(driveInList) driveInList.IsReady).ToList()
                    Dim recycleBinPath As String = Path.Combine(drive.RootDirectory.FullName, "$Recycle.Bin", GetUserSid(Environment.GetEnvironmentVariable("USERNAME")))
                    Try
                        RemoveRecursive(recycleBinPath)
                    Catch ex As Exception
                        Continue For
                    End Try
                Next
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
    End Class

End Namespace
