Imports System.IO
Imports Microsoft.VisualBasic.ApplicationServices

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
                Dim drives() As DriveInfo = DriveInfo.GetDrives()
                For Each drive As DriveInfo In drives
                    If drive.IsReady Then
                        Dim recycleBinPath As String = Path.Combine(drive.RootDirectory.FullName, GetUserSid(Environment.GetEnvironmentVariable("USERNAME")))
                        Try
                            RemoveRecursive(recycleBinPath)
                        Catch ex As Exception
                        End Try
                    End If
                Next
                Return 1
            Catch ex As Exception
                Return 0
            End Try
        End Function
    End Class

End Namespace
