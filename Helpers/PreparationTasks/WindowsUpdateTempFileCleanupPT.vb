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
            ' Return RunProcess(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "system32", "cmd.exe"),
            '                  "/c del %WINDIR%\SoftwareDistribution\Download\*.* /F /S /Q") = PROC_SUCCESS
            Dim downloadPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "SoftwareDistribution\Download")

            If Directory.Exists(downloadPath) Then
                For Each file In Directory.GetFiles(downloadPath, "*", SearchOption.AllDirectories)
                    Try
                        Delete(downloadPath, file)
                    Catch ex As Exception

                    End Try
                Next
            End If

            Return PROC_SUCCESS
        End Function

    End Class

End Namespace
