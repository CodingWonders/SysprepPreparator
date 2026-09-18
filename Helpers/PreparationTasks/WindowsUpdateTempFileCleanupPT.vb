Imports System.IO
Imports System.ServiceProcess

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
        Public Overrides Function RunPreparationTask() As PreparationTaskStatus
            DynaLog.LogMessage("Clearing Windows Update Cache...")
            If IsInTestMode Then Return PreparationTaskStatus.Skipped

            Dim SystemServices() As ServiceController = ServiceController.GetServices(),
                ServiceNamesToAlter() As String = New String(1) {"wuauserv", "BITS"}

            For Each ServiceToAlter In ServiceNamesToAlter
                ReportSubProcessStatus(String.Format("Stopping service {0}", ServiceToAlter))

                Dim serviceInfo As ServiceController = ServiceController.GetServices().FirstOrDefault(Function(svc) svc.ServiceName = ServiceToAlter)
                If serviceInfo Is Nothing OrElse serviceInfo.Status <> ServiceControllerStatus.Running Then Continue For

                Dim svcStopExitCode As Integer = RunProcess(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "system32", "net.exe"), String.Format("stop {0}", ServiceToAlter), HideWindow:=True)
                If svcStopExitCode <> 0 Then
                    DynaLog.LogMessage("Could not stop service " & ServiceToAlter & ". The process exited with code " & svcStopExitCode)
                End If
            Next

            ReportSubProcessStatus("Clearing Windows Update cache...")
            Dim downloadPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "SoftwareDistribution\Download")
            If Not RemoveRecursive(downloadPath) Then Return PreparationTaskStatus.Failed

            ' start bits to get rid of download jobs
            Dim bitsSvc As ServiceController = SystemServices.FirstOrDefault(Function(svc) svc.ServiceName = "BITS"),
                auSvc As ServiceController = SystemServices.FirstOrDefault(Function(svc) svc.ServiceName = "wuauserv")
            ReportSubProcessStatus("Starting service BITS")
            If bitsSvc IsNot Nothing Then
                If bitsSvc.Status = ServiceControllerStatus.Stopped Then
                    Dim svcStopExitCode As Integer = RunProcess(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "system32", "net.exe"), "start BITS", HideWindow:=True)
                    If svcStopExitCode <> 0 Then
                        DynaLog.LogMessage("Could not start BITS. The process exited with code " & svcStopExitCode)
                    End If
                End If
            End If

            ' clear all update jobs
            ReportSubProcessStatus("Clearing BITS transfer queue...")
            If RunProcess(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "system32", "bitsadmin.exe"), "/RESET /ALLUSERS", HideWindow:=True) <> 0 Then
                Return PreparationTaskStatus.Failed
            End If

            ' restart automatic updates
            ReportSubProcessStatus("Starting service wuauserv")
            If auSvc IsNot Nothing Then
                If auSvc.Status = ServiceControllerStatus.Stopped Then
                    Dim svcStopExitCode As Integer = RunProcess(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "system32", "net.exe"), "start wuauserv", HideWindow:=True)
                    If svcStopExitCode <> 0 Then
                        DynaLog.LogMessage("Could not start WUAUSERV. The process exited with code " & svcStopExitCode)
                    End If
                End If
            End If

            Return PreparationTaskStatus.Succeeded
        End Function

    End Class

End Namespace
