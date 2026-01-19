Imports System.IO
Imports Microsoft.Dism

Namespace Helpers.PreparationTasks

    Public Class SCSIAdapterDriverExportPT
        Inherits PreparationTask

        Private ReadOnly WorkDir As String = "ScsiAdapter"

        Private Function GetSystemDrivers() As DismDriverPackageCollection
            Dim obtainedDrivers As DismDriverPackageCollection = Nothing
            Try
                DynaLog.LogMessage("Initializing API...")
                DismApi.Initialize(DismLogLevel.LogErrors)
                DynaLog.LogMessage("Getting drivers...")
                Using session As DismSession = DismApi.OpenOnlineSession()
                    ' We only care about the third-party drivers
                    obtainedDrivers = DismApi.GetDrivers(session, False)
                End Using
            Catch ex As Exception
                DynaLog.LogMessage("Could not get drivers. Error message: " & ex.Message)
            Finally
                Try
                    DynaLog.LogMessage("Attempting to shut down API...")
                    DismApi.Shutdown()
                Catch ex As Exception

                End Try
            End Try
            Return obtainedDrivers
        End Function

        Public Overrides Function RunPreparationTask() As PreparationTaskStatus
            ReportSubProcessStatus("Preparing to export drivers...")
            CreateWorkingDirForPT(WorkDir)
            ReportSubProcessStatus("Getting SCSI Adapter/Storage Controller drivers...")
            Dim installedDrivers As DismDriverPackageCollection = GetSystemDrivers()
            Dim installedScsiAdapters As IEnumerable(Of DismDriverPackage)
            If installedDrivers IsNot Nothing Then
                ' Filter the drivers. We only want the scsi adapters
                DynaLog.LogMessage("Getting SCSI Adapters/Storage Controllers from driver list...")
                installedScsiAdapters = installedDrivers.Where(Function(driver) driver.ClassName = "SCSIAdapter")

                ' The driver export operation will export every driver. We don't want this, so we'll intervene
                ' with our functionality.
                If installedScsiAdapters Is Nothing Then Return PreparationTaskStatus.Failed
                DynaLog.LogMessage("SCSI Adapters/Storage Controllers in installation: " & installedScsiAdapters.Count)

                For Each scsiAdapter In installedScsiAdapters
                    ' Extract the name from the original path
                    Dim drvName As String = Path.GetFileName(scsiAdapter.OriginalFileName)
                    Dim destinationAdapterPath As String = Path.Combine(BaseWorkDir, WorkDir, drvName)
                    ReportSubProcessStatus(String.Format("Exporting driver {0} ...", drvName))
                    DynaLog.LogMessage("Exporting driver " & drvName & " ...")
                    CopyRecursive(Path.GetDirectoryName(scsiAdapter.OriginalFileName), destinationAdapterPath)
                Next
            Else
                Return PreparationTaskStatus.Failed
            End If

            Return PreparationTaskStatus.Succeeded
        End Function
    End Class

End Namespace
