Imports System.IO
Imports Microsoft.Dism

Namespace Helpers.PreparationTasks

    Public Class EssentialDriverExportPT
        Inherits PreparationTask

        Protected Friend Overrides Property PTWorkDir As String = "EssentialDrivers"

        ''' <summary>
        ''' Gets the information about the installed drivers of the active Windows installation.
        ''' </summary>
        ''' <returns>The installed drivers of the active Windows installation</returns>
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

        ''' <summary>
        ''' Exports all the SCSI adapters and storage controllers from an active installation
        ''' to the PT working directory.
        ''' </summary>
        ''' <returns></returns>
        Public Overrides Function RunPreparationTask() As PreparationTaskStatus
            ReportSubProcessStatus(GetValueFromLanguageData("EssentialDriverExportPT_SubProcessReporting.SPR_Message1"))
            If Not CreateWorkingDirForPT(PTWorkDir) Then Return PreparationTaskStatus.Failed
            ReportSubProcessStatus(GetValueFromLanguageData("EssentialDriverExportPT_SubProcessReporting.SPR_Message2"))
            Dim installedDrivers As DismDriverPackageCollection = GetSystemDrivers()
            Dim installedEssentialDrivers As IEnumerable(Of DismDriverPackage)
            If installedDrivers IsNot Nothing Then
                ' Filter the drivers. We only want the scsi adapters
                DynaLog.LogMessage("Getting SCSI Adapters/Storage Controllers and network controllers from driver list...")
                installedEssentialDrivers = installedDrivers.Where(Function(driver) {"SCSIAdapter", "Net"}.Contains(driver.ClassName))

                ' The driver export operation will export every driver. We don't want this, so we'll intervene
                ' with our functionality.
                If installedEssentialDrivers Is Nothing Then Return PreparationTaskStatus.Failed
                DynaLog.LogMessage("SCSI Adapters/Storage Controllers in installation: " & installedEssentialDrivers.Count)

                For Each essentialDriver In installedEssentialDrivers
                    ' Extract the name from the original path
                    Dim drvName As String = Path.GetFileName(essentialDriver.OriginalFileName)
                    Dim destinationAdapterPath As String = Path.Combine(BaseWorkDir, PTWorkDir, drvName)
                    ReportSubProcessStatus(String.Format(GetValueFromLanguageData("EssentialDriverExportPT_SubProcessReporting.SPR_Message3"), drvName))
                    DynaLog.LogMessage("Exporting driver " & drvName & " ...")
                    CopyRecursive(Path.GetDirectoryName(essentialDriver.OriginalFileName), destinationAdapterPath)
                Next
            Else
                Return PreparationTaskStatus.Failed
            End If

            Return PreparationTaskStatus.Succeeded
        End Function
    End Class

End Namespace
