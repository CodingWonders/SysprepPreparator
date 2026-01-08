Imports System.Diagnostics.Eventing
Imports System.Diagnostics.Eventing.Reader
Imports System.IO

Namespace Helpers.PreparationTasks

    ''' <summary>
    ''' The Event Log Preparation Task (PT)
    ''' </summary>
    ''' <remarks></remarks>
    Public Class EventLogPT
        Inherits PreparationTask

        ''' <summary>
        ''' Clears the event log
        ''' </summary>
        ''' <returns>Whether the operation succeeded</returns>
        ''' <remarks></remarks>
        Private Function ClearEventLogs() As Boolean
            DynaLog.LogMessage("Preparing to clear event logs...")
            Try
                DynaLog.LogMessage("Getting event logs...")
                Dim EventLogs() As String = GetSystemEventLogs()
                If EventLogs.Count > 0 Then
                    Using session As New EventLogSession()
                        For Each EventLogEntry In EventLogs
                            ReportSubProcessStatus(String.Format(GetValueFromLanguageData("EventLogPT_SubProcessReporting.SPR_Message4"),
                                                                 Array.IndexOf(EventLogs, EventLogEntry) + 1,
                                                                 EventLogs.Count,
                                                                 Math.Round((Array.IndexOf(EventLogs, EventLogEntry) / EventLogs.Count) * 100, 0),
                                                                 EventLogEntry))
                            Try
                                DynaLog.LogMessage("Clearing event log " & EventLogEntry & "...")
                                session.ClearLog(EventLogEntry)
                            Catch ex As Exception
                                DynaLog.LogMessage("Could not clear event log " & EventLogEntry & ". Message: " & ex.Message)
                            End Try
                        Next
                    End Using
                End If
                Return True
            Catch ex As Exception
                DynaLog.LogMessage("An error occurred. Message: " & ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Exports event logs to a path specified by the user
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub ExportEventLogs()
            DynaLog.LogMessage("Preparing to export event logs...")
            Dim targetEvtxPath As String = ShowFolderBrowserDialog(GetValueFromLanguageData("EventLogPT.FBD_TargetEvtxPath"), True)
            DynaLog.LogMessage("Target path for EVTXs: " & targetEvtxPath)
            If targetEvtxPath = "" Then
                DynaLog.LogMessage("No EVTX Path is specified. Aborting procedure!")
                Exit Sub
            End If

            DynaLog.LogMessage("Determining if directory contains files...")
            If Directory.Exists(targetEvtxPath) AndAlso Directory.GetFiles(targetEvtxPath, "*.evtx").Count > 0 Then DynaLog.LogMessage("Directory exists and there are already EVTX files.")

            DynaLog.LogMessage("Getting event logs...")
            Dim EventLogs() As String = GetSystemEventLogs()
            DynaLog.LogMessage("Event Log Count: " & EventLogs.Count)
            If EventLogs.Count > 0 Then
                Using session As New EventLogSession()
                    For Each EventLogEntry In EventLogs
                        ReportSubProcessStatus(String.Format(GetValueFromLanguageData("EventLogPT_SubProcessReporting.SPR_Message3"),
                                                             Array.IndexOf(EventLogs, EventLogEntry) + 1,
                                                             EventLogs.Count,
                                                             Math.Round((Array.IndexOf(EventLogs, EventLogEntry) / EventLogs.Count) * 100, 0),
                                                             EventLogEntry))

                        Dim evtxFileName As String = String.Format("{0}.evtx", EventLogEntry.Replace("/", "--"))
                        DynaLog.LogMessage("Determining if evtx file " & evtxFileName & " exists...")
                        If File.Exists(Path.Combine(targetEvtxPath, evtxFileName)) Then
                            DynaLog.LogMessage("File exists. Attempting to delete...")
                            Try
                                File.Delete(Path.Combine(targetEvtxPath, evtxFileName))
                            Catch ex As Exception
                                DynaLog.LogMessage("Could not delete file " & targetEvtxPath & ". This log will not be exported")
                                Continue For
                            End Try
                        End If
                        DynaLog.LogMessage("File does not exist.")
                        DynaLog.LogMessage("Exporting " & EventLogEntry & " to " & Path.Combine(targetEvtxPath, evtxFileName) & "...")
                        session.ExportLogAndMessages(EventLogEntry, PathType.LogName, "*", Path.Combine(targetEvtxPath, evtxFileName), True, My.Computer.Info.InstalledUICulture)
                    Next
                End Using
            End If
        End Sub

        ''' <summary>
        ''' Exports and then cleans up the target system event log
        ''' </summary>
        ''' <returns>Whether the cleanup process succeeded</returns>
        ''' <remarks>Event log cleanup will not be performed when in test mode</remarks>
        Public Overrides Function RunPreparationTask() As PreparationTaskStatus
            ReportSubProcessStatus(GetValueFromLanguageData("EventLogPT_SubProcessReporting.SPR_Message1"))
            If Not IsInAutoMode Then ExportEventLogs()              ' exporting event logs will not be done in auto mode for now
            If IsInTestMode Then Return PreparationTaskStatus.Skipped
            ReportSubProcessStatus(GetValueFromLanguageData("EventLogPT_SubProcessReporting.SPR_Message2"))
            Return If(ClearEventLogs(), PreparationTaskStatus.Succeeded, PreparationTaskStatus.Failed)
        End Function

        ''' <summary>
        ''' Gets the names of the event logs in the system
        ''' </summary>
        ''' <returns>The event log names</returns>
        ''' <remarks></remarks>
        Private Function GetSystemEventLogs() As String()
            Return New EventLogSession().GetLogNames().ToArray()
        End Function
    End Class

End Namespace
