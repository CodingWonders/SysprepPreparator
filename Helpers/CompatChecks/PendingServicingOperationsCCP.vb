Imports Microsoft.Win32
Imports System.IO

Namespace Helpers.CompatChecks

    ''' <summary>
    ''' The Pending Servicing Operations Compatibility Checker Provider (CCP)
    ''' </summary>
    ''' <remarks></remarks>
    Public Class PendingServicingOperationsCCP
        Inherits CompatibilityCheckerProvider

        Private SystemPendingServicingFile As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "WinSxS", "pending.xml")

        ''' <summary>
        ''' Determines whether there are pending servicing operations (like updates or other tasks that rely on Windows CBS).
        ''' App installations that leave temporary files behind are also detected
        ''' </summary>
        ''' <returns>Whether there are pending operations</returns>
        ''' <remarks>
        ''' An event detailing results.
        ''' </remarks>
        Public Overrides Function PerformCompatibilityCheck() As Classes.CompatibilityCheckerProviderStatus
            DynaLog.LogMessage("Detecting if there are pending servicing operations, or temporary files...")
            DynaLog.LogMessage("Pending XML: " & SystemPendingServicingFile)
            ' Check pending.xml first before moving to registry stuff
            If File.Exists(SystemPendingServicingFile) Then
                DynaLog.LogMessage("Pending XML exists. The Windows installation needs a reboot for updates")
                Status.Compatible = False
                Status.StatusMessage = New Classes.StatusMessage(GetValueFromLanguageData("PendingServicingOperationsCCP.CCPTitle"),
                                                                 GetValueFromLanguageData("CCP_NotOK_SysServ"),
                                                                 GetValueFromLanguageData("CCP_NotOK_Resolution_SysServ"),
                                                                 Classes.StatusMessage.StatusMessageSeverity.Warning)
                Return Status
            End If

            DynaLog.LogMessage("Pending XML does not exist")

            ' Our marker doesn't appear to be in WinSxS. May it be in the registry?
            ' Let's check this!

            DynaLog.LogMessage("Proceeding with Session Manager value checking. Temporary files from app installs are most likely going to be caught")

            ' 1. Pending File Rename Operations
            Try
                Dim NTSMRk As RegistryKey = Registry.LocalMachine.OpenSubKey("SYSTEM\CurrentControlSet\Control\Session Manager", False)
                Dim PFROValue As String() = NTSMRk.GetValue("PendingFileRenameOperations", New String() {})
                NTSMRk.Close()
                DynaLog.LogMessage("Pending File Rename Operations: " & PFROValue.Count)
                If PFROValue.Count > 0 Then
                    DynaLog.LogMessage("There are some files in there.")
                    ' A reboot is required to rename files that were in use
                    Status.Compatible = True
                    Status.StatusMessage = New Classes.StatusMessage(GetValueFromLanguageData("PendingServicingOperationsCCP.CCPTitle"),
                                                                     GetValueFromLanguageData("PendingServicingOperationsCCP.CCP_NotOK_NtSmPfro"),
                                                                     GetValueFromLanguageData("PendingServicingOperationsCCP.CCP_NotOK_Resolution_NtSmPfro"),
                                                                     Classes.StatusMessage.StatusMessageSeverity.Info)
                    Return Status
                End If
                DynaLog.LogMessage("There aren't any files in there.")
            Catch ex As Exception
                DynaLog.LogMessage("An error occurred. Message: " & ex.Message)
            End Try

            DynaLog.LogMessage("Proceeding with CBS value. This is less likely to pick any things because pending.xml is more likely to be found")

            ' 2. CBS
            Try
                Dim CBSRk As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing", False)
                Dim RPValue As Integer = CBSRk.GetValue("RebootPending", 0)
                CBSRk.Close()
                DynaLog.LogMessage("RebootPending Value: " & RPValue)
                If RPValue <> 0 Then
                    ' Windows Updates are still pending
                    DynaLog.LogMessage("Value is not 0.")
                    Status.Compatible = False
                    Status.StatusMessage = New Classes.StatusMessage(GetValueFromLanguageData("PendingServicingOperationsCCP.CCPTitle"),
                                                                     GetValueFromLanguageData("PendingServicingOperationsCCP.CCP_NotOK_SysServ"),
                                                                     GetValueFromLanguageData("PendingServicingOperationsCCP.CCP_NotOK_Resolution_SysServ"),
                                                                     Classes.StatusMessage.StatusMessageSeverity.Warning)
                    Return Status
                End If
                DynaLog.LogMessage("Value is 0.")

                DynaLog.LogMessage("No checks found the indicator. We'll assume everything is alright")

                ' None of our checks failed. We're good to go here!
                Status.Compatible = True
                Status.StatusMessage = New Classes.StatusMessage(GetValueFromLanguageData("PendingServicingOperationsCCP.CCPTitle"),
                                                                 GetValueFromLanguageData("PendingServicingOperationsCCP.CCP_OK"),
                                                                 Classes.StatusMessage.StatusMessageSeverity.Info)
            Catch ex As Exception
                DynaLog.LogMessage("An error occurred. Message: " & ex.Message)
            End Try

            Return Status
        End Function
    End Class

End Namespace
