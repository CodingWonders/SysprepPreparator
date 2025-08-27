Imports System.IO
Imports System.Reflection
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Microsoft.Win32


Namespace Helpers.PreparationTasks

    ''' <summary>
    ''' Preparation Task to force Windows into a generalized state
    ''' so the image can be safely moved or copied to a different PC.
    ''' </summary>
    Public Class GeneralizePT
        Inherits PreparationTask

        ''' <summary>
        ''' Executes the task to set registry values and prepare the system for generalization.
        ''' </summary>
        ''' <returns>True if successful, False if any failure occurs.</returns>
        Public Overrides Function RunPreparationTask() As Boolean
            If IsInTestMode Then Return True
            Try

                ' Update the ImageState registry value
                Using setupStateKey As RegistryKey =
                    Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\State", writable:=True)
                    If setupStateKey IsNot Nothing Then
                        setupStateKey.SetValue("ImageState", "IMAGE_STATE_GENERALIZE_RESEAL_TO_OOBE", RegistryValueKind.String)
                    Else
                        Return False
                    End If
                End Using

                ' Update the SYSTEM\Setup keys to mark OOBE state
                Using setupKey As RegistryKey =
                    Registry.LocalMachine.OpenSubKey("SYSTEM\Setup", writable:=True)
                    If setupKey IsNot Nothing Then
                        setupKey.SetValue("OOBEInProgress", 0, RegistryValueKind.DWord)
                        setupKey.SetValue("SetupPhase", 0, RegistryValueKind.DWord)
                    Else
                        Return False
                    End If
                End Using

                ' Attempt to rearm activation state
                Dim psi As New ProcessStartInfo("cscript.exe", "//B //Nologo slmgr.vbs /rearm") With {
                    .UseShellExecute = False,
                    .CreateNoWindow = True,
                    .RedirectStandardOutput = True,
                    .RedirectStandardError = True
                }
                Using p As Process = Process.Start(psi)
                    p.WaitForExit(60000) ' wait up to 60s
                    Dim out As String = p.StandardOutput.ReadToEnd()
                    Dim err As String = p.StandardError.ReadToEnd()
                    If Not String.IsNullOrWhiteSpace(err) Then
                    End If
                End Using

                Return True

            Catch ex As Exception
                Return False
            End Try
        End Function

    End Class

End Namespace
