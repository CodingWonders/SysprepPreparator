Imports System.ComponentModel
Imports System.IO
Imports Microsoft.Dism

Namespace Helpers.PreparationTasks

    Public Class DTImageCapturePT
        Inherits PreparationTask

        Private WillPrepareBootImage As Boolean = Environment.GetCommandLineArgs().Contains("/dt_capture")

        Private Function GetImageFileInformation(ImageFile As String) As DismImageInfoCollection
            Dim ImageCollection As DismImageInfoCollection = Nothing

            If Not File.Exists(ImageFile) Then Return Nothing

            Try
                DismApi.Initialize(DismLogLevel.LogErrors)
                ImageCollection = DismApi.GetImageInfo(ImageFile)
            Catch ex As Exception

            Finally
                Try
                    DismApi.Shutdown()
                Catch ex As Exception
                    ' ignore exceptions
                End Try
            End Try

            Return ImageCollection
        End Function

        Private Function GetMountedImages() As DismMountedImageInfoCollection
            Dim MountedImageCollection As DismMountedImageInfoCollection = Nothing

            Try
                DismApi.Initialize(DismLogLevel.LogErrors)
                MountedImageCollection = DismApi.GetMountedImages()
            Catch ex As Exception

            Finally
                Try
                    DismApi.Shutdown()
                Catch ex As Exception
                    ' ignore exceptions
                End Try
            End Try

            Return MountedImageCollection
        End Function

        Private Function MountImage(ImageFile As String, Index As Integer, MountDir As String, Optional ReadOnlyMount As Boolean = False, Optional ProgressOutput As DismProgressCallback = Nothing) As Boolean
            If Not File.Exists(ImageFile) Then
                ' TODO  DynaLog call
                Return False
            End If
            If Index < 1 Then Return False      ' For now we'll only check that. Then we can check index count

            If Not Directory.Exists(MountDir) Then
                ' If the mount directory does not exist, we'll try creating it. If we couldn't,
                ' we simply give up.
                Try
                    Directory.CreateDirectory(MountDir)
                Catch ex As Exception
                    Return False
                End Try
            End If

            Dim infoCollection As DismImageInfoCollection = GetImageFileInformation(ImageFile)
            If infoCollection Is Nothing OrElse Index > infoCollection.Count Then

                Return False
            End If

            Dim mounted As Boolean = False

            Try
                DismApi.Initialize(DismLogLevel.LogErrors)
                ' If we haven't defined a callback object for the operation progress, we call
                ' the API without passing it anything. Otherwise, we pass it.
                If ProgressOutput IsNot Nothing Then
                    DismApi.MountImage(ImageFile, MountDir, Index, ReadOnlyMount, progressCallback:=ProgressOutput)
                Else
                    DismApi.MountImage(ImageFile, MountDir, Index, ReadOnlyMount)
                End If

                mounted = True
            Catch ex As Exception
                ' TODO  implement dynalog
            Finally
                Try
                    DismApi.Shutdown()
                Catch ex As Exception
                    ' Ignore exceptions
                End Try
            End Try

            Return mounted
        End Function

        ''' <summary>
        ''' Runs BCDEdit with the provided arguments
        ''' </summary>
        ''' <param name="Arguments">The command-line arguments to pass to the command</param>
        ''' <param name="DontWorryBeHappy">(Optional) Determines whether or not to throw an exception if the process exits with a code different from 0</param>
        ''' <remarks>Arguments need to be passed. Otherwise, BCDEdit will simply return a basic list of entries on the BCD</remarks>
        Private Sub RunBCDConfigurator(Arguments As String, Optional DontWorryBeHappy As Boolean = False)
            DynaLog.LogMessage("Preparing to modify boot configuration data...")
            DynaLog.LogMessage("- Arguments: " & Arguments)
            DynaLog.LogMessage("- Ignore error messages? " & If(DontWorryBeHappy, "Yes", "No"))
            Try
                Dim bcdEditExitCode As Integer = RunProcess(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "system32", "bcdedit.exe"), Arguments, HideWindow:=Debugger.IsAttached, Inconditional:=DontWorryBeHappy)

                If bcdEditExitCode <> PROC_SUCCESS Then Throw New Win32Exception(bcdEditExitCode)
            Catch ex As Exception
                Throw
            End Try
        End Sub

        Private Sub UpdateBcd(sdiPath As String, bootImagePath As String)
            Dim targetGuidOutput As String = "",
                targetGuid As String = ""

            ' Configure bootmgr settings
            RunBCDConfigurator("/set {default} bootmenupolicy legacy", True)
            RunBCDConfigurator("/set {current} bootmenupolicy legacy", True)
            RunBCDConfigurator("/set {bootmgr} timeout 3", True)

            ' Configure ramdisk
            RunBCDConfigurator("/create {ramdiskoptions}")
            RunBCDConfigurator("/set {ramdiskoptions} ramdisksdidevice partition=" & Environment.GetEnvironmentVariable("SYSTEMDRIVE"))
            RunBCDConfigurator("/set {ramdiskoptions} ramdisksdipath " & sdiPath.Replace(Path.GetPathRoot(sdiPath), "\"))

            ' Grab BCD entry GUID. For this, we don't use the BCD configurator as we need to parse the output and extract it from there
            Dim bcdeditCreateBcdEntryProc As New Process() With {
                .StartInfo = New ProcessStartInfo() With {
                    .FileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "system32", "bcdedit.exe"),
                    .Arguments = String.Format("/create /d {0} /application osloader", ControlChars.Quote & "Sysprep Preparation Tool -- DISMTools Image Capture" & ControlChars.Quote),
                    .RedirectStandardOutput = True,
                    .RedirectStandardError = True
                }
            }
            bcdeditCreateBcdEntryProc.Start()
            targetGuidOutput = bcdeditCreateBcdEntryProc.StandardOutput.ReadToEnd()
            bcdeditCreateBcdEntryProc.WaitForExit()

            If bcdeditCreateBcdEntryProc.ExitCode <> 0 Then Throw New Exception("bcdedit finished with exit code " & Hex(bcdeditCreateBcdEntryProc.ExitCode))

            Dim startIndex As Integer = targetGuidOutput.IndexOf("{"),
                endIndex As Integer = targetGuidOutput.LastIndexOf("}")

            targetGuid = targetGuidOutput.Substring(startIndex, endIndex - startIndex + 1)
            ' TODO  dynalog log target guid

            ' Update our BCD entry
            Dim osloaderPath As String = If(Environment.GetEnvironmentVariable("FIRMWARE_TYPE") = "UEFI",
                "\Windows\system32\Boot\winload.efi",
                "\Windows\system32\winload.exe")
            RunBCDConfigurator(String.Format("/set {0} device ramdisk=[{1}]\{2},{{ramdiskoptions}}", targetGuid, Environment.GetEnvironmentVariable("SYSTEMDRIVE"), bootImagePath))
            RunBCDConfigurator(String.Format("/set {0} osdevice ramdisk=[{1}]\{2},{{ramdiskoptions}}", targetGuid, Environment.GetEnvironmentVariable("SYSTEMDRIVE"), bootImagePath))
            RunBCDConfigurator(String.Format("/set {0} path {1}", targetGuid, osloaderPath))
            RunBCDConfigurator(String.Format("/set {0} locale en-US", targetGuid))
            RunBCDConfigurator(String.Format("/set {0} systemroot \Windows", targetGuid))
            RunBCDConfigurator(String.Format("/set {0} detecthal Yes", targetGuid))
            RunBCDConfigurator(String.Format("/set {0} winpe Yes", targetGuid))

            ' Modify display order
            RunBCDConfigurator(String.Format("/displayorder {0} /addfirst", targetGuid))
            RunBCDConfigurator(String.Format("/default {0}", targetGuid))
        End Sub

        Private Function UnmountImage(MountDir As String, Commit As Boolean, Optional ProgressOutput As DismProgressCallback = Nothing) As Boolean
            If Not Directory.Exists(MountDir) Then Return False

            Dim mountedImages As DismMountedImageInfoCollection = GetMountedImages()
            If mountedImages Is Nothing OrElse mountedImages.FirstOrDefault(Function(mountedImage) mountedImage.MountPath = MountDir) Is Nothing Then
                ' Our image is not in the mounted images list. Don't do anything else
                Return False
            End If

            Dim unmounted As Boolean = False

            Try
                DismApi.Initialize(DismLogLevel.LogErrors)
                ' If we haven't defined a callback object for the operation progress, we call
                ' the API without passing it anything. Otherwise, we pass it.
                If ProgressOutput IsNot Nothing Then
                    DismApi.UnmountImage(MountDir, Commit, ProgressOutput)
                Else
                    DismApi.UnmountImage(MountDir, Commit)
                End If
                unmounted = True
            Catch ex As Exception
                ' TODO  implement dynalog
            Finally
                Try
                    DismApi.Shutdown()
                Catch ex As Exception
                    ' ignore exceptions
                End Try
            End Try

            Return unmounted
        End Function

        Private Sub CleanupOnFailure(folder As String)
            Try
                RemoveRecursive(folder)
            Catch ex As Exception

            End Try
        End Sub

        Public Overrides Function RunPreparationTask() As Boolean
            If Not WillPrepareBootImage Then
                Return True
            End If

            ' We'll adapt the HotInstall code:
            ' 1. We copy files (only boot.wim AND boot files for now)
            ' 2. We set bootmgr entries (NOT IN TEST MODE)
            ' 3. We mount the image
            ' 4. We modify it
            ' 5. We unmount it
            ' Afterwards, the system will be restarted with the image capture script opening immediately.
            ' In this mode, we remove the bootmgr entry and exclude the DT.BT folder so it isn't captured.

            Dim sourceFile As String = Path.Combine(Path.GetPathRoot(Application.StartupPath), "sources", "boot.wim")
            Dim bootSourceFolder As String = String.Format("{0}\{1}", Path.GetPathRoot(sourceFile), "Boot")
            Dim destinationFolder As String = String.Format("{0}\$DISMTOOLS.~BT", Environment.GetEnvironmentVariable("SYSTEMDRIVE"))
            Dim destinationFile As String = String.Format("{0}\boot.wim", destinationFolder)
            Dim bootFileDestinationFolder As String = String.Format("{0}\Boot", destinationFolder)
            Dim destinationMountDir As String = String.Format("{0}\$DISMTOOLS.~WS", Environment.GetEnvironmentVariable("SYSTEMDRIVE"))

            If Not File.Exists(sourceFile) Then Return False
            If Not Directory.Exists(destinationFolder) Then
                Try
                    Directory.CreateDirectory(destinationFolder)
                Catch ex As Exception
                    Return False
                End Try
            End If

            Try
                File.Copy(sourceFile, destinationFile, True)
                ' To make sure we can make changes we'll get file attributes and remove the readonly attribute
                ' in the destination, if present.
                Dim attrs As FileAttributes = File.GetAttributes(destinationFile)
                If (attrs And FileAttributes.ReadOnly) = FileAttributes.ReadOnly Then
                    File.SetAttributes(destinationFile, attrs And Not FileAttributes.ReadOnly)
                End If

                ' To modify the BCD to allow the system to boot to the DT PE, we need to copy the boot files
                ' over to the local disk. Then we call bcdedit as usual. Since we do a recursive copy, we'll
                ' simply call an external method.
                CopyRecursive(bootSourceFolder, bootFileDestinationFolder)

                If Not IsInTestMode Then
                    ' Update BCD only when we AREN'T in test mode.
                    UpdateBcd(Path.Combine(bootFileDestinationFolder, "boot.sdi"), destinationFile)
                End If

                MountImage(destinationFile, 1, destinationMountDir, False, Sub(progress As DismProgress)
                                                                               If progress.Current > 100 Then Exit Sub

                                                                               ' TODO  dynalog logger for mount process
                                                                           End Sub)
                ' Perform modifications to image
                ModifyImage(destinationMountDir)
                ' Unmount the image
                UnmountImage(destinationMountDir, True, Sub(progress As DismProgress)
                                                            If (progress.Current / 2) > 100 Then Exit Sub

                                                            ' TODO  dynalog logger for unmount process
                                                        End Sub)


            Catch ex As Exception
                CleanupOnFailure(destinationFolder)
                Return False
            End Try

            Return True
        End Function

        Private Sub ModifyImage(mountDir As String)
            Try
                Directory.CreateDirectory(Path.Combine(mountDir, "SysprepPrepTool"))
            Catch ex As Exception
                ' TODO  implement dynalog
            End Try
        End Sub
    End Class

End Namespace