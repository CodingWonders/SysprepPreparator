Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports Microsoft.Dism

Namespace Helpers.PreparationTasks

    Public Class DTImageCapturePT
        Inherits PreparationTask

        Private WillPrepareBootImage As Boolean = Environment.GetCommandLineArgs().Contains("/dt_capture")

        ''' <summary>
        ''' Gets the information of a given Windows image.
        ''' </summary>
        ''' <param name="ImageFile">The path of the Windows image to get information about</param>
        ''' <returns>A DismImageInfoCollection object containing information about the Windows image</returns>
        Private Function GetImageFileInformation(ImageFile As String) As DismImageInfoCollection
            DynaLog.LogMessage("Getting information about this image file...")
            DynaLog.LogMessage("- Image file: " & ImageFile)

            Dim ImageCollection As DismImageInfoCollection = Nothing

            If Not File.Exists(ImageFile) Then
                DynaLog.LogMessage("Specified image file does not exist. Stopping...")
                Return Nothing
            End If

            Try
                DynaLog.LogMessage("Initializing API...")
                DismApi.Initialize(DismLogLevel.LogErrors)
                DynaLog.LogMessage("Getting image information...")
                ImageCollection = DismApi.GetImageInfo(ImageFile)
            Catch ex As Exception
                DynaLog.LogMessage("Could not grab image information. Error message: " & ex.Message)
            Finally
                Try
                    DynaLog.LogMessage("Attempting to shut down API...")
                    DismApi.Shutdown()
                Catch ex As Exception
                    ' ignore exceptions
                End Try
            End Try

            If ImageCollection IsNot Nothing Then DynaLog.LogMessage("Amount of images obtained: " & ImageCollection.Count)     ' report this info

            Return ImageCollection
        End Function

        ''' <summary>
        ''' Gets the mounted images in the system.
        ''' </summary>
        ''' <returns>A DismMountedImageInfoCollection object containing mounted images</returns>
        Private Function GetMountedImages() As DismMountedImageInfoCollection
            DynaLog.LogMessage("Getting mounted images...")
            Dim MountedImageCollection As DismMountedImageInfoCollection = Nothing

            Try
                DynaLog.LogMessage("Initializing API...")
                DismApi.Initialize(DismLogLevel.LogErrors)
                DynaLog.LogMessage("Getting mounted image information...")
                MountedImageCollection = DismApi.GetMountedImages()
            Catch ex As Exception
                DynaLog.LogMessage("Could not grab mounted image information. Error message: " & ex.Message)
            Finally
                Try
                    DynaLog.LogMessage("Attempting to shut down API...")
                    DismApi.Shutdown()
                Catch ex As Exception
                    ' ignore exceptions
                End Try
            End Try

            Return MountedImageCollection
        End Function

        ''' <summary>
        ''' Mounts a specified Windows image to a specified directory.
        ''' </summary>
        ''' <param name="ImageFile">The path of the Windows image to mount</param>
        ''' <param name="Index">The index of the Windows image to mount</param>
        ''' <param name="MountDir">The directory to mount the Windows image to</param>
        ''' <param name="ReadOnlyMount">(Optional) Whether to perform a read-only mount. Changes will not be saved if readonly is used</param>
        ''' <param name="ProgressOutput">(Optional) A callback for the mount operation</param>
        ''' <returns>True if the mount operation succeeded, False otherwise.</returns>
        ''' <remarks>
        ''' This function will return False:
        ''' <list type="bullet">
        '''     <item>
        '''         If the Windows image does not exist
        '''     </item>
        '''     <item>
        '''         <term>If a bad index is used</term>
        '''         <description>Indexes lower than 1 and those that exceed the total amount of image(s) of <paramref name="ImageFile"/> are treated as bad</description>
        '''     </item>
        '''     <item>
        '''         If the mount directory does not exist and it could not be created
        '''     </item>
        ''' </list>
        ''' </remarks>
        Private Function MountImage(ImageFile As String, Index As Integer, MountDir As String, Optional ReadOnlyMount As Boolean = False, Optional ProgressOutput As DismProgressCallback = Nothing) As Boolean
            DynaLog.LogMessage("Preparing to mount Windows image...")
            DynaLog.LogMessage("- Image file to mount: " & ImageFile)
            DynaLog.LogMessage("- Image index: " & Index)
            DynaLog.LogMessage("- Mount Directory: " & MountDir)
            DynaLog.LogMessage("- Perform read-only mount? " & If(ReadOnlyMount, "Yes", "No"))
            DynaLog.LogMessage("- Is a progress callback defined? " & If(ProgressOutput IsNot Nothing, "Yes", "No"))

            If Not File.Exists(ImageFile) Then
                DynaLog.LogMessage("Image file does not exist. Stopping...")
                Return False
            End If
            If Index < 1 Then Return False      ' For now we'll only check that. Then we can check index count

            If Not Directory.Exists(MountDir) Then
                DynaLog.LogMessage("Mount directory does not exist. Attempting to create it...")
                ' If the mount directory does not exist, we'll try creating it. If we couldn't,
                ' we simply give up.
                Try
                    Directory.CreateDirectory(MountDir)
                Catch ex As Exception
                    DynaLog.LogMessage("Could not create the directory. Error message: " & ex.Message)
                    Return False
                End Try
            End If

            DynaLog.LogMessage("Getting information about this image file...")
            Dim infoCollection As DismImageInfoCollection = GetImageFileInformation(ImageFile)
            If infoCollection Is Nothing OrElse Index > infoCollection.Count Then
                DynaLog.LogMessage("Either we could not grab image information or specified index exceeds image count. Stopping...")
                Return False
            End If

            Dim mounted As Boolean = False

            Try
                DynaLog.LogMessage("Initializing API...")
                DismApi.Initialize(DismLogLevel.LogErrors)
                DynaLog.LogMessage("Mounting Windows image...")
                ' If we haven't defined a callback object for the operation progress, we call
                ' the API without passing it anything. Otherwise, we pass it.
                If ProgressOutput IsNot Nothing Then
                    DismApi.MountImage(ImageFile, MountDir, Index, ReadOnlyMount, progressCallback:=ProgressOutput)
                Else
                    DismApi.MountImage(ImageFile, MountDir, Index, ReadOnlyMount)
                End If

                mounted = True
            Catch ex As Exception
                DynaLog.LogMessage("Could not mount Windows image. Error message: " & ex.Message)
            Finally
                Try
                    DynaLog.LogMessage("Attempting to shut down API...")
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

        ''' <summary>
        ''' Updates the Boot Configuration Data to boot to the DT PE image.
        ''' </summary>
        ''' <param name="sdiPath">The path of a boot.sdi file used by Windows PE</param>
        ''' <param name="bootImagePath">The path of a Windows PE image</param>
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

        ''' <summary>
        ''' Unmounts a specified Windows image.
        ''' </summary>
        ''' <param name="MountDir">The directory the Windows image is mounted on</param>
        ''' <param name="Commit">Whether to save the changes of the Windows image that has been mounted</param>
        ''' <param name="ProgressOutput">(Optional) A callback for the unmount operation</param>
        ''' <returns>True if the unmount operation succeeded, False otherwise.</returns>
        ''' <remarks>
        ''' This function will return False:
        ''' <list type="bullet">
        '''     <item>
        '''         If the mount directory does not exist
        '''     </item>
        '''     <item>
        '''         If the specified mount directory is not in the list of mounted images
        '''     </item>
        ''' </list>
        ''' </remarks>
        Private Function UnmountImage(MountDir As String, Commit As Boolean, Optional ProgressOutput As DismProgressCallback = Nothing) As Boolean
            DynaLog.LogMessage("Preparing to unmount Windows image...")
            DynaLog.LogMessage("- Mount Directory: " & MountDir)
            DynaLog.LogMessage("- Commit changes? " & If(Commit, "Yes", "No"))
            DynaLog.LogMessage("- Is a progress callback defined? " & If(ProgressOutput IsNot Nothing, "Yes", "No"))

            If Not Directory.Exists(MountDir) Then
                DynaLog.LogMessage("Mount directory does not exist. Stopping...")
                Return False
            End If

            DynaLog.LogMessage("Getting mounted images to see if the mount directory is there...")
            Dim mountedImages As DismMountedImageInfoCollection = GetMountedImages()
            If mountedImages Is Nothing OrElse mountedImages.FirstOrDefault(Function(mountedImage) mountedImage.MountPath = MountDir) Is Nothing Then
                ' Our image is not in the mounted images list. Don't do anything else
                DynaLog.LogMessage("Either we could not get mounted image information or the mount directory is not in the list. Stopping...")
                Return False
            End If

            Dim unmounted As Boolean = False

            Try
                DynaLog.LogMessage("Initializing API...")
                DismApi.Initialize(DismLogLevel.LogErrors)
                DynaLog.LogMessage("Unmounting Windows image...")
                ' If we haven't defined a callback object for the operation progress, we call
                ' the API without passing it anything. Otherwise, we pass it.
                If ProgressOutput IsNot Nothing Then
                    DismApi.UnmountImage(MountDir, Commit, ProgressOutput)
                Else
                    DismApi.UnmountImage(MountDir, Commit)
                End If
                unmounted = True
            Catch ex As Exception
                DynaLog.LogMessage("Could not unmount Windows image. Error message: " & ex.Message)
            Finally
                Try
                    DynaLog.LogMessage("Attempting to shut down API...")
                    DismApi.Shutdown()
                Catch ex As Exception
                    ' ignore exceptions
                End Try
            End Try

            Return unmounted
        End Function

        ''' <summary>
        ''' Performs cleanup operations for a specified folder if this Preparation Task fails.
        ''' </summary>
        ''' <param name="folder">The folder to clean up</param>
        Private Sub CleanupOnFailure(folder As String)
            Try
                RemoveRecursive(folder)
            Catch ex As Exception

            End Try
        End Sub

        ''' <summary>
        ''' Prepares a DISMTools PE image to load the image capture script on startup.
        ''' </summary>
        ''' <returns>Whether the operation succeeded</returns>
        Public Overrides Function RunPreparationTask() As Boolean
            If Not WillPrepareBootImage Then
                DynaLog.LogMessage("The boot image will not be prepared. Stopping...")
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

            If Not File.Exists(sourceFile) Then
                DynaLog.LogMessage("Source file " & sourceFile & " does not exist. Stopping...")
                Return False
            End If
            If Not Directory.Exists(destinationFolder) Then
                DynaLog.LogMessage("Destination folder does not exist. Attempting to create it...")
                Try
                    Directory.CreateDirectory(destinationFolder)
                Catch ex As Exception
                    DynaLog.LogMessage("Could not create destination folder. Error message: " & ex.Message)
                    Return False
                End Try
            End If

            Try
                ReportSubProcessStatus("Copying image and boot files...")
                DynaLog.LogMessage("Copying boot image to destination...")
                File.Copy(sourceFile, destinationFile, True)
                ' To make sure we can make changes we'll get file attributes and remove the readonly attribute
                ' in the destination, if present.
                DynaLog.LogMessage("Getting attributes and removing readonly attribute (if present)...")
                Dim attrs As FileAttributes = File.GetAttributes(destinationFile)
                If (attrs And FileAttributes.ReadOnly) = FileAttributes.ReadOnly Then
                    File.SetAttributes(destinationFile, attrs And Not FileAttributes.ReadOnly)
                End If

                DynaLog.LogMessage("Copying boot files to destination...")
                ' To modify the BCD to allow the system to boot to the DT PE, we need to copy the boot files
                ' over to the local disk. Then we call bcdedit as usual. Since we do a recursive copy, we'll
                ' simply call an external method.
                CopyRecursive(bootSourceFolder, bootFileDestinationFolder)

                If Not IsInTestMode Then
                    ReportSubProcessStatus("Updating boot configuration data...")
                    DynaLog.LogMessage("Sysprep Preparator is not in test mode. Proceeding with BCD update...")
                    ' Update BCD only when we AREN'T in test mode.
                    UpdateBcd(Path.Combine(bootFileDestinationFolder, "boot.sdi"), destinationFile)
                End If

                DynaLog.LogMessage("Mounting DT PE image...")
                ReportSubProcessStatus("Mounting Windows image...")
                MountImage(destinationFile, 1, destinationMountDir, False, Sub(progress As DismProgress)
                                                                               If progress.Current > 100 Then Exit Sub
                                                                               DynaLog.LogMessage("Mount operation progress: " & progress.Current & "%")
                                                                               ReportSubProcessStatus(String.Format("Mounting Windows image... ({0}%)", progress.Current))
                                                                           End Sub)
                ' Perform modifications to image
                DynaLog.LogMessage("Allowing the DT PE to load the image capture script automatically...")
                DynaLog.LogMessage("Modifying boot image to load capture script...")
                ModifyImage(destinationMountDir)
                ' Unmount the image
                DynaLog.LogMessage("Unmounting DT PE image...")
                ReportSubProcessStatus("Unmounting Windows image...")
                UnmountImage(destinationMountDir, True, Sub(progress As DismProgress)
                                                            If (progress.Current / 2) > 100 Then Exit Sub
                                                            DynaLog.LogMessage("Unmount operation progress - reported by API: " & progress.Current & "% - actual progress: " & (progress.Current / 2) & "%")
                                                            ReportSubProcessStatus(String.Format("Unmounting Windows image... ({0}%)", Math.Round((progress.Current / 2), 0)))
                                                        End Sub)


            Catch ex As Exception
                DynaLog.LogMessage("An error occurred while preparing the Windows image. Error message: " & ex.Message)
                DynaLog.LogMessage("Cleaning up files...")
                CleanupOnFailure(destinationFolder)
                Return False
            End Try

            Return True
        End Function

        ''' <summary>
        ''' Modifies the Windows image mounted to a given mount directory to prepare it for image capture
        ''' </summary>
        ''' <param name="mountDir">The location of the mounted Windows image</param>
        Private Sub ModifyImage(mountDir As String)
            Try
                Directory.CreateDirectory(Path.Combine(mountDir, "SysprepPrepTool"))
            Catch ex As Exception
                ' TODO  implement dynalog
            End Try
        End Sub
    End Class

End Namespace