Imports SysprepPreparator.Helpers
Imports System.Windows.Forms
Imports System.IO
Imports System.Management
Imports Microsoft.VisualBasic.ControlChars
Imports System.Threading

Namespace Helpers.PreparationTasks

    ''' <summary>
    ''' The base class for Preparation Tasks (PTs)
    ''' </summary>
    ''' <remarks>
    ''' To integrate a Preparation Task into this program, create a class in this namespace
    ''' and inherit this base class. More information can be found in the documentation
    ''' </remarks>
    Public MustInherit Class PreparationTask
        Implements IUserInterfaceInterop, IProcessRunner, IRegistryRunner, IFileProcessor, IWmiUserProcessor

        Public Enum PreparationTaskStatus As Integer
            Succeeded = 0
            Failed = 1
            Skipped = 2
        End Enum

        ''' <summary>
        ''' Runs a preparation task
        ''' </summary>
        ''' <returns>Whether the preparation task succeeded</returns>
        ''' <remarks>This must not be called from this parent class, but from classes that inherit this</remarks>
        Public MustOverride Function RunPreparationTask() As PreparationTaskStatus

        ''' <summary>
        ''' Constant for external processes that run successfully
        ''' </summary>
        ''' <remarks></remarks>
        Protected Friend Const PROC_SUCCESS As Integer = 0

        ''' <summary>
        ''' Determines whether the Sysprep Preparation Tool is in test mode
        ''' </summary>
        ''' <remarks>
        ''' This is useful when prototyping Preparation Tasks so you don't run tasks that are typically
        ''' run on reference computers.
        ''' </remarks>
        Protected Friend IsInTestMode As Boolean = Environment.GetCommandLineArgs().Contains("/test")

        ''' <summary>
        ''' Determines whether the Sysprep Preparation Tool is in Automatic mode
        ''' </summary>
        Protected Friend IsInAutoMode As Boolean = Environment.GetCommandLineArgs().Contains("/auto")

        ''' <summary>
        ''' An event sender for subprocess status changes
        ''' </summary>
        Protected Friend SubProcessReporter As Action(Of String) = Nothing

        ''' <summary>
        ''' Reports a subprocess status change with a given status message.
        ''' </summary>
        ''' <param name="Status">The status message to report</param>
        <CodeAnalysis.SuppressMessage("Style", "IDE0031:Usar propagación de null", Justification:="<pendiente>")>
        Public Sub ReportSubProcessStatus(Status As String)
            ' null propagation is not used here to enable backcompat with vs2012
            If SubProcessReporter IsNot Nothing Then
                SubProcessReporter.Invoke(Status)
            End If
        End Sub

        ''' <summary>
        ''' Shows a file picker to open a file
        ''' </summary>
        ''' <param name="MultiSelect">Whether to allow file picker to select multiple files</param>
        ''' <returns>The path, or paths, of the chosen file</returns>
        ''' <remarks></remarks>
        Public Function ShowOpenFileDialog(Optional MultiSelect As Boolean = False) As Object Implements IUserInterfaceInterop.ShowOpenFileDialog
            Dim ofd As New OpenFileDialog() With {
                .SupportMultiDottedExtensions = True,
                .Multiselect = MultiSelect
            }

            If ofd.ShowDialog() = DialogResult.OK Then
                If MultiSelect Then
                    Return ofd.FileNames
                Else
                    Return ofd.FileName
                End If
            End If
            Return ""
        End Function

        ''' <summary>
        ''' Shows a file picker to save a file
        ''' </summary>
        ''' <returns>The path of the new file</returns>
        ''' <remarks></remarks>
        Public Function ShowSaveFileDialog() As String Implements IUserInterfaceInterop.ShowSaveFileDialog
            Dim sfd As New SaveFileDialog() With {
                .SupportMultiDottedExtensions = True
            }

            If sfd.ShowDialog() = DialogResult.OK Then
                Return sfd.FileName
            End If
            Return ""
        End Function

        ''' <summary>
        ''' Shows a folder picker
        ''' </summary>
        ''' <param name="Description">The description to show in the folder picker</param>
        ''' <param name="ShowNewFolderButton">Determines whether to show a &quot;New folder&quot; button in the dialog</param>
        ''' <returns>The selected path in the folder picker</returns>
        ''' <remarks></remarks>
        Public Function ShowFolderBrowserDialog(Description As String, ShowNewFolderButton As Boolean) As String Implements IUserInterfaceInterop.ShowFolderBrowserDialog
            Dim selectedPath As String = ""
            ' The FBD will not show in multi-threaded apartment threads, therefore making end-users think tasks that call this function
            ' will never complete. So we create a separate, single-threaded apartment, thread and we wait for it to finish instead.
            ' That DOES work
            Dim thread As New Thread(Sub()
                                         Dim fbd As New FolderBrowserDialog() With {
                                                       .RootFolder = Environment.SpecialFolder.MyComputer,
                                                       .ShowNewFolderButton = ShowNewFolderButton,
                                                       .Description = Description
                                                   }

                                         If fbd.ShowDialog() = DialogResult.OK Then
                                             selectedPath = fbd.SelectedPath
                                         End If
                                     End Sub)
            thread.SetApartmentState(Threading.ApartmentState.STA)
            thread.Start()
            thread.Join()
            Return selectedPath
        End Function

        ''' <summary>
        ''' Shows a message box
        ''' </summary>
        ''' <param name="Message">The message to display</param>
        ''' <param name="Caption">The title of the message</param>
        ''' <remarks></remarks>
        Public Sub ShowMessage(Message As String, Optional Caption As String = "") Implements IUserInterfaceInterop.ShowMessage
            MessageBox.Show(Message, Caption, MessageBoxButtons.OK)
        End Sub

        ''' <summary>
        ''' The REG application is not found on the system
        ''' </summary>
        Const DTERR_RegNotFound As Integer = 2
        ''' <summary>
        ''' A required argument object for REG operations is either null or empty
        ''' </summary>
        Const DTERR_RegItemObjectNull As Integer = 1

        ''' <summary>
        ''' Starts an external process
        ''' </summary>
        ''' <param name="FileName">The file to run</param>
        ''' <param name="Arguments">The arguments to pass to the file</param>
        ''' <param name="WorkingDirectory">The directory the program should run on</param>
        ''' <param name="HideWindow">Whether to hide a window, if created by the program</param>
        ''' <param name="Inconditional">Whether to consider the exit code of the process</param>
        ''' <returns>The exit code of the process if Inconditional is set to False, 0 otherwise.</returns>
        ''' <remarks>
        ''' If a working directory is not specified, this function will use the directory the program specified in FileName is located on
        ''' as the working directory. Please consider changing this to a different directory in your Preparation Task
        ''' if you experience path issues on the external program
        ''' </remarks>
        Public Function RunProcess(FileName As String, Optional Arguments As String = "", Optional WorkingDirectory As String = "", Optional HideWindow As Boolean = False, Optional Inconditional As Boolean = False) As Integer Implements IProcessRunner.RunProcess
            DynaLog.LogMessage("Running external process...")
            DynaLog.LogMessage(String.Format("{0} {1}", FileName, Arguments))
            DynaLog.LogMessage("- Working Directory: " & If(WorkingDirectory <> "", WorkingDirectory, "get from process directory"))
            DynaLog.LogMessage("- Attempt to hide windows the process creates? " & If(HideWindow, "Yes", "No"))
            DynaLog.LogMessage("- Consider process exit code? " & If(Inconditional, "No", "Yes"))

            Dim result As Integer = 0

            Try
                Dim Proc As New Process() With {
                    .StartInfo = New ProcessStartInfo() With {
                        .FileName = FileName,
                        .Arguments = Arguments,
                        .WorkingDirectory = If(WorkingDirectory <> "", WorkingDirectory, Path.GetDirectoryName(FileName)),
                        .CreateNoWindow = HideWindow,
                        .WindowStyle = If(HideWindow, ProcessWindowStyle.Hidden, ProcessWindowStyle.Normal)
                    }
                }

                Proc.Start()
                Proc.WaitForExit()
                result = Proc.ExitCode
            Catch ex As Exception
                DynaLog.LogMessage("Something happened with the process. Error Message: " & ex.Message)

                result = ex.HResult
            End Try

            If Inconditional Then result = PROC_SUCCESS

            DynaLog.LogMessage("Exit code: " & Hex(result))
            Return result
        End Function

        ''' <summary>
        ''' Runs a REG process
        ''' </summary>
        ''' <param name="CommandLine">The command line arguments to pass to the REG program</param>
        ''' <returns>The exit code of REG process</returns>
        ''' <remarks></remarks>
        Public Function RunRegProcess(CommandLine As String) As Integer Implements IRegistryRunner.RunRegProcess
            DynaLog.LogMessage("Running REG process...")
            DynaLog.LogMessage("- Command-line Arguments: " & CommandLine)
            DynaLog.LogMessage("Checking presence of REG program...")
            If Not File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "system32", "reg.exe")) Then
                DynaLog.LogMessage("REG is not found. Aborting this procedure!")
                Return DTERR_RegNotFound
            End If
            DynaLog.LogMessage("REG found. Running...")
            Return RunProcess(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "system32", "reg.exe"),
                              CommandLine,
                              HideWindow:=Debugger.IsAttached)
        End Function

        ''' <summary>
        ''' Gets an appropriate representation of registry value types for REG commands
        ''' </summary>
        ''' <param name="ValueType">The registry value type</param>
        ''' <returns>The representation for REG commands</returns>
        ''' <remarks></remarks>
        Public Function GetRegValueTypeFromEnum(ValueType As RegistryItem.ValueType) As String Implements IRegistryRunner.GetRegValueTypeFromEnum
            Select Case ValueType
                Case RegistryItem.ValueType.RegNone
                    Return "REG_NONE"
                Case RegistryItem.ValueType.RegSz
                    Return "REG_SZ"
                Case RegistryItem.ValueType.RegExpandSz
                    Return "REG_EXPAND_SZ"
                Case RegistryItem.ValueType.RegMultiSz
                    Return "REG_MULTI_SZ"
                Case RegistryItem.ValueType.RegBinary
                    Return "REG_BINARY"
                Case RegistryItem.ValueType.RegDword
                    Return "REG_DWORD"
                Case RegistryItem.ValueType.RegQword
                    Return "REG_QWORD"
            End Select
            Return ""
        End Function

        ''' <summary>
        ''' Adds a registry item to the system
        ''' </summary>
        ''' <param name="RegItem">The new registry item</param>
        ''' <returns>The exit code of the underlying REG process call</returns>
        ''' <remarks></remarks>
        Public Function AddRegistryItem(RegItem As RegistryItem) As Integer Implements IRegistryRunner.AddRegistryItem
            If RegItem Is Nothing Then Return DTERR_RegItemObjectNull

            DynaLog.LogMessage("Adding registry item " & RegItem.ToString())

            Return RunRegProcess(String.Format("add {0} {1} /t {2} /d {3} /f",
                                               RegItem.RegistryKeyLocation,
                                               If(String.IsNullOrEmpty(RegItem.RegistryValueName),
                                                  "/ve",
                                                  String.Format("/v {0}",
                                                                RegItem.RegistryValueName)),
                                               GetRegValueTypeFromEnum(RegItem.RegistryValueType),
                                               RegItem.RegistryValueData))
        End Function

        ''' <summary>
        ''' Removes a registry item from the system
        ''' </summary>
        ''' <param name="RegPath">The absolute path to the item (key or value)</param>
        ''' <param name="DeletionArgs">Deletion arguments to pass to REG</param>
        ''' <returns>The exit code of the underlying REG process call</returns>
        ''' <remarks></remarks>
        Public Function RemoveRegistryItem(RegPath As String, DeletionArgs As String) As Integer Implements IRegistryRunner.RemoveRegistryItem
            If String.IsNullOrEmpty(RegPath) Then Return DTERR_RegItemObjectNull
            If String.IsNullOrEmpty(DeletionArgs) Then Return DTERR_RegItemObjectNull

            DynaLog.LogMessage(String.Format("Removing registry item {0} with provided REG argument {1}", RegPath, DeletionArgs))

            Return RunRegProcess(String.Format("delete {0} {1}",
                                               ControlChars.Quote & RegPath & ControlChars.Quote,
                                               DeletionArgs))
        End Function

        ''' <summary>
        ''' Loads a registry hive to the system
        ''' </summary>
        ''' <param name="RegHivePath">The path of the registry hive</param>
        ''' <param name="RegMountPath">The path to mount the registry hive to</param>
        ''' <returns>The exit code of the underlying REG process call</returns>
        ''' <remarks></remarks>
        Public Function LoadRegistryHive(RegHivePath As String, RegMountPath As String) As Integer Implements IRegistryRunner.LoadRegistryHive
            If String.IsNullOrEmpty(RegHivePath) Then Return DTERR_RegItemObjectNull
            If String.IsNullOrEmpty(RegMountPath) Then Return DTERR_RegItemObjectNull

            DynaLog.LogMessage(String.Format("Loading registry hive {0} to path {1}", RegHivePath, RegMountPath))

            Return RunRegProcess(String.Format("load {0} {1}",
                                               RegMountPath,
                                               ControlChars.Quote & RegHivePath & ControlChars.Quote))
        End Function

        ''' <summary>
        ''' Unloads a registry hive from the system
        ''' </summary>
        ''' <param name="RegMountPath">The path of the mounted hive to unload</param>
        ''' <returns>The exit code of the underlying REG process call</returns>
        ''' <remarks></remarks>
        Public Function UnloadRegistryHive(RegMountPath As String) As Integer Implements IRegistryRunner.UnloadRegistryHive
            If String.IsNullOrEmpty(RegMountPath) Then Return DTERR_RegItemObjectNull

            DynaLog.LogMessage(String.Format("Unloading registry hive {0}", RegMountPath))

            Return RunRegProcess(String.Format("unload {0}",
                                               RegMountPath))
        End Function

        ''' <summary>
        ''' Determines whether a given folder path is a root path
        ''' </summary>
        ''' <param name="FolderPath">The path to a folder</param>
        ''' <returns>Whether the folder is a root path</returns>
        Private Function IsRootPath(FolderPath As String) As Boolean
            Return Path.GetPathRoot(FolderPath) = FolderPath
        End Function

        ''' <summary>
        ''' Removes the contents of a directory, and any subdirectories within the directory, automatically
        ''' and, then, removes the directory
        ''' </summary>
        ''' <param name="DirectoryToRemove">The directory to remove</param>
        ''' <returns>Whether removal succeeded</returns>
        Public Function RemoveRecursive(DirectoryToRemove As String) As Boolean Implements IFileProcessor.RemoveRecursive
            If IsRootPath(DirectoryToRemove) Then Return False
            If Directory.Exists(DirectoryToRemove) Then
                Try
                    Directory.Delete(DirectoryToRemove, True)
                Catch ex As Exception
                    Return RunProcess(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "system32", "cmd.exe"),
                                      "/c del " & Quote & DirectoryToRemove & Quote & " /F /S /Q", HideWindow:=True, Inconditional:=True) = PROC_SUCCESS
                End Try
            End If
            Return True
        End Function

        ''' <summary>
        ''' Copies the contents of a directory, and any subdirectories within the directory,
        ''' to a given destination.
        ''' </summary>
        ''' <param name="SourceDirectory">The directory to copy</param>
        ''' <param name="DestinationDirectory">The destination of the copied files</param>
        ''' <returns>Whether the copy succeeded</returns>
        Public Function CopyRecursive(SourceDirectory As String, DestinationDirectory As String) As Boolean Implements IFileProcessor.CopyRecursive
            ' We make sure the directory exists, if it doesn't exist, we stop.
            If Not Directory.Exists(SourceDirectory) Then Return False

            ' If the destination folder does not exist, then we try creating it. If we couldn't,
            ' we simply give up.
            If Not Directory.Exists(DestinationDirectory) Then
                Try
                    Directory.CreateDirectory(DestinationDirectory)
                Catch ex As Exception
                    Return False
                End Try
            End If

            Try
                ' Now, we create all the directories of the source folder to the destination
                Dim dirsInSource As String() = Directory.GetDirectories(SourceDirectory, "*", SearchOption.AllDirectories)
                For Each dirInSource In dirsInSource
                    Dim sourcePath As String = dirInSource.Substring(SourceDirectory.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                    Dim destinationPath As String = Path.Combine(DestinationDirectory, sourcePath)

                    If Not Directory.Exists(destinationPath) Then
                        Directory.CreateDirectory(destinationPath)
                    End If
                Next

                ' Next, we copy all the files in the source directory to the destination
                For Each FileToCopy In Directory.GetFiles(SourceDirectory, "*", SearchOption.AllDirectories)
                    Dim sourcePath As String = FileToCopy.Substring(SourceDirectory.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                    Dim destinationPath As String = Path.Combine(DestinationDirectory, sourcePath)

                    File.Copy(FileToCopy, destinationPath, True)
                Next
            Catch ex As Exception
                Return False
            End Try

            Return True
        End Function

        ''' <summary>
        ''' Gets the Security Identifier, or SID, of a user given its name.
        ''' </summary>
        ''' <param name="UserName">The user name to get the SID of</param>
        ''' <returns>The SID of the user</returns>
        Public Function GetUserSid(UserName As String) As String Implements IWmiUserProcessor.GetUserSid
            Dim UserSidCollection As ManagementObjectCollection = GetResultsFromManagementQuery("SELECT SID FROM Win32_UserAccount WHERE LocalAccount = True AND Name LIKE " & Quote & UserName & Quote)
            If UserSidCollection IsNot Nothing Then
                Return GetObjectValue(UserSidCollection(0), "SID")
            End If
            Return ""
        End Function
    End Class

End Namespace