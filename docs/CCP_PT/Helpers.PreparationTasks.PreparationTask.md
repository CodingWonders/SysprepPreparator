# Helpers.PreparationTasks.PreparationTask

Namespace: `Helpers.PreparationTasks`

Summary: The base class for Preparation Tasks (PTs).

Remarks: To integrate a Preparation Task into this program, create a class in this namespace and inherit this base class. More information can be found in the documentation.

## Implements

- `IUserInterfaceInterop`
- `IProcessRunner`
- `IRegistryRunner`
- `IFileProcessor`
- `IWmiUserProcessor`

## Constructors

This class is `MustInherit` (abstract in VB). No public constructors are defined.

## Methods

- `RunPreparationTask`
  - Summary: Runs a preparation task.
  - Returns: A `PreparationTaskStatus` value indicating the final status of the task: `Succeeded`, `Failed`, or `Skipped`.
  - Remarks: This must not be called from this parent class, but from classes that inherit this. Classes that inherit this base class must override `RunPreparationTask` and return an appropriate `PreparationTaskStatus` enum value.

- `ReportSubProcessStatus`
  - Summary: Reports a subprocess status change with a given status message.
  - Parameters:
    - `Status` (String): The status message to report.

- `ShowOpenFileDialog`
  - Summary: Shows a file picker to open a file.
  - Parameters:
    - `MultiSelect` (Boolean): Whether to allow file picker to select multiple files.
  - Returns: The path, or paths, of the chosen file.
  - Implements: `IUserInterfaceInterop`

- `ShowSaveFileDialog`
  - Summary: Shows a file picker to save a file.
  - Returns: The path of the new file.
  - Implements: `IUserInterfaceInterop`

- `ShowFolderBrowserDialog`
  - Summary: Shows a folder picker.
  - Parameters:
    - `Description` (String): The description to show in the folder picker.
    - `ShowNewFolderButton` (Boolean): Determines whether to show a "New folder" button in the dialog.
  - Returns: The selected path in the folder picker.
  - Remarks: The FBD will not show in multi-threaded apartment threads, therefore making end-users think tasks that call this function will never complete. So a separate STA thread is used and waited on.
  - Implements: `IUserInterfaceInterop`

- `ShowMessage`
  - Summary: Shows a message box.
  - Parameters:
    - `Message` (String): The message to display.
    - `Caption` (String): The title of the message.
  - Implements: `IUserInterfaceInterop`

- `RunProcess`
  - Summary: Starts an external process.
  - Parameters:
    - `FileName` (String): The file to run.
    - `Arguments` (String): The arguments to pass to the file.
    - `WorkingDirectory` (String): The directory the program should run on.
    - `HideWindow` (Boolean): Whether to hide a window, if created by the program.
    - `Inconditional` (Boolean): Whether to consider the exit code of the process.
  - Returns: The exit code of the process if `Inconditional` is set to False, `0` otherwise.
  - Remarks: If a working directory is not specified, this function will use the directory the program specified in `FileName` is located on as the working directory. Consider changing this in your Preparation Task if you experience path issues.
  - Implements: `IProcessRunner`

- `RunRegProcess`
  - Summary: Runs a REG process.
  - Parameters:
    - `CommandLine` (String): The command line arguments to pass to the REG program.
  - Returns: The exit code of the REG process or a constant error code when `reg.exe` is not found.
  - Implements: `IRegistryRunner`

- `GetRegValueTypeFromEnum`
  - Summary: Gets an appropriate representation of registry value types for REG commands.
  - Parameters:
    - `ValueType` (RegistryItem.ValueType): The registry value type.
  - Returns: The representation for REG commands.
  - Implements: `IRegistryRunner`

- `AddRegistryItem`
  - Summary: Adds a registry item to the system.
  - Parameters:
    - `RegItem` (RegistryItem): The new registry item.
  - Returns: The exit code of the underlying REG process call or an error constant when `RegItem` is null.
  - Implements: `IRegistryRunner`

- `RemoveRegistryItem`
  - Summary: Removes a registry item from the system.
  - Parameters:
    - `RegPath` (String): The absolute path to the item (key or value).
    - `DeletionArgs` (String): Deletion arguments to pass to REG.
  - Returns: The exit code of the underlying REG process call or an error constant when an argument is null or empty.
  - Implements: `IRegistryRunner`

- `LoadRegistryHive`
  - Summary: Loads a registry hive to the system.
  - Parameters:
    - `RegHivePath` (String): The path of the registry hive.
    - `RegMountPath` (String): The path to mount the registry hive to.
  - Returns: The exit code of the underlying REG process call or an error constant when an argument is null or empty.
  - Implements: `IRegistryRunner`

- `UnloadRegistryHive`
  - Summary: Unloads a registry hive from the system.
  - Parameters:
    - `RegMountPath` (String): The path of the mounted hive to unload.
  - Returns: The exit code of the underlying REG process call or an error constant when an argument is null or empty.
  - Implements: `IRegistryRunner`

- `RemoveRecursive`
  - Summary: Removes the contents of a directory, and any subdirectories within the directory, automatically and then removes the directory.
  - Parameters:
    - `DirectoryToRemove` (String): The directory to remove.
  - Returns: Whether removal succeeded. The method will refuse to remove root paths and, on error deleting via .NET APIs, will fallback to using `cmd.exe /c del` and return success based on that command's result.
  - Implements: `IFileProcessor`

- `CopyRecursive`
  - Summary: Copies the contents of a directory, and any subdirectories within the directory, to a given destination.
  - Parameters:
    - `SourceDirectory` (String): The directory to copy.
    - `DestinationDirectory` (String): The destination of the copied files.
  - Returns: Whether the copy succeeded.
  - Implements: `IFileProcessor`

- `GetUserSid`
  - Summary: Gets the Security Identifier (SID) of a user given its name.
  - Parameters:
    - `UserName` (String): The user name to get the SID of.
  - Returns: The SID of the user.
  - Implements: `IWmiUserProcessor`

- `CreateWorkingDirForPT`
  - Summary: Creates the working directory for a specific Preparation Task.
  - Parameters:
    - `PTWorkDir` (String): The name of the working directory for the Preparation Task.
  - Returns: Whether the directory creation operation succeeded for both the creation of the base working directory (if non-existent) and the creation of the Preparation Task working directory (if non-existent). If both folders already exist, True is returned.
  - Remarks: The method attempts to create a base work directory under the system drive and the PT-specific subdirectory. It logs progress to `DynaLog` and returns False on failure to create either folder.

- `PTWorkDirExists`
  - Summary: Detects whether a working directory of a Preparation Task exists in the file system.
  - Parameters:
    - `WorkDirName` (String): The name of the working directory
  - Returns: Whether the full working directory path exists.

## Enums

- `PreparationTaskStatus`
  - Summary: Enum returned by `RunPreparationTask`.
  - Values:
    - `Succeeded` - The preparation task completed successfully.
    - `Failed` - The preparation task failed.
    - `Skipped` - The preparation task did not run (skipped).

## Constants and Fields

- `PROC_SUCCESS`
  - Summary: Constant for external processes that run successfully.
  - Type: `Integer` (value `0`)

- `IsInTestMode`
  - Summary: Determines whether the Sysprep Preparation Tool is in test mode.
  - Remarks: Useful when prototyping Preparation Tasks so you don't run tasks that are typically run on reference computers.
  - Type: `Boolean` (True if program was started with `/test` argument)

- `IsInAutoMode`
  - Summary: Determines whether the Sysprep Preparation Tool is in Automatic mode.
  - Type: `Boolean` (True if program was started with `/auto` argument)

- `SubProcessReporter`
  - Summary: An event sender for subprocess status changes.
  - Type: `Action(Of String)`

- `BaseWorkDir`
  - Summary: The base working directory for all Preparation Tasks.
  - Type: `String` (default: `%SYSTEMDRIVE%\CWS_SYSPRP`)

- `PTWorkDir`
  - Summary: The working directory for a specific Preparation Task (overridable per PT).
  - Type: `String` (empty by default) - inheritors should override this property to provide a per-task working directory name.

- `DTERR_RegNotFound`
  - Summary: Error code returned when `reg.exe` is not found on the system.
  - Type: `Integer` (value `2`)

- `DTERR_RegItemObjectNull`
  - Summary: Error code returned when a required registry item or argument is null or empty.
  - Type: `Integer` (value `1`)

## Remarks

This documentation summarizes the public surface of `PreparationTask` based on the current source file. Implementation details and example usage can be found in the project source.

Working directory guidance:
- `BaseWorkDir` is created under the system drive and is intended to host per-preparation-task working folders.
- Preparation Tasks that need to persist temporary files should override the `PTWorkDir` property to return their working directory name and call `CreateWorkingDirForPT(PTWorkDir)` before writing files.
- Other Preparation Tasks may read from another PT's work folder but should not write to it; each PT should write only to its own working directory.
