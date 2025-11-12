Namespace Helpers

    ''' <summary>
    ''' This interface defines methods to run external processes
    ''' </summary>
    ''' <remarks></remarks>
    Public Interface IProcessRunner

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
        Function RunProcess(FileName As String, Optional Arguments As String = "", Optional WorkingDirectory As String = "", Optional HideWindow As Boolean = False, Optional Inconditional As Boolean = False) As Integer

    End Interface

End Namespace