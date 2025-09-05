Namespace Helpers

    ''' <summary>
    ''' This interface provides interoperability with the user interface
    ''' </summary>
    ''' <remarks></remarks>
    Public Interface IUserInterfaceInterop

        ''' <summary>
        ''' Shows a file picker to open a file
        ''' </summary>
        ''' <param name="MultiSelect">Whether to allow file picker to select multiple files</param>
        ''' <returns>The path, or paths, of the chosen file</returns>
        ''' <remarks></remarks>
        Function ShowOpenFileDialog(Optional MultiSelect As Boolean = False) As Object

        ''' <summary>
        ''' Shows a file picker to save a file
        ''' </summary>
        ''' <returns>The path of the new file</returns>
        ''' <remarks></remarks>
        Function ShowSaveFileDialog() As String

        ''' <summary>
        ''' Shows a folder picker
        ''' </summary>
        ''' <param name="Description">The description to show in the folder picker</param>
        ''' <returns>The selected path in the folder picker</returns>
        ''' <remarks></remarks>
        Function ShowFolderBrowserDialog(Description As String) As String

        ''' <summary>
        ''' Shows a message box
        ''' </summary>
        ''' <param name="Message">The message to display</param>
        ''' <param name="Caption">The title of the message</param>
        ''' <remarks></remarks>
        Sub ShowMessage(Message As String, Optional Caption As String = "")

    End Interface

End Namespace
