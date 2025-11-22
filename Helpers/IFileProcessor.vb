Namespace Helpers

    Public Interface IFileProcessor

        ''' <summary>
        ''' Removes the contents of a directory, and any subdirectories within the directory, automatically
        ''' and, then, removes the directory
        ''' </summary>
        ''' <param name="DirectoryToRemove">The directory to remove</param>
        ''' <returns>Whether removal succeeded</returns>
        Function RemoveRecursive(DirectoryToRemove As String) As Boolean

        ''' <summary>
        ''' Copies the contents of a directory, and any subdirectories within the directory,
        ''' to a given destination.
        ''' </summary>
        ''' <param name="SourceDirectory">The directory to copy</param>
        ''' <param name="DestinationDirectory">The destination of the copied files</param>
        ''' <returns>Whether the copy succeeded</returns>
        Function CopyRecursive(SourceDirectory As String, DestinationDirectory As String) As Boolean

    End Interface

End Namespace
