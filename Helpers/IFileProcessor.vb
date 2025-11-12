Namespace Helpers

    Public Interface IFileProcessor

        ''' <summary>
        ''' Removes the contents of a directory, and any subdirectories within the directory, automatically
        ''' and, then, removes the directory
        ''' </summary>
        ''' <param name="DirectoryToRemove">The directory to remove</param>
        ''' <returns>Whether removal succeeded</returns>
        Function RemoveRecursive(DirectoryToRemove As String) As Boolean

    End Interface

End Namespace
