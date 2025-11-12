Namespace Helpers

    Public Interface IWmiUserProcessor

        ''' <summary>
        ''' Gets the Security Identifier, or SID, of a user given its name.
        ''' </summary>
        ''' <param name="UserName">The user name to get the SID of</param>
        ''' <returns>The SID of the user</returns>
        Function GetUserSid(UserName As String) As String

    End Interface

End Namespace
