Namespace Classes

    ''' <summary>
    ''' This class provides properties for status messages
    ''' </summary>
    ''' <remarks></remarks>
    Public Class StatusMessage

        ''' <summary>
        ''' The severity levels of a status message
        ''' </summary>
        ''' <remarks></remarks>
        Enum StatusMessageSeverity
            ''' <summary>
            ''' An information severity
            ''' </summary>
            ''' <remarks></remarks>
            Info
            ''' <summary>
            ''' A warning severity. It may be possible to continue but there may be issues ahead
            ''' </summary>
            ''' <remarks>Like the warning lights in a car</remarks>
            Warning
            ''' <summary>
            ''' A critical severity. The end-user cannot continue if an event with this severity is logged
            ''' </summary>
            ''' <remarks></remarks>
            Critical
            ''' <summary>
            ''' An unknown severity level
            ''' </summary>
            ''' <remarks></remarks>
            Unknown
        End Enum

        ''' <summary>
        ''' The title of the status message
        ''' </summary>
        Public Property StatusTitle As String
        ''' <summary>
        ''' The description of the status message
        ''' </summary>
        Public Property StatusDescription As String
        ''' <summary>
        ''' The possible resolution of the status message
        ''' </summary>
        ''' <remarks>
        ''' This only needs to be filled when a NON-INFO severity is set (eg: warning, error). The possible resolution
        ''' needs to be set accordingly, based on what you tried to do in order to fix this problem. Use experience as
        ''' an advantage.
        ''' </remarks>
        Public Property StatusPossibleResolution As String
        ''' <summary>
        ''' The severity level of the status message
        ''' </summary>
        ''' <remarks>Look at values in the StatusMessageSeverity enumeration for more information on values</remarks>
        Public Property StatusSeverity As StatusMessageSeverity

        ''' <summary>
        ''' Initializes a default instance of a status message
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            StatusTitle = ""
            StatusDescription = ""
            StatusSeverity = StatusMessageSeverity.Unknown
        End Sub

        ''' <summary>
        ''' Initializes an instance of a status message with title, description, and severity parameters.
        ''' No resolution is provided in this constructor.
        ''' </summary>
        ''' <param name="title">The title of the status message</param>
        ''' <param name="description">The description of the status message</param>
        ''' <param name="severity">The severity level of the status message</param>
        ''' <remarks></remarks>
        Public Sub New(title As String, description As String, severity As StatusMessageSeverity)
            StatusTitle = title
            StatusDescription = description
            StatusPossibleResolution = GetValueFromLanguageData("CCPBase.STMSG_NoResolution")
            StatusSeverity = severity
        End Sub

        ''' <summary>
        ''' Initializes an instance of a status message with title, description, resolution, and severity parameters.
        ''' </summary>
        ''' <param name="title">The title of the status message</param>
        ''' <param name="description">The description of the status message</param>
        ''' <param name="resolution">The possible resolution for the status message</param>
        ''' <param name="severity">The severity level of the status message</param>
        ''' <remarks></remarks>
        Public Sub New(title As String, description As String, resolution As String, severity As StatusMessageSeverity)
            StatusTitle = title
            StatusDescription = description
            StatusPossibleResolution = resolution
            StatusSeverity = severity
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format("{0}: {1}. Resolution: {2}; Severity: {3}", StatusTitle, StatusDescription.TrimEnd("."), StatusPossibleResolution, StatusSeverity)
        End Function

        ''' <summary>
        ''' Converts a severity level to a human-readable string
        ''' </summary>
        ''' <param name="Severity">The severity level of the status message</param>
        ''' <returns>The human-readable representation</returns>
        ''' <remarks></remarks>
        Public Function SeverityToString(Severity As StatusMessageSeverity) As String
            Select Case Severity
                Case StatusMessageSeverity.Info
                    Return GetValueFromLanguageData("CCPBase.STMSG_Severity_Info")
                Case StatusMessageSeverity.Warning
                    Return GetValueFromLanguageData("CCPBase.STMSG_Severity_Warning")
                Case StatusMessageSeverity.Critical
                    Return GetValueFromLanguageData("CCPBase.STMSG_Severity_Critical")
                Case Else
                    Return GetValueFromLanguageData("CCPBase.STMSG_Severity_Unknown")
            End Select
        End Function

    End Class

End Namespace
