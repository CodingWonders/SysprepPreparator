Namespace Classes

    Public Class WindowsServerRole

        Private ReadOnly OSVER_MAXIMUM As Version = New Version(65535, 0, 0, 0)

        ''' <summary>
        ''' The name of the feature as it appears on DISM
        ''' </summary>
        ''' <returns></returns>
        Public Property FeatureName As String

        ''' <summary>
        ''' The name of the feature to display to the user
        ''' </summary>
        ''' <returns></returns>
        Public Property DisplayName As String

        ''' <summary>
        ''' The minimum version in which Sysprep supports this feature being enabled
        ''' </summary>
        ''' <returns></returns>
        Public Property BaselineVersion As Version

        ''' <summary>
        ''' The maximum version in which Sysprep supports this feature being enabled
        ''' </summary>
        ''' <returns></returns>
        Public Property MaximumVersion As Version

        ''' <summary>
        ''' A factor, or series of factors, that make the feature being enabled unsupported by Sysprep unless not met
        ''' </summary>
        ''' <returns></returns>
        Public Property CompatibilityCaveat As String

        ''' <summary>
        ''' Initializes a role object containing names and a baseline version
        ''' </summary>
        ''' <param name="featName">The feature name</param>
        ''' <param name="dispName">The feature display name</param>
        ''' <param name="baseline">The minimum version for this role to be supported</param>
        Public Sub New(featName As String, dispName As String, baseline As Version)
            FeatureName = featName
            DisplayName = dispName
            MaximumVersion = OSVER_MAXIMUM
            BaselineVersion = baseline
            CompatibilityCaveat = GetValueFromLanguageData("Common.Common_None")
        End Sub

        ''' <summary>
        ''' Initializes a role object containing names, a baseline version, and a caveat
        ''' </summary>
        ''' <param name="featName">The feature name</param>
        ''' <param name="dispName">The feature display name</param>
        ''' <param name="baseline">The minimum version for this role to be supported</param>
        ''' <param name="caveat">A caveat or series of caveats</param>
        Public Sub New(featName As String, dispName As String, baseline As Version, caveat As String)
            FeatureName = featName
            DisplayName = dispName
            MaximumVersion = OSVER_MAXIMUM
            BaselineVersion = baseline
            CompatibilityCaveat = caveat
        End Sub

        ''' <summary>
        ''' Initializes a role object containing names, a baseline version, and a maximum version
        ''' </summary>
        ''' <param name="featName">The feature name</param>
        ''' <param name="dispName">The feature display name</param>
        ''' <param name="baseline">The minimum version for this role to be supported</param>
        ''' <param name="maximum">The maximum version for this role to be supported</param>
        Public Sub New(featName As String, dispName As String, baseline As Version, maximum As Version)
            FeatureName = featName
            DisplayName = dispName
            BaselineVersion = baseline
            MaximumVersion = maximum
            CompatibilityCaveat = GetValueFromLanguageData("Common.Common_None")
        End Sub

        ''' <summary>
        ''' Initializes a role object containing names, a baseline version, a maximum version, and a caveat
        ''' </summary>
        ''' <param name="featName">The feature name</param>
        ''' <param name="dispName">The feature display name</param>
        ''' <param name="baseline">The minimum version for this role to be supported</param>
        ''' <param name="maximum">The maximum version for this role to be supported</param>
        ''' <param name="caveat">A caveat or series of caveats</param>
        Public Sub New(featName As String, dispName As String, baseline As Version, maximum As Version, caveat As String)
            FeatureName = featName
            DisplayName = dispName
            BaselineVersion = baseline
            MaximumVersion = maximum
            CompatibilityCaveat = caveat
        End Sub

    End Class

End Namespace