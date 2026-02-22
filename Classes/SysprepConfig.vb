Namespace Classes

    ''' <summary>
    ''' This class provides properties for Sysprep launch configurations
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SysprepConfig

        ''' <summary>
        ''' The cleanup action to take
        ''' </summary>
        ''' <remarks>Look at values in the CleanupAction enumeration for more information</remarks>
        Public Property Cleanup As CleanupAction
        ''' <summary>
        ''' Whether to generalize the system for deployment on multiple machines
        ''' </summary>
        ''' <remarks>By default, this will be true</remarks>
        Public Property Generalize As Boolean
        ''' <summary>
        ''' The shutdown action to take once Sysprep finishes
        ''' </summary>
        ''' <remarks>Look at values in the ShutdownMode enumeration for more information</remarks>
        Public Property Shutdown As ShutdownMode
        ''' <summary>
        ''' The path to an answer file, if specified
        ''' </summary>
        ''' <remarks></remarks>
        Public Property AnswerFile As String
        ''' <summary>
        ''' Determines whether Sysprep will process the installation in virtual machine mode (/mode:vm)
        ''' </summary>
        ''' <remarks>
        ''' According to Microsoft documentation, VM mode is useful on installations done in VMs to prepare
        ''' that target installation for capture as a virtual hard disk (VHD) that can be used on other machines
        ''' of the same VM solution (eg: Hyper-V). See more at
        ''' https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/sysprep-command-line-options?view=windows-11#modevm
        ''' </remarks>
        Public Property VMMode As Boolean

        ''' <summary>
        ''' Determines whether Sysprep will copy all registry preferences of the current user for new user profiles.
        ''' </summary>
        ''' <returns></returns>
        Public Property CopyProfile As Boolean

        ''' <summary>
        ''' Initializes a default instance of Sysprep configuration with default settings
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            Cleanup = CleanupAction.EnterSystemOobe
            Generalize = True
            Shutdown = ShutdownMode.Shutdown
            AnswerFile = ""
            VMMode = False
        End Sub

        ''' <summary>
        ''' The action Sysprep should take when processing cleanup items
        ''' </summary>
        ''' <remarks></remarks>
        Enum CleanupAction As Integer
            ''' <summary>
            ''' The system will boot to the OOBE (Out of Box Experience)
            ''' </summary>
            EnterSystemOobe = 0
            ''' <summary>
            ''' The system will boot back to Audit mode
            ''' </summary>
            EnterSystemAudit = 1
        End Enum

        ''' <summary>
        ''' The action Sysprep should take after finishing
        ''' </summary>
        ''' <remarks></remarks>
        Enum ShutdownMode As Integer
            ''' <summary>
            ''' The system should be rebooted
            ''' </summary>
            Reboot = 0
            ''' <summary>
            ''' The system should be shut down
            ''' </summary>
            Shutdown = 1
            ''' <summary>
            ''' The system should stay active. Sysprep should just quit after finishing
            ''' </summary>
            Quit = 2
        End Enum

    End Class

End Namespace
