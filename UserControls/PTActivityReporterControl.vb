Imports SysprepPreparator.Helpers.PreparationTasks

Public Class PTActivityReporterControl

    Private _taskStatusPicture As Image
    Private _taskName As String

    Public Property TaskName As String
        Get
            Return _taskName
        End Get
        Set(value As String)
            _taskName = value
            lblTaskName.Text = If(_taskName <> "", _taskName, "Task Name")
        End Set
    End Property

    Public WriteOnly Property TaskStatus As PreparationTask.PreparationTaskStatus
        Set(value As PreparationTask.PreparationTaskStatus)
            Select Case value
                Case PreparationTask.PreparationTaskStatus.Succeeded : pbTaskStatus.Image = GetGlyphResource("succeeded")
                Case PreparationTask.PreparationTaskStatus.Failed : pbTaskStatus.Image = GetGlyphResource("failed")
                Case PreparationTask.PreparationTaskStatus.Skipped : pbTaskStatus.Image = GetGlyphResource("skipped")
            End Select
        End Set
    End Property

    Public Sub SetColors()
        BackColor = CurrentTheme.SectionBackgroundColor
        ForeColor = CurrentTheme.ForegroundColor
    End Sub

    Public Sub SetActive()
        lblTaskName.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        lblTaskName.ForeColor = CurrentTheme.ForegroundColor
    End Sub

    Public Sub SetInactive()
        lblTaskName.Font = New Font("Segoe UI", 9)
        lblTaskName.ForeColor = CurrentTheme.DisabledForegroundColor
    End Sub


End Class
