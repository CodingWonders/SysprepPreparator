<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PTActivityReporterControl
    Inherits System.Windows.Forms.UserControl

    'UserControl reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.pbTaskStatus = New System.Windows.Forms.PictureBox()
        Me.lblTaskName = New System.Windows.Forms.Label()
        CType(Me.pbTaskStatus, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pbTaskStatus
        '
        Me.pbTaskStatus.Dock = System.Windows.Forms.DockStyle.Left
        Me.pbTaskStatus.Location = New System.Drawing.Point(0, 0)
        Me.pbTaskStatus.Name = "pbTaskStatus"
        Me.pbTaskStatus.Size = New System.Drawing.Size(24, 24)
        Me.pbTaskStatus.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.pbTaskStatus.TabIndex = 0
        Me.pbTaskStatus.TabStop = False
        '
        'lblTaskName
        '
        Me.lblTaskName.AutoEllipsis = True
        Me.lblTaskName.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblTaskName.Location = New System.Drawing.Point(24, 0)
        Me.lblTaskName.Name = "lblTaskName"
        Me.lblTaskName.Padding = New System.Windows.Forms.Padding(8, 0, 0, 0)
        Me.lblTaskName.Size = New System.Drawing.Size(744, 24)
        Me.lblTaskName.TabIndex = 1
        Me.lblTaskName.Text = "Task Name"
        Me.lblTaskName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PTActivityReporterControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.Controls.Add(Me.lblTaskName)
        Me.Controls.Add(Me.pbTaskStatus)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "PTActivityReporterControl"
        Me.Size = New System.Drawing.Size(768, 24)
        CType(Me.pbTaskStatus, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Private WithEvents pbTaskStatus As PictureBox
    Private WithEvents lblTaskName As Label
End Class
