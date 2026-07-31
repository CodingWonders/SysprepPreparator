<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SysvolBDEDecryptDialog
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SysvolBDEDecryptDialog))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.pbEncrypted = New System.Windows.Forms.ProgressBar()
        Me.lblDeviceID = New System.Windows.Forms.Label()
        Me.lblPersistentVolumeID = New System.Windows.Forms.Label()
        Me.lblConversionStatus = New System.Windows.Forms.Label()
        Me.lblPercentEncrypted = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoEllipsis = True
        Me.Label1.Location = New System.Drawing.Point(13, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(759, 64)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = resources.GetString("Label1.Text")
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.TableLayoutPanel1)
        Me.GroupBox1.Location = New System.Drawing.Point(13, 91)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(759, 178)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Encryption Details"
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 21.11554!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 78.88446!))
        Me.TableLayoutPanel1.Controls.Add(Me.Label2, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label3, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label4, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label5, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.pbEncrypted, 0, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.lblDeviceID, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.lblPersistentVolumeID, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.lblConversionStatus, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.lblPercentEncrypted, 1, 3)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(3, 19)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 5
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(753, 156)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label2.Location = New System.Drawing.Point(3, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(153, 31)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Device ID:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label3.Location = New System.Drawing.Point(3, 31)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(153, 31)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Persistent Volume ID:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label4.Location = New System.Drawing.Point(3, 62)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(153, 31)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Conversion Status:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label5.Location = New System.Drawing.Point(3, 93)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(153, 31)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "% Encrypted:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'pbEncrypted
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.pbEncrypted, 2)
        Me.pbEncrypted.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pbEncrypted.Location = New System.Drawing.Point(3, 127)
        Me.pbEncrypted.Name = "pbEncrypted"
        Me.pbEncrypted.Size = New System.Drawing.Size(747, 26)
        Me.pbEncrypted.TabIndex = 1
        '
        'lblDeviceID
        '
        Me.lblDeviceID.AutoSize = True
        Me.lblDeviceID.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblDeviceID.Location = New System.Drawing.Point(162, 0)
        Me.lblDeviceID.Name = "lblDeviceID"
        Me.lblDeviceID.Size = New System.Drawing.Size(588, 31)
        Me.lblDeviceID.TabIndex = 0
        '
        'lblPersistentVolumeID
        '
        Me.lblPersistentVolumeID.AutoSize = True
        Me.lblPersistentVolumeID.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblPersistentVolumeID.Location = New System.Drawing.Point(162, 31)
        Me.lblPersistentVolumeID.Name = "lblPersistentVolumeID"
        Me.lblPersistentVolumeID.Size = New System.Drawing.Size(588, 31)
        Me.lblPersistentVolumeID.TabIndex = 0
        '
        'lblConversionStatus
        '
        Me.lblConversionStatus.AutoSize = True
        Me.lblConversionStatus.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblConversionStatus.Location = New System.Drawing.Point(162, 62)
        Me.lblConversionStatus.Name = "lblConversionStatus"
        Me.lblConversionStatus.Size = New System.Drawing.Size(588, 31)
        Me.lblConversionStatus.TabIndex = 0
        '
        'lblPercentEncrypted
        '
        Me.lblPercentEncrypted.AutoSize = True
        Me.lblPercentEncrypted.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblPercentEncrypted.Location = New System.Drawing.Point(162, 93)
        Me.lblPercentEncrypted.Name = "lblPercentEncrypted"
        Me.lblPercentEncrypted.Size = New System.Drawing.Size(588, 31)
        Me.lblPercentEncrypted.TabIndex = 0
        '
        'SysvolBDEDecryptDialog
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(784, 281)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SysvolBDEDecryptDialog"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Decrypting volume..."
        Me.GroupBox1.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents pbEncrypted As ProgressBar
    Friend WithEvents lblDeviceID As Label
    Friend WithEvents lblPersistentVolumeID As Label
    Friend WithEvents lblConversionStatus As Label
    Friend WithEvents lblPercentEncrypted As Label
End Class
