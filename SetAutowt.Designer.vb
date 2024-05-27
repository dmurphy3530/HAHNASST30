<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SetAutowt
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.Command1 = New System.Windows.Forms.Button()
        Me.Command2 = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.SuspendLayout()
        '
        'Command1
        '
        Me.Command1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Command1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command1.Location = New System.Drawing.Point(53, 43)
        Me.Command1.Margin = New System.Windows.Forms.Padding(2)
        Me.Command1.Name = "Command1"
        Me.Command1.Size = New System.Drawing.Size(56, 19)
        Me.Command1.TabIndex = 0
        Me.Command1.Text = "Do It"
        Me.ToolTip1.SetToolTip(Me.Command1, "Apply Holistic weight numbers to all selected symptoms")
        Me.Command1.UseVisualStyleBackColor = True
        '
        'Command2
        '
        Me.Command2.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Command2.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command2.Location = New System.Drawing.Point(158, 43)
        Me.Command2.Margin = New System.Windows.Forms.Padding(2)
        Me.Command2.Name = "Command2"
        Me.Command2.Size = New System.Drawing.Size(56, 19)
        Me.Command2.TabIndex = 1
        Me.Command2.Text = "DONE"
        Me.ToolTip1.SetToolTip(Me.Command2, "Close this form")
        Me.Command2.UseVisualStyleBackColor = True
        '
        'SetAutowt
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(265, 96)
        Me.Controls.Add(Me.Command2)
        Me.Controls.Add(Me.Command1)
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "SetAutowt"
        Me.Text = "Set Up Autoweight"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Command1 As Button
    Friend WithEvents Command2 As Button
    Friend WithEvents ToolTip1 As ToolTip
End Class
