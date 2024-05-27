<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DispRem
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
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Print = New System.Windows.Forms.Button()
        Me.Done = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'Print
        '
        Me.Print.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Print.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Print.Location = New System.Drawing.Point(176, 347)
        Me.Print.Margin = New System.Windows.Forms.Padding(2)
        Me.Print.Name = "Print"
        Me.Print.Size = New System.Drawing.Size(76, 30)
        Me.Print.TabIndex = 1
        Me.Print.Text = "Print"
        Me.ToolTip1.SetToolTip(Me.Print, "This button prints the displayed remedy and symptoms.")
        Me.Print.UseVisualStyleBackColor = True
        '
        'Done
        '
        Me.Done.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Done.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Done.Location = New System.Drawing.Point(466, 347)
        Me.Done.Margin = New System.Windows.Forms.Padding(2)
        Me.Done.Name = "Done"
        Me.Done.Size = New System.Drawing.Size(76, 30)
        Me.Done.TabIndex = 2
        Me.Done.Text = "Done"
        Me.ToolTip1.SetToolTip(Me.Done, "This button pops-down the Display Remedy dialog.")
        Me.Done.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(606, 347)
        Me.Button1.Margin = New System.Windows.Forms.Padding(2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(76, 30)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "Help"
        Me.ToolTip1.SetToolTip(Me.Button1, "This button pops-down the Display Remedy dialog.")
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(12, 12)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBox1.Size = New System.Drawing.Size(670, 330)
        Me.TextBox1.TabIndex = 4
        '
        'DispRem
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(694, 387)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Done)
        Me.Controls.Add(Me.Print)
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "DispRem"
        Me.Text = "Remedy & Associated Symptoms"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents Print As Button
    Friend WithEvents Done As Button
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Button1 As Button
End Class
