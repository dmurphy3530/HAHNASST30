<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PrescRem
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PrescRem))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Command1 = New System.Windows.Forms.Button()
        Me.Command4 = New System.Windows.Forms.Button()
        Me.Command2 = New System.Windows.Forms.Button()
        Me.Command3 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(12, 28)
        Me.ListBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(667, 212)
        Me.ListBox1.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.ListBox1, "Displays a list of remedies for your symptom(s).")
        '
        'Command1
        '
        Me.Command1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Command1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command1.Location = New System.Drawing.Point(146, 244)
        Me.Command1.Margin = New System.Windows.Forms.Padding(2)
        Me.Command1.Name = "Command1"
        Me.Command1.Size = New System.Drawing.Size(94, 41)
        Me.Command1.TabIndex = 2
        Me.Command1.Text = "Show Remedy"
        Me.ToolTip1.SetToolTip(Me.Command1, "Displays a symptom list for the highlighted remedy.")
        Me.Command1.UseVisualStyleBackColor = True
        '
        'Command4
        '
        Me.Command4.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Command4.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command4.Location = New System.Drawing.Point(244, 243)
        Me.Command4.Margin = New System.Windows.Forms.Padding(2)
        Me.Command4.Name = "Command4"
        Me.Command4.Size = New System.Drawing.Size(94, 42)
        Me.Command4.TabIndex = 3
        Me.Command4.Text = "Export Presc."
        Me.ToolTip1.SetToolTip(Me.Command4, "Writes the remedy list to a text-format file.")
        Me.Command4.UseVisualStyleBackColor = True
        '
        'Command2
        '
        Me.Command2.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Command2.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command2.Location = New System.Drawing.Point(344, 243)
        Me.Command2.Margin = New System.Windows.Forms.Padding(2)
        Me.Command2.Name = "Command2"
        Me.Command2.Size = New System.Drawing.Size(94, 42)
        Me.Command2.TabIndex = 4
        Me.Command2.Text = "Print Presc."
        Me.ToolTip1.SetToolTip(Me.Command2, "Prints the remedy list.")
        Me.Command2.UseVisualStyleBackColor = True
        '
        'Command3
        '
        Me.Command3.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Command3.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command3.Location = New System.Drawing.Point(442, 243)
        Me.Command3.Margin = New System.Windows.Forms.Padding(2)
        Me.Command3.Name = "Command3"
        Me.Command3.Size = New System.Drawing.Size(94, 42)
        Me.Command3.TabIndex = 5
        Me.Command3.Text = "Done"
        Me.ToolTip1.SetToolTip(Me.Command3, "Pops down the ""Prescribed Remedies"" dialog.")
        Me.Command3.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(10, 11)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(330, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Indication     Remedy   Abrev., Latin name, English name"
        '
        'Button1
        '
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(585, 244)
        Me.Button1.Margin = New System.Windows.Forms.Padding(2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(94, 42)
        Me.Button1.TabIndex = 6
        Me.Button1.Text = "Help"
        Me.ToolTip1.SetToolTip(Me.Button1, "Pops down the ""Prescribed Remedies"" dialog.")
        Me.Button1.UseVisualStyleBackColor = True
        '
        'PrescRem
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(687, 288)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Command3)
        Me.Controls.Add(Me.Command2)
        Me.Controls.Add(Me.Command4)
        Me.Controls.Add(Me.Command1)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "PrescRem"
        Me.Text = "Prescribed Remedies"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents Label1 As Label
    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents Command1 As Button
    Friend WithEvents Command4 As Button
    Friend WithEvents Command2 As Button
    Friend WithEvents Command3 As Button
    Friend WithEvents Button1 As Button
End Class
