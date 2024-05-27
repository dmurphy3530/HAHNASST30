<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SetPref
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SetPref))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Option2 = New System.Windows.Forms.CheckBox()
        Me.Option1 = New System.Windows.Forms.CheckBox()
        Me.ColorDialog1 = New System.Windows.Forms.ColorDialog()
        Me.FontDialog1 = New System.Windows.Forms.FontDialog()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Option4 = New System.Windows.Forms.CheckBox()
        Me.Option3 = New System.Windows.Forms.CheckBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Option5 = New System.Windows.Forms.CheckBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.Text1 = New System.Windows.Forms.TextBox()
        Me.Check1 = New System.Windows.Forms.CheckBox()
        Me.Check3 = New System.Windows.Forms.CheckBox()
        Me.Check2 = New System.Windows.Forms.CheckBox()
        Me.ApplyButton = New System.Windows.Forms.Button()
        Me.Command1 = New System.Windows.Forms.Button()
        Me.Command2 = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Text2 = New System.Windows.Forms.TextBox()
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Option2)
        Me.GroupBox1.Controls.Add(Me.Option1)
        Me.GroupBox1.Location = New System.Drawing.Point(10, 11)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Size = New System.Drawing.Size(256, 57)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'Option2
        '
        Me.Option2.AutoSize = True
        Me.Option2.Location = New System.Drawing.Point(15, 32)
        Me.Option2.Margin = New System.Windows.Forms.Padding(2)
        Me.Option2.Name = "Option2"
        Me.Option2.Size = New System.Drawing.Size(84, 17)
        Me.Option2.TabIndex = 1
        Me.Option2.Text = "Screen Font"
        Me.ToolTip1.SetToolTip(Me.Option2, "Use this option to set font name and type for screens.")
        Me.Option2.UseVisualStyleBackColor = True
        '
        'Option1
        '
        Me.Option1.AutoSize = True
        Me.Option1.Location = New System.Drawing.Point(15, 11)
        Me.Option1.Margin = New System.Windows.Forms.Padding(2)
        Me.Option1.Name = "Option1"
        Me.Option1.Size = New System.Drawing.Size(80, 17)
        Me.Option1.TabIndex = 0
        Me.Option1.Text = "Printer Font"
        Me.ToolTip1.SetToolTip(Me.Option1, "Use this option to set font name and type for printing.")
        Me.Option1.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Option4)
        Me.GroupBox2.Controls.Add(Me.Option3)
        Me.GroupBox2.Location = New System.Drawing.Point(10, 72)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Size = New System.Drawing.Size(256, 57)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        '
        'Option4
        '
        Me.Option4.AutoSize = True
        Me.Option4.Location = New System.Drawing.Point(15, 32)
        Me.Option4.Margin = New System.Windows.Forms.Padding(2)
        Me.Option4.Name = "Option4"
        Me.Option4.Size = New System.Drawing.Size(131, 17)
        Me.Option4.TabIndex = 3
        Me.Option4.Text = "Text Foreground Color"
        Me.ToolTip1.SetToolTip(Me.Option4, "Use this option to set the text color for screens.")
        Me.Option4.UseVisualStyleBackColor = True
        '
        'Option3
        '
        Me.Option3.AutoSize = True
        Me.Option3.Location = New System.Drawing.Point(15, 11)
        Me.Option3.Margin = New System.Windows.Forms.Padding(2)
        Me.Option3.Name = "Option3"
        Me.Option3.Size = New System.Drawing.Size(135, 17)
        Me.Option3.TabIndex = 2
        Me.Option3.Text = "Text Background Color"
        Me.ToolTip1.SetToolTip(Me.Option3, "Use this option to set the color behind the text for screens.")
        Me.Option3.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Option5)
        Me.GroupBox3.Location = New System.Drawing.Point(10, 134)
        Me.GroupBox3.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox3.Size = New System.Drawing.Size(256, 43)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        '
        'Option5
        '
        Me.Option5.AutoSize = True
        Me.Option5.Location = New System.Drawing.Point(15, 18)
        Me.Option5.Margin = New System.Windows.Forms.Padding(2)
        Me.Option5.Name = "Option5"
        Me.Option5.Size = New System.Drawing.Size(131, 17)
        Me.Option5.TabIndex = 4
        Me.Option5.Text = "Restore Initial Settings"
        Me.ToolTip1.SetToolTip(Me.Option5, "Use this option to restore fonts and colors to original settings.")
        Me.Option5.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Text1)
        Me.GroupBox4.Location = New System.Drawing.Point(10, 182)
        Me.GroupBox4.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox4.Size = New System.Drawing.Size(256, 65)
        Me.GroupBox4.TabIndex = 3
        Me.GroupBox4.TabStop = False
        '
        'Text1
        '
        Me.Text1.BackColor = System.Drawing.SystemColors.Window
        Me.Text1.Location = New System.Drawing.Point(4, 17)
        Me.Text1.Margin = New System.Windows.Forms.Padding(2)
        Me.Text1.Multiline = True
        Me.Text1.Name = "Text1"
        Me.Text1.Size = New System.Drawing.Size(248, 36)
        Me.Text1.TabIndex = 4
        Me.Text1.TabStop = False
        Me.ToolTip1.SetToolTip(Me.Text1, "This box displays a sample of your screen text.")
        '
        'Check1
        '
        Me.Check1.AutoSize = True
        Me.Check1.Location = New System.Drawing.Point(25, 311)
        Me.Check1.Margin = New System.Windows.Forms.Padding(2)
        Me.Check1.Name = "Check1"
        Me.Check1.Size = New System.Drawing.Size(92, 17)
        Me.Check1.TabIndex = 5
        Me.Check1.Text = "Show Toolbar"
        Me.ToolTip1.SetToolTip(Me.Check1, "If this option is checked, a toolbar will be displayed near the top of the main s" &
        "creen.")
        Me.Check1.UseVisualStyleBackColor = True
        '
        'Check3
        '
        Me.Check3.AutoSize = True
        Me.Check3.Location = New System.Drawing.Point(25, 333)
        Me.Check3.Margin = New System.Windows.Forms.Padding(2)
        Me.Check3.Name = "Check3"
        Me.Check3.Size = New System.Drawing.Size(105, 17)
        Me.Check3.TabIndex = 6
        Me.Check3.Text = "Show Status Bar"
        Me.ToolTip1.SetToolTip(Me.Check3, "If this option is checked, a status bar with tips will be displayed at the bottom" &
        " of the main screen.")
        Me.Check3.UseVisualStyleBackColor = True
        '
        'Check2
        '
        Me.Check2.AutoSize = True
        Me.Check2.Location = New System.Drawing.Point(25, 356)
        Me.Check2.Margin = New System.Windows.Forms.Padding(2)
        Me.Check2.Name = "Check2"
        Me.Check2.Size = New System.Drawing.Size(97, 17)
        Me.Check2.TabIndex = 7
        Me.Check2.Text = "Show ToolTips"
        Me.ToolTip1.SetToolTip(Me.Check2, "If this option is checked, ToolTips will be displayed after you move the mouse po" &
        "inter over an object.")
        Me.Check2.UseVisualStyleBackColor = True
        '
        'ApplyButton
        '
        Me.ApplyButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.ApplyButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ApplyButton.Location = New System.Drawing.Point(46, 416)
        Me.ApplyButton.Margin = New System.Windows.Forms.Padding(2)
        Me.ApplyButton.Name = "ApplyButton"
        Me.ApplyButton.Size = New System.Drawing.Size(56, 24)
        Me.ApplyButton.TabIndex = 9
        Me.ApplyButton.Text = "Apply"
        Me.ToolTip1.SetToolTip(Me.ApplyButton, "This button causes your font and color settings to take effect, and pops-down the" &
        " ""Set Preferences"" dialog.")
        Me.ApplyButton.UseVisualStyleBackColor = True
        '
        'Command1
        '
        Me.Command1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Command1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command1.Location = New System.Drawing.Point(166, 416)
        Me.Command1.Margin = New System.Windows.Forms.Padding(2)
        Me.Command1.Name = "Command1"
        Me.Command1.Size = New System.Drawing.Size(56, 24)
        Me.Command1.TabIndex = 10
        Me.Command1.Text = "Help"
        Me.ToolTip1.SetToolTip(Me.Command1, "Display help for Set Preferences.")
        Me.Command1.UseVisualStyleBackColor = True
        '
        'Command2
        '
        Me.Command2.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Command2.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command2.Location = New System.Drawing.Point(106, 416)
        Me.Command2.Margin = New System.Windows.Forms.Padding(2)
        Me.Command2.Name = "Command2"
        Me.Command2.Size = New System.Drawing.Size(56, 24)
        Me.Command2.TabIndex = 11
        Me.Command2.Text = "Cancel"
        Me.ToolTip1.SetToolTip(Me.Command2, "This button pops-down the """"Set Preferences"""" dialog and discards your selected s" &
        "ettings.")
        Me.Command2.UseVisualStyleBackColor = True
        '
        'Text2
        '
        Me.Text2.Location = New System.Drawing.Point(13, 256)
        Me.Text2.Margin = New System.Windows.Forms.Padding(2)
        Me.Text2.Multiline = True
        Me.Text2.Name = "Text2"
        Me.Text2.Size = New System.Drawing.Size(248, 36)
        Me.Text2.TabIndex = 12
        Me.Text2.TabStop = False
        Me.ToolTip1.SetToolTip(Me.Text2, "This box displays a sample of your screen text.")
        '
        'PrintDialog1
        '
        Me.PrintDialog1.UseEXDialog = True
        '
        'SetPref
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(275, 450)
        Me.Controls.Add(Me.Text2)
        Me.Controls.Add(Me.Command2)
        Me.Controls.Add(Me.Command1)
        Me.Controls.Add(Me.ApplyButton)
        Me.Controls.Add(Me.Check2)
        Me.Controls.Add(Me.Check3)
        Me.Controls.Add(Me.Check1)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "SetPref"
        Me.Text = "Set Preferences"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Option2 As CheckBox
    Friend WithEvents Option1 As CheckBox
    Friend WithEvents ColorDialog1 As ColorDialog
    Friend WithEvents FontDialog1 As FontDialog
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents Option4 As CheckBox
    Friend WithEvents Option3 As CheckBox
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents Option5 As CheckBox
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents Text1 As TextBox
    Friend WithEvents Check1 As CheckBox
    Friend WithEvents Check3 As CheckBox
    Friend WithEvents Check2 As CheckBox
    Friend WithEvents ApplyButton As Button
    Friend WithEvents Command1 As Button
    Friend WithEvents Command2 As Button
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents PrintDialog1 As PrintDialog
    Friend WithEvents Text2 As TextBox
End Class
