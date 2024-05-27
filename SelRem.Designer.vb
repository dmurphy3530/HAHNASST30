<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SelRem
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SelRem))
        Me.List1 = New System.Windows.Forms.ListBox()
        Me.Command1 = New System.Windows.Forms.Button()
        Me.Command2 = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'List1
        '
        Me.List1.FormattingEnabled = True
        Me.List1.HorizontalScrollbar = True
        Me.List1.Location = New System.Drawing.Point(10, 11)
        Me.List1.Margin = New System.Windows.Forms.Padding(2)
        Me.List1.Name = "List1"
        Me.List1.Size = New System.Drawing.Size(645, 329)
        Me.List1.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.List1, "Double-click on a remedy or click on it and click the ""Select"" button to display " &
        "its symptoms.")
        '
        'Command1
        '
        Me.Command1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Command1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command1.Location = New System.Drawing.Point(169, 348)
        Me.Command1.Margin = New System.Windows.Forms.Padding(2)
        Me.Command1.Name = "Command1"
        Me.Command1.Size = New System.Drawing.Size(82, 28)
        Me.Command1.TabIndex = 26
        Me.Command1.Text = "Select"
        Me.ToolTip1.SetToolTip(Me.Command1, "Click this button to display a list of symptoms for the selected remedy.")
        Me.Command1.UseVisualStyleBackColor = True
        '
        'Command2
        '
        Me.Command2.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Command2.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command2.Location = New System.Drawing.Point(411, 348)
        Me.Command2.Margin = New System.Windows.Forms.Padding(2)
        Me.Command2.Name = "Command2"
        Me.Command2.Size = New System.Drawing.Size(82, 28)
        Me.Command2.TabIndex = 27
        Me.Command2.Text = "Done"
        Me.ToolTip1.SetToolTip(Me.Command2, "This button pops-down the ""Select Remedy"" dialog.")
        Me.Command2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(570, 348)
        Me.Button1.Margin = New System.Windows.Forms.Padding(2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(82, 28)
        Me.Button1.TabIndex = 28
        Me.Button1.Text = "Help"
        Me.ToolTip1.SetToolTip(Me.Button1, "This button pops-down the ""Select Remedy"" dialog.")
        Me.Button1.UseVisualStyleBackColor = True
        '
        'SelRem
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(663, 386)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Command2)
        Me.Controls.Add(Me.Command1)
        Me.Controls.Add(Me.List1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "SelRem"
        Me.Text = "Select Remedy to Display"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents List1 As ListBox
    Friend WithEvents Command1 As Button
    Friend WithEvents Command2 As Button
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents Button1 As Button
End Class
