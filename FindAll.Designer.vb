<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FindAll
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FindAll))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Lst1 = New System.Windows.Forms.ListBox()
        Me.SelectRemedy = New System.Windows.Forms.Button()
        Me.MustUse = New System.Windows.Forms.Button()
        Me.Command3 = New System.Windows.Forms.Button()
        Me.Weight = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(10, 11)
        Me.TextBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBox1.Size = New System.Drawing.Size(346, 65)
        Me.TextBox1.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.TextBox1, "Displays first highlighted symptom in ""Symptoms"" list.")
        '
        'Lst1
        '
        Me.Lst1.FormattingEnabled = True
        Me.Lst1.HorizontalScrollbar = True
        Me.Lst1.Location = New System.Drawing.Point(12, 122)
        Me.Lst1.Margin = New System.Windows.Forms.Padding(2)
        Me.Lst1.Name = "Lst1"
        Me.Lst1.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
        Me.Lst1.Size = New System.Drawing.Size(344, 355)
        Me.Lst1.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.Lst1, "Select one or more items from this list to build your case.")
        '
        'SelectRemedy
        '
        Me.SelectRemedy.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.SelectRemedy.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SelectRemedy.Location = New System.Drawing.Point(12, 504)
        Me.SelectRemedy.Margin = New System.Windows.Forms.Padding(2)
        Me.SelectRemedy.Name = "SelectRemedy"
        Me.SelectRemedy.Size = New System.Drawing.Size(64, 27)
        Me.SelectRemedy.TabIndex = 2
        Me.SelectRemedy.Text = "Select"
        Me.ToolTip1.SetToolTip(Me.SelectRemedy, "Selects all highlighted symptoms in ""Symptoms"" list.")
        Me.SelectRemedy.UseVisualStyleBackColor = True
        '
        'MustUse
        '
        Me.MustUse.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.MustUse.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MustUse.Location = New System.Drawing.Point(81, 504)
        Me.MustUse.Margin = New System.Windows.Forms.Padding(2)
        Me.MustUse.Name = "MustUse"
        Me.MustUse.Size = New System.Drawing.Size(64, 27)
        Me.MustUse.TabIndex = 3
        Me.MustUse.Text = "Must Use"
        Me.ToolTip1.SetToolTip(Me.MustUse, "Selects last highlighted symptom in ""Symptoms"" box for ""Must Use"".")
        Me.MustUse.UseVisualStyleBackColor = True
        '
        'Command3
        '
        Me.Command3.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Command3.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command3.Location = New System.Drawing.Point(150, 504)
        Me.Command3.Margin = New System.Windows.Forms.Padding(2)
        Me.Command3.Name = "Command3"
        Me.Command3.Size = New System.Drawing.Size(64, 27)
        Me.Command3.TabIndex = 4
        Me.Command3.Text = "Done"
        Me.ToolTip1.SetToolTip(Me.Command3, "Pops down the ""Find-All"" dialog.")
        Me.Command3.UseVisualStyleBackColor = True
        '
        'Weight
        '
        Me.Weight.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Weight.FormattingEnabled = True
        Me.Weight.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6"})
        Me.Weight.Location = New System.Drawing.Point(229, 508)
        Me.Weight.Margin = New System.Windows.Forms.Padding(2)
        Me.Weight.Name = "Weight"
        Me.Weight.Size = New System.Drawing.Size(92, 21)
        Me.Weight.TabIndex = 15
        Me.ToolTip1.SetToolTip(Me.Weight, "Set to desired weighting factor, then ""Select"" or ""Must Use"" button; or highlight" &
        " symptom to change weight in ""SELECTED SYMPTOMS"" or ""MUST USE"", then set to desi" &
        "red weighting factor")
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 105)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(105, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Pg/[Sym], Quest, Wt"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(226, 481)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(47, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Weight"
        '
        'Button1
        '
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(292, 87)
        Me.Button1.Margin = New System.Windows.Forms.Padding(2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(64, 27)
        Me.Button1.TabIndex = 16
        Me.Button1.Text = "Help"
        Me.ToolTip1.SetToolTip(Me.Button1, "Pops down the ""Find-All"" dialog.")
        Me.Button1.UseVisualStyleBackColor = True
        '
        'FindAll
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(367, 545)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Weight)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Command3)
        Me.Controls.Add(Me.MustUse)
        Me.Controls.Add(Me.SelectRemedy)
        Me.Controls.Add(Me.Lst1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "FindAll"
        Me.Text = "Find All"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Lst1 As ListBox
    Friend WithEvents SelectRemedy As Button
    Friend WithEvents MustUse As Button
    Friend WithEvents Command3 As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents Weight As ComboBox
    Friend WithEvents Button1 As Button
End Class
