<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class OneSymp
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(OneSymp))
        Me.Text1 = New System.Windows.Forms.TextBox()
        Me.Text2 = New System.Windows.Forms.TextBox()
        Me.Lst1 = New System.Windows.Forms.ListBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Command1 = New System.Windows.Forms.Button()
        Me.MustUse = New System.Windows.Forms.Button()
        Me.SelectRemedy = New System.Windows.Forms.Button()
        Me.Lst2 = New System.Windows.Forms.ListBox()
        Me.Lst3 = New System.Windows.Forms.ListBox()
        Me.Lst4 = New System.Windows.Forms.ListBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Weight = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Text1
        '
        Me.Text1.Location = New System.Drawing.Point(10, 58)
        Me.Text1.Margin = New System.Windows.Forms.Padding(2)
        Me.Text1.Name = "Text1"
        Me.Text1.Size = New System.Drawing.Size(321, 20)
        Me.Text1.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.Text1, "Displays first highlighted symptom in ""Symptoms"" list.")
        '
        'Text2
        '
        Me.Text2.Location = New System.Drawing.Point(334, 33)
        Me.Text2.Margin = New System.Windows.Forms.Padding(2)
        Me.Text2.Multiline = True
        Me.Text2.Name = "Text2"
        Me.Text2.ReadOnly = True
        Me.Text2.Size = New System.Drawing.Size(321, 43)
        Me.Text2.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.Text2, "This box displays the cumulative symptom from your selected phrases.")
        '
        'Lst1
        '
        Me.Lst1.BackColor = System.Drawing.SystemColors.Window
        Me.Lst1.FormattingEnabled = True
        Me.Lst1.HorizontalScrollbar = True
        Me.Lst1.Location = New System.Drawing.Point(10, 81)
        Me.Lst1.Margin = New System.Windows.Forms.Padding(2)
        Me.Lst1.Name = "Lst1"
        Me.Lst1.Size = New System.Drawing.Size(321, 316)
        Me.Lst1.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.Lst1, "Select one or more items from this list to build your symptom.")
        '
        'Command1
        '
        Me.Command1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Command1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command1.Location = New System.Drawing.Point(166, 418)
        Me.Command1.Margin = New System.Windows.Forms.Padding(2)
        Me.Command1.Name = "Command1"
        Me.Command1.Size = New System.Drawing.Size(74, 27)
        Me.Command1.TabIndex = 6
        Me.Command1.Text = "Done"
        Me.ToolTip1.SetToolTip(Me.Command1, "Pops down all ""Repertory-Style symptoms"" lists.")
        Me.Command1.UseVisualStyleBackColor = True
        '
        'MustUse
        '
        Me.MustUse.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.MustUse.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MustUse.Location = New System.Drawing.Point(88, 418)
        Me.MustUse.Margin = New System.Windows.Forms.Padding(2)
        Me.MustUse.Name = "MustUse"
        Me.MustUse.Size = New System.Drawing.Size(74, 27)
        Me.MustUse.TabIndex = 4
        Me.MustUse.Text = "Must Use"
        Me.ToolTip1.SetToolTip(Me.MustUse, "Selects last highlighted symptom in ""Symptoms"" box for ""Must Use"".")
        Me.MustUse.UseVisualStyleBackColor = True
        '
        'SelectRemedy
        '
        Me.SelectRemedy.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.SelectRemedy.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SelectRemedy.Location = New System.Drawing.Point(10, 418)
        Me.SelectRemedy.Margin = New System.Windows.Forms.Padding(2)
        Me.SelectRemedy.Name = "SelectRemedy"
        Me.SelectRemedy.Size = New System.Drawing.Size(74, 27)
        Me.SelectRemedy.TabIndex = 3
        Me.SelectRemedy.Text = "Select"
        Me.ToolTip1.SetToolTip(Me.SelectRemedy, "Selects all highlighted symptoms in ""Symptoms"" list.")
        Me.SelectRemedy.UseVisualStyleBackColor = True
        '
        'Lst2
        '
        Me.Lst2.FormattingEnabled = True
        Me.Lst2.HorizontalScrollbar = True
        Me.Lst2.Location = New System.Drawing.Point(334, 81)
        Me.Lst2.Margin = New System.Windows.Forms.Padding(2)
        Me.Lst2.Name = "Lst2"
        Me.Lst2.Size = New System.Drawing.Size(321, 316)
        Me.Lst2.TabIndex = 13
        Me.ToolTip1.SetToolTip(Me.Lst2, "Select one or more items from this list to build your symptom.")
        '
        'Lst3
        '
        Me.Lst3.FormattingEnabled = True
        Me.Lst3.HorizontalScrollbar = True
        Me.Lst3.Location = New System.Drawing.Point(659, 81)
        Me.Lst3.Margin = New System.Windows.Forms.Padding(2)
        Me.Lst3.Name = "Lst3"
        Me.Lst3.Size = New System.Drawing.Size(321, 173)
        Me.Lst3.TabIndex = 14
        Me.ToolTip1.SetToolTip(Me.Lst3, "Select one or more items from this list to build your symptom.")
        '
        'Lst4
        '
        Me.Lst4.FormattingEnabled = True
        Me.Lst4.HorizontalScrollbar = True
        Me.Lst4.Location = New System.Drawing.Point(659, 263)
        Me.Lst4.Margin = New System.Windows.Forms.Padding(2)
        Me.Lst4.Name = "Lst4"
        Me.Lst4.Size = New System.Drawing.Size(321, 134)
        Me.Lst4.TabIndex = 15
        Me.ToolTip1.SetToolTip(Me.Lst4, "Select one or more items from this list to build your symptom.")
        '
        'Button1
        '
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(479, 4)
        Me.Button1.Margin = New System.Windows.Forms.Padding(2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(95, 27)
        Me.Button1.TabIndex = 18
        Me.Button1.Text = "Clear Phrase"
        Me.ToolTip1.SetToolTip(Me.Button1, "Removes phrase and selected items in list(s).")
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Weight
        '
        Me.Weight.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Weight.FormattingEnabled = True
        Me.Weight.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6"})
        Me.Weight.Location = New System.Drawing.Point(244, 423)
        Me.Weight.Margin = New System.Windows.Forms.Padding(2)
        Me.Weight.Name = "Weight"
        Me.Weight.Size = New System.Drawing.Size(92, 21)
        Me.Weight.TabIndex = 20
        Me.ToolTip1.SetToolTip(Me.Weight, "Set to desired weighting factor, then ""Select"" or ""Must Use"" button; or highlight" &
        " symptom to change weight in ""SELECTED SYMPTOMS"" or ""MUST USE"", then set to desi" &
        "red weighting factor")
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(241, 399)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(47, 13)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Weight"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 33)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(131, 13)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "Enter search word(s) here:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(334, 11)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(132, 13)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "Composit symptom phrase:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.Color.Cornsilk
        Me.Label4.Location = New System.Drawing.Point(170, 33)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(94, 13)
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "Denotes active list"
        '
        'Button2
        '
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(906, 33)
        Me.Button2.Margin = New System.Windows.Forms.Padding(2)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(74, 27)
        Me.Button2.TabIndex = 21
        Me.Button2.Text = "Help"
        Me.ToolTip1.SetToolTip(Me.Button2, "Pops down all ""Repertory-Style symptoms"" lists.")
        Me.Button2.UseVisualStyleBackColor = True
        '
        'OneSymp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(989, 455)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Weight)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Lst4)
        Me.Controls.Add(Me.Lst3)
        Me.Controls.Add(Me.Lst2)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Command1)
        Me.Controls.Add(Me.MustUse)
        Me.Controls.Add(Me.SelectRemedy)
        Me.Controls.Add(Me.Lst1)
        Me.Controls.Add(Me.Text2)
        Me.Controls.Add(Me.Text1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "OneSymp"
        Me.Text = "Repertory-format Symptoms"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Text1 As TextBox
    Friend WithEvents Text2 As TextBox
    Friend WithEvents Lst1 As ListBox
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents Command1 As Button
    Friend WithEvents MustUse As Button
    Friend WithEvents SelectRemedy As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents Lst2 As ListBox
    Friend WithEvents Lst3 As ListBox
    Friend WithEvents Lst4 As ListBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents Label4 As Label
    Friend WithEvents Weight As ComboBox
    Friend WithEvents Button2 As Button
End Class
