<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PrintSelections
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
        Me.PrintSymp = New System.Windows.Forms.CheckBox()
        Me.PrintPresc = New System.Windows.Forms.CheckBox()
        Me.PrintRem = New System.Windows.Forms.CheckBox()
        Me.PrintQuest = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.PrintAll = New System.Windows.Forms.CheckBox()
        Me.CancelPButton = New System.Windows.Forms.Button()
        Me.PrintButton = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.SuspendLayout()
        '
        'PrintSymp
        '
        Me.PrintSymp.AutoSize = True
        Me.PrintSymp.Location = New System.Drawing.Point(13, 45)
        Me.PrintSymp.Name = "PrintSymp"
        Me.PrintSymp.Size = New System.Drawing.Size(88, 17)
        Me.PrintSymp.TabIndex = 0
        Me.PrintSymp.Text = "Symptom List"
        Me.ToolTip1.SetToolTip(Me.PrintSymp, "Include list of selected symptoms in print")
        Me.PrintSymp.UseVisualStyleBackColor = True
        '
        'PrintPresc
        '
        Me.PrintPresc.AutoSize = True
        Me.PrintPresc.Location = New System.Drawing.Point(13, 68)
        Me.PrintPresc.Name = "PrintPresc"
        Me.PrintPresc.Size = New System.Drawing.Size(127, 17)
        Me.PrintPresc.TabIndex = 1
        Me.PrintPresc.Text = "Suggested Remedies"
        Me.ToolTip1.SetToolTip(Me.PrintPresc, "Include list of suggested remedies in print")
        Me.PrintPresc.UseVisualStyleBackColor = True
        '
        'PrintRem
        '
        Me.PrintRem.AutoSize = True
        Me.PrintRem.Location = New System.Drawing.Point(13, 92)
        Me.PrintRem.Name = "PrintRem"
        Me.PrintRem.Size = New System.Drawing.Size(126, 17)
        Me.PrintRem.TabIndex = 2
        Me.PrintRem.Text = "Remedy Descriptions"
        Me.ToolTip1.SetToolTip(Me.PrintRem, "Include description of all suggested remedies in print")
        Me.PrintRem.UseVisualStyleBackColor = True
        '
        'PrintQuest
        '
        Me.PrintQuest.AutoSize = True
        Me.PrintQuest.Location = New System.Drawing.Point(13, 115)
        Me.PrintQuest.Name = "PrintQuest"
        Me.PrintQuest.Size = New System.Drawing.Size(91, 17)
        Me.PrintQuest.TabIndex = 3
        Me.PrintQuest.Text = "Questionnaire"
        Me.ToolTip1.SetToolTip(Me.PrintQuest, "Include questionnaire in print")
        Me.PrintQuest.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(173, 20)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Select Items to Print"
        '
        'PrintAll
        '
        Me.PrintAll.AutoSize = True
        Me.PrintAll.Location = New System.Drawing.Point(13, 162)
        Me.PrintAll.Name = "PrintAll"
        Me.PrintAll.Size = New System.Drawing.Size(107, 17)
        Me.PrintAll.TabIndex = 5
        Me.PrintAll.Text = "Print All Available"
        Me.ToolTip1.SetToolTip(Me.PrintAll, "Print all of the above")
        Me.PrintAll.UseVisualStyleBackColor = True
        '
        'CancelPButton
        '
        Me.CancelPButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CancelPButton.Location = New System.Drawing.Point(97, 190)
        Me.CancelPButton.Name = "CancelPButton"
        Me.CancelPButton.Size = New System.Drawing.Size(75, 23)
        Me.CancelPButton.TabIndex = 6
        Me.CancelPButton.Text = "Cancel"
        Me.ToolTip1.SetToolTip(Me.CancelPButton, "Click to not print")
        Me.CancelPButton.UseVisualStyleBackColor = True
        '
        'PrintButton
        '
        Me.PrintButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PrintButton.Location = New System.Drawing.Point(16, 190)
        Me.PrintButton.Name = "PrintButton"
        Me.PrintButton.Size = New System.Drawing.Size(75, 23)
        Me.PrintButton.TabIndex = 7
        Me.PrintButton.Text = "Print"
        Me.ToolTip1.SetToolTip(Me.PrintButton, "Click to print all selections")
        Me.PrintButton.UseVisualStyleBackColor = True
        '
        'PrintSelections
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(197, 235)
        Me.Controls.Add(Me.PrintButton)
        Me.Controls.Add(Me.CancelPButton)
        Me.Controls.Add(Me.PrintAll)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PrintQuest)
        Me.Controls.Add(Me.PrintRem)
        Me.Controls.Add(Me.PrintPresc)
        Me.Controls.Add(Me.PrintSymp)
        Me.HelpButton = True
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "PrintSelections"
        Me.Text = "PrintSelections"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents PrintSymp As CheckBox
    Friend WithEvents PrintPresc As CheckBox
    Friend WithEvents PrintRem As CheckBox
    Friend WithEvents PrintQuest As CheckBox
    Friend WithEvents Label1 As Label
    Friend WithEvents PrintAll As CheckBox
    Friend WithEvents CancelPButton As Button
    Friend WithEvents PrintButton As Button
    Friend WithEvents ToolTip1 As ToolTip
End Class
