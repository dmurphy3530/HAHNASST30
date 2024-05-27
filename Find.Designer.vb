<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Find
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Find))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.List1 = New System.Windows.Forms.ListBox()
        Me.Check1 = New System.Windows.Forms.CheckBox()
        Me.FindFirst = New System.Windows.Forms.Button()
        Me.FindNext = New System.Windows.Forms.Button()
        Me.FindPrev = New System.Windows.Forms.Button()
        Me.FindAllButton = New System.Windows.Forms.Button()
        Me.CancelFindButton = New System.Windows.Forms.Button()
        Me.HelpFindButton = New System.Windows.Forms.Button()
        Me.Clear = New System.Windows.Forms.Button()
        Me.Text1 = New System.Windows.Forms.TextBox()
        Me.HelpProvider1 = New System.Windows.Forms.HelpProvider()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 11)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Find what:"
        '
        'List1
        '
        Me.List1.FormattingEnabled = True
        Me.List1.Location = New System.Drawing.Point(97, 37)
        Me.List1.Margin = New System.Windows.Forms.Padding(2)
        Me.List1.Name = "List1"
        Me.List1.Size = New System.Drawing.Size(214, 342)
        Me.List1.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.List1, "This list contains all the significant words in the Symptom List; you select a wo" &
        "rd of this list for your search.")
        '
        'Check1
        '
        Me.Check1.AutoSize = True
        Me.Check1.Location = New System.Drawing.Point(12, 28)
        Me.Check1.Margin = New System.Windows.Forms.Padding(2)
        Me.Check1.Name = "Check1"
        Me.Check1.Size = New System.Drawing.Size(86, 17)
        Me.Check1.TabIndex = 3
        Me.Check1.Text = "Exact Match"
        Me.ToolTip1.SetToolTip(Me.Check1, "When this is checked, it searches for the whole word you select; otherwise, it al" &
        "so finds words which contain your selected word.")
        Me.Check1.UseVisualStyleBackColor = True
        '
        'FindFirst
        '
        Me.FindFirst.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.FindFirst.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FindFirst.Location = New System.Drawing.Point(331, 11)
        Me.FindFirst.Margin = New System.Windows.Forms.Padding(2)
        Me.FindFirst.Name = "FindFirst"
        Me.FindFirst.Size = New System.Drawing.Size(70, 19)
        Me.FindFirst.TabIndex = 4
        Me.FindFirst.Text = "Find First"
        Me.ToolTip1.SetToolTip(Me.FindFirst, "This button finds the first occurrence of your word in the Symptom List.")
        Me.FindFirst.UseVisualStyleBackColor = True
        '
        'FindNext
        '
        Me.FindNext.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.FindNext.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FindNext.Location = New System.Drawing.Point(331, 34)
        Me.FindNext.Margin = New System.Windows.Forms.Padding(2)
        Me.FindNext.Name = "FindNext"
        Me.FindNext.Size = New System.Drawing.Size(70, 19)
        Me.FindNext.TabIndex = 5
        Me.FindNext.Text = "Find Next"
        Me.ToolTip1.SetToolTip(Me.FindNext, "This button starts looking at the current location of the Symptom List and finds " &
        "the first occurrence of your word below this point.")
        Me.FindNext.UseVisualStyleBackColor = True
        '
        'FindPrev
        '
        Me.FindPrev.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.FindPrev.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FindPrev.Location = New System.Drawing.Point(331, 58)
        Me.FindPrev.Margin = New System.Windows.Forms.Padding(2)
        Me.FindPrev.Name = "FindPrev"
        Me.FindPrev.Size = New System.Drawing.Size(70, 19)
        Me.FindPrev.TabIndex = 6
        Me.FindPrev.Text = "Find Prev"
        Me.ToolTip1.SetToolTip(Me.FindPrev, "This button starts looking at the current location of the Symptom List and finds " &
        "the first occurrence of your word above this point.")
        Me.FindPrev.UseVisualStyleBackColor = True
        '
        'FindAllButton
        '
        Me.FindAllButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.FindAllButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FindAllButton.Location = New System.Drawing.Point(331, 81)
        Me.FindAllButton.Margin = New System.Windows.Forms.Padding(2)
        Me.FindAllButton.Name = "FindAllButton"
        Me.FindAllButton.Size = New System.Drawing.Size(70, 19)
        Me.FindAllButton.TabIndex = 7
        Me.FindAllButton.Text = "Find All"
        Me.ToolTip1.SetToolTip(Me.FindAllButton, "This button finds all symptoms which contain your word and pops them up in a list" &
        ", from which you can select them.")
        Me.FindAllButton.UseVisualStyleBackColor = True
        '
        'CancelFindButton
        '
        Me.CancelFindButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.CancelFindButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CancelFindButton.Location = New System.Drawing.Point(331, 105)
        Me.CancelFindButton.Margin = New System.Windows.Forms.Padding(2)
        Me.CancelFindButton.Name = "CancelFindButton"
        Me.CancelFindButton.Size = New System.Drawing.Size(70, 19)
        Me.CancelFindButton.TabIndex = 8
        Me.CancelFindButton.Text = "Cancel"
        Me.ToolTip1.SetToolTip(Me.CancelFindButton, "This button pops-down the Find dialog.")
        Me.CancelFindButton.UseVisualStyleBackColor = True
        '
        'HelpFindButton
        '
        Me.HelpFindButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.HelpFindButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.HelpFindButton.Location = New System.Drawing.Point(331, 128)
        Me.HelpFindButton.Margin = New System.Windows.Forms.Padding(2)
        Me.HelpFindButton.Name = "HelpFindButton"
        Me.HelpFindButton.Size = New System.Drawing.Size(70, 19)
        Me.HelpFindButton.TabIndex = 9
        Me.HelpFindButton.Text = "Help"
        Me.ToolTip1.SetToolTip(Me.HelpFindButton, "This button provides help for the Find dialog.")
        Me.HelpFindButton.UseVisualStyleBackColor = True
        '
        'Clear
        '
        Me.Clear.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Clear.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Clear.Location = New System.Drawing.Point(12, 56)
        Me.Clear.Margin = New System.Windows.Forms.Padding(2)
        Me.Clear.Name = "Clear"
        Me.Clear.Size = New System.Drawing.Size(70, 19)
        Me.Clear.TabIndex = 10
        Me.Clear.Text = "Clear"
        Me.ToolTip1.SetToolTip(Me.Clear, "This button clears your entry.")
        Me.Clear.UseVisualStyleBackColor = True
        '
        'Text1
        '
        Me.Text1.Location = New System.Drawing.Point(97, 8)
        Me.Text1.Margin = New System.Windows.Forms.Padding(2)
        Me.Text1.Name = "Text1"
        Me.Text1.Size = New System.Drawing.Size(214, 20)
        Me.Text1.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.Text1, "Begin typing your search word in this box, or pop-up the list and select it by do" &
        "uble-clicking.")
        '
        'Find
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(410, 394)
        Me.Controls.Add(Me.Text1)
        Me.Controls.Add(Me.Clear)
        Me.Controls.Add(Me.HelpFindButton)
        Me.Controls.Add(Me.CancelFindButton)
        Me.Controls.Add(Me.FindAllButton)
        Me.Controls.Add(Me.FindPrev)
        Me.Controls.Add(Me.FindNext)
        Me.Controls.Add(Me.FindFirst)
        Me.Controls.Add(Me.Check1)
        Me.Controls.Add(Me.List1)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "Find"
        Me.Text = "Find"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents List1 As ListBox
    Friend WithEvents Check1 As CheckBox
    Friend WithEvents FindFirst As Button
    Friend WithEvents FindNext As Button
    Friend WithEvents FindPrev As Button
    Friend WithEvents FindAllButton As Button
    Friend WithEvents CancelFindButton As Button
    Friend WithEvents HelpFindButton As Button
    Friend WithEvents Clear As Button
    Friend WithEvents Text1 As TextBox
    Friend WithEvents HelpProvider1 As HelpProvider
End Class
