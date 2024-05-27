<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class GridSearch
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(GridSearch))
        Me.FindFirst = New System.Windows.Forms.Button()
        Me.FindNext = New System.Windows.Forms.Button()
        Me.FindPrev = New System.Windows.Forms.Button()
        Me.FindAllButton = New System.Windows.Forms.Button()
        Me.Command2 = New System.Windows.Forms.Button()
        Me.Command3 = New System.Windows.Forms.Button()
        Me.Clear = New System.Windows.Forms.Button()
        Me.TextA1 = New System.Windows.Forms.TextBox()
        Me.TextB1 = New System.Windows.Forms.TextBox()
        Me.TextC1 = New System.Windows.Forms.TextBox()
        Me.TextD1 = New System.Windows.Forms.TextBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.List1 = New System.Windows.Forms.ListBox()
        Me.TextD2 = New System.Windows.Forms.TextBox()
        Me.TextC2 = New System.Windows.Forms.TextBox()
        Me.TextB2 = New System.Windows.Forms.TextBox()
        Me.TextD3 = New System.Windows.Forms.TextBox()
        Me.TextC3 = New System.Windows.Forms.TextBox()
        Me.TextB3 = New System.Windows.Forms.TextBox()
        Me.TextD4 = New System.Windows.Forms.TextBox()
        Me.TextC4 = New System.Windows.Forms.TextBox()
        Me.TextB4 = New System.Windows.Forms.TextBox()
        Me.TextD5 = New System.Windows.Forms.TextBox()
        Me.TextC5 = New System.Windows.Forms.TextBox()
        Me.TextB5 = New System.Windows.Forms.TextBox()
        Me.TextD6 = New System.Windows.Forms.TextBox()
        Me.TextC6 = New System.Windows.Forms.TextBox()
        Me.TextB6 = New System.Windows.Forms.TextBox()
        Me.TextA6 = New System.Windows.Forms.TextBox()
        Me.TextA5 = New System.Windows.Forms.TextBox()
        Me.TextA4 = New System.Windows.Forms.TextBox()
        Me.TextA3 = New System.Windows.Forms.TextBox()
        Me.TextA2 = New System.Windows.Forms.TextBox()
        Me.List2 = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.HelpProvider1 = New System.Windows.Forms.HelpProvider()
        Me.SuspendLayout()
        '
        'FindFirst
        '
        Me.FindFirst.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.FindFirst.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FindFirst.Location = New System.Drawing.Point(34, 11)
        Me.FindFirst.Margin = New System.Windows.Forms.Padding(2)
        Me.FindFirst.Name = "FindFirst"
        Me.FindFirst.Size = New System.Drawing.Size(69, 19)
        Me.FindFirst.TabIndex = 0
        Me.FindFirst.Text = "Find First"
        Me.ToolTip1.SetToolTip(Me.FindFirst, "This button finds the first occurrence of your words in the Symptom List.")
        Me.FindFirst.UseVisualStyleBackColor = True
        '
        'FindNext
        '
        Me.FindNext.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.FindNext.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FindNext.Location = New System.Drawing.Point(134, 11)
        Me.FindNext.Margin = New System.Windows.Forms.Padding(2)
        Me.FindNext.Name = "FindNext"
        Me.FindNext.Size = New System.Drawing.Size(69, 19)
        Me.FindNext.TabIndex = 1
        Me.FindNext.Text = "Find Next"
        Me.ToolTip1.SetToolTip(Me.FindNext, "This button starts looking at the current location of the Symptom List and finds " &
        "the first occurrence of your words below this point.")
        Me.FindNext.UseVisualStyleBackColor = True
        '
        'FindPrev
        '
        Me.FindPrev.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.FindPrev.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FindPrev.Location = New System.Drawing.Point(235, 11)
        Me.FindPrev.Margin = New System.Windows.Forms.Padding(2)
        Me.FindPrev.Name = "FindPrev"
        Me.FindPrev.Size = New System.Drawing.Size(69, 19)
        Me.FindPrev.TabIndex = 2
        Me.FindPrev.Text = "Find Prev"
        Me.ToolTip1.SetToolTip(Me.FindPrev, ".This button starts looking at the current location of the Symptom List and finds" &
        " the first occurrence of your words above this point")
        Me.FindPrev.UseVisualStyleBackColor = True
        '
        'FindAllButton
        '
        Me.FindAllButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.FindAllButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FindAllButton.Location = New System.Drawing.Point(335, 11)
        Me.FindAllButton.Margin = New System.Windows.Forms.Padding(2)
        Me.FindAllButton.Name = "FindAllButton"
        Me.FindAllButton.Size = New System.Drawing.Size(69, 19)
        Me.FindAllButton.TabIndex = 3
        Me.FindAllButton.Text = "Find All"
        Me.ToolTip1.SetToolTip(Me.FindAllButton, "This button finds all symptoms which contain your words and pops them up in a lis" &
        "t, from which you can select them.")
        Me.FindAllButton.UseVisualStyleBackColor = True
        '
        'Command2
        '
        Me.Command2.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Command2.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command2.Location = New System.Drawing.Point(436, 11)
        Me.Command2.Margin = New System.Windows.Forms.Padding(2)
        Me.Command2.Name = "Command2"
        Me.Command2.Size = New System.Drawing.Size(69, 19)
        Me.Command2.TabIndex = 4
        Me.Command2.Text = "Cancel"
        Me.ToolTip1.SetToolTip(Me.Command2, "This button pops-down the Grid Search dialog.")
        Me.Command2.UseVisualStyleBackColor = True
        '
        'Command3
        '
        Me.Command3.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Command3.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command3.Location = New System.Drawing.Point(665, 10)
        Me.Command3.Margin = New System.Windows.Forms.Padding(2)
        Me.Command3.Name = "Command3"
        Me.Command3.Size = New System.Drawing.Size(69, 19)
        Me.Command3.TabIndex = 5
        Me.Command3.Text = "Help"
        Me.ToolTip1.SetToolTip(Me.Command3, "This button provides help for the Grid Search dialog.")
        Me.Command3.UseVisualStyleBackColor = True
        '
        'Clear
        '
        Me.Clear.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Clear.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Clear.Location = New System.Drawing.Point(665, 33)
        Me.Clear.Margin = New System.Windows.Forms.Padding(2)
        Me.Clear.Name = "Clear"
        Me.Clear.Size = New System.Drawing.Size(69, 19)
        Me.Clear.TabIndex = 6
        Me.Clear.Text = "Clear All"
        Me.ToolTip1.SetToolTip(Me.Clear, "This button clears all your entries.")
        Me.Clear.UseVisualStyleBackColor = True
        '
        'TextA1
        '
        Me.TextA1.Location = New System.Drawing.Point(34, 80)
        Me.TextA1.Margin = New System.Windows.Forms.Padding(2)
        Me.TextA1.Name = "TextA1"
        Me.TextA1.Size = New System.Drawing.Size(155, 20)
        Me.TextA1.TabIndex = 8
        Me.ToolTip1.SetToolTip(Me.TextA1, "Begin typing a search word in one of these boxes, then select it in the list that" &
        " appears below.")
        '
        'TextB1
        '
        Me.TextB1.Location = New System.Drawing.Point(210, 80)
        Me.TextB1.Margin = New System.Windows.Forms.Padding(2)
        Me.TextB1.Name = "TextB1"
        Me.TextB1.Size = New System.Drawing.Size(155, 20)
        Me.TextB1.TabIndex = 10
        Me.ToolTip1.SetToolTip(Me.TextB1, "Begin typing a search word in one of these boxes, or pop-up the list and select a" &
        " word by double-clicking.")
        '
        'TextC1
        '
        Me.TextC1.Location = New System.Drawing.Point(386, 80)
        Me.TextC1.Margin = New System.Windows.Forms.Padding(2)
        Me.TextC1.Name = "TextC1"
        Me.TextC1.Size = New System.Drawing.Size(155, 20)
        Me.TextC1.TabIndex = 12
        Me.ToolTip1.SetToolTip(Me.TextC1, "Begin typing a search word in one of these boxes, or pop-up the list and select a" &
        " word by double-clicking.")
        '
        'TextD1
        '
        Me.TextD1.Location = New System.Drawing.Point(562, 80)
        Me.TextD1.Margin = New System.Windows.Forms.Padding(2)
        Me.TextD1.Name = "TextD1"
        Me.TextD1.Size = New System.Drawing.Size(155, 20)
        Me.TextD1.TabIndex = 14
        Me.ToolTip1.SetToolTip(Me.TextD1, "Begin typing a search word in one of these boxes, or pop-up the list and select a" &
        " word by double-clicking.")
        '
        'List1
        '
        Me.List1.FormattingEnabled = True
        Me.List1.HorizontalScrollbar = True
        Me.List1.Location = New System.Drawing.Point(34, 271)
        Me.List1.Margin = New System.Windows.Forms.Padding(2)
        Me.List1.Name = "List1"
        Me.List1.Size = New System.Drawing.Size(702, 199)
        Me.List1.TabIndex = 16
        Me.ToolTip1.SetToolTip(Me.List1, "This list contains all the significant words in the Symptom List; you select word" &
        "s from this list for your search.")
        '
        'TextD2
        '
        Me.TextD2.Location = New System.Drawing.Point(562, 104)
        Me.TextD2.Margin = New System.Windows.Forms.Padding(2)
        Me.TextD2.Name = "TextD2"
        Me.TextD2.Size = New System.Drawing.Size(155, 20)
        Me.TextD2.TabIndex = 22
        Me.ToolTip1.SetToolTip(Me.TextD2, "Begin typing a search word in one of these boxes, or pop-up the list and select a" &
        " word by double-clicking.")
        '
        'TextC2
        '
        Me.TextC2.Location = New System.Drawing.Point(386, 104)
        Me.TextC2.Margin = New System.Windows.Forms.Padding(2)
        Me.TextC2.Name = "TextC2"
        Me.TextC2.Size = New System.Drawing.Size(155, 20)
        Me.TextC2.TabIndex = 20
        Me.ToolTip1.SetToolTip(Me.TextC2, "Begin typing a search word in one of these boxes, or pop-up the list and select a" &
        " word by double-clicking.")
        '
        'TextB2
        '
        Me.TextB2.Location = New System.Drawing.Point(210, 104)
        Me.TextB2.Margin = New System.Windows.Forms.Padding(2)
        Me.TextB2.Name = "TextB2"
        Me.TextB2.Size = New System.Drawing.Size(155, 20)
        Me.TextB2.TabIndex = 18
        Me.ToolTip1.SetToolTip(Me.TextB2, "Begin typing a search word in one of these boxes, or pop-up the list and select a" &
        " word by double-clicking.")
        '
        'TextD3
        '
        Me.TextD3.Location = New System.Drawing.Point(562, 127)
        Me.TextD3.Margin = New System.Windows.Forms.Padding(2)
        Me.TextD3.Name = "TextD3"
        Me.TextD3.Size = New System.Drawing.Size(155, 20)
        Me.TextD3.TabIndex = 29
        Me.ToolTip1.SetToolTip(Me.TextD3, "Begin typing a search word in one of these boxes, or pop-up the list and select a" &
        " word by double-clicking.")
        '
        'TextC3
        '
        Me.TextC3.Location = New System.Drawing.Point(386, 127)
        Me.TextC3.Margin = New System.Windows.Forms.Padding(2)
        Me.TextC3.Name = "TextC3"
        Me.TextC3.Size = New System.Drawing.Size(155, 20)
        Me.TextC3.TabIndex = 27
        Me.ToolTip1.SetToolTip(Me.TextC3, "Begin typing a search word in one of these boxes, or pop-up the list and select a" &
        " word by double-clicking.")
        '
        'TextB3
        '
        Me.TextB3.Location = New System.Drawing.Point(210, 127)
        Me.TextB3.Margin = New System.Windows.Forms.Padding(2)
        Me.TextB3.Name = "TextB3"
        Me.TextB3.Size = New System.Drawing.Size(155, 20)
        Me.TextB3.TabIndex = 25
        Me.ToolTip1.SetToolTip(Me.TextB3, "Begin typing a search word in one of these boxes, or pop-up the list and select a" &
        " word by double-clicking.")
        '
        'TextD4
        '
        Me.TextD4.Location = New System.Drawing.Point(562, 150)
        Me.TextD4.Margin = New System.Windows.Forms.Padding(2)
        Me.TextD4.Name = "TextD4"
        Me.TextD4.Size = New System.Drawing.Size(155, 20)
        Me.TextD4.TabIndex = 36
        Me.ToolTip1.SetToolTip(Me.TextD4, "Begin typing a search word in one of these boxes, or pop-up the list and select a" &
        " word by double-clicking.")
        '
        'TextC4
        '
        Me.TextC4.Location = New System.Drawing.Point(386, 150)
        Me.TextC4.Margin = New System.Windows.Forms.Padding(2)
        Me.TextC4.Name = "TextC4"
        Me.TextC4.Size = New System.Drawing.Size(155, 20)
        Me.TextC4.TabIndex = 34
        Me.ToolTip1.SetToolTip(Me.TextC4, "Begin typing a search word in one of these boxes, or pop-up the list and select a" &
        " word by double-clicking.")
        '
        'TextB4
        '
        Me.TextB4.Location = New System.Drawing.Point(210, 150)
        Me.TextB4.Margin = New System.Windows.Forms.Padding(2)
        Me.TextB4.Name = "TextB4"
        Me.TextB4.Size = New System.Drawing.Size(155, 20)
        Me.TextB4.TabIndex = 32
        Me.ToolTip1.SetToolTip(Me.TextB4, "Begin typing a search word in one of these boxes, or pop-up the list and select a" &
        " word by double-clicking.")
        '
        'TextD5
        '
        Me.TextD5.Location = New System.Drawing.Point(562, 172)
        Me.TextD5.Margin = New System.Windows.Forms.Padding(2)
        Me.TextD5.Name = "TextD5"
        Me.TextD5.Size = New System.Drawing.Size(155, 20)
        Me.TextD5.TabIndex = 43
        Me.ToolTip1.SetToolTip(Me.TextD5, "Begin typing a search word in one of these boxes, or pop-up the list and select a" &
        " word by double-clicking.")
        '
        'TextC5
        '
        Me.TextC5.Location = New System.Drawing.Point(386, 172)
        Me.TextC5.Margin = New System.Windows.Forms.Padding(2)
        Me.TextC5.Name = "TextC5"
        Me.TextC5.Size = New System.Drawing.Size(155, 20)
        Me.TextC5.TabIndex = 41
        Me.ToolTip1.SetToolTip(Me.TextC5, "Begin typing a search word in one of these boxes, or pop-up the list and select a" &
        " word by double-clicking.")
        '
        'TextB5
        '
        Me.TextB5.Location = New System.Drawing.Point(210, 172)
        Me.TextB5.Margin = New System.Windows.Forms.Padding(2)
        Me.TextB5.Name = "TextB5"
        Me.TextB5.Size = New System.Drawing.Size(155, 20)
        Me.TextB5.TabIndex = 39
        Me.ToolTip1.SetToolTip(Me.TextB5, "Begin typing a search word in one of these boxes, or pop-up the list and select a" &
        " word by double-clicking.")
        '
        'TextD6
        '
        Me.TextD6.Location = New System.Drawing.Point(562, 195)
        Me.TextD6.Margin = New System.Windows.Forms.Padding(2)
        Me.TextD6.Name = "TextD6"
        Me.TextD6.Size = New System.Drawing.Size(155, 20)
        Me.TextD6.TabIndex = 50
        Me.ToolTip1.SetToolTip(Me.TextD6, "Begin typing a search word in one of these boxes, or pop-up the list and select a" &
        " word by double-clicking.")
        '
        'TextC6
        '
        Me.TextC6.Location = New System.Drawing.Point(386, 195)
        Me.TextC6.Margin = New System.Windows.Forms.Padding(2)
        Me.TextC6.Name = "TextC6"
        Me.TextC6.Size = New System.Drawing.Size(155, 20)
        Me.TextC6.TabIndex = 48
        Me.ToolTip1.SetToolTip(Me.TextC6, "Begin typing a search word in one of these boxes, or pop-up the list and select a" &
        " word by double-clicking.")
        '
        'TextB6
        '
        Me.TextB6.Location = New System.Drawing.Point(210, 195)
        Me.TextB6.Margin = New System.Windows.Forms.Padding(2)
        Me.TextB6.Name = "TextB6"
        Me.TextB6.Size = New System.Drawing.Size(155, 20)
        Me.TextB6.TabIndex = 46
        Me.ToolTip1.SetToolTip(Me.TextB6, "Begin typing a search word in one of these boxes, or pop-up the list and select a" &
        " word by double-clicking.")
        '
        'TextA6
        '
        Me.TextA6.Location = New System.Drawing.Point(34, 195)
        Me.TextA6.Margin = New System.Windows.Forms.Padding(2)
        Me.TextA6.Name = "TextA6"
        Me.TextA6.Size = New System.Drawing.Size(155, 20)
        Me.TextA6.TabIndex = 65
        Me.ToolTip1.SetToolTip(Me.TextA6, "Begin typing a search word in one of these boxes, or pop-up the list and select a" &
        " word by double-clicking.")
        '
        'TextA5
        '
        Me.TextA5.Location = New System.Drawing.Point(34, 172)
        Me.TextA5.Margin = New System.Windows.Forms.Padding(2)
        Me.TextA5.Name = "TextA5"
        Me.TextA5.Size = New System.Drawing.Size(155, 20)
        Me.TextA5.TabIndex = 64
        Me.ToolTip1.SetToolTip(Me.TextA5, "Begin typing a search word in one of these boxes, or pop-up the list and select a" &
        " word by double-clicking.")
        '
        'TextA4
        '
        Me.TextA4.Location = New System.Drawing.Point(34, 150)
        Me.TextA4.Margin = New System.Windows.Forms.Padding(2)
        Me.TextA4.Name = "TextA4"
        Me.TextA4.Size = New System.Drawing.Size(155, 20)
        Me.TextA4.TabIndex = 63
        Me.ToolTip1.SetToolTip(Me.TextA4, "Begin typing a search word in one of these boxes, or pop-up the list and select a" &
        " word by double-clicking.")
        '
        'TextA3
        '
        Me.TextA3.Location = New System.Drawing.Point(34, 127)
        Me.TextA3.Margin = New System.Windows.Forms.Padding(2)
        Me.TextA3.Name = "TextA3"
        Me.TextA3.Size = New System.Drawing.Size(155, 20)
        Me.TextA3.TabIndex = 62
        Me.ToolTip1.SetToolTip(Me.TextA3, "Begin typing a search word in one of these boxes, or pop-up the list and select a" &
        " word by double-clicking.")
        '
        'TextA2
        '
        Me.TextA2.Location = New System.Drawing.Point(34, 104)
        Me.TextA2.Margin = New System.Windows.Forms.Padding(2)
        Me.TextA2.Name = "TextA2"
        Me.TextA2.Size = New System.Drawing.Size(155, 20)
        Me.TextA2.TabIndex = 61
        Me.ToolTip1.SetToolTip(Me.TextA2, "Begin typing a search word in one of these boxes, or pop-up the list and select a" &
        " word by double-clicking.")
        '
        'List2
        '
        Me.List2.FormattingEnabled = True
        Me.List2.HorizontalScrollbar = True
        Me.List2.Location = New System.Drawing.Point(34, 271)
        Me.List2.Margin = New System.Windows.Forms.Padding(2)
        Me.List2.Name = "List2"
        Me.List2.Size = New System.Drawing.Size(702, 199)
        Me.List2.Sorted = True
        Me.List2.TabIndex = 66
        Me.ToolTip1.SetToolTip(Me.List2, "This list contains all the significant words in the Symptom List; you select word" &
        "s from this list for your search.")
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(186, 62)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(33, 13)
        Me.Label1.TabIndex = 52
        Me.Label1.Text = "AND"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(358, 64)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(33, 13)
        Me.Label2.TabIndex = 53
        Me.Label2.Text = "AND"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(535, 64)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(33, 13)
        Me.Label3.TabIndex = 54
        Me.Label3.Text = "AND"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(9, 94)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(25, 13)
        Me.Label4.TabIndex = 55
        Me.Label4.Text = "OR"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(9, 117)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(25, 13)
        Me.Label5.TabIndex = 56
        Me.Label5.Text = "OR"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(9, 140)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(25, 13)
        Me.Label6.TabIndex = 57
        Me.Label6.Text = "OR"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(9, 162)
        Me.Label7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(25, 13)
        Me.Label7.TabIndex = 58
        Me.Label7.Text = "OR"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(9, 185)
        Me.Label8.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(25, 13)
        Me.Label8.TabIndex = 59
        Me.Label8.Text = "OR"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'GridSearch
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(746, 479)
        Me.Controls.Add(Me.List2)
        Me.Controls.Add(Me.TextA6)
        Me.Controls.Add(Me.TextA5)
        Me.Controls.Add(Me.TextA4)
        Me.Controls.Add(Me.TextA3)
        Me.Controls.Add(Me.TextA2)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextD6)
        Me.Controls.Add(Me.TextC6)
        Me.Controls.Add(Me.TextB6)
        Me.Controls.Add(Me.TextD5)
        Me.Controls.Add(Me.TextC5)
        Me.Controls.Add(Me.TextB5)
        Me.Controls.Add(Me.TextD4)
        Me.Controls.Add(Me.TextC4)
        Me.Controls.Add(Me.TextB4)
        Me.Controls.Add(Me.TextD3)
        Me.Controls.Add(Me.TextC3)
        Me.Controls.Add(Me.TextB3)
        Me.Controls.Add(Me.TextD2)
        Me.Controls.Add(Me.TextC2)
        Me.Controls.Add(Me.TextB2)
        Me.Controls.Add(Me.List1)
        Me.Controls.Add(Me.TextD1)
        Me.Controls.Add(Me.TextC1)
        Me.Controls.Add(Me.TextB1)
        Me.Controls.Add(Me.TextA1)
        Me.Controls.Add(Me.Clear)
        Me.Controls.Add(Me.Command3)
        Me.Controls.Add(Me.Command2)
        Me.Controls.Add(Me.FindAllButton)
        Me.Controls.Add(Me.FindPrev)
        Me.Controls.Add(Me.FindNext)
        Me.Controls.Add(Me.FindFirst)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "GridSearch"
        Me.Text = "Power Grid Search"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents FindFirst As Button
    Friend WithEvents FindNext As Button
    Friend WithEvents FindPrev As Button
    Friend WithEvents FindAllButton As Button
    Friend WithEvents Command2 As Button
    Friend WithEvents Command3 As Button
    Friend WithEvents Clear As Button
    Friend WithEvents TextA1 As TextBox
    Friend WithEvents TextB1 As TextBox
    Friend WithEvents TextC1 As TextBox
    Friend WithEvents TextD1 As TextBox
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents List1 As ListBox
    Friend WithEvents TextD2 As TextBox
    Friend WithEvents TextC2 As TextBox
    Friend WithEvents TextB2 As TextBox
    Friend WithEvents TextD3 As TextBox
    Friend WithEvents TextC3 As TextBox
    Friend WithEvents TextB3 As TextBox
    Friend WithEvents TextD4 As TextBox
    Friend WithEvents TextC4 As TextBox
    Friend WithEvents TextB4 As TextBox
    Friend WithEvents TextD5 As TextBox
    Friend WithEvents TextC5 As TextBox
    Friend WithEvents TextB5 As TextBox
    Friend WithEvents TextD6 As TextBox
    Friend WithEvents TextC6 As TextBox
    Friend WithEvents TextB6 As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents TextA6 As TextBox
    Friend WithEvents TextA5 As TextBox
    Friend WithEvents TextA4 As TextBox
    Friend WithEvents TextA3 As TextBox
    Friend WithEvents TextA2 As TextBox
    Friend WithEvents HelpProvider1 As HelpProvider
    Friend WithEvents List2 As ListBox
End Class
