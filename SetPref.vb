Option Explicit On
Public Class SetPref
    Private Sub SetPref_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim msg As String   ' error message string

        Try
            Form1.ErrorFlag = 11001

            ' Set attributes from globals.
            Text1.Font = Form1.TextBox1.Font
            Text1.BackColor = Form1.TextBox1.BackColor
            Text1.ForeColor = Form1.TextBox1.ForeColor

            Option1.Checked = False
            Option2.Checked = False
            Option3.Checked = False
            Option4.Checked = False

            If Form1.ToolStrip1.Visible Then
                Check1.Checked = 1
            Else
                Check1.Checked = 0
            End If

            If Form1.ToolText Then
                Check2.Checked = True
            Else
                Check2.Checked = False
            End If

            If Form1.StatusStrip1.Visible Then
                Check3.Checked = True
            Else
                Check3.Checked = False
            End If

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub SetPref_Closing(sender As Object, e As EventArgs) Handles MyBase.Closing
        Dim msg As String   ' error message string

        Try
            Form1.ErrorFlag = 11002
            SaveSetting("Hahnasst", "Startup",
        "ShowToolbar", Form1.ShowToolbar)

            SaveSetting("Hahnasst", "Startup",
        "ShowStatusbar", Form1.ShowStatusbar)

            SaveSetting("Hahnasst", "Startup",
        "ToolText", Form1.ToolText)

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub Option1_CheckedChanged(sender As Object, e As EventArgs) Handles Option1.CheckedChanged 'Printer font
        Dim msg As String   ' error message string

        Try
            Form1.ErrorFlag = 11003

            If Option1.Checked() Then
                If Form1.FontDialog1.ShowDialog() = vbOK Then ' Display Font common dialog box.

                    '   Set printer attributes.
                    '            Printer.FontName = Form1.CommonDialog1.FontName
                    '            Printer.FontSize = Form1.CommonDialog1.FontSize

                    '   Set printer globals.
                    Form1.PrinterFontName = Form1.FontDialog1.Font
                    Text1.Font = Form1.PrinterFontName
                    Text1.Text = "Printer Sample"
                End If
            Else
                Text1.Text = "" ' Option was de-selected.
            End If

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub Option2_CheckedChanged(sender As Object, e As EventArgs) Handles Option2.CheckedChanged ' Screen font
        Dim msg As String   ' error message string

        Try
            Form1.ErrorFlag = 11004
            If Option2.Checked() Then
                If FontDialog1.ShowDialog() = vbOK Then ' Display Font common dialog box.

                    '   Preview selected settings.
                    Text2.Font = FontDialog1.Font
                    Text2.Text = "Screen Sample"
                End If
            Else
                Text2.Text = "" ' Option was de-selected.
            End If
        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub Option3_CheckedChanged(sender As Object, e As EventArgs) Handles Option3.CheckedChanged ' Screen (textbox) background color
        Dim msg As String   ' error message string

        Try
            Form1.ErrorFlag = 11005
            If Option3.Checked() Then
                If ColorDialog1.ShowDialog() = vbOK Then

                    '   Preview color setting.
                    Text2.BackColor = ColorDialog1.Color
                    Text2.Text = "Screen Sample"
                End If
            End If

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub Option4_CheckedChanged(sender As Object, e As EventArgs) Handles Option4.CheckedChanged ' Screen (textbox) foreground color
        Dim msg As String   ' error message string

        Try
            Form1.ErrorFlag = 11006
            If Option4.Checked() Then

                If ColorDialog1.ShowDialog() = vbOK Then

                    '   Preview color setting.
                    Text2.ForeColor = ColorDialog1.Color
                    Text2.Text = "Screen Sample"
                End If
            End If

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub Option5_CheckedChanged(sender As Object, e As EventArgs) Handles Option5.CheckedChanged ' Restore initial settings
        Dim msg As String   ' dialog message string

        Try
            Form1.ErrorFlag = 11007
            If Option5.Checked() Then
                msg = "Are you sure you want to restore printer and screen attributes to initial program settings?"

                If (MsgBox(msg, vbYesNo + vbQuestion) = 6) Then
                    Text1.Font = Control.DefaultFont
                    Text1.Text = "Printer Sample"
                    Text2.ForeColor = Color.FromName("WindowText")
                    Text2.BackColor = Color.FromName("Window")
                    Text2.Font = Control.DefaultFont
                    Text2.Text = "Screen Sample"
                End If
            End If

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub




    Private Sub ApplyButton_Click(sender As Object, e As EventArgs) Handles ApplyButton.Click
        ApplyButtonClick()
        Close()
    End Sub

    Private Sub ApplyButtonClick()
        Dim msg As String   ' error message string

        Try
            Form1.ErrorFlag = 11011
            If Text1.Text <> "" Then
                ' No need to set attribute globals; printer fonts are read directly from registration database at print time.
                ' Save printer settings in registration database.
                '   Save settings to registration database.
                SaveSetting("Hahnasst", "Startup",
                    "PrinterFontName", Form1.FontDialog1.Font.Name)
                SaveSetting("Hahnasst", "Startup",
                        "PrinterFontSize", Form1.FontDialog1.Font.Size)
            End If
            If Text2.Text <> "" Then
                ' Set attribute globals.
                Form1.ScreenFontName = Text1.Font
                Form1.ScreenFontSize = Text1.Font.Size
                Form1.ScreenBackColor = Text1.BackColor
                Form1.ScreenForeColor = Text1.ForeColor

                '  Set attributes.
                Form1.TextBox1.BackColor = Form1.BackColor
                Form1.Lst1.BackColor = Form1.ScreenBackColor
                Form1.MustBox.BackColor = Form1.ScreenBackColor
                Form1.SelLst.BackColor = Form1.ScreenBackColor

                Form1.TextBox1.ForeColor = Form1.ScreenForeColor
                Form1.Lst1.ForeColor = Form1.ScreenForeColor
                Form1.MustBox.ForeColor = Form1.ScreenForeColor
                Form1.SelLst.ForeColor = Form1.ScreenForeColor

                Form1.TextBox1.Font = Form1.ScreenFontName
                Form1.Lst1.Font = Form1.ScreenFontName
                Form1.MustBox.Font = Form1.ScreenFontName
                Form1.SelLst.Font = Form1.ScreenFontName

                ' Save screen settings in registration database.
                SaveSetting("Hahnasst", "Startup",
            "ScreenFontName", Text1.Font.Name)
                SaveSetting("Hahnasst", "Startup",
            "ScreenFontBold", Text1.Font.Bold)
                SaveSetting("Hahnasst", "Startup",
            "ScreenFontItalic", Text1.Font.Italic)
                SaveSetting("Hahnasst", "Startup",
            "ScreenFontUnderline", Text1.Font.Underline)
                SaveSetting("Hahnasst", "Startup",
            "ScreenFontStrikeout", Text1.Font.Strikeout)
                SaveSetting("Hahnasst", "Startup",
            "ScreenFontSize", Text1.Font.Size)
                SaveSetting("Hahnasst", "Startup",
            "BackColor", Text2.BackColor.ToString)
                SaveSetting("Hahnasst", "Startup",
            "ForeColor", Text2.ForeColor.ToString)
            End If

            ' Handle ToolTip, Toolbar, and Status Bar settings
            If Check1.Checked = True Then   ' Show Toolbar
                Form1.ShowToolbar = True
                Form1.ToolStrip1.Visible = True
            Else
                Form1.ShowToolbar = False
                Form1.ToolStrip1.Visible = False
            End If
            If Check3.Checked = True Then   ' Show Status Bar
                Form1.ShowStatusbar = True
                Form1.StatusStrip1.Visible = True
            Else
                Form1.ShowStatusbar = False
                Form1.StatusStrip1.Visible = False
            End If
            If Check2.Checked = True Then   ' Show ToolTips
                Form1.ToolText = True
                Form1.TurnOnToolTips()
            Else
                Form1.ToolText = False
                Form1.TurnOffToolTips()
            End If
        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
            Buttons:=vbOKOnly + vbCritical,
            Title:="ERROR!")
        End Try
    End Sub


    Private Sub Command1_Click(sender As Object, e As EventArgs) Handles Command1.Click
        Dim msg As String   ' error message string
        Dim myProcess As New Process

        Try
            myProcess.StartInfo.FileName = AppContext.BaseDirectory + Form1.HELP_DIR + "OptionsMenuCommands.htm"
            myProcess.StartInfo.UseShellExecute = True
            myProcess.StartInfo.RedirectStandardOutput = False
            myProcess.Start()
            myProcess.Dispose()
        Catch
            msg = "Help file not found.  Please re-install program to fix this problem."
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs) Handles Command2.Click
        Close()
    End Sub
End Class