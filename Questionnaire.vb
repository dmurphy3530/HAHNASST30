Option Explicit On
Public Class Questionnaire
    Public Shared Text1(48) As TextBox
    Public Shared Check1(126) As CheckBox
    Public Shared Label(30) As Label

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Questionnaire_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Idx As Integer  ' Check box index
        Dim msg As String   ' Dialog box message

        Try
            Form1.ErrorFlag = 8001
            If Form1.QuestFirstLoad Then
                Label83.Text = "No questionnaire file"
                Label85.Text = "1"  ' Symptom #
                ' Set up pointers to the text and check boxes and labels so they can be indexed easily.
                ' Page 1
                Text1(0) = TextBox1
                Text1(1) = TextBox2
                Text1(3) = TextBox4
                Text1(4) = TextBox5
                Text1(5) = TextBox6
                Text1(6) = TextBox7
                Text1(7) = TextBox8
                Text1(8) = TextBox9
                Text1(9) = TextBox10
                ' Page 2
                Text1(10) = TextBox11
                Text1(11) = TextBox12
                Text1(12) = TextBox13
                Text1(13) = TextBox14
                Text1(14) = TextBox15
                Text1(15) = TextBox16
                Text1(16) = TextBox17
                ' Page 3
                Text1(17) = TextBox18
                Text1(18) = TextBox19
                ' Page 4
                Text1(19) = TextBox20
                Text1(20) = TextBox21
                ' Page 5
                Text1(21) = TextBox22
                Text1(22) = TextBox23
                Text1(23) = TextBox24
                Text1(24) = TextBox25
                Text1(25) = TextBox26
                Text1(26) = TextBox27
                ' Page 6
                Text1(27) = TextBox28
                Text1(28) = TextBox29
                Text1(29) = TextBox30
                Text1(30) = TextBox31
                Text1(31) = TextBox32
                Text1(32) = TextBox33
                ' Page 7
                Text1(33) = TextBox34
                Text1(34) = TextBox35
                Text1(35) = TextBox36
                Text1(36) = TextBox37
                Text1(37) = TextBox38
                Text1(38) = TextBox39
                Text1(39) = TextBox40
                Text1(40) = TextBox41
                Text1(41) = TextBox42
                Text1(42) = TextBox43
                Text1(43) = TextBox44
                Text1(44) = TextBox45
                Text1(45) = TextBox46
                Text1(46) = TextBox47
                Text1(47) = TextBox48

                ' I would have prefered to use CheckListBox, but to make the form readable it is necessary to use
                ' multicolumn, and there is an issue with multicolumn (at least as of Visual Studio 16.9.4) where
                ' the control often (but not always) behaves as if it is not enabled.
                ' Page 1
                Check1(0) = CheckBox1
                Check1(1) = CheckBox2
                Check1(2) = CheckBox3
                Check1(3) = CheckBox4
                Check1(4) = CheckBox5
                ' Page 3
                Check1(5) = CheckBox6
                Check1(6) = CheckBox7
                Check1(7) = CheckBox8
                Check1(8) = CheckBox9
                Check1(9) = CheckBox10
                Check1(10) = CheckBox11
                Check1(11) = CheckBox12
                Check1(12) = CheckBox13
                Check1(13) = CheckBox14
                Check1(14) = CheckBox15
                Check1(15) = CheckBox16
                Check1(16) = CheckBox17
                Check1(17) = CheckBox18
                Check1(18) = CheckBox19
                Check1(19) = CheckBox20
                Check1(20) = CheckBox21
                Check1(21) = CheckBox22
                Check1(22) = CheckBox23
                Check1(23) = CheckBox24
                Check1(24) = CheckBox25
                Check1(25) = CheckBox26
                Check1(26) = CheckBox27
                Check1(27) = CheckBox28
                Check1(28) = CheckBox29
                Check1(29) = CheckBox30
                Check1(30) = CheckBox31
                Check1(31) = CheckBox32
                Check1(32) = CheckBox33
                Check1(33) = CheckBox34
                Check1(34) = CheckBox35
                Check1(35) = CheckBox36
                Check1(36) = CheckBox37
                Check1(37) = CheckBox38
                Check1(38) = CheckBox39
                Check1(39) = CheckBox40
                Check1(40) = CheckBox41
                Check1(41) = CheckBox42
                Check1(42) = CheckBox43
                Check1(43) = CheckBox44
                Check1(44) = CheckBox45
                Check1(45) = CheckBox46
                Check1(46) = CheckBox47
                Check1(47) = CheckBox48
                Check1(48) = CheckBox49
                Check1(49) = CheckBox50
                Check1(50) = CheckBox51
                Check1(51) = CheckBox52
                Check1(52) = CheckBox53
                Check1(53) = CheckBox54
                Check1(54) = CheckBox55
                ' Page 4
                Check1(55) = CheckBox56
                Check1(56) = CheckBox57
                Check1(57) = CheckBox58
                Check1(58) = CheckBox59
                ' Page 5
                Check1(59) = CheckBox60
                Check1(60) = CheckBox61
                Check1(61) = CheckBox62
                Check1(62) = CheckBox63
                Check1(63) = CheckBox64
                Check1(64) = CheckBox65
                Check1(65) = CheckBox66
                ' Page 6
                Check1(66) = CheckBox67
                Check1(67) = CheckBox68
                Check1(68) = CheckBox69
                Check1(69) = CheckBox70
                Check1(70) = CheckBox71
                Check1(71) = CheckBox72
                Check1(72) = CheckBox73
                Check1(73) = CheckBox74
                Check1(74) = CheckBox75
                Check1(75) = CheckBox76
                Check1(76) = CheckBox77
                Check1(77) = CheckBox78
                Check1(78) = CheckBox79
                Check1(79) = CheckBox80
                Check1(80) = CheckBox81
                ' Page 7
                Check1(81) = CheckBox82
                Check1(82) = CheckBox83
                Check1(83) = CheckBox84
                Check1(84) = CheckBox85
                Check1(85) = CheckBox86
                Check1(86) = CheckBox87
                Check1(87) = CheckBox88
                Check1(88) = CheckBox89
                Check1(89) = CheckBox90
                Check1(90) = CheckBox91
                Check1(91) = CheckBox92
                Check1(92) = CheckBox93
                Check1(93) = CheckBox94
                Check1(94) = CheckBox95
                Check1(95) = CheckBox96
                Check1(96) = CheckBox97
                Check1(97) = CheckBox98
                Check1(98) = CheckBox99
                Check1(99) = CheckBox100
                Check1(100) = CheckBox101
                Check1(101) = CheckBox102
                Check1(102) = CheckBox103
                Check1(103) = CheckBox104
                Check1(104) = CheckBox105
                Check1(105) = CheckBox106
                Check1(106) = CheckBox107
                Check1(107) = CheckBox108
                Check1(108) = CheckBox109
                Check1(109) = CheckBox110
                Check1(110) = CheckBox111
                Check1(111) = CheckBox112
                Check1(112) = CheckBox113
                Check1(113) = CheckBox114
                Check1(114) = CheckBox115
                Check1(115) = CheckBox116
                Check1(116) = CheckBox117
                Check1(117) = CheckBox118
                Check1(118) = CheckBox119
                Check1(119) = CheckBox120
                Check1(120) = CheckBox121
                Check1(121) = CheckBox122
                Check1(122) = CheckBox123
                Check1(123) = CheckBox124
                Check1(124) = CheckBox125
                Check1(125) = CheckBox126

                Label(0) = Label26
                Label(1) = Label27
                Label(2) = Label28
                Label(3) = Label29
                Label(4) = Label30
                Label(5) = Label31
                Label(6) = Label32
                Label(7) = Label33
                Label(8) = Label34
                Label(9) = Label35
                Label(10) = Label36
                Label(11) = Label37
                Label(12) = Label38
                Label(13) = Label39
                Label(14) = Label40
                Label(15) = Label63
                Label(16) = Label68
                Label(17) = Label69
                Label(18) = Label70
                Label(19) = Label71
                Label(20) = Label72
                Label(21) = Label73
                Label(22) = Label74
                Label(23) = Label75
                Label(24) = Label76
                Label(25) = Label77
                Label(26) = Label78
                Label(27) = Label79
                Label(28) = Label80

                ' Set up initial dim for symptom list
                ReDim Form1.Symptom.SDate(1)
                ReDim Form1.Symptom.Title(1)
                ReDim Form1.Symptom.P1(6, 1)
                ReDim Form1.Symptom.P1Cb(4, 1)
                ReDim Form1.Symptom.P2(6, 1)
                ReDim Form1.Symptom.P3(1, 1)
                ReDim Form1.Symptom.P3Cb(49, 1)
                ReDim Form1.Symptom.P4Cb(4, 1)
                ReDim Form1.Symptom.P4(1, 1)

                ' Set up dim for patient data
                ReDim Form1.Record.P5AgeCb(4)
                ReDim Form1.Record.P5SexCb(1)
                ReDim Form1.Record.P5(5)
                ReDim Form1.Record.P6(5)
                ReDim Form1.Record.P6Cb(14)
                ReDim Form1.Record.P7(14)
                ReDim Form1.Record.P7Cb(44)

                TabPage1.Show()
                If Form1.SymPtr = 0 Then DisablePrevSymptom()   ' Disable "Prev Symptom" button

                For Idx = 0 To 47
                    If Idx <> 2 Then
                        Text1(Idx).BackColor = Form1.TextBox1.BackColor
                        Text1(Idx).ForeColor = Form1.TextBox1.ForeColor
                        Text1(Idx).Font = Form1.TextBox1.Font
                        Text1(Idx).Text = ""
                    End If
                Next Idx

                For Idx = 0 To 4
                    Check1(Idx).Font = Form1.Font
                    Check1(Idx).Enabled = False
                Next Idx
                For Idx = 5 To 65
                    Check1(Idx).Font = Form1.Font
                    Check1(Idx).Enabled = True
                Next Idx
                For Idx = 66 To 125
                    Check1(Idx).Font = Form1.Font
                    Check1(Idx).Enabled = False
                Next Idx

                Form1.EnableFileQuestNew()
                Form1.EnableFileQuestOpen()
                Form1.DisableFileQuestSave()
                Form1.DisableFileQuestSaveAs()
                Form1.DisableFileQuestClose()

                Me.Visible = True
                Form1.QuestionnaireVisible = True
                Form1.QuestFirstLoad = False

                ' Disable Prev / Next symptom buttons.
                DisablePrevSymptom()
                DisableNextSymptom()
            End If
            Exit Sub

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
            Buttons:=vbOKOnly + vbCritical,
            Title:="ERROR!")
            Close()
        End Try
    End Sub

    Private Sub mnuNewQuestFile_Click(sender As Object, e As EventArgs) Handles mnuNewQuestFile.Click
        Form1.FileQuestNew()
    End Sub

    Private Sub mnuOpenQuestFile_Click(sender As Object, e As EventArgs) Handles mnuOpenQuestFile.Click
        Form1.FileQuestOpen()
        Form1.WriteQ1Screen()
        Form1.WriteQ2Screen()
        Form1.WriteQ3Screen()
        Form1.WriteQ4Screen()
        Form1.WriteQ5Screen()
        Form1.WriteQ6Screen()
        Form1.WriteQ7Screen()
        Form1.QDirty = False
    End Sub

    Private Sub mnuSaveQuestFile_Click(sender As Object, e As EventArgs) Handles mnuSaveQuestFile.Click
        If Form1.QuestFileName = "Untitled" Then
            Form1.FileQuestSaveAs()
        Else
            Form1.SaveCurrentQuestFile()
        End If
    End Sub

    Private Sub mnuSaveQuestFileAs_Click(sender As Object, e As EventArgs) Handles mnuSaveQuestFileAs.Click
        Form1.FileQuestSaveAs()
    End Sub

    Private Sub mnuCloseQuestFile_Click(sender As Object, e As EventArgs) Handles mnuCloseQuestFile.Click
        Form1.FileQuestClose()
    End Sub

    Private Sub mnuPrintQuestionnaire_Click(sender As Object, e As EventArgs)
        Dim pdResult As DialogResult

        Form1.stringBuffer = ""   ' Clear out any previous print job
        HandlePrint()   ' Load the print buffer

        ' Allow the user to choose the page range he or she would
        ' like to print.
        PrintDialog1.AllowSomePages = True

        PrintDialog1.AllowCurrentPage = True
        PrintDialog1.AllowPrintToFile = True
        PrintDialog1.AllowSelection = True

        ' Show the help button.
        PrintDialog1.ShowHelp = True


        PrintDialog1.Document = Form1.PrintDocument1

        pdResult = PrintDialog1.ShowDialog()
        If pdResult = vbOK Then
            Form1.SetupAndPrint()
            Form1.PrintDocument1.Print()  ' Perform print
        End If
        PrintDialog1.Dispose()
    End Sub

    Public Sub HandlePrint()
        Dim pdResult As DialogResult
        Form1.QuestPrint()
        If Form1.PrintPreviewShow Then
            ' Get printer font from registry
            Form1.PrinterFontName = New Font(GetSetting("Hahnasst", "Startup", "PrinterFontName", [Default]:="Microsoft Sans Serif"),
                GetSetting("Hahnasst", "Startup", "PrinterFontSize", [Default]:="8"))
            Form1.PrintPreviewDialog1.ShowIcon = False
            Form1.PrintPreviewDialog1.WindowState = FormWindowState.Normal
            Form1.PrintPreviewDialog1.StartPosition = FormStartPosition.CenterScreen
            Form1.PrintPreviewDialog1.ClientSize = New Size(800, 1200)
            Form1.PrintPreviewDialog1.UseAntiAlias = True
            Form1.PrintPreviewDialog1.Document = Form1.PrintDocument1
            Form1.PrintPreviewDialog1.ShowDialog()

        Else
            ' Allow the user to choose the page range he or she would
            ' like to print.
            Form1.PrintDialog1.AllowSomePages = True

            Form1.PrintDialog1.AllowCurrentPage = True
            Form1.PrintDialog1.AllowPrintToFile = True
            Form1.PrintDialog1.AllowSelection = True

            ' Show the help button.
            Form1.PrintDialog1.ShowHelp = True

            pdResult = Form1.PrintDialog1.ShowDialog()
            If pdResult = vbOK Then
                Form1.SetupAndPrint()
                Form1.PrintDocument1.Print()  ' Perform print
                Form1.SympPrint = False
                Form1.CasePrint = False
                Form1.RemPrint = False
                Form1.QPrint = False
            End If
            Form1.PrintDialog1.Dispose()
        End If
    End Sub


    ' It would be more efficient to do the copy of page items to structs in _LostFocus.  Unfortunately the
    ' lost focus event is not raised when selecting the menu items to save files, so the last-edited
    ' item would not get saved.
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Form1.QDirty = True
        Form1.Record.Name = Text1(0).Text
        Text_Changed1(0)
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Form1.QDirty = True
        Form1.Symptom.Title(Form1.SymPtr) = Text1(1).Text
        Text_Changed1(1)
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        Form1.QDirty = True
        Form1.Symptom.P1(0, Form1.SymPtr) = Text1(3).Text
        Text_Changed1(3)
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        Form1.QDirty = True
        Form1.Symptom.P1(1, Form1.SymPtr) = Text1(4).Text
        Text_Changed1(4)
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        Form1.QDirty = True
        Form1.Symptom.P1(2, Form1.SymPtr) = Text1(5).Text
        Text_Changed1(5)
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        Form1.QDirty = True
        Form1.Symptom.P1(3, Form1.SymPtr) = Text1(6).Text
        Text_Changed1(6)
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        Form1.QDirty = True
        Form1.Symptom.P1(4, Form1.SymPtr) = Text1(7).Text
        Text_Changed1(7)
    End Sub

    Private Sub Text_Changed1(Index As Integer)
        Dim Idx As Integer  ' check box index
        Dim msg As String   ' error message string

        Try
            Form1.ErrorFlag = 8003

            If Index >= 3 And Index <= 9 And Text1(Index).Text <> "" Then
                For Idx = 0 To 4
                    Check1(Idx).Enabled = True
                Next Idx
            ElseIf Text1(1).Text = "" And
                    Text1(3).Text = "" And
                    Text1(4).Text = "" And
                    Text1(5).Text = "" And
                    Text1(6).Text = "" And
                    Text1(7).Text = "" And
                    Text1(8).Text = "" And
                    Text1(9).Text = "" Then
                For Idx = 0 To 4
                    Check1(Idx).Enabled = False
                Next Idx
            End If

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub Text_Changed2(Index As Integer)
        Dim msg As String   ' error message string

        Try
            Form1.ErrorFlag = 8004
            If Index >= 10 Then
                If Form1.Symptom.P2(Index - 10, Form1.SymPtr) <> Text1(Index).Text Then
                    Form1.QDirty = True
                    Form1.Symptom.P2(Index - 10, Form1.SymPtr) = Text1(Index).Text
                End If
            End If

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub Text_Changed3(Index As Integer)
        Dim msg As String   ' error message string

        Try
            Form1.ErrorFlag = 8005
            If Index >= 17 Then
                If Form1.Symptom.P3(Index - 17, Form1.SymPtr) <> Text1(Index).Text Then
                    Form1.QDirty = True
                    Form1.Symptom.P3(Index - 17, Form1.SymPtr) = Text1(Index).Text
                End If
            End If

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub Text_Changed4(Index As Integer)
        Dim msg As String   ' error message string

        Try
            Form1.ErrorFlag = 8006
            If Index >= 19 Then
                If Form1.Symptom.P4(Index - 19, Form1.SymPtr) <> Text1(Index).Text Then
                    Form1.QDirty = True
                    Form1.Symptom.P4(Index - 19, Form1.SymPtr) = Text1(Index).Text
                End If
            End If

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub Text_Changed5(Index As Integer)
        Dim msg As String   ' error message string

        Try
            Form1.ErrorFlag = 8007
            If Index >= 21 Then
                If Form1.Record.P5(Index - 21) <> Text1(Index).Text Then
                    Form1.QDirty = True
                    Form1.Record.P5(Index - 21) = Text1(Index).Text
                End If
            End If

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub Text_Changed6(Index As Integer)
        Dim msg As String   ' error message string

        Try
            Form1.ErrorFlag = 8008
            If Index >= 27 Then
                If Form1.Record.P6(Index - 27) <> Text1(Index).Text Then
                    Form1.QDirty = True
                    Form1.Record.P6(Index - 27) = Text1(Index).Text
                End If
            End If

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub Text_Changed7(Index As Integer)
        Dim msg As String   ' error message string

        Try
            Form1.ErrorFlag = 8009
            If Index >= 33 Then
                If Form1.Record.P7(Index - 33) <> Text1(Index).Text Then
                    Form1.QDirty = True
                    Form1.Record.P7(Index - 33) = Text1(Index).Text
                End If
            End If

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub
    Private Sub Check_radio2(index As Integer, clickIndex As Integer)
        Dim msg As String   ' error message string

        Try
            Form1.ErrorFlag = 8010
            If clickIndex = index Then
                If Check1(index).Checked Then
                    Check1(index + 1).Checked = False
                End If
            Else
                If Check1(index + 1).Checked Then
                    Check1(index).Checked = False
                End If
            End If

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub Check_radio3(index As Integer, clickIndex As Integer)
        Dim msg As String   ' error message string

        Try
            Form1.ErrorFlag = 8011
            If clickIndex = index Then
                If Check1(index).Checked Then
                    Check1(index + 1).Checked = False
                    Check1(index + 2).Checked = False
                End If
            ElseIf clickIndex = index + 1 Then
                If Check1(index + 1).Checked Then
                    Check1(index).Checked = False
                    Check1(index + 2).Checked = False
                End If
            Else
                If Check1(index + 2).Checked Then
                    Check1(index).Checked = False
                    Check1(index + 1).Checked = False
                End If
            End If

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub Check_radio4(index As Integer, clickIndex As Integer)
        Dim msg As String   ' error message string

        Try
            Form1.ErrorFlag = 8012
            If clickIndex = index Then
                If Check1(index).Checked Then
                    Check1(index + 1).Checked = False
                    Check1(index + 2).Checked = False
                    Check1(index + 3).Checked = False
                End If
            ElseIf clickIndex = index + 1 Then
                If Check1(index + 1).Checked Then
                    Check1(index).Checked = False
                    Check1(index + 2).Checked = False
                    Check1(index + 3).Checked = False
                End If
            ElseIf clickIndex = index + 2 Then
                If Check1(index + 2).Checked Then
                    Check1(index).Checked = False
                    Check1(index + 1).Checked = False
                    Check1(index + 3).Checked = False
                End If
            Else
                If Check1(index + 3).Checked Then
                    Check1(index).Checked = False
                    Check1(index + 1).Checked = False
                    Check1(index + 2).Checked = False
                End If
            End If

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub Check_radio5(index As Integer, clickIndex As Integer)
        Dim msg As String   ' error message string

        Try
            Form1.ErrorFlag = 8013
            If clickIndex = index Then
                If Check1(index).Checked Then
                    Check1(index + 1).Checked = False
                    Check1(index + 2).Checked = False
                    Check1(index + 3).Checked = False
                    Check1(index + 4).Checked = False
                End If
            ElseIf clickIndex = index + 1 Then
                If Check1(index + 1).Checked Then
                    Check1(index).Checked = False
                    Check1(index + 2).Checked = False
                    Check1(index + 3).Checked = False
                    Check1(index + 4).Checked = False
                End If
            ElseIf clickIndex = index + 2 Then
                If Check1(index + 2).Checked Then
                    Check1(index).Checked = False
                    Check1(index + 1).Checked = False
                    Check1(index + 3).Checked = False
                    Check1(index + 4).Checked = False
                End If
            ElseIf clickIndex = index + 3 Then
                If Check1(index + 3).Checked Then
                    Check1(index).Checked = False
                    Check1(index + 1).Checked = False
                    Check1(index + 2).Checked = False
                    Check1(index + 4).Checked = False
                End If
            Else
                If Check1(index + 4).Checked Then
                    Check1(index).Checked = False
                    Check1(index + 1).Checked = False
                    Check1(index + 2).Checked = False
                    Check1(index + 3).Checked = False
                End If
            End If

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Check_radio2(0, 0) ' Make it behave as a radio button
        HandleP1Cb0()
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        Check_radio2(0, 1) ' Make it behave as a radio button
        HandleP1Cb0()
    End Sub

    Private Sub HandleP1Cb0()
        Form1.Symptom.P1Cb(0, Form1.SymPtr) = CheckBox1.Checked
        Form1.Symptom.P1Cb(1, Form1.SymPtr) = CheckBox2.Checked
        Form1.QDirty = True
    End Sub
    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        Form1.QDirty = True
        Form1.Symptom.P1(6, Form1.SymPtr) = Text1(9).Text
        Text_Changed1(9)
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        Check_radio3(2, 2)
        HandleP1CB3()
    End Sub
    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        Check_radio3(2, 3)
        HandleP1CB3()
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        Check_radio3(2, 4)
        HandleP1CB3()
    End Sub

    Public Sub HandleP1CB3()
        Form1.Symptom.P1Cb(2, Form1.SymPtr) = CheckBox3.Checked
        Form1.Symptom.P1Cb(3, Form1.SymPtr) = CheckBox4.Checked
        Form1.Symptom.P1Cb(4, Form1.SymPtr) = CheckBox5.Checked
        Form1.QDirty = True
    End Sub

    Private Sub Command1_Click(sender As Object, e As EventArgs) Handles Command1.Click ' Tab 1 Continue
        TabControl1.SelectedTab = TabPage2
    End Sub

    Private Sub Command2_Click(sender As Object, e As EventArgs) Handles Command2.Click ' Tab 1 Cancel
        Close()
    End Sub

    Private Sub Command7_Click(sender As Object, e As EventArgs) Handles Command7.Click ' P1 Prev Symptom
        Form1.SymPtr = Form1.SymPtr - 1
        EnableNextSymptom() ' Enable Next Symptom
        If Form1.SymPtr = 0 Then DisablePrevSymptom()
        Form1.WriteQ1Screen()
        Form1.WriteQ2Screen()
        Form1.WriteQ3Screen()
        Form1.WriteQ4Screen()
        TabControl1.SelectedTab = TabPage1
    End Sub

    Private Sub Command6_Click(sender As Object, e As EventArgs) Handles Command6.Click ' P1 Next Symptom
        Form1.SymPtr += 1
        EnablePrevSymptom() ' Enable Prev Symptom
        If Form1.SymPtr = Form1.SympSize - 1 Then DisableNextSymptom()
        Form1.WriteQ1Screen()
        Form1.WriteQ2Screen()
        Form1.WriteQ3Screen()
        Form1.WriteQ4Screen()
        TabControl1.SelectedTab = TabPage1
    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged
        Form1.QDirty = True
        Form1.Symptom.P1(5, Form1.SymPtr) = Text1(8).Text
        Text_Changed1(8)
    End Sub

    Private Sub Label19_Click(sender As Object, e As EventArgs) Handles Label19.Click

    End Sub

    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Click
        Dim msg As Integer

        Try
            Form1.ErrorFlag = 8014


            Form1.WriteQ2Screen()

            Exit Sub

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            Close()
        End Try
    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged
        Text_Changed2(10)
    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged
        Text_Changed2(11)
    End Sub

    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged
        Text_Changed2(12)
    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged
        Text_Changed2(13)
    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged
        Text_Changed2(14)
    End Sub

    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles TextBox16.TextChanged
        Text_Changed2(15)
    End Sub

    Private Sub TextBox17_TextChanged(sender As Object, e As EventArgs) Handles TextBox17.TextChanged
        Text_Changed2(16)
    End Sub

    Private Sub Command8_Click(sender As Object, e As EventArgs) Handles Command8.Click ' Tab 2 Continue
        TabControl1.SelectedTab = TabPage3
    End Sub

    Private Sub Command9_Click(sender As Object, e As EventArgs) Handles Command9.Click ' Tab 2 Cancle
        Close()
    End Sub

    Private Sub Command10_Click(sender As Object, e As EventArgs) Handles Command10.Click   ' Page 2 Prev Symptom
        Form1.SymPtr = Form1.SymPtr - 1
        EnableNextSymptom() ' Enable Next Symptom
        If Form1.SymPtr = 0 Then DisablePrevSymptom()
        Form1.WriteQ1Screen()
        Form1.WriteQ2Screen()
        Form1.WriteQ3Screen()
        Form1.WriteQ4Screen()
        TabControl1.SelectedTab = TabPage2
    End Sub

    Private Sub Command11_Click(sender As Object, e As EventArgs) Handles Command11.Click   ' Page 2 Next Symptom
        Form1.SymPtr += 1
        EnablePrevSymptom() ' Enable Prev Symptom
        If Form1.SymPtr = Form1.SympSize - 1 Then DisableNextSymptom()
        Form1.WriteQ1Screen()
        Form1.WriteQ2Screen()
        Form1.WriteQ3Screen()
        Form1.WriteQ4Screen()
        TabControl1.SelectedTab = TabPage2
    End Sub

    Private Sub TabPage3_Click(sender As Object, e As EventArgs) Handles TabPage3.Click
        Dim Idx As Integer
        Dim msg As String

        Try
            Form1.ErrorFlag = 8015

            For Idx = 17 To 18
                Text1(Idx).BackColor = Form1.BackColor
                Text1(Idx).ForeColor = Form1.ForeColor
                Text1(Idx).Font = Form1.Font
            Next Idx

            Form1.WriteQ3Screen()

            Text = "Questionnaire Page 3 - Symptom " + Str(Form1.SymPtr + 1)

            Visible = True
            Exit Sub

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            Close()
        End Try
    End Sub

    Private Sub TextBox18_TextChanged(sender As Object, e As EventArgs) Handles TextBox18.TextChanged
        Text_Changed3(17)
    End Sub

    Private Sub TextBox19_TextChanged(sender As Object, e As EventArgs) Handles TextBox19.TextChanged
        Text_Changed3(18)
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        Check_radio4(5, 5) ' Make it behave as a radio button
        HandleP3CB6()
    End Sub

    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        Check_radio4(5, 6) ' Make it behave as a radio button
        HandleP3CB6()
    End Sub

    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
        Check_radio4(5, 7) ' Make it behave as a radio button
        HandleP3CB6()
    End Sub

    Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox9.CheckedChanged
        Check_radio4(5, 8) ' Make it behave as a radio button
        HandleP3CB6()
    End Sub

    Private Sub HandleP3CB6()
        Dim ItmPtr As Integer   'pointer to page item
        For ItmPtr = 0 To 3
            Form1.Symptom.P3Cb(ItmPtr, Form1.SymPtr) = Check1(ItmPtr + 5).Checked
        Next ItmPtr
        Form1.QDirty = True
    End Sub

    Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged
        Check_radio4(9, 9) ' Make it behave as a radio button
        HandleP3CB10()
    End Sub

    Private Sub CheckBox11_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox11.CheckedChanged
        Check_radio4(9, 10) ' Make it behave as a radio button
        HandleP3CB10()
    End Sub

    Private Sub CheckBox12_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox12.CheckedChanged
        Check_radio4(9, 11) ' Make it behave as a radio button
        HandleP3CB10()
    End Sub

    Private Sub CheckBox13_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox13.CheckedChanged
        Check_radio4(9, 12) ' Make it behave as a radio button
        HandleP3CB10()
    End Sub

    Private Sub HandleP3CB10()
        Dim ItmPtr As Integer   'pointer to page item
        For ItmPtr = 4 To 7
            Form1.Symptom.P3Cb(ItmPtr, Form1.SymPtr) = Check1(ItmPtr + 5).Checked
        Next ItmPtr
        Form1.QDirty = True
    End Sub

    Private Sub CheckBox14_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox14.CheckedChanged
        Check_radio3(13, 13) ' Make it behave as a radio button
        HandleP3CB14()
    End Sub

    Private Sub CheckBox15_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox15.CheckedChanged
        Check_radio3(13, 14) ' Make it behave as a radio button
        HandleP3CB14()
    End Sub

    Private Sub CheckBox16_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox16.CheckedChanged
        Check_radio3(13, 15) ' Make it behave as a radio button
        HandleP3CB14()
    End Sub

    Private Sub HandleP3CB14()
        Dim ItmPtr As Integer   'pointer to page item
        For ItmPtr = 8 To 10
            Form1.Symptom.P3Cb(ItmPtr, Form1.SymPtr) = Check1(ItmPtr + 5).Checked
        Next ItmPtr
        Form1.QDirty = True
    End Sub

    Private Sub CheckBox17_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox17.CheckedChanged
        Check_radio3(16, 16) ' Make it behave as a radio button
        HandleP3CB17()
    End Sub

    Private Sub CheckBox18_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox18.CheckedChanged
        Check_radio3(16, 17) ' Make it behave as a radio button
        HandleP3CB17()
    End Sub

    Private Sub CheckBox19_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox19.CheckedChanged
        Check_radio3(16, 18) ' Make it behave as a radio button
        HandleP3CB17()
    End Sub

    Private Sub HandleP3CB17()
        Dim ItmPtr As Integer   'pointer to page item
        For ItmPtr = 11 To 13
            Form1.Symptom.P3Cb(ItmPtr, Form1.SymPtr) = Check1(ItmPtr + 5).Checked
        Next ItmPtr
        Form1.QDirty = True
    End Sub

    Private Sub CheckBox20_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox20.CheckedChanged
        Check_radio3(19, 19) ' Make it behave as a radio button
        HandleP3CB20()
    End Sub

    Private Sub CheckBox21_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox21.CheckedChanged
        Check_radio3(19, 20) ' Make it behave as a radio button
        HandleP3CB20()
    End Sub

    Private Sub CheckBox22_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox22.CheckedChanged
        Check_radio3(19, 21) ' Make it behave as a radio button
        HandleP3CB20()
    End Sub

    Private Sub HandleP3CB20()
        Dim ItmPtr As Integer   'pointer to page item
        For ItmPtr = 14 To 16
            Form1.Symptom.P3Cb(ItmPtr, Form1.SymPtr) = Check1(ItmPtr + 5).Checked
        Next ItmPtr
        Form1.QDirty = True
    End Sub

    Private Sub CheckBox23_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox23.CheckedChanged
        Check_radio3(22, 22) ' Make it behave as a radio button
        HandleP3CB23()
    End Sub

    Private Sub CheckBox24_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox24.CheckedChanged
        Check_radio3(22, 23) ' Make it behave as a radio button
        HandleP3CB23()
    End Sub

    Private Sub CheckBox25_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox25.CheckedChanged
        Check_radio3(22, 24) ' Make it behave as a radio button
        HandleP3CB23()
    End Sub

    Private Sub HandleP3CB23()
        Dim ItmPtr As Integer   'pointer to page item
        For ItmPtr = 17 To 19
            Form1.Symptom.P3Cb(ItmPtr, Form1.SymPtr) = Check1(ItmPtr + 5).Checked
        Next ItmPtr
        Form1.QDirty = True
    End Sub

    Private Sub CheckBox26_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox26.CheckedChanged
        Check_radio3(25, 25) ' Make it behave as a radio button
        HandleP3CB26()
    End Sub

    Private Sub CheckBox27_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox27.CheckedChanged
        Check_radio3(25, 26) ' Make it behave as a radio button
        HandleP3CB26()
    End Sub

    Private Sub CheckBox28_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox28.CheckedChanged
        Check_radio3(25, 27) ' Make it behave as a radio button
        HandleP3CB26()
    End Sub

    Private Sub HandleP3CB26()
        Dim ItmPtr As Integer   'pointer to page item
        For ItmPtr = 20 To 22
            Form1.Symptom.P3Cb(ItmPtr, Form1.SymPtr) = Check1(ItmPtr + 5).Checked
        Next ItmPtr
        Form1.QDirty = True
    End Sub

    Private Sub CheckBox29_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox29.CheckedChanged
        Check_radio3(28, 28) ' Make it behave as a radio button
        HandleP3B29()
    End Sub

    Private Sub CheckBox30_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox30.CheckedChanged
        Check_radio3(28, 29) ' Make it behave as a radio button
        HandleP3B29()
    End Sub

    Private Sub CheckBox31_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox31.CheckedChanged
        Check_radio3(28, 30) ' Make it behave as a radio button
        HandleP3B29()
    End Sub

    Private Sub HandleP3B29()
        Dim ItmPtr As Integer   'pointer to page item
        For ItmPtr = 23 To 25
            Form1.Symptom.P3Cb(ItmPtr, Form1.SymPtr) = Check1(ItmPtr + 5).Checked
        Next ItmPtr
        Form1.QDirty = True
    End Sub

    Private Sub CheckBox32_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox32.CheckedChanged
        Check_radio3(31, 31) ' Make it behave as a radio button
        HandleP3B32()
    End Sub

    Private Sub CheckBox33_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox33.CheckedChanged
        Check_radio3(31, 32) ' Make it behave as a radio button
        HandleP3B32()
    End Sub

    Private Sub CheckBox34_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox34.CheckedChanged
        Check_radio3(31, 33) ' Make it behave as a radio button
        HandleP3B32()
    End Sub

    Private Sub HandleP3B32()
        Dim ItmPtr As Integer   'pointer to page item
        For ItmPtr = 26 To 28
            Form1.Symptom.P3Cb(ItmPtr, Form1.SymPtr) = Check1(ItmPtr + 5).Checked
        Next ItmPtr
        Form1.QDirty = True
    End Sub

    Private Sub CheckBox35_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox35.CheckedChanged
        Check_radio3(34, 34) ' Make it behave as a radio button
        HandleP3B35()
    End Sub

    Private Sub CheckBox36_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox36.CheckedChanged
        Check_radio3(34, 35) ' Make it behave as a radio button
        HandleP3B35()
    End Sub

    Private Sub CheckBox37_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox37.CheckedChanged
        Check_radio3(34, 36) ' Make it behave as a radio button
        HandleP3B35()
    End Sub

    Private Sub HandleP3B35()
        Dim ItmPtr As Integer   'pointer to page item
        For ItmPtr = 29 To 31
            Form1.Symptom.P3Cb(ItmPtr, Form1.SymPtr) = Check1(ItmPtr + 5).Checked
        Next ItmPtr
        Form1.QDirty = True
    End Sub

    Private Sub CheckBox38_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox38.CheckedChanged
        Check_radio3(37, 37) ' Make it behave as a radio button
        HandleP3B38()
    End Sub

    Private Sub CheckBox39_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox39.CheckedChanged
        Check_radio3(37, 38) ' Make it behave as a radio button
        HandleP3B38()
    End Sub

    Private Sub CheckBox40_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox40.CheckedChanged
        Check_radio3(37, 39) ' Make it behave as a radio button
        HandleP3B38()
    End Sub

    Private Sub HandleP3B38()
        Dim ItmPtr As Integer   'pointer to page item
        For ItmPtr = 32 To 34
            Form1.Symptom.P3Cb(ItmPtr, Form1.SymPtr) = Check1(ItmPtr + 5).Checked
        Next ItmPtr
        Form1.QDirty = True
    End Sub

    Private Sub CheckBox41_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox41.CheckedChanged
        Check_radio3(40, 40) ' Make it behave as a radio button
        HandleP3B41()
    End Sub

    Private Sub CheckBox42_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox42.CheckedChanged
        Check_radio3(40, 41) ' Make it behave as a radio button
        HandleP3B41()
    End Sub

    Private Sub CheckBox43_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox43.CheckedChanged
        Check_radio3(40, 42) ' Make it behave as a radio button
        HandleP3B41()
    End Sub

    Private Sub HandleP3B41()
        Dim ItmPtr As Integer   'pointer to page item
        For ItmPtr = 35 To 37
            Form1.Symptom.P3Cb(ItmPtr, Form1.SymPtr) = Check1(ItmPtr + 5).Checked
        Next ItmPtr
        Form1.QDirty = True
    End Sub

    Private Sub CheckBox44_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox44.CheckedChanged
        Check_radio3(43, 43) ' Make it behave as a radio button
        HandleP3B44()
    End Sub

    Private Sub CheckBox45_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox45.CheckedChanged
        Check_radio3(43, 44) ' Make it behave as a radio button
        HandleP3B44()
    End Sub

    Private Sub CheckBox46_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox46.CheckedChanged
        Check_radio3(43, 45) ' Make it behave as a radio button
        HandleP3B44()
    End Sub

    Private Sub HandleP3B44()
        Dim ItmPtr As Integer   'pointer to page item
        For ItmPtr = 38 To 40
            Form1.Symptom.P3Cb(ItmPtr, Form1.SymPtr) = Check1(ItmPtr + 5).Checked
        Next ItmPtr
        Form1.QDirty = True
    End Sub

    Private Sub CheckBox47_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox47.CheckedChanged
        Check_radio3(46, 46) ' Make it behave as a radio button
        HandleP3B47()
    End Sub

    Private Sub CheckBox48_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox48.CheckedChanged
        Check_radio3(46, 47) ' Make it behave as a radio button
        HandleP3B47()
    End Sub

    Private Sub CheckBo49_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox49.CheckedChanged
        Check_radio3(46, 48) ' Make it behave as a radio button
        HandleP3B47()
    End Sub

    Private Sub HandleP3B47()
        Dim ItmPtr As Integer   'pointer to page item
        For ItmPtr = 41 To 43
            Form1.Symptom.P3Cb(ItmPtr, Form1.SymPtr) = Check1(ItmPtr + 5).Checked
        Next ItmPtr
        Form1.QDirty = True
    End Sub

    Private Sub CheckBox50_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox50.CheckedChanged
        Check_radio3(49, 49) ' Make it behave as a radio button
        HandleP3B50()
    End Sub

    Private Sub CheckBox51_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox51.CheckedChanged
        Check_radio3(49, 50) ' Make it behave as a radio button
        HandleP3B50()
    End Sub

    Private Sub CheckBox52_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox52.CheckedChanged
        Check_radio3(49, 51) ' Make it behave as a radio button
        HandleP3B50()
    End Sub

    Private Sub HandleP3B50()
        Dim ItmPtr As Integer   'pointer to page item

        For ItmPtr = 44 To 46
            Form1.Symptom.P3Cb(ItmPtr, Form1.SymPtr) = Check1(ItmPtr + 5).Checked
        Next ItmPtr
        Form1.QDirty = True
    End Sub

    Private Sub CheckBox53_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox53.CheckedChanged
        Check_radio3(52, 52) ' Make it behave as a radio button
        HandleP3B53()
    End Sub

    Private Sub CheckBox54_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox54.CheckedChanged
        Check_radio3(52, 53) ' Make it behave as a radio button
        HandleP3B53()
    End Sub

    Private Sub CheckBox55_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox55.CheckedChanged
        Check_radio3(52, 54) ' Make it behave as a radio button
        HandleP3B53()
    End Sub

    Private Sub HandleP3B53()
        Dim ItmPtr As Integer   'pointer to page item
        For ItmPtr = 47 To 49
            Form1.Symptom.P3Cb(ItmPtr, Form1.SymPtr) = Check1(ItmPtr + 5).Checked
        Next ItmPtr
        Form1.QDirty = True
    End Sub
    Private Sub Command12_Click(sender As Object, e As EventArgs) Handles Command12.Click   ' Tab 3 Continue
        '        Form1.ReadQ3Screen()
        TabControl1.SelectedTab = TabPage4
    End Sub

    Private Sub Command13_Click(sender As Object, e As EventArgs) Handles Command13.Click   ' Tab 3 Cancel
        Close()
    End Sub

    Private Sub Command14_Click(sender As Object, e As EventArgs) Handles Command14.Click   ' Page 3 Prev Symptom
        Form1.SymPtr = Form1.SymPtr - 1
        EnableNextSymptom() ' Enable Next Symptom
        If Form1.SymPtr = 0 Then DisablePrevSymptom()
        Form1.WriteQ1Screen()
        Form1.WriteQ2Screen()
        Form1.WriteQ3Screen()
        Form1.WriteQ4Screen()
        TabControl1.SelectedTab = TabPage3
    End Sub

    Private Sub Command15_Click(sender As Object, e As EventArgs) Handles Command15.Click   ' Page 3 Next Symptom
        Form1.SymPtr += 1
        EnablePrevSymptom() ' Enable Prev Symptom
        If Form1.SymPtr = Form1.SympSize - 1 Then DisableNextSymptom()
        Form1.WriteQ1Screen()
        Form1.WriteQ2Screen()
        Form1.WriteQ3Screen()
        Form1.WriteQ4Screen()
        TabControl1.SelectedTab = TabPage3
    End Sub

    Private Sub Label43_Click(sender As Object, e As EventArgs) Handles Label43.Click

    End Sub

    Private Sub TextBox21_TextChanged(sender As Object, e As EventArgs) Handles TextBox21.TextChanged
        Text_Changed4(20)
    End Sub

    Private Sub TabPage4_Click(sender As Object, e As EventArgs) Handles TabPage4.Click
        Dim Idx As Integer
        Dim msg As String

        Try
            Form1.ErrorFlag = 8016

            For Idx = 19 To 20
                Text1(Idx).BackColor = Form1.BackColor
                Text1(Idx).ForeColor = Form1.ForeColor
                Text1(Idx).Font = Form1.Font
            Next Idx

            Form1.WriteQ4Screen()

            Text = "Questionnaire Page 4 - Symptom " + Str(Form1.SymPtr + 1)

            Visible = True
            Exit Sub

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
            Buttons:=vbOKOnly + vbCritical,
            Title:="ERROR!")
            Close()
        End Try
    End Sub

    Private Sub CheckBox56_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox56.CheckedChanged
        Check_radio4(55, 55) ' Make it behave as a radio button
        HandleP4B56()
    End Sub

    Private Sub CheckBox57_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox57.CheckedChanged
        Check_radio4(55, 56) ' Make it behave as a radio button
        HandleP4B56()
    End Sub

    Private Sub CheckBox58_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox58.CheckedChanged
        Check_radio4(55, 57) ' Make it behave as a radio button
        HandleP4B56()
    End Sub

    Private Sub CheckBox59_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox59.CheckedChanged
        Check_radio4(55, 58) ' Make it behave as a radio button
        HandleP4B56()
    End Sub

    Private Sub HandleP4B56()
        Dim ItmPtr As Integer   'pointer to page item
        For ItmPtr = 0 To 3
            Form1.Symptom.P4Cb(ItmPtr, Form1.SymPtr) = Check1(ItmPtr + 55).Checked
        Next ItmPtr

        Form1.QDirty = True
    End Sub

    Private Sub TextBox20_TextChanged(sender As Object, e As EventArgs) Handles TextBox20.TextChanged
        Text_Changed4(19)
    End Sub

    Private Sub Command16_Click(sender As Object, e As EventArgs) Handles Command16.Click   ' Tab 4 Continue
        '        Form1.ReadQ4Screen()
        TabControl1.SelectedTab = TabPage5
    End Sub

    Private Sub Command17_Click(sender As Object, e As EventArgs) Handles Command17.Click   ' Tab 4 Cancel
        Close()
    End Sub

    Private Sub Command18_Click(sender As Object, e As EventArgs) Handles Command18.Click   ' Page 4 Prev Symptom
        Form1.SymPtr = Form1.SymPtr - 1
        EnableNextSymptom() ' Enable Next Symptom
        If Form1.SymPtr = 0 Then DisablePrevSymptom()
        Form1.WriteQ1Screen()
        Form1.WriteQ2Screen()
        Form1.WriteQ3Screen()
        Form1.WriteQ4Screen()
        TabControl1.SelectedTab = TabPage4
    End Sub

    Private Sub Command19_Click(sender As Object, e As EventArgs) Handles Command19.Click   ' Page 4 Next Symptom
        Form1.SymPtr += 1
        EnablePrevSymptom() ' Enable Prev Symptom
        If Form1.SymPtr = Form1.SympSize - 1 Then DisableNextSymptom()
        Form1.WriteQ1Screen()
        Form1.WriteQ2Screen()
        Form1.WriteQ3Screen()
        Form1.WriteQ4Screen()
        TabControl1.SelectedTab = TabPage4
    End Sub

    Private Sub TabPage5_Click(sender As Object, e As EventArgs) Handles TabPage5.Click
        Dim Idx As Integer
        Dim msg As String

        Try
            Form1.ErrorFlag = 8017

            For Idx = 21 To 26
                Text1(Idx).BackColor = Form1.BackColor
                Text1(Idx).ForeColor = Form1.ForeColor
                Text1(Idx).Font = Form1.Font
            Next Idx

            For Idx = 19 To 20
                Check1(Idx).Font = Form1.Font
            Next Idx

            Form1.WriteQ5Screen()

            Visible = True
            Exit Sub

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            Close()
        End Try
    End Sub

    Private Sub CheckBox60_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox60.CheckedChanged
        Check_radio5(59, 59) ' Make it behave as a radio button
        HandleP5B59()
    End Sub

    Private Sub CheckBox61_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox61.CheckedChanged
        Check_radio5(59, 60) ' Make it behave as a radio button
        HandleP5B59()
    End Sub

    Private Sub CheckBox62_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox62.CheckedChanged
        Check_radio5(59, 61) ' Make it behave as a radio button
        HandleP5B59()
    End Sub

    Private Sub CheckBox63_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox63.CheckedChanged
        Check_radio5(59, 62) ' Make it behave as a radio button
        HandleP5B59()
    End Sub

    Private Sub CheckBox64_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox64.CheckedChanged
        Check_radio5(59, 63) ' Make it behave as a radio button
        HandleP5B59()
    End Sub

    Private Sub HandleP5B59()
        Form1.Record.P5AgeCb(0) = CheckBox60.Checked
        Form1.Record.P5AgeCb(1) = CheckBox61.Checked
        Form1.Record.P5AgeCb(2) = CheckBox62.Checked
        Form1.Record.P5AgeCb(3) = CheckBox63.Checked
        Form1.Record.P5AgeCb(4) = CheckBox64.Checked
        Form1.QDirty = True
    End Sub

    Private Sub CheckBox65_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox65.CheckedChanged
        Check_radio2(64, 64) ' Make it behave as a radio button
        HandleP5B64()
    End Sub

    Private Sub CheckBox66_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox66.CheckedChanged
        Check_radio2(64, 65) ' Make it behave as a radio button
        HandleP5B64()
    End Sub

    Private Sub HandleP5B64()
        Form1.Record.P5SexCb(0) = CheckBox65.Checked
        Form1.Record.P5SexCb(1) = CheckBox66.Checked
        Form1.QDirty = True
    End Sub

    Private Sub TextBox22_TextChanged(sender As Object, e As EventArgs) Handles TextBox22.TextChanged
        Text_Changed5(21)
    End Sub

    Private Sub TextBox23_TextChanged(sender As Object, e As EventArgs) Handles TextBox23.TextChanged
        Text_Changed5(22)
    End Sub

    Private Sub TextBox24_TextChanged(sender As Object, e As EventArgs) Handles TextBox24.TextChanged
        Text_Changed5(23)
    End Sub

    Private Sub TextBox25_TextChanged(sender As Object, e As EventArgs) Handles TextBox25.TextChanged
        Text_Changed5(24)
    End Sub

    Private Sub TextBox26_TextChanged(sender As Object, e As EventArgs) Handles TextBox26.TextChanged
        Text_Changed5(25)
    End Sub

    Private Sub TextBox27_TextChanged(sender As Object, e As EventArgs) Handles TextBox27.TextChanged
        Text_Changed5(26)
    End Sub

    Private Sub Command21_Click(sender As Object, e As EventArgs) Handles Command21.Click   ' Tab 5 Cancel
        Close()
    End Sub

    Private Sub Command20_Click(sender As Object, e As EventArgs) Handles Command20.Click   ' Tab 5 Continue
        TabControl1.SelectedTab = TabPage6
    End Sub

    Private Sub TextBox28_TextChanged(sender As Object, e As EventArgs) Handles TextBox28.TextChanged
        Text_Changed6(27)
    End Sub

    Private Sub TabPage6_Click(sender As Object, e As EventArgs) Handles TabPage6.Click
        Dim Idx As Integer
        Dim msg As String

        Try
            Form1.ErrorFlag = 8018

            For Idx = 27 To 32
                Text1(Idx).BackColor = Form1.BackColor
                Text1(Idx).ForeColor = Form1.ForeColor
                Text1(Idx).Font = Form1.Font
            Next Idx

            Form1.WriteQ6Screen()

            Visible = True
            Exit Sub

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            Close()
        End Try
    End Sub

    Private Sub TextBox29_TextChanged(sender As Object, e As EventArgs) Handles TextBox29.TextChanged
        Text_Changed6(28)
        If TextBox29.Text <> "" Then
            CheckBox67.Enabled = True
            CheckBox68.Enabled = True
            CheckBox69.Enabled = True
        Else
            CheckBox67.Enabled = False
            CheckBox68.Enabled = False
            CheckBox69.Enabled = False
        End If
    End Sub

    Private Sub CheckBox67_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox67.CheckedChanged
        Check_radio3(66, 66)    ' Make it behave like a radio button.
        HandleP6B66()
    End Sub

    Private Sub CheckBox68_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox68.CheckedChanged
        Check_radio3(66, 67)    ' Make it behave like a radio button.
        HandleP6B66()
    End Sub

    Private Sub CheckBox69_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox69.CheckedChanged
        Check_radio3(66, 68)    ' Make it behave like a radio button.
        HandleP6B66()
    End Sub

    Private Sub HandleP6B66()
        Dim ItmPtr As Integer   'pointer to page item

        For ItmPtr = 0 To 2
            Form1.Record.P6Cb(ItmPtr) = Check1(ItmPtr + 66).Checked
        Next
    End Sub

    Private Sub TextBox30_TextChanged(sender As Object, e As EventArgs) Handles TextBox30.TextChanged
        Text_Changed6(29)
        If TextBox30.Text <> "" Then
            CheckBox70.Enabled = True
            CheckBox71.Enabled = True
            CheckBox72.Enabled = True
        Else
            CheckBox70.Enabled = False
            CheckBox71.Enabled = False
            CheckBox72.Enabled = False
        End If
    End Sub

    Private Sub CheckBox70_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox70.CheckedChanged
        Check_radio3(69, 69)
        HandleP6B69()
    End Sub

    Private Sub CheckBox71_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox71.CheckedChanged
        Check_radio3(69, 70)
        HandleP6B69()
    End Sub

    Private Sub CheckBox72_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox72.CheckedChanged
        Check_radio3(69, 71)
        HandleP6B69()
    End Sub

    Private Sub HandleP6B69()
        Dim ItmPtr As Integer
        For ItmPtr = 3 To 5
            Form1.Record.P6Cb(ItmPtr) = Check1(ItmPtr + 66).Checked
        Next
    End Sub

    Private Sub TextBox31_TextChanged(sender As Object, e As EventArgs) Handles TextBox31.TextChanged
        Text_Changed6(30)
        If TextBox31.Text <> "" Then
            CheckBox73.Enabled = True
            CheckBox74.Enabled = True
            CheckBox75.Enabled = True
        Else
            CheckBox73.Enabled = False
            CheckBox74.Enabled = False
            CheckBox75.Enabled = False
        End If
    End Sub

    Private Sub CheckBox73_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox73.CheckedChanged
        Check_radio3(72, 72)
        HandleP6B72()
    End Sub

    Private Sub CheckBox74_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox74.CheckedChanged
        Check_radio3(72, 73)
        HandleP6B72()
    End Sub

    Private Sub CheckBox75_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox75.CheckedChanged
        Check_radio3(72, 74)
        HandleP6B72()
    End Sub

    Private Sub HandleP6B72()
        Dim ItmPtr As Integer
        For ItmPtr = 6 To 8
            Form1.Record.P6Cb(ItmPtr) = Check1(ItmPtr + 66).Checked
        Next
    End Sub

    Private Sub TextBox32_TextChanged(sender As Object, e As EventArgs) Handles TextBox32.TextChanged
        Text_Changed6(31)
        If TextBox32.Text <> "" Then
            CheckBox76.Enabled = True
            CheckBox77.Enabled = True
            CheckBox78.Enabled = True
        Else
            CheckBox76.Enabled = False
            CheckBox77.Enabled = False
            CheckBox78.Enabled = False
        End If
    End Sub

    Private Sub TextBox33_TextChanged(sender As Object, e As EventArgs) Handles TextBox33.TextChanged
        Text_Changed6(32)
        If TextBox33.Text <> "" Then
            CheckBox79.Enabled = True
            CheckBox80.Enabled = True
            CheckBox81.Enabled = True
        Else
            CheckBox79.Enabled = False
            CheckBox80.Enabled = False
            CheckBox81.Enabled = False
        End If
    End Sub

    Private Sub CheckBox76_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox76.CheckedChanged
        Check_radio3(75, 75)
        HandleP6B75()
    End Sub

    Private Sub CheckBox77_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox77.CheckedChanged
        Check_radio3(75, 76)
        HandleP6B75()
    End Sub

    Private Sub CheckBox78_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox78.CheckedChanged
        Check_radio3(75, 77)
        HandleP6B75()
    End Sub

    Private Sub HandleP6B75()
        Dim ItmPtr As Integer

        For ItmPtr = 9 To 11
            Form1.Record.P6Cb(ItmPtr) = Check1(ItmPtr + 66).Checked
        Next
    End Sub

    Private Sub CheckBox79_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox79.CheckedChanged
        Check_radio3(78, 78)
        HandleP6B78()
    End Sub

    Private Sub CheckBox80_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox80.CheckedChanged
        Check_radio3(78, 79)
        HandleP6B78()
    End Sub

    Private Sub CheckBox81_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox81.CheckedChanged
        Check_radio3(78, 80)
        HandleP6B78()
    End Sub

    Private Sub HandleP6B78()
        Dim ItmPtr As Integer
        For ItmPtr = 12 To 14
            Form1.Record.P6Cb(ItmPtr) = Check1(ItmPtr + 66).Checked
        Next
    End Sub

    Private Sub Command22_Click(sender As Object, e As EventArgs) Handles Command22.Click   ' Tab 6 Continue
        TabControl1.SelectedTab = TabPage7
    End Sub

    Private Sub Command23_Click(sender As Object, e As EventArgs) Handles Command23.Click   ' Tab 6 Cancel
        Close()
    End Sub

    Private Sub TabPage7_Click(sender As Object, e As EventArgs) Handles TabPage7.Click
        Dim Idx As Integer  ' Check box index
        Dim msg As String

        Try
            Form1.ErrorFlag = 8019

            For Idx = 33 To 47
                Text1(Idx).BackColor = Form1.BackColor
                Text1(Idx).ForeColor = Form1.ForeColor
                Text1(Idx).Font = Form1.Font
            Next Idx

            '        For Idx = 26 To 40
            '        Check1(Idx).Font = Form1.Font
            '        Next Idx

            Form1.WriteQ7Screen()

            Visible = True
            Exit Sub

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            Close()
        End Try
    End Sub

    Private Sub TextBox34_TextChanged(sender As Object, e As EventArgs) Handles TextBox34.TextChanged
        Text_Changed7(33)
        If TextBox34.Text <> "" Then
            CheckBox82.Enabled = True
            CheckBox83.Enabled = True
            CheckBox84.Enabled = True
        Else
            CheckBox82.Enabled = False
            CheckBox83.Enabled = False
            CheckBox84.Enabled = False
        End If
    End Sub

    Private Sub CheckBox82_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox82.CheckedChanged
        Check_radio3(81, 81)
        HandleP7B81()
    End Sub

    Private Sub CheckBox83_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox83.CheckedChanged
        Check_radio3(81, 82)
        HandleP7B81()
    End Sub

    Private Sub CheckBox84_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox84.CheckedChanged
        Check_radio3(81, 83)
        HandleP7B81()
    End Sub

    Private Sub HandleP7B81()
        Dim Index As Integer
        For Index = 0 To 2
            Form1.Record.P7Cb(Index) = Check1(Index + 81).Checked
        Next
    End Sub

    Private Sub TextBox35_TextChanged(sender As Object, e As EventArgs) Handles TextBox35.TextChanged
        Text_Changed7(34)
        If TextBox35.Text <> "" Then
            CheckBox85.Enabled = True
            CheckBox86.Enabled = True
            CheckBox87.Enabled = True
        Else
            CheckBox85.Enabled = False
            CheckBox86.Enabled = False
            CheckBox87.Enabled = False
        End If
    End Sub

    Private Sub CheckBox85_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox85.CheckedChanged
        Check_radio3(84, 84)
        HandleP7B84()
    End Sub

    Private Sub CheckBox86_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox86.CheckedChanged
        Check_radio3(84, 85)
        HandleP7B84()
    End Sub

    Private Sub CheckBox87_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox87.CheckedChanged
        Check_radio3(84, 86)
        HandleP7B84()
    End Sub

    Private Sub HandleP7B84()
        Dim Index As Integer
        For Index = 3 To 5
            Form1.Record.P7Cb(Index) = Check1(Index + 81).Checked
        Next
    End Sub

    Private Sub TextBox36_TextChanged(sender As Object, e As EventArgs) Handles TextBox36.TextChanged
        Text_Changed7(35)
        If TextBox36.Text <> "" Then
            CheckBox88.Enabled = True
            CheckBox89.Enabled = True
            CheckBox90.Enabled = True
        Else
            CheckBox88.Enabled = False
            CheckBox89.Enabled = False
            CheckBox90.Enabled = False
        End If
    End Sub

    Private Sub CheckBox88_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox88.CheckedChanged
        Check_radio3(87, 87)
        HandleP7B87()
    End Sub

    Private Sub CheckBox89_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox89.CheckedChanged
        Check_radio3(87, 88)
        HandleP7B87()
    End Sub

    Private Sub CheckBox90_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox90.CheckedChanged
        Check_radio3(87, 89)
        HandleP7B87()
    End Sub

    Private Sub HandleP7B87()
        Dim Index As Integer
        For Index = 6 To 8
            Form1.Record.P7Cb(Index) = Check1(Index + 81).Checked
        Next
    End Sub

    Private Sub TextBox37_TextChanged(sender As Object, e As EventArgs) Handles TextBox37.TextChanged
        Text_Changed7(36)
        If TextBox37.Text <> "" Then
            CheckBox91.Enabled = True
            CheckBox92.Enabled = True
            CheckBox93.Enabled = True
        Else
            CheckBox91.Enabled = False
            CheckBox92.Enabled = False
            CheckBox93.Enabled = False
        End If
    End Sub

    Private Sub CheckBox91_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox91.CheckedChanged
        Check_radio3(90, 90)
        HandleP7B90()
    End Sub

    Private Sub CheckBox92_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox92.CheckedChanged
        Check_radio3(90, 91)
        HandleP7B90()
    End Sub

    Private Sub CheckBox93_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox93.CheckedChanged
        Check_radio3(90, 92)
        HandleP7B90()
    End Sub

    Private Sub HandleP7B90()
        Dim Index As Integer
        For Index = 9 To 11
            Form1.Record.P7Cb(Index) = Check1(Index + 81).Checked
        Next
    End Sub

    Private Sub TextBox38_TextChanged(sender As Object, e As EventArgs) Handles TextBox38.TextChanged
        Text_Changed7(37)
        If TextBox38.Text <> "" Then
            CheckBox94.Enabled = True
            CheckBox95.Enabled = True
            CheckBox96.Enabled = True
        Else
            CheckBox94.Enabled = False
            CheckBox95.Enabled = False
            CheckBox96.Enabled = False
        End If
    End Sub

    Private Sub CheckBox94_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox94.CheckedChanged
        Check_radio3(93, 93)
        HandleP7B93()
    End Sub

    Private Sub CheckBox95_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox95.CheckedChanged
        Check_radio3(93, 94)
        HandleP7B93()
    End Sub

    Private Sub CheckBox96_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox96.CheckedChanged
        Check_radio3(93, 95)
        HandleP7B93()
    End Sub

    Private Sub HandleP7B93()
        Dim Index As Integer
        For Index = 12 To 14
            Form1.Record.P7Cb(Index) = Check1(Index + 81).Checked
        Next
    End Sub

    Private Sub TextBox39_TextChanged(sender As Object, e As EventArgs) Handles TextBox39.TextChanged
        Text_Changed7(38)
        If TextBox39.Text <> "" Then
            CheckBox97.Enabled = True
            CheckBox98.Enabled = True
            CheckBox99.Enabled = True
        Else
            CheckBox97.Enabled = False
            CheckBox98.Enabled = False
            CheckBox99.Enabled = False
        End If
    End Sub

    Private Sub CheckBox97_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox97.CheckedChanged
        Check_radio3(96, 96)
        HandleP7B96()
    End Sub

    Private Sub CheckBox98_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox98.CheckedChanged
        Check_radio3(96, 97)
        HandleP7B96()
    End Sub

    Private Sub CheckBox99_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox99.CheckedChanged
        Check_radio3(96, 98)
        HandleP7B96()
    End Sub

    Private Sub HandleP7B96()
        Dim Index As Integer
        For Index = 15 To 17
            Form1.Record.P7Cb(Index) = Check1(Index + 81).Checked
        Next
    End Sub

    Private Sub TextBox40_TextChanged(sender As Object, e As EventArgs) Handles TextBox40.TextChanged
        Text_Changed7(39)
        If TextBox40.Text <> "" Then
            CheckBox102.Enabled = True
            CheckBox101.Enabled = True
            CheckBox100.Enabled = True
        Else
            CheckBox102.Enabled = False
            CheckBox101.Enabled = False
            CheckBox100.Enabled = False
        End If
    End Sub

    Private Sub CheckBox100_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox100.CheckedChanged
        Check_radio3(99, 99)
        HandleP7B99()
    End Sub

    Private Sub CheckBox101_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox101.CheckedChanged
        Check_radio3(99, 100)
        HandleP7B99()
    End Sub

    Private Sub CheckBox102_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox102.CheckedChanged
        Check_radio3(99, 101)
        HandleP7B99()
    End Sub

    Private Sub HandleP7B99()
        Dim Index As Integer
        For Index = 18 To 20
            Form1.Record.P7Cb(Index) = Check1(Index + 81).Checked
        Next
    End Sub

    Private Sub TextBox41_TextChanged(sender As Object, e As EventArgs) Handles TextBox41.TextChanged
        Text_Changed7(40)
        If TextBox41.Text <> "" Then
            CheckBox103.Enabled = True
            CheckBox104.Enabled = True
            CheckBox105.Enabled = True
        Else
            CheckBox103.Enabled = False
            CheckBox104.Enabled = False
            CheckBox105.Enabled = False
        End If
    End Sub

    Private Sub CheckBox103_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox103.CheckedChanged
        Check_radio3(102, 102)
        HandleP7B102()
    End Sub

    Private Sub CheckBox104_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox104.CheckedChanged
        Check_radio3(102, 103)
        HandleP7B102()
    End Sub

    Private Sub CheckBox105_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox105.CheckedChanged
        Check_radio3(102, 104)
        HandleP7B102()
    End Sub

    Private Sub HandleP7B102()
        Dim Index As Integer
        For Index = 21 To 23
            Form1.Record.P7Cb(Index) = Check1(Index + 81).Checked
        Next
    End Sub

    Private Sub TextBox42_TextChanged(sender As Object, e As EventArgs) Handles TextBox42.TextChanged
        Text_Changed7(41)
        If TextBox42.Text <> "" Then
            CheckBox106.Enabled = True
            CheckBox107.Enabled = True
            CheckBox108.Enabled = True
        Else
            CheckBox106.Enabled = False
            CheckBox107.Enabled = False
            CheckBox108.Enabled = False
        End If
    End Sub

    Private Sub CheckBox106_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox106.CheckedChanged
        Check_radio3(105, 105)
        HandleP7B105()
    End Sub

    Private Sub CheckBox107_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox107.CheckedChanged
        Check_radio3(105, 106)
        HandleP7B105()
    End Sub

    Private Sub CheckBox108_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox108.CheckedChanged
        Check_radio3(105, 107)
        HandleP7B105()
    End Sub

    Private Sub HandleP7B105()
        Dim Index As Integer
        For Index = 24 To 26
            Form1.Record.P7Cb(Index) = Check1(Index + 81).Checked
        Next
    End Sub

    Private Sub TextBox43_TextChanged(sender As Object, e As EventArgs) Handles TextBox43.TextChanged
        Text_Changed7(42)
        If TextBox43.Text <> "" Then
            CheckBox109.Enabled = True
            CheckBox110.Enabled = True
            CheckBox111.Enabled = True
        Else
            CheckBox109.Enabled = False
            CheckBox110.Enabled = False
            CheckBox111.Enabled = False
        End If
    End Sub

    Private Sub CheckBox109_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox109.CheckedChanged
        Check_radio3(108, 108)
        HandleP7B108()
    End Sub

    Private Sub CheckBox110_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox110.CheckedChanged
        Check_radio3(108, 109)
        HandleP7B108()
    End Sub

    Private Sub CheckBox111_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox111.CheckedChanged
        Check_radio3(108, 110)
        HandleP7B108()
    End Sub

    Private Sub HandleP7B108()
        Dim Index As Integer
        For Index = 27 To 29
            Form1.Record.P7Cb(Index) = Check1(Index + 81).Checked
        Next
    End Sub

    Private Sub TextBox44_TextChanged(sender As Object, e As EventArgs) Handles TextBox44.TextChanged
        Text_Changed7(43)
        If TextBox44.Text <> "" Then
            CheckBox112.Enabled = True
            CheckBox113.Enabled = True
            CheckBox114.Enabled = True
        Else
            CheckBox112.Enabled = False
            CheckBox113.Enabled = False
            CheckBox114.Enabled = False
        End If
    End Sub

    Private Sub CheckBox112_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox112.CheckedChanged
        Check_radio3(111, 111)
        HandleP7B111()
    End Sub

    Private Sub CheckBox113_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox113.CheckedChanged
        Check_radio3(111, 112)
        HandleP7B111()
    End Sub

    Private Sub CheckBox114_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox114.CheckedChanged
        Check_radio3(111, 113)
        HandleP7B111()
    End Sub

    Private Sub HandleP7B111()
        Dim Index As Integer
        For Index = 30 To 32
            Form1.Record.P7Cb(Index) = Check1(Index + 81).Checked
        Next
    End Sub

    Private Sub TextBox45_TextChanged(sender As Object, e As EventArgs) Handles TextBox45.TextChanged
        Text_Changed7(44)
        If TextBox45.Text <> "" Then
            CheckBox115.Enabled = True
            CheckBox116.Enabled = True
            CheckBox117.Enabled = True
        Else
            CheckBox115.Enabled = False
            CheckBox116.Enabled = False
            CheckBox117.Enabled = False
        End If
    End Sub

    Private Sub CheckBox115_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox115.CheckedChanged
        Check_radio3(114, 114)
        HandleP7B114()
    End Sub

    Private Sub CheckBox116_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox116.CheckedChanged
        Check_radio3(114, 115)
        HandleP7B114()
    End Sub

    Private Sub CheckBox117_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox117.CheckedChanged
        Check_radio3(114, 116)
        HandleP7B114()
    End Sub

    Private Sub HandleP7B114()
        Dim Index As Integer
        For Index = 33 To 35
            Form1.Record.P7Cb(Index) = Check1(Index + 81).Checked
        Next
    End Sub

    Private Sub TextBox46_TextChanged(sender As Object, e As EventArgs) Handles TextBox46.TextChanged
        Text_Changed7(45)
        If TextBox46.Text <> "" Then
            CheckBox118.Enabled = True
            CheckBox119.Enabled = True
            CheckBox120.Enabled = True
        Else
            CheckBox118.Enabled = False
            CheckBox119.Enabled = False
            CheckBox120.Enabled = False
        End If
    End Sub

    Private Sub CheckBox118_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox118.CheckedChanged
        Check_radio3(117, 117)
        HandleP7B117()
    End Sub

    Private Sub CheckBox119_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox119.CheckedChanged
        Check_radio3(117, 118)
        HandleP7B117()
    End Sub

    Private Sub CheckBox120_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox120.CheckedChanged
        Check_radio3(117, 119)
        HandleP7B117()
    End Sub

    Private Sub HandleP7B117()
        Dim Index As Integer
        For Index = 36 To 38
            Form1.Record.P7Cb(Index) = Check1(Index + 81).Checked
        Next
    End Sub

    Private Sub TextBox47_TextChanged(sender As Object, e As EventArgs) Handles TextBox47.TextChanged
        Text_Changed7(46)
        If TextBox47.Text <> "" Then
            CheckBox121.Enabled = True
            CheckBox122.Enabled = True
            CheckBox123.Enabled = True
        Else
            CheckBox121.Enabled = False
            CheckBox122.Enabled = False
            CheckBox123.Enabled = False
        End If
    End Sub

    Private Sub CheckBox121_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox121.CheckedChanged
        Check_radio3(120, 120)
        HandleP7B120()
    End Sub

    Private Sub CheckBox122_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox122.CheckedChanged
        Check_radio3(120, 121)
        HandleP7B120()
    End Sub

    Private Sub CheckBox123_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox123.CheckedChanged
        Check_radio3(120, 122)
        HandleP7B120()
    End Sub

    Private Sub HandleP7B120()
        Dim Index As Integer
        For Index = 39 To 41
            Form1.Record.P7Cb(Index) = Check1(Index + 81).Checked
        Next
    End Sub

    Private Sub TextBox48_TextChanged(sender As Object, e As EventArgs) Handles TextBox48.TextChanged
        Text_Changed7(47)
        If TextBox48.Text <> "" Then
            CheckBox124.Enabled = True
            CheckBox125.Enabled = True
            CheckBox126.Enabled = True
        Else
            CheckBox124.Enabled = False
            CheckBox125.Enabled = False
            CheckBox126.Enabled = False
        End If
    End Sub

    Private Sub CheckBox124_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox124.CheckedChanged
        Check_radio3(123, 123)
        HandleP7B123()
    End Sub

    Private Sub CheckBox125_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox125.CheckedChanged
        Check_radio3(123, 124)
        HandleP7B123()
    End Sub

    Private Sub CheckBox126_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox126.CheckedChanged
        Check_radio3(123, 125)
        HandleP7B123()
    End Sub

    Private Sub HandleP7B123()
        Dim Index As Integer
        For Index = 42 To 44
            Form1.Record.P7Cb(Index) = Check1(Index + 81).Checked
        Next
    End Sub

    Private Sub Command24_Click(sender As Object, e As EventArgs) Handles Command24.Click   ' Tab 7 Find Symptoms
        If QuestProgress.Visible Then QuestProgress.Close()   ' Need to re-initialize in case it is open.
        QuestProgress.Show()
    End Sub

    Private Sub Command25_Click(sender As Object, e As EventArgs) Handles Command25.Click   ' Tab 7 Cancel
        Close()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Form1.QDirty = True
        Form1.Record.PDate = DateTimePicker1.Value
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        Form1.QDirty = True
        Form1.Record.DOB = DateTimePicker2.Value
    End Sub

    Private Sub DateTimePicker3_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker3.ValueChanged
        Form1.QDirty = True
        Form1.Symptom.SDate(Form1.SymPtr) = DateTimePicker3.Value
    End Sub

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Private Sub AddSymptomButton_Click(sender As Object, e As EventArgs) Handles AddSymptomButton.Click
        Dim Ptr As Integer  ' Array element pointer

        Form1.SympSize += 1
        ReDim Preserve Form1.Symptom.SDate(Form1.SympSize - 1)
        ReDim Preserve Form1.Symptom.Title(Form1.SympSize - 1)
        ReDim Preserve Form1.Symptom.P1(6, Form1.SympSize - 1)
        ReDim Preserve Form1.Symptom.P1Cb(4, Form1.SympSize - 1)
        ReDim Preserve Form1.Symptom.P2(6, Form1.SympSize - 1)
        ReDim Preserve Form1.Symptom.P3(1, Form1.SympSize - 1)
        ReDim Preserve Form1.Symptom.P3Cb(49, Form1.SympSize - 1)
        ReDim Preserve Form1.Symptom.P4Cb(4, Form1.SympSize - 1)
        ReDim Preserve Form1.Symptom.P4(1, Form1.SympSize - 1)
        Form1.SymPtr += 1

        ' Initialize new symptom array element.
        For Ptr = 0 To 6
            Form1.Symptom.P1(Ptr, Form1.SymPtr) = ""
        Next
        For Ptr = 0 To 4
            Form1.Symptom.P1Cb(Ptr, Form1.SymPtr) = False
        Next
        For Ptr = 0 To 6
            Form1.Symptom.P2(Ptr, Form1.SymPtr) = ""
        Next
        For Ptr = 0 To 1
            Form1.Symptom.P3(Ptr, Form1.SymPtr) = ""
        Next
        For Ptr = 0 To 49
            Form1.SympData.P3Cb(Ptr, Form1.SymPtr) = False
        Next
        For Ptr = 0 To 4
            Form1.Symptom.P4Cb(Ptr, Form1.SymPtr) = False
        Next
        For Ptr = 0 To 1
            Form1.Symptom.P4(Ptr, Form1.SymPtr) = ""
        Next

        ' Write the initialized symptom array to the tab forms.
        Form1.WriteQ1Screen()
        Form1.WriteQ2Screen()
        Form1.WriteQ3Screen()
        Form1.WriteQ4Screen()

        ' Display first page for entering next symptom.
        TabControl1.SelectedTab = TabPage1
    End Sub

    Public Sub EnableNextSymptom()
        Command6.Enabled = True ' Page 1
        Command11.Enabled = True    ' Page 2
        Command15.Enabled = True    ' Page 3
        Command19.Enabled = True    ' Page 4
    End Sub

    Public Sub DisableNextSymptom()
        Command6.Enabled = False ' Page 1
        Command11.Enabled = False    ' Page 2
        Command15.Enabled = False    ' Page 3
        Command19.Enabled = False    ' Page 4
    End Sub

    Public Sub EnablePrevSymptom()
        Command7.Enabled = True ' Page 1
        Command10.Enabled = True    ' Page 2
        Command14.Enabled = True    ' Page 3
        Command18.Enabled = True    ' Page 4
    End Sub

    Public Sub DisablePrevSymptom()
        Command7.Enabled = False ' Page 1
        Command10.Enabled = False    ' Page 2
        Command14.Enabled = False    ' Page 3
        Command18.Enabled = False    ' Page 4
    End Sub

    Private Sub TextBox2_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Form1.QDirty = True
        Form1.Symptom.Title(Form1.SymPtr) = Text1(1).Text
        Form1.EnableFileQuestSave()
    End Sub


    Private Sub mnuPrintPreviewQuestionnaire_Click(sender As Object, e As EventArgs)

        Form1.stringBuffer = ""   ' Clear out any previous print job

        PrintPreviewDialog1.Document = Form1.PrintDocument1
        Form1.QuestPrint()  ' Load the print buffer
        ' Get printer font from registry
        Form1.PrinterFontName = New Font(GetSetting("Hahnasst", "Startup", "PrinterFontName", [Default]:="Microsoft Sans Serif"),
            GetSetting("Hahnasst", "Startup", "PrinterFontSize", [Default]:="8"))
        PrintPreviewDialog1.ShowIcon = False
        PrintPreviewDialog1.WindowState = FormWindowState.Normal
        PrintPreviewDialog1.StartPosition = FormStartPosition.CenterScreen
        PrintPreviewDialog1.ClientSize = New Size(800, 1200)
        PrintPreviewDialog1.UseAntiAlias = True
        PrintPreviewDialog1.Document = Form1.PrintDocument1
        PrintPreviewDialog1.ShowDialog()
        PrintPreviewDialog1.Close()
        PrintPreviewDialog1.Dispose()
    End Sub

    Private Sub PrintPreviewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintPreviewToolStripMenuItem.Click

        Form1.stringBuffer = ""   ' Clear out any previous print job

        PrintPreviewDialog1.Document = Form1.PrintDocument1
        Form1.QuestPrint()  ' Load the print buffer
        ' Get printer font from registry
        Form1.PrinterFontName = New Font(GetSetting("Hahnasst", "Startup", "PrinterFontName", [Default]:="Microsoft Sans Serif"),
            GetSetting("Hahnasst", "Startup", "PrinterFontSize", [Default]:="8"))
        PrintPreviewDialog1.ShowIcon = False
        PrintPreviewDialog1.WindowState = FormWindowState.Normal
        PrintPreviewDialog1.StartPosition = FormStartPosition.CenterScreen
        PrintPreviewDialog1.ClientSize = New Size(800, 1200)
        PrintPreviewDialog1.UseAntiAlias = True
        PrintPreviewDialog1.Document = Form1.PrintDocument1
        PrintPreviewDialog1.ShowDialog()
        PrintPreviewDialog1.Close()
        PrintPreviewDialog1.Dispose()
    End Sub

    Private Sub PrintToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintToolStripMenuItem.Click
        Dim pdResult As DialogResult

        Form1.stringBuffer = ""   ' Clear out any previous print job
        HandlePrint()   ' Load the print buffer

        ' Allow the user to choose the page range he or she would
        ' like to print.
        PrintDialog1.AllowSomePages = True

        PrintDialog1.AllowCurrentPage = True
        PrintDialog1.AllowPrintToFile = True
        PrintDialog1.AllowSelection = True

        ' Show the help button.
        PrintDialog1.ShowHelp = True


        PrintDialog1.Document = Form1.PrintDocument1

        pdResult = PrintDialog1.ShowDialog()
        If pdResult = vbOK Then
            Form1.SetupAndPrint()
            Form1.PrintDocument1.Print()  ' Perform print' Perform print
        End If
        PrintDialog1.Dispose()
    End Sub

    Private Sub mnuQuestHelp_Click(sender As Object, e As EventArgs) Handles mnuQuestHelp.Click
        Dim msg As String
        Dim myProcess As New Process

        Try
            myProcess.StartInfo.FileName = AppContext.BaseDirectory + Form1.HELP_DIR + "Questionnaire.htm"
            myProcess.StartInfo.UseShellExecute = True
            myProcess.StartInfo.RedirectStandardOutput = False
            myProcess.Start()
            myProcess.Dispose()
        Catch
            msg = "Help file not found.  Please reinstall software to fix this problem."
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub
End Class