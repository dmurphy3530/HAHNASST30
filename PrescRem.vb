Option Explicit On
Public Class PrescRem
    Private Sub PrescRem_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim ErrorFlag As Integer    ' error number
        Dim msg As String       ' error message string

        Try
            ErrorFlag = 7001
            ListBox1.BackColor = Form1.Lst1.BackColor
            ListBox1.ForeColor = Form1.Lst1.ForeColor
            ListBox1.Font = Form1.Lst1.Font

            If Form1.SelectionMode = 1 Then
                Me.Text = Me.Text + ":  Default remedy weighting"
            ElseIf Form1.SelectionMode = 2 Then
                Me.Text = Me.Text + ":  Equal remedy weighting"
            ElseIf Form1.SelectionMode = 3 Then
                Me.Text = Me.Text + ":  Straight remedy symptom count"
            End If
            If Form1.Normalize Then
                Me.Text = Me.Text + ", Normalization On"
            Else
                Me.Text = Me.Text + ", Normalization Off"
            End If

            Command1.Enabled = False

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            Close()
        End Try
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Command1.Enabled = True
    End Sub

    Private Sub ListBox1_DblClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        Dim msg As String   ' error message string

        Try
            Form1.ErrorFlag = 7003
            Command1.Enabled = True
            Call Command1_Click(sender, e)

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub Command1_Click(sender As Object, e As EventArgs) Handles Command1.Click ' Show Remedy
        Dim ErrorFlag As Integer    ' error number
        Dim msg As String   ' error message string

        Try
            ErrorFlag = 7004

            If ListBox1.Items.Count > 0 Then
                If DispRem.Visible Then DispRem.Close() ' Need to re-initialize in case it is open.
                Form1.RemedyNum = Val(Form1.SelRemData(ListBox1.SelectedIndex).SelRemData.ToString) - 1 'Note DispRem List1 is 0-based
                DispRem.Left = 270
                DispRem.Top = 1620
                DispRem.Show()
            End If

        Catch
            Cursor = Cursors.Default    ' Change cursor to default
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
               Buttons:=vbOKOnly + vbCritical,
               Title:="ERROR!")
        End Try
    End Sub

    Private Sub Command4_Click(sender As Object, e As EventArgs) Handles Command4.Click     ' Export Presc.
        Dim Filedat As String   ' line of data for presc. file
        Dim FileStatus As Boolean   ' true = good
        Dim msg As String       ' dialog message string
        Dim PrescFileName As String   ' name of presc. file
        Dim PrescPtr As Integer     ' presc. list pointer
        Dim ptr As String           ' string pointer
        Dim tempFileName As String  ' file name to display in SaveFileDialog
        Dim WrFile As Boolean   ' Write file permissive

        Try
            Form1.ErrorFlag = 7005
            ' Set Filters.
            Form1.SaveFileDialog1.Filter = "Presc. Files (*.PRE) |*.PRE| All Files (*.*) | *.*"
            ' Specify default filter.
            Form1.SaveFileDialog1.FilterIndex = 1
            WrFile = False
            ' Display the Save dialog box.
            Form1.SaveFileDialog1.InitialDirectory = Form1.CaseFilePath
            tempFileName = Form1.CaseFileName
            ' Strip file type from tempFileName
            ptr = InStr(tempFileName, ".")
            If ptr <> 0 Then
                tempFileName = Strings.Left(tempFileName, ptr - 1)
            End If
            Form1.SaveFileDialog1.FileName = tempFileName
            If (Form1.SaveFileDialog1.ShowDialog() = vbOK) Then

                ' Call the save file procedure.
                If Form1.SaveFileDialog1.FileName <> "" Then    'get the file information
                    Form1.ErrorFlag = 7006
                    PrescFileName = Form1.SaveFileDialog1.FileName

                    Form1.ErrorFlag = 7007
                    Filedat = " "   ' Avoid warning variable used before assigning value
                    FileStatus = Form1.OpenFileWrite(PrescFileName, Form1.PrescFile)
                    If Not FileStatus Then Throw New ArgumentException("Exception Occured")

                    Form1.ErrorFlag = 7008
                    Seek(Form1.PrescFile, 1)
                    Form1.ErrorFlag = 7009
                    If Form1.SelectionMode = 1 Then
                        Filedat = "Default remedy weighting"
                    ElseIf Form1.SelectionMode = 2 Then
                        Filedat = "Equal remedy weighting"
                    ElseIf Form1.SelectionMode = 3 Then
                        Filedat = "Straight remedy symptom count"
                    End If
                    If Form1.Normalize Then
                        Filedat += ", Normalization On" + vbCrLf
                    Else
                        Filedat += ", Normalization Off" + vbCrLf
                    End If
                    Form1.ErrorFlag = 7010
                    Print(Form1.PrescFile, Filedat)

                    For PrescPtr = 0 To ListBox1.Items.Count - 1
                        Filedat = vbCrLf + ListBox1.Items(PrescPtr).SelRemDesc.ToString
                        Print(Form1.PrescFile, Filedat)
                    Next PrescPtr
                    FileSystem.FileClose(Form1.PrescFile)

                End If
            End If

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            Form1.CloseAll()
        End Try
    End Sub

    Private Sub Command2_Click(sender As Object, e As EventArgs) Handles Command2.Click     ' Print Presc.
        Form1.stringBuffer = ""   ' Clear out any previous print job

        Dim pdResult As DialogResult

        Form1.stringBuffer = ""   ' Clear out any previous print job
        ' Allow the user to choose the page range he or she would
        ' like to print.
        Form1.PrintDialog1.AllowSomePages = True

        Form1.PrintDialog1.AllowCurrentPage = False
        Form1.PrintDialog1.AllowPrintToFile = True
        Form1.PrintDialog1.AllowSelection = False

        ' Show the help button.
        Form1.PrintDialog1.ShowHelp = True


        Form1.PrintDialog1.Document = Form1.PrintDocument1

        pdResult = Form1.PrintDialog1.ShowDialog()
        If pdResult = vbOK Then
            HandlePrint()   ' Load print buffer

            Form1.SetupAndPrint()
            Form1.PrintDocument1.Print()
        End If
        Form1.PrintDialog1.Dispose()
    End Sub

    Public Sub HandlePrint()
        ' Load the print buffer

        Dim Ch As String
        Dim ErrorFlag As Integer    ' error number
        Dim ListSize As Integer     ' # of lines in list box
        Dim msg As String           ' Error message string
        Dim Ptr As Integer          ' loop pointer
        Dim S As String             ' line from list box
        Dim tabPtr As Integer       ' position of first vbTab
        Dim tempStr As String       ' temp string

        Try
            ErrorFlag = 7011
            S = " "   ' Avoid warning variable used before assigning value

            Form1.PrintDialog1.Document = Form1.PrintDocument1

            '   Print the page title.
            Ch = "Case: " + Form1.CaseFileName + vbCrLf
            Form1.stringBuffer += Ch

            ListSize = ListBox1.Items.Count

            Form1.ErrorFlag = 7012
            If Form1.SelectionMode = 1 Then
                S = "Default remedy weighting"
            ElseIf Form1.SelectionMode = 2 Then
                S = "Equal remedy weighting"
            ElseIf Form1.SelectionMode = 3 Then
                S = "Straight remedy symptom count"
            End If
            If Form1.Normalize Then
                S = S + ", Normalization On"
            Else
                S = S + ", Normalization Off"
            End If
            Form1.stringBuffer += S + vbCrLf + vbLf
            For Ptr = 0 To ListSize - 1
                tempStr = ListBox1.Items(Ptr).SelRemDesc.ToString + vbCrLf
                tabPtr = InStr(tempStr, vbTab)
                If tabPtr >= 6 Then ' need to remove one vbTab so columns will come out even
                    tempStr = Strings.Left(tempStr, 5) + Mid(tempStr, 7)
                End If
                Form1.stringBuffer += tempStr + vbCrLf
            Next Ptr

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub Command3_Click(sender As Object, e As EventArgs) Handles Command3.Click     ' Done
        Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim msg As String
        Dim myProcess As New Process
        Try
            myProcess.StartInfo.FileName = AppContext.BaseDirectory + Form1.HELP_DIR + "PrescribeForm.htm"
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
End Class