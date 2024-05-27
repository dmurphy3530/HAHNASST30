Public Class PrintSelections
    Private Sub PrintSymp_CheckedChanged(sender As Object, e As EventArgs) Handles PrintSymp.CheckedChanged
        If PrintSymp.Checked = True Then
            Form1.SympPrint = True
        Else
            Form1.SympPrint = False
        End If
    End Sub

    Private Sub PrintPresc_CheckedChanged(sender As Object, e As EventArgs) Handles PrintPresc.CheckedChanged
        If PrintPresc.Checked = True Then
            Form1.CasePrint = True
        Else
            Form1.CasePrint = False
        End If
    End Sub

    Private Sub PrintRem_CheckedChanged(sender As Object, e As EventArgs) Handles PrintRem.CheckedChanged
        If PrintRem.Checked = True Then
            Form1.RemPrint = True
        Else
            Form1.RemPrint = False
        End If
    End Sub

    Private Sub PrintQuest_CheckedChanged(sender As Object, e As EventArgs) Handles PrintQuest.CheckedChanged
        If PrintQuest.Checked = True Then
            Form1.QPrint = True
        Else
            Form1.QPrint = False
        End If
    End Sub

    Private Sub PrintAll_CheckedChanged(sender As Object, e As EventArgs) Handles PrintAll.CheckedChanged
        If PrintAll.Checked = True Then
            If PrintSymp.Enabled Then PrintSymp.Checked = True
            If PrintPresc.Enabled Then PrintPresc.Checked = True
            If PrintRem.Enabled Then PrintRem.Checked = True
            If PrintQuest.Enabled Then PrintQuest.Checked = True
        End If
    End Sub

    Private Sub PrintButton_Click(sender As Object, e As EventArgs) Handles PrintButton.Click
        ' Print according to the following flags from PrintSelections dialog:
        ' SympPrint = Print list of selected symptoms in MustBox and SelSymp list box
        ' CasePrint = Print list of remedies in PrescRem ListBox1
        ' RemPrint = Print description of each remedy in PrescRem ListBox1
        ' QPrint = Print questionnaire

        Form1.stringBuffer = ""   ' Clear out any previous print job
        HandlePrint()   ' Load print buffer and print
    End Sub

    Public Sub HandlePrint()
        ' Load print buffer and print

        Dim Ptr As Integer ' Points to PrescRem ListBox1 items
        Dim ErrorFlag As Integer    ' error message number
        Dim Largest As Integer  ' sort parameter
        Dim LID As Integer
        Dim LinePtr As Integer  ' points to matmed file line
        Dim ListDat As String   ' selected symptom data line
        Dim ListPtr As Integer  ' list box item pointer
        Dim ListSize As Integer ' # of items in select list
        Dim LocSympCount As Integer
        Dim LocSeqSympNo() As Integer  ' local sequential symptom numbers
        Dim LocSeqSympNo2() As Integer ' local sequential symptom numberss minus duplicates
        Dim msg As String       ' message dialog string
        Dim OutLine As String   ' Text to output
        Dim pdResult As DialogResult
        Dim Ptr1, Ptr2 As Integer   ' sort pointers
        Dim RemCtr As Integer   ' remedy counter for finding remedy in matmed file
        Dim StringRemID As String   ' remedy ID from list box string
        Dim StringSeqSympNo As String  ' String form of sequential symptom number from list box string
        Dim StrPtr As Integer   ' string pointer
        Dim SympInList As Boolean   ' symptom is in must box or selected symptom list if true
        Dim SymptomNumber As Integer    ' sequential number of symptom in MatMed remedy text
        Dim T1 As Integer               ' sort item temporary variable
        Dim strRemData As String   ' Remedy ID string from PrescRem list box

        If Form1.QPrint Then
            Questionnaire.HandlePrint()    ' Print questionnaire
            Form1.QuestPrint()    ' Print questionnaire
            If Form1.SympPrint Or Form1.CasePrint Or Form1.RemPrint Then Form1.stringBuffer += (vbCrLf + vbLf) ' If more to print, add a separater.
        End If
        If Form1.SympPrint Then
            Form1.FilePrint()   ' Print selected symptoms list
            If Form1.CasePrint Or Form1.RemPrint Then Form1.stringBuffer += (vbCrLf + vbLf) ' If more to print, add a separater.
        End If
        If Form1.CasePrint Then
            PrescRem.HandlePrint()  ' Print remedy list
            If Form1.RemPrint Then Form1.stringBuffer += (vbCrLf + vbLf) ' If more to print, add a separater.
        End If
        If Form1.RemPrint Then  ' Use main code from DispRem Print Handler for each remedy, except output to print buffer instead of text box
            'Print the page title.
            Form1.stringBuffer += "Remedies:" + vbCrLf
            For Ptr = 0 To PrescRem.ListBox1.Items.Count - 1
                strRemData = PrescRem.ListBox1.Items(Ptr).SelRemData.ToString + vbCrLf
                Form1.RemedyNum = Val(strRemData)
                ErrorFlag = 12001
                ReDim LocSeqSympNo(0)
                ReDim LocSeqSympNo2(0)

                Try
                    ErrorFlag = 12002
                    LocSympCount = 0
                    ' Add a ">>>" in front of any symptoms selected for the current case.
                    ' First, get a list of all selected symptoms, beginning with the "must box".
                    If Form1.MustBox.Text <> "" Then
                        ListDat = Form1.MustData
                        StrPtr = InStr(ListDat, ",") + 1    ' Skip symptom ID

                        While (StrPtr < Len(ListDat) And Not SympInList)
                            StringRemID = Mid(ListDat, StrPtr)
                            ListDat = ListDat.Substring(StrPtr)
                            StrPtr = InStr(ListDat, ",") + 1
                            If Val(StringRemID) = Form1.RemedyNum + 1 Then    ' Save sequential symptom numbers.
                                StringSeqSympNo = Mid(ListDat, StrPtr)
                                If LocSympCount + 1 > UBound(LocSeqSympNo) Then _
                                        ReDim Preserve LocSeqSympNo(LocSympCount + 1)

                                LocSeqSympNo(LocSympCount) = Val(StringSeqSympNo)
                                LocSympCount = LocSympCount + 1
                            End If
                            ListDat = ListDat.Substring(StrPtr)
                            If InStr(ListDat, ",") <> 0 Then
                                StrPtr = InStr(ListDat, ",") + 1    ' Point to next number
                            Else
                                StrPtr = Len(ListDat)
                            End If
                        End While
                    End If

                    'Continue by finding all symptoms for remedy in prescribe list.
                    ListSize = Form1.SelLst.Items.Count
                    For ListPtr = 0 To ListSize - 1
                        ListDat = Form1.CombSympData(Form1.Lst1.SelectedIndices(ListPtr)).SympData.ToString()
                        StrPtr = InStr(ListDat, ",") + 1    ' Skip symptom ID
                        While (StrPtr < Len(ListDat) And Not SympInList)
                            StringRemID = Mid(ListDat, StrPtr)
                            ListDat = ListDat.Substring(StrPtr)
                            StrPtr = InStr(ListDat, ",") + 1
                            If Val(StringRemID) = Form1.RemedyNum + 1 Then    ' Save sequential symptom numbers.
                                StringSeqSympNo = Mid(ListDat, StrPtr)
                                If LocSympCount + 1 > UBound(LocSeqSympNo) Then _
                                        ReDim Preserve LocSeqSympNo(LocSympCount + 1)

                                LocSeqSympNo(LocSympCount) = Val(StringSeqSympNo)
                                LocSympCount = LocSympCount + 1
                            End If
                            ListDat = ListDat.Substring(StrPtr)
                            If InStr(ListDat, ",") <> 0 Then
                                StrPtr = InStr(ListDat, ",") + 1    ' Point to next number
                            Else
                                StrPtr = Len(ListDat)
                            End If
                        End While
                    Next ListPtr

                    ' Sort LocSympID and eliminate duplicates.
                    Largest = -1
                    For Ptr1 = 0 To LocSympCount - 1
                        For Ptr2 = Ptr1 To LocSympCount - 1
                            If (LocSeqSympNo(Ptr2) >= Largest) Then
                                Largest = LocSeqSympNo(Ptr2)
                                LID = Ptr2
                            End If
                        Next Ptr2
                        T1 = LocSeqSympNo(Ptr1)
                        LocSeqSympNo(Ptr1) = LocSeqSympNo(LID)
                        LocSeqSympNo(LID) = T1
                    Next Ptr1
                    ReDim LocSeqSympNo2(UBound(LocSeqSympNo))
                    Ptr2 = 1
                    LocSeqSympNo2(0) = LocSeqSympNo(0)
                    For Ptr1 = 1 To LocSympCount - 1
                        If LocSeqSympNo(Ptr1) <> LocSeqSympNo(Ptr1 - 1) Then
                            LocSeqSympNo2(Ptr2) = LocSeqSympNo(Ptr1)
                            Ptr2 = Ptr2 + 1
                        End If
                    Next Ptr1
                    LocSympCount = Ptr2

                    ErrorFlag = 12004
                    SymptomNumber = 0

                    ErrorFlag = 12006
                    LinePtr = 1 ' Point past first blank line
                    For RemCtr = 1 To Form1.RemedyNum
                        While SelRem.MatMedText(LinePtr) <> ""
                            LinePtr += 1
                        End While
                        LinePtr += 1    ' Point past the blank line
                    Next
                    ' Should now be pointing to selected remedy; add remedy text to List1
                    While SelRem.MatMedText(LinePtr) <> ""
                        If (InStr(Mid(SelRem.MatMedText(LinePtr), 1, 1), Chr(9)) > 0) Then    ' if a symptom
                            SympInList = False
                            SymptomNumber = SymptomNumber + 1
                            Ptr1 = 0
                            While Not SympInList And Ptr1 < LocSympCount
                                If LocSeqSympNo2(Ptr1) = SymptomNumber Then SympInList = True
                                Ptr1 = Ptr1 + 1
                            End While
                            ErrorFlag = 12009

                            If (SympInList) Then
                                OutLine = ">>>" + SelRem.MatMedText(LinePtr)
                            Else
                                OutLine = SelRem.MatMedText(LinePtr)
                            End If
                        Else
                            OutLine = SelRem.MatMedText(LinePtr)
                        End If
                        Form1.stringBuffer += OutLine + vbCrLf
                        LinePtr += 1
                        If LinePtr = SelRem.MatMedText.Length Then Exit While
                    End While
                    Cursor = Cursors.Default    ' Change cursor to default.

                Catch
                    Cursor = Cursors.Default    ' Change cursor to default.
                    msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
                    Dim unused = MsgBox(Prompt:=msg,
                            Buttons:=vbOKOnly + vbCritical,
                            Title:="ERROR!")
                End Try
            Next Ptr
        End If

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
            Form1.PrintPreviewDialog1.Close()
            Form1.PrintPreviewDialog1.Dispose()
        Else
            ' Allow the user to choose the page range he or she would
            ' like to print.
            Form1.PrintDialog1.AllowSomePages = True

            Form1.PrintDialog1.AllowCurrentPage = False
            Form1.PrintDialog1.AllowPrintToFile = True
            Form1.PrintDialog1.AllowSelection = False

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
        Close()
    End Sub

    Private Sub CancelPButton_Click(sender As Object, e As EventArgs) Handles CancelPButton.Click
        Close()
    End Sub

    Private Sub PrintSelections_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Form1.SympPrint = True
        Form1.CasePrint = False
        Form1.RemPrint = False
        Form1.QPrint = False
        If Form1.SelLst.Items.Count > 0 Then
            PrintSymp.Visible = True
            PrintSymp.Enabled = True
            PrintSymp.Checked = True    ' Assume we are printing at least the symptom list; user can un-checked this if desired.
        Else
            PrintSymp.Visible = False
            PrintSymp.Enabled = False
        End If
        If PrescRem.Visible = True Then
            PrintPresc.Visible = True
            PrintPresc.Enabled = True
            PrintRem.Visible = True
            PrintRem.Enabled = True
        Else
            PrintPresc.Visible = False
            PrintPresc.Enabled = False
            PrintRem.Visible = False
            PrintRem.Enabled = False
        End If
        If Questionnaire.TextBox1.Text <> "" Or Questionnaire.TextBox2.Text <> "" Then
            PrintQuest.Visible = True
            PrintQuest.Enabled = True
        Else
            PrintQuest.Visible = False
            PrintQuest.Enabled = False
        End If
    End Sub
End Class