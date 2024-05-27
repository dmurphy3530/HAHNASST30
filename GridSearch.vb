Option Explicit On
Public Class GridSearch
    Public CompText As String   ' text to compare
    Public First As Integer    ' top of binary search
    Public Text1(24) As TextBox
    Public SymptomWords(26771, 33) As Integer    ' symptom ID word lists
    Public NumWords(26771) As Integer  ' Number of word indexes in each symptom
    Public GridLocLineDat As String ' array of grid index strings
    Public Last As Integer     ' bottom of binary search
    Public locLineDat As String     ' selected list item
    Public NumElements(24) As Long   ' # of elements in GIn() arrays
    Public SkipChange As Boolean   ' skips text change routine
    Public TextData(24) As String  ' Data from List1/2 element associated with word in Text1(index)
    Public TextIndex As Long        ' textbox array element pointer
    Public gridSearData As New ArrayList()  ' List1 items
    Public grid2SearData As New ArrayList() ' List2 items
    Public idxLineDat As String ' index string line data
    Public SupTxtChgHnd As Boolean  ' used to suppress action of text change handler if software updated text box
    Public textHasChanged As Boolean   ' used to suppress action of list box handler if user is updating text box
    Public listLength(24) As Integer    ' # of items in list per text box index
    Public idxArray() As Integer    ' array of all indexes in idxLineDat
    Public NumIndexes As Integer    ' number of symptom IDs in idxArray
    Public idxPtr As Integer    ' pointer to indexArray element

    Private Sub Command3_Click(sender As Object, e As EventArgs) Handles Command3.Click
        Dim msg As String
        Dim myProcess As New Process

        Try
            myProcess.StartInfo.FileName = AppContext.BaseDirectory + Form1.HELP_DIR + "GridSearchForm.htm"
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

    Private Sub GridSearch_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim ErrorFlag As Integer    ' error number
        Dim msg As String
        Dim Ptr As Integer  ' loop counter
        Dim RecLen As Integer   ' search index index file record length

        Try
            ErrorFlag = 5001

            ' Set up pointers to the text boxes so they can be indexed easily.
            Text1(0) = TextA1
            Text1(1) = TextB1
            Text1(2) = TextC1
            Text1(3) = TextD1
            Text1(4) = TextA2
            Text1(5) = TextB2
            Text1(6) = TextC2
            Text1(7) = TextD2
            Text1(8) = TextA3
            Text1(9) = TextB3
            Text1(10) = TextC3
            Text1(11) = TextD3
            Text1(12) = TextA4
            Text1(13) = TextB4
            Text1(14) = TextC4
            Text1(15) = TextD4
            Text1(16) = TextA5
            Text1(17) = TextB5
            Text1(18) = TextC5
            Text1(19) = TextD5
            Text1(20) = TextA6
            Text1(21) = TextB6
            Text1(22) = TextC6
            Text1(23) = TextD6

            '            Form1.EditFind.Enabled = False

            SkipChange = False  ' Allows text change to run.
            FindNext.Enabled = False
            Command2.Enabled = True     ' Cancel button
            FindFirst.Enabled = False
            FindPrev.Enabled = False
            FindAllButton.Enabled = False

            ' Set colors derived from registry in Form1; only enable TextA1 initially.
            For Ptr = 0 To 23
                Text1(Ptr).BackColor = Form1.TextBox1.BackColor
                Text1(Ptr).ForeColor = Form1.TextBox1.ForeColor
                Text1(Ptr).Font = Form1.TextBox1.Font
                Text1(Ptr).Enabled = False
            Next Ptr
            TextA1.Enabled = True
            FindNext.Enabled = False
            FindFirst.Enabled = False
            FindPrev.Enabled = False
            FindAllButton.Enabled = False
            GridLocLineDat = ""
            List1.BackColor = Form1.TextBox1.BackColor
            List1.ForeColor = Form1.TextBox1.ForeColor
            List1.Font = Form1.TextBox1.Font

            Form1.ErrorFlag = 5002
            textHasChanged = True  ' Force it to skip List1_SelectedIndexChanged code when loading
            TextA1.Select()
            Command2.Text = "Cancel"
            RecLen = 8
            Form1.ErrorFlag = 5003

            gridSearData = Find.searData    ' Load List1 with all the words (file seartxtidx.dat); First line is word, second line is pointers to words in symptextdata.dat.
            List1.DataSource = Nothing ' Remove and re-associate database in order to refresh List1
            List1.DataSource = gridSearData
            List1.DisplayMember = "SearDataName"
            List1.ValueMember = "SearIdxName"
            List1.ClearSelected()
            List1.Visible = False   ' List1 is pre-loaded and used for text boxes in the first column
            List2.Visible = False   ' List2 is used for text boxes in columns to the right of the first column
            LoadSympWordPtrs()  ' Loads SymptomWords(<symptom pointer>, <word list>) and NumWords()

            TextA1.Select()
            Text1_Click(0)  ' Pop up List1 and allow user to select a word.
            Exit Sub

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            Form1.EditFind.Enabled = True
            Close()
        End Try
    End Sub

    Private Sub GridSearch_Unload(Cancel As Integer)
        Dim ErrorFlag As Integer    ' error number
        Dim msg As String   ' Error message string
        Dim Ptr As Integer

        Try
            ErrorFlag = 5004

            ' Save strings.
            For Ptr = 0 To 23
                Form1.GridString(Ptr) = Text1(Ptr).Text
            Next Ptr
            Form1.GridString(0) = TextA1.Text
            Form1.GridString(1) = TextB1.Text
            Form1.GridString(2) = TextC1.Text
            Form1.GridString(3) = TextD1.Text
            Form1.GridString(5) = TextB2.Text
            Form1.GridString(6) = TextC2.Text
            Form1.GridString(7) = TextD2.Text
            Form1.GridString(9) = TextB3.Text
            Form1.GridString(10) = TextC3.Text
            Form1.GridString(11) = TextD3.Text
            Form1.GridString(13) = TextB4.Text
            Form1.GridString(14) = TextC4.Text
            Form1.GridString(15) = TextD4.Text
            Form1.GridString(17) = TextB5.Text
            Form1.GridString(18) = TextC5.Text
            Form1.GridString(19) = TextD5.Text
            Form1.GridString(21) = TextB6.Text
            Form1.GridString(22) = TextC6.Text
            Form1.GridString(23) = TextD6.Text
            Form1.FindAllCaller = 0
            Form1.EditFind.Enabled = True
        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub GridSearch_Click()

        List1.Visible = False
    End Sub

    Private Sub GridSearch_GotFocus()

        TextA1.Select()
    End Sub

    Private Sub LoadSympWordPtrs()
        ' Reads the List1 item data, copies word pointers to SymptomWords array based on symptom pointers; then sorts the
        ' word pointers for each symptom pointer, eliminates duplicates, and stores counts in NumWords array.

        Dim ElementLength As Integer
        Dim IdxList() As String
        Dim MoreIDs As Boolean
        Dim msg As String   ' Error message string
        Dim Ptr1, Ptr2, SmPtr As Integer   ' sort pointers
        Dim Smallest As Integer ' used for sort
        Dim StrSympID As String
        Dim SympID As Integer
        Dim SympList As String  ' List of word's symptom ID's from List1 data
        Dim WordCount As Integer    ' Used for duplicate elimination loop
        Dim WordPtr As Integer  ' For outer loop, points to each word index in List1

        Try
            Form1.ErrorFlag = 5005
            ElementLength = List1.Items.Count
            ReDim IdxList(ElementLength)

            ' Load up the indexes from List1
            For Ptr1 = 0 To List1.Items.Count - 1
                IdxList(Ptr1) = gridSearData(Ptr1).SearIdxName.ToString()
            Next Ptr1

            ' For each word index
            For WordPtr = 0 To ElementLength - 1    ' The word being updated
                SympList = IdxList(WordPtr)
                If InStr(SympList, ";") > 0 Then
                    SympList = Strings.Left(SympList, (InStr(SympList, ";")) - 1)
                End If
                MoreIDs = True
                While MoreIDs
                    If InStr(SympList, ",") > 0 Then
                        StrSympID = Strings.Left(SympList, InStr(SympList, ",") - 1)
                        SympList = SympList.Substring(InStr(SympList, ","))
                    Else
                        StrSympID = SympList
                        MoreIDs = False
                    End If
                    SympID = Int(StrSympID)
                    SymptomWords(SympID, NumWords(SympID)) = WordPtr + 1    ' Need to decrement these sometime later.
                    NumWords(SympID) += 1
                End While
            Next

            ' Now sort the word ID's for each symptom
            For SympID = 0 To NumWords.Length - 1
                If NumWords(SympID) > 1 Then
                    For Ptr1 = 0 To NumWords(SympID) - 1
                        Smallest = 32767
                        For Ptr2 = Ptr1 To NumWords(SympID) - 1
                            If SymptomWords(SympID, Ptr2) < Smallest Then
                                Smallest = SymptomWords(SympID, Ptr2)
                                SmPtr = Ptr2
                            End If
                        Next Ptr2
                        SymptomWords(SympID, SmPtr) = SymptomWords(SympID, Ptr1)
                        SymptomWords(SympID, Ptr1) = Smallest
                    Next Ptr1
                End If

                ' Need to get rid of any duplicates.
                WordCount = NumWords(SympID)    ' Save off for loop counters
                For Ptr1 = 1 To WordCount
                    If SymptomWords(SympID, Ptr1 - 1) = SymptomWords(SympID, Ptr1) Then 'If a dumplicate
                        For Ptr2 = Ptr1 To NumWords(SympID)
                            ' Move all word pointers down one index, overwriting the duplicate; 0's replace upper-most one since ID's above top are init to 0.
                            SymptomWords(SympID, Ptr2) = SymptomWords(SympID, Ptr2 + 1)
                        Next
                        NumWords(SympID) -= 1
                    End If
                Next
            Next
        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Function GenWordPtrsFromSympIDs(SympIDs As String) As String
        ' For each symptom id in SympIDs, generates a list of word pointers; then sorts the word pointers, removes any
        ' duplicates and returns the list of word pointers.
        Dim CommaPtr As Integer ' points to next comma in IDs string
        Dim IntSympID As Integer    ' integer form of StrSympID
        Dim LocNumWds As Integer    ' number of words in WordIDs array
        Dim LocSympIDs As String    ' local copy of SympIDs
        Dim msg As String   ' Error message string
        Dim Ptr1, Ptr2, Smptr As Integer
        Dim Smallest As Integer ' comparison used for sort
        Dim StrSympID As String ' individual symptom ID from SympIDs
        Dim WordIDs() As Integer    ' array of word IDs from SymptomWords
        Dim WordPtr As Integer  ' pointer to WordIDs element

        Try
            Form1.ErrorFlag = 5006
            LocSympIDs = SympIDs
            GenWordPtrsFromSympIDs = "" ' Prevent null pointer exception in case return before this is set.

            CommaPtr = InStr(LocSympIDs, ",")
            WordPtr = 0
            Ptr1 = 0    ' In case While does not execute
            While CommaPtr > 0    ' While more than one ID in LocSympIDs string
                StrSympID = Strings.Left(LocSympIDs, CommaPtr - 1)
                IntSympID = Int(StrSympID)
                LocSympIDs = LocSympIDs.Substring(CommaPtr)
                CommaPtr = InStr(LocSympIDs, ",")
                For Ptr1 = 0 To NumWords(IntSympID) - 1
                    ReDim Preserve WordIDs(WordPtr + 1)
                    WordIDs(WordPtr) = SymptomWords(IntSympID, Ptr1)
                    WordPtr += 1
                Next Ptr1
            End While

            ' Handle last or only word pointer.
            StrSympID = LocSympIDs
            IntSympID = Int(StrSympID)
            ReDim Preserve WordIDs(WordPtr + 1)
            WordIDs(WordPtr) = SymptomWords(IntSympID, Ptr1)

            LocNumWds = WordPtr
            ' Sort the words in InterimWordData
            For Ptr1 = 0 To LocNumWds - 1
                Smallest = 32767
                For Ptr2 = Ptr1 To LocNumWds - 1
                    If WordIDs(Ptr2) < Smallest Then
                        Smallest = WordIDs(Ptr2)
                        Smptr = Ptr2
                    End If
                Next Ptr2
                WordIDs(Smptr) = WordIDs(Ptr1)
                WordIDs(Ptr1) = Smallest
            Next Ptr1

            ' Eliminate any duplicates
            Ptr1 = 1
            While Ptr1 < LocNumWds
                While WordIDs(Ptr1 - 1) = WordIDs(Ptr1)
                    For Ptr2 = Ptr1 To LocNumWds - 2
                        ' Move all word pointers down one index, overwriting the duplicate; 0's replace upper-most one since ID's above top are init to 0.
                        WordIDs(Ptr2) = WordIDs(Ptr2 + 1)
                    Next
                    WordIDs(LocNumWds - 1) = 0
                    LocNumWds -= 1
                End While
                Ptr1 += 1
            End While

            ' Add symptom pointers from WordIDs to GenWordPtrsFromSympIDs.
            Ptr1 = 0
            GenWordPtrsFromSympIDs = "; "
            'nnnn GenWordPtrsFromSympIDs = ", "
            While WordIDs(Ptr1) <> 0   ' While more pointers to add
                GenWordPtrsFromSympIDs += (WordIDs(Ptr1) - 1).ToString  ' Word ID's are 0-based so need to subtract 1.
                Ptr1 += 1
                If WordIDs(Ptr1) <> 0 Then GenWordPtrsFromSympIDs += "; "  ' If more to add, append a semicolon.
                'nnnn If WordIDs(Ptr1) <> 0 Then GenWordPtrsFromSympIDs += ", "  ' If more to add, append a comma.
            End While
        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            GenWordPtrsFromSympIDs = ""
        End Try
    End Function

    Private Sub FindFirst_Click(sender As Object, e As EventArgs) Handles FindFirst.Click
        Dim El(6) As Integer    ' selected element from each row
        Dim Found(4) As Boolean    ' a common element was found
        Dim loc(4, 6) As Integer     ' local copy of GIn array
        Dim msg As String       ' error message string

        Try
            Form1.ErrorFlag = 5007

            Cursor = Cursors.WaitCursor    ' Change cursor to hourglass.

            Call Load_Indexes() ' Load index array from List1 selected item data

            ' Set idxPtr back to 0 because we are finding first.
            idxPtr = 0

            ' Scroll Form1.Lst1 to appropriate line
            ' Unfortunately the ListBox.FindString method will not scroll the list box if the string is visible in the list.
            ' This means the found item will not always appear at the top of the list.
            ' This needs to be detailed in the instructions / help files.
            Form1.Lst1.TopIndex = idxArray(idxPtr) - 1  ' List item lines are 0-based
            If NumIndexes > 1 Then
                Me.FindNext.Enabled = True
            End If

            FindNext.Select()
            Cursor = Cursors.Default    ' Change cursor to default symbol.
            Exit Sub

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            Close()
        End Try
    End Sub

    Private Sub Load_Indexes()
        ' Loads the index array based on the symptom pointers indicated in the grid result.
        ' The right-most TextData element in each row already contains the AND'ed result for the row.
        ' So we just need to OR the right-most TextData element from each row that contains words.

        Dim locIdxLineDat As String
        Dim Ptr1, Ptr2 As Integer  ' loop pointers
        Dim CommaPtr As Integer ' Points to first comma in string
        Dim Smallest As Integer         ' smallest sort element
        Dim SmPtr As Integer            ' position of Smallest in array

        ' Find right-most TextData from each row.  Strip off the word pointers, OR them together, sort, then eliminate duplicates.
        locIdxLineDat = ""
        For Ptr1 = 0 To 20 Step 4    ' Do for each row
            ' Find the right-most data
            If TextData(Ptr1 + 3) <> "" Then
                locIdxLineDat += "," + Strings.Left(TextData(Ptr1 + 3), InStr(TextData(Ptr1 + 3), ";") - 1)  ' Strip off word pointers used by grid search
            ElseIf TextData(Ptr1 + 2) <> "" Then
                locIdxLineDat += "," + Strings.Left(TextData(Ptr1 + 2), InStr(TextData(Ptr1 + 2), ";") - 1)  ' Strip off word pointers used by grid search
            ElseIf TextData(Ptr1 + 1) <> "" Then
                locIdxLineDat += "," + Strings.Left(TextData(Ptr1 + 1), InStr(TextData(Ptr1 + 1), ";") - 1)  ' Strip off word pointers used by grid search
            ElseIf TextData(Ptr1) <> "" Then
                locIdxLineDat += "," + Strings.Left(TextData(Ptr1), InStr(TextData(Ptr1), ";") - 1)  ' Strip off word pointers used by grid search
            Else
                ' This row doesn't contain any words, there's nothing to do.
            End If
        Next

        NumIndexes = CommaCount(locIdxLineDat)
        ReDim idxArray(NumIndexes - 1)
        For Ptr1 = 0 To idxArray.Length - 1
            CommaPtr = InStr(locIdxLineDat, ",")
            If CommaPtr <> 0 Then
                locIdxLineDat = Strings.Mid(locIdxLineDat, InStr(locIdxLineDat, ",") + 1)
                idxArray(Ptr1) = Val(locIdxLineDat)
            Else
                idxArray(Ptr1) = Val(locIdxLineDat)
            End If
        Next

        ' Now sort the elements of idxArray()
        For Ptr1 = 0 To idxArray.Length - 1
            Smallest = 32767
            For Ptr2 = Ptr1 To idxArray.Length - 1
                If idxArray(Ptr2) < Smallest Then
                    Smallest = idxArray(Ptr2)
                    SmPtr = Ptr2
                End If
            Next Ptr2
            ' InterimWordData(SmPtr) = InterimWordData(Pntr1)
            idxArray(Ptr1) = Smallest
        Next Ptr1

        ' Eliminate any duplicates
        If NumIndexes > 1 Then
            Ptr1 = 1
            'While Ptr1 <= NumIndexes
            While Ptr1 < NumIndexes
                If idxArray(Ptr1) = idxArray(Ptr1 - 1) Then
                    For Ptr2 = Ptr1 To NumIndexes - 1
                        If Ptr2 < (NumIndexes - 1) Then
                            idxArray(Ptr2) = idxArray(Ptr2 + 1)
                        Else
                            idxArray(Ptr2) = 0
                        End If
                    Next
                    NumIndexes -= 1
                Else
                    Ptr1 += 1
                End If
            End While
        End If

    End Sub

    Private Sub FindNext_Click(sender As Object, e As EventArgs) Handles FindNext.Click
        Dim ErrorFlag As Integer    ' error number
        Dim msg As String   ' Error message string

        Try
            ErrorFlag = 5008

            If idxPtr < NumIndexes Then ' Should always be true
                FindPrev.Enabled = True
                idxPtr += 1  ' Point to next index
                ' Scroll Form1.Lst1 to appropriate line
                ' Unfortunately the ListBox.FindString method will not scroll the list box if the string is visible in the list.
                ' This means the found item will not always appear at the top of the list.
                ' This needs to be detailed in the instructions / help files.
                Form1.Lst1.TopIndex = idxArray(idxPtr) - 1  ' List item lines are 0-based
                If idxPtr = NumIndexes - 1 Then ' No more indexes to find
                    Me.FindNext.Enabled = False
                    FindPrev.Select()
                Else
                    FindNext.Select()
                End If
            End If

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub FindPrev_Click(sender As Object, e As EventArgs) Handles FindPrev.Click
        Dim ErrorFlag As Integer    ' error number
        Dim msg As String   ' Error message string

        Try
            Form1.ErrorFlag = 5009
            If idxPtr > 0 Then  ' Should always be true
                idxPtr -= 1
                ' Scroll Form1.Lst1 to appropriate line
                ' Unfortunately the ListBox.FindString method will not scroll the list box if the string is visible in the list.
                ' This means the found item will not always appear at the top of the list.
                Form1.Lst1.TopIndex = idxArray(idxPtr) - 1  ' List item lines are 0-based
                If idxPtr > 0 Then
                    Me.FindPrev.Enabled = False
                End If
            End If

            Me.FindNext.Enabled = True

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub FindAllButton_Click(sender As Object, e As EventArgs) Handles FindAllButton.Click
        Dim AllToFind(1) As Integer  ' array of all found indexes
        Dim loc(4, 6) As Integer     ' local copy of GIn array
        Dim msg As String       ' error message string
        Dim findallSymptomID As Integer ' symptom ID of Lst1 symptom

        Try
            Form1.ErrorFlag = 5010

            Cursor = Cursors.WaitCursor    ' Change cursor to hourglass.
            If FindAll.Visible Then FindAll.Close()   ' Need to re-initialize in case it is open.
            Form1.FindAllCaller = 2
            FindAll.Left = 0
            FindAll.Top = 0
            FindAll.Label1.Visible = False
            FindAll.Show()

            Call Load_Indexes() ' Load index array from List1 selected item data

            Form1.ErrorFlag = 5011
            FindAll.FindAllSympData.Clear()
            For idxPtr = 0 To NumIndexes - 1
                ' Load FindAll Lst1 with Form1.Lst1 items pointed to by idxArray
                HAHN30VBDev.FindAll.FindAllSympData.Add(New FindAllSympDat(Form1.CombSympData(idxArray(idxPtr) - 1).SympDesc.ToString, Form1.CombSympData(idxArray(idxPtr) - 1).SympData.ToString))
            Next

            ' Associate FindAll.Lst1 with FindAllSympData
            FindAll.Lst1.DataSource = Nothing   ' Clear any list items
            FindAll.Lst1.DataSource = FindAll.FindAllSympData

            FindAll.Lst1.DisplayMember = "SympDesc"
            FindAll.Lst1.ValueMember = "SympData"
            FindAll.Lst1.ClearSelected()
            ' Now select any symptoms that are selected in Form1.Lst1
            For idxPtr = 0 To NumIndexes - 1
                findallSymptomID = Form1.GetSymptomID(FindAll.FindAllSympData(idxPtr).SympData.ToString)
                If Form1.IsSymptomSelected(Form1.MapSymIDToIndex(findallSymptomID)) Then
                    FindAll.Lst1.SetSelected(idxPtr, True)
                End If
                If Form1.RaiseException Then Throw New System.Exception("An exception has occurred.")
            Next

            ' Clear any symptom that may have been shown in FindAll TextBox1 as a result of adding symptoms.
            FindAll.TextBox1.Text = ""
            Cursor = Cursors.Default    ' Change cursor to default symbol.
            Exit Sub

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
        Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        Close()
        End Try
    End Sub

    Private Sub Command2_Click(sender As Object, e As EventArgs) Handles Command2.Click ' Cancel
        'Close()
        Me.Visible = False
    End Sub

    Private Sub Clear_Click(sender As Object, e As EventArgs) Handles Clear.Click
        Dim mbSelection As Integer  ' Indicates which message box button was clicked
        Dim msg As String   ' Message box text to display
        Dim TextPtr As Integer

        Try
            Form1.ErrorFlag = 5012
            msg = "Warning:  all text selections in this form will be deleted.  Continue?"
            mbSelection = MsgBox(Prompt:=msg,
                Buttons:=vbYesNo + vbQuestion,
                            Title:="?")
            If mbSelection = vbYes Then

                For TextPtr = 0 To 23
                    Text1(TextPtr).Text = ""
                    TextData(TextPtr) = ""
                    If TextPtr <> 0 Then Text1(TextPtr).Enabled = False
                Next TextPtr
                TextA1.Select()
                Text1_Click(0)  ' Pop up List1 and allow user to select a word.
            End If

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub Text1_Changed(Index As Integer)
        ' Three events are used to control the action of the grid search:
        ' When the first letter is typed in a text box (indicating that it will be used), a list is produced of all possible
        ' words that could be used to select symptoms.  If it is in the first column, the complete word list (all words from
        ' all symptoms) shows in List1, which is pre-loaded for performance purposes due to its large size.  For the second column,
        ' the list is all words from all symptoms that contain the first column word.  For the third and fourth columns, the
        ' possible symptoms from the previous columns are AND'ed together to produce lists of possible words.  Columns two
        ' through four use List2, which is generated as needed for each text box.
        ' As letters are typed (including the first letter), the list is scrolled so that the first word in the alphabetical
        ' list which begins with the typed letter(s) is at the top.
        ' In order for a word to be "selected", the user needs to click on it in the list.  If the word is in the
        ' first column, the text box below (if there is one) is enabled.  Also as each word is selected, if more than one symptom
        ' contains all the words in the row, the text box to the right is enabled (unless this is the fourth column).

        Dim listIndex As Integer    ' list index of found string
        Dim InterimWordData(7800) As Integer
        Dim msg As String   ' Error message string

        ' Changing text scrolls the list to first occurance of letter(s) entered.  All text boxes to the right are disabled.
        ' Selecting the item in List1 updates the text (if different) and "selects" the word.
        Try
            If Not SupTxtChgHnd Then    ' Only handle text change if changed by user.
                textHasChanged = True
                If SkipChange Then Exit Sub
                Form1.ErrorFlag = 5013
                TextIndex = Index   ' Tells List1 SelectIndexChanged which text box to use.

                If Text1(Index).Text = "" Then NumElements(Index) = 0

                ' Disable any text boxes to the right.
                If Index <> 3 And Index <> 7 And Index <> 11 And Index <> 15 And Index <> 19 And Index <> 23 Then
                    Text1(Index + 1).Enabled = False
                End If

                Command2.Text = "Done"

                Form1.ErrorFlag = 5014

                ' Unfortunately the ListBox.FindString method will not scroll the list box if the string is visible in the list.
                ' This means the found item will not always appear at the top of the list.
                If Index = 0 Or Index = 4 Or Index = 8 Or Index = 12 Or Index = 16 Or Index = 20 Then    ' If first column
                    ' Scroll List1 to the first occurance of the text letter(S)
                    listIndex = List1.FindString(Text1(Index).Text)
                    List1.TopIndex = listIndex
                Else    ' The rest of the columns use List2
                    ' Scroll List2 to the first occurance of the text letter(S)
                    listIndex = List2.FindString(Text1(Index).Text)
                    List2.TopIndex = listIndex
                End If

            End If

            textHasChanged = False
        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
             Buttons:=vbOKOnly + vbCritical,
             Title:="ERROR!")
            Close()
        End Try
    End Sub

    Private Function GetNumberPointers(TData As String) As Integer
        ' Return the number of comma-separated elements in the string
        Dim CommaCount As Integer
        Dim ErrorFlag As Integer    ' error number
        Dim msg As String   ' Error message string
        Dim Ptr As Integer

        GetNumberPointers = 0
        Try
            ErrorFlag = 5015
            CommaCount = 0
            Ptr = 1
            While Ptr < Len(TData)
                If Mid(TData, Ptr, 1) = "," Then CommaCount += 1
                Ptr = Ptr + 1
            End While
            GetNumberPointers = CommaCount + 1  ' There will be one more pointer than the number of commas
        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Function

    Private Sub TextA1_TextChanged(sender As Object, e As EventArgs) Handles TextA1.TextChanged
        Text1_Changed(0)
    End Sub

    Private Function CommaCount(InString As String) As Integer
        ' Returns the number of commas in the string
        Dim Cntr As Integer ' keeps track of number of commas
        Dim ErrorFlag As Integer    ' error number
        Dim msg As String   ' Error flag string
        Dim Pntr As Integer ' points to next comma in string

        CommaCount = 0
        Try
            ErrorFlag = 5016
            Cntr = 0

            While InStr(InString, ",") > 0
                Pntr = InStr(InString, ",")
                Cntr += 1
                InString = InString.Substring(Pntr)
            End While

            CommaCount = Cntr
        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
               Buttons:=vbOKOnly + vbCritical,
               Title:="ERROR!")
        End Try
    End Function

    Private Sub List1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles List1.SelectedIndexChanged
        ' Called when user selects a word from the list; List1 is used for text boxes in the first column
        Dim msg As String   ' Error message string

        Try
            Form1.ErrorFlag = 5017
            If Not textHasChanged Then
                ' The word has been selected.  Copy it to the selecting text box and enable the text box to the right if there is one.
                If List1.SelectedIndex >= 0 Then    ' If an item is selected.
                    Text1(TextIndex).Text = gridSearData(List1.SelectedIndex).SearDataName.ToString
                    TextData(TextIndex) = gridSearData(List1.SelectedIndex).SearIdxName.ToString    ' From seartxtidx.dat
                    ' Get word pointers for these Symptom IDs and add them to TextData; they will each be preceded by a semicolon.
                    TextData(TextIndex) += GenWordPtrsFromSympIDs(TextData(TextIndex))

                End If
                ' If there is more than one possible symptom, enable next text box.
                If CommaCount(TextData(TextIndex)) > 0 Then
                    Text1(TextIndex + 1).Enabled = True
                    If TextIndex <> 20 Then ' If not last text box in first column, enable the text box below it.
                        Text1(TextIndex + 4).Enabled = True
                    End If
                End If
                ' Enable "Find First" and "Find All" buttons.
                FindFirst.Enabled = True
                FindAllButton.Enabled = True
            End If
        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub List1_Click(sender As Object, e As EventArgs) Handles List1.Click
        Dim ErrorFlag As Integer    ' error number
        Dim lPtr1 As Integer     ' string pointer
        Dim msg As String       ' Error message string

        Try
            ErrorFlag = 5018
            SupTxtChgHnd = True ' Don't allow Text1 text changed handler code to execute
            If List1.SelectedIndex <> -1 Then    ' If an item is selected.
                locLineDat = Find.searData(List1.SelectedIndex).SearDataName.ToString
                lPtr1 = InStr(locLineDat, Chr(9))
                If lPtr1 > 0 Then locLineDat = Mid(locLineDat, 1, lPtr1 - 1)
                Text1(TextIndex).Text = locLineDat
            End If
            SupTxtChgHnd = False

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub TextB1_TextChanged(sender As Object, e As EventArgs) Handles TextB1.TextChanged
        Text1_Changed(1)
    End Sub

    Private Sub TextC1_TextChanged(sender As Object, e As EventArgs) Handles TextC1.TextChanged
        Text1_Changed(2)
    End Sub

    Private Sub TextD1_TextChanged(sender As Object, e As EventArgs) Handles TextD1.TextChanged
        Text1_Changed(3)
    End Sub

    Private Sub TextB2_TextChanged(sender As Object, e As EventArgs) Handles TextB2.TextChanged
        Text1_Changed(5)
    End Sub

    Private Sub TextC2_TextChanged(sender As Object, e As EventArgs) Handles TextC2.TextChanged
        Text1_Changed(6)
    End Sub

    Private Sub TextD2_TextChanged(sender As Object, e As EventArgs) Handles TextD2.TextChanged
        Text1_Changed(7)
    End Sub

    Private Sub TextB3_TextChanged(sender As Object, e As EventArgs) Handles TextB3.TextChanged
        Text1_Changed(9)
    End Sub

    Private Sub TextC3_TextChanged(sender As Object, e As EventArgs) Handles TextC3.TextChanged
        Text1_Changed(10)
    End Sub

    Private Sub TextD3_TextChanged(sender As Object, e As EventArgs) Handles TextD3.TextChanged
        Text1_Changed(11)
    End Sub

    Private Sub TextB4_TextChanged(sender As Object, e As EventArgs) Handles TextB4.TextChanged
        Text1_Changed(13)
    End Sub

    Private Sub TextC4_TextChanged(sender As Object, e As EventArgs) Handles TextC4.TextChanged
        Text1_Changed(14)
    End Sub

    Private Sub TextD4_TextChanged(sender As Object, e As EventArgs) Handles TextD4.TextChanged
        Text1_Changed(15)
    End Sub

    Private Sub TextB5_TextChanged(sender As Object, e As EventArgs) Handles TextB5.TextChanged
        Text1_Changed(17)
    End Sub

    Private Sub TextC5_TextChanged(sender As Object, e As EventArgs) Handles TextC5.TextChanged
        Text1_Changed(18)
    End Sub

    Private Sub TextD5_TextChanged(sender As Object, e As EventArgs) Handles TextD5.TextChanged
        Text1_Changed(19)
    End Sub

    Private Sub TextB6_TextChanged(sender As Object, e As EventArgs) Handles TextB6.TextChanged
        Text1_Changed(21)
    End Sub

    Private Sub TextC6_TextChanged(sender As Object, e As EventArgs) Handles TextC6.TextChanged
        Text1_Changed(22)
    End Sub

    Private Sub TextD6_TextChanged(sender As Object, e As EventArgs) Handles TextD6.TextChanged
        Text1_Changed(23)
    End Sub

    Private Sub TextA1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextA1.Click
        Text1_Click(0)
    End Sub
    Private Sub TextB1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextB1.Click
        Text1_Click(1)
    End Sub
    Private Sub TextC1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextC1.Click
        Text1_Click(2)
    End Sub
    Private Sub TextD1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextD1.Click
        Text1_Click(3)
    End Sub
    Private Sub TextA2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextA2.Click
        Text1_Click(4)
    End Sub
    Private Sub TextB2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextB2.Click
        Text1_Click(5)
    End Sub
    Private Sub TextC2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextC2.Click
        Text1_Click(6)
    End Sub
    Private Sub TextD2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextD2.Click
        Text1_Click(7)
    End Sub
    Private Sub TextA3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextA3.Click
        Text1_Click(8)
    End Sub
    Private Sub TextB3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextB3.Click
        Text1_Click(9)
    End Sub
    Private Sub TextC3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextC3.Click
        Text1_Click(10)
    End Sub
    Private Sub TextD3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextD3.Click
        Text1_Click(11)
    End Sub
    Private Sub TextA4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextA4.Click
        Text1_Click(12)
    End Sub
    Private Sub TextB4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextB4.Click
        Text1_Click(13)
    End Sub
    Private Sub TextC4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextC4.Click
        Text1_Click(14)
    End Sub
    Private Sub TextD4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextD4.Click
        Text1_Click(15)
    End Sub
    Private Sub TextA5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextA5.Click
        Text1_Click(16)
    End Sub
    Private Sub TextB5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextB5.Click
        Text1_Click(17)
    End Sub
    Private Sub TextC5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextC5.Click
        Text1_Click(18)
    End Sub
    Private Sub TextD5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextD5.Click
        Text1_Click(19)
    End Sub
    Private Sub TextA6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextA6.Click
        Text1_Click(20)
    End Sub
    Private Sub TextB6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextB6.Click
        Text1_Click(21)
    End Sub
    Private Sub TextC6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextC6.Click
        Text1_Click(22)
    End Sub
    Private Sub TextD6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextD6.Click
        Text1_Click(23)
    End Sub

    Public Sub Text1_Click(Index As Integer)
        ' When text box is clicked, need to display List1 if in first column, otherwise load and display List2.
        Dim mbSelection As Integer  ' msg box selection
        Dim msg As String   ' Dialog box message
        Dim OkToContinue As Boolean ' Either next text box doesn't contain text or user selected to delete next box text.
        Dim Ptr As Integer  ' Points to next ";"
        Dim WordData As String      ' local copy of TextData from previous text box.
        Dim WordIndex As Integer    ' List1 index of word to display
        Dim WordPointer As String   ' String form of word pointer from WordData

        '        Try
        Form1.ErrorFlag = 5019
            TextIndex = Index   ' Point to list box index.
        If Index = 0 Or Index = 4 Or Index = 8 Or Index = 12 Or Index = 16 Or Index = 20 Then   ' If first column
            List2.Visible = False
            List1.Visible = True
            ' If the text box to the right already contains text, pop-up dialog to warn that any words to the right will be
            ' deleted and allow user to back out.
            If Text1(Index + 1).Text <> "" Then
                msg = "Warning:  any text selections to the right of this box will be deleted.  Continue?"
                mbSelection = MsgBox(Prompt:=msg,
                Buttons:=vbYesNo + vbQuestion,
                            Title:="?")
                If mbSelection = vbYes Then
                    Text1(Index).Text = ""
                    TextData(Index) = ""
                    Text1(Index + 1).Text = ""
                    TextData(Index + 1) = ""
                    Text1(Index + 1).Enabled = False
                    Text1(Index + 2).Text = ""
                    TextData(Index + 2) = ""
                    Text1(Index + 2).Enabled = False
                    Text1(Index + 3).Text = ""
                    TextData(Index + 3) = ""
                    Text1(Index + 3).Enabled = False
                End If
            End If
            ' In case List1 was used previously to select a word, clear any selection and scroll it back to the top.
            List1.ClearSelected()
            List1.TopIndex = 0
        Else    ' Other than first column
            OkToContinue = True
                If Index = 1 Or Index = 5 Or Index = 9 Or Index = 13 Or Index = 17 Or Index = 21 Then   ' If second column
                    If Text1(Index + 1).Text <> "" Then
                        msg = "Warning:  any text selections to the right of this box will be deleted.  Continue?"
                        mbSelection = MsgBox(Prompt:=msg,
                Buttons:=vbYesNo + vbQuestion,
                            Title:="?")
                        If mbSelection = vbYes Then
                            Text1(Index).Text = ""
                            TextData(Index) = ""
                            Text1(Index + 1).Text = ""
                            TextData(Index + 1) = ""
                            Text1(Index + 1).Enabled = False
                            Text1(Index + 2).Text = ""
                            TextData(Index + 2) = ""
                            Text1(Index + 2).Enabled = False
                        Else
                            OkToContinue = False
                        End If
                    End If
                ElseIf Index = 2 Or Index = 6 Or Index = 10 Or Index = 14 Or Index = 18 Or Index = 22 Then   ' If third column
                    If Text1(Index + 1).Text <> "" Then
                        msg = "Warning:  any text selections to the right of this box will be deleted.  Continue?"
                        mbSelection = MsgBox(Prompt:=msg,
                        Buttons:=vbYesNo + vbQuestion,
                        Title:="?")
                        If mbSelection = vbYes Then
                            Text1(Index).Text = ""
                            TextData(Index) = ""
                            Text1(Index + 1).Text = ""
                            TextData(Index + 1) = ""
                            Text1(Index + 1).Enabled = False
                        Else
                            OkToContinue = False
                        End If
                    End If
                Else    ' This is the fourth column; no need to warn about deletions.
                    Text1(Index).Text = ""
                    TextData(Index) = ""
                End If

                If OkToContinue Then
                    ' Clear old words out of List2.
                    grid2SearData.Clear()
                ' Load List2 with required words and display it.
                WordData = TextData(Index - 1)  ' Get word information passed in from text box to the left; first part is word pointers from seartxtidx.dat separated by commas, second part is word ID's generated from symptom pointers, separated by ";".
                ' At this point it is known that there will be more than one word (otherwise there is no point in making a selection); thus there will always be a ";" in WordData.
                WordData = WordData.Substring(InStr(WordData, ";") - 1) ' Remove symptom pointers portion of string.
                    While InStr(WordData, ";") > 0
                        WordData = WordData.Substring(2)    ' Remove leading "; "
                        If InStr(WordData, ";") > 0 Then
                            Ptr = InStr(WordData, ";") - 1  ' Skip the space after the ";"
                            WordPointer = (Strings.Left(WordData, Ptr))
                        WordIndex = Val(WordPointer)
                        WordData = WordData.Substring(Ptr)
                        Else
                        WordIndex = Val(WordData)
                    End If
                        ' Add List1 word pointed to by WordIndex to List2.
                        grid2SearData.Add(New grid2SearDat(gridSearData(WordIndex).SearDataName.ToString, gridSearData(WordIndex).SearIdxName.ToString))
                    End While

                    textHasChanged = True ' Suppress action of List2_SelectedIndexChanged()
                    List2.DataSource = Nothing ' Remove and re-associate database in order to refresh List1
                    List2.DataSource = grid2SearData
                    List2.DisplayMember = "Sear2DataName"
                    List2.ValueMember = "Sear2IdxName"
                    List2.ClearSelected()
                    List1.Visible = False   ' List1 is pre-loaded and used for text boxes in the first column
                    List2.Visible = True   ' List2 is used for text boxes in columns to the right of the first column
                    textHasChanged = False
                End If
            End If

        '        Catch
        '        msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
        '        Dim unused = MsgBox(Prompt:=msg,
        '        Buttons:=vbOKOnly + vbCritical,
        '        Title:="ERROR!")
        '        End Try
    End Sub


    Private Sub TextA2_TextChanged(sender As Object, e As EventArgs) Handles TextA2.TextChanged
        Text1_Changed(4)
    End Sub

    Private Sub TextA3_TextChanged(sender As Object, e As EventArgs) Handles TextA3.TextChanged
        Text1_Changed(8)
    End Sub

    Private Sub TextA4_TextChanged(sender As Object, e As EventArgs) Handles TextA4.TextChanged
        Text1_Changed(12)
    End Sub

    Private Sub TextA5_TextChanged(sender As Object, e As EventArgs) Handles TextA5.TextChanged
        Text1_Changed(16)
    End Sub

    Private Sub List2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles List2.SelectedIndexChanged
        ' Called when user selects a word in the list; List2 is used for words not in the first column.
        Dim cp1, cp2 As Integer    ' comma pointers
        Dim InterimWordData(5000) As Integer    ' unsorted word pointers from AND comparison of the two TextData indices
        Dim msg As String   ' Error message string
        Dim NumIndex As Integer ' number of symptom pointers in InterimWordData
        Dim Ptr1, Ptr2, Smptr As Integer
        Dim SelText1, SelText2 As String   ' data string from selected word in previous columns
        Dim Sel1, Sel2 As Integer   ' sizes (# of symptom pointers in) SelText<n>
        Dim Smallest As Integer ' comparison used for sort
        Dim WP1, WP2 As Integer ' word pointers

        Try
            Form1.ErrorFlag = 5020
            SupTxtChgHnd = True ' Tell Textbox event handler we will be changing the text from here.
            If Not textHasChanged Then ' Suppress this subroutine action if call resulted from initializing List2 with new items.
                ' The word has been selected.  Copy it to the selecting text box and enable the text box to the right if there is one.
                If Me.List2.SelectedIndex >= 0 Then    ' If an item is selected.
                    Text1(TextIndex).Text = grid2SearData(List2.SelectedIndex).Sear2DataName.ToString
                    SelText1 = TextData(TextIndex - 1)
                    SelText2 = grid2SearData(List2.SelectedIndex).Sear2IdxName.ToString
                    Sel1 = CommaCount(SelText1) + 1 ' First word doesn't have a preceding comma
                    'Sel2 = CommaCount(SelText2)
                    Sel2 = CommaCount(SelText2) + 1 ' First word doesn't have a preceding comma
                    Ptr1 = InStr(SelText1, ",")
                    NumIndex = 0    ' Init pointer for InterimWordData array
                    For Pntr1 = 0 To Sel1
                        SelText2 = grid2SearData(List2.SelectedIndex).Sear2IdxName.ToString
                        For Ptr2 = 0 To Sel2
                            cp1 = InStr(SelText1, ",") - 1
                            cp2 = InStr(SelText2, ",") - 1
                            If cp1 >= 0 Then
                                WP1 = Int(Strings.Left(SelText1, cp1))
                            Else
                                If InStr(SelText1, ";") <> 0 Then
                                    cp1 = InStr(SelText1, ";") - 1
                                    WP1 = Int(Strings.Left(SelText1, cp1))
                                Else
                                    WP1 = Int(SelText1)
                                End If
                            End If
                            If cp2 >= 0 Then
                                WP2 = Int(Strings.Left(SelText2, cp2))
                            Else
                                If InStr(SelText2, ";") <> 0 Then
                                    cp2 = InStr(SelText2, ";") - 1
                                    WP2 = Int(Strings.Left(SelText2, cp2))
                                Else
                                    WP2 = Int(SelText2)
                                End If
                            End If
                            If WP1 = WP2 Then
                                InterimWordData(NumIndex) += WP1
                                NumIndex += 1
                            End If
                            SelText2 = SelText2.Substring(cp2 + 1)
                        Next
                        SelText1 = SelText1.Substring(cp1 + 1)
                    Next

                    NumIndex -= 1
                    ' Sort the words in InterimWordData
                    For Ptr1 = 0 To NumIndex - 1
                        Smallest = 32767
                        For Ptr2 = Ptr1 To NumIndex - 1
                            If InterimWordData(Ptr2) < Smallest Then
                                Smallest = InterimWordData(Ptr2)
                                Smptr = Ptr2
                            End If
                        Next Ptr2
                        '                    InterimWordData(Smptr) = InterimWordData(Pntr1)
                        InterimWordData(Ptr1) = Smallest
                    Next Ptr1

                    ' Eliminate any duplicates
                    For Pntr1 = 1 To NumIndex
                        If InterimWordData(Pntr1) = InterimWordData(Pntr1 - 1) Then
                            For Ptr2 = Ptr1 To NumIndex - 1
                                InterimWordData(Ptr2) = InterimWordData(Ptr2 + 1)
                            Next
                            InterimWordData(NumIndex) = 0
                            NumIndex -= 1
                        End If
                    Next

                    ' Add symptom pointers from InterimWordData to TextData.
                    Ptr1 = 0
                    TextData(TextIndex) = ""
                    While InterimWordData(Ptr1) <> 0   ' While more pointers to add
                        TextData(TextIndex) += InterimWordData(Ptr1).ToString
                        Ptr1 += 1
                        If InterimWordData(Ptr1) <> 0 Then TextData(TextIndex) += ", "  ' If more to add, append a comma.
                    End While

                    ' Get word pointers for these Symptom IDs and add them to TextData.
                    TextData(TextIndex) += GenWordPtrsFromSympIDs(TextData(TextIndex))

                End If
                ' If there is a text box to the right and more than one possible symptom, enable it.
                If TextIndex <> 3 And
            TextIndex <> 7 And
            TextIndex <> 11 And
            TextIndex <> 15 And
            TextIndex <> 19 And
            TextIndex <> 23 And
            CommaCount(TextData(TextIndex)) > 0 Then
                    Text1(TextIndex + 1).Enabled = True
                End If
            End If
            SupTxtChgHnd = False

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
            Buttons:=vbOKOnly + vbCritical,
            Title:="ERROR!")
        End Try
    End Sub

    Private Sub TextA6_TextChanged(sender As Object, e As EventArgs) Handles TextA6.TextChanged
        Text1_Changed(20)
    End Sub

End Class

Public Class gridSearDat
    Private mySearIdxName As String
    Private mySearDataName As String

    Public Sub New(ByVal strSearDataName As String, ByVal strSearIdxName As String)
        Me.mySearIdxName = strSearIdxName
        Me.mySearDataName = strSearDataName
    End Sub

    Public ReadOnly Property SearIdxName() As String
        Get
            Return mySearIdxName
        End Get
    End Property

    Public ReadOnly Property SearDataName() As String
        Get
            Return mySearDataName
        End Get
    End Property

End Class

Public Class grid2SearDat
    Private mySear2IdxName As String
    Private mySear2DataName As String

    Public Sub New(ByVal strSear2DataName As String, ByVal strSear2IdxName As String)
        Me.mySear2IdxName = strSear2IdxName
        Me.mySear2DataName = strSear2DataName
    End Sub

    Public ReadOnly Property Sear2IdxName() As String
        Get
            Return mySear2IdxName
        End Get
    End Property

    Public ReadOnly Property Sear2DataName() As String
        Get
            Return mySear2DataName
        End Get
    End Property

End Class