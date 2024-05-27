Option Explicit On
Public Class QuestProgress
    Dim First As Integer    ' top of binary search
    Dim Index As Integer    ' GI word index list pointer
    Dim Last As Integer     ' bottom of binary search
    Dim GIPtr As Integer    ' GI array pointer
    Dim GridLocLineDat As String ' array of grid index strings
    Dim locLineDat As String     ' selected list item (line from thes file containing symptom word and its synonyms)
    Dim loc_sfile_pos As Long    ' local symptom file position
    Dim loc_syfile_pos As Long    ' local synonym file position of symptom word (and its synonyms)
    Dim NumWords As Integer     ' size of WordArray
    Dim old_sfile_pos As Integer    ' last position in symptom file
    Dim PageNumber As String   ' page number of question
    Dim Ptr1, Ptr2 As Integer   ' string pointers
    Dim QuestionNumber As String    ' question # on page
    Public Shared SearWords() As String   ' array for words from seartxtidx.dat
    Public Shared SearPointers() As String    ' array for pointers from seartxtidx.dat
    Dim SenPtr As Integer       ' sentence pointer
    Dim SympWord As String      ' symp word from file seartxtidx.dat
    Dim SympPtrStr As String    ' string containing list of symptom pointers from files for single word
    Dim SympPointers As String  ' string containing all lists of symptom pointers for phrase
    Public Shared WordArray() As String   ' array for questionaire words (all words from each symptom that are contained in the thesaurus)
    Public Shared ThesWords() As String    ' array list to contain Thestxtidx.dat file Words
    Public Shared ThesPointers() As String    ' array list to contain Thestxtidx.dat file pointers
    Dim NumMatchesFounded As Integer   ' Number of matching words / synonyms found
    Public Shared PercentageRequired As Decimal   ' Percent of words required to match
    Dim StrsToDisplay(10000) As String ' Additional string data to display with symptom descriptions in symptoms pointed to by SympsToDisplay()
    Dim SympsToDisplay(10000) As Integer ' Symptom ID's of symptoms meeting minimum number criteria
    Public Shared STDCount As Integer ' Number of symptoms to display
    Dim WordPtrStr As String    ' string containing list of word pointers from thes file for word and synonyms

    Public Sub SymParse(SymText As String, Wt As Integer)
        Dim ErrorFlag As Integer
        Dim LocWordArray() As String    ' local words to OR together
        Dim MoreToDo As Byte        ' loop exit flag
        Dim msg As String
        Dim NumSen As Integer       ' # of sentences
        Dim Ptr As Integer          ' string pointer
        Dim Sentences() As String   ' sentences from the string
        Dim Temp As String          ' string buffer
        Dim WordPtr As Integer      ' word pointer

        Try
            ErrorFlag = 8080

            ' Break the string into sentences.
            If SymText <> "" Then
                ReDim Sentences(0)
                ReDim LocWordArray(0)
                NumSen = 1
                MoreToDo = True
                Temp = SymText

                ErrorFlag = 8081
                ' Store the significant words in an array (WordArray); separate them with logicals (And / Or).
                Index = 0
                Call GetSigWord(SymText)    ' Copies significant words to WordArray(); NumWords = WordArray size
                WordPtr = 1
                For Ptr = 0 To NumWords - 1
                    ReDim Preserve LocWordArray(WordPtr - 1)
                    LocWordArray(WordPtr - 1) = WordArray(Ptr)
                    WordPtr = WordPtr + 1
                Next Ptr

                Call FindWords(LocWordArray)    ' Generate symptom pointers in SympsToDisplay array, show symptoms in FindAll selection list.
            End If

            Exit Sub

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            Close()
                End Try
    End Sub

    Public Sub AddList()    ' Adds items to FindAll list box.

        Dim decrSTDCount As Integer ' amount to decrement STDCount due to SympsToDisplay 0-value members
        Dim LineDat1, LineDat2 As String
        Dim msg As String           ' error message string
        Dim MinIdx As Integer   ' SympsToDisplay index of minimum value item in swap
        Dim Ptr As Integer     ' string pointer
        Dim Ptr1, Ptr2 As Integer   ' loop counters
        Dim questprogressSymptomID As Integer ' symptom ID of Lst1 symptom
        Dim sTemp As String     ' Used to hold swap string
        Dim Temp As Integer     ' Used to hold swap item in sort

        ' Populate the "FindAll" symptom list from SympsToDisplay() and StrsToDisplay() arrays and SymptomData() element pointed to, and display
        ' the "FindAll" dialog.

        Try
            Form1.ErrorFlag = 8082
            Form1.RaiseException = False

            ' Need to sort SympsToDisplay() items so list item number will match FindAllSymptomData element number.
            ' NOTE:  Using the ListBox "Sort" option won't work here because it only sorts the display, not the object.
            ' StrsToDisplay() will be sorted along with this so they can be added to the correct symptom descriptions.
            For Ptr1 = 0 To STDCount - 2
                MinIdx = Ptr1
                For Ptr2 = Ptr1 + 1 To STDCount - 1
                    If SympsToDisplay(Ptr2) < SympsToDisplay(MinIdx) Then
                        MinIdx = Ptr2
                    End If
                Next
                If MinIdx <> Ptr1 Then
                    Temp = SympsToDisplay(Ptr1)
                    sTemp = StrsToDisplay(Ptr1)
                    SympsToDisplay(Ptr1) = SympsToDisplay(MinIdx)
                    StrsToDisplay(Ptr1) = StrsToDisplay(MinIdx)
                    SympsToDisplay(MinIdx) = Temp
                    StrsToDisplay(MinIdx) = sTemp
                End If
            Next

            ' FindAll Label1 is only visible for Questionnaire-produced symptom list.
            ' The first number is the page, followed by a "/" and the symptom number for pages 1 – 4; or a blank for pages 5 – 7 because these are patient data,
            ' which is not associated with any particular symptom.  The next number is the question on the page, numbering consecutively from the top as question #1.
            ' Following this is the symptom weight (based on your questionnaire input), then the description. 
            If FindAll.Visible Then FindAll.Close()   ' Need to re-initialize in case it is open.
            Form1.FindAllCaller = 3
            FindAll.Left = 0
            FindAll.Top = 0
            FindAll.Label1.Visible = True   ' Show "page - question" heading.
            FindAll.Show()
            FindAll.Activate()
            Form1.FindAllVisible = True

            Form1.ErrorFlag = 8083
            FindAll.FindAllSympData.Clear()
            If Form1.CalledFromShowSymSymps Then FindAll.Label1.Visible = False ' Don't show title for symptom#, page#, etc.

            decrSTDCount = 0
            For Ptr = 0 To (STDCount - 1)
                ' Note that since the array SympsToDisplay is 0-based, we need to subtract.
                If SympsToDisplay(Ptr) > 0 Then    ' Protect against causing array out of bounds trap
                    If Not Form1.CalledFromShowSymSymps Then  ' If called from QuestProgress
                        LineDat1 = StrsToDisplay(Ptr) + Form1.SymptomData((SympsToDisplay(Ptr) * 2) - 2) ' Get the requested symptom line from symptextdata.dat file
                    Else    ' If called from Form1
                        LineDat1 = Form1.SymptomData((SympsToDisplay(Ptr) * 2) - 2) ' Get the requested symptom line from symptextdata.dat file
                    End If
                    LineDat2 = Form1.SymptomData((SympsToDisplay(Ptr) * 2) - 1) ' Get the requested data line from symptextdata.dat file
                    FindAll.FindAllSympData.Add(New FindAllSympDat(LineDat1, LineDat2))
                Else
                    decrSTDCount += 1
                End If
            Next
            ' Associate FindAll.Lst1 with FindAllSympData
            FindAll.Lst1.DataSource = Nothing ' Remove and re-associate database in order to refresh Lst1
            FindAll.Lst1.DataSource = FindAll.FindAllSympData

            FindAll.Lst1.DisplayMember = "SympDesc"
            FindAll.Lst1.ValueMember = "SympData"
            FindAll.Lst1.ClearSelected()
            ' Now select any symptoms that are selected in Form1.Lst1
            If decrSTDCount > 0 Then STDCount -= decrSTDCount   ' Necessary to prevent FindAllSympData member out of bounds due to 0-value SympsToDisplay member
            For idxPtr = 0 To STDCount - 1
                questprogressSymptomID = Form1.GetSymptomID(FindAll.FindAllSympData(idxPtr).SympData.ToString)
                If Form1.IsSymptomSelected(Form1.MapSymIDToIndex(questprogressSymptomID)) Then
                    FindAll.Lst1.SetSelected(idxPtr, True)
                End If
                If Form1.RaiseException Then Throw New System.Exception("An exception has occurred.")
            Next

            ' Clear any symptom that may have been shown in FindAll TextBox1 as a result of adding symptoms.
            FindAll.TextBox1.Text = ""
        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
             Buttons:=vbOKOnly + vbCritical,
            Title:="ERROR!")
            Close()
        End Try
    End Sub

    Public Sub FindWords(LocWordArray() As String)  ' Finds all words possible in LocWordArray, saves their pointer strings in .
        Dim Done As Boolean     ' done searching file flag
        Dim DoneWithWords As Boolean    ' done with all words in text flag
        Dim Duplicate As Boolean    ' duplicate symptom pointer flag
        Dim ErrStyle As Integer ' error message format
        Dim LocStrsToDisplay As String  ' local version of string to be added to StrsToDisplay()
        Dim LocWordArraySize As Integer ' # of elements in array
        Dim MinNumber As Integer    ' Minimum number matched pointers needed.
        Dim msg As String       ' error message string
        Dim NumMatches As Integer
        Dim NumPointers(10000) As Integer    ' number of symptom pointers for word in WordSymptoms array
        Dim ptr, Ptr1, Ptr2 As Integer ' loop counters
        Dim SynIndx As String    ' synonym index string
        Dim SynPtr As Integer   ' SynWord element pointer

        Dim ThesFound As Boolean   ' text was found in list synonyms list
        Dim ThesWord As String  ' search word from thestxtidx file
        Dim TmpInt As Integer   ' interim symptom pointer, integer format
        Dim WordFound As Boolean   ' text was found in word list
        Dim WordPtr1, WordPtr2 As Integer  ' LocWordArray element pointers
        Dim WordSymptoms(5000, 10000)    ' 1st dim is word #, 2nd dim is symptom pointer

        ' Searches for words contained in LocWordArray(), in Thesaurus and word list.  Gets pointers to symptoms from these.  If the same pointers
        ' are found for a sufficient number of words or their synonyms, adds those pointers to a "symptoms to display" array.  That array is then
        ' used to populate the "Find All" symptom list so the user can select from these symptoms.
        '
        ' Algorithm:
        ' For each word in LocWordArray()
        ' If not (the first word and requirement set to match first word exactly), binary search the file thestxtidx.dat words contained in ThesWords() array
        ' Format for thestxtidx.dat (and ThesWords() array:  a comma-delimited line of the word to match, followed by a list of synonyms; the next line is
        ' a comma-delimited list of word pointers associated with each thes word.  If the word to match is not contained in any symptom, its pointer is -1.
        ' If a match is found, it copies the line of word pointers pointed to by loc_syfile_pos.
        ' These word pointers point to a line in file seartxtidx.dat.  This file contains a line with the symptom word to match, followed by a comma-
        ' delimited line of pointers to symptoms.  This line of pointers to symptoms is copied to variable SympPtrStr.
        ' If the word was not found in Thes.dat (or if this is the first word and there is a requirement set to match first word exactly), it is necessary
        ' to search for the words in file seartxtidx.dat.  If a match is found, it copies the line of pointers to symptoms to variable SympPtrStr.
        ' At this point, whether the word was found in the thesaurus or not, the variable SympPtrStr will contain a comma-delimited line of pointers to symptoms,
        ' if the word was found.
        ' Now for each pointer in SympPtrStr, convert it to integer and add it to the WordSymptoms() array (if it is not already in this array for the
        ' particular LocWordArray() word).  The first diminsion of this array is the number of the LocWordArray() word the pointers are associated with;
        ' the second dimension is the SymPtrStr pointers.
        ' This process is continued for all the words in LocWordArray().  When this is completed for all the words, the WordSymptoms() array is checked for
        ' symptom pointers that match the criteria for minumum number and / or pointer contained in first word.  Symptom pointers that meet these criteria
        ' are copied to the SympsToDisplay() array if they are not in that array already.  Also the string that will appear before the symptom description
        ' in the "FindAll" symptom list is generated and stored in the StrsToDisplay() array.  Next the SympsToDisplay() pointers are sorted numerically (and the 
        ' StrsToDisplay() array is manipulated to match the sorted SympStoDisplay() array).  Sub AddList() then needs to be called to populate the
        ' symptom list and display the "FindAll" dialog.

        Try
            ' LocWordArray contains the list of significant words to search for
            ' For each word in LocWordArray, the SynWord array will contain a list of synonyms (if any)
            Form1.ErrorFlag = 8084

            LocWordArraySize = UBound(LocWordArray) + 1

            WordPtr1 = 0
            DoneWithWords = False
            ThesFound = False
            SympPtrStr = ""     ' String of pointers for each word and synonyms
            SympPointers = ""   ' String of all symptom pointers for phrase
            WordPtrStr = ""     ' String of pointers for each symptom

            While (WordPtr1 <= LocWordArraySize And Not DoneWithWords) ' while more words to process
                Form1.ErrorFlag = 8085
                SynIndx = ""
                ThesWord = ""
                SynPtr = 0
                GIPtr = 0

                ' First search in thestxtidx.dat; if word found, find syn; if
                ' not, look for word in seartxtidx.dat (not all words in seartxtidx.dat are in thestxtidx.dat).  Note:
                ' if a word in thestxtidx.dat is not in a symptom (and thus in seartxtidx.dat), it's index will be -1.

                ' Set up binary search parameters; and only parse synonyms if not (1st word and must contain exact match of 1st word)
                If LocWordArray(WordPtr1) <> "" And Not ((WordPtr1 = 0) And (RadioButton2.Checked = True)) Then
                    First = 0
                    Last = Form1.Syfile_size / 2   ' File contains interleaved words plus indexes
                    If First < 0 Or First >= Last Then First = 0

                    ErrStyle = vbOKOnly + vbCritical

                    ' Look for word in data read from thestextidx.dat contained in ThesWords().
                    ThesFound = False
                    Done = False
                    While Not Done  ' process word and synonyms
                        loc_syfile_pos = (First + Last + 0.1) / 2
                        ThesWord = ThesWords(loc_syfile_pos)

                        If Last - First <= 1 Then
                            If (First >= 0) Then
                                loc_syfile_pos = First
                                ' Get thes file line containing symptom word and synonyms.
                                ThesWord = ThesWords(loc_syfile_pos)

                                If LocWordArray(WordPtr1) = Mid(ThesWord, 1, Len(LocWordArray(WordPtr1))) Then
                                    ThesFound = True
                                ElseIf Last <> First Then
                                    loc_syfile_pos = Last
                                    ThesWord = ThesWords(loc_syfile_pos)

                                    If LocWordArray(WordPtr1) = Mid(ThesWord, 1, Len(LocWordArray(WordPtr1))) Then
                                        ThesFound = True
                                    End If
                                End If
                            End If
                            Done = True
                        End If

                        If Not Done Then
                            Dim LetterPtr As Integer
                            LetterPtr = InStr(ThesWord, " ")
                            If LetterPtr = 0 Then LetterPtr = Len(ThesWord)
                            If LocWordArray(WordPtr1) >= Mid(ThesWord, 1, LetterPtr) Then
                                First = (First + Last + 0.1) / 2
                            Else
                                Last = (First + Last + 0.1) / 2
                            End If

                        End If

                    End While

                    If ThesFound Then
                        Form1.ErrorFlag = 8086
                        SynIndx = 0
                        WordPtrStr = ThesPointers(loc_syfile_pos)   ' Pointers from thestxtidx.dat
                        ' Thes file contains a word pointer for the original word and each of its synonyms; these point to the word record in seartxtidx.dat.
                        ' Need to get the associated symptom pointers and place them in SympPtrStr.
                        SympPtrStr = ""
                        If WordPtrStr <> "" Then
                            If InStr(WordPtrStr, ",") > 0 Then
                                TmpInt = CInt(Strings.Left(WordPtrStr, InStr(WordPtrStr, ",") - 1)) ' Get word pointer, convert to int
                            Else
                                TmpInt = CInt(WordPtrStr)
                            End If
                            While WordPtrStr <> "" And TmpInt >= 0
                                If SympPtrStr <> "" Then
                                    SympPtrStr += ", " + Form1.GlobalSearchData((2 * TmpInt) + 1)
                                Else
                                    SympPtrStr = Form1.GlobalSearchData((2 * TmpInt) + 1)
                                End If
                                If InStr(WordPtrStr, ",") > 0 Then
                                    WordPtrStr = WordPtrStr.Substring(InStr(WordPtrStr, ", ") + 1)  ' Remove the pointer from WordPtrStr
                                    If InStr(WordPtrStr, ",") > 0 Then
                                        TmpInt = CInt(Strings.Left(WordPtrStr, InStr(WordPtrStr, ",") - 1)) ' Get word pointer, convert to int
                                    Else
                                        TmpInt = CInt(WordPtrStr)
                                    End If
                                Else
                                    WordPtrStr = ""
                                End If
                            End While

                        End If
                    End If
                End If

                ' If not found or 1st word and exact match required
                If Not ThesFound Then
                    ' Need to search in Word list.
                    WordFound = False
                    First = 0
                    Done = False
                    Last = Form1.Sfile_size / 2   ' File contains interleaved words plus indexes
                    ' Find text position in file.
                    While Not Done
                        loc_sfile_pos = (First + Last + 0.1) / 2
                        If (loc_sfile_pos < 1) Then loc_sfile_pos = 1
                        If (loc_sfile_pos >= (Form1.Sfile_size - Find.List1.Items.Count)) Then _
                                                        loc_sfile_pos = Form1.Sfile_size - (Find.List1.Items.Count + 1)

                        Form1.ErrorFlag = 8087
                        If (loc_sfile_pos > Form1.Sfile_size Or loc_sfile_pos < 1) Then Last = Last / 0 ' Force it to "catch"
                        Call GetSearfileWord()  ' Copy indexes line (pointed to by loc_sfile_pos) from word file Seartxtidx.dat to variable SympWord.

                        If Last - First <= 1 Then
                            If (First > 0) Then
                                loc_sfile_pos = First
                                Call GetSearfileWord()

                                If LocWordArray(WordPtr1) = SympWord Then
                                    WordFound = True
                                End If
                                If Not WordFound And Last - First = 1 Then
                                    loc_sfile_pos = First + 1
                                    Call GetSearfileWord()

                                    If LocWordArray(WordPtr1) = SympWord Then
                                        WordFound = True
                                    End If
                                End If
                            End If
                            Done = True
                            SympPtrStr = Form1.GlobalSearchData((2 * loc_sfile_pos) + 1)
                        End If

                        If Not Done Then
                            If LocWordArray(WordPtr1) >= SympWord Then
                                First = (First + Last + 0.1) / 2
                            Else
                                Last = (First + Last + 0.1) / 2
                            End If
                        End If
                    End While

                End If

                Ptr1 = 0
                While InStr(SympPtrStr, ",") > 0 And NumPointers(WordPtr1) < 10000
                    TmpInt = CInt(Strings.Left(SympPtrStr, InStr(SympPtrStr, ",") - 1)) ' Get pointer, convert to int
                    SympPtrStr = SympPtrStr.Substring(InStr(SympPtrStr, ", ") + 1)  ' Remove the pointer from SympPtrStr
                    ' Check for duplicate with previously saved pointers for that word
                    Duplicate = False
                    If Not ((WordPtr1 = 0) And (RadioButton2.Checked = True)) Then    ' No need to check since thes not used.
                        Ptr2 = 0
                        While Duplicate = False And Ptr2 < NumPointers(WordPtr1)
                            If WordSymptoms(WordPtr1, Ptr2) = TmpInt Then Duplicate = True
                            Ptr2 += 1
                        End While
                    End If

                    If Duplicate = False Then
                        WordSymptoms(WordPtr1, NumPointers(WordPtr1)) = TmpInt
                        NumPointers(WordPtr1) += 1
                    End If

                End While

                ' Now add the last pointer in SympPtrStr if it isn't a duplicate.
                If SympPtrStr <> "" Then TmpInt = CInt(SympPtrStr)
                ' Check for duplicate with previously saved pointers for that word
                Duplicate = False
                If Not ((WordPtr1 = 0) And (RadioButton2.Checked = True)) Then    ' No need to check since thes not used.
                    Ptr2 = 0
                    While Duplicate = False And Ptr2 < NumPointers(WordPtr1)
                        If WordSymptoms(WordPtr1, Ptr2) = TmpInt Then Duplicate = True
                        Ptr2 += 1
                    End While
                End If

                If Duplicate = False Then
                    WordSymptoms(WordPtr1, NumPointers(WordPtr1)) = TmpInt
                    NumPointers(WordPtr1) += 1
                End If

                WordPtr1 += 1
                If WordPtr1 = (UBound(LocWordArray) + 1) Or NumPointers(WordPtr1) >= 10000 Then DoneWithWords = True ' Finished with all the words, force loop exit.
            End While

            ' Determine which symptoms to display (if any) based on user selections.
            ' Calculate number matches required to display symptom based on # significant words and required percentage.
            ' Must match at least 1 word.
            ' Here, MinNumber will get rounded down in case it's not an even integer (the decimal part is removed).
            MinNumber = CInt((NumWords) * PercentageRequired)
            If MinNumber = 0 Then MinNumber = 1 ' Need at least one pointer.

            ' Build the string to be added to StrsToDisplay()
            LocStrsToDisplay = PageNumber + " " + QuestionNumber + " " + Str(Form1.QuestWeight) + " "

            ' 1st dim is WordPtr; will range from 0 to NumPointers(WordPtr) - 1.
            ' 2nd dim is the symptom pointers for each word; may contain duplicates due to synonyms; need to check for these, set dups to -2.
            ' For each element in the word, check for duplicates in the other words.

            ' If symptoms must contain 1st word OR 1st word and its synonyms
            If (RadioButton2.Checked Or RadioButton3.Checked Or Form1.CalledFromShowSymSymps) Then
                ' If MinNumber is 1, that says display symptoms for every pointer, no need to check for matches.
                ' And since we must use pointers from 1st word, no need to check pointers from any other word(s).
                WordPtr1 = 0    ' Only include matches if also contained in 1st word (or syns)
                If MinNumber = 1 Then
                    For Ptr1 = 0 To NumPointers(WordPtr1)  ' For each symptom pointer in 1st word and its synonyms (if any)
                        ' Check for duplicate, only add if it is not a duplicate.
                        Duplicate = False
                        For ptr = 0 To STDCount
                            If WordSymptoms(WordPtr1, Ptr1) = SympsToDisplay(ptr) Then Duplicate = True
                        Next
                        If Not Duplicate And (WordSymptoms(WordPtr1, Ptr1) >= 0) Then
                            SympsToDisplay(STDCount) = WordSymptoms(WordPtr1, Ptr1)
                            StrsToDisplay(STDCount) = LocStrsToDisplay
                            STDCount += 1
                        End If
                    Next
                Else    ' More than one word required to point to each symptom AND must use 1st word (WordSymptoms(0, x)

                    ' Duplicates were already checked for in each word / syn group, no need to do it here.
                    For Ptr1 = 0 To NumPointers(WordPtr1)  ' For each symptom pointer in 1st word and its synonyms (if any)
                        NumMatches = 1  ' Count the first instance.
                        For WordPtr2 = WordPtr1 + 1 To NumWords - 1 ' for groups of symptom pointers for each word
                            For Ptr2 = 0 To NumPointers(WordPtr2)   ' for each symptom pointer in group
                                If (WordSymptoms(WordPtr1, Ptr1) > 0) And (WordSymptoms(WordPtr1, Ptr1) = WordSymptoms(WordPtr2, Ptr2)) Then
                                    NumMatches += 1
                                    WordSymptoms(WordPtr2, Ptr2) = -2   ' So it won't get counted twice
                                    If NumMatches >= MinNumber Then
                                        ' Check for duplicate (from other fields), only add if it is not a duplicate.
                                        Duplicate = False
                                        For ptr = 0 To STDCount
                                            If WordSymptoms(WordPtr1, Ptr1) = SympsToDisplay(ptr) Then Duplicate = True
                                        Next
                                        If Not Duplicate And (WordSymptoms(WordPtr1, Ptr1) >= 0) Then
                                            SympsToDisplay(STDCount) = WordSymptoms(WordPtr1, Ptr1)
                                            StrsToDisplay(STDCount) = LocStrsToDisplay
                                            STDCount += 1
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    Next

                End If
            Else    ' Don't necessarily need to include 1st word or its synonyms.
                ' If MinNumber is 1, that says display symptoms for every pointer, no need to check for matches.
                ' And since we don't necessarily need to use pointers from 1st word, need to add all pointers from all words.
                If MinNumber = 1 Then
                    For WordPtr1 = 0 To NumWords - 1
                        For Ptr1 = 0 To NumPointers(WordPtr1)
                            ' Check for duplicate, only add if it is not a duplicate.
                            Duplicate = False
                            For ptr = 0 To STDCount
                                If WordSymptoms(WordPtr1, Ptr1) = SympsToDisplay(ptr) Then Duplicate = True
                            Next
                            If Not Duplicate And (WordSymptoms(WordPtr1, Ptr1) >= 0) Then
                                SympsToDisplay(STDCount) = WordSymptoms(WordPtr1, Ptr1)
                                StrsToDisplay(STDCount) = LocStrsToDisplay
                                STDCount += 1
                            End If
                        Next
                    Next
                Else    ' More than one word required to point to each symptom
                    ' Duplicates were already checked for in each word, no need to do it here.
                    For WordPtr1 = 0 To NumWords - 1
                        For Ptr1 = 0 To NumPointers(WordPtr1)
                            NumMatches = 1  ' Count the first instance.
                            For WordPtr2 = WordPtr1 + 1 To NumWords - 1 ' for each symptom pointer in group
                                For Ptr2 = 0 To NumPointers(WordPtr2)
                                    If (WordSymptoms(WordPtr1, Ptr1) > 0) And (WordSymptoms(WordPtr1, Ptr1) = WordSymptoms(WordPtr2, Ptr2)) Then
                                        NumMatches += 1
                                        WordSymptoms(WordPtr2, Ptr2) = -2   ' So it won't get counted twice                                       WordSymptoms(WordPtr2, Ptr2) = -2   ' So it won't get counted twice
                                        If NumMatches >= MinNumber Then
                                            ' Check for duplicate (from other fields), only add if it is not a duplicate.
                                            Duplicate = False
                                            For ptr = 0 To STDCount
                                                If WordSymptoms(WordPtr1, Ptr1) = SympsToDisplay(ptr) Then Duplicate = True
                                            Next
                                            If Not Duplicate And (WordSymptoms(WordPtr1, Ptr1) >= 0) Then
                                                SympsToDisplay(STDCount) = WordSymptoms(WordPtr1, Ptr1)
                                                StrsToDisplay(STDCount) = LocStrsToDisplay
                                                STDCount += 1
                                            End If
                                        End If
                                    End If
                                Next
                            Next
                        Next

                    Next
                End If
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

    Public Sub GetSearfileWord()
        ' This subroutine copies the word from word file
        ' Seartxtidx.dat specified by variable loc_sfile_pos
        ' to variable SympWord.  It also does some
        ' range checking of the file pointers.
        Dim msg As String     ' error message string

        Try
            Form1.ErrorFlag = 8088

            If (loc_sfile_pos < 1) Then loc_sfile_pos = 1
            If (loc_sfile_pos >= (Form1.Sfile_size)) Then _
            loc_sfile_pos = Form1.Sfile_size - 1

            If (loc_sfile_pos > Form1.Sfile_size Or loc_sfile_pos < 1) Then Throw New System.Exception("An exception has occurred.") ' Force it to "catch"

            Form1.ErrorFlag = 8089
            SympWord = Find.searData(loc_sfile_pos).SearDataName.ToString
            Form1.ErrorFlag = 8090

            Form1.ErrorFlag = 8091
            If (Form1.SCombSympPtr > 8 * Form1.SMaxSympPtr Or Form1.SCombSympPtr < 1) Then Last = Last / 0 ' Force it to "catch"

            Exit Sub

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            Close()
        End Try
    End Sub

    Public Sub GetSigWord(FieldText As String)
        ' Returns significant words from Sentence in global array WordArray().

        Dim msg As String
        Dim MoreToDo As Boolean     ' loop exit flag
        Dim Ptr As Integer          ' string pointer
        Dim W As String             ' word from the sentence

        Try
            Form1.ErrorFlag = 8092
            W = " "   ' Avoid warning variable used before assigning value

            NumWords = 0
            MoreToDo = True
            FieldText = LCase(FieldText)
            ' Delete the punctuation.
            For Ptr = 1 To Len(FieldText)
                If Mid(FieldText, Ptr, 1) = "," Or
                Mid(FieldText, Ptr, 1) = ";" Or
                Mid(FieldText, Ptr, 1) = ":" Or
                Mid(FieldText, Ptr, 1) = "(" Or
                Mid(FieldText, Ptr, 1) = ")" Or
                Mid(FieldText, Ptr, 1) = """" Or
                Mid(FieldText, Ptr, 1) = "''" Or
                Mid(FieldText, Ptr, 1) = "." Or
                Mid(FieldText, Ptr, 1) = "?" Or
                Mid(FieldText, Ptr, 1) = "!" Then
                    FieldText = Mid(FieldText, 1, Ptr - 1) + Mid(FieldText, Ptr + 1)
                End If
            Next Ptr

            ' Delete any and all trailing blanks
            While Mid(FieldText, 1, Len(FieldText)) = " "
                FieldText = Mid(FieldText, 1, Len(FieldText) - 1)
            End While
            While Ptr > 0
                ' Delete leading blanks - handles case where more than one blank between words.
                While Mid(FieldText, 1, 1) = " "
                    FieldText = Mid(FieldText, 2)
                End While

                Ptr = InStr(FieldText, " ")
                If Ptr > 0 Then
                    W = Mid(FieldText, 1, Ptr - 1)
                Else
                    W = FieldText
                End If
                If W <> "(" And
                    W <> ")" And
                    W <> "''" And
                        W <> """" And
                        W <> "a" And
                        W <> "an" And
                        W <> "and" And
                        W <> "as" And
                        W <> "at" And
                        W <> "be" And
                        W <> "do" And
                        W <> "I" And
                        W <> "if" And
                        W <> "in" And
                        W <> "is" And
                        W <> "it" And
                        W <> "me" And
                        W <> "not" And
                        W <> "of" And
                        W <> "oh" And
                        W <> "on" And
                        W <> "or" And
                        W <> "the" And
                        W <> "we" And
                        W <> "with" Then

                    If Not CheckForDuplicate(W) Then
                        NumWords += 1
                        ReDim Preserve WordArray(NumWords - 1)
                        WordArray(NumWords - 1) = W
                    End If
                End If
                FieldText = Mid(FieldText, Ptr + 1)
            End While
            Exit Sub

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            Close()
        End Try
    End Sub

    Private Function CheckForDuplicate(WordToCheck As String) As Boolean
        Dim Duplicate As Boolean    ' flag denoting a duplicate word
        Dim ptr As Integer

        Duplicate = False
        If NumWords > 0 Then    ' No need to check for duplicate if no words in array
            For ptr = 0 To NumWords - 1 ' NumWords is already incremented and array is zero-based, so need to subtract 2
                If WordToCheck = WordArray(ptr) Then Duplicate = True
            Next
        End If
        CheckForDuplicate = Duplicate
    End Function
    Private Sub QuestProgress_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim ErrorFlag As Integer    ' make it local to preserve it after calls
        Dim msg As String       ' error pop-up string
        Dim ptr As Integer
        Try
            ErrorFlag = 8093
            RadioButton3.Checked = True ' Default to Must contain first word or one of its synonyms
            RadioButton6.Checked = True ' Default to 75% match
            ProgressBar1.Value = 0

            ProgressBar1.Visible = False
            If Form1.SympSize = 1 Then
                Label2.Visible = False  ' "Select symptom(s) to search"
                Lst1.Visible = False    ' symptom selection list
            Else    ' More than one symptom, need to allow selection
                Lst1.Items.Clear()    ' Init list in case it was populated previously
                For ptr = 0 To Form1.SympSize - 1   ' Add number and titles of all symptoms to list box
                    Lst1.Items.Add((ptr + 1).ToString + Form1.SympData.Title(ptr))
                Next
                Label2.Visible = True  ' "Select symptom(s) to search"
                Lst1.Visible = True    ' symptom selection list
                Lst1.SetSelected(Form1.SympSize - 1, True)   ' Set the currently-displayed symptom
            End If
            STDCount = 0

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click   ' Start symptom search
        Dim CheckWord As String ' word associated with check box
        Dim ItmPtr As Integer   'pointer to page item
        Dim msg As String       ' error pop-up string
        Dim ptr1, ptr2 As Integer      ' loop pointers
        Dim RecPtr As Integer   ' Record structure pointer
        Dim SympPtr As Integer  ' loop pointer to symptom

        Try
            Label1.Text = "Starting search . . ."
            ProgressBar1.Visible = True

            Form1.SMaxSympPtr = Form1.SCombSympPtr

            Form1.ErrorFlag = 8094

            ' Copy seartxtidx.dat to 2 string arrays to make word searching easier.
            ReDim SearWords(Form1.Sfile_size / 2)
            ReDim SearPointers(Form1.Sfile_size / 2)
            Form1.CalledFromShowSymSymps = False

            ptr2 = 0
            For ptr1 = 0 To Form1.Sfile_size - 1 Step 2
                SearWords(ptr2) = Form1.GlobalSearchData(ptr1)
                SearPointers(ptr2) = Form1.GlobalSearchData(ptr1 + 1)
                ptr2 += 1
            Next

            ReDim WordArray(0)  ' Clear out any previous values, dim array to one element

            If RadioButton4.Checked Then
                PercentageRequired = 0.25
            ElseIf RadioButton5.Checked Then
                PercentageRequired = 0.5
            Else
                PercentageRequired = 0.75
            End If

            '   Assign weights to the symptoms based on check boxes on
            '   page 1.
            Label1.Text = "Assigning symptom weights . . ."
            Me.Refresh()

            SympPtr = Lst1.SelectedIndex  ' Get number of symptom currently selected.
            If SympPtr = -1 Then SympPtr = 0    ' Handle case where only one symptom so Lst1 wasn't displayed.
            If Form1.SympData.P1Cb(2, SympPtr) Then
                Form1.QuestWeight = 1
            ElseIf Form1.SympData.P1Cb(3, SympPtr) Then
                Form1.QuestWeight = 2
            ElseIf Form1.SympData.P1Cb(4, SympPtr) Then
                Form1.QuestWeight = 3
            Else
                Form1.QuestWeight = 0
            End If

            ' Read text and check boxes and parse the text or selected check box meanings, for each page.
            ' page 1
            PageNumber = "1/" + Str(SympPtr + 1)
            For ItmPtr = 0 To 6
                QuestionNumber = Str(ItmPtr + 1)
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                If Form1.SympData.P1(ItmPtr, SympPtr) <> "" Then
                    Call SymParse(Form1.SympData.P1(ItmPtr, SympPtr), Form1.QuestWeight)
                End If
                ProgressBar1.PerformStep()
            Next ItmPtr
            If Form1.SympData.P1Cb(0, SympPtr) Then
                QuestionNumber = "C1"
                CheckWord = "recurring"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            ' page 2
            PageNumber = "2/" + Str(SympPtr + 1)
            If Form1.SympData.P2(0, SympPtr) <> "" Then
                QuestionNumber = 1
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(Form1.SympData.P2(0, SympPtr) + " worsen", Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()

            If Form1.SympData.P2(1, SympPtr) <> "" Then
                QuestionNumber = 2
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(Form1.SympData.P2(1, SympPtr) + " improve", Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.SympData.P2(2, SympPtr) <> "" Then
                QuestionNumber = 3
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(Form1.SympData.P2(2, SympPtr) + " relieves", Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.SympData.P2(3, SympPtr) <> "" Then
                QuestionNumber = 4
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(Form1.SympData.P2(3, SympPtr), Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.SympData.P2(4, SympPtr) <> "" Then
                QuestionNumber = 5
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(Form1.SympData.P2(4, SympPtr) + " worst", Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.SympData.P2(5, SympPtr) <> "" Then
                QuestionNumber = 6
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse("cannot " + Form1.SympData.P2(5, SympPtr), Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.SympData.P2(6, SympPtr) <> "" Then
                QuestionNumber = 7
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse("must " + Form1.SympData.P2(6, SympPtr), Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()

            ' page 3
            PageNumber = "3/" + Str(SympPtr + 1)
            If Form1.SympData.P3(0, SympPtr) <> "" Then
                QuestionNumber = "1"
                CheckWord = "worst " + Form1.SympData.P3(0, SympPtr)
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.SympData.P3(1, SympPtr) <> "" Then
                QuestionNumber = "2"
                CheckWord = "best " + Form1.SympData.P3(1, SympPtr)
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.SympData.P3Cb(0, SympPtr) Then
                QuestionNumber = "C1"
                CheckWord = "worse winter"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            ElseIf Form1.SympData.P3Cb(1, SympPtr) Then
                QuestionNumber = "C2"
                CheckWord = "worse spring"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            ElseIf Form1.SympData.P3Cb(2, SympPtr) Then
                QuestionNumber = "C3"
                CheckWord = "worse summer"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            ElseIf Form1.SympData.P3Cb(3, SympPtr) Then
                QuestionNumber = "C4"
                CheckWord = "worse fall"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.SympData.P3Cb(4, SympPtr) Then
                QuestionNumber = "C5"
                CheckWord = "best winter"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            ElseIf Form1.SympData.P3Cb(5, SympPtr) Then
                QuestionNumber = "C6"
                CheckWord = "best spring"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            ElseIf Form1.SympData.P3Cb(6, SympPtr) Then
                QuestionNumber = "C7"
                CheckWord = "best summer"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            ElseIf Form1.SympData.P3Cb(7, SympPtr) Then
                QuestionNumber = "C8"
                CheckWord = "best fall"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.SympData.P3Cb(8, SympPtr) Then
                QuestionNumber = "C9"
                CheckWord = "cold improves"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            ElseIf Form1.SympData.P3Cb(9, SympPtr) Then
                QuestionNumber = "C10"
                CheckWord = "cold worsens"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.SympData.P3Cb(11, SympPtr) Then
                QuestionNumber = "C12"
                CheckWord = "heat improves or hot improves"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            ElseIf Form1.SympData.P3Cb(12, SympPtr) Then
                QuestionNumber = "C13"
                CheckWord = "heat worsens or hot worsens"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.SympData.P3Cb(14, SympPtr) Then
                QuestionNumber = "C15"
                CheckWord = "wet improves or damp improves"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            ElseIf Form1.SympData.P3Cb(15, SympPtr) Then
                QuestionNumber = "C16"
                CheckWord = "wet worsens or damp worsens"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.SympData.P3Cb(17, SympPtr) Then
                QuestionNumber = "C18"
                CheckWord = "dry improves"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            ElseIf Form1.SympData.P3Cb(18, SympPtr) Then
                QuestionNumber = "C19"
                CheckWord = "dry worsens"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.SympData.P3Cb(20, SympPtr) Then
                QuestionNumber = "C21"
                CheckWord = "fog improves"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            ElseIf Form1.SympData.P3Cb(21, SympPtr) Then
                QuestionNumber = "C22"
                CheckWord = "fog worsens"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.SympData.P3Cb(23, SympPtr) Then
                QuestionNumber = "C24"
                CheckWord = "rain improves"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            ElseIf Form1.SympData.P3Cb(24, SympPtr) Then
                QuestionNumber = "C25"
                CheckWord = "rain worsens"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.SympData.P3Cb(26, SympPtr) Then
                QuestionNumber = "C27"
                CheckWord = "thunder improves"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            ElseIf Form1.SympData.P3Cb(27, SympPtr) Then
                QuestionNumber = "C28"
                CheckWord = "thunder worsens"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.SympData.P3Cb(29, SympPtr) Then
                QuestionNumber = "C30"
                CheckWord = "snow improves"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            ElseIf Form1.SympData.P3Cb(30, SympPtr) Then
                QuestionNumber = "C31"
                CheckWord = "snow worsens"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.SympData.P3Cb(32, SympPtr) Then
                QuestionNumber = "C33"
                CheckWord = "sun improves"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            ElseIf Form1.SympData.P3Cb(33, SympPtr) Then
                QuestionNumber = "C34"
                CheckWord = "sun worsens"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.SympData.P3Cb(35, SympPtr) Then
                QuestionNumber = "C36"
                CheckWord = "wind improves"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            ElseIf Form1.SympData.P3Cb(36, SympPtr) Then
                QuestionNumber = "C37"
                CheckWord = "wind worsens"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.SympData.P3Cb(38, SympPtr) Then
                QuestionNumber = "C39"
                CheckWord = "warm room improves"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            ElseIf Form1.SympData.P3Cb(39, SympPtr) Then
                QuestionNumber = "C40"
                CheckWord = "warm room worsens"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.SympData.P3Cb(41, SympPtr) Then
                QuestionNumber = "C42"
                CheckWord = "cool room improves"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            ElseIf Form1.SympData.P3Cb(42, SympPtr) Then
                QuestionNumber = "C43"
                CheckWord = "cool room worsens"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.SympData.P3Cb(44, SympPtr) Then
                QuestionNumber = "C45"
                CheckWord = "draft room improves or draught room improves"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            ElseIf Form1.SympData.P3Cb(45, SympPtr) Then
                QuestionNumber = "C46"
                CheckWord = "draft room worsens or draught room worsens"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.SympData.P3Cb(47, SympPtr) Then
                QuestionNumber = "C48"
                CheckWord = "damp room improves"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            ElseIf Form1.SympData.P3Cb(48, SympPtr) Then
                QuestionNumber = "C49"
                CheckWord = "damp room worsens"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()

            ' page 4
            PageNumber = "4/" + Str(SympPtr + 1)
            If Form1.SympData.P4Cb(0, SympPtr) Then
                QuestionNumber = "C1"
                CheckWord = "worse standing"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            ElseIf Form1.SympData.P4Cb(1, SympPtr) Then
                QuestionNumber = "C2"
                CheckWord = "worse sitting"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            ElseIf Form1.SympData.P4Cb(2, SympPtr) Then
                QuestionNumber = "C3"
                CheckWord = "worse lying"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.SympData.P4(0, SympPtr) <> "" Then
                QuestionNumber = "1"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(Form1.SympData.P4(0, SympPtr) + " worsen", Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.SympData.P4(1, SympPtr) <> "" Then
                QuestionNumber = "2"
                Label1.Text = "Symptom " + Str(SympPtr + 1) + " Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(Form1.SympData.P4(1, SympPtr), Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            '                    End If
            '                End If

            ' Page 5
            PageNumber = "5/ "
            Form1.QuestWeight = 1
            If Form1.PatientData.P5AgeCb(0) Then
                QuestionNumber = "C1"
                CheckWord = "infant"
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
                If Form1.PatientData.P5SexCb(0) Then
                    QuestionNumber = "C6"
                    CheckWord = "boy"
                    Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                    Me.Refresh()
                    Call SymParse(CheckWord, Form1.QuestWeight)
                ElseIf Form1.PatientData.P5SexCb(1) Then
                    QuestionNumber = "C7"
                    CheckWord = "girl"
                    Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                    Me.Refresh()
                    Call SymParse(CheckWord, Form1.QuestWeight)
                End If
            ElseIf Form1.PatientData.P5AgeCb(1) Then
                QuestionNumber = "C2"
                CheckWord = "child"
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
                If Form1.PatientData.P5SexCb(0) Then
                    QuestionNumber = "C6"
                    CheckWord = "boy"
                    Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                    Me.Refresh()
                    Call SymParse(CheckWord, Form1.QuestWeight)
                ElseIf Form1.PatientData.P5SexCb(1) Then
                    QuestionNumber = "C7"
                    CheckWord = "girl"
                    Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                    Me.Refresh()
                    Call SymParse(CheckWord, Form1.QuestWeight)
                End If
            ElseIf Form1.PatientData.P5AgeCb(2) Then
                QuestionNumber = "C3"
                CheckWord = "adolescent"
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
                If Form1.PatientData.P5SexCb(0) Then
                    QuestionNumber = "C6"
                    CheckWord = "boy"
                    Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                    Me.Refresh()
                    Call SymParse(CheckWord, Form1.QuestWeight)
                ElseIf Form1.PatientData.P5SexCb(1) Then
                    QuestionNumber = "C7"
                    CheckWord = "girl"
                    Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                    Me.Refresh()
                    Call SymParse(CheckWord, Form1.QuestWeight)
                End If
            ElseIf Form1.PatientData.P5AgeCb(3) Then
                QuestionNumber = "C4"
                CheckWord = "adult"
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
                If Form1.PatientData.P5SexCb(0) Then
                    QuestionNumber = "C6"
                    CheckWord = "man"
                    Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                    Me.Refresh()
                    Call SymParse(CheckWord, Form1.QuestWeight)
                ElseIf Form1.PatientData.P5SexCb(1) Then
                    QuestionNumber = "C7"
                    CheckWord = "woman"
                    Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                    Me.Refresh()
                    Call SymParse(CheckWord, Form1.QuestWeight)
                End If
            ElseIf Form1.PatientData.P5AgeCb(4) Then
                QuestionNumber = "C5"
                CheckWord = "senior"
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
                If Form1.PatientData.P5SexCb(0) Then
                    QuestionNumber = "C6"
                    CheckWord = "man"
                    Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                    Me.Refresh()
                    Call SymParse(CheckWord, Form1.QuestWeight)
                ElseIf Form1.PatientData.P5SexCb(1) Then
                    QuestionNumber = "C7"
                    CheckWord = "woman"
                    Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                    Me.Refresh()
                    Call SymParse(CheckWord, Form1.QuestWeight)
                End If
            End If
            ProgressBar1.PerformStep()
            If Form1.PatientData.P5AgeCb(0) Then
                QuestionNumber = "C6"
                CheckWord = "male"
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            ElseIf Form1.PatientData.P5AgeCb(1) Then
                QuestionNumber = "C7"
                CheckWord = "female"
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(CheckWord, Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.PatientData.P5(0) <> "" Then
                QuestionNumber = "1"
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(Form1.PatientData.P5(0), Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.PatientData.P5(1) <> "" Then
                QuestionNumber = "2"
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(Form1.PatientData.P5(1) + " dislikes", Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.PatientData.P5(2) <> "" Then
                QuestionNumber = "3"
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(Form1.PatientData.P5(2) + " aversion to", Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.PatientData.P5(3) <> "" Then
                QuestionNumber = "4"
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse("craves " + Form1.PatientData.P5(3), Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.PatientData.P5(4) <> "" Then
                QuestionNumber = "5"
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse("habitual " + Form1.PatientData.P5(4), Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()
            If Form1.PatientData.P5(5) <> "" Then
                QuestionNumber = "6"
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse(Form1.PatientData.P5(5) + " disagrees", Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()

            ' Page 6
            PageNumber = "6/ "
            For ItmPtr = 0 To 5
                RecPtr = 3 * (ItmPtr - 1)
                If Form1.PatientData.P6(ItmPtr) <> "" Then
                    If Form1.PatientData.P6Cb(ItmPtr) Then
                        Form1.QuestWeight = 3
                    ElseIf Form1.PatientData.P6Cb(ItmPtr + 1) Then
                        Form1.QuestWeight = 2
                    ElseIf Form1.PatientData.P6Cb(ItmPtr + 2) Then
                        Form1.QuestWeight = 1
                    Else
                        Form1.QuestWeight = 0
                    End If

                    QuestionNumber = Str(ItmPtr + 1)
                    Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                    Me.Refresh()
                    Call SymParse(Form1.PatientData.P6(ItmPtr), Form1.QuestWeight)
                End If
                ProgressBar1.PerformStep()
            Next ItmPtr

            ' Page 7
            PageNumber = "7/ "
            RecPtr = 0
            If Form1.PatientData.P7(0) <> "" Then
                If Form1.PatientData.P7Cb(0) Then
                    Form1.QuestWeight = 3
                ElseIf Form1.PatientData.P7Cb(1) Then
                    Form1.QuestWeight = 2
                ElseIf Form1.PatientData.P7Cb(2) Then
                    Form1.QuestWeight = 1
                Else
                    Form1.QuestWeight = 0
                End If
                QuestionNumber = Str(1)
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse("Head, " + Form1.PatientData.P7(0), Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()

            RecPtr = 1
            If Form1.PatientData.P7(1) <> "" Then
                If Form1.PatientData.P7Cb(3) Then
                    Form1.QuestWeight = 3
                ElseIf Form1.PatientData.P7Cb(4) Then
                    Form1.QuestWeight = 2
                ElseIf Form1.PatientData.P7Cb(5) Then
                    Form1.QuestWeight = 1
                Else
                    Form1.QuestWeight = 0
                End If
                QuestionNumber = Str(2)
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse("Eyes, " + Form1.PatientData.P7(1), Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()

            RecPtr = 2
            If Form1.PatientData.P7(2) <> "" Then
                If Form1.PatientData.P7Cb(6) Then
                    Form1.QuestWeight = 3
                ElseIf Form1.PatientData.P7Cb(7) Then
                    Form1.QuestWeight = 2
                ElseIf Form1.PatientData.P7Cb(8) Then
                    Form1.QuestWeight = 1
                Else
                    Form1.QuestWeight = 0
                End If
                QuestionNumber = Str(3)
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse("Ears, " + Form1.PatientData.P7(2), Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()

            RecPtr = 3
            If Form1.PatientData.P7(3) <> "" Then
                If Form1.PatientData.P7Cb(9) Then
                    Form1.QuestWeight = 3
                ElseIf Form1.PatientData.P7Cb(10) Then
                    Form1.QuestWeight = 2
                ElseIf Form1.PatientData.P7Cb(11) Then
                    Form1.QuestWeight = 1
                Else
                    Form1.QuestWeight = 0
                End If
                QuestionNumber = Str(4)
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse("Nose, " + Form1.PatientData.P7(3), Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()

            RecPtr = 4
            If Form1.PatientData.P7(4) <> "" Then
                If Form1.PatientData.P7Cb(12) Then
                    Form1.QuestWeight = 3
                ElseIf Form1.PatientData.P7Cb(13) Then
                    Form1.QuestWeight = 2
                ElseIf Form1.PatientData.P7Cb(14) Then
                    Form1.QuestWeight = 1
                Else
                    Form1.QuestWeight = 0
                End If
                QuestionNumber = Str(5)
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse("Mouth, " + Form1.PatientData.P7(4), Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()

            RecPtr = 5
            If Form1.PatientData.P7(5) <> "" Then
                If Form1.PatientData.P7Cb(15) Then
                    Form1.QuestWeight = 3
                ElseIf Form1.PatientData.P7Cb(16) Then
                    Form1.QuestWeight = 2
                ElseIf Form1.PatientData.P7Cb(17) Then
                    Form1.QuestWeight = 1
                Else
                    Form1.QuestWeight = 0
                End If
                QuestionNumber = Str(6)
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse("Face, " + Form1.PatientData.P7(5), Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()

            RecPtr = 6
            If Form1.PatientData.P7(6) <> "" Then
                If Form1.PatientData.P7Cb(18) Then
                    Form1.QuestWeight = 3
                ElseIf Form1.PatientData.P7Cb(19) Then
                    Form1.QuestWeight = 2
                ElseIf Form1.PatientData.P7Cb(20) Then
                    Form1.QuestWeight = 1
                Else
                    Form1.QuestWeight = 0
                End If
                QuestionNumber = Str(7)
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse("Tongue, " + Form1.PatientData.P7(6), Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()

            RecPtr = 7
            If Form1.PatientData.P7(7) <> "" Then
                If Form1.PatientData.P7Cb(21) Then
                    Form1.QuestWeight = 3
                ElseIf Form1.PatientData.P7Cb(22) Then
                    Form1.QuestWeight = 2
                ElseIf Form1.PatientData.P7Cb(23) Then
                    Form1.QuestWeight = 1
                Else
                    Form1.QuestWeight = 0
                End If
                QuestionNumber = Str(8)
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse("Throat, " + Form1.PatientData.P7(7), Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()

            RecPtr = 8
            If Form1.PatientData.P7(8) <> "" Then
                If Form1.PatientData.P7Cb(24) Then
                    Form1.QuestWeight = 3
                ElseIf Form1.PatientData.P7Cb(25) Then
                    Form1.QuestWeight = 2
                ElseIf Form1.PatientData.P7Cb(26) Then
                    Form1.QuestWeight = 1
                Else
                    Form1.QuestWeight = 0
                End If
                QuestionNumber = Str(9)
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse("Stomach, " + Form1.PatientData.P7(8), Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()

            RecPtr = 9
            If Form1.PatientData.P7(9) <> "" Then
                If Form1.PatientData.P7Cb(27) Then
                    Form1.QuestWeight = 3
                ElseIf Form1.PatientData.P7Cb(28) Then
                    Form1.QuestWeight = 2
                ElseIf Form1.PatientData.P7Cb(29) Then
                    Form1.QuestWeight = 1
                Else
                    Form1.QuestWeight = 0
                End If
                QuestionNumber = Str(10)
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse("Abdomen, " + Form1.PatientData.P7(9), Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()

            RecPtr = 10
            If Form1.PatientData.P7(10) <> "" Then
                If Form1.PatientData.P7Cb(30) Then
                    Form1.QuestWeight = 3
                ElseIf Form1.PatientData.P7Cb(31) Then
                    Form1.QuestWeight = 2
                ElseIf Form1.PatientData.P7Cb(32) Then
                    Form1.QuestWeight = 1
                Else
                    Form1.QuestWeight = 0
                End If
                QuestionNumber = Str(11)
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse("Bowels, " + Form1.PatientData.P7(10), Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()

            RecPtr = 11
            If Form1.PatientData.P7(11) <> "" Then
                If Form1.PatientData.P7Cb(33) Then
                    Form1.QuestWeight = 3
                ElseIf Form1.PatientData.P7Cb(34) Then
                    Form1.QuestWeight = 2
                ElseIf Form1.PatientData.P7Cb(35) Then
                    Form1.QuestWeight = 1
                Else
                    Form1.QuestWeight = 0
                End If
                QuestionNumber = Str(12)
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse("Urination, " + Form1.PatientData.P7(11), Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()

            RecPtr = 12
            If Form1.PatientData.P7(12) <> "" Then
                If Form1.PatientData.P7Cb(36) Then
                    Form1.QuestWeight = 3
                ElseIf Form1.PatientData.P7Cb(37) Then
                    Form1.QuestWeight = 2
                ElseIf Form1.PatientData.P7Cb(38) Then
                    Form1.QuestWeight = 1
                Else
                    Form1.QuestWeight = 0
                End If
                QuestionNumber = Str(13)
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse("Neck, " + Form1.PatientData.P7(12), Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()

            RecPtr = 13
            If Form1.PatientData.P7(13) <> "" Then
                If Form1.PatientData.P7Cb(39) Then
                    Form1.QuestWeight = 3
                ElseIf Form1.PatientData.P7Cb(40) Then
                    Form1.QuestWeight = 2
                ElseIf Form1.PatientData.P7Cb(41) Then
                    Form1.QuestWeight = 1
                Else
                    Form1.QuestWeight = 0
                End If
                QuestionNumber = Str(14)
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse("Back, " + Form1.PatientData.P7(13), Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()

            RecPtr = 14
            If Form1.PatientData.P7(14) <> "" Then
                If Form1.PatientData.P7Cb(42) Then
                    Form1.QuestWeight = 3
                ElseIf Form1.PatientData.P7Cb(43) Then
                    Form1.QuestWeight = 2
                ElseIf Form1.PatientData.P7Cb(44) Then
                    Form1.QuestWeight = 1
                Else
                    Form1.QuestWeight = 0
                End If
                QuestionNumber = Str(15)
                Label1.Text = "Page " + Strings.Left(PageNumber, 1) + " Question " + QuestionNumber + " . . ."
                Me.Refresh()
                Call SymParse("Limbs, " + Form1.PatientData.P7(14), Form1.QuestWeight)
            End If
            ProgressBar1.PerformStep()

            Cursor = Cursors.Default    ' Change cursor to default symbol.

            If STDCount > 0 Then
                Call AddList()  ' Show FindAll, populate list with found symptoms.
            End If

            ' Handle case where no matches were found.
            If Not Form1.FindAllVisible Then
                '       Re-display questionaire 7.
                MsgBox("No symptom matches were found.", vbOKOnly + vbInformation)
                Questionnaire.TabPage7.Show()
            End If

            Close()
            Exit Sub

        Catch
            Cursor = Cursors.Default    ' Change cursor to default symbol.
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click   ' Cancel
        Close()
    End Sub

    Private Sub ProgressBar1_Click(sender As Object, e As EventArgs) Handles ProgressBar1.Click

    End Sub
End Class