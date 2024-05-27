Option Explicit On
Imports System.IO
Public Class Find
    Public First As Integer    ' top of binary search
    Public dontHandleSelIdxChg As Boolean   ' suppress selected index changed handler for List1
    Public idxArray() As Integer    ' array of all indexes in idsLineDat
    Public idxLineDat As String ' index string line data
    Public idxPtr As Integer    ' pointer to indexArray element
    Public Last As Integer     ' bottom of binary search
    Public locLineDat As String     ' selected list item
    Public loc_file_pos As Long    ' local symptom file position
    Public numIndexes As Integer    ' number of elements in indesArray
    Public Ptr1, Ptr2 As Long   ' string pointers
    Public StrSeartxtidx As String  ' contents of file Seartxtidx.dat
    Public SupTxtChgHnd As Boolean  ' used to suppress action of text change handler if software updated text box
    Public searData As New ArrayList()

    Private Sub Find_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim ErrorFlag As Integer    ' error number
        Dim msg As String   ' error message string

        Try
            If Form1.FindLoaded = False Then
                ErrorFlag = 3001
                List1.Visible = True

                Text1.BackColor = Form1.TextBox1.BackColor
                Text1.ForeColor = Form1.TextBox1.ForeColor
                Text1.Font = Form1.TextBox1.Font

                List1.BackColor = Form1.Lst1.BackColor
                List1.ForeColor = Form1.Lst1.ForeColor
                List1.Font = Form1.Lst1.Font

                FindNext.Enabled = False
                FindPrev.Enabled = False
                Form1.ErrorFlag = 3002
                Check1.Checked = 0
                CancelFindButton.Enabled = True
                CancelFindButton.Text = "Cancel"

                Form1.ErrorFlag = 3003

                dontHandleSelIdxChg = True  ' Suppress selected index changed handler.
                List1.DataSource = Nothing ' Remove and re-associate database in order to refresh List1
                List1.DataSource = searData
                List1.DisplayMember = "SearDataName"
                List1.ValueMember = "SearIdxName"
                List1.ClearSelected()
                dontHandleSelIdxChg = False
                Me.Visible = False

                'Me.Visible = True
                'Text1.Select()
                'List1.Visible = False
                Form1.FindLoaded = True
            End If

        Catch
            Cursor = Cursors.Default    ' Change cursor to default.
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub Find_Click()
        List1.Visible = False
    End Sub

    Private Sub Find_GotFocus()
        Text1.Select()
        FindFirst.Enabled = False
        FindAllButton.Enabled = False
    End Sub

    Private Sub List1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles List1.SelectedIndexChanged
        Dim lPtr1 As Integer     ' string pointer

        SupTxtChgHnd = True ' Don't allow Text1 text changed handler code to execute
        If dontHandleSelIdxChg = False Then
            If Me.List1.SelectedIndex >= 0 Then    ' If an item is selected.
                locLineDat = searData(List1.SelectedIndex).SearDataName.ToString
                lPtr1 = InStr(locLineDat, Chr(9))
                If lPtr1 > 0 Then locLineDat = Mid(locLineDat, 1, lPtr1 - 1)
                Text1.Text = locLineDat
                FindFirst.Enabled = True
                FindAllButton.Enabled = True
            End If
        End If
        SupTxtChgHnd = False
    End Sub

    Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Me.List1.Select()
    End Sub

    Private Sub FindFirst_Click(sender As Object, e As EventArgs) Handles FindFirst.Click ' Find first
        Dim dummy As Integer

        If List1.FindString(Text1.Text) = ListBox.NoMatches Then
            dummy = MsgBox("The search word you typed is not in the list; please select another word.", vbOKOnly + vbExclamation)
            Exit Sub
        End If

        Cursor = Cursors.WaitCursor    ' Change cursor to hourglass.

        Call Load_Indexes() ' Load index array from List1 selected item data

        ' Set idxPtr back to 0 because we are finding first.
        idxPtr = 0

        ' Scroll Form1.Lst1 to appropriate line
        ' Unfortunately the ListBox.FindString method will not scroll the list box if the string is visible in the list.
        ' This means the found item will not always appear at the top of the list.
        Form1.Lst1.TopIndex = idxArray(idxPtr) - 1  ' List item lines are 0-based
        If numIndexes > 1 Then
            Me.FindNext.Enabled = True
        End If

        FindNext.Select()
        Cursor = Cursors.Default    ' Change cursor to default symbol.

    End Sub

    Private Sub Load_Indexes()
        ' Reads index string from List1 and loads the index array
        Dim locIdxLineDat As String


        ' Get index string from List1
        idxLineDat = searData(List1.SelectedIndex).SearIdxName.ToString()

        numIndexes = 1  ' There will be at least one index in the string
        ' Find how many indexes are in the string
        locIdxLineDat = idxLineDat
        While InStr(locIdxLineDat, ",") <> 0
            'ddddWhile InStr(locIdxLineDat, ";") <> 0    'nnnn
            locIdxLineDat = locIdxLineDat.Substring(InStr(locIdxLineDat, ",") + 1)
            numIndexes += 1
        End While

        ' ReDim the idxArray to the appropriate size
        ReDim idxArray(numIndexes)

        ' Populate the idxArray from idxLineDat
        locIdxLineDat = idxLineDat
        For idxPtr = 0 To numIndexes - 1
            If InStr(locIdxLineDat, ",") > 0 Then

                idxArray(idxPtr) = Int(Strings.Left(locIdxLineDat, InStr(locIdxLineDat, ",") - 1))
                locIdxLineDat = locIdxLineDat.Substring(InStr(locIdxLineDat, ",") + 1)
            Else
                idxArray(idxPtr) = Int(locIdxLineDat)
            End If
        Next

    End Sub
    Private Sub FindNext_Click(sender As Object, e As EventArgs) Handles FindNext.Click
        If idxPtr < numIndexes Then ' Should always be true
            FindPrev.Enabled = True
            idxPtr += 1  ' Point to next index
            ' Scroll Form1.Lst1 to appropriate line
            ' Unfortunately the ListBox.FindString method will not scroll the list box if the string is visible in the list.
            ' This means the found item will not always appear at the top of the list.
            ' This needs to be detailed in the instructions / help files.
            Form1.Lst1.TopIndex = idxArray(idxPtr) - 1  ' List item lines are 0-based
            If idxPtr = numIndexes - 1 Then ' No more indexes to find
                Me.FindNext.Enabled = False
                FindPrev.Select()
            Else
                FindNext.Select()
            End If
        End If
    End Sub

    Private Sub FindPrev_Click(sender As Object, e As EventArgs) Handles FindPrev.Click
        If idxPtr > 0 Then  ' Should always be true
            idxPtr -= 1
            ' Scroll Form1.Lst1 to appropriate line
            ' Unfortunately the ListBox.FindString method will not scroll the list box if the string is visible in the list.
            ' This means the found item will not always appear at the top of the list.
            ' This needs to be detailed in the instructions / help files.
            Form1.Lst1.TopIndex = idxArray(idxPtr) - 1  ' List item lines are 0-based
            If idxPtr > 0 Then
                FindPrev.Enabled = False
            End If
        End If

        FindNext.Enabled = True
    End Sub

    Private Sub FindAllButton_Click(sender As Object, e As EventArgs) Handles FindAllButton.Click
        Dim dummy As Integer
        Dim ErrorFlag As Integer    ' error number
        Dim findallSymptomID As Integer ' symptom ID of Lst1 symptom
        Dim msg As String

        Try
            ErrorFlag = 3004
            Form1.RaiseException = False
            If List1.FindString(Text1.Text) = ListBox.NoMatches Then
                dummy = MsgBox("The search word you typed is not in the list; please select another word.", vbOKOnly + vbExclamation)
                Exit Sub
            End If

            Form1.FindAllCaller = 1
            If FindAll.Visible Then FindAll.Close()      ' Need to re-initialize in case it is open.
            FindAll.Left = 0
            FindAll.Top = 0
            FindAll.Label1.Visible = False
            FindAll.Show()
            Cursor = Cursors.WaitCursor    ' Change cursor to hourglass.

            Call Load_Indexes() ' Load index array from List1 selected item data

            FindAll.FindAllSympData.Clear()
            For idxPtr = 0 To numIndexes - 1
                ' Load FindAll Lst1 with Form1.Lst1 items pointed to by idxArray
                FindAll.FindAllSympData.Add(New FindAllSympDat(Form1.CombSympData(idxArray(idxPtr) - 1).SympDesc.ToString, Form1.CombSympData(idxArray(idxPtr) - 1).SympData.ToString))
            Next

            ' Associate FindAll.Lst1 with FindAllSympData
            FindAll.Lst1.DataSource = Nothing ' Remove and re-associate database in order to refresh Lst1
            FindAll.Lst1.DataSource = FindAll.FindAllSympData

            FindAll.Lst1.DisplayMember = "SympDesc"
            FindAll.Lst1.ValueMember = "SympData"
            FindAll.Lst1.ClearSelected()
            ' Now select any symptoms that are selected in Form1.Lst1
            For idxPtr = 0 To numIndexes - 1
                findallSymptomID = Form1.GetSymptomID(FindAll.FindAllSympData(idxPtr).SympData.ToString)
                If Form1.IsSymptomSelected(findallSymptomID) Then
                    FindAll.Lst1.SetSelected(idxPtr, True)
                End If
                If Form1.RaiseException Then Throw New System.Exception("An exception has occurred.")
            Next
            ' Clear any symptom that may have been shown in FindAll TextBox1 as a result of adding symptoms.
            FindAll.TextBox1.Text = ""
            Cursor = Cursors.Default    ' Change cursor to default symbol.
        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            Close()
        End Try
    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs) Handles CancelFindButton.Click
        Form1.SearchString = Text1.Text
        Form1.FindAllCaller = 0
        Form1.GridSearchButton.Enabled = True
        'Close()
        Me.Visible = False
    End Sub

    Private Sub HelpButton_Click(sender As Object, e As EventArgs) Handles HelpFindButton.Click
        Dim msg As String
        Dim myProcess As New Process

        Try
            myProcess.StartInfo.FileName = AppContext.BaseDirectory + Form1.HELP_DIR + "FindForm.htm"
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

    Private Sub Clear_Click(sender As Object, e As EventArgs) Handles Clear.Click
        SupTxtChgHnd = True ' Don't allow Text1 text changed handler code to execute
        Me.Text1.Text = ""
        Me.Text1.Focus()
        SupTxtChgHnd = False
    End Sub

    Private Sub Text1_TextChanged(sender As Object, e As EventArgs) Handles Text1.TextChanged
        Dim ErrorFlag As Integer    ' error number
        Dim Found As Boolean    ' text was found in list
        Dim listIndex As Integer    ' list index of found string
        Dim msg As String       ' error message string

        Static old_file_pos As Integer

        Try
            ErrorFlag = 3005
            If Not SupTxtChgHnd Then
                Me.List1.Visible = True
                Me.CancelFindButton.Text = "Done"

                ' Set up binary search parameters.
                If Len(Text1.Text) <= 1 Then
                    First = 0
                Else
                    First = old_file_pos - 2
                    If First < 0 Then First = 0
                End If

                Me.FindNext.Enabled = False
                Me.FindPrev.Enabled = False
                Last = Form1.Sfile_size

                ' Find text position in file.
                Found = False

                listIndex = List1.FindString(Text1.Text)
                List1.TopIndex = listIndex

            End If

        Catch
            Cursor = Cursors.Default    ' Change cursor to default.
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            Close()
        End Try
    End Sub

End Class

Public Class searDat
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