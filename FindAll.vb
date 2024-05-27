Option Explicit On
Public Class FindAll
    Public Shared FindAllSympData As New ArrayList()  ' Holds ListBox symptom strings and ID data
    Public Shared FindAllSelectedIndicesArray(1) As Integer
    Public Shared FindAllSelectedSize As Integer
    Public Shared FindAllSymSelected(1) As Integer ' list of selected symptom ID's in Lst1
    Public Shared FindAllSymSelectedSize As Integer  ' # of selected symptom ID's in Lst1
    Private Sub FindAll_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' This form lets you select symptoms one at a time from the list of all wymptoms which contain the search word.
        ' Symptom will also appear selected in the Form1 Lst1 box.
        ' Until symptom(s) are selected (for "Selected" or "Must Use") only the latest symptom appears selected in the list box;
        ' once they are selected, they stay selected in the list.  But still only the latest of the ones clicked on since
        ' then appear selected.

        ' Behavior:
        ' 1 When list loads, if item is selected in Form1.Lst1, select it in FindAll.
        ' 2 Allow multiple selections, behavior identical to Form1.
        ' 3 If more than one new selection and "Must" selected, add the symptom in TextBox1.
        ' 4 Use the button handlers from Form1 for Select and Must Use.
        ' 5 (This will result in that all new selections will be selected in Form1 Lst1.)

        '  Load default settings.
        TextBox1.BackColor = Form1.TextBox1.BackColor
        TextBox1.ForeColor = Form1.TextBox1.ForeColor
        TextBox1.Font = Form1.TextBox1.Font

        Lst1.BackColor = Form1.Lst1.BackColor
        Lst1.ForeColor = Form1.Lst1.ForeColor
        Lst1.Font = Form1.Lst1.Font

        BackColor = Form1.BackColor
        ForeColor = Form1.ForeColor

        Form1.FindAllVisible = True
        If Form1.FindAllCaller = 1 Then
            Find.Visible = False
        ElseIf Form1.FindAllCaller = 2 Then
            GridSearch.Visible = False
        ElseIf Form1.FindAllCaller = 3 Then
            Questionnaire.Visible = False
        End If
        ' No need to handle case 4 (Show Similar Symptoms) since Form1 is always visible.

        FindAllSelectedIndicesArray(0) = -1
        FindAllSelectedSize = 0
    End Sub

    Private Sub Form_Unload(Cancel As Integer)

        If Form1.FindAllCaller = 1 Then
            Find.Visible = True
        ElseIf Form1.FindAllCaller = 2 Then
            GridSearch.Visible = True
        ElseIf Form1.FindAllCaller = 3 Then
            Questionnaire.Visible = True
        End If
        Me.Lst1.SelectedItem(0) = True

    End Sub

    Private Sub Lst1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Lst1.SelectedIndexChanged
        ' Since _SelectedIndexChanged gets triggered after _Click, we can put this here to save off the latest selection.
        ' This will also cover the case where index is changed programatically and _Click isn't triggered.
        FindAllSymSelectedSize = Lst1.SelectedIndices.Count
        ReDim FindAllSymSelected(FindAllSymSelectedSize)
        For ptr = 0 To Lst1.SelectedIndices.Count - 1
            FindAllSymSelected(ptr) = Lst1.SelectedIndices(ptr)
        Next
    End Sub

    Private Sub Lst1_Click(sender As Object, e As EventArgs) Handles Lst1.Click
        Dim FindAllListItem As Integer  ' points to Lst1 item
        Dim Form1SympID As Integer  ' Symptom ID in Form1.Lst1
        Dim found As Boolean    ' selected index is found
        Dim locSympIDString As String   ' SympData of Lst1 item
        Dim locSympID As Integer ' symptom ID of Lst1 item
        Dim locSympIndex As Integer ' symptom index of Lst1 item
        Dim msg As String   ' error message string
        Dim symptomAdded As Boolean ' Indicates click resulted in a new selected item if TRUE, or de-selected it if FALSE

        Try
            ' Find item not in selected list, add it to TextBox.
            ' NOTE:  Can't use .SelectedIndex because this always points to first selected in a group.
            Form1.ErrorFlag = 4001
            Form1.RaiseException = False

            symptomAdded = False
            For listItem = 0 To Lst1.SelectedIndices.Count - 1
                If Not IsFindAllSymptomSelected(Lst1.SelectedIndices(listItem)) Then  ' Display item in TextBox1
                    TextBox1.Text = FindAllSympData(Lst1.SelectedIndices(listItem)).SympDesc.ToString()
                    ' Need to select symptom in Form1.Lst1
                    locSympIDString = FindAllSympData(Lst1.SelectedIndices(listItem)).SympData.ToString()
                    locSympID = Form1.GetSymptomID(locSympIDString)
                    For FindAllListItem = 0 To Lst1.SelectedIndices.Count - 1
                        locSympIDString = FindAllSympData(FindAllListItem).SympData.ToString()
                        Form1SympID = Form1.GetSymptomID(locSympIDString)
                        If Form1SympID = locSympID Then
                            ' Need to select it in Form1.Lst1 and add it to FindAllSelectedSymp array
                            Form1.Lst1.SetSelected(Form1.MapSymIDToIndex(locSympID), True)
                        End If
                    Next

                    symptomAdded = True
                    Exit For
                End If
            Next

            If Not symptomAdded Then    ' Click resulted in de-selecting an item, if in (Form1) selected list need to re-select it in Lst1.
                ' Find symptom that is in FindAllSymSelected but not in Lst1.Selected Indices
                For SIItem = 0 To FindAllSymSelectedSize - 1
                    locSympIndex = FindAllSymSelected(SIItem)
                    locSympID = Form1.MapSymIDToIndex(locSympIndex)
                    found = False
                    For listItem = 0 To Lst1.SelectedIndices.Count - 1
                        If Lst1.SelectedIndices(listItem) = locSympIndex Then
                            found = True
                        End If
                        If found = False Then Exit For
                    Next
                    If found = False Then Exit For
                Next
                If found = False Then    ' Should always be the case here
                    ' Need to check if symptom is in one of the selected lists.
                    If (Form1.IsSymptomInSelectedList(locSympID) Or Form1.IsSymptomInMustUseList(locSympID)) Then
                        ' Can't un-select it in Lst1.
                        Lst1.Items(locSympIndex).SetSelected(True)  ' Re-select the symptom.
                    Else
                        ' If un-selected symptom was displayed in TextBox1, delete the text
                        If TextBox1.Text = FindAllSympData(locSympIndex).SympDesc.ToString() Then
                            TextBox1.Text = ""
                        End If
                        ' Need to unselect symptom in Form1.Lst1.
                        For listItem = 0 To Form1.Lst1.SelectedIndices.Count - 1
                            locSympIDString = Form1.CombSympData(Form1.Lst1.SelectedIndices(listItem)).SympData.ToString()
                            Form1SympID = Form1.GetSymptomID(locSympIDString)
                            If Form1SympID = locSympID Then
                                ' Need to un-select it in Form1.Lst1
                                Form1.Lst1.Items(Form1.Lst1.SelectedIndices(listItem)).SetSelected(False)
                            End If
                        Next
                    End If
                    If Form1.RaiseException Then Throw New System.Exception("An exception has occurred.")
                End If
            End If

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            Close()
        End Try
    End Sub

    Public Function IsFindAllSymptomSelected(sympIndex As Integer) As Boolean
        ' Returns True if symptom has been selected in the Lst1 (symptom selection) box, otherwise False.
        Dim ptr As Integer

        IsFindAllSymptomSelected = False
        For ptr = 0 To FindAllSymSelectedSize - 1
            If sympIndex = FindAllSymSelected(ptr) Then
                IsFindAllSymptomSelected = True
                Exit For
            End If
        Next
    End Function

    Public Function InFindAllSelectedIndicesArray(SelIndex As Integer) As Boolean
        Dim FindAllListItem As Integer

        ' If SelIndex is contained in the SelectedIndicesArray it returns True, otherwise False
        InFindAllSelectedIndicesArray = False
        For FindAllListItem = 0 To FindAllSymSelectedSize
            If FindAllSelectedIndicesArray(FindAllListItem) = SelIndex Then
                InFindAllSelectedIndicesArray = True
            End If
        Next
    End Function


    Private Sub SelectRemedy_Click(sender As Object, e As EventArgs) Handles SelectRemedy.Click
        Dim msg As String               ' for diagnostics
        Dim NumItems As Integer         ' number of items selected
        Dim SympData As New ArrayList() ' Holds ListBox symptom strings and ID data
        Dim SymDatTxt As String         ' Text of symptom data from list box

        Try
            Form1.ErrorFlag = 4002
            Form1.RaiseException = False
            ' Point HandleSelectRemedy to Lst1
            SympData = FindAllSympData
            SymDatTxt = ""
            NumItems = Me.Lst1.SelectedItems.Count

            For Ptr = 0 To NumItems - 1
                SymDatTxt = SympData(Lst1.SelectedIndices(Ptr)).SympData.ToString

                Form1.HandleSelectRemedy(SymDatTxt)
            Next

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            Close()
        End Try
    End Sub

    Private Sub MustUse_Click(sender As Object, e As EventArgs) Handles MustUse.Click
        Dim IDString As String          ' string of ID's
        Dim msg As String               ' for diagnostics
        Dim Ptr1 As Integer
        Dim SympData As ArrayList  ' Holds ListBox symptom strings and ID data
        Dim SymptomData As String   ' temporary string
        Dim SymptomText As String       ' Text of the symptom

        Try
            Form1.ErrorFlag = 4005
            Form1.RaiseException = False
            SympData = FindAllSympData
            SymptomText = SympData(Lst1.SelectedIndices(0)).SympText.ToString   ' Use first selected item in the list
            IDString = SympData(Lst1.SelectedIndices(0)).SympData.ToString
            Form1.Handle_MustUse(IDString, SymptomText)
            ' Set any newly-selected symptoms as "selected" in Form1 Lst1.
            For Ptr1 = 0 To Lst1.Items.Count - 1
                SymptomData = Form1.CombSympData(Ptr1).SympData.ToString()
                If (Form1.SymptomIsSelected(SymptomData)) Then Form1.Lst1.SetSelected(Ptr1, True)
            Next Ptr1
        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            Close()
        End Try
        Exit Sub
    End Sub

    Private Sub Command3_Click(sender As Object, e As EventArgs) Handles Command3.Click 'Done button

        Me.Label1.Visible = False   ' Default:  don't show "page - question" heading.
        Me.Visible = False
        Form1.FindAllVisible = False
    End Sub

    Private Sub Weight_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Weight.SelectedIndexChanged
        Form1.Weight.Text = Weight.Text ' Synchronize Form1 Weight; it will be used to set selected symptom(s) weight
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click   'Help buton
        Dim msg As String
        Dim myProcess As New Process

        Try
            myProcess.StartInfo.FileName = AppContext.BaseDirectory + Form1.HELP_DIR + "FindAllForm.htm"
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

Public Class FindAllSympDat
    Private mySympData As String
    Private mySympDesc As String

    Public Sub New(ByVal strSympDesc As String, ByVal strSympData As String)
        Me.mySympData = strSympData
        Me.mySympDesc = strSympDesc
    End Sub

    Public ReadOnly Property SympData() As String
        Get
            Return mySympData
        End Get
    End Property

    Public ReadOnly Property SympDesc() As String
        Get
            Return mySympDesc
        End Get
    End Property

End Class