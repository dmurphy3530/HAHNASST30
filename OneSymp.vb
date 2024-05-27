Option Explicit On
Imports System.Windows.Forms
Imports System.Collections
Imports System.Drawing
Public Class OneSymp
    Public First As Integer    ' top of binary search
    Public Last As Integer     ' bottom of binary search
    Dim LineCount As Integer    ' Number of lines in input file
    Dim locLineDat As String     ' selected list item
    Public loc_file_pos As Long    ' local symptom file position
    Public SupressText1Change As Boolean    ' bypasses Text1 change sub
    Public SympIDData As String   ' Data string from selected symptom
    Public oneSympData As New ArrayList()
    Public twoSympData As New ArrayList()
    Public treSympData As New ArrayList()
    Public fourSympData As New ArrayList()
    Public Shared TwoSympFileData() = System.IO.File.ReadAllLines(AppContext.BaseDirectory + Form1.DATA_DIR + "twosymptextdata.dat") ' holds the data from twosymptextdata.dat
    Public Shared TreSympFileData() = System.IO.File.ReadAllLines(AppContext.BaseDirectory + Form1.DATA_DIR + "tresymptextdata.dat") ' holds the data from tresymptextdata.dat
    Public Shared FourSympFileData() = System.IO.File.ReadAllLines(AppContext.BaseDirectory + Form1.DATA_DIR + "foursymptextdata.dat") ' holds the data from foursymptextdata.dat
    Public SupTxtChgHnd As Boolean  ' used to suppress action of text change handler if software updated text box

    Private Sub OneSymp_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim ErrorFlag As Integer    ' error number
        Dim InName As String    ' name of input file
        Dim msg As String       ' Error message string

        Try
            ErrorFlag = 6001

            SelectRemedy.Enabled = False
            MustUse.Enabled = False

            Text1.BackColor = Form1.TextBox1.BackColor
            Text1.ForeColor = Form1.TextBox1.ForeColor
            Text1.Font = Form1.TextBox1.Font

            Text2.BackColor = Form1.TextBox1.BackColor
            Text2.ForeColor = Form1.TextBox1.ForeColor
            Text2.Font = Form1.TextBox1.Font

            Lst1.BackColor = Color.FromName("Cornsilk")
            Lst1.ForeColor = Form1.Lst1.ForeColor
            Lst1.Font = Form1.Lst1.Font

            Lst2.BackColor = Form1.Lst1.BackColor
            Lst2.ForeColor = Form1.Lst1.ForeColor
            Lst2.Font = Form1.Lst1.Font

            Lst3.BackColor = Form1.Lst1.BackColor
            Lst3.ForeColor = Form1.Lst1.ForeColor
            Lst3.Font = Form1.Lst1.Font

            Lst4.BackColor = Form1.Lst1.BackColor
            Lst4.ForeColor = Form1.Lst1.ForeColor
            Lst4.Font = Form1.Lst1.Font

            Lst2.Visible = False
            Lst3.Visible = False
            Lst4.Visible = False

            Text1.Text = ""
            Text2.Text = ""
            SupressText1Change = False
            InName = AppContext.BaseDirectory + "\onesymptextdata.dat"
            Form1.ErrorFlag = 6002

            Lst1.DisplayMember = "OneSympDataName"
            Lst1.ValueMember = "OneSympIdxName"
            Lst1.DataSource = oneSympData
            Lst1.ClearSelected()
            Me.Visible = False

            Command1.Enabled = True

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            Cursor = Cursors.Default    ' Change cursor to default
            Close()
        End Try
    End Sub

    Private Sub Form_Unload(sender As Object, e As EventArgs) Handles MyBase.Closing
        Dim ErrorFlag As Integer    ' error flag
        Dim msg As String   ' Error message string

        Try
            ErrorFlag = 6003
            Form1.SelectRemedy.Enabled = True
            Form1.MustUse.Enabled = True

            Cursor = Cursors.Default    ' Change cursor to default.

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            Cursor = Cursors.Default    ' Change cursor to default.
        End Try
    End Sub

    Private Sub Text1_TextChanged(sender As Object, e As EventArgs) Handles Text1.TextChanged
        Dim ErrorFlag As Integer    ' error number
        Dim listIndex As Integer    ' list index of found string
        Dim msg As String       ' error message string

        Try
            ErrorFlag = 7020
            If Not SupTxtChgHnd Then

                If SupressText1Change Then
                    SupressText1Change = False
                    Exit Sub
                End If
                If Text1.Text = "" Then Exit Sub

                If Lst4.Visible Then
                    listIndex = Lst4.FindString(Text1.Text)
                    Lst4.TopIndex = listIndex
                ElseIf Lst3.Visible Then
                    listIndex = Lst3.FindString(Text1.Text)
                    Lst3.TopIndex = listIndex
                ElseIf Lst2.Visible Then
                    listIndex = Lst2.FindString(Text1.Text)
                    Lst2.TopIndex = listIndex
                Else
                    listIndex = Lst1.FindString(Text1.Text)
                    Lst1.TopIndex = listIndex
                End If
            End If

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            Close()
        End Try
    End Sub

    Private Sub Text2_TextChanged(sender As Object, e As EventArgs) Handles Text2.TextChanged

    End Sub

    Private Sub Lst1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Lst1.SelectedIndexChanged
    End Sub

    Private Sub Lst1_Click(sender As Object, e As EventArgs) Handles Lst1.Click
        If Lst1.SelectedIndex <> -1 Then ' If click did not result in de-selection
            Handle_Lst1_Selection()
        End If
    End Sub

    Private Sub Handle_Lst1_Selection()
        Dim msg As String       ' error message string
        Dim SympText As String      ' text from selected symptom
        Dim SympData As String      ' data from selected symptom
        Dim StrPtr As String    ' String form of RecordPointer element
        Dim recPointer As Integer   ' pointer to line in file for item to display in next list box

        Try
            Form1.ErrorFlag = 6004
            SupTxtChgHnd = True ' Don't allow Text1 text changed handler code to execute

            ' Hide the other lists in case form was already loaded and a different symptom-portion was selected.
            Lst2.Visible = False
            Lst3.Visible = False
            Lst4.Visible = False

            If Me.Lst1.SelectedIndex >= 0 Then    ' If an item is selected.
                Lst2.Visible = False    ' In case a new item is selected after loading
                Lst3.Visible = False
                Lst4.Visible = False
                SelectRemedy.Enabled = False
                MustUse.Enabled = False

                SympText = oneSympData(Lst1.SelectedIndex).OneSympDataName.ToString
                SympData = oneSympData(Lst1.SelectedIndex).OneSympIdxName.ToString
                Text1.Text = ""
                Text2.Text = SympText
                If Strings.Left(SympData, 1) = "," Then ' If pointer is Symptom ID, there's no more to display.
                    SympIDData = SympData
                    SelectRemedy.Enabled = True
                    MustUse.Enabled = True
                Else
                    twoSympData.Clear()   ' Clear any previous data
                    Lst2.DataSource = Nothing
                    While InStr(SympData, ";") > 0
                        SympData = SympData.Substring(1)    ' Remove leading ";"
                        If InStr(SympData, ";") > 0 Then
                            StrPtr = Strings.Left(SympData, InStr(SympData, ";") - 1)
                        Else
                            StrPtr = SympData
                        End If
                        recPointer = Val(StrPtr)
                        twoSympData.Add(New twoSympDat(TwoSympFileData(2 * recPointer), TwoSympFileData((2 * recPointer) + 1)))
                        SympData = SympData.Substring(Len(StrPtr))
                        '                    Next
                    End While

                    Lst2.DataSource = twoSympData
                    Lst2.DisplayMember = "TwoSympDataName"
                    Lst2.ValueMember = "TwoSympIdxName"
                    Lst1.BackColor = Form1.Lst1.BackColor
                    Lst2.BackColor = Color.FromName("Cornsilk")
                    Lst2.ClearSelected()
                    Lst2.Visible = True
                    Text1.Select()  ' Set focus to Text1 for next entry.
                    SupTxtChgHnd = False    ' Allow Text1 text changed handler to handle text for Lst2 selection
                    If Lst2.Items.Count = 1 Then    ' No need to wait for selection, only one item
                        Lst2.SetSelected(0, True)   ' Select the only item.
                        Handle_Lst2_Selection()
                    End If
                End If
                Exit Sub
            End If
            SupTxtChgHnd = False

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            Close()
        End Try
    End Sub
    Private Sub List1_GotFocus()
        Dim ErrorFlag As Integer    ' error number
        Dim lPtr1 As Integer     ' string pointer
        Dim msg As String   ' error message string

        Try
            ErrorFlag = 6005
            If Lst1.SelectedIndex >= 0 Then    ' If an item is selected.
                locLineDat = Lst1.SelectedItem
                lPtr1 = InStr(locLineDat, Chr(9))
                If lPtr1 > 0 Then locLineDat = Mid(locLineDat, 1, lPtr1 - 1)
                SupressText1Change = True
                Text1.Text = locLineDat
            End If

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub


    Private Sub VScroll1_Scroll(sender As Object, e As ScrollEventArgs)

    End Sub

    Private Sub SelectRemedy_Click(sender As Object, e As EventArgs) Handles SelectRemedy.Click
        Dim msg As String               ' for diagnostics
        Dim SymDatTxt As String         ' Text of symptom data from list box
        Try
            ' The list box that contains the final portion of the symptom phrase has the symptom data string as its data.  When the last portion is
            ' selected in a list box, the symptom data string is copied to variable SympIDData and the "Select" and "Must Use" buttons are enabled.

            Form1.ErrorFlag = 6006
            Form1.RaiseException = False

            SymDatTxt = SympIDData.Substring(1)   ' Remove leading comma
            Form1.HandleSelectRemedy(SymDatTxt)
            If Form1.RaiseException Then Throw New System.Exception("An exception has occurred.")
            Form1.EnableFilePrint()
            Form1.Dirty = True

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            Close()
        End Try
    End Sub

    Private Sub MustUse_Click(sender As Object, e As EventArgs) Handles MustUse.Click
        Dim msg As String               ' error message string
        Dim SymDatTxt As String         ' Text of symptom data from list box

        Try
            ' The list box that contains the final portion of the symptom phrase has the symptom data string as its data.  When the last portion is
            ' selected in a list box, the symptom data string is copied to variable SympIDData and the "Select" and "Must Use" buttons are enabled.
            Form1.ErrorFlag = 6009
            Form1.RaiseException = False
            SymDatTxt = SympIDData.Substring(1)   ' Remove leading comma
            Form1.Handle_MustUse(SymDatTxt, Text2.Text)

            Exit Sub

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            Close()
        End Try
    End Sub

    Private Sub Command1_Click(sender As Object, e As EventArgs) Handles Command1.Click
        Dim ErrorFlag As Integer    ' error number
        Dim msg As String   ' error message string

        Try
            ErrorFlag = 6015
            Me.Visible = False

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub Weight_SelectedIndexChanged(sender As Object, e As EventArgs)
        Form1.Weight.Text = Weight.Text ' Synchronize Form1 Weight; it will be used to set selected symptom(s) weight
    End Sub

    Private Sub Lst2_Click(sender As Object, e As EventArgs) Handles Lst2.Click
        Dim msg As String   ' error message string

        Try
            Form1.ErrorFlag = 6017
            If Lst2.SelectedIndex <> -1 Then ' If click did not result in de-selection
                Handle_Lst2_Selection()
            End If

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub Handle_Lst2_Selection()
        Dim msg As String       ' error message string
        Dim SympText As String      ' text from selected symptom
        Dim SympData As String      ' data from selected symptom
        Dim StrPtr As String    ' String form of RecordPointer element
        Dim recPointer As Integer   ' pointer to line in file for item to display in next list box

        Try
            Form1.ErrorFlag = 6018
            SupTxtChgHnd = True ' Don't allow Text1 text changed handler code to execute

            ' Hide the other lists in case form was already loaded and a different symptom-portion was selected.
            Lst3.Visible = False
            Lst4.Visible = False

            If Me.Lst2.SelectedIndex >= 0 Then    ' If an item is selected.
                Lst3.Visible = False    ' In case a new item is selected after loading
                Lst4.Visible = False
                SelectRemedy.Enabled = False
                MustUse.Enabled = False

                SympText = twoSympData(Lst2.SelectedIndex).TwoSympDataName.ToString
                SympData = twoSympData(Lst2.SelectedIndex).TwoSympIdxName.ToString
                Text1.Text = ""
                Text2.Text = oneSympData(Lst1.SelectedIndex).OneSympDataName.ToString + SympText
                If Strings.Left(SympData, 1) = "," Then ' If pointer is Symptom ID, there's no more to display.
                    SympIDData = SympData
                    SelectRemedy.Enabled = True
                    MustUse.Enabled = True
                Else
                    treSympData.Clear()   ' Clear any previous data
                    Lst3.DataSource = Nothing
                    While InStr(SympData, ";") > 0
                        SympData = SympData.Substring(1)    ' Remove leading ";"
                        If InStr(SympData, ";") > 0 Then
                            StrPtr = Strings.Left(SympData, InStr(SympData, ";") - 1)
                        Else
                            StrPtr = SympData
                        End If
                        recPointer = Val(StrPtr)
                        treSympData.Add(New treSympDat(TreSympFileData(2 * recPointer), TreSympFileData((2 * recPointer) + 1)))
                        SympData = SympData.Substring(Len(StrPtr))
                    End While
                    Lst3.DataSource = treSympData
                    Lst3.DisplayMember = "TreSympDataName"
                    Lst3.ValueMember = "TreSympIdxName"
                    Lst2.BackColor = Form1.Lst1.BackColor
                    Lst3.BackColor = Color.FromName("Cornsilk")
                    Lst3.ClearSelected()
                    Lst3.Visible = True
                    Text1.Select()  ' Set focus to Text1 for next entry.
                    SupTxtChgHnd = False    ' Allow Text1 text changed handler to handle text for Lst3 selection
                    If Lst3.Items.Count = 1 Then    ' No need to wait for selection, only one item
                        Lst3.SetSelected(0, True)   ' Select the only item.
                        Handle_Lst3_Selection()
                    End If
                End If
                Exit Sub
            End If
            SupTxtChgHnd = False

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
            Buttons:=vbOKOnly + vbCritical,
            Title:="ERROR!")
            Close()
        End Try
    End Sub

    Private Sub Lst3_Click(sender As Object, e As EventArgs) Handles Lst3.Click
        Dim msg As String   ' error message string

        Try
            Form1.ErrorFlag = 6019
            If Lst3.SelectedIndex <> -1 Then ' If click did not result in de-selection
                Handle_Lst3_Selection()
            End If

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try

    End Sub

    Private Sub Handle_Lst3_Selection()
        Dim msg As String       'error message String
        Dim SympText As String      ' text from selected symptom
        Dim SympData As String      ' data from selected symptom
        Dim StrPtr As String    ' String form of RecordPointer element
        Dim recPointer As Integer   ' pointer to line in file for item to display in next list box

        Try
            Form1.ErrorFlag = 6020
            SupTxtChgHnd = True ' Don't allow Text1 text changed handler code to execute

            ' Hide the other lists in case form was already loaded and a different symptom-portion was selected.
            Lst4.Visible = False

            If Me.Lst3.SelectedIndex >= 0 Then    ' If an item is selected.
                Lst4.Visible = False    ' In case a new item is selected after loading
                SelectRemedy.Enabled = False
                MustUse.Enabled = False

                SympText = treSympData(Lst3.SelectedIndex).TreSympDataName.ToString
                SympData = treSympData(Lst3.SelectedIndex).TreSympIdxName.ToString
                Text1.Text = ""
                'Text2.Text = Text2.Text + " " + SympText
                Text2.Text = oneSympData(Lst1.SelectedIndex).OneSympDataName.ToString +
                    twoSympData(Lst2.SelectedIndex).TwoSympDataName.ToString + SympText
                If Strings.Left(SympData, 1) = "," Then ' If pointer is Symptom ID, there's no more to display.
                    SympIDData = SympData
                    SelectRemedy.Enabled = True
                    MustUse.Enabled = True
                Else
                    fourSympData.Clear()   ' Clear any previous data
                    Lst4.DataSource = Nothing
                    While InStr(SympData, ";") > 0
                        SympData = SympData.Substring(1)    ' Remove leading ";"
                        If InStr(SympData, ";") > 0 Then
                            StrPtr = Strings.Left(SympData, InStr(SympData, ";") - 1)
                        Else
                            StrPtr = SympData
                        End If
                        recPointer = Val(StrPtr)
                        fourSympData.Add(New fourSympDat(FourSympFileData(2 * recPointer), FourSympFileData((2 * recPointer) + 1)))
                        SympData = SympData.Substring(Len(StrPtr))
                    End While
                    Lst4.DataSource = fourSympData
                    Lst4.DisplayMember = "FourSympDataName"
                    Lst4.ValueMember = "FourSympIdxName"
                    Lst3.BackColor = Form1.Lst1.BackColor
                    Lst4.BackColor = Color.FromName("Cornsilk")
                    Lst4.ClearSelected()
                    Lst4.Visible = True
                    Text1.Select()  ' Set focus to Text1 for next entry.
                    SupTxtChgHnd = False    ' Allow Text1 text changed handler to handle text for Lst4 selection
                    If Lst4.Items.Count = 1 Then    ' No need to wait for selection, only one item
                        Lst4.SetSelected(0, True)   ' Select the only item.
                        Handle_Lst4_Selection()
                    End If
                End If
                Exit Sub
            End If
            SupTxtChgHnd = False

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
            Buttons:=vbOKOnly + vbCritical,
            Title:="ERROR!")
            Close()
        End Try
    End Sub

    Private Sub Lst4_Click(sender As Object, e As EventArgs) Handles Lst4.Click
        Dim msg As String   ' error message string

        Try
            Form1.ErrorFlag = 6021
            If Lst4.SelectedIndex <> -1 Then ' If click did not result in de-selection
                Handle_Lst4_Selection()
            End If

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub Handle_Lst4_Selection()
        Dim ErrorFlag As Integer    ' error number
        Dim msg As String       ' error message string
        Dim SympText As String      ' text from selected symptom
        Dim SympData As String      ' data from selected symptom

        Try
            ErrorFlag = 6022
            SupTxtChgHnd = True ' Don't allow Text1 text changed handler code to execute

            If Me.Lst3.SelectedIndex >= 0 Then    ' If an item is selected.
                SelectRemedy.Enabled = False
                MustUse.Enabled = False

                SympText = fourSympData(Lst4.SelectedIndex).FourSympDataName.ToString
                SympData = fourSympData(Lst4.SelectedIndex).FourSympIdxName.ToString
                Text1.Text = ""
                Text2.Text = Text2.Text + " " + SympText
                Text2.Text = oneSympData(Lst1.SelectedIndex).OneSympDataName.ToString +
                    twoSympData(Lst2.SelectedIndex).TwoSympDataName.ToString +
                    treSympData(Lst3.SelectedIndex).TreSympDataName.ToString + SympText
                If Strings.Left(SympData, 1) = "," Then ' If pointer is Symptom ID, there's no more to display.
                    SympIDData = SympData
                    SelectRemedy.Enabled = True
                    MustUse.Enabled = True
                End If
            End If  ' Note:  Lst4 records do not contain any symptom list pointers, only symptom-ID's, denoted by leading ",".
            Exit Sub
            SupTxtChgHnd = False

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
            Buttons:=vbOKOnly + vbCritical,
            Title:="ERROR!")
            Close()
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Lst1.ClearSelected()
        Text1.Text = ""
        Text2.Text = ""
        Lst1.BackColor = Color.FromName("Cornsilk")
        If Lst2.Visible Then
            Lst2.ClearSelected()
            Lst2.Visible = False
        End If
        If Lst3.Visible Then
            Lst3.ClearSelected()
            Lst3.Visible = False
        End If
        If Lst4.Visible Then
            Lst4.ClearSelected()
            Lst4.Visible = False
        End If
        SupTxtChgHnd = False    ' Ensure text entry is (re)enabled.
    End Sub

    Private Sub Weight_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles Weight.SelectedIndexChanged
        Form1.Weight.Text = Weight.Text ' Synchronize Form1 Weight; it will be used to set selected symptom(s) weight
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim msg As String
        Dim myProcess As New Process

        Try
            myProcess.StartInfo.FileName = AppContext.BaseDirectory + Form1.HELP_DIR + "RepertoryStyleSymptomForms.htm"
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

Public Class oneSympDat
    Private myOneSympIdxName As String
    Private myOneSympDataName As String

    Public Sub New(ByVal strOneSympDataName As String, ByVal strOneSympIdxName As String)
        Me.myOneSympIdxName = strOneSympIdxName
        Me.myOneSympDataName = strOneSympDataName
    End Sub

    Public ReadOnly Property OneSympIdxName() As String
        Get
            Return myOneSympIdxName
        End Get
    End Property

    Public ReadOnly Property OneSympDataName() As String
        Get
            Return myOneSympDataName
        End Get
    End Property

End Class

Public Class twoSympDat
    Private myTwoSympIdxName As String
    Private myTwoSympDataName As String

    Public Sub New(ByVal strTwoSympDataName As String, ByVal strTwoSympIdxName As String)
        Me.myTwoSympIdxName = strTwoSympIdxName
        Me.myTwoSympDataName = strTwoSympDataName
    End Sub

    Public ReadOnly Property TwoSympIdxName() As String
        Get
            Return myTwoSympIdxName
        End Get
    End Property

    Public ReadOnly Property TwoSympDataName() As String
        Get
            Return myTwoSympDataName
        End Get
    End Property

End Class



Public Class treSympDat

    Private myTreSympIdxName As String
    Private myTreSympDataName As String

    Public Sub New(ByVal strTreSympDataName As String, ByVal strTreSympIdxName As String)
        Me.myTreSympIdxName = strTreSympIdxName
        Me.myTreSympDataName = strTreSympDataName
    End Sub

    Public ReadOnly Property TreSympIdxName() As String
        Get
            Return myTreSympIdxName
        End Get
    End Property

    Public ReadOnly Property TreSympDataName() As String
        Get
            Return myTreSympDataName
        End Get
    End Property

End Class


Public Class fourSympDat
    Private myFourSympIdxName As String
    Private myFourSympDataName As String

    Public Sub New(ByVal strFourSympDataName As String, ByVal strFourSympIdxName As String)
        Me.myFourSympIdxName = strFourSympIdxName
        Me.myFourSympDataName = strFourSympDataName
    End Sub

    Public ReadOnly Property FourSympIdxName() As String
        Get
            Return myFourSympIdxName
        End Get
    End Property

    Public ReadOnly Property FourSympDataName() As String
        Get
            Return myFourSympDataName
        End Get
    End Property

End Class