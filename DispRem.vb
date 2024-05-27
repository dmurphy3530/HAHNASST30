Option Explicit On
Public Class DispRem
    Public Shared DispRemData As New ArrayList()  ' Holds ListBox symptom strings and ID data

    Dim MatFileName As String   ' name of remedies file

    Private Sub DispRem_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        Dim Ptr1, Ptr2 As Integer   ' sort pointers
        Dim RemCtr As Integer   ' remedy counter for finding remedy in matmed file
        Dim StringRemID As String   ' remedy ID from list box string
        Dim StringSeqSympNo As String  ' String form of sequential symptom number from list box string
        Dim StrPtr As Integer   ' string pointer
        Dim SympInList As Boolean   ' symptom is in must box or selected symptom list if true
        Dim SymptomNumber As Integer    ' sequential number of symptom in MatMed remedy text
        Dim T1 As Integer               ' sort item temporary variable

        ErrorFlag = 2001
        ReDim LocSeqSympNo(0)
        ReDim LocSeqSympNo2(0)

        TextBox1.BackColor = Form1.Lst1.BackColor
        TextBox1.ForeColor = Form1.Lst1.ForeColor
        TextBox1.Font = Form1.Lst1.Font

        Try
            ErrorFlag = 2002
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

            ' Sort LocSeqSympNo and eliminate duplicates.
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

            ErrorFlag = 2004
            SymptomNumber = 0

            ErrorFlag = 2006
            LinePtr = 1 ' Point past first blank line
            For RemCtr = 1 To Form1.RemedyNum
                While SelRem.MatMedText(LinePtr) <> ""
                    LinePtr += 1
                End While
                LinePtr += 1    ' Point past the blank line
            Next
            ' Should now be pointing to selected remedy; add remedy text to List1
            Dim outString As String
            outString = ""
            While SelRem.MatMedText(LinePtr) <> ""
                If (InStr(Mid(SelRem.MatMedText(LinePtr), 1, 1), Chr(9)) > 0) Then    ' if a symptom
                    SympInList = False
                    SymptomNumber = SymptomNumber + 1
                    Ptr1 = 0
                    While Not SympInList And Ptr1 < LocSympCount
                        If LocSeqSympNo2(Ptr1) = SymptomNumber Then SympInList = True
                        Ptr1 = Ptr1 + 1
                    End While
                    ErrorFlag = 2009

                    If (SympInList) Then
                        OutLine = ">>>" + SelRem.MatMedText(LinePtr)
                    Else
                        OutLine = SelRem.MatMedText(LinePtr)
                    End If
                Else
                    OutLine = SelRem.MatMedText(LinePtr)
                End If
                outString += OutLine + vbCrLf
                LinePtr += 1
                If LinePtr = SelRem.MatMedText.Length Then Exit While
            End While
            TextBox1.Text = outString
            Cursor = Cursors.Default    ' Change cursor to default.

            Exit Sub

        Catch
            Cursor = Cursors.Default    ' Change cursor to default.
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            Close()
        End Try
    End Sub


    Private Sub Print_Click(sender As Object, e As EventArgs) Handles Print.Click
        Form1.stringBuffer = ""   ' Clear out any previous print job

        Dim pdResult As DialogResult

        HandlePrint()
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
        End If
        Form1.PrintDialog1.Dispose()
    End Sub

    Public Sub HandlePrint()
        Dim ErrorFlag As Integer    ' Error number
        Dim msg As String           ' Error message string

        Try
            ErrorFlag = 2020

            'Print the page title.
            Form1.stringBuffer += "Remedy" + vbCrLf + vbLf

            ' print remedy directly from list box
            Form1.ErrorFlag = 2021
            Form1.stringBuffer += TextBox1.Text
        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
            Buttons:=vbOKOnly + vbCritical,
            Title:="ERROR!")
        End Try
    End Sub

    Private Sub Done_Click(sender As Object, e As EventArgs) Handles Done.Click
        Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim msg As String
        Dim myProcess As New Process

        Try
            myProcess.StartInfo.FileName = AppContext.BaseDirectory + Form1.HELP_DIR + "DisplayRemediesForm.htm"
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

Public Class DispRemDat
    Private myDispRemDesc As String
    Private myDispRemData As String

    Public Sub New(ByVal strDispRemDesc As String, ByVal strDispRemData As String)
        Me.myDispRemDesc = strDispRemDesc
        Me.myDispRemData = strDispRemData
    End Sub

    Public ReadOnly Property DispRemData() As String
        Get
            Return myDispRemData
        End Get
    End Property

    Public ReadOnly Property DispRemDesc() As String
        Get
            Return myDispRemDesc
        End Get
    End Property

End Class