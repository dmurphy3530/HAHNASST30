Option Explicit On
Public Class SetAutowt
    Public Sub OpenFileReadRandom(FileName As String, FileNumber As Integer, RecLen As Integer, FileStatus As Boolean)
        On Error Resume Next
        FileStatus = True   ' Assume success.
        Err.Description = ""
        FileOpen(FileName, FileNumber, OpenMode.Random, RecordLength:=RecLen)
        If Err.Description = "File already open" Then
            FileSystem.FileClose(FileNumber)
            Err.Description = ""
            FileOpen(FileName, FileNumber, OpenMode.Random, RecordLength:=RecLen)
            If Err.Description = "File Already Open" Then FileStatus = False
        End If
    End Sub

    Public Sub OpenFileWrite(FileName As String, FileNumber As Integer, FileStatus As Boolean)
        On Error Resume Next
        FileStatus = True   ' Assume success.
        Err.Description = ""
        FileOpen(FileName, FileNumber, OpenMode.Output)
        If Err.Description = "File already open" Then
            FileSystem.FileClose(FileNumber)
            Err.Description = ""
            FileOpen(FileName, FileNumber, OpenMode.Output)
            If Err.Description = "File Already Open" Then FileStatus = False
        End If
    End Sub


    Public Sub OpenFileRead(FileName As String, FileNumber As Integer, FileStatus As Boolean)
        On Error Resume Next
        FileStatus = True   ' Assume success.
        Err.Description = ""
        FileOpen(FileName, FileNumber, OpenMode.Input)
        If Err.Description = "File already open" Then
            FileSystem.FileClose(FileNumber)
            Err.Description = ""
            FileOpen(FileName, FileNumber, OpenMode.Input)
            If Err.Description = "File Already Open" Then FileStatus = False
        End If
    End Sub

    Public Function GetAutoweight(Symptom As String) As Integer

        Dim AFileWord As String ' AutoWt.dat word
        Dim AW As Integer   ' weight determined for Symptom
        Dim AWStr As String ' string form of AW
        Dim DONE As Boolean     ' weight was found flag
        Dim First As Integer    ' top of binary search
        Dim Found As Boolean    ' word was found flag
        Dim Last As Integer     ' bottom of binary search
        Dim loc_Afile_pos As Integer    ' local search file position
        Dim locLineDat As String    ' line read from AutoWt.dat
        Dim msg As String
        Dim Ptr1 As Integer   ' string pointer
        Dim SympWord As String  ' word being looked up
        Dim AutoWtText() = System.IO.File.ReadAllLines(AppContext.BaseDirectory + Form1.DATA_DIR + "AutoWt.dat")

        Try
            Form1.ErrorFlag = 10001
            AW = 0
            GetAutoweight = 0

            '   Do a binary search of AutoWt.dat for words in the string.
            DONE = False
            Ptr1 = 0
            While (Not DONE Or AW = 2) And Ptr1 < Len(Symptom)
                Ptr1 = InStr(Symptom, " ") - 1
                If Ptr1 = -1 Then Ptr1 = Len(Symptom)
                SympWord = LCase(Mid(Symptom, 1, Ptr1))
                If SympWord = "" Then
                    Ptr1 = Ptr1 + 1
                    SympWord = LCase(Mid(Symptom, 1, Ptr1))
                End If

                If Ptr1 <> Len(Symptom) Then Symptom = Mid(Symptom, Ptr1 + 2)
                If Mid(SympWord, Len(SympWord), 1) = "," Or
               Mid(SympWord, Len(SympWord), 1) = ";" Or
               Mid(SympWord, Len(SympWord), 1) = "." Or
               Mid(SympWord, Len(SympWord), 1) = ")" Then _
               SympWord = Mid(SympWord, 1, Len(SympWord) - 1)
                Last = Form1.Afile_size
                First = 0

                ' Find text position in file.
                Found = False
                AW = 0
                While Not Found
                    loc_Afile_pos = (First + Last + 0.1) / 2
                    If (loc_Afile_pos < 1) Then loc_Afile_pos = 1
                    If (loc_Afile_pos >= Form1.Afile_size) Then _
                    loc_Afile_pos = Form1.Afile_size - 1
                    locLineDat = AutoWtText(loc_Afile_pos)

                    If Last - First <= 1 Then
                        If (First > 0) Then
                            loc_Afile_pos = First
                            locLineDat = AutoWtText(loc_Afile_pos)
                            AFileWord = Mid(locLineDat, 1, InStr(locLineDat, " ") - 1)
                            If SympWord <> AFileWord Then _
                            loc_Afile_pos = Last
                        End If
                        Found = True
                        AWStr = Mid(locLineDat, InStr(locLineDat, " ") + 1, 1)
                        AW = CInt(AWStr)
                    End If
                    AFileWord = Mid(locLineDat, 1, InStr(locLineDat, " ") - 1)
                    If SympWord = AFileWord Then
                        Found = True
                        AWStr = Mid(locLineDat, InStr(locLineDat, " ") + 1, 1)
                        AW = CInt(AWStr)
                    ElseIf SympWord > AFileWord Then
                        First = (First + Last + 0.1) / 2
                    Else
                        Last = (First + Last + 0.1) / 2
                    End If
                End While
                If AW <> 0 Then DONE = True
            End While
            GetAutoweight = AW
            Exit Function

        Catch
            GetAutoweight = 0
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            Exit Function
        End Try
    End Function
    Private Sub SetAutowt_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Command1_Click(sender As Object, e As EventArgs) Handles Command1.Click
        Dim AW As Integer
        Dim FileStatus As Boolean   ' true = OK
        Dim LastAW As Integer
        Dim locLineDat As String       ' file line
        Dim msg As String       ' error message string
        Dim OldAW As Integer
        Dim OldLineDat As String
        Dim Pt As Integer
        Dim RemedyNumber As Integer
        Dim SympCount1, SympCount2, SympCount3, SympCount4 As Integer
        Dim Temp As String

        Try
            Form1.ErrorFlag = 10002
            OldLineDat = " "   ' Avoid warning variable used before assigning value

            Call OpenFileRead(AppContext.BaseDirectory + "\Matmed.dat", 1, FileStatus)
            Call OpenFileWrite(AppContext.BaseDirectory + "\NSympcnt.txt", 2, FileStatus)
            Seek(1, 1)
            locLineDat = LineInput(1)
            RemedyNumber = 1
            While Not (EOF(1))
                locLineDat = LineInput(1)
                SympCount1 = 0
                SympCount2 = 0
                SympCount3 = 0
                SympCount4 = 0
                AW = 0
                OldAW = 0
                While Len(locLineDat) > 0
                    Try
                        If Mid(locLineDat, 1, 1) = "+" Then
                            locLineDat = Mid(locLineDat, 2)
                            While Mid(locLineDat, 1, 1) = Chr(9)
                                locLineDat = Mid(locLineDat, 2)
                            End While
                            locLineDat = OldLineDat + " " + locLineDat
                            If OldAW = 2 Then AW = GetAutoweight(locLineDat)
                            If InStr(locLineDat, "**") Then
                                SympCount2 = SympCount2 + 3
                                If OldAW = 2 Then SympCount3 = SympCount3 - LastAW + 4 * (3 + AW)
                                locLineDat = Mid(locLineDat, 1, InStr(locLineDat, "**") - 1)
                            ElseIf InStr(locLineDat, "*") Then
                                SympCount2 = SympCount2 + 1
                                If OldAW = 2 Then SympCount3 = SympCount3 - LastAW + 2 * (3 + AW)
                                locLineDat = Mid(locLineDat, 1, InStr(locLineDat, "*") - 1)
                            ElseIf OldAW = 2 Then
                                SympCount3 = SympCount3 - LastAW + 3 + AW
                            End If
                            If OldAW = 2 Then SympCount4 = SympCount4 - LastAW + 3 + AW
                        ElseIf Mid(locLineDat, 1, 1) = Chr(9) Then
                            If Mid(locLineDat, 1, 2) = Chr(9) + Chr(9) Then
                                While Mid(locLineDat, 1, 1) = Chr(9)
                                    locLineDat = Mid(locLineDat, 2)
                                End While
                                locLineDat = OldLineDat + " " + locLineDat
                            End If
                            If Mid(locLineDat, 1, 1) = Chr(9) Then locLineDat = Mid(locLineDat, 2)

                            SympCount1 = SympCount1 + 1
                            Pt = InStr(locLineDat, "(")
                            If Pt <> 0 Then
                                If Pt = 1 Then
                                    locLineDat = Mid(locLineDat, 2)
                                Else
                                    locLineDat = Mid(locLineDat, 1, Pt - 1) + Mid(locLineDat, Pt + 1)
                                End If
                            End If
                            Pt = InStr(locLineDat, ")")
                            If Pt <> 0 Then
                                If Pt = 1 Then
                                    locLineDat = Mid(locLineDat, 2)
                                Else
                                    locLineDat = Mid(locLineDat, 1, Pt - 1) + Mid(locLineDat, Pt + 1)
                                End If
                            End If
                            Temp = locLineDat
                            AW = GetAutoweight(locLineDat)
                            locLineDat = Temp
                            OldAW = AW
                            If InStr(locLineDat, "**") Then
                                SympCount2 = SympCount2 + 4
                                SympCount3 = SympCount3 + 4 * (3 + AW)
                                LastAW = 4 * (3 + AW)
                                locLineDat = Mid(locLineDat, 1, InStr(locLineDat, "**") - 1)
                            ElseIf InStr(locLineDat, "*") Then
                                SympCount2 = SympCount2 + 2
                                SympCount3 = SympCount3 + 2 * (3 + AW)
                                LastAW = 2 * (3 + AW)
                                locLineDat = Mid(locLineDat, 1, InStr(locLineDat, "*") - 1)
                            Else
                                SympCount2 = SympCount2 + 1
                                SympCount3 = SympCount3 + 3 + AW
                                LastAW = 3 + AW
                            End If
                            SympCount4 = SympCount4 + 3 + AW
                        End If
                        OldLineDat = locLineDat
                        locLineDat = LineInput(1)
                    Catch
                        Exit While
                    End Try
                End While

                Print(RemedyNumber, SympCount1, SympCount2, SympCount3, SympCount4)
                RemedyNumber = RemedyNumber + 1
            End While
            FileSystem.FileClose(1)
            FileSystem.FileClose(2)

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(Form1.ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub Command2_Click(sender As Object, e As EventArgs) Handles Command2.Click
        Close()
    End Sub
End Class