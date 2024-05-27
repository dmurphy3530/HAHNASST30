Option Explicit On
Public Class SelRem
    Public Shared MatMedText() = System.IO.File.ReadAllLines(AppContext.BaseDirectory + Form1.DATA_DIR + "matmed.dat")
    Private Sub SelRem_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim LineCount As Integer    ' # of lines in MatMed.dat file
        Dim MatFileName As String   ' Name of Materia Medica file

        Try
            Command1.Enabled = False
            List1.BackColor = Form1.Lst1.BackColor
            List1.ForeColor = Form1.Lst1.ForeColor
            List1.Font = Form1.Lst1.Font

            ' Populate List1 from remedy title lines read from Matmed file

            MatFileName = AppContext.BaseDirectory + Form1.DATA_DIR + "matmed.dat"
            LineCount = Val(System.IO.File.ReadAllLines(MatFileName).Length.ToString())
            For Ptr1 = 0 To LineCount - 1
                ' Look for a line that contains just CR/LF with no text; remedy title line will be next array element.
                If MatMedText(Ptr1) = "" Then
                    List1.Items.Add(MatMedText(Ptr1 + 1))
                End If
            Next

            Exit Sub

        Catch
            Dim unused = MsgBox(Prompt:="matmed.dat file missing or corrupted; please re-install Dr. Hahnemann's Assistant to fix this problem",
            Buttons:=vbOKOnly + vbCritical,
            Title:="ERROR!")
            Close()
        End Try
    End Sub

    Private Sub List1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles List1.SelectedIndexChanged

    End Sub

    Private Sub List1_Click(sender As Object, e As EventArgs) Handles List1.Click
        Command1.Enabled = True
    End Sub
    Private Sub List1_DblClick(sender As Object, e As EventArgs) Handles List1.DoubleClick
        Command1.Enabled = True
        Call Select_Click()
    End Sub

    Private Sub Command1_Click(sender As Object, e As EventArgs) Handles Command1.Click
        Call Select_Click()
    End Sub

    Private Sub Select_Click()
        Dim SelPtr As Integer   ' selected remedy pointer

        SelPtr = -1
        Form1.RemedyNum = List1.SelectedIndex
        If DispRem.Visible Then DispRem.Close()      ' Need to re-initialize in case it is open.
        DispRem.Top = 0
        DispRem.Left = 0
        DispRem.Show()
    End Sub

    Private Sub Command2_Click(sender As Object, e As EventArgs) Handles Command2.Click ' Cancel
        'Close()
        Me.Visible = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim msg As String
        Dim myProcess As New Process

        Try
            myProcess.StartInfo.FileName = AppContext.BaseDirectory + Form1.HELP_DIR + "SelectRemediesForm.htm"
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