Option Explicit On

Public Class AutoWeightSeverity
    Public AutoWeightSev As Integer ' severity from Auto-Weight pop-up
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Form1.AutoWeightSev = 3
        Close()
    End Sub

    Private Sub AutoWeightSeverity_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = Form1.AutoWeightText
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form1.AutoWeightSev = 1
        Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form1.AutoWeightSev = 2
        Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim msg As String
        Dim myProcess As New Process

        Try
            myProcess.StartInfo.FileName = AppContext.BaseDirectory + Form1.HELP_DIR + "AutomaticallyAssignWeightCommand.htm"
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