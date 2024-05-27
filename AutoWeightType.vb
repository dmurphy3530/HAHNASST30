Public Class AutoWeightType

    Private Sub AutoWeightType_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = Form1.AutoWeightText
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form1.AutoWeightTp = 1
        Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form1.AutoWeightTp = 2
        Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Form1.AutoWeightTp = 3
        Close()
    End Sub
End Class