Public Class Treshold
    Private Sub Treshold_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = My.Computer.FileSystem.ReadAllText(Application.StartupPath & "\treshold.txt")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        My.Computer.FileSystem.WriteAllText(Application.StartupPath & "\treshold.txt", TextBox1.Text, False)
        FaceRecognizer.fModel.distTreshold = Convert.ToInt64(TextBox1.Text)
        MsgBox("Treshold Changed!")
    End Sub
End Class