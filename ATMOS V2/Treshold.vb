Public Class Treshold
    Private Sub Treshold_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = My.Computer.FileSystem.ReadAllText(Application.StartupPath & "\treshold.txt")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim q As String = "select * from staff where Type='Admin' and Name='" & MainForm.adminName & "' and password ='" & TextBox2.Text & "'"
        Dim dt As DataTable = SQLConnection.executeQuery(q)

        If dt.Rows.Count > 0 Then
            My.Computer.FileSystem.WriteAllText(Application.StartupPath & "\treshold.txt", TextBox1.Text, False)
            FaceRecognizer.fModel.distTreshold = Convert.ToInt64(TextBox1.Text)
            TextBox2.Text = ""
            MsgBox("Treshold Changed!")
        Else
            MsgBox("Invalid Password!")
        End If
    End Sub
End Class