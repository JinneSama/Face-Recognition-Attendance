Public Class AdminPassword
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim dt As DataTable = SQLConnection.executeQuery("select * from staff where Type='Admin' and Name='" & MainForm.adminName & "' and password ='" & TextBox1.Text & "'")
        If dt.Rows.Count > 0 Then
            SQLConnection.executeCommand("Update staff set password='" & TextBox2.Text & "' where Type='Admin' and Name='" & MainForm.adminName & "'")
            TextBox1.Text = ""
            TextBox2.Text = ""
            MsgBox("Password Changed!" & vbCrLf & "You will now be Logged out of your Account!")
            Me.Close()
            MainForm.Close()
            LoginForm.Show()
        Else
            MsgBox("Old Password is Wrong!")
        End If
    End Sub
End Class