Public Class SetRoom
    Private Sub SetRoom_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim _query As String = "select distinct(Room) from subject"
        Dim roomDT As DataTable = SQLConnection.executeQuery(_query)
        For Each items As DataRow In roomDT.Rows
            ComboBox1.Items.Add(items.Item(0).ToString())
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MainForm._room = ComboBox1.Text
        Me.Hide()
        MainForm.Show()
    End Sub
End Class