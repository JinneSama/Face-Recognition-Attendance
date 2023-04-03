Imports System.Threading
Public Class LoginForm
    Private pbVal = 0

    Private sqlConnection As SQLConnection
    Private Sub BunifuThinButton21_Click(sender As Object, e As EventArgs) Handles BunifuThinButton21.Click
        startLogin()
    End Sub

    Private Sub startLogin()
        Dim query As String = "select * from staff where username='" & BunifuMaterialTextbox1.Text & "' and password='" & BunifuMaterialTextbox2.Text & "' and Type='" & ComboBox1.Text & "'"
        Dim datatable As DataTable = SQLConnection.executeQuery(query)
        If datatable.Rows.Count >= 1 Then
            If ComboBox1.Text = "Admin" Then
                MainForm.loginName = ""
                MainForm.adminName = BunifuMaterialTextbox1.Text

                MainForm.FACULTYToolStripMenuItem.Visible = True
                MainForm.STUDENTToolStripMenuItem.Visible = True
                MainForm.CHANGEPWORDToolStripMenuItem.Visible = True
                MainForm.CAPTUREToolStripMenuItem.Visible = False
            Else
                MainForm.loginName = BunifuMaterialTextbox1.Text

                MainForm.CAPTUREToolStripMenuItem.Visible = False
                MainForm.FACULTYToolStripMenuItem.Visible = False
                MainForm.STUDENTToolStripMenuItem.Visible = False
                MainForm.CHANGEPWORDToolStripMenuItem.Visible = False
                MainForm.ComboBox1.Text = ComboBox1.Text
                MainForm.ComboBox1.Enabled = False
            End If
            MainForm.loginType = ComboBox1.Text
            BunifuTransition1.HideSync(Panel1)
            BunifuTransition1.ShowSync(Panel2)


            BunifuMaterialTextbox1.Text = ""
            BunifuMaterialTextbox2.Text = ""
            Timer1.Start()
        Else
            MsgBox("Please Check the Credentials you Entered!")
            BunifuMaterialTextbox1.Text = ""
            BunifuMaterialTextbox2.Text = ""
            ComboBox1.Text = ""
        End If
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        pbVal += 5
        If pbVal > 5 Then
            Timer1.Stop()


            Panel2.Hide()
            Panel1.Show()
            MainForm.Show()
            Me.Hide()
        End If
        BunifuCircleProgressbar1.Value = pbVal
    End Sub

    Private Sub LoginForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        sqlConnection = New SQLConnection
        FaceRecognizer.init()
        FTPControl.initServerURI()
    End Sub

    Private Sub BunifuMaterialTextbox2_OnValueChanged(sender As Object, e As EventArgs) Handles BunifuMaterialTextbox2.OnValueChanged
        BunifuMaterialTextbox2.isPassword = True
    End Sub
    Private Sub LoginForm_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MyBase.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            startLogin()
        End If
    End Sub
End Class
