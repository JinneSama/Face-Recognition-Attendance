Imports System.IO
Imports System.Timers
Public Class CaptureImage
    Dim faceModel As FaceModel
    Public id As String

    Private pbList As List(Of PictureBox)

    Public images As ImageList
    Public imagesToSave As ImageList
    Public imageNames As New List(Of ListViewItem)
    Private imageCounter As Integer = 0

    Private imageCount As Integer
    Private imageMillisDelay As Integer

    Private type As String
    Dim timer As Timer

    Public Sub initFaceModel(s As String)
        type = s
    End Sub
    Private Sub BunifuImageButton1_Click(sender As Object, e As EventArgs) Handles BunifuImageButton1.Click
        faceModel.stopDetection()
        BunifuThinButton22.Enabled = True
        Me.Close()
    End Sub
    Private Sub CaptureImage_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        faceModel = New FaceModel(PictureBox1, PictureBox2, False)
    End Sub

    Private Sub BunifuThinButton22_Click(sender As Object, e As EventArgs) Handles BunifuThinButton22.Click
        faceModel.startDetection()
        BunifuThinButton22.Enabled = False
    End Sub
    Private Sub BunifuThinButton21_Click(sender As Object, e As EventArgs) Handles BunifuThinButton21.Click
        timer = New Timer
        timer.AutoReset = True

        AddHandler timer.Elapsed, AddressOf timer_Elapsed
        images = New ImageList
        imagesToSave = New ImageList
        imagesToSave.ImageSize = New Point(200, 200)

        ListView1.Items.Clear()

        images.ImageSize = New Point(64, 64)

        imageCount = NumericUpDown1.Value
        imageMillisDelay = 1000 / NumericUpDown2.Value

        timer.Interval = imageMillisDelay

        imageCounter = 0
        timer.Enabled = True
        'faceModel.saveTrainingImage(id & "_" & ComboBox1.Text)
    End Sub
    Private Sub CaptureImage_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ListView1.View = View.LargeIcon
    End Sub
    Private Sub CaptureImage_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        faceModel.stopDetection()
    End Sub

    Private Sub timer_Elapsed(sender As Object, e As ElapsedEventArgs)
        'imageNames.Add(New ListViewItem(id & "_" & e.SignalTime.ToString()) With {.ImageIndex = images.Images.Count - 1})
        Me.Invoke(Sub() AddImage(imageCounter))
        If imageCounter >= imageCount Then
            timer.Stop()
            'faceModel.stopDetection()
        Else
            imageCounter += 1
            Me.Invoke(Sub() UpdateLabel(imageCounter.ToString()))
        End If
    End Sub

    Private Sub UpdateLabel(s As String)
        Label1.Text = "Captured Images: " & s
    End Sub
    Private Sub AddImage(s As String)
        images.Images.Add(PictureBox2.Image)
        imagesToSave.Images.Add(PictureBox2.Image)
        ListView1.LargeImageList = images
        ListView1.Items.Add(New ListViewItem(id & "_" & s) With {.ImageIndex = images.Images.Count - 1})
        imageNames.Add(New ListViewItem(id & "_" & s) With {.ImageIndex = images.Images.Count - 1})
        'ListView1.Items.AddRange(imageNames.ToArray)
    End Sub

    Private Sub BunifuThinButton23_Click(sender As Object, e As EventArgs) Handles BunifuThinButton23.Click
        Dim _path As String = Directory.GetCurrentDirectory() & "\TrainingImages"
        For Each path In Directory.GetFiles(_path)
            If path.Contains(id) Then

            End If
        Next

        If type = "Register" Then
            MainForm.RefreshRegisterImage()
        Else
            MainForm.RefreshListImage()
        End If
        MsgBox("Added To List!")
    End Sub



    Private Sub BunifuThinButton24_Click(sender As Object, e As EventArgs) Handles BunifuThinButton24.Click
        If type = "Register" Then
            MainForm.profilePB.Image = faceModel.bmp
        Else
            MainForm.editProfilePB.Image = faceModel.bmp
        End If
    End Sub
End Class