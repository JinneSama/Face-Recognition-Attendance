Imports AForge
Imports AForge.Video
Imports AForge.Video.DirectShow
Imports System.IO
Imports System.Threading
Imports Emgu.CV
Imports Emgu.CV.Face
Imports Emgu.CV.Util
Imports Emgu.CV.CvEnum
Imports Emgu.CV.Structure
Public Class FaceModel

    Public cam As VideoCapture
    Public captureMat As New Mat
    'Private filter As FilterInfoCollection
    'Private cam As VideoCaptureDevice

    Private pen As Pen = New Pen(Color.Green, 2)
    Private gImg As Image(Of Bgr, Byte)
    Private m As Bitmap

    Dim rects As Rectangle()
    Public bmp As Bitmap
    Shared ReadOnly classifier As CascadeClassifier = New CascadeClassifier("haarcascade_frontalface_alt.xml")

    Private stream As MJPEGStream

    Private resultImage As Image(Of Bgr, Byte)

    Private model As LBPHFaceRecognizer
    Private trainedImages As New VectorOfMat
    Private faceLabelVector As New VectorOfInt
    Private faceLabels As New List(Of Integer)
    Private labelNames As New List(Of String)

    Public isTrained As Boolean = False
    Public recordedStudents As New List(Of String)
    Public _pb As PictureBox
    Public _facePB As PictureBox
    Public useTrained As Boolean
    Public Event recordChanged(sender As Object, e As EventArgs)

    Public distTreshold As Integer = Convert.ToInt64(My.Computer.FileSystem.ReadAllText(Application.StartupPath & "\treshold.txt"))

    Public Sub New(ByRef pb As PictureBox, ByRef facePB As PictureBox, ByVal trained As Boolean)
        _pb = pb
        _facePB = facePB
        useTrained = trained
    End Sub

    Public Sub startDetection()
        recordedStudents.Clear()

        If IsNothing(cam) Then
            cam = New VideoCapture(0, VideoCapture.API.DShow)
            cam.SetCaptureProperty(CapProp.FourCC, VideoWriter.Fourcc("M", "J", "P", "G"))
            AddHandler cam.ImageGrabbed, AddressOf camFrame
        End If
        cam.Start()
    End Sub

    Public Sub stopDetection()
        If Not IsNothing(cam) Then
            cam.Stop()
            cam.Dispose()
        End If
    End Sub

    Public Sub pauseDetection()
        If Not IsNothing(cam) Then
            cam.Pause()
        End If
    End Sub

    Public Sub saveTrainingImage(ByVal lbl As String)
        System.IO.Directory.CreateDirectory("TrainingImages")
        Dim path As String = Directory.GetCurrentDirectory() & "\TrainingImages"
        Dim grayImageResult As New Mat
        CvInvoke.CvtColor(resultImage, grayImageResult, ColorConversion.Bgr2Gray)
        CvInvoke.EqualizeHist(grayImageResult, grayImageResult)
        Dim saveGrayImage As Image(Of Gray, Byte) = grayImageResult.ToImage(Of Gray, Byte)
        If File.Exists(path & "\" + lbl + ".jpg") Then
            File.Delete(path & "\" + lbl + ".jpg")
        End If
        saveGrayImage.Save(path & "\" + lbl + ".jpg")
    End Sub

    Public Sub camFrame(sender As Object, e As EventArgs)
        cam.Read(captureMat)
        If captureMat.IsEmpty Then
            Return
        End If
        bmp = captureMat.ToImage(Of Bgr, Byte).ToBitmap()
        bmp.RotateFlip(RotateFlipType.Rotate180FlipY)
        gImg = New Image(Of Bgr, Byte)(bmp)
        gImg = gImg.Resize(320, 240, Inter.Cubic)

        Dim grayImage As New Mat
        CvInvoke.CvtColor(gImg, grayImage, ColorConversion.Bgr2Gray)
        CvInvoke.EqualizeHist(grayImage, grayImage)

        rects = classifier.DetectMultiScale(grayImage, 1.2, 3)

        For Each rect As Rectangle In rects

            resultImage = gImg.Convert(Of Bgr, Byte)
            resultImage.ROI = rect
            resultImage = resultImage.Resize(200, 200, Inter.Cubic)

            If isTrained And useTrained Then
                Dim grayImageResult As New Mat
                CvInvoke.CvtColor(resultImage, grayImageResult, ColorConversion.Bgr2Gray)
                CvInvoke.EqualizeHist(grayImageResult, grayImageResult)
                Dim PR = model.Predict(grayImageResult)
                '  CvInvoke.PutText(gImg, PR.Distance.ToString(), New Drawing.Point(2, 15), FontFace.HersheyComplex, 0.5, New Bgr(Color.Orange).MCvScalar)
                If PR.Label > 0 And PR.Distance < distTreshold Then
                    CvInvoke.PutText(gImg, labelNames(PR.Label), rect.Location + New Drawing.Point(0, -3), FontFace.HersheyComplex, 0.5, New Bgr(Color.Orange).MCvScalar)
                    CvInvoke.Rectangle(gImg, rect, New Bgr(Color.Green).MCvScalar, 2)
                    If Not recordedStudents.Contains(labelNames(PR.Label)) Then
                        recordedStudents.Add(labelNames(PR.Label))
                        RaiseEvent recordChanged(Me, New EventArgs())
                        'Threading.Thread.Sleep(500)
                    End If
                Else
                    CvInvoke.PutText(gImg, "Unknown", rect.Location + New Drawing.Point(0, -3), FontFace.HersheyComplex, 0.5, New Bgr(Color.Orange).MCvScalar)
                    CvInvoke.Rectangle(gImg, rect, New Bgr(Color.Red).MCvScalar, 2)
                End If
            Else
                Dim graphics As Graphics = Graphics.FromImage(gImg.Bitmap)
                graphics.DrawRectangle(pen, rect)
            End If
        Next
        If Not IsNothing(_facePB) Then
            If IsNothing(resultImage) Then
                _facePB.Image = Nothing
            Else
                Dim grayImageResult As New Mat
                CvInvoke.CvtColor(resultImage, grayImageResult, ColorConversion.Bgr2Gray)
                CvInvoke.EqualizeHist(grayImageResult, grayImageResult)
                _facePB.Image = grayImageResult.Bitmap
            End If
        End If
        _pb.Image = gImg.Bitmap
    End Sub

    Public Sub TrainModel()
        FaceRecog.PictureBox1.BackColor = Color.Red
        Dim imgCount As Integer = 0
        Dim eigenTreshold As Double = 30000

        Dim path As String = Directory.GetCurrentDirectory() + "\TrainingImages"
        Dim files As String() = Directory.GetFiles(path, "*.jpg", SearchOption.AllDirectories)

        Dim Iterations As Integer = 20
        faceLabels.Clear()
        labelNames.Clear()
        For i As Integer = 1 To Iterations
            For Each file As String In files
                Dim names As String() = file.Split("\\").Last().Split("_")
                Dim name As String = names(0)
                trainedImages.Push(CvInvoke.Imread(file, ImreadModes.Grayscale))
                faceLabels.Add(imgCount)
                labelNames.Add(name)
                imgCount += 1
            Next
        Next
        If imgCount > 0 And Not isTrained Then
            faceLabelVector.Push(faceLabels.ToArray())
            model = New LBPHFaceRecognizer()
            model.Train(trainedImages, faceLabelVector)
            isTrained = True
        End If
    End Sub
End Class
