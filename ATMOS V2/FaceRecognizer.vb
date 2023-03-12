Public Class FaceRecognizer
    Public Shared fModel As FaceModel
    Public Shared Sub init()
        fModel = New FaceModel(Nothing, Nothing, False)
    End Sub
End Class
