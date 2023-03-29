Imports System.Net
Public Class FTPControl
    Private Shared serverUri As String = "ftp://192.168.100.2"
    Private Shared username As String = "user"
    Private Shared password As String = "user"
    Private Shared remoteFolderPath As String = ""
    Private Shared localFolderPath As String = Application.StartupPath & "/TrainingImages/"

    Public Shared Sub downloadAllFiles()
        Dim request As FtpWebRequest = WebRequest.Create(serverUri + remoteFolderPath)
        request.Method = WebRequestMethods.Ftp.ListDirectory
        request.Credentials = New NetworkCredential(username, password)

        Using response As FtpWebResponse = CType(request.GetResponse(), FtpWebResponse)
            Using streamReader As New IO.StreamReader(response.GetResponseStream())
                Dim files() As String = streamReader.ReadToEnd().Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                For Each file In files
                    If System.IO.File.Exists(localFolderPath & "\" & file) Then
                        Continue For
                    End If
                    If Not file.StartsWith(".") And file.Contains(".") Then

                        Dim remoteFilePath As String = remoteFolderPath + "/" + file
                        Dim localFilePath As String = localFolderPath + file
                        Dim downloadRequest As FtpWebRequest = WebRequest.Create(serverUri + remoteFilePath)
                        downloadRequest.Method = WebRequestMethods.Ftp.DownloadFile
                        downloadRequest.Credentials = New NetworkCredential(username, password)
                        Using downloadResponse As FtpWebResponse = CType(downloadRequest.GetResponse(), FtpWebResponse)
                            Using downloadStream As IO.Stream = downloadResponse.GetResponseStream()
                                Using fileStream As IO.FileStream = IO.File.Create(localFilePath)
                                    Dim buffer(1024) As Byte
                                    Dim bytesRead As Integer
                                    Do
                                        bytesRead = downloadStream.Read(buffer, 0, buffer.Length)
                                        If bytesRead > 0 Then
                                            fileStream.Write(buffer, 0, bytesRead)
                                        End If
                                    Loop While bytesRead > 0
                                End Using
                            End Using
                        End Using
                    End If
                Next
            End Using
        End Using
    End Sub

    Private Shared Function FileExistsOnServer(serverUri As String, username As String, password As String, fileName As String) As Boolean
        Dim request As FtpWebRequest = DirectCast(WebRequest.Create(serverUri + "/" + fileName), FtpWebRequest)
        request.Credentials = New NetworkCredential(username, password)
        request.Method = WebRequestMethods.Ftp.GetFileSize
        Try
            request.GetResponse()
            Return True
        Catch ex As WebException
            Dim response As FtpWebResponse = DirectCast(ex.Response, FtpWebResponse)
            If response.StatusCode = FtpStatusCode.ActionNotTakenFileUnavailable Then
                Return False
            End If
            Throw
        End Try
    End Function

    Public Shared Sub uploadAllFiles()
        Dim request As FtpWebRequest = DirectCast(WebRequest.Create(serverUri), FtpWebRequest)
        request.Credentials = New NetworkCredential(username, password)
        request.Method = WebRequestMethods.Ftp.UploadFile
        Dim fileList As String() = System.IO.Directory.GetFiles(localFolderPath)
        For Each file In fileList
            Dim fileName As String = System.IO.Path.GetFileName(file)
            Dim uploadUrl As String = serverUri + "/" + fileName
            Dim exists As Boolean = FileExistsOnServer(serverUri, username, password, fileName)
            If Not exists Then
                request = DirectCast(WebRequest.Create(uploadUrl), FtpWebRequest)
                Dim fileStream As System.IO.FileStream = System.IO.File.OpenRead(file)
                Dim buffer(1023) As Byte
                Dim bytesSent As Integer = fileStream.Read(buffer, 0, buffer.Length)
                request.Method = WebRequestMethods.Ftp.UploadFile

                request.Credentials = New NetworkCredential(username, password)
                Dim uploadStream As System.IO.Stream = request.GetRequestStream()

                Do While bytesSent <> 0
                    uploadStream.Write(buffer, 0, bytesSent)
                    bytesSent = fileStream.Read(buffer, 0, buffer.Length)
                Loop

                fileStream.Close()
                uploadStream.Close()
            Else
            End If
        Next
    End Sub
End Class