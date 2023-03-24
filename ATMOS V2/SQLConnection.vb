Imports System.Data
Imports System.IO
Imports MySql.Data
Imports MySql.Data.MySqlClient

Public Class SQLConnection
    Private Shared connString As String
    Public Shared sqlConnection As New MySqlConnection

    Public itemListTB As New DataTable

    Public itemListDT As New DataTable

    Public Sub New()
        Dim filePath As String = Application.StartupPath & "\connstring.txt"

        If Not File.Exists(filePath) Then
            Using sw As StreamWriter = File.CreateText(filePath)
                sw.WriteLine("Server=localhost;Port=3306;User=root;Password=root;Database=atmos")
            End Using
        End If

        connString = My.Computer.FileSystem.ReadAllText(Application.StartupPath & "\connstring.txt")
        connectToDatabase()
    End Sub

    Public Shared Function executeQuery(ByVal query As String)
        Dim dataAdapter As MySqlDataAdapter = New MySqlDataAdapter(query, sqlConnection)
        Dim dataTable As New DataTable
        dataAdapter.Fill(dataTable)
        Return dataTable
    End Function

    Public Shared Sub connectToDatabase()
        If sqlConnection.State = ConnectionState.Closed Then
            sqlConnection.ConnectionString = connString
            Try
                If sqlConnection.State = ConnectionState.Closed Then
                    sqlConnection.Open()
                Else
                    MsgBox("Error!")
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub


    Public Shared Sub executeCommand(ByVal query As String)
        Dim mysqlCommand As New MySqlCommand
        With mysqlCommand
            .CommandText = query
            .Connection = sqlConnection
            .CommandType = CommandType.Text
            .ExecuteNonQuery()
        End With
    End Sub

    Public Shared Sub executePhotoCommand(ByVal query As String, ByVal photo As Object)
        Dim mysqlCommand As New MySqlCommand
        With mysqlCommand
            .CommandText = query
            .Connection = sqlConnection
            .Parameters.AddWithValue("@photo", photo)
            .CommandType = CommandType.Text
            .ExecuteNonQuery()
        End With
    End Sub

End Class
