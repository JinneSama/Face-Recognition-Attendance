Imports BunifuAnimatorNS

Public Class FaceRecog
    Public subjectCode As String = ""
    Public studentList As New List(Of String)
    Private studentAdded As New List(Of String)
    Private Delegate Sub UpdateListInvoker(ByVal StringList As List(Of String))

    Private SchedList As New List(Of DataRow)

    Private timeOutCamera As DateTime
    Private _startDetection As Boolean = False
    Private Sub FaceRecog_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        BackgroundWorker1.WorkerSupportsCancellation = True
        BunifuThinButton21.Enabled = True
        Panel2.Hide()
        Panel1.Show()
    End Sub

    Private Sub FaceRecog_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        BunifuCustomDataGrid4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        FaceRecognizer.fModel._pb = PictureBox1
        FaceRecognizer.fModel.useTrained = True
        'BackgroundWorker1.RunWorkerAsync()
        AddHandler FaceRecognizer.fModel.recordChanged, AddressOf updateList
    End Sub

    Public Sub updateList()
        'PictureBox3.Image = My.Resources.angryimg
        'If Me.PictureBox3.InvokeRequired Then
        'Me.PictureBox3.Invoke(New UpdateListInvoker(AddressOf updateList), faceModel.recordedStudents)
        'Else

        If Me.IDLB.InvokeRequired Then
            IDLB.Invoke(New UpdateListInvoker(AddressOf updateList), FaceRecognizer.fModel.recordedStudents)
            Return
        End If
        If Not studentList.Contains(FaceRecognizer.fModel.recordedStudents.Last.ToString()) Then
            Return
        End If
        Dim query As String = "Insert into attendance (Student_Code , Student_Subject , Time_Taken) values ('" & FaceRecognizer.fModel.recordedStudents.Last.ToString() &
                "','" & subjectCode & "','" & DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") & "')"
        SQLConnection.executeQuery(query)
        query = "Select * from student where ID='" & FaceRecognizer.fModel.recordedStudents.Last.ToString() & "'"
        Dim dtFill = SQLConnection.executeQuery(query)
        Dim nms As New IO.MemoryStream(CType(dtFill.Rows().Item(0).Item(4), Byte()))
        Dim nreturnImage As Image = Image.FromStream(nms)

        If Not IsNothing(nreturnImage) Then
            PictureBox3.Image = nreturnImage
        End If

        If Not studentAdded.Contains(dtFill.Rows().Item(0).Item(1).ToString()) Then
            studentAdded.Add(dtFill.Rows().Item(0).Item(1).ToString())
        End If

        IDLB.Text = dtFill.Rows().Item(0).Item(0).ToString()
            NameLB.Text = dtFill.Rows().Item(0).Item(1).ToString()
            ListBox1.Items.Clear()

        For Each items As String In studentAdded
            ListBox1.Items.Add(items)
        Next

        'End If
    End Sub


    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        'faceModel.TrainModel()
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        BunifuTransition1.HideSync(Panel2)
        BunifuTransition1.ShowSync(Panel1)
    End Sub

    Private Sub BunifuImageButton1_Click(sender As Object, e As EventArgs) Handles BunifuImageButton1.Click
        Me.Hide()
        MainForm.Show()
    End Sub

    Private Sub BunifuThinButton21_Click(sender As Object, e As EventArgs) Handles BunifuThinButton21.Click
        BunifuThinButton21.Enabled = False
        Dim query As String = "select Distinct Room from subject"
        Select Case DateTime.Now.DayOfWeek.ToString()
            Case "Monday"
                query &= " where Mon='1'"
            Case "Tuesday"
                query &= " where Tue='1'"
            Case "Wednesday"
                query &= " where Wed='1'"
            Case "Thursday"
                query &= " where Thu='1'"
            Case "Friday"
                query &= " where Fri='1'"
            Case "Saturday"
                query &= " where Sat='1'"
            Case "Sunday"
                query &= " where Sun='1'"

        End Select
        Dim dt As DataTable = SQLConnection.executeQuery(query)
        BunifuCustomDataGrid4.DataSource = dt
        BunifuTransition1.ShowSync(Panel3, False, Animation.HorizSlide)
    End Sub

    Private Sub BunifuThinButton22_Click(sender As Object, e As EventArgs) Handles BunifuThinButton22.Click
        SchedList.Clear()
        FaceRecognizer.fModel.pauseDetection()
        BunifuThinButton21.Enabled = True
    End Sub

    Private Sub BunifuThinButton219_Click(sender As Object, e As EventArgs) Handles BunifuThinButton219.Click
        BunifuTransition1.HideSync(Panel3, False, Animation.HorizSlide)
    End Sub

    Private Sub BunifuThinButton220_Click(sender As Object, e As EventArgs) Handles BunifuThinButton220.Click
        If Not FaceRecognizer.fModel.isTrained Then
            MsgBox("Please Wait for the Model to be Trained")
            Return
        End If
        Dim nquery As String = "select name,code,time_in,time_out from subject where Room='" & BunifuCustomDataGrid4.SelectedRows.Item(0).Cells.Item(0).Value.ToString() & "'"
        Dim newDT2 As DataTable = SQLConnection.executeQuery(nquery)
        Label3.Text = "Waiting for Subject"
        For Each item In newDT2.Rows
            SchedList.Add(item)
        Next

        For Each item In SchedList.ToList()
            If Convert.ToDateTime(item.Item(3).ToString()) < DateTime.Now Then
                SchedList.Remove(item)
            End If
        Next
        BunifuTransition1.HideSync(Panel3)
        BackgroundWorker2.WorkerSupportsCancellation = True
        BackgroundWorker2.RunWorkerAsync()
    End Sub

    Private Sub BackgroundWorker2_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        While Not SchedList.Count = 0
            For Each item In SchedList.ToList()
                If Convert.ToDateTime(item.Item(3).ToString()) > DateTime.Now And Convert.ToDateTime(item.Item(2).ToString()) < DateTime.Now Then
                    subjectCode = item.Item(1).ToString()
                    Me.Invoke(Sub() refreshLabel(item.Item(0).ToString()))
                    If Not _startDetection Then
                        Me.Invoke(Sub() startDetection(item.Item(1).ToString()))
                        
                        _startDetection = True
                    End If
                End If

                If Convert.ToDateTime(item.Item(3).ToString()) < DateTime.Now Then
                    Me.Invoke(Sub() refreshLabel("Waiting for next Subject"))
                    SQLConnection.executeCommand("Insert into attendanceschedule (subjectCode, dateSet) values ('" & item.Item(1).ToString() & "','" & DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") & "')")
                    SchedList.Remove(item)
                    _startDetection = False
                    FaceRecognizer.fModel.pauseDetection()
                End If
            Next
        End While
    End Sub

    Private Sub startDetection(ByVal code As String)
        Dim query As String = "Select student_id from student_subjects where subject_id ='" & subjectCode & "'"
        Dim newDT As DataTable = SQLConnection.executeQuery(query)
        studentList.Clear()
        For i As Integer = 0 To newDT.Rows.Count - 1
            studentList.Add(newDT.Rows.Item(i).Item(0).ToString)
        Next
        FaceRecognizer.fModel.recordedStudents.Clear()
        FaceRecognizer.fModel.startDetection()
    End Sub

    Private Sub refreshLabel(ByVal txt As String)
        Label3.Text = "Now Checking Attendance : " & txt
    End Sub

    Private Sub FaceRecog_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        SchedList.Clear()
        If BackgroundWorker2.IsBusy Then
            BackgroundWorker2.CancelAsync()
            e.Cancel = True
        End If
        FaceRecognizer.fModel.stopDetection()
    End Sub
End Class