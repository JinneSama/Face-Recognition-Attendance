Imports BunifuAnimatorNS
Imports AForge
Imports AForge.Video
Imports AForge.Video.DirectShow
Imports System.IO
Imports System.Text
Imports System.Drawing
Imports System.Data

Public Class MainForm
    Private videocapture As VideoCaptureDevice
    Public loginType As String
    Public loginName As String = ""
    Public adminName As String = ""
    Private subjectID As String
    Private dt As DataTable
    Private dtStaff As DataTable
    Private dtFaculty As DataTable
    Private dtsubject As DataTable
    Private bmp As Bitmap
    Private FacultyList As DataRowCollection
    Private SubjectList As DataRowCollection
    Private FacultyAssignList As DataRowCollection
    Private editPW As String
    Private LogoutState As Boolean = False
    Private editSubjectState As Boolean = False

    Private facemodel As FaceModel
    Private closeForm As Boolean = False
    Private switchExport As Boolean = False

    Public images As ImageList
    Public imagesToSave As ImageList
    Public imageNames As New List(Of ListViewItem)
    Private cbBoxContents As List(Of personDetail)
    Private _dtForSubject As DataTable
    Public _room As String = ""

    Private Sub BunifuImageButton3_Click(sender As Object, e As EventArgs)
        BunifuTransition1.HideSync(registerStudentPanel)
    End Sub

    Private Function checkDuplicateStudentD(id As String)
        Dim dup As Boolean = False
        Dim idDT As DataTable = SQLConnection.executeQuery("Select ID from student")
        For Each items As DataRow In idDT.Rows
            If id = items.Item(0).ToString() Then
                dup = True
            End If
        Next
        Return dup
    End Function

    Private Sub BunifuThinButton22_Click(sender As Object, e As EventArgs) Handles BunifuThinButton22.Click
        If checkDuplicateStudentD(idTB.Text) Then
            MsgBox("ID has Duplicate!")
            Return
        End If
        If yearTB.Text = "" Or fnameTB.Text = "" Or lNameTB.Text = "" Or yearTB.Text = "" Then
            MsgBox("Please Fill all Required Fields First!")
            Return
        End If
        If profilePB Is Nothing Then
            MsgBox("capture profile first!")
            Return
        End If
        Dim fname As String = fnameTB.Text
        Dim mname As String = mNameTB.Text
        Dim lname As String = lNameTB.Text
        Dim id As String = idTB.Text
        Dim course As String = courseTB.Text
        Dim year As String = yearTB.Text
        Dim query As String = "Insert into student (ID , Name , Course , Year) values ('" & id & "','" & fname & " " & mname & " " & lname & "','" & course & "','" & year & "')"
        SQLConnection.executeQuery(query)
        dt = SQLConnection.executeQuery("select Name as 'Full Name' , ID as 'ID Number' , CONCAT(Course,' ' ,Year) as 'Course and Year' from student")
        BunifuCustomDataGrid1.DataSource = dt
        saveImage()
        defaultAddStudent()

        For i As Integer = 0 To ListView1.Items.Count - 1
            saveTrainingImage(ListView1.Items(i).Text, imagesToSave.Images(i))
        Next
        MsgBox("Student Added!")
        FTPControl.uploadAllFiles()
        BackgroundWorker2.RunWorkerAsync()

    End Sub

    Private Sub defaultAddStudent()
        fnameTB.Text = ""
        mNameTB.Text = ""
        lNameTB.Text = ""
        idTB.Text = ""
        yearTB.Text = ""


        profilePB.Image = Nothing
    End Sub
    Private Sub saveImage()
        If profilePB.Image Is Nothing Then
            MsgBox("Capture your Profile Photo First")
            Return
        End If
        Dim newImage As Image = profilePB.Image

        Dim picStream As New System.IO.MemoryStream()
        newImage.Save(picStream, System.Drawing.Imaging.ImageFormat.Png)
        Dim imgArray() As Byte = picStream.GetBuffer()
        picStream.Close()

        Dim photoQuery As String = "Update student Set img=@photo where id = '" & idTB.Text & "'"
        SQLConnection.executePhotoCommand(photoQuery, imgArray)
    End Sub
    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'FTPControl.downloadCertain(_room)
        SetStyle(ControlStyles.SupportsTransparentBackColor, True)
        Me.BackColor = Color.Transparent

        BunifuCustomDataGrid2.RowHeadersVisible = False
        facemodel = New FaceModel(Nothing, Nothing, True)
        BunifuCustomDataGrid1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        assignStudentSubjectDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        listofStudentSubjectDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        dt = SQLConnection.executeQuery("select Name as 'Full Name' , ID as 'ID Number' , CONCAT(Course,' ' ,Year) as 'Course and Year' from student")
        BunifuCustomDataGrid1.DataSource = dt

        DateTimePicker3.CustomFormat = "hh:mm tt"
        DateTimePicker3.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        DateTimePicker3.ShowUpDown = True

        DateTimePicker4.CustomFormat = "hh:mm tt"
        DateTimePicker4.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        DateTimePicker4.ShowUpDown = True

        BunifuCustomDataGrid2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        BunifuCustomDataGrid2.Columns.Add("No.", "Num")
        BunifuCustomDataGrid2.Columns.Add("Full Name", "Name")
        BunifuCustomDataGrid2.Columns.Add("Username", "User")
        'BunifuCustomDataGrid2.Rows.Clear()

        'BunifuCustomDataGrid2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        'dtStaff = SQLConnection.executeQuery("select name,id,username,password from staff where type = 'faculty'")
        'BunifuCustomDataGrid2.DataSource = dtStaff



        BunifuCustomDataGrid4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        BunifuCustomDataGrid3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        BunifuCustomDataGrid4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        BunifuCustomDataGrid5.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        Dim dtNew As DataTable = SQLConnection.executeQuery("select name,code,TIME_FORMAT(time_in, '%I:%i %p') as Time_In ,TIME_FORMAT(time_out, '%I:%i %p') as Time_Out, Room from subject")
        BunifuCustomDataGrid5.DataSource = dtNew

        HideAllPanel()
    End Sub

    Private Sub BunifuThinButton21_Click(sender As Object, e As EventArgs) Handles BunifuThinButton21.Click
        If idTB.Text = "" Then
            MsgBox("Please Fill all Required Fields First!")
        Else
            CaptureImage.id = idTB.Text
            CaptureImage.initFaceModel("Register")
            CaptureImage.Show()
        End If
    End Sub

    Private Sub BunifuCustomDataGrid1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles BunifuCustomDataGrid1.CellContentClick
        If e.RowIndex >= 0 Then
            Dim dtFill As DataTable = SQLConnection.executeQuery("select * from student where id='" & dt.Rows.Item(e.RowIndex).Item(1).ToString() & "'")
            showFullNameTB.Text = dtFill.Rows().Item(0).Item(1).ToString()
            showIDNumber.Text = dtFill.Rows().Item(0).Item(0).ToString()
            showCoursTB.Text = dtFill.Rows().Item(0).Item(2).ToString() & " " & dtFill.Rows().Item(0).Item(3).ToString()
            PictureBox4.Image = Nothing
            Dim nms As New IO.MemoryStream(CType(dtFill.Rows().Item(0).Item(4), Byte()))
            Dim nreturnImage As Image = Image.FromStream(nms)
            PictureBox4.Image = nreturnImage

            Dim qry As String = "select subject.name as 'Subject' , subject.code as 'Code' From student_subjects " &
            "INNER JOIN subject ON subject.code = student_subjects.subject_id WHERE student_subjects.student_id ='" & dt.Rows.Item(e.RowIndex).Item(1).ToString() & "'"

            BunifuCustomDataGrid6.DataSource = SQLConnection.executeQuery(qry)
        End If
    End Sub

    Private Sub BunifuImageButton1_Click_1(sender As Object, e As EventArgs)
        BunifuTransition1.HideSync(ListStudentPanel)
    End Sub
    Private Sub BunifuThinButton29_Click(sender As Object, e As EventArgs) Handles BunifuThinButton29.Click
        BunifuTransition1.HideSync(editSubjectsPanel, False, Animation.HorizSlide)
    End Sub

    Private Sub BunifuThinButton26_Click(sender As Object, e As EventArgs) Handles BunifuThinButton26.Click
        BunifuTransition1.ShowSync(editSubjectsPanel, False, Animation.HorizSlide)
    End Sub

    Private Sub BunifuThinButton25_Click(sender As Object, e As EventArgs) Handles BunifuThinButton25.Click
        CaptureImage.initFaceModel("List")
        CaptureImage.id = editIDNumTB.Text
        CaptureImage.Show()
    End Sub

    Private Sub BunifuImageButton4_Click(sender As Object, e As EventArgs)
        ListStudentPanel.Show()
        BunifuTransition1.HideSync(editStudentPanel)
    End Sub
    Private Sub BunifuThinButton230_Click(sender As Object, e As EventArgs) Handles BunifuThinButton230.Click
        Dim q As String = "Update student set Name='" & editFullNameTB.Text & "' where ID='" & editIDNumTB.Text & "'"
        SQLConnection.executeCommand(q)
        saveNewEditStudentImage()
        BackgroundWorker1.RunWorkerAsync()
        For i As Integer = 0 To ListView2.Items.Count - 1
            saveTrainingImage(ListView2.Items(i).Text, imagesToSave.Images(i))
        Next
        FTPControl.uploadAllFiles()
    End Sub
    Public Sub RefreshRegisterImage()
        imagesToSave = CaptureImage.imagesToSave
        ListView1.LargeImageList = CaptureImage.images
        ListView1.Items.AddRange(CaptureImage.imageNames.ToArray)
    End Sub

    Public Sub RefreshListImage()
        imagesToSave = CaptureImage.imagesToSave
        ListView2.LargeImageList = CaptureImage.images
        ListView2.Items.AddRange(CaptureImage.imageNames.ToArray)
    End Sub
    Private Sub BunifuThinButton23_Click(sender As Object, e As EventArgs) Handles BunifuThinButton23.Click
        If showFullNameTB.Text = "" Then
            MsgBox("Please Select Student First!")
            Return
        End If
        editStudentPanel.Show()
        BunifuTransition2.HideSync(ListStudentPanel)
        editFullNameTB.Text = showFullNameTB.Text
        editIDNumTB.Text = showIDNumber.Text
        editProfilePB.Image = PictureBox4.Image
        'editFrontPB.ImageLocation = Directory.GetCurrentDirectory() & "\TrainingImages\" & showIDNumber.Text & "_Front.jpg"
        'editLeftPB.ImageLocation = Directory.GetCurrentDirectory() & "\TrainingImages\" & showIDNumber.Text & "_Left.jpg"
        'editRightPB.ImageLocation = Directory.GetCurrentDirectory() & "\TrainingImages\" & showIDNumber.Text & "_Right.jpg"
        'editTopPB.ImageLocation = Directory.GetCurrentDirectory() & "\TrainingImages\" & showIDNumber.Text & "_Top.jpg"
        'editBottomPB.ImageLocation = Directory.GetCurrentDirectory() & "\TrainingImages\" & showIDNumber.Text & "_Bottom.jpg"

        RefreshListImage()

        Dim query As String = "select subject.name as 'Subject' ,student_subjects.subject_id as 'Code', TIME_FORMAT(subject.time_in,'%I:%i %p') as 'Time in', TIME_FORMAT(subject.time_out,'%I:%i %p') as 'Time out' From student_subjects " &
            "INNER JOIN subject ON subject.code = student_subjects.subject_id WHERE student_subjects.student_id ='" & editIDNumTB.Text & "'"
        listofStudentSubjectDGV.DataSource = SQLConnection.executeQuery(query)
        Dim q As String = "SELECT s.name,t.code,TIME_FORMAT(s.time_in, '%I:%i %p') as Time_In ,TIME_FORMAT(s.time_out, '%I:%i %p') as Time_Out " &
        "FROM(SELECT subject_id AS code FROM student_subjects where student_id='" & editIDNumTB.Text & "' " &
        "union all SELECT code FROM subject ) t " &
        "left join subject s on s.code = t.code"
        If Not loginName = "" Then
            q &= " where s.faculty='" & loginName & "' "
        End If
        q &= " group by t.code having count(*) = 1"
        dtsubject = SQLConnection.executeQuery(q)
        assignStudentSubjectDGV.DataSource = dtsubject
    End Sub

    Private Sub BunifuThinButton28_Click(sender As Object, e As EventArgs) Handles BunifuThinButton28.Click
        Dim query As String = "Insert into student_subjects (student_id , subject_id) values ('" & editIDNumTB.Text & "','" & assignStudentSubjectDGV.SelectedRows.Item(0).Cells.Item(1).Value.ToString & "')"
        SQLConnection.executeCommand(query)

        query = "select subject.name as 'Subject' ,student_subjects.subject_id as 'Code', TIME_FORMAT(subject.time_in,'%I:%i %p') as 'Time in', TIME_FORMAT(subject.time_out,'%I:%i %p') as 'Time out' From student_subjects " &
            "INNER JOIN subject ON subject.code = student_subjects.subject_id WHERE student_subjects.student_id ='" & editIDNumTB.Text & "'"

        listofStudentSubjectDGV.DataSource = SQLConnection.executeQuery(query)
        MsgBox("Succesfully Added to your Subjects")


        Dim q As String = "SELECT s.name,t.code,TIME_FORMAT(s.time_in, '%I:%i %p') as Time_In ,TIME_FORMAT(s.time_out, '%I:%i %p') as Time_Out " &
        "FROM(SELECT subject_id AS code FROM student_subjects where student_id='" & editIDNumTB.Text & "' " &
        "union all SELECT code FROM subject ) t " &
        "left join subject s on s.code = t.code"

        If Not loginName = "" Then
            q &= " where s.faculty='" & loginName & "' "
        End If

        q &= " group by t.code having count(*) = 1"

        dtsubject = SQLConnection.executeQuery(q)
        assignStudentSubjectDGV.DataSource = dtsubject
    End Sub

    Private Sub BunifuThinButton27_Click(sender As Object, e As EventArgs) Handles BunifuThinButton27.Click
        Dim selectedSubject As String = listofStudentSubjectDGV.SelectedRows.Item(0).Cells.Item(1).Value.ToString
        Dim query As String = "delete from student_subjects where student_id='" & editIDNumTB.Text & "' and subject_id = '" & selectedSubject & "'"
        SQLConnection.executeCommand(query)

        query = "select subject.name as 'Subject' ,student_subjects.subject_id as 'Code', TIME_FORMAT(subject.time_in,'%I:%i %p') as 'Time in', TIME_FORMAT(subject.time_out,'%I:%i %p') as 'Time out' From student_subjects " &
            "INNER JOIN subject ON subject.code = student_subjects.subject_id WHERE student_subjects.student_id ='" & editIDNumTB.Text & "'"

        listofStudentSubjectDGV.DataSource = SQLConnection.executeQuery(query)

        Dim q As String = "SELECT s.name,t.code,TIME_FORMAT(s.time_in, '%I:%i %p') as Time_In ,TIME_FORMAT(s.time_out, '%I:%i %p') as Time_Out " &
        "FROM(SELECT subject_id AS code FROM student_subjects where student_id='" & editIDNumTB.Text & "' " &
        "union all SELECT code FROM subject ) t " &
        "left join subject s on s.code = t.code"

        If Not loginName = "" Then
            q &= " where s.faculty='" & loginName & "' "
        End If

        q &= " group by t.code having count(*) = 1"

        dtsubject = SQLConnection.executeQuery(q)
        assignStudentSubjectDGV.DataSource = dtsubject

        MsgBox("Subject Successfully Deleted")
    End Sub

    Private Sub registerStudentToolStripMenuItem_click(sender As Object, e As EventArgs) Handles REGISTERSTUDENTToolStripMenuItem.Click
        HideAllPanel()
        registerStudentPanel.Show()
    End Sub
    Private Sub LISTSTUDENTToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LISTSTUDENTToolStripMenuItem.Click
        Dim dtx As DataTable = SQLConnection.executeQuery("select Name as 'Full Name' , ID as 'ID Number' , CONCAT(Course,' ' ,Year) as 'Course and Year' from student")
        BunifuCustomDataGrid1.DataSource = dtx
        HideAllPanel()
        ListStudentPanel.Show()
    End Sub

    Private Sub EXITToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Private Sub CAPTUREToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CAPTUREToolStripMenuItem.Click
        FaceRecog.Show()
    End Sub

    Private Sub BunifuThinButton210_Click(sender As Object, e As EventArgs) Handles BunifuThinButton210.Click
        defaultAddStudent()
    End Sub

    Private Sub BunifuThinButton211_Click(sender As Object, e As EventArgs) Handles BunifuThinButton211.Click
        HideAllPanel()
        ListStudentPanel.Show()
    End Sub

    Private Sub clearAddSubject()
        BunifuMaterialTextbox1.Text = ""
        BunifuMaterialTextbox2.Text = ""
        BunifuMaterialTextbox16.Text = ""

        CheckBox6.Checked = False
        CheckBox7.Checked = False
        CheckBox8.Checked = False
        CheckBox9.Checked = False
        CheckBox9.Checked = False
        CheckBox10.Checked = False
        CheckBox11.Checked = False
    End Sub

    Private Sub BunifuThinButton212_Click(sender As Object, e As EventArgs) Handles BunifuThinButton212.Click
        clearAddSubject()
    End Sub

    Private Function checkIfOverlap(_code As String, _room As String, _timeIn As DateTime, _timeOut As DateTime, _week As List(Of Boolean))
        Dim overlap As Boolean = False

        Dim _weekWord As New List(Of String)
        _weekWord.Add("Mon")
        _weekWord.Add("Tue")
        _weekWord.Add("Wed")
        _weekWord.Add("Thu")
        _weekWord.Add("Fri")
        _weekWord.Add("Sat")
        _weekWord.Add("Sun")

        Dim _index As Integer = -1
        For Each _day As Boolean In _week
            _index += 1
            If Not _day Then
                Continue For
            Else
                Dim overlapDT As DataTable = SQLConnection.executeQuery("select time_in,time_out,name from subject where Room='" & _room & "' and " &
                                                                        _weekWord(_index) & "=1")
                For Each items As DataRow In overlapDT.Rows
                    Dim tIn As DateTime = Convert.ToDateTime(items.Item(0).ToString())
                    Dim tOut As DateTime = Convert.ToDateTime(items.Item(1).ToString())
                    If (_timeIn.TimeOfDay >= tIn.TimeOfDay And _timeIn.TimeOfDay <= tOut.TimeOfDay) Or (_timeOut.TimeOfDay >= tIn.TimeOfDay And _timeOut.TimeOfDay <= tOut.TimeOfDay) Then
                        overlap = True
                        MsgBox("The Subject you are creating has a conflict at Room " & _room & ",Subject " & items.Item(2).ToString() &
                               ",where time in is " & items.Item(0).ToString() & " and time out is " & items.Item(1).ToString() &
                               "on " & _weekWord(_index))
                    End If
                Next
            End If
        Next
        Return overlap
    End Function
    Private Sub BunifuThinButton213_Click(sender As Object, e As EventArgs) Handles BunifuThinButton213.Click
        Dim _week As New List(Of Boolean)
        _week.Add(CheckBox6.Checked)
        _week.Add(CheckBox7.Checked)
        _week.Add(CheckBox8.Checked)
        _week.Add(CheckBox9.Checked)
        _week.Add(CheckBox10.Checked)
        _week.Add(CheckBox11.Checked)
        _week.Add(CheckBox12.Checked)


        If checkIfOverlap(BunifuMaterialTextbox2.Text, BunifuMaterialTextbox16.Text, Convert.ToDateTime(DateTimePicker3.Value.ToString("HH:mm:ss")), Convert.ToDateTime(DateTimePicker4.Value.ToString("HH:mm:ss")), _week) Then
            Return
        End If
        If editSubjectState Then
            Dim q As String = "Update subject set name='" & BunifuMaterialTextbox1.Text & "',time_in='" &
                    DateTimePicker3.Value.ToString("HH:mm:ss") & "',time_out='" & DateTimePicker4.Value.ToString("HH:mm:ss") & "',Mon='" & Convert.ToInt32(CheckBox6.Checked) &
                    "',Tue='" & Convert.ToInt32(CheckBox7.Checked) & "',Wed='" & Convert.ToInt32(CheckBox8.Checked) & "',Thu='" &
                    Convert.ToInt32(CheckBox9.Checked) & "',Fri='" & Convert.ToInt32(CheckBox10.Checked) & "',Sat='" &
                    Convert.ToInt32(CheckBox11.Checked) & "',Sun='" & Convert.ToInt32(CheckBox12.Checked) & "',Room='" &
                    BunifuMaterialTextbox16.Text & "' where code='" & BunifuMaterialTextbox2.Text & "'"
            SQLConnection.executeCommand(q)
            editSubjectState = False

            BunifuThinButton232.Visible = False
            Label22.Text = "Add new Subject"

            Dim dtNews As DataTable = SQLConnection.executeQuery("select name,code,TIME_FORMAT(time_in, '%I:%i %p') as Time_In ,TIME_FORMAT(time_out, '%I:%i %p') as Time_Out,Room  from subject")
            BunifuCustomDataGrid5.DataSource = dtNews

            MsgBox("Edited Successfully!!")

            BunifuMaterialTextbox1.Enabled = True
            Return
        End If
        Dim schedString As String = ""
        Dim schedStatus As String = ""
        If CheckBox6.Checked Then
            schedString += ",Mon"
            schedStatus += ",'1'"
        End If
        If CheckBox7.Checked Then
            schedString += ",Tue"
            schedStatus += ",'1'"
        End If
        If CheckBox8.Checked Then
            schedString += ",Wed"
            schedStatus += ",'1'"
        End If
        If CheckBox9.Checked Then
            schedString += ",Thu"
            schedStatus += ",'1'"
        End If
        If CheckBox10.Checked Then
            schedString += ",Fri"
            schedStatus += ",'1'"
        End If

        If CheckBox11.Checked Then
            schedString += ",Sat"
            schedStatus += ",'1'"
        End If

        If CheckBox12.Checked Then
            schedString += ",Sun"
            schedStatus += ",'1'"
        End If
        If BunifuMaterialTextbox1.Text = "" Or BunifuMaterialTextbox2.Text = "" Then
            MsgBox("Please fill in all Fields!")
            Return
        End If
        Dim query As String = "Insert into subject (name,code,time_in,time_out" & schedString & ",Room) values ('" & BunifuMaterialTextbox1.Text &
            "','" & BunifuMaterialTextbox2.Text & "','" & DateTimePicker3.Value.ToString("HH:mm:ss") & "','" &
            DateTimePicker4.Value.ToString("HH:mm:ss") & "'" & schedStatus & ",'" & BunifuMaterialTextbox16.Text & "')"
        Try
            SQLConnection.executeCommand(query)
        Catch ex As Exception
            If ex.ToString().Contains("Duplicate") Then
                MsgBox("Duplicate!")
            End If
            Return
        End Try
        BunifuMaterialTextbox1.Text = ""
        BunifuMaterialTextbox2.Text = ""
        BunifuMaterialTextbox16.Text = ""

        CheckBox6.Checked = False
        CheckBox7.Checked = False
        CheckBox8.Checked = False
        CheckBox9.Checked = False
        CheckBox10.Checked = False
        CheckBox11.Checked = False
        CheckBox12.Checked = False

        MsgBox("Subject Successfully Saved!")
        Dim dtNew As DataTable = SQLConnection.executeQuery("Select name,code,TIME_FORMAT(time_in, '%I:%i %p') as Time_In ,TIME_FORMAT(time_out, '%I:%i %p') as Time_Out, Room  from subject")
        BunifuCustomDataGrid5.DataSource = dtNew
    End Sub

    Private Sub ADDSUBJECTToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ADDSUBJECTToolStripMenuItem.Click

        Label22.Text = "Add new Subject"
        BunifuThinButton232.Visible = False
        editSubjectState = False
        clearAddSubject()
        BunifuMaterialTextbox1.Enabled = True
        HideAllPanel()
        AddFacultyPnl.Show()
    End Sub

    Public Sub hideAll()

    End Sub

    Private Sub BunifuThinButton217_Click(sender As Object, e As EventArgs) Handles BunifuThinButton217.Click
        Dim fiF As FilterInfoCollection = New FilterInfoCollection(FilterCategory.VideoInputDevice)
        videocapture = New VideoCaptureDevice(fiF(0).MonikerString)
        AddHandler videocapture.NewFrame, New NewFrameEventHandler(AddressOf camFrame)
        videocapture.Start()
    End Sub

    Public Sub camFrame(sender As Object, e As NewFrameEventArgs)
        bmp = DirectCast(e.Frame.Clone(), Bitmap)
        bmp.RotateFlip(RotateFlipType.Rotate180FlipY)
        PictureBox6.Image = bmp
    End Sub

    Private Sub BunifuThinButton216_Click(sender As Object, e As EventArgs) Handles BunifuThinButton216.Click
        If IsNothing(videocapture) Then
            Return
        End If
        If videocapture.IsRunning Then
            videocapture.Stop()

        End If
    End Sub

    Private Sub BunifuThinButton215_Click(sender As Object, e As EventArgs) Handles BunifuThinButton215.Click
        If checkDuplicateID(BunifuMaterialTextbox7.Text) Then
            MsgBox("ID has Duplicate!")
            Return
        End If

        If checkDuplicateUserName(BunifuMaterialTextbox5.Text) Then
            MsgBox("Username has Duplicate!")
            Return
        End If

        If BunifuMaterialTextbox5.Text = "" Or BunifuMaterialTextbox6.Text = "" Or BunifuMaterialTextbox7.Text = "" Then
            MsgBox("Please fill credentials first!")
        End If
        Dim query As String = "Insert into staff (Name , username , password , ID , Type) values ('" & BunifuMaterialTextbox4.Text &
            "','" & BunifuMaterialTextbox5.Text & "','" & BunifuMaterialTextbox6.Text & "','" & BunifuMaterialTextbox7.Text & "', 'Faculty')"
        SQLConnection.executeCommand(query)
        saveFacultyImage()
        clearAddFaculty()
        'dtStaff = SQLConnection.executeQuery("select name,id,username,password from staff where type = 'faculty'")
        'BunifuCustomDataGrid2.DataSource = dtStaff
    End Sub

    Private Sub saveFacultyImage()
        If PictureBox6.Image Is Nothing Then
            MsgBox("Capture your Profile Photo First")
            Return
        End If
        Dim newImage As Image = PictureBox6.Image

        Dim picStream As New System.IO.MemoryStream()
        newImage.Save(picStream, System.Drawing.Imaging.ImageFormat.Png)
        Dim imgArray() As Byte = picStream.GetBuffer()
        picStream.Close()

        Dim photoQuery As String = "Update staff Set Image=@photo where username = '" & BunifuMaterialTextbox5.Text & "'"
        SQLConnection.executePhotoCommand(photoQuery, imgArray)

        MsgBox("Faculty Saved Successfully!")
    End Sub

    Private Sub clearAddFaculty()
        BunifuMaterialTextbox4.Text = ""
        BunifuMaterialTextbox5.Text = ""
        BunifuMaterialTextbox6.Text = ""
        BunifuMaterialTextbox7.Text = ""

        PictureBox6.Image = Nothing
    End Sub

    Private Sub BunifuThinButton214_Click(sender As Object, e As EventArgs) Handles BunifuThinButton214.Click
        clearAddFaculty()
    End Sub

    Private Sub BunifuCustomDataGrid2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles BunifuCustomDataGrid2.CellContentClick
        If e.RowIndex >= 0 Then
            Dim dtFill As DataTable = SQLConnection.executeQuery("select * from staff where username='" & dtStaff.Rows.Item(e.RowIndex).Item(2).ToString() & "'")
            BunifuMaterialTextbox9.Text = dtFill.Rows().Item(0).Item(1).ToString()
            BunifuMaterialTextbox8.Text = dtFill.Rows().Item(0).Item(3).ToString()
            BunifuMaterialTextbox10.Text = dtFill.Rows().Item(0).Item(0).ToString()
            editPW = dtFill.Rows().Item(0).Item(4).ToString()
            If Not IsDBNull(dtFill.Rows().Item(0).Item(5)) Then

                Dim nms As New IO.MemoryStream(CType(dtFill.Rows().Item(0).Item(5), Byte()))
                Dim nreturnImage As Image = Image.FromStream(nms)
                If Not IsNothing(nreturnImage) Then
                    PictureBox7.Image = Nothing
                    PictureBox7.Image = nreturnImage
                End If

                BunifuCustomDataGrid8.DataSource = SQLConnection.executeQuery("select name,code from subject where faculty = '" & dtStaff.Rows.Item(e.RowIndex).Item(2).ToString() & "'")

            End If
        End If
    End Sub

    Private Sub BunifuThinButton218_Click(sender As Object, e As EventArgs) Handles BunifuThinButton218.Click
        If BunifuMaterialTextbox9.Text = "" Then
            MsgBox("Please select Faculty!")
            Return
        End If
        HideAllPanel()
        PictureBox8.Image = PictureBox7.Image
        BunifuMaterialTextbox14.Text = BunifuMaterialTextbox9.Text
        BunifuMaterialTextbox13.Text = BunifuMaterialTextbox8.Text
        BunifuMaterialTextbox11.Text = BunifuMaterialTextbox10.Text
        BunifuMaterialTextbox12.Text = editPW
        EditFacultyDetailsPanel.Show()

        BunifuMaterialTextbox9.Text = ""
        BunifuMaterialTextbox8.Text = ""
        BunifuMaterialTextbox10.Text = ""
        PictureBox7.Image = Nothing
    End Sub

    Private Sub BunifuThinButton219_Click(sender As Object, e As EventArgs)
        BunifuTransition1.HideSync(Panel2, False, Animation.HorizSlide)
    End Sub

    Private Sub BunifuThinButton221_Click(sender As Object, e As EventArgs) Handles BunifuThinButton221.Click
        Dim query As String = "Update subject set faculty = null where code='" & BunifuCustomDataGrid3.SelectedRows.Item(0).Cells.Item(1).Value.ToString() & "'"
        SQLConnection.executeCommand(query)

        MsgBox("Remove Successfull!")

        Dim dtSubs As DataTable = SQLConnection.executeQuery("select name as 'Subject' , code as 'Code' from subject where faculty is null")
        BunifuCustomDataGrid4.DataSource = dtSubs

        dtFaculty = SQLConnection.executeQuery("select name as 'Subject' , code as 'Code' from subject where faculty='" & FacultyAssignList.Item(ComboBox4.SelectedIndex).Item(3) & "'")
        BunifuCustomDataGrid3.DataSource = dtFaculty
    End Sub

    Private Sub BunifuCustomDataGrid3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles BunifuCustomDataGrid3.CellContentClick
        'If (e.ColumnIndex = -1) Then
        '    Return
        'End If
        Dim dtFill As DataTable = SQLConnection.executeQuery("select * from subject where code='" & dtFaculty.Rows.Item(e.RowIndex).Item(1).ToString() & "'")
    End Sub

    Private Sub BunifuThinButton220_Click(sender As Object, e As EventArgs)
    End Sub

    Private Sub BunifuThinButton222_Click(sender As Object, e As EventArgs) Handles BunifuThinButton222.Click
        If ComboBox4.SelectedIndex = -1 Then
            MsgBox("Please Select Faculty!")
            Return
        End If
        Dim query As String = "Update subject set faculty ='" & FacultyAssignList.Item(ComboBox4.SelectedIndex).Item(3).ToString() &
            "' where code='" & BunifuCustomDataGrid4.SelectedRows.Item(0).Cells.Item(1).Value.ToString() & "'"
        SQLConnection.executeQuery(query)
        MsgBox("Successfully Added!")

        dtFaculty = SQLConnection.executeQuery("select name as 'Subject' , code as 'Code' from subject where faculty='" & FacultyAssignList.Item(ComboBox4.SelectedIndex).Item(3) & "'")
        BunifuCustomDataGrid3.DataSource = dtFaculty

        Dim dtSubs As DataTable = SQLConnection.executeQuery("select name as 'Subject' , code as 'Code' from subject where faculty is null")
        BunifuCustomDataGrid4.DataSource = dtSubs

    End Sub

    Private Sub HideAllPanel()
        EditFacultyDetailsPanel.Hide()
        AssignSubjectFacultyPanel.Hide()
        ListOfFacultyPanel.Hide()
        AddFacultyPnl.Hide()
        AddNewFacultyPanel.Hide()
        editStudentPanel.Hide()
        ListStudentPanel.Hide()
        registerStudentPanel.Hide()
        FreePanel.Hide()
    End Sub

    Private Sub ASSIGNSUBJECTToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ASSIGNSUBJECTToolStripMenuItem.Click
        HideAllPanel()
        AssignSubjectFacultyPanel.Show()
        ComboBox4.Text = ""
        ComboBox4.Items.Clear()
        BunifuCustomDataGrid3.DataSource = Nothing
        BunifuCustomDataGrid4.DataSource = Nothing
        Dim dtFacultyList As DataTable = SQLConnection.executeQuery("Select * from staff where Type='Faculty'")
        For i As Integer = 0 To dtFacultyList.Rows.Count - 1
            ComboBox4.Items.Add(dtFacultyList.Rows.Item(i).Item(1).ToString())
        Next
        FacultyAssignList = dtFacultyList.Rows

        Dim dtSubs As DataTable = SQLConnection.executeQuery("select name as 'Subject' , code as 'Code' from subject where faculty is null")
        BunifuCustomDataGrid4.DataSource = dtSubs
    End Sub

    Private Sub LISTFACULTYToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LISTFACULTYToolStripMenuItem.Click
        HideAllPanel()
        dtStaff = SQLConnection.executeQuery("select num as 'No.', Name as 'Full Name' , username as 'Username' from staff where Type='Faculty'")

        BunifuCustomDataGrid2.Rows.Clear()
        Dim dgvID As Integer = 0
        For i As Integer = 0 To dtStaff.Rows.Count - 1
            dgvID = BunifuCustomDataGrid2.Rows.Add()
            BunifuCustomDataGrid2.Rows(dgvID).Cells.Item(0).Value = i + 1
            BunifuCustomDataGrid2.Rows(dgvID).Cells.Item(1).Value = dtStaff.Rows.Item(i).Item(1).ToString()
            BunifuCustomDataGrid2.Rows(dgvID).Cells.Item(2).Value = dtStaff.Rows.Item(i).Item(2).ToString()
        Next
        ListOfFacultyPanel.Show()
    End Sub

    Private Sub ADDFACULTYToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ADDFACULTYToolStripMenuItem.Click
        HideAllPanel()
        AddNewFacultyPanel.Show()
    End Sub

    Private Sub REPORTSToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles REPORTSToolStripMenuItem.Click
        ComboBox1.Items.Clear()
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        Panel3.Show()

        If loginName = "" Then
            Dim query As String = "Select * from staff where Type='Faculty'"
            Dim dt As DataTable = SQLConnection.executeQuery(query)
            FacultyList = dt.Rows

            For i As Integer = 0 To dt.Rows.Count - 1
                ComboBox1.Items.Add(dt.Rows.Item(i).Item(1).ToString)
            Next
        Else
            Dim query As String = "Select * from staff where username='" & loginName & "'"
            Dim dt As DataTable = SQLConnection.executeQuery(query)
            FacultyList = dt.Rows
            For i As Integer = 0 To dt.Rows.Count - 1
                ComboBox1.Items.Add(dt.Rows.Item(i).Item(1).ToString)
            Next
            ComboBox1.SelectedIndex = 0


            ComboBox2.Items.Clear()
            ComboBox2.Text = ""
            query = "Select * from subject where faculty='" & loginName & "'"
            _dtForSubject = SQLConnection.executeQuery(query)
            SubjectList = _dtForSubject.Rows
            cbBoxContents = New List(Of personDetail)

            For i As Integer = 0 To _dtForSubject.Rows.Count - 1
                cbBoxContents.Add(New personDetail(_dtForSubject.Rows.Item(i).Item(0).ToString, _dtForSubject.Rows.Item(i).Item(1).ToString))
                ComboBox2.Items.Add(_dtForSubject.Rows.Item(i).Item(0).ToString)
            Next
        End If
        HideAllPanel()
        FreePanel.Show()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox2.Items.Clear()
        ComboBox2.Text = ""
        Dim query As String = "Select * from subject where faculty='" & FacultyList.Item(ComboBox1.SelectedIndex).Item(3).ToString & "'"
        _dtForSubject = SQLConnection.executeQuery(query)
        SubjectList = _dtForSubject.Rows
        cbBoxContents = New List(Of personDetail)

        For i As Integer = 0 To _dtForSubject.Rows.Count - 1
            cbBoxContents.Add(New personDetail(_dtForSubject.Rows.Item(i).Item(0).ToString, _dtForSubject.Rows.Item(i).Item(1).ToString))
            ComboBox2.Items.Add(_dtForSubject.Rows.Item(i).Item(0).ToString)
        Next

    End Sub

    Private Sub BunifuThinButton223_Click(sender As Object, e As EventArgs) Handles BunifuThinButton223.Click
        Panel3.Hide()
        BunifuThinButton225.Visible = True
        BunifuThinButton224.Visible = True

        BunifuThinButton233.Visible = True
        BunifuCustomDataGrid7.Show()
        Label65.Visible = True
        Label65.Text = "Subject: " & ComboBox2.SelectedItem.ToString()
        BunifuCustomDataGrid7.Columns.Clear()
        BunifuCustomDataGrid7.Rows.Clear()

        BunifuCustomDataGrid7.Columns.Add("ID", "ID Number")
        BunifuCustomDataGrid7.Columns.Add("Name", "Name")
        Dim reportsDT As DataTable = SQLConnection.executeQuery("Select dateSet from attendanceschedule where subjectCode = '" & _dtForSubject.Rows.Item(ComboBox2.SelectedIndex).Item(1).ToString() & "'")

        Dim lastDate As String = ""
        For Each items As DataRow In reportsDT.Rows
            Dim d As String = Convert.ToDateTime(items.Item(0).ToString()).ToString("MMMM dd")
            If Not lastDate = d Then
                BunifuCustomDataGrid7.Columns.Add(d, d)
            End If
            lastDate = d
        Next

        Dim idList As New List(Of personDetail)
        reportsDT = SQLConnection.executeQuery("select sub.student_id , stud.name from student_subjects sub join student stud on sub.student_id = stud.ID " &
             "where sub.subject_id = '" & _dtForSubject.Rows.Item(ComboBox2.SelectedIndex).Item(1).ToString() & "'")
        For Each items As DataRow In reportsDT.Rows
            Dim studentDetail As New personDetail(items.Item(1).ToString(), items.Item(0).ToString())
            idList.Add(studentDetail)
        Next

        reportsDT = SQLConnection.executeQuery("Select time_in from subject where code = '" & cbBoxContents(ComboBox2.SelectedIndex).id & "'")
        Dim TimeOutCurrent As DateTime = Convert.ToDateTime(reportsDT.Rows.Item(0).Item(0).ToString())

        For Each item As personDetail In idList
            Dim rowID As Integer = BunifuCustomDataGrid7.Rows.Add()
            BunifuCustomDataGrid7.Rows.Item(rowID).Cells.Item(0).Value = item.id
            BunifuCustomDataGrid7.Rows.Item(rowID).Cells.Item(1).Value = item.name
            reportsDT = SQLConnection.executeQuery("Select * from attendance where Student_Code = '" & item.id & "' and Student_Subject='" & _dtForSubject.Rows.Item(ComboBox2.SelectedIndex).Item(1).ToString() & "'")
            For Each items As DataRow In reportsDT.Rows
                Dim d As String = Convert.ToDateTime(items.Item(2).ToString()).ToString("MMMM dd")
                If TimeOutCurrent.AddMinutes(15).TimeOfDay < Convert.ToDateTime(reportsDT.Rows.Item(0).Item(2).ToString).TimeOfDay Then
                    BunifuCustomDataGrid7.Rows.Item(rowID).Cells.Item(d).Value = "T"
                Else
                    BunifuCustomDataGrid7.Rows.Item(rowID).Cells.Item(d).Value = "P"
                End If
            Next
            For Each items As DataGridViewCell In BunifuCustomDataGrid7.Rows.Item(rowID).Cells
                If items.Value = Nothing Or items.Value = "" Then
                    items.Value = "A"
                End If
            Next
        Next
    End Sub

    Private Sub BunifuThinButton225_Click(sender As Object, e As EventArgs) Handles BunifuThinButton225.Click
        BunifuCustomDataGrid7.Rows.Clear()
        BunifuCustomDataGrid7.Visible = False
        BunifuThinButton225.Visible = False
        BunifuThinButton224.Visible = False
        BunifuThinButton233.Visible = False
        Label65.Visible = False
        Panel3.Show()
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged

        dtFaculty = SQLConnection.executeQuery("select name as 'Subject' , code as 'Code' from subject where faculty='" & FacultyAssignList.Item(ComboBox4.SelectedIndex).Item(3) & "'")
        BunifuCustomDataGrid3.DataSource = dtFaculty
    End Sub

    Private Sub LOGOUTToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LOGOUTToolStripMenuItem.Click
        If MessageBox.Show("Are you sure you want to Logout?", "Logout?", MessageBoxButtons.OKCancel) = DialogResult.OK Then
            Me.Close()
            LoginForm.Show()
            LogoutState = True
        End If
    End Sub

    Private Sub BunifuThinButton219_Click_1(sender As Object, e As EventArgs) Handles BunifuThinButton219.Click
        If BunifuCustomDataGrid5.SelectedRows.Count < 1 Then
            Return
        End If
        Dim query As String = "Delete from subject where code='" & BunifuCustomDataGrid5.SelectedRows.Item(0).Cells.Item(1).Value.ToString() & "'"
        SQLConnection.executeCommand(query)
        query = "Delete from student_subjects where subject_id='" & BunifuCustomDataGrid5.SelectedRows.Item(0).Cells.Item(1).Value.ToString() & "'"
        SQLConnection.executeCommand(query)
        MsgBox("Successfully Removed!")
        Dim dtNew As DataTable = SQLConnection.executeQuery("select name,code,time_in,time_out,Room from subject")
        BunifuCustomDataGrid5.DataSource = dtNew
    End Sub
    Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If LogoutState Then
            LogoutState = False
            closeForm = True
            If BackgroundWorker1.IsBusy Then
                BackgroundWorker1.CancelAsync()
            End If
        Else
            NotifyIcon1.Visible = True
            Me.Hide()
        End If
    End Sub

    Private Sub BunifuThinButton24_Click(sender As Object, e As EventArgs) Handles BunifuThinButton24.Click
        Dim fiF As FilterInfoCollection = New FilterInfoCollection(FilterCategory.VideoInputDevice)
        videocapture = New VideoCaptureDevice(fiF(0).MonikerString)
        AddHandler videocapture.NewFrame, New NewFrameEventHandler(AddressOf camFrame2)
        videocapture.Start()
    End Sub

    Private Sub BunifuThinButton220_Click_1(sender As Object, e As EventArgs) Handles BunifuThinButton220.Click
        If IsNothing(videocapture) Then
            Return
        End If
        If videocapture.IsRunning Then
            videocapture.Stop()
        End If
    End Sub

    Public Sub camFrame2(sender As Object, e As NewFrameEventArgs)
        bmp = DirectCast(e.Frame.Clone(), Bitmap)
        bmp.RotateFlip(RotateFlipType.Rotate180FlipY)
        PictureBox8.Image = bmp
    End Sub

    Private Function checkDuplicateID(id As String)
        Dim dup As Boolean = False
        Dim idDT As DataTable = SQLConnection.executeQuery("Select ID from staff")
        For Each items As DataRow In idDT.Rows
            If id = items.Item(0).ToString() Then
                dup = True
            End If
        Next
        Return dup
    End Function

    Private Function checkDuplicateUserName(_username As String)
        Dim dup As Boolean = False
        Dim idDT As DataTable = SQLConnection.executeQuery("Select username from staff")
        For Each items As DataRow In idDT.Rows
            If _username = items.Item(0).ToString() Then
                dup = True
            End If
        Next
        Return dup
    End Function

    Private Sub BunifuThinButton227_Click(sender As Object, e As EventArgs) Handles BunifuThinButton227.Click
        Dim query As String = "Update staff set Name='" & BunifuMaterialTextbox14.Text & "' , username='" & BunifuMaterialTextbox13.Text &
            "', password='" & BunifuMaterialTextbox12.Text & "', ID='" & BunifuMaterialTextbox11.Text & "' where ID='" & BunifuMaterialTextbox11.Text & "'"
        SQLConnection.executeCommand(query)

        dtStaff = SQLConnection.executeQuery("select num as 'No.', Name as 'Full Name' , username as 'Username' from staff where Type='Faculty'")
        SQLConnection.executeCommand("update subject set faculty=Null where faculty='" & BunifuMaterialTextbox13.Text & "'")
        BunifuCustomDataGrid2.Rows.Clear()
        Dim dgvID As Integer = 0
        For i As Integer = 0 To dtStaff.Rows.Count - 1
            dgvID = BunifuCustomDataGrid2.Rows.Add()
            BunifuCustomDataGrid2.Rows(dgvID).Cells.Item(0).Value = i + 1
            BunifuCustomDataGrid2.Rows(dgvID).Cells.Item(1).Value = dtStaff.Rows.Item(i).Item(1).ToString()
            BunifuCustomDataGrid2.Rows(dgvID).Cells.Item(2).Value = dtStaff.Rows.Item(i).Item(2).ToString()
        Next

        saveNewFacultyImage()
    End Sub

    Private Sub saveNewFacultyImage()
        If PictureBox8.Image Is Nothing Then
            MsgBox("Capture your Profile Photo First")
            Return
        End If
        Dim newImage As Image = PictureBox8.Image

        Dim picStream As New System.IO.MemoryStream()
        newImage.Save(picStream, System.Drawing.Imaging.ImageFormat.Png)
        Dim imgArray() As Byte = picStream.GetBuffer()
        picStream.Close()

        Dim photoQuery As String = "Update staff Set Image=@photo where ID = '" & BunifuMaterialTextbox11.Text & "'"
        SQLConnection.executePhotoCommand(photoQuery, imgArray)

        MsgBox("Faculty Edited Successfully!")
    End Sub
    Private Sub BunifuThinButton226_Click(sender As Object, e As EventArgs) Handles BunifuThinButton226.Click
        HideAllPanel()
        ListOfFacultyPanel.Show()
    End Sub
    Private Sub NotifyIcon1_Click(sender As Object, e As EventArgs) Handles NotifyIcon1.Click
        Me.Show()
        NotifyIcon1.Visible = False
    End Sub
    Private Sub BunifuThinButton228_Click(sender As Object, e As EventArgs) Handles BunifuThinButton228.Click
        If BunifuMaterialTextbox9.Text = "" Then
            MsgBox("Select Teacher to delete")
            Return
        End If
        If MessageBox.Show("Are you sure you want to delete this Teacher", "Delete", MessageBoxButtons.OKCancel) = DialogResult.OK Then
            SQLConnection.executeCommand("DELETE from staff where username='" & BunifuCustomDataGrid2.SelectedRows.Item(0).Cells.Item(2).Value.ToString() & "'")
            BunifuCustomDataGrid2.Rows.Clear()
            dtStaff = SQLConnection.executeQuery("select Name as 'Full Name' , username as 'Username' from staff where Type='Faculty'")
            Dim dgvID As Integer = 0

            For i As Integer = 0 To dtStaff.Rows.Count - 1
                dgvID = BunifuCustomDataGrid2.Rows.Add()
                BunifuCustomDataGrid2.Rows(dgvID).Cells.Item(0).Value = i + 1
                BunifuCustomDataGrid2.Rows(dgvID).Cells.Item(1).Value = dtStaff.Rows.Item(0).Item(1).ToString()
                BunifuCustomDataGrid2.Rows(dgvID).Cells.Item(2).Value = dtStaff.Rows.Item(0).Item(2).ToString()
            Next
            MsgBox("Deleted Successfully!")
        End If
    End Sub
    Private Sub BunifuThinButton229_Click(sender As Object, e As EventArgs) Handles BunifuThinButton229.Click
        If BunifuCustomDataGrid5.SelectedRows.Count <= 0 Then
            Return
        End If

        If BunifuCustomDataGrid5.SelectedRows.Item(0).Cells.Item(1).Value.ToString() = "" Then
            Return
        End If

        BunifuThinButton232.Visible = True
        Dim q As String = "Select * from subject where code='" & BunifuCustomDataGrid5.SelectedRows.Item(0).Cells.Item(1).Value.ToString() & "'"
        Dim editSubjectDT As DataTable = SQLConnection.executeQuery(q)
        editSubjectState = True
        BunifuMaterialTextbox1.Text = editSubjectDT.Rows.Item(0).Item(0).ToString()
        BunifuMaterialTextbox2.Text = editSubjectDT.Rows.Item(0).Item(1).ToString()
        CheckBox6.Checked = editSubjectDT.Rows.Item(0).Item(5)
        CheckBox7.Checked = editSubjectDT.Rows.Item(0).Item(6)
        CheckBox8.Checked = editSubjectDT.Rows.Item(0).Item(7)
        CheckBox9.Checked = editSubjectDT.Rows.Item(0).Item(8)
        CheckBox10.Checked = editSubjectDT.Rows.Item(0).Item(9)
        CheckBox11.Checked = editSubjectDT.Rows.Item(0).Item(10)
        CheckBox12.Checked = editSubjectDT.Rows.Item(0).Item(11)
        BunifuMaterialTextbox16.Text = editSubjectDT.Rows.Item(0).Item(12).ToString()
        DateTimePicker3.Value = Convert.ToDateTime(editSubjectDT.Rows.Item(0).Item(2).ToString())
        DateTimePicker4.Value = Convert.ToDateTime(editSubjectDT.Rows.Item(0).Item(3).ToString())
        Label22.Text = "Editing Subject"
        MsgBox("You are now in Edit Mode!")
    End Sub
    Private Sub BunifuMaterialTextbox3_OnValueChanged(sender As Object, e As EventArgs) Handles BunifuMaterialTextbox3.OnValueChanged
        Dim q As String = "SELECT s.name,t.code,TIME_FORMAT(s.time_in, '%I:%i %p') as Time_In ,TIME_FORMAT(s.time_out, '%I:%i %p') as Time_Out " &
        "FROM(SELECT subject_id AS code FROM student_subjects where student_id='" & editIDNumTB.Text & "' " &
        "union all SELECT code FROM subject ) t " &
        "left join subject s on s.code = t.code"
        If Not loginName = "" Then
            q &= " where s.faculty='" & loginName & "' "
        End If
        q &= " group by t.code having count(*) = 1"
        dtsubject = SQLConnection.executeQuery(q)
        assignStudentSubjectDGV.DataSource = dtsubject
    End Sub
    Private Sub saveNewEditStudentImage()
        If editProfilePB.Image Is Nothing Then
            MsgBox("Capture your Profile Photo First")
            Return
        End If
        Dim newImage As Image = editProfilePB.Image
        Dim picStream As New System.IO.MemoryStream()
        newImage.Save(picStream, System.Drawing.Imaging.ImageFormat.Png)
        Dim imgArray() As Byte = picStream.GetBuffer()
        picStream.Close()
        Dim photoQuery As String = "Update student Set Img=@photo where ID = '" & editIDNumTB.Text & "'"
        SQLConnection.executePhotoCommand(photoQuery, imgArray)
        MsgBox("Student Edited Successfully!")
    End Sub
    Private Sub BunifuThinButton231_Click(sender As Object, e As EventArgs) Handles BunifuThinButton231.Click
        Dim i As Integer
        If showFullNameTB.Text = "" Then
            MsgBox("Select Student to delete")
            Return
        End If
        If MessageBox.Show("Are you sure you want to delete Student", "Delete", MessageBoxButtons.OKCancel) = DialogResult.OK Then
            SQLConnection.executeCommand("DELETE from student where id='" & showIDNumber.Text & "'")
            SQLConnection.executeCommand("DELETE from attendance where Student_Code='" & showIDNumber.Text & "'")
            SQLConnection.executeCommand("DELETE from student_subjects where student_id='" & showIDNumber.Text & "'")
            For i = 0 To 100 Step 1
                System.IO.File.Delete(Directory.GetCurrentDirectory() & "\TrainingImages\" & showIDNumber.Text & "_" & i & ".jpg")
            Next
            dt = SQLConnection.executeQuery("select Name as 'Full Name' , ID as 'ID Number' , CONCAT(Course,' ' ,Year) as 'Course and Year' from student")
            BunifuCustomDataGrid1.DataSource = dt
            showFullNameTB.Text = ""
            showIDNumber.Text = ""
            showCoursTB.Text = ""
            PictureBox4.Image = Nothing
            BunifuCustomDataGrid6.DataSource = Nothing
        End If
    End Sub
    Private Sub FACULTYToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ComboBox4.Text = ""
    End Sub

    Private Sub BunifuCustomDataGrid5_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles BunifuCustomDataGrid5.CellContentClick

    End Sub

    Private Sub BunifuThinButton224_Click(sender As Object, e As EventArgs) Handles BunifuThinButton224.Click
        saveToExcel()
        'wew watdahek men
    End Sub
    Private Sub saveToExcel()
        Dim xlApp As New Microsoft.Office.Interop.Excel.Application()
        Dim xlWorkbook As Microsoft.Office.Interop.Excel.Workbook = xlApp.Workbooks.Open(Application.StartupPath & "/ExcelFiles/template.xlsx")

        Dim xlWorksheet As Microsoft.Office.Interop.Excel.Worksheet = xlWorkbook.ActiveSheet

        xlWorksheet.Cells(1, 3) = "Instructor: " & ComboBox1.Text
        xlWorksheet.Cells(2, 3) = "Subject Code: " & ComboBox2.Text
        Dim q As DataTable = SQLConnection.executeQuery("Select Room from subject where code='" & ComboBox2.Text & "'")
        xlWorksheet.Cells(3, 3) = "Room: " & q.Rows.Item(0).Item(0).ToString()

        Dim colIndex As Integer = 4
        For Each col As DataGridViewColumn In BunifuCustomDataGrid7.Columns
            xlWorksheet.Cells(4, colIndex) = col.Name
            colIndex += 1
        Next

        Dim rowIndex As Integer = 5
        For Each row As DataGridViewRow In BunifuCustomDataGrid7.Rows
            colIndex = 4
            xlWorksheet.Cells(rowIndex, 3) = rowIndex - 4
            For Each col As DataGridViewColumn In BunifuCustomDataGrid7.Columns
                xlWorksheet.Cells(rowIndex, colIndex) = row.Cells(colIndex - 4).Value.ToString()
                colIndex += 1
            Next
            rowIndex += 1
        Next


        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx"
        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            xlWorkbook.SaveAs(saveFileDialog.FileName)
        End If

        xlWorkbook.Close()
        xlApp.Quit()
        ReleaseObject(xlWorksheet)
        ReleaseObject(xlWorkbook)
        ReleaseObject(xlApp)
    End Sub
    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Me.Invoke(Sub() showStatus())
        FaceRecognizer.fModel.isTrained = False
        FaceRecognizer.fModel.TrainModel()
        ToolStripStatusLabel1.Text = "Face Detection Training Completed!"
        Threading.Thread.Sleep(10000)
        If Not closeForm Then
            If Me.InvokeRequired Then
                Me.Invoke(Sub() hideStatus())
            End If
        End If
    End Sub

    Private Sub hideStatus()
        StatusStrip1.Visible = False
    End Sub

    Private Sub showStatus()
        ToolStripStatusLabel1.Text = "Training Face Detection Model"
        StatusStrip1.Visible = True
    End Sub

    Private Sub BunifuMaterialTextbox15_OnValueChanged(sender As Object, e As EventArgs) Handles BunifuMaterialTextbox15.OnValueChanged
        dt = SQLConnection.executeQuery("Select Name As 'Full Name' , ID as 'ID Number' , CONCAT(Course,' ' ,Year) as 'Course and Year' from student where Name like '%" & BunifuMaterialTextbox15.Text & "%'")
        BunifuCustomDataGrid1.DataSource = dt
    End Sub

    Private Sub BunifuMaterialTextbox17_OnValueChanged(sender As Object, e As EventArgs) Handles BunifuMaterialTextbox17.OnValueChanged
        dtStaff = SQLConnection.executeQuery("select num as 'No.', Name as 'Full Name' , username as 'Username' from staff where Type='Faculty' and Name like '%" &
                                             BunifuMaterialTextbox17.Text & "%'")

        BunifuCustomDataGrid2.Rows.Clear()
        Dim dgvID As Integer = 0
        For i As Integer = 0 To dtStaff.Rows.Count - 1
            dgvID = BunifuCustomDataGrid2.Rows.Add()
            BunifuCustomDataGrid2.Rows(dgvID).Cells.Item(0).Value = i + 1
            BunifuCustomDataGrid2.Rows(dgvID).Cells.Item(1).Value = dtStaff.Rows.Item(i).Item(1).ToString()
            BunifuCustomDataGrid2.Rows(dgvID).Cells.Item(2).Value = dtStaff.Rows.Item(i).Item(2).ToString()
        Next
    End Sub

    Private Sub BunifuThinButton232_Click(sender As Object, e As EventArgs) Handles BunifuThinButton232.Click
        BunifuThinButton232.Visible = False
        editSubjectState = False
        clearAddSubject()

        Label22.Text = "Add new Subject"
        MsgBox("Edit View Exited!")
    End Sub

    Private Sub BunifuCustomDataGrid3_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles BunifuCustomDataGrid3.RowHeaderMouseClick

    End Sub

    Private Sub BunifuThinButton233_Click(sender As Object, e As EventArgs) Handles BunifuThinButton233.Click
        switchExport = Not switchExport
        If switchExport Then
            BunifuThinButton233.ButtonText = "SHOW ATTENDANCE"
            BunifuCustomDataGrid7.Columns.Clear()
            BunifuCustomDataGrid7.Rows.Clear()

            BunifuCustomDataGrid7.Columns.Add("ID", "ID Number")
            BunifuCustomDataGrid7.Columns.Add("Name", "Name")

            Dim reportsDT As DataTable = SQLConnection.executeQuery("Select dateSet from attendanceschedule where subjectCode = '" & _dtForSubject.Rows.Item(ComboBox2.SelectedIndex).Item(1).ToString() & "'")

            Dim lastDate As String = ""
            For Each items As DataRow In reportsDT.Rows
                Dim d As String = Convert.ToDateTime(items.Item(0).ToString()).ToString("MMMM dd")
                If Not lastDate = d Then
                    BunifuCustomDataGrid7.Columns.Add(d, d)
                End If
                lastDate = d
            Next

            Dim idList As New List(Of personDetail)
            reportsDT = SQLConnection.executeQuery("select sub.student_id , stud.name from student_subjects sub join student stud on sub.student_id = stud.ID " &
                 "where sub.subject_id = '" & _dtForSubject.Rows.Item(ComboBox2.SelectedIndex).Item(1).ToString() & "'")
            For Each items As DataRow In reportsDT.Rows
                Dim studentDetail As New personDetail(items.Item(1).ToString(), items.Item(0).ToString())
                idList.Add(studentDetail)
            Next

            reportsDT = SQLConnection.executeQuery("Select time_in from subject where code = '" & cbBoxContents(ComboBox2.SelectedIndex).id & "'")
            Dim TimeOutCurrent As DateTime = Convert.ToDateTime(reportsDT.Rows.Item(0).Item(0).ToString())

            For Each item As personDetail In idList
                Dim rowID As Integer = BunifuCustomDataGrid7.Rows.Add()
                BunifuCustomDataGrid7.Rows.Item(rowID).Cells.Item(0).Value = item.id
                BunifuCustomDataGrid7.Rows.Item(rowID).Cells.Item(1).Value = item.name
                reportsDT = SQLConnection.executeQuery("Select * from attendance where Student_Code = '" & item.id & "' and Student_Subject='" & _dtForSubject.Rows.Item(ComboBox2.SelectedIndex).Item(1).ToString() & "'")
                For Each items As DataRow In reportsDT.Rows
                    Dim d As String = Convert.ToDateTime(items.Item(2).ToString()).ToString("MMMM dd")
                    BunifuCustomDataGrid7.Rows.Item(rowID).Cells.Item(d).Value = Convert.ToDateTime(items.Item(2).ToString()).ToString("hh:mm tt")
                Next
                For Each items As DataGridViewCell In BunifuCustomDataGrid7.Rows.Item(rowID).Cells
                    If items.Value = Nothing Or items.Value = "" Then
                        items.Value = "A"
                    End If
                Next
            Next
        Else
            BunifuThinButton233.ButtonText = "SHOW TIME"
            BunifuCustomDataGrid7.Columns.Clear()
            BunifuCustomDataGrid7.Rows.Clear()

            BunifuCustomDataGrid7.Columns.Add("ID", "ID Number")
            BunifuCustomDataGrid7.Columns.Add("Name", "Name")

            Dim reportsDT As DataTable = SQLConnection.executeQuery("Select dateSet from attendanceschedule where subjectCode = '" & _dtForSubject.Rows.Item(ComboBox2.SelectedIndex).Item(1).ToString() & "'")

            Dim lastDate As String = ""
            For Each items As DataRow In reportsDT.Rows
                Dim d As String = Convert.ToDateTime(items.Item(0).ToString()).ToString("MMMM dd")
                If Not lastDate = d Then
                    BunifuCustomDataGrid7.Columns.Add(d, d)
                End If
                lastDate = d
            Next

            Dim idList As New List(Of personDetail)
            reportsDT = SQLConnection.executeQuery("select sub.student_id , stud.name from student_subjects sub join student stud on sub.student_id = stud.ID " &
             "where sub.subject_id = '" & _dtForSubject.Rows.Item(ComboBox2.SelectedIndex).Item(1).ToString() & "'")
            For Each items As DataRow In reportsDT.Rows
                Dim studentDetail As New personDetail(items.Item(1).ToString(), items.Item(0).ToString())
                idList.Add(studentDetail)
            Next

            reportsDT = SQLConnection.executeQuery("Select time_in from subject where code = '" & _dtForSubject.Rows.Item(ComboBox2.SelectedIndex).Item(1).ToString() & "'")
            Dim TimeOutCurrent As DateTime = Convert.ToDateTime(reportsDT.Rows.Item(0).Item(0).ToString())

            For Each item As personDetail In idList
                Dim rowID As Integer = BunifuCustomDataGrid7.Rows.Add()
                BunifuCustomDataGrid7.Rows.Item(rowID).Cells.Item(0).Value = item.id
                BunifuCustomDataGrid7.Rows.Item(rowID).Cells.Item(1).Value = item.name
                reportsDT = SQLConnection.executeQuery("Select * from attendance where Student_Code = '" & item.id & "' and Student_Subject='" & _dtForSubject.Rows.Item(ComboBox2.SelectedIndex).Item(1).ToString() & "'")
                For Each items As DataRow In reportsDT.Rows
                    Dim d As String = Convert.ToDateTime(items.Item(2).ToString()).ToString("MMMM dd")
                    If TimeOutCurrent.AddMinutes(15).TimeOfDay < Convert.ToDateTime(reportsDT.Rows.Item(0).Item(2).ToString).TimeOfDay Then
                        BunifuCustomDataGrid7.Rows.Item(rowID).Cells.Item(d).Value = "T"
                    Else
                        BunifuCustomDataGrid7.Rows.Item(rowID).Cells.Item(d).Value = "P"
                    End If
                Next
                For Each items As DataGridViewCell In BunifuCustomDataGrid7.Rows.Item(rowID).Cells
                    If items.Value = Nothing Or items.Value = "" Then
                        items.Value = "A"
                    End If
                Next
            Next
        End If


        Dim colData As New List(Of DataGridViewColumn)
        For Each items As DataGridViewColumn In BunifuCustomDataGrid7.Columns
            colData.Add(items)
        Next
        For Each col As DataGridViewColumn In colData
            Dim rowCounter As Integer = 0
            For Each row As DataGridViewRow In BunifuCustomDataGrid7.Rows
                If BunifuCustomDataGrid7.Rows.Item(row.Index).Cells(col.Index).Value = Nothing Or BunifuCustomDataGrid7.Rows.Item(row.Index).Cells(col.Index).Value = "" Then
                    rowCounter += 1
                End If
            Next
            If rowCounter > 0 Then
                BunifuCustomDataGrid7.Columns.Remove(col)
            End If
        Next
    End Sub

    Private Sub BunifuMaterialTextbox18_OnValueChanged(sender As Object, e As EventArgs) Handles BunifuMaterialTextbox18.OnValueChanged
        Dim dtNew As DataTable = SQLConnection.executeQuery("Select name,code,TIME_FORMAT(time_in, '%I:%i %p') as Time_In ,TIME_FORMAT(time_out, '%I:%i %p') as Time_Out, Room  from subject where name like '%" & BunifuMaterialTextbox18.Text & "%'")
        BunifuCustomDataGrid5.DataSource = dtNew
    End Sub

    Private Sub CHANGEPWORDToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CHANGEPWORDToolStripMenuItem.Click
        AdminPassword.Show()
    End Sub
    Public Sub saveTrainingImage(ByVal lbl As String, img As Image)
        Dim path As String = Directory.GetCurrentDirectory() & "\TrainingImages"
        Dim saveGrayImage As Image = img
        If File.Exists(path & "\" + lbl + ".jpg") Then
            File.Delete(path & "\" + lbl + ".jpg")
        End If
        saveGrayImage.Save(path & "\" + lbl + ".jpg", Imaging.ImageFormat.Jpeg)
    End Sub

    Private Sub MainForm_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        If Not BackgroundWorker1.IsBusy And Not FaceRecognizer.fModel.isTrained Then
            BackgroundWorker1.RunWorkerAsync()
        End If
        BackgroundWorker1.WorkerSupportsCancellation = True
        BackgroundWorker2.WorkerSupportsCancellation = True
    End Sub

    Private Sub BackgroundWorker2_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        While BackgroundWorker1.IsBusy

        End While
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub MainForm_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        If WindowState = FormWindowState.Maximized Then
            BunifuCustomDataGrid3.Size = New System.Drawing.Point(830, (AssignSubjectFacultyPanel.Size.Height / 2) - 100)
            Label57.Location = New System.Drawing.Point(541, (AssignSubjectFacultyPanel.Size.Height / 2) + 22)
            BunifuThinButton222.Location = New System.Drawing.Point(1203, (AssignSubjectFacultyPanel.Size.Height / 2) + 12)
            BunifuCustomDataGrid4.Size = New System.Drawing.Point(830, (AssignSubjectFacultyPanel.Size.Height / 2) - 100)
            BunifuCustomDataGrid4.Location = New System.Drawing.Point(541, (AssignSubjectFacultyPanel.Size.Height / 2) + 53)
        Else
            BunifuCustomDataGrid3.Size = New System.Drawing.Point(830, 189)
            Label57.Location = New System.Drawing.Point(11, 306)
            BunifuThinButton222.Location = New System.Drawing.Point(674, 292)
            BunifuCustomDataGrid4.Location = New System.Drawing.Point(12, 335)
        End If
    End Sub
End Class

Public Structure personDetail
    Public Property name As String
    Public Property id As String
    Public Sub New(_name As String, _id As String)
        name = _name
        id = _id
    End Sub
End Structure
