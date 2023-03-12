Imports BunifuAnimatorNS
Imports AForge
Imports AForge.Video
Imports AForge.Video.DirectShow
Imports System.IO
Imports System.Text
Imports System.Drawing

Public Class MainForm
    Private videocapture As VideoCaptureDevice
    Public loginType As String
    Public loginName As String = ""
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
    Private workCancelled As Boolean = False

    Private facemodel As FaceModel

    Private Sub BunifuTileButton1_Click(sender As Object, e As EventArgs) Handles FacultyTile.Click
        BunifuTransition1.HideSync(FacultyTile)
        BunifuTransition2.ShowSync(FacultyPanel)
    End Sub


    Private Sub DefaultAll()
        BunifuTransition2.HideSync(FacultyPanel)
        BunifuTransition1.ShowSync(FacultyTile)

        BunifuTransition1.HideSync(regStudentTile)
        BunifuTransition1.HideSync(listStudentTile)
        BunifuTransition2.ShowSync(studentTile)
    End Sub

    Private Sub BunifuTileButton2_Click(sender As Object, e As EventArgs) Handles studentTile.Click
        BunifuTransition2.HideSync(studentTile)
        BunifuTransition1.ShowSync(regStudentTile)
        BunifuTransition1.ShowSync(listStudentTile)
    End Sub

    Public Sub numberring()
        For rowNum As Integer = 0 To BunifuCustomDataGrid2.Rows.Count - 1
            BunifuCustomDataGrid2.Rows(rowNum).HeaderCell.Value = (rowNum + 1).ToString
        Next
    End Sub
    Private Sub bunifucustomdatagrid2_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles BunifuCustomDataGrid2.CellFormatting
        'numberring()
    End Sub


    Private Sub BunifuImageButton1_Click(sender As Object, e As EventArgs)
        DefaultAll()
    End Sub

    Private Sub Panel2_Click(sender As Object, e As EventArgs) Handles mainPanel.Click
        DefaultAll()
    End Sub

    Private Sub BunifuTileButton6_Click(sender As Object, e As EventArgs) Handles BunifuTileButton6.Click
        Me.Hide()
        FaceRecog.Show()
    End Sub

    Private Sub regStudentTile_Click(sender As Object, e As EventArgs) Handles regStudentTile.Click
        DefaultAll()
        registerStudentPanel.Show()
        BunifuTransition2.HideSync(mainPanel)
    End Sub

    Private Sub BunifuImageButton3_Click(sender As Object, e As EventArgs) Handles BunifuImageButton3.Click
        mainPanel.Show()
        BunifuTransition1.HideSync(registerStudentPanel)
    End Sub

    Private Sub BunifuThinButton22_Click(sender As Object, e As EventArgs) Handles BunifuThinButton22.Click
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
        MsgBox("Student Added!")
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub defaultAddStudent()
        fnameTB.Text = ""
        mNameTB.Text = ""
        lNameTB.Text = ""
        idTB.Text = ""
        yearTB.Text = ""


        profilePB.Image = My.Resources.icons8_student_male_96
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
        BunifuCustomDataGrid2.RowHeadersVisible = False
        facemodel = New FaceModel(Nothing, Nothing, True)
        BackgroundWorker1.RunWorkerAsync()
        BackgroundWorker1.WorkerSupportsCancellation = True
        If loginType = "Faculty" Then
            FACULTYToolStripMenuItem.Visible = False
            STUDENTToolStripMenuItem.Visible = False
        End If
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

        Dim dtNew As DataTable = SQLConnection.executeQuery("select name,code,TIME_FORMAT(time_in, '%H:%i') as Time_In ,TIME_FORMAT(time_out, '%H:%i') as Time_Out, Room from subject")
        BunifuCustomDataGrid5.DataSource = dtNew

        HideAllPanel()
        HideAllPanel()
        HideAllPanel()
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

    Private Sub BunifuImageButton1_Click_1(sender As Object, e As EventArgs) Handles BunifuImageButton1.Click
        mainPanel.Show()
        BunifuTransition1.HideSync(ListStudentPanel)
    End Sub

    Private Sub ListStudentTile_Click(sender As Object, e As EventArgs) Handles listStudentTile.Click
        DefaultAll()
        ListStudentPanel.Show()
        BunifuTransition1.HideSync(mainPanel)
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

    Private Sub BunifuImageButton4_Click(sender As Object, e As EventArgs) Handles BunifuImageButton4.Click
        ListStudentPanel.Show()
        BunifuTransition1.HideSync(editStudentPanel)
    End Sub
    Private Sub BunifuThinButton230_Click(sender As Object, e As EventArgs) Handles BunifuThinButton230.Click
        Dim q As String = "Update student set Name='" & editFullNameTB.Text & "' where ID='" & editIDNumTB.Text & "'"
        SQLConnection.executeCommand(q)
        saveNewEditStudentImage()
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Public Sub RefreshRegisterImage()
        Dim _images As New ImageList
        _images.ImageSize = New Drawing.Point(64, 64)
        Dim _imgNames As New List(Of ListViewItem)
        Dim _path As String = Directory.GetCurrentDirectory() & "\TrainingImages"
        For Each path In Directory.GetFiles(_path)
            If path.Contains(showIDNumber.Text) Then
                _images.Images.Add(Image.FromFile(path))
                _imgNames.Add(New ListViewItem(IO.Path.GetFileName(path)) With {.ImageIndex = _images.Images.Count - 1})
            End If
        Next
        ListView1.Items.Clear()
        ListView1.LargeImageList = _images
        ListView1.Items.AddRange(_imgNames.ToArray)
    End Sub

    Public Sub RefreshListImage()
        Dim _images As New ImageList
        _images.ImageSize = New Drawing.Point(64, 64)
        Dim _imgNames As New List(Of ListViewItem)
        Dim _path As String = Directory.GetCurrentDirectory() & "\TrainingImages"
        For Each path In Directory.GetFiles(_path)
            If path.Contains(editIDNumTB.Text) Then
                _images.Images.Add(Image.FromFile(path))
                _imgNames.Add(New ListViewItem(IO.Path.GetFileName(path)) With {.ImageIndex = _images.Images.Count - 1})
            End If
        Next
        ListView2.Items.Clear()
        ListView2.LargeImageList = _images
        ListView2.Items.AddRange(_imgNames.ToArray)
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

        Dim query As String = "select subject.name as 'Subject' ,student_subjects.subject_id as 'Code', TIME_FORMAT(subject.time_in,'%H:%i') as 'Time in', TIME_FORMAT(subject.time_out,'%H:%i') as 'Time out' From student_subjects " &
            "INNER JOIN subject ON subject.code = student_subjects.subject_id WHERE student_subjects.student_id ='" & editIDNumTB.Text & "'"
        listofStudentSubjectDGV.DataSource = SQLConnection.executeQuery(query)
        Dim q As String = "SELECT s.name,t.code,TIME_FORMAT(s.time_in, '%H:%i') as Time_In ,TIME_FORMAT(s.time_out, '%H:%i') as Time_Out " &
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

        query = "select subject.name as 'Subject' ,student_subjects.subject_id as 'Code', TIME_FORMAT(subject.time_in,'%H:%i') as 'Time in', TIME_FORMAT(subject.time_out,'%H:%i') as 'Time out' From student_subjects " &
            "INNER JOIN subject ON subject.code = student_subjects.subject_id WHERE student_subjects.student_id ='" & editIDNumTB.Text & "'"

        listofStudentSubjectDGV.DataSource = SQLConnection.executeQuery(query)
        MsgBox("Succesfully Added to your Subjects")


        Dim q As String = "SELECT s.name,t.code,TIME_FORMAT(s.time_in, '%H:%i') as Time_In ,TIME_FORMAT(s.time_out, '%H:%i') as Time_Out " &
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

        query = "select subject.name as 'Subject' ,student_subjects.subject_id as 'Code', TIME_FORMAT(subject.time_in,'%H:%i') as 'Time in', TIME_FORMAT(subject.time_out,'%H:%i') as 'Time out' From student_subjects " &
            "INNER JOIN subject ON subject.code = student_subjects.subject_id WHERE student_subjects.student_id ='" & editIDNumTB.Text & "'"

        listofStudentSubjectDGV.DataSource = SQLConnection.executeQuery(query)

        Dim q As String = "SELECT s.name,t.code,TIME_FORMAT(s.time_in, '%H:%i') as Time_In ,TIME_FORMAT(s.time_out, '%H:%i') as Time_Out " &
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

    Private Sub EXITToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EXITToolStripMenuItem.Click
        Application.Exit()
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

    Private Sub BunifuThinButton213_Click(sender As Object, e As EventArgs) Handles BunifuThinButton213.Click
        If editSubjectState Then
            Dim q As String = "Update subject set name='" & BunifuMaterialTextbox1.Text & "',time_in='" &
                DateTimePicker3.Value.ToString("HH:mm:ss") & "',time_out='" & DateTimePicker4.Value.ToString("HH:mm:ss") & "',Mon='" & Convert.ToInt32(CheckBox6.Checked) &
                "',Tue='" & Convert.ToInt32(CheckBox7.Checked) & "',Wed='" & Convert.ToInt32(CheckBox8.Checked) & "',Thu='" &
                Convert.ToInt32(CheckBox9.Checked) & "',Fri='" & Convert.ToInt32(CheckBox10.Checked) & "',Sat='" &
                Convert.ToInt32(CheckBox11.Checked) & "',Sun='" & Convert.ToInt32(CheckBox12.Checked) & "',Room='" &
                BunifuMaterialTextbox16.Text & "' where code='" & BunifuMaterialTextbox2.Text & "'"
            SQLConnection.executeCommand(q)
            editSubjectState = False

            BunifuThinButton232.Enabled = False
            BunifuThinButton229.Enabled = True
            Label22.Text = "Add new Subject"

            Dim dtNews As DataTable = SQLConnection.executeQuery("select name,code,TIME_FORMAT(time_in, '%H:%i') as Time_In ,TIME_FORMAT(time_out, '%H:%i') as Time_Out,Room  from subject")
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
                MsgBox("Error")
            End If
            Return
        End Try
        BunifuMaterialTextbox1.Text = ""
        BunifuMaterialTextbox2.Text = ""

        MsgBox("Subject Successfully Saved!")
        Dim dtNew As DataTable = SQLConnection.executeQuery("Select name,code,TIME_FORMAT(time_in, '%H:%i') as Time_In ,TIME_FORMAT(time_out, '%H:%i') as Time_Out, Room  from subject")
        BunifuCustomDataGrid5.DataSource = dtNew
    End Sub

    Private Sub ADDSUBJECTToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ADDSUBJECTToolStripMenuItem.Click

        Label22.Text = "Add new Subject"
        BunifuThinButton232.Enabled = False
        BunifuThinButton229.Enabled = True
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
            Dim nms As New IO.MemoryStream(CType(dtFill.Rows().Item(0).Item(5), Byte()))
            Dim nreturnImage As Image = Image.FromStream(nms)
            If Not IsNothing(nreturnImage) Then
                PictureBox7.Image = Nothing
                PictureBox7.Image = nreturnImage
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
        ComboBox3.Text = ""
        Panel3.Show()

        If loginName = "" Then
            Dim query As String = "Select * from staff where Type='Faculty'"
            Dim dt As DataTable = SQLConnection.executeQuery(query)
            FacultyList = dt.Rows

            For i As Integer = 0 To dt.Rows.Count - 1
                ComboBox1.Items.Add(dt.Rows.Item(i).Item(1).ToString)
            Next
        Else
            Dim query As String = "Select * from staff where Type='Faculty' and username='" & loginName & "'"
            Dim dt As DataTable = SQLConnection.executeQuery(query)
            FacultyList = dt.Rows
            For i As Integer = 0 To dt.Rows.Count - 1
                ComboBox1.Items.Add(dt.Rows.Item(i).Item(1).ToString)
            Next
            ComboBox1.SelectedIndex = 0
        End If
        HideAllPanel()
        FreePanel.Show()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox2.Items.Clear()
        Dim query As String = "Select * from subject where faculty='" & FacultyList.Item(ComboBox1.SelectedIndex).Item(3).ToString & "'"
        Dim dt As DataTable = SQLConnection.executeQuery(query)
        SubjectList = dt.Rows

        For i As Integer = 0 To dt.Rows.Count - 1
            ComboBox2.Items.Add(dt.Rows.Item(i).Item(0).ToString)
        Next

    End Sub

    Private Sub BunifuThinButton223_Click(sender As Object, e As EventArgs) Handles BunifuThinButton223.Click
        Panel3.Hide()
        BunifuThinButton225.Visible = True
        BunifuThinButton224.Visible = True
        dgvAttendance.Show()
        Label65.Visible = True
        Label65.Text = "Month: " & ComboBox3.SelectedItem.ToString() & " / " & "Subject: " & ComboBox2.SelectedItem.ToString()
        dgvAttendance.Columns.Clear()
        dgvAttendance.Rows.Clear()

        dgvAttendance.Columns.Add("ID", "ID Number")
        dgvAttendance.Columns.Add("Name", "Name")

        For i As Integer = 1 To 31
            dgvAttendance.Columns.Add("day" & i, i.ToString)
        Next

        Dim selectedDate As Integer = ComboBox3.SelectedIndex - 1
        Dim query As String = "select * from student_subjects where subject_id='" & SubjectList.Item(ComboBox2.SelectedIndex).Item(1).ToString & "'"
        Dim dts As DataTable = SQLConnection.executeQuery(query)

        Dim idList As New List(Of String)
        For i As Integer = 0 To dts.Rows.Count - 1
            If Not idList.Contains(dts.Rows.Item(i).Item(0).ToString) Then
                idList.Add(dts.Rows.Item(i).Item(0).ToString)
            End If
        Next

        query = "Select time_out from subject where code='" & SubjectList.Item(ComboBox2.SelectedIndex).Item(1).ToString & "'"
        Dim timeOutdt = SQLConnection.executeQuery(query)
        'Dim newDateTaken As DateTime
        Dim TimeOutCurrent As DateTime = Convert.ToDateTime(timeOutdt.Rows.Item(0).Item(0).ToString)

        Dim attendanceScheduleDT As DataTable = SQLConnection.executeQuery("Select * from attendanceschedule where subjectCode = '" & SubjectList.Item(ComboBox2.SelectedIndex).Item(1).ToString & "' and MONTH(dateSet)=" & (ComboBox3.SelectedIndex + 1))

        TimeOutCurrent.AddMinutes(15)
        'For Each ids In idList
        '    Dim dgvID = dgvAttendance.Rows.Add()
        '    dgvAttendance.Rows(dgvID).Cells(0).Value = ids
        '    Dim dt As DataTable = SQLConnection.executeQuery("select Name from student where ID='" & ids & "'")
        '    dgvAttendance.Rows(dgvID).Cells(1).Value = dt.Rows.Item(0).Item(0).ToString()
        '    For i As Integer = 0 To dts.Rows.Count - 1
        '        If ids = dts.Rows.Item(i).Item(0).ToString Then
        '            newDateTaken = Convert.ToDateTime(dts.Rows.Item(i).Item(2).ToString)
        '            If newDateTaken < TimeOutCurrent Then
        '                dgvAttendance.Rows(dgvID).Cells(newDateTaken.Day + 1).Value = "P"
        '            ElseIf newDateTaken > TimeOutCurrent Then
        '                dgvAttendance.Rows(dgvID).Cells(newDateTaken.Day + 1).Value = "T"
        '            Else
        '                dgvAttendance.Rows(dgvID).Cells(newDateTaken.Day + 1).Value = "A"
        '            End If
        '        End If
        '    Next
        'Next
        For Each ids In idList
            Dim dgvID = dgvAttendance.Rows.Add()
            dgvAttendance.Rows(dgvID).Cells(0).Value = ids
            Dim dt As DataTable = SQLConnection.executeQuery("select Name from student where ID='" & ids & "'")
            dgvAttendance.Rows(dgvID).Cells(1).Value = dt.Rows.Item(0).Item(0).ToString()
            For Each items In attendanceScheduleDT.Rows
                Dim attDate = Convert.ToDateTime(items.Item(1).ToString)
                Dim q As String = "Select * from attendance where MONTH(Time_Taken)=" & (ComboBox3.SelectedIndex + 1) & " and DAY(Time_Taken)=" & attDate.Day.ToString() & " and Student_Code='" & ids &
                    "' and Student_Subject='" & SubjectList.Item(ComboBox2.SelectedIndex).Item(1).ToString & "'"

                Dim studentsAtt As DataTable = SQLConnection.executeQuery(q)

                If studentsAtt.Rows.Count <= 0 Then
                    dgvAttendance.Rows(dgvID).Cells(attDate.Day + 1).Value = "A"
                Else
                    dgvAttendance.Rows(dgvID).Cells(attDate.Day + 1).Value = "P"
                End If
            Next
        Next
    End Sub

    Private Sub BunifuThinButton225_Click(sender As Object, e As EventArgs) Handles BunifuThinButton225.Click
        dgvAttendance.Rows.Clear()
        dgvAttendance.Visible = False
        BunifuThinButton225.Visible = False
        BunifuThinButton224.Visible = False
        Label65.Visible = False
        Panel3.Show()
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        dtFaculty = SQLConnection.executeQuery("select name as 'Subject' , code as 'Code' from subject where faculty='" & FacultyAssignList.Item(ComboBox4.SelectedIndex).Item(3) & "'")
                BunifuCustomDataGrid3.DataSource = dtFaculty
    End Sub

    Private Sub LOGOUTToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LOGOUTToolStripMenuItem.Click
        If MessageBox.Show("Are you sure you want to Logout?", "Logout?", MessageBoxButtons.OKCancel) = DialogResult.OK Then
            Me.Hide()
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
        MsgBox("Successfully Removed!")
        Dim dtNew As DataTable = SQLConnection.executeQuery("select name,code,time_in,time_out,Room from subject")
        BunifuCustomDataGrid5.DataSource = dtNew
    End Sub
    Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If BackgroundWorker1.IsBusy Then
            workCancelled = True
            BackgroundWorker1.CancelAsync()
            NotifyIcon1.Dispose()
            e.Cancel = True
        End If
        If LogoutState Then
            LogoutState = False
        Else
            Me.Hide()
            NotifyIcon1.Visible = True
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

    Private Sub BunifuThinButton227_Click(sender As Object, e As EventArgs) Handles BunifuThinButton227.Click
        Dim query As String = "Update staff set Name='" & BunifuMaterialTextbox14.Text & "' , username='" & BunifuMaterialTextbox13.Text &
            "', password='" & BunifuMaterialTextbox12.Text & "', ID='" & BunifuMaterialTextbox11.Text & "' where ID='" & BunifuMaterialTextbox11.Text & "'"
        SQLConnection.executeCommand(query)
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

        BunifuThinButton232.Enabled = True
        BunifuThinButton229.Enabled = False
        Dim q As String = "Select * from subject where code='" & BunifuCustomDataGrid5.SelectedRows.Item(0).Cells.Item(1).Value.ToString() & "'"
        Dim editSubjectDT As DataTable = SQLConnection.executeQuery(q)
        editSubjectState = True
        BunifuMaterialTextbox1.Enabled = False
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
        Dim q As String = "SELECT s.name,t.code,TIME_FORMAT(s.time_in, '%H:%i') as Time_In ,TIME_FORMAT(s.time_out, '%H:%i') as Time_Out " &
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
        MsgBox("Faculty Edited Successfully!")
    End Sub
    Private Sub BunifuThinButton231_Click(sender As Object, e As EventArgs) Handles BunifuThinButton231.Click

        If showFullNameTB.Text = "" Then
            MsgBox("Select Student to delete")
            Return
        End If
        If MessageBox.Show("Are you sure you want to delete Student", "Delete", MessageBoxButtons.OKCancel) = DialogResult.OK Then
            SQLConnection.executeCommand("DELETE from student where id='" & showIDNumber.Text & "'")
            dt = SQLConnection.executeQuery("select Name as 'Full Name' , ID as 'ID Number' , CONCAT(Course,' ' ,Year) as 'Course and Year' from student")
            BunifuCustomDataGrid1.DataSource = dt
            showFullNameTB.Text = ""
            showIDNumber.Text = ""
            showCoursTB.Text = ""
            PictureBox4.Image = Nothing
            System.IO.File.Delete(Directory.GetCurrentDirectory() & "\TrainingImages\" & showIDNumber.Text & "_Front.jpg")
            System.IO.File.Delete(Directory.GetCurrentDirectory() & "\TrainingImages\" & showIDNumber.Text & "_Top.jpg")
            System.IO.File.Delete(Directory.GetCurrentDirectory() & "\TrainingImages\" & showIDNumber.Text & "_Left.jpg")
            System.IO.File.Delete(Directory.GetCurrentDirectory() & "\TrainingImages\" & showIDNumber.Text & "_Right.jpg")
            System.IO.File.Delete(Directory.GetCurrentDirectory() & "\TrainingImages\" & showIDNumber.Text & "_Bottom.jpg")
        End If
    End Sub
    Private Sub FACULTYToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ComboBox4.Text = ""
    End Sub

    Private Sub BunifuCustomDataGrid5_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles BunifuCustomDataGrid5.CellContentClick

    End Sub

    Private Sub BunifuThinButton224_Click(sender As Object, e As EventArgs) Handles BunifuThinButton224.Click
        DataGridToCSV(dgvAttendance, " ")
    End Sub
    Private Sub DataGridToCSV(ByRef dt As DataGridView, Qualifier As String)
        Dim TempDirectory As String = "A temp Directory"
        System.IO.Directory.CreateDirectory(TempDirectory)
        Dim oWrite As System.IO.StreamWriter
        Dim file As String = ComboBox1.Text & "/" & ComboBox2.Text & "/" & DateTime.Now.ToString() & ".csv"
        oWrite = IO.File.CreateText(TempDirectory & "\" & file)

        Dim CSV As StringBuilder = New StringBuilder()

        Dim i As Integer = 1
        Dim CSVHeader As StringBuilder = New StringBuilder()
        For Each c As DataGridViewColumn In dt.Columns
            If i = 1 Then
                CSVHeader.Append(Qualifier & c.HeaderText.ToString() & Qualifier)
            Else
                CSVHeader.Append("," & Qualifier & c.HeaderText.ToString() & Qualifier)
            End If
            i += 1
        Next

        'CSV.AppendLine(CSVHeader.ToString())
        oWrite.WriteLine(CSVHeader.ToString())
        oWrite.Flush()

        For r As Integer = 0 To dt.Rows.Count - 1

            Dim CSVLine As StringBuilder = New StringBuilder()
            Dim s As String = ""
            For c As Integer = 0 To dt.Columns.Count - 1
                If c = 0 Then
                    Dim val As String
                    If dt.Rows.Item(r).Cells.Item(c).Value = Nothing Then
                        val = ""
                    Else
                        val = dt.Rows.Item(r).Cells.Item(c).Value.ToString
                    End If
                    'CSVLine.Append(Qualifier & gridResults.Rows(r).Cells(c).Value.ToString() & Qualifier)
                    s = s & Qualifier & val & Qualifier
                Else
                    'CSVLine.Append("," & Qualifier & gridResults.Rows(r).Cells(c).Value.ToString() & Qualifier)
                    Dim val As String
                    If dt.Rows.Item(r).Cells.Item(c).Value = Nothing Then
                        val = ""
                    Else
                        val = dt.Rows.Item(r).Cells.Item(c).Value.ToString
                    End If
                    s = s & "," & Qualifier & val & Qualifier
                End If

            Next
            oWrite.WriteLine(s)
            oWrite.Flush()
            'CSV.AppendLine(CSVLine.ToString())
            'CSVLine.Clear()
        Next

        'oWrite.Write(CSV.ToString())

        oWrite.Close()
        oWrite = Nothing

        System.Diagnostics.Process.Start(TempDirectory & "\" & file)

        GC.Collect()

    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Me.Invoke(Sub() showStatus())
        FaceRecognizer.fModel.isTrained = False
        FaceRecognizer.fModel.TrainModel()

        ToolStripStatusLabel1.Text = "Face Detection Training Completed!"
        Threading.Thread.Sleep(10000)
        If Not workCancelled Then
            Me.Invoke(Sub() hideStatus())
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
        BunifuThinButton232.Enabled = False
        BunifuThinButton229.Enabled = True
        editSubjectState = False
        clearAddSubject()

        Label22.Text = "Add new Subject"
        MsgBox("Edit View Exited!")
    End Sub
End Class