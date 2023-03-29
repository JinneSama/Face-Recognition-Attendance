<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FaceRecog
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FaceRecog))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim Animation1 As BunifuAnimatorNS.Animation = New BunifuAnimatorNS.Animation()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.BunifuThinButton23 = New ns1.BunifuThinButton2()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.BunifuThinButton219 = New ns1.BunifuThinButton2()
        Me.BunifuThinButton220 = New ns1.BunifuThinButton2()
        Me.BunifuCustomDataGrid4 = New ns1.BunifuCustomDataGrid()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.IDLB = New System.Windows.Forms.Label()
        Me.NameLB = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.PictureBox3 = New System.Windows.Forms.PictureBox()
        Me.BunifuThinButton22 = New ns1.BunifuThinButton2()
        Me.BunifuThinButton21 = New ns1.BunifuThinButton2()
        Me.BunifuImageButton1 = New ns1.BunifuImageButton()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.BunifuTransition1 = New BunifuAnimatorNS.BunifuTransition(Me.components)
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker()
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        CType(Me.BunifuCustomDataGrid4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BunifuImageButton1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Gray
        Me.Panel1.Controls.Add(Me.BunifuThinButton23)
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.ListBox1)
        Me.Panel1.Controls.Add(Me.IDLB)
        Me.Panel1.Controls.Add(Me.NameLB)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.PictureBox3)
        Me.Panel1.Controls.Add(Me.BunifuThinButton22)
        Me.Panel1.Controls.Add(Me.BunifuThinButton21)
        Me.Panel1.Controls.Add(Me.BunifuImageButton1)
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.BunifuTransition1.SetDecoration(Me.Panel1, BunifuAnimatorNS.DecorationType.None)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(774, 485)
        Me.Panel1.TabIndex = 0
        Me.Panel1.Visible = False
        '
        'BunifuThinButton23
        '
        Me.BunifuThinButton23.ActiveBorderThickness = 1
        Me.BunifuThinButton23.ActiveCornerRadius = 20
        Me.BunifuThinButton23.ActiveFillColor = System.Drawing.Color.Black
        Me.BunifuThinButton23.ActiveForecolor = System.Drawing.Color.White
        Me.BunifuThinButton23.ActiveLineColor = System.Drawing.Color.Black
        Me.BunifuThinButton23.BackColor = System.Drawing.Color.Gray
        Me.BunifuThinButton23.BackgroundImage = CType(resources.GetObject("BunifuThinButton23.BackgroundImage"), System.Drawing.Image)
        Me.BunifuThinButton23.ButtonText = "Treshold"
        Me.BunifuThinButton23.Cursor = System.Windows.Forms.Cursors.Hand
        Me.BunifuTransition1.SetDecoration(Me.BunifuThinButton23, BunifuAnimatorNS.DecorationType.None)
        Me.BunifuThinButton23.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BunifuThinButton23.ForeColor = System.Drawing.Color.Black
        Me.BunifuThinButton23.IdleBorderThickness = 1
        Me.BunifuThinButton23.IdleCornerRadius = 20
        Me.BunifuThinButton23.IdleFillColor = System.Drawing.Color.Transparent
        Me.BunifuThinButton23.IdleForecolor = System.Drawing.Color.Black
        Me.BunifuThinButton23.IdleLineColor = System.Drawing.Color.Black
        Me.BunifuThinButton23.Location = New System.Drawing.Point(471, 418)
        Me.BunifuThinButton23.Margin = New System.Windows.Forms.Padding(5)
        Me.BunifuThinButton23.Name = "BunifuThinButton23"
        Me.BunifuThinButton23.Size = New System.Drawing.Size(138, 43)
        Me.BunifuThinButton23.TabIndex = 69
        Me.BunifuThinButton23.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Gray
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.BunifuThinButton219)
        Me.Panel3.Controls.Add(Me.BunifuThinButton220)
        Me.Panel3.Controls.Add(Me.BunifuCustomDataGrid4)
        Me.BunifuTransition1.SetDecoration(Me.Panel3, BunifuAnimatorNS.DecorationType.None)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Left
        Me.Panel3.Location = New System.Drawing.Point(0, 0)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(442, 485)
        Me.Panel3.TabIndex = 61
        Me.Panel3.Visible = False
        '
        'BunifuThinButton219
        '
        Me.BunifuThinButton219.ActiveBorderThickness = 1
        Me.BunifuThinButton219.ActiveCornerRadius = 20
        Me.BunifuThinButton219.ActiveFillColor = System.Drawing.Color.Black
        Me.BunifuThinButton219.ActiveForecolor = System.Drawing.Color.Transparent
        Me.BunifuThinButton219.ActiveLineColor = System.Drawing.Color.Black
        Me.BunifuThinButton219.BackColor = System.Drawing.Color.Gray
        Me.BunifuThinButton219.BackgroundImage = CType(resources.GetObject("BunifuThinButton219.BackgroundImage"), System.Drawing.Image)
        Me.BunifuThinButton219.ButtonText = "Close"
        Me.BunifuThinButton219.Cursor = System.Windows.Forms.Cursors.Hand
        Me.BunifuTransition1.SetDecoration(Me.BunifuThinButton219, BunifuAnimatorNS.DecorationType.None)
        Me.BunifuThinButton219.Font = New System.Drawing.Font("Consolas", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BunifuThinButton219.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.BunifuThinButton219.IdleBorderThickness = 1
        Me.BunifuThinButton219.IdleCornerRadius = 20
        Me.BunifuThinButton219.IdleFillColor = System.Drawing.Color.Transparent
        Me.BunifuThinButton219.IdleForecolor = System.Drawing.Color.Black
        Me.BunifuThinButton219.IdleLineColor = System.Drawing.Color.Black
        Me.BunifuThinButton219.Location = New System.Drawing.Point(195, 434)
        Me.BunifuThinButton219.Margin = New System.Windows.Forms.Padding(4)
        Me.BunifuThinButton219.Name = "BunifuThinButton219"
        Me.BunifuThinButton219.Size = New System.Drawing.Size(100, 38)
        Me.BunifuThinButton219.TabIndex = 59
        Me.BunifuThinButton219.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'BunifuThinButton220
        '
        Me.BunifuThinButton220.ActiveBorderThickness = 1
        Me.BunifuThinButton220.ActiveCornerRadius = 20
        Me.BunifuThinButton220.ActiveFillColor = System.Drawing.Color.Black
        Me.BunifuThinButton220.ActiveForecolor = System.Drawing.Color.Transparent
        Me.BunifuThinButton220.ActiveLineColor = System.Drawing.Color.Black
        Me.BunifuThinButton220.BackColor = System.Drawing.Color.Gray
        Me.BunifuThinButton220.BackgroundImage = CType(resources.GetObject("BunifuThinButton220.BackgroundImage"), System.Drawing.Image)
        Me.BunifuThinButton220.ButtonText = "Select"
        Me.BunifuThinButton220.Cursor = System.Windows.Forms.Cursors.Hand
        Me.BunifuTransition1.SetDecoration(Me.BunifuThinButton220, BunifuAnimatorNS.DecorationType.None)
        Me.BunifuThinButton220.Font = New System.Drawing.Font("Consolas", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BunifuThinButton220.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.BunifuThinButton220.IdleBorderThickness = 1
        Me.BunifuThinButton220.IdleCornerRadius = 20
        Me.BunifuThinButton220.IdleFillColor = System.Drawing.Color.Transparent
        Me.BunifuThinButton220.IdleForecolor = System.Drawing.Color.Black
        Me.BunifuThinButton220.IdleLineColor = System.Drawing.Color.Black
        Me.BunifuThinButton220.Location = New System.Drawing.Point(302, 434)
        Me.BunifuThinButton220.Margin = New System.Windows.Forms.Padding(4)
        Me.BunifuThinButton220.Name = "BunifuThinButton220"
        Me.BunifuThinButton220.Size = New System.Drawing.Size(100, 38)
        Me.BunifuThinButton220.TabIndex = 58
        Me.BunifuThinButton220.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'BunifuCustomDataGrid4
        '
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.BunifuCustomDataGrid4.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.BunifuCustomDataGrid4.BackgroundColor = System.Drawing.Color.Gainsboro
        Me.BunifuCustomDataGrid4.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.BunifuCustomDataGrid4.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.Gray
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.BunifuCustomDataGrid4.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.BunifuCustomDataGrid4.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.BunifuTransition1.SetDecoration(Me.BunifuCustomDataGrid4, BunifuAnimatorNS.DecorationType.None)
        Me.BunifuCustomDataGrid4.DoubleBuffered = True
        Me.BunifuCustomDataGrid4.EnableHeadersVisualStyles = False
        Me.BunifuCustomDataGrid4.HeaderBgColor = System.Drawing.Color.Gray
        Me.BunifuCustomDataGrid4.HeaderForeColor = System.Drawing.Color.White
        Me.BunifuCustomDataGrid4.Location = New System.Drawing.Point(37, 31)
        Me.BunifuCustomDataGrid4.Name = "BunifuCustomDataGrid4"
        Me.BunifuCustomDataGrid4.ReadOnly = True
        Me.BunifuCustomDataGrid4.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        Me.BunifuCustomDataGrid4.RowHeadersWidth = 51
        Me.BunifuCustomDataGrid4.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.BunifuCustomDataGrid4.Size = New System.Drawing.Size(365, 378)
        Me.BunifuCustomDataGrid4.TabIndex = 24
        Me.BunifuCustomDataGrid4.VirtualMode = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.BunifuTransition1.SetDecoration(Me.Label3, BunifuAnimatorNS.DecorationType.None)
        Me.Label3.Font = New System.Drawing.Font("Consolas", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(25, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(54, 19)
        Me.Label3.TabIndex = 68
        Me.Label3.Text = "Name:"
        '
        'ListBox1
        '
        Me.BunifuTransition1.SetDecoration(Me.ListBox1, BunifuAnimatorNS.DecorationType.None)
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(27, 366)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(282, 108)
        Me.ListBox1.TabIndex = 67
        '
        'IDLB
        '
        Me.IDLB.AutoSize = True
        Me.BunifuTransition1.SetDecoration(Me.IDLB, BunifuAnimatorNS.DecorationType.None)
        Me.IDLB.Font = New System.Drawing.Font("Consolas", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.IDLB.ForeColor = System.Drawing.Color.Black
        Me.IDLB.Location = New System.Drawing.Point(112, 337)
        Me.IDLB.Name = "IDLB"
        Me.IDLB.Size = New System.Drawing.Size(27, 19)
        Me.IDLB.TabIndex = 66
        Me.IDLB.Text = "ID"
        '
        'NameLB
        '
        Me.NameLB.AutoSize = True
        Me.BunifuTransition1.SetDecoration(Me.NameLB, BunifuAnimatorNS.DecorationType.None)
        Me.NameLB.Font = New System.Drawing.Font("Consolas", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NameLB.ForeColor = System.Drawing.Color.Black
        Me.NameLB.Location = New System.Drawing.Point(76, 318)
        Me.NameLB.Name = "NameLB"
        Me.NameLB.Size = New System.Drawing.Size(45, 19)
        Me.NameLB.TabIndex = 65
        Me.NameLB.Text = "Name"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.BunifuTransition1.SetDecoration(Me.Label2, BunifuAnimatorNS.DecorationType.None)
        Me.Label2.Font = New System.Drawing.Font("Consolas", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(25, 337)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(81, 19)
        Me.Label2.TabIndex = 64
        Me.Label2.Text = "ID No. :"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.BunifuTransition1.SetDecoration(Me.Label1, BunifuAnimatorNS.DecorationType.None)
        Me.Label1.Font = New System.Drawing.Font("Consolas", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(25, 318)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(54, 19)
        Me.Label1.TabIndex = 63
        Me.Label1.Text = "Name:"
        '
        'PictureBox3
        '
        Me.PictureBox3.BackColor = System.Drawing.Color.White
        Me.PictureBox3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.BunifuTransition1.SetDecoration(Me.PictureBox3, BunifuAnimatorNS.DecorationType.None)
        Me.PictureBox3.Location = New System.Drawing.Point(27, 44)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(282, 272)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox3.TabIndex = 62
        Me.PictureBox3.TabStop = False
        '
        'BunifuThinButton22
        '
        Me.BunifuThinButton22.ActiveBorderThickness = 1
        Me.BunifuThinButton22.ActiveCornerRadius = 20
        Me.BunifuThinButton22.ActiveFillColor = System.Drawing.Color.Black
        Me.BunifuThinButton22.ActiveForecolor = System.Drawing.Color.White
        Me.BunifuThinButton22.ActiveLineColor = System.Drawing.Color.Black
        Me.BunifuThinButton22.BackColor = System.Drawing.Color.Gray
        Me.BunifuThinButton22.BackgroundImage = CType(resources.GetObject("BunifuThinButton22.BackgroundImage"), System.Drawing.Image)
        Me.BunifuThinButton22.ButtonText = "STOP"
        Me.BunifuThinButton22.Cursor = System.Windows.Forms.Cursors.Hand
        Me.BunifuTransition1.SetDecoration(Me.BunifuThinButton22, BunifuAnimatorNS.DecorationType.None)
        Me.BunifuThinButton22.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BunifuThinButton22.ForeColor = System.Drawing.Color.Black
        Me.BunifuThinButton22.IdleBorderThickness = 1
        Me.BunifuThinButton22.IdleCornerRadius = 20
        Me.BunifuThinButton22.IdleFillColor = System.Drawing.Color.Transparent
        Me.BunifuThinButton22.IdleForecolor = System.Drawing.Color.Black
        Me.BunifuThinButton22.IdleLineColor = System.Drawing.Color.Black
        Me.BunifuThinButton22.Location = New System.Drawing.Point(619, 418)
        Me.BunifuThinButton22.Margin = New System.Windows.Forms.Padding(5)
        Me.BunifuThinButton22.Name = "BunifuThinButton22"
        Me.BunifuThinButton22.Size = New System.Drawing.Size(138, 43)
        Me.BunifuThinButton22.TabIndex = 20
        Me.BunifuThinButton22.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'BunifuThinButton21
        '
        Me.BunifuThinButton21.ActiveBorderThickness = 1
        Me.BunifuThinButton21.ActiveCornerRadius = 20
        Me.BunifuThinButton21.ActiveFillColor = System.Drawing.Color.Black
        Me.BunifuThinButton21.ActiveForecolor = System.Drawing.Color.White
        Me.BunifuThinButton21.ActiveLineColor = System.Drawing.Color.Black
        Me.BunifuThinButton21.BackColor = System.Drawing.Color.Gray
        Me.BunifuThinButton21.BackgroundImage = CType(resources.GetObject("BunifuThinButton21.BackgroundImage"), System.Drawing.Image)
        Me.BunifuThinButton21.ButtonText = "START"
        Me.BunifuThinButton21.Cursor = System.Windows.Forms.Cursors.Hand
        Me.BunifuTransition1.SetDecoration(Me.BunifuThinButton21, BunifuAnimatorNS.DecorationType.None)
        Me.BunifuThinButton21.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BunifuThinButton21.ForeColor = System.Drawing.Color.Black
        Me.BunifuThinButton21.IdleBorderThickness = 1
        Me.BunifuThinButton21.IdleCornerRadius = 20
        Me.BunifuThinButton21.IdleFillColor = System.Drawing.Color.Transparent
        Me.BunifuThinButton21.IdleForecolor = System.Drawing.Color.Black
        Me.BunifuThinButton21.IdleLineColor = System.Drawing.Color.Black
        Me.BunifuThinButton21.Location = New System.Drawing.Point(619, 379)
        Me.BunifuThinButton21.Margin = New System.Windows.Forms.Padding(5)
        Me.BunifuThinButton21.Name = "BunifuThinButton21"
        Me.BunifuThinButton21.Size = New System.Drawing.Size(138, 43)
        Me.BunifuThinButton21.TabIndex = 19
        Me.BunifuThinButton21.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'BunifuImageButton1
        '
        Me.BunifuImageButton1.BackColor = System.Drawing.Color.Transparent
        Me.BunifuTransition1.SetDecoration(Me.BunifuImageButton1, BunifuAnimatorNS.DecorationType.None)
        Me.BunifuImageButton1.Image = Global.ATMOS_V2.My.Resources.Resources.icons8_left_96
        Me.BunifuImageButton1.ImageActive = Nothing
        Me.BunifuImageButton1.Location = New System.Drawing.Point(27, 44)
        Me.BunifuImageButton1.Name = "BunifuImageButton1"
        Me.BunifuImageButton1.Size = New System.Drawing.Size(39, 36)
        Me.BunifuImageButton1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.BunifuImageButton1.TabIndex = 12
        Me.BunifuImageButton1.TabStop = False
        Me.BunifuImageButton1.Visible = False
        Me.BunifuImageButton1.Zoom = 10
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.White
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.BunifuTransition1.SetDecoration(Me.PictureBox1, BunifuAnimatorNS.DecorationType.None)
        Me.PictureBox1.Location = New System.Drawing.Point(329, 44)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(418, 312)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.PictureBox2)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.BunifuTransition1.SetDecoration(Me.Panel2, BunifuAnimatorNS.DecorationType.None)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(774, 485)
        Me.Panel2.TabIndex = 1
        '
        'PictureBox2
        '
        Me.BunifuTransition1.SetDecoration(Me.PictureBox2, BunifuAnimatorNS.DecorationType.None)
        Me.PictureBox2.Image = Global.ATMOS_V2.My.Resources.Resources.loading_loading_forever
        Me.PictureBox2.Location = New System.Drawing.Point(329, 173)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(100, 100)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 2
        Me.PictureBox2.TabStop = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.BunifuTransition1.SetDecoration(Me.Label6, BunifuAnimatorNS.DecorationType.None)
        Me.Label6.Font = New System.Drawing.Font("Consolas", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label6.Location = New System.Drawing.Point(187, 305)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(370, 24)
        Me.Label6.TabIndex = 1
        Me.Label6.Text = "Loading Face Recognition Model"
        '
        'BackgroundWorker1
        '
        '
        'BunifuTransition1
        '
        Me.BunifuTransition1.AnimationType = BunifuAnimatorNS.AnimationType.ScaleAndRotate
        Me.BunifuTransition1.Cursor = Nothing
        Animation1.AnimateOnlyDifferences = True
        Animation1.BlindCoeff = CType(resources.GetObject("Animation1.BlindCoeff"), System.Drawing.PointF)
        Animation1.LeafCoeff = 0!
        Animation1.MaxTime = 1.0!
        Animation1.MinTime = 0!
        Animation1.MosaicCoeff = CType(resources.GetObject("Animation1.MosaicCoeff"), System.Drawing.PointF)
        Animation1.MosaicShift = CType(resources.GetObject("Animation1.MosaicShift"), System.Drawing.PointF)
        Animation1.MosaicSize = 0
        Animation1.Padding = New System.Windows.Forms.Padding(30)
        Animation1.RotateCoeff = 0.5!
        Animation1.RotateLimit = 0.2!
        Animation1.ScaleCoeff = CType(resources.GetObject("Animation1.ScaleCoeff"), System.Drawing.PointF)
        Animation1.SlideCoeff = CType(resources.GetObject("Animation1.SlideCoeff"), System.Drawing.PointF)
        Animation1.TimeCoeff = 0!
        Animation1.TransparencyCoeff = 0!
        Me.BunifuTransition1.DefaultAnimation = Animation1
        '
        'BackgroundWorker2
        '
        '
        'FaceRecog
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(774, 485)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel2)
        Me.BunifuTransition1.SetDecoration(Me, BunifuAnimatorNS.DecorationType.None)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "FaceRecog"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FaceRecog"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        CType(Me.BunifuCustomDataGrid4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BunifuImageButton1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As Panel
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents BunifuImageButton1 As ns1.BunifuImageButton
    Friend WithEvents BunifuThinButton22 As ns1.BunifuThinButton2
    Friend WithEvents BunifuThinButton21 As ns1.BunifuThinButton2
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Label6 As Label
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents PictureBox2 As PictureBox
    Friend WithEvents BunifuTransition1 As BunifuAnimatorNS.BunifuTransition
    Friend WithEvents Panel3 As Panel
    Friend WithEvents BunifuThinButton219 As ns1.BunifuThinButton2
    Friend WithEvents BunifuThinButton220 As ns1.BunifuThinButton2
    Friend WithEvents BunifuCustomDataGrid4 As ns1.BunifuCustomDataGrid
    Friend WithEvents IDLB As Label
    Friend WithEvents NameLB As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents PictureBox3 As PictureBox
    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label3 As Label
    Friend WithEvents BunifuThinButton23 As ns1.BunifuThinButton2
End Class
