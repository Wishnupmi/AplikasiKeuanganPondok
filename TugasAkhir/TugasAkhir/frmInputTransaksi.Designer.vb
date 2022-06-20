<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmInputTransaksi
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim AsalLabel As System.Windows.Forms.Label
        Dim Label2 As System.Windows.Forms.Label
        Dim Label11 As System.Windows.Forms.Label
        Dim Label12 As System.Windows.Forms.Label
        Dim Label1 As System.Windows.Forms.Label
        Dim Label10 As System.Windows.Forms.Label
        Dim Tgl_diterima As System.Windows.Forms.Label
        Dim Label9 As System.Windows.Forms.Label
        Dim Label3 As System.Windows.Forms.Label
        Dim Label14 As System.Windows.Forms.Label
        Dim Label13 As System.Windows.Forms.Label
        Dim Label4 As System.Windows.Forms.Label
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmInputTransaksi))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.txtKodeTransaksi = New System.Windows.Forms.TextBox()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.cmbJnsTransaksi = New System.Windows.Forms.ComboBox()
        Me.txtTgl_Transaksi = New System.Windows.Forms.DateTimePicker()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.txtKeterangan = New System.Windows.Forms.TextBox()
        Me.txtId = New System.Windows.Forms.TextBox()
        Me.txtNominal = New System.Windows.Forms.TextBox()
        Me.txtNmAkun = New System.Windows.Forms.TextBox()
        Me.txtNoAkun = New System.Windows.Forms.TextBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.txtSaldo = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.txtDebetKredit = New System.Windows.Forms.TextBox()
        AsalLabel = New System.Windows.Forms.Label()
        Label2 = New System.Windows.Forms.Label()
        Label11 = New System.Windows.Forms.Label()
        Label12 = New System.Windows.Forms.Label()
        Label1 = New System.Windows.Forms.Label()
        Label10 = New System.Windows.Forms.Label()
        Tgl_diterima = New System.Windows.Forms.Label()
        Label9 = New System.Windows.Forms.Label()
        Label3 = New System.Windows.Forms.Label()
        Label14 = New System.Windows.Forms.Label()
        Label13 = New System.Windows.Forms.Label()
        Label4 = New System.Windows.Forms.Label()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'AsalLabel
        '
        AsalLabel.AutoSize = True
        AsalLabel.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        AsalLabel.Location = New System.Drawing.Point(19, 60)
        AsalLabel.Name = "AsalLabel"
        AsalLabel.Size = New System.Drawing.Size(94, 15)
        AsalLabel.TabIndex = 43
        AsalLabel.Text = "Kode Transaksi"
        '
        'Label2
        '
        Label2.AutoSize = True
        Label2.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Label2.Location = New System.Drawing.Point(19, 156)
        Label2.Name = "Label2"
        Label2.Size = New System.Drawing.Size(71, 15)
        Label2.TabIndex = 47
        Label2.Text = "Keterangan"
        '
        'Label11
        '
        Label11.AutoSize = True
        Label11.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Label11.Location = New System.Drawing.Point(147, 156)
        Label11.Name = "Label11"
        Label11.Size = New System.Drawing.Size(10, 15)
        Label11.TabIndex = 64
        Label11.Text = ":"
        '
        'Label12
        '
        Label12.AutoSize = True
        Label12.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Label12.Location = New System.Drawing.Point(147, 61)
        Label12.Name = "Label12"
        Label12.Size = New System.Drawing.Size(10, 15)
        Label12.TabIndex = 65
        Label12.Text = ":"
        '
        'Label1
        '
        Label1.AutoSize = True
        Label1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Label1.Location = New System.Drawing.Point(19, 82)
        Label1.Name = "Label1"
        Label1.Size = New System.Drawing.Size(95, 15)
        Label1.TabIndex = 75
        Label1.Text = "Jenis Transaksi"
        '
        'Label10
        '
        Label10.AutoSize = True
        Label10.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Label10.Location = New System.Drawing.Point(147, 83)
        Label10.Name = "Label10"
        Label10.Size = New System.Drawing.Size(10, 15)
        Label10.TabIndex = 76
        Label10.Text = ":"
        '
        'Tgl_diterima
        '
        Tgl_diterima.AutoSize = True
        Tgl_diterima.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Tgl_diterima.Location = New System.Drawing.Point(17, 25)
        Tgl_diterima.Name = "Tgl_diterima"
        Tgl_diterima.Size = New System.Drawing.Size(109, 15)
        Tgl_diterima.TabIndex = 78
        Tgl_diterima.Text = "Tanggal Transaksi"
        '
        'Label9
        '
        Label9.AutoSize = True
        Label9.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Label9.Location = New System.Drawing.Point(146, 25)
        Label9.Name = "Label9"
        Label9.Size = New System.Drawing.Size(10, 15)
        Label9.TabIndex = 79
        Label9.Text = ":"
        '
        'Label3
        '
        Label3.AutoSize = True
        Label3.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Label3.Location = New System.Drawing.Point(147, 134)
        Label3.Name = "Label3"
        Label3.Size = New System.Drawing.Size(10, 15)
        Label3.TabIndex = 85
        Label3.Text = ":"
        '
        'Label14
        '
        Label14.AutoSize = True
        Label14.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Label14.Location = New System.Drawing.Point(19, 107)
        Label14.Name = "Label14"
        Label14.Size = New System.Drawing.Size(74, 15)
        Label14.TabIndex = 91
        Label14.Text = "Nomor Akun"
        '
        'Label13
        '
        Label13.AutoSize = True
        Label13.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Label13.Location = New System.Drawing.Point(147, 107)
        Label13.Name = "Label13"
        Label13.Size = New System.Drawing.Size(10, 15)
        Label13.TabIndex = 92
        Label13.Text = ":"
        '
        'Label4
        '
        Label4.AutoSize = True
        Label4.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Label4.Location = New System.Drawing.Point(19, 134)
        Label4.Name = "Label4"
        Label4.Size = New System.Drawing.Size(54, 15)
        Label4.TabIndex = 95
        Label4.Text = "Nominal"
        '
        'cmdExit
        '
        Me.cmdExit.Image = CType(resources.GetObject("cmdExit.Image"), System.Drawing.Image)
        Me.cmdExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdExit.Location = New System.Drawing.Point(743, 425)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(90, 32)
        Me.cmdExit.TabIndex = 9
        Me.cmdExit.Text = "     &Keluar"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(9, 424)
        Me.Button1.Margin = New System.Windows.Forms.Padding(2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(85, 25)
        Me.Button1.TabIndex = 8
        Me.Button1.Text = "Cetak"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'DataGridView1
        '
        Me.DataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.None
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.Red
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(9, 269)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(825, 149)
        Me.DataGridView1.TabIndex = 89
        '
        'txtKodeTransaksi
        '
        Me.txtKodeTransaksi.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtKodeTransaksi.Enabled = False
        Me.txtKodeTransaksi.Location = New System.Drawing.Point(163, 58)
        Me.txtKodeTransaksi.Name = "txtKodeTransaksi"
        Me.txtKodeTransaksi.Size = New System.Drawing.Size(102, 20)
        Me.txtKodeTransaksi.TabIndex = 3
        Me.txtKodeTransaksi.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btnCancel
        '
        Me.btnCancel.Image = CType(resources.GetObject("btnCancel.Image"), System.Drawing.Image)
        Me.btnCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCancel.Location = New System.Drawing.Point(338, 221)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(90, 31)
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "        &Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Image = CType(resources.GetObject("btnDelete.Image"), System.Drawing.Image)
        Me.btnDelete.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDelete.Location = New System.Drawing.Point(246, 221)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(87, 31)
        Me.btnDelete.TabIndex = 7
        Me.btnDelete.Text = "        &Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Image = CType(resources.GetObject("btnUpdate.Image"), System.Drawing.Image)
        Me.btnUpdate.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnUpdate.Location = New System.Drawing.Point(159, 221)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(81, 31)
        Me.btnUpdate.TabIndex = 5
        Me.btnUpdate.Text = "       Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        Me.btnUpdate.Visible = False
        '
        'cmbJnsTransaksi
        '
        Me.cmbJnsTransaksi.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.cmbJnsTransaksi.FormattingEnabled = True
        Me.cmbJnsTransaksi.Items.AddRange(New Object() {"Transaksi Kas Masuk", "Transaksi Kas Keluar"})
        Me.cmbJnsTransaksi.Location = New System.Drawing.Point(163, 82)
        Me.cmbJnsTransaksi.Name = "cmbJnsTransaksi"
        Me.cmbJnsTransaksi.Size = New System.Drawing.Size(171, 21)
        Me.cmbJnsTransaksi.TabIndex = 1
        '
        'txtTgl_Transaksi
        '
        Me.txtTgl_Transaksi.CustomFormat = "dd/MM/yyyy"
        Me.txtTgl_Transaksi.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtTgl_Transaksi.Location = New System.Drawing.Point(161, 25)
        Me.txtTgl_Transaksi.Name = "txtTgl_Transaksi"
        Me.txtTgl_Transaksi.Size = New System.Drawing.Size(171, 20)
        Me.txtTgl_Transaksi.TabIndex = 0
        '
        'btnSave
        '
        Me.btnSave.Image = CType(resources.GetObject("btnSave.Image"), System.Drawing.Image)
        Me.btnSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnSave.Location = New System.Drawing.Point(135, 221)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(81, 31)
        Me.btnSave.TabIndex = 4
        Me.btnSave.Text = "         &Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'txtKeterangan
        '
        Me.txtKeterangan.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtKeterangan.Location = New System.Drawing.Point(163, 156)
        Me.txtKeterangan.Multiline = True
        Me.txtKeterangan.Name = "txtKeterangan"
        Me.txtKeterangan.Size = New System.Drawing.Size(331, 48)
        Me.txtKeterangan.TabIndex = 2
        '
        'txtId
        '
        Me.txtId.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtId.Location = New System.Drawing.Point(338, 25)
        Me.txtId.Name = "txtId"
        Me.txtId.Size = New System.Drawing.Size(54, 20)
        Me.txtId.TabIndex = 81
        Me.txtId.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtId.Visible = False
        '
        'txtNominal
        '
        Me.txtNominal.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtNominal.Location = New System.Drawing.Point(163, 133)
        Me.txtNominal.Name = "txtNominal"
        Me.txtNominal.Size = New System.Drawing.Size(102, 20)
        Me.txtNominal.TabIndex = 83
        Me.txtNominal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtNmAkun
        '
        Me.txtNmAkun.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtNmAkun.Enabled = False
        Me.txtNmAkun.Location = New System.Drawing.Point(262, 107)
        Me.txtNmAkun.Name = "txtNmAkun"
        Me.txtNmAkun.Size = New System.Drawing.Size(224, 20)
        Me.txtNmAkun.TabIndex = 88
        '
        'txtNoAkun
        '
        Me.txtNoAkun.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtNoAkun.Location = New System.Drawing.Point(163, 107)
        Me.txtNoAkun.Name = "txtNoAkun"
        Me.txtNoAkun.Size = New System.Drawing.Size(54, 20)
        Me.txtNoAkun.TabIndex = 93
        Me.txtNoAkun.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(223, 107)
        Me.Button2.Margin = New System.Windows.Forms.Padding(2)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(34, 21)
        Me.Button2.TabIndex = 94
        Me.Button2.Text = "..."
        Me.Button2.UseVisualStyleBackColor = True
        '
        'txtSaldo
        '
        Me.txtSaldo.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtSaldo.Enabled = False
        Me.txtSaldo.Location = New System.Drawing.Point(269, 133)
        Me.txtSaldo.Name = "txtSaldo"
        Me.txtSaldo.Size = New System.Drawing.Size(102, 20)
        Me.txtSaldo.TabIndex = 96
        Me.txtSaldo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtSaldo.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Button3)
        Me.GroupBox1.Controls.Add(Me.txtDebetKredit)
        Me.GroupBox1.Controls.Add(Me.txtSaldo)
        Me.GroupBox1.Controls.Add(Label4)
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.txtNoAkun)
        Me.GroupBox1.Controls.Add(Label13)
        Me.GroupBox1.Controls.Add(Label14)
        Me.GroupBox1.Controls.Add(Me.txtNmAkun)
        Me.GroupBox1.Controls.Add(Label3)
        Me.GroupBox1.Controls.Add(Me.txtNominal)
        Me.GroupBox1.Controls.Add(Me.txtId)
        Me.GroupBox1.Controls.Add(Me.txtKeterangan)
        Me.GroupBox1.Controls.Add(Me.btnSave)
        Me.GroupBox1.Controls.Add(Label9)
        Me.GroupBox1.Controls.Add(Me.txtTgl_Transaksi)
        Me.GroupBox1.Controls.Add(Tgl_diterima)
        Me.GroupBox1.Controls.Add(Label10)
        Me.GroupBox1.Controls.Add(Me.cmbJnsTransaksi)
        Me.GroupBox1.Controls.Add(Label1)
        Me.GroupBox1.Controls.Add(Me.btnUpdate)
        Me.GroupBox1.Controls.Add(Me.btnDelete)
        Me.GroupBox1.Controls.Add(Me.btnCancel)
        Me.GroupBox1.Controls.Add(Label12)
        Me.GroupBox1.Controls.Add(Label11)
        Me.GroupBox1.Controls.Add(Label2)
        Me.GroupBox1.Controls.Add(AsalLabel)
        Me.GroupBox1.Controls.Add(Me.txtKodeTransaksi)
        Me.GroupBox1.Location = New System.Drawing.Point(9, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(825, 263)
        Me.GroupBox1.TabIndex = 47
        Me.GroupBox1.TabStop = False
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(13, 198)
        Me.Button3.Margin = New System.Windows.Forms.Padding(2)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(50, 25)
        Me.Button3.TabIndex = 90
        Me.Button3.Text = "Button3"
        Me.Button3.UseVisualStyleBackColor = True
        Me.Button3.Visible = False
        '
        'txtDebetKredit
        '
        Me.txtDebetKredit.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtDebetKredit.Enabled = False
        Me.txtDebetKredit.Location = New System.Drawing.Point(375, 133)
        Me.txtDebetKredit.Name = "txtDebetKredit"
        Me.txtDebetKredit.Size = New System.Drawing.Size(102, 20)
        Me.txtDebetKredit.TabIndex = 97
        Me.txtDebetKredit.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtDebetKredit.Visible = False
        '
        'frmInputTransaksi
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(843, 469)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "frmInputTransaksi"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FORM INPUT TRANSAKSI"
        Me.TopMost = True
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents txtKodeTransaksi As System.Windows.Forms.TextBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents cmbJnsTransaksi As System.Windows.Forms.ComboBox
    Friend WithEvents txtTgl_Transaksi As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents txtKeterangan As System.Windows.Forms.TextBox
    Friend WithEvents txtId As System.Windows.Forms.TextBox
    Friend WithEvents txtNominal As System.Windows.Forms.TextBox
    Friend WithEvents txtNmAkun As System.Windows.Forms.TextBox
    Friend WithEvents txtNoAkun As System.Windows.Forms.TextBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents txtSaldo As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtDebetKredit As System.Windows.Forms.TextBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
End Class
