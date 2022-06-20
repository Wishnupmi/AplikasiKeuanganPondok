<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form4
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
        Dim Label10 As System.Windows.Forms.Label
        Dim Label1 As System.Windows.Forms.Label
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form4))
        Dim Label12 As System.Windows.Forms.Label
        Dim Label11 As System.Windows.Forms.Label
        Dim Label2 As System.Windows.Forms.Label
        Dim AsalLabel As System.Windows.Forms.Label
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim Label9 As System.Windows.Forms.Label
        Dim Tgl_diterimaLabel As System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cmbJnsSurat = New System.Windows.Forms.ComboBox()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.txtNmrSurat = New System.Windows.Forms.TextBox()
        Me.txtAsalSurat = New System.Windows.Forms.TextBox()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.txtTgl_terima = New System.Windows.Forms.DateTimePicker()
        Label10 = New System.Windows.Forms.Label()
        Label1 = New System.Windows.Forms.Label()
        Label12 = New System.Windows.Forms.Label()
        Label11 = New System.Windows.Forms.Label()
        Label2 = New System.Windows.Forms.Label()
        AsalLabel = New System.Windows.Forms.Label()
        Label9 = New System.Windows.Forms.Label()
        Tgl_diterimaLabel = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Label9)
        Me.GroupBox1.Controls.Add(Me.txtTgl_terima)
        Me.GroupBox1.Controls.Add(Tgl_diterimaLabel)
        Me.GroupBox1.Controls.Add(Label10)
        Me.GroupBox1.Controls.Add(Me.cmbJnsSurat)
        Me.GroupBox1.Controls.Add(Label1)
        Me.GroupBox1.Controls.Add(Me.btnUpdate)
        Me.GroupBox1.Controls.Add(Me.btnDelete)
        Me.GroupBox1.Controls.Add(Me.btnCancel)
        Me.GroupBox1.Controls.Add(Me.btnSave)
        Me.GroupBox1.Controls.Add(Label12)
        Me.GroupBox1.Controls.Add(Label11)
        Me.GroupBox1.Controls.Add(Me.txtNmrSurat)
        Me.GroupBox1.Controls.Add(Label2)
        Me.GroupBox1.Controls.Add(AsalLabel)
        Me.GroupBox1.Controls.Add(Me.txtAsalSurat)
        Me.GroupBox1.Location = New System.Drawing.Point(13, 14)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox1.Size = New System.Drawing.Size(1160, 324)
        Me.GroupBox1.TabIndex = 47
        Me.GroupBox1.TabStop = False
        '
        'Label10
        '
        Label10.AutoSize = True
        Label10.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Label10.Location = New System.Drawing.Point(146, 149)
        Label10.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Label10.Name = "Label10"
        Label10.Size = New System.Drawing.Size(15, 21)
        Label10.TabIndex = 76
        Label10.Text = ":"
        '
        'cmbJnsSurat
        '
        Me.cmbJnsSurat.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.cmbJnsSurat.FormattingEnabled = True
        Me.cmbJnsSurat.Items.AddRange(New Object() {"Email", "Fax", "Biasa", "Asli", "Tembusan"})
        Me.cmbJnsSurat.Location = New System.Drawing.Point(170, 145)
        Me.cmbJnsSurat.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmbJnsSurat.Name = "cmbJnsSurat"
        Me.cmbJnsSurat.Size = New System.Drawing.Size(254, 28)
        Me.cmbJnsSurat.TabIndex = 74
        '
        'Label1
        '
        Label1.AutoSize = True
        Label1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Label1.Location = New System.Drawing.Point(7, 145)
        Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Label1.Name = "Label1"
        Label1.Size = New System.Drawing.Size(133, 21)
        Label1.TabIndex = 75
        Label1.Text = "Kelompok Akun"
        '
        'btnUpdate
        '
        Me.btnUpdate.Image = CType(resources.GetObject("btnUpdate.Image"), System.Drawing.Image)
        Me.btnUpdate.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnUpdate.Location = New System.Drawing.Point(169, 243)
        Me.btnUpdate.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(135, 49)
        Me.btnUpdate.TabIndex = 10
        Me.btnUpdate.Text = "       Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Image = CType(resources.GetObject("btnDelete.Image"), System.Drawing.Image)
        Me.btnDelete.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDelete.Location = New System.Drawing.Point(469, 243)
        Me.btnDelete.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(142, 49)
        Me.btnDelete.TabIndex = 12
        Me.btnDelete.Text = "        &Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Image = CType(resources.GetObject("btnCancel.Image"), System.Drawing.Image)
        Me.btnCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCancel.Location = New System.Drawing.Point(313, 243)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(147, 49)
        Me.btnCancel.TabIndex = 11
        Me.btnCancel.Text = "        &Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Image = CType(resources.GetObject("btnSave.Image"), System.Drawing.Image)
        Me.btnSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnSave.Location = New System.Drawing.Point(170, 243)
        Me.btnSave.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(135, 49)
        Me.btnSave.TabIndex = 9
        Me.btnSave.Text = "         &Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'Label12
        '
        Label12.AutoSize = True
        Label12.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Label12.Location = New System.Drawing.Point(147, 114)
        Label12.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Label12.Name = "Label12"
        Label12.Size = New System.Drawing.Size(15, 21)
        Label12.TabIndex = 65
        Label12.Text = ":"
        '
        'Label11
        '
        Label11.AutoSize = True
        Label11.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Label11.Location = New System.Drawing.Point(146, 73)
        Label11.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Label11.Name = "Label11"
        Label11.Size = New System.Drawing.Size(15, 21)
        Label11.TabIndex = 64
        Label11.Text = ":"
        '
        'txtNmrSurat
        '
        Me.txtNmrSurat.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtNmrSurat.Location = New System.Drawing.Point(170, 109)
        Me.txtNmrSurat.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtNmrSurat.Name = "txtNmrSurat"
        Me.txtNmrSurat.Size = New System.Drawing.Size(254, 26)
        Me.txtNmrSurat.TabIndex = 3
        '
        'Label2
        '
        Label2.AutoSize = True
        Label2.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Label2.Location = New System.Drawing.Point(7, 111)
        Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Label2.Name = "Label2"
        Label2.Size = New System.Drawing.Size(101, 21)
        Label2.TabIndex = 47
        Label2.Text = "Nama Akun"
        '
        'AsalLabel
        '
        AsalLabel.AutoSize = True
        AsalLabel.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        AsalLabel.Location = New System.Drawing.Point(7, 73)
        AsalLabel.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        AsalLabel.Name = "AsalLabel"
        AsalLabel.Size = New System.Drawing.Size(107, 21)
        AsalLabel.TabIndex = 43
        AsalLabel.Text = "Nomor Akun"
        '
        'txtAsalSurat
        '
        Me.txtAsalSurat.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtAsalSurat.Location = New System.Drawing.Point(170, 73)
        Me.txtAsalSurat.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtAsalSurat.Name = "txtAsalSurat"
        Me.txtAsalSurat.Size = New System.Drawing.Size(254, 26)
        Me.txtAsalSurat.TabIndex = 2
        '
        'DataGridView2
        '
        Me.DataGridView2.BorderStyle = System.Windows.Forms.BorderStyle.None
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.Red
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView2.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Location = New System.Drawing.Point(13, 348)
        Me.DataGridView2.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.ReadOnly = True
        Me.DataGridView2.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView2.Size = New System.Drawing.Size(1156, 242)
        Me.DataGridView2.TabIndex = 46
        '
        'Label9
        '
        Label9.AutoSize = True
        Label9.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Label9.Location = New System.Drawing.Point(146, 40)
        Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Label9.Name = "Label9"
        Label9.Size = New System.Drawing.Size(15, 21)
        Label9.TabIndex = 79
        Label9.Text = ":"
        '
        'txtTgl_terima
        '
        Me.txtTgl_terima.CustomFormat = "dd/MM/yyyy"
        Me.txtTgl_terima.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtTgl_terima.Location = New System.Drawing.Point(170, 37)
        Me.txtTgl_terima.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtTgl_terima.Name = "txtTgl_terima"
        Me.txtTgl_terima.Size = New System.Drawing.Size(254, 26)
        Me.txtTgl_terima.TabIndex = 77
        '
        'Tgl_diterimaLabel
        '
        Tgl_diterimaLabel.AutoSize = True
        Tgl_diterimaLabel.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Tgl_diterimaLabel.Location = New System.Drawing.Point(5, 40)
        Tgl_diterimaLabel.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Tgl_diterimaLabel.Name = "Tgl_diterimaLabel"
        Tgl_diterimaLabel.Size = New System.Drawing.Size(107, 21)
        Tgl_diterimaLabel.TabIndex = 78
        Tgl_diterimaLabel.Text = "Tgl Diterima"
        '
        'Form4
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1180, 604)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.DataGridView2)
        Me.Name = "Form4"
        Me.Text = "FORM INPUT TRANSAKSI"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtTgl_terima As System.Windows.Forms.DateTimePicker
    Friend WithEvents cmbJnsSurat As System.Windows.Forms.ComboBox
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents txtNmrSurat As System.Windows.Forms.TextBox
    Friend WithEvents txtAsalSurat As System.Windows.Forms.TextBox
    Friend WithEvents DataGridView2 As System.Windows.Forms.DataGridView
End Class
