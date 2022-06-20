<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MDIParent1
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
        Me.components = New System.ComponentModel.Container()
        Me.MenuStrip = New System.Windows.Forms.MenuStrip()
        Me.FileMenu = New System.Windows.Forms.ToolStripMenuItem()
        Me.DataUser = New System.Windows.Forms.ToolStripMenuItem()
        Me.DataSantri = New System.Windows.Forms.ToolStripMenuItem()
        Me.DataPerkiraanAkun = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Transaksi = New System.Windows.Forms.ToolStripMenuItem()
        Me.TransaksiKasMasuk = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProsesAkhirBulan = New System.Windows.Forms.ToolStripMenuItem()
        Me.Laporan = New System.Windows.Forms.ToolStripMenuItem()
        Me.LaporanKasMasuk = New System.Windows.Forms.ToolStripMenuItem()
        Me.LaporanKasKeluar = New System.Windows.Forms.ToolStripMenuItem()
        Me.LaporanPosisiKeuangan = New System.Windows.Forms.ToolStripMenuItem()
        Me.LaporanAktivitas = New System.Windows.Forms.ToolStripMenuItem()
        Me.LaporanArusKas = New System.Windows.Forms.ToolStripMenuItem()
        Me.StatusStrip = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.MenuStrip.SuspendLayout()
        Me.StatusStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip
        '
        Me.MenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileMenu, Me.Transaksi, Me.Laporan})
        Me.MenuStrip.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip.Name = "MenuStrip"
        Me.MenuStrip.Size = New System.Drawing.Size(632, 24)
        Me.MenuStrip.TabIndex = 5
        Me.MenuStrip.Text = "MenuStrip"
        '
        'FileMenu
        '
        Me.FileMenu.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.DataUser, Me.DataSantri, Me.DataPerkiraanAkun, Me.ToolStripSeparator5, Me.ExitToolStripMenuItem})
        Me.FileMenu.ImageTransparentColor = System.Drawing.SystemColors.ActiveBorder
        Me.FileMenu.Name = "FileMenu"
        Me.FileMenu.Size = New System.Drawing.Size(82, 20)
        Me.FileMenu.Text = "&Data Master"
        '
        'DataUser
        '
        Me.DataUser.Name = "DataUser"
        Me.DataUser.Size = New System.Drawing.Size(181, 22)
        Me.DataUser.Text = "Data User"
        '
        'DataSantri
        '
        Me.DataSantri.Name = "DataSantri"
        Me.DataSantri.Size = New System.Drawing.Size(181, 22)
        Me.DataSantri.Text = "Data Santri"
        '
        'DataPerkiraanAkun
        '
        Me.DataPerkiraanAkun.Name = "DataPerkiraanAkun"
        Me.DataPerkiraanAkun.Size = New System.Drawing.Size(181, 22)
        Me.DataPerkiraanAkun.Text = "Data Perkiraan Akun"
        '
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(178, 6)
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(181, 22)
        Me.ExitToolStripMenuItem.Text = "E&xit"
        '
        'Transaksi
        '
        Me.Transaksi.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TransaksiKasMasuk, Me.ProsesAkhirBulan})
        Me.Transaksi.Name = "Transaksi"
        Me.Transaksi.Size = New System.Drawing.Size(66, 20)
        Me.Transaksi.Text = "&Transaksi"
        '
        'TransaksiKasMasuk
        '
        Me.TransaksiKasMasuk.Name = "TransaksiKasMasuk"
        Me.TransaksiKasMasuk.Size = New System.Drawing.Size(172, 22)
        Me.TransaksiKasMasuk.Text = "Input Transaksi"
        '
        'ProsesAkhirBulan
        '
        Me.ProsesAkhirBulan.Name = "ProsesAkhirBulan"
        Me.ProsesAkhirBulan.Size = New System.Drawing.Size(172, 22)
        Me.ProsesAkhirBulan.Text = "Proses Akhir Bulan"
        '
        'Laporan
        '
        Me.Laporan.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.LaporanKasMasuk, Me.LaporanKasKeluar, Me.ToolStripSeparator1, Me.LaporanPosisiKeuangan, Me.ToolStripSeparator2, Me.LaporanAktivitas, Me.ToolStripSeparator3, Me.LaporanArusKas})
        Me.Laporan.Name = "Laporan"
        Me.Laporan.Size = New System.Drawing.Size(62, 20)
        Me.Laporan.Text = "&Laporan"
        '
        'LaporanKasMasuk
        '
        Me.LaporanKasMasuk.Name = "LaporanKasMasuk"
        Me.LaporanKasMasuk.Size = New System.Drawing.Size(206, 22)
        Me.LaporanKasMasuk.Text = "Laporan Kas Masuk"
        '
        'LaporanKasKeluar
        '
        Me.LaporanKasKeluar.Name = "LaporanKasKeluar"
        Me.LaporanKasKeluar.Size = New System.Drawing.Size(206, 22)
        Me.LaporanKasKeluar.Text = "Laporan Kas Keluar"
        '
        'LaporanPosisiKeuangan
        '
        Me.LaporanPosisiKeuangan.Name = "LaporanPosisiKeuangan"
        Me.LaporanPosisiKeuangan.Size = New System.Drawing.Size(206, 22)
        Me.LaporanPosisiKeuangan.Text = "Laporan Posisi Keuangan"
        '
        'LaporanAktivitas
        '
        Me.LaporanAktivitas.Name = "LaporanAktivitas"
        Me.LaporanAktivitas.Size = New System.Drawing.Size(206, 22)
        Me.LaporanAktivitas.Text = "Laporan Aktivitas"
        '
        'LaporanArusKas
        '
        Me.LaporanArusKas.Name = "LaporanArusKas"
        Me.LaporanArusKas.Size = New System.Drawing.Size(206, 22)
        Me.LaporanArusKas.Text = "Laporan Arus Kas"
        '
        'StatusStrip
        '
        Me.StatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel})
        Me.StatusStrip.Location = New System.Drawing.Point(0, 431)
        Me.StatusStrip.Name = "StatusStrip"
        Me.StatusStrip.Size = New System.Drawing.Size(632, 22)
        Me.StatusStrip.TabIndex = 7
        Me.StatusStrip.Text = "StatusStrip"
        '
        'ToolStripStatusLabel
        '
        Me.ToolStripStatusLabel.Name = "ToolStripStatusLabel"
        Me.ToolStripStatusLabel.Size = New System.Drawing.Size(39, 17)
        Me.ToolStripStatusLabel.Text = "Status"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(203, 6)
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(203, 6)
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(203, 6)
        '
        'MDIParent1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(632, 453)
        Me.Controls.Add(Me.MenuStrip)
        Me.Controls.Add(Me.StatusStrip)
        Me.IsMdiContainer = True
        Me.MainMenuStrip = Me.MenuStrip
        Me.Name = "MDIParent1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Aplikasi Laporan Keuangan Pondok Miftahul Huda"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MenuStrip.ResumeLayout(False)
        Me.MenuStrip.PerformLayout()
        Me.StatusStrip.ResumeLayout(False)
        Me.StatusStrip.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ToolTip As System.Windows.Forms.ToolTip
    Friend WithEvents ToolStripStatusLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip As System.Windows.Forms.StatusStrip
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator5 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents FileMenu As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MenuStrip As System.Windows.Forms.MenuStrip
    Friend WithEvents DataSantri As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DataPerkiraanAkun As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Transaksi As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Laporan As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TransaksiKasMasuk As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LaporanKasMasuk As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LaporanKasKeluar As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LaporanPosisiKeuangan As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LaporanAktivitas As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DataUser As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ProsesAkhirBulan As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LaporanArusKas As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripSeparator3 As System.Windows.Forms.ToolStripSeparator

End Class
