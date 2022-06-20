Imports System.Data.SqlClient
Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared


Public Class frmInputTransaksi
    Dim koneksi = Module2.Koneksi
    Dim con As New OleDbConnection
    'Dim ds As New DataSet
    Dim da As OleDbDataAdapter
    Dim dispoBaris As Integer
    Public nmr_urut As String
    Dim bln_romawi As String
    Dim Sql As String = String.Empty
    Dim Sql2 As String = String.Empty
    Dim Sql3 As String = String.Empty
    Dim Sql4 As String = String.Empty
    Dim Sql5 As String = String.Empty
    Dim Sql6 As String = String.Empty
    Dim Kd_Akun, Nm_Akun As String
    Dim trans As String
    Dim urutan As String
    Dim hitung As Long
    Dim bln, thn As String
    Dim TG As Date

    Private Property cmd2 As SqlCommand

    Sub reset()
        cmbJnsTransaksi.Text = ""
        txtNmAkun.Text = ""
        txtKodeTransaksi.Text = ""
        txtNominal.Text = ""
        txtDebetKredit.Text = ""
        txtKeterangan.Text = ""
        txtId.Text = ""
        txtNoAkun.Text = ""
        txtSaldo.Text = ""

        Kodenomor()

        btnSave.Visible = True
        btnUpdate.Visible = False
        cmbJnsTransaksi.Enabled = True
        txtNoAkun.Enabled = True
        Button2.Enabled = True
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        'MsgBox(trans)
        If txtNominal.Text = "0" Or txtNominal.Text = "" Then
            MsgBox("Nominal Tidak Boleh Kosong atau Nol, Silakan Lengkapi lagi.")
        Else



            tutupkoneksi()
            bukakoneksi()
            cmd = New SqlCommand("Select * from Table_Akun WHERE Nomor_Akun='" & Trim(txtNoAkun.Text) & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'", conn)

            rd = cmd.ExecuteReader
            rd.Read()
            If Not rd.HasRows Then
                MsgBox("Nomor Akun Tidak ada di bulan dan tahun yang Anda Pilih. Sesuaikan Tanggal dengan Tanggal Akun yang ada di Database")
            Else



                If ((CStr(rd.Item("Saldo_Debet")) = "0") And (CStr(rd.Item("Saldo_Kredit")) = "0")) And (cmbJnsTransaksi.Text = "Transaksi Kas Masuk") Then
                    'MsgBox("nol nol")

                    Sql3 = "INSERT INTO Data_Transaksi(Tanggal_Input, Jenis_Transaksi, Keterangan, Debit, Nama_Akun, Kode_Transaksi, No_Akun, Kredit, Saldo_Debet, Saldo_Kredit,Saldo_Debet2,Saldo_Kredit2)" & _
                          " VALUES ('" & Format(txtTgl_Transaksi.Value, "MM-dd-yyyy") & "','" & cmbJnsTransaksi.Text & "','" & txtKeterangan.Text & "','" & txtNominal.Text & "','" & txtNmAkun.Text & "','" & txtKodeTransaksi.Text & "','" & txtNoAkun.Text & "','0','" & txtSaldo.Text & "','0','0','" & Val(txtNominal.Text) & "')"

                    ' Periksa hati-hati tanda kutip untuk setiap variabel, salah ketik mengakibatkan query anda tidak akan terbaca.
                    ' Persiapan execusi Query Insert
                    Dim command3 As New SqlCommand(Sql3, Module2.Koneksi)
                    command3.ExecuteNonQuery()
                    Module2.Koneksi.Close()


                    Sql5 = "UPDATE Table_Akun Set Saldo_Kredit='" & Val(txtNominal.Text) & "' Where Nomor_Akun='" & txtNoAkun.Text & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'"

                    Dim command5 As New SqlCommand(Sql5, Module2.Koneksi)
                    command5.ExecuteNonQuery()
                    Module2.Koneksi.Close()

                    Kodenomor()
                    'MsgBox(urutan)

                    'MsgBox(urutan + " " + bln_thn)

                    Sql2 = "UPDATE Table_Nomor Set nomor='" & urutan & "'"

                    Dim command As New SqlCommand(Sql2, Module2.Koneksi)
                    command.ExecuteNonQuery()
                    Module2.Koneksi.Close()

                    MsgBox("Data berhasil tersimpan", vbInformation, "Konfirmasi")
                    Call tampil()
                    Call reset()


                ElseIf ((CStr(rd.Item("Saldo_Debet")) = "0") And (CStr(rd.Item("Saldo_Kredit")) = "0")) And (cmbJnsTransaksi.Text = "Transaksi Kas Keluar") Then
                    'MsgBox("nol nol")

                    Sql3 = "INSERT INTO Data_Transaksi(Tanggal_Input, Jenis_Transaksi, Keterangan, Debit, Nama_Akun, Kode_Transaksi, No_Akun, Kredit, Saldo_Debet, Saldo_Kredit,Saldo_Debet2,Saldo_Kredit2)" & _
                          " VALUES ('" & Format(txtTgl_Transaksi.Value, "MM-dd-yyyy") & "','" & cmbJnsTransaksi.Text & "','" & txtKeterangan.Text & "','0','" & txtNmAkun.Text & "','" & txtKodeTransaksi.Text & "','" & txtNoAkun.Text & "','" & txtNominal.Text & "','0','" & txtSaldo.Text & "','" & Val(txtNominal.Text) & "','0')"

                    ' Periksa hati-hati tanda kutip untuk setiap variabel, salah ketik mengakibatkan query anda tidak akan terbaca.
                    ' Persiapan execusi Query Insert
                    Dim command3 As New SqlCommand(Sql3, Module2.Koneksi)
                    command3.ExecuteNonQuery()
                    Module2.Koneksi.Close()


                    Sql5 = "UPDATE Table_Akun Set Saldo_Debet='" & Val(txtNominal.Text) & "' Where Nomor_Akun='" & txtNoAkun.Text & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'"

                    Dim command5 As New SqlCommand(Sql5, Module2.Koneksi)
                    command5.ExecuteNonQuery()
                    Module2.Koneksi.Close()

                    Kodenomor()
                    'MsgBox(urutan)

                    'MsgBox(urutan + " " + bln_thn)

                    Sql2 = "UPDATE Table_Nomor Set nomor='" & urutan & "'"

                    Dim command As New SqlCommand(Sql2, Module2.Koneksi)
                    command.ExecuteNonQuery()
                    Module2.Koneksi.Close()

                    MsgBox("Data berhasil tersimpan", vbInformation, "Konfirmasi")
                    Call tampil()
                    Call reset()


                ElseIf (CStr(rd.Item("Saldo_Debet")) <> "0") And (CStr(rd.Item("Saldo_Kredit")) = "0") Then
                    If trans = "debit" Then

                        Sql3 = "INSERT INTO Data_Transaksi(Tanggal_Input, Jenis_Transaksi, Keterangan, Debit, Nama_Akun, Kode_Transaksi, No_Akun, Kredit, Saldo_Debet, Saldo_Kredit,Saldo_Debet2,Saldo_Kredit2)" & _
                          " VALUES ('" & Format(txtTgl_Transaksi.Value, "MM-dd-yyyy") & "','" & cmbJnsTransaksi.Text & "','" & txtKeterangan.Text & "','" & txtNominal.Text & "','" & txtNmAkun.Text & "','" & txtKodeTransaksi.Text & "','" & txtNoAkun.Text & "','0','" & txtSaldo.Text & "','0','" & Val(txtSaldo.Text) - Val(txtNominal.Text) & "','0')"

                        ' Periksa hati-hati tanda kutip untuk setiap variabel, salah ketik mengakibatkan query anda tidak akan terbaca.
                        ' Persiapan execusi Query Insert
                        Dim command3 As New SqlCommand(Sql3, Module2.Koneksi)
                        command3.ExecuteNonQuery()
                        Module2.Koneksi.Close()


                        Sql5 = "UPDATE Table_Akun Set Saldo_Debet='" & Val(txtSaldo.Text) - Val(txtNominal.Text) & "' Where Nomor_Akun='" & txtNoAkun.Text & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'"

                        Dim command5 As New SqlCommand(Sql5, Module2.Koneksi)
                        command5.ExecuteNonQuery()
                        Module2.Koneksi.Close()

                        Kodenomor()
                        'MsgBox(urutan)

                        'MsgBox(urutan + " " + bln_thn)

                        Sql2 = "UPDATE Table_Nomor Set nomor='" & urutan & "'"

                        Dim command As New SqlCommand(Sql2, Module2.Koneksi)
                        command.ExecuteNonQuery()
                        Module2.Koneksi.Close()

                        MsgBox("Data berhasil tersimpan", vbInformation, "Konfirmasi")
                        Call tampil()
                        Call reset()

                    Else

                        Sql4 = "INSERT INTO Data_Transaksi(Tanggal_Input, Jenis_Transaksi, Keterangan, Kredit, Nama_Akun, Kode_Transaksi, No_Akun, Debit, Saldo_Kredit, Saldo_Debet,Saldo_Debet2,Saldo_Kredit2)" & _
                          " VALUES ('" & Format(txtTgl_Transaksi.Value, "MM-dd-yyyy") & "','" & cmbJnsTransaksi.Text & "','" & txtKeterangan.Text & "','" & txtNominal.Text & "','" & txtNmAkun.Text & "','" & txtKodeTransaksi.Text & "','" & txtNoAkun.Text & "','0','" & txtSaldo.Text & "','0','" & Val(txtSaldo.Text) + Val(txtNominal.Text) & "','0')"

                        ' Periksa hati-hati tanda kutip untuk setiap variabel, salah ketik mengakibatkan query anda tidak akan terbaca.
                        ' Persiapan execusi Query Insert
                        Dim command4 As New SqlCommand(Sql4, Module2.Koneksi)
                        command4.ExecuteNonQuery()
                        Module2.Koneksi.Close()

                        Sql6 = "UPDATE Table_Akun Set Saldo_Debet='" & Val(txtSaldo.Text) + Val(txtNominal.Text) & "' Where Nomor_Akun='" & txtNoAkun.Text & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'"

                        Dim command6 As New SqlCommand(Sql6, Module2.Koneksi)
                        command6.ExecuteNonQuery()
                        Module2.Koneksi.Close()

                        Kodenomor()
                        'MsgBox(urutan)

                        'MsgBox(urutan + " " + bln_thn)

                        Sql2 = "UPDATE Table_Nomor Set nomor='" & urutan & "'"

                        Dim command As New SqlCommand(Sql2, Module2.Koneksi)
                        command.ExecuteNonQuery()
                        Module2.Koneksi.Close()

                        MsgBox("Data berhasil tersimpan", vbInformation, "Konfirmasi")
                        Call tampil()
                        Call reset()

                    End If

                Else
                    If trans = "debit" Then

                        Sql3 = "INSERT INTO Data_Transaksi(Tanggal_Input, Jenis_Transaksi, Keterangan, Debit, Nama_Akun, Kode_Transaksi, No_Akun, Kredit, Saldo_Debet, Saldo_Kredit,Saldo_Debet2,Saldo_Kredit2)" & _
                          " VALUES ('" & Format(txtTgl_Transaksi.Value, "MM-dd-yyyy") & "','" & cmbJnsTransaksi.Text & "','" & txtKeterangan.Text & "','" & txtNominal.Text & "','" & txtNmAkun.Text & "','" & txtKodeTransaksi.Text & "','" & txtNoAkun.Text & "','0','" & txtSaldo.Text & "','0','0','" & Val(txtSaldo.Text) + Val(txtNominal.Text) & "')"

                        ' Periksa hati-hati tanda kutip untuk setiap variabel, salah ketik mengakibatkan query anda tidak akan terbaca.
                        ' Persiapan execusi Query Insert
                        Dim command3 As New SqlCommand(Sql3, Module2.Koneksi)
                        command3.ExecuteNonQuery()
                        Module2.Koneksi.Close()

                        Sql5 = "UPDATE Table_Akun Set Saldo_Kredit='" & Val(txtSaldo.Text) + Val(txtNominal.Text) & "' Where Nomor_Akun='" & txtNoAkun.Text & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'"

                        Dim command5 As New SqlCommand(Sql5, Module2.Koneksi)
                        command5.ExecuteNonQuery()
                        Module2.Koneksi.Close()

                        MsgBox("Data berhasil tersimpan", vbInformation, "Konfirmasi")
                        Call tampil()
                        Call reset()
                    Else

                        Sql4 = "INSERT INTO Data_Transaksi(Tanggal_Input, Jenis_Transaksi, Keterangan, Kredit, Nama_Akun, Kode_Transaksi, No_Akun, Debit, Saldo_Kredit, Saldo_Debet,Saldo_Debet2,Saldo_Kredit2)" & _
                          " VALUES ('" & Format(txtTgl_Transaksi.Value, "MM-dd-yyyy") & "','" & cmbJnsTransaksi.Text & "','" & txtKeterangan.Text & "','" & txtNominal.Text & "','" & txtNmAkun.Text & "','" & txtKodeTransaksi.Text & "','" & txtNoAkun.Text & "','0','" & txtSaldo.Text & "','0','0','" & Val(txtSaldo.Text) + Val(txtNominal.Text) & "')"

                        ' Periksa hati-hati tanda kutip untuk setiap variabel, salah ketik mengakibatkan query anda tidak akan terbaca.
                        ' Persiapan execusi Query Insert
                        Dim command4 As New SqlCommand(Sql4, Module2.Koneksi)
                        command4.ExecuteNonQuery()
                        Module2.Koneksi.Close()


                        Sql6 = "UPDATE Table_Akun Set Saldo_Kredit='" & Val(txtSaldo.Text) - Val(txtNominal.Text) & "' Where Nomor_Akun='" & txtNoAkun.Text & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'"

                        Dim command6 As New SqlCommand(Sql6, Module2.Koneksi)
                        command6.ExecuteNonQuery()
                        Module2.Koneksi.Close()

                        MsgBox("Data berhasil tersimpan", vbInformation, "Konfirmasi")
                        Call tampil()
                        Call reset()
                    End If
                End If













                'If trans = "debit" Then
                '    Sql3 = "INSERT INTO Data_Transaksi(Tanggal_Input, Jenis_Transaksi, Keterangan, Debit, Nama_Akun, Kode_Transaksi, No_Akun, Kredit, Saldo_Debet, Saldo_Kredit,Saldo_Debet2,Saldo_Kredit2)" & _
                '          " VALUES ('" & Format(txtTgl_Transaksi.Value, "MM-dd-yyyy") & "','" & cmbJnsTransaksi.Text & "','" & txtKeterangan.Text & "','" & txtNominal.Text & "','" & txtNmAkun.Text & "','" & txtKodeTransaksi.Text & "','" & txtNoAkun.Text & "','0','" & txtSaldo.Text & "','0','" & Val(txtSaldo.Text) + Val(txtNominal.Text) & "','0')"

                '    ' Periksa hati-hati tanda kutip untuk setiap variabel, salah ketik mengakibatkan query anda tidak akan terbaca.
                '    ' Persiapan execusi Query Insert
                '    Dim command3 As New SqlCommand(Sql3, Module2.Koneksi)
                '    command3.ExecuteNonQuery()
                '    Module2.Koneksi.Close()


                'Else
                '    Sql4 = "INSERT INTO Data_Transaksi(Tanggal_Input, Jenis_Transaksi, Keterangan, Kredit, Nama_Akun, Kode_Transaksi, No_Akun, Debit, Saldo_Kredit, Saldo_Debet,Saldo_Debet2,Saldo_Kredit2)" & _
                '          " VALUES ('" & Format(txtTgl_Transaksi.Value, "MM-dd-yyyy") & "','" & cmbJnsTransaksi.Text & "','" & txtKeterangan.Text & "','" & txtNominal.Text & "','" & txtNmAkun.Text & "','" & txtKodeTransaksi.Text & "','" & txtNoAkun.Text & "','0','" & txtSaldo.Text & "','0','0','" & Val(txtSaldo.Text) + Val(txtNominal.Text) & "')"

                '    ' Periksa hati-hati tanda kutip untuk setiap variabel, salah ketik mengakibatkan query anda tidak akan terbaca.
                '    ' Persiapan execusi Query Insert
                '    Dim command4 As New SqlCommand(Sql4, Module2.Koneksi)
                '    command4.ExecuteNonQuery()
                '    Module2.Koneksi.Close()

                'End If







            End If


        End If
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Close()
    End Sub

    Private Sub frmInputTransaksi_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        frmCariAkun.Visible = False
        txtNoAkun.Text = frmCariAkun.Kd_Akun
        txtNmAkun.Text = frmCariAkun.Nm_Akun
        'MsgBox(frmCariAkun.Kd_Akun)


    End Sub


    Sub Kodenomor()




        tutupkoneksi()
        bukakoneksi()
        cmd = New SqlCommand("Select nomor from Table_Nomor", conn)

        rd = cmd.ExecuteReader
        rd.Read()
        If Not rd.HasRows Then
            urutan = "001"
        Else
            hitung = Microsoft.VisualBasic.Right(rd.GetString(0), 3) + 1
            urutan = Microsoft.VisualBasic.Right("000" & hitung, 3)
        End If
        bln = Format(txtTgl_Transaksi.Value, "MM")
        thn = Format(txtTgl_Transaksi.Value, "yyyy")
        txtKodeTransaksi.Text = "TR-" + urutan + "-" + bln + thn




    End Sub



    Private Sub frmInputTransaksi_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cmbJnsTransaksi.Focus()
        Kodenomor()
        txtTgl_Transaksi.Value = Format(Now, "dd-MM-yyyy")
        'txtTgl_diteruskan.Value = Format(Now, "dd-MM-yyyy")

        DataGridView1.RowsDefaultCellStyle.BackColor = Color.AliceBlue
        DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.LightSteelBlue

        With DataGridView1
            .Columns.Add("No", "NO.")
            .Columns.Add("Tanggal", "TANGGAL")
            .Columns.Add("JenisTransaksi", "JENIS TRANSAKSI")
            .Columns.Add("NamaAkun", "NAMA AKUN")
            .Columns.Add("Debit", "DEBIT")
            .Columns.Add("Kredit", "KREDIT")
            .Columns.Add("Keterangan", "KETERANGAN")
            .Columns.Add("Id", "Id")
            .Columns.Add("NomorTrans", "NOMOR TRANSAKSI")
            .Columns.Add("NomorAkun", "NOMOR AKUN")
            .Columns.Add("saldebet", "SALDO DEBET")
            .Columns.Add("salkredit", "SALDO KREDIT")
            .Columns.Add("saldebet2", "SALDO DEBET2")
            .Columns.Add("salkredit2", "SALDO KREDIT2")

        End With

        DataGridView1.Columns.Item("No").Width = 50
        DataGridView1.Columns.Item("Tanggal").Width = 120
        DataGridView1.Columns.Item("JenisTransaksi").Width = 150
        DataGridView1.Columns.Item("NamaAkun").Width = 150
        DataGridView1.Columns.Item("Debit").Width = 150
        DataGridView1.Columns.Item("Kredit").Width = 150
        DataGridView1.Columns.Item("Keterangan").Width = 180
        DataGridView1.Columns.Item("Id").Width = 40
        DataGridView1.Columns.Item("NomorTrans").Width = 180
        DataGridView1.Columns.Item("NomorAkun").Width = 180
        DataGridView1.Columns.Item("saldebet").Width = 180
        DataGridView1.Columns.Item("salkredit").Width = 180
        DataGridView1.Columns.Item("saldebet2").Width = 180
        DataGridView1.Columns.Item("salkredit2").Width = 180


        DataGridView1.Columns("No").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Tanggal").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("JenisTransaksi").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("NamaAkun").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Debit").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Kredit").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Keterangan").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Id").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("NomorTrans").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("NomorAkun").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("saldebet").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("salkredit").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("saldebet2").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("salkredit2").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter


        DataGridView1.Columns("No").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Columns("Tanggal").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Columns("JenisTransaksi").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        DataGridView1.Columns("NamaAkun").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        DataGridView1.Columns("Debit").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
        DataGridView1.Columns("Kredit").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
        DataGridView1.Columns("Keterangan").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        DataGridView1.Columns("Id").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Columns("NomorTrans").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Columns("NomorAkun").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Columns("saldebet").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Columns("salkredit").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Columns("saldebet2").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Columns("salkredit2").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter


        'DataGridView1.Columns("Id").Visible = False
        DataGridView1.Columns("NomorAkun").Visible = False
        DataGridView1.Columns("saldebet").Visible = False
        DataGridView1.Columns("salkredit").Visible = False
        DataGridView1.Columns("saldebet2").Visible = False
        DataGridView1.Columns("salkredit2").Visible = False

        Call tampil()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim cryRpt As New ReportDocument
        Dim crTableLogonInfos As New TableLogOnInfos
        Dim crTableLogonInfo As New TableLogOnInfo
        Dim crConnectionInfo As New ConnectionInfo
        Dim crTables As Tables
        Dim crTable As Table

        'cryRpt.Load("E:\TugasAkhir\TugasAkhir\CrystalReport1.rpt")
        'With crConnectionInfo
        '    .ServerName = "(local)"
        '    .DatabaseName = "Tugas_Akhir"
        '    .UserID = "sa"
        '    .Password = "DATAPMI2012"
        'End With
        'crTables = cryRpt.Database.Tables
        'For Each crTable In crTables
        '    crTableLogonInfo = crTable.LogOnInfo
        '    crTableLogonInfo.ConnectionInfo = crConnectionInfo
        '    crTable.ApplyLogOnInfo(crTableLogonInfo)
        'Next
        ''Form2.CrystalReportViewer1.SelectionFormula = "{Disposisi.Tgl_terima} = #" & CDate(Format(DateTimePicker1.Value, "yyyy-MM-dd")) & "#"

        ''LapAktivitas.CrystalReportViewer1.SelectionFormula = "{Disposisi.Tgl_terima} in date ('" & CDate(Format(DateTimePicker1.Value, "yyyy-MM-dd")) & "') TO DateValue ('" & CDate(Format(DateTimePicker2.Value, "yyyy-MM-dd")) & "')"

        ''LaporanPengeluaran.CrystalReportViewer2.SelectionFormula = "{pengeluaran_info.tanggal_pengeluaran} >= #" & CDate(Format(DTTgl1.Value, "yyyy/MM/dd")) & "# and {pengeluaran_info.tanggal_pengeluaran} <= #" & CDate(Format(DTTgl2.Value, "yyyy/MM/dd")) & "#"



        ''Form2.CrystalReportViewer1.SelectionFormula = "{Disposisi.Tgl_terima} = #" & CDate(Format(DateTimePicker1.Value, "yyyy-MM-dd")) & "# "
        'LapAktivitas.CrystalReportViewer1.SelectionFormula = "{Table_Akun.Nomor_Akun} = '" & txtNominal.Text & "'"

        'LapAktivitas.CrystalReportViewer1.ReportSource = cryRpt
        'HasCurrentValue = true;


        'LapAktivitas.CrystalReportViewer1.RefreshReport()
        'LapAktivitas.CrystalReportViewer1.Refresh()

        'LapAktivitas.Show()
        'Me.Cursor = Cursors.Default






        Dim rep As New CrystalReport1 'crJual merupakan nama file CrystalReport

        Dim crPFDs As ParameterFieldDefinitions
        Dim crPFD As ParameterFieldDefinition
        Dim crPVs As New ParameterValues
        Dim crPDV As New ParameterDiscreteValue


        rep.Load("E:\TugasAkhir\TugasAkhir\CrystalReport1.rpt")
        With crConnectionInfo
            .ServerName = "(local)"
            .DatabaseName = "Tugas_Akhir"
            .UserID = "sa"
            .Password = "DATAPMI2012"
        End With
        crTables = rep.Database.Tables
        For Each crTable In crTables
            crTableLogonInfo = crTable.LogOnInfo
            crTableLogonInfo.ConnectionInfo = crConnectionInfo
            crTable.ApplyLogOnInfo(crTableLogonInfo)
        Next










        'LapAktivitas.CrystalReportViewer1.SelectionFormula = "{Nomor_Akun.Tgl_terima} in date ('" & CDate(Format(DateTimePicker1.Value, "yyyy-MM-dd")) & "') TO DateValue ('" & CDate(Format(DateTimePicker2.Value, "yyyy-MM-dd")) & "')"
        LapAktivitas.CrystalReportViewer1.SelectionFormula = "{Table_Akun.Nomor_Akun} = '" & txtKodeTransaksi.Text & "'"
        'LapAktivitas.CrystalReportViewer1.ReportSource = cryRpt


        crPDV.Value = txtKodeTransaksi.Text

        crPFDs = rep.DataDefinition.ParameterFields
        crPFD = crPFDs.Item("Nomor_Akun")
        crPVs.Clear()
        crPVs.Add(crPDV)
        crPFD.ApplyCurrentValues(crPVs)


        LapAktivitas.CrystalReportViewer1.ReportSource = rep
        'LapAktivitas.CrystalReportViewer1.RefreshReport()
        'LapAktivitas.CrystalReportViewer1.Refresh()
        LapAktivitas.Show()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub tampil()
        tutupkoneksi()
        bukakoneksi()
        'NomorAgenda()
        'WHERE Tanggal_Input='" & Format(txtTgl_Transaksi.Value, "yyyy") & "' 
        Sql = "SELECT * From Data_Transaksi Order by Id desc"

        Dim da2 As New SqlDataAdapter(Sql, conn)
        Dim ds2 As New DataSet
        da2.Fill(ds2, "DataTransaksi")
        Dim dt As New DataTable

        Dim i, hitung As Integer
        Dim urutan As String

        DataGridView1.Rows.Clear()
        Dim dsTables As DataTable = ds2.Tables("DataTransaksi")

        For i = 0 To ds2.Tables("DataTransaksi").Rows.Count - 1

            DataGridView1.Rows.Add()

            With DataGridView1.Rows(i)

                'If Format(dsTables.Rows(i).Item("Tanggal_Input"), "MM") = Format(txtTgl_Transaksi.Value, "MM") And Format(dsTables.Rows(i).Item("Tanggal_Input"), "yyyy") = Format(txtTgl_Transaksi.Value, "yyyy") Then

                hitung = Len(Trim$(i))
                If hitung = 1 Then
                    urutan = "00" & i
                ElseIf hitung = 2 Then
                    urutan = "0" & i
                    'ElseIf hitung = 3 Then
                    '    urutan = "0" & i
                Else
                    urutan = i
                End If

                .Cells(0).Value = urutan
                '.Cells(1).Value = Format(dsTables.Rows(i).Item("Tgl_terima"), "dd-MM-yyyy")
                .Cells(1).Value = Format(dsTables.Rows(i).Item("Tanggal_Input"), "dd-MM-yyyy")
                .Cells(2).Value = dsTables.Rows(i).Item("Jenis_Transaksi")
                .Cells(3).Value = dsTables.Rows(i).Item("Nama_Akun")
                .Cells(4).Value = dsTables.Rows(i).Item("Debit")
                .Cells(5).Value = dsTables.Rows(i).Item("Kredit")
                .Cells(6).Value = dsTables.Rows(i).Item("Keterangan")
                .Cells(7).Value = dsTables.Rows(i).Item("Id")
                .Cells(8).Value = dsTables.Rows(i).Item("Kode_Transaksi")
                .Cells(9).Value = dsTables.Rows(i).Item("No_Akun")
                .Cells(10).Value = dsTables.Rows(i).Item("Saldo_Debet")
                .Cells(11).Value = dsTables.Rows(i).Item("Saldo_Kredit")
                .Cells(12).Value = dsTables.Rows(i).Item("Saldo_Debet2")
                .Cells(13).Value = dsTables.Rows(i).Item("Saldo_Kredit2")



                'tutupkoneksi()
                'bukakoneksi()
                'cmd = New SqlCommand("Select * from Table_Akun WHERE Nomor_Akun='" & Trim(CStr(.Cells(9).Value)) & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'", conn)

                'rd = cmd.ExecuteReader
                'rd.Read()
                'If Not rd.HasRows Then
                '    'MsgBox("Nomor Akun Tidak ada di bulan dan tahun yang Anda Pilih. Sesuaikan Tanggal dengan Tanggal Akun yang ada di Database")
                'Else

                '    MsgBox(CStr(rd.Item("No_Akun")))

                'End If

                '.Cells(6).Value = dsTables.Rows(i).Item("Perihal")
                '.Cells(7).Value = dsTables.Rows(i).Item("Id")
                '.Cells(8).Value = dsTables.Rows(i).Item("Nmr_agenda")
                '.Cells(9).Value = dsTables.Rows(i).Item("Sifat_surat")
                '.Cells(10).Value = Format(dsTables.Rows(i).Item("Tgl_diteruskan"), "dd-MM-yyyy")


                'End If
            End With

        Next
        'Dim j As Integer
        'For j = DataGridView1.RowCount - 1 To 0 Step -1
        '    If Trim(DataGridView1(0, j).Value) = "" Then DataGridView1.Rows.RemoveAt(j)
        'Next
        'For Each dt In ds.Tables
        '    DataGridView2.DataSource = dt
        'Next
        'tutupkoneksi()
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        Dim i As Integer
        i = Me.DataGridView1.CurrentRow.Index


        With DataGridView1.Rows.Item(i)
            'MsgBox(CStr(Convert.ToDateTime(.Cells(1).Value)))
            If (CStr(Convert.ToDateTime(.Cells(1).Value)) <> "00.00.00") Then
                cmbJnsTransaksi.Text = .Cells(2).Value
                txtNmAkun.Text = .Cells(3).Value
                txtKodeTransaksi.Text = .Cells(8).Value
                'TG = Format(.Cells(1).Value, "dd") & "/" & Format(.Cells(1).Value, "MM") & "/" & Format(.Cells(1).Value, "yyyy")
                'Dim b As String
                txtTgl_Transaksi.Value = Convert.ToDateTime(.Cells(1).Value)
                'MsgBox(b)
                'txtTgl_Transaksi.Value = New System.DateTime(2018, 6, b, 0, 0, 0, 0)
                If .Cells(2).Value = "Transaksi Kas Masuk" Then
                    txtNominal.Text = .Cells(4).Value
                    txtSaldo.Text = IIf(IsDBNull(.Cells(10).Value), "0", .Cells(10).Value)
                    txtDebetKredit.Text = .Cells(4).Value
                Else
                    txtNominal.Text = .Cells(5).Value
                    txtSaldo.Text = IIf(IsDBNull(.Cells(11).Value), "0", .Cells(11).Value)
                    txtDebetKredit.Text = .Cells(5).Value
                End If
                'txtNominal.Text = .Cells(5).Value
                'txtKredit.Text = .Cells(6).Value
                txtKeterangan.Text = .Cells(6).Value
                txtId.Text = .Cells(7).Value
                txtNoAkun.Text = .Cells(9).Value

                tutupkoneksi()
                bukakoneksi()
                cmd = New SqlCommand("Select * from Table_Akun WHERE Nomor_Akun='" & Trim(txtNoAkun.Text) & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'", conn)

                rd = cmd.ExecuteReader
                rd.Read()
                If Not rd.HasRows Then
                    'MsgBox("Nomor Akun Tidak ada di bulan dan tahun yang Anda Pilih. Sesuaikan Tanggal dengan Tanggal Akun yang ada di Database")
                Else
                    '.Cells(14).Value = rd.Item("Saldo_Debet")
                    If rd.Item("Saldo_Debet") <> "0" Then
                        txtSaldo.Text = rd.Item("Saldo_Debet")
                    Else

                        txtSaldo.Text = rd.Item("Saldo_Kredit")
                    End If
                End If

                'tutupkoneksi()
                'bukakoneksi()
                'cmd = New SqlCommand("Select * from Table_Akun WHERE Nomor_Akun='" & Trim(txtNoAkun.Text) & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'", conn)

                'rd = cmd.ExecuteReader
                'rd.Read()
                'If Not rd.HasRows Then
                '    'MsgBox("Nomor Akun Tidak ada di bulan dan tahun yang Anda Pilih. Sesuaikan Tanggal dengan Tanggal Akun yang ada di Database")
                'Else

                '    MsgBox(rd.Item("Nomor_Akun"))

                'End If
                '.Columns.Add("No", "NO.")
                '.Columns.Add("Tanggal", "TANGGAL")
                '.Columns.Add("JenisTransaksi", "JENIS TRANSAKSI")
                '.Columns.Add("NamaAkun", "NAMA AKUN")
                '.Columns.Add("Debit", "DEBIT")
                '.Columns.Add("Kredit", "KREDIT")
                '.Columns.Add("Keterangan", "KETERANGAN")
                '.Columns.Add("Id", "Id")
                cmbJnsTransaksi.Enabled = False
                txtNoAkun.Enabled = False
                Button2.Enabled = False

                btnSave.Visible = False
                btnUpdate.Visible = True

            End If
        End With
        'MsgBox(i)
        'Call reset()


    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click


        'Sql3 = "INSERT INTO Data_Transaksi(Tanggal_Input, Jenis_Transaksi, Keterangan, Debit, Nama_Akun, Kode_Transaksi, No_Akun, Kredit, Saldo_Debet, Saldo_Kredit)" & _
        '              " VALUES ('" & Format(txtTgl_Transaksi.Value, "MM-dd-yyyy") & "','" & cmbJnsTransaksi.Text & "','" & txtKeterangan.Text & "','" & txtNominal.Text & "','" & txtNmAkun.Text & "','" & txtKodeTransaksi.Text & "','" & txtNoAkun.Text & "','0','" & txtSaldo.Text & "','0')"


        If txtNominal.Text = "0" Or txtNominal.Text = "" Then
            MsgBox("Nominal Tidak Boleh Kosong atau Nol, Silakan Lengkapi lagi.")
        Else



            tutupkoneksi()
            bukakoneksi()
            cmd = New SqlCommand("Select * from Table_Akun WHERE Nomor_Akun='" & Trim(txtNoAkun.Text) & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'", conn)

            rd = cmd.ExecuteReader
            rd.Read()
            If Not rd.HasRows Then
                MsgBox("silakan pilih tanggal")
            Else
                If CStr(rd.Item("Saldo_Debet")) <> "0" Then
                    If cmbJnsTransaksi.Text = "Transaksi Kas Masuk" Then
                        'MsgBox("masuk")
                        Sql2 = "UPDATE Data_Transaksi Set Tanggal_Input='" & Format(txtTgl_Transaksi.Value, "MM-dd-yyyy") & "',Jenis_Transaksi='" & CStr(cmbJnsTransaksi.Text) & "',keterangan='" & CStr(txtKeterangan.Text) & "',Debit='" & txtNominal.Text & "',Nama_Akun='" & CStr(txtNmAkun.Text) & "',No_Akun='" & CStr(txtNoAkun.Text) & "',Saldo_Debet='" & CStr(txtSaldo.Text) & "',Saldo_Kredit='0',Kredit='0',Saldo_Debet2='" & (Val(txtSaldo.Text) + Val(txtDebetKredit.Text)) - Val(txtNominal.Text) & "',Saldo_Kredit2='0' WHERE Id = '" & CInt(txtId.Text) & "'"

                        Dim command As New SqlCommand(Sql2, Module2.Koneksi)
                        command.ExecuteNonQuery()
                        Module2.Koneksi.Close()

                        Sql5 = "UPDATE Table_Akun Set Saldo_Debet='" & (Val(txtSaldo.Text) + Val(txtDebetKredit.Text)) - Val(txtNominal.Text) & "' Where Nomor_Akun='" & Trim(txtNoAkun.Text) & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'"

                        Dim command5 As New SqlCommand(Sql5, Module2.Koneksi)
                        command5.ExecuteNonQuery()
                        Module2.Koneksi.Close()
                    Else
                        Sql2 = "UPDATE Data_Transaksi Set Tanggal_Input='" & Format(txtTgl_Transaksi.Value, "MM-dd-yyyy") & "',Jenis_Transaksi='" & CStr(cmbJnsTransaksi.Text) & "',keterangan='" & CStr(txtKeterangan.Text) & "',Kredit='" & txtNominal.Text & "',Nama_Akun='" & CStr(txtNmAkun.Text) & "',No_Akun='" & CStr(txtNoAkun.Text) & "',Saldo_Kredit='" & CStr(txtSaldo.Text) & "',Saldo_Debet='0',Debit='0',Saldo_Kredit2='0',Saldo_Debet2='" & (Val(txtSaldo.Text) - Val(txtDebetKredit.Text)) + Val(txtNominal.Text) & "' WHERE Id = '" & CInt(txtId.Text) & "'"

                        Dim command As New SqlCommand(Sql2, Module2.Koneksi)
                        command.ExecuteNonQuery()
                        Module2.Koneksi.Close()

                        'MsgBox("keluar")
                        Sql5 = "UPDATE Table_Akun Set Saldo_Debet='" & (Val(txtSaldo.Text) - Val(txtDebetKredit.Text)) + Val(txtNominal.Text) & "' Where Nomor_Akun='" & Trim(txtNoAkun.Text) & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'"

                        Dim command5 As New SqlCommand(Sql5, Module2.Koneksi)
                        command5.ExecuteNonQuery()
                        Module2.Koneksi.Close()
                    End If
                Else
                    If cmbJnsTransaksi.Text = "Transaksi Kas Masuk" Then
                        'MsgBox("transaksi masuk")

                        If Val(txtSaldo.Text) > Val(txtDebetKredit.Text) Then

                            Sql2 = "UPDATE Data_Transaksi Set Tanggal_Input='" & Format(txtTgl_Transaksi.Value, "MM-dd-yyyy") & "',Jenis_Transaksi='" & CStr(cmbJnsTransaksi.Text) & "',keterangan='" & CStr(txtKeterangan.Text) & "',Debit='" & txtNominal.Text & "',Nama_Akun='" & CStr(txtNmAkun.Text) & "',No_Akun='" & CStr(txtNoAkun.Text) & "',Saldo_Debet='" & CStr(txtSaldo.Text) & "',Saldo_Kredit='0',Kredit='0',Saldo_Kredit2='" & (Val(txtSaldo.Text) - Val(txtDebetKredit.Text)) + Val(txtNominal.Text) & "',Saldo_Debet2='0' WHERE Id = '" & CInt(txtId.Text) & "'"

                            Dim command As New SqlCommand(Sql2, Module2.Koneksi)
                            command.ExecuteNonQuery()
                            Module2.Koneksi.Close()

                            Sql5 = "UPDATE Table_Akun Set Saldo_Kredit='" & (Val(txtSaldo.Text) - Val(txtDebetKredit.Text)) + Val(txtNominal.Text) & "' Where Nomor_Akun='" & Trim(txtNoAkun.Text) & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'"

                            Dim command5 As New SqlCommand(Sql5, Module2.Koneksi)
                            command5.ExecuteNonQuery()
                            Module2.Koneksi.Close()
                        Else
                            Sql2 = "UPDATE Data_Transaksi Set Tanggal_Input='" & Format(txtTgl_Transaksi.Value, "MM-dd-yyyy") & "',Jenis_Transaksi='" & CStr(cmbJnsTransaksi.Text) & "',keterangan='" & CStr(txtKeterangan.Text) & "',Debit='" & txtNominal.Text & "',Nama_Akun='" & CStr(txtNmAkun.Text) & "',No_Akun='" & CStr(txtNoAkun.Text) & "',Saldo_Debet='" & CStr(txtSaldo.Text) & "',Saldo_Kredit='0',Kredit='0',Saldo_Kredit2='" & (Val(txtDebetKredit.Text) - Val(txtSaldo.Text)) + Val(txtNominal.Text) & "',Saldo_Debet2='0' WHERE Id = '" & CInt(txtId.Text) & "'"

                            Dim command As New SqlCommand(Sql2, Module2.Koneksi)
                            command.ExecuteNonQuery()
                            Module2.Koneksi.Close()

                            Sql5 = "UPDATE Table_Akun Set Saldo_Kredit='" & Val(txtDebetKredit.Text) - (Val(txtSaldo.Text)) + Val(txtNominal.Text) & "' Where Nomor_Akun='" & Trim(txtNoAkun.Text) & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'"

                            Dim command5 As New SqlCommand(Sql5, Module2.Koneksi)
                            command5.ExecuteNonQuery()
                            Module2.Koneksi.Close()
                        End If


                        'Sql5 = "UPDATE Table_Akun Set Saldo_Kredit='" & Val(txtSaldo.Text) - Val(txtNominal.Text) & "' Where Nomor_Akun='" & Trim(txtNoAkun.Text) & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'"

                        'Dim command5 As New SqlCommand(Sql5, Module2.Koneksi)
                        'command5.ExecuteNonQuery()
                        'Module2.Koneksi.Close()
                    Else
                        'MsgBox("transaksi keluar")



                        Sql2 = "UPDATE Data_Transaksi Set Tanggal_Input='" & Format(txtTgl_Transaksi.Value, "MM-dd-yyyy") & "',Jenis_Transaksi='" & CStr(cmbJnsTransaksi.Text) & "',keterangan='" & CStr(txtKeterangan.Text) & "',Kredit='" & txtNominal.Text & "',Nama_Akun='" & CStr(txtNmAkun.Text) & "',No_Akun='" & CStr(txtNoAkun.Text) & "',Saldo_Kredit='" & CStr(txtSaldo.Text) & "',Saldo_Debet='0',Debit='0',Saldo_Debet2='0',Saldo_Kredit2='" & (Val(txtSaldo.Text) + Val(txtDebetKredit.Text)) - Val(txtNominal.Text) & "' WHERE Id = '" & CInt(txtId.Text) & "'"

                        Dim command As New SqlCommand(Sql2, Module2.Koneksi)
                        command.ExecuteNonQuery()
                        Module2.Koneksi.Close()

                        'MsgBox("keluar")
                        Sql5 = "UPDATE Table_Akun Set Saldo_Kredit='" & (Val(txtSaldo.Text) + Val(txtDebetKredit.Text)) - Val(txtNominal.Text) & "' Where Nomor_Akun='" & Trim(txtNoAkun.Text) & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'"

                        Dim command5 As New SqlCommand(Sql5, Module2.Koneksi)
                        command5.ExecuteNonQuery()
                        Module2.Koneksi.Close()



                        'Sql5 = "UPDATE Table_Akun Set Saldo_Kredit='" & Val(txtSaldo.Text) + Val(txtNominal.Text) & "' Where Nomor_Akun='" & Trim(txtNoAkun.Text) & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'"

                        'Dim command5 As New SqlCommand(Sql5, Module2.Koneksi)
                        'command5.ExecuteNonQuery()
                        'Module2.Koneksi.Close()
                    End If
                End If

                MsgBox("Data sukses terupdate", vbInformation, "Konfirmasi")

                Call tampil()
                Call reset()

                btnSave.Visible = True
                btnUpdate.Visible = False

            End If


        End If

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        reset()
        tampil()
        txtTgl_Transaksi.Value = Format(Now, "dd-MM-yyyy")
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click


        tutupkoneksi()
        bukakoneksi()
        cmd = New SqlCommand("Select * from Data_Transaksi WHERE Id='" & Trim(txtId.Text) & "'", conn)

        rd = cmd.ExecuteReader
        rd.Read()
        'MsgBox(rd.HasRows)

        'Dim rs As Integer

        'rs = cmd.ExecuteScalar

        'Dim JmlhRec = cmd.ExecuteScalar
        'MsgBox("Jumlah Record: " & JmlhRec)


        'MsgBox(rs)
        If Not rd.HasRows Then
        Else
            If Format(txtTgl_Transaksi.Value, "MM") = Format(rd.Item("Tanggal_Input"), "MM") And Format(txtTgl_Transaksi.Value, "yyyy") = Format(rd.Item("Tanggal_Input"), "yyyy") Then
                'Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'
                'MsgBox(rd.Item("Saldo_Debet") & " - " & rd.Item("Saldo_Debet"))
                tutupkoneksi()
                bukakoneksi()
                cmd2 = New SqlCommand("Select count(*) from Data_Transaksi WHERE No_Akun='" & Trim(txtNoAkun.Text) & "'", conn)


                Dim rs As Integer

                rs = cmd2.ExecuteScalar

                If rs = 1 And cmbJnsTransaksi.Text = "Transaksi Kas Keluar" Then
                    'MsgBox(rs & " " & cmbJnsTransaksi.Text)






                    Dim Sql As String = String.Empty

                    Sql = "Delete from Data_Transaksi where Id = '" & Trim(txtId.Text) & "'"

                    Dim command As New SqlCommand(Sql, Module2.Koneksi)

                    command.ExecuteNonQuery()

                    Module2.Koneksi.Close()

                    'tutupkoneksi()
                    'bukakoneksi()

                    'Format(Now, "MM")

                    'cmd = New SqlCommand("Select * from Table_Akun WHERE Nomor_Akun='" & Trim(Kd_Akun) & "'", conn)

                    'rd = cmd.ExecuteReader
                    'rd.Read()
                    'If Not rd.HasRows Then
                    'Else
                    '    MsgBox(rd.Item("Saldo_Debet"))
                    'End If


                    tutupkoneksi()
                    bukakoneksi()
                    cmd = New SqlCommand("Select * from Table_Akun WHERE Nomor_Akun='" & Trim(txtNoAkun.Text) & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'", conn)

                    rd = cmd.ExecuteReader
                    rd.Read()
                    If Not rd.HasRows Then
                    Else
                        If CStr(rd.Item("Saldo_Debet")) <> "0" Then
                            If cmbJnsTransaksi.Text = "Transaksi Kas Masuk" Then
                                Sql5 = "UPDATE Table_Akun Set Saldo_Debet='" & Val(txtSaldo.Text) + Val(txtNominal.Text) & "' Where Nomor_Akun='" & Trim(txtNoAkun.Text) & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'"

                                Dim command5 As New SqlCommand(Sql5, Module2.Koneksi)
                                command5.ExecuteNonQuery()
                                Module2.Koneksi.Close()
                            Else
                                Sql5 = "UPDATE Table_Akun Set Saldo_Debet='" & Val(txtSaldo.Text) - Val(txtNominal.Text) & "' Where Nomor_Akun='" & Trim(txtNoAkun.Text) & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'"

                                Dim command5 As New SqlCommand(Sql5, Module2.Koneksi)
                                command5.ExecuteNonQuery()
                                Module2.Koneksi.Close()
                            End If
                        Else
                            If cmbJnsTransaksi.Text = "Transaksi Kas Masuk" Then
                                Sql5 = "UPDATE Table_Akun Set Saldo_Kredit='" & Val(txtSaldo.Text) + Val(txtNominal.Text) & "' Where Nomor_Akun='" & Trim(txtNoAkun.Text) & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'"

                                Dim command5 As New SqlCommand(Sql5, Module2.Koneksi)
                                command5.ExecuteNonQuery()
                                Module2.Koneksi.Close()
                            Else
                                Sql5 = "UPDATE Table_Akun Set Saldo_Kredit='" & Val(txtSaldo.Text) + Val(txtNominal.Text) & "' Where Nomor_Akun='" & Trim(txtNoAkun.Text) & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'"

                                Dim command5 As New SqlCommand(Sql5, Module2.Koneksi)
                                command5.ExecuteNonQuery()
                                Module2.Koneksi.Close()
                            End If
                        End If
                    End If

                    MessageBox.Show("sukses terhapus")
                    reset()
                    tampil()




                Else










                    Dim Sql As String = String.Empty

                    Sql = "Delete from Data_Transaksi where Id = '" & Trim(txtId.Text) & "'"

                    Dim command As New SqlCommand(Sql, Module2.Koneksi)

                    command.ExecuteNonQuery()

                    Module2.Koneksi.Close()

                    'tutupkoneksi()
                    'bukakoneksi()

                    'Format(Now, "MM")

                    'cmd = New SqlCommand("Select * from Table_Akun WHERE Nomor_Akun='" & Trim(Kd_Akun) & "'", conn)

                    'rd = cmd.ExecuteReader
                    'rd.Read()
                    'If Not rd.HasRows Then
                    'Else
                    '    MsgBox(rd.Item("Saldo_Debet"))
                    'End If


                    tutupkoneksi()
                    bukakoneksi()
                    cmd = New SqlCommand("Select * from Table_Akun WHERE Nomor_Akun='" & Trim(txtNoAkun.Text) & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'", conn)

                    rd = cmd.ExecuteReader
                    rd.Read()
                    If Not rd.HasRows Then
                    Else
                        If CStr(rd.Item("Saldo_Debet")) <> "0" Then
                            If cmbJnsTransaksi.Text = "Transaksi Kas Masuk" Then
                                Sql5 = "UPDATE Table_Akun Set Saldo_Debet='" & Val(txtSaldo.Text) + Val(txtNominal.Text) & "' Where Nomor_Akun='" & Trim(txtNoAkun.Text) & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'"

                                Dim command5 As New SqlCommand(Sql5, Module2.Koneksi)
                                command5.ExecuteNonQuery()
                                Module2.Koneksi.Close()
                            Else
                                Sql5 = "UPDATE Table_Akun Set Saldo_Debet='" & Val(txtSaldo.Text) - Val(txtNominal.Text) & "' Where Nomor_Akun='" & Trim(txtNoAkun.Text) & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'"

                                Dim command5 As New SqlCommand(Sql5, Module2.Koneksi)
                                command5.ExecuteNonQuery()
                                Module2.Koneksi.Close()
                            End If
                        Else
                            If cmbJnsTransaksi.Text = "Transaksi Kas Masuk" Then
                                Sql5 = "UPDATE Table_Akun Set Saldo_Kredit='" & Val(txtSaldo.Text) - Val(txtNominal.Text) & "' Where Nomor_Akun='" & Trim(txtNoAkun.Text) & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'"

                                Dim command5 As New SqlCommand(Sql5, Module2.Koneksi)
                                command5.ExecuteNonQuery()
                                Module2.Koneksi.Close()
                            Else
                                Sql5 = "UPDATE Table_Akun Set Saldo_Kredit='" & Val(txtSaldo.Text) + Val(txtNominal.Text) & "' Where Nomor_Akun='" & Trim(txtNoAkun.Text) & "' AND Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'"

                                Dim command5 As New SqlCommand(Sql5, Module2.Koneksi)
                                command5.ExecuteNonQuery()
                                Module2.Koneksi.Close()
                            End If
                        End If
                    End If

                    MessageBox.Show("sukses terhapus")
                    reset()
                    tampil()




                End If
            End If
        End If















    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Me.Hide()
        frmCariAkun.Bln = Format(txtTgl_Transaksi.Value, "MM")
        frmCariAkun.Thn = Format(txtTgl_Transaksi.Value, "yyyy")
        frmCariAkun.Show()
        'Me.Visible = False

        'tanggalTrans = txtTgl_Transaksi.Value
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Hide()
        frmCariAkunKredit.Show()

    End Sub

    Private Sub cmbJnsTransaksi_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbJnsTransaksi.SelectedIndexChanged
        If cmbJnsTransaksi.Text = "Transaksi Kas Masuk" Then
            trans = "debit"
        Else
            trans = "kredit"
        End If
    End Sub


    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        tutupkoneksi()
        bukakoneksi()
        cmd = New SqlCommand("Select * from Data_Transaksi WHERE No_Akun='" & Trim(txtNoAkun.Text) & "'", conn)

        rd = cmd.ExecuteReader
        rd.Read()
        'MsgBox(rd.HasRows)

        'Dim rs As Integer

        'rs = cmd.ExecuteScalar

        'Dim JmlhRec = cmd.ExecuteScalar
        'MsgBox("Jumlah Record: " & JmlhRec)


        'MsgBox(rs)
        If Not rd.HasRows Then
        Else
            If Format(txtTgl_Transaksi.Value, "MM") = Format(rd.Item("Tanggal_Input"), "MM") And Format(txtTgl_Transaksi.Value, "yyyy") = Format(rd.Item("Tanggal_Input"), "yyyy") Then
                'Bulan='" & Format(txtTgl_Transaksi.Value, "MM") & "' AND Tahun='" & Format(txtTgl_Transaksi.Value, "yyyy") & "'
                'MsgBox(rd.Item("Saldo_Debet") & " - " & rd.Item("Saldo_Debet"))
                tutupkoneksi()
                bukakoneksi()
                cmd2 = New SqlCommand("Select count(*) from Data_Transaksi WHERE No_Akun='" & Trim(txtNoAkun.Text) & "'", conn)


                Dim rs As Integer

                rs = cmd2.ExecuteScalar
                MsgBox(rs & " " & cmbJnsTransaksi.Text)
            End If
        End If
    End Sub

  
End Class