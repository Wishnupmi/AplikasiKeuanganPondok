Imports System.Data.SqlClient
Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class frmArusKas
    Dim koneksi = Module2.Koneksi
    Dim con As New OleDbConnection
    'Dim ds As New DataSet
    Dim da As OleDbDataAdapter
    Dim dispoBaris As Integer
    Public nmr_urut As String
    Dim bln_romawi As String

    Dim Sql100 As String = String.Empty
    Dim Sql101 As String = String.Empty
    Dim Sql102 As String = String.Empty
    Dim Sql103 As String = String.Empty
    Dim Sql104 As String = String.Empty
    Dim Sql As String = String.Empty
    Dim Sql1 As String = String.Empty
    Dim Sql2 As String = String.Empty
    Dim Sql3 As String = String.Empty
    Dim Sql4 As String = String.Empty
    Dim Sql5 As String = String.Empty
    Dim Sql6 As String = String.Empty
    Dim Sql7 As String = String.Empty
    Dim Sql8 As String = String.Empty
    Dim Sql9 As String = String.Empty
    Dim Sql10 As String = String.Empty
    Dim Sql11 As String = String.Empty
    Dim Sql12 As String = String.Empty
    Dim Sql13 As String = String.Empty
    Dim Sql14 As String = String.Empty
    Dim Sql15 As String = String.Empty
    Dim Sql16 As String = String.Empty
    Dim Sql17 As String = String.Empty
    Dim Sql18 As String = String.Empty
    Dim Sql19 As String = String.Empty
    Dim Sql20 As String = String.Empty
    Dim Sql21 As String = String.Empty
    Dim Sql22 As String = String.Empty
    Dim Sql23 As String = String.Empty
    Dim Sql24 As String = String.Empty
    Dim Sql25 As String = String.Empty
    Dim Sql26 As String = String.Empty
    Dim Sql27 As String = String.Empty
    Dim Sql28 As String = String.Empty
    Dim Sql29 As String = String.Empty
    Dim Sql30 As String = String.Empty
    Dim Sql31 As String = String.Empty
    Dim Sql32 As String = String.Empty
    Dim Sql33 As String = String.Empty
    Dim Sql34 As String = String.Empty
    Dim Sql35 As String = String.Empty
    Dim Sql36 As String = String.Empty
    Dim Sql37 As String = String.Empty
    Dim Sql38 As String = String.Empty
    Dim Sql39 As String = String.Empty
    Dim Sql40 As String = String.Empty
    Dim Sql41 As String = String.Empty
    Dim Sql42 As String = String.Empty
    Dim Sql43 As String = String.Empty
    Dim Sql44 As String = String.Empty
    Dim Sql45 As String = String.Empty
    Dim Sql46 As String = String.Empty
    Dim Sql47 As String = String.Empty
    Dim Sql48 As String = String.Empty
    Dim Sql49 As String = String.Empty
    Dim Sql50 As String = String.Empty
    Dim Sql51 As String = String.Empty
    Dim Sql52 As String = String.Empty
    Dim Sql53 As String = String.Empty
    Dim Sql54 As String = String.Empty
    Dim Sql55 As String = String.Empty
    Dim Sql56 As String = String.Empty
    Dim Sql57 As String = String.Empty
    Dim Sql58 As String = String.Empty
    Dim Sql59 As String = String.Empty
    Dim Sql60 As String = String.Empty
    Dim Sql61 As String = String.Empty
    Dim Sql62 As String = String.Empty
    Dim Sql63 As String = String.Empty
    Dim Sql64 As String = String.Empty
    Dim Sql65 As String = String.Empty
    Dim Sql66 As String = String.Empty

    Dim Sql67 As String = String.Empty
    Dim Sql68 As String = String.Empty
    Dim Sql69 As String = String.Empty
    Dim Sql70 As String = String.Empty
    Dim Sql71 As String = String.Empty
    Dim Sql72 As String = String.Empty
    Dim Sql73 As String = String.Empty
    Dim Sql74 As String = String.Empty
    Dim Sql75 As String = String.Empty

    Dim bln, bln_tahun As String
    Dim Nomor_Akun100, Nama_Akun100, Kelompok_Akun100, Tanggal_Input100, Bulan100, Tahun100, header100 As String
    Dim Nomor_Akun102, Nama_Akun102, Kelompok_Akun102, Tanggal_Input102, Bulan102, Tahun102, header102 As String
    Dim Nomor_Akun103, Nama_Akun103, Kelompok_Akun103, Tanggal_Input103, Bulan103, Tahun103, header103 As String
    Dim Nomor_Akun104, Nama_Akun104, Kelompok_Akun104, Tanggal_Input104, Bulan104, Tahun104, header104 As String
    Dim Nomor_Akun1, Nama_Akun1, Kelompok_Akun1, Tanggal_Input1, Bulan1, Tahun1, header1 As String
    Dim Nomor_Akun2, Nama_Akun2, Kelompok_Akun2, Tanggal_Input2, Bulan2, Tahun2, header2 As String
    Dim Nomor_Akun3, Nama_Akun3, Kelompok_Akun3, Tanggal_Input3, Bulan3, Tahun3, header3 As String
    Dim Nomor_Akun4, Nama_Akun4, Kelompok_Akun4, Tanggal_Input4, Bulan4, Tahun4, header4 As String
    Dim Nomor_Akun5, Nama_Akun5, Kelompok_Akun5, Tanggal_Input5, Bulan5, Tahun5, header5 As String
    Dim Nomor_Akun6, Nama_Akun6, Kelompok_Akun6, Tanggal_Input6, Bulan6, Tahun6, header6 As String
    Dim Nomor_Akun7, Nama_Akun7, Kelompok_Akun7, Tanggal_Input7, Bulan7, Tahun7, header7 As String
    Dim Nomor_Akun8, Nama_Akun8, Kelompok_Akun8, Tanggal_Input8, Bulan8, Tahun8, header8 As String
    Dim Nomor_Akun9, Nama_Akun9, Kelompok_Akun9, Tanggal_Input9, Bulan9, Tahun9, header9 As String
    Dim Nomor_Akun10, Nama_Akun10, Kelompok_Akun10, Tanggal_Input10, Bulan10, Tahun10, header10 As String
    Dim Nomor_Akun11, Nama_Akun11, Kelompok_Akun11, Tanggal_Input11, Bulan11, Tahun11, header11 As String
    Dim Nomor_Akun12, Nama_Akun12, Kelompok_Akun12, Tanggal_Input12, Bulan12, Tahun12, header12 As String
    Dim Nomor_Akun13, Nama_Akun13, Kelompok_Akun13, Tanggal_Input13, Bulan13, Tahun13, header13 As String
    Dim Nomor_Akun14, Nama_Akun14, Kelompok_Akun14, Tanggal_Input14, Bulan14, Tahun14, header14 As String
    Dim Nomor_Akun15, Nama_Akun15, Kelompok_Akun15, Tanggal_Input15, Bulan15, Tahun15, header15 As String
    Dim Nomor_Akun16, Nama_Akun16, Kelompok_Akun16, Tanggal_Input16, Bulan16, Tahun16, header16 As String
    Dim Nomor_Akun17, Nama_Akun17, Kelompok_Akun17, Tanggal_Input17, Bulan17, Tahun17, header17 As String
    Dim Nomor_Akun18, Nama_Akun18, Kelompok_Akun18, Tanggal_Input18, Bulan18, Tahun18, header18 As String
    Dim Nomor_Akun19, Nama_Akun19, Kelompok_Akun19, Tanggal_Input19, Bulan19, Tahun19, header19 As String
    Dim Nomor_Akun20, Nama_Akun20, Kelompok_Akun20, Tanggal_Input20, Bulan20, Tahun20, header20 As String
    Dim Nomor_Akun21, Nama_Akun21, Kelompok_Akun21, Tanggal_Input21, Bulan21, Tahun21, header21 As String
    Dim Nomor_Akun22, Nama_Akun22, Kelompok_Akun22, Tanggal_Input22, Bulan22, Tahun22, header22 As String
    Dim Nomor_Akun23, Nama_Akun23, Kelompok_Akun23, Tanggal_Input23, Bulan23, Tahun23, header23 As String
    Dim Nomor_Akun24, Nama_Akun24, Kelompok_Akun24, Tanggal_Input24, Bulan24, Tahun24, header24 As String
    Dim Nomor_Akun25, Nama_Akun25, Kelompok_Akun25, Tanggal_Input25, Bulan25, Tahun25, header25 As String
    Dim Nomor_Akun26, Nama_Akun26, Kelompok_Akun26, Tanggal_Input26, Bulan26, Tahun26, header26 As String
    Dim Nomor_Akun27, Nama_Akun27, Kelompok_Akun27, Tanggal_Input27, Bulan27, Tahun27, header27 As String
    Dim Nomor_Akun28, Nama_Akun28, Kelompok_Akun28, Tanggal_Input28, Bulan28, Tahun28, header28 As String
    Dim Nomor_Akun29, Nama_Akun29, Kelompok_Akun29, Tanggal_Input29, Bulan29, Tahun29, header29 As String
    Dim Nomor_Akun30, Nama_Akun30, Kelompok_Akun30, Tanggal_Input30, Bulan30, Tahun30, header30 As String
    Dim Nomor_Akun31, Nama_Akun31, Kelompok_Akun31, Tanggal_Input31, Bulan31, Tahun31, header31 As String
    Dim Nomor_Akun32, Nama_Akun32, Kelompok_Akun32, Tanggal_Input32, Bulan32, Tahun32, header32 As String
    Dim Nomor_Akun33, Nama_Akun33, Kelompok_Akun33, Tanggal_Input33, Bulan33, Tahun33, header33 As String
    Dim Nomor_Akun34, Nama_Akun34, Kelompok_Akun34, Tanggal_Input34, Bulan34, Tahun34, header34 As String
    Dim Nomor_Akun35, Nama_Akun35, Kelompok_Akun35, Tanggal_Input35, Bulan35, Tahun35, header35 As String
    Dim Nomor_Akun36, Nama_Akun36, Kelompok_Akun36, Tanggal_Input36, Bulan36, Tahun36, header36 As String
    Dim Nomor_Akun37, Nama_Akun37, Kelompok_Akun37, Tanggal_Input37, Bulan37, Tahun37, header37 As String
    Dim Nomor_Akun38, Nama_Akun38, Kelompok_Akun38, Tanggal_Input38, Bulan38, Tahun38, header38 As String
    Dim Nomor_Akun39, Nama_Akun39, Kelompok_Akun39, Tanggal_Input39, Bulan39, Tahun39, header39 As String
    Dim Nomor_Akun40, Nama_Akun40, Kelompok_Akun40, Tanggal_Input40, Bulan40, Tahun40, header40 As String
    Dim Nomor_Akun41, Nama_Akun41, Kelompok_Akun41, Tanggal_Input41, Bulan41, Tahun41, header41 As String
    Dim Nomor_Akun42, Nama_Akun42, Kelompok_Akun42, Tanggal_Input42, Bulan42, Tahun42, header42 As String
    Dim Nomor_Akun43, Nama_Akun43, Kelompok_Akun43, Tanggal_Input43, Bulan43, Tahun43, header43 As String
    Dim Nomor_Akun44, Nama_Akun44, Kelompok_Akun44, Tanggal_Input44, Bulan44, Tahun44, header44 As String
    Dim Nomor_Akun45, Nama_Akun45, Kelompok_Akun45, Tanggal_Input45, Bulan45, Tahun45, header45 As String
    Dim Nomor_Akun46, Nama_Akun46, Kelompok_Akun46, Tanggal_Input46, Bulan46, Tahun46, header46 As String
    Dim Nomor_Akun47, Nama_Akun47, Kelompok_Akun47, Tanggal_Input47, Bulan47, Tahun47, header47 As String
    Dim Nomor_Akun48, Nama_Akun48, Kelompok_Akun48, Tanggal_Input48, Bulan48, Tahun48, header48 As String
    Dim Nomor_Akun49, Nama_Akun49, Kelompok_Akun49, Tanggal_Input49, Bulan49, Tahun49, header49 As String
    Dim Nomor_Akun50, Nama_Akun50, Kelompok_Akun50, Tanggal_Input50, Bulan50, Tahun50, header50 As String
    Dim Nomor_Akun51, Nama_Akun51, Kelompok_Akun51, Tanggal_Input51, Bulan51, Tahun51, header51 As String
    Dim Nomor_Akun52, Nama_Akun52, Kelompok_Akun52, Tanggal_Input52, Bulan52, Tahun52, header52 As String
    Dim Nomor_Akun53, Nama_Akun53, Kelompok_Akun53, Tanggal_Input53, Bulan53, Tahun53, header53 As String
    Dim Nomor_Akun54, Nama_Akun54, Kelompok_Akun54, Tanggal_Input54, Bulan54, Tahun54, header54 As String
    Dim Nomor_Akun55, Nama_Akun55, Kelompok_Akun55, Tanggal_Input55, Bulan55, Tahun55, header55 As String
    Dim Nomor_Akun56, Nama_Akun56, Kelompok_Akun56, Tanggal_Input56, Bulan56, Tahun56, header56 As String
    Dim Nomor_Akun57, Nama_Akun57, Kelompok_Akun57, Tanggal_Input57, Bulan57, Tahun57, header57 As String
    Dim Nomor_Akun58, Nama_Akun58, Kelompok_Akun58, Tanggal_Input58, Bulan58, Tahun58, header58 As String
    Dim Nomor_Akun59, Nama_Akun59, Kelompok_Akun59, Tanggal_Input59, Bulan59, Tahun59, header59 As String
    Dim Nomor_Akun60, Nama_Akun60, Kelompok_Akun60, Tanggal_Input60, Bulan60, Tahun60, header60 As String
    Dim Nomor_Akun61, Nama_Akun61, Kelompok_Akun61, Tanggal_Input61, Bulan61, Tahun61, header61 As String
    Dim Nomor_Akun62, Nama_Akun62, Kelompok_Akun62, Tanggal_Input62, Bulan62, Tahun62, header62 As String
    Dim Nomor_Akun63, Nama_Akun63, Kelompok_Akun63, Tanggal_Input63, Bulan63, Tahun63, header63 As String
    Dim Nomor_Akun64, Nama_Akun64, Kelompok_Akun64, Tanggal_Input64, Bulan64, Tahun64, header64 As String
    Dim Nomor_Akun65, Nama_Akun65, Kelompok_Akun65, Tanggal_Input65, Bulan65, Tahun65, header65 As String
    Dim Nomor_Akun66, Nama_Akun66, Kelompok_Akun66, Tanggal_Input66, Bulan66, Tahun66, header66 As String




    Dim Nomor_Akun67, Nama_Akun67, Kelompok_Akun67, Tanggal_Input67, Bulan67, Tahun67, header67 As String
    Dim Nomor_Akun68, Nama_Akun68, Kelompok_Akun68, Tanggal_Input68, Bulan68, Tahun68, header68 As String
    Dim Nomor_Akun69, Nama_Akun69, Kelompok_Akun69, Tanggal_Input69, Bulan69, Tahun69, header69 As String
    Dim Nomor_Akun70, Nama_Akun70, Kelompok_Akun70, Tanggal_Input70, Bulan70, Tahun70, header70 As String
    Dim Nomor_Akun71, Nama_Akun71, Kelompok_Akun71, Tanggal_Input71, Bulan71, Tahun71, header71 As String
    Dim Nomor_Akun72, Nama_Akun72, Kelompok_Akun72, Tanggal_Input72, Bulan72, Tahun72, header72 As String
    Dim Nomor_Akun73, Nama_Akun73, Kelompok_Akun73, Tanggal_Input73, Bulan73, Tahun73, header73 As String
    Dim Nomor_Akun74, Nama_Akun74, Kelompok_Akun74, Tanggal_Input74, Bulan74, Tahun74, header74 As String
    Dim Nomor_Akun75, Nama_Akun75, Kelompok_Akun75, Tanggal_Input75, Bulan75, Tahun75, header75 As String

    Dim i, hitung As Integer
    Dim urutan As String
    'Dim bln As String




    Private Sub frmArusKas_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        


        Dim awal, akhir, tahun, i As Integer

        awal = 2020
        akhir = 2030
        tahun = 0



        For i = awal + 1 To akhir

            txtTahun.Items.Add(i)
            'i = i + 1
        Next
    End Sub


    Public saldoakhir, jumlahneto, Kas, Piutang, Perlengkapan, BebanOperasional, BebanAdministrasi, SaldoAwal, Pendapatan_Lain, Hibah_Pendiri, AssetNeto, UangPangkalSantri, BantuanPemerintah, PendapatanInfaq, Sumbangan, Jumlah_Pend As String
    'Public BebanKonsumsiSantri, BebanListrik, BebanPerlengkapan, BebanKegiatanSantri, BebanKebersihan, BebanKesehatan, BebanIklan, BebanRapat, Jumlah_beban As String
    'Public Jumlah1, NaikTurun, SaldoAkhir, SaldoAwal, SaldoAkhir2, SaldoAwl, SaldoAkr, AsetNetoAkhir As String

    Private Sub tampil()
        tutupkoneksi()
        bukakoneksi()
        'NomorAgenda()s

        Sql = "SELECT * From Table_AkunArusKas2 WHERE Tahun='" & Trim(txtTahun.Text) & "' Order by Id_Akun asc"

        Dim da2 As New SqlDataAdapter(Sql, conn)
        Dim ds2 As New DataSet
        da2.Fill(ds2, "Table_AkunArusKas2")
        Dim dt As New DataTable



        'DataGridView1.Rows.Clear()
        Dim dsTables As DataTable = ds2.Tables("Table_AkunArusKas2")
        'MsgBox(dsTables.Rows(2).Item("Nama_Akun"))

        Kas = dsTables.Rows(3).Item("Saldo_Debet")
        Piutang = dsTables.Rows(5).Item("Saldo_Debet")
        Perlengkapan = dsTables.Rows(6).Item("Saldo_Debet")
        BebanOperasional = dsTables.Rows(0).Item("Saldo_Debet")
        BebanAdministrasi = dsTables.Rows(1).Item("Saldo_Debet")
        SaldoAwal = dsTables.Rows(2).Item("Saldo_Debet")

        jumlahneto = Val(Kas) - (Val(Piutang) + Val(Perlengkapan) + Val(BebanOperasional) + Val(BebanAdministrasi))
        saldoakhir = Val(jumlahneto) + Val(SaldoAwal)

        'PendapatanInfaq = dsTables.Rows(34).Item("Saldo_Kredit")
        'Sumbangan = dsTables.Rows(36).Item("Saldo_Kredit")

        'Jumlah_Pend = Val(kontribusi_santri) + Val(Hibah_Pendiri) + Val(UangPangkalSantri) + Val(BantuanPemerintah) + Val(PendapatanInfaq) + Val(Sumbangan)

        'BebanKonsumsiSantri = dsTables.Rows(37).Item("Saldo_Debet")
        'BebanListrik = dsTables.Rows(38).Item("Saldo_Debet")
        'BebanPerlengkapan = dsTables.Rows(39).Item("Saldo_Debet")
        'BebanKegiatanSantri = dsTables.Rows(40).Item("Saldo_Debet")
        'BebanKebersihan = dsTables.Rows(41).Item("Saldo_Debet")
        'BebanKesehatan = dsTables.Rows(42).Item("Saldo_Debet")
        'BebanIklan = dsTables.Rows(43).Item("Saldo_Debet")
        'BebanRapat = dsTables.Rows(44).Item("Saldo_Debet")

        'Jumlah_beban = Val(BebanKonsumsiSantri) + Val(BebanListrik) + Val(BebanPerlengkapan) + Val(BebanKegiatanSantri) + Val(BebanKebersihan) + Val(BebanKesehatan) + Val(BebanIklan) + Val(BebanRapat)
        'NaikTurun = Val(Jumlah_Pend) - Val(Jumlah_beban)

        'SaldoAkhir = Val(NaikTurun) + Val(AssetNeto)

        'SaldoAwal = dsTables.Rows(45).Item("Saldo_Kredit")
        'SaldoAkhir2 = dsTables.Rows(45).Item("Saldo_Kredit")

        'SaldoAwl = dsTables.Rows(46).Item("Saldo_Kredit")
        'SaldoAkr = dsTables.Rows(46).Item("Saldo_Kredit")

        'AsetNetoAkhir = Val(SaldoAkr) + Val(SaldoAkhir2) + Val(SaldoAkhir)

        'MsgBox(Jumlah1)
        'For i = 0 To ds2.Tables("DataTransaksi").Rows.Count - 1
        '    DataGridView1.Rows.Add()

        '    With DataGridView1.Rows(i)
        '        hitung = Len(Trim$(i))
        '        If hitung = 1 Then
        '            urutan = "00" & i
        '        ElseIf hitung = 2 Then
        '            urutan = "0" & i
        '            'ElseIf hitung = 3 Then
        '            '    urutan = "0" & i
        '        Else
        '            urutan = i
        '        End If

        '        .Cells(0).Value = urutan
        '        '.Cells(1).Value = Format(dsTables.Rows(i).Item("Tgl_terima"), "dd-MM-yyyy")
        '        .Cells(1).Value = Format(dsTables.Rows(i).Item("Tanggal_Input"), "dd-MM-yyyy")
        '        .Cells(2).Value = dsTables.Rows(i).Item("Jenis_Transaksi")
        '        .Cells(3).Value = dsTables.Rows(i).Item("Nama_Akun")
        '        .Cells(4).Value = dsTables.Rows(i).Item("Nominal")
        '        .Cells(5).Value = dsTables.Rows(i).Item("Debit")
        '        .Cells(6).Value = dsTables.Rows(i).Item("Kredit")
        '        .Cells(7).Value = dsTables.Rows(i).Item("Keterangan")
        '        .Cells(8).Value = dsTables.Rows(i).Item("Id")



        '        '.Cells(6).Value = dsTables.Rows(i).Item("Perihal")
        '        '.Cells(7).Value = dsTables.Rows(i).Item("Id")
        '        '.Cells(8).Value = dsTables.Rows(i).Item("Nmr_agenda")
        '        '.Cells(9).Value = dsTables.Rows(i).Item("Sifat_surat")
        '        '.Cells(10).Value = Format(dsTables.Rows(i).Item("Tgl_diteruskan"), "dd-MM-yyyy")


        '    End With
        'Next


        'For Each dt In ds.Tables
        '    DataGridView2.DataSource = dt
        'Next
        'tutupkoneksi()
    End Sub



    Private Sub cmdPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrint.Click
        

        If txtTahun.Text = "" Then
            MsgBox("Silakan pilih bulan dan tahun Laporan Keuangan", vbInformation, "Konfirmasi")
        Else


            tutupkoneksi()
            bukakoneksi()
            cmd = New SqlCommand("Select * from Table_AkunArusKas1 WHERE Tahun='" & txtTahun.Text & "'", conn)

            rd = cmd.ExecuteReader
            rd.Read()
            If Not rd.HasRows Then
                MsgBox("Tidak ada laporan di Bulan tersebut. Silakan pilih kembali bulan dan tahun Laporan Keuangan.", vbInformation, "Konfirmasi")
            Else









                Dim Sql As String = String.Empty
                tutupkoneksi()
                bukakoneksi()

                Sql = "Delete from Table_AkunArusKas2"
                Dim command As New SqlCommand(Sql, Module2.Koneksi)
                command.ExecuteNonQuery()
                Module2.Koneksi.Close()












                Sql100 = "SELECT * From Table_AkunArusKas1 WHERE Tahun='" & CStr(txtTahun.Text) & "' Order by Nomor_Akun asc"

                Dim da3 As New SqlDataAdapter(Sql100, conn)
                Dim ds3 As New DataSet
                da3.Fill(ds3, "Table_AkunArusKas1")
                Dim dt3 As New DataTable

                Dim z As Integer
                Dim debet100, debet102, debet103, debet104 As Integer
                Dim kredit100, kredit102, kredit103, kredit104 As Integer
                Dim dsTables3 As DataTable = ds3.Tables("Table_AkunArusKas1")





                For z = 0 To ds3.Tables("Table_AkunArusKas1").Rows.Count - 1

                    If dsTables3.Rows(z).Item("Nomor_Akun") = "111111" Then
                        Nomor_Akun100 = dsTables3.Rows(z).Item("Nomor_Akun")
                        Nama_Akun100 = dsTables3.Rows(z).Item("Nama_Akun")
                        Kelompok_Akun100 = dsTables3.Rows(z).Item("Kelompok_Akun")
                        Tanggal_Input100 = dsTables3.Rows(z).Item("Tanggal_Input")
                        'Bulan100 = dsTables3.Rows(z).Item("Bulan")
                        Tahun100 = dsTables3.Rows(z).Item("Tahun")
                        'header100 = dsTables3.Rows(z).Item("header")

                        debet100 += dsTables3.Rows(z).Item("Saldo_Debet")
                        kredit100 += dsTables3.Rows(z).Item("Saldo_Kredit")
                    ElseIf dsTables3.Rows(z).Item("Nomor_Akun") = "222222" Then
                        Nomor_Akun102 = dsTables3.Rows(z).Item("Nomor_Akun")
                        Nama_Akun102 = dsTables3.Rows(z).Item("Nama_Akun")
                        Kelompok_Akun102 = dsTables3.Rows(z).Item("Kelompok_Akun")
                        Tanggal_Input102 = dsTables3.Rows(z).Item("Tanggal_Input")
                        'Bulan100 = dsTables3.Rows(z).Item("Bulan")
                        Tahun102 = dsTables3.Rows(z).Item("Tahun")
                        'header100 = dsTables3.Rows(z).Item("header")

                        debet102 += dsTables3.Rows(z).Item("Saldo_Debet")
                        kredit102 += dsTables3.Rows(z).Item("Saldo_Kredit")
                    ElseIf dsTables3.Rows(z).Item("Nomor_Akun") = "333333" Then
                        Nomor_Akun103 = dsTables3.Rows(z).Item("Nomor_Akun")
                        Nama_Akun103 = dsTables3.Rows(z).Item("Nama_Akun")
                        Kelompok_Akun103 = IIf(IsDBNull(dsTables3.Rows(z).Item("Kelompok_Akun")), "", dsTables3.Rows(z).Item("Kelompok_Akun"))
                        Tanggal_Input103 = dsTables3.Rows(z).Item("Tanggal_Input")
                        'Bulan100 = dsTables3.Rows(z).Item("Bulan")
                        Tahun103 = dsTables3.Rows(z).Item("Tahun")
                        'header100 = dsTables3.Rows(z).Item("header")

                        debet103 += dsTables3.Rows(z).Item("Saldo_Debet")
                        kredit103 += dsTables3.Rows(z).Item("Saldo_Kredit")
                    ElseIf dsTables3.Rows(z).Item("Nomor_Akun") = "000000" Then
                        Nomor_Akun104 = dsTables3.Rows(z).Item("Nomor_Akun")
                        Nama_Akun104 = dsTables3.Rows(z).Item("Nama_Akun")
                        Kelompok_Akun104 = IIf(IsDBNull(dsTables3.Rows(z).Item("Kelompok_Akun")), "", dsTables3.Rows(z).Item("Kelompok_Akun"))
                        Tanggal_Input104 = dsTables3.Rows(z).Item("Tanggal_Input")
                        'Bulan100 = dsTables3.Rows(z).Item("Bulan")
                        Tahun104 = dsTables3.Rows(z).Item("Tahun")
                        'header100 = dsTables3.Rows(z).Item("header")

                        debet104 += dsTables3.Rows(z).Item("Saldo_Debet")
                        kredit104 += dsTables3.Rows(z).Item("Saldo_Kredit")
                    End If
                Next


                Sql101 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
               " VALUES ('" & Nomor_Akun100 & "','" & Nama_Akun100 & "','" & Kelompok_Akun100 & "','" & Tanggal_Input100 & "','" & Tahun100 & "','" & debet100 & "','" & kredit100 & "','" & header1 & "')"

                Sql102 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
              " VALUES ('" & Nomor_Akun102 & "','" & Nama_Akun102 & "','" & Kelompok_Akun102 & "','" & Tanggal_Input102 & "','" & Tahun102 & "','" & debet102 & "','" & kredit102 & "','" & header1 & "')"

                Sql103 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
             " VALUES ('" & Nomor_Akun103 & "','" & Nama_Akun103 & "','" & Kelompok_Akun103 & "','" & Tanggal_Input103 & "','" & Tahun103 & "','" & debet103 & "','" & kredit103 & "','" & header1 & "')"

                Sql104 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
            " VALUES ('" & Nomor_Akun104 & "','" & Nama_Akun104 & "','" & Kelompok_Akun104 & "','" & Tanggal_Input104 & "','" & Tahun104 & "','" & debet104 & "','" & kredit104 & "','" & header1 & "')"

                Dim command101 As New SqlCommand(Sql101, Module2.Koneksi)
                command101.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command102 As New SqlCommand(Sql102, Module2.Koneksi)
                command102.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command103 As New SqlCommand(Sql103, Module2.Koneksi)
                command103.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command104 As New SqlCommand(Sql104, Module2.Koneksi)
                command104.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Sql = "SELECT * From Table_Akun WHERE Tahun='" & CStr(txtTahun.Text) & "' Order by Nomor_Akun asc"

                Dim da2 As New SqlDataAdapter(Sql, conn)
                Dim ds2 As New DataSet
                da2.Fill(ds2, "Table_Akun")
                Dim dt As New DataTable

                Dim i As Integer
                Dim debet1, debet2, debet3, debet4, debet5, debet6, debet7, debet8, debet9, debet10, debet11, debet12, debet13, debet14, debet15, debet16, debet18, debet19, debet20, debet21, debet22, debet23, debet24, debet28, debet29, debet30, debet31, debet32, debet33, debet34, debet35, debet36, debet37, debet38, debet39, debet40, debet41, debet50, debet43, debet44, debet45, debet46, debet47, debet48, debet49, debet65, debet66, debet67, debet68, debet69, debet70, debet71, debet72, debet73, debet75 As Integer
                Dim kredit1, kredit2, kredit3, kredit4, kredit5, kredit6, kredit7, kredit8, kredit9, kredit10, kredit11, kredit12, kredit13, kredit14, kredit15, kredit16, kredit18, kredit19, kredit20, kredit21, kredit22, kredit23, kredit24, kredit28, kredit29, kredit30, kredit31, kredit32, kredit33, kredit34, kredit35, kredit36, kredit37, kredit38, kredit39, kredit40, kredit41, kredit43, kredit44, kredit45, kredit46, kredit47, kredit48, kredit49, kredit50, kredit65, kredit66, kredit67, kredit68, kredit69, kredit70, kredit71, kredit72, kredit73, kredit75 As Integer
                Dim dsTables As DataTable = ds2.Tables("Table_Akun")





                For i = 0 To ds2.Tables("Table_Akun").Rows.Count - 1



                    If dsTables.Rows(i).Item("Nomor_Akun") = "1101" Then
                        Nomor_Akun1 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun1 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun1 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input1 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan1 = dsTables.Rows(i).Item("Bulan")
                        Tahun1 = dsTables.Rows(i).Item("Tahun")
                        header1 = dsTables.Rows(i).Item("header")

                        debet1 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit1 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1102" Then
                        Nomor_Akun2 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun2 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun2 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input2 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan2 = dsTables.Rows(i).Item("Bulan")
                        Tahun2 = dsTables.Rows(i).Item("Tahun")
                        header2 = dsTables.Rows(i).Item("header")

                        debet2 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit2 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1103" Then
                        Nomor_Akun3 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun3 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun3 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input3 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan3 = dsTables.Rows(i).Item("Bulan")
                        Tahun3 = dsTables.Rows(i).Item("Tahun")
                        header3 = dsTables.Rows(i).Item("header")

                        debet3 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit3 += dsTables.Rows(i).Item("Saldo_Kredit")

                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1201" Then
                        Nomor_Akun4 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun4 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun4 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input4 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan4 = dsTables.Rows(i).Item("Bulan")
                        Tahun4 = dsTables.Rows(i).Item("Tahun")
                        header4 = dsTables.Rows(i).Item("header")

                        debet4 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit4 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1202" Then
                        Nomor_Akun5 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun5 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun5 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input5 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan5 = dsTables.Rows(i).Item("Bulan")
                        Tahun5 = dsTables.Rows(i).Item("Tahun")
                        header5 = dsTables.Rows(i).Item("header")

                        debet5 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit5 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1203" Then
                        Nomor_Akun6 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun6 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun6 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input6 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan6 = dsTables.Rows(i).Item("Bulan")
                        Tahun6 = dsTables.Rows(i).Item("Tahun")
                        header6 = dsTables.Rows(i).Item("header")

                        debet6 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit6 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1204" Then
                        Nomor_Akun7 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun7 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun7 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input7 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan7 = dsTables.Rows(i).Item("Bulan")
                        Tahun7 = dsTables.Rows(i).Item("Tahun")
                        header7 = dsTables.Rows(i).Item("header")

                        debet7 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit7 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1205" Then
                        Nomor_Akun8 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun8 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun8 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input8 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan8 = dsTables.Rows(i).Item("Bulan")
                        Tahun8 = dsTables.Rows(i).Item("Tahun")
                        header8 = dsTables.Rows(i).Item("header")

                        debet8 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit8 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1206" Then
                        Nomor_Akun9 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun9 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun9 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input9 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan9 = dsTables.Rows(i).Item("Bulan")
                        Tahun9 = dsTables.Rows(i).Item("Tahun")
                        header9 = dsTables.Rows(i).Item("header")

                        debet9 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit9 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1207" Then
                        Nomor_Akun10 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun10 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun10 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input10 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan10 = dsTables.Rows(i).Item("Bulan")
                        Tahun10 = dsTables.Rows(i).Item("Tahun")
                        header10 = dsTables.Rows(i).Item("header")

                        debet10 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit10 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1208" Then
                        Nomor_Akun11 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun11 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun11 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input11 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan11 = dsTables.Rows(i).Item("Bulan")
                        Tahun11 = dsTables.Rows(i).Item("Tahun")
                        header11 = dsTables.Rows(i).Item("header")

                        debet11 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit11 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1209" Then
                        Nomor_Akun12 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun12 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun12 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input12 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan12 = dsTables.Rows(i).Item("Bulan")
                        Tahun12 = dsTables.Rows(i).Item("Tahun")
                        header12 = dsTables.Rows(i).Item("header")

                        debet12 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit12 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1210" Then
                        Nomor_Akun13 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun13 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun13 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input13 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan13 = dsTables.Rows(i).Item("Bulan")
                        Tahun13 = dsTables.Rows(i).Item("Tahun")
                        header13 = dsTables.Rows(i).Item("header")

                        debet13 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit13 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1211" Then
                        Nomor_Akun14 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun14 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun14 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input14 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan14 = dsTables.Rows(i).Item("Bulan")
                        Tahun14 = dsTables.Rows(i).Item("Tahun")
                        header14 = dsTables.Rows(i).Item("header")

                        debet14 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit14 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1212" Then
                        Nomor_Akun15 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun15 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun15 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input15 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan15 = dsTables.Rows(i).Item("Bulan")
                        Tahun15 = dsTables.Rows(i).Item("Tahun")
                        header15 = dsTables.Rows(i).Item("header")

                        debet15 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit15 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1213" Then
                        Nomor_Akun16 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun16 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun16 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input16 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan16 = dsTables.Rows(i).Item("Bulan")
                        Tahun16 = dsTables.Rows(i).Item("Tahun")
                        header16 = dsTables.Rows(i).Item("header")

                        debet16 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit16 += dsTables.Rows(i).Item("Saldo_Kredit")
                        'ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1214" Then
                        '    Nomor_Akun17 = dsTables.Rows(i).Item("Nomor_Akun")
                        '    Nama_Akun17 = dsTables.Rows(i).Item("Nama_Akun")
                        '    Kelompok_Akun17 = dsTables.Rows(i).Item("Kelompok_Akun")
                        '    Tanggal_Input17 = dsTables.Rows(i).Item("Tanggal_Input")
                        '    Bulan17 = dsTables.Rows(i).Item("Bulan")
                        '    Tahun17 = dsTables.Rows(i).Item("Tahun")
                        '    header17 = dsTables.Rows(i).Item("header")

                        '    debet17 += dsTables.Rows(i).Item("Saldo_Debet")
                        '    kredit17 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1215" Then
                        Nomor_Akun18 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun18 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun18 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input18 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan18 = dsTables.Rows(i).Item("Bulan")
                        Tahun18 = dsTables.Rows(i).Item("Tahun")
                        header18 = dsTables.Rows(i).Item("header")

                        debet18 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit18 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1216" Then
                        Nomor_Akun19 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun19 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun19 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input19 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan19 = dsTables.Rows(i).Item("Bulan")
                        Tahun19 = dsTables.Rows(i).Item("Tahun")
                        header19 = dsTables.Rows(i).Item("header")

                        debet19 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit19 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1217" Then
                        Nomor_Akun20 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun20 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun20 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input20 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan20 = dsTables.Rows(i).Item("Bulan")
                        Tahun20 = dsTables.Rows(i).Item("Tahun")
                        header20 = dsTables.Rows(i).Item("header")

                        debet20 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit20 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1218" Then
                        Nomor_Akun21 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun21 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun21 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input21 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan21 = dsTables.Rows(i).Item("Bulan")
                        Tahun21 = dsTables.Rows(i).Item("Tahun")
                        header21 = dsTables.Rows(i).Item("header")

                        debet21 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit21 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1219" Then
                        Nomor_Akun22 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun22 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun22 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input22 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan22 = dsTables.Rows(i).Item("Bulan")
                        Tahun22 = dsTables.Rows(i).Item("Tahun")
                        header22 = dsTables.Rows(i).Item("header")

                        debet22 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit22 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1220" Then
                        Nomor_Akun23 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun23 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun23 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input23 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan23 = dsTables.Rows(i).Item("Bulan")
                        Tahun23 = dsTables.Rows(i).Item("Tahun")
                        header23 = dsTables.Rows(i).Item("header")

                        debet23 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit23 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1221" Then
                        Nomor_Akun24 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun24 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun24 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input24 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan24 = dsTables.Rows(i).Item("Bulan")
                        Tahun24 = dsTables.Rows(i).Item("Tahun")
                        header24 = dsTables.Rows(i).Item("header")

                        debet24 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit24 += dsTables.Rows(i).Item("Saldo_Kredit")
                        'ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1222" Then
                        '    Nomor_Akun25 = dsTables.Rows(i).Item("Nomor_Akun")
                        '    Nama_Akun25 = dsTables.Rows(i).Item("Nama_Akun")
                        '    Kelompok_Akun25 = dsTables.Rows(i).Item("Kelompok_Akun")
                        '    Tanggal_Input25 = dsTables.Rows(i).Item("Tanggal_Input")
                        '    Bulan25 = dsTables.Rows(i).Item("Bulan")
                        '    Tahun25 = dsTables.Rows(i).Item("Tahun")
                        '    header25 = dsTables.Rows(i).Item("header")

                        '    debet25 += dsTables.Rows(i).Item("Saldo_Debet")
                        '    kredit25 += dsTables.Rows(i).Item("Saldo_Kredit")
                        'ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1223" Then
                        '    Nomor_Akun26 = dsTables.Rows(i).Item("Nomor_Akun")
                        '    Nama_Akun26 = dsTables.Rows(i).Item("Nama_Akun")
                        '    Kelompok_Akun26 = dsTables.Rows(i).Item("Kelompok_Akun")
                        '    Tanggal_Input26 = dsTables.Rows(i).Item("Tanggal_Input")
                        '    Bulan26 = dsTables.Rows(i).Item("Bulan")
                        '    Tahun26 = dsTables.Rows(i).Item("Tahun")
                        '    header26 = dsTables.Rows(i).Item("header")

                        '    debet26 += dsTables.Rows(i).Item("Saldo_Debet")
                        '    kredit26 += dsTables.Rows(i).Item("Saldo_Kredit")
                        'ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1224" Then
                        '    Nomor_Akun27 = dsTables.Rows(i).Item("Nomor_Akun")
                        '    Nama_Akun27 = dsTables.Rows(i).Item("Nama_Akun")
                        '    Kelompok_Akun27 = dsTables.Rows(i).Item("Kelompok_Akun")
                        '    Tanggal_Input27 = dsTables.Rows(i).Item("Tanggal_Input")
                        '    Bulan27 = dsTables.Rows(i).Item("Bulan")
                        '    Tahun27 = dsTables.Rows(i).Item("Tahun")
                        '    header27 = dsTables.Rows(i).Item("header")

                        '    debet27 += dsTables.Rows(i).Item("Saldo_Debet")
                        '    kredit27 += dsTables.Rows(i).Item("Saldo_Kredit")

                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "2101" Then
                        Nomor_Akun28 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun28 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun28 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input28 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan28 = dsTables.Rows(i).Item("Bulan")
                        Tahun28 = dsTables.Rows(i).Item("Tahun")
                        header28 = dsTables.Rows(i).Item("header")

                        debet28 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit28 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "2102" Then
                        Nomor_Akun29 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun29 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun29 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input29 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan29 = dsTables.Rows(i).Item("Bulan")
                        Tahun29 = dsTables.Rows(i).Item("Tahun")
                        header29 = dsTables.Rows(i).Item("header")

                        debet29 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit29 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "2103" Then
                        Nomor_Akun30 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun30 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun30 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input30 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan30 = dsTables.Rows(i).Item("Bulan")
                        Tahun30 = dsTables.Rows(i).Item("Tahun")
                        header30 = dsTables.Rows(i).Item("header")

                        debet30 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit30 += dsTables.Rows(i).Item("Saldo_Kredit")

                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "2201" Then
                        Nomor_Akun31 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun31 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun31 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input31 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan31 = dsTables.Rows(i).Item("Bulan")
                        Tahun31 = dsTables.Rows(i).Item("Tahun")
                        header31 = dsTables.Rows(i).Item("header")

                        debet31 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit31 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "2202" Then
                        Nomor_Akun32 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun32 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun32 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input32 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan32 = dsTables.Rows(i).Item("Bulan")
                        Tahun32 = dsTables.Rows(i).Item("Tahun")
                        header32 = dsTables.Rows(i).Item("header")

                        debet32 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit32 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "2203" Then
                        Nomor_Akun33 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun33 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun33 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input33 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan33 = dsTables.Rows(i).Item("Bulan")
                        Tahun33 = dsTables.Rows(i).Item("Tahun")
                        header33 = dsTables.Rows(i).Item("header")

                        debet33 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit33 += dsTables.Rows(i).Item("Saldo_Kredit")


                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "3100" Then
                        Nomor_Akun34 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun34 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun34 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input34 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan34 = dsTables.Rows(i).Item("Bulan")
                        Tahun34 = dsTables.Rows(i).Item("Tahun")
                        header34 = dsTables.Rows(i).Item("header")

                        debet34 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit34 += dsTables.Rows(i).Item("Saldo_Kredit")

                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "3200" Then
                        Nomor_Akun35 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun35 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun35 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input35 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan35 = dsTables.Rows(i).Item("Bulan")
                        Tahun35 = dsTables.Rows(i).Item("Tahun")
                        header35 = dsTables.Rows(i).Item("header")

                        debet35 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit35 += dsTables.Rows(i).Item("Saldo_Kredit")


                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "4101" Then
                        Nomor_Akun36 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun36 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun36 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input36 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan36 = dsTables.Rows(i).Item("Bulan")
                        Tahun36 = dsTables.Rows(i).Item("Tahun")
                        header36 = dsTables.Rows(i).Item("header")

                        debet36 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit36 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "4102" Then
                        Nomor_Akun37 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun37 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun37 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input37 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan37 = dsTables.Rows(i).Item("Bulan")
                        Tahun37 = dsTables.Rows(i).Item("Tahun")
                        header37 = dsTables.Rows(i).Item("header")

                        debet37 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit37 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "4103" Then
                        Nomor_Akun38 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun38 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun38 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input38 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan38 = dsTables.Rows(i).Item("Bulan")
                        Tahun38 = dsTables.Rows(i).Item("Tahun")
                        header38 = dsTables.Rows(i).Item("header")

                        debet38 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit38 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "4104" Then
                        Nomor_Akun39 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun39 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun39 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input39 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan39 = dsTables.Rows(i).Item("Bulan")
                        Tahun39 = dsTables.Rows(i).Item("Tahun")
                        header39 = dsTables.Rows(i).Item("header")

                        debet39 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit39 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "4105" Then
                        Nomor_Akun40 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun40 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun40 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input40 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan40 = dsTables.Rows(i).Item("Bulan")
                        Tahun40 = dsTables.Rows(i).Item("Tahun")

                        header40 = dsTables.Rows(i).Item("header")
                        debet40 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit40 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "4106" Then
                        Nomor_Akun41 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun41 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun41 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input41 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan41 = dsTables.Rows(i).Item("Bulan")
                        Tahun41 = dsTables.Rows(i).Item("Tahun")

                        header41 = dsTables.Rows(i).Item("header")
                        debet41 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit41 += dsTables.Rows(i).Item("Saldo_Kredit")
                        'ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "4107" Then
                        '    Nomor_Akun42 = dsTables.Rows(i).Item("Nomor_Akun")
                        '    Nama_Akun42 = dsTables.Rows(i).Item("Nama_Akun")
                        '    Kelompok_Akun42 = dsTables.Rows(i).Item("Kelompok_Akun")
                        '    Tanggal_Input42 = dsTables.Rows(i).Item("Tanggal_Input")
                        '    Bulan42 = dsTables.Rows(i).Item("Bulan")
                        '    Tahun42 = dsTables.Rows(i).Item("Tahun")
                        '    header42 = dsTables.Rows(i).Item("header")

                        '    debet42 += dsTables.Rows(i).Item("Saldo_Debet")
                        '    kredit42 += dsTables.Rows(i).Item("Saldo_Kredit")


                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "5101" Then
                        Nomor_Akun43 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun43 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun43 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input43 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan43 = dsTables.Rows(i).Item("Bulan")
                        Tahun43 = dsTables.Rows(i).Item("Tahun")

                        header43 = dsTables.Rows(i).Item("header")
                        debet43 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit43 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "5102" Then
                        Nomor_Akun44 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun44 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun44 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input44 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan44 = dsTables.Rows(i).Item("Bulan")
                        Tahun44 = dsTables.Rows(i).Item("Tahun")

                        header44 = dsTables.Rows(i).Item("header")
                        debet44 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit44 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "5103" Then
                        Nomor_Akun45 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun45 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun45 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input45 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan45 = dsTables.Rows(i).Item("Bulan")
                        Tahun45 = dsTables.Rows(i).Item("Tahun")
                        header45 = dsTables.Rows(i).Item("header")

                        debet45 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit45 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "5104" Then
                        Nomor_Akun46 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun46 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun46 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input46 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan46 = dsTables.Rows(i).Item("Bulan")
                        Tahun46 = dsTables.Rows(i).Item("Tahun")

                        header46 = dsTables.Rows(i).Item("header")
                        debet46 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit46 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "5105" Then
                        Nomor_Akun47 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun47 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun47 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input47 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan47 = dsTables.Rows(i).Item("Bulan")
                        Tahun47 = dsTables.Rows(i).Item("Tahun")
                        header47 = dsTables.Rows(i).Item("header")

                        debet47 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit47 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "5106" Then
                        Nomor_Akun48 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun48 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun48 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input48 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan48 = dsTables.Rows(i).Item("Bulan")
                        Tahun48 = dsTables.Rows(i).Item("Tahun")
                        header48 = dsTables.Rows(i).Item("header")

                        debet48 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit48 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "5107" Then
                        Nomor_Akun49 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun49 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun49 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input49 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan49 = dsTables.Rows(i).Item("Bulan")
                        Tahun49 = dsTables.Rows(i).Item("Tahun")
                        header49 = dsTables.Rows(i).Item("header")

                        debet49 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit49 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "5108" Then
                        Nomor_Akun50 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun50 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun50 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input50 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan50 = dsTables.Rows(i).Item("Bulan")
                        Tahun50 = dsTables.Rows(i).Item("Tahun")

                        header50 = dsTables.Rows(i).Item("header")
                        debet50 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit50 += dsTables.Rows(i).Item("Saldo_Kredit")
                        'ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "5109" Then
                        '    Nomor_Akun51 = dsTables.Rows(i).Item("Nomor_Akun")
                        '    Nama_Akun51 = dsTables.Rows(i).Item("Nama_Akun")
                        '    Kelompok_Akun51 = dsTables.Rows(i).Item("Kelompok_Akun")
                        '    Tanggal_Input51 = dsTables.Rows(i).Item("Tanggal_Input")
                        '    Bulan51 = dsTables.Rows(i).Item("Bulan")
                        '    Tahun51 = dsTables.Rows(i).Item("Tahun")
                        '    header51 = dsTables.Rows(i).Item("header")

                        '    debet51 += dsTables.Rows(i).Item("Saldo_Debet")
                        '    kredit51 += dsTables.Rows(i).Item("Saldo_Kredit")
                        'ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "5110" Then
                        '    Nomor_Akun52 = dsTables.Rows(i).Item("Nomor_Akun")
                        '    Nama_Akun52 = dsTables.Rows(i).Item("Nama_Akun")
                        '    Kelompok_Akun52 = dsTables.Rows(i).Item("Kelompok_Akun")
                        '    Tanggal_Input52 = dsTables.Rows(i).Item("Tanggal_Input")
                        '    Bulan52 = dsTables.Rows(i).Item("Bulan")
                        '    Tahun52 = dsTables.Rows(i).Item("Tahun")
                        '    header52 = dsTables.Rows(i).Item("header")

                        '    debet52 += dsTables.Rows(i).Item("Saldo_Debet")
                        '    kredit52 += dsTables.Rows(i).Item("Saldo_Kredit")
                        'ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "5111" Then
                        '    Nomor_Akun53 = dsTables.Rows(i).Item("Nomor_Akun")
                        '    Nama_Akun53 = dsTables.Rows(i).Item("Nama_Akun")
                        '    Kelompok_Akun53 = dsTables.Rows(i).Item("Kelompok_Akun")
                        '    Tanggal_Input53 = dsTables.Rows(i).Item("Tanggal_Input")
                        '    Bulan53 = dsTables.Rows(i).Item("Bulan")
                        '    Tahun53 = dsTables.Rows(i).Item("Tahun")
                        '    header53 = dsTables.Rows(i).Item("header")

                        '    debet53 += dsTables.Rows(i).Item("Saldo_Debet")
                        '    kredit53 += dsTables.Rows(i).Item("Saldo_Kredit")
                        'ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "5112" Then
                        '    Nomor_Akun54 = dsTables.Rows(i).Item("Nomor_Akun")
                        '    Nama_Akun54 = dsTables.Rows(i).Item("Nama_Akun")
                        '    Kelompok_Akun54 = dsTables.Rows(i).Item("Kelompok_Akun")
                        '    Tanggal_Input54 = dsTables.Rows(i).Item("Tanggal_Input")
                        '    Bulan54 = dsTables.Rows(i).Item("Bulan")
                        '    Tahun54 = dsTables.Rows(i).Item("Tahun")
                        '    header54 = dsTables.Rows(i).Item("header")

                        '    debet54 += dsTables.Rows(i).Item("Saldo_Debet")
                        '    kredit54 += dsTables.Rows(i).Item("Saldo_Kredit")
                        'ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "5113" Then
                        '    Nomor_Akun55 = dsTables.Rows(i).Item("Nomor_Akun")
                        '    Nama_Akun55 = dsTables.Rows(i).Item("Nama_Akun")
                        '    Kelompok_Akun55 = dsTables.Rows(i).Item("Kelompok_Akun")
                        '    Tanggal_Input55 = dsTables.Rows(i).Item("Tanggal_Input")
                        '    Bulan55 = dsTables.Rows(i).Item("Bulan")
                        '    Tahun55 = dsTables.Rows(i).Item("Tahun")

                        '    header55 = dsTables.Rows(i).Item("header")
                        '    debet55 += dsTables.Rows(i).Item("Saldo_Debet")
                        '    kredit55 += dsTables.Rows(i).Item("Saldo_Kredit")
                        'ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "5114" Then
                        '    Nomor_Akun56 = dsTables.Rows(i).Item("Nomor_Akun")
                        '    Nama_Akun56 = dsTables.Rows(i).Item("Nama_Akun")
                        '    Kelompok_Akun56 = dsTables.Rows(i).Item("Kelompok_Akun")
                        '    Tanggal_Input56 = dsTables.Rows(i).Item("Tanggal_Input")
                        '    Bulan56 = dsTables.Rows(i).Item("Bulan")
                        '    Tahun56 = dsTables.Rows(i).Item("Tahun")
                        '    header56 = dsTables.Rows(i).Item("header")

                        '    debet56 += dsTables.Rows(i).Item("Saldo_Debet")
                        '    kredit56 += dsTables.Rows(i).Item("Saldo_Kredit")
                        'ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "5115" Then
                        '    Nomor_Akun57 = dsTables.Rows(i).Item("Nomor_Akun")
                        '    Nama_Akun57 = dsTables.Rows(i).Item("Nama_Akun")
                        '    Kelompok_Akun57 = dsTables.Rows(i).Item("Kelompok_Akun")
                        '    Tanggal_Input57 = dsTables.Rows(i).Item("Tanggal_Input")
                        '    Bulan57 = dsTables.Rows(i).Item("Bulan")
                        '    Tahun57 = dsTables.Rows(i).Item("Tahun")
                        '    header57 = dsTables.Rows(i).Item("header")

                        '    debet57 += dsTables.Rows(i).Item("Saldo_Debet")
                        '    kredit57 += dsTables.Rows(i).Item("Saldo_Kredit")
                        'ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "5116" Then
                        '    Nomor_Akun58 = dsTables.Rows(i).Item("Nomor_Akun")
                        '    Nama_Akun58 = dsTables.Rows(i).Item("Nama_Akun")
                        '    Kelompok_Akun58 = dsTables.Rows(i).Item("Kelompok_Akun")
                        '    Tanggal_Input58 = dsTables.Rows(i).Item("Tanggal_Input")
                        '    Bulan58 = dsTables.Rows(i).Item("Bulan")
                        '    Tahun58 = dsTables.Rows(i).Item("Tahun")
                        '    header58 = dsTables.Rows(i).Item("header")

                        '    debet58 += dsTables.Rows(i).Item("Saldo_Debet")
                        '    kredit58 += dsTables.Rows(i).Item("Saldo_Kredit")
                        'ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "5117" Then
                        '    Nomor_Akun59 = dsTables.Rows(i).Item("Nomor_Akun")
                        '    Nama_Akun59 = dsTables.Rows(i).Item("Nama_Akun")
                        '    Kelompok_Akun59 = dsTables.Rows(i).Item("Kelompok_Akun")
                        '    Tanggal_Input59 = dsTables.Rows(i).Item("Tanggal_Input")
                        '    Bulan59 = dsTables.Rows(i).Item("Bulan")
                        '    Tahun59 = dsTables.Rows(i).Item("Tahun")

                        '    header59 = dsTables.Rows(i).Item("header")
                        '    debet59 += dsTables.Rows(i).Item("Saldo_Debet")
                        '    kredit59 += dsTables.Rows(i).Item("Saldo_Kredit")
                        'ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "5118" Then
                        '    Nomor_Akun60 = dsTables.Rows(i).Item("Nomor_Akun")
                        '    Nama_Akun60 = dsTables.Rows(i).Item("Nama_Akun")
                        '    Kelompok_Akun60 = dsTables.Rows(i).Item("Kelompok_Akun")
                        '    Tanggal_Input60 = dsTables.Rows(i).Item("Tanggal_Input")
                        '    Bulan60 = dsTables.Rows(i).Item("Bulan")
                        '    Tahun60 = dsTables.Rows(i).Item("Tahun")

                        '    header60 = dsTables.Rows(i).Item("header")
                        '    debet60 += dsTables.Rows(i).Item("Saldo_Debet")
                        '    kredit60 += dsTables.Rows(i).Item("Saldo_Kredit")
                        'ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "5119" Then
                        '    Nomor_Akun61 = dsTables.Rows(i).Item("Nomor_Akun")
                        '    Nama_Akun61 = dsTables.Rows(i).Item("Nama_Akun")
                        '    Kelompok_Akun61 = dsTables.Rows(i).Item("Kelompok_Akun")
                        '    Tanggal_Input61 = dsTables.Rows(i).Item("Tanggal_Input")
                        '    Bulan61 = dsTables.Rows(i).Item("Bulan")
                        '    Tahun61 = dsTables.Rows(i).Item("Tahun")
                        '    header61 = dsTables.Rows(i).Item("header")

                        '    debet61 += dsTables.Rows(i).Item("Saldo_Debet")
                        '    kredit61 += dsTables.Rows(i).Item("Saldo_Kredit")
                        'ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "5120" Then
                        '    Nomor_Akun62 = dsTables.Rows(i).Item("Nomor_Akun")
                        '    Nama_Akun62 = dsTables.Rows(i).Item("Nama_Akun")
                        '    Kelompok_Akun62 = dsTables.Rows(i).Item("Kelompok_Akun")
                        '    Tanggal_Input62 = dsTables.Rows(i).Item("Tanggal_Input")
                        '    Bulan62 = dsTables.Rows(i).Item("Bulan")
                        '    Tahun62 = dsTables.Rows(i).Item("Tahun")

                        '    header62 = dsTables.Rows(i).Item("header")
                        '    debet62 += dsTables.Rows(i).Item("Saldo_Debet")
                        '    kredit62 += dsTables.Rows(i).Item("Saldo_Kredit")
                        'ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "5121" Then
                        '    Nomor_Akun63 = dsTables.Rows(i).Item("Nomor_Akun")
                        '    Nama_Akun63 = dsTables.Rows(i).Item("Nama_Akun")
                        '    Kelompok_Akun63 = dsTables.Rows(i).Item("Kelompok_Akun")
                        '    Tanggal_Input63 = dsTables.Rows(i).Item("Tanggal_Input")
                        '    Bulan63 = dsTables.Rows(i).Item("Bulan")
                        '    Tahun63 = dsTables.Rows(i).Item("Tahun")

                        '    header63 = dsTables.Rows(i).Item("header")
                        '    debet63 += dsTables.Rows(i).Item("Saldo_Debet")
                        '    kredit63 += dsTables.Rows(i).Item("Saldo_Kredit")
                        'ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "5122" Then
                        '    Nomor_Akun64 = dsTables.Rows(i).Item("Nomor_Akun")
                        '    Nama_Akun64 = dsTables.Rows(i).Item("Nama_Akun")
                        '    Kelompok_Akun64 = dsTables.Rows(i).Item("Kelompok_Akun")
                        '    Tanggal_Input64 = dsTables.Rows(i).Item("Tanggal_Input")
                        '    Bulan64 = dsTables.Rows(i).Item("Bulan")
                        '    Tahun64 = dsTables.Rows(i).Item("Tahun")

                        '    header64 = dsTables.Rows(i).Item("header")
                        '    debet64 += dsTables.Rows(i).Item("Saldo_Debet")
                        '    kredit64 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "3300" Then
                        Nomor_Akun65 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun65 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun65 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input65 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan65 = dsTables.Rows(i).Item("Bulan")
                        Tahun65 = dsTables.Rows(i).Item("Tahun")
                        header65 = dsTables.Rows(i).Item("header")

                        debet65 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit65 += dsTables.Rows(i).Item("Saldo_Kredit")
                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "3400" Then
                        Nomor_Akun66 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun66 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun66 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input66 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan66 = dsTables.Rows(i).Item("Bulan")
                        Tahun66 = dsTables.Rows(i).Item("Tahun")
                        header66 = dsTables.Rows(i).Item("header")

                        debet66 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit66 += dsTables.Rows(i).Item("Saldo_Kredit")











                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1000" Then
                        Nomor_Akun67 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun67 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun67 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input67 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan67 = dsTables.Rows(i).Item("Bulan")
                        Tahun67 = dsTables.Rows(i).Item("Tahun")
                        header67 = dsTables.Rows(i).Item("header")

                        debet67 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit67 += dsTables.Rows(i).Item("Saldo_Kredit")

                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1100" Then
                        Nomor_Akun68 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun68 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun68 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input68 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan68 = dsTables.Rows(i).Item("Bulan")
                        Tahun68 = dsTables.Rows(i).Item("Tahun")
                        header68 = dsTables.Rows(i).Item("header")

                        debet68 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit68 += dsTables.Rows(i).Item("Saldo_Kredit")

                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "1200" Then
                        Nomor_Akun69 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun69 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun69 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input69 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan69 = dsTables.Rows(i).Item("Bulan")
                        Tahun69 = dsTables.Rows(i).Item("Tahun")
                        header69 = dsTables.Rows(i).Item("header")

                        debet69 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit69 += dsTables.Rows(i).Item("Saldo_Kredit")

                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "2000" Then
                        Nomor_Akun70 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun70 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun70 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input70 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan70 = dsTables.Rows(i).Item("Bulan")
                        Tahun70 = dsTables.Rows(i).Item("Tahun")
                        header70 = dsTables.Rows(i).Item("header")

                        debet70 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit70 += dsTables.Rows(i).Item("Saldo_Kredit")

                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "2100" Then
                        Nomor_Akun71 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun71 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun71 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input71 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan71 = dsTables.Rows(i).Item("Bulan")
                        Tahun71 = dsTables.Rows(i).Item("Tahun")
                        header71 = dsTables.Rows(i).Item("header")

                        debet71 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit71 += dsTables.Rows(i).Item("Saldo_Kredit")

                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "2200" Then
                        Nomor_Akun72 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun72 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun72 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input72 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan72 = dsTables.Rows(i).Item("Bulan")
                        Tahun72 = dsTables.Rows(i).Item("Tahun")
                        header72 = dsTables.Rows(i).Item("header")

                        debet72 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit72 += dsTables.Rows(i).Item("Saldo_Kredit")

                    ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "3000" Then
                        Nomor_Akun73 = dsTables.Rows(i).Item("Nomor_Akun")
                        Nama_Akun73 = dsTables.Rows(i).Item("Nama_Akun")
                        Kelompok_Akun73 = dsTables.Rows(i).Item("Kelompok_Akun")
                        Tanggal_Input73 = dsTables.Rows(i).Item("Tanggal_Input")
                        Bulan73 = dsTables.Rows(i).Item("Bulan")
                        Tahun73 = dsTables.Rows(i).Item("Tahun")
                        header73 = dsTables.Rows(i).Item("header")

                        debet73 += dsTables.Rows(i).Item("Saldo_Debet")
                        kredit73 += dsTables.Rows(i).Item("Saldo_Kredit")

                        'ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "4000" Then
                        '    Nomor_Akun74 = dsTables.Rows(i).Item("Nomor_Akun")
                        '    Nama_Akun74 = dsTables.Rows(i).Item("Nama_Akun")
                        '    Kelompok_Akun74 = dsTables.Rows(i).Item("Kelompok_Akun")
                        '    Tanggal_Input74 = dsTables.Rows(i).Item("Tanggal_Input")
                        '    Bulan74 = dsTables.Rows(i).Item("Bulan")
                        '    Tahun74 = dsTables.Rows(i).Item("Tahun")
                        '    header74 = dsTables.Rows(i).Item("header")

                        '    debet74 += dsTables.Rows(i).Item("Saldo_Debet")
                        '    kredit74 += dsTables.Rows(i).Item("Saldo_Kredit")


                        'ElseIf dsTables.Rows(i).Item("Nomor_Akun") = "5000" Then
                        '    Nomor_Akun75 = dsTables.Rows(i).Item("Nomor_Akun")
                        '    Nama_Akun75 = dsTables.Rows(i).Item("Nama_Akun")
                        '    Kelompok_Akun75 = dsTables.Rows(i).Item("Kelompok_Akun")
                        '    Tanggal_Input75 = dsTables.Rows(i).Item("Tanggal_Input")
                        '    Bulan75 = dsTables.Rows(i).Item("Bulan")
                        '    Tahun75 = dsTables.Rows(i).Item("Tahun")
                        '    header75 = dsTables.Rows(i).Item("header")

                        '    debet75 += dsTables.Rows(i).Item("Saldo_Debet")
                        '    kredit75 += dsTables.Rows(i).Item("Saldo_Kredit")










                    End If

                    'MsgBox(Total_debet)
                    'Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Bulan, Tahun, Saldo_Debet, Saldo_Kredit, Header



                Next




                Sql1 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun1 & "','" & Nama_Akun1 & "','" & Kelompok_Akun1 & "','" & Tanggal_Input1 & "','" & Tahun1 & "','" & debet1 & "','" & kredit1 & "','" & header1 & "')"

                Sql2 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun2 & "','" & Nama_Akun2 & "','" & Kelompok_Akun2 & "','" & Tanggal_Input1 & "','" & Tahun2 & "','" & debet2 & "','" & kredit2 & "','" & header2 & "')"

                Sql3 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun3 & "','" & Nama_Akun3 & "','" & Kelompok_Akun3 & "','" & Tanggal_Input1 & "','" & Tahun3 & "','" & debet3 & "','" & kredit3 & "','" & header3 & "')"

                Sql4 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun4 & "','" & Nama_Akun4 & "','" & Kelompok_Akun4 & "','" & Tanggal_Input1 & "','" & Tahun4 & "','" & debet4 & "','" & kredit4 & "','" & header4 & "')"

                Sql5 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun5 & "','" & Nama_Akun5 & "','" & Kelompok_Akun5 & "','" & Tanggal_Input1 & "','" & Tahun5 & "','" & debet5 & "','" & kredit5 & "','" & header5 & "')"

                Sql6 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun6 & "','" & Nama_Akun6 & "','" & Kelompok_Akun6 & "','" & Tanggal_Input1 & "','" & Tahun6 & "','" & debet6 & "','" & kredit6 & "','" & header6 & "')"

                Sql7 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun7 & "','" & Nama_Akun7 & "','" & Kelompok_Akun7 & "','" & Tanggal_Input1 & "','" & Tahun7 & "','" & debet7 & "','" & kredit7 & "','" & header7 & "')"

                Sql8 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun8 & "','" & Nama_Akun8 & "','" & Kelompok_Akun8 & "','" & Tanggal_Input1 & "','" & Tahun8 & "','" & debet8 & "','" & kredit8 & "','" & header8 & "')"

                Sql9 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun9 & "','" & Nama_Akun9 & "','" & Kelompok_Akun9 & "','" & Tanggal_Input1 & "','" & Tahun9 & "','" & debet9 & "','" & kredit9 & "','" & header9 & "')"

                Sql10 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun10 & "','" & Nama_Akun10 & "','" & Kelompok_Akun10 & "','" & Tanggal_Input1 & "','" & Tahun10 & "','" & debet10 & "','" & kredit10 & "','" & header10 & "')"

                Sql11 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun11 & "','" & Nama_Akun11 & "','" & Kelompok_Akun11 & "','" & Tanggal_Input1 & "','" & Tahun11 & "','" & debet11 & "','" & kredit11 & "','" & header11 & "')"

                Sql12 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun12 & "','" & Nama_Akun12 & "','" & Kelompok_Akun12 & "','" & Tanggal_Input1 & "','" & Tahun12 & "','" & debet12 & "','" & kredit12 & "','" & header12 & "')"

                Sql13 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun13 & "','" & Nama_Akun13 & "','" & Kelompok_Akun13 & "','" & Tanggal_Input1 & "','" & Tahun13 & "','" & debet13 & "','" & kredit13 & "','" & header13 & "')"

                Sql14 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun14 & "','" & Nama_Akun14 & "','" & Kelompok_Akun14 & "','" & Tanggal_Input1 & "','" & Tahun14 & "','" & debet14 & "','" & kredit14 & "','" & header14 & "')"

                Sql15 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun15 & "','" & Nama_Akun15 & "','" & Kelompok_Akun15 & "','" & Tanggal_Input1 & "','" & Tahun15 & "','" & debet15 & "','" & kredit15 & "','" & header15 & "')"

                Sql16 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun16 & "','" & Nama_Akun16 & "','" & Kelompok_Akun16 & "','" & Tanggal_Input1 & "','" & Tahun16 & "','" & debet16 & "','" & kredit16 & "','" & header16 & "')"

                'Sql17 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                '" VALUES ('" & Nomor_Akun17 & "','" & Nama_Akun17 & "','" & Kelompok_Akun17 & "','" & Tanggal_Input1 & "','" & Tahun17 & "','" & header17 & "','" & debet17 & "','" & kredit17 & "')"

                Sql18 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun18 & "','" & Nama_Akun18 & "','" & Kelompok_Akun18 & "','" & Tanggal_Input1 & "','" & Tahun18 & "','" & debet18 & "','" & kredit18 & "','" & header18 & "')"

                Sql19 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun19 & "','" & Nama_Akun19 & "','" & Kelompok_Akun19 & "','" & Tanggal_Input1 & "','" & Tahun19 & "','" & debet19 & "','" & kredit19 & "','" & header19 & "')"

                Sql20 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun20 & "','" & Nama_Akun20 & "','" & Kelompok_Akun20 & "','" & Tanggal_Input1 & "','" & Tahun20 & "','" & debet20 & "','" & kredit20 & "','" & header20 & "')"

                Sql21 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun21 & "','" & Nama_Akun21 & "','" & Kelompok_Akun21 & "','" & Tanggal_Input1 & "','" & Tahun21 & "','" & debet21 & "','" & kredit21 & "','" & header21 & "')"

                Sql22 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun22 & "','" & Nama_Akun22 & "','" & Kelompok_Akun22 & "','" & Tanggal_Input1 & "','" & Tahun22 & "','" & debet22 & "','" & kredit22 & "','" & header22 & "')"

                Sql23 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun23 & "','" & Nama_Akun23 & "','" & Kelompok_Akun23 & "','" & Tanggal_Input1 & "','" & Tahun23 & "','" & debet23 & "','" & kredit23 & "','" & header23 & "')"

                Sql24 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun24 & "','" & Nama_Akun24 & "','" & Kelompok_Akun24 & "','" & Tanggal_Input1 & "','" & Tahun24 & "','" & debet24 & "','" & kredit24 & "','" & header24 & "')"

                'Sql25 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                '" VALUES ('" & Nomor_Akun25 & "','" & Nama_Akun25 & "','" & Kelompok_Akun25 & "','" & Tanggal_Input1 & "','" & Tahun25 & "','" & header25 & "','" & debet25 & "','" & kredit25 & "')"

                'Sql26 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                '" VALUES ('" & Nomor_Akun26 & "','" & Nama_Akun26 & "','" & Kelompok_Akun26 & "','" & Tanggal_Input1 & "','" & Tahun26 & "','" & header26 & "','" & debet26 & "','" & kredit26 & "')"

                'Sql27 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                '" VALUES ('" & Nomor_Akun27 & "','" & Nama_Akun27 & "','" & Kelompok_Akun27 & "','" & Tanggal_Input1 & "','" & Tahun27 & "','" & header27 & "','" & debet27 & "','" & kredit27 & "')"

                Sql28 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun28 & "','" & Nama_Akun28 & "','" & Kelompok_Akun28 & "','" & Tanggal_Input1 & "','" & Tahun28 & "','" & debet28 & "','" & kredit28 & "','" & header28 & "')"

                Sql29 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun29 & "','" & Nama_Akun29 & "','" & Kelompok_Akun29 & "','" & Tanggal_Input1 & "','" & Tahun29 & "','" & debet29 & "','" & kredit29 & "','" & header29 & "')"

                Sql30 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun30 & "','" & Nama_Akun30 & "','" & Kelompok_Akun30 & "','" & Tanggal_Input1 & "','" & Tahun30 & "','" & debet30 & "','" & kredit30 & "','" & header30 & "')"

                Sql31 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun31 & "','" & Nama_Akun31 & "','" & Kelompok_Akun31 & "','" & Tanggal_Input1 & "','" & Tahun31 & "','" & debet31 & "','" & kredit31 & "','" & header31 & "')"

                Sql32 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun32 & "','" & Nama_Akun32 & "','" & Kelompok_Akun32 & "','" & Tanggal_Input1 & "','" & Tahun32 & "','" & debet32 & "','" & kredit32 & "','" & header32 & "')"

                Sql33 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun33 & "','" & Nama_Akun33 & "','" & Kelompok_Akun33 & "','" & Tanggal_Input1 & "','" & Tahun33 & "','" & debet33 & "','" & kredit33 & "','" & header33 & "')"

                Sql34 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun34 & "','" & Nama_Akun34 & "','" & Kelompok_Akun34 & "','" & Tanggal_Input1 & "','" & Tahun34 & "','" & debet34 & "','" & kredit34 & "','" & header34 & "')"

                Sql35 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun35 & "','" & Nama_Akun35 & "','" & Kelompok_Akun35 & "','" & Tanggal_Input1 & "','" & Tahun35 & "','" & debet35 & "','" & kredit35 & "','" & header35 & "')"

                Sql36 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun36 & "','" & Nama_Akun36 & "','" & Kelompok_Akun36 & "','" & Tanggal_Input1 & "','" & Tahun36 & "','" & debet36 & "','" & kredit36 & "','" & header36 & "')"

                Sql37 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun37 & "','" & Nama_Akun37 & "','" & Kelompok_Akun37 & "','" & Tanggal_Input1 & "','" & Tahun37 & "','" & debet37 & "','" & kredit37 & "','" & header37 & "')"

                Sql38 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun38 & "','" & Nama_Akun38 & "','" & Kelompok_Akun38 & "','" & Tanggal_Input1 & "','" & Tahun38 & "','" & debet38 & "','" & kredit38 & "','" & header38 & "')"

                Sql39 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun39 & "','" & Nama_Akun39 & "','" & Kelompok_Akun39 & "','" & Tanggal_Input1 & "','" & Tahun39 & "','" & debet39 & "','" & kredit39 & "','" & header39 & "')"

                Sql40 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun40 & "','" & Nama_Akun40 & "','" & Kelompok_Akun40 & "','" & Tanggal_Input1 & "','" & Tahun40 & "','" & debet40 & "','" & kredit40 & "','" & header40 & "')"

                Sql41 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun41 & "','" & Nama_Akun41 & "','" & Kelompok_Akun41 & "','" & Tanggal_Input1 & "','" & Tahun41 & "','" & debet41 & "','" & kredit41 & "','" & header41 & "')"

                'Sql42 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                '" VALUES ('" & Nomor_Akun42 & "','" & Nama_Akun42 & "','" & Kelompok_Akun42 & "','" & Tanggal_Input1 & "','" & Tahun42 & "','" & debet42 & "','" & kredit42 & "','" & header42 & "')"

                Sql43 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun43 & "','" & Nama_Akun43 & "','" & Kelompok_Akun43 & "','" & Tanggal_Input1 & "','" & Tahun43 & "','" & debet43 & "','" & kredit43 & "','" & header43 & "')"

                Sql44 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun44 & "','" & Nama_Akun44 & "','" & Kelompok_Akun44 & "','" & Tanggal_Input1 & "','" & Tahun44 & "','" & debet44 & "','" & kredit44 & "','" & header44 & "')"

                Sql45 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun45 & "','" & Nama_Akun45 & "','" & Kelompok_Akun45 & "','" & Tanggal_Input1 & "','" & Tahun45 & "','" & debet45 & "','" & kredit45 & "','" & header45 & "')"

                Sql46 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun46 & "','" & Nama_Akun46 & "','" & Kelompok_Akun46 & "','" & Tanggal_Input1 & "','" & Tahun46 & "','" & debet46 & "','" & kredit46 & "','" & header46 & "')"

                Sql47 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun47 & "','" & Nama_Akun47 & "','" & Kelompok_Akun47 & "','" & Tanggal_Input1 & "','" & Tahun47 & "','" & debet47 & "','" & kredit47 & "','" & header47 & "')"

                Sql48 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun48 & "','" & Nama_Akun48 & "','" & Kelompok_Akun48 & "','" & Tanggal_Input1 & "','" & Tahun48 & "','" & debet48 & "','" & kredit48 & "','" & header48 & "')"

                Sql49 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun49 & "','" & Nama_Akun49 & "','" & Kelompok_Akun49 & "','" & Tanggal_Input1 & "','" & Tahun49 & "','" & debet49 & "','" & kredit49 & "','" & header49 & "')"

                Sql50 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun50 & "','" & Nama_Akun50 & "','" & Kelompok_Akun50 & "','" & Tanggal_Input1 & "','" & Tahun50 & "','" & debet50 & "','" & kredit50 & "','" & header50 & "')"

                'Sql51 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                '" VALUES ('" & Nomor_Akun51 & "','" & Nama_Akun51 & "','" & Kelompok_Akun51 & "','" & Tanggal_Input1 & "','" & Tahun51 & "','" & debet51 & "','" & kredit51 & "','" & header51 & "')"

                'Sql52 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                '" VALUES ('" & Nomor_Akun52 & "','" & Nama_Akun52 & "','" & Kelompok_Akun52 & "','" & Tanggal_Input1 & "','" & Tahun52 & "','" & debet52 & "','" & kredit52 & "','" & header52 & "')"

                'Sql53 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                '" VALUES ('" & Nomor_Akun53 & "','" & Nama_Akun53 & "','" & Kelompok_Akun53 & "','" & Tanggal_Input1 & "','" & Tahun53 & "','" & debet53 & "','" & kredit53 & "','" & header53 & "')"

                'Sql54 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                '" VALUES ('" & Nomor_Akun54 & "','" & Nama_Akun54 & "','" & Kelompok_Akun54 & "','" & Tanggal_Input1 & "','" & Tahun54 & "','" & debet54 & "','" & kredit54 & "','" & header54 & "')"

                'Sql55 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                '" VALUES ('" & Nomor_Akun55 & "','" & Nama_Akun55 & "','" & Kelompok_Akun55 & "','" & Tanggal_Input1 & "','" & Tahun55 & "','" & debet55 & "','" & kredit55 & "','" & header55 & "')"

                'Sql56 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                '" VALUES ('" & Nomor_Akun56 & "','" & Nama_Akun56 & "','" & Kelompok_Akun56 & "','" & Tanggal_Input1 & "','" & Tahun56 & "','" & debet56 & "','" & kredit56 & "','" & header56 & "')"

                'Sql57 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                '" VALUES ('" & Nomor_Akun57 & "','" & Nama_Akun57 & "','" & Kelompok_Akun57 & "','" & Tanggal_Input1 & "','" & Tahun57 & "','" & debet57 & "','" & kredit57 & "','" & header57 & "')"

                'Sql58 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                '" VALUES ('" & Nomor_Akun58 & "','" & Nama_Akun58 & "','" & Kelompok_Akun58 & "','" & Tanggal_Input1 & "','" & Tahun58 & "','" & debet58 & "','" & kredit58 & "','" & header58 & "')"

                'Sql59 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                '" VALUES ('" & Nomor_Akun59 & "','" & Nama_Akun59 & "','" & Kelompok_Akun59 & "','" & Tanggal_Input1 & "','" & Tahun59 & "','" & debet59 & "','" & kredit59 & "','" & header59 & "')"

                'Sql60 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                '" VALUES ('" & Nomor_Akun60 & "','" & Nama_Akun60 & "','" & Kelompok_Akun60 & "','" & Tanggal_Input1 & "','" & Tahun60 & "','" & debet60 & "','" & kredit60 & "''" & header60 & "',)"

                'Sql61 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                '" VALUES ('" & Nomor_Akun61 & "','" & Nama_Akun61 & "','" & Kelompok_Akun61 & "','" & Tanggal_Input1 & "','" & Tahun61 & "','" & debet61 & "','" & kredit61 & "''" & header61 & "',)"

                'Sql62 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                '" VALUES ('" & Nomor_Akun62 & "','" & Nama_Akun62 & "','" & Kelompok_Akun62 & "','" & Tanggal_Input1 & "','" & Tahun62 & "','" & debet62 & "','" & kredit62 & "''" & header62 & "',)"

                'Sql63 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                '" VALUES ('" & Nomor_Akun63 & "','" & Nama_Akun63 & "','" & Kelompok_Akun63 & "','" & Tanggal_Input1 & "','" & Tahun63 & "','" & debet63 & "','" & kredit63 & "'" & header63 & "',')"

                'Sql64 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                '" VALUES ('" & Nomor_Akun64 & "','" & Nama_Akun64 & "','" & Kelompok_Akun64 & "','" & Tanggal_Input1 & "','" & Tahun64 & "','" & debet64 & "','" & kredit64 & "''" & header64 & "',)"

                Sql65 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun65 & "','" & Nama_Akun65 & "','" & Kelompok_Akun65 & "','" & Tanggal_Input1 & "','" & Tahun65 & "','" & debet65 & "','" & kredit65 & "','" & header65 & "')"

                Sql66 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun66 & "','" & Nama_Akun66 & "','" & Kelompok_Akun66 & "','" & Tanggal_Input1 & "','" & Tahun66 & "','" & debet66 & "','" & kredit66 & "','" & header66 & "')"




                Sql67 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
               " VALUES ('" & Nomor_Akun67 & "','" & Nama_Akun67 & "','" & Kelompok_Akun67 & "','" & Tanggal_Input1 & "','" & Tahun67 & "','" & debet67 & "','" & kredit67 & "','" & header67 & "')"

                Sql68 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun68 & "','" & Nama_Akun68 & "','" & Kelompok_Akun68 & "','" & Tanggal_Input1 & "','" & Tahun68 & "','" & debet68 & "','" & kredit68 & "','" & header68 & "')"

                Sql69 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun69 & "','" & Nama_Akun69 & "','" & Kelompok_Akun69 & "','" & Tanggal_Input1 & "','" & Tahun69 & "','" & debet69 & "','" & kredit69 & "','" & header69 & "')"

                Sql70 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun70 & "','" & Nama_Akun70 & "','" & Kelompok_Akun70 & "','" & Tanggal_Input1 & "','" & Tahun70 & "','" & debet70 & "','" & kredit70 & "','" & header70 & "')"

                Sql71 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun71 & "','" & Nama_Akun71 & "','" & Kelompok_Akun71 & "','" & Tanggal_Input1 & "','" & Tahun71 & "','" & debet71 & "','" & kredit71 & "','" & header71 & "')"

                Sql72 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun72 & "','" & Nama_Akun72 & "','" & Kelompok_Akun72 & "','" & Tanggal_Input1 & "','" & Tahun72 & "','" & debet72 & "','" & kredit72 & "','" & header72 & "')"

                Sql73 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun73 & "','" & Nama_Akun73 & "','" & Kelompok_Akun73 & "','" & Tanggal_Input1 & "','" & Tahun73 & "','" & debet73 & "','" & kredit73 & "','" & header73 & "')"

                'Sql74 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                '" VALUES ('" & Nomor_Akun74 & "','" & Nama_Akun74 & "','" & Kelompok_Akun74 & "','" & Tanggal_Input1 & "','" & Tahun74 & "','" & debet74 & "','" & kredit74 & "','" & header74 & "')"

                Sql75 = "INSERT INTO Table_AkunArusKas2(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Tahun, Saldo_Debet, Saldo_Kredit, header)" & _
                " VALUES ('" & Nomor_Akun75 & "','" & Nama_Akun75 & "','" & Kelompok_Akun75 & "','" & Tanggal_Input1 & "','" & Tahun75 & "','" & debet75 & "','" & kredit75 & "','" & header75 & "')"




                Dim command1 As New SqlCommand(Sql1, Module2.Koneksi)
                command1.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command2 As New SqlCommand(Sql2, Module2.Koneksi)
                command2.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command3 As New SqlCommand(Sql3, Module2.Koneksi)
                command3.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command4 As New SqlCommand(Sql4, Module2.Koneksi)
                command4.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command5 As New SqlCommand(Sql5, Module2.Koneksi)
                command5.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command6 As New SqlCommand(Sql6, Module2.Koneksi)
                command6.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command7 As New SqlCommand(Sql7, Module2.Koneksi)
                command7.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command8 As New SqlCommand(Sql8, Module2.Koneksi)
                command8.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command9 As New SqlCommand(Sql9, Module2.Koneksi)
                command9.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command10 As New SqlCommand(Sql10, Module2.Koneksi)
                command10.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command11 As New SqlCommand(Sql11, Module2.Koneksi)
                command11.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command12 As New SqlCommand(Sql12, Module2.Koneksi)
                command12.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command13 As New SqlCommand(Sql13, Module2.Koneksi)
                command13.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command14 As New SqlCommand(Sql14, Module2.Koneksi)
                command14.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command15 As New SqlCommand(Sql15, Module2.Koneksi)
                command15.ExecuteNonQuery()
                Module2.Koneksi.Close()


                Dim command16 As New SqlCommand(Sql16, Module2.Koneksi)
                command16.ExecuteNonQuery()
                Module2.Koneksi.Close()

                'Dim command17 As New SqlCommand(Sql17, Module2.Koneksi)
                'command17.ExecuteNonQuery()
                'Module2.Koneksi.Close()

                Dim command18 As New SqlCommand(Sql18, Module2.Koneksi)
                command18.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command19 As New SqlCommand(Sql19, Module2.Koneksi)
                command19.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command20 As New SqlCommand(Sql20, Module2.Koneksi)
                command20.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command21 As New SqlCommand(Sql21, Module2.Koneksi)
                command21.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command22 As New SqlCommand(Sql22, Module2.Koneksi)
                command22.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command23 As New SqlCommand(Sql23, Module2.Koneksi)
                command23.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command24 As New SqlCommand(Sql24, Module2.Koneksi)
                command24.ExecuteNonQuery()
                Module2.Koneksi.Close()

                'Dim command25 As New SqlCommand(Sql25, Module2.Koneksi)
                'command25.ExecuteNonQuery()
                'Module2.Koneksi.Close()

                'Dim command26 As New SqlCommand(Sql26, Module2.Koneksi)
                'command26.ExecuteNonQuery()
                'Module2.Koneksi.Close()

                'Dim command27 As New SqlCommand(Sql27, Module2.Koneksi)
                'command27.ExecuteNonQuery()
                'Module2.Koneksi.Close()


                Dim command28 As New SqlCommand(Sql28, Module2.Koneksi)
                command28.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command29 As New SqlCommand(Sql29, Module2.Koneksi)
                command29.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command30 As New SqlCommand(Sql30, Module2.Koneksi)
                command30.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command31 As New SqlCommand(Sql31, Module2.Koneksi)
                command31.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command32 As New SqlCommand(Sql32, Module2.Koneksi)
                command32.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command33 As New SqlCommand(Sql33, Module2.Koneksi)
                command33.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command34 As New SqlCommand(Sql34, Module2.Koneksi)
                command34.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command35 As New SqlCommand(Sql35, Module2.Koneksi)
                command35.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command36 As New SqlCommand(Sql36, Module2.Koneksi)
                command36.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command37 As New SqlCommand(Sql37, Module2.Koneksi)
                command37.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command38 As New SqlCommand(Sql38, Module2.Koneksi)
                command38.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command39 As New SqlCommand(Sql39, Module2.Koneksi)
                command39.ExecuteNonQuery()
                Module2.Koneksi.Close()


                Dim command40 As New SqlCommand(Sql40, Module2.Koneksi)
                command40.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command41 As New SqlCommand(Sql41, Module2.Koneksi)
                command41.ExecuteNonQuery()
                Module2.Koneksi.Close()

                'Dim command42 As New SqlCommand(Sql42, Module2.Koneksi)
                'command42.ExecuteNonQuery()
                'Module2.Koneksi.Close()

                Dim command43 As New SqlCommand(Sql43, Module2.Koneksi)
                command43.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command44 As New SqlCommand(Sql44, Module2.Koneksi)
                command44.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command45 As New SqlCommand(Sql45, Module2.Koneksi)
                command45.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command46 As New SqlCommand(Sql46, Module2.Koneksi)
                command46.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command47 As New SqlCommand(Sql47, Module2.Koneksi)
                command47.ExecuteNonQuery()
                Module2.Koneksi.Close()


                Dim command48 As New SqlCommand(Sql48, Module2.Koneksi)
                command48.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command49 As New SqlCommand(Sql49, Module2.Koneksi)
                command49.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command50 As New SqlCommand(Sql50, Module2.Koneksi)
                command50.ExecuteNonQuery()
                Module2.Koneksi.Close()

                'Dim command51 As New SqlCommand(Sql51, Module2.Koneksi)
                'command51.ExecuteNonQuery()
                'Module2.Koneksi.Close()

                'Dim command52 As New SqlCommand(Sql52, Module2.Koneksi)
                'command52.ExecuteNonQuery()
                'Module2.Koneksi.Close()


                'Dim command53 As New SqlCommand(Sql53, Module2.Koneksi)
                'command53.ExecuteNonQuery()
                'Module2.Koneksi.Close()

                'Dim command54 As New SqlCommand(Sql54, Module2.Koneksi)
                'command54.ExecuteNonQuery()
                'Module2.Koneksi.Close()

                'Dim command55 As New SqlCommand(Sql55, Module2.Koneksi)
                'command55.ExecuteNonQuery()
                'Module2.Koneksi.Close()


                'Dim command56 As New SqlCommand(Sql56, Module2.Koneksi)
                'command56.ExecuteNonQuery()
                'Module2.Koneksi.Close()

                'Dim command57 As New SqlCommand(Sql57, Module2.Koneksi)
                'command57.ExecuteNonQuery()
                'Module2.Koneksi.Close()

                'Dim command58 As New SqlCommand(Sql58, Module2.Koneksi)
                'command58.ExecuteNonQuery()
                'Module2.Koneksi.Close()

                'Dim command59 As New SqlCommand(Sql59, Module2.Koneksi)
                'command59.ExecuteNonQuery()
                'Module2.Koneksi.Close()

                'Dim command60 As New SqlCommand(Sql60, Module2.Koneksi)
                'command60.ExecuteNonQuery()
                'Module2.Koneksi.Close()

                'Dim command61 As New SqlCommand(Sql61, Module2.Koneksi)
                'command61.ExecuteNonQuery()
                'Module2.Koneksi.Close()

                'Dim command62 As New SqlCommand(Sql62, Module2.Koneksi)
                'command62.ExecuteNonQuery()
                'Module2.Koneksi.Close()

                'Dim command63 As New SqlCommand(Sql63, Module2.Koneksi)
                'command63.ExecuteNonQuery()
                'Module2.Koneksi.Close()

                'Dim command64 As New SqlCommand(Sql64, Module2.Koneksi)
                'command64.ExecuteNonQuery()
                'Module2.Koneksi.Close()

                Dim command65 As New SqlCommand(Sql65, Module2.Koneksi)
                command65.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command66 As New SqlCommand(Sql66, Module2.Koneksi)
                command66.ExecuteNonQuery()
                Module2.Koneksi.Close()








                Dim command67 As New SqlCommand(Sql67, Module2.Koneksi)
                command67.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command68 As New SqlCommand(Sql68, Module2.Koneksi)
                command68.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command69 As New SqlCommand(Sql69, Module2.Koneksi)
                command69.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command70 As New SqlCommand(Sql70, Module2.Koneksi)
                command70.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command71 As New SqlCommand(Sql71, Module2.Koneksi)
                command71.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command72 As New SqlCommand(Sql72, Module2.Koneksi)
                command72.ExecuteNonQuery()
                Module2.Koneksi.Close()

                Dim command73 As New SqlCommand(Sql73, Module2.Koneksi)
                command73.ExecuteNonQuery()
                Module2.Koneksi.Close()

                'Dim command74 As New SqlCommand(Sql74, Module2.Koneksi)
                'command74.ExecuteNonQuery()
                'Module2.Koneksi.Close()

                'Dim command75 As New SqlCommand(Sql75, Module2.Koneksi)
                'command75.ExecuteNonQuery()
                'Module2.Koneksi.Close()





                Dim cryRpt As New ReportDocument
                Dim crTableLogonInfos As New TableLogOnInfos
                Dim crTableLogonInfo As New TableLogOnInfo
                Dim crConnectionInfo As New ConnectionInfo
                Dim crTables As Tables
                Dim crTable As Table


                Dim rep2 As New CrystalReport4 'crJual merupakan nama file CrystalReport

                Dim crPFDs As ParameterFieldDefinitions
                Dim crPFD As ParameterFieldDefinition
                Dim crPVs As New ParameterValues
                Dim crPDV As New ParameterDiscreteValue

                'Dim crPFDs2 As ParameterFieldDefinitions
                'Dim crPFD2 As ParameterFieldDefinition
                'Dim crPDV2 As New ParameterDiscreteValue
                'Dim crPVs2 As New ParameterValues

                'Dim crPFDs3 As ParameterFieldDefinitions
                'Dim crPFD3 As ParameterFieldDefinition
                'Dim crPDV3 As New ParameterDiscreteValue
                'Dim crPVs3 As New ParameterValues

                rep2.Load(".\CrystalReport4.rpt")
                With crConnectionInfo
                    .ServerName = "(local)"
                    .DatabaseName = "Tugas_Akhir"
                    .UserID = "sa"
                    .Password = "otadmin"
                End With
                crTables = rep2.Database.Tables
                For Each crTable In crTables
                    crTableLogonInfo = crTable.LogOnInfo
                    crTableLogonInfo.ConnectionInfo = crConnectionInfo
                    crTable.ApplyLogOnInfo(crTableLogonInfo)
                Next








                'Dim umur As String
                'umur = "12"
                ''LapAktivitas.CrystalReportViewer1.SelectionFormula = "{Nomor_Akun.Tgl_terima} in date ('" & CDate(Format(DateTimePicker1.Value, "yyyy-MM-dd")) & "') TO DateValue ('" & CDate(Format(DateTimePicker2.Value, "yyyy-MM-dd")) & "')"
                ''LapAktivitas.CrystalReportViewer1.SelectionFormula = "{Table_Akun.Tanggal_Input} in date ('" & CDate(Format(DateTimePicker1.Value, "yyyy-MM-dd")) & "') TO DateValue ('" & CDate(Format(DateTimePicker2.Value, "yyyy-MM-dd")) & "')"
                ''LapAktivitas.CrystalReportViewer1.SelectionFormula = "{Table_Akun.Kode_Neraca} = '1'"
                'LapAktivitas.CrystalReportViewer1.SelectionFormula = "{Table_Akun.Bulan}='" & bln & "' AND {Table_Akun.Tahun}='" & txtTahun.Text & "'"
                ''Form2.CrystalReportViewer1.SelectionFormula = "{Disposisi.Tgl_terima} in date ('" & CDate(Format(DateTimePicker1.Value, "yyyy-MM-dd")) & "') TO DateValue ('" & CDate(Format(DateTimePicker2.Value, "yyyy-MM-dd")) & "')"
                ''"{Table_Akun.Bulan} = #" & Cstr(txtTahun.text)) & "# and {Table_Akun.Tahun} = #" & cStr(txtBulan.text) & "#"
                ''LapAktivitas.CrystalReportViewer1.ReportSource = cryRpt
                LapArusKas.CrystalReportViewer1.SelectionFormula = "{Table_AkunArusKas2.Tahun}='" & Trim(txtTahun.Text) & "' AND {Table_AkunArusKas2.Nomor_Akun}='1101' Or {Table_AkunArusKas2.Nomor_Akun}='1102' Or {Table_AkunArusKas2.Nomor_Akun}='1103' Or {Table_AkunArusKas2.Nomor_Akun}='111111' Or {Table_AkunArusKas2.Nomor_Akun}='222222' Or {Table_AkunArusKas2.Nomor_Akun}='333333'"

                Dim Bulan As String
                GetTanggalAkhirBulan()
                'Bulan = Format(DateTimePicker1.Value, "MM")
                'Tahun = Format(DateTimePicker1.Value, "yyyy")
                Bulan = bln_tahun
                crPDV.Value = Bulan
                'crPDV2.Value = Tahun
                'crPDV3.Value = "df"

                crPFDs = rep2.DataDefinition.ParameterFields
                'crPFDs2 = rep.DataDefinition.ParameterFields
                'crPFDs3 = rep.DataDefinition.ParameterFields

                crPFD = crPFDs.Item("Bulan")
                'crPFD2 = crPFDs2.Item("Tahun")
                'crPFD3 = crPFDs3.Item("KontribusiSantri")

                crPVs.Clear()
                crPVs.Add(crPDV)

                'crPVs2.Clear()
                'crPVs2.Add(crPDV2)

                'crPVs2.Clear()
                'crPVs2.Add(crPDV3)

                crPFD.ApplyCurrentValues(crPVs)
                'crPFD2.ApplyCurrentValues(crPVs2)
                'crPFD3.ApplyCurrentValues(crPVs3)



                Dim paramFields As New CrystalDecisions.Shared.ParameterFields()
                Dim paramField As New CrystalDecisions.Shared.ParameterField()
                Dim discreteVal As New CrystalDecisions.Shared.ParameterDiscreteValue()

                Dim paramField2 As New CrystalDecisions.Shared.ParameterField()
                Dim discreteVal2 As New CrystalDecisions.Shared.ParameterDiscreteValue()

                Dim paramField3 As New CrystalDecisions.Shared.ParameterField()
                Dim discreteVal3 As New CrystalDecisions.Shared.ParameterDiscreteValue()

                Dim paramField4 As New CrystalDecisions.Shared.ParameterField()
                Dim discreteVal4 As New CrystalDecisions.Shared.ParameterDiscreteValue()

                Dim paramField5 As New CrystalDecisions.Shared.ParameterField()
                Dim discreteVal5 As New CrystalDecisions.Shared.ParameterDiscreteValue()

                Dim paramField6 As New CrystalDecisions.Shared.ParameterField()
                Dim discreteVal6 As New CrystalDecisions.Shared.ParameterDiscreteValue()

                Dim paramField7 As New CrystalDecisions.Shared.ParameterField()
                Dim discreteVal7 As New CrystalDecisions.Shared.ParameterDiscreteValue()

                Dim paramField8 As New CrystalDecisions.Shared.ParameterField()
                Dim discreteVal8 As New CrystalDecisions.Shared.ParameterDiscreteValue()

                'Dim paramField9 As New CrystalDecisions.Shared.ParameterField()
                'Dim discreteVal9 As New CrystalDecisions.Shared.ParameterDiscreteValue()

                'Dim paramField10 As New CrystalDecisions.Shared.ParameterField()
                'Dim discreteVal10 As New CrystalDecisions.Shared.ParameterDiscreteValue()

                'Dim paramField11 As New CrystalDecisions.Shared.ParameterField()
                'Dim discreteVal11 As New CrystalDecisions.Shared.ParameterDiscreteValue()

                'Dim paramField12 As New CrystalDecisions.Shared.ParameterField()
                'Dim discreteVal12 As New CrystalDecisions.Shared.ParameterDiscreteValue()

                'Dim paramField13 As New CrystalDecisions.Shared.ParameterField()
                'Dim discreteVal13 As New CrystalDecisions.Shared.ParameterDiscreteValue()

                'Dim paramField14 As New CrystalDecisions.Shared.ParameterField()
                'Dim discreteVal14 As New CrystalDecisions.Shared.ParameterDiscreteValue()

                'Dim paramField15 As New CrystalDecisions.Shared.ParameterField()
                'Dim discreteVal15 As New CrystalDecisions.Shared.ParameterDiscreteValue()

                'Dim paramField16 As New CrystalDecisions.Shared.ParameterField()
                'Dim discreteVal16 As New CrystalDecisions.Shared.ParameterDiscreteValue()

                'Dim paramField17 As New CrystalDecisions.Shared.ParameterField()
                'Dim discreteVal17 As New CrystalDecisions.Shared.ParameterDiscreteValue()

                'Dim paramField18 As New CrystalDecisions.Shared.ParameterField()
                'Dim discreteVal18 As New CrystalDecisions.Shared.ParameterDiscreteValue()

                'Dim paramField19 As New CrystalDecisions.Shared.ParameterField()
                'Dim discreteVal19 As New CrystalDecisions.Shared.ParameterDiscreteValue()

                'Dim paramField20 As New CrystalDecisions.Shared.ParameterField()
                'Dim discreteVal20 As New CrystalDecisions.Shared.ParameterDiscreteValue()

                'Dim paramField21 As New CrystalDecisions.Shared.ParameterField()
                'Dim discreteVal21 As New CrystalDecisions.Shared.ParameterDiscreteValue()

                'Dim paramField22 As New CrystalDecisions.Shared.ParameterField()
                'Dim discreteVal22 As New CrystalDecisions.Shared.ParameterDiscreteValue()

                'Dim paramField23 As New CrystalDecisions.Shared.ParameterField()
                'Dim discreteVal23 As New CrystalDecisions.Shared.ParameterDiscreteValue()

                'Dim paramField24 As New CrystalDecisions.Shared.ParameterField()
                'Dim discreteVal24 As New CrystalDecisions.Shared.ParameterDiscreteValue()

                tampil()

                paramField.ParameterFieldName = "Kas"
                'Dim str As String = txtBulan.Text.ToString
                Dim str As String = FormatCurrency(Kas)
                discreteVal.Value = str
                paramField.CurrentValues.Add(discreteVal)
                paramFields.Add(paramField)

                paramField2.ParameterFieldName = "Piutang"
                Dim str2 As String = FormatCurrency(Piutang.ToString)
                discreteVal2.Value = str2
                paramField2.CurrentValues.Add(discreteVal2)
                paramFields.Add(paramField2)

                paramField3.ParameterFieldName = "Perlengkapan"
                Dim str3 As String = FormatCurrency(Val(Perlengkapan.ToString))
                discreteVal3.Value = str3
                paramField3.CurrentValues.Add(discreteVal3)
                paramFields.Add(paramField3)


                paramField4.ParameterFieldName = "BebanOperasional"
                Dim str4 As String = FormatCurrency(Val(BebanOperasional.ToString))
                discreteVal4.Value = str4
                paramField4.CurrentValues.Add(discreteVal4)
                paramFields.Add(paramField4)

                paramField5.ParameterFieldName = "BebanAdministrasi"
                Dim str5 As String = FormatCurrency(Val(BebanAdministrasi.ToString))
                discreteVal5.Value = str5
                paramField5.CurrentValues.Add(discreteVal5)
                paramFields.Add(paramField5)

                paramField6.ParameterFieldName = "SaldoAwal"
                Dim str6 As String = FormatCurrency(Val(SaldoAwal.ToString))
                discreteVal6.Value = str6
                paramField6.CurrentValues.Add(discreteVal6)
                paramFields.Add(paramField6)

                paramField7.ParameterFieldName = "jumlahneto"
                Dim str7 As String = FormatCurrency(Val(jumlahneto.ToString))
                discreteVal7.Value = str7
                paramField7.CurrentValues.Add(discreteVal7)
                paramFields.Add(paramField7)

                paramField8.ParameterFieldName = "saldoakhir"
                Dim str8 As String = FormatCurrency(Val(saldoakhir.ToString))
                discreteVal8.Value = str8
                paramField8.CurrentValues.Add(discreteVal8)
                paramFields.Add(paramField8)

                'paramField9.ParameterFieldName = "BebanKonsumsiSantri"
                'Dim str9 As String = FormatCurrency(Val(BebanKonsumsiSantri.ToString))
                'discreteVal9.Value = str9
                'paramField9.CurrentValues.Add(discreteVal9)
                'paramFields.Add(paramField9)

                'paramField10.ParameterFieldName = "BebanListrik"
                'Dim str10 As String = FormatCurrency(Val(BebanListrik.ToString))
                'discreteVal10.Value = str10
                'paramField10.CurrentValues.Add(discreteVal10)
                'paramFields.Add(paramField10)

                'paramField11.ParameterFieldName = "BebanPerlengkapan"
                'Dim str11 As String = FormatCurrency(Val(BebanPerlengkapan.ToString))
                'discreteVal11.Value = str11
                'paramField11.CurrentValues.Add(discreteVal11)
                'paramFields.Add(paramField11)

                'paramField12.ParameterFieldName = "BebanKegiatanSantri"
                'Dim str12 As String = FormatCurrency(Val(BebanKegiatanSantri.ToString))
                'discreteVal12.Value = str12
                'paramField12.CurrentValues.Add(discreteVal12)
                'paramFields.Add(paramField12)

                'paramField13.ParameterFieldName = "BebanKebersihan"
                'Dim str13 As String = FormatCurrency(Val(BebanKebersihan.ToString))
                'discreteVal13.Value = str13
                'paramField13.CurrentValues.Add(discreteVal13)
                'paramFields.Add(paramField13)

                'paramField14.ParameterFieldName = "BebanKesehatan"
                'Dim str14 As String = FormatCurrency(Val(BebanKesehatan.ToString))
                'discreteVal14.Value = str14
                'paramField14.CurrentValues.Add(discreteVal14)
                'paramFields.Add(paramField14)

                'paramField15.ParameterFieldName = "BebanIklan"
                'Dim str15 As String = FormatCurrency(Val(BebanIklan.ToString))
                'discreteVal15.Value = str15
                'paramField15.CurrentValues.Add(discreteVal15)
                'paramFields.Add(paramField15)

                'paramField16.ParameterFieldName = "BebanRapat"
                'Dim str16 As String = FormatCurrency(Val(BebanRapat.ToString))
                'discreteVal16.Value = str16
                'paramField16.CurrentValues.Add(discreteVal16)
                'paramFields.Add(paramField16)

                'paramField17.ParameterFieldName = "Jumlah_beban"
                'Dim str17 As String = FormatCurrency(Val(Jumlah_beban.ToString))
                'discreteVal17.Value = str17
                'paramField17.CurrentValues.Add(discreteVal17)
                'paramFields.Add(paramField17)

                'paramField18.ParameterFieldName = "NaikTurun"
                'Dim str18 As String = FormatCurrency(Val(NaikTurun.ToString))
                'discreteVal18.Value = str18
                'paramField18.CurrentValues.Add(discreteVal18)
                'paramFields.Add(paramField18)

                'paramField19.ParameterFieldName = "SaldoAkhir"
                'Dim str19 As String = FormatCurrency(Val(SaldoAkhir.ToString))
                'discreteVal19.Value = str19
                'paramField19.CurrentValues.Add(discreteVal19)
                'paramFields.Add(paramField19)

                'paramField20.ParameterFieldName = "SaldoAwal"
                'Dim str20 As String = FormatCurrency(Val(SaldoAwal.ToString))
                'discreteVal20.Value = str20
                'paramField20.CurrentValues.Add(discreteVal20)
                'paramFields.Add(paramField20)

                'paramField21.ParameterFieldName = "SaldoAkhir2"
                'Dim str21 As String = FormatCurrency(Val(SaldoAkhir2.ToString))
                'discreteVal21.Value = str21
                'paramField21.CurrentValues.Add(discreteVal21)
                'paramFields.Add(paramField21)

                'paramField22.ParameterFieldName = "SaldoAwl"
                'Dim str22 As String = FormatCurrency(Val(SaldoAwl.ToString))
                'discreteVal22.Value = str22
                'paramField22.CurrentValues.Add(discreteVal22)
                'paramFields.Add(paramField22)

                'paramField23.ParameterFieldName = "SaldoAkr"
                'Dim str23 As String = FormatCurrency(Val(SaldoAkr.ToString))
                'discreteVal23.Value = str23
                'paramField23.CurrentValues.Add(discreteVal23)
                'paramFields.Add(paramField23)

                'paramField24.ParameterFieldName = "AsetNetoAkhir"
                'Dim str24 As String = FormatCurrency(Val(AsetNetoAkhir.ToString))
                'discreteVal24.Value = str24
                'paramField24.CurrentValues.Add(discreteVal24)
                'paramFields.Add(paramField24)

                LapArusKas.CrystalReportViewer1.ParameterFieldInfo = paramFields


                LapArusKas.CrystalReportViewer1.ReportSource = rep2
                'LapAktivitas.CrystalReportViewer1.RefreshReport()
                'LapAktivitas.CrystalReportViewer1.Refresh()
                'rep.PrintOptions.PaperOrientation = PaperOrientation.Landscape
                LapArusKas.Show()
                Me.Cursor = Cursors.Default



            End If
        End If
    End Sub

    Private Sub txtTahun_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTahun.SelectedIndexChanged
        GetTanggalAkhirBulan()
    End Sub

    Sub GetTanggalAkhirBulan()
        Dim Tanggal As Date
        Dim TanggalAkhir As Date
        Tanggal = DateSerial(txtTahun.Text, 12, 1) 'membuat tanggal 6 June 2021
        'rumus mencari tanggal akhir dalam suatu bulan
        TanggalAkhir = DateSerial(Year(Tanggal), Month(Tanggal) + 1, 0)
        bln_tahun = (Format(TanggalAkhir, "MMMM yyyy"))
        'hasil = 30 June 2021
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Close()
    End Sub
End Class