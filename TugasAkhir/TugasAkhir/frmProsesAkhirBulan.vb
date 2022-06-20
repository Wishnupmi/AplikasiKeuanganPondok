Imports System.Data.SqlClient
Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class frmAkhirBulan
    Dim koneksi = Module2.Koneksi
    Dim con As New OleDbConnection
    'Dim ds As New DataSet
    Dim da As OleDbDataAdapter
    Dim dispoBaris As Integer
    Public nmr_urut As String
    Dim bln_romawi As String
    Dim Sql As String = String.Empty
    Dim Sql2 As String = String.Empty
    Dim bln, thn_new As String

    Private Sub frmAkhirBulan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim a As Integer
        For a = 1 To 12
            Dim b As String
            b = MonthName(a)
            txtBulan.Items.Add(b)
        Next


        Dim awal, akhir, tahun, i As Integer

        awal = 2020
        akhir = 2030
        tahun = 0



        For i = awal + 1 To akhir

            txtTahun.Items.Add(i)
            'i = i + 1
        Next




        DataGridView1.RowsDefaultCellStyle.BackColor = Color.AliceBlue
        DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.LightSteelBlue

        With DataGridView1
            .Columns.Add("No", "NO.")
            .Columns.Add("Nomor_Akun", "Nomor_Akun")
            .Columns.Add("Nama_Akun", "Nama_Akun")
            .Columns.Add("Kelompok_Akun", "Kelompok_Akun")
            .Columns.Add("Tanggal_Input", "Tanggal_Input")
            .Columns.Add("Bulan", "Bulan")
            .Columns.Add("Tahun", "Tahun")
            .Columns.Add("Saldo_Debet", "Saldo_Debet")
            .Columns.Add("Saldo_Kredit", "Saldo_Kredit")
            .Columns.Add("Id_Akun", "Id_Akun")
            .Columns.Add("Header", "Header")
            '.Columns.Add("Kredit", "Kredit")
        End With

        DataGridView1.Columns.Item("No").Width = 50
        DataGridView1.Columns.Item("Nomor_Akun").Width = 80
        DataGridView1.Columns.Item("Nama_Akun").Width = 150
        DataGridView1.Columns.Item("Kelompok_Akun").Width = 150
        DataGridView1.Columns.Item("Tanggal_Input").Width = 150
        DataGridView1.Columns.Item("Bulan").Width = 50
        DataGridView1.Columns.Item("Tahun").Width = 50
        DataGridView1.Columns.Item("Saldo_Debet").Width = 100
        DataGridView1.Columns.Item("Saldo_Kredit").Width = 100
        DataGridView1.Columns.Item("Id_Akun").Width = 40
        DataGridView1.Columns.Item("Header").Width = 150
        'DataGridView1.Columns.Item("Kredit").Width = 150

        DataGridView1.Columns("No").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Nomor_Akun").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Nama_Akun").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Kelompok_Akun").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Tanggal_Input").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Bulan").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Tahun").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Saldo_Debet").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Saldo_Kredit").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Id_Akun").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Header").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        'DataGridView1.Columns("Kredit").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter


        DataGridView1.Columns("No").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Columns("Nomor_Akun").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        DataGridView1.Columns("Nama_Akun").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        DataGridView1.Columns("Kelompok_Akun").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        DataGridView1.Columns("Tanggal_Input").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        DataGridView1.Columns("Bulan").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        DataGridView1.Columns("Tahun").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        DataGridView1.Columns("Saldo_Debet").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        DataGridView1.Columns("Saldo_Kredit").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        DataGridView1.Columns("Id_Akun").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Columns("Header").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        'DataGridView1.Columns("Kredit").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft

        DataGridView1.Columns("Id_Akun").Visible = False
        DataGridView1.Columns("Tanggal_Input").Visible = False





    End Sub

    Private Sub tampil()
        tutupkoneksi()
        bukakoneksi()
        'NomorAgenda()
        'Where Nama_Akun like '" & txtCariAkun.Text & "'

        'TG = Format(.Cells(1).Value, "dd") & "/" & Format(.Cells(1).Value, "MM") & "/" & Format(.Cells(1).Value, "yyyy")
        'Dim b As String
        'txtTgl_Transaksi.Value = Convert.ToDateTime(.Cells(1).Value)

        Sql = "SELECT * From Table_Akun WHERE Bulan='" & CStr(bln) & "' AND Tahun='" & CStr(txtTahun.Text) & "' Order by Nomor_Akun asc"

        Dim da2 As New SqlDataAdapter(sql, conn)
        Dim ds2 As New DataSet
        da2.Fill(ds2, "Table_Akun")
        Dim dt As New DataTable

        Dim i, hitung As Integer
        Dim urutan As String

        DataGridView1.Rows.Clear()
        Dim dsTables As DataTable = ds2.Tables("Table_Akun")

        For i = 0 To ds2.Tables("Table_Akun").Rows.Count - 1
            DataGridView1.Rows.Add()

            With DataGridView1.Rows(i)
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
                .Cells(1).Value = dsTables.Rows(i).Item("Nomor_Akun")
                .Cells(2).Value = dsTables.Rows(i).Item("Nama_Akun")
                .Cells(3).Value = dsTables.Rows(i).Item("Kelompok_Akun")
                .Cells(4).Value = dsTables.Rows(i).Item("Tanggal_Input")
                .Cells(10).Value = dsTables.Rows(i).Item("header")
                Dim bln2 As Integer
                If Val(bln) <= 11 Then
                    bln2 = bln + 1
                Else
                    bln2 = "01"
                    thn_new = Val(txtTahun.Text) + 1
                End If

                If Len(Trim(bln2)) = 2 Then
                    .Cells(5).Value = bln2
                Else
                    .Cells(5).Value = "0" & bln2
                End If

                .Cells(6).Value = thn_new

                If dsTables.Rows(i).Item("Nomor_Akun") = "1101" Then
                    .Cells(7).Value = dsTables.Rows(i).Item("Saldo_Debet")
                    .Cells(8).Value = dsTables.Rows(i).Item("Saldo_Kredit")
                Else
                    .Cells(7).Value = "0"
                    .Cells(8).Value = "0"
                End If
                .Cells(9).Value = dsTables.Rows(i).Item("Id_Akun")

                '.Cells(9).Value = dsTables.Rows(i).Item("Kredit")
                '.Cells(7).Value = dsTables.Rows(i).Item("Id")
                '.Cells(8).Value = dsTables.Rows(i).Item("Nmr_agenda")
                '.Cells(9).Value = dsTables.Rows(i).Item("Sifat_surat")
                '.Cells(10).Value = Format(dsTables.Rows(i).Item("Tgl_diteruskan"), "dd-MM-yyyy")



            End With
        Next


        'For Each dt In ds.Tables
        '    DataGridView1.DataSource = dt
        'Next
        'tutupkoneksi()
    End Sub


    Private Sub btnProses_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProses.Click
        Timer1.Start()
        bln = ""
        If Trim(txtBulan.Text) = "Januari" Then
            bln = "01"
        ElseIf Trim(txtBulan.Text) = "Februari" Then
            bln = "02"
        ElseIf Trim(txtBulan.Text) = "Maret" Then
            bln = "03"
        ElseIf Trim(txtBulan.Text) = "April" Then
            bln = "04"
        ElseIf Trim(txtBulan.Text) = "Mei" Then
            bln = "05"
        ElseIf Trim(txtBulan.Text) = "Juni" Then
            bln = "06"
        ElseIf Trim(txtBulan.Text) = "Juli" Then
            bln = "07"
        ElseIf Trim(txtBulan.Text) = "Agustus" Then
            bln = "08"
        ElseIf Trim(txtBulan.Text) = "September" Then
            bln = "09"
        ElseIf Trim(txtBulan.Text) = "Oktober" Then
            bln = "10"
        ElseIf Trim(txtBulan.Text) = "November" Then
            bln = "11"
        Else
            bln = "12"
        End If
        'MsgBox(bln)
        tampil()

        Dim i As Integer
        Dim pesan2, process2 As Boolean
        pesan2 = False
        process2 = False
        ProgressBar1.Value = 0
        For i = 1 To Me.DataGridView1.RowCount
            If Me.DataGridView1.Rows(i - 1).Cells(0).Value <> "" Then

                tutupkoneksi()
                bukakoneksi()
                cmd = New SqlCommand("Select * from Table_Akun WHERE Nomor_Akun='" & Trim(Me.DataGridView1.Rows(i - 1).Cells(1).Value) & "' AND (Bulan='" & Trim(Me.DataGridView1.Rows(i - 1).Cells(5).Value) & "' AND Tahun='" & Trim(Me.DataGridView1.Rows(i - 1).Cells(6).Value) & "')", conn)

                rd = cmd.ExecuteReader
                rd.Read()
                If Not rd.HasRows Then
                    'MsgBox(Me.DataGridView1.Rows(i - 1).Cells(1).Value & " " & Me.DataGridView1.Rows(i - 1).Cells(5).Value & " " & Me.DataGridView1.Rows(i - 1).Cells(6).Value)
                    'MsgBox(i & " " & Me.DataGridView1.RowCount - 1)

                    Sql = "INSERT INTO Table_Akun(Nomor_Akun, Bulan, Tahun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Saldo_Debet, Saldo_Kredit, header)" & _
                          " VALUES ('" & CStr(Me.DataGridView1.Rows(i - 1).Cells(1).Value) & "','" & CStr(Me.DataGridView1.Rows(i - 1).Cells(5).Value) & "','" & CStr(Me.DataGridView1.Rows(i - 1).Cells(6).Value) & "','" & CStr(Me.DataGridView1.Rows(i - 1).Cells(2).Value) & "','" & CStr(Me.DataGridView1.Rows(i - 1).Cells(3).Value) & "','" & Format(Now, "MM-dd-yyyy") & "','" & CStr(Me.DataGridView1.Rows(i - 1).Cells(7).Value) & "','" & CStr(Me.DataGridView1.Rows(i - 1).Cells(8).Value) & "','" & CStr(Me.DataGridView1.Rows(i - 1).Cells(10).Value) & "')"


                    ' Periksa hati-hati tanda kutip untuk setiap variabel, salah ketik mengakibatkan query anda tidak akan terbaca.
                    ' Persiapan execusi Query Insert

                    Dim command As New SqlCommand(Sql, Module2.Koneksi)
                    command.ExecuteNonQuery()
                    Module2.Koneksi.Close()



                    If ProgressBar1.Value < 100 Then
                        ProgressBar1.Value += 2
                    ElseIf ProgressBar1.Value = 100 Then
                        Timer1.Stop()
                        'process2 = True
                        'GoTo process
                    End If
                Else
                    pesan2 = True
                    GoTo pesan
                End If


            End If
        Next
        '        If process2 = True Then
        'Process:
        MsgBox("Data Berhasil Di Proses, Anda sudah bisa Input Data Transaksi bulan Berikutnya", vbInformation, "Konfirmasi")
        'End If
        If pesan2 = True Then
pesan:
            MsgBox("Tidak Bisa di Proses karena sudah Tutup Buku Akhir Bulan", vbInformation, "Konfirmasi")
        End If

    End Sub

    'Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    '    If ProgressBar1.Value < 100 Then
    '        ProgressBar1.Value += 2
    '    ElseIf ProgressBar1.Value = 100 Then
    '        Timer1.Stop()
    '    End If
    'End Sub
End Class