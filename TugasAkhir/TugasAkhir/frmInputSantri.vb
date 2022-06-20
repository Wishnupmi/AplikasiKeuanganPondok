Imports System.Data.SqlClient
Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class frmInputSantri
    Dim koneksi = Module2.Koneksi
    Dim con As New OleDbConnection
    'Dim ds As New DataSet
    Dim da As OleDbDataAdapter
    Dim dispoBaris As Integer
    Public nmr_urut As String
    Dim bln_romawi As String
    Dim Sql As String = String.Empty
    Dim Sql2 As String = String.Empty

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Sql2 = "UPDATE Data_Santri Set nama_santri='" & CStr(txtNmSantri.Text) & "',nomor_induk='" & CStr(txtNmrInduk.Text) & "'" & _
        ",jenis_kelamin='" & CStr(cmbJnsKelamin.Text) & "',alamat='" & CStr(txtAlamat.Text) & "',nama_wali_santri='" & CStr(txtNmWaliSantri.Text) & "'" & _
        ",tahun_masuk='" & CStr(thn_masuk.Text) & "',semester='" & CStr(txtSemester.Text) & "' WHERE Id = '" & CInt(txtId.Text) & "'"

        Dim command As New SqlCommand(Sql2, Module2.Koneksi)
        command.ExecuteNonQuery()
        Module2.Koneksi.Close()

        MsgBox("Data sukses terupdate", vbInformation, "Konfirmasi")

        Call tampil()
        Call kosong()

        btnSave.Visible = True
        btnUpdate.Visible = False


        'Sql = "INSERT INTO Table_User(username)" & _
        '      " VALUES ('" & txtNmSantri.Text & "')"

        '' Periksa hati-hati tanda kutip untuk setiap variabel, salah ketik mengakibatkan query anda tidak akan terbaca.
        '' Persiapan execusi Query Insert

        'Dim command As New SqlCommand(Sql, Module2.Koneksi)
        'command.ExecuteNonQuery()
        'Module2.Koneksi.Close()

        'MessageBox.Show("suskses tersimpan")

    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim Sql As String = String.Empty

        Sql = "Delete from Data_Santri where Id = '" & txtId.Text & "'"

        Dim command As New SqlCommand(Sql, Module2.Koneksi)

        command.ExecuteNonQuery()

        Module2.Koneksi.Close()

        MessageBox.Show("sukses terhapus")
        kosong()
        tampil()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        'Sql2 = "UPDATE Table_User Set password='" & CStr(txtNmSantri.Text) & "' WHERE Id = '" & CInt(txtNmrSurat.Text) & "'"

        '' Periksa hati-hati tanda kutip untuk setiap variabel, salah ketik mengakibatkan query anda tidak akan terbaca.
        '' Persiapan execusi Query Insert

        'Dim command As New SqlCommand(Sql2, Module2.Koneksi)
        'command.ExecuteNonQuery()
        'Module2.Koneksi.Close()

        kosong()
        tampil()
        'MessageBox.Show("Berhasil terupdate", "Konfirmasi", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub tampil()
        tutupkoneksi()
        bukakoneksi()
        'NomorAgenda()

        Sql = "SELECT nama_santri, nomor_induk, jenis_kelamin, alamat, nama_wali_santri, tahun_masuk, semester, Id From Data_Santri Order by Id asc"

        Dim da2 As New SqlDataAdapter(Sql, conn)
        Dim ds2 As New DataSet
        da2.Fill(ds2, "DataSantri")
        Dim dt As New DataTable

        Dim i, hitung As Integer
        Dim urutan As String

        DataGridView1.Rows.Clear()
        Dim dsTables As DataTable = ds2.Tables("DataSantri")

        For i = 0 To ds2.Tables("DataSantri").Rows.Count - 1
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
                .Cells(1).Value = dsTables.Rows(i).Item("nomor_induk")
                .Cells(2).Value = dsTables.Rows(i).Item("nama_santri")
                .Cells(3).Value = dsTables.Rows(i).Item("jenis_kelamin")
                .Cells(4).Value = dsTables.Rows(i).Item("alamat")
                .Cells(5).Value = dsTables.Rows(i).Item("tahun_masuk")
                .Cells(6).Value = dsTables.Rows(i).Item("nama_wali_santri")
                .Cells(7).Value = dsTables.Rows(i).Item("semester")
                .Cells(8).Value = dsTables.Rows(i).Item("Id")
                
                '.Cells(10).Value = Format(dsTables.Rows(i).Item("Tgl_diteruskan"), "dd-MM-yyyy")

                

            End With
        Next


        'For Each dt In ds.Tables
        '    DataGridView1.DataSource = dt
        'Next
        'tutupkoneksi()
    End Sub

    Private Sub frmInputSantri_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'txtTgl_terima.Value = Format(Now, "dd-MM-yyyy")
        'txtTgl_diteruskan.Value = Format(Now, "dd-MM-yyyy")

        DataGridView1.RowsDefaultCellStyle.BackColor = Color.AliceBlue
        DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.LightSteelBlue

        With DataGridView1
            .Columns.Add("No", "NO.")
            .Columns.Add("NomorInduk", "NOMOR INDUK")
            .Columns.Add("NamaSantri", "NAMA SANTRI")
            .Columns.Add("jenis_kelamin", "JENIS KELAMIN")
            .Columns.Add("Alamat", "ALAMAT")
            .Columns.Add("NmWali", "NAMA WALI SANTRI")
            .Columns.Add("ThnMasuk", "TAHUN MASUK")
            .Columns.Add("Semester", "SEMESTER")
            .Columns.Add("Id", "Id")

        End With
        
        DataGridView1.Columns.Item("No").Width = 50
        DataGridView1.Columns.Item("NomorInduk").Width = 120
        DataGridView1.Columns.Item("NamaSantri").Width = 150
        DataGridView1.Columns.Item("Alamat").Width = 160
        DataGridView1.Columns.Item("NmWali").Width = 160
        DataGridView1.Columns.Item("ThnMasuk").Width = 120
        DataGridView1.Columns.Item("Semester").Width = 120
        DataGridView1.Columns.Item("Id").Width = 40
        DataGridView1.Columns.Item("jenis_kelamin").Width = 160

        DataGridView1.Columns("No").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("NomorInduk").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("NamaSantri").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Alamat").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("NmWali").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("ThnMasuk").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Semester").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Id").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("jenis_kelamin").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView1.Columns("No").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Columns("NomorInduk").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Columns("NamaSantri").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        DataGridView1.Columns("Alamat").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        DataGridView1.Columns("NmWali").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        DataGridView1.Columns("ThnMasuk").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Columns("Semester").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Columns("Id").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Columns("jenis_kelamin").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft

        DataGridView1.Columns("Id").Visible = False
        'DataGridView1.Columns("jenis_kelamin").Visible = False

        Call tampil()

    End Sub

    'Private Sub frmInputSantri_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
    '    Application.close()


    'End Sub

    
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Close()

    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Sql = "INSERT INTO Data_Santri(nama_santri,nomor_induk,jenis_kelamin,alamat,nama_wali_santri,tahun_masuk,semester)" & _
              " VALUES ('" & txtNmSantri.Text & "','" & txtNmrInduk.Text & "','" & cmbJnsKelamin.Text & "','" & txtAlamat.Text & "','" & txtNmWaliSantri.Text & "','" & thn_masuk.Text & "','" & txtSemester.Text & "')"

        ' Periksa hati-hati tanda kutip untuk setiap variabel, salah ketik mengakibatkan query anda tidak akan terbaca.
        ' Persiapan execusi Query Insert

        Dim command As New SqlCommand(Sql, Module2.Koneksi)
        command.ExecuteNonQuery()
        Module2.Koneksi.Close()

        MsgBox("sukses tersimpan", vbInformation, "Konfirmasi")
        kosong()
        Call tampil()
    End Sub

    Sub kosong()
        txtNmSantri.Text = ""
        txtNmrInduk.Text = ""
        cmbJnsKelamin.Text = ""
        txtAlamat.Text = ""
        txtNmWaliSantri.Text = ""
        thn_masuk.Text = ""
        txtSemester.Text = ""
        txtId.Text = ""
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        Dim i As Integer
        i = Me.DataGridView1.CurrentRow.Index
        With DataGridView1.Rows.Item(i)
            txtNmSantri.Text = .Cells(2).Value
            txtNmrInduk.Text = .Cells(1).Value
            txtAlamat.Text = .Cells(4).Value
            cmbJnsKelamin.Text = .Cells(3).Value
            txtId.Text = .Cells(8).Value
            txtNmWaliSantri.Text = .Cells(5).Value
            thn_masuk.Text = .Cells(6).Value
            txtSemester.Text = .Cells(7).Value
        End With

        btnSave.Visible = False
        btnUpdate.Visible = True
    End Sub
End Class