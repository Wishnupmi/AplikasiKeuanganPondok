Imports System.Data.SqlClient
Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class frmInputAkun
    Dim koneksi = Module2.Koneksi
    Dim con As New OleDbConnection
    'Dim ds As New DataSet
    Dim da As OleDbDataAdapter
    Dim dispoBaris As Integer
    Public nmr_urut As String
    Dim bln_romawi As String
    Dim Sql As String = String.Empty
    Dim Sql2 As String = String.Empty

    Private Sub frmInputAkun_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'txtTgl_terima.Value = Format(Now, "dd-MM-yyyy")
        'txtTgl_diteruskan.Value = Format(Now, "dd-MM-yyyy")
        Me.TopMost = True
        txtKelompokAkun.Focus()
        DataGridView2.RowsDefaultCellStyle.BackColor = Color.AliceBlue
        DataGridView2.AlternatingRowsDefaultCellStyle.BackColor = Color.LightSteelBlue

        With DataGridView2
            .Columns.Add("No", "NO.")
            .Columns.Add("NomorAkun", "NOMOR AKUN")
            .Columns.Add("KelompokAkun", "KELOMPOK AKUN")
            .Columns.Add("NamaAkun", "NAMA AKUN")
            .Columns.Add("Id", "ID")
            .Columns.Add("Header", "HEADER")
        End With

        DataGridView2.Columns.Item("No").Width = 50
        DataGridView2.Columns.Item("NomorAkun").Width = 120
        DataGridView2.Columns.Item("KelompokAkun").Width = 150
        DataGridView2.Columns.Item("NamaAkun").Width = 160
        DataGridView2.Columns.Item("Id").Width = 50
        DataGridView2.Columns.Item("Header").Width = 60

        DataGridView2.Columns("No").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns("NomorAkun").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns("KelompokAkun").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns("NamaAkun").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns("Id").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns("Header").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView2.Columns("No").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView2.Columns("NomorAkun").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView2.Columns("KelompokAkun").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        DataGridView2.Columns("NamaAkun").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        DataGridView2.Columns("Id").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView2.Columns("Header").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        DataGridView2.Columns("Id").Visible = False

        Call tampil()

    End Sub

    Private Sub tampil()
        tutupkoneksi()
        bukakoneksi()
        'NomorAgenda()

        Sql = "SELECT * From Table_Akun Order by Nomor_Akun asc"

        Dim da2 As New SqlDataAdapter(Sql, conn)
        Dim ds2 As New DataSet
        da2.Fill(ds2, "TabelAkun")
        Dim dt As New DataTable

        Dim i, hitung As Integer
        Dim urutan As String

        DataGridView2.Rows.Clear()
        Dim dsTables As DataTable = ds2.Tables("TabelAkun")

        For i = 0 To ds2.Tables("TabelAkun").Rows.Count - 1
            DataGridView2.Rows.Add()

            With DataGridView2.Rows(i)
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
                .Cells(2).Value = dsTables.Rows(i).Item("Kelompok_Akun")
                .Cells(3).Value = dsTables.Rows(i).Item("Nama_Akun")
                .Cells(4).Value = dsTables.Rows(i).Item("Id_Akun")
                .Cells(5).Value = dsTables.Rows(i).Item("Header")
                '.Cells(7).Value = dsTables.Rows(i).Item("Id")
                '.Cells(8).Value = dsTables.Rows(i).Item("Nmr_agenda")
                '.Cells(9).Value = dsTables.Rows(i).Item("Sifat_surat")
                '.Cells(10).Value = Format(dsTables.Rows(i).Item("Tgl_diteruskan"), "dd-MM-yyyy")



            End With
        Next


        'For Each dt In ds.Tables
        '    DataGridView2.DataSource = dt
        'Next
        'tutupkoneksi()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Sql = "INSERT INTO Table_Akun(Nomor_Akun, Nama_Akun, Kelompok_Akun, Tanggal_Input, Bulan, Tahun, Saldo_Debet, Saldo_Kredit, Header)" & _
              " VALUES ('" & txtNmrAkun.Text & "','" & txtNmAkun.Text & "','" & txtKelompokAkun.Text & "','" & Format(Now, "MM-dd-yyyy") & "','" & Format(Now, "MM") & "','" & Format(Now, "yyyy") & "','0','0','" & cmbHeader.Text & "')"

        ' Periksa hati-hati tanda kutip untuk setiap variabel, salah ketik mengakibatkan query anda tidak akan terbaca.
        ' Persiapan execusi Query Insert

        Dim command As New SqlCommand(Sql, Module2.Koneksi)
        command.ExecuteNonQuery()
        Module2.Koneksi.Close()

        MsgBox("Data berhasil tersimpan", vbInformation, "Konfirmasi")
        Call tampil()
        Call reset()
    End Sub


   

    Private Sub DataGridView2_CellMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs)

    End Sub

    Private Sub DataGridView2_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView2.DoubleClick
        Dim i As Integer
        i = Me.DataGridView2.CurrentRow.Index
        With DataGridView2.Rows.Item(i)
            txtNmrAkun.Text = .Cells(1).Value
            txtNmAkun.Text = .Cells(3).Value
            txtKelompokAkun.Text = .Cells(2).Value
            txtIdAkun.Text = .Cells(4).Value
            cmbHeader.Text = IIf(IsDBNull(.Cells(5).Value), "", .Cells(5).Value)
        End With
        btnSave.Visible = False
        btnUpdate.Visible = True


    End Sub

    Private Sub DataGridView2_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Close()
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim Sql As String = String.Empty

        Sql = "Delete from Table_Akun where Id_Akun = '" & txtIdAkun.Text & "'"

        Dim command As New SqlCommand(Sql, Module2.Koneksi)

        command.ExecuteNonQuery()

        Module2.Koneksi.Close()


        MsgBox("Data sukses terhapus", vbInformation, "Konfirmasi")

        Call tampil()
        Call reset()
        btnSave.Visible = True
        btnUpdate.Visible = False
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Sql2 = "UPDATE Table_Akun Set Nama_Akun='" & CStr(txtNmAkun.Text) & "',Kelompok_Akun='" & CStr(txtKelompokAkun.Text) & "',Header='" & cmbHeader.Text & "' WHERE Id_Akun = '" & CInt(txtIdAkun.Text) & "'"

        Dim command As New SqlCommand(Sql2, Module2.Koneksi)
        command.ExecuteNonQuery()
        Module2.Koneksi.Close()

        MsgBox("Data sukses terupdate", vbInformation, "Konfirmasi")

        Call tampil()
        Call reset()
        btnSave.Visible = True
        btnUpdate.Visible = False
    End Sub

    Sub reset()
        txtNmrAkun.Text = ""
        txtKelompokAkun.Text = ""
        txtNmAkun.Text = ""
        txtIdAkun.Text = ""
        cmbHeader.Text = ""
    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        reset()
        tampil()
    End Sub
End Class