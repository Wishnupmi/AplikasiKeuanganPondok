Imports System.Data.SqlClient
Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class frmCariAkun
    Dim koneksi = Module2.Koneksi
    Dim con As New OleDbConnection
    'Dim ds As New DataSet
    Dim da As OleDbDataAdapter
    Dim dispoBaris As Integer
    Public nmr_urut As String
    Dim bln_romawi As String
    Dim Sql As String = String.Empty
    Dim Sql2 As String = String.Empty
    Dim Kd_Akun As String

    Private Sub frmCariAkun_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DataGridView1.RowsDefaultCellStyle.BackColor = Color.AliceBlue
        DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.LightSteelBlue

        With DataGridView1
            .Columns.Add("No", "NO.")
            .Columns.Add("Nomor_Akun", "Nomor_Akun")
            .Columns.Add("Nama_Akun", "Nama_Akun")
            .Columns.Add("Kelompok_Akun", "Kelompok_Akun")
            .Columns.Add("Tanggal_Input", "Tanggal_Input")
            .Columns.Add("Saldo_Debet", "Saldo_Debet")
            .Columns.Add("Saldo_Kredit", "Saldo_Kredit")
            .Columns.Add("Id_Akun", "Id_Akun")
        End With

        DataGridView1.Columns.Item("No").Width = 50
        DataGridView1.Columns.Item("Nomor_Akun").Width = 130
        DataGridView1.Columns.Item("Nama_Akun").Width = 150
        DataGridView1.Columns.Item("Kelompok_Akun").Width = 150
        DataGridView1.Columns.Item("Tanggal_Input").Width = 150
        DataGridView1.Columns.Item("Saldo_Debet").Width = 150
        DataGridView1.Columns.Item("Saldo_Kredit").Width = 150
        DataGridView1.Columns.Item("Id_Akun").Width = 40

        DataGridView1.Columns("No").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Nomor_Akun").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Nama_Akun").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Kelompok_Akun").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Tanggal_Input").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Saldo_Debet").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Saldo_Kredit").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Id_Akun").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter


        DataGridView1.Columns("No").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Columns("Nomor_Akun").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        DataGridView1.Columns("Nama_Akun").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        DataGridView1.Columns("Kelompok_Akun").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        DataGridView1.Columns("Tanggal_Input").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        DataGridView1.Columns("Saldo_Debet").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        DataGridView1.Columns("Saldo_Kredit").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        DataGridView1.Columns("Id_Akun").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        DataGridView1.Columns("Id_Akun").Visible = False
        DataGridView1.Columns("Tanggal_Input").Visible = False

        tampil()

    End Sub

    Private Sub tampil()
        tutupkoneksi()
        bukakoneksi()
        'NomorAgenda()
        'Where Nama_Akun like '" & txtCariAkun.Text & "'
        Sql = "SELECT * From Table_Akun Order by Id_Akun asc"

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
                .Cells(5).Value = dsTables.Rows(i).Item("Saldo_Debet")
                .Cells(6).Value = dsTables.Rows(i).Item("Saldo_Kredit")
                .Cells(7).Value = dsTables.Rows(i).Item("Id_Akun")
                '.Cells(5).Value = dsTables.Rows(i).Item("Asal")
                '.Cells(6).Value = dsTables.Rows(i).Item("Perihal")
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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'txtCariAkun.Text 
        tampil()
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub



    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        Dim i As Integer


        i = Me.DataGridView1.CurrentRow.Index
        With DataGridView1.Rows.Item(i)
            Kd_Akun = .Cells(1).Value
            'txtPassword.Text = .Cells(2).Value
            'txtId.Text = .Cells(3).Value
        End With
        Close()
        'MsgBox(Kd_Akun)frmInputTransaksi.Show()
        frmInputTransaksi.txtNoAkun.Text = Kd_Akun


    End Sub
End Class