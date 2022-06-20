Imports System.Data.SqlClient
Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class frmDataUser
    Dim koneksi = Module2.Koneksi
    Dim con As New OleDbConnection
    'Dim ds As New DataSet
    Dim da As OleDbDataAdapter
    Dim dispoBaris As Integer
    Public nmr_urut As String
    Dim bln_romawi As String
    Dim Sql As String = String.Empty
    Dim Sql2 As String = String.Empty

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Sql = "INSERT INTO Table_User(username, password)" & _
              " VALUES ('" & txtUsername.Text & "','" & txtPassword.Text & "')"

        ' Periksa hati-hati tanda kutip untuk setiap variabel, salah ketik mengakibatkan query anda tidak akan terbaca.
        ' Persiapan execusi Query Insert

        Dim command As New SqlCommand(Sql, Module2.Koneksi)
        command.ExecuteNonQuery()
        Module2.Koneksi.Close()

        MsgBox("sukses tersimpan", vbInformation, "konfirmasi")
        kosong()
        Call tampil()
    End Sub

    Sub kosong()
        txtUsername.Text = ""
        txtPassword.Text = ""
    End Sub

    Private Sub frmDataUser_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'txtTgl_terima.Value = Format(Now, "dd-MM-yyyy")
        'txtTgl_diteruskan.Value = Format(Now, "dd-MM-yyyy")

        DataGridView1.RowsDefaultCellStyle.BackColor = Color.AliceBlue
        DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.LightSteelBlue

        With DataGridView1
            .Columns.Add("No", "NO.")
            .Columns.Add("Username", "USERNAME")
            .Columns.Add("Password", "PASSWORD")
            .Columns.Add("Id", "Id")
        End With

        DataGridView1.Columns.Item("No").Width = 50
        DataGridView1.Columns.Item("Username").Width = 160
        DataGridView1.Columns.Item("Password").Width = 150
        DataGridView1.Columns.Item("Id").Width = 40

        DataGridView1.Columns("No").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Username").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Password").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Id").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter


        DataGridView1.Columns("No").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Columns("Username").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        DataGridView1.Columns("Password").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        DataGridView1.Columns("Id").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        DataGridView1.Columns("Id").Visible = False

        Call tampil()
    End Sub

    Private Sub tampil()
        tutupkoneksi()
        bukakoneksi()
        'NomorAgenda()

        Sql = "SELECT username, password, Id From Table_User Order by Id asc"

        Dim da2 As New SqlDataAdapter(Sql, conn)
        Dim ds2 As New DataSet
        da2.Fill(ds2, "DataUser")
        Dim dt As New DataTable

        Dim i, hitung As Integer
        Dim urutan As String

        DataGridView1.Rows.Clear()
        Dim dsTables As DataTable = ds2.Tables("DataUser")

        For i = 0 To ds2.Tables("DataUser").Rows.Count - 1
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
                .Cells(1).Value = dsTables.Rows(i).Item("username")
                .Cells(2).Value = dsTables.Rows(i).Item("password")
                .Cells(3).Value = dsTables.Rows(i).Item("Id")
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

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Close()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        kosong()
        btnSave.Visible = True
        btnUpdate.Visible = False
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        Dim i As Integer
        i = Me.DataGridView1.CurrentRow.Index
        With DataGridView1.Rows.Item(i)
            txtUsername.Text = .Cells(1).Value
            txtPassword.Text = .Cells(2).Value
            txtId.Text = .Cells(3).Value
        End With
        btnSave.Visible = False
        btnUpdate.Visible = True
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim Sql As String = String.Empty

        Sql = "Delete from Table_User where Id = '" & txtId.Text & "'"

        Dim command As New SqlCommand(Sql, Module2.Koneksi)

        command.ExecuteNonQuery()

        Module2.Koneksi.Close()


        MsgBox("Data sukses terhapus", vbInformation, "Konfirmasi")

        Call tampil()
        Call kosong()
        btnSave.Visible = True
        btnUpdate.Visible = False
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Sql2 = "UPDATE Table_User Set username='" & CStr(txtUsername.Text) & "',password='" & CStr(txtPassword.Text) & "' WHERE Id = '" & CInt(txtId.Text) & "'"

        Dim command As New SqlCommand(Sql2, Module2.Koneksi)
        command.ExecuteNonQuery()
        Module2.Koneksi.Close()

        MsgBox("Data sukses terupdate", vbInformation, "Konfirmasi")

        Call tampil()
        Call kosong()
        btnSave.Visible = True
        btnUpdate.Visible = False
    End Sub
End Class