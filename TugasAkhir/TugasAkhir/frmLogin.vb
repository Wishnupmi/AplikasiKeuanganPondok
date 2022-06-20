Imports System.Data
Imports System.Data.SqlClient
Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data.OleDb

Module Modulekoneksi
    Public conn1 As SqlConnection
    Public cmd As SqlCommand
    Public ds As DataSet
    Public adapter As SqlDataAdapter
    Public rd As SqlDataReader
    Public lokasidb As String

    Public Sub ambilkoneksi()
        lokasidb = "Data Source=(local);Initial Catalog=Tugas_Akhir;Persist Security Info=True;User ID=sa; password=otadmin"
        conn = New SqlConnection(lokasidb)
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
                'MsgBox("Koneksi ke database berhasil", MsgBoxStyle.Information, "Informasi")

            End If
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Module

Public Class frmLogin

    Public koneksi As String
    'Public conn As SqlClient.SqlConnection = Nothing
    Public cmd As SqlClient.SqlCommand = Nothing
    Public baca As SqlDataReader
    Public dtadapter As New SqlClient.SqlDataAdapter
    Dim Sql, user, pass As String

    Sub konek()
        lokasidb = "Data Source=.;Initial Catalog=TugasAkhir;Persist Security Info=True;User ID=sa; password=otadmin"
        conn = New SqlConnection(lokasidb)
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
                'MsgBox("Koneksi ke database berhasil", MsgBoxStyle.Information, "Informasi")

            End If
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        End
    End Sub

    Private Sub frmLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.TopMost = True
    End Sub

    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click
        'If UserIdTextBox.Text = "admin" And PasswordTextBox.Text = "admin" Then
        '    Me.Hide()
        '    frmLoading.Show()
        'Else
        '    MsgBox("User Password yang Anda Masukkan Salah. Silakan Masukkan Kembali.", vbInformation, "Konfirmasi Login")
        '    UserIdTextBox.Focus()
        'End If



        ambilkoneksi()

        cmd = New SqlCommand("select * from Table_User where username='" & UserIdTextBox.Text & "' and password='" & PasswordTextBox.Text & "'", conn)
        baca = cmd.ExecuteReader
        baca.Read()
        If baca.HasRows Then
            'Dim x As New MenuUtama
            'x.Show()
            Me.Hide()
            'Form2.Show()
            frmLoading.Show()
            'MDIParent1.Show()
            'MsgBox("login sukses")
        Else
            MsgBox("User Password yang Anda Masukkan Salah. Silakan Masukkan Kembali.", vbInformation, "Konfirmasi Login")
            UserIdTextBox.Focus()
            'txuser.Focus()
        End If

    End Sub
End Class
