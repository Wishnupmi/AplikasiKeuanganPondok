Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient

Module Module2
    


    Public comSQL As New SqlClient.SqlCommand
    Public sql As String
    Public conn As New SqlConnection

    Public str As String = "Data Source=(local);Initial Catalog=Tugas_Akhir;Persist Security Info=True;User ID=sa; password=otadmin"


    Public Function Koneksi() As SqlConnection
        Dim conn As New SqlConnection
        lokasidb = "Data Source=(local);Initial Catalog=Tugas_Akhir;Persist Security Info=True;User ID=sa; password=otadmin"
        conn = New SqlConnection(lokasidb)
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
                'MsgBox("Koneksi ke database berhasil", MsgBoxStyle.Information, "Informasi")
                'Return conn
            End If
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Error")
        End Try
        Return conn
    End Function

    Public Sub bukakoneksi()
        If conn.State = ConnectionState.Closed Then
            conn.ConnectionString = str

            Try
                conn.Open()
                'MsgBox("Koneksi Berhasil")
            Catch ex As Exception
                MsgBox("Koneksi Gagal: " & ex.ToString)
            End Try
        End If
    End Sub

    Public Sub tutupkoneksi()
        If conn.State = ConnectionState.Open Then
            Try
                conn.Close()
            Catch ex As Exception
                MsgBox("Gagal menutup koneksi: " & ex.ToString)
            End Try
        End If

    End Sub

End Module