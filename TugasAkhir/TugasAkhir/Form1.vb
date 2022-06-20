Public Class frmLogin

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        End
    End Sub

    Private Sub frmLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.TopMost = True
    End Sub

    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click
        If UserIdTextBox.Text = "admin" Then
            Me.Hide()
            frmLoading.Show()
        Else
            MsgBox("User Password yang Anda Masukkan Salah. Silakan Masukkan Kembali.", vbInformation, "Konfirmasi Login")
            UserIdTextBox.Focus()
        End If

    End Sub
End Class
