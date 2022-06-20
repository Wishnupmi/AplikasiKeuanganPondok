Imports System.Windows.Forms

Public Class MDIParent1

    'Private Sub ShowNewForm(ByVal sender As Object, ByVal e As EventArgs) Handles NewToolStripMenuItem.Click
    '    ' Create a new instance of the child form.
    '    Dim ChildForm As New System.Windows.Forms.Form
    '    ' Make it a child of this MDI form before showing it.
    '    ChildForm.MdiParent = Me

    '    m_ChildFormNumber += 1
    '    ChildForm.Text = "Window " & m_ChildFormNumber

    '    ChildForm.Show()
    'End Sub

    'Private Sub OpenFile(ByVal sender As Object, ByVal e As EventArgs) Handles OpenToolStripMenuItem.Click
    '    Dim OpenFileDialog As New OpenFileDialog
    '    OpenFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
    '    OpenFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    '    If (OpenFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
    '        Dim FileName As String = OpenFileDialog.FileName
    '        ' TODO: Add code here to open the file.
    '    End If
    'End Sub

    'Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles SaveAsToolStripMenuItem.Click
    '    Dim SaveFileDialog As New SaveFileDialog
    '    SaveFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
    '    SaveFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*" 

    '    If (SaveFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
    '        Dim FileName As String = SaveFileDialog.FileName
    '        ' TODO: Add code here to save the current contents of the form to a file.
    '    End If
    'End Sub


    Private Sub ExitToolsStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    'Private Sub CutToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
    '    ' Use My.Computer.Clipboard to insert the selected text or images into the clipboard
    'End Sub

    'Private Sub CopyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
    '    ' Use My.Computer.Clipboard to insert the selected text or images into the clipboard
    'End Sub

    'Private Sub PasteToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
    '    'Use My.Computer.Clipboard.GetText() or My.Computer.Clipboard.GetData to retrieve information from the clipboard.
    'End Sub

    'Private Sub ToolBarToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
    '    'Me.ToolStrip.Visible = Me.ToolBarToolStripMenuItem.Checked
    'End Sub

    'Private Sub StatusBarToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
    '    Me.StatusStrip.Visible = Me.StatusBarToolStripMenuItem.Checked
    'End Sub

    'Private Sub CascadeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
    '    Me.LayoutMdi(MdiLayout.Cascade)
    'End Sub

    'Private Sub TileVerticalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
    '    Me.LayoutMdi(MdiLayout.TileVertical)
    'End Sub

    'Private Sub TileHorizontalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
    '    Me.LayoutMdi(MdiLayout.TileHorizontal)
    'End Sub

    'Private Sub ArrangeIconsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
    '    Me.LayoutMdi(MdiLayout.ArrangeIcons)
    'End Sub

    'Private Sub CloseAllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
    '    ' Close all child forms of the parent.
    '    For Each ChildForm As Form In Me.MdiChildren
    '        ChildForm.Close()
    '    Next
    'End Sub

    Private m_ChildFormNumber As Integer

    Private Sub DataSantri_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataSantri.Click
        'frmInputSantri.Show()
        Dim newMDIChild As New frmInputSantri()
        ' Set the Parent Form of the Child window.
        newMDIChild.MdiParent = Me
        ' Display the new form.
        newMDIChild.Show()

    End Sub

    Private Sub DataPerkiraanAkun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataPerkiraanAkun.Click
        'frmInputAkun.Show()
        'frmInputSantri.Show()
        Dim newMDIChild As New frmInputAkun()
        ' Set the Parent Form of the Child window.
        newMDIChild.MdiParent = Me
        ' Display the new form.
        newMDIChild.Show()
    End Sub

    Private Sub TransaksiKasMasuk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TransaksiKasMasuk.Click
        frmInputTransaksi.Show()
        'frmInputSantri.Show()
        'Dim newMDIChild As New frmInputTransaksi()
        '' Set the Parent Form of the Child window.
        'newMDIChild.MdiParent = Me
        '' Display the new form.
        'newMDIChild.Show()

    End Sub

    Private Sub LaporanKasMasuk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LaporanKasMasuk.Click
        'LapPosisiKeuangan.Show()
        'frmInputSantri.Show()
        Dim newMDIChild As New frmKasMasuk()
        ' Set the Parent Form of the Child window.
        newMDIChild.MdiParent = Me
        ' Display the new form.
        newMDIChild.Show()

    End Sub

    Private Sub LaporanKasKeluar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LaporanKasKeluar.Click
        'LapPosisiKeuangan.Show()
        'frmInputSantri.Show()
        Dim newMDIChild As New frmKasKeluar()
        ' Set the Parent Form of the Child window.
        newMDIChild.MdiParent = Me
        ' Display the new form.
        newMDIChild.Show()

    End Sub

    Private Sub LaporanPosisiKeuangan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LaporanPosisiKeuangan.Click
        'LapPosisiKeuangan.Show()
        'frmInputSantri.Show()
        Dim newMDIChild As New frmCetakNeraca()
        ' Set the Parent Form of the Child window.
        newMDIChild.MdiParent = Me
        ' Display the new form.
        newMDIChild.Show()
    End Sub

    Private Sub LaporanAktivitas_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LaporanAktivitas.Click
        'LapPosisiKeuangan.Show()
        'frmInputSantri.Show()
        Dim newMDIChild As New frmAktivitas()
        ' Set the Parent Form of the Child window.
        newMDIChild.MdiParent = Me
        ' Display the new form.
        newMDIChild.Show()

    End Sub

    Private Sub DataUser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataUser.Click
        'frmInputSantri.Show()
        Dim newMDIChild As New frmDataUser()
        ' Set the Parent Form of the Child window.
        newMDIChild.MdiParent = Me
        ' Display the new form.
        newMDIChild.Show()
    End Sub

    Private Sub ProsesAkhirBulan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProsesAkhirBulan.Click
        frmAkhirBulan.Show()
        'frmInputSantri.Show()
        'Dim newMDIChild As New frmInputTransaksi()
        '' Set the Parent Form of the Child window.
        'newMDIChild.MdiParent = Me
        '' Display the new form.
        'newMDIChild.Show()
    End Sub

    Private Sub LaporanArusKas_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LaporanArusKas.Click
        'LapPosisiKeuangan.Show()
        'frmInputSantri.Show()
        Dim newMDIChild As New frmArusKas()
        ' Set the Parent Form of the Child window.
        newMDIChild.MdiParent = Me
        ' Display the new form.
        newMDIChild.Show()
    End Sub
End Class
