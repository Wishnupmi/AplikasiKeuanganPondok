Imports System.Data.SqlClient
Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class frmKasKeluar
    Dim koneksi = Module2.Koneksi
    Dim con As New OleDbConnection
    'Dim ds As New DataSet
    Dim da As OleDbDataAdapter
    Dim dispoBaris As Integer
    Public nmr_urut As String
    Dim bln_romawi As String
    Dim Sql As String = String.Empty
    Dim Sql2 As String = String.Empty

    Dim i, hitung As Integer
    Dim urutan, bln As String

    Public kontribusi_santri, Pendapatan_Lain As String
    Public Jumlah1 As String

    Private Sub tampil()
        tutupkoneksi()
        bukakoneksi()
        'NomorAgenda()s

        Sql = "SELECT * From Table_Akun Order by Id_Akun asc"

        Dim da2 As New SqlDataAdapter(Sql, conn)
        Dim ds2 As New DataSet
        da2.Fill(ds2, "Table_Akun")
        Dim dt As New DataTable



        'DataGridView1.Rows.Clear()
        Dim dsTables As DataTable = ds2.Tables("Table_Akun")
        'MsgBox(dsTables.Rows(2).Item("Nama_Akun"))

        kontribusi_santri = dsTables.Rows(39).Item("Saldo_Kredit")
        Pendapatan_Lain = dsTables.Rows(45).Item("Saldo_Kredit")
        Jumlah1 = Val(dsTables.Rows(39).Item("Saldo_Kredit")) + Val(dsTables.Rows(45).Item("Saldo_Kredit"))
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

    Private Sub frmKasKeluar_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
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
    End Sub

    Private Sub cmdPrint_Click(sender As System.Object, e As System.EventArgs) Handles cmdPrint.Click
        Dim cryRpt As New ReportDocument
        Dim crTableLogonInfos As New TableLogOnInfos
        Dim crTableLogonInfo As New TableLogOnInfo
        Dim crConnectionInfo As New ConnectionInfo
        Dim crTables As Tables
        Dim crTable As Table

        'cryRpt.Load("E:\TugasAkhir\TugasAkhir\CrystalReport1.rpt")
        'With crConnectionInfo
        '    .ServerName = "(local)"
        '    .DatabaseName = "Tugas_Akhir"
        '    .UserID = "sa"
        '    .Password = "DATAPMI2012"
        'End With
        'crTables = cryRpt.Database.Tables
        'For Each crTable In crTables
        '    crTableLogonInfo = crTable.LogOnInfo
        '    crTableLogonInfo.ConnectionInfo = crConnectionInfo
        '    crTable.ApplyLogOnInfo(crTableLogonInfo)
        'Next
        ''Form2.CrystalReportViewer1.SelectionFormula = "{Disposisi.Tgl_terima} = #" & CDate(Format(DateTimePicker1.Value, "yyyy-MM-dd")) & "#"

        ''LapAktivitas.CrystalReportViewer1.SelectionFormula = "{Disposisi.Tgl_terima} in date ('" & CDate(Format(DateTimePicker1.Value, "yyyy-MM-dd")) & "') TO DateValue ('" & CDate(Format(DateTimePicker2.Value, "yyyy-MM-dd")) & "')"

        ''LaporanPengeluaran.CrystalReportViewer2.SelectionFormula = "{pengeluaran_info.tanggal_pengeluaran} >= #" & CDate(Format(DTTgl1.Value, "yyyy/MM/dd")) & "# and {pengeluaran_info.tanggal_pengeluaran} <= #" & CDate(Format(DTTgl2.Value, "yyyy/MM/dd")) & "#"



        ''Form2.CrystalReportViewer1.SelectionFormula = "{Disposisi.Tgl_terima} = #" & CDate(Format(DateTimePicker1.Value, "yyyy-MM-dd")) & "# "
        'LapAktivitas.CrystalReportViewer1.SelectionFormula = "{Table_Akun.Nomor_Akun} = '" & txtNominal.Text & "'"

        'LapAktivitas.CrystalReportViewer1.ReportSource = cryRpt
        'HasCurrentValue = true;


        'LapAktivitas.CrystalReportViewer1.RefreshReport()
        'LapAktivitas.CrystalReportViewer1.Refresh()

        'LapAktivitas.Show()
        'Me.Cursor = Cursors.Default






        Dim rep As New CrystalReport6 'crJual merupakan nama file CrystalReport

        Dim crPFDs As ParameterFieldDefinitions
        Dim crPFD As ParameterFieldDefinition
        Dim crPVs As New ParameterValues
        Dim crPDV As New ParameterDiscreteValue

        Dim crPFDs2 As ParameterFieldDefinitions
        Dim crPFD2 As ParameterFieldDefinition
        Dim crPDV2 As New ParameterDiscreteValue
        Dim crPVs2 As New ParameterValues

        'Dim crPFDs3 As ParameterFieldDefinitions
        'Dim crPFD3 As ParameterFieldDefinition
        'Dim crPDV3 As New ParameterDiscreteValue
        'Dim crPVs3 As New ParameterValues

        rep.Load(".\CrystalReport6.rpt")
        With crConnectionInfo
            .ServerName = "(local)"
            .DatabaseName = "Tugas_Akhir"
            .UserID = "sa"
            .Password = "otadmin"
        End With
        crTables = rep.Database.Tables
        For Each crTable In crTables
            crTableLogonInfo = crTable.LogOnInfo
            crTableLogonInfo.ConnectionInfo = crConnectionInfo
            crTable.ApplyLogOnInfo(crTableLogonInfo)
        Next



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

        'If Trim(txtBulan.Text) = "January" Then
        '    bln = "01"
        'ElseIf Trim(txtBulan.Text) = "February" Then
        '    bln = "02"
        'ElseIf Trim(txtBulan.Text) = "March" Then
        '    bln = "03"
        'ElseIf Trim(txtBulan.Text) = "April" Then
        '    bln = "04"
        'ElseIf Trim(txtBulan.Text) = "May" Then
        '    bln = "05"
        'ElseIf Trim(txtBulan.Text) = "June" Then
        '    bln = "06"
        'ElseIf Trim(txtBulan.Text) = "July" Then
        '    bln = "07"
        'ElseIf Trim(txtBulan.Text) = "August" Then
        '    bln = "08"
        'ElseIf Trim(txtBulan.Text) = "September" Then
        '    bln = "09"
        'ElseIf Trim(txtBulan.Text) = "October" Then
        '    bln = "10"
        'ElseIf Trim(txtBulan.Text) = "November" Then
        '    bln = "11"
        'Else
        '    bln = "12"
        'End If




        Dim umur As String
        umur = "12"
        'LapAktivitas.CrystalReportViewer1.SelectionFormula = "{Nomor_Akun.Tgl_terima} in date ('" & CDate(Format(DateTimePicker1.Value, "yyyy-MM-dd")) & "') TO DateValue ('" & CDate(Format(DateTimePicker2.Value, "yyyy-MM-dd")) & "')"
        'LapAktivitas.CrystalReportViewer1.SelectionFormula = "{Table_Akun.Tanggal_Input} in date ('" & CDate(Format(DateTimePicker1.Value, "yyyy-MM-dd")) & "') TO DateValue ('" & CDate(Format(DateTimePicker2.Value, "yyyy-MM-dd")) & "')"
        'LapAktivitas.CrystalReportViewer1.SelectionFormula = "{Table_Akun.Kode_Neraca} = '1'"
        LapKasKeluar.CrystalReportViewer1.SelectionFormula = "{Table_Akun.Bulan}='" & bln & "' AND {Table_Akun.Tahun}='" & txtTahun.Text & "' AND {Table_Akun.Saldo_Debet}<>0"
        'Form2.CrystalReportViewer1.SelectionFormula = "{Disposisi.Tgl_terima} in date ('" & CDate(Format(DateTimePicker1.Value, "yyyy-MM-dd")) & "') TO DateValue ('" & CDate(Format(DateTimePicker2.Value, "yyyy-MM-dd")) & "')"
        '"{Table_Akun.Bulan} = #" & Cstr(txtTahun.text)) & "# and {Table_Akun.Tahun} = #" & cStr(txtBulan.text) & "#"
        'LapAktivitas.CrystalReportViewer1.ReportSource = cryRpt

        Dim Bulan, Tahun As String

        Bulan = bln
        Tahun = txtTahun.Text
        'crPDV.Value = txtNominal.Text
        crPDV.Value = Bulan
        crPDV2.Value = Tahun
        'crPDV3.Value = "df"

        crPFDs = rep.DataDefinition.ParameterFields
        crPFDs2 = rep.DataDefinition.ParameterFields
        'crPFDs3 = rep.DataDefinition.ParameterFields

        crPFD = crPFDs.Item("Bulan")
        crPFD2 = crPFDs2.Item("Tahun")
        'crPFD3 = crPFDs3.Item("KontribusiSantri")

        crPVs.Clear()
        crPVs.Add(crPDV)

        crPVs2.Clear()
        crPVs2.Add(crPDV2)

        'crPVs2.Clear()
        'crPVs2.Add(crPDV3)

        crPFD.ApplyCurrentValues(crPVs)
        crPFD2.ApplyCurrentValues(crPVs2)
        'crPFD3.ApplyCurrentValues(crPVs3)



        'Dim paramFields As New CrystalDecisions.Shared.ParameterFields()
        'Dim paramField As New CrystalDecisions.Shared.ParameterField()
        'Dim discreteVal As New CrystalDecisions.Shared.ParameterDiscreteValue()

        'Dim paramField2 As New CrystalDecisions.Shared.ParameterField()
        'Dim discreteVal2 As New CrystalDecisions.Shared.ParameterDiscreteValue()

        'Dim paramField3 As New CrystalDecisions.Shared.ParameterField()
        'Dim discreteVal3 As New CrystalDecisions.Shared.ParameterDiscreteValue()

        'tampil()

        'paramField.ParameterFieldName = "KontribusiSantri"
        ''Dim str As String = txtBulan.Text.ToString
        'Dim str As String = FormatCurrency(kontribusi_santri)
        'discreteVal.Value = str
        'paramField.CurrentValues.Add(discreteVal)
        'paramFields.Add(paramField)

        'paramField2.ParameterFieldName = "PendapatanLain"
        'Dim str2 As String = FormatCurrency(Pendapatan_Lain.ToString)
        'discreteVal2.Value = str2
        'paramField2.CurrentValues.Add(discreteVal2)
        'paramFields.Add(paramField2)

        'paramField3.ParameterFieldName = "Jumlah1"
        'Dim str3 As String = FormatCurrency(Val(Jumlah1.ToString))
        'discreteVal3.Value = str3
        'paramField3.CurrentValues.Add(discreteVal3)
        'paramFields.Add(paramField3)



        'LapAktivitas.CrystalReportViewer1.ParameterFieldInfo = paramFields


        LapKasKeluar.CrystalReportViewer1.ReportSource = rep
        'LapAktivitas.CrystalReportViewer1.RefreshReport()
        'LapAktivitas.CrystalReportViewer1.Refresh()
        'rep.PrintOptions.PaperOrientation = PaperOrientation.Landscape
        LapKasKeluar.Show()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub cmdExit_Click(sender As System.Object, e As System.EventArgs) Handles cmdExit.Click
        Close()
    End Sub
End Class