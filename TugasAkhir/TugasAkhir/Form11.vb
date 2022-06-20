Imports System.Data.SqlClient
Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class frmCetakNeraca
    Dim koneksi = Module2.Koneksi
    Dim con As New OleDbConnection
    'Dim ds As New DataSet
    Dim da As OleDbDataAdapter
    Dim dispoBaris As Integer
    Public nmr_urut As String
    Dim bln_romawi As String
    Dim Sql As String = String.Empty
    Dim Sql2 As String = String.Empty

    Private Sub Form11_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub cmdPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrint.Click
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






        Dim rep As New CrystalReport1 'crJual merupakan nama file CrystalReport

        Dim crPFDs As ParameterFieldDefinitions
        Dim crPFD As ParameterFieldDefinition
        Dim crPVs As New ParameterValues
        Dim crPDV As New ParameterDiscreteValue


        rep.Load("E:\TugasAkhir\TugasAkhir\CrystalReport1.rpt")
        With crConnectionInfo
            .ServerName = "(local)"
            .DatabaseName = "Tugas_Akhir"
            .UserID = "sa"
            .Password = "DATAPMI2012"
        End With
        crTables = rep.Database.Tables
        For Each crTable In crTables
            crTableLogonInfo = crTable.LogOnInfo
            crTableLogonInfo.ConnectionInfo = crConnectionInfo
            crTable.ApplyLogOnInfo(crTableLogonInfo)
        Next










        'LapAktivitas.CrystalReportViewer1.SelectionFormula = "{Nomor_Akun.Tgl_terima} in date ('" & CDate(Format(DateTimePicker1.Value, "yyyy-MM-dd")) & "') TO DateValue ('" & CDate(Format(DateTimePicker2.Value, "yyyy-MM-dd")) & "')"
        LapAktivitas.CrystalReportViewer1.SelectionFormula = "{Tabel_Akun.Tanggal_Input} in date ('" & CDate(Format(DateTimePicker1.Value, "yyyy-MM-dd")) & "') TO DateValue ('" & CDate(Format(DateTimePicker2.Value, "yyyy-MM-dd")) & "')"

        'Form2.CrystalReportViewer1.SelectionFormula = "{Disposisi.Tgl_terima} in date ('" & CDate(Format(DateTimePicker1.Value, "yyyy-MM-dd")) & "') TO DateValue ('" & CDate(Format(DateTimePicker2.Value, "yyyy-MM-dd")) & "')"

        'LapAktivitas.CrystalReportViewer1.ReportSource = cryRpt
        Dim PeriodeNeraca As Date
        PeriodeNeraca = Format(DateTimePicker1.Value, "dd-MM-yyyy")
        'crPDV.Value = txtNominal.Text
        crPDV.Value = PeriodeNeraca

        crPFDs = rep.DataDefinition.ParameterFields
        crPFD = crPFDs.Item("Nomor_Akun")
        crPVs.Clear()
        crPVs.Add(crPDV)
        crPFD.ApplyCurrentValues(crPVs)


        LapAktivitas.CrystalReportViewer1.ReportSource = rep
        'LapAktivitas.CrystalReportViewer1.RefreshReport()
        'LapAktivitas.CrystalReportViewer1.Refresh()
        LapAktivitas.Show()
        Me.Cursor = Cursors.Default
    End Sub
End Class