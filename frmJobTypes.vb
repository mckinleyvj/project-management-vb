Imports System.IO
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.Sql
Imports System.Xml
Imports System.Data.SqlClient
'Imports System.Windows.Forms.PrintPreviewDialog
Imports System.Xml.Xsl

Public Class frmJobTypes

    Dim sqlConn As String = frmConfig.theConnection
    Dim connect As OdbcConnection

    Dim headerFont As New Font("Segoe UI", 9, FontStyle.Bold)
    Dim detailFont As New Font("Segoe UI", 8, FontStyle.Regular)

    Private Sub frmJobTypes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim command1 As OdbcCommand
        Dim myAdapter As OdbcDataAdapter
        Dim myDataSet As DataSet

        connect = New OdbcConnection(sqlConn)

        Dim show1 As String = "SELECT job_type, job_name, adjust_type FROM m_jobs_type where deleted = 0 ORDER BY job_type;"

        Try

            connect.Open()
            command1 = New OdbcCommand(show1, connect)

            myDataSet = New DataSet()
            myDataSet.Tables.Clear()
            myAdapter = New OdbcDataAdapter()
            myAdapter.SelectCommand = command1
            myAdapter.Fill(myDataSet, "job")

            Dim dtRetrievedData As DataTable = myDataSet.Tables("job")

            Dim dtData As New DataTable
            Dim dtDataRows As DataRow

            Dim chkboxcol As New DataGridViewCheckBoxColumn

            dtData.TableName = "JobTable"
            'dtData.Columns.Add("X", GetType(Boolean))
            dtData.Columns.Add("Code")
            dtData.Columns.Add("Description")
            dtData.Columns.Add("Type")
            'dtData.Columns.Add("Phone 2")
            'dtData.Columns.Add("Mobile")
            'dtData.Columns.Add("Deleted")

            For Each dtDataRows In dtRetrievedData.Rows

                Dim cusCode = dtDataRows("job_type").ToString().Trim()
                Dim cusName As String = dtDataRows("job_name").ToString.Trim()
                Dim adj_type As String = dtDataRows("adjust_type").ToString.Trim()
                'Dim cusTel2 As String = dtDataRows("db_tel2").ToString.Trim()
                'Dim cusHP As String = dtDataRows("db_hp").ToString.Trim()
                'Dim cusDel As Boolean = dtDataRows("deleted").ToString.Trim()

                dtData.Rows.Add(New Object() {cusCode.ToString.Trim(), cusName.ToString.Trim(), adj_type.ToString.Trim()})
            Next

            Dim finalDataSet As New DataSet
            finalDataSet.Tables.Add(dtData)

            GridView1.DataSource = finalDataSet.Tables(0)
            'GridView1.ColumnHeadersVisible = False
            'GridView1.Columns.Item("X").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            'GridView1.Columns.Item("X").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            'GridView1.Columns.Item("X").Resizable = DataGridViewTriState.False
            'GridView1.Columns.Item("X").Width = 21

            GridView1.Columns.Item("Code").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Code").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("Code").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("Code").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Code").Width = 100

            GridView1.Columns.Item("Description").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Description").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("Description").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("Description").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Description").FillWeight = 400
            GridView1.Columns.Item("Description").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

            GridView1.Columns.Item("Type").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Type").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("Type").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("Type").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Type").Width = 50

            GridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            GridView1.Refresh()

            connect.Close()
        Catch ex As Exception
            connect.Close()
        End Try

    End Sub

    Private Sub AddToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddToolStripMenuItem.Click
        diaJobType.Close()
        diaJobType.cmbSymbol.SelectedItem = "+"
        diaJobType.StartPosition = FormStartPosition.CenterParent
        diaJobType.ShowDialog()
    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        Dim iRowIndex As Integer
        Dim itm1 As String
        Dim itm2 As String
        Dim itm3 As String

        Try

            iRowIndex = GridView1.CurrentRow.Index

            diaJobType.Close()
            itm1 = GridView1.Item(0, iRowIndex).Value
            itm2 = GridView1.Item(1, iRowIndex).Value
            itm3 = GridView1.Item(2, iRowIndex).Value

            diaJobType.txt1.Text = itm1.ToString
            diaJobType.txt2.Text = itm2.ToString
            diaJobType.cmbSymbol.SelectedItem = itm3.ToString

            'String.Format("{0:n2}", aNumber)

            diaJobType.ValidateExisting.Text = itm1.ToString.ToUpper.Trim
            diaJobType.txt1.Enabled = False
            diaJobType.txt2.Focus()
            diaJobType.txt2.SelectAll()
            diaJobType.StartPosition = FormStartPosition.CenterScreen
            diaJobType.Show()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Public Sub RefreshListToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshListToolStripMenuItem.Click
        GridView1.DataSource = ""
        GridView1.Refresh()

        frmJobTypes_Load(Nothing, Nothing)
    End Sub
End Class