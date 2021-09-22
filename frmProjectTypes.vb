Imports System.IO
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.Sql
Imports System.Xml
Imports System.Data.SqlClient
'Imports System.Windows.Forms.PrintPreviewDialog
Imports System.Xml.Xsl

Public Class frmProjectTypes

    Dim sqlConn As String = frmConfig.theConnection
    Dim connect As OdbcConnection

    Dim headerFont As New Font("Segoe UI", 9, FontStyle.Bold)
    Dim detailFont As New Font("Segoe UI", 8, FontStyle.Regular)

    Private Sub frmProjectTypes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim command1 As OdbcCommand
        Dim myAdapter As OdbcDataAdapter
        Dim myDataSet As DataSet

        connect = New OdbcConnection(sqlConn)

        Dim show1 As String = "SELECT project_code, project_name, project_price FROM m_project_type WHERE deleted = 0 ORDER BY project_code;"

        Try

            connect.Open()
            command1 = New OdbcCommand(show1, connect)

            myDataSet = New DataSet()
            myDataSet.Tables.Clear()
            myAdapter = New OdbcDataAdapter()
            myAdapter.SelectCommand = command1
            myAdapter.Fill(myDataSet, "proj")

            Dim dtRetrievedData As DataTable = myDataSet.Tables("proj")

            Dim dtData As New DataTable
            Dim dtDataRows As DataRow

            Dim chkboxcol As New DataGridViewCheckBoxColumn

            dtData.TableName = "ProjTable"
            'dtData.Columns.Add("X", GetType(Boolean))
            dtData.Columns.Add("Code")
            dtData.Columns.Add("Description")
            dtData.Columns.Add("Price (B$)")
            'dtData.Columns.Add("Phone 1")
            'dtData.Columns.Add("Phone 2")
            'dtData.Columns.Add("Mobile")
            'dtData.Columns.Add("Deleted")

            For Each dtDataRows In dtRetrievedData.Rows

                Dim cusCode = dtDataRows("project_code").ToString().Trim()
                Dim cusName As String = dtDataRows("project_name").ToString.Trim()
                Dim cusName2 As String = dtDataRows("project_price").ToString.Trim()
                'Dim cusTel1 As String = dtDataRows("db_tel1").ToString.Trim()
                'Dim cusTel2 As String = dtDataRows("db_tel2").ToString.Trim()
                'Dim cusHP As String = dtDataRows("db_hp").ToString.Trim()
                'Dim cusDel As Boolean = dtDataRows("deleted").ToString.Trim()

                dtData.Rows.Add(New Object() {cusCode.ToString.Trim(), cusName.ToString.Trim(), cusName2.ToString.Trim()})
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
            GridView1.Columns.Item("Description").FillWeight = 255
            GridView1.Columns.Item("Description").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            'GridView1.Columns.Item("Description").DefaultCellStyle.WrapMode = DataGridViewTriState.True
            '.DefaultCellStyle.WrapMode = DataGridViewTriState.True

            GridView1.Columns.Item("Price (B$)").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
            GridView1.Columns.Item("Price (B$)").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("Price (B$)").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("Price (B$)").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            GridView1.Columns.Item("Price (B$)").Resizable = DataGridViewTriState.False
            GridView1.Columns.Item("Price (B$)").Width = 85

            'GridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            'GridView1.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True
            'GridView1.Refresh()

            'GridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True
            'Gridview1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            'GridView1.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True
            GridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            connect.Close()
        Catch ex As Exception
            connect.Close()
        End Try

    End Sub

    Private Sub AddToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddToolStripMenuItem.Click
        diaProjType.Close()
        diaProjType.StartPosition = FormStartPosition.CenterParent
        diaProjType.ShowDialog()
    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        Dim iRowIndex As Integer
        Dim itm1 As String
        Dim itm2 As String
        Dim itm7 As String

        Try

            iRowIndex = GridView1.CurrentRow.Index

            diaProjType.Close()
            itm1 = GridView1.Item(0, iRowIndex).Value
            itm2 = GridView1.Item(1, iRowIndex).Value
            itm7 = GridView1.Item(2, iRowIndex).Value

            diaProjType.txt1.Text = itm1.ToString
            diaProjType.txt2.Text = itm2.ToString
            diaProjType.txt7.Text = itm7.ToString

            'String.Format("{0:n2}", aNumber)

            diaProjType.ValidateExisting.Text = itm1.ToString.ToUpper.Trim
            diaProjType.txt1.Enabled = False
            diaProjType.txt2.Focus()
            diaProjType.txt2.SelectAll()
            diaProjType.StartPosition = FormStartPosition.CenterScreen
            diaProjType.Show()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Public Sub RefreshListToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshListToolStripMenuItem.Click
        GridView1.DataSource = ""
        GridView1.Refresh()

        frmProjectTypes_Load(Nothing, Nothing)
    End Sub

    Private Sub GridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridView1.CellContentDoubleClick
        Dim iRowIndex As Integer

        iRowIndex = GridView1.CurrentRow.Index

        Dim jobDesc As String = GridView1.CurrentRow.Cells("Description").Value

        MessageBox.Show(jobDesc, "Project Description", MessageBoxButtons.OK)
    End Sub
End Class