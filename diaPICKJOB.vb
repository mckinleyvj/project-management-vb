Imports System.IO
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.Sql
Imports System.Xml
Imports System.Data.SqlClient
'Imports System.Windows.Forms.PrintPreviewDialog
Imports System.Xml.Xsl

Public Class diaPICKJOB

    Dim sqlConn As String = frmConfig.theConnection
    Dim connect As OdbcConnection

    Dim headerFont As New Font("Segoe UI", 9, FontStyle.Bold)
    Dim detailFont As New Font("Segoe UI", 8, FontStyle.Regular)

    Private theCode As String
    Private theDesc As String
    Private theType As String

    Public Sub loadGridview()

        Dim myDataset5 As DataSet
        Dim command5 As OdbcCommand
        Dim myAdapter5 As OdbcDataAdapter

        connect = New OdbcConnection(sqlConn)
        connect.Open()

        Dim theValue As String = txtDB.Text.ToString
        Dim comboValue As String = cmbType.Text.ToString.Trim
        Dim commandCode As String = ""

        If comboValue = "Code" Then
            commandCode = "AND job_type LIKE '" & theValue.ToString.Trim & "%' "
        End If

        If comboValue = "Name" Then
            commandCode = "AND job_name LIKE '%" & theValue.ToString.Trim & "%' "
        End If

        Try

            Dim str_Jobs_show As String = "SELECT job_type, job_name, adjust_type FROM m_jobs_type where deleted = 0 " & commandCode & " order by job_type;"

            command5 = New OdbcCommand(str_Jobs_show, connect)

            myDataset5 = New DataSet()
            myDataset5.Tables.Clear()
            myAdapter5 = New OdbcDataAdapter()
            myAdapter5.SelectCommand = command5
            myAdapter5.Fill(myDataset5, "Db")

            Dim dtRetrievedData As DataTable = myDataset5.Tables("Db")

            Dim dtData As New DataTable
            Dim dtDataRows As DataRow

            Dim chkboxcol As New DataGridViewCheckBoxColumn

            dtData.TableName = "ProjTable"
            dtData.Columns.Add("Code")
            dtData.Columns.Add("Description")
            dtData.Columns.Add("Type")

            For Each dtDataRows In dtRetrievedData.Rows

                Dim cod As String = dtDataRows("job_type").ToString().Trim()
                Dim nam As String = dtDataRows("job_name").ToString.Trim()
                Dim typ As String = dtDataRows("adjust_type").ToString.Trim()

                'dtData.Rows.Add(New Object() {del.ToString.Trim(), close.ToString.Trim(), dat.ToString("dd-MM-yyyy"), ref.ToString.Trim(), cust.ToString.Trim(), custName.ToString.Trim(), proj.ToString.Trim(), projName.ToString.Trim(), ttlHrs.ToString.Trim(), ttlMins.ToString.Trim(), theDoc.ToString.Trim(), theRef.ToString.Trim(), theStat.ToString.Trim(), descp.ToString.Trim()})
                dtData.Rows.Add(New Object() {cod.ToString, nam.ToString, typ.ToString})
            Next

            Dim finalDataSet As New DataSet
            finalDataSet.Tables.Add(dtData)

            GridView1.DataSource = finalDataSet.Tables(0)
            'GridView1.ColumnHeadersVisible = False
            GridView1.Columns.Item("Code").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Code").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("Code").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("Code").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Code").Width = 75

            GridView1.Columns.Item("Description").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Description").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("Description").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("Description").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Description").FillWeight = 150
            GridView1.Columns.Item("Description").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

            GridView1.Columns.Item("Type").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Type").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("Type").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("Type").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Type").Width = 50

            'GridView1.ColumnHeadersDefaultCellStyle.Font = headerFont
            GridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            'GridView1.AllowUserToResizeColumns = True
            GridView1.Refresh()
            GridView1.ClearSelection()

            connect.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        connect.Close()

    End Sub

    Private Sub diaPICKJOB_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        diaJOBSHEET.GridView1.Enabled = True
    End Sub

    Private Sub diaPICKJOB_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            If txtJobDesc.ReadOnly = True Then
                Me.Close()
            Else
                'Do nothing
            End If

        End If
    End Sub

    Private Sub diaPICKJOB_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If cmbType.Items.Count <> 0 Then
            cmbType.SelectedIndex = 0
        End If

        btnSearch_Click(Nothing, Nothing)

        txtJobDesc.Text = ""
        txtJobDesc.ReadOnly = True
        btnCancel.Visible = False
        btnDetlSave.Visible = False
        txtDB.Select()
        txtDB.Focus()

    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click

        loadGridview()
 
    End Sub

    Public Sub select_item()
        Dim iRowIndex As Integer

        iRowIndex = GridView1.CurrentRow.Index

        theCode = GridView1.CurrentRow.Cells("Code").Value
        theDesc = GridView1.CurrentRow.Cells("Description").Value
        theType = GridView1.CurrentRow.Cells("Type").Value

        diaJOBSHEET.txtCustomer.Text = ""
        diaJOBSHEET.txtProject.Text = ""
        diaJOBSHEET.txtRemarks.Text = ""
        diaJOBSHEET.txtEstHours.Text = ""
        diaJOBSHEET.txtEstMinutes.Text = ""

        diaJOBSHEET.Label7.Text = theCode.ToString.Trim
        diaJOBSHEET.Label8.Text = theDesc.ToString.Trim
        diaJOBSHEET.lblType.Text = theType.ToString.Trim

        diaJOBSHEET.btnAddJob.Enabled = False
        diaJOBSHEET.btnAddProj.Enabled = False

        diaJOBSHEET.btnDetlCancel.Visible = False
        diaJOBSHEET.btnDetlSave.Visible = False
        diaJOBSHEET.txtJobDesc.Focus()
        diaJOBSHEET.txtJobDesc.Select()
        diaJOBSHEET.txtJobDesc.ReadOnly = False
        diaJOBSHEET.txtJobDesc.Text = txtJobDesc.Text.ToString
        diaJOBSHEET.dtpFrom.Enabled = True
        diaJOBSHEET.dtpTo.Enabled = True
        diaJOBSHEET.dtpFrom.Text = dtpFrom.Text.ToString
        diaJOBSHEET.dtpTo.Text = dtpTo.Text.ToString

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub GridView1_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridView1.CellClick

        Dim i As Integer
        Dim count As Integer = GridView1.Rows.Count

        i = 0
        If count > 0 Then
            txtJobDesc.ReadOnly = False
            btnCancel.Visible = True
            btnDetlSave.Visible = True
            dtpFrom.Enabled = True
            dtpTo.Enabled = True
        Else
            txtJobDesc.ReadOnly = True
            btnCancel.Visible = False
            btnDetlSave.Visible = False
            dtpFrom.Enabled = False
            dtpTo.Enabled = False
        End If

    End Sub

    Private Sub GridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridView1.CellContentDoubleClick
        Dim i As Integer
        Dim count As Integer = GridView1.Rows.Count

        i = 0
        If count > 0 Then
            txtJobDesc.ReadOnly = False
            btnCancel.Visible = True
            btnDetlSave.Visible = True
            dtpFrom.Enabled = True
            dtpTo.Enabled = True
        Else
            txtJobDesc.ReadOnly = True
            btnCancel.Visible = False
            btnDetlSave.Visible = False
            dtpFrom.Enabled = False
            dtpTo.Enabled = False
        End If
    End Sub

    Private Sub btnDetlSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDetlSave.Click

        If txtJobDesc.Text = "" Then
            validation1.Visible = True
            MessageBox.Show("Cannot leave `Job Description` empty. Please fill in.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            txtJobDesc.Focus()
        Else
            select_item()
            diaJOBSHEET.btnJobSave_Click(Nothing, Nothing)
            Me.Close()
        End If

    End Sub

    Private Sub dtpTo_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dtpTo.Validating
        Dim time_from As DateTime = dtpFrom.Text.ToString.Trim
        Dim time_to As DateTime = dtpTo.Text.ToString.Trim

        If time_to <= time_from Then
            MessageBox.Show("`Time To` cannot be less/equal to `Time From`", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            dtpTo.Text = time_from.AddMinutes(1)
            dtpTo.Focus()
        End If
    End Sub

End Class