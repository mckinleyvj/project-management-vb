Imports System.IO
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.Sql
Imports System.Xml
Imports System.Data.SqlClient
'Imports System.Windows.Forms.PrintPreviewDialog
Imports System.Xml.Xsl

Public Class diaPICKPROJ

    Dim sqlConn As String = frmConfig.theConnection
    Dim connect As OdbcConnection

    Dim headerFont As New Font("Segoe UI", 9, FontStyle.Bold)
    Dim detailFont As New Font("Segoe UI", 8, FontStyle.Regular)

    Private theCustomer As String
    Private theCustCode As String
    Private theProject As String
    Private theRemarks As String
    Private theEstHours As String
    Private theEstMins As String
    Private thedoc As String
    Private theref As String

    Public Sub loadGridview()

        Dim myDataset5 As DataSet
        Dim command5 As OdbcCommand
        Dim myAdapter5 As OdbcDataAdapter

        connect = New OdbcConnection(sqlConn)
        connect.Open()

        Dim theValue As String = txtDB.Text.ToString
        Dim comboValue As String = cmbType.Text.ToString.Trim
        Dim commandCode As String = ""

        Dim jobDate1 As DateTime = diaJOBSHEET.dtpDate.Value
        Dim jobDate As String = jobDate1.ToString("yyyy-MM-dd")

        If comboValue = "Code" Then
            commandCode = "AND T.dbcode LIKE '" & theValue.ToString.Trim & "%' "
        End If

        If comboValue = "Name" Then
            commandCode = "AND M.db_name1 LIKE '%" & theValue.ToString.Trim & "%' "
        End If

        If comboValue = "Reference No" Then
            commandCode = "AND T.doc_ref = '" & theValue.ToString.Trim & "' "
        End If

        Try

            Dim str_Customers_Show As String = "SELECT T.doc_doc as DoC, T.doc_ref as ReF, concat(T.doc_doc, '-', T.doc_ref) AS referenceNumber, T.doc_date as theDate, T.dbcode as custCode, M.db_name1 as custName, " _
                                               & "T.project_code as projCode, T.desc as projDesc, T.est_hours as esthours, T.est_minutes as estmins " _
                                               & "FROM t_projects T, m_armaster M WHERE(T.dbcode = M.dbcode) " _
                                               & "AND T.closed = 0 and T.deleted = 0 and T.cancel = 0 and T.status = 'In Progress' and T.doc_date <= '" & jobDate & "' " & commandCode & ";"

            command5 = New OdbcCommand(str_Customers_Show, connect)

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
            dtData.Columns.Add("Date")
            dtData.Columns.Add("Ref No.")
            dtData.Columns.Add("Code")
            dtData.Columns.Add("Name")
            dtData.Columns.Add("Project")
            dtData.Columns.Add("Remarks")
            dtData.Columns.Add("DoC")
            dtData.Columns.Add("ReF")
            dtData.Columns.Add("Hours")
            dtData.Columns.Add("Minutes")

            For Each dtDataRows In dtRetrievedData.Rows

                Dim dat As Date = dtDataRows("theDate").ToString().Trim()
                Dim ref As String = dtDataRows("referenceNumber").ToString.Trim()
                Dim cust As String = dtDataRows("custCode").ToString.Trim()
                Dim custName As String = dtDataRows("custName").ToString.Trim()
                Dim proj As String = dtDataRows("projCode").ToString.Trim()
                Dim descp As String = dtDataRows("projDesc").ToString.Trim()

                'ESTIMATED TIME
                Dim hr1 As Double = Convert.ToDouble(dtDataRows("esthours").ToString.Trim())
                Dim min1 As Double = Convert.ToDouble(dtDataRows("estmins").ToString.Trim())
                Dim ttlHrs1 As Double = hr1
                Dim ttlMins1 As Double
                ttlMins1 = min1

                Do While ttlMins1 >= 60
                    ttlMins1 = ttlMins1 - 60
                    ttlHrs1 = ttlHrs1 + 1
                Loop

                Dim ttlMin1 As Double = ttlMins1 * 0.01
                Dim ttlTime1 As Double = ttlHrs1 + ttlMin1
                Dim totalTime1 As String = String.Format("{0:0.00}", ttlTime1)

                Dim theDoc As String = dtDataRows("DoC").ToString.Trim()
                Dim theRef As String = dtDataRows("ReF").ToString.Trim()

                'dtData.Rows.Add(New Object() {del.ToString.Trim(), close.ToString.Trim(), dat.ToString("dd-MM-yyyy"), ref.ToString.Trim(), cust.ToString.Trim(), custName.ToString.Trim(), proj.ToString.Trim(), projName.ToString.Trim(), ttlHrs.ToString.Trim(), ttlMins.ToString.Trim(), theDoc.ToString.Trim(), theRef.ToString.Trim(), theStat.ToString.Trim(), descp.ToString.Trim()})
                dtData.Rows.Add(New Object() {dat.ToString("dd-MM-yyyy"), ref.ToString.Trim(), cust.ToString.Trim(), custName.ToString.Trim(), proj.ToString.Trim(), descp.ToString.Trim(), theDoc.ToString.Trim(), theRef.ToString.Trim(), hr1.ToString, min1.ToString})
            Next

            Dim finalDataSet As New DataSet
            finalDataSet.Tables.Add(dtData)

            GridView1.DataSource = finalDataSet.Tables(0)
            'GridView1.ColumnHeadersVisible = False
            GridView1.Columns.Item("Date").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Date").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("Date").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("Date").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Date").Width = 70

            GridView1.Columns.Item("Ref No.").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Ref No.").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("Ref No.").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("Ref No.").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Ref No.").Width = 80

            GridView1.Columns.Item("Code").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Code").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("Code").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("Code").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Code").Width = 50

            GridView1.Columns.Item("Name").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Name").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("Name").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("Name").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Name").Width = 200

            GridView1.Columns.Item("Project").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Project").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("Project").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("Project").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Project").Width = 100

            GridView1.Columns.Item("Remarks").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Remarks").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("Remarks").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("Remarks").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Remarks").Width = 220
            GridView1.Columns.Item("Remarks").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            GridView1.Columns.Item("Remarks").Visible = True

            GridView1.Columns.Item("DoC").Visible = False
            GridView1.Columns.Item("ReF").Visible = False

            GridView1.Columns.Item("Hours").Visible = False
            GridView1.Columns.Item("Minutes").Visible = False

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

    Private Sub diaPICKPROJ_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        diaJOBSHEET.GridView1.Enabled = True
    End Sub

    Private Sub diaPICKPROJ_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            If txtJobDesc.ReadOnly = True Then
                Me.Close()
            Else
                'Do nothing
            End If

        End If
    End Sub

    Private Sub diaPICKPROJ_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

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

        theCustomer = GridView1.CurrentRow.Cells("Name").Value
        theCustCode = GridView1.CurrentRow.Cells("Code").Value
        theProject = GridView1.CurrentRow.Cells("Project").Value
        theRemarks = GridView1.CurrentRow.Cells("Remarks").Value
        theEstHours = GridView1.CurrentRow.Cells("Hours").Value
        theEstMins = GridView1.CurrentRow.Cells("Minutes").Value
        thedoc = GridView1.CurrentRow.Cells("DoC").Value
        theref = GridView1.CurrentRow.Cells("ReF").Value

        diaJOBSHEET.txtCustomer.Text = theCustomer.ToString.Trim
        diaJOBSHEET.lblCustCode.text = theCustCode.ToString.Trim
        diaJOBSHEET.txtProject.Text = theProject.ToString.Trim
        diaJOBSHEET.txtRemarks.Text = theRemarks.ToString.Trim
        diaJOBSHEET.txtEstHours.Text = theEstHours.ToString.Trim
        diaJOBSHEET.txtEstMinutes.Text = theEstMins.ToString.Trim

        diaJOBSHEET.Label7.Text = thedoc.ToString.Trim
        diaJOBSHEET.Label8.Text = theref.ToString.Trim


        diaJOBSHEET.btnAddJob.Enabled = False
        '
        diaJOBSHEET.btnDetlCancel.Visible = False
        diaJOBSHEET.btnDetlSave.Visible = False
        '
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
            diaJOBSHEET.btnDetlSave_Click(Nothing, Nothing)
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