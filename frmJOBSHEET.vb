Imports System.IO
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.Sql
Imports System.Xml
Imports System.Data.SqlClient
'Imports System.Windows.Forms.PrintPreviewDialog
Imports System.Xml.Xsl

Public Class frmJOBSHEET

    Dim sqlConn As String = frmConfig.theConnection
    Dim connect As OdbcConnection

    Public getLength As String

    Dim theLoginUser As String = frmMain.theUser.ToString.Trim()

    Dim headerFont As New Font("Segoe UI", 9, FontStyle.Bold)
    Dim detailFont As New Font("Segoe UI", 8, FontStyle.Regular)

    Private Sub CloseWindowToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseWindowToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub frmJOBSHEET_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Public Sub SetDateTime()
        dtpFrom.Format = DateTimePickerFormat.Custom
        dtpTo.Format = DateTimePickerFormat.Custom

        dtpFrom.CustomFormat = "dd/MM/yyyy"
        dtpTo.CustomFormat = "dd/MM/yyyy"

        'dtpFrom.Value = Now.Date.AddDays(-7)
        dtpFrom.Value = Now.Date.AddMonths(-1)
        dtpTo.Value = Now.Date
    End Sub

    Private Sub frmJOBSHEET_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged
        'GroupBox1.Width = Me.Width - 30
        'GroupBox2.Width = Me.Width - 30

        'GridView1.Width = GroupBox2.Width - 15

        'GroupBox2.Height = Me.Height - 175
        'GridView1.Height = GroupBox2.Height - 50


        btnSearch.Location = New Point(GroupBox1.Width - 125, Me.btnSearch.Location.Y)
    End Sub

    Private Sub frmJOBSHEET_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.BackColor = Color.White
        Me.ForeColor = Color.Black
        MenuStrip1.BackColor = Color.Gainsboro
        MenuStrip1.ForeColor = Color.Black

        SetDateTime()

        dtpFrom.Focus()

        Me.Text = "Job Sheet Listing - " & theLoginUser

        If theLoginUser.ToString = "SUPERVISOR" Then
            NewJobSheetToolStripMenuItem.Enabled = False
            EditJobSheetToolStripMenuItem.Enabled = False
            ViewJobSheetToolStripMenuItem.Enabled = True
            'CloseJobSheetToolStripMenuItem.Visible = True
        Else
            NewJobSheetToolStripMenuItem.Enabled = True
            EditJobSheetToolStripMenuItem.Enabled = True
            ViewJobSheetToolStripMenuItem.Enabled = True
            'CloseJobSheetToolStripMenuItem.Visible = False
        End If

        btnSearch_Click(Nothing, Nothing)

        'Dim no_ofRows As Integer
        'GridView1.FirstDisplayedScrollingRowIndex = GridView1.RowCount - 1
        'GridView1.ClearSelection()
        'GridView1.CurrentCell = GridView1.Rows(no_ofRows).Cells(0)
        'GridView1.Rows(GridView1.RowCount - 1).Selected = True
        'GridView1.Refresh()

    End Sub

    Public Sub getRunno()
        Dim com2 As OdbcCommand
        Dim adpt2 As OdbcDataAdapter
        Dim DS2 As DataSet

        connect = New OdbcConnection(sqlConn)

        Try

            Dim str_runno = "SELECT doc_doc, run_no, length FROM m_runno where doc_doc = 'JB';"

            connect.Open()
            com2 = New OdbcCommand(str_runno, connect)

            DS2 = New DataSet()
            DS2.Tables.Clear()
            adpt2 = New OdbcDataAdapter()
            adpt2.SelectCommand = com2
            adpt2.Fill(DS2, "Runno")

            diaJOBSHEET.lblDocCode.Text = DS2.Tables("Runno").Rows(0)(0).ToString
            Dim getNumber As String = DS2.Tables("Runno").Rows(0)(1).ToString
            getLength = DS2.Tables("Runno").Rows(0)(2).ToString

            Dim finalSetting As String = getNumber.PadLeft(getLength, "0")
            diaJOBSHEET.txtDocRef.Text = finalSetting.ToString.Trim
        Catch ex As Exception
            MessageBox.Show("Something went wrong. Please check below error message. " & vbCrLf & vbCrLf & ex.ToString)
        End Try
    End Sub

    Private Sub NewJobSheetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewJobSheetToolStripMenuItem.Click

        diaJOBSHEET.Close()
        diaJOBSHEET.lblTimeSpent.Text = "00:00"

        getRunno()

        diaJOBSHEET.SaveProjectToolStripMenuItem.Visible = False
        diaJOBSHEET.EditJobsheetToolStripMenuItem.Visible = False
        diaJOBSHEET.CancelEditToolStripMenuItem.Visible = False
        diaJOBSHEET.PRINTToolStripMenuItem.Visible = False
        diaJOBSHEET.txtRemarks1.Enabled = True

        diaJOBSHEET.StartPosition = FormStartPosition.CenterScreen
        diaJOBSHEET.ShowDialog()
    End Sub

    Public Sub refresh_gridview1()

        Dim dateFrom As DateTime = dtpFrom.Text.ToString
        Dim dateFromStr As String = dateFrom.ToString("yyyy-MM-dd")
        Dim dateTo As DateTime = dtpTo.Text.ToString
        Dim dateToStr As String = dateTo.ToString("yyyy-MM-dd")

        Dim empString As String

        If txtEmployee.Text = "" Then
            empString = ""
        Else
            empString = " AND J.username LIKE '" & txtEmployee.Text.ToUpper & "%' "
        End If

        Dim command1 As OdbcCommand
        Dim myAdapter As OdbcDataAdapter
        Dim myDataSet As DataSet

        connect = New OdbcConnection(sqlConn)

        Try

            Dim str_jobsheet_show As String

            If theLoginUser.ToString = "SUPERVISOR" Then
                str_jobsheet_show = "SELECT J.deleted as del, J.completed as complete, J.job_date as jDate, J.job_code as jCode, J.job_number as JNo, " _
                                              & "concat(J.job_code, '-', J.job_number) as refNo, J.username as uName, S.emp_name1 as name, J.ttl_hours as hours, J.ttl_minutes as mins " _
                                              & "from t_jobsheet_trn J, m_employee S " _
                                              & "where J.username = S.username " _
                                              & " AND J.job_date >= '" & dateFromStr & "' AND J.job_date <= '" & dateToStr & "' " & empString.ToString & " ORDER BY jDate, JNo;"
            Else
                str_jobsheet_show = "SELECT J.deleted as del, J.completed as complete, J.job_date as jDate, J.job_code as jCode, J.job_number as JNo, " _
                                             & "concat(J.job_code, '-', J.job_number) as refNo, J.username as uName, S.emp_name1 as name, J.ttl_hours as hours, J.ttl_minutes as mins " _
                                             & "from t_jobsheet_trn J, m_employee S " _
                                             & "where J.username = S.username and J.username = '" & theLoginUser.ToString.ToUpper.Trim & "' " _
                                             & " AND J.job_date >= '" & dateFromStr & "' AND J.job_date <= '" & dateToStr & "' " & empString.ToString & " ORDER BY jDate, JNo;"
            End If

            connect.Open()
            command1 = New OdbcCommand(str_jobsheet_show, connect)

            myDataSet = New DataSet()
            myDataSet.Tables.Clear()
            myAdapter = New OdbcDataAdapter()
            myAdapter.SelectCommand = command1
            myAdapter.Fill(myDataSet, "Db")

            Dim dtRetrievedData As DataTable = myDataSet.Tables("Db")

            Dim dtData As New DataTable
            Dim dtDataRows As DataRow

            dtData.TableName = "JSTable"
            dtData.Columns.Add("X", GetType(Boolean))
            dtData.Columns.Add("C", GetType(Boolean))
            dtData.Columns.Add("Date")
            dtData.Columns.Add("Job No")
            dtData.Columns.Add("User")
            dtData.Columns.Add("Name")
            dtData.Columns.Add("       Spent")
            dtData.Columns.Add("Doc")
            dtData.Columns.Add("Ref")

            For Each dtDataRows In dtRetrievedData.Rows

                Dim del As Boolean = dtDataRows("del").ToString().Trim()
                Dim complete As Boolean = dtDataRows("complete").ToString().Trim()
                Dim dat As Date = dtDataRows("jDate").ToString().Trim()
                'Dim wordsDate As String = dat.ToString("dddd, dd-MM-yyyy")
                Dim ref As String = dtDataRows("refNo").ToString.Trim()

                Dim userN As String = dtDataRows("uName").ToString.Trim()
                Dim userName As String = dtDataRows("name").ToString.Trim()

                'TIME TAKEN
                Dim hr As Double = Convert.ToDouble(dtDataRows("hours").ToString.Trim())
                Dim min As Double = Convert.ToDouble(dtDataRows("mins").ToString.Trim())
                Dim ttlHrs As Double = hr
                Dim ttlMins As Double
                ttlMins = min

                Do While ttlMins >= 60
                    ttlMins = ttlMins - 60
                    ttlHrs = ttlHrs + 1
                Loop

                Dim ttlMin As Double = ttlMins
                Dim ttlTime As Double = ttlHrs + ttlMin
                Dim totalTime As String = ttlHrs.ToString & ":" & ttlMin.ToString("00")

                Dim theDoc As String = dtDataRows("jCode").ToString.Trim()
                Dim theRef As String = dtDataRows("JNo").ToString.Trim()

                'dtData.Rows.Add(New Object() {del.ToString.Trim(), close.ToString.Trim(), dat.ToString("dd-MM-yyyy"), ref.ToString.Trim(), cust.ToString.Trim(), custName.ToString.Trim(), proj.ToString.Trim(), projName.ToString.Trim(), ttlHrs.ToString.Trim(), ttlMins.ToString.Trim(), theDoc.ToString.Trim(), theRef.ToString.Trim(), theStat.ToString.Trim(), descp.ToString.Trim()})
                dtData.Rows.Add(New Object() {del.ToString.Trim(), complete.ToString.Trim(), dat.ToString("dd/MM/yyyy, dd MMMM yyyy - dddd"), ref.ToString.Trim(), userN.ToString.Trim(), userName.ToString.Trim(), totalTime.ToString.Trim(), theDoc.ToString.Trim(), theRef.ToString.Trim()})
            Next

            Dim finalDataSet As New DataSet
            finalDataSet.Tables.Add(dtData)

            GridView1.DataSource = finalDataSet.Tables(0)
            'GridView1.ColumnHeadersVisible = False
            GridView1.Columns.Item("X").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("X").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("X").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("X").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("X").Resizable = DataGridViewTriState.False
            GridView1.Columns.Item("X").Width = 19

            GridView1.Columns.Item("C").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("C").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("C").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("C").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("C").Resizable = DataGridViewTriState.False
            GridView1.Columns.Item("C").Width = 19

            GridView1.Columns.Item("Date").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Date").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("Date").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("Date").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Date").Width = 250

            GridView1.Columns.Item("Job No").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Job No").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("Job No").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("Job No").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Job No").Width = 100

            GridView1.Columns.Item("User").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("User").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("User").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("User").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("User").Width = 100

            GridView1.Columns.Item("Name").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Name").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("Name").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("Name").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Name").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            GridView1.Columns.Item("Name").Width = 100

            GridView1.Columns.Item("       Spent").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("       Spent").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("       Spent").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("       Spent").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            GridView1.Columns.Item("       Spent").Resizable = DataGridViewTriState.False
            GridView1.Columns.Item("       Spent").Width = 67

            GridView1.Columns.Item("Doc").Visible = False
            GridView1.Columns.Item("ReF").Visible = False

            'GridView1.ColumnHeadersDefaultCellStyle.Font = headerFont
            GridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            'GridView1.AllowUserToResizeColumns = True
            GridView1.Refresh()

            connect.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        connect.Close()
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click

        refresh_gridview1()

    End Sub

    Public Sub disableAll()
        diaJOBSHEET.Label7.Text = ""
        diaJOBSHEET.Label8.Text = ""
        diaJOBSHEET.txtTimeFrom.Text = ""
        diaJOBSHEET.txtTimeTo.Text = ""
        diaJOBSHEET.txtJobDesc.Text = ""

        diaJOBSHEET.txtCustomer.Text = ""
        diaJOBSHEET.txtProject.Text = ""
        diaJOBSHEET.txtRemarks.Text = ""
        diaJOBSHEET.txtEstHours.Text = ""
        diaJOBSHEET.txtEstMinutes.Text = ""

        diaJOBSHEET.validation1.Visible = False
        diaJOBSHEET.validation2.Visible = False

        diaJOBSHEET.btnAddJob.Visible = False
        diaJOBSHEET.btnAddProj.Visible = False
        'btnPick.Focus()

        diaJOBSHEET.btnDetlCancel.Visible = False
        diaJOBSHEET.btnDelJob.Visible = False
        diaJOBSHEET.btnDetlEdit.Visible = False
        diaJOBSHEET.btnDetlSave.Visible = False
        diaJOBSHEET.txtJobDesc.ReadOnly = True
        diaJOBSHEET.txtTimeFrom.Enabled = False
        diaJOBSHEET.txtTimeTo.Enabled = False

        diaJOBSHEET.dtpDate.Enabled = False

        'diaJOBSHEET.SaveProjectToolStripMenuItem.Enabled = False

    End Sub

    Private Sub ViewJobSheetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewJobSheetToolStripMenuItem.Click

        Dim theLoginUser As String = frmMain.theUser.ToString.Trim()

        Dim theUser As String

        Try

            Dim command1 As OdbcCommand
            Dim myAdapter As OdbcDataAdapter
            Dim myDataSet As DataSet

            connect = New OdbcConnection(sqlConn)

            Dim str_Employee_Show As String = "SELECT username, emp_name1 FROM m_employee where username = '" & theLoginUser.ToString.Trim.ToUpper & "';"

            connect.Open()
            command1 = New OdbcCommand(str_Employee_Show, connect)

            myDataSet = New DataSet()
            myDataSet.Tables.Clear()
            myAdapter = New OdbcDataAdapter()
            myAdapter.SelectCommand = command1
            myAdapter.Fill(myDataSet, "Db")

            If myDataSet.Tables("Db").Rows.Count = 0 Then
                theUser = ""
            Else
                theUser = myDataSet.Tables("Db").Rows(0)(0).ToString
            End If

            If theUser = "" Then

                If GridView1.RowCount.ToString <> 0 Then
                    Dim theDoc As String = GridView1.CurrentRow.Cells("DoC").Value
                    Dim theRef As String = GridView1.CurrentRow.Cells("ReF").Value
                    Dim theDate As String = GridView1.CurrentRow.Cells("Date").Value

                    diaJOBSHEET.Close()

                    diaJOBSHEET.Label12.Text = theDoc.ToString.Trim.ToUpper
                    diaJOBSHEET.Label13.Text = theRef.ToString.Trim.ToUpper
                    diaJOBSHEET.lblJobDate.Text = theDate.ToString.Trim

                    disableAll()
                    diaJOBSHEET.EditJobsheetToolStripMenuItem.Visible = False
                    diaJOBSHEET.CancelEditToolStripMenuItem.Visible = False
                    diaJOBSHEET.SaveProjectToolStripMenuItem.Visible = False
                    diaJOBSHEET.PRINTToolStripMenuItem.Visible = True
                    diaJOBSHEET.txtRemarks1.Enabled = False

                    diaJOBSHEET.StartPosition = FormStartPosition.CenterScreen
                    diaJOBSHEET.ShowDialog()
                Else
                    MessageBox.Show("No Record(s) found.")
                End If

            Else

                If GridView1.RowCount.ToString <> 0 Then
                    Dim theDoc As String = GridView1.CurrentRow.Cells("DoC").Value
                    Dim theRef As String = GridView1.CurrentRow.Cells("ReF").Value
                    Dim theDate As String = GridView1.CurrentRow.Cells("Date").Value

                    diaJOBSHEET.Close()

                    diaJOBSHEET.Label12.Text = theDoc.ToString.Trim.ToUpper
                    diaJOBSHEET.Label13.Text = theRef.ToString.Trim.ToUpper
                    diaJOBSHEET.lblJobDate.Text = theDate.ToString.Trim

                    diaJOBSHEET.CancelEditToolStripMenuItem.Visible = False
                    diaJOBSHEET.SaveProjectToolStripMenuItem.Visible = False
                    diaJOBSHEET.EditJobsheetToolStripMenuItem.Visible = True
                    diaJOBSHEET.PRINTToolStripMenuItem.Visible = True
                    diaJOBSHEET.txtRemarks1.Enabled = False

                    disableAll()

                    diaJOBSHEET.StartPosition = FormStartPosition.CenterScreen
                    diaJOBSHEET.ShowDialog()
                Else
                    MessageBox.Show("No Record(s) found.")
                End If

            End If

            connect.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub EditJobSheetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditJobSheetToolStripMenuItem.Click
        If GridView1.RowCount.ToString <> 0 Then
            Dim theDoc As String = GridView1.CurrentRow.Cells("DoC").Value
            Dim theRef As String = GridView1.CurrentRow.Cells("ReF").Value
            Dim theDate As String = GridView1.CurrentRow.Cells("Date").Value

            diaJOBSHEET.Close()

            diaJOBSHEET.Label12.Text = theDoc.ToString.Trim.ToUpper
            diaJOBSHEET.Label13.Text = theRef.ToString.Trim.ToUpper
            diaJOBSHEET.lblJobDate.Text = theDate.ToString.Trim

            disableAll()

            diaJOBSHEET.btnAddJob.Visible = True
            diaJOBSHEET.btnAddProj.Visible = True
            diaJOBSHEET.btnDetlEdit.Visible = True
            diaJOBSHEET.btnDelJob.Visible = True
            diaJOBSHEET.dtpDate.Enabled = True
            diaJOBSHEET.SaveProjectToolStripMenuItem.Visible = True
            diaJOBSHEET.EditJobsheetToolStripMenuItem.Visible = False
            diaJOBSHEET.PRINTToolStripMenuItem.Visible = False
            diaJOBSHEET.CancelEditToolStripMenuItem.Visible = True
            diaJOBSHEET.CloseWindowToolStripMenuItem.Visible = False
            diaJOBSHEET.txtRemarks1.Enabled = True

            diaJOBSHEET.StartPosition = FormStartPosition.CenterScreen
            diaJOBSHEET.ShowDialog()
        Else
            MessageBox.Show("No Record(s) found.")
        End If

    End Sub

    Private Sub PRINTToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PRINTToolStripMenuItem.Click
        If GridView1.RowCount.ToString <> 0 Then
            Dim theDoc As String = GridView1.CurrentRow.Cells("DoC").Value
            Dim theRef As String = GridView1.CurrentRow.Cells("ReF").Value
            'Dim theEmp As String = GridView1.CurrentRow.Cells("Name").Value
            'Dim theDate As String = GridView1.CurrentRow.Cells("Date").Value

            'diaJOBSHEET.Close()
            'diaJOBSHEET.getInfoData()

            'prnJOBSHT.lblJobCode.Text = theDoc
            'prnJOBSHT.lblJobNumber.Text = theRef
            'prnJOBSHT.lblEmployee.Text = theEmp
            'prnJOBSHT.lblJobDate.Text = theDate

            prnJOBSHT.lblJobCode.Text = theDoc
            prnJOBSHT.lblJobNumber.Text = theRef
            'prnJOBSHT.lblEmployee.Text = theEmp
            'prnJOBSHT.lblJobDate.Text = theDate
            'prnJOBSHT.lblRemarks.Text = diaJOBSHEET
            'prnJOBSHT.lblTotalHours.Text = totalJSTimeHr
            'prnJOBSHT.lblTotalMins.Text = totalJSTimeMin
            'prnJOBSHT.lblWorkHr.Text = totalWorkHr + " Hour(s)"
            'prnJOBSHT.lblWorkMin.Text = totalWorkMin + " Minute(s)"

            prnJOBSHT.StartPosition = FormStartPosition.CenterScreen
            prnJOBSHT.Show()
        Else
            MessageBox.Show("No Record(s) found.")
        End If
    End Sub

End Class