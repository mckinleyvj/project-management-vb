Imports System.IO
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.Sql
Imports System.Xml
Imports System.Data.SqlClient
'Imports System.Windows.Forms.PrintPreviewDialog
Imports System.Xml.Xsl
Imports System.Drawing.Printing
Imports System.Text
Imports Excel = Microsoft.Office.Interop.Excel

Public Class diaJSListing

    Dim sqlConn As String = frmConfig.theConnection
    Dim connect As OdbcConnection

    Dim theLoginUser As String = frmMain.theUser.ToString.Trim()

    Dim headerFont As New Font("Segoe UI", 9, FontStyle.Bold)
    Dim detailFont As New Font("Segoe UI", 8, FontStyle.Regular)

    Dim mRow As Integer = 0
    Dim mCol As Integer = 0
    Dim newpage As Boolean = True

    Dim SummaryDataSet As DataSet

    Dim expPath As String

    Private Sub diaJSListing_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        dtpFrom.Value = Now.Date.AddDays(-7)

        loadEmployee()

        cbTickAll.Enabled = False
        cbTickAll1.Enabled = False
        cbDefault.Enabled = False
        cbPlus.Enabled = False
        cbMinus.Enabled = False
        'rbSummary.Checked = True
        'rbSummary_CheckedChanged(Nothing, Nothing)

        'loadSummaryColumn()
        'loadFilters()

    End Sub

    Public Sub loadFilters()

        cbTickAll1.Enabled = True
        cbPlus.Enabled = True
        cbMinus.Enabled = True

        Dim theEmp As String = cmbEmployee.SelectedValue
        'Dim dateFrom As DateTime = dtpFrom.Value
        'Dim dateFromString As String = dateFrom.ToString("yyyy-MM-dd")
        'Dim dateTo As DateTime = dtpTo.Value
        'Dim dateToString As String = dateTo.ToString("yyyy-MM-dd")

        Try
            Dim command2 As OdbcCommand
            Dim myAdapter2 As OdbcDataAdapter
            Dim myDataSet2 As DataSet

            Dim work_type As String = "SELECT DISTINCT(D.work_project_type) as 'work_type', D.adjust_type as 'sign' " _
                                        & "FROM t_jobsheet_trn T, t_jobsheet_detl D " _
                                        & "WHERE T.job_code = D.job_code " _
                                        & "AND T.job_number = D.job_number " _
                                        & "AND T.username = '" & theEmp.ToString.Trim & "' " _
                                        & "AND D.deleted = 0 " _
                                        & "ORDER BY D.work_project_type;"

            command2 = New OdbcCommand(work_type, connect)

            myDataSet2 = New DataSet()
            myDataSet2.Tables.Clear()
            myAdapter2 = New OdbcDataAdapter()
            myAdapter2.SelectCommand = command2
            myAdapter2.Fill(myDataSet2, "work_type")

            Dim dt As New DataTable()
            Dim dr As DataRow

            Dim colCheckbox As CheckBox = New CheckBox()

            dt.Columns.Add("Select", GetType(Boolean))
            dt.Columns.Add("work_type")
            dt.Columns.Add("sign")

            Dim k As Integer = 0

            Do While k < myDataSet2.Tables("work_type").Rows.Count
                dr = dt.NewRow()
                dr("work_type") = myDataSet2.Tables("work_type").Rows(k)("work_type").ToString
                dr("sign") = myDataSet2.Tables("work_type").Rows(k)("sign").ToString
                dt.Rows.Add(dr)
                k += 1
            Loop

            'dgvFilter.Visible = True
            dgvFilter.DataSource = dt
            dgvFilter.Columns(0).Width = 20
            dgvFilter.Columns(1).FillWeight = 100
            dgvFilter.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            dgvFilter.Columns(2).Width = 20
            dgvFilter.Refresh()

            connect.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Public Sub loadEmployee()

        Dim command1 As OdbcCommand
        Dim myAdapter As OdbcDataAdapter
        Dim myDataSet As DataSet

        connect = New OdbcConnection(sqlConn)

        Try

            Dim emp_list As String = "SELECT DISTINCT username FROM t_jobsheet_trn where deleted = 0;"

            connect.Open()
            command1 = New OdbcCommand(emp_list, connect)

            myDataSet = New DataSet()
            myDataSet.Tables.Clear()
            myAdapter = New OdbcDataAdapter()
            myAdapter.SelectCommand = command1
            myAdapter.Fill(myDataSet, "emp")

            cmbEmployee.DataSource = myDataSet.Tables("emp")
            cmbEmployee.DisplayMember = "username"
            cmbEmployee.ValueMember = "username"
            cmbEmployee.Refresh()

            Dim i As Integer = 0

            Do While i < myDataSet.Tables("emp").Rows.Count()

                If myDataSet.Tables("emp").Rows(i)(0) = theLoginUser.ToString Then
                    cmbEmployee.SelectedValue = myDataSet.Tables("emp").Rows(i)(0).ToString
                Else
                    'skip. do nothing
                    'MessageBox.Show("It is not the user. " & myDataSet.Tables("emp").Rows(i)(0).ToString)
                End If
                i += 1
            Loop

            connect.Close()

            '    'lblRowCount.Text = Gridview1.Rows.Count
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        connect.Close()

    End Sub

    Private Sub btnGenerate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenerate.Click

        dgvReport.DataSource = ""
        dgvReport.Refresh()

        If rbSummary.Checked = False And rbDetailed.Checked = False Then
            MessageBox.Show("Select a Report type", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        ElseIf rbSummary.Checked = True And rbDetailed.Checked = False Then

            loadSummaryReport()

        ElseIf rbSummary.Checked = False And rbDetailed.Checked = True Then

            loadDetailedReport()

        ElseIf rbSummary.Checked = True And rbDetailed.Checked = True Then
            MessageBox.Show("This shouldn't happen. Please contact Developer.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

    End Sub

    Private Sub diaJSListing_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize

        If Me.WindowState = FormWindowState.Maximized Then

            dgvReport.Width = Me.Width - GroupBox1.Width - 45
            grpSearch.Width = Me.Width - 40
            'MessageBox.Show(Me.Width)
            dgvReport.Height = Me.Height - grpSearch.Height - 60
            btnGenerate.Location = New Point(grpSearch.Right - 214, 18)
            btnExport.Location = New Point(grpSearch.Right - 80, 18)

        ElseIf Me.WindowState = FormWindowState.Normal Then

            dgvReport.Width = Me.Width - GroupBox1.Width - 45
            grpSearch.Width = Me.Width - 40

            dgvReport.Height = Me.Height - grpSearch.Height - 60

            btnGenerate.Location = New Point(742, 18)
            btnExport.Location = New Point(876, 18)

        End If
    End Sub

    Private Sub btnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPreview.Click

        If dgvReport.Rows.Count = 0 Then
            MessageBox.Show("Nothing to print.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Dim ppd As New PrintPreviewDialog
            ppd.Document = PrintDocument
            ppd.WindowState = FormWindowState.Maximized
            ppd.ShowDialog()
            'Me.Close()
        End If

    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        PrintDialog.Document = PrintDocument
        PrintDialog.PrinterSettings = PrintDocument.PrinterSettings
        PrintDialog.AllowSomePages = True
        PrintDialog.AllowSelection = True
        If PrintDialog.ShowDialog = DialogResult.OK Then
            PrintDocument.PrinterSettings = PrintDialog.PrinterSettings
            PrintDocument.Print()
        End If
        'Me.Close()
    End Sub

    Private Sub PrintDocument_BeginPrint(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintEventArgs) Handles PrintDocument.BeginPrint

        PrintDocument.DefaultPageSettings.Landscape = True
        'PrintDocument.DefaultPageSettings.PaperSize = New PaperSize("Letter", 850, 1100)
        PrintDocument.DefaultPageSettings.PaperSize = New PaperSize("A4", 827, 1169)
        'If use A4, width is longer
        PrintDocument.DefaultPageSettings.Margins = New Margins(0, 0, 0, 0)

    End Sub

    Private Sub PrintDocument_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument.PrintPage
        Dim fromDate As String = dtpFrom.Value.ToString("dd/MM/yyyy")
        Dim toDate As String = dtpTo.Value.ToString("dd/MM/yyyy")

        Try

            Dim drawBrush As New SolidBrush(Color.Black)
            Dim drawBrushJSNumber As New SolidBrush(Color.Red)
            Dim _titlefont As Font = New Drawing.Font("Segoe UI", 14, FontStyle.Bold)
            Dim _gridheaderfont As Font = New Drawing.Font("Segoe UI", 8, FontStyle.Bold)
            Dim _griddetailfont As Font = New Drawing.Font("Segoe UI", 7, FontStyle.Regular)
            Dim _griddetailfont_adjust As Font = New Drawing.Font("Segoe UI", 6, FontStyle.Regular)

            Dim pageWidth = e.PageSettings.PrintableArea.Width
            Dim pageHeight = e.PageSettings.PrintableArea.Height - 528
            Dim footerPrint = e.PageSettings.PrintableArea.Height - 378
            'MessageBox.Show("Printable Height = " & pageHeight)
            Dim pageOrientation = e.PageSettings.Landscape = True
            Dim Lbound = e.PageSettings.PrintableArea.Left
            Dim Rbound = e.PageSettings.PrintableArea.Right
            Dim Tbound = e.PageSettings.PrintableArea.Top
            'Dim Bbound = e.PageSettings.PrinterSettings.DefaultPageSettings.Margins.Bottom
            'Dim Bbound = e.PageSettings.PrintableArea.Bottom
            'MessageBox.Show("BBound is " & Bbound)
            '250 is the optimal bottom margin

            'HEADING
            e.Graphics.DrawString("JOBSHEET LISTING REPORT", _titlefont, drawBrush, Lbound, Tbound)

            e.Graphics.DrawString("Date Printed    : " & Now.Date.ToString("dd/MM/yyyy"), _gridheaderfont, drawBrush, Rbound - 210, Tbound + 22)
            e.Graphics.DrawString("Employee : " & cmbEmployee.SelectedValue.ToString, _gridheaderfont, drawBrush, Lbound, Tbound + 22)
            e.Graphics.DrawString("Dates : " & fromDate & " To : " & toDate, _gridheaderfont, drawBrush, Lbound, Tbound + 35)

            Dim fmt As StringFormat = New StringFormat(StringFormatFlags.LineLimit)
            fmt.LineAlignment = StringAlignment.Near
            fmt.Trimming = StringTrimming.None
            Dim y As Integer = Tbound + 50
            Dim rc As Rectangle
            Dim x As Integer
            Dim h As Integer = 0
            Dim row As DataGridViewRow
            Dim dgvZZ As DataGridView = dgvReport

            'on new page, print the header columns only
            If newpage Then

                row = dgvZZ.Rows(mRow)
                x = Lbound
                For Each cell As DataGridViewCell In row.Cells

                    If cell.Visible = True Then

                        Dim theLength As Integer = dgvZZ.Columns(cell.ColumnIndex).Width
                        If dgvZZ.Columns(cell.ColumnIndex).HeaderText = "description" Then
                            theLength = Rbound - 210
                            'theLength = Rbound - 290 ''if use Letter papersize 
                        End If

                        rc = New Rectangle(x, y, theLength, 15)

                        e.Graphics.FillRectangle(Brushes.LightGray, rc)
                        e.Graphics.DrawRectangle(Pens.Black, rc)

                        e.Graphics.DrawString(dgvZZ.Columns(cell.ColumnIndex).HeaderText, _gridheaderfont, drawBrush, rc)

                        x += rc.Width
                        h = rc.Height
                    End If
                Next
                y += h

            End If
            newpage = False

            'now print the data for each row
            Dim thisNDX As Integer
            For thisNDX = mRow To dgvZZ.RowCount - 1
                ' no need to try to print the new row
                If dgvZZ.Rows(thisNDX).IsNewRow Then Exit For

                row = dgvZZ.Rows(thisNDX)
                x = Lbound
                h = 0

                ' reset X for data
                x = Lbound

                If y < pageHeight Then
                    'MessageBox.Show(y & " " & pageHeight)
                    ' print the data
                    For Each cell As DataGridViewCell In row.Cells
                        If cell.Visible = True Then

                            Dim theString As String = cell.FormattedValue.ToString()

                            Dim theSize As SizeF = e.Graphics.MeasureString(theString, _griddetailfont)

                            Dim theLength As Integer = dgvZZ.Columns(cell.ColumnIndex).Width
                            Dim theHeight As Integer = dgvZZ.Rows(cell.RowIndex).Height

                            If dgvZZ.Columns(cell.ColumnIndex).HeaderText = "description" Then
                                theLength = Rbound - 210
                                'theLength = Rbound - 290 ''if use Letter papersize 
                            End If

                            If theSize.Height < theHeight Then
                                theHeight = theSize.Height
                                If theHeight < cell.Size.Height Then
                                    theHeight = cell.Size.Height
                                End If
                            End If

                            rc = New Rectangle(x, y, theLength, theHeight)

                            e.Graphics.DrawRectangle(Pens.Black, rc)

                            e.Graphics.DrawString(theString, _griddetailfont, drawBrush, rc, fmt)

                            x += rc.Width
                            h = Math.Max(h, rc.Height)
                        End If

                    Next
                    y += h
                    ' next row to print
                    mRow = thisNDX + 1
                    newpage = False
                Else
                    'mRow = thisNDX + 1
                    e.HasMorePages = True
                    newpage = True
                    Return
                End If
                'If y + h > pageHeight Then
                '    e.HasMorePages = True
                '    'mRow -= 1   'causes last row to rePrint on next page
                '    Return
                'End If
            Next

            'END COMMENT
            'e.Graphics.DrawString("Jobsheet Date :", _gridheaderfont, drawBrush, Lbound, Tbound + 55)
            'e.Graphics.DrawString(lblJobDate.Text.ToString, _gridheaderfont, drawBrush, Lbound + 100, Tbound + 55)

            'e.Graphics.DrawString("Total Work Hours : " & lblTotalHours.Text & " " & lblTotalMins.Text, _gridheaderfont, drawBrush, Lbound, y + 10)
            e.Graphics.DrawString("**End of Report**", _gridheaderfont, drawBrush, Lbound, footerPrint)

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub

    Private Sub rbSummary_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbSummary.CheckedChanged

        dgvColumns.DataSource = ""
        dgvColumns.Refresh()

        Try

            Dim theEmp As String = cmbEmployee.SelectedValue
            Dim dateFrom As DateTime = dtpFrom.Value
            Dim dateFromString As String = dateFrom.ToString("yyyy-MM-dd")
            Dim dateTo As DateTime = dtpTo.Value
            Dim dateToString As String = dateTo.ToString("yyyy-MM-dd")

            Dim command1 As OdbcCommand
            Dim myAdapter As OdbcDataAdapter
            Dim myDataSet As DataSet

            connect = New OdbcConnection(sqlConn)

            Dim columnHeaders As String = "SELECT T.job_date, DAYNAME(T.job_date) as 'day', " _
                                          & "MIN(D.time_from) as 'first_task', " _
                                          & "MAX(D.time_to) as 'last_task', " _
                                          & "COUNT(D.job_number) as 'job_count', " _
                                          & "SEC_TO_TIME(SUM(TIME_TO_SEC(D.time_to) - TIME_TO_SEC(D.time_from))) AS 'totaltime', " _
                                          & "HOUR(SEC_TO_TIME(SUM(TIME_TO_SEC(D.time_to) - TIME_TO_SEC(D.time_from)))) as 'totaltime_hours', " _
                                          & "MINUTE(SEC_TO_TIME(SUM(TIME_TO_SEC(D.time_to) - TIME_TO_SEC(D.time_from)))) as 'totaltime_mins', " _
                                          & "SEC_TO_TIME(SUM(TIME_TO_SEC(CASE WHEN D.adjust_type = '+' THEN D.time_to ELSE 0 END) - TIME_TO_SEC(CASE WHEN D.adjust_type = '+' THEN D.time_from ELSE 0 END))) as 'worktime', " _
                                          & "HOUR(SEC_TO_TIME(SUM(TIME_TO_SEC(CASE WHEN D.adjust_type = '+' THEN D.time_to ELSE 0 END) - TIME_TO_SEC(CASE WHEN D.adjust_type = '+' THEN D.time_from ELSE 0 END)))) AS 'worktime_hours', " _
                                          & "MINUTE(SEC_TO_TIME(SUM(TIME_TO_SEC(CASE WHEN D.adjust_type = '+' THEN D.time_to ELSE 0 END) - TIME_TO_SEC(CASE WHEN D.adjust_type = '+' THEN D.time_from ELSE 0 END)))) AS 'worktime_mins', " _
                                          & "SEC_TO_TIME(SUM(TIME_TO_SEC(CASE WHEN D.adjust_type = '-' THEN D.time_to ELSE 0 END) - TIME_TO_SEC(CASE WHEN D.adjust_type = '-' THEN D.time_from ELSE 0 END))) as 'non_worktime', " _
                                          & "HOUR(SEC_TO_TIME(SUM(TIME_TO_SEC(CASE WHEN D.adjust_type = '-' THEN D.time_to ELSE 0 END) - TIME_TO_SEC(CASE WHEN D.adjust_type = '-' THEN D.time_from ELSE 0 END)))) AS 'non_worktime_hours', " _
                                          & "MINUTE(SEC_TO_TIME(SUM(TIME_TO_SEC(CASE WHEN D.adjust_type = '-' THEN D.time_to ELSE 0 END) - TIME_TO_SEC(CASE WHEN D.adjust_type = '-' THEN D.time_from ELSE 0 END)))) AS 'non_worktime_mins', " _
                                          & "T.remarks as 'remarks'" _
                                          & "FROM t_jobsheet_trn T, t_jobsheet_detl D " _
                                          & "WHERE(T.job_code = D.job_code) " _
                                          & "AND T.job_number = D.job_number " _
                                          & "AND T.username = '" & theEmp.ToString.Trim & "' " _
                                          & "AND T.job_date >= '" & dateFromString & "' " _
                                          & "AND T.job_date <= '" & dateToString & "' " _
                                          & "AND D.deleted = 0 " _
                                          & "GROUP BY T.job_date;"

            connect.Open()

            command1 = New OdbcCommand(columnHeaders, connect)

            myDataSet = New DataSet()
            myDataSet.Tables.Clear()
            myAdapter = New OdbcDataAdapter()
            myAdapter.SelectCommand = command1
            myAdapter.Fill(myDataSet, "headers")

            Dim name(myDataSet.Tables("headers").Columns.Count) As String

            Dim dt As New DataTable()
            Dim dr As DataRow

            Dim colCheckbox As CheckBox = New CheckBox()

            dt.Columns.Add("Select", GetType(Boolean))
            dt.Columns.Add("header")

            For i As Integer = 0 To myDataSet.Tables("headers").Columns.Count - 1
                dr = dt.NewRow()

                name(i) = myDataSet.Tables("headers").Columns.Item(i).ColumnName
                dr("header") = name(i)
                dt.Rows.Add(dr)
            Next

            dgvColumns.DataSource = dt
            dgvColumns.Columns(0).Width = 20
            dgvColumns.Columns(1).FillWeight = 100
            dgvColumns.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            dgvColumns.Refresh()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        loadFilters()
        loadSummaryColumn()

        cbDefault.Checked = False
        cbTickAll.Checked = False
        cbTickAll1.Checked = False

        cbTickAll.Enabled = True
        cbDefault.Enabled = True

    End Sub

    Private Sub rbDetailed_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbDetailed.CheckedChanged

        dgvColumns.DataSource = ""
        dgvColumns.Refresh()

        Try

            Dim theEmp As String = cmbEmployee.SelectedValue
            Dim dateFrom As DateTime = dtpFrom.Value
            Dim dateFromString As String = dateFrom.ToString("yyyy-MM-dd")
            Dim dateTo As DateTime = dtpTo.Value
            Dim dateToString As String = dateTo.ToString("yyyy-MM-dd")


            Dim command1 As OdbcCommand
            Dim myAdapter As OdbcDataAdapter
            Dim myDataSet As DataSet

            connect = New OdbcConnection(sqlConn)

            Dim columnHeaders As String = "SELECT T.job_date as 'job_date', " _
                                          & "DAYNAME(T.job_date) as 'day', " _
                                          & "CONCAT(D.job_code,'-',D.job_number) as 'job_ref', " _
                                          & "D.work_project_type as 'work_type', " _
                                          & "D.time_from as 'time_from', " _
                                          & "D.time_to as 'time_to', " _
                                          & "D.dbcode as 'customer', " _
                                          & "CONCAT(D.doc_doc,'-',D.doc_ref) as 'prj_ref', " _
                                          & "D.job_desc as 'description' " _
                                          & "FROM t_jobsheet_trn T, t_jobsheet_detl D " _
                                          & "WHERE T.job_code = D.job_code " _
                                          & "AND T.job_number = D.job_number " _
                                          & "AND T.username = '" & theEmp.ToString.Trim & "' " _
                                          & "AND T.job_date >= '" & dateFromString & "' " _
                                          & "AND T.job_date <= '" & dateToString & "' " _
                                          & "AND D.deleted = 0 " _
                                          & "ORDER BY T.job_date, D.time_from; "

            connect.Open()

            command1 = New OdbcCommand(columnHeaders, connect)

            myDataSet = New DataSet()
            myDataSet.Tables.Clear()
            myAdapter = New OdbcDataAdapter()
            myAdapter.SelectCommand = command1
            myAdapter.Fill(myDataSet, "headers")

            Dim name(myDataSet.Tables("headers").Columns.Count) As String

            Dim dt As New DataTable()
            Dim dr As DataRow

            Dim colCheckbox As CheckBox = New CheckBox()

            dt.Columns.Add("Select", GetType(Boolean))
            dt.Columns.Add("header")

            For i As Integer = 0 To myDataSet.Tables("headers").Columns.Count - 1
                dr = dt.NewRow()

                name(i) = myDataSet.Tables("headers").Columns.Item(i).ColumnName
                dr("header") = name(i)
                dt.Rows.Add(dr)
            Next

            dgvColumns.DataSource = dt
            dgvColumns.Columns(0).Width = 20
            dgvColumns.Columns(1).FillWeight = 100
            dgvColumns.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            dgvColumns.Refresh()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        loadFilters()
        loadDetailColumn()

        cbDefault.Checked = False
        cbTickAll.Checked = False
        cbTickAll1.Checked = True

        cbTickAll.Enabled = True
        cbDefault.Enabled = True

    End Sub

    Public Sub loadSummaryColumn()

        If rbSummary.Checked = True And dgvColumns.Rows.Count <> 0 Then

            Dim o As Integer

            Do While o < dgvColumns.Rows.Count
                If o = 0 Then
                    dgvColumns.Rows(o).Cells("Select").Value = True
                ElseIf o = 1 Then
                    dgvColumns.Rows(o).Cells("Select").Value = True
                ElseIf o = 4 Then
                    dgvColumns.Rows(o).Cells("Select").Value = True
                ElseIf o = 5 Then
                    dgvColumns.Rows(o).Cells("Select").Value = True
                ElseIf o = 8 Then
                    dgvColumns.Rows(o).Cells("Select").Value = True
                ElseIf o = 11 Then
                    dgvColumns.Rows(o).Cells("Select").Value = True
                ElseIf o = 14 Then
                    dgvColumns.Rows(o).Cells("Select").Value = True
                Else
                    dgvColumns.Rows(o).Cells("Select").Value = False
                End If
                o += 1
            Loop

            Dim j As Integer

            Do While j < dgvFilter.Rows.Count
                dgvFilter.Rows(j).Cells("Select").Value = False
                j += 1
            Loop

        Else
            'do nothing
        End If

    End Sub

    Public Sub loadDetailColumn()

        If rbDetailed.Checked = True And dgvColumns.Rows.Count <> 0 Then
            Dim i As Integer
            Do While i < dgvColumns.Rows.Count

                If i = 0 Then
                    dgvColumns.Rows(i).Cells("Select").Value = True
                ElseIf i = 1 Then
                    dgvColumns.Rows(i).Cells("Select").Value = True
                ElseIf i = 3 Then
                    dgvColumns.Rows(i).Cells("Select").Value = True
                ElseIf i = 4 Then
                    dgvColumns.Rows(i).Cells("Select").Value = True
                ElseIf i = 5 Then
                    dgvColumns.Rows(i).Cells("Select").Value = True
                ElseIf i = 6 Then
                    dgvColumns.Rows(i).Cells("Select").Value = True
                ElseIf i = 8 Then
                    dgvColumns.Rows(i).Cells("Select").Value = True
                End If
                i = i + 1

            Loop

            Dim j As Integer

            Do While j < dgvFilter.Rows.Count
                dgvFilter.Rows(j).Cells("Select").Value = True
                j += 1
            Loop

        Else
            'do nothing
        End If

    End Sub

    Public Sub loadSummaryReport()

        Dim theEmp As String = cmbEmployee.SelectedValue
        Dim dateFrom As DateTime = dtpFrom.Value
        Dim dateFromString As String = dateFrom.ToString("yyyy-MM-dd")
        Dim dateTo As DateTime = dtpTo.Value
        Dim dateToString As String = dateTo.ToString("yyyy-MM-dd")

        Try
            Dim i As Integer = 0

            Dim finalstring As String

            Dim List As String
            Dim commandString As StringBuilder = New StringBuilder

            Dim rmList As String
            Dim remarkString As StringBuilder = New StringBuilder

            Dim List1 As String
            Dim commandString1 As StringBuilder = New StringBuilder

            Do While i < dgvColumns.Rows.Count

                If dgvColumns.Rows(i).Cells("Select").Value.ToString = "True" Then

                    Dim selectedColumn As String = dgvColumns.Rows(i).Cells("header").FormattedValue.ToString

                    'Dim stringCommand1 As String = "T.job_date"
                    Dim stringCommand1 As String = "DATE_FORMAT(T.job_date,'%d/%m/%Y') as 'job_date'"
                    Dim stringCommand2 As String = "DAYNAME(T.job_date) as 'day'"
                    Dim stringCommand3 As String = "MIN(D.time_from) as 'first_task'"
                    Dim stringcommand3point5 As String = "MAX(D.time_to) as 'last_task'"
                    Dim stringCommand4 As String = "COUNT(D.job_number) as 'job_count'"
                    Dim stringCommand5 As String = "TIME_FORMAT(SEC_TO_TIME(SUM(TIME_TO_SEC(D.time_to) - TIME_TO_SEC(D.time_from))), '%H:%i') AS 'totaltime'"
                    Dim stringCommand6 As String = "HOUR(SEC_TO_TIME(SUM(TIME_TO_SEC(D.time_to) - TIME_TO_SEC(D.time_from)))) as 'totaltime_hours'"
                    Dim stringCommand7 As String = "MINUTE(SEC_TO_TIME(SUM(TIME_TO_SEC(D.time_to) - TIME_TO_SEC(D.time_from)))) as 'totaltime_mins'"
                    Dim stringCommand8 As String = "TIME_FORMAT(SEC_TO_TIME(SUM(TIME_TO_SEC(CASE WHEN D.adjust_type = '+' THEN D.time_to ELSE 0 END) - " _
                                                   & "TIME_TO_SEC(CASE WHEN D.adjust_type = '+' THEN D.time_from ELSE 0 END))), '%H:%i') as 'worktime'"
                    Dim stringCommand9 As String = "HOUR(SEC_TO_TIME(SUM(TIME_TO_SEC(CASE WHEN D.adjust_type = '+' THEN D.time_to ELSE 0 END) - " _
                                                   & "TIME_TO_SEC(CASE WHEN D.adjust_type = '+' THEN D.time_from ELSE 0 END)))) AS 'worktime_hours'"
                    Dim stringCommand10 As String = "MINUTE(SEC_TO_TIME(SUM(TIME_TO_SEC(CASE WHEN D.adjust_type = '+' THEN D.time_to ELSE 0 END) - " _
                                                    & "TIME_TO_SEC(CASE WHEN D.adjust_type = '+' THEN D.time_from ELSE 0 END)))) AS 'worktime_mins'"
                    Dim stringCommand11 As String = "TIME_FORMAT(SEC_TO_TIME(SUM(TIME_TO_SEC(CASE WHEN D.adjust_type = '-' THEN D.time_to ELSE 0 END) - " _
                                                    & "TIME_TO_SEC(CASE WHEN D.adjust_type = '-' THEN D.time_from ELSE 0 END))), '%H:%i') as 'non_worktime'"
                    Dim stringCommand12 As String = "HOUR(SEC_TO_TIME(SUM(TIME_TO_SEC(CASE WHEN D.adjust_type = '-' THEN D.time_to ELSE 0 END) - " _
                                                    & "TIME_TO_SEC(CASE WHEN D.adjust_type = '-' THEN D.time_from ELSE 0 END)))) AS 'non_worktime_hours'"
                    Dim stringCommand13 As String = "MINUTE(SEC_TO_TIME(SUM(TIME_TO_SEC(CASE WHEN D.adjust_type = '-' THEN D.time_to ELSE 0 END) - " _
                                                    & "TIME_TO_SEC(CASE WHEN D.adjust_type = '-' THEN D.time_from ELSE 0 END)))) AS 'non_worktime_mins'"
                    Dim stringCommand14 As String = "T.remarks as 'remarks'"

                    If selectedColumn = "job_date" Then
                        commandString.Append(stringCommand1 & ", ")
                    ElseIf selectedColumn = "day" Then
                        commandString.Append(stringCommand2 & ", ")
                    ElseIf selectedColumn = "first_task" Then
                        commandString.Append(stringCommand3 & ", ")
                    ElseIf selectedColumn = "last_task" Then
                        commandString.Append(stringcommand3point5 & ", ")
                    ElseIf selectedColumn = "job_count" Then
                        commandString.Append(stringCommand4 & ", ")
                    ElseIf selectedColumn = "totaltime" Then
                        commandString.Append(stringCommand5 & ", ")
                    ElseIf selectedColumn = "totaltime_hours" Then
                        commandString.Append(stringCommand6 & ", ")
                    ElseIf selectedColumn = "totaltime_mins" Then
                        commandString.Append(stringCommand7 & ", ")
                    ElseIf selectedColumn = "worktime" Then
                        commandString.Append(stringCommand8 & ", ")
                    ElseIf selectedColumn = "worktime_hours" Then
                        commandString.Append(stringCommand9 & ", ")
                    ElseIf selectedColumn = "worktime_mins" Then
                        commandString.Append(stringCommand10 & ", ")
                    ElseIf selectedColumn = "non_worktime" Then
                        commandString.Append(stringCommand11 & ", ")
                    ElseIf selectedColumn = "non_worktime_hours" Then
                        commandString.Append(stringCommand12 & ", ")
                    ElseIf selectedColumn = "non_worktime_mins" Then
                        commandString.Append(stringCommand13 & ", ")
                    ElseIf selectedColumn = "remarks" Then
                        remarkString.Append(stringCommand14)
                    End If

                Else
                    'do nothing
                End If
                i += 1
            Loop

            List = commandString.ToString
            rmList = remarkString.ToString

            Dim stringCommand15 As String = ""

            Dim j As Integer = 0
            Do While j < dgvFilter.Rows.Count
                If dgvFilter.Rows(j).Cells("Select").Value.ToString = "True" Then

                    Dim selectedFilter As String = dgvFilter.Rows(j).Cells("work_type").FormattedValue.ToString

                    stringCommand15 = "TIME_FORMAT(SEC_TO_TIME(SUM(TIME_TO_SEC(CASE WHEN D.work_project_type = '" & selectedFilter & "' THEN D.time_to ELSE 0 END) - " _
                                                & "TIME_TO_SEC(CASE WHEN D.work_project_type = '" & selectedFilter & "' THEN D.time_from ELSE 0 END))), '%H:%i') as '" & selectedFilter & "',"
                    commandString1.Append(stringCommand15)
                Else
                    'do nothing
                End If
                j += 1
            Loop

            List1 = commandString1.ToString

            If List = "" Then
                dgvReport.DataSource = Nothing
                dgvReport.Refresh()
                Exit Sub
            Else
                finalstring = List.Substring(0, commandString.Length - 2)
            End If

            If List1 = "" Then
                'do nothing
            Else
                finalstring = List.Substring(0, commandString.Length - 1) & List1.Substring(0, commandString1.Length - 1)
            End If

            If rmList = "" Then
                'do nothing
            Else
                finalstring = finalstring & "," & rmList.Substring(0, remarkString.Length)
            End If

            Dim command1 As OdbcCommand
            Dim myAdapter As OdbcDataAdapter
            'Dim SummaryDataSet as DataSet

            connect = New OdbcConnection(sqlConn)

            Dim summary As String = "SELECT " & finalstring & " FROM t_jobsheet_trn T, t_jobsheet_detl D " _
                              & "WHERE(T.job_code = D.job_code) " _
                              & "AND T.job_number = D.job_number " _
                              & "AND T.username = '" & theEmp.ToString.Trim & "' " _
                              & "AND T.job_date >= '" & dateFromString & "' " _
                              & "AND T.job_date <= '" & dateToString & "' " _
                              & "AND D.deleted = 0 " _
                              & "GROUP BY T.job_date;"

            connect.Open()

            command1 = New OdbcCommand(summary, connect)

            SummaryDataSet = New DataSet()
            SummaryDataSet.Tables.Clear()
            myAdapter = New OdbcDataAdapter()
            myAdapter.SelectCommand = command1
            myAdapter.Fill(SummaryDataSet, "report")

            If SummaryDataSet.Tables("report").Rows.Count <> 0 Then

                dgvReport.DataSource = SummaryDataSet.Tables("report")
                dgvReport.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

                Dim l As Integer = 0
                Do While l < dgvColumns.Rows.Count

                    If dgvColumns.Rows(l).Cells("header").Value = "remarks" Then
                        If dgvColumns.Rows(l).Cells("Select").Value = "True" Then

                            dgvReport.Columns("remarks").FillWeight = 200
                            dgvReport.Columns("remarks").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                            dgvReport.Columns("remarks").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        Else
                            'do nothing
                        End If
                        Exit Do
                    Else
                        l += 1
                    End If

                Loop


                'dgvReport.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgvReport.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
                dgvReport.AutoResizeColumns()
                dgvReport.Refresh()

            Else
                MessageBox.Show("There are no records found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                dgvReport.DataSource = Nothing
                dgvReport.Refresh()
            End If

            connect.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Public Sub loadDetailedReport()

        Dim theEmp As String = cmbEmployee.SelectedValue
        Dim dateFrom As DateTime = dtpFrom.Value
        Dim dateFromString As String = dateFrom.ToString("yyyy-MM-dd")
        Dim dateTo As DateTime = dtpTo.Value
        Dim dateToString As String = dateTo.ToString("yyyy-MM-dd")

        Try
            Dim i As Integer = 0
            Dim List As String
            Dim finalstring As String
            Dim commandString As StringBuilder = New StringBuilder

            Do While i < dgvColumns.Rows.Count

                If dgvColumns.Rows(i).Cells("Select").Value.ToString = "True" Then

                    Dim selectedColumn As String = dgvColumns.Rows(i).Cells("header").FormattedValue.ToString

                    'Dim stringCommand1 As String = "T.job_date as 'job_date'"
                    Dim stringCommand1 As String = "DATE_FORMAT(T.job_date,'%d/%m/%Y') as 'job_date'"
                    Dim stringCommand2 As String = "DAYNAME(T.job_date) as 'day'"
                    Dim stringCommand3 As String = "CONCAT(D.job_code,'-',D.job_number) as 'job_ref'"
                    Dim stringCommand5 As String = "D.work_project_type as 'work_type'"
                    Dim stringCommand6 As String = "TIME_FORMAT(D.time_from, '%H:%i') as 'time_from'"
                    Dim stringCommand7 As String = "TIME_FORMAT(D.time_to, '%H:%i') as 'time_to'"
                    Dim stringCommand8 As String = "D.dbcode as 'customer'"
                    Dim stringCommand9 As String = "CONCAT(D.doc_doc,'-',D.doc_ref) as 'prj_ref'"
                    Dim stringCommand12 As String = "D.job_desc as 'description'"

                    If selectedColumn = "job_date" Then
                        commandString.Append(stringCommand1 & ", ")
                    ElseIf selectedColumn = "day" Then
                        commandString.Append(stringCommand2 & ", ")
                    ElseIf selectedColumn = "job_ref" Then
                        commandString.Append(stringCommand3 & ", ")
                    ElseIf selectedColumn = "work_type" Then
                        commandString.Append(stringCommand5 & ", ")
                    ElseIf selectedColumn = "time_from" Then
                        commandString.Append(stringCommand6 & ", ")
                    ElseIf selectedColumn = "time_to" Then
                        commandString.Append(stringCommand7 & ", ")
                    ElseIf selectedColumn = "customer" Then
                        commandString.Append(stringCommand8 & ", ")
                    ElseIf selectedColumn = "prj_ref" Then
                        commandString.Append(stringCommand9 & ", ")
                    ElseIf selectedColumn = "description" Then
                        commandString.Append(stringCommand12 & ", ")
                    End If

                Else
                    'do nothing
                End If
                i += 1
            Loop

            List = commandString.ToString

            If List = "" Then
                dgvReport.DataSource = Nothing
                dgvReport.Refresh()
                Exit Sub
            Else
                finalstring = List.Substring(0, commandString.Length - 2)
            End If

            Dim List1 As String
            Dim finalstring1 As String
            Dim commandString1 As StringBuilder = New StringBuilder

            Dim j As Integer = 0
            Dim countSelect As Integer = 0
            Dim filterCommand As String = "AND D.work_project_type IN ("

            Do While j < dgvFilter.Rows.Count
                If dgvFilter.Rows(j).Cells("Select").Value.ToString = "True" Then

                    Dim selectedFilter As String = dgvFilter.Rows(j).Cells("work_type").FormattedValue.ToString
                    countSelect += 1

                    If countSelect > 1 Then
                        Dim strFilter As String = ",'" & selectedFilter.ToString.Trim & "'"
                        commandString1.Append(strFilter)
                    Else
                        Dim strFilter As String = "AND D.work_project_type IN ('" & selectedFilter.ToString.Trim & "'"
                        commandString1.Append(strFilter)
                    End If
                Else
                    filterCommand = ""
                End If
                j += 1
            Loop

            List1 = commandString1.ToString

            If List1 = "" Then
                finalstring1 = ""
            Else
                finalstring1 = List1.Substring(0, commandString1.Length) & ")"
            End If

            Dim command1 As OdbcCommand
            Dim myAdapter As OdbcDataAdapter
            Dim myDataSet As DataSet

            connect = New OdbcConnection(sqlConn)

            Dim detailed As String = "SELECT " & finalstring & " " _
                                    & "FROM t_jobsheet_trn T, t_jobsheet_detl D " _
                                    & "WHERE T.job_code = D.job_code " _
                                    & "AND T.job_number = D.job_number " _
                                    & "AND T.username = '" & theEmp.ToString.Trim & "' " _
                                    & "AND T.job_date >= '" & dateFromString & "' " _
                                    & "AND T.job_date <= '" & dateToString & "' " & finalstring1 & " " _
                                    & "AND D.deleted = 0 " _
                                    & "ORDER BY T.job_date, D.time_from;"

            connect.Open()

            command1 = New OdbcCommand(detailed, connect)

            myDataSet = New DataSet()
            myDataSet.Tables.Clear()
            myAdapter = New OdbcDataAdapter()
            myAdapter.SelectCommand = command1
            myAdapter.Fill(myDataSet, "report")

            If myDataSet.Tables("report").Rows.Count <> 0 Then

                dgvReport.DataSource = myDataSet.Tables("report")
                dgvReport.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

                Dim l As Integer = 0
                Do While l < dgvColumns.Rows.Count

                    If dgvColumns.Rows(l).Cells("header").Value = "job_date" Then
                        If dgvColumns.Rows(l).Cells("Select").Value.ToString = "True" Then
                            dgvReport.Columns("job_date").Width = 70
                            l += 1
                        Else
                            'do nothing
                            l += 1
                        End If
                    End If

                    If dgvColumns.Rows(l).Cells("header").Value = "day" Then
                        If dgvColumns.Rows(l).Cells("Select").Value.ToString = "True" Then
                            dgvReport.Columns("day").Width = 80
                            l += 1
                        Else
                            'do nothing
                            l += 1
                        End If
                    End If

                    If dgvColumns.Rows(l).Cells("header").Value = "job_ref" Then
                        If dgvColumns.Rows(l).Cells("Select").Value.ToString = "True" Then
                            dgvReport.Columns("job_ref").Width = 70
                            l += 1
                        Else
                            l += 1
                        End If
                    End If

                    If dgvColumns.Rows(l).Cells("header").Value = "work_type" Then
                        If dgvColumns.Rows(l).Cells("Select").Value.ToString = "True" Then
                            dgvReport.Columns("work_type").Width = 70
                            l += 1
                        Else
                            'do nothing
                            l += 1
                        End If
                    End If

                    If dgvColumns.Rows(l).Cells("header").Value = "time_from" Then
                        If dgvColumns.Rows(l).Cells("Select").Value.ToString = "True" Then
                            dgvReport.Columns("time_from").Width = 70
                            l += 1
                        Else
                            'do nothing
                            l += 1
                        End If
                    End If

                    If dgvColumns.Rows(l).Cells("header").Value = "time_to" Then
                        If dgvColumns.Rows(l).Cells("Select").Value.ToString = "True" Then
                            dgvReport.Columns("time_to").Width = 70
                            l += 1
                        Else
                            'do nothing
                            l += 1
                        End If
                    End If

                    If dgvColumns.Rows(l).Cells("header").Value = "customer" Then
                        If dgvColumns.Rows(l).Cells("Select").Value.ToString = "True" Then
                            dgvReport.Columns("customer").Width = 70
                            l += 1
                        Else
                            'do nothing
                            l += 1
                        End If
                    End If

                    If dgvColumns.Rows(l).Cells("header").Value = "prj_ref" Then
                        If dgvColumns.Rows(l).Cells("Select").Value.ToString = "True" Then
                            dgvReport.Columns("prj_ref").Width = 70
                            l += 1
                        Else
                            'do nothing
                            l += 1
                        End If
                    End If

                    If dgvColumns.Rows(l).Cells("header").Value = "description" Then
                        If dgvColumns.Rows(l).Cells("Select").Value.ToString = "True" Then

                            'use this if export
                            dgvReport.Columns("description").FillWeight = 1000
                            dgvReport.Columns("description").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                            'end

                            'if printing. use width.
                            'dgvReport.Columns("description").Width = 550
                            'end
                            dgvReport.Columns("description").DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft
                            dgvReport.Columns("description").DefaultCellStyle.WrapMode = DataGridViewTriState.True
                            l += 1
                        Else
                            'do nothing
                            l += 1
                        End If
                    End If
                    l += 1
                Loop

                dgvReport.DefaultCellStyle.WrapMode = DataGridViewTriState.True
                dgvReport.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True
                dgvReport.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                dgvReport.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgvReport.Refresh()

            Else
                MessageBox.Show("There are no records found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
                dgvReport.DataSource = Nothing
                dgvReport.Refresh()
            End If

            connect.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub cbTickAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbTickAll.CheckedChanged

        If cbTickAll.Checked = True Then
            Dim i As Integer
            Do While i < dgvColumns.Rows.Count
                dgvColumns.Rows(i).Cells("Select").Value = True
                i = i + 1
            Loop
        Else
            Dim i As Integer
            Do While i < dgvColumns.Rows.Count
                dgvColumns.Rows(i).Cells("Select").Value = False
                i = i + 1
            Loop
        End If

    End Sub

    Private Sub cbDefault_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbDefault.CheckedChanged

        If cbTickAll.Checked = True Then
            cbTickAll.Checked = False
            If rbSummary.Checked = True Then
                loadSummaryColumn()
                cbDefault.Checked = False
            Else
                loadDetailColumn()
                cbDefault.Checked = False
            End If
        Else
            If rbSummary.Checked = True Then
                loadSummaryColumn()
                cbDefault.Checked = False
            Else
                loadDetailColumn()
                cbDefault.Checked = False
            End If
        End If


    End Sub

    Private Sub chkCustomer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCustomer.CheckedChanged

        If chkCustomer.Checked = False Then
            cmbCustomer.DataSource = ""
            cmbCustomer.Refresh()
            cmbCustomer.Enabled = False
        Else

            Dim command As OdbcCommand
            Dim myAdapter As OdbcDataAdapter
            Dim myDataSet As DataSet

            connect = New OdbcConnection(sqlConn)
            connect.Open()

            Try
                Dim str_Customers_Show As String = "SELECT dbcode, db_name1, concat(dbcode, ' - ', db_name1) as full FROM m_armaster where dbcode <> '-' and deleted = 0 ORDER BY dbcode;"

                command = New OdbcCommand(str_Customers_Show, connect)

                myDataSet = New DataSet()
                myDataSet.Tables.Clear()
                myAdapter = New OdbcDataAdapter()
                myAdapter.SelectCommand = command
                myAdapter.Fill(myDataSet, "Db")

                cmbCustomer.DataSource = myDataSet.Tables("Db")
                cmbCustomer.DisplayMember = "full"
                cmbCustomer.ValueMember = "dbcode"
                If cmbCustomer.Enabled = True Then
                    cmbCustomer.SelectedIndex = -1
                Else
                    'NOTHING
                End If

                cmbCustomer.Refresh()

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

            connect.Close()

            cmbCustomer.Enabled = True
        End If

    End Sub

    Private Sub cbTickAll1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbTickAll1.CheckedChanged

        If cbTickAll1.Checked = True Then
            cbPlus.Checked = False
            cbMinus.Checked = False

            cbPlus.Enabled = True
            cbMinus.Enabled = True

            Dim i As Integer
            Do While i < dgvFilter.Rows.Count
                dgvFilter.Rows(i).Cells("Select").Value = True
                i = i + 1
            Loop
        Else
            Dim i As Integer
            Do While i < dgvFilter.Rows.Count
                dgvFilter.Rows(i).Cells("Select").Value = False
                i = i + 1
            Loop
        End If

    End Sub

    Private Sub btnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport.Click

        Dim dir As String
        dir = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

        Dim type As String

        If rbSummary.Checked = True Then
            type = "summary"
        ElseIf rbDetailed.Checked = True Then
            type = "detail"
        Else
            type = ""
        End If

        If dgvReport.Rows.Count <> 0 Then
            Dim SaveFileDialog As New SaveFileDialog

            With SaveFileDialog
                .FileName = cmbEmployee.Text.ToString.ToLower & "_" & dtpFrom.Value.ToString("dd") & "_" & dtpTo.Value.ToString("dd") & "_" & type.ToString
                .DefaultExt = "*.xml"
                .Filter = "XML Files|*.xml"
                .InitialDirectory = dir
                .OverwritePrompt = True
                If .ShowDialog = DialogResult.OK Then
                    'MessageBox.Show(.FileName)
                    expPath = (.FileName)
                    'Me.WindowState = FormWindowState.Minimized
                    'frmMain.WindowState = FormWindowState.Minimized
                    DATAGRIDVIEW_TO_EXCEL((dgvReport)) ' PARAMETER: YOUR DATAGRIDVIEW
                End If
            End With

        Else
            MessageBox.Show("Please generate data first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub

    Public Sub DATAGRIDVIEW_TO_EXCEL(ByVal DGV As DataGridView)

        Try
            Dim DTB = New DataTable, RWS As Integer, CLS As Integer

            For CLS = 0 To DGV.ColumnCount - 1 ' COLUMNS OF DTB
                DTB.Columns.Add(DGV.Columns(CLS).Name.ToString)
            Next

            Dim DRW As DataRow

            For RWS = 0 To DGV.Rows.Count - 1 ' FILL DTB WITH DATAGRIDVIEW
                DRW = DTB.NewRow

                For CLS = 0 To DGV.ColumnCount - 1
                    Try
                        DRW(DTB.Columns(CLS).ColumnName.ToString) = DGV.Rows(RWS).Cells(CLS).Value.ToString
                    Catch ex As Exception

                    End Try
                Next

                DTB.Rows.Add(DRW)
            Next

            DTB.AcceptChanges()

            Dim DST As New DataSet
            DST.Tables.Add(DTB)
            Dim FLE As String = expPath.ToString ' PATH AND FILE NAME WHERE THE XML WIL BE CREATED (EXEMPLE: C:\REPS\XML.xml)
            DTB.WriteXml(FLE)

            Dim xl As New Excel.Application

            Dim wb = xl.Workbooks.OpenXML(Filename:=FLE, LoadOption:=Excel.XlXmlLoadOption.xlXmlLoadImportToList)
            xl.Visible = True


        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub cbPlus_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbPlus.CheckedChanged

        If cbPlus.Checked = True Then
            cbTickAll1.Checked = False
            cbMinus.Checked = False
            cbMinus.Enabled = True
            Dim i As Integer
            Do While i < dgvFilter.Rows.Count

                If dgvFilter.Rows(i).Cells("sign").Value = "+" Then
                    dgvFilter.Rows(i).Cells("Select").Value = True
                    i += 1
                Else
                    dgvFilter.Rows(i).Cells("Select").Value = False
                    i += 1
                End If
            Loop

            cbPlus.Enabled = False

            'Else
            '    Dim i As Integer
            '    Do While i < dgvFilter.Rows.Count

            '        If dgvFilter.Rows(i).Cells("sign").Value = "+" Then
            '            dgvFilter.Rows(i).Cells("Select").Value = False
            '            i += 1
            '        Else
            '            i += 1
            '        End If
            '    Loop
        End If



    End Sub

    Private Sub cbMinus_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbMinus.CheckedChanged

        If cbMinus.Checked = True Then
            cbTickAll1.Checked = False
            cbPlus.Checked = False
            cbPlus.Enabled = True
            Dim i As Integer
            Do While i < dgvFilter.Rows.Count

                If dgvFilter.Rows(i).Cells("sign").Value = "-" Then
                    dgvFilter.Rows(i).Cells("Select").Value = True
                    i += 1
                Else
                    dgvFilter.Rows(i).Cells("Select").Value = False
                    i += 1
                End If
            Loop

            cbMinus.Enabled = False
            'Else
            '    Dim i As Integer
            '    Do While i < dgvFilter.Rows.Count

            '        If dgvFilter.Rows(i).Cells("sign").Value = "-" Then
            '            dgvFilter.Rows(i).Cells("Select").Value = False
            '            i += 1
            '        Else
            '            i += 1
            '        End If
            '    Loop
        End If

    End Sub

    Private Sub dgvFilter_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvFilter.CellContentClick

        If e.ColumnIndex = 0 Then
            cbPlus.Checked = False
            cbPlus.Enabled = True

            cbMinus.Checked = False
            cbMinus.Enabled = True

            cbTickAll1.Checked = False
        End If

    End Sub
End Class