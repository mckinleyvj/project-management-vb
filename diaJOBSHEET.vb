Imports System.IO
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.Sql
Imports System.Xml
Imports System.Data.SqlClient
'Imports System.Windows.Forms.PrintPreviewDialog
Imports System.Xml.Xsl

Public Class diaJOBSHEET

    Dim sqlConn As String = frmConfig.theConnection
    Dim connect As OdbcConnection

    Dim theLoginUser As String = frmMain.theUser.ToString.Trim()

    Dim headerFont As New Font("Segoe UI", 9, FontStyle.Bold)
    Dim detailFont As New Font("Segoe UI", 8, FontStyle.Regular)

    Public myXML As String = Application.StartupPath & "\prnJobsheet.XML"

    Private finalDataSet As New DataSet

    'total JS time
    Public totalJSTimeHr As String
    Public totalJSTimeMin As String

    'total W time
    Public totalWHr As String
    Public totalWMin As String

    'total P time
    Public totalPHr As String
    Public totalPMin As String

    'total J Time
    Public totalJHr As String
    Public totalJMin As String

    'total D Time
    Public totalDHr As String
    Public totalDMin As String

    Public totalC As String

    Public Sub getInfoData()
        Try
            connect = New OdbcConnection(sqlConn)
            connect.Open()

            'GET whole detl
            Dim command1 As OdbcCommand
            Dim myAdapter1 As OdbcDataAdapter
            Dim myDataSet1 As DataSet

            Dim str_getGrid As String = "SELECT J.job_code, J.job_number, J.doc_doc, J.doc_ref, J.time_from, J.time_to, J.dbcode, J.work_project_type " _
                                        & "FROM t_jobsheet_detl j where job_code = '" & Label12.Text.ToString.Trim.ToUpper & "' and job_number = '" & Label13.Text.ToString.Trim.ToUpper & "'" _
                                        & "AND deleted = 0;"

            command1 = New OdbcCommand(str_getGrid, connect)

            myDataSet1 = New DataSet()
            myDataSet1.Tables.Clear()
            myAdapter1 = New OdbcDataAdapter()
            myAdapter1.SelectCommand = command1
            myAdapter1.Fill(myDataSet1, "str_getGrid")
            'END

            '1. GET total hour and min of jobsheet
            Dim command2 As OdbcCommand
            Dim myAdapter2 As OdbcDataAdapter
            Dim myDataSet2 As DataSet

            Dim str_getTotalTime As String = "select HOUR(SEC_TO_TIME(SUM(TIME_TO_SEC(time_to) - TIME_TO_SEC(time_from)))) AS totalhours, MINUTE(SEC_TO_TIME(SUM(TIME_TO_SEC(time_to) - TIME_TO_SEC(time_from)))) as totalmins " _
                                             & "FROM t_jobsheet_detl j where job_code = '" & Label12.Text.ToString.Trim.ToUpper & "' and job_number = '" & Label13.Text.ToString.Trim.ToUpper & "'" _
                                             & "AND deleted = 0;"

            command2 = New OdbcCommand(str_getTotalTime, connect)

            myDataSet2 = New DataSet()
            myDataSet2.Tables.Clear()
            myAdapter2 = New OdbcDataAdapter()
            myAdapter2.SelectCommand = command2
            myAdapter2.Fill(myDataSet2, "str_getTotalTime")

            totalJSTimeHr = myDataSet2.Tables("str_getTotalTime").Rows(0)(0).ToString
            totalJSTimeMin = myDataSet2.Tables("str_getTotalTime").Rows(0)(1).ToString
            'end

            'GET total Jobs hour and min
            Dim command3 As OdbcCommand
            Dim myAdapter3 As OdbcDataAdapter
            Dim myDataSet3 As DataSet

            Dim str_getTotalNonProjectTime As String = "select HOUR(SEC_TO_TIME(SUM(TIME_TO_SEC(time_to) - TIME_TO_SEC(time_from)))) AS totalhours, MINUTE(SEC_TO_TIME(SUM(TIME_TO_SEC(time_to) - TIME_TO_SEC(time_from)))) as totalmins " _
                                                       & "FROM t_jobsheet_detl j where job_code = '" & Label12.Text.ToString.Trim.ToUpper & "' and job_number = '" & Label13.Text.ToString.Trim.ToUpper & "' and doc_doc = ' ' " _
                                                       & "AND deleted = 0;"

            command3 = New OdbcCommand(str_getTotalNonProjectTime, connect)

            myDataSet3 = New DataSet()
            myDataSet3.Tables.Clear()
            myAdapter3 = New OdbcDataAdapter()
            myAdapter3.SelectCommand = command3
            myAdapter3.Fill(myDataSet3, "str_getTotalNonProjectTime")

            totalJHr = myDataSet3.Tables("str_getTotalNonProjectTime").Rows(0)(0).ToString
            totalJMin = myDataSet3.Tables("str_getTotalNonProjectTime").Rows(0)(1).ToString
            'end

            'get total project hour and min
            Dim command4 As OdbcCommand
            Dim myAdapter4 As OdbcDataAdapter
            Dim myDataSet4 As DataSet

            Dim str_getTotalProjectTime As String = "select HOUR(SEC_TO_TIME(SUM(TIME_TO_SEC(time_to) - TIME_TO_SEC(time_from)))) AS totalhours, MINUTE(SEC_TO_TIME(SUM(TIME_TO_SEC(time_to) - TIME_TO_SEC(time_from)))) as totalmins " _
                                                    & "FROM t_jobsheet_detl j where job_code = '" & Label12.Text.ToString.Trim.ToUpper & "' and job_number = '" & Label13.Text.ToString.Trim.ToUpper & "' and doc_doc <> ' ' " _
                                                    & "AND deleted = 0;"

            command4 = New OdbcCommand(str_getTotalProjectTime, connect)

            myDataSet4 = New DataSet()
            myDataSet4.Tables.Clear()
            myAdapter4 = New OdbcDataAdapter()
            myAdapter4.SelectCommand = command4
            myAdapter4.Fill(myDataSet4, "str_getTotalProjectTime")

            totalPHr = myDataSet4.Tables("str_getTotalProjectTime").Rows(0)(0).ToString
            totalPHr = myDataSet4.Tables("str_getTotalProjectTime").Rows(0)(1).ToString
            'end

            'get total deducted time
            Dim command5 As OdbcCommand
            Dim myAdapter5 As OdbcDataAdapter
            Dim myDataSet5 As DataSet

            Dim str_getTotalDeductTime As String = "select HOUR(SEC_TO_TIME(SUM(TIME_TO_SEC(time_to) - TIME_TO_SEC(time_from)))) AS totalhours, MINUTE(SEC_TO_TIME(SUM(TIME_TO_SEC(time_to) - TIME_TO_SEC(time_from)))) as totalmins " _
                                                  & "from(t_jobsheet_detl) where job_code = '" & Label12.Text.ToString.Trim.ToUpper & "' and job_number = '" & Label13.Text.ToString.Trim.ToUpper & "' and adjust_type = '-' " _
                                                  & "AND deleted = 0;"

            command5 = New OdbcCommand(str_getTotalDeductTime, connect)

            myDataSet5 = New DataSet()
            myDataSet5.Tables.Clear()
            myAdapter5 = New OdbcDataAdapter()
            myAdapter5.SelectCommand = command5
            myAdapter5.Fill(myDataSet5, "str_getTotalDeductTime")

            totalDHr = myDataSet5.Tables("str_getTotalDeductTime").Rows(0)(0).ToString
            totalDHr = myDataSet5.Tables("str_getTotalDeductTime").Rows(0)(1).ToString
            'end

            'GET count jobs without deleted
            Dim command6 As OdbcCommand
            Dim myAdapter6 As OdbcDataAdapter
            Dim myDataSet6 As DataSet

            Dim str_getCount As String = "SELECT count(J.job_code) FROM t_jobsheet_detl J " _
                                         & "where job_code = '" & Label12.Text.ToString.Trim.ToUpper & "' and job_number = '" & Label13.Text.ToString.Trim.ToUpper & "' " _
                                         & "and deleted = 0;"

            command6 = New OdbcCommand(str_getCount, connect)

            myDataSet6 = New DataSet()
            myDataSet6.Tables.Clear()
            myAdapter6 = New OdbcDataAdapter()
            myAdapter6.SelectCommand = command6
            myAdapter6.Fill(myDataSet6, "str_getCount")

            totalC = myDataSet6.Tables("str_getCount").Rows(0)(0).ToString
            'end

            connect.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub PRINTToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PRINTToolStripMenuItem.Click

        Try
            Dim jsDoc As String = Label12.Text.ToString.Trim
            Dim jsRef As String = Label13.Text.ToString.Trim
            'Dim theEmp As String = txtName.Text.ToString.Trim
            'Dim theDate As String = lblJobDate.Text.ToString.Trim
            'Dim theRemark As String = txtRemarks1.Text.ToString.Trim

            prnJOBSHT.lblJobCode.Text = jsDoc
            prnJOBSHT.lblJobNumber.Text = jsRef
            'prnJOBSHT.lblEmployee.Text = theEmp
            'prnJOBSHT.lblJobDate.Text = Mid(theDate, 12)
            'prnJOBSHT.lblRemarks.Text = theRemark
            'prnJOBSHT.lblRowCount.Text = ""
            'prnJOBSHT.lblTotalHours.Text = Convert.ToDouble(totalJSTimeHr).ToString("00") + " Hour(s)"
            'prnJOBSHT.lblTotalMins.Text = Convert.ToDouble(totalJSTimeMin).ToString("00") + " Minute(s)"
            'prnJOBSHT.lblWorkHr.Text = Convert.ToDouble(totalWHr).ToString("00") + " Hour(s)"
            'prnJOBSHT.lblWorkMin.Text = Convert.ToDouble(totalWMin).ToString("00") + " Minute(s)"

            'Dim nonworkhr As String = Convert.ToDouble(totalJSTimeHr) - Convert.ToDouble(totalWHr)
            'Dim nonworkmin As String = Convert.ToDouble(totalJSTimeMin) - Convert.ToDouble(totalWMin)

            'If nonworkmin > 60 Then
            '    nonworkhr = nonworkhr + 1
            '    nonworkmin = nonworkmin - 60
            'ElseIf nonworkmin < 0 Then
            '    nonworkhr = nonworkhr - 1
            '    nonworkmin = nonworkmin + 60
            'Else
            '    nonworkhr = nonworkhr
            '    nonworkmin = nonworkmin
            'End If

            'prnJOBSHT.lblNonWorkHours.Text = Convert.ToDouble(nonworkhr).ToString("00") + " Hour(s)"
            'prnJOBSHT.lblNonWorkMins.Text = Convert.ToDouble(nonworkmin).ToString("00") + " Minute(s)"

            prnJOBSHT.StartPosition = FormStartPosition.CenterScreen
            prnJOBSHT.Show()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub diaJOBSHEET_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.BackColor = Color.White
        Me.ForeColor = Color.Black
        MenuStrip1.BackColor = Color.Gainsboro
        MenuStrip1.ForeColor = Color.Black

        dtpDate.Format = DateTimePickerFormat.Custom

        dtpDate.CustomFormat = "dd/MM/yyyy"
        dtpDate.DropDownAlign = LeftRightAlignment.Right

        txtDate.Text = dtpDate.Value.ToString("dddd, dd-MMMM-yyyy")

        txtDateNow.Text = Format(Now, "dd/MM/yyyy")
        txtTimeNow.Text = Format(Now, "HH:mm:ss")

        Dim theLoginUser As String = frmMain.theUser.ToString.Trim()

        connect = New OdbcConnection(sqlConn)

        If Label12.Text <> "" And Label13.Text <> "" Then
            txtDocRef.Text = Label13.Text.ToString.Trim

            Me.Text = "Job Sheet Entry Screen - " & Label12.Text.ToString.Trim.ToUpper & "-" & Label13.Text.ToString.Trim.ToUpper

            load_info()
            refresh_grid()
            load_details()

            getInfoData()

            Gridview1.Select()
        ElseIf Label12.Text = "" And Label13.Text = "" Then
            'JUST LOAD FOR NEW JOB SHEET
            Me.Text = "Job Sheet Entry Screen"

            Try
                Dim command1 As OdbcCommand
                Dim myAdapter As OdbcDataAdapter
                Dim myDataSet As DataSet

                Dim str_Employee_Show As String = "SELECT username, emp_name1 FROM m_employee where username = '" & theLoginUser.ToString.Trim.ToUpper & "';"

                connect.Open()
                command1 = New OdbcCommand(str_Employee_Show, connect)

                myDataSet = New DataSet()
                myDataSet.Tables.Clear()
                myAdapter = New OdbcDataAdapter()
                myAdapter.SelectCommand = command1
                myAdapter.Fill(myDataSet, "Db")

                txtUser.Text = myDataSet.Tables("Db").Rows(0)(0).ToString
                txtName.Text = myDataSet.Tables("Db").Rows(0)(1).ToString

                connect.Close()
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
            'END OF LOAD USER DETAIL AT LOAD

        Else
            'do nothing
        End If

        'lblFocusedDate.Text = dtpDate.Value.ToString("yyyy-MM-dd")
        connect.Close()

    End Sub

    Private Sub diaJOBSHEET_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        If e.KeyCode = Keys.Escape Then

            If EditJobsheetToolStripMenuItem.Visible = True Then
                Me.Close()
            Else
                'do nothing
                If txtJobDesc.ReadOnly = False Then
                    Gridview1.Enabled = True

                    btnAddJob.Enabled = True
                    btnAddProj.Enabled = True
                    btnAddJob.Focus()

                    btnDetlCancel.Visible = False
                    btnDelJob.Visible = False
                    btnDetlEdit.Visible = False
                    btnDetlSave.Visible = False
                    btnUpdateCancel.Visible = False
                    btnProjUpdate.Visible = False
                    txtJobDesc.ReadOnly = True
                    dtpFrom.Enabled = False
                    dtpTo.Enabled = False

                    load_details()

                    If Label12.Text = "" Then
                        btnDelJob.Visible = False
                        btnDetlEdit.Visible = False
                    Else
                        btnDelJob.Visible = True
                        btnDetlEdit.Visible = True
                    End If
                Else
                    'Me.Close()
                    'do nothing
                End If
            End If

        End If
    End Sub

    Public Sub refresh_grid()
        Dim command1 As OdbcCommand
        Dim myAdapter As OdbcDataAdapter
        Dim myDataSet As DataSet

        connect = New OdbcConnection(sqlConn)

        Try

            Dim str_Detail_Show As String = "SELECT J.deleted as Deleted, J.closed as Closed, M.dbcode as dbcode, M.db_name1 as name, J.doc_doc as DoC, J.doc_ref as ReF," _
                                            & "concat(J.doc_doc, '-', J.doc_ref) as refNo, J.job_desc as jobDesc, J.time_from as timefr, J.time_to as timeto," _
                                            & "J.create_date as CDate, J.edit_date as EDate, J.ID as theID, J.work_project_type as workCode " _
                                            & "FROM t_jobsheet_detl J, t_jobsheet_trn T, m_armaster M " _
                                            & "where T.job_code = J.job_code and T.job_number = J.job_number and J.dbcode = M.dbcode and " _
                                            & "J.job_code = 'JB' and J.job_number = '" & txtDocRef.Text.ToString.Trim & "' and J.deleted = 0 order by J.time_from;"

            connect.Open()
            command1 = New OdbcCommand(str_Detail_Show, connect)

            myDataSet = New DataSet()
            myDataSet.Tables.Clear()
            myAdapter = New OdbcDataAdapter()
            myAdapter.SelectCommand = command1
            myAdapter.Fill(myDataSet, "detl")

            Dim dtRetrievedData1 As DataTable = myDataSet.Tables("detl")

            Dim dtData As New DataTable
            Dim dtDataRows As DataRow

            Dim chkboxcol As New DataGridViewCheckBoxColumn

            dtData.TableName = "ProjTable"
            'dtData.Columns.Add("X", GetType(Boolean))
            dtData.Columns.Add("C", GetType(Boolean))
            dtData.Columns.Add("Cust.")
            dtData.Columns.Add("Name")
            dtData.Columns.Add("Ref No.")
            dtData.Columns.Add("Job Desc")
            dtData.Columns.Add("Create")
            dtData.Columns.Add("Edit")
            dtData.Columns.Add("DoC")
            dtData.Columns.Add("ReF")
            dtData.Columns.Add("From")
            dtData.Columns.Add("To")
            dtData.Columns.Add("theID")
            dtData.Columns.Add("workCode")

            For Each dtDataRows In dtRetrievedData1.Rows

                'Dim del As Boolean = dtDataRows("Deleted").ToString().Trim()
                Dim close As Boolean = dtDataRows("Closed").ToString().Trim()
                Dim db As String = dtDataRows("dbcode").ToString().Trim()
                Dim dbname As String = dtDataRows("name").ToString().Trim()
                Dim ref As String = dtDataRows("refNo").ToString.Trim()
                Dim jobDesc As String = dtDataRows("jobDesc").ToString.Trim()
                Dim C_Date As Date = dtDataRows("CDate").ToString.Trim()
                Dim EDate As Date = dtDataRows("EDate").ToString.Trim()
                Dim doc_doc As String = dtDataRows("DoC").ToString.Trim()
                Dim doc_ref As String = dtDataRows("ReF").ToString.Trim()
                Dim theID As String = dtDataRows("theID").ToString.Trim()
                Dim theWork As String = dtDataRows("workCode").ToString.Trim()

                Dim timeFrom As String = Convert.ToDateTime((dtDataRows("timefr").ToString.Trim())).ToString("HH:mm")
                Dim timeTo As String = Convert.ToDateTime((dtDataRows("timeto").ToString.Trim())).ToString("HH:mm")

                'dtData.Rows.Add(New Object() {del.ToString.Trim(), close.ToString.Trim(), db.ToString.Trim(), dbname.ToString.Trim(), ref.ToString.Trim(), jobDesc.ToString.Trim, C_Date.ToString, EDate.ToString, doc_doc.ToString.Trim, doc_ref.ToString.Trim, timeFrom.ToString.Trim(), timeTo.ToString.Trim(), theID.ToString, theWork.ToString})
                dtData.Rows.Add(New Object() {close.ToString.Trim(), db.ToString.Trim(), dbname.ToString.Trim(), ref.ToString.Trim(), jobDesc.ToString.Trim, C_Date.ToString, EDate.ToString, doc_doc.ToString.Trim, doc_ref.ToString.Trim, timeFrom.ToString.Trim(), timeTo.ToString.Trim(), theID.ToString, theWork.ToString})
            Next

            Dim finalDataSet As New DataSet
            finalDataSet.Tables.Add(dtData)

            Gridview1.DataSource = finalDataSet.Tables(0)

            'Gridview1.Columns.Item("X").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            'Gridview1.Columns.Item("X").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            'Gridview1.Columns.Item("X").HeaderCell.Style.Font = headerFont
            'Gridview1.Columns.Item("X").DefaultCellStyle.Font = detailFont
            'Gridview1.Columns.Item("X").Resizable = DataGridViewTriState.False
            'Gridview1.Columns.Item("X").Width = 19
            'Gridview1.Columns.Item("X").Visible = False

            Gridview1.Columns.Item("C").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            Gridview1.Columns.Item("C").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            Gridview1.Columns.Item("C").HeaderCell.Style.Font = headerFont
            Gridview1.Columns.Item("C").DefaultCellStyle.Font = detailFont
            Gridview1.Columns.Item("C").Resizable = DataGridViewTriState.False
            Gridview1.Columns.Item("C").Width = 19

            Gridview1.Columns.Item("Cust.").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            Gridview1.Columns.Item("Cust.").HeaderCell.Style.Font = headerFont
            Gridview1.Columns.Item("Cust.").DefaultCellStyle.Font = detailFont
            Gridview1.Columns.Item("Cust.").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            Gridview1.Columns.Item("Cust.").Width = 50

            Gridview1.Columns.Item("Name").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            Gridview1.Columns.Item("Name").HeaderCell.Style.Font = headerFont
            Gridview1.Columns.Item("Name").DefaultCellStyle.Font = detailFont
            Gridview1.Columns.Item("Name").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            Gridview1.Columns.Item("Name").Width = 200

            Gridview1.Columns.Item("Ref No.").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            Gridview1.Columns.Item("Ref No.").HeaderCell.Style.Font = headerFont
            Gridview1.Columns.Item("Ref No.").DefaultCellStyle.Font = detailFont
            Gridview1.Columns.Item("Ref No.").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            Gridview1.Columns.Item("Ref No.").Width = 80

            Gridview1.Columns.Item("Job Desc").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            Gridview1.Columns.Item("Job Desc").HeaderCell.Style.Font = headerFont
            Gridview1.Columns.Item("Job Desc").DefaultCellStyle.Font = detailFont
            Gridview1.Columns.Item("Job Desc").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            Gridview1.Columns.Item("Job Desc").FillWeight = 500
            Gridview1.Columns.Item("Job Desc").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            Gridview1.Columns.Item("Job Desc").Visible = True

            Gridview1.Columns.Item("Create").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            Gridview1.Columns.Item("Create").HeaderCell.Style.Font = headerFont
            Gridview1.Columns.Item("Create").DefaultCellStyle.Font = detailFont
            Gridview1.Columns.Item("Create").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            Gridview1.Columns.Item("Create").Width = 95
            Gridview1.Columns.Item("Create").Visible = False

            Gridview1.Columns.Item("Edit").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            Gridview1.Columns.Item("Edit").HeaderCell.Style.Font = headerFont
            Gridview1.Columns.Item("Edit").DefaultCellStyle.Font = detailFont
            Gridview1.Columns.Item("Edit").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            Gridview1.Columns.Item("Edit").Width = 95
            Gridview1.Columns.Item("Edit").Visible = False

            Gridview1.Columns.Item("DoC").Visible = False
            Gridview1.Columns.Item("ReF").Visible = False

            Gridview1.Columns.Item("From").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            Gridview1.Columns.Item("From").HeaderCell.Style.Font = headerFont
            Gridview1.Columns.Item("From").DefaultCellStyle.Font = detailFont
            Gridview1.Columns.Item("From").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Gridview1.Columns.Item("From").Resizable = DataGridViewTriState.False
            Gridview1.Columns.Item("From").Width = 50

            Gridview1.Columns.Item("To").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            Gridview1.Columns.Item("To").HeaderCell.Style.Font = headerFont
            Gridview1.Columns.Item("To").DefaultCellStyle.Font = detailFont
            Gridview1.Columns.Item("To").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Gridview1.Columns.Item("To").Resizable = DataGridViewTriState.False
            Gridview1.Columns.Item("To").Width = 50

            Gridview1.Columns.Item("theID").Visible = False
            Gridview1.Columns.Item("workCode").Visible = False

            'GridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True
            Gridview1.Refresh()

            If finalDataSet.Tables(0).Rows.Count = 0 Then
                'File.Delete(myXML)
            Else
                'File.Delete(myXML)
                finalDataSet.Tables(0).WriteXml(myXML)
            End If

            lblRows.Text = Gridview1.Rows.Count

            connect.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        connect.Close()
    End Sub

    Public Sub load_info()
        'To load the Header Information
        If Label12.Text.ToString <> "" And Label13.Text.ToString <> "" Then
            Try

                Dim command1 As OdbcCommand
                Dim myAdapter As OdbcDataAdapter
                Dim myDataSet As DataSet

                Dim str_Employee_Show As String = "SELECT J.job_number, J.ttl_hours, J.ttl_minutes, J.job_date, J.username, M.emp_name1, J.remarks FROM t_jobsheet_trn J, m_employee M " _
                                                 & "where J.username = M.username and job_code = '" & Label12.Text.ToString.Trim.ToUpper & "' and job_number = '" & Label13.Text.ToString.Trim.ToUpper & "';"

                connect.Open()
                command1 = New OdbcCommand(str_Employee_Show, connect)

                myDataSet = New DataSet()
                myDataSet.Tables.Clear()
                myAdapter = New OdbcDataAdapter()
                myAdapter.SelectCommand = command1
                myAdapter.Fill(myDataSet, "header")

                txtDocRef.Text = myDataSet.Tables("header").Rows(0)(0).ToString
                txtUser.Text = myDataSet.Tables("header").Rows(0)(4).ToString
                txtName.Text = myDataSet.Tables("header").Rows(0)(5).ToString
                dtpDate.Value = myDataSet.Tables("header").Rows(0)(3).ToString
                txtDate.Text = dtpDate.Value.ToString("dddd, dd-MM-yyyy")
                txtRemarks1.Text = myDataSet.Tables("header").Rows(0)(6).ToString
                Dim hoursJ As Double = myDataSet.Tables("header").Rows(0)(1).ToString
                Dim minsJ As Double = myDataSet.Tables("header").Rows(0)(2).ToString

                'txtTimeSpent.Text = hoursJ.ToString("00") & ":" & minsJ.ToString("00")
                lblTimeSpent.Text = hoursJ.ToString("00") & ":" & minsJ.ToString("00")
                totalWHr = hoursJ.ToString
                totalWMin = minsJ.ToString

                connect.Close()
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        Else
            'do nothing
        End If
        connect.Close()

    End Sub

    Public Sub load_projectDetails()
        Try

            Dim command11 As OdbcCommand
            Dim myAdapter As OdbcDataAdapter
            Dim myDataSet As DataSet

            Dim str_project_Show As String = "SELECT T.doc_doc, T.doc_ref, T.dbcode, T.project_code, T.desc, T.est_hours, T.est_minutes FROM t_projects T " _
                                             & "where T.doc_doc = '" & Label7.Text.ToString.Trim.ToUpper & "' and T.doc_ref = '" & Label8.Text.ToString.Trim.ToUpper & "';"

            connect = New OdbcConnection(sqlConn)
            connect.Open()

            command11 = New OdbcCommand(str_project_Show, connect)

            myDataSet = New DataSet()
            myDataSet.Tables.Clear()
            myAdapter = New OdbcDataAdapter()
            myAdapter.SelectCommand = command11
            myAdapter.Fill(myDataSet, "header")

            txtProject.Text = myDataSet.Tables("header").Rows(0)(3).ToString
            txtRemarks.Text = myDataSet.Tables("header").Rows(0)(4).ToString
            txtEstHours.Text = myDataSet.Tables("header").Rows(0)(5).ToString
            txtEstMinutes.Text = myDataSet.Tables("header").Rows(0)(6).ToString

            connect.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        connect.Close()

    End Sub

    Public Sub load_details()

        Dim iRowIndex As Integer
        Dim itm1 As String
        Dim itm2 As String
        Dim itm3 As String
        Dim itm4 As String
        Dim itm5 As String
        Dim itm6 As String
        Dim itm7 As String
        Dim itm8 As String
        Dim itm9 As String
        Dim itm10 As String

        If Gridview1.RowCount.ToString <> 0 Then

            Try
                iRowIndex = Gridview1.CurrentRow.Index

                itm1 = Gridview1.Rows(iRowIndex).Cells("DoC").Value
                itm2 = Gridview1.Rows(iRowIndex).Cells("ReF").Value
                itm3 = Gridview1.Rows(iRowIndex).Cells("Name").Value
                itm4 = Gridview1.Rows(iRowIndex).Cells("Job Desc").Value
                itm5 = Gridview1.Rows(iRowIndex).Cells("From").Value
                itm6 = Gridview1.Rows(iRowIndex).Cells("To").Value
                itm7 = Gridview1.Rows(iRowIndex).Cells("theID").Value
                itm8 = Gridview1.Rows(iRowIndex).Cells("workCode").Value
                itm9 = Gridview1.Rows(iRowIndex).Cells("Create").Value
                itm10 = Gridview1.Rows(iRowIndex).Cells("Edit").Value

                Label7.Text = itm1.ToString.Trim
                Label8.Text = itm2.ToString.Trim
                lblID.Text = itm7.ToString.Trim
                lblWorkCode.Text = itm8.ToString.Trim
                txtCreateDate.Text = itm9.ToString
                txtEditDate.Text = itm10.ToString

                If Label7.Text = "" And Label8.Text = "" Then
                    txtCustomer.Text = "-"
                    'test
                    txtProject.Text = itm8.ToString.Trim
                    'end test
                    txtRemarks.Text = "-"
                    txtEstHours.Text = "-"
                    txtEstMinutes.Text = "-"
                    txtJobDesc.Text = itm4.ToString
                    dtpFrom.Text = itm5.ToString
                    dtpTo.Text = itm6.ToString
                Else
                    txtCustomer.Text = itm3.ToString
                    txtJobDesc.Text = itm4.ToString
                    dtpFrom.Text = itm5.ToString
                    dtpTo.Text = itm6.ToString
                    load_projectDetails()
                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        Else
            'DO NOTHING
        End If

    End Sub

    Public Sub CLEARall()
        Label7.Text = ""
        Label8.Text = ""
        txtJobDesc.Text = ""

        txtCustomer.Text = ""
        txtProject.Text = ""
        txtRemarks.Text = ""
        txtEstHours.Text = ""
        txtEstMinutes.Text = ""

        validation1.Visible = False
        validation2.Visible = False

        btnAddJob.Enabled = True
        btnAddJob.Focus()

        btnProjUpdate.Visible = False
        btnDetlCancel.Visible = False
        btnDelJob.Visible = False
        btnDetlEdit.Visible = False
        btnDetlSave.Visible = False
        txtJobDesc.ReadOnly = True
        dtpFrom.Enabled = False
        dtpTo.Enabled = False
        dtpFrom.Text = Now().ToString("HH:mm")
        dtpTo.Text = Now().ToString("HH:mm")

    End Sub

    Public Sub SaveProjectToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveProjectToolStripMenuItem.Click

        Try
            Dim focusedDate As String = dtpDate.Value.ToString("yyyy-MM-dd")
            Dim jobNumber As String = txtDocRef.Text.ToString.Trim
            Dim user As String = txtUser.Text.ToString.Trim.ToUpper
            Dim theRemark As String = txtRemarks1.Text.ToString

            Dim command10 As OdbcCommand
            Dim myAdapter10 As OdbcDataAdapter
            Dim myDataSet10 As DataSet

            connect = New OdbcConnection(sqlConn)
            connect.Open()

            Dim checkJobSheetDate As String = "SELECT job_date FROM t_jobsheet_trn where username = '" & user & "' and job_number = '" & jobNumber & "';"

            command10 = New OdbcCommand(checkJobSheetDate, connect)
            myDataSet10 = New DataSet()
            myDataSet10.Tables.Clear()
            myAdapter10 = New OdbcDataAdapter()
            myAdapter10.SelectCommand = command10
            myAdapter10.Fill(myDataSet10, "oldDate")

            Dim theOldDate1 As DateTime = myDataSet10.Tables("oldDate").Rows(0)(0)
            Dim theOldDate As String = theOldDate1.ToString("yyyy-MM-dd")

            'lblOldDate.Text = theOldDate
            If focusedDate = theOldDate Then
                'if the old date is same as the new date, do nothing loh
                btnAddJob.Visible = False
                btnAddProj.Visible = False
                btnDetlEdit.Visible = False
                btnDelJob.Visible = False
                SaveProjectToolStripMenuItem.Visible = False
                dtpDate.Enabled = False
                EditJobsheetToolStripMenuItem.Visible = True
                CloseWindowToolStripMenuItem.Visible = True
                CancelEditToolStripMenuItem.Visible = False
                PRINTToolStripMenuItem.Visible = True

                'just update the remarks with change or no change
                Dim updtcommand7 As OdbcCommand
                Dim updtadapter7 As OdbcDataAdapter

                connect = New OdbcConnection(sqlConn)
                connect.Open()

                Dim update_remark As String = "update t_jobsheet_trn set remarks = '" & theRemark.ToString & "' where username = '" & user & "' and job_number = '" & jobNumber & "';"

                updtcommand7 = New OdbcCommand(update_remark, connect)
                updtadapter7 = New OdbcDataAdapter()

                updtadapter7.UpdateCommand = updtcommand7
                updtadapter7.UpdateCommand.ExecuteNonQuery()

                txtRemarks1.Enabled = False

                'MessageBox.Show("Same date, No update", "Toinks", MessageBoxButtons.OK, MessageBoxIcon.Information)
                connect.Close()
                Exit Sub
                'end
            Else
                'if old date is not same as new date, better check database for existing jobsheet for that user on that date
                Dim command11 As OdbcCommand
                Dim myAdapter11 As OdbcDataAdapter
                Dim myDataSet11 As DataSet

                connect = New OdbcConnection(sqlConn)
                connect.Open()

                Dim checkAvailableDates As String = "SELECT job_date, job_number FROM t_jobsheet_trn where username = '" & user & "' and job_date = '" & focusedDate & "';"

                command11 = New OdbcCommand(checkAvailableDates, connect)
                myDataSet11 = New DataSet()
                myDataSet11.Tables.Clear()
                myAdapter11 = New OdbcDataAdapter()
                myAdapter11.SelectCommand = command11
                myAdapter11.Fill(myDataSet11, "checkAvail")

                If myDataSet11.Tables("checkAvail").Rows.Count <> 0 Then

                    Dim jobSheet As String = myDataSet11.Tables("checkAvail").Rows(0)(1)
                    Dim jobDate As String = myDataSet11.Tables("checkAvail").Rows(0)(0)
                    'Added on 07/08 - To show name of day
                    Dim theDate As DateTime = myDataSet11.Tables("checkAvail").Rows(0)(0)
                    Dim theDay As String = theDate.ToString("dddd")

                    MessageBox.Show("This date has been taken." & vbCrLf & "Jobsheet : " & jobSheet & vbCrLf & "Date : " & jobDate & ", " & theDay & vbCrLf & "Please select another date.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    dtpDate.Focus()
                    connect.Close()
                    Exit Try
                Else
                    'if there is no return, then update !!
                    'Dim jobSheet As String = myDataSet11.Tables("checkAvail").Rows(0)(1)
                    'Dim jobDate As String = myDataSet11.Tables("checkAvail").Rows(0)(0)

                    Dim result As Integer = MessageBox.Show("Confirm update Jobsheet date?" & vbCrLf & "Jobsheet : " & jobNumber & vbCrLf & "Date : " & focusedDate, "Update Jobsheet Date", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
                    If result = DialogResult.OK Then

                        Dim updtcommand7 As OdbcCommand
                        Dim updtadapter7 As OdbcDataAdapter

                        Dim update_JSDate As String = "update t_jobsheet_trn set job_date = '" & focusedDate & "', remarks = '" & theRemark.ToString & "' where username = '" & user & "' and job_number = '" & jobNumber & "';"

                        updtcommand7 = New OdbcCommand(update_JSDate, connect)
                        updtadapter7 = New OdbcDataAdapter()

                        updtadapter7.UpdateCommand = updtcommand7
                        updtadapter7.UpdateCommand.ExecuteNonQuery()

                        btnAddJob.Visible = False
                        btnAddProj.Visible = False
                        btnDetlEdit.Visible = False
                        btnDelJob.Visible = False
                        SaveProjectToolStripMenuItem.Visible = False
                        dtpDate.Enabled = False
                        EditJobsheetToolStripMenuItem.Visible = True
                        CloseWindowToolStripMenuItem.Visible = True
                        CancelEditToolStripMenuItem.Visible = False
                        PRINTToolStripMenuItem.Visible = True
                        txtRemarks1.Enabled = False

                        MessageBox.Show("Update OK" & vbCrLf & "Jobsheet : " & jobNumber & vbCrLf & "Date : " & focusedDate, "Cool", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        connect.Close()
                        Exit Sub
                    ElseIf result = DialogResult.Cancel Then
                        'DO NOTHING
                        connect.Close()
                    End If
                End If
            End If
            connect.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        connect.Close()

    End Sub

    Private Sub dtpDate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpDate.KeyDown
        e.SuppressKeyPress = True
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged

        txtDate.Text = dtpDate.Value.ToString("dddd, dd-MM-yyyy")
        lblFocusedDate.Text = dtpDate.Value.ToString("yyyy-MM-dd")

    End Sub

    Private Sub btnAddJob_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddJob.Click
        Gridview1.Enabled = False

        diaPICKJOB.Close()
        diaPICKJOB.StartPosition = FormStartPosition.CenterScreen
        diaPICKJOB.ShowDialog()
    End Sub

    Private Sub btnDetlCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDetlCancel.Click

        Gridview1.Enabled = True

        CLEARall()

        If Label12.Text = "" Then
            btnDelJob.Visible = False
            btnDetlEdit.Visible = False
        Else
            btnDelJob.Visible = True
            btnDetlEdit.Visible = True
        End If

    End Sub

    Public Sub btnDetlSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDetlSave.Click
        Dim jobCode As String = "JB"
        Dim jobNumber As String = txtDocRef.Text.ToString.Trim
        Dim jobDate1 As DateTime = dtpDate.Value
        Dim jobDate As String = jobDate1.ToString("yyyy-MM-dd")
        Dim docDoc As String = Label7.Text.ToString.Trim
        Dim docRef As String = Label8.Text.ToString.Trim
        Dim time_from As DateTime = dtpFrom.Text.ToString.Trim
        Dim time_to As DateTime = dtpTo.Text.ToString.Trim
        Dim theDbCode As String = lblCustCode.Text.ToString

        Dim theProject As String = txtProject.Text.ToString.Trim

        Dim duration As TimeSpan = time_to - time_from

        Dim user As String = txtUser.Text.ToString.Trim.ToUpper

        'lets check if there is jobsheet already
        connect = New OdbcConnection(sqlConn)
        connect.Open()

        If txtJobDesc.Text = "" Then
            validation1.Visible = True
            Exit Sub
        Else

            If Label7.Text <> "" And Label8.Text <> "" Then

                'connect = New OdbcConnection(sqlConn)
                'connect.Open()

                Try
                    'To check if jobsheet trn has this job number
                    Dim command8 As OdbcCommand
                    Dim myAdapter8 As OdbcDataAdapter
                    Dim myDataSet8 As DataSet

                    Dim check_trn As String = "select job_code, job_number from t_jobsheet_trn where job_code = '" & jobCode & "' and job_number = '" & jobNumber & "';"

                    command8 = New OdbcCommand(check_trn, connect)
                    myDataSet8 = New DataSet()
                    myDataSet8.Tables.Clear()
                    myAdapter8 = New OdbcDataAdapter()
                    myAdapter8.SelectCommand = command8
                    myAdapter8.Fill(myDataSet8, "trnCheck")

                    Dim inscommand5 As OdbcCommand
                    Dim insadapter5 As OdbcDataAdapter

                    Dim updtCommand As OdbcCommand
                    Dim updtAdapter As OdbcDataAdapter

                    If myDataSet8.Tables("trnCheck").Rows.Count <> 0 Then
                        'Do Nothing
                        'GO Start Insert to Detl
                    Else
                        'This happens if Record does not exists. It will add a new Number at Trn table.
                        Try
                            Dim str_JobSht_trn_Insert As String = "insert into t_jobsheet_trn(job_code, job_number, job_date, username) VALUES " _
                                                              & "('" & jobCode & "','" & jobNumber & "','" & jobDate & "','" & user & "');"

                            inscommand5 = New OdbcCommand(str_JobSht_trn_Insert, connect)
                            insadapter5 = New OdbcDataAdapter()

                            insadapter5.InsertCommand = inscommand5
                            insadapter5.InsertCommand.ExecuteNonQuery()

                            'NOW LETS SHOW THE NEW OR EXISTING NUMBER OUT /BUG FIX
                            Dim command18 As OdbcCommand
                            Dim myAdapter18 As OdbcDataAdapter
                            Dim myDataSet18 As DataSet

                            Dim check_new_JS As String = "select job_code, job_number from t_jobsheet_trn where job_code = '" & jobCode & "' and job_number = '" & jobNumber & "';"

                            command18 = New OdbcCommand(check_new_JS, connect)
                            myDataSet18 = New DataSet()
                            myDataSet18.Tables.Clear()
                            myAdapter18 = New OdbcDataAdapter()
                            myAdapter18.SelectCommand = command18
                            myAdapter18.Fill(myDataSet18, "js_check")

                            Label12.Text = myDataSet18.Tables("js_check").Rows(0)(0).ToString
                            Label13.Text = myDataSet18.Tables("js_check").Rows(0)(1).ToString
                            'END OF BUG FIX

                        Catch ex As Exception
                            MessageBox.Show("Jobsheet for " & dtpDate.Text & " already exists. " & vbCrLf & "Please select another date.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            txtJobDesc.Text = ""
                            txtJobDesc.Enabled = False
                            dtpFrom.Enabled = False
                            dtpTo.Enabled = False
                            btnAddJob.Enabled = True
                            btnAddProj.Enabled = True
                            txtCustomer.Text = ""
                            txtProject.Text = ""
                            txtRemarks.Text = ""
                            txtEstHours.Text = ""
                            txtEstMinutes.Text = ""
                            Exit Sub
                        End Try
                        Dim finalRunno As String = Convert.ToString(Convert.ToDouble(jobNumber) + 1)

                        Dim updtcommand7 As OdbcCommand
                        Dim updtadapter7 As OdbcDataAdapter

                        Dim update_running_no As String = "update m_runno set run_no = '" & finalRunno & "' where doc_doc = '" & lblDocCode.Text.ToString & "';"

                        updtcommand7 = New OdbcCommand(update_running_no, connect)
                        updtadapter7 = New OdbcDataAdapter()

                        updtadapter7.UpdateCommand = updtcommand7
                        updtadapter7.UpdateCommand.ExecuteNonQuery()
                    End If
                    'END OF CHECK

                    'Start Insert to Detl
                    Dim inscommand6 As OdbcCommand
                    Dim insadapter6 As OdbcDataAdapter

                    Dim str_JobSht_detl_Insert As String = "insert into t_jobsheet_detl(job_code, job_number, doc_doc, doc_ref, job_desc, time_from, time_to, " _
                                                           & "create_date, edit_date, dbcode, work_project_type, adjust_type) VALUES " _
                                                           & "('" & jobCode & "','" & jobNumber & "','" & docDoc & "','" & docRef & "'," _
                                                           & "?,'" & Convert.ToDateTime(time_from).ToString("HH:mm") & "','" & Convert.ToDateTime(time_to).ToString("HH:mm") & "','" & Now.ToString("yyyy-MM-dd HH:mm:ss") & "'" _
                                                           & ",'" & Now.ToString("yyyy-MM-dd HH:mm:ss") & "','" & theDbCode.ToString & "','" & theProject.ToString & "','+');"

                    inscommand6 = New OdbcCommand(str_JobSht_detl_Insert, connect)
                    insadapter6 = New OdbcDataAdapter()

                    inscommand6.Parameters.AddWithValue("?", txtJobDesc.Text)
                    insadapter6.InsertCommand = inscommand6
                    insadapter6.InsertCommand.ExecuteNonQuery()
                    'End of Insert to Detl

                    'This part starts calculating the total time of the Jobsheet
                    Dim command7 As OdbcCommand
                    Dim myAdapter7 As OdbcDataAdapter
                    Dim myDataSet7 As DataSet

                    Dim select_totaltime As String = "select HOUR(SEC_TO_TIME(SUM(TIME_TO_SEC(time_to) - TIME_TO_SEC(time_from)))) AS totalhours, MINUTE(SEC_TO_TIME(SUM(TIME_TO_SEC(time_to) - TIME_TO_SEC(time_from)))) as totalmins from t_jobsheet_detl where " _
                                                     & "job_code = '" & jobCode & "' AND job_number = '" & jobNumber & "' and deleted = 0 and adjust_type = '+';"

                    command7 = New OdbcCommand(select_totaltime, connect)
                    myDataSet7 = New DataSet()
                    myDataSet7.Tables.Clear()
                    myAdapter7 = New OdbcDataAdapter()
                    myAdapter7.SelectCommand = command7
                    myAdapter7.Fill(myDataSet7, "jobtime")

                    Dim hoursJ As String = myDataSet7.Tables("jobtime").Rows(0)(0).ToString
                    Dim minsJ As String = myDataSet7.Tables("jobtime").Rows(0)(1).ToString
                    'End of calculation and storing into variables hoursJ and minsJ

                    'Start adding total hours/mins with new Job Time elapsed
                    Dim hours = Math.Floor(duration.TotalHours)
                    Dim minutes = duration.Minutes

                    Dim newTotalHours As Double = hoursJ
                    Dim newTotalMins As Double = minsJ
                    'End of adding

                    'Update the total hours and minutes into jobsheet trn table
                    Dim update_proj_time As String = "update t_jobsheet_trn set ttl_hours = " & newTotalHours & ", ttl_minutes = " & newTotalMins & " " _
                                                     & "where job_code = '" & jobCode & "' and job_number = '" & jobNumber & "';"

                    updtCommand = New OdbcCommand(update_proj_time, connect)
                    updtAdapter = New OdbcDataAdapter()

                    updtAdapter.UpdateCommand = updtCommand
                    updtAdapter.UpdateCommand.ExecuteNonQuery()
                    'End of updating the total time and mins to jobsheet trn table
                    ''END OF INSERT TO JOBSHEET

                    'txtTimeSpent.Text = newTotalHours.ToString("00") & ":" & newTotalMins.ToString("00")
                    lblTimeSpent.Text = newTotalHours.ToString("00") & ":" & newTotalMins.ToString("00")

                    ''UPDATE INTO PROJECT TABLE (Just updating the total hours and minutes)
                    'Check total time used for Project
                    Dim command12 As OdbcCommand
                    Dim myAdapter12 As OdbcDataAdapter
                    Dim myDataSet12 As DataSet

                    Dim hours_minutes = "select ttl_hours, ttl_minutes from t_projects where doc_doc = '" & docDoc.ToString.Trim & "' and doc_ref = " & docRef.ToString.Trim & ";"

                    command12 = New OdbcCommand(hours_minutes, connect)

                    myDataSet12 = New DataSet()
                    myDataSet12.Tables.Clear()
                    myAdapter12 = New OdbcDataAdapter()
                    myAdapter12.SelectCommand = command12
                    myAdapter12.Fill(myDataSet12, "time")

                    Dim total_hours As Double = myDataSet12.Tables("time").Rows(0)(0).ToString
                    Dim total_minutes As Double = myDataSet12.Tables("time").Rows(0)(1).ToString
                    'End of checking total project time

                    'Add Project time to existing project time
                    Dim calTotal_hours As Double = total_hours + hours
                    Dim calTotal_minutes As Double = total_minutes + minutes

                    Dim newtotal_hours As Double
                    Dim newtotal_minutes As Double

                    If calTotal_minutes > 60 Then
                        newtotal_hours = calTotal_hours + 1
                        newtotal_minutes = calTotal_minutes - 60
                    Else
                        newtotal_hours = calTotal_hours
                        newtotal_minutes = calTotal_minutes
                    End If
                    'End of addition

                    'Update into projects table with new time
                    Dim updtcommand8 As OdbcCommand
                    Dim updtadapter8 As OdbcDataAdapter

                    Dim update_projects_detl As String = "update t_projects set ttl_hours = '" & newtotal_hours & "', ttl_minutes = '" & newtotal_minutes & "' where doc_doc = '" & docDoc.ToString.Trim & "' and doc_ref = '" & docRef.ToString.Trim & "';"

                    updtcommand8 = New OdbcCommand(update_projects_detl, connect)
                    updtadapter8 = New OdbcDataAdapter()

                    updtadapter8.UpdateCommand = updtcommand8
                    updtadapter8.UpdateCommand.ExecuteNonQuery()
                    'End of update
                    ''END OF UPDATE TO PROJECT TABLE

                    'To clear up the screen after updating
                    validation1.Visible = False
                    validation2.Visible = False

                    btnAddJob.Enabled = True
                    btnAddJob.Focus()
                    btnAddProj.Enabled = True

                    btnProjUpdate.Visible = False
                    btnDetlCancel.Visible = False
                    btnDelJob.Visible = False
                    btnDetlEdit.Visible = False
                    btnDetlSave.Visible = False
                    txtJobDesc.ReadOnly = True
                    dtpFrom.Enabled = False
                    dtpTo.Enabled = False

                    refresh_grid()
                    connect.Close()

                    load_details()

                    btnDelJob.Visible = True
                    btnDetlEdit.Visible = True
                    connect.Close()
                    CloseWindowToolStripMenuItem.Visible = False
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try
                connect.Close()
                SaveProjectToolStripMenuItem.Visible = True
                dtpDate.Enabled = False
            Else
                MessageBox.Show("Please assign a Project first before saving.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        End If

        Gridview1.Enabled = True
        btnDelJob.Visible = True
        btnDetlEdit.Visible = True

    End Sub

    Private Sub CloseWindowToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseWindowToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub Gridview1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles Gridview1.CellClick
        load_details()
    End Sub

    Private Sub btnDelJob_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelJob.Click

        If IsNothing(Me.Gridview1.CurrentRow) Then

            MessageBox.Show("No Record Selected", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else

            Dim iRowIndex As Integer

            iRowIndex = Gridview1.CurrentRow.Index

            'Dim del As Boolean
            'del = GridView1.Item(0, iRowIndex).Value
            Dim closed As Boolean

            Try
                'If del = True Then
                '    MessageBox.Show("Job Detail already Deleted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                '    Exit Try
                'Else
                closed = Gridview1.Item(0, iRowIndex).Value
                If closed = True Then
                    MessageBox.Show("Project already Closed. Deleting not allowed.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Try
                Else
                    'This is to check Project reference
                    Dim docDoc As String = Label7.Text.ToString.Trim
                    Dim docRef As String = Label8.Text.ToString.Trim
                    'end

                    'this is to check jobsheet reference
                    Dim doc_cod As String = Label12.Text.ToString.Trim
                    Dim doc_no As String = Label13.Text.ToString.Trim
                    'end

                    'taking the ID of the record
                    Dim theID As String = lblID.Text.ToString.Trim
                    'end

                    'Grab the time as per timepicker
                    Dim time_from As DateTime = dtpFrom.Text.ToString.Trim
                    Dim time_to As DateTime = dtpTo.Text.ToString.Trim
                    'end

                    'calculate the hours and minutes of the time
                    Dim duration As TimeSpan = time_to - time_from
                    'end of calculation

                    If Gridview1.RowCount <> 0 Then

                        Dim updtCommand As OdbcCommand
                        Dim updtAdapter As OdbcDataAdapter

                        Try

                            connect = New OdbcConnection(sqlConn)
                            connect.Open()

                            Dim result As Integer = MessageBox.Show("Confirm Delete Job Detail?", "Delete Job Detail", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
                            If result = DialogResult.OK Then
                                'lets start by checking if it's a Job or Project
                                If Label7.Text.ToString = "" And Label8.Text.ToString = "" Then
                                    'THIS IS TO UPDATE JOBS ONLY (JOBSHEET ONLY)
                                    Dim workCode As String = lblWorkCode.Text.ToString.Trim

                                    Dim update_proj_time As String = "update t_jobsheet_detl set deleted = '1' " _
                                                                     & "where job_code = '" & doc_cod.ToString.ToUpper & "' and job_number = '" & doc_no & "' and ID = '" & theID.ToString.Trim & "';"

                                    updtCommand = New OdbcCommand(update_proj_time, connect)
                                    updtAdapter = New OdbcDataAdapter()

                                    updtAdapter.UpdateCommand = updtCommand
                                    updtAdapter.UpdateCommand.ExecuteNonQuery()
                                    'we have set deleted flag to 1

                                    'now lets check if this detl item job_type is + or -
                                    Dim command8 As OdbcCommand
                                    Dim myAdapter8 As OdbcDataAdapter
                                    Dim myDataSet8 As DataSet

                                    Dim select_adjust As String = "select job_type, adjust_type from m_jobs_type where job_Type = '" & workCode.ToString.Trim & "';"

                                    command8 = New OdbcCommand(select_adjust, connect)
                                    myDataSet8 = New DataSet()
                                    myDataSet8.Tables.Clear()
                                    myAdapter8 = New OdbcDataAdapter()
                                    myAdapter8.SelectCommand = command8
                                    myAdapter8.Fill(myDataSet8, "adjust")

                                    Dim type As String = myDataSet8.Tables("adjust").Rows(0)(1).ToString
                                    'end of check
                                    If type = "+" Then

                                        ''now lets recalculate the time of that jobsheet detl excluding deleted = 1
                                        Dim command7 As OdbcCommand
                                        Dim myAdapter7 As OdbcDataAdapter
                                        Dim myDataSet7 As DataSet

                                        Dim select_totaltime As String = "select HOUR(SEC_TO_TIME(SUM(TIME_TO_SEC(time_to) - TIME_TO_SEC(time_from)))) AS totalhours, MINUTE(SEC_TO_TIME(SUM(TIME_TO_SEC(time_to) - TIME_TO_SEC(time_from)))) as totalmins from t_jobsheet_detl where " _
                                                                         & "job_code = '" & doc_cod & "' AND job_number = '" & doc_no & "' and deleted = 0 and adjust_type = '+';"

                                        command7 = New OdbcCommand(select_totaltime, connect)
                                        myDataSet7 = New DataSet()
                                        myDataSet7.Tables.Clear()
                                        myAdapter7 = New OdbcDataAdapter()
                                        myAdapter7.SelectCommand = command7
                                        myAdapter7.Fill(myDataSet7, "jobtime")

                                        'BUG FIX if delete one item and leave nothing on jobsheet
                                        Dim hoursJ As String = myDataSet7.Tables("jobtime").Rows(0)(0).ToString
                                        Dim minutesJ As String = myDataSet7.Tables("jobtime").Rows(0)(1).ToString
                                        'END bug fix
                                        ''calculation done

                                        'now lets update the jobsheet header with the new time
                                        Dim update_JS_time As String = "update t_jobsheet_trn set ttl_hours = '" & hoursJ & "', ttl_minutes = '" & minutesJ & "' " _
                                                 & "where job_code = '" & doc_cod & "' and job_number = '" & doc_no & "';"

                                        updtCommand = New OdbcCommand(update_JS_time, connect)
                                        updtAdapter = New OdbcDataAdapter()

                                        updtAdapter.UpdateCommand = updtCommand
                                        updtAdapter.UpdateCommand.ExecuteNonQuery()
                                        'end of update

                                        Dim theHours As Double
                                        Dim theMins As Double

                                        Double.TryParse(hoursJ, theHours)
                                        Double.TryParse(minutesJ, theMins)

                                        'txtTimeSpent.Text = theHours.ToString("00") & ":" & theMins.ToString("00")
                                        lblTimeSpent.Text = theHours.ToString("00") & ":" & theMins.ToString("00")
                                    Else
                                        'DO NOTHING to update the time
                                    End If
                                    'To clear up the screen after updating
                                    validation1.Visible = False
                                    validation2.Visible = False

                                    btnAddJob.Enabled = True
                                    btnAddProj.Enabled = True
                                    btnAddJob.Focus()

                                    btnProjUpdate.Visible = False
                                    btnDetlCancel.Visible = False
                                    btnDetlSave.Visible = False
                                    txtJobDesc.ReadOnly = True
                                    txtJobDesc.Text = ""
                                    dtpFrom.Enabled = False
                                    dtpTo.Enabled = False
                                    dtpFrom.Text = Now.ToString("HH:mm")
                                    dtpTo.Text = Now.ToString("HH:mm")

                                    refresh_grid()
                                    connect.Close()

                                    load_details()
                                    'THIS ENDS THE DELETING PROCESS FOR JOBS

                                ElseIf Label7.Text.ToString <> "" And Label8.Text.ToString <> "" Then
                                    'THIS WILL DELETE PROJECTS
                                    Dim update_proj_time As String = "update t_jobsheet_detl set deleted = '1' " _
                                                                     & "where job_code = '" & doc_cod.ToString.ToUpper & "' and job_number = '" & doc_no & "' and ID = '" & theID.ToString.Trim & "';"

                                    updtCommand = New OdbcCommand(update_proj_time, connect)
                                    updtAdapter = New OdbcDataAdapter()

                                    updtAdapter.UpdateCommand = updtCommand
                                    updtAdapter.UpdateCommand.ExecuteNonQuery()
                                    'we have set deleted flag to 1

                                    'now lets recalculate the time of that jobsheet detl excluding deleted = 1
                                    Dim command7 As OdbcCommand
                                    Dim myAdapter7 As OdbcDataAdapter
                                    Dim myDataSet7 As DataSet

                                    Dim select_totaltime As String = "select HOUR(SEC_TO_TIME(SUM(TIME_TO_SEC(time_to) - TIME_TO_SEC(time_from)))) AS totalhours, MINUTE(SEC_TO_TIME(SUM(TIME_TO_SEC(time_to) - TIME_TO_SEC(time_from)))) as totalmins from t_jobsheet_detl where " _
                                                                     & "job_code = '" & doc_cod & "' AND job_number = '" & doc_no & "' and deleted = 0 and adjust_type = '+';"

                                    command7 = New OdbcCommand(select_totaltime, connect)
                                    myDataSet7 = New DataSet()
                                    myDataSet7.Tables.Clear()
                                    myAdapter7 = New OdbcDataAdapter()
                                    myAdapter7.SelectCommand = command7
                                    myAdapter7.Fill(myDataSet7, "jobtime")

                                    'THE BUG IS HERE
                                    Dim hoursJ As String = myDataSet7.Tables("jobtime").Rows(0)(0).ToString
                                    Dim minutesJ As String = myDataSet7.Tables("jobtime").Rows(0)(1).ToString

                                    'calculation done
                                    Dim newHours As Double
                                    Dim newMins As Double

                                    Double.TryParse(hoursJ, newHours)
                                    Double.TryParse(minutesJ, newMins)

                                    'now lets update the jobsheet header with the new time
                                    Dim update_JS_time As String = "update t_jobsheet_trn set ttl_hours = " & newHours & ", ttl_minutes = " & newMins & " " _
                                             & "where job_code = '" & doc_cod & "' and job_number = '" & doc_no & "';"

                                    updtCommand = New OdbcCommand(update_JS_time, connect)
                                    updtAdapter = New OdbcDataAdapter()

                                    updtAdapter.UpdateCommand = updtCommand
                                    updtAdapter.UpdateCommand.ExecuteNonQuery()
                                    'end of update

                                    'txtTimeSpent.Text = newHours.ToString("00") & ":" & newMins.ToString("00")
                                    lblTimeSpent.Text = newHours.ToString("00") & ":" & newMins.ToString("00")

                                    ''Since the jobsheet total time has been updated, it is time to update the time for the project
                                    'CHECK TOTAL HOURS AND MINUTES ON PROJECTS
                                    Dim command12 As OdbcCommand
                                    Dim myAdapter12 As OdbcDataAdapter
                                    Dim myDataSet12 As DataSet

                                    Dim hours_minutes = "select ttl_hours, ttl_minutes from t_projects where doc_doc = '" & docDoc.ToString.Trim & "' and doc_ref = " & docRef.ToString.Trim & ";"

                                    command12 = New OdbcCommand(hours_minutes, connect)

                                    myDataSet12 = New DataSet()
                                    myDataSet12.Tables.Clear()
                                    myAdapter12 = New OdbcDataAdapter()
                                    myAdapter12.SelectCommand = command12
                                    myAdapter12.Fill(myDataSet12, "proj_time")

                                    Dim proj_total_hours As Double = myDataSet12.Tables("proj_time").Rows(0)(0).ToString
                                    Dim proj_total_minutes As Double = myDataSet12.Tables("proj_time").Rows(0)(1).ToString
                                    'END OF CHECKING

                                    'lets pull the hours and minutes from the dtpfrom and to
                                    Dim hours = Math.Floor(duration.TotalHours)
                                    Dim minutes = duration.Minutes

                                    'Add project time to new time (+ve / -ve) as long as it adds
                                    Dim calTotal_hours As Double = proj_total_hours - hours
                                    Dim calTotal_minutes As Double = proj_total_minutes - minutes

                                    Dim newtotal_hours1 As Double
                                    Dim newtotal_minutes1 As Double

                                    If calTotal_minutes < 0 Then
                                        newtotal_hours1 = calTotal_hours - 1
                                        newtotal_minutes1 = calTotal_minutes + 60
                                    Else
                                        newtotal_hours1 = calTotal_hours
                                        newtotal_minutes1 = calTotal_minutes
                                    End If

                                    'Now lets update project table!
                                    Dim updtcommand8 As OdbcCommand
                                    Dim updtadapter8 As OdbcDataAdapter

                                    Dim update_projects_detl As String = "update t_projects set ttl_hours = '" & newtotal_hours1 & "', ttl_minutes = '" & newtotal_minutes1 & "' where doc_doc = '" & docDoc.ToString.Trim & "' and doc_ref = '" & docRef.ToString.Trim & "';"

                                    updtcommand8 = New OdbcCommand(update_projects_detl, connect)
                                    updtadapter8 = New OdbcDataAdapter()

                                    updtadapter8.UpdateCommand = updtcommand8
                                    updtadapter8.UpdateCommand.ExecuteNonQuery()
                                    ''UPDATE PROJECT TIME DONE!

                                    ''end of updating project table

                                    'To clear up the screen after updating
                                    validation1.Visible = False
                                    validation2.Visible = False

                                    btnAddJob.Enabled = True
                                    btnAddProj.Enabled = True
                                    btnAddJob.Focus()

                                    btnProjUpdate.Visible = False
                                    btnDetlCancel.Visible = False
                                    btnDetlSave.Visible = False
                                    txtJobDesc.ReadOnly = True
                                    dtpFrom.Enabled = False
                                    dtpTo.Enabled = False
                                    txtJobDesc.Text = ""
                                    dtpFrom.Text = Now.ToString("HH:mm")
                                    dtpTo.Text = Now.ToString("HH:mm")
                                    txtCustomer.Text = ""
                                    txtProject.Text = ""
                                    txtRemarks.Text = ""
                                    txtEstHours.Text = ""
                                    txtEstMinutes.Text = ""

                                    refresh_grid()
                                    connect.Close()

                                    load_details()
                                    'END OF DELETING PROJECTS
                                Else
                                    MessageBox.Show("Oops! Something went wrong", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                    'Do nothing
                                End If

                            ElseIf result = DialogResult.Cancel Then
                                'DO NOTHING
                            End If

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    Else
                        MessageBox.Show("No Record(s) Found", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End If

                'End If

            Catch ex As Exception
                MessageBox.Show("Error Occured. Please check error message below. " & vbCrLf & vbCrLf & ex.ToString, "Cancel Project", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End If

    End Sub

    Private Sub btnDetlEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDetlEdit.Click

        If Gridview1.RowCount <> 0 Then
            Dim iRowIndex As Integer

            iRowIndex = Gridview1.CurrentRow.Index

            'Dim del As Boolean
            'del = GridView1.Item(0, iRowIndex).Value
            Dim closed As Boolean
            closed = Gridview1.Item(0, iRowIndex).Value
            Try
                'If del = True Then
                '    MessageBox.Show("Job Detail already Deleted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                '    Exit Try
                'Else
                closed = Gridview1.Item(0, iRowIndex).Value
                If closed = True Then
                    MessageBox.Show("Project already Closed. Job Detail editing not Allowed.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Try
                Else

                    Gridview1.Enabled = False

                    Dim docDoc As String = Label7.Text.ToString.ToUpper
                    Dim docRef As String = Label8.Text.ToString.ToUpper

                    'This is to check if it is a JOB or a PROJECT
                    If docDoc.ToString = "" And docRef.ToString = "" Then
                        btnJobUpdate.Visible = True
                        btnProjUpdate.Visible = False
                        btnUpdateCancel.Visible = True
                    Else
                        btnJobUpdate.Visible = False
                        btnProjUpdate.Visible = True
                        btnUpdateCancel.Visible = True
                    End If
                    'End check

                    Dim theID As String = ""

                    txtJobDesc.ReadOnly = False
                    dtpFrom.Enabled = True
                    dtpTo.Enabled = True

                    btnDetlEdit.Visible = False
                    btnDelJob.Visible = False

                    btnAddJob.Enabled = False
                    btnAddProj.Enabled = False
                    txtJobDesc.Focus()

                End If
                'End If


            Catch ex As Exception
                MessageBox.Show("Error Occured. Please check error message below. " & vbCrLf & vbCrLf & ex.ToString, "Cancel Project", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else
            MessageBox.Show("Please select record(s).", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
    End Sub

    Private Sub btnProjUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProjUpdate.Click

        Dim jobCode As String = "JB"
        Dim theID As String = lblID.Text.ToString.Trim
        Dim jobNumber As String = txtDocRef.Text.ToString.Trim
        Dim jobDate1 As DateTime = dtpDate.Value
        Dim jobDate As String = jobDate1.ToString("yyyy-MM-dd")
        Dim docDoc As String = Label7.Text.ToString.Trim
        Dim docRef As String = Label8.Text.ToString.Trim
        Dim time_from As DateTime = dtpFrom.Text.ToString
        Dim time_to As DateTime = dtpTo.Text.ToString
        Dim user As String = txtUser.Text.ToString.Trim.ToUpper

        Dim new_duration As TimeSpan = time_to - time_from

        If txtJobDesc.Text = "" Then
            validation1.Visible = True
            Exit Sub
        Else
            connect = New OdbcConnection(sqlConn)
            connect.Open()

            Try
                'Lets get the original time_from and to
                Dim command8 As OdbcCommand
                Dim myAdapter8 As OdbcDataAdapter
                Dim myDataSet8 As DataSet

                Dim check_trn As String = "SELECT time_from, time_to FROM t_jobsheet_detl where job_code = '" & jobCode & "' and job_number = '" & jobNumber & "' AND " _
                                                       & "doc_doc = '" & docDoc.ToString.Trim & "' and doc_ref = '" & docRef.ToString.Trim & "' and ID = '" & theID.ToString.Trim & "';"

                command8 = New OdbcCommand(check_trn, connect)
                myDataSet8 = New DataSet()
                myDataSet8.Tables.Clear()
                myAdapter8 = New OdbcDataAdapter()
                myAdapter8.SelectCommand = command8
                myAdapter8.Fill(myDataSet8, "taketime")

                Dim ori_time_from As DateTime = myDataSet8.Tables("taketime").Rows(0)(0).ToString
                Dim ori_time_to As DateTime = myDataSet8.Tables("taketime").Rows(0)(1).ToString

                Dim ori_duration As TimeSpan = ori_time_to - ori_time_from
                'end of job

                'Update the jobsheet detl
                Dim updcommand6 As OdbcCommand
                Dim updadapter6 As OdbcDataAdapter

                Dim updtCommand As OdbcCommand
                Dim updtAdapter As OdbcDataAdapter

                Dim str_JobSht_detl_Update As String = "update t_jobsheet_detl set job_desc = ?, time_from = '" & Convert.ToDateTime(time_from).ToString("HH:mm") & "', " _
                                                       & "time_to = '" & Convert.ToDateTime(time_to).ToString("HH:mm") & "', edit_date = '" & Now.ToString("yyyy-MM-dd HH:mm:ss") & "' where " _
                                                       & "job_code = '" & jobCode & "' AND job_number = '" & jobNumber & "' AND " _
                                                       & "doc_doc = '" & docDoc.ToString.Trim & "' and doc_ref = '" & docRef.ToString.Trim & "' and ID = '" & theID.ToString.Trim & "';"

                updcommand6 = New OdbcCommand(str_JobSht_detl_Update, connect)
                updadapter6 = New OdbcDataAdapter()

                updcommand6.Parameters.AddWithValue("?", txtJobDesc.Text)
                updadapter6.UpdateCommand = updcommand6
                updadapter6.UpdateCommand.ExecuteNonQuery()
                'End of update to jobsheet detail

                'select the new time for jobsheet
                Dim command7 As OdbcCommand
                Dim myAdapter7 As OdbcDataAdapter
                Dim myDataSet7 As DataSet

                Dim select_totaltime As String = "select HOUR(SEC_TO_TIME(SUM(TIME_TO_SEC(time_to) - TIME_TO_SEC(time_from)))) AS totalhours, MINUTE(SEC_TO_TIME(SUM(TIME_TO_SEC(time_to) - TIME_TO_SEC(time_from)))) as totalmins from t_jobsheet_detl where " _
                                                 & "job_code = '" & jobCode & "' AND job_number = '" & jobNumber & "' and deleted = 0 AND adjust_type = '+';"

                command7 = New OdbcCommand(select_totaltime, connect)
                myDataSet7 = New DataSet()
                myDataSet7.Tables.Clear()
                myAdapter7 = New OdbcDataAdapter()
                myAdapter7.SelectCommand = command7
                myAdapter7.Fill(myDataSet7, "jobtime")

                Dim hoursJ As Double = myDataSet7.Tables("jobtime").Rows(0)(0).ToString
                Dim minsJ As Double = myDataSet7.Tables("jobtime").Rows(0)(1).ToString
                'end of selection

                'UPDATE TOTAL_TIME
                Dim update_proj_time As String = "update t_jobsheet_trn set ttl_hours = " & hoursJ & ", ttl_minutes = " & minsJ & " " _
                                                 & "where job_code = '" & jobCode & "' and job_number = '" & jobNumber & "';"

                updtCommand = New OdbcCommand(update_proj_time, connect)
                updtAdapter = New OdbcDataAdapter()

                updtAdapter.UpdateCommand = updtCommand
                updtAdapter.UpdateCommand.ExecuteNonQuery()

                'txtTimeSpent.Text = hoursJ.ToString("00") & ":" & minsJ.ToString("00")
                lblTimeSpent.Text = hoursJ.ToString("00") & ":" & minsJ.ToString("00")

                'END OF UPDATE
                ''END OF INSERT TO JOBSHEET

                ''UPDATING THE PROJECT TABLE
                'Calculate the time difference of old and new times
                Dim ori_hours = Math.Floor(ori_duration.TotalHours)
                Dim ori_minutes = ori_duration.Minutes

                Dim new_hours = Math.Floor(new_duration.TotalHours)
                Dim new_minutes = new_duration.Minutes

                Dim diff_hours = new_hours - ori_hours
                Dim diff_mins = new_minutes - ori_minutes
                'end of calculation

                'CHECK TOTAL HOURS AND MINUTES ON PROJECTS
                Dim command12 As OdbcCommand
                Dim myAdapter12 As OdbcDataAdapter
                Dim myDataSet12 As DataSet

                Dim hours_minutes = "select ttl_hours, ttl_minutes from t_projects where doc_doc = '" & docDoc.ToString.Trim & "' and doc_ref = " & docRef.ToString.Trim & ";"

                command12 = New OdbcCommand(hours_minutes, connect)

                myDataSet12 = New DataSet()
                myDataSet12.Tables.Clear()
                myAdapter12 = New OdbcDataAdapter()
                myAdapter12.SelectCommand = command12
                myAdapter12.Fill(myDataSet12, "proj_time")

                Dim proj_total_hours As Double = myDataSet12.Tables("proj_time").Rows(0)(0).ToString
                Dim proj_total_minutes As Double = myDataSet12.Tables("proj_time").Rows(0)(1).ToString
                'END OF CHECKING

                'Add project time to new time (+ve / -ve) as long as it adds
                Dim calTotal_hours As Double = proj_total_hours + (diff_hours)
                Dim calTotal_minutes As Double = proj_total_minutes + (diff_mins)

                Dim newtotal_hours1 As Double
                Dim newtotal_minutes1 As Double

                If calTotal_minutes > 60 Then
                    Do While calTotal_minutes > 60
                        calTotal_minutes = calTotal_minutes - 60
                        calTotal_hours = calTotal_hours + 1
                    Loop
                    newtotal_hours1 = calTotal_hours
                    newtotal_minutes1 = calTotal_minutes
                ElseIf calTotal_minutes < 0 Then
                    Do While calTotal_minutes < 0
                        calTotal_minutes = calTotal_minutes + 60
                        calTotal_hours = calTotal_hours - 1
                    Loop
                    newtotal_hours1 = calTotal_hours
                    newtotal_minutes1 = calTotal_minutes
                Else
                    newtotal_hours1 = calTotal_hours
                    newtotal_minutes1 = calTotal_minutes
                End If

                'Now lets update project table!
                Dim updtcommand8 As OdbcCommand
                Dim updtadapter8 As OdbcDataAdapter

                Dim update_projects_detl As String = "update t_projects set ttl_hours = '" & newtotal_hours1 & "', ttl_minutes = '" & newtotal_minutes1 & "' where doc_doc = '" & docDoc.ToString.Trim & "' and doc_ref = '" & docRef.ToString.Trim & "';"

                updtcommand8 = New OdbcCommand(update_projects_detl, connect)
                updtadapter8 = New OdbcDataAdapter()

                updtadapter8.UpdateCommand = updtcommand8
                updtadapter8.UpdateCommand.ExecuteNonQuery()
                ''UPDATE PROJECT TIME DONE!

                'To clear up the screen after updating
                validation1.Visible = False
                validation2.Visible = False

                btnAddJob.Enabled = True
                btnAddProj.Enabled = True
                btnAddJob.Focus()

                btnProjUpdate.Visible = False
                btnDetlCancel.Visible = False
                btnDelJob.Visible = False
                btnDetlEdit.Visible = False
                btnDetlSave.Visible = False
                txtJobDesc.ReadOnly = True
                dtpFrom.Enabled = False
                dtpTo.Enabled = False

                refresh_grid()
                load_details()
                'connect.Close()

                btnDelJob.Visible = True
                btnDetlEdit.Visible = True
                btnDetlCancel.Visible = False
                btnUpdateCancel.Visible = False

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
            connect.Close()
            SaveProjectToolStripMenuItem.Visible = True
        End If

        Gridview1.Enabled = True
        btnDelJob.Visible = True
        btnDetlEdit.Visible = True

    End Sub

    Private Sub btnUpdateCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateCancel.Click

        Gridview1.Enabled = True

        validation1.Visible = False
        validation2.Visible = False

        btnAddJob.Enabled = True
        btnAddProj.Enabled = True
        btnAddJob.Focus()

        btnDetlSave.Visible = False
        btnProjUpdate.Visible = False
        btnJobUpdate.Visible = False
        btnUpdateCancel.Visible = False

        txtJobDesc.ReadOnly = True
        dtpFrom.Enabled = False
        dtpTo.Enabled = False

        load_details()
        connect.Close()

        btnDetlCancel.Visible = False
        btnUpdateCancel.Visible = False

        If Label12.Text = "" Then
            btnDelJob.Visible = False
            btnDetlEdit.Visible = False
        Else
            btnDelJob.Visible = True
            btnDetlEdit.Visible = True
        End If

    End Sub

    Private Sub EditJobsheetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditJobsheetToolStripMenuItem.Click
        btnAddJob.Visible = True
        btnAddProj.Visible = True
        btnDetlEdit.Visible = True
        btnDelJob.Visible = True

        SaveProjectToolStripMenuItem.Visible = True
        dtpDate.Enabled = True
        EditJobsheetToolStripMenuItem.Visible = False
        CloseWindowToolStripMenuItem.Visible = False
        CancelEditToolStripMenuItem.Visible = True
        PRINTToolStripMenuItem.Visible = False
        txtRemarks1.Enabled = True

        dtpDate.Focus()
    End Sub

    Private Sub btnAddProj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddProj.Click
        Gridview1.Enabled = False
        diaPICKPROJ.Close()
        diaPICKPROJ.StartPosition = FormStartPosition.CenterScreen
        diaPICKPROJ.ShowDialog()
    End Sub

    Public Sub btnJobSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnJobSave.Click
        Dim jobCode As String = "JB"
        Dim jobNumber As String = txtDocRef.Text.ToString.Trim
        Dim jobDate1 As DateTime = dtpDate.Value
        Dim jobDate As String = jobDate1.ToString("yyyy-MM-dd")
        Dim docDoc As String = Label7.Text.ToString.Trim
        Dim docRef As String = Label8.Text.ToString.Trim
        Dim time_from As DateTime = dtpFrom.Text.ToString.Trim
        Dim time_to As DateTime = dtpTo.Text.ToString.Trim
        Dim adjustment_type As String = lblType.Text.ToString.Trim

        Dim duration As TimeSpan = time_to - time_from

        Dim user As String = txtUser.Text.ToString.Trim.ToUpper

        connect = New OdbcConnection(sqlConn)
        connect.Open()

        If txtJobDesc.Text = "" Then
            validation1.Visible = True
            Exit Sub
        Else

            If Label7.Text <> "" And Label8.Text <> "" Then

                Try
                    'To check if jobsheet trn has this job number
                    Dim command8 As OdbcCommand
                    Dim myAdapter8 As OdbcDataAdapter
                    Dim myDataSet8 As DataSet

                    Dim check_trn As String = "select job_code, job_number from t_jobsheet_trn where job_code = '" & jobCode & "' and job_number = '" & jobNumber & "';"

                    command8 = New OdbcCommand(check_trn, connect)
                    myDataSet8 = New DataSet()
                    myDataSet8.Tables.Clear()
                    myAdapter8 = New OdbcDataAdapter()
                    myAdapter8.SelectCommand = command8
                    myAdapter8.Fill(myDataSet8, "trnCheck")

                    Dim inscommand5 As OdbcCommand
                    Dim insadapter5 As OdbcDataAdapter

                    Dim updtCommand As OdbcCommand
                    Dim updtAdapter As OdbcDataAdapter

                    If myDataSet8.Tables("trnCheck").Rows.Count <> 0 Then
                        'Do Nothing
                        'Go start Insert to Detl
                    Else
                        Try
                            'This happens if Record does not exists. It will add a new Number at Trn table.
                            Dim str_JobSht_trn_Insert As String = "insert into t_jobsheet_trn(job_code, job_number, job_date, username) VALUES " _
                                                          & "('" & jobCode & "','" & jobNumber & "','" & jobDate & "','" & user & "');"

                            inscommand5 = New OdbcCommand(str_JobSht_trn_Insert, connect)
                            insadapter5 = New OdbcDataAdapter()

                            insadapter5.InsertCommand = inscommand5
                            insadapter5.InsertCommand.ExecuteNonQuery()

                            'NOW LETS SHOW THE NEW OR EXISTING NUMBER OUT /BUG FIX
                            Dim command18 As OdbcCommand
                            Dim myAdapter18 As OdbcDataAdapter
                            Dim myDataSet18 As DataSet

                            Dim check_new_JS As String = "select job_code, job_number from t_jobsheet_trn where job_code = '" & jobCode & "' and job_number = '" & jobNumber & "';"

                            command18 = New OdbcCommand(check_new_JS, connect)
                            myDataSet18 = New DataSet()
                            myDataSet18.Tables.Clear()
                            myAdapter18 = New OdbcDataAdapter()
                            myAdapter18.SelectCommand = command18
                            myAdapter18.Fill(myDataSet18, "js_check")

                            Label12.Text = myDataSet18.Tables("js_check").Rows(0)(0).ToString
                            Label13.Text = myDataSet18.Tables("js_check").Rows(0)(1).ToString
                            'END OF BUG FIX

                        Catch ex As Exception
                            MessageBox.Show("Jobsheet for " & dtpDate.Text & " already exists. " & vbCrLf & "Please select another date.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            txtJobDesc.Text = ""
                            txtJobDesc.Enabled = False
                            dtpFrom.Enabled = False
                            dtpTo.Enabled = False
                            btnAddJob.Enabled = True
                            btnAddProj.Enabled = True
                            Exit Sub
                        End Try

                        Dim finalRunno As String = Convert.ToString(Convert.ToDouble(jobNumber) + 1)

                        Dim updtcommand7 As OdbcCommand
                        Dim updtadapter7 As OdbcDataAdapter

                        Dim update_running_no As String = "update m_runno set run_no = '" & finalRunno & "' where doc_doc = '" & lblDocCode.Text.ToString & "';"

                        updtcommand7 = New OdbcCommand(update_running_no, connect)
                        updtadapter7 = New OdbcDataAdapter()

                        updtadapter7.UpdateCommand = updtcommand7
                        updtadapter7.UpdateCommand.ExecuteNonQuery()
                    End If
                    'END OF CHECK

                    'Start Insert to Detl
                    Dim inscommand6 As OdbcCommand
                    Dim insadapter6 As OdbcDataAdapter

                    Dim str_JobSht_detl_Insert As String = "insert into t_jobsheet_detl(job_code, job_number, doc_doc, doc_ref, job_desc, time_from, time_to, " _
                                                           & "create_date, edit_date, dbcode, work_project_type, adjust_type) VALUES ('" & jobCode & "','" & jobNumber & "','',''," _
                                                           & "?,'" & Convert.ToDateTime(time_from).ToString("HH:mm") & "','" & Convert.ToDateTime(time_to).ToString("HH:mm") & "','" & Now.ToString("yyyy-MM-dd HH:mm:ss") & "','" & Now.ToString("yyyy-MM-dd HH:mm:ss") & "','-', " _
                                                           & "'" & docDoc.ToString.Trim & "','" & adjustment_type.ToString & "');"

                    inscommand6 = New OdbcCommand(str_JobSht_detl_Insert, connect)
                    insadapter6 = New OdbcDataAdapter()

                    inscommand6.Parameters.AddWithValue("?", txtJobDesc.Text)
                    insadapter6.InsertCommand = inscommand6
                    insadapter6.InsertCommand.ExecuteNonQuery()
                    ' End of Insert to Detl

                    'This is to check if the adjustment type is + then you add the time to jobsheet
                    If lblType.Text.ToString = "+" Then

                        'This part starts calculating the total time of the Jobsheet
                        Dim command7 As OdbcCommand
                        Dim myAdapter7 As OdbcDataAdapter
                        Dim myDataSet7 As DataSet

                        Dim select_totaltime As String = "select HOUR(SEC_TO_TIME(SUM(TIME_TO_SEC(time_to) - TIME_TO_SEC(time_from)))) AS totalhours, MINUTE(SEC_TO_TIME(SUM(TIME_TO_SEC(time_to) - TIME_TO_SEC(time_from)))) as totalmins from t_jobsheet_detl where " _
                                                         & "job_code = '" & jobCode & "' AND job_number = '" & jobNumber & "' and deleted = 0 and adjust_type = '+';"

                        command7 = New OdbcCommand(select_totaltime, connect)
                        myDataSet7 = New DataSet()
                        myDataSet7.Tables.Clear()
                        myAdapter7 = New OdbcDataAdapter()
                        myAdapter7.SelectCommand = command7
                        myAdapter7.Fill(myDataSet7, "jobtime")

                        Dim hoursJ As String = myDataSet7.Tables("jobtime").Rows(0)(0).ToString
                        Dim minsJ As String = myDataSet7.Tables("jobtime").Rows(0)(1).ToString
                        'End of calculation and storing into variables hoursJ and minsJ

                        'Start adding total hours/mins with new Job Time elapsed
                        Dim hours = Math.Floor(duration.TotalHours)
                        Dim minutes = duration.Minutes

                        Dim newTotalHours As Double = hoursJ
                        Dim newTotalMins As Double = minsJ

                        'End of adding

                        'Update the total hours and minutes into jobsheet trn table
                        Dim update_proj_time As String = "update t_jobsheet_trn set ttl_hours = " & newTotalHours & ", ttl_minutes = " & newTotalMins & " " _
                                                         & "where job_code = '" & jobCode & "' and job_number = '" & jobNumber & "';"

                        updtCommand = New OdbcCommand(update_proj_time, connect)
                        updtAdapter = New OdbcDataAdapter()

                        updtAdapter.UpdateCommand = updtCommand
                        updtAdapter.UpdateCommand.ExecuteNonQuery()
                        'End of updating the total time and mins to jobsheet trn table

                        ' ''END OF INSERT TO JOBSHEET

                        'txtTimeSpent.Text = newTotalHours.ToString("00") & ":" & newTotalMins.ToString("00")
                        lblTimeSpent.Text = newTotalHours.ToString("00") & ":" & newTotalMins.ToString("00")
                    Else
                        'but if the symbol is '-' then you do NOTHING.
                    End If
                    'To clear up the screen after updating
                    validation1.Visible = False
                    validation2.Visible = False

                    btnAddJob.Enabled = True
                    btnAddJob.Focus()
                    btnAddProj.Enabled = True

                    btnProjUpdate.Visible = False
                    btnDetlCancel.Visible = False
                    btnDelJob.Visible = False
                    btnDetlEdit.Visible = False
                    btnDetlSave.Visible = False
                    txtJobDesc.ReadOnly = True
                    dtpFrom.Enabled = False
                    dtpTo.Enabled = False

                    refresh_grid()
                    load_details()
                    connect.Close()

                    btnDelJob.Visible = True
                    btnDetlEdit.Visible = True
                    CloseWindowToolStripMenuItem.Visible = False
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try
                connect.Close()
                SaveProjectToolStripMenuItem.Visible = True
                dtpDate.Enabled = False
            Else
                MessageBox.Show("Please assign a Job/Task first before saving.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        End If

        Gridview1.Enabled = True
        btnDelJob.Visible = True
        btnDetlEdit.Visible = True

    End Sub

    Private Sub btnJobUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnJobUpdate.Click

        Dim jobCode As String = "JB"
        Dim theID As String = lblID.Text.ToString.Trim
        Dim jobNumber As String = txtDocRef.Text.ToString.Trim
        Dim jobDate1 As DateTime = dtpDate.Value
        Dim jobDate As String = jobDate1.ToString("yyyy-MM-dd")
        Dim docDoc As String = Label7.Text.ToString.Trim
        Dim docRef As String = Label8.Text.ToString.Trim
        Dim time_from As DateTime = dtpFrom.Text.ToString
        Dim time_to As DateTime = dtpTo.Text.ToString
        Dim user As String = txtUser.Text.ToString.Trim.ToUpper

        'MsgBox(Convert.ToDateTime(time_from).ToString("HH:mm") + vbCrLf + time_to)

        Dim adjustment_type As String = lblType.Text.ToString.Trim
        Dim workCode As String = lblWorkCode.Text.ToString.Trim

        Dim duration As TimeSpan = time_to - time_from

        If txtJobDesc.Text = "" Then
            validation1.Visible = True
            Exit Sub
        Else
            connect = New OdbcConnection(sqlConn)
            connect.Open()

            Try
                'Update jobsheet detail based on description and new time
                Dim updcommand6 As OdbcCommand
                Dim updadapter6 As OdbcDataAdapter

                Dim updtCommand As OdbcCommand
                Dim updtAdapter As OdbcDataAdapter

                Dim str_JobSht_detl_Update As String = "update t_jobsheet_detl set job_desc = ?, time_from = '" & Convert.ToDateTime(time_from).ToString("HH:mm") & "', " _
                                                       & "time_to = '" & Convert.ToDateTime(time_to).ToString("HH:mm") & "', edit_date = '" & Now.ToString("yyyy-MM-dd HH:mm:ss") & "' where " _
                                                       & "job_code = '" & jobCode & "' AND job_number = '" & jobNumber & "' AND " _
                                                       & "doc_doc = '" & docDoc.ToString.Trim & "' and doc_ref = '" & docRef.ToString.Trim & "' and ID = '" & theID.ToString.Trim & "';"

                updcommand6 = New OdbcCommand(str_JobSht_detl_Update, connect)
                updadapter6 = New OdbcDataAdapter()

                updcommand6.Parameters.AddWithValue("?", txtJobDesc.Text)
                updadapter6.UpdateCommand = updcommand6
                updadapter6.UpdateCommand.ExecuteNonQuery()
                'End of update to jobsheet detail

                'now lets check if this detl item job_type is + or -
                Dim command8 As OdbcCommand
                Dim myAdapter8 As OdbcDataAdapter
                Dim myDataSet8 As DataSet

                Dim select_adjust As String = "select job_type, adjust_type from m_jobs_type where job_Type = '" & workCode.ToString.Trim & "';"

                command8 = New OdbcCommand(select_adjust, connect)
                myDataSet8 = New DataSet()
                myDataSet8.Tables.Clear()
                myAdapter8 = New OdbcDataAdapter()
                myAdapter8.SelectCommand = command8
                myAdapter8.Fill(myDataSet8, "adjust")

                Dim type As String = myDataSet8.Tables("adjust").Rows(0)(1).ToString
                'end of check
                If type = "+" Then
                    'This is to update the jobsheet trn total time and minutes based on new updated time
                    Dim command7 As OdbcCommand
                    Dim myAdapter7 As OdbcDataAdapter
                    Dim myDataSet7 As DataSet

                    Dim select_totaltime As String = "select HOUR(SEC_TO_TIME(SUM(TIME_TO_SEC(time_to) - TIME_TO_SEC(time_from)))) AS totalhours, MINUTE(SEC_TO_TIME(SUM(TIME_TO_SEC(time_to) - TIME_TO_SEC(time_from)))) as totalmins from t_jobsheet_detl where " _
                                                     & "job_code = '" & jobCode & "' AND job_number = '" & jobNumber & "' and deleted = 0 and adjust_type = '+';"

                    command7 = New OdbcCommand(select_totaltime, connect)
                    myDataSet7 = New DataSet()
                    myDataSet7.Tables.Clear()
                    myAdapter7 = New OdbcDataAdapter()
                    myAdapter7.SelectCommand = command7
                    myAdapter7.Fill(myDataSet7, "jobtime")

                    Dim hours_ori As Double = myDataSet7.Tables("jobtime").Rows(0)(0).ToString
                    Dim minutes_ori As Double = myDataSet7.Tables("jobtime").Rows(0)(1).ToString
                    'Checking of total hours and minutes from jobsheet trn done

                    'Now lets update to jobsheet trn
                    Dim update_proj_time As String = "update t_jobsheet_trn set ttl_hours = " & hours_ori & ", ttl_minutes = " & minutes_ori & " " _
                                                     & "where job_code = '" & jobCode & "' and job_number = '" & jobNumber & "';"

                    updtCommand = New OdbcCommand(update_proj_time, connect)
                    updtAdapter = New OdbcDataAdapter()

                    updtAdapter.UpdateCommand = updtCommand
                    updtAdapter.UpdateCommand.ExecuteNonQuery()
                    'End of update to jobsheet trn.
                    'txtTimeSpent.Text = hours_ori.ToString("00") & ":" & minutes_ori.ToString("00")
                    lblTimeSpent.Text = hours_ori.ToString("00") & ":" & minutes_ori.ToString("00")
                Else
                    'DO NOTHING
                End If
                'To clear up the screen after updating
                validation1.Visible = False
                validation2.Visible = False

                btnAddJob.Enabled = True
                btnAddProj.Enabled = True
                btnAddJob.Focus()

                btnProjUpdate.Visible = False
                btnJobUpdate.Visible = False
                btnDetlCancel.Visible = False
                btnDelJob.Visible = False
                btnDetlEdit.Visible = False
                btnDetlSave.Visible = False
                txtJobDesc.ReadOnly = True
                dtpFrom.Enabled = False
                dtpTo.Enabled = False

                refresh_grid()
                load_details()
                connect.Close()

                btnDelJob.Visible = True
                btnDetlEdit.Visible = True
                btnDetlCancel.Visible = False
                btnUpdateCancel.Visible = False

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
            connect.Close()

            SaveProjectToolStripMenuItem.Visible = True
        End If

        Gridview1.Enabled = True
        btnDelJob.Visible = True
        btnDetlEdit.Visible = True
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

    Private Sub CancelEditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CancelEditToolStripMenuItem.Click
        load_info()
        load_details()

        btnAddJob.Visible = False
        btnAddProj.Visible = False
        btnDetlEdit.Visible = False
        btnDelJob.Visible = False
        SaveProjectToolStripMenuItem.Visible = False
        dtpDate.Enabled = False
        EditJobsheetToolStripMenuItem.Visible = True
        CancelEditToolStripMenuItem.Visible = False
        CloseWindowToolStripMenuItem.Visible = True
        PRINTToolStripMenuItem.Visible = True
        txtRemarks1.Enabled = False

    End Sub

    Private Sub timeTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles timeTimer.Tick
        txtTimeNow.Text = TimeString
    End Sub

    Private Sub Gridview1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Gridview1.SelectionChanged
        load_details()
    End Sub
End Class