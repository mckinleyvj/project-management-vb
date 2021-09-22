Imports System.IO
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.Sql
Imports System.Xml
Imports System.Data.SqlClient
'Imports System.Windows.Forms.PrintPreviewDialog
Imports System.Xml.Xsl
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine.ReportDocument
Imports CrystalDecisions
Imports System.Drawing.Printing

Public Class prnJOBSHT

    Dim sqlConn As String = frmConfig.theConnection
    Dim connect As OdbcConnection

    Dim headerFont As New Font("Segoe UI", 8, FontStyle.Bold)
    Dim detailFont As New Font("Segoe UI", 8, FontStyle.Regular)

    Dim theLoginUser As String = frmMain.theUser.ToString.Trim()

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

    Dim JobDate As String
    Dim userName As String
    Dim remarkz As String

    Private Sub prnJOBSHT_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'If File.Exists(myXML) = True Then
        '    File.Delete(myXML)
        'End If
    End Sub

    Private Sub frmPrintJS_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Public Sub getInfoData()

        Try

            connect = New OdbcConnection(sqlConn)
            connect.Open()

            'GET whole detl
            Dim command1 As OdbcCommand
            Dim myAdapter1 As OdbcDataAdapter
            Dim myDataSet1 As DataSet

            Dim str_getGrid As String = "SELECT J.job_code, J.job_number, J.doc_doc, J.doc_ref, J.time_from, J.time_to, J.dbcode, J.work_project_type " _
                                        & "FROM t_jobsheet_detl j where job_code = '" & lblJobCode.Text.ToString.Trim.ToUpper & "' and job_number = '" & lblJobNumber.Text.ToString.Trim.ToUpper & "'" _
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
                                             & "FROM t_jobsheet_detl j where job_code = '" & lblJobCode.Text.ToString.Trim.ToUpper & "' and job_number = '" & lblJobNumber.Text.ToString.Trim.ToUpper & "'" _
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
                                                       & "FROM t_jobsheet_detl j where job_code = '" & lblJobCode.Text.ToString.Trim.ToUpper & "' and job_number = '" & lblJobNumber.Text.ToString.Trim.ToUpper & "' and doc_doc = ' ' and adjust_type = '+' " _
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
                                                    & "FROM t_jobsheet_detl j where job_code = '" & lblJobCode.Text.ToString.Trim.ToUpper & "' and job_number = '" & lblJobNumber.Text.ToString.Trim.ToUpper & "' and doc_doc <> ' ' " _
                                                    & "AND deleted = 0;"

            command4 = New OdbcCommand(str_getTotalProjectTime, connect)

            myDataSet4 = New DataSet()
            myDataSet4.Tables.Clear()
            myAdapter4 = New OdbcDataAdapter()
            myAdapter4.SelectCommand = command4
            myAdapter4.Fill(myDataSet4, "str_getTotalProjectTime")

            totalPHr = myDataSet4.Tables("str_getTotalProjectTime").Rows(0)(0).ToString
            totalPMin = myDataSet4.Tables("str_getTotalProjectTime").Rows(0)(1).ToString
            'end

            'get total deducted time
            Dim command5 As OdbcCommand
            Dim myAdapter5 As OdbcDataAdapter
            Dim myDataSet5 As DataSet

            Dim str_getTotalDeductTime As String = "select HOUR(SEC_TO_TIME(SUM(TIME_TO_SEC(time_to) - TIME_TO_SEC(time_from)))) AS totalhours, MINUTE(SEC_TO_TIME(SUM(TIME_TO_SEC(time_to) - TIME_TO_SEC(time_from)))) as totalmins " _
                                                  & "from(t_jobsheet_detl) where job_code = '" & lblJobCode.Text.ToString.Trim.ToUpper & "' and job_number = '" & lblJobNumber.Text.ToString.Trim.ToUpper & "' and adjust_type = '-' " _
                                                  & "AND deleted = 0;"

            command5 = New OdbcCommand(str_getTotalDeductTime, connect)

            myDataSet5 = New DataSet()
            myDataSet5.Tables.Clear()
            myAdapter5 = New OdbcDataAdapter()
            myAdapter5.SelectCommand = command5
            myAdapter5.Fill(myDataSet5, "str_getTotalDeductTime")

            totalDHr = myDataSet5.Tables("str_getTotalDeductTime").Rows(0)(0).ToString
            totalDMin = myDataSet5.Tables("str_getTotalDeductTime").Rows(0)(1).ToString
            'end

            'GET count jobs without deleted
            Dim command6 As OdbcCommand
            Dim myAdapter6 As OdbcDataAdapter
            Dim myDataSet6 As DataSet

            Dim str_getCount As String = "SELECT count(J.job_code) FROM t_jobsheet_detl J " _
                                         & "where job_code = '" & lblJobCode.Text.ToString.Trim.ToUpper & "' and job_number = '" & lblJobNumber.Text.ToString.Trim.ToUpper & "' " _
                                         & "and deleted = 0;"

            command6 = New OdbcCommand(str_getCount, connect)

            myDataSet6 = New DataSet()
            myDataSet6.Tables.Clear()
            myAdapter6 = New OdbcDataAdapter()
            myAdapter6.SelectCommand = command6
            myAdapter6.Fill(myDataSet6, "str_getCount")

            totalC = myDataSet6.Tables("str_getCount").Rows(0)(0).ToString
            'end

            'get header fields
            Dim command7 As OdbcCommand
            Dim myAdapter7 As OdbcDataAdapter
            Dim myDataSet7 As DataSet

            Dim str_getHeader As String = "select J.job_code, J.job_number, J.job_date, M.emp_name1, J.remarks " _
                                          & "FROM t_jobsheet_trn J, m_employee M " _
                                          & "where J.username = M.username " _
                                          & "AND J.job_code = '" & lblJobCode.Text.ToString.Trim.ToUpper & "' and J.job_number = '" & lblJobNumber.Text.ToString.Trim.ToUpper & "' and J.deleted = 0;"

            command7 = New OdbcCommand(str_getHeader, connect)

            myDataSet7 = New DataSet()
            myDataSet7.Tables.Clear()
            myAdapter7 = New OdbcDataAdapter()
            myAdapter7.SelectCommand = command7
            myAdapter7.Fill(myDataSet7, "str_getHeader")

            JobDate = myDataSet7.Tables("str_getHeader").Rows(0)(2).ToString
            userName = myDataSet7.Tables("str_getHeader").Rows(0)(3).ToString
            remarkz = myDataSet7.Tables("str_getHeader").Rows(0)(4).ToString
            'end

            'get total added time
            Dim command8 As OdbcCommand
            Dim myAdapter8 As OdbcDataAdapter
            Dim myDataSet8 As DataSet

            Dim str_getTotalAddTime As String = "select HOUR(SEC_TO_TIME(SUM(TIME_TO_SEC(time_to) - TIME_TO_SEC(time_from)))) AS totalhours, MINUTE(SEC_TO_TIME(SUM(TIME_TO_SEC(time_to) - TIME_TO_SEC(time_from)))) as totalmins " _
                                                  & "from(t_jobsheet_detl) where job_code = '" & lblJobCode.Text.ToString.Trim.ToUpper & "' and job_number = '" & lblJobNumber.Text.ToString.Trim.ToUpper & "' and adjust_type = '+' " _
                                                  & "AND deleted = 0;"

            command8 = New OdbcCommand(str_getTotalAddTime, connect)

            myDataSet8 = New DataSet()
            myDataSet8.Tables.Clear()
            myAdapter8 = New OdbcDataAdapter()
            myAdapter8.SelectCommand = command8
            myAdapter8.Fill(myDataSet8, "str_getTotalAddTime")

            totalWHr = myDataSet8.Tables("str_getTotalAddTime").Rows(0)(0).ToString
            totalWMin = myDataSet8.Tables("str_getTotalAddTime").Rows(0)(1).ToString
            'end

            connect.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Public Sub getDataSet()
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
                                            & "J.job_code = '" & lblJobCode.Text.ToString.Trim & "' and J.job_number = '" & lblJobNumber.Text.ToString.Trim & "' and J.deleted = 0 order by J.time_from;"

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

            If finalDataSet.Tables(0).Rows.Count = 0 Then
                'File.Delete(myXML)
            Else
                'File.Delete(myXML)
                finalDataSet.Tables(0).WriteXml(myXML)
            End If

            connect.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        connect.Close()
    End Sub

    Private Sub frmPrintJS_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            getDataSet()
            getInfoData()

            lblEmployee.Text = userName.ToString
            lblJobDate.Text = JobDate.Substring(0, 10)
            lblRemarks.Text = remarkz.ToString
            lblRowCount.Text = totalC.ToString

            If totalJSTimeHr <> "" Then
                lblTotalHours.Text = Convert.ToDouble(totalJSTimeHr).ToString("00") + " Hour(s)"
            Else
                lblTotalHours.Text = "00 Hour(s)"
            End If

            If totalJSTimeMin <> "" Then
                lblTotalMins.Text = Convert.ToDouble(totalJSTimeMin).ToString("00") + " Minute(s)"
            Else
                lblTotalMins.Text = "00 Minute(s)"
            End If
            '
            If totalWHr <> "" Then
                lblWHr.Text = Convert.ToDouble(totalWHr).ToString("00") + " Hour(s)"
            Else
                lblWHr.Text = "00 Hour(s)"
            End If

            If totalWMin <> "" Then
                lblWMin.Text = Convert.ToDouble(totalWMin).ToString("00") + " Minute(s)"
            Else
                lblWMin.Text = "00 Minute(s)"
            End If
            '
            If totalPHr <> "" Then
                lblPHr.Text = Convert.ToDouble(totalPHr).ToString("00") + " Hour(s)"
            Else
                lblPHr.Text = "00 Hour(s)"
            End If

            If totalPMin <> "" Then
                lblPMin.Text = Convert.ToDouble(totalPMin).ToString("00") + " Minute(s)"
            Else
                lblPMin.Text = "00 Minute(s)"
            End If
            '
            If totalJHr <> "" Then
                lblJHr.Text = Convert.ToDouble(totalJHr).ToString("00") + " Hour(s)"
            Else
                lblJHr.Text = "00 Hour(s)"
            End If

            If totalJMin <> "" Then
                lblJMin.Text = Convert.ToDouble(totalJMin).ToString("00") + " Minute(s)"
            Else
                lblJMin.Text = "00 Minute(s)"
            End If
            '
            If totalDHr <> "" Then
                lblDHr.Text = Convert.ToDouble(totalDHr).ToString("00") + " Hour(s)"
            Else
                lblDHr.Text = "00 Hour(s)"
            End If

            If totalDMin <> "" Then
                lblDMin.Text = Convert.ToDouble(totalDMin).ToString("00") + " Minute(s)"
            Else
                lblDMin.Text = "00 Minute(s)"
            End If

            Dim jb_ref As String = lblJobCode.Text + "-" + lblJobNumber.Text
            Dim jb_name As String = userName.ToString
            Dim jb_date As String = JobDate.Substring(0, 10)
            Dim jb_remarks As String = remarkz.ToString
            Dim jb_ttlLoggedTime As String = lblTotalHours.Text + " " + lblTotalMins.Text
            Dim jb_ttlDeductTime As String = lblDHr.Text + " " + lblDMin.Text
            Dim jb_ttlWorkTime As String = lblWHr.Text + " " + lblWMin.Text
            Dim jb_ttlPTime As String = lblPHr.Text + " " + lblPMin.Text
            Dim jb_ttlOTime As String = lblJHr.Text + " " + lblJMin.Text

            Dim rptDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument

            Dim strReportPath As String = Application.StartupPath & "\r_jobsheet.rpt"
            Dim strTablePath As String = myXML

            MsgBox(strTablePath.ToString)
            rptDocument.Load(strReportPath)

            Dim job_ref As CrystalDecisions.CrystalReports.Engine.TextObject
            Dim job_name As CrystalDecisions.CrystalReports.Engine.TextObject
            Dim job_date As CrystalDecisions.CrystalReports.Engine.TextObject
            Dim job_remarks As CrystalDecisions.CrystalReports.Engine.TextObject
            Dim job_ttlLoggedTime As CrystalDecisions.CrystalReports.Engine.TextObject
            Dim job_ttlDeductTime As CrystalDecisions.CrystalReports.Engine.TextObject
            Dim job_ttlWorkTime As CrystalDecisions.CrystalReports.Engine.TextObject
            Dim job_ttlPTime As CrystalDecisions.CrystalReports.Engine.TextObject
            Dim job_ttlOTime As CrystalDecisions.CrystalReports.Engine.TextObject

            job_ref = rptDocument.ReportDefinition.ReportObjects("job_ref")
            job_name = rptDocument.ReportDefinition.ReportObjects("job_name")
            job_date = rptDocument.ReportDefinition.ReportObjects("job_date")
            job_remarks = rptDocument.ReportDefinition.ReportObjects("job_remarks")
            job_ttlLoggedTime = rptDocument.ReportDefinition.ReportObjects("job_ttlLoggedTime")
            job_ttlDeductTime = rptDocument.ReportDefinition.ReportObjects("job_ttlDeductTime")
            job_ttlWorkTime = rptDocument.ReportDefinition.ReportObjects("job_ttlWorkTime")
            job_ttlPTime = rptDocument.ReportDefinition.ReportObjects("job_ttlPTime")
            job_ttlOTime = rptDocument.ReportDefinition.ReportObjects("job_ttlOTime")
            'proj_type = rptDocument.ReportDefinition.ReportObjects("proj_type")

            job_ref.Text = jb_ref.ToString
            job_name.Text = jb_name.ToString
            job_date.Text = jb_date.ToString
            job_remarks.Text = jb_remarks.ToString
            job_ttlLoggedTime.Text = jb_ttlLoggedTime.ToString
            job_ttlDeductTime.Text = jb_ttlDeductTime.ToString
            job_ttlWorkTime.Text = jb_ttlWorkTime.ToString
            job_ttlPTime.Text = jb_ttlPTime.ToString
            job_ttlOTime.Text = jb_ttlOTime.ToString

            Dim oLayout As New CrystalDecisions.Shared.PrintLayoutSettings
            oLayout.Centered = False
            oLayout.Scaling = PrintLayoutSettings.PrintScaling.DoNotScale

            Dim margins As PageMargins
            margins = rptDocument.PrintOptions.PageMargins
            margins.bottomMargin = 0
            margins.leftMargin = 0
            margins.rightMargin = 0
            margins.topMargin = 0.3

            rptDocument.PrintOptions.ApplyPageMargins(margins)

            rptDocument.Refresh()

            rptViewer.ReportSource = rptDocument
            rptViewer.RefreshReport()
            rptViewer.Cursor = Cursors.Default
            'rptDocument.PrintToPrinter(1, False, 0, 0)

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


    End Sub

End Class