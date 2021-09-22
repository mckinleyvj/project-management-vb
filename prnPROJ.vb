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

Public Class prnPROJ

    Dim sqlConn As String = frmConfig.theConnection
    Dim connect As OdbcConnection

    Dim headerFont As New Font("Segoe UI", 8, FontStyle.Bold)
    Dim detailFont As New Font("Segoe UI", 8, FontStyle.Regular)

    Public dsTempReport As DataSet = New DataSet
    Public dsDataSet As DataSet = New DataSet

    Private finalDataSet As New DataSet

    Public myXML As String = Application.StartupPath & "\prnProject.XML"

    Private Sub prnPROJ_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'If File.Exists(myXML) = True Then
        '    File.Delete(myXML)
        'End If
    End Sub

    Public Sub getDataSet()
        Try
            'START OF GRIDVIEW POPULATING
            Dim command3 As OdbcCommand
            Dim myAdapter3 As OdbcDataAdapter
            Dim myDataSet3 As DataSet

            connect = New OdbcConnection(sqlConn)

            Dim str_show_history As String = "SELECT D.deleted AS Deleted, T.job_date as jobDate, T.job_code, T.job_number, " _
                                             & "concat(T.job_code, '-', T.job_number) as jobRef, T.username as uName, M.emp_name1 as eName, " _
                                             & "D.time_from as timefrom, D.time_to as timeto, D.create_date as CDate, D.edit_date as EDate, D.job_desc as jobDesc " _
                                             & "FROM t_projects P, t_jobsheet_trn T, t_jobsheet_detl D, m_employee M WHERE P.doc_doc = D.doc_doc " _
                                             & "AND P.doc_ref = D.doc_ref AND T.job_code = D.job_code AND T.job_number = D.job_number AND T.username = M.username " _
                                             & "AND P.doc_doc = '" & Label12.Text.ToString.Trim.ToUpper & "' and " _
                                             & "P.doc_ref = '" & Label13.Text.ToString.Trim.ToUpper & "' and D.deleted = 0;"

            command3 = New OdbcCommand(str_show_history, connect)

            myDataSet3 = New DataSet()
            myDataSet3.Tables.Clear()
            myAdapter3 = New OdbcDataAdapter()
            myAdapter3.SelectCommand = command3
            myAdapter3.Fill(myDataSet3, "Hist")

            Dim dtRetrievedData1 As DataTable = myDataSet3.Tables("Hist")

            Dim dtData As New DataTable
            Dim dtDataRows As DataRow

            Dim chkboxcol As New DataGridViewCheckBoxColumn

            dtData.TableName = "ProjTable"
            dtData.Columns.Add("X", GetType(Boolean))
            dtData.Columns.Add("Date")
            dtData.Columns.Add("Job No")
            dtData.Columns.Add("User")
            dtData.Columns.Add("Name")
            dtData.Columns.Add("Job Desc")
            dtData.Columns.Add("From")
            dtData.Columns.Add("To")

            For Each dtDataRows In dtRetrievedData1.Rows

                Dim del As Boolean = dtDataRows("Deleted").ToString().Trim()
                Dim dat As Date = dtDataRows("jobDate").ToString().Trim()
                Dim ref As String = dtDataRows("jobRef").ToString.Trim()
                Dim user As String = dtDataRows("uName").ToString.Trim()
                Dim userName As String = dtDataRows("eName").ToString.Trim()
                Dim desK As String = dtDataRows("jobDesc").ToString

                Dim timeFrom As String = Convert.ToDateTime((dtDataRows("timefrom").ToString.Trim())).ToString("HH:mm")
                Dim timeTo As String = Convert.ToDateTime((dtDataRows("timeto").ToString.Trim())).ToString("HH:mm")

                'dtData.Rows.Add(New Object() {del.ToString.Trim(), close.ToString.Trim(), dat.ToString("dd-MM-yyyy"), ref.ToString.Trim(), cust.ToString.Trim(), custName.ToString.Trim(), proj.ToString.Trim(), projName.ToString.Trim(), ttlHrs.ToString.Trim(), ttlMins.ToString.Trim(), theDoc.ToString.Trim(), theRef.ToString.Trim(), theStat.ToString.Trim(), descp.ToString.Trim()})
                dtData.Rows.Add(New Object() {del.ToString.Trim(), dat.ToString("dd-MM-yyyy"), ref.ToString.Trim(), user.ToString.Trim(), userName.ToString.Trim(), desK.ToString.Trim(), timeFrom.ToString.Trim(), timeTo.ToString.Trim()})
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
    End Sub

    Private Sub prnPROJ_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        getDataSet()

        Dim prj_ref As String = diaPROJECTS.Label12.Text + "-" + diaPROJECTS.Label13.Text
        Dim prj_date As String = diaPROJECTS.dtpDate.Value.ToString("dd/MM/yyyy")
        Dim prj_cust As String = diaPROJECTS.cmbDB.Text.ToString
        Dim prj_desc As String = StrConv(diaPROJECTS.txtRemarks.Text.ToString, VbStrConv.ProperCase)
        Dim prj_time As String = diaPROJECTS.txtHours.Text.ToString + " Hour(s) " + diaPROJECTS.txtMins.Text.ToString + " Minute(s)"
        Dim prj_type As String = "(" + diaPROJECTS.cmbProj.Text.ToString + ")"

        Try

            Dim rptDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument

            Dim strReportPath As String = Application.StartupPath & "\r_project.rpt"
            Dim strTablePath As String = myXML

            Dim builder As New System.Data.Common.DbConnectionStringBuilder()
            builder.ConnectionString = sqlConn

            Dim zServer As String = TryCast(builder("Server"), String)
            Dim zDatabase As String = TryCast(builder("DATABASE"), String)
            'Dim zPassword As String = TryCast(builder("Password"), String)

            rptDocument.Load(strReportPath)
            rptDocument.DataSourceConnections(0).SetConnection(zServer, zDatabase, "root", "!lovetbs")

            Dim proj_ref As CrystalDecisions.CrystalReports.Engine.TextObject
            Dim proj_date As CrystalDecisions.CrystalReports.Engine.TextObject
            Dim proj_cust As CrystalDecisions.CrystalReports.Engine.TextObject
            Dim proj_desc As CrystalDecisions.CrystalReports.Engine.TextObject
            Dim proj_time As CrystalDecisions.CrystalReports.Engine.TextObject
            Dim proj_type As CrystalDecisions.CrystalReports.Engine.TextObject

            proj_ref = rptDocument.ReportDefinition.ReportObjects("proj_ref")
            proj_date = rptDocument.ReportDefinition.ReportObjects("proj_date")
            proj_cust = rptDocument.ReportDefinition.ReportObjects("proj_cust")
            proj_desc = rptDocument.ReportDefinition.ReportObjects("proj_desc")
            proj_time = rptDocument.ReportDefinition.ReportObjects("proj_time")
            proj_type = rptDocument.ReportDefinition.ReportObjects("proj_type")

            proj_ref.Text = prj_ref.ToString
            proj_date.Text = prj_date.ToString
            proj_cust.Text = prj_cust.ToString
            proj_desc.Text = prj_desc.ToString
            proj_time.Text = prj_time.ToString
            proj_type.Text = prj_type.ToString

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

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub
End Class