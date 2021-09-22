Imports System.IO
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.Sql
Imports System.Xml
Imports System.Data.SqlClient
'Imports System.Windows.Forms.PrintPreviewDialog
Imports System.Xml.Xsl

Public Class frmPROJECTS

    Dim sqlConn As String = frmConfig.theConnection
    Dim connect As OdbcConnection

    Public getLength As String

    Dim headerFont As New Font("Segoe UI", 9, FontStyle.Bold)
    Dim detailFont As New Font("Segoe UI", 8, FontStyle.Regular)
    Public Sub SetDateTime()
        dtpFrom.Format = DateTimePickerFormat.Custom
        dtpTo.Format = DateTimePickerFormat.Custom

        dtpFrom.CustomFormat = "dd/MM/yyyy"
        dtpTo.CustomFormat = "dd/MM/yyyy"

        'dtpFrom.Value = Now.Date.AddYears(-1)
        dtpFrom.Value = Now.Date.AddYears(-3)
        dtpTo.Value = Now.Date
    End Sub

    Public Sub ColorTheGrid()

        Dim strExt As Boolean
        Dim strComp As String
        For Each row As DataGridViewRow In Me.GridView1.Rows
            strExt = row.Cells.Item("C").Value
            strComp = row.Cells.Item("Status").Value
            If strExt = True Then
                row.DefaultCellStyle.BackColor = Color.LightGray
                row.DefaultCellStyle.ForeColor = Color.Black
            ElseIf strExt = False Then
                If strComp = "Completed" Then
                    row.DefaultCellStyle.BackColor = Color.White
                    row.DefaultCellStyle.ForeColor = Color.Blue
                    row.DefaultCellStyle.Font = New Font("Segoe UI", 8, FontStyle.Bold)
                ElseIf strComp = "In Progress" Then
                    row.DefaultCellStyle.BackColor = Color.White
                    row.DefaultCellStyle.ForeColor = Color.Red
                    row.DefaultCellStyle.Font = New Font("Segoe UI", 8, FontStyle.Bold)
                    'Else
                    '    row.DefaultCellStyle.BackColor = Color.PaleGreen
                    '    row.DefaultCellStyle.ForeColor = Color.Black
                End If
            End If
        Next


    End Sub

    Private Sub frmPROJECTS_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.BackColor = Color.White
        Me.ForeColor = Color.Black
        MenuStrip1.BackColor = Color.Gainsboro
        MenuStrip1.ForeColor = Color.Black

        'GroupBox1.Width = Me.Width - 30
        'GroupBox2.Width = Me.Width - 30

        'GridView1.Width = GroupBox2.Width - 15

        'GroupBox2.Height = Me.Height - 175
        'GridView1.Height = GroupBox2.Height - 50

        Me.Text = "Projects Listing"

        SetDateTime()

        cmbStatus.SelectedText = ""

        dtpFrom.Focus()

        If cmbStatus.Items.Count <> 0 Then
            'cmbStatus.SelectedIndex = 1
            cmbStatus.SelectedIndex = 0
        End If

        btnSearch_Click(Nothing, Nothing)

    End Sub

    Private Sub CloseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseToolStripMenuItem.Click

        If IsNothing(Me.GridView1.CurrentRow) Then
            MessageBox.Show("No Record Selected", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Dim theDoc As String = GridView1.CurrentRow.Cells("DoC").Value
            Dim theRef As String = GridView1.CurrentRow.Cells("ReF").Value
            Dim theStatus As String = GridView1.CurrentRow.Cells("Status").Value

            Dim iRowIndex As Integer

            Try

                iRowIndex = GridView1.CurrentRow.Index

                Dim closed As Boolean
                closed = GridView1.Item(0, iRowIndex).Value
                'Dim del As Boolean
                'del = GridView1.Item(0, iRowIndex).Value
                If closed = True Then
                    MessageBox.Show("Project has been closed.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Try
                Else
                    If theStatus = "Completed" Then
                        Dim result As Integer = MessageBox.Show("Confirm Close Project : " & theDoc & "-" & theRef & "?", "Close Project", MessageBoxButtons.OKCancel)
                        If result = DialogResult.OK Then

                            Dim updtcommand As OdbcCommand
                            Dim updtadapter As OdbcDataAdapter

                            connect = New OdbcConnection(sqlConn)
                            connect.Open()

                            Dim del_emp As String = "update t_projects set closed = 1 where doc_doc = '" & theDoc & "' and doc_ref = '" & theRef & "';"

                            updtcommand = New OdbcCommand(del_emp, connect)
                            updtadapter = New OdbcDataAdapter()

                            updtadapter.UpdateCommand = updtcommand
                            updtadapter.UpdateCommand.ExecuteNonQuery()

                            'DELETE THE JOB
                            Dim delcommand As OdbcCommand
                            Dim deladapter As OdbcDataAdapter

                            Dim del_cust As String = "update t_jobsheet_detl set closed = '1' where doc_doc = '" & theDoc & "' and doc_ref = '" & theRef & "';"

                            delcommand = New OdbcCommand(del_cust, connect)
                            deladapter = New OdbcDataAdapter()

                            deladapter.UpdateCommand = delcommand
                            deladapter.UpdateCommand.ExecuteNonQuery()
                            'END OF DELETE

                            MessageBox.Show("Project " & theDoc & "-" & theRef & " successfully closed.", "Close Project", MessageBoxButtons.OK, MessageBoxIcon.Information)

                            refresh_gridview1()

                        ElseIf result = DialogResult.Cancel Then
                            'DO NOTHING
                        End If
                    Else
                        MessageBox.Show("Please change Project status to 'Completed' before closing.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Try
                    End If

                End If

            Catch ex As Exception
            MessageBox.Show("Error Occured. Please check error message below. " & vbCrLf & vbCrLf & ex.ToString, "Close Project", MessageBoxButtons.OK)
        End Try

        End If

    End Sub

    Public Sub refresh_gridview1()
        Dim command1 As OdbcCommand
        Dim myAdapter As OdbcDataAdapter
        Dim myDataSet As DataSet

        Dim d1 As DateTime = dtpFrom.Value
        Dim d2 As Date = dtpTo.Value

        connect = New OdbcConnection(sqlConn)

        Dim Astatus As String

        If cmbStatus.SelectedIndex.ToString = 1 Then
            Astatus = " AND status = 'In Progress' "

        ElseIf cmbStatus.SelectedIndex.ToString = 2 Then
            Astatus = " AND status = 'Completed' "

        ElseIf cmbStatus.SelectedIndex.ToString = 3 Then
            Astatus = " AND cancel = 1 "
        Else
            Astatus = ""
        End If

        Dim customer As String = txtDB.Text.ToString.ToUpper()

        Try

            Dim str_projs = "SELECT T.doc_type, T.doc_doc AS DoC, T.doc_ref AS ReF, T.doc_date as theDate, concat(T.doc_doc, '-', T.doc_ref) AS referenceNumber, T.dbcode as custCode, D.db_name1 as custName, " _
                            & "T.project_code as projCode, P.project_name as projName, P.project_price as projPrice, T.desc as projDesc, T.ttl_hours as totalHours, T.ttl_minutes as totalMins, T.deleted, T.cancel as docCancel, T.transtamp, T.status AS stat, T.closed as closed, T.est_hours AS esthours, T.est_minutes AS estmins " _
                            & "FROM t_projects T, m_armaster D, m_project_type P " _
                            & "WHERE T.dbcode = D.dbcode AND T.project_code = P.project_code AND T.dbcode LIKE '" & customer & "%' " _
                            & "AND T.doc_date >= '" & d1.ToString("yyyy-MM-dd") & "' AND T.doc_date <= '" & d2.ToString("yyyy-MM-dd") & "' " & Astatus & " ORDER BY T.doc_date, referenceNumber;"

            connect.Open()
            command1 = New OdbcCommand(str_projs, connect)

            myDataSet = New DataSet()
            myDataSet.Tables.Clear()
            myAdapter = New OdbcDataAdapter()
            myAdapter.SelectCommand = command1
            myAdapter.Fill(myDataSet, "Proj")

            Dim dtRetrievedData As DataTable = myDataSet.Tables(0)

            Dim dtData As New DataTable
            Dim dtDataRows As DataRow

            Dim chkboxcol As New DataGridViewCheckBoxColumn

            dtData.TableName = "ProjTable"
            'dtData.Columns.Add("X", GetType(Boolean))
            dtData.Columns.Add("C", GetType(Boolean))
            dtData.Columns.Add("Date")
            dtData.Columns.Add("Status")
            dtData.Columns.Add("Ref No.")
            dtData.Columns.Add("Code")
            dtData.Columns.Add("Name")
            dtData.Columns.Add("Project")
            dtData.Columns.Add("       Spent")
            dtData.Columns.Add("Estimated")
            dtData.Columns.Add("Remarks")
            dtData.Columns.Add("DoC")
            dtData.Columns.Add("ReF")

            For Each dtDataRows In dtRetrievedData.Rows

                Dim del As Boolean = dtDataRows("docCancel").ToString().Trim()
                Dim dat As Date = dtDataRows("theDate").ToString().Trim()
                Dim ref As String = dtDataRows("referenceNumber").ToString.Trim()
                Dim cust As String = dtDataRows("custCode").ToString.Trim()
                Dim custName As String = dtDataRows("custName").ToString.Trim()
                Dim proj As String = dtDataRows("projCode").ToString.Trim()
                Dim descp As String = dtDataRows("projDesc").ToString.Trim()

                'TIME TAKEN
                Dim hr As Double = Convert.ToDouble(dtDataRows("totalHours").ToString.Trim())
                Dim min As Double = Convert.ToDouble(dtDataRows("totalMins").ToString.Trim())
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

                'Dim ttlMin As Double = ttlMins * 0.01
                Dim ttlMin1 As Double = ttlMins1
                Dim ttlTime1 As Double = ttlHrs1 + ttlMin1

                Dim totalTime1 As String = ttlHrs1.ToString & ":" & ttlMin1.ToString("00")

                Dim theDoc As String = dtDataRows("DoC").ToString.Trim()
                Dim theRef As String = dtDataRows("ReF").ToString.Trim()
                Dim theStat As String = dtDataRows("stat").ToString.Trim()
                Dim close As Boolean = dtDataRows("closed").ToString().Trim()

                'dtData.Rows.Add(New Object() {del.ToString.Trim(), close.ToString.Trim(), dat.ToString("dd-MM-yyyy"), theStat.ToString.Trim(), ref.ToString.Trim(), cust.ToString.Trim(), custName.ToString.Trim(), proj.ToString.Trim(), totalTime.ToString.Trim(), totalTime1.ToString.Trim(), descp.ToString.Trim(), theDoc.ToString.Trim(), theRef.ToString.Trim()})
                dtData.Rows.Add(New Object() {close.ToString.Trim(), dat.ToString("dd-MM-yyyy"), theStat.ToString.Trim(), ref.ToString.Trim(), cust.ToString.Trim(), custName.ToString.Trim(), proj.ToString.Trim(), totalTime.ToString.Trim(), totalTime1.ToString.Trim(), descp.ToString.Trim(), theDoc.ToString.Trim(), theRef.ToString.Trim()})

            Next

            Dim finalDataSet As New DataSet
            finalDataSet.Tables.Add(dtData)

            GridView1.DataSource = finalDataSet.Tables(0)
            'GridView1.Columns.Item("X").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            'GridView1.Columns.Item("X").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            'GridView1.Columns.Item("X").HeaderCell.Style.Font = headerFont
            'GridView1.Columns.Item("X").DefaultCellStyle.Font = detailFont
            'GridView1.Columns.Item("X").Resizable = DataGridViewTriState.False
            'GridView1.Columns.Item("X").Width = 19

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
            GridView1.Columns.Item("Date").Width = 70

            GridView1.Columns.Item("Status").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Status").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("Status").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("Status").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Status").Width = 75

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

            GridView1.Columns.Item("       Spent").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("       Spent").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("       Spent").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("       Spent").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            GridView1.Columns.Item("       Spent").Resizable = DataGridViewTriState.False
            GridView1.Columns.Item("       Spent").Width = 67

            GridView1.Columns.Item("Estimated").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Estimated").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("Estimated").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("Estimated").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            GridView1.Columns.Item("Estimated").Resizable = DataGridViewTriState.False
            GridView1.Columns.Item("Estimated").Width = 67

            GridView1.Columns.Item("Doc").Visible = False
            GridView1.Columns.Item("ReF").Visible = False

            GridView1.Columns.Item("Remarks").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Remarks").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("Remarks").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("Remarks").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Remarks").Width = 220
            GridView1.Columns.Item("Remarks").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            GridView1.Columns.Item("Remarks").Visible = True

            GridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.NotSet
            GridView1.Refresh()
            GridView1.Focus()

            ColorTheGrid()

            connect.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        connect.Close()

        'Return

    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        refresh_gridview1()

        If GridView1.Rows.Count <> 0 Then
            GridView1_CellClick(Nothing, Nothing)
        Else
            'do nothing
        End If

    End Sub

    Private Sub frmPROJECTS_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged
        'GroupBox1.Width = Me.Width - 30
        'GroupBox2.Width = Me.Width - 30

        'GridView1.Width = GroupBox2.Width - 15

        'GroupBox2.Height = Me.Height - 175
        'GridView1.Height = GroupBox2.Height - 50


        btnSearch.Location = New Point(GroupBox1.Width - 125, Me.btnSearch.Location.Y)
    End Sub

    Public Sub GridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridView1.CellClick
        Dim iRowIndex As Integer

        iRowIndex = GridView1.CurrentRow.Index

        'Dim del As Boolean = GridView1.CurrentRow.Cells("X").Value
        Dim closed As Boolean = GridView1.CurrentRow.Cells("C").Value

        'If closed = True And del = True Then
        '    NewToolStripMenuItem.Enabled = True
        '    EditToolStripMenuItem.Enabled = False
        '    ViewProjectToolStripMenuItem.Enabled = True
        '    CloseToolStripMenuItem.Enabled = False
        'ElseIf closed = True And del = False Then
        '    NewToolStripMenuItem.Enabled = True
        '    EditToolStripMenuItem.Enabled = False
        '    ViewProjectToolStripMenuItem.Enabled = True
        '    CloseToolStripMenuItem.Enabled = False
        'ElseIf closed = False And del = True Then
        '    NewToolStripMenuItem.Enabled = True
        '    EditToolStripMenuItem.Enabled = False
        '    ViewProjectToolStripMenuItem.Enabled = True
        '    CloseToolStripMenuItem.Enabled = False
        'Else
        '    NewToolStripMenuItem.Enabled = True
        '    EditToolStripMenuItem.Enabled = True
        '    ViewProjectToolStripMenuItem.Enabled = True
        '    CloseToolStripMenuItem.Enabled = True
        'End If

        If closed = True Then
            NewToolStripMenuItem.Enabled = True
            EditToolStripMenuItem.Enabled = False
            ViewProjectToolStripMenuItem.Enabled = True
            CloseToolStripMenuItem.Enabled = False
        Else
            NewToolStripMenuItem.Enabled = True
            EditToolStripMenuItem.Enabled = True
            ViewProjectToolStripMenuItem.Enabled = True
            CloseToolStripMenuItem.Enabled = True
        End If

    End Sub

    Private Sub GridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridView1.CellContentDoubleClick

        'Dim iRowIndex As Integer

        'iRowIndex = GridView1.CurrentRow.Index


        'Dim theDoc As String = GridView1.CurrentRow.Cells("DoC").Value
        'Dim theRef As String = GridView1.CurrentRow.Cells("ReF").Value

        'diaPROJECTS.Close()

        'diaPROJECTS.Label12.Text = theDoc.ToString.Trim.ToUpper
        'diaPROJECTS.Label13.Text = theRef.ToString.Trim.ToUpper


        'disableAll()

        'diaPROJECTS.StartPosition = FormStartPosition.CenterScreen
        'diaPROJECTS.ShowDialog()
        Dim theDoc As String = GridView1.CurrentRow.Cells("DoC").Value
        Dim theRef As String = GridView1.CurrentRow.Cells("ReF").Value

        diaPROJECTS.Close()

        diaPROJECTS.Label12.Text = theDoc.ToString.Trim.ToUpper
        diaPROJECTS.Label13.Text = theRef.ToString.Trim.ToUpper

        disableAll()
        'Dim del As Boolean = GridView1.CurrentRow.Cells("X").Value
        Dim closed As Boolean = GridView1.CurrentRow.Cells("C").Value
        Dim status As String = GridView1.CurrentRow.Cells("Status").Value

        If closed = True Then
            diaPROJECTS.EditProjectToolStripMenuItem.Visible = False
            diaPROJECTS.CloseProjectToolStripMenuItem.Visible = False
            diaPROJECTS.CancelNEWToolStripMenuItem1.Visible = False
            diaPROJECTS.SaveToolStripMenuItem.Visible = False
            diaPROJECTS.CancelToolStripMenuItem.Visible = False
            diaPROJECTS.CLOSEToolStripMenuItem.Visible = True
        Else
            If status = "In Progress" Then
                diaPROJECTS.EditProjectToolStripMenuItem.Visible = True
                diaPROJECTS.CloseProjectToolStripMenuItem.Visible = False
                diaPROJECTS.CancelNEWToolStripMenuItem1.Visible = False
                diaPROJECTS.SaveToolStripMenuItem.Visible = False
                diaPROJECTS.CancelToolStripMenuItem.Visible = False
                diaPROJECTS.CLOSEToolStripMenuItem.Visible = True
            ElseIf status = "Completed" Then
                diaPROJECTS.EditProjectToolStripMenuItem.Visible = True
                diaPROJECTS.CloseProjectToolStripMenuItem.Visible = True
                diaPROJECTS.CancelNEWToolStripMenuItem1.Visible = False
                diaPROJECTS.SaveToolStripMenuItem.Visible = False
                diaPROJECTS.CancelToolStripMenuItem.Visible = False
                diaPROJECTS.CLOSEToolStripMenuItem.Visible = True
            End If
        End If

        diaPROJECTS.StartPosition = FormStartPosition.CenterScreen
        diaPROJECTS.ShowDialog()

    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click

        If IsNothing(Me.GridView1.CurrentRow) Then

            MessageBox.Show("No Record Selected", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else

            Try
                Dim iRowIndex As Double
                iRowIndex = Convert.ToDouble(GridView1.CurrentRow.Index)

                Dim theDoc As String = ""
                Dim theRef As String = ""
                Dim theStatus As String = ""
                Dim closed As String = ""
                'Dim theDoc As String = GridView1.CurrentRow.Cells("DoC").Value
                'Dim theRef As String = GridView1.CurrentRow.Cells("ReF").Value
                'Dim theStatus As String = GridView1.CurrentRow.Cells("Status").Value
                Dim dgvRow As DataGridViewRow

                For Each dgvRow In GridView1.SelectedRows
                    Dim theDoc1 As String = dgvRow.Cells("DoC").Value
                    Dim theRef1 As String = dgvRow.Cells("ReF").Value
                    Dim theStatus1 As String = dgvRow.Cells("Status").Value
                    Dim closed1 As String = dgvRow.Cells("C").Value

                    theDoc = theDoc1.ToString
                    theRef = theRef1.ToString
                    theStatus = theStatus1.ToString
                    closed = closed1.ToString
                Next

                'MsgBox(theDoc + vbCrLf + theRef + vbCrLf + closed)
                'Dim closed As Boolean

                If closed = True Then
                    MessageBox.Show("Project Closed. Edit not allowed.", "Error", MessageBoxButtons.OK)
                    Exit Try
                Else
                    diaPROJECTS.Close()

                    diaPROJECTS.Label12.Text = theDoc.ToString.Trim.ToUpper
                    diaPROJECTS.Label13.Text = theRef.ToString.Trim.ToUpper
                    diaPROJECTS.EditProjectToolStripMenuItem.Visible = False
                    diaPROJECTS.SaveToolStripMenuItem.Visible = True
                    diaPROJECTS.CancelToolStripMenuItem.Visible = True
                    diaPROJECTS.CancelNEWToolStripMenuItem1.Visible = False
                    diaPROJECTS.CLOSEToolStripMenuItem.Visible = False
                    diaPROJECTS.btnPrintPreview.Visible = False

                    If theStatus = "In Progress" Then
                        diaPROJECTS.CloseProjectToolStripMenuItem.Visible = False
                    ElseIf theStatus = "Completed" Then
                        diaPROJECTS.CloseProjectToolStripMenuItem.Visible = True
                    End If

                    diaPROJECTS.StartPosition = FormStartPosition.CenterScreen
                    diaPROJECTS.ShowDialog()
                End If

            Catch ex As Exception
                MessageBox.Show("Error Occured. Please check error message below. " & vbCrLf & vbCrLf & ex.ToString, "Edit Project", MessageBoxButtons.OK)
            End Try

        End If

    End Sub

    Public Sub getRunno()
        Dim com2 As OdbcCommand
        Dim adpt2 As OdbcDataAdapter
        Dim DS2 As DataSet

        connect = New OdbcConnection(sqlConn)

        Try

            Dim str_runno = "SELECT doc_doc, run_no, length FROM m_runno where doc_doc = 'PRJ';"

            connect.Open()
            com2 = New OdbcCommand(str_runno, connect)

            DS2 = New DataSet()
            DS2.Tables.Clear()
            adpt2 = New OdbcDataAdapter()
            adpt2.SelectCommand = com2
            adpt2.Fill(DS2, "Runno")

            diaPROJECTS.txtDocCode.Text = DS2.Tables("Runno").Rows(0)(0).ToString
            Dim getNumber As String = DS2.Tables("Runno").Rows(0)(1).ToString
            getLength = DS2.Tables("Runno").Rows(0)(2).ToString

            Dim finalSetting As String = getNumber.PadLeft(getLength, "0")
            diaPROJECTS.txtDocRef.Text = finalSetting.ToString.Trim
        Catch ex As Exception
            MessageBox.Show("Something went wrong. Please check below error message. " & vbCrLf & vbCrLf & ex.ToString)
        End Try
    End Sub

    Private Sub NewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripMenuItem.Click
        diaPROJECTS.Close()
        diaPROJECTS.cmbStatus.SelectedIndex = 0
        'diaPROJECTS.cmbDB.SelectedIndex = -1
        'diaPROJECTS.cmbDB.SelectedValue = -1
        diaPROJECTS.txtHours.Text = "0"
        diaPROJECTS.txtMins.Text = "0"
        diaPROJECTS.txtTotalTime.Text = "0.00"

        getRunno()

        diaPROJECTS.EditProjectToolStripMenuItem.Visible = False
        diaPROJECTS.SaveToolStripMenuItem.Visible = True
        diaPROJECTS.CancelToolStripMenuItem.Visible = False
        diaPROJECTS.CancelNEWToolStripMenuItem1.Visible = True
        diaPROJECTS.CloseProjectToolStripMenuItem.Visible = False
        diaPROJECTS.CLOSEToolStripMenuItem.Visible = False
        diaPROJECTS.StartPosition = FormStartPosition.CenterScreen
        diaPROJECTS.ShowDialog()
    End Sub

    Public Sub disableAll()
        diaPROJECTS.SaveToolStripMenuItem.Visible = False
        diaPROJECTS.CloseProjectToolStripMenuItem.Visible = False
        diaPROJECTS.EditProjectToolStripMenuItem.Visible = False

        diaPROJECTS.txtDocCode.ReadOnly = True
        diaPROJECTS.txtDocRef.ReadOnly = True
        diaPROJECTS.txtRemarks.ReadOnly = True
        diaPROJECTS.txtHours.ReadOnly = True
        diaPROJECTS.txtMins.ReadOnly = True
        diaPROJECTS.txtEstHours.Enabled = False
        diaPROJECTS.txtEstHours.BackColor = Color.White
        diaPROJECTS.txtEstMinutes.Enabled = False
        diaPROJECTS.txtEstMinutes.BackColor = Color.White
        diaPROJECTS.txtProjPrice.Enabled = False
        diaPROJECTS.txtProjPrice.BackColor = Color.White

        diaPROJECTS.cmbDB.Enabled = False
        If diaPROJECTS.cmbDB.Enabled = False Then
            diaPROJECTS.cmbDB.DropDownStyle = ComboBoxStyle.DropDownList
            diaPROJECTS.cmbDB.AutoCompleteMode = AutoCompleteMode.None
            diaPROJECTS.cmbDB.AutoCompleteSource = AutoCompleteSource.None
            diaPROJECTS.cmbDB.BackColor = Color.White
        Else
            diaPROJECTS.cmbDB.DropDownStyle = ComboBoxStyle.DropDown
        End If

        diaPROJECTS.cmbStatus.Enabled = False
        If diaPROJECTS.cmbStatus.Enabled = False Then
            diaPROJECTS.cmbStatus.DropDownStyle = ComboBoxStyle.DropDownList
            diaPROJECTS.cmbStatus.AutoCompleteMode = AutoCompleteMode.None
            diaPROJECTS.cmbStatus.AutoCompleteSource = AutoCompleteSource.None
            diaPROJECTS.cmbStatus.BackColor = Color.White
        Else
            diaPROJECTS.cmbStatus.DropDownStyle = ComboBoxStyle.DropDown
        End If

        diaPROJECTS.cmbProj.Enabled = False
        If diaPROJECTS.cmbProj.Enabled = False Then
            diaPROJECTS.cmbProj.DropDownStyle = ComboBoxStyle.DropDownList
            diaPROJECTS.cmbProj.AutoCompleteMode = AutoCompleteMode.None
            diaPROJECTS.cmbProj.AutoCompleteSource = AutoCompleteSource.None
            diaPROJECTS.cmbProj.BackColor = Color.White
        Else
            diaPROJECTS.cmbProj.DropDownStyle = ComboBoxStyle.DropDown
        End If


        diaPROJECTS.dtpDate.Enabled = False
        diaPROJECTS.dtpDate.CalendarTitleBackColor = Color.White

        diaPROJECTS.btnAdd.Enabled = False
        diaPROJECTS.btnAddProject.Enabled = False

        diaPROJECTS.txtRemarks.TabStop = False

        diaPROJECTS.GridView1.Focus()
        diaPROJECTS.GridView1.Select()
    End Sub

    Private Sub ViewProjectToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewProjectToolStripMenuItem.Click
        If IsNothing(Me.GridView1.CurrentRow) Then

            MessageBox.Show("No Record Selected", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else

            Dim theDoc As String = GridView1.CurrentRow.Cells("DoC").Value
            Dim theRef As String = GridView1.CurrentRow.Cells("ReF").Value
            diaPROJECTS.Close()

            diaPROJECTS.Label12.Text = theDoc.ToString.Trim.ToUpper
            diaPROJECTS.Label13.Text = theRef.ToString.Trim.ToUpper

            disableAll()
            'Dim del As Boolean = GridView1.CurrentRow.Cells("X").Value
            Dim closed As Boolean = GridView1.CurrentRow.Cells("C").Value
            Dim status As String = GridView1.CurrentRow.Cells("Status").Value

            If closed = True Then
                diaPROJECTS.EditProjectToolStripMenuItem.Visible = False
                diaPROJECTS.CloseProjectToolStripMenuItem.Visible = False
                diaPROJECTS.CancelNEWToolStripMenuItem1.Visible = False
                diaPROJECTS.SaveToolStripMenuItem.Visible = False
                diaPROJECTS.CancelToolStripMenuItem.Visible = False
                diaPROJECTS.CLOSEToolStripMenuItem.Visible = True
            Else
                If status = "In Progress" Then
                    diaPROJECTS.EditProjectToolStripMenuItem.Visible = True
                    diaPROJECTS.CloseProjectToolStripMenuItem.Visible = False
                    diaPROJECTS.CancelNEWToolStripMenuItem1.Visible = False
                    diaPROJECTS.SaveToolStripMenuItem.Visible = False
                    diaPROJECTS.CancelToolStripMenuItem.Visible = False
                    diaPROJECTS.CLOSEToolStripMenuItem.Visible = True
                ElseIf status = "Completed" Then
                    diaPROJECTS.EditProjectToolStripMenuItem.Visible = True
                    diaPROJECTS.CloseProjectToolStripMenuItem.Visible = True
                    diaPROJECTS.CancelNEWToolStripMenuItem1.Visible = False
                    diaPROJECTS.SaveToolStripMenuItem.Visible = False
                    diaPROJECTS.CancelToolStripMenuItem.Visible = False
                    diaPROJECTS.CLOSEToolStripMenuItem.Visible = True
                End If
            End If

            diaPROJECTS.StartPosition = FormStartPosition.CenterScreen
            diaPROJECTS.ShowDialog()
        End If

    End Sub

    Private Sub CancelToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If IsNothing(Me.GridView1.CurrentRow) Then

            MessageBox.Show("No Record Selected", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else

            Try
                Dim theDoc As String = GridView1.CurrentRow.Cells("DoC").Value
                Dim theRef As String = GridView1.CurrentRow.Cells("ReF").Value

                Dim iRowIndex As Integer

                iRowIndex = GridView1.CurrentRow.Index

                Dim del As Boolean
                del = GridView1.Item(0, iRowIndex).Value
                Dim closed As Boolean
                If del = True Then
                    MessageBox.Show("Project already cancelled.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Try
                Else
                    closed = GridView1.Item(1, iRowIndex).Value
                    If closed = True Then
                        MessageBox.Show("Project already closed. Cancel not allowed.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Try
                    Else
                        Dim result As Integer = MessageBox.Show("Confirm Cancel Project : " & theDoc & "-" & theRef & "?", "Cancel Project", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
                        If result = DialogResult.OK Then

                            Dim updtcommand As OdbcCommand
                            Dim updtadapter As OdbcDataAdapter

                            connect = New OdbcConnection(sqlConn)
                            connect.Open()

                            'CANCEL THE PROJECT
                            Dim del_emp As String = "update t_projects set cancel = 1, status = 'Cancelled' where doc_doc = '" & theDoc & "' and doc_ref = '" & theRef & "';"

                            updtcommand = New OdbcCommand(del_emp, connect)
                            updtadapter = New OdbcDataAdapter()

                            updtadapter.UpdateCommand = updtcommand
                            updtadapter.UpdateCommand.ExecuteNonQuery()
                            'END OF CANCEL

                            'DELETE THE JOB
                            Dim delcommand As OdbcCommand
                            Dim deladapter As OdbcDataAdapter

                            Dim del_cust As String = "update t_jobsheet_detl set deleted = '1' where doc_doc = '" & theDoc & "' and doc_ref = '" & theRef & "';"

                            delcommand = New OdbcCommand(del_cust, connect)
                            deladapter = New OdbcDataAdapter()

                            deladapter.UpdateCommand = delcommand
                            deladapter.UpdateCommand.ExecuteNonQuery()
                            'END OF DELETE

                            MessageBox.Show("Project " & theDoc & "-" & theRef & " has been cancelled.", "Cancel Project", MessageBoxButtons.OK, MessageBoxIcon.Information)

                            refresh_gridview1()

                        ElseIf result = DialogResult.Cancel Then

                        End If
                    End If

                End If

            Catch ex As Exception
                MessageBox.Show("Error Occured. Please check error message below. " & vbCrLf & vbCrLf & ex.ToString, "Cancel Project", MessageBoxButtons.OK)
            End Try
        End If

    End Sub

    Private Sub CloseWindowToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseWindowToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted
        ColorTheGrid()
    End Sub

    'Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click
    '    If IsNothing(Me.GridView1.CurrentRow) Then

    '        MessageBox.Show("No Record Selected", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    Else

    '        Dim theDoc As String = GridView1.CurrentRow.Cells("DoC").Value
    '        Dim theRef As String = GridView1.CurrentRow.Cells("ReF").Value
    '        diaPROJECTS.Close()

    '        diaPROJECTS.Label12.Text = theDoc.ToString.Trim.ToUpper
    '        diaPROJECTS.Label13.Text = theRef.ToString.Trim.ToUpper

    '        disableAll()
    '        'Dim del As Boolean = GridView1.CurrentRow.Cells("X").Value
    '        Dim closed As Boolean = GridView1.CurrentRow.Cells("C").Value
    '        Dim status As String = GridView1.CurrentRow.Cells("Status").Value

    '        If closed = True Then
    '            diaPROJECTS.EditProjectToolStripMenuItem.Visible = False
    '            diaPROJECTS.CloseProjectToolStripMenuItem.Visible = False
    '            diaPROJECTS.CancelNEWToolStripMenuItem1.Visible = False
    '            diaPROJECTS.SaveToolStripMenuItem.Visible = False
    '            diaPROJECTS.CancelToolStripMenuItem.Visible = False
    '            diaPROJECTS.CLOSEToolStripMenuItem.Visible = True
    '        Else
    '            If status = "In Progress" Then
    '                diaPROJECTS.EditProjectToolStripMenuItem.Visible = True
    '                diaPROJECTS.CloseProjectToolStripMenuItem.Visible = False
    '                diaPROJECTS.CancelNEWToolStripMenuItem1.Visible = False
    '                diaPROJECTS.SaveToolStripMenuItem.Visible = False
    '                diaPROJECTS.CancelToolStripMenuItem.Visible = False
    '                diaPROJECTS.CLOSEToolStripMenuItem.Visible = True
    '            ElseIf status = "Completed" Then
    '                diaPROJECTS.EditProjectToolStripMenuItem.Visible = True
    '                diaPROJECTS.CloseProjectToolStripMenuItem.Visible = True
    '                diaPROJECTS.CancelNEWToolStripMenuItem1.Visible = False
    '                diaPROJECTS.SaveToolStripMenuItem.Visible = False
    '                diaPROJECTS.CancelToolStripMenuItem.Visible = False
    '                diaPROJECTS.CLOSEToolStripMenuItem.Visible = True
    '            End If
    '        End If

    '        diaPROJECTS.StartPosition = FormStartPosition.CenterScreen
    '        diaPROJECTS.ShowDialog()
    '        diaPROJECTS.Hide()
    '        diaPROJECTS.btnPrintPreview_Click(Nothing, Nothing)
    '    End If
    'End Sub
End Class